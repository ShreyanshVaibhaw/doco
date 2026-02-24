use std::{
    collections::VecDeque,
    fs,
    path::{Path, PathBuf},
    sync::mpsc::{self, Receiver},
    time::{Duration, Instant},
};

use notify::{Event, RecommendedWatcher, RecursiveMode, Watcher};
use windows::Win32::Graphics::Direct2D::ID2D1DeviceContext;

use crate::{
    document::model::{Block, BlockId, DocumentModel, Heading},
    render::animation::{Animation, Easing},
    theme::Theme,
    ui::{InputEvent, Point, Rect, UIComponent},
};

const SIDEBAR_MIN_WIDTH: f32 = 200.0;
const SIDEBAR_MAX_WIDTH: f32 = 400.0;
const SIDEBAR_DEFAULT_WIDTH: f32 = 260.0;
const COLLAPSE_DURATION_S: f32 = 0.20;
const TOOLTIP_DELAY: Duration = Duration::from_millis(450);
const SIDEBAR_ITEM_HEIGHT: f32 = 24.0;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SidebarPanel {
    Files,
    Outline,
    Bookmarks,
    SearchResults,
}

impl SidebarPanel {
    pub fn all() -> [Self; 4] {
        [Self::Files, Self::Outline, Self::Bookmarks, Self::SearchResults]
    }

    pub fn title(self) -> &'static str {
        match self {
            SidebarPanel::Files => "Files",
            SidebarPanel::Outline => "Outline",
            SidebarPanel::Bookmarks => "Bookmarks",
            SidebarPanel::SearchResults => "Search Results",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SidebarAction {
    Open,
    OpenInNewTab,
    Rename,
    Delete,
    ShowInExplorer,
    CopyPath,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileIcon {
    Folder,
    Docx,
    Pdf,
    Text,
    Markdown,
    Generic,
}

impl FileIcon {
    pub fn marker(self) -> &'static str {
        match self {
            FileIcon::Folder => "[DIR]",
            FileIcon::Docx => "[DOCX]",
            FileIcon::Pdf => "[PDF]",
            FileIcon::Text => "[TXT]",
            FileIcon::Markdown => "[MD]",
            FileIcon::Generic => "[FILE]",
        }
    }
}

#[derive(Debug, Clone)]
pub struct FileNode {
    pub path: PathBuf,
    pub name: String,
    pub is_dir: bool,
    pub expanded: bool,
    pub children: Vec<FileNode>,
    pub size_bytes: Option<u64>,
    pub modified_unix_secs: Option<i64>,
}

impl FileNode {
    fn from_path(path: &Path) -> std::io::Result<Self> {
        let metadata = fs::metadata(path)?;
        let is_dir = metadata.is_dir();
        let name = path
            .file_name()
            .and_then(|s| s.to_str())
            .map(|s| s.to_string())
            .unwrap_or_else(|| path.display().to_string());

        let modified_unix_secs = metadata
            .modified()
            .ok()
            .and_then(|m| m.duration_since(std::time::SystemTime::UNIX_EPOCH).ok())
            .map(|d| d.as_secs() as i64);

        Ok(Self {
            path: path.to_path_buf(),
            name,
            is_dir,
            expanded: false,
            children: Vec::new(),
            size_bytes: if is_dir { None } else { Some(metadata.len()) },
            modified_unix_secs,
        })
    }
}

#[derive(Debug, Clone)]
pub struct OutlineItem {
    pub block_id: BlockId,
    pub title: String,
    pub level: u8,
    pub collapsed: bool,
}

#[derive(Debug, Clone)]
pub struct Bookmark {
    pub id: u64,
    pub name: String,
    pub page_number: usize,
    pub block_id: BlockId,
    pub snippet: String,
}

#[derive(Debug, Clone)]
pub struct SearchResultItem {
    pub block_id: BlockId,
    pub line_or_page: usize,
    pub snippet: String,
    pub start: usize,
    pub end: usize,
}

pub struct Sidebar {
    bounds: Rect,
    visible: bool,
    pub active_panel: SidebarPanel,
    pub width: f32,
    target_width: f32,
    collapse_anim: Option<Animation>,
    pub is_collapsed: bool,
    pub resizing: bool,
    pub file_root: Option<PathBuf>,
    pub file_tree: Vec<FileNode>,
    watcher: Option<RecommendedWatcher>,
    watch_rx: Option<Receiver<notify::Result<Event>>>,
    watch_pending: VecDeque<Event>,
    pub outline_items: Vec<OutlineItem>,
    pub bookmarks: Vec<Bookmark>,
    pub search_results: Vec<SearchResultItem>,
    pub search_term: String,
    selected_index: usize,
    hovered_item: Option<PathBuf>,
    hover_started: Option<Instant>,
    pub show_tooltip: bool,
    next_bookmark_id: u64,
    current_outline_block: Option<BlockId>,
    pending_intent: Option<SidebarIntent>,
}

impl Default for Sidebar {
    fn default() -> Self {
        Self::new()
    }
}

impl Sidebar {
    pub fn new() -> Self {
        Self {
            bounds: Rect::default(),
            visible: true,
            active_panel: SidebarPanel::Files,
            width: SIDEBAR_DEFAULT_WIDTH,
            target_width: SIDEBAR_DEFAULT_WIDTH,
            collapse_anim: None,
            is_collapsed: false,
            resizing: false,
            file_root: None,
            file_tree: Vec::new(),
            watcher: None,
            watch_rx: None,
            watch_pending: VecDeque::new(),
            outline_items: Vec::new(),
            bookmarks: Vec::new(),
            search_results: Vec::new(),
            search_term: String::new(),
            selected_index: 0,
            hovered_item: None,
            hover_started: None,
            show_tooltip: false,
            next_bookmark_id: 1,
            current_outline_block: None,
            pending_intent: None,
        }
    }

    pub fn file_context_actions() -> &'static [SidebarAction] {
        &[
            SidebarAction::Open,
            SidebarAction::OpenInNewTab,
            SidebarAction::Rename,
            SidebarAction::Delete,
            SidebarAction::ShowInExplorer,
            SidebarAction::CopyPath,
        ]
    }

    pub fn file_icon_for_node(node: &FileNode) -> FileIcon {
        if node.is_dir {
            return FileIcon::Folder;
        }
        match node
            .path
            .extension()
            .and_then(|ext| ext.to_str())
            .map(|ext| ext.to_ascii_lowercase())
            .as_deref()
        {
            Some("docx") => FileIcon::Docx,
            Some("pdf") => FileIcon::Pdf,
            Some("txt") => FileIcon::Text,
            Some("md") | Some("markdown") => FileIcon::Markdown,
            _ => FileIcon::Generic,
        }
    }

    pub fn set_active_panel(&mut self, panel: SidebarPanel) {
        self.active_panel = panel;
        self.selected_index = 0;
    }

    pub fn toggle(&mut self) {
        self.is_collapsed = !self.is_collapsed;
        let from = self.width;
        self.target_width = if self.is_collapsed { 0.0 } else { SIDEBAR_DEFAULT_WIDTH };
        self.collapse_anim = Some(Animation::new(
            from,
            self.target_width,
            COLLAPSE_DURATION_S,
            Easing::Spring,
        ));
    }

    pub fn set_width(&mut self, width: f32) {
        self.target_width = width.clamp(SIDEBAR_MIN_WIDTH, SIDEBAR_MAX_WIDTH);
        self.width = self.target_width;
        self.is_collapsed = self.width <= 0.1;
    }

    pub fn resize_by(&mut self, delta_x: f32) {
        self.set_width(self.width + delta_x);
    }

    pub fn tick(&mut self, dt_s: f32) {
        if let Some(anim) = &mut self.collapse_anim {
            if anim.update(dt_s) {
                self.width = anim.current_value;
            } else {
                self.width = anim.end_value;
                self.collapse_anim = None;
            }
        }

        if let Some(rx) = &self.watch_rx {
            while let Ok(event) = rx.try_recv() {
                if let Ok(ev) = event {
                    self.watch_pending.push_back(ev);
                }
            }
        }

        if !self.watch_pending.is_empty() {
            self.watch_pending.clear();
            self.refresh_tree();
        }

        if let Some(since) = self.hover_started {
            self.show_tooltip = since.elapsed() >= TOOLTIP_DELAY;
        }
    }

    pub fn open_folder(&mut self, root: impl AsRef<Path>) -> std::io::Result<()> {
        let root = root.as_ref().to_path_buf();
        self.file_root = Some(root.clone());
        self.file_tree = build_tree(&root, 0)?;
        self.selected_index = 0;

        let (tx, rx) = mpsc::channel();
        let mut watcher = notify::recommended_watcher(move |event| {
            let _ = tx.send(event);
        })
        .map_err(map_notify)?;
        watcher
            .watch(&root, RecursiveMode::Recursive)
            .map_err(map_notify)?;

        self.watch_rx = Some(rx);
        self.watcher = Some(watcher);
        Ok(())
    }

    pub fn refresh_tree(&mut self) {
        if let Some(root) = self.file_root.clone() {
            let _ = self.open_folder(root);
        }
    }

    pub fn populate_outline(&mut self, document: &DocumentModel) {
        self.outline_items.clear();
        for block in &document.content {
            match block {
                Block::Heading(Heading { id, level, runs }) => {
                    let title = runs.iter().map(|r| r.text.as_str()).collect::<String>();
                    self.outline_items.push(OutlineItem {
                        block_id: *id,
                        title,
                        level: (*level).clamp(1, 6),
                        collapsed: false,
                    });
                }
                Block::Paragraph(p) => {
                    if let Some(style) = &p.style_id {
                        if let Some(level) = heading_level_from_style(style) {
                            let title = p.runs.iter().map(|r| r.text.as_str()).collect::<String>();
                            self.outline_items.push(OutlineItem {
                                block_id: p.id,
                                title,
                                level,
                                collapsed: false,
                            });
                        }
                    }
                }
                _ => {}
            }
        }
    }

    pub fn add_bookmark(&mut self, block_id: BlockId, page: usize, nearby_text: &str) -> u64 {
        let id = self.next_bookmark_id;
        self.next_bookmark_id += 1;
        let title = nearby_text.chars().take(28).collect::<String>();
        self.bookmarks.push(Bookmark {
            id,
            name: if title.is_empty() {
                format!("Bookmark {id}")
            } else {
                title
            },
            page_number: page,
            block_id,
            snippet: nearby_text.chars().take(120).collect(),
        });
        id
    }

    pub fn rename_bookmark(&mut self, id: u64, name: String) -> bool {
        if let Some(bookmark) = self.bookmarks.iter_mut().find(|b| b.id == id) {
            bookmark.name = name;
            true
        } else {
            false
        }
    }

    pub fn delete_bookmark(&mut self, id: u64) -> bool {
        let before = self.bookmarks.len();
        self.bookmarks.retain(|b| b.id != id);
        self.bookmarks.len() != before
    }

    pub fn set_search_results(&mut self, term: impl Into<String>, results: Vec<SearchResultItem>) {
        self.search_term = term.into();
        self.search_results = results;
    }

    pub fn search_summary(&self) -> String {
        format!(
            "{} results for '{}'",
            self.search_results.len(),
            self.search_term
        )
    }

    pub fn set_current_outline_block(&mut self, block_id: Option<BlockId>) {
        self.current_outline_block = block_id;
    }

    pub fn panel_rows(&self, max_rows: usize) -> Vec<String> {
        let mut rows = Vec::new();
        match self.active_panel {
            SidebarPanel::Files => {
                for item in flatten_tree_with_depth(&self.file_tree).into_iter().take(max_rows) {
                    let indent = "  ".repeat(item.depth.min(8));
                    let marker = Self::file_icon_for_node(item.node).marker();
                    rows.push(format!("{indent}{marker} {}", item.node.name));
                }
            }
            SidebarPanel::Outline => {
                for item in self.outline_items.iter().take(max_rows) {
                    let indent = "  ".repeat(item.level.saturating_sub(1) as usize);
                    let active = if self.current_outline_block == Some(item.block_id) {
                        "> "
                    } else {
                        ""
                    };
                    rows.push(format!("{active}{indent}{}", item.title));
                }
            }
            SidebarPanel::Bookmarks => {
                for item in self.bookmarks.iter().take(max_rows) {
                    rows.push(format!("{} (p{})", item.name, item.page_number));
                }
            }
            SidebarPanel::SearchResults => {
                for item in self.search_results.iter().take(max_rows) {
                    rows.push(format!("{}: {}", item.line_or_page, item.snippet));
                }
            }
        }
        rows
    }

    pub fn take_intent(&mut self) -> Option<SidebarIntent> {
        self.pending_intent.take()
    }

    pub fn keyboard_navigate(&mut self, key_vk: u32) -> Option<SidebarIntent> {
        match key_vk {
            0x26 => {
                self.selected_index = self.selected_index.saturating_sub(1);
                None
            }
            0x28 => {
                let max = self.active_item_count().saturating_sub(1);
                self.selected_index = (self.selected_index + 1).min(max);
                None
            }
            0x0D => self.intent_for_selected(),
            _ => None,
        }
    }

    fn active_item_count(&self) -> usize {
        match self.active_panel {
            SidebarPanel::Files => flatten_tree(&self.file_tree).len(),
            SidebarPanel::Outline => self.outline_items.len(),
            SidebarPanel::Bookmarks => self.bookmarks.len(),
            SidebarPanel::SearchResults => self.search_results.len(),
        }
    }

    fn item_index_at_point(&self, point: Point) -> Option<usize> {
        let panel = self.panel_rect();
        if !contains(panel, point) {
            return None;
        }
        let relative_y = point.y - panel.y;
        if relative_y < 0.0 {
            return None;
        }
        let index = (relative_y / SIDEBAR_ITEM_HEIGHT).floor() as usize;
        if index < self.active_item_count() {
            Some(index)
        } else {
            None
        }
    }

    fn intent_for_selected(&self) -> Option<SidebarIntent> {
        match self.active_panel {
            SidebarPanel::Files => {
                let files = flatten_tree(&self.file_tree);
                files.get(self.selected_index).map(|node| {
                    if node.is_dir {
                        SidebarIntent::ToggleFolder(node.path.clone())
                    } else {
                        SidebarIntent::OpenFile {
                            path: node.path.clone(),
                            new_tab: false,
                        }
                    }
                })
            }
            SidebarPanel::Outline => self
                .outline_items
                .get(self.selected_index)
                .map(|it| SidebarIntent::JumpToBlock(it.block_id)),
            SidebarPanel::Bookmarks => self
                .bookmarks
                .get(self.selected_index)
                .map(|it| SidebarIntent::JumpToBlock(it.block_id)),
            SidebarPanel::SearchResults => self
                .search_results
                .get(self.selected_index)
                .map(|it| SidebarIntent::JumpToBlock(it.block_id)),
        }
    }

    pub fn toggle_folder(&mut self, path: &Path) -> bool {
        toggle_node_expanded(&mut self.file_tree, path)
    }

    pub fn hover_file_item(&mut self, maybe_path: Option<PathBuf>) {
        if self.hovered_item != maybe_path {
            self.hovered_item = maybe_path;
            self.hover_started = Some(Instant::now());
            self.show_tooltip = false;
        }
    }

    pub fn tooltip_text(&self) -> Option<String> {
        if !self.show_tooltip {
            return None;
        }

        let path = self.hovered_item.as_ref()?;
        let node = flatten_tree(&self.file_tree)
            .into_iter()
            .find(|n| &n.path == path)?;

        let size = node
            .size_bytes
            .map(|s| format!("{} bytes", s))
            .unwrap_or_else(|| "Folder".to_string());
        let modified = node
            .modified_unix_secs
            .map(|v| format!("modified {}", v))
            .unwrap_or_else(|| "modified n/a".to_string());

        Some(format!("{} | {}", size, modified))
    }

    fn top_tabs_rect(&self) -> Rect {
        Rect {
            x: self.bounds.x,
            y: self.bounds.y,
            width: self.width,
            height: 34.0,
        }
    }

    fn panel_rect(&self) -> Rect {
        Rect {
            x: self.bounds.x,
            y: self.bounds.y + 34.0,
            width: self.width,
            height: (self.bounds.height - 34.0).max(0.0),
        }
    }

    pub fn tab_hit_test(&self, point: Point) -> Option<SidebarPanel> {
        if !self.visible || self.width <= 0.0 {
            return None;
        }
        let tabs = SidebarPanel::all();
        let tab_w = self.width / tabs.len() as f32;
        let tabs_rect = self.top_tabs_rect();
        if !contains(tabs_rect, point) {
            return None;
        }

        let idx = ((point.x - tabs_rect.x) / tab_w).floor().max(0.0) as usize;
        tabs.get(idx).copied()
    }

    pub fn open_folder_for_file(&mut self, file_path: &Path) -> std::io::Result<()> {
        let root = if file_path.is_dir() {
            file_path.to_path_buf()
        } else {
            file_path
                .parent()
                .map(Path::to_path_buf)
                .unwrap_or_else(|| file_path.to_path_buf())
        };
        self.open_folder(root)
    }
}

impl UIComponent for Sidebar {
    fn layout(&mut self, bounds: Rect, _dpi: f32) {
        self.bounds = bounds;
    }

    fn render(&self, _ctx: &ID2D1DeviceContext, _theme: &Theme) {
        // Drawn by shell renderer in this development stage.
    }

    fn handle_input(&mut self, event: &InputEvent) -> bool {
        match event {
            InputEvent::KeyDown(vk) => {
                if let Some(intent) = self.keyboard_navigate(*vk) {
                    self.pending_intent = Some(intent);
                    return true;
                }
                false
            }
            InputEvent::MouseDown(point) => {
                if let Some(tab) = self.tab_hit_test(*point) {
                    self.active_panel = tab;
                    self.selected_index = 0;
                    return true;
                }
                if let Some(index) = self.item_index_at_point(*point) {
                    self.selected_index = index;
                    self.pending_intent = self.intent_for_selected();
                    return self.pending_intent.is_some();
                }
                self.hit_test(*point)
            }
            InputEvent::MouseMove(point) => {
                if self.active_panel == SidebarPanel::Files {
                    let hovered = self
                        .item_index_at_point(*point)
                        .and_then(|index| flatten_tree(&self.file_tree).get(index).map(|node| node.path.clone()));
                    self.hover_file_item(hovered);
                }
                self.hit_test(*point)
            }
            _ => false,
        }
    }

    fn hit_test(&self, point: Point) -> bool {
        self.visible && self.width > 0.0 && contains(self.bounds, point)
    }

    fn set_visible(&mut self, visible: bool) {
        self.visible = visible;
    }

    fn bounds(&self) -> Rect {
        self.bounds
    }
}

#[derive(Debug, Clone)]
pub enum SidebarIntent {
    OpenFile { path: PathBuf, new_tab: bool },
    ToggleFolder(PathBuf),
    JumpToBlock(BlockId),
}

fn build_tree(root: &Path, depth: usize) -> std::io::Result<Vec<FileNode>> {
    let mut nodes = Vec::new();
    let read = fs::read_dir(root)?;
    for entry in read.flatten() {
        let path = entry.path();
        if is_hidden(&path) {
            continue;
        }

        if let Ok(mut node) = FileNode::from_path(&path) {
            if node.is_dir && depth < 8 {
                node.children = build_tree(&path, depth + 1).unwrap_or_default();
            }
            nodes.push(node);
        }
    }

    nodes.sort_by(|a, b| a.name.to_ascii_lowercase().cmp(&b.name.to_ascii_lowercase()));
    Ok(nodes)
}

fn flatten_tree(nodes: &[FileNode]) -> Vec<&FileNode> {
    let mut out = Vec::new();
    fn walk<'a>(out: &mut Vec<&'a FileNode>, items: &'a [FileNode]) {
        for node in items {
            out.push(node);
            if node.is_dir && node.expanded {
                walk(out, &node.children);
            }
        }
    }
    walk(&mut out, nodes);
    out
}

struct FlatFileNode<'a> {
    node: &'a FileNode,
    depth: usize,
}

fn flatten_tree_with_depth(nodes: &[FileNode]) -> Vec<FlatFileNode<'_>> {
    let mut out = Vec::new();
    fn walk<'a>(out: &mut Vec<FlatFileNode<'a>>, items: &'a [FileNode], depth: usize) {
        for node in items {
            out.push(FlatFileNode { node, depth });
            if node.is_dir && node.expanded {
                walk(out, &node.children, depth + 1);
            }
        }
    }
    walk(&mut out, nodes, 0);
    out
}

fn toggle_node_expanded(nodes: &mut [FileNode], path: &Path) -> bool {
    for node in nodes {
        if node.path == path {
            if node.is_dir {
                node.expanded = !node.expanded;
                return true;
            }
            return false;
        }
        if toggle_node_expanded(&mut node.children, path) {
            return true;
        }
    }
    false
}

fn heading_level_from_style(style_id: &str) -> Option<u8> {
    let lower = style_id.to_ascii_lowercase();
    lower
        .strip_prefix("heading")
        .and_then(|n| n.parse::<u8>().ok())
        .filter(|v| (1..=6).contains(v))
}

fn is_hidden(path: &Path) -> bool {
    path.file_name()
        .and_then(|s| s.to_str())
        .map(|name| name.starts_with('.'))
        .unwrap_or(false)
}

fn contains(rect: Rect, point: Point) -> bool {
    point.x >= rect.x
        && point.x <= rect.x + rect.width
        && point.y >= rect.y
        && point.y <= rect.y + rect.height
}

fn map_notify(err: notify::Error) -> std::io::Error {
    std::io::Error::other(err.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::document::model::{Block, Paragraph, Run};

    #[test]
    fn width_clamps_to_sidebar_limits() {
        let mut sidebar = Sidebar::new();
        sidebar.set_width(120.0);
        assert_eq!(sidebar.width, SIDEBAR_MIN_WIDTH);
        sidebar.set_width(999.0);
        assert_eq!(sidebar.width, SIDEBAR_MAX_WIDTH);
    }

    #[test]
    fn tab_hit_test_selects_expected_panel() {
        let mut sidebar = Sidebar::new();
        sidebar.layout(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 260.0,
                height: 600.0,
            },
            96.0,
        );
        assert_eq!(
            sidebar.tab_hit_test(Point { x: 10.0, y: 10.0 }),
            Some(SidebarPanel::Files)
        );
        assert_eq!(
            sidebar.tab_hit_test(Point { x: 140.0, y: 10.0 }),
            Some(SidebarPanel::Bookmarks)
        );
    }

    #[test]
    fn file_context_menu_actions_match_prompt() {
        assert_eq!(
            Sidebar::file_context_actions(),
            &[
                SidebarAction::Open,
                SidebarAction::OpenInNewTab,
                SidebarAction::Rename,
                SidebarAction::Delete,
                SidebarAction::ShowInExplorer,
                SidebarAction::CopyPath,
            ]
        );
    }

    #[test]
    fn file_icon_maps_common_extensions() {
        let mk = |name: &str, is_dir: bool| FileNode {
            path: PathBuf::from(name),
            name: name.to_string(),
            is_dir,
            expanded: false,
            children: Vec::new(),
            size_bytes: None,
            modified_unix_secs: None,
        };
        assert_eq!(Sidebar::file_icon_for_node(&mk("a.docx", false)), FileIcon::Docx);
        assert_eq!(Sidebar::file_icon_for_node(&mk("a.pdf", false)), FileIcon::Pdf);
        assert_eq!(Sidebar::file_icon_for_node(&mk("a.txt", false)), FileIcon::Text);
        assert_eq!(Sidebar::file_icon_for_node(&mk("a.md", false)), FileIcon::Markdown);
        assert_eq!(Sidebar::file_icon_for_node(&mk("folder", true)), FileIcon::Folder);
    }

    #[test]
    fn search_summary_reports_count_and_term() {
        let mut sidebar = Sidebar::new();
        sidebar.set_search_results(
            "needle",
            vec![
                SearchResultItem {
                    block_id: BlockId(1),
                    line_or_page: 1,
                    snippet: "needle here".to_string(),
                    start: 0,
                    end: 6,
                },
                SearchResultItem {
                    block_id: BlockId(2),
                    line_or_page: 3,
                    snippet: "needle there".to_string(),
                    start: 0,
                    end: 6,
                },
            ],
        );
        assert_eq!(sidebar.search_summary(), "2 results for 'needle'");
    }

    #[test]
    fn bookmark_lifecycle_works() {
        let mut sidebar = Sidebar::new();
        let id = sidebar.add_bookmark(BlockId(7), 2, "Nearby text");
        assert_eq!(sidebar.bookmarks.len(), 1);
        assert!(sidebar.rename_bookmark(id, "Renamed".to_string()));
        assert_eq!(sidebar.bookmarks[0].name, "Renamed");
        assert!(sidebar.delete_bookmark(id));
        assert!(sidebar.bookmarks.is_empty());
    }

    #[test]
    fn outline_populates_from_heading_and_heading_style() {
        let mut sidebar = Sidebar::new();
        let mut doc = DocumentModel::default();
        doc.content = vec![
            Block::Heading(Heading {
                id: BlockId(1),
                level: 2,
                runs: vec![Run {
                    text: "Title".to_string(),
                    style: Default::default(),
                }],
            }),
            Block::Paragraph(Paragraph {
                id: BlockId(2),
                runs: vec![Run {
                    text: "Styled heading".to_string(),
                    style: Default::default(),
                }],
                alignment: Default::default(),
                spacing: Default::default(),
                indent: Default::default(),
                style_id: Some("Heading3".to_string()),
            }),
        ];

        sidebar.populate_outline(&doc);
        assert_eq!(sidebar.outline_items.len(), 2);
        assert_eq!(sidebar.outline_items[0].level, 2);
        assert_eq!(sidebar.outline_items[1].level, 3);
    }
}
