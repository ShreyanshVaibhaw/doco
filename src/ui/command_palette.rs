use std::collections::{HashMap, VecDeque};

use windows::Win32::Graphics::Direct2D::ID2D1DeviceContext;

use crate::{
    app::AppState,
    render::animation::{Animation, Easing},
    theme::Theme,
    ui::{InputEvent, Point, Rect, UIComponent},
};

const PALETTE_WIDTH: f32 = 600.0;
const PALETTE_MAX_HEIGHT: f32 = 400.0;
const PALETTE_FADE_S: f32 = 0.10;

pub struct Command {
    pub id: &'static str,
    pub label: &'static str,
    pub category: &'static str,
    pub shortcut: Option<&'static str>,
    pub action: Box<dyn Fn(&mut AppState) + Send + Sync>,
    pub is_enabled: Box<dyn Fn(&AppState) -> bool + Send + Sync>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum QuickActionMode {
    Command,
    GoToHeading,
    GoToLineOrPage,
    GoToBookmark,
    SearchDocument,
}

impl QuickActionMode {
    fn from_query(query: &str) -> Self {
        match query.chars().next() {
            Some('#') => Self::GoToHeading,
            Some(':') => Self::GoToLineOrPage,
            Some('@') => Self::GoToBookmark,
            Some('>') => Self::Command,
            _ => Self::SearchDocument,
        }
    }
}

#[derive(Debug, Clone)]
pub struct CommandMatch {
    pub command_index: usize,
    pub score: i32,
    pub matched_chars: Vec<usize>,
}

pub struct CommandPalette {
    bounds: Rect,
    visible: bool,
    opacity: f32,
    fade: Option<Animation>,
    pub query: String,
    pub mode: QuickActionMode,
    pub selected: usize,
    commands: Vec<Command>,
    results: Vec<CommandMatch>,
    recent_ids: VecDeque<&'static str>,
    pub grouped_result_headers: Vec<(String, usize)>,
    pub close_on_click_outside: bool,
}

impl Default for CommandPalette {
    fn default() -> Self {
        Self::new()
    }
}

impl CommandPalette {
    pub fn new() -> Self {
        let mut palette = Self {
            bounds: Rect::default(),
            visible: false,
            opacity: 0.0,
            fade: None,
            query: String::new(),
            mode: QuickActionMode::Command,
            selected: 0,
            commands: default_command_registry(),
            results: Vec::new(),
            recent_ids: VecDeque::new(),
            grouped_result_headers: Vec::new(),
            close_on_click_outside: true,
        };
        palette.refresh_results(None);
        palette
    }

    pub fn open(&mut self) {
        self.visible = true;
        self.opacity = 0.0;
        self.fade = Some(Animation::new(0.0, 1.0, PALETTE_FADE_S, Easing::EaseOutCubic));
        self.refresh_results(None);
    }

    pub fn close(&mut self) {
        self.visible = false;
        self.opacity = 0.0;
        self.fade = None;
    }

    pub fn is_open(&self) -> bool {
        self.visible
    }

    pub fn tick(&mut self, dt_s: f32) {
        if let Some(anim) = &mut self.fade {
            if anim.update(dt_s) {
                self.opacity = anim.current_value.clamp(0.0, 1.0);
            } else {
                self.opacity = anim.end_value.clamp(0.0, 1.0);
                self.fade = None;
            }
        }
    }

    pub fn set_query(&mut self, text: impl Into<String>) {
        self.query = text.into();
        self.mode = QuickActionMode::from_query(self.query.as_str());
        self.selected = 0;
        self.refresh_results(None);
    }

    pub fn append_char(&mut self, ch: char) {
        self.query.push(ch);
        self.mode = QuickActionMode::from_query(self.query.as_str());
        self.selected = 0;
        self.refresh_results(None);
    }

    pub fn backspace(&mut self) {
        self.query.pop();
        self.mode = QuickActionMode::from_query(self.query.as_str());
        self.selected = 0;
        self.refresh_results(None);
    }

    pub fn move_selection(&mut self, down: bool) {
        if self.results.is_empty() {
            self.selected = 0;
            return;
        }

        if down {
            self.selected = (self.selected + 1).min(self.results.len() - 1);
        } else {
            self.selected = self.selected.saturating_sub(1);
        }
    }

    pub fn execute_selected(&mut self, app_state: &mut AppState) -> bool {
        let Some(hit) = self.results.get(self.selected) else {
            return false;
        };
        let command = &self.commands[hit.command_index];
        if !(command.is_enabled)(app_state) {
            return false;
        }

        (command.action)(app_state);
        self.register_recent(command.id);
        self.close();
        true
    }

    pub fn execute_by_id(&mut self, id: &str, app_state: &mut AppState) -> bool {
        if let Some((index, _)) = self
            .commands
            .iter()
            .enumerate()
            .find(|(_, c)| c.id.eq_ignore_ascii_case(id))
        {
            let cmd = &self.commands[index];
            if !(cmd.is_enabled)(app_state) {
                return false;
            }
            (cmd.action)(app_state);
            self.register_recent(cmd.id);
            self.close();
            true
        } else {
            false
        }
    }

    pub fn refresh_results(&mut self, app_state: Option<&AppState>) {
        let query = normalize_query(self.query.as_str());
        let command_query = strip_prefix_query(query.as_str(), self.mode);

        self.results.clear();
        self.grouped_result_headers.clear();

        if command_query.is_empty() {
            self.load_recent_or_all(app_state);
            return;
        }

        for (index, command) in self.commands.iter().enumerate() {
            if let Some(state) = app_state {
                if !(command.is_enabled)(state) {
                    continue;
                }
            }

            let haystack = format!(
                "{} {} {} {}",
                command.label,
                command.category,
                command.shortcut.unwrap_or_default(),
                command.id
            )
            .to_ascii_lowercase();

            if let Some((score, matches)) = fuzzy_score(command_query.as_str(), haystack.as_str()) {
                self.results.push(CommandMatch {
                    command_index: index,
                    score,
                    matched_chars: matches,
                });
            }
        }

        self.results.sort_by(|a, b| b.score.cmp(&a.score));
        self.build_group_headers();
    }

    pub fn results(&self) -> &[CommandMatch] {
        self.results.as_slice()
    }

    pub fn command(&self, index: usize) -> Option<&Command> {
        self.commands.get(index)
    }

    pub fn result_labels(&self, max: usize) -> Vec<String> {
        self.results
            .iter()
            .take(max)
            .filter_map(|hit| self.command(hit.command_index))
            .map(|cmd| cmd.label.to_string())
            .collect()
    }

    fn load_recent_or_all(&mut self, app_state: Option<&AppState>) {
        if !self.recent_ids.is_empty() {
            for recent_id in &self.recent_ids {
                if let Some((idx, cmd)) = self
                    .commands
                    .iter()
                    .enumerate()
                    .find(|(_, c)| c.id == *recent_id)
                {
                    if let Some(state) = app_state {
                        if !(cmd.is_enabled)(state) {
                            continue;
                        }
                    }
                    self.results.push(CommandMatch {
                        command_index: idx,
                        score: 10_000,
                        matched_chars: Vec::new(),
                    });
                }
            }
            if !self.results.is_empty() {
                self.build_group_headers();
                return;
            }
        }

        for (idx, cmd) in self.commands.iter().enumerate() {
            if let Some(state) = app_state {
                if !(cmd.is_enabled)(state) {
                    continue;
                }
            }

            self.results.push(CommandMatch {
                command_index: idx,
                score: 1,
                matched_chars: Vec::new(),
            });
        }
        self.build_group_headers();
    }

    fn build_group_headers(&mut self) {
        let mut seen: HashMap<&'static str, usize> = HashMap::new();
        for (pos, item) in self.results.iter().enumerate() {
            let category = self.commands[item.command_index].category;
            seen.entry(category).or_insert(pos);
        }

        let mut headers = seen
            .into_iter()
            .map(|(category, index)| (category.to_string(), index))
            .collect::<Vec<_>>();
        headers.sort_by_key(|(_, idx)| *idx);
        self.grouped_result_headers = headers;
    }

    fn register_recent(&mut self, id: &'static str) {
        self.recent_ids.retain(|v| *v != id);
        self.recent_ids.push_front(id);
        while self.recent_ids.len() > 12 {
            self.recent_ids.pop_back();
        }
    }

    pub fn palette_rect_in(&self, window: Rect) -> Rect {
        let width = PALETTE_WIDTH.min(window.width - 24.0).max(320.0);
        Rect {
            x: window.x + (window.width - width) * 0.5,
            y: window.y + 20.0,
            width,
            height: PALETTE_MAX_HEIGHT,
        }
    }
}

impl UIComponent for CommandPalette {
    fn layout(&mut self, bounds: Rect, _dpi: f32) {
        self.bounds = self.palette_rect_in(bounds);
    }

    fn render(&self, _ctx: &ID2D1DeviceContext, _theme: &Theme) {
        // Render by host shell in this phase.
    }

    fn handle_input(&mut self, event: &InputEvent) -> bool {
        if !self.visible {
            return false;
        }

        match event {
            InputEvent::KeyDown(vk) => match *vk {
                0x1B => {
                    self.close();
                    true
                }
                0x26 => {
                    self.move_selection(false);
                    true
                }
                0x28 => {
                    self.move_selection(true);
                    true
                }
                0x0D => true,
                0x08 => {
                    self.backspace();
                    true
                }
                _ => false,
            },
            InputEvent::Char(ch) => {
                self.append_char(*ch);
                true
            }
            InputEvent::MouseDown(p) => {
                if self.close_on_click_outside && !contains(self.bounds, *p) {
                    self.close();
                    true
                } else {
                    contains(self.bounds, *p)
                }
            }
            _ => false,
        }
    }

    fn hit_test(&self, point: Point) -> bool {
        self.visible && contains(self.bounds, point)
    }

    fn set_visible(&mut self, visible: bool) {
        self.visible = visible;
        if !visible {
            self.opacity = 0.0;
            self.fade = None;
        }
    }

    fn bounds(&self) -> Rect {
        self.bounds
    }
}

fn default_command_registry() -> Vec<Command> {
    let mut commands = Vec::new();

    let mut push = |id: &'static str,
                    label: &'static str,
                    category: &'static str,
                    shortcut: Option<&'static str>,
                    action: Box<dyn Fn(&mut AppState) + Send + Sync>| {
        commands.push(Command {
            id,
            label,
            category,
            shortcut,
            action,
            is_enabled: Box::new(|_| true),
        });
    };

    push("file.new", "New", "File", Some("Ctrl+N"), Box::new(|state| {
        state.status_text = "New document".to_string();
        state.document = crate::document::model::DocumentModel::default();
    }));
    push("file.open", "Open", "File", Some("Ctrl+O"), Box::new(|state| {
        state.status_text = "Open file".to_string();
    }));
    push("file.save", "Save", "File", Some("Ctrl+S"), Box::new(|state| {
        state.status_text = "Saved".to_string();
        state.document.dirty = false;
    }));
    push("file.save_as", "Save As", "File", Some("Ctrl+Shift+S"), Box::new(|state| {
        state.status_text = "Save As".to_string();
    }));
    push("file.export_pdf", "Export as PDF", "File", None, Box::new(|state| {
        state.status_text = "Export PDF".to_string();
    }));
    push("file.print", "Print", "File", Some("Ctrl+P"), Box::new(|state| {
        state.status_text = "Print".to_string();
    }));

    push("edit.undo", "Undo", "Edit", Some("Ctrl+Z"), Box::new(|state| {
        state.status_text = "Undo".to_string();
    }));
    push("edit.redo", "Redo", "Edit", Some("Ctrl+Y"), Box::new(|state| {
        state.status_text = "Redo".to_string();
    }));
    push("edit.find", "Find", "Edit", Some("Ctrl+F"), Box::new(|state| {
        state.status_text = "Find".to_string();
    }));
    push("edit.replace", "Replace", "Edit", Some("Ctrl+H"), Box::new(|state| {
        state.status_text = "Replace".to_string();
    }));

    push("format.bold", "Bold", "Format", Some("Ctrl+B"), Box::new(|state| {
        state.status_text = "Bold".to_string();
    }));
    push("format.italic", "Italic", "Format", Some("Ctrl+I"), Box::new(|state| {
        state.status_text = "Italic".to_string();
    }));
    push("format.underline", "Underline", "Format", Some("Ctrl+U"), Box::new(|state| {
        state.status_text = "Underline".to_string();
    }));
    push("format.clear", "Clear Formatting", "Format", Some("Ctrl+\\"), Box::new(|state| {
        state.status_text = "Clear formatting".to_string();
    }));
    push("format.heading_1", "Heading 1", "Format", None, Box::new(|state| {
        state.status_text = "Heading 1".to_string();
    }));

    push("insert.image", "Insert Image", "Insert", None, Box::new(|state| {
        state.status_text = "Insert image".to_string();
    }));
    push("insert.link", "Insert Link", "Insert", None, Box::new(|state| {
        state.status_text = "Insert link".to_string();
    }));
    push("insert.table", "Insert Table", "Insert", None, Box::new(|state| {
        state.status_text = "Insert table".to_string();
    }));

    push("view.zoom_in", "Zoom In", "View", Some("Ctrl++"), Box::new(|state| {
        state.status_text = "Zoom in".to_string();
    }));
    push("view.font_size_increase", "Font Size Increase", "View", Some("Ctrl+Shift+>"), Box::new(|state| {
        state.status_text = "Font size increased".to_string();
    }));
    push("view.font_size_decrease", "Font Size Decrease", "View", Some("Ctrl+Shift+<"), Box::new(|state| {
        state.status_text = "Font size decreased".to_string();
    }));
    push("view.zoom_out", "Zoom Out", "View", Some("Ctrl+-"), Box::new(|state| {
        state.status_text = "Zoom out".to_string();
    }));
    push("view.zoom_reset", "Reset Zoom", "View", Some("Ctrl+0"), Box::new(|state| {
        state.status_text = "Zoom 100%".to_string();
    }));
    push("view.toggle_sidebar", "Toggle Sidebar", "View", Some("Ctrl+B"), Box::new(|state| {
        state.show_sidebar = !state.show_sidebar;
    }));
    push("view.toggle_toolbar", "Toggle Toolbar", "View", None, Box::new(|state| {
        state.show_toolbar = !state.show_toolbar;
    }));
    push("view.toggle_statusbar", "Toggle Status Bar", "View", None, Box::new(|state| {
        state.show_statusbar = !state.show_statusbar;
    }));
    push("view.settings", "Open Settings", "View", Some("Ctrl+,"), Box::new(|state| {
        state.show_settings = true;
        state.status_text = "Settings".to_string();
    }));
    push("view.debug_panel", "Toggle Debug Panel", "View", Some("Ctrl+Shift+D"), Box::new(|state| {
        state.show_debug_panel = !state.show_debug_panel;
    }));
    push("view.fit_width", "Fit Width", "View", None, Box::new(|state| {
        state.status_text = "Fit width".to_string();
    }));
    push("view.fit_page", "Fit Page", "View", None, Box::new(|state| {
        state.status_text = "Fit page".to_string();
    }));
    push("view.single_page", "Single Page Mode", "View", None, Box::new(|state| {
        state.status_text = "Single page mode".to_string();
    }));
    push("view.continuous", "Continuous Mode", "View", None, Box::new(|state| {
        state.status_text = "Continuous mode".to_string();
    }));
    push("view.read_mode", "Read Mode", "View", None, Box::new(|state| {
        state.status_text = "Read mode".to_string();
    }));
    push("view.fullscreen", "Toggle Full Screen", "View", Some("F11"), Box::new(|state| {
        state.status_text = "Toggle fullscreen".to_string();
    }));
    push("view.focus_mode", "Toggle Focus Mode", "View", None, Box::new(|state| {
        state.status_text = "Toggle focus mode".to_string();
    }));

    push("theme.switch", "Switch Theme", "Theme", None, Box::new(|state| {
        state.status_text = "Switch theme".to_string();
    }));
    push("theme.toggle_dark_mode", "Toggle Dark Mode", "Theme", None, Box::new(|state| {
        state.status_text = "Toggle dark mode".to_string();
    }));
    push("background.change", "Change Background", "Background", None, Box::new(|state| {
        state.status_text = "Change background".to_string();
    }));
    push("document.word_count", "Word Count", "Document", None, Box::new(|state| {
        state.status_text = "Word count".to_string();
    }));
    push("document.properties", "Document Properties", "Document", None, Box::new(|state| {
        state.status_text = "Document properties".to_string();
    }));
    push("document.goto_page", "Go to Page", "Document", None, Box::new(|state| {
        state.status_text = "Go to page".to_string();
    }));

    push("file.open_folder", "Open Folder", "File", Some("Ctrl+K Ctrl+O"), Box::new(|state| {
        state.status_text = "Open folder".to_string();
    }));
    push("file.close_tab", "Close Tab", "File", Some("Ctrl+W"), Box::new(|state| {
        state.status_text = "Close tab".to_string();
    }));
    push("file.close_window", "Close Window", "File", Some("Alt+F4"), Box::new(|state| {
        state.status_text = "Close window".to_string();
    }));

    push("edit.cut", "Cut", "Edit", Some("Ctrl+X"), Box::new(|state| {
        state.status_text = "Cut".to_string();
    }));
    push("edit.copy", "Copy", "Edit", Some("Ctrl+C"), Box::new(|state| {
        state.status_text = "Copy".to_string();
    }));
    push("edit.paste", "Paste", "Edit", Some("Ctrl+V"), Box::new(|state| {
        state.status_text = "Paste".to_string();
    }));
    push("edit.paste_plain", "Paste Plain", "Edit", Some("Ctrl+Shift+V"), Box::new(|state| {
        state.status_text = "Paste plain".to_string();
    }));
    push("edit.select_all", "Select All", "Edit", Some("Ctrl+A"), Box::new(|state| {
        state.status_text = "Select all".to_string();
    }));

    push("format.strikethrough", "Strikethrough", "Format", Some("Ctrl+Shift+X"), Box::new(|state| {
        state.status_text = "Strikethrough".to_string();
    }));
    push("format.superscript", "Superscript", "Format", Some("Ctrl+Shift+="), Box::new(|state| {
        state.status_text = "Superscript".to_string();
    }));
    push("format.subscript", "Subscript", "Format", Some("Ctrl+="), Box::new(|state| {
        state.status_text = "Subscript".to_string();
    }));
    push("format.heading_2", "Heading 2", "Format", None, Box::new(|state| {
        state.status_text = "Heading 2".to_string();
    }));
    push("format.heading_3", "Heading 3", "Format", None, Box::new(|state| {
        state.status_text = "Heading 3".to_string();
    }));
    push("format.heading_4", "Heading 4", "Format", None, Box::new(|state| {
        state.status_text = "Heading 4".to_string();
    }));
    push("format.heading_5", "Heading 5", "Format", None, Box::new(|state| {
        state.status_text = "Heading 5".to_string();
    }));
    push("format.heading_6", "Heading 6", "Format", None, Box::new(|state| {
        state.status_text = "Heading 6".to_string();
    }));
    push("format.normal_text", "Normal Text", "Format", None, Box::new(|state| {
        state.status_text = "Normal text".to_string();
    }));
    push("format.bullet_list", "Bullet List", "Format", None, Box::new(|state| {
        state.status_text = "Bullet list".to_string();
    }));
    push("format.numbered_list", "Numbered List", "Format", None, Box::new(|state| {
        state.status_text = "Numbered list".to_string();
    }));
    push("format.increase_indent", "Increase Indent", "Format", None, Box::new(|state| {
        state.status_text = "Increase indent".to_string();
    }));
    push("format.decrease_indent", "Decrease Indent", "Format", None, Box::new(|state| {
        state.status_text = "Decrease indent".to_string();
    }));
    push("format.align_left", "Align Left", "Format", None, Box::new(|state| {
        state.status_text = "Align left".to_string();
    }));
    push("format.align_center", "Align Center", "Format", None, Box::new(|state| {
        state.status_text = "Align center".to_string();
    }));
    push("format.align_right", "Align Right", "Format", None, Box::new(|state| {
        state.status_text = "Align right".to_string();
    }));
    push("format.align_justify", "Align Justify", "Format", None, Box::new(|state| {
        state.status_text = "Align justify".to_string();
    }));
    push("format.line_spacing", "Line Spacing", "Format", None, Box::new(|state| {
        state.status_text = "Line spacing".to_string();
    }));

    push("insert.horizontal_rule", "Horizontal Rule", "Insert", None, Box::new(|state| {
        state.status_text = "Horizontal rule".to_string();
    }));
    push("insert.page_break", "Page Break", "Insert", None, Box::new(|state| {
        state.status_text = "Page break".to_string();
    }));
    push("insert.special_char", "Special Character", "Insert", None, Box::new(|state| {
        state.status_text = "Special character".to_string();
    }));
    push("insert.datetime", "Date/Time", "Insert", None, Box::new(|state| {
        state.status_text = "Date/time".to_string();
    }));

    commands
}

fn normalize_query(query: &str) -> String {
    query.trim().to_ascii_lowercase()
}

fn strip_prefix_query(query: &str, mode: QuickActionMode) -> String {
    match mode {
        QuickActionMode::Command
        | QuickActionMode::GoToHeading
        | QuickActionMode::GoToLineOrPage
        | QuickActionMode::GoToBookmark => query.chars().skip(1).collect(),
        QuickActionMode::SearchDocument => query.to_string(),
    }
}

fn fuzzy_score(needle: &str, haystack: &str) -> Option<(i32, Vec<usize>)> {
    if needle.is_empty() {
        return Some((0, Vec::new()));
    }

    let n = needle.chars().collect::<Vec<_>>();
    let h = haystack.chars().collect::<Vec<_>>();

    let mut ni = 0usize;
    let mut last_match = None;
    let mut score = 0i32;
    let mut matched = Vec::new();

    for (hi, hc) in h.iter().enumerate() {
        if ni >= n.len() {
            break;
        }

        if n[ni] == *hc {
            matched.push(hi);
            score += 6;
            if let Some(prev) = last_match {
                if hi == prev + 1 {
                    score += 4;
                }
            }
            if hi == 0 || h.get(hi.saturating_sub(1)) == Some(&' ') {
                score += 5;
            }
            last_match = Some(hi);
            ni += 1;
        }
    }

    if ni == n.len() {
        score += (40_i32 - (h.len() as i32).min(40)).max(0);
        Some((score, matched))
    } else {
        None
    }
}

fn contains(rect: Rect, point: Point) -> bool {
    point.x >= rect.x
        && point.x <= rect.x + rect.width
        && point.y >= rect.y
        && point.y <= rect.y + rect.height
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn quick_action_mode_prefixes_match_prompt() {
        assert_eq!(QuickActionMode::from_query(">save"), QuickActionMode::Command);
        assert_eq!(QuickActionMode::from_query("#h1"), QuickActionMode::GoToHeading);
        assert_eq!(QuickActionMode::from_query(":12"), QuickActionMode::GoToLineOrPage);
        assert_eq!(QuickActionMode::from_query("@intro"), QuickActionMode::GoToBookmark);
        assert_eq!(QuickActionMode::from_query("needle"), QuickActionMode::SearchDocument);
    }

    #[test]
    fn fuzzy_search_matches_font_size_increase() {
        let mut palette = CommandPalette::new();
        palette.set_query(">fsi");
        let labels = palette.result_labels(8);
        assert!(labels.iter().any(|label| label == "Font Size Increase"));
    }

    #[test]
    fn empty_query_prefers_recent_commands() {
        let mut palette = CommandPalette::new();
        let mut app_state = AppState::default();
        palette.set_query(">save");
        assert!(palette.execute_selected(&mut app_state));
        palette.open();
        let first = palette
            .results()
            .first()
            .and_then(|hit| palette.command(hit.command_index))
            .map(|cmd| cmd.id);
        assert_eq!(first, Some("file.save"));
    }

    #[test]
    fn group_headers_exist_when_results_span_categories() {
        let mut palette = CommandPalette::new();
        palette.set_query(">toggle");
        assert!(!palette.grouped_result_headers.is_empty());
    }
}
