use std::path::PathBuf;

use windows::Win32::Graphics::Direct2D::ID2D1DeviceContext;

use crate::{
    document::model::DocumentModel,
    editor::cursor::CursorState,
    render::canvas::CanvasState,
    theme::Theme,
    ui::{InputEvent, Point, Rect, UIComponent},
};

const TAB_HEIGHT: f32 = 36.0;
const TAB_MIN_WIDTH: f32 = 140.0;
const TAB_MAX_WIDTH: f32 = 260.0;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TabKind {
    Document,
    Welcome,
}

#[derive(Debug, Clone)]
pub struct TabState {
    pub id: u64,
    pub title: String,
    pub kind: TabKind,
    pub file_path: Option<PathBuf>,
    pub dirty: bool,
    pub document: DocumentModel,
    pub cursor: CursorState,
    pub canvas: CanvasState,
}

impl TabState {
    pub fn from_document(id: u64, title: String, file_path: Option<PathBuf>, document: DocumentModel) -> Self {
        Self {
            id,
            title,
            kind: TabKind::Document,
            file_path,
            dirty: document.dirty,
            document,
            cursor: CursorState::default(),
            canvas: CanvasState::default(),
        }
    }

    pub fn welcome(id: u64) -> Self {
        Self {
            id,
            title: "Welcome".to_string(),
            kind: TabKind::Welcome,
            file_path: None,
            dirty: false,
            document: DocumentModel::default(),
            cursor: CursorState::default(),
            canvas: CanvasState::default(),
        }
    }
}

#[derive(Debug, Clone)]
pub struct TabsBar {
    bounds: Rect,
    visible: bool,
    pub tabs: Vec<TabState>,
    pub active: usize,
    pub overflow_offset: usize,
    pub max_visible_tabs: usize,
    pub tab_rects: Vec<Rect>,
    pub hovered: Option<usize>,
    next_id: u64,
}

impl Default for TabsBar {
    fn default() -> Self {
        Self::new()
    }
}

impl TabsBar {
    pub fn new() -> Self {
        let mut this = Self {
            bounds: Rect::default(),
            visible: true,
            tabs: Vec::new(),
            active: 0,
            overflow_offset: 0,
            max_visible_tabs: 0,
            tab_rects: Vec::new(),
            hovered: None,
            next_id: 1,
        };
        this.ensure_welcome_tab();
        this
    }

    pub fn active_tab(&self) -> Option<&TabState> {
        self.tabs.get(self.active)
    }

    pub fn active_tab_mut(&mut self) -> Option<&mut TabState> {
        self.tabs.get_mut(self.active)
    }

    pub fn new_blank_tab(&mut self) -> usize {
        let id = self.next_id;
        self.next_id += 1;

        let tab = TabState::from_document(
            id,
            format!("Untitled {}", id),
            None,
            DocumentModel::default(),
        );

        self.tabs.push(tab);
        self.active = self.tabs.len() - 1;
        self.remove_welcome_if_needed();
        self.active
    }

    pub fn open_document_tab(
        &mut self,
        title: String,
        path: Option<PathBuf>,
        document: DocumentModel,
    ) -> usize {
        let id = self.next_id;
        self.next_id += 1;
        self.tabs
            .push(TabState::from_document(id, title, path, document));
        self.active = self.tabs.len() - 1;
        self.remove_welcome_if_needed();
        self.active
    }

    pub fn close_active_tab(&mut self) -> bool {
        self.close_tab(self.active)
    }

    pub fn close_tab(&mut self, index: usize) -> bool {
        if index >= self.tabs.len() {
            return false;
        }
        self.tabs.remove(index);

        if self.tabs.is_empty() {
            self.ensure_welcome_tab();
        }

        if self.active >= self.tabs.len() {
            self.active = self.tabs.len().saturating_sub(1);
        }

        true
    }

    pub fn close_tabs_to_right(&mut self, index: usize) {
        if index + 1 < self.tabs.len() {
            self.tabs.truncate(index + 1);
            self.active = self.active.min(self.tabs.len().saturating_sub(1));
        }
    }

    pub fn close_others(&mut self, index: usize) {
        if index >= self.tabs.len() {
            return;
        }

        let keep = self.tabs[index].clone();
        self.tabs.clear();
        self.tabs.push(keep);
        self.active = 0;
    }

    pub fn set_active(&mut self, index: usize) {
        self.active = index.min(self.tabs.len().saturating_sub(1));
    }

    pub fn switch_next(&mut self) {
        if self.tabs.is_empty() {
            return;
        }
        self.active = (self.active + 1) % self.tabs.len();
    }

    pub fn switch_prev(&mut self) {
        if self.tabs.is_empty() {
            return;
        }
        if self.active == 0 {
            self.active = self.tabs.len() - 1;
        } else {
            self.active -= 1;
        }
    }

    pub fn switch_to_number(&mut self, one_based: usize) {
        if one_based == 0 {
            return;
        }
        let index = one_based - 1;
        if index < self.tabs.len() {
            self.active = index;
        }
    }

    pub fn reorder_tab(&mut self, from: usize, to: usize) -> bool {
        if from >= self.tabs.len() || to >= self.tabs.len() || from == to {
            return false;
        }

        let tab = self.tabs.remove(from);
        self.tabs.insert(to, tab);

        if self.active == from {
            self.active = to;
        } else if from < self.active && to >= self.active {
            self.active = self.active.saturating_sub(1);
        } else if from > self.active && to <= self.active {
            self.active += 1;
        }

        true
    }

    pub fn mark_active_dirty(&mut self, dirty: bool) {
        if let Some(tab) = self.active_tab_mut() {
            tab.dirty = dirty;
            tab.document.dirty = dirty;
        }
    }

    pub fn tab_hit_test(&self, point: Point) -> Option<usize> {
        self.tab_rects
            .iter()
            .enumerate()
            .find(|(_, rect)| contains(**rect, point))
            .map(|(idx, _)| idx + self.overflow_offset)
    }

    pub fn middle_click_close(&mut self, point: Point) -> bool {
        if let Some(index) = self.tab_hit_test(point) {
            return self.close_tab(index);
        }
        false
    }

    fn ensure_welcome_tab(&mut self) {
        if self.tabs.is_empty() {
            let id = self.next_id;
            self.next_id += 1;
            self.tabs.push(TabState::welcome(id));
            self.active = 0;
        }
    }

    fn remove_welcome_if_needed(&mut self) {
        if self.tabs.len() <= 1 {
            return;
        }
        self.tabs.retain(|tab| tab.kind != TabKind::Welcome);
        self.active = self.active.min(self.tabs.len().saturating_sub(1));
    }

    fn recalc_tab_layout(&mut self) {
        self.tab_rects.clear();

        let reserved_for_new_btn = 42.0;
        let available = (self.bounds.width - reserved_for_new_btn).max(0.0);
        self.max_visible_tabs = ((available / TAB_MIN_WIDTH).floor() as usize).max(1);

        let visible_count = self
            .tabs
            .len()
            .saturating_sub(self.overflow_offset)
            .min(self.max_visible_tabs);
        if visible_count == 0 {
            return;
        }

        let tab_width = (available / visible_count as f32).clamp(TAB_MIN_WIDTH, TAB_MAX_WIDTH);
        let mut x = self.bounds.x;

        for _ in 0..visible_count {
            self.tab_rects.push(Rect {
                x,
                y: self.bounds.y,
                width: tab_width,
                height: TAB_HEIGHT,
            });
            x += tab_width;
        }
    }
}

impl UIComponent for TabsBar {
    fn layout(&mut self, bounds: Rect, _dpi: f32) {
        self.bounds = Rect {
            x: bounds.x,
            y: bounds.y,
            width: bounds.width,
            height: TAB_HEIGHT,
        };
        self.recalc_tab_layout();
    }

    fn render(&self, _ctx: &ID2D1DeviceContext, _theme: &Theme) {
        // Drawn in global shell renderer.
    }

    fn handle_input(&mut self, event: &InputEvent) -> bool {
        match event {
            InputEvent::MouseMove(point) => {
                self.hovered = self.tab_hit_test(*point);
                self.hovered.is_some()
            }
            InputEvent::MouseDown(point) => {
                if let Some(index) = self.tab_hit_test(*point) {
                    self.set_active(index);
                    return true;
                }
                false
            }
            InputEvent::KeyDown(vk) => match *vk {
                0x54 => {
                    self.new_blank_tab();
                    true
                }
                0x57 => self.close_active_tab(),
                _ => false,
            },
            _ => false,
        }
    }

    fn hit_test(&self, point: Point) -> bool {
        self.visible && contains(self.bounds, point)
    }

    fn set_visible(&mut self, visible: bool) {
        self.visible = visible;
    }

    fn bounds(&self) -> Rect {
        self.bounds
    }
}

fn contains(rect: Rect, point: Point) -> bool {
    point.x >= rect.x
        && point.x <= rect.x + rect.width
        && point.y >= rect.y
        && point.y <= rect.y + rect.height
}
