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
const TAB_GAP: f32 = 6.0;
const TAB_BAR_PADDING: f32 = 8.0;
const NEW_TAB_BUTTON_WIDTH: f32 = 28.0;
const OVERFLOW_BUTTON_WIDTH: f32 = 24.0;

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
    pub close_rects: Vec<Rect>,
    pub new_tab_rect: Rect,
    pub overflow_left_rect: Rect,
    pub overflow_right_rect: Rect,
    pub hovered: Option<usize>,
    dragging_tab: Option<usize>,
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
            close_rects: Vec::new(),
            new_tab_rect: Rect::default(),
            overflow_left_rect: Rect::default(),
            overflow_right_rect: Rect::default(),
            hovered: None,
            dragging_tab: None,
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
        self.ensure_active_visible();
        self.recalc_tab_layout();
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
        self.ensure_active_visible();
        self.recalc_tab_layout();
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

        if self.active > index {
            self.active = self.active.saturating_sub(1);
        } else if self.active >= self.tabs.len() {
            self.active = self.tabs.len().saturating_sub(1);
        }

        self.ensure_active_visible();
        self.recalc_tab_layout();
        true
    }

    pub fn close_tabs_to_right(&mut self, index: usize) {
        if index + 1 < self.tabs.len() {
            self.tabs.truncate(index + 1);
            self.active = self.active.min(self.tabs.len().saturating_sub(1));
            self.ensure_active_visible();
            self.recalc_tab_layout();
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
        self.ensure_active_visible();
        self.recalc_tab_layout();
    }

    pub fn set_active(&mut self, index: usize) {
        self.active = index.min(self.tabs.len().saturating_sub(1));
        self.ensure_active_visible();
        self.recalc_tab_layout();
    }

    pub fn switch_next(&mut self) {
        if self.tabs.is_empty() {
            return;
        }
        self.active = (self.active + 1) % self.tabs.len();
        self.ensure_active_visible();
        self.recalc_tab_layout();
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
        self.ensure_active_visible();
        self.recalc_tab_layout();
    }

    pub fn switch_to_number(&mut self, one_based: usize) {
        if one_based == 0 {
            return;
        }
        let index = one_based - 1;
        if index < self.tabs.len() {
            self.active = index;
            self.ensure_active_visible();
            self.recalc_tab_layout();
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

        self.ensure_active_visible();
        self.recalc_tab_layout();
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

    pub fn tab_close_hit_test(&self, point: Point) -> Option<usize> {
        self.close_rects
            .iter()
            .enumerate()
            .find(|(_, rect)| contains(**rect, point))
            .map(|(idx, _)| idx + self.overflow_offset)
    }

    pub fn new_button_hit_test(&self, point: Point) -> bool {
        contains(self.new_tab_rect, point)
    }

    pub fn overflow_left_hit_test(&self, point: Point) -> bool {
        self.overflow_left_rect.width > 0.0 && contains(self.overflow_left_rect, point)
    }

    pub fn overflow_right_hit_test(&self, point: Point) -> bool {
        self.overflow_right_rect.width > 0.0 && contains(self.overflow_right_rect, point)
    }

    pub fn is_tab_bar_hit(&self, point: Point) -> bool {
        self.visible && contains(self.bounds, point)
    }

    pub fn is_empty_tab_bar_space(&self, point: Point) -> bool {
        self.is_tab_bar_hit(point)
            && self.tab_hit_test(point).is_none()
            && self.tab_close_hit_test(point).is_none()
            && !self.new_button_hit_test(point)
            && !self.overflow_left_hit_test(point)
            && !self.overflow_right_hit_test(point)
    }

    pub fn scroll_overflow_left(&mut self) -> bool {
        if self.overflow_offset == 0 {
            return false;
        }
        self.overflow_offset = self.overflow_offset.saturating_sub(1);
        self.recalc_tab_layout();
        true
    }

    pub fn scroll_overflow_right(&mut self) -> bool {
        let visible = self.tab_rects.len().max(1);
        if self.overflow_offset + visible >= self.tabs.len() {
            return false;
        }
        self.overflow_offset += 1;
        self.recalc_tab_layout();
        true
    }

    pub fn middle_click_close(&mut self, point: Point) -> bool {
        if let Some(index) = self.tab_hit_test(point) {
            return self.close_tab(index);
        }
        false
    }

    fn ensure_active_visible(&mut self) {
        if self.tabs.is_empty() {
            self.overflow_offset = 0;
            return;
        }
        let visible = self.max_visible_tabs.max(1);
        if self.active < self.overflow_offset {
            self.overflow_offset = self.active;
        } else if self.active >= self.overflow_offset + visible {
            self.overflow_offset = self.active + 1 - visible;
        }
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
        self.close_rects.clear();
        self.new_tab_rect = Rect::default();
        self.overflow_left_rect = Rect::default();
        self.overflow_right_rect = Rect::default();

        if self.bounds.width <= 0.0 {
            return;
        }

        let left_edge = self.bounds.x + TAB_BAR_PADDING;
        let right_edge = (self.bounds.x + self.bounds.width - TAB_BAR_PADDING).max(left_edge);
        self.new_tab_rect = Rect {
            x: (right_edge - NEW_TAB_BUTTON_WIDTH).max(left_edge),
            y: self.bounds.y + 5.0,
            width: NEW_TAB_BUTTON_WIDTH,
            height: TAB_HEIGHT - 10.0,
        };

        let mut tabs_left = left_edge;
        let tabs_right = (self.new_tab_rect.x - TAB_GAP).max(tabs_left);

        let available_without_overflow = (tabs_right - tabs_left).max(0.0);
        let max_without_overflow = ((available_without_overflow + TAB_GAP) / (TAB_MIN_WIDTH + TAB_GAP))
            .floor() as usize;
        let overflow_needed = self.tabs.len() > max_without_overflow.max(1);

        if overflow_needed {
            self.overflow_left_rect = Rect {
                x: tabs_left,
                y: self.bounds.y + 6.0,
                width: OVERFLOW_BUTTON_WIDTH,
                height: TAB_HEIGHT - 12.0,
            };
            self.overflow_right_rect = Rect {
                x: tabs_left + OVERFLOW_BUTTON_WIDTH + 4.0,
                y: self.bounds.y + 6.0,
                width: OVERFLOW_BUTTON_WIDTH,
                height: TAB_HEIGHT - 12.0,
            };
            tabs_left = self.overflow_right_rect.x + self.overflow_right_rect.width + TAB_GAP;
        } else {
            self.overflow_offset = 0;
        }

        let available = (tabs_right - tabs_left).max(0.0);
        self.max_visible_tabs = ((available + TAB_GAP) / (TAB_MIN_WIDTH + TAB_GAP)).floor() as usize;
        self.max_visible_tabs = self.max_visible_tabs.max(1);

        if self.overflow_offset + self.max_visible_tabs > self.tabs.len() {
            self.overflow_offset = self.tabs.len().saturating_sub(self.max_visible_tabs);
        }

        let visible_count = self
            .tabs
            .len()
            .saturating_sub(self.overflow_offset)
            .min(self.max_visible_tabs);
        if visible_count == 0 {
            return;
        }

        let total_gap = TAB_GAP * visible_count.saturating_sub(1) as f32;
        let tab_width = ((available - total_gap) / visible_count as f32).clamp(TAB_MIN_WIDTH, TAB_MAX_WIDTH);
        let mut x = tabs_left;

        for _ in 0..visible_count {
            self.tab_rects.push(Rect {
                x,
                y: self.bounds.y,
                width: tab_width,
                height: TAB_HEIGHT,
            });
            self.close_rects.push(Rect {
                x: (x + tab_width - 18.0).max(x + 4.0),
                y: self.bounds.y + 10.0,
                width: 12.0,
                height: 12.0,
            });
            x += tab_width + TAB_GAP;
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
                let mut changed = false;
                let hovered = self.tab_hit_test(*point);
                if self.hovered != hovered {
                    self.hovered = hovered;
                    changed = true;
                }
                if let Some(dragging) = self.dragging_tab
                    && let Some(target) = self.tab_hit_test(*point)
                    && target != dragging
                    && self.reorder_tab(dragging, target)
                {
                    self.dragging_tab = Some(target);
                    changed = true;
                }
                changed
            }
            InputEvent::MouseDown(point) => {
                if self.overflow_left_hit_test(*point) {
                    return self.scroll_overflow_left();
                }
                if self.overflow_right_hit_test(*point) {
                    return self.scroll_overflow_right();
                }
                if self.new_button_hit_test(*point) {
                    self.new_blank_tab();
                    return true;
                }
                if let Some(index) = self.tab_hit_test(*point) {
                    self.set_active(index);
                    self.dragging_tab = Some(index);
                    return true;
                }
                false
            }
            InputEvent::MouseUp(_) => {
                if self.dragging_tab.take().is_some() {
                    true
                } else {
                    false
                }
            }
            InputEvent::KeyDown(_) => false,
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn layout_exposes_new_tab_and_overflow_controls() {
        let mut tabs = TabsBar::new();
        for _ in 0..8 {
            tabs.new_blank_tab();
        }
        tabs.layout(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 640.0,
                height: TAB_HEIGHT,
            },
            96.0,
        );

        assert!(tabs.new_tab_rect.width > 0.0);
        assert!(!tabs.tab_rects.is_empty());
        if tabs.tabs.len() > tabs.max_visible_tabs {
            assert!(tabs.overflow_right_rect.width > 0.0);
        }
    }

    #[test]
    fn drag_reorders_tabs() {
        let mut tabs = TabsBar::new();
        tabs.new_blank_tab();
        tabs.new_blank_tab();
        tabs.new_blank_tab();
        tabs.layout(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 820.0,
                height: TAB_HEIGHT,
            },
            96.0,
        );

        let first = tabs.tabs[0].id;
        let p0 = Point {
            x: tabs.tab_rects[0].x + 12.0,
            y: tabs.tab_rects[0].y + 12.0,
        };
        let p2 = Point {
            x: tabs.tab_rects[2].x + 12.0,
            y: tabs.tab_rects[2].y + 12.0,
        };

        let _ = tabs.handle_input(&InputEvent::MouseDown(p0));
        let _ = tabs.handle_input(&InputEvent::MouseMove(p2));
        let _ = tabs.handle_input(&InputEvent::MouseUp(p2));

        assert_ne!(tabs.tabs[0].id, first);
        assert!(tabs.tabs.iter().any(|tab| tab.id == first));
    }
}
