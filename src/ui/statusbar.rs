use windows::Win32::Graphics::Direct2D::ID2D1DeviceContext;

use crate::{
    theme::Theme,
    ui::{InputEvent, Point, Rect, UIComponent},
};

const STATUSBAR_HEIGHT: f32 = 28.0;
const SEGMENT_PADDING: f32 = 12.0;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StatusAction {
    OpenZoomPopup,
    ChangeEncoding,
}

#[derive(Debug, Clone)]
pub struct StatusBarInfo {
    pub page_index: usize,
    pub page_count: usize,
    pub word_count: usize,
    pub character_count: usize,
    pub view_mode: String,
    pub line: usize,
    pub column: usize,
    pub zoom_percent: u16,
    pub file_format: String,
    pub encoding: String,
}

impl Default for StatusBarInfo {
    fn default() -> Self {
        Self {
            page_index: 1,
            page_count: 1,
            word_count: 0,
            character_count: 0,
            view_mode: "Page".to_string(),
            line: 1,
            column: 1,
            zoom_percent: 100,
            file_format: "DOCX".to_string(),
            encoding: "UTF-8".to_string(),
        }
    }
}

#[derive(Debug, Default)]
pub struct StatusBar {
    bounds: Rect,
    visible: bool,
    pub info: StatusBarInfo,
    pub pending_action: Option<StatusAction>,
}

impl StatusBar {
    pub fn set_info(&mut self, info: StatusBarInfo) {
        self.info = info;
    }

    pub fn left_text(&self) -> String {
        format!(
            "Page {} of {} | Words: {} | Chars: {}",
            self.info.page_index, self.info.page_count, self.info.word_count, self.info.character_count
        )
    }

    pub fn right_text(&self) -> String {
        format!(
            "{} | {}:{} | {}% | {} | {}",
            self.info.view_mode,
            self.info.line,
            self.info.column,
            self.info.zoom_percent,
            self.info.file_format,
            self.info.encoding
        )
    }

    fn zoom_rect(&self) -> Rect {
        Rect {
            x: self.bounds.x + self.bounds.width - 220.0,
            y: self.bounds.y,
            width: 62.0,
            height: self.bounds.height,
        }
    }

    fn encoding_rect(&self) -> Rect {
        Rect {
            x: self.bounds.x + self.bounds.width - 82.0,
            y: self.bounds.y,
            width: 82.0,
            height: self.bounds.height,
        }
    }
}

impl UIComponent for StatusBar {
    fn layout(&mut self, bounds: Rect, _dpi: f32) {
        self.bounds = Rect {
            x: bounds.x,
            y: bounds.y + bounds.height - STATUSBAR_HEIGHT,
            width: bounds.width,
            height: STATUSBAR_HEIGHT,
        };
    }

    fn render(&self, _ctx: &ID2D1DeviceContext, _theme: &Theme) {
        // Drawn in host renderer.
    }

    fn handle_input(&mut self, event: &InputEvent) -> bool {
        if !self.visible {
            return false;
        }

        match event {
            InputEvent::MouseDown(point) => {
                if contains(self.zoom_rect(), *point) {
                    self.pending_action = Some(StatusAction::OpenZoomPopup);
                    return true;
                }
                if contains(self.encoding_rect(), *point) {
                    self.pending_action = Some(StatusAction::ChangeEncoding);
                    return true;
                }
                false
            }
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

#[allow(dead_code)]
fn _layout_hint_example() -> f32 {
    SEGMENT_PADDING
}
