use std::time::{Duration, Instant};

use windows::Win32::Graphics::Direct2D::ID2D1DeviceContext;

use crate::{
    theme::Theme,
    ui::{Point, Rect, UIComponent},
};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ToolbarButtonType {
    Icon,
    Toggle,
    Dropdown,
    Split,
    Separator,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ToolbarAction {
    FileMenu,
    Cut,
    Copy,
    Paste,
    Undo,
    Redo,
    Bold,
    Italic,
    Underline,
    Strikethrough,
    FontFamily,
    FontSize,
    TextColor,
    AlignLeft,
    AlignCenter,
    AlignRight,
    AlignJustify,
    List,
    Heading,
    InsertImage,
    InsertLink,
    InsertTable,
    CommandPalette,
    More,
}

#[derive(Debug, Clone)]
pub struct ToolbarButton {
    pub id: &'static str,
    pub label: &'static str,
    pub tooltip: &'static str,
    pub icon_glyph: &'static str,
    pub action: ToolbarAction,
    pub kind: ToolbarButtonType,
    pub width: f32,
    pub enabled: bool,
    pub active: bool,
    pub indeterminate: bool,
}

#[derive(Debug, Clone, Default)]
pub struct ToolbarFormatState {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub font_family: String,
    pub font_size: f32,
}

#[derive(Debug, Clone)]
pub struct Toolbar {
    bounds: Rect,
    visible: bool,
    pub buttons: Vec<ToolbarButton>,
    pub overflow: Vec<ToolbarButton>,
    pub hovered_index: Option<usize>,
    hover_started: Option<Instant>,
    pub show_tooltip: bool,
    pub format_state: ToolbarFormatState,
}

impl Default for Toolbar {
    fn default() -> Self {
        Self::new()
    }
}

impl Toolbar {
    pub fn new() -> Self {
        Self {
            bounds: Rect::default(),
            visible: true,
            buttons: default_buttons(),
            overflow: Vec::new(),
            hovered_index: None,
            hover_started: None,
            show_tooltip: false,
            format_state: ToolbarFormatState {
                font_family: "Segoe UI".to_string(),
                font_size: 12.0,
                ..ToolbarFormatState::default()
            },
        }
    }

    pub fn recalc_overflow(&mut self, available_width: f32) {
        self.overflow.clear();
        let mut used = 8.0_f32;
        let mut kept = Vec::with_capacity(self.buttons.len());

        for button in &self.buttons {
            if button.kind == ToolbarButtonType::Separator {
                used += 8.0;
                kept.push(button.clone());
                continue;
            }

            if used + button.width + 36.0 > available_width {
                self.overflow.push(button.clone());
            } else {
                used += button.width;
                kept.push(button.clone());
            }
        }

        self.buttons = kept;
    }

    pub fn set_format_state(&mut self, state: ToolbarFormatState) {
        self.format_state = state;

        for button in &mut self.buttons {
            match button.action {
                ToolbarAction::Bold => button.active = self.format_state.bold,
                ToolbarAction::Italic => button.active = self.format_state.italic,
                ToolbarAction::Underline => button.active = self.format_state.underline,
                ToolbarAction::Strikethrough => button.active = self.format_state.strikethrough,
                _ => {}
            }
        }
    }

    pub fn hit_button(&self, point: Point) -> Option<usize> {
        if !self.visible {
            return None;
        }

        let mut x = self.bounds.x + 8.0;
        for (idx, b) in self.buttons.iter().enumerate() {
            let w = if b.kind == ToolbarButtonType::Separator {
                8.0
            } else {
                b.width
            };
            let rect = Rect {
                x,
                y: self.bounds.y + 6.0,
                width: w,
                height: self.bounds.height - 12.0,
            };
            if contains(rect, point) {
                return Some(idx);
            }
            x += w;
        }

        None
    }

    pub fn begin_hover(&mut self, index: Option<usize>) {
        if self.hovered_index != index {
            self.hover_started = Some(Instant::now());
            self.show_tooltip = false;
            self.hovered_index = index;
        } else if let Some(t0) = self.hover_started {
            self.show_tooltip = t0.elapsed() >= Duration::from_millis(500);
        }
    }

    pub fn invoke(&self, index: usize) -> Option<ToolbarAction> {
        self.buttons
            .get(index)
            .filter(|b| b.enabled)
            .map(|b| b.action)
    }
}

impl UIComponent for Toolbar {
    fn layout(&mut self, bounds: Rect, _dpi: f32) {
        self.bounds = bounds;
        self.recalc_overflow(bounds.width);
    }

    fn render(&self, _ctx: &ID2D1DeviceContext, _theme: &Theme) {
        // Rendered by global D2D shell for now.
    }

    fn handle_input(&mut self, event: &crate::ui::InputEvent) -> bool {
        match event {
            crate::ui::InputEvent::MouseMove(p) => {
                let idx = self.hit_button(*p);
                self.begin_hover(idx);
                idx.is_some()
            }
            crate::ui::InputEvent::MouseDown(p) => self.hit_button(*p).is_some(),
            _ => false,
        }
    }

    fn hit_test(&self, point: Point) -> bool {
        contains(self.bounds, point)
    }

    fn set_visible(&mut self, visible: bool) {
        self.visible = visible;
    }

    fn bounds(&self) -> Rect {
        self.bounds
    }
}

fn default_buttons() -> Vec<ToolbarButton> {
    vec![
        btn("file", "File", "File menu", "", ToolbarAction::FileMenu, ToolbarButtonType::Dropdown, 64.0),
        btn("cut", "Cut", "Cut", "", ToolbarAction::Cut, ToolbarButtonType::Icon, 36.0),
        btn("copy", "Copy", "Copy", "", ToolbarAction::Copy, ToolbarButtonType::Icon, 36.0),
        btn("paste", "Paste", "Paste", "", ToolbarAction::Paste, ToolbarButtonType::Split, 44.0),
        btn("undo", "Undo", "Undo", "", ToolbarAction::Undo, ToolbarButtonType::Icon, 36.0),
        btn("redo", "Redo", "Redo", "", ToolbarAction::Redo, ToolbarButtonType::Icon, 36.0),
        sep(),
        btn("bold", "B", "Bold", "B", ToolbarAction::Bold, ToolbarButtonType::Toggle, 36.0),
        btn("italic", "I", "Italic", "I", ToolbarAction::Italic, ToolbarButtonType::Toggle, 36.0),
        btn("underline", "U", "Underline", "U", ToolbarAction::Underline, ToolbarButtonType::Toggle, 36.0),
        btn("strike", "S", "Strikethrough", "S", ToolbarAction::Strikethrough, ToolbarButtonType::Toggle, 36.0),
        sep(),
        btn("font", "Font", "Font family", "A", ToolbarAction::FontFamily, ToolbarButtonType::Dropdown, 120.0),
        btn("size", "Size", "Font size", "12", ToolbarAction::FontSize, ToolbarButtonType::Dropdown, 68.0),
        btn("color", "Color", "Text color", "", ToolbarAction::TextColor, ToolbarButtonType::Dropdown, 64.0),
        sep(),
        btn("align_left", "Left", "Align left", "", ToolbarAction::AlignLeft, ToolbarButtonType::Toggle, 36.0),
        btn("align_center", "Center", "Align center", "", ToolbarAction::AlignCenter, ToolbarButtonType::Toggle, 36.0),
        btn("align_right", "Right", "Align right", "", ToolbarAction::AlignRight, ToolbarButtonType::Toggle, 36.0),
        btn("justify", "Justify", "Align justify", "", ToolbarAction::AlignJustify, ToolbarButtonType::Toggle, 36.0),
        sep(),
        btn("list", "List", "Lists", "", ToolbarAction::List, ToolbarButtonType::Dropdown, 58.0),
        btn("heading", "Heading", "Heading styles", "", ToolbarAction::Heading, ToolbarButtonType::Dropdown, 72.0),
        btn("image", "Image", "Insert image", "", ToolbarAction::InsertImage, ToolbarButtonType::Icon, 36.0),
        btn("link", "Link", "Insert link", "", ToolbarAction::InsertLink, ToolbarButtonType::Icon, 36.0),
        btn("table", "Table", "Insert table", "", ToolbarAction::InsertTable, ToolbarButtonType::Icon, 36.0),
        btn("cmd", "Cmd", "Command palette", "", ToolbarAction::CommandPalette, ToolbarButtonType::Icon, 36.0),
        btn("more", "More", "More actions", "…", ToolbarAction::More, ToolbarButtonType::Dropdown, 40.0),
    ]
}

fn btn(
    id: &'static str,
    label: &'static str,
    tooltip: &'static str,
    icon_glyph: &'static str,
    action: ToolbarAction,
    kind: ToolbarButtonType,
    width: f32,
) -> ToolbarButton {
    ToolbarButton {
        id,
        label,
        tooltip,
        icon_glyph,
        action,
        kind,
        width,
        enabled: true,
        active: false,
        indeterminate: false,
    }
}

fn sep() -> ToolbarButton {
    ToolbarButton {
        id: "sep",
        label: "",
        tooltip: "",
        icon_glyph: "",
        action: ToolbarAction::More,
        kind: ToolbarButtonType::Separator,
        width: 8.0,
        enabled: false,
        active: false,
        indeterminate: false,
    }
}

fn contains(rect: Rect, point: Point) -> bool {
    point.x >= rect.x
        && point.x <= rect.x + rect.width
        && point.y >= rect.y
        && point.y <= rect.y + rect.height
}
