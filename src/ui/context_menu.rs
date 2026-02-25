use windows::Win32::Graphics::Direct2D::ID2D1DeviceContext;

use crate::{
    render::animation::{Animation, Easing},
    theme::Theme,
    ui::{InputEvent, Point, Rect, UIComponent},
};

const MENU_WIDTH: f32 = 260.0;
const MENU_ITEM_HEIGHT: f32 = 30.0;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ContextMenuKind {
    Canvas,
    Tab,
    Sidebar,
    Image,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ContextAction {
    Cut,
    Copy,
    Paste,
    PastePlainText,
    FontDialog,
    ParagraphDialog,
    InsertImage,
    InsertLink,
    InsertTable,
    SelectAll,
    CloseTab,
    CloseOthers,
    CloseAll,
    CloseToRight,
    CopyFilePath,
    ShowInExplorer,
    SaveImageAs,
    ImageProperties,
    BringToFront,
    SendToBack,
}

#[derive(Debug, Clone)]
pub struct ContextMenuItem {
    pub label: &'static str,
    pub action: ContextAction,
    pub enabled: bool,
}

#[derive(Debug, Default)]
pub struct ContextMenu {
    bounds: Rect,
    visible: bool,
    pub reduce_motion: bool,
    pub kind: Option<ContextMenuKind>,
    pub items: Vec<ContextMenuItem>,
    pub selected_index: Option<usize>,
    fade_anim: Option<Animation>,
    scale_anim: Option<Animation>,
    pub opacity: f32,
    pub scale: f32,
}

impl ContextMenu {
    pub fn open(&mut self, kind: ContextMenuKind, origin: Point) {
        self.kind = Some(kind);
        self.items = default_items(kind);
        self.bounds = Rect {
            x: origin.x,
            y: origin.y,
            width: MENU_WIDTH,
            height: MENU_ITEM_HEIGHT * self.items.len() as f32,
        };
        self.visible = true;
        self.selected_index = None;
        if self.reduce_motion {
            self.opacity = 1.0;
            self.scale = 1.0;
            self.fade_anim = None;
            self.scale_anim = None;
        } else {
            self.fade_anim = Some(Animation::new(0.0, 1.0, 0.08, Easing::EaseOutCubic));
            self.scale_anim = Some(Animation::new(0.95, 1.0, 0.08, Easing::EaseOutCubic));
        }
    }

    pub fn close(&mut self) {
        self.visible = false;
        self.kind = None;
        self.selected_index = None;
        self.items.clear();
        self.fade_anim = None;
        self.scale_anim = None;
        self.opacity = 0.0;
        self.scale = 0.95;
    }

    pub fn tick(&mut self, dt_s: f32) {
        if let Some(anim) = &mut self.fade_anim {
            if anim.update_respecting_motion_pref(dt_s, self.reduce_motion) {
                self.opacity = anim.current_value.clamp(0.0, 1.0);
            } else {
                self.opacity = anim.end_value.clamp(0.0, 1.0);
                self.fade_anim = None;
            }
        }

        if let Some(anim) = &mut self.scale_anim {
            if anim.update_respecting_motion_pref(dt_s, self.reduce_motion) {
                self.scale = anim.current_value;
            } else {
                self.scale = anim.end_value;
                self.scale_anim = None;
            }
        }
    }

    pub fn selected_action(&self) -> Option<ContextAction> {
        self.selected_index
            .and_then(|index| self.items.get(index))
            .filter(|item| item.enabled)
            .map(|item| item.action)
    }

    fn item_rect(&self, index: usize) -> Rect {
        Rect {
            x: self.bounds.x,
            y: self.bounds.y + MENU_ITEM_HEIGHT * index as f32,
            width: self.bounds.width,
            height: MENU_ITEM_HEIGHT,
        }
    }
}

impl UIComponent for ContextMenu {
    fn layout(&mut self, bounds: Rect, _dpi: f32) {
        self.bounds = bounds;
    }

    fn render(&self, _ctx: &ID2D1DeviceContext, _theme: &Theme) {
        // Drawn in host shell renderer.
    }

    fn handle_input(&mut self, event: &InputEvent) -> bool {
        if !self.visible {
            return false;
        }

        match event {
            InputEvent::MouseMove(point) => {
                self.selected_index = self
                    .items
                    .iter()
                    .enumerate()
                    .find(|(index, _)| contains(self.item_rect(*index), *point))
                    .map(|(index, _)| index);
                self.selected_index.is_some()
            }
            InputEvent::MouseDown(point) => {
                if !contains(self.bounds, *point) {
                    self.close();
                    return true;
                }

                self.selected_index = self
                    .items
                    .iter()
                    .enumerate()
                    .find(|(index, _)| contains(self.item_rect(*index), *point))
                    .map(|(index, _)| index);
                self.selected_action().is_some()
            }
            InputEvent::KeyDown(vk) => match *vk {
                0x1B => {
                    self.close();
                    true
                }
                0x26 => {
                    let selected = self.selected_index.unwrap_or(0).saturating_sub(1);
                    self.selected_index = Some(selected);
                    true
                }
                0x28 => {
                    let selected = (self.selected_index.unwrap_or(0) + 1).min(self.items.len().saturating_sub(1));
                    self.selected_index = Some(selected);
                    true
                }
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

fn default_items(kind: ContextMenuKind) -> Vec<ContextMenuItem> {
    let mut entries = Vec::new();

    let mut push = |label: &'static str, action: ContextAction| {
        entries.push(ContextMenuItem {
            label,
            action,
            enabled: true,
        });
    };

    match kind {
        ContextMenuKind::Canvas => {
            push("Cut", ContextAction::Cut);
            push("Copy", ContextAction::Copy);
            push("Paste", ContextAction::Paste);
            push("Paste Plain Text", ContextAction::PastePlainText);
            push("Font...", ContextAction::FontDialog);
            push("Paragraph...", ContextAction::ParagraphDialog);
            push("Insert Image", ContextAction::InsertImage);
            push("Insert Link", ContextAction::InsertLink);
            push("Insert Table", ContextAction::InsertTable);
            push("Select All", ContextAction::SelectAll);
        }
        ContextMenuKind::Tab => {
            push("Close", ContextAction::CloseTab);
            push("Close Others", ContextAction::CloseOthers);
            push("Close All", ContextAction::CloseAll);
            push("Close to the Right", ContextAction::CloseToRight);
            push("Copy File Path", ContextAction::CopyFilePath);
            push("Show in Explorer", ContextAction::ShowInExplorer);
        }
        ContextMenuKind::Sidebar => {
            push("Copy File Path", ContextAction::CopyFilePath);
            push("Show in Explorer", ContextAction::ShowInExplorer);
        }
        ContextMenuKind::Image => {
            push("Cut", ContextAction::Cut);
            push("Copy", ContextAction::Copy);
            push("Save Image As...", ContextAction::SaveImageAs);
            push("Image Properties...", ContextAction::ImageProperties);
            push("Bring to Front", ContextAction::BringToFront);
            push("Send to Back", ContextAction::SendToBack);
        }
    }

    entries
}

fn contains(rect: Rect, point: Point) -> bool {
    point.x >= rect.x
        && point.x <= rect.x + rect.width
        && point.y >= rect.y
        && point.y <= rect.y + rect.height
}
