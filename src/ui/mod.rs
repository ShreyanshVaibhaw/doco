use serde::{Deserialize, Serialize};
use windows::Win32::Graphics::Direct2D::ID2D1DeviceContext;

use crate::theme::Theme;

#[derive(Debug, Clone, Copy, Default)]
pub struct Point {
    pub x: f32,
    pub y: f32,
}

#[derive(Debug, Clone, Copy, Default)]
pub struct Size {
    pub width: f32,
    pub height: f32,
}

#[derive(Debug, Clone, Copy, Default)]
pub struct Rect {
    pub x: f32,
    pub y: f32,
    pub width: f32,
    pub height: f32,
}

#[derive(Debug, Clone, Copy, Default, Serialize, Deserialize, PartialEq)]
pub struct Color {
    pub r: f32,
    pub g: f32,
    pub b: f32,
    pub a: f32,
}

impl Color {
    pub const fn rgba(r: f32, g: f32, b: f32, a: f32) -> Self {
        Self { r, g, b, a }
    }

    pub const fn rgb(r: f32, g: f32, b: f32) -> Self {
        Self { r, g, b, a: 1.0 }
    }

    pub fn as_d2d(self) -> windows::Win32::Graphics::Direct2D::Common::D2D1_COLOR_F {
        windows::Win32::Graphics::Direct2D::Common::D2D1_COLOR_F {
            r: self.r,
            g: self.g,
            b: self.b,
            a: self.a,
        }
    }
}

#[derive(Debug, Clone, Copy, Default)]
pub struct Margin {
    pub left: f32,
    pub top: f32,
    pub right: f32,
    pub bottom: f32,
}

#[derive(Debug, Clone, Copy, Default)]
pub struct Padding {
    pub left: f32,
    pub top: f32,
    pub right: f32,
    pub bottom: f32,
}

#[derive(Debug, Clone)]
pub enum InputEvent {
    MouseMove(Point),
    MouseDown(Point),
    MouseUp(Point),
    MouseWheel { delta: f32, position: Point },
    KeyDown(u32),
    KeyUp(u32),
    Char(char),
}

#[derive(Debug, Clone, Copy)]
pub enum LayoutDirection {
    Horizontal,
    Vertical,
}

#[derive(Debug, Clone, Copy)]
pub enum Alignment {
    Start,
    Center,
    End,
    Stretch,
}

pub trait UIComponent {
    fn layout(&mut self, bounds: Rect, dpi: f32);
    fn render(&self, ctx: &ID2D1DeviceContext, theme: &Theme);
    fn handle_input(&mut self, event: &InputEvent) -> bool;
    fn hit_test(&self, point: Point) -> bool;
    fn set_visible(&mut self, visible: bool);
    fn bounds(&self) -> Rect;
}

pub struct LayoutEngine;

impl LayoutEngine {
    pub fn stack(
        bounds: Rect,
        direction: LayoutDirection,
        spacing: f32,
        sizes: &[f32],
    ) -> Vec<Rect> {
        let mut out = Vec::with_capacity(sizes.len());
        let mut cursor_x = bounds.x;
        let mut cursor_y = bounds.y;

        for size in sizes {
            let rect = match direction {
                LayoutDirection::Horizontal => Rect {
                    x: cursor_x,
                    y: bounds.y,
                    width: *size,
                    height: bounds.height,
                },
                LayoutDirection::Vertical => Rect {
                    x: bounds.x,
                    y: cursor_y,
                    width: bounds.width,
                    height: *size,
                },
            };

            out.push(rect);

            match direction {
                LayoutDirection::Horizontal => cursor_x += *size + spacing,
                LayoutDirection::Vertical => cursor_y += *size + spacing,
            }
        }

        out
    }

    pub fn absolute(origin: Rect, offset: Point, size: Size) -> Rect {
        Rect {
            x: origin.x + offset.x,
            y: origin.y + offset.y,
            width: size.width,
            height: size.height,
        }
    }

    pub fn flex(bounds: Rect, grow: &[f32], spacing: f32, direction: LayoutDirection) -> Vec<Rect> {
        let total_grow: f32 = grow.iter().copied().sum::<f32>().max(1.0);
        let axis_size = match direction {
            LayoutDirection::Horizontal => bounds.width,
            LayoutDirection::Vertical => bounds.height,
        };

        let free = (axis_size - spacing * (grow.len().saturating_sub(1) as f32)).max(0.0);
        let sizes: Vec<f32> = grow.iter().map(|g| free * (*g / total_grow)).collect();
        Self::stack(bounds, direction, spacing, &sizes)
    }
}

#[derive(Debug, Clone, Copy, Default)]
pub struct FocusRing {
    pub visible: bool,
    pub bounds: Rect,
    pub thickness: f32,
}

#[derive(Debug, Clone, Copy)]
pub struct AccessibilityPreferences {
    pub high_contrast: bool,
    pub reduce_motion: bool,
}

impl Default for AccessibilityPreferences {
    fn default() -> Self {
        Self {
            high_contrast: false,
            reduce_motion: false,
        }
    }
}

pub mod command_palette;
pub mod context_menu;
pub mod dialog;
pub mod sidebar;
pub mod statusbar;
pub mod tabs;
pub mod toast;
pub mod toolbar;
