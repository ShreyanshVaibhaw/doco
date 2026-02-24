use std::time::{Duration, Instant};

use windows::Win32::Graphics::Direct2D::ID2D1DeviceContext;

use crate::{
    render::animation::{Animation, Easing},
    theme::Theme,
    ui::{InputEvent, Point, Rect, UIComponent},
};

const TOAST_WIDTH: f32 = 320.0;
const TOAST_HEIGHT: f32 = 64.0;
const TOAST_GAP: f32 = 10.0;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ToastLevel {
    Info,
    Success,
    Warning,
    Error,
}

#[derive(Debug, Clone)]
pub struct ToastEntry {
    pub id: u64,
    pub level: ToastLevel,
    pub title: String,
    pub body: String,
    pub created_at: Instant,
    pub ttl: Duration,
    pub opacity: f32,
    pub slide: f32,
    fade_anim: Option<Animation>,
    slide_anim: Option<Animation>,
}

#[derive(Debug, Default)]
pub struct Toast {
    bounds: Rect,
    visible: bool,
    next_id: u64,
    pub reduce_motion: bool,
    pub entries: Vec<ToastEntry>,
}

impl Toast {
    pub fn push(&mut self, level: ToastLevel, title: impl Into<String>, body: impl Into<String>) -> u64 {
        let id = self.next_id.max(1);
        self.next_id = id + 1;

        let entry = ToastEntry {
            id,
            level,
            title: title.into(),
            body: body.into(),
            created_at: Instant::now(),
            ttl: Duration::from_secs(4),
            opacity: 0.0,
            slide: 26.0,
            fade_anim: Some(Animation::new(0.0, 1.0, 0.16, Easing::EaseOutCubic)),
            slide_anim: Some(Animation::new(26.0, 0.0, 0.20, Easing::Spring)),
        };

        self.entries.insert(0, entry);
        self.entries.truncate(4);
        id
    }

    pub fn push_export_complete(&mut self, path: &str) {
        self.push(
            ToastLevel::Success,
            "Export completed",
            format!("Saved to {}", path),
        );
    }

    pub fn push_recovery_saved(&mut self, path: &str) {
        self.push(
            ToastLevel::Info,
            "Recovery saved",
            format!("Backup written to {}", path),
        );
    }

    pub fn dismiss(&mut self, id: u64) -> bool {
        let before = self.entries.len();
        self.entries.retain(|entry| entry.id != id);
        self.entries.len() != before
    }

    pub fn tick(&mut self, dt_s: f32) {
        for entry in &mut self.entries {
            if !self.reduce_motion {
                if let Some(anim) = &mut entry.fade_anim {
                    if anim.update(dt_s) {
                        entry.opacity = anim.current_value.clamp(0.0, 1.0);
                    } else {
                        entry.opacity = anim.end_value.clamp(0.0, 1.0);
                        entry.fade_anim = None;
                    }
                }

                if let Some(anim) = &mut entry.slide_anim {
                    if anim.update(dt_s) {
                        entry.slide = anim.current_value;
                    } else {
                        entry.slide = anim.end_value;
                        entry.slide_anim = None;
                    }
                }
            } else {
                entry.opacity = 1.0;
                entry.slide = 0.0;
                entry.fade_anim = None;
                entry.slide_anim = None;
            }
        }

        self.entries.retain(|entry| entry.created_at.elapsed() < entry.ttl);
    }

    pub fn entry_rect(&self, index: usize) -> Rect {
        Rect {
            x: self.bounds.x + self.bounds.width - TOAST_WIDTH - 12.0,
            y: self.bounds.y + self.bounds.height
                - TOAST_HEIGHT
                - 16.0
                - (index as f32 * (TOAST_HEIGHT + TOAST_GAP)),
            width: TOAST_WIDTH,
            height: TOAST_HEIGHT,
        }
    }
}

impl UIComponent for Toast {
    fn layout(&mut self, bounds: Rect, _dpi: f32) {
        self.bounds = bounds;
    }

    fn render(&self, _ctx: &ID2D1DeviceContext, _theme: &Theme) {
        // Drawn by host renderer.
    }

    fn handle_input(&mut self, event: &InputEvent) -> bool {
        if !self.visible {
            return false;
        }

        match event {
            InputEvent::MouseDown(point) => {
                for (idx, entry) in self.entries.iter().enumerate() {
                    if contains(self.entry_rect(idx), *point) {
                        return self.dismiss(entry.id);
                    }
                }
                false
            }
            _ => false,
        }
    }

    fn hit_test(&self, point: Point) -> bool {
        self.visible
            && self
                .entries
                .iter()
                .enumerate()
                .any(|(idx, _)| contains(self.entry_rect(idx), point))
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
