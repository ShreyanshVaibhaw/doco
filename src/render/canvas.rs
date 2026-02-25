use std::collections::HashMap;

use crate::{
    document::model::DocumentModel,
    render::animation::{Animation, Easing},
    ui::{Point, Rect, Size},
};

pub const PAGE_GAP: f32 = 24.0;
pub const ZOOM_MIN: f32 = 0.25;
pub const ZOOM_MAX: f32 = 5.0;
pub const ZOOM_DEFAULT: f32 = 1.0;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PageLayoutMode {
    SinglePage,
    #[default]
    Continuous,
    ReadMode,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ZoomPreset {
    FitWidth,
    FitPage,
    ActualSize,
}

#[derive(Debug, Clone, Copy, Default)]
pub struct ScrollState {
    pub x: f32,
    pub y: f32,
    pub velocity_x: f32,
    pub velocity_y: f32,
}

#[derive(Debug, Clone, Copy, Default)]
pub struct ScrollbarState {
    pub visible: bool,
    pub alpha: f32,
    pub thickness: f32,
    pub idle_seconds: f32,
}

#[derive(Debug, Clone, Copy, Default)]
pub struct CursorVisualState {
    pub blink_timer_s: f32,
    pub visible: bool,
}

#[derive(Debug, Clone, Copy, Default)]
pub struct CachedPage {
    pub page_index: usize,
    pub zoom_bucket: u16,
}

#[derive(Debug, Clone)]
pub struct CanvasState {
    pub viewport: Size,
    pub layout_mode: PageLayoutMode,
    pub scroll: ScrollState,
    pub zoom: f32,
    pub zoom_target: f32,
    pub zoom_anim: Option<Animation>,
    pub scroll_anim_x: Option<Animation>,
    pub scroll_anim_y: Option<Animation>,
    pub show_margin_guides: bool,
    pub reduce_motion: bool,
    pub scrollbar: ScrollbarState,
    pub cursor: CursorVisualState,
    pub page_cache: HashMap<usize, CachedPage>,
    pub dirty_rects: Vec<Rect>,
}

impl Default for CanvasState {
    fn default() -> Self {
        Self {
            viewport: Size {
                width: 800.0,
                height: 600.0,
            },
            layout_mode: PageLayoutMode::Continuous,
            scroll: ScrollState::default(),
            zoom: ZOOM_DEFAULT,
            zoom_target: ZOOM_DEFAULT,
            zoom_anim: None,
            scroll_anim_x: None,
            scroll_anim_y: None,
            show_margin_guides: false,
            reduce_motion: false,
            scrollbar: ScrollbarState {
                visible: false,
                alpha: 0.0,
                thickness: 6.0,
                idle_seconds: 0.0,
            },
            cursor: CursorVisualState {
                blink_timer_s: 0.0,
                visible: true,
            },
            page_cache: HashMap::new(),
            dirty_rects: Vec::new(),
        }
    }
}

impl CanvasState {
    pub fn set_viewport(&mut self, width: f32, height: f32) {
        self.viewport = Size {
            width: width.max(1.0),
            height: height.max(1.0),
        };
        self.mark_dirty_full();
    }

    pub fn set_layout_mode(&mut self, mode: PageLayoutMode) {
        self.layout_mode = mode;
        self.mark_dirty_full();
    }

    pub fn apply_zoom_preset(&mut self, preset: ZoomPreset, page_size: Size) {
        let target = match preset {
            ZoomPreset::FitWidth => (self.viewport.width / page_size.width).clamp(ZOOM_MIN, ZOOM_MAX),
            ZoomPreset::FitPage => {
                (self.viewport.height / page_size.height)
                    .min(self.viewport.width / page_size.width)
                    .clamp(ZOOM_MIN, ZOOM_MAX)
            }
            ZoomPreset::ActualSize => 1.0,
        };

        self.set_zoom(target, None);
    }

    pub fn set_zoom(&mut self, target_zoom: f32, cursor_pos: Option<Point>) {
        let clamped = target_zoom.clamp(ZOOM_MIN, ZOOM_MAX);

        if let Some(cursor) = cursor_pos {
            let rel_x = (cursor.x + self.scroll.x) / self.zoom.max(0.001);
            let rel_y = (cursor.y + self.scroll.y) / self.zoom.max(0.001);
            self.scroll.x = rel_x * clamped - cursor.x;
            self.scroll.y = rel_y * clamped - cursor.y;
        }

        self.zoom_target = clamped;
        if self.reduce_motion {
            self.zoom = self.zoom_target;
            self.zoom_anim = None;
        } else {
            self.zoom_anim = Some(Animation::new(
                self.zoom,
                self.zoom_target,
                0.15,
                Easing::EaseOutCubic,
            ));
        }
        self.page_cache.clear();
        self.mark_dirty_full();
    }

    pub fn handle_mouse_wheel(&mut self, delta: f32, ctrl_down: bool, cursor: Point) {
        if ctrl_down {
            let step = if delta > 0.0 { 0.1 } else { -0.1 };
            self.set_zoom(self.zoom * (1.0 + step), Some(cursor));
        } else {
            let impulse = -delta * 3.0;
            self.scroll.velocity_y += impulse;
            if self.reduce_motion {
                self.scroll.y += impulse * 0.35;
                self.scroll_anim_y = None;
            } else {
                self.scroll_anim_y = Some(Animation::new(
                    self.scroll.y,
                    self.scroll.y + impulse * 0.35,
                    0.20,
                    Easing::Spring,
                ));
            }
            self.scrollbar.visible = true;
            self.scrollbar.alpha = self.scrollbar.alpha.max(0.65);
            self.scrollbar.idle_seconds = 0.0;
            self.mark_dirty_full();
        }
    }

    pub fn handle_horizontal_wheel(&mut self, delta: f32) {
        let impulse = -delta * 3.0;
        self.scroll.velocity_x += impulse;
        if self.reduce_motion {
            self.scroll.x += impulse * 0.35;
            self.scroll_anim_x = None;
        } else {
            self.scroll_anim_x = Some(Animation::new(
                self.scroll.x,
                self.scroll.x + impulse * 0.35,
                0.20,
                Easing::Spring,
            ));
        }
        self.scrollbar.visible = true;
        self.scrollbar.alpha = self.scrollbar.alpha.max(0.65);
        self.scrollbar.idle_seconds = 0.0;
        self.mark_dirty_full();
    }

    pub fn content_size(&self, document: &DocumentModel) -> Size {
        let (page_width, page_height) = page_dimensions_points(document);
        let scaled_w = page_width * self.zoom;
        let scaled_h = page_height * self.zoom;

        match self.layout_mode {
            PageLayoutMode::ReadMode => Size {
                width: (self.viewport.width * 0.88).max(540.0),
                height: (document.content.len() as f32 * 26.0 * self.zoom).max(self.viewport.height),
            },
            PageLayoutMode::SinglePage => Size {
                width: scaled_w,
                height: scaled_h,
            },
            PageLayoutMode::Continuous => {
                let page_count = document.pages.len().max(1);
                Size {
                    width: scaled_w,
                    height: (page_count as f32 * scaled_h)
                        + ((page_count.saturating_sub(1)) as f32 * PAGE_GAP),
                }
            }
        }
    }

    pub fn clamp_scroll(&mut self, document: &DocumentModel) {
        let content = self.content_size(document);
        let max_x = (content.width - self.viewport.width).max(0.0);
        let max_y = (content.height - self.viewport.height).max(0.0);
        self.scroll.x = self.scroll.x.clamp(0.0, max_x);
        self.scroll.y = self.scroll.y.clamp(0.0, max_y);
    }

    pub fn update(&mut self, dt_s: f32) -> bool {
        let mut animating = false;

        if let Some(anim) = &mut self.zoom_anim {
            if anim.update_respecting_motion_pref(dt_s, self.reduce_motion) {
                self.zoom = anim.current_value;
                animating = true;
                self.mark_dirty_full();
            } else {
                self.zoom = anim.end_value;
                self.zoom_anim = None;
                self.mark_dirty_full();
            }
        }

        if let Some(anim) = &mut self.scroll_anim_x {
            if anim.update_respecting_motion_pref(dt_s, self.reduce_motion) {
                self.scroll.x = anim.current_value;
                animating = true;
                self.scrollbar.idle_seconds = 0.0;
                self.mark_dirty_full();
            } else {
                self.scroll.x = anim.end_value;
                self.scroll_anim_x = None;
                self.mark_dirty_full();
            }
        }

        if let Some(anim) = &mut self.scroll_anim_y {
            if anim.update_respecting_motion_pref(dt_s, self.reduce_motion) {
                self.scroll.y = anim.current_value;
                animating = true;
                self.scrollbar.idle_seconds = 0.0;
                self.mark_dirty_full();
            } else {
                self.scroll.y = anim.end_value;
                self.scroll_anim_y = None;
                self.mark_dirty_full();
            }
        }

        self.cursor.blink_timer_s += dt_s;
        if self.cursor.blink_timer_s >= 0.53 {
            self.cursor.blink_timer_s = 0.0;
            self.cursor.visible = !self.cursor.visible;
            animating = true;
            self.mark_dirty_full();
        }

        if self.scrollbar.visible {
            if self.scroll_anim_x.is_some() || self.scroll_anim_y.is_some() {
                self.scrollbar.idle_seconds = 0.0;
                self.scrollbar.alpha = (self.scrollbar.alpha + dt_s * 4.0).clamp(0.0, 1.0);
            } else {
                self.scrollbar.idle_seconds += dt_s;
                if self.scrollbar.idle_seconds >= 1.5 {
                    self.scrollbar.alpha = (self.scrollbar.alpha - dt_s * 2.0).clamp(0.0, 1.0);
                    if self.scrollbar.alpha <= 0.01 {
                        self.scrollbar.visible = false;
                        self.scrollbar.idle_seconds = 0.0;
                    }
                } else {
                    self.scrollbar.alpha = (self.scrollbar.alpha + dt_s * 4.0).clamp(0.0, 1.0);
                }
            }
            self.mark_dirty_full();
        }

        animating
    }

    pub fn set_reduce_motion(&mut self, reduce_motion: bool) {
        self.reduce_motion = reduce_motion;
        if reduce_motion {
            self.zoom_anim = None;
            self.scroll_anim_x = None;
            self.scroll_anim_y = None;
            self.zoom = self.zoom_target;
        }
    }

    pub fn page_rects(&self, document: &DocumentModel) -> Vec<Rect> {
        let (page_width, page_height) = page_dimensions_points(document);
        let scaled_w = page_width * self.zoom;
        let scaled_h = page_height * self.zoom;

        match self.layout_mode {
            PageLayoutMode::ReadMode => vec![Rect {
                x: -self.scroll.x,
                y: -self.scroll.y,
                width: (self.viewport.width * 0.88).max(540.0),
                height: (document.content.len() as f32 * 26.0 * self.zoom).max(self.viewport.height),
            }],
            PageLayoutMode::SinglePage => vec![Rect {
                x: ((self.viewport.width - scaled_w) * 0.5).max(0.0) - self.scroll.x,
                y: ((self.viewport.height - scaled_h) * 0.5).max(0.0) - self.scroll.y,
                width: scaled_w,
                height: scaled_h,
            }],
            PageLayoutMode::Continuous => {
                let page_count = document.pages.len().max(1);
                let left = ((self.viewport.width - scaled_w) * 0.5).max(0.0) - self.scroll.x;
                (0..page_count)
                    .map(|i| Rect {
                        x: left,
                        y: (i as f32) * (scaled_h + PAGE_GAP) - self.scroll.y,
                        width: scaled_w,
                        height: scaled_h,
                    })
                    .collect()
            }
        }
    }

    pub fn visible_page_indices(&self, document: &DocumentModel) -> Vec<usize> {
        self.page_rects(document)
            .into_iter()
            .enumerate()
            .filter_map(|(index, rect)| {
                let visible = rect.y < self.viewport.height && (rect.y + rect.height) > 0.0;
                if visible { Some(index) } else { None }
            })
            .collect()
    }

    pub fn cull_and_cache_visible_pages(&mut self, document: &DocumentModel) -> Vec<usize> {
        let visible = self.visible_page_indices(document);
        let bucket = (self.zoom * 100.0) as u16;

        self.page_cache.retain(|index, _| visible.contains(index));
        for index in &visible {
            self.page_cache.insert(
                *index,
                CachedPage {
                    page_index: *index,
                    zoom_bucket: bucket,
                },
            );
        }

        visible
    }

    pub fn mark_dirty_rect(&mut self, rect: Rect) {
        self.dirty_rects.push(rect);
    }

    pub fn mark_dirty_full(&mut self) {
        self.mark_dirty_rect(Rect {
            x: 0.0,
            y: 0.0,
            width: self.viewport.width,
            height: self.viewport.height,
        });
    }

    pub fn take_dirty_rects(&mut self) -> Vec<Rect> {
        std::mem::take(&mut self.dirty_rects)
    }
}

fn page_dimensions_points(document: &DocumentModel) -> (f32, f32) {
    use crate::document::model::PageSize;

    match document.metadata.page_size {
        PageSize::Letter => (612.0, 792.0),
        PageSize::A4 => (595.0, 842.0),
        PageSize::Legal => (612.0, 1008.0),
        PageSize::Custom {
            width_points,
            height_points,
        } => (width_points, height_points),
    }
}

#[cfg(test)]
mod tests {
    use super::{CanvasState, Point};

    #[test]
    fn scrollbar_waits_before_fading_out() {
        let mut canvas = CanvasState::default();
        canvas.handle_mouse_wheel(1.0, false, Point { x: 40.0, y: 40.0 });
        canvas.scroll_anim_y = None;
        canvas.scrollbar.alpha = 1.0;
        let _ = canvas.update(1.0);
        assert!(canvas.scrollbar.visible);
        assert!(canvas.scrollbar.alpha > 0.9);
        let _ = canvas.update(0.7);
        assert!(canvas.scrollbar.alpha < 1.0);
    }

    #[test]
    fn reduce_motion_disables_zoom_animation() {
        let mut canvas = CanvasState::default();
        canvas.set_reduce_motion(true);
        canvas.set_zoom(1.5, None);
        assert!(canvas.zoom_anim.is_none());
        assert_eq!(canvas.zoom, 1.5);
    }
}
