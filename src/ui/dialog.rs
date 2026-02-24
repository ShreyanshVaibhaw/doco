use windows::Win32::Graphics::Direct2D::ID2D1DeviceContext;

use crate::{
    settings::{
        SettingSearchHit,
        SettingsStore,
        schema::{Settings, SettingsCategory},
        search_settings,
    },
    theme::Theme,
    ui::{InputEvent, Point, Rect, UIComponent},
};

const CATEGORY_PANE_WIDTH: f32 = 250.0;
const HEADER_HEIGHT: f32 = 56.0;
const SEARCH_HEIGHT: f32 = 36.0;
const ROW_HEIGHT: f32 = 30.0;

pub struct Dialog {
    bounds: Rect,
    visible: bool,
    search_query: String,
    selected_category: SettingsCategory,
    filtered_hits: Vec<SettingSearchHit>,
    scroll_y: f32,
    last_save_error: Option<String>,
    store: Option<SettingsStore>,
    fallback_settings: Settings,
}

impl Default for Dialog {
    fn default() -> Self {
        Self::new()
    }
}

impl Dialog {
    pub fn new() -> Self {
        match SettingsStore::load() {
            Ok(store) => Self {
                bounds: Rect::default(),
                visible: false,
                search_query: String::new(),
                selected_category: SettingsCategory::Appearance,
                filtered_hits: search_settings(""),
                scroll_y: 0.0,
                last_save_error: None,
                fallback_settings: Settings::default(),
                store: Some(store),
            },
            Err(err) => Self {
                bounds: Rect::default(),
                visible: false,
                search_query: String::new(),
                selected_category: SettingsCategory::Appearance,
                filtered_hits: search_settings(""),
                scroll_y: 0.0,
                last_save_error: Some(err.to_string()),
                fallback_settings: Settings::default(),
                store: None,
            },
        }
    }

    pub fn open(&mut self) {
        self.visible = true;
    }

    pub fn close(&mut self) {
        self.visible = false;
    }

    pub fn toggle(&mut self) {
        self.visible = !self.visible;
    }

    pub fn is_open(&self) -> bool {
        self.visible
    }

    pub fn handle_shortcut(&mut self, vk: u32, ctrl: bool, shift: bool) -> bool {
        if ctrl && !shift && vk == 0xBC {
            self.toggle();
            return true;
        }
        false
    }

    pub fn select_category(&mut self, category: SettingsCategory) {
        self.selected_category = category;
    }

    pub fn set_search_query(&mut self, query: impl Into<String>) {
        self.search_query = query.into();
        self.filtered_hits = search_settings(self.search_query.as_str());
    }

    pub fn search_query(&self) -> &str {
        self.search_query.as_str()
    }

    pub fn visible_categories(&self) -> Vec<SettingsCategory> {
        if self.search_query.trim().is_empty() {
            return SettingsCategory::all().to_vec();
        }

        let mut categories = SettingsCategory::all()
            .into_iter()
            .filter(|category| self.filtered_hits.iter().any(|hit| hit.category == *category))
            .collect::<Vec<_>>();
        categories.sort_unstable();
        categories
    }

    pub fn visible_setting_hits(&self) -> Vec<&SettingSearchHit> {
        self.filtered_hits
            .iter()
            .filter(|hit| hit.category == self.selected_category)
            .collect()
    }

    pub fn settings(&self) -> &Settings {
        if let Some(store) = &self.store {
            store.settings()
        } else {
            &self.fallback_settings
        }
    }

    pub fn apply_change<F>(&mut self, mutator: F)
    where
        F: FnOnce(&mut Settings),
    {
        if let Some(store) = &mut self.store {
            store.update(mutator);
        } else {
            mutator(&mut self.fallback_settings);
        }
    }

    pub fn tick(&mut self) {
        if let Some(store) = &mut self.store {
            if let Err(err) = store.flush_if_due() {
                self.last_save_error = Some(err.to_string());
            }
        }
    }

    pub fn force_flush(&mut self) {
        if let Some(store) = &mut self.store {
            if let Err(err) = store.force_flush() {
                self.last_save_error = Some(err.to_string());
            }
        }
    }

    pub fn has_conflicting_shortcuts(&self) -> bool {
        !self
            .settings()
            .keyboard_shortcuts
            .detect_conflicts()
            .is_empty()
    }

    pub fn set_shortcut(&mut self, command_id: &str, keys: impl Into<String>) {
        self.apply_change(|settings| settings.keyboard_shortcuts.set_binding(command_id, keys));
    }

    pub fn reset_shortcuts(&mut self) {
        self.apply_change(|settings| settings.keyboard_shortcuts.reset_to_defaults());
    }

    pub fn last_save_error(&self) -> Option<&str> {
        self.last_save_error.as_deref()
    }

    fn category_rect(&self, index: usize) -> Rect {
        Rect {
            x: self.bounds.x + 12.0,
            y: self.bounds.y + HEADER_HEIGHT + SEARCH_HEIGHT + 8.0 + (index as f32 * ROW_HEIGHT),
            width: CATEGORY_PANE_WIDTH - 24.0,
            height: ROW_HEIGHT,
        }
    }

    fn list_rect(&self) -> Rect {
        Rect {
            x: self.bounds.x + CATEGORY_PANE_WIDTH + 8.0,
            y: self.bounds.y + HEADER_HEIGHT + SEARCH_HEIGHT + 8.0,
            width: (self.bounds.width - CATEGORY_PANE_WIDTH - 20.0).max(0.0),
            height: (self.bounds.height - HEADER_HEIGHT - SEARCH_HEIGHT - 12.0).max(0.0),
        }
    }
}

impl UIComponent for Dialog {
    fn layout(&mut self, bounds: Rect, _dpi: f32) {
        self.bounds = bounds;
    }

    fn render(&self, _ctx: &ID2D1DeviceContext, _theme: &Theme) {
        // Settings dialog is composed in the shell renderer.
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
                    self.scroll_y = (self.scroll_y - 24.0).max(0.0);
                    true
                }
                0x28 => {
                    self.scroll_y += 24.0;
                    true
                }
                _ => false,
            },
            InputEvent::Char(ch) => {
                self.search_query.push(*ch);
                self.filtered_hits = search_settings(self.search_query.as_str());
                true
            }
            InputEvent::MouseWheel { delta, .. } => {
                self.scroll_y = (self.scroll_y - (delta * 14.0)).max(0.0);
                true
            }
            InputEvent::MouseDown(point) => {
                for (index, category) in self.visible_categories().into_iter().enumerate() {
                    if contains(self.category_rect(index), *point) {
                        self.select_category(category);
                        return true;
                    }
                }

                contains(self.list_rect(), *point)
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LoadingKind {
    Document,
    PdfPage,
    Image,
    LongOperation,
}

#[derive(Debug, Clone)]
pub struct LoadingState {
    pub kind: LoadingKind,
    pub progress: Option<f32>,
    pub message: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ErrorDialogKind {
    FileNotFound,
    CorruptedDocument,
    SaveFailed,
    OutOfMemory,
    PanicRecovery,
}

#[derive(Debug, Clone)]
pub struct ErrorDialogState {
    pub kind: ErrorDialogKind,
    pub title: String,
    pub body: String,
    pub retry_label: Option<String>,
}
