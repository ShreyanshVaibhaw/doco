use windows::Win32::Graphics::Direct2D::ID2D1DeviceContext;

use crate::{
    settings::{
        SettingSearchHit,
        SettingsStore,
        schema::{
            AutoSaveInterval, CursorStyle, DefaultMargins, DefaultOpenFolder, DefaultPageSize,
            DefaultViewMode, PatternQuality, Settings, SettingsCategory, ShowWhitespaceMode,
            SidebarDefaultPanel, ThemePreference, UiScale, WordWrapMode,
        },
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
    selected_hit: usize,
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
                selected_hit: 0,
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
                selected_hit: 0,
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
        self.selected_hit = 0;
        self.scroll_y = 0.0;
    }

    pub fn set_search_query(&mut self, query: impl Into<String>) {
        self.search_query = query.into();
        self.filtered_hits = search_settings(self.search_query.as_str());
        self.selected_hit = 0;
        self.scroll_y = 0.0;
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

    pub fn selected_category(&self) -> SettingsCategory {
        self.selected_category
    }

    pub fn selected_setting_row(&self) -> usize {
        let total = self.visible_setting_hits().len();
        if total == 0 {
            0
        } else {
            self.selected_hit.min(total - 1)
        }
    }

    pub fn setting_rows(&self) -> Vec<String> {
        let settings = self.settings();
        self.visible_setting_hits()
            .into_iter()
            .map(|hit| {
                format!(
                    "{}: {}",
                    hit.title,
                    setting_value_preview(settings, hit.setting_key)
                )
            })
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

    fn current_hits(&self) -> Vec<SettingSearchHit> {
        self.visible_setting_hits()
            .into_iter()
            .cloned()
            .collect::<Vec<_>>()
    }

    fn clamp_selected_hit(&mut self) {
        let total = self.visible_setting_hits().len();
        if total == 0 {
            self.selected_hit = 0;
        } else {
            self.selected_hit = self.selected_hit.min(total - 1);
        }
    }

    fn advance_category(&mut self, delta: isize) {
        let categories = self.visible_categories();
        if categories.is_empty() {
            return;
        }

        let current = categories
            .iter()
            .position(|c| *c == self.selected_category)
            .unwrap_or(0) as isize;
        let next = (current + delta).rem_euclid(categories.len() as isize) as usize;
        self.select_category(categories[next]);
    }

    fn activate_selected_hit(&mut self) -> bool {
        let hits = self.current_hits();
        if hits.is_empty() {
            return false;
        }
        let index = self.selected_hit.min(hits.len() - 1);
        self.selected_hit = index;
        self.apply_setting_cycle(hits[index].setting_key)
    }

    fn row_index_from_point(&self, point: Point) -> Option<usize> {
        let list = self.list_rect();
        if !contains(list, point) {
            return None;
        }

        let offset_y = (point.y - list.y + self.scroll_y).max(0.0);
        let row = (offset_y / ROW_HEIGHT).floor() as usize;
        Some(row)
    }

    fn apply_setting_cycle(&mut self, key: &str) -> bool {
        let before = serde_json::to_string(self.settings()).ok();
        self.apply_change(|settings| match key {
            "appearance.theme" => {
                settings.appearance.theme = match &settings.appearance.theme {
                    ThemePreference::SystemAuto => ThemePreference::Named("Light".to_string()),
                    ThemePreference::Named(name) if name.eq_ignore_ascii_case("light") => {
                        ThemePreference::Named("Dark".to_string())
                    }
                    _ => ThemePreference::SystemAuto,
                };
            }
            "appearance.canvas_background" => {
                let next = match settings.appearance.canvas_background.preset_id.as_str() {
                    "paper" => "clean",
                    "clean" => "blueprint",
                    "blueprint" => "cozy",
                    "cozy" => "midnight",
                    _ => "paper",
                };
                settings.appearance.canvas_background.preset_id = next.to_string();
            }
            "appearance.ui_font" => {
                settings.appearance.ui_font = if settings.appearance.ui_font == "Segoe UI Variable" {
                    "Segoe UI".to_string()
                } else {
                    "Segoe UI Variable".to_string()
                };
            }
            "appearance.ui_scale" => {
                settings.appearance.ui_scale = match settings.appearance.ui_scale {
                    UiScale::Percent100 => UiScale::Percent125,
                    UiScale::Percent125 => UiScale::Percent150,
                    UiScale::Percent150 => UiScale::Percent175,
                    UiScale::Percent175 => UiScale::Percent200,
                    UiScale::Percent200 => UiScale::Percent100,
                };
            }
            "appearance.show_toolbar" => {
                settings.appearance.show_toolbar = !settings.appearance.show_toolbar;
            }
            "appearance.show_sidebar" => {
                settings.appearance.show_sidebar = !settings.appearance.show_sidebar;
            }
            "appearance.show_status_bar" => {
                settings.appearance.show_status_bar = !settings.appearance.show_status_bar;
            }
            "appearance.show_tab_bar" => {
                settings.appearance.show_tab_bar = !settings.appearance.show_tab_bar;
            }
            "appearance.sidebar_default_panel" => {
                settings.appearance.sidebar_default_panel =
                    match settings.appearance.sidebar_default_panel {
                        SidebarDefaultPanel::Files => SidebarDefaultPanel::Outline,
                        SidebarDefaultPanel::Outline => SidebarDefaultPanel::Bookmarks,
                        SidebarDefaultPanel::Bookmarks => SidebarDefaultPanel::Files,
                    };
            }
            "editor.default_font_family" => {
                settings.editor.default_font_family = match settings.editor.default_font_family.as_str()
                {
                    "Segoe UI" => "Calibri".to_string(),
                    "Calibri" => "Consolas".to_string(),
                    _ => "Segoe UI".to_string(),
                };
            }
            "editor.default_font_size_pt" => {
                settings.editor.default_font_size_pt = match settings.editor.default_font_size_pt {
                    12 => 14,
                    14 => 16,
                    16 => 10,
                    _ => 12,
                };
            }
            "editor.tab_size" => {
                settings.editor.tab_size = match settings.editor.tab_size {
                    2 => 4,
                    4 => 8,
                    _ => 2,
                };
            }
            "editor.insert_spaces_instead_of_tabs" => {
                settings.editor.insert_spaces_instead_of_tabs =
                    !settings.editor.insert_spaces_instead_of_tabs;
            }
            "editor.word_wrap" => {
                settings.editor.word_wrap = match settings.editor.word_wrap {
                    WordWrapMode::On => WordWrapMode::Off,
                    WordWrapMode::Off => WordWrapMode::AtColumn(80),
                    WordWrapMode::AtColumn(_) => WordWrapMode::On,
                };
            }
            "editor.show_line_numbers" => {
                settings.editor.show_line_numbers = !settings.editor.show_line_numbers;
            }
            "editor.cursor_style" => {
                settings.editor.cursor_style = match settings.editor.cursor_style {
                    CursorStyle::Line => CursorStyle::Block,
                    CursorStyle::Block => CursorStyle::Underline,
                    CursorStyle::Underline => CursorStyle::Line,
                };
            }
            "editor.cursor_blink" => {
                settings.editor.cursor_blink = !settings.editor.cursor_blink;
            }
            "editor.auto_indent" => {
                settings.editor.auto_indent = !settings.editor.auto_indent;
            }
            "editor.auto_close_brackets" => {
                settings.editor.auto_close_brackets = !settings.editor.auto_close_brackets;
            }
            "editor.show_whitespace" => {
                settings.editor.show_whitespace = match settings.editor.show_whitespace {
                    ShowWhitespaceMode::Off => ShowWhitespaceMode::Selection,
                    ShowWhitespaceMode::Selection => ShowWhitespaceMode::All,
                    ShowWhitespaceMode::All => ShowWhitespaceMode::Off,
                };
            }
            "document.default_page_size" => {
                settings.document.default_page_size = match settings.document.default_page_size {
                    DefaultPageSize::Letter => DefaultPageSize::A4,
                    DefaultPageSize::A4 => DefaultPageSize::Legal,
                    DefaultPageSize::Legal => DefaultPageSize::Letter,
                };
            }
            "document.default_margins" => {
                settings.document.default_margins = match settings.document.default_margins {
                    DefaultMargins::Normal => DefaultMargins::Narrow,
                    DefaultMargins::Narrow => DefaultMargins::Wide,
                    DefaultMargins::Wide => DefaultMargins::Normal,
                };
            }
            "document.default_line_spacing" => {
                settings.document.default_line_spacing = match settings.document.default_line_spacing {
                    v if (v - 1.15).abs() < f32::EPSILON => 1.5,
                    v if (v - 1.5).abs() < f32::EPSILON => 2.0,
                    v if (v - 2.0).abs() < f32::EPSILON => 1.0,
                    _ => 1.15,
                };
            }
            "document.default_view_mode" => {
                settings.document.default_view_mode = match settings.document.default_view_mode {
                    DefaultViewMode::Page => DefaultViewMode::Continuous,
                    DefaultViewMode::Continuous => DefaultViewMode::Read,
                    DefaultViewMode::Read => DefaultViewMode::Page,
                };
            }
            "document.default_zoom_percent" => {
                settings.document.default_zoom_percent = match settings.document.default_zoom_percent {
                    100 => 125,
                    125 => 150,
                    150 => 175,
                    175 => 200,
                    _ => 100,
                };
            }
            "document.spelling_check" => {
                settings.document.spelling_check = !settings.document.spelling_check;
            }
            "files.auto_save_interval" => {
                settings.files.auto_save_interval = match settings.files.auto_save_interval {
                    AutoSaveInterval::Seconds(60) => AutoSaveInterval::Seconds(30),
                    AutoSaveInterval::Seconds(30) => AutoSaveInterval::Seconds(120),
                    AutoSaveInterval::Seconds(120) => AutoSaveInterval::Off,
                    AutoSaveInterval::Off => AutoSaveInterval::Seconds(60),
                    _ => AutoSaveInterval::Seconds(60),
                };
            }
            "files.create_backup_before_save" => {
                settings.files.create_backup_before_save = !settings.files.create_backup_before_save;
            }
            "files.default_save_format" => {
                settings.files.default_save_format = match settings.files.default_save_format.as_str() {
                    ".docx" => ".txt".to_string(),
                    ".txt" => ".md".to_string(),
                    ".md" => ".pdf".to_string(),
                    _ => ".docx".to_string(),
                };
            }
            "files.recent_files_count" => {
                settings.files.recent_files_count = match settings.files.recent_files_count {
                    20 => 30,
                    30 => 10,
                    _ => 20,
                };
            }
            "files.default_open_folder" => {
                settings.files.default_open_folder = match &settings.files.default_open_folder {
                    DefaultOpenFolder::LastUsed => DefaultOpenFolder::Documents,
                    DefaultOpenFolder::Documents => {
                        DefaultOpenFolder::SpecificPath("%USERPROFILE%\\Documents".to_string())
                    }
                    DefaultOpenFolder::SpecificPath(_) => DefaultOpenFolder::LastUsed,
                };
            }
            "keyboard_shortcuts.bindings" => {
                let next = settings
                    .keyboard_shortcuts
                    .bindings
                    .get("view.settings")
                    .map(|binding| binding.keys.as_str())
                    .map(|keys| {
                        if keys.eq_ignore_ascii_case("Ctrl+,") {
                            "Ctrl+Alt+S"
                        } else {
                            "Ctrl+,"
                        }
                    })
                    .unwrap_or("Ctrl+,");
                settings.keyboard_shortcuts.set_binding("view.settings", next);
            }
            "keyboard_shortcuts.reset_defaults" => {
                settings.keyboard_shortcuts.reset_to_defaults();
            }
            "performance.hardware_acceleration" => {
                settings.performance.hardware_acceleration =
                    !settings.performance.hardware_acceleration;
            }
            "performance.max_undo_history" => {
                settings.performance.max_undo_history = match settings.performance.max_undo_history {
                    1000 => 2000,
                    2000 => 500,
                    _ => 1000,
                };
            }
            "performance.background_pattern_quality" => {
                settings.performance.background_pattern_quality =
                    match settings.performance.background_pattern_quality {
                        PatternQuality::High => PatternQuality::Low,
                        PatternQuality::Low => PatternQuality::High,
                    };
            }
            "performance.animated_backgrounds" => {
                settings.performance.animated_backgrounds =
                    !settings.performance.animated_backgrounds;
            }
            "performance.max_image_cache_mb" => {
                settings.performance.max_image_cache_mb = match settings.performance.max_image_cache_mb {
                    200 => 400,
                    400 => 100,
                    _ => 200,
                };
            }
            "about.check_updates_on_startup" => {
                settings.about.check_updates_on_startup = !settings.about.check_updates_on_startup;
            }
            "about.check_for_updates" => {
                settings.about.last_update_check_utc = Some(chrono::Utc::now());
            }
            _ => {}
        });
        let after = serde_json::to_string(self.settings()).ok();
        before != after
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
                0x08 => {
                    remove_last_char(&mut self.search_query);
                    self.filtered_hits = search_settings(self.search_query.as_str());
                    self.selected_hit = 0;
                    self.scroll_y = 0.0;
                    true
                }
                0x26 => {
                    if self.selected_hit > 0 {
                        self.selected_hit -= 1;
                    }
                    self.scroll_y = (self.scroll_y - ROW_HEIGHT).max(0.0);
                    true
                }
                0x28 => {
                    let total = self.visible_setting_hits().len();
                    if total > 0 {
                        self.selected_hit = (self.selected_hit + 1).min(total - 1);
                    }
                    self.scroll_y += ROW_HEIGHT;
                    true
                }
                0x09 => {
                    self.advance_category(1);
                    true
                }
                0x0D => self.activate_selected_hit(),
                _ => false,
            },
            InputEvent::Char(ch) => {
                if ch.is_control() {
                    false
                } else {
                    self.search_query.push(*ch);
                    self.filtered_hits = search_settings(self.search_query.as_str());
                    self.selected_hit = 0;
                    self.scroll_y = 0.0;
                    true
                }
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
                if let Some(row) = self.row_index_from_point(*point) {
                    let hits = self.current_hits();
                    if row < hits.len() {
                        self.selected_hit = row;
                        self.clamp_selected_hit();
                        return self.activate_selected_hit();
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

fn setting_value_preview(settings: &Settings, key: &str) -> String {
    match key {
        "appearance.theme" => match &settings.appearance.theme {
            ThemePreference::SystemAuto => "System (auto)".to_string(),
            ThemePreference::Named(name) => name.clone(),
        },
        "appearance.canvas_background" => settings.appearance.canvas_background.preset_id.clone(),
        "appearance.ui_font" => settings.appearance.ui_font.clone(),
        "appearance.ui_scale" => format!("{:.0}%", settings.appearance.ui_scale.as_factor() * 100.0),
        "appearance.show_toolbar" => bool_text(settings.appearance.show_toolbar),
        "appearance.show_sidebar" => bool_text(settings.appearance.show_sidebar),
        "appearance.show_status_bar" => bool_text(settings.appearance.show_status_bar),
        "appearance.show_tab_bar" => bool_text(settings.appearance.show_tab_bar),
        "appearance.sidebar_default_panel" => match settings.appearance.sidebar_default_panel {
            SidebarDefaultPanel::Files => "Files".to_string(),
            SidebarDefaultPanel::Outline => "Outline".to_string(),
            SidebarDefaultPanel::Bookmarks => "Bookmarks".to_string(),
        },
        "editor.default_font_family" => settings.editor.default_font_family.clone(),
        "editor.default_font_size_pt" => format!("{} pt", settings.editor.default_font_size_pt),
        "editor.tab_size" => settings.editor.tab_size.to_string(),
        "editor.insert_spaces_instead_of_tabs" => {
            bool_text(settings.editor.insert_spaces_instead_of_tabs)
        }
        "editor.word_wrap" => match settings.editor.word_wrap {
            WordWrapMode::On => "On".to_string(),
            WordWrapMode::Off => "Off".to_string(),
            WordWrapMode::AtColumn(col) => format!("At column {}", col),
        },
        "editor.show_line_numbers" => bool_text(settings.editor.show_line_numbers),
        "editor.cursor_style" => match settings.editor.cursor_style {
            CursorStyle::Line => "Line".to_string(),
            CursorStyle::Block => "Block".to_string(),
            CursorStyle::Underline => "Underline".to_string(),
        },
        "editor.cursor_blink" => bool_text(settings.editor.cursor_blink),
        "editor.auto_indent" => bool_text(settings.editor.auto_indent),
        "editor.auto_close_brackets" => bool_text(settings.editor.auto_close_brackets),
        "editor.show_whitespace" => match settings.editor.show_whitespace {
            ShowWhitespaceMode::Off => "Off".to_string(),
            ShowWhitespaceMode::Selection => "Selection".to_string(),
            ShowWhitespaceMode::All => "All".to_string(),
        },
        "document.default_page_size" => match settings.document.default_page_size {
            DefaultPageSize::Letter => "Letter".to_string(),
            DefaultPageSize::A4 => "A4".to_string(),
            DefaultPageSize::Legal => "Legal".to_string(),
        },
        "document.default_margins" => match settings.document.default_margins {
            DefaultMargins::Normal => "Normal (1\")".to_string(),
            DefaultMargins::Narrow => "Narrow (0.5\")".to_string(),
            DefaultMargins::Wide => "Wide (1.25\")".to_string(),
        },
        "document.default_line_spacing" => format!("{:.2}", settings.document.default_line_spacing),
        "document.default_view_mode" => match settings.document.default_view_mode {
            DefaultViewMode::Page => "Page".to_string(),
            DefaultViewMode::Continuous => "Continuous".to_string(),
            DefaultViewMode::Read => "Read".to_string(),
        },
        "document.default_zoom_percent" => format!("{}%", settings.document.default_zoom_percent),
        "document.spelling_check" => bool_text(settings.document.spelling_check),
        "files.auto_save_interval" => match settings.files.auto_save_interval {
            AutoSaveInterval::Off => "Off".to_string(),
            AutoSaveInterval::Seconds(seconds) => format!("{}s", seconds),
        },
        "files.create_backup_before_save" => bool_text(settings.files.create_backup_before_save),
        "files.default_save_format" => settings.files.default_save_format.clone(),
        "files.recent_files_count" => settings.files.recent_files_count.to_string(),
        "files.default_open_folder" => match &settings.files.default_open_folder {
            DefaultOpenFolder::LastUsed => "Last used".to_string(),
            DefaultOpenFolder::Documents => "Documents".to_string(),
            DefaultOpenFolder::SpecificPath(path) => path.clone(),
        },
        "keyboard_shortcuts.bindings" => format!("{} bindings", settings.keyboard_shortcuts.bindings.len()),
        "keyboard_shortcuts.reset_defaults" => "Reset all to defaults".to_string(),
        "performance.hardware_acceleration" => bool_text(settings.performance.hardware_acceleration),
        "performance.max_undo_history" => settings.performance.max_undo_history.to_string(),
        "performance.background_pattern_quality" => match settings.performance.background_pattern_quality {
            PatternQuality::High => "High".to_string(),
            PatternQuality::Low => "Low".to_string(),
        },
        "performance.animated_backgrounds" => bool_text(settings.performance.animated_backgrounds),
        "performance.max_image_cache_mb" => format!("{} MB", settings.performance.max_image_cache_mb),
        "about.version" => settings.about.version.clone(),
        "about.check_updates_on_startup" => bool_text(settings.about.check_updates_on_startup),
        "about.licenses_url" => settings.about.licenses_url.clone(),
        "about.system_info_snapshot" => {
            if settings.about.system_info_snapshot.trim().is_empty() {
                "Not captured".to_string()
            } else {
                settings.about.system_info_snapshot.clone()
            }
        }
        "about.check_for_updates" => settings
            .about
            .last_update_check_utc
            .map(|v| format!("Last checked: {}", v.to_rfc3339()))
            .unwrap_or_else(|| "Never checked".to_string()),
        _ => "Tap to edit".to_string(),
    }
}

fn bool_text(value: bool) -> String {
    if value {
        "On".to_string()
    } else {
        "Off".to_string()
    }
}

fn remove_last_char(text: &mut String) {
    let _ = text.pop();
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

impl LoadingState {
    pub fn for_kind(kind: LoadingKind, progress: Option<f32>) -> Self {
        let message = match kind {
            LoadingKind::Document => "Loading document...".to_string(),
            LoadingKind::PdfPage => "Rendering PDF page...".to_string(),
            LoadingKind::Image => "Loading image...".to_string(),
            LoadingKind::LongOperation => "Working...".to_string(),
        };
        Self {
            kind,
            progress: progress.map(|v| v.clamp(0.0, 1.0)),
            message,
        }
    }
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

impl ErrorDialogState {
    pub fn from_kind(kind: ErrorDialogKind, detail: impl Into<String>) -> Self {
        let detail = detail.into();
        match kind {
            ErrorDialogKind::FileNotFound => Self {
                kind,
                title: "File not found".to_string(),
                body: format!("{detail}\nUse Browse to locate the file."),
                retry_label: Some("Browse".to_string()),
            },
            ErrorDialogKind::CorruptedDocument => Self {
                kind,
                title: "Document looks corrupted".to_string(),
                body: format!("{detail}\nDoco will show recoverable content only."),
                retry_label: None,
            },
            ErrorDialogKind::SaveFailed => Self {
                kind,
                title: "Save failed".to_string(),
                body: format!("{detail}\nTry saving to a different location."),
                retry_label: Some("Retry Save".to_string()),
            },
            ErrorDialogKind::OutOfMemory => Self {
                kind,
                title: "Out of memory".to_string(),
                body: format!("{detail}\nClose large documents or reduce caches and try again."),
                retry_label: Some("Retry".to_string()),
            },
            ErrorDialogKind::PanicRecovery => Self {
                kind,
                title: "Unexpected error".to_string(),
                body: format!("{detail}\nRecovery data may be available."),
                retry_label: Some("Open Recovery".to_string()),
            },
        }
    }
}

#[cfg(test)]
mod polish_tests {
    use super::{ErrorDialogKind, ErrorDialogState, LoadingKind, LoadingState};

    #[test]
    fn loading_factory_clamps_progress() {
        let state = LoadingState::for_kind(LoadingKind::LongOperation, Some(1.4));
        assert_eq!(state.progress, Some(1.0));
        assert!(state.message.contains("Working"));
    }

    #[test]
    fn save_failed_dialog_exposes_retry() {
        let dialog = ErrorDialogState::from_kind(ErrorDialogKind::SaveFailed, "Disk full");
        assert_eq!(dialog.retry_label.as_deref(), Some("Retry Save"));
        assert!(dialog.body.contains("different location"));
    }
}
