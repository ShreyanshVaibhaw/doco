use std::collections::BTreeMap;

use chrono::{DateTime, Utc};
use serde::{Deserialize, Serialize};

pub const SETTINGS_SCHEMA_VERSION: u32 = 1;

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub enum SettingsCategory {
    Appearance,
    Editor,
    Document,
    Files,
    KeyboardShortcuts,
    Performance,
    About,
}

impl SettingsCategory {
    pub const fn title(self) -> &'static str {
        match self {
            Self::Appearance => "Appearance",
            Self::Editor => "Editor",
            Self::Document => "Document",
            Self::Files => "Files",
            Self::KeyboardShortcuts => "Keyboard Shortcuts",
            Self::Performance => "Performance",
            Self::About => "About",
        }
    }

    pub const fn all() -> [Self; 7] {
        [
            Self::Appearance,
            Self::Editor,
            Self::Document,
            Self::Files,
            Self::KeyboardShortcuts,
            Self::Performance,
            Self::About,
        ]
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct Settings {
    pub schema_version: u32,
    pub appearance: AppearanceSettings,
    pub editor: EditorSettings,
    pub document: DocumentSettings,
    pub files: FileSettings,
    pub keyboard_shortcuts: KeyboardShortcutsSettings,
    pub performance: PerformanceSettings,
    pub about: AboutSettings,
}

impl Default for Settings {
    fn default() -> Self {
        Self {
            schema_version: SETTINGS_SCHEMA_VERSION,
            appearance: AppearanceSettings::default(),
            editor: EditorSettings::default(),
            document: DocumentSettings::default(),
            files: FileSettings::default(),
            keyboard_shortcuts: KeyboardShortcutsSettings::default(),
            performance: PerformanceSettings::default(),
            about: AboutSettings::default(),
        }
    }
}

impl Settings {
    pub fn migrate(mut self) -> Self {
        if self.schema_version > SETTINGS_SCHEMA_VERSION {
            return self;
        }

        self.schema_version = SETTINGS_SCHEMA_VERSION;
        self
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct AppearanceSettings {
    pub theme: ThemePreference,
    pub canvas_background: CanvasBackgroundPreference,
    pub ui_font: String,
    pub ui_scale: UiScale,
    pub show_toolbar: bool,
    pub show_sidebar: bool,
    pub show_status_bar: bool,
    pub show_tab_bar: bool,
    pub sidebar_default_panel: SidebarDefaultPanel,
}

impl Default for AppearanceSettings {
    fn default() -> Self {
        Self {
            theme: ThemePreference::SystemAuto,
            canvas_background: CanvasBackgroundPreference::default(),
            ui_font: "Segoe UI Variable".to_string(),
            ui_scale: UiScale::Percent100,
            show_toolbar: true,
            show_sidebar: true,
            show_status_bar: true,
            show_tab_bar: true,
            sidebar_default_panel: SidebarDefaultPanel::Files,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum ThemePreference {
    SystemAuto,
    Named(String),
}

impl Default for ThemePreference {
    fn default() -> Self {
        Self::SystemAuto
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct CanvasBackgroundPreference {
    pub preset_id: String,
    pub custom_payload: Option<String>,
}

impl Default for CanvasBackgroundPreference {
    fn default() -> Self {
        Self {
            preset_id: "paper".to_string(),
            custom_payload: None,
        }
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum UiScale {
    Percent100,
    Percent125,
    Percent150,
    Percent175,
    Percent200,
}

impl UiScale {
    pub fn as_factor(self) -> f32 {
        match self {
            Self::Percent100 => 1.0,
            Self::Percent125 => 1.25,
            Self::Percent150 => 1.5,
            Self::Percent175 => 1.75,
            Self::Percent200 => 2.0,
        }
    }
}

impl Default for UiScale {
    fn default() -> Self {
        Self::Percent100
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum SidebarDefaultPanel {
    Files,
    Outline,
    Bookmarks,
}

impl Default for SidebarDefaultPanel {
    fn default() -> Self {
        Self::Files
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct EditorSettings {
    pub default_font_family: String,
    pub default_font_size_pt: u16,
    pub tab_size: u8,
    pub insert_spaces_instead_of_tabs: bool,
    pub word_wrap: WordWrapMode,
    pub show_line_numbers: bool,
    pub cursor_style: CursorStyle,
    pub cursor_blink: bool,
    pub auto_indent: bool,
    pub auto_close_brackets: bool,
    pub show_whitespace: ShowWhitespaceMode,
}

impl Default for EditorSettings {
    fn default() -> Self {
        Self {
            default_font_family: "Segoe UI".to_string(),
            default_font_size_pt: 12,
            tab_size: 4,
            insert_spaces_instead_of_tabs: true,
            word_wrap: WordWrapMode::On,
            show_line_numbers: false,
            cursor_style: CursorStyle::Line,
            cursor_blink: true,
            auto_indent: true,
            auto_close_brackets: true,
            show_whitespace: ShowWhitespaceMode::Off,
        }
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum WordWrapMode {
    On,
    Off,
    AtColumn(u16),
}

impl Default for WordWrapMode {
    fn default() -> Self {
        Self::On
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum CursorStyle {
    Line,
    Block,
    Underline,
}

impl Default for CursorStyle {
    fn default() -> Self {
        Self::Line
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum ShowWhitespaceMode {
    Off,
    Selection,
    All,
}

impl Default for ShowWhitespaceMode {
    fn default() -> Self {
        Self::Off
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct DocumentSettings {
    pub default_page_size: DefaultPageSize,
    pub default_margins: DefaultMargins,
    pub default_line_spacing: f32,
    pub default_view_mode: DefaultViewMode,
    pub default_zoom_percent: u16,
    pub spelling_check: bool,
}

impl Default for DocumentSettings {
    fn default() -> Self {
        Self {
            default_page_size: DefaultPageSize::Letter,
            default_margins: DefaultMargins::Normal,
            default_line_spacing: 1.15,
            default_view_mode: DefaultViewMode::Page,
            default_zoom_percent: 100,
            spelling_check: true,
        }
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum DefaultPageSize {
    Letter,
    A4,
    Legal,
}

impl Default for DefaultPageSize {
    fn default() -> Self {
        Self::Letter
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum DefaultMargins {
    Normal,
    Narrow,
    Wide,
}

impl Default for DefaultMargins {
    fn default() -> Self {
        Self::Normal
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum DefaultViewMode {
    Page,
    Continuous,
    Read,
}

impl Default for DefaultViewMode {
    fn default() -> Self {
        Self::Page
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct FileSettings {
    pub auto_save_interval: AutoSaveInterval,
    pub create_backup_before_save: bool,
    pub default_save_format: String,
    pub recent_files_count: u16,
    pub default_open_folder: DefaultOpenFolder,
}

impl Default for FileSettings {
    fn default() -> Self {
        Self {
            auto_save_interval: AutoSaveInterval::Seconds(60),
            create_backup_before_save: true,
            default_save_format: ".docx".to_string(),
            recent_files_count: 20,
            default_open_folder: DefaultOpenFolder::LastUsed,
        }
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum AutoSaveInterval {
    Off,
    Seconds(u64),
}

impl AutoSaveInterval {
    pub fn as_seconds(self) -> Option<u64> {
        match self {
            Self::Off => None,
            Self::Seconds(v) => Some(v),
        }
    }
}

impl Default for AutoSaveInterval {
    fn default() -> Self {
        Self::Seconds(60)
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub enum DefaultOpenFolder {
    LastUsed,
    Documents,
    SpecificPath(String),
}

impl Default for DefaultOpenFolder {
    fn default() -> Self {
        Self::LastUsed
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct KeyboardShortcutsSettings {
    pub bindings: BTreeMap<String, ShortcutBinding>,
}

impl Default for KeyboardShortcutsSettings {
    fn default() -> Self {
        Self {
            bindings: default_shortcuts(),
        }
    }
}

impl KeyboardShortcutsSettings {
    pub fn set_binding(&mut self, command_id: &str, new_keys: impl Into<String>) {
        let entry = self
            .bindings
            .entry(command_id.to_string())
            .or_insert_with(|| ShortcutBinding::new("", ""));
        entry.keys = new_keys.into();
        entry.customized = true;
    }

    pub fn reset_to_defaults(&mut self) {
        self.bindings = default_shortcuts();
    }

    pub fn detect_conflicts(&self) -> Vec<ShortcutConflict> {
        let mut by_keys: BTreeMap<String, Vec<String>> = BTreeMap::new();
        for (command, binding) in &self.bindings {
            if binding.keys.trim().is_empty() {
                continue;
            }
            by_keys
                .entry(binding.keys.to_ascii_lowercase())
                .or_default()
                .push(command.clone());
        }

        by_keys
            .into_iter()
            .filter_map(|(keys, commands)| {
                if commands.len() > 1 {
                    Some(ShortcutConflict { keys, commands })
                } else {
                    None
                }
            })
            .collect()
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct ShortcutBinding {
    pub label: String,
    pub keys: String,
    pub default_keys: String,
    pub customized: bool,
}

impl ShortcutBinding {
    pub fn new(label: impl Into<String>, keys: impl Into<String>) -> Self {
        let keys = keys.into();
        Self {
            label: label.into(),
            keys: keys.clone(),
            default_keys: keys,
            customized: false,
        }
    }
}

impl Default for ShortcutBinding {
    fn default() -> Self {
        Self {
            label: String::new(),
            keys: String::new(),
            default_keys: String::new(),
            customized: false,
        }
    }
}

#[derive(Debug, Clone)]
pub struct ShortcutConflict {
    pub keys: String,
    pub commands: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct PerformanceSettings {
    pub hardware_acceleration: bool,
    pub max_undo_history: usize,
    pub background_pattern_quality: PatternQuality,
    pub animated_backgrounds: bool,
    pub max_image_cache_mb: u32,
}

impl Default for PerformanceSettings {
    fn default() -> Self {
        Self {
            hardware_acceleration: true,
            max_undo_history: 1000,
            background_pattern_quality: PatternQuality::High,
            animated_backgrounds: false,
            max_image_cache_mb: 200,
        }
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
pub enum PatternQuality {
    High,
    Low,
}

impl Default for PatternQuality {
    fn default() -> Self {
        Self::High
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct AboutSettings {
    pub version: String,
    pub check_updates_on_startup: bool,
    pub last_update_check_utc: Option<DateTime<Utc>>,
    pub licenses_url: String,
    pub system_info_snapshot: String,
}

impl Default for AboutSettings {
    fn default() -> Self {
        Self {
            version: env!("CARGO_PKG_VERSION").to_string(),
            check_updates_on_startup: true,
            last_update_check_utc: None,
            licenses_url: "https://opensource.org/licenses".to_string(),
            system_info_snapshot: String::new(),
        }
    }
}

fn default_shortcuts() -> BTreeMap<String, ShortcutBinding> {
    let mut map = BTreeMap::new();
    map.insert(
        "file.new".to_string(),
        ShortcutBinding::new("New Document", "Ctrl+N"),
    );
    map.insert(
        "file.open".to_string(),
        ShortcutBinding::new("Open", "Ctrl+O"),
    );
    map.insert(
        "file.save".to_string(),
        ShortcutBinding::new("Save", "Ctrl+S"),
    );
    map.insert(
        "edit.find".to_string(),
        ShortcutBinding::new("Find", "Ctrl+F"),
    );
    map.insert(
        "edit.replace".to_string(),
        ShortcutBinding::new("Replace", "Ctrl+H"),
    );
    map.insert(
        "view.command_palette".to_string(),
        ShortcutBinding::new("Command Palette", "Ctrl+Shift+P"),
    );
    map.insert(
        "view.settings".to_string(),
        ShortcutBinding::new("Settings", "Ctrl+,"),
    );
    map.insert(
        "view.debug_panel".to_string(),
        ShortcutBinding::new("Debug Panel", "Ctrl+Shift+D"),
    );
    map
}
