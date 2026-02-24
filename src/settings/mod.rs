pub mod schema;

use std::{
    fs,
    path::{Path, PathBuf},
    time::{Duration, Instant},
};

use windows::core::Result;

use schema::{Settings, SettingsCategory};

const SAVE_DEBOUNCE_MS: u64 = 500;

#[derive(Debug, Clone)]
pub struct SettingSearchHit {
    pub category: SettingsCategory,
    pub setting_key: &'static str,
    pub title: &'static str,
    pub summary: &'static str,
}

pub struct SettingsStore {
    path: PathBuf,
    settings: Settings,
    pending_write: bool,
    last_change_at: Option<Instant>,
    debounce: Duration,
}

impl SettingsStore {
    pub fn load() -> Result<Self> {
        let path = settings_path();
        let settings = load_settings()?;
        Ok(Self {
            path,
            settings,
            pending_write: false,
            last_change_at: None,
            debounce: Duration::from_millis(SAVE_DEBOUNCE_MS),
        })
    }

    pub fn with_path(path: impl Into<PathBuf>) -> Result<Self> {
        let path = path.into();
        let settings = load_settings_from(path.as_path())?;
        Ok(Self {
            path,
            settings,
            pending_write: false,
            last_change_at: None,
            debounce: Duration::from_millis(SAVE_DEBOUNCE_MS),
        })
    }

    pub fn settings(&self) -> &Settings {
        &self.settings
    }

    pub fn settings_mut(&mut self) -> &mut Settings {
        self.pending_write = true;
        self.last_change_at = Some(Instant::now());
        &mut self.settings
    }

    pub fn update<F>(&mut self, mutator: F)
    where
        F: FnOnce(&mut Settings),
    {
        mutator(&mut self.settings);
        self.pending_write = true;
        self.last_change_at = Some(Instant::now());
    }

    pub fn flush_if_due(&mut self) -> Result<bool> {
        let Some(last_change) = self.last_change_at else {
            return Ok(false);
        };
        if !self.pending_write || last_change.elapsed() < self.debounce {
            return Ok(false);
        }

        save_settings_to(self.path.as_path(), &self.settings)?;
        self.pending_write = false;
        self.last_change_at = None;
        Ok(true)
    }

    pub fn force_flush(&mut self) -> Result<()> {
        if self.pending_write {
            save_settings_to(self.path.as_path(), &self.settings)?;
            self.pending_write = false;
            self.last_change_at = None;
        }
        Ok(())
    }
}

pub fn settings_path() -> PathBuf {
    if let Some(root) = portable_root() {
        return root.join("settings.json");
    }

    if let Some(base) = dirs::config_dir() {
        base.join("Doco").join("settings.json")
    } else {
        PathBuf::from("settings.json")
    }
}

pub fn portable_root() -> Option<PathBuf> {
    let exe = std::env::current_exe().ok()?;
    let dir = exe.parent()?.to_path_buf();
    let marker = dir.join("doco.ini");
    if marker.exists() {
        Some(dir)
    } else {
        None
    }
}

pub fn load_settings() -> Result<Settings> {
    load_settings_from(settings_path().as_path())
}

pub fn save_settings(settings: &Settings) -> Result<()> {
    save_settings_to(settings_path().as_path(), settings)
}

pub fn load_settings_from(path: &Path) -> Result<Settings> {
    if let Ok(data) = fs::read_to_string(path) {
        if let Ok(settings) = serde_json::from_str::<Settings>(&data) {
            return Ok(settings.migrate());
        }
    }
    Ok(Settings::default())
}

pub fn save_settings_to(path: &Path, settings: &Settings) -> Result<()> {
    if let Some(parent) = path.parent() {
        let _ = fs::create_dir_all(parent);
    }
    if let Ok(data) = serde_json::to_string_pretty(&settings.clone().migrate()) {
        let _ = fs::write(path, data);
    }
    Ok(())
}

pub fn search_settings(query: &str) -> Vec<SettingSearchHit> {
    let needle = query.trim().to_ascii_lowercase();
    if needle.is_empty() {
        return settings_catalog().to_vec();
    }

    settings_catalog()
        .iter()
        .filter(|item| {
            item.title.to_ascii_lowercase().contains(needle.as_str())
                || item.summary.to_ascii_lowercase().contains(needle.as_str())
                || item.setting_key.to_ascii_lowercase().contains(needle.as_str())
                || item
                    .category
                    .title()
                    .to_ascii_lowercase()
                    .contains(needle.as_str())
        })
        .cloned()
        .collect()
}

fn settings_catalog() -> &'static [SettingSearchHit] {
    &[
        SettingSearchHit {
            category: SettingsCategory::Appearance,
            setting_key: "appearance.theme",
            title: "Theme",
            summary: "Choose system auto or a named theme.",
        },
        SettingSearchHit {
            category: SettingsCategory::Appearance,
            setting_key: "appearance.canvas_background",
            title: "Canvas Background",
            summary: "Preset paper/background styling plus custom option.",
        },
        SettingSearchHit {
            category: SettingsCategory::Appearance,
            setting_key: "appearance.ui_scale",
            title: "UI Scale",
            summary: "Scale shell UI independent of document zoom.",
        },
        SettingSearchHit {
            category: SettingsCategory::Appearance,
            setting_key: "appearance.sidebar_default_panel",
            title: "Sidebar Default Panel",
            summary: "Open Files, Outline, or Bookmarks by default.",
        },
        SettingSearchHit {
            category: SettingsCategory::Editor,
            setting_key: "editor.default_font_family",
            title: "Default Font Family",
            summary: "Choose the default editor font family.",
        },
        SettingSearchHit {
            category: SettingsCategory::Editor,
            setting_key: "editor.tab_size",
            title: "Tab Size",
            summary: "Use 2, 4, or 8 spaces for indentation.",
        },
        SettingSearchHit {
            category: SettingsCategory::Editor,
            setting_key: "editor.word_wrap",
            title: "Word Wrap",
            summary: "Enable wrapping or wrap at a specific column.",
        },
        SettingSearchHit {
            category: SettingsCategory::Editor,
            setting_key: "editor.cursor_style",
            title: "Cursor Style",
            summary: "Line, block, or underline cursor.",
        },
        SettingSearchHit {
            category: SettingsCategory::Document,
            setting_key: "document.default_page_size",
            title: "Default Page Size",
            summary: "Letter, A4, or Legal for new documents.",
        },
        SettingSearchHit {
            category: SettingsCategory::Document,
            setting_key: "document.default_view_mode",
            title: "Default View Mode",
            summary: "Page, continuous, or read mode.",
        },
        SettingSearchHit {
            category: SettingsCategory::Files,
            setting_key: "files.auto_save_interval",
            title: "Auto-save Interval",
            summary: "Configure auto-save frequency or disable it.",
        },
        SettingSearchHit {
            category: SettingsCategory::Files,
            setting_key: "files.default_save_format",
            title: "Default Save Format",
            summary: "Use .docx or another preferred format.",
        },
        SettingSearchHit {
            category: SettingsCategory::KeyboardShortcuts,
            setting_key: "keyboard_shortcuts.bindings",
            title: "Keyboard Shortcuts",
            summary: "Customize key bindings and detect conflicts.",
        },
        SettingSearchHit {
            category: SettingsCategory::Performance,
            setting_key: "performance.hardware_acceleration",
            title: "Hardware Acceleration",
            summary: "Toggle GPU rendering acceleration.",
        },
        SettingSearchHit {
            category: SettingsCategory::Performance,
            setting_key: "performance.max_image_cache_mb",
            title: "Image Cache Limit",
            summary: "Maximum memory available to decoded images.",
        },
        SettingSearchHit {
            category: SettingsCategory::About,
            setting_key: "about.version",
            title: "Version",
            summary: "Build and release version details.",
        },
    ]
}

