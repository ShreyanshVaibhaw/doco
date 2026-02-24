use std::{
    collections::HashMap,
    fs,
    path::PathBuf,
    sync::Arc,
};

use parking_lot::RwLock;
use serde::{Deserialize, Serialize};
use windows::core::Result;

use crate::ui::Color;

pub mod backgrounds;
pub mod colors;
pub mod mica;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Theme {
    pub name: String,
    pub is_dark: bool,
    pub window_bg: Color,
    pub mica_tint: Color,
    pub mica_opacity: f32,
    pub surface_primary: Color,
    pub surface_secondary: Color,
    pub surface_tertiary: Color,
    pub surface_hover: Color,
    pub surface_pressed: Color,
    pub surface_selected: Color,
    pub border_default: Color,
    pub border_subtle: Color,
    pub border_focus: Color,
    pub text_primary: Color,
    pub text_secondary: Color,
    pub text_disabled: Color,
    pub text_accent: Color,
    pub text_on_accent: Color,
    pub accent: Color,
    pub accent_hover: Color,
    pub accent_pressed: Color,
    pub canvas_bg: Color,
    pub page_bg: Color,
    pub page_shadow: Color,
    pub scrollbar_track: Color,
    pub scrollbar_thumb: Color,
    pub scrollbar_thumb_hover: Color,
    pub heading_color: Color,
    pub link_color: Color,
    pub selection_bg: Color,
    pub cursor_color: Color,
    pub line_number_color: Color,
}

pub struct ThemeManager {
    shared: Arc<RwLock<Theme>>,
    themes: HashMap<String, Theme>,
}

impl ThemeManager {
    pub fn load() -> Result<Self> {
        let mut themes = built_in_themes()
            .into_iter()
            .map(|t| (t.name.clone(), t))
            .collect::<HashMap<_, _>>();

        for path in [PathBuf::from("assets/themes"), appdata_themes_dir()] {
            if let Ok(entries) = fs::read_dir(path) {
                for entry in entries.flatten() {
                    let p = entry.path();
                    if p.extension().and_then(|s| s.to_str()) == Some("json") {
                        if let Ok(text) = fs::read_to_string(&p) {
                            if let Ok(theme) = serde_json::from_str::<Theme>(&text) {
                                themes.insert(theme.name.clone(), theme);
                            }
                        }
                    }
                }
            }
        }

        let active = themes
            .get("Dark")
            .cloned()
            .or_else(|| themes.values().next().cloned())
            .unwrap_or_else(default_dark_theme);

        Ok(Self {
            shared: Arc::new(RwLock::new(active)),
            themes,
        })
    }

    pub fn active(&self) -> Theme {
        self.shared.read().clone()
    }

    pub fn shared(&self) -> Arc<RwLock<Theme>> {
        Arc::clone(&self.shared)
    }

    pub fn set_active(&self, name: &str) -> bool {
        if let Some(next) = self.themes.get(name) {
            *self.shared.write() = next.clone();
            true
        } else {
            false
        }
    }

    pub fn names(&self) -> Vec<String> {
        let mut names = self.themes.keys().cloned().collect::<Vec<_>>();
        names.sort();
        names
    }
}

fn appdata_themes_dir() -> PathBuf {
    if let Some(portable) = crate::settings::portable_root() {
        return portable.join("themes");
    }

    if let Some(base) = dirs::data_dir() {
        base.join("Doco").join("themes")
    } else {
        PathBuf::from("assets/themes")
    }
}

fn c(hex: u32) -> Color {
    let r = ((hex >> 16) & 0xFF) as f32 / 255.0;
    let g = ((hex >> 8) & 0xFF) as f32 / 255.0;
    let b = (hex & 0xFF) as f32 / 255.0;
    Color::rgb(r, g, b)
}

fn theme_template(name: &str, is_dark: bool, bg: u32, surface: u32, accent: u32, text: u32) -> Theme {
    Theme {
        name: name.to_string(),
        is_dark,
        window_bg: c(bg),
        mica_tint: c(bg),
        mica_opacity: if is_dark { 0.85 } else { 0.75 },
        surface_primary: c(surface),
        surface_secondary: c(if is_dark { 0x252A34 } else { 0xF2F3F5 }),
        surface_tertiary: c(if is_dark { 0x2A2F3A } else { 0xECEDEE }),
        surface_hover: c(if is_dark { 0x343B49 } else { 0xE7E9ED }),
        surface_pressed: c(if is_dark { 0x3D4657 } else { 0xDCE0E6 }),
        surface_selected: c(if is_dark { 0x2C3441 } else { 0xE4E8EE }),
        border_default: c(if is_dark { 0x3A4355 } else { 0xD0D6DE }),
        border_subtle: c(if is_dark { 0x303845 } else { 0xE0E5EC }),
        border_focus: c(accent),
        text_primary: c(text),
        text_secondary: c(if is_dark { 0xB4BACC } else { 0x4A566B }),
        text_disabled: c(if is_dark { 0x6C7385 } else { 0x9BA4B1 }),
        text_accent: c(accent),
        text_on_accent: c(if is_dark { 0x101317 } else { 0xFFFFFF }),
        accent: c(accent),
        accent_hover: c(if is_dark { 0x6BA8FF } else { 0x005DBF }),
        accent_pressed: c(if is_dark { 0x4B8FEF } else { 0x004A9A }),
        canvas_bg: c(if is_dark { 0x11151D } else { 0xEEF2F8 }),
        page_bg: c(if is_dark { 0x1B202A } else { 0xFFFFFF }),
        page_shadow: c(if is_dark { 0x000000 } else { 0xB7C0CF }),
        scrollbar_track: c(if is_dark { 0x1A1F29 } else { 0xE7EBF1 }),
        scrollbar_thumb: c(if is_dark { 0x495369 } else { 0xC1CAD7 }),
        scrollbar_thumb_hover: c(if is_dark { 0x5D6984 } else { 0xA9B5C6 }),
        heading_color: c(text),
        link_color: c(accent),
        selection_bg: Color::rgba(c(accent).r, c(accent).g, c(accent).b, 0.33),
        cursor_color: c(text),
        line_number_color: c(if is_dark { 0x8992A8 } else { 0x8B94A4 }),
    }
}

pub fn default_dark_theme() -> Theme {
    theme_template("Dark", true, 0x141821, 0x1D2330, 0x5EA1FF, 0xE6EAF2)
}

pub fn built_in_themes() -> Vec<Theme> {
    vec![
        theme_template("Light", false, 0xF4F7FC, 0xFFFFFF, 0x0A6DDA, 0x1E2530),
        default_dark_theme(),
        theme_template("Nord", true, 0x2E3440, 0x3B4252, 0x88C0D0, 0xECEFF4),
        theme_template("Catppuccin Mocha", true, 0x1E1E2E, 0x313244, 0x89B4FA, 0xCDD6F4),
        theme_template("Catppuccin Latte", false, 0xEFF1F5, 0xE6E9EF, 0x1E66F5, 0x4C4F69),
        theme_template("Solarized Light", false, 0xFDF6E3, 0xEEE8D5, 0x268BD2, 0x586E75),
        theme_template("Solarized Dark", true, 0x002B36, 0x073642, 0x268BD2, 0x93A1A1),
        theme_template("One Dark", true, 0x282C34, 0x2F343F, 0x61AFEF, 0xABB2BF),
        theme_template("Dracula", true, 0x282A36, 0x343746, 0xBD93F9, 0xF8F8F2),
        theme_template("Gruvbox Dark", true, 0x282828, 0x3C3836, 0xD79921, 0xEBDBB2),
    ]
}

