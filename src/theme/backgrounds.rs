use serde::{Deserialize, Serialize};

use crate::settings::schema::CanvasBackgroundPreference;
use crate::ui::Color;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum BackgroundFillMode {
    Fill,
    Fit,
    Stretch,
    Tile,
    Center,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum PatternStyle {
    Dots,
    LinesHorizontal,
    LinesVertical,
    LinesDiagonal,
    GraphPaper,
    CrossHatch,
    Noise,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum BackgroundKind {
    Solid { color: Color },
    Gradient {
        start: Color,
        end: Color,
        angle_degrees: f32,
    },
    Pattern {
        style: PatternStyle,
        foreground: Color,
        background: Color,
        scale: f32,
    },
    Image {
        path: String,
        mode: BackgroundFillMode,
        blur_px: f32,
        opacity: f32,
    },
    AnimatedGradient {
        colors: Vec<Color>,
        speed: f32,
    },
    Preset(String),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BackgroundSettings {
    pub kind: BackgroundKind,
    pub enable_particles: bool,
}

impl Default for BackgroundSettings {
    fn default() -> Self {
        Self {
            kind: BackgroundKind::Solid {
                color: Color::rgb(0.08, 0.1, 0.14),
            },
            enable_particles: false,
        }
    }
}

pub fn preset_by_id(id: &str) -> BackgroundSettings {
    match id.to_ascii_lowercase().as_str() {
        "clean" => BackgroundSettings {
            kind: BackgroundKind::Solid {
                color: Color::rgb(0.09, 0.11, 0.15),
            },
            enable_particles: false,
        },
        "paper" => BackgroundSettings {
            kind: BackgroundKind::Pattern {
                style: PatternStyle::Noise,
                foreground: Color::rgba(0.0, 0.0, 0.0, 0.025),
                background: Color::rgb(0.96, 0.95, 0.92),
                scale: 1.0,
            },
            enable_particles: false,
        },
        "blueprint" => BackgroundSettings {
            kind: BackgroundKind::Pattern {
                style: PatternStyle::GraphPaper,
                foreground: Color::rgba(0.72, 0.86, 1.0, 0.24),
                background: Color::rgb(0.07, 0.15, 0.30),
                scale: 1.0,
            },
            enable_particles: false,
        },
        "cozy" => BackgroundSettings {
            kind: BackgroundKind::Gradient {
                start: Color::rgb(0.28, 0.18, 0.14),
                end: Color::rgb(0.60, 0.34, 0.24),
                angle_degrees: 22.0,
            },
            enable_particles: false,
        },
        "midnight" => BackgroundSettings {
            kind: BackgroundKind::Pattern {
                style: PatternStyle::Dots,
                foreground: Color::rgba(0.82, 0.88, 1.0, 0.22),
                background: Color::rgb(0.03, 0.05, 0.11),
                scale: 1.0,
            },
            enable_particles: true,
        },
        _ => BackgroundSettings::default(),
    }
}

pub fn from_canvas_preference(preference: &CanvasBackgroundPreference) -> BackgroundSettings {
    let preset = preference.preset_id.trim();
    if preset.is_empty() {
        BackgroundSettings::default()
    } else {
        preset_by_id(preset)
    }
}
