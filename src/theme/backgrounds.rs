use serde::{Deserialize, Serialize};

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
