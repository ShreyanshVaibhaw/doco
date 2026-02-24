#[derive(Debug, Clone, Copy)]
pub struct MicaSettings {
    pub enabled: bool,
    pub tint_opacity: f32,
}

impl Default for MicaSettings {
    fn default() -> Self {
        Self {
            enabled: true,
            tint_opacity: 0.8,
        }
    }
}
