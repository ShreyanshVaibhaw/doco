#[derive(Clone, Copy, Debug, Default)]
pub struct KeyModifiers {
    pub ctrl: bool,
    pub shift: bool,
    pub alt: bool,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum AppShortcut {
    OpenSettings,
    Print,
    ToggleDebugPanel,
}

#[derive(Clone, Debug)]
pub enum InputEvent {
    MouseMove { x: i32, y: i32 },
    MouseDown { x: i32, y: i32 },
    MouseUp { x: i32, y: i32 },
    MouseWheel { delta: i16, x: i32, y: i32 },
    KeyDown { vk: u32, modifiers: KeyModifiers },
    KeyUp { vk: u32, modifiers: KeyModifiers },
    Char(char),
    FilesDropped(usize),
}

pub fn resolve_shortcut(vk: u32, modifiers: KeyModifiers) -> Option<AppShortcut> {
    match (vk, modifiers.ctrl, modifiers.shift, modifiers.alt) {
        (0xBC, true, false, false) => Some(AppShortcut::OpenSettings), // Ctrl+,
        (0x50, true, false, false) => Some(AppShortcut::Print),         // Ctrl+P
        (0x44, true, true, false) => Some(AppShortcut::ToggleDebugPanel), // Ctrl+Shift+D
        _ => None,
    }
}
