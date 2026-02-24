use crate::editor::commands::{shortcut_from_vk, Shortcut as EditorShortcut};

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

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ResolvedShortcut {
    App(AppShortcut),
    Editor(EditorShortcut),
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

pub fn resolve_editor_shortcut(vk: u32, modifiers: KeyModifiers) -> Option<EditorShortcut> {
    shortcut_from_vk(modifiers.ctrl, modifiers.shift, vk)
}

pub fn resolve_any_shortcut(vk: u32, modifiers: KeyModifiers) -> Option<ResolvedShortcut> {
    if let Some(editor) = resolve_editor_shortcut(vk, modifiers) {
        return Some(ResolvedShortcut::Editor(editor));
    }
    resolve_shortcut(vk, modifiers).map(ResolvedShortcut::App)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn resolves_editor_and_app_shortcuts() {
        let ctrl = KeyModifiers {
            ctrl: true,
            shift: false,
            alt: false,
        };
        assert_eq!(
            resolve_any_shortcut(0x42, ctrl),
            Some(ResolvedShortcut::Editor(EditorShortcut::Bold))
        );
        assert_eq!(resolve_any_shortcut(0xBC, ctrl), Some(ResolvedShortcut::App(AppShortcut::OpenSettings)));
    }
}
