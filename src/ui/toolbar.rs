use std::time::{Duration, Instant};

use windows::Win32::Graphics::Direct2D::ID2D1DeviceContext;

use crate::{
    theme::Theme,
    ui::{Color, InputEvent, Point, Rect, UIComponent},
};

const TOOLBAR_PADDING_X: f32 = 8.0;
const BUTTON_GAP: f32 = 4.0;
const BUTTON_HEIGHT: f32 = 32.0;
const SPLIT_DROPDOWN_WIDTH: f32 = 14.0;
const TOOLTIP_DELAY: Duration = Duration::from_millis(500);
const QUICK_ANIMATION: Duration = Duration::from_millis(100);
const DEFAULT_OVERFLOW_PANEL_WIDTH: f32 = 220.0;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ToolbarButtonType {
    Icon,
    Toggle,
    Dropdown,
    Split,
    Separator,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ToolbarAction {
    FileMenu,
    Cut,
    Copy,
    Paste,
    Undo,
    Redo,
    Bold,
    Italic,
    Underline,
    Strikethrough,
    FontFamily,
    FontSize,
    TextColor,
    AlignLeft,
    AlignCenter,
    AlignRight,
    AlignJustify,
    List,
    Heading,
    InsertImage,
    InsertLink,
    InsertTable,
    CommandPalette,
    More,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ToggleState {
    #[default]
    Off,
    On,
    Mixed,
}

impl ToggleState {
    pub fn is_on(self) -> bool {
        matches!(self, Self::On)
    }

    pub fn is_mixed(self) -> bool {
        matches!(self, Self::Mixed)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AlignmentState {
    Left,
    Center,
    Right,
    Justify,
    Mixed,
}

impl Default for AlignmentState {
    fn default() -> Self {
        Self::Left
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HeadingState {
    Normal,
    H1,
    H2,
    H3,
    H4,
    H5,
    H6,
    Mixed,
}

impl HeadingState {
    fn display_label(self) -> &'static str {
        match self {
            Self::Normal => "Normal",
            Self::H1 => "H1",
            Self::H2 => "H2",
            Self::H3 => "H3",
            Self::H4 => "H4",
            Self::H5 => "H5",
            Self::H6 => "H6",
            Self::Mixed => "Mixed",
        }
    }
}

impl Default for HeadingState {
    fn default() -> Self {
        Self::Normal
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ListState {
    None,
    Bulleted,
    Numbered,
    Checkbox,
    Mixed,
}

impl ListState {
    fn display_label(self) -> &'static str {
        match self {
            Self::None => "List",
            Self::Bulleted => "Bullet",
            Self::Numbered => "Numbered",
            Self::Checkbox => "Checkbox",
            Self::Mixed => "Mixed",
        }
    }
}

impl Default for ListState {
    fn default() -> Self {
        Self::None
    }
}

#[derive(Debug, Clone, Default)]
pub struct ToolbarFormatState {
    pub bold: ToggleState,
    pub italic: ToggleState,
    pub underline: ToggleState,
    pub strikethrough: ToggleState,
    pub superscript: ToggleState,
    pub subscript: ToggleState,
    pub font_family: String,
    pub font_size: f32,
    pub text_color: Option<Color>,
    pub highlight_color: Option<Color>,
    pub alignment: AlignmentState,
    pub heading: HeadingState,
    pub list: ListState,
}

#[derive(Debug, Clone)]
pub struct ToolbarButton {
    pub id: &'static str,
    pub label: String,
    pub tooltip: String,
    pub icon_glyph: &'static str,
    pub action: ToolbarAction,
    pub kind: ToolbarButtonType,
    pub width: f32,
    pub enabled: bool,
    pub active: bool,
    pub indeterminate: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ToolbarDropdownKind {
    File,
    Paste,
    FontFamily,
    FontSize,
    TextColor,
    Heading,
    List,
    More,
}

#[derive(Debug, Clone)]
pub struct FontPickerState {
    pub all_fonts: Vec<String>,
    pub recent_fonts: Vec<String>,
    pub query: String,
}

impl Default for FontPickerState {
    fn default() -> Self {
        Self {
            all_fonts: default_font_catalog(),
            recent_fonts: vec!["Segoe UI".to_string()],
            query: String::new(),
        }
    }
}

impl FontPickerState {
    pub fn set_query(&mut self, query: impl Into<String>) {
        self.query = query.into();
    }

    pub fn remember_font(&mut self, font_name: impl Into<String>) {
        let font = font_name.into();
        self.recent_fonts.retain(|f| !f.eq_ignore_ascii_case(&font));
        self.recent_fonts.insert(0, font);
        if self.recent_fonts.len() > 8 {
            self.recent_fonts.truncate(8);
        }
    }

    pub fn visible_fonts(&self) -> Vec<String> {
        let needle = self.query.trim().to_ascii_lowercase();
        let mut out = Vec::new();

        let include = |font: &str, q: &str| -> bool { q.is_empty() || font.to_ascii_lowercase().contains(q) };
        for font in &self.recent_fonts {
            if include(font, &needle) {
                out.push(font.clone());
            }
        }

        for font in &self.all_fonts {
            if !include(font, &needle) {
                continue;
            }
            if out.iter().any(|f| f.eq_ignore_ascii_case(font)) {
                continue;
            }
            out.push(font.clone());
        }

        out
    }
}

#[derive(Debug, Clone)]
pub struct SizePickerState {
    pub common_sizes: Vec<f32>,
    pub custom_input: Option<f32>,
}

impl Default for SizePickerState {
    fn default() -> Self {
        Self {
            common_sizes: vec![8.0, 9.0, 10.0, 11.0, 12.0, 14.0, 16.0, 18.0, 20.0, 24.0, 28.0, 32.0, 36.0, 48.0, 72.0],
            custom_input: None,
        }
    }
}

#[derive(Debug, Clone)]
pub struct ColorPickerState {
    pub columns: usize,
    pub theme_colors: Vec<Color>,
    pub recent_colors: Vec<Color>,
    pub selected: Option<Color>,
}

impl Default for ColorPickerState {
    fn default() -> Self {
        Self {
            columns: 10,
            theme_colors: default_theme_colors(),
            recent_colors: Vec::new(),
            selected: None,
        }
    }
}

impl ColorPickerState {
    pub fn remember_color(&mut self, color: Color) {
        self.recent_colors.retain(|c| *c != color);
        self.recent_colors.insert(0, color);
        if self.recent_colors.len() > 10 {
            self.recent_colors.truncate(10);
        }
        self.selected = Some(color);
    }
}

#[derive(Debug, Clone)]
pub struct HeadingPickerState {
    pub options: Vec<HeadingState>,
}

impl Default for HeadingPickerState {
    fn default() -> Self {
        Self {
            options: vec![
                HeadingState::Normal,
                HeadingState::H1,
                HeadingState::H2,
                HeadingState::H3,
                HeadingState::H4,
                HeadingState::H5,
                HeadingState::H6,
            ],
        }
    }
}

#[derive(Debug, Clone)]
pub struct ListPickerState {
    pub options: Vec<ListState>,
}

impl Default for ListPickerState {
    fn default() -> Self {
        Self {
            options: vec![ListState::Bulleted, ListState::Numbered, ListState::Checkbox, ListState::None],
        }
    }
}

#[derive(Debug, Clone)]
pub struct ToolbarDropdownState {
    pub open: Option<ToolbarDropdownKind>,
    pub anchor: Option<Rect>,
    pub panel_width: f32,
    pub font_picker: FontPickerState,
    pub size_picker: SizePickerState,
    pub color_picker: ColorPickerState,
    pub heading_picker: HeadingPickerState,
    pub list_picker: ListPickerState,
}

impl Default for ToolbarDropdownState {
    fn default() -> Self {
        Self {
            open: None,
            anchor: None,
            panel_width: DEFAULT_OVERFLOW_PANEL_WIDTH,
            font_picker: FontPickerState::default(),
            size_picker: SizePickerState::default(),
            color_picker: ColorPickerState::default(),
            heading_picker: HeadingPickerState::default(),
            list_picker: ListPickerState::default(),
        }
    }
}

impl ToolbarDropdownState {
    pub fn open(&mut self, kind: ToolbarDropdownKind, anchor: Rect) {
        self.open = Some(kind);
        self.anchor = Some(anchor);
    }

    pub fn close(&mut self) {
        self.open = None;
        self.anchor = None;
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ToolbarHitPart {
    Main,
    Dropdown,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ToolbarIntent {
    Action(ToolbarAction),
    OpenDropdown(ToolbarDropdownKind),
}

#[derive(Debug, Clone)]
pub struct Toolbar {
    bounds: Rect,
    visible: bool,
    all_buttons: Vec<ToolbarButton>,
    pub buttons: Vec<ToolbarButton>,
    pub overflow: Vec<ToolbarButton>,
    button_rects: Vec<Rect>,
    pub hovered_index: Option<usize>,
    hover_started: Option<Instant>,
    pressed_index: Option<usize>,
    pressed_started: Option<Instant>,
    pub show_tooltip: bool,
    pub tooltip_text: Option<String>,
    pub pending_intent: Option<ToolbarIntent>,
    pub dropdown: ToolbarDropdownState,
    pub format_state: ToolbarFormatState,
}

impl Default for Toolbar {
    fn default() -> Self {
        Self::new()
    }
}

impl Toolbar {
    pub fn new() -> Self {
        let all_buttons = default_buttons();
        let mut toolbar = Self {
            bounds: Rect::default(),
            visible: true,
            all_buttons: all_buttons.clone(),
            buttons: all_buttons,
            overflow: Vec::new(),
            button_rects: Vec::new(),
            hovered_index: None,
            hover_started: None,
            pressed_index: None,
            pressed_started: None,
            show_tooltip: false,
            tooltip_text: None,
            pending_intent: None,
            dropdown: ToolbarDropdownState::default(),
            format_state: ToolbarFormatState {
                font_family: "Segoe UI".to_string(),
                font_size: 12.0,
                ..ToolbarFormatState::default()
            },
        };
        toolbar.set_format_state(toolbar.format_state.clone());
        toolbar
    }

    pub fn recalc_overflow(&mut self, available_width: f32) {
        let more_button = self
            .all_buttons
            .iter()
            .find(|b| b.id == "more")
            .cloned()
            .unwrap_or_else(default_more_button);

        let mut core_buttons = self
            .all_buttons
            .iter()
            .filter(|b| b.id != "more")
            .cloned()
            .collect::<Vec<_>>();
        trim_separators(&mut core_buttons);

        let horizontal_padding = TOOLBAR_PADDING_X * 2.0;
        if available_width <= horizontal_padding + button_width(&more_button) {
            self.buttons = vec![more_button];
            self.overflow = core_buttons;
            trim_separators(&mut self.overflow);
            self.sync_button_states_from_format();
            return;
        }

        let mut visible = Vec::with_capacity(core_buttons.len() + 1);
        let mut overflow = Vec::new();
        let budget = (available_width - horizontal_padding - button_width(&more_button)).max(0.0);
        let mut used = 0.0_f32;

        for button in core_buttons {
            let needed = if visible.is_empty() {
                button_width(&button)
            } else {
                BUTTON_GAP + button_width(&button)
            };

            if used + needed <= budget {
                used += needed;
                visible.push(button);
            } else {
                overflow.push(button);
            }
        }

        trim_separators(&mut visible);
        trim_separators(&mut overflow);

        visible.push(more_button);
        self.buttons = visible;
        self.overflow = overflow;
        self.sync_button_states_from_format();
    }

    pub fn set_format_state(&mut self, state: ToolbarFormatState) {
        self.format_state = state;
        self.sync_button_states_from_format();
    }

    pub fn set_font_catalog(&mut self, fonts: Vec<String>) {
        self.dropdown.font_picker.all_fonts = fonts;
    }

    pub fn set_font_query(&mut self, query: impl Into<String>) {
        self.dropdown.font_picker.set_query(query);
    }

    pub fn visible_fonts(&self) -> Vec<String> {
        self.dropdown.font_picker.visible_fonts()
    }

    pub fn remember_font(&mut self, font_name: impl Into<String>) {
        self.dropdown.font_picker.remember_font(font_name);
    }

    pub fn remember_color(&mut self, color: Color) {
        self.dropdown.color_picker.remember_color(color);
    }

    pub fn button_rect(&self, index: usize) -> Option<Rect> {
        self.button_rects.get(index).copied()
    }

    pub fn hit_button_part(&self, point: Point) -> Option<(usize, ToolbarHitPart)> {
        if !self.visible {
            return None;
        }

        for (idx, rect) in self.button_rects.iter().enumerate() {
            if !contains(*rect, point) {
                continue;
            }

            let button = self.buttons.get(idx)?;
            if !button.enabled {
                return None;
            }

            let part = match button.kind {
                ToolbarButtonType::Split => {
                    if point.x >= rect.x + rect.width - SPLIT_DROPDOWN_WIDTH {
                        ToolbarHitPart::Dropdown
                    } else {
                        ToolbarHitPart::Main
                    }
                }
                ToolbarButtonType::Dropdown => ToolbarHitPart::Dropdown,
                _ => ToolbarHitPart::Main,
            };

            return Some((idx, part));
        }

        None
    }

    pub fn hit_button(&self, point: Point) -> Option<usize> {
        self.hit_button_part(point).map(|(idx, _)| idx)
    }

    pub fn begin_hover(&mut self, index: Option<usize>) {
        if self.hovered_index != index {
            self.hover_started = if index.is_some() { Some(Instant::now()) } else { None };
            self.show_tooltip = false;
            self.hovered_index = index;
            self.tooltip_text = index.and_then(|i| self.buttons.get(i).map(|b| b.tooltip.clone()));
            return;
        }

        if index.is_none() {
            self.show_tooltip = false;
            return;
        }

        if let Some(t0) = self.hover_started {
            self.show_tooltip = t0.elapsed() >= TOOLTIP_DELAY;
        }
    }

    pub fn hover_progress(&self, index: usize) -> f32 {
        if self.hovered_index != Some(index) {
            return 0.0;
        }
        if let Some(t0) = self.hover_started {
            return (t0.elapsed().as_secs_f32() / QUICK_ANIMATION.as_secs_f32()).clamp(0.0, 1.0);
        }
        0.0
    }

    pub fn press_progress(&self, index: usize) -> f32 {
        if self.pressed_index != Some(index) {
            return 0.0;
        }
        if let Some(t0) = self.pressed_started {
            return (t0.elapsed().as_secs_f32() / QUICK_ANIMATION.as_secs_f32()).clamp(0.0, 1.0);
        }
        0.0
    }

    pub fn invoke(&self, index: usize) -> Option<ToolbarAction> {
        self.buttons.get(index).filter(|b| b.enabled).map(|b| b.action)
    }

    pub fn invoke_with_point(&mut self, point: Point) -> Option<ToolbarIntent> {
        let (index, part) = self.hit_button_part(point)?;
        let button = self.buttons.get(index)?;
        let intent = match (button.kind, part) {
            (ToolbarButtonType::Dropdown, _) => {
                let kind = dropdown_kind_for_action(button.action)?;
                ToolbarIntent::OpenDropdown(kind)
            }
            (ToolbarButtonType::Split, ToolbarHitPart::Dropdown) => {
                let kind = dropdown_kind_for_action(button.action)?;
                ToolbarIntent::OpenDropdown(kind)
            }
            _ => ToolbarIntent::Action(button.action),
        };

        match intent {
            ToolbarIntent::OpenDropdown(kind) => {
                if let Some(anchor) = self.button_rect(index) {
                    self.dropdown.open(kind, anchor);
                }
            }
            ToolbarIntent::Action(_) => {}
        }

        self.pending_intent = Some(intent);
        Some(intent)
    }

    fn sync_button_states_from_format(&mut self) {
        for button in &mut self.all_buttons {
            apply_format_to_button(button, &self.format_state);
        }

        for button in &mut self.buttons {
            apply_format_to_button(button, &self.format_state);
        }

        for button in &mut self.overflow {
            apply_format_to_button(button, &self.format_state);
        }
    }

    fn layout_button_rects(&mut self) {
        self.button_rects.clear();
        if !self.visible || self.bounds.height <= 0.0 {
            return;
        }

        let mut x = self.bounds.x + TOOLBAR_PADDING_X;
        let y = self.bounds.y + ((self.bounds.height - BUTTON_HEIGHT).max(0.0) * 0.5);

        for button in &self.buttons {
            let rect = Rect {
                x,
                y,
                width: button_width(button),
                height: BUTTON_HEIGHT,
            };
            self.button_rects.push(rect);
            x += rect.width + BUTTON_GAP;
        }
    }
}

impl UIComponent for Toolbar {
    fn layout(&mut self, bounds: Rect, _dpi: f32) {
        self.bounds = bounds;
        self.recalc_overflow(bounds.width);
        self.layout_button_rects();
    }

    fn render(&self, _ctx: &ID2D1DeviceContext, _theme: &Theme) {
        // Toolbar visuals are currently rendered by the shell renderer.
    }

    fn handle_input(&mut self, event: &InputEvent) -> bool {
        match event {
            InputEvent::MouseMove(point) => {
                let idx = self.hit_button(*point);
                self.begin_hover(idx);
                idx.is_some()
            }
            InputEvent::MouseDown(point) => {
                self.pressed_index = self.hit_button(*point);
                self.pressed_started = self.pressed_index.map(|_| Instant::now());
                self.pressed_index.is_some()
            }
            InputEvent::MouseUp(point) => {
                let was_pressed = self.pressed_index.is_some();
                if was_pressed {
                    let _ = self.invoke_with_point(*point);
                }
                self.pressed_index = None;
                self.pressed_started = None;
                was_pressed
            }
            InputEvent::KeyDown(vk) => {
                if *vk == 0x1B {
                    self.dropdown.close();
                    return true;
                }
                false
            }
            _ => false,
        }
    }

    fn hit_test(&self, point: Point) -> bool {
        contains(self.bounds, point)
    }

    fn set_visible(&mut self, visible: bool) {
        self.visible = visible;
    }

    fn bounds(&self) -> Rect {
        self.bounds
    }
}

fn default_buttons() -> Vec<ToolbarButton> {
    vec![
        btn("file", "File", "File menu", "", ToolbarAction::FileMenu, ToolbarButtonType::Dropdown, 68.0),
        btn("cut", "", "Cut", "", ToolbarAction::Cut, ToolbarButtonType::Icon, 32.0),
        btn("copy", "", "Copy", "", ToolbarAction::Copy, ToolbarButtonType::Icon, 32.0),
        btn("paste", "", "Paste", "", ToolbarAction::Paste, ToolbarButtonType::Split, 40.0),
        btn("undo", "", "Undo", "", ToolbarAction::Undo, ToolbarButtonType::Icon, 32.0),
        btn("redo", "", "Redo", "", ToolbarAction::Redo, ToolbarButtonType::Icon, 32.0),
        sep(),
        btn("bold", "B", "Bold", "B", ToolbarAction::Bold, ToolbarButtonType::Toggle, 32.0),
        btn("italic", "I", "Italic", "I", ToolbarAction::Italic, ToolbarButtonType::Toggle, 32.0),
        btn("underline", "U", "Underline", "U", ToolbarAction::Underline, ToolbarButtonType::Toggle, 32.0),
        btn("strike", "S", "Strikethrough", "S", ToolbarAction::Strikethrough, ToolbarButtonType::Toggle, 32.0),
        sep(),
        btn("font", "Segoe UI", "Font family", "A", ToolbarAction::FontFamily, ToolbarButtonType::Dropdown, 128.0),
        btn("size", "12", "Font size", "12", ToolbarAction::FontSize, ToolbarButtonType::Dropdown, 64.0),
        btn("color", "Color", "Text color", "", ToolbarAction::TextColor, ToolbarButtonType::Dropdown, 68.0),
        sep(),
        btn("align_left", "", "Align left", "", ToolbarAction::AlignLeft, ToolbarButtonType::Toggle, 32.0),
        btn("align_center", "", "Align center", "", ToolbarAction::AlignCenter, ToolbarButtonType::Toggle, 32.0),
        btn("align_right", "", "Align right", "", ToolbarAction::AlignRight, ToolbarButtonType::Toggle, 32.0),
        btn("justify", "", "Align justify", "", ToolbarAction::AlignJustify, ToolbarButtonType::Toggle, 32.0),
        sep(),
        btn("list", "List", "Lists", "", ToolbarAction::List, ToolbarButtonType::Dropdown, 72.0),
        btn("heading", "Normal", "Heading styles", "", ToolbarAction::Heading, ToolbarButtonType::Dropdown, 76.0),
        btn("image", "", "Insert image", "", ToolbarAction::InsertImage, ToolbarButtonType::Icon, 32.0),
        btn("link", "", "Insert link", "", ToolbarAction::InsertLink, ToolbarButtonType::Icon, 32.0),
        btn("table", "", "Insert table", "", ToolbarAction::InsertTable, ToolbarButtonType::Icon, 32.0),
        btn("cmd", "", "Command palette", "", ToolbarAction::CommandPalette, ToolbarButtonType::Icon, 32.0),
        default_more_button(),
    ]
}

fn btn(
    id: &'static str,
    label: &'static str,
    tooltip: &'static str,
    icon_glyph: &'static str,
    action: ToolbarAction,
    kind: ToolbarButtonType,
    width: f32,
) -> ToolbarButton {
    ToolbarButton {
        id,
        label: label.to_string(),
        tooltip: tooltip.to_string(),
        icon_glyph,
        action,
        kind,
        width,
        enabled: true,
        active: false,
        indeterminate: false,
    }
}

fn default_more_button() -> ToolbarButton {
    btn("more", "...", "More actions", "…", ToolbarAction::More, ToolbarButtonType::Dropdown, 40.0)
}

fn sep() -> ToolbarButton {
    ToolbarButton {
        id: "sep",
        label: String::new(),
        tooltip: String::new(),
        icon_glyph: "",
        action: ToolbarAction::More,
        kind: ToolbarButtonType::Separator,
        width: 8.0,
        enabled: false,
        active: false,
        indeterminate: false,
    }
}

fn button_width(button: &ToolbarButton) -> f32 {
    match button.kind {
        ToolbarButtonType::Separator => button.width.max(8.0),
        _ => button.width.max(32.0),
    }
}

fn trim_separators(buttons: &mut Vec<ToolbarButton>) {
    while matches!(buttons.first().map(|b| b.kind), Some(ToolbarButtonType::Separator)) {
        buttons.remove(0);
    }
    while matches!(buttons.last().map(|b| b.kind), Some(ToolbarButtonType::Separator)) {
        buttons.pop();
    }

    let mut i = 1;
    while i < buttons.len() {
        if buttons[i - 1].kind == ToolbarButtonType::Separator
            && buttons[i].kind == ToolbarButtonType::Separator
        {
            buttons.remove(i);
        } else {
            i += 1;
        }
    }
}

fn apply_format_to_button(button: &mut ToolbarButton, state: &ToolbarFormatState) {
    button.active = false;
    button.indeterminate = false;

    match button.action {
        ToolbarAction::Bold => {
            button.active = state.bold.is_on();
            button.indeterminate = state.bold.is_mixed();
        }
        ToolbarAction::Italic => {
            button.active = state.italic.is_on();
            button.indeterminate = state.italic.is_mixed();
        }
        ToolbarAction::Underline => {
            button.active = state.underline.is_on();
            button.indeterminate = state.underline.is_mixed();
        }
        ToolbarAction::Strikethrough => {
            button.active = state.strikethrough.is_on();
            button.indeterminate = state.strikethrough.is_mixed();
        }
        ToolbarAction::AlignLeft => {
            button.active = matches!(state.alignment, AlignmentState::Left);
            button.indeterminate = matches!(state.alignment, AlignmentState::Mixed);
        }
        ToolbarAction::AlignCenter => {
            button.active = matches!(state.alignment, AlignmentState::Center);
            button.indeterminate = matches!(state.alignment, AlignmentState::Mixed);
        }
        ToolbarAction::AlignRight => {
            button.active = matches!(state.alignment, AlignmentState::Right);
            button.indeterminate = matches!(state.alignment, AlignmentState::Mixed);
        }
        ToolbarAction::AlignJustify => {
            button.active = matches!(state.alignment, AlignmentState::Justify);
            button.indeterminate = matches!(state.alignment, AlignmentState::Mixed);
        }
        ToolbarAction::FontFamily => {
            button.label = truncate_label(&state.font_family, 16);
        }
        ToolbarAction::FontSize => {
            button.label = format_size(state.font_size);
        }
        ToolbarAction::Heading => {
            button.label = state.heading.display_label().to_string();
        }
        ToolbarAction::List => {
            button.label = state.list.display_label().to_string();
        }
        _ => {}
    }
}

fn format_size(size: f32) -> String {
    if (size - size.round()).abs() < f32::EPSILON {
        format!("{}", size.round() as i32)
    } else {
        format!("{size:.1}")
    }
}

fn truncate_label(text: &str, max_chars: usize) -> String {
    if text.chars().count() <= max_chars {
        return text.to_string();
    }

    let mut out = String::new();
    for ch in text.chars().take(max_chars.saturating_sub(1)) {
        out.push(ch);
    }
    out.push('…');
    out
}

fn dropdown_kind_for_action(action: ToolbarAction) -> Option<ToolbarDropdownKind> {
    match action {
        ToolbarAction::FileMenu => Some(ToolbarDropdownKind::File),
        ToolbarAction::Paste => Some(ToolbarDropdownKind::Paste),
        ToolbarAction::FontFamily => Some(ToolbarDropdownKind::FontFamily),
        ToolbarAction::FontSize => Some(ToolbarDropdownKind::FontSize),
        ToolbarAction::TextColor => Some(ToolbarDropdownKind::TextColor),
        ToolbarAction::Heading => Some(ToolbarDropdownKind::Heading),
        ToolbarAction::List => Some(ToolbarDropdownKind::List),
        ToolbarAction::More => Some(ToolbarDropdownKind::More),
        _ => None,
    }
}

fn default_font_catalog() -> Vec<String> {
    vec![
        "Segoe UI".to_string(),
        "Calibri".to_string(),
        "Arial".to_string(),
        "Cambria".to_string(),
        "Consolas".to_string(),
        "Georgia".to_string(),
        "Times New Roman".to_string(),
        "Verdana".to_string(),
        "Tahoma".to_string(),
        "Trebuchet MS".to_string(),
        "Courier New".to_string(),
    ]
}

fn default_theme_colors() -> Vec<Color> {
    vec![
        rgb(0x00, 0x00, 0x00),
        rgb(0x43, 0x43, 0x43),
        rgb(0x66, 0x66, 0x66),
        rgb(0x99, 0x99, 0x99),
        rgb(0xB7, 0xB7, 0xB7),
        rgb(0xCC, 0xCC, 0xCC),
        rgb(0xD9, 0xD9, 0xD9),
        rgb(0xEF, 0xEF, 0xEF),
        rgb(0xF3, 0xF3, 0xF3),
        rgb(0xFF, 0xFF, 0xFF),
        rgb(0x98, 0x00, 0x00),
        rgb(0xFF, 0x00, 0x00),
        rgb(0xFF, 0x99, 0x00),
        rgb(0xFF, 0xFF, 0x00),
        rgb(0x00, 0xFF, 0x00),
        rgb(0x00, 0xFF, 0xFF),
        rgb(0x4A, 0x86, 0xE8),
        rgb(0x00, 0x00, 0xFF),
        rgb(0x99, 0x00, 0xFF),
        rgb(0xFF, 0x00, 0xFF),
        rgb(0xE6, 0xB8, 0xAF),
        rgb(0xF4, 0xCC, 0xCC),
        rgb(0xFC, 0xE5, 0xCD),
        rgb(0xFF, 0xF2, 0xCC),
        rgb(0xD9, 0xEA, 0xD3),
        rgb(0xD0, 0xE0, 0xE3),
        rgb(0xC9, 0xDA, 0xF8),
        rgb(0xC9, 0xDA, 0xF8),
        rgb(0xD9, 0xD2, 0xE9),
        rgb(0xEA, 0xD1, 0xDC),
    ]
}

fn rgb(r: u8, g: u8, b: u8) -> Color {
    Color::rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0)
}

fn contains(rect: Rect, point: Point) -> bool {
    point.x >= rect.x
        && point.x <= rect.x + rect.width
        && point.y >= rect.y
        && point.y <= rect.y + rect.height
}

#[cfg(test)]
mod tests {
    use super::*;

    fn find_button<'a>(toolbar: &'a Toolbar, id: &str) -> Option<&'a ToolbarButton> {
        toolbar.buttons.iter().find(|b| b.id == id)
    }

    fn find_button_index(toolbar: &Toolbar, id: &str) -> Option<usize> {
        toolbar.buttons.iter().position(|b| b.id == id)
    }

    #[test]
    fn overflow_moves_tail_into_more_menu() {
        let mut toolbar = Toolbar::new();
        toolbar.layout(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 420.0,
                height: 44.0,
            },
            96.0,
        );

        assert!(!toolbar.overflow.is_empty());
        assert_eq!(toolbar.buttons.last().map(|b| b.id), Some("more"));
        assert!(toolbar.overflow.iter().any(|b| b.id == "heading" || b.id == "table"));
    }

    #[test]
    fn format_state_updates_toggle_and_labels() {
        let mut toolbar = Toolbar::new();
        toolbar.set_format_state(ToolbarFormatState {
            bold: ToggleState::On,
            italic: ToggleState::Mixed,
            underline: ToggleState::Off,
            strikethrough: ToggleState::On,
            font_family: "Segoe Print".to_string(),
            font_size: 14.0,
            alignment: AlignmentState::Center,
            heading: HeadingState::H2,
            list: ListState::Numbered,
            ..ToolbarFormatState::default()
        });
        toolbar.layout(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 1600.0,
                height: 44.0,
            },
            96.0,
        );

        let bold = find_button(&toolbar, "bold").expect("bold button missing");
        let italic = find_button(&toolbar, "italic").expect("italic button missing");
        let align_center = find_button(&toolbar, "align_center").expect("align_center missing");
        let heading = find_button(&toolbar, "heading").expect("heading button missing");
        let list = find_button(&toolbar, "list").expect("list button missing");
        let size = find_button(&toolbar, "size").expect("size button missing");

        assert!(bold.active);
        assert!(italic.indeterminate);
        assert!(align_center.active);
        assert_eq!(heading.label, "H2");
        assert_eq!(list.label, "Numbered");
        assert_eq!(size.label, "14");
    }

    #[test]
    fn split_button_reports_dropdown_zone() {
        let mut toolbar = Toolbar::new();
        toolbar.layout(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 1400.0,
                height: 44.0,
            },
            96.0,
        );

        let idx = find_button_index(&toolbar, "paste").expect("paste button missing");
        let rect = toolbar.button_rect(idx).expect("paste rect missing");
        let point = Point {
            x: rect.x + rect.width - 2.0,
            y: rect.y + rect.height * 0.5,
        };

        assert_eq!(toolbar.hit_button_part(point), Some((idx, ToolbarHitPart::Dropdown)));
    }

    #[test]
    fn tooltip_shows_after_hover_delay() {
        let mut toolbar = Toolbar::new();
        toolbar.layout(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 1400.0,
                height: 44.0,
            },
            96.0,
        );

        let idx = find_button_index(&toolbar, "copy").expect("copy button missing");
        toolbar.begin_hover(Some(idx));
        toolbar.hover_started = Some(Instant::now() - Duration::from_millis(600));
        toolbar.begin_hover(Some(idx));
        assert!(toolbar.show_tooltip);
    }

    #[test]
    fn font_picker_prioritizes_recent_fonts() {
        let mut toolbar = Toolbar::new();
        toolbar.set_font_catalog(vec![
            "Arial".to_string(),
            "Segoe UI".to_string(),
            "Segoe Print".to_string(),
            "Tahoma".to_string(),
        ]);
        toolbar.remember_font("Segoe Print");
        toolbar.set_font_query("segoe");

        let visible = toolbar.visible_fonts();
        assert_eq!(visible.first().map(|s| s.as_str()), Some("Segoe Print"));
        assert!(visible.iter().any(|f| f == "Segoe UI"));
    }
}
