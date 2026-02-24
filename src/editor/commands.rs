use crate::{
    document::model::{
        Block,
        BlockId,
        ListType,
        Paragraph,
        ParagraphAlignment,
        Run,
        RunStyle,
    },
    ui::Color,
};

#[derive(Debug, Clone)]
pub enum EditCommand {
    InsertText {
        block_id: BlockId,
        offset: usize,
        text: String,
    },
    DeleteText {
        block_id: BlockId,
        start: usize,
        end: usize,
    },
    ReplaceText {
        block_id: BlockId,
        start: usize,
        end: usize,
        text: String,
    },
    SplitBlock {
        block_id: BlockId,
        offset: usize,
    },
    MergeBlocks {
        first: BlockId,
        second: BlockId,
    },
    InsertBlock {
        at_index: usize,
    },
    RestoreBlock {
        at_index: usize,
        block: Block,
    },
    DeleteBlock {
        block_id: BlockId,
    },
    MoveBlock {
        block_id: BlockId,
        to_index: usize,
    },
    ReplaceRuns {
        block_id: BlockId,
        runs: Vec<Run>,
    },
    ReplaceParagraph {
        block_id: BlockId,
        paragraph: Paragraph,
    },
    FormatRun {
        block_id: BlockId,
        start: usize,
        end: usize,
        style_patch: RunStylePatch,
    },
    ClearFormatting {
        block_id: BlockId,
        start: usize,
        end: usize,
    },
    FormatParagraph {
        block_id: BlockId,
        op: ParagraphFormatOp,
    },
}

#[derive(Debug, Clone, Default)]
pub struct RunStylePatch {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub strikethrough: Option<bool>,
    pub superscript: Option<bool>,
    pub subscript: Option<bool>,
    pub font_family: Option<String>,
    pub font_size: Option<f32>,
    pub color: Option<Color>,
    pub background: Option<Color>,
}

#[derive(Debug, Clone)]
pub enum ParagraphFormatOp {
    Alignment(ParagraphAlignment),
    HeadingLevel(Option<u8>),
    ListType(Option<ListType>),
    IndentDelta(f32),
    LineSpacing(f32),
    ParagraphSpacing { before: f32, after: f32 },
    BlockQuoteToggle,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Shortcut {
    Bold,
    Italic,
    Underline,
    Strikethrough,
    Superscript,
    Subscript,
    ClearFormatting,
    AlignLeft,
    AlignCenter,
    AlignRight,
    AlignJustify,
    Save,
    Open,
    New,
    Undo,
    Redo,
    Copy,
    Cut,
    Paste,
    PastePlain,
    Find,
    Replace,
    Print,
    CommandPalette,
    ZoomIn,
    ZoomOut,
    ZoomReset,
    SelectAll,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct SelectionRange {
    pub start: usize,
    pub end: usize,
}

impl SelectionRange {
    pub fn new(start: usize, end: usize) -> Self {
        if start <= end {
            Self { start, end }
        } else {
            Self {
                start: end,
                end: start,
            }
        }
    }

    pub fn is_empty(self) -> bool {
        self.start >= self.end
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SelectionToggleState {
    Off,
    On,
    Mixed,
}

impl SelectionToggleState {
    pub fn toggled_target(self) -> bool {
        !matches!(self, Self::On)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SelectionToggleSnapshot {
    pub bold: SelectionToggleState,
    pub italic: SelectionToggleState,
    pub underline: SelectionToggleState,
    pub strikethrough: SelectionToggleState,
    pub superscript: SelectionToggleState,
    pub subscript: SelectionToggleState,
}

pub fn shortcut_from_vk(ctrl: bool, shift: bool, vk: u32) -> Option<Shortcut> {
    if !ctrl {
        return None;
    }

    match (shift, vk) {
        (false, 0x42) => Some(Shortcut::Bold),
        (false, 0x49) => Some(Shortcut::Italic),
        (false, 0x55) => Some(Shortcut::Underline),
        (true, 0x58) => Some(Shortcut::Strikethrough),
        (true, 0xBB) => Some(Shortcut::Superscript),
        (false, 0xBB) => Some(Shortcut::Subscript),
        (false, 0xDC) => Some(Shortcut::ClearFormatting),
        (false, 0x4C) => Some(Shortcut::AlignLeft),
        (false, 0x45) => Some(Shortcut::AlignCenter),
        (false, 0x52) => Some(Shortcut::AlignRight),
        (false, 0x4A) => Some(Shortcut::AlignJustify),
        (false, 0x53) => Some(Shortcut::Save),
        (false, 0x4F) => Some(Shortcut::Open),
        (false, 0x4E) => Some(Shortcut::New),
        (false, 0x5A) => Some(Shortcut::Undo),
        (false, 0x59) => Some(Shortcut::Redo),
        (true, 0x5A) => Some(Shortcut::Redo),
        (false, 0x43) => Some(Shortcut::Copy),
        (false, 0x58) => Some(Shortcut::Cut),
        (false, 0x56) => Some(Shortcut::Paste),
        (true, 0x56) => Some(Shortcut::PastePlain),
        (false, 0x46) => Some(Shortcut::Find),
        (false, 0x48) => Some(Shortcut::Replace),
        (false, 0x50) => Some(Shortcut::Print),
        (true, 0x50) => Some(Shortcut::CommandPalette),
        (false, 0x6B) => Some(Shortcut::ZoomIn),
        (false, 0x6D) | (false, 0xBD) => Some(Shortcut::ZoomOut),
        (false, 0x30) => Some(Shortcut::ZoomReset),
        (false, 0x41) => Some(Shortcut::SelectAll),
        _ => None,
    }
}

pub fn insert_text(block_id: BlockId, offset: usize, text: impl Into<String>) -> EditCommand {
    EditCommand::InsertText {
        block_id,
        offset,
        text: text.into(),
    }
}

pub fn backspace(block_id: BlockId, cursor_offset: usize) -> Option<EditCommand> {
    if cursor_offset == 0 {
        return None;
    }
    Some(EditCommand::DeleteText {
        block_id,
        start: cursor_offset.saturating_sub(1),
        end: cursor_offset,
    })
}

pub fn delete_forward(block_id: BlockId, cursor_offset: usize) -> EditCommand {
    EditCommand::DeleteText {
        block_id,
        start: cursor_offset,
        end: cursor_offset + 1,
    }
}

pub fn split_paragraph(block_id: BlockId, offset: usize) -> EditCommand {
    EditCommand::SplitBlock { block_id, offset }
}

pub fn tab(block_id: BlockId, offset: usize, as_indent: bool) -> EditCommand {
    if as_indent {
        increase_indent(block_id)
    } else {
        EditCommand::InsertText {
            block_id,
            offset,
            text: "\t".to_string(),
        }
    }
}

pub fn outdent(block_id: BlockId) -> EditCommand {
    decrease_indent(block_id)
}

pub fn format_selection(
    block_id: BlockId,
    start: usize,
    end: usize,
    style_patch: RunStylePatch,
) -> EditCommand {
    EditCommand::FormatRun {
        block_id,
        start,
        end,
        style_patch,
    }
}

pub fn detect_selection_toggles(runs: &[Run], range: SelectionRange) -> SelectionToggleSnapshot {
    SelectionToggleSnapshot {
        bold: detect_toggle_state(runs, range, |style| style.bold),
        italic: detect_toggle_state(runs, range, |style| style.italic),
        underline: detect_toggle_state(runs, range, |style| style.underline),
        strikethrough: detect_toggle_state(runs, range, |style| style.strikethrough),
        superscript: detect_toggle_state(runs, range, |style| style.superscript),
        subscript: detect_toggle_state(runs, range, |style| style.subscript),
    }
}

pub fn detect_toggle_state<F>(
    runs: &[Run],
    range: SelectionRange,
    selector: F,
) -> SelectionToggleState
where
    F: Fn(&RunStyle) -> bool,
{
    let selected = selected_runs(runs, range);
    if selected.is_empty() {
        return SelectionToggleState::Off;
    }

    let on = selected.iter().filter(|run| selector(&run.style)).count();
    if on == 0 {
        SelectionToggleState::Off
    } else if on == selected.len() {
        SelectionToggleState::On
    } else {
        SelectionToggleState::Mixed
    }
}

pub fn toggle_bold(block_id: BlockId, range: SelectionRange, runs: &[Run]) -> EditCommand {
    let next = detect_toggle_state(runs, range, |style| style.bold).toggled_target();
    format_selection(
        block_id,
        range.start,
        range.end,
        RunStylePatch {
            bold: Some(next),
            ..RunStylePatch::default()
        },
    )
}

pub fn toggle_italic(block_id: BlockId, range: SelectionRange, runs: &[Run]) -> EditCommand {
    let next = detect_toggle_state(runs, range, |style| style.italic).toggled_target();
    format_selection(
        block_id,
        range.start,
        range.end,
        RunStylePatch {
            italic: Some(next),
            ..RunStylePatch::default()
        },
    )
}

pub fn toggle_underline(block_id: BlockId, range: SelectionRange, runs: &[Run]) -> EditCommand {
    let next = detect_toggle_state(runs, range, |style| style.underline).toggled_target();
    format_selection(
        block_id,
        range.start,
        range.end,
        RunStylePatch {
            underline: Some(next),
            ..RunStylePatch::default()
        },
    )
}

pub fn toggle_strikethrough(block_id: BlockId, range: SelectionRange, runs: &[Run]) -> EditCommand {
    let next = detect_toggle_state(runs, range, |style| style.strikethrough).toggled_target();
    format_selection(
        block_id,
        range.start,
        range.end,
        RunStylePatch {
            strikethrough: Some(next),
            ..RunStylePatch::default()
        },
    )
}

pub fn toggle_superscript(block_id: BlockId, range: SelectionRange, runs: &[Run]) -> EditCommand {
    let next = detect_toggle_state(runs, range, |style| style.superscript).toggled_target();
    format_selection(
        block_id,
        range.start,
        range.end,
        RunStylePatch {
            superscript: Some(next),
            subscript: if next { Some(false) } else { None },
            ..RunStylePatch::default()
        },
    )
}

pub fn toggle_subscript(block_id: BlockId, range: SelectionRange, runs: &[Run]) -> EditCommand {
    let next = detect_toggle_state(runs, range, |style| style.subscript).toggled_target();
    format_selection(
        block_id,
        range.start,
        range.end,
        RunStylePatch {
            subscript: Some(next),
            superscript: if next { Some(false) } else { None },
            ..RunStylePatch::default()
        },
    )
}

pub fn set_font_family(
    block_id: BlockId,
    range: SelectionRange,
    font_family: impl Into<String>,
) -> EditCommand {
    format_selection(
        block_id,
        range.start,
        range.end,
        RunStylePatch {
            font_family: Some(font_family.into()),
            ..RunStylePatch::default()
        },
    )
}

pub fn set_font_size(block_id: BlockId, range: SelectionRange, font_size: f32) -> EditCommand {
    format_selection(
        block_id,
        range.start,
        range.end,
        RunStylePatch {
            font_size: Some(font_size.max(1.0)),
            ..RunStylePatch::default()
        },
    )
}

pub fn set_text_color(block_id: BlockId, range: SelectionRange, color: Color) -> EditCommand {
    format_selection(
        block_id,
        range.start,
        range.end,
        RunStylePatch {
            color: Some(color),
            ..RunStylePatch::default()
        },
    )
}

pub fn set_highlight_color(block_id: BlockId, range: SelectionRange, color: Color) -> EditCommand {
    format_selection(
        block_id,
        range.start,
        range.end,
        RunStylePatch {
            background: Some(color),
            ..RunStylePatch::default()
        },
    )
}

pub fn clear_selection_formatting(block_id: BlockId, range: SelectionRange) -> EditCommand {
    EditCommand::ClearFormatting {
        block_id,
        start: range.start,
        end: range.end,
    }
}

pub fn apply_or_toggle_bold(
    block_id: BlockId,
    selection: Option<SelectionRange>,
    runs: &[Run],
    pending: &mut RunStyle,
) -> Option<EditCommand> {
    apply_or_pending_toggle(selection, pending, |range| toggle_bold(block_id, range, runs), |style| {
        style.bold = !style.bold;
    })
}

pub fn apply_or_toggle_italic(
    block_id: BlockId,
    selection: Option<SelectionRange>,
    runs: &[Run],
    pending: &mut RunStyle,
) -> Option<EditCommand> {
    apply_or_pending_toggle(selection, pending, |range| toggle_italic(block_id, range, runs), |style| {
        style.italic = !style.italic;
    })
}

pub fn apply_or_toggle_underline(
    block_id: BlockId,
    selection: Option<SelectionRange>,
    runs: &[Run],
    pending: &mut RunStyle,
) -> Option<EditCommand> {
    apply_or_pending_toggle(selection, pending, |range| toggle_underline(block_id, range, runs), |style| {
        style.underline = !style.underline;
    })
}

pub fn apply_or_toggle_strikethrough(
    block_id: BlockId,
    selection: Option<SelectionRange>,
    runs: &[Run],
    pending: &mut RunStyle,
) -> Option<EditCommand> {
    apply_or_pending_toggle(selection, pending, |range| toggle_strikethrough(block_id, range, runs), |style| {
        style.strikethrough = !style.strikethrough;
    })
}

pub fn apply_or_toggle_superscript(
    block_id: BlockId,
    selection: Option<SelectionRange>,
    runs: &[Run],
    pending: &mut RunStyle,
) -> Option<EditCommand> {
    apply_or_pending_toggle(selection, pending, |range| toggle_superscript(block_id, range, runs), |style| {
        style.superscript = !style.superscript;
        if style.superscript {
            style.subscript = false;
        }
    })
}

pub fn apply_or_toggle_subscript(
    block_id: BlockId,
    selection: Option<SelectionRange>,
    runs: &[Run],
    pending: &mut RunStyle,
) -> Option<EditCommand> {
    apply_or_pending_toggle(selection, pending, |range| toggle_subscript(block_id, range, runs), |style| {
        style.subscript = !style.subscript;
        if style.subscript {
            style.superscript = false;
        }
    })
}

pub fn apply_or_set_font_family(
    block_id: BlockId,
    selection: Option<SelectionRange>,
    font_family: impl Into<String>,
    pending: &mut RunStyle,
) -> Option<EditCommand> {
    let family = font_family.into();
    apply_or_pending_patch(
        block_id,
        selection,
        RunStylePatch {
            font_family: Some(family.clone()),
            ..RunStylePatch::default()
        },
        pending,
        move |style| style.font_family = Some(family),
    )
}

pub fn apply_or_set_font_size(
    block_id: BlockId,
    selection: Option<SelectionRange>,
    font_size: f32,
    pending: &mut RunStyle,
) -> Option<EditCommand> {
    let size = font_size.max(1.0);
    apply_or_pending_patch(
        block_id,
        selection,
        RunStylePatch {
            font_size: Some(size),
            ..RunStylePatch::default()
        },
        pending,
        move |style| style.font_size = Some(size),
    )
}

pub fn apply_or_set_text_color(
    block_id: BlockId,
    selection: Option<SelectionRange>,
    color: Color,
    pending: &mut RunStyle,
) -> Option<EditCommand> {
    apply_or_pending_patch(
        block_id,
        selection,
        RunStylePatch {
            color: Some(color),
            ..RunStylePatch::default()
        },
        pending,
        move |style| style.color = Some(color),
    )
}

pub fn apply_or_set_highlight_color(
    block_id: BlockId,
    selection: Option<SelectionRange>,
    color: Color,
    pending: &mut RunStyle,
) -> Option<EditCommand> {
    apply_or_pending_patch(
        block_id,
        selection,
        RunStylePatch {
            background: Some(color),
            ..RunStylePatch::default()
        },
        pending,
        move |style| style.background = Some(color),
    )
}

pub fn apply_or_clear_formatting(
    block_id: BlockId,
    selection: Option<SelectionRange>,
    pending: &mut RunStyle,
) -> Option<EditCommand> {
    if let Some(range) = normalize_selection(selection) {
        return Some(clear_selection_formatting(block_id, range));
    }

    *pending = RunStyle::default();
    None
}

pub fn set_alignment(block_id: BlockId, alignment: ParagraphAlignment) -> EditCommand {
    EditCommand::FormatParagraph {
        block_id,
        op: ParagraphFormatOp::Alignment(alignment),
    }
}

pub fn align_left(block_id: BlockId) -> EditCommand {
    set_alignment(block_id, ParagraphAlignment::Left)
}

pub fn align_center(block_id: BlockId) -> EditCommand {
    set_alignment(block_id, ParagraphAlignment::Center)
}

pub fn align_right(block_id: BlockId) -> EditCommand {
    set_alignment(block_id, ParagraphAlignment::Right)
}

pub fn align_justify(block_id: BlockId) -> EditCommand {
    set_alignment(block_id, ParagraphAlignment::Justify)
}

pub fn set_heading_level(block_id: BlockId, level: Option<u8>) -> EditCommand {
    EditCommand::FormatParagraph {
        block_id,
        op: ParagraphFormatOp::HeadingLevel(level.map(|v| v.clamp(1, 6))),
    }
}

pub fn set_list_type(block_id: BlockId, list_type: Option<ListType>) -> EditCommand {
    EditCommand::FormatParagraph {
        block_id,
        op: ParagraphFormatOp::ListType(list_type),
    }
}

pub fn increase_indent(block_id: BlockId) -> EditCommand {
    EditCommand::FormatParagraph {
        block_id,
        op: ParagraphFormatOp::IndentDelta(24.0),
    }
}

pub fn decrease_indent(block_id: BlockId) -> EditCommand {
    EditCommand::FormatParagraph {
        block_id,
        op: ParagraphFormatOp::IndentDelta(-24.0),
    }
}

pub fn set_line_spacing(block_id: BlockId, line_spacing: f32) -> EditCommand {
    EditCommand::FormatParagraph {
        block_id,
        op: ParagraphFormatOp::LineSpacing(line_spacing.max(0.5)),
    }
}

pub fn set_paragraph_spacing(block_id: BlockId, before: f32, after: f32) -> EditCommand {
    EditCommand::FormatParagraph {
        block_id,
        op: ParagraphFormatOp::ParagraphSpacing {
            before: before.max(0.0),
            after: after.max(0.0),
        },
    }
}

pub fn toggle_block_quote(block_id: BlockId) -> EditCommand {
    EditCommand::FormatParagraph {
        block_id,
        op: ParagraphFormatOp::BlockQuoteToggle,
    }
}

pub fn patch_toggle_bold() -> RunStylePatch {
    RunStylePatch {
        bold: Some(true),
        ..RunStylePatch::default()
    }
}

pub fn patch_toggle_italic() -> RunStylePatch {
    RunStylePatch {
        italic: Some(true),
        ..RunStylePatch::default()
    }
}

pub fn patch_toggle_underline() -> RunStylePatch {
    RunStylePatch {
        underline: Some(true),
        ..RunStylePatch::default()
    }
}

fn selected_runs<'a>(runs: &'a [Run], range: SelectionRange) -> Vec<&'a Run> {
    if runs.is_empty() || range.is_empty() {
        return Vec::new();
    }

    let mut selected = Vec::new();
    let mut cursor = 0usize;

    for run in runs {
        let next = cursor + run.text.len();
        if next > range.start && cursor < range.end {
            selected.push(run);
        }
        cursor = next;
    }

    selected
}

fn normalize_selection(selection: Option<SelectionRange>) -> Option<SelectionRange> {
    let range = selection?;
    if range.is_empty() {
        None
    } else {
        Some(SelectionRange::new(range.start, range.end))
    }
}

fn apply_or_pending_toggle<F, P>(
    selection: Option<SelectionRange>,
    pending: &mut RunStyle,
    command_builder: F,
    pending_mutator: P,
) -> Option<EditCommand>
where
    F: FnOnce(SelectionRange) -> EditCommand,
    P: FnOnce(&mut RunStyle),
{
    if let Some(range) = normalize_selection(selection) {
        return Some(command_builder(range));
    }

    pending_mutator(pending);
    None
}

fn apply_or_pending_patch<P>(
    block_id: BlockId,
    selection: Option<SelectionRange>,
    patch: RunStylePatch,
    pending: &mut RunStyle,
    pending_mutator: P,
) -> Option<EditCommand>
where
    P: FnOnce(&mut RunStyle),
{
    if let Some(range) = normalize_selection(selection) {
        return Some(format_selection(block_id, range.start, range.end, patch));
    }

    pending_mutator(pending);
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    fn run_with_style(text: &str, style: RunStyle) -> Run {
        Run {
            text: text.to_string(),
            style,
        }
    }

    #[test]
    fn maps_required_shortcuts() {
        assert_eq!(shortcut_from_vk(true, false, 0x42), Some(Shortcut::Bold));
        assert_eq!(shortcut_from_vk(true, false, 0x49), Some(Shortcut::Italic));
        assert_eq!(shortcut_from_vk(true, false, 0x55), Some(Shortcut::Underline));
        assert_eq!(shortcut_from_vk(true, true, 0x58), Some(Shortcut::Strikethrough));
        assert_eq!(shortcut_from_vk(true, true, 0xBB), Some(Shortcut::Superscript));
        assert_eq!(shortcut_from_vk(true, false, 0xBB), Some(Shortcut::Subscript));
        assert_eq!(shortcut_from_vk(true, false, 0xDC), Some(Shortcut::ClearFormatting));
        assert_eq!(shortcut_from_vk(true, false, 0x53), Some(Shortcut::Save));
        assert_eq!(shortcut_from_vk(true, false, 0x4F), Some(Shortcut::Open));
        assert_eq!(shortcut_from_vk(true, false, 0x4E), Some(Shortcut::New));
        assert_eq!(shortcut_from_vk(true, false, 0x5A), Some(Shortcut::Undo));
        assert_eq!(shortcut_from_vk(true, false, 0x59), Some(Shortcut::Redo));
        assert_eq!(shortcut_from_vk(true, true, 0x5A), Some(Shortcut::Redo));
        assert_eq!(shortcut_from_vk(true, false, 0x43), Some(Shortcut::Copy));
        assert_eq!(shortcut_from_vk(true, false, 0x58), Some(Shortcut::Cut));
        assert_eq!(shortcut_from_vk(true, false, 0x56), Some(Shortcut::Paste));
        assert_eq!(shortcut_from_vk(true, true, 0x56), Some(Shortcut::PastePlain));
        assert_eq!(shortcut_from_vk(true, false, 0x46), Some(Shortcut::Find));
        assert_eq!(shortcut_from_vk(true, false, 0x48), Some(Shortcut::Replace));
        assert_eq!(shortcut_from_vk(true, false, 0x50), Some(Shortcut::Print));
        assert_eq!(shortcut_from_vk(true, true, 0x50), Some(Shortcut::CommandPalette));
        assert_eq!(shortcut_from_vk(true, false, 0x6B), Some(Shortcut::ZoomIn));
        assert_eq!(shortcut_from_vk(true, false, 0xBD), Some(Shortcut::ZoomOut));
        assert_eq!(shortcut_from_vk(true, false, 0x30), Some(Shortcut::ZoomReset));
        assert_eq!(shortcut_from_vk(true, false, 0x30), Some(Shortcut::ZoomReset));
        assert_eq!(shortcut_from_vk(true, false, 0x41), Some(Shortcut::SelectAll));
    }

    #[test]
    fn detects_mixed_toggle_state() {
        let runs = vec![
            run_with_style(
                "ab",
                RunStyle {
                    bold: true,
                    ..RunStyle::default()
                },
            ),
            run_with_style("cd", RunStyle::default()),
        ];

        let state = detect_toggle_state(&runs, SelectionRange::new(0, 4), |style| style.bold);
        assert_eq!(state, SelectionToggleState::Mixed);
    }

    #[test]
    fn toggle_bold_uses_detected_state() {
        let runs = vec![run_with_style(
            "abcd",
            RunStyle {
                bold: true,
                ..RunStyle::default()
            },
        )];

        let cmd = toggle_bold(BlockId(9), SelectionRange::new(0, 4), &runs);
        match cmd {
            EditCommand::FormatRun {
                block_id,
                start,
                end,
                style_patch,
            } => {
                assert_eq!(block_id, BlockId(9));
                assert_eq!((start, end), (0, 4));
                assert_eq!(style_patch.bold, Some(false));
            }
            _ => panic!("expected format run command"),
        }
    }

    #[test]
    fn pending_font_family_updates_when_no_selection() {
        let mut pending = RunStyle::default();
        let cmd = apply_or_set_font_family(BlockId(1), None, "Calibri", &mut pending);

        assert!(cmd.is_none());
        assert_eq!(pending.font_family.as_deref(), Some("Calibri"));
    }

    #[test]
    fn pending_subscript_clears_superscript() {
        let mut pending = RunStyle {
            superscript: true,
            ..RunStyle::default()
        };

        let cmd = apply_or_toggle_subscript(BlockId(1), None, &[], &mut pending);
        assert!(cmd.is_none());
        assert!(pending.subscript);
        assert!(!pending.superscript);
    }

    #[test]
    fn paragraph_format_builders_emit_expected_commands() {
        match set_heading_level(BlockId(1), Some(9)) {
            EditCommand::FormatParagraph {
                op: ParagraphFormatOp::HeadingLevel(level),
                ..
            } => assert_eq!(level, Some(6)),
            _ => panic!("expected heading command"),
        }

        match decrease_indent(BlockId(1)) {
            EditCommand::FormatParagraph {
                op: ParagraphFormatOp::IndentDelta(delta),
                ..
            } => assert!(delta < 0.0),
            _ => panic!("expected indent command"),
        }

        match set_alignment(BlockId(1), ParagraphAlignment::Justify) {
            EditCommand::FormatParagraph {
                op: ParagraphFormatOp::Alignment(ParagraphAlignment::Justify),
                ..
            } => {}
            _ => panic!("expected alignment command"),
        }
    }
}
