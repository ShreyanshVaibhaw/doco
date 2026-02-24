use crate::document::model::{
    Block,
    BlockId,
    ListType,
    Paragraph,
    ParagraphAlignment,
    Run,
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
    pub color: Option<crate::ui::Color>,
    pub background: Option<crate::ui::Color>,
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

pub fn shortcut_from_vk(ctrl: bool, shift: bool, vk: u32) -> Option<Shortcut> {
    if !ctrl {
        return None;
    }

    match (shift, vk) {
        (false, 0x42) => Some(Shortcut::Bold),
        (false, 0x49) => Some(Shortcut::Italic),
        (false, 0x55) => Some(Shortcut::Underline),
        (true, 0x58) => Some(Shortcut::Strikethrough),
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
        (false, 0x6B) | (true, 0xBB) => Some(Shortcut::ZoomIn),
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
        EditCommand::FormatParagraph {
            block_id,
            op: ParagraphFormatOp::IndentDelta(24.0),
        }
    } else {
        EditCommand::InsertText {
            block_id,
            offset,
            text: "\t".to_string(),
        }
    }
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn maps_required_shortcuts() {
        assert_eq!(shortcut_from_vk(true, false, 0x42), Some(Shortcut::Bold));
        assert_eq!(shortcut_from_vk(true, false, 0x49), Some(Shortcut::Italic));
        assert_eq!(shortcut_from_vk(true, false, 0x55), Some(Shortcut::Underline));
        assert_eq!(shortcut_from_vk(true, false, 0x53), Some(Shortcut::Save));
        assert_eq!(shortcut_from_vk(true, false, 0x4F), Some(Shortcut::Open));
        assert_eq!(shortcut_from_vk(true, false, 0x4E), Some(Shortcut::New));
        assert_eq!(shortcut_from_vk(true, false, 0x5A), Some(Shortcut::Undo));
        assert_eq!(shortcut_from_vk(true, false, 0x59), Some(Shortcut::Redo));
        assert_eq!(shortcut_from_vk(true, true, 0x5A), Some(Shortcut::Redo));
        assert_eq!(shortcut_from_vk(true, false, 0x43), Some(Shortcut::Copy));
        assert_eq!(shortcut_from_vk(true, false, 0x58), Some(Shortcut::Cut));
        assert_eq!(shortcut_from_vk(true, false, 0x56), Some(Shortcut::Paste));
        assert_eq!(shortcut_from_vk(true, false, 0x46), Some(Shortcut::Find));
        assert_eq!(shortcut_from_vk(true, false, 0x48), Some(Shortcut::Replace));
        assert_eq!(shortcut_from_vk(true, false, 0x50), Some(Shortcut::Print));
        assert_eq!(shortcut_from_vk(true, true, 0x50), Some(Shortcut::CommandPalette));
        assert_eq!(shortcut_from_vk(true, true, 0xBB), Some(Shortcut::ZoomIn));
        assert_eq!(shortcut_from_vk(true, false, 0xBD), Some(Shortcut::ZoomOut));
        assert_eq!(shortcut_from_vk(true, false, 0x30), Some(Shortcut::ZoomReset));
    }
}
