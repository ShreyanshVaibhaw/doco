use std::time::Instant;

use crate::{
    document::model::{Block, DocumentModel, Paragraph, ParagraphAlignment, Run, RunStyle},
    editor::{
        commands::{EditCommand, ParagraphFormatOp, RunStylePatch, Shortcut},
        cursor::CursorState,
        undo::{UndoEntry, UndoStack},
    },
};

pub mod clipboard;
pub mod commands;
pub mod cursor;
pub mod image_ops;
pub mod search;
pub mod table;
pub mod undo;

#[derive(Default)]
pub struct EditEngine {
    pub cursor: CursorState,
    pub undo: UndoStack,
    pub pending_format: RunStyle,
}

impl EditEngine {
    pub fn apply_command(&mut self, doc: &mut DocumentModel, command: EditCommand) {
        if let Some(inverse) = apply_to_document(doc, &command) {
            let bytes = estimate_command_size(&command);
            self.undo.push(UndoEntry {
                command,
                inverse,
                bytes,
                timestamp: Instant::now(),
            });
            doc.dirty = true;
        }
    }

    pub fn undo(&mut self, doc: &mut DocumentModel) {
        if let Some(entry) = self.undo.pop_undo() {
            let _ = apply_to_document(doc, &entry.inverse);
            doc.dirty = true;
        }
    }

    pub fn redo(&mut self, doc: &mut DocumentModel) {
        if let Some(entry) = self.undo.pop_redo() {
            let _ = apply_to_document(doc, &entry.command);
            doc.dirty = true;
        }
    }

    pub fn handle_shortcut(&mut self, shortcut: Shortcut) {
        match shortcut {
            Shortcut::Bold => self.pending_format.bold = !self.pending_format.bold,
            Shortcut::Italic => self.pending_format.italic = !self.pending_format.italic,
            Shortcut::Underline => self.pending_format.underline = !self.pending_format.underline,
            Shortcut::Strikethrough => {
                self.pending_format.strikethrough = !self.pending_format.strikethrough
            }
            Shortcut::Superscript => {
                self.pending_format.superscript = !self.pending_format.superscript;
                if self.pending_format.superscript {
                    self.pending_format.subscript = false;
                }
            }
            Shortcut::Subscript => {
                self.pending_format.subscript = !self.pending_format.subscript;
                if self.pending_format.subscript {
                    self.pending_format.superscript = false;
                }
            }
            Shortcut::ClearFormatting => self.pending_format = RunStyle::default(),
            _ => {}
        }
    }
}

fn apply_to_document(doc: &mut DocumentModel, command: &EditCommand) -> Option<EditCommand> {
    match command {
        EditCommand::InsertText {
            block_id,
            offset,
            text,
        } => {
            let (run, _) = find_or_create_run(doc, *block_id)?;
            let off = (*offset).min(run.text.len());
            run.text.insert_str(off, text);
            Some(EditCommand::DeleteText {
                block_id: *block_id,
                start: off,
                end: off + text.len(),
            })
        }
        EditCommand::DeleteText {
            block_id,
            start,
            end,
        } => {
            let (run, _) = find_or_create_run(doc, *block_id)?;
            if *start >= *end || *start >= run.text.len() {
                return None;
            }
            let end = (*end).min(run.text.len());
            let removed = run.text[*start..end].to_string();
            run.text.replace_range(*start..end, "");
            Some(EditCommand::InsertText {
                block_id: *block_id,
                offset: *start,
                text: removed,
            })
        }
        EditCommand::ReplaceText {
            block_id,
            start,
            end,
            text,
        } => {
            let (run, _) = find_or_create_run(doc, *block_id)?;
            let s = (*start).min(run.text.len());
            let e = (*end).min(run.text.len()).max(s);
            let replaced = run.text[s..e].to_string();
            run.text.replace_range(s..e, text);
            Some(EditCommand::ReplaceText {
                block_id: *block_id,
                start: s,
                end: s + text.len(),
                text: replaced,
            })
        }
        EditCommand::SplitBlock { block_id, offset } => {
            let paragraph = find_paragraph_mut(doc, *block_id)?.clone();
            let paragraph_idx = find_block_index_by_id(doc, *block_id)?;
            let base_style = paragraph
                .runs
                .first()
                .map(|r| r.style.clone())
                .unwrap_or_default();

            let mut text = paragraph
                .runs
                .iter()
                .map(|r| r.text.as_str())
                .collect::<String>();
            let cut = (*offset).min(text.len());
            let mut boundary = cut;
            while boundary > 0 && !text.is_char_boundary(boundary) {
                boundary -= 1;
            }
            let right = text.split_off(boundary);

            let mut left_paragraph = paragraph.clone();
            left_paragraph.runs = vec![Run {
                text,
                style: base_style.clone(),
            }];

            let new_id = next_block_id(doc);
            let mut right_paragraph = paragraph;
            right_paragraph.id = new_id;
            right_paragraph.runs = vec![Run {
                text: right,
                style: base_style,
            }];

            doc.content[paragraph_idx] = Block::Paragraph(left_paragraph.clone());
            doc.content
                .insert(paragraph_idx + 1, Block::Paragraph(right_paragraph.clone()));

            Some(EditCommand::MergeBlocks {
                first: left_paragraph.id,
                second: right_paragraph.id,
            })
        }
        EditCommand::MergeBlocks { first, second } => {
            let first_idx = find_block_index_by_id(doc, *first)?;
            let second_idx = find_block_index_by_id(doc, *second)?;
            if first_idx == second_idx {
                return None;
            }

            let (left_idx, right_idx) = if first_idx < second_idx {
                (first_idx, second_idx)
            } else {
                (second_idx, first_idx)
            };

            let split_offset = paragraph_text_len(match &doc.content[left_idx] {
                Block::Paragraph(p) => p,
                _ => return None,
            });

            let right_paragraph = match doc.content.remove(right_idx) {
                Block::Paragraph(p) => p,
                _ => return None,
            };

            let left_paragraph = match &mut doc.content[left_idx] {
                Block::Paragraph(p) => p,
                _ => return None,
            };
            left_paragraph.runs.extend(right_paragraph.runs);
            merge_adjacent_runs(&mut left_paragraph.runs);

            Some(EditCommand::SplitBlock {
                block_id: left_paragraph.id,
                offset: split_offset,
            })
        }
        EditCommand::InsertBlock { at_index } => {
            let id = next_block_id(doc);
            let paragraph = make_empty_paragraph(id);
            let idx = (*at_index).min(doc.content.len());
            doc.content.insert(idx, Block::Paragraph(paragraph));
            Some(EditCommand::DeleteBlock { block_id: id })
        }
        EditCommand::RestoreBlock { at_index, block } => {
            let idx = (*at_index).min(doc.content.len());
            let block_id = block_id_of(block)?;
            doc.content.insert(idx, block.clone());
            Some(EditCommand::DeleteBlock { block_id })
        }
        EditCommand::DeleteBlock { block_id } => {
            let idx = find_block_index_by_id(doc, *block_id)?;
            let removed = doc.content.remove(idx);
            Some(EditCommand::RestoreBlock {
                at_index: idx,
                block: removed,
            })
        }
        EditCommand::MoveBlock { block_id, to_index } => {
            let from_idx = find_block_index_by_id(doc, *block_id)?;
            let block = doc.content.remove(from_idx);
            let to_idx = (*to_index).min(doc.content.len());
            doc.content.insert(to_idx, block);
            Some(EditCommand::MoveBlock {
                block_id: *block_id,
                to_index: from_idx,
            })
        }
        EditCommand::ReplaceRuns { block_id, runs } => {
            let paragraph = find_paragraph_mut(doc, *block_id)?;
            let old = paragraph.runs.clone();
            paragraph.runs = runs.clone();
            Some(EditCommand::ReplaceRuns {
                block_id: *block_id,
                runs: old,
            })
        }
        EditCommand::ReplaceParagraph {
            block_id,
            paragraph,
        } => {
            let current = find_paragraph_mut(doc, *block_id)?;
            let old = current.clone();
            *current = paragraph.clone();
            Some(EditCommand::ReplaceParagraph {
                block_id: *block_id,
                paragraph: old,
            })
        }
        EditCommand::FormatRun {
            block_id,
            start,
            end,
            style_patch,
        } => {
            let paragraph = find_paragraph_mut(doc, *block_id)?;
            let old_runs = paragraph.runs.clone();
            apply_style_patch(paragraph, *start, *end, style_patch);
            Some(EditCommand::ReplaceRuns {
                block_id: *block_id,
                runs: old_runs,
            })
        }
        EditCommand::ClearFormatting {
            block_id,
            start,
            end,
        } => {
            let paragraph = find_paragraph_mut(doc, *block_id)?;
            let old_runs = paragraph.runs.clone();
            apply_style_patch(
                paragraph,
                *start,
                *end,
                &RunStylePatch {
                    bold: Some(false),
                    italic: Some(false),
                    underline: Some(false),
                    strikethrough: Some(false),
                    superscript: Some(false),
                    subscript: Some(false),
                    font_family: Some("Segoe UI".to_string()),
                    font_size: Some(12.0),
                    color: None,
                    background: None,
                },
            );
            Some(EditCommand::ReplaceRuns {
                block_id: *block_id,
                runs: old_runs,
            })
        }
        EditCommand::FormatParagraph { block_id, op } => {
            let paragraph = find_paragraph_mut(doc, *block_id)?;
            let old = paragraph.clone();

            match op {
                ParagraphFormatOp::Alignment(a) => paragraph.alignment = a.clone(),
                ParagraphFormatOp::HeadingLevel(level) => {
                    paragraph.style_id = level.map(|l| format!("Heading{l}"));
                }
                ParagraphFormatOp::ListType(list_type) => {
                    paragraph.style_id = match list_type {
                        Some(crate::document::model::ListType::Bullet) => {
                            Some("ListBullet".to_string())
                        }
                        Some(crate::document::model::ListType::Numbered) => {
                            Some("ListNumber".to_string())
                        }
                        Some(crate::document::model::ListType::Checkbox) => {
                            Some("ListCheckbox".to_string())
                        }
                        None => None,
                    }
                }
                ParagraphFormatOp::IndentDelta(delta) => {
                    paragraph.indent.left = (paragraph.indent.left + *delta).max(0.0);
                }
                ParagraphFormatOp::LineSpacing(line) => paragraph.spacing.line = *line,
                ParagraphFormatOp::ParagraphSpacing { before, after } => {
                    paragraph.spacing.before = *before;
                    paragraph.spacing.after = *after;
                }
                ParagraphFormatOp::BlockQuoteToggle => {
                    paragraph.indent.left = if paragraph.indent.left < 18.0 {
                        24.0
                    } else {
                        0.0
                    };
                }
            }

            Some(EditCommand::ReplaceParagraph {
                block_id: *block_id,
                paragraph: old,
            })
        }
        _ => None,
    }
}

fn apply_style_patch(paragraph: &mut Paragraph, start: usize, end: usize, patch: &RunStylePatch) {
    if paragraph.runs.is_empty() {
        paragraph.runs.push(Run::default());
    }

    let text_len = paragraph.runs.iter().map(|r| r.text.len()).sum::<usize>();
    let s = start.min(text_len);
    let e = end.min(text_len).max(s);

    let start_idx = split_runs_at(&mut paragraph.runs, s);
    let end_idx = split_runs_at(&mut paragraph.runs, e);

    for run in paragraph.runs.iter_mut().take(end_idx).skip(start_idx) {
        if let Some(v) = patch.bold {
            run.style.bold = v;
        }
        if let Some(v) = patch.italic {
            run.style.italic = v;
        }
        if let Some(v) = patch.underline {
            run.style.underline = v;
        }
        if let Some(v) = patch.strikethrough {
            run.style.strikethrough = v;
        }
        if let Some(v) = patch.superscript {
            run.style.superscript = v;
            if v {
                run.style.subscript = false;
            }
        }
        if let Some(v) = patch.subscript {
            run.style.subscript = v;
            if v {
                run.style.superscript = false;
            }
        }
        if let Some(f) = &patch.font_family {
            run.style.font_family = Some(f.clone());
        }
        if let Some(sz) = patch.font_size {
            run.style.font_size = Some(sz);
        }
        if let Some(c) = patch.color {
            run.style.color = Some(c);
        }
        if let Some(bg) = patch.background {
            run.style.background = Some(bg);
        }
    }

    merge_adjacent_runs(&mut paragraph.runs);
}

fn split_runs_at(runs: &mut Vec<Run>, offset: usize) -> usize {
    if offset == 0 {
        return 0;
    }

    let mut acc = 0_usize;
    for i in 0..runs.len() {
        let len = runs[i].text.len();
        let end = acc + len;
        if offset == acc {
            return i;
        }
        if offset == end {
            return i + 1;
        }
        if offset > acc && offset < end {
            let mut cut = offset - acc;
            while cut > 0 && !runs[i].text.is_char_boundary(cut) {
                cut -= 1;
            }

            let tail = runs[i].text[cut..].to_string();
            let style = runs[i].style.clone();
            runs[i].text.truncate(cut);
            runs.insert(i + 1, Run { text: tail, style });
            return i + 1;
        }
        acc = end;
    }

    runs.len()
}

fn merge_adjacent_runs(runs: &mut Vec<Run>) {
    let mut i = 0;
    while i + 1 < runs.len() {
        if runs[i].style == runs[i + 1].style {
            let tail = runs[i + 1].text.clone();
            runs[i].text.push_str(&tail);
            runs.remove(i + 1);
        } else {
            i += 1;
        }
    }
}

fn find_or_create_run(
    doc: &mut DocumentModel,
    block_id: crate::document::model::BlockId,
) -> Option<(&mut Run, usize)> {
    let idx = doc.content.iter().position(|b| match b {
        Block::Paragraph(p) => p.id == block_id,
        _ => false,
    })?;

    let paragraph = match &mut doc.content[idx] {
        Block::Paragraph(p) => p,
        _ => return None,
    };

    if paragraph.runs.is_empty() {
        paragraph.runs.push(Run::default());
    }

    Some((&mut paragraph.runs[0], idx))
}

fn find_paragraph_mut(
    doc: &mut DocumentModel,
    block_id: crate::document::model::BlockId,
) -> Option<&mut Paragraph> {
    doc.content.iter_mut().find_map(|b| match b {
        Block::Paragraph(p) if p.id == block_id => Some(p),
        _ => None,
    })
}

fn find_block_index_by_id(
    doc: &DocumentModel,
    block_id: crate::document::model::BlockId,
) -> Option<usize> {
    doc.content
        .iter()
        .position(|block| block_id_of(block) == Some(block_id))
}

fn block_id_of(block: &Block) -> Option<crate::document::model::BlockId> {
    match block {
        Block::Paragraph(p) => Some(p.id),
        Block::Table(t) => Some(t.id),
        Block::Image(i) => Some(i.id),
        Block::BlockQuote(q) => Some(q.id),
        Block::CodeBlock(c) => Some(c.id),
        Block::Heading(h) => Some(h.id),
        _ => None,
    }
}

fn paragraph_text_len(paragraph: &Paragraph) -> usize {
    paragraph.runs.iter().map(|r| r.text.len()).sum()
}

fn make_empty_paragraph(id: crate::document::model::BlockId) -> Paragraph {
    Paragraph {
        id,
        runs: vec![Run::default()],
        alignment: ParagraphAlignment::Left,
        spacing: crate::document::model::ParagraphSpacing::default(),
        indent: crate::document::model::Indent::default(),
        style_id: None,
    }
}

fn next_block_id(doc: &DocumentModel) -> crate::document::model::BlockId {
    let max_id = doc
        .content
        .iter()
        .filter_map(block_id_of)
        .map(|id| id.0)
        .max()
        .unwrap_or(0);
    crate::document::model::BlockId(max_id + 1)
}

fn estimate_command_size(cmd: &EditCommand) -> usize {
    match cmd {
        EditCommand::InsertText { text, .. } => text.len(),
        EditCommand::ReplaceText { text, .. } => text.len(),
        EditCommand::ReplaceRuns { runs, .. } => runs.iter().map(|r| r.text.len() + 32).sum(),
        EditCommand::RestoreBlock { block, .. } => match block {
            Block::Paragraph(p) => p.runs.iter().map(|r| r.text.len() + 32).sum::<usize>() + 64,
            _ => 128,
        },
        EditCommand::ReplaceParagraph { paragraph, .. } => {
            paragraph
                .runs
                .iter()
                .map(|r| r.text.len() + 32)
                .sum::<usize>()
                + 64
        }
        _ => 24,
    }
}

#[allow(dead_code)]
fn _new_paragraph_with_id(id: crate::document::model::BlockId) -> Paragraph {
    Paragraph {
        id,
        runs: vec![Run::default()],
        alignment: ParagraphAlignment::Left,
        spacing: crate::document::model::ParagraphSpacing::default(),
        indent: crate::document::model::Indent::default(),
        style_id: None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::document::model::{BlockId, DocumentModel};

    fn model_with_text(text: &str) -> DocumentModel {
        let mut doc = DocumentModel::default();
        doc.content.push(Block::Paragraph(Paragraph {
            id: BlockId(1),
            runs: vec![Run {
                text: text.to_string(),
                style: RunStyle::default(),
            }],
            alignment: ParagraphAlignment::Left,
            spacing: crate::document::model::ParagraphSpacing::default(),
            indent: crate::document::model::Indent::default(),
            style_id: None,
        }));
        doc
    }

    #[test]
    fn split_and_merge_are_inverse() {
        let mut doc = model_with_text("hello world");
        let inverse = apply_to_document(
            &mut doc,
            &EditCommand::SplitBlock {
                block_id: BlockId(1),
                offset: 5,
            },
        )
        .expect("split should produce inverse");

        assert_eq!(doc.content.len(), 2);
        let left = match &doc.content[0] {
            Block::Paragraph(p) => p.runs[0].text.clone(),
            _ => String::new(),
        };
        let right = match &doc.content[1] {
            Block::Paragraph(p) => p.runs[0].text.clone(),
            _ => String::new(),
        };
        assert_eq!(left, "hello");
        assert_eq!(right, " world");

        let _ = apply_to_document(&mut doc, &inverse).expect("merge should succeed");
        assert_eq!(doc.content.len(), 1);
        let merged = match &doc.content[0] {
            Block::Paragraph(p) => p.runs[0].text.clone(),
            _ => String::new(),
        };
        assert_eq!(merged, "hello world");
    }

    #[test]
    fn engine_undo_redo_roundtrip() {
        let mut doc = model_with_text("abc");
        let mut engine = EditEngine::default();

        engine.apply_command(&mut doc, EditCommand::InsertBlock { at_index: 1 });
        assert_eq!(doc.content.len(), 2);

        engine.undo(&mut doc);
        assert_eq!(doc.content.len(), 1);

        engine.redo(&mut doc);
        assert_eq!(doc.content.len(), 2);
    }
}
