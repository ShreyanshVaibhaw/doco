use std::{fs, path::Path};

use encoding_rs::{Encoding, UTF_16BE, UTF_16LE, UTF_8, WINDOWS_1252};
use ropey::Rope;

use crate::document::DocumentFormat;
use crate::document::model::{
    Block,
    BlockId,
    DocumentModel,
    Paragraph,
    ParagraphAlignment,
    ParagraphSpacing,
    Run,
    RunStyle,
};

pub mod renderer;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextWrapMode {
    None,
    WordBoundary,
    Character,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextEditError {
    LineOutOfBounds,
    ColumnOutOfBounds,
}

#[derive(Debug)]
pub struct TextDocument {
    pub rope: Rope,
    pub encoding_name: String,
    pub monospaced: bool,
    pub line_numbers: bool,
    pub wrap_mode: TextWrapMode,
}

impl TextDocument {
    pub fn load_from_path(path: &Path) -> std::io::Result<Self> {
        let bytes = fs::read(path)?;
        Ok(Self::from_bytes(&bytes))
    }

    pub fn from_bytes(bytes: &[u8]) -> Self {
        let (text, encoding_name) = decode_text(bytes);
        Self {
            rope: Rope::from_str(&text),
            encoding_name,
            monospaced: true,
            line_numbers: true,
            wrap_mode: TextWrapMode::WordBoundary,
        }
    }

    pub fn from_text(text: &str) -> Self {
        Self {
            rope: Rope::from_str(text),
            encoding_name: "UTF-8".to_string(),
            monospaced: true,
            line_numbers: true,
            wrap_mode: TextWrapMode::WordBoundary,
        }
    }

    pub fn line_count(&self) -> usize {
        self.rope.len_lines()
    }

    pub fn char_count(&self) -> usize {
        self.rope.len_chars()
    }

    pub fn set_monospaced(&mut self, monospaced: bool) {
        self.monospaced = monospaced;
    }

    pub fn set_line_numbers(&mut self, enabled: bool) {
        self.line_numbers = enabled;
    }

    pub fn set_wrap_mode(&mut self, wrap_mode: TextWrapMode) {
        self.wrap_mode = wrap_mode;
    }

    pub fn line_text(&self, line: usize) -> Option<String> {
        if line >= self.line_count() {
            return None;
        }
        Some(trim_line_breaks(self.rope.line(line).to_string()))
    }

    pub fn insert_text(
        &mut self,
        line: usize,
        column: usize,
        text: &str,
    ) -> Result<(), TextEditError> {
        let index = self.position_to_char_index(line, column)?;
        self.rope.insert(index, text);
        Ok(())
    }

    pub fn delete_range(
        &mut self,
        start: (usize, usize),
        end: (usize, usize),
    ) -> Result<String, TextEditError> {
        let start_index = self.position_to_char_index(start.0, start.1)?;
        let end_index = self.position_to_char_index(end.0, end.1)?;
        let (from, to) = if start_index <= end_index {
            (start_index, end_index)
        } else {
            (end_index, start_index)
        };

        if from == to {
            return Ok(String::new());
        }

        let removed = self.rope.slice(from..to).to_string();
        self.rope.remove(from..to);
        Ok(removed)
    }

    pub fn replace_range(
        &mut self,
        start: (usize, usize),
        end: (usize, usize),
        replacement: &str,
    ) -> Result<String, TextEditError> {
        let start_index = self.position_to_char_index(start.0, start.1)?;
        let end_index = self.position_to_char_index(end.0, end.1)?;
        let (from, to) = if start_index <= end_index {
            (start_index, end_index)
        } else {
            (end_index, start_index)
        };

        let removed = if from == to {
            String::new()
        } else {
            let removed = self.rope.slice(from..to).to_string();
            self.rope.remove(from..to);
            removed
        };
        self.rope.insert(from, replacement);
        Ok(removed)
    }

    pub fn to_document_model(&self) -> DocumentModel {
        let mut model = DocumentModel::default();
        model.metadata.format = DocumentFormat::Text;
        model.content = self
            .rope
            .lines()
            .enumerate()
            .map(|(i, line)| {
                Block::Paragraph(Paragraph {
                    id: BlockId(i as u64 + 1),
                    runs: vec![Run {
                        text: trim_line_breaks(line.to_string()),
                        style: RunStyle {
                            font_family: Some(if self.monospaced {
                                "Cascadia Mono".to_string()
                            } else {
                                "Segoe UI".to_string()
                            }),
                            ..RunStyle::default()
                        },
                    }],
                    alignment: ParagraphAlignment::Left,
                    spacing: ParagraphSpacing::default(),
                    indent: crate::document::model::Indent::default(),
                    style_id: None,
                })
            })
            .collect();
        model
    }

    fn position_to_char_index(&self, line: usize, column: usize) -> Result<usize, TextEditError> {
        let line_count = self.line_count();
        if line > line_count {
            return Err(TextEditError::LineOutOfBounds);
        }
        if line == line_count {
            if column == 0 {
                return Ok(self.char_count());
            }
            return Err(TextEditError::ColumnOutOfBounds);
        }

        let line_slice = self.rope.line(line);
        if column > line_slice.len_chars() {
            return Err(TextEditError::ColumnOutOfBounds);
        }

        Ok(self.rope.line_to_char(line) + column)
    }
}

fn decode_text(bytes: &[u8]) -> (String, String) {
    if bytes.is_empty() {
        return (String::new(), "UTF-8".to_string());
    }

    if let Some((encoding, bom_len)) = Encoding::for_bom(bytes) {
        let (text, _, _) = encoding.decode(&bytes[bom_len..]);
        return (text.into_owned(), encoding.name().to_string());
    }

    if likely_utf16(bytes, true) {
        let (text, _, _) = UTF_16LE.decode(bytes);
        return (text.into_owned(), "UTF-16LE (heuristic)".to_string());
    }

    if likely_utf16(bytes, false) {
        let (text, _, _) = UTF_16BE.decode(bytes);
        return (text.into_owned(), "UTF-16BE (heuristic)".to_string());
    }

    if let Ok(as_utf8) = std::str::from_utf8(bytes) {
        return (as_utf8.to_string(), UTF_8.name().to_string());
    }

    // Legacy fallback for ANSI/Latin-1-like text files on Windows.
    decode_with_encoding(bytes, WINDOWS_1252)
}

fn decode_with_encoding(bytes: &[u8], encoding: &'static Encoding) -> (String, String) {
    let (text, _, _) = encoding.decode(bytes);
    (text.into_owned(), encoding.name().to_string())
}

fn likely_utf16(bytes: &[u8], little_endian: bool) -> bool {
    if bytes.len() < 4 {
        return false;
    }

    let sample_len = bytes.len().min(4096) & !1;
    if sample_len < 4 {
        return false;
    }

    let mut zero_even = 0usize;
    let mut zero_odd = 0usize;
    for (i, b) in bytes[..sample_len].iter().enumerate() {
        if *b == 0 {
            if i % 2 == 0 {
                zero_even += 1;
            } else {
                zero_odd += 1;
            }
        }
    }

    let half = sample_len / 2;
    let even_ratio = zero_even as f32 / half as f32;
    let odd_ratio = zero_odd as f32 / half as f32;

    if little_endian {
        odd_ratio > 0.35 && even_ratio < 0.15
    } else {
        even_ratio > 0.35 && odd_ratio < 0.15
    }
}

fn trim_line_breaks(mut s: String) -> String {
    while s.ends_with('\n') || s.ends_with('\r') {
        s.pop();
    }
    s
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::document::model::Block;

    #[test]
    fn decodes_utf8_and_windows1252() {
        let utf8 = TextDocument::from_bytes("hello".as_bytes());
        assert_eq!(utf8.encoding_name, "UTF-8");
        assert_eq!(utf8.line_text(0).as_deref(), Some("hello"));

        let win1252 = TextDocument::from_bytes(&[0x63, 0x61, 0x66, 0xE9]);
        assert_eq!(win1252.encoding_name, "windows-1252");
        assert_eq!(win1252.line_text(0).as_deref(), Some("café"));
    }

    #[test]
    fn decodes_utf16_with_bom() {
        let mut bytes = vec![0xFF, 0xFE];
        for unit in "Hi".encode_utf16() {
            bytes.extend_from_slice(&unit.to_le_bytes());
        }
        let doc = TextDocument::from_bytes(&bytes);
        assert_eq!(doc.line_text(0).as_deref(), Some("Hi"));
    }

    #[test]
    fn supports_basic_edit_operations() {
        let mut doc = TextDocument::from_text("abc\ndef");
        doc.insert_text(0, 3, "!").expect("insert should succeed");
        assert_eq!(doc.line_text(0).as_deref(), Some("abc!"));

        let removed = doc
            .delete_range((0, 1), (0, 3))
            .expect("delete should succeed");
        assert_eq!(removed, "bc");
        assert_eq!(doc.line_text(0).as_deref(), Some("a!"));

        let replaced = doc
            .replace_range((1, 0), (1, 3), "XYZ")
            .expect("replace should succeed");
        assert_eq!(replaced, "def");
        assert_eq!(doc.line_text(1).as_deref(), Some("XYZ"));
    }

    #[test]
    fn to_document_model_respects_font_option() {
        let mut doc = TextDocument::from_text("line");
        doc.set_monospaced(false);
        let model = doc.to_document_model();
        let block = model.content.first().expect("first block expected");
        match block {
            Block::Paragraph(paragraph) => {
                assert_eq!(
                    paragraph.runs.first().and_then(|r| r.style.font_family.as_deref()),
                    Some("Segoe UI")
                );
            }
            _ => panic!("expected paragraph"),
        }
    }
}
