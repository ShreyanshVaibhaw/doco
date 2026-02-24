use std::{fs, path::Path};

use encoding_rs::{Encoding, UTF_16BE, UTF_16LE, UTF_8, WINDOWS_1252};
use ropey::Rope;

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
use crate::document::DocumentFormat;

pub mod renderer;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextWrapMode {
    None,
    WordBoundary,
    Character,
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
        let (text, encoding_name) = decode_text(&bytes);

        Ok(Self {
            rope: Rope::from_str(&text),
            encoding_name,
            monospaced: true,
            line_numbers: true,
            wrap_mode: TextWrapMode::WordBoundary,
        })
    }

    pub fn line_count(&self) -> usize {
        self.rope.len_lines()
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
                        text: line.to_string(),
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
}

fn decode_text(bytes: &[u8]) -> (String, String) {
    if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
        let (text, _, _) = UTF_8.decode(&bytes[3..]);
        return (text.into_owned(), "UTF-8".to_string());
    }
    if bytes.starts_with(&[0xFF, 0xFE]) {
        let (text, _, _) = UTF_16LE.decode(&bytes[2..]);
        return (text.into_owned(), "UTF-16LE".to_string());
    }
    if bytes.starts_with(&[0xFE, 0xFF]) {
        let (text, _, _) = UTF_16BE.decode(&bytes[2..]);
        return (text.into_owned(), "UTF-16BE".to_string());
    }

    if let Ok(as_utf8) = std::str::from_utf8(bytes) {
        return (as_utf8.to_string(), "UTF-8".to_string());
    }

    // Fallback for legacy Windows text files.
    decode_with_encoding(bytes, WINDOWS_1252)
}

fn decode_with_encoding(bytes: &[u8], encoding: &'static Encoding) -> (String, String) {
    let (text, _, _) = encoding.decode(bytes);
    (text.into_owned(), encoding.name().to_string())
}
