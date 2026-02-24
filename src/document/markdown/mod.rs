use std::{fs, path::Path};

use pulldown_cmark::{Options, Parser};

use crate::document::{
    DocumentFormat,
    model::{Block, DocumentModel},
};

pub mod renderer;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MarkdownViewMode {
    Rendered,
    Source,
    Split,
}

#[derive(Debug)]
pub struct MarkdownDocument {
    pub source: String,
    pub mode: MarkdownViewMode,
}

impl MarkdownDocument {
    pub fn load_from_path(path: &Path) -> std::io::Result<Self> {
        let source = fs::read_to_string(path)?;
        Ok(Self {
            source,
            mode: MarkdownViewMode::Rendered,
        })
    }

    pub fn parser(&self) -> Parser<'_> {
        let options = Options::ENABLE_TABLES
            | Options::ENABLE_STRIKETHROUGH
            | Options::ENABLE_TASKLISTS
            | Options::ENABLE_FOOTNOTES
            | Options::ENABLE_HEADING_ATTRIBUTES;
        Parser::new_ext(&self.source, options)
    }

    pub fn to_document_model(&self) -> DocumentModel {
        let mut model = renderer::markdown_to_model(self);
        model.metadata.format = DocumentFormat::Markdown;
        if model.content.is_empty() {
            model.content.push(Block::Paragraph(crate::document::model::Paragraph {
                id: crate::document::model::BlockId(1),
                runs: vec![crate::document::model::Run {
                    text: self.source.clone(),
                    style: crate::document::model::RunStyle::default(),
                }],
                alignment: crate::document::model::ParagraphAlignment::Left,
                spacing: crate::document::model::ParagraphSpacing::default(),
                indent: crate::document::model::Indent::default(),
                style_id: None,
            }));
        }
        model
    }
}
