use std::{fs, path::{Path, PathBuf}};

use pulldown_cmark::{Options, Parser};

use crate::document::{
    DocumentFormat,
    model::DocumentModel,
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
    pub source_path: Option<PathBuf>,
    pub mode: MarkdownViewMode,
}

impl MarkdownDocument {
    pub fn load_from_path(path: &Path) -> std::io::Result<Self> {
        let source = fs::read(path)?;
        Ok(Self::from_source_with_path(
            String::from_utf8_lossy(&source).into_owned(),
            Some(path.to_path_buf()),
        ))
    }

    pub fn from_source(source: impl Into<String>) -> Self {
        Self::from_source_with_path(source.into(), None)
    }

    pub fn from_source_with_path(source: String, source_path: Option<PathBuf>) -> Self {
        Self {
            source,
            source_path,
            mode: MarkdownViewMode::Rendered,
        }
    }

    pub fn parser(&self) -> Parser<'_> {
        let options = Options::ENABLE_TABLES
            | Options::ENABLE_STRIKETHROUGH
            | Options::ENABLE_TASKLISTS
            | Options::ENABLE_FOOTNOTES
            | Options::ENABLE_HEADING_ATTRIBUTES
            | Options::ENABLE_GFM
            | Options::ENABLE_MATH
            | Options::ENABLE_WIKILINKS;
        Parser::new_ext(&self.source, options)
    }

    pub fn set_mode(&mut self, mode: MarkdownViewMode) {
        self.mode = mode;
    }

    pub fn set_source(&mut self, source: impl Into<String>) {
        self.source = source.into();
    }

    pub fn source(&self) -> &str {
        self.source.as_str()
    }

    pub fn source_highlights(&self) -> Vec<renderer::MarkdownSourceHighlightSpan> {
        renderer::highlight_markdown_source(&self.source)
    }

    pub fn view_snapshot(&self) -> renderer::MarkdownViewSnapshot {
        renderer::build_view_snapshot(self)
    }

    pub fn outline(&self) -> Vec<renderer::MarkdownOutlineEntry> {
        renderer::collect_outline(self)
    }

    pub fn to_document_model(&self) -> DocumentModel {
        let mut model = renderer::markdown_to_model(self, self.source_path.as_deref());
        model.metadata.format = DocumentFormat::Markdown;
        model
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn snapshot_respects_mode() {
        let mut doc = MarkdownDocument::from_source("# Title\n\ntext");

        doc.set_mode(MarkdownViewMode::Rendered);
        match doc.view_snapshot() {
            renderer::MarkdownViewSnapshot::Rendered { model } => assert!(!model.content.is_empty()),
            _ => panic!("expected rendered snapshot"),
        }

        doc.set_mode(MarkdownViewMode::Source);
        match doc.view_snapshot() {
            renderer::MarkdownViewSnapshot::Source { lines, .. } => assert!(!lines.is_empty()),
            _ => panic!("expected source snapshot"),
        }

        doc.set_mode(MarkdownViewMode::Split);
        match doc.view_snapshot() {
            renderer::MarkdownViewSnapshot::Split { lines, preview, .. } => {
                assert!(!lines.is_empty());
                assert!(!preview.content.is_empty());
            }
            _ => panic!("expected split snapshot"),
        }
    }
}
