use std::{collections::HashMap, path::PathBuf};

use chrono::{DateTime, Utc};
use serde::{Deserialize, Serialize};

use crate::{document::DocumentFormat, ui::Color};

pub type DocumentModel = Document;

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct Document {
    pub metadata: DocumentMetadata,
    pub pages: Vec<Page>,
    pub content: Vec<Block>,
    pub styles: StyleSheet,
    pub images: HashMap<String, ImageData>,
    pub dirty: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DocumentMetadata {
    pub title: String,
    pub author: String,
    pub created: Option<DateTime<Utc>>,
    pub modified: Option<DateTime<Utc>>,
    pub file_path: Option<PathBuf>,
    pub format: DocumentFormat,
    pub page_size: PageSize,
    pub margins: Margins,
}

impl Default for DocumentMetadata {
    fn default() -> Self {
        Self {
            title: String::new(),
            author: String::new(),
            created: None,
            modified: None,
            file_path: None,
            format: DocumentFormat::Unknown,
            page_size: PageSize::Letter,
            margins: Margins::default(),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct Page {
    pub index: usize,
    pub width: f32,
    pub height: f32,
    pub block_ids: Vec<BlockId>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum Block {
    Paragraph(Paragraph),
    Table(Table),
    Image(ImageBlock),
    PageBreak,
    HorizontalRule,
    List(List),
    BlockQuote(BlockQuote),
    CodeBlock(CodeBlock),
    Heading(Heading),
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq, Hash, Default)]
pub struct BlockId(pub u64);

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Paragraph {
    pub id: BlockId,
    pub runs: Vec<Run>,
    pub alignment: ParagraphAlignment,
    pub spacing: ParagraphSpacing,
    pub indent: Indent,
    pub style_id: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct Run {
    pub text: String,
    pub style: RunStyle,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default, PartialEq)]
pub struct RunStyle {
    pub font_family: Option<String>,
    pub font_size: Option<f32>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub color: Option<Color>,
    pub background: Option<Color>,
    pub superscript: bool,
    pub subscript: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct Table {
    pub id: BlockId,
    pub rows: Vec<TableRow>,
    pub column_widths: Vec<f32>,
    pub borders: TableBorders,
    pub style: TableStylePreset,
    pub cell_padding: f32,
    pub header_row: bool,
    pub alternating_rows: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct Heading {
    pub level: u8,
    pub runs: Vec<Run>,
    pub id: BlockId,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct List {
    pub items: Vec<ListItem>,
    pub list_type: ListType,
    pub start_number: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct ListItem {
    pub id: BlockId,
    pub content: Vec<Block>,
    pub checked: Option<bool>,
    pub children: Vec<ListItem>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct ImageBlock {
    pub id: BlockId,
    pub data: ImageDataRef,
    pub original_width: u32,
    pub original_height: u32,
    pub caption: Option<Vec<Run>>,
    pub border: Option<ImageBorder>,
    pub crop: Option<CropRect>,
    pub key: String,
    pub alt_text: String,
    pub source_path: Option<PathBuf>,
    pub width: f32,
    pub height: f32,
    pub alignment: ImageAlignment,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct BlockQuote {
    pub id: BlockId,
    pub blocks: Vec<Block>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct CodeBlock {
    pub id: BlockId,
    pub language: Option<String>,
    pub code: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct TableCell {
    pub blocks: Vec<Block>,
    pub rowspan: u16,
    pub colspan: u16,
    pub background: Option<Color>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum PageSize {
    Letter,
    A4,
    Legal,
    Custom {
        width_points: f32,
        height_points: f32,
    },
}

impl Default for PageSize {
    fn default() -> Self {
        Self::Letter
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum ParagraphAlignment {
    Left,
    Center,
    Right,
    Justify,
}

impl Default for ParagraphAlignment {
    fn default() -> Self {
        Self::Left
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct ParagraphSpacing {
    pub before: f32,
    pub after: f32,
    pub line: f32,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct Indent {
    pub left: f32,
    pub right: f32,
    pub first_line: f32,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct Margins {
    pub top: f32,
    pub right: f32,
    pub bottom: f32,
    pub left: f32,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct StyleSheet {
    pub styles: HashMap<String, NamedStyle>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct NamedStyle {
    pub id: String,
    pub name: String,
    pub run_style: RunStyle,
    pub paragraph_style: Option<ParagraphStyle>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct ParagraphStyle {
    pub alignment: ParagraphAlignment,
    pub spacing: ParagraphSpacing,
    pub indent: Indent,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct ImageData {
    pub bytes: Vec<u8>,
    pub mime: String,
    pub width: u32,
    pub height: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub enum ImageDataRef {
    #[default]
    Empty,
    Embedded(ImageData),
    LinkedPath(PathBuf),
    Key(String),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum ImageAlignment {
    Inline,
    Left,
    Center,
    Right,
    Float,
}

impl Default for ImageAlignment {
    fn default() -> Self {
        Self::Inline
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum ImageBorderStyle {
    Solid,
    Dashed,
    Dotted,
}

impl Default for ImageBorderStyle {
    fn default() -> Self {
        Self::Solid
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct ImageBorder {
    pub style: ImageBorderStyle,
    pub width: f32,
    pub color: Color,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct CropRect {
    pub left: f32,
    pub top: f32,
    pub right: f32,
    pub bottom: f32,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct TableBorders {
    pub outer: BorderStyle,
    pub inner_horizontal: BorderStyle,
    pub inner_vertical: BorderStyle,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum TableStylePreset {
    Plain,
    Grid,
    HeaderAccent,
    AlternatingRows,
    Professional,
}

impl Default for TableStylePreset {
    fn default() -> Self {
        Self::Grid
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BorderStyle {
    pub width: f32,
    pub color: Color,
}

impl Default for BorderStyle {
    fn default() -> Self {
        Self {
            width: 1.0,
            color: Color::rgb(0.7, 0.7, 0.7),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum ListType {
    Bullet,
    Numbered,
    Checkbox,
}

impl Default for ListType {
    fn default() -> Self {
        Self::Bullet
    }
}

impl Document {
    pub fn next_block_id(&self) -> BlockId {
        fn walk(block: &Block, max: &mut u64) {
            match block {
                Block::Paragraph(p) => *max = (*max).max(p.id.0),
                Block::Heading(h) => *max = (*max).max(h.id.0),
                Block::CodeBlock(c) => *max = (*max).max(c.id.0),
                Block::Image(i) => *max = (*max).max(i.id.0),
                Block::Table(t) => {
                    *max = (*max).max(t.id.0);
                    for row in &t.rows {
                        for cell in &row.cells {
                            for nested in &cell.blocks {
                                walk(nested, max);
                            }
                        }
                    }
                }
                Block::List(list) => {
                    for item in &list.items {
                        *max = (*max).max(item.id.0);
                        for nested in &item.content {
                            walk(nested, max);
                        }
                        for child in &item.children {
                            *max = (*max).max(child.id.0);
                            for nested in &child.content {
                                walk(nested, max);
                            }
                        }
                    }
                }
                Block::BlockQuote(q) => {
                    *max = (*max).max(q.id.0);
                    for nested in &q.blocks {
                        walk(nested, max);
                    }
                }
                Block::PageBreak | Block::HorizontalRule => {}
            }
        }

        let mut max = 0;
        for block in &self.content {
            walk(block, &mut max);
        }
        BlockId(max + 1)
    }

    pub fn insert_embedded_image_after(
        &mut self,
        after_block_id: Option<BlockId>,
        bytes: Vec<u8>,
        mime: String,
        width: u32,
        height: u32,
        source_path: Option<PathBuf>,
        alt_text: String,
    ) -> BlockId {
        let block_id = self.next_block_id();
        let key = format!("image-{}", block_id.0);
        let image_data = ImageData {
            bytes,
            mime,
            width,
            height,
        };
        self.images.insert(key.clone(), image_data.clone());

        let image_block = Block::Image(ImageBlock {
            id: block_id,
            data: ImageDataRef::Embedded(image_data),
            original_width: width,
            original_height: height,
            width: width as f32,
            height: height as f32,
            alignment: ImageAlignment::Inline,
            caption: None,
            alt_text,
            border: None,
            crop: None,
            key,
            source_path,
        });

        let insert_index = after_block_id
            .and_then(|target| {
                self.content
                    .iter()
                    .position(|block| block_id_for_block(block) == Some(target))
                    .map(|idx| idx + 1)
            })
            .unwrap_or(self.content.len());
        self.content.insert(insert_index, image_block);
        self.dirty = true;
        block_id
    }

    pub fn insert_embedded_image(
        &mut self,
        block_id: BlockId,
        bytes: Vec<u8>,
        mime: String,
        width: u32,
        height: u32,
    ) {
        let key = format!("image-{}", block_id.0);
        let image_data = ImageData {
            bytes,
            mime,
            width,
            height,
        };
        self.images.insert(key.clone(), image_data.clone());
        self.content.push(Block::Image(ImageBlock {
            id: block_id,
            data: ImageDataRef::Embedded(image_data),
            original_width: width,
            original_height: height,
            width: width as f32,
            height: height as f32,
            alignment: ImageAlignment::Inline,
            caption: None,
            alt_text: String::new(),
            border: None,
            crop: None,
            key,
            source_path: None,
        }));
        self.dirty = true;
    }

    pub fn find_image_block_mut(&mut self, block_id: BlockId) -> Option<&mut ImageBlock> {
        fn walk(block: &mut Block, block_id: BlockId) -> Option<&mut ImageBlock> {
            match block {
                Block::Image(image) if image.id == block_id => Some(image),
                Block::Table(table) => {
                    for row in &mut table.rows {
                        for cell in &mut row.cells {
                            for nested in &mut cell.blocks {
                                if let Some(image) = walk(nested, block_id) {
                                    return Some(image);
                                }
                            }
                        }
                    }
                    None
                }
                Block::List(list) => {
                    for item in &mut list.items {
                        for nested in &mut item.content {
                            if let Some(image) = walk(nested, block_id) {
                                return Some(image);
                            }
                        }
                        for child in &mut item.children {
                            for nested in &mut child.content {
                                if let Some(image) = walk(nested, block_id) {
                                    return Some(image);
                                }
                            }
                        }
                    }
                    None
                }
                Block::BlockQuote(q) => {
                    for nested in &mut q.blocks {
                        if let Some(image) = walk(nested, block_id) {
                            return Some(image);
                        }
                    }
                    None
                }
                _ => None,
            }
        }

        for block in &mut self.content {
            if let Some(image) = walk(block, block_id) {
                return Some(image);
            }
        }
        None
    }

    pub fn remove_image_block(&mut self, block_id: BlockId) -> bool {
        let mut removed = false;
        let mut removed_key = None;
        self.content.retain(|block| {
            let keep = match block {
                Block::Image(image) if image.id == block_id => {
                    removed_key = Some(image.key.clone());
                    false
                }
                _ => true,
            };
            if !keep {
                removed = true;
            }
            keep
        });

        if let Some(key) = removed_key {
            self.images.remove(key.as_str());
        }
        if removed {
            self.dirty = true;
        }
        removed
    }
}

fn block_id_for_block(block: &Block) -> Option<BlockId> {
    match block {
        Block::Paragraph(p) => Some(p.id),
        Block::Heading(h) => Some(h.id),
        Block::CodeBlock(c) => Some(c.id),
        Block::Table(t) => Some(t.id),
        Block::Image(i) => Some(i.id),
        Block::BlockQuote(q) => Some(q.id),
        Block::List(list) => list.items.first().map(|item| item.id),
        Block::PageBreak | Block::HorizontalRule => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn paragraph(id: u64, text: &str) -> Block {
        Block::Paragraph(Paragraph {
            id: BlockId(id),
            runs: vec![Run {
                text: text.to_string(),
                style: RunStyle::default(),
            }],
            alignment: ParagraphAlignment::Left,
            spacing: ParagraphSpacing::default(),
            indent: Indent::default(),
            style_id: None,
        })
    }

    #[test]
    fn insert_embedded_image_after_cursor_block() {
        let mut doc = Document::default();
        doc.content.push(paragraph(1, "A"));
        doc.content.push(paragraph(2, "B"));

        let inserted = doc.insert_embedded_image_after(
            Some(BlockId(1)),
            vec![1, 2, 3],
            "image/png".to_string(),
            100,
            60,
            None,
            "alt".to_string(),
        );

        assert_eq!(inserted, BlockId(3));
        assert!(matches!(doc.content[1], Block::Image(_)));
        assert_eq!(doc.images.len(), 1);
        assert!(doc.dirty);
    }

    #[test]
    fn remove_image_block_cleans_document_and_map() {
        let mut doc = Document::default();
        let inserted = doc.insert_embedded_image_after(
            None,
            vec![9, 8, 7],
            "image/png".to_string(),
            32,
            32,
            None,
            String::new(),
        );

        assert!(doc.remove_image_block(inserted));
        assert!(doc.images.is_empty());
        assert!(doc.content.iter().all(|b| !matches!(b, Block::Image(_))));
    }
}
