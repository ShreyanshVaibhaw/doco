use std::{collections::HashMap, path::PathBuf};

use chrono::{DateTime, Utc};
use serde::{Deserialize, Serialize};

use crate::{
    document::DocumentFormat,
    ui::Color,
};

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
    Custom { width_points: f32, height_points: f32 },
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
}
