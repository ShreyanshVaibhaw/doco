use std::collections::{HashMap, HashSet};
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::mpsc::{self, Receiver, TryRecvError};
use std::thread;

use image::GenericImageView;
use pulldown_cmark::{Alignment, CodeBlockKind, Event, HeadingLevel, Tag, TagEnd};
use regex::Regex;

use crate::document::{
    markdown::{MarkdownDocument, MarkdownViewMode},
    model::{
        Block, BlockId, BlockQuote, CodeBlock, DocumentModel, Heading, ImageAlignment, ImageBlock,
        ImageDataRef, List, ListItem, ListType, Paragraph, ParagraphAlignment, ParagraphSpacing,
        Run, RunStyle, Table, TableCell, TableRow,
    },
};
use crate::ui::Color;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MarkdownOutlineEntry {
    pub block_id: BlockId,
    pub level: u8,
    pub title: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MarkdownSourceTokenKind {
    Heading,
    Emphasis,
    CodeFence,
    Link,
    Quote,
    ListMarker,
    Rule,
    TableRow,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MarkdownSourceHighlightSpan {
    pub start: usize,
    pub end: usize,
    pub kind: MarkdownSourceTokenKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MarkdownCodeTokenKind {
    Keyword,
    Number,
    String,
    Comment,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MarkdownCodeHighlightToken {
    pub start: usize,
    pub end: usize,
    pub kind: MarkdownCodeTokenKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MarkdownSourceLine {
    pub line_number: usize,
    pub text: String,
}

#[derive(Debug, Clone)]
pub enum MarkdownViewSnapshot {
    Rendered {
        model: DocumentModel,
    },
    Source {
        lines: Vec<MarkdownSourceLine>,
        highlights: Vec<MarkdownSourceHighlightSpan>,
    },
    Split {
        lines: Vec<MarkdownSourceLine>,
        highlights: Vec<MarkdownSourceHighlightSpan>,
        preview: DocumentModel,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MarkdownImageLoadStatus {
    Placeholder,
    Ready,
    Failed,
}

#[derive(Debug, Clone)]
pub struct MarkdownImageAsset {
    pub source: String,
    pub alt_text: String,
    pub status: MarkdownImageLoadStatus,
    pub width: u32,
    pub height: u32,
}

#[derive(Debug)]
pub struct MarkdownImageLoader {
    base_dir: Option<PathBuf>,
    cache: HashMap<String, MarkdownImageAsset>,
    pending: HashMap<String, Receiver<Result<(u32, u32), String>>>,
}

impl MarkdownImageLoader {
    pub fn new(base_dir: Option<PathBuf>) -> Self {
        Self {
            base_dir,
            cache: HashMap::new(),
            pending: HashMap::new(),
        }
    }

    pub fn request_image(&mut self, source: &str, alt_text: &str) -> MarkdownImageAsset {
        if let Some(existing) = self.cache.get(source) {
            return existing.clone();
        }

        let placeholder = MarkdownImageAsset {
            source: source.to_string(),
            alt_text: alt_text.to_string(),
            status: MarkdownImageLoadStatus::Placeholder,
            width: 0,
            height: 0,
        };
        self.cache.insert(source.to_string(), placeholder.clone());

        let source_owned = source.to_string();
        let base_dir = self.base_dir.clone();
        let (tx, rx) = mpsc::channel();
        thread::spawn(move || {
            let _ = tx.send(load_image_dimensions(source_owned.as_str(), base_dir.as_deref()));
        });
        self.pending.insert(source.to_string(), rx);

        placeholder
    }

    pub fn poll(&mut self) -> usize {
        let keys: Vec<String> = self.pending.keys().cloned().collect();
        let mut resolved = 0usize;

        for key in keys {
            let result = if let Some(rx) = self.pending.get(&key) {
                match rx.try_recv() {
                    Ok(value) => Some(value),
                    Err(TryRecvError::Empty) => None,
                    Err(TryRecvError::Disconnected) => {
                        Some(Err("image worker disconnected".to_string()))
                    }
                }
            } else {
                None
            };

            if let Some(result) = result {
                resolved += 1;
                self.pending.remove(&key);
                if let Some(entry) = self.cache.get_mut(&key) {
                    match result {
                        Ok((w, h)) => {
                            entry.status = MarkdownImageLoadStatus::Ready;
                            entry.width = w;
                            entry.height = h;
                        }
                        Err(_) => entry.status = MarkdownImageLoadStatus::Failed,
                    }
                }
            }
        }

        resolved
    }

    pub fn image(&self, source: &str) -> Option<&MarkdownImageAsset> {
        self.cache.get(source)
    }
}

#[derive(Default)]
struct ListBuilder {
    list_type: ListType,
    start_number: u32,
    items: Vec<ListItem>,
    current_item_runs: Vec<Run>,
    current_item_checked: Option<bool>,
}

#[derive(Default)]
struct TableBuilder {
    alignments: Vec<Alignment>,
    rows: Vec<TableRow>,
    current_row_cells: Vec<TableCell>,
    current_cell_runs: Vec<Run>,
}
pub fn markdown_to_model(doc: &MarkdownDocument, base_path: Option<&Path>) -> DocumentModel {
    let mut model = DocumentModel::default();
    let mut next_id = 1_u64;

    let mut in_paragraph = false;
    let mut current_runs: Vec<Run> = Vec::new();

    let mut in_heading: Option<u8> = None;
    let mut heading_runs: Vec<Run> = Vec::new();

    let mut list_stack: Vec<ListBuilder> = Vec::new();

    let mut in_code_block = false;
    let mut code_lang: Option<String> = None;
    let mut code_text = String::new();

    let mut in_block_quote = false;
    let mut quote_runs: Vec<Run> = Vec::new();

    let mut table_builder: Option<TableBuilder> = None;
    let mut table_in_cell = false;

    let mut image_capture: Option<(String, String)> = None;
    let mut footnote_label: Option<String> = None;

    let mut emphasis_depth = 0usize;
    let mut strong_depth = 0usize;
    let mut strike_depth = 0usize;
    let mut superscript_depth = 0usize;
    let mut subscript_depth = 0usize;
    let mut link_stack: Vec<String> = Vec::new();

    for event in doc.parser() {
        match event {
            Event::Start(tag) => match tag {
                Tag::Paragraph => {
                    in_paragraph = true;
                    current_runs.clear();
                }
                Tag::Heading { level, .. } => {
                    in_heading = Some(heading_level_to_u8(level));
                    heading_runs.clear();
                }
                Tag::List(start) => {
                    list_stack.push(ListBuilder {
                        list_type: if start.is_some() {
                            ListType::Numbered
                        } else {
                            ListType::Bullet
                        },
                        start_number: start.unwrap_or(1) as u32,
                        ..ListBuilder::default()
                    });
                }
                Tag::Item => {
                    if let Some(top) = list_stack.last_mut() {
                        top.current_item_runs.clear();
                        top.current_item_checked = None;
                    }
                }
                Tag::CodeBlock(kind) => {
                    in_code_block = true;
                    code_text.clear();
                    code_lang = match kind {
                        CodeBlockKind::Indented => None,
                        CodeBlockKind::Fenced(lang) => Some(lang.to_string()),
                    };
                }
                Tag::BlockQuote(_) => {
                    in_block_quote = true;
                    quote_runs.clear();
                }
                Tag::Table(alignments) => {
                    table_builder = Some(TableBuilder {
                        alignments,
                        ..TableBuilder::default()
                    });
                }
                Tag::TableRow => {
                    if let Some(table) = table_builder.as_mut() {
                        table.current_row_cells.clear();
                    }
                }
                Tag::TableCell => {
                    table_in_cell = true;
                    if let Some(table) = table_builder.as_mut() {
                        table.current_cell_runs.clear();
                    }
                }
                Tag::Emphasis => emphasis_depth += 1,
                Tag::Strong => strong_depth += 1,
                Tag::Strikethrough => strike_depth += 1,
                Tag::Superscript => superscript_depth += 1,
                Tag::Subscript => subscript_depth += 1,
                Tag::Link { dest_url, .. } => link_stack.push(dest_url.to_string()),
                Tag::Image { dest_url, .. } => {
                    image_capture = Some((dest_url.to_string(), String::new()));
                }
                Tag::FootnoteDefinition(label) => {
                    footnote_label = Some(label.to_string());
                }
                _ => {}
            },
            Event::End(tag) => match tag {
                TagEnd::Paragraph => {
                    if in_paragraph {
                        in_paragraph = false;
                        let block = Block::Paragraph(Paragraph {
                            id: BlockId(next_id),
                            runs: current_runs.clone(),
                            alignment: ParagraphAlignment::Left,
                            spacing: ParagraphSpacing::default(),
                            indent: crate::document::model::Indent::default(),
                            style_id: None,
                        });
                        next_id += 1;

                        if table_in_cell {
                            if let Some(table) = table_builder.as_mut() {
                                table.current_cell_runs.extend(extract_runs_from_block(&block));
                            }
                        } else if let Some(list) = list_stack.last_mut() {
                            list.current_item_runs.extend(extract_runs_from_block(&block));
                        } else if in_block_quote {
                            quote_runs.extend(extract_runs_from_block(&block));
                        } else {
                            model.content.push(block);
                        }
                    }
                }
                TagEnd::Heading(..) => {
                    if let Some(level) = in_heading.take() {
                        model.content.push(Block::Heading(Heading {
                            level,
                            runs: heading_runs.clone(),
                            id: BlockId(next_id),
                        }));
                        next_id += 1;
                    }
                }
                TagEnd::Item => {
                    if let Some(list) = list_stack.last_mut() {
                        let content = vec![Block::Paragraph(Paragraph {
                            id: BlockId(next_id),
                            runs: list.current_item_runs.clone(),
                            alignment: ParagraphAlignment::Left,
                            spacing: ParagraphSpacing::default(),
                            indent: crate::document::model::Indent::default(),
                            style_id: None,
                        })];
                        list.items.push(ListItem {
                            id: BlockId(next_id),
                            content,
                            checked: list.current_item_checked,
                            children: vec![],
                        });
                        next_id += 1;
                        list.current_item_runs.clear();
                        list.current_item_checked = None;
                    }
                }
                TagEnd::List(_) => {
                    if let Some(list) = list_stack.pop() {
                        model.content.push(Block::List(List {
                            items: list.items,
                            list_type: list.list_type,
                            start_number: list.start_number,
                        }));
                    }
                }
                TagEnd::CodeBlock => {
                    if in_code_block {
                        in_code_block = false;
                        model.content.push(Block::CodeBlock(CodeBlock {
                            id: BlockId(next_id),
                            language: code_lang.clone(),
                            code: code_text.clone(),
                        }));
                        next_id += 1;
                    }
                }
                TagEnd::BlockQuote(_) => {
                    if in_block_quote {
                        in_block_quote = false;
                        model.content.push(Block::BlockQuote(BlockQuote {
                            id: BlockId(next_id),
                            blocks: vec![Block::Paragraph(Paragraph {
                                id: BlockId(next_id + 1),
                                runs: quote_runs.clone(),
                                alignment: ParagraphAlignment::Left,
                                spacing: ParagraphSpacing::default(),
                                indent: crate::document::model::Indent::default(),
                                style_id: None,
                            })],
                        }));
                        next_id += 2;
                    }
                }
                TagEnd::TableCell => {
                    table_in_cell = false;
                    if let Some(table) = table_builder.as_mut() {
                        table.current_row_cells.push(TableCell {
                            blocks: vec![Block::Paragraph(Paragraph {
                                id: BlockId(next_id),
                                runs: table.current_cell_runs.clone(),
                                alignment: ParagraphAlignment::Left,
                                spacing: ParagraphSpacing::default(),
                                indent: crate::document::model::Indent::default(),
                                style_id: None,
                            })],
                            rowspan: 1,
                            colspan: 1,
                            background: None,
                        });
                        next_id += 1;
                    }
                }
                TagEnd::TableRow => {
                    if let Some(table) = table_builder.as_mut() {
                        table.rows.push(TableRow {
                            cells: std::mem::take(&mut table.current_row_cells),
                        });
                    }
                }
                TagEnd::Table => {
                    if let Some(table) = table_builder.take() {
                        let column_count = table
                            .alignments
                            .len()
                            .max(table.rows.first().map(|r| r.cells.len()).unwrap_or(0));
                        model.content.push(Block::Table(Table {
                            id: BlockId(next_id),
                            rows: table.rows,
                            column_widths: if column_count == 0 {
                                Vec::new()
                            } else {
                                vec![1.0 / column_count as f32; column_count]
                            },
                            borders: crate::document::model::TableBorders::default(),
                            style: crate::document::model::TableStylePreset::Grid,
                            cell_padding: 8.0,
                            header_row: true,
                            alternating_rows: true,
                        }));
                        next_id += 1;
                    }
                }
                TagEnd::Emphasis => emphasis_depth = emphasis_depth.saturating_sub(1),
                TagEnd::Strong => strong_depth = strong_depth.saturating_sub(1),
                TagEnd::Strikethrough => strike_depth = strike_depth.saturating_sub(1),
                TagEnd::Superscript => superscript_depth = superscript_depth.saturating_sub(1),
                TagEnd::Subscript => subscript_depth = subscript_depth.saturating_sub(1),
                TagEnd::Link => {
                    let _ = link_stack.pop();
                }
                TagEnd::Image => {
                    if let Some((src, alt_text)) = image_capture.take() {
                        model.content.push(Block::Image(build_image_block(
                            src.as_str(),
                            alt_text.as_str(),
                            base_path,
                            BlockId(next_id),
                        )));
                        next_id += 1;
                    }
                }
                TagEnd::FootnoteDefinition => {
                    footnote_label = None;
                }
                _ => {}
            },
            Event::Text(text) => {
                if let Some((_, alt_text)) = image_capture.as_mut() {
                    alt_text.push_str(text.as_ref());
                    continue;
                }

                let run = styled_run(
                    text.as_ref(),
                    strong_depth > 0,
                    emphasis_depth > 0,
                    strike_depth > 0,
                    superscript_depth > 0,
                    subscript_depth > 0,
                    link_stack.last().map(|s| s.as_str()),
                    false,
                );

                if in_code_block {
                    code_text.push_str(run.text.as_str());
                } else if in_heading.is_some() {
                    heading_runs.push(run);
                } else if in_paragraph {
                    current_runs.push(run);
                } else if in_block_quote {
                    quote_runs.push(run);
                }
            }
            Event::Code(text) => {
                let run = styled_run(
                    text.as_ref(),
                    strong_depth > 0,
                    emphasis_depth > 0,
                    strike_depth > 0,
                    superscript_depth > 0,
                    subscript_depth > 0,
                    link_stack.last().map(|s| s.as_str()),
                    true,
                );
                if in_heading.is_some() {
                    heading_runs.push(run);
                } else if in_paragraph {
                    current_runs.push(run);
                }
            }
            Event::TaskListMarker(checked) => {
                if let Some(list) = list_stack.last_mut() {
                    list.list_type = ListType::Checkbox;
                    list.current_item_checked = Some(checked);
                }
            }
            Event::FootnoteReference(label) => {
                if footnote_label.is_none() {
                    current_runs.push(Run {
                        text: format!("[{label}]"),
                        style: RunStyle {
                            superscript: true,
                            ..RunStyle::default()
                        },
                    });
                }
            }
            Event::Rule => model.content.push(Block::HorizontalRule),
            Event::SoftBreak | Event::HardBreak => {
                let run = Run {
                    text: "\n".to_string(),
                    style: RunStyle::default(),
                };
                if in_heading.is_some() {
                    heading_runs.push(run);
                } else if in_paragraph {
                    current_runs.push(run);
                } else if in_code_block {
                    code_text.push('\n');
                }
            }
            _ => {}
        }
    }

    model
}

pub fn render_markdown(model: &DocumentModel) -> usize {
    model.content.len()
}

pub fn collect_outline(doc: &MarkdownDocument) -> Vec<MarkdownOutlineEntry> {
    markdown_to_model(doc, doc.source_path.as_deref())
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Heading(h) => Some(MarkdownOutlineEntry {
                block_id: h.id,
                level: h.level,
                title: h.runs.iter().map(|r| r.text.as_str()).collect::<String>(),
            }),
            _ => None,
        })
        .collect()
}

pub fn build_view_snapshot(doc: &MarkdownDocument) -> MarkdownViewSnapshot {
    let lines = doc
        .source()
        .lines()
        .enumerate()
        .map(|(i, line)| MarkdownSourceLine {
            line_number: i + 1,
            text: line.to_string(),
        })
        .collect::<Vec<_>>();
    let highlights = highlight_markdown_source(doc.source());

    match doc.mode {
        MarkdownViewMode::Rendered => MarkdownViewSnapshot::Rendered {
            model: markdown_to_model(doc, doc.source_path.as_deref()),
        },
        MarkdownViewMode::Source => MarkdownViewSnapshot::Source { lines, highlights },
        MarkdownViewMode::Split => MarkdownViewSnapshot::Split {
            lines,
            highlights,
            preview: markdown_to_model(doc, doc.source_path.as_deref()),
        },
    }
}
pub fn highlight_markdown_source(source: &str) -> Vec<MarkdownSourceHighlightSpan> {
    let mut spans = Vec::new();
    collect_spans(source, r"(?m)^#{1,6}\s+.+$", MarkdownSourceTokenKind::Heading, &mut spans);
    collect_spans(source, r"(?m)^```.*$", MarkdownSourceTokenKind::CodeFence, &mut spans);
    collect_spans(source, r"\[[^\]]+\]\([^)]+\)", MarkdownSourceTokenKind::Link, &mut spans);
    collect_spans(source, r"(\*\*[^*]+\*\*|\*[^*]+\*|~~[^~]+~~)", MarkdownSourceTokenKind::Emphasis, &mut spans);
    collect_spans(source, r"(?m)^\s*>\s?.*$", MarkdownSourceTokenKind::Quote, &mut spans);
    collect_spans(source, r"(?m)^\s*([-*+]|\d+\.)\s+", MarkdownSourceTokenKind::ListMarker, &mut spans);
    collect_spans(source, r"(?m)^(\*\s*\*\s*\*|-{3,}|_{3,})\s*$", MarkdownSourceTokenKind::Rule, &mut spans);
    collect_spans(source, r"(?m)^\|.*\|$", MarkdownSourceTokenKind::TableRow, &mut spans);
    spans
}

pub fn highlight_code_block(language: Option<&str>, code: &str) -> Vec<MarkdownCodeHighlightToken> {
    let mut tokens = Vec::new();
    let lang = language.unwrap_or("").to_ascii_lowercase();
    let comment_pattern = if matches!(lang.as_str(), "python" | "yaml" | "bash" | "shell" | "sh") {
        r"(?m)#.*$"
    } else if lang == "html" {
        r"(?is)<!--.*?-->"
    } else if lang == "sql" {
        r"(?m)--.*$"
    } else {
        r"(?m)//.*$"
    };
    collect_code_tokens(code, comment_pattern, MarkdownCodeTokenKind::Comment, &mut tokens);
    collect_code_tokens(code, r#""([^"\\]|\\.)*"|'([^'\\]|\\.)*'"#, MarkdownCodeTokenKind::String, &mut tokens);
    collect_code_tokens(code, r"\b\d+(\.\d+)?\b", MarkdownCodeTokenKind::Number, &mut tokens);

    let keywords = language_keywords(lang.as_str());
    let word_re = Regex::new(r"\b[A-Za-z_][A-Za-z0-9_]*\b").expect("regex");
    for m in word_re.find_iter(code) {
        if keywords.contains(m.as_str()) {
            tokens.push(MarkdownCodeHighlightToken {
                start: m.start(),
                end: m.end(),
                kind: MarkdownCodeTokenKind::Keyword,
            });
        }
    }

    tokens.sort_by_key(|t| (t.start, t.end));
    tokens
}

fn heading_level_to_u8(level: HeadingLevel) -> u8 {
    match level {
        HeadingLevel::H1 => 1,
        HeadingLevel::H2 => 2,
        HeadingLevel::H3 => 3,
        HeadingLevel::H4 => 4,
        HeadingLevel::H5 => 5,
        HeadingLevel::H6 => 6,
    }
}

fn extract_runs_from_block(block: &Block) -> Vec<Run> {
    match block {
        Block::Paragraph(p) => p.runs.clone(),
        Block::Heading(h) => h.runs.clone(),
        _ => vec![],
    }
}

fn styled_run(
    text: &str,
    bold: bool,
    italic: bool,
    strikethrough: bool,
    superscript: bool,
    subscript: bool,
    link: Option<&str>,
    code_inline: bool,
) -> Run {
    let mut style = RunStyle {
        bold,
        italic,
        strikethrough,
        superscript,
        subscript,
        ..RunStyle::default()
    };
    if code_inline {
        style.font_family = Some("Cascadia Mono".to_string());
        style.background = Some(Color::rgba(0.13, 0.18, 0.26, 0.8));
    }
    if link.is_some() {
        style.underline = true;
        style.color = Some(Color::rgb(0.34, 0.55, 0.95));
    }
    Run {
        text: text.to_string(),
        style,
    }
}

fn build_image_block(source: &str, alt_text: &str, base_path: Option<&Path>, id: BlockId) -> ImageBlock {
    let resolved = resolve_local_image_path(source, base_path);
    let (width, height) = load_image_dimensions(source, base_path).unwrap_or((320, 180));
    ImageBlock {
        id,
        data: resolved
            .as_ref()
            .map(|p| ImageDataRef::LinkedPath(p.clone()))
            .unwrap_or(ImageDataRef::Empty),
        original_width: width,
        original_height: height,
        caption: None,
        border: None,
        crop: None,
        key: source.to_string(),
        alt_text: alt_text.to_string(),
        source_path: resolved,
        width: width as f32,
        height: height as f32,
        alignment: ImageAlignment::Inline,
    }
}

fn collect_spans(source: &str, pattern: &str, kind: MarkdownSourceTokenKind, spans: &mut Vec<MarkdownSourceHighlightSpan>) {
    let re = Regex::new(pattern).expect("regex");
    for m in re.find_iter(source) {
        spans.push(MarkdownSourceHighlightSpan {
            start: m.start(),
            end: m.end(),
            kind: kind.clone(),
        });
    }
}

fn collect_code_tokens(code: &str, pattern: &str, kind: MarkdownCodeTokenKind, tokens: &mut Vec<MarkdownCodeHighlightToken>) {
    let re = Regex::new(pattern).expect("regex");
    for m in re.find_iter(code) {
        tokens.push(MarkdownCodeHighlightToken {
            start: m.start(),
            end: m.end(),
            kind: kind.clone(),
        });
    }
}

fn language_keywords(language: &str) -> HashSet<&'static str> {
    match language {
        "rust" => ["fn", "let", "mut", "pub", "impl", "struct", "enum", "trait", "match", "if", "else", "for", "while", "return"].into_iter().collect(),
        "python" => ["def", "class", "import", "from", "if", "elif", "else", "for", "while", "return", "try", "except", "with"].into_iter().collect(),
        "javascript" | "typescript" => ["function", "const", "let", "var", "if", "else", "for", "while", "return", "class", "import", "export", "async", "await"].into_iter().collect(),
        "c" | "cpp" | "c++" => ["int", "void", "char", "float", "double", "if", "else", "for", "while", "return", "struct", "class", "const"].into_iter().collect(),
        "java" => ["class", "public", "private", "protected", "static", "void", "if", "else", "for", "while", "return", "new"].into_iter().collect(),
        "go" => ["func", "package", "import", "if", "else", "for", "range", "return", "type", "struct", "interface", "go", "defer"].into_iter().collect(),
        "html" => ["html", "head", "body", "div", "span", "script", "style"].into_iter().collect(),
        "css" => ["display", "color", "background", "border", "position", "flex", "grid"].into_iter().collect(),
        "sql" => ["select", "from", "where", "join", "left", "right", "insert", "update", "delete", "create", "table", "order", "group", "by"].into_iter().collect(),
        _ => HashSet::new(),
    }
}

fn is_remote_source(source: &str) -> bool {
    source.starts_with("http://") || source.starts_with("https://")
}

fn resolve_local_image_path(source: &str, base_path: Option<&Path>) -> Option<PathBuf> {
    if is_remote_source(source) {
        return None;
    }
    let path = Path::new(source);
    if path.is_absolute() {
        Some(path.to_path_buf())
    } else {
        base_path.map(|base| base.join(path))
    }
}

fn load_image_dimensions(source: &str, base_path: Option<&Path>) -> Result<(u32, u32), String> {
    if is_remote_source(source) {
        return Err("remote image download deferred".to_string());
    }
    let resolved = resolve_local_image_path(source, base_path).ok_or_else(|| "cannot resolve path".to_string())?;
    let bytes = fs::read(&resolved).map_err(|e| format!("read image failed: {e}"))?;
    let image = image::load_from_memory(&bytes).map_err(|e| format!("decode image failed: {e}"))?;
    Ok(image.dimensions())
}
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_core_markdown_features() {
        let md = r#"
# Title

Paragraph with **bold**, *italic*, ~~strike~~, `code`, [link](https://example.com) and ![alt](missing.png)

- [x] done

> quote

| A | B |
|---|---|
| 1 | 2 |

```rust
fn main() {}
```

---
"#;
        let doc = MarkdownDocument::from_source(md);
        let model = markdown_to_model(&doc, None);

        assert!(model.content.iter().any(|b| matches!(b, Block::Heading(_))));
        assert!(model.content.iter().any(|b| matches!(b, Block::List(_))));
        assert!(model.content.iter().any(|b| matches!(b, Block::BlockQuote(_))));
        assert!(model.content.iter().any(|b| matches!(b, Block::Table(_))));
        assert!(model.content.iter().any(|b| matches!(b, Block::CodeBlock(_))));
        assert!(model.content.iter().any(|b| matches!(b, Block::HorizontalRule)));
        assert!(model.content.iter().any(|b| matches!(b, Block::Image(_))));
    }

    #[test]
    fn builds_outline_and_highlights() {
        let doc = MarkdownDocument::from_source("# H1\n## H2\n- item\n[link](a)");
        let outline = collect_outline(&doc);
        assert_eq!(outline.len(), 2);

        let spans = highlight_markdown_source(doc.source());
        assert!(spans.iter().any(|s| s.kind == MarkdownSourceTokenKind::Heading));
        assert!(spans.iter().any(|s| s.kind == MarkdownSourceTokenKind::ListMarker));
        assert!(spans.iter().any(|s| s.kind == MarkdownSourceTokenKind::Link));
    }

    #[test]
    fn highlights_code_tokens() {
        let tokens = highlight_code_block(Some("rust"), r#"fn run() { let x = 42; println!("ok"); }"#);
        assert!(tokens.iter().any(|t| t.kind == MarkdownCodeTokenKind::Keyword));
        assert!(tokens.iter().any(|t| t.kind == MarkdownCodeTokenKind::Number));
        assert!(tokens.iter().any(|t| t.kind == MarkdownCodeTokenKind::String));
    }

    #[test]
    fn image_loader_is_async() {
        let temp_dir = std::env::temp_dir().join("doco-md-loader-tests");
        let _ = fs::create_dir_all(&temp_dir);
        let img_path = temp_dir.join("sample.png");
        image::DynamicImage::new_rgb8(2, 3)
            .save(&img_path)
            .expect("save image");

        let mut loader = MarkdownImageLoader::new(Some(temp_dir));
        let placeholder = loader.request_image("sample.png", "sample");
        assert_eq!(placeholder.status, MarkdownImageLoadStatus::Placeholder);

        for _ in 0..80 {
            if loader.poll() > 0 {
                break;
            }
            std::thread::sleep(std::time::Duration::from_millis(5));
        }

        let loaded = loader.image("sample.png").expect("entry exists");
        assert_eq!(loaded.status, MarkdownImageLoadStatus::Ready);
        assert_eq!((loaded.width, loaded.height), (2, 3));
    }
}
