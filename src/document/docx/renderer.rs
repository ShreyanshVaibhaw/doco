use std::{
    collections::HashMap,
    hash::{Hash, Hasher},
};

use crate::{
    document::model::{
        Block, DocumentModel, ListItem, PageSize, Paragraph, ParagraphAlignment, Run, RunStyle, Table,
        TableCell,
    },
    ui::{Color, Rect},
};

#[derive(Debug, Clone)]
pub struct RenderConfig {
    pub page_gap: f32,
    pub default_font_size: f32,
    pub margin: f32,
    pub widow_orphan_lines: usize,
    pub header_height: f32,
    pub footer_height: f32,
    pub different_first_page_header_footer: bool,
    pub viewport: Option<Rect>,
}

impl Default for RenderConfig {
    fn default() -> Self {
        Self {
            page_gap: 24.0,
            default_font_size: 12.0,
            margin: 72.0,
            widow_orphan_lines: 2,
            header_height: 18.0,
            footer_height: 18.0,
            different_first_page_header_footer: false,
            viewport: None,
        }
    }
}

#[derive(Debug, Clone)]
pub struct HyperlinkHitRegion {
    pub rect: Rect,
    pub text: String,
    pub target: Option<String>,
}

#[derive(Debug, Clone)]
pub struct LaidOutPage {
    pub index: usize,
    pub bounds: Rect,
    pub content_bounds: Rect,
    pub blocks: Vec<LaidOutBlock>,
    pub header_draw: Vec<DrawCommand>,
    pub footer_draw: Vec<DrawCommand>,
    pub link_regions: Vec<HyperlinkHitRegion>,
    pub from_cache: bool,
}

#[derive(Debug, Clone)]
pub struct LaidOutBlock {
    pub block_index: usize,
    pub rect: Rect,
    pub draw: Vec<DrawCommand>,
}

#[derive(Debug, Clone)]
pub enum DrawCommand {
    Text {
        text: String,
        rect: Rect,
        size: f32,
        bold: bool,
        italic: bool,
        underline: bool,
        color: Option<Color>,
        alignment: ParagraphAlignment,
        format_id: u32,
    },
    Line {
        from: (f32, f32),
        to: (f32, f32),
        width: f32,
        color: Color,
    },
    Rect {
        rect: Rect,
        fill: Option<Color>,
        stroke: Option<Color>,
    },
    Image {
        key: String,
        rect: Rect,
    },
}

#[derive(Debug, Clone)]
struct BlockLayout {
    height: f32,
    draw: Vec<DrawCommand>,
    links: Vec<HyperlinkHitRegion>,
    line_height: f32,
    line_count: usize,
}

impl Default for BlockLayout {
    fn default() -> Self {
        Self {
            height: 0.0,
            draw: Vec::new(),
            links: Vec::new(),
            line_height: 0.0,
            line_count: 0,
        }
    }
}

#[derive(Debug, Default)]
pub struct DocxRenderEngine {
    // In a real DirectWrite renderer this maps style keys to IDWriteTextFormat handles.
    text_style_cache: HashMap<String, u32>,
    next_text_style_id: u32,
    page_cache: HashMap<u64, Vec<LaidOutBlock>>,
}

impl DocxRenderEngine {
    pub fn paginate(&mut self, doc: &DocumentModel, cfg: &RenderConfig) -> Vec<LaidOutPage> {
        let (page_w, page_h) = page_size(doc);
        let (top, right, bottom, left) = effective_margins(doc, cfg);
        let content_bounds = Rect {
            x: left,
            y: top + cfg.header_height,
            width: (page_w - left - right).max(80.0),
            height: (page_h - top - bottom - cfg.header_height - cfg.footer_height).max(80.0),
        };

        let mut pages = vec![new_page(0, page_w, page_h, content_bounds)];
        let mut cursor_y = 0.0_f32;

        for (block_index, block) in doc.content.iter().enumerate() {
            if matches!(block, Block::PageBreak) {
                push_new_page(&mut pages, page_w, page_h, cfg.page_gap, content_bounds);
                cursor_y = 0.0;
                continue;
            }

            if let Block::Table(table) = block {
                self.paginate_table(table, block_index, cfg, &mut pages, &mut cursor_y, page_w, page_h);
                continue;
            }

            let layout = self.layout_block(block, doc, cfg, content_bounds.width);
            if layout.height <= 0.0 {
                continue;
            }

            let remaining = content_bounds.height - cursor_y;
            let widow_height = layout.line_height * cfg.widow_orphan_lines as f32;
            if cursor_y > 0.0
                && (cursor_y + layout.height > content_bounds.height
                    || (layout.line_count >= cfg.widow_orphan_lines.saturating_mul(2)
                        && remaining < widow_height))
            {
                push_new_page(&mut pages, page_w, page_h, cfg.page_gap, content_bounds);
                cursor_y = 0.0;
            }

            let page = pages.last_mut().expect("page exists");
            let rect = Rect {
                x: page.bounds.x + page.content_bounds.x,
                y: page.bounds.y + page.content_bounds.y + cursor_y,
                width: page.content_bounds.width,
                height: layout.height,
            };
            page.link_regions.extend(offset_links(&layout.links, rect.x, rect.y));
            page.blocks.push(LaidOutBlock {
                block_index,
                rect,
                draw: offset_commands(&layout.draw, rect.x, rect.y),
            });
            cursor_y += layout.height;
        }

        let page_count = pages.len();
        for page in &mut pages {
            self.render_header_footer(page, doc, cfg, page_count);
        }

        if let Some(viewport) = cfg.viewport {
            clip_to_viewport(&mut pages, viewport);
        }

        let doc_sig = document_signature(doc, cfg);
        for page in &mut pages {
            let sig = page_signature(page, doc_sig);
            if let Some(cached) = self.page_cache.get(&sig) {
                page.blocks = cached.clone();
                page.from_cache = true;
            } else {
                self.page_cache.insert(sig, page.blocks.clone());
                page.from_cache = false;
            }
        }

        pages
    }

    #[allow(clippy::too_many_arguments)]
    fn paginate_table(
        &mut self,
        table: &Table,
        block_index: usize,
        cfg: &RenderConfig,
        pages: &mut Vec<LaidOutPage>,
        cursor_y: &mut f32,
        page_w: f32,
        page_h: f32,
    ) {
        let page_content_h = pages.last().expect("page exists").content_bounds.height;
        let row_heights = table_row_heights(table, cfg, pages.last().expect("page exists").content_bounds.width);
        if row_heights.is_empty() {
            return;
        }

        let mut row_start = 0usize;
        while row_start < table.rows.len() {
            if *cursor_y > 0.0 && *cursor_y + row_heights[row_start] > page_content_h {
                let content_bounds = pages.last().expect("page exists").content_bounds;
                push_new_page(pages, page_w, page_h, cfg.page_gap, content_bounds);
                *cursor_y = 0.0;
            }

            let mut take = 0usize;
            let mut sum_h = 0.0;
            while row_start + take < row_heights.len() {
                let next = row_heights[row_start + take];
                if *cursor_y + sum_h + next > page_content_h && take > 0 {
                    break;
                }
                take += 1;
                sum_h += next;
                if *cursor_y + sum_h > page_content_h {
                    break;
                }
            }
            if take == 0 {
                take = 1;
            }

            let layout = self.layout_table_slice(
                table,
                row_start,
                take,
                &row_heights,
                cfg,
                pages.last().expect("page exists").content_bounds.width,
            );
            let page = pages.last_mut().expect("page exists");
            let rect = Rect {
                x: page.bounds.x + page.content_bounds.x,
                y: page.bounds.y + page.content_bounds.y + *cursor_y,
                width: page.content_bounds.width,
                height: layout.height,
            };
            page.blocks.push(LaidOutBlock {
                block_index,
                rect,
                draw: offset_commands(&layout.draw, rect.x, rect.y),
            });

            *cursor_y += layout.height;
            row_start += take;
            if row_start < table.rows.len() {
                let content_bounds = page.content_bounds;
                push_new_page(pages, page_w, page_h, cfg.page_gap, content_bounds);
                *cursor_y = 0.0;
            }
        }
    }

    fn layout_block(
        &mut self,
        block: &Block,
        doc: &DocumentModel,
        cfg: &RenderConfig,
        width: f32,
    ) -> BlockLayout {
        match block {
            Block::Paragraph(p) => self.layout_paragraph(p, cfg, width, 8.0),
            Block::Heading(h) => {
                let fake = Paragraph {
                    id: h.id,
                    runs: h.runs.clone(),
                    alignment: ParagraphAlignment::Left,
                    spacing: crate::document::model::ParagraphSpacing {
                        before: 10.0,
                        after: 8.0,
                        line: 0.0,
                    },
                    indent: crate::document::model::Indent::default(),
                    style_id: Some(format!("Heading{}", h.level)),
                };
                self.layout_paragraph(&fake, cfg, width, 14.0 + ((6 - h.level.min(6)) as f32))
            }
            Block::HorizontalRule => BlockLayout {
                height: 18.0,
                draw: vec![DrawCommand::Line {
                    from: (0.0, 9.0),
                    to: (width, 9.0),
                    width: 1.0,
                    color: Color::rgb(0.6, 0.6, 0.6),
                }],
                links: Vec::new(),
                line_height: 9.0,
                line_count: 1,
            },
            Block::Image(img) => {
                let (w, h) = resolve_image_size(img.width, img.height, img.key.as_str(), doc, width);
                BlockLayout {
                    height: h + 8.0,
                    draw: vec![DrawCommand::Image {
                        key: img.key.clone(),
                        rect: Rect {
                            x: 0.0,
                            y: 0.0,
                            width: w,
                            height: h,
                        },
                    }],
                    links: Vec::new(),
                    line_height: h,
                    line_count: 1,
                }
            }
            Block::Table(table) => {
                let heights = table_row_heights(table, cfg, width);
                self.layout_table_slice(table, 0, table.rows.len(), &heights, cfg, width)
            }
            Block::List(list) => {
                self.layout_list(list, cfg, width)
            }
            Block::PageBreak => BlockLayout::default(),
            Block::BlockQuote(q) => {
                let txt = flatten_blocks_text(&q.blocks);
                let lines = estimate_wrap_lines(&txt, (width - 12.0).max(40.0), cfg.default_font_size);
                let line_h = cfg.default_font_size * 1.35;
                let format_id = self.resolve_text_format("blockquote");
                BlockLayout {
                    height: lines as f32 * line_h + 8.0,
                    draw: vec![
                        DrawCommand::Line {
                            from: (0.0, 0.0),
                            to: (0.0, lines as f32 * line_h),
                            width: 2.0,
                            color: Color::rgb(0.55, 0.6, 0.7),
                        },
                        DrawCommand::Text {
                            text: txt,
                            rect: Rect {
                                x: 8.0,
                                y: 0.0,
                                width: (width - 8.0).max(30.0),
                                height: lines as f32 * line_h,
                            },
                            size: cfg.default_font_size,
                            bold: false,
                            italic: true,
                            underline: false,
                            color: Some(Color::rgb(0.75, 0.78, 0.86)),
                            alignment: ParagraphAlignment::Left,
                            format_id,
                        },
                    ],
                    links: Vec::new(),
                    line_height: line_h,
                    line_count: lines.max(1),
                }
            }
            Block::CodeBlock(code) => {
                let lines = code.code.lines().count().max(1);
                let line_h = cfg.default_font_size + 3.0;
                let format_id = self.resolve_text_format("code");
                BlockLayout {
                    height: lines as f32 * line_h + 12.0,
                    draw: vec![
                        DrawCommand::Rect {
                            rect: Rect {
                                x: 0.0,
                                y: 0.0,
                                width,
                                height: lines as f32 * line_h + 12.0,
                            },
                            fill: Some(Color::rgb(0.14, 0.14, 0.16)),
                            stroke: Some(Color::rgb(0.24, 0.24, 0.28)),
                        },
                        DrawCommand::Text {
                            text: code.code.clone(),
                            rect: Rect {
                                x: 8.0,
                                y: 6.0,
                                width: (width - 16.0).max(30.0),
                                height: lines as f32 * line_h,
                            },
                            size: cfg.default_font_size,
                            bold: false,
                            italic: false,
                            underline: false,
                            color: Some(Color::rgb(0.88, 0.89, 0.93)),
                            alignment: ParagraphAlignment::Left,
                            format_id,
                        },
                    ],
                    links: Vec::new(),
                    line_height: line_h,
                    line_count: lines,
                }
            }
        }
    }

    fn layout_paragraph(
        &mut self,
        paragraph: &Paragraph,
        cfg: &RenderConfig,
        width: f32,
        fallback_size: f32,
    ) -> BlockLayout {
        let mut draw = Vec::new();
        let mut links = Vec::new();
        let mut x = paragraph.indent.left + paragraph.indent.first_line;
        let mut y = paragraph.spacing.before.max(0.0);
        let max_width = (width - paragraph.indent.left - paragraph.indent.right).max(30.0);
        let mut line_count = 1usize;
        let mut max_line_height = (cfg.default_font_size.max(fallback_size) * 1.35).max(8.0);

        for run in merge_similar_runs(&paragraph.runs) {
            let size = run.style.font_size.unwrap_or(fallback_size).max(8.0);
            let line_h = size * 1.35;
            max_line_height = max_line_height.max(line_h);
            for (idx, part) in run.text.split('\n').enumerate() {
                if idx > 0 {
                    x = paragraph.indent.left;
                    y += max_line_height;
                    line_count += 1;
                }
                if part.is_empty() {
                    continue;
                }

                let token_w = estimate_text_width(part, size).max(2.0);
                if x > paragraph.indent.left && (x - paragraph.indent.left + token_w) > max_width {
                    x = paragraph.indent.left;
                    y += max_line_height;
                    line_count += 1;
                }

                let rect = Rect {
                    x,
                    y,
                    width: token_w.min(max_width),
                    height: line_h,
                };
                let format_id = self.resolve_text_format(style_key(&run, paragraph).as_str());
                draw.push(DrawCommand::Text {
                    text: part.to_string(),
                    rect,
                    size,
                    bold: run.style.bold,
                    italic: run.style.italic,
                    underline: run.style.underline,
                    color: run.style.color,
                    alignment: paragraph.alignment.clone(),
                    format_id,
                });

                let link_target = extract_link_target(part);
                if link_target.is_some() || is_hyperlink_style(&run.style) {
                    links.push(HyperlinkHitRegion {
                        rect,
                        text: part.to_string(),
                        target: link_target,
                    });
                }
                x += token_w;
            }
        }

        BlockLayout {
            height: y + max_line_height + paragraph.spacing.after.max(cfg.default_font_size * 0.25),
            draw: batch_text_draw_commands(draw),
            links,
            line_height: max_line_height,
            line_count,
        }
    }

    fn layout_table_slice(
        &mut self,
        table: &Table,
        row_start: usize,
        row_count: usize,
        row_heights: &[f32],
        cfg: &RenderConfig,
        width: f32,
    ) -> BlockLayout {
        if row_count == 0 {
            return BlockLayout::default();
        }
        let col_widths = resolve_column_widths(table, width);
        let mut draw = Vec::new();
        let mut y = 0.0_f32;

        for (offset, row) in table.rows[row_start..row_start + row_count].iter().enumerate() {
            let idx = row_start + offset;
            let row_h = row_heights[idx];
            if table.header_row && idx == 0 {
                draw.push(DrawCommand::Rect {
                    rect: Rect {
                        x: 0.0,
                        y,
                        width,
                        height: row_h,
                    },
                    fill: Some(Color::rgb(0.84, 0.9, 0.98)),
                    stroke: None,
                });
            } else if table.alternating_rows && idx % 2 == 1 {
                draw.push(DrawCommand::Rect {
                    rect: Rect {
                        x: 0.0,
                        y,
                        width,
                        height: row_h,
                    },
                    fill: Some(Color::rgba(0.7, 0.74, 0.8, 0.14)),
                    stroke: None,
                });
            }
            let mut x = 0.0;
            for (col_idx, col_w) in col_widths.iter().enumerate() {
                let cell = row.cells.get(col_idx).cloned().unwrap_or_default();
                draw.push(DrawCommand::Rect {
                    rect: Rect {
                        x,
                        y,
                        width: *col_w,
                        height: row_h,
                    },
                    fill: cell.background,
                    stroke: Some(table.borders.inner_vertical.color),
                });
                let txt = flatten_blocks_text(&cell.blocks);
                if !txt.trim().is_empty() {
                    draw.push(DrawCommand::Text {
                        text: txt,
                        rect: Rect {
                            x: x + table.cell_padding,
                            y: y + table.cell_padding,
                            width: (*col_w - table.cell_padding * 2.0).max(10.0),
                            height: (row_h - table.cell_padding * 2.0).max(8.0),
                        },
                        size: cfg.default_font_size,
                        bold: false,
                        italic: false,
                        underline: false,
                        color: None,
                        alignment: ParagraphAlignment::Left,
                        format_id: self.resolve_text_format("table"),
                    });
                }
                x += *col_w;
            }
            y += row_h;
        }

        draw.push(DrawCommand::Rect {
            rect: Rect {
                x: 0.0,
                y: 0.0,
                width,
                height: y.max(1.0),
            },
            fill: None,
            stroke: Some(table.borders.outer.color),
        });

        BlockLayout {
            height: y + 8.0,
            draw,
            links: Vec::new(),
            line_height: cfg.default_font_size * 1.3,
            line_count: row_count,
        }
    }

    fn layout_list(&mut self, list: &crate::document::model::List, cfg: &RenderConfig, width: f32) -> BlockLayout {
        let mut draw = Vec::new();
        let mut y = 0.0_f32;
        for (i, item) in list.items.iter().enumerate() {
            let bullet = match list.list_type {
                crate::document::model::ListType::Bullet => "•".to_string(),
                crate::document::model::ListType::Numbered => format!("{}.", list.start_number + i as u32),
                crate::document::model::ListType::Checkbox => {
                    if item.checked.unwrap_or(false) {
                        "[x]".to_string()
                    } else {
                        "[ ]".to_string()
                    }
                }
            };
            let txt = list_item_text(item);
            let format_id = self.resolve_text_format("list-item");
            draw.push(DrawCommand::Text {
                text: bullet,
                rect: Rect {
                    x: 0.0,
                    y,
                    width: 20.0,
                    height: 20.0,
                },
                size: cfg.default_font_size,
                bold: false,
                italic: false,
                underline: false,
                color: None,
                alignment: ParagraphAlignment::Left,
                format_id,
            });
            draw.push(DrawCommand::Text {
                text: txt,
                rect: Rect {
                    x: 20.0,
                    y,
                    width: (width - 20.0).max(20.0),
                    height: 20.0,
                },
                size: cfg.default_font_size,
                bold: false,
                italic: false,
                underline: false,
                color: None,
                alignment: ParagraphAlignment::Left,
                format_id,
            });
            y += 24.0;
        }
        BlockLayout {
            height: y + 4.0,
            draw,
            links: Vec::new(),
            line_height: cfg.default_font_size * 1.35,
            line_count: list.items.len().max(1),
        }
    }

    fn render_header_footer(
        &mut self,
        page: &mut LaidOutPage,
        doc: &DocumentModel,
        cfg: &RenderConfig,
        total_pages: usize,
    ) {
        if cfg.different_first_page_header_footer && page.index == 0 {
            return;
        }
        let title = if doc.metadata.title.is_empty() {
            "Doco".to_string()
        } else {
            doc.metadata.title.clone()
        };
        let hdr_id = self.resolve_text_format("header");
        let ftr_id = self.resolve_text_format("footer");
        if cfg.header_height > 0.0 {
            page.header_draw.push(DrawCommand::Text {
                text: title,
                rect: Rect {
                    x: page.bounds.x + page.content_bounds.x,
                    y: page.bounds.y + 4.0,
                    width: page.content_bounds.width,
                    height: cfg.header_height.max(10.0),
                },
                size: 9.5,
                bold: false,
                italic: false,
                underline: false,
                color: Some(Color::rgb(0.62, 0.66, 0.74)),
                alignment: ParagraphAlignment::Left,
                format_id: hdr_id,
            });
        }
        if cfg.footer_height > 0.0 {
            page.footer_draw.push(DrawCommand::Text {
                text: format!("Page {} / {}", page.index + 1, total_pages.max(1)),
                rect: Rect {
                    x: page.bounds.x + page.content_bounds.x,
                    y: page.bounds.y + page.bounds.height - cfg.footer_height - 4.0,
                    width: page.content_bounds.width,
                    height: cfg.footer_height.max(10.0),
                },
                size: 9.5,
                bold: false,
                italic: false,
                underline: false,
                color: Some(Color::rgb(0.62, 0.66, 0.74)),
                alignment: ParagraphAlignment::Center,
                format_id: ftr_id,
            });
        }
    }

    fn resolve_text_format(&mut self, key: &str) -> u32 {
        if let Some(id) = self.text_style_cache.get(key) {
            return *id;
        }
        self.next_text_style_id = self.next_text_style_id.saturating_add(1);
        self.text_style_cache.insert(key.to_string(), self.next_text_style_id);
        self.next_text_style_id
    }
}

pub fn draw_docx(model: &DocumentModel) -> Vec<LaidOutPage> {
    let mut engine = DocxRenderEngine::default();
    engine.paginate(model, &RenderConfig::default())
}

fn style_key(run: &Run, paragraph: &Paragraph) -> String {
    format!(
        "{}|{}|{}|{}|{}|{:?}|{:?}",
        run.style.font_family.clone().unwrap_or_else(|| "Segoe UI".to_string()),
        run.style.font_size.unwrap_or(12.0),
        run.style.bold,
        run.style.italic,
        run.style.underline,
        run.style.color,
        paragraph.style_id
    )
}

fn effective_margins(doc: &DocumentModel, cfg: &RenderConfig) -> (f32, f32, f32, f32) {
    let fallback = cfg.margin.max(8.0);
    let top = if doc.metadata.margins.top > 0.0 {
        doc.metadata.margins.top
    } else {
        fallback
    };
    let right = if doc.metadata.margins.right > 0.0 {
        doc.metadata.margins.right
    } else {
        fallback
    };
    let bottom = if doc.metadata.margins.bottom > 0.0 {
        doc.metadata.margins.bottom
    } else {
        fallback
    };
    let left = if doc.metadata.margins.left > 0.0 {
        doc.metadata.margins.left
    } else {
        fallback
    };
    (top, right, bottom, left)
}

fn estimate_wrap_lines(text: &str, width: f32, font_size: f32) -> usize {
    let avg_char_w = (font_size * 0.52).max(4.0);
    let max_chars = (width / avg_char_w).max(8.0) as usize;
    text.split('\n')
        .map(|line| line.chars().count().max(1).div_ceil(max_chars))
        .sum::<usize>()
        .max(1)
}

fn page_size(doc: &DocumentModel) -> (f32, f32) {
    match doc.metadata.page_size {
        PageSize::Letter => (612.0, 792.0),
        PageSize::A4 => (595.0, 842.0),
        PageSize::Legal => (612.0, 1008.0),
        PageSize::Custom {
            width_points,
            height_points,
        } => (width_points.max(200.0), height_points.max(200.0)),
    }
}

fn estimate_text_width(text: &str, size: f32) -> f32 {
    text.chars().count().max(1) as f32 * (size.max(8.0) * 0.52)
}

fn resolve_image_size(
    width: f32,
    height: f32,
    key: &str,
    doc: &DocumentModel,
    available_width: f32,
) -> (f32, f32) {
    let mut src_w = 0.0;
    let mut src_h = 0.0;
    if let Some(image) = doc.images.get(key) {
        src_w = image.width as f32 * 0.75;
        src_h = image.height as f32 * 0.75;
    }

    let mut draw_w = if width > 0.0 { width } else { src_w.max(160.0) };
    let mut draw_h = if height > 0.0 { height } else { src_h.max(120.0) };
    if src_w > 0.0 && src_h > 0.0 && draw_h > 0.0 {
        let ratio = src_w / src_h;
        if width > 0.0 && height <= 0.0 {
            draw_h = draw_w / ratio;
        } else if height > 0.0 && width <= 0.0 {
            draw_w = draw_h * ratio;
        }
    }
    if draw_w > available_width {
        let scale = available_width / draw_w;
        draw_w *= scale;
        draw_h *= scale;
    }
    (draw_w.max(80.0), draw_h.max(60.0))
}

fn table_row_heights(table: &Table, cfg: &RenderConfig, width: f32) -> Vec<f32> {
    let col_widths = resolve_column_widths(table, width);
    table
        .rows
        .iter()
        .map(|row| {
            let mut row_h = (20.0 + table.cell_padding * 2.0).max(20.0);
            for (col, cell) in row.cells.iter().enumerate() {
                row_h = row_h.max(estimate_cell_height(
                    cell,
                    table.cell_padding,
                    cfg.default_font_size,
                    *col_widths.get(col).unwrap_or(&32.0),
                ));
            }
            row_h
        })
        .collect()
}

fn resolve_column_widths(table: &Table, width: f32) -> Vec<f32> {
    let mut columns = if !table.column_widths.is_empty() {
        table.column_widths.clone()
    } else {
        vec![1.0; table.rows.iter().map(|r| r.cells.len()).max().unwrap_or(1).max(1)]
    };
    let total = columns.iter().copied().sum::<f32>().max(1.0);
    for col in &mut columns {
        *col = width * (*col / total);
    }
    columns
}

fn estimate_cell_height(cell: &TableCell, padding: f32, font_size: f32, width: f32) -> f32 {
    let text = flatten_blocks_text(&cell.blocks);
    let lines = estimate_wrap_lines(&text, (width - padding * 2.0).max(20.0), font_size);
    lines as f32 * (font_size * 1.3) + padding.max(3.0) * 2.0
}

fn flatten_blocks_text(blocks: &[Block]) -> String {
    let mut out = String::new();
    for (idx, block) in blocks.iter().enumerate() {
        if idx > 0 && !out.ends_with('\n') {
            out.push('\n');
        }
        match block {
            Block::Paragraph(p) => out.push_str(&p.runs.iter().map(|r| r.text.as_str()).collect::<String>()),
            Block::Heading(h) => out.push_str(&h.runs.iter().map(|r| r.text.as_str()).collect::<String>()),
            Block::CodeBlock(c) => out.push_str(c.code.as_str()),
            Block::BlockQuote(q) => out.push_str(flatten_blocks_text(&q.blocks).as_str()),
            Block::List(l) => {
                for item in &l.items {
                    let line = list_item_text(item);
                    if !line.is_empty() {
                        out.push_str(line.as_str());
                        out.push('\n');
                    }
                }
            }
            Block::Table(_) => out.push_str("[table]"),
            Block::Image(_) => out.push_str("[image]"),
            Block::PageBreak => out.push('\n'),
            Block::HorizontalRule => out.push_str("---"),
        }
    }
    out
}

fn list_item_text(item: &ListItem) -> String {
    let mut text = flatten_blocks_text(&item.content);
    for child in &item.children {
        let c = list_item_text(child);
        if !c.is_empty() {
            if !text.is_empty() && !text.ends_with('\n') {
                text.push('\n');
            }
            text.push_str(c.as_str());
        }
    }
    text
}

fn merge_similar_runs(runs: &[Run]) -> Vec<Run> {
    let mut merged: Vec<Run> = Vec::new();
    for run in runs {
        if let Some(last) = merged.last_mut() && run_style_eq(&last.style, &run.style) {
            last.text.push_str(run.text.as_str());
            continue;
        }
        merged.push(run.clone());
    }
    merged
}

fn run_style_eq(a: &RunStyle, b: &RunStyle) -> bool {
    a.font_family == b.font_family
        && a.font_size == b.font_size
        && a.bold == b.bold
        && a.italic == b.italic
        && a.underline == b.underline
        && a.strikethrough == b.strikethrough
        && a.color == b.color
        && a.background == b.background
        && a.superscript == b.superscript
        && a.subscript == b.subscript
}

fn is_hyperlink_style(style: &RunStyle) -> bool {
    style.underline && style.color.map(|c| c.b > 0.55 && c.r < 0.35).unwrap_or(false)
}

fn extract_link_target(text: &str) -> Option<String> {
    for token in text.split_whitespace() {
        let c = token.trim_matches(['(', ')', '[', ']', ',', '.', ';', '"']);
        if c.starts_with("http://") || c.starts_with("https://") || c.starts_with("mailto:") {
            return Some(c.to_string());
        }
        if c.starts_with("www.") && c.len() > 4 {
            return Some(format!("https://{c}"));
        }
    }
    None
}

fn batch_text_draw_commands(commands: Vec<DrawCommand>) -> Vec<DrawCommand> {
    let mut out = Vec::new();
    for cmd in commands {
        if let Some(last) = out.last_mut() && merge_if_possible(last, &cmd) {
            continue;
        }
        out.push(cmd);
    }
    out
}

fn merge_if_possible(existing: &mut DrawCommand, incoming: &DrawCommand) -> bool {
    match (existing, incoming) {
        (
            DrawCommand::Text {
                text: lt,
                rect: lr,
                size: ls,
                bold: lb,
                italic: li,
                underline: lu,
                color: lc,
                alignment: la,
                format_id: lf,
            },
            DrawCommand::Text {
                text: rt,
                rect: rr,
                size: rs,
                bold: rb,
                italic: ri,
                underline: ru,
                color: rc,
                alignment: ra,
                format_id: rf,
            },
        ) => {
            let same = *ls == *rs
                && *lb == *rb
                && *li == *ri
                && *lu == *ru
                && *lc == *rc
                && *lf == *rf
                && same_alignment(la, ra)
                && (lr.y - rr.y).abs() < 0.2
                && (lr.x + lr.width - rr.x).abs() < 1.5;
            if same {
                lt.push_str(rt.as_str());
                lr.width = (rr.x + rr.width - lr.x).max(lr.width);
                true
            } else {
                false
            }
        }
        _ => false,
    }
}

fn same_alignment(a: &ParagraphAlignment, b: &ParagraphAlignment) -> bool {
    matches!(
        (a, b),
        (ParagraphAlignment::Left, ParagraphAlignment::Left)
            | (ParagraphAlignment::Center, ParagraphAlignment::Center)
            | (ParagraphAlignment::Right, ParagraphAlignment::Right)
            | (ParagraphAlignment::Justify, ParagraphAlignment::Justify)
    )
}

fn offset_commands(commands: &[DrawCommand], dx: f32, dy: f32) -> Vec<DrawCommand> {
    commands
        .iter()
        .cloned()
        .map(|mut cmd| {
            match &mut cmd {
                DrawCommand::Text { rect, .. } | DrawCommand::Rect { rect, .. } | DrawCommand::Image { rect, .. } => {
                    rect.x += dx;
                    rect.y += dy;
                }
                DrawCommand::Line { from, to, .. } => {
                    from.0 += dx;
                    from.1 += dy;
                    to.0 += dx;
                    to.1 += dy;
                }
            }
            cmd
        })
        .collect()
}

fn offset_links(links: &[HyperlinkHitRegion], dx: f32, dy: f32) -> Vec<HyperlinkHitRegion> {
    links
        .iter()
        .cloned()
        .map(|mut l| {
            l.rect.x += dx;
            l.rect.y += dy;
            l
        })
        .collect()
}

fn clip_to_viewport(pages: &mut [LaidOutPage], viewport: Rect) {
    for page in pages {
        page.blocks.retain(|b| rects_intersect(b.rect, viewport));
        for block in &mut page.blocks {
            block
                .draw
                .retain(|d| command_rect(d).map(|r| rects_intersect(r, viewport)).unwrap_or(true));
        }
        page
            .header_draw
            .retain(|d| command_rect(d).map(|r| rects_intersect(r, viewport)).unwrap_or(true));
        page
            .footer_draw
            .retain(|d| command_rect(d).map(|r| rects_intersect(r, viewport)).unwrap_or(true));
        page.link_regions.retain(|l| rects_intersect(l.rect, viewport));
    }
}

fn command_rect(draw: &DrawCommand) -> Option<Rect> {
    match draw {
        DrawCommand::Text { rect, .. } | DrawCommand::Rect { rect, .. } | DrawCommand::Image { rect, .. } => {
            Some(*rect)
        }
        DrawCommand::Line { from, to, .. } => Some(Rect {
            x: from.0.min(to.0),
            y: from.1.min(to.1),
            width: (from.0 - to.0).abs().max(1.0),
            height: (from.1 - to.1).abs().max(1.0),
        }),
    }
}

fn rects_intersect(a: Rect, b: Rect) -> bool {
    a.x < b.x + b.width && a.x + a.width > b.x && a.y < b.y + b.height && a.y + a.height > b.y
}

fn new_page(index: usize, page_w: f32, page_h: f32, content_bounds: Rect) -> LaidOutPage {
    LaidOutPage {
        index,
        bounds: Rect {
            x: 0.0,
            y: index as f32 * page_h,
            width: page_w,
            height: page_h,
        },
        content_bounds,
        blocks: Vec::new(),
        header_draw: Vec::new(),
        footer_draw: Vec::new(),
        link_regions: Vec::new(),
        from_cache: false,
    }
}

fn push_new_page(
    pages: &mut Vec<LaidOutPage>,
    page_w: f32,
    page_h: f32,
    page_gap: f32,
    content_bounds: Rect,
) {
    let index = pages.len();
    let mut page = new_page(index, page_w, page_h, content_bounds);
    page.bounds.y = index as f32 * (page_h + page_gap);
    pages.push(page);
}

fn document_signature(doc: &DocumentModel, cfg: &RenderConfig) -> u64 {
    let mut h = std::collections::hash_map::DefaultHasher::new();
    doc.content.len().hash(&mut h);
    doc.images.len().hash(&mut h);
    doc.metadata.title.hash(&mut h);
    cfg.default_font_size.to_bits().hash(&mut h);
    cfg.margin.to_bits().hash(&mut h);
    cfg.widow_orphan_lines.hash(&mut h);
    h.finish()
}

fn page_signature(page: &LaidOutPage, doc_sig: u64) -> u64 {
    let mut h = std::collections::hash_map::DefaultHasher::new();
    doc_sig.hash(&mut h);
    page.index.hash(&mut h);
    for b in &page.blocks {
        b.block_index.hash(&mut h);
        b.rect.x.to_bits().hash(&mut h);
        b.rect.y.to_bits().hash(&mut h);
        b.rect.width.to_bits().hash(&mut h);
        b.rect.height.to_bits().hash(&mut h);
    }
    h.finish()
}

#[cfg(test)]
mod tests {
    use crate::document::model::{
        Block, BlockId, Margins, PageSize, Paragraph, ParagraphAlignment, ParagraphSpacing, Run, RunStyle,
        Table, TableCell, TableRow,
    };

    use super::*;

    #[test]
    fn renders_header_footer_and_link_regions() {
        let mut doc = DocumentModel::default();
        doc.metadata.title = "Project Plan".to_string();
        doc.content.push(simple_paragraph(
            1,
            "See https://example.com/spec for details",
            true,
        ));

        let mut engine = DocxRenderEngine::default();
        let pages = engine.paginate(&doc, &RenderConfig::default());
        assert_eq!(pages.len(), 1);
        assert!(!pages[0].header_draw.is_empty());
        assert!(!pages[0].footer_draw.is_empty());
        assert_eq!(pages[0].link_regions.len(), 1);
    }

    #[test]
    fn splits_table_across_pages() {
        let mut doc = DocumentModel::default();
        doc.metadata.page_size = PageSize::Custom {
            width_points: 420.0,
            height_points: 220.0,
        };
        doc.metadata.margins = Margins {
            top: 20.0,
            right: 20.0,
            bottom: 20.0,
            left: 20.0,
        };
        let mut table = Table {
            id: BlockId(2),
            column_widths: vec![1.0, 1.0],
            ..Table::default()
        };
        for i in 0..10 {
            table.rows.push(TableRow {
                cells: vec![
                    TableCell {
                        blocks: vec![simple_paragraph(20 + i as u64, "Row header", false)],
                        ..Default::default()
                    },
                    TableCell {
                        blocks: vec![simple_paragraph(
                            40 + i as u64,
                            "Long value column content that wraps and increases row height.",
                            false,
                        )],
                        ..Default::default()
                    },
                ],
            });
        }
        doc.content.push(Block::Table(table));
        let cfg = RenderConfig {
            header_height: 0.0,
            footer_height: 0.0,
            page_gap: 0.0,
            ..RenderConfig::default()
        };
        let mut engine = DocxRenderEngine::default();
        let pages = engine.paginate(&doc, &cfg);
        assert!(pages.len() >= 2);
    }

    #[test]
    fn honors_page_break_and_viewport_clipping() {
        let mut doc = DocumentModel::default();
        doc.content.push(simple_paragraph(1, "First", false));
        doc.content.push(Block::PageBreak);
        doc.content.push(simple_paragraph(2, "Second", false));
        for i in 0..8 {
            doc.content.push(simple_paragraph(
                10 + i,
                "Extra paragraph for viewport clipping",
                false,
            ));
        }

        let mut engine = DocxRenderEngine::default();
        let cfg = RenderConfig {
            viewport: Some(Rect {
                x: 0.0,
                y: 0.0,
                width: 420.0,
                height: 220.0,
            }),
            ..RenderConfig::default()
        };
        let pages = engine.paginate(&doc, &cfg);
        assert!(pages.len() >= 2);
        let vp = cfg.viewport.expect("viewport");
        for p in &pages {
            for b in &p.blocks {
                assert!(rects_intersect(b.rect, vp));
            }
        }
    }

    fn simple_paragraph(id: u64, text: &str, link_style: bool) -> Block {
        Block::Paragraph(Paragraph {
            id: BlockId(id),
            runs: vec![Run {
                text: text.to_string(),
                style: RunStyle {
                    underline: link_style,
                    color: if link_style {
                        Some(Color::rgb(0.1, 0.32, 0.9))
                    } else {
                        None
                    },
                    ..RunStyle::default()
                },
            }],
            alignment: ParagraphAlignment::Left,
            spacing: ParagraphSpacing {
                before: 0.0,
                after: 8.0,
                line: 0.0,
            },
            indent: Default::default(),
            style_id: None,
        })
    }
}
