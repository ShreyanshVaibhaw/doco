use std::collections::HashMap;

use crate::{
    document::model::{Block, DocumentModel, Paragraph, ParagraphAlignment, Run},
    ui::{Color, Rect},
};

#[derive(Debug, Clone)]
pub struct RenderConfig {
    pub page_gap: f32,
    pub default_font_size: f32,
    pub margin: f32,
}

impl Default for RenderConfig {
    fn default() -> Self {
        Self {
            page_gap: 24.0,
            default_font_size: 12.0,
            margin: 72.0,
        }
    }
}

#[derive(Debug, Clone)]
pub struct LaidOutPage {
    pub index: usize,
    pub bounds: Rect,
    pub blocks: Vec<LaidOutBlock>,
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

#[derive(Debug, Default)]
pub struct DocxRenderEngine {
    // In a real DirectWrite renderer this maps style keys to IDWriteTextFormat.
    text_style_cache: HashMap<String, usize>,
}

impl DocxRenderEngine {
    pub fn paginate(&mut self, doc: &DocumentModel, cfg: &RenderConfig) -> Vec<LaidOutPage> {
        let (page_w, page_h) = page_size(doc);
        let content_w = (page_w - cfg.margin * 2.0).max(100.0);
        let content_h = (page_h - cfg.margin * 2.0).max(100.0);

        let mut pages = vec![LaidOutPage {
            index: 0,
            bounds: Rect {
                x: 0.0,
                y: 0.0,
                width: page_w,
                height: page_h,
            },
            blocks: Vec::new(),
        }];

        let mut cursor_y = cfg.margin;

        for (block_index, block) in doc.content.iter().enumerate() {
            let (height, draw) = self.layout_block(block, cfg, content_w);
            if cursor_y + height > cfg.margin + content_h {
                let next_index = pages.len();
                let page_y = (next_index as f32) * (page_h + cfg.page_gap);
                pages.push(LaidOutPage {
                    index: next_index,
                    bounds: Rect {
                        x: 0.0,
                        y: page_y,
                        width: page_w,
                        height: page_h,
                    },
                    blocks: Vec::new(),
                });
                cursor_y = cfg.margin;
            }

            let page = pages.last_mut().expect("at least one page");
            let y = page.bounds.y + cursor_y;
            page.blocks.push(LaidOutBlock {
                block_index,
                rect: Rect {
                    x: cfg.margin,
                    y,
                    width: content_w,
                    height,
                },
                draw,
            });
            cursor_y += height;
        }

        pages
    }

    fn layout_block(&mut self, block: &Block, cfg: &RenderConfig, width: f32) -> (f32, Vec<DrawCommand>) {
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
            Block::HorizontalRule => {
                let h = 18.0;
                (
                    h,
                    vec![DrawCommand::Line {
                        from: (0.0, h * 0.5),
                        to: (width, h * 0.5),
                        width: 1.0,
                        color: Color::rgb(0.6, 0.6, 0.6),
                    }],
                )
            }
            Block::Image(img) => (
                img.height.max(80.0) + 8.0,
                vec![DrawCommand::Image {
                    key: img.key.clone(),
                    rect: Rect {
                        x: 0.0,
                        y: 0.0,
                        width: img.width.max(120.0),
                        height: img.height.max(80.0),
                    },
                }],
            ),
            Block::Table(table) => {
                let row_h = (20.0 + table.cell_padding * 2.0).max(20.0);
                let row_count = table.rows.len().max(1);
                let h = row_count as f32 * row_h + 8.0;

                let mut draw = vec![DrawCommand::Rect {
                    rect: Rect {
                        x: 0.0,
                        y: 0.0,
                        width,
                        height: h - 6.0,
                    },
                    fill: None,
                    stroke: Some(table.borders.outer.color),
                }];

                if table.header_row && row_count > 0 {
                    draw.push(DrawCommand::Rect {
                        rect: Rect {
                            x: 0.0,
                            y: 0.0,
                            width,
                            height: row_h,
                        },
                        fill: Some(Color::rgb(0.84, 0.9, 0.98)),
                        stroke: None,
                    });
                }

                if table.alternating_rows {
                    for i in 0..row_count {
                        if i % 2 == 1 {
                            draw.push(DrawCommand::Rect {
                                rect: Rect {
                                    x: 0.0,
                                    y: i as f32 * row_h,
                                    width,
                                    height: row_h,
                                },
                                fill: Some(Color::rgba(0.7, 0.74, 0.8, 0.16)),
                                stroke: None,
                            });
                        }
                    }
                }

                (h, draw)
            }
            Block::List(list) => {
                let mut all = Vec::new();
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
                    all.push(DrawCommand::Text {
                        text: bullet,
                        rect: Rect {
                            x: 0.0,
                            y,
                            width: 22.0,
                            height: 20.0,
                        },
                        size: cfg.default_font_size,
                        bold: false,
                        italic: false,
                        underline: false,
                        color: None,
                        alignment: ParagraphAlignment::Left,
                    });
                    y += 22.0;
                }
                (y + 4.0, all)
            }
            Block::PageBreak => (page_size_from_cfg(cfg).1, vec![]),
            Block::BlockQuote(_) => (32.0, vec![]),
            Block::CodeBlock(code) => (
                28.0 + code.code.lines().count() as f32 * (cfg.default_font_size + 3.0),
                vec![DrawCommand::Rect {
                    rect: Rect {
                        x: 0.0,
                        y: 0.0,
                        width,
                        height: 28.0 + code.code.lines().count() as f32 * (cfg.default_font_size + 3.0),
                    },
                    fill: Some(Color::rgb(0.14, 0.14, 0.16)),
                    stroke: Some(Color::rgb(0.24, 0.24, 0.28)),
                }],
            ),
        }
    }

    fn layout_paragraph(
        &mut self,
        paragraph: &Paragraph,
        cfg: &RenderConfig,
        width: f32,
        fallback_size: f32,
    ) -> (f32, Vec<DrawCommand>) {
        let mut commands = Vec::new();
        let mut y = paragraph.spacing.before.max(0.0);

        for run in &paragraph.runs {
            let size = run.style.font_size.unwrap_or(fallback_size).max(8.0);
            let style_key = style_key(run, paragraph);
            *self.text_style_cache.entry(style_key).or_insert(0) += 1;

            let estimated_lines = estimate_wrap_lines(&run.text, width, size);
            let height = estimated_lines as f32 * (size * 1.35);

            commands.push(DrawCommand::Text {
                text: run.text.clone(),
                rect: Rect {
                    x: paragraph.indent.left,
                    y,
                    width: (width - paragraph.indent.left - paragraph.indent.right).max(40.0),
                    height,
                },
                size,
                bold: run.style.bold,
                italic: run.style.italic,
                underline: run.style.underline,
                color: run.style.color,
                alignment: paragraph.alignment.clone(),
            });
            y += height;
        }

        y += paragraph.spacing.after.max(0.0).max(cfg.default_font_size * 0.25);
        (y, commands)
    }
}

pub fn draw_docx(model: &DocumentModel) -> Vec<LaidOutPage> {
    let mut engine = DocxRenderEngine::default();
    engine.paginate(model, &RenderConfig::default())
}

fn style_key(run: &Run, paragraph: &Paragraph) -> String {
    format!(
        "{}|{}|{}|{}|{:?}|{:?}|{:?}",
        run.style.font_family.clone().unwrap_or_else(|| "Segoe UI".to_string()),
        run.style.font_size.unwrap_or(12.0),
        run.style.bold,
        run.style.italic,
        run.style.color,
        paragraph.alignment,
        paragraph.style_id
    )
}

fn estimate_wrap_lines(text: &str, width: f32, font_size: f32) -> usize {
    let avg_char_w = (font_size * 0.52).max(4.0);
    let max_chars = (width / avg_char_w).max(8.0) as usize;
    let mut lines = 0;
    for paragraph in text.split('\n') {
        let count = paragraph.chars().count().max(1);
        lines += count.div_ceil(max_chars);
    }
    lines.max(1)
}

fn page_size(doc: &DocumentModel) -> (f32, f32) {
    use crate::document::model::PageSize;
    match doc.metadata.page_size {
        PageSize::Letter => (612.0, 792.0),
        PageSize::A4 => (595.0, 842.0),
        PageSize::Legal => (612.0, 1008.0),
        PageSize::Custom {
            width_points,
            height_points,
        } => (width_points, height_points),
    }
}

fn page_size_from_cfg(cfg: &RenderConfig) -> (f32, f32) {
    let w = (cfg.margin * 2.0 + 468.0).max(300.0);
    let h = (cfg.margin * 2.0 + 648.0).max(400.0);
    (w, h)
}
