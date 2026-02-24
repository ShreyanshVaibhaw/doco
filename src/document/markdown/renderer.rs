use pulldown_cmark::{
    CodeBlockKind,
    Event,
    HeadingLevel,
    Tag,
    TagEnd,
};

use crate::document::{
    markdown::MarkdownDocument,
    model::{
        Block,
        BlockId,
        BlockQuote,
        CodeBlock,
        DocumentModel,
        Heading,
        List,
        ListItem,
        ListType,
        Paragraph,
        ParagraphAlignment,
        ParagraphSpacing,
        Run,
        RunStyle,
    },
};

#[derive(Default)]
struct ListBuilder {
    list_type: ListType,
    start_number: u32,
    items: Vec<ListItem>,
    current_item_runs: Vec<Run>,
}

pub fn markdown_to_model(doc: &MarkdownDocument) -> DocumentModel {
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

                        if let Some(list) = list_stack.last_mut() {
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
                            checked: None,
                            children: vec![],
                        });
                        next_id += 1;
                        list.current_item_runs.clear();
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
                _ => {}
            },
            Event::Text(text) => {
                let run = Run {
                    text: text.to_string(),
                    style: RunStyle::default(),
                };

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
                let run = Run {
                    text: text.to_string(),
                    style: RunStyle {
                        font_family: Some("Cascadia Mono".to_string()),
                        ..RunStyle::default()
                    },
                };
                if in_heading.is_some() {
                    heading_runs.push(run);
                } else if in_paragraph {
                    current_runs.push(run);
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
