use crate::{
    document::txt::{TextDocument, TextWrapMode},
    ui::Rect,
};

#[derive(Debug, Clone)]
pub struct VisibleLine {
    pub line_number: usize,
    pub display_line_number: Option<usize>,
    pub text: String,
    pub x: f32,
    pub y: f32,
}

#[derive(Debug)]
pub struct TextViewportRenderer {
    pub line_height: f32,
    pub gutter_width: f32,
}

impl Default for TextViewportRenderer {
    fn default() -> Self {
        Self {
            line_height: 20.0,
            gutter_width: 56.0,
        }
    }
}

impl TextViewportRenderer {
    pub fn visible_lines(
        &self,
        doc: &TextDocument,
        viewport: Rect,
        scroll_y: f32,
        wrap_width_chars: usize,
    ) -> Vec<VisibleLine> {
        let start_line = (scroll_y / self.line_height).floor().max(0.0) as usize;
        let visual_capacity = (viewport.height / self.line_height).ceil() as usize + 4;

        let mut out = Vec::with_capacity(visual_capacity);
        let mut line_idx = start_line.min(doc.line_count());
        let mut visual_row = 0usize;

        while line_idx < doc.line_count() && visual_row < visual_capacity {
            let text = doc.line_text(line_idx).unwrap_or_default();
            let wrapped = apply_wrap(&text, doc.wrap_mode, wrap_width_chars.max(4));
            for (sub, chunk) in wrapped.into_iter().enumerate() {
                if visual_row >= visual_capacity {
                    break;
                }

                out.push(VisibleLine {
                    line_number: line_idx + 1,
                    display_line_number: if doc.line_numbers && sub == 0 {
                        Some(line_idx + 1)
                    } else {
                        None
                    },
                    text: chunk,
                    x: viewport.x
                        + if doc.line_numbers {
                            self.gutter_width
                        } else {
                            0.0
                        },
                    y: viewport.y + visual_row as f32 * self.line_height,
                });
                visual_row += 1;
            }
            line_idx += 1;
        }

        out
    }
}

pub fn render_text(doc: &TextDocument, viewport: Rect, scroll_y: f32) -> Vec<VisibleLine> {
    TextViewportRenderer::default().visible_lines(doc, viewport, scroll_y, 120)
}

fn apply_wrap(line: &str, mode: TextWrapMode, max_chars: usize) -> Vec<String> {
    if matches!(mode, TextWrapMode::None) || line.chars().count() <= max_chars {
        return vec![line.to_string()];
    }

    match mode {
        TextWrapMode::WordBoundary => wrap_word(line, max_chars),
        TextWrapMode::Character => wrap_char(line, max_chars),
        TextWrapMode::None => vec![line.to_string()],
    }
}

fn wrap_word(line: &str, max_chars: usize) -> Vec<String> {
    let chars: Vec<char> = line.chars().collect();
    if chars.is_empty() {
        return vec![String::new()];
    }

    let mut out = Vec::new();
    let mut start = 0usize;

    while start < chars.len() {
        let limit = (start + max_chars).min(chars.len());
        if limit == chars.len() {
            out.push(chars[start..limit].iter().collect());
            break;
        }

        let mut split_at = limit;
        for idx in (start..limit).rev() {
            if chars[idx].is_whitespace() {
                split_at = idx + 1;
                break;
            }
        }

        if split_at <= start {
            split_at = limit;
        }

        out.push(chars[start..split_at].iter().collect());
        start = split_at;
    }

    out
}

fn wrap_char(line: &str, max_chars: usize) -> Vec<String> {
    let mut out = Vec::new();
    let mut current = String::new();
    for ch in line.chars() {
        current.push(ch);
        if current.chars().count() >= max_chars {
            out.push(std::mem::take(&mut current));
        }
    }
    if !current.is_empty() {
        out.push(current);
    }
    if out.is_empty() {
        out.push(String::new());
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::document::txt::{TextDocument, TextWrapMode};

    #[test]
    fn wraps_by_word_and_character() {
        let line = "alpha beta gamma";
        let word = wrap_word(line, 7);
        assert_eq!(word, vec!["alpha ".to_string(), "beta ".to_string(), "gamma".to_string()]);

        let ch = wrap_char(line, 5);
        assert_eq!(
            ch,
            vec!["alpha".to_string(), " beta".to_string(), " gamm".to_string(), "a".to_string()]
        );
    }

    #[test]
    fn virtualizes_visible_lines_for_large_docs() {
        let mut text = String::new();
        for i in 0..5000 {
            text.push_str(format!("line-{i}\n").as_str());
        }

        let mut doc = TextDocument::from_text(&text);
        doc.set_wrap_mode(TextWrapMode::None);

        let renderer = TextViewportRenderer::default();
        let viewport = Rect {
            x: 0.0,
            y: 0.0,
            width: 800.0,
            height: 180.0,
        };
        let visible = renderer.visible_lines(&doc, viewport, 4000.0, 120);

        assert!(visible.len() <= 20);
        assert_eq!(visible.first().and_then(|v| v.display_line_number), Some(201));
    }

    #[test]
    fn hides_line_numbers_when_disabled() {
        let mut doc = TextDocument::from_text("a\nb\nc\n");
        doc.set_wrap_mode(TextWrapMode::Character);
        doc.set_line_numbers(false);

        let renderer = TextViewportRenderer::default();
        let viewport = Rect {
            x: 10.0,
            y: 10.0,
            width: 400.0,
            height: 120.0,
        };
        let visible = renderer.visible_lines(&doc, viewport, 0.0, 1);
        assert!(visible.iter().all(|line| line.display_line_number.is_none()));
        assert!(visible.iter().all(|line| (line.x - viewport.x).abs() < f32::EPSILON));
    }
}
