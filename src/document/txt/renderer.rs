use crate::{
    document::txt::{TextDocument, TextWrapMode},
    ui::Rect,
};

#[derive(Debug, Clone)]
pub struct VisibleLine {
    pub line_number: usize,
    pub text: String,
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
        let line_capacity = (viewport.height / self.line_height).ceil() as usize + 3;
        let end_line = (start_line + line_capacity).min(doc.line_count());

        let mut out = Vec::with_capacity(end_line.saturating_sub(start_line));

        for line_idx in start_line..end_line {
            let text = doc.rope.line(line_idx).to_string();
            let wrapped = apply_wrap(&text, doc.wrap_mode, wrap_width_chars.max(4));
            for (sub, chunk) in wrapped.into_iter().enumerate() {
                out.push(VisibleLine {
                    line_number: line_idx + 1,
                    text: chunk,
                    y: viewport.y + (line_idx - start_line) as f32 * self.line_height
                        + sub as f32 * self.line_height,
                });
            }
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
    let mut out = Vec::new();
    let mut current = String::new();
    for word in line.split_whitespace() {
        let tentative_len = current.chars().count() + word.chars().count() + 1;
        if tentative_len > max_chars && !current.is_empty() {
            out.push(std::mem::take(&mut current));
        }
        if !current.is_empty() {
            current.push(' ');
        }
        current.push_str(word);
    }
    if !current.is_empty() {
        out.push(current);
    }
    if out.is_empty() {
        out.push(String::new());
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
