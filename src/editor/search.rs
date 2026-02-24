use std::time::{Duration, Instant};

use regex::{Regex, RegexBuilder};

use crate::document::model::{
    Block,
    BlockId,
    DocumentModel,
    Heading,
    List,
    Paragraph,
    Table,
};

#[derive(Debug, Clone, Copy)]
pub struct SearchOptions {
    pub case_sensitive: bool,
    pub whole_word: bool,
    pub regex: bool,
}

impl Default for SearchOptions {
    fn default() -> Self {
        Self {
            case_sensitive: false,
            whole_word: false,
            regex: false,
        }
    }
}

#[derive(Debug, Clone)]
pub struct SearchMatch {
    pub block_id: BlockId,
    pub start: usize,
    pub end: usize,
    pub line_or_page: usize,
    pub snippet: String,
    pub capture_groups: Vec<(usize, usize)>,
}

#[derive(Debug, Clone)]
pub struct FindReplaceState {
    pub find_visible: bool,
    pub replace_visible: bool,
    pub query: String,
    pub replacement: String,
    pub options: SearchOptions,
    pub results: Vec<SearchMatch>,
    pub current_index: usize,
    pub result_count_text: String,
    pub last_replaced_count: usize,
    pub debounce_ms: u64,
    pub last_input_at: Instant,
    pub pending_live_update: bool,
}

impl Default for FindReplaceState {
    fn default() -> Self {
        Self {
            find_visible: false,
            replace_visible: false,
            query: String::new(),
            replacement: String::new(),
            options: SearchOptions::default(),
            results: Vec::new(),
            current_index: 0,
            result_count_text: "0 results".to_string(),
            last_replaced_count: 0,
            debounce_ms: 100,
            last_input_at: Instant::now(),
            pending_live_update: false,
        }
    }
}

impl FindReplaceState {
    pub fn open_find(&mut self) {
        self.find_visible = true;
        self.replace_visible = false;
    }

    pub fn open_replace(&mut self) {
        self.find_visible = true;
        self.replace_visible = true;
    }

    pub fn close(&mut self) {
        self.find_visible = false;
        self.replace_visible = false;
    }

    pub fn set_query(&mut self, query: String) {
        self.query = query;
        self.pending_live_update = true;
        self.last_input_at = Instant::now();
    }

    pub fn set_replacement(&mut self, replacement: String) {
        self.replacement = replacement;
    }

    pub fn should_live_update(&self, now: Instant) -> bool {
        self.pending_live_update
            && now.duration_since(self.last_input_at) >= Duration::from_millis(self.debounce_ms)
    }

    pub fn refresh_results(&mut self, doc: &DocumentModel) -> usize {
        let results = search_document(doc, self.query.as_str(), self.options);
        self.current_index = self.current_index.min(results.len().saturating_sub(1));
        self.result_count_text = format!("{} results for '{}'", results.len(), self.query);
        self.pending_live_update = false;
        self.results = results;
        self.results.len()
    }

    pub fn current_result(&self) -> Option<&SearchMatch> {
        self.results.get(self.current_index)
    }

    pub fn next(&mut self) -> Option<&SearchMatch> {
        if self.results.is_empty() {
            return None;
        }
        self.current_index = (self.current_index + 1) % self.results.len();
        self.current_result()
    }

    pub fn previous(&mut self) -> Option<&SearchMatch> {
        if self.results.is_empty() {
            return None;
        }
        if self.current_index == 0 {
            self.current_index = self.results.len() - 1;
        } else {
            self.current_index -= 1;
        }
        self.current_result()
    }
}

pub fn search_document(doc: &DocumentModel, query: &str, options: SearchOptions) -> Vec<SearchMatch> {
    if query.is_empty() {
        return Vec::new();
    }

    let blocks = collect_searchable_blocks(doc);
    let mut matches = Vec::new();

    if options.regex {
        let Ok(regex) = build_regex(query, options) else {
            return Vec::new();
        };

        for block in blocks {
            for cap in regex.captures_iter(block.text.as_str()) {
                if let Some(m) = cap.get(0) {
                    let groups = (1..cap.len())
                        .filter_map(|i| cap.get(i).map(|g| (g.start(), g.end())))
                        .collect::<Vec<_>>();

                    matches.push(SearchMatch {
                        block_id: block.id,
                        start: m.start(),
                        end: m.end(),
                        line_or_page: block.line_or_page,
                        snippet: snippet(block.text.as_str(), m.start(), m.end()),
                        capture_groups: groups,
                    });
                }
            }
        }

        return matches;
    }

    let (needle, hay_transform): (String, fn(&str) -> String) = if options.case_sensitive {
        (query.to_string(), |s| s.to_string())
    } else {
        (query.to_ascii_lowercase(), |s| s.to_ascii_lowercase())
    };

    for block in blocks {
        let hay = hay_transform(block.text.as_str());
        let mut from = 0usize;

        while from < hay.len() {
            let Some(rel) = hay[from..].find(needle.as_str()) else {
                break;
            };
            let start = from + rel;
            let end = start + needle.len();

            if options.whole_word && !is_whole_word(block.text.as_str(), start, end) {
                from = start.saturating_add(1);
                continue;
            }

            matches.push(SearchMatch {
                block_id: block.id,
                start,
                end,
                line_or_page: block.line_or_page,
                snippet: snippet(block.text.as_str(), start, end),
                capture_groups: Vec::new(),
            });

            from = end.max(start.saturating_add(1));
        }
    }

    matches
}

pub fn replace_current(doc: &mut DocumentModel, state: &mut FindReplaceState) -> usize {
    let Some(current) = state.current_result().cloned() else {
        return 0;
    };

    if replace_in_block(doc, &current, state.replacement.as_str(), state.options) {
        state.last_replaced_count = 1;
        state.refresh_results(doc);
        1
    } else {
        0
    }
}

pub fn replace_all(doc: &mut DocumentModel, state: &mut FindReplaceState) -> usize {
    if state.query.is_empty() {
        return 0;
    }

    let mut replaced = 0usize;

    if state.options.regex {
        let Ok(regex) = build_regex(state.query.as_str(), state.options) else {
            return 0;
        };

        mutate_text_blocks(doc, |text| {
            let count = regex.find_iter(text.as_str()).count();
            if count == 0 {
                return None;
            }
            replaced += count;
            let out = regex
                .replace_all(text.as_str(), state.replacement.as_str())
                .to_string();
            Some(out)
        }, None);
    } else {
        let options = state.options;
        let query = state.query.clone();
        let replacement = state.replacement.clone();
        mutate_text_blocks(doc, |text| {
            let mut local = 0usize;
            let transformed_hay = if options.case_sensitive {
                text.clone()
            } else {
                text.to_ascii_lowercase()
            };
            let transformed_query = if options.case_sensitive {
                query.clone()
            } else {
                query.to_ascii_lowercase()
            };

            let mut cursor = 0usize;
            let mut out = String::new();
            while cursor < transformed_hay.len() {
                let Some(rel) = transformed_hay[cursor..].find(transformed_query.as_str()) else {
                    out.push_str(&text[cursor..]);
                    break;
                };

                let start = cursor + rel;
                let end = start + transformed_query.len();

                if options.whole_word && !is_whole_word(text.as_str(), start, end) {
                    out.push_str(&text[cursor..start.saturating_add(1)]);
                    cursor = start.saturating_add(1);
                    continue;
                }

                out.push_str(&text[cursor..start]);
                out.push_str(&replacement);
                cursor = end;
                local += 1;
            }

            if local > 0 {
                replaced += local;
                Some(out)
            } else {
                None
            }
        }, None);
    }

    state.last_replaced_count = replaced;
    state.refresh_results(doc);
    replaced
}

pub fn replacement_preview(current: &SearchMatch, replacement: &str) -> String {
    format!("{} -> {}", current.snippet, replacement)
}

#[derive(Debug, Clone)]
struct SearchableBlock {
    id: BlockId,
    line_or_page: usize,
    text: String,
}

fn collect_searchable_blocks(doc: &DocumentModel) -> Vec<SearchableBlock> {
    let mut out = Vec::new();
    for (index, block) in doc.content.iter().enumerate() {
        collect_block(block, index + 1, &mut out);
    }
    out
}

fn collect_block(block: &Block, line_or_page: usize, out: &mut Vec<SearchableBlock>) {
    match block {
        Block::Paragraph(p) => {
            out.push(SearchableBlock {
                id: p.id,
                line_or_page,
                text: paragraph_text(p),
            });
        }
        Block::Heading(h) => {
            out.push(SearchableBlock {
                id: h.id,
                line_or_page,
                text: heading_text(h),
            });
        }
        Block::Table(t) => collect_table(t, line_or_page, out),
        Block::List(l) => collect_list(l, line_or_page, out),
        Block::BlockQuote(q) => {
            for block in &q.blocks {
                collect_block(block, line_or_page, out);
            }
        }
        Block::CodeBlock(code) => {
            out.push(SearchableBlock {
                id: code.id,
                line_or_page,
                text: code.code.clone(),
            });
        }
        Block::Image(_) | Block::PageBreak | Block::HorizontalRule => {}
    }
}

fn collect_table(table: &Table, line_or_page: usize, out: &mut Vec<SearchableBlock>) {
    for row in &table.rows {
        for cell in &row.cells {
            for block in &cell.blocks {
                collect_block(block, line_or_page, out);
            }
        }
    }
}

fn collect_list(list: &List, line_or_page: usize, out: &mut Vec<SearchableBlock>) {
    for item in &list.items {
        for block in &item.content {
            collect_block(block, line_or_page, out);
        }
        for child in &item.children {
            let nested = List {
                items: vec![child.clone()],
                list_type: list.list_type.clone(),
                start_number: list.start_number,
            };
            collect_list(&nested, line_or_page, out);
        }
    }
}

fn paragraph_text(p: &Paragraph) -> String {
    p.runs.iter().map(|r| r.text.as_str()).collect()
}

fn heading_text(h: &Heading) -> String {
    h.runs.iter().map(|r| r.text.as_str()).collect()
}

fn build_regex(query: &str, options: SearchOptions) -> Result<Regex, regex::Error> {
    let mut builder = RegexBuilder::new(query);
    builder.case_insensitive(!options.case_sensitive);
    builder.build()
}

fn is_whole_word(text: &str, start: usize, end: usize) -> bool {
    let before_ok = if start == 0 {
        true
    } else {
        text[..start]
            .chars()
            .last()
            .map(|c| !is_word_char(c))
            .unwrap_or(true)
    };

    let after_ok = if end >= text.len() {
        true
    } else {
        text[end..]
            .chars()
            .next()
            .map(|c| !is_word_char(c))
            .unwrap_or(true)
    };

    before_ok && after_ok
}

fn is_word_char(ch: char) -> bool {
    ch.is_ascii_alphanumeric() || ch == '_'
}

fn snippet(text: &str, start: usize, end: usize) -> String {
    let begin = start.saturating_sub(24);
    let finish = (end + 24).min(text.len());
    text[begin..finish].replace('\n', " ")
}

fn replace_in_block(
    doc: &mut DocumentModel,
    target: &SearchMatch,
    replacement: &str,
    options: SearchOptions,
) -> bool {
    let mut replaced = false;

    mutate_text_blocks(doc, |text| {
        if replaced {
            return None;
        }

        if options.regex {
            return None;
        }

        let transformed_hay = if options.case_sensitive {
            text.clone()
        } else {
            text.to_ascii_lowercase()
        };
        let query_len = target.end.saturating_sub(target.start);

        if target.end <= transformed_hay.len() {
            let start = target.start;
            let end = target.end;
            if is_char_boundary(text, start) && is_char_boundary(text, end) {
                let mut out = String::new();
                out.push_str(&text[..start]);
                out.push_str(replacement);
                out.push_str(&text[end..]);
                replaced = true;
                return Some(out);
            }
        }

        if query_len > 0 {
            None
        } else {
            None
        }
    }, Some(target.block_id));

    replaced
}

fn mutate_text_blocks<F>(doc: &mut DocumentModel, mut f: F, only_block: Option<BlockId>)
where
    F: FnMut(&String) -> Option<String>,
{
    for block in &mut doc.content {
        mutate_block(block, &mut f, only_block);
    }
}

fn mutate_block<F>(block: &mut Block, f: &mut F, only_block: Option<BlockId>)
where
    F: FnMut(&String) -> Option<String>,
{
    match block {
        Block::Paragraph(p) => {
            if only_block.is_none_or(|id| id == p.id) {
                let text = paragraph_text(p);
                if let Some(next) = f(&text) {
                    if p.runs.is_empty() {
                        p.runs.push(crate::document::model::Run::default());
                    }
                    p.runs.clear();
                    p.runs.push(crate::document::model::Run {
                        text: next,
                        style: crate::document::model::RunStyle::default(),
                    });
                }
            }
        }
        Block::Heading(h) => {
            if only_block.is_none_or(|id| id == h.id) {
                let text = heading_text(h);
                if let Some(next) = f(&text) {
                    h.runs.clear();
                    h.runs.push(crate::document::model::Run {
                        text: next,
                        style: crate::document::model::RunStyle::default(),
                    });
                }
            }
        }
        Block::Table(t) => {
            for row in &mut t.rows {
                for cell in &mut row.cells {
                    for nested in &mut cell.blocks {
                        mutate_block(nested, f, only_block);
                    }
                }
            }
        }
        Block::List(list) => {
            for item in &mut list.items {
                for nested in &mut item.content {
                    mutate_block(nested, f, only_block);
                }
                for child in &mut item.children {
                    for nested in &mut child.content {
                        mutate_block(nested, f, only_block);
                    }
                }
            }
        }
        Block::BlockQuote(q) => {
            for nested in &mut q.blocks {
                mutate_block(nested, f, only_block);
            }
        }
        Block::CodeBlock(c) => {
            if only_block.is_none_or(|id| id == c.id) {
                if let Some(next) = f(&c.code) {
                    c.code = next;
                }
            }
        }
        Block::Image(_) | Block::PageBreak | Block::HorizontalRule => {}
    }
}

fn is_char_boundary(text: &str, idx: usize) -> bool {
    idx <= text.len() && text.is_char_boundary(idx)
}
