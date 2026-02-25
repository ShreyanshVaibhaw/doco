use std::{
    collections::HashSet,
    hash::{Hash, Hasher},
    time::{Duration, Instant},
};

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

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
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

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct SearchCacheKey {
    query: String,
    options: SearchOptions,
    doc_fingerprint: u64,
}

#[derive(Debug, Clone)]
struct CachedRegex {
    query: String,
    options: SearchOptions,
    compiled: Regex,
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
    cache_key: Option<SearchCacheKey>,
    compiled_regex: Option<CachedRegex>,
    background_blocks: Vec<SearchableBlock>,
    background_cursor: usize,
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
            cache_key: None,
            compiled_regex: None,
            background_blocks: Vec::new(),
            background_cursor: 0,
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
        self.refresh_results_with_visible(doc, &[])
    }

    pub fn refresh_results_with_visible(
        &mut self,
        doc: &DocumentModel,
        visible_block_ids: &[BlockId],
    ) -> usize {
        if self.query.is_empty() {
            self.results.clear();
            self.current_index = 0;
            self.background_blocks.clear();
            self.background_cursor = 0;
            self.pending_live_update = false;
            self.cache_key = None;
            self.result_count_text = "0 results".to_string();
            return 0;
        }

        let doc_fingerprint = document_fingerprint(doc);
        let cache_key = SearchCacheKey {
            query: self.query.clone(),
            options: self.options,
            doc_fingerprint,
        };
        let background_done = self.background_cursor >= self.background_blocks.len();
        if self.cache_key.as_ref() == Some(&cache_key) && background_done {
            self.pending_live_update = false;
            self.update_result_count_text();
            return self.results.len();
        }

        let all_blocks = collect_searchable_blocks(doc);
        let (visible_blocks, background_blocks) =
            split_visible_and_background_blocks(all_blocks, visible_block_ids);
        let regex = self.ensure_compiled_regex();
        self.results = search_blocks(
            &visible_blocks,
            self.query.as_str(),
            self.options,
            regex.as_ref(),
        );
        self.current_index = self.current_index.min(self.results.len().saturating_sub(1));
        self.background_blocks = background_blocks;
        self.background_cursor = 0;
        self.cache_key = Some(cache_key);
        self.pending_live_update = false;
        self.update_result_count_text();

        self.results.len()
    }

    pub fn has_pending_background_search(&self) -> bool {
        self.background_cursor < self.background_blocks.len()
    }

    pub fn process_background_search(&mut self, budget_blocks: usize) -> bool {
        if budget_blocks == 0 || !self.has_pending_background_search() {
            return false;
        }

        let end = (self.background_cursor + budget_blocks).min(self.background_blocks.len());
        let regex = self.ensure_compiled_regex();
        let chunk = self.background_blocks[self.background_cursor..end].to_vec();
        let mut chunk_matches =
            search_blocks(&chunk, self.query.as_str(), self.options, regex.as_ref());
        let changed = !chunk_matches.is_empty();
        self.results.append(&mut chunk_matches);
        self.background_cursor = end;
        if !self.has_pending_background_search() {
            self.background_blocks.clear();
            self.background_cursor = 0;
        }
        self.current_index = self.current_index.min(self.results.len().saturating_sub(1));
        self.update_result_count_text();
        changed
    }

    pub fn invalidate_cache(&mut self) {
        self.cache_key = None;
        self.background_blocks.clear();
        self.background_cursor = 0;
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

    fn ensure_compiled_regex(&mut self) -> Option<Regex> {
        if !self.options.regex || self.query.is_empty() {
            self.compiled_regex = None;
            return None;
        }

        let needs_rebuild = self
            .compiled_regex
            .as_ref()
            .is_none_or(|cached| cached.query != self.query || cached.options != self.options);
        if needs_rebuild {
            self.compiled_regex = build_regex(self.query.as_str(), self.options)
                .ok()
                .map(|compiled| CachedRegex {
                    query: self.query.clone(),
                    options: self.options,
                    compiled,
                });
        }

        self.compiled_regex
            .as_ref()
            .map(|cached| cached.compiled.clone())
    }

    fn update_result_count_text(&mut self) {
        if self.query.is_empty() {
            self.result_count_text = "0 results".to_string();
        } else if self.has_pending_background_search() {
            self.result_count_text = format!("{}+ results for '{}'", self.results.len(), self.query);
        } else {
            self.result_count_text = format!("{} results for '{}'", self.results.len(), self.query);
        }
    }
}

pub fn search_document(doc: &DocumentModel, query: &str, options: SearchOptions) -> Vec<SearchMatch> {
    if query.is_empty() {
        return Vec::new();
    }

    let blocks = collect_searchable_blocks(doc);
    let compiled = if options.regex {
        build_regex(query, options).ok()
    } else {
        None
    };
    search_blocks(&blocks, query, options, compiled.as_ref())
}

fn search_blocks(
    blocks: &[SearchableBlock],
    query: &str,
    options: SearchOptions,
    compiled_regex: Option<&Regex>,
) -> Vec<SearchMatch> {
    if query.is_empty() {
        return Vec::new();
    }

    let mut matches = Vec::new();
    if options.regex {
        let Some(regex) = compiled_regex else {
            return Vec::new();
        };

        for block in blocks {
            for cap in regex.captures_iter(block.text.as_str()) {
                if let Some(m) = cap.get(0) {
                    if options.whole_word && !is_whole_word(block.text.as_str(), m.start(), m.end()) {
                        continue;
                    }

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

fn split_visible_and_background_blocks(
    blocks: Vec<SearchableBlock>,
    visible_block_ids: &[BlockId],
) -> (Vec<SearchableBlock>, Vec<SearchableBlock>) {
    if blocks.len() < 10_000 {
        return (blocks, Vec::new());
    }

    let visible_ids: HashSet<BlockId> = visible_block_ids.iter().copied().collect();
    let mut visible_blocks = Vec::new();
    let mut background_blocks = Vec::new();

    if visible_ids.is_empty() {
        let split = blocks.len().min(512);
        visible_blocks.extend_from_slice(&blocks[..split]);
        background_blocks.extend_from_slice(&blocks[split..]);
        return (visible_blocks, background_blocks);
    }

    for block in blocks {
        if visible_ids.contains(&block.id) {
            visible_blocks.push(block);
        } else {
            background_blocks.push(block);
        }
    }

    if visible_blocks.is_empty() && !background_blocks.is_empty() {
        let mut fallback = Vec::new();
        let split = background_blocks.len().min(512);
        fallback.extend_from_slice(&background_blocks[..split]);
        background_blocks = background_blocks[split..].to_vec();
        return (fallback, background_blocks);
    }

    (visible_blocks, background_blocks)
}

fn document_fingerprint(doc: &DocumentModel) -> u64 {
    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    for block in collect_searchable_blocks(doc) {
        block.id.hash(&mut hasher);
        block.line_or_page.hash(&mut hasher);
        block.text.hash(&mut hasher);
    }
    hasher.finish()
}

pub fn replace_current(doc: &mut DocumentModel, state: &mut FindReplaceState) -> usize {
    let Some(current) = state.current_result().cloned() else {
        return 0;
    };
    let regex = state.ensure_compiled_regex();

    if replace_in_block(
        doc,
        &current,
        state.query.as_str(),
        state.replacement.as_str(),
        state.options,
        regex.as_ref(),
    ) {
        state.last_replaced_count = 1;
        state.invalidate_cache();
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
        let Some(regex) = state.ensure_compiled_regex() else {
            return 0;
        };

        mutate_text_blocks(
            doc,
            |text| {
                let count = regex
                    .find_iter(text.as_str())
                    .filter(|m| !state.options.whole_word || is_whole_word(text.as_str(), m.start(), m.end()))
                    .count();
                if count == 0 {
                    return None;
                }
                replaced += count;
                let out = regex
                    .replace_all(text.as_str(), state.replacement.as_str())
                    .to_string();
                Some(out)
            },
            None,
        );
    } else {
        let options = state.options;
        let query = state.query.clone();
        let replacement = state.replacement.clone();
        mutate_text_blocks(
            doc,
            |text| {
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
            },
            None,
        );
    }

    state.last_replaced_count = replaced;
    state.invalidate_cache();
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
    let pattern = if options.whole_word {
        format!(r"\b(?:{})\b", query)
    } else {
        query.to_string()
    };
    let mut builder = RegexBuilder::new(pattern.as_str());
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
    query: &str,
    replacement: &str,
    options: SearchOptions,
    regex: Option<&Regex>,
) -> bool {
    let mut replaced = false;

    mutate_text_blocks(
        doc,
        |text| {
            if replaced {
                return None;
            }

            if target.end > text.len() || target.start > target.end {
                return None;
            }
            if !is_char_boundary(text, target.start) || !is_char_boundary(text, target.end) {
                return None;
            }

            if options.regex {
                let regex = regex?;
                let matched = &text[target.start..target.end];
                let captures = regex.captures(matched)?;
                let whole = captures.get(0)?;
                if whole.start() != 0 || whole.end() != matched.len() {
                    return None;
                }

                let mut expanded = String::new();
                captures.expand(replacement, &mut expanded);
                let mut out = String::new();
                out.push_str(&text[..target.start]);
                out.push_str(&expanded);
                out.push_str(&text[target.end..]);
                replaced = true;
                return Some(out);
            }

            let transformed_hay = if options.case_sensitive {
                text.clone()
            } else {
                text.to_ascii_lowercase()
            };
            let transformed_query = if options.case_sensitive {
                query.to_string()
            } else {
                query.to_ascii_lowercase()
            };
            if transformed_query.is_empty() {
                return None;
            }

            if target.end <= transformed_hay.len()
                && transformed_hay[target.start..target.end] == transformed_query
            {
                let mut out = String::new();
                out.push_str(&text[..target.start]);
                out.push_str(replacement);
                out.push_str(&text[target.end..]);
                replaced = true;
                return Some(out);
            }
            None
        },
        Some(target.block_id),
    );

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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::document::model::{
        Document,
        ImageBlock,
        ListItem,
        Run,
        TableCell,
        TableRow,
    };

    fn paragraph_block(id: u64, text: &str) -> Block {
        Block::Paragraph(Paragraph {
            id: BlockId(id),
            runs: vec![Run {
                text: text.to_string(),
                ..Run::default()
            }],
            alignment: crate::document::model::ParagraphAlignment::Left,
            spacing: crate::document::model::ParagraphSpacing::default(),
            indent: crate::document::model::Indent::default(),
            style_id: None,
        })
    }

    fn heading_block(id: u64, text: &str) -> Block {
        Block::Heading(Heading {
            id: BlockId(id),
            runs: vec![Run {
                text: text.to_string(),
                ..Run::default()
            }],
            ..Heading::default()
        })
    }

    fn table_block(id: u64, cell_block: Block) -> Block {
        Block::Table(Table {
            id: BlockId(id),
            rows: vec![TableRow {
                cells: vec![TableCell {
                    blocks: vec![cell_block],
                    ..TableCell::default()
                }],
            }],
            ..Table::default()
        })
    }

    fn list_block(item_id: u64, block: Block) -> Block {
        Block::List(List {
            items: vec![ListItem {
                id: BlockId(item_id),
                content: vec![block],
                ..ListItem::default()
            }],
            ..List::default()
        })
    }

    fn image_block(id: u64) -> Block {
        Block::Image(ImageBlock {
            id: BlockId(id),
            alt_text: "needle in image metadata".to_string(),
            ..ImageBlock::default()
        })
    }

    fn doc_with_blocks(blocks: Vec<Block>) -> DocumentModel {
        Document {
            content: blocks,
            ..Document::default()
        }
    }

    fn paragraph_text_by_id(doc: &DocumentModel, id: BlockId) -> String {
        doc.content
            .iter()
            .find_map(|block| match block {
                Block::Paragraph(p) if p.id == id => Some(
                    p.runs
                        .iter()
                        .map(|r| r.text.as_str())
                        .collect::<String>(),
                ),
                _ => None,
            })
            .unwrap_or_default()
    }

    #[test]
    fn search_walks_supported_blocks_and_skips_images() {
        let doc = doc_with_blocks(vec![
            paragraph_block(1, "Needle in paragraph"),
            heading_block(2, "needle in heading"),
            table_block(10, paragraph_block(3, "table needle")),
            list_block(20, paragraph_block(4, "list needle")),
            image_block(99),
        ]);

        let matches = search_document(
            &doc,
            "needle",
            SearchOptions {
                case_sensitive: false,
                whole_word: false,
                regex: false,
            },
        );

        let ids = matches.iter().map(|m| m.block_id.0).collect::<Vec<_>>();
        assert!(ids.contains(&1));
        assert!(ids.contains(&2));
        assert!(ids.contains(&3));
        assert!(ids.contains(&4));
        assert!(!ids.contains(&99));
    }

    #[test]
    fn whole_word_avoids_partial_matches() {
        let doc = doc_with_blocks(vec![paragraph_block(1, "cat scatter cat catalog")]);
        let matches = search_document(
            &doc,
            "cat",
            SearchOptions {
                case_sensitive: false,
                whole_word: true,
                regex: false,
            },
        );
        assert_eq!(matches.len(), 2);
    }

    #[test]
    fn regex_captures_and_replace_backreferences() {
        let mut doc = doc_with_blocks(vec![paragraph_block(1, "abc-123 def-456")]);
        let mut state = FindReplaceState {
            query: "([a-z]+)-(\\d+)".to_string(),
            replacement: "$2/$1".to_string(),
            options: SearchOptions {
                case_sensitive: false,
                whole_word: false,
                regex: true,
            },
            ..FindReplaceState::default()
        };

        state.refresh_results(&doc);
        assert_eq!(state.results.len(), 2);
        assert_eq!(state.results[0].capture_groups.len(), 2);

        let replaced = replace_all(&mut doc, &mut state);
        assert_eq!(replaced, 2);
        assert_eq!(paragraph_text_by_id(&doc, BlockId(1)), "123/abc 456/def");
    }

    #[test]
    fn replace_current_supports_regex_backreferences() {
        let mut doc = doc_with_blocks(vec![paragraph_block(1, "foo-10 bar-20")]);
        let mut state = FindReplaceState {
            query: "([a-z]+)-(\\d+)".to_string(),
            replacement: "$2:$1".to_string(),
            options: SearchOptions {
                case_sensitive: false,
                whole_word: false,
                regex: true,
            },
            ..FindReplaceState::default()
        };
        state.refresh_results(&doc);
        assert_eq!(state.current_index, 0);
        assert_eq!(replace_current(&mut doc, &mut state), 1);
        assert_eq!(paragraph_text_by_id(&doc, BlockId(1)), "10:foo bar-20");
    }

    #[test]
    fn large_document_search_runs_visible_then_background() {
        let mut blocks = Vec::new();
        for i in 0..10_050 {
            let text = if i == 2 {
                "visible hit"
            } else if i == 10_020 {
                "late hit"
            } else {
                "filler"
            };
            blocks.push(paragraph_block((i + 1) as u64, text));
        }
        let doc = doc_with_blocks(blocks);

        let mut state = FindReplaceState {
            query: "hit".to_string(),
            options: SearchOptions::default(),
            ..FindReplaceState::default()
        };

        state.refresh_results_with_visible(&doc, &[BlockId(3)]);
        assert_eq!(state.results.len(), 1);
        assert!(state.has_pending_background_search());

        state.process_background_search(20_000);
        assert_eq!(state.results.len(), 2);
        assert!(!state.has_pending_background_search());
    }

    #[test]
    fn refresh_uses_cache_until_invalidated() {
        let doc = doc_with_blocks(vec![paragraph_block(1, "alpha beta alpha")]);
        let mut state = FindReplaceState {
            query: "alpha".to_string(),
            options: SearchOptions {
                case_sensitive: false,
                whole_word: false,
                regex: true,
            },
            ..FindReplaceState::default()
        };

        state.refresh_results(&doc);
        let ptr_before = state
            .compiled_regex
            .as_ref()
            .map(|v| &v.compiled as *const Regex)
            .expect("compiled regex");
        let count_before = state.results.len();

        state.refresh_results(&doc);
        let ptr_after = state
            .compiled_regex
            .as_ref()
            .map(|v| &v.compiled as *const Regex)
            .expect("compiled regex");
        assert_eq!(count_before, state.results.len());
        assert_eq!(ptr_before, ptr_after);

        state.invalidate_cache();
        assert!(state.cache_key.is_none());
    }
}
