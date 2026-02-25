use crate::document::model::{
    Block, BlockId, DocumentModel, Table, TableBorders, TableCell, TableRow, TableStylePreset,
};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellPos {
    pub row: usize,
    pub col: usize,
}

#[derive(Debug, Clone)]
pub struct TableSelection {
    pub start: CellPos,
    pub end: CellPos,
}

impl TableSelection {
    pub fn normalized(&self) -> TableSelection {
        TableSelection {
            start: CellPos {
                row: self.start.row.min(self.end.row),
                col: self.start.col.min(self.end.col),
            },
            end: CellPos {
                row: self.start.row.max(self.end.row),
                col: self.start.col.max(self.end.col),
            },
        }
    }
}

#[derive(Debug, Clone, Default)]
pub struct TableLayoutCache {
    generation: u64,
    cached_rows: Option<(usize, usize)>,
}

impl TableLayoutCache {
    pub fn invalidate(&mut self) {
        self.generation = self.generation.saturating_add(1);
        self.cached_rows = None;
    }

    pub fn visible_rows(
        &mut self,
        table: &Table,
        scroll_y: f32,
        viewport_h: f32,
        default_row_h: f32,
    ) -> (usize, usize) {
        let range = visible_row_range(table, scroll_y, viewport_h, default_row_h);
        self.cached_rows = Some(range);
        range
    }

    pub fn generation(&self) -> u64 {
        self.generation
    }
}

pub fn insert_table(doc: &mut DocumentModel, at_index: usize, rows: usize, cols: usize) -> BlockId {
    let next_id = next_block_id(doc);
    let cols = cols.clamp(1, 64);
    let rows = rows.clamp(1, 2000);

    let table = Table {
        id: next_id,
        rows: (0..rows)
            .map(|_| TableRow {
                cells: (0..cols)
                    .map(|_| TableCell {
                        blocks: Vec::new(),
                        rowspan: 1,
                        colspan: 1,
                        background: None,
                    })
                    .collect(),
            })
            .collect(),
        column_widths: vec![120.0; cols],
        row_heights: vec![28.0; rows],
        borders: TableBorders::default(),
        style: TableStylePreset::Grid,
        cell_padding: 4.0,
        header_row: false,
        alternating_rows: false,
    };

    let idx = at_index.min(doc.content.len());
    doc.content.insert(idx, Block::Table(table));
    doc.dirty = true;
    next_id
}

pub fn insert_row(table: &mut Table, at: usize) {
    ensure_row_heights(table);
    let cols = table.column_widths.len().max(1);
    let row = TableRow {
        cells: (0..cols)
            .map(|_| TableCell {
                blocks: Vec::new(),
                rowspan: 1,
                colspan: 1,
                background: None,
            })
            .collect(),
    };
    let insert_at = at.min(table.rows.len());
    table.rows.insert(insert_at, row);
    table.row_heights.insert(insert_at, 28.0);
}

pub fn delete_row(table: &mut Table, row: usize) -> bool {
    ensure_row_heights(table);
    if row < table.rows.len() {
        table.rows.remove(row);
        if row < table.row_heights.len() {
            table.row_heights.remove(row);
        }
        true
    } else {
        false
    }
}

pub fn insert_row_above(table: &mut Table, row: usize) {
    insert_row(table, row);
}

pub fn insert_row_below(table: &mut Table, row: usize) {
    insert_row(table, row.saturating_add(1));
}

pub fn insert_column(table: &mut Table, at: usize) {
    let col = at.min(table.column_widths.len());
    table.column_widths.insert(col, 120.0);
    for row in &mut table.rows {
        row.cells.insert(
            col,
            TableCell {
                blocks: Vec::new(),
                rowspan: 1,
                colspan: 1,
                background: None,
            },
        );
    }
}

pub fn delete_column(table: &mut Table, col: usize) -> bool {
    if col >= table.column_widths.len() {
        return false;
    }

    table.column_widths.remove(col);
    for row in &mut table.rows {
        if col < row.cells.len() {
            row.cells.remove(col);
        }
    }
    true
}

pub fn insert_column_left(table: &mut Table, col: usize) {
    insert_column(table, col);
}

pub fn insert_column_right(table: &mut Table, col: usize) {
    insert_column(table, col.saturating_add(1));
}

pub fn merge_cells(table: &mut Table, selection: TableSelection) -> bool {
    let sel = selection.normalized();
    if sel.start.row >= table.rows.len() {
        return false;
    }

    let mut collected = Vec::new();
    for r in sel.start.row..=sel.end.row {
        let Some(row) = table.rows.get_mut(r) else {
            break;
        };

        for c in sel.start.col..=sel.end.col {
            if c >= row.cells.len() {
                continue;
            }

            if r == sel.start.row && c == sel.start.col {
                continue;
            }

            let cell = &mut row.cells[c];
            collected.append(&mut cell.blocks);
            cell.rowspan = 0;
            cell.colspan = 0;
        }
    }

    if let Some(anchor) = table
        .rows
        .get_mut(sel.start.row)
        .and_then(|row| row.cells.get_mut(sel.start.col))
    {
        anchor.rowspan = (sel.end.row - sel.start.row + 1) as u16;
        anchor.colspan = (sel.end.col - sel.start.col + 1) as u16;
        anchor.blocks.extend(collected);
        true
    } else {
        false
    }
}

pub fn split_cell(table: &mut Table, pos: CellPos) -> bool {
    let Some(cell) = table
        .rows
        .get_mut(pos.row)
        .and_then(|row| row.cells.get_mut(pos.col))
    else {
        return false;
    };

    if cell.rowspan <= 1 && cell.colspan <= 1 {
        return false;
    }

    cell.rowspan = 1;
    cell.colspan = 1;
    true
}

pub fn resize_column(table: &mut Table, col: usize, width: f32) -> bool {
    if let Some(slot) = table.column_widths.get_mut(col) {
        *slot = width.max(24.0);
        true
    } else {
        false
    }
}

pub fn resize_row(table: &mut Table, row: usize, height: f32) -> bool {
    ensure_row_heights(table);
    if let Some(slot) = table.row_heights.get_mut(row) {
        *slot = height.max(18.0);
        true
    } else {
        false
    }
}

pub fn distribute_columns_evenly(table: &mut Table, total_width: f32) {
    if table.column_widths.is_empty() {
        return;
    }
    let width = (total_width / table.column_widths.len() as f32).max(24.0);
    for slot in &mut table.column_widths {
        *slot = width;
    }
}

pub fn fit_columns_to_content(table: &mut Table, max_total_width: f32) {
    if table.column_widths.is_empty() {
        return;
    }

    let mut widths = vec![24.0f32; table.column_widths.len()];
    for row in &table.rows {
        for (col, cell) in row.cells.iter().enumerate() {
            if col >= widths.len() {
                break;
            }
            let chars = cell
                .blocks
                .iter()
                .filter_map(|block| match block {
                    Block::Paragraph(p) => Some(
                        p.runs
                            .iter()
                            .map(|run| run.text.chars().count())
                            .sum::<usize>(),
                    ),
                    _ => None,
                })
                .max()
                .unwrap_or(0);
            let estimate = 24.0 + (chars as f32 * 6.2);
            widths[col] = widths[col].max(estimate);
        }
    }

    let sum = widths.iter().sum::<f32>().max(1.0);
    let scale = (max_total_width / sum).clamp(0.2, 1.0);
    for (idx, slot) in widths.iter_mut().enumerate() {
        *slot = (*slot * scale).max(24.0);
        if let Some(column) = table.column_widths.get_mut(idx) {
            *column = *slot;
        }
    }
}

pub fn apply_style(table: &mut Table, style: TableStylePreset) {
    table.style = style.clone();
    match style {
        TableStylePreset::Plain => {
            table.alternating_rows = false;
            table.header_row = false;
        }
        TableStylePreset::Grid => {
            table.alternating_rows = false;
            table.header_row = false;
        }
        TableStylePreset::HeaderAccent => {
            table.header_row = true;
            table.alternating_rows = false;
        }
        TableStylePreset::AlternatingRows => {
            table.alternating_rows = true;
        }
        TableStylePreset::Professional => {
            table.header_row = true;
            table.alternating_rows = true;
            table.cell_padding = 6.0;
        }
    }
}

pub fn visible_row_range(table: &Table, scroll_y: f32, viewport_h: f32, row_h: f32) -> (usize, usize) {
    let start = (scroll_y / row_h).floor().max(0.0) as usize;
    let end = ((scroll_y + viewport_h) / row_h).ceil() as usize;
    (start.min(table.rows.len()), end.min(table.rows.len()))
}

pub fn find_table_mut(doc: &mut DocumentModel, table_id: BlockId) -> Option<&mut Table> {
    doc.content.iter_mut().find_map(|block| match block {
        Block::Table(table) if table.id == table_id => Some(table),
        _ => None,
    })
}

fn next_block_id(doc: &DocumentModel) -> BlockId {
    let max = doc
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(p) => Some(p.id.0),
            Block::Heading(h) => Some(h.id.0),
            Block::Table(t) => Some(t.id.0),
            Block::Image(i) => Some(i.id.0),
            Block::BlockQuote(b) => Some(b.id.0),
            Block::CodeBlock(c) => Some(c.id.0),
            _ => None,
        })
        .max()
        .unwrap_or(0);

    BlockId(max + 1)
}

fn ensure_row_heights(table: &mut Table) {
    if table.row_heights.len() < table.rows.len() {
        table.row_heights.resize(table.rows.len(), 28.0);
    } else if table.row_heights.len() > table.rows.len() {
        table.row_heights.truncate(table.rows.len());
    }
}

#[cfg(test)]
mod tests {
    use crate::document::model::{Paragraph, ParagraphAlignment, ParagraphSpacing, Run, RunStyle};

    use super::*;

    fn paragraph_block(id: u64, text: &str) -> Block {
        Block::Paragraph(Paragraph {
            id: BlockId(id),
            runs: vec![Run {
                text: text.to_string(),
                style: RunStyle::default(),
            }],
            alignment: ParagraphAlignment::Left,
            spacing: ParagraphSpacing::default(),
            indent: Default::default(),
            style_id: None,
        })
    }

    #[test]
    fn insert_table_sets_default_row_heights() {
        let mut doc = DocumentModel::default();
        let table_id = insert_table(&mut doc, 0, 3, 2);
        let table = find_table_mut(&mut doc, table_id).expect("table inserted");
        assert_eq!(table.row_heights.len(), 3);
        assert!(table.row_heights.iter().all(|h| (*h - 28.0).abs() < f32::EPSILON));
    }

    #[test]
    fn row_and_column_insert_delete_update_dimensions() {
        let mut doc = DocumentModel::default();
        let table_id = insert_table(&mut doc, 0, 2, 2);
        let table = find_table_mut(&mut doc, table_id).expect("table inserted");
        insert_row_below(table, 0);
        insert_column_right(table, 0);
        assert_eq!(table.rows.len(), 3);
        assert_eq!(table.row_heights.len(), 3);
        assert_eq!(table.column_widths.len(), 3);
        assert!(delete_row(table, 1));
        assert!(delete_column(table, 1));
        assert_eq!(table.rows.len(), 2);
        assert_eq!(table.row_heights.len(), 2);
        assert_eq!(table.column_widths.len(), 2);
    }

    #[test]
    fn fit_columns_to_content_biases_wider_text() {
        let mut table = Table {
            id: BlockId(1),
            rows: vec![TableRow {
                cells: vec![
                    TableCell {
                        blocks: vec![paragraph_block(10, "short")],
                        ..TableCell::default()
                    },
                    TableCell {
                        blocks: vec![paragraph_block(11, "this is a much longer cell payload")],
                        ..TableCell::default()
                    },
                ],
            }],
            column_widths: vec![120.0, 120.0],
            row_heights: vec![28.0],
            ..Table::default()
        };
        fit_columns_to_content(&mut table, 300.0);
        assert!(table.column_widths[1] > table.column_widths[0]);
        assert!(table.column_widths.iter().sum::<f32>() <= 300.0 + 0.1);
    }

    #[test]
    fn layout_cache_tracks_generation_and_visible_rows() {
        let mut doc = DocumentModel::default();
        let id = insert_table(&mut doc, 0, 20, 3);
        let table = find_table_mut(&mut doc, id).expect("table inserted");
        let mut cache = TableLayoutCache::default();
        let range = cache.visible_rows(table, 30.0, 120.0, 24.0);
        assert!(range.1 >= range.0);
        let before = cache.generation();
        cache.invalidate();
        assert!(cache.generation() > before);
    }
}
