use crate::document::model::{
    Block,
    BlockId,
    DocumentModel,
    Table,
    TableBorders,
    TableCell,
    TableRow,
    TableStylePreset,
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
    table.rows.insert(at.min(table.rows.len()), row);
}

pub fn delete_row(table: &mut Table, row: usize) -> bool {
    if row < table.rows.len() {
        table.rows.remove(row);
        true
    } else {
        false
    }
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
