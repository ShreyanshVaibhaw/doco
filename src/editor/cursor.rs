use crate::document::model::BlockId;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CursorPosition {
    pub block_id: BlockId,
    pub offset: usize,
}

impl Default for CursorPosition {
    fn default() -> Self {
        Self {
            block_id: BlockId(1),
            offset: 0,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SelectionRange {
    pub start: CursorPosition,
    pub end: CursorPosition,
}

impl SelectionRange {
    pub fn normalized(self) -> Self {
        if (self.start.block_id.0, self.start.offset) <= (self.end.block_id.0, self.end.offset) {
            self
        } else {
            Self {
                start: self.end,
                end: self.start,
            }
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Movement {
    Left,
    Right,
    CtrlLeft,
    CtrlRight,
    Home,
    End,
    CtrlHome,
    CtrlEnd,
    Up,
    Down,
    PageUp,
    PageDown,
}

#[derive(Debug, Clone)]
pub struct CursorState {
    pub primary: CursorPosition,
    pub selection: Option<SelectionRange>,
    selection_anchor: Option<CursorPosition>,
    pub extra_cursors: Vec<CursorPosition>,
}

impl Default for CursorState {
    fn default() -> Self {
        Self {
            primary: CursorPosition::default(),
            selection: None,
            selection_anchor: None,
            extra_cursors: Vec::new(),
        }
    }
}

impl CursorState {
    pub fn clear_selection(&mut self) {
        self.selection = None;
        self.selection_anchor = None;
    }

    pub fn set_selection(&mut self, start: CursorPosition, end: CursorPosition) {
        self.selection = Some(SelectionRange { start, end }.normalized());
        self.selection_anchor = Some(start);
    }

    pub fn select_all(&mut self, first_block: BlockId, last_block: BlockId, last_offset: usize) {
        let start = CursorPosition {
            block_id: first_block,
            offset: 0,
        };
        let end = CursorPosition {
            block_id: last_block,
            offset: last_offset,
        };
        self.selection = Some(SelectionRange { start, end }.normalized());
        self.selection_anchor = Some(start);
    }

    pub fn move_simple(&mut self, movement: Movement, current_block_len: usize) {
        match movement {
            Movement::Left => {
                self.primary.offset = self.primary.offset.saturating_sub(1);
            }
            Movement::Right => {
                self.primary.offset = (self.primary.offset + 1).min(current_block_len);
            }
            Movement::Home | Movement::CtrlHome => {
                self.primary.offset = 0;
            }
            Movement::End | Movement::CtrlEnd => {
                self.primary.offset = current_block_len;
            }
            Movement::CtrlLeft => {
                self.primary.offset = self.primary.offset.saturating_sub(5);
            }
            Movement::CtrlRight => {
                self.primary.offset = (self.primary.offset + 5).min(current_block_len);
            }
            Movement::Up | Movement::Down | Movement::PageUp | Movement::PageDown => {}
        }
        self.selection_anchor = None;
    }

    pub fn move_in_text(&mut self, movement: Movement, text: &str, extend_selection: bool) {
        let old = self.primary;
        let chars: Vec<char> = text.chars().collect();
        let len = chars.len();
        let mut offset = self.primary.offset.min(len);

        match movement {
            Movement::Left => offset = offset.saturating_sub(1),
            Movement::Right => offset = (offset + 1).min(len),
            Movement::Home | Movement::CtrlHome => offset = 0,
            Movement::End | Movement::CtrlEnd => offset = len,
            Movement::CtrlLeft => {
                while offset > 0 && chars[offset - 1].is_whitespace() {
                    offset -= 1;
                }
                while offset > 0 && !chars[offset - 1].is_whitespace() {
                    offset -= 1;
                }
            }
            Movement::CtrlRight => {
                while offset < len && chars[offset].is_whitespace() {
                    offset += 1;
                }
                while offset < len && !chars[offset].is_whitespace() {
                    offset += 1;
                }
            }
            Movement::Up | Movement::Down | Movement::PageUp | Movement::PageDown => {}
        }

        self.primary.offset = offset;
        self.update_selection(old, extend_selection);
    }

    pub fn add_cursor(&mut self, pos: CursorPosition) {
        if !self.extra_cursors.contains(&pos) && self.primary != pos {
            self.extra_cursors.push(pos);
        }
    }

    pub fn move_across_blocks(
        &mut self,
        movement: Movement,
        blocks: &[(BlockId, usize)],
        viewport_lines: usize,
        extend_selection: bool,
    ) {
        if blocks.is_empty() {
            return;
        }

        let (mut idx, mut len) = blocks
            .iter()
            .enumerate()
            .find_map(|(i, (id, l))| (*id == self.primary.block_id).then_some((i, *l)))
            .unwrap_or((0, blocks[0].1));

        let mut offset = self.primary.offset.min(len);
        let old = self.primary;

        match movement {
            Movement::Left => {
                if offset > 0 {
                    offset -= 1;
                } else if idx > 0 {
                    idx -= 1;
                    len = blocks[idx].1;
                    offset = len;
                }
            }
            Movement::Right => {
                if offset < len {
                    offset += 1;
                } else if idx + 1 < blocks.len() {
                    idx += 1;
                    len = blocks[idx].1;
                    offset = 0;
                }
            }
            Movement::CtrlLeft => {
                offset = word_left(offset, len);
            }
            Movement::CtrlRight => {
                offset = word_right(offset, len);
            }
            Movement::Home => offset = 0,
            Movement::End => offset = len,
            Movement::CtrlHome => {
                idx = 0;
                len = blocks[idx].1;
                offset = 0;
            }
            Movement::CtrlEnd => {
                idx = blocks.len() - 1;
                len = blocks[idx].1;
                offset = len;
            }
            Movement::Up => {
                if idx > 0 {
                    idx -= 1;
                    len = blocks[idx].1;
                    offset = offset.min(len);
                }
            }
            Movement::Down => {
                if idx + 1 < blocks.len() {
                    idx += 1;
                    len = blocks[idx].1;
                    offset = offset.min(len);
                }
            }
            Movement::PageUp => {
                let jump = viewport_lines.max(1);
                idx = idx.saturating_sub(jump);
                len = blocks[idx].1;
                offset = offset.min(len);
            }
            Movement::PageDown => {
                let jump = viewport_lines.max(1);
                idx = (idx + jump).min(blocks.len() - 1);
                len = blocks[idx].1;
                offset = offset.min(len);
            }
        }

        self.primary = CursorPosition {
            block_id: blocks[idx].0,
            offset,
        };
        self.update_selection(old, extend_selection);
    }

    pub fn select_word(&mut self, block_id: BlockId, text: &str, offset: usize) {
        let chars: Vec<char> = text.chars().collect();
        if chars.is_empty() {
            self.clear_selection();
            return;
        }

        let mut start = offset.min(chars.len());
        while start > 0 && !chars[start.saturating_sub(1)].is_whitespace() {
            start -= 1;
        }

        let mut end = offset.min(chars.len());
        while end < chars.len() && !chars[end].is_whitespace() {
            end += 1;
        }

        self.set_selection(
            CursorPosition { block_id, offset: start },
            CursorPosition { block_id, offset: end },
        );
        self.primary = CursorPosition { block_id, offset: end };
    }

    pub fn select_paragraph(&mut self, block_id: BlockId, len: usize) {
        self.set_selection(
            CursorPosition {
                block_id,
                offset: 0,
            },
            CursorPosition {
                block_id,
                offset: len,
            },
        );
        self.primary = CursorPosition {
            block_id,
            offset: len,
        };
    }

    pub fn drag_select(&mut self, start: CursorPosition, current: CursorPosition) {
        self.set_selection(start, current);
        self.primary = current;
    }

    pub fn select_next_occurrence(&mut self, pos: CursorPosition) {
        self.add_cursor(pos);
    }

    fn update_selection(&mut self, old: CursorPosition, extend_selection: bool) {
        if extend_selection {
            let anchor = self.selection_anchor.unwrap_or(old);
            self.selection_anchor = Some(anchor);
            self.selection = Some(
                SelectionRange {
                    start: anchor,
                    end: self.primary,
                }
                .normalized(),
            );
        } else {
            self.clear_selection();
        }
    }
}

fn word_left(offset: usize, _len: usize) -> usize {
    if offset == 0 {
        0
    } else {
        offset.saturating_sub(5)
    }
}

fn word_right(offset: usize, len: usize) -> usize {
    (offset + 5).min(len)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn block_navigation_moves_across_lines() {
        let blocks = vec![(BlockId(1), 5), (BlockId(2), 3), (BlockId(3), 7)];
        let mut cursor = CursorState::default();
        cursor.primary = CursorPosition {
            block_id: BlockId(1),
            offset: 5,
        };

        cursor.move_across_blocks(Movement::Right, &blocks, 2, false);
        assert_eq!(cursor.primary.block_id, BlockId(2));
        assert_eq!(cursor.primary.offset, 0);

        cursor.move_across_blocks(Movement::PageDown, &blocks, 2, false);
        assert_eq!(cursor.primary.block_id, BlockId(3));
    }

    #[test]
    fn selection_and_multicursor_work() {
        let mut cursor = CursorState::default();
        cursor.select_word(BlockId(9), "hello world", 1);
        let sel = cursor.selection.expect("selection expected");
        assert_eq!(sel.start.offset, 0);
        assert_eq!(sel.end.offset, 5);

        cursor.primary.offset = 11;
        cursor.move_in_text(Movement::CtrlLeft, "hello world", false);
        assert_eq!(cursor.primary.offset, 6);

        cursor.select_next_occurrence(CursorPosition {
            block_id: BlockId(9),
            offset: 7,
        });
        assert_eq!(cursor.extra_cursors.len(), 1);
    }
}
