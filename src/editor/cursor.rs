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
    pub extra_cursors: Vec<CursorPosition>,
}

impl Default for CursorState {
    fn default() -> Self {
        Self {
            primary: CursorPosition::default(),
            selection: None,
            extra_cursors: Vec::new(),
        }
    }
}

impl CursorState {
    pub fn clear_selection(&mut self) {
        self.selection = None;
    }

    pub fn set_selection(&mut self, start: CursorPosition, end: CursorPosition) {
        self.selection = Some(SelectionRange { start, end }.normalized());
    }

    pub fn select_all(&mut self, first_block: BlockId, last_block: BlockId, last_offset: usize) {
        self.selection = Some(
            SelectionRange {
                start: CursorPosition {
                    block_id: first_block,
                    offset: 0,
                },
                end: CursorPosition {
                    block_id: last_block,
                    offset: last_offset,
                },
            }
            .normalized(),
        );
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
    }

    pub fn add_cursor(&mut self, pos: CursorPosition) {
        self.extra_cursors.push(pos);
    }
}
