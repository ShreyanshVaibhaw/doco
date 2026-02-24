use std::{collections::VecDeque, time::Instant};

use crate::editor::commands::EditCommand;

const COALESCE_WINDOW_MS: u128 = 500;
const MAX_UNDO_STEPS: usize = 1000;
const MAX_UNDO_BYTES: usize = 10 * 1024 * 1024;

#[derive(Debug, Clone)]
pub struct UndoEntry {
    pub command: EditCommand,
    pub inverse: EditCommand,
    pub bytes: usize,
    pub timestamp: Instant,
}

#[derive(Debug)]
pub struct UndoStack {
    undo: VecDeque<UndoEntry>,
    redo: VecDeque<UndoEntry>,
    used_bytes: usize,
    max_steps: usize,
    max_bytes: usize,
}

impl UndoStack {
    pub fn with_limits(max_steps: usize, max_bytes: usize) -> Self {
        Self {
            undo: VecDeque::new(),
            redo: VecDeque::new(),
            used_bytes: 0,
            max_steps: max_steps.max(32),
            max_bytes: max_bytes.max(64 * 1024),
        }
    }

    pub fn set_limits(&mut self, max_steps: usize, max_bytes: usize) {
        self.max_steps = max_steps.max(32);
        self.max_bytes = max_bytes.max(64 * 1024);
        self.enforce_limits();
    }

    pub fn push(&mut self, mut entry: UndoEntry) {
        if self.try_coalesce(&mut entry) {
            return;
        }

        self.used_bytes += entry.bytes;
        self.undo.push_back(entry);
        self.redo.clear();
        self.enforce_limits();
    }

    pub fn pop_undo(&mut self) -> Option<UndoEntry> {
        let entry = self.undo.pop_back()?;
        self.used_bytes = self.used_bytes.saturating_sub(entry.bytes);
        self.redo.push_back(entry.clone());
        Some(entry)
    }

    pub fn pop_redo(&mut self) -> Option<UndoEntry> {
        let entry = self.redo.pop_back()?;
        self.used_bytes += entry.bytes;
        self.undo.push_back(entry.clone());
        Some(entry)
    }

    pub fn undo_len(&self) -> usize {
        self.undo.len()
    }

    pub fn redo_len(&self) -> usize {
        self.redo.len()
    }

    pub fn used_bytes(&self) -> usize {
        self.used_bytes
    }

    fn try_coalesce(&mut self, next: &mut UndoEntry) -> bool {
        let Some(last) = self.undo.back_mut() else {
            return false;
        };

        let elapsed = next.timestamp.saturating_duration_since(last.timestamp).as_millis();
        if elapsed > COALESCE_WINDOW_MS {
            return false;
        }

        match (&mut last.command, &next.command) {
            (
                EditCommand::InsertText {
                    block_id: b1,
                    offset: o1,
                    text: t1,
                },
                EditCommand::InsertText {
                    block_id: b2,
                    offset: o2,
                    text: t2,
                },
            ) if b1 == b2 && *o2 == (*o1 + t1.chars().count()) => {
                if t2.chars().all(is_coalescable_char) {
                    t1.push_str(t2);
                    last.bytes += next.bytes;
                    last.timestamp = next.timestamp;
                    self.used_bytes += next.bytes;
                    self.enforce_limits();
                    return true;
                }
            }
            _ => {}
        }

        false
    }

    fn enforce_limits(&mut self) {
        while self.undo.len() > self.max_steps || self.used_bytes > self.max_bytes {
            if let Some(front) = self.undo.pop_front() {
                self.used_bytes = self.used_bytes.saturating_sub(front.bytes);
            } else {
                break;
            }
        }

        // Compress older entries heuristically to keep memory bounded.
        if self.undo.len() > self.max_steps / 2 {
            let compress_upto = self.undo.len() / 3;
            for (idx, item) in self.undo.iter_mut().enumerate() {
                if idx < compress_upto {
                    let before = item.bytes;
                    item.bytes = (item.bytes as f32 * 0.85) as usize;
                    self.used_bytes = self.used_bytes.saturating_sub(before.saturating_sub(item.bytes));
                }
            }
        }
    }
}

impl Default for UndoStack {
    fn default() -> Self {
        Self::with_limits(MAX_UNDO_STEPS, MAX_UNDO_BYTES)
    }
}

fn is_coalescable_char(ch: char) -> bool {
    ch.is_ascii_alphanumeric() || ch == '_'
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::document::model::BlockId;

    fn insert_cmd(text: &str, offset: usize) -> EditCommand {
        EditCommand::InsertText {
            block_id: BlockId(1),
            offset,
            text: text.to_string(),
        }
    }

    #[test]
    fn coalesces_typing_but_breaks_on_space() {
        let now = Instant::now();
        let mut stack = UndoStack::default();

        stack.push(UndoEntry {
            command: insert_cmd("h", 0),
            inverse: EditCommand::DeleteText {
                block_id: BlockId(1),
                start: 0,
                end: 1,
            },
            bytes: 1,
            timestamp: now,
        });
        stack.push(UndoEntry {
            command: insert_cmd("i", 1),
            inverse: EditCommand::DeleteText {
                block_id: BlockId(1),
                start: 1,
                end: 2,
            },
            bytes: 1,
            timestamp: now,
        });
        assert_eq!(stack.undo_len(), 1);

        stack.push(UndoEntry {
            command: insert_cmd(" ", 2),
            inverse: EditCommand::DeleteText {
                block_id: BlockId(1),
                start: 2,
                end: 3,
            },
            bytes: 1,
            timestamp: now,
        });
        assert_eq!(stack.undo_len(), 2);
    }

    #[test]
    fn enforces_memory_and_step_limits() {
        let mut stack = UndoStack::with_limits(32, 64 * 1024);
        let now = Instant::now();

        for i in 0..200 {
            stack.push(UndoEntry {
                command: insert_cmd("abcdefghijklmnopqrstuvwxyz", i),
                inverse: EditCommand::DeleteText {
                    block_id: BlockId(1),
                    start: i,
                    end: i + 26,
                },
                bytes: 4096,
                timestamp: now,
            });
        }

        assert!(stack.undo_len() <= 32);
        assert!(stack.used_bytes() <= 64 * 1024);
    }
}
