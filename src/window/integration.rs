use std::{
    ffi::OsString,
    os::windows::ffi::{OsStrExt, OsStringExt},
    path::{Path, PathBuf},
};

use windows::Win32::UI::Shell::{DragFinish, DragQueryFileW, HDROP, SHARD_PATHW, SHAddToRecentDocs};

pub const SUPPORTED_DOCUMENT_EXTENSIONS: &[&str] = &["docx", "pdf", "txt", "md", "rtf"];
pub const SUPPORTED_IMAGE_EXTENSIONS: &[&str] = &["png", "jpg", "jpeg", "bmp", "gif", "webp", "tif", "tiff"];

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DropAction {
    OpenFilesInTabs,
    InsertImage,
    Ignore,
}

#[derive(Debug, Clone)]
pub struct DropPayload {
    pub files: Vec<PathBuf>,
    pub action: DropAction,
}

#[derive(Debug, Default)]
pub struct JumpListState {
    pub recent_files: Vec<PathBuf>,
    pub pinned_tasks: Vec<String>,
}

impl JumpListState {
    pub fn with_default_tasks() -> Self {
        Self {
            recent_files: Vec::new(),
            pinned_tasks: vec!["New Document".to_string(), "Open File".to_string()],
        }
    }

    pub fn add_recent_file(&mut self, path: PathBuf) {
        self.recent_files.retain(|existing| existing != &path);
        self.recent_files.insert(0, path.clone());
        self.recent_files.truncate(20);

        // Registers with Windows shell recent-docs list (backing Jump List source).
        let wide = path
            .as_os_str()
            .encode_wide()
            .chain(std::iter::once(0))
            .collect::<Vec<u16>>();
        unsafe {
            SHAddToRecentDocs(SHARD_PATHW.0 as u32, Some(wide.as_ptr().cast()));
        }
    }
}

#[derive(Debug, Default)]
pub struct PrintState {
    pub pending: bool,
    pub page_range: Option<(u32, u32)>,
    pub include_header_footer: bool,
}

impl PrintState {
    pub fn request_print_dialog(&mut self) {
        self.pending = true;
    }

    pub fn complete_print(&mut self) {
        self.pending = false;
    }
}

pub fn parse_startup_files_from_cli() -> Vec<PathBuf> {
    std::env::args_os()
        .skip(1)
        .map(PathBuf::from)
        .filter(|path| is_supported_path(path))
        .collect()
}

pub unsafe fn extract_drop_payload(hdrop: HDROP) -> DropPayload {
    let count = unsafe { DragQueryFileW(hdrop, u32::MAX, None) };
    let mut files = Vec::with_capacity(count as usize);

    for index in 0..count {
        let required = unsafe { DragQueryFileW(hdrop, index, None) } + 1;
        if required <= 1 {
            continue;
        }

        let mut buffer = vec![0u16; required as usize];
        let written = unsafe { DragQueryFileW(hdrop, index, Some(buffer.as_mut_slice())) };
        if written == 0 {
            continue;
        }

        let path = OsString::from_wide(&buffer[..written as usize]);
        files.push(PathBuf::from(path));
    }

    unsafe { DragFinish(hdrop) };

    let action = classify_drop(files.as_slice());
    DropPayload { files, action }
}

pub fn classify_drop(files: &[PathBuf]) -> DropAction {
    if files.is_empty() {
        return DropAction::Ignore;
    }

    if files.iter().all(|path| is_image_path(path)) {
        return DropAction::InsertImage;
    }

    if files.iter().all(|path| is_supported_path(path)) {
        return DropAction::OpenFilesInTabs;
    }

    DropAction::Ignore
}

pub fn is_supported_path(path: &Path) -> bool {
    let ext = path
        .extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_ascii_lowercase());

    match ext {
        Some(ext) => {
            SUPPORTED_DOCUMENT_EXTENSIONS.contains(&ext.as_str())
                || SUPPORTED_IMAGE_EXTENSIONS.contains(&ext.as_str())
        }
        None => false,
    }
}

pub fn is_image_path(path: &Path) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_ascii_lowercase())
        .map(|ext| SUPPORTED_IMAGE_EXTENSIONS.contains(&ext.as_str()))
        .unwrap_or(false)
}

pub fn file_association_extensions() -> Vec<&'static str> {
    let mut all = SUPPORTED_DOCUMENT_EXTENSIONS.to_vec();
    all.sort_unstable();
    all
}

pub fn explorer_open_with_command(exe_path: &Path) -> String {
    format!("\"{}\" \"%1\"", exe_path.display())
}

pub fn send_toast_notification(title: &str, body: &str) {
    // Placeholder implementation: integration point for WinRT toast bridge.
    eprintln!("[toast] {} - {}", title, body);
}
