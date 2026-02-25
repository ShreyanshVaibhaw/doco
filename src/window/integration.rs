use std::{
    ffi::OsString,
    os::windows::ffi::{OsStrExt, OsStringExt},
    path::{Path, PathBuf},
};

use windows::{
    Win32::{
        Foundation::HWND,
        Graphics::Gdi::DeleteDC,
        UI::{
            Accessibility::{HCF_HIGHCONTRASTON, HIGHCONTRASTW},
            Controls::Dialogs::{
                GetOpenFileNameW, GetSaveFileNameW, OFN_EXPLORER, OFN_FILEMUSTEXIST,
                OFN_OVERWRITEPROMPT, OFN_PATHMUSTEXIST, OPENFILENAMEW, PD_NOSELECTION, PD_PAGENUMS,
                PD_RETURNDC, PD_USEDEVMODECOPIESANDCOLLATE, PRINTDLGW, PrintDlgW,
            },
            Shell::{DragFinish, DragQueryFileW, HDROP, SHARD_PATHW, SHAddToRecentDocs},
            WindowsAndMessaging::{
                SPI_GETCLIENTAREAANIMATION, SPI_GETHIGHCONTRAST, SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS,
                SystemParametersInfoW,
            },
        },
    },
    core::w,
};

use crate::ui::AccessibilityPreferences;

pub const SUPPORTED_DOCUMENT_EXTENSIONS: &[&str] = &["docx", "pdf", "txt", "md", "rtf"];
pub const SUPPORTED_IMAGE_EXTENSIONS: &[&str] = &[
    "png", "jpg", "jpeg", "bmp", "gif", "webp", "tif", "tiff", "svg",
];

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

pub fn file_association_registry_commands(exe_path: &Path) -> Vec<String> {
    let mut commands = Vec::new();
    let open_command = explorer_open_with_command(exe_path);
    for ext in file_association_extensions() {
        let prog_id = format!("Doco.{}", ext.to_ascii_uppercase());
        commands.push(format!(
            "reg add HKCU\\Software\\Classes\\.{ext} /ve /d {prog_id} /f"
        ));
        commands.push(format!(
            "reg add HKCU\\Software\\Classes\\{prog_id}\\shell\\open\\command /ve /d \"{open_command}\" /f"
        ));
    }
    commands
}

#[derive(Debug, Clone, Copy, Default)]
pub struct PrintDialogResult {
    pub page_range: Option<(u32, u32)>,
    pub copies: u16,
}

pub fn open_print_dialog(hwnd: HWND) -> Option<PrintDialogResult> {
    let mut dialog = PRINTDLGW {
        lStructSize: std::mem::size_of::<PRINTDLGW>() as u32,
        hwndOwner: hwnd,
        Flags: PD_RETURNDC | PD_USEDEVMODECOPIESANDCOLLATE | PD_NOSELECTION,
        nMinPage: 1,
        nMaxPage: u16::MAX,
        nFromPage: 1,
        nToPage: 1,
        ..Default::default()
    };

    let ok = unsafe { PrintDlgW(&mut dialog).as_bool() };
    if !ok {
        return None;
    }

    let page_range = if dialog.Flags.0 & PD_PAGENUMS.0 != 0 {
        normalize_page_range(dialog.nFromPage, dialog.nToPage)
    } else {
        None
    };

    if !dialog.hDC.0.is_null() {
        unsafe {
            let _ = DeleteDC(dialog.hDC);
        }
    }

    Some(PrintDialogResult {
        page_range,
        copies: dialog.nCopies.max(1),
    })
}

pub fn send_toast_notification(title: &str, body: &str) {
    // Placeholder implementation: integration point for WinRT toast bridge.
    eprintln!("[toast] {} - {}", title, body);
}

pub fn query_accessibility_preferences() -> AccessibilityPreferences {
    let mut preferences = AccessibilityPreferences::default();

    let mut high_contrast = HIGHCONTRASTW {
        cbSize: std::mem::size_of::<HIGHCONTRASTW>() as u32,
        ..Default::default()
    };
    let high_contrast_ok = unsafe {
        SystemParametersInfoW(
            SPI_GETHIGHCONTRAST,
            high_contrast.cbSize,
            Some((&mut high_contrast as *mut HIGHCONTRASTW).cast()),
            SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS(0),
        )
        .is_ok()
    };
    if high_contrast_ok {
        preferences.high_contrast =
            (high_contrast.dwFlags & HCF_HIGHCONTRASTON) == HCF_HIGHCONTRASTON;
    }

    let mut client_area_animations: i32 = 1;
    let animation_ok = unsafe {
        SystemParametersInfoW(
            SPI_GETCLIENTAREAANIMATION,
            0,
            Some((&mut client_area_animations as *mut i32).cast()),
            SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS(0),
        )
        .is_ok()
    };
    if animation_ok {
        preferences.reduce_motion = client_area_animations == 0;
    }

    preferences
}

fn normalize_page_range(from: u16, to: u16) -> Option<(u32, u32)> {
    if from == 0 && to == 0 {
        return None;
    }
    let start = from.max(1) as u32;
    let end = to.max(from).max(1) as u32;
    Some((start, end))
}

pub fn pick_image_file(hwnd: HWND) -> Option<PathBuf> {
    let mut file_buffer = vec![0u16; 260];
    let mut filter = String::new();
    filter.push_str("Image Files (*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp;*.tif;*.tiff;*.svg)\0");
    filter.push_str("*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp;*.tif;*.tiff;*.svg\0");
    filter.push_str("All Files (*.*)\0*.*\0\0");
    let filter_wide = filter.encode_utf16().collect::<Vec<u16>>();

    let mut open = OPENFILENAMEW {
        lStructSize: std::mem::size_of::<OPENFILENAMEW>() as u32,
        hwndOwner: hwnd,
        lpstrFilter: windows::core::PCWSTR::from_raw(filter_wide.as_ptr()),
        lpstrFile: windows::core::PWSTR(file_buffer.as_mut_ptr()),
        nMaxFile: file_buffer.len() as u32,
        lpstrTitle: w!("Insert Image"),
        Flags: OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST,
        ..Default::default()
    };

    let ok = unsafe { GetOpenFileNameW(&mut open).as_bool() };
    if !ok {
        return None;
    }

    let len = file_buffer
        .iter()
        .position(|c| *c == 0)
        .unwrap_or(file_buffer.len());
    if len == 0 {
        return None;
    }
    let path = OsString::from_wide(&file_buffer[..len]);
    Some(PathBuf::from(path))
}

pub fn pick_open_file(hwnd: HWND) -> Option<PathBuf> {
    let mut file_buffer = vec![0u16; 260];
    let mut filter = String::new();
    filter.push_str("Supported Documents (*.docx;*.txt;*.md;*.rtf;*.pdf)\0");
    filter.push_str("*.docx;*.txt;*.md;*.rtf;*.pdf\0");
    filter.push_str("Word Document (*.docx)\0*.docx\0");
    filter.push_str("Text Document (*.txt)\0*.txt\0");
    filter.push_str("Markdown (*.md)\0*.md\0");
    filter.push_str("PDF (*.pdf)\0*.pdf\0");
    filter.push_str("All Files (*.*)\0*.*\0\0");
    let filter_wide = filter.encode_utf16().collect::<Vec<u16>>();

    let mut open = OPENFILENAMEW {
        lStructSize: std::mem::size_of::<OPENFILENAMEW>() as u32,
        hwndOwner: hwnd,
        lpstrFilter: windows::core::PCWSTR::from_raw(filter_wide.as_ptr()),
        lpstrFile: windows::core::PWSTR(file_buffer.as_mut_ptr()),
        nMaxFile: file_buffer.len() as u32,
        lpstrTitle: w!("Open Document"),
        Flags: OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST,
        ..Default::default()
    };

    let ok = unsafe { GetOpenFileNameW(&mut open).as_bool() };
    if !ok {
        return None;
    }

    let len = file_buffer
        .iter()
        .position(|c| *c == 0)
        .unwrap_or(file_buffer.len());
    if len == 0 {
        return None;
    }

    let path = OsString::from_wide(&file_buffer[..len]);
    Some(PathBuf::from(path))
}

pub fn pick_save_file(
    hwnd: HWND,
    suggested_name: &str,
    suggested_extension: &str,
) -> Option<PathBuf> {
    let mut file_buffer = vec![0u16; 260];
    let suggested = if suggested_name.is_empty() {
        format!("Untitled.{suggested_extension}")
    } else {
        suggested_name.to_string()
    };
    let suggested_w = suggested.encode_utf16().collect::<Vec<u16>>();
    let suggested_len = suggested_w.len().min(file_buffer.len().saturating_sub(1));
    file_buffer[..suggested_len].copy_from_slice(&suggested_w[..suggested_len]);
    file_buffer[suggested_len] = 0;

    let mut filter = String::new();
    filter.push_str("Word Document (*.docx)\0*.docx\0");
    filter.push_str("PDF (*.pdf)\0*.pdf\0");
    filter.push_str("Text Document (*.txt)\0*.txt\0");
    filter.push_str("Markdown (*.md)\0*.md\0");
    filter.push_str("HTML (*.html)\0*.html;*.htm\0");
    filter.push_str("RTF (*.rtf)\0*.rtf\0");
    filter.push_str("All Files (*.*)\0*.*\0\0");
    let filter_wide = filter.encode_utf16().collect::<Vec<u16>>();
    let def_ext = suggested_extension
        .trim_start_matches('.')
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect::<Vec<u16>>();

    let mut open = OPENFILENAMEW {
        lStructSize: std::mem::size_of::<OPENFILENAMEW>() as u32,
        hwndOwner: hwnd,
        lpstrFilter: windows::core::PCWSTR::from_raw(filter_wide.as_ptr()),
        lpstrDefExt: windows::core::PCWSTR::from_raw(def_ext.as_ptr()),
        lpstrFile: windows::core::PWSTR(file_buffer.as_mut_ptr()),
        nMaxFile: file_buffer.len() as u32,
        lpstrTitle: w!("Save Document"),
        Flags: OFN_EXPLORER | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT,
        ..Default::default()
    };

    let ok = unsafe { GetSaveFileNameW(&mut open).as_bool() };
    if !ok {
        return None;
    }

    let len = file_buffer
        .iter()
        .position(|c| *c == 0)
        .unwrap_or(file_buffer.len());
    if len == 0 {
        return None;
    }
    let path = OsString::from_wide(&file_buffer[..len]);
    Some(PathBuf::from(path))
}

#[cfg(test)]
mod tests {
    use std::path::PathBuf;

    use super::{
        DropAction,
        classify_drop,
        file_association_registry_commands,
        is_image_path,
        normalize_page_range,
        query_accessibility_preferences,
    };

    #[test]
    fn classify_svg_as_image_insert() {
        let files = vec![PathBuf::from("diagram.svg"), PathBuf::from("photo.png")];
        let action = classify_drop(files.as_slice());
        assert_eq!(action, DropAction::InsertImage);
        assert!(is_image_path(PathBuf::from("icon.svg").as_path()));
    }

    #[test]
    fn file_association_commands_cover_core_extensions() {
        let commands = file_association_registry_commands(PathBuf::from("C:\\Apps\\doco.exe").as_path());
        assert!(commands.iter().any(|line| line.contains(".docx")));
        assert!(commands.iter().any(|line| line.contains(".pdf")));
        assert!(commands.iter().any(|line| line.contains("%1")));
    }

    #[test]
    fn normalize_range_orders_bounds() {
        assert_eq!(normalize_page_range(0, 0), None);
        assert_eq!(normalize_page_range(3, 1), Some((3, 3)));
        assert_eq!(normalize_page_range(2, 6), Some((2, 6)));
    }

    #[test]
    fn accessibility_query_returns_a_valid_struct() {
        let prefs = query_accessibility_preferences();
        assert!(matches!(prefs.high_contrast, true | false));
        assert!(matches!(prefs.reduce_motion, true | false));
    }
}
