use std::{
    ffi::c_void,
    mem::size_of,
    path::{Path, PathBuf},
    time::{Duration, Instant},
};

use windows::{
    Win32::{
        Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, RECT, WPARAM},
        Graphics::{
            Dwm::{
                DWMSBT_MAINWINDOW, DWMWA_SYSTEMBACKDROP_TYPE, DWMWA_USE_IMMERSIVE_DARK_MODE,
                DwmSetWindowAttribute,
            },
            Gdi::{BeginPaint, EndPaint, InvalidateRect, PAINTSTRUCT},
        },
        System::LibraryLoader::GetModuleHandleW,
        UI::{
            HiDpi::{DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2, SetProcessDpiAwarenessContext},
            Input::KeyboardAndMouse::{
                GetKeyState, ReleaseCapture, SetCapture, VK_CONTROL, VK_DELETE, VK_SHIFT,
            },
            Shell::{DragAcceptFiles, HDROP},
            WindowsAndMessaging::{
                AdjustWindowRectEx, CREATESTRUCTW, CS_DBLCLKS, CS_HREDRAW, CS_VREDRAW,
                CreateWindowExW, DefWindowProcW, DispatchMessageW, GWLP_USERDATA, GetClientRect,
                GetMessageW, GetSystemMetrics, GetWindowLongPtrW, IDC_ARROW, LoadCursorW, MSG,
                IDCANCEL, IDNO, IDYES, MB_ICONWARNING, MB_YESNOCANCEL, MessageBoxW,
                PostQuitMessage, RegisterClassExW, SM_CXSCREEN, SM_CYSCREEN, SW_SHOW,
                SWP_NOACTIVATE, SWP_NOZORDER, SetWindowLongPtrW, SetWindowPos, ShowWindow,
                TranslateMessage, WINDOW_EX_STYLE, WM_CHAR, WM_CREATE, WM_DESTROY, WM_DPICHANGED,
                WM_DROPFILES, WM_KEYDOWN, WM_LBUTTONDBLCLK, WM_LBUTTONDOWN, WM_LBUTTONUP,
                WM_MBUTTONDOWN, WM_MOUSEMOVE, WM_MOUSEWHEEL, WM_NCCREATE, WM_NCDESTROY,
                WM_PAINT, WM_SETTINGCHANGE, WM_SIZE, WNDCLASSEXW, WS_OVERLAPPEDWINDOW, WS_VISIBLE,
            },
        },
    },
    core::{PCWSTR, Result, w},
};

use crate::{
    app::AppState,
    document::{
        DocumentFormat, detect_format,
        docx::parser::parse_docx,
        export::{export_pdf, save_with_format},
        markdown::MarkdownDocument,
        model::{
            Block, BlockId, DocumentModel, ImageAlignment, ImageBorder, ImageBorderStyle,
            TableStylePreset,
        },
        txt::TextDocument,
    },
    editor::{
        clipboard::read_clipboard_image,
        image_ops::load_supported_image,
        search::{FindReplaceState, replace_all, replace_current, replacement_preview},
        table::{
            CellPos,
            TableSelection,
            apply_style as apply_table_style,
            delete_column as delete_table_column,
            delete_row as delete_table_row,
            distribute_columns_evenly,
            fit_columns_to_content,
            find_table_mut,
            insert_column_left,
            insert_column_right,
            insert_row_above,
            insert_row_below,
            insert_table,
            merge_cells as merge_table_cells,
            resize_column as resize_table_column,
            resize_row as resize_table_row,
            split_cell as split_table_cell,
            visible_row_range,
        },
    },
    render::canvas::PageLayoutMode,
    render::d2d::{D2DRenderer, ShellRenderState},
    render::image_cache::{ImageDecodeCache, interpolation_hint, resolve_image_data},
    render::perf::emit_startup_marker,
    settings::schema::{Settings, SettingsCategory, SidebarDefaultPanel},
    theme::{
        Theme, ThemeManager,
        backgrounds::{BackgroundKind, from_canvas_preference},
    },
    ui::{
        InputEvent as UiInputEvent, Point as UiPoint, Rect as UiRect, UIComponent,
        command_palette::CommandPalette,
        dialog::Dialog,
        sidebar::{SearchResultItem, Sidebar, SidebarIntent, SidebarPanel},
        statusbar::{StatusAction, StatusBar, StatusBarInfo},
        tabs::{TabKind, TabsBar},
        toolbar::{Toolbar, ToolbarAction},
    },
    window::integration::{
        DropAction, JumpListState, PrintState, extract_drop_payload, parse_startup_files_from_cli,
        pick_image_file, pick_save_file, open_print_dialog, send_toast_notification,
    },
};

pub mod compositor;
pub mod input;
pub mod integration;

pub struct AppWindow {
    hwnd: HWND,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FindFieldFocus {
    Query,
    Replacement,
}

#[derive(Debug, Clone)]
struct CanvasImageOverlay {
    block_id: BlockId,
    rect: UiRect,
    interpolation: String,
    alt_text: String,
}

#[derive(Debug, Clone)]
struct CanvasTableOverlay {
    table_id: BlockId,
    rect: UiRect,
    rows: usize,
    cols: usize,
    cell_w: f32,
    cell_h: f32,
    header_h: f32,
    gutter_w: f32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ImageDragKind {
    Move,
    CornerResize,
    EdgeResizeHorizontal,
    EdgeResizeVertical,
}

#[derive(Debug, Clone)]
struct ImageDragState {
    block_id: BlockId,
    start_mouse: UiPoint,
    start_width: f32,
    start_height: f32,
    start_alignment: ImageAlignment,
    kind: ImageDragKind,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TableSelectionMode {
    Cell(CellPos),
    Row(usize),
    Column(usize),
    Table,
}

#[derive(Debug, Clone)]
struct TableResizeState {
    table_id: BlockId,
    row: Option<usize>,
    col: Option<usize>,
    start_mouse: UiPoint,
    start_value: f32,
}

#[derive(Debug, Clone, Copy)]
struct TablePickerLayout {
    panel: UiRect,
    grid: UiRect,
    rows_input: UiRect,
    cols_input: UiRect,
    insert_button: UiRect,
}

struct WindowState {
    renderer: Option<D2DRenderer>,
    dpi: f32,
    theme: Theme,
    theme_manager: ThemeManager,
    debug_panel_visible: bool,
    dropped_files: Vec<PathBuf>,
    jump_list: JumpListState,
    print_state: PrintState,
    startup_files: Vec<PathBuf>,
    app_state: AppState,
    tabs: TabsBar,
    sidebar: Sidebar,
    settings_dialog: Dialog,
    command_palette: CommandPalette,
    find_replace: FindReplaceState,
    find_focus: FindFieldFocus,
    image_cache: ImageDecodeCache,
    canvas_image_overlays: Vec<CanvasImageOverlay>,
    selected_image: Option<BlockId>,
    image_drag: Option<ImageDragState>,
    image_properties_visible: bool,
    table_picker_visible: bool,
    table_picker_rows: usize,
    table_picker_cols: usize,
    table_picker_custom_rows: String,
    table_picker_custom_cols: String,
    table_picker_custom_focus_rows: bool,
    canvas_table_overlays: Vec<CanvasTableOverlay>,
    selected_table: Option<BlockId>,
    table_selection_mode: Option<TableSelectionMode>,
    table_selection_range: Option<TableSelection>,
    table_resize: Option<TableResizeState>,
    goto_visible: bool,
    goto_input: String,
    toolbar: Toolbar,
    statusbar: StatusBar,
    last_ui_tick: Instant,
    sidebar_resizing: bool,
    sidebar_resize_grab_offset: f32,
}

impl AppWindow {
    pub fn new(theme_manager: ThemeManager, settings: Settings) -> Result<Self> {
        unsafe {
            let _ = SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
        }

        let hmodule = unsafe { GetModuleHandleW(None)? };
        let hinstance = HINSTANCE(hmodule.0);
        let class_name = w!("DocoMainWindow");

        let wc = WNDCLASSEXW {
            cbSize: size_of::<WNDCLASSEXW>() as u32,
            style: CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS,
            lpfnWndProc: Some(window_proc),
            hInstance: hinstance,
            hCursor: unsafe { LoadCursorW(None, IDC_ARROW)? },
            lpszClassName: class_name,
            ..Default::default()
        };

        unsafe {
            let _ = RegisterClassExW(&wc);
        }

        let mut rect = RECT {
            left: 0,
            top: 0,
            right: 1200,
            bottom: 800,
        };

        unsafe {
            AdjustWindowRectEx(&mut rect, WS_OVERLAPPEDWINDOW, false, WINDOW_EX_STYLE(0))?;
        }

        let width = rect.right - rect.left;
        let height = rect.bottom - rect.top;
        let x = (unsafe { GetSystemMetrics(SM_CXSCREEN) } - width).max(0) / 2;
        let y = (unsafe { GetSystemMetrics(SM_CYSCREEN) } - height).max(0) / 2;
        let theme = theme_manager.active();
        let mut app_state = AppState::default();
        app_state.settings = settings.clone();
        app_state.show_toolbar = settings.appearance.show_toolbar;
        app_state.show_sidebar = settings.appearance.show_sidebar;
        app_state.show_statusbar = settings.appearance.show_status_bar;
        app_state.show_tabs = settings.appearance.show_tab_bar;
        let mut sidebar = Sidebar::default();
        sidebar.set_active_panel(match settings.appearance.sidebar_default_panel {
            SidebarDefaultPanel::Files => SidebarPanel::Files,
            SidebarDefaultPanel::Outline => SidebarPanel::Outline,
            SidebarDefaultPanel::Bookmarks => SidebarPanel::Bookmarks,
        });

        let state = Box::new(WindowState {
            renderer: None,
            dpi: 96.0,
            theme,
            theme_manager,
            debug_panel_visible: false,
            dropped_files: Vec::new(),
            jump_list: JumpListState::with_default_tasks(),
            print_state: PrintState::default(),
            startup_files: parse_startup_files_from_cli(),
            app_state,
            tabs: TabsBar::default(),
            sidebar,
            settings_dialog: Dialog::default(),
            command_palette: CommandPalette::default(),
            find_replace: FindReplaceState::default(),
            find_focus: FindFieldFocus::Query,
            image_cache: ImageDecodeCache::default(),
            canvas_image_overlays: Vec::new(),
            selected_image: None,
            image_drag: None,
            image_properties_visible: false,
            table_picker_visible: false,
            table_picker_rows: 3,
            table_picker_cols: 3,
            table_picker_custom_rows: "3".to_string(),
            table_picker_custom_cols: "3".to_string(),
            table_picker_custom_focus_rows: true,
            canvas_table_overlays: Vec::new(),
            selected_table: None,
            table_selection_mode: None,
            table_selection_range: None,
            table_resize: None,
            goto_visible: false,
            goto_input: String::new(),
            toolbar: Toolbar::default(),
            statusbar: StatusBar::default(),
            last_ui_tick: Instant::now(),
            sidebar_resizing: false,
            sidebar_resize_grab_offset: 0.0,
        });
        let state_ptr = Box::into_raw(state);

        let hwnd = unsafe {
            CreateWindowExW(
                WINDOW_EX_STYLE(0),
                class_name,
                w!("Doco"),
                WS_OVERLAPPEDWINDOW | WS_VISIBLE,
                x,
                y,
                width,
                height,
                None,
                None,
                Some(hinstance),
                Some(state_ptr as *const c_void),
            )?
        };

        unsafe {
            let _ = ShowWindow(hwnd, SW_SHOW);
        }

        Ok(Self { hwnd })
    }

    pub fn run(self) -> Result<()> {
        let mut message = MSG::default();
        while unsafe { GetMessageW(&mut message, None, 0, 0) }.as_bool() {
            unsafe {
                let _ = TranslateMessage(&message);
                DispatchMessageW(&message);
            }
        }

        Ok(())
    }

    #[allow(dead_code)]
    pub fn hwnd(&self) -> HWND {
        self.hwnd
    }
}

unsafe fn apply_window_effects(hwnd: HWND, is_dark: bool) {
    let dark_mode = windows::core::BOOL(if is_dark { 1 } else { 0 });
    let _ = unsafe {
        DwmSetWindowAttribute(
            hwnd,
            DWMWA_USE_IMMERSIVE_DARK_MODE,
            &dark_mode as *const _ as *const c_void,
            size_of::<windows::core::BOOL>() as u32,
        )
    };

    let backdrop = DWMSBT_MAINWINDOW;
    let _ = unsafe {
        DwmSetWindowAttribute(
            hwnd,
            DWMWA_SYSTEMBACKDROP_TYPE,
            &backdrop as *const _ as *const c_void,
            size_of::<windows::Win32::Graphics::Dwm::DWM_SYSTEMBACKDROP_TYPE>() as u32,
        )
    };
}

unsafe fn state_from_hwnd(hwnd: HWND) -> Option<&'static mut WindowState> {
    let ptr = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut WindowState;
    if ptr.is_null() {
        None
    } else {
        Some(unsafe { &mut *ptr })
    }
}

fn point_from_lparam(lparam: LPARAM) -> UiPoint {
    let raw = lparam.0 as u32;
    let x = (raw & 0xFFFF) as i16 as f32;
    let y = ((raw >> 16) & 0xFFFF) as i16 as f32;
    UiPoint { x, y }
}

fn sidebar_splitter_hit_test(state: &WindowState, point: UiPoint) -> bool {
    if !state.app_state.show_sidebar {
        return false;
    }
    let bounds = state.sidebar.bounds();
    if bounds.width <= 0.0 || bounds.height <= 0.0 {
        return false;
    }

    let splitter_x = bounds.x + bounds.width;
    let half_width = 4.0;
    point.x >= splitter_x - half_width
        && point.x <= splitter_x + half_width
        && point.y >= bounds.y
        && point.y <= bounds.y + bounds.height
}

fn document_title_from_path(path: &Path) -> String {
    path.file_name()
        .and_then(|v| v.to_str())
        .unwrap_or("Document")
        .to_string()
}

fn load_document_for_path(path: &Path) -> DocumentModel {
    let detected = detect_format(path);
    let mut model = match detected {
        DocumentFormat::Docx => parse_docx(path).unwrap_or_default(),
        DocumentFormat::Markdown => MarkdownDocument::load_from_path(path)
            .map(|doc| doc.to_document_model())
            .unwrap_or_default(),
        DocumentFormat::Text => TextDocument::load_from_path(path)
            .map(|doc| doc.to_document_model())
            .unwrap_or_default(),
        DocumentFormat::Pdf | DocumentFormat::Unknown => DocumentModel::default(),
    };

    model.metadata.file_path = Some(path.to_path_buf());
    if model.metadata.title.is_empty() {
        model.metadata.title = document_title_from_path(path);
    }
    if matches!(model.metadata.format, DocumentFormat::Unknown) {
        model.metadata.format = detected;
    }
    model
}

fn process_startup_file_queue(state: &mut WindowState) -> bool {
    if state.startup_files.is_empty() {
        return false;
    }

    let path = state.startup_files.remove(0);
    let title = document_title_from_path(path.as_path());
    let document = load_document_for_path(path.as_path());
    state
        .tabs
        .open_document_tab(title.clone(), Some(path.clone()), document);
    state.jump_list.add_recent_file(path);

    if state.startup_files.is_empty() {
        state.app_state.status_text = format!("Opened {title}");
    } else {
        state.app_state.status_text = format!(
            "Opening startup files... {} remaining",
            state.startup_files.len()
        );
    }

    true
}

fn default_extension_for_document(state: &WindowState, format: DocumentFormat) -> String {
    let from_settings = state
        .app_state
        .settings
        .files
        .default_save_format
        .trim()
        .trim_start_matches('.')
        .to_ascii_lowercase();
    if !from_settings.is_empty() {
        return from_settings;
    }
    match format {
        DocumentFormat::Docx | DocumentFormat::Unknown => "docx".to_string(),
        DocumentFormat::Pdf => "pdf".to_string(),
        DocumentFormat::Text => "txt".to_string(),
        DocumentFormat::Markdown => "md".to_string(),
    }
}

fn suggested_save_name(tab: &crate::ui::tabs::TabState, default_ext: &str) -> String {
    if let Some(path) = tab
        .file_path
        .as_ref()
        .or(tab.document.metadata.file_path.as_ref())
        && let Some(name) = path.file_name().and_then(|v| v.to_str())
    {
        return name.to_string();
    }
    let base = if !tab.document.metadata.title.trim().is_empty() {
        tab.document.metadata.title.trim().to_string()
    } else {
        "Untitled".to_string()
    };
    format!("{base}.{default_ext}")
}

fn pick_save_target_for_active_tab(
    state: &WindowState,
    hwnd: HWND,
    forced_ext: Option<&str>,
) -> Option<PathBuf> {
    let tab = state.tabs.active_tab()?;
    let default_ext = forced_ext
        .map(|v| v.trim_start_matches('.').to_ascii_lowercase())
        .unwrap_or_else(|| default_extension_for_document(state, tab.document.metadata.format));
    let suggested = suggested_save_name(tab, default_ext.as_str());
    pick_save_file(hwnd, suggested.as_str(), default_ext.as_str())
}

fn path_is_read_only(path: &Path) -> bool {
    std::fs::metadata(path)
        .map(|meta| meta.permissions().readonly())
        .unwrap_or(false)
}

fn is_tab_dirty(tab: &crate::ui::tabs::TabState) -> bool {
    tab.dirty || tab.document.dirty
}

fn to_wide_null(value: &str) -> Vec<u16> {
    let mut wide = value.encode_utf16().collect::<Vec<u16>>();
    wide.push(0);
    wide
}

fn open_new_blank_tab(state: &mut WindowState) -> usize {
    let index = state.tabs.new_blank_tab();
    if state.tabs.tabs.len() > 20 {
        state.app_state.status_text = format!(
            "{} tabs open. Close inactive tabs to reduce memory usage.",
            state.tabs.tabs.len()
        );
    }
    index
}

fn close_tab_with_prompt(state: &mut WindowState, hwnd: HWND, index: usize) -> bool {
    let (dirty, title) = match state.tabs.tabs.get(index) {
        Some(tab) => (is_tab_dirty(tab), tab.title.clone()),
        None => return false,
    };

    if dirty {
        let prompt = format!("Save changes to '{title}' before closing?");
        let prompt_wide = to_wide_null(prompt.as_str());
        let choice = unsafe {
            MessageBoxW(
                Some(hwnd),
                PCWSTR(prompt_wide.as_ptr()),
                w!("Doco"),
                MB_YESNOCANCEL | MB_ICONWARNING,
            )
        };

        if choice == IDCANCEL {
            state.app_state.status_text = "Close cancelled".to_string();
            return true;
        }

        if choice == IDYES {
            state.tabs.set_active(index);
            let _ = save_active_document(state, hwnd, false);
            let still_dirty = state
                .tabs
                .tabs
                .get(index)
                .map(is_tab_dirty)
                .unwrap_or(false);
            if still_dirty {
                state.app_state.status_text =
                    "Close cancelled (document still has unsaved changes)".to_string();
                return true;
            }
        } else if choice != IDNO {
            return true;
        }
    }

    let closed_title = state
        .tabs
        .tabs
        .get(index)
        .map(|tab| tab.title.clone())
        .unwrap_or_else(|| "Tab".to_string());
    if state.tabs.close_tab(index) {
        let active_title = state
            .tabs
            .active_tab()
            .map(|tab| tab.title.clone())
            .unwrap_or_else(|| "Welcome".to_string());
        state.app_state.status_text = format!("Closed {closed_title}. Active: {active_title}");
        sync_sidebar_with_active_tab(state);
        return true;
    }

    false
}

fn save_active_document(state: &mut WindowState, hwnd: HWND, save_as: bool) -> bool {
    let (existing_path, document) = {
        let Some(tab) = state.tabs.active_tab() else {
            state.app_state.status_text = "No active tab to save".to_string();
            return true;
        };
        (
            tab.file_path
                .clone()
                .or_else(|| tab.document.metadata.file_path.clone()),
            tab.document.clone(),
        )
    };

    let target = if !save_as {
        existing_path.or_else(|| pick_save_target_for_active_tab(state, hwnd, None))
    } else {
        pick_save_target_for_active_tab(state, hwnd, None)
    };

    let Some(target) = target else {
        state.app_state.status_text = "Save cancelled".to_string();
        return true;
    };

    if target.exists() && path_is_read_only(target.as_path()) {
        state.app_state.status_text = format!("Save blocked (read-only): {}", target.display());
        return true;
    }

    match save_with_format(target.as_path(), &document) {
        Ok(_) => {
            if let Some(tab) = state.tabs.active_tab_mut() {
                tab.file_path = Some(target.clone());
                tab.title = document_title_from_path(target.as_path());
                tab.document.metadata.file_path = Some(target.clone());
                tab.document.metadata.format = detect_format(target.as_path());
                tab.document.dirty = false;
                tab.dirty = false;
            }
            state.jump_list.add_recent_file(target.clone());
            let _ = state.app_state.autosave.clear_recovery_files();
            state.app_state.status_text = format!("Saved {}", target.display());
            sync_sidebar_with_active_tab(state);
        }
        Err(err) => {
            state.app_state.status_text = format!("Save failed: {err}");
        }
    }
    true
}

fn export_active_document(state: &mut WindowState, hwnd: HWND, ext: &str) -> bool {
    let document = {
        let Some(tab) = state.tabs.active_tab() else {
            state.app_state.status_text = "No active tab to export".to_string();
            return true;
        };
        tab.document.clone()
    };

    let Some(path) = pick_save_target_for_active_tab(state, hwnd, Some(ext)) else {
        state.app_state.status_text = "Export cancelled".to_string();
        return true;
    };

    let result = if ext.eq_ignore_ascii_case("pdf") {
        export_pdf(path.as_path(), &document)
    } else {
        save_with_format(path.as_path(), &document)
    };

    match result {
        Ok(_) => {
            state.app_state.status_text = format!("Exported {}", path.display());
            send_toast_notification(
                "Export complete",
                format!("{}", path.display()).as_str(),
            );
        }
        Err(err) => {
            state.app_state.status_text = format!("Export failed: {err}");
        }
    }
    true
}

fn restore_recovery_tabs(state: &mut WindowState) -> usize {
    let recovery_files = state
        .app_state
        .autosave
        .list_recovery_files()
        .unwrap_or_default();
    let mut restored = 0usize;
    for recovery in recovery_files {
        let bytes = match std::fs::read(&recovery) {
            Ok(bytes) => bytes,
            Err(_) => continue,
        };
        let mut document = match serde_json::from_slice::<DocumentModel>(&bytes) {
            Ok(model) => model,
            Err(_) => continue,
        };
        document.metadata.file_path = None;
        document.dirty = true;
        let title = recovery
            .file_stem()
            .and_then(|v| v.to_str())
            .map(|v| format!("Recovered ({v})"))
            .unwrap_or_else(|| "Recovered".to_string());
        state.tabs.open_document_tab(title, None, document);
        restored += 1;
    }
    restored
}

fn sync_sidebar_with_active_tab(state: &mut WindowState) {
    let mut root_path = None;
    if let Some(tab) = state.tabs.active_tab() {
        state.sidebar.populate_outline(&tab.document);
        state
            .sidebar
            .set_current_outline_block(Some(tab.cursor.primary.block_id));
        root_path = tab
            .file_path
            .clone()
            .or_else(|| tab.document.metadata.file_path.clone());
    }

    if let Some(selected) = state.selected_image {
        if active_image_ref(state, selected).is_none() {
            state.selected_image = None;
            state.image_drag = None;
            state.image_properties_visible = false;
        }
    }
    if let Some(selected) = state.selected_table {
        if active_table_ref(state, selected).is_none() {
            state.selected_table = None;
            state.table_selection_mode = None;
            state.table_selection_range = None;
            state.table_resize = None;
        }
    }

    if root_path.is_none() {
        root_path = std::env::current_dir().ok();
    }

    if let Some(path) = root_path {
        let root = if path.is_dir() {
            path
        } else {
            path.parent().map(Path::to_path_buf).unwrap_or(path)
        };
        if state.sidebar.file_root.as_ref() != Some(&root) {
            let _ = state.sidebar.open_folder(root);
        }
    }
}

fn open_path_from_sidebar(state: &mut WindowState, path: PathBuf, new_tab: bool) {
    let title = document_title_from_path(&path);
    let document = load_document_for_path(&path);
    if new_tab {
        state
            .tabs
            .open_document_tab(title.clone(), Some(path), document);
    } else if let Some(tab) = state.tabs.active_tab_mut() {
        tab.title = title.clone();
        tab.kind = TabKind::Document;
        tab.file_path = Some(path);
        tab.document = document;
        tab.cursor = Default::default();
        tab.canvas = Default::default();
        tab.dirty = false;
    } else {
        state
            .tabs
            .open_document_tab(title.clone(), Some(path), document);
    }
    state.app_state.status_text = format!("Opened {title}");
    sync_sidebar_with_active_tab(state);
}

fn apply_pending_sidebar_intents(state: &mut WindowState) -> bool {
    let mut changed = false;
    while let Some(intent) = state.sidebar.take_intent() {
        match intent {
            SidebarIntent::OpenFile { path, new_tab } => {
                open_path_from_sidebar(state, path, new_tab);
                changed = true;
            }
            SidebarIntent::ToggleFolder(path) => {
                if state.sidebar.toggle_folder(&path) {
                    changed = true;
                }
            }
            SidebarIntent::JumpToBlock(block_id) => {
                if let Some(tab) = state.tabs.active_tab_mut() {
                    tab.cursor.primary.block_id = block_id;
                    tab.cursor.primary.offset = 0;
                    state.sidebar.set_current_outline_block(Some(block_id));
                    state.app_state.status_text = format!("Jumped to block {}", block_id.0);
                    changed = true;
                }
            }
        }
    }
    changed
}

fn canvas_origin(state: &WindowState) -> UiPoint {
    let tab_h = if state.app_state.show_tabs { 36.0 } else { 0.0 };
    let toolbar_h = if state.app_state.show_toolbar {
        44.0
    } else {
        0.0
    };
    let sidebar_w = if state.app_state.show_sidebar {
        state.app_state.sidebar_width.clamp(200.0, 400.0)
    } else {
        0.0
    };
    UiPoint {
        x: sidebar_w,
        y: tab_h + toolbar_h,
    }
}

fn contains_rect(rect: UiRect, point: UiPoint) -> bool {
    point.x >= rect.x
        && point.x <= rect.x + rect.width
        && point.y >= rect.y
        && point.y <= rect.y + rect.height
}

fn active_image_mut(
    state: &mut WindowState,
    block_id: BlockId,
) -> Option<&mut crate::document::model::ImageBlock> {
    state
        .tabs
        .active_tab_mut()
        .and_then(|tab| tab.document.find_image_block_mut(block_id))
}

fn active_image_ref(
    state: &WindowState,
    block_id: BlockId,
) -> Option<&crate::document::model::ImageBlock> {
    state.tabs.active_tab().and_then(|tab| {
        tab.document.content.iter().find_map(|block| match block {
            Block::Image(image) if image.id == block_id => Some(image),
            _ => None,
        })
    })
}

fn active_table_ref(state: &WindowState, table_id: BlockId) -> Option<&crate::document::model::Table> {
    state.tabs.active_tab().and_then(|tab| {
        tab.document.content.iter().find_map(|block| match block {
            Block::Table(table) if table.id == table_id => Some(table),
            _ => None,
        })
    })
}

fn open_table_picker(state: &mut WindowState) {
    state.table_picker_visible = true;
    state.table_picker_rows = state.table_picker_rows.clamp(1, 10);
    state.table_picker_cols = state.table_picker_cols.clamp(1, 10);
    state.table_picker_custom_rows = state.table_picker_rows.to_string();
    state.table_picker_custom_cols = state.table_picker_cols.to_string();
    state.table_picker_custom_focus_rows = true;
}

fn parse_table_picker_custom(value: &str, fallback: usize) -> usize {
    value
        .trim()
        .parse::<usize>()
        .ok()
        .unwrap_or(fallback)
        .clamp(1, 64)
}

fn table_insert_index_for_cursor(tab: &crate::ui::tabs::TabState) -> usize {
    tab.document
        .content
        .iter()
        .position(|block| match block {
            Block::Paragraph(p) => p.id == tab.cursor.primary.block_id,
            Block::Heading(h) => h.id == tab.cursor.primary.block_id,
            Block::CodeBlock(c) => c.id == tab.cursor.primary.block_id,
            Block::Image(i) => i.id == tab.cursor.primary.block_id,
            Block::Table(t) => t.id == tab.cursor.primary.block_id,
            Block::BlockQuote(q) => q.id == tab.cursor.primary.block_id,
            _ => false,
        })
        .map(|idx| idx + 1)
        .unwrap_or(tab.document.content.len())
}

fn insert_table_at_cursor(state: &mut WindowState, rows: usize, cols: usize) -> Option<BlockId> {
    let inserted = {
        let tab = state.tabs.active_tab_mut()?;
        let insert_idx = table_insert_index_for_cursor(tab);
        let id = insert_table(&mut tab.document, insert_idx, rows, cols);
        tab.cursor.primary.block_id = id;
        tab.cursor.primary.offset = 0;
        tab.dirty = true;
        id
    };

    state.selected_table = Some(inserted);
    state.table_selection_mode = Some(TableSelectionMode::Cell(CellPos { row: 0, col: 0 }));
    state.table_selection_range = Some(TableSelection {
        start: CellPos { row: 0, col: 0 },
        end: CellPos { row: 0, col: 0 },
    });
    sync_sidebar_with_active_tab(state);
    Some(inserted)
}

fn insert_table_from_picker(state: &mut WindowState) -> Option<BlockId> {
    let rows = parse_table_picker_custom(state.table_picker_custom_rows.as_str(), state.table_picker_rows);
    let cols = parse_table_picker_custom(state.table_picker_custom_cols.as_str(), state.table_picker_cols);
    state.table_picker_rows = rows.clamp(1, 10);
    state.table_picker_cols = cols.clamp(1, 10);
    state.table_picker_visible = false;
    insert_table_at_cursor(state, rows, cols)
}

fn table_picker_layout(state: &WindowState) -> TablePickerLayout {
    let origin = canvas_origin(state);
    let panel = UiRect {
        x: origin.x + 10.0,
        y: origin.y + 10.0,
        width: 292.0,
        height: 236.0,
    };
    let grid = UiRect {
        x: panel.x + 12.0,
        y: panel.y + 34.0,
        width: 160.0,
        height: 160.0,
    };
    TablePickerLayout {
        panel,
        grid,
        rows_input: UiRect {
            x: panel.x + 12.0,
            y: panel.y + 192.0,
            width: 78.0,
            height: 18.0,
        },
        cols_input: UiRect {
            x: panel.x + 96.0,
            y: panel.y + 192.0,
            width: 78.0,
            height: 18.0,
        },
        insert_button: UiRect {
            x: panel.x + panel.width - 90.0,
            y: panel.y + panel.height - 28.0,
            width: 76.0,
            height: 20.0,
        },
    }
}

fn update_table_picker_hover(state: &mut WindowState, point: UiPoint) -> bool {
    if !state.table_picker_visible {
        return false;
    }
    let layout = table_picker_layout(state);
    if !contains_rect(layout.grid, point) {
        return false;
    }

    let rel_x = (point.x - layout.grid.x).max(0.0);
    let rel_y = (point.y - layout.grid.y).max(0.0);
    let cols = ((rel_x / 16.0).floor() as usize + 1).clamp(1, 10);
    let rows = ((rel_y / 16.0).floor() as usize + 1).clamp(1, 10);
    if state.table_picker_rows == rows && state.table_picker_cols == cols {
        return false;
    }
    state.table_picker_rows = rows;
    state.table_picker_cols = cols;
    state.table_picker_custom_rows = rows.to_string();
    state.table_picker_custom_cols = cols.to_string();
    true
}

fn handle_table_picker_click(state: &mut WindowState, point: UiPoint) -> bool {
    if !state.table_picker_visible {
        return false;
    }
    let layout = table_picker_layout(state);
    if !contains_rect(layout.panel, point) {
        state.table_picker_visible = false;
        state.app_state.status_text = "Insert table cancelled".to_string();
        return true;
    }

    if contains_rect(layout.grid, point) {
        let rel_x = (point.x - layout.grid.x).max(0.0);
        let rel_y = (point.y - layout.grid.y).max(0.0);
        state.table_picker_cols = ((rel_x / 16.0).floor() as usize + 1).clamp(1, 10);
        state.table_picker_rows = ((rel_y / 16.0).floor() as usize + 1).clamp(1, 10);
        state.table_picker_custom_rows = state.table_picker_rows.to_string();
        state.table_picker_custom_cols = state.table_picker_cols.to_string();
        if let Some(id) = insert_table_from_picker(state) {
            state.app_state.status_text = format!(
                "Inserted table {} ({}x{})",
                id.0, state.table_picker_rows, state.table_picker_cols
            );
        } else {
            state.app_state.status_text = "Insert table failed".to_string();
        }
        return true;
    }

    if contains_rect(layout.rows_input, point) {
        state.table_picker_custom_focus_rows = true;
        return true;
    }
    if contains_rect(layout.cols_input, point) {
        state.table_picker_custom_focus_rows = false;
        return true;
    }
    if contains_rect(layout.insert_button, point) {
        if let Some(id) = insert_table_from_picker(state) {
            state.app_state.status_text = format!(
                "Inserted table {} ({}x{})",
                id.0, state.table_picker_rows, state.table_picker_cols
            );
        } else {
            state.app_state.status_text = "Insert table failed".to_string();
        }
        return true;
    }
    true
}

fn move_table_selection(state: &mut WindowState, row_delta: isize, col_delta: isize, expand: bool) -> bool {
    let Some(table_id) = state.selected_table else {
        return false;
    };
    let Some(table) = active_table_ref(state, table_id) else {
        return false;
    };
    let rows = table.rows.len().max(1);
    let cols = table.column_widths.len().max(1);

    let current = selected_table_cell(state).unwrap_or(CellPos { row: 0, col: 0 });
    let row = (current.row as isize + row_delta).clamp(0, rows.saturating_sub(1) as isize) as usize;
    let col = (current.col as isize + col_delta).clamp(0, cols.saturating_sub(1) as isize) as usize;
    let end = CellPos { row, col };

    if expand {
        let start = state
            .table_selection_range
            .as_ref()
            .map(|selection| selection.start)
            .unwrap_or(current);
        state.table_selection_mode = Some(TableSelectionMode::Cell(end));
        state.table_selection_range = Some(TableSelection { start, end });
    } else {
        state.table_selection_mode = Some(TableSelectionMode::Cell(end));
        state.table_selection_range = Some(TableSelection { start: end, end });
    }
    true
}

fn table_selected_row_col(state: &WindowState) -> Option<(usize, usize)> {
    match state.table_selection_mode {
        Some(TableSelectionMode::Cell(cell)) => Some((cell.row, cell.col)),
        Some(TableSelectionMode::Row(row)) => Some((row, 0)),
        Some(TableSelectionMode::Column(col)) => Some((0, col)),
        Some(TableSelectionMode::Table) | None => None,
    }
}

fn apply_table_shortcut(state: &mut WindowState, vk: u32, ctrl_down: bool, shift_down: bool) -> bool {
    let Some(table_id) = state.selected_table else {
        return false;
    };

    if vk == 0x09 {
        if navigate_table_cell(state, shift_down) {
            if let Some(cell) = selected_table_cell(state) {
                state.app_state.status_text = format!("Table cell {},{}", cell.row + 1, cell.col + 1);
            }
            return true;
        }
    }

    if !ctrl_down && (vk == 0x25 || vk == 0x26 || vk == 0x27 || vk == 0x28) {
        let (dr, dc) = match vk {
            0x25 => (0, -1),
            0x26 => (-1, 0),
            0x27 => (0, 1),
            0x28 => (1, 0),
            _ => (0, 0),
        };
        if move_table_selection(state, dr, dc, shift_down) {
            if let Some(sel) = &state.table_selection_range {
                state.app_state.status_text = format!(
                    "Table selection {}:{}, {}:{}",
                    sel.start.row + 1,
                    sel.start.col + 1,
                    sel.end.row + 1,
                    sel.end.col + 1
                );
            }
            return true;
        }
    }

    if vk == VK_DELETE.0 as u32 {
        let mut changed = false;
        if let Some(tab) = state.tabs.active_tab_mut() {
            if let Some(table) = find_table_mut(&mut tab.document, table_id) {
                match state.table_selection_mode {
                    Some(TableSelectionMode::Row(row)) => {
                        if table.rows.len() > 1 {
                            changed = delete_table_row(table, row.min(table.rows.len() - 1));
                        }
                    }
                    Some(TableSelectionMode::Column(col)) => {
                        if table.column_widths.len() > 1 {
                            changed =
                                delete_table_column(table, col.min(table.column_widths.len() - 1));
                        }
                    }
                    Some(TableSelectionMode::Table) => {
                        if let Some(idx) = tab.document.content.iter().position(|block| {
                            matches!(block, Block::Table(t) if t.id == table_id)
                        }) {
                            tab.document.content.remove(idx);
                            changed = true;
                            state.selected_table = None;
                            state.table_selection_mode = None;
                            state.table_selection_range = None;
                        }
                    }
                    _ => {}
                }
                if changed {
                    tab.document.dirty = true;
                    tab.dirty = true;
                }
            }
        }
        if changed {
            state.app_state.status_text = "Table structure updated".to_string();
            return true;
        }
    }

    if !(ctrl_down && shift_down) {
        return false;
    }

    let row_col = table_selected_row_col(state);
    let selected_cell = selected_table_cell(state);
    let selection_range = state.table_selection_range.clone();
    let mut changed = false;
    let mut message = None::<String>;
    if let Some(tab) = state.tabs.active_tab_mut() {
        if let Some(table) = find_table_mut(&mut tab.document, table_id) {
            match vk {
                0x55 => {
                    if let Some((row, _)) = row_col {
                        insert_row_above(table, row);
                        changed = true;
                        message = Some("Inserted row above".to_string());
                    }
                }
                0x4A => {
                    if let Some((row, _)) = row_col {
                        insert_row_below(table, row);
                        changed = true;
                        message = Some("Inserted row below".to_string());
                    }
                }
                0x48 => {
                    if let Some((_, col)) = row_col {
                        insert_column_left(table, col);
                        changed = true;
                        message = Some("Inserted column left".to_string());
                    }
                }
                0x4B => {
                    if let Some((_, col)) = row_col {
                        insert_column_right(table, col);
                        changed = true;
                        message = Some("Inserted column right".to_string());
                    }
                }
                0x4D => {
                    if let Some(selection) = selection_range.clone() {
                        changed = merge_table_cells(table, selection);
                        if changed {
                            message = Some("Merged selected cells".to_string());
                        }
                    }
                }
                0x59 => {
                    if let Some(cell) = selected_cell {
                        changed = split_table_cell(table, cell);
                        if changed {
                            message = Some("Split selected cell".to_string());
                        }
                    }
                }
                0x30 => {
                    let total = table.column_widths.iter().sum::<f32>().max(300.0);
                    distribute_columns_evenly(table, total);
                    changed = true;
                    message = Some("Distributed columns evenly".to_string());
                }
                0x39 => {
                    let total = table.column_widths.iter().sum::<f32>().max(420.0);
                    fit_columns_to_content(table, total);
                    changed = true;
                    message = Some("Auto-fit columns to content".to_string());
                }
                0x31 => {
                    apply_table_style(table, TableStylePreset::Plain);
                    changed = true;
                    message = Some("Applied table style: Plain".to_string());
                }
                0x32 => {
                    apply_table_style(table, TableStylePreset::Grid);
                    changed = true;
                    message = Some("Applied table style: Grid".to_string());
                }
                0x33 => {
                    apply_table_style(table, TableStylePreset::HeaderAccent);
                    changed = true;
                    message = Some("Applied table style: Header row".to_string());
                }
                0x34 => {
                    apply_table_style(table, TableStylePreset::AlternatingRows);
                    changed = true;
                    message = Some("Applied table style: Alternating rows".to_string());
                }
                0x35 => {
                    apply_table_style(table, TableStylePreset::Professional);
                    changed = true;
                    message = Some("Applied table style: Professional".to_string());
                }
                _ => {}
            }
            if changed {
                tab.document.dirty = true;
                tab.dirty = true;
            }
        }
    }

    if changed {
        state.app_state.status_text = message.unwrap_or_else(|| "Table updated".to_string());
    }
    changed
}

fn insert_image_from_path(
    state: &mut WindowState,
    path: &Path,
) -> std::result::Result<BlockId, String> {
    let asset = load_supported_image(path)?;
    let alt_text = path
        .file_stem()
        .and_then(|v| v.to_str())
        .unwrap_or("image")
        .to_string();
    let source_path = Some(path.to_path_buf());
    insert_loaded_image(state, asset, source_path, alt_text)
}

fn insert_loaded_image(
    state: &mut WindowState,
    asset: crate::editor::image_ops::LoadedImageAsset,
    source_path: Option<PathBuf>,
    alt_text: String,
) -> std::result::Result<BlockId, String> {
    let inserted = {
        let Some(tab) = state.tabs.active_tab_mut() else {
            return Err("no active tab".to_string());
        };
        let after = Some(tab.cursor.primary.block_id);
        let block_id = tab.document.insert_embedded_image_after(
            after,
            asset.bytes,
            asset.mime,
            asset.width,
            asset.height,
            source_path,
            alt_text,
        );
        tab.cursor.primary.block_id = block_id;
        tab.cursor.primary.offset = 0;
        tab.dirty = true;
        block_id
    };

    state.selected_image = Some(inserted);
    state.image_properties_visible = false;
    sync_sidebar_with_active_tab(state);
    Ok(inserted)
}

fn insert_images_from_paths(state: &mut WindowState, paths: &[PathBuf]) -> (usize, usize) {
    let mut inserted = 0usize;
    let mut failed = 0usize;
    for path in paths {
        if insert_image_from_path(state, path).is_ok() {
            inserted += 1;
        } else {
            failed += 1;
        }
    }
    (inserted, failed)
}

fn insert_image_from_clipboard(state: &mut WindowState) -> std::result::Result<BlockId, String> {
    let Some(payload) = read_clipboard_image().map_err(|e| e.to_string())? else {
        return Err("clipboard does not contain image data".to_string());
    };
    insert_loaded_image(
        state,
        crate::editor::image_ops::LoadedImageAsset {
            bytes: payload.bytes,
            mime: payload.mime,
            width: payload.width,
            height: payload.height,
        },
        None,
        "Clipboard Image".to_string(),
    )
}

fn collect_canvas_image_overlays(
    tab: &crate::ui::tabs::TabState,
    _selected_image: Option<BlockId>,
    image_cache: &mut ImageDecodeCache,
) -> Vec<CanvasImageOverlay> {
    let page_rect = tab
        .canvas
        .page_rects(&tab.document)
        .first()
        .copied()
        .unwrap_or(UiRect {
            x: 0.0,
            y: 0.0,
            width: tab.canvas.viewport.width.max(1.0),
            height: tab.canvas.viewport.height.max(1.0),
        });

    let content_left = page_rect.x + 46.0;
    let content_right = page_rect.x + page_rect.width - 46.0;
    let max_width = (content_right - content_left).max(72.0);
    let mut cursor_y = page_rect.y + 86.0;
    let bottom_limit = page_rect.y + page_rect.height - 50.0;

    let mut overlays = Vec::new();
    let mut visible_hashes = Vec::new();

    for block in &tab.document.content {
        let Block::Image(image) = block else {
            continue;
        };
        let zoom = tab.canvas.zoom.max(0.25);
        let width = (image.width * zoom * 0.72).clamp(56.0, max_width);
        let mut height = (image.height * zoom * 0.72).clamp(42.0, page_rect.height * 0.5);

        if image.width > 0.0 && image.height > 0.0 {
            let ratio = (image.height / image.width).max(0.08);
            height = (width * ratio).clamp(42.0, page_rect.height * 0.5);
        }

        if cursor_y + height > bottom_limit {
            break;
        }

        let x = match image.alignment {
            ImageAlignment::Left | ImageAlignment::Inline | ImageAlignment::Float => content_left,
            ImageAlignment::Center => content_left + (max_width - width) * 0.5,
            ImageAlignment::Right => content_right - width,
        };
        let rect = UiRect {
            x,
            y: cursor_y,
            width,
            height,
        };

        let scale = if image.original_width > 0 {
            width / image.original_width as f32
        } else {
            1.0
        };
        let interpolation = interpolation_hint(scale).to_string();

        if let Some(data) = resolve_image_data(image, &tab.document) {
            let thumbnail = if scale < 0.45 { Some(384) } else { None };
            if let Ok(decoded) = image_cache.get_or_decode(&data, thumbnail) {
                visible_hashes.push(decoded.source_hash);
            }
        }

        overlays.push(CanvasImageOverlay {
            block_id: image.id,
            rect,
            interpolation,
            alt_text: image.alt_text.clone(),
        });
        cursor_y += height + 16.0;

        if overlays.len() >= 12 {
            break;
        }
    }

    image_cache.mark_visible_hashes(visible_hashes.as_slice());
    overlays
}

fn collect_canvas_table_overlays(tab: &crate::ui::tabs::TabState) -> Vec<CanvasTableOverlay> {
    let page_rect = tab
        .canvas
        .page_rects(&tab.document)
        .first()
        .copied()
        .unwrap_or(UiRect {
            x: 0.0,
            y: 0.0,
            width: tab.canvas.viewport.width.max(1.0),
            height: tab.canvas.viewport.height.max(1.0),
        });

    let left = page_rect.x + 46.0;
    let mut top = page_rect.y + 430.0;
    let max_width = (page_rect.width - 92.0).max(140.0);
    let mut overlays = Vec::new();

    for block in &tab.document.content {
        let Block::Table(table) = block else {
            continue;
        };
        let rows = table.rows.len().max(1);
        let cols = table.column_widths.len().max(1);
        let gutter_w = 18.0;
        let header_h = 18.0;
        let cell_h = 24.0;
        let cell_w = ((max_width - gutter_w) / cols as f32).max(28.0);
        let visible = visible_row_range(table, tab.canvas.scroll.y.max(0.0), tab.canvas.viewport.height, cell_h);
        let visible_rows = (visible.1.saturating_sub(visible.0)).max(1);
        let total_h = header_h + visible_rows as f32 * cell_h;
        let total_w = gutter_w + cell_w * cols as f32;
        if top + total_h > page_rect.y + page_rect.height - 24.0 {
            break;
        }

        overlays.push(CanvasTableOverlay {
            table_id: table.id,
            rect: UiRect {
                x: left,
                y: top,
                width: total_w,
                height: total_h,
            },
            rows,
            cols,
            cell_w,
            cell_h,
            header_h,
            gutter_w,
        });
        top += total_h + 18.0;

        if overlays.len() >= 8 {
            break;
        }
    }

    overlays
}

fn begin_table_interaction(state: &mut WindowState, point: UiPoint) -> bool {
    let origin = canvas_origin(state);
    let local = UiPoint {
        x: point.x - origin.x,
        y: point.y - origin.y,
    };

    let overlay = state
        .canvas_table_overlays
        .iter()
        .rev()
        .find(|overlay| contains_rect(overlay.rect, local))
        .cloned();
    let Some(overlay) = overlay else {
        return false;
    };

    let local_x = local.x - overlay.rect.x;
    let local_y = local.y - overlay.rect.y;
    let rel_col = ((local_x - overlay.gutter_w) / overlay.cell_w).floor().max(0.0) as usize;
    let rel_row = ((local_y - overlay.header_h) / overlay.cell_h).floor().max(0.0) as usize;
    let col = rel_col.min(overlay.cols.saturating_sub(1));
    let row = rel_row.min(overlay.rows.saturating_sub(1));

    state.selected_table = Some(overlay.table_id);
    state.selected_image = None;
    state.image_drag = None;

    if local_x <= overlay.gutter_w && local_y <= overlay.header_h {
        state.table_selection_mode = Some(TableSelectionMode::Table);
        state.table_selection_range = Some(TableSelection {
            start: CellPos { row: 0, col: 0 },
            end: CellPos {
                row: overlay.rows.saturating_sub(1),
                col: overlay.cols.saturating_sub(1),
            },
        });
    } else if local_x <= overlay.gutter_w {
        state.table_selection_mode = Some(TableSelectionMode::Row(row));
        state.table_selection_range = Some(TableSelection {
            start: CellPos { row, col: 0 },
            end: CellPos {
                row,
                col: overlay.cols.saturating_sub(1),
            },
        });
    } else if local_y <= overlay.header_h {
        state.table_selection_mode = Some(TableSelectionMode::Column(col));
        state.table_selection_range = Some(TableSelection {
            start: CellPos { row: 0, col },
            end: CellPos {
                row: overlay.rows.saturating_sub(1),
                col,
            },
        });
    } else {
        state.table_selection_mode = Some(TableSelectionMode::Cell(CellPos { row, col }));
        state.table_selection_range = Some(TableSelection {
            start: CellPos { row, col },
            end: CellPos { row, col },
        });
    }

    // Column/row border drag handles.
    let near_col_border = if local_x > overlay.gutter_w {
        let x = local_x - overlay.gutter_w;
        let frac = (x / overlay.cell_w).fract();
        frac < 0.08 || frac > 0.92
    } else {
        false
    };
    if near_col_border {
        if let Some(table) = active_table_ref(state, overlay.table_id) {
            let border_idx = ((local_x - overlay.gutter_w) / overlay.cell_w).round().max(0.0) as usize;
            let col_idx = border_idx.min(overlay.cols.saturating_sub(1));
            let start_value = table.column_widths.get(col_idx).copied().unwrap_or(120.0);
            state.table_resize = Some(TableResizeState {
                table_id: overlay.table_id,
                row: None,
                col: Some(col_idx),
                start_mouse: local,
                start_value,
            });
        }
    } else if local_x <= overlay.gutter_w && local_y > overlay.header_h {
        if let Some(table) = active_table_ref(state, overlay.table_id) {
            let border_idx = ((local_y - overlay.header_h) / overlay.cell_h).round().max(0.0) as usize;
            let row_idx = border_idx.min(overlay.rows.saturating_sub(1));
            let start_value = table.row_heights.get(row_idx).copied().unwrap_or(28.0);
            state.table_resize = Some(TableResizeState {
                table_id: overlay.table_id,
                row: Some(row_idx),
                col: None,
                start_mouse: local,
                start_value,
            });
        }
    } else {
        state.table_resize = None;
    }

    state.app_state.status_text = format!("Table {} selected", overlay.table_id.0);
    true
}

fn update_table_resize(state: &mut WindowState, point: UiPoint) -> bool {
    let Some(resize) = state.table_resize.clone() else {
        return false;
    };
    let origin = canvas_origin(state);
    let local = UiPoint {
        x: point.x - origin.x,
        y: point.y - origin.y,
    };
    let dx = local.x - resize.start_mouse.x;
    let dy = local.y - resize.start_mouse.y;

    let mut changed = false;
    if let Some(tab) = state.tabs.active_tab_mut() {
        if let Some(table) = find_table_mut(&mut tab.document, resize.table_id) {
            if let Some(col) = resize.col {
                changed = resize_table_column(table, col, resize.start_value + dx);
            } else if let Some(row) = resize.row {
                changed = resize_table_row(table, row, resize.start_value + dy);
            }
            if changed {
                tab.document.dirty = true;
                tab.dirty = true;
            }
        }
    }
    changed
}

fn selected_table_cell(state: &WindowState) -> Option<CellPos> {
    match state.table_selection_mode {
        Some(TableSelectionMode::Cell(cell)) => Some(cell),
        Some(TableSelectionMode::Row(row)) => Some(CellPos { row, col: 0 }),
        Some(TableSelectionMode::Column(col)) => Some(CellPos { row: 0, col }),
        Some(TableSelectionMode::Table) | None => None,
    }
}

fn navigate_table_cell(state: &mut WindowState, backwards: bool) -> bool {
    let Some(table_id) = state.selected_table else {
        return false;
    };
    let Some(current) = selected_table_cell(state) else {
        return false;
    };
    let Some(table) = active_table_ref(state, table_id) else {
        return false;
    };
    let rows = table.rows.len().max(1);
    let cols = table.column_widths.len().max(1);
    let mut row = current.row.min(rows.saturating_sub(1));
    let mut col = current.col.min(cols.saturating_sub(1));

    if backwards {
        if col > 0 {
            col -= 1;
        } else if row > 0 {
            row -= 1;
            col = cols.saturating_sub(1);
        }
    } else if col + 1 < cols {
        col += 1;
    } else if row + 1 < rows {
        row += 1;
        col = 0;
    } else {
        if let Some(tab) = state.tabs.active_tab_mut() {
            if let Some(table_mut) = find_table_mut(&mut tab.document, table_id) {
                insert_row_below(table_mut, rows.saturating_sub(1));
                tab.document.dirty = true;
                tab.dirty = true;
            }
        }
        row = rows;
        col = 0;
    }

    state.table_selection_mode = Some(TableSelectionMode::Cell(CellPos { row, col }));
    state.table_selection_range = Some(TableSelection {
        start: CellPos { row, col },
        end: CellPos { row, col },
    });
    true
}

fn image_drag_kind_for_point(rect: UiRect, point: UiPoint) -> ImageDragKind {
    let edge = 8.0;
    let near_left = (point.x - rect.x).abs() <= edge;
    let near_right = (point.x - (rect.x + rect.width)).abs() <= edge;
    let near_top = (point.y - rect.y).abs() <= edge;
    let near_bottom = (point.y - (rect.y + rect.height)).abs() <= edge;

    if (near_left || near_right) && (near_top || near_bottom) {
        return ImageDragKind::CornerResize;
    }
    if near_left || near_right {
        return ImageDragKind::EdgeResizeHorizontal;
    }
    if near_top || near_bottom {
        return ImageDragKind::EdgeResizeVertical;
    }
    ImageDragKind::Move
}

fn begin_image_interaction(state: &mut WindowState, point: UiPoint) -> bool {
    let origin = canvas_origin(state);
    let local = UiPoint {
        x: point.x - origin.x,
        y: point.y - origin.y,
    };

    let hit = state
        .canvas_image_overlays
        .iter()
        .rev()
        .find(|overlay| contains_rect(overlay.rect, local))
        .cloned();
    let Some(hit_overlay) = hit else {
        return false;
    };

    state.selected_image = Some(hit_overlay.block_id);
    state.image_properties_visible = false;

    if let Some(image) = active_image_ref(state, hit_overlay.block_id) {
        state.image_drag = Some(ImageDragState {
            block_id: hit_overlay.block_id,
            start_mouse: local,
            start_width: image.width,
            start_height: image.height,
            start_alignment: image.alignment.clone(),
            kind: image_drag_kind_for_point(hit_overlay.rect, local),
        });
        state.app_state.status_text = format!("Selected image {}", hit_overlay.block_id.0);
        return true;
    }

    false
}

fn update_image_drag(state: &mut WindowState, point: UiPoint, shift_down: bool) -> bool {
    let Some(drag) = state.image_drag.clone() else {
        return false;
    };
    let origin = canvas_origin(state);
    let local = UiPoint {
        x: point.x - origin.x,
        y: point.y - origin.y,
    };
    let delta_x = local.x - drag.start_mouse.x;
    let delta_y = local.y - drag.start_mouse.y;

    let mut changed = false;
    let zoom = state
        .tabs
        .active_tab()
        .map(|tab| tab.canvas.zoom.max(0.25))
        .unwrap_or(1.0);

    if let Some(image) = active_image_mut(state, drag.block_id) {
        match drag.kind {
            ImageDragKind::Move => {
                image.alignment = if delta_x < -40.0 {
                    ImageAlignment::Left
                } else if delta_x > 40.0 {
                    ImageAlignment::Right
                } else {
                    drag.start_alignment.clone()
                };
            }
            ImageDragKind::CornerResize => {
                let mut width = (drag.start_width + delta_x / zoom).max(24.0);
                let mut height = (drag.start_height + delta_y / zoom).max(24.0);
                if shift_down {
                    let ratio = (drag.start_width / drag.start_height.max(1.0)).max(0.05);
                    if delta_x.abs() >= delta_y.abs() {
                        height = (width / ratio).max(24.0);
                    } else {
                        width = (height * ratio).max(24.0);
                    }
                }
                image.width = width;
                image.height = height;
            }
            ImageDragKind::EdgeResizeHorizontal => {
                image.width = (drag.start_width + delta_x / zoom).max(24.0);
            }
            ImageDragKind::EdgeResizeVertical => {
                image.height = (drag.start_height + delta_y / zoom).max(24.0);
            }
        }
        changed = true;
    }

    if changed {
        if let Some(tab) = state.tabs.active_tab_mut() {
            tab.document.dirty = true;
            tab.dirty = true;
        }
        state.app_state.status_text = if let Some(image) = active_image_ref(state, drag.block_id) {
            format!(
                "Image {} {:.0}x{:.0} ({:?})",
                drag.block_id.0, image.width, image.height, image.alignment
            )
        } else {
            "Image updated".to_string()
        };
    }

    changed
}

fn delete_selected_image(state: &mut WindowState) -> bool {
    let Some(selected) = state.selected_image else {
        return false;
    };
    let mut removed = false;
    if let Some(tab) = state.tabs.active_tab_mut() {
        removed = tab.document.remove_image_block(selected);
        if removed {
            tab.dirty = true;
            tab.cursor.primary.offset = 0;
        }
    }
    if removed {
        state.selected_image = None;
        state.image_drag = None;
        state.image_properties_visible = false;
        sync_sidebar_with_active_tab(state);
    }
    removed
}

fn align_selected_image(state: &mut WindowState, alignment: ImageAlignment) -> bool {
    let Some(selected) = state.selected_image else {
        return false;
    };
    if let Some(tab) = state.tabs.active_tab_mut() {
        if let Some(image) = tab.document.find_image_block_mut(selected) {
            image.alignment = alignment;
            tab.document.dirty = true;
            tab.dirty = true;
            return true;
        }
    }
    false
}

fn toggle_selected_image_border(state: &mut WindowState) -> bool {
    let Some(selected) = state.selected_image else {
        return false;
    };
    if let Some(tab) = state.tabs.active_tab_mut() {
        if let Some(image) = tab.document.find_image_block_mut(selected) {
            image.border = if image.border.is_some() {
                None
            } else {
                Some(ImageBorder {
                    style: ImageBorderStyle::Solid,
                    width: 1.0,
                    color: crate::ui::Color::rgb(0.35, 0.54, 0.92),
                })
            };
            tab.document.dirty = true;
            tab.dirty = true;
            return true;
        }
    }
    false
}

fn collect_visible_block_ids_for_search(tab: &mut crate::ui::tabs::TabState) -> Vec<BlockId> {
    let mut visible_ids = Vec::new();
    let visible_pages = tab.canvas.cull_and_cache_visible_pages(&tab.document);
    for page_index in visible_pages {
        if let Some(page) = tab.document.pages.get(page_index) {
            visible_ids.extend(page.block_ids.iter().copied());
        }
    }

    if visible_ids.is_empty() {
        visible_ids.push(tab.cursor.primary.block_id);
    }
    visible_ids.sort_by_key(|id| id.0);
    visible_ids.dedup();
    visible_ids
}

fn sync_sidebar_search_results(state: &mut WindowState) {
    let items = state
        .find_replace
        .results
        .iter()
        .take(500)
        .map(|m| SearchResultItem {
            block_id: m.block_id,
            line_or_page: m.line_or_page,
            snippet: m.snippet.clone(),
            start: m.start,
            end: m.end,
        })
        .collect::<Vec<_>>();
    state
        .sidebar
        .set_search_results(state.find_replace.query.clone(), items);
}

fn refresh_find_results(state: &mut WindowState) -> bool {
    let mut changed = false;
    if let Some(tab) = state.tabs.active_tab_mut() {
        let visible_ids = collect_visible_block_ids_for_search(tab);
        let previous_count = state.find_replace.results.len();
        let previous_index = state.find_replace.current_index;
        let _ = state
            .find_replace
            .refresh_results_with_visible(&tab.document, &visible_ids);
        changed = previous_count != state.find_replace.results.len()
            || previous_index != state.find_replace.current_index;
    }

    sync_sidebar_search_results(state);
    if state.find_replace.find_visible && !state.find_replace.query.is_empty() {
        state.sidebar.set_active_panel(SidebarPanel::SearchResults);
    }
    changed
}

fn process_find_background_search(state: &mut WindowState, budget_blocks: usize) -> bool {
    let changed = state.find_replace.process_background_search(budget_blocks);
    if changed || !state.find_replace.has_pending_background_search() {
        sync_sidebar_search_results(state);
    }
    changed
}

fn jump_to_search_match(
    state: &mut WindowState,
    search_match: &crate::editor::search::SearchMatch,
) {
    if let Some(tab) = state.tabs.active_tab_mut() {
        tab.cursor.primary.block_id = search_match.block_id;
        tab.cursor.primary.offset = search_match.start;
        state
            .sidebar
            .set_current_outline_block(Some(search_match.block_id));
    }
}

fn navigate_find_result(state: &mut WindowState, backwards: bool) -> bool {
    let found = if backwards {
        state.find_replace.previous().cloned()
    } else {
        state.find_replace.next().cloned()
    };
    if let Some(m) = found {
        jump_to_search_match(state, &m);
        state.app_state.status_text = format!(
            "Match {}/{}",
            state.find_replace.current_index + 1,
            state.find_replace.results.len()
        );
        return true;
    }
    false
}

fn replace_current_match(state: &mut WindowState) -> usize {
    let mut count = 0;
    if let Some(tab) = state.tabs.active_tab_mut() {
        count = replace_current(&mut tab.document, &mut state.find_replace);
    }
    if count > 0 {
        sync_sidebar_search_results(state);
    }
    count
}

fn replace_all_matches(state: &mut WindowState) -> usize {
    let mut count = 0;
    if let Some(tab) = state.tabs.active_tab_mut() {
        count = replace_all(&mut tab.document, &mut state.find_replace);
    }
    if count > 0 {
        sync_sidebar_search_results(state);
    }
    count
}

fn remove_last_char(text: &mut String) {
    let _ = text.pop();
}

fn collect_navigable_block_ids(doc: &DocumentModel, out: &mut Vec<BlockId>) {
    fn walk(block: &Block, out: &mut Vec<BlockId>) {
        match block {
            Block::Paragraph(p) => out.push(p.id),
            Block::Heading(h) => out.push(h.id),
            Block::CodeBlock(c) => out.push(c.id),
            Block::Table(table) => {
                for row in &table.rows {
                    for cell in &row.cells {
                        for nested in &cell.blocks {
                            walk(nested, out);
                        }
                    }
                }
            }
            Block::List(list) => {
                for item in &list.items {
                    for nested in &item.content {
                        walk(nested, out);
                    }
                    for child in &item.children {
                        for nested in &child.content {
                            walk(nested, out);
                        }
                    }
                }
            }
            Block::BlockQuote(q) => {
                for nested in &q.blocks {
                    walk(nested, out);
                }
            }
            Block::Image(_) | Block::PageBreak | Block::HorizontalRule => {}
        }
    }

    for block in &doc.content {
        walk(block, out);
    }
}

fn jump_to_line_or_page(state: &mut WindowState, one_based: usize) -> bool {
    if one_based == 0 {
        return false;
    }
    let Some(tab) = state.tabs.active_tab_mut() else {
        return false;
    };

    if let Some(page) = tab.document.pages.get(one_based - 1) {
        if let Some(id) = page.block_ids.first().copied() {
            tab.cursor.primary.block_id = id;
            tab.cursor.primary.offset = 0;
            state.sidebar.set_current_outline_block(Some(id));
            return true;
        }
    }

    let mut ids = Vec::new();
    collect_navigable_block_ids(&tab.document, &mut ids);
    if let Some(id) = ids.get(one_based - 1).copied() {
        tab.cursor.primary.block_id = id;
        tab.cursor.primary.offset = 0;
        state.sidebar.set_current_outline_block(Some(id));
        return true;
    }

    false
}

fn block_snippet(document: &DocumentModel, block_id: BlockId) -> String {
    for block in &document.content {
        match block {
            Block::Paragraph(p) if p.id == block_id => {
                return p.runs.iter().map(|r| r.text.as_str()).collect::<String>();
            }
            Block::Heading(h) if h.id == block_id => {
                return h.runs.iter().map(|r| r.text.as_str()).collect::<String>();
            }
            _ => {}
        }
    }
    String::new()
}

fn canvas_viewport_size(state: &WindowState, width: f32, height: f32) -> (f32, f32) {
    let ui_scale = state
        .app_state
        .settings
        .appearance
        .ui_scale
        .as_factor()
        .clamp(1.0, 2.0);
    let tab_h = if state.app_state.show_tabs {
        36.0 * ui_scale
    } else {
        0.0
    };
    let status_h = if state.app_state.show_statusbar {
        28.0 * ui_scale
    } else {
        0.0
    };
    let sidebar_w = if state.app_state.show_sidebar {
        state.app_state.sidebar_width.clamp(200.0, 400.0)
    } else {
        0.0
    };
    let toolbar_h = if state.app_state.show_toolbar {
        44.0 * ui_scale
    } else {
        0.0
    };
    (
        (width - sidebar_w).max(1.0),
        (height - tab_h - toolbar_h - status_h).max(1.0),
    )
}

fn relayout_shell(state: &mut WindowState, width: f32, height: f32) {
    let ui_scale = state
        .app_state
        .settings
        .appearance
        .ui_scale
        .as_factor()
        .clamp(1.0, 2.0);
    let tab_h = if state.app_state.show_tabs {
        36.0 * ui_scale
    } else {
        0.0
    };
    let status_h = if state.app_state.show_statusbar {
        28.0 * ui_scale
    } else {
        0.0
    };
    let sidebar_w = if state.app_state.show_sidebar {
        state.app_state.sidebar_width.clamp(200.0, 400.0)
    } else {
        0.0
    };
    let toolbar_h = if state.app_state.show_toolbar {
        44.0 * ui_scale
    } else {
        0.0
    };

    state.tabs.set_visible(state.app_state.show_tabs);
    state.sidebar.set_visible(state.app_state.show_sidebar);
    state.toolbar.set_visible(state.app_state.show_toolbar);
    state.statusbar.set_visible(state.app_state.show_statusbar);
    if state.app_state.show_sidebar {
        state.sidebar.set_width(sidebar_w);
    }

    state.tabs.layout(
        UiRect {
            x: 0.0,
            y: 0.0,
            width,
            height: tab_h,
        },
        state.dpi,
    );
    state.sidebar.layout(
        UiRect {
            x: 0.0,
            y: tab_h,
            width: sidebar_w,
            height: (height - tab_h - status_h).max(0.0),
        },
        state.dpi,
    );
    state.toolbar.layout(
        UiRect {
            x: sidebar_w,
            y: tab_h,
            width: (width - sidebar_w).max(0.0),
            height: toolbar_h,
        },
        state.dpi,
    );
    state.statusbar.layout(
        UiRect {
            x: 0.0,
            y: 0.0,
            width,
            height,
        },
        state.dpi,
    );
    state.command_palette.layout(
        UiRect {
            x: 0.0,
            y: 0.0,
            width,
            height,
        },
        state.dpi,
    );
    state.settings_dialog.layout(
        UiRect {
            x: 0.0,
            y: 0.0,
            width,
            height,
        },
        state.dpi,
    );

    let (canvas_w, canvas_h) = canvas_viewport_size(state, width, height);
    if let Some(tab) = state.tabs.active_tab_mut() {
        tab.canvas.set_viewport(canvas_w, canvas_h);
        tab.canvas.clamp_scroll(&tab.document);
    }
}

fn sync_theme_from_settings(state: &mut WindowState) -> bool {
    let previous_name = state.theme.name.clone();
    let previous_is_dark = state.theme.is_dark;
    let next = state
        .theme_manager
        .apply_preference(&state.app_state.settings.appearance.theme);
    let changed = previous_name != next.name || previous_is_dark != next.is_dark;
    state.theme = next;
    changed
}

fn sidebar_panel_from_preference(pref: SidebarDefaultPanel) -> SidebarPanel {
    match pref {
        SidebarDefaultPanel::Files => SidebarPanel::Files,
        SidebarDefaultPanel::Outline => SidebarPanel::Outline,
        SidebarDefaultPanel::Bookmarks => SidebarPanel::Bookmarks,
    }
}

fn set_settings_visible(state: &mut WindowState, visible: bool) {
    state.app_state.show_settings = visible;
    state.settings_dialog.set_visible(visible);
}

fn sync_runtime_from_settings(state: &mut WindowState, hwnd: HWND) {
    let settings = state.settings_dialog.settings().clone();

    let prev_show_toolbar = state.app_state.show_toolbar;
    let prev_show_sidebar = state.app_state.show_sidebar;
    let prev_show_statusbar = state.app_state.show_statusbar;
    let prev_show_tabs = state.app_state.show_tabs;
    let prev_ui_scale = state.app_state.settings.appearance.ui_scale.as_factor();
    let prev_sidebar_panel = state.sidebar.active_panel;

    state.app_state.settings = settings;
    state.app_state.show_toolbar = state.app_state.settings.appearance.show_toolbar;
    state.app_state.show_sidebar = state.app_state.settings.appearance.show_sidebar;
    state.app_state.show_statusbar = state.app_state.settings.appearance.show_status_bar;
    state.app_state.show_tabs = state.app_state.settings.appearance.show_tab_bar;

    let preferred_panel =
        sidebar_panel_from_preference(state.app_state.settings.appearance.sidebar_default_panel);
    if state.sidebar.active_panel != SidebarPanel::SearchResults {
        state.sidebar.set_active_panel(preferred_panel);
    }

    let autosave_seconds = state
        .app_state
        .settings
        .files
        .auto_save_interval
        .as_seconds()
        .unwrap_or(60 * 60 * 24 * 365 * 100)
        .max(5);
    let desired_interval = Duration::from_secs(autosave_seconds);
    if state.app_state.autosave.interval != desired_interval {
        state.app_state.autosave = crate::document::export::AutoSaveManager::new(autosave_seconds);
    }

    let desired_image_cache_bytes = (state.app_state.settings.performance.max_image_cache_mb as usize)
        .max(32)
        * 1024
        * 1024;
    if state.image_cache.max_bytes != desired_image_cache_bytes {
        state.image_cache.set_memory_budget(desired_image_cache_bytes);
    }

    let next_ui_scale = state.app_state.settings.appearance.ui_scale.as_factor();
    let needs_relayout = prev_show_toolbar != state.app_state.show_toolbar
        || prev_show_sidebar != state.app_state.show_sidebar
        || prev_show_statusbar != state.app_state.show_statusbar
        || prev_show_tabs != state.app_state.show_tabs
        || (prev_ui_scale - next_ui_scale).abs() > f32::EPSILON
        || prev_sidebar_panel != state.sidebar.active_panel;

    let theme_changed = sync_theme_from_settings(state);
    if theme_changed {
        if let Some(renderer) = &mut state.renderer {
            renderer.set_theme(state.theme.clone());
        }
        unsafe { apply_window_effects(hwnd, state.theme.is_dark) };
    }

    if needs_relayout {
        let mut client = RECT::default();
        let _ = unsafe { GetClientRect(hwnd, &mut client) };
        relayout_shell(
            state,
            (client.right - client.left).max(0) as f32,
            (client.bottom - client.top).max(0) as f32,
        );
    }
}

fn collect_document_stats(document: &DocumentModel) -> (usize, usize) {
    fn walk_block(block: &Block, words: &mut usize, chars: &mut usize) {
        match block {
            Block::Paragraph(p) => {
                for run in &p.runs {
                    *chars += run.text.chars().count();
                    *words += run.text.split_whitespace().count();
                }
            }
            Block::Heading(h) => {
                for run in &h.runs {
                    *chars += run.text.chars().count();
                    *words += run.text.split_whitespace().count();
                }
            }
            Block::CodeBlock(c) => {
                *chars += c.code.chars().count();
                *words += c.code.split_whitespace().count();
            }
            Block::List(list) => {
                for item in &list.items {
                    for nested in &item.content {
                        walk_block(nested, words, chars);
                    }
                }
            }
            Block::Table(table) => {
                for row in &table.rows {
                    for cell in &row.cells {
                        for nested in &cell.blocks {
                            walk_block(nested, words, chars);
                        }
                    }
                }
            }
            Block::BlockQuote(q) => {
                for nested in &q.blocks {
                    walk_block(nested, words, chars);
                }
            }
            Block::Image(_) | Block::PageBreak | Block::HorizontalRule => {}
        }
    }

    let mut words = 0;
    let mut chars = 0;
    for block in &document.content {
        walk_block(block, &mut words, &mut chars);
    }
    (words, chars)
}

fn collect_preview_lines(document: &DocumentModel, max_lines: usize) -> Vec<String> {
    fn push_block_lines(block: &Block, out: &mut Vec<String>, max_lines: usize) {
        if out.len() >= max_lines {
            return;
        }
        match block {
            Block::Paragraph(p) => {
                let text = p.runs.iter().map(|r| r.text.as_str()).collect::<String>();
                if !text.trim().is_empty() {
                    out.push(text);
                }
            }
            Block::Heading(h) => {
                let text = h.runs.iter().map(|r| r.text.as_str()).collect::<String>();
                if !text.trim().is_empty() {
                    out.push(text.to_uppercase());
                }
            }
            Block::CodeBlock(c) => {
                let line = if c.code.is_empty() {
                    "code block"
                } else {
                    &c.code
                };
                out.push(line.lines().next().unwrap_or("code block").to_string());
            }
            Block::List(list) => {
                for item in &list.items {
                    if out.len() >= max_lines {
                        break;
                    }
                    for nested in &item.content {
                        push_block_lines(nested, out, max_lines);
                    }
                }
            }
            Block::Table(table) => {
                out.push(format!("Table: {} rows", table.rows.len()));
            }
            Block::BlockQuote(q) => {
                for nested in &q.blocks {
                    if out.len() >= max_lines {
                        break;
                    }
                    push_block_lines(nested, out, max_lines);
                }
            }
            Block::Image(_) => out.push("[Image]".to_string()),
            Block::PageBreak => out.push(String::new()),
            Block::HorizontalRule => out.push("----".to_string()),
        }
    }

    let mut out = Vec::new();
    for block in &document.content {
        push_block_lines(block, &mut out, max_lines);
        if out.len() >= max_lines {
            break;
        }
    }

    if out.is_empty() {
        out.push("Start typing here...".to_string());
    }

    out
}

fn append_block_text(block: &Block, out: &mut String) {
    match block {
        Block::Paragraph(p) => {
            for run in &p.runs {
                out.push_str(run.text.as_str());
            }
            out.push('\n');
        }
        Block::Heading(h) => {
            for run in &h.runs {
                out.push_str(run.text.as_str());
            }
            out.push('\n');
        }
        Block::CodeBlock(c) => {
            out.push_str(c.code.as_str());
            out.push('\n');
        }
        Block::List(list) => {
            for item in &list.items {
                for nested in &item.content {
                    append_block_text(nested, out);
                }
            }
        }
        Block::Table(table) => {
            for row in &table.rows {
                for cell in &row.cells {
                    for nested in &cell.blocks {
                        append_block_text(nested, out);
                    }
                    out.push(' ');
                }
                out.push('\n');
            }
        }
        Block::BlockQuote(quote) => {
            for nested in &quote.blocks {
                append_block_text(nested, out);
            }
        }
        Block::Image(_) | Block::PageBreak | Block::HorizontalRule => {}
    }
}

fn collect_document_plain_text(document: &DocumentModel) -> String {
    let mut out = String::new();
    for block in &document.content {
        append_block_text(block, &mut out);
    }
    out
}

fn block_id_for_search(block: &Block) -> BlockId {
    match block {
        Block::Paragraph(p) => p.id,
        Block::Heading(h) => h.id,
        Block::CodeBlock(c) => c.id,
        Block::List(list) => list
            .items
            .first()
            .and_then(|item| item.content.first())
            .map(block_id_for_search)
            .unwrap_or(BlockId(0)),
        Block::Image(image) => image.id,
        Block::Table(table) => table.id,
        Block::BlockQuote(quote) => quote.id,
        Block::PageBreak | Block::HorizontalRule => BlockId(0),
    }
}

fn find_in_all_open_tabs(state: &mut WindowState, query: &str) -> (usize, usize) {
    let needle = query.trim();
    if needle.is_empty() {
        state.sidebar.set_search_results("", Vec::new());
        return (0, 0);
    }

    let mut tabs_with_matches = 0usize;
    let mut total_matches = 0usize;
    let mut sidebar_results = Vec::new();

    for tab in &state.tabs.tabs {
        if tab.kind == TabKind::Welcome {
            continue;
        }

        let text = collect_document_plain_text(&tab.document);
        if text.is_empty() {
            continue;
        }

        let mut offset = 0usize;
        let mut tab_matches = 0usize;
        while let Some(rel) = text[offset..].find(needle) {
            let absolute = offset + rel;
            tab_matches += 1;
            total_matches += 1;

            if sidebar_results.len() < 120 {
                let chars_before = text[..absolute].chars().count();
                let snippet_start = chars_before.saturating_sub(22);
                let snippet = text
                    .chars()
                    .skip(snippet_start)
                    .take(90)
                    .collect::<String>()
                    .replace('\n', " ");
                let block_id = tab
                    .document
                    .content
                    .first()
                    .map(block_id_for_search)
                    .unwrap_or(BlockId(0));
                sidebar_results.push(SearchResultItem {
                    block_id,
                    line_or_page: 1,
                    snippet: format!("{}: {}", tab.title, snippet.trim()),
                    start: absolute,
                    end: absolute + needle.len(),
                });
            }

            offset = absolute + needle.len().max(1);
            if offset >= text.len() {
                break;
            }
        }

        if tab_matches > 0 {
            tabs_with_matches += 1;
        }
    }

    state
        .sidebar
        .set_search_results(needle.to_string(), sidebar_results);
    state.sidebar.set_active_panel(SidebarPanel::SearchResults);

    (tabs_with_matches, total_matches)
}

fn tab_icon_label(tab: &crate::ui::tabs::TabState) -> &'static str {
    if tab.kind == TabKind::Welcome {
        return "[HOME]";
    }
    let extension = tab
        .file_path
        .as_ref()
        .or(tab.document.metadata.file_path.as_ref())
        .and_then(|path| path.extension().and_then(|v| v.to_str()))
        .map(|value| value.to_ascii_lowercase());
    match extension.as_deref() {
        Some("docx") => "[DOCX]",
        Some("pdf") => "[PDF]",
        Some("md") | Some("markdown") => "[MD]",
        Some("txt") => "[TXT]",
        _ => "[DOC]",
    }
}

fn tab_shell_title(tab: &crate::ui::tabs::TabState) -> String {
    let dirty = if is_tab_dirty(tab) { " *" } else { "" };
    format!("{} {}{}", tab_icon_label(tab), tab.title, dirty)
}

fn welcome_preview_lines(state: &WindowState) -> Vec<String> {
    let mut lines = vec![
        "DOCO".to_string(),
        "Document editor".to_string(),
        String::new(),
        "Quick actions:".to_string(),
        "  - Ctrl+O: Open file".to_string(),
        "  - Ctrl+T: New tab".to_string(),
        "  - Ctrl+Shift+S: Save As".to_string(),
        "  - Ctrl+Shift+F: Find in all open tabs".to_string(),
        String::new(),
        "Recent files:".to_string(),
    ];
    if state.jump_list.recent_files.is_empty() {
        lines.push("  (none)".to_string());
    } else {
        for recent in state.jump_list.recent_files.iter().take(8) {
            lines.push(format!("  - {}", recent.display()));
        }
    }
    lines
}

fn build_shell_render_state(state: &mut WindowState) -> ShellRenderState {
    let mut word_count = 0usize;
    let mut character_count = 0usize;
    let mut page_index = 1usize;
    let mut page_count = 1usize;
    let mut view_mode = "Page".to_string();
    let mut zoom_percent = 100u16;
    let mut file_format = "DOCX".to_string();
    let mut line = 1usize;
    let mut column = 1usize;

    let mut canvas_page_rects = Vec::new();
    let mut canvas_preview_lines = Vec::new();
    let mut canvas_show_margin_guides = false;
    let mut canvas_cursor_visible = true;
    let mut canvas_scrollbar_visible = false;
    let mut canvas_scrollbar_alpha = 0.0f32;
    let mut canvas_viewport_width = 1.0f32;
    let mut canvas_viewport_height = 1.0f32;
    let mut canvas_content_width = 1.0f32;
    let mut canvas_content_height = 1.0f32;
    let mut canvas_scroll_x = 0.0f32;
    let mut canvas_scroll_y = 0.0f32;
    let mut canvas_images = Vec::new();
    let mut canvas_tables = Vec::new();
    let mut current_block = None;
    let mut active_is_welcome = false;
    let selected_image_id = state.selected_image;

    {
        let (tabs, image_cache) = (&mut state.tabs, &mut state.image_cache);
        if let Some(tab) = tabs.active_tab_mut() {
            active_is_welcome = tab.kind == TabKind::Welcome;
            (word_count, character_count) = collect_document_stats(&tab.document);
            let visible_indices = tab.canvas.cull_and_cache_visible_pages(&tab.document);
            let all_page_rects = tab.canvas.page_rects(&tab.document);
            let first_visible_index = visible_indices.first().copied();

            for page_index_visible in visible_indices {
                if let Some(rect) = all_page_rects.get(page_index_visible).copied() {
                    canvas_page_rects.push(rect);
                }
            }

            page_count = all_page_rects.len().max(1);
            page_index = first_visible_index.map(|idx| idx + 1).unwrap_or(1);

            canvas_show_margin_guides = tab.canvas.show_margin_guides;
            canvas_cursor_visible = tab.canvas.cursor.visible;
            canvas_scrollbar_visible = tab.canvas.scrollbar.visible;
            canvas_scrollbar_alpha = tab.canvas.scrollbar.alpha;
            canvas_viewport_width = tab.canvas.viewport.width;
            canvas_viewport_height = tab.canvas.viewport.height;
            canvas_scroll_x = tab.canvas.scroll.x;
            canvas_scroll_y = tab.canvas.scroll.y;
            let content_size = tab.canvas.content_size(&tab.document);
            canvas_content_width = content_size.width.max(1.0);
            canvas_content_height = content_size.height.max(1.0);

            view_mode = match tab.canvas.layout_mode {
                PageLayoutMode::SinglePage => "Single Page".to_string(),
                PageLayoutMode::Continuous => "Continuous".to_string(),
                PageLayoutMode::ReadMode => "Read Mode".to_string(),
            };
            zoom_percent = (tab.canvas.zoom * 100.0).round().clamp(25.0, 500.0) as u16;
            file_format = format!("{:?}", tab.document.metadata.format).to_uppercase();
            column = tab.cursor.primary.offset.saturating_add(1);
            line = 1;
            current_block = Some(tab.cursor.primary.block_id);
            canvas_preview_lines = collect_preview_lines(&tab.document, 40);
            canvas_images = collect_canvas_image_overlays(tab, selected_image_id, image_cache);
            canvas_tables = collect_canvas_table_overlays(tab);
        }
    }
    if active_is_welcome {
        canvas_preview_lines = welcome_preview_lines(state);
        canvas_cursor_visible = false;
        canvas_images.clear();
        canvas_tables.clear();
    }
    state.sidebar.set_current_outline_block(current_block);
    state.canvas_image_overlays = canvas_images.clone();
    state.canvas_table_overlays = canvas_tables.clone();
    if let Some(renderer) = &mut state.renderer {
        renderer.update_image_cache_stats(state.image_cache.stats());
    }

    state.statusbar.set_info(StatusBarInfo {
        page_index,
        page_count,
        word_count,
        character_count,
        view_mode: view_mode.clone(),
        line,
        column,
        zoom_percent,
        file_format: file_format.clone(),
        ..StatusBarInfo::default()
    });

    let active_sidebar_panel = match state.sidebar.active_panel {
        SidebarPanel::Files => "Files",
        SidebarPanel::Outline => "Outline",
        SidebarPanel::Bookmarks => "Bookmarks",
        SidebarPanel::SearchResults => "Search Results",
    };
    let sidebar_summary = match state.sidebar.active_panel {
        SidebarPanel::Files => state
            .sidebar
            .file_root
            .as_ref()
            .map(|p| p.display().to_string())
            .unwrap_or_else(|| "No folder".to_string()),
        SidebarPanel::Outline => format!("{} headings", state.sidebar.outline_items.len()),
        SidebarPanel::Bookmarks => format!("{} bookmarks", state.sidebar.bookmarks.len()),
        SidebarPanel::SearchResults => state.sidebar.search_summary(),
    };
    let sidebar_rows = state.sidebar.panel_rows(24);
    let command_palette_open = state.command_palette.is_open();
    let command_palette_query = state.command_palette.query.clone();
    let command_palette_selected = state.command_palette.selected;
    let command_palette_results = state.command_palette.result_labels(8);
    let settings_visible = state.settings_dialog.is_open();
    let settings_query = state.settings_dialog.search_query().to_string();
    let settings_category = state.settings_dialog.selected_category().title().to_string();
    let settings_categories = state
        .settings_dialog
        .visible_categories()
        .into_iter()
        .map(|category| category.title().to_string())
        .collect::<Vec<_>>();
    let settings_rows = state.settings_dialog.setting_rows();
    let settings_selected_row = state.settings_dialog.selected_setting_row();
    let settings_conflicts = state.settings_dialog.has_conflicting_shortcuts();
    let settings_save_error = state
        .settings_dialog
        .last_save_error()
        .unwrap_or_default()
        .to_string();
    let find_visible = state.find_replace.find_visible;
    let replace_visible = state.find_replace.replace_visible;
    let find_query = state.find_replace.query.clone();
    let replace_query = state.find_replace.replacement.clone();
    let find_result_count = state.find_replace.result_count_text.clone();
    let find_case_sensitive = state.find_replace.options.case_sensitive;
    let find_whole_word = state.find_replace.options.whole_word;
    let find_regex = state.find_replace.options.regex;
    let find_total = state.find_replace.results.len();
    let find_current = if find_total == 0 {
        0
    } else {
        state.find_replace.current_index.saturating_add(1)
    };
    let find_preview = state
        .find_replace
        .current_result()
        .map(|m| replacement_preview(m, state.find_replace.replacement.as_str()))
        .unwrap_or_default();
    let find_capture_groups = state
        .find_replace
        .current_result()
        .map(|m| {
            m.capture_groups
                .iter()
                .enumerate()
                .map(|(idx, (s, e))| format!("${}: {}-{}", idx + 1, s, e))
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let selected_image = state
        .selected_image
        .and_then(|id| active_image_ref(state, id).map(|image| (id, image)))
        .map(|(id, image)| {
            let alignment = match image.alignment {
                ImageAlignment::Inline => "Inline",
                ImageAlignment::Left => "Left",
                ImageAlignment::Center => "Center",
                ImageAlignment::Right => "Right",
                ImageAlignment::Float => "Float",
            };
            let border = if image.border.is_some() {
                "Border on"
            } else {
                "Border off"
            };
            (
                id,
                format!("{:.0} x {:.0} pt", image.width, image.height),
                format!("{alignment} | {border}"),
                if image.alt_text.is_empty() {
                    "No alt text".to_string()
                } else {
                    image.alt_text.clone()
                },
            )
        });
    let selected_table_meta = state.selected_table.and_then(|id| {
        active_table_ref(state, id).map(|table| {
            let mode = match state.table_selection_mode {
                Some(TableSelectionMode::Cell(cell)) => format!("Cell {},{}", cell.row + 1, cell.col + 1),
                Some(TableSelectionMode::Row(row)) => format!("Row {}", row + 1),
                Some(TableSelectionMode::Column(col)) => format!("Column {}", col + 1),
                Some(TableSelectionMode::Table) => "Whole table".to_string(),
                None => "No selection".to_string(),
            };
            (
                id.0,
                table.rows.len(),
                table.column_widths.len(),
                mode,
                format!("{:?}", table.style),
            )
        })
    });

    let visible_start = state.tabs.overflow_offset.min(state.tabs.tabs.len());
    let visible_len = state.tabs.tab_rects.len();
    let visible_end = (visible_start + visible_len).min(state.tabs.tabs.len());
    let tab_titles = if visible_start < visible_end {
        state.tabs.tabs[visible_start..visible_end]
            .iter()
            .map(tab_shell_title)
            .collect::<Vec<_>>()
    } else {
        Vec::new()
    };
    let active_tab = if state.tabs.active >= visible_start && state.tabs.active < visible_end {
        state.tabs.active - visible_start
    } else {
        usize::MAX
    };

    ShellRenderState {
        ui_scale: state
            .app_state
            .settings
            .appearance
            .ui_scale
            .as_factor()
            .clamp(1.0, 2.0),
        show_tabs: state.app_state.show_tabs,
        show_sidebar: state.app_state.show_sidebar,
        sidebar_width: state.app_state.sidebar_width,
        sidebar_resizing: state.sidebar_resizing,
        show_toolbar: state.app_state.show_toolbar,
        show_statusbar: state.app_state.show_statusbar,
        status_text: state.app_state.status_text.clone(),
        tab_titles,
        active_tab,
        tab_has_overflow_left: state.tabs.overflow_offset > 0,
        tab_has_overflow_right: state.tabs.overflow_offset + state.tabs.tab_rects.len() < state.tabs.tabs.len(),
        toolbar_labels: state
            .toolbar
            .buttons
            .iter()
            .filter(|b| !b.label.is_empty())
            .map(|b| b.label.to_string())
            .collect(),
        active_sidebar_panel: active_sidebar_panel.to_string(),
        sidebar_summary,
        sidebar_rows,
        command_palette_open,
        command_palette_query,
        command_palette_results,
        command_palette_selected,
        settings_visible,
        settings_query,
        settings_category,
        settings_categories,
        settings_rows,
        settings_selected_row,
        settings_conflicts,
        settings_save_error,
        table_picker_visible: state.table_picker_visible,
        table_picker_rows: state.table_picker_rows,
        table_picker_cols: state.table_picker_cols,
        table_picker_custom_rows: state.table_picker_custom_rows.clone(),
        table_picker_custom_cols: state.table_picker_custom_cols.clone(),
        table_picker_custom_focus_rows: state.table_picker_custom_focus_rows,
        find_visible,
        replace_visible,
        find_query,
        replace_query,
        find_result_count,
        find_case_sensitive,
        find_whole_word,
        find_regex,
        find_preview,
        find_current,
        find_total,
        find_capture_groups,
        goto_visible: state.goto_visible,
        goto_input: state.goto_input.clone(),
        status_left: state.statusbar.left_text(),
        status_right: state.statusbar.right_text(),
        canvas_background: from_canvas_preference(
            &state.app_state.settings.appearance.canvas_background,
        ),
        canvas_page_rects,
        canvas_preview_lines,
        canvas_show_margin_guides,
        canvas_cursor_visible,
        canvas_scrollbar_visible,
        canvas_scrollbar_alpha,
        canvas_viewport_width,
        canvas_viewport_height,
        canvas_content_width,
        canvas_content_height,
        canvas_scroll_x,
        canvas_scroll_y,
        canvas_images: canvas_images
            .iter()
            .map(|overlay| crate::render::d2d::CanvasImageShellItem {
                block_id: overlay.block_id.0,
                rect: overlay.rect,
                selected: state.selected_image == Some(overlay.block_id),
                interpolation: overlay.interpolation.clone(),
                alt_text: overlay.alt_text.clone(),
            })
            .collect(),
        image_toolbar_visible: state.selected_image.is_some(),
        image_properties_visible: state.image_properties_visible,
        image_selected_size: selected_image
            .as_ref()
            .map(|(_, size, _, _)| size.clone())
            .unwrap_or_default(),
        image_selected_meta: selected_image
            .as_ref()
            .map(|(_, _, meta, _)| meta.clone())
            .unwrap_or_default(),
        image_selected_alt_text: selected_image
            .as_ref()
            .map(|(_, _, _, alt)| alt.clone())
            .unwrap_or_default(),
        canvas_tables: canvas_tables
            .iter()
            .map(|overlay| {
                let selected = state.selected_table == Some(overlay.table_id);
                let mut mode = 0u8;
                if selected {
                    mode = match state.table_selection_mode {
                        Some(TableSelectionMode::Cell(_)) => 1,
                        Some(TableSelectionMode::Row(_)) => 2,
                        Some(TableSelectionMode::Column(_)) => 3,
                        Some(TableSelectionMode::Table) => 4,
                        None => 0,
                    };
                }
                let selection = if selected {
                    state.table_selection_range.clone()
                } else {
                    None
                };
                let (start_row, start_col, end_row, end_col) = selection
                    .map(|sel| {
                        (
                            sel.start.row.min(sel.end.row),
                            sel.start.col.min(sel.end.col),
                            sel.start.row.max(sel.end.row),
                            sel.start.col.max(sel.end.col),
                        )
                    })
                    .unwrap_or((0, 0, 0, 0));

                crate::render::d2d::CanvasTableShellItem {
                    table_id: overlay.table_id.0,
                    rect: overlay.rect,
                    rows: overlay.rows,
                    cols: overlay.cols,
                    cell_w: overlay.cell_w,
                    cell_h: overlay.cell_h,
                    header_h: overlay.header_h,
                    gutter_w: overlay.gutter_w,
                    selected,
                    selection_mode: mode,
                    selection_start_row: start_row,
                    selection_start_col: start_col,
                    selection_end_row: end_row,
                    selection_end_col: end_col,
                }
            })
            .collect(),
        table_selected_meta: selected_table_meta
            .as_ref()
            .map(|(_, rows, cols, mode, style)| format!("{rows}x{cols} | {mode} | {style}"))
            .unwrap_or_default(),
        table_selected_id: selected_table_meta.map(|(id, _, _, _, _)| id).unwrap_or_default(),
    }
}

fn toolbar_action_text(action: ToolbarAction) -> &'static str {
    match action {
        ToolbarAction::FileMenu => "File menu",
        ToolbarAction::Cut => "Cut",
        ToolbarAction::Copy => "Copy",
        ToolbarAction::Paste => "Paste",
        ToolbarAction::Undo => "Undo",
        ToolbarAction::Redo => "Redo",
        ToolbarAction::Bold => "Bold",
        ToolbarAction::Italic => "Italic",
        ToolbarAction::Underline => "Underline",
        ToolbarAction::Strikethrough => "Strikethrough",
        ToolbarAction::FontFamily => "Font family",
        ToolbarAction::FontSize => "Font size",
        ToolbarAction::TextColor => "Text color",
        ToolbarAction::AlignLeft => "Align left",
        ToolbarAction::AlignCenter => "Align center",
        ToolbarAction::AlignRight => "Align right",
        ToolbarAction::AlignJustify => "Align justify",
        ToolbarAction::List => "List",
        ToolbarAction::Heading => "Heading",
        ToolbarAction::InsertImage => "Insert image",
        ToolbarAction::InsertLink => "Insert link",
        ToolbarAction::InsertTable => "Insert table",
        ToolbarAction::CommandPalette => "Command palette",
        ToolbarAction::More => "More actions",
    }
}

unsafe extern "system" fn window_proc(
    hwnd: HWND,
    message: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match message {
        WM_NCCREATE => {
            let create_struct = unsafe { &*(lparam.0 as *const CREATESTRUCTW) };
            let state_ptr = create_struct.lpCreateParams as *mut WindowState;
            unsafe {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, state_ptr as isize);
            }
            unsafe { DefWindowProcW(hwnd, message, wparam, lparam) }
        }
        WM_CREATE => {
            let create_begin = Instant::now();
            unsafe { DragAcceptFiles(hwnd, true) };

            let mut client = RECT::default();
            let _ = unsafe { GetClientRect(hwnd, &mut client) };

            if let Some(state) = unsafe { state_from_hwnd(hwnd) } {
                let _ = sync_theme_from_settings(state);
                unsafe { apply_window_effects(hwnd, state.theme.is_dark) };
                let width = (client.right - client.left).max(1) as u32;
                let height = (client.bottom - client.top).max(1) as u32;

                match D2DRenderer::new(hwnd, width, height, state.dpi, state.theme.clone()) {
                    Ok(renderer) => state.renderer = Some(renderer),
                    Err(error) => {
                        eprintln!("Renderer initialization failed: {error:?}");
                    }
                }
                relayout_shell(state, width as f32, height as f32);

                let mut opened_any = false;
                if !state.startup_files.is_empty() {
                    state.app_state.status_text = format!(
                        "Opening {} file(s) from command line in background",
                        state.startup_files.len()
                    );
                    opened_any = true;
                }

                let recovered = restore_recovery_tabs(state);
                if recovered > 0 {
                    state.app_state.status_text =
                        format!("Recovered {} unsaved document(s)", recovered);
                    opened_any = true;
                }

                if !opened_any {
                    state.app_state.status_text = "Welcome to Doco".to_string();
                }

                sync_sidebar_with_active_tab(state);
            }
            emit_startup_marker("window_create", create_begin.elapsed().as_secs_f64() * 1000.0);

            LRESULT(0)
        }
        WM_SIZE => {
            if let Some(state) = unsafe { state_from_hwnd(hwnd) } {
                let width = (lparam.0 as u32 & 0xFFFF) as u32;
                let height = ((lparam.0 as u32 >> 16) & 0xFFFF) as u32;
                if let Some(renderer) = &mut state.renderer {
                    let _ = renderer.resize(width, height);
                }
                relayout_shell(state, width as f32, height as f32);
            }

            LRESULT(0)
        }
        WM_DPICHANGED => {
            if let Some(state) = unsafe { state_from_hwnd(hwnd) } {
                state.dpi = (wparam.0 as u32 & 0xFFFF) as f32;

                let suggested = lparam.0 as *const RECT;
                if !suggested.is_null() {
                    let suggested = unsafe { *suggested };
                    let _ = unsafe {
                        SetWindowPos(
                            hwnd,
                            None,
                            suggested.left,
                            suggested.top,
                            suggested.right - suggested.left,
                            suggested.bottom - suggested.top,
                            SWP_NOZORDER | SWP_NOACTIVATE,
                        )
                    };
                }

                if let Some(renderer) = &mut state.renderer {
                    renderer.set_dpi(state.dpi);
                }

                let mut client = RECT::default();
                let _ = unsafe { GetClientRect(hwnd, &mut client) };
                relayout_shell(
                    state,
                    (client.right - client.left).max(0) as f32,
                    (client.bottom - client.top).max(0) as f32,
                );
            }

            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
            LRESULT(0)
        }
        WM_SETTINGCHANGE => {
            if let Some(state) = unsafe { state_from_hwnd(hwnd) } {
                let changed = sync_theme_from_settings(state);
                if changed {
                    if let Some(renderer) = &mut state.renderer {
                        renderer.set_theme(state.theme.clone());
                    }
                    unsafe { apply_window_effects(hwnd, state.theme.is_dark) };
                    state.app_state.status_text = format!("Theme updated: {}", state.theme.name);
                }
            }
            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
            LRESULT(0)
        }
        WM_PAINT => {
            let mut paint = PAINTSTRUCT::default();
            let _ = unsafe { BeginPaint(hwnd, &mut paint) };

            if let Some(state) = unsafe { state_from_hwnd(hwnd) } {
                let now = Instant::now();
                let dt = (now - state.last_ui_tick).as_secs_f32().clamp(0.0, 0.25);
                state.last_ui_tick = now;
                state.sidebar.tick(dt);
                state.command_palette.tick(dt);
                if state.app_state.show_settings != state.settings_dialog.is_open() {
                    state
                        .settings_dialog
                        .set_visible(state.app_state.show_settings);
                }
                state.settings_dialog.tick();
                sync_runtime_from_settings(state, hwnd);
                state.app_state.show_settings = state.settings_dialog.is_open();
                let mut needs_next_frame = false;
                if let Some(tab) = state.tabs.active_tab_mut() {
                    needs_next_frame |= tab.canvas.update(dt);
                    tab.canvas.clamp_scroll(&tab.document);
                    if let Ok(Some(path)) = state.app_state.autosave.tick(&tab.document) {
                        state.app_state.status_text =
                            format!("Auto-saved recovery snapshot: {}", path.display());
                        send_toast_notification(
                            "Auto-recovery saved",
                            format!("{}", path.display()).as_str(),
                        );
                    }
                }
                if state.find_replace.should_live_update(now) {
                    let refreshed = refresh_find_results(state);
                    needs_next_frame |=
                        refreshed || state.find_replace.has_pending_background_search();
                }
                if state.find_replace.has_pending_background_search() {
                    let chunk_changed = process_find_background_search(state, 256);
                    needs_next_frame = true;
                    if chunk_changed {
                        state.app_state.status_text = state.find_replace.result_count_text.clone();
                    }
                }
                let background =
                    from_canvas_preference(&state.app_state.settings.appearance.canvas_background);
                if matches!(background.kind, BackgroundKind::AnimatedGradient { .. }) {
                    needs_next_frame = true;
                }

                if !state.startup_files.is_empty() {
                    let startup_chunk_begin = Instant::now();
                    if process_startup_file_queue(state) {
                        sync_sidebar_with_active_tab(state);
                        needs_next_frame = true;
                        emit_startup_marker(
                            "startup_file_open",
                            startup_chunk_begin.elapsed().as_secs_f64() * 1000.0,
                        );
                    }
                }

                let shell = build_shell_render_state(state);
                if let Some(renderer) = &mut state.renderer {
                    let _ = renderer.render(&shell);
                }

                if needs_next_frame {
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                }
            }

            let _ = unsafe { EndPaint(hwnd, &paint) };
            LRESULT(0)
        }
        WM_MOUSEWHEEL => {
            if let Some(state) = unsafe { state_from_hwnd(hwnd) } {
                let ctrl_down = unsafe { GetKeyState(VK_CONTROL.0 as i32) } < 0;
                let shift_down = unsafe { GetKeyState(VK_SHIFT.0 as i32) } < 0;
                let delta = ((wparam.0 >> 16) as i16 as f32) / 120.0;

                let mut client = RECT::default();
                let _ = unsafe { GetClientRect(hwnd, &mut client) };
                let client_width = (client.right - client.left).max(1) as f32;
                let client_height = (client.bottom - client.top).max(1) as f32;
                let (canvas_w, canvas_h) = canvas_viewport_size(state, client_width, client_height);
                let cursor_in_canvas = UiPoint {
                    x: canvas_w * 0.5,
                    y: canvas_h * 0.5,
                };

                if state.settings_dialog.is_open() && !state.command_palette.is_open() {
                    let event = UiInputEvent::MouseWheel {
                        delta,
                        position: cursor_in_canvas,
                    };
                    let _ = state.settings_dialog.handle_input(&event);
                    sync_runtime_from_settings(state, hwnd);
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if let Some(tab) = state.tabs.active_tab_mut() {
                    tab.canvas.set_viewport(canvas_w, canvas_h);
                    if ctrl_down {
                        tab.canvas.handle_mouse_wheel(delta, true, cursor_in_canvas);
                        state.app_state.status_text =
                            format!("Zoom: {}%", (tab.canvas.zoom_target * 100.0).round() as u16);
                    } else if shift_down {
                        tab.canvas.handle_horizontal_wheel(delta);
                        state.app_state.status_text = "Horizontal scroll".to_string();
                    } else {
                        tab.canvas
                            .handle_mouse_wheel(delta, false, cursor_in_canvas);
                        state.app_state.status_text = "Scroll".to_string();
                    }
                    tab.canvas.clamp_scroll(&tab.document);
                }

                let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                return LRESULT(0);
            }
            unsafe { DefWindowProcW(hwnd, message, wparam, lparam) }
        }
        WM_KEYDOWN => {
            if let Some(state) = unsafe { state_from_hwnd(hwnd) } {
                let ctrl_down = unsafe { GetKeyState(VK_CONTROL.0 as i32) } < 0;
                let shift_down = unsafe { GetKeyState(VK_SHIFT.0 as i32) } < 0;
                let vk = wparam.0 as u32;

                if ctrl_down && shift_down && vk == 0x44 {
                    state.debug_panel_visible = !state.debug_panel_visible;
                    if let Some(renderer) = &mut state.renderer {
                        renderer.set_debug_panel_visible(state.debug_panel_visible);
                    }
                    state.app_state.show_debug_panel = state.debug_panel_visible;
                    state.app_state.status_text = if state.debug_panel_visible {
                        "Debug panel enabled".to_string()
                    } else {
                        "Debug panel hidden".to_string()
                    };
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if (ctrl_down && shift_down && vk == 0x50) || vk == 0x70 {
                    state.command_palette.open();
                    state
                        .command_palette
                        .refresh_results(Some(&state.app_state));
                    state.app_state.status_text = "Command palette".to_string();
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if state.command_palette.is_open() {
                    let event = UiInputEvent::KeyDown(vk);
                    let mut handled = state.command_palette.handle_input(&event);
                    if vk == 0x0D {
                        handled |= state.command_palette.execute_selected(&mut state.app_state);
                        if handled && state.app_state.status_text == "Find" {
                            state.find_replace.open_find();
                            state.find_focus = FindFieldFocus::Query;
                            refresh_find_results(state);
                        } else if handled && state.app_state.status_text == "Replace" {
                            state.find_replace.open_replace();
                            state.find_focus = FindFieldFocus::Replacement;
                            refresh_find_results(state);
                        } else if handled && state.app_state.status_text == "Insert image" {
                            if let Some(path) = pick_image_file(hwnd) {
                                match insert_image_from_path(state, &path) {
                                    Ok(_) => {
                                        state.app_state.status_text = format!(
                                            "Inserted image: {}",
                                            path.file_name()
                                                .and_then(|v| v.to_str())
                                                .unwrap_or("image")
                                        );
                                    }
                                    Err(err) => {
                                        state.app_state.status_text =
                                            format!("Insert image failed: {err}");
                                    }
                                }
                            } else {
                                state.app_state.status_text = "Insert image cancelled".to_string();
                            }
                        } else if handled && state.app_state.status_text == "Insert table" {
                            open_table_picker(state);
                            state.app_state.status_text = "Insert table (picker)".to_string();
                        } else if handled
                            && (state.app_state.status_text == "Saved"
                                || state.app_state.status_text == "Save")
                        {
                            let _ = save_active_document(state, hwnd, false);
                        } else if handled && state.app_state.status_text == "Save As" {
                            let _ = save_active_document(state, hwnd, true);
                        } else if handled && state.app_state.status_text == "Export PDF" {
                            let _ = export_active_document(state, hwnd, "pdf");
                        } else if handled && state.app_state.status_text == "Close tab" {
                            let active_index = state.tabs.active;
                            let _ = close_tab_with_prompt(state, hwnd, active_index);
                        }
                    }
                    if handled {
                        let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                        return LRESULT(0);
                    }
                }

                if state.app_state.show_settings && !state.settings_dialog.is_open() {
                    set_settings_visible(state, true);
                    sync_runtime_from_settings(state, hwnd);
                }

                if state.settings_dialog.is_open() && !state.command_palette.is_open() {
                    if ctrl_down
                        && !shift_down
                        && vk == 0x52
                        && state.settings_dialog.selected_category()
                            == SettingsCategory::KeyboardShortcuts
                    {
                        state.settings_dialog.reset_shortcuts();
                        sync_runtime_from_settings(state, hwnd);
                        state.app_state.status_text = "Shortcuts reset to defaults".to_string();
                        let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                        return LRESULT(0);
                    }

                    let event = UiInputEvent::KeyDown(vk);
                    let handled_settings = state.settings_dialog.handle_input(&event);
                    if handled_settings {
                        state.app_state.show_settings = state.settings_dialog.is_open();
                        sync_runtime_from_settings(state, hwnd);
                        if !state.settings_dialog.is_open() {
                            state.app_state.status_text = "Settings closed".to_string();
                        } else {
                            state.app_state.status_text = "Settings updated".to_string();
                        }
                    }
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if state.table_picker_visible {
                    match vk {
                        0x1B => {
                            state.table_picker_visible = false;
                            state.app_state.status_text = "Insert table cancelled".to_string();
                            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                            return LRESULT(0);
                        }
                        0x0D => {
                            if let Some(id) = insert_table_from_picker(state) {
                                state.app_state.status_text = format!(
                                    "Inserted table {} ({}x{})",
                                    id.0, state.table_picker_rows, state.table_picker_cols
                                );
                            } else {
                                state.app_state.status_text = "Insert table failed".to_string();
                            }
                            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                            return LRESULT(0);
                        }
                        0x09 => {
                            state.table_picker_custom_focus_rows = !state.table_picker_custom_focus_rows;
                            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                            return LRESULT(0);
                        }
                        0x08 => {
                            if state.table_picker_custom_focus_rows {
                                remove_last_char(&mut state.table_picker_custom_rows);
                                state.table_picker_rows = parse_table_picker_custom(
                                    state.table_picker_custom_rows.as_str(),
                                    state.table_picker_rows,
                                )
                                .clamp(1, 10);
                                state.table_picker_custom_rows = state.table_picker_rows.to_string();
                            } else {
                                remove_last_char(&mut state.table_picker_custom_cols);
                                state.table_picker_cols = parse_table_picker_custom(
                                    state.table_picker_custom_cols.as_str(),
                                    state.table_picker_cols,
                                )
                                .clamp(1, 10);
                                state.table_picker_custom_cols = state.table_picker_cols.to_string();
                            }
                            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                            return LRESULT(0);
                        }
                        0x25 => {
                            state.table_picker_cols = state.table_picker_cols.saturating_sub(1).max(1);
                            state.table_picker_custom_cols = state.table_picker_cols.to_string();
                            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                            return LRESULT(0);
                        }
                        0x27 => {
                            state.table_picker_cols = (state.table_picker_cols + 1).min(10);
                            state.table_picker_custom_cols = state.table_picker_cols.to_string();
                            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                            return LRESULT(0);
                        }
                        0x26 => {
                            state.table_picker_rows = state.table_picker_rows.saturating_sub(1).max(1);
                            state.table_picker_custom_rows = state.table_picker_rows.to_string();
                            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                            return LRESULT(0);
                        }
                        0x28 => {
                            state.table_picker_rows = (state.table_picker_rows + 1).min(10);
                            state.table_picker_custom_rows = state.table_picker_rows.to_string();
                            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                            return LRESULT(0);
                        }
                        _ => {}
                    }
                }

                if ctrl_down && vk == 0x53 {
                    if shift_down {
                        let _ = save_active_document(state, hwnd, true);
                    } else {
                        let _ = save_active_document(state, hwnd, false);
                    }
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if !state.command_palette.is_open()
                    && !state.find_replace.find_visible
                    && !state.goto_visible
                    && apply_table_shortcut(state, vk, ctrl_down, shift_down)
                {
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if ctrl_down
                    && !shift_down
                    && vk == 0x56
                    && !state.find_replace.find_visible
                    && !state.command_palette.is_open()
                    && !state.goto_visible
                {
                    match insert_image_from_clipboard(state) {
                        Ok(id) => {
                            state.app_state.status_text = format!("Pasted image {}", id.0);
                            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                            return LRESULT(0);
                        }
                        Err(_) => {}
                    }
                }

                if ctrl_down && shift_down && vk == 0x46 {
                    if state.find_replace.query.trim().is_empty() {
                        state.find_replace.open_find();
                        state.find_focus = FindFieldFocus::Query;
                        state.app_state.status_text =
                            "Set a Find query, then press Ctrl+Shift+F to search all tabs".to_string();
                    } else {
                        let query = state.find_replace.query.clone();
                        let (tabs_with_matches, total_matches) =
                            find_in_all_open_tabs(state, query.as_str());
                        state.app_state.status_text = format!(
                            "Find all: '{}' matched {} times in {} tab(s)",
                            query, total_matches, tabs_with_matches
                        );
                    }
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if ctrl_down && !shift_down && vk == 0x46 {
                    state.find_replace.open_find();
                    state.find_focus = FindFieldFocus::Query;
                    refresh_find_results(state);
                    state.app_state.status_text = "Find".to_string();
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if ctrl_down && !shift_down && vk == 0x48 {
                    state.find_replace.open_replace();
                    state.find_focus = FindFieldFocus::Replacement;
                    refresh_find_results(state);
                    state.app_state.status_text = "Replace".to_string();
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if ctrl_down && !shift_down && vk == 0x47 {
                    state.goto_visible = true;
                    state.goto_input.clear();
                    state.app_state.status_text = "Go to line/page".to_string();
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if state.goto_visible {
                    match vk {
                        0x1B => {
                            state.goto_visible = false;
                            state.goto_input.clear();
                            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                            return LRESULT(0);
                        }
                        0x08 => {
                            remove_last_char(&mut state.goto_input);
                            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                            return LRESULT(0);
                        }
                        0x0D => {
                            let target = state.goto_input.parse::<usize>().ok().unwrap_or(0);
                            if jump_to_line_or_page(state, target) {
                                state.app_state.status_text = format!("Jumped to {}", target);
                            } else {
                                state.app_state.status_text = "Target not found".to_string();
                            }
                            state.goto_visible = false;
                            state.goto_input.clear();
                            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                            return LRESULT(0);
                        }
                        _ => {}
                    }
                }

                if vk == 0x72 {
                    if navigate_find_result(state, shift_down) {
                        let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                        return LRESULT(0);
                    }
                }

                if state.find_replace.find_visible {
                    let mut handled_find = false;
                    match vk {
                        0x1B => {
                            state.find_replace.close();
                            handled_find = true;
                        }
                        0x09 => {
                            state.find_focus =
                                match (state.find_focus, state.find_replace.replace_visible) {
                                    (FindFieldFocus::Query, true) => FindFieldFocus::Replacement,
                                    _ => FindFieldFocus::Query,
                                };
                            handled_find = true;
                        }
                        0x08 => {
                            match state.find_focus {
                                FindFieldFocus::Query => {
                                    remove_last_char(&mut state.find_replace.query);
                                    state
                                        .find_replace
                                        .set_query(state.find_replace.query.clone());
                                }
                                FindFieldFocus::Replacement => {
                                    remove_last_char(&mut state.find_replace.replacement);
                                }
                            }
                            handled_find = true;
                        }
                        0x0D => {
                            if ctrl_down && state.find_replace.replace_visible {
                                let replaced = if shift_down {
                                    replace_all_matches(state)
                                } else {
                                    replace_current_match(state)
                                };
                                state.app_state.status_text = if replaced == 1 {
                                    "Replaced 1 occurrence".to_string()
                                } else {
                                    format!("Replaced {} occurrences", replaced)
                                };
                            } else {
                                let _ = navigate_find_result(state, shift_down);
                            }
                            handled_find = true;
                        }
                        _ => {}
                    }

                    if ctrl_down && shift_down && vk == 0x43 {
                        state.find_replace.options.case_sensitive =
                            !state.find_replace.options.case_sensitive;
                        state.find_replace.invalidate_cache();
                        state.find_replace.pending_live_update = true;
                        state.find_replace.last_input_at = Instant::now()
                            - std::time::Duration::from_millis(state.find_replace.debounce_ms);
                        handled_find = true;
                    }
                    if ctrl_down && shift_down && vk == 0x57 {
                        state.find_replace.options.whole_word =
                            !state.find_replace.options.whole_word;
                        state.find_replace.invalidate_cache();
                        state.find_replace.pending_live_update = true;
                        state.find_replace.last_input_at = Instant::now()
                            - std::time::Duration::from_millis(state.find_replace.debounce_ms);
                        handled_find = true;
                    }
                    if ctrl_down && shift_down && vk == 0x52 {
                        state.find_replace.options.regex = !state.find_replace.options.regex;
                        state.find_replace.invalidate_cache();
                        state.find_replace.pending_live_update = true;
                        state.find_replace.last_input_at = Instant::now()
                            - std::time::Duration::from_millis(state.find_replace.debounce_ms);
                        handled_find = true;
                    }

                    if handled_find {
                        refresh_find_results(state);
                        let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                        return LRESULT(0);
                    }
                }

                if state.selected_image.is_some() {
                    if vk == VK_DELETE.0 as u32 {
                        if delete_selected_image(state) {
                            state.app_state.status_text = "Image deleted".to_string();
                            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                            return LRESULT(0);
                        }
                    }

                    if ctrl_down && !shift_down && vk == 0x52 {
                        if let Some(path) = pick_image_file(hwnd) {
                            if let Some(selected) = state.selected_image {
                                if delete_selected_image(state) {
                                    if let Ok(inserted) = insert_image_from_path(state, &path) {
                                        state.selected_image = Some(inserted);
                                        state.image_properties_visible = false;
                                        state.app_state.status_text = format!(
                                            "Replaced image {} with {}",
                                            selected.0,
                                            path.file_name()
                                                .and_then(|v| v.to_str())
                                                .unwrap_or("image")
                                        );
                                        let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                                        return LRESULT(0);
                                    }
                                }
                            }
                        }
                    }

                    if ctrl_down && !shift_down && vk == 0x4C {
                        if align_selected_image(state, ImageAlignment::Left) {
                            state.app_state.status_text = "Image aligned left".to_string();
                            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                            return LRESULT(0);
                        }
                    }
                    if ctrl_down && !shift_down && vk == 0x45 {
                        if align_selected_image(state, ImageAlignment::Center) {
                            state.app_state.status_text = "Image aligned center".to_string();
                            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                            return LRESULT(0);
                        }
                    }
                    if ctrl_down && !shift_down && vk == 0x49 {
                        if align_selected_image(state, ImageAlignment::Right) {
                            state.app_state.status_text = "Image aligned right".to_string();
                            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                            return LRESULT(0);
                        }
                    }
                    if ctrl_down && shift_down && vk == 0x42 {
                        if toggle_selected_image_border(state) {
                            state.app_state.status_text = "Image border toggled".to_string();
                            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                            return LRESULT(0);
                        }
                    }
                }

                if ctrl_down && !shift_down && vk == 0x50 {
                    state.print_state.request_print_dialog();
                    if let Some(result) = open_print_dialog(hwnd) {
                        state.print_state.page_range = result.page_range;
                        state.print_state.include_header_footer = true;
                        state.print_state.complete_print();
                        state.app_state.status_text = if let Some((from, to)) = result.page_range {
                            format!("Print queued (pages {from}-{to}, copies {})", result.copies)
                        } else {
                            format!("Print queued (all pages, copies {})", result.copies)
                        };
                    } else {
                        state.print_state.complete_print();
                        state.app_state.status_text = "Print cancelled".to_string();
                    }
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if ctrl_down && !shift_down && vk == 0xBC {
                    let visible = !state.settings_dialog.is_open();
                    set_settings_visible(state, visible);
                    state.app_state.status_text = if visible {
                        "Settings toggled on".to_string()
                    } else {
                        "Settings toggled off".to_string()
                    };
                    if visible {
                        sync_runtime_from_settings(state, hwnd);
                    }
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if ctrl_down && shift_down && vk == 0x42 {
                    if let Some(tab) = state.tabs.active_tab() {
                        let block_id = tab.cursor.primary.block_id;
                        let snippet = block_snippet(&tab.document, block_id);
                        let bookmark_id = state.sidebar.add_bookmark(block_id, 1, &snippet);
                        state.app_state.status_text = format!("Bookmark added ({bookmark_id})");
                        let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                        return LRESULT(0);
                    }
                }

                if ctrl_down && !shift_down && vk == 0x42 {
                    state.app_state.show_sidebar = !state.app_state.show_sidebar;
                    let show_sidebar = state.app_state.show_sidebar;
                    state.app_state.settings.appearance.show_sidebar = show_sidebar;
                    state
                        .settings_dialog
                        .apply_change(|settings| settings.appearance.show_sidebar = show_sidebar);
                    if !state.app_state.show_sidebar && state.sidebar_resizing {
                        state.sidebar_resizing = false;
                        state.sidebar.resizing = false;
                        let _ = unsafe { ReleaseCapture() };
                    }
                    state.app_state.status_text = if state.app_state.show_sidebar {
                        "Sidebar shown".to_string()
                    } else {
                        "Sidebar hidden".to_string()
                    };
                    let mut client = RECT::default();
                    let _ = unsafe { GetClientRect(hwnd, &mut client) };
                    relayout_shell(
                        state,
                        (client.right - client.left).max(0) as f32,
                        (client.bottom - client.top).max(0) as f32,
                    );
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if ctrl_down && shift_down && vk == 0x54 {
                    state.app_state.show_toolbar = !state.app_state.show_toolbar;
                    let show_toolbar = state.app_state.show_toolbar;
                    state.app_state.settings.appearance.show_toolbar = show_toolbar;
                    state
                        .settings_dialog
                        .apply_change(|settings| settings.appearance.show_toolbar = show_toolbar);
                    state.app_state.status_text = if state.app_state.show_toolbar {
                        "Toolbar shown".to_string()
                    } else {
                        "Toolbar hidden".to_string()
                    };
                    let mut client = RECT::default();
                    let _ = unsafe { GetClientRect(hwnd, &mut client) };
                    relayout_shell(
                        state,
                        (client.right - client.left).max(0) as f32,
                        (client.bottom - client.top).max(0) as f32,
                    );
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if ctrl_down && !shift_down && vk == 0x4C {
                    state.app_state.show_statusbar = !state.app_state.show_statusbar;
                    let show_status_bar = state.app_state.show_statusbar;
                    state.app_state.settings.appearance.show_status_bar = show_status_bar;
                    state.settings_dialog.apply_change(|settings| {
                        settings.appearance.show_status_bar = show_status_bar
                    });
                    state.app_state.status_text = if state.app_state.show_statusbar {
                        "Status bar shown".to_string()
                    } else {
                        "Status bar hidden".to_string()
                    };
                    let mut client = RECT::default();
                    let _ = unsafe { GetClientRect(hwnd, &mut client) };
                    relayout_shell(
                        state,
                        (client.right - client.left).max(0) as f32,
                        (client.bottom - client.top).max(0) as f32,
                    );
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if ctrl_down && !shift_down && vk == 0x54 {
                    let index = open_new_blank_tab(state);
                    let title = state
                        .tabs
                        .tabs
                        .get(index)
                        .map(|tab| tab.title.clone())
                        .unwrap_or_else(|| "New tab".to_string());
                    state.app_state.status_text = format!("Opened {title}");
                    sync_sidebar_with_active_tab(state);
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if ctrl_down && !shift_down && vk == 0x4E {
                    let index = open_new_blank_tab(state);
                    let title = state
                        .tabs
                        .tabs
                        .get(index)
                        .map(|t| t.title.clone())
                        .unwrap_or_else(|| "New tab".to_string());
                    state.app_state.status_text = format!("Opened {title}");
                    sync_sidebar_with_active_tab(state);
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if ctrl_down && !shift_down && vk == 0x57 {
                    let active_index = state.tabs.active;
                    let _ = close_tab_with_prompt(state, hwnd, active_index);
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if ctrl_down && vk == 0x09 {
                    if shift_down {
                        state.tabs.switch_prev();
                    } else {
                        state.tabs.switch_next();
                    }
                    sync_sidebar_with_active_tab(state);
                    let active_title = state
                        .tabs
                        .active_tab()
                        .map(|t| t.title.clone())
                        .unwrap_or_else(|| "Welcome".to_string());
                    state.app_state.status_text = format!("Switched to {active_title}");
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if ctrl_down && !shift_down && (0x31..=0x39).contains(&vk) {
                    let tab_number = (vk - 0x30) as usize;
                    state.tabs.switch_to_number(tab_number);
                    sync_sidebar_with_active_tab(state);
                    let active_title = state
                        .tabs
                        .active_tab()
                        .map(|tab| tab.title.clone())
                        .unwrap_or_else(|| "Welcome".to_string());
                    state.app_state.status_text =
                        format!("Switched to tab {tab_number}: {active_title}");
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if !ctrl_down && state.app_state.show_sidebar {
                    let event = UiInputEvent::KeyDown(vk);
                    if state.sidebar.handle_input(&event) || apply_pending_sidebar_intents(state) {
                        let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                        return LRESULT(0);
                    }
                }
            }

            unsafe { DefWindowProcW(hwnd, message, wparam, lparam) }
        }
        WM_CHAR => {
            if let Some(state) = unsafe { state_from_hwnd(hwnd) } {
                let code = wparam.0 as u32;
                if state.settings_dialog.is_open() && !state.command_palette.is_open() {
                    if let Some(ch) = char::from_u32(code) {
                        let event = UiInputEvent::Char(ch);
                        if state.settings_dialog.handle_input(&event) {
                            sync_runtime_from_settings(state, hwnd);
                            state.app_state.show_settings = state.settings_dialog.is_open();
                            state.app_state.status_text = "Settings updated".to_string();
                            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                            return LRESULT(0);
                        }
                    }
                    return LRESULT(0);
                }

                if state.table_picker_visible {
                    if let Some(ch) = char::from_u32(code) {
                        if ch.is_ascii_digit() {
                            if state.table_picker_custom_focus_rows {
                                state.table_picker_custom_rows.push(ch);
                                state.table_picker_rows =
                                    parse_table_picker_custom(state.table_picker_custom_rows.as_str(), 1)
                                        .clamp(1, 10);
                                state.table_picker_custom_rows = state.table_picker_rows.to_string();
                            } else {
                                state.table_picker_custom_cols.push(ch);
                                state.table_picker_cols =
                                    parse_table_picker_custom(state.table_picker_custom_cols.as_str(), 1)
                                        .clamp(1, 10);
                                state.table_picker_custom_cols = state.table_picker_cols.to_string();
                            }
                            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                            return LRESULT(0);
                        }
                    }
                }

                if state.goto_visible {
                    if let Some(ch) = char::from_u32(code) {
                        if ch.is_ascii_digit() {
                            state.goto_input.push(ch);
                            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                            return LRESULT(0);
                        }
                    }
                }

                if state.find_replace.find_visible {
                    if let Some(ch) = char::from_u32(code) {
                        if !ch.is_control() {
                            match state.find_focus {
                                FindFieldFocus::Query => {
                                    let mut next = state.find_replace.query.clone();
                                    next.push(ch);
                                    state.find_replace.set_query(next);
                                }
                                FindFieldFocus::Replacement => {
                                    let mut next = state.find_replace.replacement.clone();
                                    next.push(ch);
                                    state.find_replace.set_replacement(next);
                                }
                            }
                            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                            return LRESULT(0);
                        }
                    }
                }

                if state.command_palette.is_open() {
                    if let Some(ch) = char::from_u32(code) {
                        if !ch.is_control() {
                            let event = UiInputEvent::Char(ch);
                            if state.command_palette.handle_input(&event) {
                                let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                                return LRESULT(0);
                            }
                        }
                    }
                }
            }
            unsafe { DefWindowProcW(hwnd, message, wparam, lparam) }
        }
        WM_DROPFILES => {
            if let Some(state) = unsafe { state_from_hwnd(hwnd) } {
                let payload = unsafe { extract_drop_payload(HDROP(wparam.0 as *mut c_void)) };
                state.dropped_files = payload.files.clone();

                state.app_state.status_text = match payload.action {
                    DropAction::OpenFilesInTabs => {
                        for path in &payload.files {
                            state.jump_list.add_recent_file(path.clone());
                            let title = path
                                .file_name()
                                .and_then(|v| v.to_str())
                                .unwrap_or("Document")
                                .to_string();
                            let document = load_document_for_path(path.as_path());
                            state
                                .tabs
                                .open_document_tab(title, Some(path.clone()), document);
                        }
                        format!("Drop to open: {} file(s)", payload.files.len())
                    }
                    DropAction::InsertImage => {
                        let (inserted, failed) = insert_images_from_paths(state, &payload.files);
                        if inserted == 0 {
                            format!("Drop image insert failed ({failed} file(s))")
                        } else if failed == 0 {
                            format!("Inserted {} dropped image(s)", inserted)
                        } else {
                            format!("Inserted {} dropped image(s), {} failed", inserted, failed)
                        }
                    }
                    DropAction::Ignore => "Unsupported dropped content".to_string(),
                };
                sync_sidebar_with_active_tab(state);

                let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
            }
            LRESULT(0)
        }
        WM_MOUSEMOVE => {
            if let Some(state) = unsafe { state_from_hwnd(hwnd) } {
                let point = point_from_lparam(lparam);
                if state.settings_dialog.is_open() && !state.command_palette.is_open() {
                    return LRESULT(0);
                }
                if state.table_picker_visible {
                    if update_table_picker_hover(state, point) {
                        let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    }
                    return LRESULT(0);
                }
                if state.sidebar_resizing {
                    let next_width =
                        (point.x + state.sidebar_resize_grab_offset).clamp(200.0, 400.0);
                    if (state.app_state.sidebar_width - next_width).abs() > f32::EPSILON {
                        state.app_state.sidebar_width = next_width;
                        state.app_state.status_text = format!("Sidebar width: {:.0}px", next_width);
                        let mut client = RECT::default();
                        let _ = unsafe { GetClientRect(hwnd, &mut client) };
                        relayout_shell(
                            state,
                            (client.right - client.left).max(0) as f32,
                            (client.bottom - client.top).max(0) as f32,
                        );
                    }
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }
                if state.table_resize.is_some() && update_table_resize(state, point) {
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }
                if state.image_drag.is_some() {
                    let shift_down = unsafe { GetKeyState(VK_SHIFT.0 as i32) } < 0;
                    if update_image_drag(state, point, shift_down) {
                        let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                        return LRESULT(0);
                    }
                }
                let event = UiInputEvent::MouseMove(point);
                let mut dirty = false;

                if state.app_state.show_tabs {
                    dirty |= state.tabs.handle_input(&event);
                }
                if state.app_state.show_toolbar {
                    dirty |= state.toolbar.handle_input(&event);
                }
                if state.app_state.show_sidebar {
                    dirty |= state.sidebar.handle_input(&event);
                }

                if dirty {
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                }
            }
            LRESULT(0)
        }
        WM_LBUTTONDOWN => {
            if let Some(state) = unsafe { state_from_hwnd(hwnd) } {
                let point = point_from_lparam(lparam);
                if state.table_picker_visible {
                    if handle_table_picker_click(state, point) {
                        let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                        return LRESULT(0);
                    }
                }
                if state.command_palette.is_open() {
                    let event = UiInputEvent::MouseDown(point);
                    if state.command_palette.handle_input(&event) {
                        let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                        return LRESULT(0);
                    }
                }
                if state.settings_dialog.is_open() && !state.command_palette.is_open() {
                    let event = UiInputEvent::MouseDown(point);
                    if state.settings_dialog.handle_input(&event) {
                        sync_runtime_from_settings(state, hwnd);
                        state.app_state.status_text = "Settings updated".to_string();
                    }
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }
                if state.app_state.show_tabs {
                    if let Some(index) = state.tabs.tab_close_hit_test(point) {
                        let _ = close_tab_with_prompt(state, hwnd, index);
                        let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                        return LRESULT(0);
                    }
                    if state.tabs.new_button_hit_test(point) {
                        let index = open_new_blank_tab(state);
                        let title = state
                            .tabs
                            .tabs
                            .get(index)
                            .map(|tab| tab.title.clone())
                            .unwrap_or_else(|| "New tab".to_string());
                        state.app_state.status_text = format!("Opened {title}");
                        sync_sidebar_with_active_tab(state);
                        let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                        return LRESULT(0);
                    }
                    if state.tabs.overflow_left_hit_test(point) && state.tabs.scroll_overflow_left() {
                        let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                        return LRESULT(0);
                    }
                    if state.tabs.overflow_right_hit_test(point) && state.tabs.scroll_overflow_right() {
                        let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                        return LRESULT(0);
                    }
                }
                if sidebar_splitter_hit_test(state, point) {
                    state.sidebar_resizing = true;
                    state.sidebar.resizing = true;
                    state.sidebar_resize_grab_offset = state.app_state.sidebar_width - point.x;
                    let _ = unsafe { SetCapture(hwnd) };
                    state.app_state.status_text =
                        format!("Resizing sidebar ({:.0}px)", state.app_state.sidebar_width);
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }
                let event = UiInputEvent::MouseDown(point);
                let mut handled = false;

                if state.app_state.show_tabs {
                    let previous_active = state.tabs.active;
                    handled |= state.tabs.handle_input(&event);
                    if state.tabs.active != previous_active {
                        sync_sidebar_with_active_tab(state);
                        if let Some(tab) = state.tabs.active_tab() {
                            state.app_state.status_text = format!("Switched to {}", tab.title);
                        }
                    }
                }

                if state.app_state.show_toolbar {
                    if let Some(index) = state.toolbar.hit_button(point) {
                        if let Some(action) = state.toolbar.invoke(index) {
                            match action {
                                ToolbarAction::CommandPalette => {
                                    state.command_palette.open();
                                    state
                                        .command_palette
                                        .refresh_results(Some(&state.app_state));
                                    state.app_state.status_text = "Command palette".to_string();
                                }
                                ToolbarAction::InsertImage => {
                                    if let Some(path) = pick_image_file(hwnd) {
                                        match insert_image_from_path(state, &path) {
                                            Ok(id) => {
                                                state.app_state.status_text = format!(
                                                    "Inserted image {} ({})",
                                                    id.0,
                                                    path.file_name()
                                                        .and_then(|v| v.to_str())
                                                        .unwrap_or("image")
                                                );
                                            }
                                            Err(err) => {
                                                state.app_state.status_text =
                                                    format!("Insert image failed: {err}");
                                            }
                                        }
                                    } else {
                                        state.app_state.status_text =
                                            "Insert image cancelled".to_string();
                                    }
                                }
                                ToolbarAction::InsertTable => {
                                    open_table_picker(state);
                                    state.app_state.status_text = "Insert table (picker)".to_string();
                                }
                                ToolbarAction::Paste => {
                                    if let Ok(id) = insert_image_from_clipboard(state) {
                                        state.app_state.status_text =
                                            format!("Pasted image {}", id.0);
                                    } else {
                                        state.app_state.status_text = "Paste".to_string();
                                    }
                                }
                                ToolbarAction::AlignLeft => {
                                    if align_selected_image(state, ImageAlignment::Left) {
                                        state.app_state.status_text =
                                            "Image aligned left".to_string();
                                    } else {
                                        state.app_state.status_text = "Align left".to_string();
                                    }
                                }
                                ToolbarAction::AlignCenter => {
                                    if align_selected_image(state, ImageAlignment::Center) {
                                        state.app_state.status_text =
                                            "Image aligned center".to_string();
                                    } else {
                                        state.app_state.status_text = "Align center".to_string();
                                    }
                                }
                                ToolbarAction::AlignRight => {
                                    if align_selected_image(state, ImageAlignment::Right) {
                                        state.app_state.status_text =
                                            "Image aligned right".to_string();
                                    } else {
                                        state.app_state.status_text = "Align right".to_string();
                                    }
                                }
                                _ => {
                                    state.app_state.status_text =
                                        format!("Toolbar action: {}", toolbar_action_text(action));
                                }
                            }
                        }
                    }
                    handled |= state.toolbar.handle_input(&event);
                }

                if state.app_state.show_sidebar {
                    let before = state.sidebar.active_panel;
                    handled |= state.sidebar.handle_input(&event);
                    handled |= apply_pending_sidebar_intents(state);
                    if before != state.sidebar.active_panel {
                        state.app_state.status_text = format!(
                            "Sidebar panel: {}",
                            match state.sidebar.active_panel {
                                SidebarPanel::Files => "Files",
                                SidebarPanel::Outline => "Outline",
                                SidebarPanel::Bookmarks => "Bookmarks",
                                SidebarPanel::SearchResults => "Search Results",
                            }
                        );
                    }
                }

                if state.app_state.show_statusbar {
                    handled |= state.statusbar.handle_input(&event);
                    if let Some(action) = state.statusbar.pending_action.take() {
                        state.app_state.status_text = match action {
                            StatusAction::OpenZoomPopup => "Zoom control requested".to_string(),
                            StatusAction::ChangeEncoding => "Encoding picker requested".to_string(),
                        };
                        handled = true;
                    }
                }

                if begin_image_interaction(state, point) {
                    handled = true;
                }
                if !handled && begin_table_interaction(state, point) {
                    if state.table_resize.is_some() {
                        let _ = unsafe { SetCapture(hwnd) };
                    }
                    handled = true;
                }

                if handled {
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }
            }
            unsafe { DefWindowProcW(hwnd, message, wparam, lparam) }
        }
        WM_MBUTTONDOWN => {
            if let Some(state) = unsafe { state_from_hwnd(hwnd) } {
                let point = point_from_lparam(lparam);
                if state.app_state.show_tabs
                    && let Some(index) = state.tabs.tab_hit_test(point)
                {
                    let _ = close_tab_with_prompt(state, hwnd, index);
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }
            }
            unsafe { DefWindowProcW(hwnd, message, wparam, lparam) }
        }
        WM_LBUTTONUP => {
            if let Some(state) = unsafe { state_from_hwnd(hwnd) } {
                let point = point_from_lparam(lparam);
                if state.app_state.show_tabs {
                    let tab_event = UiInputEvent::MouseUp(point);
                    if state.tabs.handle_input(&tab_event) {
                        let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                        return LRESULT(0);
                    }
                }
                if state.sidebar_resizing {
                    state.sidebar_resizing = false;
                    state.sidebar.resizing = false;
                    let _ = unsafe { ReleaseCapture() };
                    state.app_state.status_text = format!(
                        "Sidebar width set to {:.0}px",
                        state.app_state.sidebar_width
                    );
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }
                if state.table_resize.take().is_some() {
                    let _ = unsafe { ReleaseCapture() };
                    state.app_state.status_text = "Table resized".to_string();
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }
                if state.image_drag.take().is_some() {
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }
            }
            unsafe { DefWindowProcW(hwnd, message, wparam, lparam) }
        }
        WM_LBUTTONDBLCLK => {
            if let Some(state) = unsafe { state_from_hwnd(hwnd) } {
                let point = point_from_lparam(lparam);
                if state.app_state.show_tabs && state.tabs.is_empty_tab_bar_space(point) {
                    let index = open_new_blank_tab(state);
                    let title = state
                        .tabs
                        .tabs
                        .get(index)
                        .map(|tab| tab.title.clone())
                        .unwrap_or_else(|| "New tab".to_string());
                    state.app_state.status_text = format!("Opened {title}");
                    sync_sidebar_with_active_tab(state);
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }
                if begin_image_interaction(state, point) {
                    state.image_properties_visible = true;
                    if let Some(selected) = state.selected_image {
                        if let Some(image) = active_image_ref(state, selected) {
                            state.app_state.status_text = format!(
                                "Image properties: {:.0}x{:.0}, {:?}, alt='{}'",
                                image.width,
                                image.height,
                                image.alignment,
                                if image.alt_text.is_empty() {
                                    "(empty)"
                                } else {
                                    image.alt_text.as_str()
                                }
                            );
                        }
                    }
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }
            }
            unsafe { DefWindowProcW(hwnd, message, wparam, lparam) }
        }
        WM_DESTROY => {
            if let Some(state) = unsafe { state_from_hwnd(hwnd) } {
                state.settings_dialog.force_flush();
            }
            unsafe { PostQuitMessage(0) };
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut WindowState;
            if !ptr.is_null() {
                unsafe {
                    let _ = Box::from_raw(ptr);
                    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                }
            }
            unsafe { DefWindowProcW(hwnd, message, wparam, lparam) }
        }
        _ => unsafe { DefWindowProcW(hwnd, message, wparam, lparam) },
    }
}
