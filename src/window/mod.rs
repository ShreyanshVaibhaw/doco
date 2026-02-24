use std::{
    ffi::c_void,
    mem::size_of,
    path::PathBuf,
    time::Instant,
};

use windows::{
    Win32::{
        Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, RECT, WPARAM},
        Graphics::{
            Dwm::{
                DWMSBT_MAINWINDOW,
                DWMWA_SYSTEMBACKDROP_TYPE,
                DWMWA_USE_IMMERSIVE_DARK_MODE,
                DwmSetWindowAttribute,
            },
            Gdi::{BeginPaint, EndPaint, InvalidateRect, PAINTSTRUCT},
        },
        System::LibraryLoader::GetModuleHandleW,
        UI::{
            HiDpi::{DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2, SetProcessDpiAwarenessContext},
            Input::KeyboardAndMouse::{GetKeyState, ReleaseCapture, SetCapture, VK_CONTROL, VK_SHIFT},
            Shell::{DragAcceptFiles, HDROP},
            WindowsAndMessaging::{
                AdjustWindowRectEx, CREATESTRUCTW, CS_HREDRAW, CS_VREDRAW, CreateWindowExW,
                DefWindowProcW, DispatchMessageW, GWLP_USERDATA, GetClientRect, GetMessageW,
                GetSystemMetrics, GetWindowLongPtrW, IDC_ARROW, LoadCursorW, MSG, PostQuitMessage,
                RegisterClassExW, SM_CXSCREEN, SM_CYSCREEN, SW_SHOW, SWP_NOACTIVATE, SWP_NOZORDER,
                SetWindowLongPtrW, SetWindowPos, ShowWindow, TranslateMessage,
                WINDOW_EX_STYLE, WM_CREATE, WM_DESTROY, WM_DPICHANGED, WM_DROPFILES, WM_KEYDOWN,
                WM_LBUTTONDOWN, WM_LBUTTONUP, WM_MOUSEMOVE, WM_MOUSEWHEEL, WM_NCCREATE, WM_NCDESTROY, WM_PAINT,
                WM_SETTINGCHANGE, WM_SIZE,
                WNDCLASSEXW,
                WS_OVERLAPPEDWINDOW, WS_VISIBLE,
            },
        },
    },
    core::{Result, w},
};

use crate::{
    app::AppState,
    document::model::{Block, DocumentModel},
    render::canvas::PageLayoutMode,
    render::d2d::{D2DRenderer, ShellRenderState},
    settings::schema::Settings,
    theme::{Theme, ThemeManager, backgrounds::{BackgroundKind, from_canvas_preference}},
    ui::{
        InputEvent as UiInputEvent,
        Point as UiPoint,
        Rect as UiRect,
        UIComponent,
        sidebar::{Sidebar, SidebarPanel},
        statusbar::{StatusAction, StatusBar, StatusBarInfo},
        tabs::TabsBar,
        toolbar::{Toolbar, ToolbarAction},
    },
    window::integration::{
        DropAction,
        JumpListState,
        PrintState,
        extract_drop_payload,
        parse_startup_files_from_cli,
    },
};

pub mod compositor;
pub mod input;
pub mod integration;

pub struct AppWindow {
    hwnd: HWND,
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
            style: CS_HREDRAW | CS_VREDRAW,
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
            sidebar: Sidebar::default(),
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

fn canvas_viewport_size(state: &WindowState, width: f32, height: f32) -> (f32, f32) {
    let tab_h = if state.app_state.show_tabs { 36.0 } else { 0.0 };
    let status_h = if state.app_state.show_statusbar { 28.0 } else { 0.0 };
    let sidebar_w = if state.app_state.show_sidebar {
        state.app_state.sidebar_width.clamp(200.0, 400.0)
    } else {
        0.0
    };
    let toolbar_h = if state.app_state.show_toolbar { 44.0 } else { 0.0 };
    (
        (width - sidebar_w).max(1.0),
        (height - tab_h - toolbar_h - status_h).max(1.0),
    )
}

fn relayout_shell(state: &mut WindowState, width: f32, height: f32) {
    let tab_h = if state.app_state.show_tabs { 36.0 } else { 0.0 };
    let status_h = if state.app_state.show_statusbar { 28.0 } else { 0.0 };
    let sidebar_w = if state.app_state.show_sidebar {
        state.app_state.sidebar_width.clamp(200.0, 400.0)
    } else {
        0.0
    };
    let toolbar_h = if state.app_state.show_toolbar { 44.0 } else { 0.0 };

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
                let line = if c.code.is_empty() { "code block" } else { &c.code };
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

    if let Some(tab) = state.tabs.active_tab_mut() {
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
        canvas_preview_lines = collect_preview_lines(&tab.document, 40);
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

    ShellRenderState {
        show_tabs: state.app_state.show_tabs,
        show_sidebar: state.app_state.show_sidebar,
        sidebar_width: state.app_state.sidebar_width,
        sidebar_resizing: state.sidebar_resizing,
        show_toolbar: state.app_state.show_toolbar,
        show_statusbar: state.app_state.show_statusbar,
        status_text: state.app_state.status_text.clone(),
        tab_titles: state.tabs.tabs.iter().map(|t| t.title.clone()).collect(),
        active_tab: state.tabs.active,
        toolbar_labels: state
            .toolbar
            .buttons
            .iter()
            .filter(|b| !b.label.is_empty())
            .map(|b| b.label.to_string())
            .collect(),
        active_sidebar_panel: active_sidebar_panel.to_string(),
        status_left: state.statusbar.left_text(),
        status_right: state.statusbar.right_text(),
        canvas_background: from_canvas_preference(&state.app_state.settings.appearance.canvas_background),
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

                if !state.startup_files.is_empty() {
                    state.dropped_files = state.startup_files.clone();
                    state.app_state.status_text =
                        format!("Opening {} file(s) from command line", state.startup_files.len());
                    for path in &state.startup_files {
                        state.jump_list.add_recent_file(path.clone());
                        let title = path
                            .file_name()
                            .and_then(|v| v.to_str())
                            .unwrap_or("Document")
                            .to_string();
                        state
                            .tabs
                            .open_document_tab(title, Some(path.clone()), DocumentModel::default());
                    }
                } else {
                    let _ = state.tabs.new_blank_tab();
                }
            }

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
                let mut needs_next_frame = false;
                if let Some(tab) = state.tabs.active_tab_mut() {
                    needs_next_frame |= tab.canvas.update(dt);
                    tab.canvas.clamp_scroll(&tab.document);
                }
                let background = from_canvas_preference(&state.app_state.settings.appearance.canvas_background);
                if matches!(background.kind, BackgroundKind::AnimatedGradient { .. }) {
                    needs_next_frame = true;
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

                if let Some(tab) = state.tabs.active_tab_mut() {
                    tab.canvas.set_viewport(canvas_w, canvas_h);
                    if ctrl_down {
                        tab.canvas.handle_mouse_wheel(delta, true, cursor_in_canvas);
                        state.app_state.status_text = format!("Zoom: {}%", (tab.canvas.zoom_target * 100.0).round() as u16);
                    } else if shift_down {
                        tab.canvas.handle_horizontal_wheel(delta);
                        state.app_state.status_text = "Horizontal scroll".to_string();
                    } else {
                        tab.canvas.handle_mouse_wheel(delta, false, cursor_in_canvas);
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

                if ctrl_down && !shift_down && vk == 0x50 {
                    state.print_state.request_print_dialog();
                    state.app_state.status_text = "Print dialog requested".to_string();
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if ctrl_down && !shift_down && vk == 0xBC {
                    state.app_state.show_settings = !state.app_state.show_settings;
                    state.app_state.status_text = if state.app_state.show_settings {
                        "Settings toggled on".to_string()
                    } else {
                        "Settings toggled off".to_string()
                    };
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if ctrl_down && !shift_down && vk == 0x42 {
                    state.app_state.show_sidebar = !state.app_state.show_sidebar;
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

                if ctrl_down && !shift_down && vk == 0x54 {
                    state.app_state.show_toolbar = !state.app_state.show_toolbar;
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

                if ctrl_down && !shift_down && vk == 0x4E {
                    let index = state.tabs.new_blank_tab();
                    let title = state
                        .tabs
                        .tabs
                        .get(index)
                        .map(|t| t.title.clone())
                        .unwrap_or_else(|| "New tab".to_string());
                    state.app_state.status_text = format!("Opened {title}");
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if ctrl_down && !shift_down && vk == 0x57 {
                    if state.tabs.close_active_tab() {
                        let active_title = state
                            .tabs
                            .active_tab()
                            .map(|t| t.title.clone())
                            .unwrap_or_else(|| "Welcome".to_string());
                        state.app_state.status_text = format!("Active tab: {active_title}");
                    } else {
                        state.app_state.status_text = "No tab to close".to_string();
                    }
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if ctrl_down && vk == 0x09 {
                    if shift_down {
                        state.tabs.switch_prev();
                    } else {
                        state.tabs.switch_next();
                    }
                    let active_title = state
                        .tabs
                        .active_tab()
                        .map(|t| t.title.clone())
                        .unwrap_or_else(|| "Welcome".to_string());
                    state.app_state.status_text = format!("Switched to {active_title}");
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
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
                            state
                                .tabs
                                .open_document_tab(title, Some(path.clone()), DocumentModel::default());
                        }
                        format!("Drop to open: {} file(s)", payload.files.len())
                    }
                    DropAction::InsertImage => {
                        format!("Drop to insert image: {} file(s)", payload.files.len())
                    }
                    DropAction::Ignore => "Unsupported dropped content".to_string(),
                };

                let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
            }
            LRESULT(0)
        }
        WM_MOUSEMOVE => {
            if let Some(state) = unsafe { state_from_hwnd(hwnd) } {
                let point = point_from_lparam(lparam);
                if state.sidebar_resizing {
                    let next_width = (point.x + state.sidebar_resize_grab_offset).clamp(200.0, 400.0);
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
                        if let Some(tab) = state.tabs.active_tab() {
                            state.app_state.status_text = format!("Switched to {}", tab.title);
                        }
                    }
                }

                if state.app_state.show_toolbar {
                    if let Some(index) = state.toolbar.hit_button(point) {
                        if let Some(action) = state.toolbar.invoke(index) {
                            state.app_state.status_text =
                                format!("Toolbar action: {}", toolbar_action_text(action));
                        }
                    }
                    handled |= state.toolbar.handle_input(&event);
                }

                if state.app_state.show_sidebar {
                    let before = state.sidebar.active_panel;
                    handled |= state.sidebar.handle_input(&event);
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

                if handled {
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }
            }
            unsafe { DefWindowProcW(hwnd, message, wparam, lparam) }
        }
        WM_LBUTTONUP => {
            if let Some(state) = unsafe { state_from_hwnd(hwnd) } {
                if state.sidebar_resizing {
                    state.sidebar_resizing = false;
                    state.sidebar.resizing = false;
                    let _ = unsafe { ReleaseCapture() };
                    state.app_state.status_text =
                        format!("Sidebar width set to {:.0}px", state.app_state.sidebar_width);
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }
            }
            unsafe { DefWindowProcW(hwnd, message, wparam, lparam) }
        }
        WM_DESTROY => {
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

