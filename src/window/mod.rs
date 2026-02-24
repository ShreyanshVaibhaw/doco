use std::{
    ffi::c_void,
    mem::size_of,
    path::PathBuf,
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
            Input::KeyboardAndMouse::{GetKeyState, VK_CONTROL, VK_SHIFT},
            Shell::{DragAcceptFiles, HDROP},
            WindowsAndMessaging::{
                AdjustWindowRectEx, CREATESTRUCTW, CS_HREDRAW, CS_VREDRAW, CreateWindowExW,
                DefWindowProcW, DispatchMessageW, GWLP_USERDATA, GetClientRect, GetMessageW,
                GetSystemMetrics, GetWindowLongPtrW, IDC_ARROW, LoadCursorW, MSG, PostQuitMessage,
                RegisterClassExW, SM_CXSCREEN, SM_CYSCREEN, SW_SHOW, SWP_NOACTIVATE, SWP_NOZORDER,
                SetWindowLongPtrW, SetWindowPos, ShowWindow, TranslateMessage, WINDOW_EX_STYLE,
                WM_CREATE, WM_DESTROY, WM_DPICHANGED, WM_DROPFILES, WM_KEYDOWN, WM_NCCREATE,
                WM_NCDESTROY, WM_PAINT, WM_SETTINGCHANGE, WM_SIZE, WNDCLASSEXW,
                WS_OVERLAPPEDWINDOW, WS_VISIBLE,
            },
        },
    },
    core::{Result, w},
};

use crate::{
    render::d2d::D2DRenderer,
    theme::Theme,
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
    debug_panel_visible: bool,
    dropped_files: Vec<PathBuf>,
    drop_overlay_text: Option<String>,
    jump_list: JumpListState,
    print_state: PrintState,
    startup_files: Vec<PathBuf>,
}

impl AppWindow {
    pub fn new(theme: Theme) -> Result<Self> {
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

        let state = Box::new(WindowState {
            renderer: None,
            dpi: 96.0,
            theme,
            debug_panel_visible: false,
            dropped_files: Vec::new(),
            drop_overlay_text: None,
            jump_list: JumpListState::with_default_tasks(),
            print_state: PrintState::default(),
            startup_files: parse_startup_files_from_cli(),
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

unsafe fn apply_window_effects(hwnd: HWND) {
    let dark_mode = windows::core::BOOL(1);
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
            unsafe { apply_window_effects(hwnd) };
            unsafe { DragAcceptFiles(hwnd, true) };

            let mut client = RECT::default();
            let _ = unsafe { GetClientRect(hwnd, &mut client) };

            if let Some(state) = unsafe { state_from_hwnd(hwnd) } {
                let width = (client.right - client.left).max(1) as u32;
                let height = (client.bottom - client.top).max(1) as u32;

                match D2DRenderer::new(hwnd, width, height, state.dpi, state.theme.clone()) {
                    Ok(renderer) => state.renderer = Some(renderer),
                    Err(error) => {
                        eprintln!("Renderer initialization failed: {error:?}");
                    }
                }

                if !state.startup_files.is_empty() {
                    state.dropped_files = state.startup_files.clone();
                    state.drop_overlay_text =
                        Some(format!("Opening {} file(s) from command line", state.startup_files.len()));
                    for path in &state.startup_files {
                        state.jump_list.add_recent_file(path.clone());
                    }
                }
            }

            LRESULT(0)
        }
        WM_SIZE => {
            if let Some(state) = unsafe { state_from_hwnd(hwnd) } {
                if let Some(renderer) = &mut state.renderer {
                    let width = (lparam.0 as u32 & 0xFFFF) as u32;
                    let height = ((lparam.0 as u32 >> 16) & 0xFFFF) as u32;
                    let _ = renderer.resize(width, height);
                }
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
            }

            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
            LRESULT(0)
        }
        WM_SETTINGCHANGE => {
            let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
            LRESULT(0)
        }
        WM_PAINT => {
            let mut paint = PAINTSTRUCT::default();
            let _ = unsafe { BeginPaint(hwnd, &mut paint) };

            if let Some(state) = unsafe { state_from_hwnd(hwnd) } {
                if let Some(renderer) = &mut state.renderer {
                    let _ = renderer.render();
                }
            }

            let _ = unsafe { EndPaint(hwnd, &paint) };
            LRESULT(0)
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
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if ctrl_down && !shift_down && vk == 0x50 {
                    state.print_state.request_print_dialog();
                    state.drop_overlay_text = Some("Print dialog requested".to_string());
                    let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
                    return LRESULT(0);
                }

                if ctrl_down && !shift_down && vk == 0xBC {
                    state.drop_overlay_text = Some("Settings requested".to_string());
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

                state.drop_overlay_text = Some(match payload.action {
                    DropAction::OpenFilesInTabs => {
                        for path in &payload.files {
                            state.jump_list.add_recent_file(path.clone());
                        }
                        format!("Drop to open: {} file(s)", payload.files.len())
                    }
                    DropAction::InsertImage => {
                        format!("Drop to insert image: {} file(s)", payload.files.len())
                    }
                    DropAction::Ignore => "Unsupported dropped content".to_string(),
                });

                let _ = unsafe { InvalidateRect(Some(hwnd), None, false) };
            }
            LRESULT(0)
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

