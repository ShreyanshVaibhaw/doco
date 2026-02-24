use std::{mem::ManuallyDrop, path::Path};
use std::time::{Instant, SystemTime, UNIX_EPOCH};

use windows::{
    Win32::{
        Foundation::{HMODULE, HWND, RECT},
        Graphics::{
            Direct2D::{
                Common::{D2D_RECT_F, D2D1_ALPHA_MODE_IGNORE, D2D1_PIXEL_FORMAT},
                D2D1_BITMAP_OPTIONS_CANNOT_DRAW, D2D1_BITMAP_OPTIONS_TARGET, D2D1_BITMAP_PROPERTIES1,
                D2D1_DEVICE_CONTEXT_OPTIONS_NONE, D2D1_DRAW_TEXT_OPTIONS_NONE,
                D2D1_FACTORY_TYPE_SINGLE_THREADED, D2D1CreateFactory, ID2D1Bitmap1, ID2D1Device,
                ID2D1DeviceContext, ID2D1Factory1, ID2D1Image, ID2D1SolidColorBrush,
            },
            Direct3D::{D3D_DRIVER_TYPE_HARDWARE, D3D_DRIVER_TYPE_WARP},
            Direct3D11::{
                D3D11_CREATE_DEVICE_BGRA_SUPPORT, D3D11_SDK_VERSION, D3D11CreateDevice, ID3D11Device,
                ID3D11DeviceContext,
            },
            DirectWrite::{
                DWRITE_FACTORY_TYPE_SHARED, DWRITE_MEASURING_MODE_NATURAL, DWriteCreateFactory,
                IDWriteFactory, IDWriteTextFormat,
            },
            Dxgi::{
                Common::{
                    DXGI_ALPHA_MODE_IGNORE, DXGI_ALPHA_MODE_UNSPECIFIED, DXGI_FORMAT_B8G8R8A8_UNORM,
                    DXGI_FORMAT_UNKNOWN, DXGI_SAMPLE_DESC,
                },
                DXGI_PRESENT, DXGI_SCALING_STRETCH, DXGI_SWAP_CHAIN_DESC1, DXGI_SWAP_CHAIN_FLAG,
                DXGI_SWAP_EFFECT_DISCARD, DXGI_SWAP_EFFECT_FLIP_DISCARD, DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL,
                DXGI_USAGE_RENDER_TARGET_OUTPUT, IDXGIDevice, IDXGIFactory2, IDXGISurface,
                IDXGISwapChain1,
            },
        },
        UI::WindowsAndMessaging::GetClientRect,
    },
    core::{Result, HRESULT, Interface, w},
};
use windows_numerics::Vector2;

use crate::{
    render::perf::{DebugPerformancePanel, query_process_working_set_bytes},
    theme::{
        Theme,
        backgrounds::{BackgroundKind, BackgroundSettings, PatternStyle, preset_by_id},
    },
    ui::Rect as UiRect,
};

const D2DERR_RECREATE_TARGET: HRESULT = HRESULT(0x8899000C_u32 as i32);

#[derive(Debug, Clone, Default)]
pub struct ShellRenderState {
    pub show_tabs: bool,
    pub show_sidebar: bool,
    pub sidebar_width: f32,
    pub sidebar_resizing: bool,
    pub show_toolbar: bool,
    pub show_statusbar: bool,
    pub status_text: String,
    pub tab_titles: Vec<String>,
    pub active_tab: usize,
    pub toolbar_labels: Vec<String>,
    pub active_sidebar_panel: String,
    pub status_left: String,
    pub status_right: String,
    pub canvas_background: BackgroundSettings,
    pub canvas_page_rects: Vec<UiRect>,
    pub canvas_preview_lines: Vec<String>,
    pub canvas_show_margin_guides: bool,
    pub canvas_cursor_visible: bool,
    pub canvas_scrollbar_visible: bool,
    pub canvas_scrollbar_alpha: f32,
    pub canvas_viewport_width: f32,
    pub canvas_viewport_height: f32,
    pub canvas_content_width: f32,
    pub canvas_content_height: f32,
    pub canvas_scroll_x: f32,
    pub canvas_scroll_y: f32,
}

pub struct D2DRenderer {
    hwnd: HWND,
    dpi: f32,
    #[allow(dead_code)]
    d3d_device: ID3D11Device,
    #[allow(dead_code)]
    d3d_context: ID3D11DeviceContext,
    #[allow(dead_code)]
    d2d_factory: ID2D1Factory1,
    #[allow(dead_code)]
    d2d_device: ID2D1Device,
    d2d_context: ID2D1DeviceContext,
    swap_chain: IDXGISwapChain1,
    target_bitmap: Option<ID2D1Bitmap1>,
    dwrite_factory: IDWriteFactory,
    theme: Theme,
    debug_panel: DebugPerformancePanel,
}

impl D2DRenderer {
    pub fn new(hwnd: HWND, width: u32, height: u32, dpi: f32, theme: Theme) -> Result<Self> {
        unsafe {
            let d2d_factory: ID2D1Factory1 = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, None)?;

            let (d3d_device, d3d_context) = Self::create_d3d_device()?;

            let dxgi_device: IDXGIDevice = d3d_device.cast()?;
            let adapter = dxgi_device.GetAdapter()?;
            let dxgi_factory: IDXGIFactory2 = adapter.GetParent()?;

            let swap_chain = Self::create_swap_chain_for_hwnd(
                &dxgi_factory,
                &d3d_device,
                hwnd,
                width,
                height,
            )?;

            let d2d_device = d2d_factory.CreateDevice(&dxgi_device)?;
            let d2d_context = d2d_device.CreateDeviceContext(D2D1_DEVICE_CONTEXT_OPTIONS_NONE)?;

            let dwrite_factory: IDWriteFactory = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED)?;

            let mut renderer = Self {
                hwnd,
                dpi,
                d3d_device,
                d3d_context,
                d2d_factory,
                d2d_device,
                d2d_context,
                swap_chain,
                target_bitmap: None,
                dwrite_factory,
                theme,
                debug_panel: DebugPerformancePanel::default(),
            };

            renderer.recreate_target_bitmap()?;
            Ok(renderer)
        }
    }

    fn create_d3d_device() -> Result<(ID3D11Device, ID3D11DeviceContext)> {
        let hardware_result = Self::create_d3d_device_for_driver(D3D_DRIVER_TYPE_HARDWARE);
        if let Ok(devices) = hardware_result {
            return Ok(devices);
        }

        if let Err(error) = hardware_result {
            eprintln!(
                "Hardware D3D11 initialization failed, falling back to WARP software renderer: {error:?}"
            );
        }
        Self::create_d3d_device_for_driver(D3D_DRIVER_TYPE_WARP)
    }

    fn create_d3d_device_for_driver(
        driver_type: windows::Win32::Graphics::Direct3D::D3D_DRIVER_TYPE,
    ) -> Result<(ID3D11Device, ID3D11DeviceContext)> {
        unsafe {
            let mut d3d_device = None;
            let mut d3d_context = None;
            D3D11CreateDevice(
                None,
                driver_type,
                HMODULE::default(),
                D3D11_CREATE_DEVICE_BGRA_SUPPORT,
                None,
                D3D11_SDK_VERSION,
                Some(&mut d3d_device),
                None,
                Some(&mut d3d_context),
            )?;
            let d3d_device = d3d_device.expect("D3D11 device should be created");
            let d3d_context = d3d_context.expect("D3D11 context should be created");
            Ok((d3d_device, d3d_context))
        }
    }

    fn create_swap_chain_for_hwnd(
        dxgi_factory: &IDXGIFactory2,
        d3d_device: &ID3D11Device,
        hwnd: HWND,
        width: u32,
        height: u32,
    ) -> Result<IDXGISwapChain1> {
        let attempts = [
            (DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL, 2, DXGI_ALPHA_MODE_IGNORE, "flip_sequential"),
            (DXGI_SWAP_EFFECT_FLIP_DISCARD, 2, DXGI_ALPHA_MODE_IGNORE, "flip_discard"),
            (DXGI_SWAP_EFFECT_DISCARD, 1, DXGI_ALPHA_MODE_UNSPECIFIED, "discard"),
        ];

        let mut last_error = None;
        for (swap_effect, buffer_count, alpha_mode, label) in attempts {
            let swap_chain_desc = DXGI_SWAP_CHAIN_DESC1 {
                Width: width,
                Height: height,
                Format: DXGI_FORMAT_B8G8R8A8_UNORM,
                Stereo: false.into(),
                SampleDesc: DXGI_SAMPLE_DESC { Count: 1, Quality: 0 },
                BufferUsage: DXGI_USAGE_RENDER_TARGET_OUTPUT,
                BufferCount: buffer_count,
                Scaling: DXGI_SCALING_STRETCH,
                SwapEffect: swap_effect,
                AlphaMode: alpha_mode,
                Flags: 0,
            };

            let result = unsafe {
                dxgi_factory.CreateSwapChainForHwnd(
                    d3d_device,
                    hwnd,
                    &swap_chain_desc,
                    None,
                    None,
                )
            };

            match result {
                Ok(swap_chain) => return Ok(swap_chain),
                Err(error) => {
                    eprintln!("Swap chain creation attempt '{label}' failed: {error:?}");
                    last_error = Some(error);
                }
            }
        }

        Err(last_error.expect("At least one swap chain creation attempt must run"))
    }

    pub fn set_dpi(&mut self, dpi: f32) {
        self.dpi = dpi;
        unsafe {
            let _ = self.d2d_context.SetDpi(dpi, dpi);
        }
    }

    pub fn resize(&mut self, width: u32, height: u32) -> Result<()> {
        if width == 0 || height == 0 {
            return Ok(());
        }

        unsafe {
            self.d2d_context.SetTarget(None::<&ID2D1Image>);
            self.target_bitmap = None;
            self.swap_chain
                .ResizeBuffers(0, width, height, DXGI_FORMAT_UNKNOWN, DXGI_SWAP_CHAIN_FLAG(0))?;
            self.recreate_target_bitmap()?;
        }

        Ok(())
    }

    pub fn render(&mut self, shell: &ShellRenderState) -> Result<()> {
        crate::profile_scope!("renderer.frame");
        let frame_start = Instant::now();

        unsafe {
            self.d2d_context.BeginDraw();
            let clear = self.theme.window_bg.as_d2d();
            self.d2d_context.Clear(Some(&clear));

            self.draw_shell_placeholder(shell)?;

            match self.d2d_context.EndDraw(None, None) {
                Ok(()) => {
                    self.swap_chain.Present(1, DXGI_PRESENT(0)).ok()?;
                }
                Err(error) if error.code() == D2DERR_RECREATE_TARGET => {
                    self.recreate_target_bitmap()?;
                }
                Err(error) => return Err(error),
            }
        }

        let frame_ms = frame_start.elapsed().as_secs_f32() * 1000.0;
        self.debug_panel.update_frame_time(frame_ms);
        if let Some(bytes) = query_process_working_set_bytes() {
            self.debug_panel.update_memory_bytes(bytes);
        }

        Ok(())
    }

    pub fn set_debug_panel_visible(&mut self, visible: bool) {
        self.debug_panel.set_visible(visible);
    }

    pub fn set_theme(&mut self, theme: Theme) {
        self.theme = theme;
    }

    pub fn debug_panel(&self) -> &DebugPerformancePanel {
        &self.debug_panel
    }

    fn draw_shell_placeholder(&self, shell: &ShellRenderState) -> Result<()> {
        unsafe {
            let mut rect = RECT::default();
            GetClientRect(self.hwnd, &mut rect)?;

            let width = (rect.right - rect.left) as f32;
            let height = (rect.bottom - rect.top) as f32;

            let tab_h = if shell.show_tabs { 36.0 } else { 0.0 };
            let sidebar_w = if shell.show_sidebar {
                shell.sidebar_width.clamp(200.0, 400.0).min((width - 80.0).max(0.0))
            } else {
                0.0
            };
            let toolbar_h = if shell.show_toolbar { 44.0 } else { 0.0 };
            let status_h = if shell.show_statusbar { 28.0 } else { 0.0 };

            let tab_rect = D2D_RECT_F { left: 0.0, top: 0.0, right: width, bottom: tab_h };
            let sidebar_rect = D2D_RECT_F {
                left: 0.0,
                top: tab_h,
                right: sidebar_w,
                bottom: height - status_h,
            };
            let toolbar_rect = D2D_RECT_F {
                left: sidebar_w,
                top: tab_h,
                right: width,
                bottom: tab_h + toolbar_h,
            };
            let canvas_rect = D2D_RECT_F {
                left: sidebar_w,
                top: tab_h + toolbar_h,
                right: width,
                bottom: height - status_h,
            };
            let status_rect = D2D_RECT_F {
                left: 0.0,
                top: height - status_h,
                right: width,
                bottom: height,
            };

            let tab_brush = self.create_brush(self.theme.surface_secondary.as_d2d())?;
            let side_brush = self.create_brush(self.theme.surface_primary.as_d2d())?;
            let tool_brush = self.create_brush(self.theme.surface_secondary.as_d2d())?;
            let status_brush = self.create_brush(self.theme.surface_primary.as_d2d())?;

            if tab_h > 0.0 {
                self.d2d_context.FillRectangle(&tab_rect, &tab_brush);
            }
            if sidebar_w > 0.0 {
                self.d2d_context.FillRectangle(&sidebar_rect, &side_brush);
            }
            if toolbar_h > 0.0 {
                self.d2d_context.FillRectangle(&toolbar_rect, &tool_brush);
            }
            self.draw_canvas_background(canvas_rect, &shell.canvas_background)?;
            self.draw_document_canvas(canvas_rect, shell)?;
            if status_h > 0.0 {
                self.d2d_context.FillRectangle(&status_rect, &status_brush);
            }

            if sidebar_w > 0.0 {
                let border_brush = self.create_brush(self.theme.border_subtle.as_d2d())?;
                self.d2d_context.DrawLine(
                    Vector2 {
                        X: sidebar_w,
                        Y: tab_h,
                    },
                    Vector2 {
                        X: sidebar_w,
                        Y: height - status_h,
                    },
                    &border_brush,
                    1.0,
                    None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                );

                let splitter_color = if shell.sidebar_resizing {
                    self.theme.accent
                } else {
                    self.theme.border_default
                };
                let splitter_brush = self.create_brush(splitter_color.as_d2d())?;
                let splitter_rect = D2D_RECT_F {
                    left: (sidebar_w - 1.5).max(0.0),
                    top: tab_h + 4.0,
                    right: sidebar_w + 1.5,
                    bottom: (height - status_h - 4.0).max(tab_h + 6.0),
                };
                self.d2d_context.FillRectangle(&splitter_rect, &splitter_brush);
            }

            let text_brush = self.create_brush(self.theme.text_primary.as_d2d())?;
            let text_format = self.create_text_format()?;
            if sidebar_w > 0.0 {
                let files = "Files".encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &files,
                    &text_format,
                    &D2D_RECT_F {
                        left: 14.0,
                        top: tab_h + 8.0,
                        right: sidebar_w - 8.0,
                        bottom: tab_h + 32.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                let active_panel = shell.active_sidebar_panel.encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &active_panel,
                    &text_format,
                    &D2D_RECT_F {
                        left: 14.0,
                        top: tab_h + 34.0,
                        right: sidebar_w - 8.0,
                        bottom: tab_h + 58.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );
            }

            let status = shell.status_text.encode_utf16().collect::<Vec<u16>>();
            if status_h > 0.0 {
                self.d2d_context.DrawText(
                    &status,
                    &text_format,
                    &D2D_RECT_F {
                        left: 14.0,
                        top: height - status_h + 4.0,
                        right: width - 8.0,
                        bottom: height - 2.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                let status_right = shell.status_right.encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &status_right,
                    &text_format,
                    &D2D_RECT_F {
                        left: width - 420.0,
                        top: height - status_h + 4.0,
                        right: width - 10.0,
                        bottom: height - 2.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                let status_left = shell.status_left.encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &status_left,
                    &text_format,
                    &D2D_RECT_F {
                        left: 14.0,
                        top: height - status_h + 4.0,
                        right: (width - 440.0).max(100.0),
                        bottom: height - 2.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );
            }

            if tab_h > 0.0 && !shell.tab_titles.is_empty() {
                let mut x = 8.0;
                for (idx, title) in shell.tab_titles.iter().take(8).enumerate() {
                    let tab_w = ((width - 16.0) / shell.tab_titles.len().min(8) as f32).clamp(120.0, 220.0);
                    let rect = D2D_RECT_F {
                        left: x,
                        top: 4.0,
                        right: (x + tab_w).min(width - 8.0),
                        bottom: tab_h - 4.0,
                    };
                    let brush = if idx == shell.active_tab {
                        self.create_brush(self.theme.surface_hover.as_d2d())?
                    } else {
                        self.create_brush(self.theme.surface_secondary.as_d2d())?
                    };
                    self.d2d_context.FillRectangle(&rect, &brush);
                    let text = title.encode_utf16().collect::<Vec<u16>>();
                    self.d2d_context.DrawText(
                        &text,
                        &text_format,
                        &D2D_RECT_F {
                            left: rect.left + 10.0,
                            top: rect.top + 6.0,
                            right: rect.right - 6.0,
                            bottom: rect.bottom - 4.0,
                        },
                        &text_brush,
                        D2D1_DRAW_TEXT_OPTIONS_NONE,
                        DWRITE_MEASURING_MODE_NATURAL,
                    );
                    x += tab_w + 6.0;
                    if x > width - 80.0 {
                        break;
                    }
                }
            }

            if toolbar_h > 0.0 && !shell.toolbar_labels.is_empty() {
                let mut x = sidebar_w + 12.0;
                for label in shell.toolbar_labels.iter().take(12) {
                    let t = label.encode_utf16().collect::<Vec<u16>>();
                    self.d2d_context.DrawText(
                        &t,
                        &text_format,
                        &D2D_RECT_F {
                            left: x,
                            top: tab_h + 11.0,
                            right: x + 80.0,
                            bottom: tab_h + toolbar_h - 6.0,
                        },
                        &text_brush,
                        D2D1_DRAW_TEXT_OPTIONS_NONE,
                        DWRITE_MEASURING_MODE_NATURAL,
                    );
                    x += 74.0;
                    if x > width - 70.0 {
                        break;
                    }
                }
            }

            if self.debug_panel.visible {
                let panel_rect = D2D_RECT_F {
                    left: width - 290.0,
                    top: tab_h + 12.0,
                    right: width - 12.0,
                    bottom: tab_h + 146.0,
                };
                let panel_bg = self.create_brush(self.theme.surface_primary.as_d2d())?;
                let panel_border = self.create_brush(self.theme.border_default.as_d2d())?;
                self.d2d_context.FillRectangle(&panel_rect, &panel_bg);
                self.d2d_context.DrawRectangle(
                    &panel_rect,
                    &panel_border,
                    1.0,
                    None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                );

                let info = format!(
                    "Debug\nFPS: {:.1}\nFrame: {:.2} ms\nMemory: {:.1} MB\nCache Hit: {:.0}%\nCache: {:.1} MB",
                    self.debug_panel.snapshot.fps,
                    self.debug_panel.snapshot.frame_time_ms,
                    self.debug_panel.snapshot.process_memory_mb,
                    self.debug_panel.snapshot.image_cache_hit_rate * 100.0,
                    self.debug_panel.snapshot.image_cache_mb,
                );
                let text = info.encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &text,
                    &text_format,
                    &D2D_RECT_F {
                        left: panel_rect.left + 10.0,
                        top: panel_rect.top + 8.0,
                        right: panel_rect.right - 8.0,
                        bottom: panel_rect.bottom - 8.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );
            }
        }

        Ok(())
    }

    fn draw_canvas_background(&self, rect: D2D_RECT_F, settings: &BackgroundSettings) -> Result<()> {
        match &settings.kind {
            BackgroundKind::Solid { color } => self.fill_rect(rect, *color),
            BackgroundKind::Gradient { start, end, angle_degrees } => {
                self.fill_gradient(rect, *start, *end, *angle_degrees)
            }
            BackgroundKind::Pattern {
                style,
                foreground,
                background,
                scale,
            } => self.fill_pattern(rect, style.clone(), *foreground, *background, *scale),
            BackgroundKind::Image {
                path,
                mode,
                blur_px: _,
                opacity,
            } => {
                self.fill_rect(rect, self.theme.canvas_bg)?;
                let overlay = self.create_brush(
                    windows::Win32::Graphics::Direct2D::Common::D2D1_COLOR_F {
                        r: 0.0,
                        g: 0.0,
                        b: 0.0,
                        a: (1.0 - opacity.clamp(0.0, 1.0)).clamp(0.0, 1.0) * 0.35,
                    },
                )?;
                unsafe {
                    self.d2d_context.FillRectangle(&rect, &overlay);
                }

                let label = Path::new(path)
                    .file_name()
                    .and_then(|v| v.to_str())
                    .unwrap_or("custom image");
                let text = format!("Image background ({mode:?}): {label}");
                self.draw_canvas_label(rect, &text)
            }
            BackgroundKind::AnimatedGradient { colors, speed } => {
                if colors.len() < 2 {
                    return self.fill_rect(rect, self.theme.canvas_bg);
                }
                let now_s = SystemTime::now()
                    .duration_since(UNIX_EPOCH)
                    .map(|d| d.as_secs_f32())
                    .unwrap_or(0.0);
                let cycle = (now_s * speed.max(0.05)) % colors.len() as f32;
                let idx = cycle.floor() as usize;
                let next = (idx + 1) % colors.len();
                let t = cycle - idx as f32;
                let start = Self::lerp_color(colors[idx], colors[next], t);
                let end = Self::lerp_color(colors[next], colors[(next + 1) % colors.len()], t);
                self.fill_gradient(rect, start, end, 18.0)
            }
            BackgroundKind::Preset(id) => {
                let resolved = preset_by_id(id);
                self.draw_canvas_background(rect, &resolved)
            }
        }
    }

    fn draw_document_canvas(&self, canvas_rect: D2D_RECT_F, shell: &ShellRenderState) -> Result<()> {
        let shadow_color = crate::ui::Color::rgba(
            self.theme.page_shadow.r,
            self.theme.page_shadow.g,
            self.theme.page_shadow.b,
            if self.theme.is_dark { 0.32 } else { 0.22 },
        );
        let shadow_brush = self.create_brush(shadow_color.as_d2d())?;
        let page_brush = self.create_brush(self.theme.page_bg.as_d2d())?;
        let border_brush = self.create_brush(self.theme.border_subtle.as_d2d())?;
        let guide_brush = self.create_brush(
            crate::ui::Color::rgba(
                self.theme.border_default.r,
                self.theme.border_default.g,
                self.theme.border_default.b,
                0.28,
            )
            .as_d2d(),
        )?;

        unsafe {
            let mut drew_preview = false;
            for page in &shell.canvas_page_rects {
                let page_rect = D2D_RECT_F {
                    left: canvas_rect.left + page.x,
                    top: canvas_rect.top + page.y,
                    right: canvas_rect.left + page.x + page.width,
                    bottom: canvas_rect.top + page.y + page.height,
                };

                if page_rect.bottom < canvas_rect.top
                    || page_rect.top > canvas_rect.bottom
                    || page_rect.right < canvas_rect.left
                    || page_rect.left > canvas_rect.right
                {
                    continue;
                }

                let shadow_rect = D2D_RECT_F {
                    left: page_rect.left + 4.0,
                    top: page_rect.top + 4.0,
                    right: page_rect.right + 4.0,
                    bottom: page_rect.bottom + 4.0,
                };
                self.d2d_context.FillRectangle(&shadow_rect, &shadow_brush);
                self.d2d_context.FillRectangle(&page_rect, &page_brush);
                self.d2d_context.DrawRectangle(
                    &page_rect,
                    &border_brush,
                    1.0,
                    None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                );

                if shell.canvas_show_margin_guides {
                    let left_margin_x = page_rect.left + page.width * 0.11;
                    let right_margin_x = page_rect.right - page.width * 0.11;
                    self.d2d_context.DrawLine(
                        Vector2 {
                            X: left_margin_x,
                            Y: page_rect.top + 18.0,
                        },
                        Vector2 {
                            X: left_margin_x,
                            Y: page_rect.bottom - 18.0,
                        },
                        &guide_brush,
                        1.0,
                        None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                    );
                    self.d2d_context.DrawLine(
                        Vector2 {
                            X: right_margin_x,
                            Y: page_rect.top + 18.0,
                        },
                        Vector2 {
                            X: right_margin_x,
                            Y: page_rect.bottom - 18.0,
                        },
                        &guide_brush,
                        1.0,
                        None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                    );
                }

                if !drew_preview {
                    self.draw_page_preview_content(page_rect, shell)?;
                    drew_preview = true;
                }
            }
        }

        self.draw_canvas_scrollbars(canvas_rect, shell)
    }

    fn draw_page_preview_content(&self, page_rect: D2D_RECT_F, shell: &ShellRenderState) -> Result<()> {
        let left_pad = 44.0;
        let top_pad = 46.0;
        let right_pad = 40.0;
        let bottom_pad = 34.0;
        let text_rect = D2D_RECT_F {
            left: page_rect.left + left_pad,
            top: page_rect.top + top_pad,
            right: page_rect.right - right_pad,
            bottom: page_rect.bottom - bottom_pad,
        };

        let line_highlight = self.create_brush(
            crate::ui::Color::rgba(
                self.theme.selection_bg.r,
                self.theme.selection_bg.g,
                self.theme.selection_bg.b,
                0.16,
            )
            .as_d2d(),
        )?;
        let selection = self.create_brush(
            crate::ui::Color::rgba(
                self.theme.selection_bg.r,
                self.theme.selection_bg.g,
                self.theme.selection_bg.b,
                0.26,
            )
            .as_d2d(),
        )?;
        unsafe {
            let current_line = D2D_RECT_F {
                left: text_rect.left,
                top: text_rect.top + 2.0,
                right: text_rect.right,
                bottom: text_rect.top + 24.0,
            };
            self.d2d_context.FillRectangle(&current_line, &line_highlight);
            let selection_rect = D2D_RECT_F {
                left: text_rect.left + 2.0,
                top: text_rect.top + 3.0,
                right: (text_rect.left + 220.0).min(text_rect.right),
                bottom: text_rect.top + 22.0,
            };
            self.d2d_context.FillRectangle(&selection_rect, &selection);
        }

        let preview = shell
            .canvas_preview_lines
            .iter()
            .take(42)
            .map(|s| s.as_str())
            .collect::<Vec<_>>()
            .join("\n");
        let text = preview.encode_utf16().collect::<Vec<u16>>();
        let text_format = self.create_text_format()?;
        let text_brush = self.create_brush(self.theme.text_primary.as_d2d())?;
        unsafe {
            self.d2d_context.DrawText(
                &text,
                &text_format,
                &text_rect,
                &text_brush,
                D2D1_DRAW_TEXT_OPTIONS_NONE,
                DWRITE_MEASURING_MODE_NATURAL,
            );
        }

        if shell.canvas_cursor_visible {
            let cursor_brush = self.create_brush(self.theme.accent.as_d2d())?;
            unsafe {
                let cursor = D2D_RECT_F {
                    left: text_rect.left + 2.0,
                    top: text_rect.top + 2.0,
                    right: text_rect.left + 4.0,
                    bottom: text_rect.top + 22.0,
                };
                self.d2d_context.FillRectangle(&cursor, &cursor_brush);
            }
        }

        Ok(())
    }

    fn draw_canvas_scrollbars(&self, canvas_rect: D2D_RECT_F, shell: &ShellRenderState) -> Result<()> {
        if !shell.canvas_scrollbar_visible && shell.canvas_scrollbar_alpha <= 0.01 {
            return Ok(());
        }

        let alpha = shell.canvas_scrollbar_alpha.clamp(0.0, 1.0);
        let thumb_color = crate::ui::Color::rgba(
            self.theme.text_secondary.r,
            self.theme.text_secondary.g,
            self.theme.text_secondary.b,
            0.45 * alpha,
        );
        let thumb_brush = self.create_brush(thumb_color.as_d2d())?;
        let thickness = 6.0;
        let viewport_h = shell.canvas_viewport_height.max(1.0);
        let viewport_w = shell.canvas_viewport_width.max(1.0);
        let content_h = shell.canvas_content_height.max(viewport_h);
        let content_w = shell.canvas_content_width.max(viewport_w);

        unsafe {
            if content_h > viewport_h + 1.0 {
                let thumb_h = ((viewport_h / content_h) * viewport_h).clamp(28.0, viewport_h - 6.0);
                let max_track = (viewport_h - thumb_h).max(1.0);
                let max_scroll = (content_h - viewport_h).max(1.0);
                let thumb_offset = (shell.canvas_scroll_y / max_scroll).clamp(0.0, 1.0) * max_track;
                let vbar = D2D_RECT_F {
                    left: canvas_rect.right - thickness - 2.0,
                    top: canvas_rect.top + thumb_offset + 2.0,
                    right: canvas_rect.right - 2.0,
                    bottom: canvas_rect.top + thumb_offset + thumb_h,
                };
                self.d2d_context.FillRectangle(&vbar, &thumb_brush);
            }

            if content_w > viewport_w + 1.0 {
                let thumb_w = ((viewport_w / content_w) * viewport_w).clamp(28.0, viewport_w - 6.0);
                let max_track = (viewport_w - thumb_w).max(1.0);
                let max_scroll = (content_w - viewport_w).max(1.0);
                let thumb_offset = (shell.canvas_scroll_x / max_scroll).clamp(0.0, 1.0) * max_track;
                let hbar = D2D_RECT_F {
                    left: canvas_rect.left + thumb_offset + 2.0,
                    top: canvas_rect.bottom - thickness - 2.0,
                    right: canvas_rect.left + thumb_offset + thumb_w,
                    bottom: canvas_rect.bottom - 2.0,
                };
                self.d2d_context.FillRectangle(&hbar, &thumb_brush);
            }
        }

        Ok(())
    }

    fn fill_rect(&self, rect: D2D_RECT_F, color: crate::ui::Color) -> Result<()> {
        let brush = self.create_brush(color.as_d2d())?;
        unsafe {
            self.d2d_context.FillRectangle(&rect, &brush);
        }
        Ok(())
    }

    fn fill_gradient(
        &self,
        rect: D2D_RECT_F,
        start: crate::ui::Color,
        end: crate::ui::Color,
        angle_degrees: f32,
    ) -> Result<()> {
        let stripes = 64usize;
        let width = (rect.right - rect.left).max(1.0);
        let angle_radians = angle_degrees.to_radians();
        let dx = angle_radians.cos();
        let dy = angle_radians.sin();
        for i in 0..stripes {
            let t0 = i as f32 / stripes as f32;
            let t1 = (i + 1) as f32 / stripes as f32;
            let mid = (t0 + t1) * 0.5;
            let color = Self::lerp_color(start, end, mid);
            let brush = self.create_brush(color.as_d2d())?;
            let x0 = rect.left + t0 * width;
            let x1 = rect.left + t1 * width;
            let tilt = dy * 14.0;
            let skew = dx * 6.0;
            let stripe = D2D_RECT_F {
                left: x0 - skew,
                top: rect.top - tilt,
                right: x1 + skew,
                bottom: rect.bottom + tilt,
            };
            unsafe {
                self.d2d_context.FillRectangle(&stripe, &brush);
            }
        }
        Ok(())
    }

    fn fill_pattern(
        &self,
        rect: D2D_RECT_F,
        style: PatternStyle,
        foreground: crate::ui::Color,
        background: crate::ui::Color,
        scale: f32,
    ) -> Result<()> {
        self.fill_rect(rect, background)?;
        let brush = self.create_brush(foreground.as_d2d())?;
        let step = (10.0 * scale.max(0.25)).clamp(4.0, 48.0);
        let width = rect.right - rect.left;
        let height = rect.bottom - rect.top;

        unsafe {
            match style {
                PatternStyle::Dots => {
                    let size = (1.8 * scale.max(0.5)).clamp(1.2, 4.0);
                    let mut y = rect.top;
                    while y <= rect.bottom {
                        let mut x = rect.left;
                        while x <= rect.right {
                            let dot = D2D_RECT_F {
                                left: x,
                                top: y,
                                right: x + size,
                                bottom: y + size,
                            };
                            self.d2d_context.FillRectangle(&dot, &brush);
                            x += step;
                        }
                        y += step;
                    }
                }
                PatternStyle::LinesHorizontal => {
                    let mut y = rect.top;
                    while y <= rect.bottom {
                        self.d2d_context.DrawLine(
                            Vector2 { X: rect.left, Y: y },
                            Vector2 { X: rect.right, Y: y },
                            &brush,
                            1.0,
                            None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                        );
                        y += step;
                    }
                }
                PatternStyle::LinesVertical => {
                    let mut x = rect.left;
                    while x <= rect.right {
                        self.d2d_context.DrawLine(
                            Vector2 { X: x, Y: rect.top },
                            Vector2 { X: x, Y: rect.bottom },
                            &brush,
                            1.0,
                            None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                        );
                        x += step;
                    }
                }
                PatternStyle::LinesDiagonal | PatternStyle::CrossHatch | PatternStyle::GraphPaper => {
                    if matches!(style, PatternStyle::GraphPaper) {
                        let mut y = rect.top;
                        while y <= rect.bottom {
                            self.d2d_context.DrawLine(
                                Vector2 { X: rect.left, Y: y },
                                Vector2 { X: rect.right, Y: y },
                                &brush,
                                1.0,
                                None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                            );
                            y += step;
                        }
                        let mut x = rect.left;
                        while x <= rect.right {
                            self.d2d_context.DrawLine(
                                Vector2 { X: x, Y: rect.top },
                                Vector2 { X: x, Y: rect.bottom },
                                &brush,
                                1.0,
                                None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                            );
                            x += step;
                        }
                    } else {
                        let mut offset = -height;
                        while offset <= width {
                            self.d2d_context.DrawLine(
                                Vector2 {
                                    X: rect.left + offset,
                                    Y: rect.top,
                                },
                                Vector2 {
                                    X: rect.left + offset + height,
                                    Y: rect.bottom,
                                },
                                &brush,
                                1.0,
                                None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                            );
                            if matches!(style, PatternStyle::CrossHatch) {
                                self.d2d_context.DrawLine(
                                    Vector2 {
                                        X: rect.left + offset + height,
                                        Y: rect.top,
                                    },
                                    Vector2 {
                                        X: rect.left + offset,
                                        Y: rect.bottom,
                                    },
                                    &brush,
                                    1.0,
                                    None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                                );
                            }
                            offset += step;
                        }
                    }
                }
                PatternStyle::Noise => {
                    let size = (1.0 * scale.max(0.5)).clamp(1.0, 2.2);
                    let cell = step.max(4.0);
                    let mut yi = 0u32;
                    let mut y = rect.top;
                    while y <= rect.bottom {
                        let mut xi = 0u32;
                        let mut x = rect.left;
                        while x <= rect.right {
                            let hash = ((xi.wrapping_mul(73856093)) ^ (yi.wrapping_mul(19349663))) & 0xFF;
                            if hash < 84 {
                                let dot = D2D_RECT_F {
                                    left: x,
                                    top: y,
                                    right: x + size,
                                    bottom: y + size,
                                };
                                self.d2d_context.FillRectangle(&dot, &brush);
                            }
                            x += cell;
                            xi = xi.wrapping_add(1);
                        }
                        y += cell;
                        yi = yi.wrapping_add(1);
                    }
                }
            }
        }
        Ok(())
    }

    fn draw_canvas_label(&self, rect: D2D_RECT_F, label: &str) -> Result<()> {
        let text = label.encode_utf16().collect::<Vec<u16>>();
        let text_format = self.create_text_format()?;
        let text_brush = self.create_brush(self.theme.text_secondary.as_d2d())?;
        unsafe {
            self.d2d_context.DrawText(
                &text,
                &text_format,
                &D2D_RECT_F {
                    left: rect.left + 12.0,
                    top: rect.top + 12.0,
                    right: rect.right - 12.0,
                    bottom: rect.top + 38.0,
                },
                &text_brush,
                D2D1_DRAW_TEXT_OPTIONS_NONE,
                DWRITE_MEASURING_MODE_NATURAL,
            );
        }
        Ok(())
    }

    fn lerp_color(a: crate::ui::Color, b: crate::ui::Color, t: f32) -> crate::ui::Color {
        let tt = t.clamp(0.0, 1.0);
        crate::ui::Color::rgba(
            a.r + (b.r - a.r) * tt,
            a.g + (b.g - a.g) * tt,
            a.b + (b.b - a.b) * tt,
            a.a + (b.a - a.a) * tt,
        )
    }

    fn create_text_format(&self) -> Result<IDWriteTextFormat> {
        unsafe {
            match self.dwrite_factory.CreateTextFormat(
                w!("Segoe UI Variable"),
                None,
                windows::Win32::Graphics::DirectWrite::DWRITE_FONT_WEIGHT_NORMAL,
                windows::Win32::Graphics::DirectWrite::DWRITE_FONT_STYLE_NORMAL,
                windows::Win32::Graphics::DirectWrite::DWRITE_FONT_STRETCH_NORMAL,
                14.0,
                w!("en-US"),
            ) {
                Ok(format) => Ok(format),
                Err(_) => self.dwrite_factory.CreateTextFormat(
                    w!("Segoe UI"),
                    None,
                    windows::Win32::Graphics::DirectWrite::DWRITE_FONT_WEIGHT_NORMAL,
                    windows::Win32::Graphics::DirectWrite::DWRITE_FONT_STYLE_NORMAL,
                    windows::Win32::Graphics::DirectWrite::DWRITE_FONT_STRETCH_NORMAL,
                    14.0,
                    w!("en-US"),
                ),
            }
        }
    }

    fn create_brush(&self, color: windows::Win32::Graphics::Direct2D::Common::D2D1_COLOR_F) -> Result<ID2D1SolidColorBrush> {
        unsafe { self.d2d_context.CreateSolidColorBrush(&color, None) }
    }

    unsafe fn recreate_target_bitmap(&mut self) -> Result<()> {
        let surface: IDXGISurface = unsafe { self.swap_chain.GetBuffer(0)? };
        let bitmap_props = D2D1_BITMAP_PROPERTIES1 {
            pixelFormat: D2D1_PIXEL_FORMAT {
                format: DXGI_FORMAT_B8G8R8A8_UNORM,
                alphaMode: D2D1_ALPHA_MODE_IGNORE,
            },
            dpiX: self.dpi,
            dpiY: self.dpi,
            bitmapOptions: D2D1_BITMAP_OPTIONS_TARGET | D2D1_BITMAP_OPTIONS_CANNOT_DRAW,
            colorContext: ManuallyDrop::new(None),
        };

        let bitmap = unsafe {
            self.d2d_context
                .CreateBitmapFromDxgiSurface(&surface, Some(&bitmap_props))?
        };

        unsafe {
            self.d2d_context.SetTarget(&bitmap);
            let _ = self.d2d_context.SetDpi(self.dpi, self.dpi);
        }
        self.target_bitmap = Some(bitmap);

        Ok(())
    }
}
