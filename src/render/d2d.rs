use std::mem::ManuallyDrop;
use std::time::Instant;

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
    app::AppState,
    render::perf::{DebugPerformancePanel, query_process_working_set_bytes},
    theme::Theme,
};

const D2DERR_RECREATE_TARGET: HRESULT = HRESULT(0x8899000C_u32 as i32);

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

    pub fn render(&mut self, app_state: &AppState) -> Result<()> {
        crate::profile_scope!("renderer.frame");
        let frame_start = Instant::now();

        unsafe {
            self.d2d_context.BeginDraw();
            let clear = self.theme.window_bg.as_d2d();
            self.d2d_context.Clear(Some(&clear));

            self.draw_shell_placeholder(app_state)?;

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

    pub fn debug_panel(&self) -> &DebugPerformancePanel {
        &self.debug_panel
    }

    fn draw_shell_placeholder(&self, app_state: &AppState) -> Result<()> {
        unsafe {
            let mut rect = RECT::default();
            GetClientRect(self.hwnd, &mut rect)?;

            let width = (rect.right - rect.left) as f32;
            let height = (rect.bottom - rect.top) as f32;

            let tab_h = if app_state.show_tabs { 36.0 } else { 0.0 };
            let sidebar_w = if app_state.show_sidebar {
                app_state.sidebar_width.clamp(180.0, 420.0)
            } else {
                0.0
            };
            let toolbar_h = if app_state.show_toolbar { 44.0 } else { 0.0 };
            let status_h = if app_state.show_statusbar { 28.0 } else { 0.0 };

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
            let canvas_brush = self.create_brush(self.theme.canvas_bg.as_d2d())?;
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
            self.d2d_context.FillRectangle(&canvas_rect, &canvas_brush);
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
            }

            let status = app_state.status_text.encode_utf16().collect::<Vec<u16>>();
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
            }

            let hint = format!(
                "Ctrl+B Sidebar [{}]  Ctrl+T Toolbar [{}]  Ctrl+L Status [{}]",
                if app_state.show_sidebar { "On" } else { "Off" },
                if app_state.show_toolbar { "On" } else { "Off" },
                if app_state.show_statusbar { "On" } else { "Off" }
            );
            let hint = hint.encode_utf16().collect::<Vec<u16>>();
            self.d2d_context.DrawText(
                &hint,
                &text_format,
                &D2D_RECT_F {
                    left: sidebar_w + 16.0,
                    top: tab_h + toolbar_h + 12.0,
                    right: width - 12.0,
                    bottom: tab_h + toolbar_h + 40.0,
                },
                &text_brush,
                D2D1_DRAW_TEXT_OPTIONS_NONE,
                DWRITE_MEASURING_MODE_NATURAL,
            );

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
