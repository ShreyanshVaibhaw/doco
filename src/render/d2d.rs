use std::cell::RefCell;
use std::collections::HashMap;
use std::time::{Instant, SystemTime, UNIX_EPOCH};
use std::{mem::ManuallyDrop, path::Path};

use windows::{
    Win32::{
        Foundation::{HMODULE, HWND, RECT},
        Graphics::{
            Direct2D::{
                Common::{D2D_RECT_F, D2D1_ALPHA_MODE_IGNORE, D2D1_PIXEL_FORMAT},
                D2D1_BITMAP_OPTIONS_CANNOT_DRAW, D2D1_BITMAP_OPTIONS_TARGET,
                D2D1_BITMAP_PROPERTIES1, D2D1_DEVICE_CONTEXT_OPTIONS_NONE,
                D2D1_DRAW_TEXT_OPTIONS_CLIP, D2D1_DRAW_TEXT_OPTIONS_NONE,
                D2D1_FACTORY_TYPE_SINGLE_THREADED, D2D1CreateFactory,
                ID2D1Bitmap1, ID2D1Device, ID2D1DeviceContext, ID2D1Factory1, ID2D1Image,
                ID2D1SolidColorBrush,
            },
            Direct3D::{D3D_DRIVER_TYPE_HARDWARE, D3D_DRIVER_TYPE_WARP},
            Direct3D11::{
                D3D11_CREATE_DEVICE_BGRA_SUPPORT, D3D11_SDK_VERSION, D3D11CreateDevice,
                ID3D11Device, ID3D11DeviceContext,
            },
            DirectWrite::{
                DWRITE_FACTORY_TYPE_SHARED, DWRITE_MEASURING_MODE_NATURAL, DWriteCreateFactory,
                DWRITE_PARAGRAPH_ALIGNMENT_CENTER, DWRITE_TEXT_ALIGNMENT_CENTER,
                DWRITE_WORD_WRAPPING_NO_WRAP, IDWriteFactory, IDWriteTextFormat,
            },
            Dxgi::{
                Common::{
                    DXGI_ALPHA_MODE_IGNORE, DXGI_ALPHA_MODE_UNSPECIFIED,
                    DXGI_FORMAT_B8G8R8A8_UNORM, DXGI_FORMAT_UNKNOWN, DXGI_SAMPLE_DESC,
                },
                DXGI_PRESENT, DXGI_SCALING_STRETCH, DXGI_SWAP_CHAIN_DESC1, DXGI_SWAP_CHAIN_FLAG,
                DXGI_SWAP_EFFECT_DISCARD, DXGI_SWAP_EFFECT_FLIP_DISCARD,
                DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL, DXGI_USAGE_RENDER_TARGET_OUTPUT, IDXGIDevice,
                IDXGIFactory2, IDXGISurface, IDXGISwapChain1,
            },
        },
        UI::WindowsAndMessaging::GetClientRect,
    },
    core::{HRESULT, Interface, Result, w},
};
use windows_numerics::Vector2;

use crate::{
    render::image_cache::ImageCacheStats,
    render::perf::{DebugPerformancePanel, query_process_working_set_bytes},
    theme::{
        Theme,
        backgrounds::{BackgroundKind, BackgroundSettings, PatternStyle, preset_by_id},
    },
    ui::Rect as UiRect,
};

const D2DERR_RECREATE_TARGET: HRESULT = HRESULT(0x8899000C_u32 as i32);
const LAYOUT_DPI: f32 = 96.0;

#[derive(Debug, Clone, Default)]
pub struct CanvasImageShellItem {
    pub block_id: u64,
    pub rect: UiRect,
    pub selected: bool,
    pub interpolation: String,
    pub alt_text: String,
}

#[derive(Debug, Clone, Default)]
pub struct CanvasTableShellItem {
    pub table_id: u64,
    pub rect: UiRect,
    pub rows: usize,
    pub cols: usize,
    pub cell_w: f32,
    pub cell_h: f32,
    pub header_h: f32,
    pub gutter_w: f32,
    pub selected: bool,
    pub selection_mode: u8,
    pub selection_start_row: usize,
    pub selection_start_col: usize,
    pub selection_end_row: usize,
    pub selection_end_col: usize,
}

#[derive(Debug, Clone, Default)]
pub struct ToastShellItem {
    pub title: String,
    pub body: String,
    pub opacity: f32,
    pub slide_offset: f32,
}

#[derive(Debug, Clone, Default)]
pub struct ToolbarShellButton {
    pub rect: UiRect,
    pub label: String,
    pub icon: String,
    pub active: bool,
    pub enabled: bool,
}

#[derive(Debug, Clone, Default)]
pub struct ShellRenderState {
    pub ui_scale: f32,
    pub show_tabs: bool,
    pub show_sidebar: bool,
    pub sidebar_width: f32,
    pub sidebar_resizing: bool,
    pub show_toolbar: bool,
    pub show_statusbar: bool,
    pub status_text: String,
    pub tab_titles: Vec<String>,
    pub active_tab: usize,
    pub tab_transition_progress: f32,
    pub tab_transition_offset: f32,
    pub tab_has_overflow_left: bool,
    pub tab_has_overflow_right: bool,
    pub toolbar_buttons: Vec<ToolbarShellButton>,
    pub toolbar_dropdown_open: bool,
    pub toolbar_dropdown_opacity: f32,
    pub toolbar_dropdown_scale: f32,
    pub active_sidebar_panel: String,
    pub sidebar_summary: String,
    pub sidebar_rows: Vec<String>,
    pub command_palette_open: bool,
    pub command_palette_opacity: f32,
    pub command_palette_offset_y: f32,
    pub command_palette_query: String,
    pub command_palette_results: Vec<String>,
    pub command_palette_selected: usize,
    pub settings_visible: bool,
    pub settings_query: String,
    pub settings_category: String,
    pub settings_categories: Vec<String>,
    pub settings_rows: Vec<String>,
    pub settings_selected_row: usize,
    pub settings_conflicts: bool,
    pub settings_save_error: String,
    pub table_picker_visible: bool,
    pub table_picker_rows: usize,
    pub table_picker_cols: usize,
    pub table_picker_custom_rows: String,
    pub table_picker_custom_cols: String,
    pub table_picker_custom_focus_rows: bool,
    pub find_visible: bool,
    pub replace_visible: bool,
    pub find_query: String,
    pub replace_query: String,
    pub find_result_count: String,
    pub find_case_sensitive: bool,
    pub find_whole_word: bool,
    pub find_regex: bool,
    pub find_preview: String,
    pub find_current: usize,
    pub find_total: usize,
    pub find_capture_groups: Vec<String>,
    pub goto_visible: bool,
    pub goto_input: String,
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
    pub canvas_images: Vec<CanvasImageShellItem>,
    pub canvas_tables: Vec<CanvasTableShellItem>,
    pub toast_entries: Vec<ToastShellItem>,
    pub accessibility_high_contrast: bool,
    pub accessibility_reduce_motion: bool,
    pub image_toolbar_visible: bool,
    pub image_properties_visible: bool,
    pub image_selected_size: String,
    pub image_selected_meta: String,
    pub image_selected_alt_text: String,
    pub table_selected_meta: String,
    pub table_selected_id: u64,
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
    brush_cache: RefCell<HashMap<u32, ID2D1SolidColorBrush>>,
    default_text_format: RefCell<Option<IDWriteTextFormat>>,
    icon_text_format: RefCell<Option<IDWriteTextFormat>>,
}

impl D2DRenderer {
    pub fn new(hwnd: HWND, width: u32, height: u32, dpi: f32, theme: Theme) -> Result<Self> {
        unsafe {
            let d2d_factory: ID2D1Factory1 =
                D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, None)?;

            let (d3d_device, d3d_context) = Self::create_d3d_device()?;

            let dxgi_device: IDXGIDevice = d3d_device.cast()?;
            let adapter = dxgi_device.GetAdapter()?;
            let dxgi_factory: IDXGIFactory2 = adapter.GetParent()?;

            let swap_chain =
                Self::create_swap_chain_for_hwnd(&dxgi_factory, &d3d_device, hwnd, width, height)?;

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
                brush_cache: RefCell::new(HashMap::new()),
                default_text_format: RefCell::new(None),
                icon_text_format: RefCell::new(None),
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
            (
                DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL,
                2,
                DXGI_ALPHA_MODE_IGNORE,
                "flip_sequential",
            ),
            (
                DXGI_SWAP_EFFECT_FLIP_DISCARD,
                2,
                DXGI_ALPHA_MODE_IGNORE,
                "flip_discard",
            ),
            (
                DXGI_SWAP_EFFECT_DISCARD,
                1,
                DXGI_ALPHA_MODE_UNSPECIFIED,
                "discard",
            ),
        ];

        let mut last_error = None;
        for (swap_effect, buffer_count, alpha_mode, label) in attempts {
            let swap_chain_desc = DXGI_SWAP_CHAIN_DESC1 {
                Width: width,
                Height: height,
                Format: DXGI_FORMAT_B8G8R8A8_UNORM,
                Stereo: false.into(),
                SampleDesc: DXGI_SAMPLE_DESC {
                    Count: 1,
                    Quality: 0,
                },
                BufferUsage: DXGI_USAGE_RENDER_TARGET_OUTPUT,
                BufferCount: buffer_count,
                Scaling: DXGI_SCALING_STRETCH,
                SwapEffect: swap_effect,
                AlphaMode: alpha_mode,
                Flags: 0,
            };

            let result = unsafe {
                dxgi_factory.CreateSwapChainForHwnd(d3d_device, hwnd, &swap_chain_desc, None, None)
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
            let _ = self.d2d_context.SetDpi(LAYOUT_DPI, LAYOUT_DPI);
        }
    }

    pub fn resize(&mut self, width: u32, height: u32) -> Result<()> {
        if width == 0 || height == 0 {
            return Ok(());
        }

        unsafe {
            self.d2d_context.SetTarget(None::<&ID2D1Image>);
            self.target_bitmap = None;
            self.swap_chain.ResizeBuffers(
                0,
                width,
                height,
                DXGI_FORMAT_UNKNOWN,
                DXGI_SWAP_CHAIN_FLAG(0),
            )?;
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

    pub fn update_image_cache_stats(&mut self, stats: ImageCacheStats) {
        self.debug_panel.update_image_cache_stats(stats);
    }

    pub fn set_theme(&mut self, theme: Theme) {
        self.theme = theme;
        self.brush_cache.borrow_mut().clear();
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
            let ui_scale = shell.ui_scale.clamp(1.0, 2.0);

            let tab_h = if shell.show_tabs { 36.0 * ui_scale } else { 0.0 };
            let sidebar_w = if shell.show_sidebar {
                shell
                    .sidebar_width
                    .clamp(200.0, 400.0)
                    .min((width - 80.0).max(0.0))
            } else {
                0.0
            };
            let toolbar_h = if shell.show_toolbar { 44.0 * ui_scale } else { 0.0 };
            let status_h = if shell.show_statusbar { 28.0 * ui_scale } else { 0.0 };

            let tab_rect = D2D_RECT_F {
                left: 0.0,
                top: 0.0,
                right: width,
                bottom: tab_h,
            };
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

            let mut tab_color = self.theme.surface_secondary.as_d2d();
            let mut side_color = self.theme.surface_primary.as_d2d();
            let mut tool_color = self.theme.surface_secondary.as_d2d();
            let mut status_color = self.theme.surface_primary.as_d2d();
            if shell.accessibility_high_contrast {
                tab_color = crate::ui::Color::rgb(0.08, 0.08, 0.08).as_d2d();
                side_color = crate::ui::Color::rgb(0.0, 0.0, 0.0).as_d2d();
                tool_color = crate::ui::Color::rgb(0.08, 0.08, 0.08).as_d2d();
                status_color = crate::ui::Color::rgb(0.0, 0.0, 0.0).as_d2d();
            }
            let tab_brush = self.create_brush(tab_color)?;
            let side_brush = self.create_brush(side_color)?;
            let tool_brush = self.create_brush(tool_color)?;
            let status_brush = self.create_brush(status_color)?;

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
                self.d2d_context
                    .FillRectangle(&splitter_rect, &splitter_brush);
            }

            let text_color = if shell.accessibility_high_contrast {
                crate::ui::Color::rgb(1.0, 1.0, 1.0).as_d2d()
            } else {
                self.theme.text_primary.as_d2d()
            };
            let text_brush = self.create_brush(text_color)?;
            let text_format = self.create_text_format()?;
            if sidebar_w > 0.0 {
                let tab_titles = [
                    ("Files", "Files"),
                    ("Outline", "Outline"),
                    ("Marks", "Bookmarks"),
                    ("Search", "Search Results"),
                ];
                let tab_w = sidebar_w / tab_titles.len() as f32;
                for (idx, (title, panel_key)) in tab_titles.iter().enumerate() {
                    let x = idx as f32 * tab_w;
                    let tab_rect = D2D_RECT_F {
                        left: x,
                        top: tab_h,
                        right: (x + tab_w).min(sidebar_w),
                        bottom: tab_h + 34.0,
                    };
                    let is_active = shell
                        .active_sidebar_panel
                        .eq_ignore_ascii_case(panel_key);
                    if is_active {
                        let bg = self.create_brush(self.theme.surface_hover.as_d2d())?;
                        self.d2d_context.FillRectangle(&tab_rect, &bg);
                    }
                    let text = title.encode_utf16().collect::<Vec<u16>>();
                    self.d2d_context.DrawText(
                        &text,
                        &text_format,
                        &tab_rect,
                        &text_brush,
                        D2D1_DRAW_TEXT_OPTIONS_CLIP,
                        DWRITE_MEASURING_MODE_NATURAL,
                    );
                    if idx + 1 < tab_titles.len() {
                        let divider = self.create_brush(self.theme.border_subtle.as_d2d())?;
                        self.d2d_context.DrawLine(
                            Vector2 {
                                X: tab_rect.right,
                                Y: tab_rect.top + 5.0,
                            },
                            Vector2 {
                                X: tab_rect.right,
                                Y: tab_rect.bottom - 5.0,
                            },
                            &divider,
                            1.0,
                            None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                        );
                    }
                }

                let panel_top = tab_h + 34.0;
                let panel_bottom = (height - status_h).max(panel_top);
                let mut row_y = panel_top;
                for row in shell.sidebar_rows.iter().take(24) {
                    let row_bottom = row_y + 24.0;
                    if row_bottom > panel_bottom {
                        break;
                    }
                    let row_utf16 = row.encode_utf16().collect::<Vec<u16>>();
                    self.d2d_context.DrawText(
                        &row_utf16,
                        &text_format,
                        &D2D_RECT_F {
                            left: 10.0,
                            top: row_y,
                            right: sidebar_w - 8.0,
                            bottom: row_bottom,
                        },
                        &text_brush,
                        D2D1_DRAW_TEXT_OPTIONS_CLIP,
                        DWRITE_MEASURING_MODE_NATURAL,
                    );
                    row_y += 24.0;
                }
            }

            if shell.command_palette_open {
                let palette_w = 600.0_f32.min((width - 24.0).max(320.0));
                let palette_h = 400.0_f32.min((height - 28.0).max(120.0));
                let palette_x = (width - palette_w) * 0.5;
                let palette_y = 20.0 + shell.command_palette_offset_y;
                let palette_rect = D2D_RECT_F {
                    left: palette_x,
                    top: palette_y,
                    right: palette_x + palette_w,
                    bottom: palette_y + palette_h,
                };

                let mut overlay = self.theme.surface_primary.as_d2d();
                overlay.a = 0.94 * shell.command_palette_opacity.clamp(0.0, 1.0);
                let overlay_brush = self.create_brush(overlay)?;
                self.d2d_context
                    .FillRectangle(&palette_rect, &overlay_brush);

                let mut border = self.theme.border_default.as_d2d();
                border.a *= shell.command_palette_opacity.clamp(0.0, 1.0);
                let border_brush = self.create_brush(border)?;
                self.d2d_context.DrawRectangle(
                    &palette_rect,
                    &border_brush,
                    1.0,
                    None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                );

                let query_text = format!("> {}", shell.command_palette_query);
                let query_utf16 = query_text.encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &query_utf16,
                    &text_format,
                    &D2D_RECT_F {
                        left: palette_x + 14.0,
                        top: palette_y + 10.0,
                        right: palette_x + palette_w - 12.0,
                        bottom: palette_y + 36.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                let mut row_y = palette_y + 44.0;
                for (idx, row) in shell.command_palette_results.iter().enumerate() {
                    if row_y + 24.0 > palette_y + palette_h - 8.0 {
                        break;
                    }
                    if idx == shell.command_palette_selected {
                        let highlight_brush =
                            self.create_brush(self.theme.surface_hover.as_d2d())?;
                        self.d2d_context.FillRectangle(
                            &D2D_RECT_F {
                                left: palette_x + 8.0,
                                top: row_y - 1.0,
                                right: palette_x + palette_w - 8.0,
                                bottom: row_y + 21.0,
                            },
                            &highlight_brush,
                        );
                    }
                    let row_utf16 = row.encode_utf16().collect::<Vec<u16>>();
                    self.d2d_context.DrawText(
                        &row_utf16,
                        &text_format,
                        &D2D_RECT_F {
                            left: palette_x + 14.0,
                            top: row_y,
                            right: palette_x + palette_w - 12.0,
                            bottom: row_y + 20.0,
                        },
                        &text_brush,
                        D2D1_DRAW_TEXT_OPTIONS_NONE,
                        DWRITE_MEASURING_MODE_NATURAL,
                    );
                    row_y += 22.0;
                }
            }

            if shell.settings_visible && !shell.command_palette_open {
                let mut dim = self.theme.window_bg.as_d2d();
                dim.a = 0.58;
                let dim_brush = self.create_brush(dim)?;
                self.d2d_context.FillRectangle(
                    &D2D_RECT_F {
                        left: 0.0,
                        top: 0.0,
                        right: width,
                        bottom: height,
                    },
                    &dim_brush,
                );

                let panel = D2D_RECT_F {
                    left: 18.0,
                    top: 18.0,
                    right: width - 18.0,
                    bottom: height - 18.0,
                };
                let panel_bg = self.create_brush(self.theme.surface_primary.as_d2d())?;
                let panel_border = self.create_brush(self.theme.border_default.as_d2d())?;
                self.d2d_context.FillRectangle(&panel, &panel_bg);
                self.d2d_context.DrawRectangle(
                    &panel,
                    &panel_border,
                    1.0,
                    None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                );

                let split_x = (panel.left + 250.0).min(panel.right - 220.0);
                let left_pane = D2D_RECT_F {
                    left: panel.left,
                    top: panel.top,
                    right: split_x,
                    bottom: panel.bottom,
                };
                let left_bg = self.create_brush(self.theme.surface_secondary.as_d2d())?;
                self.d2d_context.FillRectangle(&left_pane, &left_bg);
                self.d2d_context.DrawLine(
                    Vector2 {
                        X: split_x,
                        Y: panel.top,
                    },
                    Vector2 {
                        X: split_x,
                        Y: panel.bottom,
                    },
                    &panel_border,
                    1.0,
                    None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                );

                let title = "Settings".encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &title,
                    &text_format,
                    &D2D_RECT_F {
                        left: panel.left + 14.0,
                        top: panel.top + 10.0,
                        right: panel.right - 12.0,
                        bottom: panel.top + 34.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                let search = format!("Search: {}", shell.settings_query)
                    .encode_utf16()
                    .collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &search,
                    &text_format,
                    &D2D_RECT_F {
                        left: panel.left + 14.0,
                        top: panel.top + 36.0,
                        right: panel.right - 12.0,
                        bottom: panel.top + 58.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                let mut category_y = panel.top + 66.0;
                for category in &shell.settings_categories {
                    let category_rect = D2D_RECT_F {
                        left: panel.left + 10.0,
                        top: category_y,
                        right: split_x - 8.0,
                        bottom: category_y + 24.0,
                    };
                    if category == &shell.settings_category {
                        let active = self.create_brush(self.theme.surface_hover.as_d2d())?;
                        self.d2d_context.FillRectangle(&category_rect, &active);
                    }
                    let category_text = category.encode_utf16().collect::<Vec<u16>>();
                    self.d2d_context.DrawText(
                        &category_text,
                        &text_format,
                        &category_rect,
                        &text_brush,
                        D2D1_DRAW_TEXT_OPTIONS_NONE,
                        DWRITE_MEASURING_MODE_NATURAL,
                    );
                    category_y += 26.0;
                    if category_y + 20.0 > panel.bottom - 40.0 {
                        break;
                    }
                }

                let heading = shell.settings_category.encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &heading,
                    &text_format,
                    &D2D_RECT_F {
                        left: split_x + 12.0,
                        top: panel.top + 10.0,
                        right: panel.right - 12.0,
                        bottom: panel.top + 34.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                let mut row_y = panel.top + 44.0;
                for (index, row) in shell.settings_rows.iter().enumerate() {
                    let row_rect = D2D_RECT_F {
                        left: split_x + 10.0,
                        top: row_y,
                        right: panel.right - 10.0,
                        bottom: row_y + 24.0,
                    };
                    if index == shell.settings_selected_row {
                        let selected = self.create_brush(self.theme.surface_hover.as_d2d())?;
                        self.d2d_context.FillRectangle(&row_rect, &selected);
                    }
                    let row_text = row.encode_utf16().collect::<Vec<u16>>();
                    self.d2d_context.DrawText(
                        &row_text,
                        &text_format,
                        &row_rect,
                        &text_brush,
                        D2D1_DRAW_TEXT_OPTIONS_NONE,
                        DWRITE_MEASURING_MODE_NATURAL,
                    );
                    row_y += 26.0;
                    if row_y + 24.0 > panel.bottom - 46.0 {
                        break;
                    }
                }

                let footer_text = if shell.settings_conflicts {
                    "Conflicting shortcuts detected. Adjust bindings or reset defaults."
                } else {
                    "Click a setting row to cycle values. Enter toggles selected row. Esc closes."
                };
                let footer = footer_text.encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &footer,
                    &text_format,
                    &D2D_RECT_F {
                        left: split_x + 12.0,
                        top: panel.bottom - 36.0,
                        right: panel.right - 12.0,
                        bottom: panel.bottom - 12.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                if !shell.settings_save_error.trim().is_empty() {
                    let err = format!("Save error: {}", shell.settings_save_error)
                        .encode_utf16()
                        .collect::<Vec<u16>>();
                    let err_brush = self.create_brush(self.theme.text_accent.as_d2d())?;
                    self.d2d_context.DrawText(
                        &err,
                        &text_format,
                        &D2D_RECT_F {
                            left: split_x + 12.0,
                            top: panel.bottom - 60.0,
                            right: panel.right - 12.0,
                            bottom: panel.bottom - 38.0,
                        },
                        &err_brush,
                        D2D1_DRAW_TEXT_OPTIONS_NONE,
                        DWRITE_MEASURING_MODE_NATURAL,
                    );
                }
            }

            if shell.table_picker_visible && !shell.command_palette_open && !shell.settings_visible {
                let picker_w = 292.0_f32.min((canvas_rect.right - canvas_rect.left - 20.0).max(220.0));
                let picker_h = 236.0;
                let picker_x = canvas_rect.left + 10.0;
                let picker_y = canvas_rect.top + 10.0;
                let picker = D2D_RECT_F {
                    left: picker_x,
                    top: picker_y,
                    right: picker_x + picker_w,
                    bottom: picker_y + picker_h,
                };
                let panel_bg = self.create_brush(self.theme.surface_primary.as_d2d())?;
                let panel_border = self.create_brush(self.theme.border_default.as_d2d())?;
                self.d2d_context.FillRectangle(&picker, &panel_bg);
                self.d2d_context.DrawRectangle(
                    &picker,
                    &panel_border,
                    1.0,
                    None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                );

                let title = "Insert Table".encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &title,
                    &text_format,
                    &D2D_RECT_F {
                        left: picker.left + 10.0,
                        top: picker.top + 8.0,
                        right: picker.right - 10.0,
                        bottom: picker.top + 28.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                let grid_left = picker.left + 12.0;
                let grid_top = picker.top + 34.0;
                let cell = 16.0f32;
                let selected_fill = self.create_brush(
                    crate::ui::Color::rgba(
                        self.theme.accent.r,
                        self.theme.accent.g,
                        self.theme.accent.b,
                        0.42,
                    )
                    .as_d2d(),
                )?;
                let grid_fill = self.create_brush(self.theme.surface_secondary.as_d2d())?;
                let grid_line = self.create_brush(self.theme.border_subtle.as_d2d())?;
                for r in 0..10usize {
                    for c in 0..10usize {
                        let left = grid_left + c as f32 * cell;
                        let top = grid_top + r as f32 * cell;
                        let rect = D2D_RECT_F {
                            left,
                            top,
                            right: left + cell - 1.0,
                            bottom: top + cell - 1.0,
                        };
                        let selected = r < shell.table_picker_rows && c < shell.table_picker_cols;
                        self.d2d_context.FillRectangle(
                            &rect,
                            if selected { &selected_fill } else { &grid_fill },
                        );
                        self.d2d_context.DrawRectangle(
                            &rect,
                            &grid_line,
                            1.0,
                            None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                        );
                    }
                }

                let dims = format!("{} x {}", shell.table_picker_rows, shell.table_picker_cols);
                let dims_utf16 = dims.encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &dims_utf16,
                    &text_format,
                    &D2D_RECT_F {
                        left: grid_left,
                        top: grid_top + 168.0,
                        right: picker.right - 10.0,
                        bottom: grid_top + 188.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                let fields = format!(
                    "Rows [{}]: {}   Cols [{}]: {}",
                    if shell.table_picker_custom_focus_rows {
                        "x"
                    } else {
                        " "
                    },
                    shell.table_picker_custom_rows,
                    if shell.table_picker_custom_focus_rows {
                        " "
                    } else {
                        "x"
                    },
                    shell.table_picker_custom_cols
                );
                let fields_utf16 = fields.encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &fields_utf16,
                    &text_format,
                    &D2D_RECT_F {
                        left: grid_left,
                        top: grid_top + 188.0,
                        right: picker.right - 10.0,
                        bottom: grid_top + 208.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                let hint = "[Enter] Insert  [Esc] Cancel  [Tab] Switch Field  [Arrows] Resize";
                let hint_utf16 = hint.encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &hint_utf16,
                    &text_format,
                    &D2D_RECT_F {
                        left: grid_left,
                        top: grid_top + 208.0,
                        right: picker.right - 10.0,
                        bottom: picker.bottom - 8.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );
            }

            if shell.find_visible && !shell.command_palette_open && !shell.settings_visible {
                let panel_w =
                    460.0_f32.min((canvas_rect.right - canvas_rect.left - 20.0).max(300.0));
                let panel_h = if shell.replace_visible { 188.0 } else { 136.0 };
                let panel_x = (canvas_rect.right - panel_w - 10.0).max(canvas_rect.left + 8.0);
                let panel_y = canvas_rect.top + 10.0;
                let panel_rect = D2D_RECT_F {
                    left: panel_x,
                    top: panel_y,
                    right: panel_x + panel_w,
                    bottom: panel_y + panel_h,
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

                let title = if shell.replace_visible {
                    "Find & Replace"
                } else {
                    "Find"
                };
                let title_utf16 = title.encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &title_utf16,
                    &text_format,
                    &D2D_RECT_F {
                        left: panel_x + 10.0,
                        top: panel_y + 6.0,
                        right: panel_x + panel_w - 10.0,
                        bottom: panel_y + 24.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                let find_line = format!("Find: {}", shell.find_query);
                let find_utf16 = find_line.encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &find_utf16,
                    &text_format,
                    &D2D_RECT_F {
                        left: panel_x + 10.0,
                        top: panel_y + 28.0,
                        right: panel_x + panel_w - 10.0,
                        bottom: panel_y + 48.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                let options = format!(
                    "[{}] Case  [{}] Word  [{}] Regex   [Shift+Enter] Prev  [Enter] Next  [Esc] Close",
                    if shell.find_case_sensitive { "x" } else { " " },
                    if shell.find_whole_word { "x" } else { " " },
                    if shell.find_regex { "x" } else { " " }
                );
                let options_utf16 = options.encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &options_utf16,
                    &text_format,
                    &D2D_RECT_F {
                        left: panel_x + 10.0,
                        top: panel_y + 48.0,
                        right: panel_x + panel_w - 10.0,
                        bottom: panel_y + 68.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                let count_line = format!(
                    "{} ({}/{})",
                    shell.find_result_count, shell.find_current, shell.find_total
                );
                let count_utf16 = count_line.encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &count_utf16,
                    &text_format,
                    &D2D_RECT_F {
                        left: panel_x + 10.0,
                        top: panel_y + 68.0,
                        right: panel_x + panel_w - 10.0,
                        bottom: panel_y + 88.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                if shell.replace_visible {
                    let replace_line = format!("Replace: {}", shell.replace_query);
                    let replace_utf16 = replace_line.encode_utf16().collect::<Vec<u16>>();
                    self.d2d_context.DrawText(
                        &replace_utf16,
                        &text_format,
                        &D2D_RECT_F {
                            left: panel_x + 10.0,
                            top: panel_y + 88.0,
                            right: panel_x + panel_w - 10.0,
                            bottom: panel_y + 108.0,
                        },
                        &text_brush,
                        D2D1_DRAW_TEXT_OPTIONS_NONE,
                        DWRITE_MEASURING_MODE_NATURAL,
                    );

                    let actions = "[Ctrl+Enter] Replace Current   [Ctrl+Shift+Enter] Replace All"
                        .encode_utf16()
                        .collect::<Vec<u16>>();
                    self.d2d_context.DrawText(
                        &actions,
                        &text_format,
                        &D2D_RECT_F {
                            left: panel_x + 10.0,
                            top: panel_y + 108.0,
                            right: panel_x + panel_w - 10.0,
                            bottom: panel_y + 128.0,
                        },
                        &text_brush,
                        D2D1_DRAW_TEXT_OPTIONS_NONE,
                        DWRITE_MEASURING_MODE_NATURAL,
                    );
                }

                let preview_top = if shell.replace_visible {
                    panel_y + 130.0
                } else {
                    panel_y + 90.0
                };
                let preview_line = format!("Preview: {}", shell.find_preview);
                let preview_utf16 = preview_line.encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &preview_utf16,
                    &text_format,
                    &D2D_RECT_F {
                        left: panel_x + 10.0,
                        top: preview_top,
                        right: panel_x + panel_w - 10.0,
                        bottom: panel_rect.bottom - 8.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                if shell.find_regex && !shell.find_capture_groups.is_empty() {
                    let mut gy = preview_top + 18.0;
                    for g in shell.find_capture_groups.iter().take(3) {
                        let g_utf16 = g.encode_utf16().collect::<Vec<u16>>();
                        self.d2d_context.DrawText(
                            &g_utf16,
                            &text_format,
                            &D2D_RECT_F {
                                left: panel_x + 14.0,
                                top: gy,
                                right: panel_x + panel_w - 10.0,
                                bottom: gy + 18.0,
                            },
                            &text_brush,
                            D2D1_DRAW_TEXT_OPTIONS_NONE,
                            DWRITE_MEASURING_MODE_NATURAL,
                        );
                        gy += 16.0;
                    }
                }
            }

            if shell.goto_visible && !shell.command_palette_open && !shell.settings_visible {
                let dialog_w =
                    260.0_f32.min((canvas_rect.right - canvas_rect.left - 20.0).max(180.0));
                let dialog_h = 72.0;
                let dialog_x = (canvas_rect.right - dialog_w - 10.0).max(canvas_rect.left + 8.0);
                let dialog_y = canvas_rect.top + 10.0;
                let dialog = D2D_RECT_F {
                    left: dialog_x,
                    top: dialog_y,
                    right: dialog_x + dialog_w,
                    bottom: dialog_y + dialog_h,
                };
                let dialog_bg = self.create_brush(self.theme.surface_primary.as_d2d())?;
                let dialog_border = self.create_brush(self.theme.border_default.as_d2d())?;
                self.d2d_context.FillRectangle(&dialog, &dialog_bg);
                self.d2d_context.DrawRectangle(
                    &dialog,
                    &dialog_border,
                    1.0,
                    None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                );
                let title = "Go To (Ctrl+G)".encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &title,
                    &text_format,
                    &D2D_RECT_F {
                        left: dialog.left + 10.0,
                        top: dialog.top + 8.0,
                        right: dialog.right - 10.0,
                        bottom: dialog.top + 28.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );
                let input = format!("Line/Page: {}", shell.goto_input)
                    .encode_utf16()
                    .collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &input,
                    &text_format,
                    &D2D_RECT_F {
                        left: dialog.left + 10.0,
                        top: dialog.top + 30.0,
                        right: dialog.right - 10.0,
                        bottom: dialog.bottom - 10.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );
            }

            if shell.image_toolbar_visible {
                let toolbar_w =
                    520.0_f32.min((canvas_rect.right - canvas_rect.left - 20.0).max(320.0));
                let toolbar_h = 70.0;
                let toolbar_x = canvas_rect.left + 10.0;
                let toolbar_y = canvas_rect.top + 10.0;
                let toolbar = D2D_RECT_F {
                    left: toolbar_x,
                    top: toolbar_y,
                    right: toolbar_x + toolbar_w,
                    bottom: toolbar_y + toolbar_h,
                };
                let panel_bg = self.create_brush(self.theme.surface_primary.as_d2d())?;
                let panel_border = self.create_brush(self.theme.border_default.as_d2d())?;
                self.d2d_context.FillRectangle(&toolbar, &panel_bg);
                self.d2d_context.DrawRectangle(
                    &toolbar,
                    &panel_border,
                    1.0,
                    None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                );

                let actions = "Image Toolbar: Replace(Ctrl+R)  Delete(Del)  Align Left(Ctrl+L)  Center(Ctrl+E)  Right(Ctrl+I)  Border(Ctrl+Shift+B)";
                let actions_utf16 = actions.encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &actions_utf16,
                    &text_format,
                    &D2D_RECT_F {
                        left: toolbar.left + 10.0,
                        top: toolbar.top + 8.0,
                        right: toolbar.right - 10.0,
                        bottom: toolbar.top + 28.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                let meta = format!(
                    "{} | {}",
                    shell.image_selected_size, shell.image_selected_meta
                );
                let meta_utf16 = meta.encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &meta_utf16,
                    &text_format,
                    &D2D_RECT_F {
                        left: toolbar.left + 10.0,
                        top: toolbar.top + 30.0,
                        right: toolbar.right - 10.0,
                        bottom: toolbar.top + 50.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                let alt = format!("Alt text: {}", shell.image_selected_alt_text);
                let alt_utf16 = alt.encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &alt_utf16,
                    &text_format,
                    &D2D_RECT_F {
                        left: toolbar.left + 10.0,
                        top: toolbar.top + 48.0,
                        right: toolbar.right - 10.0,
                        bottom: toolbar.bottom - 6.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );
            }

            if shell.image_properties_visible {
                let props_w =
                    360.0_f32.min((canvas_rect.right - canvas_rect.left - 20.0).max(240.0));
                let props_h = 118.0;
                let props_x = canvas_rect.left + 12.0;
                let props_y = canvas_rect.top + 86.0;
                let props = D2D_RECT_F {
                    left: props_x,
                    top: props_y,
                    right: props_x + props_w,
                    bottom: props_y + props_h,
                };
                let panel_bg = self.create_brush(self.theme.surface_secondary.as_d2d())?;
                let panel_border = self.create_brush(self.theme.border_default.as_d2d())?;
                self.d2d_context.FillRectangle(&props, &panel_bg);
                self.d2d_context.DrawRectangle(
                    &props,
                    &panel_border,
                    1.0,
                    None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                );

                let line1 = "Image Properties".encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &line1,
                    &text_format,
                    &D2D_RECT_F {
                        left: props.left + 10.0,
                        top: props.top + 8.0,
                        right: props.right - 10.0,
                        bottom: props.top + 28.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                let line2 = format!("Size: {}", shell.image_selected_size)
                    .encode_utf16()
                    .collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &line2,
                    &text_format,
                    &D2D_RECT_F {
                        left: props.left + 10.0,
                        top: props.top + 30.0,
                        right: props.right - 10.0,
                        bottom: props.top + 50.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                let line3 = format!("Alignment / Border: {}", shell.image_selected_meta)
                    .encode_utf16()
                    .collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &line3,
                    &text_format,
                    &D2D_RECT_F {
                        left: props.left + 10.0,
                        top: props.top + 50.0,
                        right: props.right - 10.0,
                        bottom: props.top + 70.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                let line4 = format!(
                    "Wrap: Inline/Float (drag to move) | Alt: {}",
                    shell.image_selected_alt_text
                )
                .encode_utf16()
                .collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &line4,
                    &text_format,
                    &D2D_RECT_F {
                        left: props.left + 10.0,
                        top: props.top + 72.0,
                        right: props.right - 10.0,
                        bottom: props.bottom - 10.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );
            }

            if !shell.table_selected_meta.is_empty() {
                let panel_w =
                    700.0_f32.min((canvas_rect.right - canvas_rect.left - 20.0).max(320.0));
                let panel_h = 68.0;
                let panel_x = canvas_rect.left + 10.0;
                let panel_y = if shell.image_properties_visible {
                    canvas_rect.top + 210.0
                } else {
                    canvas_rect.top + 10.0
                };
                let panel = D2D_RECT_F {
                    left: panel_x,
                    top: panel_y,
                    right: panel_x + panel_w,
                    bottom: panel_y + panel_h,
                };
                let panel_bg = self.create_brush(self.theme.surface_secondary.as_d2d())?;
                let panel_border = self.create_brush(self.theme.border_default.as_d2d())?;
                self.d2d_context.FillRectangle(&panel, &panel_bg);
                self.d2d_context.DrawRectangle(
                    &panel,
                    &panel_border,
                    1.0,
                    None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                );

                let title = format!("Table #{}  {}", shell.table_selected_id, shell.table_selected_meta);
                let title_utf16 = title.encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &title_utf16,
                    &text_format,
                    &D2D_RECT_F {
                        left: panel.left + 10.0,
                        top: panel.top + 8.0,
                        right: panel.right - 10.0,
                        bottom: panel.top + 28.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                let shortcuts = "Tab/Shift+Tab Move  Shift+Arrows Expand  Ctrl+Shift+M Merge  Ctrl+Shift+Y Split  Ctrl+Shift+1..5 Style  Ctrl+Shift+U/J/H/K Insert Row/Col";
                let shortcuts_utf16 = shortcuts.encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &shortcuts_utf16,
                    &text_format,
                    &D2D_RECT_F {
                        left: panel.left + 10.0,
                        top: panel.top + 30.0,
                        right: panel.right - 10.0,
                        bottom: panel.bottom - 8.0,
                    },
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );
            }

            if status_h > 0.0 {
                let status_left_text = if shell.status_text.trim().is_empty() {
                    shell.status_left.clone()
                } else if shell.status_left.trim().is_empty()
                    || shell.status_left.eq_ignore_ascii_case(&shell.status_text)
                {
                    shell.status_text.clone()
                } else {
                    format!("{} | {}", shell.status_text, shell.status_left)
                };
                let status = status_left_text.encode_utf16().collect::<Vec<u16>>();
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
            }

            if tab_h > 0.0 {
                let mut tabs_left = 8.0;
                let tabs_top = 4.0;
                let tabs_bottom = tab_h - 4.0;
                let new_btn_rect = D2D_RECT_F {
                    left: (width - 36.0).max(8.0),
                    top: 6.0,
                    right: (width - 8.0).max(20.0),
                    bottom: tab_h - 6.0,
                };

                if shell.tab_has_overflow_left || shell.tab_has_overflow_right {
                    let overflow_bg = self.create_brush(self.theme.surface_secondary.as_d2d())?;
                    let overflow_text = self.create_brush(self.theme.text_secondary.as_d2d())?;
                    let left_rect = D2D_RECT_F {
                        left: tabs_left,
                        top: 6.0,
                        right: tabs_left + 24.0,
                        bottom: tab_h - 6.0,
                    };
                    let right_rect = D2D_RECT_F {
                        left: tabs_left + 28.0,
                        top: 6.0,
                        right: tabs_left + 52.0,
                        bottom: tab_h - 6.0,
                    };
                    self.d2d_context.FillRectangle(&left_rect, &overflow_bg);
                    self.d2d_context.FillRectangle(&right_rect, &overflow_bg);
                    let left_arrow = if shell.tab_has_overflow_left { "<" } else { " " }
                        .encode_utf16()
                        .collect::<Vec<u16>>();
                    let right_arrow = if shell.tab_has_overflow_right { ">" } else { " " }
                        .encode_utf16()
                        .collect::<Vec<u16>>();
                    self.d2d_context.DrawText(
                        &left_arrow,
                        &text_format,
                        &left_rect,
                        &overflow_text,
                        D2D1_DRAW_TEXT_OPTIONS_NONE,
                        DWRITE_MEASURING_MODE_NATURAL,
                    );
                    self.d2d_context.DrawText(
                        &right_arrow,
                        &text_format,
                        &right_rect,
                        &overflow_text,
                        D2D1_DRAW_TEXT_OPTIONS_NONE,
                        DWRITE_MEASURING_MODE_NATURAL,
                    );
                    tabs_left = right_rect.right + 6.0;
                }

                let new_btn_bg = self.create_brush(self.theme.surface_secondary.as_d2d())?;
                self.d2d_context.FillRectangle(&new_btn_rect, &new_btn_bg);
                let plus = "+".encode_utf16().collect::<Vec<u16>>();
                self.d2d_context.DrawText(
                    &plus,
                    &text_format,
                    &new_btn_rect,
                    &text_brush,
                    D2D1_DRAW_TEXT_OPTIONS_NONE,
                    DWRITE_MEASURING_MODE_NATURAL,
                );

                if !shell.tab_titles.is_empty() {
                    let tabs_right = (new_btn_rect.left - 6.0).max(tabs_left + 100.0);
                    let count = shell.tab_titles.len();
                    let gap = 6.0;
                    let total_gap = gap * count.saturating_sub(1) as f32;
                    let tab_w =
                        ((tabs_right - tabs_left - total_gap) / count as f32).clamp(140.0, 260.0);
                    let accent_brush = self.create_brush(self.theme.accent.as_d2d())?;
                    let close_brush = self.create_brush(self.theme.text_secondary.as_d2d())?;
                    let mut x = tabs_left;
                    for (idx, title) in shell.tab_titles.iter().enumerate() {
                        let mut rect = D2D_RECT_F {
                            left: x,
                            top: tabs_top,
                            right: (x + tab_w).min(tabs_right),
                            bottom: tabs_bottom,
                        };
                        if idx == shell.active_tab {
                            rect.left += shell.tab_transition_offset;
                            rect.right += shell.tab_transition_offset;
                        }
                        let brush = if idx == shell.active_tab {
                            let mut active = self.theme.surface_hover.as_d2d();
                            active.a *= shell.tab_transition_progress.clamp(0.0, 1.0).max(0.65);
                            self.create_brush(active)?
                        } else {
                            self.create_brush(self.theme.surface_secondary.as_d2d())?
                        };
                        self.d2d_context.FillRectangle(&rect, &brush);
                        if idx == shell.active_tab {
                            let accent_line = D2D_RECT_F {
                                left: rect.left + 1.0,
                                top: rect.bottom - 3.0,
                                right: rect.right - 1.0,
                                bottom: rect.bottom - 1.0,
                            };
                            self.d2d_context.FillRectangle(&accent_line, &accent_brush);
                        }

                        let text = title.encode_utf16().collect::<Vec<u16>>();
                        self.d2d_context.DrawText(
                            &text,
                            &text_format,
                            &D2D_RECT_F {
                                left: rect.left + 10.0,
                                top: rect.top + 6.0,
                                right: rect.right - 18.0,
                                bottom: rect.bottom - 4.0,
                            },
                            &text_brush,
                            D2D1_DRAW_TEXT_OPTIONS_NONE,
                            DWRITE_MEASURING_MODE_NATURAL,
                        );

                        let close = "x".encode_utf16().collect::<Vec<u16>>();
                        self.d2d_context.DrawText(
                            &close,
                            &text_format,
                            &D2D_RECT_F {
                                left: rect.right - 16.0,
                                top: rect.top + 7.0,
                                right: rect.right - 4.0,
                                bottom: rect.bottom - 5.0,
                            },
                            &close_brush,
                            D2D1_DRAW_TEXT_OPTIONS_NONE,
                            DWRITE_MEASURING_MODE_NATURAL,
                        );

                        x += tab_w + gap;
                        if x + 80.0 > tabs_right {
                            break;
                        }
                    }
                }
            }

            if toolbar_h > 0.0 && !shell.toolbar_buttons.is_empty() {
                let button_bg = self.create_brush(self.theme.surface_secondary.as_d2d())?;
                let button_active_bg = self.create_brush(self.theme.surface_hover.as_d2d())?;
                let button_disabled_bg = self.create_brush(self.theme.surface_primary.as_d2d())?;
                let button_border = self.create_brush(self.theme.border_default.as_d2d())?;
                let text_disabled = self.create_brush(self.theme.text_secondary.as_d2d())?;
                let separator_brush = self.create_brush(self.theme.border_subtle.as_d2d())?;
                let icon_format = self.create_icon_text_format()?;
                for button in shell.toolbar_buttons.iter() {
                    let rect = D2D_RECT_F {
                        left: button.rect.x,
                        top: button.rect.y,
                        right: button.rect.x + button.rect.width,
                        bottom: button.rect.y + button.rect.height,
                    };

                    if button.label.trim().is_empty() && button.icon.trim().is_empty() && button.rect.width <= 12.0
                    {
                        let center = button.rect.x + (button.rect.width * 0.5);
                        self.d2d_context.DrawLine(
                            Vector2 {
                                X: center,
                                Y: rect.top + 4.0,
                            },
                            Vector2 {
                                X: center,
                                Y: rect.bottom - 4.0,
                            },
                            &separator_brush,
                            1.0,
                            None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                        );
                        continue;
                    }

                    let bg = if !button.enabled {
                        &button_disabled_bg
                    } else if button.active {
                        &button_active_bg
                    } else {
                        &button_bg
                    };
                    self.d2d_context.FillRectangle(&rect, bg);
                    self.d2d_context.DrawRectangle(
                        &rect,
                        &button_border,
                        1.0,
                        None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                    );

                    let is_icon_only = button.label.trim().is_empty() && !button.icon.trim().is_empty();
                    let caption = if is_icon_only {
                        button.icon.as_str()
                    } else if button.label.trim().is_empty() {
                        ""
                    } else {
                        button.label.as_str()
                    };
                    if caption.is_empty() {
                        continue;
                    }
                    let t = caption.encode_utf16().collect::<Vec<u16>>();
                    self.d2d_context.DrawText(
                        &t,
                        if is_icon_only { &icon_format } else { &text_format },
                        &D2D_RECT_F {
                            left: rect.left + 6.0,
                            top: rect.top + if is_icon_only { 2.0 } else { 4.0 },
                            right: rect.right - 6.0,
                            bottom: rect.bottom - 4.0,
                        },
                        if button.enabled { &text_brush } else { &text_disabled },
                        D2D1_DRAW_TEXT_OPTIONS_CLIP,
                        DWRITE_MEASURING_MODE_NATURAL,
                    );
                }
            }

            if shell.toolbar_dropdown_open && shell.toolbar_dropdown_opacity > 0.01 {
                let panel_w = 240.0 * shell.toolbar_dropdown_scale.clamp(0.9, 1.2);
                let panel_h = 180.0 * shell.toolbar_dropdown_scale.clamp(0.9, 1.2);
                let panel_x = sidebar_w + 20.0;
                let panel_y = tab_h + toolbar_h + 8.0;
                let panel_rect = D2D_RECT_F {
                    left: panel_x,
                    top: panel_y,
                    right: panel_x + panel_w,
                    bottom: panel_y + panel_h,
                };
                let mut panel_color = self.theme.surface_primary.as_d2d();
                panel_color.a = 0.95 * shell.toolbar_dropdown_opacity.clamp(0.0, 1.0);
                let panel_bg = self.create_brush(panel_color)?;
                let mut panel_border = self.theme.border_default.as_d2d();
                panel_border.a *= shell.toolbar_dropdown_opacity.clamp(0.0, 1.0);
                let panel_border_brush = self.create_brush(panel_border)?;
                self.d2d_context.FillRectangle(&panel_rect, &panel_bg);
                self.d2d_context.DrawRectangle(
                    &panel_rect,
                    &panel_border_brush,
                    1.0,
                    None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                );
            }

            if !shell.toast_entries.is_empty() {
                for (idx, entry) in shell.toast_entries.iter().enumerate().take(4) {
                    let width_toast = 320.0;
                    let height_toast = 64.0;
                    let right = width - 14.0 + entry.slide_offset;
                    let left = right - width_toast;
                    let bottom = (height - status_h - 14.0) - (idx as f32 * 74.0);
                    let top = bottom - height_toast;
                    let rect = D2D_RECT_F {
                        left,
                        top,
                        right,
                        bottom,
                    };

                    let mut bg = self.theme.surface_primary.as_d2d();
                    bg.a = 0.94 * entry.opacity.clamp(0.0, 1.0);
                    let mut border = self.theme.border_default.as_d2d();
                    border.a *= entry.opacity.clamp(0.0, 1.0);
                    let bg_brush = self.create_brush(bg)?;
                    let border_brush = self.create_brush(border)?;
                    self.d2d_context.FillRectangle(&rect, &bg_brush);
                    self.d2d_context.DrawRectangle(
                        &rect,
                        &border_brush,
                        1.0,
                        None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                    );

                    let title = entry.title.encode_utf16().collect::<Vec<u16>>();
                    self.d2d_context.DrawText(
                        &title,
                        &text_format,
                        &D2D_RECT_F {
                            left: rect.left + 10.0,
                            top: rect.top + 8.0,
                            right: rect.right - 10.0,
                            bottom: rect.top + 28.0,
                        },
                        &text_brush,
                        D2D1_DRAW_TEXT_OPTIONS_NONE,
                        DWRITE_MEASURING_MODE_NATURAL,
                    );
                    let body = entry.body.encode_utf16().collect::<Vec<u16>>();
                    self.d2d_context.DrawText(
                        &body,
                        &text_format,
                        &D2D_RECT_F {
                            left: rect.left + 10.0,
                            top: rect.top + 28.0,
                            right: rect.right - 10.0,
                            bottom: rect.bottom - 8.0,
                        },
                        &text_brush,
                        D2D1_DRAW_TEXT_OPTIONS_NONE,
                        DWRITE_MEASURING_MODE_NATURAL,
                    );
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

    fn draw_canvas_background(
        &self,
        rect: D2D_RECT_F,
        settings: &BackgroundSettings,
    ) -> Result<()> {
        match &settings.kind {
            BackgroundKind::Solid { color } => self.fill_rect(rect, *color),
            BackgroundKind::Gradient {
                start,
                end,
                angle_degrees,
            } => self.fill_gradient(rect, *start, *end, *angle_degrees),
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
                let overlay =
                    self.create_brush(windows::Win32::Graphics::Direct2D::Common::D2D1_COLOR_F {
                        r: 0.0,
                        g: 0.0,
                        b: 0.0,
                        a: (1.0 - opacity.clamp(0.0, 1.0)).clamp(0.0, 1.0) * 0.35,
                    })?;
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

    fn draw_document_canvas(
        &self,
        canvas_rect: D2D_RECT_F,
        shell: &ShellRenderState,
    ) -> Result<()> {
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
                    self.draw_page_preview_content(page_rect, canvas_rect, shell)?;
                    drew_preview = true;
                }
            }
        }

        self.draw_canvas_scrollbars(canvas_rect, shell)
    }

    fn draw_page_preview_content(
        &self,
        page_rect: D2D_RECT_F,
        canvas_rect: D2D_RECT_F,
        shell: &ShellRenderState,
    ) -> Result<()> {
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
            self.d2d_context
                .FillRectangle(&current_line, &line_highlight);
            let selection_rect = D2D_RECT_F {
                left: text_rect.left + 2.0,
                top: text_rect.top + 3.0,
                right: (text_rect.left + 220.0).min(text_rect.right),
                bottom: text_rect.top + 22.0,
            };
            self.d2d_context.FillRectangle(&selection_rect, &selection);
        }

        if shell.find_visible && shell.find_total > 0 && !shell.settings_visible {
            let all_match_brush = self.create_brush(
                crate::ui::Color::rgba(
                    self.theme.accent.r,
                    self.theme.accent.g,
                    self.theme.accent.b,
                    0.22,
                )
                .as_d2d(),
            )?;
            let current_match_brush = self.create_brush(
                crate::ui::Color::rgba(
                    self.theme.accent.r,
                    self.theme.accent.g,
                    self.theme.accent.b,
                    0.46,
                )
                .as_d2d(),
            )?;

            let markers = shell.find_total.min(24);
            if markers > 0 {
                let marker_w = ((text_rect.right - text_rect.left) / markers as f32).max(6.0);
                for idx in 0..markers {
                    let left = text_rect.left + idx as f32 * marker_w;
                    let rect = D2D_RECT_F {
                        left,
                        top: text_rect.top + 26.0,
                        right: (left + marker_w - 2.0).min(text_rect.right),
                        bottom: text_rect.top + 34.0,
                    };
                    let brush =
                        if shell.find_current > 0 && idx + 1 == shell.find_current.min(markers) {
                            &current_match_brush
                        } else {
                            &all_match_brush
                        };
                    unsafe {
                        self.d2d_context.FillRectangle(&rect, brush);
                    }
                }
            }
        }

        if !shell.canvas_images.is_empty() {
            let image_bg = self.create_brush(self.theme.surface_secondary.as_d2d())?;
            let image_border = self.create_brush(self.theme.border_default.as_d2d())?;
            let image_selected = self.create_brush(self.theme.accent.as_d2d())?;
            let image_text = self.create_brush(self.theme.text_secondary.as_d2d())?;
            let handle_brush = self.create_brush(self.theme.accent.as_d2d())?;

            for image in shell.canvas_images.iter().take(12) {
                let left = canvas_rect.left + image.rect.x;
                let top = canvas_rect.top + image.rect.y;
                let right = left + image.rect.width;
                let bottom = top + image.rect.height;
                let img_rect = D2D_RECT_F {
                    left,
                    top,
                    right,
                    bottom,
                };

                if right < page_rect.left
                    || left > page_rect.right
                    || bottom < page_rect.top
                    || top > page_rect.bottom
                {
                    continue;
                }

                unsafe {
                    self.d2d_context.FillRectangle(&img_rect, &image_bg);
                    self.d2d_context.DrawRectangle(
                        &img_rect,
                        if image.selected {
                            &image_selected
                        } else {
                            &image_border
                        },
                        if image.selected { 2.0 } else { 1.0 },
                        None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                    );
                }

                let label = if image.alt_text.is_empty() {
                    format!("[Image #{}]", image.block_id)
                } else {
                    format!("[Image #{}] {}", image.block_id, image.alt_text)
                };
                let interpolation = image.interpolation.encode_utf16().collect::<Vec<u16>>();
                let label_utf16 = label.encode_utf16().collect::<Vec<u16>>();
                unsafe {
                    self.d2d_context.DrawText(
                        &label_utf16,
                        &self.create_text_format()?,
                        &D2D_RECT_F {
                            left: left + 8.0,
                            top: top + 6.0,
                            right: right - 8.0,
                            bottom: top + 24.0,
                        },
                        &image_text,
                        D2D1_DRAW_TEXT_OPTIONS_NONE,
                        DWRITE_MEASURING_MODE_NATURAL,
                    );
                    self.d2d_context.DrawText(
                        &interpolation,
                        &self.create_text_format()?,
                        &D2D_RECT_F {
                            left: left + 8.0,
                            top: bottom - 20.0,
                            right: right - 8.0,
                            bottom: bottom - 4.0,
                        },
                        &image_text,
                        D2D1_DRAW_TEXT_OPTIONS_NONE,
                        DWRITE_MEASURING_MODE_NATURAL,
                    );
                }

                if image.selected {
                    let handle = 6.0;
                    let handle_rects = [
                        D2D_RECT_F {
                            left: left - handle * 0.5,
                            top: top - handle * 0.5,
                            right: left + handle * 0.5,
                            bottom: top + handle * 0.5,
                        },
                        D2D_RECT_F {
                            left: right - handle * 0.5,
                            top: top - handle * 0.5,
                            right: right + handle * 0.5,
                            bottom: top + handle * 0.5,
                        },
                        D2D_RECT_F {
                            left: left - handle * 0.5,
                            top: bottom - handle * 0.5,
                            right: left + handle * 0.5,
                            bottom: bottom + handle * 0.5,
                        },
                        D2D_RECT_F {
                            left: right - handle * 0.5,
                            top: bottom - handle * 0.5,
                            right: right + handle * 0.5,
                            bottom: bottom + handle * 0.5,
                        },
                    ];
                    for rect in handle_rects {
                        unsafe {
                            self.d2d_context.FillRectangle(&rect, &handle_brush);
                        }
                    }
                }
            }
        }

        if !shell.canvas_tables.is_empty() {
            let table_bg = self.create_brush(self.theme.surface_secondary.as_d2d())?;
            let table_border = self.create_brush(self.theme.border_default.as_d2d())?;
            let table_header = self.create_brush(
                crate::ui::Color::rgba(
                    self.theme.surface_hover.r,
                    self.theme.surface_hover.g,
                    self.theme.surface_hover.b,
                    0.8,
                )
                .as_d2d(),
            )?;
            let table_selected = self.create_brush(self.theme.accent.as_d2d())?;
            let table_text = self.create_brush(self.theme.text_secondary.as_d2d())?;
            let selection_fill = self.create_brush(
                crate::ui::Color::rgba(
                    self.theme.selection_bg.r,
                    self.theme.selection_bg.g,
                    self.theme.selection_bg.b,
                    0.32,
                )
                .as_d2d(),
            )?;

            for table in shell.canvas_tables.iter().take(10) {
                let left = canvas_rect.left + table.rect.x;
                let top = canvas_rect.top + table.rect.y;
                let right = left + table.rect.width;
                let bottom = top + table.rect.height;
                let table_rect = D2D_RECT_F {
                    left,
                    top,
                    right,
                    bottom,
                };

                if right < page_rect.left
                    || left > page_rect.right
                    || bottom < page_rect.top
                    || top > page_rect.bottom
                {
                    continue;
                }

                unsafe {
                    self.d2d_context.FillRectangle(&table_rect, &table_bg);
                    self.d2d_context.DrawRectangle(
                        &table_rect,
                        if table.selected {
                            &table_selected
                        } else {
                            &table_border
                        },
                        if table.selected { 2.0 } else { 1.0 },
                        None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                    );
                }

                let header_rect = D2D_RECT_F {
                    left: left + table.gutter_w,
                    top,
                    right,
                    bottom: top + table.header_h,
                };
                let gutter_rect = D2D_RECT_F {
                    left,
                    top: top + table.header_h,
                    right: left + table.gutter_w,
                    bottom,
                };
                unsafe {
                    self.d2d_context.FillRectangle(&header_rect, &table_header);
                    self.d2d_context.FillRectangle(&gutter_rect, &table_header);
                }

                for r in 0..table.rows {
                    for c in 0..table.cols {
                        let cell_left = left + table.gutter_w + c as f32 * table.cell_w;
                        let cell_top = top + table.header_h + r as f32 * table.cell_h;
                        let cell_rect = D2D_RECT_F {
                            left: cell_left,
                            top: cell_top,
                            right: cell_left + table.cell_w,
                            bottom: cell_top + table.cell_h,
                        };

                        let in_selection = if table.selection_mode == 4 {
                            true
                        } else if table.selection_mode == 2 {
                            r >= table.selection_start_row && r <= table.selection_end_row
                        } else if table.selection_mode == 3 {
                            c >= table.selection_start_col && c <= table.selection_end_col
                        } else if table.selection_mode == 1 {
                            r >= table.selection_start_row
                                && r <= table.selection_end_row
                                && c >= table.selection_start_col
                                && c <= table.selection_end_col
                        } else {
                            false
                        };

                        unsafe {
                            if in_selection {
                                self.d2d_context.FillRectangle(&cell_rect, &selection_fill);
                            }
                            self.d2d_context.DrawRectangle(
                                &cell_rect,
                                &table_border,
                                1.0,
                                None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                            );
                        }
                    }
                }

                let label = format!("[Table #{}] {}x{}", table.table_id, table.rows, table.cols);
                let label_utf16 = label.encode_utf16().collect::<Vec<u16>>();
                unsafe {
                    self.d2d_context.DrawText(
                        &label_utf16,
                        &self.create_text_format()?,
                        &D2D_RECT_F {
                            left: left + table.gutter_w + 4.0,
                            top: top + 2.0,
                            right: right - 6.0,
                            bottom: top + table.header_h - 2.0,
                        },
                        &table_text,
                        D2D1_DRAW_TEXT_OPTIONS_NONE,
                        DWRITE_MEASURING_MODE_NATURAL,
                    );
                }
            }
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

    fn draw_canvas_scrollbars(
        &self,
        canvas_rect: D2D_RECT_F,
        shell: &ShellRenderState,
    ) -> Result<()> {
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
                            Vector2 {
                                X: rect.right,
                                Y: y,
                            },
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
                            Vector2 {
                                X: x,
                                Y: rect.bottom,
                            },
                            &brush,
                            1.0,
                            None::<&windows::Win32::Graphics::Direct2D::ID2D1StrokeStyle>,
                        );
                        x += step;
                    }
                }
                PatternStyle::LinesDiagonal
                | PatternStyle::CrossHatch
                | PatternStyle::GraphPaper => {
                    if matches!(style, PatternStyle::GraphPaper) {
                        let mut y = rect.top;
                        while y <= rect.bottom {
                            self.d2d_context.DrawLine(
                                Vector2 { X: rect.left, Y: y },
                                Vector2 {
                                    X: rect.right,
                                    Y: y,
                                },
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
                                Vector2 {
                                    X: x,
                                    Y: rect.bottom,
                                },
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
                            let hash =
                                ((xi.wrapping_mul(73856093)) ^ (yi.wrapping_mul(19349663))) & 0xFF;
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
        if let Some(existing) = self.default_text_format.borrow().as_ref() {
            return Ok(existing.clone());
        }

        unsafe {
            let format = match self.dwrite_factory.CreateTextFormat(
                w!("Segoe UI Variable"),
                None,
                windows::Win32::Graphics::DirectWrite::DWRITE_FONT_WEIGHT_NORMAL,
                windows::Win32::Graphics::DirectWrite::DWRITE_FONT_STYLE_NORMAL,
                windows::Win32::Graphics::DirectWrite::DWRITE_FONT_STRETCH_NORMAL,
                14.0,
                w!("en-US"),
            ) {
                Ok(format) => format,
                Err(_) => self.dwrite_factory.CreateTextFormat(
                    w!("Segoe UI"),
                    None,
                    windows::Win32::Graphics::DirectWrite::DWRITE_FONT_WEIGHT_NORMAL,
                    windows::Win32::Graphics::DirectWrite::DWRITE_FONT_STYLE_NORMAL,
                    windows::Win32::Graphics::DirectWrite::DWRITE_FONT_STRETCH_NORMAL,
                    14.0,
                    w!("en-US"),
                )?,
            };
            let _ = format.SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);

            *self.default_text_format.borrow_mut() = Some(format.clone());
            Ok(format)
        }
    }

    fn create_icon_text_format(&self) -> Result<IDWriteTextFormat> {
        if let Some(existing) = self.icon_text_format.borrow().as_ref() {
            return Ok(existing.clone());
        }

        unsafe {
            let format = match self.dwrite_factory.CreateTextFormat(
                w!("Segoe Fluent Icons"),
                None,
                windows::Win32::Graphics::DirectWrite::DWRITE_FONT_WEIGHT_NORMAL,
                windows::Win32::Graphics::DirectWrite::DWRITE_FONT_STYLE_NORMAL,
                windows::Win32::Graphics::DirectWrite::DWRITE_FONT_STRETCH_NORMAL,
                13.0,
                w!("en-US"),
            ) {
                Ok(format) => format,
                Err(_) => match self.dwrite_factory.CreateTextFormat(
                    w!("Segoe MDL2 Assets"),
                    None,
                    windows::Win32::Graphics::DirectWrite::DWRITE_FONT_WEIGHT_NORMAL,
                    windows::Win32::Graphics::DirectWrite::DWRITE_FONT_STYLE_NORMAL,
                    windows::Win32::Graphics::DirectWrite::DWRITE_FONT_STRETCH_NORMAL,
                    13.0,
                    w!("en-US"),
                ) {
                    Ok(format) => format,
                    Err(_) => self.dwrite_factory.CreateTextFormat(
                        w!("Segoe UI Symbol"),
                        None,
                        windows::Win32::Graphics::DirectWrite::DWRITE_FONT_WEIGHT_NORMAL,
                        windows::Win32::Graphics::DirectWrite::DWRITE_FONT_STYLE_NORMAL,
                        windows::Win32::Graphics::DirectWrite::DWRITE_FONT_STRETCH_NORMAL,
                        13.0,
                        w!("en-US"),
                    )?,
                },
            };
            let _ = format.SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
            let _ = format.SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
            let _ = format.SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            *self.icon_text_format.borrow_mut() = Some(format.clone());
            Ok(format)
        }
    }

    fn create_brush(
        &self,
        color: windows::Win32::Graphics::Direct2D::Common::D2D1_COLOR_F,
    ) -> Result<ID2D1SolidColorBrush> {
        let key = Self::brush_color_key(color);
        if let Some(existing) = self.brush_cache.borrow().get(&key).cloned() {
            return Ok(existing);
        }

        let brush = unsafe { self.d2d_context.CreateSolidColorBrush(&color, None)? };
        self.brush_cache.borrow_mut().insert(key, brush.clone());
        Ok(brush)
    }

    fn brush_color_key(color: windows::Win32::Graphics::Direct2D::Common::D2D1_COLOR_F) -> u32 {
        let quantize = |v: f32| -> u32 { (v.clamp(0.0, 1.0) * 255.0).round() as u32 };
        let r = quantize(color.r);
        let g = quantize(color.g);
        let b = quantize(color.b);
        let a = quantize(color.a);
        (a << 24) | (r << 16) | (g << 8) | b
    }

    unsafe fn recreate_target_bitmap(&mut self) -> Result<()> {
        let surface: IDXGISurface = unsafe { self.swap_chain.GetBuffer(0)? };
        let bitmap_props = D2D1_BITMAP_PROPERTIES1 {
            pixelFormat: D2D1_PIXEL_FORMAT {
                format: DXGI_FORMAT_B8G8R8A8_UNORM,
                alphaMode: D2D1_ALPHA_MODE_IGNORE,
            },
            dpiX: LAYOUT_DPI,
            dpiY: LAYOUT_DPI,
            bitmapOptions: D2D1_BITMAP_OPTIONS_TARGET | D2D1_BITMAP_OPTIONS_CANNOT_DRAW,
            colorContext: ManuallyDrop::new(None),
        };

        let bitmap = unsafe {
            self.d2d_context
                .CreateBitmapFromDxgiSurface(&surface, Some(&bitmap_props))?
        };

        unsafe {
            self.d2d_context.SetTarget(&bitmap);
            let _ = self.d2d_context.SetDpi(LAYOUT_DPI, LAYOUT_DPI);
        }
        self.target_bitmap = Some(bitmap);
        self.brush_cache.borrow_mut().clear();

        Ok(())
    }
}
