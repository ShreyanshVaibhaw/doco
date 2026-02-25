use std::{mem::size_of, ptr::copy_nonoverlapping, sync::OnceLock};

use encoding_rs::WINDOWS_1252;
use image::GenericImageView;
use serde::{Deserialize, Serialize};
use windows::{
    Win32::{
        Foundation::{HANDLE, HGLOBAL, HWND},
        System::{
            DataExchange::{
                CloseClipboard, EmptyClipboard, GetClipboardData, IsClipboardFormatAvailable,
                OpenClipboard, RegisterClipboardFormatW, SetClipboardData,
            },
            Memory::{GMEM_MOVEABLE, GlobalAlloc, GlobalLock, GlobalSize, GlobalUnlock},
        },
    },
    core::{Error, Result, w},
};

use crate::{
    document::{
        DocumentFormat,
        model::{Run, RunStyle},
    },
    editor::{
        commands::EditCommand,
        cursor::{CursorPosition, SelectionRange},
    },
};

const CF_UNICODETEXT_U32: u32 = 13;
const CF_DIB_U32: u32 = 8;

static INTERNAL_CLIPBOARD_FORMAT: OnceLock<u32> = OnceLock::new();
static RTF_CLIPBOARD_FORMAT: OnceLock<u32> = OnceLock::new();
static HTML_CLIPBOARD_FORMAT: OnceLock<u32> = OnceLock::new();
static PNG_CLIPBOARD_FORMAT: OnceLock<u32> = OnceLock::new();

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PasteMode {
    RichText,
    PlainText,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ClipboardSource {
    Internal,
    Rtf,
    Html,
    UnicodeText,
}

#[derive(Debug, Clone)]
pub struct ClipboardPastePayload {
    pub source: ClipboardSource,
    pub runs: Vec<Run>,
}

impl ClipboardPastePayload {
    pub fn plain_text(&self) -> String {
        runs_to_plain_text(&self.runs)
    }
}

#[derive(Debug, Clone)]
pub struct ClipboardImageData {
    pub bytes: Vec<u8>,
    pub mime: String,
    pub width: u32,
    pub height: u32,
}

#[derive(Debug, Clone)]
pub struct TextDragSession {
    pub selection: SelectionRange,
    pub selected_text: String,
    pub copy_mode: bool,
    pub insertion_indicator: Option<CursorPosition>,
}

impl TextDragSession {
    pub fn new(selection: SelectionRange, selected_text: String, copy_mode: bool) -> Self {
        Self {
            selection,
            selected_text,
            copy_mode,
            insertion_indicator: None,
        }
    }

    pub fn update_insertion_indicator(&mut self, pos: CursorPosition) {
        self.insertion_indicator = Some(pos);
    }

    pub fn clear_insertion_indicator(&mut self) {
        self.insertion_indicator = None;
    }

    pub fn build_drop_commands(&self) -> Vec<EditCommand> {
        let Some(drop_at) = self.insertion_indicator else {
            return Vec::new();
        };
        drag_drop_commands(self.selection, drop_at, &self.selected_text, self.copy_mode)
    }
}

pub fn set_plain_text(text: &str) -> Result<()> {
    let _guard = ClipboardGuard::open()?;
    unsafe {
        EmptyClipboard()?;
    }
    set_clipboard_unicode_text(text)
}

pub fn get_plain_text() -> Result<Option<String>> {
    let _guard = ClipboardGuard::open()?;
    Ok(get_clipboard_unicode_text())
}

pub fn copy_runs_to_clipboard(runs: &[Run], source_format: DocumentFormat) -> Result<()> {
    let plain_text = runs_to_plain_text(runs);
    let include_rich = should_emit_rich_formats(runs, source_format);

    let _guard = ClipboardGuard::open()?;
    unsafe {
        EmptyClipboard()?;
    }
    set_clipboard_unicode_text(&plain_text)?;

    if include_rich {
        if let Ok(payload) = encode_internal_payload(runs, &plain_text) {
            set_clipboard_bytes(internal_clipboard_format(), &payload)?;
        }

        let rtf = runs_to_minimal_rtf(runs);
        if !rtf.is_empty() {
            set_clipboard_bytes(rtf_clipboard_format(), rtf.as_bytes())?;
        }

        let html = runs_to_cf_html(runs);
        if !html.is_empty() {
            set_clipboard_bytes(html_clipboard_format(), html.as_bytes())?;
        }
    }

    Ok(())
}

pub fn read_clipboard_for_paste(mode: PasteMode) -> Result<Option<ClipboardPastePayload>> {
    let _guard = ClipboardGuard::open()?;

    if matches!(mode, PasteMode::PlainText) {
        return Ok(get_clipboard_unicode_text().map(|t| ClipboardPastePayload {
            source: ClipboardSource::UnicodeText,
            runs: vec![Run {
                text: t,
                style: RunStyle::default(),
            }],
        }));
    }

    let mut order = vec![
        ClipboardSource::Internal,
        ClipboardSource::Rtf,
        ClipboardSource::Html,
        ClipboardSource::UnicodeText,
    ];
    let preferred = preferred_source(AvailableFormats {
        internal: is_format_available(internal_clipboard_format()),
        rtf: is_format_available(rtf_clipboard_format()),
        html: is_format_available(html_clipboard_format()),
        unicode: is_format_available(CF_UNICODETEXT_U32),
    });
    if let Some(primary) = preferred {
        order.retain(|v| *v != primary);
        order.insert(0, primary);
    }

    for source in order {
        match source {
            ClipboardSource::Internal => {
                if let Some(bytes) = get_clipboard_bytes(internal_clipboard_format())? {
                    if let Some(runs) = decode_internal_payload(&bytes) {
                        return Ok(Some(ClipboardPastePayload {
                            source: ClipboardSource::Internal,
                            runs,
                        }));
                    }
                }
            }
            ClipboardSource::Rtf => {
                if let Some(bytes) = get_clipboard_bytes(rtf_clipboard_format())? {
                    if let Ok(rtf) = String::from_utf8(bytes) {
                        let runs = parse_rtf_to_runs(&rtf);
                        if !runs.is_empty() {
                            return Ok(Some(ClipboardPastePayload {
                                source: ClipboardSource::Rtf,
                                runs,
                            }));
                        }
                    }
                }
            }
            ClipboardSource::Html => {
                if let Some(bytes) = get_clipboard_bytes(html_clipboard_format())? {
                    if let Ok(raw) = String::from_utf8(bytes) {
                        let runs = parse_html_to_runs(&raw);
                        if !runs.is_empty() {
                            return Ok(Some(ClipboardPastePayload {
                                source: ClipboardSource::Html,
                                runs,
                            }));
                        }
                    }
                }
            }
            ClipboardSource::UnicodeText => {
                if let Some(text) = get_clipboard_unicode_text() {
                    return Ok(Some(ClipboardPastePayload {
                        source: ClipboardSource::UnicodeText,
                        runs: vec![Run {
                            text,
                            style: RunStyle::default(),
                        }],
                    }));
                }
            }
        }
    }

    Ok(None)
}

pub fn read_clipboard_image() -> Result<Option<ClipboardImageData>> {
    let _guard = ClipboardGuard::open()?;

    if let Some(bytes) = get_clipboard_bytes(png_clipboard_format())? {
        if let Some(decoded) = decode_clipboard_image(bytes, "image/png") {
            return Ok(Some(decoded));
        }
    }

    if let Some(dib_bytes) = get_clipboard_bytes(CF_DIB_U32)? {
        if let Some(bmp_bytes) = dib_to_bmp_bytes(dib_bytes.as_slice()) {
            if let Some(decoded) = decode_clipboard_image(bmp_bytes, "image/bmp") {
                return Ok(Some(decoded));
            }
        }
    }

    Ok(None)
}

pub fn drag_drop_commands(
    selection: SelectionRange,
    drop_at: CursorPosition,
    selected_text: &str,
    copy_mode: bool,
) -> Vec<EditCommand> {
    if selected_text.is_empty() {
        return Vec::new();
    }

    let normalized = selection.normalized();
    let same_source_block = normalized.start.block_id == normalized.end.block_id;
    let same_drop_block = normalized.start.block_id == drop_at.block_id;

    let mut commands = Vec::with_capacity(if copy_mode { 1 } else { 2 });

    if !copy_mode
        && same_source_block
        && same_drop_block
        && drop_at.offset >= normalized.start.offset
        && drop_at.offset <= normalized.end.offset
    {
        return Vec::new();
    }

    let removed_len = normalized
        .end
        .offset
        .saturating_sub(normalized.start.offset);
    let insert_offset = if !copy_mode
        && same_source_block
        && same_drop_block
        && drop_at.offset > normalized.end.offset
    {
        drop_at.offset.saturating_sub(removed_len)
    } else {
        drop_at.offset
    };

    if !copy_mode && same_source_block {
        commands.push(EditCommand::DeleteText {
            block_id: normalized.start.block_id,
            start: normalized.start.offset,
            end: normalized.end.offset,
        });
    }

    commands.push(EditCommand::InsertText {
        block_id: drop_at.block_id,
        offset: insert_offset,
        text: selected_text.to_string(),
    });
    commands
}

pub fn copy_to_clipboard(text: &str) {
    let _ = set_plain_text(text);
}

struct ClipboardGuard;

impl ClipboardGuard {
    fn open() -> Result<Self> {
        unsafe {
            OpenClipboard(Some(HWND::default()))?;
        }
        Ok(Self)
    }
}

impl Drop for ClipboardGuard {
    fn drop(&mut self) {
        unsafe {
            let _ = CloseClipboard();
        }
    }
}

#[derive(Clone, Copy)]
struct AvailableFormats {
    internal: bool,
    rtf: bool,
    html: bool,
    unicode: bool,
}

fn preferred_source(available: AvailableFormats) -> Option<ClipboardSource> {
    if available.internal {
        return Some(ClipboardSource::Internal);
    }
    if available.rtf {
        return Some(ClipboardSource::Rtf);
    }
    if available.html {
        return Some(ClipboardSource::Html);
    }
    if available.unicode {
        return Some(ClipboardSource::UnicodeText);
    }
    None
}

fn set_clipboard_unicode_text(text: &str) -> Result<()> {
    let mut utf16: Vec<u16> = text.encode_utf16().collect();
    utf16.push(0);
    let bytes = unsafe {
        std::slice::from_raw_parts(utf16.as_ptr() as *const u8, utf16.len() * size_of::<u16>())
    };
    set_clipboard_raw(CF_UNICODETEXT_U32, bytes, false)
}

fn set_clipboard_bytes(format: u32, bytes: &[u8]) -> Result<()> {
    set_clipboard_raw(format, bytes, true)
}

fn set_clipboard_raw(format: u32, bytes: &[u8], append_nul: bool) -> Result<()> {
    let mut payload = Vec::with_capacity(bytes.len() + usize::from(append_nul));
    payload.extend_from_slice(bytes);
    if append_nul {
        payload.push(0);
    }

    unsafe {
        let handle: HGLOBAL = GlobalAlloc(GMEM_MOVEABLE, payload.len())?;
        let ptr = GlobalLock(handle) as *mut u8;
        if ptr.is_null() {
            return Err(Error::from_thread());
        }
        copy_nonoverlapping(payload.as_ptr(), ptr, payload.len());
        let _ = GlobalUnlock(handle);
        match SetClipboardData(format, Some(HANDLE(handle.0))) {
            Ok(_) => Ok(()),
            Err(err) => Err(err),
        }
    }
}

fn get_clipboard_unicode_text() -> Option<String> {
    if !is_format_available(CF_UNICODETEXT_U32) {
        return None;
    }

    unsafe {
        let handle = GetClipboardData(CF_UNICODETEXT_U32).ok()?;
        if handle.is_invalid() {
            return None;
        }

        let hglobal = HGLOBAL(handle.0);
        let ptr = GlobalLock(hglobal) as *const u16;
        if ptr.is_null() {
            return None;
        }

        let mut len = 0usize;
        while *ptr.add(len) != 0 {
            len += 1;
        }
        let slice = std::slice::from_raw_parts(ptr, len);
        let out = String::from_utf16(slice).ok();
        let _ = GlobalUnlock(hglobal);
        out
    }
}

fn get_clipboard_bytes(format: u32) -> Result<Option<Vec<u8>>> {
    if !is_format_available(format) {
        return Ok(None);
    }

    unsafe {
        let handle = GetClipboardData(format)?;
        if handle.is_invalid() {
            return Ok(None);
        }
        let hglobal = HGLOBAL(handle.0);
        let ptr = GlobalLock(hglobal) as *const u8;
        if ptr.is_null() {
            return Ok(None);
        }

        let size = GlobalSize(hglobal);
        let mut bytes = std::slice::from_raw_parts(ptr, size).to_vec();
        let _ = GlobalUnlock(hglobal);

        while bytes.last() == Some(&0) {
            bytes.pop();
        }
        Ok(Some(bytes))
    }
}

fn is_format_available(format: u32) -> bool {
    unsafe { IsClipboardFormatAvailable(format).is_ok() }
}

fn internal_clipboard_format() -> u32 {
    *INTERNAL_CLIPBOARD_FORMAT
        .get_or_init(|| unsafe { RegisterClipboardFormatW(w!("Doco.InternalRuns")) })
}

fn rtf_clipboard_format() -> u32 {
    *RTF_CLIPBOARD_FORMAT
        .get_or_init(|| unsafe { RegisterClipboardFormatW(w!("Rich Text Format")) })
}

fn html_clipboard_format() -> u32 {
    *HTML_CLIPBOARD_FORMAT.get_or_init(|| unsafe { RegisterClipboardFormatW(w!("HTML Format")) })
}

fn png_clipboard_format() -> u32 {
    *PNG_CLIPBOARD_FORMAT.get_or_init(|| unsafe { RegisterClipboardFormatW(w!("PNG")) })
}

fn decode_clipboard_image(bytes: Vec<u8>, mime: &str) -> Option<ClipboardImageData> {
    let decoded = image::load_from_memory(bytes.as_slice()).ok()?;
    let (width, height) = decoded.dimensions();
    Some(ClipboardImageData {
        bytes,
        mime: mime.to_string(),
        width,
        height,
    })
}

fn dib_to_bmp_bytes(dib: &[u8]) -> Option<Vec<u8>> {
    if dib.len() < 40 {
        return None;
    }

    let header_size = u32::from_le_bytes(dib[0..4].try_into().ok()?) as usize;
    if header_size < 40 || dib.len() < header_size {
        return None;
    }

    let bpp = u16::from_le_bytes(dib[14..16].try_into().ok()?);
    let compression = u32::from_le_bytes(dib[16..20].try_into().ok()?);
    let colors_used = u32::from_le_bytes(dib[32..36].try_into().ok()?);
    let masks_len = if compression == 3 && header_size == 40 {
        12usize
    } else {
        0usize
    };
    let palette_entries = if bpp <= 8 {
        if colors_used == 0 {
            1u32.checked_shl(bpp as u32).unwrap_or(0)
        } else {
            colors_used
        }
    } else {
        0
    };
    let palette_len = palette_entries as usize * 4;
    let pixel_offset_in_dib = header_size
        .checked_add(masks_len)?
        .checked_add(palette_len)?;
    if pixel_offset_in_dib > dib.len() {
        return None;
    }

    let file_header_len = 14usize;
    let file_size = file_header_len.checked_add(dib.len())?;
    let pixel_offset_in_file = file_header_len.checked_add(pixel_offset_in_dib)?;
    let mut bmp = Vec::with_capacity(file_size);
    bmp.extend_from_slice(b"BM");
    bmp.extend_from_slice(&(file_size as u32).to_le_bytes());
    bmp.extend_from_slice(&0u16.to_le_bytes());
    bmp.extend_from_slice(&0u16.to_le_bytes());
    bmp.extend_from_slice(&(pixel_offset_in_file as u32).to_le_bytes());
    bmp.extend_from_slice(dib);
    Some(bmp)
}

fn should_emit_rich_formats(runs: &[Run], source_format: DocumentFormat) -> bool {
    if source_format == DocumentFormat::Docx {
        return true;
    }
    runs.len() > 1 || runs.iter().any(|run| !run_style_is_plain(&run.style))
}

fn run_style_is_plain(style: &RunStyle) -> bool {
    !style.bold
        && !style.italic
        && !style.underline
        && !style.strikethrough
        && !style.superscript
        && !style.subscript
        && style.font_family.is_none()
        && style.font_size.is_none()
        && style.color.is_none()
        && style.background.is_none()
}

fn runs_to_plain_text(runs: &[Run]) -> String {
    let cap = runs.iter().map(|run| run.text.len()).sum();
    let mut out = String::with_capacity(cap);
    for run in runs {
        out.push_str(&run.text);
    }
    out
}

#[derive(Serialize, Deserialize)]
struct InternalClipboardPayload {
    version: u8,
    plain_text: String,
    runs: Vec<Run>,
}

fn encode_internal_payload(runs: &[Run], plain_text: &str) -> serde_json::Result<Vec<u8>> {
    serde_json::to_vec(&InternalClipboardPayload {
        version: 1,
        plain_text: plain_text.to_string(),
        runs: runs.to_vec(),
    })
}

fn decode_internal_payload(bytes: &[u8]) -> Option<Vec<Run>> {
    let payload = serde_json::from_slice::<InternalClipboardPayload>(bytes).ok()?;
    if payload.version != 1 {
        return None;
    }
    Some(payload.runs)
}

fn runs_to_minimal_rtf(runs: &[Run]) -> String {
    if runs.is_empty() {
        return String::new();
    }

    let mut out = String::from("{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Segoe UI;}}\\uc1\\pard ");
    let mut active = RunStyle::default();

    for run in runs {
        emit_rtf_style_delta(&mut out, &active, &run.style);
        active = run.style.clone();
        out.push_str(&escape_rtf_text(&run.text));
    }
    out.push('}');
    out
}

fn emit_rtf_style_delta(out: &mut String, prev: &RunStyle, next: &RunStyle) {
    if prev.bold != next.bold {
        out.push_str(if next.bold { "\\b " } else { "\\b0 " });
    }
    if prev.italic != next.italic {
        out.push_str(if next.italic { "\\i " } else { "\\i0 " });
    }
    if prev.underline != next.underline {
        out.push_str(if next.underline { "\\ul " } else { "\\ul0 " });
    }
    if prev.strikethrough != next.strikethrough {
        out.push_str(if next.strikethrough {
            "\\strike "
        } else {
            "\\strike0 "
        });
    }

    if next.superscript {
        if !prev.superscript {
            out.push_str("\\super ");
        }
    } else if next.subscript {
        if !prev.subscript {
            out.push_str("\\sub ");
        }
    } else if prev.superscript || prev.subscript {
        out.push_str("\\nosupersub ");
    }
}

fn escape_rtf_text(text: &str) -> String {
    let mut out = String::with_capacity(text.len() + text.len() / 4);
    for ch in text.chars() {
        match ch {
            '\\' => out.push_str("\\\\"),
            '{' => out.push_str("\\{"),
            '}' => out.push_str("\\}"),
            '\n' => out.push_str("\\par "),
            '\t' => out.push_str("\\tab "),
            c if c.is_ascii() => out.push(c),
            c => {
                out.push_str("\\u");
                out.push_str(&(c as i32).to_string());
                out.push('?');
            }
        }
    }
    out
}

fn parse_rtf_to_runs(rtf: &str) -> Vec<Run> {
    let bytes = rtf.as_bytes();
    let mut i = 0usize;
    let mut style_stack = vec![RunStyle::default()];
    let mut skip_stack = vec![false];
    let mut runs = Vec::new();
    let mut text = String::new();

    while i < bytes.len() {
        match bytes[i] {
            b'{' => {
                flush_run(
                    &mut runs,
                    &mut text,
                    style_stack.last().expect("style stack"),
                );
                style_stack.push(style_stack.last().expect("style stack").clone());
                skip_stack.push(*skip_stack.last().unwrap_or(&false));
                i += 1;
            }
            b'}' => {
                flush_run(
                    &mut runs,
                    &mut text,
                    style_stack.last().expect("style stack"),
                );
                if style_stack.len() > 1 {
                    style_stack.pop();
                }
                if skip_stack.len() > 1 {
                    skip_stack.pop();
                }
                i += 1;
            }
            b'\\' => {
                i += 1;
                if i >= bytes.len() {
                    break;
                }

                match bytes[i] {
                    b'\\' | b'{' | b'}' => {
                        if !skip_stack.last().copied().unwrap_or(false) {
                            text.push(bytes[i] as char);
                        }
                        i += 1;
                    }
                    b'\'' => {
                        if i + 2 < bytes.len() {
                            let h1 = bytes[i + 1] as char;
                            let h2 = bytes[i + 2] as char;
                            if let (Some(a), Some(b)) = (h1.to_digit(16), h2.to_digit(16)) {
                                let value = ((a << 4) | b) as u8;
                                let raw = [value];
                                let (decoded, _, _) = WINDOWS_1252.decode(&raw);
                                if !skip_stack.last().copied().unwrap_or(false) {
                                    text.push_str(decoded.as_ref());
                                }
                            }
                            i += 3;
                        } else {
                            i = bytes.len();
                        }
                    }
                    _ => {
                        let word_start = i;
                        while i < bytes.len() && (bytes[i] as char).is_ascii_alphabetic() {
                            i += 1;
                        }
                        let word = std::str::from_utf8(&bytes[word_start..i]).unwrap_or_default();

                        let mut sign = 1i32;
                        if i < bytes.len() && bytes[i] == b'-' {
                            sign = -1;
                            i += 1;
                        }

                        let mut value = 0i32;
                        let mut has_value = false;
                        while i < bytes.len() && (bytes[i] as char).is_ascii_digit() {
                            has_value = true;
                            value = value * 10 + (bytes[i] - b'0') as i32;
                            i += 1;
                        }
                        let arg = has_value.then_some(sign * value);

                        if i < bytes.len() && bytes[i] == b' ' {
                            i += 1;
                        }

                        if is_ignored_rtf_destination(word) {
                            if let Some(skip) = skip_stack.last_mut() {
                                *skip = true;
                            }
                        }

                        if !skip_stack.last().copied().unwrap_or(false) {
                            apply_rtf_control(&mut runs, &mut text, &mut style_stack, word, arg);
                        }

                        if word == "u" && i < bytes.len() && bytes[i] == b'?' {
                            i += 1;
                        }
                    }
                }
            }
            b'\r' | b'\n' => {
                i += 1;
            }
            _ => {
                let c = rtf[i..].chars().next().unwrap_or('\0');
                if !skip_stack.last().copied().unwrap_or(false) {
                    text.push(c);
                }
                i += c.len_utf8();
            }
        }
    }

    flush_run(
        &mut runs,
        &mut text,
        style_stack.last().expect("style stack"),
    );
    runs
}

fn apply_rtf_control(
    runs: &mut Vec<Run>,
    text: &mut String,
    style_stack: &mut Vec<RunStyle>,
    word: &str,
    arg: Option<i32>,
) {
    let Some(style) = style_stack.last_mut() else {
        return;
    };

    match word {
        "b" => {
            flush_run(runs, text, style);
            style.bold = arg.unwrap_or(1) != 0;
        }
        "i" => {
            flush_run(runs, text, style);
            style.italic = arg.unwrap_or(1) != 0;
        }
        "ul" => {
            flush_run(runs, text, style);
            style.underline = arg.unwrap_or(1) != 0;
        }
        "strike" => {
            flush_run(runs, text, style);
            style.strikethrough = arg.unwrap_or(1) != 0;
        }
        "super" => {
            flush_run(runs, text, style);
            style.superscript = true;
            style.subscript = false;
        }
        "sub" => {
            flush_run(runs, text, style);
            style.subscript = true;
            style.superscript = false;
        }
        "nosupersub" => {
            flush_run(runs, text, style);
            style.subscript = false;
            style.superscript = false;
        }
        "plain" => {
            flush_run(runs, text, style);
            *style = RunStyle::default();
        }
        "par" | "line" => text.push('\n'),
        "tab" => text.push('\t'),
        "u" => {
            if let Some(code) = arg {
                let normalized = if code < 0 { code + 65536 } else { code };
                if let Some(c) = char::from_u32(normalized as u32) {
                    text.push(c);
                }
            }
        }
        _ => {}
    }
}

fn runs_to_cf_html(runs: &[Run]) -> String {
    let fragment = runs_to_html_fragment(runs);
    if fragment.is_empty() {
        return String::new();
    }
    build_cf_html(&fragment)
}

fn runs_to_html_fragment(runs: &[Run]) -> String {
    let mut out = String::new();
    for run in runs {
        if run.text.is_empty() {
            continue;
        }
        let escaped = escape_html_text(&run.text);
        let css = style_to_css(&run.style);
        if css.is_empty() {
            out.push_str(&escaped);
        } else {
            out.push_str("<span style=\"");
            out.push_str(&css);
            out.push_str("\">");
            out.push_str(&escaped);
            out.push_str("</span>");
        }
    }
    out
}

fn build_cf_html(fragment: &str) -> String {
    let html_body =
        format!("<html><body><!--StartFragment-->{fragment}<!--EndFragment--></body></html>");
    let header_template = "Version:1.0\r\nStartHTML:0000000000\r\nEndHTML:0000000000\r\nStartFragment:0000000000\r\nEndFragment:0000000000\r\n";
    let start_html = header_template.len();
    let start_fragment = start_html + "<html><body><!--StartFragment-->".len();
    let end_fragment = start_fragment + fragment.len();
    let end_html = start_html + html_body.len();

    let header = format!(
        "Version:1.0\r\nStartHTML:{start_html:010}\r\nEndHTML:{end_html:010}\r\nStartFragment:{start_fragment:010}\r\nEndFragment:{end_fragment:010}\r\n"
    );
    format!("{header}{html_body}")
}

fn parse_html_to_runs(raw_html: &str) -> Vec<Run> {
    let html = extract_html_fragment(raw_html);
    let mut i = 0usize;
    let mut runs = Vec::new();
    let mut style_stack = vec![RunStyle::default()];
    let mut text = String::new();

    while i < html.len() {
        let byte = html.as_bytes()[i];
        if byte == b'<' {
            if let Some(end_rel) = html[i..].find('>') {
                let end = i + end_rel;
                let tag = html[i + 1..end].trim();
                flush_run(
                    &mut runs,
                    &mut text,
                    style_stack.last().expect("style stack"),
                );
                handle_html_tag(tag, &mut style_stack, &mut text);
                i = end + 1;
                continue;
            }
        }
        if byte == b'&' {
            if let Some((decoded, next)) = decode_html_entity(html, i) {
                text.push_str(&decoded);
                i = next;
                continue;
            }
        }

        let ch = html[i..].chars().next().unwrap_or('\0');
        text.push(ch);
        i += ch.len_utf8();
    }

    flush_run(
        &mut runs,
        &mut text,
        style_stack.last().expect("style stack"),
    );
    runs
}

fn is_ignored_rtf_destination(word: &str) -> bool {
    matches!(
        word,
        "fonttbl" | "colortbl" | "stylesheet" | "info" | "pict" | "object" | "header" | "footer"
    )
}

fn extract_html_fragment(raw_html: &str) -> &str {
    if let (Some(start), Some(end)) = (
        parse_offset_field(raw_html, "StartFragment:"),
        parse_offset_field(raw_html, "EndFragment:"),
    ) {
        if start < end && end <= raw_html.len() {
            return &raw_html[start..end];
        }
    }

    if let (Some(start_marker), Some(end_marker)) = (
        raw_html.find("<!--StartFragment-->"),
        raw_html.find("<!--EndFragment-->"),
    ) {
        let start = start_marker + "<!--StartFragment-->".len();
        if start < end_marker && end_marker <= raw_html.len() {
            return &raw_html[start..end_marker];
        }
    }

    raw_html
}

fn parse_offset_field(raw_html: &str, key: &str) -> Option<usize> {
    let idx = raw_html.find(key)?;
    let after = &raw_html[idx + key.len()..];
    let digits: String = after.chars().take_while(|c| c.is_ascii_digit()).collect();
    digits.parse::<usize>().ok()
}

fn handle_html_tag(tag: &str, style_stack: &mut Vec<RunStyle>, text: &mut String) {
    if tag.is_empty() {
        return;
    }

    let trimmed = tag.trim_matches('/');
    let name = trimmed
        .split_whitespace()
        .next()
        .unwrap_or_default()
        .to_ascii_lowercase();
    let closing = tag.starts_with('/');

    if !closing {
        match name.as_str() {
            "b" | "strong" => push_modified_style(style_stack, |s| s.bold = true),
            "i" | "em" => push_modified_style(style_stack, |s| s.italic = true),
            "u" => push_modified_style(style_stack, |s| s.underline = true),
            "s" | "strike" | "del" => push_modified_style(style_stack, |s| s.strikethrough = true),
            "sup" => push_modified_style(style_stack, |s| {
                s.superscript = true;
                s.subscript = false;
            }),
            "sub" => push_modified_style(style_stack, |s| {
                s.subscript = true;
                s.superscript = false;
            }),
            "span" => {
                let style_text = extract_attr_value(tag, "style").unwrap_or_default();
                push_modified_style(style_stack, |s| apply_inline_css(s, &style_text));
            }
            "br" => text.push('\n'),
            _ => {}
        }
    } else {
        match name.as_str() {
            "b" | "strong" | "i" | "em" | "u" | "s" | "strike" | "del" | "sup" | "sub" | "span" => {
                if style_stack.len() > 1 {
                    style_stack.pop();
                }
            }
            "p" | "div" | "li" => {
                if !text.ends_with('\n') {
                    text.push('\n');
                }
            }
            _ => {}
        }
    }
}

fn push_modified_style<F>(style_stack: &mut Vec<RunStyle>, mutator: F)
where
    F: FnOnce(&mut RunStyle),
{
    let mut next = style_stack.last().cloned().unwrap_or_default();
    mutator(&mut next);
    style_stack.push(next);
}

fn extract_attr_value(tag: &str, attr_name: &str) -> Option<String> {
    let lower = tag.to_ascii_lowercase();
    let key = format!("{attr_name}=");
    let idx = lower.find(&key)?;
    let rest = &tag[idx + key.len()..].trim_start();
    let quote = rest.chars().next()?;
    if quote == '"' || quote == '\'' {
        let end = rest[1..].find(quote)?;
        Some(rest[1..1 + end].to_string())
    } else {
        let end = rest.find(char::is_whitespace).unwrap_or(rest.len());
        Some(rest[..end].to_string())
    }
}

fn apply_inline_css(style: &mut RunStyle, css: &str) {
    for decl in css.split(';') {
        let mut parts = decl.splitn(2, ':');
        let key = parts.next().unwrap_or_default().trim().to_ascii_lowercase();
        let value = parts.next().unwrap_or_default().trim().to_ascii_lowercase();
        match key.as_str() {
            "font-weight" if value.contains("bold") || value == "700" => style.bold = true,
            "font-style" if value.contains("italic") => style.italic = true,
            "text-decoration" => {
                if value.contains("underline") {
                    style.underline = true;
                }
                if value.contains("line-through") {
                    style.strikethrough = true;
                }
            }
            "vertical-align" if value.contains("super") => {
                style.superscript = true;
                style.subscript = false;
            }
            "vertical-align" if value.contains("sub") => {
                style.subscript = true;
                style.superscript = false;
            }
            _ => {}
        }
    }
}

fn decode_html_entity(input: &str, offset: usize) -> Option<(String, usize)> {
    let tail = &input[offset..];
    for (entity, decoded) in [
        ("&amp;", "&"),
        ("&lt;", "<"),
        ("&gt;", ">"),
        ("&quot;", "\""),
        ("&#39;", "'"),
        ("&nbsp;", " "),
    ] {
        if tail.starts_with(entity) {
            return Some((decoded.to_string(), offset + entity.len()));
        }
    }

    if let Some(rest) = tail.strip_prefix("&#") {
        let end = rest.find(';')?;
        let body = &rest[..end];
        let value = if let Some(hex) = body.strip_prefix('x').or_else(|| body.strip_prefix('X')) {
            u32::from_str_radix(hex, 16).ok()?
        } else {
            body.parse::<u32>().ok()?
        };
        let decoded = char::from_u32(value)?.to_string();
        return Some((decoded, offset + 2 + end + 1));
    }

    None
}

fn escape_html_text(text: &str) -> String {
    let mut out = String::with_capacity(text.len() + text.len() / 8);
    for ch in text.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&#39;"),
            '\n' => out.push_str("<br/>"),
            _ => out.push(ch),
        }
    }
    out
}

fn style_to_css(style: &RunStyle) -> String {
    let mut css = String::new();
    if style.bold {
        css.push_str("font-weight:bold;");
    }
    if style.italic {
        css.push_str("font-style:italic;");
    }
    if style.underline || style.strikethrough {
        css.push_str("text-decoration:");
        if style.underline {
            css.push_str(" underline");
        }
        if style.strikethrough {
            css.push_str(" line-through");
        }
        css.push(';');
    }
    if style.superscript {
        css.push_str("vertical-align:super;");
    } else if style.subscript {
        css.push_str("vertical-align:sub;");
    }
    css
}

fn flush_run(runs: &mut Vec<Run>, text: &mut String, style: &RunStyle) {
    if text.is_empty() {
        return;
    }
    if let Some(last) = runs.last_mut() {
        if last.style == *style {
            last.text.push_str(text);
            text.clear();
            return;
        }
    }
    runs.push(Run {
        text: std::mem::take(text),
        style: style.clone(),
    });
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;

    use image::{DynamicImage, ImageFormat};

    use super::*;

    fn plain_run(text: &str) -> Run {
        Run {
            text: text.to_string(),
            style: RunStyle::default(),
        }
    }

    #[test]
    fn prefers_expected_paste_source_order() {
        assert_eq!(
            preferred_source(AvailableFormats {
                internal: true,
                rtf: true,
                html: true,
                unicode: true,
            }),
            Some(ClipboardSource::Internal)
        );
        assert_eq!(
            preferred_source(AvailableFormats {
                internal: false,
                rtf: true,
                html: true,
                unicode: true,
            }),
            Some(ClipboardSource::Rtf)
        );
        assert_eq!(
            preferred_source(AvailableFormats {
                internal: false,
                rtf: false,
                html: true,
                unicode: true,
            }),
            Some(ClipboardSource::Html)
        );
        assert_eq!(
            preferred_source(AvailableFormats {
                internal: false,
                rtf: false,
                html: false,
                unicode: true,
            }),
            Some(ClipboardSource::UnicodeText)
        );
    }

    #[test]
    fn rtf_roundtrip_keeps_basic_formatting() {
        let mut bold = RunStyle::default();
        bold.bold = true;
        let mut italic = RunStyle::default();
        italic.italic = true;

        let runs = vec![
            plain_run("Hello "),
            Run {
                text: "Bold".to_string(),
                style: bold,
            },
            Run {
                text: " and ".to_string(),
                style: RunStyle::default(),
            },
            Run {
                text: "Italic".to_string(),
                style: italic,
            },
        ];

        let rtf = runs_to_minimal_rtf(&runs);
        assert!(rtf.contains("\\b "));
        assert!(rtf.contains("\\i "));

        let parsed = parse_rtf_to_runs(&rtf);
        assert_eq!(runs_to_plain_text(&parsed), "Hello Bold and Italic");
        assert!(parsed.iter().any(|r| r.style.bold));
        assert!(parsed.iter().any(|r| r.style.italic));
    }

    #[test]
    fn html_roundtrip_keeps_basic_formatting() {
        let mut underline = RunStyle::default();
        underline.underline = true;
        let runs = vec![
            plain_run("A "),
            Run {
                text: "B".to_string(),
                style: underline,
            },
            plain_run("\nC"),
        ];

        let html = runs_to_cf_html(&runs);
        let parsed = parse_html_to_runs(&html);

        assert_eq!(runs_to_plain_text(&parsed), "A B\nC");
        assert!(parsed.iter().any(|r| r.style.underline));
    }

    #[test]
    fn drag_drop_move_adjusts_offset_after_delete() {
        let selection = SelectionRange {
            start: CursorPosition {
                block_id: crate::document::model::BlockId(1),
                offset: 2,
            },
            end: CursorPosition {
                block_id: crate::document::model::BlockId(1),
                offset: 5,
            },
        };
        let drop_at = CursorPosition {
            block_id: crate::document::model::BlockId(1),
            offset: 8,
        };

        let commands = drag_drop_commands(selection, drop_at, "abc", false);
        assert_eq!(commands.len(), 2);
        match &commands[0] {
            EditCommand::DeleteText { start, end, .. } => {
                assert_eq!((*start, *end), (2, 5));
            }
            _ => panic!("expected delete command first"),
        }
        match &commands[1] {
            EditCommand::InsertText { offset, text, .. } => {
                assert_eq!(*offset, 5);
                assert_eq!(text, "abc");
            }
            _ => panic!("expected insert command second"),
        }
    }

    #[test]
    fn plain_source_does_not_force_rich_formats() {
        let runs = vec![plain_run("100000 chars test")];
        assert!(!should_emit_rich_formats(&runs, DocumentFormat::Text));
        assert!(should_emit_rich_formats(&runs, DocumentFormat::Docx));
    }

    #[test]
    fn dib_payload_converts_back_to_bmp() {
        let mut bmp = Vec::new();
        DynamicImage::new_rgba8(2, 2)
            .write_to(&mut Cursor::new(&mut bmp), ImageFormat::Bmp)
            .expect("encode bmp");
        let dib = bmp[14..].to_vec();
        let rebuilt = dib_to_bmp_bytes(dib.as_slice()).expect("convert dib");
        let decoded = image::load_from_memory(rebuilt.as_slice()).expect("decode rebuilt bmp");
        assert_eq!(decoded.dimensions(), (2, 2));
    }
}
