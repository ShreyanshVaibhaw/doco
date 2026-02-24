use std::{mem::size_of, ptr::copy_nonoverlapping};

use windows::{
    Win32::{
        Foundation::{HANDLE, HGLOBAL, HWND},
        System::{
            DataExchange::{
                CloseClipboard, EmptyClipboard, GetClipboardData, OpenClipboard,
                SetClipboardData,
            },
            Memory::{
                GMEM_MOVEABLE, GlobalAlloc, GlobalLock, GlobalUnlock,
            },
        },
    },
    core::{Error, Result},
};

pub fn set_plain_text(text: &str) -> Result<()> {
    unsafe {
        OpenClipboard(Some(HWND::default()))?;
        EmptyClipboard()?;

        let mut utf16: Vec<u16> = text.encode_utf16().collect();
        utf16.push(0);

        let bytes = utf16.len() * size_of::<u16>();
        let handle: HGLOBAL = GlobalAlloc(GMEM_MOVEABLE, bytes)?;
        let ptr = GlobalLock(handle) as *mut u16;
        if ptr.is_null() {
            CloseClipboard()?;
            return Err(Error::from_thread());
        }

        copy_nonoverlapping(utf16.as_ptr(), ptr, utf16.len());
        let _ = GlobalUnlock(handle);
        let _ = SetClipboardData(CF_UNICODETEXT_U32, Some(HANDLE(handle.0)))?;
        CloseClipboard()?;
    }

    Ok(())
}

pub fn get_plain_text() -> Result<Option<String>> {
    unsafe {
        OpenClipboard(Some(HWND::default()))?;
        let handle = GetClipboardData(CF_UNICODETEXT_U32)?;
        if handle.is_invalid() {
            CloseClipboard()?;
            return Ok(None);
        }

        let ptr = GlobalLock(HGLOBAL(handle.0)) as *const u16;
        if ptr.is_null() {
            CloseClipboard()?;
            return Ok(None);
        }

        let mut len = 0usize;
        while *ptr.add(len) != 0 {
            len += 1;
        }

        let slice = std::slice::from_raw_parts(ptr, len);
        let out = String::from_utf16(slice).ok();
        let _ = GlobalUnlock(HGLOBAL(handle.0));
        CloseClipboard()?;
        Ok(out)
    }
}

pub fn copy_to_clipboard(text: &str) {
    let _ = set_plain_text(text);
}

const CF_UNICODETEXT_U32: u32 = 13;
