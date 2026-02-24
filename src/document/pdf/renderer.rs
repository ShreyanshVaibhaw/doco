use std::{
    collections::{HashMap, VecDeque},
    path::{Path, PathBuf},
};

#[derive(Debug, Clone)]
pub struct PdfDocumentHandle {
    pub path: PathBuf,
    pub page_count: usize,
}

#[derive(Debug, Clone)]
pub struct PdfPageRenderResult {
    pub page_index: usize,
    pub width: u32,
    pub height: u32,
    pub rgba: Vec<u8>,
}

#[derive(Debug, Clone)]
pub struct PdfOutlineItem {
    pub title: String,
    pub page_index: usize,
    pub children: Vec<PdfOutlineItem>,
}

#[derive(Debug)]
pub enum PdfError {
    FeatureDisabled,
    Io(String),
    Render(String),
}

#[derive(Debug, Clone)]
struct CacheKey {
    page: usize,
    zoom_permille: u32,
}

#[derive(Debug, Default)]
struct PdfPageCache {
    max_pages: usize,
    order: VecDeque<(usize, u32)>,
    items: HashMap<(usize, u32), PdfPageRenderResult>,
}

impl PdfPageCache {
    fn with_capacity(max_pages: usize) -> Self {
        Self {
            max_pages,
            ..Self::default()
        }
    }

    fn get(&mut self, key: &CacheKey) -> Option<PdfPageRenderResult> {
        let k = (key.page, key.zoom_permille);
        if let Some(hit) = self.items.get(&k) {
            self.order.retain(|v| *v != k);
            self.order.push_back(k);
            Some(hit.clone())
        } else {
            None
        }
    }

    fn put(&mut self, key: CacheKey, value: PdfPageRenderResult) {
        let k = (key.page, key.zoom_permille);
        self.order.retain(|v| *v != k);
        self.order.push_back(k);
        self.items.insert(k, value);
        while self.items.len() > self.max_pages {
            if let Some(old) = self.order.pop_front() {
                self.items.remove(&old);
            } else {
                break;
            }
        }
    }
}

pub struct PdfRenderer {
    cache: PdfPageCache,
    pub memory_budget_mb: usize,
}

impl Default for PdfRenderer {
    fn default() -> Self {
        Self {
            cache: PdfPageCache::with_capacity(10),
            memory_budget_mb: 50,
        }
    }
}

impl PdfRenderer {
    pub fn open_path(&mut self, path: &Path) -> Result<PdfDocumentHandle, PdfError> {
        #[cfg(feature = "pdf")]
        {
            return self.open_path_impl(path);
        }

        #[cfg(not(feature = "pdf"))]
        {
            let _ = path;
            Err(PdfError::FeatureDisabled)
        }
    }

    pub fn render_page(
        &mut self,
        handle: &PdfDocumentHandle,
        page_index: usize,
        zoom: f32,
        dpi: f32,
    ) -> Result<PdfPageRenderResult, PdfError> {
        let key = CacheKey {
            page: page_index,
            zoom_permille: (zoom * 1000.0) as u32,
        };

        if let Some(hit) = self.cache.get(&key) {
            return Ok(hit);
        }

        #[cfg(feature = "pdf")]
        {
            let result = self.render_page_impl(handle, page_index, zoom, dpi)?;
            self.cache.put(key, result.clone());
            return Ok(result);
        }

        #[cfg(not(feature = "pdf"))]
        {
            let _ = (handle, page_index, zoom, dpi);
            Err(PdfError::FeatureDisabled)
        }
    }

    pub fn extract_text(
        &mut self,
        handle: &PdfDocumentHandle,
        page_index: usize,
    ) -> Result<String, PdfError> {
        #[cfg(feature = "pdf")]
        {
            return self.extract_text_impl(handle, page_index);
        }
        #[cfg(not(feature = "pdf"))]
        {
            let _ = (handle, page_index);
            Err(PdfError::FeatureDisabled)
        }
    }

    pub fn outline(&mut self, handle: &PdfDocumentHandle) -> Result<Vec<PdfOutlineItem>, PdfError> {
        #[cfg(feature = "pdf")]
        {
            return self.outline_impl(handle);
        }
        #[cfg(not(feature = "pdf"))]
        {
            let _ = handle;
            Err(PdfError::FeatureDisabled)
        }
    }
}

#[cfg(feature = "pdf")]
impl PdfRenderer {
    fn open_path_impl(&mut self, path: &Path) -> Result<PdfDocumentHandle, PdfError> {
        use pdfium_render::prelude::*;

        let bindings = Pdfium::bind_to_system_library()
            .or_else(|_| Pdfium::bind_to_library("./pdfium.dll"))
            .map_err(|e| PdfError::Io(format!("failed to bind pdfium: {e}")))?;
        let pdfium = Pdfium::new(bindings);
        let document = pdfium
            .load_pdf_from_file(path, None)
            .map_err(|e| PdfError::Io(format!("failed to open pdf: {e}")))?;
        let page_count = document.pages().len() as usize;
        Ok(PdfDocumentHandle {
            path: path.to_path_buf(),
            page_count,
        })
    }

    fn render_page_impl(
        &mut self,
        handle: &PdfDocumentHandle,
        page_index: usize,
        zoom: f32,
        dpi: f32,
    ) -> Result<PdfPageRenderResult, PdfError> {
        use pdfium_render::prelude::*;

        let bindings = Pdfium::bind_to_system_library()
            .or_else(|_| Pdfium::bind_to_library("./pdfium.dll"))
            .map_err(|e| PdfError::Io(format!("failed to bind pdfium: {e}")))?;
        let pdfium = Pdfium::new(bindings);
        let doc = pdfium
            .load_pdf_from_file(&handle.path, None)
            .map_err(|e| PdfError::Io(format!("failed to reopen pdf: {e}")))?;
        let page = doc
            .pages()
            .get(page_index as u16)
            .map_err(|e| PdfError::Render(format!("page out of range: {e}")))?;

        let scale = (zoom * (dpi / 96.0)).max(0.25);
        let render = page.render_with_config(
            &PdfRenderConfig::new()
                .scale_page_by_factor(scale)
                .render_form_data(true),
        );
        let bitmap = render
            .map_err(|e| PdfError::Render(format!("render failed: {e}")))?;
        let width = bitmap.width() as u32;
        let height = bitmap.height() as u32;
        let rgba = bitmap.as_rgba_bytes();

        Ok(PdfPageRenderResult {
            page_index,
            width,
            height,
            rgba,
        })
    }

    fn extract_text_impl(
        &mut self,
        handle: &PdfDocumentHandle,
        page_index: usize,
    ) -> Result<String, PdfError> {
        use pdfium_render::prelude::*;

        let bindings = Pdfium::bind_to_system_library()
            .or_else(|_| Pdfium::bind_to_library("./pdfium.dll"))
            .map_err(|e| PdfError::Io(format!("failed to bind pdfium: {e}")))?;
        let pdfium = Pdfium::new(bindings);
        let doc = pdfium
            .load_pdf_from_file(&handle.path, None)
            .map_err(|e| PdfError::Io(format!("failed to open pdf: {e}")))?;
        let page = doc
            .pages()
            .get(page_index as u16)
            .map_err(|e| PdfError::Render(format!("page out of range: {e}")))?;
        let text = page
            .text()
            .all()
            .map_err(|e| PdfError::Render(format!("text extract failed: {e}")))?;
        Ok(text)
    }

    fn outline_impl(&mut self, _handle: &PdfDocumentHandle) -> Result<Vec<PdfOutlineItem>, PdfError> {
        // Placeholder: map PDFium bookmarks into this tree in the next pass.
        Ok(Vec::new())
    }
}
