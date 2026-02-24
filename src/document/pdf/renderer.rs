use std::{
    collections::{HashMap, VecDeque},
    path::{Path, PathBuf},
    sync::Arc,
};

#[derive(Debug, Clone)]
pub struct PdfDocumentHandle {
    pub source: PdfSource,
    pub page_count: usize,
}

#[derive(Debug, Clone)]
pub enum PdfSource {
    Path(PathBuf),
    Memory(Arc<Vec<u8>>),
}

#[derive(Debug, Clone)]
pub struct PdfPageRenderResult {
    pub page_index: usize,
    pub width: u32,
    pub height: u32,
    pub rgba: Vec<u8>,
}

#[derive(Debug, Clone)]
pub struct PdfTextSpan {
    pub text: String,
    pub bounds: (f32, f32, f32, f32),
}

#[derive(Debug, Clone)]
pub struct PdfPageTextContent {
    pub page_index: usize,
    pub full_text: String,
    pub spans: Vec<PdfTextSpan>,
}

#[derive(Debug, Clone)]
pub struct PdfPageProgressiveRender {
    pub low_res: PdfPageRenderResult,
    pub high_res: PdfPageRenderResult,
}

#[derive(Debug, Clone)]
pub struct PdfPageThumbnail {
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
    InvalidPage(usize),
    PasswordRequired,
}

#[derive(Debug, Clone)]
struct CacheKey {
    page: usize,
    zoom_permille: u32,
}

impl CacheKey {
    fn from(page: usize, zoom: f32) -> Self {
        Self {
            page,
            zoom_permille: (zoom.max(0.1) * 1000.0) as u32,
        }
    }
}

#[derive(Debug, Default)]
struct PdfPageCache {
    max_pages: usize,
    max_bytes: usize,
    current_bytes: usize,
    order: VecDeque<(usize, u32)>,
    items: HashMap<(usize, u32), PdfPageRenderResult>,
}

impl PdfPageCache {
    fn with_capacity(max_pages: usize, max_bytes: usize) -> Self {
        Self {
            max_pages: max_pages.max(1),
            max_bytes: max_bytes.max(1),
            ..Self::default()
        }
    }

    fn set_limits(&mut self, max_pages: usize, max_bytes: usize) {
        self.max_pages = max_pages.max(1);
        self.max_bytes = max_bytes.max(1);
        self.evict_if_needed();
    }

    fn clear(&mut self) {
        self.order.clear();
        self.items.clear();
        self.current_bytes = 0;
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
        let incoming_bytes = value.rgba.len();

        if let Some(existing) = self.items.get(&k) {
            self.current_bytes = self.current_bytes.saturating_sub(existing.rgba.len());
        }

        self.order.retain(|v| *v != k);
        self.order.push_back(k);
        self.items.insert(k, value);
        self.current_bytes = self.current_bytes.saturating_add(incoming_bytes);
        self.evict_if_needed();
    }

    fn evict_if_needed(&mut self) {
        while self.items.len() > self.max_pages || self.current_bytes > self.max_bytes {
            if let Some(old) = self.order.pop_front() {
                if let Some(old_value) = self.items.remove(&old) {
                    self.current_bytes = self.current_bytes.saturating_sub(old_value.rgba.len());
                }
            } else {
                break;
            }
        }
    }
}

pub struct PdfRenderer {
    cache: PdfPageCache,
    thumbnail_cache: PdfPageCache,
    pub max_cached_pages: usize,
    pub memory_budget_mb: usize,
    document_password: Option<String>,
    #[cfg(feature = "pdf")]
    pdfium: Option<pdfium_render::prelude::Pdfium>,
}

impl Default for PdfRenderer {
    fn default() -> Self {
        let max_cached_pages = 10;
        let memory_budget_mb = 50;
        let max_bytes = memory_budget_mb * 1024 * 1024;

        Self {
            cache: PdfPageCache::with_capacity(max_cached_pages, max_bytes),
            thumbnail_cache: PdfPageCache::with_capacity(max_cached_pages * 2, max_bytes / 3),
            max_cached_pages,
            memory_budget_mb,
            document_password: None,
            #[cfg(feature = "pdf")]
            pdfium: None,
        }
    }
}

impl PdfRenderer {
    pub fn set_memory_budget_mb(&mut self, budget_mb: usize) {
        self.memory_budget_mb = budget_mb.max(1);
        let bytes = self.memory_budget_mb * 1024 * 1024;
        self.cache.set_limits(self.max_cached_pages, bytes);
        self.thumbnail_cache
            .set_limits(self.max_cached_pages.saturating_mul(2), (bytes / 3).max(1));
    }

    pub fn set_max_cached_pages(&mut self, pages: usize) {
        self.max_cached_pages = pages.max(1);
        let bytes = self.memory_budget_mb * 1024 * 1024;
        self.cache.set_limits(self.max_cached_pages, bytes);
        self.thumbnail_cache
            .set_limits(self.max_cached_pages.saturating_mul(2), (bytes / 3).max(1));
    }

    pub fn set_document_password(&mut self, password: Option<String>) {
        self.document_password = password;
    }

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

    pub fn open_bytes(&mut self, bytes: Vec<u8>) -> Result<PdfDocumentHandle, PdfError> {
        #[cfg(feature = "pdf")]
        {
            return self.open_bytes_impl(bytes);
        }

        #[cfg(not(feature = "pdf"))]
        {
            let _ = bytes;
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
        if page_index >= handle.page_count {
            return Err(PdfError::InvalidPage(page_index));
        }

        let scaled_zoom = zoom * (dpi / 96.0).max(0.1);
        let key = CacheKey::from(page_index, scaled_zoom);

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

    pub fn render_page_progressive(
        &mut self,
        handle: &PdfDocumentHandle,
        page_index: usize,
        zoom: f32,
        dpi: f32,
    ) -> Result<PdfPageProgressiveRender, PdfError> {
        let low_res = self.render_page(handle, page_index, zoom, (dpi * 0.5).max(48.0))?;
        let high_res = self.render_page(handle, page_index, zoom, dpi)?;
        Ok(PdfPageProgressiveRender { low_res, high_res })
    }

    pub fn render_thumbnail(
        &mut self,
        handle: &PdfDocumentHandle,
        page_index: usize,
        max_side_px: u32,
    ) -> Result<PdfPageThumbnail, PdfError> {
        if page_index >= handle.page_count {
            return Err(PdfError::InvalidPage(page_index));
        }

        let side = max_side_px.max(32) as f32;
        let key = CacheKey::from(page_index, (side / 1000.0).max(0.1));

        if let Some(hit) = self.thumbnail_cache.get(&key) {
            return Ok(PdfPageThumbnail {
                page_index: hit.page_index,
                width: hit.width,
                height: hit.height,
                rgba: hit.rgba,
            });
        }

        #[cfg(feature = "pdf")]
        {
            let thumbnail = self.render_thumbnail_impl(handle, page_index, max_side_px)?;
            self.thumbnail_cache.put(
                key,
                PdfPageRenderResult {
                    page_index: thumbnail.page_index,
                    width: thumbnail.width,
                    height: thumbnail.height,
                    rgba: thumbnail.rgba.clone(),
                },
            );
            return Ok(thumbnail);
        }

        #[cfg(not(feature = "pdf"))]
        {
            let _ = (handle, page_index, max_side_px);
            Err(PdfError::FeatureDisabled)
        }
    }

    pub fn extract_text(
        &mut self,
        handle: &PdfDocumentHandle,
        page_index: usize,
    ) -> Result<String, PdfError> {
        if page_index >= handle.page_count {
            return Err(PdfError::InvalidPage(page_index));
        }

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

    pub fn extract_text_with_positions(
        &mut self,
        handle: &PdfDocumentHandle,
        page_index: usize,
    ) -> Result<PdfPageTextContent, PdfError> {
        if page_index >= handle.page_count {
            return Err(PdfError::InvalidPage(page_index));
        }

        #[cfg(feature = "pdf")]
        {
            return self.extract_text_with_positions_impl(handle, page_index);
        }

        #[cfg(not(feature = "pdf"))]
        {
            let _ = (handle, page_index);
            Err(PdfError::FeatureDisabled)
        }
    }

    pub fn go_to_page(
        &self,
        handle: &PdfDocumentHandle,
        one_based_page: usize,
    ) -> Result<usize, PdfError> {
        if one_based_page == 0 || one_based_page > handle.page_count {
            return Err(PdfError::InvalidPage(one_based_page.saturating_sub(1)));
        }

        Ok(one_based_page - 1)
    }

    pub fn page_count(&self, handle: &PdfDocumentHandle) -> usize {
        handle.page_count
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

    fn clear_document_caches(&mut self) {
        self.cache.clear();
        self.thumbnail_cache.clear();
    }
}

#[cfg(feature = "pdf")]
impl PdfRenderer {
    fn open_path_impl(&mut self, path: &Path) -> Result<PdfDocumentHandle, PdfError> {
        let password = self.document_password.clone();
        let page_count = {
            let pdfium = self.ensure_pdfium()?;
            let document = pdfium
                .load_pdf_from_file(path, password.as_deref())
                .map_err(|e| Self::map_pdfium_open_error("failed to open pdf from path", e))?;
            document.pages().len() as usize
        };

        self.clear_document_caches();

        Ok(PdfDocumentHandle {
            source: PdfSource::Path(path.to_path_buf()),
            page_count,
        })
    }

    fn open_bytes_impl(&mut self, bytes: Vec<u8>) -> Result<PdfDocumentHandle, PdfError> {
        let shared = Arc::new(bytes);
        let password = self.document_password.clone();

        let page_count = {
            let pdfium = self.ensure_pdfium()?;
            let document = pdfium
                .load_pdf_from_byte_slice(shared.as_slice(), password.as_deref())
                .map_err(|e| Self::map_pdfium_open_error("failed to open pdf from memory", e))?;
            document.pages().len() as usize
        };

        self.clear_document_caches();

        Ok(PdfDocumentHandle {
            source: PdfSource::Memory(shared),
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

        let password = self.document_password.clone();
        let pdfium = self.ensure_pdfium()?;
        let document = match &handle.source {
            PdfSource::Path(path) => pdfium
                .load_pdf_from_file(path, password.as_deref())
                .map_err(|e| Self::map_pdfium_open_error("failed to load pdf from path", e))?,
            PdfSource::Memory(bytes) => pdfium
                .load_pdf_from_byte_slice(bytes.as_slice(), password.as_deref())
                .map_err(|e| Self::map_pdfium_open_error("failed to load pdf from memory", e))?,
        };
        let page = document
            .pages()
            .get(page_index as u16)
            .map_err(|e| PdfError::Render(format!("page out of range: {e}")))?;

        let scale = (zoom * (dpi / 96.0)).max(0.25);
        let bitmap = page
            .render_with_config(
                &PdfRenderConfig::new()
                    .scale_page_by_factor(scale)
                    .render_annotations(true)
                    .render_form_data(true),
            )
            .map_err(|e| PdfError::Render(format!("render failed: {e}")))?;

        Ok(PdfPageRenderResult {
            page_index,
            width: bitmap.width() as u32,
            height: bitmap.height() as u32,
            rgba: bitmap.as_rgba_bytes(),
        })
    }

    fn render_thumbnail_impl(
        &mut self,
        handle: &PdfDocumentHandle,
        page_index: usize,
        max_side_px: u32,
    ) -> Result<PdfPageThumbnail, PdfError> {
        use pdfium_render::prelude::*;

        let password = self.document_password.clone();
        let pdfium = self.ensure_pdfium()?;
        let document = match &handle.source {
            PdfSource::Path(path) => pdfium
                .load_pdf_from_file(path, password.as_deref())
                .map_err(|e| Self::map_pdfium_open_error("failed to load pdf from path", e))?,
            PdfSource::Memory(bytes) => pdfium
                .load_pdf_from_byte_slice(bytes.as_slice(), password.as_deref())
                .map_err(|e| Self::map_pdfium_open_error("failed to load pdf from memory", e))?,
        };
        let page = document
            .pages()
            .get(page_index as u16)
            .map_err(|e| PdfError::Render(format!("page out of range: {e}")))?;

        let thumb_size = i32::try_from(max_side_px.max(32)).unwrap_or(i32::MAX);
        let bitmap = page
            .render_with_config(&PdfRenderConfig::new().thumbnail(thumb_size))
            .map_err(|e| PdfError::Render(format!("thumbnail render failed: {e}")))?;

        Ok(PdfPageThumbnail {
            page_index,
            width: bitmap.width() as u32,
            height: bitmap.height() as u32,
            rgba: bitmap.as_rgba_bytes(),
        })
    }

    fn extract_text_impl(
        &mut self,
        handle: &PdfDocumentHandle,
        page_index: usize,
    ) -> Result<String, PdfError> {
        let password = self.document_password.clone();
        let pdfium = self.ensure_pdfium()?;
        let document = match &handle.source {
            PdfSource::Path(path) => pdfium
                .load_pdf_from_file(path, password.as_deref())
                .map_err(|e| Self::map_pdfium_open_error("failed to load pdf from path", e))?,
            PdfSource::Memory(bytes) => pdfium
                .load_pdf_from_byte_slice(bytes.as_slice(), password.as_deref())
                .map_err(|e| Self::map_pdfium_open_error("failed to load pdf from memory", e))?,
        };
        let page = document
            .pages()
            .get(page_index as u16)
            .map_err(|e| PdfError::Render(format!("page out of range: {e}")))?;

        let text = page
            .text()
            .map_err(|e| PdfError::Render(format!("text extract failed: {e}")))?;

        Ok(text.all())
    }

    fn extract_text_with_positions_impl(
        &mut self,
        handle: &PdfDocumentHandle,
        page_index: usize,
    ) -> Result<PdfPageTextContent, PdfError> {
        let password = self.document_password.clone();
        let pdfium = self.ensure_pdfium()?;
        let document = match &handle.source {
            PdfSource::Path(path) => pdfium
                .load_pdf_from_file(path, password.as_deref())
                .map_err(|e| Self::map_pdfium_open_error("failed to load pdf from path", e))?,
            PdfSource::Memory(bytes) => pdfium
                .load_pdf_from_byte_slice(bytes.as_slice(), password.as_deref())
                .map_err(|e| Self::map_pdfium_open_error("failed to load pdf from memory", e))?,
        };
        let page = document
            .pages()
            .get(page_index as u16)
            .map_err(|e| PdfError::Render(format!("page out of range: {e}")))?;

        let page_height = page.height().value;
        let page_text = page
            .text()
            .map_err(|e| PdfError::Render(format!("text extract failed: {e}")))?;
        let full_text = page_text.all();

        let mut spans = Vec::new();
        for segment in page_text.segments().iter() {
            let rect = segment.bounds();
            let text = segment.text();
            if text.trim().is_empty() {
                continue;
            }

            spans.push(PdfTextSpan {
                text,
                bounds: (
                    rect.left().value,
                    (page_height - rect.top().value).max(0.0),
                    rect.width().value.max(0.0),
                    rect.height().value.max(0.0),
                ),
            });
        }

        Ok(PdfPageTextContent {
            page_index,
            full_text,
            spans,
        })
    }

    fn outline_impl(&mut self, handle: &PdfDocumentHandle) -> Result<Vec<PdfOutlineItem>, PdfError> {
        let password = self.document_password.clone();
        let pdfium = self.ensure_pdfium()?;
        let document = match &handle.source {
            PdfSource::Path(path) => pdfium
                .load_pdf_from_file(path, password.as_deref())
                .map_err(|e| Self::map_pdfium_open_error("failed to load pdf from path", e))?,
            PdfSource::Memory(bytes) => pdfium
                .load_pdf_from_byte_slice(bytes.as_slice(), password.as_deref())
                .map_err(|e| Self::map_pdfium_open_error("failed to load pdf from memory", e))?,
        };
        let bookmarks = document.bookmarks();

        let mut result = Vec::new();
        let mut current = bookmarks.root();
        while let Some(bookmark) = current {
            let next = bookmark.next_sibling();
            result.push(Self::bookmark_to_outline_item(bookmark));
            current = next;
        }

        Ok(result)
    }

    fn bookmark_to_outline_item(bookmark: pdfium_render::prelude::PdfBookmark<'_>) -> PdfOutlineItem {
        let mut children = Vec::new();
        let mut child = bookmark.first_child();

        while let Some(node) = child {
            let next = node.next_sibling();
            children.push(Self::bookmark_to_outline_item(node));
            child = next;
        }

        let title = bookmark
            .title()
            .unwrap_or_else(|| "Untitled bookmark".to_string());
        let page_index = bookmark
            .destination()
            .and_then(|destination| destination.page_index().ok())
            .unwrap_or(0) as usize;

        PdfOutlineItem {
            title,
            page_index,
            children,
        }
    }

    fn ensure_pdfium(&mut self) -> Result<&pdfium_render::prelude::Pdfium, PdfError> {
        use pdfium_render::prelude::*;

        if self.pdfium.is_none() {
            let bindings = Pdfium::bind_to_library("./pdfium.dll")
                .or_else(|_| Pdfium::bind_to_system_library())
                .map_err(|e| PdfError::Io(format!("failed to bind pdfium library: {e}")))?;
            self.pdfium = Some(Pdfium::new(bindings));
        }

        Ok(self.pdfium.as_ref().expect("pdfium initialized"))
    }

    fn map_pdfium_open_error(context: &str, error: pdfium_render::prelude::PdfiumError) -> PdfError {
        use pdfium_render::prelude::{PdfiumError, PdfiumInternalError};

        match error {
            PdfiumError::PdfiumLibraryInternalError(PdfiumInternalError::PasswordError) => {
                PdfError::PasswordRequired
            }
            other => PdfError::Io(format!("{context}: {other}")),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fake_render(page_index: usize, width: u32, height: u32) -> PdfPageRenderResult {
        PdfPageRenderResult {
            page_index,
            width,
            height,
            rgba: vec![0_u8; width.saturating_mul(height).saturating_mul(4) as usize],
        }
    }

    #[test]
    fn cache_evicts_lru_entries() {
        let mut cache = PdfPageCache::with_capacity(2, 1024 * 1024);
        cache.put(CacheKey::from(0, 1.0), fake_render(0, 10, 10));
        cache.put(CacheKey::from(1, 1.0), fake_render(1, 10, 10));
        assert!(cache.get(&CacheKey::from(0, 1.0)).is_some());

        cache.put(CacheKey::from(2, 1.0), fake_render(2, 10, 10));

        assert!(cache.items.contains_key(&(0, 1000)));
        assert!(!cache.items.contains_key(&(1, 1000)));
        assert!(cache.items.contains_key(&(2, 1000)));
    }

    #[test]
    fn cache_honors_memory_limit() {
        let mut cache = PdfPageCache::with_capacity(10, 16);
        cache.put(CacheKey::from(0, 1.0), fake_render(0, 2, 2));
        assert_eq!(cache.current_bytes, 16);

        cache.put(CacheKey::from(1, 1.0), fake_render(1, 1, 1));
        assert!(cache.current_bytes <= 16);
    }

    #[test]
    fn renderer_navigation_and_limits() {
        let mut renderer = PdfRenderer::default();
        let handle = PdfDocumentHandle {
            source: PdfSource::Path(PathBuf::from("dummy.pdf")),
            page_count: 3,
        };

        assert_eq!(renderer.page_count(&handle), 3);
        assert_eq!(renderer.go_to_page(&handle, 1).expect("page 1"), 0);
        assert_eq!(renderer.go_to_page(&handle, 3).expect("page 3"), 2);
        assert!(matches!(
            renderer.go_to_page(&handle, 0),
            Err(PdfError::InvalidPage(0))
        ));
        assert!(matches!(
            renderer.go_to_page(&handle, 4),
            Err(PdfError::InvalidPage(3))
        ));

        renderer.set_memory_budget_mb(0);
        renderer.set_max_cached_pages(0);
        assert_eq!(renderer.memory_budget_mb, 1);
        assert_eq!(renderer.max_cached_pages, 1);
    }

    #[cfg(not(feature = "pdf"))]
    #[test]
    fn feature_disabled_api_returns_expected_errors() {
        let mut renderer = PdfRenderer::default();
        let handle = PdfDocumentHandle {
            source: PdfSource::Path(PathBuf::from("dummy.pdf")),
            page_count: 1,
        };

        assert!(matches!(
            renderer.open_path(Path::new("dummy.pdf")),
            Err(PdfError::FeatureDisabled)
        ));
        assert!(matches!(
            renderer.open_bytes(vec![1, 2, 3]),
            Err(PdfError::FeatureDisabled)
        ));
        assert!(matches!(
            renderer.render_page(&handle, 0, 1.0, 96.0),
            Err(PdfError::FeatureDisabled)
        ));
        assert!(matches!(
            renderer.render_page_progressive(&handle, 0, 1.0, 96.0),
            Err(PdfError::FeatureDisabled)
        ));
        assert!(matches!(
            renderer.render_thumbnail(&handle, 0, 128),
            Err(PdfError::FeatureDisabled)
        ));
        assert!(matches!(
            renderer.extract_text(&handle, 0),
            Err(PdfError::FeatureDisabled)
        ));
        assert!(matches!(
            renderer.extract_text_with_positions(&handle, 0),
            Err(PdfError::FeatureDisabled)
        ));
        assert!(matches!(
            renderer.outline(&handle),
            Err(PdfError::FeatureDisabled)
        ));
    }
}
