use std::{
    collections::HashMap,
    fs,
    hash::{Hash, Hasher},
    path::Path,
    time::{Duration, Instant},
};

use image::{DynamicImage, GenericImageView, imageops::FilterType};

use crate::document::model::{
    DocumentModel,
    ImageBlock,
    ImageData,
    ImageDataRef,
};

#[derive(Debug, Clone)]
pub struct DecodedBitmap {
    pub width: u32,
    pub height: u32,
    pub rgba: Vec<u8>,
    pub source_hash: u64,
    pub is_thumbnail: bool,
}

#[derive(Debug)]
struct CacheEntry {
    bitmap: DecodedBitmap,
    bytes: usize,
    last_used: Instant,
}

#[derive(Debug, Default, Clone, Copy)]
pub struct ImageCacheStats {
    pub hits: u64,
    pub misses: u64,
    pub full_res_entries: usize,
    pub thumbnail_entries: usize,
    pub full_res_bytes: usize,
    pub thumbnail_bytes: usize,
}

impl ImageCacheStats {
    pub fn hit_rate(self) -> f32 {
        let total = self.hits + self.misses;
        if total == 0 {
            0.0
        } else {
            self.hits as f32 / total as f32
        }
    }
}

#[derive(Debug, Default)]
struct CacheStore {
    entries: HashMap<u64, CacheEntry>,
    current_bytes: usize,
}

impl CacheStore {
    fn get(&mut self, key: u64) -> Option<DecodedBitmap> {
        if let Some(entry) = self.entries.get_mut(&key) {
            entry.last_used = Instant::now();
            return Some(entry.bitmap.clone());
        }
        None
    }

    fn insert(&mut self, key: u64, bitmap: DecodedBitmap) {
        let bytes = bitmap.rgba.len();
        if let Some(prev) = self.entries.insert(
            key,
            CacheEntry {
                bitmap,
                bytes,
                last_used: Instant::now(),
            },
        ) {
            self.current_bytes = self.current_bytes.saturating_sub(prev.bytes);
        }
        self.current_bytes += bytes;
    }

    fn remove(&mut self, key: u64) {
        if let Some(old) = self.entries.remove(&key) {
            self.current_bytes = self.current_bytes.saturating_sub(old.bytes);
        }
    }

    fn clear(&mut self) {
        self.entries.clear();
        self.current_bytes = 0;
    }

    fn prune_idle(&mut self, idle_ttl: Duration) {
        let now = Instant::now();
        self.entries.retain(|_, entry| {
            let stale = now.saturating_duration_since(entry.last_used) >= idle_ttl;
            if stale {
                self.current_bytes = self.current_bytes.saturating_sub(entry.bytes);
            }
            !stale
        });
    }

    fn touch_keys(&mut self, keys: &[u64]) {
        let now = Instant::now();
        for key in keys {
            if let Some(entry) = self.entries.get_mut(key) {
                entry.last_used = now;
            }
        }
    }

    fn remove_oldest_until(&mut self, max_bytes: usize) {
        if self.current_bytes <= max_bytes {
            return;
        }

        let mut sorted = self
            .entries
            .iter()
            .map(|(key, entry)| (*key, entry.last_used))
            .collect::<Vec<_>>();
        sorted.sort_by_key(|(_, last_used)| *last_used);

        for (key, _) in sorted {
            if self.current_bytes <= max_bytes {
                break;
            }
            self.remove(key);
        }
    }
}

#[derive(Debug)]
pub struct ImageDecodeCache {
    pub max_bytes: usize,
    pub current_bytes: usize,
    full_res: CacheStore,
    thumbnails: CacheStore,
    idle_ttl: Duration,
    stats: ImageCacheStats,
}

impl Default for ImageDecodeCache {
    fn default() -> Self {
        Self {
            max_bytes: 200 * 1024 * 1024,
            current_bytes: 0,
            full_res: CacheStore::default(),
            thumbnails: CacheStore::default(),
            idle_ttl: Duration::from_secs(30),
            stats: ImageCacheStats::default(),
        }
    }
}

impl ImageDecodeCache {
    pub fn get_or_decode(
        &mut self,
        source: &ImageData,
        thumbnail_max_dim: Option<u32>,
    ) -> Result<DecodedBitmap, image::ImageError> {
        self.sweep_idle_decoded_bitmaps();

        let key = hash_image_data(source, thumbnail_max_dim);
        let target_cache = if thumbnail_max_dim.is_some() {
            &mut self.thumbnails
        } else {
            &mut self.full_res
        };

        if let Some(bitmap) = target_cache.get(key) {
            self.stats.hits += 1;
            return Ok(bitmap);
        }
        self.stats.misses += 1;

        let image = image::load_from_memory(&source.bytes)?;
        let bitmap = decode_bitmap(image, key, thumbnail_max_dim);
        target_cache.insert(key, bitmap.clone());

        self.prune_memory();
        self.update_stats();

        Ok(bitmap)
    }

    pub fn invalidate_hash(&mut self, hash: u64) {
        self.full_res.remove(hash);
        self.thumbnails.remove(hash);
        self.update_stats();
    }

    pub fn clear(&mut self) {
        self.full_res.clear();
        self.thumbnails.clear();
        self.current_bytes = 0;
        self.stats = ImageCacheStats::default();
    }

    pub fn set_memory_budget(&mut self, max_bytes: usize) {
        self.max_bytes = max_bytes.max(32 * 1024 * 1024);
        self.prune_memory();
        self.update_stats();
    }

    pub fn mark_visible_hashes(&mut self, hashes: &[u64]) {
        self.full_res.touch_keys(hashes);
        self.thumbnails.touch_keys(hashes);
    }

    pub fn sweep_idle_decoded_bitmaps(&mut self) {
        self.full_res.prune_idle(self.idle_ttl);
        self.thumbnails.prune_idle(self.idle_ttl);
        self.update_stats();
    }

    pub fn stats(&self) -> ImageCacheStats {
        self.stats
    }

    fn prune_memory(&mut self) {
        let full_budget = (self.max_bytes as f32 * 0.80) as usize;
        let thumb_budget = self.max_bytes.saturating_sub(full_budget);

        self.full_res.remove_oldest_until(full_budget);
        self.thumbnails.remove_oldest_until(thumb_budget.max(8 * 1024 * 1024));

        self.current_bytes = self.full_res.current_bytes + self.thumbnails.current_bytes;
    }

    fn update_stats(&mut self) {
        self.stats.full_res_entries = self.full_res.entries.len();
        self.stats.thumbnail_entries = self.thumbnails.entries.len();
        self.stats.full_res_bytes = self.full_res.current_bytes;
        self.stats.thumbnail_bytes = self.thumbnails.current_bytes;
        self.current_bytes = self.stats.full_res_bytes + self.stats.thumbnail_bytes;
    }
}

pub fn resolve_image_data(block: &ImageBlock, doc: &DocumentModel) -> Option<ImageData> {
    match &block.data {
        ImageDataRef::Embedded(data) => Some(data.clone()),
        ImageDataRef::Key(key) => doc.images.get(key).cloned(),
        ImageDataRef::LinkedPath(path) => load_image_from_path(path),
        ImageDataRef::Empty => {
            if let Some(data) = doc.images.get(block.key.as_str()) {
                return Some(data.clone());
            }
            block
                .source_path
                .as_ref()
                .and_then(|p| load_image_from_path(p))
        }
    }
}

pub fn interpolation_hint(scale: f32) -> &'static str {
    if scale < 1.0 {
        "D2D1_INTERPOLATION_MODE_HIGH_QUALITY_CUBIC"
    } else {
        "D2D1_INTERPOLATION_MODE_LINEAR"
    }
}

fn load_image_from_path(path: &Path) -> Option<ImageData> {
    let bytes = fs::read(path).ok()?;
    let decoded = image::load_from_memory(bytes.as_slice()).ok()?;
    let (width, height) = decoded.dimensions();
    let mime = mime_for_extension(path.extension().and_then(|e| e.to_str()));

    Some(ImageData {
        bytes,
        mime,
        width,
        height,
    })
}

fn decode_bitmap(image: DynamicImage, source_hash: u64, thumbnail_max_dim: Option<u32>) -> DecodedBitmap {
    let (width, height) = image.dimensions();

    if let Some(max_dim) = thumbnail_max_dim.filter(|m| *m > 0) {
        let thumb = image.thumbnail(max_dim, max_dim).to_rgba8();
        return DecodedBitmap {
            width: thumb.width(),
            height: thumb.height(),
            rgba: thumb.into_raw(),
            source_hash,
            is_thumbnail: true,
        };
    }

    // Lanczos3 gives a high-quality reduction path before upload to GPU.
    let raster = if width > 4096 || height > 4096 {
        image
            .resize(4096, 4096, FilterType::Lanczos3)
            .to_rgba8()
    } else {
        image.to_rgba8()
    };

    DecodedBitmap {
        width: raster.width(),
        height: raster.height(),
        rgba: raster.into_raw(),
        source_hash,
        is_thumbnail: false,
    }
}

fn hash_image_data(image: &ImageData, thumb_dim: Option<u32>) -> u64 {
    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    image.bytes.hash(&mut hasher);
    image.width.hash(&mut hasher);
    image.height.hash(&mut hasher);
    image.mime.hash(&mut hasher);
    thumb_dim.unwrap_or(0).hash(&mut hasher);
    hasher.finish()
}

fn mime_for_extension(ext: Option<&str>) -> String {
    match ext.unwrap_or_default().to_ascii_lowercase().as_str() {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "bmp" => "image/bmp",
        "gif" => "image/gif",
        "webp" => "image/webp",
        "tif" | "tiff" => "image/tiff",
        "svg" => "image/svg+xml",
        _ => "application/octet-stream",
    }
    .to_string()
}
