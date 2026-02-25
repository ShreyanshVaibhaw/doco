use std::{fs, path::Path};

use image::GenericImageView;
use regex::Regex;

#[derive(Debug, Clone)]
pub struct LoadedImageAsset {
    pub bytes: Vec<u8>,
    pub mime: String,
    pub width: u32,
    pub height: u32,
}

pub fn load_supported_image(path: &Path) -> Result<LoadedImageAsset, String> {
    let ext = path
        .extension()
        .and_then(|v| v.to_str())
        .map(|v| v.to_ascii_lowercase())
        .ok_or_else(|| "missing file extension".to_string())?;

    let mime = mime_for_extension(ext.as_str())
        .ok_or_else(|| format!("unsupported image format: {ext}"))?;
    let bytes = fs::read(path).map_err(|e| format!("failed to read image: {e}"))?;

    let (width, height) = if ext == "svg" {
        parse_svg_dimensions(bytes.as_slice()).unwrap_or((512, 512))
    } else {
        image::load_from_memory(bytes.as_slice())
            .map_err(|e| format!("failed to decode image: {e}"))?
            .dimensions()
    };

    Ok(LoadedImageAsset {
        bytes,
        mime: mime.to_string(),
        width,
        height,
    })
}

pub fn mime_for_extension(ext: &str) -> Option<&'static str> {
    match ext.to_ascii_lowercase().as_str() {
        "png" => Some("image/png"),
        "jpg" | "jpeg" => Some("image/jpeg"),
        "bmp" => Some("image/bmp"),
        "gif" => Some("image/gif"),
        "webp" => Some("image/webp"),
        "tif" | "tiff" => Some("image/tiff"),
        "svg" => Some("image/svg+xml"),
        _ => None,
    }
}

fn parse_svg_dimensions(bytes: &[u8]) -> Option<(u32, u32)> {
    let source = String::from_utf8_lossy(bytes);
    let root = Regex::new(r"(?is)<svg\b([^>]*)>").ok()?;
    let captures = root.captures(source.as_ref())?;
    let attrs = captures.get(1)?.as_str();

    let width = parse_svg_attr_length(attrs, "width");
    let height = parse_svg_attr_length(attrs, "height");
    match (width, height) {
        (Some(w), Some(h)) => return Some((w.max(1), h.max(1))),
        _ => {}
    }

    let view_box = parse_svg_view_box(attrs)?;
    Some((view_box.0.max(1), view_box.1.max(1)))
}

fn parse_svg_attr_length(attrs: &str, name: &str) -> Option<u32> {
    let pattern = format!(r#"(?i)\b{}\s*=\s*["']([^"']+)["']"#, regex::escape(name));
    let regex = Regex::new(pattern.as_str()).ok()?;
    let value = regex.captures(attrs)?.get(1)?.as_str();
    parse_svg_numeric(value)
}

fn parse_svg_view_box(attrs: &str) -> Option<(u32, u32)> {
    let regex = Regex::new(r#"(?i)\bviewBox\s*=\s*["']([^"']+)["']"#).ok()?;
    let value = regex.captures(attrs)?.get(1)?.as_str();
    let parts = value
        .split(|c: char| c.is_whitespace() || c == ',')
        .filter(|s| !s.is_empty())
        .collect::<Vec<_>>();
    if parts.len() != 4 {
        return None;
    }
    let width = parts[2].parse::<f32>().ok()?;
    let height = parts[3].parse::<f32>().ok()?;
    if !width.is_finite() || !height.is_finite() {
        return None;
    }
    Some((width.round() as u32, height.round() as u32))
}

fn parse_svg_numeric(value: &str) -> Option<u32> {
    let mut out = String::new();
    for ch in value.trim().chars() {
        if ch.is_ascii_digit() || ch == '.' {
            out.push(ch);
        } else {
            break;
        }
    }
    let parsed = out.parse::<f32>().ok()?;
    if !parsed.is_finite() || parsed <= 0.0 {
        return None;
    }
    Some(parsed.round() as u32)
}

#[cfg(test)]
mod tests {
    use std::{
        fs,
        path::PathBuf,
        time::{SystemTime, UNIX_EPOCH},
    };

    use image::DynamicImage;

    use super::load_supported_image;

    fn temp_file(name: &str) -> PathBuf {
        let tick = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        std::env::temp_dir().join(format!("doco-image-test-{tick}-{name}"))
    }

    #[test]
    fn loads_png_dimensions() {
        let path = temp_file("sample.png");
        DynamicImage::new_rgba8(2, 3)
            .save(&path)
            .expect("write png");

        let loaded = load_supported_image(&path).expect("load png");
        assert_eq!(loaded.mime, "image/png");
        assert_eq!(loaded.width, 2);
        assert_eq!(loaded.height, 3);
        assert!(!loaded.bytes.is_empty());

        let _ = fs::remove_file(path);
    }

    #[test]
    fn loads_svg_dimensions_from_viewbox() {
        let path = temp_file("sample.svg");
        fs::write(
            &path,
            r#"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 480"></svg>"#,
        )
        .expect("write svg");

        let loaded = load_supported_image(&path).expect("load svg");
        assert_eq!(loaded.mime, "image/svg+xml");
        assert_eq!(loaded.width, 640);
        assert_eq!(loaded.height, 480);

        let _ = fs::remove_file(path);
    }

    #[test]
    fn rejects_unsupported_extension() {
        let path = temp_file("sample.txt");
        fs::write(&path, "not an image").expect("write text");

        let err = load_supported_image(&path).expect_err("should reject unsupported type");
        assert!(err.contains("unsupported image format"));

        let _ = fs::remove_file(path);
    }
}
