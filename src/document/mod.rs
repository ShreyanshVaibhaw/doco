pub mod docx;
pub mod export;
pub mod markdown;
pub mod model;
pub mod pdf;
pub mod txt;

use std::path::Path;

use model::DocumentModel;
use serde::{Deserialize, Serialize};

pub trait DocumentHandler {
    fn load(path: &Path) -> std::io::Result<DocumentModel>;
    fn save(path: &Path, model: &DocumentModel) -> std::io::Result<()>;
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum DocumentFormat {
    Docx,
    Pdf,
    Markdown,
    Text,
    Unknown,
}

pub fn detect_format(path: &Path) -> DocumentFormat {
    match path.extension().and_then(|s| s.to_str()).map(|s| s.to_ascii_lowercase()) {
        Some(ext) if ext == "docx" => DocumentFormat::Docx,
        Some(ext) if ext == "pdf" => DocumentFormat::Pdf,
        Some(ext) if ext == "md" || ext == "markdown" => DocumentFormat::Markdown,
        Some(ext) if ext == "txt" => DocumentFormat::Text,
        _ => DocumentFormat::Unknown,
    }
}
