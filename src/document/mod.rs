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
        Some(ext)
            if matches!(
                ext.as_str(),
                "txt"
                    | "text"
                    | "log"
                    | "ini"
                    | "cfg"
                    | "conf"
                    | "toml"
                    | "yaml"
                    | "yml"
                    | "json"
                    | "jsonc"
                    | "xml"
                    | "csv"
                    | "tsv"
                    | "rs"
                    | "c"
                    | "h"
                    | "cpp"
                    | "hpp"
                    | "cc"
                    | "cs"
                    | "go"
                    | "java"
                    | "kt"
                    | "py"
                    | "js"
                    | "ts"
                    | "tsx"
                    | "jsx"
                    | "css"
                    | "scss"
                    | "html"
                    | "htm"
                    | "sql"
                    | "sh"
                    | "ps1"
                    | "bat"
                    | "cmd"
                    | "env"
                    | "gitignore"
                    | "gitattributes"
                    | "editorconfig"
                    | "lock"
                    | "cargo"
            ) =>
        {
            DocumentFormat::Text
        }
        _ => DocumentFormat::Unknown,
    }
}

#[cfg(test)]
mod tests {
    use super::{DocumentFormat, detect_format};
    use std::path::Path;

    #[test]
    fn detects_common_text_extensions() {
        assert_eq!(detect_format(Path::new("a.ini")), DocumentFormat::Text);
        assert_eq!(detect_format(Path::new("b.toml")), DocumentFormat::Text);
        assert_eq!(detect_format(Path::new("c.rs")), DocumentFormat::Text);
        assert_eq!(detect_format(Path::new("d.md")), DocumentFormat::Markdown);
        assert_eq!(detect_format(Path::new("e.docx")), DocumentFormat::Docx);
    }
}
