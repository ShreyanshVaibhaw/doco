use std::{
    fs,
    path::{Path, PathBuf},
    time::{Duration, Instant},
};

use chrono::Utc;

use crate::document::{
    docx::writer,
    model::{Block, DocumentModel, ListType},
};

pub fn save_docx(path: &Path, model: &DocumentModel) -> std::io::Result<()> {
    writer::write_docx(path, model)
}

pub fn export_txt(path: &Path, model: &DocumentModel) -> std::io::Result<()> {
    fs::write(path, to_plain_text(model))
}

pub fn export_markdown(path: &Path, model: &DocumentModel) -> std::io::Result<()> {
    fs::write(path, to_markdown(model))
}

pub fn export_html(path: &Path, model: &DocumentModel) -> std::io::Result<()> {
    fs::write(path, to_html(model))
}

pub fn export_pdf(path: &Path, model: &DocumentModel) -> std::io::Result<()> {
    // Minimal fallback PDF generator placeholder while the full render-to-PDF pipeline is wired.
    // The dependency stays available for richer output in the next iteration.
    let text = to_plain_text(model);
    let escaped = text.replace('(', "\\(").replace(')', "\\)");

    let body = format!(
        "%PDF-1.4\n1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj\n3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >> endobj\n4 0 obj << /Length {len} >> stream\nBT /F1 12 Tf 48 800 Td ({text}) Tj ET\nendstream endobj\n5 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj\n",
        len = escaped.len() + 29,
        text = escaped
    );

    let mut offsets = Vec::new();
    let mut out = Vec::new();
    out.extend_from_slice(b"%PDF-1.4\n");

    for object in body.split("endobj\n") {
        if object.trim().is_empty() {
            continue;
        }
        offsets.push(out.len());
        out.extend_from_slice(object.as_bytes());
        out.extend_from_slice(b"endobj\n");
    }

    let xref_pos = out.len();
    out.extend_from_slice(format!("xref\n0 {}\n", offsets.len() + 1).as_bytes());
    out.extend_from_slice(b"0000000000 65535 f \n");
    for off in &offsets {
        out.extend_from_slice(format!("{:010} 00000 n \n", off).as_bytes());
    }
    out.extend_from_slice(
        format!(
            "trailer << /Size {} /Root 1 0 R >>\nstartxref\n{}\n%%EOF\n",
            offsets.len() + 1,
            xref_pos
        )
        .as_bytes(),
    );

    fs::write(path, out)
}

pub fn to_plain_text(model: &DocumentModel) -> String {
    let mut out = String::new();
    for block in &model.content {
        match block {
            Block::Paragraph(p) => {
                out.push_str(p.runs.iter().map(|r| r.text.as_str()).collect::<String>().as_str());
                out.push('\n');
            }
            Block::Heading(h) => {
                out.push_str(h.runs.iter().map(|r| r.text.as_str()).collect::<String>().as_str());
                out.push('\n');
            }
            Block::CodeBlock(code) => {
                out.push_str(code.code.as_str());
                out.push('\n');
            }
            Block::List(list) => {
                for (i, item) in list.items.iter().enumerate() {
                    let prefix = match list.list_type {
                        ListType::Bullet => "- ".to_string(),
                        ListType::Numbered => format!("{}. ", list.start_number + i as u32),
                        ListType::Checkbox => {
                            if item.checked.unwrap_or(false) {
                                "[x] ".to_string()
                            } else {
                                "[ ] ".to_string()
                            }
                        }
                    };
                    out.push_str(prefix.as_str());
                    for block in &item.content {
                        if let Block::Paragraph(p) = block {
                            out.push_str(p.runs.iter().map(|r| r.text.as_str()).collect::<String>().as_str());
                        }
                    }
                    out.push('\n');
                }
            }
            Block::Table(table) => {
                for row in &table.rows {
                    let mut first = true;
                    for cell in &row.cells {
                        if !first {
                            out.push('\t');
                        }
                        first = false;
                        for block in &cell.blocks {
                            if let Block::Paragraph(p) = block {
                                out.push_str(
                                    p.runs
                                        .iter()
                                        .map(|r| r.text.as_str())
                                        .collect::<String>()
                                        .as_str(),
                                );
                            }
                        }
                    }
                    out.push('\n');
                }
            }
            Block::HorizontalRule => out.push_str("---\n"),
            Block::PageBreak => {
                out.push('\n');
                out.push(char::from(0x0C));
                out.push('\n');
            }
            Block::Image(img) => {
                out.push_str(format!("[Image: {}]\n", img.alt_text).as_str());
            }
            Block::BlockQuote(q) => {
                for block in &q.blocks {
                    if let Block::Paragraph(p) = block {
                        out.push_str("> ");
                        out.push_str(p.runs.iter().map(|r| r.text.as_str()).collect::<String>().as_str());
                        out.push('\n');
                    }
                }
            }
        }
    }
    out
}

pub fn to_markdown(model: &DocumentModel) -> String {
    let mut out = String::new();
    for block in &model.content {
        match block {
            Block::Heading(h) => {
                out.push_str("#".repeat(h.level.clamp(1, 6) as usize).as_str());
                out.push(' ');
                out.push_str(h.runs.iter().map(|r| r.text.as_str()).collect::<String>().as_str());
                out.push_str("\n\n");
            }
            Block::Paragraph(p) => {
                out.push_str(p.runs.iter().map(|r| r.text.as_str()).collect::<String>().as_str());
                out.push_str("\n\n");
            }
            Block::CodeBlock(c) => {
                out.push_str("```\n");
                out.push_str(c.code.as_str());
                out.push_str("\n```\n\n");
            }
            Block::HorizontalRule => out.push_str("---\n\n"),
            Block::List(list) => {
                for (i, item) in list.items.iter().enumerate() {
                    let marker = match list.list_type {
                        ListType::Bullet => "- ".to_string(),
                        ListType::Numbered => format!("{}. ", list.start_number + i as u32),
                        ListType::Checkbox => {
                            if item.checked.unwrap_or(false) {
                                "- [x] ".to_string()
                            } else {
                                "- [ ] ".to_string()
                            }
                        }
                    };
                    out.push_str(marker.as_str());
                    for block in &item.content {
                        if let Block::Paragraph(p) = block {
                            out.push_str(p.runs.iter().map(|r| r.text.as_str()).collect::<String>().as_str());
                        }
                    }
                    out.push('\n');
                }
                out.push('\n');
            }
            Block::Table(table) => {
                if let Some(first) = table.rows.first() {
                    out.push('|');
                    for _ in &first.cells {
                        out.push_str(" --- |");
                    }
                    out.push('\n');
                }
                for row in &table.rows {
                    out.push('|');
                    for cell in &row.cells {
                        let cell_text = cell
                            .blocks
                            .iter()
                            .filter_map(|b| match b {
                                Block::Paragraph(p) => Some(
                                    p.runs
                                        .iter()
                                        .map(|r| r.text.as_str())
                                        .collect::<String>(),
                                ),
                                _ => None,
                            })
                            .collect::<Vec<_>>()
                            .join(" ");
                        out.push(' ');
                        out.push_str(cell_text.as_str());
                        out.push_str(" |");
                    }
                    out.push('\n');
                }
                out.push('\n');
            }
            Block::Image(img) => {
                out.push_str(format!("![{}]({})\n\n", img.alt_text, img.key).as_str());
            }
            Block::PageBreak | Block::BlockQuote(_) => {}
        }
    }
    out
}

pub fn to_html(model: &DocumentModel) -> String {
    let mut body = String::new();
    for block in &model.content {
        match block {
            Block::Heading(h) => body.push_str(
                format!(
                    "<h{lvl}>{text}</h{lvl}>",
                    lvl = h.level.clamp(1, 6),
                    text = escape_html(h.runs.iter().map(|r| r.text.as_str()).collect::<String>().as_str())
                )
                .as_str(),
            ),
            Block::Paragraph(p) => body.push_str(
                format!(
                    "<p>{}</p>",
                    escape_html(p.runs.iter().map(|r| r.text.as_str()).collect::<String>().as_str())
                )
                .as_str(),
            ),
            Block::CodeBlock(c) => body.push_str(
                format!("<pre><code>{}</code></pre>", escape_html(c.code.as_str())).as_str(),
            ),
            Block::HorizontalRule => body.push_str("<hr/>"),
            Block::Image(img) => body.push_str(
                format!(
                    "<figure><img alt=\"{}\" src=\"{}\"/></figure>",
                    escape_html(img.alt_text.as_str()),
                    escape_html(img.key.as_str())
                )
                .as_str(),
            ),
            Block::Table(table) => {
                body.push_str("<table>");
                for row in &table.rows {
                    body.push_str("<tr>");
                    for cell in &row.cells {
                        let cell_text = cell
                            .blocks
                            .iter()
                            .filter_map(|b| match b {
                                Block::Paragraph(p) => Some(
                                    p.runs
                                        .iter()
                                        .map(|r| r.text.as_str())
                                        .collect::<String>(),
                                ),
                                _ => None,
                            })
                            .collect::<Vec<_>>()
                            .join(" ");
                        body.push_str(format!("<td>{}</td>", escape_html(cell_text.as_str())).as_str());
                    }
                    body.push_str("</tr>");
                }
                body.push_str("</table>");
            }
            Block::PageBreak | Block::List(_) | Block::BlockQuote(_) => {}
        }
    }

    format!(
        "<!doctype html><html><head><meta charset=\"utf-8\"><style>body{{font-family:Segoe UI,Arial,sans-serif;max-width:840px;margin:24px auto;line-height:1.4}}table{{border-collapse:collapse}}td{{border:1px solid #ccc;padding:6px}}</style></head><body>{}</body></html>",
        body
    )
}

fn escape_html(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

pub struct AutoSaveManager {
    pub interval: Duration,
    pub recovery_dir: PathBuf,
    last_save: Instant,
}

impl AutoSaveManager {
    pub fn new(interval_seconds: u64) -> Self {
        let recovery_dir = if let Some(portable) = crate::settings::portable_root() {
            portable.join("recovery")
        } else {
            dirs::data_dir()
                .unwrap_or_else(|| PathBuf::from("."))
                .join("Doco")
                .join("recovery")
        };
        let _ = fs::create_dir_all(&recovery_dir);

        Self {
            interval: Duration::from_secs(interval_seconds.max(5)),
            recovery_dir,
            last_save: Instant::now(),
        }
    }

    pub fn tick(&mut self, model: &DocumentModel) -> std::io::Result<Option<PathBuf>> {
        if self.last_save.elapsed() < self.interval || !model.dirty {
            return Ok(None);
        }

        let stamp = Utc::now().format("%Y%m%d-%H%M%S");
        let path = self.recovery_dir.join(format!("recovery-{}.json", stamp));
        let json = serde_json::to_vec_pretty(model).map_err(|e| std::io::Error::other(e.to_string()))?;
        fs::write(&path, json)?;
        self.last_save = Instant::now();
        Ok(Some(path))
    }

    pub fn list_recovery_files(&self) -> std::io::Result<Vec<PathBuf>> {
        let mut files = Vec::new();
        if self.recovery_dir.exists() {
            for entry in fs::read_dir(&self.recovery_dir)? {
                let path = entry?.path();
                if path.extension().and_then(|e| e.to_str()) == Some("json") {
                    files.push(path);
                }
            }
        }
        files.sort();
        Ok(files)
    }

    pub fn clear_recovery_files(&self) -> std::io::Result<()> {
        for file in self.list_recovery_files()? {
            let _ = fs::remove_file(file);
        }
        Ok(())
    }
}

pub fn save_with_format(path: &Path, model: &DocumentModel) -> std::io::Result<()> {
    match path
        .extension()
        .and_then(|v| v.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase()
        .as_str()
    {
        "docx" => save_docx(path, model),
        "pdf" => export_pdf(path, model),
        "txt" => export_txt(path, model),
        "md" | "markdown" => export_markdown(path, model),
        "html" | "htm" => export_html(path, model),
        _ => save_docx(path, model),
    }
}

