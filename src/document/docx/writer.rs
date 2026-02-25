use std::{
    collections::{BTreeSet, HashMap},
    fs::File,
    io::{self, Read, Write},
    path::Path,
};

use regex::Regex;
use zip::{CompressionMethod, ZipArchive, ZipWriter, write::SimpleFileOptions};

use crate::document::model::{
    Block, DocumentModel, ImageData, ListType, Paragraph, Run, TableStylePreset,
};

#[derive(Debug, Clone)]
struct ImageAsset {
    key: String,
    rel_id: String,
    file_name: String,
    mime: String,
    bytes: Vec<u8>,
}

#[derive(Debug, Default)]
struct PackageSnapshot {
    preserved: Vec<(String, Vec<u8>)>,
    content_types: Option<String>,
    document_rels: Option<String>,
}

pub fn write_docx(path: &Path, model: &DocumentModel) -> io::Result<()> {
    let source = roundtrip_source(path, model);
    let snapshot = if let Some(source_path) = source {
        read_package_snapshot(source_path)?
    } else {
        PackageSnapshot::default()
    };
    let images = build_image_assets(model);
    write_package(path, model, &snapshot, &images)
}

fn roundtrip_source<'a>(target: &'a Path, model: &'a DocumentModel) -> Option<&'a Path> {
    if let Some(source) = model.metadata.file_path.as_deref()
        && source.exists()
        && source
            .extension()
            .and_then(|v| v.to_str())
            .map(|v| v.eq_ignore_ascii_case("docx"))
            .unwrap_or(false)
    {
        return Some(source);
    }
    if target.exists()
        && target
            .extension()
            .and_then(|v| v.to_str())
            .map(|v| v.eq_ignore_ascii_case("docx"))
            .unwrap_or(false)
    {
        return Some(target);
    }
    None
}

fn read_package_snapshot(path: &Path) -> io::Result<PackageSnapshot> {
    let file = File::open(path)?;
    let mut archive = ZipArchive::new(file)?;
    let mut snapshot = PackageSnapshot::default();

    for i in 0..archive.len() {
        let mut entry = archive.by_index(i)?;
        let name = entry.name().to_string();
        if name.ends_with('/') {
            continue;
        }
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes)?;

        match name.as_str() {
            "[Content_Types].xml" => {
                snapshot.content_types = String::from_utf8(bytes).ok();
            }
            "word/_rels/document.xml.rels" => {
                snapshot.document_rels = String::from_utf8(bytes).ok();
            }
            _ => {
                if should_preserve_entry(name.as_str()) {
                    snapshot.preserved.push((name, bytes));
                }
            }
        }
    }

    snapshot.preserved.sort_by(|a, b| a.0.cmp(&b.0));
    Ok(snapshot)
}

fn should_preserve_entry(name: &str) -> bool {
    !matches!(
        name,
        "[Content_Types].xml" | "word/document.xml" | "word/_rels/document.xml.rels"
    ) && !name.starts_with("word/media/")
}

fn build_image_assets(model: &DocumentModel) -> Vec<ImageAsset> {
    let mut keys = model.images.keys().cloned().collect::<Vec<_>>();
    keys.sort_unstable();
    keys.into_iter()
        .enumerate()
        .filter_map(|(idx, key)| {
            let image = model.images.get(&key)?;
            let ext = ext_from_mime(image.mime.as_str());
            Some(ImageAsset {
                key,
                rel_id: format!("rDocoImg{}", idx + 1),
                file_name: format!("image{}.{}", idx + 1, ext),
                mime: image.mime.clone(),
                bytes: image.bytes.clone(),
            })
        })
        .collect()
}

fn write_package(
    output_path: &Path,
    model: &DocumentModel,
    snapshot: &PackageSnapshot,
    images: &[ImageAsset],
) -> io::Result<()> {
    let image_rel_map = images
        .iter()
        .map(|asset| (asset.key.clone(), asset.rel_id.clone()))
        .collect::<HashMap<_, _>>();
    let document_xml = document_xml(model, &image_rel_map);
    let content_types = content_types_xml(snapshot.content_types.as_deref(), images);
    let doc_rels = document_rels_xml(snapshot.document_rels.as_deref(), images);

    let file = File::create(output_path)?;
    let mut zip = ZipWriter::new(file);
    let options = SimpleFileOptions::default()
        .compression_method(CompressionMethod::Deflated)
        .unix_permissions(0o644);

    for (name, bytes) in &snapshot.preserved {
        zip.start_file(name, options)?;
        zip.write_all(bytes)?;
    }

    if !snapshot
        .preserved
        .iter()
        .any(|(name, _)| name == "_rels/.rels")
    {
        zip.start_file("_rels/.rels", options)?;
        zip.write_all(root_rels_xml().as_bytes())?;
    }

    if !snapshot
        .preserved
        .iter()
        .any(|(name, _)| name == "word/styles.xml")
    {
        zip.start_file("word/styles.xml", options)?;
        zip.write_all(default_styles_xml().as_bytes())?;
    }

    zip.start_file("[Content_Types].xml", options)?;
    zip.write_all(content_types.as_bytes())?;

    zip.start_file("word/document.xml", options)?;
    zip.write_all(document_xml.as_bytes())?;

    zip.start_file("word/_rels/document.xml.rels", options)?;
    zip.write_all(doc_rels.as_bytes())?;

    for image in images {
        let entry = format!("word/media/{}", image.file_name);
        zip.start_file(entry, options)?;
        zip.write_all(image.bytes.as_slice())?;
    }

    zip.finish()?;
    Ok(())
}

fn content_types_xml(existing: Option<&str>, images: &[ImageAsset]) -> String {
    let image_exts = images
        .iter()
        .map(|img| ext_from_mime(img.mime.as_str()).to_string())
        .collect::<BTreeSet<_>>();

    if let Some(existing_xml) = existing {
        let mut out = existing_xml.to_string();
        for ext in image_exts {
            let probe = format!("Extension=\"{ext}\"");
            if !out.contains(probe.as_str()) {
                let default = format!(
                    "<Default Extension=\"{ext}\" ContentType=\"{}\"/>",
                    mime_from_ext(ext.as_str())
                );
                out = insert_before_types_end(out, default.as_str());
            }
        }
        if !out.contains("PartName=\"/word/document.xml\"") {
            out = insert_before_types_end(
                out,
                "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>",
            );
        }
        if !out.contains("PartName=\"/word/styles.xml\"") {
            out = insert_before_types_end(
                out,
                "<Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>",
            );
        }
        return out;
    }

    let mut defaults = vec![
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>".to_string(),
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>".to_string(),
    ];
    for ext in image_exts {
        defaults.push(format!(
            "<Default Extension=\"{ext}\" ContentType=\"{}\"/>",
            mime_from_ext(ext.as_str())
        ));
    }

    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n{}\n<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>\n<Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>\n</Types>",
        defaults.join("\n")
    )
}

fn insert_before_types_end(mut xml: String, snippet: &str) -> String {
    if let Some(idx) = xml.rfind("</Types>") {
        xml.insert_str(idx, format!("\n{snippet}\n").as_str());
        return xml;
    }
    xml
}

fn document_rels_xml(existing: Option<&str>, images: &[ImageAsset]) -> String {
    let mut kept = Vec::new();
    if let Some(existing_xml) = existing
        && let Ok(re) = Regex::new(r#"<Relationship\b[^>]*/>"#)
    {
        for m in re.find_iter(existing_xml) {
            let rel = m.as_str();
            if rel.contains("/relationships/image\"") {
                continue;
            }
            kept.push(rel.to_string());
        }
    }

    kept.sort();
    for image in images {
        kept.push(format!(
            "<Relationship Id=\"{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/{}\"/>",
            image.rel_id, image.file_name
        ));
    }

    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n{}\n</Relationships>",
        kept.join("\n")
    )
}

fn root_rels_xml() -> &'static str {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>
</Relationships>"
}

fn default_styles_xml() -> &'static str {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">
  <w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\">
    <w:name w:val=\"Normal\"/>
  </w:style>
  <w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"heading 1\"/></w:style>
  <w:style w:type=\"paragraph\" w:styleId=\"Heading2\"><w:name w:val=\"heading 2\"/></w:style>
  <w:style w:type=\"paragraph\" w:styleId=\"Heading3\"><w:name w:val=\"heading 3\"/></w:style>
</w:styles>"
}

fn document_xml(model: &DocumentModel, image_rel_map: &HashMap<String, String>) -> String {
    let mut body = String::new();
    for block in &model.content {
        body.push_str(block_xml(block, image_rel_map).as_str());
    }

    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">
  <w:body>{}<w:sectPr/></w:body>
</w:document>",
        body
    )
}

fn block_xml(block: &Block, image_rel_map: &HashMap<String, String>) -> String {
    match block {
        Block::Paragraph(p) => paragraph_xml(p),
        Block::Heading(h) => {
            let paragraph = Paragraph {
                id: h.id,
                runs: h.runs.clone(),
                alignment: crate::document::model::ParagraphAlignment::Left,
                spacing: crate::document::model::ParagraphSpacing::default(),
                indent: crate::document::model::Indent::default(),
                style_id: Some(format!("Heading{}", h.level.clamp(1, 6))),
            };
            paragraph_xml(&paragraph)
        }
        Block::CodeBlock(code) => {
            let paragraph = Paragraph {
                id: code.id,
                runs: vec![Run {
                    text: code.code.clone(),
                    style: crate::document::model::RunStyle {
                        font_family: Some("Consolas".to_string()),
                        ..crate::document::model::RunStyle::default()
                    },
                }],
                alignment: crate::document::model::ParagraphAlignment::Left,
                spacing: crate::document::model::ParagraphSpacing::default(),
                indent: crate::document::model::Indent::default(),
                style_id: None,
            };
            paragraph_xml(&paragraph)
        }
        Block::List(list) => {
            let mut out = String::new();
            for (idx, item) in list.items.iter().enumerate() {
                let bullet = match list.list_type {
                    ListType::Bullet => "\u{2022} ".to_string(),
                    ListType::Numbered => format!("{}. ", list.start_number + idx as u32),
                    ListType::Checkbox => {
                        if item.checked.unwrap_or(false) {
                            "[x] ".to_string()
                        } else {
                            "[ ] ".to_string()
                        }
                    }
                };
                let mut runs = vec![Run {
                    text: bullet,
                    style: crate::document::model::RunStyle::default(),
                }];
                for nested in &item.content {
                    if let Block::Paragraph(p) = nested {
                        runs.extend(p.runs.clone());
                    }
                }
                let paragraph = Paragraph {
                    id: item.id,
                    runs,
                    alignment: crate::document::model::ParagraphAlignment::Left,
                    spacing: crate::document::model::ParagraphSpacing::default(),
                    indent: crate::document::model::Indent::default(),
                    style_id: None,
                };
                out.push_str(paragraph_xml(&paragraph).as_str());
            }
            out
        }
        Block::Table(table) => {
            let mut out = String::new();
            out.push_str("<w:tbl><w:tblPr>");
            out.push_str(match table.style {
                TableStylePreset::Plain => "<w:tblStyle w:val=\"TableNormal\"/>",
                TableStylePreset::Grid => "<w:tblStyle w:val=\"TableGrid\"/>",
                TableStylePreset::HeaderAccent => "<w:tblStyle w:val=\"LightListAccent1\"/>",
                TableStylePreset::AlternatingRows => "<w:tblStyle w:val=\"LightList\"/>",
                TableStylePreset::Professional => "<w:tblStyle w:val=\"MediumGrid1Accent1\"/>",
            });
            out.push_str("<w:tblCellMar><w:top w:w=\"72\" w:type=\"dxa\"/><w:left w:w=\"72\" w:type=\"dxa\"/><w:bottom w:w=\"72\" w:type=\"dxa\"/><w:right w:w=\"72\" w:type=\"dxa\"/></w:tblCellMar>");
            out.push_str("</w:tblPr>");
            for (row_idx, row) in table.rows.iter().enumerate() {
                let row_h = table
                    .row_heights
                    .get(row_idx)
                    .copied()
                    .unwrap_or(28.0)
                    .max(18.0);
                out.push_str(
                    format!(
                        "<w:tr><w:trPr><w:trHeight w:val=\"{}\" w:hRule=\"atLeast\"/></w:trPr>",
                        (row_h * 20.0).round() as i32
                    )
                    .as_str(),
                );
                for cell in &row.cells {
                    if cell.rowspan == 0 || cell.colspan == 0 {
                        continue;
                    }
                    out.push_str("<w:tc><w:tcPr>");
                    if cell.colspan > 1 {
                        out.push_str(format!("<w:gridSpan w:val=\"{}\"/>", cell.colspan).as_str());
                    }
                    if cell.rowspan > 1 {
                        out.push_str("<w:vMerge w:val=\"restart\"/>");
                    }
                    if let Some(color) = cell.background {
                        out.push_str(
                            format!(
                                "<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"{}\"/>",
                                to_hex(color)
                            )
                            .as_str(),
                        );
                    }
                    out.push_str("</w:tcPr>");
                    let mut has_paragraph = false;
                    for nested in &cell.blocks {
                        has_paragraph = true;
                        out.push_str(block_xml(nested, image_rel_map).as_str());
                    }
                    if !has_paragraph {
                        out.push_str("<w:p/>");
                    }
                    out.push_str("</w:tc>");
                }
                out.push_str("</w:tr>");
            }
            out.push_str("</w:tbl>");
            out
        }
        Block::Image(img) => image_drawing_xml(img, image_rel_map),
        Block::HorizontalRule => "<w:p><w:r><w:t>---</w:t></w:r></w:p>".to_string(),
        Block::PageBreak => "<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>".to_string(),
        Block::BlockQuote(quote) => {
            let mut out = String::new();
            for nested in &quote.blocks {
                out.push_str(block_xml(nested, image_rel_map).as_str());
            }
            out
        }
    }
}

fn image_drawing_xml(
    image: &crate::document::model::ImageBlock,
    image_rel_map: &HashMap<String, String>,
) -> String {
    let Some(rel_id) = image_rel_map.get(&image.key) else {
        return "<w:p/>".to_string();
    };
    let cx = (image.width.max(1.0) * 9525.0).round() as i64;
    let cy = (image.height.max(1.0) * 9525.0).round() as i64;
    format!(
        "<w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"{cx}\" cy=\"{cy}\"/><wp:docPr id=\"1\" name=\"Image\" descr=\"{alt}\"/><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic><pic:nvPicPr><pic:cNvPr id=\"0\" name=\"Image\"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"{rid}\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>",
        rid = rel_id,
        alt = escape_xml(image.alt_text.as_str())
    )
}

fn paragraph_xml(p: &Paragraph) -> String {
    let mut out = String::new();
    out.push_str("<w:p>");

    let mut has_ppr = p.style_id.is_some()
        || !matches!(
            p.alignment,
            crate::document::model::ParagraphAlignment::Left
        )
        || p.spacing.before > 0.0
        || p.spacing.after > 0.0
        || p.spacing.line > 0.0
        || p.indent.left > 0.0
        || p.indent.right > 0.0
        || p.indent.first_line > 0.0;
    if has_ppr {
        out.push_str("<w:pPr>");
        if let Some(style) = &p.style_id {
            out.push_str(format!("<w:pStyle w:val=\"{}\"/>", escape_xml(style)).as_str());
        }
        let align = match p.alignment {
            crate::document::model::ParagraphAlignment::Left => None,
            crate::document::model::ParagraphAlignment::Center => Some("center"),
            crate::document::model::ParagraphAlignment::Right => Some("right"),
            crate::document::model::ParagraphAlignment::Justify => Some("both"),
        };
        if let Some(jc) = align {
            out.push_str(format!("<w:jc w:val=\"{}\"/>", jc).as_str());
        }
        if p.spacing.before > 0.0 || p.spacing.after > 0.0 || p.spacing.line > 0.0 {
            out.push_str(
                format!(
                    "<w:spacing w:before=\"{}\" w:after=\"{}\" w:line=\"{}\"/>",
                    (p.spacing.before * 20.0).round() as i32,
                    (p.spacing.after * 20.0).round() as i32,
                    (p.spacing.line * 20.0).round() as i32,
                )
                .as_str(),
            );
        }
        if p.indent.left > 0.0 || p.indent.right > 0.0 || p.indent.first_line > 0.0 {
            out.push_str(
                format!(
                    "<w:ind w:left=\"{}\" w:right=\"{}\" w:firstLine=\"{}\"/>",
                    (p.indent.left * 20.0).round() as i32,
                    (p.indent.right * 20.0).round() as i32,
                    (p.indent.first_line * 20.0).round() as i32,
                )
                .as_str(),
            );
        }
        out.push_str("</w:pPr>");
        has_ppr = false;
    }

    for run in &p.runs {
        out.push_str(run_xml(run).as_str());
    }
    if p.runs.is_empty() {
        out.push_str("<w:r><w:t></w:t></w:r>");
    }
    out.push_str("</w:p>");
    out
}

fn run_xml(run: &Run) -> String {
    let mut out = String::new();
    out.push_str("<w:r>");
    if has_run_props(run) {
        out.push_str("<w:rPr>");
        if run.style.bold {
            out.push_str("<w:b/>");
        }
        if run.style.italic {
            out.push_str("<w:i/>");
        }
        if run.style.underline {
            out.push_str("<w:u w:val=\"single\"/>");
        }
        if run.style.strikethrough {
            out.push_str("<w:strike/>");
        }
        if run.style.superscript {
            out.push_str("<w:vertAlign w:val=\"superscript\"/>");
        }
        if run.style.subscript {
            out.push_str("<w:vertAlign w:val=\"subscript\"/>");
        }
        if let Some(size) = run.style.font_size {
            out.push_str(format!("<w:sz w:val=\"{}\"/>", (size * 2.0).round() as i32).as_str());
        }
        if let Some(ff) = &run.style.font_family {
            out.push_str(
                format!(
                    "<w:rFonts w:ascii=\"{}\" w:hAnsi=\"{}\"/>",
                    escape_xml(ff),
                    escape_xml(ff)
                )
                .as_str(),
            );
        }
        if let Some(color) = run.style.color {
            out.push_str(format!("<w:color w:val=\"{}\"/>", to_hex(color)).as_str());
        }
        if let Some(bg) = run.style.background {
            out.push_str(format!("<w:highlight w:val=\"{}\"/>", highlight_name(bg)).as_str());
        }
        out.push_str("</w:rPr>");
    }
    out.push_str(
        format!(
            "<w:t xml:space=\"preserve\">{}</w:t>",
            escape_xml(run.text.as_str())
        )
        .as_str(),
    );
    out.push_str("</w:r>");
    out
}

fn has_run_props(run: &Run) -> bool {
    run.style.bold
        || run.style.italic
        || run.style.underline
        || run.style.strikethrough
        || run.style.superscript
        || run.style.subscript
        || run.style.font_size.is_some()
        || run.style.font_family.is_some()
        || run.style.color.is_some()
        || run.style.background.is_some()
}

fn escape_xml(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn to_hex(c: crate::ui::Color) -> String {
    format!(
        "{:02X}{:02X}{:02X}",
        (c.r.clamp(0.0, 1.0) * 255.0) as u8,
        (c.g.clamp(0.0, 1.0) * 255.0) as u8,
        (c.b.clamp(0.0, 1.0) * 255.0) as u8
    )
}

fn highlight_name(c: crate::ui::Color) -> &'static str {
    let r = (c.r.clamp(0.0, 1.0) * 255.0) as i32;
    let g = (c.g.clamp(0.0, 1.0) * 255.0) as i32;
    let b = (c.b.clamp(0.0, 1.0) * 255.0) as i32;
    if r > 220 && g > 220 && b < 120 {
        "yellow"
    } else if r > 220 && g < 120 && b < 120 {
        "red"
    } else if r < 120 && g > 220 && b < 120 {
        "green"
    } else if r < 120 && g < 120 && b > 220 {
        "blue"
    } else {
        "none"
    }
}

fn ext_from_mime(mime: &str) -> &'static str {
    match mime {
        "image/png" => "png",
        "image/jpeg" => "jpg",
        "image/bmp" => "bmp",
        "image/gif" => "gif",
        "image/webp" => "webp",
        "image/tiff" => "tiff",
        "image/svg+xml" => "svg",
        _ => "bin",
    }
}

fn mime_from_ext(ext: &str) -> &'static str {
    match ext {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "bmp" => "image/bmp",
        "gif" => "image/gif",
        "webp" => "image/webp",
        "tif" | "tiff" => "image/tiff",
        "svg" => "image/svg+xml",
        _ => "application/octet-stream",
    }
}

#[allow(dead_code)]
fn _image_filename(index: usize, image: &ImageData) -> String {
    format!("image{}.{}", index + 1, ext_from_mime(image.mime.as_str()))
}

#[cfg(test)]
mod tests {
    use std::{
        fs,
        path::PathBuf,
        time::{SystemTime, UNIX_EPOCH},
    };

    use zip::{CompressionMethod, ZipWriter, write::SimpleFileOptions};

    use crate::document::model::{DocumentModel, Paragraph, ParagraphAlignment, ParagraphSpacing, Run, RunStyle};

    use super::*;

    fn unique_temp(name: &str) -> PathBuf {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        std::env::temp_dir().join(format!("doco-writer-{name}-{stamp}.docx"))
    }

    fn write_seed_docx(path: &Path) {
        let file = File::create(path).expect("create seed docx");
        let mut zip = ZipWriter::new(file);
        let options = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);

        zip.start_file("[Content_Types].xml", options)
            .expect("content types");
        zip.write_all(b"<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/><Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/></Types>")
            .expect("write content types");

        zip.start_file("_rels/.rels", options).expect("root rels");
        zip.write_all(root_rels_xml().as_bytes()).expect("write rels");

        zip.start_file("word/document.xml", options)
            .expect("document xml");
        zip.write_all(b"<?xml version=\"1.0\" encoding=\"UTF-8\"?><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>seed</w:t></w:r></w:p><w:sectPr/></w:body></w:document>")
            .expect("write document");

        zip.start_file("word/_rels/document.xml.rels", options)
            .expect("doc rels");
        zip.write_all(b"<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rIdHyper\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com\"/></Relationships>")
            .expect("write doc rels");

        zip.start_file("word/styles.xml", options).expect("styles");
        zip.write_all(b"<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:style w:type=\"paragraph\" w:styleId=\"Normal\"/></w:styles>")
            .expect("write styles");

        zip.start_file("word/theme/theme1.xml", options)
            .expect("theme");
        zip.write_all(b"<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>")
            .expect("write theme");

        zip.finish().expect("finish seed docx");
    }

    fn read_entry(path: &Path, target: &str) -> Vec<u8> {
        let file = File::open(path).expect("open docx");
        let mut archive = ZipArchive::new(file).expect("zip archive");
        let mut entry = archive.by_name(target).expect("entry exists");
        let mut buf = Vec::new();
        entry.read_to_end(&mut buf).expect("read entry");
        buf
    }

    #[test]
    fn roundtrip_preserves_unmodified_entries_and_updates_document() {
        let source = unique_temp("source");
        let output = unique_temp("out");
        write_seed_docx(&source);

        let mut doc = DocumentModel::default();
        doc.metadata.file_path = Some(source.clone());
        doc.content.push(Block::Paragraph(Paragraph {
            id: crate::document::model::BlockId(1),
            runs: vec![Run {
                text: "updated".to_string(),
                style: RunStyle::default(),
            }],
            alignment: ParagraphAlignment::Left,
            spacing: ParagraphSpacing::default(),
            indent: Default::default(),
            style_id: None,
        }));
        doc.images.insert(
            "img1".to_string(),
            ImageData {
                bytes: vec![0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A],
                mime: "image/png".to_string(),
                width: 1,
                height: 1,
            },
        );

        write_docx(&output, &doc).expect("write docx");

        let styles = read_entry(&output, "word/styles.xml");
        assert!(String::from_utf8_lossy(&styles).contains("styleId=\"Normal\""));
        assert!(!read_entry(&output, "word/theme/theme1.xml").is_empty());
        let xml = String::from_utf8_lossy(&read_entry(&output, "word/document.xml")).to_string();
        assert!(xml.contains("updated"));
        assert!(xml.contains("xmlns:pic="));

        let rels = String::from_utf8_lossy(&read_entry(&output, "word/_rels/document.xml.rels"))
            .to_string();
        assert!(rels.contains("rIdHyper"));
        assert!(rels.contains("rDocoImg1"));

        let content_types =
            String::from_utf8_lossy(&read_entry(&output, "[Content_Types].xml")).to_string();
        assert!(content_types.contains("Extension=\"png\""));
        assert!(content_types.contains("/word/styles.xml"));

        let _ = fs::remove_file(source);
        let _ = fs::remove_file(output);
    }

    #[test]
    fn fresh_docx_contains_minimal_required_parts() {
        let output = unique_temp("fresh");
        let mut doc = DocumentModel::default();
        doc.content.push(Block::Paragraph(Paragraph {
            id: crate::document::model::BlockId(1),
            runs: vec![Run {
                text: "hello".to_string(),
                style: RunStyle::default(),
            }],
            alignment: ParagraphAlignment::Left,
            spacing: ParagraphSpacing::default(),
            indent: Default::default(),
            style_id: None,
        }));

        write_docx(&output, &doc).expect("write fresh docx");
        assert!(!read_entry(&output, "_rels/.rels").is_empty());
        assert!(!read_entry(&output, "word/document.xml").is_empty());
        assert!(!read_entry(&output, "word/styles.xml").is_empty());
        assert!(!read_entry(&output, "[Content_Types].xml").is_empty());

        let _ = fs::remove_file(output);
    }
}
