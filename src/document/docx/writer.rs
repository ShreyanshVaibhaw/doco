use std::{
    fs::File,
    io::{self, Write},
    path::Path,
};

use zip::{CompressionMethod, ZipWriter, write::SimpleFileOptions};

use crate::document::model::{
    Block,
    DocumentModel,
    ImageData,
    ListType,
    Paragraph,
    Run,
};

pub fn write_docx(path: &Path, model: &DocumentModel) -> io::Result<()> {
    let file = File::create(path)?;
    let mut zip = ZipWriter::new(file);
    let options = SimpleFileOptions::default()
        .compression_method(CompressionMethod::Deflated)
        .unix_permissions(0o644);

    zip.start_file("[Content_Types].xml", options)?;
    zip.write_all(content_types_xml(model).as_bytes())?;

    zip.start_file("_rels/.rels", options)?;
    zip.write_all(root_rels_xml().as_bytes())?;

    zip.start_file("word/document.xml", options)?;
    zip.write_all(document_xml(model).as_bytes())?;

    zip.start_file("word/_rels/document.xml.rels", options)?;
    zip.write_all(document_rels_xml(model).as_bytes())?;

    for (idx, (_key, image)) in model.images.iter().enumerate() {
        let ext = ext_from_mime(image.mime.as_str());
        let entry = format!("word/media/image{}.{}", idx + 1, ext);
        zip.start_file(entry, options)?;
        zip.write_all(image.bytes.as_slice())?;
    }

    zip.finish()?;
    Ok(())
}

fn content_types_xml(model: &DocumentModel) -> String {
    let mut defaults = vec![
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>".to_string(),
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>".to_string(),
    ];

    let mut seen = std::collections::BTreeSet::new();
    for image in model.images.values() {
        let ext = ext_from_mime(image.mime.as_str());
        if seen.insert(ext.to_string()) {
            defaults.push(format!(
                "<Default Extension=\"{}\" ContentType=\"{}\"/>",
                ext, image.mime
            ));
        }
    }

    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n{}\n<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>\n</Types>",
        defaults.join("\n")
    )
}

fn root_rels_xml() -> &'static str {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>
</Relationships>"
}

fn document_rels_xml(model: &DocumentModel) -> String {
    let mut rels = Vec::new();
    for (idx, (_key, image)) in model.images.iter().enumerate() {
        let ext = ext_from_mime(image.mime.as_str());
        rels.push(format!(
            "<Relationship Id=\"rImg{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image{}.{}\"/>",
            idx + 1,
            idx + 1,
            ext
        ));
    }

    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n{}\n</Relationships>",
        rels.join("\n")
    )
}

fn document_xml(model: &DocumentModel) -> String {
    let mut body = String::new();

    for block in &model.content {
        match block {
            Block::Paragraph(p) => body.push_str(paragraph_xml(p).as_str()),
            Block::Heading(h) => {
                let mut p = Paragraph {
                    id: h.id,
                    runs: h.runs.clone(),
                    alignment: crate::document::model::ParagraphAlignment::Left,
                    spacing: crate::document::model::ParagraphSpacing::default(),
                    indent: crate::document::model::Indent::default(),
                    style_id: Some(format!("Heading{}", h.level)),
                };
                body.push_str(paragraph_xml(&p).as_str());
                p.runs.clear();
            }
            Block::CodeBlock(code) => {
                let p = Paragraph {
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
                body.push_str(paragraph_xml(&p).as_str());
            }
            Block::List(list) => {
                for (i, item) in list.items.iter().enumerate() {
                    let bullet = match list.list_type {
                        ListType::Bullet => "• ".to_string(),
                        ListType::Numbered => format!("{}. ", list.start_number + i as u32),
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
                    let p = Paragraph {
                        id: item.id,
                        runs,
                        alignment: crate::document::model::ParagraphAlignment::Left,
                        spacing: crate::document::model::ParagraphSpacing::default(),
                        indent: crate::document::model::Indent::default(),
                        style_id: None,
                    };
                    body.push_str(paragraph_xml(&p).as_str());
                }
            }
            Block::Table(table) => {
                body.push_str("<w:tbl>");
                for row in &table.rows {
                    body.push_str("<w:tr>");
                    for cell in &row.cells {
                        body.push_str("<w:tc><w:p>");
                        for block in &cell.blocks {
                            if let Block::Paragraph(p) = block {
                                for run in &p.runs {
                                    body.push_str(run_xml(run).as_str());
                                }
                            }
                        }
                        body.push_str("</w:p></w:tc>");
                    }
                    body.push_str("</w:tr>");
                }
                body.push_str("</w:tbl>");
            }
            Block::Image(img) => {
                body.push_str("<w:p><w:r><w:drawing>");
                body.push_str(format!("<wp:inline><wp:extent cx=\"{}\" cy=\"{}\"/></wp:inline>", (img.width.max(1.0) * 9525.0) as i64, (img.height.max(1.0) * 9525.0) as i64).as_str());
                body.push_str("</w:drawing></w:r></w:p>");
            }
            Block::HorizontalRule => {
                body.push_str("<w:p><w:r><w:t>---</w:t></w:r></w:p>");
            }
            Block::PageBreak => {
                body.push_str("<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>");
            }
            Block::BlockQuote(q) => {
                for nested in &q.blocks {
                    if let Block::Paragraph(p) = nested {
                        body.push_str(paragraph_xml(p).as_str());
                    }
                }
            }
        }
    }

    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\">
  <w:body>{}<w:sectPr/></w:body>
</w:document>",
        body
    )
}

fn paragraph_xml(p: &Paragraph) -> String {
    let mut out = String::new();
    out.push_str("<w:p>");

    if p.style_id.is_some() || !matches!(p.alignment, crate::document::model::ParagraphAlignment::Left) {
        out.push_str("<w:pPr>");
        if let Some(style) = &p.style_id {
            out.push_str(format!("<w:pStyle w:val=\"{}\"/>", escape_xml(style.as_str())).as_str());
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
        out.push_str("</w:pPr>");
    }

    for run in &p.runs {
        out.push_str(run_xml(run).as_str());
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
            out.push_str(format!("<w:rFonts w:ascii=\"{}\" w:hAnsi=\"{}\"/>", escape_xml(ff), escape_xml(ff)).as_str());
        }
        if let Some(color) = run.style.color {
            out.push_str(format!("<w:color w:val=\"{}\"/>", to_hex(color)).as_str());
        }
        out.push_str("</w:rPr>");
    }

    out.push_str(format!("<w:t xml:space=\"preserve\">{}</w:t>", escape_xml(run.text.as_str())).as_str());
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

#[allow(dead_code)]
fn _image_filename(index: usize, image: &ImageData) -> String {
    format!("image{}.{}", index + 1, ext_from_mime(image.mime.as_str()))
}
