use std::{
    collections::HashMap,
    fs::File,
    io::{Cursor, Read},
    path::Path,
};

use chrono::{DateTime, Utc};
use quick_xml::{
    Reader,
    events::{BytesStart, Event},
};
use zip::ZipArchive;

use crate::document::model::{
    Block,
    BlockId,
    DocumentModel,
    Heading,
    ImageAlignment,
    ImageBlock,
    ImageData,
    ImageDataRef,
    List,
    ListItem,
    ListType,
    Margins,
    Paragraph,
    ParagraphAlignment,
    ParagraphSpacing,
    Run,
    RunStyle,
    Table,
    TableBorders,
    TableCell,
    TableRow,
};

#[derive(Debug, Default)]
struct ParagraphBuilder {
    runs: Vec<Run>,
    style_id: Option<String>,
    alignment: ParagraphAlignment,
    spacing: ParagraphSpacing,
    indent: crate::document::model::Indent,
    list_type: Option<ListType>,
}

#[derive(Debug, Default)]
struct TableBuilder {
    rows: Vec<Vec<String>>,
    current_row: Vec<String>,
    current_cell_text: String,
    in_cell: bool,
}

#[derive(Debug, Default)]
struct ParsedRels {
    target_by_id: HashMap<String, String>,
    image_rel_ids: Vec<String>,
}

pub fn parse_docx(path: &Path) -> std::io::Result<DocumentModel> {
    let file = File::open(path)?;
    let mut archive = ZipArchive::new(file)?;

    let mut entries = HashMap::<String, Vec<u8>>::new();
    for i in 0..archive.len() {
        let mut f = archive.by_index(i)?;
        let name = f.name().to_string();
        let mut bytes = Vec::with_capacity(f.size() as usize);
        f.read_to_end(&mut bytes)?;
        entries.insert(name, bytes);
    }

    let mut document = DocumentModel::default();

    if let Some(core_xml) = entries.get("docProps/core.xml") {
        parse_core_metadata(core_xml, &mut document);
    }

    if let Some(styles_xml) = entries.get("word/styles.xml") {
        parse_styles(styles_xml, &mut document);
    }

    let rels = entries
        .get("word/_rels/document.xml.rels")
        .map(|v| parse_relationships(v.as_slice()))
        .unwrap_or_default();

    let content_types = entries
        .get("[Content_Types].xml")
        .map(|v| parse_content_types(v.as_slice()))
        .unwrap_or_default();

    if let Some(numbering_xml) = entries.get("word/numbering.xml") {
        parse_numbering(numbering_xml);
    }

    // Parse optional header/footer defensively to keep robustness across vendors.
    for optional in ["word/header1.xml", "word/footer1.xml"] {
        if let Some(xml) = entries.get(optional) {
            let _ = parse_plain_text_runs(xml);
        }
    }

    if let Some(document_xml) = entries.get("word/document.xml") {
        parse_document_xml(document_xml, &mut document, &rels)?;
    } else {
        return Err(std::io::Error::new(
            std::io::ErrorKind::InvalidData,
            "missing word/document.xml",
        ));
    }

    extract_images(&entries, &rels, &content_types, &mut document);
    document.metadata.file_path = Some(path.to_path_buf());
    document.dirty = false;

    Ok(document)
}

fn parse_core_metadata(xml: &[u8], doc: &mut DocumentModel) {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut current_tag: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                current_tag = Some(local_name(e.local_name().as_ref()));
            }
            Ok(Event::Text(t)) => {
                if let Some(tag) = &current_tag {
                    if let Ok(value) = t.decode() {
                        match tag.as_str() {
                            "title" => doc.metadata.title = value.to_string(),
                            "creator" => doc.metadata.author = value.to_string(),
                            "created" => doc.metadata.created = parse_datetime(value.as_ref()),
                            "modified" => doc.metadata.modified = parse_datetime(value.as_ref()),
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::End(_)) => current_tag = None,
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
}

fn parse_styles(xml: &[u8], doc: &mut DocumentModel) {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut current_style_id = String::new();
    let mut in_style = false;
    let mut current_name = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = local_name(e.local_name().as_ref());
                if name == "style" {
                    in_style = true;
                    current_style_id = attr_value(&e, "styleId", reader.decoder()).unwrap_or_default();
                    current_name = current_style_id.clone();
                } else if in_style && name == "name" {
                    if let Some(v) = attr_value(&e, "val", reader.decoder()) {
                        current_name = v;
                    }
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.local_name().as_ref()) == "style" && in_style {
                    doc.styles.styles.insert(
                        current_style_id.clone(),
                        crate::document::model::NamedStyle {
                            id: current_style_id.clone(),
                            name: current_name.clone(),
                            run_style: RunStyle::default(),
                            paragraph_style: None,
                        },
                    );
                    in_style = false;
                    current_style_id.clear();
                    current_name.clear();
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
}

fn parse_relationships(xml: &[u8]) -> ParsedRels {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut rels = ParsedRels::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                if local_name(e.local_name().as_ref()) == "Relationship" {
                    let id = attr_value(&e, "Id", reader.decoder()).unwrap_or_default();
                    let target = attr_value(&e, "Target", reader.decoder()).unwrap_or_default();
                    let rel_type = attr_value(&e, "Type", reader.decoder()).unwrap_or_default();
                    if !id.is_empty() && !target.is_empty() {
                        rels.target_by_id.insert(id.clone(), target);
                    }
                    if rel_type.ends_with("/image") {
                        rels.image_rel_ids.push(id);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    rels
}

fn parse_content_types(xml: &[u8]) -> HashMap<String, String> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = HashMap::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                let name = local_name(e.local_name().as_ref());
                if name == "Default" {
                    let ext = attr_value(&e, "Extension", reader.decoder()).unwrap_or_default();
                    let content_type =
                        attr_value(&e, "ContentType", reader.decoder()).unwrap_or_default();
                    if !ext.is_empty() && !content_type.is_empty() {
                        out.insert(ext.to_ascii_lowercase(), content_type);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    out
}

fn parse_numbering(_xml: &[u8]) {
    // Numbering definitions are intentionally tolerant here. We currently map list
    // blocks as bullet/numbered based on paragraph markers in document.xml.
}

fn parse_document_xml(
    xml: &[u8],
    doc: &mut DocumentModel,
    rels: &ParsedRels,
) -> std::io::Result<()> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut block_id = 1_u64;
    let mut paragraph: Option<ParagraphBuilder> = None;
    let mut run: Option<Run> = None;
    let mut in_text = false;
    let mut in_run_props = false;
    let mut in_paragraph_props = false;
    let mut current_table: Option<TableBuilder> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = local_name(e.local_name().as_ref());
                match name.as_str() {
                    "p" if current_table.is_none() => {
                        paragraph = Some(ParagraphBuilder::default());
                    }
                    "pPr" => in_paragraph_props = true,
                    "r" => run = Some(Run::default()),
                    "rPr" => in_run_props = true,
                    "t" => in_text = true,
                    "tbl" => current_table = Some(TableBuilder::default()),
                    "tr" => {
                        if let Some(tbl) = &mut current_table {
                            tbl.current_row.clear();
                        }
                    }
                    "tc" => {
                        if let Some(tbl) = &mut current_table {
                            tbl.current_cell_text.clear();
                            tbl.in_cell = true;
                        }
                    }
                    "pStyle" if in_paragraph_props => {
                        if let Some(p) = &mut paragraph {
                            p.style_id = attr_value(&e, "val", reader.decoder());
                        }
                    }
                    "jc" if in_paragraph_props => {
                        if let Some(p) = &mut paragraph {
                            if let Some(v) = attr_value(&e, "val", reader.decoder()) {
                                p.alignment = match v.as_str() {
                                    "center" => ParagraphAlignment::Center,
                                    "right" => ParagraphAlignment::Right,
                                    "both" => ParagraphAlignment::Justify,
                                    _ => ParagraphAlignment::Left,
                                };
                            }
                        }
                    }
                    "spacing" if in_paragraph_props => {
                        if let Some(p) = &mut paragraph {
                            p.spacing.before = twips_to_points(attr_value(&e, "before", reader.decoder()));
                            p.spacing.after = twips_to_points(attr_value(&e, "after", reader.decoder()));
                            p.spacing.line = twips_to_points(attr_value(&e, "line", reader.decoder()));
                        }
                    }
                    "ind" if in_paragraph_props => {
                        if let Some(p) = &mut paragraph {
                            p.indent.left = twips_to_points(attr_value(&e, "left", reader.decoder()));
                            p.indent.right = twips_to_points(attr_value(&e, "right", reader.decoder()));
                            p.indent.first_line =
                                twips_to_points(attr_value(&e, "firstLine", reader.decoder()));
                        }
                    }
                    "numPr" if in_paragraph_props => {
                        if let Some(p) = &mut paragraph {
                            p.list_type = Some(ListType::Numbered);
                        }
                    }
                    "numFmt" if in_paragraph_props => {
                        if let Some(p) = &mut paragraph {
                            if let Some(v) = attr_value(&e, "val", reader.decoder()) {
                                p.list_type = Some(if v == "bullet" {
                                    ListType::Bullet
                                } else {
                                    ListType::Numbered
                                });
                            }
                        }
                    }
                    "b" if in_run_props => {
                        if let Some(r) = &mut run {
                            r.style.bold = true;
                        }
                    }
                    "i" if in_run_props => {
                        if let Some(r) = &mut run {
                            r.style.italic = true;
                        }
                    }
                    "u" if in_run_props => {
                        if let Some(r) = &mut run {
                            r.style.underline = true;
                        }
                    }
                    "strike" if in_run_props => {
                        if let Some(r) = &mut run {
                            r.style.strikethrough = true;
                        }
                    }
                    "rFonts" if in_run_props => {
                        if let Some(r) = &mut run {
                            r.style.font_family = attr_value(&e, "ascii", reader.decoder())
                                .or_else(|| attr_value(&e, "cs", reader.decoder()));
                        }
                    }
                    "sz" if in_run_props => {
                        if let Some(r) = &mut run {
                            if let Some(v) = attr_value(&e, "val", reader.decoder()) {
                                if let Ok(half_points) = v.parse::<f32>() {
                                    r.style.font_size = Some(half_points / 2.0);
                                }
                            }
                        }
                    }
                    "color" if in_run_props => {
                        if let Some(r) = &mut run {
                            if let Some(v) = attr_value(&e, "val", reader.decoder()) {
                                r.style.color = parse_hex_color(&v);
                            }
                        }
                    }
                    "vertAlign" if in_run_props => {
                        if let Some(r) = &mut run {
                            if let Some(v) = attr_value(&e, "val", reader.decoder()) {
                                r.style.superscript = v == "superscript";
                                r.style.subscript = v == "subscript";
                            }
                        }
                    }
                    "br" => {
                        let break_type = attr_value(&e, "type", reader.decoder()).unwrap_or_default();
                        if break_type == "page" {
                            doc.content.push(Block::PageBreak);
                        } else if let Some(r) = &mut run {
                            r.text.push('\n');
                        }
                    }
                    "blip" => {
                        if let Some(embedded) = attr_value(&e, "embed", reader.decoder()) {
                            if rels.target_by_id.contains_key(&embedded) {
                                doc.content.push(Block::Image(ImageBlock {
                                    id: next_block_id(&mut block_id),
                                    data: ImageDataRef::Key(embedded.clone()),
                                    original_width: 0,
                                    original_height: 0,
                                    caption: None,
                                    border: None,
                                    crop: None,
                                    key: embedded,
                                    alt_text: String::new(),
                                    source_path: None,
                                    width: 120.0,
                                    height: 120.0,
                                    alignment: ImageAlignment::Left,
                                }));
                            }
                        }
                    }
                    "pgSz" => {
                        let w = twips_to_points(attr_value(&e, "w", reader.decoder()));
                        let h = twips_to_points(attr_value(&e, "h", reader.decoder()));
                        if w > 0.0 && h > 0.0 {
                            doc.metadata.page_size = crate::document::model::PageSize::Custom {
                                width_points: w,
                                height_points: h,
                            };
                        }
                    }
                    "pgMar" => {
                        doc.metadata.margins = Margins {
                            top: twips_to_points(attr_value(&e, "top", reader.decoder())),
                            right: twips_to_points(attr_value(&e, "right", reader.decoder())),
                            bottom: twips_to_points(attr_value(&e, "bottom", reader.decoder())),
                            left: twips_to_points(attr_value(&e, "left", reader.decoder())),
                        };
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                let name = local_name(e.local_name().as_ref());
                match name.as_str() {
                    "pStyle" if in_paragraph_props => {
                        if let Some(p) = &mut paragraph {
                            p.style_id = attr_value(&e, "val", reader.decoder());
                        }
                    }
                    "jc" if in_paragraph_props => {
                        if let Some(p) = &mut paragraph {
                            if let Some(v) = attr_value(&e, "val", reader.decoder()) {
                                p.alignment = match v.as_str() {
                                    "center" => ParagraphAlignment::Center,
                                    "right" => ParagraphAlignment::Right,
                                    "both" => ParagraphAlignment::Justify,
                                    _ => ParagraphAlignment::Left,
                                };
                            }
                        }
                    }
                    "br" => {
                        let break_type = attr_value(&e, "type", reader.decoder()).unwrap_or_default();
                        if break_type == "page" {
                            doc.content.push(Block::PageBreak);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(t)) => {
                let text = match t.decode() {
                    Ok(v) => v.into_owned(),
                    Err(_) => String::new(),
                };

                if in_text {
                    if let Some(r) = &mut run {
                        r.text.push_str(&text);
                    }
                }
                if let Some(tbl) = &mut current_table {
                    if tbl.in_cell {
                        tbl.current_cell_text.push_str(&text);
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = local_name(e.local_name().as_ref());
                match name.as_str() {
                    "pPr" => in_paragraph_props = false,
                    "rPr" => in_run_props = false,
                    "t" => in_text = false,
                    "r" => {
                        if let (Some(p), Some(r)) = (&mut paragraph, run.take()) {
                            if !r.text.is_empty() {
                                p.runs.push(r);
                            }
                        }
                    }
                    "p" if current_table.is_none() => {
                        if let Some(p) = paragraph.take() {
                            if p.runs.is_empty() {
                                continue;
                            }

                            let block_id_now = next_block_id(&mut block_id);
                            let is_heading = p
                                .style_id
                                .as_ref()
                                .and_then(|s| parse_heading_level(s));

                            if let Some(level) = is_heading {
                                doc.content.push(Block::Heading(Heading {
                                    level,
                                    runs: p.runs,
                                    id: block_id_now,
                                }));
                            } else if let Some(list_type) = p.list_type {
                                let para_block = Block::Paragraph(Paragraph {
                                    id: block_id_now,
                                    runs: p.runs,
                                    alignment: p.alignment,
                                    spacing: p.spacing,
                                    indent: p.indent,
                                    style_id: p.style_id,
                                });
                                doc.content.push(Block::List(List {
                                    items: vec![ListItem {
                                        id: block_id_now,
                                        content: vec![para_block],
                                        checked: None,
                                        children: vec![],
                                    }],
                                    list_type,
                                    start_number: 1,
                                }));
                            } else {
                                doc.content.push(Block::Paragraph(Paragraph {
                                    id: block_id_now,
                                    runs: p.runs,
                                    alignment: p.alignment,
                                    spacing: p.spacing,
                                    indent: p.indent,
                                    style_id: p.style_id,
                                }));
                            }
                        }
                    }
                    "tc" => {
                        if let Some(tbl) = &mut current_table {
                            tbl.current_row.push(tbl.current_cell_text.trim().to_string());
                            tbl.current_cell_text.clear();
                            tbl.in_cell = false;
                        }
                    }
                    "tr" => {
                        if let Some(tbl) = &mut current_table {
                            if !tbl.current_row.is_empty() {
                                tbl.rows.push(std::mem::take(&mut tbl.current_row));
                            }
                        }
                    }
                    "tbl" => {
                        if let Some(tbl) = current_table.take() {
                            let rows = tbl
                                .rows
                                .into_iter()
                                .map(|row| TableRow {
                                    cells: row
                                        .into_iter()
                                        .map(|text| TableCell {
                                            blocks: if text.is_empty() {
                                                vec![]
                                            } else {
                                                vec![Block::Paragraph(Paragraph {
                                                    id: next_block_id(&mut block_id),
                                                    runs: vec![Run {
                                                        text,
                                                        style: RunStyle::default(),
                                                    }],
                                                    alignment: ParagraphAlignment::Left,
                                                    spacing: ParagraphSpacing::default(),
                                                    indent: crate::document::model::Indent::default(),
                                                    style_id: None,
                                                })]
                                            },
                                            rowspan: 1,
                                            colspan: 1,
                                            background: None,
                                        })
                                        .collect(),
                                })
                                .collect::<Vec<_>>();

                            let col_count = rows.first().map(|r| r.cells.len()).unwrap_or(1);
                            doc.content.push(Block::Table(Table {
                                id: next_block_id(&mut block_id),
                                rows,
                                column_widths: vec![120.0; col_count],
                                borders: TableBorders::default(),
                                style: crate::document::model::TableStylePreset::Grid,
                                cell_padding: 4.0,
                                header_row: false,
                                alternating_rows: false,
                            }));
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(err) => {
                return Err(std::io::Error::new(
                    std::io::ErrorKind::InvalidData,
                    format!("document.xml parse error: {err}"),
                ));
            }
            _ => {}
        }
        buf.clear();
    }

    if doc.content.is_empty() {
        let fallback = parse_plain_text_runs(xml);
        if !fallback.is_empty() {
            doc.content.push(Block::Paragraph(Paragraph {
                id: next_block_id(&mut block_id),
                runs: vec![Run {
                    text: fallback,
                    style: RunStyle::default(),
                }],
                alignment: ParagraphAlignment::Left,
                spacing: ParagraphSpacing::default(),
                indent: crate::document::model::Indent::default(),
                style_id: None,
            }));
        }
    }

    Ok(())
}

fn extract_images(
    entries: &HashMap<String, Vec<u8>>,
    rels: &ParsedRels,
    content_types: &HashMap<String, String>,
    doc: &mut DocumentModel,
) {
    for rel_id in &rels.image_rel_ids {
        let Some(target) = rels.target_by_id.get(rel_id) else {
            continue;
        };

        let normalized = if target.starts_with("word/") {
            target.to_string()
        } else {
            format!("word/{}", target.trim_start_matches("./"))
        };

        if let Some(bytes) = entries.get(&normalized) {
            let ext = normalized
                .rsplit('.')
                .next()
                .unwrap_or_default()
                .to_ascii_lowercase();
            let mime = content_types
                .get(&ext)
                .cloned()
                .unwrap_or_else(|| default_mime_for_ext(&ext).to_string());

            doc.images.insert(
                rel_id.clone(),
                ImageData {
                    bytes: bytes.clone(),
                    mime,
                    width: 0,
                    height: 0,
                },
            );
        }
    }
}

fn parse_plain_text_runs(xml: &[u8]) -> String {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut in_text = false;
    let mut text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if local_name(e.local_name().as_ref()) == "t" {
                    in_text = true;
                }
            }
            Ok(Event::Text(t)) if in_text => {
                if let Ok(v) = t.decode() {
                    text.push_str(v.as_ref());
                    text.push(' ');
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.local_name().as_ref()) == "t" {
                    in_text = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    text.trim().to_string()
}

fn parse_heading_level(style_id: &str) -> Option<u8> {
    let lower = style_id.to_ascii_lowercase();
    if let Some(num) = lower.strip_prefix("heading") {
        return num.parse::<u8>().ok().filter(|v| (1..=6).contains(v));
    }
    None
}

fn parse_datetime(raw: &str) -> Option<DateTime<Utc>> {
    chrono::DateTime::parse_from_rfc3339(raw)
        .ok()
        .map(|dt| dt.with_timezone(&Utc))
}

fn parse_hex_color(value: &str) -> Option<crate::ui::Color> {
    let hex = value.trim_start_matches('#');
    if hex.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
    let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
    let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
    Some(crate::ui::Color::rgb(
        (r as f32) / 255.0,
        (g as f32) / 255.0,
        (b as f32) / 255.0,
    ))
}

fn twips_to_points(value: Option<String>) -> f32 {
    value
        .and_then(|v| v.parse::<f32>().ok())
        .map(|v| v / 20.0)
        .unwrap_or(0.0)
}

fn local_name(bytes: &[u8]) -> String {
    let full = std::str::from_utf8(bytes).unwrap_or_default();
    full.rsplit(':').next().unwrap_or(full).to_string()
}

fn attr_value(
    event: &BytesStart<'_>,
    key_suffix: &str,
    decoder: quick_xml::encoding::Decoder,
) -> Option<String> {
    event
        .attributes()
        .flatten()
        .find_map(|a| {
            let key = std::str::from_utf8(a.key.as_ref()).ok()?;
            if key.rsplit(':').next() == Some(key_suffix) {
                a.decode_and_unescape_value(decoder)
                    .ok()
                    .map(|v| v.to_string())
            } else {
                None
            }
        })
}

fn next_block_id(counter: &mut u64) -> BlockId {
    let id = *counter;
    *counter += 1;
    BlockId(id)
}

fn default_mime_for_ext(ext: &str) -> &'static str {
    match ext {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "bmp" => "image/bmp",
        "gif" => "image/gif",
        "webp" => "image/webp",
        "tif" | "tiff" => "image/tiff",
        "emf" => "image/emf",
        "wmf" => "image/wmf",
        _ => "application/octet-stream",
    }
}
