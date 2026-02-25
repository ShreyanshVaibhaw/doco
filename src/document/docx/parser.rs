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
    Indent,
    List,
    ListItem,
    ListType,
    Margins,
    NamedStyle,
    PageSize,
    Paragraph,
    ParagraphAlignment,
    ParagraphStyle,
    ParagraphSpacing,
    Run,
    RunStyle,
    StyleSheet,
    Table,
    TableBorders,
    TableCell,
    TableRow,
};
use crate::document::DocumentFormat;

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

#[derive(Debug, Clone, Default)]
struct NumberingMap {
    list_type_by_num_id: HashMap<String, ListType>,
}

#[derive(Debug, Clone, Default)]
struct RawStyle {
    id: String,
    name: String,
    based_on: Option<String>,
    run_patch: RunPatch,
    paragraph_patch: ParagraphPatch,
}

#[derive(Debug, Clone, Default)]
struct RunPatch {
    font_family: Option<String>,
    font_size: Option<f32>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    strikethrough: Option<bool>,
    color: Option<crate::ui::Color>,
    background: Option<crate::ui::Color>,
    superscript: Option<bool>,
    subscript: Option<bool>,
}

#[derive(Debug, Clone, Default)]
struct ParagraphPatch {
    alignment: Option<ParagraphAlignment>,
    spacing_before: Option<f32>,
    spacing_after: Option<f32>,
    spacing_line: Option<f32>,
    indent_left: Option<f32>,
    indent_right: Option<f32>,
    indent_first_line: Option<f32>,
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
        document.styles = parse_styles(styles_xml);
    }

    let rels = entries
        .get("word/_rels/document.xml.rels")
        .map(|v| parse_relationships(v.as_slice()))
        .unwrap_or_default();

    let content_types = entries
        .get("[Content_Types].xml")
        .map(|v| parse_content_types(v.as_slice()))
        .unwrap_or_default();

    let numbering = entries
        .get("word/numbering.xml")
        .map(|v| parse_numbering(v.as_slice()))
        .unwrap_or_default();

    let mut header_footer_fragments = Vec::new();
    for optional in ["word/header1.xml", "word/footer1.xml"] {
        if let Some(xml) = entries.get(optional) {
            let parsed = parse_plain_text_runs(xml);
            if !parsed.is_empty() {
                header_footer_fragments.push(parsed);
            }
        }
    }

    if let Some(document_xml) = entries.get("word/document.xml") {
        parse_document_xml(document_xml, &mut document, &rels, &numbering)?;
    } else {
        return Err(std::io::Error::new(
            std::io::ErrorKind::InvalidData,
            "missing word/document.xml",
        ));
    }

    extract_images(&entries, &rels, &content_types, &mut document);
    apply_embedded_image_dimensions(&mut document);
    if !header_footer_fragments.is_empty() && document.metadata.title.is_empty() {
        document.metadata.title = header_footer_fragments.join(" | ");
    }
    document.metadata.format = DocumentFormat::Docx;
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

fn parse_styles(xml: &[u8]) -> StyleSheet {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut raw_styles = HashMap::<String, RawStyle>::new();
    let mut current: Option<RawStyle> = None;
    let mut in_rpr = false;
    let mut in_ppr = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name = local_name(e.local_name().as_ref());
                match name.as_str() {
                    "style" => {
                        let style_id = attr_value(&e, "styleId", reader.decoder()).unwrap_or_default();
                        current = Some(RawStyle {
                            id: style_id.clone(),
                            name: style_id,
                            ..RawStyle::default()
                        });
                    }
                    "name" => {
                        if let Some(style) = &mut current {
                            if let Some(v) = attr_value(&e, "val", reader.decoder()) {
                                style.name = v;
                            }
                        }
                    }
                    "basedOn" => {
                        if let Some(style) = &mut current {
                            style.based_on = attr_value(&e, "val", reader.decoder());
                        }
                    }
                    "rPr" => in_rpr = true,
                    "pPr" => in_ppr = true,
                    "b" if in_rpr => {
                        if let Some(style) = &mut current {
                            style.run_patch.bold = Some(true);
                        }
                    }
                    "i" if in_rpr => {
                        if let Some(style) = &mut current {
                            style.run_patch.italic = Some(true);
                        }
                    }
                    "u" if in_rpr => {
                        if let Some(style) = &mut current {
                            style.run_patch.underline = Some(true);
                        }
                    }
                    "strike" if in_rpr => {
                        if let Some(style) = &mut current {
                            style.run_patch.strikethrough = Some(true);
                        }
                    }
                    "rFonts" if in_rpr => {
                        if let Some(style) = &mut current {
                            style.run_patch.font_family = attr_value(&e, "ascii", reader.decoder())
                                .or_else(|| attr_value(&e, "cs", reader.decoder()));
                        }
                    }
                    "sz" if in_rpr => {
                        if let Some(style) = &mut current {
                            if let Some(v) = attr_value(&e, "val", reader.decoder()) {
                                if let Ok(half_points) = v.parse::<f32>() {
                                    style.run_patch.font_size = Some(half_points / 2.0);
                                }
                            }
                        }
                    }
                    "color" if in_rpr => {
                        if let Some(style) = &mut current {
                            if let Some(v) = attr_value(&e, "val", reader.decoder()) {
                                style.run_patch.color = parse_hex_color(&v);
                            }
                        }
                    }
                    "highlight" if in_rpr => {
                        if let Some(style) = &mut current {
                            if let Some(v) = attr_value(&e, "val", reader.decoder()) {
                                style.run_patch.background = parse_named_highlight(&v);
                            }
                        }
                    }
                    "vertAlign" if in_rpr => {
                        if let Some(style) = &mut current {
                            if let Some(v) = attr_value(&e, "val", reader.decoder()) {
                                style.run_patch.superscript = Some(v == "superscript");
                                style.run_patch.subscript = Some(v == "subscript");
                            }
                        }
                    }
                    "jc" if in_ppr => {
                        if let Some(style) = &mut current {
                            style.paragraph_patch.alignment =
                                attr_value(&e, "val", reader.decoder()).map(|v| match v.as_str() {
                                    "center" => ParagraphAlignment::Center,
                                    "right" => ParagraphAlignment::Right,
                                    "both" => ParagraphAlignment::Justify,
                                    _ => ParagraphAlignment::Left,
                                });
                        }
                    }
                    "spacing" if in_ppr => {
                        if let Some(style) = &mut current {
                            if let Some(v) = attr_value(&e, "before", reader.decoder()) {
                                style.paragraph_patch.spacing_before = v.parse::<f32>().ok().map(|n| n / 20.0);
                            }
                            if let Some(v) = attr_value(&e, "after", reader.decoder()) {
                                style.paragraph_patch.spacing_after = v.parse::<f32>().ok().map(|n| n / 20.0);
                            }
                            if let Some(v) = attr_value(&e, "line", reader.decoder()) {
                                style.paragraph_patch.spacing_line = v.parse::<f32>().ok().map(|n| n / 20.0);
                            }
                        }
                    }
                    "ind" if in_ppr => {
                        if let Some(style) = &mut current {
                            if let Some(v) = attr_value(&e, "left", reader.decoder()) {
                                style.paragraph_patch.indent_left = v.parse::<f32>().ok().map(|n| n / 20.0);
                            }
                            if let Some(v) = attr_value(&e, "right", reader.decoder()) {
                                style.paragraph_patch.indent_right = v.parse::<f32>().ok().map(|n| n / 20.0);
                            }
                            if let Some(v) = attr_value(&e, "firstLine", reader.decoder()) {
                                style.paragraph_patch.indent_first_line = v.parse::<f32>().ok().map(|n| n / 20.0);
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                match local_name(e.local_name().as_ref()).as_str() {
                    "rPr" => in_rpr = false,
                    "pPr" => in_ppr = false,
                    "style" => {
                        if let Some(style) = current.take() {
                            raw_styles.insert(style.id.clone(), style);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(err) => {
                eprintln!("warning: styles.xml parse warning: {err}");
                break;
            }
            _ => {}
        }
        buf.clear();
    }

    resolve_styles(raw_styles)
}

fn resolve_styles(raw_styles: HashMap<String, RawStyle>) -> StyleSheet {
    fn resolve_one(
        style_id: &str,
        raw_styles: &HashMap<String, RawStyle>,
        cache: &mut HashMap<String, NamedStyle>,
        visiting: &mut Vec<String>,
    ) -> Option<NamedStyle> {
        if let Some(existing) = cache.get(style_id) {
            return Some(existing.clone());
        }
        if visiting.iter().any(|s| s == style_id) {
            return None;
        }
        let raw = raw_styles.get(style_id)?.clone();
        visiting.push(style_id.to_string());

        let mut resolved = if let Some(parent_id) = &raw.based_on {
            resolve_one(parent_id, raw_styles, cache, visiting).unwrap_or(NamedStyle {
                id: style_id.to_string(),
                name: raw.name.clone(),
                run_style: RunStyle::default(),
                paragraph_style: None,
            })
        } else {
            NamedStyle {
                id: style_id.to_string(),
                name: raw.name.clone(),
                run_style: RunStyle::default(),
                paragraph_style: None,
            }
        };
        resolved.id = raw.id.clone();
        resolved.name = if raw.name.is_empty() { raw.id.clone() } else { raw.name.clone() };
        apply_run_patch(&mut resolved.run_style, &raw.run_patch);
        apply_paragraph_patch(&mut resolved.paragraph_style, &raw.paragraph_patch);

        visiting.pop();
        cache.insert(style_id.to_string(), resolved.clone());
        Some(resolved)
    }

    let mut cache = HashMap::new();
    for style_id in raw_styles.keys() {
        let mut visiting = Vec::new();
        let _ = resolve_one(style_id, &raw_styles, &mut cache, &mut visiting);
    }
    StyleSheet { styles: cache }
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

fn parse_numbering(xml: &[u8]) -> NumberingMap {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut abstract_to_type = HashMap::<String, ListType>::new();
    let mut numbering = NumberingMap::default();
    let mut current_abstract_id: Option<String> = None;
    let mut current_num_id: Option<String> = None;
    let mut current_num_abstract_ref: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name = local_name(e.local_name().as_ref());
                match name.as_str() {
                    "abstractNum" => {
                        current_abstract_id = attr_value(&e, "abstractNumId", reader.decoder());
                    }
                    "numFmt" => {
                        if let Some(abs_id) = &current_abstract_id {
                            let list_type = attr_value(&e, "val", reader.decoder())
                                .map(|v| if v == "bullet" { ListType::Bullet } else { ListType::Numbered })
                                .unwrap_or(ListType::Numbered);
                            abstract_to_type.insert(abs_id.clone(), list_type);
                        }
                    }
                    "num" => {
                        current_num_id = attr_value(&e, "numId", reader.decoder());
                        current_num_abstract_ref = None;
                    }
                    "abstractNumId" => {
                        if current_num_id.is_some() {
                            current_num_abstract_ref = attr_value(&e, "val", reader.decoder());
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => match local_name(e.local_name().as_ref()).as_str() {
                "abstractNum" => current_abstract_id = None,
                "num" => {
                    if let (Some(num_id), Some(abs_ref)) = (&current_num_id, &current_num_abstract_ref) {
                        if let Some(list_type) = abstract_to_type.get(abs_ref) {
                            numbering
                                .list_type_by_num_id
                                .insert(num_id.clone(), list_type.clone());
                        }
                    }
                    current_num_id = None;
                    current_num_abstract_ref = None;
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(err) => {
                eprintln!("warning: numbering.xml parse warning: {err}");
                break;
            }
            _ => {}
        }
        buf.clear();
    }

    numbering
}

fn parse_document_xml(
    xml: &[u8],
    doc: &mut DocumentModel,
    rels: &ParsedRels,
    numbering: &NumberingMap,
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
    let mut in_hyperlink = false;
    let mut current_table: Option<TableBuilder> = None;
    let mut pending_image_size_points: Option<(f32, f32)> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = local_name(e.local_name().as_ref());
                match name.as_str() {
                    "p" if current_table.is_none() => {
                        paragraph = Some(ParagraphBuilder::default());
                    }
                    "pPr" => in_paragraph_props = true,
                    "hyperlink" => in_hyperlink = true,
                    "r" => {
                        let mut next_run = Run::default();
                        if in_hyperlink {
                            next_run.style.underline = true;
                            next_run.style.color = Some(crate::ui::Color::rgb(0.12, 0.39, 0.91));
                        }
                        run = Some(next_run);
                    }
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
                    "numId" if in_paragraph_props => {
                        if let Some(p) = &mut paragraph {
                            if let Some(num_id) = attr_value(&e, "val", reader.decoder()) {
                                p.list_type = numbering
                                    .list_type_by_num_id
                                    .get(&num_id)
                                    .cloned()
                                    .or(Some(ListType::Numbered));
                            }
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
                    "extent" => {
                        let cx = attr_value(&e, "cx", reader.decoder()).and_then(|v| v.parse::<f32>().ok());
                        let cy = attr_value(&e, "cy", reader.decoder()).and_then(|v| v.parse::<f32>().ok());
                        if let (Some(cx), Some(cy)) = (cx, cy) {
                            pending_image_size_points = Some((emu_to_points(cx), emu_to_points(cy)));
                        }
                    }
                    "blip" => {
                        if let Some(embedded) = attr_value(&e, "embed", reader.decoder()) {
                            if rels.target_by_id.contains_key(&embedded) {
                                let (image_w, image_h) = pending_image_size_points.unwrap_or((120.0, 120.0));
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
                                    width: image_w.max(24.0),
                                    height: image_h.max(24.0),
                                    alignment: ImageAlignment::Inline,
                                }));
                                pending_image_size_points = None;
                            }
                        }
                    }
                    "pgSz" => {
                        let w = twips_to_points(attr_value(&e, "w", reader.decoder()));
                        let h = twips_to_points(attr_value(&e, "h", reader.decoder()));
                        if w > 0.0 && h > 0.0 {
                            doc.metadata.page_size = PageSize::Custom {
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
                    "numId" if in_paragraph_props => {
                        if let Some(p) = &mut paragraph {
                            if let Some(num_id) = attr_value(&e, "val", reader.decoder()) {
                                p.list_type = numbering
                                    .list_type_by_num_id
                                    .get(&num_id)
                                    .cloned()
                                    .or(Some(ListType::Numbered));
                            }
                        }
                    }
                    "extent" => {
                        let cx = attr_value(&e, "cx", reader.decoder()).and_then(|v| v.parse::<f32>().ok());
                        let cy = attr_value(&e, "cy", reader.decoder()).and_then(|v| v.parse::<f32>().ok());
                        if let (Some(cx), Some(cy)) = (cx, cy) {
                            pending_image_size_points = Some((emu_to_points(cx), emu_to_points(cy)));
                        }
                    }
                    "blip" => {
                        if let Some(embedded) = attr_value(&e, "embed", reader.decoder()) {
                            if rels.target_by_id.contains_key(&embedded) {
                                let (image_w, image_h) = pending_image_size_points.unwrap_or((120.0, 120.0));
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
                                    width: image_w.max(24.0),
                                    height: image_h.max(24.0),
                                    alignment: ImageAlignment::Inline,
                                }));
                                pending_image_size_points = None;
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
                    "hyperlink" => in_hyperlink = false,
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

                            let mut paragraph_block = Paragraph {
                                id: block_id_now,
                                runs: p.runs,
                                alignment: p.alignment,
                                spacing: p.spacing,
                                indent: p.indent,
                                style_id: p.style_id,
                            };
                            apply_resolved_style_to_paragraph(&mut paragraph_block, &doc.styles);

                            if let Some(level) = is_heading {
                                doc.content.push(Block::Heading(Heading {
                                    level,
                                    runs: paragraph_block.runs,
                                    id: block_id_now,
                                }));
                            } else if let Some(list_type) = p.list_type {
                                let para_block = Block::Paragraph(paragraph_block);
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
                                doc.content.push(Block::Paragraph(paragraph_block));
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
                                                    indent: Indent::default(),
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
                            let row_count = rows.len();
                            doc.content.push(Block::Table(Table {
                                id: next_block_id(&mut block_id),
                                rows,
                                column_widths: vec![120.0; col_count],
                                row_heights: vec![28.0; row_count],
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
                eprintln!("warning: document.xml parse warning: {err}");
                break;
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
                indent: Indent::default(),
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
            let (width, height) = decode_image_dimensions(bytes, &ext);

            doc.images.insert(
                rel_id.clone(),
                ImageData {
                    bytes: bytes.clone(),
                    mime,
                    width,
                    height,
                },
            );
        }
    }
}

fn apply_run_patch(style: &mut RunStyle, patch: &RunPatch) {
    if let Some(font_family) = &patch.font_family {
        style.font_family = Some(font_family.clone());
    }
    if let Some(font_size) = patch.font_size {
        style.font_size = Some(font_size);
    }
    if let Some(bold) = patch.bold {
        style.bold = bold;
    }
    if let Some(italic) = patch.italic {
        style.italic = italic;
    }
    if let Some(underline) = patch.underline {
        style.underline = underline;
    }
    if let Some(strikethrough) = patch.strikethrough {
        style.strikethrough = strikethrough;
    }
    if let Some(color) = patch.color {
        style.color = Some(color);
    }
    if let Some(background) = patch.background {
        style.background = Some(background);
    }
    if let Some(superscript) = patch.superscript {
        style.superscript = superscript;
    }
    if let Some(subscript) = patch.subscript {
        style.subscript = subscript;
    }
}

fn apply_paragraph_patch(paragraph_style: &mut Option<ParagraphStyle>, patch: &ParagraphPatch) {
    let has_any = patch.alignment.is_some()
        || patch.spacing_before.is_some()
        || patch.spacing_after.is_some()
        || patch.spacing_line.is_some()
        || patch.indent_left.is_some()
        || patch.indent_right.is_some()
        || patch.indent_first_line.is_some();
    if !has_any {
        return;
    }

    let style = paragraph_style.get_or_insert_with(ParagraphStyle::default);
    if let Some(alignment) = &patch.alignment {
        style.alignment = alignment.clone();
    }
    if let Some(v) = patch.spacing_before {
        style.spacing.before = v;
    }
    if let Some(v) = patch.spacing_after {
        style.spacing.after = v;
    }
    if let Some(v) = patch.spacing_line {
        style.spacing.line = v;
    }
    if let Some(v) = patch.indent_left {
        style.indent.left = v;
    }
    if let Some(v) = patch.indent_right {
        style.indent.right = v;
    }
    if let Some(v) = patch.indent_first_line {
        style.indent.first_line = v;
    }
}

fn apply_resolved_style_to_paragraph(paragraph: &mut Paragraph, stylesheet: &StyleSheet) {
    let Some(style_id) = &paragraph.style_id else {
        return;
    };
    let Some(named) = stylesheet.styles.get(style_id) else {
        return;
    };

    if let Some(paragraph_style) = &named.paragraph_style {
        if matches!(paragraph.alignment, ParagraphAlignment::Left) {
            paragraph.alignment = paragraph_style.alignment.clone();
        }
        if paragraph.spacing.before == 0.0 {
            paragraph.spacing.before = paragraph_style.spacing.before;
        }
        if paragraph.spacing.after == 0.0 {
            paragraph.spacing.after = paragraph_style.spacing.after;
        }
        if paragraph.spacing.line == 0.0 {
            paragraph.spacing.line = paragraph_style.spacing.line;
        }
        if paragraph.indent.left == 0.0 {
            paragraph.indent.left = paragraph_style.indent.left;
        }
        if paragraph.indent.right == 0.0 {
            paragraph.indent.right = paragraph_style.indent.right;
        }
        if paragraph.indent.first_line == 0.0 {
            paragraph.indent.first_line = paragraph_style.indent.first_line;
        }
    }

    for run in &mut paragraph.runs {
        if run.style.font_family.is_none() {
            run.style.font_family = named.run_style.font_family.clone();
        }
        if run.style.font_size.is_none() {
            run.style.font_size = named.run_style.font_size;
        }
        if !run.style.bold {
            run.style.bold = named.run_style.bold;
        }
        if !run.style.italic {
            run.style.italic = named.run_style.italic;
        }
        if !run.style.underline {
            run.style.underline = named.run_style.underline;
        }
        if !run.style.strikethrough {
            run.style.strikethrough = named.run_style.strikethrough;
        }
        if run.style.color.is_none() {
            run.style.color = named.run_style.color;
        }
        if run.style.background.is_none() {
            run.style.background = named.run_style.background;
        }
        if !run.style.superscript {
            run.style.superscript = named.run_style.superscript;
        }
        if !run.style.subscript {
            run.style.subscript = named.run_style.subscript;
        }
    }
}

fn parse_named_highlight(name: &str) -> Option<crate::ui::Color> {
    match name.to_ascii_lowercase().as_str() {
        "yellow" => Some(crate::ui::Color::rgb(1.0, 0.95, 0.36)),
        "green" => Some(crate::ui::Color::rgb(0.63, 0.91, 0.67)),
        "cyan" => Some(crate::ui::Color::rgb(0.58, 0.89, 0.96)),
        "magenta" => Some(crate::ui::Color::rgb(0.96, 0.69, 0.93)),
        "blue" => Some(crate::ui::Color::rgb(0.62, 0.74, 0.97)),
        "red" => Some(crate::ui::Color::rgb(0.95, 0.58, 0.58)),
        "darkyellow" => Some(crate::ui::Color::rgb(0.84, 0.71, 0.33)),
        "darkgreen" => Some(crate::ui::Color::rgb(0.35, 0.64, 0.38)),
        "darkcyan" => Some(crate::ui::Color::rgb(0.31, 0.63, 0.66)),
        "darkmagenta" => Some(crate::ui::Color::rgb(0.62, 0.42, 0.62)),
        "darkblue" => Some(crate::ui::Color::rgb(0.38, 0.44, 0.74)),
        "darkred" => Some(crate::ui::Color::rgb(0.73, 0.34, 0.34)),
        _ => None,
    }
}

fn decode_image_dimensions(bytes: &[u8], ext: &str) -> (u32, u32) {
    if matches!(ext, "emf" | "wmf") {
        return (0, 0);
    }
    image::load_from_memory(bytes)
        .map(|img| (img.width(), img.height()))
        .unwrap_or((0, 0))
}

fn apply_embedded_image_dimensions(doc: &mut DocumentModel) {
    fn walk_blocks(blocks: &mut [Block], images: &HashMap<String, ImageData>) {
        for block in blocks {
            match block {
                Block::Image(image) => {
                    if let Some(data) = images.get(&image.key) {
                        image.original_width = data.width;
                        image.original_height = data.height;
                        if image.width <= 24.0 && data.width > 0 {
                            image.width = data.width as f32;
                        }
                        if image.height <= 24.0 && data.height > 0 {
                            image.height = data.height as f32;
                        }
                    }
                }
                Block::List(list) => {
                    for item in &mut list.items {
                        walk_blocks(&mut item.content, images);
                    }
                }
                Block::Table(table) => {
                    for row in &mut table.rows {
                        for cell in &mut row.cells {
                            walk_blocks(&mut cell.blocks, images);
                        }
                    }
                }
                Block::BlockQuote(quote) => {
                    walk_blocks(&mut quote.blocks, images);
                }
                Block::Paragraph(_) | Block::Heading(_) | Block::CodeBlock(_) | Block::PageBreak | Block::HorizontalRule => {}
            }
        }
    }

    let images = doc.images.clone();
    walk_blocks(&mut doc.content, &images);
}

fn emu_to_points(emu: f32) -> f32 {
    // 1 point = 12700 EMU.
    emu / 12700.0
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

#[cfg(test)]
mod tests {
    use super::parse_docx;
    use crate::document::model::Block;
    use crate::document::DocumentFormat;
    use std::{
        fs::{self, File},
        io::Write,
        path::PathBuf,
        sync::atomic::{AtomicU64, Ordering},
        time::{SystemTime, UNIX_EPOCH},
    };
    use zip::{ZipWriter, write::SimpleFileOptions};

    const DOC_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Demo Title</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Normal"/>
        <w:numPr><w:numId w:val="1"/></w:numPr>
      </w:pPr>
      <w:r><w:t>List Item A</w:t></w:r>
    </w:p>
    <w:p>
      <w:hyperlink r:id="rIdLink1"><w:r><w:t>https://example.com</w:t></w:r></w:hyperlink>
    </w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>Cell 1</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline><wp:extent cx="914400" cy="914400"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rIdImg1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>"#;

    const STYLES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr><w:b/></w:rPr>
  </w:style>
</w:styles>"#;

    const RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/pixel.png"/>
  <Relationship Id="rIdLink1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com"/>
</Relationships>"#;

    const CONTENT_TYPES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="png" ContentType="image/png"/>
</Types>"#;

    const NUMBERING_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="10"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="10"/></w:num>
</w:numbering>"#;

    const CORE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:title>Parser Test</dc:title>
  <dc:creator>Unit Test</dc:creator>
  <dcterms:created>2026-02-24T00:00:00Z</dcterms:created>
</cp:coreProperties>"#;

    const PIXEL_PNG: &[u8] = &[
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
        0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00,
        0x00, 0x90, 0x77, 0x53, 0xDE, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78,
        0x9C, 0x63, 0xF8, 0xCF, 0x00, 0x00, 0x02, 0x05, 0x01, 0x02, 0xA7, 0x69, 0x9D, 0x48,
        0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];

    fn write_test_docx(path: &PathBuf) {
        let file = File::create(path).expect("create docx");
        let mut zip = ZipWriter::new(file);
        let opts = SimpleFileOptions::default();
        let files = [
            ("word/document.xml", DOC_XML.as_bytes()),
            ("word/styles.xml", STYLES_XML.as_bytes()),
            ("word/_rels/document.xml.rels", RELS_XML.as_bytes()),
            ("[Content_Types].xml", CONTENT_TYPES_XML.as_bytes()),
            ("word/numbering.xml", NUMBERING_XML.as_bytes()),
            ("docProps/core.xml", CORE_XML.as_bytes()),
        ];

        for (name, bytes) in files {
            zip.start_file(name, opts).expect("start file");
            zip.write_all(bytes).expect("write file");
        }
        zip.start_file("word/media/pixel.png", opts).expect("start png");
        zip.write_all(PIXEL_PNG).expect("write png");
        zip.finish().expect("finish zip");
    }

    fn unique_temp_docx_path() -> PathBuf {
        static NEXT_ID: AtomicU64 = AtomicU64::new(1);
        let millis = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_millis();
        let seq = NEXT_ID.fetch_add(1, Ordering::Relaxed);
        std::env::temp_dir().join(format!("doco_parser_test_{millis}_{seq}.docx"))
    }

    #[test]
    fn parses_core_docx_blocks_and_images() {
        let path = unique_temp_docx_path();
        write_test_docx(&path);

        let parsed = parse_docx(&path).expect("parse docx");
        let _ = fs::remove_file(&path);

        assert_eq!(parsed.metadata.format, DocumentFormat::Docx);
        assert_eq!(parsed.metadata.author, "Unit Test");
        assert!(!parsed.content.is_empty());

        let has_heading = parsed.content.iter().any(|b| matches!(b, Block::Heading(h) if h.level == 1));
        let has_list = parsed.content.iter().any(|b| matches!(b, Block::List(_)));
        let has_table = parsed.content.iter().any(|b| matches!(b, Block::Table(_)));
        let has_image = parsed.content.iter().any(|b| matches!(b, Block::Image(_)));

        assert!(has_heading);
        assert!(has_list);
        assert!(has_table);
        assert!(has_image);
        assert!(parsed.images.contains_key("rIdImg1"));
    }

    #[test]
    fn resolves_style_inheritance_for_runs() {
        let path = unique_temp_docx_path();
        write_test_docx(&path);

        let parsed = parse_docx(&path).expect("parse docx");
        let _ = fs::remove_file(&path);

        let heading = parsed
            .content
            .iter()
            .find_map(|b| match b {
                Block::Heading(h) => Some(h),
                _ => None,
            })
            .expect("heading");
        let run = heading.runs.first().expect("heading run");
        assert_eq!(run.style.font_family.as_deref(), Some("Calibri"));
        assert_eq!(run.style.font_size, Some(11.0));
        assert!(run.style.bold);
    }
}
