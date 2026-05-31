use wasm_bindgen::prelude::*;
use std::collections::HashMap;
use std::io::{Cursor, Read};
use base64::{engine::general_purpose::STANDARD as B64, Engine as _};


mod markdown;

mod types;
pub use types::*;
mod styles;
use styles::*;
mod chart;
use chart::*;


// Excel built-in indexed color palette (indices 0-63)
// Standard Excel 2003 color palette
const INDEXED_COLORS: &[&str] = &[
    "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF", // 0-7
    "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF", // 8-15
    "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080", // 16-23
    "#9999FF", "#993366", "#FFFFCC", "#CCFFFF", "#660066", "#FF8080", "#0066CC", "#CCCCFF", // 24-31
    "#000080", "#FF00FF", "#FFFF00", "#00FFFF", "#800080", "#800000", "#008080", "#0000FF", // 32-39
    "#00CCFF", "#CCFFFF", "#CCFFCC", "#FFFF99", "#99CCFF", "#FF99CC", "#CC99FF", "#FFCC99", // 40-47
    "#3366FF", "#33CCCC", "#99CC00", "#FFCC00", "#FF9900", "#FF6600", "#666699", "#969696", // 48-55
    "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333", // 56-63
];

#[wasm_bindgen]
pub fn parse_xlsx(data: &[u8], max_zip_entry_bytes: Option<u64>) -> Result<String, JsValue> {
    console_error_panic_hook::set_once();
    let _guard = ooxml_common::zip::scoped_max(max_zip_entry_bytes);
    parse_xlsx_inner(data)
        .map(|wb| serde_json::to_string(&wb).unwrap())
        .map_err(|e| JsValue::from_str(&e))
}

#[wasm_bindgen]
pub fn parse_sheet(data: &[u8], sheet_index: u32, name: &str, max_zip_entry_bytes: Option<u64>) -> Result<String, JsValue> {
    console_error_panic_hook::set_once();
    let _guard = ooxml_common::zip::scoped_max(max_zip_entry_bytes);
    let cursor = Cursor::new(data);
    let mut archive = zip::ZipArchive::new(cursor).map_err(|e| e.to_string())?;

    let workbook_xml = read_zip_entry(&mut archive, "xl/workbook.xml")?;
    let wb_doc = roxmltree::Document::parse(&workbook_xml).map_err(|e| e.to_string())?;
    let sheets = parse_workbook_sheets(&wb_doc);

    let sheet_meta = sheets
        .get(sheet_index as usize)
        .ok_or_else(|| format!("sheet index {} out of range", sheet_index))?;

    // Resolve rId → target path from workbook.xml.rels
    let rels_xml = read_zip_entry(&mut archive, "xl/_rels/workbook.xml.rels")?;
    let rels_doc = roxmltree::Document::parse(&rels_xml).map_err(|e| e.to_string())?;
    let sheet_path = resolve_sheet_path(&rels_doc, &sheet_meta.r_id)
        .ok_or_else(|| format!("rId {} not found in rels", sheet_meta.r_id))?;

    let theme_colors = parse_theme_colors(&mut archive);
    let shared_strings = read_shared_strings(&mut archive, &theme_colors);
    let sheet_xml = read_zip_entry(&mut archive, &format!("xl/{}", sheet_path))?;
    let (mut ws, hyperlink_rids) = parse_worksheet(&sheet_xml, &shared_strings, &theme_colors, name)
        .map_err(|e| e.to_string())?;

    // Attach any drawing-anchored images and charts for this sheet
    ws.images = load_sheet_images(&mut archive, &sheet_path);
    ws.charts = load_sheet_charts(&mut archive, &sheet_path, &theme_colors);
    ws.shape_groups = load_sheet_shape_groups(&mut archive, &sheet_path, &theme_colors);
    ws.hyperlinks = load_hyperlinks(&mut archive, &sheet_path, hyperlink_rids);
    ws.comments = load_sheet_comments(&mut archive, &sheet_path);
    ws.comment_refs = ws.comments.iter().map(|c| c.cell_ref.clone()).collect();
    ws.defined_names = parse_defined_names_for_sheet(&wb_doc, sheet_index);
    ws.tables = load_sheet_tables(&mut archive, &sheet_path, &theme_colors);
    ws.slicers = load_sheet_slicers(&mut archive, &sheet_path);
    ws.sparkline_groups = load_sheet_sparklines(&mut archive, &sheet_xml, &sheets, &rels_doc, &theme_colors);
    let (df_family, df_size) = parse_default_font(&mut archive);
    ws.default_font_family = df_family;
    ws.default_font_size = df_size;

    serde_json::to_string(&ws).map_err(|e| JsValue::from_str(&e.to_string()))
}

fn parse_xlsx_inner(data: &[u8]) -> Result<ParsedWorkbook, String> {
    let cursor = Cursor::new(data);
    let mut archive = zip::ZipArchive::new(cursor).map_err(|e| e.to_string())?;

    let workbook_xml = read_zip_entry(&mut archive, "xl/workbook.xml")?;
    let wb_doc = roxmltree::Document::parse(&workbook_xml).map_err(|e| e.to_string())?;
    let sheets = parse_workbook_sheets(&wb_doc);

    let theme_colors = parse_theme_colors(&mut archive);
    let shared_strings = read_shared_strings(&mut archive, &theme_colors);
    let styles = parse_styles(&mut archive, &theme_colors)?;

    Ok(ParsedWorkbook {
        workbook: Workbook { sheets },
        styles,
        shared_strings,
    })
}

pub(crate) fn read_zip_entry(archive: &mut zip::ZipArchive<Cursor<&[u8]>>, name: &str) -> Result<String, String> {
    let max = ooxml_common::zip::current_max();
    let mut file = archive
        .by_name(name)
        .map_err(|e| format!("entry '{}' not found: {}", name, e))?;
    if file.size() > max {
        return Err(format!("entry '{}' exceeds size limit", name));
    }
    let mut buf = String::new();
    file.by_ref().take(max).read_to_string(&mut buf).map_err(|e| e.to_string())?;
    Ok(buf)
}

fn parse_theme_colors(archive: &mut zip::ZipArchive<Cursor<&[u8]>>) -> Vec<String> {
    let Ok(xml) = read_zip_entry(archive, "xl/theme/theme1.xml") else {
        return Vec::new();
    };
    let Ok(doc) = roxmltree::Document::parse(&xml) else {
        return Vec::new();
    };
    let a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main";

    // Find clrScheme node and collect child color elements in order
    // OOXML order: dk1, lt1, dk2, lt2, accent1, accent2, accent3, accent4, accent5, accent6, hlink, folHlink
    let mut colors: Vec<String> = Vec::new();
    for node in doc.descendants() {
        if node.tag_name().name() == "clrScheme" && node.tag_name().namespace() == Some(a_ns) {
            for child in node.children() {
                if !child.is_element() { continue; }
                // Each child is a color slot; its first child element holds the actual color
                for color_node in child.children() {
                    if !color_node.is_element() { continue; }
                    let hex = match color_node.tag_name().name() {
                        "srgbClr" => {
                            color_node.attribute("val").map(|v| format!("#{}", v.to_uppercase()))
                        }
                        "sysClr" => {
                            color_node.attribute("lastClr").map(|v| format!("#{}", v.to_uppercase()))
                        }
                        _ => None,
                    };
                    if let Some(h) = hex {
                        colors.push(h);
                        break;
                    }
                }
            }
            break;
        }
    }
    colors
}

/// Convert hex color + tint to resulting hex color using HLS model.
/// tint > 0: lighten; tint < 0: darken.
fn apply_tint(hex: &str, tint: f64) -> String {
    let hex = hex.trim_start_matches('#');
    if hex.len() < 6 { return format!("#{}", hex); }
    let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0) as f64 / 255.0;
    let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0) as f64 / 255.0;
    let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0) as f64 / 255.0;

    // RGB → HLS
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = (max + min) / 2.0;
    let s = if max == min {
        0.0
    } else if l < 0.5 {
        (max - min) / (max + min)
    } else {
        (max - min) / (2.0 - max - min)
    };
    let h = if max == min {
        0.0
    } else if max == r {
        (g - b) / (max - min) / 6.0
    } else if max == g {
        ((b - r) / (max - min) + 2.0) / 6.0
    } else {
        ((r - g) / (max - min) + 4.0) / 6.0
    };
    let h = if h < 0.0 { h + 1.0 } else { h };

    // Apply tint to luminance
    let new_l = if tint > 0.0 {
        l * (1.0 - tint) + tint
    } else {
        l * (1.0 + tint)
    };

    // HLS → RGB
    let (nr, ng, nb) = hls_to_rgb(h, new_l, s);
    format!("#{:02X}{:02X}{:02X}", (nr * 255.0).round() as u8, (ng * 255.0).round() as u8, (nb * 255.0).round() as u8)
}

fn hls_to_rgb(h: f64, l: f64, s: f64) -> (f64, f64, f64) {
    if s == 0.0 {
        return (l, l, l);
    }
    let q = if l < 0.5 { l * (1.0 + s) } else { l + s - l * s };
    let p = 2.0 * l - q;
    let r = hue_to_rgb(p, q, h + 1.0 / 3.0);
    let g = hue_to_rgb(p, q, h);
    let b = hue_to_rgb(p, q, h - 1.0 / 3.0);
    (r, g, b)
}

fn hue_to_rgb(p: f64, q: f64, mut t: f64) -> f64 {
    if t < 0.0 { t += 1.0; }
    if t > 1.0 { t -= 1.0; }
    if t < 1.0 / 6.0 { return p + (q - p) * 6.0 * t; }
    if t < 1.0 / 2.0 { return q; }
    if t < 2.0 / 3.0 { return p + (q - p) * (2.0 / 3.0 - t) * 6.0; }
    p
}

pub(crate) fn parse_color(node: &roxmltree::Node, theme_colors: &[String]) -> Option<String> {
    // rgb attribute (ARGB: 8 chars, drop alpha; or 6-char RGB)
    if let Some(rgb) = node.attribute("rgb") {
        if rgb.len() == 8 {
            return Some(format!("#{}", &rgb[2..].to_uppercase()));
        }
        return Some(format!("#{}", rgb.to_uppercase()));
    }

    // theme attribute → resolve from theme color array + optional tint
    //
    // ECMA-376 §18.8.3 stores the theme clrScheme in the order
    //   dk1, lt1, dk2, lt2, accent1..accent6, hlink, folHlink
    // but cell style references (c:color/@theme, c:fgColor/@theme, etc.) use
    // the Excel-internal index where dk1↔lt1 and dk2↔lt2 are SWAPPED:
    //   0=lt1, 1=dk1, 2=lt2, 3=dk2, 4..11 unchanged.
    // This is a well-known interoperability quirk (see Open-XML-SDK issue #46
    // and ECMA-376 §22.1.2.7 where "index values of 0 and 1 are swapped").
    if let Some(theme_str) = node.attribute("theme") {
        if let Ok(idx) = theme_str.parse::<usize>() {
            let mapped = match idx {
                0 => 1,
                1 => 0,
                2 => 3,
                3 => 2,
                n => n,
            };
            if let Some(base) = theme_colors.get(mapped) {
                let tint = node.attribute("tint").and_then(|s| s.parse::<f64>().ok()).unwrap_or(0.0);
                if tint == 0.0 {
                    return Some(base.clone());
                }
                return Some(apply_tint(base, tint));
            }
        }
    }

    // indexed attribute → Excel built-in palette
    if let Some(indexed_str) = node.attribute("indexed") {
        if let Ok(idx) = indexed_str.parse::<usize>() {
            // indices 64 (foreground) and 65 (background) are special: use black/white
            let color = match idx {
                64 => "#000000",
                65 => "#FFFFFF",
                _ => INDEXED_COLORS.get(idx).copied().unwrap_or("#000000"),
            };
            return Some(color.to_string());
        }
    }

    None
}

fn parse_workbook_sheets(doc: &roxmltree::Document) -> Vec<SheetMeta> {
    let mut sheets = Vec::new();
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    for node in doc.descendants() {
        if node.tag_name().name() == "sheet" && node.tag_name().namespace() == Some(ns) {
            let name = node.attribute("name").unwrap_or("Sheet").to_string();
            let sheet_id = node
                .attribute("sheetId")
                .and_then(|v| v.parse().ok())
                .unwrap_or(1);
            let r_id = node
                .attribute((r_ns, "id"))
                .unwrap_or("")
                .to_string();
            sheets.push(SheetMeta { name, sheet_id, r_id });
        }
    }
    sheets
}

/// Collect `<definedName>` entries from `workbook.xml`. `sheet_index` selects
/// which names are in scope: workbook-global (no `localSheetId`) plus any
/// whose `localSheetId` matches the given sheet position.
fn parse_defined_names_for_sheet(doc: &roxmltree::Document, sheet_index: u32) -> Vec<DefinedName> {
    let mut names = Vec::new();
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    for node in doc.descendants() {
        if node.tag_name().name() != "definedName" || node.tag_name().namespace() != Some(ns) {
            continue;
        }
        let local: Option<u32> = node.attribute("localSheetId").and_then(|s| s.parse().ok());
        if let Some(l) = local { if l != sheet_index { continue; } }
        let name = match node.attribute("name") { Some(n) => n.to_string(), None => continue };
        let formula = node.text().unwrap_or("").to_string();
        names.push(DefinedName { name, formula });
    }
    names
}

fn resolve_sheet_path(doc: &roxmltree::Document, r_id: &str) -> Option<String> {
    let ns = "http://schemas.openxmlformats.org/package/2006/relationships";
    for node in doc.descendants() {
        if node.tag_name().name() == "Relationship" && node.tag_name().namespace() == Some(ns) {
            if node.attribute("Id") == Some(r_id) {
                let target = node.attribute("Target")?;
                // ECMA-376 / Open Packaging Conventions: Target may be a
                // package-absolute path (`/xl/worksheets/sheet1.xml`, used
                // by openpyxl and some online tools) or a path relative to
                // the .rels file's parent (`worksheets/sheet1.xml`, the
                // common Office-saved form). Callers prepend `xl/` to the
                // returned value, so strip a leading `/xl/` to convert
                // absolute paths into the relative form they expect.
                let t = target.strip_prefix('/').unwrap_or(target);
                let t = t.strip_prefix("xl/").unwrap_or(t);
                return Some(t.to_string());
            }
        }
    }
    None
}

fn read_shared_strings(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    theme_colors: &[String],
) -> Vec<SharedString> {
    let Ok(xml) = read_zip_entry(archive, "xl/sharedStrings.xml") else {
        return Vec::new();
    };
    let Ok(doc) = roxmltree::Document::parse(&xml) else {
        return Vec::new();
    };
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let mut strings = Vec::new();
    for si in doc.descendants() {
        if si.tag_name().name() == "si" && si.tag_name().namespace() == Some(ns) {
            strings.push(parse_si_node(&si, ns, theme_colors));
        }
    }
    strings
}

/// Parse a `<si>` (shared) or `<is>` (inline) node into a SharedString.
/// The node may contain direct `<t>` text (plain) and/or multiple `<r>`
/// runs with per-run `<rPr>` font properties.
fn parse_si_node(
    node: &roxmltree::Node,
    ns: &str,
    theme_colors: &[String],
) -> SharedString {
    let mut text = String::new();
    let mut runs: Vec<Run> = Vec::new();
    let mut has_runs = false;
    for child in node.children() {
        if !child.is_element() { continue; }
        match child.tag_name().name() {
            "t" if child.tag_name().namespace() == Some(ns) => {
                if let Some(s) = child.text() {
                    text.push_str(s);
                }
            }
            "r" if child.tag_name().namespace() == Some(ns) => {
                has_runs = true;
                let mut run_text = String::new();
                let mut run_font: Option<RunFont> = None;
                for rc in child.children() {
                    match rc.tag_name().name() {
                        "t" => {
                            if let Some(s) = rc.text() {
                                run_text.push_str(s);
                            }
                        }
                        "rPr" => {
                            let mut f = RunFont::default();
                            for rp in rc.children() {
                                match rp.tag_name().name() {
                                    "b" => f.bold = parse_st_on_off(&rp),
                                    "i" => f.italic = parse_st_on_off(&rp),
                                    "u" => {
                                        // ECMA-376 §18.4.13 ST_UnderlineValues:
                                        // single (default) | double | singleAccounting |
                                        // doubleAccounting | none.
                                        let v = rp.attribute("val").unwrap_or("single");
                                        if v != "none" {
                                            f.underline = true;
                                            if v != "single" {
                                                f.underline_style = Some(v.to_string());
                                            }
                                        }
                                    }
                                    "strike" => f.strike = parse_st_on_off(&rp),
                                    "vertAlign" => {
                                        // ECMA-376 §18.4.6 ST_VerticalAlignRun.
                                        if let Some(v) = rp.attribute("val") {
                                            if v != "baseline" {
                                                f.vert_align = Some(v.to_string());
                                            }
                                        }
                                    }
                                    "sz" => {
                                        f.size = rp.attribute("val").and_then(|s| s.parse().ok());
                                    }
                                    "color" => {
                                        f.color = parse_color(&rp, theme_colors);
                                    }
                                    "rFont" | "name" => {
                                        f.name = rp.attribute("val").map(|s| s.to_string());
                                    }
                                    _ => {}
                                }
                            }
                            run_font = Some(f);
                        }
                        _ => {}
                    }
                }
                text.push_str(&run_text);
                runs.push(Run { text: run_text, font: run_font });
            }
            _ => {}
        }
    }
    SharedString {
        text,
        runs: if has_runs { Some(runs) } else { None },
    }
}
fn parse_worksheet(
    xml: &str,
    shared_strings: &[SharedString],
    theme_colors: &[String],
    name: &str,
) -> Result<(Worksheet, Vec<(u32, u32, String)>), String> {
    let doc = roxmltree::Document::parse(xml).map_err(|e| e.to_string())?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    let mut rows = Vec::new();
    let mut col_widths: HashMap<u32, f64> = HashMap::new();
    let mut row_heights: HashMap<u32, f64> = HashMap::new();
    let mut merge_cells: Vec<MergeCell> = Vec::new();
    let mut freeze_rows: u32 = 0;
    let mut freeze_cols: u32 = 0;
    let mut default_col_width = 8.43;
    // Intrinsic default row height in *points* — ECMA-376 §18.3.1.81.
    // 15 pt = 20 CSS px at 96 DPI, Excel's baseline for the Calibri 11
    // Normal style. The renderer multiplies by 4/3 at display time, so
    // both this default and per-row `<row ht="…">` values share the
    // same units across the parser/renderer boundary.
    let mut default_row_height = 15.0;
    let mut conditional_formats: Vec<ConditionalFormat> = Vec::new();
    let mut show_zeros = true;
    let mut show_gridlines = true;
    let mut tab_color: Option<String> = None;
    let mut auto_filter: Option<CellRange> = None;
    let mut hyperlink_rids: Vec<(u32, u32, String)> = Vec::new();

    // Pre-scan worksheet-level extLst for x14:dataBar extension attributes.
    // Excel 2010+ stores the `gradient` flag on `<x14:dataBar>` inside
    // `<extLst>/<ext>/<x14:conditionalFormattings>/<x14:conditionalFormatting>
    // /<x14:cfRule id="{GUID}">`, linked to the SpreadsheetML cfRule via a
    // matching `<x14:id>{GUID}</x14:id>` inside the cfRule's own extLst
    // (§2.6.3). Build a GUID → gradient map so cfRule parsing can look up
    // the override.
    let mut x14_databar_gradient: HashMap<String, bool> = HashMap::new();
    for x14_rule in doc.descendants().filter(|n| n.tag_name().name() == "cfRule" && n.attribute("type") == Some("dataBar")) {
        let Some(id) = x14_rule.attribute("id") else { continue };
        for bar in x14_rule.children().filter(|n| n.tag_name().name() == "dataBar") {
            if let Some(g) = bar.attribute("gradient") {
                x14_databar_gradient.insert(id.to_string(), !(g == "0" || g == "false"));
            }
        }
    }

    // Pre-scan worksheet-level extLst for x14:conditionalFormatting with
    // iconSet rules. Excel 2010+ stores custom icon sets (custom="1") here
    // with per-threshold `<x14:cfIcon iconSet="X" iconId="N"/>` overrides,
    // and cfvo values inside `<xm:f>` children instead of `val` attributes.
    // The sqref for x14 CF rules lives in a `<xm:sqref>` sibling.
    let mut x14_icon_formats: Vec<ConditionalFormat> = Vec::new();
    for x14_cf in doc.descendants().filter(|n| n.tag_name().name() == "conditionalFormatting" && n.tag_name().namespace().map(|u| u.contains("/spreadsheetml/2009/9")).unwrap_or(false)) {
        let sqref: Vec<CellRange> = x14_cf.children()
            .find(|n| n.tag_name().name() == "sqref")
            .and_then(|n| n.text())
            .map(parse_sqref)
            .unwrap_or_default();
        if sqref.is_empty() { continue; }
        let mut rules: Vec<CfRule> = Vec::new();
        for x14_rule in x14_cf.children().filter(|n| n.tag_name().name() == "cfRule" && n.attribute("type") == Some("iconSet")) {
            let priority: i32 = x14_rule.attribute("priority").and_then(|s| s.parse().ok()).unwrap_or(0);
            let Some(icon_node) = x14_rule.children().find(|n| n.tag_name().name() == "iconSet") else { continue };
            let custom = icon_node.attribute("custom").map(|v| v == "1" || v == "true").unwrap_or(false);
            let icon_set_name = icon_node.attribute("iconSet")
                .unwrap_or(if custom { "" } else { "3TrafficLights1" })
                .to_string();
            let reverse = icon_node.attribute("reverse").map(|v| v == "1" || v == "true").unwrap_or(false);
            let mut cfvos: Vec<CfValue> = Vec::new();
            let mut custom_icons: Vec<CfIcon> = Vec::new();
            for ch in icon_node.children().filter(|n| n.is_element()) {
                match ch.tag_name().name() {
                    "cfvo" => {
                        let kind = ch.attribute("type").unwrap_or("percent").to_string();
                        // x14:cfvo stores the value in `<xm:f>` child; attribute val fallback.
                        let value = ch.children()
                            .find(|n| n.tag_name().name() == "f")
                            .and_then(|n| n.text())
                            .map(|s| s.to_string())
                            .or_else(|| ch.attribute("val").map(|s| s.to_string()));
                        cfvos.push(CfValue { kind, value });
                    }
                    "cfIcon" => {
                        let set = ch.attribute("iconSet").unwrap_or("NoIcons").to_string();
                        let id = ch.attribute("iconId").and_then(|s| s.parse().ok()).unwrap_or(0);
                        custom_icons.push(CfIcon { icon_set: set, icon_id: id });
                    }
                    _ => {}
                }
            }
            rules.push(CfRule::IconSet {
                icon_set: icon_set_name,
                cfvos,
                reverse,
                priority,
                custom_icons: if custom { Some(custom_icons) } else { None },
            });
        }
        if !rules.is_empty() {
            x14_icon_formats.push(ConditionalFormat { sqref, rules });
        }
    }

    for node in doc.descendants() {
        match node.tag_name().name() {
            "sheetFormatPr" if node.tag_name().namespace() == Some(ns) => {
                if let Some(v) = node.attribute("defaultColWidth").and_then(|s| s.parse().ok()) {
                    default_col_width = v;
                }
                // ECMA-376 §18.3.1.81 `defaultRowHeight` is the workbook
                // default row height in points. Always honor it when present
                // — `demo/sample-1` stores `defaultRowHeight="20.1"` (no
                // customHeight) and Excel uses that 20.1 pt as the default
                // for non-customized rows. `customHeight` is metadata about
                // how the height was set, not a gate on whether to honor it.
                if let Some(v) = node
                    .attribute("defaultRowHeight")
                    .and_then(|s| s.parse::<f64>().ok())
                {
                    default_row_height = v;
                }
            }
            "col" if node.tag_name().namespace() == Some(ns) => {
                let custom = node.attribute("customWidth").map(|v| v == "1").unwrap_or(false);
                let hidden = node.attribute("hidden").map(|v| v == "1").unwrap_or(false);
                // Only record widths for custom-widthed columns OR hidden columns
                if !custom && !hidden { continue; }
                let min: u32 = node.attribute("min").and_then(|s| s.parse().ok()).unwrap_or(1);
                let max: u32 = node.attribute("max").and_then(|s| s.parse().ok()).unwrap_or(1);
                // Cap range to avoid storing 16K entries for max=16384 ranges
                let max = max.min(min + 255);
                let width: f64 = if hidden {
                    0.0
                } else {
                    node.attribute("width").and_then(|s| s.parse().ok()).unwrap_or(default_col_width)
                };
                for c in min..=max {
                    col_widths.insert(c, width);
                }
            }
            "sheetView" if node.tag_name().namespace() == Some(ns) => {
                show_zeros = node.attribute("showZeros").map(|v| v != "0").unwrap_or(true);
                show_gridlines = node.attribute("showGridLines").map(|v| v != "0").unwrap_or(true);
            }
            "tabColor" if node.tag_name().namespace() == Some(ns) => {
                tab_color = parse_color(&node, theme_colors);
            }
            "autoFilter" if node.tag_name().namespace() == Some(ns) => {
                if let Some(r) = node.attribute("ref") {
                    let parts: Vec<&str> = r.split(':').collect();
                    auto_filter = if parts.len() == 2 {
                        let (left, top) = parse_cell_ref(parts[0]);
                        let (right, bottom) = parse_cell_ref(parts[1]);
                        Some(CellRange { top, left, bottom, right })
                    } else {
                        let (col, row) = parse_cell_ref(parts[0]);
                        Some(CellRange { top: row, left: col, bottom: row, right: col })
                    };
                }
            }
            "hyperlinks" if node.tag_name().namespace() == Some(ns) => {
                for hl in node.children() {
                    if !hl.is_element() || hl.tag_name().name() != "hyperlink" { continue; }
                    let Some(ref_str) = hl.attribute("ref") else { continue };
                    // Only first cell of ref range
                    let ref_single = ref_str.split(':').next().unwrap_or(ref_str);
                    let (col, row) = parse_cell_ref(ref_single);
                    if let Some(rid) = hl.attributes()
                        .find(|a| a.name() == "id" && a.namespace() == Some(r_ns))
                        .map(|a| a.value().to_string())
                    {
                        hyperlink_rids.push((col, row, rid));
                    }
                }
            }
            "pane" if node.tag_name().namespace() == Some(ns) => {
                let state = node.attribute("state").unwrap_or("");
                if state == "frozen" || state == "frozenSplit" {
                    freeze_rows = node.attribute("ySplit")
                        .and_then(|s| s.parse::<f64>().ok())
                        .map(|v| v as u32)
                        .unwrap_or(0);
                    freeze_cols = node.attribute("xSplit")
                        .and_then(|s| s.parse::<f64>().ok())
                        .map(|v| v as u32)
                        .unwrap_or(0);
                }
            }
            "mergeCell" if node.tag_name().namespace() == Some(ns) => {
                if let Some(r) = node.attribute("ref") {
                    let parts: Vec<&str> = r.split(':').collect();
                    if parts.len() == 2 {
                        let (left, top) = parse_cell_ref(parts[0]);
                        let (right, bottom) = parse_cell_ref(parts[1]);
                        merge_cells.push(MergeCell { top, left, bottom, right });
                    }
                }
            }
            "row" if node.tag_name().namespace() == Some(ns) => {
                let row_idx: u32 = node.attribute("r").and_then(|s| s.parse().ok()).unwrap_or(0);
                let hidden = node.attribute("hidden").map(|v| v == "1").unwrap_or(false);
                // ECMA-376 §18.3.1.73 `<row>@ht` is the row height in points.
                // Gating the value on `@customHeight="1"` (0.37.0) was too
                // strict — `demo/sample-1` sheets 2-5 store `ht="36.95"` on
                // row 2 without `customHeight`, and Excel renders that row at
                // ~49 px (36.95 pt × 4/3), not the workbook default. Always
                // honor `ht` when present; `customHeight` is metadata about
                // *how* the height was set (user-edited vs auto-fit) and
                // doesn't gate the value itself.
                let height: Option<f64> = if hidden {
                    Some(0.0)
                } else {
                    node.attribute("ht").and_then(|s| s.parse().ok())
                };
                if let Some(h) = height {
                    row_heights.insert(row_idx, h);
                }
                let cells = parse_row_cells(&node, shared_strings, theme_colors, ns);
                rows.push(Row { index: row_idx, height, cells });
            }
            "conditionalFormatting" if node.tag_name().namespace() == Some(ns) => {
                let sqref = node.attribute("sqref")
                    .map(|s| parse_sqref(s))
                    .unwrap_or_default();
                let mut rules: Vec<CfRule> = Vec::new();
                for cf in node.children() {
                    if cf.tag_name().name() != "cfRule" { continue; }
                    let kind = cf.attribute("type").unwrap_or("").to_string();
                    let priority: i32 = cf.attribute("priority").and_then(|s| s.parse().ok()).unwrap_or(0);
                    let dxf_id: Option<u32> = cf.attribute("dxfId").and_then(|s| s.parse().ok());
                    match kind.as_str() {
                        "cellIs" => {
                            let operator = cf.attribute("operator").unwrap_or("equal").to_string();
                            let formulas: Vec<String> = cf.children()
                                .filter(|n| n.tag_name().name() == "formula")
                                .filter_map(|n| n.text().map(|s| s.to_string()))
                                .collect();
                            rules.push(CfRule::CellIs { operator, formulas, dxf_id, priority });
                        }
                        "expression"
                        | "containsBlanks" | "notContainsBlanks"
                        | "containsText" | "notContainsText"
                        | "beginsWith" | "endsWith"
                        | "containsErrors" | "notContainsErrors" => {
                            // For `containsBlanks`/`notContainsBlanks`/`containsText` etc.,
                            // Excel serializes an equivalent boolean formula (e.g.
                            // `LEN(TRIM(C8))>0`) as the rule's `<formula>` child
                            // (ECMA-376 §18.3.1.10). Evaluate as an expression rule.
                            let formula = cf.children()
                                .find(|n| n.tag_name().name() == "formula")
                                .and_then(|n| n.text())
                                .unwrap_or("")
                                .to_string();
                            let stop_if_true = cf.attribute("stopIfTrue")
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false);
                            rules.push(CfRule::Expression { formula, dxf_id, priority, stop_if_true });
                        }
                        "colorScale" => {
                            let scale = cf.children().find(|n| n.tag_name().name() == "colorScale");
                            let mut stop_values: Vec<(String, Option<String>)> = Vec::new();
                            let mut stop_colors: Vec<String> = Vec::new();
                            if let Some(scale_node) = scale {
                                for child in scale_node.children() {
                                    match child.tag_name().name() {
                                        "cfvo" => {
                                            stop_values.push((
                                                child.attribute("type").unwrap_or("num").to_string(),
                                                child.attribute("val").map(|s| s.to_string()),
                                            ));
                                        }
                                        "color" => {
                                            stop_colors.push(parse_color(&child, theme_colors).unwrap_or_else(|| "#FFFFFF".to_string()));
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            let stops: Vec<CfStop> = stop_values.into_iter().enumerate().map(|(i, (kind, value))| CfStop {
                                kind,
                                value,
                                color: stop_colors.get(i).cloned().unwrap_or_else(|| "#FFFFFF".to_string()),
                            }).collect();
                            rules.push(CfRule::ColorScale { stops, priority });
                        }
                        "dataBar" => {
                            let bar = cf.children().find(|n| n.tag_name().name() == "dataBar");
                            let mut cfvos: Vec<(String, Option<String>)> = Vec::new();
                            let mut color = "#638EC6".to_string();
                            if let Some(bar_node) = bar {
                                for child in bar_node.children() {
                                    match child.tag_name().name() {
                                        "cfvo" => {
                                            cfvos.push((
                                                child.attribute("type").unwrap_or("min").to_string(),
                                                child.attribute("val").map(|s| s.to_string()),
                                            ));
                                        }
                                        "color" => {
                                            if let Some(c) = parse_color(&child, theme_colors) { color = c; }
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            // Excel 2010+ x14:dataBar extension may override the
                            // gradient flag (§2.6.3, default="1"). "0" → solid
                            // fill. The override lives in a separate
                            // worksheet-level extLst and is linked via the
                            // `<x14:id>{GUID}</x14:id>` contained in this
                            // cfRule's own extLst.
                            let mut gradient = true;
                            'gradient_lookup: for ext_list in cf.children().filter(|n| n.tag_name().name() == "extLst") {
                                for ext in ext_list.children().filter(|n| n.tag_name().name() == "ext") {
                                    for id_node in ext.descendants().filter(|n| n.tag_name().name() == "id") {
                                        if let Some(guid) = id_node.text() {
                                            if let Some(&g) = x14_databar_gradient.get(guid) {
                                                gradient = g;
                                                break 'gradient_lookup;
                                            }
                                        }
                                    }
                                    // Fallback: some files embed <x14:dataBar>
                                    // directly in the cfRule's extLst.
                                    for x14_bar in ext.descendants().filter(|n| n.tag_name().name() == "dataBar") {
                                        if let Some(g) = x14_bar.attribute("gradient") {
                                            gradient = !(g == "0" || g == "false");
                                            break 'gradient_lookup;
                                        }
                                    }
                                }
                            }
                            let min = cfvos.first().map(|(k, v)| CfValue { kind: k.clone(), value: v.clone() })
                                .unwrap_or(CfValue { kind: "min".into(), value: None });
                            let max = cfvos.get(1).map(|(k, v)| CfValue { kind: k.clone(), value: v.clone() })
                                .unwrap_or(CfValue { kind: "max".into(), value: None });
                            rules.push(CfRule::DataBar { color, min, max, priority, gradient });
                        }
                        "top10" => {
                            let top = !cf.attribute("bottom").map(|v| v == "1" || v == "true").unwrap_or(false);
                            let percent = cf.attribute("percent").map(|v| v == "1" || v == "true").unwrap_or(false);
                            let rank = cf.attribute("rank").and_then(|s| s.parse().ok()).unwrap_or(10);
                            rules.push(CfRule::Top10 { top, percent, rank, dxf_id, priority });
                        }
                        "aboveAverage" => {
                            let above_average = cf.attribute("aboveAverage").map(|v| v != "0").unwrap_or(true);
                            rules.push(CfRule::AboveAverage { above_average, dxf_id, priority });
                        }
                        "iconSet" => {
                            let icon_set_node = cf.children().find(|n| n.tag_name().name() == "iconSet");
                            let icon_set = icon_set_node
                                .and_then(|n| n.attribute("iconSet"))
                                .unwrap_or("3TrafficLights1")
                                .to_string();
                            let reverse = icon_set_node
                                .and_then(|n| n.attribute("reverse"))
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false);
                            let cfvos: Vec<CfValue> = icon_set_node
                                .map(|n| n.children()
                                    .filter(|c| c.is_element() && c.tag_name().name() == "cfvo")
                                    .map(|c| CfValue {
                                        kind: c.attribute("type").unwrap_or("percent").to_string(),
                                        value: c.attribute("val").map(|s| s.to_string()),
                                    })
                                    .collect()
                                )
                                .unwrap_or_default();
                            rules.push(CfRule::IconSet { icon_set, cfvos, reverse, priority, custom_icons: None });
                        }
                        other => {
                            rules.push(CfRule::Other { kind: other.to_string(), priority });
                        }
                    }
                }
                conditional_formats.push(ConditionalFormat { sqref, rules });
            }
            _ => {}
        }
    }

    conditional_formats.extend(x14_icon_formats);

    let data_validations = parse_data_validations(doc.root_element());

    Ok((Worksheet {
        name: name.to_string(),
        rows,
        col_widths,
        row_heights,
        default_col_width,
        default_row_height,
        merge_cells,
        freeze_rows,
        freeze_cols,
        conditional_formats,
        images: Vec::new(),
        charts: Vec::new(),
        shape_groups: Vec::new(),
        show_zeros,
        show_gridlines,
        tab_color,
        auto_filter,
        hyperlinks: Vec::new(),
        comment_refs: Vec::new(),
        comments: Vec::new(),
        data_validations,
        defined_names: Vec::new(),
        tables: Vec::new(),
        slicers: Vec::new(),
        sparkline_groups: Vec::new(),
        default_font_family: None,
        default_font_size: None,
    }, hyperlink_rids))
}

/// Parse a .rels file into rId → Target map.
pub(crate) fn parse_rels_map(xml: &str) -> HashMap<String, String> {
    let Ok(doc) = roxmltree::Document::parse(xml) else {
        return HashMap::new();
    };
    let mut map = HashMap::new();
    for rel in doc.root_element().children().filter(|n| n.is_element()) {
        if let (Some(id), Some(target)) = (rel.attribute("Id"), rel.attribute("Target")) {
            map.insert(id.to_string(), target.to_string());
        }
    }
    map
}

/// Parse xl/comments{N}.xml referenced from the sheet's rels and collect the
/// list of A1-style cell refs that have a `<comment>` associated. The
/// renderer draws a small red triangle in each cell's top-right corner to
/// indicate the presence of a comment (ECMA-376 §18.7.3 commentList).
/// Reads xl/commentsN.xml for the given sheet and returns each `<comment>` as
/// a structured `XlsxComment` (cell ref, resolved author name, plain text).
/// Callers can derive `comment_refs: Vec<String>` from `c.cell_ref`.
fn load_sheet_comments(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    sheet_path: &str,
) -> Vec<XlsxComment> {
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else { return Vec::new(); };
    let sheet_rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let Ok(rels_xml) = read_zip_entry(archive, &sheet_rels_path) else { return Vec::new(); };
    let Ok(rels_doc) = roxmltree::Document::parse(&rels_xml) else { return Vec::new(); };

    // Accept both plain ("/comments") and threaded ("/threadedComment") relTypes
    // but prefer the classic comments file — threaded comments live in a
    // separate namespace and are an extension.
    let mut comments_target: Option<String> = None;
    for rel in rels_doc.root_element().children().filter(|n| n.is_element()) {
        let rel_type = rel.attribute("Type").unwrap_or("");
        if rel_type.ends_with("/comments") {
            if let Some(t) = rel.attribute("Target") {
                comments_target = Some(t.to_string());
                break;
            }
        }
    }
    let Some(target) = comments_target else { return Vec::new(); };

    let comments_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
    let Ok(comments_xml) = read_zip_entry(archive, &comments_path) else { return Vec::new(); };
    let Ok(comments_doc) = roxmltree::Document::parse(&comments_xml) else { return Vec::new(); };

    // Resolve <authors><author>…</author></authors> — `authorId` indexes here.
    let authors: Vec<String> = comments_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "authors")
        .map(|n| {
            n.children()
                .filter(|c| c.is_element() && c.tag_name().name() == "author")
                .map(|c| c.text().unwrap_or("").to_string())
                .collect()
        })
        .unwrap_or_default();

    let mut comments: Vec<XlsxComment> = Vec::new();
    for node in comments_doc.descendants() {
        if node.tag_name().name() != "comment" || !node.is_element() { continue }
        let Some(cell_ref) = node.attribute("ref") else { continue };
        let author = node
            .attribute("authorId")
            .and_then(|s| s.parse::<usize>().ok())
            .and_then(|i| authors.get(i).cloned())
            .filter(|s| !s.is_empty());
        let mut text = String::new();
        if let Some(t_node) = node.children().find(|c| c.is_element() && c.tag_name().name() == "text") {
            for r in t_node.descendants() {
                if r.is_element() && r.tag_name().name() == "t" {
                    if let Some(s) = r.text() { text.push_str(s); }
                }
            }
        }
        comments.push(XlsxComment {
            cell_ref: cell_ref.to_string(),
            author,
            text,
        });
    }
    comments
}

/// ECMA-376 §18.3.1.32 — extracts `<dataValidations>` rules from the sheet
/// XML root. Returns an empty vec when the element is absent.
fn parse_data_validations(ws_root: roxmltree::Node<'_, '_>) -> Vec<DataValidation> {
    let mut out: Vec<DataValidation> = Vec::new();
    let Some(dvs) = ws_root.children().find(|n| n.is_element() && n.tag_name().name() == "dataValidations") else {
        return out;
    };
    for dv in dvs.children().filter(|n| n.is_element() && n.tag_name().name() == "dataValidation") {
        let sqref = dv.attribute("sqref").unwrap_or("").to_string();
        if sqref.is_empty() { continue }
        let validation_type = dv.attribute("type").map(String::from);
        let operator = dv.attribute("operator").map(String::from);
        let allow_blank = dv.attribute("allowBlank").map(|v| v == "1" || v.eq_ignore_ascii_case("true")).unwrap_or(false);
        let prompt_title = dv.attribute("promptTitle").map(String::from).filter(|s| !s.is_empty());
        let prompt = dv.attribute("prompt").map(String::from).filter(|s| !s.is_empty());
        let error_title = dv.attribute("errorTitle").map(String::from).filter(|s| !s.is_empty());
        let error_message = dv.attribute("error").map(String::from).filter(|s| !s.is_empty());

        let mut formula1: Option<String> = None;
        let mut formula2: Option<String> = None;
        for child in dv.children().filter(|n| n.is_element()) {
            match child.tag_name().name() {
                "formula1" => formula1 = child.text().map(String::from).filter(|s| !s.is_empty()),
                "formula2" => formula2 = child.text().map(String::from).filter(|s| !s.is_empty()),
                _ => {}
            }
        }

        out.push(DataValidation {
            sqref,
            validation_type,
            operator,
            formula1,
            formula2,
            allow_blank,
            prompt_title,
            prompt,
            error_title,
            error_message,
        });
    }
    out
}

/// Parse `xl/tables/tableN.xml` files referenced from the sheet rels and
/// collect them for the renderer. Each table carries a ref range, style name
/// (e.g. "TableStyleLight18"), and the banded-rows / banded-cols flags from
/// `<tableStyleInfo>` (ECMA-376 §18.5).
/// Resolve a built-in table style's accent color from the theme.
///
/// Built-in style names follow the pattern `TableStyle{Light|Medium|Dark}{N}`
/// (ECMA-376 §18.5.1.4). Excel's UI lays the 21/28/11 built-ins out in a grid
/// of rows × 7 columns: column 0 is a "none" style (no accent), columns 1–6
/// map to accent1–accent6. So the accent index is `(N - 1) mod 7` where 0
/// means "no accent" and 1..=6 map to the theme's accent slots.
///
/// `theme_colors` is in OOXML natural order — accent1 lives at index 4, so
/// accent_n is at `theme_colors[3 + n]`. Falls back to a neutral gray when
/// the style name is unrecognised or the theme is missing accents.
/// dxf indices for the ECMA-376 §18.8.40 `<tableStyleElement>` roles we care
/// about. Built-in styles (`TableStyleLight18`, etc.) have no entry in the
/// file's `<tableStyles>` block and fall through to accent-based rendering;
/// custom styles (`"Gift Budget"`) reference dxfs from `<dxfs>`.
#[derive(Debug, Clone, Default)]
struct TableStyleElements {
    whole_table: Option<u32>,
    header_row: Option<u32>,
}

/// Parse `<tableStyles><tableStyle name="…"><tableStyleElement type="…" dxfId="…"/>`
/// into a lookup keyed by table-style name.
fn parse_table_styles_map(archive: &mut zip::ZipArchive<Cursor<&[u8]>>) -> std::collections::HashMap<String, TableStyleElements> {
    use std::collections::HashMap;
    let mut map: HashMap<String, TableStyleElements> = HashMap::new();
    let Ok(xml) = read_zip_entry(archive, "xl/styles.xml") else { return map; };
    let Ok(doc) = roxmltree::Document::parse(&xml) else { return map; };
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    for n in doc.descendants() {
        if n.tag_name().name() != "tableStyles" || n.tag_name().namespace() != Some(ns) { continue; }
        for ts in n.children().filter(|c| c.is_element() && c.tag_name().name() == "tableStyle") {
            let Some(name) = ts.attribute("name") else { continue; };
            let mut elems = TableStyleElements::default();
            for el in ts.children().filter(|c| c.is_element() && c.tag_name().name() == "tableStyleElement") {
                let t = el.attribute("type").unwrap_or("");
                let dxf: Option<u32> = el.attribute("dxfId").and_then(|s| s.parse().ok());
                match t {
                    "wholeTable" => elems.whole_table = dxf,
                    "headerRow"  => elems.header_row = dxf,
                    _ => {}
                }
            }
            map.insert(name.to_string(), elems);
        }
    }
    map
}

fn resolve_table_style_accent(style_name: &str, theme_colors: &[String]) -> String {
    let fallback = "#808080".to_string();
    let Some(rest) = style_name.strip_prefix("TableStyle") else { return fallback; };
    let digits_start = rest.find(|c: char| c.is_ascii_digit());
    let Some(start) = digits_start else { return fallback; };
    let Ok(n) = rest[start..].parse::<u32>() else { return fallback; };
    if n == 0 { return fallback; }
    let slot = ((n - 1) % 7) as usize;
    if slot == 0 { return fallback; }
    theme_colors.get(3 + slot).cloned().unwrap_or(fallback)
}

fn load_sheet_tables(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    sheet_path: &str,
    theme_colors: &[String],
) -> Vec<TableInfo> {
    let custom_styles = parse_table_styles_map(archive);
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else { return Vec::new(); };
    let sheet_rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let Ok(rels_xml) = read_zip_entry(archive, &sheet_rels_path) else { return Vec::new(); };
    let Ok(rels_doc) = roxmltree::Document::parse(&rels_xml) else { return Vec::new(); };

    let mut table_targets: Vec<String> = Vec::new();
    for rel in rels_doc.root_element().children().filter(|n| n.is_element()) {
        if rel.attribute("Type").unwrap_or("").ends_with("/table") {
            if let Some(t) = rel.attribute("Target") {
                table_targets.push(t.to_string());
            }
        }
    }

    let mut tables: Vec<TableInfo> = Vec::new();
    for target in table_targets {
        let table_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        let Ok(xml) = read_zip_entry(archive, &table_path) else { continue; };
        let Ok(doc) = roxmltree::Document::parse(&xml) else { continue; };
        let root = doc.root_element();
        let Some(ref_attr) = root.attribute("ref") else { continue };
        let parts: Vec<&str> = ref_attr.split(':').collect();
        let range = if parts.len() == 2 {
            let (left, top) = parse_cell_ref(parts[0]);
            let (right, bottom) = parse_cell_ref(parts[1]);
            CellRange { top, left, bottom, right }
        } else {
            let (col, row) = parse_cell_ref(parts[0]);
            CellRange { top: row, left: col, bottom: row, right: col }
        };
        let header_row_count: u32 = root.attribute("headerRowCount")
            .and_then(|s| s.parse().ok())
            .unwrap_or(1);
        let totals_row_count: u32 = root.attribute("totalsRowCount")
            .and_then(|s| s.parse().ok())
            .unwrap_or(0);
        let style_info = root.children().find(|n| n.tag_name().name() == "tableStyleInfo");
        // ECMA-376 §18.5.1.4: when `name` is absent the table has "None" style —
        // no visual table formatting. Default to "" rather than a named style so
        // the renderer can skip table-style overlay for these cells.
        let style_name = style_info
            .and_then(|n| n.attribute("name"))
            .unwrap_or("")
            .to_string();
        let bool_attr = |n: &roxmltree::Node, key: &str| n.attribute(key).map(|v| v == "1" || v == "true").unwrap_or(false);
        let (show_row_stripes, show_column_stripes, show_first_column, show_last_column) = match style_info {
            Some(n) => (
                bool_attr(&n, "showRowStripes"),
                bool_attr(&n, "showColumnStripes"),
                bool_attr(&n, "showFirstColumn"),
                bool_attr(&n, "showLastColumn"),
            ),
            None => (false, false, false, false),
        };
        let accent_color = resolve_table_style_accent(&style_name, theme_colors);
        let (whole_table_dxf, header_row_dxf) = match custom_styles.get(&style_name) {
            Some(e) => (e.whole_table, e.header_row),
            None => (None, None),
        };
        // ECMA-376 §18.5.1.3: each `<tableColumn>` may carry its own
        // `dataDxfId`, `headerRowDxfId`, `totalsRowDxfId`. We collect them in
        // document order so the renderer can index them via
        // `columns[cellCol - range.left]`.
        let columns: Vec<TableColumnInfo> = root
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "tableColumn")
            .map(|tc| TableColumnInfo {
                data_dxf_id:       tc.attribute("dataDxfId").and_then(|s| s.parse().ok()),
                header_row_dxf_id: tc.attribute("headerRowDxfId").and_then(|s| s.parse().ok()),
                totals_row_dxf_id: tc.attribute("totalsRowDxfId").and_then(|s| s.parse().ok()),
            })
            .collect();
        tables.push(TableInfo {
            range,
            style_name,
            header_row_count,
            totals_row_count,
            show_row_stripes,
            show_column_stripes,
            show_first_column,
            show_last_column,
            accent_color,
            whole_table_dxf,
            header_row_dxf,
            columns,
        });
    }
    tables
}

/// Resolve hyperlink rIds to URLs from the sheet rels file.
fn load_hyperlinks(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    sheet_path: &str,
    hyperlink_rids: Vec<(u32, u32, String)>,
) -> Vec<Hyperlink> {
    if hyperlink_rids.is_empty() { return Vec::new(); }
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else { return Vec::new(); };
    let rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let rels = read_zip_entry(archive, &rels_path)
        .ok()
        .map(|xml| parse_rels_map(&xml))
        .unwrap_or_default();
    hyperlink_rids.into_iter().map(|(col, row, rid)| Hyperlink {
        col, row, url: rels.get(&rid).cloned(),
    }).collect()
}

/// Read a binary file from the zip.
fn read_zip_bytes(archive: &mut zip::ZipArchive<Cursor<&[u8]>>, path: &str) -> Option<Vec<u8>> {
    let max = ooxml_common::zip::current_max();
    let mut file = archive.by_name(path).ok()?;
    if file.size() > max {
        return None;
    }
    let mut buf = Vec::new();
    file.by_ref().take(max).read_to_end(&mut buf).ok()?;
    Some(buf)
}

/// Resolve a relative path ("../media/image1.png") against a base dir ("xl/drawings").
pub(crate) fn resolve_zip_path(base_dir: &str, target: &str) -> String {
    let mut parts: Vec<&str> = base_dir.split('/').filter(|s| !s.is_empty()).collect();
    for seg in target.split('/') {
        match seg {
            ".." => { parts.pop(); }
            "." | "" => {}
            s => parts.push(s),
        }
    }
    parts.join("/")
}

fn mime_from_ext(path: &str) -> &'static str {
    match path.rsplit('.').next().unwrap_or("").to_ascii_lowercase().as_str() {
        "png"  => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif"  => "image/gif",
        "bmp"  => "image/bmp",
        "webp" => "image/webp",
        _      => "application/octet-stream",
    }
}

/// Parse `<xdr:twoCellAnchor>` elements from a drawing XML and resolve
/// embedded pictures into data URLs. `drawing_dir` is the folder that
/// contains `drawing_path` so relative `Target`s resolve correctly.
fn parse_drawing_anchors(
    drawing_xml: &str,
    drawing_rels: &HashMap<String, String>,
    drawing_dir: &str,
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
) -> Vec<ImageAnchor> {
    let Ok(doc) = roxmltree::Document::parse(drawing_xml) else {
        return Vec::new();
    };
    let xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    let a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main";
    let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let mut anchors: Vec<ImageAnchor> = Vec::new();

    for anchor in doc.descendants() {
        if anchor.tag_name().name() != "twoCellAnchor"
            || anchor.tag_name().namespace() != Some(xdr_ns)
        {
            continue;
        }
        let (mut from_col, mut from_col_off, mut from_row, mut from_row_off) = (0u32, 0i64, 0u32, 0i64);
        let (mut to_col,   mut to_col_off,   mut to_row,   mut to_row_off)   = (0u32, 0i64, 0u32, 0i64);
        let mut pic_rid: Option<String> = None;
        let mut native_ext_cx: i64 = 0;
        let mut native_ext_cy: i64 = 0;
        // ECMA-376 §20.5.2.33 `twoCellAnchor@editAs`. Possible values:
        // "twoCell" (default), "oneCell", "absolute". With "oneCell" Excel
        // preserves the picture's saved size from <xdr:spPr><a:xfrm><a:ext>
        // regardless of cell resizing.
        let edit_as = anchor.attribute("editAs").map(|s| s.to_string());

        for child in anchor.children() {
            if !child.is_element() { continue; }
            match child.tag_name().name() {
                "from" | "to" => {
                    let is_from = child.tag_name().name() == "from";
                    let mut col: u32 = 0;
                    let mut col_off: i64 = 0;
                    let mut row: u32 = 0;
                    let mut row_off: i64 = 0;
                    for c in child.children() {
                        match (c.tag_name().name(), c.text()) {
                            ("col",    Some(t)) => col     = t.trim().parse().unwrap_or(0),
                            ("colOff", Some(t)) => col_off = t.trim().parse().unwrap_or(0),
                            ("row",    Some(t)) => row     = t.trim().parse().unwrap_or(0),
                            ("rowOff", Some(t)) => row_off = t.trim().parse().unwrap_or(0),
                            _ => {}
                        }
                    }
                    if is_from {
                        from_col = col; from_col_off = col_off; from_row = row; from_row_off = row_off;
                    } else {
                        to_col = col; to_col_off = col_off; to_row = row; to_row_off = row_off;
                    }
                }
                "pic" => {
                    // <xdr:pic><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic>
                    let blip_fill = child.children()
                        .find(|n| n.tag_name().name() == "blipFill" && n.tag_name().namespace() == Some(xdr_ns));
                    if let Some(bf) = blip_fill {
                        let blip = bf.children()
                            .find(|n| n.tag_name().name() == "blip" && n.tag_name().namespace() == Some(a_ns));
                        if let Some(b) = blip {
                            // r:embed attribute
                            pic_rid = b.attributes()
                                .find(|a| a.name() == "embed" && a.namespace() == Some(r_ns))
                                .map(|a| a.value().to_string());
                        }
                    }
                    // <xdr:pic><xdr:spPr><a:xfrm><a:ext cx cy>: the picture's
                    // own saved EMU extent. Authoritative when editAs="oneCell".
                    if let Some(sp_pr) = child.children()
                        .find(|n| n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(xdr_ns))
                    {
                        if let Some(xfrm_n) = sp_pr.children()
                            .find(|n| n.tag_name().name() == "xfrm" && n.tag_name().namespace() == Some(a_ns))
                        {
                            if let Some(xfrm) = parse_xfrm(&xfrm_n) {
                                native_ext_cx = xfrm.ext_x as i64;
                                native_ext_cy = xfrm.ext_y as i64;
                            }
                        }
                    }
                }
                _ => {}
            }
        }

        let Some(rid) = pic_rid else { continue; };
        let Some(target) = drawing_rels.get(&rid) else { continue; };
        let media_path = resolve_zip_path(drawing_dir, target);
        let Some(bytes) = read_zip_bytes(archive, &media_path) else { continue; };
        let mime = mime_from_ext(&media_path);
        let data_url = format!("data:{mime};base64,{}", B64.encode(&bytes));

        anchors.push(ImageAnchor {
            from_col, from_col_off, from_row, from_row_off,
            to_col, to_col_off, to_row, to_row_off,
            edit_as,
            native_ext_cx,
            native_ext_cy,
            data_url,
        });
    }
    anchors
}

// ─── Shape group parsing ────────────────────────────────────────────────────
//
// ECMA-376 §20.5.2.17 `<xdr:grpSp>` / §20.1.9 DrawingML shapes. Each
// top-level grpSp inside a twoCellAnchor has its own coordinate system:
//   - grpSpPr/xfrm/off,ext     : group's position/size in parent coords
//   - grpSpPr/xfrm/chOff,chExt : origin/extent of the group's child coords
//
// A child sp at child coord (cx, cy) maps to parent coord:
//   parent.x = off.x + (cx - chOff.x) / chExt.cx * ext.cx
//
// For rendering, we chain these transforms down to the top-level grpSp and
// then normalize each leaf shape's rect into [0,1] of the top-level ext.

#[derive(Clone, Copy)]
struct Xfrm {
    off_x: f64, off_y: f64,
    ext_x: f64, ext_y: f64,
    ch_off_x: f64, ch_off_y: f64,
    ch_ext_x: f64, ch_ext_y: f64,
    has_ch: bool,
}

fn parse_xfrm(xfrm_node: &roxmltree::Node) -> Option<Xfrm> {
    let mut off = (0.0_f64, 0.0_f64);
    let mut ext = (0.0_f64, 0.0_f64);
    let mut ch_off = (0.0_f64, 0.0_f64);
    let mut ch_ext = (0.0_f64, 0.0_f64);
    let mut has_ext = false;
    let mut has_ch = false;
    for c in xfrm_node.children() {
        match c.tag_name().name() {
            "off" => {
                off.0 = c.attribute("x").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                off.1 = c.attribute("y").and_then(|s| s.parse().ok()).unwrap_or(0.0);
            }
            "ext" => {
                ext.0 = c.attribute("cx").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                ext.1 = c.attribute("cy").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                has_ext = true;
            }
            "chOff" => {
                ch_off.0 = c.attribute("x").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                ch_off.1 = c.attribute("y").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                has_ch = true;
            }
            "chExt" => {
                ch_ext.0 = c.attribute("cx").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                ch_ext.1 = c.attribute("cy").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                has_ch = true;
            }
            _ => {}
        }
    }
    if !has_ext { return None; }
    Some(Xfrm {
        off_x: off.0, off_y: off.1,
        ext_x: ext.0, ext_y: ext.1,
        ch_off_x: ch_off.0, ch_off_y: ch_off.1,
        ch_ext_x: if ch_ext.0 == 0.0 { ext.0 } else { ch_ext.0 },
        ch_ext_y: if ch_ext.1 == 0.0 { ext.1 } else { ch_ext.1 },
        has_ch,
    })
}

fn parse_solid_fill(fill_node: &roxmltree::Node, theme_colors: &[String]) -> Option<String> {
    for c in fill_node.children() {
        match c.tag_name().name() {
            "srgbClr" => {
                let v = c.attribute("val")?;
                return Some(format!("#{}", v.to_uppercase()));
            }
            "schemeClr" => {
                let v = c.attribute("val")?;
                // `theme_colors` is collected in OOXML clrScheme document
                // order: dk1, lt1, dk2, lt2, accent1..accent6, hlink,
                // folHlink. See `parse_theme_colors`. The earlier mapping
                // here had dk1/lt1 and dk2/lt2 swapped which darkened
                // shapes that painted "lt1" (the sheet paper colour).
                let idx = match v {
                    "dk1" | "tx1"    => Some(0),
                    "lt1" | "bg1"    => Some(1),
                    "dk2" | "tx2"    => Some(2),
                    "lt2" | "bg2"    => Some(3),
                    "accent1"        => Some(4),
                    "accent2"        => Some(5),
                    "accent3"        => Some(6),
                    "accent4"        => Some(7),
                    "accent5"        => Some(8),
                    "accent6"        => Some(9),
                    "hlink"          => Some(10),
                    "folHlink"       => Some(11),
                    _ => None,
                };
                return idx.and_then(|i| theme_colors.get(i).cloned());
            }
            _ => {}
        }
    }
    None
}

/// Parse a single custGeom path element. Each path has its own coordinate
/// system (`a:path/@w`, `@h`) that the renderer scales to the shape's rect.
fn parse_custom_path(path_node: &roxmltree::Node) -> PathInfo {
    let w: f64 = path_node.attribute("w").and_then(|s| s.parse().ok()).unwrap_or(0.0);
    let h: f64 = path_node.attribute("h").and_then(|s| s.parse().ok()).unwrap_or(0.0);
    let mut commands: Vec<PathCmd> = Vec::new();
    for cmd in path_node.children().filter(|n| n.is_element()) {
        let name = cmd.tag_name().name();
        // Collect `<a:pt x=.. y=..>` points in order.
        let pts: Vec<(f64, f64)> = cmd.children()
            .filter(|n| n.is_element() && n.tag_name().name() == "pt")
            .map(|n| (
                n.attribute("x").and_then(|s| s.parse().ok()).unwrap_or(0.0),
                n.attribute("y").and_then(|s| s.parse().ok()).unwrap_or(0.0),
            ))
            .collect();
        match name {
            "moveTo"       => if let Some(p) = pts.first() { commands.push(PathCmd::MoveTo { x: p.0, y: p.1 }); },
            "lnTo"         => if let Some(p) = pts.first() { commands.push(PathCmd::LineTo { x: p.0, y: p.1 }); },
            "cubicBezTo"   => if pts.len() >= 3 {
                commands.push(PathCmd::CubicBezTo {
                    x1: pts[0].0, y1: pts[0].1,
                    x2: pts[1].0, y2: pts[1].1,
                    x3: pts[2].0, y3: pts[2].1,
                });
            },
            "quadBezTo"    => if pts.len() >= 2 {
                commands.push(PathCmd::QuadBezTo {
                    x1: pts[0].0, y1: pts[0].1,
                    x2: pts[1].0, y2: pts[1].1,
                });
            },
            "close"        => commands.push(PathCmd::Close),
            "arcTo" => {
                // ECMA-376 §20.1.9.3: `wR`/`hR` in path-coord units;
                // `stAng`/`swAng` in 60000ths of a degree.
                let wr:     f64 = cmd.attribute("wR").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                let hr:     f64 = cmd.attribute("hR").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                let st_ang: f64 = cmd.attribute("stAng").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                let sw_ang: f64 = cmd.attribute("swAng").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                commands.push(PathCmd::ArcTo { wr, hr, st_ang, sw_ang });
            }
            _ => {}
        }
    }
    PathInfo { w, h, commands }
}

/// Parse `<xdr:txBody>` into a `ShapeText`. Returns `None` if the body
/// contains no visible runs. Run formatting follows ECMA-376 §21.1.2.3.1
/// (`<a:rPr>`): `sz` is hundredths of a point, `b="1"` = bold, `i="1"`
/// = italic, `<a:solidFill>` overrides shape-level font color, and
/// `<a:latin@typeface>` selects the Latin font face (we don't yet
/// distinguish East-Asian / complex-script fonts — `<a:ea>` and `<a:cs>`
/// are ignored for typeface).
fn parse_tx_body(tx_body: &roxmltree::Node, theme_colors: &[String]) -> Option<ShapeText> {
    let mut anchor = String::from("t");
    let mut wrap = String::from("square");
    let mut paragraphs: Vec<ShapeParagraph> = Vec::new();
    for c in tx_body.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "bodyPr" => {
                if let Some(a) = c.attribute("anchor") { anchor = a.to_string(); }
                if let Some(w) = c.attribute("wrap") { wrap = w.to_string(); }
            }
            "p" => {
                let mut align = String::from("l");
                let mut runs: Vec<ShapeTextRun> = Vec::new();
                for pc in c.children().filter(|n| n.is_element()) {
                    match pc.tag_name().name() {
                        "pPr" => {
                            if let Some(a) = pc.attribute("algn") { align = a.to_string(); }
                        }
                        "r" => {
                            // Run text + run-level formatting.
                            let mut text = String::new();
                            let mut bold = false;
                            let mut italic = false;
                            let mut size: f64 = 0.0;
                            let mut color: Option<String> = None;
                            let mut font_face: Option<String> = None;
                            for rc in pc.children().filter(|n| n.is_element()) {
                                match rc.tag_name().name() {
                                    "rPr" => {
                                        bold = rc.attribute("b").map(|v| v == "1").unwrap_or(false);
                                        italic = rc.attribute("i").map(|v| v == "1").unwrap_or(false);
                                        size = rc.attribute("sz")
                                            .and_then(|s| s.parse::<f64>().ok())
                                            .map(|v| v / 100.0)
                                            .unwrap_or(0.0);
                                        for rpc in rc.children().filter(|n| n.is_element()) {
                                            match rpc.tag_name().name() {
                                                "solidFill" => {
                                                    color = parse_solid_fill(&rpc, theme_colors);
                                                }
                                                "latin" => {
                                                    font_face = rpc.attribute("typeface").map(String::from);
                                                }
                                                _ => {}
                                            }
                                        }
                                    }
                                    "t" => {
                                        if let Some(t) = rc.text() { text.push_str(t); }
                                    }
                                    _ => {}
                                }
                            }
                            if !text.is_empty() {
                                runs.push(ShapeTextRun { text, bold, italic, size, color, font_face });
                            }
                        }
                        "br" => {
                            // Soft line break: emit an empty run with a newline marker.
                            // We collapse to a literal newline since the renderer paints
                            // each \n as a wrapped line.
                            runs.push(ShapeTextRun {
                                text: "\n".into(),
                                bold: false, italic: false, size: 0.0,
                                color: None, font_face: None,
                            });
                        }
                        _ => {}
                    }
                }
                if !runs.is_empty() {
                    paragraphs.push(ShapeParagraph { align, runs });
                }
            }
            _ => {}
        }
    }
    if paragraphs.is_empty() { None } else {
        Some(ShapeText { anchor, wrap, paragraphs })
    }
}

fn parse_sp_geom(sp_pr: &roxmltree::Node) -> Option<ShapeGeom> {
    for c in sp_pr.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "prstGeom" => {
                return Some(ShapeGeom::Preset {
                    name: c.attribute("prst").unwrap_or("rect").to_string(),
                });
            }
            "custGeom" => {
                let mut paths: Vec<PathInfo> = Vec::new();
                for pl in c.children().filter(|n| n.is_element() && n.tag_name().name() == "pathLst") {
                    for p in pl.children().filter(|n| n.is_element() && n.tag_name().name() == "path") {
                        paths.push(parse_custom_path(&p));
                    }
                }
                return Some(ShapeGeom::Custom { paths });
            }
            _ => {}
        }
    }
    None
}

/// Recursively walk an `xdr:grpSp` / `xdr:sp` tree, chaining coordinate
/// transforms, and push leaf shapes (normalized to [0,1] of `root_ext`) into
/// `out`.
fn collect_shapes(
    node: &roxmltree::Node,
    root_off_x: f64, root_off_y: f64,
    root_ext_x: f64, root_ext_y: f64,
    // transform from current local coords into root (top-level grpSp) coords
    scale_x: f64, scale_y: f64,
    trans_x: f64, trans_y: f64,
    theme_colors: &[String],
    rid_urls: &HashMap<String, String>,
    out: &mut Vec<ShapeInfo>,
) {
    for child in node.children().filter(|n| n.is_element()) {
        let tag = child.tag_name().name();
        if tag == "grpSp" {
            // Nested grpSp: compose the transform by the group's own xfrm.
            let grp_sp_pr = child.children().find(|n| n.is_element() && n.tag_name().name() == "grpSpPr");
            let xfrm = grp_sp_pr
                .and_then(|n| n.children().find(|c| c.is_element() && c.tag_name().name() == "xfrm"))
                .as_ref()
                .and_then(parse_xfrm);
            let (sx, sy, tx, ty) = if let Some(x) = xfrm {
                if x.has_ch && x.ch_ext_x != 0.0 && x.ch_ext_y != 0.0 {
                    let csx = x.ext_x / x.ch_ext_x;
                    let csy = x.ext_y / x.ch_ext_y;
                    // Child point (cx, cy) → (x.off_x + (cx - x.ch_off_x)*csx) in parent coords,
                    // then apply outer (scale/trans) to reach root coords.
                    (
                        scale_x * csx,
                        scale_y * csy,
                        trans_x + scale_x * (x.off_x - x.ch_off_x * csx),
                        trans_y + scale_y * (x.off_y - x.ch_off_y * csy),
                    )
                } else {
                    // No child coord system: treat as identity mapping inside the group.
                    (scale_x, scale_y,
                     trans_x + scale_x * x.off_x,
                     trans_y + scale_y * x.off_y)
                }
            } else {
                (scale_x, scale_y, trans_x, trans_y)
            };
            collect_shapes(&child, root_off_x, root_off_y, root_ext_x, root_ext_y,
                           sx, sy, tx, ty, theme_colors, rid_urls, out);
        } else if tag == "sp" {
            let sp_pr = child.children().find(|n| n.is_element() && n.tag_name().name() == "spPr");
            let Some(sp_pr) = sp_pr else { continue; };
            let xfrm_node = sp_pr.children().find(|n| n.is_element() && n.tag_name().name() == "xfrm");
            let Some(xfrm_n) = xfrm_node else { continue; };
            let Some(xfrm) = parse_xfrm(&xfrm_n) else { continue; };
            let rot_raw: f64 = xfrm_n.attribute("rot")
                .and_then(|s| s.parse().ok()).unwrap_or(0.0);

            // Shape rect in root coords
            let root_x = trans_x + scale_x * xfrm.off_x;
            let root_y = trans_y + scale_y * xfrm.off_y;
            let root_w = scale_x * xfrm.ext_x;
            let root_h = scale_y * xfrm.ext_y;

            // Normalize to [0,1] of root ext
            if root_ext_x == 0.0 || root_ext_y == 0.0 { continue; }
            let nx = (root_x - root_off_x) / root_ext_x;
            let ny = (root_y - root_off_y) / root_ext_y;
            let nw = root_w / root_ext_x;
            let nh = root_h / root_ext_y;

            let geom = parse_sp_geom(&sp_pr);
            let Some(geom) = geom else { continue; };

            // Fill
            let mut fill_color: Option<String> = None;
            let mut has_no_fill = false;
            for c in sp_pr.children().filter(|n| n.is_element()) {
                match c.tag_name().name() {
                    "solidFill" => { fill_color = parse_solid_fill(&c, theme_colors); }
                    "noFill"    => { has_no_fill = true; }
                    _ => {}
                }
            }
            if has_no_fill { fill_color = None; }

            // Stroke (line)
            let mut stroke_color: Option<String> = None;
            let mut stroke_width: i64 = 0;
            if let Some(ln) = sp_pr.children().find(|n| n.is_element() && n.tag_name().name() == "ln") {
                stroke_width = ln.attribute("w").and_then(|s| s.parse().ok()).unwrap_or(0);
                for c in ln.children().filter(|n| n.is_element()) {
                    if c.tag_name().name() == "solidFill" {
                        stroke_color = parse_solid_fill(&c, theme_colors);
                    } else if c.tag_name().name() == "noFill" {
                        stroke_color = None;
                        stroke_width = 0;
                    }
                }
            }

            // <xdr:style> drives fallbacks: <a:fillRef> supplies a fill when
            // <xdr:spPr> didn't, and <a:fontRef> supplies the run-default text
            // color (ECMA-376 §20.5.2.30 `<xdr:style>`). Real-world text boxes
            // saved by Excel often leave `<xdr:spPr>` without `<a:solidFill>`
            // and rely on the style's fillRef + fontRef pair (e.g. accent1
            // background + lt1 white text). We resolve scheme colors here
            // against the workbook theme and apply them as fallbacks.
            let style_node = child.children()
                .find(|n| n.is_element() && n.tag_name().name() == "style");
            let style_fill = style_node.as_ref()
                .and_then(|s| s.children().find(|n| n.is_element() && n.tag_name().name() == "fillRef"))
                .and_then(|n| parse_solid_fill(&n, theme_colors));
            let style_text_color = style_node.as_ref()
                .and_then(|s| s.children().find(|n| n.is_element() && n.tag_name().name() == "fontRef"))
                .and_then(|n| parse_solid_fill(&n, theme_colors));
            if fill_color.is_none() && !has_no_fill {
                fill_color = style_fill;
            }

            // Text body (txBox shapes carry visible text inside
            // `<xdr:txBody>`; non-textbox shapes may also have one).
            let mut text = child
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "txBody")
                .and_then(|tb| parse_tx_body(&tb, theme_colors));
            if let (Some(t), Some(default_color)) = (text.as_mut(), style_text_color) {
                for p in t.paragraphs.iter_mut() {
                    for r in p.runs.iter_mut() {
                        if r.color.is_none() {
                            r.color = Some(default_color.clone());
                        }
                    }
                }
            }

            out.push(ShapeInfo {
                x: nx, y: ny, w: nw, h: nh,
                rot: rot_raw / 60000.0,
                fill_color,
                stroke_color,
                stroke_width,
                geom,
                text,
            });
        } else if tag == "pic" {
            // `<xdr:pic>` leaf inside a group (ECMA-376 §20.5.2.17). The image
            // binary is resolved via the drawing's .rels file; `rid_urls` maps
            // each r:id to its pre-encoded `data:<mime>;base64,…` URL.
            let sp_pr = child.children().find(|n| n.is_element() && n.tag_name().name() == "spPr");
            let Some(sp_pr) = sp_pr else { continue; };
            let xfrm_node = sp_pr.children().find(|n| n.is_element() && n.tag_name().name() == "xfrm");
            let Some(xfrm_n) = xfrm_node else { continue; };
            let Some(xfrm) = parse_xfrm(&xfrm_n) else { continue; };
            let rot_raw: f64 = xfrm_n.attribute("rot")
                .and_then(|s| s.parse().ok()).unwrap_or(0.0);

            // Resolve <a:blip r:embed="rIdN"/>. The r:embed attribute lives in
            // the relationships namespace, not the drawingml namespace.
            let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            let pic_rid = child.descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "blip")
                .and_then(|b| {
                    b.attributes()
                        .find(|a| a.name() == "embed" && a.namespace() == Some(r_ns))
                        .map(|a| a.value().to_string())
                });
            let Some(rid) = pic_rid else { continue; };
            let Some(data_url) = rid_urls.get(&rid) else { continue; };

            let root_x = trans_x + scale_x * xfrm.off_x;
            let root_y = trans_y + scale_y * xfrm.off_y;
            let root_w = scale_x * xfrm.ext_x;
            let root_h = scale_y * xfrm.ext_y;
            if root_ext_x == 0.0 || root_ext_y == 0.0 { continue; }
            let nx = (root_x - root_off_x) / root_ext_x;
            let ny = (root_y - root_off_y) / root_ext_y;
            let nw = root_w / root_ext_x;
            let nh = root_h / root_ext_y;
            if nw <= 0.0 || nh <= 0.0 { continue; }

            out.push(ShapeInfo {
                x: nx, y: ny, w: nw, h: nh,
                rot: rot_raw / 60000.0,
                fill_color: None,
                stroke_color: None,
                stroke_width: 0,
                geom: ShapeGeom::Image { data_url: data_url.clone() },
                text: None,
            });
        }
        // Ignore `xdr:cxnSp` / text-only elements for this minimal pass.
    }
}

fn parse_shape_anchors(
    drawing_xml: &str,
    theme_colors: &[String],
    rid_urls: &HashMap<String, String>,
) -> Vec<ShapeAnchor> {
    let Ok(doc) = roxmltree::Document::parse(drawing_xml) else { return Vec::new(); };
    let xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    let mut anchors: Vec<ShapeAnchor> = Vec::new();

    for anchor in doc.descendants() {
        if anchor.tag_name().name() != "twoCellAnchor"
            || anchor.tag_name().namespace() != Some(xdr_ns) { continue; }

        // Parse from/to anchor rect (shared between grpSp and stand-alone sp paths)
        let (mut from_col, mut from_col_off, mut from_row, mut from_row_off) = (0u32, 0i64, 0u32, 0i64);
        let (mut to_col,   mut to_col_off,   mut to_row,   mut to_row_off)   = (0u32, 0i64, 0u32, 0i64);
        // ECMA-376 §20.5.2.33 `twoCellAnchor@editAs` — see ImageAnchor parsing
        // path for semantics. `"oneCell"` instructs the renderer to preserve
        // the group's saved EMU size instead of resizing with the cell rect.
        let edit_as = anchor.attribute("editAs").map(|s| s.to_string());
        let native_ext_cx: i64;
        let native_ext_cy: i64;
        for c in anchor.children() {
            if !c.is_element() { continue; }
            if c.tag_name().name() == "from" || c.tag_name().name() == "to" {
                let is_from = c.tag_name().name() == "from";
                let mut col: u32 = 0; let mut col_off: i64 = 0;
                let mut row: u32 = 0; let mut row_off: i64 = 0;
                for cc in c.children() {
                    match (cc.tag_name().name(), cc.text()) {
                        ("col",    Some(t)) => col     = t.trim().parse().unwrap_or(0),
                        ("colOff", Some(t)) => col_off = t.trim().parse().unwrap_or(0),
                        ("row",    Some(t)) => row     = t.trim().parse().unwrap_or(0),
                        ("rowOff", Some(t)) => row_off = t.trim().parse().unwrap_or(0),
                        _ => {}
                    }
                }
                if is_from {
                    from_col = col; from_col_off = col_off; from_row = row; from_row_off = row_off;
                } else {
                    to_col = col; to_col_off = col_off; to_row = row; to_row_off = row_off;
                }
            }
        }

        // Two top-level layouts ECMA-376 allows under <xdr:twoCellAnchor>:
        //   (a) <xdr:grpSp> wrapping a tree of nested groups + leaves; and
        //   (b) a single <xdr:sp> / <xdr:pic> directly under the anchor
        //       (no grouping wrapper). The grpSp path uses the group's xfrm
        //       to define the anchor's drawing-coord system; the stand-alone
        //       path treats the shape as filling 100 % of the anchor rect.
        let mut shapes: Vec<ShapeInfo> = Vec::new();
        if let Some(grp) = anchor.children().find(|n| n.is_element() && n.tag_name().name() == "grpSp") {
            let grp_sp_pr = grp.children().find(|n| n.is_element() && n.tag_name().name() == "grpSpPr");
            let xfrm = grp_sp_pr
                .and_then(|n| n.children().find(|c| c.is_element() && c.tag_name().name() == "xfrm"))
                .as_ref()
                .and_then(parse_xfrm);
            let Some(root) = xfrm else { continue; };
            if !root.has_ch || root.ch_ext_x == 0.0 || root.ch_ext_y == 0.0 { continue; }

            // Top-level grpSp ext is the group's saved on-sheet EMU size —
            // authoritative when editAs="oneCell".
            native_ext_cx = root.ext_x as i64;
            native_ext_cy = root.ext_y as i64;

            // Map child coords → root coords with the grpSp's own chOff/chExt.
            let csx = root.ext_x / root.ch_ext_x;
            let csy = root.ext_y / root.ch_ext_y;
            let tx = root.off_x - root.ch_off_x * csx;
            let ty = root.off_y - root.ch_off_y * csy;

            collect_shapes(&grp, root.off_x, root.off_y, root.ext_x, root.ext_y,
                           csx, csy, tx, ty, theme_colors, rid_urls, &mut shapes);
        } else if let Some(sp) = anchor.children().find(|n| n.is_element() && (n.tag_name().name() == "sp" || n.tag_name().name() == "pic")) {
            // Stand-alone sp/pic: the shape's own xfrm gives its absolute EMU
            // rect, but for our rendering pipeline the anchor's from/to
            // already defines the on-sheet rect, and the leaf occupies it
            // 100 %. Build a synthetic root coord-system whose origin matches
            // the shape's xfrm so collect_shapes normalizes the leaf to (0,0)
            // (1,1).
            let sp_pr = sp.children().find(|n| n.is_element() && n.tag_name().name() == "spPr");
            let Some(sp_pr_node) = sp_pr else { continue; };
            let xfrm_node = sp_pr_node.children().find(|n| n.is_element() && n.tag_name().name() == "xfrm");
            let Some(xfrm_n) = xfrm_node else { continue; };
            let Some(xfrm) = parse_xfrm(&xfrm_n) else { continue; };
            if xfrm.ext_x == 0.0 || xfrm.ext_y == 0.0 { continue; }
            native_ext_cx = xfrm.ext_x as i64;
            native_ext_cy = xfrm.ext_y as i64;
            collect_shapes(&anchor, xfrm.off_x, xfrm.off_y, xfrm.ext_x, xfrm.ext_y,
                           1.0, 1.0, 0.0, 0.0, theme_colors, rid_urls, &mut shapes);
        } else {
            continue;
        }

        if shapes.is_empty() { continue; }

        anchors.push(ShapeAnchor {
            from_col, from_col_off, from_row, from_row_off,
            to_col, to_col_off, to_row, to_row_off,
            edit_as,
            native_ext_cx,
            native_ext_cy,
            shapes,
        });
    }
    anchors
}

fn load_sheet_shape_groups(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    sheet_path: &str,
    theme_colors: &[String],
) -> Vec<ShapeAnchor> {
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else { return Vec::new(); };
    let sheet_rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let Ok(sheet_rels_xml) = read_zip_entry(archive, &sheet_rels_path) else { return Vec::new(); };
    let Ok(rels_doc) = roxmltree::Document::parse(&sheet_rels_xml) else { return Vec::new(); };
    let mut drawing_targets: Vec<String> = Vec::new();
    for rel in rels_doc.root_element().children().filter(|n| n.is_element()) {
        if rel.attribute("Type").unwrap_or("").ends_with("/drawing") {
            if let Some(t) = rel.attribute("Target") { drawing_targets.push(t.to_string()); }
        }
    }
    let mut all: Vec<ShapeAnchor> = Vec::new();
    for target in drawing_targets {
        let drawing_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        let Ok(drawing_xml) = read_zip_entry(archive, &drawing_path) else { continue; };
        let rid_urls = build_drawing_rid_urls(archive, &drawing_path);
        all.extend(parse_shape_anchors(&drawing_xml, theme_colors, &rid_urls));
    }
    all
}

/// Build a `HashMap<rId, data-URL>` for every image (png/jpg/…) target in
/// a drawing's `.rels` file. Used by `collect_shapes` to resolve `<xdr:pic>`
/// leaves inside a group. Mirrors the logic in `parse_drawing_anchors` but
/// eagerly encodes each referenced image so per-shape lookup is a single
/// HashMap hit.
fn build_drawing_rid_urls(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    drawing_path: &str,
) -> HashMap<String, String> {
    let Some((drawing_dir, drawing_file)) = drawing_path.rsplit_once('/') else {
        return HashMap::new();
    };
    let rels_path = format!("{}/_rels/{}.rels", drawing_dir, drawing_file);
    let rels = read_zip_entry(archive, &rels_path)
        .ok()
        .map(|xml| parse_rels_map(&xml))
        .unwrap_or_default();

    let mut result: HashMap<String, String> = HashMap::new();
    for (rid, target) in rels {
        let lower = target.to_lowercase();
        if !(lower.ends_with(".png") || lower.ends_with(".jpg")
            || lower.ends_with(".jpeg") || lower.ends_with(".gif")
            || lower.ends_with(".bmp")  || lower.ends_with(".webp"))
        {
            continue;
        }
        let media_path = resolve_zip_path(drawing_dir, &target);
        if let Some(bytes) = read_zip_bytes(archive, &media_path) {
            let mime = mime_from_ext(&media_path);
            result.insert(rid, format!("data:{mime};base64,{}", B64.encode(&bytes)));
        }
    }
    result
}

/// Given a sheet path (e.g. "worksheets/sheet1.xml"), locate and parse
/// its drawing(s), and return all image anchors found.
fn load_sheet_images(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    sheet_path: &str, // e.g. "worksheets/sheet1.xml"
) -> Vec<ImageAnchor> {
    // sheet rels path:  xl/worksheets/_rels/sheet1.xml.rels
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else {
        return Vec::new();
    };
    let sheet_rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let Ok(sheet_rels_xml) = read_zip_entry(archive, &sheet_rels_path) else {
        return Vec::new();
    };

    // Find all drawing relationships
    let Ok(rels_doc) = roxmltree::Document::parse(&sheet_rels_xml) else {
        return Vec::new();
    };
    let mut drawing_targets: Vec<String> = Vec::new();
    for rel in rels_doc.root_element().children().filter(|n| n.is_element()) {
        let rel_type = rel.attribute("Type").unwrap_or("");
        if rel_type.ends_with("/drawing") {
            if let Some(t) = rel.attribute("Target") {
                drawing_targets.push(t.to_string());
            }
        }
    }
    if drawing_targets.is_empty() { return Vec::new(); }

    let mut all_anchors: Vec<ImageAnchor> = Vec::new();
    for target in drawing_targets {
        // sheet_dir is "worksheets", target typically "../drawings/drawing1.xml"
        // base dir for the drawing = "xl/worksheets" + "../drawings" → "xl/drawings"
        let drawing_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        let Ok(drawing_xml) = read_zip_entry(archive, &drawing_path) else { continue; };
        // Drawing rels:  xl/drawings/_rels/drawing1.xml.rels
        let Some((drawing_dir, drawing_file)) = drawing_path.rsplit_once('/') else { continue; };
        let drawing_rels_path = format!("{}/_rels/{}.rels", drawing_dir, drawing_file);
        let drawing_rels = read_zip_entry(archive, &drawing_rels_path)
            .ok()
            .map(|xml| parse_rels_map(&xml))
            .unwrap_or_default();

        let mut anchors = parse_drawing_anchors(&drawing_xml, &drawing_rels, drawing_dir, archive);
        all_anchors.append(&mut anchors);
    }
    all_anchors
}

// ─── Slicer loading ─────────────────────────────────────────────────────────
//
// Office 2010+ extension (`sle:slicer` inside `<mc:AlternateContent>`).
// Resolving one slicer graphicFrame into a drawable anchor takes four
// XML files:
//   1. The sheet's drawing (for the anchor rect + graphicFrame name).
//   2. `xl/slicers/slicerN.xml` — slicer definition: graphicFrame name →
//      caption + cache name.
//   3. `xl/slicerCaches/slicerCacheN.xml` — cache definition: cache name →
//      source field + list of (item index, selected?).
//   4. `xl/pivotCache/pivotCacheDefinitionN.xml` — pivot cache: field name →
//      ordered string values.
// Excel also allows slicers bound to Excel Tables (`tableSlicerCache`), but
// the present sample is pivot-only; we only implement the pivot path.

#[derive(Default)]
struct SlicerCacheInfo {
    source_name: String,
    items: Vec<(u32, bool)>, // (index into pivot field, selected)
}

#[derive(Default)]
struct PivotCacheFields {
    by_name: HashMap<String, Vec<String>>, // field name → ordered string items
}

/// Parse every `xl/pivotCache/pivotCacheDefinition*.xml` and merge its
/// cacheFields (indexed by `@name`) into a single map. Sample workbooks
/// typically have one pivotCache but the loop keeps the code general.
fn load_all_pivot_cache_fields(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
) -> PivotCacheFields {
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let mut out = PivotCacheFields::default();
    let names: Vec<String> = (0..archive.len())
        .filter_map(|i| archive.by_index(i).ok().map(|f| f.name().to_string()))
        .filter(|n| n.starts_with("xl/pivotCache/pivotCacheDefinition") && n.ends_with(".xml"))
        .collect();
    for name in names {
        let Ok(xml) = read_zip_entry(archive, &name) else { continue; };
        let Ok(doc) = roxmltree::Document::parse(&xml) else { continue; };
        for field in doc.descendants() {
            if field.tag_name().name() != "cacheField"
                || field.tag_name().namespace() != Some(ns)
            { continue; }
            let Some(field_name) = field.attribute("name") else { continue; };
            let mut items: Vec<String> = Vec::new();
            for shared in field.children().filter(|n| n.is_element() && n.tag_name().name() == "sharedItems") {
                for item in shared.children().filter(|n| n.is_element()) {
                    match item.tag_name().name() {
                        "s" => items.push(item.attribute("v").unwrap_or("").to_string()),
                        "n" => items.push(item.attribute("v").unwrap_or("").to_string()),
                        "d" => items.push(item.attribute("v").unwrap_or("").to_string()),
                        "b" => items.push(item.attribute("v").unwrap_or("").to_string()),
                        "m" => items.push(String::new()),
                        _ => {}
                    }
                }
            }
            if !items.is_empty() {
                out.by_name.insert(field_name.to_string(), items);
            }
        }
    }
    out
}

/// Parse every `xl/slicerCaches/slicerCache*.xml` and build a map keyed by
/// the slicerCache's `@name` attribute (e.g. `"スライサー_贈答相手1"`). That
/// name is what `<slicer cache="…"/>` in `xl/slicers/slicerN.xml` references.
fn load_all_slicer_caches(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
) -> HashMap<String, SlicerCacheInfo> {
    let mut out: HashMap<String, SlicerCacheInfo> = HashMap::new();
    let names: Vec<String> = (0..archive.len())
        .filter_map(|i| archive.by_index(i).ok().map(|f| f.name().to_string()))
        .filter(|n| n.starts_with("xl/slicerCaches/slicerCache") && n.ends_with(".xml"))
        .collect();
    for path in names {
        let Ok(xml) = read_zip_entry(archive, &path) else { continue; };
        let Ok(doc) = roxmltree::Document::parse(&xml) else { continue; };
        let root = doc.root_element();
        let cache_name = root.attribute("name").unwrap_or("").to_string();
        let source_name = root.attribute("sourceName").unwrap_or("").to_string();
        let mut items: Vec<(u32, bool)> = Vec::new();
        for tabular in doc.descendants().filter(|n| n.is_element() && n.tag_name().name() == "tabular") {
            for i_el in tabular.descendants().filter(|n| n.is_element() && n.tag_name().name() == "i") {
                let x: u32 = i_el.attribute("x").and_then(|v| v.parse().ok()).unwrap_or(0);
                // `s` defaults to "1" (selected) when absent — ECMA-376
                // extension schema for slicer caches.
                let selected = i_el.attribute("s").map(|v| v != "0").unwrap_or(true);
                items.push((x, selected));
            }
        }
        if !cache_name.is_empty() {
            out.insert(cache_name, SlicerCacheInfo { source_name, items });
        }
    }
    out
}

/// Slicer definition (`xl/slicers/slicerN.xml`): maps each graphicFrame name
/// on the sheet to its display caption and the slicerCache it's backed by.
#[derive(Default)]
struct SlicerDef {
    caption: String,
    cache: String,
}

fn parse_slicers_xml(xml: &str) -> HashMap<String, SlicerDef> {
    let mut out: HashMap<String, SlicerDef> = HashMap::new();
    let Ok(doc) = roxmltree::Document::parse(xml) else { return out; };
    for slicer in doc.descendants().filter(|n| n.is_element() && n.tag_name().name() == "slicer") {
        let name = slicer.attribute("name").unwrap_or("").to_string();
        let caption = slicer.attribute("caption").unwrap_or("").to_string();
        let cache = slicer.attribute("cache").unwrap_or("").to_string();
        if !name.is_empty() {
            out.insert(name, SlicerDef { caption, cache });
        }
    }
    out
}

fn load_sheet_slicers(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    sheet_path: &str, // e.g. "worksheets/sheet1.xml"
) -> Vec<SlicerAnchor> {
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else {
        return Vec::new();
    };
    let sheet_rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let Ok(sheet_rels_xml) = read_zip_entry(archive, &sheet_rels_path) else {
        return Vec::new();
    };
    let Ok(rels_doc) = roxmltree::Document::parse(&sheet_rels_xml) else {
        return Vec::new();
    };

    // 1. Collect slicer-definition and drawing targets from the sheet rels.
    let mut drawing_targets: Vec<String> = Vec::new();
    let mut slicer_targets: Vec<String> = Vec::new();
    for rel in rels_doc.root_element().children().filter(|n| n.is_element()) {
        let rel_type = rel.attribute("Type").unwrap_or("");
        let Some(target) = rel.attribute("Target") else { continue; };
        if rel_type.ends_with("/drawing") {
            drawing_targets.push(target.to_string());
        } else if rel_type.ends_with("/slicer") {
            slicer_targets.push(target.to_string());
        }
    }
    if drawing_targets.is_empty() || slicer_targets.is_empty() {
        return Vec::new();
    }

    // 2. Parse all slicer definitions referenced by this sheet, keyed by
    //    graphicFrame name.
    let mut slicer_defs: HashMap<String, SlicerDef> = HashMap::new();
    for target in &slicer_targets {
        let slicer_path = resolve_zip_path(&format!("xl/{}", sheet_dir), target);
        let Ok(xml) = read_zip_entry(archive, &slicer_path) else { continue; };
        for (k, v) in parse_slicers_xml(&xml) {
            slicer_defs.insert(k, v);
        }
    }
    if slicer_defs.is_empty() { return Vec::new(); }

    // 3. Resolve caches (and their backing pivot fields) once.
    let slicer_caches = load_all_slicer_caches(archive);
    let pivot_fields = load_all_pivot_cache_fields(archive);

    // 4. Walk each drawing and pick up slicer graphicFrames.
    let mut out: Vec<SlicerAnchor> = Vec::new();
    for target in drawing_targets {
        let drawing_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        let Ok(drawing_xml) = read_zip_entry(archive, &drawing_path) else { continue; };
        out.extend(parse_slicer_anchors(&drawing_xml, &slicer_defs, &slicer_caches, &pivot_fields));
    }
    out
}

fn parse_slicer_anchors(
    drawing_xml: &str,
    slicer_defs: &HashMap<String, SlicerDef>,
    slicer_caches: &HashMap<String, SlicerCacheInfo>,
    pivot_fields: &PivotCacheFields,
) -> Vec<SlicerAnchor> {
    let Ok(doc) = roxmltree::Document::parse(drawing_xml) else {
        return Vec::new();
    };
    let xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    let mc_ns = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    let slicer_uri = "http://schemas.microsoft.com/office/drawing/2010/slicer";
    let mut out: Vec<SlicerAnchor> = Vec::new();

    for anchor in doc.descendants() {
        if anchor.tag_name().name() != "twoCellAnchor"
            || anchor.tag_name().namespace() != Some(xdr_ns)
        { continue; }

        // Anchor rect.
        let mut from = (0u32, 0i64, 0u32, 0i64);
        let mut to   = (0u32, 0i64, 0u32, 0i64);
        for child in anchor.children().filter(|n| n.is_element()) {
            match child.tag_name().name() {
                "from" | "to" => {
                    let is_from = child.tag_name().name() == "from";
                    let mut col: u32 = 0; let mut col_off: i64 = 0;
                    let mut row: u32 = 0; let mut row_off: i64 = 0;
                    for c in child.children() {
                        match (c.tag_name().name(), c.text()) {
                            ("col",    Some(t)) => col     = t.trim().parse().unwrap_or(0),
                            ("colOff", Some(t)) => col_off = t.trim().parse().unwrap_or(0),
                            ("row",    Some(t)) => row     = t.trim().parse().unwrap_or(0),
                            ("rowOff", Some(t)) => row_off = t.trim().parse().unwrap_or(0),
                            _ => {}
                        }
                    }
                    if is_from { from = (col, col_off, row, row_off); }
                    else       { to   = (col, col_off, row, row_off); }
                }
                _ => {}
            }
        }

        // Slicers live inside `<mc:AlternateContent><mc:Choice>` — descend
        // until we find a `<xdr:graphicFrame>` whose graphicData uri is the
        // 2010 slicer namespace, then harvest the graphicFrame's cNvPr name.
        let Some(frame_name) = anchor.descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "Choice" && n.tag_name().namespace() == Some(mc_ns))
            .flat_map(|choice| choice.descendants())
            .find_map(|n| {
                if n.is_element()
                    && n.tag_name().name() == "graphicData"
                    && n.attribute("uri") == Some(slicer_uri)
                {
                    // graphicData → ancestor graphicFrame → nvGraphicFramePr → cNvPr
                    let mut p = n.parent();
                    while let Some(pp) = p {
                        if pp.tag_name().name() == "graphicFrame" { break; }
                        p = pp.parent();
                    }
                    let frame = p?;
                    let cnvpr = frame.descendants()
                        .find(|d| d.is_element() && d.tag_name().name() == "cNvPr")?;
                    cnvpr.attribute("name").map(|s| s.to_string())
                } else { None }
            }) else { continue };

        let Some(slicer_def) = slicer_defs.get(&frame_name) else { continue; };

        // Resolve items via cache → pivot field; fall back to an empty list
        // if any link is broken (still renders the header and box).
        let items: Vec<SlicerItem> = slicer_caches.get(&slicer_def.cache)
            .map(|cache| {
                let field_items = pivot_fields.by_name.get(&cache.source_name);
                cache.items.iter().map(|(x, selected)| {
                    let name = field_items
                        .and_then(|list| list.get(*x as usize))
                        .cloned()
                        .unwrap_or_default();
                    SlicerItem { name, selected: *selected }
                }).collect()
            })
            .unwrap_or_default();

        let caption = if !slicer_def.caption.is_empty() {
            slicer_def.caption.clone()
        } else {
            frame_name.clone()
        };

        out.push(SlicerAnchor {
            from_col: from.0, from_col_off: from.1, from_row: from.2, from_row_off: from.3,
            to_col:   to.0,   to_col_off:   to.1,   to_row:   to.2,   to_row_off:   to.3,
            caption,
            items,
        });
    }
    out
}

// ─── Chart loading ──────────────────────────────────────────────────────────


pub(crate) fn resolve_fill_color(fill_node: &roxmltree::Node, theme_colors: &[String]) -> Option<String> {
    // Accept either a `<a:solidFill>` directly or a `<c:spPr>` whose first
    // fill-ish child is `<a:solidFill>`. Looking at *direct* children (not
    // descendants) is intentional — chart series often carry label/axis text
    // colors under `c:dLbls`/`c:txPr` which must NOT be misread as fill.
    let solid = if fill_node.tag_name().name() == "solidFill" {
        Some(*fill_node)
    } else {
        fill_node.children().find(|n| n.is_element() && n.tag_name().name() == "solidFill")
    }?;
    for n in solid.children().filter(|n| n.is_element()) {
        let tag = n.tag_name().name();
        if tag == "srgbClr" {
            if let Some(v) = n.attribute("val") {
                return Some(v.to_lowercase());
            }
        }
        if tag == "schemeClr" {
            if let Some(v) = n.attribute("val") {
                let idx = match v {
                    "dk1"  | "tx1" => Some(0),
                    "lt1"  | "bg1" => Some(1),
                    "dk2"  | "tx2" => Some(2),
                    "lt2"  | "bg2" => Some(3),
                    "accent1" => Some(4),
                    "accent2" => Some(5),
                    "accent3" => Some(6),
                    "accent4" => Some(7),
                    "accent5" => Some(8),
                    "accent6" => Some(9),
                    "hlink"    => Some(10),
                    "folHlink" => Some(11),
                    _ => None,
                };
                if let Some(i) = idx {
                    if let Some(c) = theme_colors.get(i) {
                        return Some(c.trim_start_matches('#').to_lowercase());
                    }
                }
            }
        }
    }
    None
}

fn parse_sqref(s: &str) -> Vec<CellRange> {
    s.split_whitespace().map(|range_str| {
        if let Some((a, b)) = range_str.split_once(':') {
            let (left, top) = parse_cell_ref(a);
            let (right, bottom) = parse_cell_ref(b);
            CellRange { top, left, bottom, right }
        } else {
            let (col, row) = parse_cell_ref(range_str);
            CellRange { top: row, left: col, bottom: row, right: col }
        }
    }).collect()
}

/// Split an `<xm:f>` reference like `Sheet1!A1:A10` or `'My Sheet'!$B$3:$B$8`
/// into `(sheet_name, range)`. Returns `None` if the reference has no sheet
/// qualifier — sparkline data refs always do, so unqualified is treated as
/// "same sheet" by callers.
fn split_sheet_ref(s: &str) -> (Option<String>, String) {
    let s = s.trim();
    let Some(bang) = s.rfind('!') else { return (None, s.to_string()); };
    let mut sheet = s[..bang].to_string();
    // Strip absolute-ref dollars from the range part.
    let range = s[bang + 1..].replace('$', "");
    // Quoted sheet names ('foo''s sheet' uses doubled quotes for inner ').
    if sheet.starts_with('\'') && sheet.ends_with('\'') {
        sheet = sheet[1..sheet.len() - 1].replace("''", "'");
    }
    (Some(sheet), range)
}

/// Read a worksheet XML and extract numeric `<v>` values for the cells in
/// `range`. Returns one value per cell in row-major order across the range.
/// Empty cells, non-numeric values, and cells outside the range yield `None`.
///
/// This is intentionally lighter than `parse_row_cells`: sparklines only need
/// raw numbers, no styles, formulas, or shared strings.
fn extract_range_values(sheet_xml: &str, range: &CellRange) -> Vec<Option<f64>> {
    let total = ((range.bottom - range.top + 1) as usize)
        .saturating_mul((range.right - range.left + 1) as usize);
    let mut values: Vec<Option<f64>> = vec![None; total];
    let Ok(doc) = roxmltree::Document::parse(sheet_xml) else { return values; };
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let row_span = (range.right - range.left + 1) as usize;
    for c in doc.descendants().filter(|n| n.tag_name().name() == "c" && n.tag_name().namespace() == Some(ns)) {
        let Some(r_attr) = c.attribute("r") else { continue };
        let (col, row) = parse_cell_ref(r_attr);
        if row < range.top || row > range.bottom || col < range.left || col > range.right {
            continue;
        }
        // Only honor numeric / formula-numeric cells. `t` of "s" / "str" /
        // "inlineStr" / "b" / "e" all map to None for sparkline values.
        let t = c.attribute("t").unwrap_or("");
        if matches!(t, "s" | "str" | "inlineStr" | "b" | "e") { continue; }
        let v = c.children()
            .find(|n| n.tag_name().name() == "v" && n.tag_name().namespace() == Some(ns))
            .and_then(|n| n.text())
            .and_then(|s| s.trim().parse::<f64>().ok());
        if let Some(num) = v {
            let idx = (row - range.top) as usize * row_span + (col - range.left) as usize;
            if idx < values.len() {
                values[idx] = Some(num);
            }
        }
    }
    values
}

/// Walk the worksheet XML's `<extLst>` and produce one `SparklineGroup` per
/// `<x14:sparklineGroup>`. Resolves cross-sheet `<xm:f>` data references by
/// reading the referenced sheet from the archive (cached per call to avoid
/// re-reads). Theme colors are flattened to `#RRGGBB` via `parse_color`.
fn load_sheet_sparklines(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    sheet_xml: &str,
    sheets: &[SheetMeta],
    rels_doc: &roxmltree::Document,
    theme_colors: &[String],
) -> Vec<SparklineGroup> {
    let Ok(doc) = roxmltree::Document::parse(sheet_xml) else { return Vec::new(); };
    let mut groups: Vec<SparklineGroup> = Vec::new();
    // Cache: sheet name → loaded XML. Saves re-reading when many sparklines
    // reference the same source sheet (typical: one "data" sheet feeds many
    // dashboard sparklines).
    let mut xml_cache: HashMap<String, Option<String>> = HashMap::new();

    let parse_bool_attr = |n: &roxmltree::Node, key: &str, default: bool| -> bool {
        match n.attribute(key) {
            Some(v) => v == "1" || v.eq_ignore_ascii_case("true"),
            None => default,
        }
    };
    let parse_f64_attr = |n: &roxmltree::Node, key: &str| -> Option<f64> {
        n.attribute(key).and_then(|v| v.parse::<f64>().ok())
    };

    for group_node in doc.descendants().filter(|n| n.tag_name().name() == "sparklineGroup") {
        let kind = match group_node.attribute("type").unwrap_or("line") {
            "column" => SparklineType::Column,
            "stacked" => SparklineType::Stem,  // historical alias
            "stem" => SparklineType::Stem,
            // ECMA-376 lists `line` and a planned `stairStep`; treat unknown
            // types as line (closest visual fallback).
            _ => SparklineType::Line,
        };

        let resolve_color = |child_name: &str| -> Option<String> {
            group_node.children()
                .find(|n| n.is_element() && n.tag_name().name() == child_name)
                .and_then(|n| parse_color(&n, theme_colors))
        };

        let mut sparklines: Vec<Sparkline> = Vec::new();
        // <x14:sparklines> is the wrapper; <x14:sparkline> are the children.
        for sparklines_node in group_node.children().filter(|n| n.is_element() && n.tag_name().name() == "sparklines") {
            for sl in sparklines_node.children().filter(|n| n.is_element() && n.tag_name().name() == "sparkline") {
                let f_text = sl.children()
                    .find(|n| n.is_element() && n.tag_name().name() == "f")
                    .and_then(|n| n.text())
                    .unwrap_or("");
                let sqref_text = sl.children()
                    .find(|n| n.is_element() && n.tag_name().name() == "sqref")
                    .and_then(|n| n.text())
                    .unwrap_or("");
                if f_text.is_empty() || sqref_text.is_empty() { continue; }
                let (col, row) = parse_cell_ref(sqref_text.trim());
                let (source_sheet, range_str) = split_sheet_ref(f_text);
                let ranges = parse_sqref(&range_str);
                let Some(range) = ranges.into_iter().next() else { continue };

                // Look up source sheet XML (cross-sheet ref). When the ref
                // has no sheet qualifier, fall back to the *current* sheet
                // XML.
                let source_xml: Option<&str> = match source_sheet {
                    Some(name) => {
                        if !xml_cache.contains_key(&name) {
                            let path = sheets.iter()
                                .find(|s| s.name == name)
                                .and_then(|s| resolve_sheet_path(rels_doc, &s.r_id))
                                .map(|p| format!("xl/{}", p));
                            let xml = path.and_then(|p| read_zip_entry(archive, &p).ok());
                            xml_cache.insert(name.clone(), xml);
                        }
                        xml_cache.get(&name).and_then(|o| o.as_deref())
                    }
                    None => Some(sheet_xml),
                };
                let values = source_xml
                    .map(|xml| extract_range_values(xml, &range))
                    .unwrap_or_default();

                sparklines.push(Sparkline { row, col, values });
            }
        }

        groups.push(SparklineGroup {
            kind,
            markers: parse_bool_attr(&group_node, "markers", false),
            high: parse_bool_attr(&group_node, "high", false),
            low: parse_bool_attr(&group_node, "low", false),
            first: parse_bool_attr(&group_node, "first", false),
            last: parse_bool_attr(&group_node, "last", false),
            negative: parse_bool_attr(&group_node, "negative", false),
            display_x_axis: parse_bool_attr(&group_node, "displayXAxis", false),
            display_empty_cells_as: group_node.attribute("displayEmptyCellsAs").unwrap_or("gap").to_string(),
            min_axis_type: group_node.attribute("minAxisType").unwrap_or("individual").to_string(),
            max_axis_type: group_node.attribute("maxAxisType").unwrap_or("individual").to_string(),
            manual_min: parse_f64_attr(&group_node, "manualMin"),
            manual_max: parse_f64_attr(&group_node, "manualMax"),
            line_weight: parse_f64_attr(&group_node, "lineWeight").unwrap_or(0.75),
            color_series: resolve_color("colorSeries"),
            color_negative: resolve_color("colorNegative"),
            color_axis: resolve_color("colorAxis"),
            color_markers: resolve_color("colorMarkers"),
            color_first: resolve_color("colorFirst"),
            color_last: resolve_color("colorLast"),
            color_high: resolve_color("colorHigh"),
            color_low: resolve_color("colorLow"),
            sparklines,
        });
    }
    groups
}

fn parse_row_cells(
    row_node: &roxmltree::Node,
    shared_strings: &[SharedString],
    theme_colors: &[String],
    ns: &str,
) -> Vec<Cell> {
    let mut cells = Vec::new();
    for c_node in row_node.children() {
        if c_node.tag_name().name() != "c" || c_node.tag_name().namespace() != Some(ns) {
            continue;
        }
        let cell_ref = c_node.attribute("r").unwrap_or("A1").to_string();
        let (col, row) = parse_cell_ref(&cell_ref);
        let cell_type = c_node.attribute("t").unwrap_or("");
        let style_index: u32 = c_node.attribute("s").and_then(|s| s.parse().ok()).unwrap_or(0);

        // Inline string: <c t="inlineStr"><is>...</is></c>
        let is_node = c_node.children().find(|n| n.tag_name().name() == "is");

        // Formula text, if any (<f>…</f>). Kept so the renderer can
        // recompute volatile builtins (TODAY, NOW) at display time.
        let formula: Option<String> = c_node
            .children()
            .find(|n| n.tag_name().name() == "f")
            .and_then(|n| n.text())
            .map(|s| s.trim().to_string())
            .filter(|s| !s.is_empty());

        let v_text = c_node
            .children()
            .find(|n| n.tag_name().name() == "v")
            .and_then(|n| n.text())
            .unwrap_or("")
            .to_string();

        let value = if cell_type == "inlineStr" {
            match is_node {
                Some(is) => {
                    let ss = parse_si_node(&is, ns, theme_colors);
                    CellValue::Text { text: ss.text, runs: ss.runs }
                }
                None => CellValue::Empty,
            }
        } else if v_text.is_empty() {
            CellValue::Empty
        } else {
            match cell_type {
                "s" => {
                    let idx: usize = v_text.parse().unwrap_or(0);
                    if let Some(ss) = shared_strings.get(idx) {
                        CellValue::Text { text: ss.text.clone(), runs: ss.runs.clone() }
                    } else {
                        CellValue::Text { text: String::new(), runs: None }
                    }
                }
                "str" => CellValue::Text { text: v_text, runs: None },
                "b" => CellValue::Bool { bool: v_text == "1" || v_text == "true" },
                "e" => CellValue::Error { error: v_text },
                _ => {
                    if let Ok(n) = v_text.parse::<f64>() {
                        CellValue::Number { number: n }
                    } else {
                        CellValue::Text { text: v_text, runs: None }
                    }
                }
            }
        };

        cells.push(Cell { col, row, col_ref: cell_ref, value, style_index, formula });
    }
    cells
}

fn parse_cell_ref(r: &str) -> (u32, u32) {
    let col_str: String = r.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
    let row_str: String = r.chars().skip_while(|c| c.is_ascii_alphabetic()).collect();
    let col = col_str.chars().fold(0u32, |acc, c| acc * 26 + (c as u32 - 'A' as u32 + 1));
    let row = row_str.parse().unwrap_or(1);
    (col, row)
}

// ===========================
//  Native (non-WASM) API
// ===========================

/// Returns workbook overview (sheet names and metadata) as JSON.
/// Native equivalent of `parse_xlsx` for use from the MCP server.

pub fn parse_workbook_native(data: &[u8]) -> Result<String, String> {
    parse_xlsx_inner(data)
        .and_then(|wb| serde_json::to_string(&wb.workbook).map_err(|e| e.to_string()))
}

/// Parse the workbook and project every sheet to GitHub-flavoured markdown:
/// `## SheetName` headings followed by a pipe table per sheet. Merged-cell
/// continuation cells are rendered as empty; the display value comes from the
/// WASM-callable markdown projection (mirrors `to_markdown_native`).
#[wasm_bindgen]
pub fn xlsx_to_markdown(data: &[u8], max_zip_entry_bytes: Option<u64>) -> Result<String, JsValue> {
    console_error_panic_hook::set_once();
    let _guard = ooxml_common::zip::scoped_max(max_zip_entry_bytes);
    to_markdown_impl(data).map_err(|e| JsValue::from_str(&e))
}

/// cached `<v>` so formula formulas show their results, not the formula text.
/// Designed for AI agents that need to read the spreadsheet content
/// efficiently — drops styling, formatting, charts, sparklines, drawings.

pub fn to_markdown_native(data: &[u8]) -> Result<String, String> {
    to_markdown_impl(data)
}

/// Shared implementation between `to_markdown_native` (mcp-server) and
/// `xlsx_to_markdown` (browser / Node WASM).
fn to_markdown_impl(data: &[u8]) -> Result<String, String> {
    let cursor = Cursor::new(data);
    let mut archive = zip::ZipArchive::new(cursor).map_err(|e| e.to_string())?;
    let workbook_xml = read_zip_entry(&mut archive, "xl/workbook.xml")?;
    let wb_doc = roxmltree::Document::parse(&workbook_xml).map_err(|e| e.to_string())?;
    let sheets = parse_workbook_sheets(&wb_doc);

    let mut out = String::new();
    for (idx, sheet_meta) in sheets.iter().enumerate() {
        let sheet_json =
            parse_sheet_native(data, idx as u32, &sheet_meta.name).map_err(|e| {
                format!("sheet '{}' (#{}) parse failed: {}", sheet_meta.name, idx, e)
            })?;
        let sheet: serde_json::Value =
            serde_json::from_str(&sheet_json).map_err(|e| e.to_string())?;
        markdown::render_sheet(&sheet, &mut out);
    }
    Ok(out)
}

/// Parses a single worksheet by 0-based index and returns it as JSON.
/// Native equivalent of `parse_sheet` for use from the MCP server.

pub fn parse_sheet_native(data: &[u8], sheet_index: u32, name: &str) -> Result<String, String> {
    let cursor = Cursor::new(data);
    let mut archive = zip::ZipArchive::new(cursor).map_err(|e| e.to_string())?;

    let workbook_xml = read_zip_entry(&mut archive, "xl/workbook.xml")?;
    let wb_doc = roxmltree::Document::parse(&workbook_xml).map_err(|e| e.to_string())?;
    let sheets = parse_workbook_sheets(&wb_doc);

    let sheet_meta = sheets
        .get(sheet_index as usize)
        .ok_or_else(|| format!("sheet index {} out of range", sheet_index))?;

    let rels_xml = read_zip_entry(&mut archive, "xl/_rels/workbook.xml.rels")?;
    let rels_doc = roxmltree::Document::parse(&rels_xml).map_err(|e| e.to_string())?;
    let sheet_path = resolve_sheet_path(&rels_doc, &sheet_meta.r_id)
        .ok_or_else(|| format!("rId {} not found in rels", sheet_meta.r_id))?;

    let theme_colors = parse_theme_colors(&mut archive);
    let shared_strings = read_shared_strings(&mut archive, &theme_colors);
    let sheet_xml = read_zip_entry(&mut archive, &format!("xl/{}", sheet_path))?;
    let (mut ws, hyperlink_rids) =
        parse_worksheet(&sheet_xml, &shared_strings, &theme_colors, name)
            .map_err(|e| e.to_string())?;

    ws.images = load_sheet_images(&mut archive, &sheet_path);
    ws.charts = load_sheet_charts(&mut archive, &sheet_path, &theme_colors);
    ws.shape_groups = load_sheet_shape_groups(&mut archive, &sheet_path, &theme_colors);
    ws.hyperlinks = load_hyperlinks(&mut archive, &sheet_path, hyperlink_rids);
    ws.comments = load_sheet_comments(&mut archive, &sheet_path);
    ws.comment_refs = ws.comments.iter().map(|c| c.cell_ref.clone()).collect();
    ws.defined_names = parse_defined_names_for_sheet(&wb_doc, sheet_index);
    ws.tables = load_sheet_tables(&mut archive, &sheet_path, &theme_colors);
    ws.slicers = load_sheet_slicers(&mut archive, &sheet_path);
    ws.sparkline_groups = load_sheet_sparklines(&mut archive, &sheet_xml, &sheets, &rels_doc, &theme_colors);
    let (df_family, df_size) = parse_default_font(&mut archive);
    ws.default_font_family = df_family;
    ws.default_font_size = df_size;

    serde_json::to_string(&ws).map_err(|e| e.to_string())
}
