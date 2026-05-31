use wasm_bindgen::prelude::*;
use std::collections::HashMap;
use std::io::{Cursor, Read};


mod markdown;

mod types;
pub use types::*;
mod styles;
use styles::*;
mod chart;
use chart::*;
mod drawing;
use drawing::*;
mod slicer;
use slicer::*;


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
pub(crate) fn read_zip_bytes(archive: &mut zip::ZipArchive<Cursor<&[u8]>>, path: &str) -> Option<Vec<u8>> {
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

pub(crate) fn mime_from_ext(path: &str) -> &'static str {
    match path.rsplit('.').next().unwrap_or("").to_ascii_lowercase().as_str() {
        "png"  => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif"  => "image/gif",
        "bmp"  => "image/bmp",
        "webp" => "image/webp",
        _      => "application/octet-stream",
    }
}

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
