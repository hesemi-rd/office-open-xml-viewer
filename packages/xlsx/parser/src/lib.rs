use std::collections::HashMap;
use std::io::{Cursor, Read};
use wasm_bindgen::prelude::*;

use ooxml_common::zip::read_zip_string;

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
mod table;
use table::*;

// Excel built-in indexed color palette (indices 0-63)
// Standard Excel 2003 color palette
const INDEXED_COLORS: &[&str] = &[
    "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF",
    "#00FFFF", // 0-7
    "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF",
    "#00FFFF", // 8-15
    "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0",
    "#808080", // 16-23
    "#9999FF", "#993366", "#FFFFCC", "#CCFFFF", "#660066", "#FF8080", "#0066CC",
    "#CCCCFF", // 24-31
    "#000080", "#FF00FF", "#FFFF00", "#00FFFF", "#800080", "#800000", "#008080",
    "#0000FF", // 32-39
    "#00CCFF", "#CCFFFF", "#CCFFCC", "#FFFF99", "#99CCFF", "#FF99CC", "#CC99FF",
    "#FFCC99", // 40-47
    "#3366FF", "#33CCCC", "#99CC00", "#FFCC00", "#FF9900", "#FF6600", "#666699",
    "#969696", // 48-55
    "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399",
    "#333333", // 56-63
];

/// Parse a xlsx archive's workbook index and return it as UTF-8 JSON **bytes**.
///
/// Returning `Vec<u8>` (a fresh copy on the JS side) instead of `String` keeps
/// the model out of the JsString/UTF-16 representation: the worker forwards the
/// resulting `ArrayBuffer` to the main thread as a transferable and the main
/// thread does a single `TextDecoder.decode` + `JSON.parse`, collapsing three
/// serializations (Rust String → JsString → structured clone) into one decode.
#[wasm_bindgen]
pub fn parse_xlsx(data: &[u8], max_zip_entry_bytes: Option<u64>) -> Result<Vec<u8>, JsValue> {
    console_error_panic_hook::set_once();
    let _guard = ooxml_common::zip::scoped_max(max_zip_entry_bytes);
    let wb = parse_xlsx_inner(data).map_err(|e| JsValue::from_str(&e))?;
    serde_json::to_vec(&wb).map_err(|e| JsValue::from_str(&format!("serialize error: {e}")))
}

/// Parse one worksheet's cell data + layout and return it as UTF-8 JSON
/// **bytes** (see `parse_xlsx` for the bytes-return rationale).
#[wasm_bindgen]
pub fn parse_sheet(
    data: &[u8],
    sheet_index: u32,
    name: &str,
    max_zip_entry_bytes: Option<u64>,
) -> Result<Vec<u8>, JsValue> {
    console_error_panic_hook::set_once();
    let _guard = ooxml_common::zip::scoped_max(max_zip_entry_bytes);
    let cursor = Cursor::new(data);
    let mut archive = zip::ZipArchive::new(cursor).map_err(|e| e.to_string())?;

    let workbook_xml = read_zip_string(&mut archive, "xl/workbook.xml")?;
    let wb_doc = roxmltree::Document::parse(&workbook_xml).map_err(|e| e.to_string())?;
    let sheets = parse_workbook_sheets(&wb_doc);

    let sheet_meta = sheets
        .get(sheet_index as usize)
        .ok_or_else(|| format!("sheet index {} out of range", sheet_index))?;

    // Resolve rId → target path from workbook.xml.rels
    let rels_xml = read_zip_string(&mut archive, "xl/_rels/workbook.xml.rels")?;
    let rels_doc = roxmltree::Document::parse(&rels_xml).map_err(|e| e.to_string())?;
    let sheet_path = resolve_sheet_path(&rels_doc, &sheet_meta.r_id)
        .ok_or_else(|| format!("rId {} not found in rels", sheet_meta.r_id))?;

    let theme_colors = parse_theme_colors(&mut archive);
    let shared_strings = read_shared_strings(&mut archive, &theme_colors);
    let sheet_xml = read_zip_string(&mut archive, &format!("xl/{}", sheet_path))?;
    let (mut ws, hyperlink_rids) =
        parse_worksheet(&sheet_xml, &shared_strings, &theme_colors, name)
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
    ws.sparkline_groups =
        load_sheet_sparklines(&mut archive, &sheet_xml, &sheets, &rels_doc, &theme_colors);
    let (df_family, df_size) = parse_default_font(&mut archive);
    ws.default_font_family = df_family;
    ws.default_font_size = df_size;

    serde_json::to_vec(&ws).map_err(|e| JsValue::from_str(&e.to_string()))
}

fn parse_xlsx_inner(data: &[u8]) -> Result<ParsedWorkbook, String> {
    let cursor = Cursor::new(data);
    let mut archive = zip::ZipArchive::new(cursor).map_err(|e| e.to_string())?;

    let workbook_xml = read_zip_string(&mut archive, "xl/workbook.xml")?;
    let wb_doc = roxmltree::Document::parse(&workbook_xml).map_err(|e| e.to_string())?;
    let sheets = parse_workbook_sheets(&wb_doc);

    let theme_colors = parse_theme_colors(&mut archive);
    let shared_strings = read_shared_strings(&mut archive, &theme_colors);
    let styles = parse_styles(&mut archive, &theme_colors)?;

    // Surface each sheet's tab color (`<sheetPr><tabColor>`) on the workbook
    // sheet list so the viewer can paint every tab up front. `<sheetPr>` is the
    // first child of `<worksheet>` (ECMA-376 §18.3.1.99 element order), so a
    // small head read of each sheet entry is enough — we never decompress the
    // (potentially huge) `<sheetData>` body just to read the tab color.
    let mut sheets = sheets;
    if let Ok(rels_xml) = read_zip_string(&mut archive, "xl/_rels/workbook.xml.rels") {
        if let Ok(rels_doc) = roxmltree::Document::parse(&rels_xml) {
            for sheet in sheets.iter_mut() {
                let Some(path) = resolve_sheet_path(&rels_doc, &sheet.r_id) else {
                    continue;
                };
                let Ok(head) = read_zip_entry_head(&mut archive, &format!("xl/{}", path), 16_384)
                else {
                    continue;
                };
                sheet.tab_color = extract_tab_color_from_head(&head, &theme_colors);
            }
        }
    }

    Ok(ParsedWorkbook {
        workbook: Workbook { sheets },
        styles,
        shared_strings,
    })
}

/// Read only the first `max_bytes` of a ZIP entry as text. Used to probe the
/// top of a worksheet (its `<sheetPr>`) without inflating the whole sheet.
/// Lossy UTF-8 keeps a multibyte character split at the cut from erroring; the
/// region we care about (`<sheetPr><tabColor>`) is pure ASCII near the start.
fn read_zip_entry_head(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    name: &str,
    max_bytes: u64,
) -> Result<String, String> {
    let mut file = archive
        .by_name(name)
        .map_err(|e| format!("entry '{}' not found: {}", name, e))?;
    let mut buf = Vec::new();
    file.by_ref()
        .take(max_bytes)
        .read_to_end(&mut buf)
        .map_err(|e| e.to_string())?;
    Ok(String::from_utf8_lossy(&buf).into_owned())
}

/// Extract the resolved tab color from the head of a worksheet XML. Locates the
/// single `<tabColor .../>` element (it lives in `<sheetPr>`, before
/// `<sheetData>`) and resolves its `rgb` / `theme`+`tint` / `indexed` attributes
/// through the same rules as cell colors. Returns `None` when no tab color is
/// declared or the tag is truncated by the head limit. A lightweight attribute
/// scan avoids any namespace-prefix assumptions in the partial document.
fn extract_tab_color_from_head(head: &str, theme_colors: &[String]) -> Option<String> {
    // Don't look past the data body — `tabColor` only appears in `<sheetPr>`.
    let scope = head.split("<sheetData").next().unwrap_or(head);
    let start = scope.find("tabColor")?;
    let rest = &scope[start..];
    // The element is self-closing (`<tabColor ... />`); read up to its `>`.
    let end = rest.find('>')?;
    let tag = &rest[..end];
    let attr = |name: &str| -> Option<&str> {
        let key = format!("{}=\"", name);
        let i = tag.find(&key)? + key.len();
        let j = tag[i..].find('"')? + i;
        Some(&tag[i..j])
    };
    resolve_color_attrs(
        attr("rgb"),
        attr("theme"),
        attr("tint"),
        attr("indexed"),
        theme_colors,
    )
}

/// Theme `fmtScheme > lnStyleLst` line widths (EMU), in declaration order.
/// A drawing shape's `<a:style><a:lnRef idx="N">` resolves its outline width
/// from entry N (1-based) of this list (ECMA-376 §20.1.4.2.19); an entry
/// without an explicit `w` uses the CT_LineProperties default 9525 EMU =
/// 0.75 pt (§20.1.2.2.24).
pub(crate) fn parse_theme_ln_widths(archive: &mut zip::ZipArchive<Cursor<&[u8]>>) -> Vec<i64> {
    let Ok(xml) = read_zip_string(archive, "xl/theme/theme1.xml") else {
        return Vec::new();
    };
    let Ok(doc) = roxmltree::Document::parse(&xml) else {
        return Vec::new();
    };
    let a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main";
    let mut widths: Vec<i64> = Vec::new();
    for node in doc.descendants() {
        if node.tag_name().name() == "lnStyleLst" && node.tag_name().namespace() == Some(a_ns) {
            for ln in node
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "ln")
            {
                widths.push(
                    ln.attribute("w")
                        .and_then(|s| s.parse().ok())
                        .unwrap_or(9525),
                );
            }
            break;
        }
    }
    widths
}

fn parse_theme_colors(archive: &mut zip::ZipArchive<Cursor<&[u8]>>) -> Vec<String> {
    let Ok(xml) = read_zip_string(archive, "xl/theme/theme1.xml") else {
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
                if !child.is_element() {
                    continue;
                }
                // Each child is a color slot; its first child element holds the actual color
                for color_node in child.children() {
                    if !color_node.is_element() {
                        continue;
                    }
                    let hex = match color_node.tag_name().name() {
                        "srgbClr" => color_node
                            .attribute("val")
                            .map(|v| format!("#{}", v.to_uppercase())),
                        "sysClr" => color_node
                            .attribute("lastClr")
                            .map(|v| format!("#{}", v.to_uppercase())),
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
    if hex.len() < 6 {
        return format!("#{}", hex);
    }
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
    format!(
        "#{:02X}{:02X}{:02X}",
        (nr * 255.0).round() as u8,
        (ng * 255.0).round() as u8,
        (nb * 255.0).round() as u8
    )
}

fn hls_to_rgb(h: f64, l: f64, s: f64) -> (f64, f64, f64) {
    if s == 0.0 {
        return (l, l, l);
    }
    let q = if l < 0.5 {
        l * (1.0 + s)
    } else {
        l + s - l * s
    };
    let p = 2.0 * l - q;
    let r = hue_to_rgb(p, q, h + 1.0 / 3.0);
    let g = hue_to_rgb(p, q, h);
    let b = hue_to_rgb(p, q, h - 1.0 / 3.0);
    (r, g, b)
}

fn hue_to_rgb(p: f64, q: f64, mut t: f64) -> f64 {
    if t < 0.0 {
        t += 1.0;
    }
    if t > 1.0 {
        t -= 1.0;
    }
    if t < 1.0 / 6.0 {
        return p + (q - p) * 6.0 * t;
    }
    if t < 1.0 / 2.0 {
        return q;
    }
    if t < 2.0 / 3.0 {
        return p + (q - p) * (2.0 / 3.0 - t) * 6.0;
    }
    p
}

pub(crate) fn parse_color(node: &roxmltree::Node, theme_colors: &[String]) -> Option<String> {
    resolve_color_attrs(
        node.attribute("rgb"),
        node.attribute("theme"),
        node.attribute("tint"),
        node.attribute("indexed"),
        theme_colors,
    )
}

/// Resolve a DrawingML/SpreadsheetML color from its raw attribute values
/// (`rgb` / `theme` + `tint` / `indexed`). Split out from [`parse_color`] so
/// callers that scan attributes without a roxmltree node (e.g. the bounded
/// tab-color head probe) share the exact same resolution rules.
pub(crate) fn resolve_color_attrs(
    rgb: Option<&str>,
    theme: Option<&str>,
    tint: Option<&str>,
    indexed: Option<&str>,
    theme_colors: &[String],
) -> Option<String> {
    // rgb attribute (ARGB: 8 chars, drop alpha; or 6-char RGB)
    if let Some(rgb) = rgb {
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
    // This is an index→index remap, not a logical→slot-name mapping, so the
    // shared ooxml_common::color::SCHEME_DEFAULT_SLOTS table (the canonical
    // §19.3.1.6 logical→slot names) does not apply here; this stays local.
    if let Some(theme_str) = theme {
        if let Ok(idx) = theme_str.parse::<usize>() {
            let mapped = match idx {
                0 => 1,
                1 => 0,
                2 => 3,
                3 => 2,
                n => n,
            };
            if let Some(base) = theme_colors.get(mapped) {
                let tint = tint.and_then(|s| s.parse::<f64>().ok()).unwrap_or(0.0);
                if tint == 0.0 {
                    return Some(base.clone());
                }
                return Some(apply_tint(base, tint));
            }
        }
    }

    // indexed attribute → Excel built-in palette
    if let Some(indexed_str) = indexed {
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
            let r_id = node.attribute((r_ns, "id")).unwrap_or("").to_string();
            let visibility = match node.attribute("state") {
                Some("hidden") => SheetVisibility::Hidden,
                Some("veryHidden") => SheetVisibility::VeryHidden,
                _ => SheetVisibility::Visible,
            };
            sheets.push(SheetMeta {
                name,
                sheet_id,
                r_id,
                tab_color: None,
                visibility,
            });
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
        if let Some(l) = local {
            if l != sheet_index {
                continue;
            }
        }
        let name = match node.attribute("name") {
            Some(n) => n.to_string(),
            None => continue,
        };
        let formula = node.text().unwrap_or("").to_string();
        names.push(DefinedName { name, formula });
    }
    names
}

fn resolve_sheet_path(doc: &roxmltree::Document, r_id: &str) -> Option<String> {
    let ns = "http://schemas.openxmlformats.org/package/2006/relationships";
    for node in doc.descendants() {
        if node.tag_name().name() == "Relationship"
            && node.tag_name().namespace() == Some(ns)
            && node.attribute("Id") == Some(r_id)
        {
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
    None
}

fn read_shared_strings(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    theme_colors: &[String],
) -> Vec<SharedString> {
    let Ok(xml) = read_zip_string(archive, "xl/sharedStrings.xml") else {
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
fn parse_si_node(node: &roxmltree::Node, ns: &str, theme_colors: &[String]) -> SharedString {
    let mut text = String::new();
    let mut runs: Vec<Run> = Vec::new();
    let mut has_runs = false;
    for child in node.children() {
        if !child.is_element() {
            continue;
        }
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
                runs.push(Run {
                    text: run_text,
                    font: run_font,
                });
            }
            _ => {}
        }
    }
    SharedString {
        text,
        runs: if has_runs { Some(runs) } else { None },
    }
}
/// `(row, col, relationship id)` triples for cell hyperlinks pending rels resolution.
type HyperlinkRids = Vec<(u32, u32, String)>;

fn parse_worksheet(
    xml: &str,
    shared_strings: &[SharedString],
    theme_colors: &[String],
    name: &str,
) -> Result<(Worksheet, HyperlinkRids), String> {
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
    let mut right_to_left = false;
    let mut tab_color: Option<String> = None;
    let mut auto_filter: Option<CellRange> = None;
    let mut hyperlink_rids: HyperlinkRids = Vec::new();

    // Pre-scan worksheet-level extLst for x14:dataBar extension attributes.
    // Excel 2010+ stores the `gradient` flag on `<x14:dataBar>` inside
    // `<extLst>/<ext>/<x14:conditionalFormattings>/<x14:conditionalFormatting>
    // /<x14:cfRule id="{GUID}">`, linked to the SpreadsheetML cfRule via a
    // matching `<x14:id>{GUID}</x14:id>` inside the cfRule's own extLst
    // (§2.6.3). Build a GUID → gradient map so cfRule parsing can look up
    // the override.
    let mut x14_databar_gradient: HashMap<String, bool> = HashMap::new();
    for x14_rule in doc
        .descendants()
        .filter(|n| n.tag_name().name() == "cfRule" && n.attribute("type") == Some("dataBar"))
    {
        let Some(id) = x14_rule.attribute("id") else {
            continue;
        };
        for bar in x14_rule
            .children()
            .filter(|n| n.tag_name().name() == "dataBar")
        {
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
    for x14_cf in doc.descendants().filter(|n| {
        n.tag_name().name() == "conditionalFormatting"
            && n.tag_name()
                .namespace()
                .map(|u| u.contains("/spreadsheetml/2009/9"))
                .unwrap_or(false)
    }) {
        let sqref: Vec<CellRange> = x14_cf
            .children()
            .find(|n| n.tag_name().name() == "sqref")
            .and_then(|n| n.text())
            .map(parse_sqref)
            .unwrap_or_default();
        if sqref.is_empty() {
            continue;
        }
        let mut rules: Vec<CfRule> = Vec::new();
        for x14_rule in x14_cf
            .children()
            .filter(|n| n.tag_name().name() == "cfRule" && n.attribute("type") == Some("iconSet"))
        {
            let priority: i32 = x14_rule
                .attribute("priority")
                .and_then(|s| s.parse().ok())
                .unwrap_or(0);
            let Some(icon_node) = x14_rule
                .children()
                .find(|n| n.tag_name().name() == "iconSet")
            else {
                continue;
            };
            let custom = icon_node
                .attribute("custom")
                .map(|v| v == "1" || v == "true")
                .unwrap_or(false);
            let icon_set_name = icon_node
                .attribute("iconSet")
                .unwrap_or(if custom { "" } else { "3TrafficLights1" })
                .to_string();
            let reverse = icon_node
                .attribute("reverse")
                .map(|v| v == "1" || v == "true")
                .unwrap_or(false);
            let mut cfvos: Vec<CfValue> = Vec::new();
            let mut custom_icons: Vec<CfIcon> = Vec::new();
            for ch in icon_node.children().filter(|n| n.is_element()) {
                match ch.tag_name().name() {
                    "cfvo" => {
                        let kind = ch.attribute("type").unwrap_or("percent").to_string();
                        // x14:cfvo stores the value in `<xm:f>` child; attribute val fallback.
                        let value = ch
                            .children()
                            .find(|n| n.tag_name().name() == "f")
                            .and_then(|n| n.text())
                            .map(|s| s.to_string())
                            .or_else(|| ch.attribute("val").map(|s| s.to_string()));
                        cfvos.push(CfValue { kind, value });
                    }
                    "cfIcon" => {
                        let set = ch.attribute("iconSet").unwrap_or("NoIcons").to_string();
                        let id = ch
                            .attribute("iconId")
                            .and_then(|s| s.parse().ok())
                            .unwrap_or(0);
                        custom_icons.push(CfIcon {
                            icon_set: set,
                            icon_id: id,
                        });
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
                if let Some(v) = node
                    .attribute("defaultColWidth")
                    .and_then(|s| s.parse().ok())
                {
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
                let custom = attr_bool(&node, "customWidth").unwrap_or(false);
                let hidden = attr_bool(&node, "hidden").unwrap_or(false);
                // Only record widths for custom-widthed columns OR hidden columns
                if !custom && !hidden {
                    continue;
                }
                let min: u32 = node
                    .attribute("min")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(1);
                let max: u32 = node
                    .attribute("max")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(1);
                // Cap range to avoid storing 16K entries for max=16384 ranges
                let max = max.min(min + 255);
                let width: f64 = if hidden {
                    0.0
                } else {
                    node.attribute("width")
                        .and_then(|s| s.parse().ok())
                        .unwrap_or(default_col_width)
                };
                for c in min..=max {
                    col_widths.insert(c, width);
                }
            }
            "sheetView" if node.tag_name().namespace() == Some(ns) => {
                show_zeros = attr_bool(&node, "showZeros").unwrap_or(true);
                show_gridlines = attr_bool(&node, "showGridLines").unwrap_or(true);
                // ECMA-376 §18.3.1.87 `rightToLeft` — mirrors the whole grid so
                // column A is on the right. Default false (left-to-right).
                right_to_left = attr_bool(&node, "rightToLeft").unwrap_or(false);
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
                        Some(CellRange {
                            top,
                            left,
                            bottom,
                            right,
                        })
                    } else {
                        let (col, row) = parse_cell_ref(parts[0]);
                        Some(CellRange {
                            top: row,
                            left: col,
                            bottom: row,
                            right: col,
                        })
                    };
                }
            }
            "hyperlinks" if node.tag_name().namespace() == Some(ns) => {
                for hl in node.children() {
                    if !hl.is_element() || hl.tag_name().name() != "hyperlink" {
                        continue;
                    }
                    let Some(ref_str) = hl.attribute("ref") else {
                        continue;
                    };
                    // Only first cell of ref range
                    let ref_single = ref_str.split(':').next().unwrap_or(ref_str);
                    let (col, row) = parse_cell_ref(ref_single);
                    if let Some(rid) = hl
                        .attributes()
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
                    freeze_rows = node
                        .attribute("ySplit")
                        .and_then(|s| s.parse::<f64>().ok())
                        .map(|v| v as u32)
                        .unwrap_or(0);
                    freeze_cols = node
                        .attribute("xSplit")
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
                        merge_cells.push(MergeCell {
                            top,
                            left,
                            bottom,
                            right,
                        });
                    }
                }
            }
            "row" if node.tag_name().namespace() == Some(ns) => {
                let row_idx: u32 = node
                    .attribute("r")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0);
                let hidden = attr_bool(&node, "hidden").unwrap_or(false);
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
                rows.push(Row {
                    index: row_idx,
                    height,
                    cells,
                });
            }
            "conditionalFormatting" if node.tag_name().namespace() == Some(ns) => {
                let sqref = node.attribute("sqref").map(parse_sqref).unwrap_or_default();
                let mut rules: Vec<CfRule> = Vec::new();
                for cf in node.children() {
                    if cf.tag_name().name() != "cfRule" {
                        continue;
                    }
                    let kind = cf.attribute("type").unwrap_or("").to_string();
                    let priority: i32 = cf
                        .attribute("priority")
                        .and_then(|s| s.parse().ok())
                        .unwrap_or(0);
                    let dxf_id: Option<u32> = cf.attribute("dxfId").and_then(|s| s.parse().ok());
                    match kind.as_str() {
                        "cellIs" => {
                            let operator = cf.attribute("operator").unwrap_or("equal").to_string();
                            let formulas: Vec<String> = cf
                                .children()
                                .filter(|n| n.tag_name().name() == "formula")
                                .filter_map(|n| n.text().map(|s| s.to_string()))
                                .collect();
                            rules.push(CfRule::CellIs {
                                operator,
                                formulas,
                                dxf_id,
                                priority,
                            });
                        }
                        "expression" | "containsBlanks" | "notContainsBlanks" | "containsText"
                        | "notContainsText" | "beginsWith" | "endsWith" | "containsErrors"
                        | "notContainsErrors" => {
                            // For `containsBlanks`/`notContainsBlanks`/`containsText` etc.,
                            // Excel serializes an equivalent boolean formula (e.g.
                            // `LEN(TRIM(C8))>0`) as the rule's `<formula>` child
                            // (ECMA-376 §18.3.1.10). Evaluate as an expression rule.
                            let formula = cf
                                .children()
                                .find(|n| n.tag_name().name() == "formula")
                                .and_then(|n| n.text())
                                .unwrap_or("")
                                .to_string();
                            let stop_if_true = cf
                                .attribute("stopIfTrue")
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false);
                            rules.push(CfRule::Expression {
                                formula,
                                dxf_id,
                                priority,
                                stop_if_true,
                            });
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
                                                child
                                                    .attribute("type")
                                                    .unwrap_or("num")
                                                    .to_string(),
                                                child.attribute("val").map(|s| s.to_string()),
                                            ));
                                        }
                                        "color" => {
                                            stop_colors.push(
                                                parse_color(&child, theme_colors)
                                                    .unwrap_or_else(|| "#FFFFFF".to_string()),
                                            );
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            let stops: Vec<CfStop> = stop_values
                                .into_iter()
                                .enumerate()
                                .map(|(i, (kind, value))| CfStop {
                                    kind,
                                    value,
                                    color: stop_colors
                                        .get(i)
                                        .cloned()
                                        .unwrap_or_else(|| "#FFFFFF".to_string()),
                                })
                                .collect();
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
                                                child
                                                    .attribute("type")
                                                    .unwrap_or("min")
                                                    .to_string(),
                                                child.attribute("val").map(|s| s.to_string()),
                                            ));
                                        }
                                        "color" => {
                                            if let Some(c) = parse_color(&child, theme_colors) {
                                                color = c;
                                            }
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
                            'gradient_lookup: for ext_list in
                                cf.children().filter(|n| n.tag_name().name() == "extLst")
                            {
                                for ext in
                                    ext_list.children().filter(|n| n.tag_name().name() == "ext")
                                {
                                    for id_node in
                                        ext.descendants().filter(|n| n.tag_name().name() == "id")
                                    {
                                        if let Some(guid) = id_node.text() {
                                            if let Some(&g) = x14_databar_gradient.get(guid) {
                                                gradient = g;
                                                break 'gradient_lookup;
                                            }
                                        }
                                    }
                                    // Fallback: some files embed <x14:dataBar>
                                    // directly in the cfRule's extLst.
                                    for x14_bar in ext
                                        .descendants()
                                        .filter(|n| n.tag_name().name() == "dataBar")
                                    {
                                        if let Some(g) = x14_bar.attribute("gradient") {
                                            gradient = !(g == "0" || g == "false");
                                            break 'gradient_lookup;
                                        }
                                    }
                                }
                            }
                            let min = cfvos
                                .first()
                                .map(|(k, v)| CfValue {
                                    kind: k.clone(),
                                    value: v.clone(),
                                })
                                .unwrap_or(CfValue {
                                    kind: "min".into(),
                                    value: None,
                                });
                            let max = cfvos
                                .get(1)
                                .map(|(k, v)| CfValue {
                                    kind: k.clone(),
                                    value: v.clone(),
                                })
                                .unwrap_or(CfValue {
                                    kind: "max".into(),
                                    value: None,
                                });
                            rules.push(CfRule::DataBar {
                                color,
                                min,
                                max,
                                priority,
                                gradient,
                            });
                        }
                        "top10" => {
                            let top = !cf
                                .attribute("bottom")
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false);
                            let percent = cf
                                .attribute("percent")
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false);
                            let rank = cf
                                .attribute("rank")
                                .and_then(|s| s.parse().ok())
                                .unwrap_or(10);
                            rules.push(CfRule::Top10 {
                                top,
                                percent,
                                rank,
                                dxf_id,
                                priority,
                            });
                        }
                        "aboveAverage" => {
                            let above_average = cf
                                .attribute("aboveAverage")
                                .map(|v| v != "0")
                                .unwrap_or(true);
                            // ECMA-376 §18.3.1.10: `equalAverage` (default
                            // false) and `stdDev` (optional, number of
                            // standard deviations for the band threshold).
                            let equal_average = cf
                                .attribute("equalAverage")
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false);
                            let std_dev = cf
                                .attribute("stdDev")
                                .and_then(|v| v.parse::<u32>().ok())
                                .filter(|&n| n > 0);
                            rules.push(CfRule::AboveAverage {
                                above_average,
                                equal_average,
                                std_dev,
                                dxf_id,
                                priority,
                            });
                        }
                        "iconSet" => {
                            let icon_set_node =
                                cf.children().find(|n| n.tag_name().name() == "iconSet");
                            let icon_set = icon_set_node
                                .and_then(|n| n.attribute("iconSet"))
                                .unwrap_or("3TrafficLights1")
                                .to_string();
                            let reverse = icon_set_node
                                .and_then(|n| n.attribute("reverse"))
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false);
                            let cfvos: Vec<CfValue> = icon_set_node
                                .map(|n| {
                                    n.children()
                                        .filter(|c| c.is_element() && c.tag_name().name() == "cfvo")
                                        .map(|c| CfValue {
                                            kind: c
                                                .attribute("type")
                                                .unwrap_or("percent")
                                                .to_string(),
                                            value: c.attribute("val").map(|s| s.to_string()),
                                        })
                                        .collect()
                                })
                                .unwrap_or_default();
                            rules.push(CfRule::IconSet {
                                icon_set,
                                cfvos,
                                reverse,
                                priority,
                                custom_icons: None,
                            });
                        }
                        other => {
                            rules.push(CfRule::Other {
                                kind: other.to_string(),
                                priority,
                            });
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

    Ok((
        Worksheet {
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
            right_to_left,
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
        },
        hyperlink_rids,
    ))
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
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else {
        return Vec::new();
    };
    let sheet_rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let Ok(rels_xml) = read_zip_string(archive, &sheet_rels_path) else {
        return Vec::new();
    };
    let Ok(rels_doc) = roxmltree::Document::parse(&rels_xml) else {
        return Vec::new();
    };

    // A sheet may carry classic notes (`/comments`, ECMA-376 §18.7) and/or
    // Office-365 threaded comments (`/threadedComment`, MS extension). Excel
    // writes a back-compat `xl/commentsN.xml` alongside every threaded comment,
    // so the classic file is preferred when present (it already includes the
    // threaded text). Only when there is no classic file do we fall back to the
    // threaded part — which is the case for files authored by tools that emit
    // threaded comments exclusively.
    let mut classic_target: Option<String> = None;
    let mut threaded_target: Option<String> = None;
    for rel in rels_doc
        .root_element()
        .children()
        .filter(|n| n.is_element())
    {
        let rel_type = rel.attribute("Type").unwrap_or("");
        let Some(t) = rel.attribute("Target") else {
            continue;
        };
        if rel_type.ends_with("/comments") {
            classic_target.get_or_insert_with(|| t.to_string());
        } else if rel_type.ends_with("/threadedComment") {
            threaded_target.get_or_insert_with(|| t.to_string());
        }
    }

    if let Some(target) = classic_target {
        let comments_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        if let Ok(comments_xml) = read_zip_string(archive, &comments_path) {
            return parse_comments_xml(&comments_xml);
        }
    }

    if let Some(target) = threaded_target {
        let tc_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        if let Ok(tc_xml) = read_zip_string(archive, &tc_path) {
            let persons = load_persons(archive);
            return parse_threaded_comments_xml(&tc_xml, &persons);
        }
    }

    Vec::new()
}

/// Load `personId` → display-name map from `xl/persons/person*.xml`
/// (Office-365 threaded-comment authors, MS-XLSX schema
/// `…/office/spreadsheetml/2018/threadedcomments`). `<person displayName id/>`.
/// Returns an empty map when no persons part exists.
fn load_persons(archive: &mut zip::ZipArchive<Cursor<&[u8]>>) -> HashMap<String, String> {
    let mut out: HashMap<String, String> = HashMap::new();
    // Persons live under xl/persons/ by convention; collect every part there so
    // we don't depend on the exact file name (person.xml vs person1.xml).
    let person_paths: Vec<String> = (0..archive.len())
        .filter_map(|i| archive.by_index(i).ok().map(|f| f.name().to_string()))
        .filter(|n| n.starts_with("xl/persons/") && n.ends_with(".xml"))
        .collect();
    for path in person_paths {
        let Ok(xml) = read_zip_string(archive, &path) else {
            continue;
        };
        let Ok(doc) = roxmltree::Document::parse(&xml) else {
            continue;
        };
        for p in doc
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "person")
        {
            if let (Some(id), Some(name)) = (p.attribute("id"), p.attribute("displayName")) {
                out.insert(id.to_string(), name.to_string());
            }
        }
    }
    out
}

/// Parse `xl/threadedComments/threadedCommentN.xml` (MS-XLSX threaded comments,
/// schema `…/office/spreadsheetml/2018/threadedcomments`). Each
/// `<threadedComment ref personId>` carries a `<text>`; `personId` resolves to a
/// display name via the `persons` map. Multiple comments on the same cell (a
/// thread of replies) are joined into one body, mirroring how the classic
/// back-compat file flattens a thread. Returns one `XlsxComment` per cell that
/// has at least one threaded comment.
fn parse_threaded_comments_xml(
    tc_xml: &str,
    persons: &HashMap<String, String>,
) -> Vec<XlsxComment> {
    let Ok(doc) = roxmltree::Document::parse(tc_xml) else {
        return Vec::new();
    };
    // Preserve document order of first appearance per cell ref.
    let mut order: Vec<String> = Vec::new();
    let mut by_ref: HashMap<String, (Option<String>, String)> = HashMap::new();
    for node in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "threadedComment")
    {
        let Some(cell_ref) = node.attribute("ref") else {
            continue;
        };
        let author = node
            .attribute("personId")
            .and_then(|id| persons.get(id).cloned())
            .filter(|s| !s.is_empty());
        let text = node
            .children()
            .find(|c| c.is_element() && c.tag_name().name() == "text")
            .and_then(|t| t.text())
            .unwrap_or("")
            .to_string();
        let entry = by_ref.entry(cell_ref.to_string()).or_insert_with(|| {
            order.push(cell_ref.to_string());
            (None, String::new())
        });
        // First comment in the thread sets the author; replies are appended.
        if entry.0.is_none() {
            entry.0 = author;
        }
        if !entry.1.is_empty() {
            entry.1.push('\n');
        }
        entry.1.push_str(&text);
    }
    order
        .into_iter()
        .map(|cell_ref| {
            let (author, text) = by_ref.remove(&cell_ref).unwrap();
            XlsxComment {
                cell_ref,
                author,
                text,
            }
        })
        .collect()
}

/// Parse a `xl/commentsN.xml` document (ECMA-376 §18.7) into structured
/// `XlsxComment`s. Resolves `@authorId` against the `<authors>` block and joins
/// every `<text>/<r>/<t>` run into plain text (rich-text formatting dropped).
/// Returns an empty vec on malformed XML. Split out from `load_sheet_comments`
/// so the parse path is unit-testable without a ZIP archive.
fn parse_comments_xml(comments_xml: &str) -> Vec<XlsxComment> {
    let Ok(comments_doc) = roxmltree::Document::parse(comments_xml) else {
        return Vec::new();
    };

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
        if node.tag_name().name() != "comment" || !node.is_element() {
            continue;
        }
        let Some(cell_ref) = node.attribute("ref") else {
            continue;
        };
        let author = node
            .attribute("authorId")
            .and_then(|s| s.parse::<usize>().ok())
            .and_then(|i| authors.get(i).cloned())
            .filter(|s| !s.is_empty());
        let mut text = String::new();
        if let Some(t_node) = node
            .children()
            .find(|c| c.is_element() && c.tag_name().name() == "text")
        {
            for r in t_node.descendants() {
                if r.is_element() && r.tag_name().name() == "t" {
                    if let Some(s) = r.text() {
                        text.push_str(s);
                    }
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
    let Some(dvs) = ws_root
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "dataValidations")
    else {
        return out;
    };
    for dv in dvs
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "dataValidation")
    {
        let sqref = dv.attribute("sqref").unwrap_or("").to_string();
        if sqref.is_empty() {
            continue;
        }
        let validation_type = dv.attribute("type").map(String::from);
        let operator = dv.attribute("operator").map(String::from);
        let allow_blank = dv
            .attribute("allowBlank")
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
            .unwrap_or(false);
        let prompt_title = dv
            .attribute("promptTitle")
            .map(String::from)
            .filter(|s| !s.is_empty());
        let prompt = dv
            .attribute("prompt")
            .map(String::from)
            .filter(|s| !s.is_empty());
        let error_title = dv
            .attribute("errorTitle")
            .map(String::from)
            .filter(|s| !s.is_empty());
        let error_message = dv
            .attribute("error")
            .map(String::from)
            .filter(|s| !s.is_empty());

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

/// Resolve hyperlink rIds to URLs from the sheet rels file.
fn load_hyperlinks(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    sheet_path: &str,
    hyperlink_rids: HyperlinkRids,
) -> Vec<Hyperlink> {
    if hyperlink_rids.is_empty() {
        return Vec::new();
    }
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else {
        return Vec::new();
    };
    let rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let rels = read_zip_string(archive, &rels_path)
        .ok()
        .map(|xml| parse_rels_map(&xml))
        .unwrap_or_default();
    hyperlink_rids
        .into_iter()
        .map(|(col, row, rid)| Hyperlink {
            col,
            row,
            url: rels.get(&rid).cloned(),
        })
        .collect()
}

/// Resolve a relative path ("../media/image1.png") against a base dir ("xl/drawings").
pub(crate) fn resolve_zip_path(base_dir: &str, target: &str) -> String {
    // An absolute Target (leading "/", e.g. openpyxl's
    // `/xl/drawings/drawing1.xml`) is package-root-relative and must ignore
    // `base_dir`; otherwise the base would be prepended, producing a path that
    // doesn't exist in the archive (ECMA-376 / OPC part names are root-anchored
    // when they start with "/").
    let mut parts: Vec<&str> = if target.starts_with('/') {
        Vec::new()
    } else {
        base_dir.split('/').filter(|s| !s.is_empty()).collect()
    };
    for seg in target.split('/') {
        match seg {
            ".." => {
                parts.pop();
            }
            "." | "" => {}
            s => parts.push(s),
        }
    }
    parts.join("/")
}

pub(crate) fn resolve_fill_color(
    fill_node: &roxmltree::Node,
    theme_colors: &[String],
) -> Option<String> {
    // Accept either a `<a:solidFill>` directly or a `<c:spPr>` whose first
    // fill-ish child is `<a:solidFill>`. Looking at *direct* children (not
    // descendants) is intentional — chart series often carry label/axis text
    // colors under `c:dLbls`/`c:txPr` which must NOT be misread as fill.
    let solid = if fill_node.tag_name().name() == "solidFill" {
        Some(*fill_node)
    } else {
        fill_node
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
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
                // ooxml_common::color::SCHEME_DEFAULT_SLOTS is the canonical
                // logical→slot-NAME table (§19.3.1.6). xlsx instead maps the
                // logical/slot name straight to a numeric INDEX into the theme
                // color Vec (raw clrScheme order: dk1=0, lt1=1, dk2=2, lt2=3,
                // accent1..6=4..9, hlink=10, folHlink=11). Routing through the
                // shared name→name table would add an indirection without
                // changing the result, so this stays local. (Note: the cell
                // @theme path below applies the §22.1.2.7 dk1↔lt1 / dk2↔lt2
                // index swap; this drawing path indexes the array directly.)
                let idx = match v {
                    "dk1" | "tx1" => Some(0),
                    "lt1" | "bg1" => Some(1),
                    "dk2" | "tx2" => Some(2),
                    "lt2" | "bg2" => Some(3),
                    "accent1" => Some(4),
                    "accent2" => Some(5),
                    "accent3" => Some(6),
                    "accent4" => Some(7),
                    "accent5" => Some(8),
                    "accent6" => Some(9),
                    "hlink" => Some(10),
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
    s.split_whitespace()
        .map(|range_str| {
            if let Some((a, b)) = range_str.split_once(':') {
                let (left, top) = parse_cell_ref(a);
                let (right, bottom) = parse_cell_ref(b);
                CellRange {
                    top,
                    left,
                    bottom,
                    right,
                }
            } else {
                let (col, row) = parse_cell_ref(range_str);
                CellRange {
                    top: row,
                    left: col,
                    bottom: row,
                    right: col,
                }
            }
        })
        .collect()
}

/// Split an `<xm:f>` reference like `Sheet1!A1:A10` or `'My Sheet'!$B$3:$B$8`
/// into `(sheet_name, range)`. Returns `None` if the reference has no sheet
/// qualifier — sparkline data refs always do, so unqualified is treated as
/// "same sheet" by callers.
fn split_sheet_ref(s: &str) -> (Option<String>, String) {
    let s = s.trim();
    let Some(bang) = s.rfind('!') else {
        return (None, s.to_string());
    };
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
    let Ok(doc) = roxmltree::Document::parse(sheet_xml) else {
        return values;
    };
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let row_span = (range.right - range.left + 1) as usize;
    for c in doc
        .descendants()
        .filter(|n| n.tag_name().name() == "c" && n.tag_name().namespace() == Some(ns))
    {
        let Some(r_attr) = c.attribute("r") else {
            continue;
        };
        let (col, row) = parse_cell_ref(r_attr);
        if row < range.top || row > range.bottom || col < range.left || col > range.right {
            continue;
        }
        // Only honor numeric / formula-numeric cells. `t` of "s" / "str" /
        // "inlineStr" / "b" / "e" all map to None for sparkline values.
        let t = c.attribute("t").unwrap_or("");
        if matches!(t, "s" | "str" | "inlineStr" | "b" | "e") {
            continue;
        }
        let v = c
            .children()
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
    let Ok(doc) = roxmltree::Document::parse(sheet_xml) else {
        return Vec::new();
    };
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

    for group_node in doc
        .descendants()
        .filter(|n| n.tag_name().name() == "sparklineGroup")
    {
        let kind = match group_node.attribute("type").unwrap_or("line") {
            "column" => SparklineType::Column,
            "stacked" => SparklineType::Stem, // historical alias
            "stem" => SparklineType::Stem,
            // ECMA-376 lists `line` and a planned `stairStep`; treat unknown
            // types as line (closest visual fallback).
            _ => SparklineType::Line,
        };

        let resolve_color = |child_name: &str| -> Option<String> {
            group_node
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == child_name)
                .and_then(|n| parse_color(&n, theme_colors))
        };

        let mut sparklines: Vec<Sparkline> = Vec::new();
        // <x14:sparklines> is the wrapper; <x14:sparkline> are the children.
        for sparklines_node in group_node
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "sparklines")
        {
            for sl in sparklines_node
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "sparkline")
            {
                let f_text = sl
                    .children()
                    .find(|n| n.is_element() && n.tag_name().name() == "f")
                    .and_then(|n| n.text())
                    .unwrap_or("");
                let sqref_text = sl
                    .children()
                    .find(|n| n.is_element() && n.tag_name().name() == "sqref")
                    .and_then(|n| n.text())
                    .unwrap_or("");
                if f_text.is_empty() || sqref_text.is_empty() {
                    continue;
                }
                let (col, row) = parse_cell_ref(sqref_text.trim());
                let (source_sheet, range_str) = split_sheet_ref(f_text);
                let ranges = parse_sqref(&range_str);
                let Some(range) = ranges.into_iter().next() else {
                    continue;
                };

                // Look up source sheet XML (cross-sheet ref). When the ref
                // has no sheet qualifier, fall back to the *current* sheet
                // XML.
                let source_xml: Option<&str> = match source_sheet {
                    Some(name) => {
                        if !xml_cache.contains_key(&name) {
                            let path = sheets
                                .iter()
                                .find(|s| s.name == name)
                                .and_then(|s| resolve_sheet_path(rels_doc, &s.r_id))
                                .map(|p| format!("xl/{}", p));
                            let xml = path.and_then(|p| read_zip_string(archive, &p).ok());
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
            display_empty_cells_as: group_node
                .attribute("displayEmptyCellsAs")
                .unwrap_or("gap")
                .to_string(),
            min_axis_type: group_node
                .attribute("minAxisType")
                .unwrap_or("individual")
                .to_string(),
            max_axis_type: group_node
                .attribute("maxAxisType")
                .unwrap_or("individual")
                .to_string(),
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
        let style_index: u32 = c_node
            .attribute("s")
            .and_then(|s| s.parse().ok())
            .unwrap_or(0);

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
                    CellValue::Text {
                        text: ss.text,
                        runs: ss.runs,
                    }
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
                        CellValue::Text {
                            text: ss.text.clone(),
                            runs: ss.runs.clone(),
                        }
                    } else {
                        CellValue::Text {
                            text: String::new(),
                            runs: None,
                        }
                    }
                }
                "str" => CellValue::Text {
                    text: v_text,
                    runs: None,
                },
                "b" => CellValue::Bool {
                    bool: v_text == "1" || v_text == "true",
                },
                "e" => CellValue::Error { error: v_text },
                _ => {
                    if let Ok(n) = v_text.parse::<f64>() {
                        CellValue::Number { number: n }
                    } else {
                        CellValue::Text {
                            text: v_text,
                            runs: None,
                        }
                    }
                }
            }
        };

        cells.push(Cell {
            col,
            row,
            col_ref: cell_ref,
            value,
            style_index,
            formula,
        });
    }
    cells
}

/// Parse an `ST_Boolean` (ECMA-376 §22.9.2.7, xsd:boolean) attribute value.
/// Accepts `1`/`true`/`on` as true and `0`/`false`/`off` as false (case-insensitive).
/// Returns `None` when the attribute is absent so callers can apply their own default.
pub(crate) fn attr_bool(node: &roxmltree::Node, name: &str) -> Option<bool> {
    node.attribute(name)
        .map(|v| matches!(v.trim().to_ascii_lowercase().as_str(), "1" | "true" | "on"))
}

pub(crate) fn parse_cell_ref(r: &str) -> (u32, u32) {
    let col_str: String = r.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
    let row_str: String = r.chars().skip_while(|c| c.is_ascii_alphabetic()).collect();
    let col = col_str
        .chars()
        .fold(0u32, |acc, c| acc * 26 + (c as u32 - 'A' as u32 + 1));
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

/// Extract raw bytes for a single embedded image entry (e.g.
/// "xl/media/image1.png") from an xlsx zip archive. Thin `wasm_bindgen` wrapper
/// over the shared [`ooxml_common::zip::extract_zip_entry`] reader; used by the
/// main thread to lazily materialize image blobs on demand.
#[wasm_bindgen]
pub fn extract_image(
    data: &[u8],
    path: &str,
    max_zip_entry_bytes: Option<u64>,
) -> Result<Vec<u8>, JsValue> {
    ooxml_common::zip::extract_zip_entry(data, path, max_zip_entry_bytes)
        .map_err(|e| JsValue::from_str(&e))
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
    let workbook_xml = read_zip_string(&mut archive, "xl/workbook.xml")?;
    let wb_doc = roxmltree::Document::parse(&workbook_xml).map_err(|e| e.to_string())?;
    let sheets = parse_workbook_sheets(&wb_doc);

    let mut out = String::new();
    for (idx, sheet_meta) in sheets.iter().enumerate() {
        let sheet_json = parse_sheet_native(data, idx as u32, &sheet_meta.name)
            .map_err(|e| format!("sheet '{}' (#{}) parse failed: {}", sheet_meta.name, idx, e))?;
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

    let workbook_xml = read_zip_string(&mut archive, "xl/workbook.xml")?;
    let wb_doc = roxmltree::Document::parse(&workbook_xml).map_err(|e| e.to_string())?;
    let sheets = parse_workbook_sheets(&wb_doc);

    let sheet_meta = sheets
        .get(sheet_index as usize)
        .ok_or_else(|| format!("sheet index {} out of range", sheet_index))?;

    let rels_xml = read_zip_string(&mut archive, "xl/_rels/workbook.xml.rels")?;
    let rels_doc = roxmltree::Document::parse(&rels_xml).map_err(|e| e.to_string())?;
    let sheet_path = resolve_sheet_path(&rels_doc, &sheet_meta.r_id)
        .ok_or_else(|| format!("rId {} not found in rels", sheet_meta.r_id))?;

    let theme_colors = parse_theme_colors(&mut archive);
    let shared_strings = read_shared_strings(&mut archive, &theme_colors);
    let sheet_xml = read_zip_string(&mut archive, &format!("xl/{}", sheet_path))?;
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
    ws.sparkline_groups =
        load_sheet_sparklines(&mut archive, &sheet_xml, &sheets, &rels_doc, &theme_colors);
    let (df_family, df_size) = parse_default_font(&mut archive);
    ws.default_font_family = df_family;
    ws.default_font_size = df_size;

    serde_json::to_string(&ws).map_err(|e| e.to_string())
}

#[cfg(test)]
mod tab_color_tests {
    use super::extract_tab_color_from_head;

    const THEME: &[&str] = &[
        "#000000", "#FFFFFF", "#44546A", "#E7E6E6", "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000",
        "#5B9BD5", "#70AD47", "#0563C1", "#954F72",
    ];

    fn theme() -> Vec<String> {
        THEME.iter().map(|s| s.to_string()).collect()
    }

    #[test]
    fn tab_color_rgb() {
        let head = r#"<?xml version="1.0"?><worksheet><sheetPr><tabColor rgb="FFFF0000"/></sheetPr><dimension ref="A1"/><sheetData>"#;
        assert_eq!(
            extract_tab_color_from_head(head, &theme()).as_deref(),
            Some("#FF0000")
        );
    }

    #[test]
    fn tab_color_theme_with_tint() {
        // theme="4" (Excel-internal accent1) resolves to #4472C4; a tint just
        // needs to produce *something* different — exact value covered by apply_tint.
        let head = r#"<worksheet><sheetPr codeName="S1"><tabColor theme="4" tint="-0.249977111117893"/></sheetPr><sheetData/></worksheet>"#;
        let got = extract_tab_color_from_head(head, &theme());
        assert!(got.is_some(), "theme tab color should resolve");
        assert_ne!(
            got.as_deref(),
            Some("#4472C4"),
            "tint should darken the base"
        );
    }

    #[test]
    fn tab_color_absent() {
        let head = r#"<worksheet><sheetPr/><dimension ref="A1"/><sheetData><row/></sheetData></worksheet>"#;
        assert_eq!(extract_tab_color_from_head(head, &theme()), None);
    }

    #[test]
    fn tab_color_not_searched_past_sheetdata() {
        // A stray "tabColor" token inside the body must not be misread.
        let head =
            r#"<worksheet><sheetPr/><sheetData><c><is><t>tabColor rgb="00FF00"</t></is></c>"#;
        assert_eq!(extract_tab_color_from_head(head, &theme()), None);
    }
}

#[cfg(test)]
mod sheet_view_tests {
    use super::parse_worksheet;

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// ECMA-376 §18.3.1.87 `<sheetView rightToLeft="1">` mirrors the entire
    /// grid (column A on the right).
    #[test]
    fn sheet_view_right_to_left() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetViews><sheetView rightToLeft="1" workbookViewId="0"/></sheetViews><sheetData/></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");
        assert!(ws.right_to_left, "rightToLeft=\"1\" → right_to_left true");
    }

    /// Absent `@rightToLeft` defaults to false (left-to-right).
    #[test]
    fn sheet_view_right_to_left_defaults_false() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData/></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");
        assert!(
            !ws.right_to_left,
            "absent @rightToLeft → right_to_left false"
        );
    }

    /// ECMA-376 §22.9.2.7 `ST_Boolean` allows `true`/`false` as well as `1`/`0`.
    /// LibreOffice writes `<col customWidth="true" .../>`; the parser must honor
    /// the recorded width instead of skipping the `<col>` (which would leave the
    /// column at `defaultColWidth`).
    #[test]
    fn col_custom_width_accepts_true_literal() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><cols><col customWidth="true" min="1" max="1" width="22"/></cols><sheetData/></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");
        assert_eq!(
            ws.col_widths.get(&1).copied(),
            Some(22.0),
            "customWidth=\"true\" → width 22 recorded for column 1"
        );
    }

    /// `customWidth="1"` (Excel's spelling) must keep working after the helper change.
    #[test]
    fn col_custom_width_accepts_one_literal() {
        let xml = format!(
            r#"<worksheet xmlns="{NS}"><cols><col customWidth="1" min="2" max="2" width="10"/></cols><sheetData/></worksheet>"#
        );
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");
        assert_eq!(ws.col_widths.get(&2).copied(), Some(10.0));
    }
}

#[cfg(test)]
mod resolve_zip_path_tests {
    use super::resolve_zip_path;

    /// A relative Target resolves against the base directory, honoring `..`.
    #[test]
    fn relative_target_resolves_against_base() {
        assert_eq!(
            resolve_zip_path("xl/worksheets", "../drawings/drawing1.xml"),
            "xl/drawings/drawing1.xml"
        );
        assert_eq!(
            resolve_zip_path("xl/drawings", "../media/image1.png"),
            "xl/media/image1.png"
        );
    }

    /// An absolute Target (leading "/", as openpyxl writes for drawings) is
    /// package-root-relative and ignores the base directory.
    #[test]
    fn absolute_target_ignores_base() {
        assert_eq!(
            resolve_zip_path("xl/worksheets", "/xl/drawings/drawing1.xml"),
            "xl/drawings/drawing1.xml"
        );
        assert_eq!(
            resolve_zip_path("xl/drawings", "/xl/charts/chart1.xml"),
            "xl/charts/chart1.xml"
        );
    }
}

#[cfg(test)]
mod conditional_format_tests {
    use super::parse_worksheet;
    use crate::types::CfRule;

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn parse_cf_rules(cf_xml: &str) -> Vec<CfRule> {
        let xml = format!(r#"<worksheet xmlns="{NS}"><sheetData/>{cf_xml}</worksheet>"#);
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");
        ws.conditional_formats
            .into_iter()
            .flat_map(|cf| cf.rules)
            .collect()
    }

    /// ECMA-376 §18.3.1.10: an `aboveAverage` rule with no extra attributes
    /// defaults to `aboveAverage=true`, `equalAverage=false`, no `stdDev`.
    #[test]
    fn above_average_defaults() {
        let rules = parse_cf_rules(
            r#"<conditionalFormatting sqref="A1:A5"><cfRule type="aboveAverage" dxfId="0" priority="1"/></conditionalFormatting>"#,
        );
        match &rules[..] {
            [CfRule::AboveAverage {
                above_average,
                equal_average,
                std_dev,
                ..
            }] => {
                assert!(*above_average, "aboveAverage defaults to true");
                assert!(!*equal_average, "equalAverage defaults to false");
                assert_eq!(*std_dev, None, "no stdDev by default");
            }
            other => panic!("expected one AboveAverage rule, got {other:?}"),
        }
    }

    /// `aboveAverage="0"` flips to below-average; `equalAverage="1"` is honored.
    #[test]
    fn below_average_with_equal_average() {
        let rules = parse_cf_rules(
            r#"<conditionalFormatting sqref="A1:A5"><cfRule type="aboveAverage" aboveAverage="0" equalAverage="1" dxfId="0" priority="1"/></conditionalFormatting>"#,
        );
        match &rules[..] {
            [CfRule::AboveAverage {
                above_average,
                equal_average,
                ..
            }] => {
                assert!(!*above_average, "aboveAverage=\"0\" → false");
                assert!(*equal_average, "equalAverage=\"1\" → true");
            }
            other => panic!("expected one AboveAverage rule, got {other:?}"),
        }
    }

    /// `stdDev="2"` is captured as a band multiplier (ECMA-376 §18.3.1.10).
    #[test]
    fn above_average_std_dev() {
        let rules = parse_cf_rules(
            r#"<conditionalFormatting sqref="A1:A5"><cfRule type="aboveAverage" stdDev="2" dxfId="0" priority="1"/></conditionalFormatting>"#,
        );
        match &rules[..] {
            [CfRule::AboveAverage { std_dev, .. }] => {
                assert_eq!(*std_dev, Some(2), "stdDev=\"2\" captured");
            }
            other => panic!("expected one AboveAverage rule, got {other:?}"),
        }
    }
}

#[cfg(test)]
mod data_validation_tests {
    use super::parse_worksheet;
    use crate::types::DataValidation;

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn parse_dvs(dv_xml: &str) -> Vec<DataValidation> {
        let xml = format!(r#"<worksheet xmlns="{NS}"><sheetData/>{dv_xml}</worksheet>"#);
        let (ws, _) = parse_worksheet(&xml, &[], &[], "Sheet1").expect("worksheet parses");
        ws.data_validations
    }

    /// ECMA-376 §18.3.1.33 — a `list` rule captures type, sqref and the
    /// `<formula1>` literal list. `allowBlank="1"` is honored.
    #[test]
    fn list_validation_captures_formula_and_sqref() {
        let dvs = parse_dvs(
            r#"<dataValidations count="1"><dataValidation sqref="B2:B5" type="list" allowBlank="1"><formula1>"Pending,Shipped,Delivered"</formula1></dataValidation></dataValidations>"#,
        );
        assert_eq!(dvs.len(), 1, "one rule parsed");
        let dv = &dvs[0];
        assert_eq!(dv.sqref, "B2:B5");
        assert_eq!(dv.validation_type.as_deref(), Some("list"));
        assert_eq!(
            dv.formula1.as_deref(),
            Some("\"Pending,Shipped,Delivered\"")
        );
        assert!(dv.allow_blank, "allowBlank=\"1\" → true");
    }

    /// A `whole`/`between` rule keeps both operands and the operator.
    #[test]
    fn whole_between_keeps_both_operands() {
        let dvs = parse_dvs(
            r#"<dataValidations count="1"><dataValidation sqref="C2:C5" type="whole" operator="between"><formula1>1</formula1><formula2>100</formula2></dataValidation></dataValidations>"#,
        );
        let dv = &dvs[0];
        assert_eq!(dv.validation_type.as_deref(), Some("whole"));
        assert_eq!(dv.operator.as_deref(), Some("between"));
        assert_eq!(dv.formula1.as_deref(), Some("1"));
        assert_eq!(dv.formula2.as_deref(), Some("100"));
        assert!(!dv.allow_blank, "absent allowBlank → false");
    }

    /// A rule without a `@sqref` is dropped (nothing to anchor it to).
    #[test]
    fn rule_without_sqref_is_skipped() {
        let dvs = parse_dvs(
            r#"<dataValidations count="1"><dataValidation type="list"><formula1>"A,B"</formula1></dataValidation></dataValidations>"#,
        );
        assert!(dvs.is_empty(), "missing sqref → rule dropped");
    }

    /// Multiple rules in one block are all captured, preserving order.
    #[test]
    fn multiple_rules_preserved() {
        let dvs = parse_dvs(
            r#"<dataValidations count="2"><dataValidation sqref="B2:B5" type="list"><formula1>"A,B"</formula1></dataValidation><dataValidation sqref="C2:C5" type="whole" operator="between"><formula1>1</formula1><formula2>9</formula2></dataValidation></dataValidations>"#,
        );
        assert_eq!(dvs.len(), 2);
        assert_eq!(dvs[0].sqref, "B2:B5");
        assert_eq!(dvs[1].sqref, "C2:C5");
    }

    /// Absent `<dataValidations>` yields an empty vec (no panic).
    #[test]
    fn absent_block_yields_empty() {
        let dvs = parse_dvs("");
        assert!(dvs.is_empty());
    }
}

#[cfg(test)]
mod comment_tests {
    use super::parse_comments_xml;

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// ECMA-376 §18.7 — each `<comment>` yields its cell ref, the author
    /// resolved from `@authorId`, and the joined `<t>` text.
    #[test]
    fn resolves_ref_author_and_text() {
        let xml = format!(
            r#"<comments xmlns="{NS}"><authors><author>Reviewer</author><author>Ops Team</author></authors><commentList><comment ref="B1" authorId="0"><text><t>Set the order status.</t></text></comment><comment ref="C3" authorId="1"><text><t>Verify qty.</t></text></comment></commentList></comments>"#
        );
        let cs = parse_comments_xml(&xml);
        assert_eq!(cs.len(), 2);
        assert_eq!(cs[0].cell_ref, "B1");
        assert_eq!(cs[0].author.as_deref(), Some("Reviewer"));
        assert_eq!(cs[0].text, "Set the order status.");
        assert_eq!(cs[1].cell_ref, "C3");
        assert_eq!(cs[1].author.as_deref(), Some("Ops Team"));
    }

    /// Multiple `<r><t>` runs in one comment are concatenated into plain text.
    #[test]
    fn joins_multiple_runs() {
        let xml = format!(
            r#"<comments xmlns="{NS}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="0"><text><r><t>Hello </t></r><r><t>world</t></r></text></comment></commentList></comments>"#
        );
        let cs = parse_comments_xml(&xml);
        assert_eq!(cs[0].text, "Hello world");
    }

    /// An out-of-range or absent `@authorId` leaves the author as None.
    #[test]
    fn missing_author_is_none() {
        let xml = format!(
            r#"<comments xmlns="{NS}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="9"><text><t>orphan</t></text></comment></commentList></comments>"#
        );
        let cs = parse_comments_xml(&xml);
        assert_eq!(cs.len(), 1);
        assert_eq!(cs[0].author, None, "authorId out of range → None");
        assert_eq!(cs[0].text, "orphan");
    }

    /// Malformed XML returns an empty vec instead of panicking.
    #[test]
    fn malformed_xml_yields_empty() {
        assert!(parse_comments_xml("<comments><not closed").is_empty());
    }
}

#[cfg(test)]
mod threaded_comment_tests {
    use super::parse_threaded_comments_xml;
    use std::collections::HashMap;

    const TC_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";

    fn persons() -> HashMap<String, String> {
        let mut m = HashMap::new();
        m.insert("{p1}".to_string(), "Reviewer".to_string());
        m.insert("{p2}".to_string(), "Ops Team".to_string());
        m
    }

    /// MS-XLSX threaded comments — each `<threadedComment>` yields its cell ref,
    /// the author resolved from `personId`, and the `<text>` body.
    #[test]
    fn resolves_ref_person_and_text() {
        let xml = format!(
            r#"<ThreadedComments xmlns="{TC_NS}"><threadedComment ref="B1" personId="{{p1}}" id="a"><text>Set the status.</text></threadedComment><threadedComment ref="C3" personId="{{p2}}" id="b"><text>Verify qty.</text></threadedComment></ThreadedComments>"#
        );
        let cs = parse_threaded_comments_xml(&xml, &persons());
        assert_eq!(cs.len(), 2);
        assert_eq!(cs[0].cell_ref, "B1");
        assert_eq!(cs[0].author.as_deref(), Some("Reviewer"));
        assert_eq!(cs[0].text, "Set the status.");
        assert_eq!(cs[1].author.as_deref(), Some("Ops Team"));
    }

    /// A thread of replies on one cell is flattened into a single comment body,
    /// keeping the original author.
    #[test]
    fn replies_collapse_into_one_thread() {
        let xml = format!(
            r#"<ThreadedComments xmlns="{TC_NS}"><threadedComment ref="A1" personId="{{p1}}" id="a"><text>Question?</text></threadedComment><threadedComment ref="A1" personId="{{p2}}" id="b" parentId="a"><text>Answer.</text></threadedComment></ThreadedComments>"#
        );
        let cs = parse_threaded_comments_xml(&xml, &persons());
        assert_eq!(cs.len(), 1, "one comment per cell");
        assert_eq!(cs[0].cell_ref, "A1");
        assert_eq!(
            cs[0].author.as_deref(),
            Some("Reviewer"),
            "first author kept"
        );
        assert_eq!(cs[0].text, "Question?\nAnswer.");
    }

    /// An unknown `personId` leaves the author as None (no persons part).
    #[test]
    fn unknown_person_is_none() {
        let xml = format!(
            r#"<ThreadedComments xmlns="{TC_NS}"><threadedComment ref="A1" personId="{{zzz}}" id="a"><text>hi</text></threadedComment></ThreadedComments>"#
        );
        let cs = parse_threaded_comments_xml(&xml, &HashMap::new());
        assert_eq!(cs[0].author, None);
        assert_eq!(cs[0].text, "hi");
    }
}

#[cfg(test)]
mod extract_image_tests {
    use super::extract_image;

    #[test]
    fn extract_image_reads_entry() {
        use std::io::{Cursor, Write};
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = zip::write::SimpleFileOptions::default();
            w.start_file("xl/media/i.png", o).unwrap();
            w.write_all(b"X").unwrap();
            w.finish().unwrap();
        }
        assert_eq!(extract_image(&buf, "xl/media/i.png", None).unwrap(), b"X");
    }
}

#[cfg(test)]
mod sheet_visibility_tests {
    use super::*;

    #[test]
    fn sheet_state_attr_maps_to_visibility() {
        // ECMA-376 §18.2.19 ST_SheetState: visible (default) | hidden | veryHidden.
        let xml = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId2" state="hidden"/><sheet name="C" sheetId="3" r:id="rId3" state="veryHidden"/><sheet name="D" sheetId="4" r:id="rId4" state="visible"/></sheets></workbook>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let sheets = parse_workbook_sheets(&doc);
        assert_eq!(sheets[0].visibility, SheetVisibility::Visible); // absent ⇒ visible
        assert_eq!(sheets[1].visibility, SheetVisibility::Hidden);
        assert_eq!(sheets[2].visibility, SheetVisibility::VeryHidden);
        assert_eq!(sheets[3].visibility, SheetVisibility::Visible);
    }
}
