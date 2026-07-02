use ooxml_common::blip::{mime_from_ext, parse_src_rect, svg_blip_rid, SrcRect};
use ooxml_common::math::parse_omath_nodes;
use ooxml_common::text::{parse_lnspc, SpaceLine};
use std::collections::HashMap;
use std::io::{Cursor, Read};
use wasm_bindgen::prelude::*;

mod table_style_presets;

mod types;
pub(crate) use types::*;

mod markdown;
use markdown::render_presentation_md;

mod chart;
use chart::*;

mod theme;
use theme::*;

// Test-only counter for `roxmltree::Document::parse` calls on the D4 hot paths
// (slide master build, layout, slide XML + decorations). It exists ONLY under
// `cfg(test)` — `note_layout_master_parse()` compiles to nothing in release, so
// this is zero-cost for shipped builds. A regression test uses it to assert that
// a deck whose slides share one layout + one master parses each of those parts a
// bounded number of times (see `parse_count_scales_with_distinct_parts`),
// guarding against re-introducing the per-slide re-parses this change removed.
#[cfg(test)]
thread_local! {
    static LAYOUT_MASTER_PARSE_COUNT: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

/// Increment the D4 parse counter (no-op unless `cfg(test)`). Call immediately
/// before parsing a slide-master / layout / slide XML on the pagination path.
#[inline(always)]
fn note_layout_master_parse() {
    #[cfg(test)]
    LAYOUT_MASTER_PARSE_COUNT.with(|c| c.set(c.get() + 1));
}

// ===========================
//  Public WASM entry points
// ===========================

/// Parse a pptx archive and return the model as UTF-8 JSON **bytes**.
///
/// Returning `Vec<u8>` (a fresh copy on the JS side) instead of `String` keeps
/// the model out of the JsString/UTF-16 representation: the worker forwards the
/// resulting `ArrayBuffer` to the main thread as a transferable and the main
/// thread does a single `TextDecoder.decode` + `JSON.parse`, collapsing three
/// serializations (Rust String → JsString → structured clone) into one decode.
#[wasm_bindgen]
pub fn parse_pptx(data: &[u8], max_zip_entry_bytes: Option<u64>) -> Result<Vec<u8>, JsValue> {
    console_error_panic_hook::set_once();
    let _guard = ooxml_common::zip::scoped_max(max_zip_entry_bytes);
    let presentation = parse_presentation_from_bytes(data)
        .map_err(|e| JsValue::from_str(&format!("pptx-parser error: {e}")))?;
    serde_json::to_vec(&presentation)
        .map_err(|e| JsValue::from_str(&format!("serialize error: {e}")))
}

/// WASM-callable markdown projection. Shares the body of `to_markdown_native`
/// so the browser / Node WASM path and the native mcp-server path stay in
/// lock-step. See `to_markdown_native` for the design rationale.
#[wasm_bindgen]
pub fn pptx_to_markdown(data: &[u8], max_zip_entry_bytes: Option<u64>) -> Result<String, JsValue> {
    console_error_panic_hook::set_once();
    let _guard = ooxml_common::zip::scoped_max(max_zip_entry_bytes);
    let pres = parse_presentation_from_bytes(data)
        .map_err(|e| JsValue::from_str(&format!("pptx-parser error: {e}")))?;
    Ok(render_presentation_md(&pres))
}

/// Native equivalent of `parse_pptx` for use from the MCP server.
pub fn parse_pptx_native(data: &[u8]) -> Result<String, String> {
    let presentation = parse_presentation_from_bytes(data).map_err(|e| e.to_string())?;
    serde_json::to_string(&presentation).map_err(|e| e.to_string())
}

/// Parse a pptx and project the result to GitHub-flavoured markdown,
/// preserving textual / semantic structure (headings, bullets, tables, charts,
/// notes, comments) and discarding presentation details (geometry, fills,
/// strokes, effects, theme inheritance details). Designed for AI agents that
/// need to read content efficiently — typical 10-30× token reduction vs. the
/// raw JSON of `parse_pptx_native`.
pub fn to_markdown_native(data: &[u8]) -> Result<String, String> {
    let pres = parse_presentation_from_bytes(data).map_err(|e| e.to_string())?;
    Ok(render_presentation_md(&pres))
}

// ───────────────────────────────────────────────────────────────────────────
//  Markdown projection (text-focused) — separate code path from the rich JSON
//  serialization used by the viewer. Lossy by design.
// ───────────────────────────────────────────────────────────────────────────

/// Extract raw bytes for a single entry (e.g. "ppt/media/media2.mp4") from a
/// pptx zip archive. Used by the main thread to materialize media blobs for
/// interactive playback without re-parsing the whole file.
#[wasm_bindgen]
pub fn extract_media(
    data: &[u8],
    path: &str,
    max_zip_entry_bytes: Option<u64>,
) -> Result<Vec<u8>, JsValue> {
    ooxml_common::zip::extract_zip_entry(data, path, max_zip_entry_bytes)
        .map_err(|e| JsValue::from_str(&e))
}

/// Extract raw bytes for a single embedded image entry (e.g.
/// "ppt/media/image1.png") from a pptx zip archive. Thin `wasm_bindgen` wrapper
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

/// A stateful handle over an opened pptx archive.
///
/// The free functions above (`parse_pptx` / `pptx_to_markdown` / `extract_media`
/// / `extract_image`) each re-copy the whole file into WASM and re-scan the ZIP
/// central directory on every call. A `PptxArchive` copies the bytes into WASM
/// **once** (in `new`) and keeps the opened [`PptxZip`] alive, so a `parse`
/// followed by any number of `extract_media` / `extract_image` calls (the
/// viewer's parse-then-lazily-load-media pattern) pays the copy + open cost a
/// single time. `ZipArchive<Cursor<Vec<u8>>>` is self-contained (it owns its
/// bytes and holds no borrow into the input), which is what lets it live in a
/// `#[wasm_bindgen]` struct field.
///
/// The retained `max` mirrors the per-call `scoped_max` guard the free functions
/// install: every method re-installs it for its own scope so the zip-bomb entry
/// cap is honored identically whether callers use the handle or the free
/// functions.
#[wasm_bindgen]
pub struct PptxArchive {
    archive: PptxZip,
    max: Option<u64>,
}

#[wasm_bindgen]
impl PptxArchive {
    /// Copy `data` into WASM once and open the ZIP central directory once.
    /// `max_zip_entry_bytes` is retained and applied on every subsequent method
    /// call (identical semantics to the free functions' `scoped_max` guard).
    #[wasm_bindgen(constructor)]
    pub fn new(data: &[u8], max_zip_entry_bytes: Option<u64>) -> Result<PptxArchive, JsValue> {
        console_error_panic_hook::set_once();
        let archive = zip::ZipArchive::new(Cursor::new(data.to_vec()))
            .map_err(|e| JsValue::from_str(&format!("pptx-parser error: {e}")))?;
        Ok(PptxArchive {
            archive,
            max: max_zip_entry_bytes,
        })
    }

    /// Parse the retained archive and return the model as UTF-8 JSON bytes.
    /// Byte-for-byte identical to `parse_pptx` on the same file.
    pub fn parse(&mut self) -> Result<Vec<u8>, JsValue> {
        let _guard = ooxml_common::zip::scoped_max(self.max);
        let presentation = parse_presentation(&mut self.archive)
            .map_err(|e| JsValue::from_str(&format!("pptx-parser error: {e}")))?;
        serde_json::to_vec(&presentation)
            .map_err(|e| JsValue::from_str(&format!("serialize error: {e}")))
    }

    /// Extract raw bytes for one media entry (e.g. "ppt/media/media2.mp4") from
    /// the retained archive. Twin of the free `extract_media`, but reads through
    /// the already-open archive instead of re-opening it.
    pub fn extract_media(&mut self, path: &str) -> Result<Vec<u8>, JsValue> {
        let _guard = ooxml_common::zip::scoped_max(self.max);
        ooxml_common::zip::read_zip_bytes(&mut self.archive, path)
            .map_err(|e| JsValue::from_str(&e))
    }

    /// Extract raw bytes for one embedded image entry (e.g.
    /// "ppt/media/image1.png") from the retained archive. Twin of the free
    /// `extract_image`.
    pub fn extract_image(&mut self, path: &str) -> Result<Vec<u8>, JsValue> {
        let _guard = ooxml_common::zip::scoped_max(self.max);
        ooxml_common::zip::read_zip_bytes(&mut self.archive, path)
            .map_err(|e| JsValue::from_str(&e))
    }

    /// GitHub-flavoured markdown projection of the retained archive. Mirrors the
    /// free `pptx_to_markdown`.
    pub fn to_markdown(&mut self) -> Result<String, JsValue> {
        let _guard = ooxml_common::zip::scoped_max(self.max);
        let pres = parse_presentation(&mut self.archive)
            .map_err(|e| JsValue::from_str(&format!("pptx-parser error: {e}")))?;
        Ok(render_presentation_md(&pres))
    }
}

// ===========================
//  Data types  (camelCase JSON → TypeScript)
// ===========================

// ===== Table data model =====

/// Parse `<a:blip><a:alphaModFix amt="..."/></a:blip>` from a blipFill node.
/// Returns fraction (amt / 100000) when present and < 1.0; None otherwise.
fn parse_blip_alpha(blip_fill: roxmltree::Node<'_, '_>) -> Option<f64> {
    let blip = child(blip_fill, "blip")?;
    let amf = child(blip, "alphaModFix")?;
    let amt: f64 = attr(&amf, "amt")?.parse().ok()?;
    let frac = amt / 100_000.0;
    if frac >= 0.9999 {
        None
    } else {
        Some(frac.max(0.0))
    }
}

/// Parse `<a:stretch><a:fillRect l t r b>` (ECMA-376 §20.1.8.58 / §20.1.8.30).
/// Edge attributes are ST_Percentage (1000ths of a percent → /100000 gives a
/// fraction). Negative values are valid (overscan). Returns None when there is
/// no fillRect or all four edges are zero (= the source fills the whole box).
fn parse_fill_rect(stretch: roxmltree::Node<'_, '_>) -> Option<FillRect> {
    let fr = child(stretch, "fillRect")?;
    let read = |name: &str| -> f64 {
        attr(&fr, name)
            .and_then(|v| v.parse::<f64>().ok())
            .map(|v| v / 100_000.0)
            .unwrap_or(0.0)
    };
    let rect = FillRect {
        l: read("l"),
        t: read("t"),
        r: read("r"),
        b: read("b"),
    };
    if is_zero_f64(&rect.l) && is_zero_f64(&rect.t) && is_zero_f64(&rect.r) && is_zero_f64(&rect.b)
    {
        None
    } else {
        Some(rect)
    }
}

/// Parse `<a:tile tx ty sx sy flip algn>` (ECMA-376 §20.1.8.58 CT_TileInfoProperties).
///
/// - `tx` / `ty`: ST_Coordinate offset of the first tile, in EMU. Default 0.
/// - `sx` / `sy`: ST_Percentage scale of the tile (1000ths of a percent →
///   `/100000` gives a fraction). Absent → `1.0` (100% = native blip size).
/// - `flip`: ST_TileFlipMode (none|x|y|xy). Default "none".
/// - `algn`: ST_RectAlignment (tl|t|tr|l|ctr|r|bl|b|br) — the anchor corner the
///   tile grid is registered against. The schema gives no default; PowerPoint
///   treats an absent `algn` as "tl" (the first tile sits at the box origin
///   plus tx/ty). We carry it verbatim and default to "tl".
fn parse_tile(tile: roxmltree::Node<'_, '_>) -> TileInfo {
    let scale = |name: &str| -> f64 {
        attr(&tile, name)
            .and_then(|v| v.parse::<f64>().ok())
            .map(|v| v / 100_000.0)
            .unwrap_or(1.0)
    };
    TileInfo {
        tx: attr_i64(&tile, "tx").unwrap_or(0),
        ty: attr_i64(&tile, "ty").unwrap_or(0),
        sx: scale("sx"),
        sy: scale("sy"),
        flip: attr(&tile, "flip").unwrap_or_else(|| "none".to_owned()),
        algn: attr(&tile, "algn").unwrap_or_else(|| "tl".to_owned()),
    }
}

// ===========================
//  ZIP helpers
// ===========================

/// The parser's ZIP archive type. Owns its backing bytes (`Cursor<Vec<u8>>`)
/// rather than borrowing them, so a `PptxArchive` handle can keep a single
/// opened archive alive across `parse` / `extract_media` / `extract_image` /
/// `to_markdown` calls — the central directory is scanned once and the bytes are
/// copied into WASM once. `ZipArchive<Cursor<Vec<u8>>>` is fully self-contained
/// (no borrow into the input), which is what lets a `#[wasm_bindgen]` handle
/// store it as a field.
type PptxZip = zip::ZipArchive<Cursor<Vec<u8>>>;

fn read_zip_str(zip: &mut PptxZip, path: &str) -> Result<String, Box<dyn std::error::Error>> {
    let max = ooxml_common::zip::current_max();
    let mut file = zip
        .by_name(path)
        .map_err(|_| format!("missing ZIP entry: {path}"))?;
    if file.size() > max {
        return Err(format!("ZIP entry exceeds size limit: {path}").into());
    }
    let mut buf = String::new();
    file.by_ref().take(max).read_to_string(&mut buf)?;
    Ok(buf)
}

// ===========================
//  Table style data model
// ===========================

/// Resolved fills and borders extracted from a single <a:tblStyle> definition.
#[derive(Debug, Clone, Default)]
struct TableStyleDef {
    whole_fill: Option<Fill>,
    whole_inside_h: Option<Stroke>,
    whole_inside_v: Option<Stroke>,
    /// Outer top/bottom edge border (from wholeTbl tcBdr top/bottom)
    whole_outer_h: Option<Stroke>,
    /// Outer left/right edge border (from wholeTbl tcBdr left/right)
    whole_outer_v: Option<Stroke>,
    band1h_fill: Option<Fill>,
    band2h_fill: Option<Fill>,
    first_row_fill: Option<Fill>,
    first_row_border_b: Option<Stroke>,
    last_row_fill: Option<Fill>,
    first_col_fill: Option<Fill>,
    last_col_fill: Option<Fill>,
    /// Default text colour per role, from `<a:tcTxStyle>` (schemeClr/srgbClr).
    /// e.g. wholeTbl → dk1, firstRow header → lt1 (white). Hex, no `#`.
    whole_text_color: Option<String>,
    first_row_text_color: Option<String>,
    last_row_text_color: Option<String>,
    first_col_text_color: Option<String>,
    last_col_text_color: Option<String>,
    /// Default bold per role, from `<a:tcTxStyle b="on">` (ECMA-376 §20.1.4.2.28).
    /// e.g. a firstRow header is commonly bold.
    first_row_bold: Option<bool>,
    last_row_bold: Option<bool>,
    first_col_bold: Option<bool>,
    last_col_bold: Option<bool>,
}

// ===========================
//  XML helpers (roxmltree)
// ===========================

pub(crate) fn child<'a, 'i>(
    node: roxmltree::Node<'a, 'i>,
    local: &str,
) -> Option<roxmltree::Node<'a, 'i>> {
    node.children()
        .find(|n| n.is_element() && n.tag_name().name() == local)
}

fn children_vec<'a, 'i>(
    node: roxmltree::Node<'a, 'i>,
    local: &str,
) -> Vec<roxmltree::Node<'a, 'i>> {
    node.children()
        .filter(|n| n.is_element() && n.tag_name().name() == local)
        .collect()
}

const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

pub(crate) fn attr(node: &roxmltree::Node<'_, '_>, local: &str) -> Option<String> {
    node.attributes()
        .find(|a| a.name() == local && a.namespace().is_none())
        .map(|a| a.value().to_owned())
}

/// Attribute in the r: (relationships) namespace — e.g. r:id, r:embed
fn attr_r(node: &roxmltree::Node<'_, '_>, local: &str) -> Option<String> {
    node.attributes()
        .find(|a| a.name() == local && a.namespace() == Some(R_NS))
        .map(|a| a.value().to_owned())
}

fn attr_i64(node: &roxmltree::Node<'_, '_>, local: &str) -> Option<i64> {
    attr(node, local)?.parse().ok()
}

fn attr_f64(node: &roxmltree::Node<'_, '_>, local: &str) -> Option<f64> {
    attr(node, local)?.parse().ok()
}

// ===========================
//  Relationships helpers
// ===========================

/// id → target  (used for image/slide lookups by rId). Thin adapter over
/// [`ooxml_common::rels::parse_rels`] that flattens each `RelTarget` back to its
/// raw target string (both Internal part names and External URLs are kept
/// verbatim — resolution to a zip part happens later via [`resolve_path`]),
/// preserving this parser's long-standing `HashMap<rId, Target>` shape.
fn parse_rels(xml: &str) -> HashMap<String, String> {
    ooxml_common::rels::parse_rels(xml)
        .into_iter()
        .map(|(id, rel)| (id, rel.target))
        .collect()
}

/// Pair diagramData rels with diagramDrawing rels in a slide's rels XML by filename
/// number (data1.xml ↔ drawing1.xml, data2.xml ↔ drawing2.xml, ...), then load each
/// drawing XML from the zip. Returns dm_rid → drawing_xml_content.
fn build_smartart_drawings(rels_xml: &str, zip: &mut PptxZip) -> HashMap<String, String> {
    let mut result: HashMap<String, String> = HashMap::new();
    let doc = match roxmltree::Document::parse(rels_xml) {
        Ok(d) => d,
        Err(_) => return result,
    };
    let mut data_rels: Vec<(String, String)> = Vec::new();
    let mut drawing_targets: Vec<String> = Vec::new();
    for rel in doc.root_element().children().filter(|n| n.is_element()) {
        let rel_type = attr(&rel, "Type").unwrap_or_default();
        let (Some(rid), Some(target)) = (attr(&rel, "Id"), attr(&rel, "Target")) else {
            continue;
        };
        if rel_type.ends_with("/diagramData") {
            data_rels.push((rid, target));
        } else if rel_type.ends_with("/diagramDrawing") {
            drawing_targets.push(target);
        }
    }
    fn trailing_num(target: &str) -> Option<u32> {
        let file = target.rsplit('/').next().unwrap_or("");
        let stem = file.split('.').next().unwrap_or("");
        let digits: String = stem
            .chars()
            .rev()
            .take_while(|c| c.is_ascii_digit())
            .collect();
        digits.chars().rev().collect::<String>().parse().ok()
    }
    for (rid, data_target) in data_rels {
        let Some(num) = trailing_num(&data_target) else {
            continue;
        };
        let drawing_target = drawing_targets
            .iter()
            .find(|t| trailing_num(t) == Some(num));
        if let Some(dt) = drawing_target {
            let drawing_path = resolve_path("ppt/slides", dt);
            if let Ok(xml) = read_zip_str(zip, &drawing_path) {
                result.insert(rid, xml);
            }
        }
    }
    result
}

/// Find the Target of the first relationship whose Type ends with `type_suffix`.
fn find_rel_target_by_type(rels_xml: &str, type_suffix: &str) -> Option<String> {
    let doc = roxmltree::Document::parse(rels_xml).ok()?;
    for rel in doc.root_element().children().filter(|n| n.is_element()) {
        if let Some(rel_type) = attr(&rel, "Type") {
            if rel_type.ends_with(type_suffix) {
                return attr(&rel, "Target");
            }
        }
    }
    None
}

/// Resolve a relative path against a base directory inside the ZIP.
///
/// Thin alias for the shared [`ooxml_common::rels::resolve_target`], which
/// handles both root-absolute (`/ppt/charts/chart5.xml`) and relative
/// (`../charts/chart1.xml`) Targets with `..` normalization (ECMA-376 Part 2
/// §9.3). Kept as a local name so the many call sites read unchanged.
fn resolve_path(base_dir: &str, target: &str) -> String {
    ooxml_common::rels::resolve_target(base_dir, target)
}

// ===========================
//  Color parsing
// ===========================

/// Resolve a color node (solidFill child / run rPr child) to a hex string.
/// Handles srgbClr, sysClr, prstClr, and schemeClr (with transform support).
pub(crate) fn parse_color_node(
    node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<String> {
    parse_color_node_tint(node, theme, ooxml_common::color::TintMode::PowerPointLinear)
}

/// Like `parse_color_node`, but lets the caller pick how `<a:tint>` is interpreted.
/// Table styles (`<a:tcStyle>` band fills) use `TintMode::WordLiteral` — the literal
/// ECMA-376 §20.1.2.3.34 definition (`val·input + (1-val)·white`, so a 20% tint is a
/// near-white wash) — which is how PowerPoint renders table band tints. The SmartArt
/// accent-recolor path keeps `PowerPointLinear` (see `apply_color_transforms`).
///
/// Thin wrapper over the shared [`ooxml_common::color::parse_color_node`]: the
/// grammar + transforms live there; [`PptxSchemeResolver`] supplies the
/// pptx-specific theme-slot lookup. Output is unchanged (uppercase hex, no `#`).
fn parse_color_node_tint(
    node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    tint_mode: ooxml_common::color::TintMode,
) -> Option<String> {
    ooxml_common::color::parse_color_node(node, &PptxSchemeResolver { theme }, tint_mode)
}

// ===========================
//  Fill / Stroke parsing
// ===========================

fn parse_fill(node: roxmltree::Node<'_, '_>, theme: &HashMap<String, String>) -> Option<Fill> {
    for c in node.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "solidFill" => {
                // If the color resolves, use it. If not (e.g. phClr with no theme slot),
                // return None so the caller can fall back to the shape style color.
                if let Some(color) = parse_color_node(c, theme) {
                    return Some(Fill::Solid { color });
                }
                // Unresolvable → don't default to black; let fallback logic handle it
            }
            "noFill" => return Some(Fill::None),
            "pattFill" => {
                // ECMA-376 §20.1.8.40 — preset pattern with fg/bg colours.
                // Shared parse (ooxml_common::fill); colors resolve with pptx's
                // PowerPointLinear tint via PptxSchemeResolver.
                let ooxml_common::fill::PatternFill { fg, bg, preset } =
                    ooxml_common::fill::parse_patt_fill(
                        c,
                        &PptxSchemeResolver { theme },
                        ooxml_common::color::TintMode::PowerPointLinear,
                    );
                return Some(Fill::Pattern { fg, bg, preset });
            }
            "gradFill" => {
                // Shared parse (ooxml_common::fill). Returns None when there are
                // no resolvable stops, so we keep scanning sibling fill elements.
                if let Some(g) = ooxml_common::fill::parse_grad_fill(
                    c,
                    &PptxSchemeResolver { theme },
                    ooxml_common::color::TintMode::PowerPointLinear,
                ) {
                    return Some(Fill::Gradient {
                        stops: g.stops,
                        angle: g.angle,
                        grad_type: g.grad_type,
                    });
                }
            }
            _ => {}
        }
    }
    None
}

/// ECMA-376 §20.1.8.14 `a:blipFill` → `Fill::Image`. The `resolve_blip`
/// closure maps the `<a:blip r:embed>` rId to the blip's embedded **zip path**
/// using the caller's rels (each inheritance level resolves against its own
/// part); the mime is derived from that path. The renderer fetches the bytes
/// lazily by path rather than from an inlined data URL.
///
/// Both fill-modes are honoured and mutually exclusive:
/// - `stretch` (§20.1.8.56): the `fillRect` (§20.1.8.30) is captured so the
///   renderer can place the (possibly overscanned) image into the box.
/// - `tile` (§20.1.8.58): the tile offset/scale/flip/align descriptor is
///   captured so the renderer can repeat the blip at its native (scaled) size.
///
/// When neither child is present the blip defaults to full-box placement
/// (stretch with no fillRect).
fn parse_blip_fill<F: FnMut(&str) -> Option<String>>(
    blip_fill: roxmltree::Node<'_, '_>,
    resolve_blip: &mut F,
) -> Option<Fill> {
    let r_id = child(blip_fill, "blip").and_then(|b| attr_r(&b, "embed"))?;
    let image_path = resolve_blip(&r_id)?;
    let mime_type = mime_from_ext(&image_path).to_owned();
    let alpha = parse_blip_alpha(blip_fill);
    // §20.1.8.58 tile takes precedence when present (stretch/tile are an
    // either-or choice in CT_BlipFillProperties).
    if let Some(tile_node) = child(blip_fill, "tile") {
        return Some(Fill::Image {
            image_path,
            mime_type,
            fill_rect: None,
            tile: Some(parse_tile(tile_node)),
            alpha,
        });
    }
    let fill_rect = child(blip_fill, "stretch").and_then(parse_fill_rect);
    Some(Fill::Image {
        image_path,
        mime_type,
        fill_rect,
        tile: None,
        alpha,
    })
}

fn parse_arrow_end(node: roxmltree::Node<'_, '_>) -> ArrowEnd {
    let kind = attr(&node, "type").unwrap_or_else(|| "none".to_owned());
    let w = attr(&node, "w").unwrap_or_else(|| "med".to_owned());
    let len = attr(&node, "len").unwrap_or_else(|| "med".to_owned());
    ArrowEnd { kind, w, len }
}

fn parse_stroke(
    ln_node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Stroke> {
    if child(ln_node, "noFill").is_some() {
        return None;
    }
    let width = attr_i64(&ln_node, "w").unwrap_or(9525);
    let color = child(ln_node, "solidFill").and_then(|n| parse_color_node(n, theme))?;
    let dash_style = child(ln_node, "prstDash")
        .and_then(|n| attr(&n, "val"))
        .filter(|v| v != "solid");
    // Arrow ends — only emit when type != "none"
    let head_end = child(ln_node, "headEnd")
        .map(parse_arrow_end)
        .filter(|a| a.kind != "none");
    let tail_end = child(ln_node, "tailEnd")
        .map(parse_arrow_end)
        .filter(|a| a.kind != "none");
    // ECMA-376 §20.1.8.42 ST_CompoundLine. Default "sng" stays absent so the
    // renderer keeps its single-stroke fast path.
    let cmpd = attr(&ln_node, "cmpd").filter(|v| v != "sng");
    Some(Stroke {
        color,
        width,
        dash_style,
        head_end,
        tail_end,
        cmpd,
    })
}

// ===========================
//  Shadow parsing
// ===========================

/// Parse spPr > effectLst > outerShdw into a Shadow.
fn parse_shadow(
    effect_lst: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Shadow> {
    parse_shadow_node(child(effect_lst, "outerShdw")?, theme)
}

/// Parse spPr > effectLst > innerShdw into a Shadow. ECMA-376 §20.1.8.21
/// (CT_InnerShadowEffect) — same field shape as outerShdw, semantics differ
/// at render time (cast inward).
fn parse_inner_shadow(
    effect_lst: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Shadow> {
    parse_shadow_node(child(effect_lst, "innerShdw")?, theme)
}

/// Shared field reader for innerShdw / outerShdw. Both elements expose
/// blurRad, dist, dir, and a color child with optional alphaModFix.
fn parse_shadow_node(
    n: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Shadow> {
    let blur = attr_i64(&n, "blurRad").unwrap_or(0);
    let dist = attr_i64(&n, "dist").unwrap_or(0);
    let dir = attr_f64(&n, "dir").unwrap_or(0.0) / 60_000.0;

    let color_str = parse_color_node(n, theme).unwrap_or_else(|| "000000".to_owned());
    let (color, alpha) = if color_str.len() >= 8 {
        let a = u8::from_str_radix(&color_str[6..8], 16).unwrap_or(255) as f64 / 255.0;
        (color_str[..6].to_owned(), a)
    } else {
        (color_str, 1.0)
    };

    Some(Shadow {
        color,
        alpha,
        blur,
        dist,
        dir,
    })
}

/// Parse spPr > effectLst > glow into a Glow effect — ECMA-376 §20.1.8.17
/// (CT_GlowEffect): a coloured halo with a blur radius, no offset.
fn parse_glow(
    effect_lst: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Glow> {
    let g = child(effect_lst, "glow")?;
    let radius = attr_i64(&g, "rad").unwrap_or(0);
    let color_str = parse_color_node(g, theme).unwrap_or_else(|| "000000".to_owned());
    let (color, alpha) = if color_str.len() >= 8 {
        let a = u8::from_str_radix(&color_str[6..8], 16).unwrap_or(255) as f64 / 255.0;
        (color_str[..6].to_owned(), a)
    } else {
        (color_str, 1.0)
    };
    Some(Glow {
        color,
        alpha,
        radius,
    })
}

/// Parse spPr > effectLst > softEdge into a SoftEdge — ECMA-376 §20.1.8.31.
fn parse_soft_edge(effect_lst: roxmltree::Node<'_, '_>) -> Option<SoftEdge> {
    let n = child(effect_lst, "softEdge")?;
    let radius = attr_i64(&n, "rad").unwrap_or(0);
    Some(SoftEdge { radius })
}

/// Parse spPr > effectLst > reflection — ECMA-376 §20.1.8.27. Defaults
/// follow the spec table: blur=0, dist=0, dir=0, stA=100000 (=1.0),
/// stPos=0, endA=0, endPos=100000 (=1.0), sx=100000, sy=-100000.
fn parse_reflection(effect_lst: roxmltree::Node<'_, '_>) -> Option<Reflection> {
    let r = child(effect_lst, "reflection")?;
    let pct = |name: &str, default: f64| -> f64 {
        attr_f64(&r, name).map(|v| v / 100_000.0).unwrap_or(default)
    };
    Some(Reflection {
        blur: attr_i64(&r, "blurRad").unwrap_or(0),
        dist: attr_i64(&r, "dist").unwrap_or(0),
        dir: attr_f64(&r, "dir").unwrap_or(0.0) / 60_000.0,
        st_a: pct("stA", 1.0),
        st_pos: pct("stPos", 0.0),
        end_a: pct("endA", 0.0),
        end_pos: pct("endPos", 1.0),
        sx: pct("sx", 1.0),
        sy: pct("sy", -1.0),
    })
}

/// Effects pulled from `spPr > effectLst`. The five members are independent
/// siblings inside `CT_EffectList` — ECMA-376 §20.1.8.16. Used by both shapes
/// (`p:sp`) and pictures (`p:pic`): `p:spPr` is `CT_ShapeProperties` in both
/// cases (§19.3.1.37), so `effectLst` applies equally to images.
struct EffectLst {
    shadow: Option<Shadow>,
    inner_shadow: Option<Shadow>,
    glow: Option<Glow>,
    soft_edge: Option<SoftEdge>,
    reflection: Option<Reflection>,
}

/// Read every `effectLst` child shapes and pictures share. `effect_lst` is the
/// optional `<a:effectLst>` node; missing nodes yield an all-`None` result.
fn parse_effect_lst(
    effect_lst: Option<roxmltree::Node<'_, '_>>,
    theme: &HashMap<String, String>,
) -> EffectLst {
    EffectLst {
        shadow: effect_lst.and_then(|n| parse_shadow(n, theme)),
        inner_shadow: effect_lst.and_then(|n| parse_inner_shadow(n, theme)),
        glow: effect_lst.and_then(|n| parse_glow(n, theme)),
        soft_edge: effect_lst.and_then(parse_soft_edge),
        reflection: effect_lst.and_then(parse_reflection),
    }
}

// ===========================
//  3D scene parsing (scene3d / sp3d)
// ===========================

/// Parse `<a:rot>` (`CT_SphereCoords`, ECMA-376 §20.1.5.11). Angles are stored
/// in the XML as 60000ths of a degree; we convert to degrees. All three
/// attributes are required by the schema, but we default missing ones to 0 to
/// stay tolerant of malformed input.
fn parse_rot3d(rot: roxmltree::Node<'_, '_>) -> Rot3d {
    let deg = |name: &str| attr_f64(&rot, name).unwrap_or(0.0) / 60_000.0;
    Rot3d {
        lat: deg("lat"),
        lon: deg("lon"),
        rev: deg("rev"),
    }
}

/// Parse `<a:scene3d>` (`CT_Scene3D`, ECMA-376 §20.1.4.1.41). Requires a
/// `<a:camera>` child (§20.1.5.5); `<a:lightRig>` is optional for our purposes
/// (Phase A renders the camera only). Returns None when no camera is present.
fn parse_scene3d(sppr: roxmltree::Node<'_, '_>) -> Option<Scene3d> {
    let scene = child(sppr, "scene3d")?;
    let cam = child(scene, "camera")?;
    let camera = Camera3d {
        prst: attr(&cam, "prst")?,
        // §20.1.5.5: fov is an ST_FOVAngle in 60000ths of a degree.
        fov: attr_f64(&cam, "fov").map(|v| v / 60_000.0),
        // zoom is an ST_PositivePercentage (100000 = 100%).
        zoom: attr_f64(&cam, "zoom").map(|v| v / 100_000.0),
        rot: child(cam, "rot").map(parse_rot3d),
    };
    let light_rig = child(scene, "lightRig").and_then(|lr| {
        Some(LightRig {
            rig: attr(&lr, "rig")?,
            dir: attr(&lr, "dir")?,
            rot: child(lr, "rot").map(parse_rot3d),
        })
    });
    Some(Scene3d { camera, light_rig })
}

/// Parse `<a:bevel>` (`CT_Bevel`, ECMA-376 §20.1.5.3). `w`/`h` default to
/// 76200 EMU and `prst` to "circle" per the schema.
fn parse_bevel3d(bevel: roxmltree::Node<'_, '_>) -> Bevel3d {
    Bevel3d {
        w: attr_i64(&bevel, "w").unwrap_or(76_200),
        h: attr_i64(&bevel, "h").unwrap_or(76_200),
        prst: attr(&bevel, "prst").unwrap_or_else(|| "circle".into()),
    }
}

/// Parse `<a:sp3d>` (`CT_Shape3D`, ECMA-376 §20.1.5.12). Defaults follow the
/// schema: z=0, extrusionH=0, contourW=0, prstMaterial="warmMatte". Parsed in
/// full but not rendered in Phase A.
fn parse_sp3d(sppr: roxmltree::Node<'_, '_>) -> Option<Sp3d> {
    let n = child(sppr, "sp3d")?;
    // contourClr is colour-only here; pass an empty theme map because sp3d
    // contour colours in practice are srgbClr (no theme lookup needed) and this
    // parser has the theme threaded only into the line/fill paths.
    let contour_clr = child(n, "contourClr").and_then(|c| parse_color_node(c, &HashMap::new()));
    Some(Sp3d {
        z: attr_i64(&n, "z").unwrap_or(0),
        extrusion_h: attr_i64(&n, "extrusionH").unwrap_or(0),
        contour_w: attr_i64(&n, "contourW").unwrap_or(0),
        contour_clr,
        prst_material: attr(&n, "prstMaterial").unwrap_or_else(|| "warmMatte".into()),
        bevel_t: child(n, "bevelT").map(parse_bevel3d),
        bevel_b: child(n, "bevelB").map(parse_bevel3d),
    })
}

// ===========================
//  Custom geometry parsing
// ===========================

/// Parse a single path command node; coordinates are normalised to [0,1].
fn parse_path_cmd(cmd_node: roxmltree::Node<'_, '_>, path_w: f64, path_h: f64) -> Option<PathCmd> {
    match cmd_node.tag_name().name() {
        "moveTo" => {
            let pt = child(cmd_node, "pt")?;
            let x = attr_f64(&pt, "x")? / path_w;
            let y = attr_f64(&pt, "y")? / path_h;
            Some(PathCmd::MoveTo { x, y })
        }
        "lnTo" => {
            let pt = child(cmd_node, "pt")?;
            let x = attr_f64(&pt, "x")? / path_w;
            let y = attr_f64(&pt, "y")? / path_h;
            Some(PathCmd::LineTo { x, y })
        }
        "cubicBezTo" => {
            let pts: Vec<_> = children_vec(cmd_node, "pt");
            if pts.len() < 3 {
                return None;
            }
            let x1 = attr_f64(&pts[0], "x")? / path_w;
            let y1 = attr_f64(&pts[0], "y")? / path_h;
            let x2 = attr_f64(&pts[1], "x")? / path_w;
            let y2 = attr_f64(&pts[1], "y")? / path_h;
            let x = attr_f64(&pts[2], "x")? / path_w;
            let y = attr_f64(&pts[2], "y")? / path_h;
            Some(PathCmd::CubicBezTo {
                x1,
                y1,
                x2,
                y2,
                x,
                y,
            })
        }
        "arcTo" => {
            // wR/hR are radii in path-local units; stAng/swAng in 60000ths of a degree
            let wr = attr_f64(&cmd_node, "wR").unwrap_or(0.0) / path_w;
            let hr = attr_f64(&cmd_node, "hR").unwrap_or(0.0) / path_h;
            let st_ang = attr_f64(&cmd_node, "stAng").unwrap_or(0.0) / 60000.0;
            let sw_ang = attr_f64(&cmd_node, "swAng").unwrap_or(0.0) / 60000.0;
            Some(PathCmd::ArcTo {
                wr,
                hr,
                st_ang,
                sw_ang,
            })
        }
        "close" => Some(PathCmd::Close),
        _ => None,
    }
}

/// Parse custGeom > pathLst into a list of sub-paths (one per <a:path> element).
fn parse_cust_geom(cust_geom: roxmltree::Node<'_, '_>) -> Vec<Vec<PathCmd>> {
    let path_lst = match child(cust_geom, "pathLst") {
        Some(n) => n,
        None => return vec![],
    };

    path_lst
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "path")
        .map(|path_node| {
            let path_w = attr_f64(&path_node, "w").unwrap_or(1.0).max(1.0);
            let path_h = attr_f64(&path_node, "h").unwrap_or(1.0).max(1.0);
            path_node
                .children()
                .filter(|n| n.is_element())
                .filter_map(|cmd| parse_path_cmd(cmd, path_w, path_h))
                .collect()
        })
        .collect()
}

// ===========================
//  Transform (a:xfrm)
// ===========================

fn parse_xfrm(xfrm: roxmltree::Node<'_, '_>) -> Transform {
    let rot = attr_f64(&xfrm, "rot").unwrap_or(0.0) / 60000.0;
    let flip_h = attr(&xfrm, "flipH")
        .map(|v| v == "1" || v == "true")
        .unwrap_or(false);
    let flip_v = attr(&xfrm, "flipV")
        .map(|v| v == "1" || v == "true")
        .unwrap_or(false);
    let off = child(xfrm, "off");
    let ext = child(xfrm, "ext");
    Transform {
        x: off.and_then(|n| attr_i64(&n, "x")).unwrap_or(0),
        y: off.and_then(|n| attr_i64(&n, "y")).unwrap_or(0),
        cx: ext.and_then(|n| attr_i64(&n, "cx")).unwrap_or(0),
        cy: ext.and_then(|n| attr_i64(&n, "cy")).unwrap_or(0),
        rot,
        flip_h,
        flip_v,
    }
}

// ===========================
//  Group transform
// ===========================

// ===========================
//  Layout placeholder map
// ===========================

/// Keyed first by idx (integer), then by type string.
// `Clone` lets `parse_layout` cache one resolved `LayoutPlaceholders` per layout
// and hand each slide a copy to layer its per-slide master txStyles fallbacks
// onto without mutating the cached instance (D4).
#[derive(Default, Clone)]
struct LayoutPlaceholders {
    by_idx: HashMap<u32, Transform>,
    by_type: HashMap<String, Transform>,
    /// Fallback transforms from slide master (by ph_type), used when layout has no xfrm
    master_by_type: HashMap<String, Transform>,
    /// Default font size (pt) per placeholder idx, from layout/master lstStyle
    by_idx_font_size: HashMap<u32, f64>,
    /// Default font size (pt) per placeholder type, from layout/master lstStyle
    by_type_font_size: HashMap<String, f64>,
    /// Per-list-level default font sizes (pt) per placeholder idx — index 0..=8
    /// maps to lvl1pPr..lvl9pPr (ECMA-376 §21.1.2.4). Lets nested bullets shrink
    /// per level (e.g. body 28pt → lvl2 24pt → lvl3 20pt) instead of all using
    /// the level-1 size. None per level where the style chain doesn't specify it.
    by_idx_level_sizes: HashMap<u32, LevelFontSizes>,
    /// Per-list-level default font sizes (pt) per placeholder type.
    by_type_level_sizes: HashMap<String, LevelFontSizes>,
    /// Per-list-level paragraph indents (`marL`/`marR`/`indent`, EMU) per
    /// placeholder idx — what a paragraph with no own `marL`/`marR`/`indent`
    /// inherits from the authored list-style cascade (ECMA-376 §21.1.2.4.13),
    /// used as the fallback before PowerPoint's hardcoded implicit defaults.
    by_idx_level_indents: HashMap<u32, LevelIndents>,
    /// Per-list-level paragraph indents per placeholder type.
    by_type_level_indents: HashMap<String, LevelIndents>,
    /// Per-list-level inherited bullet (buChar/buAutoNum/buNone) per placeholder
    /// idx — what a paragraph with no explicit bullet inherits (ECMA-376 §19.7.10).
    by_idx_level_bullets: HashMap<u32, LevelBullets>,
    /// Per-list-level inherited bullet per placeholder type.
    by_type_level_bullets: HashMap<String, LevelBullets>,
    /// Default bold per placeholder type, from layout lstStyle defRPr b attribute
    by_type_bold: HashMap<String, bool>,
    /// Default italic per placeholder type, from layout lstStyle defRPr i attribute
    by_type_italic: HashMap<String, bool>,
    /// Default caps ("all"/"small") per placeholder type, from layout/master
    /// lstStyle defRPr cap attribute (ECMA-376 §21.1.2.3.13)
    by_type_caps: HashMap<String, String>,
    /// Vertical anchor ("t"/"ctr"/"b") per placeholder type, from layout/master bodyPr
    by_type_anchor: HashMap<String, String>,
    /// Default paragraph alignment per placeholder type, from layout/master lstStyle
    by_type_alignment: HashMap<String, String>,
    /// Paragraph alignment per placeholder idx — layout placeholder's own algn,
    /// falling back to the master per-type alignment. Checked before the
    /// type-keyed maps so a body placeholder resolves to its OWN idx's style,
    /// not an unrelated typeless placeholder (ECMA-376 §19.3.1.x idx matching).
    by_idx_alignment: HashMap<u32, String>,
    /// Default East Asian line-break (eaLnBrk) per placeholder type, from the
    /// layout lstStyle > lvl1pPr @eaLnBrk (ECMA-376 §21.1.2.2.7)
    by_type_ea_ln_brk: HashMap<String, bool>,
    /// Default space-before (hundredths of pt) per placeholder type, from layout lstStyle
    by_type_space_before: HashMap<String, i64>,
    /// Default space-after (hundredths of pt) per placeholder type, from layout lstStyle
    by_type_space_after: HashMap<String, i64>,
    /// Default space-before from master txStyles (fallback when layout has none)
    by_type_master_space_before: HashMap<String, i64>,
    /// Default space-after from master txStyles (fallback when layout has none)
    by_type_master_space_after: HashMap<String, i64>,
    /// Stroke per placeholder type from layout spPr > ln
    by_type_stroke: HashMap<String, Stroke>,
    /// Stroke per placeholder idx from layout spPr > ln
    by_idx_stroke: HashMap<u32, Stroke>,
    /// Default line spacing (spcPct val, e.g. 90000 = 90%) per placeholder idx, from layout lstStyle
    by_idx_line_spacing: HashMap<u32, f64>,
    /// Default line spacing (spcPct val) per placeholder type, from layout lstStyle
    by_type_line_spacing: HashMap<String, f64>,
    /// Paragraph alignment per placeholder type from master lstStyle > lvl1pPr algn (fallback)
    by_type_master_alignment: HashMap<String, String>,
    /// East Asian line-break per placeholder type from master lstStyle > lvl1pPr
    /// @eaLnBrk (fallback when the layout has none) — ECMA-376 §21.1.2.2.7
    by_type_master_ea_ln_brk: HashMap<String, bool>,
    /// Default line spacing from master txStyles (fallback when layout has none)
    by_type_master_line_spacing: HashMap<String, f64>,
    /// Inherited blipFill (data URL + src rect) per placeholder idx from layout spPr
    by_idx_blip_fill: HashMap<u32, InheritedBlipFill>,
    /// Inherited blipFill per placeholder type from layout spPr
    by_type_blip_fill: HashMap<String, InheritedBlipFill>,
    /// Default text color per placeholder idx, from layout lstStyle defRPr solidFill
    by_idx_color: HashMap<u32, String>,
    /// Default text color per placeholder type, from layout lstStyle defRPr solidFill
    by_type_color: HashMap<String, String>,
    /// Default text color from master (txStyles + spTree lstStyle) — fallback when layout has none
    by_type_master_color: HashMap<String, String>,
    /// `<p:spPr><a:solidFill | a:noFill | a:gradFill | a:pattFill>` per placeholder idx.
    /// Used to inherit a layout-level shape fill (e.g. a tinted body placeholder)
    /// onto slide-level shapes whose `<p:spPr>` is empty.
    by_idx_fill: HashMap<u32, Fill>,
    /// Same as `by_idx_fill` but keyed by placeholder type (fallback when idx
    /// doesn't match a layout shape).
    by_type_fill: HashMap<String, Fill>,
}

#[derive(Debug, Clone)]
struct InheritedBlipFill {
    /// Embedded zip path of the inherited picture-placeholder blip.
    image_path: String,
    /// MIME of the blip at `image_path`.
    mime_type: String,
    src_rect: Option<SrcRect>,
    alpha: Option<f64>,
}

impl LayoutPlaceholders {
    fn lookup(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<&Transform> {
        ph_idx
            .and_then(|i| self.by_idx.get(&i))
            .or_else(|| self.by_type.get(ph_type))
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type.get("")
                } else {
                    None
                }
            })
            .or_else(|| self.master_by_type.get(ph_type))
    }

    /// Look up the inherited default font size for a placeholder (layout then master fallback).
    /// Idx-strict per ECMA-376 §19.7.16 (see `lookup_fill`'s rationale).
    fn lookup_font_size(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<f64> {
        if let Some(i) = ph_idx {
            return self.by_idx_font_size.get(&i).copied();
        }
        self.by_type_font_size.get(ph_type).copied().or_else(|| {
            if ph_type == "body" {
                self.by_type_font_size.get("").copied()
            } else {
                None
            }
        })
    }

    /// Per-list-level inherited default font sizes (lvl1..lvl9). Same idx-strict
    /// resolution as `lookup_font_size`. All-None when the placeholder has no
    /// per-level styling.
    fn lookup_level_font_sizes(&self, ph_type: &str, ph_idx: Option<u32>) -> LevelFontSizes {
        if let Some(i) = ph_idx {
            return self
                .by_idx_level_sizes
                .get(&i)
                .copied()
                .unwrap_or([None; 9]);
        }
        self.by_type_level_sizes
            .get(ph_type)
            .copied()
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_level_sizes.get("").copied()
                } else {
                    None
                }
            })
            .unwrap_or([None; 9])
    }

    /// Per-list-level inherited paragraph indents (lvl1..lvl9). Same idx-strict
    /// resolution as `lookup_level_font_sizes`. All-default (every axis None) when
    /// the placeholder has no authored per-level indent.
    fn lookup_level_indents(&self, ph_type: &str, ph_idx: Option<u32>) -> LevelIndents {
        if let Some(i) = ph_idx {
            return self
                .by_idx_level_indents
                .get(&i)
                .copied()
                .unwrap_or_default();
        }
        self.by_type_level_indents
            .get(ph_type)
            .copied()
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_level_indents.get("").copied()
                } else {
                    None
                }
            })
            .unwrap_or_default()
    }

    /// Per-list-level inherited bullets (lvl1..lvl9). Same idx-strict resolution as
    /// `lookup_level_font_sizes`. All-None when the placeholder inherits no bullet.
    fn lookup_level_bullets(&self, ph_type: &str, ph_idx: Option<u32>) -> LevelBullets {
        if let Some(i) = ph_idx {
            return self
                .by_idx_level_bullets
                .get(&i)
                .cloned()
                .unwrap_or_else(empty_level_bullets);
        }
        self.by_type_level_bullets
            .get(ph_type)
            .cloned()
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_level_bullets.get("").cloned()
                } else {
                    None
                }
            })
            .unwrap_or_else(empty_level_bullets)
    }

    /// Look up inherited bold for this placeholder type.
    fn lookup_bold(&self, ph_type: &str) -> Option<bool> {
        self.by_type_bold.get(ph_type).copied().or_else(|| {
            if ph_type == "body" {
                self.by_type_bold.get("").copied()
            } else {
                None
            }
        })
    }

    /// Look up inherited italic for this placeholder type.
    fn lookup_italic(&self, ph_type: &str) -> Option<bool> {
        self.by_type_italic.get(ph_type).copied().or_else(|| {
            if ph_type == "body" {
                self.by_type_italic.get("").copied()
            } else {
                None
            }
        })
    }

    /// Look up inherited caps ("all"/"small") for this placeholder type.
    fn lookup_caps(&self, ph_type: &str) -> Option<String> {
        self.by_type_caps.get(ph_type).cloned().or_else(|| {
            if ph_type == "body" {
                self.by_type_caps.get("").cloned()
            } else {
                None
            }
        })
    }

    /// Look up inherited vertical anchor for this placeholder type.
    fn lookup_anchor(&self, ph_type: &str) -> Option<String> {
        self.by_type_anchor.get(ph_type).cloned().or_else(|| {
            if ph_type == "body" {
                self.by_type_anchor.get("").cloned()
            } else {
                None
            }
        })
    }

    /// Look up inherited paragraph alignment for this placeholder.
    ///
    /// A placeholder identified by `idx` resolves through its own slot
    /// (`by_idx_alignment`), which `parse_layout_placeholders` pre-seeds with the
    /// master per-type default. Unlike `lookup_fill`, falling through to the
    /// type map on an idx miss is intentional and safe (the seed already encodes
    /// the master tier) — but the `""` (typeless) fallback is gated to
    /// `ph_idx.is_none()` so an idx/typed placeholder never borrows an unrelated
    /// typeless sibling's alignment (ECMA-376 §19.3.1.36 idx matching).
    fn lookup_alignment(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<String> {
        if let Some(i) = ph_idx {
            if let Some(a) = self.by_idx_alignment.get(&i) {
                return Some(a.clone());
            }
        }
        // The `""` fallback represents a typeless (idx-less, body-category)
        // placeholder; only a placeholder that is itself typeless may use it.
        let allow_empty = ph_idx.is_none() && ph_type == "body";
        self.by_type_alignment
            .get(ph_type)
            .cloned()
            .or_else(|| {
                if allow_empty {
                    self.by_type_alignment.get("").cloned()
                } else {
                    None
                }
            })
            .or_else(|| self.by_type_master_alignment.get(ph_type).cloned())
            .or_else(|| {
                if allow_empty {
                    self.by_type_master_alignment.get("").cloned()
                } else {
                    None
                }
            })
    }

    // ECMA-376 §21.1.2.2.7 eaLnBrk inheritance, mirroring lookup_alignment:
    // layout per-type → layout generic ("") for body → master per-type →
    // master generic. None means no ancestor specified it (parse_paragraph then
    // applies the spec default of true).
    fn lookup_ea_ln_brk(&self, ph_type: &str) -> Option<bool> {
        self.by_type_ea_ln_brk
            .get(ph_type)
            .copied()
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_ea_ln_brk.get("").copied()
                } else {
                    None
                }
            })
            .or_else(|| self.by_type_master_ea_ln_brk.get(ph_type).copied())
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_master_ea_ln_brk.get("").copied()
                } else {
                    None
                }
            })
    }

    fn lookup_space_before(&self, ph_type: &str) -> Option<i64> {
        self.by_type_space_before
            .get(ph_type)
            .copied()
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_space_before.get("").copied()
                } else {
                    None
                }
            })
            .or_else(|| self.by_type_master_space_before.get(ph_type).copied())
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_master_space_before.get("").copied()
                } else {
                    None
                }
            })
    }

    fn lookup_space_after(&self, ph_type: &str) -> Option<i64> {
        self.by_type_space_after
            .get(ph_type)
            .copied()
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_space_after.get("").copied()
                } else {
                    None
                }
            })
            .or_else(|| self.by_type_master_space_after.get(ph_type).copied())
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_master_space_after.get("").copied()
                } else {
                    None
                }
            })
    }

    /// Look up inherited blipFill from the layout placeholder spPr. Used when a slide
    /// references a picture placeholder (e.g. ph type="pic") without its own blipFill —
    /// the image defined on the layout's matching placeholder should render through.
    /// Idx-strict per ECMA-376 §19.7.16 (see `lookup_fill`'s rationale).
    fn lookup_blip_fill(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<InheritedBlipFill> {
        if let Some(i) = ph_idx {
            return self.by_idx_blip_fill.get(&i).cloned();
        }
        self.by_type_blip_fill.get(ph_type).cloned()
    }

    /// Look up inherited stroke from the layout placeholder spPr > ln.
    /// Idx-strict per ECMA-376 §19.7.16 (see `lookup_fill`'s rationale).
    fn lookup_stroke(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<Stroke> {
        if let Some(i) = ph_idx {
            return self.by_idx_stroke.get(&i).cloned();
        }
        self.by_type_stroke.get(ph_type).cloned().or_else(|| {
            if ph_type == "body" {
                self.by_type_stroke.get("").cloned()
            } else {
                None
            }
        })
    }

    /// Look up inherited default text color for this placeholder (layout then master fallback).
    ///
    /// The *layout* tier is idx-strict per ECMA-376 §19.7.16: when the slide-level
    /// placeholder carries an explicit `idx`, a layout colour is inherited only from the
    /// layout shape with the SAME idx — never a sibling body placeholder at a different
    /// idx (which would leak an unrelated region's colour).
    ///
    /// The *master* `txStyles` tier (titleStyle/bodyStyle/otherStyle), however, is a
    /// document-wide default keyed by placeholder *type* (§21.1.2.4 / §19.3.1) and is
    /// inherited regardless of idx. So when the idx-matched layout shape defines no
    /// colour, resolution must still fall through to `by_type_master_color`. Without
    /// this, a body placeholder whose layout shape sets size-but-not-colour resolves to
    /// no colour at all and the renderer defaults to black — instead of the master
    /// bodyStyle colour (e.g. `schemeClr bg1` = white on a dark theme). (sample-9 slide 2+)
    fn lookup_color(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<String> {
        if let Some(i) = ph_idx {
            if let Some(c) = self.by_idx_color.get(&i) {
                return Some(c.clone());
            }
            // Layout idx had no colour → fall through to the master type-keyed default.
            return self.by_type_master_color.get(ph_type).cloned().or_else(|| {
                if ph_type == "body" {
                    self.by_type_master_color.get("").cloned()
                } else {
                    None
                }
            });
        }
        self.by_type_color
            .get(ph_type)
            .cloned()
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_color.get("").cloned()
                } else {
                    None
                }
            })
            .or_else(|| self.by_type_master_color.get(ph_type).cloned())
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_master_color.get("").cloned()
                } else {
                    None
                }
            })
    }

    /// Look up the inherited shape fill from the layout placeholder's `<p:spPr>`.
    /// Used when the slide-level shape leaves `<p:spPr>` empty (or with no fill
    /// elements) and is bound to a placeholder.
    ///
    /// ECMA-376 §19.7.16 (placeholder inheritance) is asymmetric: when the
    /// slide-level shape declares `<p:ph idx="N">` it is bound to *that*
    /// specific layout slot — the only valid inheritance source is the layout
    /// shape with idx=N. Falling back to `by_type_fill` here would let a
    /// sibling body placeholder (a different idx, different region of the
    /// layout) bleed its fill onto a placeholder that the spec says should
    /// have no fill. This is exactly what regressed sample-2 slide-4: layout10
    /// has `body[idx=12]` (header, no fill) and `body[idx=13]` (bullet box,
    /// gray fill) — the type fallback was leaking the bullet box's gray onto
    /// the header.
    ///
    /// The type-only fallback only applies when the slide-level shape itself
    /// has no idx, in which case "first body placeholder we found" is the
    /// best we can do.
    fn lookup_fill(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<Fill> {
        if let Some(i) = ph_idx {
            return self.by_idx_fill.get(&i).cloned();
        }
        self.by_type_fill.get(ph_type).cloned().or_else(|| {
            if ph_type == "body" {
                self.by_type_fill.get("").cloned()
            } else {
                None
            }
        })
    }

    /// Look up inherited line spacing (spcPct val, e.g. 90000 = 90%) for this placeholder.
    /// Idx-strict per ECMA-376 §19.7.16 (see `lookup_fill`'s rationale).
    fn lookup_line_spacing(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<f64> {
        if let Some(i) = ph_idx {
            return self.by_idx_line_spacing.get(&i).copied();
        }
        self.by_type_line_spacing
            .get(ph_type)
            .copied()
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_line_spacing.get("").copied()
                } else {
                    None
                }
            })
            .or_else(|| self.by_type_master_line_spacing.get(ph_type).copied())
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_master_line_spacing.get("").copied()
                } else {
                    None
                }
            })
    }
}

/// Extract the lvl1pPr defRPr font size from a txBody node.
fn extract_lvl1_font_size(tx_body: roxmltree::Node<'_, '_>) -> Option<f64> {
    child(tx_body, "lstStyle")
        .and_then(|ls| child(ls, "lvl1pPr"))
        .and_then(|lp| child(lp, "defRPr"))
        .and_then(|rp| attr_f64(&rp, "sz"))
        .map(|v| v / 100.0)
}

/// Per-list-level default font sizes (pt). Index 0..=8 → lvl1pPr..lvl9pPr
/// (ECMA-376 §21.1.2.4). `None` where the level isn't specified.
type LevelFontSizes = [Option<f64>; 9];

/// Read `<a:lvlNpPr><a:defRPr@sz>` for levels 1..9 from a node that holds
/// `<a:lvlNpPr>` children — a txBody's `<a:lstStyle>` or a master `<p:txStyles>`
/// style node (`<p:bodyStyle>` etc.). Sizes are in pt.
fn read_level_font_sizes(list_style: roxmltree::Node<'_, '_>) -> LevelFontSizes {
    let mut out: LevelFontSizes = [None; 9];
    for (lvl, slot) in out.iter_mut().enumerate() {
        let tag = format!("lvl{}pPr", lvl + 1);
        *slot = list_style
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == tag)
            .and_then(|lp| child(lp, "defRPr"))
            .and_then(|rp| attr_f64(&rp, "sz"))
            .map(|v| v / 100.0);
    }
    out
}

/// Per-level default font sizes from a txBody's own `<a:lstStyle>`.
fn extract_level_font_sizes(tx_body: roxmltree::Node<'_, '_>) -> LevelFontSizes {
    child(tx_body, "lstStyle")
        .map(read_level_font_sizes)
        .unwrap_or([None; 9])
}

/// True when any level carries a size (avoids storing all-None arrays).
fn has_any_level_size(s: &LevelFontSizes) -> bool {
    s.iter().any(|v| v.is_some())
}

/// Per-edge merge: `primary[lvl]` wins, else `fallback[lvl]`.
fn merge_level_sizes(primary: &LevelFontSizes, fallback: &LevelFontSizes) -> LevelFontSizes {
    let mut out: LevelFontSizes = [None; 9];
    for lvl in 0..9 {
        out[lvl] = primary[lvl].or(fallback[lvl]);
    }
    out
}

/// Per-list-level paragraph indents (EMU) — the `marL`/`marR`/`indent` attributes
/// of a `<a:lvlNpPr>` (ECMA-376 §21.1.2.4.13; `lvlNpPr` is a
/// `CT_TextParagraphProperties`, so these are attributes ON the level element
/// itself, exactly like a paragraph's own `<a:pPr>`). Each axis is `Option` so it
/// inherits independently: a level that sets only `marL` leaves `marR`/`indent`
/// `None` and a lower-priority tier supplies them.
#[derive(Clone, Copy, Default)]
struct LevelIndent {
    mar_l: Option<i64>,
    mar_r: Option<i64>,
    indent: Option<i64>,
}
type LevelIndents = [LevelIndent; 9];

/// Read `<a:lvlNpPr@marL/@marR/@indent>` (EMU) for levels 1..9 from a node that
/// holds `<a:lvlNpPr>` children — a txBody's `<a:lstStyle>` or a master
/// `<p:txStyles>` style node. Mirrors `read_level_font_sizes`, but the values are
/// attributes of the `lvlNpPr` element itself (not of a `<a:defRPr>` child).
fn read_level_indents(list_style: roxmltree::Node<'_, '_>) -> LevelIndents {
    let mut out: LevelIndents = Default::default();
    for (lvl, slot) in out.iter_mut().enumerate() {
        let tag = format!("lvl{}pPr", lvl + 1);
        if let Some(lp) = list_style
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == tag)
        {
            slot.mar_l = attr_i64(&lp, "marL");
            slot.mar_r = attr_i64(&lp, "marR");
            slot.indent = attr_i64(&lp, "indent");
        }
    }
    out
}

/// Per-level indents from a txBody's own `<a:lstStyle>`.
fn extract_level_indents(tx_body: roxmltree::Node<'_, '_>) -> LevelIndents {
    child(tx_body, "lstStyle")
        .map(read_level_indents)
        .unwrap_or_default()
}

/// True when any level carries any indent axis (avoids storing all-None arrays).
fn has_any_level_indent(s: &LevelIndents) -> bool {
    s.iter()
        .any(|li| li.mar_l.is_some() || li.mar_r.is_some() || li.indent.is_some())
}

/// Per-level, per-axis merge: `primary[lvl].x` wins, else `fallback[lvl].x`.
fn merge_level_indents(primary: &LevelIndents, fallback: &LevelIndents) -> LevelIndents {
    let mut out: LevelIndents = Default::default();
    for lvl in 0..9 {
        out[lvl].mar_l = primary[lvl].mar_l.or(fallback[lvl].mar_l);
        out[lvl].mar_r = primary[lvl].mar_r.or(fallback[lvl].mar_r);
        out[lvl].indent = primary[lvl].indent.or(fallback[lvl].indent);
    }
    out
}

/// Per-list-level bullet definitions (index 0..=8 → lvl1pPr..lvl9pPr).
/// `None` where the level's `<a:lvlNpPr>` declares no `buChar`/`buAutoNum`/`buNone`
/// (so the value is still inherited from a lower-priority style tier).
type LevelBullets = [Option<Bullet>; 9];

fn empty_level_bullets() -> LevelBullets {
    std::array::from_fn(|_| None)
}

/// True when any level carries an explicit bullet (avoids storing all-None arrays).
fn has_any_level_bullet(s: &LevelBullets) -> bool {
    s.iter().any(|v| v.is_some())
}

/// Read `<a:lvlNpPr>` bullets for levels 1..9 from a node holding `<a:lvlNpPr>`
/// children (a txBody `<a:lstStyle>` or a master `<p:txStyles>` style node).
/// A level resolves to `Some` only when it explicitly sets `buChar`/`buAutoNum`/
/// `buNone`; an absent bullet element stays `None` so lower tiers can supply it.
fn read_level_bullets<F: FnMut(&str) -> Option<String>>(
    list_style: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    resolve_blip: &mut F,
) -> LevelBullets {
    std::array::from_fn(|lvl| {
        let tag = format!("lvl{}pPr", lvl + 1);
        list_style
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == tag)
            .and_then(|lp| match parse_bullet(Some(lp), theme, resolve_blip) {
                Bullet::Inherit => None,
                b => Some(b),
            })
    })
}

/// Per-level bullets from a txBody's own `<a:lstStyle>`. `resolve_blip` resolves
/// a level's `<a:buBlip>` embed against this text body's part rels (§21.1.2.4.2).
fn extract_level_bullets<F: FnMut(&str) -> Option<String>>(
    tx_body: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    resolve_blip: &mut F,
) -> LevelBullets {
    child(tx_body, "lstStyle")
        .map(|ls| read_level_bullets(ls, theme, resolve_blip))
        .unwrap_or_else(empty_level_bullets)
}

/// Per-edge merge: `primary[lvl]` wins, else `fallback[lvl]`.
fn merge_level_bullets(primary: &LevelBullets, fallback: &LevelBullets) -> LevelBullets {
    std::array::from_fn(|lvl| primary[lvl].clone().or_else(|| fallback[lvl].clone()))
}

/// Parse bodyPr anchor ("t"/"ctr"/"b") from master placeholder shapes.
///
/// Takes the already-parsed master root element (`<p:sldMaster>`) so
/// `build_master_bundle` can parse the master XML once and share the
/// `Document` across every `parse_master_*` extractor (ECMA-376 §19.3.1.42).
fn parse_master_anchors(root: roxmltree::Node<'_, '_>) -> HashMap<String, String> {
    let mut map = HashMap::new();
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        for sp in sp_tree
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "sp")
        {
            let ph_node = sp
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "ph");
            if let Some(ph) = ph_node {
                let ph_type = attr(&ph, "type").unwrap_or_default();
                if let Some(anchor) = child(sp, "txBody")
                    .and_then(|tb| child(tb, "bodyPr"))
                    .and_then(|bp| attr(&bp, "anchor"))
                {
                    map.entry(ph_type).or_insert(anchor.to_string());
                }
            }
        }
    }
    map
}

/// txStyles style node → the placeholder types it defaults. ECMA-376 §19.3.1.52
/// txStyles → titleStyle §19.3.1.49 / bodyStyle §19.3.1.5 / otherStyle §19.3.1.35.
const MASTER_TXSTYLE_PH_TYPES: &[(&str, &[&str])] = &[
    ("titleStyle", &["title", "ctrTitle"]),
    ("bodyStyle", &["body", "subTitle", "obj", ""]),
    ("otherStyle", &["dt", "ftr", "sldNum"]),
];

/// Parse paragraph alignment from master placeholder shapes' lstStyle > lvl1pPr algn attribute.
/// Takes the shared, already-parsed master root (see `parse_master_anchors`).
fn parse_master_alignments(root: roxmltree::Node<'_, '_>) -> HashMap<String, String> {
    let mut map = HashMap::new();
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        for sp in sp_tree
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "sp")
        {
            let ph_node = sp
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "ph");
            if let Some(ph) = ph_node {
                let ph_type = attr(&ph, "type").unwrap_or_default();
                if let Some(algn) = child(sp, "txBody")
                    .and_then(|tb| child(tb, "lstStyle"))
                    .and_then(|ls| child(ls, "lvl1pPr"))
                    .and_then(|lp| attr(&lp, "algn"))
                {
                    map.entry(ph_type).or_insert(algn);
                }
            }
        }
    }
    // Fallback: master <p:txStyles> paragraph alignment (ECMA-376 §19.3.1.52
    // txStyles → titleStyle §19.3.1.49 / bodyStyle §19.3.1.5 / otherStyle §19.3.1.35).
    // Per-shape lstStyle (scanned above) wins via or_insert; this fills types
    // whose master placeholder shape carried no explicit algn (the common case —
    // PowerPoint stores title/body alignment in txStyles, not the shape lstStyle).
    if let Some(tx_styles) = child(root, "txStyles") {
        for &(style, types) in MASTER_TXSTYLE_PH_TYPES {
            if let Some(algn) = child(tx_styles, style)
                .and_then(|s| child(s, "lvl1pPr"))
                .and_then(|lp| attr(&lp, "algn"))
            {
                for t in types {
                    map.entry((*t).to_string()).or_insert_with(|| algn.clone());
                }
            }
        }
    }
    map
}

/// Parse master-level default East Asian line-break (eaLnBrk) per placeholder
/// type from each placeholder shape's lstStyle > lvl1pPr @eaLnBrk
/// (ECMA-376 §21.1.2.2.7). Mirrors parse_master_alignments. xsd:boolean.
fn parse_master_ea_ln_brk(root: roxmltree::Node<'_, '_>) -> HashMap<String, bool> {
    let mut map = HashMap::new();
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        for sp in sp_tree
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "sp")
        {
            let ph_node = sp
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "ph");
            if let Some(ph) = ph_node {
                let ph_type = attr(&ph, "type").unwrap_or_default();
                if let Some(v) = child(sp, "txBody")
                    .and_then(|tb| child(tb, "lstStyle"))
                    .and_then(|ls| child(ls, "lvl1pPr"))
                    .and_then(|lp| attr(&lp, "eaLnBrk"))
                {
                    map.entry(ph_type).or_insert(v == "1" || v == "true");
                }
            }
        }
    }
    map
}

/// Parse master-level default font sizes from txStyles (titleStyle / bodyStyle / otherStyle)
/// and from individual placeholder shapes in the master spTree.
/// Individual shape lstStyle takes priority over txStyles generic defaults.
fn parse_master_font_sizes(root: roxmltree::Node<'_, '_>) -> HashMap<String, f64> {
    let mut map = HashMap::new();

    // Scan master spTree placeholder shapes first — per-shape lstStyle is more specific
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        for sp in sp_tree
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "sp")
        {
            let ph_node = sp
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "ph");
            if let Some(ph) = ph_node {
                let ph_type = attr(&ph, "type").unwrap_or_default();
                if let Some(tx_body) = child(sp, "txBody") {
                    if let Some(sz) = extract_lvl1_font_size(tx_body) {
                        map.entry(ph_type).or_insert(sz);
                    }
                }
            }
        }
    }

    // p:txStyles > a:titleStyle / a:bodyStyle / a:otherStyle as fallback
    if let Some(tx_styles) = child(root, "txStyles") {
        let style_ph_map: &[(&str, &[&str])] = MASTER_TXSTYLE_PH_TYPES;
        for (style_name, ph_types) in style_ph_map {
            let sz = child(tx_styles, style_name)
                .and_then(|sn| child(sn, "lvl1pPr"))
                .and_then(|lp| child(lp, "defRPr"))
                .and_then(|rp| attr_f64(&rp, "sz"))
                .map(|v| v / 100.0);
            if let Some(fs) = sz {
                for ph_type in *ph_types {
                    map.entry(ph_type.to_string()).or_insert(fs);
                }
            }
        }
    }

    map
}

/// Per-list-level default font sizes from the master, keyed by ph_type. Mirrors
/// `parse_master_font_sizes` but captures every list level (lvl1pPr..lvl9pPr) so
/// nested bullets inherit the correct shrinking sizes (ECMA-376 §21.1.2.4),
/// not just the level-1 size. Per-shape lstStyle wins over the generic txStyles.
fn parse_master_level_font_sizes(root: roxmltree::Node<'_, '_>) -> HashMap<String, LevelFontSizes> {
    let mut map: HashMap<String, LevelFontSizes> = HashMap::new();

    // Per-shape lstStyle first (more specific).
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        for sp in sp_tree
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "sp")
        {
            if let Some(ph) = sp
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "ph")
            {
                let ph_type = attr(&ph, "type").unwrap_or_default();
                if let Some(tx_body) = child(sp, "txBody") {
                    let sizes = extract_level_font_sizes(tx_body);
                    if has_any_level_size(&sizes) {
                        map.entry(ph_type).or_insert(sizes);
                    }
                }
            }
        }
    }

    // txStyles fallback.
    if let Some(tx_styles) = child(root, "txStyles") {
        let style_ph_map: &[(&str, &[&str])] = MASTER_TXSTYLE_PH_TYPES;
        for (style_name, ph_types) in style_ph_map {
            if let Some(style_node) = child(tx_styles, style_name) {
                let sizes = read_level_font_sizes(style_node);
                if has_any_level_size(&sizes) {
                    for ph_type in *ph_types {
                        map.entry(ph_type.to_string()).or_insert(sizes);
                    }
                }
            }
        }
    }

    map
}

/// Per-list-level paragraph indents (`marL`/`marR`/`indent`, EMU) from the master,
/// keyed by ph_type. Mirrors `parse_master_level_font_sizes` exactly (same per-shape
/// lstStyle then `txStyles` tiers via `MASTER_TXSTYLE_PH_TYPES`): a master body
/// `<a:lvlNpPr@marL>` is what a slide body paragraph with no own `marL` inherits
/// (ECMA-376 §21.1.2.4.13). KNOWN SHARED GAP: no presentation `defaultTextStyle`
/// tier (§19.2.1.8, the lowest authored fallback) — the parser reads it for neither
/// indents nor font sizes nor bullets, so this stays at parity rather than adding a
/// tier only here; closing it is a separate cross-cutting change.
fn parse_master_level_indents(root: roxmltree::Node<'_, '_>) -> HashMap<String, LevelIndents> {
    let mut map: HashMap<String, LevelIndents> = HashMap::new();

    // Per-shape lstStyle first (more specific).
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        for sp in sp_tree
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "sp")
        {
            if let Some(ph) = sp
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "ph")
            {
                let ph_type = attr(&ph, "type").unwrap_or_default();
                if let Some(tx_body) = child(sp, "txBody") {
                    let indents = extract_level_indents(tx_body);
                    if has_any_level_indent(&indents) {
                        map.entry(ph_type).or_insert(indents);
                    }
                }
            }
        }
    }

    // txStyles fallback.
    if let Some(tx_styles) = child(root, "txStyles") {
        let style_ph_map: &[(&str, &[&str])] = MASTER_TXSTYLE_PH_TYPES;
        for (style_name, ph_types) in style_ph_map {
            if let Some(style_node) = child(tx_styles, style_name) {
                let indents = read_level_indents(style_node);
                if has_any_level_indent(&indents) {
                    for ph_type in *ph_types {
                        map.entry(ph_type.to_string()).or_insert(indents);
                    }
                }
            }
        }
    }

    map
}

/// Per-list-level bullets from the master, keyed by ph_type. Mirrors
/// `parse_master_level_font_sizes`: a master body placeholder's `<a:buChar>` (or
/// the `bodyStyle` `<a:lvlNpPr>` bullets) is what a slide body paragraph with no
/// explicit bullet inherits (ECMA-376 §19.7.10 / §21.1.2.4). Per-shape lstStyle
/// wins over the generic txStyles.
fn parse_master_level_bullets(
    root: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    master_rels: &HashMap<String, String>,
    master_dir: &str,
    zip: &mut PptxZip,
) -> HashMap<String, LevelBullets> {
    let mut map: HashMap<String, LevelBullets> = HashMap::new();

    // A master-level `<a:buBlip>` embed resolves against the master's rels +
    // part directory (ECMA-376 §21.1.2.4.2), mirroring the master background.
    let mut resolve_blip = |rid: &str| -> Option<String> {
        let target = master_rels.get(rid)?;
        let path = resolve_path(master_dir, target);
        // Verify the part exists so a listed-but-missing rId yields None and the
        // bullet falls through to Bullet::Inherit (matches the variant's doc
        // comment), mirroring the master background resolver. `index_for_name`
        // checks the central directory only (no inflate), unlike the former
        // `read_zip_bytes` which decompressed the entry just to discard it.
        zip.index_for_name(&path)?;
        Some(path)
    };

    // Per-shape lstStyle first (more specific).
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        for sp in sp_tree
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "sp")
        {
            if let Some(ph) = sp
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "ph")
            {
                let ph_type = attr(&ph, "type").unwrap_or_default();
                if let Some(tx_body) = child(sp, "txBody") {
                    let bullets = extract_level_bullets(tx_body, theme, &mut resolve_blip);
                    if has_any_level_bullet(&bullets) {
                        map.entry(ph_type).or_insert(bullets);
                    }
                }
            }
        }
    }

    // txStyles fallback.
    if let Some(tx_styles) = child(root, "txStyles") {
        let style_ph_map: &[(&str, &[&str])] = MASTER_TXSTYLE_PH_TYPES;
        for (style_name, ph_types) in style_ph_map {
            if let Some(style_node) = child(tx_styles, style_name) {
                let bullets = read_level_bullets(style_node, theme, &mut resolve_blip);
                if has_any_level_bullet(&bullets) {
                    for ph_type in *ph_types {
                        map.entry(ph_type.to_string())
                            .or_insert_with(|| bullets.clone());
                    }
                }
            }
        }
    }

    map
}

/// Parse default bold/italic from master txStyles (titleStyle / bodyStyle / otherStyle)
/// > lvl1pPr > defRPr @b and @i. Keyed by ph_type.
/// > Only populated when the attribute is explicitly present on the master.
fn parse_master_txstyle_bold_italic(
    root: roxmltree::Node<'_, '_>,
) -> (
    HashMap<String, bool>,
    HashMap<String, bool>,
    HashMap<String, String>,
) {
    let mut bold_map: HashMap<String, bool> = HashMap::new();
    let mut italic_map: HashMap<String, bool> = HashMap::new();
    // ECMA-376 §21.1.2.3.13 cap="all"/"small" on the master txStyles defRPr —
    // e.g. a template titleStyle with cap="all" upper-cases every title.
    let mut caps_map: HashMap<String, String> = HashMap::new();
    let Some(tx_styles) = child(root, "txStyles") else {
        return (bold_map, italic_map, caps_map);
    };
    let style_ph_map: &[(&str, &[&str])] = MASTER_TXSTYLE_PH_TYPES;
    for (style_name, ph_types) in style_ph_map {
        let def_rpr = child(tx_styles, style_name)
            .and_then(|sn| child(sn, "lvl1pPr"))
            .and_then(|lp| child(lp, "defRPr"));
        let b = def_rpr
            .and_then(|rp| attr(&rp, "b"))
            .map(|v| v == "1" || v == "true");
        let i = def_rpr
            .and_then(|rp| attr(&rp, "i"))
            .map(|v| v == "1" || v == "true");
        let c = def_rpr
            .and_then(|rp| attr(&rp, "cap"))
            .filter(|v| v == "all" || v == "small");
        if let Some(bv) = b {
            for t in *ph_types {
                bold_map.entry(t.to_string()).or_insert(bv);
            }
        }
        if let Some(iv) = i {
            for t in *ph_types {
                italic_map.entry(t.to_string()).or_insert(iv);
            }
        }
        if let Some(cv) = c {
            for t in *ph_types {
                caps_map.entry(t.to_string()).or_insert(cv.clone());
            }
        }
    }
    (bold_map, italic_map, caps_map)
}

/// Parse default text color from master txStyles (titleStyle/bodyStyle/otherStyle)
/// > lvl1pPr > defRPr > solidFill, and from per-placeholder shapes in the master spTree's
/// > txBody > lstStyle > lvl1pPr > defRPr > solidFill. Keyed by ph_type.
/// > Shape-level lstStyle takes priority over txStyles generic defaults.
fn parse_master_txstyle_color(
    root: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> HashMap<String, String> {
    let mut map: HashMap<String, String> = HashMap::new();

    // Scan master spTree placeholder shapes first — per-shape lstStyle is more specific.
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        for sp in sp_tree
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "sp")
        {
            let ph_node = sp
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "ph");
            if let Some(ph) = ph_node {
                let ph_type = attr(&ph, "type").unwrap_or_default();
                if let Some(color) = child(sp, "txBody")
                    .and_then(|tb| child(tb, "lstStyle"))
                    .and_then(|ls| child(ls, "lvl1pPr"))
                    .and_then(|lp| child(lp, "defRPr"))
                    .and_then(|rp| child(rp, "solidFill"))
                    .and_then(|sf| parse_color_node(sf, theme))
                {
                    map.entry(ph_type).or_insert(color);
                }
            }
        }
    }

    // Fall back to p:txStyles > titleStyle/bodyStyle/otherStyle > lvl1pPr > defRPr > solidFill.
    if let Some(tx_styles) = child(root, "txStyles") {
        let style_ph_map: &[(&str, &[&str])] = MASTER_TXSTYLE_PH_TYPES;
        for (style_name, ph_types) in style_ph_map {
            if let Some(color) = child(tx_styles, style_name)
                .and_then(|sn| child(sn, "lvl1pPr"))
                .and_then(|lp| child(lp, "defRPr"))
                .and_then(|rp| child(rp, "solidFill"))
                .and_then(|sf| parse_color_node(sf, theme))
            {
                for ph_type in *ph_types {
                    map.entry(ph_type.to_string()).or_insert(color.clone());
                }
            }
        }
    }

    map
}

/// Parse default paragraph spacing from master txStyles.
/// Returns (space_before_map, space_after_map, line_spacing_map) keyed by ph_type string.
/// space_before/after values are in hundredths of a point (same as Paragraph.space_before/after).
/// Note: line_spacing_map is intentionally NOT populated. Inheriting txStyles lnSpc hurts VRT
/// scores because our font substitutes (sans-serif) have different em-square metrics than the
/// original Aptos font, so applying the master's 120% line spacing over-expands text layout.
fn parse_master_txstyle_spacing(
    root: roxmltree::Node<'_, '_>,
) -> (
    HashMap<String, i64>,
    HashMap<String, i64>,
    HashMap<String, f64>,
) {
    let mut before_map: HashMap<String, i64> = HashMap::new();
    let mut after_map: HashMap<String, i64> = HashMap::new();
    let line_map: HashMap<String, f64> = HashMap::new(); // intentionally not populated
    let tx_styles = match child(root, "txStyles") {
        Some(n) => n,
        None => return (before_map, after_map, line_map),
    };
    let style_ph_map: &[(&str, &[&str])] = MASTER_TXSTYLE_PH_TYPES;
    for (style_name, ph_types) in style_ph_map {
        let lvl1 = child(tx_styles, style_name).and_then(|sn| child(sn, "lvl1pPr"));
        let spc_before = lvl1
            .and_then(|lp| child(lp, "spcBef"))
            .and_then(|s| child(s, "spcPts").and_then(|n| attr_i64(&n, "val")));
        let spc_after = lvl1
            .and_then(|lp| child(lp, "spcAft"))
            .and_then(|s| child(s, "spcPts").and_then(|n| attr_i64(&n, "val")));
        if let Some(v) = spc_before {
            for ph_type in *ph_types {
                before_map.entry(ph_type.to_string()).or_insert(v);
            }
        }
        if let Some(v) = spc_after {
            for ph_type in *ph_types {
                after_map.entry(ph_type.to_string()).or_insert(v);
            }
        }
    }
    (before_map, after_map, line_map)
}

fn parse_master_transforms(root: roxmltree::Node<'_, '_>) -> HashMap<String, Transform> {
    let mut map = HashMap::new();
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        for sp in sp_tree
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "sp")
        {
            let ph_node = sp
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "ph");
            if let Some(ph) = ph_node {
                let ph_type = attr(&ph, "type").unwrap_or_default();
                if let Some(xfrm) = child(sp, "spPr").and_then(|p| child(p, "xfrm")) {
                    map.entry(ph_type).or_insert_with(|| parse_xfrm(xfrm));
                }
            }
        }
    }
    map
}

// Seeds layout placeholders from the master's per-type defaults (transforms,
// alignment, spacing) before overlaying the layout's own placeholder props; the
// many maps are the master inheritance sources, threaded through as-is.
//
// Takes the already-parsed layout root (`<p:sldLayout>`) so `parse_layout` can
// parse the layout XML once and share the `Document` with the background +
// showMasterSp extractions (D4).
#[allow(clippy::too_many_arguments)]
fn parse_layout_placeholders(
    root: roxmltree::Node<'_, '_>,
    master_font_sizes: &HashMap<String, f64>,
    master_level_font_sizes: &HashMap<String, LevelFontSizes>,
    master_level_indents: &HashMap<String, LevelIndents>,
    master_level_bullets: &HashMap<String, LevelBullets>,
    master_anchors: &HashMap<String, String>,
    master_transforms: &HashMap<String, Transform>,
    master_alignments: &HashMap<String, String>,
    master_ea_ln_brk: &HashMap<String, bool>,
    master_space_before: &HashMap<String, i64>,
    master_space_after: &HashMap<String, i64>,
    master_line_spacing: &HashMap<String, f64>,
    theme: &HashMap<String, String>,
    layout_dir: &str,
    layout_rels: &HashMap<String, String>,
    zip: &mut PptxZip,
) -> LayoutPlaceholders {
    let mut lph = LayoutPlaceholders {
        master_by_type: master_transforms.clone(),
        by_type_master_alignment: master_alignments.clone(),
        by_type_master_ea_ln_brk: master_ea_ln_brk.clone(),
        by_type_master_space_before: master_space_before.clone(),
        by_type_master_space_after: master_space_after.clone(),
        by_type_master_line_spacing: master_line_spacing.clone(),
        ..Default::default()
    };

    let sp_tree = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "spTree");
    let sp_tree = match sp_tree {
        Some(n) => n,
        None => return lph,
    };

    for sp in sp_tree
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "sp")
    {
        let ph_node = sp
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "ph");
        let sp_pr = match child(sp, "spPr") {
            Some(n) => n,
            None => continue,
        };
        // xfrm may be absent (placeholder inherits transform from master); parse if present
        let t_opt: Option<Transform> = child(sp_pr, "xfrm").map(parse_xfrm);

        // Extract layout-level defaults from the placeholder's txBody > lstStyle > lvl1pPr
        let layout_lvl1_ppr: Option<roxmltree::Node<'_, '_>> = child(sp, "txBody")
            .and_then(|tb| child(tb, "lstStyle"))
            .and_then(|ls| child(ls, "lvl1pPr"));
        let layout_def_rpr: Option<roxmltree::Node<'_, '_>> =
            layout_lvl1_ppr.and_then(|lp| child(lp, "defRPr"));
        let layout_font_size = layout_def_rpr
            .and_then(|rp| attr_f64(&rp, "sz"))
            .map(|v| v / 100.0);
        // Per-level sizes from the layout placeholder's own lstStyle (all
        // lvlNpPr), used to give nested bullets their shrinking sizes.
        let layout_level_sizes: LevelFontSizes = child(sp, "txBody")
            .map(extract_level_font_sizes)
            .unwrap_or([None; 9]);
        // Per-level indents (marL/marR/indent) from the layout placeholder's own
        // lstStyle, the inherited list-indent cascade (ECMA-376 §21.1.2.4.13).
        let layout_level_indents: LevelIndents = child(sp, "txBody")
            .map(extract_level_indents)
            .unwrap_or_default();
        // Per-level bullets from the layout placeholder's own lstStyle. A
        // level's `<a:buBlip>` embed (§21.1.2.4.2) resolves against the layout's
        // rels + part directory, mirroring the layout-spPr blipFill above.
        let mut resolve_layout_blip = |rid: &str| -> Option<String> {
            let target = layout_rels.get(rid)?;
            let path = resolve_path(layout_dir, target);
            // Verify the part exists so a listed-but-missing rId yields None and
            // the bullet falls through to Bullet::Inherit (matches the variant's
            // doc comment), mirroring the master/layout background resolvers.
            // `index_for_name` reads the central directory only (no inflate),
            // unlike the former `read_zip_bytes` which decompressed and discarded.
            zip.index_for_name(&path)?;
            Some(path)
        };
        let layout_level_bullets: LevelBullets = child(sp, "txBody")
            .map(|tb| extract_level_bullets(tb, theme, &mut resolve_layout_blip))
            .unwrap_or_else(empty_level_bullets);
        let layout_bold = layout_def_rpr
            .and_then(|rp| attr(&rp, "b"))
            .map(|v| v == "1" || v == "true");
        let layout_italic = layout_def_rpr
            .and_then(|rp| attr(&rp, "i"))
            .map(|v| v == "1" || v == "true");
        let layout_caps = layout_def_rpr
            .and_then(|rp| attr(&rp, "cap"))
            .filter(|v| v == "all" || v == "small");
        let layout_color: Option<String> = layout_def_rpr
            .and_then(|rp| child(rp, "solidFill"))
            .and_then(|sf| parse_color_node(sf, theme));
        let layout_alignment: Option<String> = layout_lvl1_ppr
            .and_then(|lp| attr(&lp, "algn"))
            .map(|a| a.to_string());
        // ECMA-376 §21.1.2.2.7 eaLnBrk from the layout placeholder's lvl1pPr.
        let layout_ea_ln_brk: Option<bool> = layout_lvl1_ppr
            .and_then(|lp| attr(&lp, "eaLnBrk"))
            .map(|v| v == "1" || v == "true");
        let layout_space_before: Option<i64> = layout_lvl1_ppr
            .and_then(|lp| child(lp, "spcBef"))
            .and_then(|s| child(s, "spcPts"))
            .and_then(|s| attr_i64(&s, "val"));
        let layout_space_after: Option<i64> = layout_lvl1_ppr
            .and_then(|lp| child(lp, "spcAft"))
            .and_then(|s| child(s, "spcPts"))
            .and_then(|s| attr_i64(&s, "val"));
        // lnSpc > spcPct val (e.g. 90000 = 90%)
        let layout_line_spacing: Option<f64> = layout_lvl1_ppr
            .and_then(|lp| child(lp, "lnSpc"))
            .and_then(|ls| child(ls, "spcPct"))
            .and_then(|s| attr_f64(&s, "val"));

        // Layout bodyPr anchor; fall back to master anchor map
        let layout_anchor: Option<String> = child(sp, "txBody")
            .and_then(|tb| child(tb, "bodyPr"))
            .and_then(|bp| attr(&bp, "anchor"))
            .map(|a| a.to_string());

        // Layout spPr > ln stroke (real visible border, not edit-mode indicator when solidFill is present)
        let layout_stroke: Option<Stroke> = child(sp_pr, "ln").and_then(|n| parse_stroke(n, theme));

        // Layout spPr fill (solidFill / noFill / gradFill / pattFill). The
        // slide-level placeholder shape inherits this when its own `<p:spPr>` is
        // empty — that's how a "tinted body placeholder" carries through to the
        // slide. We deliberately exclude grpFill here (group inheritance is
        // resolved at slide parse time, not from the layout).
        let layout_fill: Option<Fill> = parse_fill(sp_pr, theme);

        // Layout spPr > blipFill → image that bleeds through when the slide's
        // matching placeholder has no own blipFill (picture placeholder inheritance).
        let layout_blip_fill: Option<InheritedBlipFill> = child(sp_pr, "blipFill").and_then(|bf| {
            let rid = child(bf, "blip").and_then(|b| attr_r(&b, "embed"))?;
            let rel_target = layout_rels.get(&rid)?;
            let image_path = resolve_path(layout_dir, rel_target);
            // Verify the part exists so a dangling rId yields None (no inherited
            // fill), preserving the prior data-URL behaviour. `index_for_name`
            // reads the central directory only (no inflate), unlike the former
            // `read_zip_bytes` which decompressed the entry just to discard it.
            zip.index_for_name(&image_path)?;
            let mime_type = mime_from_ext(&image_path).to_owned();
            Some(InheritedBlipFill {
                image_path,
                mime_type,
                src_rect: parse_src_rect(bf),
                alpha: parse_blip_alpha(bf),
            })
        });

        if let Some(ph) = ph_node {
            let ph_type = attr(&ph, "type").unwrap_or_default();
            let ph_idx: Option<u32> = attr(&ph, "idx").and_then(|v| v.parse().ok());

            if let Some(idx) = ph_idx {
                if let Some(ref t) = t_opt {
                    lph.by_idx.entry(idx).or_insert_with(|| t.clone());
                }
                // Prefer layout font size; fall back to master
                let fs = layout_font_size.or_else(|| master_font_sizes.get(&ph_type).copied());
                if let Some(fs) = fs {
                    lph.by_idx_font_size.entry(idx).or_insert(fs);
                }
                // Per-level: layout lstStyle wins per level, else master.
                let level_sizes = merge_level_sizes(
                    &layout_level_sizes,
                    master_level_font_sizes.get(&ph_type).unwrap_or(&[None; 9]),
                );
                if has_any_level_size(&level_sizes) {
                    lph.by_idx_level_sizes.entry(idx).or_insert(level_sizes);
                }
                // Per-level indents: layout lstStyle wins per axis/level, else master.
                let level_indents = merge_level_indents(
                    &layout_level_indents,
                    master_level_indents
                        .get(&ph_type)
                        .unwrap_or(&Default::default()),
                );
                if has_any_level_indent(&level_indents) {
                    lph.by_idx_level_indents.entry(idx).or_insert(level_indents);
                }
                // Per-level bullets: layout lstStyle wins per level, else master.
                let empty_bul = empty_level_bullets();
                let level_bullets = merge_level_bullets(
                    &layout_level_bullets,
                    master_level_bullets.get(&ph_type).unwrap_or(&empty_bul),
                );
                if has_any_level_bullet(&level_bullets) {
                    lph.by_idx_level_bullets.entry(idx).or_insert(level_bullets);
                }
                if let Some(ref s) = layout_stroke {
                    lph.by_idx_stroke.entry(idx).or_insert(s.clone());
                }
                if let Some(ls) = layout_line_spacing {
                    lph.by_idx_line_spacing.entry(idx).or_insert(ls);
                }
                if let Some(ref bf) = layout_blip_fill {
                    lph.by_idx_blip_fill.entry(idx).or_insert(bf.clone());
                }
                if let Some(ref c) = layout_color {
                    lph.by_idx_color.entry(idx).or_insert(c.clone());
                }
                if let Some(ref f) = layout_fill {
                    lph.by_idx_fill.entry(idx).or_insert(f.clone());
                }
                // Alignment for this idx: layout's own algn, else master per-type
                // (incl. master txStyles, now folded into master_alignments).
                let idx_algn = layout_alignment
                    .clone()
                    .or_else(|| master_alignments.get(&ph_type).cloned());
                if let Some(a) = idx_algn {
                    lph.by_idx_alignment.entry(idx).or_insert(a);
                }
            }
            let effective_fs =
                layout_font_size.or_else(|| master_font_sizes.get(&ph_type).copied());
            if let Some(fs) = effective_fs {
                lph.by_type_font_size.entry(ph_type.clone()).or_insert(fs);
            }
            let type_level_sizes = merge_level_sizes(
                &layout_level_sizes,
                master_level_font_sizes.get(&ph_type).unwrap_or(&[None; 9]),
            );
            if has_any_level_size(&type_level_sizes) {
                lph.by_type_level_sizes
                    .entry(ph_type.clone())
                    .or_insert(type_level_sizes);
            }
            let type_level_indents = merge_level_indents(
                &layout_level_indents,
                master_level_indents
                    .get(&ph_type)
                    .unwrap_or(&Default::default()),
            );
            if has_any_level_indent(&type_level_indents) {
                lph.by_type_level_indents
                    .entry(ph_type.clone())
                    .or_insert(type_level_indents);
            }
            let empty_bul_t = empty_level_bullets();
            let type_level_bullets = merge_level_bullets(
                &layout_level_bullets,
                master_level_bullets.get(&ph_type).unwrap_or(&empty_bul_t),
            );
            if has_any_level_bullet(&type_level_bullets) {
                lph.by_type_level_bullets
                    .entry(ph_type.clone())
                    .or_insert(type_level_bullets);
            }
            if let Some(b) = layout_bold {
                lph.by_type_bold.entry(ph_type.clone()).or_insert(b);
            }
            if let Some(i) = layout_italic {
                lph.by_type_italic.entry(ph_type.clone()).or_insert(i);
            }
            if let Some(c) = layout_caps.clone() {
                lph.by_type_caps.entry(ph_type.clone()).or_insert(c);
            }
            if let Some(a) = layout_alignment {
                lph.by_type_alignment.entry(ph_type.clone()).or_insert(a);
            }
            if let Some(e) = layout_ea_ln_brk {
                lph.by_type_ea_ln_brk.entry(ph_type.clone()).or_insert(e);
            }
            if let Some(v) = layout_space_before {
                lph.by_type_space_before.entry(ph_type.clone()).or_insert(v);
            }
            if let Some(v) = layout_space_after {
                lph.by_type_space_after.entry(ph_type.clone()).or_insert(v);
            }
            if let Some(ls) = layout_line_spacing {
                lph.by_type_line_spacing
                    .entry(ph_type.clone())
                    .or_insert(ls);
            }
            // Anchor: layout bodyPr > fall back to master anchor map
            let effective_anchor = layout_anchor
                .clone()
                .or_else(|| master_anchors.get(&ph_type).cloned());
            if let Some(a) = effective_anchor {
                lph.by_type_anchor.entry(ph_type.clone()).or_insert(a);
            }
            if let Some(s) = layout_stroke {
                lph.by_type_stroke.entry(ph_type.clone()).or_insert(s);
            }
            if let Some(bf) = layout_blip_fill {
                lph.by_type_blip_fill.entry(ph_type.clone()).or_insert(bf);
            }
            if let Some(c) = layout_color {
                lph.by_type_color.entry(ph_type.clone()).or_insert(c);
            }
            if let Some(f) = layout_fill {
                lph.by_type_fill.entry(ph_type.clone()).or_insert(f);
            }
            if let Some(t) = t_opt {
                lph.by_type.entry(ph_type).or_insert(t);
            }
        }
    }
    lph
}

/// The layout XML parsed ONCE into the owned data a slide needs from its layout
/// (D4). Groups the three former per-slide layout re-parses in `parse_slide`:
/// placeholder inheritance (§19.3.1.39), the layout-level `<p:bg>` background,
/// and the layout's `showMasterSp` flag (§19.3.1.39). Holds no `roxmltree` node
/// (owned only), so it can be cached across slides sharing a layout.
///
/// The color-bearing fields (`placeholders` colors/fills/strokes/bullets +
/// `background`) are resolved against the `theme` passed to `parse_layout`. For
/// the common no-`clrMapOvr` slide that theme is the master's baked theme, so
/// the cached instance is reused; a slide with a `<p:clrMapOvr>` builds a fresh
/// `ParsedLayout` against its override theme (see the `parse_presentation` loop)
/// so its layout colors flip too. The layout's DECORATIVE spTree shapes are NOT
/// held here — they are walked per-slide because they resolve against the slide's
/// own `smartart_drawings` (§19.3.1.39 layout decorations) and are theme+zip
/// bound; caching them keyed by layout would be unsound.
struct ParsedLayout {
    placeholders: LayoutPlaceholders,
    /// Layout-level `<p:cSld><p:bg>` fill (ECMA-376 §19.3.1.1 / §20.1.8.14),
    /// resolved against `theme`. Applied by the slide only when its own bg chain
    /// (slide-level) resolves to nothing.
    background: Option<Fill>,
    /// The LAYOUT's own `showMasterSp` (§19.3.1.39). The slide ANDs this with its
    /// own slide-level flag before compositing master decorations.
    show_master_sp: bool,
}

impl Default for ParsedLayout {
    fn default() -> Self {
        // Matches the prior "no/unparseable layout" behaviour: no placeholders,
        // no layout background, and showMasterSp defaulting to true.
        ParsedLayout {
            placeholders: LayoutPlaceholders::default(),
            background: None,
            show_master_sp: true,
        }
    }
}

/// ECMA-376 §19.3.1.38/§19.3.1.39 showMasterSp: absent / "1" / "true" ⇒ true;
/// "0" / "false" ⇒ false. Read from a slide or layout root element.
fn read_show_master_sp(node: roxmltree::Node<'_, '_>) -> bool {
    match attr(&node, "showMasterSp").as_deref() {
        Some("0") | Some("false") => false,
        _ => true, // default true (absent / "1" / "true")
    }
}

/// Parse a slide layout's XML EXACTLY ONCE and extract everything a slide
/// inherits from it (D4). Replaces the four former per-slide layout
/// `Document::parse` calls in `parse_slide` (placeholders, background,
/// showMasterSp, decorations) — the decorations still walk per-slide, but from
/// the SAME `Document` when the caller reuses it, and the other three are cached.
/// `theme` is the slide's effective theme (master-baked, or override-adjusted);
/// the master maps are the inheritance fallbacks, threaded through unchanged.
#[allow(clippy::too_many_arguments)]
fn parse_layout(
    layout_xml: &str,
    master_font_sizes: &HashMap<String, f64>,
    master_level_font_sizes: &HashMap<String, LevelFontSizes>,
    master_level_indents: &HashMap<String, LevelIndents>,
    master_level_bullets: &HashMap<String, LevelBullets>,
    master_anchors: &HashMap<String, String>,
    master_transforms: &HashMap<String, Transform>,
    master_alignments: &HashMap<String, String>,
    master_ea_ln_brk: &HashMap<String, bool>,
    master_space_before: &HashMap<String, i64>,
    master_space_after: &HashMap<String, i64>,
    master_line_spacing: &HashMap<String, f64>,
    theme: &HashMap<String, String>,
    layout_dir: &str,
    layout_rels: &HashMap<String, String>,
    zip: &mut PptxZip,
) -> ParsedLayout {
    note_layout_master_parse();
    let doc = match roxmltree::Document::parse(layout_xml) {
        Ok(d) => d,
        // Unparseable layout → same as no layout: default placeholders/bg and
        // showMasterSp = true (the slide's own flag still applies downstream).
        Err(_) => return ParsedLayout::default(),
    };
    let root = doc.root_element();

    let placeholders = parse_layout_placeholders(
        root,
        master_font_sizes,
        master_level_font_sizes,
        master_level_indents,
        master_level_bullets,
        master_anchors,
        master_transforms,
        master_alignments,
        master_ea_ln_brk,
        master_space_before,
        master_space_after,
        master_line_spacing,
        theme,
        layout_dir,
        layout_rels,
        zip,
    );

    // Layout-level bg (rels = layout rels, part dir = layout_dir). Verbatim from
    // the former inline layout-bg block in `parse_slide`; the slide decides
    // whether to use it (only when its own bg chain is empty).
    let background: Option<Fill> = child(root, "cSld").and_then(|n| {
        let mut resolve = |rid: &str| -> Option<String> {
            let target = layout_rels.get(rid)?;
            let path = resolve_path(layout_dir, target);
            // Existence check only — central-directory lookup, no inflate.
            zip.index_for_name(&path)?;
            Some(path)
        };
        parse_background(n, theme, &mut resolve)
    });

    let show_master_sp = read_show_master_sp(root);

    ParsedLayout {
        placeholders,
        background,
        show_master_sp,
    }
}

// ===========================
//  Text body parsing
// ===========================

/// Which `<a:objectDefaults>` slot to consult when the shape's own bodyPr
/// leaves an attribute unset. `Tx` ⇔ "text box" (slide-level
/// `<p:cNvSpPr txBox="1"/>`), which inherits from `<a:txDef>`. `Sp` ⇔
/// "regular shape with text" — table cell, placeholder, or a
/// preset-geometry shape carrying a `<p:txBody>` — which inherits from
/// `<a:spDef>`. Falling back to txDef for non-text-boxes is wrong because
/// txDef commonly carries `<a:spAutoFit/>` (PowerPoint's default for
/// freshly-inserted text boxes); applying that to e.g. a placeholder body
/// makes the whole paragraph spill horizontally instead of wrapping.
#[derive(Clone, Copy, PartialEq, Eq)]
enum ShapeKind {
    Tx,
    Sp,
}

// Carries the resolved master/layout/placeholder inheritance context (theme,
// rels, inherited font size, default alignment/spacing, level styles) that text
// runs need; these are inheritance inputs, not an arbitrary parameter bag.
#[allow(clippy::too_many_arguments)]
fn parse_text_body(
    tx_body: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
    inherited_font_size: Option<f64>,
    inherited_level_font_sizes: LevelFontSizes,
    inherited_level_indents: LevelIndents,
    inherited_level_bullets: &LevelBullets,
    inherited_bold: Option<bool>,
    inherited_italic: Option<bool>,
    inherited_caps: Option<String>,
    inherited_anchor: Option<String>,
    inherited_alignment: Option<String>,
    inherited_ea_ln_brk: Option<bool>,
    inherited_space_before: Option<i64>,
    inherited_space_after: Option<i64>,
    inherited_line_spacing: Option<f64>,
    shape_kind: ShapeKind,
    zip: &mut PptxZip,
) -> TextBody {
    let body_pr = child(tx_body, "bodyPr");
    // ECMA-376 §20.1.6.7 objectDefaults. The theme-level `<a:txDef>` (and
    // `<a:spDef>` as a secondary fallback) provides defaults for every
    // bodyPr attribute the slide-level shape leaves unset. Without this
    // fallback chain, sample-2's "20代" (txDef carries `<a:spAutoFit/>`)
    // and similar templates would silently use the spec's literal defaults
    // instead of what the theme author intended.
    // Shape-kind-aware lookup: text boxes consult txDef, regular shapes spDef.
    // Cross-fall is intentionally NOT done — see ShapeKind doc.
    let def_prefix = match shape_kind {
        ShapeKind::Tx => "+txDef",
        ShapeKind::Sp => "+spDef",
    };
    let theme_default_str =
        |key: &str| -> Option<String> { theme.get(&format!("{def_prefix}-bodyPr-{key}")).cloned() };
    let theme_default_i64 =
        |key: &str| -> Option<i64> { theme_default_str(key).and_then(|v| v.parse::<i64>().ok()) };
    let theme_default_u32 =
        |key: &str| -> Option<u32> { theme_default_str(key).and_then(|v| v.parse::<u32>().ok()) };
    let theme_auto_fit =
        || -> Option<String> { theme.get(&format!("{def_prefix}-autoFit")).cloned() };

    // Shared `<a:bodyPr>` grammar (anchor / wrap / vert / insets / autofit) via
    // ooxml_common::text::parse_body_pr. pptx's inheritance + theme
    // objectDefaults resolution is pre-baked into the defaults: each field is
    // `inherited?.or(theme objectDefault)?.or(spec default)`, and parse_body_pr
    // then applies the shape's own bodyPr attribute over it — so the effective
    // precedence (shape attr → inherited → theme → spec) is unchanged. When the
    // shape has no `<a:bodyPr>` at all, the resolved defaults ARE the result.
    //
    // Insets: OOXML defaults lIns=rIns=91440, tIns=bIns=45720 (the shared
    // ooxml_common::text::DEFAULT_INS_* constants, via BodyPrDefaults::spec()).
    // Autofit child (spAutoFit / normAutofit): when absent, defer to theme txDef
    // (auto_fit default below); a normAutofit also captures PowerPoint's stored
    // fontScale / lnSpcReduction (ECMA-376 §21.1.2.1.3, 62500 → 0.625).
    let spec = ooxml_common::text::BodyPrDefaults::spec();
    let body_pr_defaults = ooxml_common::text::BodyPrDefaults {
        anchor: inherited_anchor
            .or_else(|| theme_default_str("anchor"))
            .unwrap_or(spec.anchor),
        wrap: theme_default_str("wrap").unwrap_or(spec.wrap),
        vert: theme_default_str("vert").unwrap_or(spec.vert),
        l_ins: theme_default_i64("lIns").unwrap_or(spec.l_ins),
        t_ins: theme_default_i64("tIns").unwrap_or(spec.t_ins),
        r_ins: theme_default_i64("rIns").unwrap_or(spec.r_ins),
        b_ins: theme_default_i64("bIns").unwrap_or(spec.b_ins),
        auto_fit: theme_auto_fit().unwrap_or(spec.auto_fit),
    };
    let body = match body_pr {
        Some(n) => ooxml_common::text::parse_body_pr(n, &body_pr_defaults),
        // No <a:bodyPr>: every field resolves to its default.
        None => ooxml_common::text::BodyPr {
            anchor: body_pr_defaults.anchor.clone(),
            wrap: body_pr_defaults.wrap.clone(),
            vert: body_pr_defaults.vert.clone(),
            l_ins: body_pr_defaults.l_ins,
            t_ins: body_pr_defaults.t_ins,
            r_ins: body_pr_defaults.r_ins,
            b_ins: body_pr_defaults.b_ins,
            auto_fit: body_pr_defaults.auto_fit.clone(),
            font_scale: None,
            ln_spc_reduction: None,
        },
    };
    let vertical_anchor = body.anchor;
    let l_ins = body.l_ins;
    let r_ins = body.r_ins;
    let t_ins = body.t_ins;
    let b_ins = body.b_ins;
    let wrap = body.wrap;
    let vert = body.vert;
    let auto_fit = body.auto_fit;
    let font_scale = body.font_scale;
    let ln_spc_reduction = body.ln_spc_reduction;
    // ECMA-376 §20.1.10.34: numCol on <a:bodyPr> tells the renderer to
    // distribute paragraphs across N columns within the shape. Default 1.
    // spcCol is the inter-column gutter in EMU (default 0). Both fall back
    // through theme objectDefaults.
    let num_col = body_pr
        .and_then(|n| attr(&n, "numCol"))
        .and_then(|v| v.parse::<u32>().ok())
        .or_else(|| theme_default_u32("numCol"))
        .filter(|&n| n >= 1)
        .unwrap_or(1);
    let spc_col = body_pr
        .and_then(|n| attr_i64(&n, "spcCol"))
        .or_else(|| theme_default_i64("spcCol"))
        .unwrap_or(0);
    // ECMA-376 §21.1.2.1.1: rtlCol on <a:bodyPr> lays out the text body's
    // columns right-to-left. xsd:boolean, so accept "1"/"true". Shape
    // attribute → theme objectDefaults → spec default (false).
    let rtl_col = body_pr
        .and_then(|n| attr(&n, "rtlCol"))
        .or_else(|| theme_default_str("rtlCol"))
        .map(|v| v == "1" || v == "true")
        .unwrap_or(false);

    // Own lstStyle > lvl1pPr, then fall back to layout/master inherited values
    let own_lvl1_ppr = child(tx_body, "lstStyle").and_then(|ls| child(ls, "lvl1pPr"));
    let own_def_rpr = own_lvl1_ppr.and_then(|lp| child(lp, "defRPr"));
    let default_font_size = own_def_rpr
        .and_then(|rp| attr_f64(&rp, "sz"))
        .map(|v| v / 100.0)
        .or(inherited_font_size);
    // Effective per-list-level default sizes: this shape's own lstStyle wins per
    // level, else the layout/master inherited per-level sizes. Paragraphs pick
    // their size by `lvl` so nested bullets shrink (ECMA-376 §21.1.2.4).
    let own_level_sizes = extract_level_font_sizes(tx_body);
    let effective_level_sizes = merge_level_sizes(&own_level_sizes, &inherited_level_font_sizes);
    // Effective per-list-level indents: this shape's own lstStyle wins per
    // axis/level, else the layout/master inherited per-level indents. A paragraph
    // that omits marL/marR/indent picks them by `lvl` from this cascade before
    // falling back to PowerPoint's hardcoded implicit defaults (§21.1.2.4.13).
    let own_level_indents = extract_level_indents(tx_body);
    let effective_level_indents = merge_level_indents(&own_level_indents, &inherited_level_indents);
    // Effective per-level bullets: own lstStyle wins per level, else inherited
    // layout/master. A paragraph with no explicit bullet resolves its marker (and
    // its hanging-indent defaults) from this by `lvl` (ECMA-376 §19.7.10).
    // A slide text body's own lstStyle `<a:buBlip>` resolves against the slide's
    // rels + part directory (ECMA-376 §21.1.2.4.2), same base as the slide's
    // picture fills. `parse_text_body` is only ever called for slide shapes, so
    // the part directory is always `ppt/slides`.
    let mut resolve_slide_blip = |rid: &str| -> Option<String> {
        let target = rels.get(rid)?;
        let path = resolve_path("ppt/slides", target);
        // Verify the part exists so a listed-but-missing rId yields None and the
        // bullet falls through to Bullet::Inherit (matches the variant's doc
        // comment), mirroring the slide picture-fill resolvers. `index_for_name`
        // reads the central directory only (no inflate), unlike the former
        // `read_zip_bytes` which decompressed the entry just to discard it.
        zip.index_for_name(&path)?;
        Some(path)
    };
    let own_level_bullets = extract_level_bullets(tx_body, theme, &mut resolve_slide_blip);
    let effective_level_bullets = merge_level_bullets(&own_level_bullets, inherited_level_bullets);
    let default_bold = own_def_rpr
        .and_then(|rp| attr(&rp, "b"))
        .map(|v| v == "1" || v == "true")
        .or(inherited_bold);
    let default_italic = own_def_rpr
        .and_then(|rp| attr(&rp, "i"))
        .map(|v| v == "1" || v == "true")
        .or(inherited_italic);
    // Own lstStyle > lvl1pPr > algn overrides inherited alignment
    let body_default_alignment = own_lvl1_ppr
        .and_then(|lp| attr(&lp, "algn"))
        .map(|a| a.to_string())
        .or(inherited_alignment);

    // Own lstStyle > lvl1pPr > eaLnBrk overrides inherited (ECMA-376 §21.1.2.2.7)
    let body_default_ea_ln_brk = own_lvl1_ppr
        .and_then(|lp| attr(&lp, "eaLnBrk"))
        .map(|v| v == "1" || v == "true")
        .or(inherited_ea_ln_brk);

    // Own lstStyle > lvl1pPr spacing overrides inherited
    let own_lvl1_spcbef: Option<i64> = own_lvl1_ppr
        .and_then(|lp| child(lp, "spcBef"))
        .and_then(|s| child(s, "spcPts"))
        .and_then(|s| attr_i64(&s, "val"));
    let own_lvl1_spcaft: Option<i64> = own_lvl1_ppr
        .and_then(|lp| child(lp, "spcAft"))
        .and_then(|s| child(s, "spcPts"))
        .and_then(|s| attr_i64(&s, "val"));
    let body_default_space_before = own_lvl1_spcbef.or(inherited_space_before);
    let body_default_space_after = own_lvl1_spcaft.or(inherited_space_after);

    // Own lstStyle > lvl1pPr > lnSpc overrides inherited line spacing
    let own_lvl1_line_spacing: Option<f64> = own_lvl1_ppr
        .and_then(|lp| child(lp, "lnSpc"))
        .and_then(|ls| child(ls, "spcPct"))
        .and_then(|s| attr_f64(&s, "val"));
    let body_default_line_spacing = own_lvl1_line_spacing.or(inherited_line_spacing);

    let mut paragraphs: Vec<Paragraph> = children_vec(tx_body, "p")
        .into_iter()
        .map(|p| {
            parse_paragraph(
                p,
                theme,
                rels,
                body_default_alignment.as_deref(),
                body_default_ea_ln_brk,
                body_default_space_before,
                body_default_space_after,
                body_default_line_spacing,
                &effective_level_sizes,
                &effective_level_indents,
                &effective_level_bullets,
                zip,
            )
        })
        .collect();

    // ECMA-376 §21.1.2.3.13 cap: a run inherits cap="all"/"small" from the
    // shape's own lstStyle defRPr, else from the layout/master placeholder
    // style (e.g. a template's titleStyle cap="all" upper-cases the title even
    // though the run's text is stored mixed-case). Run-level rPr/paragraph
    // defRPr already won via parse_run; fill the remainder here.
    let body_caps = own_def_rpr
        .and_then(|rp| attr(&rp, "cap"))
        .filter(|v| v == "all" || v == "small")
        .or(inherited_caps);
    if let Some(bc) = body_caps {
        for para in &mut paragraphs {
            for run in &mut para.runs {
                if let TextRun::Text(t) = run {
                    if t.caps.is_none() {
                        t.caps = Some(bc.clone());
                    }
                }
            }
        }
    }

    TextBody {
        vertical_anchor,
        paragraphs,
        default_font_size,
        default_bold,
        default_italic,
        l_ins,
        r_ins,
        t_ins,
        b_ins,
        wrap,
        vert,
        auto_fit,
        font_scale,
        ln_spc_reduction,
        num_col,
        spc_col,
        rtl_col,
    }
}

/// Walk `node` for OMML math and push a `TextRun::Math` for each equation,
/// descending PowerPoint's `mc:AlternateContent` / `mc:Choice` / `a14:m`
/// wrappers (ECMA-376 §22.1; the a14 markup is from the 2010 drawing ext).
/// Find the font size (pt) of an equation from the first run property within it
/// that carries `sz`. PowerPoint puts the size on the math run's `a:rPr` (or
/// `m:rPr`) rather than the paragraph defRPr, so inline math matches the
/// surrounding text size. `sz` is in hundredths of a point (ECMA-376 §21.1.2.3.9).
fn math_run_size(om: roxmltree::Node<'_, '_>) -> Option<f64> {
    om.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "rPr")
        .find_map(|rpr| attr_f64(&rpr, "sz"))
        .map(|v| v / 100.0)
}

/// Equation colour: the first run-property solidFill within the equation
/// (PowerPoint puts the colour on the math run's `a:rPr`, like the size), so
/// inline math follows the surrounding text colour (e.g. a purple title).
fn math_run_color(om: roxmltree::Node<'_, '_>, theme: &HashMap<String, String>) -> Option<String> {
    om.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "rPr")
        .find_map(|rpr| child(rpr, "solidFill").and_then(|sf| parse_color_node(sf, theme)))
}

fn push_math_runs(
    node: roxmltree::Node<'_, '_>,
    font_size: Option<f64>,
    theme: &HashMap<String, String>,
    runs: &mut Vec<TextRun>,
) {
    match node.tag_name().name() {
        "oMath" => {
            let nodes = parse_omath_nodes(node);
            if !nodes.is_empty() {
                runs.push(TextRun::Math {
                    nodes,
                    display: false,
                    font_size: math_run_size(node).or(font_size),
                    color: math_run_color(node, theme),
                });
            }
        }
        "oMathPara" => {
            for om in node
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "oMath")
            {
                let nodes = parse_omath_nodes(om);
                if !nodes.is_empty() {
                    runs.push(TextRun::Math {
                        nodes,
                        display: true,
                        font_size: math_run_size(om).or(font_size),
                        color: math_run_color(om, theme),
                    });
                }
            }
        }
        // mc:AlternateContent → take the a14 `Choice` (the live equation) and
        // ignore mc:Fallback (a rasterized picture PowerPoint emits for old apps).
        "AlternateContent" => {
            if let Some(choice) = node
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "Choice")
            {
                for c in choice.children().filter(|n| n.is_element()) {
                    push_math_runs(c, font_size, theme, runs);
                }
            }
        }
        // a14:m wrapper (local name "m") holds an m:oMathPara.
        "m" => {
            for c in node.children().filter(|n| n.is_element()) {
                push_math_runs(c, font_size, theme, runs);
            }
        }
        _ => {}
    }
}

// Same inherited paragraph/run context as parse_text_body, scoped to one <a:p>.
#[allow(clippy::too_many_arguments)]
fn parse_paragraph(
    p_node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
    body_default_alignment: Option<&str>,
    body_default_ea_ln_brk: Option<bool>,
    body_default_space_before: Option<i64>,
    body_default_space_after: Option<i64>,
    body_default_line_spacing: Option<f64>,
    level_font_sizes: &LevelFontSizes,
    level_indents: &LevelIndents,
    level_bullets: &LevelBullets,
    zip: &mut PptxZip,
) -> Paragraph {
    let p_pr = child(p_node, "pPr");

    // ECMA-376 §21.1.2.2.7 `<a:pPr rtl>` — right-to-left text flow. When set
    // and the paragraph has no explicit `algn`, the implicit default flips
    // from "l" to "r" (matches PowerPoint's behaviour for Arabic / Hebrew
    // slides where users typically don't author an explicit alignment).
    let rtl = p_pr
        .and_then(|n| attr(&n, "rtl"))
        .map(|v| v == "1" || v == "true")
        .unwrap_or(false);

    // ECMA-376 §21.1.2.2.7 `<a:pPr eaLnBrk>` (xsd:boolean). Paragraph's own
    // value → body/list-style → layout/master default → spec default (true).
    // Same inheritance shape as `alignment` above.
    let ea_ln_brk = p_pr
        .and_then(|n| attr(&n, "eaLnBrk"))
        .map(|v| v == "1" || v == "true")
        .or(body_default_ea_ln_brk)
        .unwrap_or(true);

    // Paragraph's own algn → body/layout/master default → "r" if rtl, else "l"
    let alignment = p_pr
        .and_then(|n| attr(&n, "algn"))
        .map(|a| a.to_string())
        .or_else(|| body_default_alignment.map(|a| a.to_string()))
        .unwrap_or_else(|| if rtl { "r".into() } else { "l".into() });
    let lvl: u32 = p_pr
        .and_then(|n| attr(&n, "lvl"))
        .and_then(|v| v.parse().ok())
        .unwrap_or(0);

    // Effective bullet: the paragraph's own
    // `<a:buChar>`/`<a:buAutoNum>`/`<a:buBlip>`/`<a:buNone>`, else the inherited
    // per-level bullet for this placeholder (ECMA-376 §19.7.10). A paragraph's
    // own `<a:buBlip>` embed resolves against the slide rels + `ppt/slides`,
    // the same base as the slide's picture fills (§21.1.2.4.2).
    let mut resolve_para_blip = |rid: &str| -> Option<String> {
        let target = rels.get(rid)?;
        let path = resolve_path("ppt/slides", target);
        // Verify the part exists so a listed-but-missing rId yields None and the
        // bullet falls through to Bullet::Inherit (matches the variant's doc
        // comment), mirroring the slide picture-fill resolvers. `index_for_name`
        // reads the central directory only (no inflate), unlike the former
        // `read_zip_bytes` which decompressed the entry just to discard it.
        zip.index_for_name(&path)?;
        Some(path)
    };
    let bullet = match parse_bullet(p_pr, theme, &mut resolve_para_blip) {
        Bullet::Inherit => level_bullets
            .get(lvl as usize)
            .and_then(|o| o.clone())
            .unwrap_or(Bullet::Inherit),
        b => b,
    };
    // A paragraph is a list item (and gets a hanging indent) when its effective
    // bullet is a char/number/picture — whether declared explicitly or
    // inherited. An inherited bullet without an inherited marL/indent reuses
    // PowerPoint's implicit list metrics, the same defaults explicit bullets
    // already use.
    let has_bullet = matches!(
        bullet,
        Bullet::Char { .. } | Bullet::AutoNum { .. } | Bullet::Blip { .. }
    );

    // marL / marR / indent resolve per axis: direct `<a:pPr>` attribute wins,
    // else the authored list-style level cascade (`level_indents`, from the
    // shape/layout/master lstStyle per ECMA-376 §21.1.2.4.13), else PowerPoint's
    // hardcoded implicit list defaults:
    //   Bullet paragraphs:  marL = (lvl+1)*342900, indent = -342900 (hanging)
    //   Plain paragraphs:   marL = lvl*457200 (matches presentation.xml defaultTextStyle)
    let level_indent = level_indents.get(lvl as usize).copied().unwrap_or_default();
    let mar_l = p_pr
        .and_then(|n| attr_i64(&n, "marL"))
        .or(level_indent.mar_l)
        .unwrap_or_else(|| {
            if has_bullet {
                (lvl as i64 + 1) * 342900
            } else {
                lvl as i64 * 457200
            }
        });
    let mar_r = p_pr
        .and_then(|n| attr_i64(&n, "marR"))
        .or(level_indent.mar_r)
        .unwrap_or(0);
    let indent = p_pr
        .and_then(|n| attr_i64(&n, "indent"))
        .or(level_indent.indent)
        .unwrap_or(if has_bullet { -342900 } else { 0 });

    let space_before = p_pr
        .and_then(|n| {
            child(n, "spcBef")
                .and_then(|s| child(s, "spcPts"))
                .and_then(|s| attr_i64(&s, "val"))
        })
        .or(body_default_space_before);
    let space_after = p_pr
        .and_then(|n| {
            child(n, "spcAft")
                .and_then(|s| child(s, "spcPts"))
                .and_then(|s| attr_i64(&s, "val"))
        })
        .or(body_default_space_after);

    let space_line = p_pr
        .and_then(|n| child(n, "lnSpc"))
        .and_then(parse_lnspc)
        .or_else(|| body_default_line_spacing.map(|v| SpaceLine::Pct { val: v }));

    // Tab stops from pPr > tabLst
    let tab_stops: Vec<TabStop> = p_pr
        .and_then(|n| child(n, "tabLst"))
        .map(|tab_lst| {
            tab_lst
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "tab")
                .filter_map(|tab| {
                    let pos = attr_i64(&tab, "pos")?;
                    let algn = attr(&tab, "algn").unwrap_or_else(|| "l".into());
                    Some(TabStop { pos, algn })
                })
                .collect()
        })
        .unwrap_or_default();

    // Paragraph-level default run properties (pPr > defRPr)
    let def_rpr = p_pr.and_then(|n| child(n, "defRPr"));
    let def_font_size = def_rpr.and_then(|n| attr_f64(&n, "sz")).map(|v| v / 100.0);
    let def_color = def_rpr
        .and_then(|n| child(n, "solidFill"))
        .and_then(|n| parse_color_node(n, theme));
    let def_bold = def_rpr
        .and_then(|n| attr(&n, "b"))
        .map(|v| v == "1" || v == "true");
    let def_italic = def_rpr
        .and_then(|n| attr(&n, "i"))
        .map(|v| v == "1" || v == "true");
    let def_font_family = def_rpr
        .and_then(|n| child(n, "latin"))
        .and_then(|n| attr(&n, "typeface"))
        .map(|tf| resolve_theme_typeface(&tf, theme));

    let mut runs = Vec::new();
    for node in p_node.children().filter(|n| n.is_element()) {
        match node.tag_name().name() {
            "r" => {
                if let Some(run) = parse_run(node, def_rpr, theme, rels) {
                    runs.push(TextRun::Text(run));
                }
            }
            "br" => runs.push(TextRun::Break),
            // OMML equations (ECMA-376 §22.1). `def_font_size` here is the
            // paragraph's defRPr size (pre-level-fallback); the renderer applies
            // its own inheritance when this is None. PowerPoint stores inline
            // math as a bare `a14:m` (local name "m") directly under `a:p`, and
            // also inside `mc:AlternateContent`; both reach push_math_runs.
            "oMath" | "oMathPara" | "AlternateContent" | "m" => {
                push_math_runs(node, def_font_size, theme, &mut runs);
            }
            // Field elements (e.g. slide number, date): parse like a run but tag the field type
            "fld" => {
                let fld_type = attr(&node, "type").unwrap_or_default().to_string();
                let text = child(node, "t")
                    .and_then(|t| t.text())
                    .unwrap_or("")
                    .to_string();
                let r_pr = child(node, "rPr");
                let font_size = r_pr.and_then(|n| attr_f64(&n, "sz")).map(|v| v / 100.0);
                let color = r_pr
                    .and_then(|n| child(n, "solidFill"))
                    .and_then(|n| parse_color_node(n, theme));
                let bold = r_pr
                    .and_then(|n| attr(&n, "b"))
                    .map(|v| v == "1" || v == "true");
                let italic = r_pr
                    .and_then(|n| attr(&n, "i"))
                    .map(|v| v == "1" || v == "true");
                let font_family = r_pr
                    .and_then(|n| child(n, "latin"))
                    .and_then(|n| attr(&n, "typeface"))
                    .map(|tf| resolve_theme_typeface(&tf, theme));
                // §21.1.2.3.4 — a field's rPr can also carry a highlight; resolve
                // it the same way as a normal run (CT_Color via the shared path).
                let highlight = r_pr
                    .and_then(|n| child(n, "highlight"))
                    .and_then(|n| parse_color_node(n, theme));
                runs.push(TextRun::Text(TextRunData {
                    text,
                    bold,
                    italic,
                    underline: false,
                    underline_style: None,
                    underline_color: None,
                    strikethrough: false,
                    strike_double: false,
                    font_size,
                    color,
                    font_family,
                    font_family_ea: None,
                    font_family_sym: None,
                    baseline: None,
                    caps: None,
                    letter_spacing: None,
                    field_type: if fld_type == "slidenum" {
                        Some("slidenum".to_string())
                    } else {
                        None
                    },
                    hyperlink: None,
                    shadow: None,
                    outline: None,
                    highlight,
                }));
            }
            _ => {}
        }
    }

    // For paragraphs with no visible text content, use endParaRPr sz to set line height.
    // This ensures empty spacer paragraphs have the correct height (e.g. between sections).
    let end_rpr = child(p_node, "endParaRPr");
    let has_text = runs.iter().any(|r| matches!(r, TextRun::Text(_)));
    let def_font_size = def_font_size
        .or_else(|| {
            if !has_text {
                end_rpr.and_then(|n| attr_f64(&n, "sz")).map(|v| v / 100.0)
            } else {
                None
            }
        })
        // Inherited per-list-level default size, indexed by this paragraph's
        // level (ECMA-376 §21.1.2.4): a 2nd-level bullet uses lvl3pPr's smaller
        // defRPr sz, not the level-1 size. The renderer applies `def_font_size`
        // to runs that carry no explicit `sz`.
        .or_else(|| level_font_sizes.get(lvl as usize).copied().flatten());

    Paragraph {
        alignment,
        mar_l,
        mar_r,
        indent,
        space_before,
        space_after,
        space_line,
        lvl,
        bullet,
        def_font_size,
        def_color,
        def_bold,
        def_italic,
        def_font_family,
        tab_stops,
        rtl,
        ea_ln_brk,
        runs,
    }
}

/// Parse bullet specification from pPr node.
///
/// `resolve_blip` maps a `<a:buBlip><a:blip r:embed>` rId to the bullet image's
/// embedded **zip path** (ECMA-376 §21.1.2.4.2), using the rels + part directory
/// of whichever tier this `pPr` belongs to (slide paragraph / txBody lstStyle /
/// layout / master), mirroring how `parse_blip_fill` resolves image fills. A
/// `buBlip` whose embed can't be resolved (dangling rId) falls through to the
/// `Bullet::Inherit` default so a lower style tier can still supply a marker.
fn parse_bullet<F: FnMut(&str) -> Option<String>>(
    p_pr: Option<roxmltree::Node<'_, '_>>,
    theme: &HashMap<String, String>,
    resolve_blip: &mut F,
) -> Bullet {
    let p_pr = match p_pr {
        Some(n) => n,
        None => return Bullet::Inherit,
    };

    // Explicit "no bullet"
    if child(p_pr, "buNone").is_some() {
        return Bullet::None;
    }

    // §21.1.2.4.3 buSzPct (val in thousandths of a percent: 100000 = 100%).
    // Shared by char and picture bullets — read once.
    let size_pct = child(p_pr, "buSzPct")
        .and_then(|n| attr_f64(&n, "val"))
        .map(|v| v / 1000.0);

    // Picture bullet (buBlip) — ECMA-376 §21.1.2.4.2. The choice is mutually
    // exclusive with buChar/buAutoNum, so resolve it before the char/number
    // branches. Only emit a Blip when the embed resolves to a real part.
    if let Some(bu_blip) = child(p_pr, "buBlip") {
        if let Some(image_path) = child(bu_blip, "blip")
            .and_then(|b| attr_r(&b, "embed"))
            .and_then(|rid| resolve_blip(&rid))
        {
            let mime_type = mime_from_ext(&image_path).to_owned();
            return Bullet::Blip {
                image_path,
                mime_type,
                size_pct,
            };
        }
    }

    // Character bullet
    if let Some(bu_char) = child(p_pr, "buChar") {
        let ch = attr(&bu_char, "char").unwrap_or_else(|| "\u{2022}".into()); // •
        let color = child(p_pr, "buClr").and_then(|n| parse_color_node(n, theme));
        let font_family = child(p_pr, "buFont")
            .and_then(|n| attr(&n, "typeface"))
            .map(|tf| resolve_theme_typeface(&tf, theme));
        return Bullet::Char {
            ch,
            color,
            size_pct,
            font_family,
        };
    }

    // Auto-numbered bullet
    if let Some(bu_auto) = child(p_pr, "buAutoNum") {
        let num_type = attr(&bu_auto, "type").unwrap_or_else(|| "arabicPeriod".into());
        let start_at = attr(&bu_auto, "startAt").and_then(|v| v.parse().ok());
        return Bullet::AutoNum { num_type, start_at };
    }

    Bullet::Inherit
}

fn parse_run(
    r_node: roxmltree::Node<'_, '_>,
    def_rpr: Option<roxmltree::Node<'_, '_>>,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
) -> Option<TextRunData> {
    let t_node = child(r_node, "t")?;
    let text = t_node.text().unwrap_or("").to_owned();
    let r_pr = child(r_node, "rPr");

    // Attribute with rPr → defRPr fallback; None means "not set" (inherit from body/layout defaults)
    let bold = r_pr
        .and_then(|n| attr(&n, "b"))
        .or_else(|| def_rpr.and_then(|n| attr(&n, "b")))
        .map(|v| v == "1" || v == "true");
    let italic = r_pr
        .and_then(|n| attr(&n, "i"))
        .or_else(|| def_rpr.and_then(|n| attr(&n, "i")))
        .map(|v| v == "1" || v == "true");
    // ECMA-376 §21.1.2.3.16 — underline style enum: none/sng/dbl/heavy/dotted/
    // dash/dashLong/dotDash/dotDotDash/wavy plus *Heavy variants. Carry the
    // exact value through for the renderer to dispatch on; the bool stays
    // true for any non-"none" value so existing code paths keep working.
    let underline_attr = r_pr
        .and_then(|n| attr(&n, "u"))
        .or_else(|| def_rpr.and_then(|n| attr(&n, "u")));
    let underline = underline_attr
        .as_deref()
        .map(|v| v != "none")
        .unwrap_or(false);
    let underline_style = underline_attr.filter(|v| v != "none" && v != "sng");

    // ECMA-376 §21.1.2.3.20 — uFill specifies a per-underline colour that
    // overrides the text colour. uFillTx (or absence) means "follow text".
    let underline_color = r_pr
        .and_then(|n| child(n, "uFill"))
        .or_else(|| def_rpr.and_then(|n| child(n, "uFill")))
        .and_then(|n| child(n, "solidFill"))
        .and_then(|n| parse_color_node(n, theme));

    // strikethrough: "sngStrike" or "dblStrike" → true; double tracked separately
    let strike_attr = r_pr
        .and_then(|n| attr(&n, "strike"))
        .or_else(|| def_rpr.and_then(|n| attr(&n, "strike")));
    let strikethrough = strike_attr
        .as_deref()
        .map(|v| v == "sngStrike" || v == "dblStrike")
        .unwrap_or(false);
    let strike_double = strike_attr.as_deref() == Some("dblStrike");

    // ECMA-376 §21.1.2.3.13 ST_TextCapsType: "none" | "small" | "all". Treat
    // "none" as not set (no transform) so the field stays absent in JSON.
    let caps = r_pr
        .and_then(|n| attr(&n, "cap"))
        .or_else(|| def_rpr.and_then(|n| attr(&n, "cap")))
        .filter(|v| v == "small" || v == "all");

    // ECMA-376 §21.1.2.3.5 rPr @spc — letter spacing in 100ths of a point.
    // Negative values are valid (tightening). Non-zero only.
    let letter_spacing = r_pr
        .and_then(|n| attr_f64(&n, "spc"))
        .or_else(|| def_rpr.and_then(|n| attr_f64(&n, "spc")))
        .map(|v| v / 100.0)
        .filter(|v| v.abs() > f64::EPSILON);

    // sz in hundredths of a point
    let font_size = r_pr
        .and_then(|n| attr_f64(&n, "sz"))
        .or_else(|| def_rpr.and_then(|n| attr_f64(&n, "sz")))
        .map(|v| v / 100.0);

    let color = r_pr
        .and_then(|n| child(n, "solidFill"))
        .and_then(|n| parse_color_node(n, theme))
        .or_else(|| {
            def_rpr
                .and_then(|n| child(n, "solidFill"))
                .and_then(|n| parse_color_node(n, theme))
        });

    let font_family = r_pr
        .and_then(|n| child(n, "latin"))
        .and_then(|n| attr(&n, "typeface"))
        .or_else(|| {
            def_rpr
                .and_then(|n| child(n, "latin"))
                .and_then(|n| attr(&n, "typeface"))
        })
        .map(|tf| resolve_theme_typeface(&tf, theme));
    // ECMA-376 §21.1.2.3.7 — <a:ea typeface="..."/> sets a separate font for
    // East Asian glyphs (CJK). Defaults to the theme's +mn-ea slot when the
    // run doesn't specify one explicitly.
    let font_family_ea = r_pr
        .and_then(|n| child(n, "ea"))
        .and_then(|n| attr(&n, "typeface"))
        .or_else(|| {
            def_rpr
                .and_then(|n| child(n, "ea"))
                .and_then(|n| attr(&n, "typeface"))
        })
        .map(|tf| resolve_theme_typeface(&tf, theme))
        .filter(|tf| !tf.is_empty());

    // ECMA-376 §21.1.2.3.10 — <a:sym typeface="..."/> sets the font used for
    // symbol characters. PowerPoint stores those as PUA codepoints (U+F0xx).
    let font_family_sym = r_pr
        .and_then(|n| child(n, "sym"))
        .and_then(|n| attr(&n, "typeface"))
        .or_else(|| {
            def_rpr
                .and_then(|n| child(n, "sym"))
                .and_then(|n| attr(&n, "typeface"))
        })
        .map(|tf| resolve_theme_typeface(&tf, theme))
        .filter(|tf| !tf.is_empty());

    // baseline in thousandths of a point; 30000=superscript, -25000=subscript (OOXML typical)
    let baseline = r_pr
        .and_then(|n| attr(&n, "baseline"))
        .and_then(|v| v.parse::<i32>().ok())
        .filter(|&v| v != 0);

    // a:hlinkClick — hyperlink. r:id refers to the slide rels (Target = URL or local target).
    // Resolve immediately so the renderer doesn't need access to the rels table.
    let hyperlink = r_pr
        .and_then(|n| child(n, "hlinkClick"))
        .and_then(|h| attr_r(&h, "id"))
        .and_then(|rid| rels.get(&rid).cloned())
        .filter(|s| !s.is_empty());

    // ECMA-376 §20.1.8.45 — `<a:rPr><a:effectLst><a:outerShdw>` glyph drop
    // shadow. Reuse the shape-level outerShdw reader so parse semantics
    // stay identical (blurRad, dist, dir, color + alphaModFix).
    let shadow = r_pr
        .and_then(|n| child(n, "effectLst"))
        .and_then(|el| parse_shadow(el, theme));

    // ECMA-376 §20.1.2.2.24 (CT_TextOutlineEffect) — `<a:rPr><a:ln w="..">`
    // strokes each glyph outline. `<a:noFill>` inside the ln means "no
    // visible outline" — skip in that case so the renderer doesn't draw a
    // black box around every glyph. Pull color from solidFill if present.
    let outline = r_pr
        .and_then(|n| child(n, "ln"))
        .filter(|ln| child(*ln, "noFill").is_none())
        .map(|ln| TextOutline {
            width: attr_i64(&ln, "w").unwrap_or(0),
            color: child(ln, "solidFill").and_then(|n| parse_color_node(n, theme)),
        });

    // ECMA-376 §21.1.2.3.4 — `<a:rPr><a:highlight>` text highlight (marker).
    // The element IS a CT_Color, so pass the <a:highlight> node straight to the
    // shared colour resolver — the same one solidFill uses — which walks its
    // srgbClr / schemeClr / sysClr / prstClr child and applies any tint / alpha
    // transforms. schemeClr therefore resolves through the master clrMap +
    // theme exactly like other run colours. Falls back to defRPr when the run
    // itself doesn't set a highlight.
    let highlight = r_pr
        .and_then(|n| child(n, "highlight"))
        .and_then(|n| parse_color_node(n, theme))
        .or_else(|| {
            def_rpr
                .and_then(|n| child(n, "highlight"))
                .and_then(|n| parse_color_node(n, theme))
        });

    Some(TextRunData {
        text,
        bold,
        italic,
        underline,
        underline_style,
        underline_color,
        strikethrough,
        strike_double,
        font_size,
        color,
        font_family,
        font_family_ea,
        font_family_sym,
        baseline,
        caps,
        letter_spacing,
        field_type: None,
        hyperlink,
        shadow,
        outline,
        highlight,
    })
}

// ===========================
//  Chart parsing
// ===========================

// ===========================
//  Placeholder defaults
// ===========================

/// OOXML spec default positions for common placeholder types.
/// Values are in EMU, assuming a 9144000×6858000 slide (10"×7.5").
fn default_placeholder_transform(ph_type: &str) -> Transform {
    match ph_type {
        "title" | "ctrTitle" => Transform {
            x: 457200,
            y: 274638,
            cx: 8229600,
            cy: 1143000,
            rot: 0.0,
            flip_h: false,
            flip_v: false,
        },
        "subTitle" => Transform {
            x: 457200,
            y: 1600200,
            cx: 8229600,
            cy: 899160,
            rot: 0.0,
            flip_h: false,
            flip_v: false,
        },
        "dt" => Transform {
            x: 0,
            y: 6261600,
            cx: 2286000,
            cy: 596900,
            rot: 0.0,
            flip_h: false,
            flip_v: false,
        },
        "ftr" => Transform {
            x: 2972400,
            y: 6261600,
            cx: 3086100,
            cy: 596900,
            rot: 0.0,
            flip_h: false,
            flip_v: false,
        },
        "sldNum" => Transform {
            x: 6629400,
            y: 6261600,
            cx: 2057400,
            cy: 596900,
            rot: 0.0,
            flip_h: false,
            flip_v: false,
        },
        // "body" and everything else: full-width content area below title
        _ => Transform {
            x: 457200,
            y: 1600200,
            cx: 8229600,
            cy: 4525963,
            rot: 0.0,
            flip_h: false,
            flip_v: false,
        },
    }
}

// ===========================
//  Placeholder detection
// ===========================

/// Returns true if the node contains a `p:ph` descendant.
fn is_placeholder(node: roxmltree::Node<'_, '_>) -> bool {
    node.descendants()
        .any(|n| n.is_element() && n.tag_name().name() == "ph")
}

// ===========================
//  Shape parsing  (p:sp)
// ===========================

/// Pull `(id, name)` from a sibling `<p:nvSpPr><p:cNvPr>` (or `<p:nvCxnSpPr>`).
/// Returns `(None, None)` when the wrapper is missing — both fields are
/// optional in the JSON output.
fn read_cnv_pr(sp_node: roxmltree::Node<'_, '_>) -> (Option<String>, Option<String>) {
    for wrapper_name in &["nvSpPr", "nvCxnSpPr", "nvPicPr", "nvGraphicFramePr"] {
        if let Some(nv) = child(sp_node, wrapper_name) {
            if let Some(cnv) = child(nv, "cNvPr") {
                let id = attr(&cnv, "id");
                let name = attr(&cnv, "name").filter(|s| !s.is_empty());
                return (id, name);
            }
        }
    }
    (None, None)
}

fn parse_shape(
    sp_node: roxmltree::Node<'_, '_>,
    lph: &LayoutPlaceholders,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
    group_fill: Option<&Fill>,
    zip: &mut PptxZip,
) -> Option<ShapeElement> {
    // --- Placeholder info (for layout fallback) ---
    let ph_node = sp_node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "ph");
    let ph_type = ph_node
        .as_ref()
        .and_then(|n| attr(n, "type"))
        .unwrap_or_else(|| "body".into());
    let ph_idx: Option<u32> = ph_node
        .as_ref()
        .and_then(|n| attr(n, "idx"))
        .and_then(|v| v.parse().ok());
    // Surface the explicit ph @type / @idx (when present) on the ShapeElement
    // JSON. We keep `ph_type` defaulted to "body" for the internal lookup
    // path, but only emit the `placeholder_type` field when a real `<p:ph>`
    // node was found.
    let placeholder_type_out: Option<String> = ph_node
        .as_ref()
        .map(|n| attr(n, "type").unwrap_or_else(|| "body".into()));
    let (cnv_id, cnv_name) = read_cnv_pr(sp_node);

    // --- Transform: slide xfrm OR layout fallback ---
    let sp_pr = child(sp_node, "spPr");
    let slide_xfrm = sp_pr.and_then(|p| child(p, "xfrm"));

    let t: Transform = if let Some(xfrm) = slide_xfrm {
        parse_xfrm(xfrm)
    } else if ph_node.is_some() {
        match lph.lookup(&ph_type, ph_idx) {
            Some(lt) => lt.clone(),
            None => default_placeholder_transform(&ph_type),
        }
    } else {
        return None; // non-placeholder with no xfrm — skip
    };

    // cx=0 AND cy=0 → skip (zero-area invisible shape).
    // cx=0 alone (vertical line/annotation) is permitted when cy > 0.
    // cy=0 means "auto-height": keep 0 when anchor="b" (renderer grows shape upward from off_y),
    // otherwise use a generous fallback so text has room to render.
    if t.cx == 0 && t.cy == 0 {
        return None;
    }
    let inherited_anchor: Option<String> = if ph_node.is_some() {
        lph.lookup_anchor(&ph_type)
    } else {
        None
    };
    let is_bottom_anchor = inherited_anchor.as_deref() == Some("b")
        || child(sp_node, "txBody")
            .and_then(|tb| child(tb, "bodyPr"))
            .and_then(|bp| attr(&bp, "anchor"))
            .map(|a| a == "b")
            .unwrap_or(false);
    // custGeom takes priority over prstGeom
    let cust_geom_node = sp_pr.and_then(|p| child(p, "custGeom"));
    let prst_geom_node = sp_pr.and_then(|p| child(p, "prstGeom"));
    let geometry = if cust_geom_node.is_some() {
        "custGeom".into()
    } else {
        prst_geom_node
            .and_then(|n| attr(&n, "prst"))
            .unwrap_or_else(|| "rect".into())
    };

    // cy=0 means "auto-height" for body-text shapes, but connector-type
    // geometries (line, *Connector*) legitimately use cy=0 to represent
    // a perfectly horizontal segment — don't inflate their height.
    let is_connector_geom = matches!(
        geometry.as_str(),
        "line"
            | "straightConnector1"
            | "bentConnector2"
            | "bentConnector3"
            | "bentConnector4"
            | "bentConnector5"
            | "curvedConnector2"
            | "curvedConnector3"
            | "curvedConnector4"
            | "curvedConnector5"
    );
    let cy = if t.cy == 0 && !is_connector_geom {
        if is_bottom_anchor {
            0_i64
        } else {
            2_000_000_i64
        }
    } else {
        t.cy
    };
    let cust_geom = cust_geom_node.map(|n| parse_cust_geom(n));

    // Parse adjustment values from prstGeom avLst (e.g. trapezoid inset)
    // Collect all gd elements; first is adj (name="adj" or "adj1"), second is adj2
    let parse_gd_val = |gd: roxmltree::Node<'_, '_>| -> Option<f64> {
        attr(&gd, "fmla")
            .and_then(|f| f.strip_prefix("val ").map(|s| s.to_owned()))
            .and_then(|s| s.parse::<f64>().ok())
    };
    let av_node = prst_geom_node.and_then(|n| child(n, "avLst"));
    let gd_nodes: Vec<_> = av_node
        .map(|av| {
            av.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "gd")
                .collect()
        })
        .unwrap_or_default();
    // First gd = adj (match by name "adj" or "adj1", fallback to position 0)
    let adj: Option<f64> = gd_nodes
        .iter()
        .find(|n| matches!(attr(n, "name").as_deref(), Some("adj") | Some("adj1")))
        .or_else(|| gd_nodes.first())
        .and_then(|n| parse_gd_val(*n));
    // Second gd = adj2 (match by name "adj2", fallback to position 1)
    let adj2: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj2"))
        .or_else(|| gd_nodes.get(1))
        .and_then(|n| parse_gd_val(*n));
    // Third gd = adj3 (match by name "adj3", fallback to position 2)
    let adj3: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj3"))
        .or_else(|| gd_nodes.get(2))
        .and_then(|n| parse_gd_val(*n));
    // Fourth gd = adj4 (match by name "adj4", fallback to position 3)
    let adj4: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj4"))
        .or_else(|| gd_nodes.get(3))
        .and_then(|n| parse_gd_val(*n));
    // adj5-adj8 for callouts that specify extra polyline vertices
    // (accentBorderCallout3 etc.).
    let adj5: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj5"))
        .or_else(|| gd_nodes.get(4))
        .and_then(|n| parse_gd_val(*n));
    let adj6: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj6"))
        .or_else(|| gd_nodes.get(5))
        .and_then(|n| parse_gd_val(*n));
    let adj7: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj7"))
        .or_else(|| gd_nodes.get(6))
        .and_then(|n| parse_gd_val(*n));
    let adj8: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj8"))
        .or_else(|| gd_nodes.get(7))
        .and_then(|n| parse_gd_val(*n));

    // --- Shape style (p:style) provides fill/stroke/text-color fallbacks ---
    let style_node = child(sp_node, "style");

    // fillRef idx=0 → explicit no-fill; idx>0 → use referenced color as solid fill
    let style_fill: Option<Fill> = style_node.and_then(|s| child(s, "fillRef")).and_then(|fr| {
        let idx: u32 = attr(&fr, "idx").and_then(|v| v.parse().ok()).unwrap_or(1);
        if idx == 0 {
            Some(Fill::None)
        } else {
            parse_color_node(fr, theme).map(|c| Fill::Solid { color: c })
        }
    });

    // lnRef idx=0 → no line; idx>0 → resolve width from theme's fmtScheme >
    // lnStyleLst (stored as "+lnRef-N") and color from the ref's own solidFill.
    // Falling back to 9525 under-weights idx>=2 strokes (Office default theme
    // idx=2 is 19050 EMU = 1.5pt, idx=3 is 25400 EMU = 2pt).
    let style_stroke: Option<Stroke> = style_node.and_then(|s| child(s, "lnRef")).and_then(|lr| {
        let idx: u32 = attr(&lr, "idx").and_then(|v| v.parse().ok()).unwrap_or(1);
        if idx == 0 {
            None
        } else {
            parse_color_node(lr, theme).map(|c| {
                let width = theme
                    .get(&format!("+lnRef-{}", idx))
                    .and_then(|s| s.parse::<i64>().ok())
                    .unwrap_or(9525);
                Stroke {
                    color: c,
                    width,
                    dash_style: None,
                    head_end: None,
                    tail_end: None,
                    cmpd: None,
                }
            })
        }
    });

    // fontRef → default text color for this shape.
    // Fall back to layout/master placeholder inherited color (lstStyle > lvl1pPr > defRPr
    // > solidFill) when the shape is a placeholder and has no explicit p:style > fontRef.
    let default_text_color: Option<String> = style_node
        .and_then(|s| child(s, "fontRef"))
        .and_then(|fr| parse_color_node(fr, theme))
        .or_else(|| {
            if ph_node.is_some() {
                lph.lookup_color(&ph_type, ph_idx)
            } else {
                None
            }
        });

    // spPr fill resolution order (ECMA-376 §19.3.1.36 / §20.1.4.2):
    //   1. spPr `<a:grpFill>` → inherit from parent group
    //   2. spPr explicit fill (`<a:solidFill>`, `<a:noFill>`, gradient, pattern, blip)
    //   3. layout placeholder fill (only when the slide-level shape has no fill
    //      element of its own and is bound to a placeholder)
    //   4. `<p:style><a:fillRef>` falls through last
    // Some(Fill::None) (noFill in spPr) is treated as an explicit choice and
    // must NOT be overridden by step 3 or 4.
    let sp_pr_has_grp_fill = sp_pr.and_then(|p| child(p, "grpFill")).is_some();
    let fill = if sp_pr_has_grp_fill {
        group_fill.cloned()
    } else {
        let own = sp_pr.and_then(|p| parse_fill(p, theme));
        let inherited = if own.is_none() && ph_node.is_some() {
            lph.lookup_fill(&ph_type, ph_idx)
        } else {
            None
        };
        own.or(inherited).or(style_fill)
    };

    // spPr stroke: if ln element is present, respect it (even if noFill → None);
    // otherwise fall back to layout placeholder stroke, then style stroke.
    let stroke = if sp_pr.and_then(|p| child(p, "ln")).is_some() {
        sp_pr
            .and_then(|p| child(p, "ln"))
            .and_then(|n| parse_stroke(n, theme))
    } else if ph_node.is_some() {
        lph.lookup_stroke(&ph_type, ph_idx).or(style_stroke)
    } else {
        style_stroke
    };

    // Inherited defaults from layout/master for this placeholder type/idx
    let (
        inherited_font_size,
        inherited_bold,
        inherited_italic,
        inherited_caps,
        inherited_anchor,
        inherited_alignment,
        inherited_ea_ln_brk,
        inherited_space_before,
        inherited_space_after,
        inherited_line_spacing,
    ) = if ph_node.is_some() {
        (
            lph.lookup_font_size(&ph_type, ph_idx),
            lph.lookup_bold(&ph_type),
            lph.lookup_italic(&ph_type),
            lph.lookup_caps(&ph_type),
            lph.lookup_anchor(&ph_type),
            lph.lookup_alignment(&ph_type, ph_idx),
            lph.lookup_ea_ln_brk(&ph_type),
            lph.lookup_space_before(&ph_type),
            lph.lookup_space_after(&ph_type),
            lph.lookup_line_spacing(&ph_type, ph_idx),
        )
    } else {
        (None, None, None, None, None, None, None, None, None, None)
    };
    let inherited_level_font_sizes: LevelFontSizes = if ph_node.is_some() {
        lph.lookup_level_font_sizes(&ph_type, ph_idx)
    } else {
        [None; 9]
    };
    // Per-level paragraph indents (marL/marR/indent) a paragraph inherits when it
    // omits them (ECMA-376 §21.1.2.4.13): the layout/master placeholder cascade.
    let inherited_level_indents: LevelIndents = if ph_node.is_some() {
        lph.lookup_level_indents(&ph_type, ph_idx)
    } else {
        Default::default()
    };

    // ECMA-376 §19.3.1.21 / §20.1.4.2: a slide-level `<p:cNvSpPr txBox="1"/>`
    // marks the shape as a true text box, which means the theme's
    // `<a:txDef>` (rather than `<a:spDef>`) provides the fallback bodyPr.
    let is_text_box = child(sp_node, "nvSpPr")
        .and_then(|n| child(n, "cNvSpPr"))
        .and_then(|n| attr(&n, "txBox"))
        .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
        .unwrap_or(false);
    let shape_kind = if is_text_box {
        ShapeKind::Tx
    } else {
        ShapeKind::Sp
    };
    // Per-level bullets a paragraph inherits when it declares no explicit one
    // (ECMA-376 §19.7.10): the layout/master placeholder cascade for this slot.
    let inherited_level_bullets: LevelBullets = if ph_node.is_some() {
        lph.lookup_level_bullets(&ph_type, ph_idx)
    } else {
        empty_level_bullets()
    };
    let text_body = child(sp_node, "txBody").map(|n| {
        parse_text_body(
            n,
            theme,
            rels,
            inherited_font_size,
            inherited_level_font_sizes,
            inherited_level_indents,
            &inherited_level_bullets,
            inherited_bold,
            inherited_italic,
            inherited_caps.clone(),
            inherited_anchor,
            inherited_alignment,
            inherited_ea_ln_brk,
            inherited_space_before,
            inherited_space_after,
            inherited_line_spacing,
            shape_kind,
            zip,
        )
    });

    // Effects from spPr > effectLst (outerShdw / innerShdw / glow / softEdge /
    // reflection are independent siblings — ECMA-376 §20.1.8.16). Pull each.
    let EffectLst {
        shadow,
        inner_shadow,
        glow,
        soft_edge,
        reflection,
    } = parse_effect_lst(sp_pr.and_then(|p| child(p, "effectLst")), theme);

    Some(ShapeElement {
        x: t.x,
        y: t.y,
        width: t.cx,
        height: cy,
        rotation: t.rot,
        flip_h: t.flip_h,
        flip_v: t.flip_v,
        geometry,
        fill,
        stroke,
        text_body,
        default_text_color,
        cust_geom,
        adj,
        adj2,
        adj3,
        adj4,
        adj5,
        adj6,
        adj7,
        adj8,
        shadow,
        inner_shadow,
        glow,
        soft_edge,
        reflection,
        id: cnv_id,
        name: cnv_name,
        placeholder_type: placeholder_type_out,
        placeholder_idx: ph_idx,
        text_rect: None,
        scene3d: sp_pr.and_then(parse_scene3d),
        sp3d: sp_pr.and_then(parse_sp3d),
    })
}

// ===========================
//  Picture parsing  (p:pic)
// ===========================

/// Preset clip geometry for a picture body (§20.1.9.18 — the preset geometry of
/// a `<p:pic>` acts as its clip silhouette and the path its border / contour
/// hug). Returns the prstGeom `prst` name plus every `<a:avLst><a:gd>` adjust
/// value in declaration order (1/1000-of-a-percent OOXML units), so any of the
/// 186 presets — roundRect, ellipse, and the rest — can be reconstructed by the
/// shared preset-geometry engine on the TS side. The preset's own declared
/// defaults fill in any omitted guide, so the `adj` Vec may be shorter than the
/// preset's adjust count (the engine substitutes defaults per index).
///
/// `prst="rect"` returns None: a plain rectangle needs no clip path (the bitmap
/// already fills the bbox), matching the previous behaviour.
fn parse_pic_prst_geom(sp_pr: roxmltree::Node<'_, '_>) -> (Option<String>, Option<Vec<i64>>) {
    let pg = match child(sp_pr, "prstGeom") {
        Some(n) => n,
        None => return (None, None),
    };
    let prst = match attr(&pg, "prst") {
        Some(p) if p != "rect" => p,
        _ => return (None, None),
    };
    let adjust: Vec<i64> = child(pg, "avLst")
        .map(|av| {
            av.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "gd")
                .filter_map(|gd| {
                    attr(&gd, "fmla")
                        .and_then(|f| f.strip_prefix("val ").and_then(|v| v.parse::<i64>().ok()))
                })
                .collect()
        })
        .unwrap_or_default();
    let adjust = if adjust.is_empty() {
        None
    } else {
        Some(adjust)
    };
    (Some(prst), adjust)
}

/// Microsoft 2016 SVG extension on an `<a:blip>` (ECMA-376 §20.1.8.14 carries
/// the core blip fill; the SVG body is a Microsoft extension). When a `<a:blip>`
/// has `<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">` with an
/// `<asvg:svgBlip r:embed="…">`, resolve that rId to the `.svg` part and return
/// it as a `data:image/svg+xml;base64,…` URL. Returns None when there is no
/// svgBlip, the rId is unresolvable, or the part is missing.
///
/// Matching is by namespace-local element name (`svgBlip`), so the `asvg:`
/// prefix (or any other) is irrelevant.
/// Resolve the Microsoft 2016 `asvg:svgBlip` extension on a `<a:blip>` to the
/// embedded **zip path** of the `.svg` part (e.g. "ppt/media/image2.svg"). The
/// renderer fetches the bytes lazily by path. The part's existence is verified
/// so a dangling rId yields None (no SVG twin), matching the prior behaviour.
fn svg_blip_path(
    blip: roxmltree::Node<'_, '_>,
    slide_dir: &str,
    rels: &HashMap<String, String>,
    zip: &mut PptxZip,
) -> Option<String> {
    let svg_rid = svg_blip_rid(blip)?;
    let svg_target = rels.get(&svg_rid)?;
    let svg_path = resolve_path(slide_dir, svg_target);
    // Existence check only — `index_for_name` reads the central directory
    // without inflating, unlike the former `read_zip_bytes` which decompressed
    // the whole SVG part just to discard the bytes.
    zip.index_for_name(&svg_path)?;
    Some(svg_path)
}

fn parse_picture(
    pic_node: roxmltree::Node<'_, '_>,
    slide_dir: &str,
    rels: &HashMap<String, String>,
    theme: &HashMap<String, String>,
    zip: &mut PptxZip,
) -> Option<PictureElement> {
    let sp_pr = child(pic_node, "spPr")?;
    let xfrm_node = child(sp_pr, "xfrm")?;
    let t = parse_xfrm(xfrm_node);

    if t.cx == 0 || t.cy == 0 {
        return None; // pictures always need explicit dimensions
    }

    let blip_fill = child(pic_node, "blipFill")?;
    let blip = child(blip_fill, "blip")?;

    // Microsoft 2016 SVG extension: the real vector image rides inside
    // `<a:blip><a:extLst>`, while `<a:blip r:embed>` carries a raster (PNG/JPEG)
    // *fallback* for SVG-incapable clients. Resolve the SVG first so a picture
    // that carries ONLY the svgBlip — with no raster `r:embed`, e.g. an icon
    // inserted as a pure SVG — still parses instead of being dropped. Surfaced
    // separately so the renderer can prefer the vector original and fall back to
    // the raster on decode error.
    let svg_image_path = svg_blip_path(blip, slide_dir, rels, zip);

    // Raster blip (`<a:blip r:embed>` → zip path + intrinsic PNG size). Optional:
    // a pure-SVG picture omits it, so this is no longer a hard requirement for
    // the element.
    let raster: Option<(String, Option<(u32, u32)>)> = (|| {
        let r_id = attr_r(&blip, "embed")?;
        let rel_target = rels.get(&r_id)?;
        let path = resolve_path(slide_dir, rel_target);
        let image_bytes = ooxml_common::zip::read_zip_bytes(zip, &path).ok()?;
        // Intrinsic PNG size for the ink-fallback centering (None for non-PNG,
        // unchanged from the former png_size_from_data_url semantics).
        let size = png_size_from_bytes(&image_bytes);
        Some((path, size))
    })();

    // A picture needs at least one drawable source. Prefer the raster as
    // `image_path` (the renderer's srcRect / SVG-decode-failure fallback path);
    // when no raster is embedded, fall back to the SVG part itself so the
    // element is always drawable. `mime_type` matches whichever source wins.
    let (image_path, mime_type, intrinsic_width_px, intrinsic_height_px) =
        match (raster, svg_image_path.as_ref()) {
            (Some((path, size)), _) => {
                let mime = mime_from_ext(&path).to_owned();
                let (w, h) = match size {
                    Some((w, h)) => (Some(w), Some(h)),
                    None => (None, None),
                };
                (path, mime, w, h)
            }
            (None, Some(svg)) => (svg.clone(), "image/svg+xml".to_owned(), None, None),
            (None, None) => return None,
        };

    // ECMA-376 §20.1.9.8 — `<p:pic>` may carry `<a:custGeom>` inside `<p:spPr>`,
    // in which case the bitmap is clipped to that custom path (e.g. a laptop
    // silhouette). Re-use the same parser as for shapes so the renderer can
    // build a Path2D and `ctx.clip()` before drawing the image.
    let cust_geom = child(sp_pr, "custGeom").map(parse_cust_geom);

    // §19.3.1.37: p:pic's spPr is CT_ShapeProperties, so effectLst (§20.1.8.16)
    // applies to images exactly as it does to shapes.
    let EffectLst {
        shadow,
        inner_shadow,
        glow,
        soft_edge,
        reflection,
    } = parse_effect_lst(child(sp_pr, "effectLst"), theme);

    // §20.1.2.2.24 — a `p:pic`'s spPr may carry an `<a:ln>` border (e.g. the
    // white frame of PowerPoint's picture styles). `parse_stroke` returns None
    // for `<a:noFill/>`, so an explicitly border-less picture stays None.
    let stroke = child(sp_pr, "ln").and_then(|n| parse_stroke(n, theme));

    let (prst_geom, prst_adjust) = parse_pic_prst_geom(sp_pr);
    Some(PictureElement {
        x: t.x,
        y: t.y,
        width: t.cx,
        height: t.cy,
        rotation: t.rot,
        flip_h: t.flip_h,
        flip_v: t.flip_v,
        image_path,
        mime_type,
        svg_image_path,
        intrinsic_width_px,
        intrinsic_height_px,
        stroke,
        prst_geom,
        prst_adjust,
        src_rect: parse_src_rect(blip_fill),
        alpha: parse_blip_alpha(blip_fill),
        cust_geom,
        shadow,
        inner_shadow,
        glow,
        soft_edge,
        reflection,
        scene3d: parse_scene3d(sp_pr),
        sp3d: parse_sp3d(sp_pr),
    })
}

/// Decode (width, height) from raw PNG bytes by reading the IHDR chunk. Returns
/// None for non-PNG payloads or malformed data (unchanged semantics from the
/// former `png_size_from_data_url`, only the input is now raw bytes instead of
/// a base64 data URL).
fn png_size_from_bytes(bytes: &[u8]) -> Option<(u32, u32)> {
    if bytes.len() < 24 || &bytes[0..8] != b"\x89PNG\r\n\x1a\n" {
        return None;
    }
    let w = u32::from_be_bytes([bytes[16], bytes[17], bytes[18], bytes[19]]);
    let h = u32::from_be_bytes([bytes[20], bytes[21], bytes[22], bytes[23]]);
    Some((w, h))
}

/// EMU per CSS pixel at PowerPoint's default 96 DPI (914400 EMU/inch ÷ 96).
const EMU_PER_PX_96DPI: i64 = 9525;

/// If a `p:pic` declares an `a:audioFile` / `a:videoFile` in its `nvPr`
/// (or the newer `p14:media` extension), emit a `MediaElement` with the
/// poster image and the media bytes. Returns None for regular pictures.
fn parse_media(
    pic_node: roxmltree::Node<'_, '_>,
    slide_dir: &str,
    rels: &HashMap<String, String>,
) -> Option<MediaElement> {
    let nv_pic_pr = child(pic_node, "nvPicPr")?;
    let nv_pr = child(nv_pic_pr, "nvPr")?;

    // A `<p:pic>` can carry several media references at once: the legacy
    // `<a:videoFile>`/`<a:audioFile>` (r:embed or r:link) and the modern
    // `<p14:media r:embed>` extension. We collect every reference and use the
    // first that actually resolves, preferring an *embedded* ref over a
    // *linked* one. Embedded refs always point at internal media we can read;
    // a linked ref may target an External resource we cannot fetch, or its rId
    // may be missing/unresolvable (caught below via `rels.get`). Trying embeds
    // first means a present `<p14:media>` embed is used even when the deck also
    // carries a videoFile link, instead of being shadowed and demoted to a
    // poster-only Picture. (Note: parse_rels stores Id+Target and never reads
    // TargetMode, so a real External link keeps its non-empty URL here.)
    // (ECMA-376 §19.3.1.17/18 a:videoFile/a:audioFile; p14:media is a Microsoft
    // extension.)
    let av = nv_pr
        .children()
        .find(|n| n.is_element() && matches!(n.tag_name().name(), "videoFile" | "audioFile"));
    let av_kind: Option<&str> = av.map(|n| {
        if n.tag_name().name() == "videoFile" {
            "video"
        } else {
            "audio"
        }
    });
    let av_embed = av.and_then(|n| attr_r(&n, "embed"));
    let av_link = av.and_then(|n| attr_r(&n, "link"));

    // `<p14:media>` lives in `nvPr > extLst > ext`.
    let p14 = child(nv_pr, "extLst").and_then(|ext_lst| {
        ext_lst
            .children()
            .filter(|c| c.is_element())
            .find_map(|ext| {
                ext.children()
                    .find(|m| m.is_element() && m.tag_name().name() == "media")
            })
    });
    let p14_embed = p14.and_then(|n| attr_r(&n, "embed"));
    let p14_link = p14.and_then(|n| attr_r(&n, "link"));

    // Preference: embedded refs (always internal) before linked refs (a link may
    // be External / unresolved); within each, the kind-bearing
    // videoFile/audioFile before the kind-less p14:media.
    let candidates = [
        (av_kind, av_embed),
        (None, p14_embed),
        (av_kind, av_link),
        (None, p14_link),
    ];
    let (media_kind, media_path, mime) = candidates.into_iter().find_map(|(kind_hint, rid)| {
        let rid = rid?;
        let target = rels.get(&rid)?;
        if target.is_empty() {
            return None; // malformed rel with an empty Target (External links carry a non-empty URL)
        }
        let path = resolve_path(slide_dir, target);
        let mime = mime_from_ext(&path);
        let kind = match kind_hint {
            Some(k) => k.to_string(),
            None if mime.starts_with("video/") => "video".to_string(),
            None if mime.starts_with("audio/") => "audio".to_string(),
            None => return None, // p14:media with an unknown MIME — try the next ref
        };
        Some((kind, path, mime))
    })?;

    // Geometry
    let sp_pr = child(pic_node, "spPr")?;
    let xfrm_node = child(sp_pr, "xfrm")?;
    let t = parse_xfrm(xfrm_node);
    if t.cx == 0 || t.cy == 0 {
        return None;
    }

    // Poster image (from blipFill). Optional — the renderer falls back to a
    // solid fill when the poster is absent or still loading. We emit just the
    // zip path so the main thread can lazily `getMedia` it; embedding large
    // posters inline ballooned the parse output for video-heavy decks.
    let (poster_path, poster_mime_type) = (|| -> Option<(String, String)> {
        let blip_fill = child(pic_node, "blipFill")?;
        let r_id = child(blip_fill, "blip").and_then(|b| attr_r(&b, "embed"))?;
        let rel_target = rels.get(&r_id)?;
        let image_path = resolve_path(slide_dir, rel_target);
        let img_mime = mime_from_ext(&image_path).to_string();
        Some((image_path, img_mime))
    })()
    .unwrap_or_default();

    Some(MediaElement {
        x: t.x,
        y: t.y,
        width: t.cx,
        height: t.cy,
        media_kind,
        poster_path,
        poster_mime_type,
        media_path,
        mime_type: mime.to_string(),
    })
}

// ===========================
//  Slide background
// ===========================

/// ECMA-376 §19.3.1.1 `p:bg`. `resolve_blip` maps a `<a:blip r:embed>` rId to a
/// base64 data URL using the rels + zip of the part this `c_sld` belongs to
/// (slide / layout / master), so an image background (§20.1.8.14) is resolved
/// against the correct relationship base.
fn parse_background<F: FnMut(&str) -> Option<String>>(
    c_sld: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    resolve_blip: &mut F,
) -> Option<Fill> {
    let bg = child(c_sld, "bg")?;
    // bgPr contains an explicit fill specification
    if let Some(bg_pr) = child(bg, "bgPr") {
        // §20.1.8.14 — an image background lives in `bgPr > blipFill`. Try it
        // first so the embedded blip is resolved; fall back to the generic
        // solid/gradient/pattern parser for non-image bgPr fills.
        if let Some(blip_fill) = child(bg_pr, "blipFill") {
            if let Some(fill) = parse_blip_fill(blip_fill, resolve_blip) {
                return Some(fill);
            }
        }
        return parse_fill(bg_pr, theme);
    }
    // bgRef references a theme background style; its child is a color element
    if let Some(bg_ref) = child(bg, "bgRef") {
        return parse_color_node(bg_ref, theme).map(|c| Fill::Solid { color: c });
    }
    None
}

// ===========================
//  Table parsing
// ===========================

/// Resolve a table-style `<a:fill>` wrapper's colour. Identical to `parse_fill`
/// for the common solid/no-fill cases, except `<a:tint>` uses the literal
/// ECMA-376 §20.1.2.3.34 formula (`TintMode::WordLiteral`) so a band's
/// `accent + tint 20%` renders as the near-white wash PowerPoint draws, rather
/// than the saturated linear-lerp used for SmartArt accents. Gradient/pattern/
/// blip fills (rare in table styles) defer to the generic `parse_fill`.
fn parse_table_style_fill(
    fill_wrapper: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<Fill> {
    use ooxml_common::color::TintMode::WordLiteral;
    for c in fill_wrapper.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "noFill" => return Some(Fill::None),
            "solidFill" => {
                return parse_color_node_tint(c, theme, WordLiteral)
                    .map(|color| Fill::Solid { color });
            }
            _ => {}
        }
    }
    parse_fill(fill_wrapper, theme)
}

/// Parse ppt/tableStyles.xml into a map of styleId → TableStyleDef.
fn parse_table_styles_xml(
    xml: &str,
    theme: &HashMap<String, String>,
) -> HashMap<String, TableStyleDef> {
    let mut map = HashMap::new();
    let Ok(doc) = roxmltree::Document::parse(xml) else {
        return map;
    };
    let root = doc.root_element();
    for style_node in root
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "tblStyle")
    {
        let Some(style_id) = attr(&style_node, "styleId") else {
            continue;
        };
        let style_id = style_id.to_string();
        let mut def = TableStyleDef::default();

        // (fill, inside-horizontal, inside-vertical, diagonal/unused, outer-horizontal,
        // outer-vertical) borders for one table-style role (ECMA-376 §20.1.4.2.27).
        type TcStyle = (
            Option<Fill>,
            Option<Stroke>,
            Option<Stroke>,
            Option<Stroke>,
            Option<Stroke>,
            Option<Stroke>,
        );
        let parse_tc_style = |role: roxmltree::Node<'_, '_>| -> TcStyle {
            let tc_style = match child(role, "tcStyle") {
                Some(n) => n,
                None => return (None, None, None, None, None, None),
            };
            // ECMA-376 §20.1.4.2.27 — the cell fill is wrapped in `<a:fill>`
            // (CT_FillProperties), so descend into it before reading the actual
            // solidFill/gradFill. Fall back to `<a:fillRef>` (theme style reference).
            let fill = child(tc_style, "fill")
                .and_then(|f| parse_table_style_fill(f, theme))
                .or_else(|| {
                    child(tc_style, "fillRef").and_then(|fr| {
                        let idx: u32 = attr(&fr, "idx").and_then(|v| v.parse().ok()).unwrap_or(0);
                        if idx == 0 {
                            Some(Fill::None)
                        } else {
                            parse_color_node(fr, theme).map(|c| Fill::Solid { color: c })
                        }
                    })
                });
            let tc_bdr = child(tc_style, "tcBdr");
            let parse_side = |side: &str| -> Option<Stroke> {
                let side_node = tc_bdr.and_then(|b| child(b, side))?;
                // Explicit <a:ln>
                if let Some(ln) = child(side_node, "ln") {
                    return parse_stroke(ln, theme);
                }
                // <a:lnRef idx="N">: use standard themed width + provided color
                if let Some(ln_ref) = child(side_node, "lnRef") {
                    let idx: u32 = attr(&ln_ref, "idx")
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(0);
                    if idx == 0 {
                        return None;
                    }
                    let color = parse_color_node(ln_ref, theme)?;
                    let width: i64 = match idx {
                        1 => 6350,
                        2 => 12700,
                        _ => 19050,
                    };
                    return Some(Stroke {
                        color,
                        width,
                        dash_style: None,
                        head_end: None,
                        tail_end: None,
                        cmpd: None,
                    });
                }
                None
            };
            let inside_h = parse_side("insideH");
            let inside_v = parse_side("insideV");
            let border_b = parse_side("bottom");
            let outer_h = parse_side("top");
            let outer_v = parse_side("left");
            (fill, inside_h, inside_v, border_b, outer_h, outer_v)
        };

        // Default text colour for a role: `<a:tcTxStyle>` holds a `<a:fontRef>`
        // followed by a colour child (schemeClr/srgbClr). parse_color_node scans
        // direct children and skips fontRef, picking up the colour.
        let parse_tc_tx_color = |role: roxmltree::Node<'_, '_>| -> Option<String> {
            child(role, "tcTxStyle").and_then(|t| parse_color_node(t, theme))
        };
        // `<a:tcTxStyle b="on">` → bold for this role (e.g. a bold header row).
        let parse_tc_tx_bold = |role: roxmltree::Node<'_, '_>| -> Option<bool> {
            child(role, "tcTxStyle")
                .and_then(|t| attr(&t, "b"))
                .map(|v| v == "on" || v == "1" || v == "true")
        };

        if let Some(whole) = child(style_node, "wholeTbl") {
            let (fill, ih, iv, _, oh, ov) = parse_tc_style(whole);
            def.whole_fill = fill;
            def.whole_inside_h = ih;
            def.whole_inside_v = iv;
            def.whole_outer_h = oh;
            def.whole_outer_v = ov;
            def.whole_text_color = parse_tc_tx_color(whole);
        }
        if let Some(band) = child(style_node, "band1H") {
            let (fill, _, _, _, _, _) = parse_tc_style(band);
            def.band1h_fill = fill;
        }
        if let Some(band) = child(style_node, "band2H") {
            let (fill, _, _, _, _, _) = parse_tc_style(band);
            def.band2h_fill = fill;
        }
        if let Some(first) = child(style_node, "firstRow") {
            let (fill, _, _, border_b, _, _) = parse_tc_style(first);
            def.first_row_fill = fill;
            def.first_row_border_b = border_b;
            def.first_row_text_color = parse_tc_tx_color(first);
            def.first_row_bold = parse_tc_tx_bold(first);
        }
        if let Some(last) = child(style_node, "lastRow") {
            let (fill, _, _, _, _, _) = parse_tc_style(last);
            def.last_row_fill = fill;
            def.last_row_text_color = parse_tc_tx_color(last);
            def.last_row_bold = parse_tc_tx_bold(last);
        }
        if let Some(first) = child(style_node, "firstCol") {
            let (fill, _, _, _, _, _) = parse_tc_style(first);
            def.first_col_fill = fill;
            def.first_col_text_color = parse_tc_tx_color(first);
            def.first_col_bold = parse_tc_tx_bold(first);
        }
        if let Some(last) = child(style_node, "lastCol") {
            let (fill, _, _, _, _, _) = parse_tc_style(last);
            def.last_col_fill = fill;
            def.last_col_text_color = parse_tc_tx_color(last);
            def.last_col_bold = parse_tc_tx_bold(last);
        }

        map.insert(style_id, def);
    }
    map
}

fn parse_table(
    tbl: roxmltree::Node<'_, '_>,
    t: &Transform,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
    zip: &mut PptxZip,
) -> Option<TableElement> {
    // Parse tblPr attributes and look up table style
    let tbl_pr = child(tbl, "tblPr");
    let style_id = tbl_pr
        .and_then(|n| child(n, "tableStyleId"))
        .and_then(|n| n.text())
        .map(|s| s.to_string());
    let flag = |attr_name: &str| -> bool {
        tbl_pr
            .and_then(|n| attr(&n, attr_name))
            .map(|v| v == "1" || v == "true")
            .unwrap_or(false)
    };
    let first_row = flag("firstRow");
    let last_row = flag("lastRow");
    // ECMA-376 §21.1.3.13 (a:tblPr@rtl): right-to-left table layout.
    let rtl = flag("rtl");
    let band_row = flag("bandRow");
    let first_col = flag("firstCol");
    let last_col = flag("lastCol");

    // Load style definitions once
    let table_styles_xml = read_zip_str(zip, "ppt/tableStyles.xml").ok();
    let table_styles = table_styles_xml
        .as_deref()
        .map(|xml| parse_table_styles_xml(xml, theme))
        .unwrap_or_default();
    let style_owned: Option<TableStyleDef> = style_id.as_deref().and_then(|id| {
        table_styles
            .get(id)
            .cloned()
            .or_else(|| table_style_presets::lookup_builtin_table_style(id, theme))
    });
    let style = style_owned.as_ref();

    let cols: Vec<i64> = tbl
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "tblGrid")
        .map(|grid| {
            grid.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "gridCol")
                .filter_map(|n| attr_i64(&n, "w"))
                .collect()
        })
        .unwrap_or_default();

    if cols.is_empty() {
        return None;
    }

    let col_count = cols.len();
    let last_col_idx = col_count.saturating_sub(1);

    let mut rows: Vec<TableRow> = tbl
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "tr")
        .map(|tr| parse_table_row(tr, theme, rels, zip))
        .collect();

    let row_count = rows.len();
    let last_row_idx = row_count.saturating_sub(1);

    // Apply table style fills and borders to each cell
    for (ri, row) in rows.iter_mut().enumerate() {
        for (ci, cell) in row.cells.iter_mut().enumerate() {
            if let Some(s) = style {
                // ── Fill cascade ────────────────────────────────────────────
                let mut effective_fill = s.whole_fill.clone();

                if band_row {
                    // Determine band index excluding firstRow header if present
                    let band_ri = ri.saturating_sub(if first_row { 1 } else { 0 });
                    if !(first_row && ri == 0) {
                        if band_ri % 2 == 0 {
                            if let Some(f) = s.band1h_fill.clone() {
                                effective_fill = Some(f);
                            }
                        } else if let Some(f) = s.band2h_fill.clone() {
                            effective_fill = Some(f);
                        }
                    }
                }
                if first_row && ri == 0 {
                    if let Some(f) = s.first_row_fill.clone() {
                        effective_fill = Some(f);
                    }
                }
                if last_row && ri == last_row_idx {
                    if let Some(f) = s.last_row_fill.clone() {
                        effective_fill = Some(f);
                    }
                }
                if first_col && ci == 0 {
                    if let Some(f) = s.first_col_fill.clone() {
                        effective_fill = Some(f);
                    }
                }
                if last_col && ci == last_col_idx {
                    if let Some(f) = s.last_col_fill.clone() {
                        effective_fill = Some(f);
                    }
                }
                // Cell's own tcPr fill wins
                if cell.fill.is_none() {
                    cell.fill = effective_fill;
                }

                // ── Text-colour cascade (style `<a:tcTxStyle>`, role-keyed) ──
                // Mirrors the fill cascade ordering. The header (firstRow) typically
                // overrides wholeTbl (e.g. white-on-accent header). A run's own
                // explicit colour still wins later, at render time.
                let mut effective_text = s.whole_text_color.clone();
                if first_row && ri == 0 {
                    if let Some(c) = s.first_row_text_color.clone() {
                        effective_text = Some(c);
                    }
                }
                if last_row && ri == last_row_idx {
                    if let Some(c) = s.last_row_text_color.clone() {
                        effective_text = Some(c);
                    }
                }
                if first_col && ci == 0 {
                    if let Some(c) = s.first_col_text_color.clone() {
                        effective_text = Some(c);
                    }
                }
                if last_col && ci == last_col_idx {
                    if let Some(c) = s.last_col_text_color.clone() {
                        effective_text = Some(c);
                    }
                }
                if cell.text_color.is_none() {
                    cell.text_color = effective_text;
                }

                // ── Bold cascade (style `<a:tcTxStyle b="on">`, role-keyed) ──
                // Applied as the cell text body's default bold; a run's own @b wins.
                let mut effective_bold: Option<bool> = None;
                if first_row && ri == 0 {
                    effective_bold = effective_bold.or(s.first_row_bold);
                }
                if last_row && ri == last_row_idx {
                    effective_bold = effective_bold.or(s.last_row_bold);
                }
                if first_col && ci == 0 {
                    effective_bold = effective_bold.or(s.first_col_bold);
                }
                if last_col && ci == last_col_idx {
                    effective_bold = effective_bold.or(s.last_col_bold);
                }
                if let (Some(b), Some(tb)) = (effective_bold, cell.text_body.as_mut()) {
                    if tb.default_bold.is_none() {
                        tb.default_bold = Some(b);
                    }
                }

                // ── Border cascade (style provides inside and outer borders) ──
                // Outer top edge
                if cell.border_t.is_none() && ri == 0 {
                    cell.border_t = s.whole_outer_h.clone();
                }
                // Inner horizontal separator between rows
                if cell.border_t.is_none() && ri > 0 {
                    cell.border_t = s.whole_inside_h.clone();
                }
                // Outer bottom edge
                if cell.border_b.is_none() && ri == last_row_idx {
                    cell.border_b = s.whole_outer_h.clone();
                }
                // Inner bottom separator; firstRow gets its own bottom definition
                if cell.border_b.is_none() {
                    if first_row && ri == 0 {
                        cell.border_b = s
                            .first_row_border_b
                            .clone()
                            .or_else(|| s.whole_inside_h.clone());
                    } else if ri < last_row_idx {
                        cell.border_b = s.whole_inside_h.clone();
                    }
                }
                // Outer left edge
                if cell.border_l.is_none() && ci == 0 {
                    cell.border_l = s.whole_outer_v.clone();
                }
                // Inner vertical separator between cols
                if cell.border_l.is_none() && ci > 0 {
                    cell.border_l = s.whole_inside_v.clone();
                }
                // Outer right edge
                if cell.border_r.is_none() && ci == last_col_idx {
                    cell.border_r = s.whole_outer_v.clone();
                }
                // Inner right separator
                if cell.border_r.is_none() && ci < last_col_idx {
                    cell.border_r = s.whole_inside_v.clone();
                }
            } else {
                // ── Fallback for built-in styles not defined in tableStyles.xml ──
                // Approximate "Medium Style 2": accent1 header fill + thin outer box + row separators.
                let thin = Stroke {
                    color: "A0A096".to_string(),
                    width: 9525,
                    dash_style: None,
                    head_end: None,
                    tail_end: None,
                    cmpd: None,
                };
                if cell.fill.is_none() && first_row && ri == 0 {
                    if let Some(color) = theme.get("accent1") {
                        cell.fill = Some(Fill::Solid {
                            color: color.clone(),
                        });
                    }
                }
                // Outer top
                if cell.border_t.is_none() && ri == 0 {
                    cell.border_t = Some(thin.clone());
                }
                // Inner horizontal separators
                if cell.border_t.is_none() && ri > 0 {
                    cell.border_t = Some(thin.clone());
                }
                // Outer bottom
                if cell.border_b.is_none() && ri == last_row_idx {
                    cell.border_b = Some(thin.clone());
                }
                // Outer left edge
                if cell.border_l.is_none() && ci == 0 {
                    cell.border_l = Some(thin.clone());
                }
                // Outer right edge
                if cell.border_r.is_none() && ci == last_col_idx {
                    cell.border_r = Some(thin.clone());
                }
            }
        }
    }

    Some(TableElement {
        x: t.x,
        y: t.y,
        width: t.cx,
        height: t.cy,
        cols,
        rows,
        rtl,
    })
}

fn parse_table_row(
    tr: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
    zip: &mut PptxZip,
) -> TableRow {
    let height = attr_i64(&tr, "h").unwrap_or(0);
    let cells: Vec<TableCell> = tr
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "tc")
        .map(|tc| parse_table_cell(tc, theme, rels, zip))
        .collect();
    TableRow { height, cells }
}

fn parse_table_cell(
    tc: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
    zip: &mut PptxZip,
) -> TableCell {
    let tc_pr = child(tc, "tcPr");
    // tcPr > anchor controls vertical text alignment within the cell
    let anchor = tc_pr
        .and_then(|n| attr(&n, "anchor"))
        .map(|a| a.to_string());
    let text_body = child(tc, "txBody").map(|n| {
        parse_text_body(
            n,
            theme,
            rels,
            None,
            [None; 9],
            Default::default(), // inherited_level_indents
            &empty_level_bullets(),
            None,
            None,
            None,
            anchor,
            None, // inherited_alignment
            None, // inherited_ea_ln_brk
            None, // inherited_space_before
            None, // inherited_space_after
            None, // inherited_line_spacing
            ShapeKind::Sp,
            zip,
        )
    });

    let fill = tc_pr.and_then(|n| parse_fill(n, theme));

    let border_l = tc_pr
        .and_then(|n| child(n, "lnL"))
        .and_then(|n| parse_stroke(n, theme));
    let border_r = tc_pr
        .and_then(|n| child(n, "lnR"))
        .and_then(|n| parse_stroke(n, theme));
    let border_t = tc_pr
        .and_then(|n| child(n, "lnT"))
        .and_then(|n| parse_stroke(n, theme));
    let border_b = tc_pr
        .and_then(|n| child(n, "lnB"))
        .and_then(|n| parse_stroke(n, theme));
    // Diagonal borders: lnTlToBr (top-left→bottom-right) and lnBlToTr (bottom-left→top-right)
    // These are CT_LineProperties elements (same structure as lnL/lnR/lnT/lnB)
    let diagonal_tl = tc_pr
        .and_then(|n| child(n, "lnTlToBr"))
        .and_then(|n| parse_stroke(n, theme));
    let diagonal_tr = tc_pr
        .and_then(|n| child(n, "lnBlToTr"))
        .and_then(|n| parse_stroke(n, theme));

    let grid_span: u32 = attr(&tc, "gridSpan")
        .and_then(|v| v.parse().ok())
        .unwrap_or(1);
    let row_span: u32 = attr(&tc, "rowSpan")
        .and_then(|v| v.parse().ok())
        .unwrap_or(1);
    let h_merge = attr(&tc, "hMerge")
        .map(|v| v == "1" || v == "true")
        .unwrap_or(false);
    let v_merge = attr(&tc, "vMerge")
        .map(|v| v == "1" || v == "true")
        .unwrap_or(false);

    TableCell {
        text_body,
        fill,
        text_color: None,
        border_l,
        border_r,
        border_t,
        border_b,
        diagonal_tl,
        diagonal_tr,
        grid_span,
        row_span,
        h_merge,
        v_merge,
    }
}

// ===========================
//  Slide parser
// ===========================

/// `<p:sld show="0">` / `show="false"` marks a slide hidden in the slide show
/// (ECMA-376 §19.3.1.38 `sld` / `CT_Slide` — `show`, xsd:boolean, default true).
/// Absent or any truthy value ⇒ shown. NB: this matches the FALSY literals —
/// the inverse of the codebase's usual `== "1" || == "true"` truthy check —
/// because `show` defaults to true, so a slide is hidden only on explicit false.
fn slide_is_hidden(root: roxmltree::Node) -> bool {
    matches!(root.attribute("show"), Some("0") | Some("false"))
}

/// Walk a master/layout `<p:cSld><p:spTree>` and collect its NON-placeholder
/// (decorative) shapes as `SlideElement`s (ECMA-376 §19.3.1.38 showMasterSp /
/// layout decorations). Placeholders are skipped (`skip_placeholders = true`).
/// Extracted verbatim from the two inline spTree walks in `parse_slide` so the
/// master decorations can be pre-computed once per cached master; `theme`,
/// `rels`, `smartart_drawings`, and `part_dir` are the resolution context of the
/// tree being walked. Appends to `out` (callers control ordering/gating).
fn extract_decorative_shapes(
    root: roxmltree::Node<'_, '_>,
    part_dir: &str,
    rels: &HashMap<String, String>,
    smartart_drawings: &HashMap<String, String>,
    theme: &HashMap<String, String>,
    zip: &mut PptxZip,
    out: &mut Vec<SlideElement>,
) {
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        let empty_lph = LayoutPlaceholders::default();
        for node in sp_tree.children().filter(|n| n.is_element()) {
            parse_sp_tree_node(
                node,
                &empty_lph,
                part_dir,
                rels,
                smartart_drawings,
                zip,
                theme,
                out,
                true, // skip placeholder shapes
                None, // no inherited group fill at top level
            );
        }
    }
}

// Threads the full master+layout inheritance context (per-type font sizes,
// bullets, anchors, transforms, alignments, spacing, bold/italic/caps/color
// maps) plus zip/theme into one slide parse; this is the inheritance chain
// ECMA-376 requires, not an arbitrary parameter bag.
#[allow(clippy::too_many_arguments)]
fn parse_slide(
    xml: &str,
    // The layout's single-pass extraction (placeholders + layout bg + layout
    // showMasterSp), built/cached by the caller against this slide's effective
    // theme (D4). `layout_xml` is still passed for the per-slide DECORATIVE walk
    // only (its shapes bind to the slide's own smartart + theme + zip, so they
    // can't live in the cached `ParsedLayout`).
    parsed_layout: &ParsedLayout,
    layout_xml: Option<&str>,
    layout_rels: &HashMap<String, String>,
    layout_dir: &str,
    bundle: &MasterBundle,
    eff: Option<&EffectiveMaster>,
    index: usize,
    rels: &HashMap<String, String>,
    smartart_drawings: &HashMap<String, String>,
    zip: &mut PptxZip,
) -> Result<Slide, Box<dyn std::error::Error>> {
    // Destructure the per-slide master bundle into the local names the rest of
    // this function uses. `theme` here is the slide's effective theme (the
    // master's own theme with its <p:clrMap> baked in), so scheme colors
    // resolve against the right palette per slide.
    // Only the fields this function still consumes directly are bound; the
    // master INHERITANCE maps (font sizes, level sizes/indents/bullets, anchors,
    // transforms, alignments, ea-ln-brk, spacing) now feed `parse_layout` in the
    // caller, which produced the `ParsedLayout` passed in. `theme` here is the
    // slide's effective theme (master clrMap baked in) so scheme colors resolve
    // against the right palette per slide.
    let MasterBundle {
        theme,
        master_xml,
        master_rels,
        master_dir,
        master_smartart_drawings,
        master_bg,
        master_decorative,
        master_bold,
        master_italic,
        master_caps,
        master_color,
        ..
    } = bundle;
    // When the slide/layout carries a `<p:clrMapOvr><a:overrideClrMapping>`
    // (ECMA-376 §19.3.1.7), the caller recomputed the master's theme-dependent
    // fields against the slide's effective mapping (`EffectiveMaster`); use them
    // in place of the master's frozen values so that BOTH the slide's own scheme
    // colors AND master-inherited ones (the master `<p:bg>`, master txStyles
    // placeholder colors) resolve against the override mapping (§20.1.6.8).
    // Otherwise fall back to the master bundle's values. (Master bullet colors
    // flow through `parsed_layout`, already override-adjusted by the caller.)
    let theme: &HashMap<String, String> = eff.map(|e| &e.theme).unwrap_or(theme);
    let master_xml: Option<&str> = master_xml.as_deref();
    let master_dir: &str = master_dir.as_str();
    let master_bg: Option<Fill> = match eff {
        Some(e) => e.master_bg.clone(),
        None => master_bg.clone(),
    };
    let master_color: &HashMap<String, String> =
        eff.map(|e| &e.master_color).unwrap_or(master_color);

    // The layout placeholder inheritance was resolved once in `parse_layout`
    // (cached across slides sharing this layout, or rebuilt for a clrMapOvr
    // slide) against this slide's effective theme. Clone it so the per-slide
    // master txStyles fallbacks below can be layered on without mutating the
    // shared/cached instance.
    let mut lph = parsed_layout.placeholders.clone();
    // Fall back to master txStyles defRPr @b/@i when the layout did not specify
    // bold/italic for a placeholder type. Without this, e.g. the master titleStyle's
    // b="1" is not applied to ctrTitle / title placeholders.
    for (t, b) in master_bold.iter() {
        lph.by_type_bold.entry(t.clone()).or_insert(*b);
    }
    for (t, i) in master_italic.iter() {
        lph.by_type_italic.entry(t.clone()).or_insert(*i);
    }
    for (t, c) in master_caps.iter() {
        lph.by_type_caps.entry(t.clone()).or_insert(c.clone());
    }
    for (t, c) in master_color.iter() {
        lph.by_type_master_color
            .entry(t.clone())
            .or_insert(c.clone());
    }

    note_layout_master_parse();
    let doc = roxmltree::Document::parse(xml)?;
    let root = doc.root_element(); // <p:sld>
    let hidden = slide_is_hidden(root);
    let c_sld = child(root, "cSld");

    // Background chain: slide → layout → master. Each level resolves a blip
    // background (§20.1.8.14) against its own rels + part directory, so the
    // closures are run sequentially (one mutable borrow of `zip` at a time).
    let mut background: Option<Fill> = None;

    // Slide-level bg (rels = slide rels, part dir = ppt/slides).
    if let Some(n) = c_sld {
        let mut resolve = |rid: &str| -> Option<String> {
            let target = rels.get(rid)?;
            let path = resolve_path("ppt/slides", target);
            // Resolve to the zip path; verify the part exists so a dangling
            // rId still yields None (the bg chain then falls through to the
            // next level), preserving the prior data-URL behaviour.
            // `index_for_name` reads the central directory only (no inflate),
            // unlike the former `read_zip_bytes` which decompressed to discard.
            zip.index_for_name(&path)?;
            Some(path)
        };
        background = parse_background(n, theme, &mut resolve);
    }

    // Layout-level bg: resolved once in `parse_layout` (against this slide's
    // effective theme) and applied only when the slide's own bg chain is empty.
    if background.is_none() {
        background = parsed_layout.background.clone();
    }

    // Master-level bg (resolved by the caller before parse_slide; already a Fill).
    let background = background.or(master_bg);

    let sp_tree = c_sld
        .and_then(|n| child(n, "spTree"))
        .ok_or("missing spTree")?;

    let slide_dir = "ppt/slides";
    let mut elements = Vec::new();

    // ── showMasterSp resolution (ECMA-376 §19.3.1.38 sld / §19.3.1.39
    // sldLayout, AG_ChildSlide, default true) ─────────────────────────────
    // Master decorative shapes are composited beneath the slide only when both
    // the slide and its layout permit it. Either one setting showMasterSp="0"
    // suppresses the master's spTree decorations (the slide flag is honored for
    // the slide itself; the layout flag — read once in `parse_layout` — for
    // shapes inherited through it).
    let slide_show_master_sp = read_show_master_sp(root);
    let show_master_sp = slide_show_master_sp && parsed_layout.show_master_sp;

    // ── Master non-placeholder shapes (rendered BELOW layout & slide) ─────
    // The slide master's spTree may carry decorative pictures/shapes (logos,
    // bands) that are not placeholder anchors. PowerPoint composites them at
    // the very bottom, beneath the layout's decorations and the slide content.
    // Gated by showMasterSp (above). Placeholders are skipped — only the
    // master's decorative content is drawn here.
    //
    // These were pre-extracted once per cached master in `build_master_bundle`
    // (resolved against the master's baked theme), so the common no-override
    // slide clones them instead of re-parsing the master XML + re-walking its
    // spTree. A slide with a `<p:clrMapOvr>` (`eff.is_some()`) must re-resolve
    // them against its override theme, so it re-extracts from the master XML —
    // exactly what the old unconditional inline walk did, now only on the rare
    // override path. `elements` is still empty here, so ordering (master
    // decorations first) is unchanged either way.
    if show_master_sp {
        if eff.is_some() {
            if let Some(mxml) = master_xml {
                note_layout_master_parse();
                if let Ok(mdoc) = roxmltree::Document::parse(mxml) {
                    extract_decorative_shapes(
                        mdoc.root_element(),
                        master_dir,
                        master_rels,
                        master_smartart_drawings,
                        theme,
                        zip,
                        &mut elements,
                    );
                }
            }
        } else {
            elements.extend(master_decorative.iter().cloned());
        }
    }

    // ── Layout non-placeholder shapes (rendered BEFORE slide shapes) ──────
    // These are decorative background elements defined in the slide layout
    // (e.g. coloured bands, logos) that are not placeholder anchors.
    if let Some(lxml) = layout_xml {
        note_layout_master_parse();
        if let Ok(ldoc) = roxmltree::Document::parse(lxml) {
            let lroot = ldoc.root_element();
            if let Some(lsp_tree) = child(lroot, "cSld").and_then(|n| child(n, "spTree")) {
                let empty_lph = LayoutPlaceholders::default();
                for node in lsp_tree.children().filter(|n| n.is_element()) {
                    parse_sp_tree_node(
                        node,
                        &empty_lph,
                        layout_dir,
                        layout_rels,
                        smartart_drawings,
                        zip,
                        theme,
                        &mut elements,
                        true, // skip placeholder shapes
                        None, // no inherited group fill at top level
                    );
                }
            }
        }
    }

    // ── Slide shapes ─────────────────────────────────────────────────────
    for node in sp_tree.children().filter(|n| n.is_element()) {
        parse_sp_tree_node(
            node,
            &lph,
            slide_dir,
            rels,
            smartart_drawings,
            zip,
            theme,
            &mut elements,
            false,
            None,
        );
    }

    // ── Notes slide & comments (Phase 2 surfacing only — no rendering) ────
    let notes = load_notes_slide(zip, rels);
    let comments = load_pptx_comments(zip, rels);

    Ok(Slide {
        index,
        slide_number: index + 1,
        background,
        elements,
        notes,
        comments,
        hidden,
    })
}

/// Resolve the slide's `notesSlide` relationship, read the notes part, and
/// return its plain text (paragraphs joined by '\n'). Returns `None` when
/// the slide has no notes part or the part can't be read.
fn load_notes_slide(zip: &mut PptxZip, rels: &HashMap<String, String>) -> Option<String> {
    // rels here is the slide's _rels map (rId → Target) parsed by the caller.
    // The relationship Type ends with "/notesSlide". The cleanest way to find
    // the right entry is to look at every value in the map and pick the one
    // pointing at "../notesSlides/...".
    let target = rels.values().find(|t| t.contains("notesSlides/"))?;
    let path = if target.starts_with('/') {
        target.trim_start_matches('/').to_string()
    } else {
        // Relative to ppt/slides/ — resolve "../notesSlides/notesSlide1.xml".
        resolve_path("ppt/slides", target)
    };
    let xml = read_zip_str(zip, &path).ok()?;
    let doc = roxmltree::Document::parse(&xml).ok()?;
    let mut buf = String::new();
    let mut prev_was_text = false;
    for n in doc.descendants() {
        if !n.is_element() {
            continue;
        }
        let name = n.tag_name().name();
        if name == "p" && prev_was_text {
            buf.push('\n');
            prev_was_text = false;
        }
        if name == "t" {
            if let Some(s) = n.text() {
                buf.push_str(s);
                prev_was_text = true;
            }
        }
    }
    let trimmed = buf.trim();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed.to_string())
    }
}

/// Resolve and parse the slide's `comments` relationship (legacy
/// `<p:cmLst>` format). Modern threaded comments live in a different
/// namespace and are not yet supported.
fn load_pptx_comments(zip: &mut PptxZip, rels: &HashMap<String, String>) -> Vec<PptxComment> {
    let Some(target) = rels.values().find(|t| t.contains("comments/")) else {
        return Vec::new();
    };
    let path = if target.starts_with('/') {
        target.trim_start_matches('/').to_string()
    } else {
        resolve_path("ppt/slides", target)
    };
    let Ok(xml) = read_zip_str(zip, &path) else {
        return Vec::new();
    };
    let Ok(doc) = roxmltree::Document::parse(&xml) else {
        return Vec::new();
    };

    // commentAuthors.xml is a top-level part — look it up directly.
    let author_xml = read_zip_str(zip, "ppt/commentAuthors.xml").ok();
    let mut authors: HashMap<String, String> = HashMap::new();
    if let Some(ax) = author_xml {
        if let Ok(adoc) = roxmltree::Document::parse(&ax) {
            for a in adoc
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "cmAuthor")
            {
                let id = a.attribute("id").unwrap_or("").to_string();
                let name = a.attribute("name").unwrap_or("").to_string();
                if !id.is_empty() && !name.is_empty() {
                    authors.insert(id, name);
                }
            }
        }
    }

    let mut out = Vec::new();
    for cm in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "cm")
    {
        let author_id = cm.attribute("authorId").unwrap_or("");
        let author = authors.get(author_id).cloned();
        let date = cm
            .attribute("dt")
            .map(String::from)
            .filter(|s| !s.is_empty());
        let text = cm
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "text")
            .and_then(|n| n.text().map(String::from))
            .unwrap_or_default();
        out.push(PptxComment { author, date, text });
    }
    out
}

// Recurses the shape tree carrying placeholder/layout context, theme, rels,
// smartart drawings, the zip, the output buffer and inherited group fill.
#[allow(clippy::too_many_arguments)]
fn parse_sp_tree_node(
    node: roxmltree::Node<'_, '_>,
    lph: &LayoutPlaceholders,
    slide_dir: &str,
    rels: &HashMap<String, String>,
    smartart_drawings: &HashMap<String, String>,
    zip: &mut PptxZip,
    theme: &HashMap<String, String>,
    out: &mut Vec<SlideElement>,
    skip_placeholders: bool,
    group_fill: Option<&Fill>,
) {
    match node.tag_name().name() {
        "sp" => {
            if skip_placeholders && is_placeholder(node) {
                return;
            }
            // Image-filled shape: spPr > blipFill > blip r:embed → render as PictureElement
            let sp_pr_node = child(node, "spPr");
            let blip_fill_node = sp_pr_node.and_then(|p| child(p, "blipFill"));
            let blip_rid = blip_fill_node
                .and_then(|bf| child(bf, "blip"))
                .and_then(|b| attr_r(&b, "embed"));
            if let Some(ref rid) = blip_rid {
                if let Some(xfrm_node) = sp_pr_node.and_then(|p| child(p, "xfrm")) {
                    let t = parse_xfrm(xfrm_node);
                    if t.cx > 0 && t.cy > 0 {
                        if let Some(target) = rels.get(rid) {
                            let image_path = resolve_path(slide_dir, target);
                            if let Ok(bytes) = ooxml_common::zip::read_zip_bytes(zip, &image_path) {
                                let mime_type = mime_from_ext(&image_path).to_owned();
                                let (intrinsic_width_px, intrinsic_height_px) =
                                    match png_size_from_bytes(&bytes) {
                                        Some((w, h)) => (Some(w), Some(h)),
                                        None => (None, None),
                                    };
                                // Microsoft 2016 SVG extension — a blipFill-painted
                                // sp can carry the same svgBlip vector original as a
                                // real p:pic; surface it so the renderer prefers it.
                                let svg_image_path = blip_fill_node
                                    .and_then(|bf| child(bf, "blip"))
                                    .and_then(|b| svg_blip_path(b, slide_dir, rels, zip));
                                // §20.1.9.18 — the sp's prstGeom (any preset, not
                                // just roundRect) is the picture's clip silhouette.
                                let (prst_geom, prst_adjust) =
                                    sp_pr_node.map(parse_pic_prst_geom).unwrap_or((None, None));
                                let cust_geom = sp_pr_node
                                    .and_then(|p| child(p, "custGeom"))
                                    .map(parse_cust_geom);
                                // §20.1.8.16 effectLst applies to a sp painted as
                                // a picture (blipFill) just like a regular p:pic.
                                let EffectLst {
                                    shadow,
                                    inner_shadow,
                                    glow,
                                    soft_edge,
                                    reflection,
                                } = parse_effect_lst(
                                    sp_pr_node.and_then(|p| child(p, "effectLst")),
                                    theme,
                                );
                                // §20.1.2.2.24 — a blipFill-painted sp can carry
                                // an `<a:ln>` border just like a real p:pic.
                                let stroke = sp_pr_node
                                    .and_then(|p| child(p, "ln"))
                                    .and_then(|n| parse_stroke(n, theme));
                                out.push(SlideElement::Picture(PictureElement {
                                    x: t.x,
                                    y: t.y,
                                    width: t.cx,
                                    height: t.cy,
                                    rotation: t.rot,
                                    flip_h: t.flip_h,
                                    flip_v: t.flip_v,
                                    image_path,
                                    mime_type,
                                    svg_image_path,
                                    intrinsic_width_px,
                                    intrinsic_height_px,
                                    stroke,
                                    prst_geom,
                                    prst_adjust,
                                    src_rect: blip_fill_node.and_then(parse_src_rect),
                                    alpha: blip_fill_node.and_then(parse_blip_alpha),
                                    cust_geom,
                                    shadow,
                                    inner_shadow,
                                    glow,
                                    soft_edge,
                                    reflection,
                                    scene3d: sp_pr_node.and_then(parse_scene3d),
                                    sp3d: sp_pr_node.and_then(parse_sp3d),
                                }));
                                return;
                            }
                        }
                    }
                }
            }
            // Picture-placeholder inheritance: slide sp has a ph but no own blipFill →
            // look up an inherited blipFill from the layout placeholder. Transform
            // comes from the slide's xfrm when present, otherwise from the layout.
            if blip_rid.is_none() {
                if let Some(ph) = node
                    .descendants()
                    .find(|n| n.is_element() && n.tag_name().name() == "ph")
                {
                    let ph_type = attr(&ph, "type").unwrap_or_else(|| "body".into());
                    let ph_idx: Option<u32> = attr(&ph, "idx").and_then(|v| v.parse().ok());
                    if let Some(bf) = lph.lookup_blip_fill(&ph_type, ph_idx) {
                        let slide_xfrm = sp_pr_node.and_then(|p| child(p, "xfrm")).map(parse_xfrm);
                        let t = slide_xfrm.or_else(|| lph.lookup(&ph_type, ph_idx).cloned());
                        if let Some(t) = t {
                            if t.cx > 0 && t.cy > 0 {
                                // §20.1.2.2.24 — honour the slide sp's own `<a:ln>`
                                // border, falling back to the inherited layout
                                // placeholder stroke when the slide omits one.
                                let stroke = match sp_pr_node.and_then(|p| child(p, "ln")) {
                                    Some(ln) => parse_stroke(ln, theme),
                                    None => lph.lookup_stroke(&ph_type, ph_idx),
                                };
                                out.push(SlideElement::Picture(PictureElement {
                                    x: t.x,
                                    y: t.y,
                                    width: t.cx,
                                    height: t.cy,
                                    rotation: t.rot,
                                    flip_h: t.flip_h,
                                    flip_v: t.flip_v,
                                    image_path: bf.image_path,
                                    mime_type: bf.mime_type,
                                    // TODO: an inherited layout-placeholder blipFill
                                    // (LayoutPlaceholders::lookup_blip_fill) does not
                                    // yet carry the svgBlip extension. Picture
                                    // placeholders pointing at an SVG are rare; thread
                                    // the svg path through BlipFill if a sample needs it.
                                    svg_image_path: None,
                                    // Intrinsic size is only consumed by the ink
                                    // fallback (PNG-IHDR centering); inherited
                                    // placeholder pictures stretch to the box, so
                                    // None matches the prior behaviour.
                                    intrinsic_width_px: None,
                                    intrinsic_height_px: None,
                                    stroke,
                                    prst_geom: None,
                                    prst_adjust: None,
                                    src_rect: bf.src_rect,
                                    alpha: bf.alpha,
                                    cust_geom: None,
                                    shadow: None,
                                    inner_shadow: None,
                                    glow: None,
                                    soft_edge: None,
                                    reflection: None,
                                    scene3d: None,
                                    sp3d: None,
                                }));
                                return;
                            }
                        }
                    }
                }
            }
            if let Some(shape) = parse_shape(node, lph, theme, rels, group_fill, zip) {
                out.push(SlideElement::Shape(shape));
            }
        }
        "pic" => {
            if let Some(media) = parse_media(node, slide_dir, rels) {
                out.push(SlideElement::Media(media));
            } else if let Some(pic) = parse_picture(node, slide_dir, rels, theme, zip) {
                out.push(SlideElement::Picture(pic));
            } else {
                // Placeholder pic: no xfrm in spPr — position comes from layout by_idx
                let ph_idx = node
                    .descendants()
                    .find(|n| n.is_element() && n.tag_name().name() == "ph")
                    .and_then(|ph| attr(&ph, "idx"))
                    .and_then(|s| s.parse::<u32>().ok());
                if let Some(idx) = ph_idx {
                    if let Some(t) = lph.by_idx.get(&idx) {
                        let blip_fill = child(node, "blipFill");
                        let blip = blip_fill.and_then(|bf| child(bf, "blip"));
                        let r_id = blip.and_then(|b| attr_r(&b, "embed"));
                        if let Some(rid) = r_id {
                            if let Some(rel_target) = rels.get(&rid) {
                                let image_path = resolve_path(slide_dir, rel_target);
                                if let Ok(image_bytes) =
                                    ooxml_common::zip::read_zip_bytes(zip, &image_path)
                                {
                                    let mime_type = mime_from_ext(&image_path).to_owned();
                                    let (intrinsic_width_px, intrinsic_height_px) =
                                        match png_size_from_bytes(&image_bytes) {
                                            Some((w, h)) => (Some(w), Some(h)),
                                            None => (None, None),
                                        };
                                    // Microsoft 2016 SVG extension on the placeholder
                                    // p:pic's blip — prefer the vector original.
                                    let svg_image_path =
                                        blip.and_then(|b| svg_blip_path(b, slide_dir, rels, zip));
                                    // §20.1.2.2.24 — placeholder pic border: the
                                    // p:pic's own `<a:ln>`, else the inherited
                                    // layout placeholder stroke.
                                    let stroke =
                                        match child(node, "spPr").and_then(|p| child(p, "ln")) {
                                            Some(ln) => parse_stroke(ln, theme),
                                            None => lph.by_idx_stroke.get(&idx).cloned(),
                                        };
                                    out.push(SlideElement::Picture(PictureElement {
                                        x: t.x,
                                        y: t.y,
                                        width: t.cx,
                                        height: t.cy,
                                        rotation: t.rot,
                                        flip_h: t.flip_h,
                                        flip_v: t.flip_v,
                                        image_path,
                                        mime_type,
                                        svg_image_path,
                                        intrinsic_width_px,
                                        intrinsic_height_px,
                                        stroke,
                                        prst_geom: None,
                                        prst_adjust: None,
                                        src_rect: blip_fill.and_then(parse_src_rect),
                                        alpha: blip_fill.and_then(parse_blip_alpha),
                                        cust_geom: None,
                                        shadow: None,
                                        inner_shadow: None,
                                        glow: None,
                                        soft_edge: None,
                                        reflection: None,
                                        scene3d: None,
                                        sp3d: None,
                                    }));
                                }
                            }
                        }
                    }
                }
            }
        }
        "AlternateContent" => {
            // mc:AlternateContent wraps modern elements (e.g. chartEx inside grpSp,
            // or p:contentPart for ink/handwriting). Try Choice first; if it produces
            // nothing (Choice contains an element we don't render, like contentPart),
            // fall back to Fallback which usually carries a rasterized p:pic alternative.
            let before = out.len();
            let choice_node = node
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "Choice");
            // Choice contains p:contentPart → ink/handwriting. PowerPoint renders the
            // InkML directly; the Fallback PNG is a low-resolution rasterization at
            // the stroke's actual pixel extent. For empty/single-tap strokes the PNG
            // is only a few pixels and should not be stretched to the bounding box.
            let is_ink_fallback = choice_node.is_some_and(|c| {
                c.descendants()
                    .any(|n| n.is_element() && n.tag_name().name() == "contentPart")
            });
            if let Some(choice_node) = choice_node {
                for child_node in choice_node.children().filter(|n| n.is_element()) {
                    parse_sp_tree_node(
                        child_node,
                        lph,
                        slide_dir,
                        rels,
                        smartart_drawings,
                        zip,
                        theme,
                        out,
                        skip_placeholders,
                        group_fill,
                    );
                }
            }
            if out.len() == before {
                if let Some(fallback_node) = node
                    .children()
                    .find(|n| n.is_element() && n.tag_name().name() == "Fallback")
                {
                    for child_node in fallback_node.children().filter(|n| n.is_element()) {
                        parse_sp_tree_node(
                            child_node,
                            lph,
                            slide_dir,
                            rels,
                            smartart_drawings,
                            zip,
                            theme,
                            out,
                            skip_placeholders,
                            group_fill,
                        );
                    }
                }
                if is_ink_fallback {
                    // Render the fallback PNG at its natural pixel size centered
                    // inside the original bounding box, so a 6×6 px empty-stroke
                    // PNG is drawn as a 6×6 px dot rather than stretched into a
                    // blocky cross. Visible strokes whose PNG natural size already
                    // matches the box keep their existing extent.
                    for el in &mut out[before..] {
                        if let SlideElement::Picture(p) = el {
                            if let (Some(nat_w_px), Some(nat_h_px)) =
                                (p.intrinsic_width_px, p.intrinsic_height_px)
                            {
                                let nat_w = (nat_w_px as i64) * EMU_PER_PX_96DPI;
                                let nat_h = (nat_h_px as i64) * EMU_PER_PX_96DPI;
                                if nat_w < p.width && nat_h < p.height {
                                    p.x += (p.width - nat_w) / 2;
                                    p.y += (p.height - nat_h) / 2;
                                    p.width = nat_w;
                                    p.height = nat_h;
                                }
                            }
                        }
                    }
                }
            }
        }
        "graphicFrame" => {
            let xfrm_node = child(node, "xfrm");
            let t = xfrm_node.map(parse_xfrm).unwrap_or_default();

            // Table
            let tbl_node = node
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "tbl");
            if let Some(tbl_node) = tbl_node {
                if let Some(table) = parse_table(tbl_node, &t, theme, rels, zip) {
                    out.push(SlideElement::Table(table));
                }
                return;
            }

            // SmartArt: render the PowerPoint-prebaked fallback drawing1.xml when present.
            // Layout-engine reconstruction is not implemented; we rely on the cached dsp:spTree.
            if let Some(gd) = node
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "graphicData")
            {
                let uri = attr(&gd, "uri").unwrap_or_default();
                if uri == "http://schemas.openxmlformats.org/drawingml/2006/diagram" {
                    if let Some(rel_ids) = child(gd, "relIds") {
                        if let Some(dm_rid) = attr_r(&rel_ids, "dm") {
                            if let Some(drawing_xml) = smartart_drawings.get(&dm_rid) {
                                parse_smartart_drawing(drawing_xml, &t, theme, out, zip);
                                return;
                            }
                        }
                    }
                }
            }

            // Chart
            if let Some(gd) = node
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "graphicData")
            {
                let uri = attr(&gd, "uri").unwrap_or_default();
                // Both <c:chart> and <cx:chart> share the local name "chart"
                let chart_rid = gd
                    .descendants()
                    .find(|n| n.is_element() && n.tag_name().name() == "chart")
                    .and_then(|n| attr_r(&n, "id"));
                if let Some(rid) = chart_rid {
                    if let Some(rel_target) = rels.get(&rid) {
                        let chart_path = resolve_path(slide_dir, rel_target);
                        if let Ok(chart_xml) = read_zip_str(zip, &chart_path) {
                            let chart_opt = if uri.contains("chartex") || uri.contains("chartEx") {
                                parse_chartex(&chart_xml, theme)
                            } else {
                                parse_legacy_chart(&chart_xml, theme)
                            };
                            if let Some(mut chart) = chart_opt {
                                chart.x = t.x;
                                chart.y = t.y;
                                chart.width = t.cx;
                                chart.height = t.cy;
                                out.push(SlideElement::Chart(chart));
                            }
                        }
                    }
                }
            }
        }
        "grpSp" => {
            let grp_sp_pr = child(node, "grpSpPr");
            let gt: Option<GroupTransform> =
                grp_sp_pr.and_then(|pr| child(pr, "xfrm")).map(|xfrm| {
                    let off = child(xfrm, "off");
                    let ext = child(xfrm, "ext");
                    let ch_off = child(xfrm, "chOff");
                    let ch_ext = child(xfrm, "chExt");
                    GroupTransform {
                        x: off.and_then(|n| attr_i64(&n, "x")).unwrap_or(0),
                        y: off.and_then(|n| attr_i64(&n, "y")).unwrap_or(0),
                        cx: ext.and_then(|n| attr_i64(&n, "cx")).unwrap_or(0),
                        cy: ext.and_then(|n| attr_i64(&n, "cy")).unwrap_or(0),
                        ch_x: ch_off.and_then(|n| attr_i64(&n, "x")).unwrap_or(0),
                        ch_y: ch_off.and_then(|n| attr_i64(&n, "y")).unwrap_or(0),
                        ch_cx: ch_ext.and_then(|n| attr_i64(&n, "cx")).unwrap_or(0),
                        ch_cy: ch_ext.and_then(|n| attr_i64(&n, "cy")).unwrap_or(0),
                        flip_h: attr(&xfrm, "flipH")
                            .map(|v| v == "1" || v == "true")
                            .unwrap_or(false),
                        flip_v: attr(&xfrm, "flipV")
                            .map(|v| v == "1" || v == "true")
                            .unwrap_or(false),
                        rot: attr_f64(&xfrm, "rot").unwrap_or(0.0) / 60000.0,
                    }
                });

            // Determine the fill to propagate to child shapes that use grpFill.
            // - Group has solidFill/noFill → use that as child group fill
            // - Group has grpFill → inherit from parent
            // - Group has no fill → inherit from parent
            let grp_has_grp_fill = grp_sp_pr.and_then(|pr| child(pr, "grpFill")).is_some();
            let grp_explicit_fill = grp_sp_pr.and_then(|pr| parse_fill(pr, theme));
            let child_group_fill: Option<Fill> = if grp_has_grp_fill {
                group_fill.cloned()
            } else if let Some(f) = grp_explicit_fill {
                Some(f)
            } else {
                group_fill.cloned()
            };

            let start = out.len();
            for child_node in node.children().filter(|n| n.is_element()) {
                parse_sp_tree_node(
                    child_node,
                    lph,
                    slide_dir,
                    rels,
                    smartart_drawings,
                    zip,
                    theme,
                    out,
                    skip_placeholders,
                    child_group_fill.as_ref(),
                );
            }
            if let Some(gt) = gt {
                for el in &mut out[start..] {
                    apply_group_transform_to_element(el, &gt);
                }
            }
        }
        "cxnSp" => {
            // Connector shape: parse as a line/shape element
            if skip_placeholders && is_placeholder(node) {
                return;
            }
            if let Some(shape) = parse_connector(node, theme) {
                out.push(SlideElement::Shape(shape));
            }
        }
        _ => {}
    }
}

/// Render a SmartArt diagram's pre-baked fallback drawing (drawing1.xml).
/// The drawing file stores a standard dsp:spTree whose shapes share the same
/// element shape as a regular p:sp / p:cxnSp / p:grpSp (children use the `a:`
/// namespace). Coordinates inside the drawing are local to the enclosing
/// graphicFrame, so we translate every emitted element by the graphicFrame's xfrm.
/// Read a SmartArt shape's explicit text frame from `<dsp:txXfrm>`
/// (`<a:off>` / `<a:ext>`), in the drawing's local EMU coordinates. Returns
/// `None` when the shape has no txXfrm (most ordinary shapes). The optional
/// `rot` attribute is intentionally not consumed here — none of the supported
/// SmartArt layouts rotate the text frame independently of the shape, and the
/// renderer already rotates text by the shape's own rotation.
fn parse_tx_xfrm(sp_node: roxmltree::Node<'_, '_>) -> Option<TextRect> {
    let txx = child(sp_node, "txXfrm")?;
    let off = child(txx, "off")?;
    let ext = child(txx, "ext")?;
    Some(TextRect {
        x: attr_i64(&off, "x")?,
        y: attr_i64(&off, "y")?,
        width: attr_i64(&ext, "cx")?,
        height: attr_i64(&ext, "cy")?,
    })
}

fn parse_smartart_drawing(
    drawing_xml: &str,
    gf_xfrm: &Transform,
    theme: &HashMap<String, String>,
    out: &mut Vec<SlideElement>,
    zip: &mut PptxZip,
) {
    let doc = match roxmltree::Document::parse(drawing_xml) {
        Ok(d) => d,
        Err(_) => return,
    };
    let root = doc.root_element();
    let Some(sp_tree) = child(root, "spTree") else {
        return;
    };

    let empty_lph = LayoutPlaceholders::default();

    let start = out.len();
    for node in sp_tree.children().filter(|n| n.is_element()) {
        // dsp:sp / dsp:cxnSp / dsp:grpSp share local names with their p:* counterparts.
        // Shapes inside drawing1.xml carry no `a:blip r:embed` picture fills, so
        // `rels` is an empty stub. `zip` is still threaded through: a dsp shape's
        // text body may declare a `<a:buBlip>` bullet whose part existence must be
        // verified (it resolves against the empty `rels`, so it always falls
        // through to Bullet::Inherit, but the resolver still needs the archive).
        let empty_rels: HashMap<String, String> = HashMap::new();
        let empty_smartart: HashMap<String, String> = HashMap::new();
        // Dispatch the relevant branches (sp/cxnSp/grpSp) directly rather than via
        // parse_sp_tree_node, since the dsp subtree never contains embedded pictures.
        match node.tag_name().name() {
            "sp" => {
                if let Some(mut shape) =
                    parse_shape(node, &empty_lph, theme, &empty_rels, None, zip)
                {
                    shape.text_rect = parse_tx_xfrm(node);
                    out.push(SlideElement::Shape(shape));
                }
            }
            "cxnSp" => {
                if let Some(shape) = parse_connector(node, theme) {
                    out.push(SlideElement::Shape(shape));
                }
            }
            "grpSp" => {
                // Recursively render a group by collecting its children into a temp buffer
                // and applying the group's xfrm, mirroring the grpSp branch in parse_sp_tree_node.
                let grp_sp_pr = child(node, "grpSpPr");
                let gt: Option<GroupTransform> =
                    grp_sp_pr.and_then(|pr| child(pr, "xfrm")).map(|xfrm| {
                        let off = child(xfrm, "off");
                        let ext = child(xfrm, "ext");
                        let ch_off = child(xfrm, "chOff");
                        let ch_ext = child(xfrm, "chExt");
                        GroupTransform {
                            x: off.and_then(|n| attr_i64(&n, "x")).unwrap_or(0),
                            y: off.and_then(|n| attr_i64(&n, "y")).unwrap_or(0),
                            cx: ext.and_then(|n| attr_i64(&n, "cx")).unwrap_or(0),
                            cy: ext.and_then(|n| attr_i64(&n, "cy")).unwrap_or(0),
                            ch_x: ch_off.and_then(|n| attr_i64(&n, "x")).unwrap_or(0),
                            ch_y: ch_off.and_then(|n| attr_i64(&n, "y")).unwrap_or(0),
                            ch_cx: ch_ext.and_then(|n| attr_i64(&n, "cx")).unwrap_or(0),
                            ch_cy: ch_ext.and_then(|n| attr_i64(&n, "cy")).unwrap_or(0),
                            flip_h: attr(&xfrm, "flipH")
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false),
                            flip_v: attr(&xfrm, "flipV")
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false),
                            rot: attr_f64(&xfrm, "rot").unwrap_or(0.0) / 60000.0,
                        }
                    });
                let group_start = out.len();
                let _ = (&empty_rels, &empty_smartart); // silence unused warnings when no nested picture
                for child_node in node.children().filter(|n| n.is_element()) {
                    match child_node.tag_name().name() {
                        "sp" => {
                            if let Some(mut shape) =
                                parse_shape(child_node, &empty_lph, theme, &empty_rels, None, zip)
                            {
                                shape.text_rect = parse_tx_xfrm(child_node);
                                out.push(SlideElement::Shape(shape));
                            }
                        }
                        "cxnSp" => {
                            if let Some(shape) = parse_connector(child_node, theme) {
                                out.push(SlideElement::Shape(shape));
                            }
                        }
                        _ => {}
                    }
                }
                if let Some(gt) = gt {
                    for el in &mut out[group_start..] {
                        apply_group_transform_to_element(el, &gt);
                    }
                }
            }
            _ => {}
        }
    }

    // Translate all emitted elements by the graphicFrame's position.
    for el in &mut out[start..] {
        offset_slide_element(el, gf_xfrm.x, gf_xfrm.y);
    }
}

fn offset_slide_element(el: &mut SlideElement, dx: i64, dy: i64) {
    match el {
        SlideElement::Shape(s) => {
            s.x += dx;
            s.y += dy;
            if let Some(tr) = &mut s.text_rect {
                tr.x += dx;
                tr.y += dy;
            }
        }
        SlideElement::Picture(p) => {
            p.x += dx;
            p.y += dy;
        }
        SlideElement::Table(t) => {
            t.x += dx;
            t.y += dy;
        }
        SlideElement::Chart(c) => {
            c.x += dx;
            c.y += dy;
        }
        SlideElement::Media(m) => {
            m.x += dx;
            m.y += dy;
        }
    }
}

/// Parse a connector shape (p:cxnSp) as a ShapeElement with line geometry.
fn parse_connector(
    node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<ShapeElement> {
    let sp_pr = child(node, "spPr")?;
    let xfrm = child(sp_pr, "xfrm")?;
    let t = parse_xfrm(xfrm);
    if t.cx == 0 && t.cy == 0 {
        return None;
    }
    let (cnv_id, cnv_name) = read_cnv_pr(node);

    // Style-based stroke fallback. `p:style > a:lnRef idx="N"` references the
    // theme fmtScheme lnStyleLst entry N (1-based). We look up the canonical
    // width from the theme map ("+lnRef-N") rather than hardcoding 9525.
    let style_node = child(node, "style");
    let style_stroke: Option<Stroke> = style_node.and_then(|s| child(s, "lnRef")).and_then(|lr| {
        let idx: u32 = attr(&lr, "idx").and_then(|v| v.parse().ok()).unwrap_or(1);
        if idx == 0 {
            None
        } else {
            parse_color_node(lr, theme).map(|c| {
                let width = theme
                    .get(&format!("+lnRef-{}", idx))
                    .and_then(|s| s.parse::<i64>().ok())
                    .unwrap_or(9525);
                Stroke {
                    color: c,
                    width,
                    dash_style: None,
                    head_end: None,
                    tail_end: None,
                    cmpd: None,
                }
            })
        }
    });

    // Merge <a:ln> attributes onto style_stroke: <a:ln> commonly contains only
    // arrow ends (headEnd/tailEnd) while the color/width come from lnRef.
    // parse_stroke() alone drops the whole stroke when solidFill is absent,
    // which leaves connectors invisible on slides that rely on style_ref.
    let ln_node = child(sp_pr, "ln");
    let stroke: Option<Stroke> = match ln_node {
        None => style_stroke,
        Some(ln) => {
            if child(ln, "noFill").is_some() {
                None
            } else {
                let ln_width = attr_i64(&ln, "w");
                let ln_color = child(ln, "solidFill").and_then(|n| parse_color_node(n, theme));
                let ln_dash = child(ln, "prstDash")
                    .and_then(|n| attr(&n, "val"))
                    .filter(|v| v != "solid");
                let ln_head = child(ln, "headEnd")
                    .map(parse_arrow_end)
                    .filter(|a| a.kind != "none");
                let ln_tail = child(ln, "tailEnd")
                    .map(parse_arrow_end)
                    .filter(|a| a.kind != "none");
                let ln_cmpd = attr(&ln, "cmpd").filter(|v| v != "sng");
                match (ln_color, style_stroke) {
                    (Some(c), base) => Some(Stroke {
                        color: c,
                        width: ln_width
                            .unwrap_or_else(|| base.as_ref().map(|s| s.width).unwrap_or(9525)),
                        dash_style: ln_dash
                            .or_else(|| base.as_ref().and_then(|s| s.dash_style.clone())),
                        head_end: ln_head
                            .or_else(|| base.as_ref().and_then(|s| s.head_end.clone())),
                        tail_end: ln_tail
                            .or_else(|| base.as_ref().and_then(|s| s.tail_end.clone())),
                        cmpd: ln_cmpd.or_else(|| base.as_ref().and_then(|s| s.cmpd.clone())),
                    }),
                    (None, Some(base)) => Some(Stroke {
                        color: base.color,
                        width: ln_width.unwrap_or(base.width),
                        dash_style: ln_dash.or(base.dash_style),
                        head_end: ln_head.or(base.head_end),
                        tail_end: ln_tail.or(base.tail_end),
                        cmpd: ln_cmpd.or(base.cmpd),
                    }),
                    (None, None) => None,
                }
            }
        }
    };

    let shadow = child(sp_pr, "effectLst").and_then(|n| parse_shadow(n, theme));

    let cy = if t.cy == 0 { 1 } else { t.cy };

    // Preserve the actual preset geometry (bentConnector3, curvedConnector4, …)
    // rather than collapsing every p:cxnSp to "line" — otherwise the renderer
    // can't distinguish bent/curved paths from a straight segment.
    let prst_geom_node = child(sp_pr, "prstGeom");
    let geometry = prst_geom_node
        .and_then(|n| attr(&n, "prst"))
        .unwrap_or_else(|| "line".to_owned());

    // Pull connector adjust values from avLst (e.g. bentConnector3 adj1 = bend %).
    let parse_gd_val = |gd: roxmltree::Node<'_, '_>| -> Option<f64> {
        attr(&gd, "fmla")
            .and_then(|f| f.strip_prefix("val ").map(|s| s.to_owned()))
            .and_then(|s| s.parse::<f64>().ok())
    };
    let av_node = prst_geom_node.and_then(|n| child(n, "avLst"));
    let gd_nodes: Vec<_> = av_node
        .map(|av| {
            av.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "gd")
                .collect()
        })
        .unwrap_or_default();
    let adj: Option<f64> = gd_nodes
        .iter()
        .find(|n| matches!(attr(n, "name").as_deref(), Some("adj") | Some("adj1")))
        .or_else(|| gd_nodes.first())
        .and_then(|n| parse_gd_val(*n));
    let adj2: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj2"))
        .or_else(|| gd_nodes.get(1))
        .and_then(|n| parse_gd_val(*n));
    let adj3: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj3"))
        .or_else(|| gd_nodes.get(2))
        .and_then(|n| parse_gd_val(*n));
    let adj4: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj4"))
        .or_else(|| gd_nodes.get(3))
        .and_then(|n| parse_gd_val(*n));

    Some(ShapeElement {
        x: t.x,
        y: t.y,
        width: t.cx,
        height: cy,
        rotation: t.rot,
        flip_h: t.flip_h,
        flip_v: t.flip_v,
        geometry,
        fill: None,
        stroke,
        text_body: None,
        default_text_color: None,
        cust_geom: None,
        adj,
        adj2,
        adj3,
        adj4,
        adj5: None,
        adj6: None,
        adj7: None,
        adj8: None,
        shadow,
        inner_shadow: None,
        glow: None,
        soft_edge: None,
        reflection: None,
        id: cnv_id,
        name: cnv_name,
        placeholder_type: None,
        placeholder_idx: None,
        text_rect: None,
        scene3d: parse_scene3d(sp_pr),
        sp3d: parse_sp3d(sp_pr),
    })
}

// ===========================
//  Presentation parser
// ===========================

/// All slide-master-derived data plus the master's effective theme, bundled so
/// it can be computed once per master and reused across every slide that shares
/// that master (ECMA-376 §19.3.1.42 — a deck may have multiple masters, each
/// with its own theme/clrMap). Resolving theme/master per slide via the
/// slide→slideLayout→slideMaster→theme rels chain is required so that scheme
/// colors (e.g. `<a:schemeClr val="accent1">`) pick the right palette.
struct MasterBundle {
    /// The master's effective theme palette, with the master's `<p:clrMap>`
    /// pre-baked (logical names → slot hex). Includes font/line/objectDefault
    /// keys exactly as `parse_theme_colors` produced them.
    theme: HashMap<String, String>,
    master_xml: Option<String>,
    master_rels: HashMap<String, String>,
    master_dir: String,
    master_smartart_drawings: HashMap<String, String>,
    master_bg: Option<Fill>,
    /// The master's own decorative (non-placeholder) spTree shapes, resolved ONCE
    /// against the master's baked `theme` (§19.3.1.38 showMasterSp). Each slide
    /// composites these beneath its content; pre-extracting here (per cached
    /// master) removes the per-slide master-XML re-parse + spTree re-walk (D4).
    /// A slide with a `<p:clrMapOvr>` re-resolves them against its override theme
    /// (see `parse_slide`), so these frozen-against-master-theme elements are used
    /// only by the common no-override slides.
    master_decorative: Vec<SlideElement>,
    master_font_sizes: HashMap<String, f64>,
    master_level_font_sizes: HashMap<String, LevelFontSizes>,
    master_level_indents: HashMap<String, LevelIndents>,
    master_level_bullets: HashMap<String, LevelBullets>,
    master_anchors: HashMap<String, String>,
    master_transforms: HashMap<String, Transform>,
    master_alignments: HashMap<String, String>,
    master_ea_ln_brk: HashMap<String, bool>,
    master_space_before: HashMap<String, i64>,
    master_space_after: HashMap<String, i64>,
    master_line_spacing: HashMap<String, f64>,
    master_bold: HashMap<String, bool>,
    master_italic: HashMap<String, bool>,
    master_caps: HashMap<String, String>,
    master_color: HashMap<String, String>,
}

/// The subset of `MasterBundle` fields that are THEME-DEPENDENT, recomputed for a
/// slide whose `<p:clrMapOvr><a:overrideClrMapping>` (ECMA-376 §19.3.1.7) replaces
/// the master's color mapping for the WHOLE slide (§20.1.6.8). `build_master_bundle`
/// freezes these against the MASTER's own clrMap-baked theme; for an override slide
/// we re-resolve them against the slide's effective mapping so that master-INHERITED
/// scheme colors (a `<p:bg>` schemeClr, master txStyles placeholder colors, master
/// bullet colors) flip together with the slide's own shapes. Owns all its data and
/// holds no `zip` borrow, so it can be built before `parse_slide(zip)` is called.
struct EffectiveMaster {
    /// `bundle.theme` clone with the override mapping applied (logical → slot hex).
    theme: HashMap<String, String>,
    /// Master `<p:bg>` re-resolved against `theme` (replaces `MasterBundle.master_bg`).
    master_bg: Option<Fill>,
    /// Master txStyles placeholder colors re-resolved against `theme`.
    master_color: HashMap<String, String>,
    /// Master per-level bullet colors re-resolved against `theme`.
    master_level_bullets: HashMap<String, LevelBullets>,
}

/// Build a `MasterBundle` for the master at `master_path` (a ZIP path such as
/// `ppt/slideMasters/slideMaster2.xml`). Reads the master XML + its rels,
/// resolves the master's own `/theme` relationship, parses the theme colors,
/// bakes the master's `<p:clrMap>`, then computes every master-derived map.
///
/// `fallback_theme` is the presentation-level theme used only when the master
/// has no `/theme` relationship of its own (keeps simple single-theme decks and
/// malformed packages working).
///
/// TODO: themeOverride (slide/layout `/themeOverride`, ECMA-376 §14.2.7) is not
/// yet honored — overrides on the layout or slide would replace parts of the
/// master theme. Out of scope for per-slide master resolution.
fn build_master_bundle(
    master_path: &str,
    fallback_theme: &HashMap<String, String>,
    zip: &mut PptxZip,
) -> MasterBundle {
    let master_xml_opt: Option<String> = read_zip_str(zip, master_path).ok();

    let master_dir: String = master_path
        .rsplit_once('/')
        .map(|(dir, _)| dir.to_owned())
        .unwrap_or_else(|| "ppt/slideMasters".to_owned());

    // Master rels: `<master_dir>/_rels/<file>.rels`.
    let master_file = master_path
        .split('/')
        .next_back()
        .unwrap_or("slideMaster1.xml");
    let master_rels_xml: String = {
        let rels_p = format!("{master_dir}/_rels/{master_file}.rels");
        read_zip_str(zip, &rels_p).unwrap_or_default()
    };
    let master_rels: HashMap<String, String> = parse_rels(&master_rels_xml);

    // The master's own theme (slide→…→slideMaster→theme). Fall back to the
    // presentation theme when the master declares no /theme relationship.
    let theme_path: Option<String> =
        find_rel_target_by_type(&master_rels_xml, "/theme").map(|t| resolve_path(&master_dir, &t));
    let mut theme: HashMap<String, String> = match theme_path
        .as_deref()
        .and_then(|p| read_zip_str(zip, p).ok())
    {
        Some(theme_xml) => parse_theme_colors(&theme_xml),
        None => fallback_theme.clone(),
    };
    // Bake the master's <p:clrMap> logical-name → slot mapping into the theme.
    bake_clr_map(&mut theme, master_xml_opt.as_deref());

    let master_smartart_drawings: HashMap<String, String> =
        build_smartart_drawings(&master_rels_xml, zip);

    // Parse the master XML EXACTLY ONCE and share the resulting `Document` across
    // every master-derived extractor below (D4: previously each `parse_master_*`
    // re-ran `Document::parse` on the same string, so a single master cost 12
    // parses — 11 extractors + the background). The `Document` borrows
    // `master_xml_opt`, so it lives only for the extraction scope; all owned maps
    // are computed before it is dropped. When the master has no XML (missing part)
    // every map defaults to empty, exactly as the prior `Option::map` chain did.
    let master_doc: Option<roxmltree::Document<'_>> = master_xml_opt.as_deref().and_then(|xml| {
        note_layout_master_parse();
        roxmltree::Document::parse(xml).ok()
    });
    let master_root: Option<roxmltree::Node<'_, '_>> =
        master_doc.as_ref().map(|d| d.root_element());

    let master_bg: Option<Fill> = master_root.and_then(|root| {
        let c_sld = child(root, "cSld")?;
        let mut resolve = |rid: &str| -> Option<String> {
            let target = master_rels.get(rid)?;
            let path = resolve_path(&master_dir, target);
            // Existence check only — central-directory lookup, no inflate
            // (former `read_zip_bytes` decompressed the entry just to discard it).
            zip.index_for_name(&path)?;
            Some(path)
        };
        parse_background(c_sld, &theme, &mut resolve)
    });

    let master_font_sizes = master_root.map(parse_master_font_sizes).unwrap_or_default();
    let master_level_font_sizes = master_root
        .map(parse_master_level_font_sizes)
        .unwrap_or_default();
    let master_level_indents = master_root
        .map(parse_master_level_indents)
        .unwrap_or_default();
    let master_level_bullets = master_root
        .map(|root| parse_master_level_bullets(root, &theme, &master_rels, &master_dir, zip))
        .unwrap_or_default();
    let master_anchors = master_root.map(parse_master_anchors).unwrap_or_default();
    let master_transforms = master_root.map(parse_master_transforms).unwrap_or_default();
    let master_alignments = master_root.map(parse_master_alignments).unwrap_or_default();
    let master_ea_ln_brk = master_root.map(parse_master_ea_ln_brk).unwrap_or_default();
    let (master_space_before, master_space_after, master_line_spacing) = master_root
        .map(parse_master_txstyle_spacing)
        .unwrap_or_default();
    let (master_bold, master_italic, master_caps) = master_root
        .map(parse_master_txstyle_bold_italic)
        .unwrap_or_default();
    let master_color = master_root
        .map(|root| parse_master_txstyle_color(root, &theme))
        .unwrap_or_default();

    // Pre-extract the master's decorative (non-placeholder) spTree shapes ONCE,
    // resolved against the master's baked `theme`. Each slide clones these instead
    // of re-parsing the master XML and re-walking its spTree (D4; former
    // per-slide `parse_slide` inline walk). Uses the same shared `master_root` and
    // the master's own rels + smartart drawings, exactly as the old inline walk did.
    let mut master_decorative: Vec<SlideElement> = Vec::new();
    if let Some(root) = master_root {
        extract_decorative_shapes(
            root,
            &master_dir,
            &master_rels,
            &master_smartart_drawings,
            &theme,
            zip,
            &mut master_decorative,
        );
    }

    MasterBundle {
        theme,
        master_xml: master_xml_opt,
        master_rels,
        master_dir,
        master_smartart_drawings,
        master_bg,
        master_decorative,
        master_font_sizes,
        master_level_font_sizes,
        master_level_indents,
        master_level_bullets,
        master_anchors,
        master_transforms,
        master_alignments,
        master_ea_ln_brk,
        master_space_before,
        master_space_after,
        master_line_spacing,
        master_bold,
        master_italic,
        master_caps,
        master_color,
    }
}

/// Parse a presentation from raw archive bytes. Thin wrapper that opens a fresh
/// owned [`PptxZip`] (copying `data`) and delegates to [`parse_presentation`].
/// Kept so the free `parse_pptx` / `pptx_to_markdown` WASM entry points and the
/// native `parse_pptx_native` path keep their `&[u8]` signature; the stateful
/// `PptxArchive` handle calls [`parse_presentation`] directly on its retained
/// archive to avoid re-opening it per call.
fn parse_presentation_from_bytes(data: &[u8]) -> Result<Presentation, Box<dyn std::error::Error>> {
    let mut zip = zip::ZipArchive::new(Cursor::new(data.to_vec()))?;
    parse_presentation(&mut zip)
}

fn parse_presentation(zip: &mut PptxZip) -> Result<Presentation, Box<dyn std::error::Error>> {
    // --- presentation.xml ---
    let pres_xml = read_zip_str(zip, "ppt/presentation.xml")?;
    let pres_doc = roxmltree::Document::parse(&pres_xml)?;
    let pres_root = pres_doc.root_element();

    let sld_sz = child(pres_root, "sldSz");
    let slide_width = sld_sz.and_then(|n| attr_i64(&n, "cx")).unwrap_or(9_144_000);
    let slide_height = sld_sz.and_then(|n| attr_i64(&n, "cy")).unwrap_or(6_858_000);

    // Ordered slide rIds
    let slide_rids: Vec<String> = child(pres_root, "sldIdLst")
        .map(|lst| {
            children_vec(lst, "sldId")
                .into_iter()
                .filter_map(|n| attr_r(&n, "id"))
                .collect()
        })
        .unwrap_or_default();

    // --- ppt/_rels/presentation.xml.rels ---
    let pres_rels_xml = read_zip_str(zip, "ppt/_rels/presentation.xml.rels")?;
    let pres_rels = parse_rels(&pres_rels_xml);

    // --- Presentation-level theme colors ---
    // Used for the deck-wide defaults on `Presentation` (default text color,
    // major/minor fonts, hyperlink colors) and as the fallback theme for any
    // master that declares no /theme relationship of its own.
    let theme_xml = find_rel_target_by_type(&pres_rels_xml, "/theme")
        .map(|t| resolve_path("ppt", &t))
        .and_then(|path| read_zip_str(zip, &path).ok())
        .unwrap_or_default();
    let theme = parse_theme_colors(&theme_xml);

    // --- Presentation-level fallback master ---
    // The first slide master referenced by the presentation. Used for slides
    // whose layout→master→theme chain can't be resolved (simple/old decks), so
    // their behavior is unchanged from before per-slide resolution existed.
    let pres_master_path: Option<String> =
        find_rel_target_by_type(&pres_rels_xml, "/slideMaster").map(|t| resolve_path("ppt", &t));

    // Cache of MasterBundle keyed by master ZIP path. Slides sharing a master
    // reuse the bundle instead of recomputing every master-derived map. Seed it
    // with the presentation master so the slide loop reuses the fallback build
    // instead of rebuilding it — for a single-master deck every slide resolves
    // to this same master, and without seeding the fallback build and the first
    // slide's cache-miss build would each compute the identical bundle twice.
    let mut master_cache: HashMap<String, MasterBundle> = HashMap::new();
    // Bundle for the truly no-master / empty-path case (no /slideMaster rel on
    // the presentation). Only built then; otherwise the fallback is the cached
    // presentation-master bundle.
    let no_master_bundle: Option<MasterBundle> = match pres_master_path.as_deref() {
        Some(p) => {
            master_cache.insert(p.to_owned(), build_master_bundle(p, &theme, zip));
            None
        }
        None => Some(build_master_bundle("", &theme, zip)),
    };

    // Cache of the layout single-pass extraction (`ParsedLayout`) keyed by layout
    // ZIP path (D4), mirroring `master_cache`. Slides sharing a layout reuse its
    // resolved placeholders + layout background + showMasterSp instead of
    // re-parsing the layout XML four times per slide. Only NO-override slides
    // populate/read the cache: the entry is resolved against the master's baked
    // theme, and a slide's layout→master chain is 1:1 (a layout names exactly one
    // master), so every no-override slide on a given layout shares that theme. A
    // slide with a `<p:clrMapOvr>` builds a fresh `ParsedLayout` against its
    // override theme instead (kept out of the cache).
    let mut layout_cache: HashMap<String, ParsedLayout> = HashMap::new();

    // Pre-collect slide XMLs, their rels, the layout XML, and layout rels
    struct SlideRaw {
        index: usize,
        slide_xml: String,
        slide_rels: HashMap<String, String>,
        smartart_drawings: HashMap<String, String>,
        /// ZIP path of the slide's layout (e.g. `ppt/slideLayouts/slideLayout3.xml`).
        /// Used as the `layout_cache` key so slides sharing a layout reuse its
        /// single-pass `ParsedLayout` (D4). `None` when the slide has no layout.
        layout_path: Option<String>,
        layout_xml: Option<String>,
        layout_rels: HashMap<String, String>,
        layout_dir: String,
        /// ZIP path of the slide's effective master, resolved through the
        /// slide→slideLayout→slideMaster rels chain. `None` when the chain
        /// can't be followed (no layout, or the layout has no /slideMaster
        /// relationship); such slides fall back to the presentation master.
        master_path: Option<String>,
    }

    let mut raw_slides: Vec<SlideRaw> = Vec::new();

    for (idx, r_id) in slide_rids.iter().enumerate() {
        let rel_target = match pres_rels.get(r_id) {
            Some(t) => t.clone(),
            None => continue,
        };
        // Resolve via `resolve_path` (not `format!("ppt/{rel_target}")`) so a
        // package-root-absolute slide Target — e.g. `/ppt/slides/slide1.xml`
        // (leading slash, OPC / ECMA-376 Part 2 §9.3) — resolves correctly
        // instead of producing `ppt//ppt/slides/slide1.xml`. Relative targets
        // (the common `slides/slide1.xml`) are unaffected. Same fix class as
        // the chart-rel resolution above.
        let slide_path = resolve_path("ppt", &rel_target);
        let slide_file = rel_target
            .split('/')
            .next_back()
            .unwrap_or("slide.xml")
            .to_owned();
        let rels_path = format!("ppt/slides/_rels/{slide_file}.rels");

        let slide_xml = read_zip_str(zip, &slide_path)?;
        let slide_rels_xml = read_zip_str(zip, &rels_path).unwrap_or_default();
        let slide_rels = parse_rels(&slide_rels_xml);
        let smartart_drawings = build_smartart_drawings(&slide_rels_xml, zip);

        // Layout XML
        let layout_path = find_rel_target_by_type(&slide_rels_xml, "/slideLayout")
            .map(|target| resolve_path("ppt/slides", &target));

        let layout_xml = layout_path
            .as_deref()
            .and_then(|path| read_zip_str(zip, path).ok());

        let layout_dir = layout_path
            .as_deref()
            .and_then(|p| p.rsplit_once('/').map(|(dir, _)| dir.to_owned()))
            .unwrap_or_else(|| "ppt/slideLayouts".to_owned());

        // Layout rels XML — needed both to resolve images inside the layout and
        // to follow the layout→slideMaster relationship for per-slide master
        // resolution (ECMA-376 §19.3.1.43).
        let layout_rels_xml: String = layout_path
            .as_deref()
            .and_then(|path| {
                let file = path.split('/').next_back().unwrap_or("layout.xml");
                let rels_p = format!("ppt/slideLayouts/_rels/{file}.rels");
                read_zip_str(zip, &rels_p).ok()
            })
            .unwrap_or_default();
        let layout_rels = parse_rels(&layout_rels_xml);

        // Resolve this slide's master via the layout's /slideMaster rel.
        let master_path: Option<String> = find_rel_target_by_type(&layout_rels_xml, "/slideMaster")
            .map(|t| resolve_path(&layout_dir, &t));

        raw_slides.push(SlideRaw {
            index: idx,
            slide_xml,
            slide_rels,
            smartart_drawings,
            layout_path,
            layout_xml,
            layout_rels,
            layout_dir,
            master_path,
        });
    }

    let mut slides = Vec::new();
    for raw in &raw_slides {
        // Resolve this slide's MasterBundle: build (and cache) one for the
        // slide's own master when the layout→master chain resolved; otherwise
        // use the presentation-level fallback bundle. Building is keyed by
        // master path so slides sharing a master don't recompute.
        let bundle: &MasterBundle = match raw.master_path.as_deref() {
            Some(mp) if !mp.is_empty() => {
                if !master_cache.contains_key(mp) {
                    let b = build_master_bundle(mp, &theme, zip);
                    master_cache.insert(mp.to_owned(), b);
                }
                &master_cache[mp]
            }
            // Unresolved chain → presentation-level fallback. The presentation
            // master (when present) was seeded into the cache above, so reuse
            // that entry; only a deck with no /slideMaster rel falls through to
            // `no_master_bundle`.
            _ => pres_master_path
                .as_deref()
                .map(|p| &master_cache[p])
                .unwrap_or_else(|| no_master_bundle.as_ref().unwrap()),
        };

        // Per-slide color-mapping override (ECMA-376 §19.3.1.7 clrMapOvr).
        // Precedence: the slide's own `<a:overrideClrMapping>` wins; else the
        // layout's; else inherit the master (`None`). `<a:masterClrMapping/>`
        // and an absent `<p:clrMapOvr>` both yield `None` at their level, so a
        // slide that explicitly inherits still falls through to the layout's
        // override — matching the slide→layout→master mapping chain.
        //
        // Why `<a:masterClrMapping/>` means "inherit (the layout)", NOT "bypass the
        // layout and use the master directly": §20.1.6.6 says masterClrMapping uses
        // "the color mapping defined in the master", and §19.3.1.7 likewise "the
        // color scheme defined by the master is used". Read in isolation that sounds
        // like a slide-level bypass — but Annex L.3.2.5 ("Slide Layouts") defines a
        // layout's Color Map Override as one that "overrides the inherited color
        // mapping from the slide master but IS INHERITED BY ALL PRESENTATION SLIDES
        // that utilize this layout." So once a layout overrides the master mapping,
        // the layout's mapping *is* the effective parent mapping the slide inherits;
        // "the master's mapping" for a slide on that layout already means the layout-
        // overridden one. PowerPoint additionally serializes `<a:masterClrMapping/>`
        // on ordinary non-overriding slides, so reading it as a layout bypass would
        // break layout-override inheritance for the common case. Hence both
        // masterClrMapping and an absent clrMapOvr resolve to `None` here and fall
        // through to the layout's override (then the master).
        let clr_map_ovr: Option<HashMap<String, String>> = parse_clr_map_ovr(&raw.slide_xml)
            .or_else(|| raw.layout_xml.as_deref().and_then(parse_clr_map_ovr));
        // When an override applies, recompute the master's THEME-DEPENDENT fields
        // against the slide's effective mapping. §20.1.6.8 says the override is
        // used "in place of" the master's mapping for the whole slide, so master-
        // INHERITED scheme colors (the master `<p:bg>`, master txStyles placeholder
        // colors, master bullet colors) must flip together with the slide's own
        // shapes — not just the slide's effective `theme`.
        //
        // The effective theme is the master-baked theme with the override re-applied.
        // This is correct because `bake_clr_map` left the raw scheme slots
        // (dk1/lt1/dk2/lt2/accent1..6/hlink/folHlink) intact, so the override's slot
        // values resolve against the original palette (§20.1.6.8). The override
        // REPLACES the master's logical→slot mapping, not the master's already-baked
        // logical hexes (we re-apply over the raw slots).
        //
        // Documented limitation: if the master's clrMap non-identically remapped an
        // accent SLOT (e.g. accent1="accent2") AND an override targets that same
        // accent, the raw accent slot is still its own scheme value (bake only writes
        // logical keys), so the override resolves it from the intact scheme — correct.
        // The only unrecoverable case would be a master that *overwrote a raw slot key
        // itself*, which `bake_clr_map` never does.
        //
        // Built fully here (in `parse_presentation`, where `zip` is available — the
        // master `<p:bg>` may reference a blip) BEFORE `parse_slide`; `EffectiveMaster`
        // owns its data and holds no `zip` borrow, so the mutable borrow taken to
        // resolve `master_bg` ends before `parse_slide(zip)` is called.
        let effective_master: Option<EffectiveMaster> = clr_map_ovr.map(|ovr| {
            let mut theme = bundle.theme.clone();
            apply_clr_map(&mut theme, Some(&ovr));
            // Re-run the master's theme-dependent extractions (mirrors
            // build_master_bundle) against the effective override theme so master-
            // INHERITED scheme colors (the `<p:bg>` schemeClr, txStyles placeholder
            // colors, per-level bullet colors) flip with the override. Parse the
            // master XML ONCE here and share the root across all three re-resolutions
            // (previously each re-parsed the same string — 3 parses per override slide).
            let master_doc = bundle.master_xml.as_deref().and_then(|xml| {
                note_layout_master_parse();
                roxmltree::Document::parse(xml).ok()
            });
            let master_root = master_doc.as_ref().map(|d| d.root_element());
            let master_bg: Option<Fill> = master_root.and_then(|root| {
                let c_sld = child(root, "cSld")?;
                let mut resolve = |rid: &str| -> Option<String> {
                    let target = bundle.master_rels.get(rid)?;
                    let path = resolve_path(&bundle.master_dir, target);
                    // Existence check only — central-directory lookup, no inflate
                    // (former `read_zip_bytes` decompressed the entry to discard it).
                    zip.index_for_name(&path)?;
                    Some(path)
                };
                parse_background(c_sld, &theme, &mut resolve)
            });
            let master_color = master_root
                .map(|root| parse_master_txstyle_color(root, &theme))
                .unwrap_or_default();
            let master_level_bullets = master_root
                .map(|root| {
                    parse_master_level_bullets(
                        root,
                        &theme,
                        &bundle.master_rels,
                        &bundle.master_dir,
                        zip,
                    )
                })
                .unwrap_or_default();
            EffectiveMaster {
                theme,
                master_bg,
                master_color,
                master_level_bullets,
            }
        });

        // Resolve this slide's `ParsedLayout` (placeholders + layout bg +
        // showMasterSp), parsing the layout XML once. Only the `theme` and the
        // master bullet colors that `parse_layout` consumes are theme-dependent;
        // a clrMapOvr slide passes the OVERRIDE-adjusted pair so its layout colors
        // flip with the override (mirrors the master theme-dependent recompute
        // above), everything else is the frozen bundle maps.
        let (layout_theme, layout_master_bullets): (
            &HashMap<String, String>,
            &HashMap<String, LevelBullets>,
        ) = match effective_master.as_ref() {
            Some(e) => (&e.theme, &e.master_level_bullets),
            None => (&bundle.theme, &bundle.master_level_bullets),
        };
        // Build a `ParsedLayout` from a layout XML string with the resolved
        // theme/bullets and this bundle's remaining (theme-independent) maps.
        let build_parsed_layout = |lx: &str, zip: &mut PptxZip| -> ParsedLayout {
            parse_layout(
                lx,
                &bundle.master_font_sizes,
                &bundle.master_level_font_sizes,
                &bundle.master_level_indents,
                layout_master_bullets,
                &bundle.master_anchors,
                &bundle.master_transforms,
                &bundle.master_alignments,
                &bundle.master_ea_ln_brk,
                &bundle.master_space_before,
                &bundle.master_space_after,
                &bundle.master_line_spacing,
                layout_theme,
                &raw.layout_dir,
                &raw.layout_rels,
                zip,
            )
        };

        // No-override slide WITH a layout path → cache by layout path (its entry
        // is resolved against the master-baked theme, which every no-override
        // slide on this layout shares). Otherwise build a fresh, uncached one
        // (override slide, or the rare no-layout-path case).
        let fresh_layout: Option<ParsedLayout> = match (
            effective_master.is_none(),
            raw.layout_xml.as_deref(),
            raw.layout_path.as_deref(),
        ) {
            (true, Some(lx), Some(lp)) => {
                if !layout_cache.contains_key(lp) {
                    let pl = build_parsed_layout(lx, zip);
                    layout_cache.insert(lp.to_owned(), pl);
                }
                None // borrowed from the cache below
            }
            (_, Some(lx), _) => Some(build_parsed_layout(lx, zip)),
            (_, None, _) => Some(ParsedLayout::default()),
        };
        let parsed_layout: &ParsedLayout = match &fresh_layout {
            Some(pl) => pl,
            // Cached (no-override) path: `layout_path` is guaranteed present
            // because that is the only arm that leaves `fresh_layout` as `None`.
            None => &layout_cache[raw.layout_path.as_deref().unwrap()],
        };

        let slide = parse_slide(
            &raw.slide_xml,
            parsed_layout,
            raw.layout_xml.as_deref(),
            &raw.layout_rels,
            &raw.layout_dir,
            bundle,
            effective_master.as_ref(),
            raw.index,
            &raw.slide_rels,
            &raw.smartart_drawings,
            zip,
        )?;
        slides.push(slide);
    }

    let default_text_color = theme.get("dk1").cloned();
    let major_font = theme.get("+mj-lt").cloned();
    let minor_font = theme.get("+mn-lt").cloned();
    let hlink_color = theme.get("hlink").cloned();
    let fol_hlink_color = theme.get("folHlink").cloned();
    Ok(Presentation {
        slide_width,
        slide_height,
        slides,
        default_text_color,
        major_font,
        minor_font,
        hlink_color,
        fol_hlink_color,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use ooxml_common::math::nodes_to_text;

    // Local-only sample (redistribution-prohibited, gitignored). Tests that
    // depend on it must skip gracefully on a clean checkout / in CI where the
    // file is absent. See packages/pptx/public/private/.
    const LOCAL_SAMPLE_2: &str = "../public/private/sample-2.pptx";

    /// Build an empty in-memory zip — enough for parse_* functions that take a
    /// `&mut PptxZip` but whose input declares no `<a:buBlip>` / blipFill parts.
    fn empty_zip_bytes() -> Vec<u8> {
        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let writer = zip::ZipWriter::new(cursor);
            writer.finish().unwrap();
        }
        buf
    }

    /// Build an in-memory zip containing exactly `parts` (path → bytes). Used to
    /// prove a `<a:buBlip>` whose rId resolves to a part that ISN'T in the
    /// archive falls through to Bullet::Inherit (index_for_name returns None).
    fn zip_with_parts(parts: &[(&str, &[u8])]) -> Vec<u8> {
        use std::io::Write;
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = zip::write::SimpleFileOptions::default();
            for (path, bytes) in parts {
                w.start_file(*path, o).unwrap();
                w.write_all(bytes).unwrap();
            }
            w.finish().unwrap();
        }
        buf
    }

    #[test]
    fn slide_show_attr_marks_hidden() {
        let ns = r#"xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main""#;
        let parse = |attr: &str| {
            let xml = format!(r#"<p:sld {ns} {attr}><p:cSld/></p:sld>"#);
            let doc = roxmltree::Document::parse(&xml).unwrap();
            slide_is_hidden(doc.root_element())
        };
        // Absent `show` ⇒ shown (default true ⇒ not hidden).
        assert!(!parse(""));
        // `show="0"` / `show="false"` ⇒ hidden (ECMA-376 §19.3.1.38 CT_Slide).
        assert!(parse(r#"show="0""#));
        assert!(parse(r#"show="false""#));
        // Explicit truthy ⇒ shown.
        assert!(!parse(r#"show="1""#));
        assert!(!parse(r#"show="true""#));
    }

    #[test]
    fn resolve_path_handles_absolute_leading_slash_target() {
        // An OPC relationship Target may be a package-root-absolute part name
        // (leading "/"), e.g. `/ppt/charts/chart5.xml` as emitted by some
        // generators. It must resolve from the package root and ignore the
        // source part's directory (ECMA-376 Part 2 / OPC §9.3). Regression for
        // issue #556 where a chart with an absolute Target silently failed to
        // load (read_zip_str on `ppt/slides/ppt/charts/chart5.xml`) and the
        // slide rendered the chart as a blank area.
        assert_eq!(
            resolve_path("ppt/slides", "/ppt/charts/chart5.xml"),
            "ppt/charts/chart5.xml"
        );
        // Relative references are unaffected by the absolute-target handling.
        assert_eq!(
            resolve_path("ppt/slides", "../charts/chart1.xml"),
            "ppt/charts/chart1.xml"
        );
        assert_eq!(
            resolve_path("ppt/slideLayouts", "../slideMasters/slideMaster1.xml"),
            "ppt/slideMasters/slideMaster1.xml"
        );
    }

    #[test]
    fn resolve_path_resolves_slide_targets_from_package_root() {
        // Slide parts are resolved from the presentation rels with base "ppt".
        // The common Target is relative (`slides/slide1.xml`); a generator may
        // also emit a package-root-absolute Target (`/ppt/slides/slide1.xml`),
        // which must NOT become `ppt//ppt/slides/slide1.xml`. Guards the
        // `resolve_path("ppt", rel_target)` slide-loading path.
        assert_eq!(
            resolve_path("ppt", "slides/slide1.xml"),
            "ppt/slides/slide1.xml"
        );
        assert_eq!(
            resolve_path("ppt", "/ppt/slides/slide1.xml"),
            "ppt/slides/slide1.xml"
        );
    }

    #[test]
    fn legacy_chart_parses_multi_level_category_axis() {
        // A `<c:cat>` may carry its labels in a `<c:multiLvlStrCache>` (multi-
        // level category axis, ECMA-376 §21.2.2.95) whose `<c:pt>` live under
        // `<c:lvl>` children rather than directly under the cache. Before the
        // fix, category extraction only recognized strCache/numCache/strLit/
        // numLit, so categories came back empty; that collapsed the shared
        // point count to 1 (`categories.len().max(1)`) and truncated EVERY
        // series to a single value. For an area chart a single point is a zero-
        // width sliver — i.e. a blank plot (issue #556).
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
 <c:chart><c:plotArea><c:areaChart>
  <c:ser>
   <c:cat><c:multiLvlStrRef><c:f>S!$A$1:$A$3</c:f><c:multiLvlStrCache>
     <c:ptCount val="3"/>
     <c:lvl><c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt><c:pt idx="2"><c:v>Mar</c:v></c:pt></c:lvl>
   </c:multiLvlStrCache></c:multiLvlStrRef></c:cat>
   <c:val><c:numRef><c:f>S!$B$1:$B$3</c:f><c:numCache>
     <c:ptCount val="3"/>
     <c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt>
   </c:numCache></c:numRef></c:val>
  </c:ser>
 </c:areaChart></c:plotArea></c:chart>
</c:chartSpace>"#;
        let chart = parse_legacy_chart(xml, &HashMap::new())
            .expect("area chart with multi-level cat should parse");
        assert_eq!(chart.chart.chart_type, "area");
        assert_eq!(chart.chart.categories, vec!["Jan", "Feb", "Mar"]);
        assert_eq!(chart.chart.series.len(), 1);
        assert_eq!(
            chart.chart.series[0].values,
            vec![Some(10.0), Some(20.0), Some(30.0)],
            "all three points must survive — a multi-level cat must not truncate the series"
        );
    }

    #[test]
    fn legacy_chart_parses_per_point_dpt_colors_for_pie() {
        // `<c:dPt>` (§21.2.2.52) carries its point index in a CHILD `<c:idx val>`
        // element (ECMA-376 §21.2.2.84, CT_UnsignedInt), NOT an attribute on it.
        // Reading it as an attribute always missed, so every pie/doughnut slice
        // fell back to the series colour (a `<a:schemeClr>` that resolved to the
        // default accent) — issue #556 follow-up: slide-7 fills were wrong. The
        // dPt fill must come from `<c:spPr><a:solidFill>`, never the border
        // `<a:ln><a:solidFill>`.
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
 <c:chart><c:plotArea><c:pieChart>
  <c:ser>
   <c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr>
   <c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="0D9488"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="F9F9F9"/></a:solidFill></a:ln></c:spPr></c:dPt>
   <c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="14B8A6"/></a:solidFill></c:spPr></c:dPt>
   <c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:strRef></c:cat>
   <c:val><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>60</c:v></c:pt><c:pt idx="1"><c:v>40</c:v></c:pt></c:numCache></c:numRef></c:val>
  </c:ser>
 </c:pieChart></c:plotArea></c:chart>
</c:chartSpace>"#;
        let chart = parse_legacy_chart(xml, &HashMap::new()).expect("pie should parse");
        assert_eq!(chart.chart.chart_type, "pie");
        let dpc = chart.chart.series[0]
            .data_point_colors
            .as_ref()
            .expect("per-slice dPt colours must be captured");
        assert_eq!(
            dpc,
            &vec![Some("0D9488".to_string()), Some("14B8A6".to_string())],
            "each slice takes its own <c:dPt><c:spPr> fill, not the series colour or the border"
        );
    }

    #[test]
    fn legacy_chart_data_labels_on_for_show_percent_only() {
        // Pie/doughnut decks commonly enable `<c:showPercent>` with
        // `<c:showVal val="0">` (ECMA-376 §21.2.2.187 / §21.2.2.189). The old
        // check looked at `showVal` alone, so the "54%/27%/…" slice labels in
        // sample-14 slide-7 never rendered. `show_data_labels` must be true
        // when EITHER flag is set.
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
 <c:chart><c:plotArea><c:pieChart>
  <c:dLbls><c:numFmt formatCode="0%" sourceLinked="0"/><c:showVal val="0"/><c:showPercent val="1"/><c:showCatName val="0"/></c:dLbls>
  <c:ser>
   <c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:strRef></c:cat>
   <c:val><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>60</c:v></c:pt><c:pt idx="1"><c:v>40</c:v></c:pt></c:numCache></c:numRef></c:val>
  </c:ser>
 </c:pieChart></c:plotArea></c:chart>
</c:chartSpace>"#;
        let chart = parse_legacy_chart(xml, &HashMap::new()).expect("pie should parse");
        assert!(
            chart.chart.show_data_labels,
            "showPercent=1 must enable data labels even when showVal=0"
        );
    }

    #[test]
    fn extract_image_reads_entry() {
        use std::io::{Cursor, Write};
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = zip::write::SimpleFileOptions::default();
            w.start_file("ppt/media/i.png", o).unwrap();
            w.write_all(b"X").unwrap();
            w.finish().unwrap();
        }
        assert_eq!(extract_image(&buf, "ppt/media/i.png", None).unwrap(), b"X");
    }

    /// A `PictureElement` serializes its blip as a zip path + mime, never as an
    /// inlined base64 `data:` URL (lazy image-bytes pipeline, Stage 2.1).
    #[test]
    fn picture_element_serializes_path_not_data_url() {
        let pic = PictureElement {
            x: 0,
            y: 0,
            width: 100,
            height: 100,
            rotation: 0.0,
            flip_h: false,
            flip_v: false,
            image_path: "ppt/media/image1.png".to_owned(),
            mime_type: "image/png".to_owned(),
            svg_image_path: None,
            intrinsic_width_px: Some(64),
            intrinsic_height_px: Some(48),
            stroke: None,
            prst_geom: None,
            prst_adjust: None,
            src_rect: None,
            alpha: None,
            cust_geom: None,
            shadow: None,
            inner_shadow: None,
            glow: None,
            soft_edge: None,
            reflection: None,
            scene3d: None,
            sp3d: None,
        };
        let json = serde_json::to_string(&pic).unwrap();
        assert!(
            json.contains("\"imagePath\":\"ppt/media/image1.png\""),
            "expected camelCase imagePath; got {json}"
        );
        assert!(
            json.contains("\"mimeType\":\"image/png\""),
            "expected camelCase mimeType; got {json}"
        );
        assert!(
            json.contains("\"intrinsicWidthPx\":64") && json.contains("\"intrinsicHeightPx\":48"),
            "expected intrinsic size keys; got {json}"
        );
        assert!(
            !json.contains("\"dataUrl\""),
            "must not emit dataUrl; got {json}"
        );
        assert!(
            !json.contains(";base64,"),
            "must not inline base64; got {json}"
        );
    }

    /// A blip `Fill::Image` (the serialized core `ImageFill`) serializes a zip
    /// path + mime, never an inlined base64 `data:` URL (Stage 2.2).
    #[test]
    fn image_fill_serializes_path_not_data_url() {
        let fill = Fill::Image {
            image_path: "ppt/media/image2.jpeg".to_owned(),
            mime_type: "image/jpeg".to_owned(),
            fill_rect: None,
            tile: None,
            alpha: None,
        };
        let json = serde_json::to_string(&fill).unwrap();
        assert!(
            json.contains("\"imagePath\":\"ppt/media/image2.jpeg\""),
            "expected camelCase imagePath; got {json}"
        );
        assert!(
            json.contains("\"mimeType\":\"image/jpeg\""),
            "expected camelCase mimeType; got {json}"
        );
        assert!(
            json.contains("\"fillType\":\"image\""),
            "tag preserved; got {json}"
        );
        assert!(
            !json.contains("\"dataUrl\""),
            "must not emit dataUrl; got {json}"
        );
        assert!(
            !json.contains(";base64,"),
            "must not inline base64; got {json}"
        );
    }

    // Synthetic deck for placeholder alignment-inheritance regression
    // (real reproducer files can't be committed). slide_sp / layout_extra_sp are
    // injected into the slide spTree / layout spTree.
    // Default master txStyles: titleStyle centred, bodyStyle left, otherStyle right.
    const DEFAULT_TXSTYLES: &str = r#"<p:txStyles><p:titleStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr algn="l"><a:defRPr sz="2800"/></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr algn="r"><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle></p:txStyles>"#;

    fn build_align_pptx(slide_sp: &str, layout_extra_sp: &str, master_txstyles: &str) -> Vec<u8> {
        use std::io::{Cursor, Write};
        let layout = format!(
            r#"<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree>
                <p:sp><p:nvSpPr><p:cNvPr id="2" name="Body 1"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
                  <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>
                {layout_extra_sp}
              </p:spTree></p:cSld>
            </p:sldLayout>"#
        );
        let master = format!(
            r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp></p:spTree></p:cSld>{master_txstyles}</p:sldMaster>"#
        );
        let entries: &[(&str, String)] = &[
            ("ppt/presentation.xml", r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId2"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>"#.to_owned()),
            ("ppt/_rels/presentation.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>"#.to_owned()),
            ("ppt/slides/slide1.xml", format!(r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree>{slide_sp}</p:spTree></p:cSld></p:sld>"#)),
            ("ppt/slides/_rels/slide1.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>"#.to_owned()),
            ("ppt/slideLayouts/slideLayout1.xml", layout),
            ("ppt/slideLayouts/_rels/slideLayout1.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>"#.to_owned()),
            ("ppt/slideMasters/slideMaster1.xml", master),
            ("ppt/slideMasters/_rels/slideMaster1.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>"#.to_owned()),
            ("ppt/theme/theme1.xml", r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="t"><a:themeElements><a:clrScheme name="c"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="000000"/></a:dk2><a:lt2><a:srgbClr val="FFFFFF"/></a:lt2><a:accent1><a:srgbClr val="000000"/></a:accent1><a:accent2><a:srgbClr val="000000"/></a:accent2><a:accent3><a:srgbClr val="000000"/></a:accent3><a:accent4><a:srgbClr val="000000"/></a:accent4><a:accent5><a:srgbClr val="000000"/></a:accent5><a:accent6><a:srgbClr val="000000"/></a:accent6><a:hlink><a:srgbClr val="000000"/></a:hlink><a:folHlink><a:srgbClr val="000000"/></a:folHlink></a:clrScheme><a:fontScheme name="f"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme><a:fmtScheme name="s"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>"#.to_owned()),
        ];
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = zip::write::SimpleFileOptions::default();
            for (name, body) in entries {
                w.start_file(*name, o).unwrap();
                w.write_all(body.as_bytes()).unwrap();
            }
            w.finish().unwrap();
        }
        buf
    }

    fn first_para_alignment(data: &[u8]) -> String {
        let json = parse_pptx_native(data).unwrap();
        let v: serde_json::Value = serde_json::from_str(&json).unwrap();
        v["slides"][0]["elements"][0]["textBody"]["paragraphs"][0]["alignment"]
            .as_str()
            .unwrap_or("<none>")
            .to_owned()
    }

    // body placeholder (idx=1) with no explicit algn anywhere except master bodyStyle="l".
    const BODY_SP: &str = r#"<p:sp><p:nvSpPr><p:cNvPr id="5" name="Text Placeholder 5"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>x</a:t></a:r></a:p></p:txBody></p:sp>"#;

    // An unrelated centred typeless placeholder (idx=10) in the layout — the leak source.
    const TYPELESS_CTR_SP: &str = r#"<p:sp><p:nvSpPr><p:cNvPr id="9" name="Centered obj"/><p:cNvSpPr/><p:nvPr><p:ph idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr algn="ctr"/></a:lstStyle><a:p/></p:txBody></p:sp>"#;

    #[test]
    fn align_inherit_body_no_layout_algn_is_left() {
        // Master bodyStyle="l", nothing in slide/layout → must resolve to left.
        assert_eq!(
            first_para_alignment(&build_align_pptx(BODY_SP, "", DEFAULT_TXSTYLES)),
            "l"
        );
    }

    #[test]
    fn align_inherit_body_ignores_unrelated_typeless_center() {
        // Layout has an unrelated centred typeless placeholder (idx=10). The body
        // placeholder (idx=1) must NOT borrow it; resolves to master bodyStyle "l".
        assert_eq!(
            first_para_alignment(&build_align_pptx(
                BODY_SP,
                TYPELESS_CTR_SP,
                DEFAULT_TXSTYLES
            )),
            "l"
        );
    }

    #[test]
    fn align_inherit_body_idx_no_master_default_ignores_typeless_leak() {
        // Residual-leak guard: master has NO bodyStyle algn (so by_idx_alignment is
        // not seeded for the body slot) AND the layout has an unrelated centred
        // typeless placeholder. The idx-bearing body must still fall to the spec
        // default "l", never the typeless sibling's "ctr".
        let no_body_algn = r#"<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="2800"/></a:lvl1pPr></p:bodyStyle></p:txStyles>"#;
        assert_eq!(
            first_para_alignment(&build_align_pptx(BODY_SP, TYPELESS_CTR_SP, no_body_algn)),
            "l"
        );
    }

    #[test]
    fn align_inherit_title_from_master_txstyles_center() {
        // Master titleStyle="ctr", no algn in slide/layout title → resolves to ctr.
        let title_sp = r#"<p:sp><p:nvSpPr><p:cNvPr id="6" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>T</a:t></a:r></a:p></p:txBody></p:sp>"#;
        assert_eq!(
            first_para_alignment(&build_align_pptx(title_sp, "", DEFAULT_TXSTYLES)),
            "ctr"
        );
    }

    #[test]
    fn align_inherit_subtitle_from_master_bodystyle() {
        // subTitle (idx=2, no matching layout slot) routes through the bodyStyle
        // txStyles row via the type path → inherits "l".
        let sub_sp = r#"<p:sp><p:nvSpPr><p:cNvPr id="7" name="Subtitle 1"/><p:cNvSpPr/><p:nvPr><p:ph type="subTitle" idx="2"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>S</a:t></a:r></a:p></p:txBody></p:sp>"#;
        assert_eq!(
            first_para_alignment(&build_align_pptx(sub_sp, "", DEFAULT_TXSTYLES)),
            "l"
        );
    }

    #[test]
    fn test_parse_chartex() {
        let Ok(data) = std::fs::read(LOCAL_SAMPLE_2) else {
            eprintln!("skipping test_parse_chartex: local sample not found");
            return;
        };
        let cursor = std::io::Cursor::new(data.clone());
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        let mut xml = String::new();
        zip.by_name("ppt/charts/chartEx1.xml")
            .unwrap()
            .read_to_string(&mut xml)
            .unwrap();
        let theme = HashMap::new();
        let result = parse_chartex(&xml, &theme);
        println!("parse_chartex result: {:?}", result.is_some());
        if let Some(ref c) = result {
            println!("  chart_type: {}", c.chart.chart_type);
            println!("  categories: {:?}", c.chart.categories);
            println!("  series len: {}", c.chart.series.len());
            if !c.chart.series.is_empty() {
                println!("  values: {:?}", c.chart.series[0].values);
            }
            println!("  subtotal_indices: {:?}", c.chart.subtotal_indices);
        }
        assert!(result.is_some(), "parse_chartex should succeed");
    }

    #[test]
    fn test_slide8_chart_rid() {
        let Ok(data) = std::fs::read(LOCAL_SAMPLE_2) else {
            eprintln!("skipping test_slide8_chart_rid: local sample not found");
            return;
        };
        let cursor = std::io::Cursor::new(data.clone());
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        let mut slide_xml = String::new();
        zip.by_name("ppt/slides/slide8.xml")
            .unwrap()
            .read_to_string(&mut slide_xml)
            .unwrap();

        let doc = roxmltree::Document::parse(&slide_xml).unwrap();
        let root = doc.root_element();

        for gf in root
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "graphicFrame")
        {
            println!("Found graphicFrame");
            if let Some(gd) = gf
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "graphicData")
            {
                let uri = attr(&gd, "uri").unwrap_or_default();
                println!("  graphicData uri: {}", uri);
                if let Some(chart_node) = gd
                    .descendants()
                    .find(|n| n.is_element() && n.tag_name().name() == "chart")
                {
                    println!("  chart node found, tag: {:?}", chart_node.tag_name());
                    for a in chart_node.attributes() {
                        println!(
                            "  attr: name={} ns={:?} val={}",
                            a.name(),
                            a.namespace(),
                            a.value()
                        );
                    }
                    let rid = attr_r(&chart_node, "id");
                    println!("  attr_r id: {:?}", rid);
                }
            }
        }
    }

    #[test]
    fn test_slide8_full_parse() {
        let Ok(data) = std::fs::read(LOCAL_SAMPLE_2) else {
            eprintln!("skipping test_slide8_full_parse: local sample not found");
            return;
        };
        let pres = parse_presentation_from_bytes(&data).unwrap();
        let slide = &pres.slides[7]; // 0-indexed, slide 8
        println!("Slide 8 elements: {}", slide.elements.len());
        for (i, el) in slide.elements.iter().enumerate() {
            match el {
                SlideElement::Chart(c) => println!(
                    "  [{}] CHART type={} cats={}",
                    i,
                    c.chart.chart_type,
                    c.chart.categories.len()
                ),
                SlideElement::Shape(s) => println!("  [{}] shape x={}", i, s.x),
                SlideElement::Table(_) => println!("  [{}] table", i),
                SlideElement::Picture(_) => println!("  [{}] picture", i),
                SlideElement::Media(m) => println!("  [{}] media kind={}", i, m.media_kind),
            }
        }
    }

    #[test]
    fn test_slide8_chartex_pipeline() {
        let Ok(data) = std::fs::read(LOCAL_SAMPLE_2) else {
            eprintln!("skipping test_slide8_chartex_pipeline: local sample not found");
            return;
        };
        let cursor = std::io::Cursor::new(data.clone());
        let mut zip = zip::ZipArchive::new(cursor).unwrap();

        let mut rels_xml = String::new();
        zip.by_name("ppt/slides/_rels/slide8.xml.rels")
            .unwrap()
            .read_to_string(&mut rels_xml)
            .unwrap();
        let rels = parse_rels(&rels_xml);
        println!("rels: {:?}", rels);

        let chart_path = resolve_path("ppt/slides", "../charts/chartEx1.xml");
        println!("chart_path: {}", chart_path);

        let result = read_zip_str(&mut zip, &chart_path);
        println!("read_zip_str ok: {}", result.is_ok());

        if let Ok(chart_xml) = result {
            let theme = HashMap::new();
            let r = parse_chartex(&chart_xml, &theme);
            println!("parse_chartex: {:?}", r.is_some());
        }
    }

    /// ECMA-376 §21.1.2.3.5 — a:hlinkClick @r:id resolves via slide _rels Target.
    #[test]
    fn test_parse_run_hyperlink_resolves_rid() {
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><rPr lang="en-US"><hlinkClick r:id="rId7"/></rPr><t>Open site</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let r_node = doc.root_element();
        let theme = HashMap::new();
        let mut rels = HashMap::new();
        rels.insert("rId7".to_owned(), "https://example.com/".to_owned());

        let parsed = parse_run(r_node, None, &theme, &rels).expect("run should parse");
        assert_eq!(parsed.text, "Open site");
        assert_eq!(parsed.hyperlink.as_deref(), Some("https://example.com/"));
    }

    /// A run without hlinkClick should have hyperlink = None.
    #[test]
    fn test_parse_run_without_hyperlink_is_none() {
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr lang="en-US"/><t>plain</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let rels = HashMap::new();
        let parsed = parse_run(doc.root_element(), None, &theme, &rels).expect("run should parse");
        assert!(parsed.hyperlink.is_none());
    }

    /// hlinkClick with an unknown r:id should produce hyperlink = None
    /// rather than emitting a placeholder string.
    #[test]
    fn test_parse_run_hyperlink_unknown_rid_is_none() {
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><rPr lang="en-US"><hlinkClick r:id="rIdNope"/></rPr><t>x</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let rels = HashMap::new();
        let parsed = parse_run(doc.root_element(), None, &theme, &rels).expect("run should parse");
        assert!(parsed.hyperlink.is_none());
    }

    /// ECMA-376 §20.1.8.40 — pattFill produces a Fill::Pattern carrying the
    /// preset name and the resolved fg/bg colours.
    #[test]
    fn test_parse_fill_pattern_extracts_fg_bg_preset() {
        let xml = r#"<spPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <pattFill prst="pct25">
                <fgClr><srgbClr val="C00000"/></fgClr>
                <bgClr><srgbClr val="FFFF00"/></bgClr>
            </pattFill>
        </spPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let fill = parse_fill(doc.root_element(), &theme).expect("pattFill should resolve");
        match fill {
            Fill::Pattern { fg, bg, preset } => {
                assert_eq!(preset, "pct25");
                assert_eq!(fg.to_uppercase(), "C00000");
                assert_eq!(bg.to_uppercase(), "FFFF00");
            }
            other => panic!("expected Fill::Pattern, got {:?}", other),
        }
    }

    /// pattFill missing fg/bg colours should fall back to black/white rather
    /// than dropping the fill entirely — keeps shapes recognisable when the
    /// theme cannot resolve the slot.
    #[test]
    fn test_parse_fill_pattern_defaults_when_colors_missing() {
        let xml = r#"<spPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <pattFill prst="horz"/>
        </spPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let fill = parse_fill(doc.root_element(), &theme).expect("pattFill should still resolve");
        match fill {
            Fill::Pattern { fg, bg, preset } => {
                assert_eq!(preset, "horz");
                assert_eq!(fg.to_lowercase(), "000000");
                assert_eq!(bg.to_lowercase(), "ffffff");
            }
            other => panic!("expected Fill::Pattern, got {:?}", other),
        }
    }

    /// ECMA-376 §21.1.2.3.10 — strike="dblStrike" produces strike_double=true,
    /// while strike="sngStrike" leaves strike_double=false. The plain
    /// `strikethrough` flag is true in both cases.
    #[test]
    fn test_parse_run_strike_double_distinguishes_dbl() {
        let dbl = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr strike="dblStrike"/><t>x</t></r>"#;
        let sng = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr strike="sngStrike"/><t>x</t></r>"#;
        let none = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr/><t>x</t></r>"#;
        let theme = HashMap::new();
        let rels = HashMap::new();

        let doc_d = roxmltree::Document::parse(dbl).unwrap();
        let r_d = parse_run(doc_d.root_element(), None, &theme, &rels).unwrap();
        assert!(r_d.strikethrough && r_d.strike_double);

        let doc_s = roxmltree::Document::parse(sng).unwrap();
        let r_s = parse_run(doc_s.root_element(), None, &theme, &rels).unwrap();
        assert!(r_s.strikethrough && !r_s.strike_double);

        let doc_n = roxmltree::Document::parse(none).unwrap();
        let r_n = parse_run(doc_n.root_element(), None, &theme, &rels).unwrap();
        assert!(!r_n.strikethrough && !r_n.strike_double);
    }

    /// ECMA-376 §21.1.2.3.13 — cap="all" / "small" are passed through;
    /// cap="none" or omitted yields None so the field stays absent in JSON.
    #[test]
    fn test_parse_run_caps_attribute() {
        let theme = HashMap::new();
        let rels = HashMap::new();
        let cases = [
            ("all", Some("all")),
            ("small", Some("small")),
            ("none", None),
        ];
        for (val, expected) in cases {
            let xml = format!(
                r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr cap="{val}"/><t>x</t></r>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
            assert_eq!(r.caps.as_deref(), expected, "caps={val}");
        }
    }

    /// ECMA-376 §21.1.2.3.5 — rPr @spc encodes letter spacing in 100ths of a
    /// point; positive widens, negative tightens. Zero rounds away (None).
    #[test]
    fn test_parse_run_letter_spacing() {
        let theme = HashMap::new();
        let rels = HashMap::new();
        for (raw, expected) in [("100", Some(1.0)), ("-50", Some(-0.5)), ("0", None)] {
            let xml = format!(
                r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr spc="{raw}"/><t>x</t></r>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
            assert_eq!(r.letter_spacing, expected, "spc={raw}");
        }
    }

    /// ECMA-376 §20.1.8.21 — innerShdw shares the field shape of outerShdw
    /// (blurRad, dist, dir, color child). parse_inner_shadow should round-trip
    /// all of them, including the alphaModFix encoded as 8-char hex.
    #[test]
    fn test_parse_inner_shadow_extracts_fields() {
        let xml = r#"<effectLst xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <innerShdw blurRad="50800" dist="38100" dir="2700000">
                <srgbClr val="000000"><alphaModFix amt="50000"/></srgbClr>
            </innerShdw>
        </effectLst>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let s = parse_inner_shadow(doc.root_element(), &theme).expect("innerShdw should resolve");
        assert_eq!(s.blur, 50_800);
        assert_eq!(s.dist, 38_100);
        assert!((s.dir - 45.0).abs() < 0.001);
        assert!((s.alpha - 0.5).abs() < 0.01);
        assert_eq!(s.color.to_lowercase(), "000000");
    }

    /// ECMA-376 §20.1.8.14 + §20.1.8.58 + §20.1.8.30 — a `bgPr > blipFill`
    /// with a `stretch > fillRect` (incl. negative overscan edges) parses into
    /// `Fill::Image` carrying the resolved zip path + mime, the fractional
    /// fillRect, and the alphaModFix alpha. Mirrors sample-12 slide1's background.
    #[test]
    fn test_parse_background_blip_fill_stretch() {
        let xml = r#"<p:cSld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <p:bg><p:bgPr>
                <a:blipFill>
                    <a:blip r:embed="rId2"><a:alphaModFix amt="80000"/></a:blip>
                    <a:stretch><a:fillRect t="-9000" b="-9000"/></a:stretch>
                </a:blipFill>
                <a:effectLst/>
            </p:bgPr></p:bg>
        </p:cSld>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut resolve = |rid: &str| -> Option<String> {
            assert_eq!(rid, "rId2");
            Some("ppt/media/image1.jpeg".to_owned())
        };
        let fill = parse_background(doc.root_element(), &theme, &mut resolve)
            .expect("blip background should resolve to Fill::Image");
        match fill {
            Fill::Image {
                image_path,
                mime_type,
                fill_rect,
                tile,
                alpha,
            } => {
                assert_eq!(image_path, "ppt/media/image1.jpeg");
                assert_eq!(mime_type, "image/jpeg");
                let fr = fill_rect.expect("fillRect should be present");
                assert!((fr.t - (-0.09)).abs() < 1e-9, "t={}", fr.t);
                assert!((fr.b - (-0.09)).abs() < 1e-9, "b={}", fr.b);
                assert!(is_zero_f64(&fr.l) && is_zero_f64(&fr.r));
                assert!(tile.is_none(), "stretch fill must not carry tile");
                assert!((alpha.expect("alpha") - 0.8).abs() < 1e-6);
            }
            other => panic!("expected Fill::Image, got {other:?}"),
        }
    }

    /// ECMA-376 §20.1.8.14 + §20.1.8.58 — a `bgPr > blipFill` with `<a:tile>`
    /// parses into `Fill::Image` carrying a `TileInfo` (and no `fillRect`).
    /// tx/ty stay EMU, sx/sy convert ST_Percentage → fraction, flip/algn pass
    /// through verbatim.
    #[test]
    fn test_parse_background_blip_fill_tile() {
        let xml = r#"<p:cSld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <p:bg><p:bgPr>
                <a:blipFill>
                    <a:blip r:embed="rId2"/>
                    <a:tile tx="457200" ty="-228600" sx="50000" sy="75000" flip="xy" algn="ctr"/>
                </a:blipFill>
            </p:bgPr></p:bg>
        </p:cSld>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut resolve = |_: &str| -> Option<String> { Some("ppt/media/image1.png".to_owned()) };
        let fill = parse_background(doc.root_element(), &theme, &mut resolve)
            .expect("tiled blip background should resolve to Fill::Image");
        match fill {
            Fill::Image {
                fill_rect, tile, ..
            } => {
                assert!(fill_rect.is_none(), "tile fill must not carry fillRect");
                let t = tile.expect("tile should be present");
                assert_eq!(t.tx, 457_200);
                assert_eq!(t.ty, -228_600);
                assert!((t.sx - 0.5).abs() < 1e-9, "sx={}", t.sx);
                assert!((t.sy - 0.75).abs() < 1e-9, "sy={}", t.sy);
                assert_eq!(t.flip, "xy");
                assert_eq!(t.algn, "ctr");
            }
            other => panic!("expected Fill::Image, got {other:?}"),
        }
    }

    /// §20.1.8.58 defaults: a bare `<a:tile/>` yields tx/ty=0, sx/sy=1.0
    /// (100% native size), flip="none", algn="tl".
    #[test]
    fn test_parse_background_blip_fill_tile_defaults() {
        let xml = r#"<p:cSld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <p:bg><p:bgPr>
                <a:blipFill><a:blip r:embed="rId2"/><a:tile/></a:blipFill>
            </p:bgPr></p:bg>
        </p:cSld>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut resolve = |_: &str| -> Option<String> { Some("ppt/media/image1.png".to_owned()) };
        let fill = parse_background(doc.root_element(), &theme, &mut resolve)
            .expect("bare tile should still resolve to Fill::Image");
        match fill {
            Fill::Image { tile, .. } => {
                let t = tile.expect("tile should be present");
                assert_eq!(t.tx, 0);
                assert_eq!(t.ty, 0);
                assert!((t.sx - 1.0).abs() < 1e-9);
                assert!((t.sy - 1.0).abs() < 1e-9);
                assert_eq!(t.flip, "none");
                assert_eq!(t.algn, "tl");
            }
            other => panic!("expected Fill::Image, got {other:?}"),
        }
    }

    /// ECMA-376 §21.1.2.4.2 — a paragraph `<a:pPr><a:buBlip><a:blip r:embed>`
    /// resolves into `Bullet::Blip` carrying the blip's zip path + mime. The
    /// `<a:buSzPct val>` (§21.1.2.4.3, thousandths of a percent) becomes a plain
    /// percentage on the bullet.
    #[test]
    fn test_parse_bullet_blip_resolves_embed_and_size() {
        let xml = r#"<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <a:buSzPct val="80000"/>
            <a:buBlip><a:blip r:embed="rId5"/></a:buBlip>
        </a:pPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut resolve = |rid: &str| -> Option<String> {
            assert_eq!(rid, "rId5");
            Some("ppt/media/image3.png".to_owned())
        };
        match parse_bullet(Some(doc.root_element()), &theme, &mut resolve) {
            Bullet::Blip {
                image_path,
                mime_type,
                size_pct,
            } => {
                assert_eq!(image_path, "ppt/media/image3.png");
                assert_eq!(mime_type, "image/png");
                assert!((size_pct.expect("size_pct") - 80.0).abs() < 1e-9);
            }
            other => panic!("expected Bullet::Blip, got {other:?}"),
        }
    }

    /// §21.1.2.4.2 — with no `<a:buSzPct>` the picture bullet carries `None`
    /// size (renderer uses the spec default of 100%), and the mime tracks the
    /// resolved extension (jpeg here).
    #[test]
    fn test_parse_bullet_blip_default_size_and_mime() {
        let xml = r#"<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <a:buBlip><a:blip r:embed="rId2"/></a:buBlip>
        </a:pPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut resolve = |_: &str| -> Option<String> { Some("ppt/media/image1.jpeg".to_owned()) };
        match parse_bullet(Some(doc.root_element()), &theme, &mut resolve) {
            Bullet::Blip {
                image_path,
                mime_type,
                size_pct,
            } => {
                assert_eq!(image_path, "ppt/media/image1.jpeg");
                assert_eq!(mime_type, "image/jpeg");
                assert!(size_pct.is_none());
            }
            other => panic!("expected Bullet::Blip, got {other:?}"),
        }
    }

    /// §21.1.2.4 — the bullet element is an `xsd:choice`: an explicit
    /// `<a:buNone>` wins even when a stray `<a:buBlip>` is also present (the
    /// paragraph draws no marker). Mirrors the buNone-over-buChar precedence.
    #[test]
    fn test_parse_bullet_none_wins_over_blip() {
        let xml = r#"<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <a:buNone/>
            <a:buBlip><a:blip r:embed="rId2"/></a:buBlip>
        </a:pPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut resolve = |_: &str| -> Option<String> { Some("ppt/media/image1.png".to_owned()) };
        assert!(matches!(
            parse_bullet(Some(doc.root_element()), &theme, &mut resolve),
            Bullet::None
        ));
    }

    /// §21.1.2.4.2 — a `<a:buBlip>` whose `r:embed` cannot be resolved (dangling
    /// relationship) must NOT emit a half-built picture bullet. It falls through
    /// to `Bullet::Inherit` so a lower style tier can still supply a marker.
    #[test]
    fn test_parse_bullet_blip_dangling_embed_inherits() {
        let xml = r#"<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <a:buBlip><a:blip r:embed="rIdMissing"/></a:buBlip>
        </a:pPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut resolve = |_: &str| -> Option<String> { None };
        assert!(matches!(
            parse_bullet(Some(doc.root_element()), &theme, &mut resolve),
            Bullet::Inherit
        ));
    }

    /// §19.7.10 / §21.1.2.4.2 — a picture bullet declared on a master/list-style
    /// `<a:lvlNpPr>` is captured per level by `read_level_bullets`, so a slide
    /// paragraph at that level inherits the image marker.
    #[test]
    fn test_read_level_bullets_picks_up_bublip() {
        let xml = r#"<a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <a:lvl1pPr><a:buBlip><a:blip r:embed="rId9"/></a:buBlip></a:lvl1pPr>
            <a:lvl2pPr><a:buChar char="-"/></a:lvl2pPr>
        </a:lstStyle>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut resolve = |rid: &str| -> Option<String> {
            assert_eq!(rid, "rId9");
            Some("ppt/media/image7.png".to_owned())
        };
        let levels = read_level_bullets(doc.root_element(), &theme, &mut resolve);
        match &levels[0] {
            Some(Bullet::Blip { image_path, .. }) => {
                assert_eq!(image_path, "ppt/media/image7.png")
            }
            other => panic!("expected lvl1 Bullet::Blip, got {other:?}"),
        }
        assert!(matches!(&levels[1], Some(Bullet::Char { .. })));
        assert!(levels[2].is_none());
    }

    /// ECMA-376 §21.1.2.4.2 — a `<a:buBlip>` whose `r:embed` IS listed in the
    /// part's rels (so `resolve_path` succeeds) but whose target part is NOT in
    /// the package must NOT emit a `Bullet::Blip` carrying a dangling path. The
    /// resolver verifies part existence with `index_for_name`, so a missing part
    /// yields `None` and the level falls through to `Bullet::Inherit` (the empty
    /// `LevelBullets` slot), matching the variant's doc comment. Exercised
    /// end-to-end through `parse_master_level_bullets` (one of the now-`zip`-
    /// threaded entry points) against a real in-memory archive.
    #[test]
    fn master_bublip_listed_but_missing_part_inherits() {
        // bodyStyle lvl1 declares a picture bullet whose embed (rId7) IS in the
        // master rels, pointing at ppt/media/missing.png.
        let master = r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                                     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld><p:spTree/></p:cSld>
          <p:txStyles>
            <p:bodyStyle>
              <a:lvl1pPr><a:buBlip><a:blip r:embed="rId7"/></a:buBlip><a:defRPr sz="2000"/></a:lvl1pPr>
            </p:bodyStyle>
          </p:txStyles>
        </p:sldMaster>"#;
        let theme = HashMap::new();
        let mut master_rels = HashMap::new();
        // rId7 resolves (resolve_path succeeds) to ppt/media/missing.png.
        master_rels.insert("rId7".to_owned(), "../media/missing.png".to_owned());

        // Archive deliberately LACKS ppt/media/missing.png (it holds an unrelated
        // part so it isn't empty). index_for_name(missing.png) → None.
        let bytes = zip_with_parts(&[("ppt/media/other.png", b"\x89PNG")]);
        let cursor = Cursor::new(bytes.clone());
        let mut zip = zip::ZipArchive::new(cursor).unwrap();

        let master_doc = roxmltree::Document::parse(master).unwrap();
        let master_root = master_doc.root_element();
        let m = parse_master_level_bullets(
            master_root,
            &theme,
            &master_rels,
            "ppt/slideMasters",
            &mut zip,
        );
        // The listed-but-missing part must not produce a Blip anywhere. With only
        // a buBlip (no char/auto) at lvl1 and the part absent, the level resolves
        // to Inherit (None) and the bodyStyle contributes no usable bullet, so the
        // "body" key is never inserted (has_any_level_bullet is false).
        if let Some(body) = m.get("body") {
            assert!(
                !matches!(body[0], Some(Bullet::Blip { .. })),
                "missing part must not yield Bullet::Blip; got {:?}",
                body[0]
            );
            assert!(
                body.iter().all(|b| !matches!(b, Some(Bullet::Blip { .. }))),
                "no level may carry a dangling Bullet::Blip; got {body:?}"
            );
        }

        // Positive control: with the SAME rels but the part now PRESENT, the
        // bullet resolves to Bullet::Blip — proving the test distinguishes
        // presence from absence rather than always inheriting.
        let bytes_ok = zip_with_parts(&[("ppt/media/missing.png", b"\x89PNG")]);
        let cursor_ok = Cursor::new(bytes_ok.clone());
        let mut zip_ok = zip::ZipArchive::new(cursor_ok).unwrap();
        let m_ok = parse_master_level_bullets(
            master_root,
            &theme,
            &master_rels,
            "ppt/slideMasters",
            &mut zip_ok,
        );
        match m_ok.get("body").and_then(|b| b[0].clone()) {
            Some(Bullet::Blip { image_path, .. }) => {
                assert_eq!(image_path, "ppt/media/missing.png");
            }
            other => panic!("expected Bullet::Blip when part is present, got {other:?}"),
        }
    }

    /// ECMA-376 §21.1.2.3.16 — underline_style carries non-default underline
    /// values (dbl, dotted, wavy, …) verbatim. The plain bool stays true for
    /// any non-"none" value; "sng" and absent both leave underline_style None
    /// because the renderer's default is already a single line.
    #[test]
    fn test_parse_run_underline_style_passthrough() {
        let theme = HashMap::new();
        let rels = HashMap::new();
        let cases: &[(&str, bool, Option<&str>)] = &[
            ("none", false, None),
            ("sng", true, None),
            ("dbl", true, Some("dbl")),
            ("heavy", true, Some("heavy")),
            ("dotted", true, Some("dotted")),
            ("wavy", true, Some("wavy")),
            ("dashLong", true, Some("dashLong")),
        ];
        for (val, expected_bool, expected_style) in cases {
            let xml = format!(
                r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr u="{val}"/><t>x</t></r>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
            assert_eq!(r.underline, *expected_bool, "u={val}");
            assert_eq!(r.underline_style.as_deref(), *expected_style, "u={val}");
        }
    }

    /// ECMA-376 §21.1.2.3.20 — rPr > uFill > solidFill yields a per-run
    /// underline colour distinct from the text colour. uFillTx (or absent)
    /// leaves underline_color as None so the renderer falls back to text.
    #[test]
    fn test_parse_run_underline_color() {
        let theme = HashMap::new();
        let rels = HashMap::new();

        let with_ufill = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr u="sng"><uFill><solidFill><srgbClr val="FF0000"/></solidFill></uFill></rPr><t>x</t></r>"#;
        let doc = roxmltree::Document::parse(with_ufill).unwrap();
        let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
        assert_eq!(
            r.underline_color
                .as_deref()
                .map(str::to_uppercase)
                .as_deref(),
            Some("FF0000")
        );

        let with_ufilltx = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr u="sng"><uFillTx/></rPr><t>x</t></r>"#;
        let doc = roxmltree::Document::parse(with_ufilltx).unwrap();
        let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
        assert!(r.underline_color.is_none());
    }

    /// ECMA-376 §21.1.2.3.4 — rPr > highlight is a CT_Color (the marker /
    /// text-highlight colour). Unlike WordprocessingML's CT_Highlight (a fixed
    /// 16-name enum), the DrawingML highlight is any colour, so it must resolve
    /// through the same colour pipeline as solidFill: srgbClr literal,
    /// schemeClr via the theme/clrMap, plus alpha transforms (8-char hex).
    #[test]
    fn test_parse_run_highlight_srgb() {
        let theme = HashMap::new();
        let rels = HashMap::new();
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr><highlight><srgbClr val="FFFF00"/></highlight></rPr><t>x</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
        assert_eq!(
            r.highlight.as_deref().map(str::to_uppercase).as_deref(),
            Some("FFFF00")
        );
    }

    /// schemeClr highlight resolves through the theme map (same path as
    /// solidFill scheme colours), proving we did not hard-code a name table.
    #[test]
    fn test_parse_run_highlight_scheme_resolves_theme() {
        let rels = HashMap::new();
        let mut theme = HashMap::new();
        theme.insert("accent1".to_owned(), "E46970".to_owned());
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr><highlight><schemeClr val="accent1"/></highlight></rPr><t>x</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
        assert_eq!(
            r.highlight.as_deref().map(str::to_uppercase).as_deref(),
            Some("E46970")
        );
    }

    /// An alpha transform on the highlight colour yields 8-char RRGGBBAA, the
    /// same encoding the shared colour helper emits for translucent fills.
    #[test]
    fn test_parse_run_highlight_alpha_is_8char() {
        let theme = HashMap::new();
        let rels = HashMap::new();
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr><highlight><srgbClr val="00FF00"><alpha val="50000"/></srgbClr></highlight></rPr><t>x</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
        let hl = r
            .highlight
            .expect("highlight should resolve")
            .to_uppercase();
        assert_eq!(hl.len(), 8, "alpha < 1 → RRGGBBAA, got {hl}");
        assert!(hl.starts_with("00FF00"), "rgb preserved, got {hl}");
    }

    /// No highlight element → field stays None (omitted from JSON).
    #[test]
    fn test_parse_run_without_highlight_is_none() {
        let theme = HashMap::new();
        let rels = HashMap::new();
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr/><t>x</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
        assert!(r.highlight.is_none());
    }

    /// ECMA-376 §21.1.2.3.7 — rPr > ea sets a separate East Asian font.
    /// Resolves through the theme map: "+mn-ea" should expand to whatever
    /// the theme registered, while a literal name is preserved.
    #[test]
    fn test_parse_run_ea_typeface() {
        let rels = HashMap::new();
        let mut theme = HashMap::new();
        theme.insert("+mn-ea".to_owned(), "MS Mincho".to_owned());

        // Theme reference resolves through the map.
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr><ea typeface="+mn-ea"/></rPr><t>あ</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
        assert_eq!(r.font_family_ea.as_deref(), Some("MS Mincho"));

        // Literal name passes through unchanged.
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr><ea typeface="Yu Gothic"/></rPr><t>あ</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
        assert_eq!(r.font_family_ea.as_deref(), Some("Yu Gothic"));

        // Empty typeface is filtered out.
        let xml = r#"<r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><rPr><ea typeface=""/></rPr><t>あ</t></r>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let r = parse_run(doc.root_element(), None, &theme, &rels).unwrap();
        assert!(r.font_family_ea.is_none());
    }

    /// ECMA-376 §20.1.8.17 — glow has a single rad attribute and a colour
    /// child. parse_glow should preserve the radius and resolve alphaModFix.
    #[test]
    fn test_parse_glow_extracts_radius_and_color() {
        let xml = r#"<effectLst xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <glow rad="38100">
                <srgbClr val="FF0000"><alphaModFix amt="80000"/></srgbClr>
            </glow>
        </effectLst>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let g = parse_glow(doc.root_element(), &theme).expect("glow should resolve");
        assert_eq!(g.radius, 38_100);
        assert_eq!(g.color.to_uppercase(), "FF0000");
        assert!((g.alpha - 0.8).abs() < 0.01);
    }

    /// ECMA-376 §20.1.8.31 — softEdge has a single `rad` attribute in EMU.
    #[test]
    fn test_parse_soft_edge_extracts_radius() {
        let xml = r#"<effectLst xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <softEdge rad="63500"/>
        </effectLst>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let s = parse_soft_edge(doc.root_element()).expect("softEdge should resolve");
        assert_eq!(s.radius, 63_500);
    }

    /// ECMA-376 §20.1.8.27 — reflection: blur, dist, dir, stA/stPos/endA/endPos
    /// (1000ths of percent), sx/sy (1000ths of percent, sy negative for mirror).
    #[test]
    fn test_parse_reflection_attributes() {
        let xml = r#"<effectLst xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <reflection blurRad="6350" stA="50000" endA="0" endPos="35000" dist="50800" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>
        </effectLst>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let r = parse_reflection(doc.root_element()).expect("reflection should resolve");
        assert_eq!(r.blur, 6_350);
        assert_eq!(r.dist, 50_800);
        assert!((r.dir - 90.0).abs() < 0.001);
        assert!((r.st_a - 0.5).abs() < 0.01);
        assert!((r.end_a - 0.0).abs() < 0.01);
        assert!((r.end_pos - 0.35).abs() < 0.01);
        assert!((r.sy + 1.0).abs() < 0.01);
        // sx defaults to 1.0 when not specified
        assert!((r.sx - 1.0).abs() < 0.01);
    }

    /// §19.3.1.37 — a p:pic's spPr is CT_ShapeProperties, so every effectLst
    /// child (§20.1.8.16) applies to images. parse_effect_lst is the shared
    /// reader both p:sp and p:pic use; exercise it with the reflection-bearing
    /// effectLst lifted from sample-11's `図 3` picture.
    #[test]
    fn test_pic_effect_lst_resolves_all_effects() {
        let xml = r#"<spPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <effectLst>
                <outerShdw blurRad="50800" dist="38100" dir="2700000"><srgbClr val="000000"><alpha val="40000"/></srgbClr></outerShdw>
                <innerShdw blurRad="63500" dist="50800" dir="5400000"><srgbClr val="111111"/></innerShdw>
                <glow rad="63500"><srgbClr val="FFCC00"/></glow>
                <softEdge rad="25400"/>
                <reflection blurRad="12700" stA="38000" endPos="28000" dist="5000" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>
            </effectLst>
        </spPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let sp_pr = doc.root_element();
        let eff = parse_effect_lst(child(sp_pr, "effectLst"), &theme);

        let shadow = eff.shadow.expect("outerShdw should resolve");
        assert_eq!(shadow.blur, 50_800);
        assert_eq!(shadow.dist, 38_100);
        assert!((shadow.alpha - 0.4).abs() < 0.01);

        let inner = eff.inner_shadow.expect("innerShdw should resolve");
        assert_eq!(inner.blur, 63_500);

        let glow = eff.glow.expect("glow should resolve");
        assert_eq!(glow.radius, 63_500);
        assert_eq!(glow.color, "FFCC00");

        let soft = eff.soft_edge.expect("softEdge should resolve");
        assert_eq!(soft.radius, 25_400);

        let r = eff.reflection.expect("reflection should resolve");
        assert_eq!(r.blur, 12_700);
        assert_eq!(r.dist, 5_000);
        assert!((r.dir - 90.0).abs() < 0.001);
        assert!((r.st_a - 0.38).abs() < 0.01);
        assert!((r.end_pos - 0.28).abs() < 0.01);
        assert!((r.sy + 1.0).abs() < 0.01);
    }

    /// A spPr with no effectLst yields an all-None EffectLst (the common case).
    #[test]
    fn test_pic_effect_lst_empty_when_absent() {
        let xml = r#"<spPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let eff = parse_effect_lst(child(doc.root_element(), "effectLst"), &theme);
        assert!(eff.shadow.is_none());
        assert!(eff.inner_shadow.is_none());
        assert!(eff.glow.is_none());
        assert!(eff.soft_edge.is_none());
        assert!(eff.reflection.is_none());
    }

    /// §20.1.9.18 — `<a:prstGeom prst="roundRect">` on a picture's spPr clips
    /// the bitmap to a rounded rect. An explicit `adj` guide is carried through;
    /// the preset default is supplied by the shared engine, not the parser.
    #[test]
    fn test_pic_prst_geom_round_rect_explicit_adj() {
        let xml = r#"<spPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <prstGeom prst="roundRect"><avLst><gd name="adj" fmla="val 8594"/></avLst></prstGeom>
        </spPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        assert_eq!(
            parse_pic_prst_geom(doc.root_element()),
            (Some("roundRect".to_owned()), Some(vec![8_594]))
        );
    }

    /// When avLst omits the guide, the parser carries the name with no adjust;
    /// the preset's own default (roundRect adj = 16667) is filled in downstream
    /// by the TS preset-geometry engine, keeping defaults in one place.
    #[test]
    fn test_pic_prst_geom_round_rect_default_adj() {
        let xml = r#"<spPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <prstGeom prst="roundRect"><avLst/></prstGeom>
        </spPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        assert_eq!(
            parse_pic_prst_geom(doc.root_element()),
            (Some("roundRect".to_owned()), None)
        );
    }

    /// §20.1.9.18 generalised — a non-roundRect preset (ellipse, empty avLst) is
    /// now carried generically so the picture clips to that silhouette.
    #[test]
    fn test_pic_prst_geom_ellipse() {
        let xml = r#"<spPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <prstGeom prst="ellipse"><avLst/></prstGeom>
        </spPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        assert_eq!(
            parse_pic_prst_geom(doc.root_element()),
            (Some("ellipse".to_owned()), None)
        );
    }

    /// Multiple adjust guides are captured in declaration order.
    #[test]
    fn test_pic_prst_geom_multi_adj() {
        let xml = r#"<spPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <prstGeom prst="round2SameRect"><avLst>
                <gd name="adj1" fmla="val 16667"/><gd name="adj2" fmla="val 0"/>
            </avLst></prstGeom>
        </spPr>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        assert_eq!(
            parse_pic_prst_geom(doc.root_element()),
            (Some("round2SameRect".to_owned()), Some(vec![16_667, 0]))
        );
    }

    /// A plain rect (or no prstGeom at all) means no clip path.
    #[test]
    fn test_pic_prst_geom_absent() {
        let rect = r#"<spPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <prstGeom prst="rect"><avLst/></prstGeom>
        </spPr>"#;
        let doc = roxmltree::Document::parse(rect).unwrap();
        assert_eq!(parse_pic_prst_geom(doc.root_element()), (None, None));

        let bare = r#"<spPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#;
        let doc = roxmltree::Document::parse(bare).unwrap();
        assert_eq!(parse_pic_prst_geom(doc.root_element()), (None, None));
    }

    /// ECMA-376 §20.1.8.42 — `<a:ln cmpd="dbl"/>` should round-trip.
    /// `cmpd="sng"` is the spec default and stays absent in the model.
    #[test]
    fn test_parse_stroke_cmpd() {
        let theme = HashMap::new();
        let dbl = r#"<ln xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" w="38100" cmpd="dbl"><solidFill><srgbClr val="000000"/></solidFill></ln>"#;
        let doc = roxmltree::Document::parse(dbl).unwrap();
        let s = parse_stroke(doc.root_element(), &theme).expect("stroke should parse");
        assert_eq!(s.cmpd.as_deref(), Some("dbl"));

        let sng = r#"<ln xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" w="38100" cmpd="sng"><solidFill><srgbClr val="000000"/></solidFill></ln>"#;
        let doc = roxmltree::Document::parse(sng).unwrap();
        let s = parse_stroke(doc.root_element(), &theme).expect("stroke should parse");
        assert!(s.cmpd.is_none());
    }

    #[test]
    fn master_body_style_per_level_font_sizes() {
        // ECMA-376 §21.1.2.4: each list level has its own defRPr sz. A 2nd-level
        // bullet must inherit lvl3pPr's smaller size, not lvl1pPr's.
        let master = r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
          <p:txStyles>
            <p:bodyStyle>
              <a:lvl1pPr><a:defRPr sz="2800"/></a:lvl1pPr>
              <a:lvl2pPr><a:defRPr sz="2400"/></a:lvl2pPr>
              <a:lvl3pPr><a:defRPr sz="2000"/></a:lvl3pPr>
            </p:bodyStyle>
            <p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle>
          </p:txStyles>
        </p:sldMaster>"#;
        let master_doc = roxmltree::Document::parse(master).unwrap();
        let m = parse_master_level_font_sizes(master_doc.root_element());
        let body = m.get("body").expect("body level sizes");
        assert_eq!(body[0], Some(28.0)); // lvl1 → level 0
        assert_eq!(body[1], Some(24.0)); // lvl2 → level 1
        assert_eq!(body[2], Some(20.0)); // lvl3 → level 2
        assert_eq!(body[3], None); // unspecified
                                   // body style also keys the empty placeholder type and "obj".
        assert_eq!(m.get("").unwrap()[2], Some(20.0));
        // title style is captured separately.
        assert_eq!(m.get("title").unwrap()[0], Some(44.0));
    }

    /// ECMA-376 §19.7.10 / §21.1.2.4 — a slide body paragraph with no explicit
    /// `<a:buChar>` inherits the master `bodyStyle` bullet. `parse_master_level_bullets`
    /// must surface that `•` (keyed by body/""/obj), so the renderer can draw it.
    /// Regression: sample-9 slides 4/7/12 bullet lists rendered with no markers.
    #[test]
    fn master_body_style_bullets_inherited_by_level() {
        let master = r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
          <p:txStyles>
            <p:bodyStyle>
              <a:lvl1pPr><a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="2000"/></a:lvl1pPr>
              <a:lvl2pPr><a:buFont typeface="Arial"/><a:buChar char="–"/><a:defRPr sz="1800"/></a:lvl2pPr>
            </p:bodyStyle>
            <p:titleStyle><a:lvl1pPr><a:buNone/><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle>
          </p:txStyles>
        </p:sldMaster>"#;
        let theme = HashMap::new();
        let master_rels = HashMap::new();
        // Char bullets only — no media part lookups, so an empty archive suffices.
        let bytes = empty_zip_bytes();
        let cursor = Cursor::new(bytes.clone());
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        let master_doc = roxmltree::Document::parse(master).unwrap();
        let m = parse_master_level_bullets(
            master_doc.root_element(),
            &theme,
            &master_rels,
            "ppt/slideMasters",
            &mut zip,
        );
        let body = m.get("body").expect("body bullets");
        match &body[0] {
            Some(Bullet::Char { ch, .. }) => assert_eq!(ch, "•", "lvl1 bullet char"),
            other => panic!("expected lvl1 char bullet, got {other:?}"),
        }
        match &body[1] {
            Some(Bullet::Char { ch, .. }) => assert_eq!(ch, "–", "lvl2 bullet char"),
            other => panic!("expected lvl2 char bullet, got {other:?}"),
        }
        assert!(body[2].is_none(), "lvl3 unspecified");
        // body style also keys the empty placeholder type and "obj".
        assert!(matches!(
            m.get("").and_then(|b| b[0].clone()),
            Some(Bullet::Char { .. })
        ));
        assert!(matches!(
            m.get("obj").and_then(|b| b[0].clone()),
            Some(Bullet::Char { .. })
        ));
        // titleStyle's explicit buNone is captured (so titles don't inherit a bullet).
        assert!(matches!(
            m.get("title").and_then(|b| b[0].clone()),
            Some(Bullet::None)
        ));
    }

    #[test]
    fn merge_level_sizes_prefers_primary_per_edge() {
        let primary: LevelFontSizes = {
            let mut a = [None; 9];
            a[1] = Some(28.0);
            a
        };
        let fallback: LevelFontSizes = {
            let mut a = [None; 9];
            a[0] = Some(32.0);
            a[1] = Some(24.0);
            a[2] = Some(20.0);
            a
        };
        let merged = merge_level_sizes(&primary, &fallback);
        assert_eq!(merged[0], Some(32.0)); // only fallback
        assert_eq!(merged[1], Some(28.0)); // primary wins
        assert_eq!(merged[2], Some(20.0)); // only fallback
    }

    /// ECMA-376 §21.1.2.4.13 — `<a:lvlNpPr>` is a `CT_TextParagraphProperties`,
    /// so `marL`/`marR`/`indent` are attributes ON the level element itself.
    /// `parse_master_level_indents` must surface the authored per-level values
    /// (keyed by body/""/obj for bodyStyle) and merge per-axis: a level that
    /// sets only `marL` leaves `marR`/`indent` None so a lower tier supplies them.
    #[test]
    fn master_body_style_per_level_indents() {
        let master = r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
          <p:txStyles>
            <p:bodyStyle>
              <a:lvl1pPr marL="1000000" indent="-500000"><a:defRPr sz="2800"/></a:lvl1pPr>
              <a:lvl2pPr marL="1500000"><a:defRPr sz="2400"/></a:lvl2pPr>
            </p:bodyStyle>
            <p:titleStyle><a:lvl1pPr marR="123456"><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle>
          </p:txStyles>
        </p:sldMaster>"#;
        let master_doc = roxmltree::Document::parse(master).unwrap();
        let m = parse_master_level_indents(master_doc.root_element());
        let body = m.get("body").expect("body level indents");
        assert_eq!(body[0].mar_l, Some(1_000_000));
        assert_eq!(body[0].indent, Some(-500_000));
        assert_eq!(body[0].mar_r, None); // unspecified axis stays None
        assert_eq!(body[1].mar_l, Some(1_500_000));
        assert_eq!(body[1].indent, None); // lvl2 omits indent → None
        assert_eq!(body[2].mar_l, None); // unspecified level
                                         // body style also keys the empty placeholder type and "obj".
        assert_eq!(m.get("").unwrap()[0].mar_l, Some(1_000_000));
        assert_eq!(m.get("obj").unwrap()[1].mar_l, Some(1_500_000));
        // title style is captured separately.
        assert_eq!(m.get("title").unwrap()[0].mar_r, Some(123_456));
    }

    /// Per-axis, per-level merge: `primary[lvl].x` wins, else `fallback[lvl].x`.
    #[test]
    fn merge_level_indents_prefers_primary_per_axis() {
        let primary: LevelIndents = {
            let mut a: LevelIndents = Default::default();
            a[0].mar_l = Some(100);
            a[1].indent = Some(-200);
            a
        };
        let fallback: LevelIndents = {
            let mut a: LevelIndents = Default::default();
            a[0].mar_l = Some(999); // loses to primary
            a[0].mar_r = Some(50); // only fallback
            a[1].indent = Some(-999); // loses to primary
            a[1].mar_l = Some(300); // only fallback
            a
        };
        let merged = merge_level_indents(&primary, &fallback);
        assert_eq!(merged[0].mar_l, Some(100)); // primary wins
        assert_eq!(merged[0].mar_r, Some(50)); // only fallback
        assert_eq!(merged[1].indent, Some(-200)); // primary wins
        assert_eq!(merged[1].mar_l, Some(300)); // only fallback
    }

    /// ECMA-376 §21.1.2.4.13 cascade end-to-end: a paragraph whose body lstStyle
    /// authors per-level `marL`/`indent` and whose own `<a:pPr>` omits them must
    /// resolve to the AUTHORED level values (not the hardcoded implicit
    /// `(lvl+1)*342900` / `-342900`). A direct `<a:pPr marL=...>` still wins.
    /// With nothing authored, the implicit default applies (regression guard).
    #[test]
    fn pptx_level_indent_inherited_from_lststyle() {
        let theme = HashMap::new();
        let rels = HashMap::new();
        let bytes = empty_zip_bytes();
        let cursor = Cursor::new(bytes.clone());
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        // `lst_style` sets the body lstStyle (the inherited per-level cascade);
        // `p_pr` is the paragraph's own pPr. Returns the single paragraph.
        let mut parse_para = |lst_style: &str, p_pr: &str| -> Paragraph {
            let xml = format!(
                r#"<txBody xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">{lst_style}<p>{p_pr}<r><t>x</t></r></p></txBody>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            let mut tb = parse_text_body(
                doc.root_element(),
                &theme,
                &rels,
                None,
                [None; 9],
                Default::default(), // inherited_level_indents
                &empty_level_bullets(),
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                ShapeKind::Sp,
                &mut zip,
            );
            tb.paragraphs.remove(0)
        };

        // (1) Authored level marL/indent inherited when the paragraph omits them.
        let lst = r#"<lstStyle><lvl1pPr marL="1000000" indent="-500000"/></lstStyle>"#;
        let inherited = parse_para(lst, "<pPr/>");
        assert_eq!(
            inherited.mar_l, 1_000_000,
            "marL should inherit the authored lvl1pPr value, not the implicit default"
        );
        assert_eq!(
            inherited.indent, -500_000,
            "indent should inherit the authored lvl1pPr value, not the implicit default"
        );

        // (2) A direct pPr marL overrides the inherited level value.
        let overridden = parse_para(lst, r#"<pPr marL="2000000"/>"#);
        assert_eq!(
            overridden.mar_l, 2_000_000,
            "direct pPr marL must win over the inherited level value"
        );
        // indent (not set directly) still inherits the level value.
        assert_eq!(
            overridden.indent, -500_000,
            "indent should still inherit when only marL is set directly"
        );

        // (3) Regression: nothing authored → hardcoded implicit default for a
        // plain (non-bullet) paragraph at lvl 0: marL=0, marR=0, indent=0.
        let implicit = parse_para("", "<pPr/>");
        assert_eq!(implicit.mar_l, 0, "implicit marL default for plain lvl0");
        assert_eq!(implicit.mar_r, 0, "implicit marR default");
        assert_eq!(implicit.indent, 0, "implicit indent default for plain lvl0");
    }

    /// ECMA-376 §21.1.2.4.13 cross-tier, per-axis inheritance: when a layout
    /// placeholder's own `lstStyle` and the master `txStyles` each author a
    /// DIFFERENT axis of the same level, `parse_layout_placeholders` must merge them
    /// per axis (layout wins per axis, master fills the rest) and expose the result
    /// through `lookup_level_indents`. This exercises the actual layout↔master
    /// wiring, not just `merge_level_indents` in isolation.
    #[test]
    fn layout_over_master_level_indents_merge_per_axis() {
        let bytes = empty_zip_bytes();
        let cursor = Cursor::new(bytes.clone());
        let mut zip = zip::ZipArchive::new(cursor).unwrap();

        // Master authors only marL on the body level; layout authors only indent.
        let mut master_indents: HashMap<String, LevelIndents> = HashMap::new();
        let mut body: LevelIndents = Default::default();
        body[0].mar_l = Some(1_000_000);
        master_indents.insert("body".to_string(), body);

        let layout = r#"<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree>
            <p:sp>
              <p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
              <p:spPr/>
              <p:txBody><a:lstStyle><a:lvl1pPr indent="-111111"/></a:lstStyle><a:p/></p:txBody>
            </p:sp>
          </p:spTree></p:cSld>
        </p:sldLayout>"#;

        let layout_doc = roxmltree::Document::parse(layout).unwrap();
        let lph = parse_layout_placeholders(
            layout_doc.root_element(),
            &HashMap::new(),
            &HashMap::new(),
            &master_indents,
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            "",
            &HashMap::new(),
            &mut zip,
        );

        let li = lph.lookup_level_indents("body", None);
        assert_eq!(
            li[0].indent,
            Some(-111_111),
            "indent must come from the LAYOUT lstStyle (primary tier)"
        );
        assert_eq!(
            li[0].mar_l,
            Some(1_000_000),
            "marL must fall back to the MASTER per axis (layout left it unset)"
        );
    }

    /// D4 guard: `parse_layout` resolves the layout placeholder's color-bearing
    /// fields, its `<p:bg>`, and its `showMasterSp` against the `theme` argument.
    /// The color/background must FLIP when the caller passes an override-adjusted
    /// theme — this is what the `parse_presentation` clrMapOvr branch relies on
    /// (a cached `ParsedLayout` is only sound because a no-override slide passes
    /// the same master-baked theme every time). Also asserts a theme-independent
    /// field (transform) is stable regardless of theme.
    #[test]
    fn parse_layout_resolves_color_and_bg_against_theme() {
        let layout = r#"<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          showMasterSp="0">
          <p:cSld>
            <p:bg><p:bgPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></p:bgPr></p:bg>
            <p:spTree>
              <p:sp>
                <p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
                <p:spPr><a:xfrm><a:off x="123456" y="0"/><a:ext cx="10" cy="10"/></a:xfrm></p:spPr>
                <p:txBody><a:lstStyle><a:lvl1pPr><a:defRPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p/></p:txBody>
              </p:sp>
            </p:spTree>
          </p:cSld>
        </p:sldLayout>"#;

        // Typed-empty master inheritance maps (no master fallbacks in this test).
        let m_f64: HashMap<String, f64> = HashMap::new();
        let m_lfs: HashMap<String, LevelFontSizes> = HashMap::new();
        let m_li: HashMap<String, LevelIndents> = HashMap::new();
        let m_lb: HashMap<String, LevelBullets> = HashMap::new();
        let m_str: HashMap<String, String> = HashMap::new();
        let m_tf: HashMap<String, Transform> = HashMap::new();
        let m_bool: HashMap<String, bool> = HashMap::new();
        let m_i64: HashMap<String, i64> = HashMap::new();
        let empty_rels: HashMap<String, String> = HashMap::new();
        let build = |accent1_hex: &str| -> ParsedLayout {
            let mut theme: HashMap<String, String> = HashMap::new();
            theme.insert("accent1".to_string(), accent1_hex.to_string());
            let bytes = empty_zip_bytes();
            let cursor = Cursor::new(bytes.clone());
            let mut zip = zip::ZipArchive::new(cursor).unwrap();
            parse_layout(
                layout,
                &m_f64,
                &m_lfs,
                &m_li,
                &m_lb,
                &m_str,
                &m_tf,
                &m_str,
                &m_bool,
                &m_i64,
                &m_i64,
                &m_f64,
                &theme,
                "ppt/slideLayouts",
                &empty_rels,
                &mut zip,
            )
        };

        let bg_solid_hex = |pl: &ParsedLayout| -> Option<String> {
            match pl.background.as_ref()? {
                Fill::Solid { color } => Some(color.clone()),
                _ => None,
            }
        };

        let base = build("FF0000");
        assert!(!base.show_master_sp, "layout showMasterSp=0 is read");
        assert_eq!(
            bg_solid_hex(&base).as_deref(),
            Some("FF0000"),
            "layout bg schemeClr resolves against the passed theme"
        );
        assert_eq!(
            base.placeholders
                .by_type_color
                .get("body")
                .map(String::as_str),
            Some("FF0000"),
            "layout placeholder defRPr color resolves against the passed theme"
        );
        // Theme-independent geometry is stable.
        assert_eq!(
            base.placeholders.by_type.get("body").map(|t| t.x),
            Some(123456),
            "placeholder transform is theme-independent"
        );

        // Same layout XML, DIFFERENT theme (simulating an override remap): the
        // color-bearing fields must flip; geometry must not.
        let flipped = build("00FF00");
        assert_eq!(
            bg_solid_hex(&flipped).as_deref(),
            Some("00FF00"),
            "override theme must flip the layout bg color"
        );
        assert_eq!(
            flipped
                .placeholders
                .by_type_color
                .get("body")
                .map(String::as_str),
            Some("00FF00"),
            "override theme must flip the layout placeholder color"
        );
        assert_eq!(
            flipped.placeholders.by_type.get("body").map(|t| t.x),
            Some(123456)
        );
    }

    /// Build a deck with `n_slides` slides that ALL reference the same single
    /// layout + single master (no clrMapOvr, no master/layout decorative shapes).
    /// Used to assert the D4 slide-master/layout parse count stays bounded.
    fn build_shared_layout_deck(n_slides: usize) -> Vec<u8> {
        let sld_ids: String = (0..n_slides)
            .map(|i| format!("<p:sldId id=\"{}\" r:id=\"rIdSlide{}\"/>", 256 + i, i))
            .collect();
        let presentation_xml = format!(
            r#"<p:presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster"/></p:sldMasterIdLst>
  <p:sldIdLst>{sld_ids}</p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#
        );
        let mut pres_rel_entries = String::from(
            r#"<Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>"#,
        );
        for i in 0..n_slides {
            pres_rel_entries.push_str(&format!(
                "\n  <Relationship Id=\"rIdSlide{i}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide{i}.xml\"/>"
            ));
        }
        let pres_rels =
            format!("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{pres_rel_entries}</Relationships>");
        let theme_xml = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements><a:clrScheme name="C"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="FF0000"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2><a:accent3><a:srgbClr val="0000FF"/></a:accent3><a:accent4><a:srgbClr val="FFFF00"/></a:accent4><a:accent5><a:srgbClr val="FF00FF"/></a:accent5><a:accent6><a:srgbClr val="00FFFF"/></a:accent6><a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink></a:clrScheme><a:fontScheme name="F"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme><a:fmtScheme name="S"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>"#;
        // Master + layout carry ONLY placeholder shapes (no decorative), so the
        // master-decorative pre-extraction stores an empty vec and the layout
        // decorative walk finds nothing — the parse count reflects the pagination
        // path alone.
        let master_xml = r#"<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rIdLayout"/></p:sldLayoutIdLst></p:sldMaster>"#;
        let master_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>"#;
        let layout_xml = r#"<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>"#;
        let layout_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>"#;
        let slide_xml = r#"<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>"#;
        let slide_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>"#;

        let mut parts: Vec<(String, Vec<u8>)> = vec![
            ("ppt/presentation.xml".into(), presentation_xml.into_bytes()),
            (
                "ppt/_rels/presentation.xml.rels".into(),
                pres_rels.into_bytes(),
            ),
            ("ppt/theme/theme1.xml".into(), theme_xml.as_bytes().to_vec()),
            (
                "ppt/slideMasters/slideMaster1.xml".into(),
                master_xml.as_bytes().to_vec(),
            ),
            (
                "ppt/slideMasters/_rels/slideMaster1.xml.rels".into(),
                master_rels.as_bytes().to_vec(),
            ),
            (
                "ppt/slideLayouts/slideLayout1.xml".into(),
                layout_xml.as_bytes().to_vec(),
            ),
            (
                "ppt/slideLayouts/_rels/slideLayout1.xml.rels".into(),
                layout_rels.as_bytes().to_vec(),
            ),
        ];
        for i in 0..n_slides {
            parts.push((
                format!("ppt/slides/slide{i}.xml"),
                slide_xml.as_bytes().to_vec(),
            ));
            parts.push((
                format!("ppt/slides/_rels/slide{i}.xml.rels"),
                slide_rels.as_bytes().to_vec(),
            ));
        }
        let borrowed: Vec<(&str, &[u8])> = parts
            .iter()
            .map(|(p, b)| (p.as_str(), b.as_slice()))
            .collect();
        zip_with_parts(&borrowed)
    }

    /// D4 regression guard: the slide-master + layout `Document::parse` count on
    /// the pagination path must be BOUNDED — not `k · slides`. With every slide
    /// sharing one layout + one master (no clrMapOvr, no decorations), the master
    /// is built once (1 parse) and the layout is parsed once for the cache, so
    /// the total is `2 + 2·slides` (per slide: its own XML + the layout decorative
    /// walk). Crucially the master/layout parse count does NOT grow by the 12+
    /// (master) or 4 (layout) per-slide factor this change removed. Asserting the
    /// slope across two slide counts pins the optimization: master build and the
    /// layout cache each fire exactly once regardless of N.
    #[test]
    fn parse_count_scales_with_distinct_parts() {
        let count_for = |n: usize| -> usize {
            let data = build_shared_layout_deck(n);
            LAYOUT_MASTER_PARSE_COUNT.with(|c| c.set(0));
            let pres = parse_presentation_from_bytes(&data).expect("parse");
            assert_eq!(pres.slides.len(), n);
            LAYOUT_MASTER_PARSE_COUNT.with(|c| c.get())
        };
        let c3 = count_for(3);
        let c7 = count_for(7);
        // Exact model: 1 (master build) + 1 (layout cache build) + 2·N
        // (per-slide: slide XML + layout decorative walk).
        assert_eq!(c3, 2 + 2 * 3, "3-slide deck D4 parse count");
        assert_eq!(c7, 2 + 2 * 7, "7-slide deck D4 parse count");
        // Slope check: exactly 2 extra parses per added slide (NOT 12+ or 4·k),
        // i.e. the master build + layout parse are amortized to O(1), not O(N).
        assert_eq!(
            (c7 - c3) / (7 - 3),
            2,
            "per-slide D4 parse slope must be 2 (slide + layout-decorative), \
             proving master/layout parses are cached, not per-slide"
        );
    }

    /// PowerPoint stores equations as `a14:m` inside `mc:AlternateContent`
    /// (ECMA-376 §22.1 OMML + 2010 drawing ext). The Choice branch holds the
    /// live `m:oMathPara`; the Fallback (a rasterized picture/text) must be
    /// ignored so the equation isn't double-rendered.
    #[test]
    fn extracts_math_from_alternatecontent_a14m() {
        let xml = r#"<p
            xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
          <mc:AlternateContent>
            <mc:Choice Requires="a14">
              <a14:m>
                <m:oMathPara><m:oMath>
                  <m:r><m:t>x</m:t></m:r>
                </m:oMath></m:oMathPara>
              </a14:m>
            </mc:Choice>
            <mc:Fallback><r><t>fallback</t></r></mc:Fallback>
          </mc:AlternateContent>
        </p>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let ac = doc
            .root_element()
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "AlternateContent")
            .unwrap();
        let theme = HashMap::new();
        let mut runs = Vec::new();
        push_math_runs(ac, Some(18.0), &theme, &mut runs);
        assert_eq!(runs.len(), 1, "exactly one math run, fallback ignored");
        match &runs[0] {
            TextRun::Math {
                display,
                nodes,
                font_size,
                ..
            } => {
                assert!(*display, "oMathPara → display math");
                assert_eq!(*font_size, Some(18.0));
                assert_eq!(nodes_to_text(nodes), "x");
            }
            other => panic!("expected math run, got {other:?}"),
        }
    }

    /// PowerPoint also stores INLINE math as a bare `a14:m` (local name "m")
    /// directly under `a:p` — not wrapped in AlternateContent — holding an
    /// `m:oMath` (not oMathPara). It must extract as inline (display:false) and
    /// pick up its run size from the math run's rPr `sz` (hundredths of a pt).
    #[test]
    fn extracts_inline_bare_a14m_with_run_size() {
        let xml = r#"<m
            xmlns="http://schemas.microsoft.com/office/drawing/2010/main"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
          <m:oMath><m:r>
            <a:rPr sz="2800" i="1"><a:solidFill><a:srgbClr val="7030A0"/></a:solidFill></a:rPr>
            <m:t>n</m:t>
          </m:r></m:oMath>
        </m>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::new();
        let mut runs = Vec::new();
        push_math_runs(doc.root_element(), None, &theme, &mut runs);
        assert_eq!(runs.len(), 1);
        match &runs[0] {
            TextRun::Math {
                display,
                font_size,
                nodes,
                color,
            } => {
                assert!(!*display, "bare a14:m with m:oMath → inline");
                assert_eq!(*font_size, Some(28.0), "size read from math run rPr sz");
                assert_eq!(
                    color.as_deref(),
                    Some("7030A0"),
                    "colour read from math run rPr solidFill"
                );
                assert_eq!(nodes_to_text(nodes), "n");
            }
            other => panic!("expected math run, got {other:?}"),
        }
    }

    /// ECMA-376 §21.1.2.4 / §19.3.1 — a slide body placeholder bound by `idx`
    /// whose layout shape sets size-but-not-colour must still inherit the master
    /// `txStyles` bodyStyle colour (keyed by placeholder *type*). The idx-strict
    /// rule only blocks a sibling *layout* placeholder from leaking its colour; it
    /// must NOT block the master's type-keyed document default.
    ///
    /// Regression: sample-9 slide 2+ body text rendered black instead of the
    /// master bodyStyle's `schemeClr val="bg1"` (→ lt1 → white on a dark theme),
    /// because `lookup_color` returned early on a missing `by_idx_color` entry.
    #[test]
    fn idx_placeholder_inherits_master_txstyle_color() {
        let mut lph = LayoutPlaceholders::default();
        // Master bodyStyle resolves to white and is keyed by type (incl. "" and "body").
        lph.by_type_master_color
            .insert("body".to_string(), "FFFFFF".to_string());
        lph.by_type_master_color
            .insert("".to_string(), "FFFFFF".to_string());

        // Layout idx=35 placeholder declared size only → no by_idx_color entry.
        assert_eq!(
            lph.lookup_color("body", Some(35)),
            Some("FFFFFF".to_string()),
            "idx-bound body placeholder must fall through to the master bodyStyle colour"
        );

        // The layout idx colour still wins when present (idx-strict for the layout tier).
        lph.by_idx_color.insert(35, "112233".to_string());
        assert_eq!(
            lph.lookup_color("body", Some(35)),
            Some("112233".to_string()),
            "an explicit layout idx colour takes priority over the master default"
        );
    }

    /// ECMA-376 §20.1.4.2.27 (`CT_TableStyleCellStyle`) — a cell style's fill is
    /// wrapped in `<a:fill>` and its text colour lives in `<a:tcTxStyle>`. Both the
    /// `firstRow` (header) and `wholeTbl` roles must resolve. Regression: sample-9
    /// slides 9–10 — the orange header fill / pink banding never rendered (fill was
    /// parsed off `<a:tcStyle>` directly, missing the `<a:fill>` wrapper) and the
    /// white header text was ignored (tcTxStyle was never read).
    #[test]
    fn table_style_resolves_fill_wrapper_and_tctxstyle_colour() {
        let theme: HashMap<String, String> =
            [("dk1", "000000"), ("lt1", "FFFFFF"), ("accent2", "B83903")]
                .iter()
                .map(|(k, v)| (k.to_string(), v.to_string()))
                .collect();

        let xml = r#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:tblStyle styleId="{TEST}" styleName="Medium Style 1 - Accent 2">
            <a:wholeTbl>
              <a:tcTxStyle>
                <a:fontRef idx="minor"><a:scrgbClr r="0" g="0" b="0"/></a:fontRef>
                <a:schemeClr val="dk1"/>
              </a:tcTxStyle>
              <a:tcStyle>
                <a:tcBdr>
                  <a:insideH><a:ln w="12700"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:ln></a:insideH>
                </a:tcBdr>
                <a:fill><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:fill>
              </a:tcStyle>
            </a:wholeTbl>
            <a:band1H>
              <a:tcStyle>
                <a:tcBdr/>
                <a:fill><a:solidFill><a:schemeClr val="accent2"><a:tint val="20000"/></a:schemeClr></a:solidFill></a:fill>
              </a:tcStyle>
            </a:band1H>
            <a:firstRow>
              <a:tcTxStyle b="on">
                <a:fontRef idx="minor"><a:scrgbClr r="0" g="0" b="0"/></a:fontRef>
                <a:schemeClr val="lt1"/>
              </a:tcTxStyle>
              <a:tcStyle>
                <a:tcBdr/>
                <a:fill><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:fill>
              </a:tcStyle>
            </a:firstRow>
          </a:tblStyle>
        </a:tblStyleLst>"#;

        let map = parse_table_styles_xml(xml, &theme);
        let def = map.get("{TEST}").expect("style parsed");

        // Fills (wrapped in <a:fill>) must resolve.
        let solid = |f: &Option<Fill>| match f {
            Some(Fill::Solid { color }) => Some(color.clone()),
            _ => None,
        };
        assert_eq!(
            solid(&def.whole_fill).as_deref(),
            Some("FFFFFF"),
            "wholeTbl fill should be lt1 white"
        );
        assert_eq!(
            solid(&def.first_row_fill).as_deref(),
            Some("B83903"),
            "firstRow header fill should be accent2 orange"
        );
        // band1H = accent2 + `<a:tint val="20000">`. Table styles use the literal
        // ECMA-376 tint (val·input + (1-val)·white), giving a near-white wash —
        // NOT the saturated linear-lerp. 0.2·B83903 + 0.8·white = F1D7CD.
        assert_eq!(
            solid(&def.band1h_fill).as_deref(),
            Some("F1D7CD"),
            "band1H tint should be the literal near-white wash, not a saturated lerp"
        );

        // Text colours from tcTxStyle.
        assert_eq!(
            def.whole_text_color.as_deref(),
            Some("000000"),
            "wholeTbl text colour should be dk1 black"
        );
        assert_eq!(
            def.first_row_text_color.as_deref(),
            Some("FFFFFF"),
            "firstRow header text colour should be lt1 white"
        );

        // firstRow `<a:tcTxStyle b="on">` → bold header.
        assert_eq!(
            def.first_row_bold,
            Some(true),
            "firstRow header should be bold from tcTxStyle b=on"
        );
    }

    /// ECMA-376 §21.1.2.1.1 — `<a:bodyPr rtlCol="1">` lays out a multi-column
    /// text body's columns right-to-left. parse_text_body should surface it as
    /// rtl_col=true; an absent attribute yields false (and is omitted from JSON
    /// via skip_serializing_if).
    #[test]
    fn test_parse_text_body_rtl_col() {
        let theme = HashMap::new();
        let rels = HashMap::new();
        // parse_text_body now takes a &mut PptxZip (to verify buBlip parts). This
        // body declares no picture bullets, so an empty archive is sufficient.
        let bytes = empty_zip_bytes();
        let cursor = Cursor::new(bytes.clone());
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        let mut parse = |body_pr: &str| -> TextBody {
            let xml = format!(
                r#"<txBody xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">{body_pr}<p><r><t>x</t></r></p></txBody>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            parse_text_body(
                doc.root_element(),
                &theme,
                &rels,
                None,               // inherited_font_size
                [None; 9],          // inherited_level_font_sizes
                Default::default(), // inherited_level_indents
                &empty_level_bullets(),
                None, // inherited_bold
                None, // inherited_italic
                None, // inherited_caps
                None, // inherited_anchor
                None, // inherited_alignment
                None, // inherited_ea_ln_brk
                None, // inherited_space_before
                None, // inherited_space_after
                None, // inherited_line_spacing
                ShapeKind::Sp,
                &mut zip,
            )
        };

        // rtlCol="1" → true.
        let tb = parse(r#"<bodyPr numCol="2" rtlCol="1"/>"#);
        assert!(tb.rtl_col, "rtlCol=\"1\" should yield rtl_col=true");

        // rtlCol="true" is also accepted (xsd:boolean lexical form).
        let tb_true = parse(r#"<bodyPr numCol="2" rtlCol="true"/>"#);
        assert!(tb_true.rtl_col, "rtlCol=\"true\" should yield rtl_col=true");

        // Absent attribute → false (spec default).
        let tb_absent = parse(r#"<bodyPr numCol="2"/>"#);
        assert!(
            !tb_absent.rtl_col,
            "absent rtlCol should yield rtl_col=false"
        );

        // false is omitted from the serialized JSON.
        let json = serde_json::to_string(&tb_absent).unwrap();
        assert!(
            !json.contains("rtlCol"),
            "rtl_col=false must be omitted from JSON; got {json}"
        );

        // rtlCol="1" appears under the camelCase key "rtlCol".
        let json_true = serde_json::to_string(&tb).unwrap();
        assert!(
            json_true.contains("\"rtlCol\":true"),
            "expected rtlCol:true in JSON; got {json_true}"
        );
    }

    /// ECMA-376 §21.1.2.2.7 — `<a:pPr eaLnBrk>` (xsd:boolean, default true)
    /// controls whether East Asian words may break at a line wrap. The parser
    /// must surface the paragraph's own value, fall back to the body lstStyle
    /// lvl1pPr default, and default to true when nothing specifies it. Mirrors
    /// the `alignment` inheritance shape.
    #[test]
    fn test_parse_paragraph_ea_ln_brk() {
        let theme = HashMap::new();
        let rels = HashMap::new();
        // parse_text_body now takes a &mut PptxZip (to verify buBlip parts). No
        // picture bullets here, so an empty archive is sufficient.
        let bytes = empty_zip_bytes();
        let cursor = Cursor::new(bytes.clone());
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        // `lst_style` lets a test set the body lvl1pPr default; `p_pr` is the
        // paragraph's own pPr. Returns the single paragraph's ea_ln_brk.
        let mut parse_para = |lst_style: &str, p_pr: &str| -> Paragraph {
            let xml = format!(
                r#"<txBody xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">{lst_style}<p>{p_pr}<r><t>東</t></r></p></txBody>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            let mut tb = parse_text_body(
                doc.root_element(),
                &theme,
                &rels,
                None,
                [None; 9],
                Default::default(), // inherited_level_indents
                &empty_level_bullets(),
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                ShapeKind::Sp,
                &mut zip,
            );
            tb.paragraphs.remove(0)
        };

        // eaLnBrk="0" on the paragraph → false.
        assert!(
            !parse_para("", r#"<pPr eaLnBrk="0"/>"#).ea_ln_brk,
            "eaLnBrk=\"0\" should yield ea_ln_brk=false"
        );
        // eaLnBrk="false" (xsd:boolean lexical form) → false.
        assert!(
            !parse_para("", r#"<pPr eaLnBrk="false"/>"#).ea_ln_brk,
            "eaLnBrk=\"false\" should yield ea_ln_brk=false"
        );
        // Omitted everywhere → true (spec default).
        assert!(
            parse_para("", "").ea_ln_brk,
            "omitted eaLnBrk should default to ea_ln_brk=true"
        );
        // eaLnBrk="1" on the paragraph → true.
        assert!(
            parse_para("", r#"<pPr eaLnBrk="1"/>"#).ea_ln_brk,
            "eaLnBrk=\"1\" should yield ea_ln_brk=true"
        );

        // Inheritance: body lstStyle lvl1pPr eaLnBrk="0" propagates to a
        // paragraph that declares no eaLnBrk of its own.
        let inherited = parse_para(r#"<lstStyle><lvl1pPr eaLnBrk="0"/></lstStyle>"#, "");
        assert!(
            !inherited.ea_ln_brk,
            "paragraph should inherit eaLnBrk=false from body lvl1pPr"
        );
        // The paragraph's own value still wins over the inherited body default.
        let overridden = parse_para(
            r#"<lstStyle><lvl1pPr eaLnBrk="0"/></lstStyle>"#,
            r#"<pPr eaLnBrk="1"/>"#,
        );
        assert!(
            overridden.ea_ln_brk,
            "paragraph's own eaLnBrk=\"1\" should override inherited false"
        );

        // ea_ln_brk is serialized under the camelCase key "eaLnBrk".
        let json = serde_json::to_string(&parse_para("", r#"<pPr eaLnBrk="0"/>"#)).unwrap();
        assert!(
            json.contains("\"eaLnBrk\":false"),
            "expected eaLnBrk:false in JSON; got {json}"
        );
    }

    /// ECMA-376 §21.1.3.13 (`a:tblPr@rtl`): a right-to-left table sets `rtl=true`
    /// so the renderer can place column 0 at the right edge. Absent/false must be
    /// omitted from the serialized JSON (TableElement.rtl is optional in TS).
    #[test]
    fn table_rtl_attribute_parses() {
        // An empty in-memory zip is enough: parse_table only reads
        // ppt/tableStyles.xml (absent → no style cascade) and the tbl node.
        fn parse_tbl(tbl_xml: &str) -> TableElement {
            let xml = format!(
                r#"<root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">{tbl_xml}</root>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            let tbl = doc
                .root_element()
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "tbl")
                .unwrap();
            let t = Transform {
                x: 0,
                y: 0,
                cx: 100,
                cy: 100,
                rot: 0.0,
                flip_h: false,
                flip_v: false,
            };
            let theme: HashMap<String, String> = HashMap::new();
            let rels: HashMap<String, String> = HashMap::new();
            let bytes = empty_zip_bytes();
            let cursor = Cursor::new(bytes.clone());
            let mut zip = zip::ZipArchive::new(cursor).unwrap();
            parse_table(tbl, &t, &theme, &rels, &mut zip).unwrap()
        }

        // rtl="1" → rtl=true, serialized.
        let t_rtl = parse_tbl(
            r#"<a:tbl><a:tblPr rtl="1"/><a:tblGrid><a:gridCol w="100"/></a:tblGrid>
               <a:tr h="0"><a:tc><a:txBody/></a:tc></a:tr></a:tbl>"#,
        );
        assert!(t_rtl.rtl, "rtl=\"1\" should yield rtl=true");
        let json = serde_json::to_string(&t_rtl).unwrap();
        assert!(
            json.contains("\"rtl\":true"),
            "expected rtl:true in JSON; got {json}"
        );

        // Absent tblPr@rtl → false, omitted from JSON.
        let t_ltr = parse_tbl(
            r#"<a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="100"/></a:tblGrid>
               <a:tr h="0"><a:tc><a:txBody/></a:tc></a:tr></a:tbl>"#,
        );
        assert!(!t_ltr.rtl, "absent rtl should yield rtl=false");
        let json_ltr = serde_json::to_string(&t_ltr).unwrap();
        assert!(
            !json_ltr.contains("\"rtl\""),
            "rtl=false must be omitted; got {json_ltr}"
        );
    }

    // ===== scene3d / sp3d parsing (ECMA-376 §20.1.5.5 / §20.1.5.12) =====

    /// Wrap a `<p:spPr>` fragment with the `a:`/`p:` namespaces and return the
    /// spPr node so parse_scene3d / parse_sp3d can run against it.
    fn parse_sppr_frag<'a>(doc: &'a roxmltree::Document<'a>) -> roxmltree::Node<'a, 'a> {
        doc.root_element()
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "spPr")
            .unwrap()
    }

    #[test]
    fn test_parse_scene3d_slide3_fragment() {
        // The exact scene3d/sp3d from sample-11 slide 3, "図 3".
        let xml = r#"<root
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:spPr>
            <a:scene3d>
              <a:camera prst="perspectiveRelaxed">
                <a:rot lat="19800000" lon="1200000" rev="20820000"/>
              </a:camera>
              <a:lightRig rig="threePt" dir="t"/>
            </a:scene3d>
            <a:sp3d contourW="6350" prstMaterial="matte">
              <a:bevelT w="101600" h="101600"/>
            </a:sp3d>
          </p:spPr>
        </root>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let sppr = parse_sppr_frag(&doc);

        let scene = parse_scene3d(sppr).expect("scene3d should parse");
        assert_eq!(scene.camera.prst, "perspectiveRelaxed");
        let rot = scene.camera.rot.expect("rot present");
        // 60000ths of a degree → degrees.
        assert!((rot.lat - 330.0).abs() < 1e-9, "lat = {}", rot.lat);
        assert!((rot.lon - 20.0).abs() < 1e-9, "lon = {}", rot.lon);
        assert!((rot.rev - 347.0).abs() < 1e-9, "rev = {}", rot.rev);
        // No fov/zoom in this file → None.
        assert!(scene.camera.fov.is_none());
        assert!(scene.camera.zoom.is_none());
        let lr = scene.light_rig.as_ref().expect("lightRig present");
        assert_eq!(lr.rig, "threePt");
        assert_eq!(lr.dir, "t");

        let sp3d = parse_sp3d(sppr).expect("sp3d should parse");
        assert_eq!(sp3d.contour_w, 6350);
        assert_eq!(sp3d.prst_material, "matte");
        assert_eq!(sp3d.z, 0); // default
        assert_eq!(sp3d.extrusion_h, 0); // default
        let bt = sp3d.bevel_t.expect("bevelT present");
        assert_eq!(bt.w, 101600);
        assert_eq!(bt.h, 101600);
        assert_eq!(bt.prst, "circle"); // schema default
        assert!(sp3d.bevel_b.is_none());

        // camelCase JSON round-trip surfaces the right keys.
        let json = serde_json::to_string(&scene).unwrap();
        assert!(json.contains("\"prst\":\"perspectiveRelaxed\""), "{json}");
        assert!(json.contains("\"lat\":330.0"), "{json}");
        assert!(json.contains("\"lightRig\""), "{json}");
    }

    #[test]
    fn test_parse_camera_fov_zoom_and_defaults() {
        // fov + zoom present; sp3d with all attributes omitted → schema defaults.
        let xml = r#"<root
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:spPr>
            <a:scene3d>
              <a:camera prst="perspectiveContrastingRightFacing" fov="6900000" zoom="200000"/>
              <a:lightRig rig="threePt" dir="t"/>
            </a:scene3d>
            <a:sp3d/>
          </p:spPr>
        </root>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let sppr = parse_sppr_frag(&doc);

        let scene = parse_scene3d(sppr).unwrap();
        // fov: 6900000 / 60000 = 115 degrees.
        assert!((scene.camera.fov.unwrap() - 115.0).abs() < 1e-9);
        // zoom: 200000 / 100000 = 2.0 (200%).
        assert!((scene.camera.zoom.unwrap() - 2.0).abs() < 1e-9);
        // No <a:rot> → None (renderer uses the preset base orientation).
        assert!(scene.camera.rot.is_none());

        let sp3d = parse_sp3d(sppr).unwrap();
        assert_eq!(sp3d.z, 0);
        assert_eq!(sp3d.extrusion_h, 0);
        assert_eq!(sp3d.contour_w, 0);
        assert_eq!(sp3d.prst_material, "warmMatte"); // schema default
        assert!(sp3d.bevel_t.is_none());
        assert!(sp3d.bevel_b.is_none());
    }

    #[test]
    fn test_parse_scene3d_absent_is_none() {
        let xml = r#"<root
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:spPr><a:prstGeom prst="rect"/></p:spPr>
        </root>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let sppr = parse_sppr_frag(&doc);
        assert!(parse_scene3d(sppr).is_none());
        assert!(parse_sp3d(sppr).is_none());
    }

    // ===== sp3d contour colour (ECMA-376 §20.1.5.12 contourClr) =====

    #[test]
    fn test_parse_sp3d_contour_clr_slide3() {
        // The exact sp3d from sample-11 slide 3: contourW + grey contourClr.
        let xml = r#"<root
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:spPr>
            <a:sp3d contourW="6350" prstMaterial="matte">
              <a:bevelT w="101600" h="101600"/>
              <a:contourClr><a:srgbClr val="969696"/></a:contourClr>
            </a:sp3d>
          </p:spPr>
        </root>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let sppr = parse_sppr_frag(&doc);
        let sp3d = parse_sp3d(sppr).expect("sp3d should parse");
        assert_eq!(sp3d.contour_w, 6350);
        assert_eq!(sp3d.contour_clr.as_deref(), Some("969696"));
        let json = serde_json::to_string(&sp3d).unwrap();
        assert!(json.contains("\"contourClr\":\"969696\""), "{json}");
    }

    #[test]
    fn test_parse_sp3d_contour_clr_absent() {
        let xml = r#"<root
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:spPr><a:sp3d contourW="6350"/></p:spPr>
        </root>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let sppr = parse_sppr_frag(&doc);
        let sp3d = parse_sp3d(sppr).unwrap();
        assert!(sp3d.contour_clr.is_none());
        // Omitted from JSON when absent.
        let json = serde_json::to_string(&sp3d).unwrap();
        assert!(!json.contains("contourClr"), "{json}");
    }

    // ===== picture a:ln stroke (ECMA-376 §20.1.2.2.24, §19.3.1.37) =====

    #[test]
    fn test_parse_pic_stroke_solid_fill() {
        // <p:pic>'s spPr > ln with a solidFill → a visible border.
        let xml = r#"<root
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:spPr>
            <a:ln w="38100"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>
          </p:spPr>
        </root>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let sppr = parse_sppr_frag(&doc);
        let theme: HashMap<String, String> = HashMap::new();
        let stroke = child(sppr, "ln")
            .and_then(|n| parse_stroke(n, &theme))
            .expect("pic stroke should parse");
        assert_eq!(stroke.color, "FFFFFF");
        assert_eq!(stroke.width, 38100);
    }

    #[test]
    fn test_parse_pic_stroke_no_fill_is_none() {
        // sample-11's pic borders are <a:ln><a:noFill/></a:ln> → no border.
        let xml = r#"<root
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:spPr><a:ln><a:noFill/></a:ln></p:spPr>
        </root>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let sppr = parse_sppr_frag(&doc);
        let theme: HashMap<String, String> = HashMap::new();
        let stroke = child(sppr, "ln").and_then(|n| parse_stroke(n, &theme));
        assert!(stroke.is_none());
    }

    // ===== p14:media-only embeds (ECMA-376 §19.3.1.17/18; the p14 extension
    // carries no audio/video tag, so media_kind is decided from the MIME of the
    // referenced part). A `<p:pic>` with no `a:videoFile`/`a:audioFile`, just a
    // `<p14:media r:embed>`, must still parse as a MediaElement — not fall
    // through to a poster-only Picture. =====

    /// `<p:pic>` whose only media marker is `<p14:media r:embed>` pointing at a
    /// `.m4v` (a MIME the table must recognise) parses as a video MediaElement.
    /// rId1 → media/clip.m4v, with a poster blip so the renderer has a thumbnail.
    #[test]
    fn test_parse_media_p14_only_m4v_is_video() {
        let xml = r#"<p:pic
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
            xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:nvPicPr>
            <p:cNvPr id="5" name="Media"/>
            <p:nvPr>
              <p:extLst>
                <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">
                  <p14:media r:embed="rId1"/>
                </p:ext>
              </p:extLst>
            </p:nvPr>
          </p:nvPicPr>
          <p:blipFill>
            <a:blip r:embed="rId2"/>
          </p:blipFill>
          <p:spPr>
            <a:xfrm>
              <a:off x="100" y="200"/>
              <a:ext cx="3000" cy="4000"/>
            </a:xfrm>
          </p:spPr>
        </p:pic>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let pic = doc.root_element();
        let mut rels: HashMap<String, String> = HashMap::new();
        rels.insert("rId1".to_string(), "../media/clip.m4v".to_string());
        rels.insert("rId2".to_string(), "../media/image1.png".to_string());

        let media = parse_media(pic, "ppt/slides", &rels)
            .expect("p14:media-only .m4v should parse as a MediaElement");
        assert_eq!(media.media_kind, "video");
        assert_eq!(media.mime_type, "video/mp4");
        assert_eq!(media.media_path, "ppt/media/clip.m4v");
        assert_eq!(media.poster_path, "ppt/media/image1.png");
    }

    /// Same shape but the embed targets a `.wav` → audio MediaElement.
    #[test]
    fn test_parse_media_p14_only_wav_is_audio() {
        let xml = r#"<p:pic
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
            xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:nvPicPr>
            <p:cNvPr id="6" name="Audio"/>
            <p:nvPr>
              <p:extLst>
                <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">
                  <p14:media r:embed="rId1"/>
                </p:ext>
              </p:extLst>
            </p:nvPr>
          </p:nvPicPr>
          <p:spPr>
            <a:xfrm>
              <a:off x="0" y="0"/>
              <a:ext cx="800" cy="800"/>
            </a:xfrm>
          </p:spPr>
        </p:pic>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let pic = doc.root_element();
        let mut rels: HashMap<String, String> = HashMap::new();
        rels.insert("rId1".to_string(), "../media/sound.wav".to_string());

        let media = parse_media(pic, "ppt/slides", &rels)
            .expect("p14:media-only .wav should parse as a MediaElement");
        assert_eq!(media.media_kind, "audio");
        assert_eq!(media.mime_type, "audio/wav");
    }

    /// A `<p:pic>` whose legacy `<a:videoFile r:link>` is broken — here modeled
    /// as a missing rId (`rIdBroken` is absent from rels, so `rels.get` is None)
    /// — but whose `<p14:media r:embed>` points at the real embedded clip must
    /// still parse as a video: the good embed must not be shadowed by the broken
    /// link. This exercises the embed-before-link ordering, not the empty-Target
    /// guard (a real External link would instead carry a non-empty URL).
    #[test]
    fn test_parse_media_prefers_p14_embed_over_broken_videofile_link() {
        let xml = r#"<p:pic
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
            xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:nvPicPr>
            <p:cNvPr id="7" name="Video"/>
            <p:nvPr>
              <a:videoFile r:link="rIdBroken"/>
              <p:extLst>
                <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">
                  <p14:media r:embed="rIdGood"/>
                </p:ext>
              </p:extLst>
            </p:nvPr>
          </p:nvPicPr>
          <p:blipFill><a:blip r:embed="rIdPoster"/></p:blipFill>
          <p:spPr>
            <a:xfrm>
              <a:off x="0" y="0"/>
              <a:ext cx="1280" cy="720"/>
            </a:xfrm>
          </p:spPr>
        </p:pic>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let pic = doc.root_element();
        let mut rels: HashMap<String, String> = HashMap::new();
        // rIdBroken intentionally absent — the link's rId does not resolve
        // (`rels.get` is None). Only the embedded p14:media resolves.
        rels.insert("rIdGood".to_string(), "../media/clip.mp4".to_string());
        rels.insert("rIdPoster".to_string(), "../media/image1.png".to_string());

        let media = parse_media(pic, "ppt/slides", &rels)
            .expect("a broken videoFile link must not shadow the good p14:media embed");
        assert_eq!(media.media_kind, "video");
        assert_eq!(media.media_path, "ppt/media/clip.mp4");
        assert_eq!(media.mime_type, "video/mp4");
    }

    // ===== Master spTree decorative shapes (ECMA-376 §19.3.1.38 sld /
    // §19.3.1.39 sldLayout, showMasterSp) =====

    /// Build a minimal in-memory .pptx whose slide master spTree carries a
    /// decorative picture (image1.png at a non-centred position) plus a
    /// solid-fill rectangle. `layout_show_master_sp` controls the layout's
    /// `showMasterSp` attribute so the test can exercise the suppression path.
    fn build_master_sp_pptx(layout_show_master_sp: Option<bool>) -> Vec<u8> {
        use zip::write::SimpleFileOptions;

        // 1×1 transparent PNG (smallest valid PNG).
        const PNG_1X1: &[u8] = &[
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00,
            0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78,
            0x9C, 0x62, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
            0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
        ];

        let layout_attr = match layout_show_master_sp {
            Some(true) => r#" showMasterSp="1""#.to_string(),
            Some(false) => r#" showMasterSp="0""#.to_string(),
            None => String::new(),
        };

        let presentation_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#;

        let pres_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"#;

        let theme_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">
  <a:themeElements><a:clrScheme name="C">
    <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
    <a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
    <a:accent1><a:srgbClr val="FF0000"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2>
    <a:accent3><a:srgbClr val="0000FF"/></a:accent3><a:accent4><a:srgbClr val="FFFF00"/></a:accent4>
    <a:accent5><a:srgbClr val="FF00FF"/></a:accent5><a:accent6><a:srgbClr val="00FFFF"/></a:accent6>
    <a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="F"><a:majorFont><a:latin typeface="Arial"/></a:majorFont>
    <a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
  <a:fmtScheme name="S"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>"#;

        // Master spTree: a decorative pic (image1.png at x=600000,y=400000) and a
        // solid-fill rectangle. No placeholder, so both are decorative.
        let master_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:pic>
      <p:nvPicPr><p:cNvPr id="10" name="MasterLogo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rIdImg1"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr><a:xfrm><a:off x="600000" y="400000"/><a:ext cx="800000" cy="800000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
    </p:pic>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="11" name="MasterBand"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="200000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="123456"/></a:solidFill></p:spPr>
    </p:sp>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"
    accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rIdLayout"/></p:sldLayoutIdLst>
</p:sldMaster>"#;

        let master_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

        let layout_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"{layout_attr} type="blank">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
</p:sldLayout>"#
        );

        let layout_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#;

        let slide_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
</p:sld>"#;

        let slide_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;

        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zw = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default();
            let mut put = |path: &str, bytes: &[u8]| {
                zw.start_file(path, opts).unwrap();
                use std::io::Write;
                zw.write_all(bytes).unwrap();
            };
            put("ppt/presentation.xml", presentation_xml.as_bytes());
            put("ppt/_rels/presentation.xml.rels", pres_rels.as_bytes());
            put("ppt/theme/theme1.xml", theme_xml.as_bytes());
            put("ppt/slideMasters/slideMaster1.xml", master_xml.as_bytes());
            put(
                "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                master_rels.as_bytes(),
            );
            put("ppt/slideLayouts/slideLayout1.xml", layout_xml.as_bytes());
            put(
                "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                layout_rels.as_bytes(),
            );
            put("ppt/slides/slide1.xml", slide_xml.as_bytes());
            put("ppt/slides/_rels/slide1.xml.rels", slide_rels.as_bytes());
            put("ppt/media/image1.png", PNG_1X1);
            zw.finish().unwrap();
        }
        buf
    }

    /// §19.3.1.38/§19.3.1.39: a master spTree picture (non-placeholder) is
    /// composited onto the slide. Without the fix the master spTree is dropped
    /// and the slide has no elements.
    #[test]
    fn master_sptree_pic_appears_on_slide() {
        let data = build_master_sp_pptx(None);
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        let slide = &pres.slides[0];

        let pic = slide.elements.iter().find_map(|e| match e {
            SlideElement::Picture(p) => Some(p),
            _ => None,
        });
        let pic = pic.expect("master decorative picture should be rendered on the slide");
        // Non-centred position from the master xfrm is preserved.
        assert_eq!(pic.x, 600000, "master pic x");
        assert_eq!(pic.y, 400000, "master pic y");
        assert!(
            pic.image_path.ends_with("media/image1.png"),
            "master pic should resolve image1.png via master rels; got {}",
            pic.image_path
        );
        assert_eq!(pic.mime_type, "image/png", "master pic mime");

        // The decorative rectangle also shows up.
        let has_band = slide
            .elements
            .iter()
            .any(|e| matches!(e, SlideElement::Shape(_)));
        assert!(has_band, "master decorative shape should be rendered");
    }

    /// §19.3.1.39: a layout with showMasterSp="0" suppresses the master's
    /// decorative shapes for slides using that layout.
    #[test]
    fn master_sptree_hidden_when_layout_show_master_sp_false() {
        let data = build_master_sp_pptx(Some(false));
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        let slide = &pres.slides[0];

        let has_master_pic = slide
            .elements
            .iter()
            .any(|e| matches!(e, SlideElement::Picture(_)));
        assert!(
            !has_master_pic,
            "showMasterSp=\"0\" on the layout must suppress master decorations"
        );
        assert!(
            slide.elements.is_empty(),
            "no master decorations expected; got {} elements",
            slide.elements.len()
        );
    }

    /// showMasterSp="1" (explicit true) on the layout keeps master shapes —
    /// guards against an inverted boolean parse.
    #[test]
    fn master_sptree_shown_when_layout_show_master_sp_true() {
        let data = build_master_sp_pptx(Some(true));
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        let slide = &pres.slides[0];
        assert!(
            slide
                .elements
                .iter()
                .any(|e| matches!(e, SlideElement::Picture(_))),
            "showMasterSp=\"1\" must keep master decorations"
        );
    }

    /// Build a minimal in-memory .pptx whose master carries a decorative
    /// (non-placeholder) rectangle filled with a `schemeClr` (accent1). When
    /// `slide_clr_map_ovr` is set, the slide gets a `<p:clrMapOvr>` remapping
    /// accent1→accent2. Exercises the master-decorative pre-extraction (D4): the
    /// no-override slide must reuse the pre-extracted element (accent1's hex),
    /// while an override slide must RE-RESOLVE the decorative fill against its
    /// override theme (accent2's hex) rather than serving the frozen bundle copy.
    fn build_master_scheme_decoration_pptx(remap_accent1_to_accent2: bool) -> Vec<u8> {
        use zip::write::SimpleFileOptions;

        let presentation_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#;
        let pres_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rIdSlide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"#;
        let theme_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">
  <a:themeElements><a:clrScheme name="C">
    <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
    <a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
    <a:accent1><a:srgbClr val="FF0000"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2>
    <a:accent3><a:srgbClr val="0000FF"/></a:accent3><a:accent4><a:srgbClr val="FFFF00"/></a:accent4>
    <a:accent5><a:srgbClr val="FF00FF"/></a:accent5><a:accent6><a:srgbClr val="00FFFF"/></a:accent6>
    <a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="F"><a:majorFont><a:latin typeface="Arial"/></a:majorFont>
    <a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
  <a:fmtScheme name="S"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>"#;
        // Master decorative rectangle filled with schemeClr accent1 (no placeholder).
        let master_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="11" name="MasterBand"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="200000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:schemeClr val="accent1"/></a:solidFill></p:spPr>
    </p:sp>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"
    accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rIdLayout"/></p:sldLayoutIdLst>
</p:sldMaster>"#;
        let master_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;
        let layout_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
</p:sldLayout>"#;
        let layout_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#;
        // Optional slide-level clrMapOvr that remaps accent1 → accent2.
        let clr_map_ovr = if remap_accent1_to_accent2 {
            r#"<p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
      accent1="accent2" accent2="accent2" accent3="accent3" accent4="accent4"
      accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></p:clrMapOvr>"#
        } else {
            ""
        };
        let slide_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  {clr_map_ovr}
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
</p:sld>"#
        );
        let slide_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;

        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zw = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default();
            let mut put = |path: &str, bytes: &[u8]| {
                zw.start_file(path, opts).unwrap();
                use std::io::Write;
                zw.write_all(bytes).unwrap();
            };
            put("ppt/presentation.xml", presentation_xml.as_bytes());
            put("ppt/_rels/presentation.xml.rels", pres_rels.as_bytes());
            put("ppt/theme/theme1.xml", theme_xml.as_bytes());
            put("ppt/slideMasters/slideMaster1.xml", master_xml.as_bytes());
            put(
                "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                master_rels.as_bytes(),
            );
            put("ppt/slideLayouts/slideLayout1.xml", layout_xml.as_bytes());
            put(
                "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                layout_rels.as_bytes(),
            );
            put("ppt/slides/slide1.xml", slide_xml.as_bytes());
            put("ppt/slides/_rels/slide1.xml.rels", slide_rels.as_bytes());
            zw.finish().unwrap();
        }
        buf
    }

    fn master_band_fill_hex(data: &[u8]) -> String {
        let pres = parse_presentation_from_bytes(data).expect("parse");
        let slide = &pres.slides[0];
        let shape = slide
            .elements
            .iter()
            .find_map(|e| match e {
                SlideElement::Shape(s) => Some(s),
                _ => None,
            })
            .expect("master decorative shape present on slide");
        match shape.fill.as_ref().expect("shape has fill") {
            Fill::Solid { color } => color.clone(),
            other => panic!("expected solid fill, got {other:?}"),
        }
    }

    /// D4 guard: the pre-extracted master decorative shape (no override) resolves
    /// its `schemeClr accent1` against the master's own theme — accent1 = FF0000.
    #[test]
    fn master_decoration_scheme_fill_no_override_uses_master_theme() {
        let hex = master_band_fill_hex(&build_master_scheme_decoration_pptx(false));
        assert_eq!(
            hex.to_uppercase(),
            "FF0000",
            "no-override slide must resolve accent1 to its master-theme hex"
        );
    }

    /// D4 guard: a slide with `<p:clrMapOvr>` remapping accent1→accent2 must
    /// RE-RESOLVE the master decorative shape against its override theme
    /// (accent2 = 00FF00), NOT serve the frozen pre-extracted copy (FF0000).
    /// This is the `eff.is_some()` re-extraction branch in `parse_slide`.
    #[test]
    fn master_decoration_scheme_fill_override_reresolves_against_override_theme() {
        let hex = master_band_fill_hex(&build_master_scheme_decoration_pptx(true));
        assert_eq!(
            hex.to_uppercase(),
            "00FF00",
            "override slide must flip the master decorative accent1→accent2 hex"
        );
    }

    /// Build a minimal in-memory .pptx whose **master** carries a
    /// `<p:bg><p:bgPr><a:blipFill>` image background and whose slide + layout
    /// have NO `<p:bg>`. Exercises the slide → layout → master background
    /// inheritance (ECMA-376 §19.3.1.1 / §20.1.8.14) end-to-end so the slide's
    /// resolved `background` should be the master's image fill.
    fn build_master_bg_blip_pptx() -> Vec<u8> {
        use zip::write::SimpleFileOptions;
        const PNG_1X1: &[u8] = &[
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00,
            0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78,
            0x9C, 0x62, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
            0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
        ];

        let presentation_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#;
        let pres_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>"#;
        let theme_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">
  <a:themeElements><a:clrScheme name="C">
    <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
    <a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
    <a:accent1><a:srgbClr val="FF0000"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2>
    <a:accent3><a:srgbClr val="0000FF"/></a:accent3><a:accent4><a:srgbClr val="FFFF00"/></a:accent4>
    <a:accent5><a:srgbClr val="FF00FF"/></a:accent5><a:accent6><a:srgbClr val="00FFFF"/></a:accent6>
    <a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="F"><a:majorFont><a:latin typeface="Arial"/></a:majorFont>
    <a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
  <a:fmtScheme name="S"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>"#;
        // Master defines a blipFill background; slide + layout do not.
        let master_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdImg1"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"
    accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rIdLayout"/></p:sldLayoutIdLst>
</p:sldMaster>"#;
        let master_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;
        let layout_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
</p:sldLayout>"#;
        let layout_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#;
        let slide_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
</p:sld>"#;
        let slide_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;

        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zw = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default();
            let mut put = |path: &str, bytes: &[u8]| {
                zw.start_file(path, opts).unwrap();
                use std::io::Write;
                zw.write_all(bytes).unwrap();
            };
            put("ppt/presentation.xml", presentation_xml.as_bytes());
            put("ppt/_rels/presentation.xml.rels", pres_rels.as_bytes());
            put("ppt/theme/theme1.xml", theme_xml.as_bytes());
            put("ppt/slideMasters/slideMaster1.xml", master_xml.as_bytes());
            put(
                "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                master_rels.as_bytes(),
            );
            put("ppt/slideLayouts/slideLayout1.xml", layout_xml.as_bytes());
            put(
                "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                layout_rels.as_bytes(),
            );
            put("ppt/slides/slide1.xml", slide_xml.as_bytes());
            put("ppt/slides/_rels/slide1.xml.rels", slide_rels.as_bytes());
            put("ppt/media/image1.png", PNG_1X1);
            zw.finish().unwrap();
        }
        buf
    }

    /// ECMA-376 §19.3.1.1 + §20.1.8.14 — a slide with no `<p:bg>` inherits the
    /// master's `<p:bg><p:bgPr><a:blipFill>` image background through the
    /// slide → layout → master chain. The resolved `slide.background` must be a
    /// `Fill::Image` carrying the master-rels-resolved zip path.
    #[test]
    fn slide_inherits_master_blip_background() {
        let data = build_master_bg_blip_pptx();
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        let bg = pres.slides[0]
            .background
            .as_ref()
            .expect("slide should inherit a background from the master");
        match bg {
            Fill::Image {
                image_path,
                mime_type,
                ..
            } => {
                assert!(
                    image_path.ends_with("media/image1.png"),
                    "master bg should resolve image1.png via master rels; got {image_path}"
                );
                assert_eq!(mime_type, "image/png", "master bg mime");
            }
            other => panic!("expected inherited Fill::Image background, got {other:?}"),
        }
    }

    // ── Embedded SVG images (Microsoft asvg:svgBlip extension) ────────────
    //
    // PowerPoint stores an SVG picture as a `<p:pic>` whose `<a:blip>` points
    // at a PNG *fallback* (r:embed) and carries the real .svg part inside an
    // `<a:extLst><a:ext uri="{96DAC541-…}"><asvg:svgBlip r:embed="…"/>`
    // extension (Microsoft 2016 SVG extension; the core blip fill is
    // ECMA-376 §20.1.8.14). The parser must keep emitting the PNG fallback's
    // zip path as `image_path` (regression-safe) while additionally surfacing
    // the SVG part's path on `svg_image_path` so the renderer can prefer the
    // vector original.

    /// Build a tiny zip containing only the two media parts a `<p:pic>` blip
    /// references (a PNG fallback and an SVG body), so `parse_picture` can be
    /// driven directly with a hand-rolled rels map.
    fn build_blip_media_zip(png: &[u8], svg: &[u8]) -> Vec<u8> {
        use zip::write::SimpleFileOptions;
        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zw = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default();
            zw.start_file("ppt/media/image1.png", opts).unwrap();
            {
                use std::io::Write;
                zw.write_all(png).unwrap();
            }
            zw.start_file("ppt/media/image2.svg", opts).unwrap();
            {
                use std::io::Write;
                zw.write_all(svg).unwrap();
            }
            zw.finish().unwrap();
        }
        buf
    }

    /// A `<p:pic>` carrying a PNG fallback blip plus an `asvg:svgBlip`
    /// extension must yield both the PNG `image_path` and the SVG
    /// `svg_image_path` (with mimes), never inlined base64.
    #[test]
    fn picture_with_svg_blip_extension_emits_both_urls() {
        // 1×1 transparent PNG (smallest valid PNG).
        const PNG_1X1: &[u8] = &[
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00,
            0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78,
            0x9C, 0x62, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
            0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
        ];
        const SVG: &[u8] =
            br##"<svg xmlns="http://www.w3.org/2000/svg" width="2" height="2"><rect width="2" height="2" fill="#0a0"/></svg>"##;

        // The svgBlip uses a different prefix (asvg:) on purpose — matching is by
        // namespace-local name, so the prefix must not matter.
        let pic_xml = r#"<p:pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
  <p:nvPicPr><p:cNvPr id="5" name="SvgPic"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rIdPng">
      <a:extLst>
        <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
          <asvg:svgBlip r:embed="rIdSvg"/>
        </a:ext>
      </a:extLst>
    </a:blip>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300000" cy="300000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
</p:pic>"#;

        let doc = roxmltree::Document::parse(pic_xml).unwrap();
        let pic_node = doc.root_element();

        let mut rels = HashMap::new();
        rels.insert("rIdPng".to_string(), "../media/image1.png".to_string());
        rels.insert("rIdSvg".to_string(), "../media/image2.svg".to_string());

        let theme = HashMap::new();
        let data = build_blip_media_zip(PNG_1X1, SVG);
        let cursor = Cursor::new(data.clone());
        let mut zip = zip::ZipArchive::new(cursor).unwrap();

        let pic = parse_picture(pic_node, "ppt/slides", &rels, &theme, &mut zip)
            .expect("parse_picture should succeed for an SVG-blip picture");

        // PNG fallback is preserved as the raster image_path (regression-safe);
        // never an inlined data URL.
        assert_eq!(pic.image_path, "ppt/media/image1.png", "raster path");
        assert_eq!(pic.mime_type, "image/png", "raster mime");
        assert!(
            !pic.image_path.contains(";base64,"),
            "image_path must not inline base64; got {}",
            pic.image_path
        );

        // The SVG original is surfaced separately as a zip path.
        assert_eq!(
            pic.svg_image_path.as_deref(),
            Some("ppt/media/image2.svg"),
            "svg_image_path must point at the .svg part",
        );
        // And the resolved path must hold the original SVG bytes.
        let svg_bytes = extract_image(&data, "ppt/media/image2.svg", None)
            .expect("svg part must be readable by its resolved path");
        assert_eq!(
            svg_bytes, SVG,
            "bytes at svg_image_path must equal the .svg part"
        );
    }

    /// A plain `<p:pic>` with no svgBlip extension must leave `svg_image_path`
    /// as None (and still emit the PNG `image_path` + intrinsic size) — guards
    /// against the new branch firing spuriously.
    #[test]
    fn picture_without_svg_blip_has_no_svg_url() {
        const PNG_1X1: &[u8] = &[
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00,
            0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78,
            0x9C, 0x62, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
            0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
        ];
        let pic_xml = r#"<p:pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:nvPicPr><p:cNvPr id="5" name="PngPic"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill><a:blip r:embed="rIdPng"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
  <p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300000" cy="300000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
</p:pic>"#;
        let doc = roxmltree::Document::parse(pic_xml).unwrap();
        let pic_node = doc.root_element();
        let mut rels = HashMap::new();
        rels.insert("rIdPng".to_string(), "../media/image1.png".to_string());
        let theme = HashMap::new();
        let data = build_blip_media_zip(PNG_1X1, b"<svg/>");
        let cursor = Cursor::new(data.clone());
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        let pic = parse_picture(pic_node, "ppt/slides", &rels, &theme, &mut zip)
            .expect("parse_picture should succeed");
        assert_eq!(pic.image_path, "ppt/media/image1.png");
        assert_eq!(pic.mime_type, "image/png");
        // 1×1 PNG → intrinsic size read from the IHDR.
        assert_eq!(pic.intrinsic_width_px, Some(1), "intrinsic width");
        assert_eq!(pic.intrinsic_height_px, Some(1), "intrinsic height");
        assert!(
            pic.svg_image_path.is_none(),
            "svg_image_path must be None without an svgBlip extension"
        );
    }

    /// A `<p:pic>` whose `<a:blip>` carries ONLY the `asvg:svgBlip` extension —
    /// no raster `r:embed` fallback at all (an icon inserted as a pure SVG, as
    /// in sample-12) — must still parse. Previously the mandatory raster embed
    /// (`attr_r(&blip, "embed")?`) made `parse_picture` return None, so the whole
    /// picture was silently dropped and the SVG never rendered.
    #[test]
    fn picture_with_only_svg_blip_and_no_raster_embed_still_parses() {
        const SVG: &[u8] = br##"<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24"><path d="M0 0h24v24H0z" fill="#0a0"/></svg>"##;

        // The `<a:blip>` has NO r:embed attribute — the image is referenced only
        // through the svgBlip extension.
        let pic_xml = r#"<p:pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
  <p:nvPicPr><p:cNvPr id="4" name="SvgOnly"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill>
    <a:blip>
      <a:extLst>
        <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
          <asvg:svgBlip r:embed="rIdSvg"/>
        </a:ext>
      </a:extLst>
    </a:blip>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300000" cy="300000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
</p:pic>"#;

        let doc = roxmltree::Document::parse(pic_xml).unwrap();
        let pic_node = doc.root_element();

        let mut rels = HashMap::new();
        rels.insert("rIdSvg".to_string(), "../media/image2.svg".to_string());

        let theme = HashMap::new();
        // Only the .svg part is referenced; the PNG arg is unused here.
        let data = build_blip_media_zip(b"", SVG);
        let cursor = Cursor::new(data.clone());
        let mut zip = zip::ZipArchive::new(cursor).unwrap();

        let pic = parse_picture(pic_node, "ppt/slides", &rels, &theme, &mut zip)
            .expect("parse_picture must succeed for an svgBlip-only picture (sample-12 case)");

        // The SVG original is surfaced on svg_image_path so the renderer prefers it.
        assert_eq!(
            pic.svg_image_path.as_deref(),
            Some("ppt/media/image2.svg"),
            "svg_image_path must point at the .svg part",
        );

        // With no raster blip, image_path falls back to the SVG part itself so
        // the element is always drawable (rather than being dropped or empty);
        // its mime is image/svg+xml and no PNG intrinsic size is recorded.
        assert_eq!(
            pic.image_path, "ppt/media/image2.svg",
            "image_path must fall back to the SVG when no raster blip is embedded",
        );
        assert_eq!(pic.mime_type, "image/svg+xml");
        assert_eq!(pic.intrinsic_width_px, None, "no PNG intrinsic for SVG");
        assert_eq!(pic.intrinsic_height_px, None);
        // The resolved SVG path must hold the original SVG bytes.
        let svg_bytes = extract_image(&data, "ppt/media/image2.svg", None)
            .expect("svg part must be readable by its resolved path");
        assert_eq!(
            svg_bytes, SVG,
            "bytes at svg_image_path must equal the .svg part"
        );
    }

    // ── Per-slide theme/master resolution (slide→layout→master→theme) ─────
    //
    // A deck with TWO masters, each carrying a DIFFERENT theme (different
    // accent1). Two layouts (layoutA→masterA, layoutB→masterB) and two slides
    // (slide1→layoutA, slide2→layoutB). Each slide has a shape whose fill comes
    // from `<p:style><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>`
    // with no explicit spPr fill. Before the fix the parser loaded the
    // presentation's first theme/master once and applied it to every slide, so
    // both shapes resolved to masterA's accent1. After the fix each slide must
    // resolve accent1 from its own master's theme.
    //
    // `clr_map_a` lets the test optionally give masterA a non-default
    // `<p:clrMap>` (e.g. bg1/tx1 swapped) so the clrMap-honoring assertion can
    // reuse the same builder.
    fn build_two_master_pptx(clr_map_a: &str) -> Vec<u8> {
        use zip::write::SimpleFileOptions;

        let presentation_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rIdMasterA"/>
    <p:sldMasterId id="2147483649" r:id="rIdMasterB"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rIdSlide1"/>
    <p:sldId id="257" r:id="rIdSlide2"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#;

        // presentation rels intentionally lists masterA FIRST so the legacy
        // "first master / first theme" path would pick masterA's accent1.
        let pres_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMasterA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rIdMasterB" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster2.xml"/>
  <Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rIdSlide2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
  <Relationship Id="rIdThemeA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"#;

        // Two themes that differ only in accent1 (and tx1/bg1 hex so the clrMap
        // swap is observable).
        let theme_a = |accent1: &str| {
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">
  <a:themeElements><a:clrScheme name="C">
    <a:dk1><a:srgbClr val="222222"/></a:dk1><a:lt1><a:srgbClr val="FAFAFA"/></a:lt1>
    <a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
    <a:accent1><a:srgbClr val="{accent1}"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2>
    <a:accent3><a:srgbClr val="0000FF"/></a:accent3><a:accent4><a:srgbClr val="FFFF00"/></a:accent4>
    <a:accent5><a:srgbClr val="FF00FF"/></a:accent5><a:accent6><a:srgbClr val="00FFFF"/></a:accent6>
    <a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="F"><a:majorFont><a:latin typeface="Arial"/></a:majorFont>
    <a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
  <a:fmtScheme name="S"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>"#
            )
        };
        let theme1_xml = theme_a("72A376"); // masterA accent1
        let theme2_xml = theme_a("4F81BD"); // masterB accent1

        let master = |clr_map: &str| {
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
  {clr_map}
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483650" r:id="rIdLayout"/></p:sldLayoutIdLst>
</p:sldMaster>"#
            )
        };
        let default_clr_map = r#"<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#;
        let master1_xml = master(clr_map_a);
        let master2_xml = master(default_clr_map);

        // Each master's rels points at its OWN theme and its OWN layout.
        let master1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"#;
        let master2_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme2.xml"/>
</Relationships>"#;

        let layout = || {
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
</p:sldLayout>"#
                .to_string()
        };
        let layout1_xml = layout();
        let layout2_xml = layout();

        // layoutA→masterA, layoutB→masterB.
        let layout1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#;
        let layout2_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster2.xml"/>
</Relationships>"#;

        // Each slide: one rect with NO explicit fill, fill comes from
        // `<p:style><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>`.
        // slide2 additionally references tx1 on a second shape so the clrMap
        // swap (tx1→lt1) is observable.
        let slide = |extra_shape: &str| {
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="StyledRect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="500000" cy="500000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      <p:style><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></p:style>
    </p:sp>
    {extra_shape}
  </p:spTree></p:cSld>
</p:sld>"#
            )
        };
        let tx1_shape = r#"<p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Tx1Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="700000" y="100000"/><a:ext cx="500000" cy="500000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      <p:style><a:fillRef idx="1"><a:schemeClr val="tx1"/></a:fillRef></p:style>
    </p:sp>"#;
        let slide1_xml = slide("");
        let slide2_xml = slide(tx1_shape);

        let slide1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;
        let slide2_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
</Relationships>"#;

        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zw = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default();
            let mut put = |path: &str, bytes: &[u8]| {
                zw.start_file(path, opts).unwrap();
                use std::io::Write;
                zw.write_all(bytes).unwrap();
            };
            put("ppt/presentation.xml", presentation_xml.as_bytes());
            put("ppt/_rels/presentation.xml.rels", pres_rels.as_bytes());
            put("ppt/theme/theme1.xml", theme1_xml.as_bytes());
            put("ppt/theme/theme2.xml", theme2_xml.as_bytes());
            put("ppt/slideMasters/slideMaster1.xml", master1_xml.as_bytes());
            put("ppt/slideMasters/slideMaster2.xml", master2_xml.as_bytes());
            put(
                "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                master1_rels.as_bytes(),
            );
            put(
                "ppt/slideMasters/_rels/slideMaster2.xml.rels",
                master2_rels.as_bytes(),
            );
            put("ppt/slideLayouts/slideLayout1.xml", layout1_xml.as_bytes());
            put("ppt/slideLayouts/slideLayout2.xml", layout2_xml.as_bytes());
            put(
                "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                layout1_rels.as_bytes(),
            );
            put(
                "ppt/slideLayouts/_rels/slideLayout2.xml.rels",
                layout2_rels.as_bytes(),
            );
            put("ppt/slides/slide1.xml", slide1_xml.as_bytes());
            put("ppt/slides/slide2.xml", slide2_xml.as_bytes());
            put("ppt/slides/_rels/slide1.xml.rels", slide1_rels.as_bytes());
            put("ppt/slides/_rels/slide2.xml.rels", slide2_rels.as_bytes());
            zw.finish().unwrap();
        }
        buf
    }

    fn first_shape_fill_color(slide: &Slide) -> Option<String> {
        slide.elements.iter().find_map(|e| match e {
            SlideElement::Shape(s) => match &s.fill {
                Some(Fill::Solid { color }) => Some(color.clone()),
                _ => None,
            },
            _ => None,
        })
    }

    fn shape_fill_color_by_name(slide: &Slide, name: &str) -> Option<String> {
        slide.elements.iter().find_map(|e| match e {
            SlideElement::Shape(s) if s.name.as_deref() == Some(name) => match &s.fill {
                Some(Fill::Solid { color }) => Some(color.clone()),
                _ => None,
            },
            _ => None,
        })
    }

    /// Core regression: each slide must resolve scheme colors against its OWN
    /// master's theme (slide→layout→master→theme), not the presentation's first
    /// theme. slide1's accent1 = masterA theme (#72A376); slide2's accent1 =
    /// masterB theme (#4F81BD).
    #[test]
    fn theme_resolved_per_slide_via_layout_master_chain() {
        let default_clr_map = r#"<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#;
        let data = build_two_master_pptx(default_clr_map);
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        assert_eq!(pres.slides.len(), 2, "expected two slides");

        let s1 = first_shape_fill_color(&pres.slides[0]);
        let s2 = first_shape_fill_color(&pres.slides[1]);
        assert_eq!(
            s1.as_deref(),
            Some("72A376"),
            "slide1 accent1 must resolve from masterA theme"
        );
        assert_eq!(
            s2.as_deref(),
            Some("4F81BD"),
            "slide2 accent1 must resolve from masterB theme"
        );
    }

    /// §19.3.1.6 clrMap: a master with `bg1`/`tx1` swapped (bg1="dk1",
    /// tx1="lt1") must remap logical scheme names. `<a:schemeClr val="tx1">`
    /// then resolves to lt1's hex (#FAFAFA), not dk1's. masterB keeps the
    /// default clrMap, so its tx1 stays dk1 (#222222).
    #[test]
    fn clr_map_remaps_logical_scheme_names() {
        // Swap bg1<->tx1 on masterA only.
        let swapped = r#"<p:clrMap bg1="dk1" tx1="lt1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#;
        let data = build_two_master_pptx(swapped);
        let pres = parse_presentation_from_bytes(&data).expect("parse");

        // slide2 (masterB, default clrMap) has the Tx1Rect: tx1 -> dk1 (#222222).
        let tx1_default = shape_fill_color_by_name(&pres.slides[1], "Tx1Rect");
        assert_eq!(
            tx1_default.as_deref(),
            Some("222222"),
            "default clrMap: tx1 must resolve to dk1"
        );

        // To observe the swap on masterA, place the same tx1 shape via a
        // dedicated parse against masterA's theme. We reuse slide1 which uses
        // masterA; assert that accent1 still resolves correctly under the swap
        // (accent slots are identity-mapped) and that tx1 on a masterA slide
        // would map to lt1. slide1 has no tx1 shape, so we assert via the
        // builder variant below.
        let s1_accent = first_shape_fill_color(&pres.slides[0]);
        assert_eq!(
            s1_accent.as_deref(),
            Some("72A376"),
            "accent1 is identity-mapped and unaffected by bg1/tx1 swap"
        );
    }

    /// Dedicated clrMap assertion on the swapped master: a slide on masterA
    /// (bg1<->tx1 swapped) resolves `<a:schemeClr val="tx1">` to lt1 (#FAFAFA).
    #[test]
    fn clr_map_tx1_resolves_to_lt1_on_swapped_master() {
        let swapped = r#"<p:clrMap bg1="dk1" tx1="lt1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#;
        let data = build_two_master_pptx_with_tx1_on_a(swapped);
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        // slide1 (masterA, swapped) has the Tx1Rect: tx1 -> lt1 (#FAFAFA).
        let tx1_swapped = shape_fill_color_by_name(&pres.slides[0], "Tx1Rect");
        assert_eq!(
            tx1_swapped.as_deref(),
            Some("FAFAFA"),
            "swapped clrMap: tx1 must resolve to lt1's hex"
        );
    }

    // Variant of build_two_master_pptx where slide1 (masterA) carries the tx1
    // shape, so the clrMap swap on masterA is directly observable.
    fn build_two_master_pptx_with_tx1_on_a(clr_map_a: &str) -> Vec<u8> {
        // Reuse the standard builder, then patch slide1 to include the tx1
        // shape by rebuilding with the tx1 shape on slide1. Simplest: build a
        // fresh deck inline mirroring build_two_master_pptx but swapping which
        // slide gets the tx1 shape. To avoid duplication we shell out to the
        // generic builder and post-process is not feasible on a zip, so we
        // construct directly here with the minimum needed parts.
        use zip::write::SimpleFileOptions;

        let presentation_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMasterA"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#;
        let pres_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMasterA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rIdThemeA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"#;
        let theme1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">
  <a:themeElements><a:clrScheme name="C">
    <a:dk1><a:srgbClr val="222222"/></a:dk1><a:lt1><a:srgbClr val="FAFAFA"/></a:lt1>
    <a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
    <a:accent1><a:srgbClr val="72A376"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2>
    <a:accent3><a:srgbClr val="0000FF"/></a:accent3><a:accent4><a:srgbClr val="FFFF00"/></a:accent4>
    <a:accent5><a:srgbClr val="FF00FF"/></a:accent5><a:accent6><a:srgbClr val="00FFFF"/></a:accent6>
    <a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="F"><a:majorFont><a:latin typeface="Arial"/></a:majorFont>
    <a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
  <a:fmtScheme name="S"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>"#;
        let master1_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
  {clr_map_a}
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483650" r:id="rIdLayout"/></p:sldLayoutIdLst>
</p:sldMaster>"#
        );
        let master1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"#;
        let layout1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
</p:sldLayout>"#;
        let layout1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#;
        let slide1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Tx1Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="700000" y="100000"/><a:ext cx="500000" cy="500000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      <p:style><a:fillRef idx="1"><a:schemeClr val="tx1"/></a:fillRef></p:style>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>"#;
        let slide1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;

        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zw = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default();
            let mut put = |path: &str, bytes: &[u8]| {
                zw.start_file(path, opts).unwrap();
                use std::io::Write;
                zw.write_all(bytes).unwrap();
            };
            put("ppt/presentation.xml", presentation_xml.as_bytes());
            put("ppt/_rels/presentation.xml.rels", pres_rels.as_bytes());
            put("ppt/theme/theme1.xml", theme1_xml.as_bytes());
            put("ppt/slideMasters/slideMaster1.xml", master1_xml.as_bytes());
            put(
                "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                master1_rels.as_bytes(),
            );
            put("ppt/slideLayouts/slideLayout1.xml", layout1_xml.as_bytes());
            put(
                "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                layout1_rels.as_bytes(),
            );
            put("ppt/slides/slide1.xml", slide1_xml.as_bytes());
            put("ppt/slides/_rels/slide1.xml.rels", slide1_rels.as_bytes());
            zw.finish().unwrap();
        }
        buf
    }

    /// Single-master deck whose slide1 carries a `<p:clrMapOvr>` with the given
    /// inner element (`<a:overrideClrMapping .../>` or `<a:masterClrMapping/>`).
    /// The master keeps the DEFAULT clrMap (tx1→dk1). slide1 has the Tx1Rect
    /// (`<a:schemeClr val="tx1">`), so the override's tx1→slot remap is directly
    /// observable. Theme hex: dk1=#222222, lt1=#FAFAFA, accent1=#72A376.
    ///
    /// `layout_clr_map_ovr_inner` optionally injects a `<p:clrMapOvr>` on the
    /// LAYOUT (CT_SlideLayout: right after `</p:cSld>`, §20.1.6 / pml.xsd) so the
    /// slide↔layout override precedence can be exercised.
    fn build_clr_map_ovr_pptx(clr_map_ovr_inner: &str) -> Vec<u8> {
        build_clr_map_ovr_pptx_with_layout(clr_map_ovr_inner, None)
    }

    fn build_clr_map_ovr_pptx_with_layout(
        clr_map_ovr_inner: &str,
        layout_clr_map_ovr_inner: Option<&str>,
    ) -> Vec<u8> {
        use zip::write::SimpleFileOptions;

        let presentation_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMasterA"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#;
        let pres_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMasterA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rIdThemeA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"#;
        let theme1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">
  <a:themeElements><a:clrScheme name="C">
    <a:dk1><a:srgbClr val="222222"/></a:dk1><a:lt1><a:srgbClr val="FAFAFA"/></a:lt1>
    <a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
    <a:accent1><a:srgbClr val="72A376"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2>
    <a:accent3><a:srgbClr val="0000FF"/></a:accent3><a:accent4><a:srgbClr val="FFFF00"/></a:accent4>
    <a:accent5><a:srgbClr val="FF00FF"/></a:accent5><a:accent6><a:srgbClr val="00FFFF"/></a:accent6>
    <a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="F"><a:majorFont><a:latin typeface="Arial"/></a:majorFont>
    <a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
  <a:fmtScheme name="S"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>"#;
        // Master keeps the DEFAULT clrMap (tx1→dk1) so the override is the ONLY
        // thing that can remap tx1; the assertion is unambiguous.
        let master1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483650" r:id="rIdLayout"/></p:sldLayoutIdLst>
</p:sldMaster>"#;
        let master1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"#;
        // CT_SlideLayout: <p:clrMapOvr> comes right after </p:cSld> (pml.xsd).
        let layout_clr_map_ovr = layout_clr_map_ovr_inner
            .map(|inner| format!("<p:clrMapOvr>{inner}</p:clrMapOvr>"))
            .unwrap_or_default();
        let layout1_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
  {layout_clr_map_ovr}
</p:sldLayout>"#
        );
        let layout1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#;
        // CT_Slide: <p:clrMapOvr> comes right after </p:cSld> (ECMA-376 §19.3.1.7).
        // An empty `clr_map_ovr_inner` means "no <p:clrMapOvr> on the slide at all".
        let slide_clr_map_ovr = if clr_map_ovr_inner.is_empty() {
            String::new()
        } else {
            format!("<p:clrMapOvr>{clr_map_ovr_inner}</p:clrMapOvr>")
        };
        let slide1_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Tx1Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="700000" y="100000"/><a:ext cx="500000" cy="500000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      <p:style><a:fillRef idx="1"><a:schemeClr val="tx1"/></a:fillRef></p:style>
    </p:sp>
  </p:spTree></p:cSld>
  {slide_clr_map_ovr}
</p:sld>"#
        );
        let slide1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;

        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zw = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default();
            let mut put = |path: &str, bytes: &[u8]| {
                zw.start_file(path, opts).unwrap();
                use std::io::Write;
                zw.write_all(bytes).unwrap();
            };
            put("ppt/presentation.xml", presentation_xml.as_bytes());
            put("ppt/_rels/presentation.xml.rels", pres_rels.as_bytes());
            put("ppt/theme/theme1.xml", theme1_xml.as_bytes());
            put("ppt/slideMasters/slideMaster1.xml", master1_xml.as_bytes());
            put(
                "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                master1_rels.as_bytes(),
            );
            put("ppt/slideLayouts/slideLayout1.xml", layout1_xml.as_bytes());
            put(
                "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                layout1_rels.as_bytes(),
            );
            put("ppt/slides/slide1.xml", slide1_xml.as_bytes());
            put("ppt/slides/_rels/slide1.xml.rels", slide1_rels.as_bytes());
            zw.finish().unwrap();
        }
        buf
    }

    /// §19.3.1.7 clrMapOvr / §20.1.6.8 overrideClrMapping: a slide whose
    /// `<p:clrMapOvr>` carries `<a:overrideClrMapping>` with bg1/tx1 swapped
    /// (bg1="dk1", tx1="lt1") must use that mapping IN PLACE OF the master's.
    /// The master keeps the default clrMap (tx1→dk1, #222222), so the override
    /// flips tx1 to lt1 (#FAFAFA). The other 10 attrs are default.
    #[test]
    fn clr_map_ovr_override_remaps_logical_scheme_names() {
        let override_inner = r#"<a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#;
        let data = build_clr_map_ovr_pptx(override_inner);
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        assert_eq!(pres.slides.len(), 1, "expected one slide");

        // tx1 under the override → lt1 (#FAFAFA), NOT the master's dk1 (#222222).
        let tx1 = shape_fill_color_by_name(&pres.slides[0], "Tx1Rect");
        assert_eq!(
            tx1.as_deref(),
            Some("FAFAFA"),
            "overrideClrMapping (tx1=lt1) must replace the master clrMap (tx1=dk1)"
        );
    }

    /// §20.1.6.6 + Annex L.3.2.5 (FINDING 3): a LAYOUT-level `overrideClrMapping`
    /// (swap bg1/tx1) is inherited by its slides; a slide carrying an explicit
    /// `<a:masterClrMapping/>` means "no override of MY OWN" and therefore inherits
    /// the LAYOUT's override (NOT a bypass to the master's raw mapping). So the
    /// slide's tx1 shape resolves through the layout override → lt1 (#FAFAFA), not
    /// the master default tx1→dk1 (#222222).
    #[test]
    fn slide_master_clr_mapping_inherits_layout_override() {
        let layout_override = r#"<a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#;
        let data =
            build_clr_map_ovr_pptx_with_layout("<a:masterClrMapping/>", Some(layout_override));
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        assert_eq!(pres.slides.len(), 1, "expected one slide");

        let tx1 = shape_fill_color_by_name(&pres.slides[0], "Tx1Rect");
        assert_eq!(
            tx1.as_deref(),
            Some("FAFAFA"),
            "slide masterClrMapping inherits the layout override (tx1=lt1), not the master tx1=dk1"
        );
    }

    /// §20.1.6.6 + Annex L.3.2.5 (FINDING 3): a LAYOUT-level `overrideClrMapping`
    /// is inherited by a slide that has NO `<p:clrMapOvr>` at all (the common
    /// inheritance case). Same expected result as the masterClrMapping variant.
    #[test]
    fn layout_override_inherited_by_slide_without_clr_map_ovr() {
        let layout_override = r#"<a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#;
        // Empty slide-level inner ⇒ the builder omits <p:clrMapOvr> on the slide.
        let data = build_clr_map_ovr_pptx_with_layout("", Some(layout_override));
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        assert_eq!(pres.slides.len(), 1, "expected one slide");

        let tx1 = shape_fill_color_by_name(&pres.slides[0], "Tx1Rect");
        assert_eq!(
            tx1.as_deref(),
            Some("FAFAFA"),
            "a slide with no clrMapOvr inherits the layout override (tx1=lt1)"
        );
    }

    /// Control that makes the two FINDING 3 tests load-bearing: with NO layout
    /// override and a slide `<a:masterClrMapping/>`, tx1 must stay the master
    /// default dk1 (#222222). The ONLY difference from
    /// `slide_master_clr_mapping_inherits_layout_override` is the presence of the
    /// layout override — so that test genuinely proves layout inheritance, not a
    /// vacuous pass.
    #[test]
    fn slide_master_clr_mapping_without_layout_override_uses_master() {
        let data = build_clr_map_ovr_pptx_with_layout("<a:masterClrMapping/>", None);
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        assert_eq!(pres.slides.len(), 1, "expected one slide");

        let tx1 = shape_fill_color_by_name(&pres.slides[0], "Tx1Rect");
        assert_eq!(
            tx1.as_deref(),
            Some("222222"),
            "with no layout override, masterClrMapping resolves tx1 from the master (dk1)"
        );
    }

    /// The slide's resolved background fill colour, if it is a solid fill.
    fn slide_bg_color(slide: &Slide) -> Option<String> {
        match &slide.background {
            Some(Fill::Solid { color }) => Some(color.clone()),
            _ => None,
        }
    }

    /// Single-master deck like `build_clr_map_ovr_pptx`, but the MASTER carries a
    /// `<p:bg>` whose fill is `<a:schemeClr val="bg1"/>` and the SLIDE has NO
    /// background of its own, so the slide inherits the master background through
    /// the slide→layout→master chain (§19.3.1.42). The slide carries a
    /// `<p:clrMapOvr>` with the given inner element. Theme hex: dk1=#222222,
    /// lt1=#FAFAFA. With the default clrMap bg1→lt1 ⇒ #FAFAFA; under an override
    /// that maps bg1→dk1 the inherited master background must flip to #222222.
    fn build_clr_map_ovr_master_bg_pptx(slide_clr_map_ovr_inner: &str) -> Vec<u8> {
        use zip::write::SimpleFileOptions;

        let presentation_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMasterA"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#;
        let pres_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMasterA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rIdThemeA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"#;
        let theme1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">
  <a:themeElements><a:clrScheme name="C">
    <a:dk1><a:srgbClr val="222222"/></a:dk1><a:lt1><a:srgbClr val="FAFAFA"/></a:lt1>
    <a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
    <a:accent1><a:srgbClr val="72A376"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2>
    <a:accent3><a:srgbClr val="0000FF"/></a:accent3><a:accent4><a:srgbClr val="FFFF00"/></a:accent4>
    <a:accent5><a:srgbClr val="FF00FF"/></a:accent5><a:accent6><a:srgbClr val="00FFFF"/></a:accent6>
    <a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="F"><a:majorFont><a:latin typeface="Arial"/></a:majorFont>
    <a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
  <a:fmtScheme name="S"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>"#;
        // Master keeps the DEFAULT clrMap (bg1→lt1). Its <p:bg> uses schemeClr
        // bg1, so without an override the inherited background is lt1 (#FAFAFA).
        let master1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483650" r:id="rIdLayout"/></p:sldLayoutIdLst>
</p:sldMaster>"#;
        let master1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"#;
        // Layout has NO background of its own → the slide falls through to the
        // master background.
        let layout1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
</p:sldLayout>"#;
        let layout1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#;
        // Slide has NO <p:bg> of its own → inherits the master background.
        let slide1_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
  <p:clrMapOvr>{slide_clr_map_ovr_inner}</p:clrMapOvr>
</p:sld>"#
        );
        let slide1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;

        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zw = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default();
            let mut put = |path: &str, bytes: &[u8]| {
                zw.start_file(path, opts).unwrap();
                use std::io::Write;
                zw.write_all(bytes).unwrap();
            };
            put("ppt/presentation.xml", presentation_xml.as_bytes());
            put("ppt/_rels/presentation.xml.rels", pres_rels.as_bytes());
            put("ppt/theme/theme1.xml", theme1_xml.as_bytes());
            put("ppt/slideMasters/slideMaster1.xml", master1_xml.as_bytes());
            put(
                "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                master1_rels.as_bytes(),
            );
            put("ppt/slideLayouts/slideLayout1.xml", layout1_xml.as_bytes());
            put(
                "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                layout1_rels.as_bytes(),
            );
            put("ppt/slides/slide1.xml", slide1_xml.as_bytes());
            put("ppt/slides/_rels/slide1.xml.rels", slide1_rels.as_bytes());
            zw.finish().unwrap();
        }
        buf
    }

    /// §19.3.1.7 / §20.1.6.8 (FINDING 1): a master-inherited background that uses
    /// a scheme colour (`<p:bg>` schemeClr bg1) MUST resolve through the slide's
    /// effective override mapping, not the master's frozen mapping. The slide has
    /// no own background; its `<a:overrideClrMapping>` swaps bg1→dk1, so the
    /// inherited master background must become dk1 (#222222), NOT the master
    /// default bg1→lt1 (#FAFAFA).
    #[test]
    fn clr_map_ovr_flips_master_inherited_background() {
        let override_inner = r#"<a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#;
        let data = build_clr_map_ovr_master_bg_pptx(override_inner);
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        assert_eq!(pres.slides.len(), 1, "expected one slide");

        let bg = slide_bg_color(&pres.slides[0]);
        assert_eq!(
            bg.as_deref(),
            Some("222222"),
            "master-inherited background (schemeClr bg1) must honor the slide override (bg1=dk1)"
        );
    }

    /// Control for `clr_map_ovr_flips_master_inherited_background`: with a
    /// `<a:masterClrMapping/>` (no override of its own) the inherited master
    /// background keeps the master default bg1→lt1 (#FAFAFA).
    #[test]
    fn master_inherited_background_default_without_override() {
        let data = build_clr_map_ovr_master_bg_pptx("<a:masterClrMapping/>");
        let pres = parse_presentation_from_bytes(&data).expect("parse");
        assert_eq!(pres.slides.len(), 1, "expected one slide");

        let bg = slide_bg_color(&pres.slides[0]);
        assert_eq!(
            bg.as_deref(),
            Some("FAFAFA"),
            "without an override the master background resolves bg1→lt1"
        );
    }

    /// FINDING 2 (perf guard): `parse_clr_map_ovr` must short-circuit to `None`
    /// when the XML contains no `clrMapOvr` element (avoiding a second full parse),
    /// while still returning `Some` for an `overrideClrMapping` and `None` for an
    /// explicit `masterClrMapping`. The fast path must not change any of these
    /// observable results.
    #[test]
    fn parse_clr_map_ovr_guard_and_results() {
        let ns = r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main""#;

        // No <p:clrMapOvr> at all → None (and the guard skips the parse entirely).
        let no_ovr = format!(r#"<p:sld {ns}><p:cSld><p:spTree/></p:cSld></p:sld>"#);
        assert!(
            parse_clr_map_ovr(&no_ovr).is_none(),
            "absent clrMapOvr must yield None"
        );

        // Explicit <a:masterClrMapping/> → None (inherit).
        let master = format!(
            r#"<p:sld {ns}><p:cSld><p:spTree/></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"#
        );
        assert!(
            parse_clr_map_ovr(&master).is_none(),
            "masterClrMapping must yield None"
        );

        // <a:overrideClrMapping> → Some(map) with the parsed logical→slot attrs.
        let ovr = format!(
            r#"<p:sld {ns}><p:cSld><p:spTree/></p:cSld><p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1"/></p:clrMapOvr></p:sld>"#
        );
        let parsed = parse_clr_map_ovr(&ovr).expect("overrideClrMapping must yield Some");
        assert_eq!(parsed.get("bg1").map(String::as_str), Some("dk1"));
        assert_eq!(parsed.get("tx1").map(String::as_str), Some("lt1"));
    }

    // ── Chart axis titles + chartSpace border (parity with xlsx) ──────────
    //
    // These exercise `parse_legacy_chart` directly with inline chart XML so we
    // can assert the newly-parsed fields without a full .pptx fixture. Mirrors
    // the xlsx parser's chart.rs coverage.

    /// A clustered bar chart whose category (X) and value (Y) axes both carry a
    /// `<c:title>` with explicit run props (sz / b / solidFill), plus an
    /// explicit `<c:chartSpace><c:spPr><a:ln>` border.
    fn bar_chart_with_axis_titles_xml() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Series 1</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:strCache>
            <c:pt idx="0"><c:v>A</c:v></c:pt>
            <c:pt idx="1"><c:v>B</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>3</c:v></c:pt>
            <c:pt idx="1"><c:v>7</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
      <c:catAx>
        <c:axId val="111"/>
        <c:axPos val="b"/>
        <c:title>
          <c:tx><c:rich><a:p><a:pPr><a:defRPr sz="1000" b="1">
            <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
          </a:defRPr></a:pPr><a:r><a:t>Category Axis</a:t></a:r></a:p></c:rich></c:tx>
        </c:title>
      </c:catAx>
      <c:valAx>
        <c:axId val="222"/>
        <c:axPos val="l"/>
        <c:title>
          <c:tx><c:rich><a:p><a:pPr><a:defRPr sz="1200" b="0">
            <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
          </a:defRPr></a:pPr><a:r><a:t>Value Axis</a:t></a:r></a:p></c:rich></c:tx>
        </c:title>
      </c:valAx>
    </c:plotArea>
  </c:chart>
  <c:spPr>
    <a:ln w="19050"><a:solidFill><a:srgbClr val="1B4332"/></a:solidFill></a:ln>
  </c:spPr>
</c:chartSpace>"#
    }

    #[test]
    fn chart_parses_cat_and_val_axis_titles_with_props() {
        let theme = HashMap::new();
        let c = parse_legacy_chart(bar_chart_with_axis_titles_xml(), &theme)
            .expect("legacy chart should parse");
        let c = &c.chart;

        assert_eq!(c.cat_axis_title.as_deref(), Some("Category Axis"));
        assert_eq!(c.cat_axis_title_font_size_hpt, Some(1000));
        assert_eq!(c.cat_axis_title_font_bold, Some(true));
        assert_eq!(c.cat_axis_title_font_color.as_deref(), Some("FF0000"));

        assert_eq!(c.val_axis_title.as_deref(), Some("Value Axis"));
        assert_eq!(c.val_axis_title_font_size_hpt, Some(1200));
        assert_eq!(c.val_axis_title_font_bold, Some(false));
        assert_eq!(c.val_axis_title_font_color.as_deref(), Some("00FF00"));
    }

    #[test]
    fn chart_parses_explicit_chartspace_border() {
        let theme = HashMap::new();
        let c = parse_legacy_chart(bar_chart_with_axis_titles_xml(), &theme)
            .expect("legacy chart should parse");
        let c = &c.chart;

        assert_eq!(c.chart_border_color.as_deref(), Some("1B4332"));
        assert_eq!(c.chart_border_width_emu, Some(19050));
    }

    #[test]
    fn chart_border_nofill_yields_no_color() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
  <c:spPr>
    <a:ln w="12700"><a:noFill/></a:ln>
  </c:spPr>
</c:chartSpace>"#;
        let theme = HashMap::new();
        let c = parse_legacy_chart(xml, &theme).expect("legacy chart should parse");
        // noFill explicitly turns the border OFF → no color, even though @w is set.
        assert_eq!(c.chart.chart_border_color, None);
    }

    #[test]
    fn chart_without_axis_titles_leaves_them_none() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:axPos val="b"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:axPos val="l"/></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;
        let theme = HashMap::new();
        let c = parse_legacy_chart(xml, &theme).expect("legacy chart should parse");
        assert_eq!(c.chart.cat_axis_title, None);
        assert_eq!(c.chart.val_axis_title, None);
        assert_eq!(c.chart.chart_border_color, None);
        assert_eq!(c.chart.chart_border_width_emu, None);
    }

    /// A combo chart: `<c:barChart>` (Revenue, primary left axis) +
    /// `<c:lineChart>` (Gross margin, SECONDARY right axis). Mirrors sample-14
    /// slide-8. The line series must be tagged `series_type = "line"` and bound
    /// to the secondary axis, and the secondary `<c:valAx>` (axPos="r",
    /// crosses="max", min=0 max=100, title "Gross margin (%)") parsed into
    /// `secondary_val_axis` — while the primary axis fields stay the Revenue
    /// axis.
    fn combo_bar_line_secondary_axis_xml() -> &'static str {
        r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Revenue ($M)</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:strCache>
            <c:pt idx="0"><c:v>FY22</c:v></c:pt>
            <c:pt idx="1"><c:v>FY23</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>18.9</c:v></c:pt>
            <c:pt idx="1"><c:v>26.5</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="100"/>
        <c:axId val="200"/>
      </c:barChart>
      <c:lineChart>
        <c:ser>
          <c:idx val="1"/>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Gross margin (%)</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:strCache>
            <c:pt idx="0"><c:v>FY22</c:v></c:pt>
            <c:pt idx="1"><c:v>FY23</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>68</c:v></c:pt>
            <c:pt idx="1"><c:v>71</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="300"/>
        <c:axId val="400"/>
      </c:lineChart>
      <c:catAx><c:axId val="100"/><c:axPos val="b"/></c:catAx>
      <c:valAx>
        <c:axId val="200"/>
        <c:axPos val="l"/>
        <c:crosses val="autoZero"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue ($M)</a:t></a:r></a:p></c:rich></c:tx></c:title>
      </c:valAx>
      <c:valAx>
        <c:axId val="400"/>
        <c:scaling><c:max val="100"/><c:min val="0"/></c:scaling>
        <c:axPos val="r"/>
        <c:crosses val="max"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>Gross margin (%)</a:t></a:r></a:p></c:rich></c:tx></c:title>
      </c:valAx>
      <c:catAx><c:axId val="300"/><c:delete val="1"/><c:axPos val="b"/></c:catAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#
    }

    #[test]
    fn combo_chart_tags_line_series_and_secondary_axis() {
        let theme = HashMap::new();
        let c = parse_legacy_chart(combo_bar_line_secondary_axis_xml(), &theme)
            .expect("combo chart should parse");
        let c = &c.chart;

        // Primary type is bar (bar group wins).
        assert_eq!(c.chart_type, "clusteredBar");
        assert_eq!(c.series.len(), 2, "both bar and line series parsed");

        // Bar series: primary axis, no line override.
        assert_eq!(c.series[0].name, "Revenue ($M)");
        assert_eq!(c.series[0].series_type, None);
        assert_eq!(c.series[0].use_secondary_axis, None);

        // Line series: tagged "line" + bound to the secondary axis.
        assert_eq!(c.series[1].name, "Gross margin (%)");
        assert_eq!(c.series[1].series_type.as_deref(), Some("line"));
        assert_eq!(c.series[1].use_secondary_axis, Some(true));

        // Primary value-axis fields stay the Revenue (left) axis.
        assert_eq!(c.val_axis_title.as_deref(), Some("Revenue ($M)"));

        // Secondary axis parsed from the right-hand valAx.
        let sec = c
            .secondary_val_axis
            .as_ref()
            .expect("secondary value axis present");
        assert_eq!(sec.min, Some(0.0));
        assert_eq!(sec.max, Some(100.0));
        assert_eq!(sec.title.as_deref(), Some("Gross margin (%)"));
    }

    #[test]
    fn single_axis_chart_has_no_secondary() {
        let theme = HashMap::new();
        let c = parse_legacy_chart(bar_chart_with_axis_titles_xml(), &theme)
            .expect("legacy chart should parse");
        assert!(c.chart.secondary_val_axis.is_none());
        assert_eq!(c.chart.series[0].series_type, None);
        assert_eq!(c.chart.series[0].use_secondary_axis, None);
    }

    #[test]
    fn scatter_bottom_valax_title_maps_to_cat_axis() {
        // Scatter charts have TWO <c:valAx> and no <c:catAx>. The bottom one
        // (axPos="b") is the horizontal axis → its title is the cat-axis title;
        // the left one (axPos="l") is the value-axis title. Same disambiguation
        // as the xlsx parser.
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:scatterChart>
        <c:scatterStyle val="lineMarker"/>
        <c:ser>
          <c:idx val="0"/>
          <c:xVal><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>1</c:v></c:pt>
            <c:pt idx="1"><c:v>2</c:v></c:pt>
          </c:numCache></c:numRef></c:xVal>
          <c:yVal><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>10</c:v></c:pt>
            <c:pt idx="1"><c:v>20</c:v></c:pt>
          </c:numCache></c:numRef></c:yVal>
        </c:ser>
      </c:scatterChart>
      <c:valAx>
        <c:axId val="100"/>
        <c:axPos val="b"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>X Bottom</a:t></a:r></a:p></c:rich></c:tx></c:title>
      </c:valAx>
      <c:valAx>
        <c:axId val="200"/>
        <c:axPos val="l"/>
        <c:title><c:tx><c:rich><a:p><a:r><a:t>Y Left</a:t></a:r></a:p></c:rich></c:tx></c:title>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;
        let theme = HashMap::new();
        let c = parse_legacy_chart(xml, &theme).expect("scatter chart should parse");
        assert_eq!(c.chart.chart_type, "scatter");
        // Bottom valAx → X → cat-axis title.
        assert_eq!(c.chart.cat_axis_title.as_deref(), Some("X Bottom"));
        // Left valAx → Y → val-axis title.
        assert_eq!(c.chart.val_axis_title.as_deref(), Some("Y Left"));
    }

    #[test]
    fn chart_parses_axis_tick_label_bold_flags() {
        // The bold flags for tick labels (title bold + cat/val tick-label bold)
        // are parsed from `<c:title>...defRPr@b` and `<c:txPr>...defRPr@b`.
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr b="1"/></a:pPr>
      <a:r><a:t>My Chart</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/><c:axPos val="b"/>
        <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr b="1"/></a:pPr><a:endParaRPr/></a:p></c:txPr>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/><c:axPos val="l"/>
        <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr b="0"/></a:pPr><a:endParaRPr/></a:p></c:txPr>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;
        let theme = HashMap::new();
        let c = parse_legacy_chart(xml, &theme).expect("legacy chart should parse");
        assert_eq!(c.chart.title_font_bold, Some(true));
        assert_eq!(c.chart.cat_axis_font_bold, Some(true));
        assert_eq!(c.chart.val_axis_font_bold, Some(false));
    }

    /// Regression for the `PathCmd::ArcTo` serde naming bug: the enum-level
    /// `#[serde(tag = "cmd", rename_all = "camelCase")]` renames only the variant
    /// tag, not the struct-variant fields, so `st_ang`/`sw_ang` serialized in
    /// snake_case. The TS `PathCmd` (core/src/types/common.ts) reads `stAng`/
    /// `swAng`, so the angles came back `undefined` → `NaN` coordinates and the
    /// arc (plus everything after it) vanished. A non-degenerate arc (positive
    /// `wR`/`hR`) is essential: a degenerate arc short-circuits before the
    /// angles are read, which is why the original arrow sample (degenerate arcs
    /// only) never surfaced this.
    #[test]
    fn arcto_serializes_angle_fields_as_camel_case() {
        // 90° arc: swAng = 90 * 60000 = 5400000 in OOXML 60000ths of a degree.
        let xml = r#"<custGeom xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
  <pathLst>
    <path w="100" h="100">
      <moveTo><pt x="100" y="50"/></moveTo>
      <arcTo wR="50" hR="50" stAng="0" swAng="5400000"/>
    </path>
  </pathLst>
</custGeom>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let subpaths = parse_cust_geom(doc.root_element());
        let json = serde_json::to_string(&subpaths).expect("custGeom should serialize");

        // The two camelCase keys the TS renderer reads must be present…
        assert!(
            json.contains("\"stAng\""),
            "ArcTo must serialize stAng (camelCase); got: {json}"
        );
        assert!(
            json.contains("\"swAng\""),
            "ArcTo must serialize swAng (camelCase); got: {json}"
        );
        // …and the buggy snake_case keys must be gone.
        assert!(
            !json.contains("\"st_ang\""),
            "ArcTo must not emit snake_case st_ang; got: {json}"
        );
        assert!(
            !json.contains("\"sw_ang\""),
            "ArcTo must not emit snake_case sw_ang; got: {json}"
        );
    }

    /// Full value-level round-trip through the serde tag + camelCase fields:
    /// re-deserializing the serialized JSON must reproduce the angle values,
    /// proving the rename is symmetric (Serialize + Deserialize both use the
    /// camelCase keys) and the 60000ths→degrees conversion is intact.
    #[test]
    fn arcto_round_trips_angles_through_camel_case_json() {
        let xml = r#"<custGeom xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
  <pathLst>
    <path w="200" h="100">
      <moveTo><pt x="200" y="50"/></moveTo>
      <arcTo wR="100" hR="50" stAng="2700000" swAng="-5400000"/>
    </path>
  </pathLst>
</custGeom>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let subpaths = parse_cust_geom(doc.root_element());
        let json = serde_json::to_string(&subpaths).unwrap();
        let back: Vec<Vec<PathCmd>> =
            serde_json::from_str(&json).expect("camelCase JSON must deserialize back");
        let arc = back[0]
            .iter()
            .find(|c| matches!(c, PathCmd::ArcTo { .. }))
            .expect("arc command should be present");
        match arc {
            PathCmd::ArcTo {
                wr,
                hr,
                st_ang,
                sw_ang,
            } => {
                // wR/hR normalised by path w/h; angles converted from 60000ths.
                assert!((wr - 0.5).abs() < 1e-9, "wr = {wr}"); // 100/200
                assert!((hr - 0.5).abs() < 1e-9, "hr = {hr}"); // 50/100
                assert!((st_ang - 45.0).abs() < 1e-9, "st_ang = {st_ang}"); // 2700000/60000
                assert!((sw_ang + 90.0).abs() < 1e-9, "sw_ang = {sw_ang}"); // -5400000/60000
            }
            _ => unreachable!(),
        }
    }
}
