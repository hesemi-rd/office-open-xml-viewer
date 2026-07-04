//! Shared OOXML chart-XML extractors used by both the xlsx and pptx Rust
//! parsers.
//!
//! Both crates parse a chart `<c:chartSpace>` (and the modern
//! `<cx:chartSpace>` for waterfall / treemap / box-and-whisker etc.) but
//! historically did so with two near-identical bodies sitting in
//! `packages/xlsx/parser/src/lib.rs` and `packages/pptx/parser/src/lib.rs`.
//! The result was that fields added on one side stayed missing on the other
//! until somebody noticed (e.g. PowerPoint sample-2 slide-7 displaying its
//! legend on the right because the pptx adapter had a hard-coded
//! `legendPos: null` while xlsx already passed it through).
//!
//! This module hosts the helpers that don't need any crate-private state:
//! they're pure XML probes that take a roxmltree node and return the parsed
//! property. The data-structure layer (xlsx's `ChartData`, pptx's
//! `ChartElement`) intentionally stays in each crate so we don't pull
//! schema-specific types into the shared one.
//!
//! ## Namespace handling
//!
//! All helpers match elements by local name only. Real chart documents put
//! everything under either the `c:` (chart 2006) or `cx:` (chartEx 2014)
//! namespace and never mix non-chart elements at these paths, so the strict
//! `tag_name().namespace() == Some(c_ns)` check in xlsx adds nothing in
//! practice — this module drops it for symmetry with the pptx side and to
//! keep the API simple. If a future format wedges a non-chart element into
//! `<c:plotArea>` the caller can pre-filter before delegating here.
//!
//! All field references are to ECMA-376 / ISO-29500 part 1 §21.2 (DrawingML
//! Charts) unless stated otherwise.

use roxmltree::Node;
use serde::{Deserialize, Serialize};

// ============================================================================
// Shared chart data model
// ============================================================================
//
// These structs are the Rust mirror of the TypeScript `ChartModel` in
// `packages/core/src/types/chart.ts`. Both the pptx and xlsx Rust parsers build
// a `ChartModel` and emit it as a single nested `chart` object, so the TS
// renderer (`@silurus/ooxml-core`'s `renderChart`) receives a value that is
// already `ChartModel`-shaped and needs no per-field adapter.
//
// Field-for-field parity with the TS interface is the contract. Serde
// `rename_all = "camelCase"` matches the TS key names. The REQUIRED TS fields
// (no `?`) are serialized unconditionally so the wire object always carries
// them — an `Option<T>` REQUIRED field emits `null` when `None` (matching
// `T | null`), and a `bool`/`Vec` REQUIRED field emits `false`/`[]`. The
// OPTIONAL TS fields (`field?: …`) keep `skip_serializing_if` so they drop off
// the wire when unset; the renderer treats a missing key and an explicit `null`
// identically (every read is `?? default` / `!= null`), so this is
// render-equivalent to emitting `null`.
//
// All field references are ECMA-376 / ISO-29500 part 1 §21.2 (DrawingML Charts)
// as documented on the TS side; see that file for the per-field spec citations.

/// Mirror of TS `ChartModel`. Built by each parser and emitted as the single
/// `chart` object consumed by the core chart renderer.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartModel {
    // ── Required (always serialized) ────────────────────────────────────────
    pub chart_type: String,
    pub title: Option<String>,
    pub categories: Vec<String>,
    pub series: Vec<ChartSeries>,
    pub show_data_labels: bool,
    pub val_min: Option<f64>,
    pub val_max: Option<f64>,
    pub cat_axis_title: Option<String>,
    pub val_axis_title: Option<String>,
    pub cat_axis_hidden: bool,
    pub val_axis_hidden: bool,
    pub cat_axis_line_hidden: bool,
    pub val_axis_line_hidden: bool,
    pub plot_area_bg: Option<String>,
    pub chart_bg: Option<String>,
    pub show_legend: bool,
    pub legend_pos: Option<String>,
    pub cat_axis_cross_between: String,
    pub val_axis_major_tick_mark: String,
    pub cat_axis_major_tick_mark: String,
    pub title_font_size_hpt: Option<i32>,
    pub title_font_color: Option<String>,
    pub title_font_face: Option<String>,
    pub cat_axis_font_size_hpt: Option<i32>,
    pub val_axis_font_size_hpt: Option<i32>,
    pub data_label_font_size_hpt: Option<i32>,
    pub subtotal_indices: Vec<u32>,
    // ── Optional (skipped when unset) ───────────────────────────────────────
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_minor_tick_mark: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_minor_tick_mark: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend_manual_layout: Option<LegendManualLayout>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_format_code: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bar_gap_width: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bar_overlap: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_position: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_format_code: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_title_font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_title_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_title_font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_border_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_border_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_crosses: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_crosses_at: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_crosses: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_crosses_at: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_axis_line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_format_code: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_min: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cat_axis_max: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_manual_layout: Option<ChartManualLayout>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_area_manual_layout: Option<ChartManualLayout>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scatter_style: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub radar_style: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub secondary_val_axis: Option<SecondaryValueAxis>,
    /// `<c:date1904>` (ECMA-376 §21.2.2.38). `true` = the chart's serial dates
    /// resolve against the 1904 date system. Omitted from JSON when false (the
    /// default 1900 system) for wire parity.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub date1904: bool,
}

/// Mirror of TS `ChartSeries`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartSeries {
    pub name: String,
    pub color: Option<String>,
    pub values: Vec<Option<f64>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_point_colors: Option<Vec<Option<String>>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_colors: Option<Vec<Option<String>>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub series_type: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub use_secondary_axis: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub categories: Option<Vec<String>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_marker: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_format_code: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_symbol: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_fill: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_line: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_point_overrides: Option<Vec<ChartDataPointOverride>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_label_overrides: Option<Vec<ChartDataLabelOverride>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub series_data_labels: Option<ChartSeriesDataLabels>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub err_bars: Option<Vec<ChartErrBars>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bubble_sizes: Option<Vec<Option<f64>>>,
    /// `<c:ser><c:smooth val>` (ECMA-376 §21.2.2.194) — line/area series flag
    /// requesting a smoothed (spline) curve. `None` (omitted) = straight
    /// polyline (the default); only serialized when the file sets it.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub smooth: Option<bool>,
}

/// Mirror of TS `ChartDataPointOverride`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartDataPointOverride {
    pub idx: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_symbol: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_fill: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker_line: Option<String>,
}

/// Mirror of TS `ChartDataLabelOverride`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartDataLabelOverride {
    pub idx: u32,
    pub text: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_bold: Option<bool>,
}

/// Mirror of TS `ChartSeriesDataLabels`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
#[serde(rename_all = "camelCase")]
pub struct ChartSeriesDataLabels {
    pub show_val: bool,
    pub show_cat_name: bool,
    pub show_ser_name: bool,
    pub show_percent: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub format_code: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size_hpt: Option<i32>,
}

/// Mirror of TS `ChartErrBars`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartErrBars {
    pub dir: String,
    pub bar_type: String,
    pub plus: Vec<Option<f64>>,
    pub minus: Vec<Option<f64>>,
    pub no_end_cap: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dash: Option<String>,
}

/// Mirror of TS `SecondaryValueAxis`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SecondaryValueAxis {
    pub min: Option<f64>,
    pub max: Option<f64>,
    pub title: Option<String>,
    pub hidden: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub format_code: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    pub line_hidden: bool,
    pub major_tick_mark: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_size_hpt: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_font_color: Option<String>,
}

/// Mirror of TS `ChartManualLayout`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ChartManualLayout {
    pub x_mode: String,
    pub y_mode: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout_target: Option<String>,
    pub x: f64,
    pub y: f64,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub w: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h: Option<f64>,
}

/// Mirror of TS `LegendManualLayout`.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct LegendManualLayout {
    pub x_mode: String,
    pub y_mode: String,
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
}

/// Combine a chart-type family (`bar` / `line` / `area`) with its bar direction
/// and grouping into the canonical `ChartModel.chart_type` vocabulary the core
/// renderer dispatches on.
///
/// This is the Rust home of the logic the xlsx TS renderer used to run in
/// `canonicalChartType` (pptx already emitted the canonical string). `bar_dir`
/// is ECMA-376 §21.2.3.4 `ST_BarDir`: `"bar"` = horizontal, `"col"` (or any
/// other value) = vertical. `grouping` is §21.2.3.17 `ST_Grouping`. Non-bar /
/// non-line / non-area families are returned unchanged.
pub fn canonical_chart_type(chart_type: &str, bar_dir: &str, grouping: &str) -> String {
    match chart_type {
        "bar" => {
            let is_h = bar_dir == "bar";
            match (grouping, is_h) {
                ("stacked", true) => "stackedBarH",
                ("stacked", false) => "stackedBar",
                ("percentStacked", true) => "stackedBarHPct",
                ("percentStacked", false) => "stackedBarPct",
                (_, true) => "clusteredBarH",
                (_, false) => "clusteredBar",
            }
            .to_string()
        }
        "line" => match grouping {
            "stacked" => "stackedLine",
            "percentStacked" => "stackedLinePct",
            _ => "line",
        }
        .to_string(),
        "area" => match grouping {
            "stacked" => "stackedArea",
            "percentStacked" => "stackedAreaPct",
            _ => "area",
        }
        .to_string(),
        other => other.to_string(),
    }
}

/// Find a direct child of `parent` whose local name is `name`.
fn child<'a, 'i>(parent: Node<'a, 'i>, name: &str) -> Option<Node<'a, 'i>> {
    parent
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == name)
}

/// Theme-aware color resolution for chart text-color helpers.
///
/// pptx and xlsx store their theme palettes in different shapes
/// (`HashMap<String, String>` vs. `&[String]`) and apply DrawingML
/// transforms with different `tint` formulas (Word-literal vs. linear
/// sRGB lerp), so each crate keeps its own resolver. The shared chart
/// helpers take a `&dyn ColorResolver` instead of either concrete type so
/// fields like `<c:dLbls><c:txPr>...<a:solidFill>` can be extracted once.
pub trait ColorResolver {
    /// Resolve an `<a:solidFill>` node to a hex string (no leading `#`),
    /// or `None` when the contained color child can't be mapped to a
    /// concrete RGB value (for example a `<a:schemeClr val="phClr"/>`
    /// that the implementation chooses not to substitute).
    ///
    /// The node passed in is the `<a:solidFill>` element itself; the
    /// implementation reads its direct children for the actual color
    /// (`<a:srgbClr>` / `<a:schemeClr>` / `<a:sysClr>` / `<a:prstClr>`)
    /// and applies the surrounding lumMod/lumOff/tint/shade transforms.
    fn resolve_solid_fill(&self, node: Node) -> Option<String>;
}

/// `<c:legend>` presence + `<c:legendPos val>` (ECMA-376 §21.2.2.10).
///
/// `(show_legend, legend_pos)`. When the chart omits `<c:legend>` Office
/// hides the legend even if a default position would otherwise apply, so
/// `show_legend = false` is the authoritative "no legend" signal.
pub fn extract_legend(root: Node) -> (bool, Option<String>) {
    // The legend can sit anywhere inside `<c:chart>` but in practice it's a
    // direct child of `<c:chart>`. Use descendants to be tolerant of either
    // structure — there's only one `<c:legend>` element per chart.
    let legend = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "legend");
    let show = legend.is_some();
    let pos = legend.and_then(|ln| {
        child(ln, "legendPos")
            .and_then(|p| p.attribute("val"))
            .map(|s| s.to_string())
    });
    (show, pos)
}

/// `<c:barChart><c:gapWidth val>` / `<c:overlap val>` (ECMA-376 §21.2.2.13,
/// §21.2.2.25). Returns `(gap%, overlap%)`. Defaults to (None, None) when
/// the file relies on Office's defaults (gap 150, overlap 0).
pub fn extract_bar_gap_overlap(root: Node) -> (Option<i32>, Option<i32>) {
    let gap = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "gapWidth")
        .and_then(|n| n.attribute("val").and_then(|v| v.parse::<i32>().ok()));
    let ov = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "overlap")
        .and_then(|n| n.attribute("val").and_then(|v| v.parse::<i32>().ok()));
    (gap, ov)
}

/// First `<c:dLbls><c:dLblPos val>` found anywhere in the chart (chart-level
/// or per-series). ECMA-376 §21.2.2.49.
pub fn extract_data_label_position(root: Node) -> Option<String> {
    root.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "dLbls")
        .find_map(|dlbls| {
            child(dlbls, "dLblPos")
                .and_then(|n| n.attribute("val"))
                .map(|s| s.to_string())
        })
}

/// First non-`General` `<c:dLbls><c:numFmt formatCode>` in the chart.
/// ECMA-376 §21.2.2.37.
pub fn extract_data_label_format_code(root: Node) -> Option<String> {
    root.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "dLbls")
        .find_map(|dlbls| {
            child(dlbls, "numFmt")
                .and_then(|n| n.attribute("formatCode"))
                .map(|s| s.to_string())
                .filter(|s| !s.is_empty() && s != "General")
        })
}

/// `<c:catAx|valAx><c:numFmt formatCode>` — the value-axis tick label
/// number format (ECMA-376 §21.2.2.21). Caller passes the already-located
/// `<c:catAx>` / `<c:valAx>` node.
pub fn extract_axis_format_code(axis_node: Node) -> Option<String> {
    child(axis_node, "numFmt")
        .and_then(|n| n.attribute("formatCode"))
        .map(|s| s.to_string())
        .filter(|s| !s.is_empty() && s != "General")
}

/// `<c:catAx|valAx><c:scaling>` — read explicit `<c:min val>` / `<c:max val>`.
/// Returns `(min, max)`; either can be `None` when the file leaves Excel to
/// pick the auto bound.
pub fn extract_axis_min_max(axis_node: Node) -> (Option<f64>, Option<f64>) {
    let Some(scaling) = child(axis_node, "scaling") else {
        return (None, None);
    };
    let mn = child(scaling, "min")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok());
    let mx = child(scaling, "max")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok());
    (mn, mx)
}

/// `<c:catAx|valAx><c:delete val="1"/>` — true when the axis (labels, ticks
/// and line) should be hidden. ECMA-376 §21.2.2.40.
pub fn axis_is_deleted(axis_node: Node) -> bool {
    child(axis_node, "delete")
        .and_then(|n| n.attribute("val"))
        .map(|v| v != "0" && !v.eq_ignore_ascii_case("false"))
        .unwrap_or(false)
}

/// `<c:catAx|valAx><c:majorTickMark val>` / `<c:minorTickMark val>`. Values
/// are the ECMA-376 §21.2.3.48 ST_TickMark enum: `none` | `out` | `in` |
/// `cross`. Returns the raw string (None when the element is absent).
pub fn extract_axis_tick_mark(axis_node: Node, name: &str) -> Option<String> {
    child(axis_node, name)
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string())
}

/// Like [`extract_axis_tick_mark`] but applies the schema default `"out"` when
/// the element is absent (CT_TickMark `val` defaults to `out` — ECMA-376
/// §21.2.3.48 ST_TickMark). Keeps pptx/xlsx in agreement: the xlsx renderer
/// already defaults to `"out"`; the legacy pptx `"cross"` default was a bug
/// (it drew crossing ticks on charts that omit `<c:majorTickMark>`).
pub fn extract_axis_tick_mark_or_default(axis_node: Node, name: &str) -> String {
    extract_axis_tick_mark(axis_node, name).unwrap_or_else(|| "out".to_string())
}

/// First `<a:defRPr@sz>` or `<a:rPr@sz>` found inside the axis's `<c:txPr>`.
/// Sizes are OOXML hundredths of a point (e.g. 1200 = 12 pt).
pub fn extract_axis_tick_label_size(axis_node: Node) -> Option<i32> {
    let txpr = child(axis_node, "txPr")?;
    txpr.descendants().find_map(|n| {
        if !n.is_element() {
            return None;
        }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" {
            return None;
        }
        n.attribute("sz").and_then(|v| v.parse::<i32>().ok())
    })
}

/// First `<a:defRPr@b>` / `<a:rPr@b>` bold flag inside the axis's `<c:txPr>`
/// — the tick-label bold flag (ECMA-376 §21.2.2.17). `None` when unspecified.
pub fn extract_axis_tick_label_bold(axis_node: Node) -> Option<bool> {
    let txpr = child(axis_node, "txPr")?;
    txpr.descendants().find_map(|n| {
        if !n.is_element() {
            return None;
        }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" {
            return None;
        }
        n.attribute("b")
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
    })
}

/// Plain text of `node`'s direct-child `<c:title>` (ECMA-376 §21.2.2.6
/// `CT_Title`). Works for the `<c:chart>` element (chart title) or a
/// `<c:catAx>` / `<c:valAx>` (axis title). Walks `<a:t>` (rich text runs) and
/// `<c:v>` (string-ref cache) descendants and concatenates their text.
/// Returns `None` when there is no `<c:title>` child or it carries no text.
pub fn extract_chart_title_text(node: Node) -> Option<String> {
    let title = child(node, "title")?;
    let mut text = String::new();
    for d in title.descendants().filter(|n| n.is_element()) {
        match d.tag_name().name() {
            "t" | "v" => {
                if let Some(t) = d.text() {
                    text.push_str(t);
                }
            }
            _ => {}
        }
    }
    if text.is_empty() {
        None
    } else {
        Some(text)
    }
}

/// First `<a:defRPr@sz>` / `<a:rPr@sz>` (hundredths of a point) inside `node`'s
/// direct-child `<c:title>`. `None` when absent.
pub fn extract_chart_title_size(node: Node) -> Option<i32> {
    let title = child(node, "title")?;
    title.descendants().find_map(|n| {
        if !n.is_element() {
            return None;
        }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" {
            return None;
        }
        n.attribute("sz").and_then(|v| v.parse::<i32>().ok())
    })
}

/// First `<a:defRPr@b>` / `<a:rPr@b>` bold flag inside `node`'s direct-child
/// `<c:title>`. `None` when not specified (renderer treats as not bold).
pub fn extract_chart_title_bold(node: Node) -> Option<bool> {
    let title = child(node, "title")?;
    title.descendants().find_map(|n| {
        if !n.is_element() {
            return None;
        }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" {
            return None;
        }
        n.attribute("b")
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
    })
}

/// First `<a:solidFill>/<a:srgbClr@val>` (hex without `#`) inside `node`'s
/// direct-child `<c:title>`. Only an `<a:srgbClr>` that is a direct child of a
/// `<a:solidFill>` is honored — this skips gradient stops and other non-fill
/// color nodes. `<a:schemeClr>` is left unresolved here (the theme palette is
/// not wired through to chart title/border parsing yet — a known limitation
/// shared by both parsers). `None` = renderer default.
pub fn extract_chart_title_srgb(node: Node) -> Option<String> {
    let title = child(node, "title")?;
    title.descendants().find_map(|n| {
        if !n.is_element() || n.tag_name().name() != "srgbClr" {
            return None;
        }
        // Skip srgbClr nodes that aren't inside a solidFill (e.g. a gradient stop).
        let parent_is_solid = n
            .parent()
            .map(|p| p.tag_name().name() == "solidFill")
            .unwrap_or(false);
        if !parent_is_solid {
            return None;
        }
        n.attribute("val").map(|s| s.to_string())
    })
}

/// Axis title text + run props from a `<c:catAx>` / `<c:valAx>` node. Reuses
/// the chart-title helpers (which scope to the node's direct-child `<c:title>`);
/// run props are resolved only when title text is present, so an axis with no
/// title yields all `None`. Returns `(text, size_hpt, bold, srgb_color)`.
pub fn extract_axis_title_with_props(
    axis_node: Node,
) -> (Option<String>, Option<i32>, Option<bool>, Option<String>) {
    match extract_chart_title_text(axis_node) {
        None => (None, None, None, None),
        Some(text) => (
            Some(text),
            extract_chart_title_size(axis_node),
            extract_chart_title_bold(axis_node),
            extract_chart_title_srgb(axis_node),
        ),
    }
}

/// Explicit chart-frame border from `<c:chartSpace><c:spPr><a:ln>` (ECMA-376
/// §21.2.2.5 / DrawingML §20.1.2.2.24). `chart_space_root` is the
/// `<c:chartSpace>` element. Returns `(srgb_color, width_emu)` under the locked
/// policy shared by both parsers: a border is drawn ONLY when the XML explicitly
/// declares a paintable line.
///
///  - no `<a:ln>` (or no `<c:spPr>`) → `(None, None)` — no default border;
///  - `<a:ln><a:noFill/>` → border explicitly off → color `None` (width still
///    reported when `@w` is present);
///  - `<a:ln><a:solidFill><a:srgbClr@val>` → `(Some(hex), width)`.
///
/// `@w` (EMU) is captured as `u32` regardless of the fill. `<a:schemeClr>` is
/// intentionally left unresolved here (theme not wired through to chart border
/// parsing yet).
/// `<c:date1904>` (ECMA-376 §21.2.2.38) as a direct child of `<c:chartSpace>`.
/// The element is a `CT_Boolean`: `val` defaults to `true` when the element is
/// present but the attribute is omitted, so `<c:date1904/>` alone means
/// date1904=true. `val="0"` / `"false"` disable it. Absent element ⇒ false (the
/// default 1900 date system, §18.17.4.1).
pub fn extract_chart_date1904(chart_space_root: Node) -> bool {
    match child(chart_space_root, "date1904") {
        Some(n) => match n.attribute("val") {
            None => true, // element present, val implied true
            Some(v) => v == "1" || v.eq_ignore_ascii_case("true"),
        },
        None => false,
    }
}

/// `<c:ser><c:smooth val>` (ECMA-376 §21.2.2.194) — line/area series smoothing
/// flag. `ser_node` is the `<c:ser>` element. Returns `Some(true/false)` when
/// the element is present (CT_Boolean: `val` implied true when omitted),
/// `None` when the series has no `<c:smooth>` (straight-polyline default). Shared
/// so the pptx and xlsx parsers honor the flag identically.
pub fn extract_series_smooth(ser_node: Node) -> Option<bool> {
    child(ser_node, "smooth").map(|n| match n.attribute("val") {
        None => true, // element present, val implied true
        Some(v) => v == "1" || v.eq_ignore_ascii_case("true"),
    })
}

pub fn extract_chart_space_border(chart_space_root: Node) -> (Option<String>, Option<u32>) {
    let Some(ln) = child(chart_space_root, "spPr").and_then(|sp| child(sp, "ln")) else {
        return (None, None);
    };
    let width = ln.attribute("w").and_then(|v| v.parse::<u32>().ok());
    // An explicit `<a:noFill/>` turns the border off → no color.
    if child(ln, "noFill").is_some() {
        return (None, width);
    }
    // Only an srgbClr inside a direct `<a:solidFill>` is honored.
    let color = child(ln, "solidFill")
        .and_then(|sf| child(sf, "srgbClr"))
        .and_then(|srgb| srgb.attribute("val"))
        .map(|s| s.to_string());
    (color, width)
}

/// First `<c:dLbls><c:txPr>` font size (hpt). Mirrors the per-series + chart
/// fallback chain: walk every `<c:dLbls>` in document order, returning the
/// first inner `<a:defRPr@sz>` / `<a:rPr@sz>` we find.
pub fn extract_data_label_font_size(root: Node) -> Option<i32> {
    root.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "dLbls")
        .find_map(|dl| {
            child(dl, "txPr").and_then(|tx| {
                tx.descendants().find_map(|n| {
                    if !n.is_element() {
                        return None;
                    }
                    let tag = n.tag_name().name();
                    if tag != "defRPr" && tag != "rPr" {
                        return None;
                    }
                    n.attribute("sz").and_then(|v| v.parse::<i32>().ok())
                })
            })
        })
}

/// First `<c:dLbls><c:txPr>...<a:solidFill>` resolved to a hex color.
///
/// Walks each `<c:dLbls>` (chart-level + per-series) in document order,
/// drills into its `<c:txPr>` and looks for the first descendant
/// `<a:solidFill>` whose color the resolver can map. Stops on the first
/// successful resolution — this matches the chart-then-series fallback
/// pattern Office writers actually emit (e.g. a top-level `<c:dLbls>`
/// declaring the label color globally and the `<c:ser><c:dLbls>` blocks
/// inheriting it).
///
/// Note we deliberately scope the search to inside `<c:txPr>` so a
/// sibling `<c:dLbls><c:spPr><a:solidFill>` (the label *background*
/// fill, distinct from the text color) can't shadow the answer.
pub fn extract_data_label_font_color(root: Node, resolver: &dyn ColorResolver) -> Option<String> {
    for dlbls in root
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "dLbls")
    {
        let Some(txpr) = child(dlbls, "txPr") else {
            continue;
        };
        for desc in txpr.descendants().filter(|n| n.is_element()) {
            if desc.tag_name().name() != "solidFill" {
                continue;
            }
            if let Some(c) = resolver.resolve_solid_fill(desc) {
                return Some(c);
            }
        }
    }
    None
}

/// `<c:catAx|valAx><c:txPr>` tick-label text color, resolved to a hex string
/// (no leading `#`). Walks the axis's `<c:txPr>` for the first descendant
/// `<a:solidFill>` the resolver can map — this is the `<a:defRPr><a:solidFill>`
/// that ECMA-376 §21.2.2.* / §21.1.2.2.* uses to color the axis tick labels
/// (e.g. PowerPoint's "category labels in gray"). Scoped to `<c:txPr>` so the
/// sibling `<c:spPr>` axis-line fill can't shadow the answer.
pub fn extract_axis_tick_label_color(
    axis_node: Node,
    resolver: &dyn ColorResolver,
) -> Option<String> {
    let txpr = child(axis_node, "txPr")?;
    for desc in txpr.descendants().filter(|n| n.is_element()) {
        if desc.tag_name().name() != "solidFill" {
            continue;
        }
        if let Some(c) = resolver.resolve_solid_fill(desc) {
            return Some(c);
        }
    }
    None
}

/// `<c:catAx|valAx><c:spPr><a:ln>` axis-line style (ECMA-376 §21.2.2.* line
/// properties via DrawingML §20.1.2.2.24). Returns `(color, width_emu, no_fill)`:
///
///  - `color`: resolved hex (no `#`) when the line carries a `<a:solidFill>`.
///  - `width_emu`: the `<a:ln w>` width in EMU when present.
///  - `no_fill`: true when the line is explicitly `<a:noFill>` — the axis still
///    shows its labels/ticks but the rule itself is suppressed.
///
/// When the axis has no `<c:spPr><a:ln>` at all the tuple is
/// `(None, None, false)` and the caller falls back to its default rule.
pub fn extract_axis_line_style(
    axis_node: Node,
    resolver: &dyn ColorResolver,
) -> (Option<String>, Option<u32>, bool) {
    let Some(sp_pr) = child(axis_node, "spPr") else {
        return (None, None, false);
    };
    let Some(ln) = child(sp_pr, "ln") else {
        return (None, None, false);
    };
    let width = ln.attribute("w").and_then(|v| v.parse::<u32>().ok());
    let no_fill = child(ln, "noFill").is_some();
    let color = child(ln, "solidFill").and_then(|sf| resolver.resolve_solid_fill(sf));
    (color, width, no_fill)
}

/// chartEx (`<cx:chartSpace>`) axis visibility. ChartEx encodes the
/// scale type via a `<cx:catScaling>` / `<cx:valScaling>` child rather
/// than separate `<c:catAx>` / `<c:valAx>` elements, so callers can't just
/// reuse `axis_is_deleted` — this helper walks `<cx:axis hidden="1">` and
/// pairs each one with its scaling kind.
///
/// Returns `(cat_hidden, val_hidden)`. Defaults to `(false, false)` when no
/// `<cx:axis>` declares `hidden`.
pub fn extract_chartex_axis_hidden(root: Node) -> (bool, bool) {
    let mut cat_hidden = false;
    let mut val_hidden = false;
    for ax in root
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "axis")
    {
        let hidden = ax.attribute("hidden").map(|v| v == "1").unwrap_or(false);
        if !hidden {
            continue;
        }
        let is_val = ax
            .children()
            .any(|c| c.is_element() && c.tag_name().name() == "valScaling");
        let is_cat = ax
            .children()
            .any(|c| c.is_element() && c.tag_name().name() == "catScaling");
        if is_val {
            val_hidden = true;
        }
        if is_cat {
            cat_hidden = true;
        }
    }
    (cat_hidden, val_hidden)
}

#[cfg(test)]
mod tests {
    use super::*;
    use roxmltree::Document;

    fn root_of(xml: &str) -> Document<'_> {
        Document::parse(xml).expect("parse fixture")
    }

    #[test]
    fn canonical_chart_type_bar_matrix() {
        // Mirrors the TS `canonicalChartType` bar branch (ST_BarDir "bar" =
        // horizontal). Every (grouping, dir) pair must map to the same string
        // the renderer dispatches on.
        assert_eq!(
            canonical_chart_type("bar", "col", "clustered"),
            "clusteredBar"
        );
        assert_eq!(
            canonical_chart_type("bar", "bar", "clustered"),
            "clusteredBarH"
        );
        assert_eq!(canonical_chart_type("bar", "col", "stacked"), "stackedBar");
        assert_eq!(canonical_chart_type("bar", "bar", "stacked"), "stackedBarH");
        assert_eq!(
            canonical_chart_type("bar", "col", "percentStacked"),
            "stackedBarPct"
        );
        assert_eq!(
            canonical_chart_type("bar", "bar", "percentStacked"),
            "stackedBarHPct"
        );
        // Unknown grouping → clustered fallback (matches the TS default arm).
        assert_eq!(
            canonical_chart_type("bar", "col", "standard"),
            "clusteredBar"
        );
    }

    #[test]
    fn canonical_chart_type_line_area_and_passthrough() {
        assert_eq!(canonical_chart_type("line", "col", "standard"), "line");
        assert_eq!(
            canonical_chart_type("line", "col", "stacked"),
            "stackedLine"
        );
        assert_eq!(
            canonical_chart_type("line", "col", "percentStacked"),
            "stackedLinePct"
        );
        assert_eq!(canonical_chart_type("area", "col", "standard"), "area");
        assert_eq!(
            canonical_chart_type("area", "col", "stacked"),
            "stackedArea"
        );
        assert_eq!(
            canonical_chart_type("area", "col", "percentStacked"),
            "stackedAreaPct"
        );
        // Families the renderer already names canonically pass through verbatim.
        for t in ["pie", "doughnut", "scatter", "bubble", "radar", "waterfall"] {
            assert_eq!(canonical_chart_type(t, "col", "clustered"), t);
        }
    }

    /// The wire contract: a `ChartModel` must serialize with the same camelCase
    /// keys the TS `ChartModel` declares, REQUIRED fields present even when
    /// `None`/`false`/empty, OPTIONAL fields dropped when unset. This is the
    /// Rust-side oracle that pins the emitted JSON shape.
    #[test]
    fn chart_model_serializes_canonical_shape() {
        let m = ChartModel {
            chart_type: "clusteredBar".to_string(),
            title: None,
            categories: vec!["A".to_string(), "B".to_string()],
            series: vec![ChartSeries {
                name: "S1".to_string(),
                color: Some("FF0000".to_string()),
                values: vec![Some(1.0), None, Some(3.0)],
                data_point_colors: None,
                data_label_colors: None,
                label_color: None,
                series_type: None,
                use_secondary_axis: None,
                categories: None,
                show_marker: None,
                val_format_code: None,
                marker_symbol: None,
                marker_size: None,
                marker_fill: None,
                marker_line: None,
                data_point_overrides: None,
                data_label_overrides: None,
                series_data_labels: None,
                err_bars: None,
                bubble_sizes: None,
                smooth: None,
            }],
            show_data_labels: false,
            val_min: None,
            val_max: None,
            cat_axis_title: None,
            val_axis_title: None,
            cat_axis_hidden: false,
            val_axis_hidden: false,
            cat_axis_line_hidden: false,
            val_axis_line_hidden: false,
            plot_area_bg: None,
            chart_bg: Some("FFFFFF".to_string()),
            show_legend: false,
            legend_pos: None,
            cat_axis_cross_between: "between".to_string(),
            val_axis_major_tick_mark: "out".to_string(),
            cat_axis_major_tick_mark: "out".to_string(),
            title_font_size_hpt: None,
            title_font_color: None,
            title_font_face: None,
            cat_axis_font_size_hpt: None,
            val_axis_font_size_hpt: None,
            data_label_font_size_hpt: None,
            subtotal_indices: vec![],
            val_axis_minor_tick_mark: None,
            cat_axis_minor_tick_mark: None,
            cat_axis_font_color: None,
            val_axis_font_color: None,
            legend_manual_layout: None,
            val_axis_format_code: None,
            bar_gap_width: None,
            bar_overlap: None,
            data_label_position: None,
            data_label_font_color: None,
            data_label_format_code: None,
            title_font_bold: None,
            cat_axis_font_bold: None,
            val_axis_font_bold: None,
            cat_axis_title_font_size_hpt: None,
            cat_axis_title_font_bold: None,
            cat_axis_title_font_color: None,
            val_axis_title_font_size_hpt: None,
            val_axis_title_font_bold: None,
            val_axis_title_font_color: None,
            chart_border_color: None,
            chart_border_width_emu: None,
            cat_axis_crosses: None,
            cat_axis_crosses_at: None,
            val_axis_crosses: None,
            val_axis_crosses_at: None,
            cat_axis_line_color: None,
            cat_axis_line_width_emu: None,
            val_axis_line_color: None,
            val_axis_line_width_emu: None,
            cat_axis_format_code: None,
            cat_axis_min: None,
            cat_axis_max: None,
            title_manual_layout: None,
            plot_area_manual_layout: None,
            scatter_style: None,
            radar_style: None,
            secondary_val_axis: None,
            date1904: false,
        };
        let v = serde_json::to_value(&m).unwrap();
        let obj = v.as_object().unwrap();
        // Required scalar keys present with camelCase names, even when None/false.
        assert_eq!(obj["chartType"], "clusteredBar");
        assert!(obj["title"].is_null());
        assert_eq!(obj["showDataLabels"], false);
        assert_eq!(obj["catAxisHidden"], false);
        assert_eq!(obj["catAxisCrossBetween"], "between");
        assert_eq!(obj["valAxisMajorTickMark"], "out");
        assert!(obj["plotAreaBg"].is_null());
        assert_eq!(obj["chartBg"], "FFFFFF");
        assert_eq!(obj["subtotalIndices"], serde_json::json!([]));
        // Optional unset keys dropped from the wire.
        assert!(!obj.contains_key("barGapWidth"));
        assert!(!obj.contains_key("secondaryValAxis"));
        assert!(!obj.contains_key("catAxisFontColor"));
        // date1904 is dropped from the wire when false (default 1900 system).
        assert!(!obj.contains_key("date1904"));
        // Series: required present, optional dropped; array null preserved.
        let s0 = &obj["series"][0];
        assert_eq!(s0["name"], "S1");
        assert_eq!(s0["color"], "FF0000");
        assert_eq!(s0["values"], serde_json::json!([1.0, null, 3.0]));
        assert!(!s0.as_object().unwrap().contains_key("showMarker"));
        // Round-trips back to an equal model (Deserialize parity).
        let back: ChartModel = serde_json::from_value(v).unwrap();
        assert_eq!(back, m);
    }

    #[test]
    fn legend_present_with_pos() {
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:legend><c:legendPos val="t"/></c:legend>
        </c:chart>"#;
        let d = root_of(xml);
        let (show, pos) = extract_legend(d.root_element());
        assert!(show);
        assert_eq!(pos.as_deref(), Some("t"));
    }

    #[test]
    fn legend_absent() {
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        let (show, pos) = extract_legend(d.root_element());
        assert!(!show);
        assert!(pos.is_none());
    }

    #[test]
    fn bar_gap_overlap_default_to_none() {
        let xml =
            r#"<c:barChart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        assert_eq!(extract_bar_gap_overlap(d.root_element()), (None, None));
    }

    #[test]
    fn bar_gap_overlap_explicit() {
        let xml = r#"<c:barChart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:gapWidth val="50"/>
            <c:overlap val="100"/>
        </c:barChart>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_bar_gap_overlap(d.root_element()),
            (Some(50), Some(100))
        );
    }

    #[test]
    fn data_label_position() {
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:plotArea><c:dLbls><c:dLblPos val="ctr"/></c:dLbls></c:plotArea>
        </c:chart>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_data_label_position(d.root_element()).as_deref(),
            Some("ctr")
        );
    }

    #[test]
    fn axis_delete_truthy_variants() {
        for (val, expect) in [
            ("1", true),
            ("0", false),
            ("true", true),
            ("false", false),
            ("True", true),
        ] {
            let xml = format!(
                r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:delete val="{val}"/>
            </c:valAx>"#
            );
            let d = root_of(&xml);
            assert_eq!(axis_is_deleted(d.root_element()), expect, "val={val}");
        }
    }

    #[test]
    fn axis_delete_default_false() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        assert!(!axis_is_deleted(d.root_element()));
    }

    #[test]
    fn axis_min_max() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:scaling><c:max val="2500"/><c:min val="0"/></c:scaling>
        </c:valAx>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_axis_min_max(d.root_element()),
            (Some(0.0), Some(2500.0))
        );
    }

    #[test]
    fn series_smooth_present_and_absent() {
        // No `<c:smooth>` → None (straight-polyline default).
        let none =
            root_of(r#"<c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#);
        assert_eq!(extract_series_smooth(none.root_element()), None);
        // `<c:smooth val="1"/>` → Some(true).
        let on = root_of(
            r#"<c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:smooth val="1"/></c:ser>"#,
        );
        assert_eq!(extract_series_smooth(on.root_element()), Some(true));
        // `<c:smooth val="0"/>` → Some(false) (explicit off).
        let off = root_of(
            r#"<c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:smooth val="0"/></c:ser>"#,
        );
        assert_eq!(extract_series_smooth(off.root_element()), Some(false));
        // Bare `<c:smooth/>` → Some(true) (CT_Boolean implied-true).
        let bare = root_of(
            r#"<c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:smooth/></c:ser>"#,
        );
        assert_eq!(extract_series_smooth(bare.root_element()), Some(true));
    }

    #[test]
    fn axis_format_code_skips_general() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:numFmt formatCode="General"/>
        </c:valAx>"#;
        let d = root_of(xml);
        assert!(extract_axis_format_code(d.root_element()).is_none());
    }

    #[test]
    fn axis_format_code_passes_through() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:numFmt formatCode="0.0%"/>
        </c:valAx>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_axis_format_code(d.root_element()).as_deref(),
            Some("0.0%")
        );
    }

    /// Test resolver: returns the schemeClr@val verbatim, or the srgbClr@val
    /// uppercased. Just enough to drive `extract_data_label_font_color`.
    struct StubResolver;
    impl ColorResolver for StubResolver {
        fn resolve_solid_fill(&self, node: Node) -> Option<String> {
            for c in node.children().filter(|n| n.is_element()) {
                match c.tag_name().name() {
                    "srgbClr" => return c.attribute("val").map(|v| v.to_uppercase()),
                    "schemeClr" => return c.attribute("val").map(|v| v.to_string()),
                    _ => {}
                }
            }
            None
        }
    }

    #[test]
    fn data_label_font_color_resolves_via_resolver() {
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:plotArea>
                <c:dLbls>
                    <c:txPr><a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr></a:r></a:p></c:txPr>
                </c:dLbls>
            </c:plotArea>
        </c:chart>"#;
        let d = root_of(xml);
        let got = extract_data_label_font_color(d.root_element(), &StubResolver);
        assert_eq!(got.as_deref(), Some("bg1"));
    }

    #[test]
    fn data_label_font_color_skips_label_background_fill() {
        // `<c:spPr><a:solidFill>` (label background) must not be picked up;
        // only the text fill inside `<c:txPr>` counts.
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:plotArea>
                <c:dLbls>
                    <c:spPr><a:solidFill><a:srgbClr val="aabbcc"/></a:solidFill></c:spPr>
                </c:dLbls>
            </c:plotArea>
        </c:chart>"#;
        let d = root_of(xml);
        let got = extract_data_label_font_color(d.root_element(), &StubResolver);
        assert!(
            got.is_none(),
            "spPr fill must not leak into the font color: got {got:?}"
        );
    }

    #[test]
    fn data_label_font_color_first_dlbls_wins() {
        // Mimics Office writers that put a chart-level dLbls block AND
        // per-series ones — the first txPr resolution wins.
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:plotArea>
                <c:dLbls>
                    <c:txPr><a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="ffffff"/></a:solidFill></a:rPr></a:r></a:p></c:txPr>
                </c:dLbls>
                <c:barChart>
                    <c:ser><c:dLbls>
                        <c:txPr><a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:rPr></a:r></a:p></c:txPr>
                    </c:dLbls></c:ser>
                </c:barChart>
            </c:plotArea>
        </c:chart>"#;
        let d = root_of(xml);
        let got = extract_data_label_font_color(d.root_element(), &StubResolver);
        assert_eq!(got.as_deref(), Some("FFFFFF"));
    }

    #[test]
    fn axis_tick_label_color_from_txpr() {
        let xml = r#"<c:catAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:spPr><a:ln><a:solidFill><a:srgbClr val="d9d9d9"/></a:solidFill></a:ln></c:spPr>
            <c:txPr><a:p><a:pPr><a:defRPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
        </c:catAx>"#;
        let d = root_of(xml);
        // The txPr text fill (bg1) is returned — the spPr line fill must not leak.
        let got = extract_axis_tick_label_color(d.root_element(), &StubResolver);
        assert_eq!(got.as_deref(), Some("bg1"));
    }

    #[test]
    fn axis_tick_label_color_absent() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        assert!(extract_axis_tick_label_color(d.root_element(), &StubResolver).is_none());
    }

    #[test]
    fn axis_line_style_solid_with_width() {
        let xml = r#"<c:catAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:spPr><a:noFill/><a:ln w="9525"><a:solidFill><a:srgbClr val="d9d9d9"/></a:solidFill></a:ln></c:spPr>
        </c:catAx>"#;
        let d = root_of(xml);
        let (color, width, no_fill) = extract_axis_line_style(d.root_element(), &StubResolver);
        assert_eq!(color.as_deref(), Some("D9D9D9"));
        assert_eq!(width, Some(9525));
        assert!(!no_fill);
    }

    #[test]
    fn axis_line_style_nofill_line() {
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:spPr><a:ln w="9525"><a:noFill/></a:ln></c:spPr>
        </c:valAx>"#;
        let d = root_of(xml);
        let (color, width, no_fill) = extract_axis_line_style(d.root_element(), &StubResolver);
        assert!(color.is_none());
        assert_eq!(width, Some(9525));
        assert!(no_fill);
    }

    #[test]
    fn axis_line_style_absent() {
        let xml = r#"<c:catAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_axis_line_style(d.root_element(), &StubResolver),
            (None, None, false)
        );
    }

    #[test]
    fn chartex_axis_hidden_value_only() {
        let xml = r#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
            <cx:axis id="0"><cx:catScaling/></cx:axis>
            <cx:axis id="1" hidden="1"><cx:valScaling/></cx:axis>
        </cx:chartSpace>"#;
        let d = root_of(xml);
        assert_eq!(extract_chartex_axis_hidden(d.root_element()), (false, true));
    }

    #[test]
    fn chart_title_text_size_bold_srgb() {
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:title><c:tx><c:rich>
                <a:p><a:pPr><a:defRPr sz="1400" b="1"><a:solidFill><a:srgbClr val="1B4332"/></a:solidFill></a:defRPr></a:pPr>
                <a:r><a:t>Carbon &amp; Growth</a:t></a:r></a:p>
            </c:rich></c:tx></c:title>
        </c:chart>"#;
        let d = root_of(xml);
        let root = d.root_element();
        assert_eq!(
            extract_chart_title_text(root).as_deref(),
            Some("Carbon & Growth")
        );
        assert_eq!(extract_chart_title_size(root), Some(1400));
        assert_eq!(extract_chart_title_bold(root), Some(true));
        assert_eq!(extract_chart_title_srgb(root).as_deref(), Some("1B4332"));
    }

    #[test]
    fn chart_title_text_from_strref_cache() {
        // Title sourced from a strRef cache (`<c:v>`) rather than rich runs.
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:title><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Sales</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title>
        </c:chart>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_chart_title_text(d.root_element()).as_deref(),
            Some("Sales")
        );
    }

    #[test]
    fn chart_title_srgb_skips_non_solidfill_srgb() {
        // An `<a:srgbClr>` that is NOT a direct child of `<a:solidFill>` (here a
        // gradient stop) must be ignored.
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:title><c:tx><c:rich><a:p><a:r><a:rPr>
                <a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="ABCDEF"/></a:gs></a:gsLst></a:gradFill>
            </a:rPr><a:t>T</a:t></a:r></a:p></c:rich></c:tx></c:title>
        </c:chart>"#;
        let d = root_of(xml);
        assert!(extract_chart_title_srgb(d.root_element()).is_none());
    }

    #[test]
    fn chart_title_helpers_absent() {
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        let root = d.root_element();
        assert!(extract_chart_title_text(root).is_none());
        assert!(extract_chart_title_size(root).is_none());
        assert!(extract_chart_title_bold(root).is_none());
        assert!(extract_chart_title_srgb(root).is_none());
    }

    #[test]
    fn axis_title_with_props_full() {
        let xml = r#"<c:catAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:axPos val="b"/>
            <c:title><c:tx><c:rich>
                <a:p><a:pPr><a:defRPr sz="1000" b="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:defRPr></a:pPr>
                <a:r><a:t>Category Axis</a:t></a:r></a:p>
            </c:rich></c:tx></c:title>
        </c:catAx>"#;
        let d = root_of(xml);
        let (text, size, bold, color) = extract_axis_title_with_props(d.root_element());
        assert_eq!(text.as_deref(), Some("Category Axis"));
        assert_eq!(size, Some(1000));
        assert_eq!(bold, Some(true));
        assert_eq!(color.as_deref(), Some("FF0000"));
    }

    #[test]
    fn axis_title_with_props_text_absent_all_none() {
        // Axis with no `<c:title>` → text None gates the props to None even
        // though run props could in theory be read elsewhere.
        let xml = r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:axPos val="l"/>
        </c:valAx>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_axis_title_with_props(d.root_element()),
            (None, None, None, None)
        );
    }

    #[test]
    fn axis_tick_label_bold_variants() {
        for (b, expect) in [("1", Some(true)), ("0", Some(false)), ("true", Some(true))] {
            let xml = format!(
                r#"<c:catAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                    <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr b="{b}"/></a:pPr><a:endParaRPr/></a:p></c:txPr>
                </c:catAx>"#
            );
            let d = root_of(&xml);
            assert_eq!(
                extract_axis_tick_label_bold(d.root_element()),
                expect,
                "b={b}"
            );
        }
    }

    #[test]
    fn axis_tick_label_bold_absent() {
        let xml = r#"<c:catAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        assert!(extract_axis_tick_label_bold(d.root_element()).is_none());
    }

    #[test]
    fn chart_space_border_solid() {
        let xml = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="1B4332"/></a:solidFill></a:ln></c:spPr>
        </c:chartSpace>"#;
        let d = root_of(xml);
        assert_eq!(
            extract_chart_space_border(d.root_element()),
            (Some("1B4332".to_string()), Some(19050))
        );
    }

    #[test]
    fn chart_space_border_nofill_color_none_width_kept() {
        let xml = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:spPr><a:ln w="12700"><a:noFill/></a:ln></c:spPr>
        </c:chartSpace>"#;
        let d = root_of(xml);
        // noFill turns the border off → color None, but @w is still reported.
        assert_eq!(
            extract_chart_space_border(d.root_element()),
            (None, Some(12700))
        );
    }

    #[test]
    fn chart_space_border_absent() {
        let xml =
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let d = root_of(xml);
        assert_eq!(extract_chart_space_border(d.root_element()), (None, None));
    }

    #[test]
    fn chart_date1904_variants() {
        // §21.2.2.38: CT_Boolean. Element present + val omitted ⇒ true.
        let ns = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        let bare = format!(r#"<c:chartSpace xmlns:c="{ns}"><c:date1904/></c:chartSpace>"#);
        assert!(extract_chart_date1904(root_of(&bare).root_element()));

        let one = format!(r#"<c:chartSpace xmlns:c="{ns}"><c:date1904 val="1"/></c:chartSpace>"#);
        assert!(extract_chart_date1904(root_of(&one).root_element()));

        let word =
            format!(r#"<c:chartSpace xmlns:c="{ns}"><c:date1904 val="true"/></c:chartSpace>"#);
        assert!(extract_chart_date1904(root_of(&word).root_element()));

        let zero = format!(r#"<c:chartSpace xmlns:c="{ns}"><c:date1904 val="0"/></c:chartSpace>"#);
        assert!(!extract_chart_date1904(root_of(&zero).root_element()));

        // Word form of the falsey value: `val="false"` also disables the 1904
        // system (CT_Boolean accepts both "0" and "false").
        let false_word =
            format!(r#"<c:chartSpace xmlns:c="{ns}"><c:date1904 val="false"/></c:chartSpace>"#);
        assert!(!extract_chart_date1904(root_of(&false_word).root_element()));

        // Absent element ⇒ false (default 1900 system).
        let absent = format!(r#"<c:chartSpace xmlns:c="{ns}"/>"#);
        assert!(!extract_chart_date1904(root_of(&absent).root_element()));
    }
}
