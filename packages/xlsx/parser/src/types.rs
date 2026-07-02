use serde::Serialize;
use std::collections::HashMap;

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Workbook {
    pub sheets: Vec<SheetMeta>,
}

/// Sheet visibility (`<sheet state>`, ECMA-376 §18.2.19 `ST_SheetState`).
/// `Hidden` = hidden but user-unhideable via the UI; `VeryHidden` = revealable
/// only programmatically. Default is `Visible`.
#[derive(Debug, Serialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub enum SheetVisibility {
    Visible,
    Hidden,
    VeryHidden,
}

impl SheetVisibility {
    /// For `skip_serializing_if`: omit the default (`Visible`) so existing
    /// workbook JSON snapshots are unchanged.
    pub fn is_visible(&self) -> bool {
        matches!(self, SheetVisibility::Visible)
    }
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetMeta {
    pub name: String,
    pub sheet_id: u32,
    pub r_id: String,
    /// Sheet tab color (`<sheetPr><tabColor>`, ECMA-376 §18.3.1.93) resolved
    /// to `#RRGGBB`. Surfaced at workbook-list time so the viewer can paint
    /// every tab without eagerly parsing each worksheet. `None` when the
    /// sheet declares no tab color.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tab_color: Option<String>,
    /// Sheet visibility (`<sheet state>`, ECMA-376 §18.2.19). Omitted from JSON
    /// when `Visible` (the default) so existing snapshots are unchanged.
    #[serde(skip_serializing_if = "SheetVisibility::is_visible")]
    pub visibility: SheetVisibility,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MergeCell {
    pub top: u32,
    pub left: u32,
    pub bottom: u32,
    pub right: u32,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Worksheet {
    pub name: String,
    pub rows: Vec<Row>,
    pub col_widths: HashMap<u32, f64>,
    pub row_heights: HashMap<u32, f64>,
    pub default_col_width: f64,
    pub default_row_height: f64,
    pub merge_cells: Vec<MergeCell>,
    pub freeze_rows: u32,
    pub freeze_cols: u32,
    pub conditional_formats: Vec<ConditionalFormat>,
    pub images: Vec<ImageAnchor>,
    pub charts: Vec<ChartAnchor>,
    /// Grouped shapes from `<xdr:grpSp>` inside a twoCellAnchor (ECMA-376
    /// §20.5.2.17). Each anchor flattens its shape tree to a list of leaf
    /// shapes with normalized geometry for the renderer.
    pub shape_groups: Vec<ShapeAnchor>,
    /// Whether to display zero values in cells (ECMA-376 §18.3.1.94)
    pub show_zeros: bool,
    /// Whether to draw default grid lines on this sheet. Mirrors the "View →
    /// Gridlines" checkbox in Excel; parsed from `<sheetView showGridLines>`
    /// (ECMA-376 §18.3.1.83). Defaults to true.
    pub show_gridlines: bool,
    /// Whether the sheet grid is laid out right-to-left, mirroring the entire
    /// grid so column A sits on the right. Parsed from `<sheetView rightToLeft>`
    /// (ECMA-376 §18.3.1.87). Defaults to false (left-to-right).
    pub right_to_left: bool,
    /// Tab color for the sheet tab (ECMA-376 §18.3.1.79)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tab_color: Option<String>,
    /// AutoFilter range (ECMA-376 §18.3.1.2)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub auto_filter: Option<CellRange>,
    /// Hyperlinks in this worksheet (ECMA-376 §18.3.1.47)
    pub hyperlinks: Vec<Hyperlink>,
    /// Cell refs (A1-style) that have an associated <comment> in xl/commentsN.xml.
    /// Excel shows a small red triangle in the top-right corner of each.
    pub comment_refs: Vec<String>,
    /// Full-fidelity comment bodies (text + author) for each `<comment>` in
    /// xl/commentsN.xml — agents that want to read the comment content (not
    /// just the cell that has one) consume this. Empty when the sheet has no
    /// comments file. Keep in sync with `comment_refs` (one entry per ref).
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub comments: Vec<XlsxComment>,
    /// `<dataValidations>` rules (ECMA-376 §18.3.1.32). Each rule covers one
    /// or more cell ranges and constrains permitted input ("list", "decimal",
    /// "date", "textLength", "custom", …). Empty when the sheet declares
    /// none. Renderer ignores this for now — exposed for tooling.
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub data_validations: Vec<DataValidation>,
    /// Defined names in scope for this sheet. Includes workbook-global names and
    /// any names whose `localSheetId` matches this sheet's position in the
    /// workbook. Used by conditional-formatting `expression` rules that call
    /// named ranges like `task_start`, `today`, etc. (ECMA-376 §18.2.5).
    pub defined_names: Vec<DefinedName>,
    /// Excel Tables defined for this sheet (ECMA-376 §18.5). Rendered with a
    /// built-in table style (bold header, banded rows, etc.) on top of the
    /// cells' own styles.
    pub tables: Vec<TableInfo>,
    /// Slicers anchored to the sheet's drawing (Office 2010+ extension —
    /// `http://schemas.microsoft.com/office/drawing/2010/slicer` inside
    /// `<mc:AlternateContent>/<mc:Choice>`). Each slicer resolves its cache
    /// and referenced pivotCacheDefinition so the renderer can draw a static
    /// button list with the saved selection state.
    pub slicers: Vec<SlicerAnchor>,
    /// Sparkline groups defined in the worksheet's `<extLst>` (Office 2010
    /// extension `http://schemas.microsoft.com/office/spreadsheetml/2009/9/main`,
    /// element `<x14:sparklineGroup>`, ECMA-376 §18.2 / Part 4).
    pub sparkline_groups: Vec<SparklineGroup>,
    /// Family name of the workbook's Normal-style font, resolved from
    /// `<cellStyleXfs>[0].fontId` → `<fonts>[fontId].name.val`. Used by the
    /// renderer to compute the Max Digit Width (ECMA-376 §18.3.1.13) for the
    /// active sheet's column widths. Denormalized onto every worksheet for
    /// renderer convenience; the value is workbook-wide.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default_font_family: Option<String>,
    /// Point size of the workbook's Normal-style font (`<fonts>[N].sz.val`).
    /// Used together with `default_font_family` to compute Max Digit Width.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default_font_size: Option<f64>,
}

/// Single sparkline group (`<x14:sparklineGroup>`). Holds the shared formatting
/// for every individual sparkline cell that belongs to the group.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SparklineGroup {
    /// `line` (default) | `column` | `stem` (win/loss).
    pub kind: SparklineType,
    /// Show a marker dot at every data point (line type only).
    pub markers: bool,
    /// Highlight high / low / first / last / negative points.
    pub high: bool,
    pub low: bool,
    pub first: bool,
    pub last: bool,
    pub negative: bool,
    /// Show the horizontal axis line when data crosses zero.
    pub display_x_axis: bool,
    /// `gap` (default) | `zero` | `span` — how empty cells in the source
    /// range are treated. We only honor `gap` (default) at render time.
    pub display_empty_cells_as: String,
    /// Per-axis-bound type: `individual` (default) / `group` / `custom`.
    pub min_axis_type: String,
    pub max_axis_type: String,
    /// Used when *AxisType=`custom`. f64::NAN otherwise.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub manual_min: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub manual_max: Option<f64>,
    /// Stroke weight in pt (line type). ECMA-376 default 0.75.
    pub line_weight: f64,
    /// Resolved RGB hex strings (e.g. `#4472C4`) — theme + tint flattened
    /// at parse time so the renderer never sees a theme index. `None` means
    /// the property was not specified and the renderer should fall back.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_series: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_negative: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_axis: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_markers: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_first: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_last: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_high: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_low: Option<String>,
    /// Individual sparklines (one per destination cell).
    pub sparklines: Vec<Sparkline>,
}

#[derive(Debug, Serialize, Clone, Copy)]
#[serde(rename_all = "camelCase")]
pub enum SparklineType {
    Line,
    Column,
    Stem,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Sparkline {
    /// 1-based row of the cell that displays this sparkline (`<xm:sqref>`).
    pub row: u32,
    /// 1-based column of the cell.
    pub col: u32,
    /// Numeric values resolved from the `<xm:f>` data range. `None` for
    /// empty / non-numeric / out-of-bounds cells.
    pub values: Vec<Option<f64>>,
}

/// Excel Table metadata (ECMA-376 §18.5 `<table>`). The renderer overlays a
/// built-in style on top of the cell styles inside `range`.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableInfo {
    /// Inclusive table area including the header row.
    pub range: CellRange,
    /// Built-in style name like "TableStyleLight18" (ECMA-376 §18.5.1.4).
    pub style_name: String,
    /// Number of header rows (default 1).
    pub header_row_count: u32,
    /// Number of totals rows at the bottom (default 0).
    pub totals_row_count: u32,
    /// `<tableStyleInfo showRowStripes>` — banded rows in the data region.
    pub show_row_stripes: bool,
    /// `<tableStyleInfo showColumnStripes>`.
    pub show_column_stripes: bool,
    /// `<tableStyleInfo showFirstColumn>`.
    pub show_first_column: bool,
    /// `<tableStyleInfo showLastColumn>`.
    pub show_last_column: bool,
    /// Accent color resolved from the built-in style name against this file's
    /// theme accents (e.g. `TableStyleLight18` → accent3 of theme1.xml). Used
    /// by the renderer to draw banding, header background, and rules.
    pub accent_color: String,
    /// `true` when `style_name` is defined in the file's `<tableStyles>` block,
    /// i.e. a *custom* style (ECMA-376 §18.5.1.2). The renderer must draw such
    /// tables strictly from their declared element dxfs and must NOT apply the
    /// accent-based approximation (banding, synthesized rules/header) that is
    /// reserved for built-in style names whose definitions are absent.
    #[serde(default)]
    pub is_custom: bool,
    /// Dxf index for the `wholeTable` element of a custom `<tableStyle>`
    /// (ECMA-376 §18.8.83). When set, its border/fill apply to every cell
    /// of the table as a base layer. Built-in style names use the renderer's
    /// accent-based fallback, not this field.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub whole_table_dxf: Option<u32>,
    /// Dxf index for the `headerRow` element of a custom `<tableStyle>`.
    /// Provides the header background fill, font color/weight, and any
    /// vertical separator borders shown between header cells.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_row_dxf: Option<u32>,
    /// Dxf index for the `totalRow` element (ECMA-376 §18.18.93).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub total_row_dxf: Option<u32>,
    /// Dxf index for the `firstColumn` element.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_column_dxf: Option<u32>,
    /// Dxf index for the `lastColumn` element.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub last_column_dxf: Option<u32>,
    /// Dxf index for the `firstRowStripe` (band1 horizontal) element — the odd
    /// banded-row stripe applied when `show_row_stripes` is set.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub band1_horizontal_dxf: Option<u32>,
    /// Dxf index for the `secondRowStripe` (band2 horizontal) element — the
    /// even banded-row stripe.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub band2_horizontal_dxf: Option<u32>,
    /// Per-column DXF references (ECMA-376 §18.5.1.3 `tableColumn`). Length
    /// matches the number of `<tableColumn>` children in the table XML, so
    /// `columns[c - range.left]` gives the DXFs for the cell column. The
    /// renderer can use these to apply column-level overlays for named-style
    /// tables that don't pre-bake column DXFs into the cell `xf`. For files
    /// where Excel pre-bakes the result into `xf` (the common case), reading
    /// the cell `xf` already reflects the column DXF and these fields are
    /// purely informational.
    pub columns: Vec<TableColumnInfo>,
}

/// Per-column DXF references inside a `<table>` element
/// (ECMA-376 §18.5.1.3 `tableColumn`).
#[derive(Debug, Serialize, Default, Clone)]
#[serde(rename_all = "camelCase")]
pub struct TableColumnInfo {
    /// `<tableColumn dataDxfId>` — applied to data-region cells in this column.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data_dxf_id: Option<u32>,
    /// `<tableColumn headerRowDxfId>` — applied to the header cell of this column.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_row_dxf_id: Option<u32>,
    /// `<tableColumn totalsRowDxfId>` — applied to the totals cell of this column.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub totals_row_dxf_id: Option<u32>,
}

/// One cell comment. Sourced from the classic notes file `xl/commentsN.xml`
/// (ECMA-376 §18.7) when present, otherwise from the Office-365 threaded
/// comments part `xl/threadedComments/threadedCommentN.xml` (MS-XLSX schema
/// `…/office/spreadsheetml/2018/threadedcomments`, `personId` resolved against
/// `xl/persons/person*.xml`). `text` is the joined plain text — every `<r><t>`
/// run for classic notes, every reply in the thread (newline-joined) for
/// threaded comments; rich-text formatting is dropped.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct XlsxComment {
    /// A1-style cell reference (`@ref` on the comment element).
    pub cell_ref: String,
    /// Resolved author name. For classic notes this is the `<authors>` entry
    /// indexed by `@authorId`; for threaded comments it is the `<person>`
    /// `displayName` matching `@personId`. `None` when unresolved / absent.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    /// Concatenated plain text (every run for classic; every threaded reply,
    /// newline-joined, for threaded).
    pub text: String,
}

/// One `<dataValidation>` rule (ECMA-376 §18.3.1.32). `type` covers the
/// constraint class ("list", "whole", "decimal", "date", "time", "textLength",
/// "custom"). `operator` qualifies it ("between", "notBetween", "equal",
/// "notEqual", "lessThan", …). `formula1` / `formula2` are the rule operands.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DataValidation {
    /// Affected cell ranges, written verbatim from `@sqref` (space-separated).
    pub sqref: String,
    /// Constraint class. None means the validator is the spec's default
    /// (`"none"`, treated as no constraint).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub validation_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub operator: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula1: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula2: Option<String>,
    #[serde(skip_serializing_if = "std::ops::Not::not", default)]
    pub allow_blank: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub prompt_title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub prompt: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_message: Option<String>,
}

/// Workbook- or sheet-scoped defined name (ECMA-376 §18.2.5 `definedName`).
/// `formula` is the raw formula text (typically a cell/range reference, e.g.
/// `ProjectSchedule!$E1`). Relative references inside are shifted relative to
/// A1 when substituted into a formula.
#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct DefinedName {
    pub name: String,
    pub formula: String,
}

// ─── Chart types ────────────────────────────────────────────────────────────

/// A data series inside a chart.
#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ChartSeries {
    /// Display name of the series.
    pub name: String,
    /// Chart type for this series ("bar"|"line"|"area"|"pie"|"radar"|"scatter").
    /// Allows mixed charts (e.g. bar + line sharing the same axes).
    pub series_type: String,
    /// Category labels (X-axis for most charts).
    pub categories: Vec<String>,
    /// Numeric values; `None` = missing data point.
    pub values: Vec<Option<f64>>,
    /// Explicit fill color hex. When the series has a `<c:spPr>` fill we resolve
    /// it at parse time; otherwise we fall back to `theme.accent[idx+1]` (the
    /// default Excel palette, keyed by `<c:idx>` per ECMA-376 §21.2.2.27) so the
    /// renderer doesn't need theme access. None only when neither applies.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// Whether to draw data-point markers on line/scatter series. Resolved at
    /// parse time from `<c:ser><c:marker><c:symbol val>` (ECMA-376 §21.2.2.32)
    /// falling back to the chart-type-level `<c:lineChart><c:marker val>`
    /// (§21.2.2.33). Absent markers default to hidden for line charts.
    pub show_marker: bool,
    /// `<c:val>/<c:numRef>/<c:numCache>/<c:formatCode>` — Excel number format
    /// code applied to the series' numeric values (e.g. `"¥"#,##0`). When the
    /// chart-level `<c:dLbls><c:numFmt>` is absent this governs how data
    /// labels are formatted per ECMA-376 §21.2.2.37.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_format_code: Option<String>,
    /// `<c:ser><c:dLbls><c:txPr>…<a:solidFill>` — series-level data-label text
    /// color (ECMA-376 §21.2.2.216). Resolved against the workbook theme.
    /// Stacked/clustered charts can color each series's labels independently
    /// (e.g. white on a dark bar, black on a light one), which a single
    /// chart-level color can't express. None = fall back to the chart-level
    /// `data_label_font_color` in the renderer.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub label_color: Option<String>,
    /// `<c:ser><c:order val>` — display order of this series (stacking
    /// ordering + legend ordering) per ECMA-376 §21.2.2.28. Lower `order`
    /// renders first/below; Excel's legend for horizontal bar charts reverses
    /// this so low-order series end up at the bottom of the plot.
    pub order: usize,
    /// `<c:marker><c:symbol val>` (ECMA-376 §21.2.2.32). One of
    /// "circle"|"square"|"diamond"|"triangle"|"x"|"plus"|"star"|"dot"|
    /// "dash"|"picture"|"none". Absent = renderer-default circle.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker_symbol: Option<String>,
    /// `<c:marker><c:size val>` (ECMA-376 §21.2.2.34) — marker side length
    /// in points. Default is renderer-defined (~5 pt).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker_size: Option<u32>,
    /// `<c:marker><c:spPr><a:solidFill>` resolved hex (no `#`). Marker fill
    /// independent of series stroke color.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker_fill: Option<String>,
    /// `<c:marker><c:spPr><a:ln><a:solidFill>` resolved hex (no `#`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker_line: Option<String>,
    /// Per-data-point overrides (`<c:dPt idx>`, ECMA-376 §21.2.2.39). Each
    /// entry overrides marker / fill for a single point in the series.
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub data_point_overrides: Vec<DataPointOverride>,
    /// Per-data-point custom labels (`<c:dLbl idx>` inside `<c:dLbls>`,
    /// ECMA-376 §21.2.2.45). Custom rich text — including
    /// `<a:fld type="CELLRANGE">` field references — is resolved to a plain
    /// string at parse time.
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub data_label_overrides: Vec<DataLabelOverride>,
    /// Series-level `<c:dLbls>` defaults (showVal / showSerName / etc.). When
    /// neither this nor `data_label_overrides` is present, series labels are
    /// suppressed even when the chart-level `data_label_position` is set.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub series_data_labels: Option<SeriesDataLabels>,
    /// Error bars (`<c:errBars>`, ECMA-376 §21.2.2.20). Up to two per series
    /// (one for X, one for Y). `cust` / `fixedVal` / `stdErr` / `stdDev` /
    /// `percentage` are all resolved to absolute plus/minus arrays at parse
    /// time so the renderer just draws lines.
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub err_bars: Vec<ErrBars>,
}

/// Per-point override pulled from `<c:dPt idx="N">` siblings of a series
/// (ECMA-376 §21.2.2.39). Any unset field falls back to the series-level
/// value.
#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct DataPointOverride {
    pub idx: u32,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker_symbol: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker_size: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker_fill: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker_line: Option<String>,
}

/// Custom data label for one point (`<c:dLbl idx="N">`).
#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct DataLabelOverride {
    pub idx: u32,
    /// Resolved text. Empty when the label is intentionally blank
    /// (e.g. an empty `<a:p>`); the renderer should still skip drawing
    /// for that idx because Excel deletes it.
    pub text: String,
    /// `<c:dLblPos val>` — "l"|"r"|"t"|"b"|"ctr"|"outEnd"|"bestFit". When
    /// `None` the series-level / chart-level position applies.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
    /// Resolved font color hex (no `#`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    /// Font size in OOXML hundredths of a point.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size_hpt: Option<i32>,
    /// `<a:defRPr b="1">` inside the per-idx rich text. None = inherit.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_bold: Option<bool>,
}

/// Series-level `<c:dLbls>` defaults applied to every point that doesn't
/// have its own `<c:dLbl>` override.
#[derive(Debug, Serialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct SeriesDataLabels {
    pub show_val: bool,
    pub show_cat_name: bool,
    pub show_ser_name: bool,
    pub show_percent: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub format_code: Option<String>,
    /// `<c:dLbls><c:txPr>...defRPr@b>` series-level bold default.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_bold: Option<bool>,
    /// `<c:dLbls><c:txPr>...defRPr@sz>` series-level font size in OOXML
    /// hundredths of a point (e.g. 1200 = 12 pt).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size_hpt: Option<i32>,
}

/// Error bars (`<c:errBars>`, ECMA-376 §21.2.2.20).
///
/// `errValType="cust"` — values come from `<c:plus>/<c:minus>/<c:numRef>`,
/// possibly cross-sheet. We resolve those at parse time into arrays. For
/// `fixedVal` / `stdErr` / `stdDev` / `percentage` the parser computes the
/// absolute per-point delta from the series values so the renderer just
/// draws line segments without needing the raw type info.
#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ErrBars {
    /// "x" | "y".
    pub dir: String,
    /// "plus" | "minus" | "both" — drives whether plus / minus is rendered.
    pub bar_type: String,
    /// Absolute positive deltas per point. `None` = no bar in that
    /// direction at that index.
    pub plus: Vec<Option<f64>>,
    /// Absolute negative deltas per point.
    pub minus: Vec<Option<f64>>,
    /// `<c:noEndCap val>` — when true the renderer skips the perpendicular
    /// cap line at the bar tip (ECMA-376 §21.2.2.21).
    pub no_end_cap: bool,
    /// Resolved stroke hex (no `#`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// Stroke width in EMU (`<a:ln w>`). 12700 EMU = 1 pt.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    /// `<a:prstDash val>` — "solid"|"dash"|"dot"|"dashDot"|...
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dash: Option<String>,
}

/// Parsed chart data extracted from `xl/charts/chartN.xml`.
#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ChartData {
    /// Primary chart type: "bar"|"line"|"area"|"pie"|"doughnut"|"radar"|"scatter"
    pub chart_type: String,
    /// Bar direction: "col" (vertical) | "row" (horizontal). Only relevant for bar charts.
    pub bar_dir: String,
    /// Grouping mode: "clustered"|"stacked"|"standard"|"percentStacked"
    pub grouping: String,
    /// Optional chart title.
    pub title: Option<String>,
    /// Shared category list (from first series that has categories).
    pub categories: Vec<String>,
    /// All series across all chart-type elements in plotArea.
    pub series: Vec<ChartSeries>,
    /// Whether data labels are enabled (c:dLbls with showVal or showPercent).
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub show_data_labels: bool,
    /// Category axis title (c:catAx/c:title).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_title: Option<String>,
    /// Value axis title (c:valAx/c:title).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_title: Option<String>,
    /// `c:valAx/c:numFmt@formatCode` — custom number format for the value axis
    /// tick labels (e.g. `"$"#,##0`). When unset, tick labels use a plain
    /// numeric format. `sourceLinked="1"` is treated as a non-override (i.e.
    /// the axis inherits the data's format code); we still capture it so the
    /// renderer can honor it.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_format_code: Option<String>,
    /// True when `<c:legend>` is present in the chart; false means no legend.
    pub show_legend: bool,
    /// `<c:legend><c:legendPos val>` — "r" (default) | "l" | "t" | "b" | "tr".
    /// None = default ("r").
    #[serde(skip_serializing_if = "Option::is_none")]
    pub legend_pos: Option<String>,
    /// Chart title font size in OOXML hundredths of a point (e.g. 1400 = 14pt).
    /// Taken from the first `defRPr@sz` or `rPr@sz` inside `c:title`. None =
    /// not specified; renderer falls back to a proportional default.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title_font_size_hpt: Option<i32>,
    /// Chart title font color as a hex string without '#'. Taken from the
    /// first `a:solidFill/a:srgbClr@val` inside `c:title`. None = default.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title_font_color: Option<String>,
    /// Chart title font family (ECMA-376 DrawingML §20.1.4.2.24 `a:latin@typeface`).
    /// Taken from the first `a:latin` element inside `c:title`. None = default.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title_font_face: Option<String>,
    /// Category axis tick-label font size in hundredths of a point
    /// (ECMA-376 §21.2.2.17 `c:txPr/a:defRPr@sz`). None = not specified.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_font_size_hpt: Option<i32>,
    /// Value axis tick-label font size in hundredths of a point.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_font_size_hpt: Option<i32>,
    /// ECMA-376 §21.2.2.40 — `<c:catAx><c:delete val="1"/>` hides the category
    /// axis (labels, ticks, and axis line).
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub cat_axis_hidden: bool,
    /// ECMA-376 §21.2.2.40 — `<c:valAx><c:delete val="1"/>` hides the value
    /// axis (labels, ticks, and axis line).
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub val_axis_hidden: bool,
    /// `<c:catAx><c:spPr><a:ln><a:noFill>` — hides just the category axis
    /// LINE while keeping labels/ticks visible. Distinct from `cat_axis_hidden`
    /// (full `<c:delete val="1"/>` removes labels too).
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub cat_axis_line_hidden: bool,
    /// `<c:valAx><c:spPr><a:ln><a:noFill>` — hides just the value axis LINE
    /// while keeping labels/ticks visible. Sample-1's "Carbon & Growth" line
    /// chart uses this to suppress the leftmost vertical rule.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub val_axis_line_hidden: bool,
    /// Outer `<c:chartSpace><c:spPr>` fill resolution (ECMA-376 §21.2.2.5).
    /// `Some(hex)` for `<a:solidFill>`, `None` for `<a:noFill>` or when spPr is
    /// absent *and* no explicit fill is declared. An absent `chart_bg` on the
    /// JSON side (serde skips None) tells the renderer "no outer frame" —
    /// i.e. transparent, so the underlying cell panel shows through.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub chart_bg: Option<String>,
    /// True when the parser saw a `<c:chartSpace><c:spPr>` element at all.
    /// Lets the renderer distinguish "spec explicitly said noFill" (present
    /// but `chart_bg` is None) from "no spPr — use the default opaque white"
    /// (absent).
    pub has_chart_sp_pr: bool,
    /// `<c:legend><c:layout><c:manualLayout>` absolute placement (ECMA-376
    /// §21.2.2.31). All four fractions are relative to the chart space.
    /// None = use the default side-based layout from `legend_pos`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub legend_manual_layout: Option<LegendManualLayout>,
    /// `<c:barChart><c:gapWidth val>` — space between category groups as a
    /// percentage of bar width (ECMA-376 §21.2.2.13). Default per spec is 150.
    /// None = renderer default.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bar_gap_width: Option<i32>,
    /// `<c:barChart><c:overlap val>` — overlap/gap between bars in a cluster as
    /// a signed percentage (ECMA-376 §21.2.2.25). Positive = bars overlap,
    /// negative = gap between bars, 0 = flush. Range [-100, 100].
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bar_overlap: Option<i32>,
    /// `<c:dLbls><c:dLblPos val>` (ECMA-376 §21.2.2.16). One of
    /// "ctr"|"inBase"|"inEnd"|"outEnd"|"l"|"r"|"t"|"b"|"bestFit" etc.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data_label_position: Option<String>,
    /// Hex (no `#`) resolved from `<c:dLbls><c:txPr>` text fill. Used for data
    /// label text color (e.g. "FFFFFF" when labels sit inside filled bars).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data_label_font_color: Option<String>,
    /// `<c:dLbls><c:numFmt@formatCode>` — optional chart-level override for
    /// data label number format (ECMA-376 §21.2.2.35). Takes precedence over
    /// per-series `val_format_code` at render time.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data_label_format_code: Option<String>,
    /// `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr@b>` — bold flag for
    /// the chart title. None = inherit (treat as not bold).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title_font_bold: Option<bool>,
    /// `<c:catAx><c:txPr>...defRPr@b>` — bold flag for X-axis tick labels.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_font_bold: Option<bool>,
    /// `<c:valAx><c:txPr>...defRPr@b>` — bold flag for Y-axis tick labels.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_font_bold: Option<bool>,
    /// `<c:catAx><c:majorTickMark val>` (ECMA-376 §21.2.2.49) — one of
    /// `none` / `out` / `in` / `cross`. Default `out`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_major_tick_mark: Option<String>,
    /// `<c:catAx><c:minorTickMark val>` — same vocabulary. Default `none`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_minor_tick_mark: Option<String>,
    /// `<c:valAx><c:majorTickMark val>` and `<c:minorTickMark val>`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_major_tick_mark: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_minor_tick_mark: Option<String>,
    /// `<c:catAx><c:spPr><a:ln>` resolved color (hex without `#`) and
    /// width in EMU. Renders the X-axis line itself; default light gray
    /// at 1 px when absent.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_line_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_line_width_emu: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_line_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_line_width_emu: Option<u32>,
    /// `<c:catAx><c:crosses val>` — where the X axis sits along the Y axis
    /// (`autoZero` = at value 0, `min` = at the data min, `max` = at the
    /// data max). Default `autoZero`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_crosses: Option<String>,
    /// `<c:catAx><c:crossesAt val>` — explicit numeric crossing point;
    /// takes precedence over `cat_axis_crosses`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_crosses_at: Option<f64>,
    /// `<c:valAx><c:crosses val>` and `<c:crossesAt val>` mirroring the
    /// catAx fields above.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_crosses: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_crosses_at: Option<f64>,
    /// `<c:catAx><c:numFmt@formatCode>` (or scatter's X-axis valAx). Used
    /// to format the bottom-axis tick labels, e.g. `"m/d/yyyy"`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_format_code: Option<String>,
    /// `<c:catAx><c:scaling><c:min/max>` — only meaningful for scatter
    /// charts (where the X axis is numeric). None = derive from data.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_min: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_max: Option<f64>,
    /// `<c:valAx><c:scaling><c:min/max>` — explicit Y-axis range. None =
    /// derive from data.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_min: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_max: Option<f64>,
    /// `<c:title><c:layout><c:manualLayout>` (ECMA-376 §21.2.2.27) absolute
    /// placement for the chart title. When present, overrides the renderer's
    /// auto-positioning; absent = current default layout.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title_manual_layout: Option<ManualLayout>,
    /// `<c:plotArea><c:layout><c:manualLayout>` absolute placement for the
    /// plot area, with `layoutTarget=inner` (default) describing the inner
    /// plot rect (no axes / labels) or `outer` describing the outer rect
    /// (axes + labels included).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub plot_area_manual_layout: Option<ManualLayout>,
    /// `<c:radarChart><c:radarStyle val>` — "standard" (line only),
    /// "marker" (line + markers), "filled" (closed polygon with area fill).
    /// Default per ECMA-376 §21.2.3.10 is "standard" — no area fill.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub radar_style: Option<String>,
    /// `<c:catAx><c:title>` run-property font size in hundredths of a point
    /// (first `a:defRPr@sz`/`a:rPr@sz` inside the axis title). None = not set;
    /// renderer falls back to its proportional default. Distinct from
    /// `cat_axis_font_size_hpt`, which is the tick-label size (`c:txPr`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_size: Option<i32>,
    /// `<c:catAx><c:title>` run-property bold flag. None = inherit (not bold).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_bold: Option<bool>,
    /// `<c:catAx><c:title>` run-property color (hex without `#`) from the first
    /// `a:solidFill/a:srgbClr@val`. schemeClr is not resolved (theme not wired
    /// to charts). None = renderer default.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_title_color: Option<String>,
    /// `<c:valAx><c:title>` run-property font size in hundredths of a point.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_title_size: Option<i32>,
    /// `<c:valAx><c:title>` run-property bold flag. None = inherit.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_title_bold: Option<bool>,
    /// `<c:valAx><c:title>` run-property color (hex without `#`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_title_color: Option<String>,
    /// `<c:chartSpace><c:spPr><a:ln><a:solidFill><a:srgbClr@val>` — explicit
    /// chart border color (hex without `#`). Only populated when the XML
    /// explicitly declares a paintable line; `<a:noFill/>` or an absent `<a:ln>`
    /// leaves this None (no default Excel-style border). schemeClr is not
    /// resolved (theme not wired to charts).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub chart_border_color: Option<String>,
    /// `<c:chartSpace><c:spPr><a:ln@w>` — explicit chart border width in EMU.
    /// None = unset (renderer uses a 1px hairline when a color is present).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub chart_border_width_emu: Option<u32>,
}

/// Generic `<c:manualLayout>` block (used for title, plotArea, legend).
#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ManualLayout {
    /// "edge" (fraction is from top-left of chart) | "factor" (fraction
    /// from default position). ECMA-376 §21.2.2.32 ST_LayoutMode.
    pub x_mode: String,
    pub y_mode: String,
    /// "inner" (excludes axes) | "outer" (includes axes). Only meaningful
    /// for plotArea; harmless for title.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub layout_target: Option<String>,
    pub x: f64,
    pub y: f64,
    /// Width / height fractions. `None` = let the renderer auto-fit
    /// (corresponds to ECMA-376 only-(x,y) layouts where the size keeps its
    /// auto value).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub w: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub h: Option<f64>,
}

/// `<c:manualLayout>` coordinates for a legend (ECMA-376 §21.2.2.31). Fractions
/// are of the chart space's width/height. `xMode`/`yMode` select between "edge"
/// (fraction from top-left) and "factor" (fraction from the default position).
#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct LegendManualLayout {
    pub x_mode: String,
    pub y_mode: String,
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
}

/// A chart anchored to a rectangular range of cells (ECMA-376 §20.5 twoCellAnchor).
/// Offsets are EMU (914400 EMU = 1 inch, 9525 EMU = 1 px @ 96 DPI).
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartAnchor {
    pub from_col: u32,
    pub from_col_off: i64,
    pub from_row: u32,
    pub from_row_off: i64,
    pub to_col: u32,
    pub to_col_off: i64,
    pub to_row: u32,
    pub to_row_off: i64,
    pub chart: ChartData,
}

/// A grouped-shape anchor (ECMA-376 §20.5.2.17, `<xdr:grpSp>` inside a
/// `<xdr:twoCellAnchor>`). Leaf shape elements (`<xdr:sp>`) from any nesting
/// level are flattened into `shapes` with normalized coordinates so the
/// renderer only needs to scale to the anchor rect.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeAnchor {
    pub from_col: u32,
    pub from_col_off: i64,
    pub from_row: u32,
    pub from_row_off: i64,
    pub to_col: u32,
    pub to_col_off: i64,
    pub to_row: u32,
    pub to_row_off: i64,
    /// `twoCellAnchor@editAs` (ECMA-376 §20.5.2.33). With `"oneCell"` the
    /// renderer uses `native_ext_cx/cy` for the on-sheet size instead of the
    /// from/to-derived rect (Excel's "Move but don't size with cells").
    #[serde(skip_serializing_if = "Option::is_none")]
    pub edit_as: Option<String>,
    /// Group's `<xdr:grpSpPr><a:xfrm><a:ext cx cy>` (or `<xdr:spPr><a:xfrm>`
    /// for stand-alone sp/pic) in EMU. The saved on-sheet size, used as the
    /// authoritative extent when `editAs == "oneCell"`. 0 = unavailable.
    pub native_ext_cx: i64,
    pub native_ext_cy: i64,
    pub shapes: Vec<ShapeInfo>,
}

/// A leaf shape extracted from a grpSp/sp tree. Position/size are normalized
/// to [0,1] relative to the top-level grpSp extent (which itself maps to the
/// anchor rect).
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeInfo {
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
    /// Rotation in degrees (clockwise). DrawingML `a:xfrm/@rot` is in 60000ths
    /// of a degree; the parser converts to degrees here.
    pub rot: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stroke_color: Option<String>,
    /// Stroke width in EMU (914400 = 1 inch). 0 = no stroke.
    pub stroke_width: i64,
    pub geom: ShapeGeom,
    /// Text content from `<xdr:txBody>` (ECMA-376 §20.5.2.34). Present on
    /// shapes that carry visible text — typically text boxes (`txBox="1"`)
    /// but also any `<xdr:sp>` whose `<a:p>` runs render visibly. `None`
    /// when the shape has no text body or only contains an empty paragraph.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text: Option<ShapeText>,
}

/// Paragraph line spacing (`<a:pPr>/<a:lnSpc>`, ECMA-376 §21.1.2.2.5). Shared
/// with the pptx parser (and mirroring core's TS `SpaceLine`): a percentage of
/// the natural single line, or an absolute per-line height in points.
pub use ooxml_common::text::SpaceLine;

/// Text body inside a shape (`<xdr:txBody>`, ECMA-376 §20.1.2.2). Holds
/// the paragraphs plus body-level formatting (`<a:bodyPr>`).
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeText {
    /// `<a:bodyPr@anchor>` — vertical alignment of the text block within the
    /// shape rect. `t` (top, default), `ctr` (middle), `b` (bottom),
    /// `just`/`dist` (treated as top).
    pub anchor: String,
    /// `<a:bodyPr@wrap>` — `square` (default = wrap to shape width),
    /// `none` (no wrap).
    pub wrap: String,
    /// `<a:bodyPr>` autofit child (ECMA-376 §21.1.2.1.1-.3): `sp`
    /// (`spAutoFit`), `norm` (`normAutofit`), or `none` (`noAutofit`/absent).
    /// Always emitted (default `none`), mirroring the pptx `TextBody.autoFit`.
    pub auto_fit: String,
    /// `<a:normAutofit@fontScale>` — PowerPoint/Excel's stored font-shrink
    /// fraction (e.g. 0.625 for `fontScale="62500"`). `None` when unset.
    /// Modeled for parity; the xlsx renderer does not currently apply it.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_scale: Option<f64>,
    /// `<a:normAutofit@lnSpcReduction>` — stored line-spacing reduction fraction
    /// (e.g. 0.20 for `lnSpcReduction="20000"`). `None` when unset.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ln_spc_reduction: Option<f64>,
    pub paragraphs: Vec<ShapeParagraph>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeParagraph {
    /// `<a:pPr@algn>` — `l` (default), `ctr`, `r`, `just`, `dist`.
    pub align: String,
    /// `<a:pPr@rtl>` — whether the paragraph reads right-to-left (ECMA-376
    /// §21.1.2.2.7). Omitted from JSON when false so existing output stays
    /// byte-identical (additive, like the other Phase 0 direction flags).
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub rtl: bool,
    /// `<a:pPr@marL>` — left margin (EMU). ECMA-376 §21.1.2.2.7
    /// (`CT_TextParagraphProperties`). Direct attribute only (xlsx text boxes
    /// have no lstStyle/level cascade). `None` = unset. Omitted from JSON when
    /// `None` so existing output stays byte-identical (additive).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub mar_l: Option<i64>,
    /// `<a:pPr@marR>` — right margin (EMU). ECMA-376 §21.1.2.2.7. `None` = unset.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub mar_r: Option<i64>,
    /// `<a:pPr@indent>` — first-line indent (EMU; negative = hanging). ECMA-376
    /// §21.1.2.2.7. `None` = unset.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indent: Option<i64>,
    /// `<a:pPr>/<a:lnSpc>` line spacing (ECMA-376 §21.1.2.2.5). Direct-only.
    /// `None` = unset. Omitted from JSON when `None` so existing output stays
    /// byte-identical (additive).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub space_line: Option<SpaceLine>,
    pub runs: Vec<ShapeTextRun>,
}

/// A run within a shape paragraph. Tagged union (mirrors the pptx `TextRun`
/// shape) so a run is either styled text, a soft line break, or an OMML
/// equation. Excel stores "Insert > Equation" as OMML inside the shared
/// DrawingML `<xdr:txBody>` grammar (ECMA-376 §22.1), exactly like PowerPoint.
#[derive(Debug, Serialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum ShapeTextRun {
    /// The enum-level `rename_all = "camelCase"` (tag = "type") renames only the
    /// variant tags, not the fields, so the multi-word `font_face` needs a
    /// per-variant `rename_all` to serialize as the camelCase key the TS
    /// renderer reads (`renderer.ts`: `run.fontFace`). Without it the key lands
    /// as snake_case, the renderer reads `undefined`, and the shape text's font
    /// face is silently ignored (falls back to the default stack). Same root
    /// cause as the pptx fix in PR #489 / the xlsx ArcTo fix in PR #491.
    #[serde(rename_all = "camelCase")]
    Text {
        text: String,
        bold: bool,
        italic: bool,
        /// Font size in points (`<a:rPr@sz>` is in 100ths of a point; this
        /// field is already converted). 0 means "inherit from default" →
        /// renderer falls back to its own default.
        size: f64,
        #[serde(skip_serializing_if = "Option::is_none")]
        color: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        font_face: Option<String>,
        /// East-Asian typeface (`<a:ea@typeface>`, ECMA-376 §21.1.2.3.1). The
        /// common encoding for Japanese shape text sets Meiryo here while
        /// leaving `<a:latin>` default. The renderer floors the line box by
        /// this face's design line too. Additive — omitted from JSON when None
        /// so shapes without `<a:ea>` stay byte-identical.
        #[serde(skip_serializing_if = "Option::is_none")]
        font_face_ea: Option<String>,
        /// Complex-script typeface (`<a:cs@typeface>`, ECMA-376 §21.1.2.3.1).
        /// Additive — omitted from JSON when None.
        #[serde(skip_serializing_if = "Option::is_none")]
        font_face_cs: Option<String>,
    },
    /// Soft line break (`<a:br>`).
    Break,
    /// Inline (`display:false`) or block (`display:true`) OMML equation.
    ///
    /// Like `Text`, the multi-word `font_size` needs a per-variant
    /// `rename_all = "camelCase"` so it serializes as `fontSize` (the key the
    /// renderer reads at `renderer.ts`: `run.fontSize`). Without it the
    /// snake_case key is read as `undefined` and the equation always falls back
    /// to the inherited/default size instead of its explicit `rPr@sz`.
    #[serde(rename_all = "camelCase")]
    Math {
        nodes: Vec<ooxml_common::math::MathNode>,
        display: bool,
        /// Point size for the equation, when the run carries an explicit
        /// `rPr@sz`; otherwise the renderer inherits the surrounding size.
        #[serde(skip_serializing_if = "Option::is_none")]
        font_size: Option<f64>,
        #[serde(skip_serializing_if = "Option::is_none")]
        color: Option<String>,
    },
}

#[derive(Debug, Serialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum ShapeGeom {
    /// Preset geometry (rect, ellipse, roundRect, triangle, etc.).
    /// ECMA-376 §20.1.9.18 `a:prstGeom/@prst`.
    ///
    /// `adj` carries the shape's adjust handles from `<a:avLst><a:gd>` in
    /// declaration order (index 0 = adj/adj1, 1 = adj2, …); `None` entries mean
    /// "use the preset's declared default". ECMA-376 §19.5.31.3 / §20.1.9.5.
    Preset {
        name: String,
        #[serde(skip_serializing_if = "Vec::is_empty")]
        adj: Vec<Option<f64>>,
    },
    /// Freeform path geometry (ECMA-376 §20.1.9.2 `a:custGeom`).
    Custom { paths: Vec<PathInfo> },
    /// Bitmap (or vector) image leaf inside a `<xdr:grpSp>` tree (ECMA-376
    /// §20.5.2.17). `image_path` is the zip path of the drawing's relationship
    /// target (png/jpg/gif/svg/…) — the blip's raster `r:embed` fallback, or the
    /// SVG itself when no raster is embedded — and `mime_type` its MIME via the
    /// shared `mime_from_ext`. `svg_image_path` carries the Microsoft svgBlip
    /// extension's vector original when present, so the renderer can prefer it
    /// and fall back to `image_path` on a decode failure. `None` otherwise. Its
    /// MIME is always `image/svg+xml` and is owned by the SVG decoder. The
    /// renderer fetches bytes lazily via `extract_image`.
    #[serde(rename_all = "camelCase")]
    Image {
        image_path: String,
        mime_type: String,
        #[serde(skip_serializing_if = "Option::is_none")]
        svg_image_path: Option<String>,
        /// ECMA-376 §20.1.8.55 `<a:srcRect>` source-image crop on the leaf pic,
        /// present only when cropped. Honored by the renderer like the top-level
        /// `ImageAnchor.src_rect` so a `oneCellAnchor` / `grpSp` leaf pic crops
        /// the same as a `twoCellAnchor` picture.
        #[serde(skip_serializing_if = "Option::is_none")]
        src_rect: Option<SrcRect>,
    },
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PathInfo {
    /// Path's own coordinate system width.
    pub w: f64,
    /// Path's own coordinate system height.
    pub h: f64,
    pub commands: Vec<PathCmd>,
}

#[derive(Debug, Serialize)]
#[serde(tag = "op", rename_all = "camelCase")]
pub enum PathCmd {
    MoveTo {
        x: f64,
        y: f64,
    },
    LineTo {
        x: f64,
        y: f64,
    },
    CubicBezTo {
        x1: f64,
        y1: f64,
        x2: f64,
        y2: f64,
        x3: f64,
        y3: f64,
    },
    QuadBezTo {
        x1: f64,
        y1: f64,
        x2: f64,
        y2: f64,
    },
    /// ECMA-376 §20.1.9.3 `a:arcTo`. `stAng`/`swAng` are in 60000ths of a
    /// degree. The start point is the current pen position; the ellipse
    /// center is derived so the pen lies on the ellipse at `stAng`.
    ///
    /// The enum-level `rename_all = "camelCase"` (tag = "op") renames only the
    /// variant tags, not the fields, so `st_ang`/`sw_ang` need a per-variant
    /// `rename_all` to serialize as the camelCase keys the TS renderer reads
    /// (`renderer.ts`: `cmd.stAng`/`cmd.swAng`, then `/ 60000`). Without it the
    /// keys land as snake_case, the renderer reads `undefined` → `NaN`, and the
    /// arc fails to draw. Same root cause as the pptx fix in PR #489. The raw
    /// 60000ths convention is unchanged (the renderer does the division).
    #[serde(rename_all = "camelCase")]
    ArcTo {
        wr: f64,
        hr: f64,
        st_ang: f64,
        sw_ang: f64,
    },
    Close,
}

/// Slicer anchor — a button bank that filters a connected pivot table or
/// Excel Table. Office stores slicers in a 2010 extension namespace
/// (`sle:slicer`) wrapped in `<mc:AlternateContent>`, with the cache data in
/// `xl/slicerCaches/*.xml` and the underlying item list in the linked
/// pivotCache's `sharedItems`.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SlicerAnchor {
    pub from_col: u32,
    pub from_col_off: i64,
    pub from_row: u32,
    pub from_row_off: i64,
    pub to_col: u32,
    pub to_col_off: i64,
    pub to_row: u32,
    pub to_row_off: i64,
    /// Slicer header text. Typically the `caption` of the slicer definition
    /// (`xl/slicers/slicerN.xml`), falling back to the drawing `cNvPr` name.
    pub caption: String,
    /// One row per cache item in display order. Items flagged "selected" are
    /// the ones currently active in the filter; non-selected items are drawn
    /// with the ghost style. When the cache selection state is unavailable,
    /// all items are emitted as selected (Excel's default).
    pub items: Vec<SlicerItem>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SlicerItem {
    pub name: String,
    pub selected: bool,
}

/// An image anchored to a rectangular range of cells
/// (ECMA-376 §20.5, `<xdr:twoCellAnchor>`). Offsets are EMU (English
/// Metric Unit): 914400 EMU = 1 inch, and 9525 EMU = 1 pixel at 96 DPI.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ImageAnchor {
    pub from_col: u32,
    pub from_col_off: i64,
    pub from_row: u32,
    pub from_row_off: i64,
    pub to_col: u32,
    pub to_col_off: i64,
    pub to_row: u32,
    pub to_row_off: i64,
    /// `twoCellAnchor@editAs` (ECMA-376 §20.5.2.33). Possible values: `"twoCell"`
    /// (default), `"oneCell"`, `"absolute"`. With `"oneCell"`, Excel preserves
    /// the picture's native EMU size (below) when cells are resized; with
    /// `"twoCell"`, the from/to anchor rect IS the rendered size.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub edit_as: Option<String>,
    /// `<xdr:pic><xdr:spPr><a:xfrm><a:ext cx cy>` in EMU. The picture's saved
    /// size at insert/edit time. Used as the authoritative size when
    /// `editAs == "oneCell"`. 0 = absent / use from/to rect.
    pub native_ext_cx: i64,
    pub native_ext_cy: i64,
    /// Zip path of the blip inside the package (e.g. `xl/media/image1.png`).
    /// The blip's own `r:embed` raster fallback when an svgBlip extension is
    /// present; otherwise the only source. Falls back to the SVG part itself
    /// when the picture has no raster `r:embed` (an icon inserted as a pure
    /// SVG), so the element is always drawable. The renderer fetches the bytes
    /// lazily via `extract_image` rather than receiving an inlined base64 URL.
    pub image_path: String,
    /// MIME of the blip at `image_path`, derived from its extension via the
    /// shared `mime_from_ext` (e.g. `image/png`, or `image/svg+xml` for the
    /// SVG-only fallback).
    pub mime_type: String,
    /// Microsoft 2016 SVG extension (`<a:blip><a:extLst><a:ext
    /// uri="{96DAC541-…}"><asvg:svgBlip r:embed>`, MS-ODRAWXML): the zip path of
    /// the vector *original*, so the renderer can prefer it and fall back to
    /// `image_path` on a decode failure. `None` when the picture carries no
    /// svgBlip extension (the common case). Its MIME is always `image/svg+xml`
    /// and is owned by the SVG decoder.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub svg_image_path: Option<String>,
    /// ECMA-376 §20.1.8.55 `<a:srcRect>` source-image crop, present only when the
    /// picture is cropped. `None` (the common case) ⇒ the whole blip fills the
    /// anchor rect. When set, the renderer draws only the visible sub-rectangle.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub src_rect: Option<SrcRect>,
}

/// ECMA-376 §20.1.8.55 `<a:srcRect>` source-image crop, shared across the docx,
/// pptx and xlsx parsers (see `ooxml_common::blip`).
pub use ooxml_common::blip::SrcRect;

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CellRange {
    pub top: u32,
    pub left: u32,
    pub bottom: u32,
    pub right: u32,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ConditionalFormat {
    pub sqref: Vec<CellRange>,
    pub rules: Vec<CfRule>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase", tag = "type")]
pub enum CfRule {
    #[serde(rename_all = "camelCase")]
    CellIs {
        operator: String,
        formulas: Vec<String>,
        dxf_id: Option<u32>,
        priority: i32,
    },
    #[serde(rename_all = "camelCase")]
    Expression {
        formula: String,
        dxf_id: Option<u32>,
        priority: i32,
        stop_if_true: bool,
    },
    #[serde(rename_all = "camelCase")]
    ColorScale { stops: Vec<CfStop>, priority: i32 },
    #[serde(rename_all = "camelCase")]
    DataBar {
        color: String,
        min: CfValue,
        max: CfValue,
        priority: i32,
        gradient: bool,
    },
    #[serde(rename_all = "camelCase")]
    Top10 {
        top: bool,
        percent: bool,
        rank: u32,
        dxf_id: Option<u32>,
        priority: i32,
    },
    #[serde(rename_all = "camelCase")]
    AboveAverage {
        above_average: bool,
        /// ECMA-376 §18.3.1.10 `equalAverage`: include cells equal to the
        /// average in the highlighted set. Default false.
        #[serde(skip_serializing_if = "std::ops::Not::not")]
        equal_average: bool,
        /// ECMA-376 §18.3.1.10 `stdDev`: when present, the threshold becomes
        /// `mean ± stdDev · σ` (population standard deviation) instead of the
        /// plain mean. Absent (None) means a simple above/below-average rule.
        #[serde(skip_serializing_if = "Option::is_none")]
        std_dev: Option<u32>,
        dxf_id: Option<u32>,
        priority: i32,
    },
    #[serde(rename_all = "camelCase")]
    IconSet {
        icon_set: String,
        cfvos: Vec<CfValue>,
        reverse: bool,
        priority: i32,
        #[serde(skip_serializing_if = "Option::is_none")]
        custom_icons: Option<Vec<CfIcon>>,
    },
    #[serde(rename_all = "camelCase")]
    Other { kind: String, priority: i32 },
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CfIcon {
    pub icon_set: String,
    pub icon_id: u32,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CfStop {
    pub kind: String,
    pub value: Option<String>,
    pub color: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CfValue {
    pub kind: String,
    pub value: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Hyperlink {
    pub col: u32,
    pub row: u32,
    pub url: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Row {
    pub index: u32,
    pub height: Option<f64>,
    pub cells: Vec<Cell>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Cell {
    pub col: u32,
    pub row: u32,
    pub col_ref: String,
    pub value: CellValue,
    pub style_index: u32,
    /// Raw `<f>` formula text (ECMA-376 §18.3.1.40), when present. The
    /// renderer uses this to recompute volatile functions like TODAY() /
    /// NOW() at display time so the cached `<v>` (frozen when the file was
    /// last saved) doesn't show a stale date.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase", tag = "type")]
pub enum CellValue {
    #[default]
    Empty,
    Text {
        text: String,
        #[serde(skip_serializing_if = "Option::is_none")]
        runs: Option<Vec<Run>>,
    },
    Number {
        number: f64,
    },
    Bool {
        bool: bool,
    },
    Error {
        error: String,
    },
}

#[derive(Debug, Clone, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Run {
    pub text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font: Option<RunFont>,
}

#[derive(Debug, Clone, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct RunFont {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// ECMA-376 §18.4.13 ST_UnderlineValues — see Font.underline_style.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline_style: Option<String>,
    /// ECMA-376 §18.4.6 ST_VerticalAlignRun — "superscript" | "subscript" |
    /// "baseline". Absent leaves the run on the baseline.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<String>,
}

#[derive(Debug, Clone, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct SharedString {
    pub text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub runs: Option<Vec<Run>>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Styles {
    pub fonts: Vec<Font>,
    pub fills: Vec<Fill>,
    pub borders: Vec<Border>,
    pub cell_xfs: Vec<CellXf>,
    pub num_fmts: Vec<NumFmt>,
    pub dxfs: Vec<Dxf>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Dxf {
    pub font: Option<Font>,
    pub fill: Option<Fill>,
    pub border: Option<Border>,
    /// Number format override applied when the conditional-formatting rule
    /// matches. ECMA-376 §18.8.17 allows `<dxf>` to carry a `<numFmt>` that
    /// replaces the cell's own style numFmt (e.g. switching a calendar cell
    /// from `d` to `m"月"d"日"` on the first of each month).
    pub num_fmt: Option<NumFmt>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Font {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    pub size: f64,
    pub color: Option<String>,
    pub name: Option<String>,
    /// ECMA-376 §18.4.13 ST_UnderlineValues. Only emitted when not the default
    /// "single" — values: "double", "singleAccounting", "doubleAccounting".
    /// "none" sets `underline = false` and leaves this field absent.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline_style: Option<String>,
    /// ECMA-376 §18.4.6 ST_VerticalAlignRun on a cell-level <font> —
    /// "superscript" | "subscript". Absent leaves the run on the baseline.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<String>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Fill {
    pub pattern_type: String,
    pub fg_color: Option<String>,
    pub bg_color: Option<String>,
    /// When the fill element is a <gradientFill>, this carries the gradient
    /// stops + type + rotation. patternType stays "none" because xlsx does
    /// not mix gradient + pattern in the same fill.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub gradient: Option<GradientFillSpec>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct GradientFillSpec {
    /// "linear" (default) or "path". Linear uses `degree`; path uses top/bottom/left/right.
    pub gradient_type: String,
    /// Linear-gradient rotation in degrees (0 = left→right).
    pub degree: f64,
    /// Path-gradient bounding box (0..1 within the cell). Unused for linear.
    pub left: f64,
    pub right: f64,
    pub top: f64,
    pub bottom: f64,
    pub stops: Vec<GradientStopSpec>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct GradientStopSpec {
    pub position: f64,
    pub color: String,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Border {
    pub left: Option<BorderEdge>,
    pub right: Option<BorderEdge>,
    pub top: Option<BorderEdge>,
    pub bottom: Option<BorderEdge>,
    /// Diagonal line from bottom-left to top-right (ECMA-376 §18.8.4 diagonalUp)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub diagonal_up: Option<BorderEdge>,
    /// Diagonal line from top-left to bottom-right (ECMA-376 §18.8.4 diagonalDown)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub diagonal_down: Option<BorderEdge>,
    /// Inner horizontal rule between rows inside a region (ECMA-376 §18.8.40
    /// `tableStyleElement/border/horizontal`). Only set on table-style dxfs;
    /// ignored on cell-level borders.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub horizontal: Option<BorderEdge>,
    /// Inner vertical rule between columns inside a region (same ECMA-376
    /// section). Only set on table-style dxfs.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vertical: Option<BorderEdge>,
}

#[derive(Debug, Clone, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct BorderEdge {
    pub style: String,
    pub color: Option<String>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct CellXf {
    pub font_id: u32,
    pub fill_id: u32,
    pub border_id: u32,
    pub num_fmt_id: u32,
    pub align_h: Option<String>,
    pub align_v: Option<String>,
    pub wrap_text: bool,
    /// Text indentation level (each level ≈ 3 characters wide, ECMA-376 §18.8.44)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indent: Option<u32>,
    /// Text rotation in degrees: 0–90 = counter-clockwise, 91–180 = (value−90)° clockwise, 255 = stacked (ECMA-376 §18.8.44)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_rotation: Option<u32>,
    /// Shrink text to fit the cell width (ECMA-376 §18.8.44)
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub shrink_to_fit: bool,
    /// `<alignment readingOrder>` (ECMA-376 §18.8.1) — 0 = context (default),
    /// 1 = left-to-right, 2 = right-to-left. Drives canvas `direction`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub reading_order: Option<u32>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct NumFmt {
    pub num_fmt_id: u32,
    pub format_code: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ParsedWorkbook {
    pub workbook: Workbook,
    pub styles: Styles,
    pub shared_strings: Vec<SharedString>,
}
