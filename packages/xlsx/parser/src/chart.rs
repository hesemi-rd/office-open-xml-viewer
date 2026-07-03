use crate::types::*;
use crate::{parse_rels_map, resolve_fill_color, resolve_zip_path};
use ooxml_common::chart::{
    canonical_chart_type, ChartDataLabelOverride as CmDataLabelOverride,
    ChartDataPointOverride as CmDataPointOverride, ChartErrBars as CmErrBars,
    ChartManualLayout as CmManualLayout, ChartModel, ChartSeries as CmSeries,
    ChartSeriesDataLabels as CmSeriesDataLabels, LegendManualLayout as CmLegendManualLayout,
};
use ooxml_common::ns::{is_a_ns, is_c_ns, is_r_ns, is_xdr_ns};
use ooxml_common::zip::read_zip_string;
use std::collections::HashMap;

// ─── ChartData → shared ChartModel (was the TS `adaptChartData`) ─────────────
//
// The xlsx parser builds a `ChartData` (parser-native: `chartType` + `barDir` +
// `grouping`, always-present series `categories`/`showMarker`/`order`). Emitting
// the chart to JSON goes through this conversion, which is the Rust home of the
// former TS `adaptChartData` + `canonicalChartType` renderer helpers: it
// applies every default and conditional the adapter did so the wire object is
// already a canonical `ChartModel` and the TS side needs no adapter.

impl From<ManualLayout> for CmManualLayout {
    fn from(m: ManualLayout) -> Self {
        CmManualLayout {
            x_mode: m.x_mode,
            y_mode: m.y_mode,
            layout_target: m.layout_target,
            x: m.x,
            y: m.y,
            w: m.w,
            h: m.h,
        }
    }
}

impl From<LegendManualLayout> for CmLegendManualLayout {
    fn from(m: LegendManualLayout) -> Self {
        CmLegendManualLayout {
            x_mode: m.x_mode,
            y_mode: m.y_mode,
            x: m.x,
            y: m.y,
            w: m.w,
            h: m.h,
        }
    }
}

impl From<DataPointOverride> for CmDataPointOverride {
    fn from(o: DataPointOverride) -> Self {
        CmDataPointOverride {
            idx: o.idx,
            color: o.color,
            marker_symbol: o.marker_symbol,
            marker_size: o.marker_size.map(|v| v as f64),
            marker_fill: o.marker_fill,
            marker_line: o.marker_line,
        }
    }
}

impl From<DataLabelOverride> for CmDataLabelOverride {
    fn from(o: DataLabelOverride) -> Self {
        CmDataLabelOverride {
            idx: o.idx,
            text: o.text,
            position: o.position,
            font_color: o.font_color,
            font_size_hpt: o.font_size_hpt,
            font_bold: o.font_bold,
        }
    }
}

impl From<SeriesDataLabels> for CmSeriesDataLabels {
    fn from(d: SeriesDataLabels) -> Self {
        CmSeriesDataLabels {
            show_val: d.show_val,
            show_cat_name: d.show_cat_name,
            show_ser_name: d.show_ser_name,
            show_percent: d.show_percent,
            position: d.position,
            font_color: d.font_color,
            format_code: d.format_code,
            font_bold: d.font_bold,
            font_size_hpt: d.font_size_hpt,
        }
    }
}

impl From<ErrBars> for CmErrBars {
    fn from(e: ErrBars) -> Self {
        CmErrBars {
            dir: e.dir,
            bar_type: e.bar_type,
            plus: e.plus,
            minus: e.minus,
            no_end_cap: e.no_end_cap,
            color: e.color,
            line_width_emu: e.line_width_emu,
            dash: e.dash,
        }
    }
}

/// Maps `Vec<T>` → `Some(Vec<U>)` when non-empty, `None` when empty — the Rust
/// form of the adapter's `xs.length > 0 ? xs : null` / `xs ?? null` on the
/// per-series override / errbar arrays.
fn some_if_nonempty<T, U: From<T>>(xs: Vec<T>) -> Option<Vec<U>> {
    if xs.is_empty() {
        None
    } else {
        Some(xs.into_iter().map(U::from).collect())
    }
}

impl From<ChartSeries> for CmSeries {
    fn from(s: ChartSeries) -> Self {
        CmSeries {
            name: s.name,
            color: s.color,
            values: s.values,
            // xlsx never resolves these (pptx chartEx-only) — matches the
            // adapter, which omitted them.
            data_point_colors: None,
            data_label_colors: None,
            label_color: s.label_color,
            series_type: Some(s.series_type),
            use_secondary_axis: None,
            // `categories.length > 0 ? categories : null`.
            categories: if s.categories.is_empty() {
                None
            } else {
                Some(s.categories)
            },
            show_marker: Some(s.show_marker),
            val_format_code: s.val_format_code,
            marker_symbol: s.marker_symbol,
            marker_size: s.marker_size.map(|v| v as f64),
            marker_fill: s.marker_fill,
            marker_line: s.marker_line,
            data_point_overrides: some_if_nonempty(s.data_point_overrides),
            data_label_overrides: some_if_nonempty(s.data_label_overrides),
            series_data_labels: s.series_data_labels.map(CmSeriesDataLabels::from),
            err_bars: some_if_nonempty(s.err_bars),
            bubble_sizes: None,
            // `order` is xlsx parse-time only (used for stacking/legend sort
            // before emit); core `ChartSeries` has no such field.
        }
    }
}

impl From<ChartData> for ChartModel {
    fn from(c: ChartData) -> Self {
        // The white-default rule (adapter): when a `<c:chartSpace><c:spPr>` was
        // present we honor whatever it resolved to (solid hex or noFill→None);
        // when spPr was absent the file relies on Excel's default opaque-white
        // chart area, so we substitute white.
        let chart_bg = if c.has_chart_sp_pr {
            c.chart_bg
        } else {
            Some("FFFFFF".to_string())
        };
        ChartModel {
            chart_type: canonical_chart_type(&c.chart_type, &c.bar_dir, &c.grouping),
            title: c.title,
            categories: c.categories,
            series: c.series.into_iter().map(CmSeries::from).collect(),
            show_data_labels: c.show_data_labels,
            // Adapter mapped `valAxisMin/Max` → `valMin/Max`.
            val_min: c.val_axis_min,
            val_max: c.val_axis_max,
            cat_axis_title: c.cat_axis_title,
            val_axis_title: c.val_axis_title,
            cat_axis_hidden: c.cat_axis_hidden,
            val_axis_hidden: c.val_axis_hidden,
            cat_axis_line_hidden: c.cat_axis_line_hidden,
            val_axis_line_hidden: c.val_axis_line_hidden,
            // Adapter hard-coded `plotAreaBg: null` (xlsx never resolves it).
            plot_area_bg: None,
            chart_bg,
            show_legend: c.show_legend,
            legend_pos: c.legend_pos,
            // Adapter hard-coded `catAxisCrossBetween: 'between'`.
            cat_axis_cross_between: "between".to_string(),
            // Adapter default `?? 'out'` (ECMA-376 §21.2.2.49 ST_TickMark).
            val_axis_major_tick_mark: c
                .val_axis_major_tick_mark
                .unwrap_or_else(|| "out".to_string()),
            cat_axis_major_tick_mark: c
                .cat_axis_major_tick_mark
                .unwrap_or_else(|| "out".to_string()),
            title_font_size_hpt: c.title_font_size_hpt,
            title_font_color: c.title_font_color,
            title_font_face: c.title_font_face,
            cat_axis_font_size_hpt: c.cat_axis_font_size_hpt,
            val_axis_font_size_hpt: c.val_axis_font_size_hpt,
            // Adapter hard-coded `dataLabelFontSizeHpt: null` (xlsx never sets it).
            data_label_font_size_hpt: None,
            // Adapter hard-coded `subtotalIndices: []` (waterfall is pptx-only).
            subtotal_indices: vec![],
            val_axis_minor_tick_mark: c.val_axis_minor_tick_mark,
            cat_axis_minor_tick_mark: c.cat_axis_minor_tick_mark,
            // xlsx never resolves axis tick-label colors → adapter omitted them.
            cat_axis_font_color: None,
            val_axis_font_color: None,
            legend_manual_layout: c.legend_manual_layout.map(CmLegendManualLayout::from),
            val_axis_format_code: c.val_axis_format_code,
            bar_gap_width: c.bar_gap_width,
            bar_overlap: c.bar_overlap,
            data_label_position: c.data_label_position,
            data_label_font_color: c.data_label_font_color,
            data_label_format_code: c.data_label_format_code,
            title_font_bold: c.title_font_bold,
            cat_axis_font_bold: c.cat_axis_font_bold,
            val_axis_font_bold: c.val_axis_font_bold,
            cat_axis_title_font_size_hpt: c.cat_axis_title_size,
            cat_axis_title_font_bold: c.cat_axis_title_bold,
            cat_axis_title_font_color: c.cat_axis_title_color,
            val_axis_title_font_size_hpt: c.val_axis_title_size,
            val_axis_title_font_bold: c.val_axis_title_bold,
            val_axis_title_font_color: c.val_axis_title_color,
            chart_border_color: c.chart_border_color,
            chart_border_width_emu: c.chart_border_width_emu,
            cat_axis_crosses: c.cat_axis_crosses,
            cat_axis_crosses_at: c.cat_axis_crosses_at,
            val_axis_crosses: c.val_axis_crosses,
            val_axis_crosses_at: c.val_axis_crosses_at,
            cat_axis_line_color: c.cat_axis_line_color,
            cat_axis_line_width_emu: c.cat_axis_line_width_emu,
            val_axis_line_color: c.val_axis_line_color,
            val_axis_line_width_emu: c.val_axis_line_width_emu,
            cat_axis_format_code: c.cat_axis_format_code,
            cat_axis_min: c.cat_axis_min,
            cat_axis_max: c.cat_axis_max,
            title_manual_layout: c.title_manual_layout.map(CmManualLayout::from),
            plot_area_manual_layout: c.plot_area_manual_layout.map(CmManualLayout::from),
            // xlsx combo charts are not implemented → no scatter style / secondary
            // axis (adapter omitted both). scatterStyle is scatter-only and xlsx
            // doesn't parse it; secondaryValAxis is pptx-only.
            scatter_style: None,
            radar_style: c.radar_style,
            secondary_val_axis: None,
            date1904: c.date1904,
        }
    }
}

/// Given a sheet path (e.g. "worksheets/sheet1.xml"), locate and parse
/// its drawing(s) for chart anchors (`<xdr:graphicFrame>` elements).
pub(crate) fn load_sheet_charts(
    archive: &mut crate::XlsxZip,
    sheet_path: &str,
    theme_colors: &[String],
) -> Vec<ChartAnchor> {
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else {
        return Vec::new();
    };
    let sheet_rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let Ok(sheet_rels_xml) = read_zip_string(archive, &sheet_rels_path) else {
        return Vec::new();
    };
    let Ok(rels_doc) = roxmltree::Document::parse(&sheet_rels_xml) else {
        return Vec::new();
    };

    // Collect all drawing relationship targets
    let mut drawing_targets: Vec<String> = Vec::new();
    for rel in rels_doc
        .root_element()
        .children()
        .filter(|n| n.is_element())
    {
        if rel.attribute("Type").unwrap_or("").ends_with("/drawing") {
            if let Some(t) = rel.attribute("Target") {
                drawing_targets.push(t.to_string());
            }
        }
    }
    if drawing_targets.is_empty() {
        return Vec::new();
    }

    let mut all_charts: Vec<ChartAnchor> = Vec::new();

    for target in drawing_targets {
        // Resolve drawing path relative to the sheet directory
        let drawing_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        let Ok(drawing_xml) = read_zip_string(archive, &drawing_path) else {
            continue;
        };
        let Ok(draw_doc) = roxmltree::Document::parse(&drawing_xml) else {
            continue;
        };

        // Load drawing rels (to resolve chart rIds)
        let Some((drawing_dir, drawing_file)) = drawing_path.rsplit_once('/') else {
            continue;
        };
        let drawing_rels_path = format!("{}/_rels/{}.rels", drawing_dir, drawing_file);
        let drawing_rels = read_zip_string(archive, &drawing_rels_path)
            .ok()
            .map(|xml| parse_rels_map(&xml))
            .unwrap_or_default();

        // Iterate over chart anchors. Charts may be saved either as a
        // `<xdr:twoCellAnchor>` (from + to cells — Excel's default) or a
        // `<xdr:oneCellAnchor>` (from cell + a saved `<xdr:ext cx cy>` EMU
        // size, ECMA-376 §20.5.2.16). openpyxl and some other writers emit
        // oneCellAnchor; both must produce a chart.
        for anchor in draw_doc
            .root_element()
            .children()
            .filter(|n| n.is_element())
        {
            let anchor_tag = anchor.tag_name().name();
            let is_one_cell = anchor_tag == "oneCellAnchor";
            if (anchor_tag != "twoCellAnchor" && !is_one_cell)
                || !is_xdr_ns(anchor.tag_name().namespace())
            {
                continue;
            }

            let (mut from_col, mut from_col_off, mut from_row, mut from_row_off) =
                (0u32, 0i64, 0u32, 0i64);
            let (mut to_col, mut to_col_off, mut to_row, mut to_row_off) = (0u32, 0i64, 0u32, 0i64);
            // oneCellAnchor size: `<xdr:ext cx cy>` in EMU.
            let (mut ext_cx, mut ext_cy) = (0i64, 0i64);
            let mut chart_rid: Option<String> = None;

            for child in anchor.children() {
                if !child.is_element() {
                    continue;
                }
                match child.tag_name().name() {
                    "from" | "to" => {
                        let is_from = child.tag_name().name() == "from";
                        let mut col: u32 = 0;
                        let mut col_off: i64 = 0;
                        let mut row: u32 = 0;
                        let mut row_off: i64 = 0;
                        for c in child.children() {
                            match (c.tag_name().name(), c.text()) {
                                ("col", Some(t)) => col = t.trim().parse().unwrap_or(0),
                                ("colOff", Some(t)) => col_off = t.trim().parse().unwrap_or(0),
                                ("row", Some(t)) => row = t.trim().parse().unwrap_or(0),
                                ("rowOff", Some(t)) => row_off = t.trim().parse().unwrap_or(0),
                                _ => {}
                            }
                        }
                        if is_from {
                            from_col = col;
                            from_col_off = col_off;
                            from_row = row;
                            from_row_off = row_off;
                        } else {
                            to_col = col;
                            to_col_off = col_off;
                            to_row = row;
                            to_row_off = row_off;
                        }
                    }
                    "ext" => {
                        // oneCellAnchor's `<xdr:ext cx cy>` size in EMU.
                        ext_cx = child
                            .attribute("cx")
                            .and_then(|v| v.parse().ok())
                            .unwrap_or(0);
                        ext_cy = child
                            .attribute("cy")
                            .and_then(|v| v.parse().ok())
                            .unwrap_or(0);
                    }
                    "graphicFrame" => {
                        // ECMA-376 §20.1.2.2.8 CT_NonVisualDrawingProps@hidden:
                        // a hidden chart's own `<xdr:nvGraphicFramePr>/<xdr:cNvPr
                        // hidden="1">` marks it not rendered. This is the chart's
                        // own graphicFrame walk (independent of the shared shape
                        // walker in drawing.rs::collect_shapes), so it needs its
                        // own check.
                        if crate::drawing::xdr_node_hidden(&child) {
                            continue;
                        }
                        // Look for a:graphic/a:graphicData/c:chart[@r:id]
                        for gf_child in child.descendants() {
                            if gf_child.tag_name().name() == "chart"
                                && is_c_ns(gf_child.tag_name().namespace())
                            {
                                if let Some(rid) = gf_child
                                    .attributes()
                                    .find(|a| a.name() == "id" && is_r_ns(a.namespace()))
                                    .map(|a| a.value().to_string())
                                {
                                    chart_rid = Some(rid);
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }

            // For a oneCellAnchor the `<to>` corner is absent; the chart's
            // size is the saved EMU extent (ECMA-376 §20.5.2.16). Encode it as
            // a `to` corner pinned to the `from` cell plus the extent offset so
            // the renderer's from/to → pixel math yields exactly `ext` px.
            if is_one_cell {
                to_col = from_col;
                to_row = from_row;
                to_col_off = from_col_off + ext_cx;
                to_row_off = from_row_off + ext_cy;
            }

            let Some(rid) = chart_rid else {
                continue;
            };
            let Some(chart_target) = drawing_rels.get(&rid) else {
                continue;
            };
            let chart_path = resolve_zip_path(drawing_dir, chart_target);
            let Ok(chart_xml) = read_zip_string(archive, &chart_path) else {
                continue;
            };
            let Some(chart_data) = parse_chart_xml(&chart_xml, theme_colors) else {
                continue;
            };

            all_charts.push(ChartAnchor {
                from_col,
                from_col_off,
                from_row,
                from_row_off,
                to_col,
                to_col_off,
                to_row,
                to_row_off,
                // Adapt the parser-native `ChartData` into the canonical
                // `ChartModel` at emit time (was the TS `adaptChartData`).
                chart: chart_data.into(),
            });
        }
    }
    all_charts
}

// ─── Chart XML parser ────────────────────────────────────────────────────────

/// Parse a `xl/charts/chartN.xml` file into a `ChartData`.
pub(crate) fn parse_chart_xml(xml: &str, theme_colors: &[String]) -> Option<ChartData> {
    let doc = roxmltree::Document::parse(xml).ok()?;

    // Find c:chart root element
    let chart_root = doc
        .descendants()
        .find(|n| n.tag_name().name() == "chart" && is_c_ns(n.tag_name().namespace()))?;

    // Parse optional title
    let title = extract_chart_title(&chart_root);
    let title_font_size_hpt = extract_chart_title_size(&chart_root);
    let title_font_color = extract_chart_title_color(&chart_root);
    let title_font_face = extract_chart_title_face(&chart_root);
    let title_font_bold = extract_chart_title_bold(&chart_root);

    // Legend presence: <c:chart><c:legend> is the authoritative signal. Absence
    // means Excel hides the legend (default for a single-series chart with no
    // explicit legend element). `<c:legendPos val>` picks a side per
    // ECMA-376 §21.2.2.10 — both parts come from the shared ooxml-common helper
    // so pptx & xlsx stay in lockstep.
    let (show_legend, legend_pos) = ooxml_common::chart::extract_legend(chart_root);
    let legend_node = chart_root
        .children()
        .find(|n| n.tag_name().name() == "legend" && is_c_ns(n.tag_name().namespace()));
    // Legend <c:layout><c:manualLayout> (ECMA-376 §21.2.2.31) — when present,
    // gives explicit x/y/w/h fractions of the chart space. Used by the Excel
    // templates that position a top legend into a narrow band, e.g. over the
    // left half of the chart. We just collect the raw fractions here; the
    // renderer decides whether to honor `edge` vs `factor` placement.
    let legend_manual_layout = legend_node.and_then(|ln| {
        let layout = ln
            .children()
            .find(|n| n.tag_name().name() == "layout" && is_c_ns(n.tag_name().namespace()))?;
        let manual = layout
            .children()
            .find(|n| n.tag_name().name() == "manualLayout" && is_c_ns(n.tag_name().namespace()))?;
        let val = |tag: &str| {
            manual
                .children()
                .find(|n| n.tag_name().name() == tag && is_c_ns(n.tag_name().namespace()))
                .and_then(|n| n.attribute("val").and_then(|v| v.parse::<f64>().ok()))
        };
        let mode = |tag: &str| {
            manual
                .children()
                .find(|n| n.tag_name().name() == tag && is_c_ns(n.tag_name().namespace()))
                .and_then(|n| n.attribute("val").map(|v| v.to_string()))
                .unwrap_or_else(|| "edge".to_string())
        };
        Some(LegendManualLayout {
            x_mode: mode("xMode"),
            y_mode: mode("yMode"),
            x: val("x").unwrap_or(0.0),
            y: val("y").unwrap_or(0.0),
            w: val("w").unwrap_or(0.0),
            h: val("h").unwrap_or(0.0),
        })
    });

    // `<c:chartSpace><c:spPr>` outer fill (ECMA-376 §21.2.2.5). When the
    // element exists and carries `<a:noFill/>` the chart space is
    // transparent — this sample explicitly does that so the underlying
    // gray cell panel shows through. `<a:solidFill>` is resolved against
    // the theme just like series fills. When the element is absent we leave
    // `chart_bg` unset and tell the adapter to use the default opaque white
    // via `has_chart_sp_pr=false`.
    let chart_space_root = doc
        .descendants()
        .find(|n| n.tag_name().name() == "chartSpace" && is_c_ns(n.tag_name().namespace()));
    let chart_sp_pr = chart_space_root.and_then(|cs| {
        cs.children()
            .find(|n| n.tag_name().name() == "spPr" && is_c_ns(n.tag_name().namespace()))
    });
    let has_chart_sp_pr = chart_sp_pr.is_some();
    let chart_bg = chart_sp_pr.and_then(|sp| {
        // Walk direct children: noFill → None, solidFill → resolved color.
        let mut resolved: Option<String> = None;
        for ch in sp.children().filter(|n| n.is_element()) {
            match ch.tag_name().name() {
                "noFill" => {
                    return None;
                }
                "solidFill" => {
                    resolved = resolve_fill_color(&ch, theme_colors);
                    break;
                }
                _ => {}
            }
        }
        resolved
    });
    // Explicit chart border from `<c:chartSpace><c:spPr><a:ln>` (ECMA-376
    // §21.2.2.5 / DrawingML §20.1.2.2.24). Shared with the pptx parser via
    // `ooxml_common::chart::extract_chart_space_border` so the locked policy
    // (border only on an explicit paintable line; noFill → color None; srgb
    // inside solidFill → hex; schemeClr left unresolved) stays in lockstep.
    let (chart_border_color, chart_border_width_emu) = chart_space_root
        .map(ooxml_common::chart::extract_chart_space_border)
        .unwrap_or((None, None));

    // `<c:date1904>` (ECMA-376 §21.2.2.38) — a direct child of `<c:chartSpace>`.
    // Shared with the pptx chart parser via ooxml-common so both stay in
    // lockstep on the CT_Boolean val semantics (implied-true when present).
    let date1904 = chart_space_root
        .map(ooxml_common::chart::extract_chart_date1904)
        .unwrap_or(false);

    // `<c:title><c:layout><c:manualLayout>` (ECMA-376 §21.2.2.27).
    let title_manual_layout = chart_root
        .children()
        .find(|n| n.tag_name().name() == "title" && is_c_ns(n.tag_name().namespace()))
        .and_then(|t| {
            t.children()
                .find(|n| n.tag_name().name() == "layout" && is_c_ns(n.tag_name().namespace()))
        })
        .and_then(|l| extract_manual_layout(&l));

    // Find c:plotArea
    let plot_area = chart_root
        .children()
        .find(|n| n.tag_name().name() == "plotArea" && is_c_ns(n.tag_name().namespace()))?;
    let plot_area_manual_layout = plot_area
        .children()
        .find(|n| n.tag_name().name() == "layout" && is_c_ns(n.tag_name().namespace()))
        .and_then(|l| extract_manual_layout(&l));

    let mut primary_type = String::new();
    let mut bar_dir = "col".to_string();
    let mut grouping = "clustered".to_string();
    // `grouping` is recorded only from the first non-line chart-type element that
    // explicitly sets it. In combo charts (e.g. `<c:barChart grouping="stacked">`
    // followed by `<c:lineChart grouping="standard">`) the lineChart's grouping
    // must not overwrite the bar's, since stacking is a bar/area concept.
    let mut grouping_locked = false;
    let mut all_series: Vec<ChartSeries> = Vec::new();
    let mut shared_categories: Vec<String> = Vec::new();
    let mut show_data_labels = false;
    let mut cat_axis_title: Option<String> = None;
    let mut val_axis_title: Option<String> = None;
    // Axis-title run properties (sz/b/color). Only populated when the axis has a
    // `<c:title>` — the same helpers used for the chart title work unchanged on
    // an axis node because they scope to the node's direct-child `<c:title>`.
    let mut cat_axis_title_size: Option<i32> = None;
    let mut cat_axis_title_bold: Option<bool> = None;
    let mut cat_axis_title_color: Option<String> = None;
    let mut val_axis_title_size: Option<i32> = None;
    let mut val_axis_title_bold: Option<bool> = None;
    let mut val_axis_title_color: Option<String> = None;
    let mut cat_axis_font_size_hpt: Option<i32> = None;
    let mut val_axis_font_size_hpt: Option<i32> = None;
    let mut val_axis_format_code: Option<String> = None;
    // ECMA-376 §21.2.2.40 — `<c:delete val="1"/>` on a `<c:catAx>`/`<c:valAx>`
    // hides the axis (labels, ticks, and lines). Default is "0" (visible).
    let mut cat_axis_hidden = false;
    let mut val_axis_hidden = false;
    // `<c:catAx|valAx><c:spPr><a:ln><a:noFill>` — line-only hide; labels stay.
    let mut cat_axis_line_hidden = false;
    let mut val_axis_line_hidden = false;
    let mut cat_axis_format_code: Option<String> = None;
    let mut cat_axis_min: Option<f64> = None;
    let mut cat_axis_max: Option<f64> = None;
    let mut val_axis_min: Option<f64> = None;
    let mut val_axis_max: Option<f64> = None;
    let mut cat_axis_font_bold: Option<bool> = None;
    let mut val_axis_font_bold: Option<bool> = None;
    let mut cat_axis_line_color: Option<String> = None;
    let mut cat_axis_line_width_emu: Option<u32> = None;
    let mut val_axis_line_color: Option<String> = None;
    let mut val_axis_line_width_emu: Option<u32> = None;
    let mut cat_axis_major_tick_mark: Option<String> = None;
    let mut cat_axis_minor_tick_mark: Option<String> = None;
    let mut val_axis_major_tick_mark: Option<String> = None;
    let mut val_axis_minor_tick_mark: Option<String> = None;
    let mut cat_axis_crosses: Option<String> = None;
    let mut cat_axis_crosses_at: Option<f64> = None;
    let mut val_axis_crosses: Option<String> = None;
    let mut val_axis_crosses_at: Option<f64> = None;
    let mut bar_gap_width: Option<i32> = None;
    let mut bar_overlap: Option<i32> = None;
    let mut radar_style: Option<String> = None;
    let mut data_label_position: Option<String> = None;
    let mut data_label_format_code: Option<String> = None;
    // Data-label text color is theme-aware, so we hoist the extraction to a
    // shared helper backed by an `XlsxColorResolver`. It walks chart-level
    // and per-series `<c:dLbls><c:txPr>...<a:solidFill>` in document order
    // and returns the first one it can resolve.
    let xlsx_resolver = XlsxColorResolver { theme_colors };
    let data_label_font_color =
        ooxml_common::chart::extract_data_label_font_color(chart_root, &xlsx_resolver);

    // Recognised chart-type element names → our internal type strings
    let type_map: &[(&str, &str)] = &[
        ("barChart", "bar"),
        ("lineChart", "line"),
        ("areaChart", "area"),
        ("pieChart", "pie"),
        ("doughnutChart", "doughnut"),
        ("radarChart", "radar"),
        ("scatterChart", "scatter"),
        ("bubbleChart", "scatter"), // treat bubble as scatter
    ];

    for child in plot_area.children() {
        if !child.is_element() {
            continue;
        }
        if !is_c_ns(child.tag_name().namespace()) {
            continue;
        }
        let elem_name = child.tag_name().name();

        // Axis title + tick label font size extraction (ECMA-376 §21.2.2.17
        // c:txPr/a:defRPr@sz gives tick labels their hpt size; absent = default).
        match elem_name {
            // `<c:dateAx>` (§21.2.2.39) is the date/time-series category axis:
            // same child grammar as `<c:catAx>` (title, numFmt, delete, ticks,
            // scaling, line style). Excel emits it when the category source is
            // dates; the numFmt formatCode drives serial-date label formatting on
            // the TS side. Advanced date-unit control (baseTimeUnit/majorTimeUnit)
            // is out of scope here — treated identically to catAx.
            "catAx" | "dateAx" => {
                if cat_axis_title.is_none() {
                    let (t, sz, b, c) = extract_axis_title_with_props(&child);
                    if t.is_some() {
                        cat_axis_title = t;
                        cat_axis_title_size = sz;
                        cat_axis_title_bold = b;
                        cat_axis_title_color = c;
                    }
                }
                if cat_axis_font_size_hpt.is_none() {
                    cat_axis_font_size_hpt = extract_axis_tick_label_size(&child);
                }
                if cat_axis_format_code.is_none() {
                    cat_axis_format_code = extract_axis_format_code(&child);
                }
                if cat_axis_font_bold.is_none() {
                    cat_axis_font_bold = extract_axis_tick_label_bold(&child);
                }
                if cat_axis_line_color.is_none() || cat_axis_line_width_emu.is_none() {
                    let (col, w) = extract_axis_line_style(&child, theme_colors);
                    if cat_axis_line_color.is_none() {
                        cat_axis_line_color = col;
                    }
                    if cat_axis_line_width_emu.is_none() {
                        cat_axis_line_width_emu = w;
                    }
                }
                if cat_axis_major_tick_mark.is_none() {
                    cat_axis_major_tick_mark = extract_axis_tick_mark(&child, "majorTickMark");
                }
                if cat_axis_minor_tick_mark.is_none() {
                    cat_axis_minor_tick_mark = extract_axis_tick_mark(&child, "minorTickMark");
                }
                if let Some((mn, mx)) = extract_axis_scaling(&child) {
                    if cat_axis_min.is_none() {
                        cat_axis_min = mn;
                    }
                    if cat_axis_max.is_none() {
                        cat_axis_max = mx;
                    }
                }
                let (cr, cra) = extract_axis_crosses(&child);
                if cat_axis_crosses.is_none() {
                    cat_axis_crosses = cr;
                }
                if cat_axis_crosses_at.is_none() {
                    cat_axis_crosses_at = cra;
                }
                if axis_is_deleted(&child) {
                    cat_axis_hidden = true;
                }
                if axis_line_is_hidden(&child) {
                    cat_axis_line_hidden = true;
                }
                continue;
            }
            "valAx" => {
                // Scatter charts use two `<c:valAx>` (no catAx). Disambiguate
                // by `<c:axPos val>` — `b`(bottom)/`t`(top) → X axis, `l`/`r`
                // → Y axis. For non-scatter charts the first valAx hit is
                // always Y.
                let ax_pos = child
                    .children()
                    .find(|n| n.tag_name().name() == "axPos" && is_c_ns(n.tag_name().namespace()))
                    .and_then(|n| n.attribute("val"))
                    .unwrap_or("");
                let is_x_axis = matches!(ax_pos, "b" | "t");
                if is_x_axis {
                    // A scatter chart's bottom `<c:valAx>` is the horizontal axis,
                    // so its title is the cat-axis (X) title. Without this the
                    // X-axis title of every scatter chart was dropped (the catAx
                    // branch above never runs for scatter — there is no catAx).
                    if cat_axis_title.is_none() {
                        let (t, sz, b, c) = extract_axis_title_with_props(&child);
                        if t.is_some() {
                            cat_axis_title = t;
                            cat_axis_title_size = sz;
                            cat_axis_title_bold = b;
                            cat_axis_title_color = c;
                        }
                    }
                    if cat_axis_format_code.is_none() {
                        cat_axis_format_code = extract_axis_format_code(&child);
                    }
                    if let Some((mn, mx)) = extract_axis_scaling(&child) {
                        if cat_axis_min.is_none() {
                            cat_axis_min = mn;
                        }
                        if cat_axis_max.is_none() {
                            cat_axis_max = mx;
                        }
                    }
                    if cat_axis_font_size_hpt.is_none() {
                        cat_axis_font_size_hpt = extract_axis_tick_label_size(&child);
                    }
                    if cat_axis_font_bold.is_none() {
                        cat_axis_font_bold = extract_axis_tick_label_bold(&child);
                    }
                    if cat_axis_line_color.is_none() || cat_axis_line_width_emu.is_none() {
                        let (col, w) = extract_axis_line_style(&child, theme_colors);
                        if cat_axis_line_color.is_none() {
                            cat_axis_line_color = col;
                        }
                        if cat_axis_line_width_emu.is_none() {
                            cat_axis_line_width_emu = w;
                        }
                    }
                    if cat_axis_major_tick_mark.is_none() {
                        cat_axis_major_tick_mark = extract_axis_tick_mark(&child, "majorTickMark");
                    }
                    if cat_axis_minor_tick_mark.is_none() {
                        cat_axis_minor_tick_mark = extract_axis_tick_mark(&child, "minorTickMark");
                    }
                    let (cr, cra) = extract_axis_crosses(&child);
                    if cat_axis_crosses.is_none() {
                        cat_axis_crosses = cr;
                    }
                    if cat_axis_crosses_at.is_none() {
                        cat_axis_crosses_at = cra;
                    }
                    if axis_is_deleted(&child) {
                        cat_axis_hidden = true;
                    }
                    if axis_line_is_hidden(&child) {
                        cat_axis_line_hidden = true;
                    }
                } else {
                    if val_axis_title.is_none() {
                        let (t, sz, b, c) = extract_axis_title_with_props(&child);
                        if t.is_some() {
                            val_axis_title = t;
                            val_axis_title_size = sz;
                            val_axis_title_bold = b;
                            val_axis_title_color = c;
                        }
                    }
                    if val_axis_font_size_hpt.is_none() {
                        val_axis_font_size_hpt = extract_axis_tick_label_size(&child);
                    }
                    if val_axis_format_code.is_none() {
                        val_axis_format_code = extract_axis_format_code(&child);
                    }
                    if val_axis_font_bold.is_none() {
                        val_axis_font_bold = extract_axis_tick_label_bold(&child);
                    }
                    if let Some((mn, mx)) = extract_axis_scaling(&child) {
                        if val_axis_min.is_none() {
                            val_axis_min = mn;
                        }
                        if val_axis_max.is_none() {
                            val_axis_max = mx;
                        }
                    }
                    let (cr, cra) = extract_axis_crosses(&child);
                    if val_axis_crosses.is_none() {
                        val_axis_crosses = cr;
                    }
                    if val_axis_crosses_at.is_none() {
                        val_axis_crosses_at = cra;
                    }
                    if val_axis_line_color.is_none() || val_axis_line_width_emu.is_none() {
                        let (col, w) = extract_axis_line_style(&child, theme_colors);
                        if val_axis_line_color.is_none() {
                            val_axis_line_color = col;
                        }
                        if val_axis_line_width_emu.is_none() {
                            val_axis_line_width_emu = w;
                        }
                    }
                    if val_axis_major_tick_mark.is_none() {
                        val_axis_major_tick_mark = extract_axis_tick_mark(&child, "majorTickMark");
                    }
                    if val_axis_minor_tick_mark.is_none() {
                        val_axis_minor_tick_mark = extract_axis_tick_mark(&child, "minorTickMark");
                    }
                    if axis_is_deleted(&child) {
                        val_axis_hidden = true;
                    }
                    if axis_line_is_hidden(&child) {
                        val_axis_line_hidden = true;
                    }
                }
                continue;
            }
            _ => {}
        }

        let ser_type = match type_map.iter().find(|(k, _)| *k == elem_name) {
            Some((_, v)) => *v,
            None => continue,
        };

        if primary_type.is_empty() {
            primary_type = ser_type.to_string();
        }

        // barDir / grouping / dLbls / marker (only meaningful for bar/line/area).
        // <c:marker val> at the chart-type element is the default for all line
        // series in that element (ECMA-376 §21.2.2.33). "1" = markers visible.
        let mut chart_marker_default = false;
        for attr_node in child.children().filter(|n| n.is_element()) {
            match attr_node.tag_name().name() {
                "barDir" => {
                    bar_dir = attr_node.attribute("val").unwrap_or("col").to_string();
                }
                "grouping" => {
                    let val = attr_node
                        .attribute("val")
                        .unwrap_or("clustered")
                        .to_string();
                    if !grouping_locked && ser_type != "line" {
                        grouping = val;
                        grouping_locked = true;
                    }
                }
                "marker" => {
                    chart_marker_default = attr_node.attribute("val").unwrap_or("0") != "0";
                }
                "gapWidth" => {
                    // ECMA-376 §21.2.2.13 — percent of bar width between category
                    // groups (default 150 per spec). Only meaningful on bar charts.
                    if bar_gap_width.is_none() {
                        bar_gap_width = attr_node.attribute("val").and_then(|v| v.parse().ok());
                    }
                }
                "overlap" => {
                    // ECMA-376 §21.2.2.25 — signed percent of bar width for the
                    // overlap within a cluster (negative = gap).
                    if bar_overlap.is_none() {
                        bar_overlap = attr_node.attribute("val").and_then(|v| v.parse().ok());
                    }
                }
                "radarStyle" => {
                    // ECMA-376 §21.2.3.10 — "standard" (line only, default),
                    // "marker" (line + markers), "filled" (closed polygon
                    // with area fill). Only the "filled" variant should
                    // produce a translucent area in the renderer.
                    if radar_style.is_none() {
                        radar_style = attr_node.attribute("val").map(|s| s.to_string());
                    }
                }
                "dLbls" => {
                    for d in attr_node.children().filter(|n| n.is_element()) {
                        match d.tag_name().name() {
                            "showVal" | "showPercent" => {
                                if d.attribute("val").unwrap_or("1") != "0" {
                                    show_data_labels = true;
                                }
                            }
                            "dLblPos" => {
                                if data_label_position.is_none() {
                                    data_label_position = d.attribute("val").map(|s| s.to_string());
                                }
                            }
                            "numFmt" if data_label_format_code.is_none() => {
                                data_label_format_code = d
                                    .attribute("formatCode")
                                    .map(|s| s.to_string())
                                    .filter(|s| !s.is_empty() && s != "General");
                            }
                            // txPr (data-label text color) is now resolved
                            // up front via ooxml_common::chart::extract_data_label_font_color.
                            _ => {}
                        }
                    }
                }
                _ => {}
            }
        }

        // Parse series
        for ser_node in child.children().filter(|n| {
            n.is_element() && n.tag_name().name() == "ser" && is_c_ns(n.tag_name().namespace())
        }) {
            let s = parse_chart_series(&ser_node, ser_type, chart_marker_default, theme_colors);
            if shared_categories.is_empty() && !s.categories.is_empty() {
                shared_categories = s.categories.clone();
            }
            all_series.push(s);

            // Per-series `<c:ser><c:dLbls>` fallback for chart-level
            // properties that Excel commonly writes on each series instead
            // of (or in addition to) the chart-level `<c:dLbls>`. Per
            // ECMA-376 §21.2.2.47 a series-level dLbls applies to that
            // series's data points; for the renderer we only need one
            // value, so the first series encountered "wins". This is how
            // the default-color/position/format travels when Excel emits
            // `<c:dLblPos val="inBase"/>` + `<c:txPr>` per series rather
            // than on the chart.
            if let Some(ser_dlbls) = ser_node.children().find(|n| {
                n.is_element()
                    && n.tag_name().name() == "dLbls"
                    && is_c_ns(n.tag_name().namespace())
            }) {
                for d in ser_dlbls.children().filter(|n| n.is_element()) {
                    match d.tag_name().name() {
                        "showVal" | "showPercent" => {
                            if d.attribute("val").unwrap_or("1") != "0" {
                                show_data_labels = true;
                            }
                        }
                        "dLblPos" => {
                            if data_label_position.is_none() {
                                data_label_position = d.attribute("val").map(|s| s.to_string());
                            }
                        }
                        "numFmt" if data_label_format_code.is_none() => {
                            data_label_format_code = d
                                .attribute("formatCode")
                                .map(|s| s.to_string())
                                .filter(|s| !s.is_empty() && s != "General");
                        }
                        // txPr (data-label text color) handled by the
                        // shared helper before this loop runs.
                        _ => {}
                    }
                }
            }
        }
    }

    if primary_type.is_empty() {
        return None;
    }

    // Fill in categories for series that have none (mixed charts share categories)
    for s in &mut all_series {
        if s.categories.is_empty() {
            s.categories = shared_categories.clone();
        }
    }

    // Stable-sort by `c:order` so the array is in Excel's display order.
    // ECMA-376 §21.2.2.28 — `<c:order>` is the authoritative stacking /
    // legend order, independent of document order.
    all_series.sort_by_key(|s| s.order);

    Some(ChartData {
        chart_type: primary_type,
        bar_dir,
        grouping,
        title,
        categories: shared_categories,
        series: all_series,
        show_data_labels,
        cat_axis_title,
        val_axis_title,
        show_legend,
        legend_pos,
        title_font_size_hpt,
        title_font_color,
        title_font_face,
        title_font_bold,
        cat_axis_font_bold,
        val_axis_font_bold,
        cat_axis_crosses,
        cat_axis_crosses_at,
        val_axis_crosses,
        val_axis_crosses_at,
        cat_axis_line_color,
        cat_axis_line_width_emu,
        val_axis_line_color,
        val_axis_line_width_emu,
        cat_axis_major_tick_mark,
        cat_axis_minor_tick_mark,
        val_axis_major_tick_mark,
        val_axis_minor_tick_mark,
        cat_axis_font_size_hpt,
        val_axis_font_size_hpt,
        val_axis_format_code,
        chart_bg,
        has_chart_sp_pr,
        legend_manual_layout,
        cat_axis_hidden,
        val_axis_hidden,
        cat_axis_line_hidden,
        val_axis_line_hidden,
        bar_gap_width,
        bar_overlap,
        data_label_position,
        data_label_font_color,
        data_label_format_code,
        cat_axis_format_code,
        cat_axis_min,
        cat_axis_max,
        val_axis_min,
        val_axis_max,
        title_manual_layout,
        plot_area_manual_layout,
        radar_style,
        cat_axis_title_size,
        cat_axis_title_bold,
        cat_axis_title_color,
        val_axis_title_size,
        val_axis_title_bold,
        val_axis_title_color,
        chart_border_color,
        chart_border_width_emu,
        date1904,
    })
}

/// `<c:catAx|valAx><c:numFmt@formatCode>` (ECMA-376 §21.2.2.21). Thin
/// wrapper around `ooxml_common::chart::extract_axis_format_code` so the
/// pptx and xlsx parsers stay in lockstep on the format-code rules.
pub(crate) fn extract_axis_format_code(axis_node: &roxmltree::Node) -> Option<String> {
    ooxml_common::chart::extract_axis_format_code(*axis_node)
}

/// `<c:catAx|valAx><c:majorTickMark val>` / `<c:minorTickMark val>` —
/// `none` / `out` / `in` / `cross` (ECMA-376 §21.2.2.49 ST_TickMark).
pub(crate) fn extract_axis_tick_mark(axis_node: &roxmltree::Node, name: &str) -> Option<String> {
    ooxml_common::chart::extract_axis_tick_mark(*axis_node, name)
}

/// `<c:catAx|valAx><c:spPr><a:ln>` — resolved color (no `#`) and width
/// (EMU) for the axis line itself. None when not set.
pub(crate) fn extract_axis_line_style(
    axis_node: &roxmltree::Node,
    theme_colors: &[String],
) -> (Option<String>, Option<u32>) {
    let Some(sp_pr) = axis_node
        .children()
        .find(|n| n.tag_name().name() == "spPr" && is_c_ns(n.tag_name().namespace()))
    else {
        return (None, None);
    };
    let Some(ln) = sp_pr.children().find(|n| n.tag_name().name() == "ln") else {
        return (None, None);
    };
    let width = ln.attribute("w").and_then(|v| v.parse::<u32>().ok());
    let color = extract_solid_fill_in_drawingml(&ln, theme_colors);
    (color, width)
}

/// `<c:catAx|valAx><c:spPr><a:ln><a:noFill>` — true when the axis line is
/// explicitly hidden (labels and tick marks still render). Distinct from
/// `<c:delete val="1"/>` which hides the entire axis. Sample-1 "Carbon &
/// Growth" uses this on `<c:valAx>` to keep the Y-axis numbers visible
/// while suppressing the vertical rule.
pub(crate) fn axis_line_is_hidden(axis_node: &roxmltree::Node) -> bool {
    let Some(sp_pr) = axis_node
        .children()
        .find(|n| n.tag_name().name() == "spPr" && is_c_ns(n.tag_name().namespace()))
    else {
        return false;
    };
    let Some(ln) = sp_pr.children().find(|n| n.tag_name().name() == "ln") else {
        return false;
    };
    ln.children()
        .any(|n| n.is_element() && n.tag_name().name() == "noFill")
}

/// `<c:catAx|valAx><c:txPr>...defRPr@b>` — bold flag for axis tick labels.
/// Thin wrapper around `ooxml_common::chart::extract_axis_tick_label_bold` so
/// the pptx and xlsx parsers stay in lockstep.
pub(crate) fn extract_axis_tick_label_bold(axis_node: &roxmltree::Node) -> Option<bool> {
    ooxml_common::chart::extract_axis_tick_label_bold(*axis_node)
}

/// `<c:catAx|valAx><c:crosses>` and `<c:crossesAt>` — where the axis sits
/// along its perpendicular axis. `crosses` is a string ("autoZero" |
/// "min" | "max"); `crossesAt` is an explicit numeric override that
/// takes precedence at render time.
pub(crate) fn extract_axis_crosses(axis_node: &roxmltree::Node) -> (Option<String>, Option<f64>) {
    let crosses = axis_node
        .children()
        .find(|n| n.tag_name().name() == "crosses" && is_c_ns(n.tag_name().namespace()))
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string());
    let crosses_at = axis_node
        .children()
        .find(|n| n.tag_name().name() == "crossesAt" && is_c_ns(n.tag_name().namespace()))
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok());
    (crosses, crosses_at)
}

/// Read explicit `<c:scaling><c:min>` / `<c:scaling><c:max>` values, returning
/// `(min, max)` where each is `None` if the axis didn't override that bound.
/// Returns `None` only when neither bound is set (matches the prior xlsx
/// callsite shape `if let Some((mn, mx)) = …`).
pub(crate) fn extract_axis_scaling(
    axis_node: &roxmltree::Node,
) -> Option<(Option<f64>, Option<f64>)> {
    let (mn, mx) = ooxml_common::chart::extract_axis_min_max(*axis_node);
    if mn.is_some() || mx.is_some() {
        Some((mn, mx))
    } else {
        None
    }
}

/// `<c:catAx|valAx><c:delete val="1"/>` — true when the axis is hidden
/// (ECMA-376 §21.2.2.40). Thin wrapper around the shared helper.
pub(crate) fn axis_is_deleted(axis_node: &roxmltree::Node) -> bool {
    ooxml_common::chart::axis_is_deleted(*axis_node)
}

/// Extract a `<c:layout><c:manualLayout>` block. The given `layout_node` is
/// `<c:layout>` (parent of `<c:manualLayout>`). Returns None when the layout
/// is auto (no `manualLayout` child).
pub(crate) fn extract_manual_layout(layout_node: &roxmltree::Node) -> Option<ManualLayout> {
    let ml = layout_node
        .children()
        .find(|n| n.tag_name().name() == "manualLayout" && is_c_ns(n.tag_name().namespace()))?;
    let mut x_mode = "edge".to_string();
    let mut y_mode = "edge".to_string();
    let mut layout_target: Option<String> = None;
    let mut x = 0.0_f64;
    let mut y = 0.0_f64;
    let mut w: Option<f64> = None;
    let mut h: Option<f64> = None;
    for ch in ml
        .children()
        .filter(|n| n.is_element() && is_c_ns(n.tag_name().namespace()))
    {
        let val = ch.attribute("val").map(|s| s.to_string());
        match ch.tag_name().name() {
            "xMode" => {
                if let Some(v) = val {
                    x_mode = v;
                }
            }
            "yMode" => {
                if let Some(v) = val {
                    y_mode = v;
                }
            }
            "layoutTarget" => {
                layout_target = val;
            }
            "x" => {
                if let Some(v) = ch.attribute("val").and_then(|s| s.parse::<f64>().ok()) {
                    x = v;
                }
            }
            "y" => {
                if let Some(v) = ch.attribute("val").and_then(|s| s.parse::<f64>().ok()) {
                    y = v;
                }
            }
            "w" => {
                w = ch.attribute("val").and_then(|s| s.parse::<f64>().ok());
            }
            "h" => {
                h = ch.attribute("val").and_then(|s| s.parse::<f64>().ok());
            }
            _ => {}
        }
    }
    Some(ManualLayout {
        x_mode,
        y_mode,
        layout_target,
        x,
        y,
        w,
        h,
    })
}

/// Extract a category/value axis tick-label font size (hundredths of a point)
/// from the first `a:defRPr@sz` (or `a:rPr@sz`) inside the axis' `c:txPr`.
/// ECMA-376 §21.2.2.17 — `<c:txPr>` controls tick label text properties.
pub(crate) fn extract_axis_tick_label_size(axis_node: &roxmltree::Node) -> Option<i32> {
    ooxml_common::chart::extract_axis_tick_label_size(*axis_node)
}

/// Extract the chart title's font size (hundredths of a point) from the first
/// `a:defRPr@sz` or `a:rPr@sz` found under `c:title`. Returns None when absent.
/// Thin wrapper around `ooxml_common::chart::extract_chart_title_size`.
pub(crate) fn extract_chart_title_size(chart_root: &roxmltree::Node) -> Option<i32> {
    ooxml_common::chart::extract_chart_title_size(*chart_root)
}

/// Extract chart title bold flag from the first `a:defRPr@b` / `a:rPr@b`
/// inside `c:title`. Returns None when not specified (renderer treats as
/// not bold). Thin wrapper around `ooxml_common::chart::extract_chart_title_bold`.
pub(crate) fn extract_chart_title_bold(chart_root: &roxmltree::Node) -> Option<bool> {
    ooxml_common::chart::extract_chart_title_bold(*chart_root)
}

/// Extract the chart title's font color (hex without '#') from the first
/// `a:solidFill/a:srgbClr@val` inside `c:title`. Only srgb is resolved here —
/// scheme colors would require the workbook theme, which isn't wired through
/// to chart parsing yet. Thin wrapper around
/// `ooxml_common::chart::extract_chart_title_srgb`.
pub(crate) fn extract_chart_title_color(chart_root: &roxmltree::Node) -> Option<String> {
    ooxml_common::chart::extract_chart_title_srgb(*chart_root)
}

/// Extract the chart title's font family from the first `a:latin@typeface`
/// descendant of `c:title` (ECMA-376 DrawingML §20.1.4.2.24).
pub(crate) fn extract_chart_title_face(chart_root: &roxmltree::Node) -> Option<String> {
    let title_node = chart_root
        .children()
        .find(|n| n.tag_name().name() == "title" && is_c_ns(n.tag_name().namespace()))?;
    title_node.descendants().find_map(|n| {
        if !n.is_element() {
            return None;
        }
        if !is_a_ns(n.tag_name().namespace()) {
            return None;
        }
        if n.tag_name().name() != "latin" {
            return None;
        }
        n.attribute("typeface").map(|s| s.to_string())
    })
}

/// Extract plain text from `c:chart/c:title`. Thin wrapper around
/// `ooxml_common::chart::extract_chart_title_text`.
pub(crate) fn extract_chart_title(chart_root: &roxmltree::Node) -> Option<String> {
    ooxml_common::chart::extract_chart_title_text(*chart_root)
}

/// Extract an axis title's text + run props from a `<c:catAx>`/`<c:valAx>` node.
/// Reuses the chart-title helpers (which scope to the node's direct-child
/// `<c:title>`); run props are resolved only when title text is present.
/// Returns (text, size_hpt, bold, color_hex). Thin wrapper around
/// `ooxml_common::chart::extract_axis_title_with_props`.
fn extract_axis_title_with_props(
    axis_node: &roxmltree::Node,
) -> (Option<String>, Option<i32>, Option<bool>, Option<String>) {
    ooxml_common::chart::extract_axis_title_with_props(*axis_node)
}

/// Parse one `<c:ser>` element.
/// Resolve the fill color from a single DrawingML fill element. The caller
/// passes either a `<c:spPr>` (in which case we look for the first `<a:solidFill>`
/// **as a direct child** to avoid picking up text fills nested under `<c:dLbls>`
/// / `<c:txPr>`) or the `<a:solidFill>` directly. Supports `a:srgbClr` (explicit
/// hex) and `a:schemeClr` (theme accent/dark/light).
/// Theme colors use drawingML names (`accent1`..`accent6`, `dk1`/`dk2`/`lt1`/`lt2`)
/// which map to the parser's natural-order theme array (dk1@0, lt1@1, dk2@2,
/// lt2@3, accent1@4 … accent6@9).
/// `ooxml_common::chart::ColorResolver` implementation backed by xlsx's
/// indexed `theme_colors` slice. Drives the chart helpers in ooxml-common
/// that need theme-aware color resolution (e.g. data-label text color)
/// without leaking the slice shape into the shared crate.
struct XlsxColorResolver<'a> {
    theme_colors: &'a [String],
}

impl ooxml_common::chart::ColorResolver for XlsxColorResolver<'_> {
    fn resolve_solid_fill(&self, node: roxmltree::Node<'_, '_>) -> Option<String> {
        resolve_fill_color(&node, self.theme_colors)
    }
}
/// Series fill color from `<c:ser><c:spPr><a:solidFill>`. Returns None when
/// the series has no direct `<c:spPr>` or its fill isn't a recognised solid.
pub(crate) fn resolve_series_color(
    ser_node: &roxmltree::Node,
    theme_colors: &[String],
) -> Option<String> {
    let sp_pr = ser_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spPr")?;
    // Line / scatter / radar series carry their color on `<a:ln><a:solidFill>`
    // (the series IS the stroke), not on `<a:solidFill>` directly under spPr.
    // Bar / area / pie series carry it on the direct `<a:solidFill>` (fill).
    // Try fill first (handles bar/area/pie/marker fill); if absent, fall back
    // to the line color (handles line/scatter/radar). Without this fallback
    // the renderer would lose the explicit `<a:ln><a:solidFill><a:srgbClr>`
    // overrides on line-chart series and rotate through the theme accents
    // instead (`demo/sample-1` "Carbon & Growth" "Year" series should be
    // #2D6A4F but rendered as accent1 = #156082 blue).
    if let Some(c) = resolve_fill_color(&sp_pr, theme_colors) {
        return Some(c);
    }
    let ln = sp_pr
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "ln")?;
    resolve_fill_color(&ln, theme_colors)
}

pub(crate) fn parse_chart_series(
    node: &roxmltree::Node,
    ser_type: &str,
    chart_marker_default: bool,
    theme_colors: &[String],
) -> ChartSeries {
    let name = extract_series_name(node);

    // `<c:idx val>` (ECMA-376 §21.2.2.27) — the canonical series index Excel
    // uses for default color selection. When absent, fall back to 0 so we
    // still produce a deterministic palette pick. `<c:order>` is the display
    // order (legend / stacking) and is intentionally ignored for coloring.
    let idx: usize = node
        .children()
        .find(|n| n.tag_name().name() == "idx" && is_c_ns(n.tag_name().namespace()))
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(0);

    // `<c:order val>` (ECMA-376 §21.2.2.28) — series display order. Used for
    // stacking and legend ordering. Defaults to 0.
    let order: usize = node
        .children()
        .find(|n| n.tag_name().name() == "order" && is_c_ns(n.tag_name().namespace()))
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(0);

    // For scatter: xVal → categories (as strings), yVal → values
    // For others:  cat  → categories,             val  → values
    let (cat_tag, val_tag) = if ser_type == "scatter" {
        ("xVal", "yVal")
    } else {
        ("cat", "val")
    };

    let categories = collect_str_cache(node, cat_tag);
    let values = collect_num_cache(node, val_tag);
    // `<c:val><c:numRef><c:numCache><c:formatCode>` (ECMA-376 §21.2.2.37)
    // preserves the Excel number format Excel stamped onto the cached values
    // at save time; absent "General" codes return None so the renderer can
    // fall back cleanly.
    let val_format_code = node
        .children()
        .find(|n| n.tag_name().name() == val_tag && is_c_ns(n.tag_name().namespace()))
        .and_then(|val_node| {
            val_node.descendants().find(|n| {
                n.is_element()
                    && n.tag_name().name() == "formatCode"
                    && is_c_ns(n.tag_name().namespace())
            })
        })
        .and_then(|n| n.text().map(|s| s.to_string()))
        .filter(|s| !s.is_empty() && s != "General");

    // Series fill color from c:spPr/a:solidFill (supports a:srgbClr and a:schemeClr).
    // For schemeClr, resolves "accentN"/"dk1"/etc. against the workbook theme.
    //
    // When the series has no explicit fill, Excel's default palette assigns
    // `theme.accent[idx % 6 + 1]` — i.e. accent1, accent2, … cycling by
    // `<c:idx>`. That's the rule behind "first series = green, second = red"
    // when the theme's accent1/accent2 are green/red. We inline that
    // resolution here so the renderer doesn't need theme access.
    let color = resolve_series_color(node, theme_colors).or_else(|| {
        // Theme order in `theme_colors`: dk1@0, lt1@1, dk2@2, lt2@3, accent1@4 … accent6@9.
        theme_colors
            .get(4 + (idx % 6))
            .map(|c| c.trim_start_matches('#').to_lowercase())
    });

    // Marker visibility (ECMA-376 §21.2.2.32 — c:marker/c:symbol default is
    // "none"). A per-series <c:marker><c:symbol> overrides; otherwise fall
    // back to the chart-type-level <c:lineChart><c:marker val> flag. Scatter
    // charts default to visible markers even without an explicit flag.
    let marker_node = node
        .children()
        .find(|n| n.tag_name().name() == "marker" && is_c_ns(n.tag_name().namespace()));
    let (marker_symbol, marker_size, marker_fill, marker_line) =
        parse_marker_block(marker_node, theme_colors);
    let show_marker = match (&marker_symbol, ser_type) {
        (Some(sym), _) => sym != "none",
        (None, "scatter") => true,
        _ => chart_marker_default,
    };

    // Per-series data-label text color from `<c:ser><c:dLbls><c:txPr>…solidFill`.
    // Scoped to THIS series (not chart-root) so stacked/clustered charts keep
    // their independent label colors instead of all collapsing to the single
    // chart-level color. Reuses the shared dLbls-txPr walker + xlsx theme
    // resolver. None when the series has no own dLbls text color.
    let label_color = ooxml_common::chart::extract_data_label_font_color(
        *node,
        &XlsxColorResolver { theme_colors },
    );

    let data_point_overrides = parse_data_point_overrides(node, theme_colors);
    // `<c15:datalabelsRange>` lookup table for `<a:fld type="CELLRANGE">`
    // labels. Excel saves the actual cached label strings here; we resolve
    // CELLRANGE field placeholders against this at parse time so the
    // renderer just receives plain strings.
    let dlbl_range_cache = collect_dlbl_range_cache(node);
    let (series_data_labels, data_label_overrides) =
        parse_data_labels(node, theme_colors, &dlbl_range_cache);
    let err_bars = parse_error_bars(node, &values, theme_colors);

    ChartSeries {
        name,
        series_type: ser_type.to_string(),
        categories,
        values,
        color,
        show_marker,
        val_format_code,
        label_color,
        order,
        marker_symbol,
        marker_size,
        marker_fill,
        marker_line,
        data_point_overrides,
        data_label_overrides,
        series_data_labels,
        err_bars,
    }
}

/// Parse `<c:marker>` into (symbol, size, fill, line) — all hex colors are
/// returned without `#`. ECMA-376 §21.2.2.32 / §21.2.2.34. The fill and
/// line colors come from `<c:spPr>` nested inside marker.
pub(crate) fn parse_marker_block(
    marker_node: Option<roxmltree::Node>,
    theme_colors: &[String],
) -> (Option<String>, Option<u32>, Option<String>, Option<String>) {
    let Some(mk) = marker_node else {
        return (None, None, None, None);
    };
    let symbol = mk
        .children()
        .find(|n| n.tag_name().name() == "symbol" && is_c_ns(n.tag_name().namespace()))
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string());
    let size = mk
        .children()
        .find(|n| n.tag_name().name() == "size" && is_c_ns(n.tag_name().namespace()))
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<u32>().ok());
    let sp_pr = mk
        .children()
        .find(|n| n.tag_name().name() == "spPr" && is_c_ns(n.tag_name().namespace()));
    let fill = sp_pr.and_then(|p| extract_solid_fill_in_drawingml(&p, theme_colors));
    let line = sp_pr.and_then(|p| {
        let ln = p.children().find(|n| n.tag_name().name() == "ln");
        ln.and_then(|l| extract_solid_fill_in_drawingml(&l, theme_colors))
    });
    (symbol, size, fill, line)
}

/// Locate the first resolvable `<a:solidFill>` among `parent`'s direct children
/// (children only, not deep descendants — chart spPr is structured shallowly)
/// and resolve its color to hex **without** `#` (uppercase). The chart wire
/// model prepends `#` on the TS side, so this matches every other chart color
/// field.
///
/// Delegates the DrawingML color grammar (`srgbClr`/`sysClr`/`prstClr`/
/// `schemeClr` + `lumMod`/`lumOff`/`tint`/`shade`/`alpha` transforms) to the
/// shared [`ooxml_common::color::parse_color_node`] via the crate-wide
/// [`XlsxSchemeResolver`], so scheme slots resolve through the §20.1.6.2 default
/// clrMap (`tx2`→`dk2`, `bg2`→`lt2`) and luminance transforms apply in HLS space
/// (§20.1.2.3.20/.21). The prior private copy in this module mapped `bg2`/`tx2`
/// to the wrong slots and multiplied `lumMod`/`lumOff` in RGB space.
pub(crate) fn extract_solid_fill_in_drawingml(
    parent: &roxmltree::Node,
    theme_colors: &[String],
) -> Option<String> {
    parent
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "solidFill")
        .find_map(|fill| {
            ooxml_common::color::parse_color_node(
                fill,
                &crate::drawing::XlsxSchemeResolver { theme_colors },
                ooxml_common::color::TintMode::PowerPointLinear,
            )
        })
}

/// Walk every `<c:dPt>` direct child of the series and collect per-point
/// overrides. Multiple `<c:dPt>` per series is normal; each one targets a
/// single `<c:idx>` (ECMA-376 §21.2.2.39).
pub(crate) fn parse_data_point_overrides(
    ser_node: &roxmltree::Node,
    theme_colors: &[String],
) -> Vec<DataPointOverride> {
    let mut result = Vec::new();
    for dpt in ser_node.children().filter(|n| {
        n.is_element() && n.tag_name().name() == "dPt" && is_c_ns(n.tag_name().namespace())
    }) {
        let idx = dpt
            .children()
            .find(|n| n.tag_name().name() == "idx" && is_c_ns(n.tag_name().namespace()))
            .and_then(|n| n.attribute("val"))
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(0);
        let sp_pr = dpt
            .children()
            .find(|n| n.tag_name().name() == "spPr" && is_c_ns(n.tag_name().namespace()));
        let color = sp_pr.and_then(|p| extract_solid_fill_in_drawingml(&p, theme_colors));
        let mk = dpt
            .children()
            .find(|n| n.tag_name().name() == "marker" && is_c_ns(n.tag_name().namespace()));
        let (marker_symbol, marker_size, marker_fill, marker_line) =
            parse_marker_block(mk, theme_colors);
        result.push(DataPointOverride {
            idx,
            color,
            marker_symbol,
            marker_size,
            marker_fill,
            marker_line,
        });
    }
    result
}

/// Resolve `<c:ser><c:extLst><c:ext><c15:datalabelsRange>` cache: index →
/// label text. Used to substitute `<a:fld type="CELLRANGE">` placeholders.
/// Returns indices in 0..ptCount; missing entries are empty strings.
pub(crate) fn collect_dlbl_range_cache(ser_node: &roxmltree::Node) -> HashMap<u32, String> {
    let mut map: HashMap<u32, String> = HashMap::new();
    let Some(ext_lst) = ser_node
        .children()
        .find(|n| n.tag_name().name() == "extLst" && is_c_ns(n.tag_name().namespace()))
    else {
        return map;
    };
    for ext in ext_lst.children().filter(|n| {
        n.is_element() && n.tag_name().name() == "ext" && is_c_ns(n.tag_name().namespace())
    }) {
        for range in ext
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "datalabelsRange")
        {
            for cache in range
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "dlblRangeCache")
            {
                for pt in cache.children().filter(|n| {
                    n.is_element()
                        && n.tag_name().name() == "pt"
                        && is_c_ns(n.tag_name().namespace())
                }) {
                    let Some(idx) = pt.attribute("idx").and_then(|v| v.parse::<u32>().ok()) else {
                        continue;
                    };
                    let v = pt
                        .children()
                        .find(|n| n.tag_name().name() == "v" && is_c_ns(n.tag_name().namespace()))
                        .and_then(|n| n.text())
                        .unwrap_or("")
                        .to_string();
                    map.insert(idx, v);
                }
            }
        }
    }
    map
}

/// Walk a `<c:tx><c:rich>` (or any DrawingML rich-text root) and reduce it
/// to plain text. `<a:fld type="CELLRANGE">` placeholders are substituted
/// from `cellrange_cache` keyed by `idx`. Other field types and runs are
/// concatenated. Newlines come from paragraph breaks.
pub(crate) fn flatten_rich_text(
    rich_root: &roxmltree::Node,
    cellrange_cache: Option<&str>,
) -> String {
    let mut out = String::new();
    let mut first_para = true;
    for p in rich_root
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "p")
    {
        if !first_para {
            out.push('\n');
        }
        first_para = false;
        for child in p.children().filter(|n| n.is_element()) {
            match child.tag_name().name() {
                "r" => {
                    if let Some(t) = child.children().find(|n| n.tag_name().name() == "t") {
                        if let Some(s) = t.text() {
                            out.push_str(s);
                        }
                    }
                }
                "fld" => {
                    let typ = child.attribute("type").unwrap_or("");
                    if typ == "CELLRANGE" {
                        if let Some(s) = cellrange_cache {
                            out.push_str(s);
                        }
                    } else {
                        // VALUE/SERIESNAME/CATEGORYNAME field placeholders are
                        // resolved by the renderer using the series data, since
                        // they don't need cell-range expansion. We embed a marker
                        // so the renderer can recognise them.
                        if let Some(t) = child.children().find(|n| n.tag_name().name() == "t") {
                            if let Some(s) = t.text() {
                                out.push_str(s);
                            }
                        }
                    }
                }
                _ => {}
            }
        }
    }
    out
}

/// Parse `<c:dLbls>` (series-level defaults + per-idx overrides).
pub(crate) fn parse_data_labels(
    ser_node: &roxmltree::Node,
    theme_colors: &[String],
    cellrange_cache: &HashMap<u32, String>,
) -> (Option<SeriesDataLabels>, Vec<DataLabelOverride>) {
    let Some(d_lbls) = ser_node
        .children()
        .find(|n| n.tag_name().name() == "dLbls" && is_c_ns(n.tag_name().namespace()))
    else {
        return (None, Vec::new());
    };

    let bool_attr = |n: &roxmltree::Node, name: &str| {
        n.children()
            .find(|c| c.tag_name().name() == name && is_c_ns(c.tag_name().namespace()))
            .and_then(|c| c.attribute("val"))
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
            .unwrap_or(false)
    };

    let position = d_lbls
        .children()
        .find(|n| n.tag_name().name() == "dLblPos" && is_c_ns(n.tag_name().namespace()))
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string());
    let format_code = d_lbls
        .children()
        .find(|n| n.tag_name().name() == "numFmt" && is_c_ns(n.tag_name().namespace()))
        .and_then(|n| n.attribute("formatCode"))
        .map(|s| s.to_string());
    let font_color = d_lbls
        .children()
        .find(|n| n.tag_name().name() == "txPr" && is_c_ns(n.tag_name().namespace()))
        .and_then(|tx| {
            tx.descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "defRPr")
                .and_then(|def| extract_solid_fill_in_drawingml(&def, theme_colors))
        });
    let font_bold_default = d_lbls
        .children()
        .find(|n| n.tag_name().name() == "txPr" && is_c_ns(n.tag_name().namespace()))
        .and_then(|tx| {
            tx.descendants()
                .find(|n| {
                    n.is_element()
                        && (n.tag_name().name() == "defRPr" || n.tag_name().name() == "rPr")
                })
                .and_then(|n| n.attribute("b"))
                .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
        });
    let font_size_default = d_lbls
        .children()
        .find(|n| n.tag_name().name() == "txPr" && is_c_ns(n.tag_name().namespace()))
        .and_then(|tx| {
            tx.descendants()
                .find(|n| {
                    n.is_element()
                        && (n.tag_name().name() == "defRPr" || n.tag_name().name() == "rPr")
                })
                .and_then(|n| n.attribute("sz"))
                .and_then(|v| v.parse::<i32>().ok())
        });

    let series_defaults = SeriesDataLabels {
        show_val: bool_attr(&d_lbls, "showVal"),
        show_cat_name: bool_attr(&d_lbls, "showCatName"),
        show_ser_name: bool_attr(&d_lbls, "showSerName"),
        show_percent: bool_attr(&d_lbls, "showPercent"),
        position: position.clone(),
        font_color: font_color.clone(),
        format_code,
        font_bold: font_bold_default,
        font_size_hpt: font_size_default,
    };

    let mut overrides = Vec::new();
    for dl in d_lbls.children().filter(|n| {
        n.is_element() && n.tag_name().name() == "dLbl" && is_c_ns(n.tag_name().namespace())
    }) {
        let idx = dl
            .children()
            .find(|n| n.tag_name().name() == "idx" && is_c_ns(n.tag_name().namespace()))
            .and_then(|n| n.attribute("val"))
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(0);
        // <c:delete val="1"/> — the user explicitly removed this point's
        // label. Render as empty text so the renderer skips it.
        let deleted = dl
            .children()
            .find(|n| n.tag_name().name() == "delete" && is_c_ns(n.tag_name().namespace()))
            .and_then(|n| n.attribute("val"))
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
            .unwrap_or(false);
        let pos = dl
            .children()
            .find(|n| n.tag_name().name() == "dLblPos" && is_c_ns(n.tag_name().namespace()))
            .and_then(|n| n.attribute("val"))
            .map(|s| s.to_string());
        let cache_for_idx = cellrange_cache.get(&idx).map(|s| s.as_str());
        let text = if deleted {
            String::new()
        } else {
            // Custom text lives at `<c:tx><c:rich>` (ECMA-376 §21.2.2.46).
            // Without `<c:tx>` the override is metadata-only (e.g. only a
            // position change); show the cellrange cache value when
            // available, else empty.
            let tx = dl
                .children()
                .find(|n| n.tag_name().name() == "tx" && is_c_ns(n.tag_name().namespace()));
            match tx {
                Some(tx_node) => flatten_rich_text(&tx_node, cache_for_idx),
                None => cache_for_idx.unwrap_or("").to_string(),
            }
        };
        let font_color = dl
            .children()
            .find(|n| n.tag_name().name() == "txPr" && is_c_ns(n.tag_name().namespace()))
            .and_then(|tx| {
                tx.descendants()
                    .find(|n| n.is_element() && n.tag_name().name() == "defRPr")
                    .and_then(|def| extract_solid_fill_in_drawingml(&def, theme_colors))
            });
        let font_size_hpt = dl
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "defRPr")
            .and_then(|def| def.attribute("sz"))
            .and_then(|v| v.parse::<i32>().ok());
        let font_bold = dl
            .descendants()
            .find(|n| {
                n.is_element() && (n.tag_name().name() == "defRPr" || n.tag_name().name() == "rPr")
            })
            .and_then(|def| def.attribute("b"))
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"));
        overrides.push(DataLabelOverride {
            idx,
            text,
            position: pos,
            font_color,
            font_size_hpt,
            font_bold,
        });
    }

    let any_default = series_defaults.show_val
        || series_defaults.show_cat_name
        || series_defaults.show_ser_name
        || series_defaults.show_percent
        || series_defaults.position.is_some()
        || series_defaults.font_color.is_some()
        || series_defaults.format_code.is_some()
        || series_defaults.font_bold.is_some()
        || series_defaults.font_size_hpt.is_some();
    let series_out = if any_default {
        Some(series_defaults)
    } else {
        None
    };
    (series_out, overrides)
}

/// Parse all `<c:errBars>` direct children of a series and resolve per-
/// point plus / minus deltas to absolute numbers. Each errBars block fixes
/// a direction (x|y); a series can have at most one of each direction.
pub(crate) fn parse_error_bars(
    ser_node: &roxmltree::Node,
    series_values: &[Option<f64>],
    theme_colors: &[String],
) -> Vec<ErrBars> {
    let mut result = Vec::new();
    for eb in ser_node.children().filter(|n| {
        n.is_element() && n.tag_name().name() == "errBars" && is_c_ns(n.tag_name().namespace())
    }) {
        let dir = eb
            .children()
            .find(|n| n.tag_name().name() == "errDir" && is_c_ns(n.tag_name().namespace()))
            .and_then(|n| n.attribute("val"))
            .unwrap_or("y")
            .to_string();
        let bar_type = eb
            .children()
            .find(|n| n.tag_name().name() == "errBarType" && is_c_ns(n.tag_name().namespace()))
            .and_then(|n| n.attribute("val"))
            .unwrap_or("both")
            .to_string();
        let val_type = eb
            .children()
            .find(|n| n.tag_name().name() == "errValType" && is_c_ns(n.tag_name().namespace()))
            .and_then(|n| n.attribute("val"))
            .unwrap_or("fixedVal")
            .to_string();
        let no_end_cap = eb
            .children()
            .find(|n| n.tag_name().name() == "noEndCap" && is_c_ns(n.tag_name().namespace()))
            .and_then(|n| n.attribute("val"))
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
            .unwrap_or(false);

        let n_points = series_values.len();
        let mut plus: Vec<Option<f64>> = vec![None; n_points];
        let mut minus: Vec<Option<f64>> = vec![None; n_points];

        match val_type.as_str() {
            "cust" => {
                for (slot, target) in [("plus", &mut plus), ("minus", &mut minus)] {
                    let Some(side) = eb
                        .children()
                        .find(|n| n.tag_name().name() == slot && is_c_ns(n.tag_name().namespace()))
                    else {
                        continue;
                    };
                    let vals = extract_num_block(&side, n_points);
                    if !vals.is_empty() {
                        let len = vals.len().min(target.len());
                        target[..len].copy_from_slice(&vals[..len]);
                    }
                }
            }
            "fixedVal" => {
                let v = eb
                    .children()
                    .find(|n| n.tag_name().name() == "val" && is_c_ns(n.tag_name().namespace()))
                    .and_then(|n| n.attribute("val"))
                    .and_then(|s| s.parse::<f64>().ok())
                    .unwrap_or(0.0);
                for i in 0..n_points {
                    plus[i] = Some(v);
                    minus[i] = Some(v);
                }
            }
            "percentage" => {
                // Each point's bar = abs(value) * pct/100.
                let pct = eb
                    .children()
                    .find(|n| n.tag_name().name() == "val" && is_c_ns(n.tag_name().namespace()))
                    .and_then(|n| n.attribute("val"))
                    .and_then(|s| s.parse::<f64>().ok())
                    .unwrap_or(0.0);
                for (i, v) in series_values.iter().enumerate() {
                    if let Some(val) = v {
                        let d = val.abs() * pct / 100.0;
                        plus[i] = Some(d);
                        minus[i] = Some(d);
                    }
                }
            }
            "stdErr" | "stdDev" => {
                let nums: Vec<f64> = series_values.iter().filter_map(|v| *v).collect();
                if !nums.is_empty() {
                    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
                    let var =
                        nums.iter().map(|v| (v - mean).powi(2)).sum::<f64>() / nums.len() as f64;
                    let std = var.sqrt();
                    let mult = eb
                        .children()
                        .find(|n| n.tag_name().name() == "val" && is_c_ns(n.tag_name().namespace()))
                        .and_then(|n| n.attribute("val"))
                        .and_then(|s| s.parse::<f64>().ok())
                        .unwrap_or(1.0);
                    let sample = if val_type == "stdErr" {
                        std / (nums.len() as f64).sqrt()
                    } else {
                        std
                    };
                    let delta = sample * mult;
                    for i in 0..n_points {
                        plus[i] = Some(delta);
                        minus[i] = Some(delta);
                    }
                }
            }
            _ => {}
        }

        let sp_pr = eb
            .children()
            .find(|n| n.tag_name().name() == "spPr" && is_c_ns(n.tag_name().namespace()));
        let color = sp_pr.and_then(|p| {
            let ln = p.children().find(|n| n.tag_name().name() == "ln");
            match ln {
                Some(l) => extract_solid_fill_in_drawingml(&l, theme_colors),
                None => extract_solid_fill_in_drawingml(&p, theme_colors),
            }
        });
        let line_width_emu = sp_pr
            .and_then(|p| p.children().find(|n| n.tag_name().name() == "ln"))
            .and_then(|ln| ln.attribute("w"))
            .and_then(|v| v.parse::<u32>().ok());
        let dash = sp_pr
            .and_then(|p| p.children().find(|n| n.tag_name().name() == "ln"))
            .and_then(|ln| ln.children().find(|n| n.tag_name().name() == "prstDash"))
            .and_then(|n| n.attribute("val"))
            .map(|s| s.to_string());

        result.push(ErrBars {
            dir,
            bar_type,
            plus,
            minus,
            no_end_cap,
            color,
            line_width_emu,
            dash,
        });
    }
    result
}

/// Read a `<c:numRef><c:numCache>` or `<c:numLit>` block under `parent` and
/// return per-point values keyed by `<c:pt idx>`. Length is at least
/// `expected_len` (padded with None).
pub(crate) fn extract_num_block(parent: &roxmltree::Node, expected_len: usize) -> Vec<Option<f64>> {
    let cache = parent.descendants().find(|n| {
        n.is_element()
            && (n.tag_name().name() == "numCache" || n.tag_name().name() == "numLit")
            && is_c_ns(n.tag_name().namespace())
    });
    let Some(cache) = cache else {
        return Vec::new();
    };
    let pt_count: usize = cache
        .children()
        .find(|n| n.tag_name().name() == "ptCount" && is_c_ns(n.tag_name().namespace()))
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(expected_len);
    let len = pt_count.max(expected_len);
    let mut values: Vec<Option<f64>> = vec![None; len];
    for pt in cache
        .children()
        .filter(|n| n.tag_name().name() == "pt" && is_c_ns(n.tag_name().namespace()))
    {
        let Some(idx) = pt.attribute("idx").and_then(|v| v.parse::<usize>().ok()) else {
            continue;
        };
        let v = pt
            .children()
            .find(|n| n.tag_name().name() == "v" && is_c_ns(n.tag_name().namespace()))
            .and_then(|n| n.text())
            .and_then(|s| s.trim().parse::<f64>().ok());
        if idx < values.len() {
            values[idx] = v;
        }
    }
    values
}

/// Extract series name from `c:tx`.
pub(crate) fn extract_series_name(node: &roxmltree::Node) -> String {
    // c:tx/c:strRef/c:strCache/c:pt[@idx=0]/c:v
    // or c:tx/c:v
    if let Some(tx) = node
        .children()
        .find(|n| n.tag_name().name() == "tx" && is_c_ns(n.tag_name().namespace()))
    {
        for desc in tx.descendants() {
            if desc.tag_name().name() == "v" && is_c_ns(desc.tag_name().namespace()) {
                if let Some(t) = desc.text() {
                    if !t.is_empty() {
                        return t.to_string();
                    }
                }
            }
        }
    }
    String::new()
}

/// Collect string values from a cache child element (e.g. `<c:cat>` or `<c:xVal>`).
/// Reads `c:strRef/c:strCache`, `c:multiLvlStrRef/c:multiLvlStrCache`, or
/// `c:numRef/c:numCache` (formats numbers as strings).
pub(crate) fn collect_str_cache(ser_node: &roxmltree::Node, child_tag: &str) -> Vec<String> {
    let Some(child) = ser_node
        .children()
        .find(|n| n.tag_name().name() == child_tag && is_c_ns(n.tag_name().namespace()))
    else {
        return Vec::new();
    };

    // Multi-level categories: use only the first (innermost) lvl to get primary labels.
    if let Some(multi_cache) = child
        .descendants()
        .find(|n| n.tag_name().name() == "multiLvlStrCache" && is_c_ns(n.tag_name().namespace()))
    {
        let pt_count: usize = multi_cache
            .children()
            .find(|n| n.tag_name().name() == "ptCount" && is_c_ns(n.tag_name().namespace()))
            .and_then(|n| n.attribute("val"))
            .and_then(|v| v.parse().ok())
            .unwrap_or(0);
        if let Some(first_lvl) = multi_cache
            .children()
            .find(|n| n.tag_name().name() == "lvl" && is_c_ns(n.tag_name().namespace()))
        {
            let mut pts: Vec<(usize, String)> = Vec::new();
            for pt in first_lvl.children().filter(|n| {
                n.is_element() && n.tag_name().name() == "pt" && is_c_ns(n.tag_name().namespace())
            }) {
                let idx: usize = pt
                    .attribute("idx")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                let val = pt
                    .children()
                    .find(|n| n.tag_name().name() == "v")
                    .and_then(|n| n.text())
                    .unwrap_or("")
                    .to_string();
                pts.push((idx, val));
            }
            let len = pt_count.max(pts.iter().map(|(i, _)| i + 1).max().unwrap_or(0));
            let mut result = vec![String::new(); len];
            for (idx, val) in pts {
                if idx < result.len() {
                    result[idx] = val;
                }
            }
            return result;
        }
    }

    // Standard strRef/strCache or numRef/numCache
    let mut pt_count: usize = 0;
    let mut pts: Vec<(usize, String)> = Vec::new();
    for desc in child.descendants() {
        match desc.tag_name().name() {
            "ptCount" if is_c_ns(desc.tag_name().namespace()) => {
                pt_count = desc
                    .attribute("val")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
            }
            "pt" if is_c_ns(desc.tag_name().namespace()) => {
                let idx: usize = desc
                    .attribute("idx")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                let val = desc
                    .children()
                    .find(|n| n.tag_name().name() == "v")
                    .and_then(|n| n.text())
                    .unwrap_or("")
                    .to_string();
                pts.push((idx, val));
            }
            _ => {}
        }
    }
    if pt_count == 0 {
        pt_count = pts.len();
    }
    let mut result = vec![String::new(); pt_count];
    for (idx, val) in pts {
        if idx < result.len() {
            result[idx] = val;
        }
    }
    result
}

/// Collect numeric values from a cache child element (e.g. `<c:val>` or `<c:yVal>`).
pub(crate) fn collect_num_cache(ser_node: &roxmltree::Node, child_tag: &str) -> Vec<Option<f64>> {
    let Some(child) = ser_node
        .children()
        .find(|n| n.tag_name().name() == child_tag && is_c_ns(n.tag_name().namespace()))
    else {
        return Vec::new();
    };

    let mut pt_count: usize = 0;
    let mut pts: Vec<(usize, f64)> = Vec::new();
    for desc in child.descendants() {
        match desc.tag_name().name() {
            "ptCount" if is_c_ns(desc.tag_name().namespace()) => {
                pt_count = desc
                    .attribute("val")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
            }
            "pt" if is_c_ns(desc.tag_name().namespace()) => {
                let idx: usize = desc
                    .attribute("idx")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                if let Some(v) = desc
                    .children()
                    .find(|n| n.tag_name().name() == "v")
                    .and_then(|n| n.text())
                    .and_then(|t| t.parse::<f64>().ok())
                {
                    pts.push((idx, v));
                }
            }
            _ => {}
        }
    }
    if pt_count == 0 {
        pt_count = pts.len();
    }
    let mut result: Vec<Option<f64>> = vec![None; pt_count];
    for (idx, val) in pts {
        if idx < result.len() {
            result[idx] = Some(val);
        }
    }
    result
}

#[cfg(test)]
mod pie_doughnut_tests {
    use super::*;

    const C_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    fn theme() -> Vec<String> {
        vec!["111111".into(); 12]
    }

    /// A single-series chartSpace whose plotArea holds `chart_elem`
    /// (`pieChart` or `doughnutChart`). One series, three categories.
    fn pie_chart_xml(chart_elem: &str) -> String {
        format!(
            r#"<c:chartSpace xmlns:c="{c}" xmlns:a="{a}">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>Share</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:layout/>
      <c:{ce}>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Region</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:strCache>
            <c:pt idx="0"><c:v>North</c:v></c:pt>
            <c:pt idx="1"><c:v>South</c:v></c:pt>
            <c:pt idx="2"><c:v>East</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>30</c:v></c:pt>
            <c:pt idx="1"><c:v>50</c:v></c:pt>
            <c:pt idx="2"><c:v>20</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
      </c:{ce}>
    </c:plotArea>
    <c:legend><c:legendPos val="r"/></c:legend>
  </c:chart>
</c:chartSpace>"#,
            c = C_NS,
            a = A_NS,
            ce = chart_elem,
        )
    }

    /// ECMA-376 §21.2.2.141 `<c:pieChart>` → chart_type "pie"; categories and
    /// the single series' values are carried through unchanged.
    #[test]
    fn pie_chart_series_and_categories() {
        let xml = pie_chart_xml("pieChart");
        let chart = parse_chart_xml(&xml, &theme()).expect("pie chart parses");
        assert_eq!(chart.chart_type, "pie");
        assert_eq!(chart.title.as_deref(), Some("Share"));
        assert_eq!(chart.categories, vec!["North", "South", "East"]);
        assert_eq!(chart.series.len(), 1);
        assert_eq!(
            chart.series[0].values,
            vec![Some(30.0), Some(50.0), Some(20.0)]
        );
    }

    /// ECMA-376 §21.2.2.50 `<c:doughnutChart>` → chart_type "doughnut".
    #[test]
    fn doughnut_chart_type() {
        let xml = pie_chart_xml("doughnutChart");
        let chart = parse_chart_xml(&xml, &theme()).expect("doughnut chart parses");
        assert_eq!(chart.chart_type, "doughnut");
        assert_eq!(chart.categories.len(), 3);
        assert_eq!(chart.series.len(), 1);
        assert_eq!(
            chart.series[0].values,
            vec![Some(30.0), Some(50.0), Some(20.0)]
        );
    }

    // ── Oracle for the `ChartData → ChartModel` conversion (was adaptChartData) ──
    //
    // These pin the adapter defaults / conditionals that moved from TS to Rust.
    // The emitted `ChartModel` must carry them exactly so chart rendering is
    // unchanged.

    /// A vertical clustered bar chart, no `<c:chartSpace><c:spPr>` (relies on
    /// Excel's default opaque-white chart area). Exercises: canonical chartType,
    /// white chartBg default, `between` crossBetween, `out` tick defaults,
    /// null plotAreaBg / dataLabelFontSizeHpt, empty subtotalIndices, and the
    /// series adapter (categories→Some, showMarker→Some, seriesType→Some,
    /// `order` dropped).
    fn bar_chart_xml() -> String {
        format!(
            r#"<c:chartSpace xmlns:c="{c}" xmlns:a="{a}">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Sales</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:strCache>
            <c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
      <c:valAx><c:axId val="1"/><c:scaling><c:max val="30"/><c:min val="0"/></c:scaling><c:axPos val="l"/></c:valAx>
      <c:catAx><c:axId val="2"/><c:axPos val="b"/></c:catAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b"/></c:legend>
  </c:chart>
</c:chartSpace>"#,
            c = C_NS,
            a = A_NS,
        )
    }

    #[test]
    fn adapter_bar_defaults_and_series_mapping() {
        let data = parse_chart_xml(&bar_chart_xml(), &theme()).expect("bar parses");
        // `order` exists on the parse-time series but not on the emitted model.
        assert_eq!(data.series[0].order, 0);
        let m = ChartModel::from(data);

        // Canonical chartType (bar + col + clustered).
        assert_eq!(m.chart_type, "clusteredBar");
        // No `<c:spPr>` on chartSpace → Excel default opaque white.
        assert_eq!(m.chart_bg.as_deref(), Some("FFFFFF"));
        // Hard-coded adapter constants.
        assert_eq!(m.cat_axis_cross_between, "between");
        assert!(m.plot_area_bg.is_none());
        assert!(m.data_label_font_size_hpt.is_none());
        assert!(m.subtotal_indices.is_empty());
        // Tick-mark default `out` (adapter `?? 'out'`).
        assert_eq!(m.val_axis_major_tick_mark, "out");
        assert_eq!(m.cat_axis_major_tick_mark, "out");
        // valAxisMin/Max → valMin/Max rename.
        assert_eq!(m.val_min, Some(0.0));
        assert_eq!(m.val_max, Some(30.0));
        assert_eq!(m.legend_pos.as_deref(), Some("b"));

        // Series adapter: categories non-empty → Some; showMarker/seriesType
        // wrapped in Some; the `order` field is gone.
        assert_eq!(m.series.len(), 1);
        let s = &m.series[0];
        assert_eq!(s.name, "Sales");
        assert_eq!(
            s.categories.as_deref(),
            Some(&["Q1".to_string(), "Q2".to_string()][..])
        );
        assert_eq!(s.series_type.as_deref(), Some("bar"));
        assert_eq!(s.show_marker, Some(false));
        assert_eq!(s.values, vec![Some(10.0), Some(20.0)]);
        // xlsx never populates these (pptx chartEx-only).
        assert!(s.data_point_colors.is_none());
        assert!(s.bubble_sizes.is_none());
        assert!(s.use_secondary_axis.is_none());
    }

    /// A horizontal stacked bar → `stackedBarH`, and an explicit
    /// `<c:chartSpace><c:spPr><a:noFill/>` → transparent chartBg (None), NOT the
    /// white default. Proves the `has_chart_sp_pr` branch of the adapter.
    #[test]
    fn adapter_horizontal_stacked_and_nofill_bg() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{c}" xmlns:a="{a}">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="bar"/>
        <c:grouping val="stacked"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>5</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
  <c:spPr><a:noFill/></c:spPr>
</c:chartSpace>"#,
            c = C_NS,
            a = A_NS,
        );
        let m = ChartModel::from(parse_chart_xml(&xml, &theme()).expect("parses"));
        assert_eq!(m.chart_type, "stackedBarH");
        // spPr present but noFill → transparent, so NOT the white default.
        assert!(
            m.chart_bg.is_none(),
            "explicit noFill spPr must yield transparent chartBg, got {:?}",
            m.chart_bg
        );
    }

    /// A series with no `<c:cat>` → empty categories → adapter emits `None`
    /// (not an empty array), matching `categories.length > 0 ? … : null`.
    #[test]
    fn adapter_empty_categories_become_none() {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{c}" xmlns:a="{a}">
  <c:chart><c:plotArea><c:lineChart>
    <c:ser><c:idx val="0"/><c:order val="0"/>
      <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
    </c:ser>
  </c:lineChart></c:plotArea></c:chart>
</c:chartSpace>"#,
            c = C_NS,
            a = A_NS,
        );
        let m = ChartModel::from(parse_chart_xml(&xml, &theme()).expect("parses"));
        assert_eq!(m.chart_type, "line");
        assert!(m.series[0].categories.is_none());
    }

    /// `<c:date1904/>` as a direct child of `<c:chartSpace>` (ECMA-376
    /// §21.2.2.38) must surface as `ChartModel.date1904 = true`, threaded from
    /// `parse_chart_xml` through the `ChartData → ChartModel` adapter. Absence
    /// of the element leaves it at the default 1900 system (false).
    #[test]
    fn adapter_chart_space_date1904_element() {
        let with = format!(
            r#"<c:chartSpace xmlns:c="{c}" xmlns:a="{a}">
  <c:date1904/>
  <c:chart><c:plotArea><c:lineChart>
    <c:ser><c:idx val="0"/><c:order val="0"/>
      <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
    </c:ser>
  </c:lineChart></c:plotArea></c:chart>
</c:chartSpace>"#,
            c = C_NS,
            a = A_NS,
        );
        let m = ChartModel::from(parse_chart_xml(&with, &theme()).expect("parses"));
        assert!(
            m.date1904,
            "<c:date1904/> must set ChartModel.date1904 = true"
        );

        let without = format!(
            r#"<c:chartSpace xmlns:c="{c}" xmlns:a="{a}">
  <c:chart><c:plotArea><c:lineChart>
    <c:ser><c:idx val="0"/><c:order val="0"/>
      <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
    </c:ser>
  </c:lineChart></c:plotArea></c:chart>
</c:chartSpace>"#,
            c = C_NS,
            a = A_NS,
        );
        let m0 = ChartModel::from(parse_chart_xml(&without, &theme()).expect("parses"));
        assert!(
            !m0.date1904,
            "absent <c:date1904> must leave the 1900 default"
        );
    }
}

#[cfg(test)]
mod label_color_tests {
    use super::*;
    use roxmltree::Document;

    const C_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    fn ser_xml(scheme: &str) -> String {
        format!(
            r#"<c:ser xmlns:c="{c}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                 <c:idx val="0"/><c:order val="0"/>
                 <c:dLbls>
                   <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr>
                     <a:solidFill><a:schemeClr val="{s}"/></a:solidFill>
                   </a:defRPr></a:pPr></a:p></c:txPr>
                   <c:showVal val="1"/>
                 </c:dLbls>
                 <c:val><c:numRef><c:numCache>
                   <c:pt idx="0"><c:v>1</c:v></c:pt>
                 </c:numCache></c:numRef></c:val>
               </c:ser>"#,
            c = C_NS,
            s = scheme
        )
    }

    // Theme order: dk1@0, lt1@1, dk2@2, lt2@3, accent1@4 …
    fn theme() -> Vec<String> {
        vec![
            "111111".into(),
            "fefefe".into(),
            "222222".into(),
            "eeeeee".into(),
            "aa0000".into(),
            "00aa00".into(),
            "0000aa".into(),
            "aaaa00".into(),
            "00aaaa".into(),
            "aa00aa".into(),
        ]
    }

    #[test]
    fn per_series_label_color_resolves_tx1_to_dk1() {
        let xml = ser_xml("tx1");
        let doc = Document::parse(&xml).unwrap();
        let s = parse_chart_series(&doc.root_element(), "bar", false, &theme());
        // tx1 maps to dk1 → theme[0]
        assert_eq!(s.label_color.as_deref(), Some("111111"));
    }

    #[test]
    fn per_series_label_color_resolves_bg1_to_lt1() {
        let xml = ser_xml("bg1");
        let doc = Document::parse(&xml).unwrap();
        let s = parse_chart_series(&doc.root_element(), "bar", false, &theme());
        // bg1 maps to lt1 → theme[1]
        assert_eq!(s.label_color.as_deref(), Some("fefefe"));
    }

    #[test]
    fn no_series_dlbls_color_is_none() {
        let xml = format!(
            r#"<c:ser xmlns:c="{c}"><c:idx val="0"/><c:order val="0"/>
                 <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
               </c:ser>"#,
            c = C_NS
        );
        let doc = Document::parse(&xml).unwrap();
        let s = parse_chart_series(&doc.root_element(), "bar", false, &theme());
        assert_eq!(s.label_color, None);
    }
}

#[cfg(test)]
mod axis_title_and_border_tests {
    use super::*;

    const C_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    fn theme() -> Vec<String> {
        vec!["111111".into(); 12]
    }

    /// A bar chartSpace whose axis titles + optional chartSpace spPr are
    /// supplied verbatim. `cat_title` / `val_title` are full `<c:title>…`
    /// fragments (or empty), `sp_pr` is a full `<c:spPr>…` fragment (or empty).
    fn bar_chart_xml(cat_title: &str, val_title: &str, sp_pr: &str) -> String {
        format!(
            r#"<c:chartSpace xmlns:c="{c}" xmlns:a="{a}">
  {sp}
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:cat><c:strRef><c:strCache>
            <c:pt idx="0"><c:v>A</c:v></c:pt>
            <c:pt idx="1"><c:v>B</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>3</c:v></c:pt>
            <c:pt idx="1"><c:v>5</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
      <c:catAx>
        <c:axPos val="b"/>
        {cat}
      </c:catAx>
      <c:valAx>
        <c:axPos val="l"/>
        {val}
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#,
            c = C_NS,
            a = A_NS,
            sp = sp_pr,
            cat = cat_title,
            val = val_title,
        )
    }

    /// Fixture A — a `valAx` title with `sz="1800" b="1"` → size Some(1800),
    /// bold Some(true). The cat axis has no title.
    #[test]
    fn fixture_a_val_axis_title_size_and_bold() {
        let val_title = r#"<c:title><c:tx><c:rich><a:p><a:pPr>
            <a:defRPr sz="1800" b="1"/>
          </a:pPr><a:r><a:rPr sz="1800" b="1"/><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx></c:title>"#;
        let xml = bar_chart_xml("", val_title, "");
        let chart = parse_chart_xml(&xml, &theme()).expect("chart parses");
        assert_eq!(chart.val_axis_title.as_deref(), Some("Revenue"));
        assert_eq!(chart.val_axis_title_size, Some(1800));
        assert_eq!(chart.val_axis_title_bold, Some(true));
    }

    /// Fixture B — axis title with `sz="1800"` only (no `b`) → bold None.
    #[test]
    fn fixture_b_size_only_leaves_bold_none() {
        let val_title = r#"<c:title><c:tx><c:rich><a:p>
            <a:r><a:rPr sz="1800"/><a:t>Units</a:t></a:r></a:p></c:rich></c:tx></c:title>"#;
        let xml = bar_chart_xml("", val_title, "");
        let chart = parse_chart_xml(&xml, &theme()).expect("chart parses");
        assert_eq!(chart.val_axis_title_size, Some(1800));
        assert_eq!(chart.val_axis_title_bold, None);
    }

    /// Fixture A/cat — confirm the cat-axis path populates its own fields too
    /// (the renderer fix touches both axes uniformly).
    #[test]
    fn cat_axis_title_size_and_color() {
        let cat_title = r#"<c:title><c:tx><c:rich><a:p>
            <a:r><a:rPr sz="1800"><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill></a:rPr><a:t>Month</a:t></a:r></a:p></c:rich></c:tx></c:title>"#;
        let xml = bar_chart_xml(cat_title, "", "");
        let chart = parse_chart_xml(&xml, &theme()).expect("chart parses");
        assert_eq!(chart.cat_axis_title.as_deref(), Some("Month"));
        assert_eq!(chart.cat_axis_title_size, Some(1800));
        assert_eq!(chart.cat_axis_title_color.as_deref(), Some("ff0000"));
    }

    /// Fixture C — `<c:chartSpace>` with no spPr → no border, no spPr flag.
    #[test]
    fn fixture_c_no_sp_pr_no_border() {
        let xml = bar_chart_xml("", "", "");
        let chart = parse_chart_xml(&xml, &theme()).expect("chart parses");
        assert!(!chart.has_chart_sp_pr);
        assert_eq!(chart.chart_border_color, None);
        assert_eq!(chart.chart_border_width_emu, None);
    }

    /// Fixture D — `<a:ln w="9525"><a:solidFill><a:srgbClr val="808080"/>` →
    /// color Some("808080"), width Some(9525).
    #[test]
    fn fixture_d_explicit_ln_solid_fill_border() {
        let sp_pr = r#"<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="808080"/></a:solidFill></a:ln></c:spPr>"#;
        let xml = bar_chart_xml("", "", sp_pr);
        let chart = parse_chart_xml(&xml, &theme()).expect("chart parses");
        assert!(chart.has_chart_sp_pr);
        assert_eq!(chart.chart_border_color.as_deref(), Some("808080"));
        assert_eq!(chart.chart_border_width_emu, Some(9525));
    }

    /// Fixture E — `<a:ln><a:noFill/>` → border explicitly off → color None.
    #[test]
    fn fixture_e_ln_nofill_no_border() {
        let sp_pr = r#"<c:spPr><a:ln w="12700"><a:noFill/></a:ln></c:spPr>"#;
        let xml = bar_chart_xml("", "", sp_pr);
        let chart = parse_chart_xml(&xml, &theme()).expect("chart parses");
        assert!(chart.has_chart_sp_pr);
        assert_eq!(chart.chart_border_color, None);
    }

    /// A scatter chartSpace: two `<c:valAx>` (no `<c:catAx>`); the bottom one
    /// (`axPos="b"`) is the horizontal/X axis. `x_title`/`y_title` are full
    /// `<c:title>…` fragments (or empty).
    fn scatter_chart_xml(x_title: &str, y_title: &str) -> String {
        format!(
            r#"<c:chartSpace xmlns:c="{c}" xmlns:a="{a}">
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <c:scatterChart>
        <c:scatterStyle val="lineMarker"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:xVal><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>1</c:v></c:pt>
            <c:pt idx="1"><c:v>2</c:v></c:pt>
          </c:numCache></c:numRef></c:xVal>
          <c:yVal><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>3</c:v></c:pt>
            <c:pt idx="1"><c:v>5</c:v></c:pt>
          </c:numCache></c:numRef></c:yVal>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:scatterChart>
      <c:valAx>
        <c:axId val="1"/>
        <c:axPos val="b"/>
        {x}
      </c:valAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:axPos val="l"/>
        {y}
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#,
            c = C_NS,
            a = A_NS,
            x = x_title,
            y = y_title,
        )
    }

    /// Regression (sample-30) — a scatter chart's bottom (`axPos="b"`) `<c:valAx>`
    /// title must map to the cat-axis (horizontal) title with its run props.
    /// Before the fix the `is_x_axis` branch never called `extract_chart_title`,
    /// so the X-axis title of every scatter chart silently vanished while the
    /// Y-axis title rendered fine.
    #[test]
    fn scatter_bottom_valax_title_maps_to_cat_axis() {
        let x_title = r#"<c:title><c:tx><c:rich><a:p>
            <a:r><a:rPr sz="1800" b="1"/><a:t>Acid/Bromate</a:t></a:r></a:p></c:rich></c:tx></c:title>"#;
        let y_title = r#"<c:title><c:tx><c:rich><a:p>
            <a:r><a:rPr sz="1800"/><a:t>Period</a:t></a:r></a:p></c:rich></c:tx></c:title>"#;
        let xml = scatter_chart_xml(x_title, y_title);
        let chart = parse_chart_xml(&xml, &theme()).expect("scatter parses");
        // Bottom (X) valAx → cat-axis title, bold 18pt.
        assert_eq!(chart.cat_axis_title.as_deref(), Some("Acid/Bromate"));
        assert_eq!(chart.cat_axis_title_size, Some(1800));
        assert_eq!(chart.cat_axis_title_bold, Some(true));
        // Left (Y) valAx → val-axis title, 18pt not bold.
        assert_eq!(chart.val_axis_title.as_deref(), Some("Period"));
        assert_eq!(chart.val_axis_title_size, Some(1800));
        assert_eq!(chart.val_axis_title_bold, None);
    }
}

#[cfg(test)]
mod solid_fill_color_tests {
    use super::*;
    use roxmltree::Document;

    const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    // Theme in clrScheme document order: dk1@0, lt1@1, dk2@2, lt2@3,
    // accent1@4 … folHlink@11. Distinct hexes so a mis-index is obvious.
    fn theme() -> Vec<String> {
        vec![
            "#111111".into(), // dk1 @0
            "#FEFEFE".into(), // lt1 @1
            "#222222".into(), // dk2 @2
            "#EEEEEE".into(), // lt2 @3
            "#4472C4".into(), // accent1 @4
            "#00AA00".into(), // accent2 @5
            "#0000AA".into(), // accent3 @6
            "#AAAA00".into(), // accent4 @7
            "#00AAAA".into(), // accent5 @8
            "#AA00AA".into(), // accent6 @9
            "#0563C1".into(), // hlink @10
            "#954F72".into(), // folHlink @11
        ]
    }

    fn solid_fill(inner: &str) -> String {
        format!(r#"<a:spPr xmlns:a="{A_NS}"><a:solidFill>{inner}</a:solidFill></a:spPr>"#)
    }

    /// §20.1.6.2 default clrMap: `tx2` → `dk2` (theme slot 2), NOT `lt1`.
    #[test]
    fn scheme_tx2_resolves_to_dk2_slot() {
        let xml = solid_fill(r#"<a:schemeClr val="tx2"/>"#);
        let doc = Document::parse(&xml).unwrap();
        let out = extract_solid_fill_in_drawingml(&doc.root_element(), &theme());
        // tx2 → dk2 → theme[2] = "222222" (uppercase, no `#`).
        assert_eq!(out.as_deref(), Some("222222"));
    }

    /// §20.1.6.2 default clrMap: `bg2` → `lt2` (theme slot 3), NOT `dk1`.
    #[test]
    fn scheme_bg2_resolves_to_lt2_slot() {
        let xml = solid_fill(r#"<a:schemeClr val="bg2"/>"#);
        let doc = Document::parse(&xml).unwrap();
        let out = extract_solid_fill_in_drawingml(&doc.root_element(), &theme());
        // bg2 → lt2 → theme[3] = "EEEEEE".
        assert_eq!(out.as_deref(), Some("EEEEEE"));
    }

    /// tx1 → dk1 and bg1 → lt1 (unchanged, but pinned so a refactor can't drift).
    #[test]
    fn scheme_tx1_bg1_resolve_to_dk1_lt1() {
        let tx1 = solid_fill(r#"<a:schemeClr val="tx1"/>"#);
        let doc = Document::parse(&tx1).unwrap();
        assert_eq!(
            extract_solid_fill_in_drawingml(&doc.root_element(), &theme()).as_deref(),
            Some("111111") // dk1 @0
        );
        let bg1 = solid_fill(r#"<a:schemeClr val="bg1"/>"#);
        let doc = Document::parse(&bg1).unwrap();
        assert_eq!(
            extract_solid_fill_in_drawingml(&doc.root_element(), &theme()).as_deref(),
            Some("FEFEFE") // lt1 @1
        );
    }

    /// `lumMod` is a luminance modulation applied to the HLS `L` channel
    /// (§20.1.2.3.20), NOT a per-RGB-component multiply. For `4472C4` at
    /// `lumMod 50000`, the HLS result is `203864`; the (wrong) RGB-space
    /// multiply would give `223962`.
    #[test]
    fn lummod_applies_in_hls_space_not_rgb() {
        let xml = solid_fill(r#"<a:srgbClr val="4472C4"><a:lumMod val="50000"/></a:srgbClr>"#);
        let doc = Document::parse(&xml).unwrap();
        let out = extract_solid_fill_in_drawingml(&doc.root_element(), &theme());
        assert_eq!(out.as_deref(), Some("203864"));
    }

    /// A plain srgbClr with no transforms passes through (uppercased, no `#`).
    #[test]
    fn plain_srgb_passthrough() {
        let xml = solid_fill(r#"<a:srgbClr val="ff8000"/>"#);
        let doc = Document::parse(&xml).unwrap();
        let out = extract_solid_fill_in_drawingml(&doc.root_element(), &theme());
        assert_eq!(out.as_deref(), Some("FF8000"));
    }
}

#[cfg(test)]
mod date_axis_tests {
    use super::*;

    const C_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    fn theme() -> Vec<String> {
        vec!["#111111".into(); 12]
    }

    /// A line chart whose horizontal axis is a `<c:dateAx>` (§21.2.2.39), the
    /// time-series category axis Excel emits when the category source is dates.
    /// `axis_inner` is spliced into the `<c:dateAx>` element.
    fn date_axis_chart_xml(axis_inner: &str) -> String {
        format!(
            r#"<c:chartSpace xmlns:c="{c}" xmlns:a="{a}">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:cat><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>44927</c:v></c:pt><c:pt idx="1"><c:v>44958</c:v></c:pt>
          </c:numCache></c:numRef></c:cat>
          <c:val><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:lineChart>
      <c:dateAx>
        <c:axId val="2"/>
        <c:axPos val="b"/>
        {axis}
      </c:dateAx>
      <c:valAx><c:axId val="1"/><c:axPos val="l"/></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#,
            c = C_NS,
            a = A_NS,
            axis = axis_inner,
        )
    }

    /// `<c:dateAx>` is a category axis (§21.2.2.39): its `<c:numFmt>` formatCode
    /// must populate `cat_axis_format_code` exactly as `<c:catAx>` does, so the
    /// TS side formats the serial dates instead of showing raw numbers.
    #[test]
    fn date_axis_format_code_populates_cat_axis_format_code() {
        let xml = date_axis_chart_xml(r#"<c:numFmt formatCode="m/d/yyyy" sourceLinked="0"/>"#);
        let chart = parse_chart_xml(&xml, &theme()).expect("dateAx chart parses");
        assert_eq!(chart.cat_axis_format_code.as_deref(), Some("m/d/yyyy"));
    }

    /// A deleted `<c:dateAx>` (`<c:delete val="1"/>`) hides the category axis,
    /// matching the catAx convention.
    #[test]
    fn date_axis_delete_hides_cat_axis() {
        let xml = date_axis_chart_xml(r#"<c:delete val="1"/>"#);
        let chart = parse_chart_xml(&xml, &theme()).expect("dateAx chart parses");
        assert!(chart.cat_axis_hidden);
    }

    /// A dateAx title maps to the cat-axis title (same as catAx).
    #[test]
    fn date_axis_title_maps_to_cat_axis_title() {
        let title = r#"<c:title><c:tx><c:rich><a:p>
            <a:r><a:rPr sz="1200"/><a:t>Date</a:t></a:r></a:p></c:rich></c:tx></c:title>"#;
        let xml = date_axis_chart_xml(title);
        let chart = parse_chart_xml(&xml, &theme()).expect("dateAx chart parses");
        assert_eq!(chart.cat_axis_title.as_deref(), Some("Date"));
        assert_eq!(chart.cat_axis_title_size, Some(1200));
    }
}

/// ISO/IEC 29500 Strict-conformance fixture. `parse_chart_xml` and its helpers
/// match `c:`/`a:` elements via `is_c_ns`/`is_a_ns` (see `fix(xlsx): accept
/// Strict namespace URIs across the parser`); this pins that a Strict
/// chartSpace — `xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart"` +
/// `xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"` — yields the exact
/// same `ChartData` (series values, categories, axis format code) as the
/// Transitional fixture. Before the `is_c_ns`/`is_a_ns` conversion this
/// document parsed to `None` (the top-level `<c:chart>` lookup itself is
/// namespace-pinned), so this test only passes through the predicate path.
#[cfg(test)]
mod strict_namespace_tests {
    use super::*;

    const C_NS_STRICT: &str = "http://purl.oclc.org/ooxml/drawingml/chart";
    const A_NS_STRICT: &str = "http://purl.oclc.org/ooxml/drawingml/main";

    fn theme() -> Vec<String> {
        vec!["#111111".into(); 12]
    }

    #[test]
    fn strict_chart_series_categories_and_axis_format_code() {
        let xml = format!(
            r##"<c:chartSpace xmlns:c="{c}" xmlns:a="{a}">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>Strict Share</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Sales</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:strCache>
            <c:pt idx="0"><c:v>Q1</c:v></c:pt>
            <c:pt idx="1"><c:v>Q2</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>10</c:v></c:pt>
            <c:pt idx="1"><c:v>20</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:valAx>
        <c:axId val="1"/><c:axPos val="l"/>
        <c:numFmt formatCode="#,##0.00" sourceLinked="0"/>
      </c:valAx>
      <c:catAx><c:axId val="2"/><c:axPos val="b"/></c:catAx>
    </c:plotArea>
    <c:legend><c:legendPos val="r"/></c:legend>
  </c:chart>
</c:chartSpace>"##,
            c = C_NS_STRICT,
            a = A_NS_STRICT,
        );

        let chart = parse_chart_xml(&xml, &theme()).expect("Strict chartSpace must parse");
        assert_eq!(chart.chart_type, "bar");
        assert_eq!(chart.title.as_deref(), Some("Strict Share"));
        assert_eq!(chart.categories, vec!["Q1", "Q2"]);
        assert_eq!(chart.series.len(), 1);
        assert_eq!(chart.series[0].values, vec![Some(10.0), Some(20.0)]);
        assert_eq!(chart.val_axis_format_code.as_deref(), Some("#,##0.00"));
    }
}

/// §20.1.2.2.8 — an `<xdr:cNvPr hidden="1">` graphicFrame is not rendered.
/// `load_sheet_charts` walks `<xdr:graphicFrame>` independently of the shared
/// shape walker in `drawing.rs::collect_shapes`, so it needs its own hidden
/// check — this covers that walk specifically (full sheet → drawing → chart
/// zip round trip, since `load_sheet_charts` reads from the archive).
#[cfg(test)]
mod hidden_tests {
    use super::*;
    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;

    const NS: &str = concat!(
        r#"xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" "#,
        r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
        r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#,
    );

    fn theme() -> Vec<String> {
        vec!["#111111".into(); 12]
    }

    fn minimal_chart_xml() -> String {
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>10</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:valAx><c:axId val="1"/><c:axPos val="l"/></c:valAx>
      <c:catAx><c:axId val="2"/><c:axPos val="b"/></c:catAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#
            .to_string()
    }

    fn drawing_xml(hidden_attr: &str) -> String {
        format!(
            r#"<xdr:wsDr {NS}><xdr:twoCellAnchor>
              <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
              <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>16</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
              <xdr:graphicFrame>
                <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"{hidden}/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
                <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="3000000"/></xdr:xfrm>
                <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                  <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart"/>
                </a:graphicData></a:graphic>
              </xdr:graphicFrame>
              <xdr:clientData/>
            </xdr:twoCellAnchor></xdr:wsDr>"#,
            NS = NS,
            hidden = hidden_attr,
        )
    }

    /// Builds a minimal zip archive wiring `xl/worksheets/sheet1.xml`'s rels →
    /// `xl/drawings/drawing1.xml` → its own rels → `xl/charts/chart1.xml`, the
    /// same part chain `load_sheet_charts` walks in production.
    fn archive_with_chart(hidden_attr: &str) -> crate::XlsxZip {
        let mut buf = Vec::new();
        {
            let mut zw = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = SimpleFileOptions::default();

            zw.start_file("xl/worksheets/_rels/sheet1.xml.rels", o)
                .unwrap();
            zw.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#).unwrap();

            zw.start_file("xl/drawings/drawing1.xml", o).unwrap();
            zw.write_all(drawing_xml(hidden_attr).as_bytes()).unwrap();

            zw.start_file("xl/drawings/_rels/drawing1.xml.rels", o)
                .unwrap();
            zw.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#).unwrap();

            zw.start_file("xl/charts/chart1.xml", o).unwrap();
            zw.write_all(minimal_chart_xml().as_bytes()).unwrap();

            zw.finish().unwrap();
        }
        zip::ZipArchive::new(Cursor::new(buf)).unwrap()
    }

    #[test]
    fn hidden_chart_graphicframe_is_not_emitted() {
        for attr in [r#" hidden="1""#, r#" hidden="true""#] {
            let mut archive = archive_with_chart(attr);
            let charts = load_sheet_charts(&mut archive, "worksheets/sheet1.xml", &theme());
            assert!(charts.is_empty(), "hidden chart emitted (attr={attr})");
        }
    }

    #[test]
    fn visible_chart_graphicframe_is_emitted_unchanged() {
        for attr in ["", r#" hidden="0""#, r#" hidden="false""#] {
            let mut archive = archive_with_chart(attr);
            let charts = load_sheet_charts(&mut archive, "worksheets/sheet1.xml", &theme());
            assert_eq!(charts.len(), 1, "visible chart dropped (attr={attr})");
        }
    }
}
