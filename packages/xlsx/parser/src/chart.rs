use crate::types::*;
use crate::{resolve_fill_color, read_zip_entry, resolve_zip_path, parse_rels_map};
use std::io::Cursor;
use std::collections::HashMap;

/// Given a sheet path (e.g. "worksheets/sheet1.xml"), locate and parse
/// its drawing(s) for chart anchors (`<xdr:graphicFrame>` elements).
pub(crate) fn load_sheet_charts(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    sheet_path: &str,
    theme_colors: &[String],
) -> Vec<ChartAnchor> {
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

    // Collect all drawing relationship targets
    let mut drawing_targets: Vec<String> = Vec::new();
    for rel in rels_doc.root_element().children().filter(|n| n.is_element()) {
        if rel.attribute("Type").unwrap_or("").ends_with("/drawing") {
            if let Some(t) = rel.attribute("Target") {
                drawing_targets.push(t.to_string());
            }
        }
    }
    if drawing_targets.is_empty() { return Vec::new(); }

    let mut all_charts: Vec<ChartAnchor> = Vec::new();
    let xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    let a_ns   = "http://schemas.openxmlformats.org/drawingml/2006/main";
    let c_ns   = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    let r_ns   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    for target in drawing_targets {
        // Resolve drawing path relative to the sheet directory
        let drawing_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        let Ok(drawing_xml) = read_zip_entry(archive, &drawing_path) else { continue; };
        let Ok(draw_doc) = roxmltree::Document::parse(&drawing_xml) else { continue; };

        // Load drawing rels (to resolve chart rIds)
        let Some((drawing_dir, drawing_file)) = drawing_path.rsplit_once('/') else { continue; };
        let drawing_rels_path = format!("{}/_rels/{}.rels", drawing_dir, drawing_file);
        let drawing_rels = read_zip_entry(archive, &drawing_rels_path)
            .ok()
            .map(|xml| parse_rels_map(&xml))
            .unwrap_or_default();

        // Iterate over twoCellAnchor elements
        for anchor in draw_doc.root_element().children().filter(|n| n.is_element()) {
            if anchor.tag_name().name() != "twoCellAnchor"
                || anchor.tag_name().namespace() != Some(xdr_ns)
            {
                continue;
            }

            let (mut from_col, mut from_col_off, mut from_row, mut from_row_off) = (0u32, 0i64, 0u32, 0i64);
            let (mut to_col,   mut to_col_off,   mut to_row,   mut to_row_off)   = (0u32, 0i64, 0u32, 0i64);
            let mut chart_rid: Option<String> = None;

            for child in anchor.children() {
                if !child.is_element() { continue; }
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
                        if is_from { from_col = col; from_col_off = col_off; from_row = row; from_row_off = row_off; }
                        else       { to_col   = col; to_col_off   = col_off; to_row   = row; to_row_off   = row_off; }
                    }
                    "graphicFrame" => {
                        // Look for a:graphic/a:graphicData/c:chart[@r:id]
                        for gf_child in child.descendants() {
                            if gf_child.tag_name().name() == "chart"
                                && gf_child.tag_name().namespace() == Some(c_ns)
                            {
                                if let Some(rid) = gf_child.attributes()
                                    .find(|a| a.name() == "id" && a.namespace() == Some(r_ns))
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

            let Some(rid) = chart_rid else { continue; };
            let Some(chart_target) = drawing_rels.get(&rid) else { continue; };
            let chart_path = resolve_zip_path(drawing_dir, chart_target);
            let Ok(chart_xml) = read_zip_entry(archive, &chart_path) else { continue; };
            let Some(chart_data) = parse_chart_xml(&chart_xml, c_ns, a_ns, theme_colors) else { continue; };

            all_charts.push(ChartAnchor {
                from_col, from_col_off, from_row, from_row_off,
                to_col,   to_col_off,   to_row,   to_row_off,
                chart: chart_data,
            });
        }
    }
    all_charts
}

// ─── Chart XML parser ────────────────────────────────────────────────────────

/// Parse a `xl/charts/chartN.xml` file into a `ChartData`.
pub(crate) fn parse_chart_xml(xml: &str, c_ns: &str, a_ns: &str, theme_colors: &[String]) -> Option<ChartData> {
    let doc = roxmltree::Document::parse(xml).ok()?;

    // Find c:chart root element
    let chart_root = doc.descendants()
        .find(|n| n.tag_name().name() == "chart" && n.tag_name().namespace() == Some(c_ns))?;

    // Parse optional title
    let title = extract_chart_title(&chart_root, c_ns, a_ns);
    let title_font_size_hpt = extract_chart_title_size(&chart_root, c_ns, a_ns);
    let title_font_color = extract_chart_title_color(&chart_root, c_ns, a_ns);
    let title_font_face = extract_chart_title_face(&chart_root, c_ns, a_ns);
    let title_font_bold = extract_chart_title_bold(&chart_root, c_ns, a_ns);

    // Legend presence: <c:chart><c:legend> is the authoritative signal. Absence
    // means Excel hides the legend (default for a single-series chart with no
    // explicit legend element). `<c:legendPos val>` picks a side per
    // ECMA-376 §21.2.2.10 — both parts come from the shared ooxml-common helper
    // so pptx & xlsx stay in lockstep.
    let (show_legend, legend_pos) = ooxml_common::chart::extract_legend(chart_root);
    let legend_node = chart_root.children()
        .find(|n| n.tag_name().name() == "legend" && n.tag_name().namespace() == Some(c_ns));
    // Legend <c:layout><c:manualLayout> (ECMA-376 §21.2.2.31) — when present,
    // gives explicit x/y/w/h fractions of the chart space. Used by the Excel
    // templates that position a top legend into a narrow band, e.g. over the
    // left half of the chart. We just collect the raw fractions here; the
    // renderer decides whether to honor `edge` vs `factor` placement.
    let legend_manual_layout = legend_node.and_then(|ln| {
        let layout = ln.children()
            .find(|n| n.tag_name().name() == "layout" && n.tag_name().namespace() == Some(c_ns))?;
        let manual = layout.children()
            .find(|n| n.tag_name().name() == "manualLayout" && n.tag_name().namespace() == Some(c_ns))?;
        let val = |tag: &str| manual.children()
            .find(|n| n.tag_name().name() == tag && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.attribute("val").and_then(|v| v.parse::<f64>().ok()));
        let mode = |tag: &str| manual.children()
            .find(|n| n.tag_name().name() == tag && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.attribute("val").map(|v| v.to_string()))
            .unwrap_or_else(|| "edge".to_string());
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
    let chart_space_root = doc.descendants()
        .find(|n| n.tag_name().name() == "chartSpace" && n.tag_name().namespace() == Some(c_ns));
    let chart_sp_pr = chart_space_root.and_then(|cs| cs.children()
        .find(|n| n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(c_ns)));
    let has_chart_sp_pr = chart_sp_pr.is_some();
    let chart_bg = chart_sp_pr.and_then(|sp| {
        // Walk direct children: noFill → None, solidFill → resolved color.
        let mut resolved: Option<String> = None;
        for ch in sp.children().filter(|n| n.is_element()) {
            match ch.tag_name().name() {
                "noFill"    => { return None; }
                "solidFill" => { resolved = resolve_fill_color(&ch, theme_colors); break; }
                _ => {}
            }
        }
        resolved
    });

    // `<c:title><c:layout><c:manualLayout>` (ECMA-376 §21.2.2.27).
    let title_manual_layout = chart_root.children()
        .find(|n| n.tag_name().name() == "title" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|t| t.children().find(|n| n.tag_name().name() == "layout" && n.tag_name().namespace() == Some(c_ns)))
        .and_then(|l| extract_manual_layout(&l, c_ns));

    // Find c:plotArea
    let plot_area = chart_root.children()
        .find(|n| n.tag_name().name() == "plotArea" && n.tag_name().namespace() == Some(c_ns))?;
    let plot_area_manual_layout = plot_area.children()
        .find(|n| n.tag_name().name() == "layout" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|l| extract_manual_layout(&l, c_ns));

    let mut primary_type = String::new();
    let mut bar_dir      = "col".to_string();
    let mut grouping     = "clustered".to_string();
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
        ("barChart",      "bar"),
        ("lineChart",     "line"),
        ("areaChart",     "area"),
        ("pieChart",      "pie"),
        ("doughnutChart", "doughnut"),
        ("radarChart",    "radar"),
        ("scatterChart",  "scatter"),
        ("bubbleChart",   "scatter"), // treat bubble as scatter
    ];

    for child in plot_area.children() {
        if !child.is_element() { continue; }
        if child.tag_name().namespace() != Some(c_ns) { continue; }
        let elem_name = child.tag_name().name();

        // Axis title + tick label font size extraction (ECMA-376 §21.2.2.17
        // c:txPr/a:defRPr@sz gives tick labels their hpt size; absent = default).
        match elem_name {
            "catAx" => {
                if cat_axis_title.is_none() {
                    cat_axis_title = extract_chart_title(&child, c_ns, a_ns);
                }
                if cat_axis_font_size_hpt.is_none() {
                    cat_axis_font_size_hpt = extract_axis_tick_label_size(&child, c_ns, a_ns);
                }
                if cat_axis_format_code.is_none() {
                    cat_axis_format_code = extract_axis_format_code(&child, c_ns);
                }
                if cat_axis_font_bold.is_none() {
                    cat_axis_font_bold = extract_axis_tick_label_bold(&child, c_ns);
                }
                if cat_axis_line_color.is_none() || cat_axis_line_width_emu.is_none() {
                    let (col, w) = extract_axis_line_style(&child, c_ns, theme_colors);
                    if cat_axis_line_color.is_none() { cat_axis_line_color = col; }
                    if cat_axis_line_width_emu.is_none() { cat_axis_line_width_emu = w; }
                }
                if cat_axis_major_tick_mark.is_none() {
                    cat_axis_major_tick_mark = extract_axis_tick_mark(&child, c_ns, "majorTickMark");
                }
                if cat_axis_minor_tick_mark.is_none() {
                    cat_axis_minor_tick_mark = extract_axis_tick_mark(&child, c_ns, "minorTickMark");
                }
                if let Some((mn, mx)) = extract_axis_scaling(&child, c_ns) {
                    if cat_axis_min.is_none() { cat_axis_min = mn; }
                    if cat_axis_max.is_none() { cat_axis_max = mx; }
                }
                let (cr, cra) = extract_axis_crosses(&child, c_ns);
                if cat_axis_crosses.is_none() { cat_axis_crosses = cr; }
                if cat_axis_crosses_at.is_none() { cat_axis_crosses_at = cra; }
                if axis_is_deleted(&child, c_ns) { cat_axis_hidden = true; }
                if axis_line_is_hidden(&child, c_ns) { cat_axis_line_hidden = true; }
                continue;
            }
            "valAx" => {
                // Scatter charts use two `<c:valAx>` (no catAx). Disambiguate
                // by `<c:axPos val>` — `b`(bottom)/`t`(top) → X axis, `l`/`r`
                // → Y axis. For non-scatter charts the first valAx hit is
                // always Y.
                let ax_pos = child.children()
                    .find(|n| n.tag_name().name() == "axPos" && n.tag_name().namespace() == Some(c_ns))
                    .and_then(|n| n.attribute("val"))
                    .unwrap_or("");
                let is_x_axis = matches!(ax_pos, "b" | "t");
                if is_x_axis {
                    if cat_axis_format_code.is_none() {
                        cat_axis_format_code = extract_axis_format_code(&child, c_ns);
                    }
                    if let Some((mn, mx)) = extract_axis_scaling(&child, c_ns) {
                        if cat_axis_min.is_none() { cat_axis_min = mn; }
                        if cat_axis_max.is_none() { cat_axis_max = mx; }
                    }
                    if cat_axis_font_size_hpt.is_none() {
                        cat_axis_font_size_hpt = extract_axis_tick_label_size(&child, c_ns, a_ns);
                    }
                    if cat_axis_font_bold.is_none() {
                        cat_axis_font_bold = extract_axis_tick_label_bold(&child, c_ns);
                    }
                    if cat_axis_line_color.is_none() || cat_axis_line_width_emu.is_none() {
                        let (col, w) = extract_axis_line_style(&child, c_ns, theme_colors);
                        if cat_axis_line_color.is_none() { cat_axis_line_color = col; }
                        if cat_axis_line_width_emu.is_none() { cat_axis_line_width_emu = w; }
                    }
                    if cat_axis_major_tick_mark.is_none() {
                        cat_axis_major_tick_mark = extract_axis_tick_mark(&child, c_ns, "majorTickMark");
                    }
                    if cat_axis_minor_tick_mark.is_none() {
                        cat_axis_minor_tick_mark = extract_axis_tick_mark(&child, c_ns, "minorTickMark");
                    }
                    let (cr, cra) = extract_axis_crosses(&child, c_ns);
                    if cat_axis_crosses.is_none() { cat_axis_crosses = cr; }
                    if cat_axis_crosses_at.is_none() { cat_axis_crosses_at = cra; }
                    if axis_is_deleted(&child, c_ns) { cat_axis_hidden = true; }
                    if axis_line_is_hidden(&child, c_ns) { cat_axis_line_hidden = true; }
                } else {
                    if val_axis_title.is_none() {
                        val_axis_title = extract_chart_title(&child, c_ns, a_ns);
                    }
                    if val_axis_font_size_hpt.is_none() {
                        val_axis_font_size_hpt = extract_axis_tick_label_size(&child, c_ns, a_ns);
                    }
                    if val_axis_format_code.is_none() {
                        val_axis_format_code = extract_axis_format_code(&child, c_ns);
                    }
                    if val_axis_font_bold.is_none() {
                        val_axis_font_bold = extract_axis_tick_label_bold(&child, c_ns);
                    }
                    if let Some((mn, mx)) = extract_axis_scaling(&child, c_ns) {
                        if val_axis_min.is_none() { val_axis_min = mn; }
                        if val_axis_max.is_none() { val_axis_max = mx; }
                    }
                    let (cr, cra) = extract_axis_crosses(&child, c_ns);
                    if val_axis_crosses.is_none() { val_axis_crosses = cr; }
                    if val_axis_crosses_at.is_none() { val_axis_crosses_at = cra; }
                    if val_axis_line_color.is_none() || val_axis_line_width_emu.is_none() {
                        let (col, w) = extract_axis_line_style(&child, c_ns, theme_colors);
                        if val_axis_line_color.is_none() { val_axis_line_color = col; }
                        if val_axis_line_width_emu.is_none() { val_axis_line_width_emu = w; }
                    }
                    if val_axis_major_tick_mark.is_none() {
                        val_axis_major_tick_mark = extract_axis_tick_mark(&child, c_ns, "majorTickMark");
                    }
                    if val_axis_minor_tick_mark.is_none() {
                        val_axis_minor_tick_mark = extract_axis_tick_mark(&child, c_ns, "minorTickMark");
                    }
                    if axis_is_deleted(&child, c_ns) { val_axis_hidden = true; }
                    if axis_line_is_hidden(&child, c_ns) { val_axis_line_hidden = true; }
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
                "barDir"   => { bar_dir  = attr_node.attribute("val").unwrap_or("col").to_string(); }
                "grouping" => {
                    let val = attr_node.attribute("val").unwrap_or("clustered").to_string();
                    if !grouping_locked && ser_type != "line" {
                        grouping = val;
                        grouping_locked = true;
                    }
                }
                "marker"   => {
                    chart_marker_default = attr_node.attribute("val").unwrap_or("0") != "0";
                }
                "gapWidth" => {
                    // ECMA-376 §21.2.2.13 — percent of bar width between category
                    // groups (default 150 per spec). Only meaningful on bar charts.
                    if bar_gap_width.is_none() {
                        bar_gap_width = attr_node.attribute("val").and_then(|v| v.parse().ok());
                    }
                }
                "overlap"  => {
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
                "dLbls"    => {
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
                            "numFmt" => {
                                if data_label_format_code.is_none() {
                                    data_label_format_code = d.attribute("formatCode")
                                        .map(|s| s.to_string())
                                        .filter(|s| !s.is_empty() && s != "General");
                                }
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
        for ser_node in child.children()
            .filter(|n| n.is_element() && n.tag_name().name() == "ser" && n.tag_name().namespace() == Some(c_ns))
        {
            let s = parse_chart_series(&ser_node, c_ns, ser_type, chart_marker_default, theme_colors);
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
            if let Some(ser_dlbls) = ser_node.children()
                .find(|n| n.is_element()
                    && n.tag_name().name() == "dLbls"
                    && n.tag_name().namespace() == Some(c_ns))
            {
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
                        "numFmt" => {
                            if data_label_format_code.is_none() {
                                data_label_format_code = d.attribute("formatCode")
                                    .map(|s| s.to_string())
                                    .filter(|s| !s.is_empty() && s != "General");
                            }
                        }
                        // txPr (data-label text color) handled by the
                        // shared helper before this loop runs.
                        _ => {}
                    }
                }
            }
        }
    }

    if primary_type.is_empty() { return None; }

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
    })
}

/// `<c:catAx|valAx><c:numFmt@formatCode>` (ECMA-376 §21.2.2.21). Thin
/// wrapper around `ooxml_common::chart::extract_axis_format_code` so the
/// pptx and xlsx parsers stay in lockstep on the format-code rules.
pub(crate) fn extract_axis_format_code(axis_node: &roxmltree::Node, _c_ns: &str) -> Option<String> {
    ooxml_common::chart::extract_axis_format_code(*axis_node)
}

/// `<c:catAx|valAx><c:majorTickMark val>` / `<c:minorTickMark val>` —
/// `none` / `out` / `in` / `cross` (ECMA-376 §21.2.2.49 ST_TickMark).
pub(crate) fn extract_axis_tick_mark(axis_node: &roxmltree::Node, _c_ns: &str, name: &str) -> Option<String> {
    ooxml_common::chart::extract_axis_tick_mark(*axis_node, name)
}

/// `<c:catAx|valAx><c:spPr><a:ln>` — resolved color (no `#`) and width
/// (EMU) for the axis line itself. None when not set.
pub(crate) fn extract_axis_line_style(
    axis_node: &roxmltree::Node,
    c_ns: &str,
    theme_colors: &[String],
) -> (Option<String>, Option<u32>) {
    let Some(sp_pr) = axis_node.children()
        .find(|n| n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(c_ns))
    else { return (None, None); };
    let Some(ln) = sp_pr.children().find(|n| n.tag_name().name() == "ln") else { return (None, None); };
    let width = ln.attribute("w").and_then(|v| v.parse::<u32>().ok());
    let color = extract_solid_fill_in_drawingml(&ln, theme_colors);
    (color, width)
}

/// `<c:catAx|valAx><c:spPr><a:ln><a:noFill>` — true when the axis line is
/// explicitly hidden (labels and tick marks still render). Distinct from
/// `<c:delete val="1"/>` which hides the entire axis. Sample-1 "Carbon &
/// Growth" uses this on `<c:valAx>` to keep the Y-axis numbers visible
/// while suppressing the vertical rule.
pub(crate) fn axis_line_is_hidden(axis_node: &roxmltree::Node, c_ns: &str) -> bool {
    let Some(sp_pr) = axis_node.children()
        .find(|n| n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(c_ns))
    else { return false; };
    let Some(ln) = sp_pr.children().find(|n| n.tag_name().name() == "ln") else { return false; };
    ln.children().any(|n| n.is_element() && n.tag_name().name() == "noFill")
}

/// `<c:catAx|valAx><c:txPr>...defRPr@b>` — bold flag for axis tick labels.
pub(crate) fn extract_axis_tick_label_bold(axis_node: &roxmltree::Node, c_ns: &str) -> Option<bool> {
    let txpr = axis_node.children()
        .find(|n| n.tag_name().name() == "txPr" && n.tag_name().namespace() == Some(c_ns))?;
    txpr.descendants().find_map(|n| {
        if !n.is_element() { return None; }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" { return None; }
        n.attribute("b").map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
    })
}

/// `<c:catAx|valAx><c:crosses>` and `<c:crossesAt>` — where the axis sits
/// along its perpendicular axis. `crosses` is a string ("autoZero" |
/// "min" | "max"); `crossesAt` is an explicit numeric override that
/// takes precedence at render time.
pub(crate) fn extract_axis_crosses(axis_node: &roxmltree::Node, c_ns: &str) -> (Option<String>, Option<f64>) {
    let crosses = axis_node.children()
        .find(|n| n.tag_name().name() == "crosses" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string());
    let crosses_at = axis_node.children()
        .find(|n| n.tag_name().name() == "crossesAt" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok());
    (crosses, crosses_at)
}

/// Read explicit `<c:scaling><c:min>` / `<c:scaling><c:max>` values, returning
/// `(min, max)` where each is `None` if the axis didn't override that bound.
/// Returns `None` only when neither bound is set (matches the prior xlsx
/// callsite shape `if let Some((mn, mx)) = …`).
pub(crate) fn extract_axis_scaling(axis_node: &roxmltree::Node, _c_ns: &str) -> Option<(Option<f64>, Option<f64>)> {
    let (mn, mx) = ooxml_common::chart::extract_axis_min_max(*axis_node);
    if mn.is_some() || mx.is_some() { Some((mn, mx)) } else { None }
}

/// `<c:catAx|valAx><c:delete val="1"/>` — true when the axis is hidden
/// (ECMA-376 §21.2.2.40). Thin wrapper around the shared helper.
pub(crate) fn axis_is_deleted(axis_node: &roxmltree::Node, _c_ns: &str) -> bool {
    ooxml_common::chart::axis_is_deleted(*axis_node)
}

/// Extract a `<c:layout><c:manualLayout>` block. The given `layout_node` is
/// `<c:layout>` (parent of `<c:manualLayout>`). Returns None when the layout
/// is auto (no `manualLayout` child).
pub(crate) fn extract_manual_layout(layout_node: &roxmltree::Node, c_ns: &str) -> Option<ManualLayout> {
    let ml = layout_node.children()
        .find(|n| n.tag_name().name() == "manualLayout" && n.tag_name().namespace() == Some(c_ns))?;
    let mut x_mode = "edge".to_string();
    let mut y_mode = "edge".to_string();
    let mut layout_target: Option<String> = None;
    let mut x = 0.0_f64;
    let mut y = 0.0_f64;
    let mut w: Option<f64> = None;
    let mut h: Option<f64> = None;
    for ch in ml.children().filter(|n| n.is_element() && n.tag_name().namespace() == Some(c_ns)) {
        let val = ch.attribute("val").map(|s| s.to_string());
        match ch.tag_name().name() {
            "xMode" => { if let Some(v) = val { x_mode = v; } }
            "yMode" => { if let Some(v) = val { y_mode = v; } }
            "layoutTarget" => { layout_target = val; }
            "x" => { if let Some(v) = ch.attribute("val").and_then(|s| s.parse::<f64>().ok()) { x = v; } }
            "y" => { if let Some(v) = ch.attribute("val").and_then(|s| s.parse::<f64>().ok()) { y = v; } }
            "w" => { w = ch.attribute("val").and_then(|s| s.parse::<f64>().ok()); }
            "h" => { h = ch.attribute("val").and_then(|s| s.parse::<f64>().ok()); }
            _ => {}
        }
    }
    Some(ManualLayout { x_mode, y_mode, layout_target, x, y, w, h })
}

/// Extract a category/value axis tick-label font size (hundredths of a point)
/// from the first `a:defRPr@sz` (or `a:rPr@sz`) inside the axis' `c:txPr`.
/// ECMA-376 §21.2.2.17 — `<c:txPr>` controls tick label text properties.
pub(crate) fn extract_axis_tick_label_size(axis_node: &roxmltree::Node, _c_ns: &str, _a_ns: &str) -> Option<i32> {
    ooxml_common::chart::extract_axis_tick_label_size(*axis_node)
}

/// Extract the chart title's font size (hundredths of a point) from the first
/// `a:defRPr@sz` or `a:rPr@sz` found under `c:title`. Returns None when absent.
pub(crate) fn extract_chart_title_size(chart_root: &roxmltree::Node, c_ns: &str, a_ns: &str) -> Option<i32> {
    let title_node = chart_root.children()
        .find(|n| n.tag_name().name() == "title" && n.tag_name().namespace() == Some(c_ns))?;
    title_node.descendants().find_map(|n| {
        if !n.is_element() { return None; }
        if n.tag_name().namespace() != Some(a_ns) { return None; }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" { return None; }
        n.attribute("sz").and_then(|v| v.parse::<i32>().ok())
    })
}

/// Extract chart title bold flag from the first `a:defRPr@b` / `a:rPr@b`
/// inside `c:title`. Returns None when not specified (renderer treats as
/// not bold).
pub(crate) fn extract_chart_title_bold(chart_root: &roxmltree::Node, c_ns: &str, a_ns: &str) -> Option<bool> {
    let title_node = chart_root.children()
        .find(|n| n.tag_name().name() == "title" && n.tag_name().namespace() == Some(c_ns))?;
    title_node.descendants().find_map(|n| {
        if !n.is_element() { return None; }
        if n.tag_name().namespace() != Some(a_ns) { return None; }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" { return None; }
        n.attribute("b").map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
    })
}

/// Extract the chart title's font color (hex without '#') from the first
/// `a:solidFill/a:srgbClr@val` inside `c:title`. Only srgb is resolved here —
/// scheme colors would require the workbook theme, which isn't wired through
/// to chart parsing yet.
pub(crate) fn extract_chart_title_color(chart_root: &roxmltree::Node, c_ns: &str, a_ns: &str) -> Option<String> {
    let title_node = chart_root.children()
        .find(|n| n.tag_name().name() == "title" && n.tag_name().namespace() == Some(c_ns))?;
    title_node.descendants().find_map(|n| {
        if !n.is_element() { return None; }
        if n.tag_name().namespace() != Some(a_ns) { return None; }
        if n.tag_name().name() != "srgbClr" { return None; }
        // Skip srgbClr nodes that aren't inside a solidFill (e.g. a gradient stop).
        let parent_is_solid = n.parent()
            .map(|p| p.tag_name().name() == "solidFill" && p.tag_name().namespace() == Some(a_ns))
            .unwrap_or(false);
        if !parent_is_solid { return None; }
        n.attribute("val").map(|s| s.to_string())
    })
}

/// Extract the chart title's font family from the first `a:latin@typeface`
/// descendant of `c:title` (ECMA-376 DrawingML §20.1.4.2.24).
pub(crate) fn extract_chart_title_face(chart_root: &roxmltree::Node, c_ns: &str, a_ns: &str) -> Option<String> {
    let title_node = chart_root.children()
        .find(|n| n.tag_name().name() == "title" && n.tag_name().namespace() == Some(c_ns))?;
    title_node.descendants().find_map(|n| {
        if !n.is_element() { return None; }
        if n.tag_name().namespace() != Some(a_ns) { return None; }
        if n.tag_name().name() != "latin" { return None; }
        n.attribute("typeface").map(|s| s.to_string())
    })
}

/// Extract plain text from `c:chart/c:title`.
pub(crate) fn extract_chart_title(chart_root: &roxmltree::Node, c_ns: &str, a_ns: &str) -> Option<String> {
    let title_node = chart_root.children()
        .find(|n| n.tag_name().name() == "title" && n.tag_name().namespace() == Some(c_ns))?;
    // c:title/c:tx/c:rich/a:p/a:r/a:t  or  c:title/c:tx/c:strRef/c:strCache/c:pt/c:v
    let mut text = String::new();
    for node in title_node.descendants() {
        if node.tag_name().name() == "t" && node.tag_name().namespace() == Some(a_ns) {
            if let Some(t) = node.text() { text.push_str(t); }
        }
        if node.tag_name().name() == "v" && node.tag_name().namespace() == Some(c_ns) {
            if let Some(t) = node.text() { text.push_str(t); }
        }
    }
    if text.is_empty() { None } else { Some(text) }
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
pub(crate) fn resolve_series_color(ser_node: &roxmltree::Node, theme_colors: &[String]) -> Option<String> {
    let sp_pr = ser_node.children()
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
    let ln = sp_pr.children().find(|n| n.is_element() && n.tag_name().name() == "ln")?;
    resolve_fill_color(&ln, theme_colors)
}

pub(crate) fn parse_chart_series(
    node: &roxmltree::Node,
    c_ns: &str,
    ser_type: &str,
    chart_marker_default: bool,
    theme_colors: &[String],
) -> ChartSeries {
    let name = extract_series_name(node, c_ns);

    // `<c:idx val>` (ECMA-376 §21.2.2.27) — the canonical series index Excel
    // uses for default color selection. When absent, fall back to 0 so we
    // still produce a deterministic palette pick. `<c:order>` is the display
    // order (legend / stacking) and is intentionally ignored for coloring.
    let idx: usize = node.children()
        .find(|n| n.tag_name().name() == "idx" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(0);

    // `<c:order val>` (ECMA-376 §21.2.2.28) — series display order. Used for
    // stacking and legend ordering. Defaults to 0.
    let order: usize = node.children()
        .find(|n| n.tag_name().name() == "order" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(0);

    // For scatter: xVal → categories (as strings), yVal → values
    // For others:  cat  → categories,             val  → values
    let (cat_tag, val_tag) = if ser_type == "scatter" { ("xVal", "yVal") } else { ("cat", "val") };

    let categories = collect_str_cache(node, c_ns, cat_tag);
    let values     = collect_num_cache(node, c_ns, val_tag);
    // `<c:val><c:numRef><c:numCache><c:formatCode>` (ECMA-376 §21.2.2.37)
    // preserves the Excel number format Excel stamped onto the cached values
    // at save time; absent "General" codes return None so the renderer can
    // fall back cleanly.
    let val_format_code = node.children()
        .find(|n| n.tag_name().name() == val_tag && n.tag_name().namespace() == Some(c_ns))
        .and_then(|val_node| val_node.descendants()
            .find(|n| n.is_element()
                && n.tag_name().name() == "formatCode"
                && n.tag_name().namespace() == Some(c_ns)))
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
    let color = resolve_series_color(node, theme_colors)
        .or_else(|| {
            // Theme order in `theme_colors`: dk1@0, lt1@1, dk2@2, lt2@3, accent1@4 … accent6@9.
            theme_colors.get(4 + (idx % 6)).map(|c| c.trim_start_matches('#').to_lowercase())
        });

    // Marker visibility (ECMA-376 §21.2.2.32 — c:marker/c:symbol default is
    // "none"). A per-series <c:marker><c:symbol> overrides; otherwise fall
    // back to the chart-type-level <c:lineChart><c:marker val> flag. Scatter
    // charts default to visible markers even without an explicit flag.
    let marker_node = node.children()
        .find(|n| n.tag_name().name() == "marker" && n.tag_name().namespace() == Some(c_ns));
    let (marker_symbol, marker_size, marker_fill, marker_line) = parse_marker_block(marker_node, c_ns, theme_colors);
    let show_marker = match (&marker_symbol, ser_type) {
        (Some(sym), _)   => sym != "none",
        (None, "scatter") => true,
        _                 => chart_marker_default,
    };

    let data_point_overrides = parse_data_point_overrides(node, c_ns, theme_colors);
    // `<c15:datalabelsRange>` lookup table for `<a:fld type="CELLRANGE">`
    // labels. Excel saves the actual cached label strings here; we resolve
    // CELLRANGE field placeholders against this at parse time so the
    // renderer just receives plain strings.
    let dlbl_range_cache = collect_dlbl_range_cache(node, c_ns);
    let (series_data_labels, data_label_overrides) = parse_data_labels(node, c_ns, theme_colors, &dlbl_range_cache);
    let err_bars = parse_error_bars(node, c_ns, &values, theme_colors);

    ChartSeries {
        name,
        series_type: ser_type.to_string(),
        categories,
        values,
        color,
        show_marker,
        val_format_code,
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
    c_ns: &str,
    theme_colors: &[String],
) -> (Option<String>, Option<u32>, Option<String>, Option<String>) {
    let Some(mk) = marker_node else { return (None, None, None, None); };
    let symbol = mk.children()
        .find(|n| n.tag_name().name() == "symbol" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string());
    let size = mk.children()
        .find(|n| n.tag_name().name() == "size" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<u32>().ok());
    let sp_pr = mk.children()
        .find(|n| n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(c_ns));
    let fill = sp_pr.and_then(|p| extract_solid_fill_in_drawingml(&p, theme_colors));
    let line = sp_pr.and_then(|p| {
        let ln = p.children().find(|n| n.tag_name().name() == "ln");
        ln.and_then(|l| extract_solid_fill_in_drawingml(&l, theme_colors))
    });
    (symbol, size, fill, line)
}

/// Locate the first `a:solidFill > a:srgbClr@val` or `a:schemeClr@val` under
/// `node` (children only, not deep descendants — chart spPr is structured
/// shallowly). Returns the resolved hex without `#`. Handles theme refs and
/// `lumMod`/`lumOff`/`tint`/`shade`/`alpha` color transforms by delegating
/// to `apply_color_transforms`.
pub(crate) fn extract_solid_fill_in_drawingml(
    parent: &roxmltree::Node,
    theme_colors: &[String],
) -> Option<String> {
    for fill in parent.children().filter(|n| n.is_element() && n.tag_name().name() == "solidFill") {
        for clr in fill.children().filter(|n| n.is_element()) {
            match clr.tag_name().name() {
                "srgbClr" => {
                    if let Some(rgb) = clr.attribute("val") {
                        return Some(apply_color_transforms(rgb, &clr));
                    }
                }
                "schemeClr" => {
                    if let Some(scheme) = clr.attribute("val") {
                        let base = resolve_scheme_color(scheme, theme_colors);
                        if let Some(b) = base {
                            return Some(apply_color_transforms(&b, &clr));
                        }
                    }
                }
                _ => {}
            }
        }
    }
    None
}

/// Look up a scheme color name ("dk1"/"lt1"/"dk2"/"lt2"/"accent1"…"accent6"
/// /"hlink"/"folHlink") in the workbook theme color table. Returns hex
/// (no `#`) or None when unknown.
pub(crate) fn resolve_scheme_color(name: &str, theme_colors: &[String]) -> Option<String> {
    // Theme order (parse_theme_colors): dk1@0, lt1@1, dk2@2, lt2@3,
    // accent1@4..accent6@9, hlink@10, folHlink@11.
    let idx = match name {
        "dk1" | "tx1" | "bg2" => 0,
        "lt1" | "bg1" | "tx2" => 1,
        "dk2" => 2,
        "lt2" => 3,
        "accent1" => 4, "accent2" => 5, "accent3" => 6,
        "accent4" => 7, "accent5" => 8, "accent6" => 9,
        "hlink" => 10, "folHlink" => 11,
        _ => return None,
    };
    theme_colors.get(idx).map(|s| s.trim_start_matches('#').to_string())
}

/// Apply DrawingML color transforms (`lumMod`/`lumOff`/`tint`/`shade`/
/// `alpha` — drop alpha) found as children of a color element. Returns a
/// hex string without `#`. Already-existing `apply_tint` handles
/// lumMod-style brightness changes for the simpler `lumMod-only` case;
/// this widens it to combine multiple transforms.
pub(crate) fn apply_color_transforms(base_hex: &str, color_el: &roxmltree::Node) -> String {
    let cleaned = base_hex.trim_start_matches('#');
    let r = u8::from_str_radix(&cleaned.get(0..2).unwrap_or("00"), 16).unwrap_or(0);
    let g = u8::from_str_radix(&cleaned.get(2..4).unwrap_or("00"), 16).unwrap_or(0);
    let b = u8::from_str_radix(&cleaned.get(4..6).unwrap_or("00"), 16).unwrap_or(0);
    let mut rf = r as f64 / 255.0;
    let mut gf = g as f64 / 255.0;
    let mut bf = b as f64 / 255.0;
    for child in color_el.children().filter(|n| n.is_element()) {
        let pct = child.attribute("val")
            .and_then(|v| v.parse::<f64>().ok())
            .map(|v| v / 100000.0);
        let Some(p) = pct else { continue };
        match child.tag_name().name() {
            "lumMod"  => { rf *= p; gf *= p; bf *= p; }
            "lumOff"  => { rf += p; gf += p; bf += p; }
            "tint"    => {
                // ECMA-376: lighten toward 1.0 by `p` (0..1).
                rf = rf + (1.0 - rf) * p;
                gf = gf + (1.0 - gf) * p;
                bf = bf + (1.0 - bf) * p;
            }
            "shade"   => {
                // Darken toward 0 by `1 - p`.
                rf *= p; gf *= p; bf *= p;
            }
            // alpha is dropped — we render opaque.
            _ => {}
        }
    }
    let clamp = |v: f64| -> u8 { (v.max(0.0).min(1.0) * 255.0).round() as u8 };
    format!("{:02X}{:02X}{:02X}", clamp(rf), clamp(gf), clamp(bf))
}

/// Walk every `<c:dPt>` direct child of the series and collect per-point
/// overrides. Multiple `<c:dPt>` per series is normal; each one targets a
/// single `<c:idx>` (ECMA-376 §21.2.2.39).
pub(crate) fn parse_data_point_overrides(
    ser_node: &roxmltree::Node,
    c_ns: &str,
    theme_colors: &[String],
) -> Vec<DataPointOverride> {
    let mut result = Vec::new();
    for dpt in ser_node.children().filter(|n| n.is_element() && n.tag_name().name() == "dPt" && n.tag_name().namespace() == Some(c_ns)) {
        let idx = dpt.children()
            .find(|n| n.tag_name().name() == "idx" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.attribute("val"))
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(0);
        let sp_pr = dpt.children()
            .find(|n| n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(c_ns));
        let color = sp_pr.and_then(|p| extract_solid_fill_in_drawingml(&p, theme_colors));
        let mk = dpt.children()
            .find(|n| n.tag_name().name() == "marker" && n.tag_name().namespace() == Some(c_ns));
        let (marker_symbol, marker_size, marker_fill, marker_line) = parse_marker_block(mk, c_ns, theme_colors);
        result.push(DataPointOverride {
            idx, color, marker_symbol, marker_size, marker_fill, marker_line,
        });
    }
    result
}

/// Resolve `<c:ser><c:extLst><c:ext><c15:datalabelsRange>` cache: index →
/// label text. Used to substitute `<a:fld type="CELLRANGE">` placeholders.
/// Returns indices in 0..ptCount; missing entries are empty strings.
pub(crate) fn collect_dlbl_range_cache(ser_node: &roxmltree::Node, c_ns: &str) -> HashMap<u32, String> {
    let mut map: HashMap<u32, String> = HashMap::new();
    let Some(ext_lst) = ser_node.children().find(|n| n.tag_name().name() == "extLst" && n.tag_name().namespace() == Some(c_ns)) else { return map; };
    for ext in ext_lst.children().filter(|n| n.is_element() && n.tag_name().name() == "ext" && n.tag_name().namespace() == Some(c_ns)) {
        for range in ext.descendants().filter(|n| n.is_element() && n.tag_name().name() == "datalabelsRange") {
            for cache in range.children().filter(|n| n.is_element() && n.tag_name().name() == "dlblRangeCache") {
                for pt in cache.children().filter(|n| n.is_element() && n.tag_name().name() == "pt" && n.tag_name().namespace() == Some(c_ns)) {
                    let Some(idx) = pt.attribute("idx").and_then(|v| v.parse::<u32>().ok()) else { continue };
                    let v = pt.children()
                        .find(|n| n.tag_name().name() == "v" && n.tag_name().namespace() == Some(c_ns))
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
    for p in rich_root.descendants().filter(|n| n.is_element() && n.tag_name().name() == "p") {
        if !first_para { out.push('\n'); }
        first_para = false;
        for child in p.children().filter(|n| n.is_element()) {
            match child.tag_name().name() {
                "r" => {
                    if let Some(t) = child.children().find(|n| n.tag_name().name() == "t") {
                        if let Some(s) = t.text() { out.push_str(s); }
                    }
                }
                "fld" => {
                    let typ = child.attribute("type").unwrap_or("");
                    if typ == "CELLRANGE" {
                        if let Some(s) = cellrange_cache { out.push_str(s); }
                    } else {
                        // VALUE/SERIESNAME/CATEGORYNAME field placeholders are
                        // resolved by the renderer using the series data, since
                        // they don't need cell-range expansion. We embed a marker
                        // so the renderer can recognise them.
                        if let Some(t) = child.children().find(|n| n.tag_name().name() == "t") {
                            if let Some(s) = t.text() { out.push_str(s); }
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
    c_ns: &str,
    theme_colors: &[String],
    cellrange_cache: &HashMap<u32, String>,
) -> (Option<SeriesDataLabels>, Vec<DataLabelOverride>) {
    let Some(d_lbls) = ser_node.children()
        .find(|n| n.tag_name().name() == "dLbls" && n.tag_name().namespace() == Some(c_ns))
    else { return (None, Vec::new()); };

    let bool_attr = |n: &roxmltree::Node, name: &str| {
        n.children()
            .find(|c| c.tag_name().name() == name && c.tag_name().namespace() == Some(c_ns))
            .and_then(|c| c.attribute("val"))
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
            .unwrap_or(false)
    };

    let position = d_lbls.children()
        .find(|n| n.tag_name().name() == "dLblPos" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string());
    let format_code = d_lbls.children()
        .find(|n| n.tag_name().name() == "numFmt" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("formatCode"))
        .map(|s| s.to_string());
    let font_color = d_lbls.children()
        .find(|n| n.tag_name().name() == "txPr" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|tx| {
            tx.descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "defRPr")
                .and_then(|def| extract_solid_fill_in_drawingml(&def, theme_colors))
        });
    let font_bold_default = d_lbls.children()
        .find(|n| n.tag_name().name() == "txPr" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|tx| {
            tx.descendants()
                .find(|n| n.is_element() && (n.tag_name().name() == "defRPr" || n.tag_name().name() == "rPr"))
                .and_then(|n| n.attribute("b"))
                .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
        });
    let font_size_default = d_lbls.children()
        .find(|n| n.tag_name().name() == "txPr" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|tx| {
            tx.descendants()
                .find(|n| n.is_element() && (n.tag_name().name() == "defRPr" || n.tag_name().name() == "rPr"))
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
    for dl in d_lbls.children().filter(|n| n.is_element() && n.tag_name().name() == "dLbl" && n.tag_name().namespace() == Some(c_ns)) {
        let idx = dl.children()
            .find(|n| n.tag_name().name() == "idx" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.attribute("val"))
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(0);
        // <c:delete val="1"/> — the user explicitly removed this point's
        // label. Render as empty text so the renderer skips it.
        let deleted = dl.children()
            .find(|n| n.tag_name().name() == "delete" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.attribute("val"))
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
            .unwrap_or(false);
        let pos = dl.children()
            .find(|n| n.tag_name().name() == "dLblPos" && n.tag_name().namespace() == Some(c_ns))
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
            let tx = dl.children()
                .find(|n| n.tag_name().name() == "tx" && n.tag_name().namespace() == Some(c_ns));
            match tx {
                Some(tx_node) => flatten_rich_text(&tx_node, cache_for_idx),
                None => cache_for_idx.unwrap_or("").to_string(),
            }
        };
        let font_color = dl.children()
            .find(|n| n.tag_name().name() == "txPr" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|tx| {
                tx.descendants()
                    .find(|n| n.is_element() && n.tag_name().name() == "defRPr")
                    .and_then(|def| extract_solid_fill_in_drawingml(&def, theme_colors))
            });
        let font_size_hpt = dl.descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "defRPr")
            .and_then(|def| def.attribute("sz"))
            .and_then(|v| v.parse::<i32>().ok());
        let font_bold = dl.descendants()
            .find(|n| n.is_element() && (n.tag_name().name() == "defRPr" || n.tag_name().name() == "rPr"))
            .and_then(|def| def.attribute("b"))
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"));
        overrides.push(DataLabelOverride { idx, text, position: pos, font_color, font_size_hpt, font_bold });
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
    let series_out = if any_default { Some(series_defaults) } else { None };
    (series_out, overrides)
}

/// Parse all `<c:errBars>` direct children of a series and resolve per-
/// point plus / minus deltas to absolute numbers. Each errBars block fixes
/// a direction (x|y); a series can have at most one of each direction.
pub(crate) fn parse_error_bars(
    ser_node: &roxmltree::Node,
    c_ns: &str,
    series_values: &[Option<f64>],
    theme_colors: &[String],
) -> Vec<ErrBars> {
    let mut result = Vec::new();
    for eb in ser_node.children().filter(|n| n.is_element() && n.tag_name().name() == "errBars" && n.tag_name().namespace() == Some(c_ns)) {
        let dir = eb.children()
            .find(|n| n.tag_name().name() == "errDir" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.attribute("val"))
            .unwrap_or("y")
            .to_string();
        let bar_type = eb.children()
            .find(|n| n.tag_name().name() == "errBarType" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.attribute("val"))
            .unwrap_or("both")
            .to_string();
        let val_type = eb.children()
            .find(|n| n.tag_name().name() == "errValType" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.attribute("val"))
            .unwrap_or("fixedVal")
            .to_string();
        let no_end_cap = eb.children()
            .find(|n| n.tag_name().name() == "noEndCap" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.attribute("val"))
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
            .unwrap_or(false);

        let n_points = series_values.len();
        let mut plus: Vec<Option<f64>> = vec![None; n_points];
        let mut minus: Vec<Option<f64>> = vec![None; n_points];

        match val_type.as_str() {
            "cust" => {
                for (slot, target) in [("plus", &mut plus), ("minus", &mut minus)] {
                    let Some(side) = eb.children().find(|n| n.tag_name().name() == slot && n.tag_name().namespace() == Some(c_ns)) else { continue };
                    let vals = extract_num_block(&side, c_ns, n_points);
                    if !vals.is_empty() {
                        let len = vals.len().min(target.len());
                        for i in 0..len { target[i] = vals[i]; }
                    }
                }
            }
            "fixedVal" => {
                let v = eb.children()
                    .find(|n| n.tag_name().name() == "val" && n.tag_name().namespace() == Some(c_ns))
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
                let pct = eb.children()
                    .find(|n| n.tag_name().name() == "val" && n.tag_name().namespace() == Some(c_ns))
                    .and_then(|n| n.attribute("val"))
                    .and_then(|s| s.parse::<f64>().ok())
                    .unwrap_or(0.0);
                for (i, v) in series_values.iter().enumerate() {
                    if let Some(val) = v {
                        let d = val.abs() * pct / 100.0;
                        plus[i] = Some(d); minus[i] = Some(d);
                    }
                }
            }
            "stdErr" | "stdDev" => {
                let nums: Vec<f64> = series_values.iter().filter_map(|v| *v).collect();
                if !nums.is_empty() {
                    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
                    let var = nums.iter().map(|v| (v - mean).powi(2)).sum::<f64>() / nums.len() as f64;
                    let std = var.sqrt();
                    let mult = eb.children()
                        .find(|n| n.tag_name().name() == "val" && n.tag_name().namespace() == Some(c_ns))
                        .and_then(|n| n.attribute("val"))
                        .and_then(|s| s.parse::<f64>().ok())
                        .unwrap_or(1.0);
                    let sample = if val_type == "stdErr" {
                        std / (nums.len() as f64).sqrt()
                    } else { std };
                    let delta = sample * mult;
                    for i in 0..n_points {
                        plus[i] = Some(delta); minus[i] = Some(delta);
                    }
                }
            }
            _ => {}
        }

        let sp_pr = eb.children()
            .find(|n| n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(c_ns));
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
            dir, bar_type,
            plus, minus,
            no_end_cap,
            color, line_width_emu, dash,
        });
    }
    result
}

/// Read a `<c:numRef><c:numCache>` or `<c:numLit>` block under `parent` and
/// return per-point values keyed by `<c:pt idx>`. Length is at least
/// `expected_len` (padded with None).
pub(crate) fn extract_num_block(parent: &roxmltree::Node, c_ns: &str, expected_len: usize) -> Vec<Option<f64>> {
    let cache = parent.descendants()
        .find(|n| n.is_element()
            && (n.tag_name().name() == "numCache" || n.tag_name().name() == "numLit")
            && n.tag_name().namespace() == Some(c_ns));
    let Some(cache) = cache else { return Vec::new(); };
    let pt_count: usize = cache.children()
        .find(|n| n.tag_name().name() == "ptCount" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(expected_len);
    let len = pt_count.max(expected_len);
    let mut values: Vec<Option<f64>> = vec![None; len];
    for pt in cache.children().filter(|n| n.tag_name().name() == "pt" && n.tag_name().namespace() == Some(c_ns)) {
        let Some(idx) = pt.attribute("idx").and_then(|v| v.parse::<usize>().ok()) else { continue };
        let v = pt.children()
            .find(|n| n.tag_name().name() == "v" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.text())
            .and_then(|s| s.trim().parse::<f64>().ok());
        if idx < values.len() { values[idx] = v; }
    }
    values
}

/// Extract series name from `c:tx`.
pub(crate) fn extract_series_name(node: &roxmltree::Node, c_ns: &str) -> String {
    // c:tx/c:strRef/c:strCache/c:pt[@idx=0]/c:v
    // or c:tx/c:v
    if let Some(tx) = node.children().find(|n| n.tag_name().name() == "tx" && n.tag_name().namespace() == Some(c_ns)) {
        for desc in tx.descendants() {
            if desc.tag_name().name() == "v" && desc.tag_name().namespace() == Some(c_ns) {
                if let Some(t) = desc.text() {
                    if !t.is_empty() { return t.to_string(); }
                }
            }
        }
    }
    String::new()
}

/// Collect string values from a cache child element (e.g. `<c:cat>` or `<c:xVal>`).
/// Reads `c:strRef/c:strCache`, `c:multiLvlStrRef/c:multiLvlStrCache`, or
/// `c:numRef/c:numCache` (formats numbers as strings).
pub(crate) fn collect_str_cache(ser_node: &roxmltree::Node, c_ns: &str, child_tag: &str) -> Vec<String> {
    let Some(child) = ser_node.children()
        .find(|n| n.tag_name().name() == child_tag && n.tag_name().namespace() == Some(c_ns))
    else { return Vec::new(); };

    // Multi-level categories: use only the first (innermost) lvl to get primary labels.
    if let Some(multi_cache) = child.descendants()
        .find(|n| n.tag_name().name() == "multiLvlStrCache" && n.tag_name().namespace() == Some(c_ns))
    {
        let pt_count: usize = multi_cache.children()
            .find(|n| n.tag_name().name() == "ptCount" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.attribute("val"))
            .and_then(|v| v.parse().ok())
            .unwrap_or(0);
        if let Some(first_lvl) = multi_cache.children()
            .find(|n| n.tag_name().name() == "lvl" && n.tag_name().namespace() == Some(c_ns))
        {
            let mut pts: Vec<(usize, String)> = Vec::new();
            for pt in first_lvl.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "pt" && n.tag_name().namespace() == Some(c_ns))
            {
                let idx: usize = pt.attribute("idx").and_then(|v| v.parse().ok()).unwrap_or(0);
                let val = pt.children()
                    .find(|n| n.tag_name().name() == "v")
                    .and_then(|n| n.text())
                    .unwrap_or("")
                    .to_string();
                pts.push((idx, val));
            }
            let len = pt_count.max(pts.iter().map(|(i, _)| i + 1).max().unwrap_or(0));
            let mut result = vec![String::new(); len];
            for (idx, val) in pts {
                if idx < result.len() { result[idx] = val; }
            }
            return result;
        }
    }

    // Standard strRef/strCache or numRef/numCache
    let mut pt_count: usize = 0;
    let mut pts: Vec<(usize, String)> = Vec::new();
    for desc in child.descendants() {
        match desc.tag_name().name() {
            "ptCount" if desc.tag_name().namespace() == Some(c_ns) => {
                pt_count = desc.attribute("val").and_then(|v| v.parse().ok()).unwrap_or(0);
            }
            "pt" if desc.tag_name().namespace() == Some(c_ns) => {
                let idx: usize = desc.attribute("idx").and_then(|v| v.parse().ok()).unwrap_or(0);
                let val = desc.children()
                    .find(|n| n.tag_name().name() == "v")
                    .and_then(|n| n.text())
                    .unwrap_or("")
                    .to_string();
                pts.push((idx, val));
            }
            _ => {}
        }
    }
    if pt_count == 0 { pt_count = pts.len(); }
    let mut result = vec![String::new(); pt_count];
    for (idx, val) in pts {
        if idx < result.len() { result[idx] = val; }
    }
    result
}

/// Collect numeric values from a cache child element (e.g. `<c:val>` or `<c:yVal>`).
pub(crate) fn collect_num_cache(ser_node: &roxmltree::Node, c_ns: &str, child_tag: &str) -> Vec<Option<f64>> {
    let Some(child) = ser_node.children()
        .find(|n| n.tag_name().name() == child_tag && n.tag_name().namespace() == Some(c_ns))
    else { return Vec::new(); };

    let mut pt_count: usize = 0;
    let mut pts: Vec<(usize, f64)> = Vec::new();
    for desc in child.descendants() {
        match desc.tag_name().name() {
            "ptCount" if desc.tag_name().namespace() == Some(c_ns) => {
                pt_count = desc.attribute("val").and_then(|v| v.parse().ok()).unwrap_or(0);
            }
            "pt" if desc.tag_name().namespace() == Some(c_ns) => {
                let idx: usize = desc.attribute("idx").and_then(|v| v.parse().ok()).unwrap_or(0);
                if let Some(v) = desc.children()
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
    if pt_count == 0 { pt_count = pts.len(); }
    let mut result: Vec<Option<f64>> = vec![None; pt_count];
    for (idx, val) in pts {
        if idx < result.len() { result[idx] = Some(val); }
    }
    result
}

