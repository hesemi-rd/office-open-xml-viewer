//! Chart element parsing: the legacy DrawingML chart (`c:` namespace) and the
//! newer chartEx (`cx:` namespace) parsers, plus the pptx `ColorResolver` the
//! shared `ooxml_common::chart` helpers use to resolve `<a:solidFill>` colours.
//! Extracted verbatim from `lib.rs`. The general colour grammar
//! (`parse_color_node`) and the shared XML helpers (`child`, `attr`) stay in
//! `lib.rs` and are imported here.

use crate::types::*;
use crate::{attr, parse_color_node};
use std::collections::HashMap;

/// `ooxml_common::chart::ColorResolver` implementation backed by pptx's
/// `HashMap<String, String>` theme palette and PowerPoint's tint formula.
/// Used by chart helpers in ooxml-common that need to resolve
/// `<a:solidFill>` text colors without owning the theme storage.
pub(crate) struct PptxColorResolver<'a> {
    pub(crate) theme: &'a HashMap<String, String>,
}

impl ooxml_common::chart::ColorResolver for PptxColorResolver<'_> {
    fn resolve_solid_fill(&self, node: roxmltree::Node<'_, '_>) -> Option<String> {
        parse_color_node(node, self.theme)
    }
}

/// Parse a legacy OOXML chart (c: namespace) — barChart / lineChart etc.
pub(crate) fn parse_legacy_chart(
    xml: &str,
    theme: &HashMap<String, String>,
) -> Option<ChartElement> {
    let doc = roxmltree::Document::parse(xml).ok()?;
    let root = doc.root_element();

    // Determine chart type by finding the first recognized chart element
    let find_chart = |name: &str| {
        root.descendants()
            .find(|n| n.is_element() && n.tag_name().name() == name)
    };

    let chart_type = if let Some(bc) = find_chart("barChart") {
        let grouping = bc
            .children()
            .find(|c| c.is_element() && c.tag_name().name() == "grouping")
            .and_then(|n| attr(&n, "val"))
            .unwrap_or_else(|| "clustered".into());
        let bar_dir = bc
            .children()
            .find(|c| c.is_element() && c.tag_name().name() == "barDir")
            .and_then(|n| attr(&n, "val"))
            .unwrap_or_else(|| "col".into());
        let horizontal = bar_dir == "bar";
        match (grouping.as_str(), horizontal) {
            ("stacked" | "percentStacked", false) => "stackedBar".to_string(),
            ("stacked" | "percentStacked", true) => "stackedBarH".to_string(),
            (_, false) => "clusteredBar".to_string(),
            (_, true) => "clusteredBarH".to_string(),
        }
    } else if let Some(lc) = find_chart("lineChart") {
        let grouping = lc
            .children()
            .find(|c| c.is_element() && c.tag_name().name() == "grouping")
            .and_then(|n| attr(&n, "val"))
            .unwrap_or_else(|| "standard".into());
        match grouping.as_str() {
            "stacked" => "stackedLine".to_string(),
            "percentStacked" => "stackedLinePct".to_string(),
            _ => "line".to_string(),
        }
    } else if find_chart("pieChart").is_some() {
        "pie".to_string()
    } else if find_chart("doughnutChart").is_some() {
        "doughnut".to_string()
    } else if let Some(ac) = find_chart("areaChart") {
        let grouping = ac
            .children()
            .find(|c| c.is_element() && c.tag_name().name() == "grouping")
            .and_then(|n| attr(&n, "val"))
            .unwrap_or_else(|| "standard".into());
        match grouping.as_str() {
            "stacked" => "stackedArea".to_string(),
            _ => "area".to_string(),
        }
    } else if find_chart("scatterChart").is_some() {
        "scatter".to_string()
    } else if find_chart("bubbleChart").is_some() {
        "bubble".to_string()
    } else if find_chart("radarChart").is_some() {
        "radar".to_string()
    } else {
        "unknown".to_string()
    };

    // Title text
    let title_node_opt = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "title");
    let title = title_node_opt.and_then(|title_node| {
        let texts: Vec<String> = title_node
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "t")
            .filter_map(|n| n.text().map(|t| t.to_string()))
            .collect();
        if texts.is_empty() {
            None
        } else {
            Some(texts.join(""))
        }
    });
    // Title font size in hundredths of a point — taken from the first
    // defRPr@sz or rPr@sz we find inside the title. ECMA-376 uses hpt for size.
    let title_font_size_hpt = title_node_opt.and_then(|t| {
        t.descendants().find_map(|n| {
            if !n.is_element() {
                return None;
            }
            let tag = n.tag_name().name();
            if tag != "defRPr" && tag != "rPr" {
                return None;
            }
            attr(&n, "sz").and_then(|v| v.parse::<i32>().ok())
        })
    });

    // val axis max / min and visibility — shared helpers in ooxml-common
    // so xlsx & pptx stay in sync (`<c:scaling><c:min|max val>` §21.2.2.160
    // scaling and `<c:delete val>` ECMA-376 §21.2.2.40 delete).
    // Combo charts (bar + line) declare TWO `<c:valAx>`: a PRIMARY (axPos="l",
    // `<c:crosses val="autoZero">`) and a SECONDARY (axPos="r",
    // `<c:crosses val="max">`). Collect both so series can be mapped to the
    // right scale and the right-hand axis drawn. The primary axis keeps driving
    // every existing axis read below; only the secondary is new.
    let val_ax_nodes: Vec<roxmltree::Node> = root
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "valAx")
        .collect();
    let ax_pos = |n: &roxmltree::Node| -> Option<String> {
        n.children()
            .find(|c| c.is_element() && c.tag_name().name() == "axPos")
            .and_then(|c| attr(&c, "val"))
    };
    let ax_id_of = |n: &roxmltree::Node| -> Option<String> {
        n.children()
            .find(|c| c.is_element() && c.tag_name().name() == "axId")
            .and_then(|c| attr(&c, "val"))
    };
    // Primary = first value axis that isn't on the right; secondary (only when a
    // second value axis exists) = a value axis on the right edge.
    let val_ax = val_ax_nodes
        .iter()
        .find(|n| ax_pos(n).as_deref() != Some("r"))
        .or_else(|| val_ax_nodes.first())
        .copied();
    let secondary_val_ax = if val_ax_nodes.len() >= 2 {
        val_ax_nodes
            .iter()
            .find(|n| ax_pos(n).as_deref() == Some("r"))
            .copied()
    } else {
        None
    };
    let secondary_ax_id = secondary_val_ax.as_ref().and_then(ax_id_of);
    // The category axis is `<c:catAx>` or, for a date/time-series X axis,
    // `<c:dateAx>` (§21.2.2.39) — same child grammar, so every cat-axis read
    // below (hidden, tick label size/color/bold, line style, tick marks, format
    // code) treats them identically.
    let cat_ax = root
        .descendants()
        .find(|n| n.is_element() && matches!(n.tag_name().name(), "catAx" | "dateAx"));
    let (val_min, val_max) = val_ax
        .map(ooxml_common::chart::extract_axis_min_max)
        .unwrap_or((None, None));
    let val_axis_hidden = val_ax
        .map(ooxml_common::chart::axis_is_deleted)
        .unwrap_or(false);
    let cat_axis_hidden = cat_ax
        .map(ooxml_common::chart::axis_is_deleted)
        .unwrap_or(false);

    // Series
    let plot_area = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "plotArea")?;

    // Plot area background: <c:plotArea><c:spPr><a:solidFill>
    let plot_area_bg = plot_area
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spPr")
        .and_then(|sp| {
            sp.children()
                .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
        })
        .and_then(|fill| parse_color_node(fill, theme));

    let ser_nodes: Vec<_> = plot_area
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "ser")
        .collect();

    if ser_nodes.is_empty() {
        return None;
    }

    // Helper: collect <c:pt> values from a cache node (strCache or numCache)
    let collect_pt_strings = |cache: roxmltree::Node<'_, '_>| -> Vec<String> {
        cache
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "pt")
            .filter_map(|pt| {
                pt.children()
                    .find(|n| n.is_element() && n.tag_name().name() == "v")
            })
            .filter_map(|v| v.text().map(|t| t.to_string()))
            .collect()
    };

    // ECMA-376 §21.2.2: category data may be in a *Cache (backing a *Ref) or a *Lit (inline literal).
    // Accept strCache/numCache (external refs with cached values) AND strLit/numLit (inline literals).
    let is_pt_container =
        |name: &str| matches!(name, "strCache" | "numCache" | "strLit" | "numLit");

    // Category labels live under `<c:cat>` in one of two shapes:
    //   * single-level: a strCache/numCache/strLit/numLit holding `<c:pt>`
    //     children directly, or
    //   * multi-level: a `<c:multiLvlStrCache>` (ECMA-376 §21.2.2.95) whose
    //     `<c:pt>` are nested one level deeper inside `<c:lvl>` elements.
    // For a multi-level axis we use the innermost level (the first `<c:lvl>`),
    // which carries one label per data point — matching how Word / PowerPoint
    // label the category axis. Missing this path left `categories` empty, which
    // collapsed `pt_count` to 1 and truncated every series to a single value
    // (issue #556: area chart rendered as a blank zero-width sliver).
    let cat_node = ser_nodes[0]
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "cat");
    let categories: Vec<String> = cat_node
        .and_then(|cat| {
            cat.descendants()
                .find(|n| n.is_element() && is_pt_container(n.tag_name().name()))
                .map(&collect_pt_strings)
                .filter(|v| !v.is_empty())
                .or_else(|| {
                    // Multi-level: collect from the first `<c:lvl>` of the cache.
                    cat.descendants()
                        .find(|n| n.is_element() && n.tag_name().name() == "lvl")
                        .map(&collect_pt_strings)
                })
        })
        .unwrap_or_default();

    let pt_count = categories.len().max(1);

    let is_scatter_like = chart_type == "scatter" || chart_type == "bubble";

    // A combo chart's primary type is the first recognized group (bar wins);
    // line series only need the "line" override when the chart isn't ALREADY a
    // line chart (otherwise every series in a pure line chart would carry a
    // redundant override).
    let primary_is_line =
        chart_type == "line" || chart_type == "stackedLine" || chart_type == "stackedLinePct";

    let series: Vec<ChartSeriesData> = ser_nodes
        .iter()
        .map(|ser| {
            // Each `<c:ser>` is a direct child of its chart-group element
            // (`<c:barChart>`/`<c:lineChart>`/…). Tag line-group series of a
            // combo chart so the renderer draws them as a line over the columns
            // (ECMA-376 §21.2.2.97), and flag series whose group references the
            // secondary value axis so they plot against the right-hand scale.
            let group = ser.parent();
            let series_type = match group.map(|p| p.tag_name().name()) {
                Some("lineChart") if !primary_is_line => Some("line".to_string()),
                _ => None,
            };
            let use_secondary_axis = match (group, secondary_ax_id.as_deref()) {
                (Some(g), Some(sec)) => g
                    .children()
                    .filter(|c| c.is_element() && c.tag_name().name() == "axId")
                    .any(|c| attr(&c, "val").as_deref() == Some(sec)),
                _ => false,
            };

            // Series name from <c:tx>  (can be strRef/strCache, strLit, or a bare <c:v>)
            let name = ser
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "tx")
                .and_then(|tx| {
                    // Preferred: first pt > v inside any cache/lit container
                    tx.descendants()
                        .find(|n| n.is_element() && is_pt_container(n.tag_name().name()))
                        .and_then(|cache| {
                            cache
                                .children()
                                .find(|n| n.is_element() && n.tag_name().name() == "pt")
                                .and_then(|pt| {
                                    pt.children()
                                        .find(|n| n.is_element() && n.tag_name().name() == "v")
                                })
                                .and_then(|v| v.text().map(|t| t.to_string()))
                        })
                        // Fallback: <c:tx><c:v>Name</c:v></c:tx>
                        .or_else(|| {
                            tx.children()
                                .find(|n| n.is_element() && n.tag_name().name() == "v")
                                .and_then(|v| v.text().map(|t| t.to_string()))
                        })
                })
                .unwrap_or_default();

            // Per-series X values for scatter/bubble: ECMA-376 §21.2.2.43 puts numeric
            // X data in `<c:xVal>` (with its own numCache / numLit) instead of the
            // shared `<c:cat>`. Read it as strings so the core ChartSeries.categories
            // field can stay string-typed (renderScatterChart parses each entry back
            // to a float).
            let x_cache = if is_scatter_like {
                ser.children()
                    .find(|n| n.is_element() && n.tag_name().name() == "xVal")
                    .and_then(|x| {
                        x.descendants()
                            .find(|n| n.is_element() && is_pt_container(n.tag_name().name()))
                    })
            } else {
                None
            };
            let series_categories: Option<Vec<String>> = x_cache.map(&collect_pt_strings);

            // Y values: scatter/bubble use `<c:yVal>`, everything else uses `<c:val>`.
            // Restrict the descendant walk to the matching tag so a sibling `<c:xVal>`
            // (also a numCache) can't be picked up as the Y series.
            let val_cache = if is_scatter_like {
                ser.children()
                    .find(|n| n.is_element() && n.tag_name().name() == "yVal")
                    .and_then(|y| {
                        y.descendants().find(|n| {
                            n.is_element()
                                && (n.tag_name().name() == "numCache"
                                    || n.tag_name().name() == "numLit")
                        })
                    })
            } else {
                ser.children()
                    .find(|n| n.is_element() && n.tag_name().name() == "val")
                    .and_then(|v| {
                        v.descendants().find(|n| {
                            n.is_element()
                                && (n.tag_name().name() == "numCache"
                                    || n.tag_name().name() == "numLit")
                        })
                    })
            };

            // For scatter/bubble the point count comes from this series' xVal (each
            // series can have a different point count). For other charts it's the
            // shared category count.
            let series_pt_count = if is_scatter_like {
                series_categories
                    .as_ref()
                    .map(|c| c.len())
                    .unwrap_or(0)
                    .max(1)
            } else {
                pt_count
            };

            let mut values: Vec<Option<f64>> = vec![None; series_pt_count];
            if let Some(cache) = val_cache {
                for pt in cache
                    .children()
                    .filter(|n| n.is_element() && n.tag_name().name() == "pt")
                {
                    let idx: usize = attr(&pt, "idx").and_then(|v| v.parse().ok()).unwrap_or(0);
                    let val: Option<f64> = pt
                        .children()
                        .find(|n| n.is_element() && n.tag_name().name() == "v")
                        .and_then(|v| v.text())
                        .and_then(|t| t.parse().ok());
                    if idx < values.len() {
                        values[idx] = val;
                    }
                }
            }

            // Bubble per-point sizes (ECMA-376 §21.2.2.4 `<c:bubbleSize>`).
            // Only meaningful for bubble charts; scatter / others ignore.
            let bubble_sizes: Option<Vec<Option<f64>>> = if chart_type == "bubble" {
                let bub_cache = ser
                    .children()
                    .find(|n| n.is_element() && n.tag_name().name() == "bubbleSize")
                    .and_then(|b| {
                        b.descendants().find(|n| {
                            n.is_element()
                                && (n.tag_name().name() == "numCache"
                                    || n.tag_name().name() == "numLit")
                        })
                    });
                bub_cache.map(|cache| {
                    let mut sizes: Vec<Option<f64>> = vec![None; series_pt_count];
                    for pt in cache
                        .children()
                        .filter(|n| n.is_element() && n.tag_name().name() == "pt")
                    {
                        let idx: usize = attr(&pt, "idx").and_then(|v| v.parse().ok()).unwrap_or(0);
                        let val: Option<f64> = pt
                            .children()
                            .find(|n| n.is_element() && n.tag_name().name() == "v")
                            .and_then(|v| v.text())
                            .and_then(|t| t.parse().ok());
                        if idx < sizes.len() {
                            sizes[idx] = val;
                        }
                    }
                    sizes
                })
            } else {
                None
            };

            // Series color from spPr > solidFill (bar/area/pie) or spPr > ln > solidFill (line)
            let color = ser
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "spPr")
                .and_then(|sp| {
                    sp.children()
                        .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
                        .or_else(|| {
                            sp.children()
                                .find(|n| n.is_element() && n.tag_name().name() == "ln")
                                .and_then(|ln| {
                                    ln.children().find(|n| {
                                        n.is_element() && n.tag_name().name() == "solidFill"
                                    })
                                })
                        })
                })
                .and_then(|fill| parse_color_node(fill, theme));

            // Per-data-point colors from <c:dPt> (§21.2.2.52; important for
            // pie charts). The point index is the CHILD element `<c:idx val>`
            // (ECMA-376 §21.2.2.84, CT_UnsignedInt), not an attribute on
            // `<c:dPt>` — the old `attr(dpt, "idx")` always returned None, so
            // every slice fell
            // back to the series colour. The fill is `<c:spPr><a:solidFill>`;
            // restrict to spPr's direct child so a border `<a:ln><a:solidFill>`
            // can't be mistaken for the slice fill.
            let data_point_colors: Vec<Option<String>> = (0..series_pt_count)
                .map(|i| {
                    ser.children()
                        .filter(|n| n.is_element() && n.tag_name().name() == "dPt")
                        .find(|dpt| {
                            dpt.children()
                                .find(|n| n.is_element() && n.tag_name().name() == "idx")
                                .and_then(|n| attr(&n, "val"))
                                .and_then(|v| v.parse::<usize>().ok())
                                == Some(i)
                        })
                        .and_then(|dpt| {
                            dpt.children()
                                .find(|n| n.is_element() && n.tag_name().name() == "spPr")
                        })
                        .and_then(|sp| {
                            sp.children()
                                .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
                        })
                        .and_then(|fill| parse_color_node(fill, theme))
                })
                .collect();

            let has_dpt_colors = data_point_colors.iter().any(|c| c.is_some());

            // Per-point `<c:dPt><c:explosion>` (§21.2.2.61) — pie/doughnut slice
            // pull-out. Only the explosion is captured here (fills already flow
            // through `data_point_colors`); a dPt with no explosion yields no
            // override so the wire stays clean for the common non-exploded pie.
            let data_point_overrides: Vec<ooxml_common::chart::ChartDataPointOverride> = ser
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "dPt")
                .filter_map(|dpt| {
                    let idx = dpt
                        .children()
                        .find(|n| n.is_element() && n.tag_name().name() == "idx")
                        .and_then(|n| attr(&n, "val"))
                        .and_then(|v| v.parse::<u32>().ok())?;
                    let explosion = ooxml_common::chart::extract_dpt_explosion(dpt)?;
                    Some(ooxml_common::chart::ChartDataPointOverride {
                        idx,
                        color: None,
                        marker_symbol: None,
                        marker_size: None,
                        marker_fill: None,
                        marker_line: None,
                        explosion: Some(explosion),
                    })
                })
                .collect();

            // Series value number format from `<c:val>…<c:numCache><c:formatCode>`.
            // Used for data labels when `<c:dLbls>` carries no explicit `<c:numFmt>`
            // (ECMA-376 §21.2.2.121). "General" means "no format" → drop it so the
            // renderer's default integer/decimal formatter takes over.
            let val_format_code = val_cache
                .and_then(|cache| {
                    cache
                        .children()
                        .find(|n| n.is_element() && n.tag_name().name() == "formatCode")
                        .and_then(|fc| fc.text().map(|t| t.to_string()))
                })
                .filter(|s| !s.is_empty() && s != "General");

            // Series-level data-label text colour from `<c:dLbls><c:txPr>…solidFill`.
            // Scoped to this `<c:ser>` (not chart-root) so stacked-bar segments keep
            // their independent label colours (white on dark fill, black on light).
            let label_color = ser
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "dLbls")
                .and_then(|dlbls| {
                    dlbls
                        .children()
                        .find(|n| n.is_element() && n.tag_name().name() == "txPr")
                })
                .and_then(|txpr| {
                    txpr.descendants()
                        .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
                })
                .and_then(|fill| parse_color_node(fill, theme));

            ChartSeriesData {
                name,
                values,
                color,
                data_point_colors: if has_dpt_colors {
                    Some(data_point_colors)
                } else {
                    None
                },
                // Legacy `<c:chart>` per-point label colors are extracted via
                // `<c:dLbls><c:dLbl idx>` — not yet wired here; chartEx is the only
                // path that needs it for sample-2's waterfall.
                data_label_colors: None,
                categories: series_categories,
                bubble_sizes,
                val_format_code,
                label_color,
                series_type,
                // Shared `ChartSeries.use_secondary_axis` is `Option<bool>`; the
                // legacy default (false) is expressed as `None` so it drops off
                // the wire exactly as the old `skip_serializing_if = "Not::not"`
                // did.
                use_secondary_axis: if use_secondary_axis { Some(true) } else { None },
                // Fields the legacy `<c:chart>` path doesn't populate (marker
                // styling, per-point overrides, per-series dLbls, error bars) —
                // pptx renders these from the series color today.
                show_marker: None,
                marker_symbol: None,
                marker_size: None,
                marker_fill: None,
                marker_line: None,
                data_point_overrides: if data_point_overrides.is_empty() {
                    None
                } else {
                    Some(data_point_overrides)
                },
                data_label_overrides: None,
                series_data_labels: None,
                err_bars: None,
                // `<c:ser><c:smooth>` (§21.2.2.194) — line/area spline flag.
                // Shared with the xlsx parser via ooxml-common so both honor the
                // CT_Boolean implied-true semantics.
                smooth: ooxml_common::chart::extract_series_smooth(*ser),
            }
        })
        .collect();

    // Data labels are on when `<c:dLbls>` enables `<c:showVal>` OR
    // `<c:showPercent>` (ECMA-376 §21.2.2.189 / §21.2.2.187) — at chart level
    // or in any series. Pie/doughnut decks commonly use showPercent only (e.g.
    // sample-14 slide-7's "54%/27%/…" slice labels); the renderer draws the
    // slice percentage for pie/doughnut and the raw value for bar/line.
    let show_data_labels = root
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "dLbls")
        .any(|d_lbls| {
            d_lbls.children().any(|c| {
                c.is_element()
                    && matches!(c.tag_name().name(), "showVal" | "showPercent")
                    && attr(&c, "val").as_deref() == Some("1")
            })
        });

    // Outer chartSpace spPr: we want the child of chartSpace (not plotArea).
    let chart_bg = root
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spPr")
        .and_then(|sp| {
            sp.children()
                .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
        })
        .and_then(|fill| parse_color_node(fill, theme));

    // <c:legend> + <c:legendPos val> — shared helper.
    let (show_legend, legend_pos) = ooxml_common::chart::extract_legend(root);

    // ECMA-376 §21.2.2.35: `<c:crossBetween>` lives on the VALUE axis (not cat),
    // and describes whether value gridlines land between or on category ticks.
    // Default is "between" (categories inset by half a step each side).
    let cat_axis_cross_between = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "valAx")
        .and_then(|ax| {
            ax.children()
                .find(|n| n.is_element() && n.tag_name().name() == "crossBetween")
        })
        .and_then(|n| attr(&n, "val"))
        .unwrap_or_else(|| "between".to_string());

    // Major tick marks (ECMA-376 §21.2.2.49 ST_TickMark, default "cross").
    // Schema default is `out` (ST_TickMark §21.2.3.48), shared with xlsx via
    // the ooxml-common helper so the two parsers don't diverge on the default.
    let read_major_tick_mark = |ax: Option<roxmltree::Node>| -> String {
        ax.map(|n| ooxml_common::chart::extract_axis_tick_mark_or_default(n, "majorTickMark"))
            .unwrap_or_else(|| "out".to_string())
    };
    let val_axis_major_tick_mark = read_major_tick_mark(val_ax);
    let cat_axis_major_tick_mark = read_major_tick_mark(cat_ax);

    // Axis tick-label font size from `<c:txPr>` (in OOXML hundredths of a point).
    let cat_axis_font_size_hpt = cat_ax.and_then(ooxml_common::chart::extract_axis_tick_label_size);
    let val_axis_font_size_hpt = val_ax.and_then(ooxml_common::chart::extract_axis_tick_label_size);

    // Data-label font size — first `<c:dLbls><c:txPr>` defRPr/rPr@sz we find.
    let data_label_font_size_hpt = ooxml_common::chart::extract_data_label_font_size(root);

    // Bar gap / overlap, dLblPos and numFmt — all shared helpers so any new
    // chart property added to the xlsx side stays applied to pptx without
    // a manual port (the slide-7 / sample-2 issue this PR avoids).
    let (bar_gap_width, bar_overlap) = ooxml_common::chart::extract_bar_gap_overlap(root);
    let data_label_position = ooxml_common::chart::extract_data_label_position(root);
    let data_label_format_code = ooxml_common::chart::extract_data_label_format_code(root);

    // Data-label font color uses the shared helper too — pptx supplies a
    // ColorResolver wrapper around `parse_color_node` so the
    // ECMA-376 §21.2.2.16 dLbls > txPr > solidFill walk lives in one place.
    let resolver = PptxColorResolver { theme };
    let data_label_font_color = ooxml_common::chart::extract_data_label_font_color(root, &resolver);

    // Axis tick-label text color + axis-line style (color / width / noFill).
    // ECMA-376 §21.2.2.* — `<c:catAx|valAx><c:txPr>…<a:solidFill>` colors the
    // tick labels and `<c:spPr><a:ln>` styles the axis rule. Shared helpers so
    // the gray "2025年3月期" category labels and the light-gray category-axis
    // line in sample-2 slide-16's horizontal bar chart resolve the same way.
    let cat_axis_font_color =
        cat_ax.and_then(|n| ooxml_common::chart::extract_axis_tick_label_color(n, &resolver));
    let val_axis_font_color =
        val_ax.and_then(|n| ooxml_common::chart::extract_axis_tick_label_color(n, &resolver));
    let (cat_axis_line_color, cat_axis_line_width_emu, cat_axis_line_hidden) = cat_ax
        .map(|n| ooxml_common::chart::extract_axis_line_style(n, &resolver))
        .unwrap_or((None, None, false));
    let (val_axis_line_color, val_axis_line_width_emu, val_axis_line_hidden) = val_ax
        .map(|n| ooxml_common::chart::extract_axis_line_style(n, &resolver))
        .unwrap_or((None, None, false));

    // `<c:valAx><c:numFmt formatCode>` — value-axis tick label number format.
    let val_axis_format_code = val_ax.and_then(ooxml_common::chart::extract_axis_format_code);
    // `<c:catAx|dateAx><c:numFmt formatCode>` — category-axis number format. For
    // a `<c:dateAx>` this is the date serial format code (e.g. "m/d/yyyy") the TS
    // side needs to format category labels. Reaches parity with the xlsx parser,
    // which already wires this field (pptx previously hardcoded it to None).
    let cat_axis_format_code = cat_ax.and_then(ooxml_common::chart::extract_axis_format_code);

    // Secondary value axis (combo charts) — parse the right-hand `<c:valAx>`
    // into a self-contained spec using the same shared helpers as the primary
    // axis. None for the common single value-axis case.
    let secondary_val_axis = secondary_val_ax.map(|ax| {
        let (min, max) = ooxml_common::chart::extract_axis_min_max(ax);
        let (t, title_size, title_bold, title_color) =
            ooxml_common::chart::extract_axis_title_with_props(ax);
        let (line_color, line_width_emu, line_hidden) =
            ooxml_common::chart::extract_axis_line_style(ax, &resolver);
        SecondaryValueAxis {
            min,
            max,
            title: t,
            hidden: ooxml_common::chart::axis_is_deleted(ax),
            format_code: ooxml_common::chart::extract_axis_format_code(ax),
            font_color: ooxml_common::chart::extract_axis_tick_label_color(ax, &resolver),
            font_size_hpt: ooxml_common::chart::extract_axis_tick_label_size(ax),
            line_color,
            line_width_emu,
            line_hidden,
            major_tick_mark: ooxml_common::chart::extract_axis_tick_mark_or_default(
                ax,
                "majorTickMark",
            ),
            title_font_size_hpt: title_size,
            title_font_bold: title_bold,
            title_font_color: title_color,
        }
    });

    // `<c:plotArea><c:layout><c:manualLayout>` — explicit plot-area rectangle
    // (fractions of chart space). ECMA-376 §21.2.2.32. Sample-2 slide-16 uses
    // this to keep its horizontal bar chart from spilling into the side
    // annotation column. We parse the same shape xlsx already exposes
    // (xMode, yMode, layoutTarget, x, y, w?, h?).
    let plot_area_manual_layout = plot_area
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "layout")
        .and_then(|layout| {
            layout
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "manualLayout")
        })
        .map(|ml| {
            let mut x_mode = "edge".to_string();
            let mut y_mode = "edge".to_string();
            let mut layout_target: Option<String> = None;
            let mut x = 0.0_f64;
            let mut y = 0.0_f64;
            let mut w: Option<f64> = None;
            let mut h: Option<f64> = None;
            for ch in ml.children().filter(|n| n.is_element()) {
                let val_str = attr(&ch, "val");
                match ch.tag_name().name() {
                    "xMode" => {
                        if let Some(v) = val_str {
                            x_mode = v;
                        }
                    }
                    "yMode" => {
                        if let Some(v) = val_str {
                            y_mode = v;
                        }
                    }
                    "layoutTarget" => {
                        layout_target = val_str;
                    }
                    "x" => {
                        if let Some(v) = val_str.and_then(|s| s.parse::<f64>().ok()) {
                            x = v;
                        }
                    }
                    "y" => {
                        if let Some(v) = val_str.and_then(|s| s.parse::<f64>().ok()) {
                            y = v;
                        }
                    }
                    "w" => {
                        w = val_str.and_then(|s| s.parse::<f64>().ok());
                    }
                    "h" => {
                        h = val_str.and_then(|s| s.parse::<f64>().ok());
                    }
                    _ => {}
                }
            }
            ChartManualLayout {
                x_mode,
                y_mode,
                layout_target,
                x,
                y,
                w,
                h,
            }
        });

    // `<c:scatterChart><c:scatterStyle val>` — ECMA-376 §21.2.2.42. Lives
    // directly under scatterChart, so a plot_area descendant walk is enough.
    let scatter_style = if chart_type == "scatter" {
        plot_area
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "scatterStyle")
            .and_then(|n| attr(&n, "val"))
    } else {
        None
    };

    // Axis titles + run props (ECMA-376 §21.2.2.6 `CT_Title`). Iterate every
    // `<c:catAx>`/`<c:valAx>` so the scatter case — two `<c:valAx>`, no
    // `<c:catAx>` — resolves correctly: a `<c:valAx>` whose `<c:axPos val>` is
    // `b`/`t` is the horizontal (X) axis → cat-axis title; `l`/`r` is the
    // vertical (Y) axis → val-axis title. A real `<c:catAx>` always feeds the
    // cat-axis title. First title wins for each axis (matches the xlsx parser).
    let mut cat_axis_title: Option<String> = None;
    let mut cat_axis_title_size: Option<i32> = None;
    let mut cat_axis_title_bold: Option<bool> = None;
    let mut cat_axis_title_color: Option<String> = None;
    let mut cat_axis_title_face: Option<String> = None;
    let mut val_axis_title: Option<String> = None;
    let mut val_axis_title_size: Option<i32> = None;
    let mut val_axis_title_bold: Option<bool> = None;
    let mut val_axis_title_color: Option<String> = None;
    let mut val_axis_title_face: Option<String> = None;
    for ax in plot_area
        .children()
        .filter(|n| n.is_element() && matches!(n.tag_name().name(), "catAx" | "dateAx" | "valAx"))
    {
        let is_cat = if matches!(ax.tag_name().name(), "catAx" | "dateAx") {
            true
        } else {
            // valAx: disambiguate by axPos (b/t → X/cat, l/r → Y/val).
            let ax_pos = ax
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "axPos")
                .and_then(|n| attr(&n, "val"))
                .unwrap_or_default();
            matches!(ax_pos.as_str(), "b" | "t")
        };
        if is_cat {
            if cat_axis_title.is_none() {
                let (t, sz, b, col) = ooxml_common::chart::extract_axis_title_with_props(ax);
                if t.is_some() {
                    cat_axis_title = t;
                    cat_axis_title_size = sz;
                    cat_axis_title_bold = b;
                    cat_axis_title_color = col;
                    cat_axis_title_face = ooxml_common::chart::extract_axis_title_face(ax);
                }
            }
        } else if val_axis_title.is_none() {
            let (t, sz, b, col) = ooxml_common::chart::extract_axis_title_with_props(ax);
            if t.is_some() {
                val_axis_title = t;
                val_axis_title_size = sz;
                val_axis_title_bold = b;
                val_axis_title_color = col;
                val_axis_title_face = ooxml_common::chart::extract_axis_title_face(ax);
            }
        }
    }

    // Axis tick-label bold flags (`<c:txPr>…defRPr@b`) and the chart-title bold
    // flag (`<c:title>…defRPr@b`). These were never serialized before; wiring
    // them through reaches parity with the xlsx parser so the renderer's
    // ST_Style bold handling applies uniformly. All three come from the shared
    // ooxml-common helpers so the two parsers stay in lockstep. The chart-title
    // bold helper expects the `<c:title>`'s parent, so pass `title_node_opt`'s
    // parent (the element that holds it as a direct child).
    let cat_axis_font_bold = cat_ax.and_then(ooxml_common::chart::extract_axis_tick_label_bold);
    let val_axis_font_bold = val_ax.and_then(ooxml_common::chart::extract_axis_tick_label_bold);
    let title_font_bold = title_node_opt
        .and_then(|t| t.parent())
        .and_then(ooxml_common::chart::extract_chart_title_bold);

    // Explicit chartSpace border from `<c:chartSpace><c:spPr><a:ln>` (ECMA-376
    // §21.2.2.5 / DrawingML §20.1.2.2.24). Shared with the xlsx parser via
    // `ooxml_common::chart::extract_chart_space_border` so the locked policy
    // (border only on an explicit paintable line; `<a:noFill/>` → color None;
    // srgb inside solidFill → hex; `@w` captured as u32; schemeClr unresolved)
    // stays in lockstep. `root` is the `<c:chartSpace>` element here.
    let (chart_border_color, chart_border_width_emu) =
        ooxml_common::chart::extract_chart_space_border(root);

    // `<c:date1904>` (ECMA-376 §21.2.2.38) — direct child of `<c:chartSpace>`
    // (`root`). Shared with the xlsx parser via ooxml-common so both honor the
    // CT_Boolean implied-true semantics.
    let date1904 = ooxml_common::chart::extract_chart_date1904(root);

    // `<c:chart><c:dispBlanksAs>` (ECMA-376 §21.2.2.42) — null-cell plotting for
    // line/area. Shared with the xlsx parser via ooxml-common.
    let disp_blanks_as = ooxml_common::chart::extract_disp_blanks_as(root);

    // ── Chart text font faces (CH10) ────────────────────────────────────────
    // Tick-label faces (`<c:catAx|valAx><c:txPr>…<a:latin>`), data-label face
    // (`<c:dLbls><c:txPr>…<a:latin>`) and legend text props, all via the shared
    // ooxml-common extractors so pptx/xlsx stay in lockstep. Absent faces stay
    // None; the renderer falls back to the theme body/heading font.
    let cat_axis_font_face = cat_ax.and_then(ooxml_common::chart::extract_axis_tick_label_face);
    let val_axis_font_face = val_ax.and_then(ooxml_common::chart::extract_axis_tick_label_face);
    let data_label_font_face = ooxml_common::chart::extract_data_label_face(root);
    let (legend_font_face, legend_font_size_hpt, legend_font_bold) =
        ooxml_common::chart::extract_legend_text_props(root);
    let legend_font_color = {
        let resolver = PptxColorResolver { theme };
        ooxml_common::chart::extract_legend_font_color(root, &resolver)
    };
    // Theme fallback fonts: pptx stores the theme's major/minor Latin faces in
    // the `+mj-lt` / `+mn-lt` keys of the color+font map (see lib.rs
    // parse_theme_colors). None when the theme lacks a fontScheme.
    let theme_major_font_latin = theme.get("+mj-lt").cloned();
    let theme_minor_font_latin = theme.get("+mn-lt").cloned();

    // ── Pie / doughnut geometry (CH8) ───────────────────────────────────────
    // holeSize (doughnut) / firstSliceAng (pie + doughnut), shared extractors.
    let hole_size = ooxml_common::chart::extract_hole_size(root);
    let first_slice_angle = ooxml_common::chart::extract_first_slice_angle(root);

    // ── Axis scale model (CH6) ──────────────────────────────────────────────
    // Gridline presence, manual major/minor units, log scale and orientation —
    // all via the shared ooxml-common extractors on the primary val/cat axes.
    // `<c:majorGridlines>` presence: Office writes it on the value axis by
    // default (renderer keeps its historical always-on when the field is None),
    // so we only emit `Some(false)` when a value axis EXISTS without the element.
    let val_axis_major_gridlines =
        val_ax.map(|ax| ooxml_common::chart::axis_has_major_gridlines(ax));
    let cat_axis_major_gridlines =
        cat_ax.map(|ax| ooxml_common::chart::axis_has_major_gridlines(ax));
    let val_axis_minor_gridlines =
        val_ax.map(|ax| ooxml_common::chart::axis_has_minor_gridlines(ax));
    let val_axis_major_unit = val_ax.and_then(ooxml_common::chart::extract_axis_major_unit);
    let val_axis_minor_unit = val_ax.and_then(ooxml_common::chart::extract_axis_minor_unit);
    let val_axis_log_base = val_ax.and_then(ooxml_common::chart::extract_axis_log_base);
    let val_axis_orientation = val_ax.and_then(ooxml_common::chart::extract_axis_orientation);
    let cat_axis_orientation = cat_ax.and_then(ooxml_common::chart::extract_axis_orientation);
    let cat_axis_tick_label_pos = cat_ax.and_then(ooxml_common::chart::extract_axis_tick_label_pos);
    let val_axis_tick_label_pos = val_ax.and_then(ooxml_common::chart::extract_axis_tick_label_pos);
    let cat_axis_label_rotation =
        cat_ax.and_then(ooxml_common::chart::extract_axis_tick_label_rotation);

    // Chart title font face (`<c:title>…<a:latin>`) — parity with xlsx, which
    // already extracts it. `extract_axis_title_face` scopes to a node's
    // direct-child `<c:title>`, so pass the title's parent (`<c:chart>`).
    let title_font_face = title_node_opt
        .and_then(|t| t.parent())
        .and_then(ooxml_common::chart::extract_axis_title_face);

    Some(ChartElement {
        x: 0,
        y: 0,
        width: 0,
        height: 0,
        chart: ChartModel {
            chart_type,
            title,
            categories,
            series,
            val_max,
            val_min,
            subtotal_indices: vec![],
            show_data_labels,
            cat_axis_hidden,
            val_axis_hidden,
            plot_area_bg,
            chart_bg,
            show_legend,
            cat_axis_cross_between,
            val_axis_major_tick_mark,
            cat_axis_major_tick_mark,
            title_font_size_hpt,
            title_font_color: None,
            title_font_face,
            cat_axis_font_size_hpt,
            val_axis_font_size_hpt,
            cat_axis_font_color,
            val_axis_font_color,
            cat_axis_line_color,
            cat_axis_line_width_emu,
            cat_axis_line_hidden,
            val_axis_line_color,
            val_axis_line_width_emu,
            val_axis_line_hidden,
            data_label_font_size_hpt,
            legend_pos,
            bar_gap_width,
            bar_overlap,
            data_label_position,
            data_label_font_color,
            data_label_format_code,
            val_axis_format_code,
            plot_area_manual_layout,
            scatter_style,
            cat_axis_title,
            val_axis_title,
            // TS `ChartElement` renamed the axis-title run-prop fields to the
            // core `ChartModel` names (`…TitleFontSizeHpt/Bold/Color`); the
            // parser locals keep the shorter legacy names.
            cat_axis_title_font_size_hpt: cat_axis_title_size,
            cat_axis_title_font_bold: cat_axis_title_bold,
            cat_axis_title_font_color: cat_axis_title_color,
            val_axis_title_font_size_hpt: val_axis_title_size,
            val_axis_title_font_bold: val_axis_title_bold,
            val_axis_title_font_color: val_axis_title_color,
            title_font_bold,
            cat_axis_font_bold,
            val_axis_font_bold,
            chart_border_color,
            chart_border_width_emu,
            secondary_val_axis,
            // Pie/doughnut geometry (CH8) + chart text font faces (CH10).
            hole_size,
            first_slice_angle,
            cat_axis_font_face,
            val_axis_font_face,
            cat_axis_title_font_face: cat_axis_title_face,
            val_axis_title_font_face: val_axis_title_face,
            data_label_font_face,
            legend_font_face,
            legend_font_color,
            legend_font_size_hpt,
            legend_font_bold,
            theme_major_font_latin,
            theme_minor_font_latin,
            // ChartModel fields the legacy pptx `<c:chart>` path leaves unset
            // (they were never in the pptx `ChartElement` copy, so they defaulted
            // to `undefined` on the TS side and stay absent on the wire).
            val_axis_minor_tick_mark: None,
            cat_axis_minor_tick_mark: None,
            legend_manual_layout: None,
            title_manual_layout: None,
            cat_axis_crosses: None,
            cat_axis_crosses_at: None,
            val_axis_crosses: None,
            val_axis_crosses_at: None,
            cat_axis_format_code,
            cat_axis_min: None,
            cat_axis_max: None,
            radar_style: None,
            date1904,
            disp_blanks_as,
            // ── Axis scale model (CH6) ──────────────────────────────────────
            val_axis_major_gridlines,
            cat_axis_major_gridlines,
            val_axis_minor_gridlines,
            val_axis_major_unit,
            val_axis_minor_unit,
            val_axis_log_base,
            val_axis_orientation,
            cat_axis_orientation,
            cat_axis_tick_label_pos,
            val_axis_tick_label_pos,
            cat_axis_label_rotation,
        },
    })
}

/// Parse a modern chartEx (cx: namespace) — waterfall, treemap, etc.
pub(crate) fn parse_chartex(xml: &str, theme: &HashMap<String, String>) -> Option<ChartElement> {
    let doc = roxmltree::Document::parse(xml).ok()?;
    let root = doc.root_element();

    // Chart type from series layoutId attribute
    let series_node = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "series")?;
    let layout_id = attr(&series_node, "layoutId").unwrap_or_default();
    let chart_type = layout_id; // "waterfall", "treemap", etc.

    // Categories from chartData > data > strDim[@type="cat"] > lvl > pt
    let categories: Vec<String> = root
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "strDim"
                && attr(n, "type").as_deref() == Some("cat")
        })
        .and_then(|dim| {
            dim.descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "lvl")
        })
        .map(|lvl| {
            lvl.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "pt")
                .filter_map(|pt| pt.text().map(|t| t.replace('\n', " ")))
                .collect()
        })
        .unwrap_or_default();

    let pt_count = categories.len().max(1);

    // Values from chartData > data > numDim[@type="val"] > lvl > pt
    let raw_values: Vec<Option<f64>> = root
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "numDim"
                && attr(n, "type").as_deref() == Some("val")
        })
        .and_then(|dim| {
            dim.descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "lvl")
        })
        .map(|lvl| {
            let mut vals: Vec<Option<f64>> = vec![None; pt_count];
            for (i, pt) in lvl
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "pt")
                .enumerate()
            {
                if i < vals.len() {
                    vals[i] = pt.text().and_then(|t| t.parse().ok());
                }
            }
            vals
        })
        .unwrap_or_else(|| vec![None; pt_count]);

    // Subtotal indices (idx=0 is always implicit; add from cx:subtotals)
    let mut subtotal_indices: Vec<u32> = vec![0];
    if let Some(subtotals_node) = series_node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "subtotals")
    {
        for idx_node in subtotals_node
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "idx")
        {
            if let Some(v) = attr(&idx_node, "val").and_then(|v| v.parse::<u32>().ok()) {
                if v != 0 {
                    subtotal_indices.push(v);
                }
            }
        }
    }

    // Series color (first dataPt or series spPr)
    let color = series_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spPr")
        .and_then(|sp| {
            sp.children()
                .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
        })
        .and_then(|fill| parse_color_node(fill, theme));

    // Per-idx data-label colors. ChartEx writes `<cx:dataLabels>` with
    // `<cx:dataLabel idx="N">` overrides; each carries its own `<cx:txPr>`
    // whose first `<a:solidFill>` is the label colour for that bar. Sample-2
    // waterfall uses this to paint negative-bar labels in accent1 (red) while
    // positive-bar labels stay tx1 (black).
    let mut data_label_colors_vec: Vec<Option<String>> = vec![None; raw_values.len().max(1)];
    let mut has_per_label_color = false;
    for dl in series_node
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "dataLabel")
    {
        let Some(idx) = attr(&dl, "idx").and_then(|v| v.parse::<usize>().ok()) else {
            continue;
        };
        if idx >= data_label_colors_vec.len() {
            continue;
        }
        // First `<a:solidFill>` inside the per-idx <cx:txPr>.
        let txpr = match dl
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "txPr")
        {
            Some(n) => n,
            None => continue,
        };
        for desc in txpr.descendants().filter(|n| n.is_element()) {
            if desc.tag_name().name() != "solidFill" {
                continue;
            }
            if let Some(c) = parse_color_node(desc, theme) {
                data_label_colors_vec[idx] = Some(c);
                has_per_label_color = true;
                break;
            }
        }
    }

    let series = vec![ChartSeriesData {
        name: String::new(),
        values: raw_values,
        color,
        data_point_colors: None,
        data_label_colors: if has_per_label_color {
            Some(data_label_colors_vec)
        } else {
            None
        },
        categories: None,
        bubble_sizes: None,
        val_format_code: None,
        label_color: None,
        series_type: None,
        use_secondary_axis: None,
        show_marker: None,
        marker_symbol: None,
        marker_size: None,
        marker_fill: None,
        marker_line: None,
        data_point_overrides: None,
        data_label_overrides: None,
        series_data_labels: None,
        err_bars: None,
        // chartEx (waterfall) has no `<c:smooth>` concept.
        smooth: None,
    }];

    // ChartEx axis visibility — shared helper that pairs each `<cx:axis hidden>`
    // with its `<cx:catScaling>` / `<cx:valScaling>` child to disambiguate cat
    // vs. val (chartEx doesn't declare axis kind via the `id` attribute).
    let (cat_axis_hidden, val_axis_hidden) = ooxml_common::chart::extract_chartex_axis_hidden(root);

    // `<cx:catScaling gapWidth>` (chartEx) — same semantics as legacy
    // `<c:gapWidth>` but stored as a *fraction* (e.g. 0.8 ≡ 80%) instead of
    // an integer percentage. Convert to the legacy percentage form so the
    // shared renderer's `barW = catGap / (1 + gapWidth/100)` formula works
    // uniformly across chart types. Default 1.5 (= legacy 150%) per PowerPoint
    // when the attribute is omitted.
    let bar_gap_width = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "catScaling")
        .and_then(|n| attr(&n, "gapWidth"))
        .and_then(|v| v.parse::<f64>().ok())
        .map(|frac| (frac * 100.0).round() as i32);

    Some(ChartElement {
        x: 0,
        y: 0,
        width: 0,
        height: 0,
        chart: ChartModel {
            chart_type,
            title: None,
            categories,
            series,
            val_max: None,
            val_min: None,
            subtotal_indices,
            show_data_labels: false,
            cat_axis_hidden,
            val_axis_hidden,
            plot_area_bg: None,
            chart_bg: None,
            show_legend: false,
            cat_axis_cross_between: "between".to_string(),
            val_axis_major_tick_mark: "cross".to_string(),
            cat_axis_major_tick_mark: "cross".to_string(),
            title_font_size_hpt: None,
            title_font_color: None,
            title_font_face: None,
            cat_axis_font_size_hpt: None,
            val_axis_font_size_hpt: None,
            cat_axis_font_color: None,
            val_axis_font_color: None,
            cat_axis_line_color: None,
            cat_axis_line_width_emu: None,
            cat_axis_line_hidden: false,
            val_axis_line_color: None,
            val_axis_line_width_emu: None,
            val_axis_line_hidden: false,
            data_label_font_size_hpt: None,
            legend_pos: None,
            bar_gap_width,
            bar_overlap: None,
            data_label_position: None,
            data_label_font_color: None,
            data_label_format_code: None,
            val_axis_format_code: None,
            plot_area_manual_layout: None,
            scatter_style: None,
            // chartEx (waterfall/treemap/etc.) has its own axis model and is not
            // wired for axis titles or an explicit chartSpace border yet.
            cat_axis_title: None,
            val_axis_title: None,
            cat_axis_title_font_size_hpt: None,
            cat_axis_title_font_bold: None,
            cat_axis_title_font_color: None,
            val_axis_title_font_size_hpt: None,
            val_axis_title_font_bold: None,
            val_axis_title_font_color: None,
            title_font_bold: None,
            cat_axis_font_bold: None,
            val_axis_font_bold: None,
            chart_border_color: None,
            chart_border_width_emu: None,
            secondary_val_axis: None,
            // chartEx charts (waterfall/treemap/etc.) are not pie/doughnut and
            // don't carry `<c:txPr>` axis/legend faces; only the theme fallback
            // fonts are threaded so their data labels can pick up the body font.
            hole_size: None,
            first_slice_angle: None,
            cat_axis_font_face: None,
            val_axis_font_face: None,
            cat_axis_title_font_face: None,
            val_axis_title_font_face: None,
            data_label_font_face: None,
            legend_font_face: None,
            legend_font_color: None,
            legend_font_size_hpt: None,
            legend_font_bold: None,
            theme_major_font_latin: theme.get("+mj-lt").cloned(),
            theme_minor_font_latin: theme.get("+mn-lt").cloned(),
            val_axis_minor_tick_mark: None,
            cat_axis_minor_tick_mark: None,
            legend_manual_layout: None,
            title_manual_layout: None,
            cat_axis_crosses: None,
            cat_axis_crosses_at: None,
            val_axis_crosses: None,
            val_axis_crosses_at: None,
            cat_axis_format_code: None,
            cat_axis_min: None,
            cat_axis_max: None,
            radar_style: None,
            // chartEx (cx: namespace) has its own date-axis model; the legacy
            // `<c:date1904>` element does not apply here, so keep the 1900
            // default until/unless a chartEx date system is wired.
            date1904: false,
            // chartEx waterfall has no line/area blanks to display.
            disp_blanks_as: None,
            // chartEx (cx:) has its own axis model (`<cx:axis>`); the classic
            // `<c:catAx>`/`<c:valAx>` scale properties don't apply, so leave the
            // CH6 fields unset — the renderer keeps its defaults (byte-stable).
            val_axis_major_gridlines: None,
            cat_axis_major_gridlines: None,
            val_axis_minor_gridlines: None,
            val_axis_major_unit: None,
            val_axis_minor_unit: None,
            val_axis_log_base: None,
            val_axis_orientation: None,
            cat_axis_orientation: None,
            cat_axis_tick_label_pos: None,
            val_axis_tick_label_pos: None,
            cat_axis_label_rotation: None,
        },
    })
}
