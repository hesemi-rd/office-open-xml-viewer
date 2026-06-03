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

/// Find a direct child of `parent` whose local name is `name`.
fn child<'a, 'i>(parent: Node<'a, 'i>, name: &str) -> Option<Node<'a, 'i>> {
    parent.children().find(|n| n.is_element() && n.tag_name().name() == name)
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
    let legend = root.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "legend");
    let show = legend.is_some();
    let pos = legend.and_then(|ln| {
        child(ln, "legendPos").and_then(|p| p.attribute("val")).map(|s| s.to_string())
    });
    (show, pos)
}

/// `<c:barChart><c:gapWidth val>` / `<c:overlap val>` (ECMA-376 §21.2.2.13,
/// §21.2.2.25). Returns `(gap%, overlap%)`. Defaults to (None, None) when
/// the file relies on Office's defaults (gap 150, overlap 0).
pub fn extract_bar_gap_overlap(root: Node) -> (Option<i32>, Option<i32>) {
    let gap = root.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "gapWidth")
        .and_then(|n| n.attribute("val").and_then(|v| v.parse::<i32>().ok()));
    let ov = root.descendants()
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
            child(dlbls, "dLblPos").and_then(|n| n.attribute("val")).map(|s| s.to_string())
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
/// are the ECMA-376 §21.2.2.49 ST_TickMark enum: `none` | `out` | `in` |
/// `cross`. Returns the raw string.
pub fn extract_axis_tick_mark(axis_node: Node, name: &str) -> Option<String> {
    child(axis_node, name)
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string())
}

/// First `<a:defRPr@sz>` or `<a:rPr@sz>` found inside the axis's `<c:txPr>`.
/// Sizes are OOXML hundredths of a point (e.g. 1200 = 12 pt).
pub fn extract_axis_tick_label_size(axis_node: Node) -> Option<i32> {
    let txpr = child(axis_node, "txPr")?;
    txpr.descendants().find_map(|n| {
        if !n.is_element() { return None; }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" { return None; }
        n.attribute("sz").and_then(|v| v.parse::<i32>().ok())
    })
}

/// First `<c:dLbls><c:txPr>` font size (hpt). Mirrors the per-series + chart
/// fallback chain: walk every `<c:dLbls>` in document order, returning the
/// first inner `<a:defRPr@sz>` / `<a:rPr@sz>` we find.
pub fn extract_data_label_font_size(root: Node) -> Option<i32> {
    root.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "dLbls")
        .find_map(|dl| {
            child(dl, "txPr").and_then(|tx| tx.descendants().find_map(|n| {
                if !n.is_element() { return None; }
                let tag = n.tag_name().name();
                if tag != "defRPr" && tag != "rPr" { return None; }
                n.attribute("sz").and_then(|v| v.parse::<i32>().ok())
            }))
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
    for dlbls in root.descendants().filter(|n| n.is_element() && n.tag_name().name() == "dLbls") {
        let Some(txpr) = child(dlbls, "txPr") else { continue; };
        for desc in txpr.descendants().filter(|n| n.is_element()) {
            if desc.tag_name().name() != "solidFill" { continue; }
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
pub fn extract_axis_tick_label_color(axis_node: Node, resolver: &dyn ColorResolver) -> Option<String> {
    let txpr = child(axis_node, "txPr")?;
    for desc in txpr.descendants().filter(|n| n.is_element()) {
        if desc.tag_name().name() != "solidFill" { continue; }
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
    let Some(sp_pr) = child(axis_node, "spPr") else { return (None, None, false); };
    let Some(ln) = child(sp_pr, "ln") else { return (None, None, false); };
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
    for ax in root.descendants().filter(|n| n.is_element() && n.tag_name().name() == "axis") {
        let hidden = ax.attribute("hidden").map(|v| v == "1").unwrap_or(false);
        if !hidden { continue; }
        let is_val = ax.children().any(|c| c.is_element() && c.tag_name().name() == "valScaling");
        let is_cat = ax.children().any(|c| c.is_element() && c.tag_name().name() == "catScaling");
        if is_val { val_hidden = true; }
        if is_cat { cat_hidden = true; }
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
        let xml = r#"<c:barChart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
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
        assert_eq!(extract_bar_gap_overlap(d.root_element()), (Some(50), Some(100)));
    }

    #[test]
    fn data_label_position() {
        let xml = r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:plotArea><c:dLbls><c:dLblPos val="ctr"/></c:dLbls></c:plotArea>
        </c:chart>"#;
        let d = root_of(xml);
        assert_eq!(extract_data_label_position(d.root_element()).as_deref(), Some("ctr"));
    }

    #[test]
    fn axis_delete_truthy_variants() {
        for (val, expect) in [("1", true), ("0", false), ("true", true), ("false", false), ("True", true)] {
            let xml = format!(r#"<c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:delete val="{val}"/>
            </c:valAx>"#);
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
        assert_eq!(extract_axis_min_max(d.root_element()), (Some(0.0), Some(2500.0)));
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
        assert_eq!(extract_axis_format_code(d.root_element()).as_deref(), Some("0.0%"));
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
        assert!(got.is_none(), "spPr fill must not leak into the font color: got {got:?}");
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
        assert_eq!(extract_axis_line_style(d.root_element(), &StubResolver), (None, None, false));
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
}
