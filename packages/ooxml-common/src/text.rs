//! Shared DrawingML text-body helpers used by the pptx and xlsx parsers.
//!
//! Both hosts embed the same DrawingML text grammar: a paragraph's line spacing
//! rides in `<a:lnSpc>` (ECMA-376 §21.1.2.2.5) and a text body's autofit mode in
//! a `<a:bodyPr>` child (`<a:spAutoFit>` / `<a:normAutofit>` / `<a:noAutofit>`,
//! §21.1.2.1.1-.4). These leaves were previously read inline in each parser with
//! byte-identical serde shapes; sharing the type + leaf parse keeps the two
//! formats' line-spacing / autofit handling identical.
//!
//! Following the `parse_src_rect` precedent in [`crate::blip`], each caller
//! *locates* the node (`<a:lnSpc>` / `<a:bodyPr>`) and keeps its own inheritance
//! and defaults; the shared function only parses the located leaf.

use roxmltree::Node;
use serde::{Deserialize, Serialize};

/// Paragraph line spacing (`<a:lnSpc>`, ECMA-376 §21.1.2.2.5): a percentage of
/// the natural single line, or an absolute per-line height in points. The serde
/// shape (`{"type":"pct","val":..}` / `{"type":"pts","val":..}`) mirrors core's
/// TS `SpaceLine`, so the pptx and xlsx JSON stays byte-identical.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum SpaceLine {
    /// `<a:spcPct@val>` (e.g. 100000 = 100%, 150000 = 150%).
    Pct { val: f64 },
    /// `<a:spcPts@val>` in points (the raw ST_TextSpacingPoint hundredths-of-a-
    /// point value is divided by 100 by this parser, matching core / pptx).
    Pts { val: f64 },
}

/// Parse a located `<a:lnSpc>` node (ECMA-376 §21.1.2.2.5). A `<a:spcPct@val>`
/// child yields the raw percentage (e.g. `150000`); otherwise a `<a:spcPts@val>`
/// child yields points (its raw hundredths-of-a-point `@val` divided by 100).
/// Returns `None` when neither child carries a parseable `@val`. The caller
/// passes the located `<a:lnSpc>` node and keeps its own inheritance
/// (lstStyle / master / body defaults).
pub fn parse_lnspc(ln_spc: Node<'_, '_>) -> Option<SpaceLine> {
    if let Some(v) = ln_spc
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spcPct")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok())
    {
        return Some(SpaceLine::Pct { val: v });
    }
    ln_spc
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spcPts")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok())
        .map(|v| SpaceLine::Pts { val: v / 100.0 })
}

/// Parse a located `<a:bodyPr>` node's autofit child (ECMA-376 §21.1.2.1.1-.4).
/// Returns `None` when the `<a:bodyPr>` has *no* autofit child, so the caller
/// applies its own default (pptx defers to the theme txDef; xlsx uses `none`).
/// Otherwise returns `Some((auto_fit, font_scale, ln_spc_reduction))`:
///
/// - `<a:spAutoFit>` → `("sp", None, None)`
/// - `<a:normAutofit fontScale? lnSpcReduction?>` → `("norm", fontScale?, lnSpcReduction?)`,
///   each scale being the raw ST_Percentage (1000ths of a percent) divided by
///   100000 to a fraction (e.g. `62500` → `0.625`)
/// - `<a:noAutofit>` → `("none", None, None)`
///
/// The OOXML spelling is `normAutofit` (lowercase `f`).
pub fn parse_autofit(body_pr: Node<'_, '_>) -> Option<(String, Option<f64>, Option<f64>)> {
    for c in body_pr.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "spAutoFit" => return Some(("sp".to_owned(), None, None)),
            "normAutofit" => {
                let font_scale = c
                    .attribute("fontScale")
                    .and_then(|v| v.parse::<f64>().ok())
                    .map(|v| v / 100_000.0);
                let ln_spc_reduction = c
                    .attribute("lnSpcReduction")
                    .and_then(|v| v.parse::<f64>().ok())
                    .map(|v| v / 100_000.0);
                return Some(("norm".to_owned(), font_scale, ln_spc_reduction));
            }
            "noAutofit" => return Some(("none".to_owned(), None, None)),
            _ => {}
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;
    use roxmltree::Document;

    const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// `<a:spcPct@val>` → a percent SpaceLine carrying the raw `@val`.
    #[test]
    fn parse_lnspc_reads_pct_raw_val() {
        let xml = format!(r#"<a:lnSpc xmlns:a="{A_NS}"><a:spcPct val="150000"/></a:lnSpc>"#);
        let doc = Document::parse(&xml).unwrap();
        assert_eq!(
            parse_lnspc(doc.root_element()),
            Some(SpaceLine::Pct { val: 150000.0 })
        );
    }

    /// `<a:spcPts@val>` → a points SpaceLine (raw hundredths of a point / 100).
    #[test]
    fn parse_lnspc_reads_pts_hundredths_of_point() {
        let xml = format!(r#"<a:lnSpc xmlns:a="{A_NS}"><a:spcPts val="1800"/></a:lnSpc>"#);
        let doc = Document::parse(&xml).unwrap();
        // 1800 hundredths of a point → 18 pt.
        assert_eq!(
            parse_lnspc(doc.root_element()),
            Some(SpaceLine::Pts { val: 18.0 })
        );
    }

    /// spcPct takes precedence over spcPts, and an empty/absent-val lnSpc → None.
    #[test]
    fn parse_lnspc_none_when_no_parseable_child() {
        let empty = format!(r#"<a:lnSpc xmlns:a="{A_NS}"/>"#);
        assert!(parse_lnspc(Document::parse(&empty).unwrap().root_element()).is_none());
        // spcPct with no @val falls through and (here) so does absent spcPts.
        let no_val = format!(r#"<a:lnSpc xmlns:a="{A_NS}"><a:spcPct/></a:lnSpc>"#);
        assert!(parse_lnspc(Document::parse(&no_val).unwrap().root_element()).is_none());
    }

    /// The serde shape matches the two enums it replaces: tag "type", camelCase
    /// variant tags, plain `val`.
    #[test]
    fn space_line_serializes_with_type_tag_and_camelcase() {
        let pct = serde_json::to_value(SpaceLine::Pct { val: 150000.0 }).unwrap();
        assert_eq!(pct["type"], "pct");
        assert_eq!(pct["val"], 150000.0);
        let pts = serde_json::to_value(SpaceLine::Pts { val: 18.0 }).unwrap();
        assert_eq!(pts["type"], "pts");
        assert_eq!(pts["val"], 18.0);
    }

    /// `<a:normAutofit fontScale lnSpcReduction>` → ("norm", scale/100000,
    /// reduction/100000).
    #[test]
    fn parse_autofit_normautofit_reads_scales() {
        let xml = format!(
            r#"<a:bodyPr xmlns:a="{A_NS}"><a:normAutofit fontScale="62500" lnSpcReduction="20000"/></a:bodyPr>"#
        );
        let doc = Document::parse(&xml).unwrap();
        assert_eq!(
            parse_autofit(doc.root_element()),
            Some(("norm".to_owned(), Some(0.625), Some(0.20)))
        );
    }

    /// `<a:normAutofit>` with no scale attributes → ("norm", None, None).
    #[test]
    fn parse_autofit_normautofit_without_scales() {
        let xml = format!(r#"<a:bodyPr xmlns:a="{A_NS}"><a:normAutofit/></a:bodyPr>"#);
        let doc = Document::parse(&xml).unwrap();
        assert_eq!(
            parse_autofit(doc.root_element()),
            Some(("norm".to_owned(), None, None))
        );
    }

    /// `<a:spAutoFit>` → ("sp", None, None).
    #[test]
    fn parse_autofit_spautofit() {
        let xml = format!(r#"<a:bodyPr xmlns:a="{A_NS}"><a:spAutoFit/></a:bodyPr>"#);
        let doc = Document::parse(&xml).unwrap();
        assert_eq!(
            parse_autofit(doc.root_element()),
            Some(("sp".to_owned(), None, None))
        );
    }

    /// `<a:noAutofit>` → ("none", None, None).
    #[test]
    fn parse_autofit_noautofit() {
        let xml = format!(r#"<a:bodyPr xmlns:a="{A_NS}"><a:noAutofit/></a:bodyPr>"#);
        let doc = Document::parse(&xml).unwrap();
        assert_eq!(
            parse_autofit(doc.root_element()),
            Some(("none".to_owned(), None, None))
        );
    }

    /// A `<a:bodyPr>` with NO autofit child → None, so the caller applies its own
    /// default (pptx: theme txDef; xlsx: "none").
    #[test]
    fn parse_autofit_none_when_no_child() {
        let xml = format!(r#"<a:bodyPr xmlns:a="{A_NS}" anchor="ctr" wrap="square"/>"#);
        let doc = Document::parse(&xml).unwrap();
        assert!(parse_autofit(doc.root_element()).is_none());
    }
}
