//! Shared DrawingML non-visual-properties helpers (`cNvPr` / `docPr`).
//!
//! Every OOXML host embeds the same DrawingML non-visual drawing properties
//! (ECMA-376 §20.1.2.2.8 `CT_NonVisualDrawingProps`, exposed as `<*:cNvPr>` in
//! pptx/xlsx drawings and as `<wp:docPr>` on a WordprocessingML `<wp:inline>` /
//! `<wp:anchor>`). Its `hidden` attribute (`xsd:boolean`, default `false`)
//! marks a drawing object that must **not** be rendered. This predicate is the
//! single source of truth for reading that flag across the three parsers.

use roxmltree::Node;

/// Parse an `xsd:boolean` attribute value. Per the W3C XML Schema lexical
/// space, the four valid literals are `true` / `false` / `1` / `0`; any other
/// text is not a valid boolean and yields `None` (callers apply the schema
/// default themselves). Surrounding whitespace is tolerated.
pub fn parse_xsd_bool(value: &str) -> Option<bool> {
    match value.trim() {
        "true" | "1" => Some(true),
        "false" | "0" => Some(false),
        _ => None,
    }
}

/// True when a DrawingML non-visual-properties node
/// (`<*:cNvPr>` or `<wp:docPr>`) carries `hidden` set to a truthy
/// `xsd:boolean` (§20.1.2.2.8 / §20.4.2.5). Absent or `false`/`0` → not hidden
/// (the schema default). A hidden drawing object is not rendered.
pub fn nv_props_hidden(nv_props: Node) -> bool {
    nv_props
        .attribute("hidden")
        .and_then(parse_xsd_bool)
        .unwrap_or(false)
}

#[cfg(test)]
mod tests {
    use super::*;
    use roxmltree::Document;

    #[test]
    fn parse_xsd_bool_accepts_all_four_literals() {
        assert_eq!(parse_xsd_bool("true"), Some(true));
        assert_eq!(parse_xsd_bool("1"), Some(true));
        assert_eq!(parse_xsd_bool("false"), Some(false));
        assert_eq!(parse_xsd_bool("0"), Some(false));
        // Whitespace tolerated.
        assert_eq!(parse_xsd_bool("  1 "), Some(true));
        // Anything else is not a valid boolean.
        assert_eq!(parse_xsd_bool("yes"), None);
        assert_eq!(parse_xsd_bool(""), None);
        assert_eq!(parse_xsd_bool("TRUE"), None); // case-sensitive per XSD
    }

    fn node_from(xml: &str) -> Document<'_> {
        Document::parse(xml).unwrap()
    }

    #[test]
    fn nv_props_hidden_reads_boolean_default_false() {
        // hidden="1" and hidden="true" → hidden.
        for attr in ["1", "true"] {
            let xml = format!(r#"<cNvPr id="2" name="x" hidden="{attr}"/>"#);
            let doc = node_from(&xml);
            assert!(nv_props_hidden(doc.root_element()), "hidden={attr}");
        }
        // hidden="0" / "false" / absent → not hidden.
        for xml in [
            r#"<cNvPr id="2" name="x" hidden="0"/>"#,
            r#"<cNvPr id="2" name="x" hidden="false"/>"#,
            r#"<cNvPr id="2" name="x"/>"#,
        ] {
            let doc = node_from(xml);
            assert!(!nv_props_hidden(doc.root_element()), "xml={xml}");
        }
    }
}
