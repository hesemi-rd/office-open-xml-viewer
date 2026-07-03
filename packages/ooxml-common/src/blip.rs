//! Shared DrawingML blip helpers used by the docx, pptx and xlsx parsers.
//!
//! Every OOXML host references embedded media through the same DrawingML
//! `<a:blip>` element (ECMA-376 §20.1.8.13). A blip's raster image rides in its
//! `@r:embed` relationship; when the picture is a vector graphic, Microsoft's
//! 2016 SVG extension (MS-ODRAWXML) nests an `<asvg:svgBlip r:embed="…">` inside
//! the blip's `<a:extLst>`, and the raster `@r:embed` becomes a *fallback* for
//! SVG-incapable clients. These helpers were previously re-implemented (or, for
//! the SVG extension, only implemented in pptx) per parser; sharing them keeps
//! the three formats' blip handling identical.

use crate::ns::relationships;
use roxmltree::Node;
use serde::{Deserialize, Serialize};

/// Transitional relationships namespace (`r:`) — where a blip's `embed` / `link`
/// rIds live. Runtime lookups accept the Strict URI too (see [`blip_embed_rid`]);
/// this constant is kept for the crate's test fixtures.
pub const R_NS: &str = relationships::TRANSITIONAL;

/// Microsoft 2016 SVG extension URI (MS-ODRAWXML). An `<a:blip>` wrapping a
/// vector image carries `<a:extLst><a:ext uri="{96DAC541-…}"><asvg:svgBlip
/// r:embed="…"/></a:ext></a:extLst>`.
pub const SVG_BLIP_EXT_URI: &str = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";

/// Resolve a blip-like node's `r:embed` relationship id (the raster image, or
/// an `svgBlip`'s vector target). Reads the `embed` attribute in the
/// relationships namespace — Transitional or Strict (ISO/IEC 29500) — tolerating
/// the literal `r:embed` form some producers emit without binding the namespace.
pub fn blip_embed_rid(node: &Node<'_, '_>) -> Option<String> {
    node.attribute((relationships::TRANSITIONAL, "embed"))
        .or_else(|| node.attribute((relationships::STRICT, "embed")))
        .or_else(|| node.attribute("r:embed"))
        .map(str::to_string)
}

/// Resolve the relationship id of the vector original from an `<a:blip>`'s
/// `asvg:svgBlip` extension. Matching is by namespace-local element name
/// (`svgBlip`), so the producer's prefix (`asvg:` etc.) does not matter. Returns
/// `None` when the blip carries no svgBlip extension (the common, raster-only
/// case).
pub fn svg_blip_rid(blip: Node<'_, '_>) -> Option<String> {
    blip.children()
        .find(|n| n.is_element() && n.tag_name().name() == "extLst")?
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "ext")
        .find_map(|ext| {
            ext.children()
                .find(|n| n.is_element() && n.tag_name().name() == "svgBlip")
        })
        .and_then(|svg_blip| blip_embed_rid(&svg_blip))
}

/// Map a part name / path extension to its MIME type, covering the image, audio
/// and video parts OOXML blips reference. Unknown extensions fall back to
/// `application/octet-stream`. This is the single source of truth for all three
/// parsers — notably it is the only place `svg` ⇒ `image/svg+xml` is decided.
pub fn mime_from_ext(path: &str) -> &'static str {
    match path
        .rsplit('.')
        .next()
        .unwrap_or("")
        .to_ascii_lowercase()
        .as_str()
    {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "svg" => "image/svg+xml",
        "webp" => "image/webp",
        // Windows Metafile / Enhanced Metafile. Office embeds these for charts
        // and diagrams; the renderer rasterizes WMF via a minimal player and
        // skips EMF (a follow-up). The conventional MIME types per IANA / Windows.
        "wmf" => "image/wmf",
        "emf" => "image/emf",
        "mp3" => "audio/mpeg",
        "m4a" => "audio/mp4",
        "wav" => "audio/wav",
        "aac" => "audio/aac",
        "wma" => "audio/x-ms-wma",
        "flac" => "audio/flac",
        "ogg" | "oga" => "audio/ogg",
        "mp4" | "m4v" => "video/mp4",
        "mov" | "qt" => "video/quicktime",
        "avi" => "video/x-msvideo",
        "wmv" => "video/x-ms-wmv",
        "mpg" | "mpeg" => "video/mpeg",
        "3gp" => "video/3gpp",
        "mkv" => "video/x-matroska",
        "webm" => "video/webm",
        "ogv" => "video/ogg",
        _ => "application/octet-stream",
    }
}

/// ECMA-376 §20.1.8.55 `<a:srcRect>` source-image crop (DrawingML
/// `CT_RelativeRect`). Each edge inset is a fraction `0..1` of the *source*
/// bitmap, measured inward from that edge, so the visible source region is
/// `[l, t, 1−r, 1−b]`. The raw attributes are `ST_Percentage` in 1000ths of a
/// percent; the parser divides by 100000 to a fraction so renderers need no unit
/// knowledge. Absent edges default to `0`. Shared by the docx, pptx and xlsx
/// parsers so the three formats crop identically.
#[derive(Serialize, Deserialize, Debug, Clone, Copy, Default, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SrcRect {
    pub l: f64,
    pub t: f64,
    pub r: f64,
    pub b: f64,
}

/// Parse `<a:srcRect l t r b>` from a `<*:blipFill>` node (the parent of the
/// `<a:blip>`), matched by namespace-local name so a `p:`, `a:` or `xdr:`
/// blipFill all work. Each edge is divided by 100000 to a fraction; an absent
/// edge defaults to `0`. Returns `None` when there is no `srcRect` or all four
/// edges are zero (no crop), so an uncropped picture never forces a renderer
/// onto the sub-rectangle draw path.
pub fn parse_src_rect(blip_fill: Node<'_, '_>) -> Option<SrcRect> {
    let sr = blip_fill
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "srcRect")?;
    let read = |name: &str| -> f64 {
        sr.attribute(name)
            .and_then(|v| v.parse::<f64>().ok())
            .map(|v| v / 100_000.0)
            .unwrap_or(0.0)
    };
    let rect = SrcRect {
        l: read("l"),
        t: read("t"),
        r: read("r"),
        b: read("b"),
    };
    if rect.l.abs() < 1e-9 && rect.t.abs() < 1e-9 && rect.r.abs() < 1e-9 && rect.b.abs() < 1e-9 {
        None
    } else {
        Some(rect)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use roxmltree::Document;

    const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const ASVG_NS: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";

    // ISO/IEC 29500 Strict: a blip declaring the relationships namespace under
    // the `http://purl.oclc.org/ooxml/officeDocument/relationships` URI still
    // resolves its `r:embed` rId (the `relationships::STRICT` branch of
    // `blip_embed_rid`).
    #[test]
    fn blip_embed_rid_reads_strict_relationships_ns() {
        let r_strict = crate::ns::relationships::STRICT;
        let xml = format!(r#"<a:blip xmlns:a="{A_NS}" xmlns:r="{r_strict}" r:embed="rIdStrict"/>"#);
        let doc = Document::parse(&xml).unwrap();
        assert_eq!(
            blip_embed_rid(&doc.root_element()).as_deref(),
            Some("rIdStrict")
        );
    }

    #[test]
    fn svg_blip_rid_reads_extension_regardless_of_prefix() {
        // The svgBlip uses a different prefix (asvg:) on purpose: matching is by
        // namespace-local name, so the prefix must not matter.
        let xml = format!(
            r#"<a:blip xmlns:a="{A_NS}" xmlns:r="{R_NS}" xmlns:asvg="{ASVG_NS}" r:embed="rIdPng">
                 <a:extLst>
                   <a:ext uri="{SVG_BLIP_EXT_URI}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext>
                 </a:extLst>
               </a:blip>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let blip = doc.root_element();
        assert_eq!(blip_embed_rid(&blip).as_deref(), Some("rIdPng"));
        assert_eq!(svg_blip_rid(blip).as_deref(), Some("rIdSvg"));
    }

    #[test]
    fn parse_src_rect_reads_fractions_and_defaults_absent_edges() {
        // l/r given, t/b absent → fractions = raw/100000, absent ⇒ 0.
        let xml = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:blip/><a:srcRect l="32560" r="3829"/></a:blipFill>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let sr = parse_src_rect(doc.root_element()).expect("srcRect present");
        assert!((sr.l - 0.3256).abs() < 1e-9);
        assert!((sr.r - 0.03829).abs() < 1e-9);
        assert_eq!(sr.t, 0.0);
        assert_eq!(sr.b, 0.0);
    }

    #[test]
    fn parse_src_rect_none_when_absent_or_all_zero() {
        let none = format!(r#"<a:blipFill xmlns:a="{A_NS}"><a:blip/></a:blipFill>"#);
        assert!(parse_src_rect(Document::parse(&none).unwrap().root_element()).is_none());
        let zero = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:srcRect l="0" t="0" r="0" b="0"/></a:blipFill>"#
        );
        assert!(parse_src_rect(Document::parse(&zero).unwrap().root_element()).is_none());
    }

    #[test]
    fn svg_blip_only_has_no_raster_embed() {
        // sample-12 shape: the blip has no r:embed at all, only the svgBlip ext.
        let xml = format!(
            r#"<a:blip xmlns:a="{A_NS}" xmlns:r="{R_NS}" xmlns:asvg="{ASVG_NS}">
                 <a:extLst>
                   <a:ext uri="{SVG_BLIP_EXT_URI}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext>
                 </a:extLst>
               </a:blip>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let blip = doc.root_element();
        assert_eq!(blip_embed_rid(&blip), None);
        assert_eq!(svg_blip_rid(blip).as_deref(), Some("rIdSvg"));
    }

    #[test]
    fn plain_blip_has_no_svg() {
        let xml = format!(r#"<a:blip xmlns:a="{A_NS}" xmlns:r="{R_NS}" r:embed="rId1"/>"#);
        let doc = Document::parse(&xml).unwrap();
        let blip = doc.root_element();
        assert_eq!(blip_embed_rid(&blip).as_deref(), Some("rId1"));
        assert_eq!(svg_blip_rid(blip), None);
    }

    #[test]
    fn mime_table_covers_svg_and_common_rasters() {
        assert_eq!(mime_from_ext("../media/image2.svg"), "image/svg+xml");
        assert_eq!(mime_from_ext("a/b/image1.PNG"), "image/png");
        assert_eq!(mime_from_ext("x.jpeg"), "image/jpeg");
        assert_eq!(mime_from_ext("x.webp"), "image/webp");
        assert_eq!(mime_from_ext("noext"), "application/octet-stream");
    }

    #[test]
    fn mime_table_covers_metafiles() {
        // WMF/EMF blips (charts, diagrams) get the conventional metafile MIMEs,
        // not the application/octet-stream fallback, so the renderer's content
        // sniff has a sensible MIME to pair with each part.
        assert_eq!(mime_from_ext("word/media/image1.wmf"), "image/wmf");
        assert_eq!(mime_from_ext("ppt/media/image3.emf"), "image/emf");
        // Case-insensitive, like the raster entries.
        assert_eq!(mime_from_ext("a/b/CHART.WMF"), "image/wmf");
        assert_eq!(mime_from_ext("x.EMF"), "image/emf");
    }
}
