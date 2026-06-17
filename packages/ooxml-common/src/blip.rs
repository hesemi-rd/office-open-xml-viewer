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

use roxmltree::Node;

/// Relationships namespace (`r:`) — where a blip's `embed` / `link` rIds live.
pub const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Microsoft 2016 SVG extension URI (MS-ODRAWXML). An `<a:blip>` wrapping a
/// vector image carries `<a:extLst><a:ext uri="{96DAC541-…}"><asvg:svgBlip
/// r:embed="…"/></a:ext></a:extLst>`.
pub const SVG_BLIP_EXT_URI: &str = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";

/// Resolve a blip-like node's `r:embed` relationship id (the raster image, or
/// an `svgBlip`'s vector target). Reads the `embed` attribute in the
/// relationships namespace, tolerating the literal `r:embed` form some producers
/// emit without binding the namespace.
pub fn blip_embed_rid(node: &Node<'_, '_>) -> Option<String> {
    node.attribute((R_NS, "embed"))
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

#[cfg(test)]
mod tests {
    use super::*;
    use roxmltree::Document;

    const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const ASVG_NS: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";

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
}
