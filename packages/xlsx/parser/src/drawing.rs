use crate::types::*;
use crate::{parse_rels_map, resolve_zip_path};
use ooxml_common::zip::read_zip_string;
// Shared DrawingML blip helpers (ECMA-376 §20.1.8.13 + Microsoft 2016 SVG
// extension, MS-ODRAWXML). `mime_from_ext` is the single source of truth for
// `.svg ⇒ image/svg+xml`; `svg_blip_rid` resolves the vector original nested in
// a blip's `<a:extLst>`. Replaces xlsx's former local `mime_from_ext` (a strict
// subset that lacked the `svg` arm and so dropped SVG parts).
use ooxml_common::blip::{blip_embed_rid, mime_from_ext, parse_src_rect, svg_blip_rid};
use std::collections::HashMap;
use std::io::Cursor;

/// Parse `<xdr:twoCellAnchor>` elements from a drawing XML and resolve
/// embedded pictures into data URLs. `drawing_dir` is the folder that
/// contains `drawing_path` so relative `Target`s resolve correctly.
pub(crate) fn parse_drawing_anchors(
    drawing_xml: &str,
    drawing_rels: &HashMap<String, String>,
    drawing_dir: &str,
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
) -> Vec<ImageAnchor> {
    let Ok(doc) = roxmltree::Document::parse(drawing_xml) else {
        return Vec::new();
    };
    let xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    let a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main";
    let mut anchors: Vec<ImageAnchor> = Vec::new();

    for anchor in doc.descendants() {
        if anchor.tag_name().name() != "twoCellAnchor"
            || anchor.tag_name().namespace() != Some(xdr_ns)
        {
            continue;
        }
        let (mut from_col, mut from_col_off, mut from_row, mut from_row_off) =
            (0u32, 0i64, 0u32, 0i64);
        let (mut to_col, mut to_col_off, mut to_row, mut to_row_off) = (0u32, 0i64, 0u32, 0i64);
        let mut pic_rid: Option<String> = None;
        // Microsoft 2016 SVG extension (MS-ODRAWXML): the blip's vector original,
        // nested in `<a:blip><a:extLst>`. The raster `r:embed` (pic_rid) is only a
        // fallback for SVG-incapable clients.
        let mut svg_rid: Option<String> = None;
        let mut native_ext_cx: i64 = 0;
        let mut native_ext_cy: i64 = 0;
        // ECMA-376 §20.1.8.55 `<a:srcRect>` source-image crop (None ⇒ uncropped).
        let mut src_rect: Option<SrcRect> = None;
        // ECMA-376 §20.5.2.33 `twoCellAnchor@editAs`. Possible values:
        // "twoCell" (default), "oneCell", "absolute". With "oneCell" Excel
        // preserves the picture's saved size from <xdr:spPr><a:xfrm><a:ext>
        // regardless of cell resizing.
        let edit_as = anchor.attribute("editAs").map(|s| s.to_string());

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
                "pic" => {
                    // <xdr:pic><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic>
                    let blip_fill = child.children().find(|n| {
                        n.tag_name().name() == "blipFill"
                            && n.tag_name().namespace() == Some(xdr_ns)
                    });
                    if let Some(bf) = blip_fill {
                        let blip = bf.children().find(|n| {
                            n.tag_name().name() == "blip" && n.tag_name().namespace() == Some(a_ns)
                        });
                        if let Some(b) = blip {
                            // Raster fallback (`<a:blip r:embed>`); tolerate the
                            // literal `r:embed` form via the shared helper.
                            pic_rid = blip_embed_rid(&b);
                            // Vector original (`<asvg:svgBlip r:embed>` inside the
                            // blip's `<a:extLst>`), matched by namespace-local name.
                            svg_rid = svg_blip_rid(b);
                        }
                        // `<a:srcRect>` is a sibling of `<a:blip>` inside blipFill.
                        src_rect = parse_src_rect(bf);
                    }
                    // <xdr:pic><xdr:spPr><a:xfrm><a:ext cx cy>: the picture's
                    // own saved EMU extent. Authoritative when editAs="oneCell".
                    if let Some(sp_pr) = child.children().find(|n| {
                        n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(xdr_ns)
                    }) {
                        if let Some(xfrm_n) = sp_pr.children().find(|n| {
                            n.tag_name().name() == "xfrm" && n.tag_name().namespace() == Some(a_ns)
                        }) {
                            if let Some(xfrm) = parse_xfrm(&xfrm_n) {
                                native_ext_cx = xfrm.ext_x as i64;
                                native_ext_cy = xfrm.ext_y as i64;
                            }
                        }
                    }
                }
                _ => {}
            }
        }

        // Resolve a blip rId → the media part's zip path (verifying the entry
        // exists, so a dangling rId is dropped exactly as before). The renderer
        // fetches the bytes lazily via `extract_image`; no base64 is inlined.
        let mut resolve = |rid: &str| -> Option<String> {
            let target = drawing_rels.get(rid)?;
            let media_path = resolve_zip_path(drawing_dir, target);
            // Confirm the entry resolves before emitting its path (preserves the
            // previous "drop when bytes are missing" semantics). `index_for_name`
            // reads only the central directory — no inflate, unlike the former
            // `read_zip_bytes` which decompressed the entry only to discard it.
            archive.index_for_name(&media_path)?;
            Some(media_path)
        };

        // Vector original first (so an svg-only picture is never dropped); raster
        // fallback second.
        let svg_image_path = svg_rid.as_deref().and_then(&mut resolve);
        let raster_path = pic_rid.as_deref().and_then(&mut resolve);

        // A picture needs at least one drawable source. Prefer the raster as
        // `image_path` (Excel's compatibility fallback); when no raster is
        // embedded, fall back to the SVG itself so the element is always
        // drawable. Drop only when neither resolves. ECMA-376 §20.1.8.13 +
        // MS-ODRAWXML svgBlip.
        let image_path = match (raster_path, svg_image_path.as_ref()) {
            (Some(raster), _) => raster,
            (None, Some(svg)) => svg.clone(),
            (None, None) => continue,
        };
        let mime_type = mime_from_ext(&image_path).to_string();

        anchors.push(ImageAnchor {
            from_col,
            from_col_off,
            from_row,
            from_row_off,
            to_col,
            to_col_off,
            to_row,
            to_row_off,
            edit_as,
            native_ext_cx,
            native_ext_cy,
            image_path,
            mime_type,
            svg_image_path,
            src_rect,
        });
    }
    anchors
}

// ─── Shape group parsing ────────────────────────────────────────────────────
//
// ECMA-376 §20.5.2.17 `<xdr:grpSp>` / §20.1.9 DrawingML shapes. Each
// top-level grpSp inside a twoCellAnchor has its own coordinate system:
//   - grpSpPr/xfrm/off,ext     : group's position/size in parent coords
//   - grpSpPr/xfrm/chOff,chExt : origin/extent of the group's child coords
//
// A child sp at child coord (cx, cy) maps to parent coord:
//   parent.x = off.x + (cx - chOff.x) / chExt.cx * ext.cx
//
// For rendering, we chain these transforms down to the top-level grpSp and
// then normalize each leaf shape's rect into [0,1] of the top-level ext.

#[derive(Clone, Copy)]
pub(crate) struct Xfrm {
    off_x: f64,
    off_y: f64,
    ext_x: f64,
    ext_y: f64,
    ch_off_x: f64,
    ch_off_y: f64,
    ch_ext_x: f64,
    ch_ext_y: f64,
    has_ch: bool,
}

pub(crate) fn parse_xfrm(xfrm_node: &roxmltree::Node) -> Option<Xfrm> {
    let mut off = (0.0_f64, 0.0_f64);
    let mut ext = (0.0_f64, 0.0_f64);
    let mut ch_off = (0.0_f64, 0.0_f64);
    let mut ch_ext = (0.0_f64, 0.0_f64);
    let mut has_ext = false;
    let mut has_ch = false;
    for c in xfrm_node.children() {
        match c.tag_name().name() {
            "off" => {
                off.0 = c.attribute("x").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                off.1 = c.attribute("y").and_then(|s| s.parse().ok()).unwrap_or(0.0);
            }
            "ext" => {
                ext.0 = c
                    .attribute("cx")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0.0);
                ext.1 = c
                    .attribute("cy")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0.0);
                has_ext = true;
            }
            "chOff" => {
                ch_off.0 = c.attribute("x").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                ch_off.1 = c.attribute("y").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                has_ch = true;
            }
            "chExt" => {
                ch_ext.0 = c
                    .attribute("cx")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0.0);
                ch_ext.1 = c
                    .attribute("cy")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0.0);
                has_ch = true;
            }
            _ => {}
        }
    }
    if !has_ext {
        return None;
    }
    Some(Xfrm {
        off_x: off.0,
        off_y: off.1,
        ext_x: ext.0,
        ext_y: ext.1,
        ch_off_x: ch_off.0,
        ch_off_y: ch_off.1,
        ch_ext_x: if ch_ext.0 == 0.0 { ext.0 } else { ch_ext.0 },
        ch_ext_y: if ch_ext.1 == 0.0 { ext.1 } else { ch_ext.1 },
        has_ch,
    })
}

/// Apply the DrawingML color-transform children (`lumMod`, `lumOff`, `shade`,
/// `tint`, `satMod`, …) declared inside a `<a:srgbClr>` / `<a:schemeClr>` to a
/// resolved base hex, via the shared transform. Without this, an accent with
/// e.g. `lumMod 20% + lumOff 80%` (a light tint) renders at full strength — so
/// "light fill + dark border" pairs collapse to one solid mid-tone.
fn apply_clr_mods(base_with_hash: &str, clr_node: &roxmltree::Node) -> String {
    let base = base_with_hash.trim_start_matches('#');
    let out = ooxml_common::color::apply_color_transforms(
        base,
        *clr_node,
        ooxml_common::color::TintMode::PowerPointLinear,
    );
    format!("#{}", out.to_uppercase())
}

pub(crate) fn parse_solid_fill(
    fill_node: &roxmltree::Node,
    theme_colors: &[String],
) -> Option<String> {
    for c in fill_node.children() {
        match c.tag_name().name() {
            "srgbClr" => {
                let v = c.attribute("val")?;
                return Some(apply_clr_mods(&format!("#{}", v.to_uppercase()), &c));
            }
            "sysClr" => {
                // System colour (e.g. windowText / window). `lastClr` is the
                // concrete value the authoring app last resolved it to.
                let last = c.attribute("lastClr")?;
                return Some(apply_clr_mods(&format!("#{}", last.to_uppercase()), &c));
            }
            "schemeClr" => {
                let v = c.attribute("val")?;
                // `theme_colors` is collected in OOXML clrScheme document
                // order: dk1, lt1, dk2, lt2, accent1..accent6, hlink,
                // folHlink. See `parse_theme_colors`. The earlier mapping
                // here had dk1/lt1 and dk2/lt2 swapped which darkened
                // shapes that painted "lt1" (the sheet paper colour).
                let idx = match v {
                    "dk1" | "tx1" => Some(0),
                    "lt1" | "bg1" => Some(1),
                    "dk2" | "tx2" => Some(2),
                    "lt2" | "bg2" => Some(3),
                    "accent1" => Some(4),
                    "accent2" => Some(5),
                    "accent3" => Some(6),
                    "accent4" => Some(7),
                    "accent5" => Some(8),
                    "accent6" => Some(9),
                    "hlink" => Some(10),
                    "folHlink" => Some(11),
                    _ => None,
                };
                return idx
                    .and_then(|i| theme_colors.get(i).cloned())
                    .map(|base| apply_clr_mods(&base, &c));
            }
            _ => {}
        }
    }
    None
}

/// Parse a single custGeom path element. Each path has its own coordinate
/// system (`a:path/@w`, `@h`) that the renderer scales to the shape's rect.
pub(crate) fn parse_custom_path(path_node: &roxmltree::Node) -> PathInfo {
    let w: f64 = path_node
        .attribute("w")
        .and_then(|s| s.parse().ok())
        .unwrap_or(0.0);
    let h: f64 = path_node
        .attribute("h")
        .and_then(|s| s.parse().ok())
        .unwrap_or(0.0);
    let mut commands: Vec<PathCmd> = Vec::new();
    for cmd in path_node.children().filter(|n| n.is_element()) {
        let name = cmd.tag_name().name();
        // Collect `<a:pt x=.. y=..>` points in order.
        let pts: Vec<(f64, f64)> = cmd
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "pt")
            .map(|n| {
                (
                    n.attribute("x").and_then(|s| s.parse().ok()).unwrap_or(0.0),
                    n.attribute("y").and_then(|s| s.parse().ok()).unwrap_or(0.0),
                )
            })
            .collect();
        match name {
            "moveTo" => {
                if let Some(p) = pts.first() {
                    commands.push(PathCmd::MoveTo { x: p.0, y: p.1 });
                }
            }
            "lnTo" => {
                if let Some(p) = pts.first() {
                    commands.push(PathCmd::LineTo { x: p.0, y: p.1 });
                }
            }
            "cubicBezTo" => {
                if pts.len() >= 3 {
                    commands.push(PathCmd::CubicBezTo {
                        x1: pts[0].0,
                        y1: pts[0].1,
                        x2: pts[1].0,
                        y2: pts[1].1,
                        x3: pts[2].0,
                        y3: pts[2].1,
                    });
                }
            }
            "quadBezTo" => {
                if pts.len() >= 2 {
                    commands.push(PathCmd::QuadBezTo {
                        x1: pts[0].0,
                        y1: pts[0].1,
                        x2: pts[1].0,
                        y2: pts[1].1,
                    });
                }
            }
            "close" => commands.push(PathCmd::Close),
            "arcTo" => {
                // ECMA-376 §20.1.9.3: `wR`/`hR` in path-coord units;
                // `stAng`/`swAng` in 60000ths of a degree.
                let wr: f64 = cmd
                    .attribute("wR")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0.0);
                let hr: f64 = cmd
                    .attribute("hR")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0.0);
                let st_ang: f64 = cmd
                    .attribute("stAng")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0.0);
                let sw_ang: f64 = cmd
                    .attribute("swAng")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(0.0);
                commands.push(PathCmd::ArcTo {
                    wr,
                    hr,
                    st_ang,
                    sw_ang,
                });
            }
            _ => {}
        }
    }
    PathInfo { w, h, commands }
}

/// Parse `<xdr:txBody>` into a `ShapeText`. Returns `None` if the body
/// contains no visible runs. Run formatting follows ECMA-376 §21.1.2.3.1
/// (`<a:rPr>`): `sz` is hundredths of a point, `b="1"` = bold, `i="1"`
/// = italic, `<a:solidFill>` overrides shape-level font color,
/// `<a:latin@typeface>` selects the Latin font face, and `<a:ea@typeface>` /
/// `<a:cs@typeface>` the East-Asian / complex-script faces. The renderer
/// floors the line box by whichever of these declares a tabled (Meiryo /
/// Sakkal Majalla) design line — the common Japanese encoding sets Meiryo on
/// `<a:ea>` while leaving `<a:latin>` default. Per-glyph font switching is
/// still not modeled (a larger change); only the line-box floor uses ea/cs.
/// Point size for an equation: the first `a:rPr@sz` on an actual math RUN
/// (`m:r`), in pt. Scoped to `m:r` so the size on `<m:ctrlPr>` (control
/// properties — structural delimiters, not visible glyphs) is ignored.
fn math_run_size_shape(om: roxmltree::Node) -> Option<f64> {
    om.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "r")
        .find_map(|r| {
            r.descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "rPr")
                .find_map(|rpr| rpr.attribute("sz").and_then(|s| s.parse::<f64>().ok()))
        })
        .map(|v| v / 100.0)
}

/// Equation colour: the first `a:rPr/solidFill` on an actual math RUN (`m:r`).
/// Scoped to `m:r` so a colour on `<m:ctrlPr>` (structural control properties,
/// which Excel does NOT use as the visible glyph colour) is ignored — equations
/// with no explicit run colour then fall back to the shape font colour (black).
fn math_run_color_shape(om: roxmltree::Node, theme_colors: &[String]) -> Option<String> {
    om.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "r")
        .find_map(|r| {
            r.descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "rPr")
                .find_map(|rpr| {
                    rpr.children()
                        .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
                        .and_then(|sf| parse_solid_fill(&sf, theme_colors))
                })
        })
}

/// Emit `ShapeTextRun::Math` runs from an OMML wrapper inside a `<a:p>`. Handles
/// the four shapes Excel/PowerPoint produce — bare `m:oMath`, `m:oMathPara`
/// (block), the `a14:m` wrapper (local name `m`), and `mc:AlternateContent`
/// (take the `a14` `Choice`, ignore the rasterized `Fallback`). Direct port of
/// the pptx parser's `push_math_runs` (ECMA-376 §22.1); reuses the shared
/// `ooxml_common::math` OMML→AST parser unchanged.
fn push_math_runs_shape(
    node: roxmltree::Node,
    theme_colors: &[String],
    runs: &mut Vec<ShapeTextRun>,
) {
    match node.tag_name().name() {
        "oMath" => {
            let nodes = ooxml_common::math::parse_omath_nodes(node);
            if !nodes.is_empty() {
                runs.push(ShapeTextRun::Math {
                    nodes,
                    display: false,
                    font_size: math_run_size_shape(node),
                    color: math_run_color_shape(node, theme_colors),
                });
            }
        }
        "oMathPara" => {
            for om in node
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "oMath")
            {
                let nodes = ooxml_common::math::parse_omath_nodes(om);
                if !nodes.is_empty() {
                    runs.push(ShapeTextRun::Math {
                        nodes,
                        display: true,
                        font_size: math_run_size_shape(om),
                        color: math_run_color_shape(om, theme_colors),
                    });
                }
            }
        }
        // mc:AlternateContent → the a14 `Choice` holds the live equation; the
        // `Fallback` holds a rasterized picture we must NOT also draw.
        "AlternateContent" => {
            if let Some(choice) = node
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "Choice")
            {
                for c in choice.children().filter(|n| n.is_element()) {
                    push_math_runs_shape(c, theme_colors, runs);
                }
            }
        }
        // a14:m wrapper (local name "m") holds an m:oMathPara / m:oMath.
        "m" => {
            for c in node.children().filter(|n| n.is_element()) {
                push_math_runs_shape(c, theme_colors, runs);
            }
        }
        _ => {}
    }
}

pub(crate) fn parse_tx_body(
    tx_body: &roxmltree::Node,
    theme_colors: &[String],
) -> Option<ShapeText> {
    let mut anchor = String::from("t");
    let mut wrap = String::from("square");
    // `<a:bodyPr>` autofit child (ECMA-376 §21.1.2.1.1-.3). Default "none"
    // (xlsx has no theme txDef fallback). Mirrors the pptx parser; for
    // normAutofit also capture the stored fontScale / lnSpcReduction
    // (ST_TextFontScalePercentOrPercentString / ST_TextSpacingPercentOrPercentString,
    // 1000ths of a percent → fraction).
    let mut auto_fit = String::from("none");
    let mut font_scale: Option<f64> = None;
    let mut ln_spc_reduction: Option<f64> = None;
    let mut paragraphs: Vec<ShapeParagraph> = Vec::new();
    for c in tx_body.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "bodyPr" => {
                if let Some(a) = c.attribute("anchor") {
                    anchor = a.to_string();
                }
                if let Some(w) = c.attribute("wrap") {
                    wrap = w.to_string();
                }
                // Autofit child (spAutoFit / normAutofit / noAutofit). Shared
                // with pptx; xlsx keeps the "none" default when there is no
                // autofit child (parse_autofit returns None).
                if let Some((af, fs, lsr)) = ooxml_common::text::parse_autofit(c) {
                    auto_fit = af;
                    font_scale = fs;
                    ln_spc_reduction = lsr;
                }
            }
            "p" => {
                let mut align = String::from("l");
                let mut rtl = false;
                let mut mar_l: Option<i64> = None;
                let mut mar_r: Option<i64> = None;
                let mut indent: Option<i64> = None;
                let mut space_line: Option<SpaceLine> = None;
                let mut runs: Vec<ShapeTextRun> = Vec::new();
                for pc in c.children().filter(|n| n.is_element()) {
                    match pc.tag_name().name() {
                        "pPr" => {
                            if let Some(a) = pc.attribute("algn") {
                                align = a.to_string();
                            }
                            // ECMA-376 §21.1.2.2.7 `<a:pPr@rtl>` — right-to-left
                            // paragraph. Default false (left-to-right).
                            rtl = pc
                                .attribute("rtl")
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false);
                            // ECMA-376 §21.1.2.2.7 (`CT_TextParagraphProperties`)
                            // direct indent attributes (EMU). `indent` may be
                            // negative (hanging). Direct-attribute-only — xlsx
                            // text boxes have no lstStyle/level cascade.
                            mar_l = pc.attribute("marL").and_then(|v| v.parse().ok());
                            mar_r = pc.attribute("marR").and_then(|v| v.parse().ok());
                            indent = pc.attribute("indent").and_then(|v| v.parse().ok());
                            // ECMA-376 §21.1.2.2.5 `<a:lnSpc>`: spcPct is a
                            // percentage of the natural single line; spcPts is an
                            // absolute per-line height (raw @val is hundredths of
                            // a point → divide by 100). Shared with pptx.
                            if let Some(ln_spc) = pc
                                .children()
                                .find(|n| n.is_element() && n.tag_name().name() == "lnSpc")
                            {
                                space_line = ooxml_common::text::parse_lnspc(ln_spc);
                            }
                        }
                        "r" => {
                            // Run text + run-level formatting.
                            let mut text = String::new();
                            let mut bold = false;
                            let mut italic = false;
                            let mut size: f64 = 0.0;
                            let mut color: Option<String> = None;
                            let mut font_face: Option<String> = None;
                            // ECMA-376 §21.1.2.3.1 `<a:ea>` / `<a:cs>` typefaces.
                            // Parsed alongside `<a:latin>` so the renderer can
                            // floor the line box by a tabled East-Asian face
                            // (e.g. Meiryo set only on `<a:ea>`).
                            let mut font_face_ea: Option<String> = None;
                            let mut font_face_cs: Option<String> = None;
                            for rc in pc.children().filter(|n| n.is_element()) {
                                match rc.tag_name().name() {
                                    "rPr" => {
                                        bold = rc.attribute("b").map(|v| v == "1").unwrap_or(false);
                                        italic =
                                            rc.attribute("i").map(|v| v == "1").unwrap_or(false);
                                        size = rc
                                            .attribute("sz")
                                            .and_then(|s| s.parse::<f64>().ok())
                                            .map(|v| v / 100.0)
                                            .unwrap_or(0.0);
                                        for rpc in rc.children().filter(|n| n.is_element()) {
                                            match rpc.tag_name().name() {
                                                "solidFill" => {
                                                    color = parse_solid_fill(&rpc, theme_colors);
                                                }
                                                "latin" => {
                                                    font_face =
                                                        rpc.attribute("typeface").map(String::from);
                                                }
                                                "ea" => {
                                                    font_face_ea =
                                                        rpc.attribute("typeface").map(String::from);
                                                }
                                                "cs" => {
                                                    font_face_cs =
                                                        rpc.attribute("typeface").map(String::from);
                                                }
                                                _ => {}
                                            }
                                        }
                                    }
                                    "t" => {
                                        if let Some(t) = rc.text() {
                                            text.push_str(t);
                                        }
                                    }
                                    _ => {}
                                }
                            }
                            if !text.is_empty() {
                                runs.push(ShapeTextRun::Text {
                                    text,
                                    bold,
                                    italic,
                                    size,
                                    color,
                                    font_face,
                                    font_face_ea,
                                    font_face_cs,
                                });
                            }
                        }
                        "br" => {
                            // Soft line break (`<a:br>`).
                            runs.push(ShapeTextRun::Break);
                        }
                        // OMML equations (ECMA-376 §22.1) embedded in shape text:
                        // bare `m:oMath` / `m:oMathPara`, the `a14:m` wrapper
                        // (local name "m"), or `mc:AlternateContent`.
                        "oMath" | "oMathPara" | "AlternateContent" | "m" => {
                            push_math_runs_shape(pc, theme_colors, &mut runs);
                        }
                        _ => {}
                    }
                }
                if !runs.is_empty() {
                    paragraphs.push(ShapeParagraph {
                        align,
                        rtl,
                        mar_l,
                        mar_r,
                        indent,
                        space_line,
                        runs,
                    });
                }
            }
            _ => {}
        }
    }
    if paragraphs.is_empty() {
        None
    } else {
        Some(ShapeText {
            anchor,
            wrap,
            auto_fit,
            font_scale,
            ln_spc_reduction,
            paragraphs,
        })
    }
}

/// Parse the adjust handles from a `<a:prstGeom>`'s `<a:avLst>` into an
/// `adj1..adj8`-ordered vector. ECMA-376 §19.5.31.3 / §20.1.9.5: each
/// `<a:gd name="adjN" fmla="val X"/>` supplies one handle. Matching is by name
/// (`adj`/`adj1`, `adj2`, …) with a positional fallback, mirroring the pptx
/// parser so the shared `renderPresetShape` engine receives an identical shape.
/// Trailing `None`s are trimmed so a plain rect/ellipse yields an empty vec.
fn parse_preset_adj(prst_geom: &roxmltree::Node) -> Vec<Option<f64>> {
    let av = prst_geom
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "avLst");
    let Some(av) = av else {
        return Vec::new();
    };
    let gd_nodes: Vec<roxmltree::Node> = av
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "gd")
        .collect();
    if gd_nodes.is_empty() {
        return Vec::new();
    }

    let parse_val = |gd: &roxmltree::Node| -> Option<f64> {
        gd.attribute("fmla")
            .and_then(|f| f.strip_prefix("val "))
            .and_then(|s| s.trim().parse::<f64>().ok())
    };

    // When the handles carry `name` attributes, match strictly by name so a
    // missing adjN leaves a `None` gap (its index still aligns with the engine's
    // adj1..adj8 expectations). Positional fallback is used only for legacy
    // unnamed `<a:gd>` lists.
    let named = gd_nodes.iter().any(|n| n.attribute("name").is_some());
    let mut out: Vec<Option<f64>> = (1..=8)
        .map(|i| {
            let by_name = gd_nodes.iter().find(|n| {
                let name = n.attribute("name");
                if i == 1 {
                    matches!(name, Some("adj") | Some("adj1"))
                } else {
                    name == Some(&format!("adj{i}")[..])
                }
            });
            if named {
                by_name.and_then(parse_val)
            } else {
                gd_nodes.get(i - 1).and_then(parse_val)
            }
        })
        .collect();

    while matches!(out.last(), Some(None)) {
        out.pop();
    }
    out
}

pub(crate) fn parse_sp_geom(sp_pr: &roxmltree::Node) -> Option<ShapeGeom> {
    for c in sp_pr.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "prstGeom" => {
                return Some(ShapeGeom::Preset {
                    name: c.attribute("prst").unwrap_or("rect").to_string(),
                    adj: parse_preset_adj(&c),
                });
            }
            "custGeom" => {
                let mut paths: Vec<PathInfo> = Vec::new();
                for pl in c
                    .children()
                    .filter(|n| n.is_element() && n.tag_name().name() == "pathLst")
                {
                    for p in pl
                        .children()
                        .filter(|n| n.is_element() && n.tag_name().name() == "path")
                    {
                        paths.push(parse_custom_path(&p));
                    }
                }
                return Some(ShapeGeom::Custom { paths });
            }
            _ => {}
        }
    }
    None
}

/// Recursively walk an `xdr:grpSp` / `xdr:sp` tree, chaining coordinate
/// transforms, and push leaf shapes (normalized to [0,1] of `root_ext`) into
/// `out`.
// The arguments are the running coordinate-transform state threaded through the
// recursion (root offset/extent, accumulated scale/translation, theme + rels);
// bundling them into a struct would only move the same fields elsewhere without
// improving clarity.
#[allow(clippy::too_many_arguments)]
pub(crate) fn collect_shapes(
    node: &roxmltree::Node,
    root_off_x: f64,
    root_off_y: f64,
    root_ext_x: f64,
    root_ext_y: f64,
    // transform from current local coords into root (top-level grpSp) coords
    scale_x: f64,
    scale_y: f64,
    trans_x: f64,
    trans_y: f64,
    theme_colors: &[String],
    theme_ln_widths: &[i64],
    rid_urls: &HashMap<String, String>,
    out: &mut Vec<ShapeInfo>,
) {
    for child in node.children().filter(|n| n.is_element()) {
        let tag = child.tag_name().name();
        if tag == "grpSp" {
            // Nested grpSp: compose the transform by the group's own xfrm.
            let grp_sp_pr = child
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "grpSpPr");
            let xfrm = grp_sp_pr
                .and_then(|n| {
                    n.children()
                        .find(|c| c.is_element() && c.tag_name().name() == "xfrm")
                })
                .as_ref()
                .and_then(parse_xfrm);
            let (sx, sy, tx, ty) = if let Some(x) = xfrm {
                if x.has_ch && x.ch_ext_x != 0.0 && x.ch_ext_y != 0.0 {
                    let csx = x.ext_x / x.ch_ext_x;
                    let csy = x.ext_y / x.ch_ext_y;
                    // Child point (cx, cy) → (x.off_x + (cx - x.ch_off_x)*csx) in parent coords,
                    // then apply outer (scale/trans) to reach root coords.
                    (
                        scale_x * csx,
                        scale_y * csy,
                        trans_x + scale_x * (x.off_x - x.ch_off_x * csx),
                        trans_y + scale_y * (x.off_y - x.ch_off_y * csy),
                    )
                } else {
                    // No child coord system: treat as identity mapping inside the group.
                    (
                        scale_x,
                        scale_y,
                        trans_x + scale_x * x.off_x,
                        trans_y + scale_y * x.off_y,
                    )
                }
            } else {
                (scale_x, scale_y, trans_x, trans_y)
            };
            collect_shapes(
                &child,
                root_off_x,
                root_off_y,
                root_ext_x,
                root_ext_y,
                sx,
                sy,
                tx,
                ty,
                theme_colors,
                theme_ln_widths,
                rid_urls,
                out,
            );
        } else if tag == "sp" {
            let sp_pr = child
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "spPr");
            let Some(sp_pr) = sp_pr else {
                continue;
            };
            let xfrm_node = sp_pr
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "xfrm");
            let Some(xfrm_n) = xfrm_node else {
                continue;
            };
            let Some(xfrm) = parse_xfrm(&xfrm_n) else {
                continue;
            };
            let rot_raw: f64 = xfrm_n
                .attribute("rot")
                .and_then(|s| s.parse().ok())
                .unwrap_or(0.0);

            // Shape rect in root coords
            let root_x = trans_x + scale_x * xfrm.off_x;
            let root_y = trans_y + scale_y * xfrm.off_y;
            let root_w = scale_x * xfrm.ext_x;
            let root_h = scale_y * xfrm.ext_y;

            // Normalize to [0,1] of root ext
            if root_ext_x == 0.0 || root_ext_y == 0.0 {
                continue;
            }
            let nx = (root_x - root_off_x) / root_ext_x;
            let ny = (root_y - root_off_y) / root_ext_y;
            let nw = root_w / root_ext_x;
            let nh = root_h / root_ext_y;

            let geom = parse_sp_geom(&sp_pr);
            let Some(geom) = geom else {
                continue;
            };

            // Fill
            let mut fill_color: Option<String> = None;
            let mut has_no_fill = false;
            for c in sp_pr.children().filter(|n| n.is_element()) {
                match c.tag_name().name() {
                    "solidFill" => {
                        fill_color = parse_solid_fill(&c, theme_colors);
                    }
                    "noFill" => {
                        has_no_fill = true;
                    }
                    _ => {}
                }
            }
            if has_no_fill {
                fill_color = None;
            }

            // Stroke (line)
            let mut stroke_color: Option<String> = None;
            let mut stroke_width: i64 = 0;
            // An explicit `<a:ln><a:noFill/>` is a hard "no outline" and must
            // suppress the theme lnRef fallback below.
            let mut ln_no_stroke = false;
            if let Some(ln) = sp_pr
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "ln")
            {
                let w_attr = ln.attribute("w");
                stroke_width = w_attr.and_then(|s| s.parse().ok()).unwrap_or(0);
                for c in ln.children().filter(|n| n.is_element()) {
                    if c.tag_name().name() == "solidFill" {
                        stroke_color = parse_solid_fill(&c, theme_colors);
                        ln_no_stroke = false;
                    } else if c.tag_name().name() == "noFill" {
                        stroke_color = None;
                        stroke_width = 0;
                        ln_no_stroke = true;
                    }
                }
                // An `<a:ln>` with a fill but no explicit `w` still draws an
                // outline — Excel uses a thin default (~0.75pt = 9525 EMU).
                // Without this the border (e.g. a dark-green outline on a
                // light-green box) silently disappears.
                if w_attr.is_none() && stroke_color.is_some() && stroke_width == 0 {
                    stroke_width = 9525;
                }
            }

            // <xdr:style> drives fallbacks: <a:fillRef> supplies a fill when
            // <xdr:spPr> didn't, and <a:fontRef> supplies the run-default text
            // color (ECMA-376 §20.5.2.30 `<xdr:style>`). Real-world text boxes
            // saved by Excel often leave `<xdr:spPr>` without `<a:solidFill>`
            // and rely on the style's fillRef + fontRef pair (e.g. accent1
            // background + lt1 white text). We resolve scheme colors here
            // against the workbook theme and apply them as fallbacks.
            let style_node = child
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "style");
            let style_fill = style_node
                .as_ref()
                .and_then(|s| {
                    s.children()
                        .find(|n| n.is_element() && n.tag_name().name() == "fillRef")
                })
                .and_then(|n| parse_solid_fill(&n, theme_colors));
            let style_text_color = style_node
                .as_ref()
                .and_then(|s| {
                    s.children()
                        .find(|n| n.is_element() && n.tag_name().name() == "fontRef")
                })
                .and_then(|n| parse_solid_fill(&n, theme_colors));
            if fill_color.is_none() && !has_no_fill {
                fill_color = style_fill;
            }

            // <a:lnRef> — ECMA-376 §20.1.4.2.19: when `<xdr:spPr>` carries no
            // explicit `<a:ln>`, the shape's outline comes from the theme's
            // style matrix: `idx="N"` selects entry N (1-based) of
            // `fmtScheme > lnStyleLst` for the line WIDTH, and the lnRef's
            // child color (e.g. `<a:schemeClr val="accent1"/>`) substitutes
            // the entry's `phClr` placeholder for the line COLOR. This is how
            // Excel-inserted shapes get their default border, so without it
            // every default shape renders outline-less. An explicit
            // `<a:ln>` (including `<a:noFill/>`) always wins.
            if stroke_color.is_none() && !ln_no_stroke {
                if let Some(ln_ref) = style_node.as_ref().and_then(|s| {
                    s.children()
                        .find(|n| n.is_element() && n.tag_name().name() == "lnRef")
                }) {
                    if let Some(c) = parse_solid_fill(&ln_ref, theme_colors) {
                        stroke_color = Some(c);
                        let idx: usize = ln_ref
                            .attribute("idx")
                            .and_then(|v| v.parse().ok())
                            .unwrap_or(0);
                        // Missing/out-of-range entries fall back to the
                        // CT_LineProperties default width (§20.1.2.2.24,
                        // 9525 EMU = 0.75 pt).
                        stroke_width = if idx >= 1 {
                            theme_ln_widths.get(idx - 1).copied().unwrap_or(9525)
                        } else {
                            9525
                        };
                    }
                }
            }

            // Text body (txBox shapes carry visible text inside
            // `<xdr:txBody>`; non-textbox shapes may also have one).
            let mut text = child
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "txBody")
                .and_then(|tb| parse_tx_body(&tb, theme_colors));
            if let (Some(t), Some(default_color)) = (text.as_mut(), style_text_color) {
                for p in t.paragraphs.iter_mut() {
                    for r in p.runs.iter_mut() {
                        match r {
                            ShapeTextRun::Text { color, .. } | ShapeTextRun::Math { color, .. } => {
                                if color.is_none() {
                                    *color = Some(default_color.clone());
                                }
                            }
                            ShapeTextRun::Break => {}
                        }
                    }
                }
            }

            out.push(ShapeInfo {
                x: nx,
                y: ny,
                w: nw,
                h: nh,
                rot: rot_raw / 60000.0,
                fill_color,
                stroke_color,
                stroke_width,
                geom,
                text,
            });
        } else if tag == "pic" {
            // `<xdr:pic>` leaf inside a group (ECMA-376 §20.5.2.17). The image
            // binary is resolved via the drawing's .rels file; `rid_urls` maps
            // each r:id to its zip path inside the package.
            let sp_pr = child
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "spPr");
            let Some(sp_pr) = sp_pr else {
                continue;
            };
            let xfrm_node = sp_pr
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "xfrm");
            let Some(xfrm_n) = xfrm_node else {
                continue;
            };
            let Some(xfrm) = parse_xfrm(&xfrm_n) else {
                continue;
            };
            let rot_raw: f64 = xfrm_n
                .attribute("rot")
                .and_then(|s| s.parse().ok())
                .unwrap_or(0.0);

            // Resolve `<a:blip>`. The raster fallback rides in `r:embed`; the
            // vector original (Microsoft 2016 svgBlip extension, MS-ODRAWXML) is
            // nested in `<a:blip><a:extLst>`. `rid_urls` maps every media rId
            // (raster *and* svg) to its zip path inside the package.
            let blip = child
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "blip");
            let pic_rid = blip.as_ref().and_then(blip_embed_rid);
            let svg_rid = blip.and_then(svg_blip_rid);

            let svg_image_path = svg_rid.as_deref().and_then(|r| rid_urls.get(r)).cloned();
            let raster_path = pic_rid.as_deref().and_then(|r| rid_urls.get(r)).cloned();

            // `<a:srcRect>` crop lives in the leaf's `<xdr:blipFill>` (§20.1.8.55).
            let src_rect = child
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "blipFill")
                .and_then(parse_src_rect);

            // Prefer the raster as `image_path`; fall back to the SVG when no
            // raster is embedded so an svg-only leaf is never dropped. Drop only
            // when neither resolves. ECMA-376 §20.1.8.13 + MS-ODRAWXML svgBlip.
            let image_path = match (raster_path, svg_image_path.as_ref()) {
                (Some(raster), _) => raster,
                (None, Some(svg)) => svg.clone(),
                (None, None) => continue,
            };
            let mime_type = mime_from_ext(&image_path).to_string();

            let root_x = trans_x + scale_x * xfrm.off_x;
            let root_y = trans_y + scale_y * xfrm.off_y;
            let root_w = scale_x * xfrm.ext_x;
            let root_h = scale_y * xfrm.ext_y;
            if root_ext_x == 0.0 || root_ext_y == 0.0 {
                continue;
            }
            let nx = (root_x - root_off_x) / root_ext_x;
            let ny = (root_y - root_off_y) / root_ext_y;
            let nw = root_w / root_ext_x;
            let nh = root_h / root_ext_y;
            if nw <= 0.0 || nh <= 0.0 {
                continue;
            }

            out.push(ShapeInfo {
                x: nx,
                y: ny,
                w: nw,
                h: nh,
                rot: rot_raw / 60000.0,
                fill_color: None,
                stroke_color: None,
                stroke_width: 0,
                geom: ShapeGeom::Image {
                    image_path,
                    mime_type,
                    svg_image_path,
                    src_rect,
                },
                text: None,
            });
        }
        // Ignore `xdr:cxnSp` / text-only elements for this minimal pass.
    }
}

pub(crate) fn parse_shape_anchors(
    drawing_xml: &str,
    theme_colors: &[String],
    theme_ln_widths: &[i64],
    rid_urls: &HashMap<String, String>,
) -> Vec<ShapeAnchor> {
    let Ok(doc) = roxmltree::Document::parse(drawing_xml) else {
        return Vec::new();
    };
    let xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    let mut anchors: Vec<ShapeAnchor> = Vec::new();

    for anchor in doc.descendants() {
        let anchor_tag = anchor.tag_name().name();
        // ECMA-376 §20.5.2: a drawing shape is anchored with either
        // `twoCellAnchor` (from+to) or `oneCellAnchor` (from + a saved `<ext>`
        // size). Excel authors equation text boxes as oneCellAnchor, so we must
        // accept both or those shapes (and their math) are silently dropped.
        if (anchor_tag != "twoCellAnchor" && anchor_tag != "oneCellAnchor")
            || anchor.tag_name().namespace() != Some(xdr_ns)
        {
            continue;
        }

        // Parse from/to anchor rect (shared between grpSp and stand-alone sp paths)
        let (mut from_col, mut from_col_off, mut from_row, mut from_row_off) =
            (0u32, 0i64, 0u32, 0i64);
        let (mut to_col, mut to_col_off, mut to_row, mut to_row_off) = (0u32, 0i64, 0u32, 0i64);
        // ECMA-376 §20.5.2.33 `twoCellAnchor@editAs` — see ImageAnchor parsing
        // path for semantics. `"oneCell"` instructs the renderer to preserve
        // the group's saved EMU size instead of resizing with the cell rect.
        // oneCellAnchor has no `<to>`; its size is the saved EMU `<ext>` (which
        // equals the shape's own xfrm ext), positioned at `<from>` — i.e. the
        // "oneCell" (move-but-don't-size) semantics the renderer already
        // implements via nativeExtCx/Cy. Force editAs accordingly so the
        // renderer sizes from the saved ext rather than a (missing) `to` rect.
        let edit_as = if anchor_tag == "oneCellAnchor" {
            Some("oneCell".to_string())
        } else {
            anchor.attribute("editAs").map(|s| s.to_string())
        };
        let native_ext_cx: i64;
        let native_ext_cy: i64;
        for c in anchor.children() {
            if !c.is_element() {
                continue;
            }
            if c.tag_name().name() == "from" || c.tag_name().name() == "to" {
                let is_from = c.tag_name().name() == "from";
                let mut col: u32 = 0;
                let mut col_off: i64 = 0;
                let mut row: u32 = 0;
                let mut row_off: i64 = 0;
                for cc in c.children() {
                    match (cc.tag_name().name(), cc.text()) {
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
        }

        // Excel wraps a modern shape in an anchor-level `<mc:AlternateContent>`
        // (`<mc:Choice Requires="…">` = the live shape, `<mc:Fallback>` = a
        // legacy fallback). Unwrap to the Choice's content so the grpSp/sp/pic
        // lookups below see the real shape; otherwise the whole drawing — and
        // any equation in it — is silently dropped. (Equations inside the shape
        // text body are handled separately by `push_math_runs_shape`; this is
        // the distinct, shape-level wrapper.)
        let content = anchor
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "AlternateContent")
            .and_then(|ac| {
                ac.children()
                    .find(|n| n.is_element() && n.tag_name().name() == "Choice")
            })
            .unwrap_or(anchor);

        // Two top-level layouts ECMA-376 allows under <xdr:twoCellAnchor>:
        //   (a) <xdr:grpSp> wrapping a tree of nested groups + leaves; and
        //   (b) a single <xdr:sp> / <xdr:pic> directly under the anchor
        //       (no grouping wrapper). The grpSp path uses the group's xfrm
        //       to define the anchor's drawing-coord system; the stand-alone
        //       path treats the shape as filling 100 % of the anchor rect.
        let mut shapes: Vec<ShapeInfo> = Vec::new();
        if let Some(grp) = content
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "grpSp")
        {
            let grp_sp_pr = grp
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "grpSpPr");
            let xfrm = grp_sp_pr
                .and_then(|n| {
                    n.children()
                        .find(|c| c.is_element() && c.tag_name().name() == "xfrm")
                })
                .as_ref()
                .and_then(parse_xfrm);
            let Some(root) = xfrm else {
                continue;
            };
            if !root.has_ch || root.ch_ext_x == 0.0 || root.ch_ext_y == 0.0 {
                continue;
            }

            // Top-level grpSp ext is the group's saved on-sheet EMU size —
            // authoritative when editAs="oneCell".
            native_ext_cx = root.ext_x as i64;
            native_ext_cy = root.ext_y as i64;

            // Map child coords → root coords with the grpSp's own chOff/chExt.
            let csx = root.ext_x / root.ch_ext_x;
            let csy = root.ext_y / root.ch_ext_y;
            let tx = root.off_x - root.ch_off_x * csx;
            let ty = root.off_y - root.ch_off_y * csy;

            collect_shapes(
                &grp,
                root.off_x,
                root.off_y,
                root.ext_x,
                root.ext_y,
                csx,
                csy,
                tx,
                ty,
                theme_colors,
                theme_ln_widths,
                rid_urls,
                &mut shapes,
            );
        } else if let Some(sp) = content.children().find(|n| {
            if !n.is_element() {
                return false;
            }
            let t = n.tag_name().name();
            // A standalone `<xdr:sp>` always renders through the shape path. A
            // standalone `<xdr:pic>` under a `twoCellAnchor`, however, is ALSO a
            // plain image anchor that `parse_drawing_anchors` already emits into
            // `ws.images` — WITH its `<a:srcRect>` crop, svgBlip vector original,
            // and `editAs` size. Capturing it here too would draw the picture a
            // second time as an uncropped full-image shape, overwriting the
            // cropped anchor draw (the sample-27 regression). So a `twoCellAnchor`
            // pic is owned solely by `ws.images`; we only keep a standalone pic
            // for `oneCellAnchor`, which `parse_drawing_anchors` does not handle.
            t == "sp" || (t == "pic" && anchor_tag == "oneCellAnchor")
        }) {
            // Stand-alone sp/pic: the shape's own xfrm gives its absolute EMU
            // rect, but for our rendering pipeline the anchor's from/to
            // already defines the on-sheet rect, and the leaf occupies it
            // 100 %. Build a synthetic root coord-system whose origin matches
            // the shape's xfrm so collect_shapes normalizes the leaf to (0,0)
            // (1,1).
            let sp_pr = sp
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "spPr");
            let Some(sp_pr_node) = sp_pr else {
                continue;
            };
            let xfrm_node = sp_pr_node
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "xfrm");
            let Some(xfrm_n) = xfrm_node else {
                continue;
            };
            let Some(xfrm) = parse_xfrm(&xfrm_n) else {
                continue;
            };
            if xfrm.ext_x == 0.0 || xfrm.ext_y == 0.0 {
                continue;
            }
            native_ext_cx = xfrm.ext_x as i64;
            native_ext_cy = xfrm.ext_y as i64;
            collect_shapes(
                &content,
                xfrm.off_x,
                xfrm.off_y,
                xfrm.ext_x,
                xfrm.ext_y,
                1.0,
                1.0,
                0.0,
                0.0,
                theme_colors,
                theme_ln_widths,
                rid_urls,
                &mut shapes,
            );
        } else {
            continue;
        }

        if shapes.is_empty() {
            continue;
        }

        anchors.push(ShapeAnchor {
            from_col,
            from_col_off,
            from_row,
            from_row_off,
            to_col,
            to_col_off,
            to_row,
            to_row_off,
            edit_as,
            native_ext_cx,
            native_ext_cy,
            shapes,
        });
    }
    anchors
}

pub(crate) fn load_sheet_shape_groups(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    sheet_path: &str,
    theme_colors: &[String],
) -> Vec<ShapeAnchor> {
    // Theme line-style widths for <a:lnRef> resolution (§20.1.4.2.19).
    let theme_ln_widths = crate::parse_theme_ln_widths(archive);
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
    let mut all: Vec<ShapeAnchor> = Vec::new();
    for target in drawing_targets {
        let drawing_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        let Ok(drawing_xml) = read_zip_string(archive, &drawing_path) else {
            continue;
        };
        let rid_urls = build_drawing_rid_urls(archive, &drawing_path);
        all.extend(parse_shape_anchors(
            &drawing_xml,
            theme_colors,
            &theme_ln_widths,
            &rid_urls,
        ));
    }
    all
}

/// Build a `HashMap<rId, zip-path>` for every image (png/jpg/…) target in
/// a drawing's `.rels` file. Used by `collect_shapes` to resolve `<xdr:pic>`
/// leaves inside a group. Mirrors the logic in `parse_drawing_anchors`; the
/// renderer fetches each referenced image's bytes lazily via `extract_image`,
/// so this maps to zip paths rather than inlining base64.
pub(crate) fn build_drawing_rid_urls(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    drawing_path: &str,
) -> HashMap<String, String> {
    let Some((drawing_dir, drawing_file)) = drawing_path.rsplit_once('/') else {
        return HashMap::new();
    };
    let rels_path = format!("{}/_rels/{}.rels", drawing_dir, drawing_file);
    let rels = read_zip_string(archive, &rels_path)
        .ok()
        .map(|xml| parse_rels_map(&xml))
        .unwrap_or_default();

    let mut result: HashMap<String, String> = HashMap::new();
    for (rid, target) in rels {
        let lower = target.to_lowercase();
        // `.svg` is included so the Microsoft svgBlip extension's vector original
        // is collected (it was previously excluded, dropping svg-only pictures).
        if !(lower.ends_with(".png")
            || lower.ends_with(".jpg")
            || lower.ends_with(".jpeg")
            || lower.ends_with(".gif")
            || lower.ends_with(".bmp")
            || lower.ends_with(".webp")
            || lower.ends_with(".svg"))
        {
            continue;
        }
        let media_path = resolve_zip_path(drawing_dir, &target);
        // Only emit the path when the entry actually resolves (preserves the
        // previous behavior of dropping rIds whose bytes are missing).
        // `index_for_name` checks the central directory only — no inflate, unlike
        // the former `read_zip_bytes` which decompressed the entry to discard it.
        if archive.index_for_name(&media_path).is_some() {
            result.insert(rid, media_path);
        }
    }
    result
}

/// Given a sheet path (e.g. "worksheets/sheet1.xml"), locate and parse
/// its drawing(s), and return all image anchors found.
pub(crate) fn load_sheet_images(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    sheet_path: &str, // e.g. "worksheets/sheet1.xml"
) -> Vec<ImageAnchor> {
    // sheet rels path:  xl/worksheets/_rels/sheet1.xml.rels
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else {
        return Vec::new();
    };
    let sheet_rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let Ok(sheet_rels_xml) = read_zip_string(archive, &sheet_rels_path) else {
        return Vec::new();
    };

    // Find all drawing relationships
    let Ok(rels_doc) = roxmltree::Document::parse(&sheet_rels_xml) else {
        return Vec::new();
    };
    let mut drawing_targets: Vec<String> = Vec::new();
    for rel in rels_doc
        .root_element()
        .children()
        .filter(|n| n.is_element())
    {
        let rel_type = rel.attribute("Type").unwrap_or("");
        if rel_type.ends_with("/drawing") {
            if let Some(t) = rel.attribute("Target") {
                drawing_targets.push(t.to_string());
            }
        }
    }
    if drawing_targets.is_empty() {
        return Vec::new();
    }

    let mut all_anchors: Vec<ImageAnchor> = Vec::new();
    for target in drawing_targets {
        // sheet_dir is "worksheets", target typically "../drawings/drawing1.xml"
        // base dir for the drawing = "xl/worksheets" + "../drawings" → "xl/drawings"
        let drawing_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        let Ok(drawing_xml) = read_zip_string(archive, &drawing_path) else {
            continue;
        };
        // Drawing rels:  xl/drawings/_rels/drawing1.xml.rels
        let Some((drawing_dir, drawing_file)) = drawing_path.rsplit_once('/') else {
            continue;
        };
        let drawing_rels_path = format!("{}/_rels/{}.rels", drawing_dir, drawing_file);
        let drawing_rels = read_zip_string(archive, &drawing_rels_path)
            .ok()
            .map(|xml| parse_rels_map(&xml))
            .unwrap_or_default();

        let mut anchors = parse_drawing_anchors(&drawing_xml, &drawing_rels, drawing_dir, archive);
        all_anchors.append(&mut anchors);
    }
    all_anchors
}

#[cfg(test)]
mod math_tests {
    use super::*;
    use ooxml_common::math::nodes_to_text;

    const NS: &str = r#"xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
        xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math""#;

    /// Excel stores "Insert > Equation" as OMML inside the shared DrawingML
    /// `<xdr:txBody>` grammar — a block equation as `mc:AlternateContent` →
    /// `a14:m` → `m:oMathPara` (ECMA-376 §22.1). The Fallback (a rasterized
    /// picture) must be ignored so the equation isn't double-rendered.
    #[test]
    fn parses_block_math_from_alternatecontent() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:bodyPr/>
              <a:p>
                <mc:AlternateContent>
                  <mc:Choice Requires="a14"><a14:m>
                    <m:oMathPara><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara>
                  </a14:m></mc:Choice>
                  <mc:Fallback><a:r><a:t>fallback</a:t></a:r></mc:Fallback>
                </mc:AlternateContent>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        assert_eq!(text.paragraphs.len(), 1);
        let runs = &text.paragraphs[0].runs;
        assert_eq!(runs.len(), 1, "one math run, fallback ignored");
        match &runs[0] {
            ShapeTextRun::Math { display, nodes, .. } => {
                assert!(*display, "oMathPara → display (block) math");
                assert_eq!(nodes_to_text(nodes), "x");
            }
            other => panic!("expected Math run, got {other:?}"),
        }
    }

    /// Inline math is a bare `a14:m` (local name "m") directly under `a:p`,
    /// holding `m:oMath` (not oMathPara). It extracts as inline (display:false)
    /// and inherits size + colour from the math run's `a:rPr`.
    #[test]
    fn parses_inline_bare_a14m_with_size_and_color() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:bodyPr/>
              <a:p>
                <a14:m><m:oMath><m:r>
                  <a:rPr sz="2800" i="1"><a:solidFill><a:srgbClr val="7030A0"/></a:solidFill></a:rPr>
                  <m:t>n</m:t>
                </m:r></m:oMath></a14:m>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        let runs = &text.paragraphs[0].runs;
        assert_eq!(runs.len(), 1);
        match &runs[0] {
            ShapeTextRun::Math {
                display,
                font_size,
                color,
                nodes,
            } => {
                assert!(!*display, "bare a14:m + m:oMath → inline");
                assert_eq!(*font_size, Some(28.0), "size from math run rPr sz/100");
                assert_eq!(
                    color.as_deref(),
                    Some("#7030A0"),
                    "colour from math run rPr solidFill (xlsx #-prefixed convention)"
                );
                assert_eq!(nodes_to_text(nodes), "n");
            }
            other => panic!("expected Math run, got {other:?}"),
        }
    }

    /// Text + break + math coexist in one paragraph; the run order is preserved
    /// and `<a:br>` becomes a `Break` variant.
    #[test]
    fn mixes_text_break_and_math_runs() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:p>
                <a:r><a:rPr sz="1100"/><a:t>E=</a:t></a:r>
                <a14:m><m:oMath><m:r><m:t>mc</m:t></m:r></m:oMath></a14:m>
                <a:br/>
                <a:r><a:t>done</a:t></a:r>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        let runs = &text.paragraphs[0].runs;
        assert_eq!(runs.len(), 4, "text, math, break, text");
        assert!(matches!(runs[0], ShapeTextRun::Text { .. }));
        assert!(matches!(runs[1], ShapeTextRun::Math { display: false, .. }));
        assert!(matches!(runs[2], ShapeTextRun::Break));
        assert!(matches!(runs[3], ShapeTextRun::Text { .. }));
    }

    /// `<a:pPr rtl="1">` (ECMA-376 §21.1.2.2.7) marks the paragraph as
    /// right-to-left.
    #[test]
    fn parses_paragraph_rtl_attribute() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:p>
                <a:pPr rtl="1"/>
                <a:r><a:t>שלום</a:t></a:r>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        assert_eq!(text.paragraphs.len(), 1);
        assert!(text.paragraphs[0].rtl, "pPr rtl=\"1\" → rtl true");
    }

    /// Absent `@rtl` defaults to false (left-to-right).
    #[test]
    fn paragraph_rtl_defaults_false_when_absent() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:p>
                <a:pPr algn="ctr"/>
                <a:r><a:t>hello</a:t></a:r>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        assert_eq!(text.paragraphs.len(), 1);
        assert!(!text.paragraphs[0].rtl, "absent @rtl → rtl false");
    }

    /// `<a:pPr marL marR indent>` (ECMA-376 §21.1.2.2.7,
    /// `CT_TextParagraphProperties`) are the direct paragraph indent attributes
    /// in EMU. `indent` may be negative (hanging). They parse onto the
    /// `ShapeParagraph` as `Option<i64>` so an absent attribute stays `None`.
    #[test]
    fn parses_paragraph_indent_attributes() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:p>
                <a:pPr marL="457200" marR="91440" indent="-228600"/>
                <a:r><a:t>indented</a:t></a:r>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        assert_eq!(text.paragraphs.len(), 1);
        let p = &text.paragraphs[0];
        assert_eq!(p.mar_l, Some(457200), "marL parses as EMU i64");
        assert_eq!(p.mar_r, Some(91440), "marR parses as EMU i64");
        assert_eq!(
            p.indent,
            Some(-228600),
            "indent parses (negative = hanging)"
        );
    }

    /// Absent indent attributes (or absent `<a:pPr>` entirely) leave all three
    /// fields `None` so the JSON output stays byte-identical (additive).
    #[test]
    fn paragraph_indent_defaults_none_when_absent() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:p>
                <a:r><a:t>plain</a:t></a:r>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        assert_eq!(text.paragraphs.len(), 1);
        let p = &text.paragraphs[0];
        assert_eq!(p.mar_l, None, "absent marL → None");
        assert_eq!(p.mar_r, None, "absent marR → None");
        assert_eq!(p.indent, None, "absent indent → None");

        // None of the three keys should appear in the serialized JSON.
        let v: serde_json::Value = serde_json::to_value(p).unwrap();
        assert!(v.get("marL").is_none(), "marL omitted when None");
        assert!(v.get("marR").is_none(), "marR omitted when None");
        assert!(v.get("indent").is_none(), "indent omitted when None");
        // spaceLine is additive/omitted when absent; autoFit always present.
        assert!(p.space_line.is_none(), "absent lnSpc → None");
        assert!(v.get("spaceLine").is_none(), "spaceLine omitted when None");
        let vt: serde_json::Value = serde_json::to_value(&text).unwrap();
        assert_eq!(vt["autoFit"], "none", "no autofit child → autoFit=none");
        assert!(
            vt.get("fontScale").is_none(),
            "fontScale omitted when unset"
        );
        assert!(
            vt.get("lnSpcReduction").is_none(),
            "lnSpcReduction omitted when unset"
        );
    }

    /// `<a:pPr>/<a:lnSpc>` line spacing (ECMA-376 §21.1.2.2.5): spcPct → a
    /// percent SpaceLine (raw @val), spcPts → a points SpaceLine (raw hundredths
    /// of a point divided by 100). Mirrors the pptx SpaceLine JSON contract.
    #[test]
    fn parses_lnspc_pct_and_pts() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:p>
                <a:pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc></a:pPr>
                <a:r><a:t>pct</a:t></a:r>
              </a:p>
              <a:p>
                <a:pPr><a:lnSpc><a:spcPts val="1800"/></a:lnSpc></a:pPr>
                <a:r><a:t>pts</a:t></a:r>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        assert_eq!(text.paragraphs.len(), 2);

        let vp0: serde_json::Value = serde_json::to_value(&text.paragraphs[0]).unwrap();
        assert_eq!(vp0["spaceLine"]["type"], "pct");
        assert_eq!(vp0["spaceLine"]["val"], 150000.0);

        let vp1: serde_json::Value = serde_json::to_value(&text.paragraphs[1]).unwrap();
        assert_eq!(vp1["spaceLine"]["type"], "pts");
        // 1800 hundredths of a point → 18 pt.
        assert_eq!(vp1["spaceLine"]["val"], 18.0);
    }

    /// `<a:bodyPr>/<a:normAutofit>` (ECMA-376 §21.1.2.1.3): autoFit="norm" plus
    /// the stored fontScale / lnSpcReduction (1000ths of a percent → fraction).
    #[test]
    fn parses_normautofit_scales() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:bodyPr><a:normAutofit fontScale="62500" lnSpcReduction="20000"/></a:bodyPr>
              <a:p><a:r><a:t>x</a:t></a:r></a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        let v: serde_json::Value = serde_json::to_value(&text).unwrap();
        assert_eq!(v["autoFit"], "norm");
        assert_eq!(v["fontScale"], 0.625);
        assert_eq!(v["lnSpcReduction"], 0.2);
    }

    /// `<a:bodyPr>/<a:spAutoFit>` → autoFit="sp" (no scales); `<a:noAutofit>` →
    /// autoFit="none". Both leave rendering unchanged (renderer applies neither).
    #[test]
    fn parses_spautofit_and_noautofit() {
        let sp_xml = format!(
            r#"<xdr:txBody {NS}>
              <a:bodyPr><a:spAutoFit/></a:bodyPr>
              <a:p><a:r><a:t>x</a:t></a:r></a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&sp_xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        let v: serde_json::Value = serde_json::to_value(&text).unwrap();
        assert_eq!(v["autoFit"], "sp");
        assert!(
            v.get("fontScale").is_none(),
            "spAutoFit stores no fontScale"
        );

        let no_xml = format!(
            r#"<xdr:txBody {NS}>
              <a:bodyPr><a:noAutofit/></a:bodyPr>
              <a:p><a:r><a:t>x</a:t></a:r></a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&no_xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        let v: serde_json::Value = serde_json::to_value(&text).unwrap();
        assert_eq!(v["autoFit"], "none");
    }

    /// `ShapeTextRun` uses an enum-level `#[serde(tag = "type", rename_all =
    /// "camelCase")]`, which renames only the variant tags — not the fields. The
    /// `Text` variant's `font_face` and the `Math` variant's `font_size`
    /// therefore need per-variant `rename_all` to serialize as the camelCase
    /// keys the TS renderer reads (`run.fontFace` / `run.fontSize`). This locks
    /// the JSON contract so the keys never regress to snake_case (which the
    /// renderer reads as `undefined`). Same bug class as the pptx serde fix
    /// (PR #489) and the xlsx ArcTo fix (PR #491).
    #[test]
    fn serializes_text_run_font_face_as_camel_case() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:p>
                <a:r>
                  <a:rPr sz="1400"><a:latin typeface="Calibri"/></a:rPr>
                  <a:t>hi</a:t>
                </a:r>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        let run = &text.paragraphs[0].runs[0];
        assert!(
            matches!(run, ShapeTextRun::Text { .. }),
            "expected Text run"
        );

        let v: serde_json::Value = serde_json::to_value(run).unwrap();
        assert_eq!(v["type"], "text", "tag key is `type`");
        assert_eq!(
            v["fontFace"], "Calibri",
            "font_face must serialize as camelCase `fontFace` (renderer reads run.fontFace)"
        );
        assert!(
            v.get("font_face").is_none(),
            "snake_case `font_face` must not appear"
        );
    }

    /// ECMA-376 §21.1.2.3.1 `<a:ea>` (East-Asian) / `<a:cs>` (complex-script)
    /// typefaces parse onto the run alongside `<a:latin>`. The common Japanese
    /// encoding sets Meiryo on `<a:ea>` while leaving `<a:latin>` default; the
    /// renderer floors the line box by the tabled face. They serialize as
    /// camelCase `fontFaceEa` / `fontFaceCs` and are additive (omitted when
    /// absent), so shapes without ea/cs stay byte-identical.
    #[test]
    fn parses_ea_and_cs_typefaces() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:p>
                <a:r>
                  <a:rPr sz="1400">
                    <a:ea typeface="Meiryo"/>
                    <a:cs typeface="Sakkal Majalla"/>
                  </a:rPr>
                  <a:t>あ</a:t>
                </a:r>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        let run = &text.paragraphs[0].runs[0];
        match run {
            ShapeTextRun::Text {
                font_face,
                font_face_ea,
                font_face_cs,
                ..
            } => {
                assert!(font_face.is_none(), "latin left default → font_face None");
                assert_eq!(
                    font_face_ea.as_deref(),
                    Some("Meiryo"),
                    "<a:ea> → fontFaceEa"
                );
                assert_eq!(
                    font_face_cs.as_deref(),
                    Some("Sakkal Majalla"),
                    "<a:cs> → fontFaceCs"
                );
            }
            other => panic!("expected Text run, got {other:?}"),
        }

        let v: serde_json::Value = serde_json::to_value(run).unwrap();
        assert_eq!(
            v["fontFaceEa"], "Meiryo",
            "serializes as camelCase fontFaceEa"
        );
        assert_eq!(
            v["fontFaceCs"], "Sakkal Majalla",
            "serializes as camelCase fontFaceCs"
        );
    }

    /// Additive guarantee: a run with no `<a:ea>` / `<a:cs>` omits both keys
    /// from JSON entirely (serde `skip_serializing_if = "Option::is_none"`), so
    /// existing shape JSON stays byte-identical.
    #[test]
    fn omits_ea_and_cs_when_absent() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:p>
                <a:r>
                  <a:rPr sz="1400"><a:latin typeface="Calibri"/></a:rPr>
                  <a:t>hi</a:t>
                </a:r>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        let run = &text.paragraphs[0].runs[0];
        let v: serde_json::Value = serde_json::to_value(run).unwrap();
        assert!(
            v.get("fontFaceEa").is_none(),
            "fontFaceEa omitted when absent"
        );
        assert!(
            v.get("fontFaceCs").is_none(),
            "fontFaceCs omitted when absent"
        );
    }

    /// Companion to the Text-run case: the `Math` variant's `font_size` must
    /// serialize as `fontSize` (renderer reads `run.fontSize`). Without the
    /// per-variant `rename_all` the equation falls back to the inherited size.
    #[test]
    fn serializes_math_run_font_size_as_camel_case() {
        let xml = format!(
            r#"<xdr:txBody {NS}>
              <a:p>
                <a14:m><m:oMath><m:r>
                  <a:rPr sz="2800"/>
                  <m:t>n</m:t>
                </m:r></m:oMath></a14:m>
              </a:p>
            </xdr:txBody>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let text = parse_tx_body(&doc.root_element(), &[]).expect("txBody parses");
        let run = &text.paragraphs[0].runs[0];
        assert!(
            matches!(run, ShapeTextRun::Math { .. }),
            "expected Math run"
        );

        let v: serde_json::Value = serde_json::to_value(run).unwrap();
        assert_eq!(v["type"], "math", "tag key is `type`");
        assert_eq!(
            v["fontSize"], 28.0,
            "font_size must serialize as camelCase `fontSize` (renderer reads run.fontSize)"
        );
        assert!(
            v.get("font_size").is_none(),
            "snake_case `font_size` must not appear"
        );
    }
}

#[cfg(test)]
mod style_lnref_tests {
    use super::*;

    const NS: &str = r#"xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#;

    /// dk1, lt1, dk2, lt2, accent1 — enough slots for schemeClr accent1 (idx 4).
    fn theme() -> Vec<String> {
        ["#000000", "#FFFFFF", "#222222", "#EEEEEE", "#4472C4"]
            .iter()
            .map(|s| s.to_string())
            .collect()
    }

    fn shape_of(sp_pr_inner: &str, style: &str) -> (Option<String>, i64) {
        let xml = format!(
            r#"<xdr:wsDr {NS}><xdr:twoCellAnchor>
              <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
              <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
              <xdr:sp>
                <xdr:nvSpPr><xdr:cNvPr id="2" name="P"/><xdr:cNvSpPr/></xdr:nvSpPr>
                <xdr:spPr>
                  <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                  <a:prstGeom prst="parallelogram"><a:avLst/></a:prstGeom>
                  {sp_pr_inner}
                </xdr:spPr>
                {style}
              </xdr:sp>
              <xdr:clientData/>
            </xdr:twoCellAnchor></xdr:wsDr>"#
        );
        let anchors =
            parse_shape_anchors(&xml, &theme(), &[6_350, 12_700, 19_050], &HashMap::new());
        assert_eq!(anchors.len(), 1, "one anchor expected");
        let s = &anchors[0].shapes[0];
        (s.stroke_color.clone(), s.stroke_width)
    }

    /// §20.1.4.2.19 — with no explicit `<a:ln>`, the outline comes from the
    /// theme style matrix: lnRef idx picks the lnStyleLst width, the lnRef's
    /// child scheme colour substitutes phClr.
    #[test]
    fn lnref_supplies_stroke_when_sppr_has_no_ln() {
        let (color, width) = shape_of(
            "",
            r#"<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef></xdr:style>"#,
        );
        assert_eq!(color.as_deref(), Some("#4472C4"));
        assert_eq!(width, 12_700);
    }

    /// An explicit `<a:ln><a:noFill/>` is a hard "no outline" and must beat
    /// the theme lnRef fallback.
    #[test]
    fn explicit_nofill_ln_beats_lnref() {
        let (color, width) = shape_of(
            r#"<a:ln><a:noFill/></a:ln>"#,
            r#"<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef></xdr:style>"#,
        );
        assert!(color.is_none());
        assert_eq!(width, 0);
    }

    /// An explicit `<a:ln>` with its own fill/width wins over lnRef.
    #[test]
    fn explicit_ln_beats_lnref() {
        let (color, width) = shape_of(
            r#"<a:ln w="28575"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln>"#,
            r#"<xdr:style><a:lnRef idx="3"><a:schemeClr val="accent1"/></a:lnRef></xdr:style>"#,
        );
        assert_eq!(color.as_deref(), Some("#00FF00"));
        assert_eq!(width, 28_575);
    }

    /// lnRef idx out of range (or theme without lnStyleLst) still draws with
    /// the CT_LineProperties default width (§20.1.2.2.24).
    #[test]
    fn lnref_out_of_range_uses_default_width() {
        let (color, width) = shape_of(
            "",
            r#"<xdr:style><a:lnRef idx="9"><a:schemeClr val="accent1"/></a:lnRef></xdr:style>"#,
        );
        assert_eq!(color.as_deref(), Some("#4472C4"));
        assert_eq!(width, 9_525);
    }

    /// A standalone `<xdr:pic>` directly under a `twoCellAnchor` is a plain image
    /// anchor owned by `ws.images` (parse_drawing_anchors, WITH its `<a:srcRect>`
    /// crop). `parse_shape_anchors` must NOT also capture it — otherwise the
    /// renderer draws the picture twice and the uncropped shape draw overwrites
    /// the cropped anchor draw (the sample-27 double-draw regression).
    #[test]
    fn standalone_twocellanchor_pic_is_not_a_shape_anchor() {
        let xml = format!(
            r#"<xdr:wsDr {NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor editAs="oneCell">
              <xdr:from><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
              <xdr:to><xdr:col>11</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
              <xdr:pic>
                <xdr:nvPicPr><xdr:cNvPr id="2" name="P"/><xdr:cNvPicPr/></xdr:nvPicPr>
                <xdr:blipFill><a:blip r:embed="rId1"/><a:srcRect l="32560" r="3829"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
                <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2905125" cy="2181225"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
              </xdr:pic>
              <xdr:clientData/>
            </xdr:twoCellAnchor></xdr:wsDr>"#
        );
        let anchors = parse_shape_anchors(&xml, &theme(), &[6_350], &HashMap::new());
        assert!(
            anchors.is_empty(),
            "twoCellAnchor pic belongs to ws.images, not ws.shape_groups: {anchors:?}"
        );
    }

    /// A standalone `<xdr:sp>` under the same anchor IS a shape anchor — only the
    /// `pic` is deduped, never a real shape.
    #[test]
    fn standalone_sp_is_still_a_shape_anchor() {
        let (color, _width) = shape_of(
            r#"<a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>"#,
            "",
        );
        assert_eq!(color.as_deref(), Some("#FF0000"), "sp still captured");
    }

    /// A `oneCellAnchor` pic is NOT handled by `parse_drawing_anchors` (which only
    /// scans `twoCellAnchor`), so the shape path must keep capturing it — else the
    /// image is dropped entirely. Its `<a:srcRect>` crop must also be surfaced on
    /// the `ShapeGeom::Image` so the renderer can crop it like a top-level anchor.
    #[test]
    fn standalone_onecellanchor_pic_is_still_captured_with_crop() {
        let xml = format!(
            r#"<xdr:wsDr {NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:oneCellAnchor>
              <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
              <xdr:ext cx="914400" cy="914400"/>
              <xdr:pic>
                <xdr:nvPicPr><xdr:cNvPr id="2" name="P"/><xdr:cNvPicPr/></xdr:nvPicPr>
                <xdr:blipFill><a:blip r:embed="rId1"/><a:srcRect l="10000" t="20000"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
                <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
              </xdr:pic>
              <xdr:clientData/>
            </xdr:oneCellAnchor></xdr:wsDr>"#
        );
        let mut rids = HashMap::new();
        rids.insert("rId1".to_string(), "xl/media/image1.png".to_string());
        let anchors = parse_shape_anchors(&xml, &theme(), &[6_350], &rids);
        assert_eq!(anchors.len(), 1, "oneCellAnchor pic must still be captured");
        match &anchors[0].shapes[0].geom {
            ShapeGeom::Image { src_rect, .. } => {
                let sr = src_rect
                    .as_ref()
                    .expect("leaf pic surfaces its srcRect crop");
                assert!((sr.l - 0.1).abs() < 1e-9);
                assert!((sr.t - 0.2).abs() < 1e-9);
            }
            other => panic!("expected an image-geom shape, got {other:?}"),
        }
    }
}

#[cfg(test)]
mod geom_tests {
    use super::*;

    const NS: &str = r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#;

    fn geom_of(sp_pr_xml: &str) -> ShapeGeom {
        let doc = roxmltree::Document::parse(sp_pr_xml).unwrap();
        parse_sp_geom(&doc.root_element()).expect("sp_pr has geometry")
    }

    /// A preset with no `<a:avLst>` yields an empty adjust vector — the engine
    /// then falls back to each shape's declared defaults.
    #[test]
    fn preset_without_avlst_has_empty_adj() {
        let geom = geom_of(&format!(
            r#"<a:spPr {NS}><a:prstGeom prst="parallelogram"/></a:spPr>"#
        ));
        match geom {
            ShapeGeom::Preset { name, adj } => {
                assert_eq!(name, "parallelogram");
                assert!(adj.is_empty(), "no avLst → empty adj, got {adj:?}");
            }
            other => panic!("expected Preset, got {other:?}"),
        }
    }

    /// `<a:gd name="adj" fmla="val X"/>` is read into the first adjust slot.
    #[test]
    fn preset_reads_single_named_adj() {
        let geom = geom_of(&format!(
            r#"<a:spPr {NS}><a:prstGeom prst="parallelogram">
                 <a:avLst><a:gd name="adj" fmla="val 41667"/></a:avLst>
               </a:prstGeom></a:spPr>"#
        ));
        match geom {
            ShapeGeom::Preset { name, adj } => {
                assert_eq!(name, "parallelogram");
                assert_eq!(adj, vec![Some(41667.0)]);
            }
            other => panic!("expected Preset, got {other:?}"),
        }
    }

    /// Multiple named handles (`adj1`/`adj2`) land in declaration order; a gap
    /// produces a `None` placeholder so later indices stay aligned.
    #[test]
    fn preset_reads_named_adj1_adj2_and_fills_gaps() {
        let geom = geom_of(&format!(
            r#"<a:spPr {NS}><a:prstGeom prst="wedgeRectCallout">
                 <a:avLst>
                   <a:gd name="adj1" fmla="val -20000"/>
                   <a:gd name="adj3" fmla="val 55000"/>
                 </a:avLst>
               </a:prstGeom></a:spPr>"#
        ));
        match geom {
            ShapeGeom::Preset { adj, .. } => {
                // adj1 set, adj2 missing → None, adj3 set, rest trimmed.
                assert_eq!(adj, vec![Some(-20000.0), None, Some(55000.0)]);
            }
            other => panic!("expected Preset, got {other:?}"),
        }
    }

    /// Unnamed `<a:gd>` handles fall back to declaration position.
    #[test]
    fn preset_positional_fallback_for_unnamed_gd() {
        let geom = geom_of(&format!(
            r#"<a:spPr {NS}><a:prstGeom prst="roundRect">
                 <a:avLst><a:gd fmla="val 16667"/></a:avLst>
               </a:prstGeom></a:spPr>"#
        ));
        match geom {
            ShapeGeom::Preset { adj, .. } => assert_eq!(adj, vec![Some(16667.0)]),
            other => panic!("expected Preset, got {other:?}"),
        }
    }
}

#[cfg(test)]
mod blip_svg_tests {
    // Microsoft 2016 SVG extension (MS-ODRAWXML) on a `<xdr:pic>` blip: the real
    // vector image rides in `<a:blip><a:extLst><a:ext
    // uri="{96DAC541-…}"><asvg:svgBlip r:embed>`, while `<a:blip r:embed>` is a
    // PNG/JPEG *fallback* (which may be absent for a pure-SVG insert). The parser
    // must keep emitting the raster fallback as `image_path` (regression-safe)
    // while additionally surfacing the SVG body in `svg_image_path`, and must
    // never drop a picture that carries only the svgBlip. The model now carries
    // zip paths + mime (no inlined base64); the renderer fetches bytes lazily via
    // `extract_image`. Mirrors the pptx/docx parser tests.
    use super::*;

    // 1×1 transparent PNG (smallest valid PNG).
    const PNG_1X1: &[u8] = &[
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F,
        0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x62, 0x00,
        0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49,
        0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];
    const SVG: &[u8] =
        br##"<svg xmlns="http://www.w3.org/2000/svg" width="2" height="2"><rect width="2" height="2" fill="#0a0"/></svg>"##;

    /// Build a tiny zip with the two media parts a `<xdr:pic>` blip references
    /// (a PNG fallback and an SVG body) under `xl/media/`, so
    /// `parse_drawing_anchors` can be driven with a hand-rolled rels map.
    fn build_media_zip(png: &[u8], svg: &[u8]) -> Vec<u8> {
        use std::io::Write;
        use zip::write::SimpleFileOptions;
        let mut buf = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zw = zip::ZipWriter::new(cursor);
            let opts = SimpleFileOptions::default();
            zw.start_file("xl/media/image1.png", opts).unwrap();
            zw.write_all(png).unwrap();
            zw.start_file("xl/media/image2.svg", opts).unwrap();
            zw.write_all(svg).unwrap();
            zw.finish().unwrap();
        }
        buf
    }

    /// `<xdr:wsDr>` wrapping a single `twoCellAnchor` picture whose `<a:blip>`
    /// inner XML is supplied by the caller (so each test varies only the blip).
    fn drawing_xml(blip_inner: &str) -> String {
        format!(
            r#"<xdr:wsDr
  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="2" name="P"/><xdr:cNvPicPr/></xdr:nvPicPr>
      <xdr:blipFill>{blip_inner}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
      <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="300000" cy="300000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#
        )
    }

    fn parse_one(blip_inner: &str, rels: &HashMap<String, String>) -> ImageAnchor {
        let xml = drawing_xml(blip_inner);
        let data = build_media_zip(PNG_1X1, SVG);
        let cursor = Cursor::new(data.as_slice());
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let anchors = parse_drawing_anchors(&xml, rels, "xl/drawings", &mut archive);
        assert_eq!(anchors.len(), 1, "exactly one picture anchor expected");
        anchors.into_iter().next().unwrap()
    }

    /// A `<xdr:pic>` carrying a PNG fallback blip plus an `asvg:svgBlip` extension
    /// must keep the PNG on `image_path` and surface the SVG on `svg_image_path`,
    /// both as zip paths with mime derived from the extension (no base64). The
    /// svgBlip uses a different prefix (asvg:) on purpose — matching is by
    /// namespace-local name, so the prefix must not matter.
    #[test]
    fn picture_with_svg_blip_extension_emits_both_paths() {
        let blip = r#"<a:blip r:embed="rIdPng">
          <a:extLst>
            <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
              <asvg:svgBlip r:embed="rIdSvg"/>
            </a:ext>
          </a:extLst>
        </a:blip>"#;
        let mut rels = HashMap::new();
        rels.insert("rIdPng".to_string(), "../media/image1.png".to_string());
        rels.insert("rIdSvg".to_string(), "../media/image2.svg".to_string());

        let anchor = parse_one(blip, &rels);

        // PNG fallback preserved on image_path (regression-safe). No base64.
        assert_eq!(anchor.image_path, "xl/media/image1.png");
        assert_eq!(anchor.mime_type, "image/png");
        assert!(
            !anchor.image_path.contains(";base64,"),
            "image_path must be a zip path, not a data URL"
        );
        // The SVG part is surfaced separately as a zip path.
        assert_eq!(
            anchor.svg_image_path.as_deref(),
            Some("xl/media/image2.svg")
        );
        // native ext is retained alongside the new path refs.
        assert_eq!(anchor.native_ext_cx, 300000);
        assert_eq!(anchor.native_ext_cy, 300000);
    }

    /// End-to-end wiring guard: a `<xdr:pic>` whose `<xdr:blipFill>` carries an
    /// `<a:srcRect>` sibling of the blip must surface the crop on the parsed
    /// `ImageAnchor.src_rect` (sample-27's horizontal crop). Catches a regression
    /// that drops the `parse_src_rect(bf)` wiring from the pic branch.
    #[test]
    fn picture_with_src_rect_surfaces_crop() {
        let blip = r#"<a:blip r:embed="rIdPng"/><a:srcRect l="32560" r="3829"/>"#;
        let mut rels = HashMap::new();
        rels.insert("rIdPng".to_string(), "../media/image1.png".to_string());

        let anchor = parse_one(blip, &rels);

        let sr = anchor
            .src_rect
            .as_ref()
            .expect("srcRect surfaced on the anchor");
        assert!((sr.l - 0.3256).abs() < 1e-9);
        assert!((sr.r - 0.03829).abs() < 1e-9);
        assert_eq!(sr.t, 0.0);
        assert_eq!(sr.b, 0.0);

        // It serializes as camelCase `srcRect` so the TS renderer can read it.
        let json = serde_json::to_string(&anchor).unwrap();
        assert!(json.contains("\"srcRect\""), "emits srcRect: {json}");
    }

    /// An uncropped `<xdr:pic>` (no `<a:srcRect>`) leaves `src_rect == None`, and
    /// the serialized JSON omits the key entirely (skip_serializing_if), so the
    /// common case stays on the cheap full-blip draw path.
    #[test]
    fn picture_without_src_rect_omits_crop() {
        let blip = r#"<a:blip r:embed="rIdPng"/>"#;
        let mut rels = HashMap::new();
        rels.insert("rIdPng".to_string(), "../media/image1.png".to_string());

        let anchor = parse_one(blip, &rels);

        assert!(anchor.src_rect.is_none());
        let json = serde_json::to_string(&anchor).unwrap();
        assert!(
            !json.contains("srcRect"),
            "omits srcRect when absent: {json}"
        );
    }

    /// A `<xdr:pic>` whose `<a:blip>` carries ONLY the `asvg:svgBlip` extension —
    /// no raster `r:embed` fallback (an icon inserted as a pure SVG) — must still
    /// parse. Previously the media filter excluded `.svg` and the resolution
    /// required a raster embed, so the whole picture was silently dropped.
    #[test]
    fn picture_with_only_svg_blip_and_no_raster_embed_still_parses() {
        let blip = r#"<a:blip>
          <a:extLst>
            <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
              <asvg:svgBlip r:embed="rIdSvg"/>
            </a:ext>
          </a:extLst>
        </a:blip>"#;
        let mut rels = HashMap::new();
        rels.insert("rIdSvg".to_string(), "../media/image2.svg".to_string());

        let anchor = parse_one(blip, &rels);

        // With no raster embed, image_path falls back to the SVG path itself so
        // the element is always drawable …
        assert_eq!(
            anchor.image_path, "xl/media/image2.svg",
            "image_path must fall back to the SVG path when no raster is embedded"
        );
        assert_eq!(
            anchor.mime_type, "image/svg+xml",
            "mime must follow the fallback SVG path's extension"
        );
        // … and svg_image_path carries the same vector source.
        assert_eq!(
            anchor.svg_image_path.as_deref(),
            Some("xl/media/image2.svg"),
            "svg_image_path must be Some for a pure-SVG picture"
        );
    }

    /// A plain `<xdr:pic>` with no svgBlip extension must leave `svg_image_path`
    /// as None (and still emit the PNG `image_path`) — guards against the new
    /// branch firing spuriously.
    #[test]
    fn picture_without_svg_blip_has_no_svg_path() {
        let blip = r#"<a:blip r:embed="rIdPng"/>"#;
        let mut rels = HashMap::new();
        rels.insert("rIdPng".to_string(), "../media/image1.png".to_string());

        let anchor = parse_one(blip, &rels);
        assert_eq!(anchor.image_path, "xl/media/image1.png");
        assert_eq!(anchor.mime_type, "image/png");
        assert!(
            anchor.svg_image_path.is_none(),
            "svg_image_path must be None without an svgBlip extension"
        );
    }

    /// Struct-serialize guard: `ImageAnchor` serializes the new camelCase path
    /// refs (`imagePath`/`mimeType`/`svgImagePath`) and retains
    /// `nativeExtCx`/`nativeExtCy`, and never emits the old `dataUrl` /
    /// `;base64,` form. Fixture-free so it always runs in CI.
    #[test]
    fn image_anchor_serializes_path_refs_not_data_url() {
        let anchor = ImageAnchor {
            from_col: 1,
            from_col_off: 0,
            from_row: 1,
            from_row_off: 0,
            to_col: 3,
            to_col_off: 0,
            to_row: 5,
            to_row_off: 0,
            edit_as: Some("oneCell".to_string()),
            native_ext_cx: 300000,
            native_ext_cy: 200000,
            image_path: "xl/media/image1.png".to_string(),
            mime_type: "image/png".to_string(),
            svg_image_path: Some("xl/media/image2.svg".to_string()),
            src_rect: None,
        };
        let json = serde_json::to_string(&anchor).unwrap();
        assert!(json.contains("\"imagePath\":\"xl/media/image1.png\""));
        assert!(json.contains("\"mimeType\":\"image/png\""));
        assert!(json.contains("\"svgImagePath\":\"xl/media/image2.svg\""));
        assert!(json.contains("\"nativeExtCx\":300000"));
        assert!(json.contains("\"nativeExtCy\":200000"));
        assert!(!json.contains("dataUrl"), "must not emit dataUrl: {json}");
        assert!(!json.contains(";base64,"), "must not inline base64: {json}");
    }

    /// Struct-serialize guard for the group-leaf `ShapeGeom::Image` variant:
    /// same camelCase path refs, no `dataUrl` / base64.
    #[test]
    fn shape_geom_image_serializes_path_refs_not_data_url() {
        let geom = ShapeGeom::Image {
            image_path: "xl/media/image1.png".to_string(),
            mime_type: "image/png".to_string(),
            svg_image_path: None,
            src_rect: None,
        };
        let json = serde_json::to_string(&geom).unwrap();
        assert!(json.contains("\"type\":\"image\""));
        assert!(json.contains("\"imagePath\":\"xl/media/image1.png\""));
        assert!(json.contains("\"mimeType\":\"image/png\""));
        assert!(!json.contains("dataUrl"), "must not emit dataUrl: {json}");
        assert!(!json.contains(";base64,"), "must not inline base64: {json}");
    }
}

#[cfg(test)]
mod custom_path_arc_tests {
    use super::*;

    const NS: &str = r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#;

    /// `PathCmd`'s enum-level `rename_all = "camelCase"` (tag = "op") renames
    /// only the variant tags, not the fields. Without a per-variant
    /// `rename_all`, `ArcTo { st_ang, sw_ang }` serialized as snake_case keys
    /// `st_ang`/`sw_ang`, which the TS renderer never reads — it reads
    /// `cmd.stAng`/`cmd.swAng` (`renderer.ts`), so the values came back
    /// `undefined` → `NaN` and the arc failed to draw. Same root cause as the
    /// pptx fix in PR #489. This guards the JSON keys the renderer relies on.
    #[test]
    fn arc_to_serializes_angle_fields_as_camelcase() {
        // Non-degenerate arc (wR/hR > 0) — the renderer's degenerate-arc guard
        // (`if (rx <= 0 || ry <= 0) break;`) short-circuits before the angles
        // are read, so only non-degenerate arcs surface the missing keys.
        let xml = format!(
            r#"<a:path {NS} w="100" h="100">
                 <a:moveTo><a:pt x="0" y="50"/></a:moveTo>
                 <a:arcTo wR="50" hR="50" stAng="0" swAng="5400000"/>
               </a:path>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let path = parse_custom_path(&doc.root_element());

        // The parser keeps angles as raw 60000ths of a degree (the xlsx
        // convention — the TS renderer divides by 60000). swAng = 5_400_000 =
        // 90°. wR/hR stay in path-coord units.
        let arc = path
            .commands
            .iter()
            .find(|c| matches!(c, PathCmd::ArcTo { .. }))
            .expect("path contains an arcTo command");
        match arc {
            PathCmd::ArcTo {
                wr,
                hr,
                st_ang,
                sw_ang,
            } => {
                assert_eq!(*wr, 50.0);
                assert_eq!(*hr, 50.0);
                assert_eq!(*st_ang, 0.0);
                assert_eq!(
                    *sw_ang, 5_400_000.0,
                    "raw 60000ths preserved (renderer divides by 60000)"
                );
            }
            other => panic!("expected ArcTo, got {other:?}"),
        }

        let value = serde_json::to_value(arc).expect("arcTo serializes");
        let obj = value
            .as_object()
            .expect("arcTo serializes to a JSON object");

        // Tag key is "op" (not "cmd"), variant tag is camelCase "arcTo".
        assert_eq!(
            obj.get("op").and_then(|v| v.as_str()),
            Some("arcTo"),
            "variant tag serializes under key \"op\" as camelCase \"arcTo\""
        );

        // The angle fields must be camelCase — the TS renderer reads
        // cmd.stAng / cmd.swAng. snake_case keys here regress the bug.
        assert!(
            obj.contains_key("stAng"),
            "stAng must be camelCase, got keys: {:?}",
            obj.keys().collect::<Vec<_>>()
        );
        assert!(
            obj.contains_key("swAng"),
            "swAng must be camelCase, got keys: {:?}",
            obj.keys().collect::<Vec<_>>()
        );
        assert!(
            !obj.contains_key("st_ang") && !obj.contains_key("sw_ang"),
            "snake_case angle keys must not appear (regresses the NaN bug)"
        );
        assert_eq!(obj.get("swAng").and_then(|v| v.as_f64()), Some(5_400_000.0));
    }
}

#[cfg(test)]
mod src_rect_tests {
    use super::*;

    const NS: &str = r#"xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;

    /// sample-27: `<a:srcRect l="32560" r="3829"/>` ⇒ left 0.3256, right 0.03829,
    /// top/bottom 0 (absent edges default 0). Edge attrs are ST_Percentage in
    /// 1000ths of a percent, so the fraction is the raw value / 100000.
    #[test]
    fn parses_horizontal_crop_fractions() {
        let xml = format!(
            r#"<xdr:blipFill {NS}>
              <a:blip r:embed="rId1"/>
              <a:srcRect l="32560" r="3829"/>
              <a:stretch><a:fillRect/></a:stretch>
            </xdr:blipFill>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let sr = parse_src_rect(doc.root_element()).expect("srcRect present");
        assert!((sr.l - 0.3256).abs() < 1e-9, "l = l_attr / 100000");
        assert!((sr.r - 0.03829).abs() < 1e-9, "r = r_attr / 100000");
        assert_eq!(sr.t, 0.0, "absent top defaults to 0");
        assert_eq!(sr.b, 0.0, "absent bottom defaults to 0");
    }

    /// No `<a:srcRect>` ⇒ None (uncropped picture, the common case).
    #[test]
    fn absent_src_rect_is_none() {
        let xml = format!(
            r#"<xdr:blipFill {NS}><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        assert!(parse_src_rect(doc.root_element()).is_none());
    }

    /// An all-zero `<a:srcRect>` ⇒ None: an explicit no-op crop must not push the
    /// renderer onto the 9-arg sub-rect path (which would be an identity draw).
    #[test]
    fn all_zero_src_rect_is_none() {
        let xml = format!(
            r#"<xdr:blipFill {NS}><a:blip r:embed="rId1"/><a:srcRect l="0" t="0" r="0" b="0"/></xdr:blipFill>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        assert!(parse_src_rect(doc.root_element()).is_none());
    }
}
