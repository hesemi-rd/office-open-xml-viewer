use crate::types::*;
use crate::{read_zip_entry, read_zip_bytes, resolve_zip_path, parse_rels_map, mime_from_ext};
use base64::{engine::general_purpose::STANDARD as B64, Engine as _};
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
    let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let mut anchors: Vec<ImageAnchor> = Vec::new();

    for anchor in doc.descendants() {
        if anchor.tag_name().name() != "twoCellAnchor"
            || anchor.tag_name().namespace() != Some(xdr_ns)
        {
            continue;
        }
        let (mut from_col, mut from_col_off, mut from_row, mut from_row_off) = (0u32, 0i64, 0u32, 0i64);
        let (mut to_col,   mut to_col_off,   mut to_row,   mut to_row_off)   = (0u32, 0i64, 0u32, 0i64);
        let mut pic_rid: Option<String> = None;
        let mut native_ext_cx: i64 = 0;
        let mut native_ext_cy: i64 = 0;
        // ECMA-376 §20.5.2.33 `twoCellAnchor@editAs`. Possible values:
        // "twoCell" (default), "oneCell", "absolute". With "oneCell" Excel
        // preserves the picture's saved size from <xdr:spPr><a:xfrm><a:ext>
        // regardless of cell resizing.
        let edit_as = anchor.attribute("editAs").map(|s| s.to_string());

        for child in anchor.children() {
            if !child.is_element() { continue; }
            match child.tag_name().name() {
                "from" | "to" => {
                    let is_from = child.tag_name().name() == "from";
                    let mut col: u32 = 0;
                    let mut col_off: i64 = 0;
                    let mut row: u32 = 0;
                    let mut row_off: i64 = 0;
                    for c in child.children() {
                        match (c.tag_name().name(), c.text()) {
                            ("col",    Some(t)) => col     = t.trim().parse().unwrap_or(0),
                            ("colOff", Some(t)) => col_off = t.trim().parse().unwrap_or(0),
                            ("row",    Some(t)) => row     = t.trim().parse().unwrap_or(0),
                            ("rowOff", Some(t)) => row_off = t.trim().parse().unwrap_or(0),
                            _ => {}
                        }
                    }
                    if is_from {
                        from_col = col; from_col_off = col_off; from_row = row; from_row_off = row_off;
                    } else {
                        to_col = col; to_col_off = col_off; to_row = row; to_row_off = row_off;
                    }
                }
                "pic" => {
                    // <xdr:pic><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic>
                    let blip_fill = child.children()
                        .find(|n| n.tag_name().name() == "blipFill" && n.tag_name().namespace() == Some(xdr_ns));
                    if let Some(bf) = blip_fill {
                        let blip = bf.children()
                            .find(|n| n.tag_name().name() == "blip" && n.tag_name().namespace() == Some(a_ns));
                        if let Some(b) = blip {
                            // r:embed attribute
                            pic_rid = b.attributes()
                                .find(|a| a.name() == "embed" && a.namespace() == Some(r_ns))
                                .map(|a| a.value().to_string());
                        }
                    }
                    // <xdr:pic><xdr:spPr><a:xfrm><a:ext cx cy>: the picture's
                    // own saved EMU extent. Authoritative when editAs="oneCell".
                    if let Some(sp_pr) = child.children()
                        .find(|n| n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(xdr_ns))
                    {
                        if let Some(xfrm_n) = sp_pr.children()
                            .find(|n| n.tag_name().name() == "xfrm" && n.tag_name().namespace() == Some(a_ns))
                        {
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

        let Some(rid) = pic_rid else { continue; };
        let Some(target) = drawing_rels.get(&rid) else { continue; };
        let media_path = resolve_zip_path(drawing_dir, target);
        let Some(bytes) = read_zip_bytes(archive, &media_path) else { continue; };
        let mime = mime_from_ext(&media_path);
        let data_url = format!("data:{mime};base64,{}", B64.encode(&bytes));

        anchors.push(ImageAnchor {
            from_col, from_col_off, from_row, from_row_off,
            to_col, to_col_off, to_row, to_row_off,
            edit_as,
            native_ext_cx,
            native_ext_cy,
            data_url,
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
    off_x: f64, off_y: f64,
    ext_x: f64, ext_y: f64,
    ch_off_x: f64, ch_off_y: f64,
    ch_ext_x: f64, ch_ext_y: f64,
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
                ext.0 = c.attribute("cx").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                ext.1 = c.attribute("cy").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                has_ext = true;
            }
            "chOff" => {
                ch_off.0 = c.attribute("x").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                ch_off.1 = c.attribute("y").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                has_ch = true;
            }
            "chExt" => {
                ch_ext.0 = c.attribute("cx").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                ch_ext.1 = c.attribute("cy").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                has_ch = true;
            }
            _ => {}
        }
    }
    if !has_ext { return None; }
    Some(Xfrm {
        off_x: off.0, off_y: off.1,
        ext_x: ext.0, ext_y: ext.1,
        ch_off_x: ch_off.0, ch_off_y: ch_off.1,
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

pub(crate) fn parse_solid_fill(fill_node: &roxmltree::Node, theme_colors: &[String]) -> Option<String> {
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
                    "dk1" | "tx1"    => Some(0),
                    "lt1" | "bg1"    => Some(1),
                    "dk2" | "tx2"    => Some(2),
                    "lt2" | "bg2"    => Some(3),
                    "accent1"        => Some(4),
                    "accent2"        => Some(5),
                    "accent3"        => Some(6),
                    "accent4"        => Some(7),
                    "accent5"        => Some(8),
                    "accent6"        => Some(9),
                    "hlink"          => Some(10),
                    "folHlink"       => Some(11),
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
    let w: f64 = path_node.attribute("w").and_then(|s| s.parse().ok()).unwrap_or(0.0);
    let h: f64 = path_node.attribute("h").and_then(|s| s.parse().ok()).unwrap_or(0.0);
    let mut commands: Vec<PathCmd> = Vec::new();
    for cmd in path_node.children().filter(|n| n.is_element()) {
        let name = cmd.tag_name().name();
        // Collect `<a:pt x=.. y=..>` points in order.
        let pts: Vec<(f64, f64)> = cmd.children()
            .filter(|n| n.is_element() && n.tag_name().name() == "pt")
            .map(|n| (
                n.attribute("x").and_then(|s| s.parse().ok()).unwrap_or(0.0),
                n.attribute("y").and_then(|s| s.parse().ok()).unwrap_or(0.0),
            ))
            .collect();
        match name {
            "moveTo"       => if let Some(p) = pts.first() { commands.push(PathCmd::MoveTo { x: p.0, y: p.1 }); },
            "lnTo"         => if let Some(p) = pts.first() { commands.push(PathCmd::LineTo { x: p.0, y: p.1 }); },
            "cubicBezTo"   => if pts.len() >= 3 {
                commands.push(PathCmd::CubicBezTo {
                    x1: pts[0].0, y1: pts[0].1,
                    x2: pts[1].0, y2: pts[1].1,
                    x3: pts[2].0, y3: pts[2].1,
                });
            },
            "quadBezTo"    => if pts.len() >= 2 {
                commands.push(PathCmd::QuadBezTo {
                    x1: pts[0].0, y1: pts[0].1,
                    x2: pts[1].0, y2: pts[1].1,
                });
            },
            "close"        => commands.push(PathCmd::Close),
            "arcTo" => {
                // ECMA-376 §20.1.9.3: `wR`/`hR` in path-coord units;
                // `stAng`/`swAng` in 60000ths of a degree.
                let wr:     f64 = cmd.attribute("wR").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                let hr:     f64 = cmd.attribute("hR").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                let st_ang: f64 = cmd.attribute("stAng").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                let sw_ang: f64 = cmd.attribute("swAng").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                commands.push(PathCmd::ArcTo { wr, hr, st_ang, sw_ang });
            }
            _ => {}
        }
    }
    PathInfo { w, h, commands }
}

/// Parse `<xdr:txBody>` into a `ShapeText`. Returns `None` if the body
/// contains no visible runs. Run formatting follows ECMA-376 §21.1.2.3.1
/// (`<a:rPr>`): `sz` is hundredths of a point, `b="1"` = bold, `i="1"`
/// = italic, `<a:solidFill>` overrides shape-level font color, and
/// `<a:latin@typeface>` selects the Latin font face (we don't yet
/// distinguish East-Asian / complex-script fonts — `<a:ea>` and `<a:cs>`
/// are ignored for typeface).
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

pub(crate) fn parse_tx_body(tx_body: &roxmltree::Node, theme_colors: &[String]) -> Option<ShapeText> {
    let mut anchor = String::from("t");
    let mut wrap = String::from("square");
    let mut paragraphs: Vec<ShapeParagraph> = Vec::new();
    for c in tx_body.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "bodyPr" => {
                if let Some(a) = c.attribute("anchor") { anchor = a.to_string(); }
                if let Some(w) = c.attribute("wrap") { wrap = w.to_string(); }
            }
            "p" => {
                let mut align = String::from("l");
                let mut rtl = false;
                let mut runs: Vec<ShapeTextRun> = Vec::new();
                for pc in c.children().filter(|n| n.is_element()) {
                    match pc.tag_name().name() {
                        "pPr" => {
                            if let Some(a) = pc.attribute("algn") { align = a.to_string(); }
                            // ECMA-376 §21.1.2.2.7 `<a:pPr@rtl>` — right-to-left
                            // paragraph. Default false (left-to-right).
                            rtl = pc.attribute("rtl").map(|v| v == "1" || v == "true").unwrap_or(false);
                        }
                        "r" => {
                            // Run text + run-level formatting.
                            let mut text = String::new();
                            let mut bold = false;
                            let mut italic = false;
                            let mut size: f64 = 0.0;
                            let mut color: Option<String> = None;
                            let mut font_face: Option<String> = None;
                            for rc in pc.children().filter(|n| n.is_element()) {
                                match rc.tag_name().name() {
                                    "rPr" => {
                                        bold = rc.attribute("b").map(|v| v == "1").unwrap_or(false);
                                        italic = rc.attribute("i").map(|v| v == "1").unwrap_or(false);
                                        size = rc.attribute("sz")
                                            .and_then(|s| s.parse::<f64>().ok())
                                            .map(|v| v / 100.0)
                                            .unwrap_or(0.0);
                                        for rpc in rc.children().filter(|n| n.is_element()) {
                                            match rpc.tag_name().name() {
                                                "solidFill" => {
                                                    color = parse_solid_fill(&rpc, theme_colors);
                                                }
                                                "latin" => {
                                                    font_face = rpc.attribute("typeface").map(String::from);
                                                }
                                                _ => {}
                                            }
                                        }
                                    }
                                    "t" => {
                                        if let Some(t) = rc.text() { text.push_str(t); }
                                    }
                                    _ => {}
                                }
                            }
                            if !text.is_empty() {
                                runs.push(ShapeTextRun::Text { text, bold, italic, size, color, font_face });
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
                    paragraphs.push(ShapeParagraph { align, rtl, runs });
                }
            }
            _ => {}
        }
    }
    if paragraphs.is_empty() { None } else {
        Some(ShapeText { anchor, wrap, paragraphs })
    }
}

pub(crate) fn parse_sp_geom(sp_pr: &roxmltree::Node) -> Option<ShapeGeom> {
    for c in sp_pr.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "prstGeom" => {
                return Some(ShapeGeom::Preset {
                    name: c.attribute("prst").unwrap_or("rect").to_string(),
                });
            }
            "custGeom" => {
                let mut paths: Vec<PathInfo> = Vec::new();
                for pl in c.children().filter(|n| n.is_element() && n.tag_name().name() == "pathLst") {
                    for p in pl.children().filter(|n| n.is_element() && n.tag_name().name() == "path") {
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
pub(crate) fn collect_shapes(
    node: &roxmltree::Node,
    root_off_x: f64, root_off_y: f64,
    root_ext_x: f64, root_ext_y: f64,
    // transform from current local coords into root (top-level grpSp) coords
    scale_x: f64, scale_y: f64,
    trans_x: f64, trans_y: f64,
    theme_colors: &[String],
    rid_urls: &HashMap<String, String>,
    out: &mut Vec<ShapeInfo>,
) {
    for child in node.children().filter(|n| n.is_element()) {
        let tag = child.tag_name().name();
        if tag == "grpSp" {
            // Nested grpSp: compose the transform by the group's own xfrm.
            let grp_sp_pr = child.children().find(|n| n.is_element() && n.tag_name().name() == "grpSpPr");
            let xfrm = grp_sp_pr
                .and_then(|n| n.children().find(|c| c.is_element() && c.tag_name().name() == "xfrm"))
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
                    (scale_x, scale_y,
                     trans_x + scale_x * x.off_x,
                     trans_y + scale_y * x.off_y)
                }
            } else {
                (scale_x, scale_y, trans_x, trans_y)
            };
            collect_shapes(&child, root_off_x, root_off_y, root_ext_x, root_ext_y,
                           sx, sy, tx, ty, theme_colors, rid_urls, out);
        } else if tag == "sp" {
            let sp_pr = child.children().find(|n| n.is_element() && n.tag_name().name() == "spPr");
            let Some(sp_pr) = sp_pr else { continue; };
            let xfrm_node = sp_pr.children().find(|n| n.is_element() && n.tag_name().name() == "xfrm");
            let Some(xfrm_n) = xfrm_node else { continue; };
            let Some(xfrm) = parse_xfrm(&xfrm_n) else { continue; };
            let rot_raw: f64 = xfrm_n.attribute("rot")
                .and_then(|s| s.parse().ok()).unwrap_or(0.0);

            // Shape rect in root coords
            let root_x = trans_x + scale_x * xfrm.off_x;
            let root_y = trans_y + scale_y * xfrm.off_y;
            let root_w = scale_x * xfrm.ext_x;
            let root_h = scale_y * xfrm.ext_y;

            // Normalize to [0,1] of root ext
            if root_ext_x == 0.0 || root_ext_y == 0.0 { continue; }
            let nx = (root_x - root_off_x) / root_ext_x;
            let ny = (root_y - root_off_y) / root_ext_y;
            let nw = root_w / root_ext_x;
            let nh = root_h / root_ext_y;

            let geom = parse_sp_geom(&sp_pr);
            let Some(geom) = geom else { continue; };

            // Fill
            let mut fill_color: Option<String> = None;
            let mut has_no_fill = false;
            for c in sp_pr.children().filter(|n| n.is_element()) {
                match c.tag_name().name() {
                    "solidFill" => { fill_color = parse_solid_fill(&c, theme_colors); }
                    "noFill"    => { has_no_fill = true; }
                    _ => {}
                }
            }
            if has_no_fill { fill_color = None; }

            // Stroke (line)
            let mut stroke_color: Option<String> = None;
            let mut stroke_width: i64 = 0;
            if let Some(ln) = sp_pr.children().find(|n| n.is_element() && n.tag_name().name() == "ln") {
                let w_attr = ln.attribute("w");
                stroke_width = w_attr.and_then(|s| s.parse().ok()).unwrap_or(0);
                for c in ln.children().filter(|n| n.is_element()) {
                    if c.tag_name().name() == "solidFill" {
                        stroke_color = parse_solid_fill(&c, theme_colors);
                    } else if c.tag_name().name() == "noFill" {
                        stroke_color = None;
                        stroke_width = 0;
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
            let style_node = child.children()
                .find(|n| n.is_element() && n.tag_name().name() == "style");
            let style_fill = style_node.as_ref()
                .and_then(|s| s.children().find(|n| n.is_element() && n.tag_name().name() == "fillRef"))
                .and_then(|n| parse_solid_fill(&n, theme_colors));
            let style_text_color = style_node.as_ref()
                .and_then(|s| s.children().find(|n| n.is_element() && n.tag_name().name() == "fontRef"))
                .and_then(|n| parse_solid_fill(&n, theme_colors));
            if fill_color.is_none() && !has_no_fill {
                fill_color = style_fill;
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
                x: nx, y: ny, w: nw, h: nh,
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
            // each r:id to its pre-encoded `data:<mime>;base64,…` URL.
            let sp_pr = child.children().find(|n| n.is_element() && n.tag_name().name() == "spPr");
            let Some(sp_pr) = sp_pr else { continue; };
            let xfrm_node = sp_pr.children().find(|n| n.is_element() && n.tag_name().name() == "xfrm");
            let Some(xfrm_n) = xfrm_node else { continue; };
            let Some(xfrm) = parse_xfrm(&xfrm_n) else { continue; };
            let rot_raw: f64 = xfrm_n.attribute("rot")
                .and_then(|s| s.parse().ok()).unwrap_or(0.0);

            // Resolve <a:blip r:embed="rIdN"/>. The r:embed attribute lives in
            // the relationships namespace, not the drawingml namespace.
            let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            let pic_rid = child.descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "blip")
                .and_then(|b| {
                    b.attributes()
                        .find(|a| a.name() == "embed" && a.namespace() == Some(r_ns))
                        .map(|a| a.value().to_string())
                });
            let Some(rid) = pic_rid else { continue; };
            let Some(data_url) = rid_urls.get(&rid) else { continue; };

            let root_x = trans_x + scale_x * xfrm.off_x;
            let root_y = trans_y + scale_y * xfrm.off_y;
            let root_w = scale_x * xfrm.ext_x;
            let root_h = scale_y * xfrm.ext_y;
            if root_ext_x == 0.0 || root_ext_y == 0.0 { continue; }
            let nx = (root_x - root_off_x) / root_ext_x;
            let ny = (root_y - root_off_y) / root_ext_y;
            let nw = root_w / root_ext_x;
            let nh = root_h / root_ext_y;
            if nw <= 0.0 || nh <= 0.0 { continue; }

            out.push(ShapeInfo {
                x: nx, y: ny, w: nw, h: nh,
                rot: rot_raw / 60000.0,
                fill_color: None,
                stroke_color: None,
                stroke_width: 0,
                geom: ShapeGeom::Image { data_url: data_url.clone() },
                text: None,
            });
        }
        // Ignore `xdr:cxnSp` / text-only elements for this minimal pass.
    }
}

pub(crate) fn parse_shape_anchors(
    drawing_xml: &str,
    theme_colors: &[String],
    rid_urls: &HashMap<String, String>,
) -> Vec<ShapeAnchor> {
    let Ok(doc) = roxmltree::Document::parse(drawing_xml) else { return Vec::new(); };
    let xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    let mut anchors: Vec<ShapeAnchor> = Vec::new();

    for anchor in doc.descendants() {
        let anchor_tag = anchor.tag_name().name();
        // ECMA-376 §20.5.2: a drawing shape is anchored with either
        // `twoCellAnchor` (from+to) or `oneCellAnchor` (from + a saved `<ext>`
        // size). Excel authors equation text boxes as oneCellAnchor, so we must
        // accept both or those shapes (and their math) are silently dropped.
        if (anchor_tag != "twoCellAnchor" && anchor_tag != "oneCellAnchor")
            || anchor.tag_name().namespace() != Some(xdr_ns) { continue; }

        // Parse from/to anchor rect (shared between grpSp and stand-alone sp paths)
        let (mut from_col, mut from_col_off, mut from_row, mut from_row_off) = (0u32, 0i64, 0u32, 0i64);
        let (mut to_col,   mut to_col_off,   mut to_row,   mut to_row_off)   = (0u32, 0i64, 0u32, 0i64);
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
            if !c.is_element() { continue; }
            if c.tag_name().name() == "from" || c.tag_name().name() == "to" {
                let is_from = c.tag_name().name() == "from";
                let mut col: u32 = 0; let mut col_off: i64 = 0;
                let mut row: u32 = 0; let mut row_off: i64 = 0;
                for cc in c.children() {
                    match (cc.tag_name().name(), cc.text()) {
                        ("col",    Some(t)) => col     = t.trim().parse().unwrap_or(0),
                        ("colOff", Some(t)) => col_off = t.trim().parse().unwrap_or(0),
                        ("row",    Some(t)) => row     = t.trim().parse().unwrap_or(0),
                        ("rowOff", Some(t)) => row_off = t.trim().parse().unwrap_or(0),
                        _ => {}
                    }
                }
                if is_from {
                    from_col = col; from_col_off = col_off; from_row = row; from_row_off = row_off;
                } else {
                    to_col = col; to_col_off = col_off; to_row = row; to_row_off = row_off;
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
        if let Some(grp) = content.children().find(|n| n.is_element() && n.tag_name().name() == "grpSp") {
            let grp_sp_pr = grp.children().find(|n| n.is_element() && n.tag_name().name() == "grpSpPr");
            let xfrm = grp_sp_pr
                .and_then(|n| n.children().find(|c| c.is_element() && c.tag_name().name() == "xfrm"))
                .as_ref()
                .and_then(parse_xfrm);
            let Some(root) = xfrm else { continue; };
            if !root.has_ch || root.ch_ext_x == 0.0 || root.ch_ext_y == 0.0 { continue; }

            // Top-level grpSp ext is the group's saved on-sheet EMU size —
            // authoritative when editAs="oneCell".
            native_ext_cx = root.ext_x as i64;
            native_ext_cy = root.ext_y as i64;

            // Map child coords → root coords with the grpSp's own chOff/chExt.
            let csx = root.ext_x / root.ch_ext_x;
            let csy = root.ext_y / root.ch_ext_y;
            let tx = root.off_x - root.ch_off_x * csx;
            let ty = root.off_y - root.ch_off_y * csy;

            collect_shapes(&grp, root.off_x, root.off_y, root.ext_x, root.ext_y,
                           csx, csy, tx, ty, theme_colors, rid_urls, &mut shapes);
        } else if let Some(sp) = content.children().find(|n| n.is_element() && (n.tag_name().name() == "sp" || n.tag_name().name() == "pic")) {
            // Stand-alone sp/pic: the shape's own xfrm gives its absolute EMU
            // rect, but for our rendering pipeline the anchor's from/to
            // already defines the on-sheet rect, and the leaf occupies it
            // 100 %. Build a synthetic root coord-system whose origin matches
            // the shape's xfrm so collect_shapes normalizes the leaf to (0,0)
            // (1,1).
            let sp_pr = sp.children().find(|n| n.is_element() && n.tag_name().name() == "spPr");
            let Some(sp_pr_node) = sp_pr else { continue; };
            let xfrm_node = sp_pr_node.children().find(|n| n.is_element() && n.tag_name().name() == "xfrm");
            let Some(xfrm_n) = xfrm_node else { continue; };
            let Some(xfrm) = parse_xfrm(&xfrm_n) else { continue; };
            if xfrm.ext_x == 0.0 || xfrm.ext_y == 0.0 { continue; }
            native_ext_cx = xfrm.ext_x as i64;
            native_ext_cy = xfrm.ext_y as i64;
            collect_shapes(&content, xfrm.off_x, xfrm.off_y, xfrm.ext_x, xfrm.ext_y,
                           1.0, 1.0, 0.0, 0.0, theme_colors, rid_urls, &mut shapes);
        } else {
            continue;
        }

        if shapes.is_empty() { continue; }

        anchors.push(ShapeAnchor {
            from_col, from_col_off, from_row, from_row_off,
            to_col, to_col_off, to_row, to_row_off,
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
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else { return Vec::new(); };
    let sheet_rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let Ok(sheet_rels_xml) = read_zip_entry(archive, &sheet_rels_path) else { return Vec::new(); };
    let Ok(rels_doc) = roxmltree::Document::parse(&sheet_rels_xml) else { return Vec::new(); };
    let mut drawing_targets: Vec<String> = Vec::new();
    for rel in rels_doc.root_element().children().filter(|n| n.is_element()) {
        if rel.attribute("Type").unwrap_or("").ends_with("/drawing") {
            if let Some(t) = rel.attribute("Target") { drawing_targets.push(t.to_string()); }
        }
    }
    let mut all: Vec<ShapeAnchor> = Vec::new();
    for target in drawing_targets {
        let drawing_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        let Ok(drawing_xml) = read_zip_entry(archive, &drawing_path) else { continue; };
        let rid_urls = build_drawing_rid_urls(archive, &drawing_path);
        all.extend(parse_shape_anchors(&drawing_xml, theme_colors, &rid_urls));
    }
    all
}

/// Build a `HashMap<rId, data-URL>` for every image (png/jpg/…) target in
/// a drawing's `.rels` file. Used by `collect_shapes` to resolve `<xdr:pic>`
/// leaves inside a group. Mirrors the logic in `parse_drawing_anchors` but
/// eagerly encodes each referenced image so per-shape lookup is a single
/// HashMap hit.
pub(crate) fn build_drawing_rid_urls(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    drawing_path: &str,
) -> HashMap<String, String> {
    let Some((drawing_dir, drawing_file)) = drawing_path.rsplit_once('/') else {
        return HashMap::new();
    };
    let rels_path = format!("{}/_rels/{}.rels", drawing_dir, drawing_file);
    let rels = read_zip_entry(archive, &rels_path)
        .ok()
        .map(|xml| parse_rels_map(&xml))
        .unwrap_or_default();

    let mut result: HashMap<String, String> = HashMap::new();
    for (rid, target) in rels {
        let lower = target.to_lowercase();
        if !(lower.ends_with(".png") || lower.ends_with(".jpg")
            || lower.ends_with(".jpeg") || lower.ends_with(".gif")
            || lower.ends_with(".bmp")  || lower.ends_with(".webp"))
        {
            continue;
        }
        let media_path = resolve_zip_path(drawing_dir, &target);
        if let Some(bytes) = read_zip_bytes(archive, &media_path) {
            let mime = mime_from_ext(&media_path);
            result.insert(rid, format!("data:{mime};base64,{}", B64.encode(&bytes)));
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
    let Ok(sheet_rels_xml) = read_zip_entry(archive, &sheet_rels_path) else {
        return Vec::new();
    };

    // Find all drawing relationships
    let Ok(rels_doc) = roxmltree::Document::parse(&sheet_rels_xml) else {
        return Vec::new();
    };
    let mut drawing_targets: Vec<String> = Vec::new();
    for rel in rels_doc.root_element().children().filter(|n| n.is_element()) {
        let rel_type = rel.attribute("Type").unwrap_or("");
        if rel_type.ends_with("/drawing") {
            if let Some(t) = rel.attribute("Target") {
                drawing_targets.push(t.to_string());
            }
        }
    }
    if drawing_targets.is_empty() { return Vec::new(); }

    let mut all_anchors: Vec<ImageAnchor> = Vec::new();
    for target in drawing_targets {
        // sheet_dir is "worksheets", target typically "../drawings/drawing1.xml"
        // base dir for the drawing = "xl/worksheets" + "../drawings" → "xl/drawings"
        let drawing_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        let Ok(drawing_xml) = read_zip_entry(archive, &drawing_path) else { continue; };
        // Drawing rels:  xl/drawings/_rels/drawing1.xml.rels
        let Some((drawing_dir, drawing_file)) = drawing_path.rsplit_once('/') else { continue; };
        let drawing_rels_path = format!("{}/_rels/{}.rels", drawing_dir, drawing_file);
        let drawing_rels = read_zip_entry(archive, &drawing_rels_path)
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
            ShapeTextRun::Math { display, font_size, color, nodes } => {
                assert!(!*display, "bare a14:m + m:oMath → inline");
                assert_eq!(*font_size, Some(28.0), "size from math run rPr sz/100");
                assert_eq!(color.as_deref(), Some("#7030A0"), "colour from math run rPr solidFill (xlsx #-prefixed convention)");
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
}

