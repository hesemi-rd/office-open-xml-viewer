//! Shape / tree / picture / media / table / connector / SmartArt assembly:
//! turns a `<p:spTree>` and its children into `SlideElement`s. Extracted
//! verbatim from `lib.rs`. Pulls fill/stroke/effect/geometry parsing from
//! `fill`, text-body parsing from `text`, chart parsing from `chart`, and the
//! shared XML/zip helpers + `TableStyleDef` from the crate root.

use crate::chart::{parse_chartex, parse_legacy_chart};
use crate::fill::{
    parse_arrow_end, parse_blip_alpha, parse_color_node, parse_cust_geom, parse_effect_lst,
    parse_fill, parse_scene3d, parse_shadow, parse_sp3d, parse_stroke, parse_table_style_fill,
    parse_xfrm, EffectLst,
};
use crate::master::LayoutPlaceholders;
use crate::text::{
    empty_level_bullets, parse_text_body, LevelBullets, LevelFontSizes, LevelIndents, ShapeKind,
};
use crate::types::*;
use crate::{
    attr, attr_f64, attr_i64, attr_r, child, read_zip_str, resolve_path, table_style_presets,
    PptxZip, TableStyleDef,
};
use ooxml_common::blip::{mime_from_ext, parse_src_rect, svg_blip_rid};
use std::collections::HashMap;

/// OOXML spec default positions for common placeholder types.
/// Values are in EMU, assuming a 9144000×6858000 slide (10"×7.5").
pub(crate) fn default_placeholder_transform(ph_type: &str) -> Transform {
    match ph_type {
        "title" | "ctrTitle" => Transform {
            x: 457200,
            y: 274638,
            cx: 8229600,
            cy: 1143000,
            rot: 0.0,
            flip_h: false,
            flip_v: false,
        },
        "subTitle" => Transform {
            x: 457200,
            y: 1600200,
            cx: 8229600,
            cy: 899160,
            rot: 0.0,
            flip_h: false,
            flip_v: false,
        },
        "dt" => Transform {
            x: 0,
            y: 6261600,
            cx: 2286000,
            cy: 596900,
            rot: 0.0,
            flip_h: false,
            flip_v: false,
        },
        "ftr" => Transform {
            x: 2972400,
            y: 6261600,
            cx: 3086100,
            cy: 596900,
            rot: 0.0,
            flip_h: false,
            flip_v: false,
        },
        "sldNum" => Transform {
            x: 6629400,
            y: 6261600,
            cx: 2057400,
            cy: 596900,
            rot: 0.0,
            flip_h: false,
            flip_v: false,
        },
        // "body" and everything else: full-width content area below title
        _ => Transform {
            x: 457200,
            y: 1600200,
            cx: 8229600,
            cy: 4525963,
            rot: 0.0,
            flip_h: false,
            flip_v: false,
        },
    }
}

/// Returns true if the node contains a `p:ph` descendant.
pub(crate) fn is_placeholder(node: roxmltree::Node<'_, '_>) -> bool {
    node.descendants()
        .any(|n| n.is_element() && n.tag_name().name() == "ph")
}

/// Pull `(id, name)` from a sibling `<p:nvSpPr><p:cNvPr>` (or `<p:nvCxnSpPr>`).
/// Returns `(None, None)` when the wrapper is missing — both fields are
/// optional in the JSON output.
pub(crate) fn read_cnv_pr(sp_node: roxmltree::Node<'_, '_>) -> (Option<String>, Option<String>) {
    for wrapper_name in &["nvSpPr", "nvCxnSpPr", "nvPicPr", "nvGraphicFramePr"] {
        if let Some(nv) = child(sp_node, wrapper_name) {
            if let Some(cnv) = child(nv, "cNvPr") {
                let id = attr(&cnv, "id");
                let name = attr(&cnv, "name").filter(|s| !s.is_empty());
                return (id, name);
            }
        }
    }
    (None, None)
}

pub(crate) fn parse_shape(
    sp_node: roxmltree::Node<'_, '_>,
    lph: &LayoutPlaceholders,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
    group_fill: Option<&Fill>,
    zip: &mut PptxZip,
) -> Option<ShapeElement> {
    // --- Placeholder info (for layout fallback) ---
    let ph_node = sp_node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "ph");
    let ph_type = ph_node
        .as_ref()
        .and_then(|n| attr(n, "type"))
        .unwrap_or_else(|| "body".into());
    let ph_idx: Option<u32> = ph_node
        .as_ref()
        .and_then(|n| attr(n, "idx"))
        .and_then(|v| v.parse().ok());
    // Surface the explicit ph @type / @idx (when present) on the ShapeElement
    // JSON. We keep `ph_type` defaulted to "body" for the internal lookup
    // path, but only emit the `placeholder_type` field when a real `<p:ph>`
    // node was found.
    let placeholder_type_out: Option<String> = ph_node
        .as_ref()
        .map(|n| attr(n, "type").unwrap_or_else(|| "body".into()));
    let (cnv_id, cnv_name) = read_cnv_pr(sp_node);

    // --- Transform: slide xfrm OR layout fallback ---
    let sp_pr = child(sp_node, "spPr");
    let slide_xfrm = sp_pr.and_then(|p| child(p, "xfrm"));

    let t: Transform = if let Some(xfrm) = slide_xfrm {
        parse_xfrm(xfrm)
    } else if ph_node.is_some() {
        match lph.lookup(&ph_type, ph_idx) {
            Some(lt) => lt.clone(),
            None => default_placeholder_transform(&ph_type),
        }
    } else {
        return None; // non-placeholder with no xfrm — skip
    };

    // cx=0 AND cy=0 → skip (zero-area invisible shape).
    // cx=0 alone (vertical line/annotation) is permitted when cy > 0.
    // cy=0 means "auto-height": keep 0 when anchor="b" (renderer grows shape upward from off_y),
    // otherwise use a generous fallback so text has room to render.
    if t.cx == 0 && t.cy == 0 {
        return None;
    }
    let inherited_anchor: Option<String> = if ph_node.is_some() {
        lph.lookup_anchor(&ph_type)
    } else {
        None
    };
    let is_bottom_anchor = inherited_anchor.as_deref() == Some("b")
        || child(sp_node, "txBody")
            .and_then(|tb| child(tb, "bodyPr"))
            .and_then(|bp| attr(&bp, "anchor"))
            .map(|a| a == "b")
            .unwrap_or(false);
    // custGeom takes priority over prstGeom
    let cust_geom_node = sp_pr.and_then(|p| child(p, "custGeom"));
    let prst_geom_node = sp_pr.and_then(|p| child(p, "prstGeom"));
    let geometry = if cust_geom_node.is_some() {
        "custGeom".into()
    } else {
        prst_geom_node
            .and_then(|n| attr(&n, "prst"))
            .unwrap_or_else(|| "rect".into())
    };

    // cy=0 means "auto-height" for body-text shapes, but connector-type
    // geometries (line, *Connector*) legitimately use cy=0 to represent
    // a perfectly horizontal segment — don't inflate their height.
    let is_connector_geom = matches!(
        geometry.as_str(),
        "line"
            | "straightConnector1"
            | "bentConnector2"
            | "bentConnector3"
            | "bentConnector4"
            | "bentConnector5"
            | "curvedConnector2"
            | "curvedConnector3"
            | "curvedConnector4"
            | "curvedConnector5"
    );
    let cy = if t.cy == 0 && !is_connector_geom {
        if is_bottom_anchor {
            0_i64
        } else {
            2_000_000_i64
        }
    } else {
        t.cy
    };
    let cust_geom = cust_geom_node.map(|n| parse_cust_geom(n));

    // Parse adjustment values from prstGeom avLst (e.g. trapezoid inset)
    // Collect all gd elements; first is adj (name="adj" or "adj1"), second is adj2
    let parse_gd_val = |gd: roxmltree::Node<'_, '_>| -> Option<f64> {
        attr(&gd, "fmla")
            .and_then(|f| f.strip_prefix("val ").map(|s| s.to_owned()))
            .and_then(|s| s.parse::<f64>().ok())
    };
    let av_node = prst_geom_node.and_then(|n| child(n, "avLst"));
    let gd_nodes: Vec<_> = av_node
        .map(|av| {
            av.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "gd")
                .collect()
        })
        .unwrap_or_default();
    // First gd = adj (match by name "adj" or "adj1", fallback to position 0)
    let adj: Option<f64> = gd_nodes
        .iter()
        .find(|n| matches!(attr(n, "name").as_deref(), Some("adj") | Some("adj1")))
        .or_else(|| gd_nodes.first())
        .and_then(|n| parse_gd_val(*n));
    // Second gd = adj2 (match by name "adj2", fallback to position 1)
    let adj2: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj2"))
        .or_else(|| gd_nodes.get(1))
        .and_then(|n| parse_gd_val(*n));
    // Third gd = adj3 (match by name "adj3", fallback to position 2)
    let adj3: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj3"))
        .or_else(|| gd_nodes.get(2))
        .and_then(|n| parse_gd_val(*n));
    // Fourth gd = adj4 (match by name "adj4", fallback to position 3)
    let adj4: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj4"))
        .or_else(|| gd_nodes.get(3))
        .and_then(|n| parse_gd_val(*n));
    // adj5-adj8 for callouts that specify extra polyline vertices
    // (accentBorderCallout3 etc.).
    let adj5: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj5"))
        .or_else(|| gd_nodes.get(4))
        .and_then(|n| parse_gd_val(*n));
    let adj6: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj6"))
        .or_else(|| gd_nodes.get(5))
        .and_then(|n| parse_gd_val(*n));
    let adj7: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj7"))
        .or_else(|| gd_nodes.get(6))
        .and_then(|n| parse_gd_val(*n));
    let adj8: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj8"))
        .or_else(|| gd_nodes.get(7))
        .and_then(|n| parse_gd_val(*n));

    // --- Shape style (p:style) provides fill/stroke/text-color fallbacks ---
    let style_node = child(sp_node, "style");

    // fillRef idx=0 → explicit no-fill; idx>0 → use referenced color as solid fill
    let style_fill: Option<Fill> = style_node.and_then(|s| child(s, "fillRef")).and_then(|fr| {
        let idx: u32 = attr(&fr, "idx").and_then(|v| v.parse().ok()).unwrap_or(1);
        if idx == 0 {
            Some(Fill::None)
        } else {
            parse_color_node(fr, theme).map(|c| Fill::Solid { color: c })
        }
    });

    // lnRef idx=0 → no line; idx>0 → resolve width from theme's fmtScheme >
    // lnStyleLst (stored as "+lnRef-N") and color from the ref's own solidFill.
    // Falling back to 9525 under-weights idx>=2 strokes (Office default theme
    // idx=2 is 19050 EMU = 1.5pt, idx=3 is 25400 EMU = 2pt).
    let style_stroke: Option<Stroke> = style_node.and_then(|s| child(s, "lnRef")).and_then(|lr| {
        let idx: u32 = attr(&lr, "idx").and_then(|v| v.parse().ok()).unwrap_or(1);
        if idx == 0 {
            None
        } else {
            parse_color_node(lr, theme).map(|c| {
                let width = theme
                    .get(&format!("+lnRef-{}", idx))
                    .and_then(|s| s.parse::<i64>().ok())
                    .unwrap_or(9525);
                Stroke {
                    color: c,
                    width,
                    dash_style: None,
                    head_end: None,
                    tail_end: None,
                    cmpd: None,
                }
            })
        }
    });

    // fontRef → default text color for this shape.
    // Fall back to layout/master placeholder inherited color (lstStyle > lvl1pPr > defRPr
    // > solidFill) when the shape is a placeholder and has no explicit p:style > fontRef.
    let default_text_color: Option<String> = style_node
        .and_then(|s| child(s, "fontRef"))
        .and_then(|fr| parse_color_node(fr, theme))
        .or_else(|| {
            if ph_node.is_some() {
                lph.lookup_color(&ph_type, ph_idx)
            } else {
                None
            }
        });

    // spPr fill resolution order (ECMA-376 §19.3.1.36 / §20.1.4.2):
    //   1. spPr `<a:grpFill>` → inherit from parent group
    //   2. spPr explicit fill (`<a:solidFill>`, `<a:noFill>`, gradient, pattern, blip)
    //   3. layout placeholder fill (only when the slide-level shape has no fill
    //      element of its own and is bound to a placeholder)
    //   4. `<p:style><a:fillRef>` falls through last
    // Some(Fill::None) (noFill in spPr) is treated as an explicit choice and
    // must NOT be overridden by step 3 or 4.
    let sp_pr_has_grp_fill = sp_pr.and_then(|p| child(p, "grpFill")).is_some();
    let fill = if sp_pr_has_grp_fill {
        group_fill.cloned()
    } else {
        let own = sp_pr.and_then(|p| parse_fill(p, theme));
        let inherited = if own.is_none() && ph_node.is_some() {
            lph.lookup_fill(&ph_type, ph_idx)
        } else {
            None
        };
        own.or(inherited).or(style_fill)
    };

    // spPr stroke: if ln element is present, respect it (even if noFill → None);
    // otherwise fall back to layout placeholder stroke, then style stroke.
    let stroke = if sp_pr.and_then(|p| child(p, "ln")).is_some() {
        sp_pr
            .and_then(|p| child(p, "ln"))
            .and_then(|n| parse_stroke(n, theme))
    } else if ph_node.is_some() {
        lph.lookup_stroke(&ph_type, ph_idx).or(style_stroke)
    } else {
        style_stroke
    };

    // Inherited defaults from layout/master for this placeholder type/idx
    let (
        inherited_font_size,
        inherited_bold,
        inherited_italic,
        inherited_caps,
        inherited_anchor,
        inherited_alignment,
        inherited_ea_ln_brk,
        inherited_space_before,
        inherited_space_after,
        inherited_line_spacing,
    ) = if ph_node.is_some() {
        (
            lph.lookup_font_size(&ph_type, ph_idx),
            lph.lookup_bold(&ph_type),
            lph.lookup_italic(&ph_type),
            lph.lookup_caps(&ph_type),
            lph.lookup_anchor(&ph_type),
            lph.lookup_alignment(&ph_type, ph_idx),
            lph.lookup_ea_ln_brk(&ph_type),
            lph.lookup_space_before(&ph_type),
            lph.lookup_space_after(&ph_type),
            lph.lookup_line_spacing(&ph_type, ph_idx),
        )
    } else {
        (None, None, None, None, None, None, None, None, None, None)
    };
    let inherited_level_font_sizes: LevelFontSizes = if ph_node.is_some() {
        lph.lookup_level_font_sizes(&ph_type, ph_idx)
    } else {
        [None; 9]
    };
    // Per-level paragraph indents (marL/marR/indent) a paragraph inherits when it
    // omits them (ECMA-376 §21.1.2.4.13): the layout/master placeholder cascade.
    let inherited_level_indents: LevelIndents = if ph_node.is_some() {
        lph.lookup_level_indents(&ph_type, ph_idx)
    } else {
        Default::default()
    };

    // ECMA-376 §19.3.1.21 / §20.1.4.2: a slide-level `<p:cNvSpPr txBox="1"/>`
    // marks the shape as a true text box, which means the theme's
    // `<a:txDef>` (rather than `<a:spDef>`) provides the fallback bodyPr.
    let is_text_box = child(sp_node, "nvSpPr")
        .and_then(|n| child(n, "cNvSpPr"))
        .and_then(|n| attr(&n, "txBox"))
        .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
        .unwrap_or(false);
    let shape_kind = if is_text_box {
        ShapeKind::Tx
    } else {
        ShapeKind::Sp
    };
    // Per-level bullets a paragraph inherits when it declares no explicit one
    // (ECMA-376 §19.7.10): the layout/master placeholder cascade for this slot.
    let inherited_level_bullets: LevelBullets = if ph_node.is_some() {
        lph.lookup_level_bullets(&ph_type, ph_idx)
    } else {
        empty_level_bullets()
    };
    let text_body = child(sp_node, "txBody").map(|n| {
        parse_text_body(
            n,
            theme,
            rels,
            inherited_font_size,
            inherited_level_font_sizes,
            inherited_level_indents,
            &inherited_level_bullets,
            inherited_bold,
            inherited_italic,
            inherited_caps.clone(),
            inherited_anchor,
            inherited_alignment,
            inherited_ea_ln_brk,
            inherited_space_before,
            inherited_space_after,
            inherited_line_spacing,
            shape_kind,
            zip,
        )
    });

    // Effects from spPr > effectLst (outerShdw / innerShdw / glow / softEdge /
    // reflection are independent siblings — ECMA-376 §20.1.8.16). Pull each.
    let EffectLst {
        shadow,
        inner_shadow,
        glow,
        soft_edge,
        reflection,
    } = parse_effect_lst(sp_pr.and_then(|p| child(p, "effectLst")), theme);

    Some(ShapeElement {
        x: t.x,
        y: t.y,
        width: t.cx,
        height: cy,
        rotation: t.rot,
        flip_h: t.flip_h,
        flip_v: t.flip_v,
        geometry,
        fill,
        stroke,
        text_body,
        default_text_color,
        cust_geom,
        adj,
        adj2,
        adj3,
        adj4,
        adj5,
        adj6,
        adj7,
        adj8,
        shadow,
        inner_shadow,
        glow,
        soft_edge,
        reflection,
        id: cnv_id,
        name: cnv_name,
        placeholder_type: placeholder_type_out,
        placeholder_idx: ph_idx,
        text_rect: None,
        scene3d: sp_pr.and_then(parse_scene3d),
        sp3d: sp_pr.and_then(parse_sp3d),
    })
}

/// Preset clip geometry for a picture body (§20.1.9.18 — the preset geometry of
/// a `<p:pic>` acts as its clip silhouette and the path its border / contour
/// hug). Returns the prstGeom `prst` name plus every `<a:avLst><a:gd>` adjust
/// value in declaration order (1/1000-of-a-percent OOXML units), so any of the
/// 186 presets — roundRect, ellipse, and the rest — can be reconstructed by the
/// shared preset-geometry engine on the TS side. The preset's own declared
/// defaults fill in any omitted guide, so the `adj` Vec may be shorter than the
/// preset's adjust count (the engine substitutes defaults per index).
///
/// `prst="rect"` returns None: a plain rectangle needs no clip path (the bitmap
/// already fills the bbox), matching the previous behaviour.
pub(crate) fn parse_pic_prst_geom(
    sp_pr: roxmltree::Node<'_, '_>,
) -> (Option<String>, Option<Vec<i64>>) {
    let pg = match child(sp_pr, "prstGeom") {
        Some(n) => n,
        None => return (None, None),
    };
    let prst = match attr(&pg, "prst") {
        Some(p) if p != "rect" => p,
        _ => return (None, None),
    };
    let adjust: Vec<i64> = child(pg, "avLst")
        .map(|av| {
            av.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "gd")
                .filter_map(|gd| {
                    attr(&gd, "fmla")
                        .and_then(|f| f.strip_prefix("val ").and_then(|v| v.parse::<i64>().ok()))
                })
                .collect()
        })
        .unwrap_or_default();
    let adjust = if adjust.is_empty() {
        None
    } else {
        Some(adjust)
    };
    (Some(prst), adjust)
}

/// Microsoft 2016 SVG extension on an `<a:blip>` (ECMA-376 §20.1.8.14 carries
/// the core blip fill; the SVG body is a Microsoft extension). When a `<a:blip>`
/// has `<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">` with an
/// `<asvg:svgBlip r:embed="…">`, resolve that rId to the `.svg` part and return
/// it as a `data:image/svg+xml;base64,…` URL. Returns None when there is no
/// svgBlip, the rId is unresolvable, or the part is missing.
///
/// Matching is by namespace-local element name (`svgBlip`), so the `asvg:`
/// prefix (or any other) is irrelevant.
/// Resolve the Microsoft 2016 `asvg:svgBlip` extension on a `<a:blip>` to the
/// embedded **zip path** of the `.svg` part (e.g. "ppt/media/image2.svg"). The
/// renderer fetches the bytes lazily by path. The part's existence is verified
/// so a dangling rId yields None (no SVG twin), matching the prior behaviour.
pub(crate) fn svg_blip_path(
    blip: roxmltree::Node<'_, '_>,
    slide_dir: &str,
    rels: &HashMap<String, String>,
    zip: &mut PptxZip,
) -> Option<String> {
    let svg_rid = svg_blip_rid(blip)?;
    let svg_target = rels.get(&svg_rid)?;
    let svg_path = resolve_path(slide_dir, svg_target);
    // Existence check only — `index_for_name` reads the central directory
    // without inflating, unlike the former `read_zip_bytes` which decompressed
    // the whole SVG part just to discard the bytes.
    zip.index_for_name(&svg_path)?;
    Some(svg_path)
}

pub(crate) fn parse_picture(
    pic_node: roxmltree::Node<'_, '_>,
    slide_dir: &str,
    rels: &HashMap<String, String>,
    theme: &HashMap<String, String>,
    zip: &mut PptxZip,
) -> Option<PictureElement> {
    let sp_pr = child(pic_node, "spPr")?;
    let xfrm_node = child(sp_pr, "xfrm")?;
    let t = parse_xfrm(xfrm_node);

    if t.cx == 0 || t.cy == 0 {
        return None; // pictures always need explicit dimensions
    }

    let blip_fill = child(pic_node, "blipFill")?;
    let blip = child(blip_fill, "blip")?;

    // Microsoft 2016 SVG extension: the real vector image rides inside
    // `<a:blip><a:extLst>`, while `<a:blip r:embed>` carries a raster (PNG/JPEG)
    // *fallback* for SVG-incapable clients. Resolve the SVG first so a picture
    // that carries ONLY the svgBlip — with no raster `r:embed`, e.g. an icon
    // inserted as a pure SVG — still parses instead of being dropped. Surfaced
    // separately so the renderer can prefer the vector original and fall back to
    // the raster on decode error.
    let svg_image_path = svg_blip_path(blip, slide_dir, rels, zip);

    // Raster blip (`<a:blip r:embed>` → zip path + intrinsic PNG size). Optional:
    // a pure-SVG picture omits it, so this is no longer a hard requirement for
    // the element.
    let raster: Option<(String, Option<(u32, u32)>)> = (|| {
        let r_id = attr_r(&blip, "embed")?;
        let rel_target = rels.get(&r_id)?;
        let path = resolve_path(slide_dir, rel_target);
        let image_bytes = ooxml_common::zip::read_zip_bytes(zip, &path).ok()?;
        // Intrinsic PNG size for the ink-fallback centering (None for non-PNG,
        // unchanged from the former png_size_from_data_url semantics).
        let size = png_size_from_bytes(&image_bytes);
        Some((path, size))
    })();

    // A picture needs at least one drawable source. Prefer the raster as
    // `image_path` (the renderer's srcRect / SVG-decode-failure fallback path);
    // when no raster is embedded, fall back to the SVG part itself so the
    // element is always drawable. `mime_type` matches whichever source wins.
    let (image_path, mime_type, intrinsic_width_px, intrinsic_height_px) =
        match (raster, svg_image_path.as_ref()) {
            (Some((path, size)), _) => {
                let mime = mime_from_ext(&path).to_owned();
                let (w, h) = match size {
                    Some((w, h)) => (Some(w), Some(h)),
                    None => (None, None),
                };
                (path, mime, w, h)
            }
            (None, Some(svg)) => (svg.clone(), "image/svg+xml".to_owned(), None, None),
            (None, None) => return None,
        };

    // ECMA-376 §20.1.9.8 — `<p:pic>` may carry `<a:custGeom>` inside `<p:spPr>`,
    // in which case the bitmap is clipped to that custom path (e.g. a laptop
    // silhouette). Re-use the same parser as for shapes so the renderer can
    // build a Path2D and `ctx.clip()` before drawing the image.
    let cust_geom = child(sp_pr, "custGeom").map(parse_cust_geom);

    // §19.3.1.37: p:pic's spPr is CT_ShapeProperties, so effectLst (§20.1.8.16)
    // applies to images exactly as it does to shapes.
    let EffectLst {
        shadow,
        inner_shadow,
        glow,
        soft_edge,
        reflection,
    } = parse_effect_lst(child(sp_pr, "effectLst"), theme);

    // §20.1.2.2.24 — a `p:pic`'s spPr may carry an `<a:ln>` border (e.g. the
    // white frame of PowerPoint's picture styles). `parse_stroke` returns None
    // for `<a:noFill/>`, so an explicitly border-less picture stays None.
    let stroke = child(sp_pr, "ln").and_then(|n| parse_stroke(n, theme));

    let (prst_geom, prst_adjust) = parse_pic_prst_geom(sp_pr);
    Some(PictureElement {
        x: t.x,
        y: t.y,
        width: t.cx,
        height: t.cy,
        rotation: t.rot,
        flip_h: t.flip_h,
        flip_v: t.flip_v,
        image_path,
        mime_type,
        svg_image_path,
        intrinsic_width_px,
        intrinsic_height_px,
        stroke,
        prst_geom,
        prst_adjust,
        src_rect: parse_src_rect(blip_fill),
        alpha: parse_blip_alpha(blip_fill),
        cust_geom,
        shadow,
        inner_shadow,
        glow,
        soft_edge,
        reflection,
        scene3d: parse_scene3d(sp_pr),
        sp3d: parse_sp3d(sp_pr),
    })
}

/// Decode (width, height) from raw PNG bytes by reading the IHDR chunk. Returns
/// None for non-PNG payloads or malformed data (unchanged semantics from the
/// former `png_size_from_data_url`, only the input is now raw bytes instead of
/// a base64 data URL).
pub(crate) fn png_size_from_bytes(bytes: &[u8]) -> Option<(u32, u32)> {
    if bytes.len() < 24 || &bytes[0..8] != b"\x89PNG\r\n\x1a\n" {
        return None;
    }
    let w = u32::from_be_bytes([bytes[16], bytes[17], bytes[18], bytes[19]]);
    let h = u32::from_be_bytes([bytes[20], bytes[21], bytes[22], bytes[23]]);
    Some((w, h))
}

/// EMU per CSS pixel at PowerPoint's default 96 DPI (914400 EMU/inch ÷ 96).
pub(crate) const EMU_PER_PX_96DPI: i64 = 9525;

/// If a `p:pic` declares an `a:audioFile` / `a:videoFile` in its `nvPr`
/// (or the newer `p14:media` extension), emit a `MediaElement` with the
/// poster image and the media bytes. Returns None for regular pictures.
pub(crate) fn parse_media(
    pic_node: roxmltree::Node<'_, '_>,
    slide_dir: &str,
    rels: &HashMap<String, String>,
) -> Option<MediaElement> {
    let nv_pic_pr = child(pic_node, "nvPicPr")?;
    let nv_pr = child(nv_pic_pr, "nvPr")?;

    // A `<p:pic>` can carry several media references at once: the legacy
    // `<a:videoFile>`/`<a:audioFile>` (r:embed or r:link) and the modern
    // `<p14:media r:embed>` extension. We collect every reference and use the
    // first that actually resolves, preferring an *embedded* ref over a
    // *linked* one. Embedded refs always point at internal media we can read;
    // a linked ref may target an External resource we cannot fetch, or its rId
    // may be missing/unresolvable (caught below via `rels.get`). Trying embeds
    // first means a present `<p14:media>` embed is used even when the deck also
    // carries a videoFile link, instead of being shadowed and demoted to a
    // poster-only Picture. (Note: parse_rels stores Id+Target and never reads
    // TargetMode, so a real External link keeps its non-empty URL here.)
    // (ECMA-376 §19.3.1.17/18 a:videoFile/a:audioFile; p14:media is a Microsoft
    // extension.)
    let av = nv_pr
        .children()
        .find(|n| n.is_element() && matches!(n.tag_name().name(), "videoFile" | "audioFile"));
    let av_kind: Option<&str> = av.map(|n| {
        if n.tag_name().name() == "videoFile" {
            "video"
        } else {
            "audio"
        }
    });
    let av_embed = av.and_then(|n| attr_r(&n, "embed"));
    let av_link = av.and_then(|n| attr_r(&n, "link"));

    // `<p14:media>` lives in `nvPr > extLst > ext`.
    let p14 = child(nv_pr, "extLst").and_then(|ext_lst| {
        ext_lst
            .children()
            .filter(|c| c.is_element())
            .find_map(|ext| {
                ext.children()
                    .find(|m| m.is_element() && m.tag_name().name() == "media")
            })
    });
    let p14_embed = p14.and_then(|n| attr_r(&n, "embed"));
    let p14_link = p14.and_then(|n| attr_r(&n, "link"));

    // Preference: embedded refs (always internal) before linked refs (a link may
    // be External / unresolved); within each, the kind-bearing
    // videoFile/audioFile before the kind-less p14:media.
    let candidates = [
        (av_kind, av_embed),
        (None, p14_embed),
        (av_kind, av_link),
        (None, p14_link),
    ];
    let (media_kind, media_path, mime) = candidates.into_iter().find_map(|(kind_hint, rid)| {
        let rid = rid?;
        let target = rels.get(&rid)?;
        if target.is_empty() {
            return None; // malformed rel with an empty Target (External links carry a non-empty URL)
        }
        let path = resolve_path(slide_dir, target);
        let mime = mime_from_ext(&path);
        let kind = match kind_hint {
            Some(k) => k.to_string(),
            None if mime.starts_with("video/") => "video".to_string(),
            None if mime.starts_with("audio/") => "audio".to_string(),
            None => return None, // p14:media with an unknown MIME — try the next ref
        };
        Some((kind, path, mime))
    })?;

    // Geometry
    let sp_pr = child(pic_node, "spPr")?;
    let xfrm_node = child(sp_pr, "xfrm")?;
    let t = parse_xfrm(xfrm_node);
    if t.cx == 0 || t.cy == 0 {
        return None;
    }

    // Poster image (from blipFill). Optional — the renderer falls back to a
    // solid fill when the poster is absent or still loading. We emit just the
    // zip path so the main thread can lazily `getMedia` it; embedding large
    // posters inline ballooned the parse output for video-heavy decks.
    let (poster_path, poster_mime_type) = (|| -> Option<(String, String)> {
        let blip_fill = child(pic_node, "blipFill")?;
        let r_id = child(blip_fill, "blip").and_then(|b| attr_r(&b, "embed"))?;
        let rel_target = rels.get(&r_id)?;
        let image_path = resolve_path(slide_dir, rel_target);
        let img_mime = mime_from_ext(&image_path).to_string();
        Some((image_path, img_mime))
    })()
    .unwrap_or_default();

    Some(MediaElement {
        x: t.x,
        y: t.y,
        width: t.cx,
        height: t.cy,
        media_kind,
        poster_path,
        poster_mime_type,
        media_path,
        mime_type: mime.to_string(),
    })
}

/// Parse ppt/tableStyles.xml into a map of styleId → TableStyleDef.
pub(crate) fn parse_table_styles_xml(
    xml: &str,
    theme: &HashMap<String, String>,
) -> HashMap<String, TableStyleDef> {
    let mut map = HashMap::new();
    let Ok(doc) = roxmltree::Document::parse(xml) else {
        return map;
    };
    let root = doc.root_element();
    for style_node in root
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "tblStyle")
    {
        let Some(style_id) = attr(&style_node, "styleId") else {
            continue;
        };
        let style_id = style_id.to_string();
        let mut def = TableStyleDef::default();

        // (fill, inside-horizontal, inside-vertical, diagonal/unused, outer-horizontal,
        // outer-vertical) borders for one table-style role (ECMA-376 §20.1.4.2.27).
        type TcStyle = (
            Option<Fill>,
            Option<Stroke>,
            Option<Stroke>,
            Option<Stroke>,
            Option<Stroke>,
            Option<Stroke>,
        );
        let parse_tc_style = |role: roxmltree::Node<'_, '_>| -> TcStyle {
            let tc_style = match child(role, "tcStyle") {
                Some(n) => n,
                None => return (None, None, None, None, None, None),
            };
            // ECMA-376 §20.1.4.2.27 — the cell fill is wrapped in `<a:fill>`
            // (CT_FillProperties), so descend into it before reading the actual
            // solidFill/gradFill. Fall back to `<a:fillRef>` (theme style reference).
            let fill = child(tc_style, "fill")
                .and_then(|f| parse_table_style_fill(f, theme))
                .or_else(|| {
                    child(tc_style, "fillRef").and_then(|fr| {
                        let idx: u32 = attr(&fr, "idx").and_then(|v| v.parse().ok()).unwrap_or(0);
                        if idx == 0 {
                            Some(Fill::None)
                        } else {
                            parse_color_node(fr, theme).map(|c| Fill::Solid { color: c })
                        }
                    })
                });
            let tc_bdr = child(tc_style, "tcBdr");
            let parse_side = |side: &str| -> Option<Stroke> {
                let side_node = tc_bdr.and_then(|b| child(b, side))?;
                // Explicit <a:ln>
                if let Some(ln) = child(side_node, "ln") {
                    return parse_stroke(ln, theme);
                }
                // <a:lnRef idx="N">: use standard themed width + provided color
                if let Some(ln_ref) = child(side_node, "lnRef") {
                    let idx: u32 = attr(&ln_ref, "idx")
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(0);
                    if idx == 0 {
                        return None;
                    }
                    let color = parse_color_node(ln_ref, theme)?;
                    let width: i64 = match idx {
                        1 => 6350,
                        2 => 12700,
                        _ => 19050,
                    };
                    return Some(Stroke {
                        color,
                        width,
                        dash_style: None,
                        head_end: None,
                        tail_end: None,
                        cmpd: None,
                    });
                }
                None
            };
            let inside_h = parse_side("insideH");
            let inside_v = parse_side("insideV");
            let border_b = parse_side("bottom");
            let outer_h = parse_side("top");
            let outer_v = parse_side("left");
            (fill, inside_h, inside_v, border_b, outer_h, outer_v)
        };

        // Default text colour for a role: `<a:tcTxStyle>` holds a `<a:fontRef>`
        // followed by a colour child (schemeClr/srgbClr). parse_color_node scans
        // direct children and skips fontRef, picking up the colour.
        let parse_tc_tx_color = |role: roxmltree::Node<'_, '_>| -> Option<String> {
            child(role, "tcTxStyle").and_then(|t| parse_color_node(t, theme))
        };
        // `<a:tcTxStyle b="on">` → bold for this role (e.g. a bold header row).
        let parse_tc_tx_bold = |role: roxmltree::Node<'_, '_>| -> Option<bool> {
            child(role, "tcTxStyle")
                .and_then(|t| attr(&t, "b"))
                .map(|v| v == "on" || v == "1" || v == "true")
        };

        if let Some(whole) = child(style_node, "wholeTbl") {
            let (fill, ih, iv, _, oh, ov) = parse_tc_style(whole);
            def.whole_fill = fill;
            def.whole_inside_h = ih;
            def.whole_inside_v = iv;
            def.whole_outer_h = oh;
            def.whole_outer_v = ov;
            def.whole_text_color = parse_tc_tx_color(whole);
        }
        if let Some(band) = child(style_node, "band1H") {
            let (fill, _, _, _, _, _) = parse_tc_style(band);
            def.band1h_fill = fill;
        }
        if let Some(band) = child(style_node, "band2H") {
            let (fill, _, _, _, _, _) = parse_tc_style(band);
            def.band2h_fill = fill;
        }
        if let Some(first) = child(style_node, "firstRow") {
            let (fill, _, _, border_b, _, _) = parse_tc_style(first);
            def.first_row_fill = fill;
            def.first_row_border_b = border_b;
            def.first_row_text_color = parse_tc_tx_color(first);
            def.first_row_bold = parse_tc_tx_bold(first);
        }
        if let Some(last) = child(style_node, "lastRow") {
            let (fill, _, _, _, _, _) = parse_tc_style(last);
            def.last_row_fill = fill;
            def.last_row_text_color = parse_tc_tx_color(last);
            def.last_row_bold = parse_tc_tx_bold(last);
        }
        if let Some(first) = child(style_node, "firstCol") {
            let (fill, _, _, _, _, _) = parse_tc_style(first);
            def.first_col_fill = fill;
            def.first_col_text_color = parse_tc_tx_color(first);
            def.first_col_bold = parse_tc_tx_bold(first);
        }
        if let Some(last) = child(style_node, "lastCol") {
            let (fill, _, _, _, _, _) = parse_tc_style(last);
            def.last_col_fill = fill;
            def.last_col_text_color = parse_tc_tx_color(last);
            def.last_col_bold = parse_tc_tx_bold(last);
        }

        map.insert(style_id, def);
    }
    map
}

pub(crate) fn parse_table(
    tbl: roxmltree::Node<'_, '_>,
    t: &Transform,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
    zip: &mut PptxZip,
) -> Option<TableElement> {
    // Parse tblPr attributes and look up table style
    let tbl_pr = child(tbl, "tblPr");
    let style_id = tbl_pr
        .and_then(|n| child(n, "tableStyleId"))
        .and_then(|n| n.text())
        .map(|s| s.to_string());
    let flag = |attr_name: &str| -> bool {
        tbl_pr
            .and_then(|n| attr(&n, attr_name))
            .map(|v| v == "1" || v == "true")
            .unwrap_or(false)
    };
    let first_row = flag("firstRow");
    let last_row = flag("lastRow");
    // ECMA-376 §21.1.3.13 (a:tblPr@rtl): right-to-left table layout.
    let rtl = flag("rtl");
    let band_row = flag("bandRow");
    let first_col = flag("firstCol");
    let last_col = flag("lastCol");

    // Load style definitions once
    let table_styles_xml = read_zip_str(zip, "ppt/tableStyles.xml").ok();
    let table_styles = table_styles_xml
        .as_deref()
        .map(|xml| parse_table_styles_xml(xml, theme))
        .unwrap_or_default();
    let style_owned: Option<TableStyleDef> = style_id.as_deref().and_then(|id| {
        table_styles
            .get(id)
            .cloned()
            .or_else(|| table_style_presets::lookup_builtin_table_style(id, theme))
    });
    let style = style_owned.as_ref();

    let cols: Vec<i64> = tbl
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "tblGrid")
        .map(|grid| {
            grid.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "gridCol")
                .filter_map(|n| attr_i64(&n, "w"))
                .collect()
        })
        .unwrap_or_default();

    if cols.is_empty() {
        return None;
    }

    let col_count = cols.len();
    let last_col_idx = col_count.saturating_sub(1);

    let mut rows: Vec<TableRow> = tbl
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "tr")
        .map(|tr| parse_table_row(tr, theme, rels, zip))
        .collect();

    let row_count = rows.len();
    let last_row_idx = row_count.saturating_sub(1);

    // Apply table style fills and borders to each cell
    for (ri, row) in rows.iter_mut().enumerate() {
        for (ci, cell) in row.cells.iter_mut().enumerate() {
            if let Some(s) = style {
                // ── Fill cascade ────────────────────────────────────────────
                let mut effective_fill = s.whole_fill.clone();

                if band_row {
                    // Determine band index excluding firstRow header if present
                    let band_ri = ri.saturating_sub(if first_row { 1 } else { 0 });
                    if !(first_row && ri == 0) {
                        if band_ri % 2 == 0 {
                            if let Some(f) = s.band1h_fill.clone() {
                                effective_fill = Some(f);
                            }
                        } else if let Some(f) = s.band2h_fill.clone() {
                            effective_fill = Some(f);
                        }
                    }
                }
                if first_row && ri == 0 {
                    if let Some(f) = s.first_row_fill.clone() {
                        effective_fill = Some(f);
                    }
                }
                if last_row && ri == last_row_idx {
                    if let Some(f) = s.last_row_fill.clone() {
                        effective_fill = Some(f);
                    }
                }
                if first_col && ci == 0 {
                    if let Some(f) = s.first_col_fill.clone() {
                        effective_fill = Some(f);
                    }
                }
                if last_col && ci == last_col_idx {
                    if let Some(f) = s.last_col_fill.clone() {
                        effective_fill = Some(f);
                    }
                }
                // Cell's own tcPr fill wins
                if cell.fill.is_none() {
                    cell.fill = effective_fill;
                }

                // ── Text-colour cascade (style `<a:tcTxStyle>`, role-keyed) ──
                // Mirrors the fill cascade ordering. The header (firstRow) typically
                // overrides wholeTbl (e.g. white-on-accent header). A run's own
                // explicit colour still wins later, at render time.
                let mut effective_text = s.whole_text_color.clone();
                if first_row && ri == 0 {
                    if let Some(c) = s.first_row_text_color.clone() {
                        effective_text = Some(c);
                    }
                }
                if last_row && ri == last_row_idx {
                    if let Some(c) = s.last_row_text_color.clone() {
                        effective_text = Some(c);
                    }
                }
                if first_col && ci == 0 {
                    if let Some(c) = s.first_col_text_color.clone() {
                        effective_text = Some(c);
                    }
                }
                if last_col && ci == last_col_idx {
                    if let Some(c) = s.last_col_text_color.clone() {
                        effective_text = Some(c);
                    }
                }
                if cell.text_color.is_none() {
                    cell.text_color = effective_text;
                }

                // ── Bold cascade (style `<a:tcTxStyle b="on">`, role-keyed) ──
                // Applied as the cell text body's default bold; a run's own @b wins.
                let mut effective_bold: Option<bool> = None;
                if first_row && ri == 0 {
                    effective_bold = effective_bold.or(s.first_row_bold);
                }
                if last_row && ri == last_row_idx {
                    effective_bold = effective_bold.or(s.last_row_bold);
                }
                if first_col && ci == 0 {
                    effective_bold = effective_bold.or(s.first_col_bold);
                }
                if last_col && ci == last_col_idx {
                    effective_bold = effective_bold.or(s.last_col_bold);
                }
                if let (Some(b), Some(tb)) = (effective_bold, cell.text_body.as_mut()) {
                    if tb.default_bold.is_none() {
                        tb.default_bold = Some(b);
                    }
                }

                // ── Border cascade (style provides inside and outer borders) ──
                // Outer top edge
                if cell.border_t.is_none() && ri == 0 {
                    cell.border_t = s.whole_outer_h.clone();
                }
                // Inner horizontal separator between rows
                if cell.border_t.is_none() && ri > 0 {
                    cell.border_t = s.whole_inside_h.clone();
                }
                // Outer bottom edge
                if cell.border_b.is_none() && ri == last_row_idx {
                    cell.border_b = s.whole_outer_h.clone();
                }
                // Inner bottom separator; firstRow gets its own bottom definition
                if cell.border_b.is_none() {
                    if first_row && ri == 0 {
                        cell.border_b = s
                            .first_row_border_b
                            .clone()
                            .or_else(|| s.whole_inside_h.clone());
                    } else if ri < last_row_idx {
                        cell.border_b = s.whole_inside_h.clone();
                    }
                }
                // Outer left edge
                if cell.border_l.is_none() && ci == 0 {
                    cell.border_l = s.whole_outer_v.clone();
                }
                // Inner vertical separator between cols
                if cell.border_l.is_none() && ci > 0 {
                    cell.border_l = s.whole_inside_v.clone();
                }
                // Outer right edge
                if cell.border_r.is_none() && ci == last_col_idx {
                    cell.border_r = s.whole_outer_v.clone();
                }
                // Inner right separator
                if cell.border_r.is_none() && ci < last_col_idx {
                    cell.border_r = s.whole_inside_v.clone();
                }
            } else {
                // ── Fallback for built-in styles not defined in tableStyles.xml ──
                // Approximate "Medium Style 2": accent1 header fill + thin outer box + row separators.
                let thin = Stroke {
                    color: "A0A096".to_string(),
                    width: 9525,
                    dash_style: None,
                    head_end: None,
                    tail_end: None,
                    cmpd: None,
                };
                if cell.fill.is_none() && first_row && ri == 0 {
                    if let Some(color) = theme.get("accent1") {
                        cell.fill = Some(Fill::Solid {
                            color: color.clone(),
                        });
                    }
                }
                // Outer top
                if cell.border_t.is_none() && ri == 0 {
                    cell.border_t = Some(thin.clone());
                }
                // Inner horizontal separators
                if cell.border_t.is_none() && ri > 0 {
                    cell.border_t = Some(thin.clone());
                }
                // Outer bottom
                if cell.border_b.is_none() && ri == last_row_idx {
                    cell.border_b = Some(thin.clone());
                }
                // Outer left edge
                if cell.border_l.is_none() && ci == 0 {
                    cell.border_l = Some(thin.clone());
                }
                // Outer right edge
                if cell.border_r.is_none() && ci == last_col_idx {
                    cell.border_r = Some(thin.clone());
                }
            }
        }
    }

    Some(TableElement {
        x: t.x,
        y: t.y,
        width: t.cx,
        height: t.cy,
        cols,
        rows,
        rtl,
    })
}

pub(crate) fn parse_table_row(
    tr: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
    zip: &mut PptxZip,
) -> TableRow {
    let height = attr_i64(&tr, "h").unwrap_or(0);
    let cells: Vec<TableCell> = tr
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "tc")
        .map(|tc| parse_table_cell(tc, theme, rels, zip))
        .collect();
    TableRow { height, cells }
}

pub(crate) fn parse_table_cell(
    tc: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
    zip: &mut PptxZip,
) -> TableCell {
    let tc_pr = child(tc, "tcPr");
    // tcPr > anchor controls vertical text alignment within the cell
    let anchor = tc_pr
        .and_then(|n| attr(&n, "anchor"))
        .map(|a| a.to_string());
    let text_body = child(tc, "txBody").map(|n| {
        parse_text_body(
            n,
            theme,
            rels,
            None,
            [None; 9],
            Default::default(), // inherited_level_indents
            &empty_level_bullets(),
            None,
            None,
            None,
            anchor,
            None, // inherited_alignment
            None, // inherited_ea_ln_brk
            None, // inherited_space_before
            None, // inherited_space_after
            None, // inherited_line_spacing
            ShapeKind::Sp,
            zip,
        )
    });

    let fill = tc_pr.and_then(|n| parse_fill(n, theme));

    let border_l = tc_pr
        .and_then(|n| child(n, "lnL"))
        .and_then(|n| parse_stroke(n, theme));
    let border_r = tc_pr
        .and_then(|n| child(n, "lnR"))
        .and_then(|n| parse_stroke(n, theme));
    let border_t = tc_pr
        .and_then(|n| child(n, "lnT"))
        .and_then(|n| parse_stroke(n, theme));
    let border_b = tc_pr
        .and_then(|n| child(n, "lnB"))
        .and_then(|n| parse_stroke(n, theme));
    // Diagonal borders: lnTlToBr (top-left→bottom-right) and lnBlToTr (bottom-left→top-right)
    // These are CT_LineProperties elements (same structure as lnL/lnR/lnT/lnB)
    let diagonal_tl = tc_pr
        .and_then(|n| child(n, "lnTlToBr"))
        .and_then(|n| parse_stroke(n, theme));
    let diagonal_tr = tc_pr
        .and_then(|n| child(n, "lnBlToTr"))
        .and_then(|n| parse_stroke(n, theme));

    let grid_span: u32 = attr(&tc, "gridSpan")
        .and_then(|v| v.parse().ok())
        .unwrap_or(1);
    let row_span: u32 = attr(&tc, "rowSpan")
        .and_then(|v| v.parse().ok())
        .unwrap_or(1);
    let h_merge = attr(&tc, "hMerge")
        .map(|v| v == "1" || v == "true")
        .unwrap_or(false);
    let v_merge = attr(&tc, "vMerge")
        .map(|v| v == "1" || v == "true")
        .unwrap_or(false);

    TableCell {
        text_body,
        fill,
        text_color: None,
        border_l,
        border_r,
        border_t,
        border_b,
        diagonal_tl,
        diagonal_tr,
        grid_span,
        row_span,
        h_merge,
        v_merge,
    }
}

/// Walk a master/layout `<p:cSld><p:spTree>` and collect its NON-placeholder
/// (decorative) shapes as `SlideElement`s (ECMA-376 §19.3.1.38 showMasterSp /
/// layout decorations). Placeholders are skipped (`skip_placeholders = true`).
/// Extracted verbatim from the two inline spTree walks in `parse_slide` so the
/// master decorations can be pre-computed once per cached master; `theme`,
/// `rels`, `smartart_drawings`, and `part_dir` are the resolution context of the
/// tree being walked. Appends to `out` (callers control ordering/gating).
pub(crate) fn extract_decorative_shapes(
    root: roxmltree::Node<'_, '_>,
    part_dir: &str,
    rels: &HashMap<String, String>,
    smartart_drawings: &HashMap<String, String>,
    theme: &HashMap<String, String>,
    zip: &mut PptxZip,
    out: &mut Vec<SlideElement>,
) {
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        let empty_lph = LayoutPlaceholders::default();
        for node in sp_tree.children().filter(|n| n.is_element()) {
            parse_sp_tree_node(
                node,
                &empty_lph,
                part_dir,
                rels,
                smartart_drawings,
                zip,
                theme,
                out,
                true, // skip placeholder shapes
                None, // no inherited group fill at top level
            );
        }
    }
}

// Recurses the shape tree carrying placeholder/layout context, theme, rels,
// smartart drawings, the zip, the output buffer and inherited group fill.
#[allow(clippy::too_many_arguments)]
pub(crate) fn parse_sp_tree_node(
    node: roxmltree::Node<'_, '_>,
    lph: &LayoutPlaceholders,
    slide_dir: &str,
    rels: &HashMap<String, String>,
    smartart_drawings: &HashMap<String, String>,
    zip: &mut PptxZip,
    theme: &HashMap<String, String>,
    out: &mut Vec<SlideElement>,
    skip_placeholders: bool,
    group_fill: Option<&Fill>,
) {
    match node.tag_name().name() {
        "sp" => {
            if skip_placeholders && is_placeholder(node) {
                return;
            }
            // Image-filled shape: spPr > blipFill > blip r:embed → render as PictureElement
            let sp_pr_node = child(node, "spPr");
            let blip_fill_node = sp_pr_node.and_then(|p| child(p, "blipFill"));
            let blip_rid = blip_fill_node
                .and_then(|bf| child(bf, "blip"))
                .and_then(|b| attr_r(&b, "embed"));
            if let Some(ref rid) = blip_rid {
                if let Some(xfrm_node) = sp_pr_node.and_then(|p| child(p, "xfrm")) {
                    let t = parse_xfrm(xfrm_node);
                    if t.cx > 0 && t.cy > 0 {
                        if let Some(target) = rels.get(rid) {
                            let image_path = resolve_path(slide_dir, target);
                            if let Ok(bytes) = ooxml_common::zip::read_zip_bytes(zip, &image_path) {
                                let mime_type = mime_from_ext(&image_path).to_owned();
                                let (intrinsic_width_px, intrinsic_height_px) =
                                    match png_size_from_bytes(&bytes) {
                                        Some((w, h)) => (Some(w), Some(h)),
                                        None => (None, None),
                                    };
                                // Microsoft 2016 SVG extension — a blipFill-painted
                                // sp can carry the same svgBlip vector original as a
                                // real p:pic; surface it so the renderer prefers it.
                                let svg_image_path = blip_fill_node
                                    .and_then(|bf| child(bf, "blip"))
                                    .and_then(|b| svg_blip_path(b, slide_dir, rels, zip));
                                // §20.1.9.18 — the sp's prstGeom (any preset, not
                                // just roundRect) is the picture's clip silhouette.
                                let (prst_geom, prst_adjust) =
                                    sp_pr_node.map(parse_pic_prst_geom).unwrap_or((None, None));
                                let cust_geom = sp_pr_node
                                    .and_then(|p| child(p, "custGeom"))
                                    .map(parse_cust_geom);
                                // §20.1.8.16 effectLst applies to a sp painted as
                                // a picture (blipFill) just like a regular p:pic.
                                let EffectLst {
                                    shadow,
                                    inner_shadow,
                                    glow,
                                    soft_edge,
                                    reflection,
                                } = parse_effect_lst(
                                    sp_pr_node.and_then(|p| child(p, "effectLst")),
                                    theme,
                                );
                                // §20.1.2.2.24 — a blipFill-painted sp can carry
                                // an `<a:ln>` border just like a real p:pic.
                                let stroke = sp_pr_node
                                    .and_then(|p| child(p, "ln"))
                                    .and_then(|n| parse_stroke(n, theme));
                                out.push(SlideElement::Picture(PictureElement {
                                    x: t.x,
                                    y: t.y,
                                    width: t.cx,
                                    height: t.cy,
                                    rotation: t.rot,
                                    flip_h: t.flip_h,
                                    flip_v: t.flip_v,
                                    image_path,
                                    mime_type,
                                    svg_image_path,
                                    intrinsic_width_px,
                                    intrinsic_height_px,
                                    stroke,
                                    prst_geom,
                                    prst_adjust,
                                    src_rect: blip_fill_node.and_then(parse_src_rect),
                                    alpha: blip_fill_node.and_then(parse_blip_alpha),
                                    cust_geom,
                                    shadow,
                                    inner_shadow,
                                    glow,
                                    soft_edge,
                                    reflection,
                                    scene3d: sp_pr_node.and_then(parse_scene3d),
                                    sp3d: sp_pr_node.and_then(parse_sp3d),
                                }));
                                return;
                            }
                        }
                    }
                }
            }
            // Picture-placeholder inheritance: slide sp has a ph but no own blipFill →
            // look up an inherited blipFill from the layout placeholder. Transform
            // comes from the slide's xfrm when present, otherwise from the layout.
            if blip_rid.is_none() {
                if let Some(ph) = node
                    .descendants()
                    .find(|n| n.is_element() && n.tag_name().name() == "ph")
                {
                    let ph_type = attr(&ph, "type").unwrap_or_else(|| "body".into());
                    let ph_idx: Option<u32> = attr(&ph, "idx").and_then(|v| v.parse().ok());
                    if let Some(bf) = lph.lookup_blip_fill(&ph_type, ph_idx) {
                        let slide_xfrm = sp_pr_node.and_then(|p| child(p, "xfrm")).map(parse_xfrm);
                        let t = slide_xfrm.or_else(|| lph.lookup(&ph_type, ph_idx).cloned());
                        if let Some(t) = t {
                            if t.cx > 0 && t.cy > 0 {
                                // §20.1.2.2.24 — honour the slide sp's own `<a:ln>`
                                // border, falling back to the inherited layout
                                // placeholder stroke when the slide omits one.
                                let stroke = match sp_pr_node.and_then(|p| child(p, "ln")) {
                                    Some(ln) => parse_stroke(ln, theme),
                                    None => lph.lookup_stroke(&ph_type, ph_idx),
                                };
                                out.push(SlideElement::Picture(PictureElement {
                                    x: t.x,
                                    y: t.y,
                                    width: t.cx,
                                    height: t.cy,
                                    rotation: t.rot,
                                    flip_h: t.flip_h,
                                    flip_v: t.flip_v,
                                    image_path: bf.image_path,
                                    mime_type: bf.mime_type,
                                    // TODO: an inherited layout-placeholder blipFill
                                    // (LayoutPlaceholders::lookup_blip_fill) does not
                                    // yet carry the svgBlip extension. Picture
                                    // placeholders pointing at an SVG are rare; thread
                                    // the svg path through BlipFill if a sample needs it.
                                    svg_image_path: None,
                                    // Intrinsic size is only consumed by the ink
                                    // fallback (PNG-IHDR centering); inherited
                                    // placeholder pictures stretch to the box, so
                                    // None matches the prior behaviour.
                                    intrinsic_width_px: None,
                                    intrinsic_height_px: None,
                                    stroke,
                                    prst_geom: None,
                                    prst_adjust: None,
                                    src_rect: bf.src_rect,
                                    alpha: bf.alpha,
                                    cust_geom: None,
                                    shadow: None,
                                    inner_shadow: None,
                                    glow: None,
                                    soft_edge: None,
                                    reflection: None,
                                    scene3d: None,
                                    sp3d: None,
                                }));
                                return;
                            }
                        }
                    }
                }
            }
            if let Some(shape) = parse_shape(node, lph, theme, rels, group_fill, zip) {
                out.push(SlideElement::Shape(shape));
            }
        }
        "pic" => {
            if let Some(media) = parse_media(node, slide_dir, rels) {
                out.push(SlideElement::Media(media));
            } else if let Some(pic) = parse_picture(node, slide_dir, rels, theme, zip) {
                out.push(SlideElement::Picture(pic));
            } else {
                // Placeholder pic: no xfrm in spPr — position comes from layout by_idx
                let ph_idx = node
                    .descendants()
                    .find(|n| n.is_element() && n.tag_name().name() == "ph")
                    .and_then(|ph| attr(&ph, "idx"))
                    .and_then(|s| s.parse::<u32>().ok());
                if let Some(idx) = ph_idx {
                    if let Some(t) = lph.by_idx.get(&idx) {
                        let blip_fill = child(node, "blipFill");
                        let blip = blip_fill.and_then(|bf| child(bf, "blip"));
                        let r_id = blip.and_then(|b| attr_r(&b, "embed"));
                        if let Some(rid) = r_id {
                            if let Some(rel_target) = rels.get(&rid) {
                                let image_path = resolve_path(slide_dir, rel_target);
                                if let Ok(image_bytes) =
                                    ooxml_common::zip::read_zip_bytes(zip, &image_path)
                                {
                                    let mime_type = mime_from_ext(&image_path).to_owned();
                                    let (intrinsic_width_px, intrinsic_height_px) =
                                        match png_size_from_bytes(&image_bytes) {
                                            Some((w, h)) => (Some(w), Some(h)),
                                            None => (None, None),
                                        };
                                    // Microsoft 2016 SVG extension on the placeholder
                                    // p:pic's blip — prefer the vector original.
                                    let svg_image_path =
                                        blip.and_then(|b| svg_blip_path(b, slide_dir, rels, zip));
                                    // §20.1.2.2.24 — placeholder pic border: the
                                    // p:pic's own `<a:ln>`, else the inherited
                                    // layout placeholder stroke.
                                    let stroke =
                                        match child(node, "spPr").and_then(|p| child(p, "ln")) {
                                            Some(ln) => parse_stroke(ln, theme),
                                            None => lph.by_idx_stroke.get(&idx).cloned(),
                                        };
                                    out.push(SlideElement::Picture(PictureElement {
                                        x: t.x,
                                        y: t.y,
                                        width: t.cx,
                                        height: t.cy,
                                        rotation: t.rot,
                                        flip_h: t.flip_h,
                                        flip_v: t.flip_v,
                                        image_path,
                                        mime_type,
                                        svg_image_path,
                                        intrinsic_width_px,
                                        intrinsic_height_px,
                                        stroke,
                                        prst_geom: None,
                                        prst_adjust: None,
                                        src_rect: blip_fill.and_then(parse_src_rect),
                                        alpha: blip_fill.and_then(parse_blip_alpha),
                                        cust_geom: None,
                                        shadow: None,
                                        inner_shadow: None,
                                        glow: None,
                                        soft_edge: None,
                                        reflection: None,
                                        scene3d: None,
                                        sp3d: None,
                                    }));
                                }
                            }
                        }
                    }
                }
            }
        }
        "AlternateContent" => {
            // mc:AlternateContent wraps modern elements (e.g. chartEx inside grpSp,
            // or p:contentPart for ink/handwriting). Try Choice first; if it produces
            // nothing (Choice contains an element we don't render, like contentPart),
            // fall back to Fallback which usually carries a rasterized p:pic alternative.
            let before = out.len();
            let choice_node = node
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "Choice");
            // Choice contains p:contentPart → ink/handwriting. PowerPoint renders the
            // InkML directly; the Fallback PNG is a low-resolution rasterization at
            // the stroke's actual pixel extent. For empty/single-tap strokes the PNG
            // is only a few pixels and should not be stretched to the bounding box.
            let is_ink_fallback = choice_node.is_some_and(|c| {
                c.descendants()
                    .any(|n| n.is_element() && n.tag_name().name() == "contentPart")
            });
            if let Some(choice_node) = choice_node {
                for child_node in choice_node.children().filter(|n| n.is_element()) {
                    parse_sp_tree_node(
                        child_node,
                        lph,
                        slide_dir,
                        rels,
                        smartart_drawings,
                        zip,
                        theme,
                        out,
                        skip_placeholders,
                        group_fill,
                    );
                }
            }
            if out.len() == before {
                if let Some(fallback_node) = node
                    .children()
                    .find(|n| n.is_element() && n.tag_name().name() == "Fallback")
                {
                    for child_node in fallback_node.children().filter(|n| n.is_element()) {
                        parse_sp_tree_node(
                            child_node,
                            lph,
                            slide_dir,
                            rels,
                            smartart_drawings,
                            zip,
                            theme,
                            out,
                            skip_placeholders,
                            group_fill,
                        );
                    }
                }
                if is_ink_fallback {
                    // Render the fallback PNG at its natural pixel size centered
                    // inside the original bounding box, so a 6×6 px empty-stroke
                    // PNG is drawn as a 6×6 px dot rather than stretched into a
                    // blocky cross. Visible strokes whose PNG natural size already
                    // matches the box keep their existing extent.
                    for el in &mut out[before..] {
                        if let SlideElement::Picture(p) = el {
                            if let (Some(nat_w_px), Some(nat_h_px)) =
                                (p.intrinsic_width_px, p.intrinsic_height_px)
                            {
                                let nat_w = (nat_w_px as i64) * EMU_PER_PX_96DPI;
                                let nat_h = (nat_h_px as i64) * EMU_PER_PX_96DPI;
                                if nat_w < p.width && nat_h < p.height {
                                    p.x += (p.width - nat_w) / 2;
                                    p.y += (p.height - nat_h) / 2;
                                    p.width = nat_w;
                                    p.height = nat_h;
                                }
                            }
                        }
                    }
                }
            }
        }
        "graphicFrame" => {
            let xfrm_node = child(node, "xfrm");
            let t = xfrm_node.map(parse_xfrm).unwrap_or_default();

            // Table
            let tbl_node = node
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "tbl");
            if let Some(tbl_node) = tbl_node {
                if let Some(table) = parse_table(tbl_node, &t, theme, rels, zip) {
                    out.push(SlideElement::Table(table));
                }
                return;
            }

            // SmartArt: render the PowerPoint-prebaked fallback drawing1.xml when present.
            // Layout-engine reconstruction is not implemented; we rely on the cached dsp:spTree.
            if let Some(gd) = node
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "graphicData")
            {
                let uri = attr(&gd, "uri").unwrap_or_default();
                if uri == "http://schemas.openxmlformats.org/drawingml/2006/diagram" {
                    if let Some(rel_ids) = child(gd, "relIds") {
                        if let Some(dm_rid) = attr_r(&rel_ids, "dm") {
                            if let Some(drawing_xml) = smartart_drawings.get(&dm_rid) {
                                parse_smartart_drawing(drawing_xml, &t, theme, out, zip);
                                return;
                            }
                        }
                    }
                }
            }

            // Chart
            if let Some(gd) = node
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "graphicData")
            {
                let uri = attr(&gd, "uri").unwrap_or_default();
                // Both <c:chart> and <cx:chart> share the local name "chart"
                let chart_rid = gd
                    .descendants()
                    .find(|n| n.is_element() && n.tag_name().name() == "chart")
                    .and_then(|n| attr_r(&n, "id"));
                if let Some(rid) = chart_rid {
                    if let Some(rel_target) = rels.get(&rid) {
                        let chart_path = resolve_path(slide_dir, rel_target);
                        if let Ok(chart_xml) = read_zip_str(zip, &chart_path) {
                            let chart_opt = if uri.contains("chartex") || uri.contains("chartEx") {
                                parse_chartex(&chart_xml, theme)
                            } else {
                                parse_legacy_chart(&chart_xml, theme)
                            };
                            if let Some(mut chart) = chart_opt {
                                chart.x = t.x;
                                chart.y = t.y;
                                chart.width = t.cx;
                                chart.height = t.cy;
                                out.push(SlideElement::Chart(chart));
                            }
                        }
                    }
                }
            }
        }
        "grpSp" => {
            let grp_sp_pr = child(node, "grpSpPr");
            let gt: Option<GroupTransform> =
                grp_sp_pr.and_then(|pr| child(pr, "xfrm")).map(|xfrm| {
                    let off = child(xfrm, "off");
                    let ext = child(xfrm, "ext");
                    let ch_off = child(xfrm, "chOff");
                    let ch_ext = child(xfrm, "chExt");
                    GroupTransform {
                        x: off.and_then(|n| attr_i64(&n, "x")).unwrap_or(0),
                        y: off.and_then(|n| attr_i64(&n, "y")).unwrap_or(0),
                        cx: ext.and_then(|n| attr_i64(&n, "cx")).unwrap_or(0),
                        cy: ext.and_then(|n| attr_i64(&n, "cy")).unwrap_or(0),
                        ch_x: ch_off.and_then(|n| attr_i64(&n, "x")).unwrap_or(0),
                        ch_y: ch_off.and_then(|n| attr_i64(&n, "y")).unwrap_or(0),
                        ch_cx: ch_ext.and_then(|n| attr_i64(&n, "cx")).unwrap_or(0),
                        ch_cy: ch_ext.and_then(|n| attr_i64(&n, "cy")).unwrap_or(0),
                        flip_h: attr(&xfrm, "flipH")
                            .map(|v| v == "1" || v == "true")
                            .unwrap_or(false),
                        flip_v: attr(&xfrm, "flipV")
                            .map(|v| v == "1" || v == "true")
                            .unwrap_or(false),
                        rot: attr_f64(&xfrm, "rot").unwrap_or(0.0) / 60000.0,
                    }
                });

            // Determine the fill to propagate to child shapes that use grpFill.
            // - Group has solidFill/noFill → use that as child group fill
            // - Group has grpFill → inherit from parent
            // - Group has no fill → inherit from parent
            let grp_has_grp_fill = grp_sp_pr.and_then(|pr| child(pr, "grpFill")).is_some();
            let grp_explicit_fill = grp_sp_pr.and_then(|pr| parse_fill(pr, theme));
            let child_group_fill: Option<Fill> = if grp_has_grp_fill {
                group_fill.cloned()
            } else if let Some(f) = grp_explicit_fill {
                Some(f)
            } else {
                group_fill.cloned()
            };

            let start = out.len();
            for child_node in node.children().filter(|n| n.is_element()) {
                parse_sp_tree_node(
                    child_node,
                    lph,
                    slide_dir,
                    rels,
                    smartart_drawings,
                    zip,
                    theme,
                    out,
                    skip_placeholders,
                    child_group_fill.as_ref(),
                );
            }
            if let Some(gt) = gt {
                for el in &mut out[start..] {
                    apply_group_transform_to_element(el, &gt);
                }
            }
        }
        "cxnSp" => {
            // Connector shape: parse as a line/shape element
            if skip_placeholders && is_placeholder(node) {
                return;
            }
            if let Some(shape) = parse_connector(node, theme) {
                out.push(SlideElement::Shape(shape));
            }
        }
        _ => {}
    }
}

/// Render a SmartArt diagram's pre-baked fallback drawing (drawing1.xml).
/// The drawing file stores a standard dsp:spTree whose shapes share the same
/// element shape as a regular p:sp / p:cxnSp / p:grpSp (children use the `a:`
/// namespace). Coordinates inside the drawing are local to the enclosing
/// graphicFrame, so we translate every emitted element by the graphicFrame's xfrm.
/// Read a SmartArt shape's explicit text frame from `<dsp:txXfrm>`
/// (`<a:off>` / `<a:ext>`), in the drawing's local EMU coordinates. Returns
/// `None` when the shape has no txXfrm (most ordinary shapes). The optional
/// `rot` attribute is intentionally not consumed here — none of the supported
/// SmartArt layouts rotate the text frame independently of the shape, and the
/// renderer already rotates text by the shape's own rotation.
fn parse_tx_xfrm(sp_node: roxmltree::Node<'_, '_>) -> Option<TextRect> {
    let txx = child(sp_node, "txXfrm")?;
    let off = child(txx, "off")?;
    let ext = child(txx, "ext")?;
    Some(TextRect {
        x: attr_i64(&off, "x")?,
        y: attr_i64(&off, "y")?,
        width: attr_i64(&ext, "cx")?,
        height: attr_i64(&ext, "cy")?,
    })
}

fn parse_smartart_drawing(
    drawing_xml: &str,
    gf_xfrm: &Transform,
    theme: &HashMap<String, String>,
    out: &mut Vec<SlideElement>,
    zip: &mut PptxZip,
) {
    let doc = match roxmltree::Document::parse(drawing_xml) {
        Ok(d) => d,
        Err(_) => return,
    };
    let root = doc.root_element();
    let Some(sp_tree) = child(root, "spTree") else {
        return;
    };

    let empty_lph = LayoutPlaceholders::default();

    let start = out.len();
    for node in sp_tree.children().filter(|n| n.is_element()) {
        // dsp:sp / dsp:cxnSp / dsp:grpSp share local names with their p:* counterparts.
        // Shapes inside drawing1.xml carry no `a:blip r:embed` picture fills, so
        // `rels` is an empty stub. `zip` is still threaded through: a dsp shape's
        // text body may declare a `<a:buBlip>` bullet whose part existence must be
        // verified (it resolves against the empty `rels`, so it always falls
        // through to Bullet::Inherit, but the resolver still needs the archive).
        let empty_rels: HashMap<String, String> = HashMap::new();
        let empty_smartart: HashMap<String, String> = HashMap::new();
        // Dispatch the relevant branches (sp/cxnSp/grpSp) directly rather than via
        // parse_sp_tree_node, since the dsp subtree never contains embedded pictures.
        match node.tag_name().name() {
            "sp" => {
                if let Some(mut shape) =
                    parse_shape(node, &empty_lph, theme, &empty_rels, None, zip)
                {
                    shape.text_rect = parse_tx_xfrm(node);
                    out.push(SlideElement::Shape(shape));
                }
            }
            "cxnSp" => {
                if let Some(shape) = parse_connector(node, theme) {
                    out.push(SlideElement::Shape(shape));
                }
            }
            "grpSp" => {
                // Recursively render a group by collecting its children into a temp buffer
                // and applying the group's xfrm, mirroring the grpSp branch in parse_sp_tree_node.
                let grp_sp_pr = child(node, "grpSpPr");
                let gt: Option<GroupTransform> =
                    grp_sp_pr.and_then(|pr| child(pr, "xfrm")).map(|xfrm| {
                        let off = child(xfrm, "off");
                        let ext = child(xfrm, "ext");
                        let ch_off = child(xfrm, "chOff");
                        let ch_ext = child(xfrm, "chExt");
                        GroupTransform {
                            x: off.and_then(|n| attr_i64(&n, "x")).unwrap_or(0),
                            y: off.and_then(|n| attr_i64(&n, "y")).unwrap_or(0),
                            cx: ext.and_then(|n| attr_i64(&n, "cx")).unwrap_or(0),
                            cy: ext.and_then(|n| attr_i64(&n, "cy")).unwrap_or(0),
                            ch_x: ch_off.and_then(|n| attr_i64(&n, "x")).unwrap_or(0),
                            ch_y: ch_off.and_then(|n| attr_i64(&n, "y")).unwrap_or(0),
                            ch_cx: ch_ext.and_then(|n| attr_i64(&n, "cx")).unwrap_or(0),
                            ch_cy: ch_ext.and_then(|n| attr_i64(&n, "cy")).unwrap_or(0),
                            flip_h: attr(&xfrm, "flipH")
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false),
                            flip_v: attr(&xfrm, "flipV")
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false),
                            rot: attr_f64(&xfrm, "rot").unwrap_or(0.0) / 60000.0,
                        }
                    });
                let group_start = out.len();
                let _ = (&empty_rels, &empty_smartart); // silence unused warnings when no nested picture
                for child_node in node.children().filter(|n| n.is_element()) {
                    match child_node.tag_name().name() {
                        "sp" => {
                            if let Some(mut shape) =
                                parse_shape(child_node, &empty_lph, theme, &empty_rels, None, zip)
                            {
                                shape.text_rect = parse_tx_xfrm(child_node);
                                out.push(SlideElement::Shape(shape));
                            }
                        }
                        "cxnSp" => {
                            if let Some(shape) = parse_connector(child_node, theme) {
                                out.push(SlideElement::Shape(shape));
                            }
                        }
                        _ => {}
                    }
                }
                if let Some(gt) = gt {
                    for el in &mut out[group_start..] {
                        apply_group_transform_to_element(el, &gt);
                    }
                }
            }
            _ => {}
        }
    }

    // Translate all emitted elements by the graphicFrame's position.
    for el in &mut out[start..] {
        offset_slide_element(el, gf_xfrm.x, gf_xfrm.y);
    }
}

fn offset_slide_element(el: &mut SlideElement, dx: i64, dy: i64) {
    match el {
        SlideElement::Shape(s) => {
            s.x += dx;
            s.y += dy;
            if let Some(tr) = &mut s.text_rect {
                tr.x += dx;
                tr.y += dy;
            }
        }
        SlideElement::Picture(p) => {
            p.x += dx;
            p.y += dy;
        }
        SlideElement::Table(t) => {
            t.x += dx;
            t.y += dy;
        }
        SlideElement::Chart(c) => {
            c.x += dx;
            c.y += dy;
        }
        SlideElement::Media(m) => {
            m.x += dx;
            m.y += dy;
        }
    }
}

/// Parse a connector shape (p:cxnSp) as a ShapeElement with line geometry.
fn parse_connector(
    node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<ShapeElement> {
    let sp_pr = child(node, "spPr")?;
    let xfrm = child(sp_pr, "xfrm")?;
    let t = parse_xfrm(xfrm);
    if t.cx == 0 && t.cy == 0 {
        return None;
    }
    let (cnv_id, cnv_name) = read_cnv_pr(node);

    // Style-based stroke fallback. `p:style > a:lnRef idx="N"` references the
    // theme fmtScheme lnStyleLst entry N (1-based). We look up the canonical
    // width from the theme map ("+lnRef-N") rather than hardcoding 9525.
    let style_node = child(node, "style");
    let style_stroke: Option<Stroke> = style_node.and_then(|s| child(s, "lnRef")).and_then(|lr| {
        let idx: u32 = attr(&lr, "idx").and_then(|v| v.parse().ok()).unwrap_or(1);
        if idx == 0 {
            None
        } else {
            parse_color_node(lr, theme).map(|c| {
                let width = theme
                    .get(&format!("+lnRef-{}", idx))
                    .and_then(|s| s.parse::<i64>().ok())
                    .unwrap_or(9525);
                Stroke {
                    color: c,
                    width,
                    dash_style: None,
                    head_end: None,
                    tail_end: None,
                    cmpd: None,
                }
            })
        }
    });

    // Merge <a:ln> attributes onto style_stroke: <a:ln> commonly contains only
    // arrow ends (headEnd/tailEnd) while the color/width come from lnRef.
    // parse_stroke() alone drops the whole stroke when solidFill is absent,
    // which leaves connectors invisible on slides that rely on style_ref.
    let ln_node = child(sp_pr, "ln");
    let stroke: Option<Stroke> = match ln_node {
        None => style_stroke,
        Some(ln) => {
            if child(ln, "noFill").is_some() {
                None
            } else {
                let ln_width = attr_i64(&ln, "w");
                let ln_color = child(ln, "solidFill").and_then(|n| parse_color_node(n, theme));
                let ln_dash = child(ln, "prstDash")
                    .and_then(|n| attr(&n, "val"))
                    .filter(|v| v != "solid");
                let ln_head = child(ln, "headEnd")
                    .map(parse_arrow_end)
                    .filter(|a| a.kind != "none");
                let ln_tail = child(ln, "tailEnd")
                    .map(parse_arrow_end)
                    .filter(|a| a.kind != "none");
                let ln_cmpd = attr(&ln, "cmpd").filter(|v| v != "sng");
                match (ln_color, style_stroke) {
                    (Some(c), base) => Some(Stroke {
                        color: c,
                        width: ln_width
                            .unwrap_or_else(|| base.as_ref().map(|s| s.width).unwrap_or(9525)),
                        dash_style: ln_dash
                            .or_else(|| base.as_ref().and_then(|s| s.dash_style.clone())),
                        head_end: ln_head
                            .or_else(|| base.as_ref().and_then(|s| s.head_end.clone())),
                        tail_end: ln_tail
                            .or_else(|| base.as_ref().and_then(|s| s.tail_end.clone())),
                        cmpd: ln_cmpd.or_else(|| base.as_ref().and_then(|s| s.cmpd.clone())),
                    }),
                    (None, Some(base)) => Some(Stroke {
                        color: base.color,
                        width: ln_width.unwrap_or(base.width),
                        dash_style: ln_dash.or(base.dash_style),
                        head_end: ln_head.or(base.head_end),
                        tail_end: ln_tail.or(base.tail_end),
                        cmpd: ln_cmpd.or(base.cmpd),
                    }),
                    (None, None) => None,
                }
            }
        }
    };

    let shadow = child(sp_pr, "effectLst").and_then(|n| parse_shadow(n, theme));

    let cy = if t.cy == 0 { 1 } else { t.cy };

    // Preserve the actual preset geometry (bentConnector3, curvedConnector4, …)
    // rather than collapsing every p:cxnSp to "line" — otherwise the renderer
    // can't distinguish bent/curved paths from a straight segment.
    let prst_geom_node = child(sp_pr, "prstGeom");
    let geometry = prst_geom_node
        .and_then(|n| attr(&n, "prst"))
        .unwrap_or_else(|| "line".to_owned());

    // Pull connector adjust values from avLst (e.g. bentConnector3 adj1 = bend %).
    let parse_gd_val = |gd: roxmltree::Node<'_, '_>| -> Option<f64> {
        attr(&gd, "fmla")
            .and_then(|f| f.strip_prefix("val ").map(|s| s.to_owned()))
            .and_then(|s| s.parse::<f64>().ok())
    };
    let av_node = prst_geom_node.and_then(|n| child(n, "avLst"));
    let gd_nodes: Vec<_> = av_node
        .map(|av| {
            av.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "gd")
                .collect()
        })
        .unwrap_or_default();
    let adj: Option<f64> = gd_nodes
        .iter()
        .find(|n| matches!(attr(n, "name").as_deref(), Some("adj") | Some("adj1")))
        .or_else(|| gd_nodes.first())
        .and_then(|n| parse_gd_val(*n));
    let adj2: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj2"))
        .or_else(|| gd_nodes.get(1))
        .and_then(|n| parse_gd_val(*n));
    let adj3: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj3"))
        .or_else(|| gd_nodes.get(2))
        .and_then(|n| parse_gd_val(*n));
    let adj4: Option<f64> = gd_nodes
        .iter()
        .find(|n| attr(n, "name").as_deref() == Some("adj4"))
        .or_else(|| gd_nodes.get(3))
        .and_then(|n| parse_gd_val(*n));

    Some(ShapeElement {
        x: t.x,
        y: t.y,
        width: t.cx,
        height: cy,
        rotation: t.rot,
        flip_h: t.flip_h,
        flip_v: t.flip_v,
        geometry,
        fill: None,
        stroke,
        text_body: None,
        default_text_color: None,
        cust_geom: None,
        adj,
        adj2,
        adj3,
        adj4,
        adj5: None,
        adj6: None,
        adj7: None,
        adj8: None,
        shadow,
        inner_shadow: None,
        glow: None,
        soft_edge: None,
        reflection: None,
        id: cnv_id,
        name: cnv_name,
        placeholder_type: None,
        placeholder_idx: None,
        text_rect: None,
        scene3d: parse_scene3d(sp_pr),
        sp3d: parse_sp3d(sp_pr),
    })
}
