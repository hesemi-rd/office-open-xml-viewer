//! Text body / paragraph / run parsing plus the list-style level-* helpers
//! (`LevelFontSizes` / `LevelIndents` / `LevelBullets` and their read/extract/
//! has/merge functions, shared with the master extractors in `lib.rs`).
//! Extracted verbatim from `lib.rs`. Shared XML helpers (`child`,
//! `children_vec`, `attr`, `attr_r`, `attr_i64`, `attr_f64`, `resolve_path`)
//! stay in `lib.rs`; the colour + theme helpers live in `fill` / `theme`.

use crate::fill::{parse_color_node, parse_shadow};
use crate::theme::resolve_theme_typeface;
use crate::types::*;
use crate::{attr, attr_f64, attr_i64, attr_r, child, children_vec, resolve_path, PptxZip};
use ooxml_common::blip::mime_from_ext;
use ooxml_common::math::parse_omath_nodes;
use ooxml_common::text::{parse_lnspc, SpaceLine};
use std::collections::HashMap;

/// Extract the lvl1pPr defRPr font size from a txBody node.
pub(crate) fn extract_lvl1_font_size(tx_body: roxmltree::Node<'_, '_>) -> Option<f64> {
    child(tx_body, "lstStyle")
        .and_then(|ls| child(ls, "lvl1pPr"))
        .and_then(|lp| child(lp, "defRPr"))
        .and_then(|rp| attr_f64(&rp, "sz"))
        .map(|v| v / 100.0)
}

/// Per-list-level default font sizes (pt). Index 0..=8 → lvl1pPr..lvl9pPr
/// (ECMA-376 §21.1.2.4). `None` where the level isn't specified.
pub(crate) type LevelFontSizes = [Option<f64>; 9];

/// Read `<a:lvlNpPr><a:defRPr@sz>` for levels 1..9 from a node that holds
/// `<a:lvlNpPr>` children — a txBody's `<a:lstStyle>` or a master `<p:txStyles>`
/// style node (`<p:bodyStyle>` etc.). Sizes are in pt.
pub(crate) fn read_level_font_sizes(list_style: roxmltree::Node<'_, '_>) -> LevelFontSizes {
    let mut out: LevelFontSizes = [None; 9];
    for (lvl, slot) in out.iter_mut().enumerate() {
        let tag = format!("lvl{}pPr", lvl + 1);
        *slot = list_style
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == tag)
            .and_then(|lp| child(lp, "defRPr"))
            .and_then(|rp| attr_f64(&rp, "sz"))
            .map(|v| v / 100.0);
    }
    out
}

/// Per-level default font sizes from a txBody's own `<a:lstStyle>`.
pub(crate) fn extract_level_font_sizes(tx_body: roxmltree::Node<'_, '_>) -> LevelFontSizes {
    child(tx_body, "lstStyle")
        .map(read_level_font_sizes)
        .unwrap_or([None; 9])
}

/// True when any level carries a size (avoids storing all-None arrays).
pub(crate) fn has_any_level_size(s: &LevelFontSizes) -> bool {
    s.iter().any(|v| v.is_some())
}

/// Per-edge merge: `primary[lvl]` wins, else `fallback[lvl]`.
pub(crate) fn merge_level_sizes(
    primary: &LevelFontSizes,
    fallback: &LevelFontSizes,
) -> LevelFontSizes {
    let mut out: LevelFontSizes = [None; 9];
    for lvl in 0..9 {
        out[lvl] = primary[lvl].or(fallback[lvl]);
    }
    out
}

/// Per-list-level paragraph indents (EMU) — the `marL`/`marR`/`indent` attributes
/// of a `<a:lvlNpPr>` (ECMA-376 §21.1.2.4.13; `lvlNpPr` is a
/// `CT_TextParagraphProperties`, so these are attributes ON the level element
/// itself, exactly like a paragraph's own `<a:pPr>`). Each axis is `Option` so it
/// inherits independently: a level that sets only `marL` leaves `marR`/`indent`
/// `None` and a lower-priority tier supplies them.
#[derive(Clone, Copy, Default)]
pub(crate) struct LevelIndent {
    pub(crate) mar_l: Option<i64>,
    pub(crate) mar_r: Option<i64>,
    pub(crate) indent: Option<i64>,
}
pub(crate) type LevelIndents = [LevelIndent; 9];

/// Read `<a:lvlNpPr@marL/@marR/@indent>` (EMU) for levels 1..9 from a node that
/// holds `<a:lvlNpPr>` children — a txBody's `<a:lstStyle>` or a master
/// `<p:txStyles>` style node. Mirrors `read_level_font_sizes`, but the values are
/// attributes of the `lvlNpPr` element itself (not of a `<a:defRPr>` child).
pub(crate) fn read_level_indents(list_style: roxmltree::Node<'_, '_>) -> LevelIndents {
    let mut out: LevelIndents = Default::default();
    for (lvl, slot) in out.iter_mut().enumerate() {
        let tag = format!("lvl{}pPr", lvl + 1);
        if let Some(lp) = list_style
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == tag)
        {
            slot.mar_l = attr_i64(&lp, "marL");
            slot.mar_r = attr_i64(&lp, "marR");
            slot.indent = attr_i64(&lp, "indent");
        }
    }
    out
}

/// Per-level indents from a txBody's own `<a:lstStyle>`.
pub(crate) fn extract_level_indents(tx_body: roxmltree::Node<'_, '_>) -> LevelIndents {
    child(tx_body, "lstStyle")
        .map(read_level_indents)
        .unwrap_or_default()
}

/// True when any level carries any indent axis (avoids storing all-None arrays).
pub(crate) fn has_any_level_indent(s: &LevelIndents) -> bool {
    s.iter()
        .any(|li| li.mar_l.is_some() || li.mar_r.is_some() || li.indent.is_some())
}

/// Per-level, per-axis merge: `primary[lvl].x` wins, else `fallback[lvl].x`.
pub(crate) fn merge_level_indents(primary: &LevelIndents, fallback: &LevelIndents) -> LevelIndents {
    let mut out: LevelIndents = Default::default();
    for lvl in 0..9 {
        out[lvl].mar_l = primary[lvl].mar_l.or(fallback[lvl].mar_l);
        out[lvl].mar_r = primary[lvl].mar_r.or(fallback[lvl].mar_r);
        out[lvl].indent = primary[lvl].indent.or(fallback[lvl].indent);
    }
    out
}

/// Per-list-level bullet definitions (index 0..=8 → lvl1pPr..lvl9pPr).
/// `None` where the level's `<a:lvlNpPr>` declares no `buChar`/`buAutoNum`/`buNone`
/// (so the value is still inherited from a lower-priority style tier).
pub(crate) type LevelBullets = [Option<Bullet>; 9];

pub(crate) fn empty_level_bullets() -> LevelBullets {
    std::array::from_fn(|_| None)
}

/// True when any level carries an explicit bullet (avoids storing all-None arrays).
pub(crate) fn has_any_level_bullet(s: &LevelBullets) -> bool {
    s.iter().any(|v| v.is_some())
}

/// Read `<a:lvlNpPr>` bullets for levels 1..9 from a node holding `<a:lvlNpPr>`
/// children (a txBody `<a:lstStyle>` or a master `<p:txStyles>` style node).
/// A level resolves to `Some` only when it explicitly sets `buChar`/`buAutoNum`/
/// `buNone`; an absent bullet element stays `None` so lower tiers can supply it.
pub(crate) fn read_level_bullets<F: FnMut(&str) -> Option<String>>(
    list_style: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    resolve_blip: &mut F,
) -> LevelBullets {
    std::array::from_fn(|lvl| {
        let tag = format!("lvl{}pPr", lvl + 1);
        list_style
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == tag)
            .and_then(|lp| match parse_bullet(Some(lp), theme, resolve_blip) {
                Bullet::Inherit => None,
                b => Some(b),
            })
    })
}

/// Per-level bullets from a txBody's own `<a:lstStyle>`. `resolve_blip` resolves
/// a level's `<a:buBlip>` embed against this text body's part rels (§21.1.2.4.2).
pub(crate) fn extract_level_bullets<F: FnMut(&str) -> Option<String>>(
    tx_body: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    resolve_blip: &mut F,
) -> LevelBullets {
    child(tx_body, "lstStyle")
        .map(|ls| read_level_bullets(ls, theme, resolve_blip))
        .unwrap_or_else(empty_level_bullets)
}

/// Per-edge merge: `primary[lvl]` wins, else `fallback[lvl]`.
pub(crate) fn merge_level_bullets(primary: &LevelBullets, fallback: &LevelBullets) -> LevelBullets {
    std::array::from_fn(|lvl| primary[lvl].clone().or_else(|| fallback[lvl].clone()))
}

// ===========================
//  Text body parsing
// ===========================

/// Which `<a:objectDefaults>` slot to consult when the shape's own bodyPr
/// leaves an attribute unset. `Tx` ⇔ "text box" (slide-level
/// `<p:cNvSpPr txBox="1"/>`), which inherits from `<a:txDef>`. `Sp` ⇔
/// "regular shape with text" — table cell, placeholder, or a
/// preset-geometry shape carrying a `<p:txBody>` — which inherits from
/// `<a:spDef>`. Falling back to txDef for non-text-boxes is wrong because
/// txDef commonly carries `<a:spAutoFit/>` (PowerPoint's default for
/// freshly-inserted text boxes); applying that to e.g. a placeholder body
/// makes the whole paragraph spill horizontally instead of wrapping.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum ShapeKind {
    Tx,
    Sp,
}

// Carries the resolved master/layout/placeholder inheritance context (theme,
// rels, inherited font size, default alignment/spacing, level styles) that text
// runs need; these are inheritance inputs, not an arbitrary parameter bag.
#[allow(clippy::too_many_arguments)]
pub(crate) fn parse_text_body(
    tx_body: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
    inherited_font_size: Option<f64>,
    inherited_level_font_sizes: LevelFontSizes,
    inherited_level_indents: LevelIndents,
    inherited_level_bullets: &LevelBullets,
    inherited_bold: Option<bool>,
    inherited_italic: Option<bool>,
    inherited_caps: Option<String>,
    inherited_anchor: Option<String>,
    inherited_alignment: Option<String>,
    inherited_ea_ln_brk: Option<bool>,
    inherited_space_before: Option<i64>,
    inherited_space_after: Option<i64>,
    inherited_line_spacing: Option<f64>,
    shape_kind: ShapeKind,
    zip: &mut PptxZip,
) -> TextBody {
    let body_pr = child(tx_body, "bodyPr");
    // ECMA-376 §20.1.6.7 objectDefaults. The theme-level `<a:txDef>` (and
    // `<a:spDef>` as a secondary fallback) provides defaults for every
    // bodyPr attribute the slide-level shape leaves unset. Without this
    // fallback chain, sample-2's "20代" (txDef carries `<a:spAutoFit/>`)
    // and similar templates would silently use the spec's literal defaults
    // instead of what the theme author intended.
    // Shape-kind-aware lookup: text boxes consult txDef, regular shapes spDef.
    // Cross-fall is intentionally NOT done — see ShapeKind doc.
    let def_prefix = match shape_kind {
        ShapeKind::Tx => "+txDef",
        ShapeKind::Sp => "+spDef",
    };
    let theme_default_str =
        |key: &str| -> Option<String> { theme.get(&format!("{def_prefix}-bodyPr-{key}")).cloned() };
    let theme_default_i64 =
        |key: &str| -> Option<i64> { theme_default_str(key).and_then(|v| v.parse::<i64>().ok()) };
    let theme_default_u32 =
        |key: &str| -> Option<u32> { theme_default_str(key).and_then(|v| v.parse::<u32>().ok()) };
    let theme_auto_fit =
        || -> Option<String> { theme.get(&format!("{def_prefix}-autoFit")).cloned() };

    // Shared `<a:bodyPr>` grammar (anchor / wrap / vert / insets / autofit) via
    // ooxml_common::text::parse_body_pr. pptx's inheritance + theme
    // objectDefaults resolution is pre-baked into the defaults: each field is
    // `inherited?.or(theme objectDefault)?.or(spec default)`, and parse_body_pr
    // then applies the shape's own bodyPr attribute over it — so the effective
    // precedence (shape attr → inherited → theme → spec) is unchanged. When the
    // shape has no `<a:bodyPr>` at all, the resolved defaults ARE the result.
    //
    // Insets: OOXML defaults lIns=rIns=91440, tIns=bIns=45720 (the shared
    // ooxml_common::text::DEFAULT_INS_* constants, via BodyPrDefaults::spec()).
    // Autofit child (spAutoFit / normAutofit): when absent, defer to theme txDef
    // (auto_fit default below); a normAutofit also captures PowerPoint's stored
    // fontScale / lnSpcReduction (ECMA-376 §21.1.2.1.3, 62500 → 0.625).
    let spec = ooxml_common::text::BodyPrDefaults::spec();
    let body_pr_defaults = ooxml_common::text::BodyPrDefaults {
        anchor: inherited_anchor
            .or_else(|| theme_default_str("anchor"))
            .unwrap_or(spec.anchor),
        wrap: theme_default_str("wrap").unwrap_or(spec.wrap),
        vert: theme_default_str("vert").unwrap_or(spec.vert),
        l_ins: theme_default_i64("lIns").unwrap_or(spec.l_ins),
        t_ins: theme_default_i64("tIns").unwrap_or(spec.t_ins),
        r_ins: theme_default_i64("rIns").unwrap_or(spec.r_ins),
        b_ins: theme_default_i64("bIns").unwrap_or(spec.b_ins),
        auto_fit: theme_auto_fit().unwrap_or(spec.auto_fit),
    };
    let body = match body_pr {
        Some(n) => ooxml_common::text::parse_body_pr(n, &body_pr_defaults),
        // No <a:bodyPr>: every field resolves to its default.
        None => ooxml_common::text::BodyPr {
            anchor: body_pr_defaults.anchor.clone(),
            wrap: body_pr_defaults.wrap.clone(),
            vert: body_pr_defaults.vert.clone(),
            l_ins: body_pr_defaults.l_ins,
            t_ins: body_pr_defaults.t_ins,
            r_ins: body_pr_defaults.r_ins,
            b_ins: body_pr_defaults.b_ins,
            auto_fit: body_pr_defaults.auto_fit.clone(),
            font_scale: None,
            ln_spc_reduction: None,
        },
    };
    let vertical_anchor = body.anchor;
    let l_ins = body.l_ins;
    let r_ins = body.r_ins;
    let t_ins = body.t_ins;
    let b_ins = body.b_ins;
    let wrap = body.wrap;
    let vert = body.vert;
    let auto_fit = body.auto_fit;
    let font_scale = body.font_scale;
    let ln_spc_reduction = body.ln_spc_reduction;
    // ECMA-376 §20.1.10.34: numCol on <a:bodyPr> tells the renderer to
    // distribute paragraphs across N columns within the shape. Default 1.
    // spcCol is the inter-column gutter in EMU (default 0). Both fall back
    // through theme objectDefaults.
    let num_col = body_pr
        .and_then(|n| attr(&n, "numCol"))
        .and_then(|v| v.parse::<u32>().ok())
        .or_else(|| theme_default_u32("numCol"))
        .filter(|&n| n >= 1)
        .unwrap_or(1);
    let spc_col = body_pr
        .and_then(|n| attr_i64(&n, "spcCol"))
        .or_else(|| theme_default_i64("spcCol"))
        .unwrap_or(0);
    // ECMA-376 §21.1.2.1.1: rtlCol on <a:bodyPr> lays out the text body's
    // columns right-to-left. xsd:boolean, so accept "1"/"true". Shape
    // attribute → theme objectDefaults → spec default (false).
    let rtl_col = body_pr
        .and_then(|n| attr(&n, "rtlCol"))
        .or_else(|| theme_default_str("rtlCol"))
        .map(|v| v == "1" || v == "true")
        .unwrap_or(false);

    // ECMA-376 §20.1.9.19 — `<a:bodyPr><a:prstTxWarp prst="…">` selects a WordArt
    // text-warp envelope (ST_TextShapeType). Its `<a:avLst>` carries `<a:gd>`
    // adjust overrides in adj1/adj2/… order (thousandths of a percent). We record
    // the preset name + adjust values; the renderer maps glyphs through the
    // matching envelope from presetTextWarpDefinitions.xml. `prst="textNoShape"`
    // means "no warp", so it is treated as absent.
    //
    // Schema note: CT_TextBodyProperties (dml-main.xsd) is an xsd:sequence whose
    // FIRST child is prstTxWarp — before EG_TextAutofit (spAutoFit/normAutofit),
    // scene3d, EG_Text3D and extLst. Real Office files always emit it in that
    // position, and PowerPoint IGNORES a prstTxWarp placed later in the
    // sequence. This name-based lookup is deliberately position-independent for
    // robustness, but any fixture/generator we author must emit the schema
    // order or PowerPoint itself will render the text un-warped.
    let text_warp = body_pr
        .and_then(|n| child(n, "prstTxWarp"))
        .and_then(|n| attr(&n, "prst"))
        .filter(|p| p != "textNoShape")
        .map(|preset| {
            let adj = body_pr
                .and_then(|n| child(n, "prstTxWarp"))
                .and_then(|w| child(w, "avLst"))
                .map(|av| {
                    av.children()
                        .filter(|c| c.is_element() && c.tag_name().name() == "gd")
                        .filter_map(|gd| {
                            // fmla is "val <n>" for avLst adjust guides.
                            attr(&gd, "fmla").and_then(|f| {
                                f.strip_prefix("val ")
                                    .and_then(|v| v.trim().parse::<i64>().ok())
                            })
                        })
                        .collect::<Vec<i64>>()
                })
                .unwrap_or_default();
            TextWarp { preset, adj }
        });

    // Own lstStyle > lvl1pPr, then fall back to layout/master inherited values
    let own_lvl1_ppr = child(tx_body, "lstStyle").and_then(|ls| child(ls, "lvl1pPr"));
    let own_def_rpr = own_lvl1_ppr.and_then(|lp| child(lp, "defRPr"));
    let default_font_size = own_def_rpr
        .and_then(|rp| attr_f64(&rp, "sz"))
        .map(|v| v / 100.0)
        .or(inherited_font_size);
    // Effective per-list-level default sizes: this shape's own lstStyle wins per
    // level, else the layout/master inherited per-level sizes. Paragraphs pick
    // their size by `lvl` so nested bullets shrink (ECMA-376 §21.1.2.4).
    let own_level_sizes = extract_level_font_sizes(tx_body);
    let effective_level_sizes = merge_level_sizes(&own_level_sizes, &inherited_level_font_sizes);
    // Effective per-list-level indents: this shape's own lstStyle wins per
    // axis/level, else the layout/master inherited per-level indents. A paragraph
    // that omits marL/marR/indent picks them by `lvl` from this cascade before
    // falling back to PowerPoint's hardcoded implicit defaults (§21.1.2.4.13).
    let own_level_indents = extract_level_indents(tx_body);
    let effective_level_indents = merge_level_indents(&own_level_indents, &inherited_level_indents);
    // Effective per-level bullets: own lstStyle wins per level, else inherited
    // layout/master. A paragraph with no explicit bullet resolves its marker (and
    // its hanging-indent defaults) from this by `lvl` (ECMA-376 §19.7.10).
    // A slide text body's own lstStyle `<a:buBlip>` resolves against the slide's
    // rels + part directory (ECMA-376 §21.1.2.4.2), same base as the slide's
    // picture fills. `parse_text_body` is only ever called for slide shapes, so
    // the part directory is always `ppt/slides`.
    let mut resolve_slide_blip = |rid: &str| -> Option<String> {
        let target = rels.get(rid)?;
        let path = resolve_path("ppt/slides", target);
        // Verify the part exists so a listed-but-missing rId yields None and the
        // bullet falls through to Bullet::Inherit (matches the variant's doc
        // comment), mirroring the slide picture-fill resolvers. `index_for_name`
        // reads the central directory only (no inflate), unlike the former
        // `read_zip_bytes` which decompressed the entry just to discard it.
        zip.index_for_name(&path)?;
        Some(path)
    };
    let own_level_bullets = extract_level_bullets(tx_body, theme, &mut resolve_slide_blip);
    let effective_level_bullets = merge_level_bullets(&own_level_bullets, inherited_level_bullets);
    let default_bold = own_def_rpr
        .and_then(|rp| attr(&rp, "b"))
        .map(|v| v == "1" || v == "true")
        .or(inherited_bold);
    let default_italic = own_def_rpr
        .and_then(|rp| attr(&rp, "i"))
        .map(|v| v == "1" || v == "true")
        .or(inherited_italic);
    // Own lstStyle > lvl1pPr > algn overrides inherited alignment
    let body_default_alignment = own_lvl1_ppr
        .and_then(|lp| attr(&lp, "algn"))
        .map(|a| a.to_string())
        .or(inherited_alignment);

    // Own lstStyle > lvl1pPr > eaLnBrk overrides inherited (ECMA-376 §21.1.2.2.7)
    let body_default_ea_ln_brk = own_lvl1_ppr
        .and_then(|lp| attr(&lp, "eaLnBrk"))
        .map(|v| v == "1" || v == "true")
        .or(inherited_ea_ln_brk);

    // Own lstStyle > lvl1pPr spacing overrides inherited
    let own_lvl1_spcbef: Option<i64> = own_lvl1_ppr
        .and_then(|lp| child(lp, "spcBef"))
        .and_then(|s| child(s, "spcPts"))
        .and_then(|s| attr_i64(&s, "val"));
    let own_lvl1_spcaft: Option<i64> = own_lvl1_ppr
        .and_then(|lp| child(lp, "spcAft"))
        .and_then(|s| child(s, "spcPts"))
        .and_then(|s| attr_i64(&s, "val"));
    let body_default_space_before = own_lvl1_spcbef.or(inherited_space_before);
    let body_default_space_after = own_lvl1_spcaft.or(inherited_space_after);

    // Own lstStyle > lvl1pPr > lnSpc overrides inherited line spacing
    let own_lvl1_line_spacing: Option<f64> = own_lvl1_ppr
        .and_then(|lp| child(lp, "lnSpc"))
        .and_then(|ls| child(ls, "spcPct"))
        .and_then(|s| attr_f64(&s, "val"));
    let body_default_line_spacing = own_lvl1_line_spacing.or(inherited_line_spacing);

    let mut paragraphs: Vec<Paragraph> = children_vec(tx_body, "p")
        .into_iter()
        .map(|p| {
            parse_paragraph(
                p,
                theme,
                rels,
                body_default_alignment.as_deref(),
                body_default_ea_ln_brk,
                body_default_space_before,
                body_default_space_after,
                body_default_line_spacing,
                &effective_level_sizes,
                &effective_level_indents,
                &effective_level_bullets,
                zip,
            )
        })
        .collect();

    // ECMA-376 §21.1.2.3.13 cap: a run inherits cap="all"/"small" from the
    // shape's own lstStyle defRPr, else from the layout/master placeholder
    // style (e.g. a template's titleStyle cap="all" upper-cases the title even
    // though the run's text is stored mixed-case). Run-level rPr/paragraph
    // defRPr already won via parse_run; fill the remainder here.
    let body_caps = own_def_rpr
        .and_then(|rp| attr(&rp, "cap"))
        .filter(|v| v == "all" || v == "small")
        .or(inherited_caps);
    if let Some(bc) = body_caps {
        for para in &mut paragraphs {
            for run in &mut para.runs {
                if let TextRun::Text(t) = run {
                    if t.caps.is_none() {
                        t.caps = Some(bc.clone());
                    }
                }
            }
        }
    }

    TextBody {
        vertical_anchor,
        paragraphs,
        default_font_size,
        default_bold,
        default_italic,
        l_ins,
        r_ins,
        t_ins,
        b_ins,
        wrap,
        vert,
        auto_fit,
        font_scale,
        ln_spc_reduction,
        num_col,
        spc_col,
        rtl_col,
        text_warp,
    }
}

/// Walk `node` for OMML math and push a `TextRun::Math` for each equation,
/// descending PowerPoint's `mc:AlternateContent` / `mc:Choice` / `a14:m`
/// wrappers (ECMA-376 §22.1; the a14 markup is from the 2010 drawing ext).
/// Find the font size (pt) of an equation from the first run property within it
/// that carries `sz`. PowerPoint puts the size on the math run's `a:rPr` (or
/// `m:rPr`) rather than the paragraph defRPr, so inline math matches the
/// surrounding text size. `sz` is in hundredths of a point (ECMA-376 §21.1.2.3.9).
pub(crate) fn math_run_size(om: roxmltree::Node<'_, '_>) -> Option<f64> {
    om.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "rPr")
        .find_map(|rpr| attr_f64(&rpr, "sz"))
        .map(|v| v / 100.0)
}

/// Equation colour: the first run-property solidFill within the equation
/// (PowerPoint puts the colour on the math run's `a:rPr`, like the size), so
/// inline math follows the surrounding text colour (e.g. a purple title).
pub(crate) fn math_run_color(
    om: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<String> {
    om.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "rPr")
        .find_map(|rpr| child(rpr, "solidFill").and_then(|sf| parse_color_node(sf, theme)))
}

pub(crate) fn push_math_runs(
    node: roxmltree::Node<'_, '_>,
    font_size: Option<f64>,
    theme: &HashMap<String, String>,
    runs: &mut Vec<TextRun>,
) {
    match node.tag_name().name() {
        "oMath" => {
            let nodes = parse_omath_nodes(node);
            if !nodes.is_empty() {
                runs.push(TextRun::Math {
                    nodes,
                    display: false,
                    font_size: math_run_size(node).or(font_size),
                    color: math_run_color(node, theme),
                });
            }
        }
        "oMathPara" => {
            for om in node
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "oMath")
            {
                let nodes = parse_omath_nodes(om);
                if !nodes.is_empty() {
                    runs.push(TextRun::Math {
                        nodes,
                        display: true,
                        font_size: math_run_size(om).or(font_size),
                        color: math_run_color(om, theme),
                    });
                }
            }
        }
        // mc:AlternateContent → take the a14 `Choice` (the live equation) and
        // ignore mc:Fallback (a rasterized picture PowerPoint emits for old apps).
        "AlternateContent" => {
            if let Some(choice) = node
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "Choice")
            {
                for c in choice.children().filter(|n| n.is_element()) {
                    push_math_runs(c, font_size, theme, runs);
                }
            }
        }
        // a14:m wrapper (local name "m") holds an m:oMathPara.
        "m" => {
            for c in node.children().filter(|n| n.is_element()) {
                push_math_runs(c, font_size, theme, runs);
            }
        }
        _ => {}
    }
}

// Same inherited paragraph/run context as parse_text_body, scoped to one <a:p>.
#[allow(clippy::too_many_arguments)]
pub(crate) fn parse_paragraph(
    p_node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
    body_default_alignment: Option<&str>,
    body_default_ea_ln_brk: Option<bool>,
    body_default_space_before: Option<i64>,
    body_default_space_after: Option<i64>,
    body_default_line_spacing: Option<f64>,
    level_font_sizes: &LevelFontSizes,
    level_indents: &LevelIndents,
    level_bullets: &LevelBullets,
    zip: &mut PptxZip,
) -> Paragraph {
    let p_pr = child(p_node, "pPr");

    // ECMA-376 §21.1.2.2.7 `<a:pPr rtl>` — right-to-left text flow. When set
    // and the paragraph has no explicit `algn`, the implicit default flips
    // from "l" to "r" (matches PowerPoint's behaviour for Arabic / Hebrew
    // slides where users typically don't author an explicit alignment).
    let rtl = p_pr
        .and_then(|n| attr(&n, "rtl"))
        .map(|v| v == "1" || v == "true")
        .unwrap_or(false);

    // ECMA-376 §21.1.2.2.7 `<a:pPr eaLnBrk>` (xsd:boolean). Paragraph's own
    // value → body/list-style → layout/master default → spec default (true).
    // Same inheritance shape as `alignment` above.
    let ea_ln_brk = p_pr
        .and_then(|n| attr(&n, "eaLnBrk"))
        .map(|v| v == "1" || v == "true")
        .or(body_default_ea_ln_brk)
        .unwrap_or(true);

    // Paragraph's own algn → body/layout/master default → "r" if rtl, else "l"
    let alignment = p_pr
        .and_then(|n| attr(&n, "algn"))
        .map(|a| a.to_string())
        .or_else(|| body_default_alignment.map(|a| a.to_string()))
        .unwrap_or_else(|| if rtl { "r".into() } else { "l".into() });
    let lvl: u32 = p_pr
        .and_then(|n| attr(&n, "lvl"))
        .and_then(|v| v.parse().ok())
        .unwrap_or(0);

    // Effective bullet: the paragraph's own
    // `<a:buChar>`/`<a:buAutoNum>`/`<a:buBlip>`/`<a:buNone>`, else the inherited
    // per-level bullet for this placeholder (ECMA-376 §19.7.10). A paragraph's
    // own `<a:buBlip>` embed resolves against the slide rels + `ppt/slides`,
    // the same base as the slide's picture fills (§21.1.2.4.2).
    let mut resolve_para_blip = |rid: &str| -> Option<String> {
        let target = rels.get(rid)?;
        let path = resolve_path("ppt/slides", target);
        // Verify the part exists so a listed-but-missing rId yields None and the
        // bullet falls through to Bullet::Inherit (matches the variant's doc
        // comment), mirroring the slide picture-fill resolvers. `index_for_name`
        // reads the central directory only (no inflate), unlike the former
        // `read_zip_bytes` which decompressed the entry just to discard it.
        zip.index_for_name(&path)?;
        Some(path)
    };
    let bullet = match parse_bullet(p_pr, theme, &mut resolve_para_blip) {
        Bullet::Inherit => level_bullets
            .get(lvl as usize)
            .and_then(|o| o.clone())
            .unwrap_or(Bullet::Inherit),
        b => b,
    };
    // A paragraph is a list item (and gets a hanging indent) when its effective
    // bullet is a char/number/picture — whether declared explicitly or
    // inherited. An inherited bullet without an inherited marL/indent reuses
    // PowerPoint's implicit list metrics, the same defaults explicit bullets
    // already use.
    let has_bullet = matches!(
        bullet,
        Bullet::Char { .. } | Bullet::AutoNum { .. } | Bullet::Blip { .. }
    );

    // marL / marR / indent resolve per axis: direct `<a:pPr>` attribute wins,
    // else the authored list-style level cascade (`level_indents`, from the
    // shape/layout/master lstStyle per ECMA-376 §21.1.2.4.13), else PowerPoint's
    // hardcoded implicit list defaults:
    //   Bullet paragraphs:  marL = (lvl+1)*342900, indent = -342900 (hanging)
    //   Plain paragraphs:   marL = lvl*457200 (matches presentation.xml defaultTextStyle)
    let level_indent = level_indents.get(lvl as usize).copied().unwrap_or_default();
    let mar_l = p_pr
        .and_then(|n| attr_i64(&n, "marL"))
        .or(level_indent.mar_l)
        .unwrap_or_else(|| {
            if has_bullet {
                (lvl as i64 + 1) * 342900
            } else {
                lvl as i64 * 457200
            }
        });
    let mar_r = p_pr
        .and_then(|n| attr_i64(&n, "marR"))
        .or(level_indent.mar_r)
        .unwrap_or(0);
    let indent = p_pr
        .and_then(|n| attr_i64(&n, "indent"))
        .or(level_indent.indent)
        .unwrap_or(if has_bullet { -342900 } else { 0 });

    let space_before = p_pr
        .and_then(|n| {
            child(n, "spcBef")
                .and_then(|s| child(s, "spcPts"))
                .and_then(|s| attr_i64(&s, "val"))
        })
        .or(body_default_space_before);
    let space_after = p_pr
        .and_then(|n| {
            child(n, "spcAft")
                .and_then(|s| child(s, "spcPts"))
                .and_then(|s| attr_i64(&s, "val"))
        })
        .or(body_default_space_after);

    let space_line = p_pr
        .and_then(|n| child(n, "lnSpc"))
        .and_then(parse_lnspc)
        .or_else(|| body_default_line_spacing.map(|v| SpaceLine::Pct { val: v }));

    // Tab stops from pPr > tabLst
    let tab_stops: Vec<TabStop> = p_pr
        .and_then(|n| child(n, "tabLst"))
        .map(|tab_lst| {
            tab_lst
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "tab")
                .filter_map(|tab| {
                    let pos = attr_i64(&tab, "pos")?;
                    let algn = attr(&tab, "algn").unwrap_or_else(|| "l".into());
                    Some(TabStop { pos, algn })
                })
                .collect()
        })
        .unwrap_or_default();

    // Paragraph-level default run properties (pPr > defRPr)
    let def_rpr = p_pr.and_then(|n| child(n, "defRPr"));
    let def_font_size = def_rpr.and_then(|n| attr_f64(&n, "sz")).map(|v| v / 100.0);
    let def_color = def_rpr
        .and_then(|n| child(n, "solidFill"))
        .and_then(|n| parse_color_node(n, theme));
    let def_bold = def_rpr
        .and_then(|n| attr(&n, "b"))
        .map(|v| v == "1" || v == "true");
    let def_italic = def_rpr
        .and_then(|n| attr(&n, "i"))
        .map(|v| v == "1" || v == "true");
    let def_font_family = def_rpr
        .and_then(|n| child(n, "latin"))
        .and_then(|n| attr(&n, "typeface"))
        .map(|tf| resolve_theme_typeface(&tf, theme));

    let mut runs = Vec::new();
    for node in p_node.children().filter(|n| n.is_element()) {
        match node.tag_name().name() {
            "r" => {
                if let Some(run) = parse_run(node, def_rpr, theme, rels) {
                    runs.push(TextRun::Text(run));
                }
            }
            "br" => runs.push(TextRun::Break),
            // OMML equations (ECMA-376 §22.1). `def_font_size` here is the
            // paragraph's defRPr size (pre-level-fallback); the renderer applies
            // its own inheritance when this is None. PowerPoint stores inline
            // math as a bare `a14:m` (local name "m") directly under `a:p`, and
            // also inside `mc:AlternateContent`; both reach push_math_runs.
            "oMath" | "oMathPara" | "AlternateContent" | "m" => {
                push_math_runs(node, def_font_size, theme, &mut runs);
            }
            // Field elements (e.g. slide number, date): parse like a run but tag the field type
            "fld" => {
                let fld_type = attr(&node, "type").unwrap_or_default().to_string();
                let text = child(node, "t")
                    .and_then(|t| t.text())
                    .unwrap_or("")
                    .to_string();
                let r_pr = child(node, "rPr");
                let font_size = r_pr.and_then(|n| attr_f64(&n, "sz")).map(|v| v / 100.0);
                let color = r_pr
                    .and_then(|n| child(n, "solidFill"))
                    .and_then(|n| parse_color_node(n, theme));
                let bold = r_pr
                    .and_then(|n| attr(&n, "b"))
                    .map(|v| v == "1" || v == "true");
                let italic = r_pr
                    .and_then(|n| attr(&n, "i"))
                    .map(|v| v == "1" || v == "true");
                let font_family = r_pr
                    .and_then(|n| child(n, "latin"))
                    .and_then(|n| attr(&n, "typeface"))
                    .map(|tf| resolve_theme_typeface(&tf, theme));
                // §21.1.2.3.4 — a field's rPr can also carry a highlight; resolve
                // it the same way as a normal run (CT_Color via the shared path).
                let highlight = r_pr
                    .and_then(|n| child(n, "highlight"))
                    .and_then(|n| parse_color_node(n, theme));
                runs.push(TextRun::Text(TextRunData {
                    text,
                    bold,
                    italic,
                    underline: false,
                    underline_style: None,
                    underline_color: None,
                    strikethrough: false,
                    strike_double: false,
                    font_size,
                    color,
                    font_family,
                    font_family_ea: None,
                    font_family_sym: None,
                    baseline: None,
                    caps: None,
                    letter_spacing: None,
                    field_type: if fld_type == "slidenum" {
                        Some("slidenum".to_string())
                    } else {
                        None
                    },
                    hyperlink: None,
                    hyperlink_action: None,
                    shadow: None,
                    outline: None,
                    highlight,
                }));
            }
            _ => {}
        }
    }

    // For paragraphs with no visible text content, use endParaRPr sz to set line height.
    // This ensures empty spacer paragraphs have the correct height (e.g. between sections).
    let end_rpr = child(p_node, "endParaRPr");
    let has_text = runs.iter().any(|r| matches!(r, TextRun::Text(_)));
    let def_font_size = def_font_size
        .or_else(|| {
            if !has_text {
                end_rpr.and_then(|n| attr_f64(&n, "sz")).map(|v| v / 100.0)
            } else {
                None
            }
        })
        // Inherited per-list-level default size, indexed by this paragraph's
        // level (ECMA-376 §21.1.2.4): a 2nd-level bullet uses lvl3pPr's smaller
        // defRPr sz, not the level-1 size. The renderer applies `def_font_size`
        // to runs that carry no explicit `sz`.
        .or_else(|| level_font_sizes.get(lvl as usize).copied().flatten());

    Paragraph {
        alignment,
        mar_l,
        mar_r,
        indent,
        space_before,
        space_after,
        space_line,
        lvl,
        bullet,
        def_font_size,
        def_color,
        def_bold,
        def_italic,
        def_font_family,
        tab_stops,
        rtl,
        ea_ln_brk,
        runs,
    }
}

/// Parse bullet specification from pPr node.
///
/// `resolve_blip` maps a `<a:buBlip><a:blip r:embed>` rId to the bullet image's
/// embedded **zip path** (ECMA-376 §21.1.2.4.2), using the rels + part directory
/// of whichever tier this `pPr` belongs to (slide paragraph / txBody lstStyle /
/// layout / master), mirroring how `parse_blip_fill` resolves image fills. A
/// `buBlip` whose embed can't be resolved (dangling rId) falls through to the
/// `Bullet::Inherit` default so a lower style tier can still supply a marker.
pub(crate) fn parse_bullet<F: FnMut(&str) -> Option<String>>(
    p_pr: Option<roxmltree::Node<'_, '_>>,
    theme: &HashMap<String, String>,
    resolve_blip: &mut F,
) -> Bullet {
    let p_pr = match p_pr {
        Some(n) => n,
        None => return Bullet::Inherit,
    };

    // Explicit "no bullet"
    if child(p_pr, "buNone").is_some() {
        return Bullet::None;
    }

    // §21.1.2.4.3 buSzPct (val in thousandths of a percent: 100000 = 100%).
    // Shared by char and picture bullets — read once.
    let size_pct = child(p_pr, "buSzPct")
        .and_then(|n| attr_f64(&n, "val"))
        .map(|v| v / 1000.0);

    // Picture bullet (buBlip) — ECMA-376 §21.1.2.4.2. The choice is mutually
    // exclusive with buChar/buAutoNum, so resolve it before the char/number
    // branches. Only emit a Blip when the embed resolves to a real part.
    if let Some(bu_blip) = child(p_pr, "buBlip") {
        if let Some(image_path) = child(bu_blip, "blip")
            .and_then(|b| attr_r(&b, "embed"))
            .and_then(|rid| resolve_blip(&rid))
        {
            let mime_type = mime_from_ext(&image_path).to_owned();
            return Bullet::Blip {
                image_path,
                mime_type,
                size_pct,
            };
        }
    }

    // Character bullet
    if let Some(bu_char) = child(p_pr, "buChar") {
        let ch = attr(&bu_char, "char").unwrap_or_else(|| "\u{2022}".into()); // •
        let color = child(p_pr, "buClr").and_then(|n| parse_color_node(n, theme));
        let font_family = child(p_pr, "buFont")
            .and_then(|n| attr(&n, "typeface"))
            .map(|tf| resolve_theme_typeface(&tf, theme));
        return Bullet::Char {
            ch,
            color,
            size_pct,
            font_family,
        };
    }

    // Auto-numbered bullet
    if let Some(bu_auto) = child(p_pr, "buAutoNum") {
        let num_type = attr(&bu_auto, "type").unwrap_or_else(|| "arabicPeriod".into());
        let start_at = attr(&bu_auto, "startAt").and_then(|v| v.parse().ok());
        return Bullet::AutoNum { num_type, start_at };
    }

    Bullet::Inherit
}

pub(crate) fn parse_run(
    r_node: roxmltree::Node<'_, '_>,
    def_rpr: Option<roxmltree::Node<'_, '_>>,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
) -> Option<TextRunData> {
    let t_node = child(r_node, "t")?;
    let text = t_node.text().unwrap_or("").to_owned();
    let r_pr = child(r_node, "rPr");

    // Attribute with rPr → defRPr fallback; None means "not set" (inherit from body/layout defaults)
    let bold = r_pr
        .and_then(|n| attr(&n, "b"))
        .or_else(|| def_rpr.and_then(|n| attr(&n, "b")))
        .map(|v| v == "1" || v == "true");
    let italic = r_pr
        .and_then(|n| attr(&n, "i"))
        .or_else(|| def_rpr.and_then(|n| attr(&n, "i")))
        .map(|v| v == "1" || v == "true");
    // ECMA-376 §21.1.2.3.16 — underline style enum: none/sng/dbl/heavy/dotted/
    // dash/dashLong/dotDash/dotDotDash/wavy plus *Heavy variants. Carry the
    // exact value through for the renderer to dispatch on; the bool stays
    // true for any non-"none" value so existing code paths keep working.
    let underline_attr = r_pr
        .and_then(|n| attr(&n, "u"))
        .or_else(|| def_rpr.and_then(|n| attr(&n, "u")));
    let underline = underline_attr
        .as_deref()
        .map(|v| v != "none")
        .unwrap_or(false);
    let underline_style = underline_attr.filter(|v| v != "none" && v != "sng");

    // ECMA-376 §21.1.2.3.20 — uFill specifies a per-underline colour that
    // overrides the text colour. uFillTx (or absence) means "follow text".
    let underline_color = r_pr
        .and_then(|n| child(n, "uFill"))
        .or_else(|| def_rpr.and_then(|n| child(n, "uFill")))
        .and_then(|n| child(n, "solidFill"))
        .and_then(|n| parse_color_node(n, theme));

    // strikethrough: "sngStrike" or "dblStrike" → true; double tracked separately
    let strike_attr = r_pr
        .and_then(|n| attr(&n, "strike"))
        .or_else(|| def_rpr.and_then(|n| attr(&n, "strike")));
    let strikethrough = strike_attr
        .as_deref()
        .map(|v| v == "sngStrike" || v == "dblStrike")
        .unwrap_or(false);
    let strike_double = strike_attr.as_deref() == Some("dblStrike");

    // ECMA-376 §21.1.2.3.13 ST_TextCapsType: "none" | "small" | "all". Treat
    // "none" as not set (no transform) so the field stays absent in JSON.
    let caps = r_pr
        .and_then(|n| attr(&n, "cap"))
        .or_else(|| def_rpr.and_then(|n| attr(&n, "cap")))
        .filter(|v| v == "small" || v == "all");

    // ECMA-376 §21.1.2.3.5 rPr @spc — letter spacing in 100ths of a point.
    // Negative values are valid (tightening). Non-zero only.
    let letter_spacing = r_pr
        .and_then(|n| attr_f64(&n, "spc"))
        .or_else(|| def_rpr.and_then(|n| attr_f64(&n, "spc")))
        .map(|v| v / 100.0)
        .filter(|v| v.abs() > f64::EPSILON);

    // sz in hundredths of a point
    let font_size = r_pr
        .and_then(|n| attr_f64(&n, "sz"))
        .or_else(|| def_rpr.and_then(|n| attr_f64(&n, "sz")))
        .map(|v| v / 100.0);

    let color = r_pr
        .and_then(|n| child(n, "solidFill"))
        .and_then(|n| parse_color_node(n, theme))
        .or_else(|| {
            def_rpr
                .and_then(|n| child(n, "solidFill"))
                .and_then(|n| parse_color_node(n, theme))
        });

    let font_family = r_pr
        .and_then(|n| child(n, "latin"))
        .and_then(|n| attr(&n, "typeface"))
        .or_else(|| {
            def_rpr
                .and_then(|n| child(n, "latin"))
                .and_then(|n| attr(&n, "typeface"))
        })
        .map(|tf| resolve_theme_typeface(&tf, theme));
    // ECMA-376 §21.1.2.3.7 — <a:ea typeface="..."/> sets a separate font for
    // East Asian glyphs (CJK). Defaults to the theme's +mn-ea slot when the
    // run doesn't specify one explicitly.
    let font_family_ea = r_pr
        .and_then(|n| child(n, "ea"))
        .and_then(|n| attr(&n, "typeface"))
        .or_else(|| {
            def_rpr
                .and_then(|n| child(n, "ea"))
                .and_then(|n| attr(&n, "typeface"))
        })
        .map(|tf| resolve_theme_typeface(&tf, theme))
        .filter(|tf| !tf.is_empty());

    // ECMA-376 §21.1.2.3.10 — <a:sym typeface="..."/> sets the font used for
    // symbol characters. PowerPoint stores those as PUA codepoints (U+F0xx).
    let font_family_sym = r_pr
        .and_then(|n| child(n, "sym"))
        .and_then(|n| attr(&n, "typeface"))
        .or_else(|| {
            def_rpr
                .and_then(|n| child(n, "sym"))
                .and_then(|n| attr(&n, "typeface"))
        })
        .map(|tf| resolve_theme_typeface(&tf, theme))
        .filter(|tf| !tf.is_empty());

    // baseline in thousandths of a point; 30000=superscript, -25000=subscript (OOXML typical)
    let baseline = r_pr
        .and_then(|n| attr(&n, "baseline"))
        .and_then(|v| v.parse::<i32>().ok())
        .filter(|&v| v != 0);

    // a:hlinkClick — hyperlink. r:id refers to the slide rels (Target = URL or
    // internal part name). Resolve immediately so the renderer doesn't need
    // access to the rels table. ECMA-376 §21.1.2.3.5 (CT_Hyperlink): the
    // optional @action holds a "ppaction://..." verb (e.g. hlinksldjump) that
    // marks the link as an INTERNAL navigation; carry it through so the TS side
    // can distinguish a slide jump from an external URL. For a slide jump the
    // rel is TargetMode=Internal, so `hyperlink` is the internal slide part.
    let hlink_click = r_pr.and_then(|n| child(n, "hlinkClick"));
    let hyperlink = hlink_click
        .and_then(|h| attr_r(&h, "id"))
        .and_then(|rid| rels.get(&rid).cloned())
        .filter(|s| !s.is_empty());
    let hyperlink_action = hlink_click
        .and_then(|h| attr(&h, "action"))
        .filter(|s| !s.is_empty());

    // ECMA-376 §20.1.8.45 — `<a:rPr><a:effectLst><a:outerShdw>` glyph drop
    // shadow. Reuse the shape-level outerShdw reader so parse semantics
    // stay identical (blurRad, dist, dir, color + alphaModFix).
    let shadow = r_pr
        .and_then(|n| child(n, "effectLst"))
        .and_then(|el| parse_shadow(el, theme));

    // ECMA-376 §20.1.2.2.24 (CT_TextOutlineEffect) — `<a:rPr><a:ln w="..">`
    // strokes each glyph outline. `<a:noFill>` inside the ln means "no
    // visible outline" — skip in that case so the renderer doesn't draw a
    // black box around every glyph. Pull color from solidFill if present.
    let outline = r_pr
        .and_then(|n| child(n, "ln"))
        .filter(|ln| child(*ln, "noFill").is_none())
        .map(|ln| TextOutline {
            width: attr_i64(&ln, "w").unwrap_or(0),
            color: child(ln, "solidFill").and_then(|n| parse_color_node(n, theme)),
        });

    // ECMA-376 §21.1.2.3.4 — `<a:rPr><a:highlight>` text highlight (marker).
    // The element IS a CT_Color, so pass the <a:highlight> node straight to the
    // shared colour resolver — the same one solidFill uses — which walks its
    // srgbClr / schemeClr / sysClr / prstClr child and applies any tint / alpha
    // transforms. schemeClr therefore resolves through the master clrMap +
    // theme exactly like other run colours. Falls back to defRPr when the run
    // itself doesn't set a highlight.
    let highlight = r_pr
        .and_then(|n| child(n, "highlight"))
        .and_then(|n| parse_color_node(n, theme))
        .or_else(|| {
            def_rpr
                .and_then(|n| child(n, "highlight"))
                .and_then(|n| parse_color_node(n, theme))
        });

    Some(TextRunData {
        text,
        bold,
        italic,
        underline,
        underline_style,
        underline_color,
        strikethrough,
        strike_double,
        font_size,
        color,
        font_family,
        font_family_ea,
        font_family_sym,
        baseline,
        caps,
        letter_spacing,
        field_type: None,
        hyperlink,
        hyperlink_action,
        shadow,
        outline,
        highlight,
    })
}
