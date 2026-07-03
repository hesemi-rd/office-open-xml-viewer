//! DrawingML theme parsing (`<a:theme>`) shared by the docx, pptx and xlsx
//! parsers.
//!
//! Every OOXML host embeds the same theme grammar (ECMA-376 §20.1.6 /
//! §14.2.7 / §20.1.4.1): a `<a:clrScheme>` of twelve named color slots, a
//! `<a:fontScheme>` with major/minor typefaces per script, and a
//! `<a:fmtScheme><a:lnStyleLst>` of reference line widths. The three parsers had
//! three partial, drifting copies — pptx resolved `<a:prstClr>` preset names
//! while docx and xlsx silently dropped them; xlsx never read the font scheme;
//! docx never read line styles. Consolidating the *parse* here fixes the prstClr
//! gap uniformly (a preset color now resolves in all three formats) while each
//! parser keeps its own thin key-format adapter and color casing.
//!
//! Scope is "types + parse + pure predicate". This module reads the theme XML
//! into owned structs and stores each color slot's hex **exactly as authored**
//! (the `srgbClr@val` / `sysClr@lastClr` string verbatim, or a preset's
//! canonical hex). Case-folding, `#` prefixing, logical-name (`clrMap`)
//! resolution and the runtime color transforms (lumMod/tint/…) are NOT here —
//! they diverge per host and stay in each parser / the renderer.

use crate::ns::is_a_ns;
use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;

/// The twelve `<a:clrScheme>` slot names in ECMA-376 §20.1.6.2 declaration
/// order: `dk1`, `lt1`, `dk2`, `lt2`, `accent1`..`accent6`, `hlink`,
/// `folHlink`. Exposed so a consumer that needs positional order (xlsx indexes
/// its palette by slot ordinal) can iterate without hard-coding the list.
pub const CLR_SCHEME_SLOTS: [&str; 12] = [
    "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
];

/// The parsed `<a:clrScheme>`: slot name → hex string, stored **raw** (as
/// authored — no `#`, no case-folding). Keys are the twelve slot names present
/// in the theme; a slot whose color could not be read (unsupported child, or a
/// `prstClr` name outside [`preset_color`]) is simply absent.
///
/// A [`BTreeMap`] backs it so any serialized form is deterministic; lookups by
/// slot name are the common case ([`ThemeColorScheme::get`]).
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct ThemeColorScheme {
    colors: BTreeMap<String, String>,
}

impl ThemeColorScheme {
    /// Parse the `<a:clrScheme>` from a theme XML document. Each slot child holds
    /// one color element; `srgbClr` contributes its `val`, `sysClr` its
    /// `lastClr`, and `prstClr` its [`preset_color`] hex. The stored string is
    /// verbatim (the caller applies any casing / `#` prefix). Missing scheme or
    /// malformed XML yields an empty scheme.
    pub fn parse(xml: &str) -> Self {
        let mut colors = BTreeMap::new();
        let Ok(doc) = roxmltree::Document::parse(xml) else {
            return Self { colors };
        };
        let Some(scheme) = doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "clrScheme")
        else {
            return Self { colors };
        };
        for slot in scheme.children().filter(|n| n.is_element()) {
            let name = slot.tag_name().name().to_owned();
            for c in slot.children().filter(|n| n.is_element()) {
                let hex = color_node_hex(c);
                if let Some(h) = hex {
                    colors.insert(name, h);
                    break;
                }
            }
        }
        Self { colors }
    }

    /// Look up a slot's raw hex by name (`dk1`, `accent1`, `hlink`, …).
    pub fn get(&self, slot: &str) -> Option<&str> {
        self.colors.get(slot).map(String::as_str)
    }

    /// The twelve slots in spec order (`CLR_SCHEME_SLOTS`), each `Some(raw hex)`
    /// when present. Lets a positional consumer (xlsx builds a `Vec` indexed by
    /// ordinal) reconstruct its ordered palette while dropping absent slots as it
    /// sees fit.
    pub fn slots_in_order(&self) -> [Option<&str>; 12] {
        CLR_SCHEME_SLOTS.map(|slot| self.get(slot))
    }

    /// Iterate slot-name → raw-hex pairs (sorted by slot name). For a host that
    /// stores the palette in its own string-keyed map.
    pub fn iter(&self) -> impl Iterator<Item = (&str, &str)> {
        self.colors.iter().map(|(k, v)| (k.as_str(), v.as_str()))
    }

    /// True when no slot color was parsed.
    pub fn is_empty(&self) -> bool {
        self.colors.is_empty()
    }
}

/// Resolve a single DrawingML color element to its raw hex, covering the three
/// forms a `<a:clrScheme>` slot uses: `srgbClr@val`, `sysClr@lastClr`, and
/// `prstClr@val` (via [`preset_color`]). Other elements (e.g. `scheme`-relative
/// colors, which never appear inside a theme's own scheme) yield `None`. The
/// returned string is verbatim — no casing or `#` is applied.
fn color_node_hex(node: roxmltree::Node<'_, '_>) -> Option<String> {
    match node.tag_name().name() {
        "srgbClr" => node.attribute("val").map(str::to_owned),
        "sysClr" => node.attribute("lastClr").map(str::to_owned),
        "prstClr" => preset_color(node.attribute("val").unwrap_or_default()),
        _ => None,
    }
}

/// The parsed `<a:fontScheme>`: the major (heading) and minor (body) typeface
/// for each script axis. Stored as owned strings; a script with no `typeface`
/// (or an empty one) is `None`. Each parser maps these onto its own key format
/// (pptx `+mj-lt`, docx `major/latin`, …) in its adapter.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ThemeFonts {
    /// `<a:majorFont>` typefaces (latin / ea / cs).
    pub major: ThemeFontGroup,
    /// `<a:minorFont>` typefaces (latin / ea / cs).
    pub minor: ThemeFontGroup,
}

/// The three script typefaces of one font group (`<a:majorFont>` or
/// `<a:minorFont>`): Latin (`<a:latin>`), East-Asian (`<a:ea>`) and
/// complex-script (`<a:cs>`). Empty `typeface=""` (common for `ea`/`cs`) is
/// normalized to `None`.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct ThemeFontGroup {
    pub latin: Option<String>,
    pub ea: Option<String>,
    pub cs: Option<String>,
}

impl ThemeFonts {
    /// Parse the `<a:fontScheme>` (major + minor × latin/ea/cs). Missing scheme
    /// or malformed XML yields all-`None`.
    pub fn parse(xml: &str) -> Self {
        let Ok(doc) = roxmltree::Document::parse(xml) else {
            return Self::default();
        };
        let Some(scheme) = doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "fontScheme")
        else {
            return Self::default();
        };
        Self {
            major: parse_font_group(scheme, "majorFont"),
            minor: parse_font_group(scheme, "minorFont"),
        }
    }
}

/// Read one `<a:majorFont>` / `<a:minorFont>` child's latin/ea/cs typefaces.
fn parse_font_group(scheme: roxmltree::Node<'_, '_>, group_name: &str) -> ThemeFontGroup {
    let Some(group) = scheme
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == group_name)
    else {
        return ThemeFontGroup::default();
    };
    let read = |axis: &str| -> Option<String> {
        group
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == axis)
            .and_then(|n| n.attribute("typeface"))
            .filter(|t| !t.is_empty())
            .map(str::to_owned)
    };
    ThemeFontGroup {
        latin: read("latin"),
        ea: read("ea"),
        cs: read("cs"),
    }
}

/// Parse `<a:fmtScheme><a:lnStyleLst>` line-reference widths (EMU), in
/// declaration order. A drawing shape's `<a:style><a:lnRef idx="N">` resolves
/// its outline width from entry N (1-based) of this list (ECMA-376 §20.1.4.2.19);
/// an `<a:ln>` without an explicit `w` uses the CT_LineProperties default 9525
/// EMU = 0.75 pt (§20.1.2.2.24). Missing scheme or malformed XML yields an empty
/// list.
pub fn parse_ln_style_widths(xml: &str) -> Vec<i64> {
    let Ok(doc) = roxmltree::Document::parse(xml) else {
        return Vec::new();
    };
    for node in doc.descendants() {
        if node.tag_name().name() == "lnStyleLst" && is_a_ns(node.tag_name().namespace()) {
            return node
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "ln")
                .map(|ln| {
                    ln.attribute("w")
                        .and_then(|s| s.parse().ok())
                        .unwrap_or(9525)
                })
                .collect();
        }
    }
    Vec::new()
}

/// Resolve a DrawingML `<a:prstClr>` preset color name (ECMA-376 §20.1.10.48
/// ST_PresetColorVal) to its canonical uppercase hex. Covers the presets that
/// appear in real theme scheme slots; unrecognized names yield `None` (the
/// caller leaves the slot unset). Single source of truth for the three parsers,
/// which previously either handled a subset (pptx) or none (docx/xlsx).
pub fn preset_color(name: &str) -> Option<String> {
    let hex = match name {
        "black" => "000000",
        "white" => "FFFFFF",
        "red" => "FF0000",
        "green" => "008000",
        "blue" => "0000FF",
        "yellow" => "FFFF00",
        "cyan" => "00FFFF",
        "magenta" => "FF00FF",
        "orange" => "FFA500",
        "gray" | "grey" => "808080",
        "darkGray" | "darkGrey" => "404040",
        "lightGray" | "lightGrey" => "D3D3D3",
        _ => return None,
    };
    Some(hex.to_owned())
}

#[cfg(test)]
mod tests {
    use super::*;

    const THEME: &str = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:themeElements>
        <a:clrScheme name="Office">
          <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
          <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
          <a:dk2><a:srgbClr val="44546A"/></a:dk2>
          <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
          <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
          <a:accent2><a:prstClr val="orange"/></a:accent2>
          <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
          <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
        </a:clrScheme>
        <a:fontScheme name="Office">
          <a:majorFont>
            <a:latin typeface="Aptos Display"/>
            <a:ea typeface="Yu Gothic"/>
            <a:cs typeface=""/>
          </a:majorFont>
          <a:minorFont>
            <a:latin typeface="Aptos"/>
            <a:ea typeface=""/>
            <a:cs typeface=""/>
          </a:minorFont>
        </a:fontScheme>
        <a:fmtScheme name="Office">
          <a:lnStyleLst>
            <a:ln w="6350"/>
            <a:ln w="12700"/>
            <a:ln/>
          </a:lnStyleLst>
        </a:fmtScheme>
      </a:themeElements>
    </a:theme>"#;

    #[test]
    fn clr_scheme_reads_srgb_sys_and_preset_raw() {
        let s = ThemeColorScheme::parse(THEME);
        // sysClr uses lastClr, srgbClr uses val — both raw (as authored).
        assert_eq!(s.get("dk1"), Some("000000"));
        assert_eq!(s.get("lt1"), Some("FFFFFF"));
        assert_eq!(s.get("dk2"), Some("44546A"));
        assert_eq!(s.get("accent1"), Some("4472C4"));
        // prstClr resolves through preset_color (uniform across formats).
        assert_eq!(s.get("accent2"), Some("FFA500"));
        assert_eq!(s.get("hlink"), Some("0563C1"));
        // Slots not present in the XML are absent (not empty strings).
        assert_eq!(s.get("accent3"), None);
    }

    #[test]
    fn clr_scheme_slots_in_order_matches_spec_positions() {
        let s = ThemeColorScheme::parse(THEME);
        let ordered = s.slots_in_order();
        assert_eq!(ordered[0], Some("000000")); // dk1
        assert_eq!(ordered[1], Some("FFFFFF")); // lt1
        assert_eq!(ordered[4], Some("4472C4")); // accent1
        assert_eq!(ordered[5], Some("FFA500")); // accent2 (preset)
        assert_eq!(ordered[6], None); // accent3 absent
        assert_eq!(ordered[10], Some("0563C1")); // hlink
    }

    #[test]
    fn font_scheme_reads_axes_and_drops_empty() {
        let f = ThemeFonts::parse(THEME);
        assert_eq!(f.major.latin.as_deref(), Some("Aptos Display"));
        assert_eq!(f.major.ea.as_deref(), Some("Yu Gothic"));
        // Empty typeface="" normalizes to None (not Some("")).
        assert_eq!(f.major.cs, None);
        assert_eq!(f.minor.latin.as_deref(), Some("Aptos"));
        assert_eq!(f.minor.ea, None);
        assert_eq!(f.minor.cs, None);
    }

    #[test]
    fn ln_style_widths_reads_list_with_default() {
        // Third <a:ln> has no w → CT_LineProperties default 9525.
        assert_eq!(parse_ln_style_widths(THEME), vec![6350, 12700, 9525]);
    }

    /// Same fixture as [`ln_style_widths_reads_list_with_default`], but declared
    /// under the ISO/IEC 29500 Strict `a:` URI
    /// (`http://purl.oclc.org/ooxml/drawingml/main`) instead of the Transitional
    /// one. `parse_ln_style_widths` must accept both — a document saved by
    /// Office in Strict conformance still has a `<a:fmtScheme><a:lnStyleLst>` to
    /// resolve `<a:lnRef idx="N">` line widths from.
    #[test]
    fn ln_style_widths_reads_list_with_default_strict_ns() {
        const STRICT_THEME: &str = r#"<a:theme xmlns:a="http://purl.oclc.org/ooxml/drawingml/main">
          <a:themeElements>
            <a:fmtScheme name="Office">
              <a:lnStyleLst>
                <a:ln w="6350"/>
                <a:ln w="12700"/>
                <a:ln/>
              </a:lnStyleLst>
            </a:fmtScheme>
          </a:themeElements>
        </a:theme>"#;
        assert_eq!(parse_ln_style_widths(STRICT_THEME), vec![6350, 12700, 9525]);
    }

    #[test]
    fn preset_color_covers_named_presets() {
        assert_eq!(preset_color("black").as_deref(), Some("000000"));
        assert_eq!(preset_color("white").as_deref(), Some("FFFFFF"));
        assert_eq!(preset_color("gray").as_deref(), Some("808080"));
        assert_eq!(preset_color("grey").as_deref(), Some("808080"));
        assert_eq!(preset_color("orange").as_deref(), Some("FFA500"));
        // Unknown → None (caller leaves the slot unset).
        assert_eq!(preset_color("chartreuse"), None);
    }

    #[test]
    fn empty_and_malformed_yield_empty() {
        assert!(ThemeColorScheme::parse("").is_empty());
        assert!(ThemeColorScheme::parse("<not xml").is_empty());
        assert_eq!(ThemeFonts::parse(""), ThemeFonts::default());
        assert!(parse_ln_style_widths("").is_empty());
    }
}
