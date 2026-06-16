use base64::engine::general_purpose::STANDARD as B64;
use base64::Engine;
use roxmltree::Document as XmlDoc;
use std::collections::HashMap;
use std::io::Read;
use zip::ZipArchive;

use crate::numbering::NumberingMap;
use crate::styles::{
    parse_para_fmt, parse_run_fmt, CondFmt, EdgeBorder, ParaFmt, RawTblBorders, RunFmt, StyleMap,
};
use crate::types::*;
use crate::xml_util::*;

const DEFAULT_FONT_SIZE: f64 = 10.0; // pt fallback

/// OMML (math) namespace — ECMA-376 §22.
const M_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";

type Zip<'a> = ZipArchive<std::io::Cursor<&'a [u8]>>;

/// Section-level header/footer references collected from sectPr.
/// Maps reference type ("default" | "first" | "even") to the target xml path (e.g. "header1.xml").
#[derive(Default)]
struct SectionRefs {
    headers: HashMap<String, String>,
    footers: HashMap<String, String>,
}

pub fn parse(data: &[u8]) -> Result<Document, String> {
    let cursor = std::io::Cursor::new(data);
    let mut zip = ZipArchive::new(cursor).map_err(|e| e.to_string())?;

    let rels_xml = read_zip_entry(&mut zip, "word/_rels/document.xml.rels").unwrap_or_default();
    let rel_map = parse_rels(&rels_xml);

    // Styles are referenced from the document relationships (Target may be
    // "styles.xml" or "styles2.xml"). Fall back to "word/styles.xml" for old files.
    let styles_path = find_rel_target(&rels_xml, "styles")
        .map(|t| {
            if t.starts_with('/') {
                t.trim_start_matches('/').to_string()
            } else {
                format!("word/{}", t)
            }
        })
        .unwrap_or_else(|| "word/styles.xml".to_string());
    let style_map = read_zip_entry(&mut zip, &styles_path)
        .map(|s| StyleMap::parse(&s))
        .unwrap_or_else(|_| StyleMap::parse(""));

    let numbering_path = find_rel_target(&rels_xml, "numbering")
        .map(|t| {
            if t.starts_with('/') {
                t.trim_start_matches('/').to_string()
            } else {
                format!("word/{}", t)
            }
        })
        .unwrap_or_else(|| "word/numbering.xml".to_string());
    let mut num_map = read_zip_entry(&mut zip, &numbering_path)
        .map(|s| NumberingMap::parse(&s))
        .unwrap_or_default();

    // Theme is referenced by a relationship with Type ending in "/theme" — resolve
    // to word/<target> and parse the clrScheme.
    let mut theme = find_rel_target(&rels_xml, "theme")
        .map(|t| {
            let p = if t.starts_with('/') {
                t.trim_start_matches('/').to_string()
            } else {
                format!("word/{}", t)
            };
            read_zip_entry(&mut zip, &p)
                .map(|s| ThemeColors::parse(&s))
                .unwrap_or_default()
        })
        .unwrap_or_default();

    // §17.15.1.88 w:themeFontLang — when the theme leaves a cs typeface empty,
    // the settings' bidi language decides the actual complex-script face.
    let settings_path = find_rel_target(&rels_xml, "settings")
        .map(|t| {
            if t.starts_with('/') {
                t.trim_start_matches('/').to_string()
            } else {
                format!("word/{}", t)
            }
        })
        .unwrap_or_else(|| "word/settings.xml".to_string());
    let mut document_settings: Option<crate::types::DocumentSettings> = None;
    if let Ok(settings_xml) = read_zip_entry(&mut zip, &settings_path) {
        if let Some(lang) = parse_theme_font_bidi_lang(&settings_xml) {
            theme.fill_default_cs_font(&lang);
        }
        document_settings = parse_document_settings(&settings_xml);
    }
    let theme = theme;

    let media_map = load_media_map(&mut zip, &rel_map, "word/");

    let doc_xml = read_zip_entry(&mut zip, "word/document.xml")?;
    let xml_doc = XmlDoc::parse(&doc_xml).map_err(|e| e.to_string())?;

    let body_node = xml_doc
        .root_element()
        .descendants()
        .find(|n| n.tag_name().name() == "body")
        .ok_or("No body element")?;

    let sect_pr = body_node
        .children()
        .rfind(|n| n.is_element())
        .filter(|n| n.tag_name().name() == "sectPr");

    let (section, refs) = parse_section(sect_pr, &rel_map);

    let body = parse_body_elements(
        body_node,
        &style_map,
        &mut num_map,
        &media_map,
        &rel_map,
        &theme,
    );

    let headers = load_header_footer_set(
        &mut zip,
        &refs.headers,
        "hdr",
        &style_map,
        &mut num_map,
        &theme,
    );
    let footers = load_header_footer_set(
        &mut zip,
        &refs.footers,
        "ftr",
        &style_map,
        &mut num_map,
        &theme,
    );

    let major_font = theme.theme_font("major", "latin");
    let minor_font = theme.theme_font("minor", "latin");

    // ECMA-376 §17.8.3.10: font family classification from fontTable.xml.
    // Resolve via relationship (Type ending in "/fontTable"); fall back to
    // "word/fontTable.xml" for documents that omit the relationship.
    let font_table_path = find_rel_target(&rels_xml, "fontTable")
        .map(|t| {
            if t.starts_with('/') {
                t.trim_start_matches('/').to_string()
            } else {
                format!("word/{}", t)
            }
        })
        .unwrap_or_else(|| "word/fontTable.xml".to_string());
    let font_family_classes = read_zip_entry(&mut zip, &font_table_path)
        .map(|s| parse_font_table(&s))
        .unwrap_or_default();

    // Track-changes events live inline in the body XML; do a second pass to
    // surface them as a flat list with author/date metadata. The body parse
    // above already merged the run text transparently — this gives consumers
    // (MCP / agents) the revision metadata without disturbing rendering.
    let revisions = collect_revisions(body_node);

    let comments = find_rel_target(&rels_xml, "comments")
        .map(|t| {
            if t.starts_with('/') {
                t.trim_start_matches('/').to_string()
            } else {
                format!("word/{}", t)
            }
        })
        .and_then(|p| read_zip_entry(&mut zip, &p).ok())
        .map(|xml| parse_comments(&xml))
        .unwrap_or_default();
    let footnotes_path = find_rel_target(&rels_xml, "footnotes").map(|t| {
        if t.starts_with('/') {
            t.trim_start_matches('/').to_string()
        } else {
            format!("word/{}", t)
        }
    });
    let footnotes = footnotes_path
        .map(|p| parse_notes(&mut zip, &p, "footnote", &style_map, &mut num_map, &theme))
        .unwrap_or_default();
    let endnotes_path = find_rel_target(&rels_xml, "endnotes").map(|t| {
        if t.starts_with('/') {
            t.trim_start_matches('/').to_string()
        } else {
            format!("word/{}", t)
        }
    });
    let endnotes = endnotes_path
        .map(|p| parse_notes(&mut zip, &p, "endnote", &style_map, &mut num_map, &theme))
        .unwrap_or_default();

    Ok(Document {
        section,
        body,
        headers,
        footers,
        major_font,
        minor_font,
        font_family_classes,
        revisions,
        comments,
        footnotes,
        endnotes,
        settings: document_settings,
    })
}

/// Walks the body looking for `<w:ins>` / `<w:del>` ancestors and returns one
/// `DocxRevision` per element. Text is collected from descendant `<w:t>` (for
/// insertions) and `<w:delText>` (for deletions).
fn collect_revisions(body: roxmltree::Node) -> Vec<crate::types::DocxRevision> {
    let mut out = Vec::new();
    for node in body.descendants().filter(|n| n.is_element()) {
        let kind = match node.tag_name().name() {
            "ins" => "insertion",
            "del" => "deletion",
            _ => continue,
        };
        let author = node
            .attributes()
            .find(|a| a.name() == "author")
            .map(|a| a.value().to_string())
            .filter(|s| !s.is_empty());
        let date = node
            .attributes()
            .find(|a| a.name() == "date")
            .map(|a| a.value().to_string())
            .filter(|s| !s.is_empty());
        let mut text = String::new();
        for t in node.descendants().filter(|n| n.is_element()) {
            // For insertions, w:t carries the new text. For deletions, the
            // original text lives in w:delText (ECMA-376 §17.13.5.13).
            let is_text = (kind == "insertion" && t.tag_name().name() == "t")
                || (kind == "deletion" && t.tag_name().name() == "delText");
            if is_text {
                if let Some(s) = t.text() {
                    text.push_str(s);
                }
            }
        }
        out.push(crate::types::DocxRevision {
            kind: kind.to_string(),
            author,
            date,
            text,
        });
    }
    out
}

/// Parse word/comments.xml into a flat list of `<w:comment>` entries.
fn parse_comments(xml: &str) -> Vec<crate::types::DocxComment> {
    let Ok(doc) = roxmltree::Document::parse(xml) else {
        return Vec::new();
    };
    let mut out = Vec::new();
    for c in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "comment")
    {
        let id = c
            .attributes()
            .find(|a| a.name() == "id")
            .map(|a| a.value().to_string())
            .unwrap_or_default();
        if id.is_empty() {
            continue;
        }
        let author = c
            .attributes()
            .find(|a| a.name() == "author")
            .map(|a| a.value().to_string())
            .filter(|s| !s.is_empty());
        let initials = c
            .attributes()
            .find(|a| a.name() == "initials")
            .map(|a| a.value().to_string())
            .filter(|s| !s.is_empty());
        let date = c
            .attributes()
            .find(|a| a.name() == "date")
            .map(|a| a.value().to_string())
            .filter(|s| !s.is_empty());
        let mut text = String::new();
        for t in c
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "t")
        {
            if let Some(s) = t.text() {
                text.push_str(s);
            }
        }
        out.push(crate::types::DocxComment {
            id,
            author,
            initials,
            date,
            text,
        });
    }
    out
}

/// Parse word/footnotes.xml or word/endnotes.xml into a list of notes, each
/// carrying its full block-level content (ECMA-376 §17.11.2 / §17.11.10).
/// Excludes the reserved entries — separator (id="-1") and
/// continuationSeparator (id="0") — and any note declared
/// `w:type="separator" | "continuationSeparator" | "continuationNotice"`.
/// Note content is parsed with the document's styles + numbering so the
/// FootnoteText/EndnoteText style and the `<w:footnoteRef/>` auto-number marker
/// resolve correctly; the note part's own relationships supply media.
fn parse_notes(
    zip: &mut Zip,
    path: &str,
    element_name: &str,
    style_map: &StyleMap,
    num_map: &mut NumberingMap,
    theme: &ThemeColors,
) -> Vec<crate::types::DocxNote> {
    let Ok(xml) = read_zip_entry(zip, path) else {
        return Vec::new();
    };

    // Per-part rels for media (e.g. an image inside a footnote). The part lives
    // at e.g. word/footnotes.xml, so its rels are word/_rels/footnotes.xml.rels.
    let (dir, file) = path.rsplit_once('/').unwrap_or(("", path));
    let rels_path = if dir.is_empty() {
        format!("_rels/{}.rels", file)
    } else {
        format!("{}/_rels/{}.rels", dir, file)
    };
    let base_dir = if dir.is_empty() {
        String::new()
    } else {
        format!("{}/", dir)
    };
    let rels_xml = read_zip_entry(zip, &rels_path).unwrap_or_default();
    let local_rel_map = parse_rels(&rels_xml);
    let local_media_map = load_media_map(zip, &local_rel_map, &base_dir);

    let Ok(doc) = XmlDoc::parse(&xml) else {
        return Vec::new();
    };
    let mut out = Vec::new();
    for n in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == element_name)
    {
        let id = attr_w(n, "id").unwrap_or_default();
        // ECMA-376 §17.11.18 ST_FtnEdn — skip reserved special notes. They are
        // tagged `w:type` (separator / continuationSeparator / continuationNotice)
        // and conventionally use ids -1 / 0.
        let note_type = attr_w(n, "type");
        let is_special = matches!(
            note_type.as_deref(),
            Some("separator") | Some("continuationSeparator") | Some("continuationNotice")
        );
        if id.is_empty() || id == "-1" || id == "0" || is_special {
            continue;
        }
        let content = parse_body_elements(
            n,
            style_map,
            num_map,
            &local_media_map,
            &local_rel_map,
            theme,
        );
        out.push(crate::types::DocxNote { id, content });
    }
    out
}

/// Resolve scheme color names (accent1..6, dk1, dk2, lt1, lt2, hlink, folHlink)
/// to hex strings parsed from word/theme/themeN.xml clrScheme.
#[derive(Debug, Default, Clone)]
pub struct ThemeColors {
    map: HashMap<String, String>,
    /// ECMA-376 §20.1.4.1.14 fontScheme: theme-referenced typefaces used via
    /// rFonts @asciiTheme / @hAnsiTheme / @eastAsiaTheme / @cstheme.
    /// Keys: "minor" + "major" crossed with "latin"/"ea"/"cs".
    fonts: HashMap<String, String>,
    /// Raw theme XML, retained so wps:style/fillRef → fillStyleLst /
    /// bgFillStyleLst lookups (ECMA-376 §20.1.4.1.7) can be resolved on demand.
    /// Re-parsing per shape is fine — the cover usually has only a handful of
    /// shapes that take a fillRef, and theme XML is small.
    theme_xml: Option<String>,
}

impl ThemeColors {
    fn parse(xml: &str) -> Self {
        let mut map: HashMap<String, String> = HashMap::new();
        let mut fonts: HashMap<String, String> = HashMap::new();
        let theme_xml = Some(xml.to_string());
        let doc = match XmlDoc::parse(xml) {
            Ok(d) => d,
            Err(_) => {
                return Self {
                    map,
                    fonts,
                    theme_xml,
                }
            }
        };
        let root = doc.root_element();
        if let Some(scheme) = root
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "clrScheme")
        {
            for child in scheme.children().filter(|n| n.is_element()) {
                let name = child.tag_name().name().to_string();
                let hex = child.children().filter(|n| n.is_element()).find_map(|n| {
                    match n.tag_name().name() {
                        "srgbClr" => n.attribute("val").map(|v| v.to_uppercase()),
                        "sysClr" => n.attribute("lastClr").map(|v| v.to_uppercase()),
                        _ => None,
                    }
                });
                if let Some(h) = hex {
                    map.insert(name, h);
                }
            }
        }
        if let Some(font_scheme) = root
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "fontScheme")
        {
            for group_name in &["majorFont", "minorFont"] {
                let prefix = if *group_name == "majorFont" {
                    "major"
                } else {
                    "minor"
                };
                if let Some(group) = font_scheme
                    .children()
                    .find(|n| n.is_element() && n.tag_name().name() == *group_name)
                {
                    for child in group.children().filter(|n| n.is_element()) {
                        let typ = match child.tag_name().name() {
                            "latin" => Some("latin"),
                            "ea" => Some("ea"),
                            "cs" => Some("cs"),
                            _ => None,
                        };
                        if let Some(t) = typ {
                            if let Some(face) = child.attribute("typeface") {
                                if !face.is_empty() {
                                    fonts.insert(format!("{prefix}/{t}"), face.to_string());
                                }
                            }
                        }
                    }
                }
            }
        }
        Self {
            map,
            fonts,
            theme_xml,
        }
    }

    fn resolve(&self, scheme_name: &str) -> Option<String> {
        // bg1/bg2/tx1/tx2 map onto lt1/lt2/dk1/dk2 per the default §19.3.1.6
        // clrMap; raw slot names (and accents/hlink) pass through unchanged.
        // Canonical table: ooxml_common::color::SCHEME_DEFAULT_SLOTS.
        let key = ooxml_common::color::default_scheme_slot(scheme_name);
        self.map.get(key).cloned()
    }

    /// Resolve a string that may be either a literal typeface (e.g. "Georgia")
    /// or an internal "@theme:<ref>" marker produced at rFonts-parse time.
    /// Theme refs that fail to resolve fall back to None so the renderer can
    /// use its own default; literal typefaces pass through unchanged.
    pub fn resolve_font_ref(&self, v: Option<String>) -> Option<String> {
        let s = v?;
        if let Some(r) = s.strip_prefix("@theme:") {
            return self.resolve_font(r);
        }
        Some(s)
    }

    /// Read a theme typeface directly by group ("major" / "minor") and axis
    /// ("latin" / "ea" / "cs"). Used by the document loader to expose the
    /// heading / body fonts to TS callers (e.g. for Google Fonts preload).
    pub fn theme_font(&self, group: &str, axis: &str) -> Option<String> {
        self.fonts.get(&format!("{group}/{axis}")).cloned()
    }

    /// Fill the complex-script theme slots when the theme leaves them EMPTY
    /// (`<a:cs typeface=""/>`). ECMA-376 §17.15.1.88 `w:themeFontLang`: the
    /// document's bidi language selects the actual face for an unspecified
    /// theme font. Word substitutes Arial for Arabic/Hebrew here (verified
    /// against sample-7.pdf: the no-rFonts table cells embed ArialMT while
    /// explicit-rFonts runs embed Sakkal Majalla).
    pub fn fill_default_cs_font(&mut self, bidi_lang: &str) {
        let primary = bidi_lang
            .split('-')
            .next()
            .unwrap_or("")
            .to_ascii_lowercase();
        let default = match primary.as_str() {
            "ar" | "he" => "Arial",
            _ => return,
        };
        for group in ["minor", "major"] {
            let key = format!("{group}/cs");
            self.fonts.entry(key).or_insert_with(|| default.to_string());
        }
    }

    /// Resolve an rFonts theme reference (e.g. "minorHAnsi" → minor.latin,
    /// "minorEastAsia" → minor.ea, "majorHAnsi" → major.latin). Returns None
    /// when the reference is unknown or the theme has no matching typeface.
    pub fn resolve_font(&self, theme_ref: &str) -> Option<String> {
        let (group, axis) = match theme_ref {
            "minorHAnsi" | "minorAscii" => ("minor", "latin"),
            "minorBidi" => ("minor", "cs"),
            "minorEastAsia" => ("minor", "ea"),
            "majorHAnsi" | "majorAscii" => ("major", "latin"),
            "majorBidi" => ("major", "cs"),
            "majorEastAsia" => ("major", "ea"),
            _ => return None,
        };
        self.fonts.get(&format!("{group}/{axis}")).cloned()
    }
}

/// Extract `<w:themeFontLang w:bidi="…"/>` from word/settings.xml (§17.15.1.88).
fn parse_theme_font_bidi_lang(settings_xml: &str) -> Option<String> {
    let doc = XmlDoc::parse(settings_xml).ok()?;
    let node = doc
        .root_element()
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "themeFontLang")?;
    attr_w(node, "bidi").filter(|s| !s.is_empty())
}

/// Parse the typography settings the renderer needs from `word/settings.xml`.
///
/// - §17.15.1.58 `w:kinsoku` — East-Asian line-breaking toggle (ST_OnOff;
///   absence means ON, which the renderer assumes, so we only surface an
///   explicit `Some(false)`/`Some(true)`).
/// - §17.15.1.60 `w:noLineBreaksBefore` / §17.15.1.59 `w:noLineBreaksAfter` —
///   custom 行頭/行末禁則 character sets. The spec says `w:val` "specifies the
///   set of characters", i.e. a present element REPLACES the application
///   default set. A document may emit one element per `w:lang`; we concatenate
///   the `w:val` strings into one set (the renderer is language-agnostic here
///   and applies the union to its single CJK break path).
///
/// Returns `None` when none of these elements are present, so the renderer
/// falls back to its built-in Japanese defaults with kinsoku ON.
fn parse_document_settings(settings_xml: &str) -> Option<crate::types::DocumentSettings> {
    let doc = XmlDoc::parse(settings_xml).ok()?;
    let root = doc.root_element();

    let kinsoku = bool_prop(root, "kinsoku");

    let collect = |tag: &str| -> Option<String> {
        let mut found = false;
        let mut acc = String::new();
        for node in root
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == tag)
        {
            found = true;
            if let Some(val) = attr_w(node, "val") {
                acc.push_str(&val);
            }
        }
        if found {
            Some(acc)
        } else {
            None
        }
    };
    let no_line_breaks_before = collect("noLineBreaksBefore");
    let no_line_breaks_after = collect("noLineBreaksAfter");

    // ECMA-376 §22.1.2.30 `m:mathPr/m:defJc@m:val` — document-wide default math
    // justification (math namespace, bare `val` fallback).
    let math_def_jc = root
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "mathPr")
        .and_then(|mp| {
            mp.children()
                .find(|n| n.is_element() && n.tag_name().name() == "defJc")
        })
        .and_then(|jc| {
            jc.attribute((M_NS, "val"))
                .or_else(|| jc.attribute("val"))
                .map(|s| s.to_string())
        });

    if kinsoku.is_none()
        && no_line_breaks_before.is_none()
        && no_line_breaks_after.is_none()
        && math_def_jc.is_none()
    {
        return None;
    }
    Some(crate::types::DocumentSettings {
        kinsoku,
        no_line_breaks_before,
        no_line_breaks_after,
        math_def_jc,
    })
}

fn find_rel_target(rels_xml: &str, type_suffix: &str) -> Option<String> {
    if rels_xml.is_empty() {
        return None;
    }
    let doc = XmlDoc::parse(rels_xml).ok()?;
    for rel in doc
        .root_element()
        .children()
        .filter(|n| n.tag_name().name() == "Relationship")
    {
        if let (Some(ty), Some(target)) = (rel.attribute("Type"), rel.attribute("Target")) {
            if ty.ends_with(&format!("/{}", type_suffix)) {
                return Some(target.to_string());
            }
        }
    }
    None
}

/// Parse `word/fontTable.xml` and build a map from font name to family class.
///
/// ECMA-376 §17.8.3.10 defines `<w:font w:name="…"><w:family w:val="…"/></w:font>`
/// where val is one of: `roman` (serif), `swiss` (sans-serif), `modern`
/// (monospace), `script`, `decorative`, `auto` (no info). The renderer uses
/// this map as the primary source of serif/sans-serif classification, falling
/// back to name-pattern matching only when the font is absent or classified
/// as `auto`.
fn parse_font_table(xml: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let Ok(doc) = XmlDoc::parse(xml) else {
        return map;
    };
    let w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    for font in doc.root_element().descendants().filter(|n| {
        n.is_element() && n.tag_name().name() == "font" && n.tag_name().namespace() == Some(w_ns)
    }) {
        let Some(name) = font.attribute((w_ns, "name")) else {
            continue;
        };
        let family = font
            .children()
            .find(|n| {
                n.is_element()
                    && n.tag_name().name() == "family"
                    && n.tag_name().namespace() == Some(w_ns)
            })
            .and_then(|n| n.attribute((w_ns, "val")));
        if let Some(f) = family {
            map.insert(name.to_string(), f.to_string());
        }
    }
    map
}

fn parse_body_elements(
    body_node: roxmltree::Node,
    style_map: &StyleMap,
    num_map: &mut NumberingMap,
    media_map: &HashMap<String, String>,
    rel_map: &HashMap<String, String>,
    theme: &ThemeColors,
) -> Vec<BodyElement> {
    let mut body: Vec<BodyElement> = Vec::new();
    // The body-level sectPr (the last element) defines the final section and
    // is not a page break. Mid-body sectPrs (nested in pPr) DO imply a page break.
    let body_children = element_children_flat(body_node);
    let body_level_sect_pr = body_children
        .iter()
        .last()
        .copied()
        .filter(|n| n.tag_name().name() == "sectPr");
    let body_level_sect_id = body_level_sect_pr.map(|n| n.id());

    for child in body_children {
        match child.tag_name().name() {
            "p" => {
                let result =
                    parse_paragraph(child, style_map, num_map, media_map, rel_map, theme, None);
                let is_page_break_only = result.runs.len() == 1
                    && matches!(
                        result.runs[0],
                        DocRun::Break {
                            break_type: BreakType::Page
                        }
                    );
                if is_page_break_only {
                    body.push(BodyElement::PageBreak { parity: None });
                    continue;
                }
                // Mid-paragraph page breaks come from <w:br w:type="page"/>
                // and from <w:lastRenderedPageBreak/>. Split the paragraph
                // around them and emit BodyElement::PageBreak between the
                // pieces — each piece keeps the same pPr so layout
                // continues correctly on the next page.
                for piece in split_para_on_page_breaks(result) {
                    match piece {
                        ParaPiece::Para(p) => body.push(BodyElement::Paragraph(p)),
                        ParaPiece::PageBreak => body.push(BodyElement::PageBreak { parity: None }),
                    }
                }
                // ECMA-376 §17.6.1: a section break inside pPr defines the
                // section that ENDS at this paragraph. The break TYPE
                // (`<w:type w:val>`) controls how the next section starts:
                //   - "continuous" → no page break
                //   - "nextPage" / missing → plain page break
                //   - "oddPage" / "evenPage" → page break + parity padding
                if let Some(sect_pr) = child_w(child, "pPr").and_then(|ppr| child_w(ppr, "sectPr"))
                {
                    match read_section_break_type(sect_pr).as_deref() {
                        Some("continuous") => { /* no page break */ }
                        Some("oddPage") => body.push(BodyElement::PageBreak {
                            parity: Some("odd".to_string()),
                        }),
                        Some("evenPage") => body.push(BodyElement::PageBreak {
                            parity: Some("even".to_string()),
                        }),
                        _ => body.push(BodyElement::PageBreak { parity: None }),
                    }
                }
            }
            "tbl" => {
                let tbl = parse_table(child, style_map, num_map, media_map, rel_map, theme);
                body.push(BodyElement::Table(tbl));
            }
            // Mid-body loose sectPr (rare) behaves like a page break. The
            // final body-level sectPr only defines section settings — skip it.
            "sectPr" if Some(child.id()) != body_level_sect_id => {
                match read_section_break_type(child).as_deref() {
                    Some("continuous") => {}
                    Some("oddPage") => body.push(BodyElement::PageBreak {
                        parity: Some("odd".to_string()),
                    }),
                    Some("evenPage") => body.push(BodyElement::PageBreak {
                        parity: Some("even".to_string()),
                    }),
                    _ => body.push(BodyElement::PageBreak { parity: None }),
                }
            }
            _ => {}
        }
    }
    body
}

// Short-lived intermediate consumed immediately by the caller into BodyElement;
// boxing the Para variant would add a heap allocation per paragraph on the hot
// parse path with no real memory benefit for a transient `Vec<ParaPiece>`.
#[allow(clippy::large_enum_variant)]
enum ParaPiece {
    Para(DocParagraph),
    PageBreak,
}

/// Split a parsed paragraph at every internal page-break run. The split
/// pieces all share the source paragraph's pPr, so layout (alignment,
/// indents, line spacing, …) is preserved across the page boundary. The
/// page-break run itself is consumed; downstream code emits
/// BodyElement::PageBreak instead.
///
/// Two break flavors are recognized:
///   - `BreakType::Page`         — hard `<w:br w:type="page"/>`, always honored.
///   - `BreakType::RenderedPage` — Word's `<w:lastRenderedPageBreak/>` hint
///     (ECMA-376 §17.3.1.20), a layout cache that is NOT authoritative. We
///     paginate ourselves, so it is ignored uniformly and stripped (never
///     mixed with self-pagination per the package CLAUDE.md). Verified
///     behavior-invariant on the ruby sample that used to need it.
///
/// `<w:lastRenderedPageBreak/>` is often emitted by Word at the very
/// start of a paragraph too (echoing the page break that produced the
/// preceding paragraph break). To avoid emitting two consecutive page
/// breaks we drop chunks that contain no visible text.
///
/// Pure-page-break paragraphs are handled upstream as
/// BodyElement::PageBreak before this function ever sees them.
fn split_para_on_page_breaks(para: DocParagraph) -> Vec<ParaPiece> {
    // `<w:lastRenderedPageBreak/>` (BreakType::RenderedPage) is Word's layout
    // cache, not an authoritative break (ECMA-376 §17.3.1.20). We paginate the
    // body ourselves (computePages, TS side), so these hints are ignored
    // uniformly and stripped here — only a hard `<w:br w:type="page"/>` splits a
    // paragraph. Per the package CLAUDE.md, honoring vs ignoring the hint must
    // not be mixed. (It was previously honored only inside ruby paragraphs as a
    // stopgap for a ruby line-height drift that the docGrid-aware line metrics
    // have since absorbed: with this gate gone, private/sample-5 — 66 ruby runs,
    // 5 lastRenderedPageBreak hints, 7 pages — is byte-identical and still
    // matches its Word export.)
    let is_break_run = |r: &DocRun| {
        matches!(
            r,
            DocRun::Break {
                break_type: BreakType::Page
            }
        )
    };
    let has_break = para.runs.iter().any(is_break_run);
    if !has_break {
        // Strip the (ignored) RenderedPage runs so they don't pollute layout.
        let mut p = para;
        p.runs.retain(|r| {
            !matches!(
                r,
                DocRun::Break {
                    break_type: BreakType::RenderedPage
                }
            )
        });
        return vec![ParaPiece::Para(p)];
    }

    // Split run chunks on hard page breaks; RenderedPage runs are dropped.
    let mut chunks: Vec<Vec<DocRun>> = vec![Vec::new()];
    for run in para.runs.iter().cloned() {
        match &run {
            DocRun::Break {
                break_type: BreakType::Page,
            } => chunks.push(Vec::new()),
            DocRun::Break {
                break_type: BreakType::RenderedPage,
            } => { /* ignored hint */ }
            _ => chunks.last_mut().unwrap().push(run),
        }
    }

    let has_visible = |runs: &Vec<DocRun>| {
        runs.iter().any(|r| {
            matches!(r,
            DocRun::Text(t) if !t.text.trim().is_empty())
                || matches!(r, DocRun::Field(_) | DocRun::Image(_) | DocRun::Shape(_))
        })
    };

    // Drop trailing chunks that carry no visible content too — this
    // happens when the paragraph ends with `<w:br w:type="page"/>`
    // (Word's anchored shapes paragraph at the cover commonly does this
    // to force the cover onto its own page; the trailing empty chunk
    // would otherwise emit an extra blank paragraph + page break).
    let chunks: Vec<Vec<DocRun>> = {
        let mut c = chunks;
        while c.last().map(|r| !has_visible(r)).unwrap_or(false) && c.len() > 1 {
            c.pop();
            // Each pop also drops the preceding break that produced this
            // empty trailing chunk — modelled here by NOT emitting a break
            // before the (now removed) chunk.
        }
        c
    };

    let mut out: Vec<ParaPiece> = Vec::new();
    let mut emitted_para = false;
    for (i, runs) in chunks.into_iter().enumerate() {
        // Drop the leading chunk when it carries no visible content — this
        // happens when the paragraph starts with <w:lastRenderedPageBreak/>
        // (Word's hint duplicating a paragraph-level break that the
        // surrounding section break already covers).
        if i == 0 && !has_visible(&runs) {
            continue;
        }
        if emitted_para {
            out.push(ParaPiece::PageBreak);
        }
        let mut chunk = para.clone();
        chunk.runs = runs;
        out.push(ParaPiece::Para(chunk));
        emitted_para = true;
    }
    if out.is_empty() {
        out.push(ParaPiece::Para(para));
    }
    out
}

/// Read `<w:type w:val>` from a sectPr node, normalized to the ECMA-376
/// ST_SectionMark values ("continuous" | "nextPage" | "oddPage" | "evenPage" |
/// "nextColumn"). `None` if the element is absent (default = "nextPage").
fn read_section_break_type(sect_pr: roxmltree::Node) -> Option<String> {
    sect_pr
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "type")
        .find_map(|n| attr_w(n, "val"))
}

fn load_media_map(
    zip: &mut Zip,
    rel_map: &HashMap<String, String>,
    base_dir: &str,
) -> HashMap<String, String> {
    let mut media_map: HashMap<String, String> = HashMap::new();
    for (rid, target) in rel_map {
        if target.contains("media/") || target.contains("image") {
            let path = if target.starts_with('/') {
                target.trim_start_matches('/').to_string()
            } else {
                format!("{}{}", base_dir, target)
            };
            if let Ok(bytes) = read_zip_bytes(zip, &path) {
                let mime = if path.ends_with(".png") {
                    "image/png"
                } else if path.ends_with(".jpg") || path.ends_with(".jpeg") {
                    "image/jpeg"
                } else if path.ends_with(".gif") {
                    "image/gif"
                } else {
                    "image/png"
                };
                let b64 = B64.encode(&bytes);
                media_map.insert(rid.clone(), format!("data:{};base64,{}", mime, b64));
            }
        }
    }
    media_map
}

fn load_header_footer_set(
    zip: &mut Zip,
    type_to_target: &HashMap<String, String>,
    root_tag: &str,
    style_map: &StyleMap,
    num_map: &mut NumberingMap,
    theme: &ThemeColors,
) -> HeadersFooters {
    let mut out = HeadersFooters::default();
    for (kind, target) in type_to_target {
        let path = format!("word/{}", target);
        let xml = match read_zip_entry(zip, &path) {
            Ok(s) => s,
            Err(_) => continue,
        };

        // Per-file rels for image resolution
        let stem = target.trim_end_matches(".xml");
        let rels_path = format!("word/_rels/{}.xml.rels", stem);
        let rels_xml = read_zip_entry(zip, &rels_path).unwrap_or_default();
        let local_rel_map = parse_rels(&rels_xml);
        let local_media_map = load_media_map(zip, &local_rel_map, "word/");

        let xml_doc = match XmlDoc::parse(&xml) {
            Ok(d) => d,
            Err(_) => continue,
        };
        let Some(root) = xml_doc
            .root_element()
            .descendants()
            .find(|n| n.tag_name().name() == root_tag)
        else {
            continue;
        };

        let body = parse_body_elements(
            root,
            style_map,
            num_map,
            &local_media_map,
            &local_rel_map,
            theme,
        );
        let hf = HeaderFooter { body };
        match kind.as_str() {
            "first" => out.first = Some(hf),
            "even" => out.even = Some(hf),
            _ => out.default = Some(hf),
        }
    }
    out
}

fn parse_section(
    sect_pr: Option<roxmltree::Node>,
    rel_map: &HashMap<String, String>,
) -> (SectionProps, SectionRefs) {
    let default = SectionProps {
        page_width: 612.0,
        page_height: 792.0,
        margin_top: 72.0,
        margin_right: 72.0,
        margin_bottom: 72.0,
        margin_left: 72.0,
        header_distance: 36.0,
        footer_distance: 36.0,
        title_page: false,
        even_and_odd_headers: false,
        doc_grid_type: None,
        doc_grid_line_pitch: None,
    };

    let Some(sp) = sect_pr else {
        return (default, SectionRefs::default());
    };

    let mut props = default;
    if let Some(pg_sz) = child_w(sp, "pgSz") {
        if let Some(w) = attr_w(pg_sz, "w") {
            props.page_width = twips_to_pt(&w);
        }
        if let Some(h) = attr_w(pg_sz, "h") {
            props.page_height = twips_to_pt(&h);
        }
    }
    if let Some(pg_mar) = child_w(sp, "pgMar") {
        if let Some(v) = attr_w(pg_mar, "top") {
            props.margin_top = twips_to_pt(&v);
        }
        if let Some(v) = attr_w(pg_mar, "right") {
            props.margin_right = twips_to_pt(&v);
        }
        if let Some(v) = attr_w(pg_mar, "bottom") {
            props.margin_bottom = twips_to_pt(&v);
        }
        if let Some(v) = attr_w(pg_mar, "left") {
            props.margin_left = twips_to_pt(&v);
        }
        if let Some(v) = attr_w(pg_mar, "header") {
            props.header_distance = twips_to_pt(&v);
        }
        if let Some(v) = attr_w(pg_mar, "footer") {
            props.footer_distance = twips_to_pt(&v);
        }
    }
    props.title_page = child_w(sp, "titlePg").is_some();

    // ECMA-376 §17.6.5 w:docGrid. When @type=lines|linesAndChars with a
    // linePitch, Word renders each line of text at intervals of linePitch
    // rather than at the font's natural line height. For auto line rule the
    // multiplier applies against linePitch, not the font — which is what
    // makes a 56pt heading with lineRule=auto value=1040 (4.33×) render at
    // ~18pt × 4.33 = 78pt rather than 56pt × 1.25 × 4.33 = 303pt.
    if let Some(dg) = child_w(sp, "docGrid") {
        if let Some(t) = attr_w(dg, "type") {
            props.doc_grid_type = Some(t);
        }
        if let Some(lp) = attr_w(dg, "linePitch") {
            // linePitch is in twentieths of a point (twips), same as other
            // w: unit attributes. twips_to_pt divides by 20.
            props.doc_grid_line_pitch = Some(twips_to_pt(&lp));
        }
    }

    // Collect header/footer references
    let mut refs = SectionRefs::default();
    for child in sp.children().filter(|n| n.is_element()) {
        let local = child.tag_name().name();
        if local != "headerReference" && local != "footerReference" {
            continue;
        }
        let kind = attr_w(child, "type").unwrap_or_else(|| "default".to_string());
        let rid = child
            .attribute((R_NS, "id"))
            .or_else(|| child.attribute("id"))
            .map(|s| s.to_string());
        let Some(rid) = rid else { continue };
        let Some(target) = rel_map.get(&rid) else {
            continue;
        };
        let target = target.trim_start_matches('/').to_string();
        if local == "headerReference" {
            refs.headers.insert(kind, target);
        } else {
            refs.footers.insert(kind, target);
        }
    }

    (props, refs)
}

fn parse_paragraph(
    node: roxmltree::Node,
    style_map: &StyleMap,
    num_map: &mut NumberingMap,
    media_map: &HashMap<String, String>,
    rel_map: &HashMap<String, String>,
    theme: &ThemeColors,
    table_style_id: Option<&str>,
) -> DocParagraph {
    // Get style ID from pPr/pStyle. When absent, resolve_para falls back to the
    // paragraph style marked w:default="1" via StyleMap::default_para_style_id.
    let ppr_node = child_w(node, "pPr");
    let explicit_style_id = ppr_node
        .and_then(|p| child_w(p, "pStyle"))
        .and_then(|s| attr_w(s, "val"));

    // Resolve base formatting from style
    let (mut base_para, mut base_run) =
        style_map.resolve_para(explicit_style_id.as_deref(), table_style_id);

    // Apply direct paragraph formatting overrides
    if let Some(ppr) = ppr_node {
        let direct = parse_para_fmt(ppr);
        apply_direct_para(&mut base_para, &direct);
        // Also merge direct rPr
        if let Some(rpr) = child_w(ppr, "rPr") {
            let direct_run = parse_run_fmt(rpr);
            apply_direct_run(&mut base_run, &direct_run);
        }
    }

    let alignment = base_para
        .alignment
        .as_deref()
        .map(normalize_align)
        .unwrap_or("left")
        .to_string();
    let indent_right = base_para.indent_right.unwrap_or(0.0);
    let space_before = base_para.space_before.unwrap_or(0.0);
    let space_after = base_para.space_after.unwrap_or(0.0);
    let line_spacing = base_para.line_spacing_val.map(|v| LineSpacing {
        value: v,
        rule: base_para
            .line_spacing_rule
            .clone()
            .unwrap_or_else(|| "auto".to_string()),
        explicit: base_para.line_spacing_explicit.unwrap_or(false),
    });

    // Numbering — extract level data before advancing counter (avoids borrow conflict)
    let numbering = if let (Some(num_id), Some(num_level)) = (base_para.num_id, base_para.num_level)
    {
        if num_id != 0 {
            let (format, ind_left, tab) = num_map
                .get_level(num_id, num_level)
                .map(|l| (l.format.clone(), l.indent_left, l.tab))
                .unwrap_or_else(|| ("decimal".to_string(), 36.0, 18.0));
            let counter = num_map.advance(num_id, num_level);
            let text = num_map.resolve_text(num_id, num_level, counter);
            Some(NumberingInfo {
                num_id,
                level: num_level,
                format,
                text,
                indent_left: ind_left,
                tab,
            })
        } else {
            None
        }
    } else {
        None
    };

    // Numbering level's pPr/ind overrides the paragraph style's indent
    let (indent_left, indent_first) = if let Some(ref num) = numbering {
        num_map
            .get_level(num.num_id, num.level)
            .map(|l| (l.indent_left, -l.tab))
            .unwrap_or((
                base_para.indent_left.unwrap_or(0.0),
                base_para.indent_first.unwrap_or(0.0),
            ))
    } else {
        (
            base_para.indent_left.unwrap_or(0.0),
            base_para.indent_first.unwrap_or(0.0),
        )
    };
    // Same for the end-side indent (w:ind@right ≡ end): an RTL list level
    // defines its indent there (e.g. w:right="720" w:hanging="360").
    let indent_right = numbering
        .as_ref()
        .and_then(|num| num_map.get_level(num.num_id, num.level))
        .and_then(|l| l.indent_right)
        .unwrap_or(indent_right);

    // Parse runs
    let mut runs = vec![];
    parse_para_content(
        node, &base_run, style_map, media_map, rel_map, theme, &mut runs, None,
    );

    // NOTE: We do NOT force display-math paragraphs to center here. The math
    // block's justification (ECMA-376 §22.1.2.88 `m:jc` / §22.1.2.30 `m:defJc`)
    // is a concept independent of the paragraph's text `w:jc`; it is resolved by
    // the renderer from the per-instance `jc` on the Math run and the document
    // default `mathDefJc` (spec default `centerGroup`). The paragraph alignment
    // stays its natural text value (e.g. Tabletext = left).

    let tab_stops = base_para
        .tab_stops
        .clone()
        .unwrap_or_default()
        .into_iter()
        .map(|(pos, alignment, leader)| TabStop {
            pos,
            alignment,
            leader,
        })
        .collect();

    DocParagraph {
        alignment,
        indent_left,
        indent_right,
        indent_first,
        space_before,
        space_after,
        line_spacing,
        numbering,
        tab_stops,
        runs,
        shading: base_para.shading.clone(),
        page_break_before: base_para.page_break_before.unwrap_or(false),
        contextual_spacing: base_para.contextual_spacing.unwrap_or(false),
        // Word's built-in "Heading 1–9" styles carry an implicit keepNext even
        // when styles.xml doesn't spell it out — demoted heading paragraphs
        // (outlineLvl 0..8) are pinned to the next paragraph so the heading
        // never orphans at page bottom. Honor an explicit false to opt out.
        keep_next: base_para
            .keep_next
            .unwrap_or_else(|| base_para.outline_level.is_some()),
        keep_lines: base_para.keep_lines.unwrap_or(false),
        // ECMA-376 §17.3.1.44: widowControl defaults to true when absent.
        widow_control: base_para.widow_control.unwrap_or(true),
        borders: base_para.para_borders.clone(),
        // Fall back to the document's default paragraph style (w:default="1")
        // rather than the literal "Normal" — international templates often use
        // locale-specific IDs ("a", "標準", etc.) for the default style, and
        // contextualSpacing needs a stable ID to group adjacent paragraphs.
        style_id: explicit_style_id
            .clone()
            .or_else(|| style_map.default_para_style_id().map(str::to_string))
            .or_else(|| Some("Normal".to_string())),
        default_font_size: base_run.font_size,
        // Resolve the paragraph's default font the same way runs do (ascii
        // first, then eastAsia, through theme refs) so empty paragraphs can be
        // sized with the intended font's line metrics.
        default_font_family: theme
            .resolve_font_ref(base_run.font_family_ascii.clone())
            .or_else(|| theme.resolve_font_ref(base_run.font_family_east_asia.clone())),
        outline_level: base_para.outline_level,
        // ECMA-376 §17.3.1.6 — RTL paragraph flag resolved through the style
        // chain + direct pPr. The renderer reads it as the paragraph base
        // direction for the UAX#9 reordering and alignment-edge passes.
        bidi: base_para.bidi,
        // ECMA-376 §17.3.1.32 — when explicitly off, the renderer skips docGrid
        // line snapping for this paragraph (e.g. footnote text).
        snap_to_grid: base_para.snap_to_grid,
    }
}

#[derive(Default)]
struct FieldState {
    /// Currently inside a field (between fldChar begin and end).
    active: bool,
    /// Have we passed the `separate` fldChar yet?
    past_separate: bool,
    /// Accumulated instruction text (PAGE, NUMPAGES, etc.)
    instruction: String,
    /// Formatting from the first instrText run — used as the field's display format.
    fmt: Option<RunFmt>,
    /// Fallback text captured between `separate` and `end`.
    fallback: String,
    /// True when we recompute this field (PAGE/NUMPAGES) and swallow its cached result.
    /// False for complex fields (TOC, PAGEREF, REF, …) whose result is rendered as-is.
    substitute: bool,
}

// Threads the immutable parse context (style/media/rel maps, theme) plus the
// running output buffer and revision state through the paragraph walk; grouping
// these into a context struct would only relocate the same fields.
#[allow(clippy::too_many_arguments)]
fn parse_para_content(
    node: roxmltree::Node,
    base_run: &RunFmt,
    style_map: &StyleMap,
    media_map: &HashMap<String, String>,
    rel_map: &HashMap<String, String>,
    theme: &ThemeColors,
    runs: &mut Vec<DocRun>,
    revision: Option<&RunRevision>,
) {
    let mut field = FieldState::default();

    for child in element_children_flat(node) {
        match child.tag_name().name() {
            "r" => {
                handle_run_in_para(
                    child, base_run, style_map, media_map, theme, runs, &mut field, None, revision,
                );
            }
            "hyperlink" => {
                // Resolve URL from r:id via relationships
                let href = child
                    .attribute((R_NS, "id"))
                    .or_else(|| child.attribute("id"))
                    .and_then(|rid| rel_map.get(rid).cloned());
                for r in child
                    .children()
                    .filter(|n| n.is_element() && n.tag_name().name() == "r")
                {
                    handle_run_in_para(
                        r,
                        base_run,
                        style_map,
                        media_map,
                        theme,
                        runs,
                        &mut field,
                        Some(href.clone()),
                        revision,
                    );
                }
            }
            "ins" | "del" => {
                // ECMA-376 §17.13.5 — build a RunRevision context covering
                // every descendant run so the renderer can paint tracked
                // changes inline. Nested ins/del isn't legal per spec; the
                // inner block wins if it occurs anyway.
                let kind = if child.tag_name().name() == "ins" {
                    "insertion"
                } else {
                    "deletion"
                };
                let inner = RunRevision {
                    kind: kind.to_string(),
                    author: attr_w(child, "author"),
                    date: attr_w(child, "date"),
                };
                parse_para_content(
                    child,
                    base_run,
                    style_map,
                    media_map,
                    rel_map,
                    theme,
                    runs,
                    Some(&inner),
                );
            }
            "smartTag" => {
                parse_para_content(
                    child, base_run, style_map, media_map, rel_map, theme, runs, revision,
                );
            }
            "fldSimple" => {
                let instr = attr_w(child, "instr").unwrap_or_default();
                // Collect formatting from the first contained run (if any)
                let mut fmt = base_run.clone();
                if let Some(r) = child
                    .children()
                    .find(|n| n.is_element() && n.tag_name().name() == "r")
                {
                    if let Some(rpr) = child_w(r, "rPr") {
                        apply_direct_run(&mut fmt, &parse_run_fmt(rpr));
                    }
                }
                let fallback = extract_text_from_runs(child);
                runs.push(make_field_run(&instr, &fmt, &fallback, theme));
            }
            "oMath" => {
                let nodes = crate::math::parse_omath_nodes(child);
                if !nodes.is_empty() {
                    runs.push(DocRun::Math {
                        nodes,
                        display: false,
                        font_size: base_run.font_size.unwrap_or(DEFAULT_FONT_SIZE),
                        jc: None,
                    });
                }
            }
            "oMathPara" => {
                // A block math paragraph wraps one or more m:oMath children.
                // ECMA-376 §22.1.2.88 — `m:oMathParaPr/m:jc@m:val` (math namespace,
                // bare `val` fallback) is the per-instance justification of the
                // display equation. Document-default (`m:defJc`) resolution is the
                // renderer's job; we only surface the explicit per-instance value.
                let para_jc = child
                    .children()
                    .find(|n| n.is_element() && n.tag_name().name() == "oMathParaPr")
                    .and_then(|pr| {
                        pr.children()
                            .find(|n| n.is_element() && n.tag_name().name() == "jc")
                    })
                    .and_then(|jc| {
                        jc.attribute((M_NS, "val"))
                            .or_else(|| jc.attribute("val"))
                            .map(|s| s.to_string())
                    });
                for om in child
                    .children()
                    .filter(|n| n.is_element() && n.tag_name().name() == "oMath")
                {
                    let nodes = crate::math::parse_omath_nodes(om);
                    if !nodes.is_empty() {
                        runs.push(DocRun::Math {
                            nodes,
                            display: true,
                            font_size: base_run.font_size.unwrap_or(DEFAULT_FONT_SIZE),
                            jc: para_jc.clone(),
                        });
                    }
                }
            }
            _ => {}
        }
    }
}

// Same parse-context threading as parse_para_content, with the additional
// hyperlink/field state carried per run.
#[allow(clippy::too_many_arguments)]
fn handle_run_in_para(
    r_node: roxmltree::Node,
    base_run: &RunFmt,
    style_map: &StyleMap,
    media_map: &HashMap<String, String>,
    theme: &ThemeColors,
    runs: &mut Vec<DocRun>,
    field: &mut FieldState,
    // Outer None = not inside a hyperlink. Some(None) = hyperlink without URL. Some(Some(url)) = hyperlink with URL.
    link_href: Option<Option<String>>,
    revision: Option<&RunRevision>,
) {
    // Inspect this run for field control characters or instruction text first.
    let mut fld_char_type: Option<String> = None;
    let mut instr_text = String::new();
    for c in r_node.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "fldChar" => {
                if let Some(t) = attr_w(c, "fldCharType") {
                    fld_char_type = Some(t);
                }
            }
            "instrText" => {
                if let Some(t) = c.text() {
                    instr_text.push_str(t);
                }
            }
            _ => {}
        }
    }

    if let Some(ct) = fld_char_type {
        match ct.as_str() {
            "begin" => {
                field.active = true;
                field.past_separate = false;
                field.instruction.clear();
                field.fallback.clear();
                field.fmt = None;
                field.substitute = false;
            }
            "separate" => {
                field.past_separate = true;
                // Only PAGE / NUMPAGES are recomputed (their cached result is swallowed).
                // Complex fields (TOC, PAGEREF, REF, HYPERLINK, …) render their result
                // content as normal runs — so multi-paragraph / nested fields like a TOC
                // keep their headings, tabs and page numbers.
                field.substitute = classify_field(&field.instruction) != "other";
            }
            "end" => {
                if field.active && field.substitute {
                    let fmt = field.fmt.clone().unwrap_or_else(|| base_run.clone());
                    runs.push(make_field_run(
                        &field.instruction,
                        &fmt,
                        &field.fallback,
                        theme,
                    ));
                }
                *field = FieldState::default();
            }
            _ => {}
        }
        return;
    }

    if field.active && !field.past_separate {
        // Inside the instruction (before `separate`). Accumulate it and remember the
        // first instruction run's formatting; the (hidden) instruction never renders.
        if !instr_text.is_empty() {
            field.instruction.push_str(&instr_text);
            if field.fmt.is_none() {
                let mut fmt = base_run.clone();
                if let Some(rpr) = child_w(r_node, "rPr") {
                    apply_direct_run(&mut fmt, &parse_run_fmt(rpr));
                }
                field.fmt = Some(fmt);
            }
        }
        return;
    }

    if field.active && field.substitute {
        // Cached result of a recomputed field (PAGE/NUMPAGES) — swallow it.
        for c in r_node
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "t")
        {
            if let Some(t) = c.text() {
                field.fallback.push_str(t);
            }
        }
        return;
    }
    // field.active && past_separate && !substitute → fall through and render the result.

    // Normal run
    parse_run_inner(
        r_node, base_run, style_map, media_map, theme, runs, link_href, revision,
    );
}

fn extract_text_from_runs(node: roxmltree::Node) -> String {
    let mut out = String::new();
    for n in node.descendants() {
        if n.is_element() && n.tag_name().name() == "t" {
            if let Some(t) = n.text() {
                out.push_str(t);
            }
        }
    }
    out
}

fn make_field_run(instr: &str, fmt: &RunFmt, fallback: &str, theme: &ThemeColors) -> DocRun {
    let field_type = classify_field(instr);
    DocRun::Field(FieldRun {
        field_type,
        instruction: instr.trim().to_string(),
        fallback_text: fallback.to_string(),
        bold: fmt.bold.unwrap_or(false),
        italic: fmt.italic.unwrap_or(false),
        underline: fmt.underline.unwrap_or(false),
        strikethrough: fmt.strikethrough.unwrap_or(false),
        font_size: fmt.font_size.unwrap_or(DEFAULT_FONT_SIZE),
        color: fmt.color.clone(),
        font_family: theme
            .resolve_font_ref(fmt.font_family_ascii.clone())
            .or_else(|| theme.resolve_font_ref(fmt.font_family_east_asia.clone())),
        background: fmt.background.clone(),
        vert_align: fmt.vert_align.clone(),
        all_caps: fmt.all_caps.unwrap_or(false),
        small_caps: fmt.small_caps.unwrap_or(false),
        double_strikethrough: fmt.dstrike.unwrap_or(false),
        highlight: fmt.highlight.clone(),
    })
}

fn classify_field(instr: &str) -> String {
    let token = instr
        .split_whitespace()
        .next()
        .unwrap_or("")
        .to_ascii_uppercase();
    match token.as_str() {
        "PAGE" => "page".to_string(),
        "NUMPAGES" => "numPages".to_string(),
        _ => "other".to_string(),
    }
}

// Same parse-context threading as handle_run_in_para.
#[allow(clippy::too_many_arguments)]
fn parse_run_inner(
    node: roxmltree::Node,
    base_run: &RunFmt,
    style_map: &StyleMap,
    media_map: &HashMap<String, String>,
    theme: &ThemeColors,
    runs: &mut Vec<DocRun>,
    link_href: Option<Option<String>>,
    revision: Option<&RunRevision>,
) {
    // Merge run-level formatting
    let rpr_node = child_w(node, "rPr");
    let mut fmt = base_run.clone();

    // Apply rStyle. ECMA-376 §17.7.5: character styles overlay on top of the
    // paragraph's run formatting; they must not re-introduce docDefaults or
    // the default paragraph style — those are already baked into base_run.
    if let Some(rpr) = rpr_node {
        if let Some(rs) = child_w(rpr, "rStyle").and_then(|n| attr_w(n, "val")) {
            let style_run = style_map.resolve_run_style(&rs);
            apply_direct_run(&mut fmt, &style_run);
        }
        let direct = parse_run_fmt(rpr);
        apply_direct_run(&mut fmt, &direct);
    }

    // Skip hidden runs entirely
    if fmt.vanish.unwrap_or(false) {
        return;
    }

    // External links (URL) vs internal-document links (anchor: TOC entries,
    // cross-references). Word renders internal links as plain body text — NOT with
    // the Hyperlink style's blue/underline — so strip that styling here.
    let is_link = link_href.is_some();
    let is_external = matches!(link_href, Some(Some(_)));
    let is_internal = matches!(link_href, Some(None));
    let hyperlink = link_href.clone().flatten();
    if is_internal {
        fmt.color = base_run.color.clone();
        fmt.underline = base_run.underline;
    }

    let bold = fmt.bold.unwrap_or(false);
    let italic = fmt.italic.unwrap_or(false);
    let underline = fmt.underline.unwrap_or(false) || is_external;
    let strikethrough = fmt.strikethrough.unwrap_or(false);
    let font_size = fmt.font_size.unwrap_or(DEFAULT_FONT_SIZE);
    let color = if is_external && fmt.color.is_none() {
        Some("0563c1".to_string())
    } else {
        fmt.color.clone()
    };
    let font_family = theme
        .resolve_font_ref(fmt.font_family_ascii.clone())
        .or_else(|| theme.resolve_font_ref(fmt.font_family_east_asia.clone()));
    let vert_align = fmt.vert_align.clone();
    let all_caps = fmt.all_caps.unwrap_or(false);
    let small_caps = fmt.small_caps.unwrap_or(false);
    let double_strikethrough = fmt.dstrike.unwrap_or(false);
    let highlight = fmt.highlight.clone();
    // RTL / complex-script properties (ECMA-376 §17.3.2.30 / §17.3.2.26 /
    // §17.3.2.39). The renderer uses `rtl` to drive complex-script shaping and
    // the UAX#9 reordering / direction-mirroring pass. The cs font goes through
    // the same theme-ref resolution as the ascii/eastAsia axes so consumers
    // receive a literal family.
    let rtl = fmt.rtl;
    let cs_toggle = fmt.cs_toggle;
    let font_family_cs = theme.resolve_font_ref(fmt.font_family_cs.clone());
    let font_size_cs = fmt.font_size_cs;
    // Complex-script bold / italic (ECMA-376 §17.3.2.3 / §17.3.2.17) and the
    // complex-script language tag (§17.3.2.20). Recorded so the renderer can
    // route cs-classified content to cs formatting and decide AN digit ordering.
    let bold_cs = fmt.bold_cs;
    let italic_cs = fmt.italic_cs;
    let lang_bidi = fmt.lang_bidi.clone();

    for child in node.children().filter(|n| n.is_element()) {
        match child.tag_name().name() {
            "t" | "delText" => {
                // ECMA-376 §17.13.5: text inside <w:del> is wrapped in
                // <w:delText> instead of <w:t>, but otherwise carries the
                // same content. Accept both and attach the revision below.
                let text = child.text().unwrap_or("").to_string();
                if !text.is_empty() {
                    runs.push(DocRun::Text(TextRun {
                        text,
                        bold,
                        italic,
                        underline,
                        strikethrough,
                        font_size,
                        color: color.clone(),
                        font_family: font_family.clone(),
                        is_link,
                        background: fmt.background.clone(),
                        vert_align: vert_align.clone(),
                        hyperlink: hyperlink.clone(),
                        all_caps,
                        small_caps,
                        double_strikethrough,
                        highlight: highlight.clone(),
                        ruby: None,
                        revision: revision.cloned(),
                        rtl,
                        cs: cs_toggle,
                        font_family_cs: font_family_cs.clone(),
                        font_size_cs,
                        bold_cs,
                        italic_cs,
                        lang_bidi: lang_bidi.clone(),
                        note_ref: None,
                    }));
                }
            }
            "tab" => {
                // w:tab emits a horizontal tab character; layout handles tab stop alignment.
                runs.push(DocRun::Text(TextRun {
                    text: "\t".to_string(),
                    bold,
                    italic,
                    underline,
                    strikethrough,
                    font_size,
                    color: color.clone(),
                    font_family: font_family.clone(),
                    is_link,
                    background: fmt.background.clone(),
                    vert_align: vert_align.clone(),
                    hyperlink: hyperlink.clone(),
                    all_caps,
                    small_caps,
                    double_strikethrough,
                    highlight: highlight.clone(),
                    ruby: None,
                    revision: revision.cloned(),
                    rtl,
                    cs: cs_toggle,
                    font_family_cs: font_family_cs.clone(),
                    font_size_cs,
                    bold_cs,
                    italic_cs,
                    lang_bidi: lang_bidi.clone(),
                    note_ref: None,
                }));
            }
            "br" => {
                let break_type = attr_w(child, "type")
                    .as_deref()
                    .map(|v| match v {
                        "page" => BreakType::Page,
                        "column" => BreakType::Column,
                        _ => BreakType::Line,
                    })
                    .unwrap_or(BreakType::Line);
                runs.push(DocRun::Break { break_type });
            }
            "lastRenderedPageBreak" => {
                // ECMA-376 §17.3.1.20: Word stores a hint at the location
                // where the previous render placed a page break. We mark
                // it as a separate `RenderedPage` break type so the
                // paragraph splitter can decide whether to honor it
                // (currently: only inside ruby-bearing paragraphs).
                runs.push(DocRun::Break {
                    break_type: BreakType::RenderedPage,
                });
            }
            "ruby" => {
                // ECMA-376 §17.3.3.25 w:ruby — phonetic guide (furigana).
                // The rubyBase carries the base glyph(s); the rt carries the
                // small annotation rendered above. We emit each rubyBase run
                // as a TextRun, attaching the rt text as a `ruby` annotation
                // so the renderer can draw it above the base.
                let rt_text: String = child_w(child, "rt")
                    .map(|rt| {
                        rt.descendants()
                            .filter(|n| n.is_element() && n.tag_name().name() == "t")
                            .filter_map(|n| n.text())
                            .collect::<String>()
                    })
                    .unwrap_or_default();
                let rt_size_pt: f64 = child_w(child, "rubyPr")
                    .and_then(|rp| child_w(rp, "hps"))
                    .and_then(|hps| attr_w(hps, "val"))
                    .and_then(|v| v.parse::<f64>().ok())
                    .map(|hp| hp / 2.0) // half-points → points
                    .unwrap_or_else(|| fmt.font_size.unwrap_or(DEFAULT_FONT_SIZE) / 2.0);
                let ruby = if !rt_text.is_empty() {
                    Some(RubyAnnotation {
                        text: rt_text,
                        font_size_pt: rt_size_pt,
                    })
                } else {
                    None
                };
                if let Some(rb) = child_w(child, "rubyBase") {
                    let before = runs.len();
                    for inner in rb
                        .children()
                        .filter(|n| n.is_element() && n.tag_name().name() == "r")
                    {
                        parse_run_inner(
                            inner,
                            &fmt,
                            style_map,
                            media_map,
                            theme,
                            runs,
                            link_href.clone(),
                            revision,
                        );
                    }
                    // Attach ruby to the FIRST text run produced from rubyBase
                    // (typical case is a single base run carrying one or two
                    // glyphs). Splitting the annotation across multiple runs
                    // is uncommon — sample data has 1 base run per ruby.
                    if let Some(rb_anno) = ruby {
                        for r in &mut runs[before..] {
                            if let DocRun::Text(t) = r {
                                t.ruby = Some(rb_anno.clone());
                                break;
                            }
                        }
                    }
                }
            }
            "drawing" => {
                for r in parse_inline_drawing(child, media_map, theme) {
                    runs.push(r);
                }
            }
            "footnoteReference" | "endnoteReference" | "footnoteRef" | "endnoteRef" => {
                // ECMA-376 §17.11.6 / §17.11.7 / §17.11.16 / §17.11.17.
                // `*Reference` is the in-body mark; `*Ref` is the auto-number
                // placeholder that sits at the start of the note's own content.
                // Both render as a superscript number. The DISPLAYED number is
                // the note's sequential position (resolved by the renderer from
                // the footnotes/endnotes ordering), not the raw `@w:id` — we keep
                // the id in `text` only as a fallback. The `*Ref` placeholder
                // carries no id of its own, so it is tagged with an empty id and
                // the renderer substitutes the enclosing note's number.
                let tag = child.tag_name().name();
                let kind = if tag.starts_with("footnote") {
                    "footnote"
                } else {
                    "endnote"
                };
                let id_str = attr_w(child, "id").unwrap_or_default();
                runs.push(DocRun::Text(TextRun {
                    text: id_str.clone(),
                    bold,
                    italic,
                    underline,
                    strikethrough,
                    font_size,
                    color: color.clone(),
                    font_family: font_family.clone(),
                    is_link,
                    background: fmt.background.clone(),
                    // Force superscript regardless of the run's original
                    // vertAlign so reference markers appear above the line.
                    // NOTE: the model value is "super" (see the styles.rs
                    // w:vertAlign mapping) — the renderer only raises the
                    // baseline for that exact token.
                    vert_align: Some("super".to_string()),
                    hyperlink: hyperlink.clone(),
                    all_caps,
                    small_caps,
                    double_strikethrough,
                    highlight: highlight.clone(),
                    ruby: None,
                    revision: revision.cloned(),
                    rtl,
                    cs: cs_toggle,
                    font_family_cs: font_family_cs.clone(),
                    font_size_cs,
                    bold_cs,
                    italic_cs,
                    lang_bidi: lang_bidi.clone(),
                    note_ref: Some(crate::types::NoteRef {
                        kind: kind.to_string(),
                        id: id_str,
                    }),
                }));
            }
            "AlternateContent" => {
                // mc:AlternateContent/mc:Choice may contain w:drawing
                if let Some(choice) = child.children().find(|n| n.tag_name().name() == "Choice") {
                    for inner in choice.children().filter(|n| n.is_element()) {
                        if inner.tag_name().name() == "drawing" {
                            for r in parse_inline_drawing(inner, media_map, theme) {
                                runs.push(r);
                            }
                        }
                    }
                }
            }
            "pict" => {
                // Legacy VML drawing (ECMA-376 Part 4 §14.1): <w:pict> wraps a
                // <v:shape>/<v:rect>/<v:roundrect> with optional <v:textbox>.
                // Word still emits these for simple text boxes. We surface the
                // shape's fill/stroke/size and its txbxContent as a ShapeRun so
                // the existing shape renderer draws the panel + RTL body text.
                if let Some(shp) = parse_vml_pict(child, theme) {
                    runs.push(DocRun::Shape(Box::new(shp)));
                }
            }
            _ => {}
        }
    }
}

fn parse_inline_drawing(
    node: roxmltree::Node,
    media_map: &HashMap<String, String>,
    theme: &ThemeColors,
) -> Vec<DocRun> {
    // Distinguish inline vs anchor
    let is_anchor = node.descendants().any(|n| n.tag_name().name() == "anchor");

    if !is_anchor {
        let container = match node.descendants().find(|n| n.tag_name().name() == "inline") {
            Some(c) => c,
            None => return vec![],
        };
        let extent = match container
            .children()
            .find(|n| n.tag_name().name() == "extent")
        {
            Some(e) => e,
            None => return vec![],
        };
        let cx: f64 = match extent.attribute("cx").and_then(|v| v.parse().ok()) {
            Some(v) => v,
            None => return vec![],
        };
        let cy: f64 = match extent.attribute("cy").and_then(|v| v.parse().ok()) {
            Some(v) => v,
            None => return vec![],
        };
        let blip = match node.descendants().find(|n| n.tag_name().name() == "blip") {
            Some(b) => b,
            None => return vec![],
        };
        let r_id = match blip
            .attribute((R_NS, "embed"))
            .or_else(|| blip.attribute("r:embed"))
        {
            Some(r) => r,
            None => return vec![],
        };
        let data_url = match media_map.get(r_id) {
            Some(u) => u.clone(),
            None => return vec![],
        };
        return vec![DocRun::Image(ImageRun {
            data_url,
            width_pt: cx / 12700.0,
            height_pt: cy / 12700.0,
            anchor: false,
            anchor_x_pt: 0.0,
            anchor_y_pt: 0.0,
            anchor_x_from_margin: false,
            anchor_y_from_para: false,
            color_replace_from: None,
            wrap_mode: None,
            dist_top: 0.0,
            dist_bottom: 0.0,
            dist_left: 0.0,
            dist_right: 0.0,
            wrap_side: None,
            // Inline image: not a float, so the overlap constraint is moot.
            // Carry the spec no-constraint value (true).
            allow_overlap: true,
        })];
    }

    // ── Anchor image/shape ─────────────────────────────────
    let container = match node.descendants().find(|n| n.tag_name().name() == "anchor") {
        Some(c) => c,
        None => return vec![],
    };

    // Parse positionH / positionV with relativeFrom
    let (pos_x, x_from_margin, x_align) = parse_anchor_pos_h(&container);
    let (pos_y, y_from_para, y_align) = parse_anchor_pos_v(&container);
    let (pct_h, pct_v, rel_h, rel_v) = parse_anchor_pct_pos(&container);
    let (size_w_pct, size_h_pct, size_w_rel, size_h_rel) = parse_anchor_size_rel(&container);
    let anchor_meta = parse_anchor_wrap(&container);

    // behindDoc="1" flag — renderer uses this to draw shapes before text
    let behind_doc = container
        .attribute("behindDoc")
        .map(|v| v == "1" || v == "true")
        .unwrap_or(false);

    let apply_pos_meta = |shp: &mut ShapeRun| {
        shp.anchor_x_align = x_align.clone();
        shp.anchor_y_align = y_align.clone();
        shp.pct_pos_h = pct_h;
        shp.pct_pos_v = pct_v;
        shp.anchor_x_relative_from = rel_h.clone();
        shp.anchor_y_relative_from = rel_v.clone();
        shp.width_pct = size_w_pct;
        shp.height_pct = size_h_pct;
        shp.width_relative_from = size_w_rel.clone();
        shp.height_relative_from = size_h_rel.clone();
    };

    // Check for wgp (Word Graphics Group) — expands to multiple per-element entries
    if let Some(wgp) = container
        .descendants()
        .find(|n| n.tag_name().name() == "wgp")
    {
        let mut out: Vec<DocRun> = Vec::new();
        for img in parse_wgp_images(
            wgp,
            media_map,
            pos_x,
            x_from_margin,
            pos_y,
            y_from_para,
            &anchor_meta,
        ) {
            out.push(DocRun::Image(img));
        }
        for mut shp in parse_wgp_shapes(
            wgp,
            theme,
            pos_x,
            x_from_margin,
            pos_y,
            y_from_para,
            &anchor_meta,
        ) {
            shp.behind_doc = behind_doc;
            apply_pos_meta(&mut shp);
            out.push(DocRun::Shape(Box::new(shp)));
        }
        return out;
    }

    // wps:wsp directly under the anchor (no wgp wrapper)
    if let Some(wsp) = container
        .descendants()
        .find(|n| n.tag_name().name() == "wsp")
    {
        if let Some(mut shp) = parse_wsp_shape(
            wsp,
            theme,
            pos_x,
            x_from_margin,
            pos_y,
            y_from_para,
            &anchor_meta,
            1.0,
            1.0,
            0.0,
            0.0,
            0,
        ) {
            shp.behind_doc = behind_doc;
            apply_pos_meta(&mut shp);
            return vec![DocRun::Shape(Box::new(shp))];
        }
    }

    // Regular single-blip anchor
    let extent = match container
        .children()
        .find(|n| n.tag_name().name() == "extent")
    {
        Some(e) => e,
        None => return vec![],
    };
    let cx: f64 = match extent.attribute("cx").and_then(|v| v.parse().ok()) {
        Some(v) => v,
        None => return vec![],
    };
    let cy: f64 = match extent.attribute("cy").and_then(|v| v.parse().ok()) {
        Some(v) => v,
        None => return vec![],
    };
    let blip = match node.descendants().find(|n| n.tag_name().name() == "blip") {
        Some(b) => b,
        None => return vec![],
    };
    let r_id = match blip
        .attribute((R_NS, "embed"))
        .or_else(|| blip.attribute("r:embed"))
    {
        Some(r) => r,
        None => return vec![],
    };
    let data_url = match media_map.get(r_id) {
        Some(u) => u.clone(),
        None => return vec![],
    };
    vec![DocRun::Image(ImageRun {
        data_url,
        width_pt: cx / 12700.0,
        height_pt: cy / 12700.0,
        anchor: true,
        anchor_x_pt: pos_x,
        anchor_y_pt: pos_y,
        anchor_x_from_margin: x_from_margin,
        anchor_y_from_para: y_from_para,
        color_replace_from: None,
        wrap_mode: anchor_meta.wrap_mode.clone(),
        dist_top: anchor_meta.dist_top,
        dist_bottom: anchor_meta.dist_bottom,
        dist_left: anchor_meta.dist_left,
        dist_right: anchor_meta.dist_right,
        wrap_side: anchor_meta.wrap_side.clone(),
        allow_overlap: anchor_meta.allow_overlap,
    })]
}

#[derive(Clone)]
struct AnchorMeta {
    wrap_mode: Option<String>,
    wrap_side: Option<String>,
    dist_top: f64,
    dist_bottom: f64,
    dist_left: f64,
    dist_right: f64,
    /// ECMA-376 §20.4.2.3 `wp:anchor/@allowOverlap` — whether this floating
    /// object may overlap other floating objects. Spec default is **true**
    /// (the attribute is optional). `false` mandates the object be repositioned
    /// to prevent overlap. We implement `Default` by hand so the spec default of
    /// `true` is preserved everywhere `AnchorMeta::default()` is used (a derived
    /// `Default` would wrongly yield `false`).
    allow_overlap: bool,
}

impl Default for AnchorMeta {
    fn default() -> Self {
        AnchorMeta {
            wrap_mode: None,
            wrap_side: None,
            dist_top: 0.0,
            dist_bottom: 0.0,
            dist_left: 0.0,
            dist_right: 0.0,
            // §20.4.2.3: omitted @allowOverlap ⇒ true.
            allow_overlap: true,
        }
    }
}

/// Parse wrap element and dist* padding from a wp:anchor container.
fn parse_anchor_wrap(container: &roxmltree::Node) -> AnchorMeta {
    let to_pt = |s: &str| s.parse::<f64>().ok().map(|v| v / 12700.0).unwrap_or(0.0);
    let dist_top = container.attribute("distT").map(to_pt).unwrap_or(0.0);
    let dist_bottom = container.attribute("distB").map(to_pt).unwrap_or(0.0);
    let dist_left = container.attribute("distL").map(to_pt).unwrap_or(0.0);
    let dist_right = container.attribute("distR").map(to_pt).unwrap_or(0.0);

    // ECMA-376 §20.4.2.3 `wp:anchor/@allowOverlap`: "1"/"true" ⇒ true,
    // "0"/"false" ⇒ false, omitted ⇒ true (spec default).
    let allow_overlap = container
        .attribute("allowOverlap")
        .map(|v| v == "1" || v == "true")
        .unwrap_or(true);

    let mut wrap_mode: Option<String> = None;
    let mut wrap_side: Option<String> = None;

    for child in container.children().filter(|n| n.is_element()) {
        let name = child.tag_name().name();
        match name {
            "wrapSquare" => {
                wrap_mode = Some("square".into());
                wrap_side = child.attribute("wrapText").map(|s| s.to_string());
                break;
            }
            "wrapTopAndBottom" => {
                wrap_mode = Some("topAndBottom".into());
                break;
            }
            "wrapNone" => {
                wrap_mode = Some("none".into());
                break;
            }
            "wrapTight" => {
                wrap_mode = Some("tight".into());
                wrap_side = child.attribute("wrapText").map(|s| s.to_string());
                break;
            }
            "wrapThrough" => {
                wrap_mode = Some("through".into());
                wrap_side = child.attribute("wrapText").map(|s| s.to_string());
                break;
            }
            _ => {}
        }
    }

    AnchorMeta {
        wrap_mode,
        wrap_side,
        dist_top,
        dist_bottom,
        dist_left,
        dist_right,
        allow_overlap,
    }
}

/// Parse positionH — returns (posOffset_pt, needs_margin_offset).
/// "column" and "margin" relative offsets both mean: add marginLeft in the renderer.
/// `<wp:positionH>` / `<wp:positionV>` may live directly under `<wp:anchor>`,
/// or be wrapped in `<mc:AlternateContent>` for Word 2010+ pct-based positioning.
/// In the wrapped form `<mc:Choice>` holds the wp14 pct-based variant and
/// `<mc:Fallback>` holds a posOffset variant. Always pick Choice (matches what
/// Word renders in 2010+); never read from Fallback.
fn find_position_node<'a, 'i>(
    container: &roxmltree::Node<'a, 'i>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'i>> {
    if let Some(n) = container.children().find(|n| n.tag_name().name() == name) {
        return Some(n);
    }
    for ac in container
        .children()
        .filter(|n| n.tag_name().name() == "AlternateContent")
    {
        if let Some(choice) = ac.children().find(|n| n.tag_name().name() == "Choice") {
            if let Some(n) = choice
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == name)
            {
                return Some(n);
            }
        }
    }
    None
}

fn parse_anchor_pos_h(container: &roxmltree::Node) -> (f64, bool, Option<String>) {
    let pos = match find_position_node(container, "positionH") {
        Some(p) => p,
        None => return (0.0, false, None),
    };
    let rel = pos.attribute("relativeFrom").unwrap_or("page");
    let offset = pos
        .children()
        .find(|n| n.tag_name().name() == "posOffset")
        .and_then(|n| n.text())
        .and_then(|t| t.parse::<f64>().ok())
        .map(|emu| emu / 12700.0)
        .unwrap_or(0.0);
    // ECMA-376 §20.4.3.1: <wp:align>left|center|right</wp:align> takes
    // precedence over posOffset when both are present.
    let align = pos
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "align")
        .and_then(|n| n.text())
        .map(|s| s.trim().to_string())
        .filter(|s| !s.is_empty());
    let from_margin = matches!(rel, "column" | "margin" | "leftMargin" | "insideMargin");
    (offset, from_margin, align)
}

/// Parse positionV — returns (posOffset_pt, is_paragraph_relative, align).
fn parse_anchor_pos_v(container: &roxmltree::Node) -> (f64, bool, Option<String>) {
    let pos = match find_position_node(container, "positionV") {
        Some(p) => p,
        None => return (0.0, false, None),
    };
    let rel = pos.attribute("relativeFrom").unwrap_or("page");
    let offset = pos
        .children()
        .find(|n| n.tag_name().name() == "posOffset")
        .and_then(|n| n.text())
        .and_then(|t| t.parse::<f64>().ok())
        .map(|emu| emu / 12700.0)
        .unwrap_or(0.0);
    let align = pos
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "align")
        .and_then(|n| n.text())
        .map(|s| s.trim().to_string())
        .filter(|s| !s.is_empty());
    let from_para = matches!(rel, "paragraph" | "line");
    (offset, from_para, align)
}

/// Read ECMA-376 §20.4.2.7 wp14:pctPos{H,V}Offset and the positionH/V
/// relativeFrom strings. Both values are normalized into a fraction in
/// `[0, 1]`; the renderer multiplies them by the relative container's
/// dimension. The relativeFrom string is captured raw ("page", "margin",
/// "topMargin", "rightMargin", "insideMargin", "paragraph", "line", …)
/// so the renderer can pick the right container and edge.
fn parse_anchor_pct_pos(
    container: &roxmltree::Node,
) -> (Option<f64>, Option<f64>, Option<String>, Option<String>) {
    let read_pct = |pos_node: roxmltree::Node| -> Option<f64> {
        // <wp14:pctPos*Offset> may be wrapped in <mc:AlternateContent>/<mc:Choice>
        // — search descendants under positionH/V to be safe.
        pos_node
            .descendants()
            .filter(|n| n.is_element())
            .find(|n| matches!(n.tag_name().name(), "pctPosHOffset" | "pctPosVOffset"))
            .and_then(|n| n.text())
            .and_then(|t| t.parse::<f64>().ok())
            .map(|v| v / 100_000.0)
    };
    let read_rel = |pos_node: roxmltree::Node| -> Option<String> {
        pos_node.attribute("relativeFrom").map(|s| s.to_string())
    };
    let h_node = find_position_node(container, "positionH");
    let v_node = find_position_node(container, "positionV");
    (
        h_node.and_then(read_pct),
        v_node.and_then(read_pct),
        h_node.and_then(read_rel),
        v_node.and_then(read_rel),
    )
}

/// Read ECMA-376 §20.4.2.18 wp14:sizeRelH / sizeRelV — width/height as a
/// fraction of the relativeFrom container. Returns
/// `(width_pct, height_pct, width_relative_from, height_relative_from)`.
/// `pct == 0` is treated as None (fall back to extent), matching Word.
fn parse_anchor_size_rel(
    container: &roxmltree::Node,
) -> (Option<f64>, Option<f64>, Option<String>, Option<String>) {
    let read = |outer_tag: &str, inner_tag: &str| -> (Option<f64>, Option<String>) {
        let node = container
            .children()
            .find(|n| n.tag_name().name() == outer_tag);
        let pct = node
            .and_then(|n| {
                n.descendants()
                    .find(|c| c.is_element() && c.tag_name().name() == inner_tag)
                    .and_then(|c| c.text())
                    .and_then(|t| t.parse::<f64>().ok())
            })
            .map(|v| v / 100_000.0)
            .filter(|v| *v > 0.0);
        let rel = node.and_then(|n| n.attribute("relativeFrom").map(|s| s.to_string()));
        (pct, rel)
    };
    let (w_pct, w_rel) = read("sizeRelH", "pctWidth");
    let (h_pct, h_rel) = read("sizeRelV", "pctHeight");
    (w_pct, h_pct, w_rel, h_rel)
}

/// Expand a wp:wgp group into individual ImageRun entries.
/// Each pic child gets page-relative coordinates: group anchor origin + child offset within group.
fn parse_wgp_images(
    wgp: roxmltree::Node,
    media_map: &HashMap<String, String>,
    anchor_pos_x: f64,
    x_from_margin: bool,
    anchor_pos_y: f64,
    y_from_para: bool,
    anchor_meta: &AnchorMeta,
) -> Vec<ImageRun> {
    // Pictures inside a wpg group live in the group's child coordinate space and
    // must be mapped to page space through the cumulative transform of every
    // group on the path from the wgp down to the pic (ECMA-376 §20.1.7.5/.6),
    // exactly like parse_wgp_shapes. The base transform is the outermost wgp's
    // own grpSpPr/xfrm (chOff/chExt → off/ext); nested wpg:grpSp groups compose
    // their transforms on top as we descend. (The old code applied only each
    // pic's own offset, ignoring both the group's scale/offset and any nested
    // grpSp transform, mis-placing/mis-sizing grouped pictures.)
    let base = match group_xfrm(wgp) {
        Some(x) => GroupTransform::IDENTITY.compose_child(x),
        None => GroupTransform::IDENTITY,
    };
    let mut results = Vec::new();
    walk_group_images(
        wgp,
        base,
        media_map,
        anchor_pos_x,
        x_from_margin,
        anchor_pos_y,
        y_from_para,
        anchor_meta,
        &mut results,
    );
    results
}

/// Recursively walk the element children of a group (`wpg:wgp` or nested
/// `wpg:grpSp`), composing each nested grpSp's transform into `xform` before
/// descending, and emit an `ImageRun` for every `pic:pic` using the cumulative
/// transform. Mirror image of `walk_group_children` (shapes); pre-order
/// preserves document order.
#[allow(clippy::too_many_arguments)]
fn walk_group_images(
    group: roxmltree::Node,
    xform: GroupTransform,
    media_map: &HashMap<String, String>,
    anchor_pos_x: f64,
    x_from_margin: bool,
    anchor_pos_y: f64,
    y_from_para: bool,
    anchor_meta: &AnchorMeta,
    results: &mut Vec<ImageRun>,
) {
    for child in group.children().filter(|n| n.is_element()) {
        match child.tag_name().name() {
            "pic" => {
                if let Some(img) = parse_group_pic(
                    child,
                    xform,
                    media_map,
                    anchor_pos_x,
                    x_from_margin,
                    anchor_pos_y,
                    y_from_para,
                    anchor_meta,
                ) {
                    results.push(img);
                }
            }
            "grpSp" => {
                let child_xform = match group_xfrm(child) {
                    Some(x) => xform.compose_child(x),
                    None => xform,
                };
                walk_group_images(
                    child,
                    child_xform,
                    media_map,
                    anchor_pos_x,
                    x_from_margin,
                    anchor_pos_y,
                    y_from_para,
                    anchor_meta,
                    results,
                );
            }
            _ => {}
        }
    }
}

/// Build an `ImageRun` for a single `pic:pic` inside a group, mapping its
/// spPr/xfrm off (position) and ext (size) from the immediate group's child
/// coordinate space to page space via the cumulative transform `xform`. The
/// position/size math matches `parse_wsp_shape`:
///   width_pt    = cx * scale_x / EMU_PER_PT
///   anchor_x_pt = anchor_pos_x + xform.off_x_pt + ox * scale_x / EMU_PER_PT
#[allow(clippy::too_many_arguments)]
fn parse_group_pic(
    pic: roxmltree::Node,
    xform: GroupTransform,
    media_map: &HashMap<String, String>,
    anchor_pos_x: f64,
    x_from_margin: bool,
    anchor_pos_y: f64,
    y_from_para: bool,
    anchor_meta: &AnchorMeta,
) -> Option<ImageRun> {
    // Position and size come from the pic's spPr > a:xfrm (child-coord EMU).
    let sp_pr = pic.children().find(|n| n.tag_name().name() == "spPr")?;
    let xfrm = sp_pr.children().find(|n| n.tag_name().name() == "xfrm")?;
    let off = xfrm.children().find(|n| n.tag_name().name() == "off")?;
    let ext = xfrm.children().find(|n| n.tag_name().name() == "ext")?;
    let ox = off
        .attribute("x")
        .and_then(|v| v.parse::<f64>().ok())
        .unwrap_or(0.0);
    let oy = off
        .attribute("y")
        .and_then(|v| v.parse::<f64>().ok())
        .unwrap_or(0.0);
    let cx = ext
        .attribute("cx")
        .and_then(|v| v.parse::<f64>().ok())
        .unwrap_or(0.0);
    let cy = ext
        .attribute("cy")
        .and_then(|v| v.parse::<f64>().ok())
        .unwrap_or(0.0);

    if cx <= 0.0 || cy <= 0.0 {
        return None;
    }

    // Find the blip inside this pic.
    let blip = pic.descendants().find(|n| n.tag_name().name() == "blip")?;
    let r_id = blip
        .attribute((R_NS, "embed"))
        .or_else(|| blip.attribute("r:embed"))?;
    let data_url = media_map.get(r_id)?.clone();

    // Parse a:clrChange if present — used to make a specific color transparent.
    // clrFrom specifies the source color; clrTo with alpha=0 means replace with transparent.
    let color_replace_from = blip
        .children()
        .find(|n| n.tag_name().name() == "clrChange")
        .and_then(|cc| cc.children().find(|n| n.tag_name().name() == "clrFrom"))
        .and_then(|cf| cf.children().find(|n| n.tag_name().name() == "srgbClr"))
        .and_then(|clr| clr.attribute("val").map(|v| v.to_uppercase()));

    Some(ImageRun {
        data_url,
        width_pt: cx * xform.scale_x / 12700.0,
        height_pt: cy * xform.scale_y / 12700.0,
        anchor: true,
        // Map the pic offset through the group chain, then add the page-space
        // anchor offset of the whole group.
        anchor_x_pt: anchor_pos_x + xform.off_x_emu / 12700.0 + ox * xform.scale_x / 12700.0,
        anchor_y_pt: anchor_pos_y + xform.off_y_emu / 12700.0 + oy * xform.scale_y / 12700.0,
        anchor_x_from_margin: x_from_margin,
        anchor_y_from_para: y_from_para,
        color_replace_from,
        wrap_mode: anchor_meta.wrap_mode.clone(),
        dist_top: anchor_meta.dist_top,
        dist_bottom: anchor_meta.dist_bottom,
        dist_left: anchor_meta.dist_left,
        dist_right: anchor_meta.dist_right,
        wrap_side: anchor_meta.wrap_side.clone(),
        allow_overlap: anchor_meta.allow_overlap,
    })
}

/// Cumulative child-coord-space → page-space affine transform built up while
/// descending a `wpg:wgp` / nested `wpg:grpSp` tree. A child point `local`
/// (EMU, in the immediate group's child coordinate system) maps to page EMU as
/// `page = off_emu + local * scale` (independent per axis; OOXML group
/// transforms have no skew). ECMA-376 §20.1.7.5 (`a:grpSpPr` group transform)
/// and §20.1.7.6 (`a:xfrm` child offset/extent) define this scale/offset, and
/// nested groups compose their transforms multiplicatively.
#[derive(Clone, Copy)]
struct GroupTransform {
    scale_x: f64,
    scale_y: f64,
    off_x_emu: f64,
    off_y_emu: f64,
}

impl GroupTransform {
    const IDENTITY: GroupTransform = GroupTransform {
        scale_x: 1.0,
        scale_y: 1.0,
        off_x_emu: 0.0,
        off_y_emu: 0.0,
    };

    /// Compose with the transform of a child group whose own `grpSpPr/xfrm`
    /// gives off/ext/chOff/chExt. The child group maps grandchild coordinates
    /// `g` to this group's child space via `mid = g_off - g_chOff*g_scale +
    /// g*g_scale`; applying `self` (child→page) on top yields the composite
    /// `page = self.off + self.scale*(mid)`. Expanding gives:
    ///   new_scale = self.scale * g_scale
    ///   new_off   = self.off + self.scale * (g_off - g_chOff * g_scale)
    fn compose_child(self, xfrm: roxmltree::Node) -> GroupTransform {
        let (off_x, off_y, ext_cx, ext_cy, ch_off_x, ch_off_y, ch_ext_cx, ch_ext_cy) =
            read_group_xfrm(xfrm);
        let g_scale_x = if ch_ext_cx > 0.0 && ext_cx > 0.0 {
            ext_cx / ch_ext_cx
        } else {
            1.0
        };
        let g_scale_y = if ch_ext_cy > 0.0 && ext_cy > 0.0 {
            ext_cy / ch_ext_cy
        } else {
            1.0
        };
        GroupTransform {
            scale_x: self.scale_x * g_scale_x,
            scale_y: self.scale_y * g_scale_y,
            off_x_emu: self.off_x_emu + self.scale_x * (off_x - ch_off_x * g_scale_x),
            off_y_emu: self.off_y_emu + self.scale_y * (off_y - ch_off_y * g_scale_y),
        }
    }
}

/// Read off/ext/chOff/chExt (EMU) from a group `a:xfrm`. Returns
/// (off_x, off_y, ext_cx, ext_cy, ch_off_x, ch_off_y, ch_ext_cx, ch_ext_cy).
fn read_group_xfrm(xfrm: roxmltree::Node) -> (f64, f64, f64, f64, f64, f64, f64, f64) {
    let attr = |node: Option<roxmltree::Node>, name: &str| {
        node.and_then(|n| n.attribute(name).and_then(|v| v.parse::<f64>().ok()))
            .unwrap_or(0.0)
    };
    let off = xfrm
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "off");
    let ext = xfrm
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "ext");
    let ch_off = xfrm
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "chOff");
    let ch_ext = xfrm
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "chExt");
    (
        attr(off, "x"),
        attr(off, "y"),
        attr(ext, "cx"),
        attr(ext, "cy"),
        attr(ch_off, "x"),
        attr(ch_off, "y"),
        attr(ch_ext, "cx"),
        attr(ch_ext, "cy"),
    )
}

/// Locate a group's `grpSpPr > xfrm` (the group's own transform), if present.
/// Applies to both `wpg:wgp` and nested `wpg:grpSp`, which carry their
/// transform in a direct `grpSpPr` child.
fn group_xfrm<'a, 'i>(group: roxmltree::Node<'a, 'i>) -> Option<roxmltree::Node<'a, 'i>> {
    group
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "grpSpPr")
        .and_then(|gsp| {
            gsp.children()
                .find(|n| n.is_element() && n.tag_name().name() == "xfrm")
        })
}

/// Expand the wps:wsp shapes of a `wpg:wgp` into ShapeRun entries, composing
/// nested `wpg:grpSp` group transforms recursively (ECMA-376 §20.1.7.5/.6).
/// Each grpSp scales and offsets its children relative to its parent group, so
/// the cumulative child→page transform must be the product of every group on
/// the path from the wgp down to the wsp — not just the outermost grpSpPr.
fn parse_wgp_shapes(
    wgp: roxmltree::Node,
    theme: &ThemeColors,
    anchor_pos_x: f64,
    x_from_margin: bool,
    anchor_pos_y: f64,
    y_from_para: bool,
    anchor_meta: &AnchorMeta,
) -> Vec<ShapeRun> {
    // Base transform = the outermost wgp grpSpPr/xfrm (chOff/chExt → off/ext).
    let base = match group_xfrm(wgp) {
        Some(x) => GroupTransform::IDENTITY.compose_child(x),
        None => GroupTransform::IDENTITY,
    };

    // Outer group dimensions in pt — passed to EVERY child (nested or not) so
    // the renderer resolves align/pctPos against the whole group's bounding
    // box. Falls back to the child-coord-space ext when the group omits an
    // outer ext (rare). This is the outermost wgp's box and is invariant to
    // nesting depth.
    let (group_w_pt, group_h_pt) = match group_xfrm(wgp) {
        Some(x) => {
            let (_, _, ext_cx, ext_cy, _, _, ch_ext_cx, ch_ext_cy) = read_group_xfrm(x);
            (
                (if ext_cx > 0.0 { ext_cx } else { ch_ext_cx }) / 12700.0,
                (if ext_cy > 0.0 { ext_cy } else { ch_ext_cy }) / 12700.0,
            )
        }
        None => (0.0, 0.0),
    };

    let mut results = Vec::new();
    let mut z_order: u32 = 0;
    walk_group_children(
        wgp,
        base,
        theme,
        anchor_pos_x,
        x_from_margin,
        anchor_pos_y,
        y_from_para,
        anchor_meta,
        group_w_pt,
        group_h_pt,
        &mut z_order,
        &mut results,
    );
    results
}

/// Recursively walk the element children of a group (`wpg:wgp` or nested
/// `wpg:grpSp`), composing each nested grpSp's transform into `xform` before
/// descending. `wps:wsp` children are emitted as ShapeRun using the cumulative
/// transform; `z_order` increments in document (pre-order) order so the
/// resulting z-index matches a flat descendant walk.
#[allow(clippy::too_many_arguments)]
fn walk_group_children(
    group: roxmltree::Node,
    xform: GroupTransform,
    theme: &ThemeColors,
    anchor_pos_x: f64,
    x_from_margin: bool,
    anchor_pos_y: f64,
    y_from_para: bool,
    anchor_meta: &AnchorMeta,
    group_w_pt: f64,
    group_h_pt: f64,
    z_order: &mut u32,
    results: &mut Vec<ShapeRun>,
) {
    for child in group.children().filter(|n| n.is_element()) {
        match child.tag_name().name() {
            "wsp" => {
                let idx = *z_order;
                *z_order += 1;
                if let Some(mut shape) = parse_wsp_shape(
                    child,
                    theme,
                    anchor_pos_x,
                    x_from_margin,
                    anchor_pos_y,
                    y_from_para,
                    anchor_meta,
                    xform.scale_x,
                    xform.scale_y,
                    xform.off_x_emu / 12700.0,
                    xform.off_y_emu / 12700.0,
                    idx,
                ) {
                    shape.group_width_pt = Some(group_w_pt);
                    shape.group_height_pt = Some(group_h_pt);
                    results.push(shape);
                }
            }
            "grpSp" => {
                // Compose this nested group's transform, then recurse.
                let child_xform = match group_xfrm(child) {
                    Some(x) => xform.compose_child(x),
                    None => xform,
                };
                walk_group_children(
                    child,
                    child_xform,
                    theme,
                    anchor_pos_x,
                    x_from_margin,
                    anchor_pos_y,
                    y_from_para,
                    anchor_meta,
                    group_w_pt,
                    group_h_pt,
                    z_order,
                    results,
                );
            }
            _ => {}
        }
    }
}

/// Parse a single wps:wsp into ShapeRun. `sx,sy` scale the shape's spPr/xfrm
/// from group child coord space to page EMU; `group_off_pt_*` are the group origin
/// on the page (in pt) so the shape's off.x/off.y (in child coord space) can be
/// translated to page-relative pt. For a standalone wsp (no wgp), pass sx=sy=1, group_off=0.
// Carries the accumulated anchor/group coordinate transform (offsets, scale,
// relative-from flags, z-order) needed to place a wsp shape in page space;
// these are interdependent transform parameters, not an arbitrary bag.
#[allow(clippy::too_many_arguments)]
fn parse_wsp_shape(
    wsp: roxmltree::Node,
    theme: &ThemeColors,
    anchor_pos_x: f64,
    x_from_margin: bool,
    anchor_pos_y: f64,
    y_from_para: bool,
    anchor_meta: &AnchorMeta,
    sx: f64,
    sy: f64,
    group_off_pt_x: f64,
    group_off_pt_y: f64,
    z_order: u32,
) -> Option<ShapeRun> {
    let sp_pr = wsp
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spPr")?;
    let xfrm = sp_pr
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "xfrm")?;
    let off = xfrm
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "off")?;
    let ext = xfrm
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "ext")?;
    let ox = off.attribute("x").and_then(|v| v.parse::<f64>().ok())?;
    let oy = off.attribute("y").and_then(|v| v.parse::<f64>().ok())?;
    let cx = ext.attribute("cx").and_then(|v| v.parse::<f64>().ok())?;
    let cy = ext.attribute("cy").and_then(|v| v.parse::<f64>().ok())?;

    // Line/connector presets (ECMA-376 §20.1.9.18 prstGeom; preset geometries
    // `line`, `straightConnector1`, `bent*Connector*`, `curved*Connector*`)
    // legitimately have a degenerate bounding box: an axis-aligned connector
    // has cx==0 (vertical) or cy==0 (horizontal). Such a shape must NOT be
    // discarded — it is the line itself. A genuine zero-area box on any other
    // geometry (rect, ellipse, …) has nothing to draw, and a negative extent
    // is always invalid, so both are still rejected.
    let prst_lower = sp_pr
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "prstGeom")
        .and_then(|n| n.attribute("prst"))
        .map(|p| p.to_ascii_lowercase());
    let is_line_geom = matches!(
        prst_lower.as_deref(),
        Some(p) if p == "line"
            || p.starts_with("straightconnector")
            || p.starts_with("bentconnector")
            || p.starts_with("curvedconnector")
    );
    if cx < 0.0 || cy < 0.0 {
        return None;
    }
    if !is_line_geom && (cx == 0.0 || cy == 0.0) {
        return None;
    }
    let rotation = xfrm
        .attribute("rot")
        .and_then(|v| v.parse::<f64>().ok())
        .map(|r| r / 60000.0) // OOXML rotation: 60000ths of a degree
        .unwrap_or(0.0);
    // §20.1.7.6 a:xfrm flipH/flipV — "1"/"true" mirror the shape.
    let flip_h = matches!(xfrm.attribute("flipH"), Some("1") | Some("true"));
    let flip_v = matches!(xfrm.attribute("flipV"), Some("1") | Some("true"));

    let width_pt = cx * sx / 12700.0;
    let height_pt = cy * sy / 12700.0;
    let anchor_x_pt = anchor_pos_x + group_off_pt_x + ox * sx / 12700.0;
    let anchor_y_pt = anchor_pos_y + group_off_pt_y + oy * sy / 12700.0;

    let cust_geom = sp_pr
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "custGeom");
    let (subpaths, preset_geometry, adj_values) = if let Some(cg) = cust_geom {
        (parse_custom_geometry(cg), None, Vec::new())
    } else {
        // Defer prstGeom rendering to core's buildShapePath. Carry the preset
        // name + adjustment values so the renderer can call into the shared
        // catalog (matches pptx coverage: ellipse, roundRect, triangles,
        // arrows, callouts, ribbons, …). Unknown presets are still passed
        // through; buildShapePath falls back to rect for anything it doesn't
        // recognize.
        let prst_node = sp_pr
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "prstGeom");
        let prst = prst_node
            .and_then(|n| n.attribute("prst"))
            .unwrap_or("rect")
            .to_string();
        let adj_values = prst_node
            .and_then(|p| {
                p.children()
                    .find(|n| n.is_element() && n.tag_name().name() == "avLst")
            })
            .map(|av| {
                av.children()
                    .filter(|n| n.is_element() && n.tag_name().name() == "gd")
                    .filter_map(|gd| {
                        gd.attribute("fmla")
                            .and_then(|f| f.strip_prefix("val "))
                            .and_then(|s| s.parse::<f64>().ok())
                    })
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default();
        (Vec::new(), Some(prst), adj_values)
    };
    if subpaths.is_empty() && preset_geometry.is_none() {
        return None;
    }

    // Direct fills on spPr take priority over wps:style/fillRef. ECMA-376
    // §20.1.4.1.30: when *no* direct fill is set, the shape's appearance comes
    // from the theme's fillStyleLst / bgFillStyleLst entry referenced by idx,
    // recolored using the schemeClr embedded in the fillRef. But §20.1.8.44:
    // an explicit `<a:noFill/>` is itself a direct fill property and overrides
    // the style reference, so only fall back to fillRef when the fill is Absent.
    let fill = match parse_shape_fill(sp_pr, theme) {
        FillSpec::Explicit(f) => Some(f),
        FillSpec::NoFill => None,
        FillSpec::Absent => wsp
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "style")
            .and_then(|st| {
                st.children()
                    .find(|n| n.is_element() && n.tag_name().name() == "fillRef")
            })
            .and_then(|fr| resolve_fill_ref(fr, theme)),
    };
    let ln_node = sp_pr
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "ln");
    let (stroke, stroke_width) = ln_node
        .map(|ln| {
            let has_no_fill = ln
                .children()
                .any(|n| n.is_element() && n.tag_name().name() == "noFill");
            if has_no_fill {
                return (None, 0.0);
            }
            let color = ln
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
                .and_then(|sf| resolve_color_element(sf, theme));
            let w_emu = ln
                .attribute("w")
                .and_then(|v| v.parse::<f64>().ok())
                .unwrap_or(9525.0);
            (color, w_emu / 12700.0)
        })
        .unwrap_or((None, 0.0));
    // ECMA-376 §20.1.8.48 prstDash and §20.1.8.3 head/tail line-end decorations.
    let stroke_dash = ln_node.and_then(|ln| {
        ln.children()
            .find(|n| n.is_element() && n.tag_name().name() == "prstDash")
            .and_then(|d| d.attribute("val"))
            .map(|v| v.to_string())
    });
    let parse_line_end = |name: &str| -> Option<LineEnd> {
        let ln = ln_node?;
        let end = ln
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == name)?;
        // CT_LineEndProperties: type defaults to "none"; w/len default to "med"
        // (ECMA-376 §20.1.8.3 — absent w/len means the medium step).
        let ty = end.attribute("type").unwrap_or("none");
        if ty == "none" {
            return None;
        }
        Some(LineEnd {
            r#type: ty.to_string(),
            w: end.attribute("w").unwrap_or("med").to_string(),
            len: end.attribute("len").unwrap_or("med").to_string(),
        })
    };
    let head_end = parse_line_end("headEnd");
    let tail_end = parse_line_end("tailEnd");

    // Shape body text: <wps:txbx><w:txbxContent>...</w:txbxContent></wps:txbx>
    // and the bodyPr (insets / vertical anchor).
    let (text_blocks, text_anchor, text_inset_l, text_inset_t, text_inset_r, text_inset_b) =
        parse_shape_text_body(wsp, theme);

    Some(ShapeRun {
        width_pt,
        height_pt,
        anchor_x_pt,
        anchor_y_pt,
        anchor_x_from_margin: x_from_margin,
        anchor_y_from_para: y_from_para,
        behind_doc: false,
        z_order,
        subpaths,
        preset_geometry,
        adj_values,
        fill,
        stroke,
        stroke_width,
        stroke_dash,
        head_end,
        tail_end,
        rotation,
        flip_h,
        flip_v,
        wrap_mode: anchor_meta.wrap_mode.clone(),
        text_blocks,
        text_anchor,
        text_inset_l,
        text_inset_t,
        text_inset_r,
        text_inset_b,
        ..Default::default()
    })
}

/// Extract text blocks and bodyPr from a wsp shape.
/// Returns (blocks, anchor, inset_l, inset_t, inset_r, inset_b).
///
/// Per ECMA-376 §21.1.2.1.1, lIns/tIns/rIns/bIns are the distance from
/// the rendered (page-space) bounding-box edge to the text, measured in
/// page-space EMU. They are independent of any enclosing group's coordinate
/// transform — the rendered shape edge and the rendered text position both
/// live in page space, so the inset is invariant to the group's sx/sy scale.
/// Defaults follow §21.1.2.1.1: lIns=rIns=91440 EMU (0.1in = 7.2pt),
/// tIns=bIns=45720 EMU (0.05in = 3.6pt).
fn parse_shape_text_body(
    wsp: roxmltree::Node,
    theme: &ThemeColors,
) -> (Vec<ShapeText>, Option<String>, f64, f64, f64, f64) {
    let txbx = wsp
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "txbx");
    let body_pr = wsp
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "bodyPr");

    let anchor = body_pr
        .and_then(|b| b.attribute("anchor"))
        .map(|s| s.to_string());
    let emu_to_pt = |v: &str| v.parse::<f64>().ok().map(|e| e / 12700.0).unwrap_or(0.0);
    // ECMA-376 §21.1.2.1.1 defaults: lIns=rIns=91440 EMU, tIns=bIns=45720 EMU
    let l = body_pr
        .and_then(|b| b.attribute("lIns"))
        .map(emu_to_pt)
        .unwrap_or(91440.0 / 12700.0);
    let t = body_pr
        .and_then(|b| b.attribute("tIns"))
        .map(emu_to_pt)
        .unwrap_or(45720.0 / 12700.0);
    let r = body_pr
        .and_then(|b| b.attribute("rIns"))
        .map(emu_to_pt)
        .unwrap_or(91440.0 / 12700.0);
    let b = body_pr
        .and_then(|b| b.attribute("bIns"))
        .map(emu_to_pt)
        .unwrap_or(45720.0 / 12700.0);

    let blocks: Vec<ShapeText> = txbx
        .and_then(|t| {
            t.children()
                .find(|n| n.is_element() && n.tag_name().name() == "txbxContent")
        })
        .map(|content| {
            children_w_flat(content, "p")
                .into_iter()
                .filter_map(|p| extract_simple_paragraph_text(p, theme))
                .collect()
        })
        .unwrap_or_default();

    (blocks, anchor, l, t, r, b)
}

/// Reduce a <w:p> inside <w:txbxContent> to a single ShapeText. Pulls
/// formatting from the FIRST run encountered; ignores mixed-format runs.
fn extract_simple_paragraph_text(p: roxmltree::Node, theme: &ThemeColors) -> Option<ShapeText> {
    let mut text = String::new();
    let mut first_rpr: Option<roxmltree::Node> = None;
    for r in p
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "r")
    {
        if first_rpr.is_none() {
            first_rpr = child_w(r, "rPr");
        }
        for t in r
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "t")
        {
            if let Some(text_node) = t.text() {
                text.push_str(text_node);
            }
        }
    }
    if text.is_empty() {
        return None;
    }

    let alignment = child_w(p, "pPr")
        .and_then(|ppr| child_w(ppr, "jc"))
        .and_then(|jc| attr_w(jc, "val"))
        .unwrap_or_else(|| "left".to_string());

    let (font_size_pt, color, font_family, bold, italic) = if let Some(rpr) = first_rpr {
        let fmt = parse_run_fmt(rpr);
        (
            fmt.font_size.unwrap_or(DEFAULT_FONT_SIZE),
            fmt.color.clone(),
            theme
                .resolve_font_ref(fmt.font_family_ascii.clone())
                .or_else(|| theme.resolve_font_ref(fmt.font_family_east_asia.clone())),
            fmt.bold.unwrap_or(false),
            fmt.italic.unwrap_or(false),
        )
    } else {
        (DEFAULT_FONT_SIZE, None, None, false, false)
    };

    Some(ShapeText {
        text,
        font_size_pt,
        color,
        font_family,
        bold,
        italic,
        alignment: normalize_align(&alignment).to_string(),
    })
}

/// Parse a legacy VML `<w:pict>` text box (ECMA-376 Part 4 §14.1) into a
/// ShapeRun. Handles the common Word-emitted form:
///   <w:pict><v:shape type="#_x0000_t202"
///       style="…width:300pt;height:60pt…" fillcolor="#fdf2d0"
///       strokecolor="#c0a000">
///     <v:textbox><w:txbxContent>…</w:txbxContent></v:textbox>
///   </v:shape></w:pict>
/// Size comes from the CSS-like `style` attribute (width/height in pt), fill
/// and stroke from `fillcolor`/`strokecolor`. The shape is anchored to the
/// paragraph's leading (top-left) corner — VML `position:relative` text boxes
/// flow with their anchor paragraph, which Word places at the left margin just
/// below the preceding content.
fn parse_vml_pict(pict: roxmltree::Node, theme: &ThemeColors) -> Option<ShapeRun> {
    // v:shape / v:rect / v:roundrect — any VML shape element with geometry.
    let shape = pict.descendants().find(|n| {
        n.is_element() && matches!(n.tag_name().name(), "shape" | "rect" | "roundrect" | "oval")
    })?;

    // CSS-like `style`: "position:relative;width:300pt;height:60pt;…"
    let style = shape.attribute("style").unwrap_or("");
    let css_pt = |prop: &str| -> Option<f64> {
        for decl in style.split(';') {
            let mut kv = decl.splitn(2, ':');
            let k = kv.next()?.trim();
            let v = kv.next()?.trim();
            if k.eq_ignore_ascii_case(prop) {
                // strip a trailing "pt" unit; VML lengths default to pt here.
                let num = v.trim_end_matches("pt").trim();
                return num.parse::<f64>().ok();
            }
        }
        None
    };
    let width_pt = css_pt("width").unwrap_or(0.0);
    let height_pt = css_pt("height").unwrap_or(0.0);
    if width_pt <= 0.0 || height_pt <= 0.0 {
        return None;
    }

    // VML colors: `fillcolor` / `strokecolor` are "#rrggbb" (or named); we keep
    // the 6-hex form the renderer expects (no leading '#').
    let hex6 = |c: &str| -> Option<String> {
        let s = c.trim().trim_start_matches('#');
        if s.len() == 6 && s.chars().all(|ch| ch.is_ascii_hexdigit()) {
            Some(s.to_ascii_lowercase())
        } else {
            None
        }
    };
    let fill = shape
        .attribute("fillcolor")
        .and_then(hex6)
        .map(|color| ShapeFill::Solid { color });
    let stroke = shape.attribute("strokecolor").and_then(hex6);
    // VML default stroke weight is 0.75pt when a stroke color is present and no
    // explicit weight is given (Part 4 §14.1.2.21 strokeweight default).
    let stroke_width = if stroke.is_some() {
        shape
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "stroke")
            .and_then(|n| n.attribute("weight"))
            .and_then(|w| w.trim_end_matches("pt").trim().parse::<f64>().ok())
            .or_else(|| {
                shape
                    .attribute("strokeweight")
                    .and_then(|w| w.trim_end_matches("pt").trim().parse::<f64>().ok())
            })
            .unwrap_or(0.75)
    } else {
        0.0
    };

    // Body text from <v:textbox><w:txbxContent>.
    let text_blocks: Vec<ShapeText> = shape
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "txbxContent")
        .map(|content| {
            children_w_flat(content, "p")
                .into_iter()
                .filter_map(|p| extract_simple_paragraph_text(p, theme))
                .collect()
        })
        .unwrap_or_default();

    Some(ShapeRun {
        width_pt,
        height_pt,
        anchor_x_pt: 0.0,
        anchor_y_pt: 0.0,
        anchor_x_from_margin: true,
        anchor_y_from_para: true,
        behind_doc: false,
        z_order: 0,
        subpaths: Vec::new(),
        preset_geometry: Some("rect".to_string()),
        adj_values: Vec::new(),
        fill,
        stroke,
        stroke_width,
        rotation: 0.0,
        // VML t202 text-box default insets are the OOXML defaults (§21.1.2.1.1).
        text_blocks,
        text_anchor: None,
        text_inset_l: 91440.0 / 12700.0,
        text_inset_t: 45720.0 / 12700.0,
        text_inset_r: 91440.0 / 12700.0,
        text_inset_b: 45720.0 / 12700.0,
        ..Default::default()
    })
}

/// Result of inspecting a shape's spPr for a direct fill.
///
/// ECMA-376 §20.1.8.44 (noFill): an explicit `<a:noFill/>` is a direct fill
/// property and therefore overrides the shape's style reference. We must
/// distinguish it from "no fill element present" (where the wps:style/fillRef
/// recipe applies). Collapsing both into `None` makes a no-fill shape pick up
/// the theme gradient referenced by fillRef, which is wrong.
enum FillSpec {
    /// A direct solidFill/gradFill was present and resolved.
    Explicit(ShapeFill),
    /// An explicit `<a:noFill/>` — the shape is intentionally unfilled.
    NoFill,
    /// No direct fill element at all — defer to wps:style/fillRef.
    Absent,
}

/// Parse a shape's direct fill (solidFill / gradFill / noFill), reporting which
/// of the three states applies so the caller can decide whether to fall back to
/// the style's fillRef.
fn parse_shape_fill(sp_pr: roxmltree::Node, theme: &ThemeColors) -> FillSpec {
    for child in sp_pr.children().filter(|n| n.is_element()) {
        match child.tag_name().name() {
            "solidFill" => {
                return match resolve_color_element(child, theme) {
                    Some(c) => FillSpec::Explicit(ShapeFill::Solid { color: c }),
                    // A solidFill whose color failed to resolve is still an
                    // explicit, direct fill declaration — do not fall back to
                    // the style reference.
                    None => FillSpec::NoFill,
                };
            }
            "gradFill" => {
                return match parse_grad_fill(child, theme) {
                    Some(f) => FillSpec::Explicit(f),
                    None => FillSpec::NoFill,
                };
            }
            "noFill" => return FillSpec::NoFill,
            _ => {}
        }
    }
    FillSpec::Absent
}

/// Resolve a wps:style/a:fillRef into a concrete ShapeFill using the theme's
/// fmtScheme/fillStyleLst (idx 1..) or bgFillStyleLst (idx 1000+). The fillRef
/// also carries a `<a:schemeClr>` child whose name substitutes for `phClr`
/// placeholders in the recipe (ECMA-376 §20.1.4.1.7 / §20.1.4.1.30).
///
/// Resume / cover templates lean on this: their backgrounds are described
/// indirectly as `fillRef idx="1003"` with the actual color and gradient
/// parameters living in the theme part. Without this lookup, the shape ends
/// up with no fill and the cover panel renders blank.
fn resolve_fill_ref(fill_ref: roxmltree::Node, theme: &ThemeColors) -> Option<ShapeFill> {
    let idx: u32 = fill_ref.attribute("idx")?.parse().ok()?;
    if idx == 0 {
        return None;
    }
    let scheme_clr = fill_ref
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "schemeClr")
        .and_then(|n| n.attribute("val"))
        .unwrap_or("dk1");

    let xml = theme.theme_xml.as_ref()?;
    let doc = XmlDoc::parse(xml).ok()?;
    let fmt = doc
        .root_element()
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "fmtScheme")?;
    // ECMA-376 §20.1.4.1.30: fillRef idx is 1-indexed.
    //   idx 1..999  → fillStyleLst[idx - 1]
    //   idx 1001+   → bgFillStyleLst[idx - 1001]
    let (lst_name, local_idx) = if idx >= 1001 {
        ("bgFillStyleLst", (idx - 1001) as usize)
    } else {
        ("fillStyleLst", (idx - 1) as usize)
    };
    let lst = fmt
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == lst_name)?;
    let entry = lst.children().filter(|n| n.is_element()).nth(local_idx)?;

    match entry.tag_name().name() {
        "solidFill" => resolve_color_element_with_phclr(entry, theme, scheme_clr)
            .map(|c| ShapeFill::Solid { color: c }),
        "gradFill" => parse_grad_fill_phclr(entry, theme, scheme_clr),
        // blipFill / pattFill recipes aren't supported yet — fall back to no fill.
        _ => None,
    }
}

fn parse_grad_fill(node: roxmltree::Node, theme: &ThemeColors) -> Option<ShapeFill> {
    parse_grad_fill_phclr(node, theme, "")
}

fn parse_grad_fill_phclr(
    node: roxmltree::Node,
    theme: &ThemeColors,
    ph_clr: &str,
) -> Option<ShapeFill> {
    let gs_lst = node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "gsLst")?;
    let mut stops: Vec<GradientStop> = Vec::new();
    for gs in gs_lst
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "gs")
    {
        let pos = gs
            .attribute("pos")
            .and_then(|v| v.parse::<f64>().ok())
            .map(|p| p / 100000.0)
            .unwrap_or(0.0);
        if let Some(color) = resolve_color_element_with_phclr(gs, theme, ph_clr) {
            stops.push(GradientStop {
                position: pos,
                color,
            });
        }
    }
    if stops.is_empty() {
        return None;
    }

    // Linear direction (a:lin ang = "60000"ths of a degree)
    let (angle, grad_type) = if let Some(lin) = node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "lin")
    {
        let ang = lin
            .attribute("ang")
            .and_then(|v| v.parse::<f64>().ok())
            .map(|a| a / 60000.0)
            .unwrap_or(0.0);
        (ang, "linear".to_string())
    } else if node
        .children()
        .any(|n| n.is_element() && n.tag_name().name() == "path")
    {
        (0.0, "radial".to_string())
    } else {
        (0.0, "linear".to_string())
    };

    Some(ShapeFill::Gradient {
        stops,
        angle,
        grad_type,
    })
}

/// Parse <a:custGeom><a:pathLst><a:path w="W" h="H">...</a:path></a:pathLst>.
/// Path coords inside each <a:path> are absolute within W×H; normalize to [0,1].
fn parse_custom_geometry(cust_geom: roxmltree::Node) -> Vec<Vec<PathCmd>> {
    let Some(path_lst) = cust_geom
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "pathLst")
    else {
        return vec![];
    };
    let mut subpaths: Vec<Vec<PathCmd>> = Vec::new();
    for path in path_lst
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "path")
    {
        let pw = path
            .attribute("w")
            .and_then(|v| v.parse::<f64>().ok())
            .unwrap_or(0.0);
        let ph = path
            .attribute("h")
            .and_then(|v| v.parse::<f64>().ok())
            .unwrap_or(0.0);
        if pw <= 0.0 || ph <= 0.0 {
            continue;
        }
        let mut cmds: Vec<PathCmd> = Vec::new();
        for cmd in path.children().filter(|n| n.is_element()) {
            let name = cmd.tag_name().name();
            let pts: Vec<(f64, f64)> = cmd
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "pt")
                .filter_map(|p| {
                    let x = p.attribute("x").and_then(|v| v.parse::<f64>().ok())?;
                    let y = p.attribute("y").and_then(|v| v.parse::<f64>().ok())?;
                    Some((x / pw, y / ph))
                })
                .collect();
            match name {
                "moveTo" => {
                    if let Some(p) = pts.first() {
                        cmds.push(PathCmd::MoveTo { x: p.0, y: p.1 });
                    }
                }
                "lnTo" => {
                    if let Some(p) = pts.first() {
                        cmds.push(PathCmd::LineTo { x: p.0, y: p.1 });
                    }
                }
                "cubicBezTo" => {
                    if pts.len() >= 3 {
                        cmds.push(PathCmd::CubicBezTo {
                            x1: pts[0].0,
                            y1: pts[0].1,
                            x2: pts[1].0,
                            y2: pts[1].1,
                            x: pts[2].0,
                            y: pts[2].1,
                        });
                    }
                }
                "arcTo" => {
                    let wr = cmd
                        .attribute("wR")
                        .and_then(|v| v.parse::<f64>().ok())
                        .unwrap_or(0.0)
                        / pw;
                    let hr = cmd
                        .attribute("hR")
                        .and_then(|v| v.parse::<f64>().ok())
                        .unwrap_or(0.0)
                        / ph;
                    let st_ang = cmd
                        .attribute("stAng")
                        .and_then(|v| v.parse::<f64>().ok())
                        .unwrap_or(0.0)
                        / 60000.0;
                    let sw_ang = cmd
                        .attribute("swAng")
                        .and_then(|v| v.parse::<f64>().ok())
                        .unwrap_or(0.0)
                        / 60000.0;
                    cmds.push(PathCmd::ArcTo {
                        wr,
                        hr,
                        st_ang,
                        sw_ang,
                    });
                }
                "close" => cmds.push(PathCmd::Close),
                _ => {}
            }
        }
        if !cmds.is_empty() {
            subpaths.push(cmds);
        }
    }
    subpaths
}

/// Resolve a color container (e.g. <a:solidFill>, <a:gs>) into a hex string by
/// inspecting its child: <a:srgbClr>, <a:schemeClr>, or <a:sysClr>, applying
/// any lumMod/lumOff/alpha modifiers declared on the inner color element.
fn resolve_color_element(container: roxmltree::Node, theme: &ThemeColors) -> Option<String> {
    resolve_color_element_with_phclr(container, theme, "")
}

/// Like `resolve_color_element` but, when an inner `<a:schemeClr val="phClr"/>`
/// is encountered, substitutes the scheme name `ph_clr` (the `<a:schemeClr>`
/// child of the wps:style/fillRef that triggered theme lookup). Pass an empty
/// string to disable the substitution.
fn resolve_color_element_with_phclr(
    container: roxmltree::Node,
    theme: &ThemeColors,
    ph_clr: &str,
) -> Option<String> {
    for c in container.children().filter(|n| n.is_element()) {
        let base = match c.tag_name().name() {
            "srgbClr" => c.attribute("val").map(|v| v.to_uppercase()),
            "schemeClr" => {
                let raw_name = c.attribute("val")?;
                let name = if raw_name == "phClr" && !ph_clr.is_empty() {
                    ph_clr
                } else {
                    raw_name
                };
                theme.resolve(name)
            }
            "sysClr" => c
                .attribute("lastClr")
                .map(|v| v.to_uppercase())
                .or_else(|| c.attribute("val").map(|v| v.to_uppercase())),
            _ => None,
        };
        let Some(hex) = base else { continue };
        return Some(apply_color_mods(&hex, c));
    }
    None
}

/// Word's interpretation of OOXML color modifiers. Wraps the shared
/// `ooxml_common::color::apply_color_transforms` with TintMode::WordLiteral
/// — the spec-literal `tint = val·input + (1-val)·white` formulation. See
/// the comment in `ooxml-common/src/color.rs` for why this differs from
/// PowerPoint's linear-sRGB tint.
fn apply_color_mods(hex: &str, color_node: roxmltree::Node) -> String {
    ooxml_common::color::apply_color_transforms(
        hex,
        color_node,
        ooxml_common::color::TintMode::WordLiteral,
    )
}

// ===== Table parsing =====

fn parse_table(
    node: roxmltree::Node,
    style_map: &StyleMap,
    num_map: &mut NumberingMap,
    media_map: &HashMap<String, String>,
    rel_map: &HashMap<String, String>,
    theme: &ThemeColors,
) -> DocTable {
    let tbl_pr = child_w(node, "tblPr");
    let tbl_grid = child_w(node, "tblGrid");

    // Resolve the table style ID (§17.4.63 w:tblStyle). Cell paragraphs
    // inherit this style's pPr — e.g. "Table Grid" (style `af3` in this
    // sample) sets line="240" after="0", which tightens cell line spacing
    // below the body-text default inherited from docDefault.
    let table_style_id = tbl_pr
        .and_then(|p| child_w(p, "tblStyle"))
        .and_then(|s| attr_w(s, "val"));

    // Column widths from tblGrid
    let col_widths: Vec<f64> = tbl_grid
        .map(|g| {
            children_w(g, "gridCol")
                .iter()
                .map(|c| attr_w(*c, "w").map(|v| twips_to_pt(&v)).unwrap_or(72.0))
                .collect()
        })
        .unwrap_or_default();

    // Resolve the table style's cell/border formatting (shading, banding, borders,
    // vAlign) — these live in styles.xml, not inline (§17.7.6). tblLook selects which
    // conditional formats are active.
    let tstyle = table_style_id
        .as_deref()
        .map(|id| style_map.resolve_table_style(id))
        .unwrap_or_default();
    let look = tbl_pr.and_then(|p| child_w(p, "tblLook"));
    let look_flag = |name: &str| look.and_then(|l| attr_w(l, name)).as_deref() == Some("1");
    let first_row = look_flag("firstRow");
    let h_band = look.and_then(|l| attr_w(l, "noHBand")).as_deref() != Some("1");

    // Table borders: inline tblBorders win; otherwise the table style's borders
    // (so styles like "Table Grid" show their gridlines).
    let mut borders = tbl_pr
        .and_then(|p| child_w(p, "tblBorders"))
        .map(|b| parse_table_borders(b))
        .unwrap_or_default();
    apply_style_borders(&mut borders, &tstyle.borders);

    // Cell margins
    let (cm_top, cm_bot, cm_left, cm_right) = tbl_pr
        .and_then(|p| child_w(p, "tblCellMar"))
        .map(|m| {
            (
                child_w(m, "top")
                    .and_then(|n| attr_w(n, "w"))
                    .map(|v| twips_to_pt(&v))
                    .unwrap_or(0.0),
                child_w(m, "bottom")
                    .and_then(|n| attr_w(n, "w"))
                    .map(|v| twips_to_pt(&v))
                    .unwrap_or(0.0),
                child_w(m, "left")
                    .and_then(|n| attr_w(n, "w"))
                    .map(|v| twips_to_pt(&v))
                    .unwrap_or(3.6),
                child_w(m, "right")
                    .and_then(|n| attr_w(n, "w"))
                    .map(|v| twips_to_pt(&v))
                    .unwrap_or(3.6),
            )
        })
        .unwrap_or((0.0, 0.0, 3.6, 3.6));

    let mut rows = vec![];
    let mut row_cnf: Vec<Option<String>> = vec![];
    for tr_node in children_w_flat(node, "tr") {
        // §17.4.7 conditional-format bitmask on the row (firstRow/band1Horz/…).
        row_cnf.push(
            child_w(tr_node, "trPr")
                .and_then(|p| child_w(p, "cnfStyle"))
                .and_then(|c| attr_w(c, "val")),
        );
        let row = parse_table_row(
            tr_node,
            style_map,
            num_map,
            media_map,
            rel_map,
            theme,
            table_style_id.as_deref(),
        );
        rows.push(row);
    }

    // Apply table-style cell shading + vAlign where the cell didn't set them inline.
    // Only treat row 0 as a non-banded "first row" when firstRow is enabled AND the
    // style's firstRow conditional actually carries formatting; an empty firstRow
    // (like EHC) must NOT shift the banding (Word bands from row 0 → 1st row = band1).
    let first_row_styled = first_row
        && tstyle
            .cond
            .get("firstRow")
            .map(|c| c.shd.is_some())
            .unwrap_or(false);
    for (r, row) in rows.iter_mut().enumerate() {
        // Pick the conditional format: the row's explicit cnfStyle wins (§17.4.7);
        // otherwise fall back to tblLook firstRow + horizontal banding by row parity.
        let cond_name: Option<String> = if let Some(cnf) = &row_cnf[r] {
            cnf_to_cond(cnf)
        } else if r == 0 && first_row_styled {
            Some("firstRow".to_string())
        } else if h_band {
            let bi = if first_row_styled {
                r as i64 - 1
            } else {
                r as i64
            };
            Some(
                if bi % 2 == 0 {
                    "band1Horz"
                } else {
                    "band2Horz"
                }
                .to_string(),
            )
        } else {
            None
        };
        let cond: Option<&CondFmt> = cond_name.as_deref().and_then(|n| tstyle.cond.get(n));
        let row_shd = cond
            .and_then(|c| c.shd.clone())
            .or_else(|| tstyle.cell_shd.clone());
        for cell in row.cells.iter_mut() {
            if cell.background.is_none() {
                cell.background = row_shd.clone();
            }
            if cell.v_align.is_empty() {
                cell.v_align = tstyle
                    .cell_valign
                    .clone()
                    .unwrap_or_else(|| "top".to_string());
            }
        }
    }

    let jc = tbl_pr
        .and_then(|p| child_w(p, "jc"))
        .and_then(|j| attr_w(j, "val"))
        .unwrap_or_else(|| "left".to_string());

    // ECMA-376 §17.4.52 w:tblLayout/@w:type — "fixed" | "autofit". Absent ⇒ None
    // (renderer applies the spec default "autofit").
    let layout = tbl_pr
        .and_then(|p| child_w(p, "tblLayout"))
        .and_then(|l| attr_w(l, "type"));

    // ECMA-376 §17.4.63 w:tblW (ST_TblWidth): dxa ⇒ width_pt, pct ⇒ width_pct
    // (50ths of a percent of available width), auto/nil/0 ⇒ both None.
    let (tbl_w_pt, tbl_w_pct) = tbl_pr
        .and_then(|p| child_w(p, "tblW"))
        .map(|w| {
            let wtype = attr_w(w, "type").unwrap_or_else(|| "dxa".to_string());
            match wtype.as_str() {
                "dxa" => (
                    attr_w(w, "w").map(|v| twips_to_pt(&v)).filter(|v| *v > 0.0),
                    None,
                ),
                "pct" => (
                    None,
                    attr_w(w, "w")
                        .and_then(|v| v.parse::<f64>().ok())
                        .filter(|v| *v > 0.0),
                ),
                _ => (None, None),
            }
        })
        .unwrap_or((None, None));

    // ECMA-376 §17.4.1 w:bidiVisual — RTL (visual) column order. On-off toggle.
    // The renderer mirrors the grid (logical column 0 rightmost) and flips
    // per-cell left/right borders when this is set.
    let bidi_visual = tbl_pr.and_then(|p| bool_prop(p, "bidiVisual"));

    DocTable {
        col_widths,
        rows,
        borders,
        cell_margin_top: cm_top,
        cell_margin_bottom: cm_bot,
        cell_margin_left: cm_left,
        cell_margin_right: cm_right,
        jc,
        layout,
        width_pt: tbl_w_pt,
        width_pct: tbl_w_pct,
        bidi_visual,
    }
}

fn parse_table_row(
    node: roxmltree::Node,
    style_map: &StyleMap,
    num_map: &mut NumberingMap,
    media_map: &HashMap<String, String>,
    rel_map: &HashMap<String, String>,
    theme: &ThemeColors,
    table_style_id: Option<&str>,
) -> DocTableRow {
    let tr_pr = child_w(node, "trPr");
    let tr_height_node = tr_pr.and_then(|p| child_w(p, "trHeight"));
    let row_height = tr_height_node
        .and_then(|h| attr_w(h, "val"))
        .map(|v| twips_to_pt(&v));
    // ECMA-376 §17.4.80: default hRule is "auto".
    let row_height_rule = tr_height_node
        .and_then(|h| attr_w(h, "hRule"))
        .unwrap_or_else(|| "auto".to_string());
    let is_header = tr_pr.and_then(|p| child_w(p, "tblHeader")).is_some();

    let mut cells = vec![];
    for tc_node in children_w_flat(node, "tc") {
        let cell = parse_table_cell(
            tc_node,
            style_map,
            num_map,
            media_map,
            rel_map,
            theme,
            table_style_id,
        );
        cells.push(cell);
    }

    DocTableRow {
        cells,
        row_height,
        row_height_rule,
        is_header,
    }
}

fn parse_table_cell(
    node: roxmltree::Node,
    style_map: &StyleMap,
    num_map: &mut NumberingMap,
    media_map: &HashMap<String, String>,
    rel_map: &HashMap<String, String>,
    theme: &ThemeColors,
    table_style_id: Option<&str>,
) -> DocTableCell {
    let tc_pr = child_w(node, "tcPr");

    let col_span = tc_pr
        .and_then(|p| child_w(p, "gridSpan"))
        .and_then(|g| attr_w(g, "val"))
        .and_then(|v| v.parse().ok())
        .unwrap_or(1);

    // ECMA-376 §17.4.85: ST_Merge default is "continue", so a <w:vMerge/>
    // element with no val attribute means the cell continues the merged
    // region from the row above. Only val="restart" begins a new merge.
    let v_merge = tc_pr
        .and_then(|p| child_w(p, "vMerge"))
        .map(|m| attr_w(m, "val").map(|v| v == "restart").unwrap_or(false));

    let borders = tc_pr
        .and_then(|p| child_w(p, "tcBorders"))
        .map(|b| parse_cell_borders(b))
        .unwrap_or_default();

    let background = tc_pr
        .and_then(|p| child_w(p, "shd"))
        .and_then(|s| attr_w(s, "fill"))
        .filter(|f| f != "auto" && f.len() == 6)
        .map(|f| f.to_lowercase());

    // Empty = not set inline; parse_table fills it from the table style (else "top").
    let v_align = tc_pr
        .and_then(|p| child_w(p, "vAlign"))
        .and_then(|v| attr_w(v, "val"))
        .unwrap_or_default();

    // ECMA-376 §17.4.71 (w:tcW) + §17.18.90 ST_TblWidth:
    //   dxa  — twentieths of a point (1/20pt) ⇒ width_pt.
    //   pct  — 50ths of a percent of the available content width (e.g.
    //          w="2500" = 50%). Resolved in the renderer where the available
    //          width is known ⇒ width_pct.
    //   auto, nil — no width preference; both None (column falls back to grid).
    let tc_w = tc_pr.and_then(|p| child_w(p, "tcW"));
    let (width_pt, width_pct) = tc_w
        .map(|w| {
            let wtype = attr_w(w, "type").unwrap_or_else(|| "dxa".to_string());
            match wtype.as_str() {
                "dxa" => (attr_w(w, "w").map(|v| twips_to_pt(&v)), None),
                "pct" => (None, attr_w(w, "w").and_then(|v| v.parse::<f64>().ok())),
                _ => (None, None),
            }
        })
        .unwrap_or((None, None));

    // Per-cell margins (ECMA-376 §17.4.42 `<w:tcPr><w:tcMar>`). Each edge,
    // when present, overrides the table-level `<w:tblCellMar>` default; absent
    // edges stay None so the renderer falls back to the table default.
    let tc_mar = tc_pr.and_then(|p| child_w(p, "tcMar"));
    let edge_mar = |name: &str| {
        tc_mar
            .and_then(|m| child_w(m, name))
            .and_then(|n| attr_w(n, "w"))
            .map(|v| twips_to_pt(&v))
    };
    let margin_top = edge_mar("top");
    let margin_bottom = edge_mar("bottom");
    let margin_left = edge_mar("left");
    let margin_right = edge_mar("right");

    // ECMA-376 §17.4.7: a cell may contain paragraphs AND nested tables in
    // any order. element_children_flat unwraps sdt wrappers like elsewhere.
    let mut content: Vec<CellElement> = vec![];
    for child in element_children_flat(node) {
        match child.tag_name().name() {
            "p" => content.push(CellElement::Paragraph(parse_paragraph(
                child,
                style_map,
                num_map,
                media_map,
                rel_map,
                theme,
                table_style_id,
            ))),
            "tbl" => content.push(CellElement::Table(parse_table(
                child, style_map, num_map, media_map, rel_map, theme,
            ))),
            _ => {}
        }
    }

    DocTableCell {
        content,
        col_span,
        v_merge,
        borders,
        background,
        v_align,
        width_pt,
        width_pct,
        margin_top,
        margin_bottom,
        margin_left,
        margin_right,
    }
}

fn parse_table_borders(node: roxmltree::Node) -> TableBorders {
    // ECMA-376 §17.4.* tblBorders: the vertical edges may be given either as the
    // legacy physical names (w:left/w:right) or as the logical edges
    // (w:start/w:end, §17.4.66-67). Prefer the physical name when both are
    // present, else fall back to the logical alias. The renderer handles
    // visual mirroring for bidiVisual tables via its `mirror` flag, so at
    // parse time start→left and end→right unconditionally.
    TableBorders {
        top: child_w(node, "top").map(parse_border_spec),
        bottom: child_w(node, "bottom").map(parse_border_spec),
        left: child_w(node, "left")
            .or_else(|| child_w(node, "start"))
            .map(parse_border_spec),
        right: child_w(node, "right")
            .or_else(|| child_w(node, "end"))
            .map(parse_border_spec),
        inside_h: child_w(node, "insideH").map(parse_border_spec),
        inside_v: child_w(node, "insideV").map(parse_border_spec),
    }
}

fn parse_cell_borders(node: roxmltree::Node) -> CellBorders {
    // ECMA-376 §17.4.* tcBorders: same logical/physical edge aliasing as
    // tblBorders. w:start/w:end (§17.4.66-67) are the logical vertical edges;
    // w:left/w:right the physical names. Prefer physical, fall back to logical.
    CellBorders {
        top: child_w(node, "top").map(parse_border_spec),
        bottom: child_w(node, "bottom").map(parse_border_spec),
        left: child_w(node, "left")
            .or_else(|| child_w(node, "start"))
            .map(parse_border_spec),
        right: child_w(node, "right")
            .or_else(|| child_w(node, "end"))
            .map(parse_border_spec),
    }
}

/// Decode a `w:cnfStyle` bitmask (12 chars) to the conditional-format key it selects.
/// Bit order (§17.4.7): firstRow,lastRow,firstCol,lastCol,band1Vert,band2Vert,
/// band1Horz,band2Horz,neCell,nwCell,seCell,swCell.
fn cnf_to_cond(cnf: &str) -> Option<String> {
    let bit = |i: usize| cnf.as_bytes().get(i).copied() == Some(b'1');
    if bit(0) {
        Some("firstRow".to_string())
    } else if bit(1) {
        Some("lastRow".to_string())
    } else if bit(6) {
        Some("band1Horz".to_string())
    } else if bit(7) {
        Some("band2Horz".to_string())
    } else {
        None
    }
}

/// Fill table-border edges from a table style where the inline table didn't set them.
fn apply_style_borders(dst: &mut TableBorders, src: &RawTblBorders) {
    let usable = |e: &EdgeBorder| e.style != "none" && e.style != "nil";
    let conv = |e: &EdgeBorder| BorderSpec {
        width: e.width,
        color: e.color.clone(),
        style: e.style.clone(),
    };
    if dst.top.is_none() {
        if let Some(e) = &src.top {
            if usable(e) {
                dst.top = Some(conv(e));
            }
        }
    }
    if dst.bottom.is_none() {
        if let Some(e) = &src.bottom {
            if usable(e) {
                dst.bottom = Some(conv(e));
            }
        }
    }
    if dst.left.is_none() {
        if let Some(e) = &src.left {
            if usable(e) {
                dst.left = Some(conv(e));
            }
        }
    }
    if dst.right.is_none() {
        if let Some(e) = &src.right {
            if usable(e) {
                dst.right = Some(conv(e));
            }
        }
    }
    if dst.inside_h.is_none() {
        if let Some(e) = &src.inside_h {
            if usable(e) {
                dst.inside_h = Some(conv(e));
            }
        }
    }
    if dst.inside_v.is_none() {
        if let Some(e) = &src.inside_v {
            if usable(e) {
                dst.inside_v = Some(conv(e));
            }
        }
    }
}

fn parse_border_spec(node: roxmltree::Node) -> BorderSpec {
    let style = attr_w(node, "val").unwrap_or_else(|| "none".to_string());
    let width = attr_w(node, "sz")
        .map(|v| {
            v.parse::<f64>().unwrap_or(4.0) / 8.0 // eighth-points → pt
        })
        .unwrap_or(0.5);
    let color = attr_w(node, "color")
        .filter(|c| c != "auto")
        .map(|c| c.to_lowercase());
    BorderSpec {
        width,
        color,
        style,
    }
}

// ===== Helpers =====

fn normalize_align(s: &str) -> &str {
    match s {
        // "both" = justified with last line left-aligned (default Word justify).
        // "distribute" = justified including last line (CJK Distribute).
        // Keep them distinct so the renderer can decide whether to stretch
        // the last line. Legacy "justify" maps to the same behavior as "both".
        "both" | "justify" => "justify",
        "distribute" => "distribute",
        "right" | "end" => "right",
        "center" => "center",
        "start" | "left" => "left",
        _ => "left",
    }
}

fn apply_direct_para(base: &mut ParaFmt, direct: &ParaFmt) {
    if direct.alignment.is_some() {
        base.alignment = direct.alignment.clone();
    }
    if direct.indent_left.is_some() {
        base.indent_left = direct.indent_left;
    }
    if direct.indent_right.is_some() {
        base.indent_right = direct.indent_right;
    }
    if direct.indent_first.is_some() {
        base.indent_first = direct.indent_first;
    }
    if direct.space_before.is_some() {
        base.space_before = direct.space_before;
    }
    if direct.space_after.is_some() {
        base.space_after = direct.space_after;
    }
    if direct.line_spacing_val.is_some() {
        base.line_spacing_val = direct.line_spacing_val;
    }
    if direct.line_spacing_rule.is_some() {
        base.line_spacing_rule = direct.line_spacing_rule.clone();
    }
    if direct.line_spacing_explicit.is_some() {
        base.line_spacing_explicit = direct.line_spacing_explicit;
    }
    if direct.outline_level.is_some() {
        base.outline_level = direct.outline_level;
    }
    if direct.num_id.is_some() {
        base.num_id = direct.num_id;
    }
    if direct.num_level.is_some() {
        base.num_level = direct.num_level;
    }
    if direct.tab_stops.is_some() {
        base.tab_stops = direct.tab_stops.clone();
    }
    if direct.bidi.is_some() {
        base.bidi = direct.bidi;
    }
    if direct.snap_to_grid.is_some() {
        base.snap_to_grid = direct.snap_to_grid;
    }
}

fn apply_direct_run(base: &mut RunFmt, direct: &RunFmt) {
    if direct.bold.is_some() {
        base.bold = direct.bold;
    }
    if direct.italic.is_some() {
        base.italic = direct.italic;
    }
    if direct.underline.is_some() {
        base.underline = direct.underline;
    }
    if direct.strikethrough.is_some() {
        base.strikethrough = direct.strikethrough;
    }
    if direct.font_size.is_some() {
        base.font_size = direct.font_size;
    }
    if direct.color.is_some() {
        base.color = direct.color.clone();
    }
    if direct.font_family_ascii.is_some() {
        base.font_family_ascii = direct.font_family_ascii.clone();
    }
    if direct.font_family_east_asia.is_some() {
        base.font_family_east_asia = direct.font_family_east_asia.clone();
    }
    if direct.background.is_some() {
        base.background = direct.background.clone();
    }
    if direct.vert_align.is_some() {
        base.vert_align = direct.vert_align.clone();
    }
    if direct.rtl.is_some() {
        base.rtl = direct.rtl;
    }
    if direct.cs_toggle.is_some() {
        base.cs_toggle = direct.cs_toggle;
    }
    if direct.font_family_cs.is_some() {
        base.font_family_cs = direct.font_family_cs.clone();
    }
    if direct.bold_cs.is_some() {
        base.bold_cs = direct.bold_cs;
    }
    if direct.italic_cs.is_some() {
        base.italic_cs = direct.italic_cs;
    }
    if direct.lang_bidi.is_some() {
        base.lang_bidi = direct.lang_bidi.clone();
    }

    // Complex-script font size: a directly-applied `w:sz` without an
    // accompanying `w:szCs` mirrors into the cs size (ECMA-376 §17.3.2.18 —
    // see `styles::apply_run` for the full rationale). An explicit `w:szCs`
    // always wins; otherwise inherited-only szCs is carried unchanged.
    if direct.font_size_cs_set_here {
        base.font_size_cs = direct.font_size_cs;
    } else if direct.font_size_set_here {
        base.font_size_cs = direct.font_size;
    } else if direct.font_size_cs.is_some() {
        base.font_size_cs = direct.font_size_cs;
    }
    base.font_size_set_here = false;
    base.font_size_cs_set_here = false;
}

fn parse_rels(xml: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    if xml.is_empty() {
        return map;
    }
    let doc = match XmlDoc::parse(xml) {
        Ok(d) => d,
        Err(_) => return map,
    };
    for rel in doc
        .root_element()
        .children()
        .filter(|n| n.tag_name().name() == "Relationship")
    {
        if let (Some(id), Some(target)) = (rel.attribute("Id"), rel.attribute("Target")) {
            map.insert(id.to_string(), target.to_string());
        }
    }
    map
}

fn read_zip_entry(zip: &mut Zip, path: &str) -> Result<String, String> {
    let max = ooxml_common::zip::current_max();
    let mut entry = zip.by_name(path).map_err(|e| format!("{}: {}", path, e))?;
    if entry.size() > max {
        return Err(format!("{}: exceeds size limit", path));
    }
    let mut s = String::new();
    entry
        .by_ref()
        .take(max)
        .read_to_string(&mut s)
        .map_err(|e| e.to_string())?;
    Ok(s)
}

fn read_zip_bytes(zip: &mut Zip, path: &str) -> Result<Vec<u8>, String> {
    let max = ooxml_common::zip::current_max();
    let mut entry = zip.by_name(path).map_err(|e| format!("{}: {}", path, e))?;
    if entry.size() > max {
        return Err(format!("{}: exceeds size limit", path));
    }
    let mut buf = vec![];
    entry
        .by_ref()
        .take(max)
        .read_to_end(&mut buf)
        .map_err(|e| e.to_string())?;
    Ok(buf)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml_util::W_NS;

    fn parse_tbl(body: &str) -> DocTable {
        let xml = format!(r#"<w:tbl xmlns:w="{ns}">{body}</w:tbl>"#, ns = W_NS);
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let style_map = StyleMap::parse("");
        let mut num_map = NumberingMap::default();
        let media: HashMap<String, String> = HashMap::new();
        let rels: HashMap<String, String> = HashMap::new();
        let theme = ThemeColors::default();
        parse_table(
            doc.root_element(),
            &style_map,
            &mut num_map,
            &media,
            &rels,
            &theme,
        )
    }

    // ECMA-376 §17.4.71 — a cell's `<w:tcW>` preferred width (type="dxa") reaches
    // the model as `width_pt` so the renderer can size autofit columns by it.
    #[test]
    fn cell_tcw_dxa_surfaces_as_width_pt() {
        let t = parse_tbl(
            r#"<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
               <w:tblGrid><w:gridCol w:w="8306"/></w:tblGrid>
               <w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
                 <w:p/></w:tc></w:tr>"#,
        );
        let cell = &t.rows[0].cells[0];
        // 2000 twips = 100 pt.
        assert_eq!(cell.width_pt, Some(100.0));
        assert_eq!(cell.width_pct, None);
        // tblGrid still parsed independently (8306 twips ≈ 415.3 pt).
        assert!((t.col_widths[0] - 415.3).abs() < 0.1);
    }

    // ECMA-376 §17.4.71 + §17.18.90 — type="pct" carries no twips; it surfaces as
    // `width_pct` (50ths of a percent) for the renderer to resolve.
    #[test]
    fn cell_tcw_pct_surfaces_as_width_pct() {
        let t = parse_tbl(
            r#"<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
               <w:tr><w:tc><w:tcPr><w:tcW w:w="2500" w:type="pct"/></w:tcPr>
                 <w:p/></w:tc></w:tr>"#,
        );
        let cell = &t.rows[0].cells[0];
        assert_eq!(cell.width_pt, None);
        assert_eq!(cell.width_pct, Some(2500.0));
    }

    // ECMA-376 §17.4.52 — absent `<w:tblLayout>` ⇒ None (renderer default autofit).
    #[test]
    fn absent_tbl_layout_is_none() {
        let t = parse_tbl(
            r#"<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
               <w:tr><w:tc><w:p/></w:tc></w:tr>"#,
        );
        assert_eq!(t.layout, None);
    }

    // ECMA-376 §17.4.52 — explicit fixed layout reaches the model verbatim.
    #[test]
    fn fixed_tbl_layout_surfaces() {
        let t = parse_tbl(
            r#"<w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>
               <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
               <w:tr><w:tc><w:p/></w:tc></w:tr>"#,
        );
        assert_eq!(t.layout.as_deref(), Some("fixed"));
    }

    // ECMA-376 §17.4.63 — `<w:tblW w:type="auto" w:w="0">` carries no preference.
    #[test]
    fn tbl_w_auto_zero_is_none() {
        let t = parse_tbl(
            r#"<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
               <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
               <w:tr><w:tc><w:p/></w:tc></w:tr>"#,
        );
        assert_eq!(t.width_pt, None);
        assert_eq!(t.width_pct, None);
    }
}

#[cfg(test)]
mod wgp_image_tests {
    use super::*;
    use crate::xml_util::R_NS;

    // ECMA-376 §20.1.7.5/.6 — a picture nested inside a wpg:grpSp must be mapped
    // to page space through the *composed* group transforms (each grpSp's
    // off/ext/chOff/chExt scales and offsets its children). The old parser
    // applied only the pic's own offset, ignoring the group's scale/offset and
    // any nested grpSp transform.
    #[test]
    fn nested_group_pic_composes_transform() {
        // Outer wgp: identity (off 0, ext == chExt). Inner grpSp: offset
        // (69850, 295275) EMU and 2× scale (ext 2000 over chExt 1000). Pic at
        // child-coord off (100000, 200000) EMU, ext (127000, 254000).
        let xml = format!(
            r#"<wpg:wgp xmlns:wpg="urn:wpg" xmlns:a="urn:a" xmlns:pic="urn:pic" xmlns:r="{r}">
                 <wpg:grpSpPr><a:xfrm>
                   <a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/>
                   <a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/>
                 </a:xfrm></wpg:grpSpPr>
                 <wpg:grpSp>
                   <wpg:grpSpPr><a:xfrm>
                     <a:off x="69850" y="295275"/><a:ext cx="2000" cy="2000"/>
                     <a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/>
                   </a:xfrm></wpg:grpSpPr>
                   <pic:pic>
                     <pic:spPr><a:xfrm>
                       <a:off x="100000" y="200000"/><a:ext cx="127000" cy="254000"/>
                     </a:xfrm></pic:spPr>
                     <pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill>
                   </pic:pic>
                 </wpg:grpSp>
               </wpg:wgp>"#,
            r = R_NS
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let mut media = HashMap::new();
        media.insert("rId1".to_string(), "data:image/png;base64,AAAA".to_string());
        let meta = AnchorMeta::default();
        let imgs = parse_wgp_images(
            doc.root_element(),
            &media,
            0.0,   // anchor_pos_x
            false, // x_from_margin
            0.0,   // anchor_pos_y
            false, // y_from_para
            &meta,
        );
        assert_eq!(imgs.len(), 1);
        let img = &imgs[0];
        // size = pic ext × inner scale(2): 127000*2/12700 = 20pt, 254000*2/12700 = 40pt.
        assert!(
            (img.width_pt - 20.0).abs() < 1e-6,
            "width_pt = {}",
            img.width_pt
        );
        assert!(
            (img.height_pt - 40.0).abs() < 1e-6,
            "height_pt = {}",
            img.height_pt
        );
        // pos = group off + pic off × scale, all /12700:
        //   x = (69850 + 100000*2)/12700 = 269850/12700 ≈ 21.2480
        //   y = (295275 + 200000*2)/12700 = 695275/12700 ≈ 54.7461
        // Pre-fix (buggy) values were x = 100000/12700 ≈ 7.874, width = 10pt.
        assert!(
            (img.anchor_x_pt - 21.248_031).abs() < 1e-3,
            "anchor_x_pt = {}",
            img.anchor_x_pt
        );
        assert!(
            (img.anchor_y_pt - 54.746_063).abs() < 1e-3,
            "anchor_y_pt = {}",
            img.anchor_y_pt
        );
    }
}

#[cfg(test)]
mod allow_overlap_tests {
    use super::*;

    fn meta(anchor_attrs: &str) -> AnchorMeta {
        let xml = format!(
            r#"<wp:anchor xmlns:wp="urn:wp" {attrs}>
                 <wp:wrapSquare wrapText="bothSides"/>
               </wp:anchor>"#,
            attrs = anchor_attrs
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        parse_anchor_wrap(&doc.root_element())
    }

    // ECMA-376 §20.4.2.3 — @allowOverlap is optional; when omitted the object
    // MAY overlap (default true).
    #[test]
    fn omitted_allow_overlap_defaults_true() {
        assert!(meta("").allow_overlap);
    }

    // §20.4.2.3 — allowOverlap="0" / "false" forbids overlap (false).
    #[test]
    fn allow_overlap_false_is_false() {
        assert!(!meta(r#"allowOverlap="0""#).allow_overlap);
        assert!(!meta(r#"allowOverlap="false""#).allow_overlap);
    }

    // §20.4.2.3 — allowOverlap="1" / "true" permits overlap (true).
    #[test]
    fn allow_overlap_true_is_true() {
        assert!(meta(r#"allowOverlap="1""#).allow_overlap);
        assert!(meta(r#"allowOverlap="true""#).allow_overlap);
    }

    // The hand-written Default must preserve the spec default of true so any
    // AnchorMeta::default() (e.g. group images without an anchor) is correct.
    #[test]
    fn default_anchor_meta_allows_overlap() {
        assert!(AnchorMeta::default().allow_overlap);
    }
}

#[cfg(test)]
mod math_jc_tests {
    use super::*;
    use crate::xml_util::W_NS;

    const M_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    /// Build a `<w:p>` declaring both the w and m namespaces and run it through
    /// `parse_paragraph`, analogous to `parse_tbl` above.
    fn parse_p(inner: &str) -> DocParagraph {
        let xml = format!(
            r#"<w:p xmlns:w="{w}" xmlns:m="{m}">{inner}</w:p>"#,
            w = W_NS,
            m = M_NS
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let style_map = StyleMap::parse("");
        let mut num_map = NumberingMap::default();
        let media: HashMap<String, String> = HashMap::new();
        let rels: HashMap<String, String> = HashMap::new();
        let theme = ThemeColors::default();
        parse_paragraph(
            doc.root_element(),
            &style_map,
            &mut num_map,
            &media,
            &rels,
            &theme,
            None,
        )
    }

    fn first_math_jc(p: &DocParagraph) -> Option<&Option<String>> {
        p.runs.iter().find_map(|r| match r {
            DocRun::Math { jc, .. } => Some(jc),
            _ => None,
        })
    }

    // ECMA-376 §22.1.2.88 `m:jc` — the per-instance justification on
    // `m:oMathPara/m:oMathParaPr` reaches the display Math run.
    #[test]
    fn omathpara_jc_left_sets_run_jc() {
        let p = parse_p(
            r#"<m:oMathPara><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr>
               <m:oMath><m:r><m:t>α</m:t></m:r></m:oMath></m:oMathPara>"#,
        );
        assert_eq!(first_math_jc(&p), Some(&Some("left".to_string())));
    }

    // ECMA-376 §22.1.2.88 — absent `m:oMathParaPr` ⇒ no per-instance jc; the
    // document default (`m:defJc`) resolution is left to the renderer.
    #[test]
    fn omathpara_without_jc_is_none() {
        let p = parse_p(r#"<m:oMathPara><m:oMath><m:r><m:t>α</m:t></m:r></m:oMath></m:oMathPara>"#);
        assert_eq!(first_math_jc(&p), Some(&None));
    }

    // ECMA-376 §22.1.2.77 `m:oMath` — inline math carries no oMathPara jc.
    #[test]
    fn inline_omath_jc_is_none() {
        let p = parse_p(r#"<m:oMath><m:r><m:t>α</m:t></m:r></m:oMath>"#);
        assert_eq!(first_math_jc(&p), Some(&None));
    }

    // ECMA-376 §22.1.2.30 `m:defJc` — document-wide default math justification in
    // `word/settings.xml` `m:mathPr` surfaces as `math_def_jc`; absent ⇒ None.
    #[test]
    fn settings_defjc_surfaces() {
        let xml = format!(
            r#"<w:settings xmlns:w="{w}" xmlns:m="{m}">
                 <m:mathPr><m:defJc m:val="centerGroup"/></m:mathPr>
               </w:settings>"#,
            w = W_NS,
            m = M_NS
        );
        let s = parse_document_settings(&xml).expect("settings present (defJc)");
        assert_eq!(s.math_def_jc.as_deref(), Some("centerGroup"));

        let empty = r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
        assert!(parse_document_settings(empty).is_none());
    }

    // ECMA-376 §22.1.2.88 + §17.3.1.13 `w:jc` — a display-math paragraph with no
    // explicit text alignment must keep its natural text alignment ("left" for
    // a Tabletext-styled cell). The math block's own justification is handled by
    // the renderer via `m:jc`/`m:defJc`, NOT by force-centering the paragraph.
    #[test]
    fn display_math_para_alignment_not_forced_center() {
        // No w:jc in pPr ⇒ natural default is "left"; oMathPara present.
        let p = parse_p(
            r#"<m:oMathPara><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr>
               <m:oMath><m:r><m:t>α</m:t></m:r></m:oMath></m:oMathPara>"#,
        );
        assert_eq!(p.alignment, "left");
    }
}

#[cfg(test)]
mod cs_toggle_tests {
    use super::*;

    fn run_of(body_inner: &str) -> TextRun {
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body_inner}</w:body></w:document>"#
        );
        let doc = XmlDoc::parse(&xml).unwrap();
        let body = doc
            .root_element()
            .descendants()
            .find(|n| n.tag_name().name() == "body")
            .unwrap();
        let style_map = StyleMap::parse("");
        let mut num_map = NumberingMap::default();
        let elems = parse_body_elements(
            body,
            &style_map,
            &mut num_map,
            &HashMap::new(),
            &HashMap::new(),
            &ThemeColors::default(),
        );
        for e in elems {
            if let BodyElement::Paragraph(p) = e {
                for r in p.runs {
                    if let DocRun::Text(t) = r {
                        return t;
                    }
                }
            }
        }
        panic!("no text run");
    }

    #[test]
    fn cs_element_sets_run_cs_toggle() {
        // §17.3.2.7 <w:cs/> — the complex-script run toggle.
        let run = run_of(r#"<w:p><w:r><w:rPr><w:cs/></w:rPr><w:t>x</w:t></w:r></w:p>"#);
        assert_eq!(run.cs, Some(true));
    }

    #[test]
    fn rfonts_cs_attribute_does_not_set_the_toggle() {
        // rFonts@cs is only a font SLOT (§17.3.2.26) — it must not force cs.
        let run =
            run_of(r#"<w:p><w:r><w:rPr><w:rFonts w:cs="Arial"/></w:rPr><w:t>x</w:t></w:r></w:p>"#);
        assert_eq!(run.cs, None);
        assert_eq!(run.font_family_cs.as_deref(), Some("Arial"));
    }
}

#[cfg(test)]
mod theme_cs_tests {
    use super::*;

    #[test]
    fn empty_theme_cs_falls_back_to_arial_for_arabic_bidi_lang() {
        // §17.15.1.88: theme <a:cs typeface=""/> + themeFontLang bidi="ar-SA"
        // → Word uses Arial for theme-referenced complex-script fonts
        // (sample-7.pdf embeds ArialMT for the no-rFonts table cells).
        let theme_xml = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:themeElements><a:fontScheme name="Office">
            <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
            <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
          </a:fontScheme></a:themeElements></a:theme>"#;
        let mut theme = ThemeColors::parse(theme_xml);
        assert_eq!(
            theme.resolve_font("minorBidi"),
            None,
            "empty cs typeface is unset"
        );
        theme.fill_default_cs_font("ar-SA");
        assert_eq!(theme.resolve_font("minorBidi").as_deref(), Some("Arial"));
        assert_eq!(theme.resolve_font("majorBidi").as_deref(), Some("Arial"));
        // Latin axes untouched
        assert_eq!(theme.resolve_font("minorHAnsi").as_deref(), Some("Calibri"));
    }

    #[test]
    fn explicit_theme_cs_font_is_not_overridden() {
        let theme_xml = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:themeElements><a:fontScheme name="X">
            <a:minorFont><a:latin typeface="Calibri"/><a:cs typeface="Sakkal Majalla"/></a:minorFont>
          </a:fontScheme></a:themeElements></a:theme>"#;
        let mut theme = ThemeColors::parse(theme_xml);
        theme.fill_default_cs_font("ar-SA");
        assert_eq!(
            theme.resolve_font("minorBidi").as_deref(),
            Some("Sakkal Majalla")
        );
    }

    #[test]
    fn non_rtl_bidi_lang_adds_no_default() {
        let mut theme = ThemeColors::parse(
            "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>",
        );
        theme.fill_default_cs_font("en-US");
        assert_eq!(theme.resolve_font("minorBidi"), None);
    }

    #[test]
    fn theme_font_lang_bidi_is_extracted_from_settings() {
        let settings = r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:themeFontLang w:val="en-AE" w:bidi="ar-SA"/></w:settings>"#;
        assert_eq!(
            parse_theme_font_bidi_lang(settings).as_deref(),
            Some("ar-SA")
        );
        let none = r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
        assert_eq!(parse_theme_font_bidi_lang(none), None);
    }

    #[test]
    fn document_settings_absent_when_no_kinsoku_elements() {
        // §17.15.1.58 default kinsoku=ON ⇒ no element ⇒ None (renderer defaults).
        let xml = r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:themeFontLang w:val="ja-JP"/></w:settings>"#;
        assert!(parse_document_settings(xml).is_none());
    }

    #[test]
    fn document_settings_kinsoku_off_is_surfaced() {
        // §17.15.1.58 explicit w:val="0" disables kinsoku.
        let xml = r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:kinsoku w:val="0"/></w:settings>"#;
        let s = parse_document_settings(xml).expect("settings present");
        assert_eq!(s.kinsoku, Some(false));
        assert_eq!(s.no_line_breaks_before, None);
        assert_eq!(s.no_line_breaks_after, None);
    }

    #[test]
    fn document_settings_kinsoku_on_explicit() {
        let xml = r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:kinsoku w:val="1"/></w:settings>"#;
        assert_eq!(parse_document_settings(xml).unwrap().kinsoku, Some(true));
    }

    #[test]
    fn document_settings_custom_no_line_breaks_replace_default() {
        // §17.15.1.59/.60 — custom sets surfaced verbatim (renderer replaces).
        let xml = r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:noLineBreaksBefore w:lang="ja-JP" w:val="、。"/>
          <w:noLineBreaksAfter w:lang="ja-JP" w:val="（「"/></w:settings>"#;
        let s = parse_document_settings(xml).unwrap();
        assert_eq!(s.no_line_breaks_before.as_deref(), Some("、。"));
        assert_eq!(s.no_line_breaks_after.as_deref(), Some("（「"));
    }

    #[test]
    fn document_settings_multiple_lang_no_line_breaks_concatenated() {
        let xml = r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:noLineBreaksBefore w:lang="ja-JP" w:val="、"/>
          <w:noLineBreaksBefore w:lang="zh-CN" w:val="。"/></w:settings>"#;
        let s = parse_document_settings(xml).unwrap();
        assert_eq!(s.no_line_breaks_before.as_deref(), Some("、。"));
    }
}

#[cfg(test)]
mod rtl_tests {
    use super::*;

    /// Parse a minimal `<w:body>` document through the real body-parse path
    /// (the layer below the wasm/zip-gated `parse` entry). Mirrors how the
    /// crate builds `BodyElement`s from `word/document.xml`.
    fn body_from(body_inner: &str) -> Vec<BodyElement> {
        let xml = format!(
            r#"<w:document xmlns:w="{ns}"><w:body>{inner}</w:body></w:document>"#,
            ns = W_NS,
            inner = body_inner,
        );
        let doc = XmlDoc::parse(&xml).unwrap();
        let body_node = doc
            .root_element()
            .descendants()
            .find(|n| n.tag_name().name() == "body")
            .unwrap();
        let style_map = StyleMap::parse("");
        let mut num_map = NumberingMap::default();
        let media_map: HashMap<String, String> = HashMap::new();
        let rel_map: HashMap<String, String> = HashMap::new();
        let theme = ThemeColors::default();
        parse_body_elements(
            body_node,
            &style_map,
            &mut num_map,
            &media_map,
            &rel_map,
            &theme,
        )
    }

    /// ECMA-376 §17.3.1.6 w:bidi, §17.3.2.30 w:rtl, §17.3.2.26 w:rFonts@cs,
    /// §17.3.2.39 w:szCs, §17.4.1 w:bidiVisual — all surfaced on the model.
    #[test]
    fn rtl_direction_attributes_are_extracted() {
        let body = body_from(
            r#"
            <w:p>
              <w:pPr><w:bidi/></w:pPr>
              <w:r>
                <w:rPr>
                  <w:rtl/>
                  <w:rFonts w:cs="Arial"/>
                  <w:szCs w:val="28"/>
                </w:rPr>
                <w:t>שלום</w:t>
              </w:r>
            </w:p>
            <w:tbl>
              <w:tblPr><w:bidiVisual/></w:tblPr>
              <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
              <w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr>
            </w:tbl>
            "#,
        );

        // Paragraph w:bidi
        let para = body
            .iter()
            .find_map(|e| match e {
                BodyElement::Paragraph(p) => Some(p),
                _ => None,
            })
            .expect("paragraph present");
        assert_eq!(para.bidi, Some(true), "w:bidi should set paragraph.bidi");

        // Run w:rtl + w:cs + w:szCs
        let run = para
            .runs
            .iter()
            .find_map(|r| match r {
                DocRun::Text(t) => Some(t),
                _ => None,
            })
            .expect("text run present");
        assert_eq!(run.rtl, Some(true), "w:rtl should set run.rtl");
        assert_eq!(
            run.font_family_cs.as_deref(),
            Some("Arial"),
            "w:rFonts@cs → run.fontFamilyCs"
        );
        assert_eq!(
            run.font_size_cs,
            Some(14.0),
            "w:szCs val=28 half-pts → 14pt run.fontSizeCs"
        );

        // Table w:bidiVisual
        let tbl = body
            .iter()
            .find_map(|e| match e {
                BodyElement::Table(t) => Some(t),
                _ => None,
            })
            .expect("table present");
        assert_eq!(
            tbl.bidi_visual,
            Some(true),
            "w:bidiVisual should set table.bidiVisual"
        );
    }

    /// On-off toggles honor an explicit `w:val="0"` (off), per §17.3.2.22, so
    /// the new flags can carry `Some(false)` distinctly from `None` (inherit).
    #[test]
    fn rtl_toggles_honor_explicit_off() {
        let body = body_from(
            r#"<w:p><w:pPr><w:bidi w:val="0"/></w:pPr><w:r><w:rPr><w:rtl w:val="false"/></w:rPr><w:t>a</w:t></w:r></w:p>"#,
        );
        let para = body
            .iter()
            .find_map(|e| match e {
                BodyElement::Paragraph(p) => Some(p),
                _ => None,
            })
            .unwrap();
        assert_eq!(para.bidi, Some(false));
        let run = para
            .runs
            .iter()
            .find_map(|r| match r {
                DocRun::Text(t) => Some(t),
                _ => None,
            })
            .unwrap();
        assert_eq!(run.rtl, Some(false));
    }

    /// ECMA-376 §17.3.2.18: a directly-applied `w:sz` with NO accompanying
    /// `w:szCs` mirrors into the complex-script size (Word draws Arabic in a
    /// run that only sets `w:sz` at that size — sample-7's underlined title sets
    /// sz=36 / no szCs and renders the Arabic at 18pt). An explicit `w:szCs`
    /// still wins independently.
    #[test]
    fn direct_sz_without_szcs_mirrors_into_cs_size() {
        // sz=36 (18pt), no szCs -> font_size_cs == font_size == 18.
        let body = body_from(
            r#"<w:p><w:r><w:rPr><w:b/><w:sz w:val="36"/></w:rPr><w:t>نبذة</w:t></w:r></w:p>"#,
        );
        let run = body
            .iter()
            .find_map(|e| match e {
                BodyElement::Paragraph(p) => p.runs.iter().find_map(|r| match r {
                    DocRun::Text(t) => Some(t),
                    _ => None,
                }),
                _ => None,
            })
            .expect("text run present");
        assert_eq!(run.font_size, 18.0, "sz=36 half-pts → 18pt");
        assert_eq!(
            run.font_size_cs,
            Some(18.0),
            "sz without szCs mirrors into the complex-script size (§17.3.2.18)"
        );

        // sz=36 AND szCs=24 -> cs size honors the explicit szCs (12pt).
        let body2 = body_from(
            r#"<w:p><w:r><w:rPr><w:sz w:val="36"/><w:szCs w:val="24"/></w:rPr><w:t>x</w:t></w:r></w:p>"#,
        );
        let run2 = body2
            .iter()
            .find_map(|e| match e {
                BodyElement::Paragraph(p) => p.runs.iter().find_map(|r| match r {
                    DocRun::Text(t) => Some(t),
                    _ => None,
                }),
                _ => None,
            })
            .unwrap();
        assert_eq!(run2.font_size, 18.0);
        assert_eq!(
            run2.font_size_cs,
            Some(12.0),
            "explicit szCs wins over sz mirroring"
        );
    }

    /// ECMA-376 §17.3.2.3 w:bCs, §17.3.2.17 w:iCs, §17.3.2.20 w:lang/@w:bidi —
    /// complex-script bold/italic toggles and the RTL language tag (lower-cased)
    /// are surfaced on the run model for the renderer's cs-formatting path.
    #[test]
    fn complex_script_bold_italic_and_lang_bidi_are_extracted() {
        let body = body_from(
            r#"<w:p><w:r><w:rPr><w:rtl/><w:bCs/><w:iCs/>
              <w:lang w:val="en-AE" w:bidi="ae-AR"/></w:rPr><w:t>28-02-2026</w:t></w:r></w:p>"#,
        );
        let run = body
            .iter()
            .find_map(|e| match e {
                BodyElement::Paragraph(p) => p.runs.iter().find_map(|r| match r {
                    DocRun::Text(t) => Some(t),
                    _ => None,
                }),
                _ => None,
            })
            .expect("text run present");
        assert_eq!(run.bold_cs, Some(true), "w:bCs → run.boldCs");
        assert_eq!(run.italic_cs, Some(true), "w:iCs → run.italicCs");
        assert_eq!(
            run.lang_bidi.as_deref(),
            Some("ae-ar"),
            "w:lang@w:bidi lower-cased → run.langBidi"
        );
    }

    /// Legacy VML text box (ECMA-376 Part 4 §14.1): `<w:pict>` with a
    /// `<v:shape type="#_x0000_t202">` surfaces as a ShapeRun carrying the
    /// fill/stroke from the VML attributes, the size from the CSS `style`, and
    /// the `<w:txbxContent>` body text — so the renderer draws the yellow box.
    #[test]
    fn vml_pict_textbox_becomes_a_shape_run() {
        let body = body_from(
            r##"<w:p><w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml">
              <v:shape id="tb1" type="#_x0000_t202"
                  style="position:relative;width:300pt;height:60pt"
                  fillcolor="#fdf2d0" strokecolor="#c0a000">
                <v:textbox><w:txbxContent>
                  <w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>
                    <w:r><w:rPr><w:rtl/></w:rPr><w:t>مربع نص 2025</w:t></w:r>
                  </w:p>
                </w:txbxContent></v:textbox>
              </v:shape>
            </w:pict></w:r></w:p>"##,
        );
        let para = body
            .iter()
            .find_map(|e| match e {
                BodyElement::Paragraph(p) => Some(p),
                _ => None,
            })
            .expect("paragraph present");
        let shape = para
            .runs
            .iter()
            .find_map(|r| match r {
                DocRun::Shape(s) => Some(s),
                _ => None,
            })
            .expect("VML pict should produce a ShapeRun");

        assert_eq!(shape.width_pt, 300.0, "width from CSS style");
        assert_eq!(shape.height_pt, 60.0, "height from CSS style");
        assert_eq!(shape.preset_geometry.as_deref(), Some("rect"));
        match &shape.fill {
            Some(ShapeFill::Solid { color }) => assert_eq!(color, "fdf2d0"),
            other => panic!("expected solid fdf2d0 fill, got {other:?}"),
        }
        assert_eq!(shape.stroke.as_deref(), Some("c0a000"));
        assert!(shape.stroke_width > 0.0, "stroke present ⇒ visible weight");
        assert_eq!(shape.text_blocks.len(), 1, "one body paragraph");
        assert_eq!(shape.text_blocks[0].text, "مربع نص 2025");
        assert_eq!(shape.text_blocks[0].alignment, "right");
    }

    /// ECMA-376 §17.4.* w:tcBorders with only the logical vertical edges
    /// (w:start/w:end, §17.4.66-67) must still yield left/right border specs.
    /// RTL/bidi-aware authoring tools emit start/end instead of left/right;
    /// the parser maps start→left, end→right (the renderer handles visual
    /// mirroring for bidiVisual tables).
    #[test]
    fn tc_borders_start_end_map_to_left_right() {
        let body = body_from(
            r#"
            <w:tbl>
              <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
              <w:tr><w:tc>
                <w:tcPr>
                  <w:tcBorders>
                    <w:start w:sz="1" w:val="single" w:color="D9D9D9"/>
                    <w:top w:sz="1" w:val="single" w:color="D9D9D9"/>
                    <w:end w:sz="1" w:val="single" w:color="D9D9D9"/>
                    <w:bottom w:sz="1" w:val="single" w:color="D9D9D9"/>
                  </w:tcBorders>
                </w:tcPr>
                <w:p><w:r><w:t>x</w:t></w:r></w:p>
              </w:tc></w:tr>
            </w:tbl>
            "#,
        );
        let tbl = body
            .iter()
            .find_map(|e| match e {
                BodyElement::Table(t) => Some(t),
                _ => None,
            })
            .expect("table present");
        let cell = &tbl.rows[0].cells[0];
        let left = cell
            .borders
            .left
            .as_ref()
            .expect("w:start should map to cell.left");
        let right = cell
            .borders
            .right
            .as_ref()
            .expect("w:end should map to cell.right");
        assert_eq!(left.color.as_deref(), Some("d9d9d9"), "start color → left");
        assert_eq!(left.style, "single");
        assert_eq!(right.color.as_deref(), Some("d9d9d9"), "end color → right");
        assert_eq!(right.style, "single");
        // top/bottom (literal physical names) still parsed.
        assert!(cell.borders.top.is_some());
        assert!(cell.borders.bottom.is_some());
    }

    /// ECMA-376 §17.4.* w:tblBorders: a table-level w:start/w:end likewise
    /// maps to the physical left/right edges.
    #[test]
    fn tbl_borders_start_end_map_to_left_right() {
        let body = body_from(
            r#"
            <w:tbl>
              <w:tblPr>
                <w:tblBorders>
                  <w:start w:sz="4" w:val="single" w:color="000000"/>
                  <w:end w:sz="4" w:val="single" w:color="000000"/>
                </w:tblBorders>
              </w:tblPr>
              <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
              <w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr>
            </w:tbl>
            "#,
        );
        let tbl = body
            .iter()
            .find_map(|e| match e {
                BodyElement::Table(t) => Some(t),
                _ => None,
            })
            .expect("table present");
        assert!(
            tbl.borders.left.is_some(),
            "w:start should map to tblBorders.left"
        );
        assert!(
            tbl.borders.right.is_some(),
            "w:end should map to tblBorders.right"
        );
    }

    /// ECMA-376 §17.3.2.4 w:bCs / §17.3.2.6 w:iCs are INDEPENDENT toggles from
    /// §17.3.2.3 w:b / §17.3.2.5 w:i AT THE PARSER LEVEL: a run that sets only
    /// `w:b` (no `w:bCs`) surfaces `bold_cs == None`. NOTE: the renderer's cs
    /// axis intentionally MIRRORS the non-cs value when bCs is absent
    /// (`csBold = boldCs ?? base.bold`) — PDF sample-7 page-1 headings carry
    /// `w:b` without `w:bCs` and render BOLD in Word. So `None` here means
    /// "not set on the cs axis at the parser level", not "renders non-bold".
    #[test]
    fn complex_script_bold_is_independent_of_non_cs_bold() {
        let body = body_from(
            r#"<w:p><w:r>
                <w:rPr><w:rtl/><w:cs/><w:b/><w:szCs w:val="24"/></w:rPr>
                <w:t>28-02-2026</w:t>
            </w:r></w:p>"#,
        );
        let para = body
            .iter()
            .find_map(|e| match e {
                BodyElement::Paragraph(p) => Some(p),
                _ => None,
            })
            .unwrap();
        let run = para
            .runs
            .iter()
            .find_map(|r| match r {
                DocRun::Text(t) => Some(t),
                _ => None,
            })
            .unwrap();
        // `w:b` sets the non-CS bold…
        assert!(run.bold, "w:b sets non-CS bold");
        // …and `w:bCs` is absent, so the parser leaves the CS bold axis None.
        // The renderer mirrors it from `bold` (boldCs ?? bold) per the PDF-
        // verified sample-7 page-1 behaviour; that fallback lives in renderer.ts.
        assert_eq!(
            run.bold_cs, None,
            "absent w:bCs stays None at the parser level"
        );
        assert_eq!(
            run.italic_cs, None,
            "absent w:iCs stays None at the parser level"
        );
    }

    /// An explicit `w:bCs` / `w:iCs` is surfaced on its own axis (and honors the
    /// `w:val="0"` off form, §17.3.2.22), independent of `w:b`/`w:i`.
    #[test]
    fn complex_script_bold_italic_surface_when_present() {
        let body = body_from(
            r#"<w:p><w:r>
                <w:rPr><w:rtl/><w:bCs/><w:iCs w:val="0"/></w:rPr>
                <w:t>عربي</w:t>
            </w:r></w:p>"#,
        );
        let para = body
            .iter()
            .find_map(|e| match e {
                BodyElement::Paragraph(p) => Some(p),
                _ => None,
            })
            .unwrap();
        let run = para
            .runs
            .iter()
            .find_map(|r| match r {
                DocRun::Text(t) => Some(t),
                _ => None,
            })
            .unwrap();
        assert_eq!(run.bold_cs, Some(true), "w:bCs sets the CS bold axis");
        assert_eq!(
            run.italic_cs,
            Some(false),
            "w:iCs val=0 turns CS italic off"
        );
    }
}

#[cfg(test)]
mod footnote_tests {
    use super::*;
    use crate::xml_util::W_NS;

    fn first_para(body_inner: &str) -> crate::types::DocParagraph {
        let xml = format!(
            r#"<w:document xmlns:w="{ns}"><w:body>{inner}</w:body></w:document>"#,
            ns = W_NS,
            inner = body_inner,
        );
        let doc = XmlDoc::parse(&xml).unwrap();
        let body_node = doc
            .root_element()
            .descendants()
            .find(|n| n.tag_name().name() == "body")
            .unwrap();
        let style_map = StyleMap::parse("");
        let mut num_map = NumberingMap::default();
        let elems = parse_body_elements(
            body_node,
            &style_map,
            &mut num_map,
            &HashMap::new(),
            &HashMap::new(),
            &ThemeColors::default(),
        );
        for e in elems {
            if let BodyElement::Paragraph(p) = e {
                return p;
            }
        }
        panic!("no paragraph parsed");
    }

    /// ECMA-376 §17.11.17 — a body `<w:footnoteReference>` becomes a superscript
    /// TextRun tagged with `note_ref { kind:"footnote", id }`.
    #[test]
    fn footnote_reference_is_tagged_as_note_ref() {
        let p = first_para(
            r#"<w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
                 <w:footnoteReference w:id="3"/></w:r></w:p>"#,
        );
        let run = p
            .runs
            .iter()
            .find_map(|r| match r {
                DocRun::Text(t) => Some(t),
                _ => None,
            })
            .expect("text run");
        let nr = run.note_ref.as_ref().expect("note_ref set");
        assert_eq!(nr.kind, "footnote");
        assert_eq!(nr.id, "3");
        // Model value is "super" (the styles.rs mapping of w:vertAlign
        // val="superscript"), NOT the raw XML token — the renderer raises
        // the baseline only for "super".
        assert_eq!(run.vert_align.as_deref(), Some("super"));
    }

    /// ECMA-376 §17.11.16 — the in-note `<w:footnoteRef>` placeholder is tagged
    /// with an empty id (the renderer substitutes the enclosing note's number).
    #[test]
    fn footnote_ref_placeholder_has_empty_id() {
        let p = first_para(r#"<w:p><w:r><w:footnoteRef/></w:r><w:r><w:t> body</w:t></w:r></w:p>"#);
        let nr = p
            .runs
            .iter()
            .find_map(|r| match r {
                DocRun::Text(t) => t.note_ref.clone(),
                _ => None,
            })
            .expect("note_ref set");
        assert_eq!(nr.kind, "footnote");
        assert_eq!(nr.id, "");
    }

    /// ECMA-376 §17.11.6 — endnote references carry kind "endnote".
    #[test]
    fn endnote_reference_kind_is_endnote() {
        let p = first_para(r#"<w:p><w:r><w:endnoteReference w:id="2"/></w:r></w:p>"#);
        let nr = p
            .runs
            .iter()
            .find_map(|r| match r {
                DocRun::Text(t) => t.note_ref.clone(),
                _ => None,
            })
            .expect("note_ref set");
        assert_eq!(nr.kind, "endnote");
        assert_eq!(nr.id, "2");
    }

    /// ECMA-376 §17.3.1.32 — w:snapToGrid val=0 surfaces as Some(false) so the
    /// renderer can opt the paragraph out of the docGrid.
    #[test]
    fn snap_to_grid_off_is_surfaced() {
        let p = first_para(
            r#"<w:p><w:pPr><w:snapToGrid w:val="0"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#,
        );
        assert_eq!(p.snap_to_grid, Some(false));
    }

    #[test]
    fn snap_to_grid_absent_is_none() {
        let p = first_para(r#"<w:p><w:r><w:t>x</w:t></w:r></w:p>"#);
        assert_eq!(p.snap_to_grid, None);
    }
}
