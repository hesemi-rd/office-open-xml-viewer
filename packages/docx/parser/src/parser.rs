use std::collections::HashMap;
use std::io::Read;
use zip::ZipArchive;
use roxmltree::Document as XmlDoc;
use base64::Engine;
use base64::engine::general_purpose::STANDARD as B64;

use crate::xml_util::*;
use crate::types::*;
use crate::styles::{StyleMap, parse_para_fmt, parse_run_fmt, ParaFmt, RunFmt};
use crate::numbering::NumberingMap;

const DEFAULT_FONT_SIZE: f64 = 10.0; // pt fallback

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

    let rels_xml = read_zip_entry(&mut zip, "word/_rels/document.xml.rels")
        .unwrap_or_default();
    let rel_map = parse_rels(&rels_xml);

    // Styles are referenced from the document relationships (Target may be
    // "styles.xml" or "styles2.xml"). Fall back to "word/styles.xml" for old files.
    let styles_path = find_rel_target(&rels_xml, "styles")
        .map(|t| if t.starts_with('/') { t.trim_start_matches('/').to_string() } else { format!("word/{}", t) })
        .unwrap_or_else(|| "word/styles.xml".to_string());
    let style_map = read_zip_entry(&mut zip, &styles_path)
        .map(|s| StyleMap::parse(&s))
        .unwrap_or_else(|_| StyleMap::parse(""));

    let numbering_path = find_rel_target(&rels_xml, "numbering")
        .map(|t| if t.starts_with('/') { t.trim_start_matches('/').to_string() } else { format!("word/{}", t) })
        .unwrap_or_else(|| "word/numbering.xml".to_string());
    let mut num_map = read_zip_entry(&mut zip, &numbering_path)
        .map(|s| NumberingMap::parse(&s))
        .unwrap_or_default();

    // Theme is referenced by a relationship with Type ending in "/theme" — resolve
    // to word/<target> and parse the clrScheme.
    let theme = find_rel_target(&rels_xml, "theme")
        .map(|t| {
            let p = if t.starts_with('/') { t.trim_start_matches('/').to_string() } else { format!("word/{}", t) };
            read_zip_entry(&mut zip, &p).map(|s| ThemeColors::parse(&s)).unwrap_or_default()
        })
        .unwrap_or_default();

    let media_map = load_media_map(&mut zip, &rel_map, "word/");

    let doc_xml = read_zip_entry(&mut zip, "word/document.xml")?;
    let xml_doc = XmlDoc::parse(&doc_xml).map_err(|e| e.to_string())?;

    let body_node = xml_doc.root_element()
        .descendants()
        .find(|n| n.tag_name().name() == "body")
        .ok_or("No body element")?;

    let sect_pr = body_node.children()
        .filter(|n| n.is_element())
        .last()
        .filter(|n| n.tag_name().name() == "sectPr");

    let (section, refs) = parse_section(sect_pr, &rel_map);

    let body = parse_body_elements(body_node, &style_map, &mut num_map, &media_map, &rel_map, &theme);

    let headers = load_header_footer_set(&mut zip, &refs.headers, "hdr", &style_map, &mut num_map, &theme);
    let footers = load_header_footer_set(&mut zip, &refs.footers, "ftr", &style_map, &mut num_map, &theme);

    let major_font = theme.theme_font("major", "latin");
    let minor_font = theme.theme_font("minor", "latin");

    // ECMA-376 §17.8.3.10: font family classification from fontTable.xml.
    // Resolve via relationship (Type ending in "/fontTable"); fall back to
    // "word/fontTable.xml" for documents that omit the relationship.
    let font_table_path = find_rel_target(&rels_xml, "fontTable")
        .map(|t| if t.starts_with('/') { t.trim_start_matches('/').to_string() } else { format!("word/{}", t) })
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
        .map(|t| if t.starts_with('/') { t.trim_start_matches('/').to_string() } else { format!("word/{}", t) })
        .and_then(|p| read_zip_entry(&mut zip, &p).ok())
        .map(|xml| parse_comments(&xml))
        .unwrap_or_default();
    let footnotes = find_rel_target(&rels_xml, "footnotes")
        .map(|t| if t.starts_with('/') { t.trim_start_matches('/').to_string() } else { format!("word/{}", t) })
        .and_then(|p| read_zip_entry(&mut zip, &p).ok())
        .map(|xml| parse_notes(&xml, "footnote"))
        .unwrap_or_default();
    let endnotes = find_rel_target(&rels_xml, "endnotes")
        .map(|t| if t.starts_with('/') { t.trim_start_matches('/').to_string() } else { format!("word/{}", t) })
        .and_then(|p| read_zip_entry(&mut zip, &p).ok())
        .map(|xml| parse_notes(&xml, "endnote"))
        .unwrap_or_default();

    Ok(Document {
        section, body, headers, footers, major_font, minor_font, font_family_classes,
        revisions, comments, footnotes, endnotes,
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
    let Ok(doc) = roxmltree::Document::parse(xml) else { return Vec::new(); };
    let mut out = Vec::new();
    for c in doc.descendants().filter(|n| n.is_element() && n.tag_name().name() == "comment") {
        let id = c.attributes().find(|a| a.name() == "id").map(|a| a.value().to_string()).unwrap_or_default();
        if id.is_empty() { continue }
        let author = c.attributes().find(|a| a.name() == "author").map(|a| a.value().to_string()).filter(|s| !s.is_empty());
        let initials = c.attributes().find(|a| a.name() == "initials").map(|a| a.value().to_string()).filter(|s| !s.is_empty());
        let date = c.attributes().find(|a| a.name() == "date").map(|a| a.value().to_string()).filter(|s| !s.is_empty());
        let mut text = String::new();
        for t in c.descendants().filter(|n| n.is_element() && n.tag_name().name() == "t") {
            if let Some(s) = t.text() { text.push_str(s); }
        }
        out.push(crate::types::DocxComment { id, author, initials, date, text });
    }
    out
}

/// Parse word/footnotes.xml or word/endnotes.xml. Excludes the two
/// reserved entries (id="-1" separator, id="0" continuation separator)
/// per ECMA-376 §17.11.10.
fn parse_notes(xml: &str, element_name: &str) -> Vec<crate::types::DocxNote> {
    let Ok(doc) = roxmltree::Document::parse(xml) else { return Vec::new(); };
    let mut out = Vec::new();
    for n in doc.descendants().filter(|n| n.is_element() && n.tag_name().name() == element_name) {
        let id = n.attributes().find(|a| a.name() == "id").map(|a| a.value().to_string()).unwrap_or_default();
        if id.is_empty() || id == "-1" || id == "0" { continue }
        let mut text = String::new();
        for t in n.descendants().filter(|t| t.is_element() && t.tag_name().name() == "t") {
            if let Some(s) = t.text() { text.push_str(s); }
        }
        out.push(crate::types::DocxNote { id, text });
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
        let doc = match XmlDoc::parse(xml) { Ok(d) => d, Err(_) => return Self { map, fonts, theme_xml } };
        let root = doc.root_element();
        if let Some(scheme) = root.descendants().find(|n| n.is_element() && n.tag_name().name() == "clrScheme") {
            for child in scheme.children().filter(|n| n.is_element()) {
                let name = child.tag_name().name().to_string();
                let hex = child.children().filter(|n| n.is_element()).find_map(|n| {
                    match n.tag_name().name() {
                        "srgbClr" => n.attribute("val").map(|v| v.to_uppercase()),
                        "sysClr" => n.attribute("lastClr").map(|v| v.to_uppercase()),
                        _ => None,
                    }
                });
                if let Some(h) = hex { map.insert(name, h); }
            }
        }
        if let Some(font_scheme) = root.descendants().find(|n| n.is_element() && n.tag_name().name() == "fontScheme") {
            for group_name in &["majorFont", "minorFont"] {
                let prefix = if *group_name == "majorFont" { "major" } else { "minor" };
                if let Some(group) = font_scheme.children().find(|n| n.is_element() && n.tag_name().name() == *group_name) {
                    for child in group.children().filter(|n| n.is_element()) {
                        let typ = match child.tag_name().name() {
                            "latin" => Some("latin"),
                            "ea"    => Some("ea"),
                            "cs"    => Some("cs"),
                            _       => None,
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
        Self { map, fonts, theme_xml }
    }

    fn resolve(&self, scheme_name: &str) -> Option<String> {
        // "bg1"/"bg2"/"tx1"/"tx2" map onto lt1/lt2/dk1/dk2 per spec
        let key = match scheme_name {
            "bg1" => "lt1",
            "bg2" => "lt2",
            "tx1" => "dk1",
            "tx2" => "dk2",
            other => other,
        };
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

    /// Resolve an rFonts theme reference (e.g. "minorHAnsi" → minor.latin,
    /// "minorEastAsia" → minor.ea, "majorHAnsi" → major.latin). Returns None
    /// when the reference is unknown or the theme has no matching typeface.
    pub fn resolve_font(&self, theme_ref: &str) -> Option<String> {
        let (group, axis) = match theme_ref {
            "minorHAnsi" | "minorAscii" => ("minor", "latin"),
            "minorBidi"                 => ("minor", "cs"),
            "minorEastAsia"             => ("minor", "ea"),
            "majorHAnsi" | "majorAscii" => ("major", "latin"),
            "majorBidi"                 => ("major", "cs"),
            "majorEastAsia"             => ("major", "ea"),
            _ => return None,
        };
        self.fonts.get(&format!("{group}/{axis}")).cloned()
    }
}

fn find_rel_target(rels_xml: &str, type_suffix: &str) -> Option<String> {
    if rels_xml.is_empty() { return None; }
    let doc = XmlDoc::parse(rels_xml).ok()?;
    for rel in doc.root_element().children().filter(|n| n.tag_name().name() == "Relationship") {
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
    let Ok(doc) = XmlDoc::parse(xml) else { return map };
    let w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    for font in doc.root_element()
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "font" && n.tag_name().namespace() == Some(w_ns))
    {
        let Some(name) = font.attribute((w_ns, "name")) else { continue };
        let family = font.children()
            .find(|n| n.is_element() && n.tag_name().name() == "family" && n.tag_name().namespace() == Some(w_ns))
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
                let result = parse_paragraph(child, style_map, num_map, media_map, rel_map, theme, None);
                let is_page_break_only = result.runs.len() == 1 && matches!(
                    result.runs[0],
                    DocRun::Break { break_type: BreakType::Page }
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
                if let Some(sect_pr) = child_w(child, "pPr").and_then(|ppr| child_w(ppr, "sectPr")) {
                    match read_section_break_type(sect_pr).as_deref() {
                        Some("continuous") => { /* no page break */ }
                        Some("oddPage") => body.push(BodyElement::PageBreak { parity: Some("odd".to_string()) }),
                        Some("evenPage") => body.push(BodyElement::PageBreak { parity: Some("even".to_string()) }),
                        _ => body.push(BodyElement::PageBreak { parity: None }),
                    }
                }
            }
            "tbl" => {
                let tbl = parse_table(child, style_map, num_map, media_map, rel_map, theme);
                body.push(BodyElement::Table(tbl));
            }
            "sectPr" => {
                // Mid-body loose sectPr (rare) would behave like a page break.
                // The final body-level sectPr only defines section settings — skip it.
                if Some(child.id()) != body_level_sect_id {
                    match read_section_break_type(child).as_deref() {
                        Some("continuous") => {}
                        Some("oddPage") => body.push(BodyElement::PageBreak { parity: Some("odd".to_string()) }),
                        Some("evenPage") => body.push(BodyElement::PageBreak { parity: Some("even".to_string()) }),
                        _ => body.push(BodyElement::PageBreak { parity: None }),
                    }
                }
            }
            _ => {}
        }
    }
    body
}

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
    let is_break_run = |r: &DocRun| matches!(r, DocRun::Break { break_type: BreakType::Page });
    let has_break = para.runs.iter().any(is_break_run);
    if !has_break {
        // Strip the (ignored) RenderedPage runs so they don't pollute layout.
        let mut p = para;
        p.runs.retain(|r| !matches!(r, DocRun::Break { break_type: BreakType::RenderedPage }));
        return vec![ParaPiece::Para(p)];
    }

    // Split run chunks on hard page breaks; RenderedPage runs are dropped.
    let mut chunks: Vec<Vec<DocRun>> = vec![Vec::new()];
    for run in para.runs.iter().cloned() {
        match &run {
            DocRun::Break { break_type: BreakType::Page } => chunks.push(Vec::new()),
            DocRun::Break { break_type: BreakType::RenderedPage } => { /* ignored hint */ }
            _ => chunks.last_mut().unwrap().push(run),
        }
    }

    let has_visible = |runs: &Vec<DocRun>| {
        runs.iter().any(|r| matches!(r,
            DocRun::Text(t) if !t.text.trim().is_empty())
            || matches!(r, DocRun::Field(_) | DocRun::Image(_) | DocRun::Shape(_)))
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
                let mime = if path.ends_with(".png") { "image/png" }
                    else if path.ends_with(".jpg") || path.ends_with(".jpeg") { "image/jpeg" }
                    else if path.ends_with(".gif") { "image/gif" }
                    else { "image/png" };
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
        let Some(root) = xml_doc.root_element().descendants().find(|n| n.tag_name().name() == root_tag) else {
            continue;
        };

        let body = parse_body_elements(root, style_map, num_map, &local_media_map, &local_rel_map, theme);
        let hf = HeaderFooter { body };
        match kind.as_str() {
            "first" => out.first = Some(hf),
            "even" => out.even = Some(hf),
            _ => out.default = Some(hf),
        }
    }
    out
}

fn parse_section(sect_pr: Option<roxmltree::Node>, rel_map: &HashMap<String, String>) -> (SectionProps, SectionRefs) {
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

    let Some(sp) = sect_pr else { return (default, SectionRefs::default()) };

    let mut props = default;
    if let Some(pg_sz) = child_w(sp, "pgSz") {
        if let Some(w) = attr_w(pg_sz, "w") { props.page_width = twips_to_pt(&w); }
        if let Some(h) = attr_w(pg_sz, "h") { props.page_height = twips_to_pt(&h); }
    }
    if let Some(pg_mar) = child_w(sp, "pgMar") {
        if let Some(v) = attr_w(pg_mar, "top") { props.margin_top = twips_to_pt(&v); }
        if let Some(v) = attr_w(pg_mar, "right") { props.margin_right = twips_to_pt(&v); }
        if let Some(v) = attr_w(pg_mar, "bottom") { props.margin_bottom = twips_to_pt(&v); }
        if let Some(v) = attr_w(pg_mar, "left") { props.margin_left = twips_to_pt(&v); }
        if let Some(v) = attr_w(pg_mar, "header") { props.header_distance = twips_to_pt(&v); }
        if let Some(v) = attr_w(pg_mar, "footer") { props.footer_distance = twips_to_pt(&v); }
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
        if local != "headerReference" && local != "footerReference" { continue; }
        let kind = attr_w(child, "type").unwrap_or_else(|| "default".to_string());
        let rid = child.attribute((R_NS, "id"))
            .or_else(|| child.attribute("id"))
            .map(|s| s.to_string());
        let Some(rid) = rid else { continue };
        let Some(target) = rel_map.get(&rid) else { continue };
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
    let (mut base_para, mut base_run) = style_map.resolve_para(explicit_style_id.as_deref(), table_style_id);

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

    let alignment = base_para.alignment.as_deref().map(normalize_align).unwrap_or("left").to_string();
    let indent_right = base_para.indent_right.unwrap_or(0.0);
    let space_before = base_para.space_before.unwrap_or(0.0);
    let space_after = base_para.space_after.unwrap_or(0.0);
    let line_spacing = base_para.line_spacing_val.map(|v| LineSpacing {
        value: v,
        rule: base_para.line_spacing_rule.clone().unwrap_or_else(|| "auto".to_string()),
        explicit: base_para.line_spacing_explicit.unwrap_or(false),
    });

    // Numbering — extract level data before advancing counter (avoids borrow conflict)
    let numbering = if let (Some(num_id), Some(num_level)) = (base_para.num_id, base_para.num_level) {
        if num_id != 0 {
            let (format, ind_left, tab) = num_map.get_level(num_id, num_level)
                .map(|l| (l.format.clone(), l.indent_left, l.tab))
                .unwrap_or_else(|| ("decimal".to_string(), 36.0, 18.0));
            let counter = num_map.advance(num_id, num_level);
            let text = num_map.resolve_text(num_id, num_level, counter);
            Some(NumberingInfo { num_id, level: num_level, format, text, indent_left: ind_left, tab })
        } else { None }
    } else { None };

    // Numbering level's pPr/ind overrides the paragraph style's indent
    let (indent_left, indent_first) = if let Some(ref num) = numbering {
        num_map.get_level(num.num_id, num.level)
            .map(|l| (l.indent_left, -l.tab))
            .unwrap_or((base_para.indent_left.unwrap_or(0.0), base_para.indent_first.unwrap_or(0.0)))
    } else {
        (base_para.indent_left.unwrap_or(0.0), base_para.indent_first.unwrap_or(0.0))
    };

    // Parse runs
    let mut runs = vec![];
    parse_para_content(node, &base_run, style_map, media_map, rel_map, theme, &mut runs, None);

    let tab_stops = base_para.tab_stops.clone().unwrap_or_default().into_iter()
        .map(|(pos, alignment, leader)| TabStop { pos, alignment, leader })
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
        keep_next: base_para.keep_next.unwrap_or_else(|| base_para.outline_level.is_some()),
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
        default_font_family: theme.resolve_font_ref(base_run.font_family_ascii.clone())
            .or_else(|| theme.resolve_font_ref(base_run.font_family_east_asia.clone())),
        outline_level: base_para.outline_level,
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
}

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
                handle_run_in_para(child, base_run, style_map, media_map, theme, runs, &mut field, None, revision);
            }
            "hyperlink" => {
                // Resolve URL from r:id via relationships
                let href = child.attribute((R_NS, "id"))
                    .or_else(|| child.attribute("id"))
                    .and_then(|rid| rel_map.get(rid).cloned());
                for r in child.children().filter(|n| n.is_element() && n.tag_name().name() == "r") {
                    handle_run_in_para(r, base_run, style_map, media_map, theme, runs, &mut field, Some(href.clone()), revision);
                }
            }
            "ins" | "del" => {
                // ECMA-376 §17.13.5 — build a RunRevision context covering
                // every descendant run so the renderer can paint tracked
                // changes inline. Nested ins/del isn't legal per spec; the
                // inner block wins if it occurs anyway.
                let kind = if child.tag_name().name() == "ins" { "insertion" } else { "deletion" };
                let inner = RunRevision {
                    kind: kind.to_string(),
                    author: attr_w(child, "author"),
                    date: attr_w(child, "date"),
                };
                parse_para_content(child, base_run, style_map, media_map, rel_map, theme, runs, Some(&inner));
            }
            "smartTag" => {
                parse_para_content(child, base_run, style_map, media_map, rel_map, theme, runs, revision);
            }
            "fldSimple" => {
                let instr = attr_w(child, "instr").unwrap_or_default();
                // Collect formatting from the first contained run (if any)
                let mut fmt = base_run.clone();
                if let Some(r) = child.children().find(|n| n.is_element() && n.tag_name().name() == "r") {
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
                    });
                }
            }
            "oMathPara" => {
                // A block math paragraph wraps one or more m:oMath children.
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
                        });
                    }
                }
            }
            _ => {}
        }
    }
}

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
            }
            "separate" => {
                field.past_separate = true;
            }
            "end" => {
                if field.active {
                    let fmt = field.fmt.clone().unwrap_or_else(|| base_run.clone());
                    runs.push(make_field_run(&field.instruction, &fmt, &field.fallback, theme));
                }
                *field = FieldState::default();
            }
            _ => {}
        }
        return;
    }

    if field.active {
        if !field.past_separate {
            // Capture instruction text and remember the formatting of the first instruction run
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
        } else {
            // Fallback/result text between separate and end — accumulate for "other" fields
            for c in r_node.children().filter(|n| n.is_element() && n.tag_name().name() == "t") {
                if let Some(t) = c.text() {
                    field.fallback.push_str(t);
                }
            }
        }
        return;
    }

    // Normal run
    parse_run_inner(r_node, base_run, style_map, media_map, theme, runs, link_href, revision);
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
    let token = instr.trim().split_whitespace().next().unwrap_or("").to_ascii_uppercase();
    match token.as_str() {
        "PAGE" => "page".to_string(),
        "NUMPAGES" => "numPages".to_string(),
        _ => "other".to_string(),
    }
}

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
    if fmt.vanish.unwrap_or(false) { return; }

    let is_link = link_href.is_some();
    let hyperlink = link_href.clone().flatten();

    let bold = fmt.bold.unwrap_or(false);
    let italic = fmt.italic.unwrap_or(false);
    let underline = fmt.underline.unwrap_or(false) || is_link;
    let strikethrough = fmt.strikethrough.unwrap_or(false);
    let font_size = fmt.font_size.unwrap_or(DEFAULT_FONT_SIZE);
    let color = if is_link && fmt.color.is_none() {
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
                }));
            }
            "br" => {
                let break_type = attr_w(child, "type").as_deref().map(|v| match v {
                    "page" => BreakType::Page,
                    "column" => BreakType::Column,
                    _ => BreakType::Line,
                }).unwrap_or(BreakType::Line);
                runs.push(DocRun::Break { break_type });
            }
            "lastRenderedPageBreak" => {
                // ECMA-376 §17.3.1.20: Word stores a hint at the location
                // where the previous render placed a page break. We mark
                // it as a separate `RenderedPage` break type so the
                // paragraph splitter can decide whether to honor it
                // (currently: only inside ruby-bearing paragraphs).
                runs.push(DocRun::Break { break_type: BreakType::RenderedPage });
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
                    Some(RubyAnnotation { text: rt_text, font_size_pt: rt_size_pt })
                } else {
                    None
                };
                if let Some(rb) = child_w(child, "rubyBase") {
                    let before = runs.len();
                    for inner in rb.children().filter(|n| n.is_element() && n.tag_name().name() == "r") {
                        parse_run_inner(inner, &fmt, style_map, media_map, theme, runs, link_href.clone(), revision);
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
            "footnoteReference" | "endnoteReference" => {
                // ECMA-376 §17.11.16: render the footnote number as superscript
                // at the reference point. Full bottom-of-page footnote
                // rendering isn't implemented; we at least place the marker.
                let id_str = attr_w(child, "id").unwrap_or_else(|| "?".to_string());
                runs.push(DocRun::Text(TextRun {
                    text: id_str,
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
                    vert_align: Some("superscript".to_string()),
                    hyperlink: hyperlink.clone(),
                    all_caps,
                    small_caps,
                    double_strikethrough,
                    highlight: highlight.clone(),
                    ruby: None,
                    revision: revision.cloned(),
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
        let extent = match container.children().find(|n| n.tag_name().name() == "extent") {
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
        let r_id = match blip.attribute((R_NS, "embed")).or_else(|| blip.attribute("r:embed")) {
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
        })];
    }

    // ── Anchor image/shape ─────────────────────────────────
    let container = match node.descendants().find(|n| n.tag_name().name() == "anchor") {
        Some(c) => c,
        None => return vec![],
    };

    // Parse positionH / positionV with relativeFrom
    let (pos_x, x_from_margin, x_align) = parse_anchor_pos_h(&container);
    let (pos_y, y_from_para,   y_align) = parse_anchor_pos_v(&container);
    let (pct_h, pct_v, rel_h, rel_v) = parse_anchor_pct_pos(&container);
    let (size_w_pct, size_h_pct, size_w_rel, size_h_rel) = parse_anchor_size_rel(&container);
    let anchor_meta = parse_anchor_wrap(&container);

    // behindDoc="1" flag — renderer uses this to draw shapes before text
    let behind_doc = container.attribute("behindDoc").map(|v| v == "1" || v == "true").unwrap_or(false);

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
    if let Some(wgp) = container.descendants().find(|n| n.tag_name().name() == "wgp") {
        let mut out: Vec<DocRun> = Vec::new();
        for img in parse_wgp_images(wgp, media_map, pos_x, x_from_margin, pos_y, y_from_para, &anchor_meta) {
            out.push(DocRun::Image(img));
        }
        for mut shp in parse_wgp_shapes(wgp, theme, pos_x, x_from_margin, pos_y, y_from_para, &anchor_meta) {
            shp.behind_doc = behind_doc;
            apply_pos_meta(&mut shp);
            out.push(DocRun::Shape(shp));
        }
        return out;
    }

    // wps:wsp directly under the anchor (no wgp wrapper)
    if let Some(wsp) = container.descendants().find(|n| n.tag_name().name() == "wsp") {
        if let Some(mut shp) = parse_wsp_shape(wsp, theme, pos_x, x_from_margin, pos_y, y_from_para, &anchor_meta, 1.0, 1.0, 0.0, 0.0, 0) {
            shp.behind_doc = behind_doc;
            apply_pos_meta(&mut shp);
            return vec![DocRun::Shape(shp)];
        }
    }

    // Regular single-blip anchor
    let extent = match container.children().find(|n| n.tag_name().name() == "extent") {
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
    let r_id = match blip.attribute((R_NS, "embed")).or_else(|| blip.attribute("r:embed")) {
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
    })]
}

#[derive(Default, Clone)]
struct AnchorMeta {
    wrap_mode: Option<String>,
    wrap_side: Option<String>,
    dist_top: f64,
    dist_bottom: f64,
    dist_left: f64,
    dist_right: f64,
}

/// Parse wrap element and dist* padding from a wp:anchor container.
fn parse_anchor_wrap(container: &roxmltree::Node) -> AnchorMeta {
    let to_pt = |s: &str| s.parse::<f64>().ok().map(|v| v / 12700.0).unwrap_or(0.0);
    let dist_top = container.attribute("distT").map(to_pt).unwrap_or(0.0);
    let dist_bottom = container.attribute("distB").map(to_pt).unwrap_or(0.0);
    let dist_left = container.attribute("distL").map(to_pt).unwrap_or(0.0);
    let dist_right = container.attribute("distR").map(to_pt).unwrap_or(0.0);

    let mut wrap_mode: Option<String> = None;
    let mut wrap_side: Option<String> = None;

    for child in container.children().filter(|n| n.is_element()) {
        let name = child.tag_name().name();
        match name {
            "wrapSquare"       => { wrap_mode = Some("square".into());       wrap_side = child.attribute("wrapText").map(|s| s.to_string()); break; }
            "wrapTopAndBottom" => { wrap_mode = Some("topAndBottom".into()); break; }
            "wrapNone"         => { wrap_mode = Some("none".into());         break; }
            "wrapTight"        => { wrap_mode = Some("tight".into());        wrap_side = child.attribute("wrapText").map(|s| s.to_string()); break; }
            "wrapThrough"      => { wrap_mode = Some("through".into());      wrap_side = child.attribute("wrapText").map(|s| s.to_string()); break; }
            _ => {}
        }
    }

    AnchorMeta { wrap_mode, wrap_side, dist_top, dist_bottom, dist_left, dist_right }
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
    for ac in container.children().filter(|n| n.tag_name().name() == "AlternateContent") {
        if let Some(choice) = ac.children().find(|n| n.tag_name().name() == "Choice") {
            if let Some(n) = choice.descendants().find(|n| n.is_element() && n.tag_name().name() == name) {
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
    let offset = pos.children()
        .find(|n| n.tag_name().name() == "posOffset")
        .and_then(|n| n.text())
        .and_then(|t| t.parse::<f64>().ok())
        .map(|emu| emu / 12700.0)
        .unwrap_or(0.0);
    // ECMA-376 §20.4.3.1: <wp:align>left|center|right</wp:align> takes
    // precedence over posOffset when both are present.
    let align = pos.children()
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
    let offset = pos.children()
        .find(|n| n.tag_name().name() == "posOffset")
        .and_then(|n| n.text())
        .and_then(|t| t.parse::<f64>().ok())
        .map(|emu| emu / 12700.0)
        .unwrap_or(0.0);
    let align = pos.children()
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
        let node = container.children().find(|n| n.tag_name().name() == outer_tag);
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
    let mut results = Vec::new();
    // Iterate all pic descendants in the wgp (covers both direct children and nested grpSp)
    for pic in wgp.descendants().filter(|n| n.tag_name().name() == "pic") {
        // Position and size come from the pic's spPr > a:xfrm
        let sp_pr = match pic.children().find(|n| n.tag_name().name() == "spPr") {
            Some(s) => s,
            None => continue,
        };
        let xfrm = match sp_pr.children().find(|n| n.tag_name().name() == "xfrm") {
            Some(x) => x,
            None => continue,
        };
        let off = match xfrm.children().find(|n| n.tag_name().name() == "off") {
            Some(o) => o,
            None => continue,
        };
        let ext = match xfrm.children().find(|n| n.tag_name().name() == "ext") {
            Some(e) => e,
            None => continue,
        };
        let ox = off.attribute("x").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 12700.0;
        let oy = off.attribute("y").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 12700.0;
        let cx = ext.attribute("cx").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 12700.0;
        let cy = ext.attribute("cy").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 12700.0;

        if cx <= 0.0 || cy <= 0.0 { continue; }

        // Find the blip inside this pic
        let blip = match pic.descendants().find(|n| n.tag_name().name() == "blip") {
            Some(b) => b,
            None => continue,
        };
        let r_id = match blip.attribute((R_NS, "embed")).or_else(|| blip.attribute("r:embed")) {
            Some(r) => r,
            None => continue,
        };
        let data_url = match media_map.get(r_id) {
            Some(u) => u.clone(),
            None => continue,
        };

        // Parse a:clrChange if present — used to make a specific color transparent.
        // clrFrom specifies the source color; clrTo with alpha=0 means replace with transparent.
        let color_replace_from = blip.children()
            .find(|n| n.tag_name().name() == "clrChange")
            .and_then(|cc| cc.children().find(|n| n.tag_name().name() == "clrFrom"))
            .and_then(|cf| cf.children().find(|n| n.tag_name().name() == "srgbClr"))
            .and_then(|clr| clr.attribute("val").map(|v| v.to_uppercase()));

        results.push(ImageRun {
            data_url,
            width_pt: cx,
            height_pt: cy,
            anchor: true,
            // Combine the group's anchor offset with this pic's offset within the group
            anchor_x_pt: anchor_pos_x + ox,
            anchor_y_pt: anchor_pos_y + oy,
            anchor_x_from_margin: x_from_margin,
            anchor_y_from_para: y_from_para,
            color_replace_from,
            wrap_mode: anchor_meta.wrap_mode.clone(),
            dist_top: anchor_meta.dist_top,
            dist_bottom: anchor_meta.dist_bottom,
            dist_left: anchor_meta.dist_left,
            dist_right: anchor_meta.dist_right,
            wrap_side: anchor_meta.wrap_side.clone(),
        });
    }
    results
}

/// Expand wps:wsp descendants of a wgp into ShapeRun entries. Applies
/// wgp grpSpPr transform (chOff/chExt → off/ext scale) to each child shape.
fn parse_wgp_shapes(
    wgp: roxmltree::Node,
    theme: &ThemeColors,
    anchor_pos_x: f64,
    x_from_margin: bool,
    anchor_pos_y: f64,
    y_from_para: bool,
    anchor_meta: &AnchorMeta,
) -> Vec<ShapeRun> {
    // Read group transform: off/ext (page-relative) vs chOff/chExt (child coord space).
    let grp_xfrm = wgp.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "grpSpPr")
        .and_then(|gsp| gsp.children().find(|n| n.is_element() && n.tag_name().name() == "xfrm"));
    let (off_x, off_y, ext_cx, ext_cy, ch_off_x, ch_off_y, ch_ext_cx, ch_ext_cy) = grp_xfrm
        .map(|x| {
            let off = x.children().find(|n| n.is_element() && n.tag_name().name() == "off");
            let ext = x.children().find(|n| n.is_element() && n.tag_name().name() == "ext");
            let ch_off = x.children().find(|n| n.is_element() && n.tag_name().name() == "chOff");
            let ch_ext = x.children().find(|n| n.is_element() && n.tag_name().name() == "chExt");
            (
                off.and_then(|o| o.attribute("x").and_then(|v| v.parse::<f64>().ok())).unwrap_or(0.0),
                off.and_then(|o| o.attribute("y").and_then(|v| v.parse::<f64>().ok())).unwrap_or(0.0),
                ext.and_then(|e| e.attribute("cx").and_then(|v| v.parse::<f64>().ok())).unwrap_or(0.0),
                ext.and_then(|e| e.attribute("cy").and_then(|v| v.parse::<f64>().ok())).unwrap_or(0.0),
                ch_off.and_then(|o| o.attribute("x").and_then(|v| v.parse::<f64>().ok())).unwrap_or(0.0),
                ch_off.and_then(|o| o.attribute("y").and_then(|v| v.parse::<f64>().ok())).unwrap_or(0.0),
                ch_ext.and_then(|e| e.attribute("cx").and_then(|v| v.parse::<f64>().ok())).unwrap_or(0.0),
                ch_ext.and_then(|e| e.attribute("cy").and_then(|v| v.parse::<f64>().ok())).unwrap_or(0.0),
            )
        })
        .unwrap_or((0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0));

    let sx = if ch_ext_cx > 0.0 && ext_cx > 0.0 { ext_cx / ch_ext_cx } else { 1.0 };
    let sy = if ch_ext_cy > 0.0 && ext_cy > 0.0 { ext_cy / ch_ext_cy } else { 1.0 };
    // Page-relative offset of the group origin, in EMU. ch_off is subtracted
    // because child coordinates are measured relative to chOff.
    let group_page_off_x_emu = off_x - ch_off_x * sx;
    let group_page_off_y_emu = off_y - ch_off_y * sy;

    // Outer group dimensions in pt — passed to each child so the renderer can
    // resolve align/pctPos against the GROUP's bounding box, then offset each
    // child within it. Falls back to the child-coord-space ext when the group
    // omits an outer ext (rare).
    let group_w_pt = (if ext_cx > 0.0 { ext_cx } else { ch_ext_cx }) / 12700.0;
    let group_h_pt = (if ext_cy > 0.0 { ext_cy } else { ch_ext_cy }) / 12700.0;

    let mut results = Vec::new();
    for (idx, wsp) in wgp.descendants().filter(|n| n.is_element() && n.tag_name().name() == "wsp").enumerate() {
        if let Some(mut shape) = parse_wsp_shape(
            wsp, theme,
            anchor_pos_x, x_from_margin,
            anchor_pos_y, y_from_para,
            anchor_meta,
            sx, sy,
            group_page_off_x_emu / 12700.0, group_page_off_y_emu / 12700.0,
            idx as u32,
        ) {
            shape.group_width_pt = Some(group_w_pt);
            shape.group_height_pt = Some(group_h_pt);
            results.push(shape);
        }
    }
    results
}

/// Parse a single wps:wsp into ShapeRun. `sx,sy` scale the shape's spPr/xfrm
/// from group child coord space to page EMU; `group_off_pt_*` are the group origin
/// on the page (in pt) so the shape's off.x/off.y (in child coord space) can be
/// translated to page-relative pt. For a standalone wsp (no wgp), pass sx=sy=1, group_off=0.
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
    let sp_pr = wsp.children().find(|n| n.is_element() && n.tag_name().name() == "spPr")?;
    let xfrm = sp_pr.children().find(|n| n.is_element() && n.tag_name().name() == "xfrm")?;
    let off = xfrm.children().find(|n| n.is_element() && n.tag_name().name() == "off")?;
    let ext = xfrm.children().find(|n| n.is_element() && n.tag_name().name() == "ext")?;
    let ox = off.attribute("x").and_then(|v| v.parse::<f64>().ok())?;
    let oy = off.attribute("y").and_then(|v| v.parse::<f64>().ok())?;
    let cx = ext.attribute("cx").and_then(|v| v.parse::<f64>().ok())?;
    let cy = ext.attribute("cy").and_then(|v| v.parse::<f64>().ok())?;
    if cx <= 0.0 || cy <= 0.0 { return None; }
    let rotation = xfrm.attribute("rot")
        .and_then(|v| v.parse::<f64>().ok())
        .map(|r| r / 60000.0) // OOXML rotation: 60000ths of a degree
        .unwrap_or(0.0);

    let width_pt = cx * sx / 12700.0;
    let height_pt = cy * sy / 12700.0;
    let anchor_x_pt = anchor_pos_x + group_off_pt_x + ox * sx / 12700.0;
    let anchor_y_pt = anchor_pos_y + group_off_pt_y + oy * sy / 12700.0;

    let cust_geom = sp_pr.children().find(|n| n.is_element() && n.tag_name().name() == "custGeom");
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
            .and_then(|p| p.children().find(|n| n.is_element() && n.tag_name().name() == "avLst"))
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
    if subpaths.is_empty() && preset_geometry.is_none() { return None; }

    // Direct fills on spPr take priority over wps:style/fillRef. ECMA-376
    // §20.1.4.1.30: when no direct fill is set, the shape's appearance comes
    // from the theme's fillStyleLst / bgFillStyleLst entry referenced by idx,
    // recolored using the schemeClr embedded in the fillRef.
    let fill = parse_shape_fill(sp_pr, theme).or_else(|| {
        wsp.children()
            .find(|n| n.is_element() && n.tag_name().name() == "style")
            .and_then(|st| st.children().find(|n| n.is_element() && n.tag_name().name() == "fillRef"))
            .and_then(|fr| resolve_fill_ref(fr, theme))
    });
    let (stroke, stroke_width) = sp_pr.children()
        .find(|n| n.is_element() && n.tag_name().name() == "ln")
        .map(|ln| {
            let has_no_fill = ln.children().any(|n| n.is_element() && n.tag_name().name() == "noFill");
            if has_no_fill { return (None, 0.0); }
            let color = ln.children()
                .find(|n| n.is_element() && n.tag_name().name() == "solidFill")
                .and_then(|sf| resolve_color_element(sf, theme));
            let w_emu = ln.attribute("w").and_then(|v| v.parse::<f64>().ok()).unwrap_or(9525.0);
            (color, w_emu / 12700.0)
        })
        .unwrap_or((None, 0.0));

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
        rotation,
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
    let txbx = wsp.children().find(|n| n.is_element() && n.tag_name().name() == "txbx");
    let body_pr = wsp.children().find(|n| n.is_element() && n.tag_name().name() == "bodyPr");

    let anchor = body_pr
        .and_then(|b| b.attribute("anchor"))
        .map(|s| s.to_string());
    let emu_to_pt = |v: &str| v.parse::<f64>().ok().map(|e| e / 12700.0).unwrap_or(0.0);
    // ECMA-376 §21.1.2.1.1 defaults: lIns=rIns=91440 EMU, tIns=bIns=45720 EMU
    let l = body_pr.and_then(|b| b.attribute("lIns")).map(emu_to_pt).unwrap_or(91440.0 / 12700.0);
    let t = body_pr.and_then(|b| b.attribute("tIns")).map(emu_to_pt).unwrap_or(45720.0 / 12700.0);
    let r = body_pr.and_then(|b| b.attribute("rIns")).map(emu_to_pt).unwrap_or(91440.0 / 12700.0);
    let b = body_pr.and_then(|b| b.attribute("bIns")).map(emu_to_pt).unwrap_or(45720.0 / 12700.0);

    let blocks: Vec<ShapeText> = txbx
        .and_then(|t| t.children().find(|n| n.is_element() && n.tag_name().name() == "txbxContent"))
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
    for r in p.descendants().filter(|n| n.is_element() && n.tag_name().name() == "r") {
        if first_rpr.is_none() {
            first_rpr = child_w(r, "rPr");
        }
        for t in r.descendants().filter(|n| n.is_element() && n.tag_name().name() == "t") {
            if let Some(text_node) = t.text() {
                text.push_str(text_node);
            }
        }
    }
    if text.is_empty() { return None; }

    let alignment = child_w(p, "pPr")
        .and_then(|ppr| child_w(ppr, "jc"))
        .and_then(|jc| attr_w(jc, "val"))
        .unwrap_or_else(|| "left".to_string());

    let (font_size_pt, color, font_family, bold, italic) = if let Some(rpr) = first_rpr {
        let fmt = parse_run_fmt(rpr);
        (
            fmt.font_size.unwrap_or(DEFAULT_FONT_SIZE),
            fmt.color.clone(),
            theme.resolve_font_ref(fmt.font_family_ascii.clone())
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

/// Parse a shape's fill (solidFill or gradFill). Returns None for noFill/missing.
fn parse_shape_fill(sp_pr: roxmltree::Node, theme: &ThemeColors) -> Option<ShapeFill> {
    for child in sp_pr.children().filter(|n| n.is_element()) {
        match child.tag_name().name() {
            "solidFill" => {
                return resolve_color_element(child, theme).map(|c| ShapeFill::Solid { color: c });
            }
            "gradFill" => {
                return parse_grad_fill(child, theme);
            }
            "noFill" => return None,
            _ => {}
        }
    }
    None
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
fn resolve_fill_ref(
    fill_ref: roxmltree::Node,
    theme: &ThemeColors,
) -> Option<ShapeFill> {
    let idx: u32 = fill_ref.attribute("idx")?.parse().ok()?;
    if idx == 0 { return None; }
    let scheme_clr = fill_ref.children()
        .find(|n| n.is_element() && n.tag_name().name() == "schemeClr")
        .and_then(|n| n.attribute("val"))
        .unwrap_or("dk1");

    let xml = theme.theme_xml.as_ref()?;
    let doc = XmlDoc::parse(xml).ok()?;
    let fmt = doc.root_element()
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
    let lst = fmt.children().find(|n| n.is_element() && n.tag_name().name() == lst_name)?;
    let entry = lst.children().filter(|n| n.is_element()).nth(local_idx)?;

    match entry.tag_name().name() {
        "solidFill" => {
            resolve_color_element_with_phclr(entry, theme, scheme_clr)
                .map(|c| ShapeFill::Solid { color: c })
        }
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
    let gs_lst = node.children().find(|n| n.is_element() && n.tag_name().name() == "gsLst")?;
    let mut stops: Vec<GradientStop> = Vec::new();
    for gs in gs_lst.children().filter(|n| n.is_element() && n.tag_name().name() == "gs") {
        let pos = gs.attribute("pos").and_then(|v| v.parse::<f64>().ok()).map(|p| p / 100000.0).unwrap_or(0.0);
        if let Some(color) = resolve_color_element_with_phclr(gs, theme, ph_clr) {
            stops.push(GradientStop { position: pos, color });
        }
    }
    if stops.is_empty() { return None; }

    // Linear direction (a:lin ang = "60000"ths of a degree)
    let (angle, grad_type) = if let Some(lin) = node.children().find(|n| n.is_element() && n.tag_name().name() == "lin") {
        let ang = lin.attribute("ang").and_then(|v| v.parse::<f64>().ok()).map(|a| a / 60000.0).unwrap_or(0.0);
        (ang, "linear".to_string())
    } else if node.children().any(|n| n.is_element() && n.tag_name().name() == "path") {
        (0.0, "radial".to_string())
    } else {
        (0.0, "linear".to_string())
    };

    Some(ShapeFill::Gradient { stops, angle, grad_type })
}

/// Parse <a:custGeom><a:pathLst><a:path w="W" h="H">...</a:path></a:pathLst>.
/// Path coords inside each <a:path> are absolute within W×H; normalize to [0,1].
fn parse_custom_geometry(cust_geom: roxmltree::Node) -> Vec<Vec<PathCmd>> {
    let Some(path_lst) = cust_geom.children().find(|n| n.is_element() && n.tag_name().name() == "pathLst") else {
        return vec![];
    };
    let mut subpaths: Vec<Vec<PathCmd>> = Vec::new();
    for path in path_lst.children().filter(|n| n.is_element() && n.tag_name().name() == "path") {
        let pw = path.attribute("w").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0);
        let ph = path.attribute("h").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0);
        if pw <= 0.0 || ph <= 0.0 { continue; }
        let mut cmds: Vec<PathCmd> = Vec::new();
        for cmd in path.children().filter(|n| n.is_element()) {
            let name = cmd.tag_name().name();
            let pts: Vec<(f64, f64)> = cmd.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "pt")
                .filter_map(|p| {
                    let x = p.attribute("x").and_then(|v| v.parse::<f64>().ok())?;
                    let y = p.attribute("y").and_then(|v| v.parse::<f64>().ok())?;
                    Some((x / pw, y / ph))
                })
                .collect();
            match name {
                "moveTo" => { if let Some(p) = pts.first() { cmds.push(PathCmd::MoveTo { x: p.0, y: p.1 }); } }
                "lnTo" => { if let Some(p) = pts.first() { cmds.push(PathCmd::LineTo { x: p.0, y: p.1 }); } }
                "cubicBezTo" => {
                    if pts.len() >= 3 {
                        cmds.push(PathCmd::CubicBezTo {
                            x1: pts[0].0, y1: pts[0].1,
                            x2: pts[1].0, y2: pts[1].1,
                            x:  pts[2].0, y:  pts[2].1,
                        });
                    }
                }
                "arcTo" => {
                    let wr = cmd.attribute("wR").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / pw;
                    let hr = cmd.attribute("hR").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / ph;
                    let st_ang = cmd.attribute("stAng").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 60000.0;
                    let sw_ang = cmd.attribute("swAng").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 60000.0;
                    cmds.push(PathCmd::ArcTo { wr, hr, st_ang, sw_ang });
                }
                "close" => cmds.push(PathCmd::Close),
                _ => {}
            }
        }
        if !cmds.is_empty() { subpaths.push(cmds); }
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
                let name = if raw_name == "phClr" && !ph_clr.is_empty() { ph_clr } else { raw_name };
                theme.resolve(name)
            }
            "sysClr" => c.attribute("lastClr").map(|v| v.to_uppercase())
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
    let col_widths: Vec<f64> = tbl_grid.map(|g| {
        children_w(g, "gridCol")
            .iter()
            .map(|c| attr_w(*c, "w").map(|v| twips_to_pt(&v)).unwrap_or(72.0))
            .collect()
    }).unwrap_or_default();

    // Table borders
    let borders = tbl_pr.and_then(|p| child_w(p, "tblBorders"))
        .map(|b| parse_table_borders(b))
        .unwrap_or_default();

    // Cell margins
    let (cm_top, cm_bot, cm_left, cm_right) = tbl_pr
        .and_then(|p| child_w(p, "tblCellMar"))
        .map(|m| (
            child_w(m, "top").and_then(|n| attr_w(n, "w")).map(|v| twips_to_pt(&v)).unwrap_or(0.0),
            child_w(m, "bottom").and_then(|n| attr_w(n, "w")).map(|v| twips_to_pt(&v)).unwrap_or(0.0),
            child_w(m, "left").and_then(|n| attr_w(n, "w")).map(|v| twips_to_pt(&v)).unwrap_or(3.6),
            child_w(m, "right").and_then(|n| attr_w(n, "w")).map(|v| twips_to_pt(&v)).unwrap_or(3.6),
        ))
        .unwrap_or((0.0, 0.0, 3.6, 3.6));

    let mut rows = vec![];
    for tr_node in children_w_flat(node, "tr") {
        let row = parse_table_row(tr_node, style_map, num_map, media_map, rel_map, theme, table_style_id.as_deref());
        rows.push(row);
    }

    DocTable {
        col_widths,
        rows,
        borders,
        cell_margin_top: cm_top,
        cell_margin_bottom: cm_bot,
        cell_margin_left: cm_left,
        cell_margin_right: cm_right,
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
        let cell = parse_table_cell(tc_node, style_map, num_map, media_map, rel_map, theme, table_style_id);
        cells.push(cell);
    }

    DocTableRow { cells, row_height, row_height_rule, is_header }
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
    let v_merge = tc_pr.and_then(|p| child_w(p, "vMerge")).map(|m| {
        attr_w(m, "val").map(|v| v == "restart").unwrap_or(false)
    });

    let borders = tc_pr.and_then(|p| child_w(p, "tcBorders"))
        .map(|b| parse_cell_borders(b))
        .unwrap_or_default();

    let background = tc_pr.and_then(|p| child_w(p, "shd"))
        .and_then(|s| attr_w(s, "fill"))
        .filter(|f| f != "auto" && f.len() == 6)
        .map(|f| f.to_lowercase());

    let v_align = tc_pr.and_then(|p| child_w(p, "vAlign"))
        .and_then(|v| attr_w(v, "val"))
        .unwrap_or_else(|| "top".to_string());

    // ECMA-376 §17.18.87 ST_TblWidth:
    //   dxa  — twentieths of a point (1/20pt)
    //   pct  — 50ths of a percent of the table width (e.g. w="2500" = 50%).
    //         We don't know the table width here, so leave None and fall
    //         back to grid allocation like the other non-dxa cases.
    //   auto, nil — width is dictated by content/grid; treat as None.
    let width_pt = tc_pr.and_then(|p| child_w(p, "tcW"))
        .and_then(|w| {
            let wtype = attr_w(w, "type").unwrap_or_else(|| "dxa".to_string());
            match wtype.as_str() {
                "dxa" => attr_w(w, "w").map(|v| twips_to_pt(&v)),
                _ => None,
            }
        });

    // Per-cell margins (ECMA-376 §17.4.42 `<w:tcPr><w:tcMar>`). Each edge,
    // when present, overrides the table-level `<w:tblCellMar>` default; absent
    // edges stay None so the renderer falls back to the table default.
    let tc_mar = tc_pr.and_then(|p| child_w(p, "tcMar"));
    let edge_mar = |name: &str| tc_mar
        .and_then(|m| child_w(m, name))
        .and_then(|n| attr_w(n, "w"))
        .map(|v| twips_to_pt(&v));
    let margin_top = edge_mar("top");
    let margin_bottom = edge_mar("bottom");
    let margin_left = edge_mar("left");
    let margin_right = edge_mar("right");

    // ECMA-376 §17.4.7: a cell may contain paragraphs AND nested tables in
    // any order. element_children_flat unwraps sdt wrappers like elsewhere.
    let mut content: Vec<CellElement> = vec![];
    for child in element_children_flat(node) {
        match child.tag_name().name() {
            "p" => content.push(CellElement::Paragraph(
                parse_paragraph(child, style_map, num_map, media_map, rel_map, theme, table_style_id)
            )),
            "tbl" => content.push(CellElement::Table(
                parse_table(child, style_map, num_map, media_map, rel_map, theme)
            )),
            _ => {}
        }
    }

    DocTableCell {
        content, col_span, v_merge, borders, background, v_align, width_pt,
        margin_top, margin_bottom, margin_left, margin_right,
    }
}

fn parse_table_borders(node: roxmltree::Node) -> TableBorders {
    TableBorders {
        top: child_w(node, "top").map(parse_border_spec),
        bottom: child_w(node, "bottom").map(parse_border_spec),
        left: child_w(node, "left").map(parse_border_spec),
        right: child_w(node, "right").map(parse_border_spec),
        inside_h: child_w(node, "insideH").map(parse_border_spec),
        inside_v: child_w(node, "insideV").map(parse_border_spec),
    }
}

fn parse_cell_borders(node: roxmltree::Node) -> CellBorders {
    CellBorders {
        top: child_w(node, "top").map(parse_border_spec),
        bottom: child_w(node, "bottom").map(parse_border_spec),
        left: child_w(node, "left").map(parse_border_spec),
        right: child_w(node, "right").map(parse_border_spec),
    }
}

fn parse_border_spec(node: roxmltree::Node) -> BorderSpec {
    let style = attr_w(node, "val").unwrap_or_else(|| "none".to_string());
    let width = attr_w(node, "sz").map(|v| {
        v.parse::<f64>().unwrap_or(4.0) / 8.0  // eighth-points → pt
    }).unwrap_or(0.5);
    let color = attr_w(node, "color").filter(|c| c != "auto").map(|c| c.to_lowercase());
    BorderSpec { width, color, style }
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
    if direct.alignment.is_some() { base.alignment = direct.alignment.clone(); }
    if direct.indent_left.is_some() { base.indent_left = direct.indent_left; }
    if direct.indent_right.is_some() { base.indent_right = direct.indent_right; }
    if direct.indent_first.is_some() { base.indent_first = direct.indent_first; }
    if direct.space_before.is_some() { base.space_before = direct.space_before; }
    if direct.space_after.is_some() { base.space_after = direct.space_after; }
    if direct.line_spacing_val.is_some() { base.line_spacing_val = direct.line_spacing_val; }
    if direct.line_spacing_rule.is_some() { base.line_spacing_rule = direct.line_spacing_rule.clone(); }
    if direct.line_spacing_explicit.is_some() { base.line_spacing_explicit = direct.line_spacing_explicit; }
    if direct.outline_level.is_some() { base.outline_level = direct.outline_level; }
    if direct.num_id.is_some() { base.num_id = direct.num_id; }
    if direct.num_level.is_some() { base.num_level = direct.num_level; }
    if direct.tab_stops.is_some() { base.tab_stops = direct.tab_stops.clone(); }
}

fn apply_direct_run(base: &mut RunFmt, direct: &RunFmt) {
    if direct.bold.is_some() { base.bold = direct.bold; }
    if direct.italic.is_some() { base.italic = direct.italic; }
    if direct.underline.is_some() { base.underline = direct.underline; }
    if direct.strikethrough.is_some() { base.strikethrough = direct.strikethrough; }
    if direct.font_size.is_some() { base.font_size = direct.font_size; }
    if direct.color.is_some() { base.color = direct.color.clone(); }
    if direct.font_family_ascii.is_some() { base.font_family_ascii = direct.font_family_ascii.clone(); }
    if direct.font_family_east_asia.is_some() { base.font_family_east_asia = direct.font_family_east_asia.clone(); }
    if direct.background.is_some() { base.background = direct.background.clone(); }
    if direct.vert_align.is_some() { base.vert_align = direct.vert_align.clone(); }
}

fn parse_rels(xml: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    if xml.is_empty() { return map; }
    let doc = match XmlDoc::parse(xml) { Ok(d) => d, Err(_) => return map };
    for rel in doc.root_element().children().filter(|n| n.tag_name().name() == "Relationship") {
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
    entry.by_ref().take(max).read_to_string(&mut s).map_err(|e| e.to_string())?;
    Ok(s)
}

fn read_zip_bytes(zip: &mut Zip, path: &str) -> Result<Vec<u8>, String> {
    let max = ooxml_common::zip::current_max();
    let mut entry = zip.by_name(path).map_err(|e| format!("{}: {}", path, e))?;
    if entry.size() > max {
        return Err(format!("{}: exceeds size limit", path));
    }
    let mut buf = vec![];
    entry.by_ref().take(max).read_to_end(&mut buf).map_err(|e| e.to_string())?;
    Ok(buf)
}
