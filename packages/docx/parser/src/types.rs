use serde::Serialize;
use std::collections::HashMap;

#[derive(Serialize, Debug, Default, Clone)]
#[serde(rename_all = "camelCase")]
pub struct Document {
    pub section: SectionProps,
    pub body: Vec<BodyElement>,
    pub headers: HeadersFooters,
    pub footers: HeadersFooters,
    /// ECMA-376 §17.7.4 theme `<a:fontScheme><a:majorFont><a:latin@typeface>`
    /// — the heading typeface declared by the document theme. `None` when
    /// the document has no theme part or no Latin major typeface is
    /// declared.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub major_font: Option<String>,
    /// ECMA-376 §17.7.4 theme `<a:fontScheme><a:minorFont><a:latin@typeface>`
    /// — the body typeface declared by the document theme. `None` when the
    /// document has no theme part or no Latin minor typeface is declared.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub minor_font: Option<String>,
    /// ECMA-376 §17.8.3.10 — font family classification from `word/fontTable.xml`.
    /// Maps font name to `<w:family @w:val>` (one of: "roman", "swiss", "modern",
    /// "script", "decorative", "auto"). Used by the renderer to select the
    /// correct CSS generic family (roman→serif, swiss→sans-serif, modern→monospace)
    /// without relying on name-pattern heuristics. Empty when fontTable.xml
    /// is absent or malformed.
    #[serde(skip_serializing_if = "HashMap::is_empty")]
    pub font_family_classes: HashMap<String, String>,
    /// ECMA-376 §17.13.5 — track-changes events found in the body. Each entry
    /// is one `<w:ins>` or `<w:del>` block, with the change author / date /
    /// text content. Empty when the document has no tracked changes.
    /// Renderer ignores this; surfaced for tools (MCP, agents).
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub revisions: Vec<DocxRevision>,
    /// ECMA-376 §17.13.4 — `word/comments.xml` flat list. Each comment carries
    /// its id (matches `<w:commentReference w:id>` in the body), author, date,
    /// and plain-text body. Empty when the document has no comments part.
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub comments: Vec<DocxComment>,
    /// ECMA-376 §17.11.10 — `word/footnotes.xml` flat list (id + text).
    /// Excludes the spec-defined separator and continuation-separator entries
    /// (id="-1" / "0"). Empty when the document has no footnotes part.
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub footnotes: Vec<DocxNote>,
    /// ECMA-376 §17.11.4 — `word/endnotes.xml` flat list. Same shape as
    /// `footnotes`.
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub endnotes: Vec<DocxNote>,
    /// ECMA-376 §17.15.1.* — document-wide compatibility / typography settings
    /// from `word/settings.xml`. Currently the Japanese line-breaking (kinsoku)
    /// configuration. `None` when settings.xml carries none of these elements
    /// (the renderer then uses spec defaults: kinsoku ON).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub settings: Option<DocumentSettings>,
}

/// Document-wide settings surfaced from `word/settings.xml`. Only the
/// typography settings the renderer needs are extracted.
#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct DocumentSettings {
    /// §17.15.1.58 `w:kinsoku` — East-Asian line-breaking toggle. `None` means
    /// the element is absent; the spec default is ON, so the renderer treats
    /// `None` and `Some(true)` identically. `Some(false)` disables kinsoku.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub kinsoku: Option<bool>,
    /// §17.15.1.60 `w:noLineBreaksBefore@w:val` — custom set of characters that
    /// cannot begin a line (行頭禁則). When present it REPLACES the application
    /// default set. Multiple per-`w:lang` elements are concatenated.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub no_line_breaks_before: Option<String>,
    /// §17.15.1.59 `w:noLineBreaksAfter@w:val` — custom set of characters that
    /// cannot end a line (行末禁則). Replaces the default when present.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub no_line_breaks_after: Option<String>,
}

/// Single track-changes event extracted from a body `<w:ins>` / `<w:del>`.
#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct DocxRevision {
    /// "insertion" | "deletion"
    pub kind: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    /// `<w:ins w:date>` / `<w:del w:date>` — ISO-8601 timestamp.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub date: Option<String>,
    /// Concatenated plain text. For deletions, this is the deleted text
    /// captured from `<w:delText>` runs.
    pub text: String,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct DocxComment {
    pub id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    /// `<w:initials>` from `word/comments.xml`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub initials: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub date: Option<String>,
    pub text: String,
}

/// ECMA-376 §17.11.2 (endnote) / §17.11.10 (footnote) — one note's block-level
/// content. Retains the full paragraph/run structure (FootnoteText style, the
/// `<w:footnoteRef/>` auto-number placeholder, inline formatting) so the
/// renderer can lay it out faithfully at page bottom / document end, and so the
/// data-only API can surface the plain text.
#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct DocxNote {
    pub id: String,
    /// Block-level content of the note (paragraphs / nested tables), parsed with
    /// the document's styles + numbering. Empty for an empty note.
    pub content: Vec<BodyElement>,
}

#[derive(Serialize, Debug, Default, Clone)]
#[serde(rename_all = "camelCase")]
pub struct HeadersFooters {
    pub default: Option<HeaderFooter>,
    pub first: Option<HeaderFooter>,
    pub even: Option<HeaderFooter>,
}

#[derive(Serialize, Debug, Default, Clone)]
#[serde(rename_all = "camelCase")]
pub struct HeaderFooter {
    pub body: Vec<BodyElement>,
}

#[derive(Serialize, Debug, Default, Clone)]
#[serde(rename_all = "camelCase")]
pub struct SectionProps {
    /// page width in pt (converted from twips)
    pub page_width: f64,
    /// page height in pt
    pub page_height: f64,
    pub margin_top: f64,
    pub margin_right: f64,
    pub margin_bottom: f64,
    pub margin_left: f64,
    /// distance from top of page to header (pt)
    pub header_distance: f64,
    /// distance from bottom of page to footer (pt)
    pub footer_distance: f64,
    /// whether first page has its own header/footer
    pub title_page: bool,
    /// whether even pages have distinct header/footer
    pub even_and_odd_headers: bool,
    /// ECMA-376 §17.6.5 w:docGrid/@w:type ("default" | "lines" |
    /// "linesAndChars" | "snapToChars"). None = default.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub doc_grid_type: Option<String>,
    /// ECMA-376 §17.6.5 w:docGrid/@w:linePitch in pt (converted from twentieths
    /// of a point). When docGridType is "lines" or "linesAndChars", this is
    /// the vertical pitch per grid line — auto line spacing multiplies against
    /// this instead of the font's natural line height.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub doc_grid_line_pitch: Option<f64>,
}

#[derive(Serialize, Debug, Clone)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum BodyElement {
    Paragraph(DocParagraph),
    Table(DocTable),
    /// Page break. `parity` carries section-break parity intent:
    /// `Some("odd")` = oddPage break (next content must start on an odd
    /// 1-based page), `Some("even")` = evenPage. `None` = a plain `nextPage`
    /// or `<w:br w:type="page"/>` break.
    PageBreak {
        #[serde(skip_serializing_if = "Option::is_none")]
        parity: Option<String>,
    },
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct DocParagraph {
    /// "left" | "center" | "right" | "both"
    pub alignment: String,
    /// pt
    pub indent_left: f64,
    /// pt
    pub indent_right: f64,
    /// pt (negative = hanging)
    pub indent_first: f64,
    /// pt
    pub space_before: f64,
    /// pt
    pub space_after: f64,
    /// None = single (1.0), Some(LineSpacing)
    pub line_spacing: Option<LineSpacing>,
    pub numbering: Option<NumberingInfo>,
    /// Explicit tab stops from w:tabs. Empty means use default tab interval.
    pub tab_stops: Vec<TabStop>,
    pub runs: Vec<DocRun>,
    /// Paragraph background hex color (w:shd fill on paragraph)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shading: Option<String>,
    /// Force a page break before this paragraph (w:pageBreakBefore)
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub page_break_before: bool,
    /// Suppress spacing between adjacent same-style paragraphs (w:contextualSpacing)
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub contextual_spacing: bool,
    /// Keep paragraph on the same page as the next paragraph (w:keepNext)
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub keep_next: bool,
    /// Keep all lines of this paragraph on the same page (w:keepLines)
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub keep_lines: bool,
    /// Widow/orphan control (w:widowControl). Default per spec: true.
    pub widow_control: bool,
    /// Paragraph borders (w:pBdr)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub borders: Option<ParagraphBorders>,
    /// Style ID of the applied paragraph style (for contextual spacing resolution)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_id: Option<String>,
    /// Default font size in pt inherited from style + direct pPr/rPr. Used for
    /// sizing empty paragraphs (lines with no runs) correctly.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default_font_size: Option<f64>,
    /// Default font family resolved from the style chain + direct pPr/rPr.
    /// Used to size empty paragraphs (no runs) when the intended font's line
    /// height differs from the fallback (e.g. an empty Meiryo cell that forms a
    /// résumé "bar" must reserve Meiryo's tall line box). None when unresolved.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default_font_family: Option<String>,
    /// ECMA-376 §17.3.1.20 `<w:outlineLvl w:val="N">` (0–8). Resolved through
    /// the style chain: explicit pPr → linked paragraph style → docDefaults.
    /// `None` for body paragraphs that don't appear in the document outline.
    /// Surfaced so MCP / agents can build a heading hierarchy without relying
    /// on the styleId string ("Heading1", "見出し1", etc.).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub outline_level: Option<u32>,
    /// ECMA-376 §17.3.1.6 `<w:bidi>` — right-to-left paragraph. `Some(true)` =
    /// RTL, `Some(false)` = explicitly LTR, `None` = unspecified (inherit).
    /// The renderer consumes this as the paragraph base direction: it seeds the
    /// UAX#9 reordering pass, swaps the left/right indents, and resolves the
    /// `w:jc` start/end alignment edges against it.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bidi: Option<bool>,
    /// ECMA-376 §17.3.1.32 `<w:snapToGrid>` — `Some(false)` opts the paragraph
    /// OUT of the section's document grid, so its lines use natural font
    /// metrics / the line-spacing multiplier directly instead of snapping to the
    /// grid pitch. `None` = inherit (default on). Set on Word's "Footnote Text"
    /// style.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub snap_to_grid: Option<bool>,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphBorders {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<ParaBorderEdge>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom: Option<ParaBorderEdge>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<ParaBorderEdge>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<ParaBorderEdge>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub between: Option<ParaBorderEdge>,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ParaBorderEdge {
    /// "single" | "double" | "dashed" | ...
    pub style: String,
    pub color: Option<String>,
    /// pt (sz / 8)
    pub width: f64,
    /// pt spacing between border and text
    pub space: f64,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct TabStop {
    /// tab stop position in pt (from the left of the paragraph content area)
    pub pos: f64,
    /// "left" | "center" | "right" | "decimal" | "bar" | "clear"
    pub alignment: String,
    /// "none" | "dot" | "hyphen" | "underscore" | "heavy" | "middleDot"
    pub leader: String,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct LineSpacing {
    /// multiplier (e.g. 1.15) or exact pt
    pub value: f64,
    /// "auto" | "exact" | "atLeast"
    pub rule: String,
    /// True when `w:spacing/@w:line` is set on the paragraph's own pPr or on
    /// a named style; false when only docDefault provides it. Controls the
    /// docGrid interaction per ECMA-376 §17.6.5 — inherited-only paragraphs
    /// snap to one grid pitch per line regardless of the multiplier.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub explicit: bool,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct NumberingInfo {
    pub num_id: u32,
    pub level: u32,
    /// "decimal" | "bullet" | "lowerLetter" | "upperLetter" | "lowerRoman" | "upperRoman"
    pub format: String,
    /// resolved text, e.g. "1." or "•"
    pub text: String,
    /// indent for the entire numbered paragraph (pt)
    pub indent_left: f64,
    /// tab stop after bullet/number (pt)
    pub tab: f64,
}

#[derive(Serialize, Debug, Clone)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum DocRun {
    Text(TextRun),
    Image(ImageRun),
    /// `rename_all` on the enum only renames variant tags; the field
    /// `break_type` would otherwise serialize as snake_case while the TS
    /// side reads `breakType`. Re-apply camelCase at the variant level so
    /// the JSON tag matches.
    #[serde(rename_all = "camelCase")]
    Break {
        break_type: BreakType,
    },
    Field(FieldRun),
    Shape(ShapeRun),
    /// An OMML equation (`m:oMath`). `display` = block (`m:oMathPara`).
    #[serde(rename_all = "camelCase")]
    Math {
        nodes: Vec<crate::math::MathNode>,
        display: bool,
        /// Resolved paragraph font size (pt) so the equation matches surrounding text.
        font_size: f64,
    },
}

/// A drawn shape (wps:wsp inside wp:anchor). Positioned like an anchor image
/// and rendered via core's buildCustomPath + paint primitives.
#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct ShapeRun {
    /// pt
    pub width_pt: f64,
    /// pt
    pub height_pt: f64,
    /// anchor X (pt). Ignored if `anchor_x_align` is set.
    pub anchor_x_pt: f64,
    /// anchor Y (pt). Ignored if `anchor_y_align` is set.
    pub anchor_y_pt: f64,
    pub anchor_x_from_margin: bool,
    pub anchor_y_from_para: bool,
    /// ECMA-376 §20.4.3.1 wp:align (positionH/wp:align). When present the
    /// renderer centers / left-aligns / right-aligns the shape within the
    /// container indicated by `anchor_x_from_margin`. Values: "left",
    /// "center", "right" (others fall back to "left").
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor_x_align: Option<String>,
    /// Vertical equivalent of anchor_x_align (positionV/wp:align).
    /// Values: "top", "center", "bottom".
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor_y_align: Option<String>,
    /// ECMA-376 §20.4.2.7 wp14:pctPosHOffset / pctPosVOffset — fraction
    /// of the relativeFrom container's width / height in 1/100,000ths of
    /// a percent. When set, renderer ignores anchor_x_pt / anchor_y_pt
    /// and computes the offset as `container_size * pct / 100000`. The
    /// container is `anchor_*_relative_from` (page / margin / etc.).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pct_pos_h: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pct_pos_v: Option<f64>,
    /// `<wp:positionH/V relativeFrom="…">` — overrides the looser
    /// `anchor_x_from_margin`/`anchor_y_from_para` booleans for the
    /// pct-pos and align paths. Common values: "page", "margin",
    /// "leftMargin"/"rightMargin", "insideMargin"/"outsideMargin",
    /// "topMargin"/"bottomMargin", "paragraph", "line".
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor_x_relative_from: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor_y_relative_from: Option<String>,
    /// ECMA-376 §20.4.2.18 wp14:sizeRelH/sizeRelV — width/height as a
    /// fraction of the relativeFrom container, in 1/100,000ths of a percent.
    /// When set, the renderer uses this in place of `width_pt` / `height_pt`
    /// for layout (text frame, align centering, shape draw rect). The
    /// container is resolved by `width_relative_from` / `height_relative_from`
    /// using the same rules as anchor positioning. `pct == 0` is treated as
    /// "fall back to extent" (matches Word's empirical behavior).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width_pct: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height_pct: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width_relative_from: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height_relative_from: Option<String>,
    /// Parent wgp group dimensions (pt). Set only when this shape is a child
    /// of a `<wpg:wgp>` group; `None` for standalone wsp anchors. The renderer
    /// uses these for align/pctPos math so the GROUP is positioned within
    /// the relativeFrom container, and per-shape offsets (`anchor_x_pt` /
    /// `anchor_y_pt`) carry the child's offset within the group.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub group_width_pt: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub group_height_pt: Option<f64>,
    /// If true, draw the shape behind text (wp:anchor behindDoc="1"). Renderer
    /// should draw background shapes BEFORE body text.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub behind_doc: bool,
    /// Document order within the wp:anchor (for correct z-ordering among shapes
    /// sharing the same behindDoc value). Lower value = drawn first.
    pub z_order: u32,
    /// normalized [0,1] custom path commands (one or more sub-paths). Empty
    /// when `preset_geometry` is set; the renderer chooses between
    /// buildCustomPath (custGeom) and buildShapePath (prstGeom).
    pub subpaths: Vec<Vec<PathCmd>>,
    /// OOXML <a:prstGeom prst="..."> name (e.g. "rect", "ellipse",
    /// "roundRect", "rtTriangle"). Empty when the shape is custGeom.
    /// `adj_values` carries up to four <a:gd name="adj{n}"> values (0–100000
    /// scale) for shapes that support adjustment handles (trapezoid, callouts).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub preset_geometry: Option<String>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub adj_values: Vec<f64>,
    /// Fill (solid or gradient). None = no fill.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill: Option<ShapeFill>,
    /// stroke hex. None = no stroke.
    pub stroke: Option<String>,
    /// stroke width in pt.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub stroke_width: f64,
    /// rotation in degrees (clockwise).
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub rotation: f64,
    /// Text blocks from <wps:txbx><w:txbxContent> — text rendered INSIDE the
    /// shape's bounding box (ECMA-376 §17.3.4.7). Each block is one paragraph
    /// reduced to plain text + the first run's formatting; advanced layout
    /// (numbering, paragraph styles, mixed-format runs within a single
    /// paragraph) isn't supported in shape bodies yet.
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub text_blocks: Vec<ShapeText>,
    /// Vertical anchor for the shape text box: "t" (top), "ctr" (center),
    /// "b" (bottom). Read from <wps:bodyPr @anchor>. Default = "t".
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_anchor: Option<String>,
    /// Body-pr text insets in pt (left/top/right/bottom). Default 0 each.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub text_inset_l: f64,
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub text_inset_t: f64,
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub text_inset_r: f64,
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub text_inset_b: f64,
    /// Wrap mode matching ImageRun.wrap_mode semantics.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap_mode: Option<String>,
}

#[derive(Serialize, Debug, Clone)]
#[serde(tag = "fillType", rename_all = "camelCase")]
pub enum ShapeFill {
    Solid {
        color: String,
    },
    #[serde(rename_all = "camelCase")]
    Gradient {
        stops: Vec<GradientStop>,
        /// degrees: 0 = left→right, 90 = top→bottom
        angle: f64,
        /// "linear" | "radial"
        grad_type: String,
    },
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct GradientStop {
    /// 0.0–1.0
    pub position: f64,
    /// hex 6-char
    pub color: String,
}

/// Custom geometry path command (shape rendering). Mirrors the pptx
/// PathCmd type to keep JSON output compatible with core's buildCustomPath.
#[derive(Serialize, Debug, Clone)]
#[serde(tag = "cmd", rename_all = "camelCase")]
pub enum PathCmd {
    MoveTo {
        x: f64,
        y: f64,
    },
    LineTo {
        x: f64,
        y: f64,
    },
    CubicBezTo {
        x1: f64,
        y1: f64,
        x2: f64,
        y2: f64,
        x: f64,
        y: f64,
    },
    ArcTo {
        wr: f64,
        hr: f64,
        st_ang: f64,
        sw_ang: f64,
    },
    Close,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct FieldRun {
    /// "page" | "numPages" | "other"
    pub field_type: String,
    /// original instruction text (e.g. "PAGE \\* MERGEFORMAT")
    pub instruction: String,
    /// fallback text captured between fldChar separate and end (shown if field_type is "other")
    pub fallback_text: String,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    /// pt
    pub font_size: f64,
    pub color: Option<String>,
    pub font_family: Option<String>,
    pub background: Option<String>,
    /// "super" | "sub" | None
    pub vert_align: Option<String>,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub all_caps: bool,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub small_caps: bool,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub double_strikethrough: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub highlight: Option<String>,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct TextRun {
    pub text: String,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    /// pt
    pub font_size: f64,
    pub color: Option<String>,
    pub font_family: Option<String>,
    pub is_link: bool,
    pub background: Option<String>,
    /// "super" | "sub" | None
    pub vert_align: Option<String>,
    /// Target URL for hyperlinks (from relationships.xml), None if not a link or no URL
    pub hyperlink: Option<String>,
    /// Transform all characters to uppercase (w:caps)
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub all_caps: bool,
    /// Render as small capitals (uppercase at ~80% size, w:smallCaps)
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub small_caps: bool,
    /// Double strikethrough (w:dstrike)
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub double_strikethrough: bool,
    /// OOXML highlight color name: "yellow" | "cyan" | "green" | ... (w:highlight)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub highlight: Option<String>,
    /// ECMA-376 §17.3.3.25 ruby annotation (furigana). Set on the rubyBase
    /// text run; renders as a small inline annotation above the base glyphs.
    /// `text` is the annotation string (e.g. "すわ" above "坐"). `font_size_pt`
    /// is the annotation's font size in pt — Word stores this as `<w:hps>`
    /// (half-points) inside the rubyPr.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ruby: Option<RubyAnnotation>,
    /// ECMA-376 §17.13.5 — set when this run sits inside a `<w:ins>` or
    /// `<w:del>` block. The renderer paints insertions with a per-author
    /// underline and deletions with a per-author strikethrough so reviewers
    /// can see tracked edits inline.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub revision: Option<RunRevision>,
    /// ECMA-376 §17.3.2.30 `<w:rtl>` — complex-script / right-to-left run.
    /// `Some(true)` = RTL, `Some(false)` = explicitly LTR, `None` = unspecified.
    /// The renderer marks a `Some(true)` run as RTL for the UAX#9 pass (forcing
    /// complex-script shaping) and draws it with `ctx.direction = 'rtl'` so the
    /// glyph order is mirrored.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rtl: Option<bool>,
    /// ECMA-376 §17.3.2.7 `<w:cs/>` — complex-script run toggle: cs
    /// formatting applies to ALL characters of the run (§17.3.2.26).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cs: Option<bool>,
    /// ECMA-376 §17.3.2.26 `<w:rFonts w:cs>` — complex-script typeface (theme
    /// references resolved to a literal family). `None` when unspecified.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family_cs: Option<String>,
    /// ECMA-376 §17.3.2.39 `<w:szCs>` — complex-script font size in pt (same
    /// units as `font_size`). `None` when unspecified.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size_cs: Option<f64>,
    /// ECMA-376 §17.3.2.3 `<w:bCs>` — complex-script bold toggle. `None` when
    /// unspecified (renderer falls back to `bold`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold_cs: Option<bool>,
    /// ECMA-376 §17.3.2.17 `<w:iCs>` — complex-script italic toggle. `None`
    /// when unspecified (renderer falls back to `italic`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic_cs: Option<bool>,
    /// ECMA-376 §17.3.2.20 `<w:lang w:bidi>` — the complex-script (RTL) language
    /// tag, lower-cased (e.g. "ar-sa", "ae-ar", "he-il"). Used to decide whether
    /// European digits in a complex-script run are classified as AN (Word's
    /// Arabic/Hebrew digit ordering). `None` when unspecified.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub lang_bidi: Option<String>,
    /// ECMA-376 §17.11.6 / §17.11.7 / §17.11.16 / §17.11.17 — set when this run
    /// is a footnote/endnote reference mark (`<w:footnoteReference>` in the body,
    /// `<w:footnoteRef>` inside the note content, and the endnote equivalents).
    /// `text` holds the raw `@w:id` as a fallback; the renderer overrides the
    /// displayed glyph with the note's sequential number (the displayed number is
    /// the note's 1-based position, not the raw id). `None` for ordinary runs.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub note_ref: Option<NoteRef>,
}

/// A footnote / endnote reference marker. The displayed number is resolved by
/// the renderer from the note's sequential position (ECMA-376 numbering), so
/// only the kind + id need to travel through the model.
#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct NoteRef {
    /// "footnote" | "endnote".
    pub kind: String,
    /// The `@w:id` linking the marker to its note in footnotes.xml / endnotes.xml.
    pub id: String,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct RunRevision {
    /// "insertion" or "deletion".
    pub kind: String,
    /// `<w:ins w:author>` / `<w:del w:author>`. Used by the renderer to pick
    /// a stable per-author colour (modulo a fixed palette).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    /// `<w:ins w:date>` / `<w:del w:date>` ISO-8601 timestamp.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub date: Option<String>,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct RubyAnnotation {
    pub text: String,
    /// pt
    pub font_size_pt: f64,
}

/// One paragraph of text rendered inside a shape (`<wps:txbx><w:txbxContent>`).
/// Reduced to a single combined string + the first run's effective formatting
/// — the shape-text layouts in our current sample corpus carry one run per
/// paragraph, so we don't yet support mixed runs / full inline layout here.
#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct ShapeText {
    pub text: String,
    pub font_size_pt: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub bold: bool,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub italic: bool,
    /// Paragraph alignment ("left" | "center" | "right" | "both").
    pub alignment: String,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ImageRun {
    /// data:<mime>;base64,...
    pub data_url: String,
    /// pt
    pub width_pt: f64,
    /// pt
    pub height_pt: f64,
    /// true = wp:anchor (absolute page position), false = wp:inline (flows with text)
    pub anchor: bool,
    /// X offset in pt (anchor only).  Interpretation depends on anchor_x_from_margin.
    pub anchor_x_pt: f64,
    /// Y offset in pt (anchor only).  Interpretation depends on anchor_y_from_para.
    pub anchor_y_pt: f64,
    /// If true anchorXPt is relative to the left margin; add section.marginLeft to get page-abs X.
    /// If false anchorXPt is already page-absolute.
    pub anchor_x_from_margin: bool,
    /// If true anchorYPt is relative to the paragraph's top Y in the renderer (add paragraphTopPx).
    /// If false anchorYPt is already page-absolute.
    pub anchor_y_from_para: bool,
    /// When set, the renderer should replace all pixels of this hex color (e.g. "FFFFFF") with
    /// full transparency. Used to implement a:clrChange (make-background-transparent).
    pub color_replace_from: Option<String>,
    /// Wrap mode for anchor images. One of:
    ///   "square" | "topAndBottom" | "none" | "tight" | "through"
    /// Inline images and anchors without an explicit wrap element use "none".
    /// "tight" and "through" fall back to "square" rendering in the MVP.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap_mode: Option<String>,
    /// distT (top padding, pt). Anchor-only.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub dist_top: f64,
    /// distB (bottom padding, pt). Anchor-only.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub dist_bottom: f64,
    /// distL (left padding, pt). Anchor-only.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub dist_left: f64,
    /// distR (right padding, pt). Anchor-only.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub dist_right: f64,
    /// wrapSquare/wrapTight "wrapText" attribute: "bothSides" | "left" | "right" | "largest".
    /// Defaults to "bothSides" (equivalent).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap_side: Option<String>,
}

fn is_zero_f64(v: &f64) -> bool {
    *v == 0.0
}

#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub enum BreakType {
    Line,
    Page,
    Column,
    /// ECMA-376 §17.3.1.20 — Word's saved hint about where the previous
    /// render placed a page break. A layout cache, not authoritative: we
    /// paginate ourselves and ignore it uniformly (the parser strips these
    /// runs). Kept distinct from `Page` only so the stripping stays explicit.
    /// Serialized as `renderedPage` for the TS side.
    RenderedPage,
}

// ===== Table =====

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct DocTable {
    /// column widths in pt
    pub col_widths: Vec<f64>,
    pub rows: Vec<DocTableRow>,
    /// table-level borders
    pub borders: TableBorders,
    /// cell margin defaults pt
    pub cell_margin_top: f64,
    pub cell_margin_bottom: f64,
    pub cell_margin_left: f64,
    pub cell_margin_right: f64,
    /// table horizontal alignment on the page: "left" | "center" | "right" (w:tblPr/w:jc).
    pub jc: String,
    /// ECMA-376 §17.4.52 `<w:tblLayout w:type>`. "fixed" or "autofit".
    /// Absent in the source ⇒ None, which the renderer treats as the spec
    /// default "autofit" (size columns by preferred widths). When "fixed" the
    /// renderer uses the tblGrid widths verbatim (scaled to fit), ignoring tcW.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub layout: Option<String>,
    /// ECMA-376 §17.4.63 `<w:tblW>` preferred table width. `width_pt` is set
    /// only for type="dxa"; `width_pct` carries type="pct" (value in 50ths of a
    /// percent of the available content width). type="auto"/"nil"/0 ⇒ both None
    /// (table width is dictated by its columns).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width_pt: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width_pct: Option<f64>,
    /// ECMA-376 §17.4.1 `<w:bidiVisual>` — render columns in right-to-left
    /// (visual) order. `Some(true)` = RTL columns, `Some(false)` = explicitly
    /// LTR, `None` = unspecified. When `Some(true)` the renderer mirrors the
    /// grid so logical column 0 is placed rightmost and flips the per-cell
    /// left/right borders accordingly.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bidi_visual: Option<bool>,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableBorders {
    pub top: Option<BorderSpec>,
    pub bottom: Option<BorderSpec>,
    pub left: Option<BorderSpec>,
    pub right: Option<BorderSpec>,
    pub inside_h: Option<BorderSpec>,
    pub inside_v: Option<BorderSpec>,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct BorderSpec {
    /// pt
    pub width: f64,
    pub color: Option<String>,
    /// "single" | "double" | "none" | ...
    pub style: String,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct DocTableRow {
    pub cells: Vec<DocTableCell>,
    /// pt, None = auto
    pub row_height: Option<f64>,
    /// ECMA-376 §17.4.80 w:trHeight/@hRule. "auto" (default) = treat row_height
    /// as informational and size to content; "atLeast" = use as a lower bound;
    /// "exact" = honor the value exactly. Stored as the raw OOXML token so the
    /// renderer can branch without re-parsing.
    pub row_height_rule: String,
    pub is_header: bool,
}

/// One block-level entry inside a table cell. ECMA-376 §17.4.7 (w:tc) allows
/// any BlockLevelElts inside a cell — paragraphs and nested tables.
#[derive(Serialize, Debug, Clone)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum CellElement {
    Paragraph(DocParagraph),
    Table(DocTable),
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct DocTableCell {
    pub content: Vec<CellElement>,
    pub col_span: u32,
    /// VMerge: None = no merge, Some(true) = start of vertical merge, Some(false) = continuation
    pub v_merge: Option<bool>,
    pub borders: CellBorders,
    /// hex color background
    pub background: Option<String>,
    /// "top" | "center" | "bottom"
    pub v_align: String,
    /// ECMA-376 §17.4.71 `<w:tcW>` preferred cell width, type="dxa", in pt.
    /// Drives autofit column sizing (the per-column preferred width is the max
    /// over the cells anchored in it).
    pub width_pt: Option<f64>,
    /// `<w:tcW>` type="pct": 50ths of a percent of the available content width.
    /// Resolved against the available width in the renderer (parse time has no
    /// table width). None unless the cell uses type="pct".
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width_pct: Option<f64>,
    /// Per-cell margins from `<w:tcPr><w:tcMar>` (ECMA-376 §17.4.42), in pt.
    /// Each edge overrides the table-level `<w:tblCellMar>` default (§17.4.41)
    /// when present; None = inherit the table default. Used e.g. by résumé
    /// templates that add a top margin to a single cell to space its content.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub margin_top: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub margin_bottom: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub margin_left: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub margin_right: Option<f64>,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct CellBorders {
    pub top: Option<BorderSpec>,
    pub bottom: Option<BorderSpec>,
    pub left: Option<BorderSpec>,
    pub right: Option<BorderSpec>,
}
