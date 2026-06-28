// ===== Output JSON model (mirrors Rust types) =====

import type { MathNode } from '@silurus/ooxml-core';

export interface DocxDocumentModel {
  section: SectionProps;
  body: BodyElement[];
  headers: HeadersFooters;
  footers: HeadersFooters;
  /** Theme `<a:fontScheme><a:majorFont><a:latin@typeface>` (heading face). */
  majorFont?: string;
  /** Theme `<a:fontScheme><a:minorFont><a:latin@typeface>` (body face). */
  minorFont?: string;
  /**
   * ECMA-376 §17.8.3.10 — font family classification from `word/fontTable.xml`.
   * Maps font name to `<w:family @w:val>`: "roman" | "swiss" | "modern" |
   * "script" | "decorative" | "auto". The renderer uses this as the primary
   * source for serif/sans-serif decisions (roman→serif, swiss→sans-serif,
   * modern→monospace), falling back to name-pattern matching only when the
   * entry is absent or classified as "auto".
   */
  fontFamilyClasses?: Record<string, string>;
  /** ECMA-376 §17.13.5 — flat list of `<w:ins>` / `<w:del>` events in the
   *  body. Each entry carries author / date / text. The renderer marks
   *  runs inline via {@link DocxTextRun.revision}; this array is primarily for
   *  tooling (MCP, agents, change-summary panels). */
  revisions?: DocRevision[];
  /** ECMA-376 §17.13.4 — `word/comments.xml`. Each comment carries id,
   *  author, initials, date, and plain-text body. */
  comments?: DocComment[];
  /** ECMA-376 §17.11.10 — `word/footnotes.xml` (id + text). Excludes the
   *  spec-defined separator / continuation-separator entries. */
  footnotes?: DocNote[];
  /** ECMA-376 §17.11.4 — `word/endnotes.xml` (id + text). Same shape as
   *  `footnotes`. */
  endnotes?: DocNote[];
  /** ECMA-376 §17.15.1.* — document-wide compatibility / typography settings
   *  from `word/settings.xml`. Currently carries the Japanese line-breaking
   *  (kinsoku) configuration. Absent when settings.xml has no relevant
   *  elements (the renderer then uses spec defaults: kinsoku ON). */
  settings?: DocSettings;
}

export interface DocSettings {
  /** §17.15.1.58 `w:kinsoku` — East-Asian line-breaking toggle. `undefined`
   *  means the element is absent; the spec default is ON (treated as `true`). */
  kinsoku?: boolean;
  /** §17.15.1.60 `w:noLineBreaksBefore@w:val` — custom set of characters that
   *  cannot begin a line (行頭禁則). When present it REPLACES the application
   *  default set. Word's per-`w:lang` sets are merged into one string. */
  noLineBreaksBefore?: string;
  /** §17.15.1.59 `w:noLineBreaksAfter@w:val` — custom set of characters that
   *  cannot end a line (行末禁則). Replaces the default when present. */
  noLineBreaksAfter?: string;
  /** §22.1.2.30 `m:mathPr/m:defJc@m:val` — document-wide default math
   *  justification (ST_Jc math: left|right|center|centerGroup). `undefined`
   *  ⇒ the renderer uses the spec default `centerGroup`. */
  mathDefJc?: string;
}

export interface DocRevision {
  /** "insertion" | "deletion" */
  kind: 'insertion' | 'deletion' | string;
  author?: string;
  /** ISO-8601 timestamp */
  date?: string;
  text: string;
}

export interface DocComment {
  id: string;
  author?: string;
  initials?: string;
  date?: string;
  text: string;
}

export interface DocNote {
  id: string;
  /** ECMA-376 §17.11.2 / §17.11.10 — the note's block-level content
   *  (paragraphs / nested tables), parsed with the document's styles +
   *  numbering. The leading run is the `<w:footnoteRef/>` auto-number marker
   *  (carries a {@link DocxTextRun.noteRef}). Use {@link noteText} to extract
   *  the plain-text body without the marker. */
  content: BodyElement[];
}

/** Flatten a footnote/endnote's content to its plain-text body, excluding the
 *  auto-number reference marker. Convenience for data-only consumers
 *  (the renderer draws {@link DocNote.content} directly). */
export function noteText(note: DocNote): string {
  const parts: string[] = [];
  for (const el of note.content) {
    if (el.type !== 'paragraph') continue;
    let s = '';
    for (const run of (el as DocParagraph).runs) {
      if (run.type === 'text' && !(run as DocxTextRun).noteRef) {
        s += (run as DocxTextRun).text;
      }
    }
    s = s.trim();
    if (s) parts.push(s);
  }
  return parts.join(' ');
}

export interface HeadersFooters {
  default: HeaderFooter | null;
  first: HeaderFooter | null;
  even: HeaderFooter | null;
}

export interface HeaderFooter {
  body: BodyElement[];
}

export interface SectionProps {
  pageWidth: number;   // pt
  pageHeight: number;  // pt
  marginTop: number;   // pt
  marginRight: number;
  marginBottom: number;
  marginLeft: number;
  headerDistance: number;   // pt — top of page to header
  footerDistance: number;   // pt — bottom of page to footer
  titlePage: boolean;
  evenAndOddHeaders: boolean;
  /** ECMA-376 §17.6.22 ST_SectionMark — the body (final) section's `<w:type>`
   *  start type ("continuous" | "nextPage" | "oddPage" | "evenPage"). Governs
   *  how the last section begins relative to the previous one; consumed by the
   *  paginator at the boundary INTO the final section. Absent ⇒ "nextPage" (the
   *  spec default). Non-final sections carry their start type on their own
   *  SectionBreak marker. */
  sectionStart?: string | null;
  /** ECMA-376 §17.6.5 w:docGrid/@w:type — "default" | "lines" | "linesAndChars" | "snapToChars". */
  docGridType?: string | null;
  /** ECMA-376 §17.6.5 w:docGrid/@w:linePitch in pt. When docGridType is "lines" or
   *  "linesAndChars", auto line spacing multiplies against this pitch instead of
   *  the font's natural line height. */
  docGridLinePitch?: number | null;
  /** ECMA-376 §17.6.5 w:docGrid/@w:charSpace (ST_DecimalNumber, signed). The
   *  raw character-grid spacing in 1/4096ths of an em (NOT twips). When
   *  docGridType is "linesAndChars" or "snapToChars", every full-width East-
   *  Asian glyph occupies a fixed cell of width `fontSizePt + charSpace/4096` pt
   *  (negative = tighter). Absent ⇒ East-Asian glyphs keep their natural em
   *  advance. */
  docGridCharSpace?: number | null;
  /** ECMA-376 §17.6.4 `<w:cols>` — newspaper-style multi-column layout. `null`
   *  (or absent) ⇒ single full-width column (unchanged behavior). When present,
   *  body text flows top-to-bottom through `count` columns (newspaper fill);
   *  see {@link computeColumns}. */
  columns?: ColumnsSpec | null;
}

/** ECMA-376 §17.6.4 `<w:cols>` — the section's multi-column configuration. */
export interface ColumnsSpec {
  /** `@w:num` — number of columns (>= 2 when emitted). */
  count: number;
  /** `@w:space` in pt — inter-column gap for equal-width columns (default 36pt
   *  = 720 twips per the spec). */
  spacePt: number;
  /** `@w:equalWidth` (default true) — all columns share one width + `spacePt`.
   *  When false, `cols` carries explicit per-column geometry. */
  equalWidth: boolean;
  /** `@w:sep` — draw vertical separator rules between columns. */
  sep: boolean;
  /** Per-column `<w:col>` entries (width + trailing space, pt). Empty when
   *  `equalWidth` is true. */
  cols: ColSpec[];
}

/** ECMA-376 §17.6.3 `<w:col>` — one column's width and trailing space (pt). */
export interface ColSpec {
  widthPt: number;
  spacePt: number;
}

/** ECMA-376 §17.6.4 — a single newspaper column resolved to its page-absolute
 *  left edge (`xPt`) and text width (`wPt`), in pt. Produced by `computeColumns`
 *  (in the renderer) from a section's {@link ColumnsSpec} + the page geometry. */
export interface ColumnGeom {
  xPt: number;
  wPt: number;
}

export type BodyElement =
  | { type: 'paragraph' } & DocParagraph
  | { type: 'table' } & DocTable
  | { type: 'pageBreak'; parity?: 'odd' | 'even' }
  /** ECMA-376 §17.3.1.20 `<w:br w:type="column"/>` — force the following content
   *  into the next newspaper column (or the next page's first column when
   *  already in the last column). Hoisted to the body level by the parser. */
  | { type: 'columnBreak' }
  /** ECMA-376 §17.6.x — a section boundary in the body. A `<w:sectPr>` carried in
   *  a paragraph's `pPr` (or a loose mid-body one) defines the section that ENDS
   *  here; the FINAL section's settings live on {@link DocxDocumentModel.section}.
   *  `columns` is the ENDING section's `<w:cols>` (§17.6.4; absent ⇒ single
   *  full-width column — the spec default for `@w:num` 1). `kind` is the
   *  ST_SectionMark (§17.18.79) controlling how the NEXT section starts:
   *  "continuous" (same page), "nextPage" (default; page break), "oddPage" /
   *  "evenPage" (page break + parity padding). The paginator switches its active
   *  newspaper-column geometry per section at each marker — fixing the regression
   *  where every section inherited the body-level section's columns.
   *
   *  `headers`/`footers` carry the ENDING section's resolved (§17.10.1-inherited)
   *  header/footer set, and `titlePage` its own `<w:titlePg>` flag, so the renderer
   *  can pick the active section's header/footer per page (mirroring how `columns`
   *  drives per-section column geometry). The body-level (final) section's sets
   *  live on {@link DocxDocumentModel.section}/`.headers`/`.footers` instead. */
  | {
      type: 'sectionBreak';
      kind: 'continuous' | 'nextPage' | 'oddPage' | 'evenPage' | string;
      columns?: ColumnsSpec | null;
      headers?: HeadersFooters;
      footers?: HeadersFooters;
      titlePage?: boolean;
    };

/** A BodyElement annotated with a line range to render. Set when the
 *  paginator splits a paragraph that doesn't fit on a single page —
 *  `lineSlice` constrains which laid-out line indices the renderer paints,
 *  and the renderer adjusts the starting Y so the slice's first line begins
 *  at the page's content top. `colIndex` records which newspaper column (0-based)
 *  the element was placed in (ECMA-376 §17.6.4); absent / 0 for single-column
 *  sections. */
export type PaginatedBodyElement = BodyElement & {
  lineSlice?: { start: number; end: number };
  colIndex?: number;
  /** ECMA-376 §17.6.4 — the column geometry of the SECTION this element belongs
   *  to (per-section newspaper columns). Stamped by the paginator so the renderer
   *  resolves `colIndex` against the right section's columns even when two
   *  sections share a page (a "continuous" section break). Absent ⇒ the renderer
   *  uses the page-level `columns` it was given (single-section / header / footer
   *  paths), so single-section documents are unaffected. Runtime-only — never
   *  emitted by the parser. */
  colGeom?: ColumnGeom[];
  /** ECMA-376 §17.6.4 — page-absolute Y (pt) of the TOP of the multi-column
   *  region this element belongs to on its page. For a section started by a
   *  "continuous" section break (§17.18.79) the columns begin partway down the
   *  page (below the preceding single-column content), not at the page content
   *  top; the paginator computes that origin once (front-loaded layout) and
   *  stamps it so the renderer resets a column's vertical cursor to the REGION
   *  top — never the page top. Also carries the region's bottom (max column
   *  depth) onto the FIRST element of the following section so it clears all
   *  columns. Absent ⇒ the renderer uses the page content top (single-column /
   *  page-spanning section). Runtime-only — never emitted by the parser. */
  colTopPt?: number;
  /** ECMA-376 §17.10.1 — the resolved header/footer set + `<w:titlePg>` flag of
   *  the SECTION this element belongs to. Stamped by the paginator (from the
   *  upcoming `SectionBreak` marker, or the body-level section for the final
   *  section) so the renderer picks the active section's header/footer per page —
   *  mirroring how `colGeom` resolves per-section columns. Absent ⇒ the renderer
   *  falls back to the body-level `doc.headers`/`doc.footers`/`section.titlePage`.
   *  Runtime-only — never emitted by the parser. */
  sectionHF?: { headers: HeadersFooters; footers: HeadersFooters; titlePage: boolean };
};

export interface DocParagraph {
  /**
   * ECMA-376 §17.18.44 ST_Jc. Renderer honors left, start, center, right, end,
   * both, distribute. Other values (kashida variants, numTab, thaiDistribute)
   * are treated as start-aligned.
   */
  alignment:
    | 'left' | 'start'
    | 'center'
    | 'right' | 'end'
    | 'justify' | 'both'
    | 'distribute'
    | string;
  indentLeft: number;   // pt
  indentRight: number;  // pt
  indentFirst: number;  // pt
  spaceBefore: number;  // pt
  spaceAfter: number;   // pt
  lineSpacing: LineSpacing | null;
  numbering: NumberingInfo | null;
  tabStops: TabStop[];
  runs: DocRun[];
  /** Paragraph background hex color (w:shd fill) */
  shading?: string | null;
  /** Force a page break before this paragraph (w:pageBreakBefore) */
  pageBreakBefore?: boolean;
  /** Suppress spacing between adjacent same-style paragraphs (w:contextualSpacing) */
  contextualSpacing?: boolean;
  /** Keep paragraph on same page as the next paragraph (w:keepNext) */
  keepNext?: boolean;
  /** Keep all lines of this paragraph on the same page (w:keepLines) */
  keepLines?: boolean;
  /** Widow/orphan control (w:widowControl). ECMA-376 default is true. */
  widowControl?: boolean;
  /** Paragraph borders (w:pBdr) */
  borders?: ParagraphBorders | null;
  /** Style ID of the applied paragraph style */
  styleId?: string | null;
  /** Default font size (pt) inherited from style + direct pPr/rPr. Falls back to 10pt. */
  defaultFontSize?: number;
  /** Default font family resolved from the style chain. Used to size empty
   *  paragraphs (no runs) with the intended font's line metrics. */
  defaultFontFamily?: string | null;
  /**
   * ECMA-376 §17.3.1.6 `<w:bidi>` — right-to-left paragraph. `true` = RTL,
   * `false` = explicitly LTR, absent = unspecified (inherit). The renderer uses
   * this as the paragraph base direction: it seeds the UAX#9 reordering pass
   * (`computeLineVisualOrder`), swaps the left/right indents, resolves the
   * `w:jc` start/end edges (`resolveAlignEdge`), and lays out lines from the
   * right.
   */
  bidi?: boolean;
  /**
   * ECMA-376 §17.3.1.32 `<w:snapToGrid>` — when `false`, this paragraph opts out
   * of the section's document grid (`w:docGrid`): its lines use natural font
   * metrics / the line-spacing multiplier directly instead of snapping to the
   * grid pitch. `undefined` = inherit (default on). Set on Word's "Footnote
   * Text" style, so footnote bodies use compact natural line height.
   */
  snapToGrid?: boolean;
  /**
   * ECMA-376 §17.3.1.11 `<w:framePr>` — text-frame / drop-cap properties.
   * Present ⇒ this paragraph is part of a text frame; the renderer positions it
   * as a frame (drop cap or generic frame) and registers a wrap exclusion so
   * following body text flows around it. Absent ⇒ ordinary in-flow paragraph.
   */
  framePr?: FramePr;
}

/**
 * ECMA-376 §17.3.1.11 `<w:framePr>` — text-frame / drop-cap properties.
 *
 * Lengths are pt (parser converts from twips). Per the spec, `x`/`y` are
 * ignored when `xAlign`/`yAlign` are set, and for a drop cap `y`/`yAlign` are
 * ignored entirely while `lines` drives the height. `w`/`h`/`x`/`y` are
 * `undefined` when the attribute was absent (distinct from an explicit 0).
 */
export interface FramePr {
  /** ST_DropCap (§17.18.20): 'none' | 'drop' | 'margin'. Default 'none'. */
  dropCap: 'none' | 'drop' | 'margin' | string;
  /** §17.3.1.11 `lines` — drop-cap vertical height in anchor lines. Default 1. */
  lines: number;
  /** ST_Wrap (§17.18.104): 'around'|'auto'|'none'|'notBeside'|'through'|'tight'. Default 'around'. */
  wrap: 'around' | 'auto' | 'none' | 'notBeside' | 'through' | 'tight' | string;
  /** ST_HAnchor (§17.18.35): 'text'(=column) | 'margin' | 'page'. Default 'page'. */
  hAnchor: 'text' | 'margin' | 'page' | string;
  /** ST_VAnchor (§17.18.100): 'text' | 'margin' | 'page'. Default 'page'. */
  vAnchor: 'text' | 'margin' | 'page' | string;
  /** ST_HeightRule (§17.18.37): 'auto' | 'atLeast' | 'exact'. Default 'auto'. */
  hRule: 'auto' | 'atLeast' | 'exact' | string;
  /** hSpace — min wrap padding L/R when wrap='around' (pt). Default 0. */
  hSpace: number;
  /** vSpace — min wrap padding top/bottom (pt). Default 0. */
  vSpace: number;
  /** w — exact frame width (pt). Absent ⇒ auto (max content line width). */
  w?: number;
  /** h — frame height (pt). Meaning gated by hRule. Absent ⇒ auto. */
  h?: number;
  /** x — absolute horizontal offset from hAnchor (pt). Ignored when xAlign set. */
  x?: number;
  /** y — absolute vertical offset from vAnchor (pt). Ignored when yAlign set / drop cap. */
  y?: number;
  /** ST_XAlign (§22.9.2.18): 'left'|'center'|'right'|'inside'|'outside'. Supersedes x. */
  xAlign?: 'left' | 'center' | 'right' | 'inside' | 'outside' | string;
  /** ST_YAlign (§22.9.2.20): 'inline'|'top'|'center'|'bottom'|'inside'|'outside'. Supersedes y. */
  yAlign?: 'inline' | 'top' | 'center' | 'bottom' | 'inside' | 'outside' | string;
}

export interface ParagraphBorders {
  top: ParaBorderEdge | null;
  bottom: ParaBorderEdge | null;
  left: ParaBorderEdge | null;
  right: ParaBorderEdge | null;
  between: ParaBorderEdge | null;
}

export interface ParaBorderEdge {
  style: string;
  color: string | null;
  /** pt (sz / 8) */
  width: number;
  /** pt spacing between border and text */
  space: number;
}

/** ECMA-376 §17.3.2.4 `<w:bdr>` — a run-level border drawn as a box around the
 *  run's text. Parallel to {@link ParaBorderEdge} but applies per run. */
export interface DocxRunBorder {
  /** "single" | "double" | "dashed" | ... (w:bdr/@w:val) */
  style: string;
  /** hex 6-char, or null for automatic (renderer falls back to text color) */
  color?: string | null;
  /** pt (sz / 8) */
  width: number;
  /** pt spacing between the border and the run text (w:space) */
  space: number;
}

export interface TabStop {
  /** tab stop position in pt (from the left of paragraph content area) */
  pos: number;
  alignment: 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'clear';
  leader: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
}

export interface LineSpacing {
  value: number;   // multiplier (auto) or pt (exact/atLeast)
  rule: 'auto' | 'exact' | 'atLeast';
  /** True when `w:spacing/@w:line` was set on the paragraph's own pPr or on a
   *  named style (not inherited solely from docDefault). Per ECMA-376 §17.6.5,
   *  an inherited-only paragraph in a docGrid section snaps to one grid pitch
   *  per line, ignoring the multiplier. Defaults to false on JSON parse. */
  explicit?: boolean;
}

export interface NumberingInfo {
  numId: number;
  level: number;
  format: string;
  text: string;       // resolved bullet text, e.g. "1." or "•"
  indentLeft: number; // pt
  tab: number;        // pt
  /** ECMA-376 §17.9.28 `<w:suff>` — "tab" (default) | "space" | "nothing".
   *  Where body text starts after the marker on the first line. */
  suff: string;
  /** ECMA-376 §17.9.8 `<w:lvlJc>` — marker justification: "left" (default) |
   *  "right" (period-aligned numerals: marker RIGHT edge at the hanging-indent
   *  position) | "center". The renderer offsets the marker draw accordingly.
   *  Always emitted by the parser; optional here so hand-built fixtures may omit
   *  it (the renderer treats absent as "left"). */
  jc?: string;
  /** ECMA-376 §17.3.2.26 ascii axis for the marker glyph, resolved through the
   *  level's `rPr` (§17.9.6) merged over the paragraph's run formatting. The
   *  renderer draws Latin marker chars (e.g. a decimal "1") with this family, so
   *  a heading whose ascii=Times renders its auto-number in Times (serif) even
   *  when eastAsia=Gothic. Absent ⇒ the renderer falls back to its default. */
  fontFamily?: string | null;
  /** ECMA-376 §17.3.2.26 eastAsia axis for the marker glyph (same resolution as
   *  {@link NumberingInfo.fontFamily}). The renderer draws CJK marker chars with
   *  this family. Absent ⇒ the renderer falls back to
   *  {@link NumberingInfo.fontFamily}. */
  fontFamilyEastAsia?: string | null;
  /** ECMA-376 §17.9.9/§17.9.20 — when the level uses a `<w:lvlPicBulletId>`,
   *  the marker is this image (zip path, e.g. `word/media/image1.gif`), drawn in
   *  place of {@link NumberingInfo.text}. Absent ⇒ ordinary text/glyph marker. */
  picBulletImagePath?: string;
  /** MIME type of {@link NumberingInfo.picBulletImagePath} (e.g. `image/gif`). */
  picBulletMimeType?: string;
  /** Picture-bullet marker width in pt (from the `<v:shape style="width">`). */
  picBulletWidthPt?: number;
  /** Picture-bullet marker height in pt (from the `<v:shape style="height">`). */
  picBulletHeightPt?: number;
}

export type DocRun =
  | { type: 'text' } & DocxTextRun
  | { type: 'image' } & ImageRun
  | { type: 'break'; breakType: 'line' | 'page' | 'column' }
  | { type: 'field' } & FieldRun
  | { type: 'shape' } & ShapeRun
  | { type: 'math'; nodes: MathNode[]; display: boolean; fontSize: number; jc?: string };

export type PathCmd =
  | { cmd: 'moveTo'; x: number; y: number }
  | { cmd: 'lineTo'; x: number; y: number }
  | { cmd: 'cubicBezTo'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { cmd: 'arcTo'; wr: number; hr: number; stAng: number; swAng: number }
  | { cmd: 'close' };

export interface ShapeRun {
  widthPt: number;
  heightPt: number;
  /** X offset in pt */
  anchorXPt: number;
  /** Y offset in pt */
  anchorYPt: number;
  anchorXFromMargin: boolean;
  anchorYFromPara: boolean;
  /** ECMA-376 §20.4.3.1 wp:align horizontal: "left" | "center" | "right" |
   *  "inside" | "outside". When set the renderer aligns the shape inside the
   *  container indicated by `anchorXFromMargin` and ignores `anchorXPt`. */
  anchorXAlign?: string | null;
  /** Vertical equivalent of anchorXAlign: "top" | "center" | "bottom". */
  anchorYAlign?: string | null;
  /** ECMA-376 §20.4.2.7 wp14:pctPosHOffset / pctPosVOffset normalised to a
   *  fraction in `[0, 1]`. When set the renderer multiplies it by the
   *  relativeFrom container's width / height and uses that as the
   *  shape's offset within the container, ignoring anchorXPt / anchorYPt. */
  pctPosH?: number | null;
  pctPosV?: number | null;
  /** Raw `relativeFrom` value from `<wp:positionH>` / `<wp:positionV>` —
   *  e.g. "page", "margin", "topMargin", "rightMargin",
   *  "insideMargin", "paragraph", "line". Drives container selection
   *  for both pctPos* and anchor*Align positioning. */
  anchorXRelativeFrom?: string | null;
  anchorYRelativeFrom?: string | null;
  /** ECMA-376 §20.4.2.18 wp14:sizeRelH/sizeRelV — width/height as a
   *  fraction of the relativeFrom container. When set, the renderer uses
   *  this in place of `widthPt` / `heightPt` for layout. `pct == 0` from
   *  the source is dropped at parse time (treated as "use extent"). */
  widthPct?: number | null;
  heightPct?: number | null;
  widthRelativeFrom?: string | null;
  heightRelativeFrom?: string | null;
  /** Parent wgp group dimensions (pt) — set only when this shape is a child
   *  of a `<wpg:wgp>`. Used by `resolveAnchor*` so align/pctPos resolve the
   *  GROUP's origin, then `anchor[XY]Pt` adds the within-group offset. */
  groupWidthPt?: number | null;
  groupHeightPt?: number | null;
  /** Draw behind text when true (wp:anchor behindDoc="1"). */
  behindDoc?: boolean;
  /** Document-order index within a group; lower values render first. */
  zOrder: number;
  /** Normalized [0,1] custom-geometry sub-paths. Empty when `presetGeometry`
   *  is set; the renderer chooses between buildCustomPath and buildShapePath. */
  subpaths: PathCmd[][];
  /** OOXML <a:prstGeom prst> name (e.g. "rect", "ellipse", "rtTriangle").
   *  When set the renderer calls core's buildShapePath with `adjValues`. */
  presetGeometry?: string | null;
  /** Up to four <a:gd name="adj{n}"> values from prstGeom/avLst (0–100000). */
  adjValues?: number[];
  fill: ShapeFill | null;
  stroke: string | null;
  strokeWidth?: number;
  /** `<a:ln><a:prstDash val>` — ECMA-376 §20.1.8.48. Absent = solid. */
  strokeDash?: string | null;
  /** `<a:ln><a:headEnd>` line-start decoration (ECMA-376 §20.1.8.3). */
  headEnd?: LineEnd | null;
  /** `<a:ln><a:tailEnd>` line-end decoration (ECMA-376 §20.1.8.3). */
  tailEnd?: LineEnd | null;
  rotation?: number;
  /** `<a:xfrm flipH>` (§20.1.7.6) — mirror about the vertical centre line. */
  flipH?: boolean;
  /** `<a:xfrm flipV>` (§20.1.7.6) — mirror about the horizontal centre line. */
  flipV?: boolean;
  wrapMode?: string | null;
  /** Padding top (pt). Anchor-only. Mirrors {@link ImageRun.distTop}; an anchored
   *  wrap-shape uses these to size its float-exclusion band (ECMA-376 §20.4.2.x). */
  distTop?: number;
  /** Padding bottom (pt). Anchor-only. */
  distBottom?: number;
  /** Padding left (pt). Anchor-only. */
  distLeft?: number;
  /** Padding right (pt). Anchor-only. */
  distRight?: number;
  /** wrapText attribute: "bothSides" | "left" | "right" | "largest". */
  wrapSide?: string | null;
  /** Text rendered INSIDE the shape's bounding box (`<wps:txbx><w:txbxContent>`). */
  textBlocks?: ShapeText[];
  /** "t" | "ctr" | "b" — vertical anchor for the shape's text body (`<wps:bodyPr @anchor>`). */
  textAnchor?: string | null;
  textInsetL?: number;  // pt
  textInsetT?: number;  // pt
  textInsetR?: number;  // pt
  textInsetB?: number;  // pt
}

/** DrawingML line-end (arrow head). ECMA-376 §20.1.8.3 CT_LineEndProperties.
 *  Maps 1:1 to core's `ArrowEnd`. */
export interface LineEnd {
  /** "triangle" | "stealth" | "diamond" | "oval" | "arrow" (never "none"). */
  type: string;
  /** Width step: "sm" | "med" | "lg". */
  w: string;
  /** Length step: "sm" | "med" | "lg". */
  len: string;
}

/** One formatting run (`<w:r>`) inside a shape-text paragraph. Mirrors the
 *  character-formatting fields of {@link ShapeText}; the renderer lays a
 *  paragraph's {@link ShapeText.runs} out as rich text so mixed bold/non-bold
 *  runs each keep their own font. */
export interface ShapeTextRun {
  text: string;
  fontSizePt: number;
  color?: string | null;
  /** ECMA-376 §17.3.2.26 ascii axis (`<w:rFonts w:ascii>`), resolved through
   *  docDefaults. Latin letters/digits in this run render with this family. */
  fontFamily?: string | null;
  /** ECMA-376 §17.3.2.26 eastAsia axis (`<w:rFonts w:eastAsia>`), resolved
   *  through docDefaults. CJK characters in this run render with this family;
   *  the renderer falls back to {@link ShapeTextRun.fontFamily} when absent. */
  fontFamilyEastAsia?: string | null;
  bold?: boolean;
  italic?: boolean;
}

export interface ShapeText {
  text: string;
  fontSizePt: number;
  color?: string | null;
  fontFamily?: string | null;
  bold?: boolean;
  italic?: boolean;
  /** Per-run formatting for this paragraph (one entry per `<w:r>` with text).
   *  When non-empty the renderer draws the block as rich text (each run's
   *  font); otherwise it uses the single block-level format fields above
   *  (image blocks / legacy single-format paragraphs). Absent for image-only
   *  paragraphs. */
  runs?: ShapeTextRun[];
  alignment: string;
  /** Zip path of an inline image inside this text-box paragraph
   *  (`<w:drawing><wp:inline><a:blip r:embed>`), e.g. `word/media/image1.emf`.
   *  Absent for a text-only paragraph. */
  imagePath?: string;
  /** MIME type of the blip at {@link ShapeText.imagePath}. */
  mimeType?: string;
  /** Zip path of the vector original (`asvg:svgBlip` extension), preferred over
   *  `imagePath` when present. */
  svgImagePath?: string;
  /** Inline image natural width in pt (from `<wp:extent cx>`). */
  imageWidthPt?: number;
  /** Inline image natural height in pt (from `<wp:extent cy>`). */
  imageHeightPt?: number;
}

export type ShapeFill =
  | { fillType: 'solid'; color: string }
  | { fillType: 'gradient'; stops: GradientStop[]; angle: number; gradType: string };

export interface GradientStop {
  /** 0.0–1.0 */
  position: number;
  /** hex 6-char */
  color: string;
}

export interface FieldRun {
  /** "page" | "numPages" | "other" */
  fieldType: string;
  instruction: string;
  fallbackText: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  fontSize: number;  // pt
  color: string | null;
  fontFamily: string | null;
  background: string | null;
  vertAlign: 'super' | 'sub' | null;
  allCaps?: boolean;
  smallCaps?: boolean;
  doubleStrikethrough?: boolean;
  highlight?: string | null;
}

export interface DocxTextRun {
  text: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  fontSize: number;  // pt
  color: string | null;
  fontFamily: string | null;
  /** ECMA-376 §17.3.2.26 eastAsia axis (`<w:rFonts w:eastAsia>`), resolved
   *  through the style chain + docDefaults. CJK code points in this run render
   *  with this family; {@link DocxTextRun.fontFamily} keeps the conflated single-
   *  font fallback (ascii → eastAsia) for paths that do not split per character.
   *  The renderer routes consecutive CJK code points to this axis (the same per-
   *  script rule {@link ShapeTextRun.fontFamilyEastAsia} uses), so a Gothic
   *  eastAsia title sits beside a serif ascii number with no name heuristics.
   *  Absent ⇒ the renderer falls back to {@link DocxTextRun.fontFamily}. */
  fontFamilyEastAsia?: string | null;
  isLink: boolean;
  background: string | null;
  /** ECMA-376 §17.3.2.6 — `<w:color w:val="auto"/>` was set on this run. When
   *  true and {@link DocxTextRun.color} is absent, the renderer resolves the
   *  glyph color from the effective background (an implementation-defined
   *  black/white contrast pick; ECMA-376 gives no normative algorithm) instead
   *  of the default text color. */
  colorAuto?: boolean | null;
  /** ECMA-376 §17.3.2.4 `<w:bdr>` — a run-level border (box) drawn around the
   *  run text. Absent when the run has no border. */
  border?: DocxRunBorder | null;
  vertAlign: 'super' | 'sub' | null;
  /** Target URL for hyperlinks (resolved from relationships.xml) */
  hyperlink: string | null;
  allCaps?: boolean;
  smallCaps?: boolean;
  doubleStrikethrough?: boolean;
  highlight?: string | null;
  /** ECMA-376 §17.3.3.25 ruby annotation (furigana). Renders above the
   *  base text in a smaller font; line height is expanded to fit it. */
  ruby?: RubyAnnotation;
  /** ECMA-376 §17.13.5 — set when this run sits inside `<w:ins>` or
   *  `<w:del>`. The renderer paints insertions with an author-coloured
   *  underline and deletions with an author-coloured strikethrough so
   *  tracked changes appear inline. */
  revision?: RunRevision;
  /** ECMA-376 §17.3.2.30 `<w:rtl>` — complex-script / right-to-left run.
   *  `true` = RTL, `false` = explicitly LTR, absent = unspecified. The renderer
   *  treats a `true` run as RTL for the UAX#9 pass (it forces complex-script
   *  shaping and marks the segment so `computeLineVisualOrder` reorders it), and
   *  draws the slice with `ctx.direction = 'rtl'` so Canvas mirrors the glyphs. */
  rtl?: boolean;
  /** ECMA-376 §17.3.2.7 `<w:cs/>` — complex-script run toggle: cs formatting
   *  applies to ALL characters of the run (§17.3.2.26). Distinct from
   *  `rFonts@cs` (`fontFamilyCs`), which is only a font slot. */
  cs?: boolean;
  /** ECMA-376 §17.3.2.26 `<w:rFonts w:cs>` — complex-script typeface
   *  (theme references resolved to a literal family). */
  fontFamilyCs?: string;
  /** ECMA-376 §17.3.2.39 `<w:szCs>` — complex-script font size in pt
   *  (same units as `fontSize`). */
  fontSizeCs?: number;
  /** ECMA-376 §17.3.2.3 `<w:bCs>` — complex-script bold toggle. */
  boldCs?: boolean;
  /** ECMA-376 §17.3.2.17 `<w:iCs>` — complex-script italic toggle. */
  italicCs?: boolean;
  /** ECMA-376 §17.3.2.20 `<w:lang w:bidi>` — complex-script (RTL) language tag,
   *  lower-cased (e.g. "ar-sa", "ae-ar"). Drives Word's AN digit ordering. */
  langBidi?: string;
  /** ECMA-376 §17.11.6/.7/.16/.17 — set when this run is a footnote/endnote
   *  reference marker (`<w:footnoteReference>` in the body, `<w:footnoteRef>` at
   *  the start of the note's content, and the endnote equivalents). `text` holds
   *  the raw `@w:id`; the renderer overrides the displayed glyph with the note's
   *  sequential number. */
  noteRef?: NoteRef;
}

/** A footnote / endnote reference marker (ECMA-376 §17.11). */
export interface NoteRef {
  /** "footnote" | "endnote" */
  kind: 'footnote' | 'endnote' | string;
  /** `@w:id` linking the marker to its note. Empty for the in-note
   *  `<w:footnoteRef/>` placeholder (the renderer uses the enclosing note). */
  id: string;
}

export interface RunRevision {
  /** "insertion" or "deletion" */
  kind: 'insertion' | 'deletion' | string;
  /** `<w:ins w:author>` / `<w:del w:author>`. Used to colour the markup. */
  author?: string;
  /** ISO-8601 timestamp. */
  date?: string;
}

export interface RubyAnnotation {
  text: string;
  /** Annotation font size in pt. Word stores this as half-points in `<w:hps>`. */
  fontSizePt: number;
}

export interface ImageRun {
  /**
   * Embedded zip path of the raster blip (e.g. `word/media/image1.png`) — the
   * raster fallback (PNG/JPEG), or the SVG part itself when no raster blip is
   * embedded. The renderer fetches the bytes lazily by path (see {@link
   * DocxDocument.getImage}) instead of inlining base64.
   */
  imagePath: string;
  /** MIME type of the blip at {@link ImageRun.imagePath} (e.g. `image/png`, or
   *  `image/svg+xml` for an svg-only picture). */
  mimeType: string;
  /**
   * Vector original from the Microsoft `asvg:svgBlip` extension (MS-ODRAWXML) —
   * the zip path of the `.svg` part. When present the renderer prefers it over
   * {@link ImageRun.imagePath} (the raster fallback). Absent for a plain raster
   * image. Its MIME is always `image/svg+xml` and is owned by the SVG decoder.
   */
  svgImagePath?: string;
  /**
   * ECMA-376 §20.1.8.55 `<a:srcRect>` — the source-rectangle crop applied to
   * the decoded bitmap before it is drawn into the display box. The four values
   * are inset FRACTIONS 0..1 of the source bitmap measured inward from each
   * edge (`l`/`t` from left/top, `r`/`b` from right/bottom); the visible source
   * region is `[l, t, 1−r, 1−b]`. The parser converts the raw ST_Percentage
   * (1000ths of a percent) to fractions, so the renderer crops in bitmap pixels
   * (`sx = l*w`, `sy = t*h`, `sw = (1−l−r)*w`, `sh = (1−t−b)*h`) without unit
   * knowledge. Absent / null when there is no crop (the full bitmap is drawn).
   */
  srcRect?: { l: number; t: number; r: number; b: number } | null;
  widthPt: number;
  heightPt: number;
  /** true = wp:anchor (absolute positioned), false/undefined = wp:inline (flows with text) */
  anchor?: boolean;
  /** X offset in pt (anchor only) */
  anchorXPt?: number;
  /** Y offset in pt (anchor only) */
  anchorYPt?: number;
  /**
   * If true, anchorXPt is relative to the left margin — add section.marginLeft to get page X.
   * If false/absent, anchorXPt is already page-absolute.
   */
  anchorXFromMargin?: boolean;
  /**
   * If true, anchorYPt is relative to the paragraph's top Y in the renderer.
   * If false/absent, anchorYPt is already page-absolute.
   */
  anchorYFromPara?: boolean;
  /**
   * When set, the renderer replaces all pixels of this hex color (e.g. "FFFFFF") with full
   * transparency. Implements a:clrChange (make-background-transparent).
   */
  colorReplaceFrom?: string;
  /**
   * Wrap mode for anchor images:
   *   "square" | "topAndBottom" | "none" | "tight" | "through"
   * Inline images and undetermined cases leave this undefined.
   * MVP renders "tight" and "through" as "square".
   */
  wrapMode?: string;
  /** Padding top (pt). Anchor-only. */
  distTop?: number;
  /** Padding bottom (pt). Anchor-only. */
  distBottom?: number;
  /** Padding left (pt). Anchor-only. */
  distLeft?: number;
  /** Padding right (pt). Anchor-only. */
  distRight?: number;
  /** wrapText attribute: "bothSides" | "left" | "right" | "largest". */
  wrapSide?: string;
  /**
   * ECMA-376 §20.4.2.3 `wp:anchor/@allowOverlap` — whether this floating object
   * may overlap other floats. Spec default is true (the attribute is optional);
   * absent/undefined is treated as true. `false` mandates the renderer
   * reposition the object to prevent any overlap.
   */
  allowOverlap?: boolean;
  /** ECMA-376 §20.4.3.1 wp:align horizontal: "left" | "center" | "right" |
   *  "inside" | "outside". When set the renderer aligns the image inside the
   *  container indicated by `anchorXFromMargin` and ignores `anchorXPt`.
   *  Mirrors {@link ShapeRun.anchorXAlign}. Absent for inline images and
   *  offset-based anchors. */
  anchorXAlign?: string | null;
  /** Vertical equivalent of anchorXAlign: "top" | "center" | "bottom". */
  anchorYAlign?: string | null;
  /**
   * ECMA-376 §20.4.3.2 `<wp:positionH/@relativeFrom>` / §20.4.3.5
   * `<wp:positionV/@relativeFrom>` — names the container the offset / align /
   * pctPos is measured against. Raw spec string: `"page"`, `"margin"`,
   * `"paragraph"`, `"line"`, `"leftMargin"`, `"rightMargin"`, `"topMargin"`,
   * `"bottomMargin"`, `"insideMargin"`, `"outsideMargin"`, `"column"`,
   * `"character"`. Mirrors {@link ShapeRun.anchorXRelativeFrom} /
   * {@link ShapeRun.anchorYRelativeFrom}. When present, supersedes the legacy
   * coarse boolean hints (`anchorXFromMargin` / `anchorYFromPara`) for the
   * align and pctPos paths so e.g. `relativeFrom="margin"` + `align="top"`
   * pins the image to the top content margin rather than the page top. Absent
   * for inline images and for anchors that omitted `<wp:positionH/V>`.
   */
  anchorXRelativeFrom?: string | null;
  anchorYRelativeFrom?: string | null;
}

// ===== Table =====

/**
 * ECMA-376 §17.4.57 `<w:tblpPr>` — floating-table positioning. Present in
 * `<w:tblPr>` ⇒ the table FLOATS (out of the main text flow, absolutely
 * positioned by its top-left corner). All fields are optional in the source.
 */
export interface TblpPr {
  /** §17.4.57 minimum distance to wrapping text (dist padding), pt. Default 0. */
  leftFromText: number;
  rightFromText: number;
  topFromText: number;
  bottomFromText: number;
  /** §17.4.57 ST_HAnchor {text,margin,page}. Default 'page'. */
  horzAnchor: 'text' | 'margin' | 'page' | string;
  /** True iff the source `<w:tblpPr>` carried ANY horizontal positioning hint
   *  (horzAnchor, tblpX, or tblpXSpec). When false, no horizontal position was
   *  given: ECMA-376's literal default is the page edge, but Word places such a
   *  table at the anchor paragraph's text/column left. computeFloatTableBox uses
   *  this flag to apply that Word-runtime placement. */
  horzSpecified: boolean;
  /** §17.4.57 ST_VAnchor {text,margin,page}. Default 'page'. */
  vertAnchor: 'text' | 'margin' | 'page' | string;
  /** §17.4.57 absolute signed offset from the horz/vert anchor edge, pt.
   *  Default 0. Ignored when the matching `*Spec` is present. */
  tblpX: number;
  tblpY: number;
  /** §17.4.57 ST_XAlign {left,center,right,inside,outside}. Supersedes tblpX. */
  tblpXSpec?: 'left' | 'center' | 'right' | 'inside' | 'outside' | string;
  /** §17.4.57 ST_YAlign {inline,top,center,bottom,inside,outside}. Supersedes
   *  tblpY, UNLESS vertAnchor='text' (relative vertical positioning is not
   *  allowed there ⇒ tblpYSpec is ignored, fall back to tblpY). */
  tblpYSpec?: 'inline' | 'top' | 'center' | 'bottom' | 'inside' | 'outside' | string;
}

export interface DocTable {
  colWidths: number[];  // pt
  rows: DocTableRow[];
  borders: TableBorders;
  cellMarginTop: number;
  cellMarginBottom: number;
  cellMarginLeft: number;
  cellMarginRight: number;
  /** table horizontal alignment on the page: 'left' | 'center' | 'right'. */
  jc: string;
  /** ECMA-376 §17.4.52 `<w:tblLayout w:type>` — 'fixed' | 'autofit'. Absent
   *  (undefined) ⇒ spec default 'autofit'. Both paths size columns from the
   *  tblGrid (§17.4.48) scaled to fit: 'fixed' uses the grid verbatim; 'autofit'
   *  additionally lets content min-width grow a column. Per-cell `widthPt`/
   *  `widthPct` (`<w:tcW>`) is NOT re-applied — Word bakes the resolved widths
   *  into the saved grid (see resolveColumnWidths). Only a degenerate all-zero
   *  grid falls back to tcW-preference sizing. */
  layout?: string;
  /** ECMA-376 §17.4.63 `<w:tblW>` preferred table width (type="dxa"), pt. */
  widthPt?: number;
  /** `<w:tblW>` type="pct": 50ths of a percent of available content width. */
  widthPct?: number;
  /**
   * ECMA-376 §17.4.1 `<w:bidiVisual>` — render columns in right-to-left
   * (visual) order. `true` = RTL columns, `false` = explicitly LTR, absent =
   * unspecified. When `true` the renderer mirrors the grid so logical column 0
   * is placed rightmost, and flips per-cell left/right borders accordingly.
   */
  bidiVisual?: boolean;
  /** ECMA-376 §17.4.57 `<w:tblpPr>` — when present the table is FLOATING
   *  (absolutely positioned, out of the main text flow). Absent ⇒ block table. */
  tblpPr?: TblpPr;
  /** ECMA-376 §17.4.56 `<w:tblOverlap w:val>` — 'never' | 'overlap'. 'never' ⇒
   *  the floating table must be repositioned to avoid overlapping other floats.
   *  Default 'overlap' (omitted ⇒ overlap allowed). Ignored when not floating. */
  overlap?: string;
}

export interface TableBorders {
  top: BorderSpec | null;
  bottom: BorderSpec | null;
  left: BorderSpec | null;
  right: BorderSpec | null;
  insideH: BorderSpec | null;
  insideV: BorderSpec | null;
}

export interface BorderSpec {
  width: number;   // pt
  color: string | null;
  style: string;
}

export interface DocTableRow {
  cells: DocTableCell[];
  rowHeight: number | null;  // pt
  /** ECMA-376 §17.4.80 hRule. "auto" (default) = informational; "atLeast" =
   *  lower bound; "exact" = fixed clip. */
  rowHeightRule: 'auto' | 'atLeast' | 'exact' | string;
  isHeader: boolean;
}

/** ECMA-376 §17.4.7: a table cell may contain paragraphs AND nested tables. */
export type CellElement =
  | { type: 'paragraph' } & DocParagraph
  | { type: 'table' } & DocTable;

export interface DocTableCell {
  content: CellElement[];
  colSpan: number;
  vMerge: boolean | null;  // null=no merge, true=start, false=continuation
  borders: CellBorders;
  background: string | null;
  vAlign: 'top' | 'center' | 'bottom';
  /** ECMA-376 §17.4.71 `<w:tcW>` preferred cell width (type="dxa"), pt. A
   *  PREFERRED width only: autofit column sizing is driven by the tblGrid
   *  (§17.4.48), not by re-applying this (Word bakes the resolved widths into
   *  the saved grid — see resolveColumnWidths). Consulted only for the
   *  degenerate all-zero-grid fallback. */
  widthPt: number | null;
  /** `<w:tcW>` type="pct": 50ths of a percent of available content width.
   *  Resolved against the available width at render time. Preferred width only
   *  (see `widthPt`). */
  widthPct?: number;
  /** Per-cell margins (pt) from `<w:tcPr><w:tcMar>` (ECMA-376 §17.4.42). Each
   *  edge overrides the table-level `cellMargin*` default when set; null/absent
   *  = inherit the table default. */
  marginTop?: number | null;
  marginBottom?: number | null;
  marginLeft?: number | null;
  marginRight?: number | null;
}

export interface CellBorders {
  top: BorderSpec | null;
  bottom: BorderSpec | null;
  left: BorderSpec | null;
  right: BorderSpec | null;
  /** ECMA-376 §17.4.34 (tcBorders w:insideH/w:insideV): the interior
   *  horizontal/vertical border this cell contributes. Folded from the cell's
   *  inline tcBorders OVER the resolved conditional table-style borders (§17.7.6)
   *  at parse time. `null` = unset (the renderer falls back to the table-level
   *  insideH/insideV); a spec with style "nil"/"none" = an explicit "no interior
   *  border" (e.g. banded data rows in Medium List 2 / Medium Shading 2). */
  insideH: BorderSpec | null;
  insideV: BorderSpec | null;
}

// ===== Worker message protocol =====

export type WorkerRequest =
  | { type: 'init'; wasmUrl: string }
  | { type: 'parse'; id: number; data: ArrayBuffer; maxZipEntryBytes?: number }
  | { type: 'extractImage'; id: number; path: string };

export type WorkerResponse =
  | { type: 'parsed'; id: number; document: DocxDocumentModel }
  | { type: 'imageExtracted'; id: number; bytes: ArrayBuffer }
  | { type: 'error'; id: number; message: string };

// ===== Public API types =====

export interface RenderPageOptions {
  /** Canvas CSS width in px; height is auto-computed from page aspect ratio */
  width?: number;
  dpr?: number;
  defaultTextColor?: string;
  /** Called for each rendered text segment. Used to build a transparent text selection overlay. */
  onTextRun?: (run: { text: string; x: number; y: number; w: number; h: number; fontSize: number; font: string }) => void;
  /** Default `true`. When false, ECMA-376 §17.13.5 track-changes runs render
   *  in their normal style (no author colour, no underline / strikethrough)
   *  — equivalent to Word's "Final / No Markup" view. */
  showTrackChanges?: boolean;
}
