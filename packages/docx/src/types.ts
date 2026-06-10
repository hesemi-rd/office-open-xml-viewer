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
  text: string;
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
  /** ECMA-376 §17.6.5 w:docGrid/@w:type — "default" | "lines" | "linesAndChars" | "snapToChars". */
  docGridType?: string | null;
  /** ECMA-376 §17.6.5 w:docGrid/@w:linePitch in pt. When docGridType is "lines" or
   *  "linesAndChars", auto line spacing multiplies against this pitch instead of
   *  the font's natural line height. */
  docGridLinePitch?: number | null;
}

export type BodyElement =
  | { type: 'paragraph' } & DocParagraph
  | { type: 'table' } & DocTable
  | { type: 'pageBreak'; parity?: 'odd' | 'even' };

/** A BodyElement annotated with a line range to render. Set when the
 *  paginator splits a paragraph that doesn't fit on a single page —
 *  `lineSlice` constrains which laid-out line indices the renderer paints,
 *  and the renderer adjusts the starting Y so the slice's first line begins
 *  at the page's content top. */
export type PaginatedBodyElement = BodyElement & {
  lineSlice?: { start: number; end: number };
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
   * `false` = explicitly LTR, absent = unspecified (inherit). Phase 0 of RTL
   * support: recorded only; alignment/column-order resolution is deferred.
   */
  bidi?: boolean;
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
}

export type DocRun =
  | { type: 'text' } & DocxTextRun
  | { type: 'image' } & ImageRun
  | { type: 'break'; breakType: 'line' | 'page' | 'column' }
  | { type: 'field' } & FieldRun
  | { type: 'shape' } & ShapeRun
  | { type: 'math'; nodes: MathNode[]; display: boolean; fontSize: number };

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
  rotation?: number;
  wrapMode?: string | null;
  /** Text rendered INSIDE the shape's bounding box (`<wps:txbx><w:txbxContent>`). */
  textBlocks?: ShapeText[];
  /** "t" | "ctr" | "b" — vertical anchor for the shape's text body (`<wps:bodyPr @anchor>`). */
  textAnchor?: string | null;
  textInsetL?: number;  // pt
  textInsetT?: number;  // pt
  textInsetR?: number;  // pt
  textInsetB?: number;  // pt
}

export interface ShapeText {
  text: string;
  fontSizePt: number;
  color?: string | null;
  fontFamily?: string | null;
  bold?: boolean;
  italic?: boolean;
  alignment: string;
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
  isLink: boolean;
  background: string | null;
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
   *  `true` = RTL, `false` = explicitly LTR, absent = unspecified.
   *  Phase 0: recorded only; glyph-order resolution is deferred. */
  rtl?: boolean;
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
  dataUrl: string;
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
}

// ===== Table =====

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
   *  (undefined) ⇒ spec default 'autofit': columns are sized by the per-column
   *  max preferred width (cell `widthPt`), tblGrid only as fallback. 'fixed'
   *  uses tblGrid widths verbatim (scaled to fit). */
  layout?: string;
  /** ECMA-376 §17.4.63 `<w:tblW>` preferred table width (type="dxa"), pt. */
  widthPt?: number;
  /** `<w:tblW>` type="pct": 50ths of a percent of available content width. */
  widthPct?: number;
  /**
   * ECMA-376 §17.4.1 `<w:bidiVisual>` — render columns in right-to-left
   * (visual) order. `true` = RTL columns, `false` = explicitly LTR, absent =
   * unspecified. Phase 0 of RTL support: recorded only; column-order
   * resolution is deferred.
   */
  bidiVisual?: boolean;
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
  /** ECMA-376 §17.4.71 `<w:tcW>` preferred cell width (type="dxa"), pt. Drives
   *  autofit column sizing: each grid column's width is the max `widthPt` over
   *  the cells anchored in it. */
  widthPt: number | null;
  /** `<w:tcW>` type="pct": 50ths of a percent of available content width.
   *  Resolved against the available width at render time. */
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
}

// ===== Worker message protocol =====

export type WorkerRequest =
  | { type: 'init'; wasmUrl: string }
  | { type: 'parse'; id: number; data: ArrayBuffer; maxZipEntryBytes?: number };

export type WorkerResponse =
  | { type: 'parsed'; id: number; document: DocxDocumentModel }
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
