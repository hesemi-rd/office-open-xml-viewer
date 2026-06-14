import type {
  DocxDocumentModel, BodyElement, PaginatedBodyElement, DocParagraph, DocTable, DocTableRow, DocTableCell, CellElement,
  DocRun, DocxTextRun, ImageRun, ShapeRun, FieldRun, HeaderFooter, LineSpacing, BorderSpec, TableBorders, CellBorders,
  TabStop, ParagraphBorders, ParaBorderEdge, SectionProps, DocNote,
} from './types';
import {
  buildCustomPath,
  buildShapePath,
  hexToRgba,
  resolveFill,
  mathToMathML,
  recolorSvg,
  PT_TO_PX,
  resolveBaseDirection,
  isHTMLCanvas,
  defaultDpr,
  classifyCjkFont,
  cjkFallbackChain,
  NON_CJK_SANS_FALLBACKS,
  NON_CJK_SERIF_FALLBACKS,
} from '@silurus/ooxml-core';
import type { MathNode, MathRenderer } from '@silurus/ooxml-core';
import { intendedSingleLinePx, correctLineMetrics } from './font-metrics.js';
import {
  segmentsHaveRtl,
  computeLineVisualOrder,
  resolveAlignEdge,
  type AlignEdge,
  type LineVisualOrder,
} from './bidi-line.js';

const HIGHLIGHT_COLORS: Record<string, string> = {
  yellow: '#FFFF00', cyan: '#00FFFF', green: '#00FF00', magenta: '#FF00FF',
  blue: '#0000FF', red: '#FF0000', darkBlue: '#000080', darkCyan: '#008080',
  darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000',
  darkYellow: '#808000', darkGray: '#808080', lightGray: '#C0C0C0',
  black: '#000000', white: '#FFFFFF',
};

// ── Math (OMML) rendering via MathJax ───────────────────────────────────────
// Each equation is converted OMML AST -> MathML -> MathJax SVG, then rasterized to
// an <img> once (async, before pagination). Layout reads cached em-extents
// synchronously; drawing blits the image. Skipped entirely for math-free documents.
interface MathRender {
  img: CanvasImageSource;
  /** baseline-relative extents in em (1em = the equation's font size). */
  widthEm: number;
  ascentEm: number;
  descentEm: number;
}
// Keyed by the run's MathNode[] reference, which is stable from parse through render.
const mathRenders = new WeakMap<MathNode[], MathRender>();

/** True if any run in the body (incl. tables) is an OMML equation. */
export function documentHasMath(body: BodyElement[]): boolean {
  return collectMathRuns(body).length > 0;
}

function collectMathRuns(body: BodyElement[]): { nodes: MathNode[]; display: boolean }[] {
  const found: { nodes: MathNode[]; display: boolean }[] = [];
  const fromRuns = (runs: DocRun[]) => {
    for (const r of runs) {
      if (r.type === 'math') found.push({ nodes: r.nodes, display: r.display });
    }
  };
  const walk = (el: BodyElement) => {
    if ('runs' in el) fromRuns((el as DocParagraph).runs);
    if ('rows' in el) {
      for (const row of (el as DocTable).rows) {
        for (const cell of row.cells) {
          for (const child of cell.content) walk(child as BodyElement);
        }
      }
    }
  };
  body.forEach(walk);
  return found;
}

/** Rasterize an SVG string to an <img> (browser). Resolves once decoded. */
function svgToImage(svg: string): Promise<HTMLImageElement> {
  const url = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  const img = new Image();
  return new Promise((resolve, reject) => {
    img.onload = () => resolve(img);
    img.onerror = reject;
    img.src = url;
  });
}

/**
 * Convert + rasterize every equation in the document. Must complete before
 * pagination/render (which read extents synchronously). Idempotent per equation.
 */
export async function prepareMathRuns(body: BodyElement[], math: MathRenderer): Promise<void> {
  const runs = collectMathRuns(body);
  if (runs.length === 0) return;
  await math.loadMathJax();
  for (const r of runs) {
    if (mathRenders.has(r.nodes)) continue;
    try {
      const out = await math.mathMLToSvg(mathToMathML(r.nodes, r.display));
      const img = await svgToImage(recolorSvg(out.svg, '#000000'));
      mathRenders.set(r.nodes, {
        img,
        widthEm: out.widthEm,
        ascentEm: out.ascentEm,
        descentEm: out.descentEm,
      });
    } catch {
      // Conversion failure: leave the equation unrendered (zero-size) rather than throw.
    }
  }
}

/** Anchor image float that affects text wrap on the current page. */
interface FloatRect {
  mode: 'square' | 'topAndBottom';
  /** Hex key of the image bitmap (used to defer drawing until final Y is known). */
  imageKey: string;
  /** Absolute canvas X of the image box (without dist padding). */
  imageX: number;
  imageY: number;
  imageW: number;
  imageH: number;
  /** Padded exclusion rectangle for text wrap. */
  xLeft: number;
  xRight: number;
  yTop: number;
  yBottom: number;
  /** wrapText: "bothSides" | "left" | "right" | "largest" (only square uses this). */
  side: string;
  /** true once the image itself has been drawn (drawn after its paragraph lays out). */
  drawn: boolean;
}

interface RenderState {
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
  scale: number;    // px per pt
  contentX: number; // left of content area (px)
  contentW: number; // width of content area (px)
  y: number;        // current Y cursor (px)
  pageH: number;    // full page height (px)
  defaultColor: string;
  /** 0-based page index currently being rendered */
  pageIndex: number;
  /** total page count in the document */
  totalPages: number;
  /** preloaded image bitmaps keyed by dataUrl */
  images: Map<string, ImageBitmap>;
  /** when true, layout is performed but nothing is drawn (used for header/footer height measurement) */
  dryRun: boolean;
  /** section left margin in pt — used to convert margin-relative anchor X to page-absolute */
  marginLeft: number;
  /** section right/top/bottom margins in pt — used by anchor positioning to
   *  resolve `<wp:positionH/V relativeFrom="margin">` and the
   *  `*Margin` family containers. */
  marginRight: number;
  marginTop: number;
  marginBottom: number;
  /** Section page width in pt. */
  pageWidth: number;
  /** Active anchor-image floats that constrain text layout on the current page. */
  floats: FloatRect[];
  /** ECMA-376 §17.6.5 docGrid (type + pitch), applied to auto line spacing. */
  docGrid: DocGridCtx;
  /** ECMA-376 §17.8.3.10 — font→family map from word/fontTable.xml. Used by
   *  resolveFontFamily as the authoritative source of serif/sans-serif classification. */
  fontFamilyClasses: Record<string, string>;
  /** ECMA-376 §17.15.1.58–.60 — resolved Japanese line-breaking rules
   *  (kinsoku enabled flag + line-start/line-end forbidden character sets).
   *  Default is the application's Japanese kinsoku table with kinsoku ON. */
  kinsoku: KinsokuRules;
  /** ECMA-376 §22.1.2.30 `m:mathPr/m:defJc` — document-wide default math
   *  justification (ST_Jc math). `undefined` ⇒ spec default `centerGroup`.
   *  Threaded from `doc.settings.mathDefJc` like `kinsoku`; consumed by the
   *  per-line alignment step for single display-math lines. */
  mathDefJc?: string;
  /** Callback for building a transparent text selection overlay. */
  onTextRun?: (run: DocxTextRunInfo) => void;
  /** When false, runs tagged with a `revision` render without the
   *  track-changes overlay (no author colour, no underline/strikethrough). */
  showTrackChanges: boolean;
  /** ECMA-376 §17.11 — footnote/endnote reference markers (`noteRef` runs)
   *  display the note's 1-based sequential number, not the raw `@w:id`. Keyed by
   *  `"footnote:<id>"` / `"endnote:<id>"`. The in-note `<w:footnoteRef>`
   *  placeholder (empty id) is substituted with the number provided via
   *  {@link currentNoteNumber} while drawing that note's content. */
  noteNumbers?: Map<string, number>;
  /** Set while laying out a footnote/endnote's own content, so the leading
   *  `<w:footnoteRef>` placeholder (which carries no id) renders the note's
   *  number. Undefined for body text. */
  currentNoteNumber?: number;
}

/** Information about a rendered text segment for building a transparent selection overlay. */
export interface DocxTextRunInfo {
  text: string;
  /** Left edge in canvas CSS px. */
  x: number;
  /** Top of line box in canvas CSS px. */
  y: number;
  /** Measured text width in CSS px. */
  w: number;
  /** Line height in CSS px. */
  h: number;
  /** Font size in CSS px. */
  fontSize: number;
  /** CSS `font` shorthand used for canvas drawing (e.g. `"bold 16px Arial"`). */
  font: string;
}

export interface RenderDocumentOptions {
  width?: number;
  dpr?: number;
  defaultTextColor?: string;
  /** total pages in the document (used to resolve NUMPAGES fields) */
  totalPages?: number;
  /** Pre-computed page splits (from computePages). When provided, skips internal pagination. */
  prebuiltPages?: PaginatedBodyElement[][];
  /** Called for each rendered text segment. Used to build a transparent text selection overlay. */
  onTextRun?: (run: DocxTextRunInfo) => void;
  /** Default `true`. When false, runs tagged with a `revision` (insertion or
   *  deletion from `<w:ins>` / `<w:del>`) render in their normal colour with
   *  no underline / strikethrough overlay — useful for a "final / no markup"
   *  view of a tracked document. */
  showTrackChanges?: boolean;
}

// ===== Image preloading =====

interface ImagePair {
  url: string;
  colorReplaceFrom?: string;
}

/** Returns a stable map key for a (url, colorReplaceFrom) pair. */
function imageKey(url: string, colorReplaceFrom?: string): string {
  return colorReplaceFrom ? `${url}|clr:${colorReplaceFrom}` : url;
}

/** Picks a stable colour for a track-changes author. Mirrors Word's behaviour
 *  of cycling through a fixed palette (Word uses 8 hues then alternates).
 *  An empty / missing author maps to the first colour. */
const TRACK_CHANGE_AUTHOR_PALETTE = [
  '#C00000', // red
  '#0070C0', // blue
  '#00B050', // green
  '#7030A0', // purple
  '#E97132', // orange
  '#196B24', // dark green
  '#9E480E', // brown
  '#525252', // grey
];
function authorColor(author?: string): string {
  if (!author) return TRACK_CHANGE_AUTHOR_PALETTE[0];
  // Simple FNV-1a style hash so the same author always gets the same colour.
  let h = 0x811c9dc5;
  for (let i = 0; i < author.length; i++) {
    h ^= author.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return TRACK_CHANGE_AUTHOR_PALETTE[Math.abs(h) % TRACK_CHANGE_AUTHOR_PALETTE.length];
}

function collectImagePairs(doc: DocxDocumentModel): ImagePair[] {
  const seen = new Map<string, ImagePair>();
  const walk = (runs: DocRun[]) => {
    for (const run of runs) {
      if (run.type === 'image') {
        const img = run as unknown as ImageRun;
        const key = imageKey(img.dataUrl, img.colorReplaceFrom);
        if (!seen.has(key)) seen.set(key, { url: img.dataUrl, colorReplaceFrom: img.colorReplaceFrom });
      }
    }
  };
  const walkTable = (tbl: DocTable) => {
    for (const row of tbl.rows)
      for (const cell of row.cells)
        for (const ce of cell.content) {
          if (ce.type === 'paragraph') walk((ce as unknown as DocParagraph).runs);
          else if (ce.type === 'table') walkTable(ce as unknown as DocTable);
        }
  };
  const walkBody = (body: BodyElement[]) => {
    for (const el of body) {
      if (el.type === 'paragraph') walk((el as unknown as DocParagraph).runs);
      if (el.type === 'table') walkTable(el as unknown as DocTable);
    }
  };
  walkBody(doc.body);
  if (doc.headers.default) walkBody(doc.headers.default.body);
  if (doc.headers.first)   walkBody(doc.headers.first.body);
  if (doc.headers.even)    walkBody(doc.headers.even.body);
  if (doc.footers.default) walkBody(doc.footers.default.body);
  if (doc.footers.first)   walkBody(doc.footers.first.body);
  if (doc.footers.even)    walkBody(doc.footers.even.body);
  return [...seen.values()];
}

/**
 * Apply a:clrChange color replacement: turn every pixel whose (R,G,B) matches colorHex into
 * fully transparent (alpha=0). Returns a new ImageBitmap with the modified pixels.
 */
async function applyColorReplacement(bmp: ImageBitmap, colorHex: string): Promise<ImageBitmap> {
  const r = parseInt(colorHex.slice(0, 2), 16);
  const g = parseInt(colorHex.slice(2, 4), 16);
  const b = parseInt(colorHex.slice(4, 6), 16);

  const offscreen = new OffscreenCanvas(bmp.width, bmp.height);
  const ctx2 = offscreen.getContext('2d')!;
  ctx2.drawImage(bmp, 0, 0);

  const imgData = ctx2.getImageData(0, 0, bmp.width, bmp.height);
  const d = imgData.data;

  for (let i = 0; i < d.length; i += 4) {
    if (d[i] === r && d[i + 1] === g && d[i + 2] === b) {
      d[i + 3] = 0; // make transparent
    }
  }

  ctx2.putImageData(imgData, 0, 0);
  return createImageBitmap(offscreen);
}

async function preloadImages(doc: DocxDocumentModel): Promise<Map<string, ImageBitmap>> {
  const pairs = collectImagePairs(doc);
  const entries = await Promise.all(
    pairs.map(async (pair): Promise<[string, ImageBitmap] | null> => {
      try {
        const res = await fetch(pair.url);
        const blob = await res.blob();
        let bmp = await createImageBitmap(blob);
        if (pair.colorReplaceFrom) {
          bmp = await applyColorReplacement(bmp, pair.colorReplaceFrom);
        }
        return [imageKey(pair.url, pair.colorReplaceFrom), bmp];
      } catch {
        return null;
      }
    }),
  );
  return new Map(entries.filter((e): e is [string, ImageBitmap] => e !== null));
}

// ===== Main entry =====

export async function renderDocumentToCanvas(
  doc: DocxDocumentModel,
  canvas: HTMLCanvasElement | OffscreenCanvas,
  pageIndex: number,
  opts: RenderDocumentOptions = {},
): Promise<void> {
  const sec = doc.section;
  const dpr = opts.dpr ?? defaultDpr();
  const cssWidth = opts.width ?? sec.pageWidth * PT_TO_PX;
  const scale = cssWidth / sec.pageWidth;  // px per pt
  const cssHeight = sec.pageHeight * scale;

  canvas.width = Math.round(cssWidth * dpr);
  canvas.height = Math.round(cssHeight * dpr);

  if (isHTMLCanvas(canvas)) {
    canvas.style.width = `${cssWidth}px`;
    canvas.style.height = `${cssHeight}px`;
    if (!canvas.style.display) canvas.style.display = 'block';
  }

  const ctx = canvas.getContext('2d') as CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
  ctx.scale(dpr, dpr);

  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, cssWidth, cssHeight);

  const kinsoku = resolveKinsokuRules(doc.settings);
  const pages = opts.prebuiltPages ?? computePages(doc.body, sec, ctx, doc.fontFamilyClasses ?? {}, kinsoku, doc.footnotes ?? []);
  const totalPages = Math.max(opts.totalPages ?? pages.length, pages.length);
  const elements = pages[pageIndex] ?? pages[0] ?? [];

  const images = await preloadImages(doc);

  // ECMA-376 §17.11: map each note id to its 1-based display number so the
  // reference markers (and the in-note footnoteRef placeholder) show the
  // sequential number, not the raw @w:id.
  const footnoteNums = buildNoteNumberMap(doc.footnotes);
  const endnoteNums = buildNoteNumberMap(doc.endnotes);
  const noteNumbers = new Map<string, number>();
  for (const [id, n] of footnoteNums) noteNumbers.set(`footnote:${id}`, n);
  for (const [id, n] of endnoteNums) noteNumbers.set(`endnote:${id}`, n);

  const baseState: RenderState = {
    ctx,
    scale,
    contentX: sec.marginLeft * scale,
    contentW: (sec.pageWidth - sec.marginLeft - sec.marginRight) * scale,
    y: sec.marginTop * scale,
    pageH: cssHeight,
    defaultColor: opts.defaultTextColor ?? '#000000',
    pageIndex,
    totalPages,
    images,
    dryRun: false,
    marginLeft: sec.marginLeft,
    marginRight: sec.marginRight,
    marginTop: sec.marginTop,
    marginBottom: sec.marginBottom,
    pageWidth: sec.pageWidth,
    floats: [],
    docGrid: { type: sec.docGridType ?? null, linePitchPt: sec.docGridLinePitch ?? null },
    fontFamilyClasses: doc.fontFamilyClasses ?? {},
    kinsoku,
    mathDefJc: doc.settings?.mathDefJc,
    onTextRun: opts.onTextRun,
    showTrackChanges: opts.showTrackChanges ?? true,
    noteNumbers,
  };

  // Header: top of page, starting at headerDistance
  const header = pickHeaderFooter(doc.headers, pageIndex, totalPages, doc.section.titlePage, doc.section.evenAndOddHeaders);
  if (header) {
    renderHeaderFooter(header, sec.headerDistance * scale, baseState);
  }

  // Footer: anchored from bottom, rising by its measured height
  const footer = pickHeaderFooter(doc.footers, pageIndex, totalPages, doc.section.titlePage, doc.section.evenAndOddHeaders);
  if (footer) {
    const footerHeight = measureHeaderFooterHeight(footer, baseState);
    const footerTopY = cssHeight - sec.footerDistance * scale - footerHeight;
    renderHeaderFooter(footer, footerTopY, baseState);
  }

  // Body
  const bodyState: RenderState = { ...baseState, y: sec.marginTop * scale };
  renderBodyElements(elements, bodyState);

  // Footnotes referenced on this page (ECMA-376 §17.11): drawn at the bottom of
  // the text column, above a short separator rule. The page area was already
  // reserved during pagination so the body stops short of them.
  drawPageFootnotes(elements, doc, baseState, scale, cssHeight, sec);

  // Endnotes (§17.11 endnotePr default position = document end) on the last
  // page, after the body flow. Minimal impl: a heading-less list at doc end.
  if (pageIndex === totalPages - 1) {
    drawEndnotes(doc, bodyState, scale, cssHeight, sec);
  }
}

/** Measure a note's content block in pt (paragraphs only), using a fresh
 *  pt-scale measure state. Returns the full height and the last paragraph's
 *  trailing spaceAfter (which overflows the bottom margin, like body text). */
function measureNoteBlockForDraw(
  note: DocNote,
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  sec: SectionProps,
  fontFamilyClasses: Record<string, string>,
  kinsoku: KinsokuRules,
): { total: number; trailingSpaceAfter: number } {
  const measure = buildMeasureState(ctx, sec, fontFamilyClasses, kinsoku);
  const contentWPt = sec.pageWidth - sec.marginLeft - sec.marginRight;
  return measureFootnoteBlockPt(note, measure, contentWPt);
}

/** Draw the footnote area for the current page: a separator rule (§17.11.9)
 *  followed by each referenced note's content, numbered. */
function drawPageFootnotes(
  elements: PaginatedBodyElement[],
  doc: DocxDocumentModel,
  baseState: RenderState,
  scale: number,
  cssHeight: number,
  sec: SectionProps,
): void {
  if (!doc.footnotes || doc.footnotes.length === 0) return;
  const noteById = indexNotes(doc.footnotes);

  // Collect referenced footnote ids in document (reading) order, de-duplicated.
  const ids: string[] = [];
  const seen = new Set<string>();
  const scan = (els: PaginatedBodyElement[]) => {
    for (const el of els) {
      if (el.type !== 'paragraph') continue;
      for (const id of footnoteRefsInRuns((el as unknown as DocParagraph).runs)) {
        if (!seen.has(id) && noteById.has(id)) {
          seen.add(id);
          ids.push(id);
        }
      }
    }
  };
  scan(elements);
  if (ids.length === 0) return;

  // Total block height (pt). The last note's trailing spaceAfter overflows the
  // bottom margin (like body text), so the block is positioned by its content
  // height — placing the last footnote line just above the bottom margin.
  let totalPt = 0;
  let lastTrailingPt = 0;
  for (const id of ids) {
    const note = noteById.get(id);
    if (!note) continue;
    const m = measureNoteBlockForDraw(note, baseState.ctx, sec, baseState.fontFamilyClasses, baseState.kinsoku);
    totalPt += m.total;
    lastTrailingPt = m.trailingSpaceAfter;
  }
  const contentPt = Math.max(0, totalPt - lastTrailingPt);
  const gapPx = FOOTNOTE_SEPARATOR_GAP_PT * scale;
  const blockTopY = cssHeight - sec.marginBottom * scale - contentPt * scale;

  // Separator rule: short left-aligned line (Word's default footnote separator,
  // ~1/3 of the text width), centered in the gap above the notes.
  const leftX = sec.marginLeft * scale;
  const ruleY = Math.round(blockTopY - gapPx);
  const ctx = baseState.ctx;
  ctx.save();
  ctx.strokeStyle = baseState.defaultColor;
  ctx.lineWidth = Math.max(1, Math.round(0.5 * scale));
  ctx.beginPath();
  ctx.moveTo(leftX, ruleY + 0.5);
  ctx.lineTo(leftX + (sec.pageWidth - sec.marginLeft - sec.marginRight) * scale / 3, ruleY + 0.5);
  ctx.stroke();
  ctx.restore();

  // Draw each note's content, numbered, flowing down from blockTopY.
  const noteState: RenderState = { ...baseState, y: blockTopY };
  for (const id of ids) {
    const note = noteById.get(id);
    if (!note) continue;
    noteState.currentNoteNumber = doc.footnotes.findIndex((n) => n.id === id) + 1;
    const paras = note.content.filter((e) => e.type === 'paragraph') as unknown as DocParagraph[];
    renderParaList(paras, noteState);
  }
}

/** Draw endnotes after the body flow on the final page (§17.11, default
 *  docEnd position). Each note is numbered with its endnote sequence. */
function drawEndnotes(
  doc: DocxDocumentModel,
  bodyState: RenderState,
  scale: number,
  cssHeight: number,
  sec: SectionProps,
): void {
  if (!doc.endnotes || doc.endnotes.length === 0) return;
  // Skip empty endnotes (only the reserved entries were filtered in the parser;
  // a note with no text contributes nothing).
  const notes = doc.endnotes.filter((n) =>
    n.content.some((e) => e.type === 'paragraph' && (e as unknown as DocParagraph).runs.length > 0),
  );
  if (notes.length === 0) return;

  // Continue from the body cursor with a small gap; clamp inside the bottom
  // margin. We do NOT spill endnotes onto a dedicated trailing page yet.
  const ctx = bodyState.ctx;
  let y = bodyState.y + FOOTNOTE_SEPARATOR_GAP_PT * 2 * scale;
  const maxY = cssHeight - sec.marginBottom * scale;
  if (y >= maxY) return;

  // Separator rule above the endnotes.
  const leftX = sec.marginLeft * scale;
  ctx.save();
  ctx.strokeStyle = bodyState.defaultColor;
  ctx.lineWidth = Math.max(1, Math.round(0.5 * scale));
  ctx.beginPath();
  ctx.moveTo(leftX, Math.round(y) + 0.5);
  ctx.lineTo(leftX + (sec.pageWidth - sec.marginLeft - sec.marginRight) * scale / 3, Math.round(y) + 0.5);
  ctx.stroke();
  ctx.restore();

  const noteState: RenderState = { ...bodyState, y: y + FOOTNOTE_SEPARATOR_GAP_PT * scale };
  for (const note of notes) {
    noteState.currentNoteNumber = doc.endnotes.findIndex((n) => n.id === note.id) + 1;
    const paras = note.content.filter((e) => e.type === 'paragraph') as unknown as DocParagraph[];
    renderParaList(paras, noteState);
  }
}

// ── Footnotes / endnotes ────────────────────────────────────────────────────
// ECMA-376 §17.11. A `<w:footnoteReference>` in the body (and the
// `<w:footnoteRef>` placeholder inside the note) is tagged on its run as
// `noteRef`. The DISPLAYED number is the note's 1-based position in document
// order (we don't honor numFmt/numStart overrides yet — the samples don't use
// them; see dev log). The note bodies are drawn at the bottom of the page that
// holds their first reference, above a short separator rule (§17.11.9).
// Endnotes (§17.11 endnotePr@pos, default docEnd) are appended after the body
// flow; we don't lay them on their own trailing page yet.

/** Vertical space (pt) Word allocates for the footnote separator region — the
 *  short rule plus the small leading above the first footnote. Word draws the
 *  separator in its own paragraph (the reserved id=-1 note, default single
 *  spacing) above the notes; one blank line's worth of leading is a faithful
 *  approximation for the common case. */
const FOOTNOTE_SEPARATOR_GAP_PT = 6;

/** Build a map from a note's `@w:id` to its 1-based sequential number, in the
 *  order the notes appear in footnotes.xml / endnotes.xml (ECMA-376 §17.11
 *  default decimal numbering, start=1, no restart). */
function buildNoteNumberMap(notes: DocNote[] | undefined): Map<string, number> {
  const m = new Map<string, number>();
  if (!notes) return m;
  notes.forEach((n, i) => m.set(n.id, i + 1));
  return m;
}

/** Index footnotes by id for content lookup. */
function indexNotes(notes: DocNote[] | undefined): Map<string, DocNote> {
  const m = new Map<string, DocNote>();
  if (!notes) return m;
  for (const n of notes) m.set(n.id, n);
  return m;
}

/** Collect, in document order, the footnote ids referenced by a paragraph's
 *  runs (kind === 'footnote'). Endnote refs are excluded — they aren't drawn
 *  per-page. */
function footnoteRefsInRuns(runs: DocRun[]): string[] {
  const ids: string[] = [];
  for (const r of runs) {
    if (r.type !== 'text') continue;
    const nr = (r as DocxTextRun).noteRef;
    if (nr && nr.kind === 'footnote' && nr.id) ids.push(nr.id);
  }
  return ids;
}

/** Measure one footnote's content block (pt). `total` is every paragraph's full
 *  height; `trailingSpaceAfter` is the last paragraph's `spaceAfter`, which —
 *  like a body paragraph's trailing space — may legally overflow the bottom
 *  margin and so is NOT counted when reserving page space. */
function measureFootnoteBlockPt(
  note: DocNote,
  state: RenderState,
  contentWPt: number,
): { total: number; trailingSpaceAfter: number } {
  let total = 0;
  let trailingSpaceAfter = 0;
  for (const el of note.content) {
    if (el.type !== 'paragraph') continue;
    const para = el as unknown as DocParagraph;
    total += estimateParagraphHeight(state, para, contentWPt, false);
    trailingSpaceAfter = para.spaceAfter;
  }
  return { total, trailingSpaceAfter };
}

/** Height (pt) to RESERVE on the page for a footnote: its content minus the
 *  trailing spaceAfter (overflows the bottom margin) plus the separator region. */
function footnoteReserveHeightPt(
  note: DocNote,
  state: RenderState,
  contentWPt: number,
  includeSeparator: boolean,
): number {
  const { total, trailingSpaceAfter } = measureFootnoteBlockPt(note, state, contentWPt);
  return Math.max(0, total - trailingSpaceAfter) + (includeSeparator ? FOOTNOTE_SEPARATOR_GAP_PT : 0);
}

/**
 * Split body into pages, honoring explicit page breaks AND measuring content
 * overflow for automatic pagination. All measurements are done in pt (scale=1).
 *
 * When `footnotes` is provided, the content area of each page shrinks by the
 * measured height of every footnote whose reference first appears on that page
 * (ECMA-376 §17.11): the footnote bodies occupy the bottom of the text column,
 * so the body must stop short of them.
 */
export function computePages(
  body: BodyElement[],
  section: SectionProps,
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  fontFamilyClasses: Record<string, string> = {},
  kinsoku: KinsokuRules = DEFAULT_KINSOKU_RULES,
  footnotes: DocNote[] = [],
): PaginatedBodyElement[][] {
  const fullContentH = section.pageHeight - section.marginTop - section.marginBottom;
  const contentW = section.pageWidth - section.marginLeft - section.marginRight;
  const measureState = buildMeasureState(ctx, section, fontFamilyClasses, kinsoku);
  const noteById = indexNotes(footnotes);
  const haveFootnotes = noteById.size > 0;
  // Per-page reserved footnote height (pt). Index 0 = first page. Grows as
  // footnote references are placed on the current page.
  const footnoteReservePt: number[] = [0];

  const pages: PaginatedBodyElement[][] = [[]];
  let y = 0;
  let prevPara: DocParagraph | null = null;
  // Word collapses the gap between two paragraphs to max(prev.spaceAfter,
  // this.spaceBefore) — i.e. it does NOT sum them. CSS-style "margin
  // collapsing." Matches Word's observed layout on demo/sample-1 (gap
  // between para 18 after=360 and para 19 before=240 is 18pt, not 30pt).
  let prevSpaceAfter = 0;
  // Keep measureState.y in sync with the current page's content Y so that
  // registerAnchorFloats/WrapLayoutCtx anchor relative to where we actually
  // are on the page. Anchor floats are registered on the measureState as
  // paragraphs are processed and cleared when we flip to a new page, exactly
  // like the real renderer does.
  measureState.y = section.marginTop;
  measureState.floats = [];
  // Footnote ids already reserved on the current page (so a paragraph that
  // references the same note twice doesn't double-count, and the renderer draws
  // each note once). Reset on every page flip.
  let pageNoteIds = new Set<string>();
  // Effective content height for the current page: the full text column minus
  // the footnote area reserved at the bottom of THIS page.
  const effContentH = () => fullContentH - (footnoteReservePt[pages.length - 1] ?? 0);
  const startPageBookkeeping = () => {
    footnoteReservePt[pages.length - 1] = 0;
    pageNoteIds = new Set<string>();
  };
  const newPage = () => {
    if (pages[pages.length - 1].length > 0) {
      pages.push([]);
      y = 0;
      prevPara = null;
      prevSpaceAfter = 0;
      measureState.y = section.marginTop;
      measureState.floats = [];
      startPageBookkeeping();
    }
  };

  // A paragraph-anchored floating object (wp:anchor with positionV
  // relativeFrom="paragraph"/"line", ECMA-376 §20.4.3.4) is positioned by its
  // anchor paragraph's top. Word keeps such a float on its page: if the float
  // would extend below the bottom margin, the layout engine relocates the
  // anchor paragraph to the next page so the float fits, rather than letting the
  // object spill past the page bottom. Return the largest distance from the
  // paragraph's content-top down to the bottom edge of any paragraph-anchored
  // float it carries, so the paginator can apply the same displacement. Returns
  // 0 when the paragraph has no paragraph-anchored floats (page-absolute floats
  // are pinned regardless of which page the anchor lands on, so they never
  // trigger a break). Measured at scale 1 (pt), matching the paginator's `y`.
  const anchoredFloatBottomOffset = (para: DocParagraph, spaceBeforePt: number): number => {
    let maxBottom = 0;
    for (const run of para.runs) {
      if (run.type !== 'image') continue;
      const img = run as unknown as ImageRun;
      if (!img.anchor || !img.anchorYFromPara) continue;
      // Wrap floats anchor after spaceBefore (registerAnchorFloats uses
      // state.y post-spaceBefore); non-wrap floats anchor at the paragraph's
      // pre-spaceBefore top (renderAnchorImages uses paragraphStartY). Mirror
      // each so the estimate matches the draw position exactly.
      const anchorBase = isWrapFloat(img.wrapMode) ? spaceBeforePt : 0;
      const bottom = anchorBase + (img.anchorYPt ?? 0) + img.heightPt;
      if (bottom > maxBottom) maxBottom = bottom;
    }
    return maxBottom;
  };

  // ECMA-376 §17.3.1.15: keepNext means this paragraph must stay on the same
  // page as the next paragraph. The simplest interpretation, and what Word
  // appears to do in practice, is "treat the keepNext chain as a single unit
  // for page-break purposes" — so here we look ahead and add the next
  // paragraph's (or first line's) height to the break decision.
  const estimateNextBlockHeight = (startIdx: number): number => {
    const nxt = body[startIdx];
    if (!nxt) return 0;
    if (nxt.type === 'paragraph') {
      // We only need enough room for the first line so that "keepNext" avoids
      // orphaning the current paragraph at the bottom of a page while the
      // next begins on a new page. Using the full paragraph is safer for a
      // single-line next; for multi-line we rely on that paragraph's own
      // break logic after placing.
      return estimateParagraphHeight(measureState, nxt as unknown as DocParagraph, contentW, false);
    }
    if (nxt.type === 'table') {
      return estimateTableHeight(measureState, nxt as unknown as DocTable, contentW);
    }
    return 0;
  };

  for (let i = 0; i < body.length; i++) {
    const el = body[i];
    if (el.type === 'pageBreak') {
      pages.push([]);
      y = 0;
      prevPara = null;
      prevSpaceAfter = 0;
      measureState.y = section.marginTop;
      measureState.floats = [];
      startPageBookkeeping();
      // ECMA-376 §17.18.79 ST_SectionMark: oddPage / evenPage breaks pad
      // with a blank page when the new section would otherwise start on the
      // wrong parity. pages.length here is the next page's 1-based index.
      if (el.parity === 'odd' && pages.length % 2 === 0) {
        pages.push([]);
        startPageBookkeeping();
      } else if (el.parity === 'even' && pages.length % 2 === 1) {
        pages.push([]);
        startPageBookkeeping();
      }
      continue;
    }
    if (el.type === 'paragraph') {
      const para = el as unknown as DocParagraph;
      if (para.pageBreakBefore) newPage();
      const suppressBefore = contextualSuppressed(prevPara, para);

      // Collapse with the previous paragraph's spaceAfter — Word takes
      // max(prev.after, this.before) between paragraphs, not the sum.
      const effectiveBefore = suppressBefore ? 0 : para.spaceBefore;
      const overlap = Math.min(prevSpaceAfter, effectiveBefore);
      y -= overlap;
      measureState.y -= overlap;

      // Register this paragraph's anchor-image floats on the measureState so
      // subsequent paragraphs estimate around them (text-wrap around images
      // adds lines that the float-unaware estimate would otherwise miss,
      // which caused page 2 of demo/sample-1 to spill past the bottom
      // margin). The renderer runs the same registerAnchorFloats call; by
      // mirroring it here the paginator sees the same layout.
      const paragraphAnchorY = measureState.y + effectiveBefore;
      registerAnchorFloats(para, measureState, paragraphAnchorY);

      const h = estimateParagraphHeight(measureState, para, contentW, suppressBefore, section.marginLeft);

      // ECMA-376 §17.11: a footnote shares the page with its reference, so the
      // body must stop short of the footnote area. Measure the footnotes this
      // paragraph newly references (not already reserved on this page) and fold
      // their height into the fit decision — if the paragraph + its footnote(s)
      // don't fit, both move to the next page together.
      let newRefIds: string[] = [];
      let addReservePt = 0;
      // Sum the reserve for a set of newly-referenced notes, charging the
      // separator region only to the first note on a page (when that page has
      // no reserve yet).
      const sumReserve = (ids: string[]): number => {
        let sum = 0;
        for (let k = 0; k < ids.length; k++) {
          const note = noteById.get(ids[k]);
          if (!note) continue;
          const firstOnPage = (footnoteReservePt[pages.length - 1] ?? 0) === 0 && k === 0;
          sum += footnoteReserveHeightPt(note, measureState, contentW, firstOnPage);
        }
        return sum;
      };
      if (haveFootnotes) {
        const seen = new Set<string>();
        for (const id of footnoteRefsInRuns(para.runs)) {
          if (pageNoteIds.has(id) || seen.has(id)) continue;
          if (!noteById.has(id)) continue;
          seen.add(id);
          newRefIds.push(id);
        }
        addReservePt = sumReserve(newRefIds);
      }

      // Break if this paragraph alone doesn't fit, OR if keepNext is set and
      // placing it would leave no room for the next block on the same page.
      // Per Word's layout behavior: `spaceAfter` is trailing whitespace that
      // can legally overflow the bottom of the page — only content + spaceBefore
      // must fit. This is what lets a closing paragraph with a large
      // `w:spacing/@w:after` land flush against the bottom margin.
      const needNext = para.keepNext ? estimateNextBlockHeight(i + 1) : 0;
      const fitHeight = h - para.spaceAfter;
      const needed = fitHeight + needNext + addReservePt;
      // A paragraph-anchored float must fit below the paragraph's top on the
      // same page. If it overflows the bottom margin here but would fit when the
      // paragraph starts a fresh page, displace the paragraph (Word's float
      // keep-on-page behavior). When the float is taller than the page content
      // area it can never fit — leave it on this page and allow the overflow
      // (no break would help, and breaking unconditionally would loop forever).
      const floatBottomOff = anchoredFloatBottomOffset(para, effectiveBefore);
      const floatOverflowsHere = floatBottomOff > 0 && y + floatBottomOff > effContentH();
      const floatFitsFresh = floatBottomOff > 0 && floatBottomOff <= effContentH();
      const breakForFloat = y > 0 && floatOverflowsHere && floatFitsFresh;
      if ((y > 0 && y + needed > effContentH()) || breakForFloat) {
        newPage();
        // newPage() cleared measureState.floats; re-register this paragraph's
        // anchor floats against the new page's top so wrap-around estimates for
        // this and later paragraphs see them at the correct (post-break) Y.
        registerAnchorFloats(para, measureState, measureState.y + effectiveBefore);
        // The references move to the new page; nothing was reserved there yet,
        // so the separator region still applies to the first footnote.
        if (haveFootnotes && newRefIds.length > 0) addReservePt = sumReserve(newRefIds);
      }

      // ECMA-376 doesn't say "paragraphs must fit on one page" — Word
      // splits long paragraphs at line boundaries. If the paragraph is
      // taller than the remaining content area, walk the laid-out lines
      // and emit one PaginatedBodyElement slice per page.
      const pageContentH = effContentH();
      const remainingH = pageContentH - y;
      if (h > remainingH && h > pageContentH * 0.5) {
        const placed = splitParagraphAcrossPages(
          measureState, para, contentW, suppressBefore, section.marginLeft,
          y, pageContentH, pages,
          () => { newPage(); },
        );
        // After splitting, `y` is the bottom of the last slice on the
        // current page (continues for the LAST slice; intermediate slices
        // filled their pages exactly, so newPage was called between them).
        y = placed.endY;
        measureState.y = section.marginTop + placed.endY;
        // A split footnote-bearing paragraph reserves on the page where it
        // ends. Rare; the separator region re-applies if that page had none.
        if (haveFootnotes && newRefIds.length > 0) {
          newRefIds = newRefIds.filter((id) => !pageNoteIds.has(id));
          addReservePt = sumReserve(newRefIds);
        }
      } else {
        pages[pages.length - 1].push(el as PaginatedBodyElement);
        y += h;
        measureState.y += h;
      }

      // Commit the footnote reserve onto the page the paragraph landed on.
      if (haveFootnotes && newRefIds.length > 0) {
        const idx = pages.length - 1;
        footnoteReservePt[idx] = (footnoteReservePt[idx] ?? 0) + addReservePt;
        for (const id of newRefIds) pageNoteIds.add(id);
      }

      prevPara = para;
      prevSpaceAfter = para.spaceAfter;
    } else if (el.type === 'table') {
      const tbl = el as unknown as DocTable;
      const rowHs = computeTableRowHeights(measureState, tbl, contentW);
      const h = rowHs.reduce((s, x) => s + x, 0);
      // Footnote references inside table cells are not folded into the reserve
      // (the per-page reserve is driven by body paragraphs); they still draw at
      // page bottom via the renderer's page scan. effContentH() respects any
      // reserve already accumulated on this page.
      const tableContentH = effContentH();
      if (h > tableContentH) {
        // Taller than a full page: split row-by-row so the overflow continues
        // onto the next page instead of being clipped (ECMA-376 table
        // pagination). Tables that fit on a page keep the simple place-whole
        // path below.
        const endY = splitTableAcrossPages(tbl, rowHs, y, tableContentH, pages, () => newPage());
        y = endY;
        measureState.y = section.marginTop + endY;
      } else {
        if (y + h > tableContentH) newPage();
        pages[pages.length - 1].push(el);
        y += h;
        measureState.y += h;
      }
      prevPara = null;
    }
  }
  return pages;
}

/** Paginate with a throwaway measure context. Pagination must use the same
 *  fontFamilyClasses + kinsoku rules as the render path, otherwise line-break
 *  decisions (and thus page breaks) diverge between measurement and paint
 *  (ECMA-376 §17.15.1.58–.60). Shared by the main-thread DocxDocument and the
 *  render worker so the two modes can never paginate differently. */
export function paginateDocument(doc: DocxDocumentModel): PaginatedBodyElement[][] {
  const ctx = new OffscreenCanvas(1, 1).getContext('2d');
  if (!ctx) return [doc.body];
  return computePages(
    doc.body,
    doc.section,
    ctx,
    doc.fontFamilyClasses ?? {},
    resolveKinsokuRules(doc.settings),
    doc.footnotes ?? [],
  );
}

function buildMeasureState(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  section: SectionProps,
  fontFamilyClasses: Record<string, string> = {},
  kinsoku: KinsokuRules = DEFAULT_KINSOKU_RULES,
): RenderState {
  return {
    ctx,
    scale: 1,
    contentX: 0,
    contentW: section.pageWidth - section.marginLeft - section.marginRight,
    y: 0,
    pageH: section.pageHeight,
    defaultColor: '#000000',
    pageIndex: 0,
    totalPages: 1,
    images: new Map(),
    dryRun: true,
    marginLeft: section.marginLeft,
    marginRight: section.marginRight,
    marginTop: section.marginTop,
    marginBottom: section.marginBottom,
    pageWidth: section.pageWidth,
    floats: [],
    docGrid: { type: section.docGridType ?? null, linePitchPt: section.docGridLinePitch ?? null },
    fontFamilyClasses,
    kinsoku,
    showTrackChanges: false,
  };
}

function estimateParagraphHeight(
  state: RenderState,
  para: DocParagraph,
  contentWPt: number,
  suppressSpaceBefore = false,
  paraXPt = 0,
): number {
  const indLeft = para.indentLeft;
  const indRight = para.indentRight;
  const paraW = Math.max(1, contentWPt - indLeft - indRight);
  const segs = buildSegments(para.runs, state);
  // Word renders ruby paragraphs with consistent line spacing — every line
  // in a paragraph that carries ANY furigana snaps to the same pitch
  // multiple, otherwise mixed-ruby paragraphs jitter and pagination drifts.
  const paraHasRuby = paragraphHasRuby(para);
  const grid = paraGrid(para, state);
  let textH: number;
  if (segs.length === 0) {
    textH = paragraphMarkLineHeight(para, 1, grid, paraHasRuby);
  } else {
    // When anchor-image floats are active on the current page the paragraph
    // wraps around them, adding lines compared to a full-width layout. Use
    // the same WrapLayoutCtx the renderer uses so estimate and render agree.
    const wrapCtx: WrapLayoutCtx | undefined = state.floats.length > 0 ? {
      startPageY: state.y,
      paraX: paraXPt,
      floats: state.floats,
      lineBoxH: (a, d, _h, is) => lineBoxHeight(para.lineSpacing, a, d, 1, grid, paraHasRuby, is ?? 0),
      pageH: state.pageH,
    } : undefined;
    const lines = layoutLines(state.ctx, segs, paraW, para.indentFirst, 1, para.tabStops, wrapCtx, state.fontFamilyClasses, indLeft, state.kinsoku);
    if (lines.length === 0) {
      // Anchor-only paragraph: layoutLines placed no inline content but the
      // paragraph mark still occupies one line (§17.3.1.29) — mirror the
      // renderer so pagination math agrees.
      textH = paragraphMarkLineHeight(para, 1, grid, paraHasRuby);
    } else if (paraHasRuby) {
      // Word uses the same line height for every line in a ruby paragraph,
      // snapped to an integer docGrid pitch.
      const uniform = snapParaLineToGrid(
        Math.max(0, ...lines.map(l => lineBoxHeight(para.lineSpacing, l.ascent, l.descent, 1, grid, true, l.intendedSingle))),
        grid,
        1,
      );
      textH = uniform * lines.length;
    } else {
      textH = lines.reduce((s, l) => s + lineBoxHeight(para.lineSpacing, l.ascent, l.descent, 1, grid, false, l.intendedSingle), 0);
    }
  }
  return textH + (suppressSpaceBefore ? 0 : para.spaceBefore) + para.spaceAfter;
}

/** Snap a paragraph's uniform line height up to an integer multiple of the
 *  docGrid pitch. Mirrors Word's docGrid handling for ruby paragraphs:
 *  the grid pitch widens to accommodate the tallest required line, and
 *  every line in the paragraph then uses that widened pitch. */
function snapParaLineToGrid(h: number, grid: DocGridCtx | undefined, scale: number): number {
  if (!grid || !grid.linePitchPt || grid.linePitchPt <= 0) return h;
  if (grid.type !== 'lines' && grid.type !== 'linesAndChars') return h;
  const pitchPx = grid.linePitchPt * scale;
  if (pitchPx <= 0) return h;
  if (h <= pitchPx) return pitchPx;
  return Math.ceil(h / pitchPx) * pitchPx;
}

/** Return true when any text run in the paragraph carries a `ruby` annotation.
 *  Used to apply paragraph-wide line-height snapping to docGrid pitch — Word
 *  renders the entire ruby paragraph with consistent line spacing so that
 *  ruby-bearing and ruby-free lines line up on the same baseline grid. */
function paragraphHasRuby(para: DocParagraph): boolean {
  for (const run of para.runs) {
    if (run.type === 'text' && (run as unknown as DocxTextRun).ruby) return true;
  }
  return false;
}

/** The docGrid that governs a paragraph's line heights. ECMA-376 §17.3.1.32:
 *  a paragraph with `w:snapToGrid` explicitly off ignores the section grid, so
 *  its lines use natural font metrics / the spacing multiplier directly. */
function paraGrid(para: DocParagraph, state: RenderState): DocGridCtx {
  return para.snapToGrid === false ? { type: null, linePitchPt: null } : state.docGrid;
}

/** Lay out a paragraph's lines, then walk the line list distributing them
 *  across pages whenever the cumulative height would exceed the page bottom.
 *  Each per-page chunk is appended to `pages` as a `lineSlice`-tagged
 *  PaginatedBodyElement — the renderer reads `lineSlice` and renders only
 *  that index range, padding the leading/trailing space-before/after on
 *  the appropriate sides.
 *
 *  Returns the Y where the FINAL slice ends on the current (last) page, so
 *  the caller can advance `y` / `measureState.y` accordingly.
 */
function splitParagraphAcrossPages(
  measureState: RenderState,
  para: DocParagraph,
  contentWPt: number,
  suppressSpaceBefore: boolean,
  marginLeftPt: number,
  initialY: number,
  contentH: number,
  pages: PaginatedBodyElement[][],
  newPage: () => void,
): { endY: number } {
  const indLeft = para.indentLeft;
  const indRight = para.indentRight;
  const paraW = Math.max(1, contentWPt - indLeft - indRight);
  const segs = buildSegments(para.runs, measureState);
  if (segs.length === 0) {
    // No layoutable content — treat as a single empty line, fits or pushes.
    pages[pages.length - 1].push(para as PaginatedBodyElement);
    return { endY: initialY + estimateParagraphHeight(measureState, para, contentWPt, suppressSpaceBefore, marginLeftPt) };
  }
  const wrapCtx: WrapLayoutCtx | undefined = measureState.floats.length > 0 ? {
    startPageY: measureState.y,
    paraX: marginLeftPt,
    floats: measureState.floats,
    lineBoxH: (a, d, _h, is) => lineBoxHeight(para.lineSpacing, a, d, 1, measureState.docGrid, paragraphHasRuby(para), is ?? 0),
    pageH: measureState.pageH,
  } : undefined;
  const lines = layoutLines(measureState.ctx, segs, paraW, para.indentFirst, 1, para.tabStops, wrapCtx, measureState.fontFamilyClasses, indLeft, measureState.kinsoku);
  if (lines.length === 0) {
    // Anchor-only paragraph: no inline lines to split, but the paragraph mark
    // still occupies one line (§17.3.1.29). Emit it as a single unit, matching
    // the literal-empty branch above so the renderer's one-line advance agrees.
    pages[pages.length - 1].push(para as PaginatedBodyElement);
    return { endY: initialY + estimateParagraphHeight(measureState, para, contentWPt, suppressSpaceBefore, marginLeftPt) };
  }
  const paraHasRuby = paragraphHasRuby(para);

  const perLineH = (l: typeof lines[number]) => lineBoxHeight(para.lineSpacing, l.ascent, l.descent, 1, measureState.docGrid, paraHasRuby, l.intendedSingle);
  const uniformH = paraHasRuby
    ? snapParaLineToGrid(Math.max(0, ...lines.map(perLineH)), measureState.docGrid, 1)
    : 0;
  const lineHeights = lines.map(l => paraHasRuby ? uniformH : perLineH(l));
  const spaceBefore = suppressSpaceBefore ? 0 : para.spaceBefore;
  const spaceAfter = para.spaceAfter;

  let lineIdx = 0;
  let cursorY = initialY;
  let isFirstSliceOnPage = true; // first slice carries spaceBefore
  while (lineIdx < lines.length) {
    // Available space on the current page from cursorY downward.
    const remaining = contentH - cursorY;
    // First slice on a page reserves spaceBefore; the LAST slice (covering
    // the final line) reserves spaceAfter.
    const sliceLeading = isFirstSliceOnPage ? spaceBefore : 0;
    let usedH = sliceLeading;
    let firstFitting = lineIdx;
    let lastFitting = lineIdx;
    while (lastFitting < lines.length && usedH + lineHeights[lastFitting] <= remaining) {
      usedH += lineHeights[lastFitting];
      lastFitting++;
    }
    if (lastFitting === firstFitting) {
      // Not even one line fits on this page — flush to a new page and retry.
      // Guard against infinite loop: if we're already at the start of a
      // fresh page (cursorY ≈ 0) and still don't fit, force-emit one line.
      if (cursorY > 0) {
        newPage();
        cursorY = 0;
        isFirstSliceOnPage = true;
        continue;
      }
      // First page, first line doesn't fit — force-emit it and let it overflow.
      lastFitting = firstFitting + 1;
      usedH += lineHeights[firstFitting];
    }
    const isFinalSlice = lastFitting === lines.length;
    if (isFinalSlice) usedH += spaceAfter;
    pages[pages.length - 1].push({
      ...(para as object),
      type: 'paragraph',
      lineSlice: { start: firstFitting, end: lastFitting },
    } as PaginatedBodyElement);
    lineIdx = lastFitting;
    cursorY += usedH;
    if (!isFinalSlice) {
      newPage();
      cursorY = 0;
      isFirstSliceOnPage = true;
    }
  }
  return { endY: cursorY };
}

/** Per-row heights used by both pagination and the height estimate. Mirrors the
 *  renderer's row sizing (exact / atLeast / auto + vMerge span distribution,
 *  ECMA-376 §17.4.80, §17.4.85). */
function computeTableRowHeights(state: RenderState, table: DocTable, contentWPt: number): number[] {
  const colWidths = resolveColumnWidths(table, contentWPt, state);

  const rowHs: number[] = [];
  const restartInfo: Array<{ ri: number; ci: number; contentH: number }> = [];

  for (let ri = 0; ri < table.rows.length; ri++) {
    const row = table.rows[ri];
    if (row.rowHeight != null && row.rowHeightRule === 'exact') {
      rowHs.push(row.rowHeight);
      continue;
    }
    let rowH = row.rowHeight != null ? row.rowHeight : 10;
    let ci = 0;
    for (const cell of row.cells) {
      const span = Math.min(cell.colSpan, colWidths.length - ci);
      const cellW = colWidths.slice(ci, ci + span).reduce((s, w) => s + w, 0);
      const cm = effCellMargins(cell, table);
      const innerW = Math.max(1, cellW - cm.left - cm.right);
      let ch = cm.top + cm.bottom;
      for (const ce of cell.content) {
        if (ce.type === 'paragraph') {
          ch += estimateParagraphHeight(state, ce as unknown as DocParagraph, innerW);
        } else if (ce.type === 'table') {
          ch += estimateTableHeight(state, ce as unknown as DocTable, innerW);
        }
      }
      // ECMA-376 §17.4.85: vMerge=restart cell content spans the merged region;
      // do not inflate the first row alone. vMerge=false (continue) renders no
      // content. Both are deferred to the post-pass below.
      if (cell.vMerge === true) {
        restartInfo.push({ ri, ci, contentH: ch });
      } else if (cell.vMerge !== false) {
        if (ch > rowH) rowH = ch;
      }
      ci += span;
    }
    rowHs.push(rowH);
  }

  for (const info of restartInfo) {
    const endRi = findMergeEndRow(table, info.ri, info.ci);
    let spanH = 0;
    for (let rj = info.ri; rj <= endRi; rj++) spanH += rowHs[rj];
    if (spanH < info.contentH) {
      rowHs[endRi] += info.contentH - spanH;
    }
  }

  return rowHs;
}

function estimateTableHeight(state: RenderState, table: DocTable, contentWPt: number): number {
  return computeTableRowHeights(state, table, contentWPt).reduce((s, x) => s + x, 0);
}

/**
 * Resolve the per-grid-column widths (pt) for a table, honoring Word's table
 * layout algorithm (ECMA-376 §17.4.52 tblLayout + §17.4.71/§17.4.63 tcW/tblW).
 *
 *   - layout === 'fixed': use the tblGrid widths verbatim (the historical
 *     behavior), then scale down proportionally if the grid total overflows the
 *     available content width.
 *   - layout absent or 'autofit' (the spec default): each grid column's width is
 *     the maximum *preferred* width (cell `widthPt`, i.e. `<w:tcW type="dxa">`,
 *     or `widthPct` resolved against `contentWPt`) over the cells anchored in
 *     it. A gridSpan cell contributes its preference distributed across the
 *     columns it spans, in proportion to those columns' tblGrid widths, but
 *     only raises a column above what single-column cells already require.
 *     Columns with no preference anywhere keep their tblGrid width. If the
 *     resulting table width exceeds `contentWPt`, all columns are scaled
 *     proportionally to fit — this is what Word does when the preferred-width
 *     sum overflows the text column.
 *
 * The returned widths sum to at most `contentWPt`. Both `renderTable` (which
 * then multiplies by the device scale) and `computeTableRowHeights` (which
 * works directly in pt) consume this.
 */
/** Minimum content width (pt) of a single cell: the widest non-breakable token
 *  across its paragraphs plus the cell's left/right margins. A column can never
 *  be narrower than this without clipping/wrapping the longest word — Word's
 *  auto-fit (ECMA-376 §17.4.52) treats it as the column's hard floor when the
 *  preferred widths overflow, which is why date strings like "2026-02-28" stay
 *  on one line even though their preferred `tcW` gets squeezed. */
function cellMinContentPt(cell: DocTableCell, table: DocTable, state: RenderState): number {
  const { ctx, fontFamilyClasses } = state;
  let maxTokenPt = 0;
  const scanPara = (para: DocParagraph): void => {
    for (const run of para.runs) {
      if (run.type !== 'text') continue;
      const t = run as unknown as DocxTextRun & { type: 'text' };
      if (!t.text) continue;
      // Resolve the complex-script (cs) axis the same way `buildSegments` does
      // (ECMA-376 §17.3.2.3/§17.3.2.17/§17.3.2.18): a run forcing cs (w:rtl or
      // the §17.3.2.7 <w:cs/> toggle) measures with the cs bold/italic/size/
      // family, each falling back to its Latin counterpart when absent. We pick
      // the cs axis for the min-content estimate when the run forces cs so wide
      // Arabic/Hebrew tokens reserve enough column width; otherwise the Latin
      // axis. NOTE rFonts@cs alone is just a font SLOT — it must not force cs
      // (a Latin heading whose style defines cstheme would wrongly take szCs).
      const forceCs = t.rtl === true || t.cs === true;
      const effBold = forceCs ? (t.boldCs ?? t.bold) : t.bold;
      const effItalic = forceCs ? (t.italicCs ?? t.italic) : t.italic;
      const effFontSize = forceCs ? (t.fontSizeCs ?? t.fontSize) : t.fontSize;
      const effFontFamily = forceCs ? (t.fontFamilyCs ?? t.fontFamily) : t.fontFamily;
      // Measure in pt-space (font size in pt, scale 1) so the result composes
      // directly with the pt-based column math.
      ctx.font = buildFont(effBold, effItalic, effFontSize, effFontFamily, fontFamilyClasses);
      // Non-breakable tokens are the same units the line layout treats as
      // atomic (UAX#14 has no break inside a word or between a digit and an
      // adjacent hyphen/period). CJK breaks per-glyph, so its min is one glyph.
      for (const piece of t.text.split('\t')) {
        for (const token of splitTextForLayout(piece)) {
          const trimmed = token.replace(/\s+$/u, '');
          if (!trimmed) continue;
          const w = hasCJKBreakOpportunity(trimmed)
            ? Math.max(...[...trimmed].map((ch) => ctx.measureText(ch).width))
            : ctx.measureText(trimmed).width;
          if (w > maxTokenPt) maxTokenPt = w;
        }
      }
    }
  };
  for (const ce of cell.content) {
    if (ce.type === 'paragraph') scanPara(ce as unknown as DocParagraph);
    // Nested tables: their own columns handle min-width; skip here.
  }
  if (maxTokenPt === 0) return 0;
  const cm = effCellMargins(cell, table);
  return maxTokenPt + cm.left + cm.right;
}

function resolveColumnWidths(table: DocTable, contentWPt: number, state: RenderState): number[] {
  const n = table.colWidths.length;
  if (n === 0) return [];

  const grid = table.colWidths;

  // Per-column minimum content width (pt). Single-column cells set a hard floor;
  // a gridSpan cell's min is distributed across its columns in proportion to the
  // tblGrid so a wide spanning cell does not over-inflate any one column.
  const minW: number[] = new Array(n).fill(0);
  for (const row of table.rows) {
    let ci = 0;
    for (const cell of row.cells) {
      const span = Math.min(Math.max(cell.colSpan, 1), n - ci);
      const m = cellMinContentPt(cell, table, state);
      if (m > 0) {
        if (span === 1) {
          if (m > minW[ci]) minW[ci] = m;
        } else {
          const spanCols = grid.slice(ci, ci + span);
          const gridSum = spanCols.reduce((s, w) => s + w, 0);
          for (let k = 0; k < span; k++) {
            const share = gridSum > 0 ? spanCols[k] / gridSum : 1 / span;
            const part = m * share;
            if (part > minW[ci + k]) minW[ci + k] = part;
          }
        }
      }
      ci += span;
    }
  }

  // Fit a desired-width vector (preferred `tcW`, floored at min content) to the
  // available width. When the desired total overflows, Word's auto-fit
  // (ECMA-376 §17.4.52) distributes the available width in proportion to each
  // column's PREFERRED width — but never below its minimum content width. A
  // column pinned to its min keeps that width and the leftover space is shared
  // among the still-shrinkable columns (again proportional to preferred). This
  // is what makes a date column with a small preferred width settle at exactly
  // the date string's width (its content min) while a column with a large
  // preferred width keeps proportionally more, matching Word's PDF layout.
  const fitToContent = (widths: number[]): number[] => {
    const total = widths.reduce((s, w) => s + w, 0);
    if (total <= contentWPt || total <= 0) return widths;
    const minTotal = minW.reduce((s, w) => s + w, 0);
    if (minTotal >= contentWPt) {
      // Even the minimums overflow — scale the minimums so the table still
      // fits (content clips, as Word does when forced narrower than its words).
      const s = contentWPt / minTotal;
      return minTotal > 0 ? minW.map((w) => w * s) : widths.map(() => contentWPt / n);
    }
    // Iteratively pin columns to their min and redistribute the rest by
    // preferred-width proportion. Converges in ≤ n passes (each pass pins at
    // least one new column or finishes). `widths` already encodes the preferred
    // width (floored at min) per column, used as the distribution weight.
    const out = widths.slice();
    const pinned = new Array(n).fill(false);
    for (let pass = 0; pass < n; pass++) {
      let free = contentWPt;
      let weightSum = 0;
      for (let c = 0; c < n; c++) {
        if (pinned[c]) free -= out[c];
        else weightSum += widths[c];
      }
      if (weightSum <= 0) break;
      let pinnedAny = false;
      for (let c = 0; c < n; c++) {
        if (pinned[c]) continue;
        const share = free * (widths[c] / weightSum);
        if (share < minW[c]) {
          out[c] = minW[c];
          pinned[c] = true;
          pinnedAny = true;
        } else {
          out[c] = share;
        }
      }
      if (!pinnedAny) break;
    }
    return out;
  };

  // Fixed layout: tblGrid is authoritative (ECMA-376 §17.4.52) and content is
  // clipped, never grown — so scale proportionally to fit, ignoring min content
  // widths (which only govern the autofit branch below).
  if (table.layout === 'fixed') {
    const g = grid.slice();
    const total = g.reduce((s, w) => s + w, 0);
    if (total > contentWPt && total > 0) {
      const s = contentWPt / total;
      return g.map((w) => w * s);
    }
    return g;
  }

  // Autofit (default): preferred widths drive the column sizes.
  // `pref[c]` accumulates the strongest single-column preference seen so far.
  const pref: number[] = new Array(n).fill(0);
  // `gridFallback[c]` is true while no cell has expressed a preference for c.
  const hasPref: boolean[] = new Array(n).fill(false);

  // Translate a cell's `<w:tcW>` into a preferred width in pt, or null when the
  // cell expresses no preference (auto/nil). pct is a fraction of contentWPt
  // (50ths of a percent — §17.18.90 ST_TblWidth).
  const cellPreferred = (cell: DocTableCell): number | null => {
    if (cell.widthPt != null) return cell.widthPt;
    if (cell.widthPct != null) return (cell.widthPct / 5000) * contentWPt;
    return null;
  };

  // First pass: single-column (non-spanning) cells set a hard per-column floor.
  for (const row of table.rows) {
    let ci = 0;
    for (const cell of row.cells) {
      const span = Math.min(Math.max(cell.colSpan, 1), n - ci);
      if (span === 1) {
        const p = cellPreferred(cell);
        if (p != null) {
          if (p > pref[ci]) pref[ci] = p;
          hasPref[ci] = true;
        }
      }
      ci += span;
    }
  }

  // Second pass: gridSpan cells distribute their preference across the spanned
  // columns in proportion to the tblGrid widths, but only raise columns that
  // are still below their share (so a single-column floor is never lowered).
  for (const row of table.rows) {
    let ci = 0;
    for (const cell of row.cells) {
      const span = Math.min(Math.max(cell.colSpan, 1), n - ci);
      if (span > 1) {
        const p = cellPreferred(cell);
        if (p != null) {
          const spanCols = grid.slice(ci, ci + span);
          const gridSum = spanCols.reduce((s, w) => s + w, 0);
          // Current resolved width across the span (grid fallback where unset).
          const curSum = spanCols.reduce(
            (s, w, k) => s + (hasPref[ci + k] ? pref[ci + k] : w),
            0,
          );
          // Only widen when the span's preference exceeds what we already have.
          if (p > curSum) {
            const extra = p - curSum;
            for (let k = 0; k < span; k++) {
              const share = gridSum > 0 ? spanCols[k] / gridSum : 1 / span;
              const base = hasPref[ci + k] ? pref[ci + k] : grid[ci + k];
              pref[ci + k] = base + extra * share;
              hasPref[ci + k] = true;
            }
          }
        }
      }
      ci += span;
    }
  }

  // Columns with no preference anywhere fall back to their tblGrid width. A
  // column's desired width is then floored at its minimum content width: a
  // preferred `tcW` narrower than the longest unbreakable token cannot actually
  // be honored (ECMA-376 §17.4.52 — auto layout grows a column to fit content).
  const widths = pref.map((p, c) => Math.max(hasPref[c] ? p : grid[c], minW[c]));
  return fitToContent(widths);
}

/** A page break before row `ri` is unsafe when `ri` continues a vertical merge
 *  started above (ECMA-376 §17.4.85): splitting there would orphan the merged
 *  cell's continuation. Such a row carries at least one `vMerge=false` cell. */
function tableBreakAllowedBefore(table: DocTable, ri: number): boolean {
  if (ri <= 0) return true;
  return !table.rows[ri].cells.some((c) => c.vMerge === false);
}

/**
 * Split a table that is taller than one page across page boundaries, row by
 * row (ECMA-376 table pagination). Each page receives a {@link DocTable} slice
 * holding the rows that fit; `w:tblHeader` rows (§17.4.78) repeat at the top of
 * every continuation. Breaks land only on vMerge-safe boundaries
 * ({@link tableBreakAllowedBefore}). Returns the Y offset after the final slice
 * on the last page.
 */
export function splitTableAcrossPages(
  table: DocTable,
  rowHs: number[],
  startY: number,
  contentH: number,
  pages: PaginatedBodyElement[][],
  newPage: () => void,
): number {
  const n = table.rows.length;
  // Leading tblHeader rows repeat on each continuation page.
  let headerCount = 0;
  while (headerCount < n && table.rows[headerCount].isHeader) headerCount++;
  const headerRows = table.rows.slice(0, headerCount);
  const headerH = rowHs.slice(0, headerCount).reduce((s, h) => s + h, 0);

  let y = startY;
  let start = 0;
  let firstSlice = true;

  while (start < n) {
    const isContinuation = !firstSlice && headerCount > 0 && start >= headerCount;
    const avail = contentH - y;
    let used = isContinuation ? headerH : 0;
    let end = start;
    // Always place at least one row to guarantee forward progress.
    while (end < n) {
      const h = rowHs[end];
      if (end > start && used + h > avail && tableBreakAllowedBefore(table, end)) break;
      used += h;
      end++;
    }

    const bodyRows = table.rows.slice(start, end);
    const sliceRows = isContinuation ? [...headerRows, ...bodyRows] : bodyRows;
    pages[pages.length - 1].push({ ...table, type: 'table', rows: sliceRows } as PaginatedBodyElement);

    y += used;
    start = end;
    firstSlice = false;
    if (start < n) {
      newPage();
      y = 0;
    }
  }
  return y;
}

/** Find the last row index in a vMerge span starting at (startRi, startCi) —
 *  i.e. walk forward while subsequent rows have a cell at column-start ci with
 *  vMerge=false (continue). ECMA-376 §17.4.85. */
function findMergeEndRow(table: DocTable, startRi: number, startCi: number): number {
  let endRi = startRi;
  for (let rj = startRi + 1; rj < table.rows.length; rj++) {
    const row = table.rows[rj];
    let ci = 0;
    let matched = false;
    for (const cell of row.cells) {
      if (ci === startCi) {
        if (cell.vMerge === false) matched = true;
        break;
      }
      if (ci > startCi) break;
      ci += cell.colSpan;
    }
    if (!matched) break;
    endRi = rj;
  }
  return endRi;
}

function pickHeaderFooter(
  set: DocxDocumentModel['headers'],
  pageIndex: number,
  _totalPages: number,
  titlePage: boolean,
  evenAndOdd: boolean,
): HeaderFooter | null {
  if (titlePage && pageIndex === 0 && set.first) return set.first;
  if (evenAndOdd && pageIndex % 2 === 1 && set.even) return set.even;
  return set.default ?? null;
}

function renderHeaderFooter(hf: HeaderFooter, topY: number, base: RenderState): void {
  const state: RenderState = { ...base, y: topY };
  renderBodyElements(hf.body, state);
}

function measureHeaderFooterHeight(hf: HeaderFooter, base: RenderState): number {
  const state: RenderState = { ...base, y: 0, dryRun: true, floats: [] };
  renderBodyElements(hf.body, state);
  return state.y;
}

// ===== Body element dispatch =====

function renderBodyElement(el: BodyElement, state: RenderState): void {
  if (el.type === 'paragraph') {
    renderParagraph(el as unknown as DocParagraph, state);
  } else if (el.type === 'table') {
    renderTable(el as unknown as DocTable, state);
  }
}

function contextualSuppressed(prev: DocParagraph | null, curr: DocParagraph): boolean {
  return !!(prev?.contextualSpacing && curr.contextualSpacing && prev.styleId && prev.styleId === curr.styleId);
}

function renderBodyElements(elements: PaginatedBodyElement[], state: RenderState): void {
  let prevPara: DocParagraph | null = null;
  let prevSpaceAfter = 0;
  for (const el of elements) {
    if (el.type === 'paragraph') {
      const para = el as unknown as DocParagraph;
      const slice = (el as PaginatedBodyElement).lineSlice;
      const suppress = contextualSuppressed(prevPara, para);
      // Collapse spaceAfter+spaceBefore like Word: use max, not sum.
      const effBefore = suppress ? 0 : para.spaceBefore;
      const overlap = Math.min(prevSpaceAfter, effBefore);
      state.y -= overlap * state.scale;
      // Continuation slices (slice.start > 0) suppress spaceBefore: the
      // earlier slice already consumed it on the previous page. Likewise
      // mid-paragraph slices (slice.end < total) suppress spaceAfter — only
      // the slice covering the FINAL line of the paragraph emits it.
      const isContinuation = !!slice && slice.start > 0;
      renderParagraph(para, state, suppress || isContinuation, slice);
      prevPara = para;
      prevSpaceAfter = para.spaceAfter;
    } else if (el.type === 'table') {
      renderTable(el as unknown as DocTable, state);
      prevPara = null;
      prevSpaceAfter = 0;
    }
  }
}

function renderParaList(paras: DocParagraph[], state: RenderState): void {
  let prevPara: DocParagraph | null = null;
  let prevSpaceAfter = 0;
  for (const para of paras) {
    const suppress = contextualSuppressed(prevPara, para);
    const effBefore = suppress ? 0 : para.spaceBefore;
    const overlap = Math.min(prevSpaceAfter, effBefore);
    state.y -= overlap * state.scale;
    renderParagraph(para, state, suppress);
    prevPara = para;
    prevSpaceAfter = para.spaceAfter;
  }
}

// ===== Paragraph rendering =====

/**
 * Map an ST_Jc (math) value to a physical alignment edge for a display
 * equation. ECMA-376 §22.1.2.88: `left`→left, `right`→right, `center` and
 * `centerGroup`→center (for a single equation centerGroup is identical to
 * center; the group-block distinction is out of scope — see spec YAGNI). The
 * math jc is an absolute position within the block, not a logical start/end, so
 * it is NOT flipped by the paragraph base direction.
 */
function mathJcToEdge(jc: string): AlignEdge {
  switch (jc) {
    case 'left':
      return 'left';
    case 'right':
      return 'right';
    case 'center':
    case 'centerGroup':
    default:
      return 'center';
  }
}

function renderParagraph(
  para: DocParagraph,
  state: RenderState,
  suppressSpaceBefore = false,
  /** When set, render only `lines[start, end)` of the laid-out paragraph,
   *  used by the paginator to split paragraphs that don't fit on one page. */
  lineSlice?: { start: number; end: number },
): void {
  const { ctx, scale, contentX, contentW, defaultColor, dryRun, fontFamilyClasses } = state;
  // Capture Y before spaceBefore — used for paragraph-relative anchor image positioning
  const paragraphStartY = state.y;

  if (!suppressSpaceBefore) state.y += para.spaceBefore * scale;

  // Register anchor floats from this paragraph (must happen after spaceBefore so that
  // paragraph-relative Y resolves against the textAreaTop, matching Word).
  registerAnchorFloats(para, state, state.y);

  // behindDoc shapes must render before text so they appear behind it.
  renderAnchorImages(para, state, paragraphStartY, 'behind');

  // If any topAndBottom float already extends past state.y, skip past it before text starts.
  state.y = skipPastTopAndBottom(state.y, state.floats);

  const textAreaTopY = state.y;

  // ECMA-376 §17.3.1.12 w:ind — the transitional left/right attributes are
  // logical start/end (Part 4 §14.11.2). In a bidi paragraph the start side is
  // the physical RIGHT, so the two indents swap physical sides here.
  const baseRtl = para.bidi === true;
  const indLeft = (baseRtl ? para.indentRight : para.indentLeft) * scale;
  const indRight = (baseRtl ? para.indentLeft : para.indentRight) * scale;
  const indFirst = para.indentFirst * scale;

  // Numbering prefix (indent is already baked into para.indentLeft / para.indentFirst)
  let numPrefix = '';
  let numTab = 0;
  if (para.numbering) {
    numPrefix = para.numbering.text + '\t';
    numTab = para.numbering.tab * scale;
  }

  const paraX = contentX + indLeft;
  const firstLineX = paraX + indFirst;
  const paraW = contentW - indLeft - indRight;

  // Collect all text segments with formatting (resolving field runs against page context)
  const segments = buildSegments(para.runs, state);
  // Word renders ruby paragraphs with consistent line spacing — every line
  // in a paragraph that carries ANY furigana snaps to the same pitch
  // multiple. Compute once at paragraph scope and share with the line loop.
  const paraHasRuby = paragraphHasRuby(para);
  const grid = paraGrid(para, state);

  if (segments.length === 0) {
    const emptyH = paragraphMarkLineHeight(para, scale, grid, paraHasRuby);
    if (para.shading && !dryRun) {
      ctx.fillStyle = `#${para.shading}`;
      ctx.fillRect(contentX + indLeft, textAreaTopY, paraW, emptyH);
    }
    state.y += emptyH;
    if (para.borders && !dryRun) {
      drawParaBorders(ctx, contentX + indLeft, textAreaTopY, paraW, emptyH, para.borders, scale);
    }
    state.y += para.spaceAfter * scale;
    renderAnchorImages(para, state, paragraphStartY);
    return;
  }

  const wrapCtx: WrapLayoutCtx | undefined = state.floats.length > 0 ? {
    startPageY: state.y,
    paraX,
    floats: state.floats,
    lineBoxH: (a, d, _h, is) => lineBoxHeight(para.lineSpacing, a, d, scale, grid, paraHasRuby, is ?? 0),
    pageH: state.pageH,
  } : undefined;

  const lines = layoutLines(ctx, segments, paraW, firstLineX - paraX, scale, para.tabStops, wrapCtx, state.fontFamilyClasses, indLeft, state.kinsoku);

  // A paragraph whose only segments are wrap-float anchors (wp:anchor) places no
  // inline content on any line, so layoutLines returns zero lines. Per ECMA-376
  // §17.3.1.29 the paragraph mark still produces one line box; §20.4.2.x removes
  // the floating object from the inline flow but does not suppress that mark.
  // Reserve the same paragraph-mark line height the literal-empty path uses, so
  // consecutive anchor-only paragraphs don't collapse onto each other. The
  // anchor floats themselves are registered and drawn by registerAnchorFloats /
  // renderAnchorImages on their own absolute-position path, so this only adds the
  // in-flow paragraph-mark advance (no double counting, no double draw).
  if (lines.length === 0) {
    const emptyH = paragraphMarkLineHeight(para, scale, grid, paraHasRuby);
    if (para.shading && !dryRun) {
      ctx.fillStyle = `#${para.shading}`;
      ctx.fillRect(contentX + indLeft, textAreaTopY, paraW, emptyH);
    }
    state.y += emptyH;
    if (para.borders && !dryRun) {
      drawParaBorders(ctx, contentX + indLeft, textAreaTopY, paraW, emptyH, para.borders, scale);
    }
    const isFinalSlice = !lineSlice || lineSlice.end >= lines.length;
    if (isFinalSlice) state.y += para.spaceAfter * scale;
    if (!lineSlice || lineSlice.start === 0) {
      renderAnchorImages(para, state, paragraphStartY);
    }
    return;
  }

  // For paragraphs that carry any ruby annotation, Word renders every line
  // at the SAME height. Per the user's note: when the section's docGrid is
  // active, Word widens the grid pitch to accommodate the tallest required
  // line (ruby + base + leading), then ALL lines in the paragraph use that
  // widened pitch — both ruby-bearing and non-ruby lines share the same
  // baseline grid, otherwise the lines drift. We mimic this by computing
  // uniformLineH = ceil(max natural / pitch) * pitch when docGrid is on,
  // else just the max natural.
  const uniformLineH = paraHasRuby
    ? snapParaLineToGrid(
        Math.max(0, ...lines.map(l => lineBoxHeight(para.lineSpacing, l.ascent, l.descent, scale, grid, true, l.intendedSingle))),
        grid,
        scale,
      )
    : 0;
  const lineHForLine = (l: typeof lines[number]): number =>
    paraHasRuby
      ? uniformLineH
      : lineBoxHeight(para.lineSpacing, l.ascent, l.descent, scale, grid, false, l.intendedSingle);

  if (para.shading && !dryRun) {
    const totalTextH = lines.reduce((s, l) => s + lineHForLine(l), 0);
    ctx.fillStyle = `#${para.shading}`;
    ctx.fillRect(contentX + indLeft, textAreaTopY, paraW, totalTextH);
  }

  // ECMA-376 §17.18.44 ST_Jc: "both" and "distribute" fully justify the line
  // by expanding inter-word spaces. The last line of a "both" paragraph is
  // traditionally left-aligned (not stretched); "distribute" also stretches
  // the last line. We count whitespace chars in trailing positions of each
  // segment and divide the slack proportionally across them.
  const isJustified =
    para.alignment === 'justify' ||
    para.alignment === 'both' ||
    para.alignment === 'distribute';
  const stretchLastLine = para.alignment === 'distribute';

  const countTrailingSpaces = (s: string) => {
    let c = 0;
    for (let i = s.length - 1; i >= 0 && s[i] === ' '; i--) c++;
    return c;
  };

  // Bidirectional text. The paragraph's base direction comes from w:bidi
  // (ECMA-376 §17.3.1.6). We engage the (exact) bidi pass only when the base is
  // RTL or the line actually contains strong-RTL characters, so pure-LTR
  // paragraphs keep their byte-identical fast path. `alignEdge` resolves
  // logical start/end against the base direction. (`baseRtl` is declared with
  // the indent swap above.)
  const paraNeedsBidi = baseRtl || segmentsHaveRtl(segments);
  const alignEdge = resolveAlignEdge(para.alignment, baseRtl);

  // Slice bounds — when the paginator split this paragraph across pages,
  // only render lines in [sliceStart, sliceEnd). The first line we paint
  // resets state.y baseline so the slice begins at the page's content top.
  const sliceStart = lineSlice ? lineSlice.start : 0;
  const sliceEnd = lineSlice ? lineSlice.end : lines.length;
  for (let li = sliceStart; li < sliceEnd; li++) {
    const line = lines[li];
    // First-line indent and numbering prefix only apply to the paragraph's
    // ORIGINAL first line, not the first line of a continuation slice.
    const firstLine = li === 0;
    // Last-line justification flips off only at the paragraph's true end —
    // mid-paragraph slices keep justifying through to the slice boundary.
    const isLastLine = li === lines.length - 1;

    // Honor wrap-computed line topY (may push past topAndBottom floats).
    if (line.topY !== undefined && line.topY > state.y) state.y = line.topY;

    // Word centers the font's natural line (ascent+descent) within the expanded
    // line box — extra space from auto/exact/atLeast goes half above and half
    // below the glyphs. Baseline = top + halfExtra + ascent. For ruby
    // paragraphs every line uses the same height (the max natural line) so
    // ruby- and non-ruby-bearing lines share a baseline grid.
    const lineH = lineHForLine(line);
    const naturalLineH = line.ascent + line.descent;
    const baseline = state.y + (lineH - naturalLineH) / 2 + line.ascent;

    // Per-line X range (may be narrower than paraW when wrapping around floats).
    const lineLeft = paraX + line.xOffset;
    const lineAvailW = line.availWidth;
    // First-line indent shifts the START edge: physical left for LTR; for RTL
    // the start is the right edge, so it narrows/widens the line's available
    // width instead of moving x (effAvailW below).
    let x = firstLine && !baseRtl ? lineLeft + indFirst : lineLeft;
    const effAvailW = baseRtl && firstLine ? lineAvailW - indFirst : lineAvailW;

    // Visual draw order. Under bidi we reorder the line's segments per UAX#9
    // (rule L2) and draw each with ctx.direction matching its resolved
    // direction; ctx.textAlign stays physical 'left' so x is always the
    // segment's left edge. The LTR fast path is untouched (visual === null).
    // Computed before justification so the stretch bookkeeping below can use
    // the same (visual) domain as the draw loop.
    const visual: LineVisualOrder | null = paraNeedsBidi
      ? computeLineVisualOrder(line.segments, baseRtl)
      : null;
    if (paraNeedsBidi) ctx.textAlign = 'left';
    const segCount = line.segments.length;
    // The segment whose trailing spaces sit at the line's PHYSICAL end (no
    // stretch there): the visually-last segment. Equals the logical last in
    // the LTR fast path.
    const lastDrawnSi = visual ? visual.order[segCount - 1] : segCount - 1;

    const lineWidth = line.segments.reduce((s, seg) => s + seg.measuredWidth, 0);
    const lineSlack = effAvailW - (x - lineLeft) - lineWidth;
    const applyJustify = isJustified && (!isLastLine || stretchLastLine);
    let alignOffset = 0;
    // ECMA-376 §22.1.2.88 `m:jc` / §22.1.2.30 `m:defJc` — a display equation's
    // justification is independent of the paragraph's text alignment. When this
    // line is exactly one display-math segment, resolve its effective math jc
    // (per-instance → document default → spec default `centerGroup`) and use it
    // for THIS line only, overriding the paragraph alignEdge.
    const onlyMathSeg =
      line.segments.length === 1 &&
      'mathNodes' in line.segments[0] &&
      (line.segments[0] as LayoutMathSeg).display
        ? (line.segments[0] as LayoutMathSeg)
        : null;
    const mathEdge = onlyMathSeg
      ? mathJcToEdge(onlyMathSeg.jc ?? state.mathDefJc ?? 'centerGroup')
      : null;
    const effEdge = mathEdge ?? alignEdge;
    if (effEdge === 'right') {
      alignOffset = lineSlack;
    } else if (effEdge === 'center') {
      alignOffset = lineSlack / 2;
    } else if (effEdge === 'justify' && baseRtl && !applyJustify) {
      // The unstretched (last) line of a justified RTL paragraph aligns to the
      // leading edge — the RIGHT margin (§17.18.44 `both`: last line is
      // start-aligned). LTR keeps alignOffset 0 as before.
      alignOffset = lineSlack;
    }
    // 'left' and stretched 'justify' keep alignOffset 0.
    x += alignOffset;

    if (firstLine && numPrefix && !dryRun) {
      const numFontSize = getDefaultFontSize(para) * scale;
      ctx.font = `${numFontSize}px sans-serif`;
      ctx.fillStyle = defaultColor;
      if (baseRtl) {
        // The RTL list marker is laid out INLINE at the line's start (right)
        // edge: its right edge sits numTab (w:hanging) to the right of the
        // text's start edge, mirroring the LTR `firstLineX - numTab` anchor,
        // and it follows the text through jc alignment (sample-8 PDF ground
        // truth: marker right edge = aligned text right edge + hanging).
        const prevAlign = ctx.textAlign;
        const prevDir = ctx.direction;
        ctx.textAlign = 'left';
        ctx.direction = 'rtl';
        const markerW = ctx.measureText(para.numbering!.text).width;
        ctx.fillText(para.numbering!.text, x + lineWidth + numTab - markerW, baseline);
        ctx.textAlign = prevAlign;
        ctx.direction = prevDir;
      } else {
        ctx.fillText(para.numbering!.text, lineLeft + indFirst - numTab, baseline);
      }
    }

    // Inter-word adjustment per whitespace char on this line. Positive slack
    // (lineWidth < availW) expands spaces to fill; negative slack (lineWidth >
    // availW, typically from canvas measuring ~1 px wider than Word) compresses
    // spaces so the final glyph lands on the right margin instead of overflowing.
    // Compression is capped so we never eat more than the natural width of a
    // space, and is only applied when the line is a candidate for justification
    // (jc=both/distribute, not the last line unless distribute).
    let extraPerSpace = 0;
    if (applyJustify) {
      let totalTrailingSpaces = 0;
      for (let si = 0; si < segCount; si++) {
        const seg = line.segments[si];
        // Trailing spaces on the visually-final segment don't stretch (they sit
        // at the physical line end) — same domain as the draw-loop skip below.
        if (si === lastDrawnSi) continue;
        if ('text' in seg) totalTrailingSpaces += countTrailingSpaces((seg as LayoutTextSeg).text);
      }
      const slack = effAvailW - (x - lineLeft) - lineWidth;
      if (totalTrailingSpaces > 0) {
        extraPerSpace = slack / totalTrailingSpaces;
        // Don't compress past zero-width spaces — limit compression to at most
        // half the widest space on the line. Estimated from default font size.
        const minExtra = -line.ascent * 0.25;
        if (extraPerSpace < minExtra) extraPerSpace = minExtra;
      }
    }

    for (let vi = 0; vi < segCount; vi++) {
      const si = visual ? visual.order[vi] : vi;
      const seg = line.segments[si];
      const isLastSeg = vi === segCount - 1;
      if (visual) ctx.direction = visual.rtl[si] ? 'rtl' : 'ltr';
      if ('isTab' in seg) {
        // Tabs render as blank space, optionally filled with a leader (TOC dots etc.).
        if (!dryRun && seg.leader && seg.leader !== 'none' && seg.measuredWidth > 1) {
          drawTabLeader(ctx, seg.leader, x, baseline, seg.measuredWidth, seg.fontSize * scale, defaultColor);
        }
        x += seg.measuredWidth;
        continue;
      }
      if ('dataUrl' in seg) {
        if (!dryRun) renderInlineImage(ctx, seg as LayoutImageSeg, x, baseline, scale, state.images);
        x += seg.measuredWidth;
        continue;
      }
      if ('mathNodes' in seg) {
        const render = mathRenders.get(seg.mathNodes);
        if (!dryRun && render) {
          const emPx = seg.fontSize * scale;
          const w = render.widthEm * emPx;
          const h = (render.ascentEm + render.descentEm) * emPx;
          const top = baseline - render.ascentEm * emPx;
          ctx.drawImage(render.img, x, top, w, h);
        }
        x += seg.measuredWidth;
        continue;
      }
      const s = seg as LayoutTextSeg;
      if (!dryRun) {
        const effSizePx = calcEffectiveFontPx(s, scale);
        const yOffset = s.vertAlign === 'super'
          ? -s.fontSize * scale * 0.35
          : s.vertAlign === 'sub'
            ? s.fontSize * scale * 0.15
            : 0;
        ctx.font = buildFont(s.bold, s.italic, effSizePx, s.fontFamily, fontFamilyClasses);

        if (s.highlight) {
          ctx.fillStyle = HIGHLIGHT_COLORS[s.highlight] ?? '#FFFF00';
          ctx.fillRect(x, baseline + yOffset - effSizePx * 0.85, s.measuredWidth, effSizePx * 1.1);
        }

        // Track-changes overlay: paint insertions / deletions in the author's
        // colour with the canonical Word markup (underline for insertions,
        // strikethrough for deletions). The author hash gives stable colours
        // for the same reviewer across pages. Disabled when
        // `showTrackChanges: false` (the "Final / No Markup" view).
        const revActive = state.showTrackChanges && !!s.revision;
        const revColor = revActive ? authorColor(s.revision!.author) : null;
        ctx.fillStyle = revColor ?? (s.color ? `#${s.color}` : defaultColor);
        ctx.fillText(s.text, x, baseline + yOffset);

        // Ruby annotation: small text centered above the base glyphs.
        if (s.ruby) {
          const rubySizePx = s.ruby.fontSizePt * scale;
          const rubyFont = buildFont(s.bold, s.italic, rubySizePx, s.fontFamily, fontFamilyClasses);
          ctx.save();
          ctx.font = rubyFont;
          const rubyW = ctx.measureText(s.ruby.text).width;
          const rubyX = x + (s.measuredWidth - rubyW) / 2;
          // Sit the ruby's baseline a small gap above the base ascent so the
          // characters don't touch. fillText baseline is at the line of the
          // characters, so subtract the ruby descent + small gap from the
          // base's ascent line to position correctly.
          const rubyBaseline = baseline + yOffset - effSizePx * 0.85 - rubySizePx * 0.1;
          ctx.fillStyle = s.color ? `#${s.color}` : defaultColor;
          ctx.fillText(s.ruby.text, rubyX, rubyBaseline);
          ctx.restore();
        }

        if (state.onTextRun && s.text) {
          state.onTextRun({
            text: s.text,
            x,
            y: state.y,
            w: s.measuredWidth,
            h: lineH,
            fontSize: effSizePx,
            font: ctx.font,
          });
        }

        const lineColor = revColor ?? (s.color ? `#${s.color}` : defaultColor);
        const lineW = Math.max(0.5, effSizePx * 0.05);
        const textW = ctx.measureText(s.text).width;

        const isInsertion = revActive && s.revision?.kind === 'insertion';
        const isDeletion = revActive && s.revision?.kind === 'deletion';

        if (s.underline || isInsertion) {
          ctx.strokeStyle = lineColor;
          ctx.lineWidth = lineW;
          const uy = baseline + yOffset + effSizePx * 0.12;
          ctx.beginPath(); ctx.moveTo(x, uy); ctx.lineTo(x + textW, uy); ctx.stroke();
        }

        if (s.strikethrough || isDeletion) {
          ctx.strokeStyle = lineColor;
          ctx.lineWidth = lineW;
          const sy = baseline + yOffset - effSizePx * 0.3;
          ctx.beginPath(); ctx.moveTo(x, sy); ctx.lineTo(x + textW, sy); ctx.stroke();
        }

        if (s.doubleStrikethrough) {
          ctx.strokeStyle = lineColor;
          ctx.lineWidth = lineW;
          const sy1 = baseline + yOffset - effSizePx * 0.35;
          const sy2 = baseline + yOffset - effSizePx * 0.22;
          ctx.beginPath(); ctx.moveTo(x, sy1); ctx.lineTo(x + textW, sy1); ctx.stroke();
          ctx.beginPath(); ctx.moveTo(x, sy2); ctx.lineTo(x + textW, sy2); ctx.stroke();
        }
      }

      x += s.measuredWidth;
      // Inter-word justification slack (applied AFTER the segment so the next
      // segment starts at a shifted baseline). Skip on the final segment —
      // trailing spaces at line end don't participate in stretching.
      if (extraPerSpace > 0 && !isLastSeg) {
        const trailing = countTrailingSpaces(s.text);
        if (trailing > 0) x += trailing * extraPerSpace;
      }
    }
    if (paraNeedsBidi) ctx.direction = 'ltr'; // reset for subsequent draws

    state.y += lineH;
  }

  if (para.borders && !dryRun) {
    const textH = state.y - textAreaTopY;
    drawParaBorders(ctx, contentX + indLeft, textAreaTopY, paraW, textH, para.borders, scale);
  }

  // spaceAfter is paragraph-level; only emit it on the slice that covers
  // the FINAL line of the paragraph (or when no slice is set at all).
  const isFinalSlice = !lineSlice || lineSlice.end >= lines.length;
  if (isFinalSlice) state.y += para.spaceAfter * scale;

  // Anchor images are absolutely positioned — draw after inline flow.
  // Skip this for continuation slices: anchor positioning is paragraph-relative
  // and the first slice already painted them.
  if (!lineSlice || lineSlice.start === 0) {
    renderAnchorImages(para, state, paragraphStartY);
  }
}

// ===== Text layout =====

interface LayoutTextSeg {
  text: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  fontSize: number;  // pt
  color: string | null;
  fontFamily: string | null;
  vertAlign: 'super' | 'sub' | null;
  measuredWidth: number;  // px (set during layout)
  smallCaps?: boolean;
  doubleStrikethrough?: boolean;
  highlight?: string | null;
  /** Ruby annotation rendered in a small font directly above this segment. */
  ruby?: { text: string; fontSizePt: number };
  /** Track-changes revision attached to this run (insertion / deletion). */
  revision?: { kind: 'insertion' | 'deletion' | string; author?: string };
  /** ECMA-376 §17.3.2.30 `<w:rtl>` — run carries right-to-left characteristics.
   *  When true the segment's text is treated as a strong-RTL embedding in the
   *  per-line bidi pass (so leading digits / neutrals resolve RTL). */
  rtl?: boolean;
  /** UAX#9 §4.3 HL1 / Word behaviour: classify this segment's European digits
   *  (U+0030–0039) as Arabic-Number (AN) in the per-line bidi pass, so a date
   *  like "28-02-2026" in an Arabic complex-script run reorders to "2026-02-28"
   *  exactly as Word renders it (ECMA-376 §17.3.2.20 w:lang w:bidi). */
  digitsAsAN?: boolean;
}

/**
 * Horizontal tab. Width is resolved during layout against paragraph tab stops
 * (or the default 36pt interval if no explicit stop is configured).
 */
interface LayoutTabSeg {
  isTab: true;
  fontSize: number;  // pt — for line-height purposes
  measuredWidth: number;
  /** tab leader to fill the gap (e.g. TOC dot leaders); set during layout. */
  leader?: TabStop['leader'];
}

interface LayoutImageSeg {
  dataUrl: string;
  widthPt: number;
  heightPt: number;
  /** true = wp:anchor: skip inline flow, draw at absolute page coords */
  anchor: boolean;
  anchorXPt: number;
  anchorYPt: number;
  anchorXFromMargin: boolean;
  anchorYFromPara: boolean;
  /** When set, pixels matching this hex color are replaced with alpha=0 before drawing. */
  colorReplaceFrom?: string;
  measuredWidth: number;
}

/** An inline OMML equation. Measured + drawn via the core math engine. */
interface LayoutMathSeg {
  mathNodes: import('@silurus/ooxml-core').MathNode[];
  display: boolean;
  fontSize: number;  // pt
  color: string | null;
  measuredWidth: number;
  /** px ascent/descent of the laid-out box at scale, cached during measurement. */
  mathAscent: number;
  mathDescent: number;
  /** ECMA-376 §22.1.2.88 `m:oMathPara/m:jc` — per-instance justification of a
   *  display equation (ST_Jc math). `undefined` for inline math; the renderer
   *  resolves the document default (`mathDefJc`, spec default `centerGroup`). */
  jc?: string;
}

/** Sentinel that forces a new line when encountered in layoutLines. */
interface LayoutLineBreak {
  lineBreak: true;
  fontSize: number;  // pt — used to set line height on empty lines
  measuredWidth: 0;
}

type LayoutSeg = LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutLineBreak | LayoutTabSeg;

interface LayoutLine {
  segments: (LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutTabSeg)[];
  height: number;  // pt — max fontSize on line (for empty-line sizing fallback)
  ascent: number;  // px — fontBoundingBoxAscent (font-metric, stable per font+size)
  descent: number; // px — fontBoundingBoxDescent
  /** px — intended single-line height (max over segments of the requested
   *  font's win line-height ratio × em), for fonts whose substituted Canvas
   *  metrics understate Word's line spacing. 0 when no segment needs it. */
  intendedSingle: number;
  /** Additional horizontal offset (px) from paraX, caused by wrap-around floats. */
  xOffset: number;
  /** Effective available width (px) for this line after float exclusion. */
  availWidth: number;
  /** When wrap context is active, the absolute canvas Y where this line begins. */
  topY?: number;
  /** Set when at least one segment on this line carries a ruby annotation —
   *  enables docGrid pitch snapping in lineBoxHeight. */
  hasRuby?: boolean;
}

/** Additional context passed to layoutLines so it can honor floats on the current page. */
interface WrapLayoutCtx {
  startPageY: number;   // absolute canvas Y where the first line should start
  paraX: number;        // absolute canvas X of the paragraph's content left edge
  floats: FloatRect[];  // floats active on the current page
  /** Per-line box-height resolver (line natural ascent+descent → total px box height). */
  lineBoxH: (ascentPx: number, descentPx: number, hasRuby?: boolean, intendedSinglePx?: number) => number;
  /** Hard cap on Y to keep layout from running past the page. */
  pageH: number;
}

/**
 * Resolve the formatting axis that actually governs a run's glyphs.
 *
 * ECMA-376 §17.3.2.30 `w:rtl` marks a run as complex-script. For such a run the
 * complex-script properties take effect — §17.3.2.4 `bCs` (bold), §17.3.2.6
 * `iCs` (italic), §17.3.2.26 `rFonts@cs` (typeface), §17.3.2.39 `szCs` (size) —
 * instead of the non-CS `b`/`i`/`rFonts@ascii`/`sz`, which apply to
 * non-complex (Latin/CJK) text. `bCs`/`iCs` are INDEPENDENT toggles: an absent
 * `bCs` does not inherit `b`'s value, so a complex-script run that carries only
 * `w:b` renders non-bold. (sample-7's date cells carry `b` without `bCs`, and
 * Word draws them at regular weight; the header carries both and is bold.)
 */
function buildSegments(runs: DocRun[], state: RenderState): LayoutSeg[] {
  const segs: LayoutSeg[] = [];
  const pushTextPiece = (
    text: string,
    base: DocxTextRun | FieldRun,
    vertAlign: 'super' | 'sub' | null,
  ) => {
    const displayText = (base.allCaps || base.smallCaps) ? text.toUpperCase() : text;
    // Ruby annotation rides with the WHOLE base text (typically 1-2 chars).
    // Splitting on word boundaries would lose the association, so attach
    // the annotation only to the first emitted segment.
    const baseRuby = (base as DocxTextRun).ruby;
    const ruby = baseRuby
      ? { text: baseRuby.text, fontSizePt: baseRuby.fontSizePt }
      : undefined;
    const revision = (base as DocxTextRun).revision;
    const r = base as DocxTextRun;
    const rtl = r.rtl === true ? true : undefined;

    // ECMA-376 §17.3.2.26 content classification. A run with `w:rtl`
    // (§17.3.2.30) or the `<w:cs/>` toggle (§17.3.2.7) applies complex-script
    // formatting to ALL of its characters; otherwise each character is routed by
    // its Unicode block (Arabic/Hebrew/... → cs; Latin/digits/CJK → ascii/hAnsi).
    // NOTE rFonts@cs (fontFamilyCs) alone is just a font SLOT and must NOT
    // force cs — e.g. sample-1's Heading1 (Latin) has cstheme + szCs=52 but
    // renders at w:sz=24; forcing cs blew its size up to 26pt.
    const forceCs = r.rtl === true || r.cs === true;

    // Complex-script (cs) formatting sources, each falling back to its Latin
    // counterpart when the cs-specific property is absent (the parser already
    // resolves szCs through the full style chain, mirroring a directly-set
    // `w:sz` per §17.3.2.18; bCs/iCs per §17.3.2.3/§17.3.2.17).
    const csFontSize = r.fontSizeCs ?? base.fontSize;
    const csFontFamily = r.fontFamilyCs ?? base.fontFamily;
    const csBold = r.boldCs ?? base.bold;
    const csItalic = r.italicCs ?? base.italic;

    // Word classifies European digits in an Arabic/Hebrew complex-script run as
    // AN (§17.3.2.20 w:lang w:bidi): use the bidi language's primary subtag when
    // present, else fall back to the run being rtl-marked.
    const digitsAsAN =
      (forceCs || r.rtl === true) && isRtlBidiLang(r.langBidi, r.rtl === true);

    let firstSeg = true;
    const emit = (word: string, cs: boolean) => {
      segs.push({
        text: word,
        bold: cs ? csBold : base.bold,
        italic: cs ? csItalic : base.italic,
        underline: base.underline,
        strikethrough: base.strikethrough,
        fontSize: cs ? csFontSize : base.fontSize,
        color: base.color,
        fontFamily: cs ? csFontFamily : base.fontFamily,
        vertAlign,
        measuredWidth: 0,
        smallCaps: base.smallCaps ?? false,
        doubleStrikethrough: base.doubleStrikethrough ?? false,
        highlight: base.highlight ?? null,
        ruby: firstSeg ? ruby : undefined,
        revision,
        rtl,
        digitsAsAN: digitsAsAN ? true : undefined,
      });
      firstSeg = false;
    };

    for (const word of splitTextForLayout(displayText)) {
      if (forceCs) {
        // When the run's digits are AN-classified, split a token into maximal
        // digit-groups and the surrounding separators so the per-line bidi pass
        // (which reorders at SEGMENT granularity) can place the groups in Word's
        // order — e.g. "28-02-2026" → segments [28][-][02][-][2026] reordered to
        // 2026-02-28. Canvas only reorders WITHIN a fillText using EN semantics,
        // so a single-segment date would otherwise stay 28-02-2026.
        if (digitsAsAN) {
          for (const slice of splitDigitGroups(word)) emit(slice, true);
        } else {
          emit(word, true);
        }
      } else {
        // Mixed Arabic+Latin word (no w:rtl / w:cs): split at script boundaries
        // so each side gets its own (cs vs Latin) size and typeface.
        for (const slice of splitByComplexScript(word)) emit(slice.text, slice.cs);
      }
    }
  };

  for (const run of runs) {
    if (run.type === 'text') {
      const t = run as unknown as DocxTextRun & { type: 'text' };
      // ECMA-376 §17.11: substitute a footnote/endnote reference marker's glyph
      // with the note's resolved sequential number. The body `*Reference` run
      // carries the id; the in-note `*Ref` placeholder carries an empty id, so
      // we fall back to the note number currently being drawn.
      const noteText =
        t.noteRef
          ? (t.noteRef.id
              ? state.noteNumbers?.get(`${t.noteRef.kind}:${t.noteRef.id}`)
              : state.currentNoteNumber)
          : undefined;
      if (t.noteRef) {
        const label = noteText != null ? String(noteText) : (t.text || '');
        if (label.length > 0) pushTextPiece(label, t, t.vertAlign ?? 'super');
        continue;
      }
      // Split on tab chars so tab alignment can be resolved during layout.
      const parts = t.text.split('\t');
      for (let i = 0; i < parts.length; i++) {
        if (parts[i].length > 0) pushTextPiece(parts[i], t, t.vertAlign);
        if (i < parts.length - 1) {
          segs.push({ isTab: true, fontSize: t.fontSize, measuredWidth: 0 });
        }
      }
    } else if (run.type === 'image') {
      const img = run as unknown as ImageRun & { type: 'image' };
      segs.push({
        dataUrl: img.dataUrl,
        widthPt: img.widthPt,
        heightPt: img.heightPt,
        anchor: img.anchor ?? false,
        anchorXPt: img.anchorXPt ?? 0,
        anchorYPt: img.anchorYPt ?? 0,
        anchorXFromMargin: img.anchorXFromMargin ?? false,
        anchorYFromPara: img.anchorYFromPara ?? false,
        colorReplaceFrom: img.colorReplaceFrom,
        measuredWidth: 0,
      });
    } else if (run.type === 'break') {
      if (run.breakType === 'line') {
        // Determine font size for the line break height from surrounding text runs
        const fontSize = findNearbyFontSize(runs, runs.indexOf(run));
        segs.push({ lineBreak: true, fontSize, measuredWidth: 0 });
      }
      // page/column breaks handled at the document level (splitPages)
    } else if (run.type === 'field') {
      const f = run as unknown as FieldRun & { type: 'field' };
      const text = resolveFieldText(f, state);
      if (text) pushTextPiece(text, f, f.vertAlign);
    } else if (run.type === 'math') {
      // The parser resolves the paragraph font size; fall back to a nearby run only
      // if it is somehow absent.
      const fontSize = run.fontSize || findNearbyFontSize(runs, runs.indexOf(run));
      segs.push({
        mathNodes: run.nodes,
        display: run.display,
        fontSize,
        color: null,
        measuredWidth: 0,
        mathAscent: 0,
        mathDescent: 0,
        jc: run.jc,
      });
    }
  }
  return segs;
}

function findNearbyFontSize(runs: DocRun[], idx: number): number {
  // Look backwards then forwards for a text or field run to get font size
  for (let i = idx - 1; i >= 0; i--) {
    const r = runs[i];
    if (r.type === 'text') return (r as unknown as DocxTextRun).fontSize;
    if (r.type === 'field') return (r as unknown as FieldRun).fontSize;
  }
  for (let i = idx + 1; i < runs.length; i++) {
    const r = runs[i];
    if (r.type === 'text') return (r as unknown as DocxTextRun).fontSize;
    if (r.type === 'field') return (r as unknown as FieldRun).fontSize;
  }
  return 10; // pt fallback
}

function resolveFieldText(f: FieldRun, state: RenderState): string {
  if (f.fieldType === 'page') return String(state.pageIndex + 1);
  if (f.fieldType === 'numPages') return String(state.totalPages);
  return f.fallbackText;
}

/** Returns true for code-points that permit line-break between adjacent characters (CJK). */
function hasCJKBreakOpportunity(text: string): boolean {
  for (let i = 0; i < text.length; ) {
    const cp = text.codePointAt(i)!;
    if (
      (cp >= 0x3000 && cp <= 0x9FFF)  ||
      (cp >= 0xF900 && cp <= 0xFAFF)  ||
      (cp >= 0xAC00 && cp <= 0xD7AF)  ||
      (cp >= 0xFF00 && cp <= 0xFFEF)
    ) return true;
    i += cp > 0xFFFF ? 2 : 1;
  }
  return false;
}

/**
 * Binary-search the longest prefix of `text` whose rendered width fits in `maxWidth`.
 * Used for CJK overflow splitting.
 */
function fitCJKPrefix(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  text: string,
  maxWidth: number,
): string {
  const chars = [...text]; // spread handles surrogate pairs
  let lo = 0, hi = chars.length;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (ctx.measureText(chars.slice(0, mid).join('')).width <= maxWidth) lo = mid;
    else hi = mid - 1;
  }
  return chars.slice(0, lo).join('');
}

/* ------------------------------------------------------------------ *
 * Japanese line-breaking (kinsoku shori / 禁則処理)
 *
 * ECMA-376 §17.15.1.58 `w:kinsoku` is a document-wide on/off toggle for
 * "East Asian typography line-breaking rules". Its default, when the
 * element is absent from settings.xml, is TRUE (the toggle is a
 * ST_OnOff whose absence Word treats as enabled for kinsoku). So a doc
 * with no <w:kinsoku> still gets Japanese line breaking — which is what
 * Word does and what users see.
 *
 * §17.15.1.59 `w:noLineBreaksAfter` / §17.15.1.60 `w:noLineBreaksBefore`
 * let a document override the character set used by the kinsoku engine
 * for a given language (`w:lang`):
 *   - noLineBreaksBefore (§.60): characters that "cannot begin a line"
 *     (行頭禁則 — line-start-forbidden).
 *   - noLineBreaksAfter  (§.59): characters that "cannot end a line"
 *     (行末禁則 — line-end-forbidden).
 * The spec states the `w:val` "specifies the set of characters" — it is
 * the COMPLETE set, so a present override REPLACES the application's
 * default set for that language (it does not extend it). When the
 * element is absent the application's own default set is used. We
 * implement replace-vs-default exactly per that wording.
 *
 * The default sets below are Word's documented Japanese kinsoku tables
 * (Tools ▸ Options ▸ Typography ▸ "Use default kinsoku rules"). They
 * coincide with JIS X 4051 §6.1 (行頭禁則文字 / 行末禁則文字). We encode
 * them as two flat string constants (data, not scattered conditionals);
 * membership is a Set lookup.
 *
 * Word applies kinsoku only to East-Asian wrapping (the per-character
 * break path). Pure-Latin word wrap is untouched: these sets contain no
 * ASCII letters/space, and the Latin path never consults them.
 * ------------------------------------------------------------------ */

/** §17.15.1.60 default 行頭禁則 — characters that may NOT begin a line.
 *  Closing brackets/quotes, mid/end punctuation, small kana, prolonged
 *  sound mark, iteration marks, and their halfwidth forms. */
const KINSOKU_DEFAULT_LINE_START_FORBIDDEN =
  // closing brackets / quotes (fullwidth)
  '”’）〕］｝〉》」』】〙〗〟｠»' +
  // mid / end punctuation (fullwidth)
  '、。，．・：；／？！‐ー゠–〜～' +
  // small kana
  'ぁぃぅぇぉっゃゅょゎゕゖ' +
  'ァィゥェォッャュョヮヵヶ' +
  'ㇰㇱㇲㇳㇴㇵㇶㇷㇸㇹㇺㇻㇼㇽㇾㇿ' +
  // iteration / sound marks
  '々〻ゝゞヽヾ゛゜' +
  // misc trailing symbols
  '％‰℃°′″' +
  // halfwidth forms (cannot start a line either)
  '｡｣､･ｰﾞﾟ' +
  '!),.:;?]}｠';

/** §17.15.1.59 default 行末禁則 — characters that may NOT end a line.
 *  Opening brackets / quotes and currency/lead symbols. */
const KINSOKU_DEFAULT_LINE_END_FORBIDDEN =
  // opening brackets / quotes (fullwidth)
  '“‘（〔［｛〈《「『【〘〖〝｟«' +
  // currency / lead symbols
  '＄￥＃￡￠' +
  // halfwidth opening forms
  '([{｟';

/** Resolved kinsoku configuration for a document.
 *  `enabled` reflects §17.15.1.58; the two sets are §.60 / §.59 (custom
 *  sets replace the defaults — see resolveKinsokuRules). */
export interface KinsokuRules {
  enabled: boolean;
  /** Code points forbidden at line START (行頭禁則). */
  lineStartForbidden: Set<number>;
  /** Code points forbidden at line END (行末禁則). */
  lineEndForbidden: Set<number>;
}

function codePointSet(text: string): Set<number> {
  const out = new Set<number>();
  for (const ch of text) out.add(ch.codePointAt(0)!);
  return out;
}

/** Build the active {@link KinsokuRules} from the document settings.
 *  - `enabled` defaults to TRUE when undefined (§17.15.1.58 default).
 *  - A non-undefined custom set REPLACES the default for that direction
 *    (§17.15.1.59 / §.60 "specifies the set of characters"). An empty
 *    string is a legitimate replacement that disables that direction.
 */
export function resolveKinsokuRules(settings?: {
  kinsoku?: boolean;
  noLineBreaksBefore?: string;
  noLineBreaksAfter?: string;
}): KinsokuRules {
  return {
    enabled: settings?.kinsoku !== false,
    lineStartForbidden: codePointSet(
      settings?.noLineBreaksBefore ?? KINSOKU_DEFAULT_LINE_START_FORBIDDEN,
    ),
    lineEndForbidden: codePointSet(
      settings?.noLineBreaksAfter ?? KINSOKU_DEFAULT_LINE_END_FORBIDDEN,
    ),
  };
}

/** The default Japanese kinsoku rules (no document overrides). */
export const DEFAULT_KINSOKU_RULES: KinsokuRules = resolveKinsokuRules();

/**
 * Adjust a CJK line-break position so it does not violate kinsoku.
 *
 * Given a line being split into `head = chars[0..splitAt)` (stays on the
 * current line) and `tail = chars[splitAt..]` (overflows to the next),
 * return the largest legal `splitAt' <= splitAt` such that:
 *   1. `tail'[0]` is not line-START-forbidden (行頭禁則 追い出し — the
 *      offending char and any preceding forbidden chars are pulled down
 *      onto the next line), AND
 *   2. `head'[last]` is not line-END-forbidden (push a dangling opener
 *      to the next line).
 *
 * Retraction is bounded: we never retract below `minSplit` (default 1,
 * so at least one code point always stays on a non-empty line and we
 * keep forward progress). If no legal split exists within that bound
 * (pathological run of forbidden chars), the original `splitAt` is
 * returned unchanged — Word likewise lets an over-long forbidden run
 * overflow rather than loop forever.
 *
 * `chars` must be an array of single code points (e.g. `[...text]`).
 */
export function kinsokuAdjustedSplit(
  chars: string[],
  splitAt: number,
  rules: KinsokuRules,
  minSplit = 1,
): number {
  if (!rules.enabled) return splitAt;
  if (splitAt <= 0 || splitAt >= chars.length) return splitAt;

  const startForbidden = (i: number): boolean =>
    i < chars.length && rules.lineStartForbidden.has(chars[i].codePointAt(0)!);
  const endForbidden = (i: number): boolean =>
    i >= 0 && rules.lineEndForbidden.has(chars[i].codePointAt(0)!);

  let s = splitAt;
  // Retract while the tail begins with a start-forbidden char OR the head
  // ends with an end-forbidden char. Each retraction moves one code point
  // from the head onto the tail (追い出し). Bounded by minSplit.
  while (s > minSplit && (startForbidden(s) || endForbidden(s - 1))) {
    s--;
  }
  // If we hit the floor and it is still illegal, no legal break exists in
  // range — fall back to the unrestricted split (never empty, never hang).
  if (s <= minSplit && (startForbidden(s) || endForbidden(s - 1))) {
    return splitAt;
  }
  return s;
}

/**
 * Split a text run into layout-segment strings.
 * Each segment is an atomic unit for word-level fitting; CJK overflow is handled in layoutLines.
 */
/** RTL primary language subtags (ISO 639) whose complex-script context makes
 *  Word classify European digits as Arabic-Number (AN). */
const RTL_PRIMARY_SUBTAGS = new Set([
  'ar', // Arabic
  'fa', // Persian
  'ur', // Urdu
  'he', 'iw', // Hebrew (iw = legacy code)
  'yi', 'ji', // Yiddish
  'ps', // Pashto
  'sd', // Sindhi
  'ug', // Uyghur
  'dv', // Divehi
  'syr', // Syriac
  'ckb', // Central Kurdish (Sorani)
]);

/**
 * Decide whether a `w:lang w:bidi` tag (§17.3.2.20) designates an RTL
 * complex-script language, so the run's European digits are classified AN
 * (Word's date ordering). The tag's primary subtag (before the first '-') is
 * matched against {@link RTL_PRIMARY_SUBTAGS}. When the tag is absent OR a
 * malformed/unknown value (e.g. the "ae-AR" seen in real-world files), fall
 * back to whether the run is explicitly rtl-marked — `w:rtl` already asserts
 * the run is complex-script RTL content.
 */
function isRtlBidiLang(langBidi: string | undefined, runIsRtl: boolean): boolean {
  if (langBidi) {
    const primary = langBidi.split('-')[0].toLowerCase();
    if (RTL_PRIMARY_SUBTAGS.has(primary)) return true;
  }
  return runIsRtl;
}

/**
 * ECMA-376 §17.3.2.26 (rFonts) content classification — does this code point
 * belong to the COMPLEX-SCRIPT (`cs`) category? Word routes such characters to
 * the run's complex-script formatting (`w:szCs` / `w:rFonts w:cs` / `w:bCs` /
 * `w:iCs`) instead of the Latin (`ascii`/`hAnsi`) formatting. The ranges are
 * the Unicode blocks Word treats as complex script: Hebrew, Arabic (incl.
 * supplements / extended / presentation forms), Syriac, Thaana, NKo,
 * Samaritan, Mandaic, and the Arabic-math / extended Plane-1 blocks. Latin,
 * digits (EN), punctuation and CJK are NOT cs.
 */
function isComplexScriptCodePoint(cp: number): boolean {
  return (
    (cp >= 0x0590 && cp <= 0x05ff) || // Hebrew
    (cp >= 0x0600 && cp <= 0x06ff) || // Arabic
    (cp >= 0x0700 && cp <= 0x074f) || // Syriac
    (cp >= 0x0750 && cp <= 0x077f) || // Arabic Supplement
    (cp >= 0x0780 && cp <= 0x07bf) || // Thaana
    (cp >= 0x07c0 && cp <= 0x07ff) || // NKo
    (cp >= 0x0800 && cp <= 0x083f) || // Samaritan
    (cp >= 0x0840 && cp <= 0x085f) || // Mandaic
    (cp >= 0x0860 && cp <= 0x08ff) || // Syriac Supp. / Arabic Extended-A/B
    (cp >= 0xfb1d && cp <= 0xfb4f) || // Hebrew presentation forms
    (cp >= 0xfb50 && cp <= 0xfdff) || // Arabic Presentation Forms-A
    (cp >= 0xfe70 && cp <= 0xfeff) || // Arabic Presentation Forms-B
    (cp >= 0x10800 && cp <= 0x10fff) || // Plane-1 RTL blocks
    (cp >= 0x1e800 && cp <= 0x1efff)    // Mende/Adlam/Arabic Math
  );
}

/**
 * Split `text` into maximal runs that are uniformly complex-script or not, per
 * §17.3.2.26 per-character classification. Returns `[{text, cs}]` in logical
 * order. Used only when a run has NO explicit `w:rtl`/`w:cs` (which would force
 * the whole run to cs); otherwise the caller treats the entire piece as cs.
 *
 * Digits / spaces / punctuation (neutral, non-cs) attach to the PRECEDING slice
 * so a number embedded in Arabic ("نص 12 نص") does not fragment the word into
 * extra segments — Word keeps such weak/neutral characters with the script they
 * border. A leading neutral run takes the first strong slice's class.
 */
function splitByComplexScript(text: string): { text: string; cs: boolean }[] {
  const out: { text: string; cs: boolean }[] = [];
  let curCs: boolean | null = null;
  let buf = '';
  for (const ch of text) {
    const cp = ch.codePointAt(0) as number;
    // Neutral (non-letter) characters do not switch the active class; they ride
    // with whatever script is currently open (or the next one if none yet).
    const isLetter = /\p{L}/u.test(ch);
    if (!isLetter) {
      buf += ch;
      continue;
    }
    const cs = isComplexScriptCodePoint(cp);
    if (curCs === null) {
      curCs = cs;
      buf += ch;
    } else if (cs === curCs) {
      buf += ch;
    } else {
      out.push({ text: buf, cs: curCs });
      curCs = cs;
      buf = ch;
    }
  }
  if (buf.length > 0) out.push({ text: buf, cs: curCs ?? false });
  return out;
}

/**
 * Split a token into maximal runs of European digits (U+0030–0039) versus
 * everything else, so a date / number in an AN-classified Arabic run can be
 * reordered group-by-group by the per-line bidi pass (which works at segment
 * granularity). "28-02-2026" → ["28","-","02","-","2026"].
 */
function splitDigitGroups(text: string): string[] {
  const out: string[] = [];
  let buf = '';
  let bufDigit: boolean | null = null;
  for (const ch of text) {
    const c = ch.charCodeAt(0);
    const isDigit = c >= 0x30 && c <= 0x39;
    if (bufDigit === null || isDigit === bufDigit) {
      buf += ch;
    } else {
      out.push(buf);
      buf = ch;
    }
    bufDigit = isDigit;
  }
  if (buf.length > 0) out.push(buf);
  return out.length ? out : [text];
}

function splitTextForLayout(text: string): string[] {
  const result: string[] = [];
  let i = 0;
  while (i < text.length) {
    let j = i;
    while (j < text.length && text[j] !== ' ') j++;
    while (j < text.length && text[j] === ' ') j++;
    if (j > i) result.push(text.slice(i, j));
    i = j;
  }
  return result.length ? result : [text];
}

function layoutLines(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  segs: LayoutSeg[],
  maxWidth: number,
  firstIndent: number,
  scale: number,
  tabStops: TabStop[] = [],
  wrapCtx?: WrapLayoutCtx,
  fontFamilyClasses: Record<string, string> = {},
  // Paragraph left-indent in px. Tab-stop positions are measured from the text
  // margin (ECMA-376 §17.3.1.37), but layout is paraX-relative, so subtract this.
  tabOriginPx: number = 0,
  // ECMA-376 §17.15.1.58–.60 Japanese line-breaking rules. Default kinsoku is
  // ON; the CJK overflow path retracts the break to a kinsoku-legal position.
  kinsoku: KinsokuRules = DEFAULT_KINSOKU_RULES,
): LayoutLine[] {
  const lines: LayoutLine[] = [];
  let currentLine: (LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutTabSeg)[] = [];
  let currentWidth = 0;
  // Sum of trailing-space widths of every text token on the current line.
  // Used for two things:
  //   1. Knuth-Plass-style shrink tolerance: a justified line may compress
  //      inter-word spaces by up to SPACE_SHRINK_RATIO, so a candidate word
  //      whose "natural" width would overflow by less than that total shrink
  //      budget is allowed to fit. This is the standard typographic approach
  //      to line-breaking and lets us absorb the ~0.1–0.3 px/glyph advance
  //      bias between Chromium canvas and Word's internal text layout,
  //      matching Word's wrap on long paragraphs.
  //   2. Trailing-space collapse at line end — the last token's trailing
  //      space disappears when no further word is added, so when deciding
  //      whether a candidate word fits we treat it as if it would become the
  //      final word (its own trailing spaces collapsible).
  let lineTotalTrailingW = 0;
  let lineHeight = 0;   // pt
  let lineAscent = 0;   // px
  let lineDescent = 0;  // px
  let lineIntendedSingle = 0; // px — max intended single-line height on the line
  let isFirst = true;
  // Effective width/offset for the current line after float exclusion.
  let lineMaxWidth = maxWidth;
  let lineXOffset = 0;
  let currentLineTopY = wrapCtx?.startPageY ?? 0;

  // Compute wrap constraints for a new line about to start. Mutates lineXOffset/lineMaxWidth/currentLineTopY.
  const startLine = (): void => {
    lineXOffset = 0;
    lineMaxWidth = maxWidth;
    if (!wrapCtx) return;
    // Probe height: the smallest plausible line height; good enough for float intersection check.
    const probeH = 10 * scale;
    // Keep pushing past any topAndBottom block we sit inside.
    for (let guard = 0; guard < 16; guard++) {
      const lineBot = currentLineTopY + probeH;
      let skip: number | null = null;
      for (const f of wrapCtx.floats) {
        if (f.mode !== 'topAndBottom') continue;
        if (lineBot > f.yTop && currentLineTopY < f.yBottom) {
          skip = skip === null ? f.yBottom : Math.max(skip, f.yBottom);
        }
      }
      if (skip === null) break;
      currentLineTopY = skip;
    }
    // Now compute horizontal constraint from square floats.
    const paraXLeft = wrapCtx.paraX;
    const paraXRight = wrapCtx.paraX + maxWidth;
    let left = paraXLeft;
    let right = paraXRight;
    const lineBot = currentLineTopY + probeH;
    for (const f of wrapCtx.floats) {
      if (f.mode !== 'square') continue;
      if (lineBot <= f.yTop || currentLineTopY >= f.yBottom) continue;
      // Decide which side text should flow on. "left"/"right" refer to the side TEXT occupies.
      const spaceLeft = f.xLeft - paraXLeft;
      const spaceRight = paraXRight - f.xRight;
      let textOnLeft: boolean;
      switch (f.side) {
        case 'left':    textOnLeft = true;  break;
        case 'right':   textOnLeft = false; break;
        case 'largest':
        case 'bothSides':
        default:        textOnLeft = spaceLeft >= spaceRight; break;
      }
      if (textOnLeft) {
        if (f.xLeft < right) right = Math.max(left, f.xLeft);
      } else {
        if (f.xRight > left) left = Math.min(right, f.xRight);
      }
    }
    const eff = Math.max(0, right - left);
    lineXOffset = Math.max(0, left - paraXLeft);
    lineMaxWidth = Math.min(maxWidth - lineXOffset, eff);
    if (lineMaxWidth < 0) lineMaxWidth = 0;
  };
  startLine();

  const availW = () => lineMaxWidth - (isFirst ? firstIndent : 0);

  // Default tab interval when no matching explicit stop exists (Word's default is 720 twips = 36pt)
  const DEFAULT_TAB_PT = 36;

  let lineHasRuby = false;
  const flush = (forceHeight?: number) => {
    const h = forceHeight !== undefined ? forceHeight : (lineHeight || 10);
    // If the line has no measured content (empty/line-break line), synthesize
    // stable ascent/descent from the effective font size so wrap/baseline math
    // stays consistent with non-empty lines.
    const hasContent = lineAscent > 0 || lineDescent > 0;
    const asc = hasContent ? lineAscent : h * scale * 0.8;
    const desc = hasContent ? lineDescent : h * scale * 0.2;
    lines.push({
      segments: currentLine,
      height: h,
      ascent: asc,
      descent: desc,
      intendedSingle: lineIntendedSingle,
      xOffset: lineXOffset,
      availWidth: lineMaxWidth,
      topY: wrapCtx ? currentLineTopY : undefined,
      hasRuby: lineHasRuby,
    });
    if (wrapCtx) {
      currentLineTopY += wrapCtx.lineBoxH(asc, desc, lineHasRuby, lineIntendedSingle);
    }
    currentLine = [];
    currentWidth = 0;
    lineTotalTrailingW = 0;
    lineHeight = 0;
    lineAscent = 0;
    lineDescent = 0;
    lineIntendedSingle = 0;
    lineHasRuby = false;
    isFirst = false;
    startLine();
  };

  const addToLine = (
    s: LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutTabSeg,
    w: number,
    h: number,
    asc: number,
    desc: number,
    trailingSpaceW: number = 0,
  ) => {
    currentLine.push(s);
    currentWidth += w;
    lineTotalTrailingW += trailingSpaceW;
    if (h > lineHeight) lineHeight = h;
    if (asc > lineAscent) lineAscent = asc;
    if (desc > lineDescent) lineDescent = desc;
    if (!('isTab' in s) && !('dataUrl' in s) && !('mathNodes' in s)) {
      const ts = s as LayoutTextSeg;
      if (ts.ruby) lineHasRuby = true;
      // Intended single-line height for fonts whose substituted Canvas metrics
      // understate Word's line spacing (font-metrics.ts). 0 for untabled fonts.
      const intended = intendedSingleLinePx(ts.fontFamily, effectiveFontPx(ts));
      if (intended > lineIntendedSingle) lineIntendedSingle = intended;
    }
  };

  const effectiveFontPx = (s: LayoutTextSeg): number => calcEffectiveFontPx(s, scale);

  const measureText = (s: LayoutTextSeg): TextMetrics => {
    ctx.font = buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses);
    return ctx.measureText(s.text);
  };

  // Width of a queued segment, for right/center tab look-ahead.
  const tabFollowWidth = (q: LayoutSeg): number => {
    if ('isTab' in q) return q.measuredWidth || 0;
    if ('dataUrl' in q) return q.widthPt * scale;
    if ('mathNodes' in q) return q.measuredWidth || 0;
    if ('lineBreak' in q) return 0;
    return measureText(q).width;
  };

  // Use an explicit queue so CJK split-tails can be re-queued
  const queue: LayoutSeg[] = [...segs];

  // A `<w:br/>` always starts a new line (§17.3.3.1) — when it is the LAST
  // content of the paragraph that new line is an EMPTY line that still
  // occupies one line height (Word reserves it; visible e.g. as extra table
  // row height). Track the trailing break so it can be flushed after the loop.
  let trailingBreakFontSize: number | null = null;

  while (queue.length > 0) {
    const seg = queue.shift()!;

    // ── Line-break sentinel ──────────────────────────────
    if ('lineBreak' in seg) {
      flush(seg.fontSize);
      trailingBreakFontSize = seg.fontSize;
      continue;
    }
    trailingBreakFontSize = null;

    // ── Tab segment ──────────────────────────────────────
    if ('isTab' in seg) {
      // Absolute position on the line measured from paraX (line origin for continuation lines)
      const absFromParaX = currentWidth + (isFirst ? firstIndent : 0);
      // Tab-stop X relative to paraX: stops are measured from the text margin, so
      // subtract the paragraph's own left indent.
      const stopXof = (t: TabStop) => t.pos * scale - tabOriginPx;
      // Find the next tab stop strictly greater than the current position
      const stop = tabStops.find((t) => stopXof(t) > absFromParaX);
      // Right/center/decimal tab: place the tab + its trailing content (up to the next
      // tab / line end) so the content ends at / centers on the stop, and commit that
      // content directly so the normal wrap check doesn't push it past the stop
      // (ECMA-376 §17.3.1.37). This is what makes TOC "heading …… page" lines work.
      if (stop && stop.alignment !== 'left' && stop.alignment !== 'bar' && stop.alignment !== 'clear') {
        const stopX = stopXof(stop);
        seg.leader = stop.leader;
        let followW = 0;
        for (const q of queue) {
          if ('isTab' in q || 'lineBreak' in q) break;
          followW += tabFollowWidth(q);
        }
        const frac = stop.alignment === 'center' ? 0.5 : 1;
        let tabW = stopX - absFromParaX - followW * frac;
        if (tabW <= 0) tabW = seg.fontSize * scale * 0.25;
        seg.measuredWidth = tabW;
        addToLine(seg, tabW, seg.fontSize, seg.fontSize * scale * 0.8, seg.fontSize * scale * 0.2);
        // Commit the trailing content onto this line without a wrap re-check.
        while (queue.length > 0) {
          const q = queue[0];
          if ('isTab' in q || 'lineBreak' in q) break;
          queue.shift();
          if ('dataUrl' in q) {
            const w = q.widthPt * scale;
            q.measuredWidth = w;
            addToLine(q, w, q.heightPt, q.heightPt * scale, 0);
          } else if ('mathNodes' in q) {
            addToLine(q, q.measuredWidth || 0, q.fontSize, q.mathAscent || 0, q.mathDescent || 0);
          } else {
            const m = measureText(q);
            q.measuredWidth = m.width;
            const asc = m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent ?? q.fontSize * scale * 0.8;
            const desc = m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent ?? q.fontSize * scale * 0.2;
            addToLine(q, m.width, q.fontSize, asc, desc);
          }
        }
        continue;
      }

      let tabWidth: number;
      if (stop) {
        tabWidth = stopXof(stop) - absFromParaX;
        seg.leader = stop.leader;
      } else {
        // Round up to the next DEFAULT_TAB_PT boundary
        const nextDefault = Math.ceil((absFromParaX + 0.01) / (DEFAULT_TAB_PT * scale)) * (DEFAULT_TAB_PT * scale);
        tabWidth = nextDefault - absFromParaX;
      }
      // Clamp to avoid negative widths; if tab would overflow the line, wrap instead
      if (tabWidth <= 0) {
        flush();
        queue.unshift(seg);
        continue;
      }
      if (currentWidth + tabWidth > availW() && currentLine.length > 0) {
        flush();
        queue.unshift(seg);
        continue;
      }
      seg.measuredWidth = tabWidth;
      addToLine(seg, tabWidth, seg.fontSize, seg.fontSize * scale * 0.8, seg.fontSize * scale * 0.2);
      continue;
    }

    // ── Image segment ────────────────────────────────────
    if ('dataUrl' in seg) {
      if (seg.anchor) { seg.measuredWidth = 0; continue; }
      const w = seg.widthPt * scale;
      const h = seg.heightPt;
      const asc = seg.heightPt * scale;
      seg.measuredWidth = w;
      if (currentLine.length > 0 && currentWidth + w > availW()) flush();
      addToLine(seg, w, h, asc, 0);
      continue;
    }

    // ── Math segment ─────────────────────────────────────
    if ('mathNodes' in seg) {
      const render = mathRenders.get(seg.mathNodes);
      if (!render) { seg.measuredWidth = 0; continue; }
      const emPx = seg.fontSize * scale;
      const w = render.widthEm * emPx;
      const asc = render.ascentEm * emPx;
      const desc = render.descentEm * emPx;
      seg.measuredWidth = w;
      // Ink extents (from the MathJax SVG viewBox) position the rasterized
      // glyph relative to the baseline when drawing.
      seg.mathAscent = asc;
      seg.mathDescent = desc;
      // …but the LINE BOX must reserve at least a normal single line for the
      // run's font size. A short equation — e.g. a lone "−" — has near-zero ink
      // height; using that as the line height would collapse the line (and the
      // table row) and pin the glyph to the very top of the cell. Floor to the
      // font's natural ascent/descent so math occupies a full line like text
      // does (tall math — fractions, big operators — keeps its larger ink box).
      const lineAsc = Math.max(asc, emPx * 0.8);
      const lineDesc = Math.max(desc, emPx * 0.2);
      if (currentLine.length > 0 && currentWidth + w > availW()) flush();
      addToLine(seg, w, seg.fontSize, lineAsc, lineDesc);
      continue;
    }

    // ── Text segment ─────────────────────────────────────
    const s = seg as LayoutTextSeg;
    const m = measureText(s);
    const w = m.width;
    // Line-height tracks the un-scaled pt font so super/sub don't shrink the line.
    const h = s.fontSize;
    // Prefer font-metric ascent/descent (stable per font+size) so baselines and
    // line boxes do not jitter based on the specific characters on each line.
    // When the document font is substituted by one with different vertical
    // metrics, rescale to the document font's design line box so the line
    // height (and thus baseline centering, row auto-heights, cell vAlign
    // §17.4.84, and pagination) match Word — see font-metrics.ts.
    const rawAsc = m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent ?? s.fontSize * scale * 0.8;
    const rawDesc = m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent ?? s.fontSize * scale * 0.2;
    const corrected = correctLineMetrics(s.fontFamily, effectiveFontPx(s), rawAsc, rawDesc);
    let asc = corrected.ascent;
    const desc = corrected.descent;
    // Ruby annotation: small text rendered above the base. Reserve ascent
    // room for the rt glyphs (ECMA-376 §17.3.3.25). The actual line spacing
    // in docGrid sections is set further down by the paragraph-wide
    // pitch snap — see the doc-grid-aware path in renderParagraph /
    // estimateParagraphHeight that takes max(perLineH) and snaps it to
    // an integer grid pitch. The reserve here just ensures the natural
    // line height EXCEEDS the docGrid pitch when ruby is present, so the
    // snap actually picks the next pitch slot. 1.5× rt-size is enough for
    // any rt size that fits in one extra pitch above the base.
    if (s.ruby) {
      asc = asc + s.ruby.fontSizePt * scale * 1.5;
    }
    // Wrap-fit check uses two standard typographic allowances:
    //   1. Trailing-space collapse: if this word becomes the last on the
    //      line, its trailing space (if any) collapses. We subtract it from
    //      the width used to test fit.
    //   2. Knuth-Plass shrink tolerance: a justified line may compress
    //      each inter-word space by up to SPACE_SHRINK_RATIO (25%) of its
    //      natural width without harming readability. This lets us absorb
    //      the canvas measureText vs Word advance-width discrepancy
    //      (~0.1–0.3 px/glyph) that would otherwise push a trailing word
    //      onto the next line. ECMA-376 doesn't prescribe a line-breaking
    //      algorithm — tolerance-based fit is standard typography (TeX,
    //      InDesign, Word) and keeps layout close to Word's output.
    const SPACE_SHRINK_RATIO = 0.25;
    const trimmed = s.text.replace(/ +$/, '');
    const trailingSpaceW = s.text.endsWith(' ')
      ? w - ctx.measureText(trimmed).width
      : 0;
    const wForFit = w - trailingSpaceW;
    const shrinkBudget = lineTotalTrailingW * SPACE_SHRINK_RATIO;

    if (currentWidth + wForFit <= availW() + shrinkBudget) {
      // Fits on current line as-is
      s.measuredWidth = w;
      addToLine(s, w, h, asc, desc, trailingSpaceW);
    } else if (hasCJKBreakOpportunity(s.text)) {
      // CJK overflow: split at the maximum prefix that fits, re-queue the tail
      const available = availW() - currentWidth;
      ctx.font = buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses);
      const rawPrefix = available > 0 ? fitCJKPrefix(ctx, s.text, available) : '';
      // Apply kinsoku to the break position: retract leftwards so the tail
      // never begins with a 行頭禁則 char and the head never ends with a
      // 行末禁則 char (ECMA-376 §17.15.1.58–.60). When the current line
      // already has content, retracting to an empty prefix is allowed — the
      // whole run moves to the next (fresh) line, which is Word's 追い出し.
      // When the line is empty we keep at least one char (minSplit=1) so we
      // never lose forward progress.
      const allChars = [...s.text];
      const rawSplit = [...rawPrefix].length;
      const minSplit = currentLine.length > 0 ? 0 : 1;
      const split = kinsokuAdjustedSplit(allChars, rawSplit, kinsoku, minSplit);
      const prefix = allChars.slice(0, split).join('');
      if (prefix.length > 0) {
        const pm = ctx.measureText(prefix);
        const headSeg: LayoutTextSeg = { ...s, text: prefix, measuredWidth: pm.width };
        addToLine(headSeg, pm.width, h, asc, desc);
        const tail = s.text.slice(prefix.length);
        if (tail) queue.unshift({ ...s, text: tail, measuredWidth: 0 });
      } else if (currentLine.length > 0) {
        // No prefix fits but line has content — flush and retry on a fresh line
        flush();
        queue.unshift(s);
      } else {
        // Empty line and not even one char fits — force-fit one char to guarantee progress
        const firstChar = [...s.text][0] ?? '';
        if (firstChar) {
          const fm = ctx.measureText(firstChar);
          const headSeg: LayoutTextSeg = { ...s, text: firstChar, measuredWidth: fm.width };
          addToLine(headSeg, fm.width, h, asc, desc);
          const tail = s.text.slice(firstChar.length);
          if (tail) queue.unshift({ ...s, text: tail, measuredWidth: 0 });
        }
      }
    } else if (currentLine.length === 0) {
      // Nothing on the line yet and no CJK break — force-fit (word wider than column)
      s.measuredWidth = w;
      addToLine(s, w, h, asc, desc);
    } else {
      // Latin word wrap: flush and put this word on the next line
      flush();
      s.measuredWidth = w;
      addToLine(s, w, h, asc, desc);
    }
  }

  if (currentLine.length > 0) flush();
  // Trailing <w:br/>: emit the empty line it opened (§17.3.3.1).
  else if (trailingBreakFontSize !== null) flush(trailingBreakFontSize);

  return lines;
}

/** Fill a tab gap with its leader characters (e.g. TOC dot leaders, ECMA-376 §17.3.1.37). */
function drawTabLeader(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  leader: NonNullable<LayoutTabSeg['leader']>,
  x: number,
  baseline: number,
  width: number,
  fontPx: number,
  color: string,
): void {
  const ch =
    leader === 'hyphen'
      ? '-'
      : leader === 'underscore' || leader === 'heavy'
        ? '_'
        : leader === 'middleDot'
          ? '·'
          : '.';
  ctx.save();
  ctx.font = `${fontPx}px serif`;
  ctx.fillStyle = color;
  const chW = ctx.measureText(ch).width;
  if (chW > 0) {
    // Dots sit on a loose grid; other leaders are drawn solid.
    const step = leader === 'dot' || leader === 'middleDot' ? chW * 1.5 : chW;
    const margin = chW * 0.5;
    // Leave a clear gap (about one dot-step) before the page number so the
    // leader doesn't run right up against it.
    const end = x + width - step - margin;
    for (let cx = x + margin; cx <= end; cx += step) {
      ctx.fillText(ch, cx, baseline);
    }
  }
  ctx.restore();
}

function renderInlineImage(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  seg: LayoutImageSeg,
  x: number,
  baseline: number,
  scale: number,
  images: Map<string, ImageBitmap>,
): void {
  // Anchor images are skipped during layout (measuredWidth=0, not added to line.segments)
  // and are drawn later by renderAnchorImages — so this function only handles inline images.
  if (seg.anchor) return;
  const bmp = images.get(imageKey(seg.dataUrl, seg.colorReplaceFrom));
  if (!bmp) return;
  const w = seg.widthPt * scale;
  const h = seg.heightPt * scale;
  ctx.drawImage(bmp, x, baseline - h, w, h);
}

/** Collect and draw anchor images with wrapMode='none' (or unspecified).
 * Wrap floats (square/topAndBottom/tight/through) are drawn by registerAnchorFloats.
 *
 * `phase` = 'behind' draws only shapes with behindDoc=true (sorted by zOrder asc);
 * `phase` = 'front' draws shapes without behindDoc + all anchor images. */
function renderAnchorImages(
  para: DocParagraph,
  state: RenderState,
  paragraphTopPx: number,
  phase: 'behind' | 'front' = 'front',
): void {
  if (state.dryRun) return;
  if (phase === 'behind') {
    const shapes = para.runs
      .filter((r): r is ShapeRun & { type: 'shape' } =>
        r.type === 'shape' && !!(r as unknown as ShapeRun).behindDoc)
      .slice()
      .sort((a, b) =>
        ((a as unknown as ShapeRun).zOrder ?? 0) - ((b as unknown as ShapeRun).zOrder ?? 0));
    for (const s of shapes) renderAnchorShape(s as unknown as ShapeRun, state, paragraphTopPx);
    return;
  }
  for (const run of para.runs) {
    if (run.type === 'shape') {
      const s = run as unknown as ShapeRun;
      if (s.behindDoc) continue;
      renderAnchorShape(s, state, paragraphTopPx);
      continue;
    }
    if (run.type !== 'image') continue;
    const img = run as unknown as ImageRun;
    if (!img.anchor) continue;
    if (isWrapFloat(img.wrapMode)) continue;  // drawn as a float
    const bmp = state.images.get(imageKey(img.dataUrl, img.colorReplaceFrom));
    if (!bmp) continue;
    const w = img.widthPt * state.scale;
    const h = img.heightPt * state.scale;

    // Resolve X: margin-relative offsets need section.marginLeft added
    const pageX = img.anchorXFromMargin
      ? (state.marginLeft + (img.anchorXPt ?? 0)) * state.scale
      : (img.anchorXPt ?? 0) * state.scale;

    // Resolve Y: paragraph-relative offsets use the paragraph's top Y in canvas px
    const pageY = img.anchorYFromPara
      ? paragraphTopPx + (img.anchorYPt ?? 0) * state.scale
      : (img.anchorYPt ?? 0) * state.scale;

    state.ctx.drawImage(bmp, pageX, pageY, w, h);
  }
}

/** Draw a wps:wsp shape via core's custGeom primitive. */
/** Resolve a shape's page X by combining the explicit `anchorXPt` offset with
 *  any `anchorXAlign` (ECMA-376 §20.4.3.1 wp:align). When align is set we
 *  position the shape inside the container indicated by `relativeFrom` (or
 *  `anchorXFromMargin` for the legacy two-state hint). When `pctPos` is set
 *  we ignore the explicit offset and place the shape at `pct` of the
 *  container's width / height (ECMA-376 §20.4.2.7 wp14:pctPosH/VOffset).
 *
 *  relativeFrom containers (ECMA-376 §20.4.3.4):
 *    - "page"          → full page rect
 *    - "margin"        → printable area between margins
 *    - "leftMargin"    → strip from x=0 to x=marginLeft
 *    - "rightMargin"   → strip from x=pageW-marginRight to x=pageW
 *    - "insideMargin"  → on odd pages = leftMargin, even = rightMargin
 *                        (we approximate as leftMargin)
 *    - "outsideMargin" → on odd pages = rightMargin, even = leftMargin
 *                        (we approximate as rightMargin)
 *    - "character"     → degrade to "margin" (no run-relative anchor data)
 *    - "topMargin"     → strip from y=0 to y=marginTop
 *    - "bottomMargin"  → strip from y=pageH-marginBottom to y=pageH
 *    - "paragraph"/"line" → relative to paragraph top (V only) */
function xContainer(
  relativeFrom: string | null | undefined,
  fromMarginHint: boolean,
  state: RenderState,
): { start: number; end: number } {
  const { scale } = state;
  const pageW = state.pageWidth * scale;
  const ml = state.marginLeft * scale;
  const mr = state.marginRight * scale;
  const rf = relativeFrom ?? (fromMarginHint ? 'margin' : 'page');
  switch (rf) {
    case 'page':          return { start: 0, end: pageW };
    case 'leftMargin':    return { start: 0, end: ml };
    case 'rightMargin':   return { start: pageW - mr, end: pageW };
    case 'insideMargin':  return { start: 0, end: ml };
    case 'outsideMargin': return { start: pageW - mr, end: pageW };
    case 'margin':
    case 'character':
    case 'column':
    default:              return { start: ml, end: pageW - mr };
  }
}

function yContainer(
  relativeFrom: string | null | undefined,
  fromParaHint: boolean,
  paragraphTopPx: number,
  state: RenderState,
): { start: number; end: number } {
  const { scale } = state;
  const mt = state.marginTop * scale;
  const mb = state.marginBottom * scale;
  const rf = relativeFrom ?? (fromParaHint ? 'paragraph' : 'page');
  switch (rf) {
    case 'page':         return { start: 0, end: state.pageH };
    case 'topMargin':    return { start: 0, end: mt };
    case 'bottomMargin': return { start: state.pageH - mb, end: state.pageH };
    case 'paragraph':
    case 'line':         return { start: paragraphTopPx, end: state.pageH };
    case 'margin':
    default:             return { start: mt, end: state.pageH - mb };
  }
}

/** Resolve the page X for an anchor or anchor-group child. `offsetPx` carries
 *  the shape's offset (within the group for wgp children, 0 for standalone
 *  anchors). `alignWidthPx` is the width used when aligning — the GROUP's
 *  width for wgp children, the shape's own width for standalone anchors. */
function resolveAnchorX(
  align: string | null | undefined,
  fromMargin: boolean,
  offsetPt: number,
  widthPx: number,
  state: RenderState,
  relativeFrom?: string | null,
  pctPos?: number | null,
  alignWidthPt?: number | null,
): number {
  const { scale } = state;
  const c = xContainer(relativeFrom, fromMargin, state);
  const offsetPx = offsetPt * scale;
  if (pctPos != null) {
    return c.start + (c.end - c.start) * pctPos + offsetPx;
  }
  if (!align) {
    return c.start + offsetPx;
  }
  const containerW = c.end - c.start;
  const aw = alignWidthPt != null ? alignWidthPt * scale : widthPx;
  switch (align) {
    case 'center': return c.start + (containerW - aw) / 2 + offsetPx;
    case 'right':
    case 'outside': return c.end - aw + offsetPx;
    case 'inside':
    case 'left':
    default:        return c.start + offsetPx;
  }
}

function resolveAnchorY(
  align: string | null | undefined,
  fromPara: boolean,
  offsetPt: number,
  heightPx: number,
  paragraphTopPx: number,
  state: RenderState,
  relativeFrom?: string | null,
  pctPos?: number | null,
  alignHeightPt?: number | null,
): number {
  const { scale } = state;
  const c = yContainer(relativeFrom, fromPara, paragraphTopPx, state);
  const offsetPx = offsetPt * scale;
  if (pctPos != null) {
    return c.start + (c.end - c.start) * pctPos + offsetPx;
  }
  if (!align) {
    return c.start + offsetPx;
  }
  const containerH = c.end - c.start;
  const ah = alignHeightPt != null ? alignHeightPt * scale : heightPx;
  switch (align) {
    case 'center': return c.start + (containerH - ah) / 2 + offsetPx;
    case 'bottom': return c.end - ah + offsetPx;
    case 'top':
    default:       return c.start + offsetPx;
  }
}

function renderAnchorShape(shape: ShapeRun, state: RenderState, paragraphTopPx: number): void {
  const { ctx, scale } = state;
  // ECMA-376 §20.4.2.18: when wp14:sizeRelH/sizeRelV is present it overrides
  // the static wp:extent for that axis. The size is `relativeFrom` container
  // size × pct.
  //
  // For a wgp group with sizeRelH, the parent group resizes and every child
  // shape scales proportionally — so a grouped child's effective width is
  // `original_width × (new_group_w / old_group_w)`, and its within-group
  // offset (carried by anchorXPt) scales by the same ratio. Standalone
  // shapes simply take `container × pct` as their width.
  let w = shape.widthPt * scale;
  let h = shape.heightPt * scale;
  let offsetXPt = shape.anchorXPt;
  let offsetYPt = shape.anchorYPt;
  let alignWidthPt = shape.groupWidthPt ?? null;
  let alignHeightPt = shape.groupHeightPt ?? null;
  if (shape.widthPct != null) {
    const c = xContainer(shape.widthRelativeFrom, false, state);
    const newSizePt = ((c.end - c.start) * shape.widthPct) / scale;
    if (shape.groupWidthPt != null && shape.groupWidthPt > 0) {
      const ratio = newSizePt / shape.groupWidthPt;
      w = shape.widthPt * scale * ratio;
      offsetXPt = shape.anchorXPt * ratio;
    } else {
      w = newSizePt * scale;
    }
    alignWidthPt = newSizePt;
  }
  if (shape.heightPct != null) {
    const c = yContainer(shape.heightRelativeFrom, false, paragraphTopPx, state);
    const newSizePt = ((c.end - c.start) * shape.heightPct) / scale;
    if (shape.groupHeightPt != null && shape.groupHeightPt > 0) {
      const ratio = newSizePt / shape.groupHeightPt;
      h = shape.heightPt * scale * ratio;
      offsetYPt = shape.anchorYPt * ratio;
    } else {
      h = newSizePt * scale;
    }
    alignHeightPt = newSizePt;
  }
  // Line/connector presets (ECMA-376 §20.1.9.18) are valid with a degenerate
  // bounding box — a horizontal line has h==0, a vertical line w==0. Stroking
  // such a path still draws a visible segment, so only bail when there is truly
  // nothing to draw (both dimensions zero) or an inverted box (negative).
  const preset = shape.presetGeometry?.toLowerCase() ?? '';
  const isLineGeom =
    preset === 'line' ||
    preset.startsWith('straightconnector') ||
    preset.startsWith('bentconnector') ||
    preset.startsWith('curvedconnector');
  if (w < 0 || h < 0) return;
  if (isLineGeom ? w === 0 && h === 0 : w === 0 || h === 0) return;
  const x = resolveAnchorX(
    shape.anchorXAlign, shape.anchorXFromMargin, offsetXPt, w, state,
    shape.anchorXRelativeFrom, shape.pctPosH, alignWidthPt,
  );
  const y = resolveAnchorY(
    shape.anchorYAlign, shape.anchorYFromPara, offsetYPt, h, paragraphTopPx, state,
    shape.anchorYRelativeFrom, shape.pctPosV, alignHeightPt,
  );

  const rot = shape.rotation ?? 0;
  ctx.save();
  if (rot !== 0) {
    ctx.translate(x + w / 2, y + h / 2);
    ctx.rotate((rot * Math.PI) / 180);
    ctx.translate(-(x + w / 2), -(y + h / 2));
  }
  ctx.beginPath();
  if (shape.presetGeometry) {
    // OOXML <a:prstGeom> — delegate to core's buildShapePath which has the
    // full preset shape catalog (rect / ellipse / triangles / arrows /
    // callouts / ribbons / flowchart / …) shared with the pptx renderer.
    const adj = shape.adjValues ?? [];
    buildShapePath(
      ctx as CanvasRenderingContext2D,
      shape.presetGeometry,
      x, y, w, h,
      adj[0] ?? null,
      adj[1] ?? null,
      adj[2] ?? null,
      adj[3] ?? null,
    );
  } else {
    buildCustomPath(ctx as CanvasRenderingContext2D, shape.subpaths, x, y, w, h);
  }
  const fillStyle = resolveFill(shape.fill, ctx as CanvasRenderingContext2D, x, y, w, h);
  if (fillStyle) {
    ctx.fillStyle = fillStyle;
    ctx.fill();
  }
  if (shape.stroke && (shape.strokeWidth ?? 0) > 0) {
    ctx.strokeStyle = hexToRgba(shape.stroke);
    ctx.lineWidth = Math.max(0.5, (shape.strokeWidth ?? 0) * scale);
    ctx.stroke();
  }
  ctx.restore();

  // Body text inside the shape (wps:txbx). Drawn AFTER fill/stroke so text
  // sits on top of the panel. Rotation is intentionally not applied to body
  // text — the cover-template usage we care about uses anchor-only text.
  if (shape.textBlocks && shape.textBlocks.length > 0) {
    renderShapeText(shape, x, y, w, h, ctx as CanvasRenderingContext2D, scale, state.fontFamilyClasses);
  }
}

/** Render a shape's body text inside its bounding box, honoring lIns/tIns/
 *  rIns/bIns and the wps:bodyPr @anchor (t / ctr / b). Alignment within each
 *  line is read from the per-block paragraph alignment. */
function renderShapeText(
  shape: ShapeRun,
  x: number, y: number, w: number, h: number,
  ctx: CanvasRenderingContext2D,
  scale: number,
  fontFamilyClasses: Record<string, string> = {},
): void {
  const blocks = shape.textBlocks ?? [];
  const lIns = (shape.textInsetL ?? 0) * scale;
  const tIns = (shape.textInsetT ?? 0) * scale;
  const rIns = (shape.textInsetR ?? 0) * scale;
  const bIns = (shape.textInsetB ?? 0) * scale;
  const innerX = x + lIns;
  const innerW = Math.max(0, w - lIns - rIns);
  const innerY = y + tIns;
  const innerH = Math.max(0, h - tIns - bIns);

  // First pass: measure each block's natural height (one line at fontSize px)
  const lineHeights = blocks.map((b) => b.fontSizePt * scale * 1.2);
  const totalH = lineHeights.reduce((s, lh) => s + lh, 0);

  const anchor = shape.textAnchor ?? 't';
  let cursorY: number;
  if (anchor === 'b') {
    cursorY = innerY + Math.max(0, innerH - totalH);
  } else if (anchor === 'ctr') {
    cursorY = innerY + Math.max(0, (innerH - totalH) / 2);
  } else {
    cursorY = innerY;
  }

  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    const fontPx = block.fontSizePt * scale;
    ctx.font = buildFont(block.bold ?? false, block.italic ?? false, fontPx, block.fontFamily ?? null, fontFamilyClasses);
    ctx.fillStyle = block.color ? `#${block.color}` : '#000000';
    // The whole block is drawn in one fillText, so Canvas applies UAX#9 over
    // the full string; we only set the base direction (for neutral resolution)
    // and resolve alignment. No explicit shape-paragraph rtl flag exists, so
    // derive the base direction from the content (first-strong).
    const baseRtl = resolveBaseDirection(undefined, block.text) === 'rtl';
    // Shape text draws one line per block with no inter-word stretching, so
    // 'distribute' keeps its pre-bidi approximation (centered) rather than the
    // justify edge resolveAlignEdge reports for paragraphs.
    const edge = block.alignment === 'distribute'
      ? 'center'
      : resolveAlignEdge(block.alignment, baseRtl);
    ctx.textAlign = 'left';
    ctx.direction = baseRtl ? 'rtl' : 'ltr';
    const m = ctx.measureText(block.text);
    let tx = innerX;
    if (edge === 'center') {
      tx = innerX + Math.max(0, (innerW - m.width) / 2);
    } else if (edge === 'right') {
      tx = innerX + Math.max(0, innerW - m.width);
    }
    // Baseline = cursorY + ascent (approx 0.85 of font size for default fonts).
    const baseline = cursorY + fontPx * 0.85;
    ctx.fillText(block.text, tx, baseline);
    cursorY += lineHeights[i];
  }
  ctx.direction = 'ltr'; // reset for subsequent draws
}

function isWrapFloat(mode?: string): boolean {
  return mode === 'square' || mode === 'topAndBottom' || mode === 'tight' || mode === 'through';
}

/** Register floats from a paragraph's anchor images and draw the image bitmap immediately. */
function registerAnchorFloats(para: DocParagraph, state: RenderState, paragraphAnchorY: number): void {
  for (const run of para.runs) {
    if (run.type !== 'image') continue;
    const img = run as unknown as ImageRun;
    if (!img.anchor) continue;
    if (!isWrapFloat(img.wrapMode)) continue;

    const mode: 'square' | 'topAndBottom' =
      img.wrapMode === 'topAndBottom' ? 'topAndBottom' : 'square';

    const scale = state.scale;
    const w = img.widthPt * scale;
    const h = img.heightPt * scale;
    const pageX = img.anchorXFromMargin
      ? (state.marginLeft + (img.anchorXPt ?? 0)) * scale
      : (img.anchorXPt ?? 0) * scale;
    const pageY = img.anchorYFromPara
      ? paragraphAnchorY + (img.anchorYPt ?? 0) * scale
      : (img.anchorYPt ?? 0) * scale;
    const dt = (img.distTop    ?? 0) * scale;
    const db = (img.distBottom ?? 0) * scale;
    const dl = (img.distLeft   ?? 0) * scale;
    const dr = (img.distRight  ?? 0) * scale;

    const key = imageKey(img.dataUrl, img.colorReplaceFrom);
    const rect: FloatRect = {
      mode,
      imageKey: key,
      imageX: pageX,
      imageY: pageY,
      imageW: w,
      imageH: h,
      xLeft: pageX - dl,
      xRight: pageX + w + dr,
      yTop: pageY - dt,
      yBottom: pageY + h + db,
      side: img.wrapSide ?? 'bothSides',
      drawn: false,
    };
    state.floats.push(rect);

    if (!state.dryRun) {
      const bmp = state.images.get(key);
      if (bmp) state.ctx.drawImage(bmp, rect.imageX, rect.imageY, rect.imageW, rect.imageH);
      rect.drawn = true;
    }
  }
}

/** If y is inside a topAndBottom float, return the float bottom; otherwise return y. */
function skipPastTopAndBottom(y: number, floats: FloatRect[]): number {
  for (let guard = 0; guard < 16; guard++) {
    let next = y;
    for (const f of floats) {
      if (f.mode !== 'topAndBottom') continue;
      if (y >= f.yTop && y < f.yBottom) next = Math.max(next, f.yBottom);
    }
    if (next === y) return y;
    y = next;
  }
  return y;
}

// ===== Table rendering =====

function renderTable(table: DocTable, state: RenderState): void {
  const { ctx, scale, contentX, contentW, dryRun } = state;

  // Resolve column widths in pt (autofit by preferred widths, or fixed grid),
  // already scaled to fit the available content width, then convert to px.
  const colWidths = resolveColumnWidths(table, contentW / scale, state).map((w) => w * scale);

  // Horizontal table alignment on the page (w:tblPr/w:jc).
  const tableW = colWidths.reduce((s, w) => s + w, 0);
  const tableX =
    table.jc === 'center'
      ? contentX + Math.max(0, (contentW - tableW) / 2)
      : table.jc === 'right'
        ? contentX + Math.max(0, contentW - tableW)
        : contentX;

  const rowHeights: number[] = [];
  for (const row of table.rows) {
    const rowH = calculateRowHeight(row, table, colWidths, scale, state);
    rowHeights.push(rowH);
  }

  // ECMA-376 §17.4.85: extend each vMerge span's last row when the restart
  // cell's content exceeds the sum of its rows. calculateRowHeight excluded
  // restart cells from per-row height to keep the FIRST row from absorbing
  // all the content of a tall merged cell.
  for (let ri = 0; ri < table.rows.length; ri++) {
    let ci = 0;
    for (const cell of table.rows[ri].cells) {
      const span = Math.min(cell.colSpan, colWidths.length - ci);
      if (cell.vMerge === true) {
        const cellW = colWidths.slice(ci, ci + span).reduce((s, w) => s + w, 0);
        const contentH = measureRestartCellContentHeight(cell, table, cellW, scale, state);
        const endRi = findMergeEndRow(table, ri, ci);
        let spanH = 0;
        for (let rj = ri; rj <= endRi; rj++) spanH += rowHeights[rj];
        if (spanH < contentH) {
          rowHeights[endRi] += contentH - spanH;
        }
      }
      ci += span;
    }
  }

  // ECMA-376 §17.4.1 `<w:bidiVisual>`: lay the grid columns right-to-left, so
  // logical column 0 sits at the table's RIGHT edge and indices advance
  // leftward. We mirror by POSITION arithmetic (not canvas transform): a cell
  // spanning [ci, ci+span) gets physical left x = tableX + tableW − (offset of
  // its right grid edge). Cell borders are mirrored too — a cell's logical
  // left/right border specs swap physical sides (its "start" edge is on the
  // right). gridSpan still consumes the same logical columns; only the mapping
  // from logical column offset to a physical x flips.
  const mirror = table.bidiVisual === true;

  let y = state.y;
  for (let ri = 0; ri < table.rows.length; ri++) {
    const row = table.rows[ri];
    const rowH = rowHeights[ri];
    let x = tableX;
    let ci = 0;

    for (const cell of row.cells) {
      const span = Math.min(cell.colSpan, colWidths.length - ci);
      const cellW = colWidths.slice(ci, ci + span).reduce((s, w) => s + w, 0);
      // Physical left edge of this cell. LTR: cumulative from the left (`x`).
      // bidiVisual: place so logical column 0 is rightmost — the cell's left
      // edge is the table's right edge minus the offset of its trailing grid
      // line (sum of widths up to and including this span).
      const leadX = mirror ? tableX + tableW - (x - tableX) - cellW : x;

      if (cell.vMerge === false) {
        // continue cell — content is rendered by its restart partner.
      } else {
        // ECMA-376 §17.4.85: a vMerge=restart cell visually occupies the full
        // merged span; use the sum of row heights for its render box.
        let drawH = rowH;
        if (cell.vMerge === true) {
          const endRi = findMergeEndRow(table, ri, ci);
          drawH = 0;
          for (let rj = ri; rj <= endRi; rj++) drawH += rowHeights[rj];
        }
        // ECMA-376 §17.4.81: an exact row height is honored verbatim and
        // content taller than the row is clipped to the row box (Word clips;
        // we would otherwise overflow into neighboring rows). A vMerge=restart
        // cell spans multiple rows, so it is never governed by a single row's
        // exact height — only single-row cells clip.
        const clipExact = row.rowHeightRule === 'exact' && cell.vMerge !== true;
        if (!dryRun) renderCell(cell, table, leadX, y, cellW, drawH, state, mirror, clipExact);
        else measureCellContent(cell, table, cellW, scale, state);
      }

      x += cellW;
      ci += span;
    }

    y += rowH;
  }

  state.y = y;
}

function calculateRowHeight(
  row: DocTableRow,
  table: DocTable,
  colWidths: number[],
  scale: number,
  state: RenderState,
): number {
  // ECMA-376 §17.4.80:
  //   exact   — honor w:trHeight verbatim, clip overflow.
  //   atLeast / auto (default) — w:trHeight is a lower bound that content
  //                               can exceed. In practice Word treats the
  //                               default `auto` like `atLeast` when a value
  //                               is present: it preserves the saved layout
  //                               height even though the spec text describes
  //                               auto as content-driven. Resume / cover
  //                               templates rely on this — their row heights
  //                               (trHeight=1872, 2448, 1152 …) shape the
  //                               whole page composition.
  if (row.rowHeight != null && row.rowHeightRule === 'exact') return row.rowHeight * scale;
  const minH = row.rowHeight != null ? row.rowHeight * scale : 10 * scale;

  let maxH = minH;
  let ci = 0;
  for (const cell of row.cells) {
    const span = Math.min(cell.colSpan, colWidths.length - ci);
    const cellW = colWidths.slice(ci, ci + span).reduce((s, w) => s + w, 0);
    const cm = effCellMargins(cell, table);
    const contentW = cellW - (cm.left + cm.right) * scale;

    // ECMA-376 §17.4.85 (w:vMerge): a vMerge=restart cell's content occupies
    // the entire merged span (this row + following rows whose same column
    // carries vMerge=continue). Including its content height in THIS row's
    // height would inflate the first row of the span and push subsequent
    // content downward — Word distributes the merged-cell content across the
    // full span instead. We exclude such cells here; the calling code applies
    // a second pass to extend the span's last row when the merged sum is
    // shorter than the restart cell's content.
    if (cell.vMerge === true) {
      ci += span;
      continue;
    }
    // vMerge=false (continue) cells contain no rendered content.
    if (cell.vMerge === false) {
      ci += span;
      continue;
    }

    let h = (cm.top + cm.bottom) * scale;
    for (const ce of cell.content) {
      h += measureCellElementHeight(state, ce, contentW, scale);
    }
    if (h > maxH) maxH = h;
    ci += span;
  }
  return maxH;
}

/** Measure a vMerge=restart cell's full content height including cell margins.
 *  Used by the pass-2 merge-span extension in renderTable / estimateTableHeight. */
function measureRestartCellContentHeight(
  cell: DocTableCell,
  table: DocTable,
  cellW: number,
  scale: number,
  state: RenderState,
): number {
  const cm = effCellMargins(cell, table);
  const contentW = cellW - (cm.left + cm.right) * scale;
  let h = (cm.top + cm.bottom) * scale;
  for (const ce of cell.content) {
    h += measureCellElementHeight(state, ce, contentW, scale);
  }
  return h;
}

function measureParaHeight(
  state: RenderState,
  para: DocParagraph,
  maxWidth: number,
  scale: number,
): number {
  const segs = buildSegments(para.runs, state);
  const paraHasRuby = paragraphHasRuby(para);
  const grid = paraGrid(para, state);
  if (segs.length === 0) {
    const fs = getDefaultFontSize(para);
    const { asc, desc } = emptyLineNaturalPx(fs, scale);
    return lineBoxHeight(para.lineSpacing, asc, desc, scale, grid, paraHasRuby, emptyIntendedSinglePx(para, scale));
  }
  const lines = layoutLines(state.ctx, segs, maxWidth, 0, scale, para.tabStops, undefined, state.fontFamilyClasses, 0, state.kinsoku);
  return lines.reduce((s, l) => s + lineBoxHeight(para.lineSpacing, l.ascent, l.descent, scale, grid, paraHasRuby, l.intendedSingle), 0);
}

/** Effective cell margins (pt). Per-cell `<w:tcMar>` overrides (ECMA-376
 *  §17.4.42) take precedence per edge over the table-level `<w:tblCellMar>`
 *  default (§17.4.41). A résumé template, for example, gives one cell a larger
 *  top margin to add space above its content. */
function effCellMargins(
  cell: DocTableCell,
  table: DocTable,
): { top: number; bottom: number; left: number; right: number } {
  return {
    top: cell.marginTop ?? table.cellMarginTop,
    bottom: cell.marginBottom ?? table.cellMarginBottom,
    left: cell.marginLeft ?? table.cellMarginLeft,
    right: cell.marginRight ?? table.cellMarginRight,
  };
}

function measureCellContent(
  cell: DocTableCell,
  table: DocTable,
  cellW: number,
  scale: number,
  state: RenderState,
): void {
  const cm = effCellMargins(cell, table);
  const ml = cm.left * scale;
  const mr = cm.right * scale;
  const innerW = cellW - ml - mr;
  for (const ce of cell.content) {
    measureCellElementHeight(state, ce, innerW, scale);
  }
}

/** Measure a cell-level element (paragraph or nested table) at the rendering
 *  scale. Returns total occupied height including paragraph spacing. */
function measureCellElementHeight(
  state: RenderState,
  ce: CellElement,
  innerWPx: number,
  scale: number,
): number {
  if (ce.type === 'paragraph') {
    const para = ce as unknown as DocParagraph;
    return measureParaHeight(state, para, innerWPx, scale)
      + (para.spaceBefore + para.spaceAfter) * scale;
  }
  // Nested table — estimateTableHeight works in pt; convert to px.
  const tbl = ce as unknown as DocTable;
  return estimateTableHeight(state, tbl, innerWPx / scale) * scale;
}

function renderCell(
  cell: DocTableCell,
  table: DocTable,
  x: number,
  y: number,
  w: number,
  h: number,
  state: RenderState,
  mirror = false,
  clipExact = false,
): void {
  const { ctx, scale } = state;

  if (cell.background) {
    ctx.fillStyle = `#${cell.background}`;
    ctx.fillRect(x, y, w, h);
  }

  drawCellBorders(ctx, x, y, w, h, cell.borders, table.borders, scale, mirror);

  const cm = effCellMargins(cell, table);
  const mt = cm.top * scale;
  const mb = cm.bottom * scale;
  const ml = cm.left * scale;
  const mr = cm.right * scale;

  // ECMA-376 §17.6.5 defines w:docGrid as a section-level constraint on
  // Cell paragraphs inherit the section's docGrid, but their line-spacing
  // rule comes from the table style's pPr (see parse_table + StyleMap's
  // `resolve_para` with a `table_style_id`). "Table Grid" sets line=240
  // (M=1.0), so with docGrid a cell line box is `max(natural, pitch × 1.0)`
  // = pitch (~18pt), matching Word's observed in-cell baseline advance
  // on demo/sample-1 page 3.
  const cellState: RenderState = {
    ...state,
    contentX: x + ml,
    contentW: w - ml - mr,
    y: y + mt,
  };

  if (cell.vAlign === 'center' || cell.vAlign === 'bottom') {
    // ECMA-376 §17.4.7 requires every <w:tc> to end with a <w:p>. When a cell's
    // visible content is a nested table, Word emits a trailing empty paragraph
    // purely as that syntactic anchor. Including it in the centering content
    // height would balloon contentH ≈ rowHeight and pin the visible block to
    // the top of the cell — matching neither Word nor LibreOffice's
    // rendering of resume "bar chart" cells. Skip a single trailing empty
    // paragraph after a non-paragraph block.
    const visibleContent = trimTrailingStructuralMarker(cell.content);
    let contentH = visibleContent.reduce(
      (s, ce) => s + measureCellElementHeight(cellState, ce, w - ml - mr, scale), 0);
    // Word collapses the LAST paragraph's space-after against the cell's bottom
    // content boundary when vertically aligning. That trailing spacing produces
    // no ink (nothing follows it inside the cell), so including it in the
    // centered block height lifts the visible text above true center. Word
    // centers the line box alone: a header cell whose paragraph carries an 8 pt
    // space-after still centers the ~16.8 pt line box, not 24.8 pt. This mirrors
    // how block space-after collapses at a container edge (ECMA-376 §17.3.1.33
    // describes spacing between paragraphs, not at the frame boundary). The
    // render path already adds this space-after after the final line where it
    // has no visual effect, so trimming it here only fixes the measurement.
    // Spacing BETWEEN two paragraphs inside the cell is left intact.
    const lastEl = visibleContent[visibleContent.length - 1];
    if (lastEl && lastEl.type === 'paragraph') {
      contentH -= (lastEl as unknown as DocParagraph).spaceAfter * scale;
    }
    if (cell.vAlign === 'center') cellState.y = y + (h - contentH) / 2;
    else cellState.y = y + h - contentH - mb;
  }

  if (clipExact) {
    // ECMA-376 §17.4.81: clip content to the exact row box so taller content
    // does not bleed into adjacent rows (Word's behavior for hRule="exact").
    ctx.save();
    ctx.beginPath();
    ctx.rect(x, y, w, h);
    ctx.clip();
    renderCellContent(cell.content, cellState);
    ctx.restore();
  } else {
    renderCellContent(cell.content, cellState);
  }
}

/** Drop a trailing empty paragraph that follows a non-paragraph block (nested
 *  table). ECMA-376 §17.4.7 requires every cell to end with a paragraph; when
 *  the visible content is a nested table, Word's emitted trailing <w:p/> is a
 *  structural anchor with no visible role. Returns the original array if no
 *  such pattern matches. */
function trimTrailingStructuralMarker(content: CellElement[]): CellElement[] {
  if (content.length < 2) return content;
  const last = content[content.length - 1];
  const prev = content[content.length - 2];
  if (last.type !== 'paragraph' || prev.type === 'paragraph') return content;
  const lastPara = last as unknown as DocParagraph;
  if (lastPara.runs.length > 0) return content;
  return content.slice(0, -1);
}

/** Render a cell's interleaved paragraphs and nested tables in document order.
 *  Mirrors renderBodyElements but without page-break handling (cells never
 *  contain page breaks in our model). */
function renderCellContent(content: CellElement[], state: RenderState): void {
  let prevPara: DocParagraph | null = null;
  let prevSpaceAfter = 0;
  for (const ce of content) {
    if (ce.type === 'paragraph') {
      const para = ce as unknown as DocParagraph;
      const suppress = contextualSuppressed(prevPara, para);
      const effBefore = suppress ? 0 : para.spaceBefore;
      const overlap = Math.min(prevSpaceAfter, effBefore);
      state.y -= overlap * state.scale;
      renderParagraph(para, state, suppress);
      prevPara = para;
      prevSpaceAfter = para.spaceAfter;
    } else if (ce.type === 'table') {
      renderTable(ce as unknown as DocTable, state);
      prevPara = null;
      prevSpaceAfter = 0;
    }
  }
}

function drawCellBorders(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
  cell: CellBorders,
  table: TableBorders,
  scale: number,
  mirror = false,
): void {
  const top = cell.top ?? table.top;
  const bottom = cell.bottom ?? table.bottom;
  // ECMA-376 §17.4.1: under bidiVisual the columns are visually reversed, so a
  // cell's logical left (start) border is drawn on its physical right edge and
  // vice versa. Borders are owned by the cell, so swap which spec each side uses.
  const left = mirror ? (cell.right ?? table.right) : (cell.left ?? table.left);
  const right = mirror ? (cell.left ?? table.left) : (cell.right ?? table.right);

  if (top && top.style !== 'none') drawBorderLine(ctx, x, y, x + w, y, top, scale);
  if (bottom && bottom.style !== 'none') drawBorderLine(ctx, x, y + h, x + w, y + h, bottom, scale);
  if (left && left.style !== 'none') drawBorderLine(ctx, x, y, x, y + h, left, scale);
  if (right && right.style !== 'none') drawBorderLine(ctx, x + w, y, x + w, y + h, right, scale);
}

function drawBorderLine(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  x1: number, y1: number, x2: number, y2: number,
  spec: BorderSpec,
  scale: number,
): void {
  ctx.save();
  ctx.strokeStyle = spec.color ? `#${spec.color}` : '#000000';
  ctx.lineWidth = Math.max(0.5, spec.width * scale);
  ctx.beginPath();
  ctx.moveTo(x1, y1);
  ctx.lineTo(x2, y2);
  ctx.stroke();
  ctx.restore();
}

function drawParaBorders(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
  borders: ParagraphBorders,
  scale: number,
): void {
  const drawEdge = (edge: ParaBorderEdge | null, x1: number, y1: number, x2: number, y2: number) => {
    if (!edge || edge.style === 'none') return;
    const spec: BorderSpec = { width: edge.width, color: edge.color, style: edge.style };
    drawBorderLine(ctx, x1, y1, x2, y2, spec, scale);
  };
  const sp = (edge: ParaBorderEdge | null) => (edge?.space ?? 0) * scale;
  drawEdge(borders.top,    x, y - sp(borders.top),         x + w, y - sp(borders.top));
  drawEdge(borders.bottom, x, y + h + sp(borders.bottom),  x + w, y + h + sp(borders.bottom));
  drawEdge(borders.left,   x - sp(borders.left), y,        x - sp(borders.left), y + h);
  drawEdge(borders.right,  x + w + sp(borders.right), y,   x + w + sp(borders.right), y + h);
}

// ===== Utilities =====

function calcEffectiveFontPx(s: LayoutTextSeg, scale: number): number {
  let size = s.fontSize * scale;
  if (s.smallCaps) size *= 0.8;
  if (s.vertAlign) size *= 0.65;
  return size;
}

function buildFont(
  bold: boolean,
  italic: boolean,
  sizePx: number,
  family: string | null,
  fontFamilyClasses: Record<string, string> = {},
): string {
  const w = bold ? 'bold' : 'normal';
  const s = italic ? 'italic' : 'normal';
  const f = normalizeFontFamily(family, fontFamilyClasses);
  return `${s} ${w} ${sizePx}px ${f}`;
}

/** Arabic-script faces that hosts rarely ship; we substitute them with Noto
 *  Naskh/Sans Arabic web fonts (see DOCX_GOOGLE_FONTS in document.ts — this
 *  list MUST mirror the Arabic entries there). A run whose font is one of these
 *  contains BOTH Arabic and Latin/digit glyphs that Word renders from the same
 *  single face, so the fallback chain must keep both scripts stylistically
 *  consistent (Arabic substitute first, serif Latin companion before the sans
 *  generics) rather than letting Latin/digits leak to a CJK sans face. */
const ARABIC_SUBSTITUTE_FONTS = new Set([
  'sakkal majalla',
  'traditional arabic',
  'simplified arabic',
  'arabic typesetting',
  'univers next arabic',
  'noto naskh arabic',
  'noto sans arabic',
]);

/** Naskh-style traditional Arabic faces ship a serif Latin companion; the
 *  geometric/modern ones pair with a sans Latin. Drives whether an Arabic-font
 *  run's Latin+digits route to Noto Naskh Arabic (serif-like) or Noto Sans
 *  Arabic, and which Latin serif/sans companion follows. */
const NASKH_SERIF_ARABIC_FONTS = new Set([
  'sakkal majalla',
  'traditional arabic',
  'simplified arabic',
  'arabic typesetting',
  'noto naskh arabic',
]);

function isArabicSubstituteFont(family: string): boolean {
  return ARABIC_SUBSTITUTE_FONTS.has(family.toLowerCase());
}

/** Quote each family for a CSS font-family list. */
function quoteAll(names: readonly string[]): string {
  return names.map((n) => `"${n}"`).join(', ');
}

/** Generic Arabic web-font fallbacks (loaded when `useGoogleFonts` is on). */
const ARABIC_TAIL_SANS = ['Noto Naskh Arabic', 'Noto Sans Arabic'] as const;

/**
 * Sans fallback TAIL (everything after the requested face) for a Latin/CJK run.
 *
 * - `cjk`: the document's CJK language inferred from the font name, or `null`
 *   for a plain Latin face — in which case the existing Japanese system-font
 *   companions (Hiragino Sans / Meiryo) lead, preserving the long-standing JP
 *   default. For a non-JP CJK language the matching Noto CJK leads so shared
 *   Han glyphs take that language's shapes (see core/fonts/scripts.ts).
 *
 * Order: [CJK companions] → Arabic → non-CJK scripts (Hebrew/Thai/Devanagari,
 * Cyrillic via Noto Sans) → `sans-serif`. The non-CJK scripts have no Han
 * collision so their position is immaterial; they sit before the generic so
 * the browser's per-glyph fallback can reach them.
 */
function sansTail(cjk: ReturnType<typeof classifyCjkFont>): string {
  const cjkPart =
    cjk && cjk !== 'jp'
      ? cjkFallbackChain(cjk, 'sans')
      : // JP / Latin default: keep the historical system-font hints first, then
        // the Noto CJK siblings so a CJK glyph still resolves on hosts lacking
        // the system faces.
        ['Noto Sans JP', 'Hiragino Sans', 'Meiryo', ...cjkFallbackChain('jp', 'sans').slice(1)];
  return `${quoteAll([...cjkPart, ...ARABIC_TAIL_SANS, ...NON_CJK_SANS_FALLBACKS])}, sans-serif`;
}

/** Serif counterpart of {@link sansTail}. */
function serifTail(cjk: ReturnType<typeof classifyCjkFont>): string {
  const cjkPart =
    cjk && cjk !== 'jp'
      ? cjkFallbackChain(cjk, 'serif')
      : // JP / Latin default: historical mincho system hints, then Noto serif
        // CJK siblings.
        [
          'Yu Mincho', 'YuMincho', 'Hiragino Mincho ProN', 'MS Mincho',
          'Noto Serif JP', ...cjkFallbackChain('jp', 'serif').slice(1),
        ];
  return `${quoteAll([...cjkPart, ...ARABIC_TAIL_SANS, ...NON_CJK_SERIF_FALLBACKS])}, serif`;
}

/** Resolve a requested font-family name to a CSS font-family string with
 *  appropriate fallback chain.
 *
 *  Classification priority:
 *  1. `fontFamilyClasses` map (from `word/fontTable.xml` §17.8.3.10):
 *     - "roman"      → serif
 *     - "swiss"      → sans-serif
 *     - "modern"     → monospace
 *     - "script"/"decorative" → sans-serif fallback
 *     - "auto" / absent       → fall through to step 2
 *  2. Name-pattern matching (fallback for fonts absent from fontTable, or
 *     where fontTable says "auto"). Retained as a safety net for theme fonts
 *     and system fonts that OOXML docs do not list in fontTable.xml.
 */
export function normalizeFontFamily(
  family: string | null,
  fontFamilyClasses: Record<string, string> = {},
): string {
  if (!family) return sansTail(null);

  const escape = (s: string) => s.replace(/"/g, '\\"');
  const head = `"${escape(family)}"`;
  const lower = family.toLowerCase();

  // CJK language inferred from the font name (null for plain Latin faces). For a
  // non-JP CJK language the matching Noto CJK leads the fallback tail so shared
  // Han glyphs render with that language's shapes; see core/fonts/scripts.ts.
  const cjk = classifyCjkFont(family);

  // 0) Arabic-script faces substituted by Noto Naskh/Sans Arabic. A single
  //    Sakkal Majalla / Traditional Arabic run carries Arabic glyphs AND
  //    Latin letters/digits; Word draws both from that one face. The browser
  //    resolves each glyph against the chain in order, so the Arabic substitute
  //    MUST come first — otherwise the Latin/digit glyphs are grabbed by the
  //    first chain member that has them (e.g. the CJK "Noto Sans JP"), and
  //    Latin/digits render in a different, sans face than the Arabic. Keeping
  //    the Arabic substitute first makes Arabic+Latin+digits all resolve from
  //    one family, matching Word's single-face rendering.
  //
  //    Latin companion: traditional Naskh faces (Sakkal Majalla, Traditional
  //    Arabic, …) ship a SERIF Latin companion — Word's PDF export of sample-7
  //    renders the Latin "first leader name" with bracketed serifs and the
  //    "2026" digits as serif figures. Noto Naskh Arabic, our substitute, also
  //    ships a serif Latin face (verified: its Latin glyphs carry bracketed
  //    serifs and closely match the PDF), so placing it first gives Latin+digits
  //    a serif look consistent with the Arabic — matching Word. "Noto Serif" is
  //    kept as a serif safety net for the rare case Noto Naskh Arabic is
  //    unavailable, so Latin still falls to a serif rather than the CJK sans.
  //    Geometric Arabic faces (Univers Next Arabic) pair with a sans Latin.
  if (isArabicSubstituteFont(family)) {
    if (NASKH_SERIF_ARABIC_FONTS.has(lower)) {
      return `${head}, "Noto Naskh Arabic", "Noto Sans Arabic", "Noto Serif", "Noto Sans JP", "Hiragino Sans", serif`;
    }
    return `${head}, "Noto Sans Arabic", "Noto Naskh Arabic", "Noto Sans JP", "Hiragino Sans", sans-serif`;
  }

  // A CJK face's stroke style (song/ming = serif, gothic/hei = sans) decides
  // which Noto CJK variant leads its tail. Names verified against the Office
  // default font set: song (宋/SimSun/Batang), ming (明/MingLiU/PMingLiU),
  // kai (楷/KaiTi/標楷體), fangsong (仿宋) and any *Mincho are serif; the rest
  // (hei/黑, gothic/ゴシック, YaHei, JhengHei, Malgun, Gulim, Dotum, 角ゴ) are sans.
  const isCjkSerif =
    cjk != null &&
    (/song|sung|simsun|nsimsun|batang|gungsuh|mincho|mingliu|pmingliu|ming\s*liu|fang\s*song|fangsong|kai\s*ti|kaiti|stsong|stkaiti|stfangsong|stzhongsong|simkai|simfang|新細明|細明|宋体|明朝|楷体|楷體|仿宋|標楷|游明朝|ＭＳ 明朝/.test(
      lower,
    ) ||
      /新細明體|細明體|宋体|明朝|楷体|楷體|仿宋|標楷體|游明朝|ＭＳ 明朝/.test(family));

  // 1) Authoritative classification from word/fontTable.xml §17.8.3.10.
  const tableClass = fontFamilyClasses[family];
  if (tableClass && tableClass !== 'auto') {
    switch (tableClass) {
      case 'roman':
        return `${head}, ${serifTail(cjk)}`;
      case 'swiss':
        return `${head}, ${sansTail(cjk)}`;
      case 'modern':
        return `${head}, "Courier New", monospace`;
      default:
        // script / decorative — fall through to name-pattern matching
        break;
    }
  }

  // 2) Name-pattern fallback for fonts absent from fontTable or classified "auto".
  const isSerif =
    isCjkSerif ||
    family.includes('明朝') ||
    family.includes('明朝体') ||
    /\bmincho\b/i.test(family) ||
    /\bmin\s*cho\b/i.test(family) ||
    family.includes('ＭＳ 明朝') ||
    family.includes('MS Mincho') ||
    family.includes('Yu Mincho') ||
    family.includes('游明朝') ||
    family.includes('Hiragino Mincho') ||
    family.includes('ヒラギノ明朝') ||
    family.includes('Cambria') ||
    family.includes('Caladea') ||
    family.includes('Times') ||
    family.includes('Georgia') ||
    family.includes('Bodoni') ||
    family.includes('Garamond') ||
    family.includes('Playfair') ||
    family.includes('Source Serif') ||
    family.includes('Noto Serif');

  if (isSerif) {
    return `${head}, ${serifTail(cjk)}`;
  }

  // Japanese system-font hints (only meaningful for JP / Latin faces; a non-JP
  // CJK face skips these so its matching Noto CJK leads the tail).
  if (cjk == null || cjk === 'jp') {
    if (lower.includes('meiryo') || family.includes('メイリオ')) {
      return `${head}, "Meiryo UI", "Meiryo", ${sansTail(cjk)}`;
    }
    if (family.includes('游ゴシック') || /\byu\s*gothic\b/i.test(family) || lower.includes('yugothic')) {
      return `${head}, "Yu Gothic", "YuGothic", ${sansTail(cjk)}`;
    }
    if (lower.includes('ipa')) {
      return `${head}, "IPAexGothic", ${sansTail(cjk)}`;
    }
    if (lower.includes('segoe')) {
      return `${head}, "Segoe UI", ${quoteAll([...ARABIC_TAIL_SANS, ...NON_CJK_SANS_FALLBACKS])}, sans-serif`;
    }
  }
  return `${head}, ${sansTail(cjk)}`;
}

function getDefaultFontSize(para: DocParagraph): number {
  for (const run of para.runs) {
    if (run.type === 'text') {
      return (run as unknown as DocxTextRun).fontSize;
    }
    if (run.type === 'field') {
      return (run as unknown as FieldRun).fontSize;
    }
  }
  if (typeof para.defaultFontSize === 'number') return para.defaultFontSize;
  return 10; // pt fallback
}

/** First text/field run's font family — used to size empty paragraphs whose
 *  intended font (e.g. Meiryo) has a larger win line height than the fallback.
 *  Empty paragraphs (no runs) fall back to the paragraph's style-resolved
 *  default font so e.g. an empty Meiryo cell that forms a résumé "bar" reserves
 *  Meiryo's tall line box rather than the generic fallback's. */
function getDefaultFontFamily(para: DocParagraph): string | null {
  for (const run of para.runs) {
    if (run.type === 'text') return (run as unknown as DocxTextRun).fontFamily;
    if (run.type === 'field') return (run as unknown as FieldRun).fontFamily;
  }
  return para.defaultFontFamily ?? null;
}

/** Intended single-line height (px) for an empty paragraph, from its default
 *  font's win line-height ratio. 0 when the font is not in the metrics table. */
function emptyIntendedSinglePx(para: DocParagraph, scale: number): number {
  return intendedSingleLinePx(getDefaultFontFamily(para), getDefaultFontSize(para) * scale);
}

/** Document-grid context passed to line-box computation.  When the section's
 *  `w:docGrid` is "lines"/"linesAndChars" with a positive pitch (ECMA-376
 *  §17.6.5), auto line spacing multiplies against the grid pitch instead of
 *  the font's natural line height. Without this, a 56-pt heading with
 *  lineRule="auto" value=4.33 would claim 56×1.25×4.33 ≈ 303pt of vertical
 *  space; with this, it claims max(natural, 18pt × 4.33) ≈ 78pt — matching
 *  Word's rendering on grids typical of Japanese/Chinese templates. */
interface DocGridCtx {
  /** "default" | "lines" | "linesAndChars" | "snapToChars" */
  type: string | null | undefined;
  /** Grid pitch in pt (already converted from twips in the parser). */
  linePitchPt: number | null | undefined;
}

function isGridLineRule(ctx: DocGridCtx | undefined): boolean {
  if (!ctx || !ctx.linePitchPt || ctx.linePitchPt <= 0) return false;
  return ctx.type === 'lines' || ctx.type === 'linesAndChars';
}

/**
 * Compute the total line-box height in px from a line's natural font metrics
 * (fontBoundingBoxAscent + fontBoundingBoxDescent) per ECMA-376 §17.3.1.33.
 *
 *   auto    → natural × value ("single" = 1 natural line, "double" = 2).
 *             When docGrid type=lines|linesAndChars is active, the
 *             multiplier applies against the grid pitch instead, with a
 *             floor of the natural line height.
 *   exact   → value in pt, converted to px (ignores font and grid).
 *   atLeast → max(natural, value in pt × scale).
 *   null    → natural, or grid pitch if the section defines one.
 *
 * Exported for unit tests only — not part of the package API (not
 * re-exported from index.ts).
 */
export function lineBoxHeight(
  ls: LineSpacing | null,
  ascentPx: number,
  descentPx: number,
  scale: number,
  grid?: DocGridCtx,
  hasRuby?: boolean,
  intendedSinglePx = 0,
): number {
  const glyphNatural = ascentPx + descentPx;
  // For `auto`/single spacing the multiplier applies to the intended font's
  // design line height (ECMA-376 §17.3.1.33). When the document's font is
  // substituted, the Canvas glyph extent (`glyphNatural`) understates that —
  // see font-metrics.ts. `base` restores the intended single-line height so
  // line spacing matches Word, while never dropping below the substituted
  // glyph extent (so glyphs are not clipped). Grid-snapped lines are governed
  // by the grid pitch instead, so the metric correction stays out of them.
  const natural = Math.max(glyphNatural, intendedSinglePx);
  const hasGrid = isGridLineRule(grid);
  const pitchPx = hasGrid ? grid!.linePitchPt! * scale : 0;
  // Per ECMA-376 §17.6.5, a paragraph whose `line` attribute is NOT
  // explicitly set — it only inherits from docDefault — snaps to one grid
  // pitch per text line in docGrid sections, regardless of the inherited
  // multiplier. Paragraphs that do set `line` on their pPr or a named style
  // multiply against the pitch as usual. This is what makes Word render
  // ESSAY (9 pt, no explicit line) at ~1 pitch (~18 pt) while a 1.33×
  // body paragraph with line="320" renders at pitch × 1.33 = ~24 pt.
  //
  // Note on ruby paragraphs: an earlier version of this function snapped
  // ruby lines up to the next integer grid pitch on the theory that Word
  // reserves whole grid slots for them. Empirical comparison against Word
  // PNG exports of sample-5 (13.5pt base + 8pt rt on pitch=18) showed
  // Word does NOT snap to integer pitches — the actual line height lands
  // around 2.3× the pitch, which is what `natural` produces directly when
  // the rt reserve in buildSegments is set generously enough. So the
  // ruby-aware logic now lives entirely in the segment-level ascent
  // reservation, and lineBoxHeight stays format-agnostic.
  const inheritedOnly = ls !== null && ls.explicit !== true;
  if (!ls) {
    // No explicit spacing → single line. Use the intended single-line height
    // (`natural`) off-grid; on-grid, snap to the pitch with the glyph extent
    // as the overflow floor (the grid, not the font metric, governs height).
    return hasGrid ? Math.max(glyphNatural, pitchPx) : natural;
  }
  // A zero/negative `w:line` is degenerate input whose behavior ECMA-376
  // §17.3.1.33 does not define (read literally, an `exact` line of 0 would
  // collapse the line box to no height; some generators emit
  // `<w:spacing w:line="0" w:lineRule="exact"/>` on table cells, e.g. sample-7).
  // Word's native model has no such state: per the [MS-DOC] LSPD structure,
  // "exact" spacing is encoded as a negative dyaLine ("the line spacing, in
  // twips, is exactly 0x10000 minus dyaLine", so an exact 0 is unrepresentable)
  // and a non-negative dyaLine in twips mode is "dyaLine or the number of twips
  // necessary for single spacing, whichever value is greater" — i.e. a stored 0
  // resolves to exactly single spacing. Word's PDF export of sample-7 confirms
  // (those rows render at normal single-line height). Match that: treat
  // exact/auto line <= 0 as single spacing. (LSPD's max() rule is the twips
  // mode; applying the same fallback to a degenerate auto multiplier <= 0 is
  // the analogous non-collapsing reading.)
  if ((ls.rule === 'exact' || ls.rule === 'auto') && ls.value <= 0) {
    return hasGrid ? Math.max(glyphNatural, pitchPx) : natural;
  }
  if (ls.rule === 'auto') {
    if (hasGrid) {
      if (inheritedOnly) return Math.max(glyphNatural, pitchPx);
      return Math.max(glyphNatural, pitchPx * ls.value);
    }
    return natural * ls.value;
  }
  if (ls.rule === 'exact') return ls.value * scale;
  if (ls.rule === 'atLeast') return Math.max(natural, ls.value * scale);
  return natural;
}

/** Natural single-line height in px for an empty paragraph (no rendered text). */
function emptyLineNaturalPx(fontSizePt: number, scale: number): { asc: number; desc: number } {
  return { asc: fontSizePt * scale * 0.8, desc: fontSizePt * scale * 0.2 };
}

/**
 * Height (px) of the paragraph-mark line box for a paragraph that places no
 * inline content on any line. Per ECMA-376 §17.3.1.29 the paragraph mark always
 * produces one line box even when the paragraph has no inline runs; floating
 * objects (§20.4.2.x `wp:anchor`) are removed from the inline flow but never
 * suppress that paragraph-mark line. This is the height used both by the
 * literal empty-paragraph path and by paragraphs whose only segments are
 * wrap-float anchors (which `layoutLines` skips, yielding zero lines).
 */
function paragraphMarkLineHeight(
  para: DocParagraph,
  scale: number,
  grid: DocGridCtx | undefined,
  paraHasRuby: boolean,
): number {
  const fs = getDefaultFontSize(para);
  const { asc, desc } = emptyLineNaturalPx(fs, scale);
  return lineBoxHeight(para.lineSpacing, asc, desc, scale, grid, paraHasRuby, emptyIntendedSinglePx(para, scale));
}
