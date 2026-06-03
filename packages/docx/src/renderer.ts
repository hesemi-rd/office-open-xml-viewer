import type {
  Document, BodyElement, PaginatedBodyElement, DocParagraph, DocTable, DocTableRow, DocTableCell, CellElement,
  DocRun, TextRun, ImageRun, ShapeRun, FieldRun, HeaderFooter, LineSpacing, BorderSpec, TableBorders, CellBorders,
  TabStop, ParagraphBorders, ParaBorderEdge, SectionProps,
} from './types';
import {
  buildCustomPath,
  buildShapePath,
  hexToRgba,
  resolveFill,
  mathToMathML,
  loadMathJax,
  mathMLToSvg,
  recolorSvg,
} from '@silurus/ooxml-core';
import type { MathNode } from '@silurus/ooxml-core';
import { intendedSingleLinePx } from './font-metrics.js';

const HIGHLIGHT_COLORS: Record<string, string> = {
  yellow: '#FFFF00', cyan: '#00FFFF', green: '#00FF00', magenta: '#FF00FF',
  blue: '#0000FF', red: '#FF0000', darkBlue: '#000080', darkCyan: '#008080',
  darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000',
  darkYellow: '#808000', darkGray: '#808080', lightGray: '#C0C0C0',
  black: '#000000', white: '#FFFFFF',
};

// 1pt = 96/72 CSS px at screen
const PT_TO_PX = 96 / 72;

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
export async function prepareMathRuns(body: BodyElement[]): Promise<void> {
  const runs = collectMathRuns(body);
  if (runs.length === 0) return;
  await loadMathJax();
  for (const r of runs) {
    if (mathRenders.has(r.nodes)) continue;
    try {
      const out = await mathMLToSvg(mathToMathML(r.nodes, r.display));
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
  /** Callback for building a transparent text selection overlay. */
  onTextRun?: (run: DocxTextRunInfo) => void;
  /** When false, runs tagged with a `revision` render without the
   *  track-changes overlay (no author colour, no underline/strikethrough). */
  showTrackChanges: boolean;
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

function collectImagePairs(doc: Document): ImagePair[] {
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

async function preloadImages(doc: Document): Promise<Map<string, ImageBitmap>> {
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
  doc: Document,
  canvas: HTMLCanvasElement | OffscreenCanvas,
  pageIndex: number,
  opts: RenderDocumentOptions = {},
): Promise<void> {
  const sec = doc.section;
  const dpr = opts.dpr ?? devicePixelRatio ?? 1;
  const cssWidth = opts.width ?? sec.pageWidth * PT_TO_PX;
  const scale = cssWidth / sec.pageWidth;  // px per pt
  const cssHeight = sec.pageHeight * scale;

  canvas.width = Math.round(cssWidth * dpr);
  canvas.height = Math.round(cssHeight * dpr);

  if (canvas instanceof HTMLCanvasElement) {
    canvas.style.width = `${cssWidth}px`;
    canvas.style.height = `${cssHeight}px`;
    if (!canvas.style.display) canvas.style.display = 'block';
  }

  const ctx = canvas.getContext('2d') as CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
  ctx.scale(dpr, dpr);

  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, cssWidth, cssHeight);

  const pages = opts.prebuiltPages ?? computePages(doc.body, sec, ctx, doc.fontFamilyClasses ?? {});
  const totalPages = Math.max(opts.totalPages ?? pages.length, pages.length);
  const elements = pages[pageIndex] ?? pages[0] ?? [];

  const images = await preloadImages(doc);

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
    onTextRun: opts.onTextRun,
    showTrackChanges: opts.showTrackChanges ?? true,
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
}

/**
 * Split body into pages, honoring explicit page breaks AND measuring content
 * overflow for automatic pagination. All measurements are done in pt (scale=1).
 */
export function computePages(
  body: BodyElement[],
  section: SectionProps,
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  fontFamilyClasses: Record<string, string> = {},
): PaginatedBodyElement[][] {
  const contentH = section.pageHeight - section.marginTop - section.marginBottom;
  const contentW = section.pageWidth - section.marginLeft - section.marginRight;
  const measureState = buildMeasureState(ctx, section, fontFamilyClasses);

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
  const newPage = () => {
    if (pages[pages.length - 1].length > 0) {
      pages.push([]);
      y = 0;
      prevPara = null;
      prevSpaceAfter = 0;
      measureState.y = section.marginTop;
      measureState.floats = [];
    }
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
      // ECMA-376 §17.18.79 ST_SectionMark: oddPage / evenPage breaks pad
      // with a blank page when the new section would otherwise start on the
      // wrong parity. pages.length here is the next page's 1-based index.
      if (el.parity === 'odd' && pages.length % 2 === 0) {
        pages.push([]);
      } else if (el.parity === 'even' && pages.length % 2 === 1) {
        pages.push([]);
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
      // Break if this paragraph alone doesn't fit, OR if keepNext is set and
      // placing it would leave no room for the next block on the same page.
      // Per Word's layout behavior: `spaceAfter` is trailing whitespace that
      // can legally overflow the bottom of the page — only content + spaceBefore
      // must fit. This is what lets a closing paragraph with a large
      // `w:spacing/@w:after` land flush against the bottom margin.
      const needNext = para.keepNext ? estimateNextBlockHeight(i + 1) : 0;
      const fitHeight = h - para.spaceAfter;
      const needed = fitHeight + needNext;
      if (y > 0 && y + needed > contentH) {
        newPage();
      }

      // ECMA-376 doesn't say "paragraphs must fit on one page" — Word
      // splits long paragraphs at line boundaries. If the paragraph is
      // taller than the remaining content area, walk the laid-out lines
      // and emit one PaginatedBodyElement slice per page.
      const remainingH = contentH - y;
      if (h > remainingH && h > contentH * 0.5) {
        const placed = splitParagraphAcrossPages(
          measureState, para, contentW, suppressBefore, section.marginLeft,
          y, contentH, pages,
          () => { newPage(); },
        );
        // After splitting, `y` is the bottom of the last slice on the
        // current page (continues for the LAST slice; intermediate slices
        // filled their pages exactly, so newPage was called between them).
        y = placed.endY;
        measureState.y = section.marginTop + placed.endY;
      } else {
        pages[pages.length - 1].push(el as PaginatedBodyElement);
        y += h;
        measureState.y += h;
      }
      prevPara = para;
      prevSpaceAfter = para.spaceAfter;
    } else if (el.type === 'table') {
      const tbl = el as unknown as DocTable;
      const h = estimateTableHeight(measureState, tbl, contentW);
      if (y + h > contentH) newPage();
      pages[pages.length - 1].push(el);
      y += h;
      measureState.y += h;
      prevPara = null;
    }
  }
  return pages;
}

function buildMeasureState(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  section: SectionProps,
  fontFamilyClasses: Record<string, string> = {},
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
  let textH: number;
  if (segs.length === 0) {
    const fs = getDefaultFontSize(para);
    const { asc, desc } = emptyLineNaturalPx(fs, 1);
    textH = lineBoxHeight(para.lineSpacing, asc, desc, 1, state.docGrid, paraHasRuby, emptyIntendedSinglePx(para, 1));
  } else {
    // When anchor-image floats are active on the current page the paragraph
    // wraps around them, adding lines compared to a full-width layout. Use
    // the same WrapLayoutCtx the renderer uses so estimate and render agree.
    const wrapCtx: WrapLayoutCtx | undefined = state.floats.length > 0 ? {
      startPageY: state.y,
      paraX: paraXPt,
      floats: state.floats,
      lineBoxH: (a, d, _h, is) => lineBoxHeight(para.lineSpacing, a, d, 1, state.docGrid, paraHasRuby, is ?? 0),
      pageH: state.pageH,
    } : undefined;
    const lines = layoutLines(state.ctx, segs, paraW, para.indentFirst, 1, para.tabStops, wrapCtx, state.fontFamilyClasses, indLeft);
    if (paraHasRuby) {
      // Word uses the same line height for every line in a ruby paragraph,
      // snapped to an integer docGrid pitch.
      const uniform = snapParaLineToGrid(
        Math.max(0, ...lines.map(l => lineBoxHeight(para.lineSpacing, l.ascent, l.descent, 1, state.docGrid, true, l.intendedSingle))),
        state.docGrid,
        1,
      );
      textH = uniform * lines.length;
    } else {
      textH = lines.reduce((s, l) => s + lineBoxHeight(para.lineSpacing, l.ascent, l.descent, 1, state.docGrid, false, l.intendedSingle), 0);
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
    if (run.type === 'text' && (run as unknown as TextRun).ruby) return true;
  }
  return false;
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
  const lines = layoutLines(measureState.ctx, segs, paraW, para.indentFirst, 1, para.tabStops, wrapCtx, measureState.fontFamilyClasses, indLeft);
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

function estimateTableHeight(state: RenderState, table: DocTable, contentWPt: number): number {
  const totalColW = table.colWidths.reduce((s, w) => s + w, 0);
  const colScale = totalColW > contentWPt ? contentWPt / totalColW : 1;
  const colWidths = table.colWidths.map((w) => w * colScale);

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

  return rowHs.reduce((s, x) => s + x, 0);
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
  set: Document['headers'],
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

  const indLeft = para.indentLeft * scale;
  const indRight = para.indentRight * scale;
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

  if (segments.length === 0) {
    const fontSizePt = getDefaultFontSize(para);
    const { asc, desc } = emptyLineNaturalPx(fontSizePt, scale);
    const emptyH = lineBoxHeight(para.lineSpacing, asc, desc, scale, state.docGrid, paraHasRuby, emptyIntendedSinglePx(para, scale));
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
    lineBoxH: (a, d, _h, is) => lineBoxHeight(para.lineSpacing, a, d, scale, state.docGrid, paraHasRuby, is ?? 0),
    pageH: state.pageH,
  } : undefined;

  const lines = layoutLines(ctx, segments, paraW, firstLineX - paraX, scale, para.tabStops, wrapCtx, state.fontFamilyClasses, indLeft * scale);

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
        Math.max(0, ...lines.map(l => lineBoxHeight(para.lineSpacing, l.ascent, l.descent, scale, state.docGrid, true, l.intendedSingle))),
        state.docGrid,
        scale,
      )
    : 0;
  const lineHForLine = (l: typeof lines[number]): number =>
    paraHasRuby
      ? uniformLineH
      : lineBoxHeight(para.lineSpacing, l.ascent, l.descent, scale, state.docGrid, false, l.intendedSingle);

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
    let x = firstLine ? lineLeft + indFirst : lineLeft;

    if (firstLine && numPrefix && !dryRun) {
      const numFontSize = getDefaultFontSize(para) * scale;
      ctx.font = `${numFontSize}px sans-serif`;
      ctx.fillStyle = defaultColor;
      ctx.fillText(para.numbering!.text, x - numTab, baseline);
    }

    const lineWidth = line.segments.reduce((s, seg) => s + seg.measuredWidth, 0);
    let alignOffset = 0;
    if (para.alignment === 'right' || para.alignment === 'end') {
      alignOffset = lineAvailW - (x - lineLeft) - lineWidth;
    } else if (para.alignment === 'center') {
      alignOffset = (lineAvailW - (x - lineLeft) - lineWidth) / 2;
    }
    x += alignOffset;

    // Inter-word adjustment per whitespace char on this line. Positive slack
    // (lineWidth < availW) expands spaces to fill; negative slack (lineWidth >
    // availW, typically from canvas measuring ~1 px wider than Word) compresses
    // spaces so the final glyph lands on the right margin instead of overflowing.
    // Compression is capped so we never eat more than the natural width of a
    // space, and is only applied when the line is a candidate for justification
    // (jc=both/distribute, not the last line unless distribute).
    let extraPerSpace = 0;
    const applyJustify = isJustified && (!isLastLine || stretchLastLine);
    if (applyJustify) {
      let totalTrailingSpaces = 0;
      for (let si = 0; si < line.segments.length; si++) {
        const seg = line.segments[si];
        if (si === line.segments.length - 1) break; // trailing spaces on final seg don't stretch
        if ('text' in seg) totalTrailingSpaces += countTrailingSpaces((seg as LayoutTextSeg).text);
      }
      const slack = lineAvailW - (x - lineLeft) - lineWidth;
      if (totalTrailingSpaces > 0) {
        extraPerSpace = slack / totalTrailingSpaces;
        // Don't compress past zero-width spaces — limit compression to at most
        // half the widest space on the line. Estimated from default font size.
        const minExtra = -line.ascent * 0.25;
        if (extraPerSpace < minExtra) extraPerSpace = minExtra;
      }
    }

    for (let si = 0; si < line.segments.length; si++) {
      const seg = line.segments[si];
      const isLastSeg = si === line.segments.length - 1;
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

function buildSegments(runs: DocRun[], state: RenderState): LayoutSeg[] {
  const segs: LayoutSeg[] = [];
  const pushTextPiece = (
    text: string,
    base: TextRun | FieldRun,
    vertAlign: 'super' | 'sub' | null,
  ) => {
    const displayText = (base.allCaps || base.smallCaps) ? text.toUpperCase() : text;
    // Ruby annotation rides with the WHOLE base text (typically 1-2 chars).
    // Splitting on word boundaries would lose the association, so attach
    // the annotation only to the first emitted segment.
    const ruby = (base as TextRun).ruby
      ? { text: (base as TextRun).ruby!.text, fontSizePt: (base as TextRun).ruby!.fontSizePt }
      : undefined;
    const revision = (base as TextRun).revision;
    let firstSeg = true;
    for (const word of splitTextForLayout(displayText)) {
      segs.push({
        text: word,
        bold: base.bold,
        italic: base.italic,
        underline: base.underline,
        strikethrough: base.strikethrough,
        fontSize: base.fontSize,
        color: base.color,
        fontFamily: base.fontFamily,
        vertAlign,
        measuredWidth: 0,
        smallCaps: base.smallCaps ?? false,
        doubleStrikethrough: base.doubleStrikethrough ?? false,
        highlight: base.highlight ?? null,
        ruby: firstSeg ? ruby : undefined,
        revision,
      });
      firstSeg = false;
    }
  };

  for (const run of runs) {
    if (run.type === 'text') {
      const t = run as unknown as TextRun & { type: 'text' };
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
      });
    }
  }
  return segs;
}

function findNearbyFontSize(runs: DocRun[], idx: number): number {
  // Look backwards then forwards for a text or field run to get font size
  for (let i = idx - 1; i >= 0; i--) {
    const r = runs[i];
    if (r.type === 'text') return (r as unknown as TextRun).fontSize;
    if (r.type === 'field') return (r as unknown as FieldRun).fontSize;
  }
  for (let i = idx + 1; i < runs.length; i++) {
    const r = runs[i];
    if (r.type === 'text') return (r as unknown as TextRun).fontSize;
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

/**
 * Split a text run into layout-segment strings.
 * Each segment is an atomic unit for word-level fitting; CJK overflow is handled in layoutLines.
 */
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

  while (queue.length > 0) {
    const seg = queue.shift()!;

    // ── Line-break sentinel ──────────────────────────────
    if ('lineBreak' in seg) {
      flush(seg.fontSize);
      continue;
    }

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
      seg.mathAscent = asc;
      seg.mathDescent = desc;
      if (currentLine.length > 0 && currentWidth + w > availW()) flush();
      addToLine(seg, w, seg.fontSize, asc, desc);
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
    let asc = m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent ?? s.fontSize * scale * 0.8;
    const desc = m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent ?? s.fontSize * scale * 0.2;
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
      const prefix = available > 0 ? fitCJKPrefix(ctx, s.text, available) : '';
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
    const end = x + width - margin;
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
  if (w <= 0 || h <= 0) return;
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
    const m = ctx.measureText(block.text);
    let tx = innerX;
    if (block.alignment === 'center' || block.alignment === 'distribute') {
      tx = innerX + Math.max(0, (innerW - m.width) / 2);
    } else if (block.alignment === 'right' || block.alignment === 'end') {
      tx = innerX + Math.max(0, innerW - m.width);
    }
    // Baseline = cursorY + ascent (approx 0.85 of font size for default fonts).
    const baseline = cursorY + fontPx * 0.85;
    ctx.fillText(block.text, tx, baseline);
    cursorY += lineHeights[i];
  }
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

  const totalColW = table.colWidths.reduce((s, w) => s + w, 0) * scale;
  const colScale = totalColW > contentW ? contentW / totalColW : 1;
  const colWidths = table.colWidths.map(w => w * scale * colScale);

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

  let y = state.y;
  for (let ri = 0; ri < table.rows.length; ri++) {
    const row = table.rows[ri];
    const rowH = rowHeights[ri];
    let x = tableX;
    let ci = 0;

    for (const cell of row.cells) {
      const span = Math.min(cell.colSpan, colWidths.length - ci);
      const cellW = colWidths.slice(ci, ci + span).reduce((s, w) => s + w, 0);

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
        if (!dryRun) renderCell(cell, table, x, y, cellW, drawH, state);
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
  if (segs.length === 0) {
    const fs = getDefaultFontSize(para);
    const { asc, desc } = emptyLineNaturalPx(fs, scale);
    return lineBoxHeight(para.lineSpacing, asc, desc, scale, state.docGrid, paraHasRuby, emptyIntendedSinglePx(para, scale));
  }
  const lines = layoutLines(state.ctx, segs, maxWidth, 0, scale, para.tabStops, undefined, state.fontFamilyClasses);
  return lines.reduce((s, l) => s + lineBoxHeight(para.lineSpacing, l.ascent, l.descent, scale, state.docGrid, paraHasRuby, l.intendedSingle), 0);
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
): void {
  const { ctx, scale } = state;

  if (cell.background) {
    ctx.fillStyle = `#${cell.background}`;
    ctx.fillRect(x, y, w, h);
  }

  drawCellBorders(ctx, x, y, w, h, cell.borders, table.borders, scale);

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
    const contentH = visibleContent.reduce(
      (s, ce) => s + measureCellElementHeight(cellState, ce, w - ml - mr, scale), 0);
    if (cell.vAlign === 'center') cellState.y = y + (h - contentH) / 2;
    else cellState.y = y + h - contentH - mb;
  }

  renderCellContent(cell.content, cellState);
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
): void {
  const top = cell.top ?? table.top;
  const bottom = cell.bottom ?? table.bottom;
  const left = cell.left ?? table.left;
  const right = cell.right ?? table.right;

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
function normalizeFontFamily(
  family: string | null,
  fontFamilyClasses: Record<string, string> = {},
): string {
  if (!family) return '"Noto Sans JP", "Hiragino Sans", "Meiryo", sans-serif';

  const escape = (s: string) => s.replace(/"/g, '\\"');
  const head = `"${escape(family)}"`;
  const lower = family.toLowerCase();

  // 1) Authoritative classification from word/fontTable.xml §17.8.3.10.
  const tableClass = fontFamilyClasses[family];
  if (tableClass && tableClass !== 'auto') {
    switch (tableClass) {
      case 'roman':
        return `${head}, "Yu Mincho", "YuMincho", "Hiragino Mincho ProN", "MS Mincho", "Noto Serif JP", "Noto Serif", serif`;
      case 'swiss':
        return `${head}, "Noto Sans JP", "Hiragino Sans", "Meiryo", sans-serif`;
      case 'modern':
        return `${head}, "Courier New", monospace`;
      default:
        // script / decorative — fall through to name-pattern matching
        break;
    }
  }

  // 2) Name-pattern fallback for fonts absent from fontTable or classified "auto".
  const isSerif =
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
    return `${head}, "Yu Mincho", "YuMincho", "Hiragino Mincho ProN", "MS Mincho", "Noto Serif JP", "Noto Serif", serif`;
  }

  if (lower.includes('meiryo') || family.includes('メイリオ')) {
    return `${head}, "Meiryo UI", "Meiryo", "Noto Sans JP", "Hiragino Sans", sans-serif`;
  }
  if (family.includes('游ゴシック') || /\byu\s*gothic\b/i.test(family) || lower.includes('yugothic')) {
    return `${head}, "Yu Gothic", "YuGothic", "Noto Sans JP", "Hiragino Sans", sans-serif`;
  }
  if (lower.includes('ipa')) {
    return `${head}, "IPAexGothic", "Noto Sans JP", "Hiragino Sans", sans-serif`;
  }
  if (lower.includes('segoe')) {
    return `${head}, "Segoe UI", sans-serif`;
  }
  return `${head}, "Noto Sans JP", "Hiragino Sans", "Meiryo", sans-serif`;
}

function getDefaultFontSize(para: DocParagraph): number {
  for (const run of para.runs) {
    if (run.type === 'text') {
      return (run as unknown as TextRun).fontSize;
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
    if (run.type === 'text') return (run as unknown as TextRun).fontFamily;
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
 */
function lineBoxHeight(
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
