import type {
  DocxDocumentModel, BodyElement, PaginatedBodyElement, DocParagraph, DocTable, DocTableRow, DocTableCell, CellElement,
  DocRun, DocxTextRun, ImageRun, ShapeRun, ShapeText, ShapeTextRun, FieldRun, HeaderFooter, HeadersFooters, LineSpacing, BorderSpec, TableBorders, CellBorders,
  TabStop, ParagraphBorders, ParaBorderEdge, DocxRunBorder, SectionProps, SectionGeom, DocNote, NumberingInfo, ColumnGeom, FramePr, TblpPr, DocSettings,
} from './types';
import type { ArrowEnd, Stroke } from '@silurus/ooxml-core';
import {
  buildCustomPath,
  buildShapePath,
  renderPresetShape,
  hasPreset,
  getCachedSvgImageByPath,
  preferVectorBlip,
  hexToRgba,
  autoContrastColor,
  resolveFill,
  applyStroke,
  drawArrowHead,
  lineEndRetract,
  retractLineEndpoint,
  getConnectorAnchors,
  mathToMathML,
  recolorSvg,
  crispOffset,
  PT_TO_PX,
  resolveBaseDirection,
  isHTMLCanvas,
  defaultDpr,
  classifyCjkFont,
  cjkFallbackChain,
  NON_CJK_SANS_FALLBACKS,
  NON_CJK_SERIF_FALLBACKS,
  resolveKinsokuRules,
  DEFAULT_KINSOKU_RULES,
  kinsokuAdjustedSplit,
  crossRunKinsokuRetract,
  isCjkBreakChar,
  classifyFontGeneric,
  isComplexScriptCodePoint,
  getCachedBitmapByPath,
  dropBitmapCacheByPath,
  drawImageCropped,
  metafileRasterSize,
  symbolFontToUnicode,
  isSymbolFontFamily,
  symbolTextToUnicodeSegments,
  docxBorderDashArray,
  fillDoubleBorder,
} from '@silurus/ooxml-core';
import type { MathNode, MathRenderer, KinsokuRules } from '@silurus/ooxml-core';
import { intendedSingleLinePx, correctLineMetrics } from './font-metrics.js';
import {
  segmentsHaveRtl,
  computeLineVisualOrder,
  resolveAlignEdge,
  type AlignEdge,
  type LineVisualOrder,
} from './bidi-line.js';
import {
  type FloatRect,
  isWrapFloat,
  resolveLineFloatWindow,
  skipPastTopAndBottom,
} from './float-layout.js';
import {
  distributeLineSlack,
  type SegStretch,
} from './text-distribute.js';
import {
  type FrameBox,
  computeFrameBox,
  registerFrameFloat,
  pushFloatRect,
} from './frame-geometry.js';
import {
  computeFloatTableBox,
  registerTableFloat,
  floatTableWrapSide,
} from './float-table-geometry.js';
import {
  xContainer,
  yContainer,
  resolveAnchorX,
  resolveAnchorY,
} from './anchor-geometry.js';
import {
  findMergeEndRow,
  resolveTableRowHeights,
  resolveSingleRowHeight,
} from './table-geometry.js';
import { justifiedPiecePositions } from '@silurus/ooxml-core';

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

export interface RenderState {
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
  scale: number;    // px per pt
  /** Device-pixel ratio the canvas was scaled by (`ctx.scale(dpr, dpr)`). Used to
   *  compute the crisp-line offset (see crispOffset) so thin axis-aligned strokes
   *  land on a single device row instead of straddling two. */
  dpr: number;
  contentX: number; // left of content area (px)
  contentW: number; // width of content area (px)
  y: number;        // current Y cursor (px)
  pageH: number;    // full page height (px)
  defaultColor: string;
  /** 0-based page index currently being rendered */
  pageIndex: number;
  /** total page count in the document */
  totalPages: number;
  /** preloaded drawable images keyed by `imageKey(imagePath, colorReplaceFrom)`
   *  (raster ImageBitmap or, for an `asvg:svgBlip` vector original, an
   *  HTMLImageElement) */
  images: Map<string, DecodedImage>;
  /** when true, layout is performed but nothing is drawn (used for header/footer height measurement) */
  dryRun: boolean;
  /** section left margin in pt — used to convert margin-relative anchor X to page-absolute */
  marginLeft: number;
  /** section right margin in pt — used by anchor positioning to resolve
   *  `<wp:positionH relativeFrom="margin">` and the `*Margin` family containers. */
  marginRight: number;
  /** ECMA-376 §17.6.11: the body's TOP/BOTTOM **inset** from the page edge in pt — the
   *  margin's MAGNITUDE (|margin|), NOT the signed pgMar value. A negative top/bottom
   *  margin (ST_SignedTwipsMeasure) measures the body |margin| from the page edge and
   *  overlaps the header/footer; `bodyMarginInsetPt` derives this at the writers
   *  (baseState, buildMeasureState). Read as the column region top (renderBodyElements /
   *  splitParagraphAcrossPages) and as the text-margin container for `relativeFrom=
   *  "topMargin"/"bottomMargin"/"margin"` anchors/frames (anchor-geometry, frame-geometry;
   *  §17.18.100 — the text-margin location IS the body edge). Do NOT treat as the signed
   *  margin: the overflow decision keeps the sign separately (header/footerOverflowPt). */
  marginTop: number;
  marginBottom: number;
  /** Section page width in pt. */
  pageWidth: number;
  /** Active anchor-image floats that constrain text layout on the current page. */
  floats: FloatRect[];
  /** Monotonic counter assigning a unique id to each registerAnchorFloats call,
   *  i.e. one id per paragraph per page. Used only to scope the implementation-
   *  defined (HEURISTIC) overlap avoidance to DIFFERENT paragraphs. Reset to 0
   *  on every page flip so measure and render assign matching paraIds. */
  floatParaSeq: number;
  /** ECMA-376 §17.6.5 docGrid (type + pitch), applied to auto line spacing. */
  docGrid: DocGridCtx;
  /** True when the document body contains East Asian text. Gates docGrid line-
   *  cell rounding of empty / anchor-only paragraph marks (see
   *  paragraphMarkLineHeight), which carry no text to classify themselves. */
  docEastAsian: boolean;
  /** ECMA-376 §17.8.3.10 — font→family map from word/fontTable.xml. Used by
   *  resolveFontFamily as the authoritative source of serif/sans-serif classification. */
  fontFamilyClasses: Record<string, string>;
  /** ECMA-376 §17.15.1.58–.60 — resolved Japanese line-breaking rules
   *  (kinsoku enabled flag + line-start/line-end forbidden character sets).
   *  Default is the application's Japanese kinsoku table with kinsoku ON. */
  kinsoku: KinsokuRules;
  /** ECMA-376 §17.15.1.25 `w:defaultTabStop` — the interval (points) at which
   *  automatic tab stops are generated after all custom stops. Threaded from
   *  `doc.settings.defaultTabStop` like `kinsoku` so the MEASURE pass matches
   *  the DRAW pass; falls back to {@link DEFAULT_TAB_PT} (720 twips = 36pt) when
   *  the document omits the element. */
  defaultTabPt: number;
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
  /** ECMA-376 §20.4.3.2/§20.4.3.5: a DrawingML anchor whose `<wp:positionV>`
   *  uses a page-level `relativeFrom` (page / margin / topMargin / bottomMargin
   *  / leftMargin / rightMargin / insideMargin / outsideMargin / column) is
   *  positioned independently of its source-order anchoring paragraph — Word
   *  lays it out as soon as the page is opened, so paragraphs that come BEFORE
   *  the anchor's paragraph in source order still wrap around it. To match,
   *  we pre-scan upcoming body paragraphs at every page-start and register
   *  these floats up front. This set records which paragraphs have had their
   *  page-level floats pre-registered on the current page, so
   *  {@link registerAnchorFloats} skips re-registering them when the main
   *  flow reaches that paragraph. Reset whenever floats are reset (page flip
   *  or column relocation that rolls back this paragraph's own floats). */
  pageAnchorPrescanned?: Set<DocParagraph>;
  /** ECMA-376 §20.4.2.10 `behindDoc` z-order: an anchored object with
   *  `behindDoc="0"` floats IN FRONT of the inline text/image flow. The flow is
   *  painted in document order, so a front-anchored shape in an EARLY paragraph
   *  would be overpainted by a LATER inline image (sample-13: the "Journal
   *  homepage" text box, anchored to the first paragraph, sat behind the inline
   *  masthead banner that follows it). When this collector is set, the body
   *  render defers each front-anchor draw into it (capturing the column band)
   *  and replays them after the whole page's flow, so front floats land on top.
   *  `null`/absent ⇒ draw in place (headers/footers and measurement passes). */
  deferFront?: Array<() => void> | null;
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
  /**
   * Lazy image-byte loader: fetch the raw bytes for an embedded image by zip
   * path, wrapped in a Blob of the given MIME (twin of pptx's `fetchImage`).
   * Supplied by {@link DocxDocument} (routing to its `getImage`), so the
   * renderer decodes images on demand instead of from inlined base64. When
   * omitted, images are skipped (no byte source).
   */
  fetchImage?: (path: string, mimeType: string) => Promise<Blob>;
  /** Called for each rendered text segment. Used to build a transparent text selection overlay. */
  onTextRun?: (run: DocxTextRunInfo) => void;
  /** Default `true`. When false, runs tagged with a `revision` (insertion or
   *  deletion from `<w:ins>` / `<w:del>`) render in their normal colour with
   *  no underline / strikethrough overlay — useful for a "final / no markup"
   *  view of a tracked document. */
  showTrackChanges?: boolean;
}

// ===== Image preloading =====

/**
 * A decoded, drawable image. Raster blips decode to an `ImageBitmap`
 * (createImageBitmap); the Microsoft `asvg:svgBlip` vector original decodes to
 * an `HTMLImageElement` (via core's path-keyed `getCachedSvgImageByPath`, since
 * `createImageBitmap` cannot rasterize SVG in every browser). Both are valid
 * `ctx.drawImage` sources with numeric `.width`/`.height`, so every draw site is
 * identical regardless of which kind was decoded.
 */
type DecodedImage = ImageBitmap | HTMLImageElement;

interface ImagePair {
  /** Zip path of the raster fallback (or the SVG part itself when no raster
   *  blip is embedded). The cache key + the byte-fetch path. */
  imagePath: string;
  /** MIME type of the blip at {@link ImagePair.imagePath}. */
  mimeType: string;
  /**
   * Zip path of the vector original from the `asvg:svgBlip` extension, when
   * present. Preferred over `imagePath`; the decoded image is still stored
   * under `imagePath`'s key so draw sites (which look up by `imagePath`) find
   * it unchanged.
   */
  svgImagePath?: string;
  colorReplaceFrom?: string;
  /**
   * Largest intended draw size (pt) over every reference to this key. Only used
   * to pick a raster target resolution for vector metafiles (WMF/EMF), which
   * have no intrinsic pixel size — the player must rasterize at a chosen size.
   * Raster (PNG/JPEG) and SVG paths ignore it (they carry/scale their own
   * resolution). Defaults to 0 when no size is known.
   */
  widthPt: number;
  heightPt: number;
  /** True when at least one reference to this image carries an `<a:srcRect>`
   *  crop, so the decode must prefer the raster (the crop math needs the
   *  bitmap's native pixel grid; an SVG vector original has none). */
  hasCrop?: boolean;
}

/** Returns a stable map key for an (imagePath, colorReplaceFrom) pair. */
function imageKey(imagePath: string, colorReplaceFrom?: string): string {
  return colorReplaceFrom ? `${imagePath}|clr:${colorReplaceFrom}` : imagePath;
}

type DocxFetchImage = (path: string, mime: string) => Promise<Blob>;

// Second-layer cache for the `a:clrChange` (colorReplaceFrom) result. The core
// path-keyed cache (getCachedBitmapByPath) holds the color-replacement-FREE
// bitmap — shared across every reference to a path and reclaimed with the
// document. The make-transparent pass (getImageData + putImageData, expensive)
// then runs once per (imagePath, colorReplaceFrom) pair and its ImageBitmap is
// kept here, so revisiting a page re-runs neither the decode NOR the recolor.
//
// Keyed FIRST by the document's `fetchImage` closure (one stable identity per
// DocxDocument), then by imageKey(imagePath, colorReplaceFrom) — mirroring the
// core cache's per-document namespacing so two documents sharing a zip path +
// replace colour don't cross-contaminate, and the whole map is reclaimed with
// the document. The stored value is an ImageBitmap (a fresh OffscreenCanvas
// raster), so on destroy it must be closed (see dropColorReplacedCache), the
// same GPU-lifecycle discipline the core cache follows through its promise.
const colorReplacedByFetch = new WeakMap<DocxFetchImage, Map<string, Promise<ImageBitmap>>>();

function colorReplacedCacheFor(fetchImage: DocxFetchImage): Map<string, Promise<ImageBitmap>> {
  let cache = colorReplacedByFetch.get(fetchImage);
  if (!cache) {
    cache = new Map();
    colorReplacedByFetch.set(fetchImage, cache);
  }
  return cache;
}

/**
 * Close every color-replaced ImageBitmap for one document's `fetchImage` and
 * forget the document. Call from `DocxDocument.destroy()` alongside
 * `dropBitmapCacheByPath` (base bitmaps) and `dropSvgImageCache` (SVG object
 * URLs) so all three per-document image caches release promptly. A no-op when no
 * clrChange image was decoded.
 */
export function dropColorReplacedCache(fetchImage: DocxFetchImage): void {
  const cache = colorReplacedByFetch.get(fetchImage);
  if (!cache) return;
  for (const p of cache.values()) p.then((b) => b.close()).catch(() => {});
  cache.clear();
  colorReplacedByFetch.delete(fetchImage);
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
  // Record one image reference (collapsing duplicate keys, tracking the max
  // intended draw size so a vector metafile is rasterized sharply enough for its
  // largest occurrence — only meaningful for WMF/EMF).
  const record = (pair: ImagePair) => {
    const key = imageKey(pair.imagePath, pair.colorReplaceFrom);
    const existing = seen.get(key);
    if (!existing) {
      seen.set(key, pair);
    } else {
      existing.widthPt = Math.max(existing.widthPt, pair.widthPt);
      existing.heightPt = Math.max(existing.heightPt, pair.heightPt);
      // If ANY reference is cropped, force the raster decode for this key.
      existing.hasCrop = existing.hasCrop || pair.hasCrop;
    }
  };
  // ECMA-376 §17.9.9/§17.9.20 — a level's picture-bullet marker is an image
  // that lives on the paragraph's numbering, not in any run. Feed it into the
  // same decode pipeline (keyed by its zip path) so the marker draw site finds
  // a decoded bitmap.
  const recordPara = (para: DocParagraph) => {
    const num = para.numbering;
    const pb = num?.picBulletImagePath;
    if (pb && num) {
      // Same §17.9.20 size resolution the draw site uses (picBulletSizePt): the
      // extent if present, else the resolved marker font size. Keeping the two in
      // lock-step matters for WMF/EMF bullets, where this size drives raster
      // sharpness — a 0 here would rasterize a vector bullet at zero size.
      const size = picBulletSizePt(num, para);
      record({
        imagePath: pb,
        mimeType: num.picBulletMimeType ?? '',
        widthPt: size.w,
        heightPt: size.h,
      });
    }
  };
  const walk = (runs: DocRun[]) => {
    for (const run of runs) {
      if (run.type === 'image') {
        const img = run as unknown as ImageRun;
        record({
          imagePath: img.imagePath,
          mimeType: img.mimeType,
          svgImagePath: img.svgImagePath,
          colorReplaceFrom: img.colorReplaceFrom,
          ...metafileRasterSize(img.mimeType, img.srcRect, img.widthPt ?? 0, img.heightPt ?? 0),
          hasCrop: img.srcRect != null,
        });
      } else if (run.type === 'shape') {
        // Inline images living inside a text box (<wps:txbx>) ride on the
        // shape's text blocks. Feed them into the same decode pipeline so the
        // WMF/EMF/raster/SVG decoders see their bytes (no colorReplace here).
        const shp = run as unknown as ShapeRun;
        for (const block of shp.textBlocks ?? []) {
          if (block.imagePath) {
            record({
              imagePath: block.imagePath,
              mimeType: block.mimeType ?? '',
              svgImagePath: block.svgImagePath,
              widthPt: block.imageWidthPt ?? 0,
              heightPt: block.imageHeightPt ?? 0,
            });
          }
        }
      }
    }
  };
  const walkTable = (tbl: DocTable) => {
    for (const row of tbl.rows)
      for (const cell of row.cells)
        for (const ce of cell.content) {
          if (ce.type === 'paragraph') {
            const p = ce as unknown as DocParagraph;
            recordPara(p);
            walk(p.runs);
          } else if (ce.type === 'table') walkTable(ce as unknown as DocTable);
        }
  };
  const walkBody = (body: BodyElement[]) => {
    for (const el of body) {
      if (el.type === 'paragraph') {
        const p = el as unknown as DocParagraph;
        recordPara(p);
        walk(p.runs);
      }
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

/**
 * Decode a raster blip to an `ImageBitmap`, pulling the bytes lazily by zip path
 * via `fetchImage(imagePath, mimeType)` (twin of pptx's `fetchImage`) rather
 * than `fetch`-ing an inlined data URL. Applies an `a:clrChange`
 * (`colorReplaceFrom`) make-transparent pass when requested — unchanged
 * post-decode behavior.
 *
 * Two-layer caching so a page revisit re-runs NEITHER the decode NOR the recolor:
 *  1. the color-replacement-free bitmap comes from the shared, per-document,
 *     path-keyed {@link getCachedBitmapByPath} (the raster/metafile cache docx,
 *     pptx and xlsx now share). It content-sniffs the bytes (extension/MIME are
 *     unreliable — sample-10's chart is a standard WMF mislabeled `.emf`),
 *     rasterizing a WMF via the minimal player at a size from `widthPt`/`heightPt`,
 *     returning `null` for a true EMF (or a geometry-less metafile), else
 *     `createImageBitmap`. A `null` throws so `preloadImages` drops the image
 *     (the existing "missing image" behavior, no crash).
 *  2. when a clrChange is requested, the make-transparent result is memoized per
 *     (imagePath, colorReplaceFrom) in {@link colorReplacedCacheFor} so the
 *     expensive getImageData/putImageData pass runs once per document.
 *
 * `suppressBoundaryFrame: true` is REQUIRED: docx's former in-tree player ran the
 * window/device-boundary edge suppression unconditionally (to hide sample-10's
 * Fig.1 cosmetic outer frame). Core defaults that heuristic OFF (spec-clean), so
 * docx must opt in here to preserve its current rendering.
 *
 * Exported for unit testing of the lazy-bytes contract.
 */
export async function decodeRaster(
  imagePath: string,
  mimeType: string,
  colorReplaceFrom: string | undefined,
  fetchImage: (path: string, mime: string) => Promise<Blob>,
  widthPt = 0,
  heightPt = 0,
): Promise<ImageBitmap> {
  // Base bitmap (no colour replacement): shared, path-keyed, per-document cache.
  const base = await getCachedBitmapByPath(imagePath, mimeType, fetchImage, {
    widthPt,
    heightPt,
    suppressBoundaryFrame: true,
  });
  if (!base) throw new Error(`${imagePath} produced no drawable output`);
  if (!colorReplaceFrom) return base;
  // Second layer: memoize the make-transparent result per (path, colour). The
  // recolor reads the SHARED base bitmap and produces a fresh independent raster,
  // so the base is never mutated and stays reusable for other references / draws.
  const cache = colorReplacedCacheFor(fetchImage);
  const key = imageKey(imagePath, colorReplaceFrom);
  let hit = cache.get(key);
  if (!hit) {
    hit = applyColorReplacement(base, colorReplaceFrom);
    // Don't poison the cache if the recolor pass rejects; let the next call retry.
    hit.catch(() => cache.delete(key));
    cache.set(key, hit);
  }
  return hit;
}

/**
 * Decode every embedded image referenced by the document into a drawable map
 * keyed by `imageKey(imagePath, colorReplaceFrom)`. Bytes are fetched lazily by
 * zip path via `fetchImage`; SVG vector originals decode through the path-keyed
 * `<img>` helper. Returns an empty map when `fetchImage` is absent (no byte
 * source) — draw sites then simply skip.
 *
 * Exported for unit testing of the keying + single-decode-per-key contract.
 */
export async function preloadImages(
  doc: DocxDocumentModel,
  fetchImage: ((path: string, mime: string) => Promise<Blob>) | undefined,
): Promise<Map<string, DecodedImage>> {
  if (!fetchImage) return new Map();
  const fetch = fetchImage;
  const pairs = collectImagePairs(doc);
  const entries = await Promise.all(
    pairs.map(async (pair): Promise<[string, DecodedImage] | null> => {
      // Unified svgBlip selection (shared with pptx/xlsx). The decoded image is
      // keyed by the raster `imagePath` regardless of which source we picked, so
      // every draw site finds it via imageKey(imagePath, …) unchanged.
      const dataIsSvg = pair.mimeType === 'image/svg+xml';
      // Shared vector-vs-raster gate (see core preferVectorBlip). `hasCrop` is
      // this format's already-aggregated "any reference to this key is cropped"
      // flag, so it stands in for srcRect presence (`|| null` normalises the
      // undefined case). When true, `blip.svgImagePath` is narrowed to string.
      const blip = { svgImagePath: pair.svgImagePath, srcRect: pair.hasCrop || null };
      try {
        let img: DecodedImage;
        if (preferVectorBlip(blip)) {
          // Prefer the vector original (Microsoft `asvg:svgBlip` extension);
          // fall back to the raster on any SVG decode failure. With an
          // `<a:srcRect>` crop (§20.1.8.55) we skip this branch and decode the
          // raster instead, because the crop math (drawImageCropped) needs the
          // bitmap's native pixel grid — an SVG element has none.
          try {
            img = await getCachedSvgImageByPath(blip.svgImagePath, fetch);
          } catch {
            img = dataIsSvg
              ? await getCachedSvgImageByPath(pair.imagePath, fetch)
              : await decodeRaster(pair.imagePath, pair.mimeType, pair.colorReplaceFrom, fetch, pair.widthPt, pair.heightPt);
          }
        } else if (dataIsSvg) {
          // svg-only picture (no svgImagePath surfaced — e.g. a non-svgBlip
          // `.svg` part): `createImageBitmap` can't rasterize SVG, so decode
          // through the path-keyed <img>-based SVG path.
          img = await getCachedSvgImageByPath(pair.imagePath, fetch);
        } else {
          img = await decodeRaster(pair.imagePath, pair.mimeType, pair.colorReplaceFrom, fetch, pair.widthPt, pair.heightPt);
        }
        return [imageKey(pair.imagePath, pair.colorReplaceFrom), img];
      } catch {
        return null;
      }
    }),
  );
  return new Map(entries.filter((e): e is [string, DecodedImage] => e !== null));
}

// ===== Main entry =====

/**
 * Per-canvas monotonic render token for the {@link renderDocumentToCanvas}
 * cancellation guard. A WeakMap keyed on the canvas replaces the previous
 * property monkey-patch (`canvas.__docxRenderToken`), so no non-standard field
 * is written onto the caller's canvas and the `as unknown as` cast is gone.
 * WeakMap keys are held weakly, so a discarded canvas is collected normally.
 * (Mirrors the pptx renderSlide guard's renderTokens map.)
 */
const renderTokens = new WeakMap<HTMLCanvasElement | OffscreenCanvas, number>();

export async function renderDocumentToCanvas(
  doc: DocxDocumentModel,
  canvas: HTMLCanvasElement | OffscreenCanvas,
  pageIndex: number,
  opts: RenderDocumentOptions = {},
): Promise<void> {
  // Cancellation guard. renderDocumentToCanvas is async (it awaits image decode
  // via preloadImages), so rapid page navigation can start a newer render of the
  // SAME canvas before this one finishes. Both clear the canvas (`canvas.width =
  // …` + a white fillRect) up front and then draw their page AFTER the await —
  // so the clears run first and the draws accumulate, ghosting several pages on
  // top of each other. Stamp a per-canvas token; once a newer render supersedes
  // us, stop at the next await so only the latest render's output survives.
  // (Mirrors the pptx renderSlide guard; the worker path renders each page on a
  // fresh OffscreenCanvas, so the token is a no-op there.)
  const myToken = (renderTokens.get(canvas) ?? 0) + 1;
  renderTokens.set(canvas, myToken);
  const superseded = () => renderTokens.get(canvas) !== myToken;

  const dpr = opts.dpr ?? defaultDpr();
  // getContext before sizing is legal: resizing a canvas after getContext resets
  // its drawing state, and the ctx.scale/fill below run AFTER canvas.width/height.
  // The pagination fallback (paginateWithHeaderFooterReserve) needs this ctx.
  const ctx = canvas.getContext('2d') as CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
  const kinsoku = resolveKinsokuRules(doc.settings);
  const pages = opts.prebuiltPages ?? paginateWithHeaderFooterReserve(doc, ctx, doc.fontFamilyClasses ?? {}, kinsoku, doc.footnotes ?? []);
  const totalPages = Math.max(opts.totalPages ?? pages.length, pages.length);
  const elements = pages[pageIndex] ?? pages[0] ?? [];

  // ECMA-376 §17.6.13 / §17.6.11 — page geometry is PER-SECTION. Size THIS page from
  // the section active at its top (resolvePageSection.geom, stamped by the paginator),
  // NOT from the single body-level `doc.section`. `sec` merges the resolved geometry
  // (size + margins + header/footer distances) over the body-level section so the
  // docGrid / columns / sectionStart / even-odd fields keep their body-level values —
  // those already flow per-section through the paginator's `colGeom`/docGrid state
  // rails, so only the page-box geometry needs the per-page swap here. For a
  // single-section document `geom` equals `doc.section`, so `sec === doc.section` in
  // value — byte-identical output.
  const pageGeom = resolvePageSection(pages, pageIndex, doc).geom;
  const sec: SectionProps = { ...doc.section, ...pageGeom };

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

  ctx.scale(dpr, dpr);
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, cssWidth, cssHeight);

  const images = await preloadImages(doc, opts.fetchImage);
  // A newer render of this canvas started while we awaited image decode — stop
  // so we don't paint this (now stale) page over the newer one.
  if (superseded()) return;

  // ECMA-376 §17.11: map each note id to its 1-based display number so the
  // reference markers (and the in-note footnoteRef placeholder) show the
  // sequential number, not the raw @w:id.
  const footnoteNums = buildNoteNumberMap(doc.footnotes);
  const endnoteNums = buildNoteNumberMap(doc.endnotes);
  const noteNumbers = new Map<string, number>();
  for (const [id, n] of footnoteNums) noteNumbers.set(`footnote:${id}`, n);
  for (const [id, n] of endnoteNums) noteNumbers.set(`endnote:${id}`, n);

  const docEA = documentHasEastAsian(doc.body);
  // ECMA-376 §17.6.11: the body is inset from each page edge by the margin's MAGNITUDE
  // (a negative margin places the body |margin| inside the edge, overlapping the
  // header/footer — see bodyMarginInsetPt). Identity for the non-negative common case.
  // The header/footer reserves below still use the SIGNED margin (header/footerOverflowPt
  // return 0 for a negative one), so a negative margin reserves nothing yet insets |margin|.
  const bodyTopPt = bodyMarginInsetPt(sec.marginTop);
  const bodyBottomPt = bodyMarginInsetPt(sec.marginBottom);
  const baseState: RenderState = {
    ctx,
    scale,
    dpr,
    contentX: sec.marginLeft * scale,
    contentW: (sec.pageWidth - sec.marginLeft - sec.marginRight) * scale,
    y: bodyTopPt * scale,
    pageH: cssHeight,
    defaultColor: opts.defaultTextColor ?? '#000000',
    pageIndex,
    totalPages,
    images,
    dryRun: false,
    marginLeft: sec.marginLeft,
    marginRight: sec.marginRight,
    // §17.6.11: store the body inset (|margin|), the value the paint pass re-adds as a
    // column's region top (state.marginTop, renderBodyElements); never the raw sign.
    marginTop: bodyTopPt,
    marginBottom: bodyBottomPt,
    pageWidth: sec.pageWidth,
    floats: [],
    floatParaSeq: 0,
    docGrid: {
      type: sec.docGridType ?? null,
      linePitchPt: sec.docGridLinePitch ?? null,
      charSpacePt: sec.docGridCharSpace != null ? sec.docGridCharSpace / 4096 : null,
    },
    docEastAsian: docEA,
    fontFamilyClasses: doc.fontFamilyClasses ?? {},
    kinsoku,
    // §17.15.1.25 — automatic tab interval, resolved once and threaded like
    // `kinsoku` so the measure and draw passes agree.
    defaultTabPt: resolveDefaultTabPt(doc.settings),
    mathDefJc: doc.settings?.mathDefJc,
    onTextRun: opts.onTextRun,
    showTrackChanges: opts.showTrackChanges ?? true,
    noteNumbers,
  };

  // ECMA-376 §17.10.1 — per-section header/footer selection. resolvePageHeader and
  // resolvePageFooter resolve the section active at the top of this page (and whether
  // it is that section's first page) from the paginated elements' stamped `sectionHF`,
  // then apply the spec precedence (first → even → default). Single-section docs
  // resolve to doc.headers/footers/section.titlePage, unchanged. The paint path uses
  // the SAME resolvers as the reserve pass (computeHeaderReserves / computeFooterReserves)
  // so the body's start/end can never drift from the gap pagination reserved.

  // Header: top of page, starting at headerDistance. A header taller than its
  // top-margin allowance (§17.6.11) overflows the content area downward; the body
  // was already paginated to clear it (paginateWithHeaderFooterReserve), and its
  // start y is pushed down by the same overflow so no body line sits over the header.
  const header = resolvePageHeader(pages, pageIndex, doc);
  let headerReservePx = 0;
  if (header) {
    const headerHeight = measureHeaderFooterHeight(header, baseState);
    renderHeaderFooter(header, sec.headerDistance * scale, baseState);
    // §17.6.11 overflow in device px (headerHeight is at canvas scale), via the shared
    // formula so the body start matches the pagination reserve exactly.
    headerReservePx = headerOverflowPt(headerHeight, sec.marginTop * scale, sec.headerDistance * scale);
  }

  // Footer: anchored from bottom, rising by its measured height. A footer taller
  // than its bottom-margin allowance (§17.6.11) overflows the content area; the body
  // was already paginated to clear it (paginateWithHeaderFooterReserve), and the same
  // overflow raises the footnote block below so notes clear it too.
  const footer = resolvePageFooter(pages, pageIndex, doc);
  let footerReservePx = 0;
  if (footer) {
    const footerHeight = measureHeaderFooterHeight(footer, baseState);
    const footerTopY = cssHeight - sec.footerDistance * scale - footerHeight;
    renderHeaderFooter(footer, footerTopY, baseState);
    // §17.6.11 overflow in device px (footerHeight is at canvas scale), via the shared
    // formula so the footnote clearance matches the pagination reserve exactly.
    footerReservePx = footerOverflowPt(footerHeight, sec.marginBottom * scale, sec.footerDistance * scale);
  }

  // Body. ECMA-376 §17.6.4: lay out body text in EACH section's newspaper columns
  // (per-section columns). `columns` is the body-level (final) section's geometry,
  // used as the fallback for elements that carry no per-section `colGeom` (single-
  // section docs, where it equals the whole-body geometry — unchanged path).
  const columns = computeColumns(sec);
  // ECMA-376 §17.6.11 (pgMar/@top): the body starts at the GREATER of the top margin
  // and the header's extent, so a tall header (headerReservePx > 0) pushes the first
  // body line down to the header's bottom. The same overflow shrank the paginated
  // content area from the top (computeHeaderReserves), so the body fits within margins.
  const bodyState: RenderState = { ...baseState, y: bodyTopPt * scale + headerReservePx };
  // Optional column separator rules (`<w:cols w:sep="1">`), drawn before the text
  // so glyphs sit on top. A thin rule is centred in each inter-column gap and
  // spans the content height. With per-section columns a page can carry more than
  // one section's geometry (a continuous break), so draw separators for each
  // DISTINCT multi-column geometry actually present on this page (derived from the
  // elements' stamped `colGeom`), falling back to the page-level `columns` when an
  // element carries none. The `sep` flag is the final section's
  // (`sec.columns?.sep`); threading a per-section sep toggle is unnecessary until a
  // document mixes differing sep settings across sections sharing a page (rare,
  // untested — both bundled samples use sep:false).
  if (sec.columns?.sep) {
    const seen = new Set<ColumnGeom[]>();
    const geoms: ColumnGeom[][] = [];
    for (const el of elements) {
      const g = (el as PaginatedBodyElement).colGeom ?? columns;
      if (g.length > 1 && !seen.has(g)) {
        seen.add(g);
        geoms.push(g);
      }
    }
    if (geoms.length === 0 && columns.length > 1) geoms.push(columns);
    for (const g of geoms) drawColumnSeparators(ctx, g, sec, scale);
  }
  renderBodyElements(elements, bodyState, columns, headerReservePx);

  // Footnotes referenced on this page (ECMA-376 §17.11): drawn at the bottom of
  // the text column, above a short separator rule. The page area was already
  // reserved during pagination so the body stops short of them.
  drawPageFootnotes(elements, doc, baseState, scale, cssHeight, sec, footerReservePx);

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
  docEastAsian: boolean,
  // §17.15.1.25 — keep the note measure pass on the same automatic tab interval.
  defaultTabPt: number = DEFAULT_TAB_PT,
): { total: number; trailingSpaceAfter: number } {
  const measure = buildMeasureState(ctx, sec, fontFamilyClasses, kinsoku, docEastAsian, defaultTabPt);
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
  footerReservePx = 0,
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
    const m = measureNoteBlockForDraw(note, baseState.ctx, sec, baseState.fontFamilyClasses, baseState.kinsoku, baseState.docEastAsian, baseState.defaultTabPt);
    totalPt += m.total;
    lastTrailingPt = m.trailingSpaceAfter;
  }
  const contentPt = Math.max(0, totalPt - lastTrailingPt);
  const gapPx = FOOTNOTE_SEPARATOR_GAP_PT * scale;
  // Clear a footer taller than the bottom margin (§17.6.11): the footnote block, like
  // the body, sits above the footer's overflow (footerReservePx), not just the margin.
  const blockTopY = cssHeight - bodyMarginInsetPt(sec.marginBottom) * scale - footerReservePx - contentPt * scale;

  // Separator rule: short left-aligned line (Word's default footnote separator,
  // ~1/3 of the text width), centered in the gap above the notes.
  const leftX = sec.marginLeft * scale;
  const ruleY = Math.round(blockTopY - gapPx);
  const ctx = baseState.ctx;
  ctx.save();
  ctx.strokeStyle = baseState.defaultColor;
  const ruleLw = Math.max(1, Math.round(0.5 * scale));
  ctx.lineWidth = ruleLw;
  // Crispness nudge (see crispOffset): a horizontal rule snaps onto the nearest
  // crisp device row from its own y. The previous hardcoded `+ 0.5` was a
  // logical-px offset that only happened to be crisp at dpr=1 and blurred at
  // dpr>1; crispOffset makes it device-correct at any dpr and fractional y.
  const ruleNudge = crispOffset(ruleY, ruleLw, baseState.dpr);
  ctx.beginPath();
  ctx.moveTo(leftX, ruleY + ruleNudge);
  ctx.lineTo(leftX + (sec.pageWidth - sec.marginLeft - sec.marginRight) * scale / 3, ruleY + ruleNudge);
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
  const maxY = cssHeight - bodyMarginInsetPt(sec.marginBottom) * scale;
  if (y >= maxY) return;

  // Separator rule above the endnotes.
  const leftX = sec.marginLeft * scale;
  ctx.save();
  ctx.strokeStyle = bodyState.defaultColor;
  const ruleLw = Math.max(1, Math.round(0.5 * scale));
  ctx.lineWidth = ruleLw;
  // Crispness nudge (see crispOffset): device-correct replacement for the old
  // hardcoded logical `+ 0.5`, which only rendered crisp at dpr=1. Snap from the
  // rule's own y so it stays crisp at any dpr and fractional position.
  const ruleY = Math.round(y);
  const ruleNudge = crispOffset(ruleY, ruleLw, bodyState.dpr);
  ctx.beginPath();
  ctx.moveTo(leftX, ruleY + ruleNudge);
  ctx.lineTo(leftX + (sec.pageWidth - sec.marginLeft - sec.marginRight) * scale / 3, ruleY + ruleNudge);
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

/** ECMA-376 §17.10.1 — an empty header/footer set (no default/first/even). Used
 *  when a per-section `SectionBreak` declares one reference type but not others:
 *  the absent types fall back to this so `pickHeaderFooter` simply finds none and
 *  renders nothing for that page kind. */
const EMPTY_HEADERS_FOOTERS: HeadersFooters = { default: null, first: null, even: null };

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
// `ColumnGeom` (one newspaper column's page-absolute x/width, pt) is defined in
// ./types (alongside ColumnsSpec/ColSpec) so it can be referenced from
// `PaginatedBodyElement` without a renderer↔types import cycle. Re-exported here
// for callers that import it from the renderer.
export type { ColumnGeom } from './types';

/**
 * ECMA-376 §17.6.4 — resolve a section's newspaper columns to page-absolute
 * left-x / width pairs (pt). The content band is `[marginLeft, pageWidth -
 * marginRight]`; columns tile it left-to-right.
 *
 * - No columns (or count <= 1): one full-width column spanning the content band
 *   (unchanged single-column behavior).
 * - Equal width (`equalWidth`): `colW = (contentW - (count-1)*space) / count`;
 *   column i sits at `marginLeft + i*(colW + space)`.
 * - Explicit `<w:col>` widths: walk the columns, advancing x by each column's
 *   own width + trailing space. The per-column widths/spaces are used verbatim
 *   (Word writes them to sum to the content band).
 *
 * `colW` is clamped to a positive minimum so a malformed/over-wide spec never
 * yields a zero/negative text width that would wedge line layout.
 */
export function computeColumns(section: SectionProps): ColumnGeom[] {
  const contentW = section.pageWidth - section.marginLeft - section.marginRight;
  const cols = section.columns;
  if (!cols || cols.count <= 1) {
    return [{ xPt: section.marginLeft, wPt: Math.max(1, contentW) }];
  }

  // Explicit per-column geometry (unequal widths).
  if (!cols.equalWidth && cols.cols.length > 0) {
    const out: ColumnGeom[] = [];
    let x = section.marginLeft;
    for (const c of cols.cols) {
      out.push({ xPt: x, wPt: Math.max(1, c.widthPt) });
      x += c.widthPt + c.spacePt;
    }
    return out;
  }

  // Equal-width columns separated by `space`.
  const count = cols.count;
  const space = cols.spacePt;
  const colW = Math.max(1, (contentW - (count - 1) * space) / count);
  const out: ColumnGeom[] = [];
  for (let i = 0; i < count; i++) {
    out.push({ xPt: section.marginLeft + i * (colW + space), wPt: colW });
  }
  return out;
}

/** ECMA-376 §17.6.11 — the per-page content-area insets (pt) reserved for a header
 *  taller than its top-margin allowance (`top`) and a footer taller than its
 *  bottom-margin allowance (`bottom`). One value per page; see headerOverflowPt /
 *  footerOverflowPt. */
interface PageReserve { top: number; bottom: number }
const ZERO_RESERVE: PageReserve = { top: 0, bottom: 0 };

export function computePages(
  body: BodyElement[],
  section: SectionProps,
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  fontFamilyClasses: Record<string, string> = {},
  kinsoku: KinsokuRules = DEFAULT_KINSOKU_RULES,
  footnotes: DocNote[] = [],
  /** Per-page tall header/footer reserve (ECMA-376 §17.6.11). Empty ⇒ no reserve
   *  (the common case). Computed by paginateWithHeaderFooterReserve's second pass. */
  pageReserves: PageReserve[] = [],
  /** ECMA-376 §17.15.1.25 — automatic tab-stop interval (pt). Threaded so the
   *  pagination measure pass advances tabs identically to the draw pass; defaults
   *  to the spec absent value when a caller has no document settings. */
  defaultTabPt: number = DEFAULT_TAB_PT,
): PaginatedBodyElement[][] {
  // ECMA-376 §17.6.11: the body is inset from each page edge by the margin's MAGNITUDE
  // (a negative margin measures the body |margin| from the edge and overlaps the
  // header/footer — see bodyMarginInsetPt). Identity for the non-negative common case.
  // ECMA-376 §17.6.11 — the body inset (|margin|) and §17.6.13 full content height
  // are PER-SECTION now: a section break can change the page size/margins. These
  // read the ACTIVE section (`currentSectionGeom`, reassigned at every break) so a
  // landscape section paginates against its own (wider/shorter) page. For a
  // single-section document `currentSectionGeom` never changes ⇒ identical values.
  // `currentSectionGeom` is declared below (TDZ-safe: these arrows only close over
  // it, and the first call happens during the element loop, after its `let` init).
  const bodyTopPt = () => bodyMarginInsetPt(currentSectionGeom.marginTop);
  const bodyBottomPt = () => bodyMarginInsetPt(currentSectionGeom.marginBottom);
  const fullContentH = () => currentSectionGeom.pageHeight - bodyTopPt() - bodyBottomPt();
  const measureState = buildMeasureState(ctx, section, fontFamilyClasses, kinsoku, documentHasEastAsian(body), defaultTabPt);
  const noteById = indexNotes(footnotes);
  const haveFootnotes = noteById.size > 0;
  // Per-page reserved footnote height (pt). Index 0 = first page. Grows as
  // footnote references are placed on the current page.
  const footnoteReservePt: number[] = [0];

  // ECMA-376 §17.6.4 newspaper columns are PER-SECTION. A `<w:sectPr>` carried in
  // a paragraph's `pPr` (or a loose mid-body one) ends a section; the parser emits
  // a `SectionBreak` marker carrying THAT section's `<w:cols>`. The FINAL section's
  // columns live on `section.columns` (the body-level sectPr). So the columns for
  // the section that STARTS at body index `startIdx` are the `.columns` of the
  // NEXT `SectionBreak` at/after `startIdx`; if there is none, the body-level
  // section's columns. A `<w:cols>`-less section (columns == null) ⇒ one full-width
  // column (computeColumns's single-column path), exactly as before.
  const sectionColumnsFrom = (startIdx: number): ColumnGeom[] => {
    for (let j = startIdx; j < body.length; j++) {
      const e = body[j];
      if (e.type === 'sectionBreak') {
        // Resolve this section's cols by swapping in its ColumnsSpec. Page
        // GEOMETRY is now per-section (`sectionGeomFrom`), but `computeColumns`
        // here still receives the BODY-LEVEL width (`section.pageWidth`/margins) —
        // a documented residual: a mid-body section with a different page WIDTH and
        // multiple columns tiles against the body width, not its own. Single-column
        // sections and uniform-width documents are unaffected. Aligning column
        // widths to the per-section width is the width-residual follow-up in the
        // design doc; do NOT change it here.
        return computeColumns({ ...section, columns: e.columns ?? null });
      }
    }
    return computeColumns(section);
  };

  // ECMA-376 §17.6.22 (ST_SectionMark): a section's `<w:type>` specifies how
  // THAT section begins relative to the previous one (continuous ⇒ same page;
  // nextPage/odd/even ⇒ a new page). The type lives on the section's OWN sectPr,
  // which the parser stamps on the SectionBreak marker that ENDS that section.
  // So the break at a boundary is governed by the UPCOMING section's type — the
  // kind of the NEXT marker at/after `startIdx`, exactly like sectionColumnsFrom
  // resolves the upcoming section's columns. (A marker's own kind is the start
  // type of the section it closes — relevant at the PREVIOUS boundary, not this
  // one. Using it here is an off-by-one that turns e.g. a title section's
  // type="nextPage" into a spurious page break before a following
  // type="continuous" body section.) The last section (no following marker) uses
  // the body sectPr's start type; absent ⇒ "nextPage" (the spec default).
  //
  // A `continuous` body that would otherwise flow up onto a full cover page does
  // NOT regress here: Word places a "Cover Pages" building block (ECMA-376
  // §17.5.2 docPartGallery) on its own page, and the parser models that by
  // emitting a PageBreak after the cover's content — so the continuous section
  // after a cover lands on the next page regardless of this upcoming-section
  // reading. See `parse_body_elements` (cover-page page-fill) in the parser.
  const sectionKindFrom = (startIdx: number): string => {
    for (let j = startIdx; j < body.length; j++) {
      const e = body[j];
      if (e.type === 'sectionBreak') return e.kind ?? 'nextPage';
    }
    return section.sectionStart ?? 'nextPage';
  };

  // Whether `body[idx]` is an empty section-break spacer whose break is CONTINUOUS
  // — the precise trigger for Word's spacing-before suppression (see
  // isSectionBreakSpacerAt). The break at idx+1 is governed by the FOLLOWING
  // section's start type (§17.6.22), so its effective kind is sectionKindFrom of
  // the section starting at idx+2 (the marker's own `.kind` can read "nextPage"
  // while the section actually continues — sample-13). A next-page/odd/even break
  // does NOT suppress (the spacer there is the closing section's trailing line on
  // a page that is ending anyway).
  const isContinuousSectionSpacer = (idx: number): boolean =>
    isSectionBreakSpacerAt(body, idx) && sectionKindFrom(idx + 2) === 'continuous';

  // A continuous-section spacer whose OWN space-before is zero. Word renders NO
  // paragraph-mark line box for it: the section mark collapses to zero height
  // instead of occupying a blank line. (A spacer that DOES carry a space-before
  // keeps its box — the before manifests as the blank line.)
  //
  // NOTE — this matches Microsoft WORD's observed layout, NOT a spec rule.
  // §17.3.1.29 mandates that every paragraph produces one mark line box; no
  // ECMA-376 clause (nor [MS-DOC] / [MS-OI29500]) documents a section-mark
  // collapse. The model is reconstructed clean-room from Word's OWN output:
  //   - sample-12's spacer is Normal with before=0. Word shows exactly ONE blank
  //     line between "[Format…single line spacing]" and "1. INTRODUCTION" (the
  //     user verified this by cursor-walk) and paints the heading ~24pt higher
  //     than our prior two-blank-line layout — i.e. the spacer's own mark line is
  //     absent (pdftotext -bbox: INTRODUCTION at 446pt, not 470pt).
  //   - sample-13's spacer is before=440 (22pt). Word keeps its mark line (TWO
  //     blank lines, heading at 376pt), so the collapse is gated on before=0.
  // Both gates were pinned against the Word PDFs; we follow Word's measured
  // behaviour, not any external implementation.
  const isCollapsedContinuousSpacer = (idx: number): boolean => {
    if (!isContinuousSectionSpacer(idx)) return false;
    const p = body[idx] as unknown as DocParagraph;
    return (p.spaceBefore ?? 0) === 0;
  };

  // ECMA-376 §17.10.1 — the resolved header/footer set + `<w:titlePg>` flag for
  // the section that OWNS the content starting at body index `startIdx`. Mirrors
  // `sectionColumnsFrom`: the owning section's set is carried on the NEXT
  // `SectionBreak` marker at/after `startIdx` (the marker ENDS that section); if
  // there is none, the content belongs to the FINAL (body-level) section whose
  // set lives on `doc.headers`/`doc.footers`/`section.titlePage`. The paginator
  // stamps this on each element so the renderer can pick the active section's
  // header/footer per page without re-deriving the body→page mapping.
  const sectionHFFrom = (startIdx: number): PaginatedBodyElement['sectionHF'] => {
    for (let j = startIdx; j < body.length; j++) {
      const e = body[j];
      if (e.type === 'sectionBreak') {
        return {
          headers: e.headers ?? EMPTY_HEADERS_FOOTERS,
          footers: e.footers ?? EMPTY_HEADERS_FOOTERS,
          titlePage: e.titlePage ?? false,
        };
      }
    }
    // Final section: the body-level set. `undefined` ⇒ the renderer's fallback
    // (doc.headers/footers/section.titlePage), which IS this same set.
    return undefined;
  };

  // ECMA-376 §17.6.13 / §17.6.11 — the page geometry (size + margins) of the
  // section that OWNS the content starting at body index `startIdx`. Mirrors
  // `sectionColumnsFrom` / `sectionHFFrom`: the owning section's geometry is on the
  // NEXT `SectionBreak` marker's `geom` at/after `startIdx` (the marker ENDS that
  // section); if there is none, the content belongs to the FINAL (body-level)
  // section whose geometry lives on `section`. A `geom`-less marker (a section that
  // inherits pgSz/pgMar) also falls back to the body-level geometry. For a
  // single-section document there are no markers, so every element gets the
  // body-level geometry — behaviour-neutral.
  const bodySectionGeom: SectionGeom = sectionGeomOf(section);
  const sectionGeomFrom = (startIdx: number): SectionGeom => {
    for (let j = startIdx; j < body.length; j++) {
      const e = body[j];
      if (e.type === 'sectionBreak') return e.geom ?? bodySectionGeom;
    }
    return bodySectionGeom;
  };

  // The active section's column geometry. Reassigned (a) here for the first
  // section and (b) at every `SectionBreak` as the flow enters the next section.
  // `colIndex` tracks which column we are filling; `colX()`/`colW()` give its
  // page-absolute left edge and text width (pt). Measurement uses the column
  // width (not the full content band) and the column's left x as the float-window
  // paraX, so square floats only constrain the column(s) their x-range intersects.
  // For a single-section document there are no SectionBreak markers, so `columns`
  // stays `computeColumns(section)` for the whole body — byte-identical to before.
  let columns = sectionColumnsFrom(0);
  let colIndex = 0;
  const colX = () => columns[colIndex].xPt;
  const colW = () => columns[colIndex].wPt;
  // ECMA-376 §17.10.1 — the active section's resolved header/footer set + titlePg,
  // tracked in lockstep with `columns` (reassigned at every SectionBreak). Stamped
  // on each element so the renderer picks the active section's header/footer per
  // page. `undefined` for the final section ⇒ the renderer's body-level fallback.
  let currentSectionHF = sectionHFFrom(0);
  // ECMA-376 §17.6.13 / §17.6.11 — the active section's page geometry, tracked in
  // lockstep with `columns`/`currentSectionHF` (reassigned at every SectionBreak).
  // Stamped on each element so the renderer sizes each page from its own section.
  // For a single-section document this stays the body-level geometry throughout.
  let currentSectionGeom = sectionGeomFrom(0);
  // ECMA-376 §17.6.4 / §17.18.79 — the CURRENT multi-column region's vertical
  // extent on the current page, in content-relative pt (0 = page content top,
  // i.e. the same frame as `y`). A region tiled into N newspaper columns is a
  // rectangle [colTopY, …]; every column STARTS at `colTopY`, not the page top.
  // For a section opened by a "continuous" section break the region begins
  // partway down the page (below the preceding single-column content), so
  // `colTopY > 0`. `maxColBottomY` accumulates the deepest column bottom reached
  // in the region so the NEXT section (after a continuous break) starts below
  // ALL columns rather than overprinting a taller one. Reset to the page top at
  // every real page open; re-seeded at each continuous section boundary. For a
  // single-section / single-column document `colTopY` stays 0 — behaviour-neutral.
  let colTopY = 0;
  let maxColBottomY = 0;
  // ECMA-376 §17.6.4 — newspaper column BALANCING target (content-relative pt per
  // column) for the CURRENT section, or null when the section fills greedily.
  // Word balances a continuous multi-column section that does NOT fill its page —
  // its content is split so all columns end at roughly the same height — but
  // leaves the FINAL (body) section greedy (it packs column 0 first). `balanceColH`
  // is the per-column height target; a non-last column breaks to the next column
  // once it reaches it (see maybeBalanceBreak). null ⇒ greedy fill to the page
  // bottom (single-column sections, the final section, or a section taller than
  // one page).
  let balanceColH: number | null = null;
  // Page-absolute pt of the current region top, stamped on elements so the paint
  // pass resets a column's cursor to the region top (front-loaded layout: the
  // renderer consumes this instead of independently deciding the column top).
  const colTopAbsPt = () => bodyTopPt() + colTopY;

  // Run `fn` with the measure state's content band temporarily re-pointed at the
  // CURRENT newspaper column (#513). The paint pass (renderBodyElements) sets
  // state.contentX/contentW = col.xPt/wPt × scale per element before resolving an
  // out-of-flow frame or floating table, because their placement math reads
  // state.contentX/contentW directly (frameXContainer for horzAnchor="text",
  // floatTableWrapSide for the wrap side). The measure pass must use the SAME band
  // so the FloatRect it registers (x, wrap side) matches what the renderer paints;
  // otherwise a horzAnchor="text" frame / table in a multi-column section diverges
  // between measure and paint. measureState.scale is 1, so colX()/colW() (pt) are
  // already the px-equivalent values the paint pass derives via col.xPt × scale.
  // contentX/contentW are saved and restored via try/finally so an early throw
  // inside `fn` cannot leak the narrowed band into subsequent elements.
  const withColumnBand = <T>(fn: () => T): T => {
    const savedX = measureState.contentX;
    const savedW = measureState.contentW;
    measureState.contentX = colX() * measureState.scale;
    measureState.contentW = colW() * measureState.scale;
    try {
      return fn();
    } finally {
      measureState.contentX = savedX;
      measureState.contentW = savedW;
    }
  };

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
  // like the real renderer does. floatParaSeq is reset together with floats so
  // the paraId採番 matches the renderer, which starts each page at 0 (a fresh
  // RenderState per page in renderDocumentToCanvas).
  measureState.y = bodyTopPt();
  measureState.floats = [];
  measureState.floatParaSeq = 0;
  // ECMA-376 §20.4.3.2/§20.4.3.5: a wp:anchor whose positionV relativeFrom is
  // page-level (page / margin / *Margin / column) is positioned independently
  // of its source-order anchoring paragraph — Word lays it out at page-open,
  // so paragraphs PRECEDING the anchor's paragraph on the same page still
  // wrap around it. We pre-register such floats at every page-start so the
  // paginator's per-paragraph height estimate matches the renderer (which
  // does the same pre-scan once per page in renderBodyElements). The
  // page-start body index is updated at every place that resets `floats`
  // (initial setup, newPage, pageBreak, sectionBreak/nextPage, +
  // pageBreakBefore). prescanFloatsFrom() reads it and rescans into the
  // (already cleared) measureState. See preRegisterPageFloats header.
  measureState.pageAnchorPrescanned = new Set();
  const prescanFloatsFrom = (idx: number): void => {
    measureState.pageAnchorPrescanned = new Set();
    preRegisterPageFloats(body, idx, measureState);
  };
  prescanFloatsFrom(0);
  // Footnote ids already reserved on the current page (so a paragraph that
  // references the same note twice doesn't double-count, and the renderer draws
  // each note once). Reset on every page flip.
  let pageNoteIds = new Set<string>();
  // Effective content height for the current page: the full text column minus the
  // footnote + tall-footer reserve at the BOTTOM and the tall-header reserve at the TOP
  // of THIS page (ECMA-376 §17.6.11). The body cursor is page-relative (0 = content
  // top); the header reserve also shifts the render-time body start down by the same
  // amount (renderDocumentToCanvas), so reducing the height here keeps the body within
  // the bottom margin while it begins below the header.
  // Clamp past the array end to the last entry: a tall header/footer shrinks every
  // page, so this (second) pass can have MORE pages than the (first) pass the reserves
  // were measured over. For a uniform header/footer the last measured reserve is the
  // right value for any extra page (and a first-page-only reserve's last entry is 0,
  // also correct). Empty ⇒ no reserve (the common no-overflow path).
  const reserveAt = (i: number): PageReserve =>
    pageReserves.length === 0 ? ZERO_RESERVE : (pageReserves[Math.min(i, pageReserves.length - 1)] ?? ZERO_RESERVE);
  const effContentH = () => {
    const r = reserveAt(pages.length - 1);
    return fullContentH() - (footnoteReservePt[pages.length - 1] ?? 0) - r.bottom - r.top;
  };
  const startPageBookkeeping = () => {
    footnoteReservePt[pages.length - 1] = 0;
    pageNoteIds = new Set<string>();
  };
  /** Open a new page (no-op when the current one is empty). `pageStartIdx`
   *  is the body index that becomes the new page's first laid-out item — the
   *  caller's current `i` (a paragraph being split / relocated / a pageBreak
   *  / a table being split). Used to pre-register page-level wrap floats
   *  (ECMA-376 §20.4.3.2/§20.4.3.5; see `prescanFloatsFrom`). */
  const newPage = (pageStartIdx: number) => {
    if (pages[pages.length - 1].length > 0) {
      pages.push([]);
      y = 0;
      colIndex = 0;
      // A fresh page: the section's columns span it from the content top.
      colTopY = 0;
      maxColBottomY = 0;
      balanceColH = null;
      prevPara = null;
      prevSpaceAfter = 0;
      measureState.y = bodyTopPt();
      // Floats are PAGE-scoped (ECMA-376 §20.4.2.x): a new page starts with a
      // clean float set. Columns of the SAME page share floats (see nextColumn).
      measureState.floats = [];
      measureState.floatParaSeq = 0;
      prescanFloatsFrom(pageStartIdx);
      startPageBookkeeping();
      // ECMA-376 §17.6.4: re-evaluate balancing for the content continuing on this
      // page. A multi-page continuous section fills its full pages greedily (both
      // columns), but its LAST page (the remainder now fitting on one page) is
      // balanced like a short section — so e.g. the tail of a 2-column body fills
      // both columns instead of packing column 0.
      setupBalancing(pageStartIdx);
    }
  };
  // Advance to the next newspaper column of the CURRENT page: reset the vertical
  // cursor to the column top but KEEP the page's floats (they are page-scoped, so
  // a full-width wrapTopAndBottom band still pushes down every column's first
  // line, and a square float keeps constraining the columns its x-range covers).
  // No new page is pushed and footnote reserve is untouched (same page).
  const nextColumn = () => {
    // Record the just-finished column's bottom, then drop the cursor back to the
    // REGION top (not the page top) — for a continuous mid-page section the next
    // column begins below the preceding single-column content (ECMA-376 §17.6.4).
    maxColBottomY = Math.max(maxColBottomY, y);
    colIndex++;
    y = colTopY;
    prevPara = null;
    prevSpaceAfter = 0;
    measureState.y = bodyTopPt() + colTopY;
  };
  // Overflow handler shared by element placement and paragraph/table splitting:
  // move to the next column if one remains on this page, otherwise to a new page.
  // `pageStartIdx` is forwarded to `newPage` so a fresh page pre-registers
  // page-level floats from the right body index (the element that triggered the
  // overflow); ignored for the same-page next-column path (floats are page-scoped).
  const nextColumnOrPage = (pageStartIdx: number) => {
    if (colIndex < columns.length - 1) nextColumn();
    else newPage(pageStartIdx);
  };

  // ECMA-376 §17.6.4 newspaper column balancing. Measure the single-column
  // content height (pt) of the section STARTING at `startIdx` — up to the next
  // section/page break — and whether a break terminates it. Uses a throwaway
  // clone of the measure state so the live cursor / floats are untouched. Mirrors
  // the per-element height the main loop computes (estimateParagraphHeight /
  // computeTableRowHeights at the column width, with paragraph spacing collapse),
  // but is an APPROXIMATION: line-level page splitting and floats are ignored,
  // which is fine for deciding a balance target. Returns `Infinity` when the
  // section can't be balanced as one block (an inner page break).
  const measureSectionColumnHeight = (
    startIdx: number,
    colWPt: number,
  ): { height: number; terminated: boolean } => {
    const ms: RenderState = { ...measureState, y: bodyTopPt(), floats: [], floatParaSeq: 0 };
    let total = 0;
    let prevAfter = 0;
    let prevP: DocParagraph | null = null;
    let terminated = false;
    for (let j = startIdx; j < body.length; j++) {
      const e = body[j];
      if (e.type === 'sectionBreak' || e.type === 'pageBreak') { terminated = true; break; }
      if (e.type === 'columnBreak') continue;
      if (e.type === 'paragraph') {
        const p = e as unknown as DocParagraph;
        if (p.pageBreakBefore) return { height: Infinity, terminated: false };
        if (p.framePr) continue; // frame is out of flow (adds no column height)
        const suppress = contextualSuppressed(prevP, p) || isContinuousSectionSpacer(j);
        const before = suppress ? 0 : p.spaceBefore;
        total += estimateParagraphHeight(ms, p, colWPt, suppress, 0) - Math.min(prevAfter, before);
        prevAfter = p.spaceAfter;
        prevP = p;
      } else if (e.type === 'table') {
        const t = e as unknown as DocTable;
        if (t.tblpPr) continue; // floating table: out of flow
        total += computeTableRowHeights(ms, t, colWPt).reduce((s, x) => s + x, 0);
        prevAfter = 0;
        prevP = null;
      }
    }
    return { height: total, terminated };
  };

  // Configure balancing for the section starting at `startIdx`. Word balances a
  // continuous multi-column section that does NOT fill its page (its content is
  // split so the columns end at roughly equal heights), but leaves greedy: a
  // single-column section, the FINAL (unterminated) section — Word packs column 0
  // there, matching the journal templates' last page — and any section taller
  // than the space available across its columns on this page.
  const setupBalancing = (startIdx: number): void => {
    balanceColH = null;
    const ncols = columns.length;
    if (ncols < 2) return;
    const { height, terminated } = measureSectionColumnHeight(startIdx, columns[0].wPt);
    if (!terminated || !Number.isFinite(height)) return; // final / page-breaking ⇒ greedy
    const avail = effContentH() - colTopY;
    if (avail <= 0 || height > ncols * avail) return; // spills past one page ⇒ greedy
    balanceColH = height / ncols;
  };

  // True when a whole element of height `fitH` should move to the next column to
  // keep a balanced section's columns even: balancing is active, the current
  // column is not the last (the last absorbs the remainder), it already holds
  // content, and the element would push it past the balanced target.
  const wantsBalanceBreak = (fitH: number): boolean =>
    balanceColH != null &&
    colIndex < columns.length - 1 &&
    y > colTopY &&
    y + fitH > colTopY + balanceColH;

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
  const anchoredFloatBottomOffset = (para: DocParagraph): number => {
    let maxBottom = 0;
    for (const run of para.runs) {
      if (run.type === 'image') {
        const img = run as unknown as ImageRun;
        if (!img.anchor || !img.anchorYFromPara) continue;
        // ECMA-376 §20.4.3.5: a `positionV relativeFrom="paragraph"/"line"` float
        // anchors against the paragraph's pre-spaceBefore TOP, regardless of wrap
        // mode (registerAnchorFloats + renderAnchorImages both use paragraphStartY).
        // So its bottom, measured from the paragraph top (the paginator's `y`), is
        // anchorYPt + height — no spaceBefore term.
        const bottom = (img.anchorYPt ?? 0) + img.heightPt;
        if (bottom > maxBottom) maxBottom = bottom;
      } else if (run.type === 'shape') {
        // An anchored shape with positionV relativeFrom="paragraph"/"line" is
        // kept on its anchor's page the same way an image is, and anchors at the
        // same pre-spaceBefore paragraph top. Take the height from resolveShapeBox
        // so sizeRelV / wgp-group scaling is honored. measureState is scale 1 (pt),
        // so box.h is already in pt.
        const shp = run as unknown as ShapeRun;
        if (!shp.anchorYFromPara) continue;
        const box = resolveShapeBox(shp, measureState, measureState.y);
        if (box.h <= 0) continue;
        const bottom = (shp.anchorYPt ?? 0) + box.h;
        if (bottom > maxBottom) maxBottom = bottom;
      }
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
      return estimateParagraphHeight(measureState, nxt as unknown as DocParagraph, colW(), false);
    }
    if (nxt.type === 'table') {
      return estimateTableHeight(measureState, nxt as unknown as DocTable, colW());
    }
    return 0;
  };

  // Stamp the active newspaper column (index + this section's geometry) on an
  // element and push it onto the current page. For single-column sections
  // colIndex is always 0 (the renderer treats 0/absent identically) and colGeom
  // equals the page-level columns, so this is behaviour-neutral there. colGeom
  // lets the renderer resolve the right section's column widths when two sections
  // share a page (a continuous break).
  const pushTagged = (el: PaginatedBodyElement) => {
    el.colIndex = colIndex;
    el.colGeom = columns;
    el.colTopPt = colTopAbsPt();
    el.sectionHF = currentSectionHF;
    el.sectionGeom = currentSectionGeom;
    pages[pages.length - 1].push(el);
  };

  // Balance the FIRST section's columns if it is a non-final multi-column section
  // that fits on page 1 (§17.6.4). Single-column / multi-page / final first
  // sections leave `balanceColH` null (greedy), so this is behaviour-neutral for
  // the common single-section document.
  setupBalancing(0);

  for (let i = 0; i < body.length; i++) {
    const el = body[i];
    if (el.type === 'columnBreak') {
      // ECMA-376 §17.3.1.20 <w:br w:type="column"/>: force the next column (or a
      // new page's first column when already in the last column — newPage() no-ops
      // on an empty page, so a column break in the last column of an as-yet-empty
      // page simply stays put). Page-start index = the body element AFTER the
      // break (i + 1) since the break itself emits no content on the new page.
      nextColumnOrPage(i + 1);
      continue;
    }
    if (el.type === 'pageBreak') {
      pages.push([]);
      y = 0;
      colTopY = 0;
      maxColBottomY = 0;
      balanceColH = null;
      prevPara = null;
      prevSpaceAfter = 0;
      measureState.y = bodyTopPt();
      measureState.floats = [];
      measureState.floatParaSeq = 0;
      // Pre-register page-level wrap floats from the next body element onward
      // (the pageBreak itself emits nothing on the new page).
      prescanFloatsFrom(i + 1);
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
    if (el.type === 'sectionBreak') {
      // ECMA-376 §17.6.x — a section boundary. Switch the active newspaper-column
      // geometry to the NEXT section (the one starting at i+1) and reset the
      // column index, then break per the section's ST_SectionMark (§17.18.79).
      // This is the per-section-columns fix: each section now lays out in its OWN
      // columns instead of every section inheriting the body-level section's.
      columns = sectionColumnsFrom(i + 1);
      colIndex = 0;
      // ECMA-376 §17.10.1 — the section starting at i+1 owns the following pages'
      // headers/footers (resolved from the NEXT marker, or the body-level set).
      currentSectionHF = sectionHFFrom(i + 1);
      // ECMA-376 §17.6.13 / §17.6.11 — and its page geometry (size + margins).
      currentSectionGeom = sectionGeomFrom(i + 1);
      // The break is governed by the UPCOMING section's start type (§17.6.22),
      // not this marker's own kind (the section it closes). See sectionKindFrom.
      // The sample-5 cover overprint that prompted the 0.66.1 hotfix is now fixed
      // at its real root: the parser emits a PageBreak after a "Cover Pages"
      // building block, so the continuous body after the cover lands on page 2
      // without forcing every nextPage→continuous boundary to break a page.
      const upcomingKind = sectionKindFrom(i + 1);
      if (upcomingKind === 'continuous') {
        // ECMA-376 §17.18.77 (ST_SectionMark) / §17.6.22 (type) — geometry residual:
        // reassigning `currentSectionGeom` above activates the incoming section's
        // page frame MID-PAGE (`bodyTopPt()`/`effContentH()` switch here). The spec
        // does NOT mandate promoting a continuous break to a new page when the page
        // size differs; it says the opposite — "continuous section breaks might not
        // specify certain page-level section properties, since they shall be
        // inherited from the following section" (§17.18.77 / §17.6.22). The only
        // spec-defined exception is a footnote reference on the break page (§17.11.14
        // ⇒ the new section begins on the following page). Word additionally defers a
        // differing page geometry to the next page in practice, but that is a runtime
        // behaviour, not a spec rule — left as a documented residual (see design doc),
        // NOT reverse-engineered here (root CLAUDE.md: spec-first, no runtime guessing
        // without approval).
        //
        // ECMA-376 §17.18.77 "continuous": NO page break — the next section's
        // content continues on the SAME page. It must start below the BOTTOM of
        // EVERY column of the section just ended (ECMA-376 §17.6.4), not at the
        // last-filled column's cursor: a 2-col section whose first column ran to
        // the page bottom while the second is short would otherwise let the new
        // (e.g. full-width) section overprint the taller column. `regionBottom`
        // is the deepest column reached; the new section's region then begins
        // there. (1-col→1-col, e.g. sample-5: maxColBottomY tracks no extra
        // column, so regionBottom == y — the sections simply stack, unchanged.)
        // Full column BALANCING of the ended section is a separate fidelity step;
        // this clears the overprint and the page-fill error regardless.
        const regionBottom = Math.max(maxColBottomY, y);
        y = regionBottom;
        measureState.y = bodyTopPt() + regionBottom;
        colTopY = regionBottom;
        maxColBottomY = regionBottom;
        prevPara = null;
        prevSpaceAfter = 0;
        // ECMA-376 §17.6.4: balance this continuous section's columns if it fits
        // on the current page below `regionBottom` (Word balances non-final
        // continuous multi-column sections; the final section stays greedy).
        setupBalancing(i + 1);
      } else {
        // nextPage (default) / oddPage / evenPage: start a new page (mirrors the
        // pageBreak path, including parity padding). A new page already resets
        // colIndex to 0 and clears page-scoped floats.
        pages.push([]);
        y = 0;
        colTopY = 0;
        maxColBottomY = 0;
        balanceColH = null;
        prevPara = null;
        prevSpaceAfter = 0;
        measureState.y = bodyTopPt();
        measureState.floats = [];
        measureState.floatParaSeq = 0;
        // Pre-register page-level wrap floats from the next body element onward
        // (the sectionBreak itself emits nothing on the new page).
        prescanFloatsFrom(i + 1);
        startPageBookkeeping();
        // Balance the new section's columns if it fits on its fresh page (§17.6.4).
        setupBalancing(i + 1);
        if (upcomingKind === 'oddPage' && pages.length % 2 === 0) {
          pages.push([]);
          startPageBookkeeping();
        } else if (upcomingKind === 'evenPage' && pages.length % 2 === 1) {
          pages.push([]);
          startPageBookkeeping();
        }
      }
      continue;
    }
    if (el.type === 'paragraph') {
      const para = el as unknown as DocParagraph;
      // `pageBreakBefore` opens a fresh page whose first item is THIS paragraph,
      // so the pre-scan should start at `i` (not i+1) to include its own
      // page-level floats. (registerAnchorFloats below will then skip the
      // pre-registered page-level ones via state.pageAnchorPrescanned.)
      if (para.pageBreakBefore) newPage(i);

      // A frame paragraph (ECMA-376 §17.3.1.11) is positioned out of flow: it
      // does not advance the page cursor and is not split. Register its wrap
      // float on the measureState so following paragraphs estimate around it,
      // and emit it onto the current page (the renderer draws it absolutely).
      // It does NOT advance y / measureState.y and leaves prevPara/spaceAfter
      // untouched so the following paragraph spaces against the paragraph
      // BEFORE the frame. (Pagination of a frame that itself overflows the page
      // bottom — moving the frame + its anchor text together — is Word runtime
      // behaviour not pinned by ECMA-376; see HEURISTIC note below. We keep the
      // minimal model: the frame stays with the anchor paragraph because it
      // adds no height here, so a normal break on the anchor paragraph carries
      // the wrap band implicitly. TODO(§17.3.1.11): once a self-contained frame
      // height/keep calc exists, drive an explicit keep-with-anchor here.)
      if (para.framePr) {
        const anchorH = frameAnchorLineHeightPx(body as PaginatedBodyElement[], el, measureState);
        // Resolve and register the frame's wrap float against the CURRENT column
        // band (#513), mirroring renderBodyElements which re-points
        // state.contentX/contentW per column before drawing a frame. For
        // hAnchor="text" the box x and exclusion x-range are column-relative
        // (frameXContainer reads contentX), so measure and paint agree.
        withColumnBand(() => {
          const box = resolveFrameBox(para, measureState, anchorH);
          registerFrameFloat(box, para.framePr!, measureState);
        });
        pushTagged(el as PaginatedBodyElement);
        continue;
      }
      const contextual = contextualSuppressed(prevPara, para);
      // An empty CONTINUOUS-section-break spacer drops only ITS OWN before (gap
      // collapses to prev.after, normally 0) — NOT the previous paragraph's after
      // (see isSectionBreakSpacerAt). contextualSpacing drops both. Stamp the
      // element so the paint pass (which gets per-page lists with the sectionBreak
      // already consumed, so it cannot re-detect the adjacency) applies the same
      // drop.
      const spacer = isContinuousSectionSpacer(i);
      if (spacer) (el as PaginatedBodyElement).sectionBreakSpacer = true;
      // A zero-before continuous-section spacer renders NO mark line box (Word's
      // section-mark collapse — see isCollapsedContinuousSpacer). Skip it entirely:
      // add no height and leave prevPara/prevSpaceAfter as the paragraph BEFORE the
      // spacer, so the next section's first paragraph spaces against it. The paint
      // pass mirrors this by skipping the stamped element. (Still pushed so the
      // per-page element sequence — and any section-geometry stamping keyed off it —
      // is unchanged.)
      if (isCollapsedContinuousSpacer(i)) {
        (el as PaginatedBodyElement).collapsedSpacer = true;
        pushTagged(el as PaginatedBodyElement);
        continue;
      }
      const suppressBefore = contextual || spacer;

      // An empty paragraph that immediately precedes a COLLAPSED continuous spacer
      // begins the section-break empty run, which Word renders FLUSH below the
      // preceding content (the section transition collapses upward): the previous
      // paragraph's spaceAfter is dropped too (sample-12: "[Format…]"'s 6pt after
      // vanishes, so "1. INTRODUCTION" sits at Word's 446pt rather than ~452pt).
      // Mirrors contextualSpacing's full drop of the previous after; no-op when the
      // spacer is NOT collapsed (sample-13, before=22 keeps normal flow).
      const leadsCollapsedRun = isInklessParagraph(para) && isCollapsedContinuousSpacer(i + 1);
      // Stamp it so the paint pass reads the decision here rather than re-deriving it
      // from per-page adjacency: the collapsed spacer this looks ahead to can fall on
      // the NEXT page's element list, where paint could not see it (lockstep).
      if (leadsCollapsedRun) (el as PaginatedBodyElement).leadsCollapsedRun = true;

      // Collapse with the previous paragraph's spaceAfter — Word takes
      // max(prev.after, this.before) between paragraphs, not the sum.
      const effectiveBefore = suppressBefore ? 0 : para.spaceBefore;
      // §17.3.1.9 contextualSpacing: same-style adjacent paragraphs drop BOTH the
      // previous after and this before (gap = 0), keeping the paginator's fill in
      // lockstep with the paint pass.
      const overlap = (contextual || leadsCollapsedRun) ? prevSpaceAfter : Math.min(prevSpaceAfter, effectiveBefore);
      y -= overlap;
      measureState.y -= overlap;

      // Register this paragraph's anchor-image floats on the measureState so
      // subsequent paragraphs estimate around them (text-wrap around images
      // adds lines that the float-unaware estimate would otherwise miss,
      // which caused page 2 of demo/sample-1 to spill past the bottom
      // margin). The renderer runs the same registerAnchorFloats call; by
      // mirroring it here the paginator sees the same layout.
      // Snapshot the float set + paraSeq BEFORE this paragraph registers, so a
      // relocation to the NEXT COLUMN (which keeps page floats) can roll back this
      // paragraph's own floats and re-register them at the new column position
      // without leaving stale duplicates. (A relocation to a new PAGE clears
      // floats wholesale via newPage(), so this snapshot is unused there.)
      const floatsBefore = measureState.floats.length;
      const floatSeqBefore = measureState.floatParaSeq;
      // ECMA-376 §20.4.3.5: a `positionV relativeFrom="paragraph"` float anchors
      // at the paragraph's TOP (pre-spaceBefore), so register it at measureState.y
      // BEFORE spaceBefore is folded in — matching the paint pass (paragraphStartY)
      // and renderAnchorImages (wrapNone). See registerAnchorFloats's call site in
      // renderParagraph for the spec rationale.
      const paragraphAnchorY = measureState.y;
      registerAnchorFloats(para, measureState, paragraphAnchorY);

      const h = estimateParagraphHeight(measureState, para, colW(), suppressBefore, colX());

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
          sum += footnoteReserveHeightPt(note, measureState, colW(), firstOnPage);
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
      const floatBottomOff = anchoredFloatBottomOffset(para);
      const floatOverflowsHere = floatBottomOff > 0 && y + floatBottomOff > effContentH();
      const floatFitsFresh = floatBottomOff > 0 && floatBottomOff <= effContentH();
      const breakForFloat = y > 0 && floatOverflowsHere && floatFitsFresh;
      // Does the content (+ keepNext look-ahead + newly-referenced footnote
      // bodies) overflow the space left on this page? spaceAfter is trailing
      // whitespace allowed to spill past the bottom margin, so it is excluded.
      const overflowsHere = y > 0 && y + needed > effContentH();
      // Relocate the WHOLE paragraph to a fresh page (rather than splitting its
      // lines) only when it must stay intact AND would then fit: keepLines
      // (§17.3.1.14 — all lines on one page), keepNext (§17.3.1.15 — stay with
      // the next block), or a footnote it carries (§17.11 — keep the body with
      // its reference). An ordinary paragraph is NOT relocated just for spilling
      // past the bottom; it is split at a line boundary by the path below
      // (Word's default behaviour). A keep-on-page float forces relocation too.
      const keepIntact = para.keepLines || needNext > 0 || addReservePt > 0;
      // A paragraph splits at line boundaries unless keepLines (§17.3.1.14) holds
      // and it still fits a page (a keepLines paragraph taller than a page must
      // split anyway). Used both for the balance whole-move below and the
      // page/balance split path further down.
      const splittable = !para.keepLines || h > effContentH();
      // ECMA-376 §17.6.4 column balancing. Move the WHOLE paragraph to the next
      // column when the current (non-last) column has reached the balanced target —
      // but ONLY for a paragraph that cannot be split at a line boundary
      // (keepLines). A splittable paragraph is instead split AT the balance target
      // by the split path below, so a long paragraph fills column 0 up to the
      // target and spills the remainder into column 1 (even columns), rather than
      // being shoved whole into the next column (which left column 0 nearly empty
      // and column 1 overfull, sample-12 p.2). No-op when balancing is off.
      const balanceBreak = wantsBalanceBreak(fitHeight) && !splittable;
      if (breakForFloat || balanceBreak || (overflowsHere && keepIntact && needed <= effContentH())) {
        const pagesBeforeRelocate = pages.length;
        // Relocating THIS paragraph to a fresh page: pre-scan from `i` so the
        // new page starts with THIS paragraph's own page-level floats counted.
        nextColumnOrPage(i);
        const movedToNewPage = pages.length > pagesBeforeRelocate;
        if (movedToNewPage) {
          // newPage() cleared measureState.floats AND reset floatParaSeq to 0, so
          // this is a REPLACE of the earlier register at the top of the loop (whose
          // floats were just discarded): this paragraph is now the first registrant
          // on the fresh page and gets paraId 0 — matching the renderer, which
          // re-registers from a fresh per-page state.
          registerAnchorFloats(para, measureState, measureState.y);
          // The references move to the new page; nothing was reserved there yet,
          // so the separator region still applies to the first footnote.
          if (haveFootnotes && newRefIds.length > 0) addReservePt = sumReserve(newRefIds);
        } else {
          // Moved to the NEXT COLUMN of the same page. Page floats persist, but
          // this paragraph's own floats were anchored at the previous column's Y —
          // roll them back and re-register against the new column top so wrap
          // estimates for this paragraph (and later ones) use the right band.
          measureState.floats.length = floatsBefore;
          measureState.floatParaSeq = floatSeqBefore;
          registerAnchorFloats(para, measureState, measureState.y);
        }
      }

      // ECMA-376 places no "a paragraph must fit on one page" requirement — Word
      // splits a paragraph at line boundaries whenever its content doesn't fit in
      // the space left on the page. Walk the laid-out lines and emit one slice
      // per page. The old `&& h > pageContentH * 0.5` guard (split only
      // paragraphs taller than half a page) was a heuristic that suppressed every
      // ordinary line-level break: a body/list paragraph that didn't fit was
      // neither split nor (in the default case) relocated, so it overflowed and
      // was clipped (sample-9 page-4). Gate on splittability instead — keepLines
      // (§17.3.1.14) forbids splitting unless the paragraph cannot fit on any
      // page (taller than the full content height), where it must split anyway.
      const pageContentH = effContentH();
      // ECMA-376 §17.6.4 — in a BALANCED newspaper section every non-last column
      // fills only to the balance target (the region top + height/ncols); the last
      // column absorbs the remainder. Cap the split at that target so a paragraph
      // taller than one balanced column is split AT the balance point instead of
      // packed whole into column 0 (which left column 0 visibly fuller than column
      // 1, sample-12 p.2). Unbalanced / greedy sections (balanceColH == null) and
      // the last column return the page bottom, so this is behaviour-neutral there.
      // Read live (colIndex / colTopY / balanceColH change as the split advances).
      const columnBottomLimit = (): number =>
        balanceColH != null && colIndex < columns.length - 1
          ? Math.min(effContentH(), colTopY + balanceColH)
          : effContentH();
      const remainingH = columnBottomLimit() - y;
      if (fitHeight > remainingH && splittable) {
        const placed = splitParagraphAcrossPages(
          measureState, para, colW(), suppressBefore, colX(),
          y, pageContentH, pages,
          // Overflow during the split advances to the next column first, then a
          // new page (newspaper fill). The just-filled column's bottom is folded
          // into `maxColBottomY` so a following continuous section clears the
          // deepest column (ECMA-376 §17.6.4). Each slice is tagged with the
          // column it landed in via the colIndex thunk, plus this section's column
          // geometry (constant — a paragraph never spans a section boundary). The
          // continuation slice belongs to THIS paragraph, so the new page's
          // pre-scan starts at `i` (this paragraph index).
          (filledColBottom: number) => { maxColBottomY = Math.max(maxColBottomY, filledColBottom); nextColumnOrPage(i); },
          () => colIndex,
          columns,
          () => colTopY,
          columnBottomLimit,
          () => currentSectionHF,
          () => currentSectionGeom,
        );
        // After splitting, `y` is the bottom of the last slice in the
        // current column (continues for the LAST slice; intermediate slices
        // filled their column/page exactly, so the break callback ran between
        // them).
        y = placed.endY;
        measureState.y = bodyTopPt() + placed.endY;
        // A split footnote-bearing paragraph reserves on the page where it
        // ends. Rare; the separator region re-applies if that page had none.
        if (haveFootnotes && newRefIds.length > 0) {
          newRefIds = newRefIds.filter((id) => !pageNoteIds.has(id));
          addReservePt = sumReserve(newRefIds);
        }
      } else {
        pushTagged(el as PaginatedBodyElement);
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

      // ECMA-376 §17.4.57 `<w:tblpPr>`: a floating table is positioned OUT OF
      // FLOW (like a frame paragraph, §17.3.1.11). It does NOT advance the page
      // cursor and is not split. Register its wrap float on the measureState so
      // following paragraphs estimate around it, then emit it onto the current
      // page (the renderer draws it absolutely). prevPara/spaceAfter are left
      // untouched so the following paragraph spaces against the paragraph BEFORE
      // the table.
      //
      // ヒューリスティック、§17.4.57 TODO: floating table page-fit is Word runtime
      // behavior, not spec-defined. Minimal: keep on current page (no break /
      // relocation across pages here; only the in-page wrap band is modeled).
      if (tbl.tblpPr) {
        // #513: re-point measureState.contentX/contentW at the CURRENT newspaper
        // column for the duration of the placement so the paginator estimates the
        // wrap band against the SAME column band the paint pass paints into
        // (renderBodyElements sets state.contentX/contentW = col.xPt/wPt × scale
        // per element). For a horzAnchor="text" floating table inside a
        // MULTI-COLUMN section, computeFloatTableBox (frameXContainer) and
        // floatTableWrapSide both read contentX/contentW, so without this the
        // measured box x and wrap side would diverge from the rendered ones.
        // measureState.scale is 1, so colW() (pt) equals the px content width
        // computeTableLayout expects (it re-divides by scale internally); colX()
        // (pt) is likewise the column's page-absolute left edge.
        withColumnBand(() => {
          const cW = colW() * measureState.scale;
          const layout = computeTableLayout(tbl, cW, measureState);
          const tableH = layout.rowHeights.reduce((s, x) => s + x, 0);
          const box = computeFloatTableBox(tbl.tblpPr!, measureState, measureState.y, layout.tableW, tableH);
          const side = floatTableWrapSide(box, measureState);
          registerTableFloat(box, tbl.tblpPr!, measureState, side, tbl.overlap !== 'never');
        });
        pushTagged(el as PaginatedBodyElement);
        continue;
      }

      // Tables in a multi-column section are sized to the column width, not the
      // full content band.
      const rowHs = computeTableRowHeights(measureState, tbl, colW());
      const h = rowHs.reduce((s, x) => s + x, 0);
      // Footnote references inside table cells are not folded into the reserve
      // (the per-page reserve is driven by body paragraphs); they still draw at
      // page bottom via the renderer's page scan. effContentH() respects any
      // reserve already accumulated on this page.
      const tableContentH = effContentH();
      if (h > tableContentH) {
        // Taller than a full column: split row-by-row so the overflow continues
        // into the next column / page instead of being clipped (ECMA-376 table
        // pagination). Tables that fit keep the simple place-whole path below.
        const endY = splitTableAcrossPages(
          tbl, rowHs, y, tableContentH, pages,
          // Table slices belong to THIS table element, so the new page's
          // pre-scan starts at `i` (this table's body index). The just-filled
          // column bottom folds into `maxColBottomY` so a following continuous
          // section clears the deepest column (ECMA-376 §17.6.4).
          (filledColBottom: number) => { maxColBottomY = Math.max(maxColBottomY, filledColBottom); nextColumnOrPage(i); },
          () => colIndex,
          columns,
          () => colTopY,
          bodyTopPt(),
          () => currentSectionHF,
          () => currentSectionGeom,
        );
        y = endY;
        measureState.y = bodyTopPt() + endY;
      } else {
        // §17.6.4 column balancing (wantsBalanceBreak) OR a table that doesn't fit
        // the rest of this column ⇒ advance to the next column / page.
        if (wantsBalanceBreak(h) || y + h > tableContentH) nextColumnOrPage(i);
        pushTagged(el as PaginatedBodyElement);
        y += h;
        measureState.y += h;
      }
      prevPara = null;
    }
  }
  return pages;
}

/** ECMA-376 §17.6.11 (pgMar/@bottom): the main-document text bottom is placed at the
 *  GREATER of the bottom margin and the footer's extent — a footer taller than the
 *  bottom-margin allowance pushes content up. Returns the pt by which a footer of
 *  height `footerH` overflows the bottom margin and must be reserved (content ends at
 *  the footer's top, `footerDistance + footerH` from the page bottom). A NEGATIVE
 *  bottom margin means the text is measured from the page bottom regardless of the
 *  footer — it overlaps the footer — so nothing is reserved. Unit-agnostic: pass all
 *  three args in one unit (pt for the pagination reserve, px for paint-time footnote
 *  clearance). */
function footerOverflowPt(footerH: number, marginBottom: number, footerDistance: number): number {
  if (marginBottom < 0) return 0;
  return Math.max(0, footerDistance + footerH - marginBottom);
}

/** ECMA-376 §17.6.11 (pgMar/@top): the SYMMETRIC twin of footerOverflowPt. The
 *  main-document text TOP is placed at the GREATER of the top margin and the header's
 *  extent — a header taller than the top-margin allowance pushes content DOWN. Returns
 *  the pt by which a header of height `headerH` overflows the top margin and must be
 *  reserved (content starts at the header's bottom, `headerDistance + headerH` from the
 *  page top). A NEGATIVE top margin means the text is measured from the page top
 *  regardless of the header — it overlaps the header — so nothing is reserved.
 *  Unit-agnostic: pass all three args in one unit (pt for the pagination reserve, px
 *  for the paint-time body start). */
function headerOverflowPt(headerH: number, marginTop: number, headerDistance: number): number {
  if (marginTop < 0) return 0;
  return Math.max(0, headerDistance + headerH - marginTop);
}

/** ECMA-376 §17.6.11 (pgMar/@top,@bottom): the body text's distance (pt) from the page
 *  edge is the margin's MAGNITUDE. A non-negative margin insets the body, and the
 *  header/footer is reserved against it (header/footerOverflowPt push the body further
 *  in). A NEGATIVE margin measures the body |margin| from the page edge "regardless of
 *  the header/footer ... and therefore shall overlap the header/footer text" — the
 *  spec's w:top="-720"/w:bottom="-720" examples place the body ½ inch (|margin|) inside
 *  the page edge while overlapping the running head/foot (those functions then reserve
 *  0). Either way the body's inset from the page edge is |margin|. Math.abs is identity
 *  for the non-negative common case, so this only changes negative-margin documents. */
function bodyMarginInsetPt(margin: number): number {
  return Math.abs(margin);
}

/** ECMA-376 §17.6.13 (pgSz) + §17.6.11 (pgMar) — project a {@link SectionProps}
 *  onto the smaller {@link SectionGeom} view (page size + margins + header/footer
 *  distances, dropping the section's title-page/columns/etc.). Used to seed the
 *  body-level fallback geometry in `computePages` (bodySectionGeom); a private
 *  module helper so the boundary-promotion path can reuse the SAME 8-field
 *  projection when comparing an incoming section's geometry to the ending one. */
function sectionGeomOf(s: SectionProps): SectionGeom {
  return {
    pageWidth: s.pageWidth,
    pageHeight: s.pageHeight,
    marginTop: s.marginTop,
    marginRight: s.marginRight,
    marginBottom: s.marginBottom,
    marginLeft: s.marginLeft,
    headerDistance: s.headerDistance,
    footerDistance: s.footerDistance,
  };
}

/** Resolve the footer that applies to `pageIndex` with the §17.10.1/§17.10.6
 *  first/even/default precedence (resolvePageSection + pickHeaderFooter), or null if
 *  none. One selection shared by the reserve pass, the footer paint, and the footnote
 *  clearance so all three size and place the SAME footer. */
function resolvePageFooter(
  pages: PaginatedBodyElement[][],
  pageIndex: number,
  doc: DocxDocumentModel,
): HeaderFooter | null {
  const ps = resolvePageSection(pages, pageIndex, doc);
  return pickHeaderFooter(
    ps.footers, ps.isFirstPageOfSection, pageIndex % 2 === 1, ps.titlePage, doc.section.evenAndOddHeaders,
  );
}

/** Resolve the header that applies to `pageIndex` with the same §17.10.1/§17.10.6
 *  first/even/default precedence as resolvePageFooter. One selection shared by the
 *  reserve pass and the header paint so both size and place the SAME header. */
function resolvePageHeader(
  pages: PaginatedBodyElement[][],
  pageIndex: number,
  doc: DocxDocumentModel,
): HeaderFooter | null {
  const ps = resolvePageSection(pages, pageIndex, doc);
  return pickHeaderFooter(
    ps.headers, ps.isFirstPageOfSection, pageIndex % 2 === 1, ps.titlePage, doc.section.evenAndOddHeaders,
  );
}

/** Below this (pt) a header's/footer's overflow is sub-point noise — skip the second
 *  pagination pass when no page's header or footer overflows by at least this much. */
const MIN_MARGIN_OVERFLOW_PT = 0.5;

/**
 * ECMA-376 §17.6.11 (pgMar/@bottom) — per-page pt to reserve at the bottom of the
 * content area for a footer taller than its bottom-margin allowance. The main text
 * bottom sits at the greater of the bottom margin and the footer extent
 * (`footerDistance + footerHeight`); when the footer extent wins, content must clear
 * it (Word never lays body text over a footer). A footer that fits the margin — or a
 * negative bottom margin (§17.6.11: text then overlaps the footer) — reserves 0. See
 * footerOverflowPt for the exact rule.
 */
function computeFooterReserves(
  pages: PaginatedBodyElement[][],
  doc: DocxDocumentModel,
  measure: RenderState,
): number[] {
  return pages.map((_unused, pageIdx) => {
    const footer = resolvePageFooter(pages, pageIdx, doc);
    if (!footer) return 0;
    const footerH = measureHeaderFooterHeight(footer, measure); // pt (measure is scale 1)
    // ECMA-376 §17.6.11 — margins/distances are PER-SECTION: read this page's stamped
    // geometry (body-level fallback only for a truly empty page). Residual: footer
    // CONTENT is still measured at the body-level width (`measure` is built once from
    // doc.section); per-section measure width is a follow-up — it matters only when
    // mixed page WIDTHS meet a wrapping footer.
    const g = pages[pageIdx]?.[0]?.sectionGeom;
    return footerOverflowPt(
      footerH,
      g?.marginBottom ?? doc.section.marginBottom,
      g?.footerDistance ?? doc.section.footerDistance,
    );
  });
}

/**
 * ECMA-376 §17.6.11 (pgMar/@top) — the SYMMETRIC twin of computeFooterReserves: the
 * per-page pt to reserve at the TOP of the content area for a header taller than its
 * top-margin allowance. The main text top sits at the greater of the top margin and the
 * header extent (`headerDistance + headerHeight`); when the header extent wins, the body
 * starts at the header's bottom (Word never lays body text over a header). A header that
 * fits the margin — or a negative top margin (§17.6.11: text then overlaps the header) —
 * reserves 0. See headerOverflowPt for the exact rule.
 */
function computeHeaderReserves(
  pages: PaginatedBodyElement[][],
  doc: DocxDocumentModel,
  measure: RenderState,
): number[] {
  return pages.map((_unused, pageIdx) => {
    const header = resolvePageHeader(pages, pageIdx, doc);
    if (!header) return 0;
    const headerH = measureHeaderFooterHeight(header, measure); // pt (measure is scale 1)
    // ECMA-376 §17.6.11 — margins/distances are PER-SECTION: read this page's stamped
    // geometry (body-level fallback only for a truly empty page). Residual: header
    // CONTENT is still measured at the body-level width (`measure` is built once from
    // doc.section); per-section measure width is a follow-up — it matters only when
    // mixed page WIDTHS meet a wrapping header.
    const g = pages[pageIdx]?.[0]?.sectionGeom;
    return headerOverflowPt(
      headerH,
      g?.marginTop ?? doc.section.marginTop,
      g?.headerDistance ?? doc.section.headerDistance,
    );
  });
}

/**
 * Paginate with header/footer-height awareness (ECMA-376 §17.6.11). Pass 1 paginates
 * without reservation; a header or footer taller than its margin allowance overflows
 * the content area (computeHeaderReserves reserves at the top, computeFooterReserves at
 * the bottom, measured per page) and both are fed into a second pass so body content
 * never overlaps either. When nothing overflows — the common case — pass 1 is returned
 * unchanged. Shared by the main render path and the worker so the two can never
 * paginate differently.
 *
 * One re-pass is used (not iterated to a fixpoint), with two bounded approximations:
 *  - Page-count growth: a tall reserve shrinks every page, so pass 2 can have MORE pages
 *    than the pass-1 reserve array covers. computePages clamps a past-the-end page to the
 *    last reserve — exact for a uniform header/footer (every page the same) and for a
 *    first-page-only reserve (its trailing entry is 0). An even/odd header of DIFFERING
 *    height on a grown page is the residual gap (exotic, untested).
 *  - Page-index shift: a CONTINUOUS section whose tall first-page header/footer's page
 *    index moves when the reserve repacks content is left for a fixpoint pass if a real
 *    document ever needs it (sample-13's masthead begins a PAGE-STARTING break, so its
 *    index is identical in both passes).
 * The reserve shrinks the content height and shifts the body start; it does NOT re-anchor
 * page-level floats (wrapTopAndBottom / page-anchored), which still measure from the bare
 * top margin. No bundled document pairs a tall header/footer with such a float; if one
 * arises, shift measureState.y by the reserve at each page open to keep floats in frame.
 */
function paginateWithHeaderFooterReserve(
  doc: DocxDocumentModel,
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  fontFamilyClasses: Record<string, string>,
  kinsoku: KinsokuRules,
  footnotes: DocNote[],
): PaginatedBodyElement[][] {
  // §17.15.1.25 — resolve once here so both pagination passes and the
  // reserve-measure state share the document's automatic tab interval.
  const defaultTabPt = resolveDefaultTabPt(doc.settings);
  const pass1 = computePages(doc.body, doc.section, ctx, fontFamilyClasses, kinsoku, footnotes, [], defaultTabPt);
  const measure = buildMeasureState(ctx, doc.section, fontFamilyClasses, kinsoku, documentHasEastAsian(doc.body), defaultTabPt);
  const footerReserves = computeFooterReserves(pass1, doc, measure);
  const headerReserves = computeHeaderReserves(pass1, doc, measure);
  const overflows = (rs: number[]): boolean => rs.some((r) => r > MIN_MARGIN_OVERFLOW_PT);
  if (!overflows(footerReserves) && !overflows(headerReserves)) return pass1;
  const pageReserves = pass1.map((_unused, i): PageReserve => ({
    top: headerReserves[i] ?? 0,
    bottom: footerReserves[i] ?? 0,
  }));
  return computePages(doc.body, doc.section, ctx, fontFamilyClasses, kinsoku, footnotes, pageReserves, defaultTabPt);
}

/** Paginate with a throwaway measure context (a fresh OffscreenCanvas, scale 1).
 *  Pagination must use the same fontFamilyClasses + kinsoku rules as the render
 *  path, otherwise line-break decisions (and thus page breaks) diverge between
 *  measurement and paint (ECMA-376 §17.15.1.58–.60). Shared by the main-thread
 *  DocxDocument and the render worker so the two modes can never paginate
 *  differently.
 *
 *  FONT-LOADING PRECONDITION: the OffscreenCanvas measurement here uses whatever
 *  fonts are loaded AT CALL TIME. Callers must ensure font loading has completed
 *  (e.g. await `document.fonts.ready` / the relevant `FontFace.load()`) before
 *  paginating — paginating against fallback metrics and painting after the real
 *  fonts arrive yields stale page breaks. This has always been true for the
 *  lineSlice indices; since the compute-once reuse (Phase 4-1 B2 Stage 1) it
 *  also covers the stamped line GEOMETRY: at paint scale 1 the renderer reuses
 *  the lines measured here verbatim (renderParagraph's reuse gate assumes
 *  measure-time and paint-time text metrics are identical), where the old
 *  recompute path would at least have re-wrapped the within-page text under the
 *  late-loaded fonts. */
export function paginateDocument(doc: DocxDocumentModel): PaginatedBodyElement[][] {
  const ctx = new OffscreenCanvas(1, 1).getContext('2d');
  if (!ctx) return [doc.body];
  return paginateWithHeaderFooterReserve(
    doc,
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
  docEastAsian = false,
  // §17.15.1.25 — threaded so the measure pass uses the SAME automatic tab
  // interval as the draw pass; defaults to the spec absent value when no doc.
  defaultTabPt: number = DEFAULT_TAB_PT,
): RenderState {
  return {
    ctx,
    scale: 1,
    dpr: 1,
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
    // §17.6.11: the measure state's marginTop is the BODY-LEVEL body inset (|margin|).
    // Per-section colTopPt stamps no longer read this field directly: the split
    // functions derive the region top from the threaded `tagSectionGeom` closure
    // (`bodyMarginInsetPt(tagSectionGeom().marginTop)`), matching pushTagged's
    // `bodyTopPt()` per-section convention. This body-level value is only the
    // single-section-equivalent fallback (identical when there is one section) and
    // still seeds contentW/pageH below. Never the raw sign. Identity for non-negative.
    marginTop: bodyMarginInsetPt(section.marginTop),
    marginBottom: bodyMarginInsetPt(section.marginBottom),
    pageWidth: section.pageWidth,
    floats: [],
    floatParaSeq: 0,
    docGrid: {
      type: section.docGridType ?? null,
      linePitchPt: section.docGridLinePitch ?? null,
      charSpacePt: section.docGridCharSpace != null ? section.docGridCharSpace / 4096 : null,
    },
    docEastAsian,
    fontFamilyClasses,
    kinsoku,
    defaultTabPt,
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
  // Float-wrap windows are evaluated at the paragraph's content left edge =
  // contentX + physical left indent, matching renderParagraph. The left/right
  // indents swap to physical sides in a bidi paragraph (§17.3.1.12 / Part 4
  // §14.11.2), so use the bidi-resolved left indent; otherwise an indented / RTL
  // paragraph wrapping a square float would measure the gap at the wrong X and
  // diverge from the paint pass.
  const paraX = paraXPt + (para.bidi === true ? indRight : indLeft);
  const segs = buildSegments(para.runs, state);
  // Word renders ruby paragraphs with consistent line spacing — every line
  // in a paragraph that carries ANY furigana snaps to the same pitch
  // multiple, otherwise mixed-ruby paragraphs jitter and pagination drifts.
  const paraHasRuby = paragraphHasRuby(para);
  const grid = paraGrid(para, state);
  // Mirror renderParagraph's vertical advancement EXACTLY so the paginator's
  // page-fill tracker matches where the renderer actually draws each line. With
  // anchor floats active a paragraph's text (and even an empty paragraph mark)
  // is pushed BELOW the float band (ECMA-376 §20.4.2.x; resolveLineFloatWindow /
  // skipPastTopAndBottom), and that vertical displacement — not just the line
  // heights — consumes page space. The previous estimate summed only the line
  // heights, so the paginator under-counted full-width-float pages and packed
  // far too much onto them (sample-9 page-4: the photo block displaced the body
  // text, but the paginator ignored the gap and let the bullet list spill past
  // the bottom margin). Reproduce the renderer's cursor walk:
  //   spaceBefore → skipPastTopAndBottom → per line: max(topY) then += lineH
  //   (empty/anchor-only: flow the mark line below the band) → spaceAfter.
  // When no floats are active skipPastTopAndBottom is a no-op and no line carries
  // a topY, so this collapses to spaceBefore + Σ lineHeights + spaceAfter — the
  // exact previous value, leaving float-free documents unchanged.
  const hasFloats = state.floats.length > 0;
  const startY = state.y;
  let cursor = startY + (suppressSpaceBefore ? 0 : para.spaceBefore);
  if (hasFloats) cursor = skipPastTopAndBottom(cursor, state.floats);
  const flowMarkLine = (): void => {
    // Empty / anchor-only paragraph: one paragraph-mark line box, flowed below a
    // full-width float band exactly like renderEmptyMarkParagraph.
    if (hasFloats) {
      const win = resolveLineFloatWindow(cursor, paragraphMarkEmPx(para, 1), 10, paraX, paraW, state.floats);
      if (win.topY > cursor) cursor = win.topY;
    }
    cursor += paragraphMarkLineHeight(para, 1, grid, paraHasRuby, state.docEastAsian, state.ctx, state.fontFamilyClasses);
  };
  if (segs.length === 0) {
    flowMarkLine();
  } else {
    // Same WrapLayoutCtx the renderer uses (startPageY is the post-spaceBefore,
    // post-skip top — matching renderParagraph) so the laid-out line `topY`s
    // and line count agree with the paint pass.
    const wrapCtx: WrapLayoutCtx | undefined = hasFloats ? {
      startPageY: cursor,
      paraX,
      floats: state.floats,
      lineBoxH: (a, d, _h, is) => lineBoxHeight(para.lineSpacing, a, d, 1, grid, paraHasRuby, is ?? 0, paragraphIsEastAsian(para)),
      pageH: state.pageH,
      markEmPx: paragraphMarkEmPx(para, 1),
    } : undefined;
    const lines = layoutLines(state.ctx, segs, paraW, para.indentFirst, 1, para.tabStops, wrapCtx, state.fontFamilyClasses, indLeft, state.kinsoku, gridCharDeltaPx(grid, 1), state.defaultTabPt);
    if (lines.length === 0) {
      // Anchor-only paragraph: no inline content, but the paragraph mark still
      // occupies one (possibly flowed) line (§17.3.1.29).
      flowMarkLine();
    } else if (paraHasRuby) {
      // Word uses the same line height for every line in a ruby paragraph,
      // snapped to an integer docGrid pitch.
      const uniform = snapParaLineToGrid(
        Math.max(0, ...lines.map(l => lineBoxHeight(para.lineSpacing, l.ascent, l.descent, 1, grid, true, l.intendedSingle, paragraphIsEastAsian(para)))),
        grid,
        1,
      );
      for (const l of lines) {
        if (l.topY !== undefined && l.topY > cursor) cursor = l.topY;
        cursor += uniform;
      }
    } else {
      for (const l of lines) {
        if (l.topY !== undefined && l.topY > cursor) cursor = l.topY;
        cursor += lineBoxHeight(para.lineSpacing, l.ascent, l.descent, 1, grid, false, l.intendedSingle, paragraphIsEastAsian(para));
      }
    }
  }
  cursor += para.spaceAfter;
  return cursor - startY;
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

/** Code points whose presence marks a line as East Asian for docGrid line-cell
 *  rounding: CJK symbols/punctuation, Hiragana, Katakana, CJK Unified +
 *  Extension A, compatibility ideographs, Hangul, and fullwidth forms. Content
 *  test only — not a font-name heuristic (cf. packages/docx/CLAUDE.md). */
const EAST_ASIAN_RE =
  /[ᄀ-ᇿ⺀-⿟　-〿぀-ヿ㄰-㆏㐀-䶿一-鿿ꥠ-꥿가-퟿豈-﫿＀-￯]/u;

/** ECMA-376 §17.6.5 docGrid line-cell rounding (see lineBoxHeight) applies to
 *  East Asian lines. A paragraph counts as East Asian when any of its text runs
 *  carries East Asian characters. (Empty / anchor-only paragraphs carry no text;
 *  their mark line is handled separately by paragraphMarkLineHeight.) */
function paragraphIsEastAsian(para: DocParagraph): boolean {
  for (const run of para.runs) {
    if (run.type === 'text' && EAST_ASIAN_RE.test((run as unknown as DocxTextRun).text)) return true;
  }
  return false;
}

/** Whether the document body contains any East Asian text. An empty / anchor-
 *  only paragraph mark carries no text to classify, so its docGrid cell rounding
 *  (paragraphMarkLineHeight) is gated on this document-level signal: in an East
 *  Asian document the paragraph default is an East Asian font and its mark snaps
 *  to whole grid cells; a purely Latin document (e.g. demo/sample-1) keeps the
 *  natural single-cell mark. Recurses into table cells. */
function documentHasEastAsian(body: BodyElement[]): boolean {
  for (const el of body) {
    if (el.type === 'paragraph') {
      if (paragraphIsEastAsian(el as unknown as DocParagraph)) return true;
    } else if (el.type === 'table') {
      for (const row of (el as unknown as DocTable).rows) {
        for (const cell of row.cells) {
          if (documentHasEastAsian(cell.content)) return true;
        }
      }
    }
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
  /** Advance to the next column / page. Receives the bottom (content-relative pt)
   *  the just-filled column reached, so the caller can track the deepest column
   *  of a multi-column region (ECMA-376 §17.6.4) for the following section. */
  newPage: (filledColBottom: number) => void,
  /** ECMA-376 §17.6.4 — current newspaper column index, read AFTER each
   *  `newPage()` (which may advance the column). When provided, every emitted
   *  slice is tagged with the column it landed in so the renderer flows it in the
   *  right column. Omitted (single-column / direct unit tests) ⇒ no tag. */
  tagColIndex?: () => number,
  /** ECMA-376 §17.6.4 — the current SECTION's column geometry. A paragraph is
   *  never split across a section boundary, so this is constant for all slices;
   *  stamped so the renderer resolves the slice's column against the right
   *  section. Omitted ⇒ the renderer uses the page-level columns. */
  colGeom?: ColumnGeom[],
  /** ECMA-376 §17.6.4 — content-relative pt of the current column-region TOP,
   *  read AFTER each `newPage()`. A continuation column of a continuous mid-page
   *  section restarts here (below the preceding single-column content), not at
   *  the page top. Omitted ⇒ the page content top (0). */
  columnTop?: () => number,
  /** ECMA-376 §17.6.4 — the content-relative BOTTOM (pt) the current column may
   *  fill to, read AFTER each `newPage()`. For a balanced newspaper section this
   *  is the balance target (`colTop + height/ncols`) on every non-last column, so
   *  a single paragraph taller than one balanced column is split at the balance
   *  point rather than packed into column 0; the last column (and any unbalanced
   *  section) is uncapped at the page content bottom. Omitted ⇒ `contentH` (the
   *  page bottom) — behaviour-neutral for the single-column / greedy paths. */
  columnBottom?: () => number,
  /** ECMA-376 §17.10.1 — the active SECTION's resolved header/footer set. A
   *  paragraph never spans a section boundary, so this is constant for all slices;
   *  stamped so the renderer picks the right section's header/footer per page.
   *  Omitted ⇒ the renderer's body-level fallback. */
  tagSectionHF?: () => PaginatedBodyElement['sectionHF'],
  /** ECMA-376 §17.6.13 / §17.6.11 — the active SECTION's page geometry (size +
   *  margins). A paragraph never spans a section boundary, so this is constant for
   *  all slices; stamped so the renderer sizes each page from the right section.
   *  Omitted ⇒ the renderer's body-level fallback. */
  tagSectionGeom?: () => SectionGeom,
): { endY: number } {
  const colTop = columnTop ?? (() => 0);
  const colBot = columnBottom ?? (() => contentH);
  const stamp = (el: PaginatedBodyElement): PaginatedBodyElement => {
    if (tagColIndex) el.colIndex = tagColIndex();
    if (colGeom) el.colGeom = colGeom;
    // Front-loaded layout: stamp the region top (page-absolute pt) so the paint
    // pass resets this slice's column cursor to it instead of the page top. The
    // top inset is PER-SECTION (§17.6.11): a slice in a mid-body section with a
    // different marginTop must anchor against ITS section, not the body-level
    // `measureState.marginTop` (frozen at buildMeasureState). `tagSectionGeom` is
    // the active section's geometry (constant across a paragraph's slices), so
    // `bodyMarginInsetPt(tagSectionGeom().marginTop)` matches pushTagged's
    // `bodyTopPt()` convention exactly. For a single-section document this equals
    // `measureState.marginTop` (both are `bodyMarginInsetPt(section.marginTop)`)
    // — value-identical. Falls back to the body-level inset when no geom thunk.
    if (columnTop) {
      const topInset = tagSectionGeom
        ? bodyMarginInsetPt(tagSectionGeom().marginTop)
        : measureState.marginTop;
      el.colTopPt = topInset + colTop();
    }
    if (tagSectionHF) el.sectionHF = tagSectionHF();
    if (tagSectionGeom) el.sectionGeom = tagSectionGeom();
    return el;
  };
  const indLeft = para.indentLeft;
  const indRight = para.indentRight;
  const paraW = Math.max(1, contentWPt - indLeft - indRight);
  // Mirror renderParagraph's paragraph-content left edge: contentX + the
  // physical left indent (ECMA-376 §17.3.1.12; the left/right indents swap to
  // physical sides in a bidi paragraph, Part 4 §14.11.2). Using the bare margin
  // here would evaluate square-float wrap windows at the wrong X for indented /
  // RTL paragraphs, re-introducing a paginate/render disagreement.
  const physLeftInd = para.bidi === true ? indRight : indLeft;
  const paraX = marginLeftPt + physLeftInd;
  // A paragraph with no layoutable inline lines (literally empty, or only
  // wrap-float anchors) is a single paragraph-mark line (§17.3.1.29) that cannot
  // be split. If it doesn't fit in the space left on this page, relocate it
  // whole to the next page — matching the wholesale break the paginator applies
  // to unsplittable paragraphs — instead of letting the mark overflow the bottom
  // margin (which the prior unconditional break never allowed).
  const placeMarkOnly = (): { endY: number } => {
    let markH = estimateParagraphHeight(measureState, para, contentWPt, suppressSpaceBefore, marginLeftPt);
    let top = initialY;
    if (initialY > 0 && initialY + markH - para.spaceAfter > colBot()) {
      newPage(initialY);
      top = colTop();
      markH = estimateParagraphHeight(measureState, para, contentWPt, suppressSpaceBefore, marginLeftPt);
    }
    pages[pages.length - 1].push(stamp(para as PaginatedBodyElement));
    return { endY: top + markH };
  };
  const segs = buildSegments(para.runs, measureState);
  if (segs.length === 0) {
    return placeMarkOnly();
  }
  const wrapCtx: WrapLayoutCtx | undefined = measureState.floats.length > 0 ? {
    startPageY: measureState.y,
    paraX,
    floats: measureState.floats,
    lineBoxH: (a, d, _h, is) => lineBoxHeight(para.lineSpacing, a, d, 1, measureState.docGrid, paragraphHasRuby(para), is ?? 0, paragraphIsEastAsian(para)),
    pageH: measureState.pageH,
    markEmPx: paragraphMarkEmPx(para, 1),
  } : undefined;
  const lines = layoutLines(measureState.ctx, segs, paraW, para.indentFirst, 1, para.tabStops, wrapCtx, measureState.fontFamilyClasses, indLeft, measureState.kinsoku, gridCharDeltaPx(paraGrid(para, measureState), 1), measureState.defaultTabPt);
  if (lines.length === 0) {
    // Anchor-only paragraph: no inline lines, but the paragraph mark still
    // occupies one (possibly relocated) line (§17.3.1.29).
    return placeMarkOnly();
  }
  const paraHasRuby = paragraphHasRuby(para);

  // Compute-once eligibility, once per paragraph: a paragraph whose segment
  // TEXT depends on the paint state (page/numPages fields, note references)
  // must not ship its measure-time lines — the stamped text would be stale.
  // See paragraphSegsStateSensitive.
  const stampLines = !paragraphSegsStateSensitive(para);

  const perLineH = (l: typeof lines[number]) => lineBoxHeight(para.lineSpacing, l.ascent, l.descent, 1, measureState.docGrid, paraHasRuby, l.intendedSingle, paragraphIsEastAsian(para));
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
    // Available space in the current column from cursorY downward (balance target
    // for a non-last balanced column; the page bottom otherwise).
    const remaining = colBot() - cursorY;
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
        newPage(cursorY);
        cursorY = colTop();
        isFirstSliceOnPage = true;
        continue;
      }
      // First page, first line doesn't fit — force-emit it and let it overflow.
      lastFitting = firstFitting + 1;
      usedH += lineHeights[firstFitting];
    }
    // ECMA-376 §17.3.1.44 widowControl (default ON): keep at least two lines of
    // the paragraph together across a page break — never strand a single trailing
    // line on a later page (widow) nor a single leading line at a page bottom
    // (orphan). Skipped when w:widowControl is explicitly off
    // (para.widowControl === false). Only applies when lines actually carry over
    // (lastFitting < lines.length); the final slice can legally be one line.
    if (para.widowControl !== false && lastFitting < lines.length) {
      // Widow: this slice would leave exactly one line for a later page. Pull one
      // line down with it so ≥2 carry over — provided this slice keeps ≥1 line.
      if (lines.length - lastFitting === 1 && lastFitting - firstFitting >= 2) {
        lastFitting--;
        usedH -= lineHeights[lastFitting];
      }
      // Orphan: the paragraph's first line would sit alone at this page's bottom
      // (more lines follow). Relocate the paragraph start to the next page so it
      // begins with ≥2 lines — only when there is room above to break from
      // (cursorY > 0); a lone line at a fresh page top cannot be helped. Also
      // catches the case the widow pull above just reduced to a single line.
      if (firstFitting === 0 && lastFitting - firstFitting === 1 && cursorY > 0) {
        newPage(cursorY);
        cursorY = colTop();
        isFirstSliceOnPage = true;
        continue;
      }
    }
    const isFinalSlice = lastFitting === lines.length;
    if (isFinalSlice) usedH += spaceAfter;
    const sliceEl = {
      ...(para as object),
      type: 'paragraph',
      lineSlice: { start: firstFitting, end: lastFitting },
    } as PaginatedElementWithLines;
    // Phase 4-1 B2 Stage 1 — hand the paint pass the scale-1 lines this split
    // already computed so it can skip re-running layoutLines (compute-once).
    // The FULL array is stamped on every slice (not `lines.slice(...)`) because
    // the paint loop indexes by absolute line number; `lineSlice` still selects
    // the sub-range to paint. The same immutable array is shared across slices
    // and across repeated renderPage calls — the draw path only reads it. The
    // recorded inputs let the paint pass verify its own layout would be
    // identical before reusing (see renderParagraph's reuse gate). Skipped for
    // state-sensitive segments (stampLines above): those keep the recompute
    // path so field text resolves against the real page context.
    if (stampLines) {
      sliceEl.layoutLines = lines;
      sliceEl.layoutLinesInputs = {
        scale: 1,
        paraW,
        firstIndent: para.indentFirst,
        tabOriginPx: indLeft,
        gridDeltaPx: gridCharDeltaPx(paraGrid(para, measureState), 1),
        hasFloats: wrapCtx !== undefined,
        kinsoku: measureState.kinsoku,
      };
    }
    pages[pages.length - 1].push(stamp(sliceEl));
    lineIdx = lastFitting;
    cursorY += usedH;
    if (!isFinalSlice) {
      newPage(cursorY);
      cursorY = colTop();
      isFirstSliceOnPage = true;
    }
  }
  return { endY: cursorY };
}

/** Per-row heights used by both pagination and the height estimate. Mirrors the
 *  renderer's row sizing (exact / atLeast / auto + vMerge span distribution,
 *  ECMA-376 §17.4.80, §17.4.85) via the shared {@link resolveTableRowHeights}
 *  skeleton. Works in pt (scale 1); the cell measurer is the paginator's
 *  float-aware `estimateParagraphHeight` cursor-walk. Adjacent paragraphs inside
 *  a cell collapse spacing the same way `renderCellContent` does (ECMA-376
 *  §17.3.1.33 contextualSpacing + spaceAfter/spaceBefore overlap = max not sum),
 *  so the measured height matches the painted height. Without this, a cell
 *  containing a nested table followed by a paragraph with `spaceBefore` would
 *  measure taller than it paints, leaving a gap below the nested table. */
function computeTableRowHeights(state: RenderState, table: DocTable, contentWPt: number): number[] {
  const colWidths = resolveColumnWidths(table, contentWPt, state);
  return resolveTableRowHeights(table, colWidths, 1, (cell, cellW) => {
    const cm = effCellMargins(cell, table);
    const innerW = Math.max(1, cellW - cm.left - cm.right);
    // pt-space: estimateParagraphHeight emits the full spaceBefore (its
    // suppressSpaceBefore flag is for page-break continuations, not intra-cell
    // collapse), so sumCellContentHeight folds in contextualSuppressed
    // (§17.3.1.33) and the prevSpaceAfter/spaceBefore overlap to match the
    // paint pass's renderCellContent. Drop the §17.4.7 trailing structural
    // empty paragraph after a nested table for the SAME reason the paint-side
    // measurer (measureCellContentHeightPx) does — otherwise the paginator
    // would reserve more height than the paint pass uses and break the page
    // early (the two are contracted to agree, per this function's docstring).
    return cm.top + cm.bottom + sumCellContentHeight(trimTrailingStructuralMarker(cell.content), (ce) => {
      if (ce.type === 'paragraph') {
        return estimateParagraphHeight(state, ce as unknown as DocParagraph, innerW);
      }
      return estimateTableHeight(state, ce as unknown as DocTable, innerW);
    }, 1);
  });
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
 *   - autofit (the spec default) WITH a preferred table width (tblW=dxa/pct,
 *     §17.4.63): the tblGrid widths (§17.4.48) ARE the column widths, scaled to
 *     fit `contentWPt`, content min-width the only grower. Per-cell `tcW`
 *     (§17.4.71) is NOT re-applied: Word bakes the resolved auto-fit widths back
 *     into the saved `<w:gridCol>`, so for a round-tripped preferred-width table
 *     the grid already is the tcW-resolved layout (sample-3).
 *   - autofit with tblW=auto ("AutoFit to Contents"), or a degenerate all-zero
 *     grid: per-cell `tcW` + content min drive the column sizes. The saved grid
 *     is the style/default full text column, not a baked layout, so it must not
 *     pin the width — otherwise the table spans the page and overrides its own
 *     `w:jc` placement (sample-7's cover tables). See the in-body comment for the
 *     full rationale and the commit-8a3d8a5 history.
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

/** Resolve a table's per-grid-column widths (pt) to fit `contentWPt`. Exported
 *  for unit tests (column-widths.test) — see {@link calculateRowHeight} for the
 *  same test-export pattern. */
export function resolveColumnWidths(table: DocTable, contentWPt: number, state: RenderState): number[] {
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

  // Autofit (default). The width source depends on the table's preferred width
  // (ECMA-376 §17.4.63 `<w:tblW>` / §17.18.87 ST_TblWidth):
  //
  // (a) tblW=dxa or pct (a FIXED preferred width) ⇒ trust the `<w:tblGrid>`
  //     (§17.4.48 / gridCol §17.4.16), scaled to fit, with content min-widths
  //     the only grower. DELIBERATE DEVIATION from the literal autofit algorithm
  //     to match Word: §17.4.16/§17.4.52 make a cell's `<w:tcW>` (§17.4.71) an
  //     input that can override the grid's initial widths, but Word does not ship
  //     the pre-autofit state — it BAKES the resolved autofit widths back into
  //     the saved `<w:gridCol>`. So for a round-tripped, preferred-width Word
  //     table the grid already IS the tcW-resolved layout; re-applying `tcW`
  //     double-counts. sample-3's résumé grid is [2137, 222, 2430, 279, 2427,
  //     279] twips (a deliberately narrow first content column), yet each row's
  //     single-column cells carry `tcW≈30%`; the old "tcW overrides grid" path
  //     equalized the columns to ~116 pt apiece, shifting every later column
  //     right and re-wrapping the description paragraphs. Trusting the grid
  //     reproduces Word exactly.
  //
  // (b) tblW=auto ("AutoFit to Contents") ⇒ the table has NO preferred width, so
  //     the saved `<w:gridCol>` is the style/layout default (commonly the FULL
  //     text column), NOT a baked autofit result. Trusting it makes the table
  //     span the page and defeats its own `w:jc` placement (sample-7's cover
  //     tables carry gridCol=full-page yet a 100 pt `tcW`; the grid path centred
  //     them and broke the right/left alignment). So tblW=auto falls through to
  //     the tcW/content-preference autofit below, exactly as before commit
  //     8a3d8a5 sized every autofit table — that change correctly fixed the
  //     preferred-width case (a) but wrongly applied the grid to (b) too.
  //
  // A degenerate grid (no `<w:gridCol>` widths) also falls through to (b): with
  // no grid to anchor the columns, tcW + content are the only sizing signal.
  //
  // LIMITATION: a `tblW=pct` table whose SAVED grid no longer matches the
  // available width (authored under different margins, or by a tool that did not
  // bake the grid) is NOT re-scaled up to the pct target — the grid is taken
  // as-is (only scaled DOWN on overflow). Tracked for a future full tblW pass.
  const hasPreferredWidth = table.widthPt != null || table.widthPct != null;
  const gridSum = grid.reduce((s, w) => s + w, 0);
  if (gridSum > 0 && hasPreferredWidth) {
    const desired = grid.map((g, c) => Math.max(g, minW[c]));
    return fitToContent(desired);
  }

  // (b) tblW=auto (or a degenerate grid): preferred-width autofit, where `tcW`
  // and content drive the column sizes. `pref[c]` accumulates the strongest
  // single-column preference seen so far.
  const pref: number[] = new Array(n).fill(0);
  // `hasPref[c]` is true once any cell has expressed a preference for c.
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
  /** Advance to the next column / page. Receives the bottom (content-relative pt)
   *  the just-filled column reached so the caller can track the deepest column of
   *  a multi-column region (ECMA-376 §17.6.4). */
  newPage: (filledColBottom: number) => void,
  /** ECMA-376 §17.6.4 — current newspaper column index, read AFTER each
   *  `newPage()`. When provided, each table slice is tagged with its column.
   *  Omitted (single-column / direct unit tests) ⇒ no tag. */
  tagColIndex?: () => number,
  /** ECMA-376 §17.6.4 — the current SECTION's column geometry (constant across
   *  the split; a table is never split across a section boundary). Stamped on
   *  each slice so the renderer resolves its column against the right section. */
  colGeom?: ColumnGeom[],
  /** ECMA-376 §17.6.4 — content-relative pt of the current column-region TOP,
   *  read AFTER each `newPage()`. A continuation column of a continuous mid-page
   *  section restarts here, not at the page top. Omitted ⇒ the page top (0). The
   *  region-top page-absolute pt (`marginTop + columnTop()`) is stamped on each
   *  slice so the paint pass resets its cursor to the region top. */
  columnTop?: () => number,
  /** Page-content top (pt) used to convert `columnTop()` to a page-absolute Y for
   *  the `colTopPt` stamp. Omitted ⇒ no `colTopPt` stamp (single-column tests). */
  marginTopPt?: number,
  /** ECMA-376 §17.10.1 — the active SECTION's resolved header/footer set (constant
   *  across the split). Stamped so the renderer picks the right section's header/
   *  footer per page. Omitted ⇒ the renderer's body-level fallback. */
  tagSectionHF?: () => PaginatedBodyElement['sectionHF'],
  /** ECMA-376 §17.6.13 / §17.6.11 — the active SECTION's page geometry (size +
   *  margins; constant across the split). Stamped so the renderer sizes each page
   *  from the right section. Omitted ⇒ the renderer's body-level fallback. */
  tagSectionGeom?: () => SectionGeom,
): number {
  const colTop = columnTop ?? (() => 0);
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
    const sliceEl = { ...table, type: 'table', rows: sliceRows } as PaginatedBodyElement;
    if (tagColIndex) sliceEl.colIndex = tagColIndex();
    if (colGeom) sliceEl.colGeom = colGeom;
    // Front-loaded layout: stamp the region top (page-absolute pt) so the paint
    // pass resets this slice's column cursor to it instead of the page top.
    if (columnTop && marginTopPt != null) sliceEl.colTopPt = marginTopPt + colTop();
    if (tagSectionHF) sliceEl.sectionHF = tagSectionHF();
    if (tagSectionGeom) sliceEl.sectionGeom = tagSectionGeom();
    pages[pages.length - 1].push(sliceEl);

    y += used;
    start = end;
    firstSlice = false;
    if (start < n) {
      newPage(y);
      y = colTop();
    }
  }
  return y;
}

/**
 * ECMA-376 §17.10.1 precedence for picking a section's header/footer for one page:
 *   1. `first` — on the section's FIRST page when `<w:titlePg>` is set;
 *   2. `even`  — on even-parity pages when `<w:evenAndOddHeaders>` is set;
 *   3. `default` — otherwise.
 * `isFirstPageOfSection` (the section's first page, not necessarily page 0) and
 * `isEvenPage` (document page parity) are resolved per page by the caller so this
 * works for per-section title pages (e.g. sample-13's masthead section gets its
 * own first-page footer on whatever document page it begins).
 */
function pickHeaderFooter(
  set: HeadersFooters,
  isFirstPageOfSection: boolean,
  isEvenPage: boolean,
  titlePage: boolean,
  evenAndOdd: boolean,
): HeaderFooter | null {
  if (titlePage && isFirstPageOfSection && set.first) return set.first;
  if (evenAndOdd && isEvenPage && set.even) return set.even;
  return set.default ?? null;
}

/**
 * ECMA-376 §17.10.1 — resolve the section active at the TOP of `pageIndex` (the
 * section that owns that page's content) and whether `pageIndex` is that section's
 * FIRST page. The paginator stamps each element's `sectionHF` (the upcoming
 * `SectionBreak`'s resolved set, or — when absent — the body-level section). So the
 * active section for a page is the `sectionHF` of its first element; `undefined`
 * means the final/body section (`doc.headers`/`doc.footers`/`section.titlePage`).
 *
 * `isFirstPageOfSection` is true when this is page 0 or the active section differs
 * from the previous page's — i.e. a section boundary fell on this page's top. Two
 * distinct sections are compared by reference identity of their stamped `sectionHF`
 * object (each section is stamped with one shared object instance).
 */
function resolvePageSection(
  pages: PaginatedBodyElement[][],
  pageIndex: number,
  doc: DocxDocumentModel,
): { headers: HeadersFooters; footers: HeadersFooters; titlePage: boolean; isFirstPageOfSection: boolean; geom: SectionGeom } {
  const sectionOf = (idx: number): PaginatedBodyElement['sectionHF'] => pages[idx]?.[0]?.sectionHF;
  const active = sectionOf(pageIndex);
  const headers = active?.headers ?? doc.headers;
  const footers = active?.footers ?? doc.footers;
  const titlePage = active?.titlePage ?? doc.section.titlePage;
  const isFirstPageOfSection = pageIndex === 0 || sectionOf(pageIndex - 1) !== active;
  // ECMA-376 §17.6.13 / §17.6.11 — the page's geometry, from the paginator's stamp
  // on this page's first element, falling back to the body-level section (the value
  // the paginator itself stamps for the final/single section, so this fallback only
  // fires for a truly empty page). Sizes the canvas + margins per page.
  const geom: SectionGeom = pages[pageIndex]?.[0]?.sectionGeom ?? sectionGeomOf(doc.section);
  return { headers, footers, titlePage, isFirstPageOfSection, geom };
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

/**
 * Whether a paragraph places NO inline content — no text, image, shape, math, or
 * break run. It still produces one paragraph-mark line box (§17.3.1.29), but
 * carries no glyphs.
 */
function isInklessParagraph(p: DocParagraph): boolean {
  return !(p.runs ?? []).some((r) => {
    const run = r as { type?: string; text?: string };
    if (run.type === 'text') return (run.text ?? '').length > 0;
    return true; // image / shape / math / break runs are visible content
  });
}

/**
 * Whether `body[i]` is an empty paragraph that carries a SECTION BREAK — an
 * inkless paragraph immediately followed by a `sectionBreak` element. (The parser
 * emits a sectPr-bearing paragraph as a `Paragraph` followed by a `SectionBreak`,
 * so the spacer paragraph and its break are adjacent siblings.)
 *
 * This adjacency is the structural test; the CALLER additionally gates on the
 * break being CONTINUOUS (isContinuousSectionSpacer) — that is the measured
 * trigger. When it holds, the spacer's spacing-BEFORE is suppressed: it sits
 * flush below the preceding paragraph (a section transition, not normal flow).
 *
 * NOTE — this matches Microsoft WORD's observed layout, NOT a spec rule. ECMA-376
 * §17.3.1.33 would apply the `before` normally (a consumer takes
 * `max(prev.after, this.before)`), and no clause — nor any [MS-DOC] /
 * [MS-OI29500] note — suppresses it for a section-break or empty paragraph
 * (verified independently). The model is reconstructed from Word's OWN output:
 *   - sample-13 has two empty `mSectionBreak` paragraphs (each `w:before="440"`,
 *     22pt) before a continuous 2-column break. Measuring the Word PDF
 *     (pdftotext -bbox), Word keeps BOTH empty line boxes but renders them flush
 *     — i.e. it drops exactly ONE 22pt `before` — so "1. INTRODUCTION" sits at
 *     64.4pt below "Keywords" instead of our pre-fix 87.3pt.
 *   - Ablation against Word's output pinned the trigger: removing the section
 *     break removes the suppression; putting text in the spacer removes it;
 *     changing the spacer's style does NOT (same-style is irrelevant); 1-col vs
 *     2-col does NOT (column count is irrelevant). So: an EMPTY paragraph at a
 *     CONTINUOUS section boundary, with only its own `before` dropped.
 * So we drop ONLY the spacer's own `before` and keep both line boxes — Word's
 * decomposition. (LibreOffice independently suppresses here too — corroborating
 * that this is real, undocumented Word-compat behaviour, tdf#166503 — but via a
 * DIFFERENT mechanism that collapses a whole paragraph (~33.5pt); we deliberately
 * do NOT follow its amount, only Word's measured −22pt.) Revisit and cite if a
 * primary source ever surfaces.
 */
function isSectionBreakSpacerAt(body: ArrayLike<{ type?: string }>, i: number): boolean {
  const el = body[i] as { type?: string } | undefined;
  if (!el || el.type !== 'paragraph') return false;
  const next = body[i + 1] as { type?: string } | undefined;
  if (next?.type !== 'sectionBreak') return false;
  return isInklessParagraph(el as unknown as DocParagraph);
}

/**
 * Sum the heights of a cell's content elements with paragraph spacing collapsed
 * the same way `renderCellContent` paints them, so a cell measured for row
 * sizing equals the height it actually paints. Two collapse rules apply
 * (mirroring the paint pass):
 *
 *   - ECMA-376 §17.3.1.33 `<w:contextualSpacing>`: a paragraph whose
 *     contextualSpacing toggle matches the previous paragraph's AND shares its
 *     styleId emits zero spaceBefore (the toggle suppresses spacing between
 *     same-style siblings).
 *   - Adjacent-paragraph spacing OVERLAP: the gap between two paragraphs is
 *     `max(prevSpaceAfter, currSpaceBefore)`, not their sum. We subtract the
 *     overlap `min(prevSpaceAfter, effBefore)` so a 12pt space-after followed
 *     by a 12pt space-before contributes 12pt of gap, not 24pt.
 *
 * A nested table (CellElement other than paragraph) resets the
 * prev-paragraph context — the next paragraph after a table spaces from a
 * fresh baseline, exactly as `renderCellContent` does.
 *
 * `perElementHeight(elem)` returns each element's full measured height: for a
 * paragraph it must include its full `spaceBefore` (no contextual or overlap
 * adjustment) so this helper can subtract the collapse correctly; for a nested
 * table it is the table's own total height. `spaceScale` converts spec spacing
 * (pt) into the same units as `perElementHeight` returns (1 for pt; the device
 * scale for px); the subtracted collapse therefore lands in matching units.
 *
 * Exported for unit tests (table-spacing-collapse.test).
 */
export function sumCellContentHeight(
  content: CellElement[],
  perElementHeight: (el: CellElement) => number,
  spaceScale: number,
): number {
  let h = 0;
  let prevPara: DocParagraph | null = null;
  let prevSpaceAfter = 0;
  for (const ce of content) {
    if (ce.type === 'paragraph') {
      const para = ce as unknown as DocParagraph;
      const suppress = contextualSuppressed(prevPara, para);
      const effBefore = suppress ? 0 : para.spaceBefore;
      // §17.3.1.9 contextualSpacing: between two same-style paragraphs that both
      // set it, BOTH the previous after and this before are dropped (gap = 0), not
      // just collapsed — so e.g. a code listing's lines sit tight.
      const overlap = suppress ? prevSpaceAfter : Math.min(prevSpaceAfter, effBefore);
      h += perElementHeight(ce)
        - (suppress ? para.spaceBefore : 0) * spaceScale
        - overlap * spaceScale;
      prevPara = para;
      prevSpaceAfter = para.spaceAfter;
    } else {
      h += perElementHeight(ce);
      prevPara = null;
      prevSpaceAfter = 0;
    }
  }
  return h;
}

/** ECMA-376 §17.6.4 `<w:cols w:sep="1">` — draw a thin vertical rule centred in
 *  each inter-column gap, spanning the section's content height. The spec does
 *  not prescribe a width/colour; Word draws a hairline, so we use ~0.5pt in the
 *  default text colour (matching the footnote separator convention). */
function drawColumnSeparators(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  columns: ColumnGeom[],
  sec: SectionProps,
  scale: number,
): void {
  // §17.6.11: the separators span the body content band, whose top/bottom are inset from
  // the page edges by the margins' MAGNITUDE (bodyMarginInsetPt). Identity for non-negative.
  const topY = bodyMarginInsetPt(sec.marginTop) * scale;
  const botY = (sec.pageHeight - bodyMarginInsetPt(sec.marginBottom)) * scale;
  ctx.save();
  ctx.strokeStyle = '#000000';
  ctx.lineWidth = Math.max(1, Math.round(0.5 * scale));
  for (let i = 0; i < columns.length - 1; i++) {
    const gapStart = columns[i].xPt + columns[i].wPt;
    const gapEnd = columns[i + 1].xPt;
    const midX = Math.round(((gapStart + gapEnd) / 2) * scale) + 0.5;
    ctx.beginPath();
    ctx.moveTo(midX, topY);
    ctx.lineTo(midX, botY);
    ctx.stroke();
  }
  ctx.restore();
}

function renderBodyElements(
  elements: PaginatedBodyElement[],
  state: RenderState,
  /** ECMA-376 §17.6.4 — page-absolute column geometry (pt) for the page. Used as
   *  the fallback when an element carries no per-section `colGeom` (header /
   *  footer, single-column). Omitted ⇒ the state's existing full-width
   *  contentX/contentW is used unchanged. */
  columns?: ColumnGeom[],
  /** ECMA-376 §17.6.11 — device-px the content top is reserved below the top
   *  margin for a header taller than its top-margin allowance. `state.y` already
   *  carries it for column 0 (the caller's `bodyState.y`); it must be re-added
   *  whenever a newspaper column (§17.6.4) restarts at its section's region top,
   *  which is otherwise a bare-margin `colTopPt`. 0 (default) for header/footer
   *  and nested renders, which have no such reserve. */
  headerReservePx = 0,
): void {
  let prevPara: DocParagraph | null = null;
  let prevSpaceAfter = 0;
  // ECMA-376 §20.4.3.2/§20.4.3.5: a wp:anchor whose positionV relativeFrom is
  // page-level (page / margin / *Margin / column — anything but
  // paragraph/line/character) is positioned independently of its source-order
  // anchoring paragraph. Word lays such floats out as soon as the page is
  // opened, so paragraphs PRECEDING the anchor's paragraph on the same page
  // still wrap around it. Mirror that by pre-registering them at every
  // page-start — `state` is fresh per page (renderDocumentToCanvas builds a
  // new bodyState with floats=[] for each render call), so this single call
  // covers the page. `elements` is already the paginated body for THIS page,
  // so scanning from index 0 stays within the page (preRegisterPageFloats
  // additionally stops at the first explicit page/non-continuous section
  // break, which is a no-op here but matches the paginator's call sites).
  state.pageAnchorPrescanned = new Set();
  preRegisterPageFloats(elements, 0, state);
  // Collect front floats (behindDoc="0") and paint them after the whole flow, so
  // they layer ON TOP of later inline content (§20.4.2.10). Saved/restored so a
  // nested render (table cell / header-footer) keeps its own top layer.
  const prevDeferFront = state.deferFront;
  state.deferFront = [];
  // The (geometry, column index) the flow `state` is currently set to. `activeCol`
  // starts at -1 so the first element always seeds the column. `activeGeom` tracks
  // the SECTION whose columns are in effect (per-section newspaper columns,
  // §17.6.4): elements carry their own section's geometry in `colGeom`, so two
  // sections sharing a page (a continuous break) each resolve their `colIndex`
  // against the right widths.
  let activeCol = -1;
  let activeGeom: ColumnGeom[] | undefined;
  // ECMA-376 §17.3.1.7 — the next IN-FLOW paragraph that is adjacent to `elements[i]`
  // in the SAME newspaper column (same colGeom + colIndex), with no intervening
  // non-paragraph element (a table/floating-table boundary closes the border box).
  // A `framePr` paragraph is out of flow (§17.3.1.11) and is skipped/blocks the run.
  // Returns null when the run is broken (column change, table between, end of page).
  // A page break never appears mid-`elements` (the paginator splits pages), so the
  // box correctly closes at the column/page bottom without a special case here.
  const sameColumn = (a: PaginatedBodyElement, b: PaginatedBodyElement): boolean =>
    (a.colGeom ?? columns) === (b.colGeom ?? columns) && (a.colIndex ?? 0) === (b.colIndex ?? 0);
  const flowParaInColumn = (cur: PaginatedBodyElement, sibling: PaginatedBodyElement | undefined): DocParagraph | null => {
    if (!sibling) return null;
    if (sibling.type !== 'paragraph') return null; // table/other ends the run
    if (!sameColumn(cur, sibling)) return null; // column change ends the run
    const p = sibling as unknown as DocParagraph;
    if (p.framePr) return null; // out-of-flow frame paragraph is not in the run
    return p;
  };
  const prevFlowParaInColumn = (i: number): DocParagraph | null =>
    flowParaInColumn(elements[i], elements[i - 1]);
  const nextFlowParaInColumn = (i: number): DocParagraph | null =>
    flowParaInColumn(elements[i], elements[i + 1]);
  for (let elIdx = 0; elIdx < elements.length; elIdx++) {
    const el = elements[elIdx];
    // Per-section column geometry: prefer the element's own (stamped by the
    // paginator), else the page-level columns. A single full-width column (or no
    // geometry) is the unchanged single-column path.
    const cols = el.colGeom ?? columns;
    const multiCol = !!cols && cols.length > 1;
    const elCol = el.colIndex ?? 0;
    // Switch the flow when this element's section geometry OR its column index
    // changed (also seeds the first element). Reset the vertical cursor to the
    // column TOP only when ADVANCING to a higher column within the page (newspaper
    // fill); a same-or-lower colIndex from a continuous section break continues at
    // the current y so the new section stacks below the previous one rather than
    // overprinting it. Floats are NOT cleared here — they are page-scoped and the
    // per-page fresh bodyState already gave a clean set.
    if (multiCol && (cols !== activeGeom || elCol !== activeCol)) {
      const col = cols[Math.min(elCol, cols.length - 1)];
      state.contentX = col.xPt * state.scale;
      state.contentW = col.wPt * state.scale;
      if (elCol > activeCol && cols === activeGeom) {
        // Region top (front-loaded layout): a continuation column restarts at the
        // SECTION's region top — for a continuous mid-page section that is BELOW
        // the preceding single-column content, not the page content top. The
        // paginator computed and stamped it (`colTopPt`, page-absolute pt); the
        // paint pass consumes it rather than re-deriving the column top. Absent ⇒
        // a page-spanning section ⇒ the page content top, unchanged. A tall header
        // (§17.6.11) shifts the whole content frame down — `colTopPt`/`state.marginTop`
        // carry the body inset (|margin|, not the signed margin), so re-add the reserve
        // to match column 0's `bodyState.y`. A negative top margin reserves 0, so the
        // re-add is a no-op there (the body already overlaps the header).
        state.y = (el.colTopPt ?? state.marginTop) * state.scale + headerReservePx;
      }
      prevPara = null;
      prevSpaceAfter = 0;
      activeCol = elCol;
      activeGeom = cols;
    } else if (!multiCol && cols && cols.length === 1 && cols !== activeGeom) {
      // A single-column SECTION on a multi-section page (e.g. sample-5 sections
      // 1–4): set the full-width content band for this section. Continue at the
      // current y (continuous), BUT when the PRECEDING section was multi-column,
      // never above its deepest column: `colTopPt` carries that section's region
      // bottom (ECMA-376 §17.6.4), so a full-width section after a 2-column run
      // clears BOTH columns instead of overprinting the taller one. Gated on the
      // previous section being multi-column so a 1-col→1-col continuous break
      // (sample-5) keeps the exact prior behavior (continue at the current y).
      if (activeGeom && activeGeom.length > 1 && el.colTopPt != null) {
        state.y = Math.max(state.y, el.colTopPt * state.scale + headerReservePx);
      }
      const col = cols[0];
      state.contentX = col.xPt * state.scale;
      state.contentW = col.wPt * state.scale;
      prevPara = null;
      prevSpaceAfter = 0;
      activeCol = elCol;
      activeGeom = cols;
    }
    if (el.type === 'paragraph') {
      const para = el as unknown as DocParagraph;
      const slice = (el as PaginatedBodyElement).lineSlice;
      // A frame paragraph (ECMA-376 §17.3.1.11) is positioned out of flow and
      // does not participate in spacing collapse or pagination splitting. Draw
      // it absolutely and leave prevPara/spaceAfter untouched so the following
      // non-frame paragraph spaces against the paragraph BEFORE the frame.
      if (para.framePr) {
        renderFrameParagraph(para, state, frameAnchorLineHeightPx(elements, el, state));
        continue;
      }
      // A zero-before continuous-section spacer renders no mark line box (Word's
      // section-mark collapse; stamped by the paginator — see
      // isCollapsedContinuousSpacer). Skip it: paint nothing, advance y by nothing,
      // and leave prevPara/prevSpaceAfter as the paragraph BEFORE it so the new
      // section's first paragraph spaces against that paragraph. Mirrors the
      // paginator's identical skip so fill and paint stay in lockstep.
      if ((el as PaginatedBodyElement).collapsedSpacer) continue;
      const contextual = contextualSuppressed(prevPara, para);
      // Empty section-break spacer: drop only its own before (see
      // isSectionBreakSpacerAt); contextualSpacing drops the previous after too.
      // The adjacency was detected during pagination and stamped (the sectionBreak
      // marker is gone from this per-page list), so read the tag here.
      const spacer = !!(el as PaginatedBodyElement).sectionBreakSpacer;
      const suppress = contextual || spacer;
      // An empty paragraph immediately before a COLLAPSED continuous spacer begins the
      // section-break empty run; Word renders it FLUSH below the preceding content,
      // dropping the previous paragraph's spaceAfter too. Read the paginator-stamped
      // flag rather than looking ahead in this per-page slice: the collapsed spacer can
      // sit on the next page, where `elements[elIdx + 1]` would be undefined.
      const leadsCollapsedRun = !!(el as PaginatedBodyElement).leadsCollapsedRun;
      // Collapse spaceAfter+spaceBefore like Word: use max, not sum.
      const effBefore = suppress ? 0 : para.spaceBefore;
      // §17.3.1.9 contextualSpacing: between two same-style paragraphs that both
      // set it, BOTH the previous after and this before are dropped (gap = 0), not
      // just collapsed — so e.g. a code listing's lines sit tight.
      const overlap = (contextual || leadsCollapsedRun) ? prevSpaceAfter : Math.min(prevSpaceAfter, effBefore);
      state.y -= overlap * state.scale;
      // Continuation slices (slice.start > 0) suppress spaceBefore: the
      // earlier slice already consumed it on the previous page. Likewise
      // mid-paragraph slices (slice.end < total) suppress spaceAfter — only
      // the slice covering the FINAL line of the paragraph emits it.
      const isContinuation = !!slice && slice.start > 0;
      // ECMA-376 §17.3.1.7 paragraph-border merge: suppress this paragraph's TOP
      // edge when the previous IN-FLOW paragraph (already null on column/section
      // change, table/frame boundary — exactly the run-breaking cases) shares its
      // border box, and its BOTTOM edge when the next adjacent same-column
      // paragraph does. The run is thus drawn as one box: top on the first member,
      // bottom on the last, `between` (if any) at the inner joins.
      const borderMerge: ParaBorderMerge | undefined =
        hasAnyBorderEdge(para.borders)
          ? {
              suppressTop: parasShareBorderBox(prevFlowParaInColumn(elIdx), para),
              suppressBottom: parasShareBorderBox(para, nextFlowParaInColumn(elIdx)),
            }
          : undefined;
      renderParagraph(para, state, suppress || isContinuation, slice, false, borderMerge);
      prevPara = para;
      prevSpaceAfter = para.spaceAfter;
    } else if (el.type === 'table') {
      const tbl = el as unknown as DocTable;
      // A floating table (ECMA-376 §17.4.57 `<w:tblpPr>`) is out of flow: it is
      // drawn absolutely by renderFloatTable and adds no flow height, so leave
      // prevPara/spaceAfter untouched (the following content spaces against the
      // paragraph BEFORE the table, exactly like a frame paragraph). A block
      // table resets them (it ends the previous spacing context).
      renderTable(tbl, state);
      if (!tbl.tblpPr) {
        prevPara = null;
        prevSpaceAfter = 0;
      }
    }
  }
  // Paint the deferred front floats on top of the page's inline flow, then
  // restore the enclosing render's collector.
  const deferredFront = state.deferFront ?? [];
  state.deferFront = prevDeferFront;
  for (const draw of deferredFront) draw();
}

function renderParaList(paras: DocParagraph[], state: RenderState): void {
  let prevPara: DocParagraph | null = null;
  let prevSpaceAfter = 0;
  for (let i = 0; i < paras.length; i++) {
    const para = paras[i];
    const suppress = contextualSuppressed(prevPara, para);
    const effBefore = suppress ? 0 : para.spaceBefore;
    const overlap = Math.min(prevSpaceAfter, effBefore);
    state.y -= overlap * state.scale;
    // ECMA-376 §17.3.1.7 paragraph-border merge. This list is a single flow (a
    // note or a table cell), so adjacency is just consecutive list members; a
    // frame paragraph (§17.3.1.11) is out of flow and breaks the run. `framePr`
    // siblings are filtered by parasShareBorderBox, so compare with the literal
    // neighbors here.
    const prevSibling = (paras[i - 1] ?? null) as DocParagraph | null;
    const nextSibling = (paras[i + 1] ?? null) as DocParagraph | null;
    const borderMerge: ParaBorderMerge | undefined = hasAnyBorderEdge(para.borders)
      ? {
          suppressTop: parasShareBorderBox(prevSibling, para),
          suppressBottom: parasShareBorderBox(para, nextSibling),
        }
      : undefined;
    renderParagraph(para, state, suppress, undefined, false, borderMerge);
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

/**
 * Render a paragraph that produces NO inline lines — either literally empty
 * (no segments) or anchor-only (its only content is wrap floats, drawn
 * separately). Per ECMA-376 §17.3.1.29 such a paragraph still emits ONE
 * paragraph-mark line box; this advances `state.y` past it, draws its shading /
 * borders, and lays its wrapNone anchor images at the (possibly float-flowed)
 * paragraph base.
 *
 * Shared by renderParagraph's `segments.length === 0` and `lines.length === 0`
 * branches (previously duplicated verbatim). The anchor-only branch's
 * slice-boundary guards (spaceAfter only on the final slice, anchor images only
 * on the first) are parameterized via `markCtx.totalLines` / `lineSlice`; for
 * the literally-empty branch `lineSlice` is always undefined (empty paragraphs
 * are never sliced), so those guards reduce to the unconditional behavior it had.
 */
function renderEmptyMarkParagraph(
  para: DocParagraph,
  state: RenderState,
  markCtx: {
    grid: DocGridCtx;
    paraHasRuby: boolean;
    contentX: number;
    indLeft: number;
    paraW: number;
    textAreaTopY: number;
    paragraphStartY: number;
    /** Flowed top of the mark line (output of resolveEmptyMarkTop). */
    markTop: number;
    /** Total laid-out line count (0 here); used by the slice guards. */
    totalLines: number;
    lineSlice?: { start: number; end: number };
    /** §17.3.1.7 paragraph-border merge (suppress top/bottom edges). */
    borderMerge?: ParaBorderMerge;
  },
): void {
  const { ctx, scale, dryRun } = state;
  const { grid, paraHasRuby, contentX, indLeft, paraW, textAreaTopY,
    paragraphStartY, markTop, totalLines, lineSlice, borderMerge } = markCtx;
  // Displacement applied by the float-flow (0 when the mark fits where it is).
  const flowShift = Math.max(0, markTop - textAreaTopY);
  if (markTop > state.y) state.y = markTop;
  const markRectTop = state.y;
  const emptyH = paragraphMarkLineHeight(para, scale, grid, paraHasRuby, state.docEastAsian, ctx, state.fontFamilyClasses);
  if (para.shading && !dryRun) {
    ctx.fillStyle = `#${para.shading}`;
    const sb = paraShadingRect(contentX + indLeft, markRectTop, paraW, emptyH, para.borders, borderMerge, scale);
    ctx.fillRect(sb.x, sb.y, sb.w, sb.h);
  }
  state.y += emptyH;
  if (para.borders && !dryRun) {
    drawParaBorders(ctx, contentX + indLeft, markRectTop, paraW, emptyH, para.borders, scale, state.dpr, borderMerge);
  }
  // Only the slice covering the FINAL line emits spaceAfter. With no inline
  // lines there is a single slice, so this is the whole paragraph.
  const isFinalSlice = !lineSlice || lineSlice.end >= totalLines;
  if (isFinalSlice) state.y += para.spaceAfter * scale;
  // wrapNone anchor images anchor relative to the paragraph (ayFromPara); when
  // the mark line flowed below a float band the paragraph (and its wrapNone
  // image) drops by the same amount, so shift the anchor base by flowShift while
  // keeping the un-flowed base (paragraphStartY) otherwise unchanged. Only the
  // first slice draws them (a continuation slice already did on its page).
  if (!lineSlice || lineSlice.start === 0) {
    renderAnchorImages(para, state, paragraphStartY + flowShift);
  }
}

// ===== Text frames & drop caps (ECMA-376 §17.3.1.11) =====

/**
 * One line height (px) of the anchor (following non-frame) paragraph, used to
 * size a drop cap by `lines` (§17.3.1.11). The drop cap height equals
 * `lines` × this. Scans `elements` after the frame element for the first
 * non-frame paragraph; falls back to the frame paragraph's own single-line
 * height when none follows (a degenerate trailing frame).
 */
function frameAnchorLineHeightPx(
  elements: PaginatedBodyElement[],
  frameEl: PaginatedBodyElement,
  state: RenderState,
): number {
  const start = elements.indexOf(frameEl);
  for (let j = start + 1; j < elements.length; j++) {
    const e = elements[j];
    if (e.type !== 'paragraph') continue;
    const p = e as unknown as DocParagraph;
    if (p.framePr) continue; // adjacent frame paragraphs are part of the frame
    return paragraphMarkLineHeight(
      p,
      state.scale,
      paraGrid(p, state),
      paragraphHasRuby(p),
      state.docEastAsian,
      state.ctx,
      state.fontFamilyClasses,
    );
  }
  const fp = frameEl as unknown as DocParagraph;
  return paragraphMarkLineHeight(
    fp,
    state.scale,
    paraGrid(fp, state),
    paragraphHasRuby(fp),
    state.docEastAsian,
    state.ctx,
    state.fontFamilyClasses,
  );
}

/**
 * Render a paragraph that is part of a text frame (`para.framePr` set), per
 * ECMA-376 §17.3.1.11.
 *
 * The frame is OUT OF FLOW: it is drawn at an absolute (anchor-relative)
 * position and does NOT advance the in-flow `state.y`, so the following
 * non-frame paragraph begins where this paragraph sat. The frame's glyphs are
 * painted at their own run sizes (a drop cap's big letter is just a large `sz`
 * run, e.g. sample-11's 58.5 pt "D"). A wrap exclusion FloatRect is registered
 * (unless wrap="none") so the following body text flows around the frame — the
 * exclusion x-range is built from the frame's COLUMN-relative band so
 * resolveLineFloatWindow only constrains the matching column (#513).
 *
 * `anchorLineHpx` is one line height of the following non-frame paragraph,
 * needed to size a drop cap by `lines`. Falls back to the frame paragraph's own
 * single-line height when there is no following paragraph.
 */
/**
 * Resolve a frame paragraph's box (canvas px) from the current flow geometry.
 * Lays the frame content out at a wide width to get its natural size, then maps
 * the framePr anchors/alignment. Shared by the renderer (draw) and the
 * paginator (wrap estimate) so both see the same band.
 */
function resolveFrameBox(
  para: DocParagraph,
  state: RenderState,
  anchorLineHpx: number,
): FrameBox {
  const fp = para.framePr!;
  const { scale } = state;
  const paraTop = state.y;
  const grid = paraGrid(para, state);
  const paraHasRuby = paragraphHasRuby(para);
  const segments = buildSegments(para.runs, state);

  // Measure the frame's natural content size at a wide width (single-line frame
  // content stays one line). No floats apply INSIDE the frame; the cap glyph's
  // own run size drives its extent.
  const measureW = 100000;
  const lines =
    segments.length === 0
      ? []
      : layoutLines(
          state.ctx,
          segments,
          measureW,
          0,
          scale,
          para.tabStops,
          undefined,
          state.fontFamilyClasses,
          0,
          state.kinsoku,
          gridCharDeltaPx(grid, scale),
          state.defaultTabPt,
        );
  const contentW =
    lines.length === 0
      ? 0
      : Math.max(...lines.map((l) => l.segments.reduce((s, sg) => s + sg.measuredWidth, 0)));
  const contentH = lines.reduce(
    (s, l) =>
      s +
      lineBoxHeight(para.lineSpacing, l.ascent, l.descent, scale, grid, paraHasRuby, l.intendedSingle, paragraphIsEastAsian(para)),
    0,
  );

  return computeFrameBox(fp, state, paraTop, contentW, contentH, anchorLineHpx);
}

// ===== Floating tables (ECMA-376 §17.4.57 w:tblpPr / §17.4.56 w:tblOverlap) =====
// Placement math (computeFloatTableBox / registerTableFloat / floatTableWrapSide)
// lives in float-table-geometry.ts; the float-table render path below consumes it.

/**
 * Render a paragraph that is part of a text frame (`para.framePr` set), per
 * ECMA-376 §17.3.1.11.
 *
 * The frame is OUT OF FLOW: it is drawn at an absolute (anchor-relative)
 * position and does NOT advance the in-flow `state.y`, so the following
 * non-frame paragraph begins where this paragraph sat. The frame's glyphs are
 * painted at their own run sizes (a drop cap's big letter is just a large `sz`
 * run, e.g. sample-11's 58.5 pt "D"). A wrap exclusion is then registered so
 * following body text flows around the frame.
 *
 * `anchorLineHpx` is one line height of the following non-frame paragraph,
 * needed to size a drop cap by `lines` (§17.3.1.11).
 */
function renderFrameParagraph(
  para: DocParagraph,
  state: RenderState,
  anchorLineHpx: number,
): void {
  const fp = para.framePr!;
  // In-flow Y the following paragraph must resume from. The frame is out of
  // flow; we restore this after drawing so state.y is untouched by the frame.
  const inFlowY = state.y;
  const box = resolveFrameBox(para, state, anchorLineHpx);

  // Draw the frame's glyphs by redirecting the flow geometry to the frame box,
  // then rendering the paragraph through the normal line path. inFrame=true
  // suppresses anchor-float re-registration and avoids re-entering the frame
  // dispatch; suppressSpaceBefore=true keeps the cap anchored to the paragraph
  // top (the frame is positioned absolutely, not in flow).
  const savedX = state.contentX;
  const savedW = state.contentW;
  state.contentX = box.x;
  state.contentW = Math.max(box.w, box.exRight - box.x);
  state.y = box.y;
  renderParagraph(para, state, true, undefined, /* inFrame */ true);
  state.contentX = savedX;
  state.contentW = savedW;

  // Restore the in-flow cursor: the frame consumes NO vertical space in the
  // body flow (§17.3.1.11 — the frame is positioned relative to the next
  // non-frame paragraph; the frame itself does not advance the flow).
  state.y = inFlowY;

  registerFrameFloat(box, fp, state);
}

/** ECMA-376 §17.15.1.25 — the ABSENT default for `<w:defaultTabStop>`: "If this
 *  element is omitted, then automatic tab stops should be generated at 720
 *  twentieths of a point (0.5")", i.e. 36 pt. Used ONLY as the fallback when a
 *  document carries no `<w:defaultTabStop>`; a document that sets one overrides
 *  this via {@link resolveDefaultTabPt}. Shared by the line layout
 *  (`layoutLines`) and the numbered-list marker's trailing-tab advance
 *  (`renderParagraph`). */
const DEFAULT_TAB_PT = 36;

/** ECMA-376 §17.15.1.25 — resolve the document's automatic tab-stop interval
 *  (pt): the explicit `<w:defaultTabStop>` value when present, else the spec
 *  absent default of 720 twips (36pt). Mirrors {@link resolveKinsokuRules}: the
 *  resolved value is threaded into both the measure and draw passes so they
 *  agree. */
function resolveDefaultTabPt(settings: DocSettings | undefined): number {
  const v = settings?.defaultTabStop;
  // §17.15.1.25 defines automatic stops as multiples of the interval, which is
  // undefined for a non-positive interval; fall back to the documented absent
  // default (720 twips = 36pt) so the automatic grid always advances.
  return v != null && v > 0 ? v : DEFAULT_TAB_PT;
}

/** ECMA-376 §17.3.1.37 (tabs) + §17.15.1.25 (defaultTabStop) — advance from a
 *  pen position to the next effective tab stop, ALL in text-margin px.
 *
 *  The effective stop set is the custom stops PLUS automatic left-aligned stops
 *  at every multiple of `intervalPx` that occurs AFTER all custom stops
 *  (§17.15.1.25: "Automatic tab stops refer to the tab stop locations which
 *  occur after all custom tab stops"; the spec example puts a 0.25" grid at
 *  2.5"/2.75"/3.0" past a 2.28" custom stop). A tab moves to the smallest
 *  effective stop strictly greater than `curMarginPx`.
 *
 *  @param curMarginPx pen position, measured from the text margin.
 *  @param customStopsPx custom stops in margin-px (`pos = tabStop.pos * scale`),
 *    order not assumed; each carries its own alignment + leader.
 *  @param intervalPx automatic-stop interval = `defaultTabPt * scale`.
 *  @returns the chosen stop (custom keeps its alignment/leader; an automatic
 *    stop is `'left'` with no leader), or `null` only when no stop exists. */
export function nextTabStop(
  curMarginPx: number,
  customStopsPx: { pos: number; alignment: TabStop['alignment']; leader?: TabStop['leader'] }[],
  intervalPx: number,
): { pos: number; alignment: TabStop['alignment']; leader?: TabStop['leader'] } | null {
  // Candidate 1 — the nearest custom stop strictly past the pen. The spec set is
  // unordered, so scan for the minimum rather than assuming sort order.
  let custom: { pos: number; alignment: TabStop['alignment']; leader?: TabStop['leader'] } | null = null;
  let maxCustomPx = 0;
  for (const t of customStopsPx) {
    if (t.pos > maxCustomPx) maxCustomPx = t.pos;
    if (t.pos > curMarginPx && (custom === null || t.pos < custom.pos)) custom = t;
  }

  // Candidate 2 — the nearest automatic multiple of the interval that is BOTH
  // past the pen AND past the last custom stop (§17.15.1.25: automatic stops
  // begin only after all custom stops). EPS forces a pen sitting exactly on a
  // boundary to advance to the NEXT one.
  let auto: { pos: number; alignment: TabStop['alignment'] } | null = null;
  if (intervalPx > 0) {
    const EPS = 1e-6;
    const from = Math.max(curMarginPx, maxCustomPx);
    let pos = Math.ceil((from + EPS) / intervalPx) * intervalPx;
    // Guard against floating-point landing on/under the pen (rounding can yield a
    // boundary that is not strictly greater): step to the next multiple.
    if (pos <= curMarginPx) pos += intervalPx;
    auto = { pos, alignment: 'left' };
  }

  if (custom && auto) {
    // Ties → the custom stop wins so its alignment/leader are honoured.
    return custom.pos <= auto.pos ? custom : auto;
  }
  return custom ?? auto;
}

interface NumberingMarkerLayout {
  numTab: number;
  picBullet: { bmp: DecodedImage; w: number; h: number } | null;
  numBodyOffset: number;
  markerJcShiftPx: number;
  hasMarker: boolean;
}

function resolveNumberingMarker(
  para: DocParagraph,
  state: RenderState,
  indLeft: number,
  indFirst: number,
): NumberingMarkerLayout {
  const { ctx, scale, fontFamilyClasses } = state;
  // Numbering marker. `hasMarker` is the "this paragraph has a marker" flag;
  // it is true for a text/glyph marker (`numMarker`) AND for a §17.9.9 picture
  // bullet (whose lvlText is typically empty — `numMarker` would be falsy).
  let numMarker = '';
  let numTab = 0;
  // §17.9.9/§17.9.20 — when the level uses a picture bullet, this holds its
  // decoded bitmap + draw size (px); the marker is drawn as an image, not text.
  let picBullet: { bmp: DecodedImage; w: number; h: number } | null = null;
  // First-line body offset (px) from paraX for an LTR numbered paragraph, set by
  // the §17.9.28 `<w:suff>` that follows the marker:
  //   tab (default) → body advances to the indentLeft tab stop (offset 0),
  //   space/nothing → body abuts the marker (marker end, + one space for space).
  let numBodyOffset = 0;
  // §17.9.8 `<w:lvlJc>` — horizontal shift (px) applied to the LTR marker draw so
  // it left/right/centre-aligns at the hanging-indent reference (firstLineX).
  // 0 = left (default); −markerW = right (period-aligned numerals: right edge at
  // firstLineX); −markerW/2 = centre. Set in the numbering block below.
  let markerJcShiftPx = 0;
  if (para.numbering) {
    numMarker = para.numbering.text;
    numTab = para.numbering.tab * scale;
    const suff = para.numbering.suff || 'tab';
    const pbPath = para.numbering.picBulletImagePath;
    if (pbPath) {
      const bmp = state.images.get(imageKey(pbPath));
      if (bmp) {
        // §17.9.20 — size from the bullet drawing's extent, else the resolved
        // marker font size (picBulletSizePt is the single source of truth shared
        // with the collect side; no magic pt default).
        const size = picBulletSizePt(para.numbering, para);
        picBullet = { bmp, w: size.w * scale, h: size.h * scale };
      }
    }
    // Marker glyph width (px) with its RESOLVED font (§17.3.2.26 + §17.9.6); the
    // picture bullet's own width when present. Needed for both the suff≠tab abut
    // and the suff=tab overrun check below, so measure once up front.
    let markerW: number;
    if (picBullet) {
      markerW = picBullet.w;
    } else {
      ctx.font = buildFont(false, false, getDefaultFontSize(para) * scale, markerFontFamily(para.numbering), fontFamilyClasses);
      markerW = ctx.measureText(markerDisplayText(para.numbering)).width;
    }
    // §17.9.8 lvlJc: shift the marker so its left/right/centre aligns at
    // firstLineX (the hanging-indent reference). The marker's RIGHT edge measured
    // from paraX (the indentLeft tab) is then `indFirst + shift + markerW`.
    const lvlJc = para.numbering.jc || 'left';
    markerJcShiftPx = lvlJc === 'right' ? -markerW : lvlJc === 'center' ? -markerW / 2 : 0;
    const markerEndFromIndent = indFirst + markerJcShiftPx + markerW;
    if (suff !== 'tab') {
      const spaceW = suff === 'space' ? ctx.measureText(' ').width : 0;
      // body abuts the marker's right edge (+ one space for suff="space").
      numBodyOffset = markerEndFromIndent + spaceW;
    } else {
      // suff=tab: the marker is followed by a tab that advances the body to the
      // numbering's indentLeft tab stop (numBodyOffset 0 — the body sits at
      // paraX). But ECMA-376 §17.9.6 + §17.3.1.37: a tab never moves BACKWARD, so
      // when the marker overruns that stop — a wide multi-level number like
      // "1.1.1." whose glyphs exceed the hanging indent (the marker `indFirst`
      // budget), e.g. in a substitute font — the tab advances to the next stop
      // PAST the marker end instead, and the body follows it. Without this the
      // body stays at indentLeft and the marker overprints it (sample-11's
      // "1.1.1. Three" collided; Word advances "Three" to the next default tab).
      // markerEndFromIndent (jc-adjusted right edge from paraX) ≤ 0 ⇒ it fits
      // (right-aligned markers always do), leave the body at indentLeft.
      if (markerEndFromIndent > 0) {
        // Next tab stop strictly past the marker end, resolved in TEXT-MARGIN
        // coordinates via the SAME helper as line layout (§17.3.1.37 +
        // §17.15.1.25): honour the paragraph's explicit stops (already in margin
        // px = pos * scale) plus the document's automatic grid AFTER all custom
        // stops, then convert back to paraX-relative px (− indLeft).
        const markerEndFromMargin = indLeft + markerEndFromIndent;
        const customStopsPx = (para.tabStops ?? []).map((ts) => ({
          pos: ts.pos * scale,
          alignment: ts.alignment,
          leader: ts.leader,
        }));
        const stop = nextTabStop(markerEndFromMargin, customStopsPx, state.defaultTabPt * scale);
        if (stop) numBodyOffset = stop.pos - indLeft;
      }
    }
  }
  // True when the paragraph has any marker to draw (text glyph OR picture bullet).
  const hasMarker = numMarker !== '' || picBullet !== null;
  return { numTab, picBullet, numBodyOffset, markerJcShiftPx, hasMarker };
}

/** Phase 4-1 B2 Stage 2 — rehydrate the paginator's scale-1 stamped lines into
 *  the paint scale, keeping the scale-1 line PARTITION but re-deriving all
 *  hinting-sensitive geometry at the paint scale.
 *
 *  Why not a pure ×scale of the scale-1 fields: a real (hinted) font's Canvas
 *  metrics are NOT scale-linear — `measureText(pt·s).width ≠ s · measureText(pt)`
 *  and `fontBoundingBoxAscent(pt·s) ≠ s · fontBoundingBoxAscent(pt)`. The draw
 *  path advances the pen by each segment's `measuredWidth` while `fillText`
 *  renders at the paint-scale font, and the line box uses `ascent/descent`; a
 *  geometric ×scale of those would drift the pen off the glyphs horizontally and
 *  the baseline off by up to ~1px vertically, accumulating down the page (seen on
 *  demo/sample-1 p.4). So RE-MEASURE every text segment at the paint scale here —
 *  the exact same per-segment measurement `layoutLines` does (advance via
 *  gridWidth, box via correctedLineMetrics, the §17.3.2.33 small-caps full-size
 *  box, the §17.3.3.25 ruby ascent reserve, and the intended-single-line floor) —
 *  so measure == draw byte-for-byte, identical to a fresh paint-scale layout of
 *  the SAME partition. Only WHICH glyphs sit on WHICH line comes from the scale-1
 *  stamp; that is the zoom-invariant part (Word lays text out in the document's
 *  coordinate space and the display scale is a viewport transform, not a
 *  re-layout — the scale-1 partition is correct at every zoom). This makes reuse
 *  a deliberate behaviour change vs. the old recompute-at-paint-scale path, which
 *  let hinting shift wrap points per zoom.
 *
 *  Geometry that IS scale-linear (a clean ×scale): the float-exclusion `xOffset` /
 *  `availWidth` / `topY` (pure page geometry, no glyph hinting) and non-text
 *  advances — image `widthPt·scale`, math `render.*Em·emPx`. Empty/synthetic
 *  lines (no text segment) fall back to the ×scale of the stamped box, matching
 *  layoutLines' `h·scale·0.8`/`0.2` synthesis.
 *
 *  Returns FRESH line + segment objects (never mutates the shared stamp — the
 *  draw path and repeated renderPage calls read the same immutable array; see
 *  layout-lines-reuse-identity.test.ts). `scale === 1` returns the stamp
 *  unchanged (identity — no copy, no re-measure), so the scale-1 reuse path stays
 *  byte-for-byte as in Stage 1. */
function rescaleLayoutLines(
  lines: LayoutLine[],
  scale: number,
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  fontFamilyClasses: Record<string, string>,
  gridDeltaPx: number,
): LayoutLine[] {
  if (scale === 1) return lines;

  // Per-text-segment measurement at the PAINT scale — the SAME model layoutLines
  // uses (segAdvance + the text-segment metric block), so measure == draw.
  const measureTextSeg = (s: LayoutTextSeg): { advance: number; asc: number; desc: number; intended: number } => {
    const effPx = calcEffectiveFontPx(s, scale);
    ctx.font = buildFont(s.bold, s.italic, effPx, s.fontFamily, fontFamilyClasses);
    const m = ctx.measureText(s.text);
    const advance = gridWidth(m.width, s.text, gridDeltaPx);
    // §17.3.2.33 — a small-caps run's LINE BOX follows the FULL run size (measure
    // the box at fullPx, not the 2pt-reduced glyphs); super/subscript keeps its
    // shrunk contribution. Mirrors layoutLines' fullPx / metricEmPx split.
    const fullPx = s.fontSize * scale;
    let metricM = m;
    let metricEmPx = effPx;
    if (s.smallCaps && !s.vertAlign && metricEmPx !== fullPx) {
      ctx.font = buildFont(s.bold, s.italic, fullPx, s.fontFamily, fontFamilyClasses);
      metricM = ctx.measureText(s.text || 'X');
      metricEmPx = fullPx;
    }
    const corrected = correctedLineMetrics(metricM, s.fontFamily, fullPx, metricEmPx);
    // §17.3.3.25 — ruby reserves extra ascent room (rt size × 1.5), same as layoutLines.
    const asc = s.ruby ? corrected.ascent + s.ruby.fontSizePt * scale * 1.5 : corrected.ascent;
    // Intended single-line floor (font-metrics.ts) — small caps keep the FULL run
    // size here too (addToLine's intendedEm).
    const intendedEm = s.smallCaps && !s.vertAlign ? fullPx : effPx;
    const intended = intendedSingleLinePx(s.fontFamily, intendedEm);
    return { advance, asc, desc: corrected.descent, intended };
  };

  return lines.map((l) => {
    let asc = 0;
    let desc = 0;
    let intended = 0;
    let hasText = false;
    const segments = l.segments.map((s) => {
      if ('isTab' in s) {
        // Tab advance is position-dependent (pen → next stop) and box-neutral; a
        // ×scale reproduces the scale-1 tab-to-stop advance scaled to paint space
        // (the stamp reuse is gated to no-float, non-marker paragraphs, so the
        // tab-stop grid scales cleanly with the box).
        return { ...s, measuredWidth: s.measuredWidth * scale };
      }
      if ('imagePath' in s) {
        // Anchored images live out of inline flow: layoutLines pins their
        // measuredWidth to 0 (they add no pen advance) and their anchor*Pt
        // fields stay in pt space for the draw-time position resolver.
        if (s.anchor) return { ...s, measuredWidth: 0 };
        return { ...s, measuredWidth: s.widthPt * scale };
      }
      if ('mathNodes' in s) {
        const copy = { ...s, measuredWidth: s.measuredWidth * scale } as LayoutMathSeg;
        copy.mathAscent *= scale;
        copy.mathDescent *= scale;
        asc = Math.max(asc, copy.mathAscent, s.fontSize * scale * 0.8);
        desc = Math.max(desc, copy.mathDescent, s.fontSize * scale * 0.2);
        return copy;
      }
      const t = s as LayoutTextSeg;
      const mm = measureTextSeg(t);
      hasText = true;
      if (mm.asc > asc) asc = mm.asc;
      if (mm.desc > desc) desc = mm.desc;
      if (mm.intended > intended) intended = mm.intended;
      return { ...t, measuredWidth: mm.advance };
    });
    // Empty/synthetic line (no text/math contributing metrics): fall back to the
    // ×scale of the stamped box (layoutLines synthesises h·scale·0.8/0.2 there).
    if (!hasText && asc === 0 && desc === 0) {
      asc = l.ascent * scale;
      desc = l.descent * scale;
      intended = l.intendedSingle * scale;
    }
    return {
      ...l,
      segments,
      ascent: asc,
      descent: desc,
      intendedSingle: intended,
      // Pure page geometry — scale-linear (no glyph hinting).
      xOffset: l.xOffset * scale,
      availWidth: l.availWidth * scale,
      topY: l.topY === undefined ? undefined : l.topY * scale,
    };
  });
}

function renderParagraph(
  para: DocParagraph,
  state: RenderState,
  suppressSpaceBefore = false,
  /** When set, render only `lines[start, end)` of the laid-out paragraph,
   *  used by the paginator to split paragraphs that don't fit on one page. */
  lineSlice?: { start: number; end: number },
  /** True when this call is the redirected draw of a `<w:framePr>` frame
   *  paragraph (from {@link renderFrameParagraph}). It suppresses the in-flow
   *  cursor bookkeeping that the frame path handles itself: anchor-float
   *  registration is skipped (the frame is the float) and the
   *  topAndBottom-skip / frame dispatch are bypassed (the geometry is already
   *  the frame box). Frame dispatch for a non-frame call lives in
   *  renderBodyElements so it can pass the anchor paragraph's line height. */
  inFrame = false,
  /** ECMA-376 §17.3.1.7 paragraph-border merge: suppress the top edge when a
   *  same-border paragraph precedes this one in the same column, and the bottom
   *  edge when one follows. Computed by the paint loop (renderBodyElements /
   *  renderParaList), which knows in-flow adjacency. Absent ⇒ draw the full box
   *  (a standalone bordered paragraph). */
  borderMerge?: ParaBorderMerge,
): void {
  const { ctx, scale, contentX, contentW, defaultColor, dryRun, fontFamilyClasses } = state;
  // Capture Y before spaceBefore — used for paragraph-relative anchor image positioning
  const paragraphStartY = state.y;

  if (!suppressSpaceBefore) state.y += para.spaceBefore * scale;

  // Register anchor floats from this paragraph. ECMA-376 §20.4.3.5: a
  // `positionV relativeFrom="paragraph"` float is positioned relative to "the
  // paragraph which contains the drawing anchor" — its TOP edge, BEFORE the
  // paragraph's spaceBefore (Word anchors the float at the paragraph top, not the
  // post-spaceBefore text area). So pass `paragraphStartY` (pre-spaceBefore),
  // identically for wrap AND wrapNone floats (renderAnchorImages below already
  // uses paragraphStartY). Anchoring wrap floats at the post-spaceBefore text top
  // placed them spaceBefore too low — e.g. sample-12's figure (anchor paragraph
  // spaceBefore=12 pt) sat 12 pt under Word, eating the gap above its caption.
  // Skipped for the frame-draw recursion: a frame paragraph's wrap exclusion is
  // its own FloatRect (renderFrameParagraph), not an anchor image/shape float.
  if (!inFrame) registerAnchorFloats(para, state, paragraphStartY);

  // behindDoc shapes must render before text so they appear behind it.
  renderAnchorImages(para, state, paragraphStartY, 'behind');

  // If any topAndBottom float already extends past state.y, skip past it before text starts.
  state.y = skipPastTopAndBottom(state.y, state.floats);

  const textAreaTopY = state.y;

  // ECMA-376 §17.3.1.12 w:ind — the transitional left/right attributes are
  // logical start/end (Part 4 §14.11.2). In a bidi paragraph the start side is
  // the physical RIGHT, so the two indents swap physical sides here.
  //
  // A frame paragraph's own body-style indents (e.g. a default first-line indent
  // inherited from the body style) do NOT apply to the frame content: the frame
  // box already positions the glyphs from its left edge, and the wrap exclusion
  // is built from that same left edge (§17.3.1.11). Honoring the indent here
  // would shift the cap glyph right of the exclusion band and let body text
  // overlap it, so zero the indents in the frame-draw recursion.
  const baseRtl = para.bidi === true;
  const indLeft = inFrame ? 0 : (baseRtl ? para.indentRight : para.indentLeft) * scale;
  const indRight = inFrame ? 0 : (baseRtl ? para.indentLeft : para.indentRight) * scale;
  const indFirst = inFrame ? 0 : para.indentFirst * scale;

  // Numbering marker layout (§17.9.x): see resolveNumberingMarker.
  const { numTab, picBullet, numBodyOffset, markerJcShiftPx, hasMarker } =
    resolveNumberingMarker(para, state, indLeft, indFirst);

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

  // A paragraph with no inline content (literally empty, or anchor-only) still
  // produces ONE paragraph-mark line box (ECMA-376 §17.3.1.29 regulates only the
  // existence of that line; the horizontal wrap geometry around a square float is
  // §20.4.2.17). The behavior below — firing an automatic "flow the mark line
  // below the float band when it cannot sit beside it" displacement, and using
  // ONE EM of the mark font as the width the gap must hold — is NOT specified by
  // ECMA-376 Part 1. It is an implementation-defined HEURISTIC chosen to match
  // Word: the only spec-mandated flow of a line onto a float-free region is the
  // explicit `<w:br w:clear>` of §17.18.3, which is not what fires here. Without
  // this heuristic an empty paragraph mark wedges into a sub-em sliver beside a
  // full-width float band and the following paragraphs (and any wrapNone image
  // they anchor) stay pinned inside the band. We resolve the mark line's flowed
  // top here and use it for the mark advance, the shading/border rect, and the
  // paragraph-relative base of any wrapNone anchor image drawn below.
  const resolveEmptyMarkTop = (): number => {
    if (state.floats.length === 0) return textAreaTopY;
    // Required width for an empty mark line: one em of the mark font (HEURISTIC,
    // see above — not a spec-defined threshold). A side gap narrower than this is
    // treated as unable to hold the line start, so the line flows below.
    const markEm = paragraphMarkEmPx(para, scale);
    const probeH = 10 * scale;
    const win = resolveLineFloatWindow(
      textAreaTopY, markEm, probeH, paraX, paraW, state.floats,
    );
    return win.topY;
  };

  if (segments.length === 0) {
    // Literally-empty paragraph: one paragraph-mark line box, no inline content
    // and (by construction in the paginator) never sliced.
    renderEmptyMarkParagraph(para, state, {
      grid, paraHasRuby, contentX, indLeft, paraW, textAreaTopY, paragraphStartY,
      markTop: resolveEmptyMarkTop(), totalLines: 0, lineSlice: undefined, borderMerge,
    });
    return;
  }

  const wrapCtx: WrapLayoutCtx | undefined = state.floats.length > 0 ? {
    startPageY: state.y,
    paraX,
    floats: state.floats,
    lineBoxH: (a, d, _h, is) => lineBoxHeight(para.lineSpacing, a, d, scale, grid, paraHasRuby, is ?? 0, paragraphIsEastAsian(para)),
    pageH: state.pageH,
    markEmPx: paragraphMarkEmPx(para, scale),
  } : undefined;

  // ECMA-376 §17.3.1.12 (hanging) + §17.3.1.38 (a hanging indent implicitly
  // creates a tab stop at indentLeft) + §17.9.28 (`<w:suff>`, default "tab"):
  // in a hanging-indent list the number glyph sits at firstLineX (= indentLeft −
  // hanging); with suff=tab it is followed by a tab that advances the body to the
  // indentLeft tab stop, so the first line's TEXT region matches the continuation
  // lines' ([paraX, paraX+paraW]) and the negative first-line indent positions
  // only the marker. With suff=space/nothing the body abuts the marker instead
  // (numBodyOffset, computed above). Non-numbered paragraphs apply the first-line
  // indent (positive firstLine, or a bare negative hanging without a marker) to
  // the body as usual. RTL lists keep their existing start-edge handling.
  const firstLineIndent = hasMarker && !baseRtl ? numBodyOffset : firstLineX - paraX;
  const paintGridDeltaPx = gridCharDeltaPx(grid, scale);
  // Phase 4-1 B2 Stage 2 — compute-once, ZOOM-INVARIANT reuse. When the paginator
  // split this paragraph it stamped the scale-1 lines it laid out
  // (splitParagraphAcrossPages). Reuse the scale-1 line PARTITION at ANY paint
  // scale — skipping this scale's line-BREAK decisions — and rehydrate to the
  // paint scale by RE-MEASURING each line's glyph geometry at the paint scale
  // (rescaleLayoutLines: advance + box + tabs, so measure == draw with no hinting
  // drift; scale 1 returns the stamp unchanged). This is a deliberate BEHAVIOUR
  // CHANGE from Stage 1, which reused only at scale 1 and otherwise re-ran the
  // break decisions at the paint scale: with a real (hinted) font those decisions
  // could move wrap points as the zoom changes, whereas Word lays text out in the
  // document's coordinate space and treats the display scale as a viewport
  // transform (the wrap partition is scale 1 at every zoom). Reuse is gated to the
  // cases where the scale-1 partition is the one this paint would intend — every
  // layout INPUT must match the paginator's, compared in the paginator's scale-1
  // space (the paint values are exactly `scale ×` the scale-1 values; see the
  // per-field derivation below). The input check still rejects the numbering
  // firstLineIndent case (measure uses para.indentFirst, paint uses numBodyOffset
  // for a marker) so those recompute. A float context is excluded outright: float
  // wrap depends on the page-absolute Y of THIS slice, which the stamped
  // (whole-paragraph, first-page) lines do not carry.
  //
  // Scale-1 input reconstruction: the pt sources are read straight off the
  // paragraph (para.indent*) — NO division — except paraW, whose only scaled term
  // is contentW (= colW·scale, no rounding), so contentW/scale recovers colW. At
  // scale 1 every reconstruction is the identity (÷1), so the gate stays
  // bit-exact there and the Stage-1 pixel-identity test is unaffected. paraW is
  // compared with a magnitude-relative epsilon because contentW/scale − indL − indR
  // and (colW − indL − indR) are two float paths to the same real width (round-off
  // ~1e-13); this is a geometric equality test, not a snapping heuristic.
  //
  // Segment stability: the reuse also assumes buildSegments(para.runs, ·) yields
  // the SAME segments under the paginator's measure state and this paint state.
  // buildSegments is pure over `para.runs` (both passes read the same paragraph
  // object) EXCEPT for two paint-context text sources — page/numPages fields and
  // noteRef labels — and paragraphs carrying those are never stamped in the
  // first place (paragraphSegsStateSensitive, checked at the stamp site), so no
  // per-segment gate comparison is needed here. If buildSegments ever gains a
  // new state-dependent text source, extend that predicate — this gate cannot
  // see inside the stamped segments.
  const stamped = para as unknown as PaginatedElementWithLines;
  // Reconstruct THIS paint's layout inputs in the paginator's scale-1 pt space.
  const paraW1 = contentW / scale - (inFrame ? 0 : (baseRtl ? para.indentRight : para.indentLeft)) - (inFrame ? 0 : (baseRtl ? para.indentLeft : para.indentRight));
  const indLeft1 = inFrame ? 0 : (baseRtl ? para.indentRight : para.indentLeft);
  const firstIndent1 = hasMarker && !baseRtl ? numBodyOffset / scale : para.indentFirst;
  const gridDelta1 = gridCharDeltaPx(grid, 1);
  const reuse =
    lineReuseEnabled &&
    stamped.layoutLines !== undefined &&
    stamped.layoutLinesInputs !== undefined &&
    stamped.layoutLinesInputs.scale === 1 &&
    !wrapCtx &&
    !stamped.layoutLinesInputs.hasFloats &&
    Math.abs(stamped.layoutLinesInputs.paraW - paraW1) <= 1e-6 * Math.max(1, Math.abs(paraW1)) &&
    stamped.layoutLinesInputs.firstIndent === firstIndent1 &&
    stamped.layoutLinesInputs.tabOriginPx === indLeft1 &&
    stamped.layoutLinesInputs.gridDeltaPx === gridDelta1 &&
    // §17.3.1.16 / §17.15.1.58–.59 — kinsoku governs CJK retract decisions in
    // layoutLines, so differing rules mean a (potentially) different partition.
    // Value equivalence, NOT `===`: the prebuiltPages path resolves the rules
    // independently in paginateDocument and here (fresh Sets per call), so the
    // references legitimately differ while the rules are identical.
    kinsokuRulesEquivalent(stamped.layoutLinesInputs.kinsoku, state.kinsoku);
  // Zoom-invariant line breaking (Phase 4-1 B2 Stage 2). Three cases, all feeding
  // the SAME draw loop below; rescaleLayoutLines re-measures the geometry at the
  // paint scale off the scale-1 PARTITION (so measure == draw, no hinting drift):
  //  1. reuse — the paginator's scale-1 stamp.
  //  2. no float context — recompute the partition in the paginator's SAME scale-1
  //     space (paraW1 / firstIndent1 / indLeft1 / gridDelta1). A paragraph that
  //     missed the stamp (never split, or its inputs differ — e.g. a numbered
  //     list's firstLineIndent) must STILL break at the scale-1 partition, or a
  //     page would MIX scale-1-broken (split) and paint-scale-broken (non-split)
  //     paragraphs — two wrap regimes side by side — and a paint-scale re-break
  //     can even overflow the height the paginator reserved (estimateParagraphHeight
  //     also lays out at scale 1). Keeping the whole non-float body on scale-1
  //     breaking makes the page coherent and paginate-aligned.
  //  3. float wrap context — stays at the paint scale (a straight layoutLines):
  //     the wrap windows are evaluated against paint-scale float rectangles at
  //     page-absolute Y, which have no scale-1 form here (the stamp reuse excludes
  //     floats for the same reason). Pre-existing paginate(scale 1)/paint(scale s)
  //     float behaviour, unchanged.
  const lines = reuse
    ? rescaleLayoutLines(stamped.layoutLines as LayoutLine[], scale, ctx, state.fontFamilyClasses, paintGridDeltaPx)
    : wrapCtx
      ? layoutLines(ctx, segments, paraW, firstLineIndent, scale, para.tabStops, wrapCtx, state.fontFamilyClasses, indLeft, state.kinsoku, paintGridDeltaPx, state.defaultTabPt)
      : rescaleLayoutLines(
          layoutLines(ctx, segments, paraW1, firstIndent1, 1, para.tabStops, undefined, state.fontFamilyClasses, indLeft1, state.kinsoku, gridDelta1, state.defaultTabPt),
          scale, ctx, state.fontFamilyClasses, paintGridDeltaPx,
        );

  // Decimal-tab auto-alignment. ECMA-376 (§17.3.1.37 tabs / §17.18.84 ST_TabJc
  // `decimal`) only positions content at a tab stop when an explicit tab
  // character advances to it; absent a tab, content starts at the indent. Word,
  // however, aligns a bare number to a leading DECIMAL tab with NO tab character
  // — the built-in "Decimal Aligned" paragraph style on table number cells does
  // exactly this. sample-11's College table proves it: 110 / 103 / +7 etc. each
  // right-align on the decimal tab at 18 pt (Word PDF bbox: per-column right
  // edges coincide), where we previously left-aligned them. This is a deliberate
  // Word-runtime deviation (user-approved); it is gated to NUMERIC content whose
  // first tab stop is `decimal` and which carries no explicit tab, so ordinary
  // paragraphs are untouched. We right-edge align the number at the stop — the
  // same approximation the explicit-tab decimal path uses (frac=1; it does not
  // split on the '.'), exact for the integers in scope. Applied at DRAW time as
  // a pure horizontal offset, so the measured row height (and the paginate/paint
  // height contract) is unaffected.
  const decimalAutoTabPx: number | null = (() => {
    if (segments.some((s) => 'isTab' in s)) return null; // explicit tab wins
    const stops = para.tabStops ?? [];
    if (stops.length === 0) return null;
    const firstStop = stops.reduce((a, b) => (b.pos < a.pos ? b : a));
    if (firstStop.alignment !== 'decimal') return null;
    const txt = para.runs.map((r) => (r as { text?: string }).text ?? '').join('').trim();
    if (txt === '' || !/^[+\-(]?[\d., ]+\)?%?$/.test(txt)) return null; // numbers only
    return firstStop.pos * scale - indLeft; // px, relative to paraX (mirrors layoutLines' stopXof)
  })();

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
    // Anchor-only paragraph: same content-less mark line as the literally-empty
    // path (the anchor floats themselves are drawn separately). Slice guards
    // honor a paginator-split slice (spaceAfter on the final slice, anchor
    // images on the first).
    renderEmptyMarkParagraph(para, state, {
      grid, paraHasRuby, contentX, indLeft, paraW, textAreaTopY, paragraphStartY,
      markTop: resolveEmptyMarkTop(), totalLines: lines.length, lineSlice, borderMerge,
    });
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
        Math.max(0, ...lines.map(l => lineBoxHeight(para.lineSpacing, l.ascent, l.descent, scale, grid, true, l.intendedSingle, paragraphIsEastAsian(para)))),
        grid,
        scale,
      )
    : 0;
  const lineHForLine = (l: typeof lines[number]): number =>
    paraHasRuby
      ? uniformLineH
      : lineBoxHeight(para.lineSpacing, l.ascent, l.descent, scale, grid, false, l.intendedSingle, paragraphIsEastAsian(para));

  // Slice bounds — when the paginator split this paragraph across pages,
  // only render lines in [sliceStart, sliceEnd). The first line we paint
  // resets state.y baseline so the slice begins at the page's content top.
  // Resolved BEFORE the shading fill so the fill height covers exactly the
  // lines this page paints (see paintedParagraphHeight): a sliced paragraph
  // must not fill to the full-paragraph height past the slice's bottom border.
  const sliceStart = lineSlice ? lineSlice.start : 0;
  const sliceEnd = lineSlice ? lineSlice.end : lines.length;
  // The slice is authoritative for WHICH lines land on this page, but THIS
  // pass's `lines` array is authoritative for how many lines the text actually
  // occupies at this scale (pagination lays out at scale 1; ctx.measureText is
  // not perfectly scale-invariant, so a long narrow paragraph can wrap to a
  // slightly different line count here than the scale-1 slice assumed). Cap the
  // iteration at lines.length so we paint every real line and never index a
  // phantom line that only existed in the scale-1 measurement (lines[i] would be
  // undefined → "Cannot read properties of undefined"). The overflow is bounded
  // (at most a line or two), so all the paragraph's text is still painted across
  // its slices. `paintEnd` also bounds the shading height below.
  const paintEnd = Math.min(sliceEnd, lines.length);

  if (para.shading && !dryRun) {
    // Shading is the BACKGROUND (text paints on top), so its height must be known
    // BEFORE the draw loop. Replay the loop's exact per-line advancement over the
    // painted slice so the fill height === the post-loop border height
    // (state.y − textAreaTopY) BY CONSTRUCTION — the fill meets the bottom border
    // in the float-clearance and page-slice cases too, not just top/left/right.
    const paintedH = paintedParagraphHeight(lines, sliceStart, paintEnd, textAreaTopY, lineHForLine);
    ctx.fillStyle = `#${para.shading}`;
    const sb = paraShadingRect(contentX + indLeft, textAreaTopY, paraW, paintedH, para.borders, borderMerge, scale);
    ctx.fillRect(sb.x, sb.y, sb.w, sb.h);
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

  // Bidirectional text. The paragraph's base direction comes from w:bidi
  // (ECMA-376 §17.3.1.6). We engage the (exact) bidi pass only when the base is
  // RTL or the line actually contains strong-RTL characters, so pure-LTR
  // paragraphs keep their byte-identical fast path. `alignEdge` resolves
  // logical start/end against the base direction. (`baseRtl` is declared with
  // the indent swap above.)
  const paraNeedsBidi = baseRtl || segmentsHaveRtl(segments);
  const alignEdge = resolveAlignEdge(para.alignment, baseRtl);

  // ECMA-376 §17.6.5 character-grid delta (px per EA glyph) for the DRAW pass —
  // the SAME value layoutLines folded into measuredWidth. A pure-EA segment is
  // drawn so its glyphs occupy exactly `measuredWidth` (= natural + len·Δ): the
  // draw uses `justifiedPiecePositions(..., letterSpacingPx = Δ)`, whose final
  // glyph lands on the box edge, so the painted advance equals measuredWidth by
  // construction. See the gridCharDeltaPx / gridSegDeltaPx header.
  const drawGridDeltaPx = gridCharDeltaPx(grid, scale);
  const drawCtx: ParagraphLineDrawCtx = { ctx, scale, state, para, dryRun, defaultColor, fontFamilyClasses, contentX, contentW, lines, grid, paraX, firstLineX, paraW, indLeft, indFirst, baseRtl, hasMarker, numTab, numBodyOffset, markerJcShiftPx, picBullet, isJustified, stretchLastLine, alignEdge, paraNeedsBidi, decimalAutoTabPx, drawGridDeltaPx, lineHForLine };
  for (let li = sliceStart; li < paintEnd; li++) {
    drawParagraphLine(li, drawCtx);
  }

  if (para.borders && !dryRun) {
    // `state.y` started this pass at `textAreaTopY` (captured above, untouched
    // until the draw loop) and the loop advanced it by exactly the per-line steps
    // `paintedParagraphHeight` replays (the topY float-clearance max-jump, then
    // `+= lineHForLine`). So `textH` here equals the `paintedH` the shading fill
    // used above — the fill meets this bottom border by construction, in the
    // normal, float-clearance and page-slice cases alike.
    const textH = state.y - textAreaTopY;
    drawParaBorders(ctx, contentX + indLeft, textAreaTopY, paraW, textH, para.borders, scale, state.dpr, borderMerge);
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

/** Per-line draw context for {@link drawParagraphLine}. Bundles the read-only
 *  paragraph-scope values the line loop reads (plus `state`, mutated by
 *  reference via `state.y`). Extracted from {@link renderParagraph} verbatim so
 *  the per-line drawing is a single thin call; no behaviour change. */
interface ParagraphLineDrawCtx {
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
  scale: number;
  state: RenderState;
  para: DocParagraph;
  dryRun: boolean;
  defaultColor: string;
  fontFamilyClasses: Record<string, string>;
  contentX: number;
  contentW: number;
  lines: LayoutLine[];
  grid: DocGridCtx;
  paraX: number;
  firstLineX: number;
  paraW: number;
  indLeft: number;
  indFirst: number;
  baseRtl: boolean;
  hasMarker: boolean;
  numTab: number;
  numBodyOffset: number;
  markerJcShiftPx: number;
  picBullet: { bmp: DecodedImage; w: number; h: number } | null;
  isJustified: boolean;
  stretchLastLine: boolean;
  alignEdge: ReturnType<typeof resolveAlignEdge>;
  paraNeedsBidi: boolean;
  decimalAutoTabPx: number | null;
  drawGridDeltaPx: number;
  lineHForLine: (l: LayoutLine) => number;
}

/** Draws line `li` of a paragraph. Extracted from {@link renderParagraph}'s
 *  per-line loop body verbatim (the loop simply calls this for each line);
 *  `state.y` is advanced by reference, exactly as before. */
function drawParagraphLine(li: number, c: ParagraphLineDrawCtx): void {
  const {
    ctx, scale, state, para, dryRun, defaultColor, fontFamilyClasses,
    contentX, contentW, lines, grid, paraX, firstLineX, paraW, indLeft,
    indFirst, baseRtl, hasMarker, numTab, numBodyOffset, markerJcShiftPx,
    picBullet, isJustified, stretchLastLine, alignEdge, paraNeedsBidi,
    decimalAutoTabPx, drawGridDeltaPx, lineHForLine,
  } = c;
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
    // For a numbered first line (LTR) the body sits at lineLeft + numBodyOffset
    // (the indentLeft tab stop for suff=tab → offset 0; the marker end for
    // space/nothing); indFirst only pulls the marker into the hanging margin
    // (drawn below). Non-numbered first lines apply indFirst to the body directly.
    let x = firstLine && !baseRtl ? (hasMarker ? lineLeft + numBodyOffset : lineLeft + indFirst) : lineLeft;
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
    // The visually-last segment (its trailing edge is the line's physical end, so
    // no gap opens there). Consumed ONLY on the bidi path of the justification
    // distribution below: an LTR line excludes no segment (the kernel's content-
    // span trim already closes the final glyph's gap). Equals the logical last in
    // the LTR fast path.
    const lastDrawnSi = visual ? visual.order[segCount - 1] : segCount - 1;

    const lineWidth = line.segments.reduce((s, seg) => s + seg.measuredWidth, 0);
    const lineSlack = effAvailW - (x - lineLeft) - lineWidth;
    // §17.18.44: a `both` line justifies UNLESS it is the paragraph's true last
    // line OR ends at a manual `<w:br/>` (§17.3.3.1) — both terminate a logical
    // line and are left-aligned. `distribute` (stretchLastLine) still spreads
    // every line, including these.
    const endsLogicalLine = isLastLine || (line.endsWithBreak ?? false);
    const applyJustify = isJustified && (!endsLogicalLine || stretchLastLine);
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
    // Decimal-tab auto-alignment (see decimalAutoTabPx above): override the
    // paragraph alignment so the number's right edge (its decimal point, for an
    // integer) lands on the decimal tab. `paraX + decimalAutoTabPx` is the stop
    // in device space; subtracting the line width and the current `x` yields the
    // left-shift, clamped ≥ 0 so a number wider than the stop simply overflows
    // right (never pulled left of its natural start).
    if (decimalAutoTabPx != null && lineWidth > 0) {
      alignOffset = Math.max(0, paraX + decimalAutoTabPx - lineWidth - x);
    }
    x += alignOffset;

    if (firstLine && hasMarker && !dryRun) {
      if (picBullet) {
        // §17.9.9/§17.9.20 — the marker is an image. It occupies the same
        // anchor a text marker would (LTR: left edge in the hanging margin;
        // RTL: right edge numTab past the start edge), and rides the line's jc
        // alignment via `x`/`lineWidth` exactly like the glyph marker below.
        // Vertically it bottom-aligns to the baseline (the inline-image
        // convention, §17.3.3 anchored to the text bottom) so a sub-em bullet
        // rests on the line like a glyph.
        const { bmp, w, h } = picBullet;
        const top = baseline - h;
        // LTR: left edge at firstLineX, shifted by lvlJc (§17.9.8); RTL keeps its
        // own mirrored anchor.
        const left = baseRtl ? x + lineWidth + numTab - w : lineLeft + indFirst + markerJcShiftPx;
        ctx.drawImage(bmp, left, top, w, h);
      } else {
        const numFontSize = getDefaultFontSize(para) * scale;
        // Draw the marker with its RESOLVED font (§17.3.2.26 + §17.9.6): the
        // ascii axis for a Latin number (a decimal "1" → Times → serif), the
        // eastAsia axis for a CJK marker. Replaces the old hardcoded sans-serif,
        // which forced every number/bullet sans regardless of the heading's font.
        ctx.font = buildFont(false, false, numFontSize, markerFontFamily(para.numbering!), fontFamilyClasses);
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
          const markerText = markerDisplayText(para.numbering!);
          const markerW = ctx.measureText(markerText).width;
          ctx.fillText(markerText, x + lineWidth + numTab - markerW, baseline);
          ctx.textAlign = prevAlign;
          ctx.direction = prevDir;
        } else {
          // Marker sits in the hanging margin at lineLeft + indFirst (= firstLineX
          // when the line isn't shifted by a float; lineLeft already includes any
          // float xOffset, so the marker tracks the body that hangs off it),
          // shifted by lvlJc (§17.9.8) so a "right" level period-aligns its right
          // edge at firstLineX. The body was advanced past the marker above
          // (numBodyOffset).
          ctx.fillText(markerDisplayText(para.numbering!), lineLeft + indFirst + markerJcShiftPx, baseline);
        }
      }
    }

    // Justified-line slack distribution (ECMA-376 §17.18.44). Positive slack
    // (lineWidth < availW) expands the line to fill the margin; negative slack
    // (lineWidth > availW, typically from canvas measuring ~1px wider than Word)
    // compresses it so the final glyph lands on the right margin instead of
    // overflowing. Gaps open at inter-word spaces AND — for expansion — inter-CJK
    // boundaries, so a pure-CJK `both`/`distribute` line fills the margin too
    // (Word fills CJK `both` lines by adding inter-character pitch; see
    // text-distribute.ts). distributeLineSlack returns, per logical segment, the
    // internal split points and a trailing-gap flag; the draw loop applies
    // `perGap` at each. Only computed when the line is a justify candidate
    // (jc=both/distribute, not the last line unless distribute).
    let segStretch: Map<number, SegStretch> | null = null;
    let distPerGap = 0;
    // First content segment in reading order. Leading-whitespace segments before
    // it (a paragraph's 字下げ indent) are NOT stretched — Word keeps the indent
    // fixed and distributes slack only across the line content (§17.18.44). Only
    // meaningful for LTR: under bidi the logical-leading segment is not the
    // visually-leading one, so leave the skip off (0) there.
    let firstContentSi = 0;
    if (applyJustify) {
      if (!paraNeedsBidi) {
        for (let i = 0; i < segCount; i++) {
          const seg = line.segments[i];
          if (!('text' in seg) || /\S/.test((seg as LayoutTextSeg).text)) { firstContentSi = i; break; }
        }
      }
      const slack = effAvailW - (x - lineLeft) - lineWidth;
      // Compression cap (negative slack): never eat more than ~a quarter em per
      // gap, estimated from the line ascent. For expansion this is unbounded.
      const minPerGap = -line.ascent * 0.25;
      // Expansion opens inter-CJK boundaries; compression touches only spaces
      // (shrinking a space is fine, overlapping ideographs is not).
      const distSegs = line.segments.map(seg =>
        'text' in seg ? { text: (seg as LayoutTextSeg).text } : {},
      );
      const dist = distributeLineSlack(
        distSegs,
        slack,
        firstContentSi,
        // §17.18.44 spreads the slack across EVERY inter-CJK boundary on the line,
        // so the visually-last segment must still distribute pitch INTERNALLY when
        // it is a multi-glyph CJK run — else it stays at the bare grid pitch while
        // earlier segments absorb all the slack (two pitches on one line). The
        // kernel already closes only the FINAL glyph's gap via its content-span
        // trim, so an LTR line excludes no segment: pass `segCount`, the no-match
        // sentinel the pptx justifier also uses. Bidi keeps the whole-segment
        // exclusion because its logical-last unit is not the visually-last glyph.
        // See the `lastDrawnSi` option doc in core/src/text/line-distribute.ts.
        paraNeedsBidi ? lastDrawnSi : segCount,
        minPerGap,
        slack > 0,
      );
      segStretch = dist ? dist.perSeg : null;
      distPerGap = dist ? dist.perGap : 0;
    }

    // ECMA-376 §17.3.2.4 (`<w:bdr>`): adjacent runs whose border attribute set
    // is identical form ONE run-border group and are "rendered within the same
    // set of borders". Accumulate the group's pixel extent as segments are
    // drawn left-to-right and stroke a single frame when the group ends (a
    // segment with a different / absent border, or end of line). Grouping is by
    // visual adjacency within this line; mixed-direction lines are an edge case
    // (the spec phrases the group in logical order) left for a follow-up.
    interface OpenBorderGroup {
      border: DocxRunBorder;
      left: number; right: number; top: number; bottom: number;
    }
    let borderGroup: OpenBorderGroup | null = null;
    const flushBorderGroup = () => {
      if (!borderGroup) return;
      const g = borderGroup;
      borderGroup = null;
      const bw = Math.max(1, g.border.width * scale); // w:sz/8 is in pt
      const sp = (g.border.space ?? 0) * scale;
      ctx.strokeStyle = g.border.color ? `#${g.border.color}` : defaultColor;
      ctx.lineWidth = bw;
      ctx.strokeRect(
        g.left - sp,
        g.top - sp,
        g.right - g.left + 2 * sp,
        g.bottom - g.top + 2 * sp,
      );
    };

    for (let vi = 0; vi < segCount; vi++) {
      const si = visual ? visual.order[vi] : vi;
      const seg = line.segments[si];
      if (visual) ctx.direction = visual.rtl[si] ? 'rtl' : 'ltr';
      // A non-text segment (tab / inline image / math) breaks run-border
      // adjacency (§17.3.2.4 groups adjacent *runs*), so close any open frame.
      if (!('text' in seg)) flushBorderGroup();
      if ('isTab' in seg) {
        // Tabs render as blank space, optionally filled with a leader (TOC dots etc.).
        if (!dryRun && seg.leader && seg.leader !== 'none' && seg.measuredWidth > 1) {
          drawTabLeader(ctx, seg.leader, x, baseline, seg.measuredWidth, seg.fontSize * scale, defaultColor, seg.bold, seg.italic);
        }
        x += seg.measuredWidth;
        continue;
      }
      if ('imagePath' in seg) {
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
      // Justification stretch for THIS segment (logical index si). `internalStretch`
      // is the px added between the segment's own glyphs (inter-CJK boundaries);
      // `splitBefore` lists the code-point offsets to advance `distPerGap` at while
      // drawing. `spanW` (the glyph advance + internalStretch) covers every glyph
      // and the interior pitch; `decoW` (below) additionally covers the segment's
      // own widened trailing SPACE so run decorations stay gap-free under `both`.
      const stretch = segStretch?.get(si);
      const internalStretch = stretch?.internalStretch ?? 0;
      if (!dryRun) {
        const effSizePx = calcEffectiveFontPx(s, scale);
        const yOffset = s.vertAlign === 'super'
          ? -s.fontSize * scale * 0.35
          : s.vertAlign === 'sub'
            ? s.fontSize * scale * 0.15
            : 0;
        ctx.font = buildFont(s.bold, s.italic, effSizePx, s.fontFamily, fontFamilyClasses);

        // Width spanned by the glyphs after justification, for ruby centring /
        // onTextRun reporting.
        const spanW = s.measuredWidth + internalStretch;

        // Width spanned by EVERY run decoration (highlight §17.3.1.15, shading
        // §17.3.2.32, border §17.3.2.4, underline §17.3.2.40, strike §17.3.2.37).
        // On a `both`/`distribute` line (§17.18.44) the justifier widens this
        // segment's TRAILING SPACE by `distPerGap` and advances the pen past it
        // (`x += distPerGap`, below). That space belongs to THIS run, so Word
        // paints its decorations across the widened advance — otherwise a GAP
        // opens between words (the bug this fixes). We extend ONLY when the
        // segment actually ends in whitespace it owns: an inter-CJK boundary gap
        // (no trailing space, e.g. a run/script split between two ideographs) is
        // NOT owned by either run, so extending there would bleed a highlight
        // past its run. Disabled under bidi: the pen advances in visual order
        // while `trailingGap` is a logical-order flag, so the widened gap is not
        // reliably the segment's own physical-right edge (kept at `spanW`, the
        // pre-justify-slack behaviour — no regression, just no slack fill).
        const ownsTrailingSlack =
          !!stretch?.trailingGap && !paraNeedsBidi && /\s$/.test(s.text);
        const decoW = spanW + (ownsTrailingSlack ? distPerGap : 0);

        // Glyph box used by every run-level box decoration (highlight fill,
        // §17.3.2.32 shading fill, §17.3.2.4 border): same vertical extent of
        // ~0.85em above the baseline to ~0.25em below it. Computed once so the
        // three decorations stay byte-identical (no duplicated 0.85 / 1.1).
        const boxTop = baseline + yOffset - effSizePx * 0.85;
        const boxHeight = effSizePx * 1.1;

        if (s.highlight) {
          ctx.fillStyle = HIGHLIGHT_COLORS[s.highlight] ?? '#FFFF00';
          ctx.fillRect(x, boxTop, decoW, boxHeight);
        }

        // ECMA-376 §17.3.2.32 run shading fill (`<w:shd w:fill>`): a solid
        // background rect behind the glyphs. Used for inverse video (black fill
        // + automatic = white text). Same rect geometry as the highlight box.
        if (s.background) {
          ctx.fillStyle = `#${s.background}`;
          ctx.fillRect(x, boxTop, decoW, boxHeight);
        }

        // ECMA-376 §17.3.2.4 run border (`<w:bdr>`, "box"): a rectangle around
        // the run, inflated by w:space (pt → px), drawn after the background so
        // the box outlines the filled area. Per the spec, adjacent runs sharing
        // an identical border render within the SAME frame, so instead of
        // stroking here we extend (or open) the current border group; the frame
        // is stroked by flushBorderGroup() when the group ends. The box bounds
        // each segment's glyph box (same rect the shading uses) unioned across
        // the group, so a mixed-size group still encloses every run.
        const activeBorder =
          s.border && s.border.style !== 'none' && s.border.style !== 'nil'
            ? s.border
            : null;
        if (activeBorder) {
          const segTop = boxTop;
          const segBottom = segTop + boxHeight;
          if (borderGroup && runBordersEqual(borderGroup.border, activeBorder)) {
            borderGroup.right = x + decoW;
            borderGroup.top = Math.min(borderGroup.top, segTop);
            borderGroup.bottom = Math.max(borderGroup.bottom, segBottom);
          } else {
            flushBorderGroup();
            borderGroup = {
              border: activeBorder,
              left: x, right: x + decoW, top: segTop, bottom: segBottom,
            };
          }
        } else {
          flushBorderGroup();
        }

        // Track-changes overlay: paint insertions / deletions in the author's
        // colour with the canonical Word markup (underline for insertions,
        // strikethrough for deletions). The author hash gives stable colours
        // for the same reviewer across pages. Disabled when
        // `showTrackChanges: false` (the "Final / No Markup" view).
        const revActive = state.showTrackChanges && !!s.revision;
        const revColor = revActive ? authorColor(s.revision!.author) : null;
        let glyphColor: string;
        if (revColor) {
          glyphColor = revColor;
        } else if (s.color) {
          glyphColor = `#${s.color}`;
        } else if (s.colorAuto) {
          // ECMA-376 §17.3.2.6 (w:color) / ST_HexColorAuto §17.18.39: automatic
          // color picks black/white for contrast against the effective
          // background. The black/white pick is implementation-defined (no
          // normative algorithm) — delegated to core's autoContrastColor.
          // TODO: the fully-conformant effective background also folds in
          // paragraph-level shading (`<w:pPr><w:shd>`) and table cell shading.
          // The draw loop currently only has the run shading at this point,
          // which covers the inverse-video case (run `w:shd w:fill="000000"`).
          // Paragraph/cell shading should be threaded into
          // `LayoutTextSeg.background` when dark enough to flip auto text white.
          glyphColor = autoContrastColor(s.background ?? null);
        } else {
          glyphColor = defaultColor;
        }
        ctx.fillStyle = glyphColor;
        // Draw the glyphs. Three cases, all anchored to the WHOLE-string
        // cumulative advance so the browser's contextual CJK metrics (most
        // visibly 約物半角, the half-width collapse of （「」。）) are honoured and
        // the painted advance equals the segment's box exactly:
        //   1. Character grid active on a pure-EA segment (segGridDelta !== 0):
        //      walk every glyph, advancing each to its cell start
        //      `measure(prefix) + i·Δ + justGaps·perGap`. The final glyph lands so
        //      the segment edge is measure(whole) + len·Δ + nGaps·perGap =
        //      measuredWidth + internalStretch — measure==draw by construction
        //      (§17.6.5). Folds in any justification pitch at the same time.
        //   2. Justified inter-CJK pitch only (no grid): the existing
        //      `justifiedPiecePositions` slice-at-gaps path.
        //   3. Neither: a single fillText (the common path).
        const segGridDelta = gridSegDeltaPx(s.text, drawGridDeltaPx);
        if (segGridDelta !== 0) {
          const cps = [...s.text]; // code points (handles surrogate pairs)
          // Draw each CONTIGUOUS piece (sliced only at justify gaps) as ONE
          // contextually-shaped `fillText`, with the per-EA-glyph grid delta
          // applied via `ctx.letterSpacing`. The previous per-code-point loop
          // painted each glyph ISOLATED (no contextual shaping) yet positioned
          // glyph i by the CONTEXTUAL cumulative `measureText(prefix_i)`. JIS X
          // 4051 約物連続 packing compresses a closing-class punctuation immediately
          // followed by an opening bracket ("：［", "、［", "）（") ~half-em in
          // measureText (a bracket next to a plain kanji/kana does NOT pack), so an
          // isolated full-width bracket plus that collapsed cumulative measure
          // pulled the following glyph half-em left, OVERLAPPING the bracket.
          // Drawing the piece contiguously makes measure and draw shape the SAME
          // way (the packing honoured ⇒ no overlap), and
          // `letterSpacing = Δ` reproduces the per-cell delta the box was measured
          // with. Build `pieces` BEFORE setting letterSpacing: justifiedPiecePositions
          // is eager and its internal `measure` calls must run at the natural
          // advance (it adds `from·Δ` itself; the canvas adds Δ between glyphs
          // WITHIN each piece — together glyph i lands at measure(prefix)+i·Δ, the
          // same target as before). See @silurus/ooxml-core → justify-positions.ts.
          const measure = (str: string): number => ctx.measureText(str).width;
          const pieces = justifiedPiecePositions(
            cps,
            stretch?.splitBefore ?? [],
            distPerGap,
            measure,
            drawGridDeltaPx,
          );
          const prevLetterSpacing = ctx.letterSpacing;
          ctx.letterSpacing = `${drawGridDeltaPx}px`;
          for (const { text: piece, dx } of pieces) {
            ctx.fillText(piece, x + dx, baseline + yOffset);
          }
          ctx.letterSpacing = prevLetterSpacing;
        } else if (stretch && stretch.splitBefore.length > 0) {
          // ECMA-376 §17.18.44 `both`/`distribute` inter-CJK justification pitch.
          // Anchor each sliced piece to the WHOLE-string cumulative advance plus
          // the accumulated pitch, instead of summing the isolated pieces'
          // advances. That sum drifts wider than the segment's box and would paint
          // the next run over this segment's tail (most visible at a CJK→Latin
          // boundary). See `@silurus/ooxml-core` → text/justify-positions.ts.
          const cps = [...s.text]; // code points (handles surrogate pairs)
          if (stretch.splitBefore.length === cps.length - 1) {
            // FULLY distributed: a gap was opened at EVERY inter-glyph boundary
            // (pure-CJK justify), so the pitch is UNIFORM. Drawing one glyph per
            // piece in isolation (the loop below) loses the browser's contextual
            // packing — JIS X 4051 約物連続 compresses a closing-class punctuation
            // immediately followed by an opening bracket ("：［", "、［", "）（")
            // ~half-em in measureText — so the bracket paints full-width while the
            // next glyph, positioned by the COLLAPSED cumulative measure, is pulled
            // left and OVERLAPS it (the justify analog of the docGrid case-1 fix,
            // PR #626).
            // Draw the whole CONTEXTUALLY-shaped run in ONE fillText with
            // ctx.letterSpacing = distPerGap: glyph i lands at measure(prefix_i) +
            // i·distPerGap (the exact justified position), so the final glyph
            // reaches measure(whole) + (n-1)·distPerGap = the segment box edge
            // (= internalStretch). Restore the prior letterSpacing afterwards; no
            // measureText runs inside the set/restore window.
            const prevLetterSpacing = ctx.letterSpacing;
            ctx.letterSpacing = `${distPerGap}px`;
            ctx.fillText(s.text, x, baseline + yOffset);
            ctx.letterSpacing = prevLetterSpacing;
          } else {
            const measure = (str: string): number => ctx.measureText(str).width;
            for (const { text: piece, dx } of justifiedPiecePositions(
              cps,
              stretch.splitBefore,
              distPerGap,
              measure,
            )) {
              ctx.fillText(piece, x + dx, baseline + yOffset);
            }
          }
        } else {
          ctx.fillText(s.text, x, baseline + yOffset);
        }

        // Ruby annotation: small text centered above the base glyphs.
        if (s.ruby) {
          const rubySizePx = s.ruby.fontSizePt * scale;
          const rubyFont = buildFont(s.bold, s.italic, rubySizePx, s.fontFamily, fontFamilyClasses);
          ctx.save();
          ctx.font = rubyFont;
          const rubyW = ctx.measureText(s.ruby.text).width;
          const rubyX = x + (spanW - rubyW) / 2;
          // Sit the ruby's baseline a small gap above the base ascent so the
          // characters don't touch. fillText baseline is at the line of the
          // characters, so subtract the ruby descent + small gap from the
          // base's ascent line to position correctly.
          const rubyBaseline = baseline + yOffset - effSizePx * 0.85 - rubySizePx * 0.1;
          // Ruby shares glyphColor, so a track-changes run's ruby inherits the
          // author revision color (a behavior change from previously ignoring
          // revColor for ruby).
          ctx.fillStyle = glyphColor;
          ctx.fillText(s.ruby.text, rubyX, rubyBaseline);
          ctx.restore();
        }

        if (state.onTextRun && s.text) {
          state.onTextRun({
            text: s.text,
            x,
            y: state.y,
            w: spanW,
            h: lineH,
            fontSize: effSizePx,
            font: ctx.font,
          });
        }

        // Underline / strike share the glyph colour, so an inverse-video run
        // (automatic colour on a dark background) draws a white rule too.
        const lineColor = glyphColor;
        const lineW = Math.max(0.5, effSizePx * 0.05);
        // Crispness nudge (see crispOffset): underline / strike-through are
        // horizontal strokes; each snaps onto the nearest crisp device row from
        // its own y (an odd device-width one would otherwise straddle two rows).
        // Compute the offset per line because each stroke sits at a different y.
        // Underline / strike run the SAME `decoW` as the box decorations: the
        // segment's grid-aware advance (§17.6.5) plus the interior + owned
        // trailing-space justification pitch. Word runs the rule under a run's
        // spaces (incl. their justified widening), so the line decoration tracks
        // the drawn advance and stays flush with the box fills (one width concept
        // for every decoration, matching the pptx renderer's run rules).
        const isInsertion = revActive && s.revision?.kind === 'insertion';
        const isDeletion = revActive && s.revision?.kind === 'deletion';

        if (s.underline || isInsertion) {
          ctx.strokeStyle = lineColor;
          ctx.lineWidth = lineW;
          const uyRaw = baseline + yOffset + effSizePx * 0.12;
          const uy = uyRaw + crispOffset(uyRaw, lineW, state.dpr);
          ctx.beginPath(); ctx.moveTo(x, uy); ctx.lineTo(x + decoW, uy); ctx.stroke();
        }

        if (s.strikethrough || isDeletion) {
          ctx.strokeStyle = lineColor;
          ctx.lineWidth = lineW;
          const syRaw = baseline + yOffset - effSizePx * 0.3;
          const sy = syRaw + crispOffset(syRaw, lineW, state.dpr);
          ctx.beginPath(); ctx.moveTo(x, sy); ctx.lineTo(x + decoW, sy); ctx.stroke();
        }

        if (s.doubleStrikethrough) {
          ctx.strokeStyle = lineColor;
          ctx.lineWidth = lineW;
          const sy1Raw = baseline + yOffset - effSizePx * 0.35;
          const sy2Raw = baseline + yOffset - effSizePx * 0.22;
          const sy1 = sy1Raw + crispOffset(sy1Raw, lineW, state.dpr);
          const sy2 = sy2Raw + crispOffset(sy2Raw, lineW, state.dpr);
          ctx.beginPath(); ctx.moveTo(x, sy1); ctx.lineTo(x + decoW, sy1); ctx.stroke();
          ctx.beginPath(); ctx.moveTo(x, sy2); ctx.lineTo(x + decoW, sy2); ctx.stroke();
        }
      }

      // Advance the pen past the segment's glyphs plus any internal justification
      // pitch added between them.
      x += s.measuredWidth + internalStretch;
      // Trailing inter-segment gap (an inter-word space or inter-CJK boundary at
      // this segment's edge), applied AFTER the segment so the next one starts
      // shifted. distributeLineSlack only sets trailingGap on gap-opening
      // segments — never the visually-final segment or a leading-indent segment —
      // so the final glyph still lands on the margin (Σgaps == slack).
      if (stretch?.trailingGap) x += distPerGap;
    }
    // End of line closes any open run-border group: a frame never spans lines
    // (each line wrap starts a fresh box on the next line).
    flushBorderGroup();
    if (paraNeedsBidi) ctx.direction = 'ltr'; // reset for subsequent draws

    state.y += lineH;
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
  /** This segment is GLUED to the preceding one (no inter-segment break): they
   *  are case-pieces of the same word emitted at different sizes for small caps
   *  (§17.3.2.33) — e.g. "I"(full)+"NTRODUCTION"(reduced). The line breaker must
   *  not start a new line before a glued segment; it retracts the whole glued
   *  group instead, so a small-caps word never splits across lines. */
  joinPrev?: boolean;
  doubleStrikethrough?: boolean;
  highlight?: string | null;
  /** ECMA-376 §17.3.2.32 `<w:shd w:fill>` — run shading fill (hex 6). Painted as
   *  a solid rect behind the glyphs; also the effective background that an
   *  automatic text color resolves against. */
  background?: string | null;
  /** ECMA-376 §17.3.2.6 — run carries `<w:color w:val="auto"/>`. The glyph
   *  color is resolved from {@link LayoutTextSeg.background} for contrast
   *  (implementation-defined black/white pick; no normative algorithm). */
  colorAuto?: boolean;
  /** ECMA-376 §17.3.2.4 `<w:bdr>` — a run-level border (box) around the text. */
  border?: DocxRunBorder | null;
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
  /** Bold/italic of the run carrying the tab (ECMA-376 §17.3.1.37 — the leader
   *  characters take the formatting of the tab's run, e.g. a bold TOC1 entry's
   *  dot leader is bold). Threaded so {@link drawTabLeader} can match the font. */
  bold?: boolean;
  italic?: boolean;
}

interface LayoutImageSeg {
  /** Zip path of the blip — also the `'imagePath' in seg` discriminant that
   *  distinguishes an image segment from text/math/tab segments. */
  imagePath: string;
  /** MIME type of the blip at {@link LayoutImageSeg.imagePath}. */
  mimeType: string;
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
  /** ECMA-376 §20.1.8.55 `<a:srcRect>` source-rectangle crop (fractions 0..1 of
   *  the decoded bitmap). When present the draw paths use the 9-arg
   *  `drawImage` to blit only `[l, t, 1−r, 1−b]` of the bitmap into the display
   *  box. `undefined` ⇒ draw the full bitmap. */
  srcRect?: { l: number; t: number; r: number; b: number };
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
  /** ECMA-376 §17.3.3.1 — this line is terminated by a MANUAL line break
   *  (`<w:br w:type="textWrapping"/>`). In a justified (`both`) paragraph it is
   *  the end of a logical line and must be left-aligned, not stretched — exactly
   *  like the paragraph's final line (§17.18.44). */
  endsWithBreak?: boolean;
}

/** Phase 4-1 B2 Stage 1 — a paginated element carrying the paginator's scale-1
 *  laid-out lines (compute-once). `splitParagraphAcrossPages` stamps the FULL
 *  paragraph line array (not the per-slice sub-range) onto every slice so the
 *  paint pass can index it by ABSOLUTE line number exactly like a freshly
 *  computed array (drawParagraphLine reads `lines[li]` / `lines.length`). Kept as
 *  a renderer-internal intersection — `LayoutLine` is private to renderer.ts, so
 *  this deliberately does NOT touch the public `PaginatedBodyElement` in types.ts
 *  (which would need a renderer↔types import for the payload type; see the
 *  ColumnGeom note above). `layoutLinesInputs` records the scale-1 inputs the
 *  paginator laid the lines out with; the paint pass reuses them ONLY when its
 *  own scale is 1 and every input matches, so no px field is ever rescaled. */
type PaginatedElementWithLines = PaginatedBodyElement & {
  layoutLines?: LayoutLine[];
  /** The line-layout inputs the paginator used (all in the paginator's scale-1 pt
   *  space). The paint pass reuses the stamped lines ONLY when its own scale is 1
   *  AND every one of these inputs equals the value it would pass to layoutLines
   *  itself — a self-verifying gate that stays correct across single/multi-column,
   *  float wrap, and the numbering firstLineIndent derivation (which differs
   *  between measure and paint for numbered lists, so those simply fail the
   *  `firstIndent` check and recompute). No px field is ever rescaled. */
  layoutLinesInputs?: {
    scale: number;      // always 1 (paginator space)
    paraW: number;      // pt
    firstIndent: number; // pt
    tabOriginPx: number; // pt (== indLeft at scale 1)
    gridDeltaPx: number; // pt
    hasFloats: boolean;  // a float context changes wrap → never reuse across it
    /** The kinsoku rules the paginator laid out with (§17.3.1.16 / §17.15.1.58–.59).
     *  Compared by {@link kinsokuRulesEquivalent} in the paint gate. NOTE this is
     *  usually NOT reference-identical to the paint state's rules: the production
     *  path resolves `resolveKinsokuRules(doc.settings)` once in paginateDocument
     *  and again in renderDocumentToCanvas, and the resolver builds fresh Set
     *  objects per call — so the gate needs value equivalence, not `===` alone. */
    kinsoku: KinsokuRules;
  };
};

/** Value equivalence of two resolved kinsoku rule sets, with a reference fast
 *  path. The reuse gate cannot rely on `===` alone: `resolveKinsokuRules` builds
 *  a FRESH object (fresh Sets) on every call, and the prebuiltPages production
 *  path (DocxDocument.renderPage) resolves it independently in paginateDocument
 *  and in renderDocumentToCanvas — same `doc.settings`, different references.
 *  Both derive from the same immutable settings so they are value-equal there;
 *  this check is pure defense so a genuinely different rule set (which would
 *  change CJK retract decisions in layoutLines) can never reuse stale lines. */
function kinsokuRulesEquivalent(a: KinsokuRules, b: KinsokuRules): boolean {
  if (a === b) return true;
  if (a.enabled !== b.enabled) return false;
  const setEq = (x: Set<number>, y: Set<number>): boolean => {
    if (x.size !== y.size) return false;
    for (const cp of x) if (!y.has(cp)) return false;
    return true;
  };
  return setEq(a.lineStartForbidden, b.lineStartForbidden) && setEq(a.lineEndForbidden, b.lineEndForbidden);
}

/** True when {@link buildSegments} would produce DIFFERENT segments for this
 *  paragraph under the paginator's measure state versus the paint state — i.e.
 *  the segment text depends on per-page render context rather than on the runs
 *  alone. Exactly two sources exist today:
 *    - `field` runs of type page / numPages: {@link resolveFieldText} returns
 *      `state.pageIndex + 1` / `state.totalPages`, and the measure state is
 *      frozen at pageIndex 0 / totalPages 1 (buildMeasureState);
 *    - `noteRef` text runs: the label resolves via `state.noteNumbers` /
 *      `state.currentNoteNumber`, which only the paint state carries
 *      (renderDocumentToCanvas builds the map; the measure state never does).
 *  Such a paragraph must NOT stamp its measured lines: the stamped segments
 *  would carry the measure-time text (a stale page number / note label) and the
 *  paint pass would draw it verbatim. Skipping the stamp keeps those paragraphs
 *  on the recompute path, which resolves fields against the real page context —
 *  the pre-reuse behaviour. Extend this predicate if buildSegments ever gains a
 *  new state-dependent text source. */
function paragraphSegsStateSensitive(para: DocParagraph): boolean {
  for (const run of para.runs) {
    if (run.type === 'field') {
      const ft = (run as unknown as FieldRun).fieldType;
      if (ft === 'page' || ft === 'numPages') return true;
    } else if (run.type === 'text' && (run as unknown as DocxTextRun).noteRef) {
      return true;
    }
  }
  return false;
}

/** Phase 4-1 B2 Stage 1 — master switch for the compute-once line reuse. Always
 *  ON in production; the pixel-identity characterization test flips it OFF to
 *  capture a fresh-recompute reference and assert the reuse path paints an
 *  IDENTICAL stream (see layout-lines-reuse-identity.test.ts). Module-local so it
 *  never leaks onto the public surface. */
let lineReuseEnabled = true;

/** Additional context passed to layoutLines so it can honor floats on the current page. */
interface WrapLayoutCtx {
  startPageY: number;   // absolute canvas Y where the first line should start
  paraX: number;        // absolute canvas X of the paragraph's content left edge
  floats: FloatRect[];  // floats active on the current page
  /** Per-line box-height resolver (line natural ascent+descent → total px box height). */
  lineBoxH: (ascentPx: number, descentPx: number, hasRuby?: boolean, intendedSinglePx?: number) => number;
  /** Hard cap on Y to keep layout from running past the page. */
  pageH: number;
  /** Paragraph-mark em width (px), from paragraphMarkEmPx(para). Shared with
   *  resolveEmptyMarkTop so an empty / anchor-only mark line that has no
   *  trailing-break font of its own falls below a float band on the SAME
   *  threshold the empty-paragraph path uses. */
  markEmPx: number;
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
/**
 * Split a `w:smallCaps` (§17.3.2.33) run into maximal pieces by character class
 * for sizing. The spec reduces "all SMALL LETTER characters ... two points
 * smaller", so ONLY lowercase letters are `reduced`; uppercase letters AND every
 * non-alphabetic character (digits, punctuation) stay at the FULL run size.
 * So "Introduction" → "I" full + "NTRODUCTION" reduced (matching the heading's
 * "1."), and "co2" → "CO" reduced + "2" full. `reduced` flags the small-cap
 * pieces; the caller still uppercases every piece for display.
 *
 * Whitespace carries no glyph, so it EXTENDS the current piece rather than
 * opening a full-size one — otherwise an inter-word space between two small-cap
 * words would fragment into its own segment and corrupt trailing-space collapse
 * / line breaking. A leading run with no lowercase letter defaults to full size.
 */
function splitSmallCapsCase(text: string): { text: string; reduced: boolean }[] {
  const out: { text: string; reduced: boolean }[] = [];
  for (const ch of text) {
    // A lowercase letter: unchanged by toLowerCase AND changed by toUpperCase.
    const isLowerLetter = ch.toLowerCase() === ch && ch.toUpperCase() !== ch;
    const reduced = /\s/.test(ch)
      ? (out[out.length - 1]?.reduced ?? false) // whitespace: keep with current piece
      : isLowerLetter;
    const last = out[out.length - 1];
    if (last && last.reduced === reduced) last.text += ch;
    else out.push({ text: ch, reduced });
  }
  return out.length ? out : [{ text, reduced: false }];
}

function buildSegments(runs: DocRun[], state: RenderState): LayoutSeg[] {
  const segs: LayoutSeg[] = [];
  const pushTextPiece = (
    text: string,
    base: DocxTextRun | FieldRun,
    vertAlign: 'super' | 'sub' | null,
  ) => {
    // §17.3.2.33 small caps are sized per character: lowercase LETTERS render two
    // points smaller, uppercase letters and non-alphabetic characters at the full
    // run size. `reduced` (set per case-piece in the loop below) carries that onto
    // each emitted segment; calcEffectiveFontPx shrinks only the reduced ones.
    // allCaps (§17.3.2.5) and non-caps runs are a single, non-reduced piece.
    let reduced = false;
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

    // ECMA-376 §17.3.2.26 eastAsia axis. Within a non-complex-script slice, CJK
    // code points take the eastAsia face while Latin/digits keep the ascii face
    // (`base.fontFamily`). Only `DocxTextRun` carries the axis; absent (field
    // runs / single-axis parser output) ⇒ fall back to ascii, exactly like
    // `shapeTokenFamily`. Bold/italic/size are NOT axis-specific here — eastAsia
    // shares the Latin (non-cs) toggles, so only the family differs.
    const eaFontFamily = (base as DocxTextRun).fontFamilyEastAsia ?? base.fontFamily;

    // Word classifies European digits in an Arabic/Hebrew complex-script run as
    // AN (§17.3.2.20 w:lang w:bidi): use the bidi language's primary subtag when
    // present, else fall back to the run being rtl-marked.
    const digitsAsAN =
      (forceCs || r.rtl === true) && isRtlBidiLang(r.langBidi, r.rtl === true);

    let firstSeg = true;
    // True while the next emitted segment should be GLUED to the previous one
    // (a small-caps case-piece that continues the same word). Consumed by the
    // first pushSeg of the piece so only that segment carries joinPrev.
    let gluePending = false;
    // Script slot for an emitted segment (§17.3.2.26): 'cs' = complex-script
    // (Arabic/Hebrew/...), 'ea' = East-Asian (CJK → eastAsia face), 'latin' =
    // Latin/digits/neutral (ascii face). Each segment stays SINGLE-FONT — one
    // family for its whole `.text` — so the measure==draw / docGrid char-grid
    // invariant holds and the draw loop needs no per-segment font switching.
    const pushSeg = (text: string, cs: boolean, fontFamily: string | null) => {
      segs.push({
        text,
        bold: cs ? csBold : base.bold,
        italic: cs ? csItalic : base.italic,
        underline: base.underline,
        strikethrough: base.strikethrough,
        fontSize: cs ? csFontSize : base.fontSize,
        color: base.color,
        fontFamily,
        vertAlign,
        measuredWidth: 0,
        smallCaps: reduced,
        joinPrev: gluePending ? true : undefined,
        doubleStrikethrough: base.doubleStrikethrough ?? false,
        highlight: base.highlight ?? null,
        background: base.background ?? null,
        colorAuto: r.colorAuto ?? false,
        border: r.border ?? null,
        ruby: firstSeg ? ruby : undefined,
        revision,
        rtl,
        digitsAsAN: digitsAsAN ? true : undefined,
      });
      firstSeg = false;
      gluePending = false; // glue applies only to a piece's FIRST segment
    };
    const emit = (word: string, slot: 'cs' | 'ea' | 'latin') => {
      const cs = slot === 'cs';
      const fontFamily = slot === 'cs' ? csFontFamily : slot === 'ea' ? eaFontFamily : base.fontFamily;
      // ECMA-376 §17.3.2.26 + §17.3.3.30: a run whose rFonts axis is Symbol or
      // Wingdings stores glyphs as the FONT's own (private) code points — Word
      // commonly in the PUA (U+F020–U+F0FF). Those render as tofu in any
      // fallback face, so normalize each character to its Unicode equivalent
      // (core `symbolTextToUnicodeSegments`, the same table the list marker uses
      // via `symbolFontToUnicode`). The string is split at mapped/unmapped
      // boundaries: a MAPPED run is drawn in a generic fallback (fontFamily=null
      // → sans tail with the dingbat glyphs; keeping the symbol family would let
      // an installed Symbol/Wingdings re-interpret the Unicode code point as the
      // WRONG glyph), while an UNMAPPED run keeps the symbol family so a host
      // that ships Symbol/Wingdings still draws its native glyph. Done once at
      // build time so measure==draw (the seg.text is never transformed later).
      if (isSymbolFontFamily(fontFamily)) {
        for (const part of symbolTextToUnicodeSegments(word, fontFamily)) {
          pushSeg(part.text, cs, part.mapped ? null : fontFamily);
        }
        return;
      }
      pushSeg(word, cs, fontFamily);
    };

    // A non-complex-script slice still mixes scripts at the CJK boundary: emit
    // its maximal CJK runs on the 'ea' (eastAsia) slot and the rest on 'latin'
    // (ascii). Keeps each emitted segment single-font (so a serif ascii digit
    // sits next to a gothic eastAsia title) without changing the cs path.
    const emitNonCs = (slice: string) => {
      for (const part of splitByEastAsia(slice)) emit(part.text, part.ea ? 'ea' : 'latin');
    };

    // Small caps split the run into full-size (uppercase-origin / non-cased) and
    // reduced (lowercase-origin) case-pieces; everything else is one piece. Each
    // piece is still UPPERCASED for display (allCaps or smallCaps), and `reduced`
    // drives its segments' size — see splitSmallCapsCase / calcEffectiveFontPx.
    const casePieces = base.smallCaps
      ? splitSmallCapsCase(text)
      : [{ text, reduced: false }];
    let prevPieceText = '';
    for (const piece of casePieces) {
      reduced = piece.reduced;
      // Glue this piece's FIRST segment to the previous piece when they continue
      // the same word (the previous piece did not end at a space) — so a
      // small-caps word's full-cap initial and reduced remainder stay on one line.
      gluePending = prevPieceText.length > 0 && !/\s$/.test(prevPieceText);
      prevPieceText = piece.text;
      const displayText = (base.allCaps || base.smallCaps) ? piece.text.toUpperCase() : piece.text;
      for (const word of splitTextForLayout(displayText)) {
        if (forceCs) {
          // When the run's digits are AN-classified, split a token into maximal
          // digit-groups and the surrounding separators so the per-line bidi pass
          // (which reorders at SEGMENT granularity) can place the groups in Word's
          // order — e.g. "28-02-2026" → segments [28][-][02][-][2026] reordered to
          // 2026-02-28. Canvas only reorders WITHIN a fillText using EN semantics,
          // so a single-segment date would otherwise stay 28-02-2026.
          if (digitsAsAN) {
            for (const slice of splitDigitGroups(word)) emit(slice, 'cs');
          } else {
            emit(word, 'cs');
          }
        } else {
          // Mixed Arabic+Latin word (no w:rtl / w:cs): split at script boundaries
          // so each side gets its own (cs vs Latin) size and typeface; the non-cs
          // side then sub-splits at CJK boundaries for the eastAsia face.
          for (const slice of splitByComplexScript(word)) {
            if (slice.cs) emit(slice.text, 'cs');
            else emitNonCs(slice.text);
          }
        }
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
          segs.push({ isTab: true, fontSize: t.fontSize, measuredWidth: 0, bold: t.bold, italic: t.italic });
        }
      }
    } else if (run.type === 'image') {
      const img = run as unknown as ImageRun & { type: 'image' };
      segs.push({
        imagePath: img.imagePath,
        mimeType: img.mimeType,
        widthPt: img.widthPt,
        heightPt: img.heightPt,
        anchor: img.anchor ?? false,
        anchorXPt: img.anchorXPt ?? 0,
        anchorYPt: img.anchorYPt ?? 0,
        anchorXFromMargin: img.anchorXFromMargin ?? false,
        anchorYFromPara: img.anchorYFromPara ?? false,
        colorReplaceFrom: img.colorReplaceFrom,
        srcRect: img.srcRect ?? undefined,
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

  // ── UAX#14 LB13 / ECMA-376 §17.15.1.59 (行頭禁則 — line-start-forbidden) ──────
  // A closing / mid-punctuation code point (comma, period, ; : ! ? ) ] } and
  // their CJK forms) carries NO line-break opportunity before it, so it may
  // never BEGIN a line. When such a char OPENS a segment that is glued to the
  // previous text segment — no intervening whitespace, e.g. a comma authored in
  // its own run as in sample-12's "…detection system" | ", metadata" — mark it
  // `joinPrev` so the group machinery in layoutLines keeps it with the preceding
  // word and wraps "system," together instead of orphaning "," at the next
  // line's head.
  //
  // This is a UNIVERSAL Latin/Western rule (UAX#14 LB13), NOT the East-Asian
  // kinsoku feature, so it consults the application's DEFAULT forbidden table
  // UNCONDITIONALLY — independent of the document's §17.3.1.16 `w:kinsoku`
  // toggle and of any custom §17.15.1.59 `w:noLineBreaksBefore` set (which
  // REPLACES the default East-Asian table for a language and so must NOT be able
  // to drop the ASCII non-starters and re-orphan a Latin comma). The document's
  // kinsoku settings still govern the separate per-character CJK retract paths
  // (kinsokuAdjustedSplit / crossRunKinsokuRetract), which read `state.kinsoku`.
  // The ASCII non-starters (!),.:;?]}) live in that default table (core
  // rules.ts), so one membership test covers Latin and (incidentally) CJK forms.
  for (let i = 1; i < segs.length; i++) {
    const cur = segs[i];
    if (!('text' in cur) || cur.joinPrev) continue;
    const firstCp = cur.text.codePointAt(0);
    if (firstCp === undefined || !DEFAULT_KINSOKU_RULES.lineStartForbidden.has(firstCp)) continue;
    const prev = segs[i - 1];
    // Only glue across a boundary that is NOT already a break opportunity: the
    // preceding unit must be text that does not end in whitespace (a trailing
    // space is a legal break, so the mark may legitimately start the line).
    if (!('text' in prev) || /\s$/.test(prev.text)) continue;
    cur.joinPrev = true;
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

/** Returns true when any code point of `text` permits a line break between
 *  adjacent characters (CJK / ideographic). The canonical ranges live in core's
 *  {@link isCjkBreakChar} (single source of truth across all renderers). */
function hasCJKBreakOpportunity(text: string): boolean {
  for (let i = 0; i < text.length; ) {
    const cp = text.codePointAt(i)!;
    if (isCjkBreakChar(cp)) return true;
    i += cp > 0xffff ? 2 : 1;
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
  // ECMA-376 §17.6.5 character-grid delta (px per EA glyph, 0 when inactive).
  // The fit must compare CELL widths so the grid's char count lands per line —
  // the same `gridWidth` the line box / draw uses, keeping the split consistent.
  gridDeltaPx = 0,
): string {
  const chars = [...text]; // spread handles surrogate pairs
  let lo = 0, hi = chars.length;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    const prefix = chars.slice(0, mid).join('');
    if (gridWidth(ctx.measureText(prefix).width, prefix, gridDeltaPx) <= maxWidth) lo = mid;
    else hi = mid - 1;
  }
  return chars.slice(0, lo).join('');
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
 * Split a (non-complex-script) string into maximal runs that are uniformly
 * East-Asian (CJK) or not, per the §17.3.2.26 ascii/eastAsia axis split. Returns
 * `[{text, ea}]` in logical order. CJK classification uses the canonical
 * {@link isCjkBreakChar} from `@silurus/ooxml-core` — the SAME predicate the
 * shape-text path ({@link shapeTokenFamily}) and the body wrap/justify paths
 * use, so the eastAsia face is picked consistently across renderers with no name
 * heuristics. Each returned slice stays single-font when emitted, preserving the
 * measure==draw / docGrid char-grid invariant.
 *
 * Boundary rule: classification is purely per code point (every CJK code point
 * opens/continues an `ea` run; every other code point a `latin` run). This is
 * intentionally simpler than {@link splitByComplexScript}'s neutral-attachment —
 * a digit between two ideographs is Latin/ascii either way (Word renders ASCII
 * digits with the ascii face), and a single fillText anchors to the cumulative
 * whole-string advance, so the visible spacing is unchanged.
 *
 * NOTE: this split decides the FONT slot only. Whether a resulting segment is
 * snapped to the §17.6.5 character grid is decided SEPARATELY by the grid's own
 * `EAST_ASIAN_RE` purity test (see `gridSegDeltaPx`/`eaGlyphCount`), not by the
 * `ea` flag here. The two CJK predicates classify slightly different code-point
 * sets; correctness of the grid total relies on `eaGlyphCount` being additive
 * over this partition (covered by docgrid-char.test.ts's mixed-token case). Keep
 * them independent — do not "unify" the predicates without re-checking that test.
 */
function splitByEastAsia(text: string): { text: string; ea: boolean }[] {
  const out: { text: string; ea: boolean }[] = [];
  let curEa: boolean | null = null;
  let buf = '';
  for (const ch of text) {
    const cp = ch.codePointAt(0) as number;
    const ea = isCjkBreakChar(cp);
    if (curEa === null || ea === curEa) {
      curEa = ea;
      buf += ch;
    } else {
      out.push({ text: buf, ea: curEa });
      curEa = ea;
      buf = ch;
    }
  }
  if (buf.length > 0) out.push({ text: buf, ea: curEa ?? false });
  return out;
}

/**
 * Split a token into maximal runs of European digits (U+0030–0039) versus the
 * separators between them, so a date in an AN-classified Arabic run can be
 * reordered group-by-group by the per-line bidi pass (which works at segment
 * granularity). "28-02-2026" → ["28","-","02","-","2026"], which the RTL reorder
 * then lays out right-to-left as Word does.
 *
 * EXCEPTION — ECMA-376 relies on UAX#9 W4: a SINGLE common separator (CS) sitting
 * between two numbers of the same type joins them into ONE number. So a decimal /
 * thousands / time separator (`.`, `,`, `:`, `/`, NBSP) flanked by European
 * digits on BOTH sides stays inside the digit group: "1234.56", "1,234.56" and
 * "12:34" are one left-to-right number, not three reorderable pieces. (A European
 * separator like `-` is ES, NOT CS, and W4's ES clause is EN-only — these run
 * digits are AN — so a hyphen still splits, preserving the date case.) Splitting
 * a decimal sent "1234.56" through the RTL segment reorder and drew it "56.1234".
 */
export function splitDigitGroups(text: string): string[] {
  const isEuDigit = (c: number) => c >= 0x30 && c <= 0x39;
  // UAX#9 Common Separator (CS) subset that can join two adjacent numbers (W4).
  // The last char is NBSP (U+00A0, e.g. a French thousands separator), itself CS;
  // a plain space is WS and never reaches here (splitTextForLayout breaks on it).
  const isJoiningCS = (ch: string) =>
    ch === '.' || ch === ',' || ch === ':' || ch === '/' || ch === ' ';
  const out: string[] = [];
  let buf = '';
  let bufDigit: boolean | null = null;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    let isDigit = isEuDigit(ch.charCodeAt(0));
    // W4: a single CS between two European digits is part of the number — keep it
    // in the current digit group so the whole number stays one (LTR) segment.
    if (!isDigit && bufDigit === true && isJoiningCS(ch) && isEuDigit(text.charCodeAt(i + 1))) {
      isDigit = true;
    }
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
  // ECMA-376 §17.6.5 docGrid CHARACTER grid: per-EA-glyph cell delta in px (0
  // when inactive). Folded into every advance via `gridWidth` so line breaking
  // packs the grid's char count per line; the draw paths add the SAME delta.
  gridDeltaPx = 0,
  // ECMA-376 §17.15.1.25 — automatic tab-stop interval (pt). The automatic-stop
  // grid (`nextTabStop`) multiplies this by `scale`; defaults to the spec absent
  // value (720 twips = 36pt) for callers without document settings.
  defaultTabPt: number = DEFAULT_TAB_PT,
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

  // Compute wrap constraints for a new line about to start. Mutates
  // lineXOffset/lineMaxWidth/currentLineTopY. `minWidth` is the smallest width
  // the upcoming line must have to be placeable here (the width of its first
  // atomic token, or the paragraph-mark em for an empty line); a free gap
  // narrower than this is treated as unusable and the line is sent below the
  // intervening float(s) — the ECMA-376 wrap rule that text which cannot fit
  // beside a floating object flows past it.
  const startLine = (minWidth: number = 0): void => {
    lineXOffset = 0;
    lineMaxWidth = maxWidth;
    if (!wrapCtx) return;
    // Small fixed probe height for float intersection (matches the historical
    // wrap behaviour for the topAndBottom skip and horizontal-gap scan).
    const probeH = 10 * scale;
    const win = resolveLineFloatWindow(
      currentLineTopY, minWidth, probeH, wrapCtx.paraX, maxWidth, wrapCtx.floats,
    );
    currentLineTopY = win.topY;
    lineXOffset = win.xOffset;
    lineMaxWidth = win.maxWidth;
  };

  const availW = () => lineMaxWidth - (isFirst ? firstIndent : 0);

  let lineHasRuby = false;
  const flush = (forceHeight?: number, brTerminated = false) => {
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
      endsWithBreak: brTerminated,
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
    startLine(requiredLineWidth());
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
    if (!('isTab' in s) && !('imagePath' in s) && !('mathNodes' in s)) {
      const ts = s as LayoutTextSeg;
      if (ts.ruby) lineHasRuby = true;
      // Intended single-line height for fonts whose substituted Canvas metrics
      // understate Word's line spacing (font-metrics.ts). 0 for untabled fonts.
      // Small caps (non-super/sub) keep the FULL run size here so the line box
      // follows the run size, not the 2pt-reduced glyphs (§17.3.2.33).
      const intendedEm = ts.smallCaps && !ts.vertAlign ? ts.fontSize * scale : effectiveFontPx(ts);
      const intended = intendedSingleLinePx(ts.fontFamily, intendedEm);
      if (intended > lineIntendedSingle) lineIntendedSingle = intended;
    }
  };

  const effectiveFontPx = (s: LayoutTextSeg): number => calcEffectiveFontPx(s, scale);

  // Measure-loop font guard: line wrapping calls measureText / strAdvance many
  // times in a row for the SAME segment (fit search, split prefixes/tails), so
  // the built font string is usually identical to the previous one. Skip the
  // redundant `ctx.font =` in that case. This tracker is written by EVERY font
  // assignment on the measure path (both helpers below route through it), so it
  // always reflects the context's current measure font — no stale skip. The
  // draw-path `ctx.font =` sites are separate and left untouched. `buildFont` is
  // now cheap (normalizeFontFamily is memoized per-doc), so this only elides the
  // setter call itself.
  let lastMeasureFont: string | null = null;
  const setMeasureFont = (font: string): void => {
    if (font !== lastMeasureFont) {
      ctx.font = font;
      lastMeasureFont = font;
    }
  };

  const measureText = (s: LayoutTextSeg): TextMetrics => {
    setMeasureFont(buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses));
    return ctx.measureText(s.text);
  };

  // The segment's laid-out ADVANCE (= its measuredWidth): natural width plus the
  // character-grid delta. This is the SINGLE source of truth shared with the
  // draw paths (gridWidth) — every line-break / fit / tab measurement uses it so
  // line wrapping packs the grid's char count and the box matches what is drawn.
  const segAdvance = (s: LayoutTextSeg): number =>
    gridWidth(measureText(s).width, s.text, gridDeltaPx);
  // Grid advance of an arbitrary string under a segment's font (for split
  // prefixes/tails). Selects the font, then applies the same gridWidth model.
  const strAdvance = (s: LayoutTextSeg, text: string): number => {
    setMeasureFont(buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses));
    return gridWidth(ctx.measureText(text).width, text, gridDeltaPx);
  };

  // Width of a queued segment, for right/center tab look-ahead.
  const tabFollowWidth = (q: LayoutSeg): number => {
    if ('isTab' in q) return q.measuredWidth || 0;
    if ('imagePath' in q) return q.widthPt * scale;
    if ('mathNodes' in q) return q.measuredWidth || 0;
    if ('lineBreak' in q) return 0;
    return segAdvance(q);
  };

  // Use an explicit queue so CJK split-tails can be re-queued
  const queue: LayoutSeg[] = [...segs];

  // Smallest width the NEXT line must have to be placeable beside a float: the
  // width of its first atomic token. For text we measure the first wrap unit
  // (CJK: one grapheme — kinsoku may force more, but one char is the floor;
  // Latin: up to the first space). For an image/math the whole object. An empty
  // line (no remaining content) still reserves the paragraph-mark em so a
  // sliver gap between full-width floats does not "fit" it. Used by startLine to
  // decide whether to wrap in a gap or send the line below the floats.
  const requiredLineWidth = (): number => {
    // First inline token that actually occupies width on the line. Anchor-image
    // segments are floats (drawn separately, measuredWidth 0) and lineBreaks
    // carry no width, so skip them — otherwise the float's own width would be
    // mistaken for the line's required width and wrongly push every wrap line
    // below the float.
    const q = queue.find((s) => !('lineBreak' in s) && !('imagePath' in s && s.anchor));
    if (!q) {
      // Empty/paragraph-mark line: reserve one em so a sub-glyph gap is rejected
      // and the mark line drops below the floats. A trailing `<w:br/>` carries
      // its own (line-local) font size; otherwise use the shared paragraph-mark
      // em (paragraphMarkEmPx, threaded via wrapCtx) so this fallback agrees with
      // resolveEmptyMarkTop. Outside a float context the result is unused
      // (startLine no-ops), so fall back to the legacy estimate.
      if (trailingBreakFontSize !== null) return trailingBreakFontSize * scale;
      if (wrapCtx) return wrapCtx.markEmPx;
      return (segs[0] && 'fontSize' in segs[0] ? segs[0].fontSize : 10) * scale;
    }
    if ('isTab' in q) return q.measuredWidth || 0;
    if ('imagePath' in q) return q.widthPt * scale;
    if ('mathNodes' in q) return q.measuredWidth || 0;
    const ts = q as LayoutTextSeg;
    // First wrap unit: leading run up to the first space (Latin word) or the
    // first character (CJK / no space). Whichever is shorter bounds the floor.
    const sp = ts.text.indexOf(' ');
    const head = sp > 0 ? ts.text.slice(0, sp) : ts.text;
    const firstChar = [...head][0] ?? '';
    const probe = { ...ts, text: firstChar };
    return segAdvance(probe);
  };

  // A `<w:br/>` always starts a new line (§17.3.3.1) — when it is the LAST
  // content of the paragraph that new line is an EMPTY line that still
  // occupies one line height (Word reserves it; visible e.g. as extra table
  // row height). Track the trailing break so it can be flushed after the loop.
  let trailingBreakFontSize: number | null = null;

  // Establish the first line's wrap window now that the content queue exists.
  startLine(requiredLineWidth());

  while (queue.length > 0) {
    const seg = queue.shift()!;

    // ── Line-break sentinel ──────────────────────────────
    if ('lineBreak' in seg) {
      // The line being flushed ends at a MANUAL break (§17.3.3.1) — mark it so a
      // justified paragraph left-aligns it like its final line (§17.18.44).
      flush(seg.fontSize, true);
      trailingBreakFontSize = seg.fontSize;
      continue;
    }
    trailingBreakFontSize = null;

    // ── Tab segment ──────────────────────────────────────
    if ('isTab' in seg) {
      // Absolute position on the line measured from paraX (line origin for continuation lines)
      const absFromParaX = currentWidth + (isFirst ? firstIndent : 0);
      // ECMA-376 §17.3.1.37 / §17.15.1.25 — resolve the next stop in TEXT-MARGIN
      // coordinates (the same origin as custom stops): the current pen position is
      // `absFromParaX + tabOriginPx`, custom stops are `pos * scale`, and the
      // automatic grid interval is `defaultTabPt * scale`. Mixing paraX and margin
      // coordinates is what diverged leading-tab rows from labeled ones; computing
      // both in margin space and converting back keeps them aligned.
      const curMarginPx = absFromParaX + tabOriginPx;
      const customStopsPx = tabStops.map((t) => ({ pos: t.pos * scale, alignment: t.alignment, leader: t.leader }));
      const stop = nextTabStop(curMarginPx, customStopsPx, defaultTabPt * scale);
      // Convert the chosen margin-space stop back to paraX-relative px.
      const stopParaX = stop ? stop.pos - tabOriginPx : absFromParaX;
      // Right/center/decimal tab: place the tab + its trailing content (up to the next
      // tab / line end) so the content ends at / centers on the stop, and commit that
      // content directly so the normal wrap check doesn't push it past the stop
      // (ECMA-376 §17.3.1.37). This is what makes TOC "heading …… page" lines work.
      // Automatic stops returned by nextTabStop are left-aligned, so they fall
      // through to the left-tab path below.
      if (stop && stop.alignment !== 'left' && stop.alignment !== 'bar' && stop.alignment !== 'clear') {
        const stopX = stopParaX;
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
          if ('imagePath' in q) {
            const w = q.widthPt * scale;
            q.measuredWidth = w;
            addToLine(q, w, q.heightPt, q.heightPt * scale, 0);
          } else if ('mathNodes' in q) {
            addToLine(q, q.measuredWidth || 0, q.fontSize, q.mathAscent || 0, q.mathDescent || 0);
          } else {
            const m = measureText(q);
            const w = gridWidth(m.width, q.text, gridDeltaPx);
            q.measuredWidth = w;
            const asc = m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent ?? q.fontSize * scale * 0.8;
            const desc = m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent ?? q.fontSize * scale * 0.2;
            addToLine(q, w, q.fontSize, asc, desc);
          }
        }
        continue;
      }

      // Left-aligned tab (custom 'left'/'bar'/'clear' or an automatic stop): the
      // pen moves to the stop's paraX. nextTabStop already applied the §17.15.1.25
      // "after all custom stops" automatic grid, so there is no separate fallback.
      let tabWidth = stopParaX - absFromParaX;
      if (stop) seg.leader = stop.leader;
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
    if ('imagePath' in seg) {
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
    // Advance = natural width + character-grid delta (the SINGLE model shared
    // with the draw paths; 0 unless an active grid AND a pure-EA segment).
    const w = gridWidth(m.width, s.text, gridDeltaPx);
    // Line-height tracks the un-scaled pt font so super/sub don't shrink the line.
    const h = s.fontSize;
    // Prefer font-metric ascent/descent (stable per font+size) so baselines and
    // line boxes do not jitter based on the specific characters on each line.
    // When the document font is substituted by one with different vertical
    // metrics, rescale to the document font's design line box so the line
    // height (and thus baseline centering, row auto-heights, cell vAlign
    // §17.4.84, and pagination) match Word — see font-metrics.ts.
    // Fallback box at the run's full size; correction at the effective size
    // (smallCaps/vertAlign shrink it). See correctedLineMetrics.
    // §17.3.2.33: a small-caps run's LINE BOX follows the FULL run size, not the
    // 2pt-reduced glyph size — so a wrapped continuation line carrying only reduced
    // (all-lowercase) small caps is not short. Measure the box metrics at the full
    // size (the advance `w` above still uses the reduced glyphs). Super/subscript
    // is excluded: it intentionally shrinks its line contribution.
    const fullPx = s.fontSize * scale;
    let metricM = m;
    let metricEmPx = effectiveFontPx(s);
    if (s.smallCaps && !s.vertAlign && metricEmPx !== fullPx) {
      const prevFont = ctx.font;
      ctx.font = buildFont(s.bold, s.italic, fullPx, s.fontFamily, fontFamilyClasses);
      metricM = ctx.measureText(s.text || 'X');
      ctx.font = prevFont;
      metricEmPx = fullPx;
    }
    const corrected = correctedLineMetrics(metricM, s.fontFamily, fullPx, metricEmPx);
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
    // Subtract the GRID width of the trimmed text (not the natural width) so the
    // grid delta on EA glyphs cancels and trailingSpaceW is the bare space
    // advance — keeping `w` and `wForFit` on the one advance model.
    const trailingSpaceW = s.text.endsWith(' ')
      ? w - gridWidth(ctx.measureText(trimmed).width, trimmed, gridDeltaPx)
      : 0;
    const wForFit = w - trailingSpaceW;
    const shrinkBudget = lineTotalTrailingW * SPACE_SHRINK_RATIO;

    // Atomic glued group: when THIS segment starts a glued group (its followers
    // in the queue are `joinPrev` pieces — small-caps case-pieces of the SAME
    // word like "I" then "NTRODUCTION", or a UAX#14 LB13 non-starter authored in
    // its own run like a trailing "," / "。"), the per-segment wrap below would
    // let the group split across lines. Pre-measure it and, if it does not fit on
    // the current (non-empty) line, flush so it starts fresh.
    //
    // ONLY when the lead segment is NOT itself CJK-breakable. A glued group whose
    // lead is a CJK run (e.g. "…通過する" + "。") is NOT atomic: the run splits at
    // an inter-CJK boundary and the trailing non-starter stays on its LAST piece
    // (§17.3.1.16 kinsoku keeps it off the next line's head when enabled — the
    // default; with kinsoku off it may lead the line, as it did before PR #602).
    // Pre-flushing the whole run instead leaves the prior line far short, which a
    // `both` line then stretches wide (sample-9). `joinPrev` stays a pure "this is
    // a non-starter" marker; the atomic-vs-breakable decision lives HERE. A
    // non-breakable Latin / small-caps lead is genuinely atomic, so the pre-flush
    // (and the over-long-word char-break path below) still applies there.
    if (
      !s.joinPrev &&
      currentLine.length > 0 &&
      (queue[0] as LayoutTextSeg | undefined)?.joinPrev &&
      !hasCJKBreakOpportunity(s.text)
    ) {
      let groupW = w;
      let groupTrail = trailingSpaceW;
      for (let k = 0; k < queue.length && (queue[k] as LayoutTextSeg).joinPrev; k++) {
        const f = queue[k] as LayoutTextSeg;
        // A CJK-BREAKABLE follower (e.g. "Roman" + "、あるいは…用いる。") is NOT
        // atomic: only its LEADING run of line-start-forbidden chars would orphan
        // at a line head (UAX#14 LB13 / §17.3.1.16); the rest splits at an
        // inter-CJK boundary and wraps on its own. So glue only that prefix's
        // advance to the lead and STOP summing here — mirror of the CJK-lead
        // direction handled by the `!hasCJKBreakOpportunity(s.text)` gate above
        // (sample-9 fb836d6). Summing the whole breakable run instead would
        // pre-flush "Roman" down alone, leaving a `both` line stretched sparse
        // (sample-16). A Latin / small-caps follower (no CJK break opportunity —
        // the "I" + "NTRODUCTION" case) stays fully atomic: keep full-add.
        if (hasCJKBreakOpportunity(f.text)) {
          const chars = [...f.text];
          let p = 0;
          while (p < chars.length && DEFAULT_KINSOKU_RULES.lineStartForbidden.has(chars[p].codePointAt(0)!)) p++;
          if (p < chars.length) {
            // Breakable rest exists past the leading non-starters: glue only the
            // prefix (it may be empty — then "Roman" is effectively unglued and
            // wraps on its own) and end the atomic group here.
            groupW += strAdvance(f, chars.slice(0, p).join(''));
            groupTrail = 0;
            break;
          }
          // Entirely non-starters (no breakable rest): fall through to full-add.
        }
        const fw = segAdvance(f);
        groupW += fw;
        const ft = f.text.replace(/ +$/, '');
        groupTrail = f.text.endsWith(' ') ? fw - gridWidth(ctx.measureText(ft).width, ft, gridDeltaPx) : 0;
      }
      if (currentWidth + (groupW - groupTrail) > availW() + shrinkBudget) flush();
    }

    if (currentWidth + wForFit <= availW() + shrinkBudget) {
      // Fits on current line as-is
      s.measuredWidth = w;
      addToLine(s, w, h, asc, desc, trailingSpaceW);
    } else if (hasCJKBreakOpportunity(s.text)) {
      // CJK overflow: split at the maximum prefix that fits, re-queue the tail.
      // (pptx's analogous CJK fit is cjk-wrap.ts `fitCjkLine`, kept intentionally
      //  separate: it sums per-char advances, whereas this path uses substring
      //  binary-search + the cross-run 追い出し below. Don't naively unify them.)
      const available = availW() - currentWidth;
      ctx.font = buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses);
      const rawPrefix = available > 0 ? fitCJKPrefix(ctx, s.text, available, gridDeltaPx) : '';
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
        // Grid advance for the head piece — the same model as the line box / draw.
        const pw = strAdvance(s, prefix);
        const headSeg: LayoutTextSeg = { ...s, text: prefix, measuredWidth: pw };
        addToLine(headSeg, pw, h, asc, desc);
        const tail = s.text.slice(prefix.length);
        if (tail) queue.unshift({ ...s, text: tail, measuredWidth: 0 });
      } else if (currentLine.length > 0) {
        // No prefix of `s` fits. If `s` would lead the next line with a 行頭禁則
        // char, kinsokuAdjustedSplit can't fix it from within `s` (the offending
        // char is its first); pull trailing graphemes of the current line's last
        // text segment down so they lead the next line ahead of `s` — cross-run
        // 追い出し (§17.3.1.16). See crossRunKinsokuRetract for the bounded,
        // re-validating, whitespace-guarded retraction count.
        let retracted: LayoutTextSeg | null = null;
        const sFirstCp = s.text.codePointAt(0);
        const lastSeg = currentLine[currentLine.length - 1];
        if (sFirstCp !== undefined && kinsoku.lineStartForbidden.has(sFirstCp) && 'text' in lastSeg) {
          const lastText = lastSeg as LayoutTextSeg;
          const chars = [...lastText.text];
          const minKeep = currentLine.length > 1 ? 0 : 1;
          const k = crossRunKinsokuRetract(chars, kinsoku, minKeep);
          if (k > 0) {
            const headText = chars.slice(0, chars.length - k).join('');
            const tailText = chars.slice(chars.length - k).join('');
            retracted = { ...lastText, text: tailText, measuredWidth: strAdvance(lastText, tailText) };
            if (headText) {
              const headW = strAdvance(lastText, headText);
              currentWidth -= lastText.measuredWidth - headW;
              currentLine[currentLine.length - 1] = { ...lastText, text: headText, measuredWidth: headW };
            } else {
              // Whole last segment moves down. Line metrics (ascent/descent) are
              // not recomputed; the retracted graphemes share the line's font in
              // practice, so the flushed box height is unaffected.
              currentWidth -= lastText.measuredWidth;
              currentLine.pop();
            }
          }
        }
        flush();
        queue.unshift(s);
        if (retracted) queue.unshift(retracted);
      } else {
        // Empty line and not even one char fits — force-fit one char to guarantee progress
        const firstChar = [...s.text][0] ?? '';
        if (firstChar) {
          const fw = strAdvance(s, firstChar);
          const headSeg: LayoutTextSeg = { ...s, text: firstChar, measuredWidth: fw };
          addToLine(headSeg, fw, h, asc, desc);
          const tail = s.text.slice(firstChar.length);
          if (tail) queue.unshift({ ...s, text: tail, measuredWidth: 0 });
        }
      }
    } else if (currentLine.length === 0) {
      // A single non-CJK word wider than the FULL line width — e.g. a long URL in
      // a narrow newspaper column. ECMA-376 prescribes no line-break algorithm;
      // Word breaks such an over-long word at the character level (overflow-wrap)
      // so it stays inside the column instead of bleeding past the right margin /
      // into the next column. Fit the widest character prefix (≥1 char so the
      // split always advances), draw it, and re-queue the remainder. Segments are
      // already one space-delimited word (splitTextForLayout), so this never
      // breaks where a space could have wrapped.
      const available = availW();
      ctx.font = buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses);
      const allChars = [...s.text];
      let split = available > 0 ? [...fitCJKPrefix(ctx, s.text, available, gridDeltaPx)].length : 0;
      if (split < 1) split = 1;
      if (split >= allChars.length) {
        // The visible glyphs actually fit (only a trailing space pushed it over the
        // fit test) — place the word whole.
        s.measuredWidth = w;
        addToLine(s, w, h, asc, desc);
      } else {
        const prefix = allChars.slice(0, split).join('');
        const pw = strAdvance(s, prefix);
        addToLine({ ...s, text: prefix, measuredWidth: pw }, pw, h, asc, desc);
        queue.unshift({ ...s, text: allChars.slice(split).join(''), measuredWidth: 0 });
      }
    } else {
      // Latin word does not fit on the current (non-empty) line: move it to a fresh
      // line and re-process. There it either fits, or — when it is wider than the
      // whole column — the empty-line branch above breaks it at the character level
      // (overflow-wrap). Re-queueing rather than force-adding is what lets that
      // over-long-word path run instead of letting the word spill the column.
      flush();
      queue.unshift(s);
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
  bold?: boolean,
  italic?: boolean,
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
  // ECMA-376 §17.3.1.37: the leader fill takes the formatting of the tab's run,
  // so a bold/italic TOC entry draws a bold/italic dot leader.
  const style = `${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}`;
  ctx.font = `${style}${fontPx}px serif`;
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
  images: Map<string, DecodedImage>,
): void {
  // Anchor images are skipped during layout (measuredWidth=0, not added to line.segments)
  // and are drawn later by renderAnchorImages — so this function only handles inline images.
  if (seg.anchor) return;
  const bmp = images.get(imageKey(seg.imagePath, seg.colorReplaceFrom));
  if (!bmp) return;
  const w = seg.widthPt * scale;
  const h = seg.heightPt * scale;
  drawImageCropped(ctx, bmp, seg.srcRect, x, baseline - h, w, h);
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
  // Front floats (behindDoc="0"): defer to the page's top layer so a later inline
  // image cannot overpaint them (§20.4.2.10). Capture the current column band so
  // the replayed draw resolves a column-relative anchor against the right widths.
  if (state.deferFront) {
    const cx = state.contentX;
    const cw = state.contentW;
    state.deferFront.push(() => {
      const sx = state.contentX;
      const sw = state.contentW;
      const sd = state.deferFront;
      state.contentX = cx;
      state.contentW = cw;
      state.deferFront = null; // draw in place this time
      renderAnchorImages(para, state, paragraphTopPx, 'front');
      state.contentX = sx;
      state.contentW = sw;
      state.deferFront = sd;
    });
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
    const bmp = state.images.get(imageKey(img.imagePath, img.colorReplaceFrom));
    if (!bmp) continue;

    // wrapNone images anchor against the paragraph's pre-spaceBefore top
    // (paragraphTopPx). Shared box resolution with the float path. By design the
    // box-resolution is symmetric but the overlap handling is NOT: wrap floats
    // (registerAnchorFloats) build an exclusion rect and run resolveFloatOverlap,
    // whereas wrapNone images carry no exclusion rect — they are positioned
    // directly in the paragraph flow (ECMA-376 wrapNone, §20.4.2.x: the object
    // does not displace text and is not displaced by other floats), so dist* is
    // unused here.
    const { x: pageX, y: pageY, w, h } = resolveAnchorBox(img, state, paragraphTopPx);
    drawImageCropped(state.ctx, bmp, img.srcRect ?? undefined, pageX, pageY, w, h);
  }
}

// Anchor placement geometry (xContainer / yContainer / resolveAnchorX /
// resolveAnchorY, ECMA-376 §20.4.3.x) lives in anchor-geometry.ts and is
// imported above; the shape/image render paths below consume it.

/** Convert a parsed docx LineEnd into core's ArrowEnd. Returns undefined when
 *  absent so the Stroke field stays unset. */
function lineEndToArrowEnd(
  end: ShapeRun['headEnd'],
): ArrowEnd | undefined {
  if (!end) return undefined;
  return { type: end.type, w: end.w, len: end.len };
}

/**
 * Resolve an anchored shape's page-space bounding box {x,y,w,h} (px). Shared by
 * renderAnchorShape (where the shape is drawn) and registerAnchorFloats (where
 * its float-exclusion rect is built), so the exclusion band matches the paint
 * box exactly — see root CLAUDE.md (no duplicated geometry).
 *
 * Mirrors the renderer's sizing: sizeRelH/sizeRelV (ECMA-376 §20.4.2.18)
 * override the static extent, and a wgp child scales by the group ratio with its
 * within-group offset scaled in step; resolveAnchorX/Y then place the box. `w`/`h`
 * may be 0/negative for degenerate line presets — the caller decides how to
 * treat those (renderAnchorShape draws a line; a wrap-shape with no area
 * registers no float).
 */
function resolveShapeBox(
  shape: ShapeRun,
  state: RenderState,
  paragraphTopPx: number,
): { x: number; y: number; w: number; h: number } {
  const { scale } = state;
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
  const x = resolveAnchorX(
    shape.anchorXAlign, shape.anchorXFromMargin, offsetXPt, w, state,
    shape.anchorXRelativeFrom, shape.pctPosH, alignWidthPt,
  );
  const y = resolveAnchorY(
    shape.anchorYAlign, shape.anchorYFromPara, offsetYPt, h, paragraphTopPx, state,
    shape.anchorYRelativeFrom, shape.pctPosV, alignHeightPt,
  );
  return { x, y, w, h };
}

function renderAnchorShape(shape: ShapeRun, state: RenderState, paragraphTopPx: number): void {
  const { ctx, scale } = state;
  const { x, y, w, h } = resolveShapeBox(shape, state, paragraphTopPx);
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
  // Straight / bent connectors whose leader we re-stroke retracted from filled
  // line-end decorations (so the line stops at the arrow base). Curved
  // connectors are excluded — their Bézier leader can't be retracted from a
  // polyline vertex without straightening it, so they keep the preset leader.
  const isRetractableLeader =
    preset === 'line' ||
    preset.startsWith('straightconnector') ||
    preset.startsWith('bentconnector');
  if (w < 0 || h < 0) return;
  if (isLineGeom ? w === 0 && h === 0 : w === 0 || h === 0) return;

  const rot = shape.rotation ?? 0;
  const flipH = shape.flipH ?? false;
  const flipV = shape.flipV ?? false;
  ctx.save();
  // §20.1.7.6 — rotate then mirror about the shape centre, matching the pptx
  // renderer. Applying flip via the canvas transform keeps the body path, the
  // connector arrow-head position, and its direction consistent (a flipped
  // connector swaps which tip carries the head/tail end).
  if (rot !== 0 || flipH || flipV) {
    ctx.translate(x + w / 2, y + h / 2);
    if (rot !== 0) ctx.rotate((rot * Math.PI) / 180);
    if (flipH) ctx.scale(-1, 1);
    if (flipV) ctx.scale(1, -1);
    ctx.translate(-(x + w / 2), -(y + h / 2));
  }
  // Dispatch to the shared spec-driven preset engine when the geometry is a
  // known <a:prstGeom> preset, mirroring the pptx renderer. `arc` (ECMA-376
  // §20.1.10.56 ST_ShapeType "arc") goes through the engine too: its
  // presetShapeDefinitions geometry is two <path>s — path 0 (stroke="false")
  // fills the pie wedge (arc + lnTo centre + close) and path 1 (fill="none")
  // strokes the open arc edge — which the engine honours per-path. The legacy
  // buildShapePath fallback could only draw the open arc, so filling it
  // auto-closed into a chord; arc was excluded here to dodge that, but the
  // engine renders it faithfully now. custGeom (no presetGeometry, subpaths
  // only) still falls through to buildCustomPath.
  const geom = shape.presetGeometry?.toLowerCase() ?? '';
  const usePresetEngine =
    !!shape.presetGeometry && hasPreset(geom);

  const adj = shape.adjValues ?? [];
  const fillStyle = resolveFill(shape.fill, ctx as CanvasRenderingContext2D, x, y, w, h);

  // Build a core Stroke so dash / line-end handling matches the pptx path.
  // `width` is in pt and `scale` is px/pt, so `width * scale` is px — the
  // same convention core's applyStroke / drawArrowHead expect.
  const coreStroke: Stroke | null =
    shape.stroke && (shape.strokeWidth ?? 0) > 0
      ? {
          color: shape.stroke,
          width: shape.strokeWidth ?? 0,
          dashStyle: shape.strokeDash ?? undefined,
          headEnd: lineEndToArrowEnd(shape.headEnd),
          tailEnd: lineEndToArrowEnd(shape.tailEnd),
        }
      : null;
  const strokeCb = coreStroke
    ? () => {
        applyStroke(ctx as CanvasRenderingContext2D, coreStroke, scale);
        ctx.stroke();
      }
    : null;

  if (usePresetEngine) {
    renderPresetShape(
      ctx as CanvasRenderingContext2D,
      geom, x, y, w, h,
      [
        adj[0] ?? null, adj[1] ?? null, adj[2] ?? null, adj[3] ?? null,
        adj[4] ?? null, adj[5] ?? null, adj[6] ?? null, adj[7] ?? null,
      ],
      fillStyle, strokeCb,
      // docx shapes carry no shadow state, so the clear-shadow hook is a no-op.
      () => {},
      // A retractable connector leader is re-stroked retracted below; suppress
      // the preset engine's full-length leader stroke to avoid a double line /
      // a cap poking through the arrow tip.
      isRetractableLeader ? { skipTrailingStroke: true } : undefined,
    );
  } else {
    ctx.beginPath();
    if (shape.presetGeometry) {
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
    if (fillStyle) {
      ctx.fillStyle = fillStyle;
      ctx.fill();
    }
    if (strokeCb) strokeCb();
  }

  // Line-end decorations (ECMA-376 §20.1.8.3). Only connector/line presets
  // expose head/tail tips with a well-defined tangent; for those we place the
  // arrow heads at the path endpoints with the path's outgoing direction. The
  // preset engine does not draw connector arrow heads, so this runs whether or
  // not the body went through it. Gate on `isLineGeom` (mirroring the pptx
  // renderer's CONNECTOR_GEOMS set): getConnectorAnchors resolves path[0] of
  // *any* preset, so a filled shape carrying an <a:ln> head/tail end would
  // otherwise get spurious arrow heads at its first subpath's endpoints.
  if (coreStroke && (coreStroke.headEnd || coreStroke.tailEnd) && isLineGeom) {
    const anchors = getConnectorAnchors(
      preset, x, y, w, h,
      shape.adjValues ?? [],
    );
    if (anchors) {
      ctx.setLineDash([]);
      // Re-stroke the leader retracted from any filled decoration so the line
      // stops at the arrow base instead of poking through its tip (Word /
      // PowerPoint behaviour). Straight/bent only; curved keeps its preset leader.
      if (isRetractableLeader && anchors.vertices.length >= 2) {
        const pts = anchors.vertices.map((v) => ({ x: v.x, y: v.y }));
        if (coreStroke.tailEnd) {
          const r = lineEndRetract(coreStroke.tailEnd, coreStroke, scale);
          pts[pts.length - 1] = retractLineEndpoint(pts[pts.length - 1], pts[pts.length - 2], r);
        }
        if (coreStroke.headEnd) {
          const r = lineEndRetract(coreStroke.headEnd, coreStroke, scale);
          pts[0] = retractLineEndpoint(pts[0], pts[1], r);
        }
        applyStroke(ctx as CanvasRenderingContext2D, coreStroke, scale);
        ctx.beginPath();
        ctx.moveTo(pts[0].x, pts[0].y);
        for (let i = 1; i < pts.length; i++) ctx.lineTo(pts[i].x, pts[i].y);
        ctx.stroke();
      }
      if (coreStroke.tailEnd) {
        drawArrowHead(ctx as CanvasRenderingContext2D, anchors.end.x, anchors.end.y, anchors.end.angle, coreStroke.tailEnd, coreStroke, scale);
      }
      if (coreStroke.headEnd) {
        drawArrowHead(ctx as CanvasRenderingContext2D, anchors.start.x, anchors.start.y, anchors.start.angle, coreStroke.headEnd, coreStroke, scale);
      }
    }
  }
  ctx.restore();

  // Body text inside the shape (wps:txbx). Drawn AFTER fill/stroke so text
  // sits on top of the panel. Rotation is intentionally not applied to body
  // text — the cover-template usage we care about uses anchor-only text.
  if (shape.textBlocks && shape.textBlocks.length > 0) {
    renderShapeText(shape, x, y, w, h, ctx as CanvasRenderingContext2D, scale, state.fontFamilyClasses, state.images);
  }
}

/** Fit an image block to the text-box inner width, preserving aspect from its
 *  natural pt size. If the natural width already fits, draw at natural size ×
 *  scale; otherwise scale down to innerW. Falls back to a square innerW box when
 *  the natural size is unknown (0). Returns px dimensions. */
/** Greedy line-wrap for text-box body text within `maxWidth` px (ECMA-376
 *  §21.1.2.1.1 — text-box content wraps to the inset box width unless wrap is
 *  off). Latin words break at spaces (the space stays with the preceding word);
 *  CJK / ideographic characters may break between any two (they carry no
 *  inter-word spaces). `ctx.font` must already be the block's font. A single
 *  token wider than `maxWidth` is left to overflow its own line (no
 *  hyphenation). Always returns at least one line. */
// Ideographic / CJK classification for the shape-text tokenizers uses the
// canonical `isCjkBreakChar` from @silurus/ooxml-core (imported above), the same
// predicate the body's wrap/justify paths use. It covers U+3000–U+9FFF, the
// Hangul Syllables block U+AC00–U+D7A3, U+F900–U+FAFF and U+FF00–U+FFEF — so
// Korean text-box text is classified as CJK (and takes the eastAsia face), which
// the previous local `isCjkCp` dropped. NOTE: that local predicate also covered
// the SIP Ext-B range U+20000–U+2FA1F; `isCjkBreakChar` does not, so it is
// intentionally dropped here. If Ext-B ideographs ever need CJK treatment, the
// fix belongs in the core predicate (shared by pptx/docx/xlsx), not a docx-local
// re-fork.

/** Per-token font family for a shape-text run, picked by the token's script
 *  (ECMA-376 §17.3.2.26): a CJK token (its first code point is East-Asian) uses
 *  the run's eastAsia axis, a Latin/digit token uses the ascii axis. The
 *  tokenizer ({@link tokenizeShapeText}) makes every token homogeneous — one CJK
 *  char, or a Latin word incl. its trailing space — so the first code point
 *  classifies the whole token. Falls back to the ascii `fontFamily` when the
 *  eastAsia axis is absent (older parser output / single-axis runs). The font
 *  CLASS (serif/sans) then comes from `fontFamilyClasses` (fontTable §17.8.3.10),
 *  so a serif ascii face and a gothic eastAsia face render in their own styles
 *  with no name heuristics. */
function shapeTokenFamily(token: string, run: ShapeTextRun): string | null {
  const cp = token.codePointAt(0) ?? 0;
  return isCjkBreakChar(cp) ? (run.fontFamilyEastAsia ?? run.fontFamily ?? null) : (run.fontFamily ?? null);
}

/** Split a string into atomic wrap units: each CJK char alone, or a run of
 *  non-CJK characters up to and including a trailing space. Shared by the
 *  single-format ({@link wrapShapeText}) and rich ({@link wrapShapeRuns})
 *  text-box layout paths so they tokenize identically. */
function tokenizeShapeText(text: string): string[] {
  const tokens: string[] = [];
  let buf = '';
  for (const ch of text) {
    const cp = ch.codePointAt(0) ?? 0;
    if (isCjkBreakChar(cp)) {
      if (buf) {
        tokens.push(buf);
        buf = '';
      }
      tokens.push(ch);
    } else if (ch === ' ') {
      buf += ch;
      tokens.push(buf);
      buf = '';
    } else {
      buf += ch;
    }
  }
  if (buf) tokens.push(buf);
  return tokens;
}

function wrapShapeText(
  ctx: CanvasRenderingContext2D,
  text: string,
  maxWidth: number,
  // ECMA-376 §17.3.1.12 — the FIRST line wraps to this width (paraW − firstLine
  // indent); continuation lines use `maxWidth` (paraW). Defaults to `maxWidth`
  // so a no-indent caller keeps the original single-width wrap exactly.
  firstLineWidth: number = maxWidth,
): string[] {
  if (!text) return [''];
  if (maxWidth <= 0) return [text];
  const tokens = tokenizeShapeText(text);

  const lines: string[] = [];
  let cur = '';
  // The first line being filled uses firstLineWidth; once a line wraps the
  // remaining lines use maxWidth.
  let limit = firstLineWidth;
  for (const tok of tokens) {
    if (cur !== '' && ctx.measureText(cur + tok).width > limit) {
      lines.push(cur.replace(/\s+$/, ''));
      cur = tok.replace(/^\s+/, ''); // a wrapped line never starts with a space
      limit = maxWidth;
    } else {
      cur += tok;
    }
  }
  if (cur !== '') lines.push(cur.replace(/\s+$/, ''));
  return lines.length ? lines : [text];
}

/** A single wrap unit tagged with the run it came from. `text` is a Latin word
 *  (incl. trailing space) or one CJK character (see {@link tokenizeShapeText}).
 *  `width` is its measured advance in px under the run's font (filled during
 *  greedy line-fill). */
interface RichToken {
  text: string;
  run: ShapeTextRun;
  width: number;
}

/** Greedy line-wrap for a rich (mixed-format) text-box paragraph. Builds one
 *  flat token stream across all `runs` (each token carrying its run's format),
 *  measures every token under its own font, and fills lines to `maxWidth` px —
 *  the same wrap rule {@link wrapShapeText} uses, but per-token-font-aware.
 *  A leading space is dropped when a line wraps (a wrapped line never starts
 *  with a space); a token wider than `maxWidth` stays on its own line. Mutates
 *  `ctx.font` while measuring. Always returns at least one (possibly empty)
 *  line. */
function wrapShapeRuns(
  ctx: CanvasRenderingContext2D,
  runs: ShapeTextRun[],
  maxWidth: number,
  scale: number,
  fontFamilyClasses: Record<string, string>,
  // ECMA-376 §17.3.1.12 — first line wraps to this width (paraW − firstLine
  // indent); continuation lines use `maxWidth` (paraW). Defaults to `maxWidth`
  // so a no-indent caller keeps the original single-width wrap exactly.
  firstLineWidth: number = maxWidth,
): RichToken[][] {
  const tokens: RichToken[] = [];
  for (const run of runs) {
    const fontPx = run.fontSizePt * scale;
    for (const text of tokenizeShapeText(run.text)) {
      // Measure each token under its OWN per-character font (ascii vs eastAsia
      // axis, §17.3.2.26) so a CJK glyph's advance is read from the eastAsia face
      // and a Latin/digit glyph's from the ascii face.
      ctx.font = buildFont(run.bold ?? false, run.italic ?? false, fontPx, shapeTokenFamily(text, run), fontFamilyClasses);
      tokens.push({ text, run, width: ctx.measureText(text).width });
    }
  }
  if (tokens.length === 0) return [[]];

  const lines: RichToken[][] = [];
  let cur: RichToken[] = [];
  let curW = 0;
  // The first line being filled uses firstLineWidth; once a line wraps the
  // remaining lines use maxWidth.
  let limit = firstLineWidth;
  for (const tok of tokens) {
    if (cur.length > 0 && curW + tok.width > limit) {
      lines.push(cur);
      limit = maxWidth;
      // A wrapped line never starts with a space — drop a leading space token.
      if (tok.text.trim() === '') {
        cur = [];
        curW = 0;
        continue;
      }
      cur = [tok];
      curW = tok.width;
    } else {
      cur.push(tok);
      curW += tok.width;
    }
  }
  if (cur.length > 0) lines.push(cur);
  return lines.length ? lines : [[]];
}

function fitShapeImage(
  widthPt: number,
  heightPt: number,
  innerW: number,
  scale: number,
): { w: number; h: number } {
  const natW = (widthPt ?? 0) * scale;
  const natH = (heightPt ?? 0) * scale;
  if (natW <= 0 || natH <= 0) {
    // No intrinsic size surfaced — reserve a square innerW box.
    return { w: innerW, h: innerW };
  }
  if (natW <= innerW) return { w: natW, h: natH };
  const s = innerW / natW;
  return { w: innerW, h: natH * s };
}

/** Render a shape's body text inside its bounding box, honoring lIns/tIns/
 *  rIns/bIns and the wps:bodyPr @anchor (t / ctr / b). Alignment within each
 *  line is read from the per-block paragraph alignment.
 *
 *  Blocks carrying an `imagePath` (an inline image inside the text box, e.g. a
 *  WMF chart wrapped as the sole content of a paragraph) draw the decoded
 *  bitmap from `images` instead of text, fitted to the inner width. The
 *  reserved height is the SAME value used by the first-pass measurement and the
 *  draw advance, so vertical anchoring (t/ctr/b) stays consistent. A missing
 *  bitmap reserves its height but draws nothing (no crash).
 *
 *  Exported for unit testing the inline-image fit/draw + missing-bitmap paths. */
export function renderShapeText(
  shape: ShapeRun,
  x: number, y: number, w: number, h: number,
  ctx: CanvasRenderingContext2D,
  scale: number,
  fontFamilyClasses: Record<string, string> = {},
  images: Map<string, DecodedImage> = new Map(),
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

  // ECMA-376 §17.3.1.12 — per-paragraph indent (px). `leftPx`/`rightPx` shrink
  // the text column from the inner box edges; `firstPx` is the SIGNED first-line
  // indent (positive = first line further right; negative = first line hangs
  // LEFT, so its width is WIDER). The body renderer honors the sign the same way
  // (Word applies a signed hanging first-line list-independently); the shape path
  // mirrors it rather than clamping. `paraW` is the continuation-line width;
  // `firstLineW` the first line's. When all indents are 0 ⇒ leftPx=rightPx=
  // firstPx=0 ⇒ paraW=firstLineW=innerW and regionLeft=innerX (no-op).
  const indentOf = (b: ShapeText) => {
    const leftPx = (b.indentLeft ?? 0) * scale;
    const rightPx = (b.indentRight ?? 0) * scale;
    const firstPx = (b.indentFirst ?? 0) * scale; // SIGNED
    const paraW = Math.max(0, innerW - leftPx - rightPx);
    const firstLineW = Math.max(0, paraW - firstPx);
    return { leftPx, firstPx, paraW, firstLineW };
  };

  // First pass: lay out each block. Text blocks WRAP to the inner width
  // (ECMA-376 §21.1.2.1.1) — a long title/abstract that exceeds the box width
  // breaks onto multiple lines instead of overflowing the page; image blocks
  // reserve their fitted height. The computed layout drives both vertical
  // anchoring (totalH) and the draw pass (no re-wrapping).
  type BlockIndent = { leftPx: number; firstPx: number; paraW: number; firstLineW: number };
  type BlockLayout =
    | { kind: 'image'; fitW: number; fitH: number; ind: BlockIndent }
    | { kind: 'text'; lines: string[]; lineH: number; baselineOffset: number; ind: BlockIndent }
    | { kind: 'rich'; lines: RichToken[][]; lineHeights: number[]; baselineOffsets: number[]; ind: BlockIndent };
  // ECMA-376 line box: the font's NATURAL line height (OS/2 win metrics, read via
  // the browser's fontBoundingBox and corrected for substituted faces by
  // correctLineMetrics), NOT a flat 1.2×em. The flat factor understates real
  // faces, so a text box's trailing line stayed inside a `noAutofit` box that
  // Word clips (sample-6's 3-line banner). The §17.3.1.33 line-spacing rule is
  // applied by the SHARED `lineBoxHeight` (single source of truth with the body
  // renderer, incl. the `intendedSingleLinePx` floor that keeps a substituted
  // CJK face — Meiryo — from under-measuring). Returns the CENTERED baseline
  // offset (half-leading) so the draw pass seats glyphs the way Word centers the
  // font's natural line within an expanded line box, instead of top-aligning.
  const shapeLineMetrics = (
    family: string | null | undefined,
    bold: boolean,
    italic: boolean,
    fontPx: number,
    b: ShapeText,
    familyEa?: string | null | undefined,
    // Explicit design-line floor (px). The body floors the line box by the MAX
    // intendedSingleLinePx across ALL segments on the line (~5684), not just the
    // tallest one, because a shorter-but-tabled face (e.g. a small Meiryo run
    // after a large untabled-ascii run) still raises the box. The rich call site
    // passes that all-runs max here; when omitted (the single-format path) the
    // floor falls back to this run's own ascii+eastAsia faces. Either way it is a
    // FLOOR (0 for all-untabled lines ⇒ unchanged).
    lineFloorPx?: number,
  ): { lineH: number; baselineOffset: number } => {
    // Floor the single-line box by the TALLEST design line among the run's
    // declared faces (ascii §17.3.2.26 + eastAsia). The common Japanese encoding
    // sets Meiryo (1.596×em) / Sakkal Majalla (1.3965×em) ONLY on
    // `<w:rFonts w:eastAsia>` while `<w:rFonts w:ascii>` stays an untabled Latin
    // default, so an ascii-only floor would leave the box flat
    // (intendedSingleLinePx(untabledAscii)=0). This is a FLOOR, not a replace —
    // intendedSingleLinePx returns 0 for every untabled face, so non-Meiryo/
    // Sakkal text boxes are unchanged. Mirrors the xlsx shape-text floor
    // (PR #646) and the docx BODY per-eastAsia-segment floor. It is NOT per-glyph
    // CJK font switching (a larger change deferred here).
    const asciiIntended = intendedSingleLinePx(family ?? null, fontPx);
    const eaIntended = intendedSingleLinePx(familyEa ?? null, fontPx);
    // Use the explicit all-runs floor when provided (rich path); else fall back
    // to this run's own faces (single-format path). Both are a floor, so the
    // per-run ascii/eastAsia max is still folded in for the fallback.
    const intended = lineFloorPx ?? Math.max(asciiIntended, eaIntended);
    // Measure the glyph metrics on the RENDERING face — the one whose design line
    // wins the floor. When the eastAsia face is tabled and taller than the ascii
    // box (the CJK glyphs are the tallest ink on the line), read c.ascent/descent
    // from `familyEa`, not the untabled ascii default, so the baseline is placed
    // relative to the CJK ink — mirroring the body's `line.ascent` being the
    // eastAsia-resolved corrected ascent. If the ascii face wins (Latin line, or
    // both untabled) keep measuring ascii, so an all-untabled line is byte-for-
    // byte unchanged from before this change.
    const measureFamily = eaIntended > asciiIntended ? (familyEa ?? null) : (family ?? null);
    ctx.font = buildFont(bold, italic, fontPx, measureFamily, fontFamilyClasses);
    const m = ctx.measureText('Mg');
    const rawAsc = m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent ?? fontPx * 0.8;
    const rawDesc = m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent ?? fontPx * 0.2;
    const c = correctLineMetrics(measureFamily, fontPx, rawAsc, rawDesc);
    // The glyph box (measured, corrected) is kept SEPARATE from the floored line
    // box: `intended` may inflate `lineH` above the glyph box, and Word centers
    // the glyph box within that expanded box (half-leading) rather than pinning
    // the ink to the top. Mirrors the body's `glyphNatural = ascent + descent`
    // (lineBoxHeight) vs the floor-inflated `natural`.
    const glyphNatural = c.ascent + c.descent;
    const ls: LineSpacing | null = b.lineSpacingRule
      ? { value: b.lineSpacingVal ?? 0, rule: b.lineSpacingRule as 'auto' | 'exact' | 'atLeast' }
      : null;
    const lineH = lineBoxHeight(ls, c.ascent, c.descent, scale, undefined, false, intended, false);
    // Word centers the font's GLYPH box within the (possibly expanded) line box
    // (half-leading): when line-spacing or the design-line floor grows the box
    // the extra space is split above and below the glyph box, so the baseline is
    // (lineH − glyphNatural)/2 below the box top plus the ascent. With no
    // expansion (glyphNatural == lineH) this reduces to c.ascent. This is the
    // SAME math as the body draw baseline (naturalLineH = ascent+descent, NOT
    // floor-inflated) — so an inflated box floats the glyphs centered with real
    // half-leading instead of top-pinning them.
    const baselineOffset = (lineH - glyphNatural) / 2 + c.ascent;
    return { lineH, baselineOffset };
  };
  const layouts: BlockLayout[] = blocks.map((b) => {
    const ind = indentOf(b);
    if (b.imagePath) {
      // The image occupies the FIRST line, so it fits to firstLineW (= paraW −
      // signed first-line indent), not the full inner width.
      const { w: fitW, h: fitH } = fitShapeImage(b.imageWidthPt ?? 0, b.imageHeightPt ?? 0, ind.firstLineW, scale);
      return { kind: 'image', fitW, fitH, ind };
    }
    // Rich path: a paragraph with explicit per-run formatting lays out as mixed
    // fonts. Each line's height is the TALLEST run's natural line box (ECMA-376
    // line box ≈ largest font on the line).
    if (b.runs && b.runs.length > 0) {
      const lines = wrapShapeRuns(ctx, b.runs, ind.paraW, scale, fontFamilyClasses, ind.firstLineW);
      const metrics = lines.map((toks) => {
        const tallest = toks.reduce<RichToken | null>(
          (best, t) => (best && best.run.fontSizePt >= t.run.fontSizePt ? best : t),
          null,
        );
        const run = tallest?.run;
        const fontPx = (run?.fontSizePt ?? b.fontSizePt) * scale;
        // Design-line floor is the MAX intendedSingleLinePx over EVERY run on the
        // line — both its ascii AND eastAsia face (§17.3.2.26), each at that run's
        // own size — not just the tallest run's faces. The tallest run may be an
        // untabled Latin face while a shorter-but-equal-or-smaller Meiryo run
        // still raises the box to Meiryo's design line; flooring on the tallest
        // run alone missed that. Mirrors the body's per-segment max (~5684-5685).
        // 0 for an all-untabled line ⇒ unchanged.
        const lineFloorPx = toks.reduce((floor, t) => {
          const runPx = t.run.fontSizePt * scale;
          return Math.max(
            floor,
            intendedSingleLinePx(t.run.fontFamily ?? null, runPx),
            intendedSingleLinePx(t.run.fontFamilyEastAsia ?? null, runPx),
          );
        }, 0);
        return shapeLineMetrics(
          run?.fontFamily ?? b.fontFamily,
          run?.bold ?? false,
          run?.italic ?? false,
          fontPx,
          b,
          // eastAsia axis (§17.3.2.26) of the TALLEST run selects the measurement
          // face inside shapeLineMetrics (the glyph box read for the baseline).
          // The design-line FLOOR, however, is the all-runs max computed above and
          // passed explicitly — so a shorter tabled run still raises the box.
          run?.fontFamilyEastAsia,
          lineFloorPx,
        );
      });
      return {
        kind: 'rich',
        lines,
        lineHeights: metrics.map((x) => x.lineH),
        baselineOffsets: metrics.map((x) => x.baselineOffset),
        ind,
      };
    }
    const fontPx = b.fontSizePt * scale;
    // Single-format (legacy, run-less) fallback. SCOPED OUT of the eastAsia
    // line-box floor: block-level `ShapeText` carries only the CONFLATED
    // `fontFamily` (ascii → eastAsia → default, parser.rs resolve_font_with_default)
    // and has NO block-level eastAsia field on the Rust `ShapeText` struct
    // (types.rs) or the TS interface (types.ts) — so there is no untabled-ascii /
    // tabled-eastAsia split to floor against here. In practice the parser emits a
    // `ShapeTextRun` for every `<w:r>` with text (parser.rs, extract loop), so a
    // TEXT paragraph always takes the rich path above; this branch only serves
    // image-less legacy blocks. Wiring an eastAsia floor here would need a new
    // block-level `font_family_east_asia` field threaded through types.rs + the
    // parser + types.ts + a WASM rebuild — nontrivial plumbing for a path the
    // current parser never feeds text into. Deferred as a follow-up rather than
    // half-plumbed. The rich path (ShapeTextRun.fontFamilyEastAsia) fixes the
    // common case (PR #646 mirror).
    const { lineH, baselineOffset } = shapeLineMetrics(b.fontFamily, b.bold ?? false, b.italic ?? false, fontPx, b);
    return { kind: 'text', lines: wrapShapeText(ctx, b.text, ind.paraW, ind.firstLineW), lineH, baselineOffset, ind };
  });
  const blockHeight = (l: BlockLayout): number => {
    if (l.kind === 'image') return l.fitH;
    if (l.kind === 'rich') return l.lineHeights.reduce((s, h) => s + h, 0);
    return l.lines.length * l.lineH;
  };
  // ECMA-376 §17.3.1.33 — each text-box paragraph's own spaceBefore/After is
  // reserved inside the box (px). sample-13's "Journal homepage" line carries
  // spaceBefore = 50 pt, which drops it well below the box top so it clears the
  // masthead banner instead of hiding behind it. The gap BETWEEN two paragraphs
  // is max(prev.after, this.before) — NOT their sum (same collapse the body uses);
  // the first block reserves only its own before, and a trailing after overflows
  // the box (it is not part of the laid-out block extent, so it is excluded from
  // totalH used for ctr/bottom anchoring).
  const spBefore = blocks.map((b) => (b.spaceBefore ?? 0) * scale);
  const spAfter = blocks.map((b) => (b.spaceAfter ?? 0) * scale);
  const gapBefore = (i: number): number =>
    i > 0 ? Math.max(spBefore[i], spAfter[i - 1]) : spBefore[i];
  const totalH = layouts.reduce((s, l, i) => s + gapBefore(i) + blockHeight(l), 0);

  const anchor = shape.textAnchor ?? 't';
  let cursorY: number;
  if (anchor === 'b') {
    cursorY = innerY + Math.max(0, innerH - totalH);
  } else if (anchor === 'ctr') {
    cursorY = innerY + Math.max(0, (innerH - totalH) / 2);
  } else {
    cursorY = innerY;
  }

  // ECMA-376 §21.1.2.1.1 — a `<a:noAutofit/>` (normalized to "none") text box
  // keeps a FIXED size and Word CLIPS text that overflows the box ("sp"/spAutoFit
  // grows the box, "norm"/normAutofit shrinks the text, so only "none" needs a
  // clip: the box is already the resolved size for the other modes). Clip to the
  // shape's box so an overflowing trailing line is hidden exactly as Word does —
  // e.g. sample-6's 3-line banner box whose 3rd line ("All mccp … Creative
  // Commons licence") sits below the 82 pt box and is not shown in Word.
  //
  // Deliberate cross-package divergence (do NOT unify): PowerPoint does NOT clip
  // its shape text — an overflowing paragraph renders past the shape bounds — so
  // the pptx renderer intentionally omits this clip (see packages/pptx/src/
  // renderer.ts, "does NOT clip text that overflows its shape"). Word clips a
  // fixed box; PowerPoint overflows. Same bodyPr concept, different app behavior.
  const clipToBox = shape.textAutofit === 'none';
  if (clipToBox) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(x, y, w, h);
    ctx.clip();
  }

  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    const layout = layouts[i];

    // Reserve the gap above this paragraph: its own spaceBefore collapsed with
    // the previous paragraph's spaceAfter (max, not sum — §17.3.1.33).
    cursorY += gapBefore(i);

    if (layout.kind === 'image') {
      // Inline image inside the text box. The image is the paragraph's FIRST
      // (only) line, so it lives in the first-line region: region-left =
      // innerX + leftPx + firstPx, region-width = firstLineW (ECMA-376
      // §17.3.1.12). Place horizontally per the paragraph alignment (figures
      // default to centered), and advance by the reserved height regardless of
      // whether a bitmap is present (a missing decode must not shift the rest of
      // the layout).
      const { fitW, fitH, ind } = layout;
      const regionLeft = innerX + ind.leftPx + ind.firstPx;
      const regionW = ind.firstLineW;
      const bmp = block.imagePath ? images.get(imageKey(block.imagePath)) : undefined;
      if (bmp) {
        let drawX = regionLeft + Math.max(0, (regionW - fitW) / 2); // default: centered
        if (block.alignment === 'left' || block.alignment === 'both') {
          drawX = regionLeft;
        } else if (block.alignment === 'right') {
          drawX = regionLeft + Math.max(0, regionW - fitW);
        }
        ctx.drawImage(bmp, drawX, cursorY, fitW, fitH);
      }
      cursorY += fitH;
      continue;
    }

    if (layout.kind === 'rich') {
      // Mixed-format paragraph: each token carries its own run's font/color.
      // 'distribute' stays centered (no inter-word stretch in shape text), like
      // the single-format path.
      const edgeFor = (rtl: boolean) =>
        block.alignment === 'distribute' ? 'center' : resolveAlignEdge(block.alignment, rtl);
      ctx.textAlign = 'left';
      const ind = layout.ind;
      for (let li = 0; li < layout.lines.length; li++) {
        const lineToks = layout.lines[li];
        const lineH = layout.lineHeights[li];
        const lineW = lineToks.reduce((s, t) => s + t.width, 0);
        // ECMA-376 §17.3.1.12 — align within the per-line INDENTED region: the
        // first line carries the signed first-line indent, continuation lines do
        // not. region-left = innerX + leftPx (+ firstPx on the first line);
        // region-width = firstLineW (first) / paraW (continuation).
        const isFirstLine = li === 0;
        const regionLeft = innerX + ind.leftPx + (isFirstLine ? ind.firstPx : 0);
        const regionW = isFirstLine ? ind.firstLineW : ind.paraW;
        // Base direction (first-strong) from the line's own text, matching the
        // single-format path's per-block resolution but resolved per line.
        const lineText = lineToks.map((t) => t.text).join('');
        const baseRtl = resolveBaseDirection(undefined, lineText) === 'rtl';
        const edge = edgeFor(baseRtl);
        let tx = regionLeft;
        if (edge === 'center') {
          tx = regionLeft + Math.max(0, (regionW - lineW) / 2);
        } else if (edge === 'right') {
          tx = regionLeft + Math.max(0, regionW - lineW);
        }
        // Baseline sits at the CENTERED offset below the line-box top (half-
        // leading, metric-based): the natural line is centered within the
        // (possibly expanded) line height used for advancing cursorY, matching
        // the body renderer instead of top-aligning.
        const baseline = cursorY + layout.baselineOffsets[li];
        // UAX#9 visual reorder (rule L2), the SAME pass body paragraphs use. A
        // rich line draws one token per fillText, so — unlike the single-fillText
        // plain path below, where the canvas reorders internally — the tokens
        // must be reordered HERE or RTL/mixed text draws in logical order (Word
        // reverses it). Shape paragraphs carry no explicit rtl flag, so the base
        // direction is the line's first-strong char; ctx.textAlign stays 'left'
        // so tx is each token's left edge and ctx.direction is set per token to
        // its resolved level (so a Latin/number island still shapes LTR).
        const visual = computeLineVisualOrder(lineToks, baseRtl);
        for (let vi = 0; vi < lineToks.length; vi++) {
          const si = visual.order[vi];
          const tok = lineToks[si];
          ctx.direction = visual.rtl[si] ? 'rtl' : 'ltr';
          const fontPx = tok.run.fontSizePt * scale;
          // Per-character font (ascii vs eastAsia axis, §17.3.2.26): a CJK token
          // draws with the eastAsia family, a Latin/digit token with the ascii
          // family. Mirrors the measure pass so advance and draw agree.
          ctx.font = buildFont(tok.run.bold ?? false, tok.run.italic ?? false, fontPx, shapeTokenFamily(tok.text, tok.run), fontFamilyClasses);
          ctx.fillStyle = tok.run.color ? `#${tok.run.color}` : '#000000';
          ctx.fillText(tok.text, tx, baseline);
          tx += tok.width;
        }
        cursorY += lineH;
      }
      continue;
    }

    const fontPx = block.fontSizePt * scale;
    ctx.font = buildFont(block.bold ?? false, block.italic ?? false, fontPx, block.fontFamily ?? null, fontFamilyClasses);
    ctx.fillStyle = block.color ? `#${block.color}` : '#000000';
    // Base direction (for neutral resolution) + alignment, derived from the
    // content (first-strong) since shape paragraphs carry no explicit rtl flag.
    const baseRtl = resolveBaseDirection(undefined, block.text) === 'rtl';
    // No inter-word stretching in shape text, so 'distribute' stays centered
    // rather than the justify edge resolveAlignEdge reports for paragraphs.
    const edge = block.alignment === 'distribute'
      ? 'center'
      : resolveAlignEdge(block.alignment, baseRtl);
    ctx.textAlign = 'left';
    ctx.direction = baseRtl ? 'rtl' : 'ltr';
    const ind = layout.ind;
    for (let li = 0; li < layout.lines.length; li++) {
      const line = layout.lines[li];
      const m = ctx.measureText(line);
      // ECMA-376 §17.3.1.12 — align within the per-line INDENTED region: the
      // first line carries the signed first-line indent, continuation lines do
      // not (region-left = innerX + leftPx (+ firstPx on the first line);
      // region-width = firstLineW (first) / paraW (continuation)).
      const isFirstLine = li === 0;
      const regionLeft = innerX + ind.leftPx + (isFirstLine ? ind.firstPx : 0);
      const regionW = isFirstLine ? ind.firstLineW : ind.paraW;
      let tx = regionLeft;
      if (edge === 'center') {
        tx = regionLeft + Math.max(0, (regionW - m.width) / 2);
      } else if (edge === 'right') {
        tx = regionLeft + Math.max(0, regionW - m.width);
      }
      // Baseline = line top + the CENTERED offset (half-leading, metric-based):
      // the natural line is centered within the (possibly expanded) line height
      // used to advance cursorY, matching the body renderer instead of top-
      // aligning.
      const baseline = cursorY + layout.baselineOffset;
      ctx.fillText(line, tx, baseline);
      cursorY += layout.lineH;
    }
  }
  if (clipToBox) ctx.restore();
  ctx.direction = 'ltr'; // reset for subsequent draws
}

/**
 * Resolve an anchor image's page-space box origin and dist* padding (px), shared
 * by registerAnchorFloats (wrap floats) and renderAnchorImages (wrapNone images).
 *
 * X: margin-relative offsets add section.marginLeft (ECMA-376 §20.4.3.4
 * relativeFrom="margin"); otherwise anchorXPt is already page-absolute.
 * Y: paragraph-relative offsets add `paraBaseY`; otherwise page-absolute. The
 * caller supplies `paraBaseY` = the paragraph's pre-spaceBefore TOP for ALL
 * paragraph-relative floats — wrap and wrapNone alike (ECMA-376 §20.4.3.5: a
 * `positionV relativeFrom="paragraph"` float is positioned relative to the
 * paragraph that contains the anchor, i.e. its top edge before spaceBefore).
 * Page-level floats pass 0 (resolveAnchorY ignores paraBaseY for them). This is
 * the box origin BEFORE any overlap displacement; resolveFloatOverlap runs on
 * top of it for floats.
 *
 * Exported under a `_test` alias for the anchor-image relativeFrom wiring test
 * (the public renderer entry points consume the box internally; pin the
 * positionH/V → xContainer/yContainer plumbing at this seam).
 */
export const __test_resolveAnchorBox = (
  img: ImageRun,
  state: RenderState,
  paraBaseY: number,
): { x: number; y: number; w: number; h: number; dl: number; dr: number; dt: number; db: number } =>
  resolveAnchorBox(img, state, paraBaseY);

/** Exported for the page-anchor pre-scan test (ECMA-376 §20.4.3.2/§20.4.3.5):
 *  drives {@link preRegisterPageFloats} from a unit test against a stub
 *  RenderState so we can pin which paragraphs get pre-registered and that
 *  duplicate calls are idempotent. */
export const __test_preRegisterPageFloats = (
  body: readonly (BodyElement | PaginatedBodyElement)[],
  startIdx: number,
  state: RenderState,
): void => preRegisterPageFloats(body, startIdx, state);

/** Exported for the page-anchor pre-scan test — pins the
 *  {paragraph,line,character} ⇒ paragraph-local vs everything-else ⇒ page-level
 *  classification (ECMA-376 §20.4.3.5 ST_RelFromV). */
export const __test_isPageLevelAnchorY = (
  rf: string | null | undefined,
  fromPara: boolean,
): boolean => isPageLevelAnchorY(rf, fromPara);

/** Exported for the scale-invariance characterization test (Phase 4-1 B2 Stage 1).
 *  {@link layoutLines} is the single line-breaking + measurement kernel called by
 *  BOTH the paginator (scale 1, pt space) and the paint pass (device scale). The
 *  compute-once refactor stamps the paginator's scale-1 lines and rehydrates them
 *  at the paint scale, so the test pins EXACTLY which `LayoutLine` fields scale
 *  linearly with the passed `scale` (px advances, ascent/descent, xOffset,
 *  availWidth, topY) versus which are scale-invariant (segment count, line count,
 *  pt `height`) — and surfaces any field/branch that is NOT a clean ×scale. */
export const __test_layoutLines = layoutLines;
export type { LayoutSeg as __test_LayoutSeg, LayoutLine as __test_LayoutLine, LayoutTextSeg as __test_LayoutTextSeg, WrapLayoutCtx as __test_WrapLayoutCtx };

/** Exported for the compute-once pixel-identity test (Phase 4-1 B2 Stage 1).
 *  Toggles the stamped-line reuse in renderParagraph so the test can render the
 *  SAME page with reuse ON and OFF and assert the paint call streams are
 *  byte-identical. Returns the previous value so the test can restore it. */
export const __test_setLineReuseEnabled = (v: boolean): boolean => {
  const prev = lineReuseEnabled;
  lineReuseEnabled = v;
  return prev;
};

function resolveAnchorBox(
  img: ImageRun,
  state: RenderState,
  paraBaseY: number,
): { x: number; y: number; w: number; h: number; dl: number; dr: number; dt: number; db: number } {
  const scale = state.scale;
  const w = img.widthPt * scale;
  const h = img.heightPt * scale;
  // ECMA-376 §20.4.3.1 wp:align — when positionH/V carry <wp:align>, the
  // renderer aligns the image within its relativeFrom container instead of
  // using the (discarded) posOffset. Mirrors resolveShapeBox (the ShapeRun
  // equivalent): we route X/Y through resolveAnchorX/Y with the image's own
  // box size as the align size. The raw §20.4.3.2/§20.4.3.5
  // `<wp:positionH/V>@relativeFrom` string (e.g. "margin", "topMargin") is
  // threaded through so xContainer/yContainer pick the correct container.
  // Without it a `relativeFrom="margin"` + `align="top"` image would degrade
  // to the page-relative top edge (Y=0 → inside the top margin), which is
  // exactly the sample-11 misplacement before this wire-up. ImageRun carries
  // no pctPos/sizeRel, so those args remain null and the legacy boolean
  // anchorXFromMargin / anchorYFromPara hints still gate page-vs-margin when
  // no raw relativeFrom is present. When align is absent, resolveAnchorX/Y
  // fall back to the offset path.
  const x = resolveAnchorX(
    img.anchorXAlign, img.anchorXFromMargin ?? false, img.anchorXPt ?? 0, w, state,
    img.anchorXRelativeFrom ?? null, null, null,
  );
  const y = resolveAnchorY(
    img.anchorYAlign, img.anchorYFromPara ?? false, img.anchorYPt ?? 0, h, paraBaseY, state,
    img.anchorYRelativeFrom ?? null, null, null,
  );
  return {
    x,
    y,
    w,
    h,
    dl: (img.distLeft   ?? 0) * scale,
    dr: (img.distRight  ?? 0) * scale,
    dt: (img.distTop    ?? 0) * scale,
    db: (img.distBottom ?? 0) * scale,
  };
}

/** ECMA-376 §20.4.3.2 / §20.4.3.5 — a `<wp:positionV>` `relativeFrom` value
 *  that resolves the float's Y INDEPENDENTLY of its anchoring paragraph (vs.
 *  `paragraph` / `line` / `character` which resolve against the paragraph's
 *  top). When Y is page-level, Word treats the float as page-positioned: it
 *  is laid out as soon as the page is opened and earlier paragraphs on the
 *  same page wrap around it. {@link preRegisterPageFloats} uses this to
 *  hoist such floats to page-start; paragraph-local Y still flows the legacy
 *  per-paragraph path.
 *
 *  An anchor with NO explicit `<wp:positionV>` (anchorYRelativeFrom absent)
 *  still resolves against the page top via the legacy hint
 *  (`anchorYFromPara=false` ⇒ page-absolute offset), so it qualifies as
 *  page-level too. */
function isPageLevelAnchorY(rf: string | null | undefined, fromPara: boolean): boolean {
  if (rf == null) return !fromPara;
  switch (rf) {
    case 'paragraph':
    case 'line':
    case 'character':
      return false;
    default:
      return true;
  }
}

/** True when this run is a wrap float whose vertical placement is page-level
 *  (independent of source-order paragraph position) — see
 *  {@link isPageLevelAnchorY}. `isWrapFloat` already filters inline images
 *  (their `wrapMode` is undefined) and non-wrapping anchors, so an extra
 *  `anchor` check is redundant here. */
function isPageLevelWrapFloat(run: ImageRun | ShapeRun): boolean {
  if (!isWrapFloat(run.wrapMode)) return false;
  return isPageLevelAnchorY(run.anchorYRelativeFrom ?? null, run.anchorYFromPara ?? false);
}

/** Register floats from a paragraph's anchor images and shapes. Anchor images
 *  are drawn immediately; anchor shapes are NOT drawn here (renderAnchorShape
 *  paints them separately) — we only reserve their float-exclusion band so body
 *  text wraps around them (ECMA-376 §20.4.2.16/.17), exactly like images.
 *
 *  Page-level floats (positionV relativeFrom ∈ {page, margin, *Margin, column},
 *  ECMA-376 §20.4.3.2/§20.4.3.5) are skipped when this paragraph was already
 *  pre-registered at the current page's start by {@link preRegisterPageFloats}
 *  — re-registering would double-stamp the FloatRect (and re-draw the image).
 *  Paragraph-local floats (`paragraph`/`line`/`character`) keep the per-
 *  paragraph path so their Y stays anchored at this paragraph's top. */
function registerAnchorFloats(para: DocParagraph, state: RenderState, paragraphAnchorY: number): void {
  // One id per registerAnchorFloats call ⇒ one id per paragraph. Floats sharing
  // a paraId (e.g. two side-by-side photos in one paragraph) never displace each
  // other; floats from different paragraphs do (de-facto overlap avoidance).
  const paraId = state.floatParaSeq++;
  const prescanned = state.pageAnchorPrescanned?.has(para) ?? false;
  for (const run of para.runs) {
    if (run.type === 'image') {
      const img = run as unknown as ImageRun;
      if (prescanned && isPageLevelWrapFloat(img)) continue;
      registerImageFloat(img, state, paragraphAnchorY, paraId);
    } else if (run.type === 'shape') {
      const shp = run as unknown as ShapeRun;
      if (prescanned && isPageLevelWrapFloat(shp)) continue;
      registerShapeFloat(shp, state, paragraphAnchorY, paraId);
    }
  }
}

/** Pre-scan upcoming body elements at a page-start moment and register any
 *  page-level (positionV relativeFrom ∈ {page, margin, *Margin, column})
 *  wrap floats they carry. Mirrors Word's layout order: page-level floats are
 *  positioned as soon as the page is opened, so paragraphs that PRECEDE the
 *  anchoring paragraph in source order on the same page wrap around them
 *  (ECMA-376 §20.4.3.2/§20.4.3.5 + §20.4.2.16/.17). Each pre-registered
 *  paragraph is recorded in `state.pageAnchorPrescanned` so the main flow's
 *  {@link registerAnchorFloats} skips its page-level runs (avoiding a
 *  duplicate FloatRect / re-drawn image) while still registering its
 *  paragraph-local floats normally.
 *
 *  Bounds: the scan stops at the next forced page boundary that the
 *  paginator/renderer is guaranteed to honor — an explicit `pageBreak`
 *  (§17.18.79 / §17.3.1.20) or a non-continuous `sectionBreak`. Content
 *  overflow may still push paragraphs to later pages mid-scan; the
 *  paginator's `newPage()` resets the float set wholesale, so those
 *  paragraphs get re-pre-scanned on the next page. (Same idempotent flow as
 *  the existing `registerAnchorFloats` post-newPage re-call at the split-
 *  relocation site.) */
function preRegisterPageFloats(
  body: readonly (BodyElement | PaginatedBodyElement)[],
  startIdx: number,
  state: RenderState,
): void {
  if (!state.pageAnchorPrescanned) state.pageAnchorPrescanned = new Set();
  for (let j = startIdx; j < body.length; j++) {
    const el = body[j];
    if (!el) continue;
    if (el.type === 'pageBreak') break;
    if (el.type === 'sectionBreak') {
      const sb = el as unknown as { kind?: string };
      if (sb.kind && sb.kind !== 'continuous') break;
      continue;
    }
    if (el.type !== 'paragraph') continue;
    const para = el as unknown as DocParagraph;
    // Skip if already pre-registered (renderer may call this once per page;
    // paginator may re-call after newPage(), but newPage clears the set).
    if (state.pageAnchorPrescanned.has(para)) continue;
    let hasPageLevel = false;
    for (const run of para.runs) {
      if (run.type === 'image') {
        if (isPageLevelWrapFloat(run as unknown as ImageRun)) { hasPageLevel = true; break; }
      } else if (run.type === 'shape') {
        if (isPageLevelWrapFloat(run as unknown as ShapeRun)) { hasPageLevel = true; break; }
      }
    }
    if (!hasPageLevel) continue;
    // Register only the page-level floats from this paragraph. paraY=0 is safe
    // because resolveAnchorY ignores it for page-level relativeFrom containers
    // (anchor-geometry.ts §20.4.3.x). Allocate a fresh paraId so overlap
    // avoidance treats these like any other anchor-paragraph float.
    const paraId = state.floatParaSeq++;
    for (const run of para.runs) {
      if (run.type === 'image') {
        const img = run as unknown as ImageRun;
        if (!isPageLevelWrapFloat(img)) continue;
        registerImageFloat(img, state, 0, paraId);
      } else if (run.type === 'shape') {
        const shp = run as unknown as ShapeRun;
        if (!isPageLevelWrapFloat(shp)) continue;
        registerShapeFloat(shp, state, 0, paraId);
      }
    }
    state.pageAnchorPrescanned.add(para);
  }
}

/** Reserve the float-exclusion rect for one anchored wrap-image and draw the
 *  bitmap immediately (the image is the float). */
function registerImageFloat(
  img: ImageRun,
  state: RenderState,
  paragraphAnchorY: number,
  paraId: number,
): void {
  if (!img.anchor) return;
  if (!isWrapFloat(img.wrapMode)) return;

  const mode: 'square' | 'topAndBottom' =
    img.wrapMode === 'topAndBottom' ? 'topAndBottom' : 'square';

  // Paragraph-relative wrap floats anchor at the pre-spaceBefore paragraph top
  // (paragraphAnchorY), per ECMA-376 §20.4.3.5 — identical to wrapNone images.
  const box = resolveAnchorBox(img, state, paragraphAnchorY);
  const { w, h, dl, dr, dt, db } = box;

  // Overlap avoidance. Spec-mandated part: allowOverlap="false" (ECMA-376
  // §20.4.2.3) REQUIRES repositioning to prevent overlap; "true"/omitted only
  // permits overlap. Default true per §20.4.2.3.
  // Implementation-defined (HEURISTIC, Word-mimicking, no ECMA-376 basis):
  // displacing the later document-order float, the "other paragraphs only"
  // gate under allowOverlap=true, and the right-then-down re-seat using dist
  // padding as the float-to-float gap. See resolveFloatOverlap header.
  const allowOverlap = img.allowOverlap ?? true;
  const key = imageKey(img.imagePath, img.colorReplaceFrom);
  const rect = pushFloatRect(state, {
    x: box.x,
    y: box.y,
    w, h, dl, dr, dt, db,
    kind: 'shape', // DrawingML anchor (§20.4.2.3); not a floating table.
    mode,
    side: img.wrapSide ?? 'bothSides',
    imageKey: key,
    drawn: false,
    paraId,
    avoidOverlap: true,
    allowOverlap,
  });

  if (!state.dryRun) {
    const bmp = state.images.get(key);
    if (bmp) drawImageCropped(state.ctx, bmp, img.srcRect ?? undefined, rect.imageX, rect.imageY, rect.imageW, rect.imageH);
    rect.drawn = true;
  }
}

/** Reserve the float-exclusion rect for one anchored wrap-shape (wps:txbx /
 *  DrawingML wp:anchor shape). The shape is drawn separately by
 *  renderAnchorShape, so here we only push the FloatRect (drawn=true ⇒ the
 *  deferred-image-draw path never tries to paint it). The box is resolved by the
 *  SAME resolveShapeBox the renderer draws with, so the band matches the shape. */
function registerShapeFloat(
  shape: ShapeRun,
  state: RenderState,
  paragraphAnchorY: number,
  paraId: number,
): void {
  if (!isWrapFloat(shape.wrapMode)) return;

  // Match resolveShapeBox's paragraphTopPx convention. resolveAnchorY reads
  // paragraphTopPx only for relativeFrom="paragraph"/"line" (anchorYFromPara);
  // wrap floats anchor at the pre-spaceBefore paragraph top (§20.4.3.5),
  // identical to the image path (resolveAnchorBox uses paragraphAnchorY there).
  const { x, y, w, h } = resolveShapeBox(shape, state, paragraphAnchorY);
  // A degenerate (zero/negative-area) box — e.g. a wrap-flagged line preset —
  // reserves no band; bail like renderAnchorShape skips drawing it.
  if (w <= 0 || h <= 0) return;

  const mode: 'square' | 'topAndBottom' =
    shape.wrapMode === 'topAndBottom' ? 'topAndBottom' : 'square';

  const scale = state.scale;
  const dl = (shape.distLeft   ?? 0) * scale;
  const dr = (shape.distRight  ?? 0) * scale;
  const dt = (shape.distTop    ?? 0) * scale;
  const db = (shape.distBottom ?? 0) * scale;

  // Overlap avoidance, kept consistent with the image path. Shapes carry no
  // parsed allowOverlap field; the spec default is true (§20.4.2.3), so
  // same-paragraph floats never displace each other and a lone shape is a no-op
  // here — but running it keeps multi-float behavior identical to images.
  pushFloatRect(state, {
    x, y, w, h, dl, dr, dt, db,
    kind: 'shape', // DrawingML wp:anchor shape (§20.4.2.3); not a floating table.
    mode,
    side: shape.wrapSide ?? 'bothSides',
    imageKey: '',
    // The shape is painted by renderAnchorShape, not by the deferred image-draw
    // path; mark it drawn so that path skips it (it has no bitmap to draw).
    drawn: true,
    paraId,
    avoidOverlap: true,
    allowOverlap: true,
  });
}

// ===== Table rendering =====

/** Per-column widths (px), total table width (px), and per-row heights (px,
 *  with the §17.4.85 vMerge-span extension applied) for a table laid out in a
 *  content band `contentWPx` wide. Shared by the block ({@link renderTable}) and
 *  floating ({@link renderFloatTable}) paths so both size the table identically. */
function computeTableLayout(
  table: DocTable,
  contentWPx: number,
  state: RenderState,
): { colWidths: number[]; tableW: number; rowHeights: number[] } {
  const { scale } = state;
  // Resolve column widths in pt (autofit by preferred widths, or fixed grid),
  // already scaled to fit the available content width, then convert to px.
  const colWidths = resolveColumnWidths(table, contentWPx / scale, state).map((w) => w * scale);
  const tableW = colWidths.reduce((s, w) => s + w, 0);

  // Shared ST_HeightRule + §17.4.85 vMerge-span skeleton (resolveTableRowHeights),
  // with the paint pass's px cell measurer. The restart-span extension is part of
  // the skeleton now — calculateRowHeight already excludes restart cells per-row,
  // and the resolver re-measures them via the same callback to grow the last row.
  const rowHeights = resolveTableRowHeights(table, colWidths, scale, (cell, cellW) =>
    measureCellContentHeightPx(cell, table, cellW, scale, state),
  );

  return { colWidths, tableW, rowHeights };
}

/** Content height (px, at `scale`) of a table cell laid out at total width
 *  `cellW`: cell top/bottom margins plus each content element measured at the
 *  paint pass's `measureCellElementHeight`. Adjacent paragraphs inside the cell
 *  collapse spacing the same way `renderCellContent` does (ECMA-376 §17.3.1.33
 *  contextualSpacing + spaceAfter/spaceBefore overlap = max not sum), so the
 *  measured height matches the painted height. Shared by the per-row skeleton
 *  (via computeTableLayout) and the exported {@link calculateRowHeight}. */
function measureCellContentHeightPx(
  cell: DocTableCell,
  table: DocTable,
  cellW: number,
  scale: number,
  state: RenderState,
): number {
  const cm = effCellMargins(cell, table);
  const contentW = cellW - (cm.left + cm.right) * scale;
  // ECMA-376 §17.4.7 requires every <w:tc> to end with a <w:p>. When the cell's
  // visible content is a nested table, Word emits a trailing empty <w:p/> purely
  // as that syntactic anchor; it carries no ink and does NOT grow the row (Word's
  // outer cell hugs the inner table — sample-11's "table inside a table" outer
  // row measures the inner table height, not inner + the structural mark's line
  // box + its inherited space-before). Drop it from the row-height measurement,
  // exactly as the vAlign block height already does (trimTrailingStructuralMarker).
  // The mark itself is still painted by renderCellContent; being empty it adds no
  // visible content, so excluding it from sizing cannot hide anything.
  const measured = trimTrailingStructuralMarker(cell.content);
  // measureCellElementHeight always includes paragraph spaceBefore+spaceAfter;
  // sumCellContentHeight folds in contextualSuppressed (§17.3.1.33) and the
  // prevSpaceAfter/spaceBefore overlap collapse to match the paint pass's
  // renderCellContent. Spacing is converted from pt to px with `scale`.
  return (cm.top + cm.bottom) * scale + sumCellContentHeight(
    measured,
    (ce) => measureCellElementHeight(state, ce, contentW, scale),
    scale,
  );
}

/** Draw all rows of a table whose grid origin is `tableX` (px) and whose top is
 *  `startY` (px), returning the Y just past the last row. Shared by the block
 *  and floating paths. Honors bidiVisual, vMerge span heights, and exact-row
 *  clipping exactly as the original inline loop did. In dryRun, measures cell
 *  content instead of drawing. */
function drawTableRows(
  table: DocTable,
  colWidths: number[],
  tableW: number,
  rowHeights: number[],
  tableX: number,
  startY: number,
  state: RenderState,
): number {
  const { scale, dryRun } = state;
  // ECMA-376 §17.4.1 `<w:bidiVisual>`: lay the grid columns right-to-left, so
  // logical column 0 sits at the table's RIGHT edge and indices advance
  // leftward. We mirror by POSITION arithmetic (not canvas transform): a cell
  // spanning [ci, ci+span) gets physical left x = tableX + tableW − (offset of
  // its right grid edge). Cell borders are mirrored too — a cell's logical
  // left/right border specs swap physical sides (its "start" edge is on the
  // right). gridSpan still consumes the same logical columns; only the mapping
  // from logical column offset to a physical x flips.
  const mirror = table.bidiVisual === true;

  // ECMA-376 §17.4.66 (border-collapse): a shared gridline must sit ON TOP of
  // every cell fill. Painting cell-by-cell (fill→border, per cell) let the next
  // column's background fill cover the half of the vertical border the previous
  // column had just drawn; with alternating row banding (e.g. Medium List 2)
  // this made a shared vertical rule look like its thickness changed row to row.
  // So walk the grid ONCE to collect every cell's paint box, then paint in TWO
  // passes: all backgrounds + content first, all borders second.
  const jobs: Array<{
    cell: DocTableCell;
    x: number;
    y: number;
    w: number;
    h: number;
    edges: CellEdgeFlags;
    clipExact: boolean;
  }> = [];

  let y = startY;
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
        let lastRowOfCell = ri;
        if (cell.vMerge === true) {
          const endRi = findMergeEndRow(table, ri, ci);
          lastRowOfCell = endRi;
          drawH = 0;
          for (let rj = ri; rj <= endRi; rj++) drawH += rowHeights[rj];
        }
        // ECMA-376 §17.4.38/§17.4.39: classify which physical edges of this cell
        // are the table's OUTER edges (vs. interior gridlines) from its grid
        // position so drawCellBorders can pick table.top/bottom/left/right vs.
        // table.insideH/insideV. `leftCol`/`rightCol` are the LOGICAL columns
        // (gridSpan-aware); the renderer flips them for bidiVisual via `mirror`.
        const edges: CellEdgeFlags = {
          topRow: ri === 0,
          bottomRow: lastRowOfCell === table.rows.length - 1,
          leftCol: ci === 0,
          rightCol: ci + span === colWidths.length,
        };
        // ECMA-376 §17.4.81: an exact row height is honored verbatim and
        // content taller than the row is clipped to the row box (Word clips;
        // we would otherwise overflow into neighboring rows). A vMerge=restart
        // cell spans multiple rows, so it is never governed by a single row's
        // exact height — only single-row cells clip.
        const clipExact = row.rowHeightRule === 'exact' && cell.vMerge !== true;
        if (dryRun) measureCellContent(cell, table, cellW, scale, state);
        else jobs.push({ cell, x: leadX, y, w: cellW, h: drawH, edges, clipExact });
      }

      x += cellW;
      ci += span;
    }

    y += rowH;
  }

  // Pass 1: backgrounds + content. Pass 2: borders, so a border is never
  // overpainted by a neighbouring cell's fill. `mirror` only swaps which
  // physical side a logical border maps to, so it is consulted in the border
  // pass alone.
  for (const j of jobs) {
    renderCell(j.cell, table, j.x, j.y, j.w, j.h, state, j.clipExact);
  }
  for (const j of jobs) {
    drawCellBorders(
      state.ctx, j.x, j.y, j.w, j.h, j.cell.borders, table.borders,
      scale, j.edges, mirror, state.dpr,
    );
  }
  return y;
}

/**
 * Render a FLOATING table (ECMA-376 §17.4.57 `<w:tblpPr>`). Like a `<w:framePr>`
 * frame, it is OUT OF FLOW: drawn at an absolute (anchor-relative) position and
 * consuming ZERO flow height, so the following content begins where the table's
 * anchor paragraph sat. A wrap-exclusion FloatRect is registered (§17.4.57
 * *FromText padding, §17.4.56 overlap) so the body text flows around it.
 *
 * Mirrors {@link renderFrameParagraph}: save contentX/contentW/y, redirect the
 * flow geometry to the resolved box, draw the rows, then RESTORE the in-flow
 * state.y (the float adds no flow height). In dryRun the rows are not drawn but
 * the float is still registered so wrap estimates see the band.
 */
function renderFloatTable(table: DocTable, state: RenderState): void {
  const tp = table.tblpPr!;
  const inFlowY = state.y;
  const savedX = state.contentX;
  const savedW = state.contentW;

  // Lay the table out in its anchor column's content band, then place its box.
  // tableW is the ACTUAL rendered width (sum of column widths), so the FloatRect
  // exclusion matches the painted table exactly (#513 column integrity: for
  // horzAnchor="text" the box.x derives from the column band via
  // frameXContainer, so the wrap stays inside this column).
  const { colWidths, tableW, rowHeights } = computeTableLayout(table, state.contentW, state);
  const tableH = rowHeights.reduce((s, h) => s + h, 0);
  const box = computeFloatTableBox(tp, state, inFlowY, tableW, tableH);
  const side = floatTableWrapSide(box, state);

  // Redirect the flow geometry to the float box and draw the rows there. The
  // table is positioned absolutely; its grid origin is box.x (no jc — a floating
  // table's position is dictated entirely by tblpPr, §17.4.57).
  state.contentX = box.x;
  state.contentW = tableW;
  drawTableRows(table, colWidths, tableW, rowHeights, box.x, box.y, state);
  state.contentX = savedX;
  state.contentW = savedW;

  // Restore the in-flow cursor: a floating table consumes NO body flow height
  // (§17.4.57 — out of flow), so the following content spaces as if it weren't
  // here. (renderFrameParagraph does the same for a frame.)
  state.y = inFlowY;

  registerTableFloat(box, tp, state, side, table.overlap !== 'never');
}

function renderTable(table: DocTable, state: RenderState): void {
  // ECMA-376 §17.4.57: a `<w:tblpPr>` table floats — divert to the out-of-flow
  // path before the normal block layout (which would advance state.y).
  if (table.tblpPr) {
    renderFloatTable(table, state);
    return;
  }

  const { contentX, contentW } = state;
  const { colWidths, tableW, rowHeights } = computeTableLayout(table, contentW, state);

  // Horizontal table alignment on the page (w:tblPr/w:jc).
  const tableX =
    table.jc === 'center'
      ? contentX + Math.max(0, (contentW - tableW) / 2)
      : table.jc === 'right'
        ? contentX + Math.max(0, contentW - tableW)
        : contentX;

  const y = drawTableRows(table, colWidths, tableW, rowHeights, tableX, state.y, state);

  state.y = y;
}

/** Height (px) of a single table row via the shared ST_HeightRule skeleton
 *  ({@link resolveSingleRowHeight}), with the paint pass's px cell measurer.
 *  EXCLUDES the §17.4.85 vMerge span extension — `computeTableLayout` applies
 *  that across the whole table. Exported for unit tests (table-row-height.test). */
export function calculateRowHeight(
  row: DocTableRow,
  table: DocTable,
  colWidths: number[],
  scale: number,
  state: RenderState,
): number {
  return resolveSingleRowHeight(row, colWidths, scale, (cell, cellW) =>
    measureCellContentHeightPx(cell, table, cellW, scale, state),
  );
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
    return lineBoxHeight(para.lineSpacing, asc, desc, scale, grid, paraHasRuby, emptyIntendedSinglePx(para, scale), paragraphIsEastAsian(para));
  }
  // ECMA-376 §17.3.1.12 (`<w:ind>`): the paragraph's own left/right indent
  // narrows the wrap width and `firstLine` insets the first line — exactly as
  // the paint pass (renderParagraph) and the paginator (estimateParagraphHeight)
  // lay it out. The row-height measurer MUST honor them too: without the indent,
  // a cell paragraph that carries a first-line/left indent (e.g. sample-11's
  // table cells, firstLine=21.6 pt) is measured for fewer wrapped lines than it
  // paints, so the row is sized too short and the overflow ("Town" /
  // "University") bleeds into the next row. `maxWidth` is the cell's inner width
  // (cell margins already removed by the caller); the paragraph indents come off
  // it here. `tabOriginPx = indentLeft` mirrors layoutLines' tab origin in the
  // paint/paginate paths.
  //
  // NOTE: like `estimateParagraphHeight` (the paginator), this passes the raw
  // `indentFirst` and does NOT model a numbering marker's `numBodyOffset` (the
  // hanging-indent first-line geometry `renderParagraph` applies for a numbered
  // paragraph). The two non-paint measurers therefore stay consistent with each
  // other, but a NUMBERED paragraph inside a cell that wraps can still measure
  // slightly differently from the paint pass. No such cell exists in the covered
  // samples; revisit together with estimateParagraphHeight if list-in-cell
  // fidelity is needed.
  const indLeftPx = para.indentLeft * scale;
  const indRightPx = para.indentRight * scale;
  const paraW = Math.max(1, maxWidth - indLeftPx - indRightPx);
  const lines = layoutLines(state.ctx, segs, paraW, para.indentFirst * scale, scale, para.tabStops, undefined, state.fontFamilyClasses, indLeftPx, state.kinsoku, gridCharDeltaPx(grid, scale), state.defaultTabPt);
  return lines.reduce((s, l) => s + lineBoxHeight(para.lineSpacing, l.ascent, l.descent, scale, grid, paraHasRuby, l.intendedSingle, paragraphIsEastAsian(para)), 0);
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
  clipExact = false,
): void {
  const { ctx, scale } = state;

  // Cell BACKGROUND + content only. Borders are painted in a separate, later
  // pass by drawTableRows (ECMA-376 §17.4.66 border-collapse): a shared gridline
  // must sit on top of every cell fill, so no neighbouring cell's background can
  // occlude the border drawn by the cell on the other side of the gridline.
  if (cell.background) {
    ctx.fillStyle = `#${cell.background}`;
    ctx.fillRect(x, y, w, h);
  }

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
    // ECMA-376 §17.3.1.33 + §17.4.84 (vAlign): Word collapses the FIRST
    // paragraph's space-before and the LAST paragraph's space-after against the
    // cell's content boundary when vertically aligning. Neither produces any ink
    // (nothing surrounds them inside the cell), so including them in the
    // vertically-aligned block height pushes the visible block off centre/bottom.
    // Word vertically aligns the INKED block alone: a header cell whose only
    // paragraph carries 6 pt space-before + 8 pt space-after still centres the
    // ~16.8 pt line box, not 30.8 pt. The symmetric trim mirrors how block
    // spacing collapses at a container edge (§17.3.1.33 describes spacing
    // BETWEEN paragraphs, not at the frame boundary). Spacing BETWEEN two
    // paragraphs inside the cell is left intact (handled by §17.3.1.33's
    // contextual / max-overlap rules inside the paint pass).
    const firstEl = visibleContent[0];
    const lastEl = visibleContent[visibleContent.length - 1];
    // Leading space-before (first paragraph only). Nested table first ⇒ 0.
    const firstSpaceBefore =
      firstEl && firstEl.type === 'paragraph'
        ? (firstEl as unknown as DocParagraph).spaceBefore * scale
        : 0;
    contentH -= firstSpaceBefore;
    // Trailing space-after (last paragraph only). Nested table last ⇒ 0.
    if (lastEl && lastEl.type === 'paragraph') {
      contentH -= (lastEl as unknown as DocParagraph).spaceAfter * scale;
    }
    // `renderParagraph` will re-consume the first paragraph's spaceBefore (it
    // unconditionally adds `para.spaceBefore * scale` to `state.y`). Pull
    // `cellState.y` up by `firstSpaceBefore` so that addition lands the inked
    // top exactly on the vertically-aligned position. Without this pull-up the
    // visible block lands `firstSpaceBefore` PAST the intended vAlign position
    // (= +3 pt down for a typical 6 pt spaceBefore at scale 1) — asymmetric with
    // the trailing-spaceAfter trim, which renderParagraph never reconsumes
    // because nothing follows it inside the cell.
    if (cell.vAlign === 'center') {
      cellState.y = y + (h - contentH) / 2 - firstSpaceBefore;
    } else {
      cellState.y = y + h - contentH - mb - firstSpaceBefore;
    }
  }

  if (clipExact) {
    // ECMA-376 §17.4.80 (trHeight) + §17.18.37 (ST_HeightRule "exact"):
    // the row height is exactly @val and content taller than that must not
    // bleed into adjacent rows. The clip is therefore **Y-axis only** — the
    // spec puts no horizontal bound on cell content. Clipping the full
    // (x, y, w, h) bbox half-masks a 0.5 pt nested-table border that lands
    // exactly on the cell's left/right edge (e.g. outer tcMar.left=0 +
    // tblCellMar.left=0 + inner tblInd=0): half the stroke straddles the clip
    // boundary and visibly disappears. Clipping by Y alone preserves the
    // anti-bleed intent without erasing borders that legitimately sit on the
    // cell edge.
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, y, ctx.canvas.width, h);
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
      // §17.3.1.9 contextualSpacing: between two same-style paragraphs that both
      // set it, BOTH the previous after and this before are dropped (gap = 0), not
      // just collapsed — so e.g. a code listing's lines sit tight.
      const overlap = suppress ? prevSpaceAfter : Math.min(prevSpaceAfter, effBefore);
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

/** Which grid edges of the table this cell touches, so {@link drawCellBorders}
 *  can pick the OUTER (table.top/bottom/left/right) vs the INNER
 *  (table.insideH/insideV) spec per physical edge (ECMA-376 §17.4.38/§17.4.39). */
interface CellEdgeFlags {
  topRow: boolean;     // cell sits in the table's first row → its top is the table outer top
  bottomRow: boolean;  // cell's bottom edge is the table outer bottom (vMerge-span aware)
  leftCol: boolean;    // cell's left edge is the table outer left
  rightCol: boolean;   // cell's right edge is the table outer right
}

/** Resolve a `nil`/`none` border to "no ink". A `null` means "not set" — the
 *  caller already substituted a fallback before reaching here. */
function paintable(b: BorderSpec | null): BorderSpec | null {
  if (!b) return null;
  if (b.style === 'none' || b.style === 'nil') return null;
  return b;
}

function drawCellBorders(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
  cell: CellBorders,
  table: TableBorders,
  scale: number,
  edges: CellEdgeFlags,
  mirror = false,
  dpr = 1,
): void {
  // ECMA-376 §17.4.38/§17.4.39: a cell's TOP/BOTTOM use the table OUTER border
  // (top/bottom) only on the table's outermost rows; an interior horizontal
  // gridline uses table.insideH. Likewise LEFT/RIGHT use the outer left/right
  // only on the outermost columns, interior verticals use table.insideV.
  // Per-edge precedence (§17.4.38/§17.4.39, the tblBorders presence cascade):
  // the cell's own explicit edge (which already folds the conditional
  // firstRow/lastRow/band tcBorders at parse time) wins; for an interior edge
  // the cell's insideH/insideV is consulted before the table inside spec; the
  // table outer spec is the last resort on outer edges.
  //
  // TODO(border-conflict §17.4.66): true adjacent-cell border conflict
  // resolution is NOT implemented. Per §17.4.66 (tcBorders), when cell spacing
  // is zero and the two cells sharing an interior gridline disagree, Word picks
  // the winner by border WEIGHT, not by sz/width:
  //   weight = (number of lines in the border) × (border style's number)
  // where "border style's number" is the spec's CT_Border style rank table
  // (single=1, thick=2, double=3, dotted=4, dashed=5, …, inset=25). The larger
  // weight wins. Ties are broken, in order:
  //   1. a style precedence list (single > thick > double > dotted > … > inset);
  //   2. luminance, darker wins, via three formulas applied in sequence —
  //      R+B+2G, then B+2G, then G (smaller value wins);
  //   3. reading order: the first border in reading order is displayed.
  // (There is no "top-then-left ownership" rule — that was a CSS-ism, not spec.)
  // Today each cell paints its own four edges, so an interior gridline is drawn
  // once by each of the two adjacent cells; for matching specs this is visually
  // identical, and `nil` on either side still suppresses that side's stroke.
  // When the two sides disagree the later-painted cell currently wins rather
  // than the §17.4.66 weight rule above.
  const horiz = (own: BorderSpec | null, outer: boolean): BorderSpec | null =>
    paintable(own ?? (outer ? table.top : (cell.insideH ?? table.insideH)));
  const horizB = (own: BorderSpec | null, outer: boolean): BorderSpec | null =>
    paintable(own ?? (outer ? table.bottom : (cell.insideH ?? table.insideH)));
  const vert = (own: BorderSpec | null, outer: boolean): BorderSpec | null =>
    paintable(own ?? (outer ? table.left : (cell.insideV ?? table.insideV)));
  const vertR = (own: BorderSpec | null, outer: boolean): BorderSpec | null =>
    paintable(own ?? (outer ? table.right : (cell.insideV ?? table.insideV)));

  const top = horiz(cell.top, edges.topRow);
  const bottom = horizB(cell.bottom, edges.bottomRow);
  // ECMA-376 §17.4.1: under bidiVisual the columns are visually reversed, so a
  // cell's logical left (start) border is drawn on its physical right edge and
  // vice versa. Borders are owned by the cell, so swap which spec each side uses,
  // AND swap which outer edge / inside-fallback applies (the start column's outer
  // edge is the physical right under bidiVisual).
  const left = mirror
    ? vertR(cell.right, edges.rightCol)
    : vert(cell.left, edges.leftCol);
  const right = mirror
    ? vert(cell.left, edges.leftCol)
    : vertR(cell.right, edges.rightCol);

  if (top) drawBorderLine(ctx, x, y, x + w, y, top, scale, dpr);
  if (bottom) drawBorderLine(ctx, x, y + h, x + w, y + h, bottom, scale, dpr);
  if (left) drawBorderLine(ctx, x, y, x, y + h, left, scale, dpr);
  if (right) drawBorderLine(ctx, x + w, y, x + w, y + h, right, scale, dpr);
}

/** Stroke one crisp axis-aligned segment. `perp` shifts the whole line
 *  perpendicular to its direction (px, pre-crisp-snap) — used to place the two
 *  rails of a `double` border on either side of the nominal edge. */
function strokeCrispSegment(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  x1: number, y1: number, x2: number, y2: number,
  lw: number,
  dpr: number,
  perp: number,
): void {
  ctx.lineWidth = lw;
  // Crispness nudge (see crispOffset): a thin (odd device-width) axis-aligned
  // stroke straddles two device rows and blurs; nudging it perpendicular to the
  // line snaps it onto the nearest crisp device position. Cell / paragraph
  // borders are always horizontal (y1===y2) or vertical (x1===x2) — never
  // diagonal — so the orientation is read directly from the endpoints, and the
  // snap delta is derived from the line's own coordinate (fractional-safe).
  const horizontal = y1 === y2;
  const vertical = x1 === x2;
  // `perp` runs along x for a horizontal line, along y for a vertical line.
  const ox = (horizontal ? 0 : perp);
  const oy = (horizontal ? perp : 0);
  const dpx = vertical ? crispOffset(x1 + ox, lw, dpr) : 0;
  const dpy = horizontal ? crispOffset(y1 + oy, lw, dpr) : 0;
  ctx.beginPath();
  ctx.moveTo(x1 + ox + dpx, y1 + oy + dpy);
  ctx.lineTo(x2 + ox + dpx, y2 + oy + dpy);
  ctx.stroke();
}

/**
 * ECMA-376 §17.18.2 ST_Border dash/dot families → a `setLineDash` pattern,
 * expressed in units of the stroked width `lw` (px). Thin wrapper over core's
 * shared `docxBorderDashArray` (which owns the §17.18.2 relative table); the ctx
 * is already `scale(dpr,dpr)`d, so `lw`-relative lengths render crisply at any
 * dpr (matching the single/double paths). Returns `[]` for solid styles.
 * Re-exported here so the existing `border-dash.test.ts` contract is preserved.
 */
export function borderDashPattern(style: string, lw: number): number[] {
  return docxBorderDashArray(style, lw);
}

function drawBorderLine(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  x1: number, y1: number, x2: number, y2: number,
  spec: BorderSpec,
  scale: number,
  dpr = 1,
): void {
  ctx.save();
  ctx.strokeStyle = spec.color ? `#${spec.color}` : '#000000';
  const lw = Math.max(0.5, spec.width * scale);

  if (spec.style === 'double') {
    // ECMA-376 §17.18.2 ST_Border "double": two parallel lines with a gap,
    // painted as device-pixel-aligned rail/gap/rail fills so a thin double
    // (e.g. sz6 ≈ 0.75px) never collapses into one line. Shared with the other
    // renderers via core's `fillDoubleBorder` (see core/draw/double-border.ts).
    ctx.fillStyle = ctx.strokeStyle;
    fillDoubleBorder(ctx, x1, y1, x2, y2, lw, dpr);
    ctx.restore();
    return;
  }

  // Dashed/dotted ST_Border families (§17.18.2). setLineDash is reset by the
  // ctx.restore() below. Solid styles get an empty pattern → continuous line.
  const dash = borderDashPattern(spec.style, lw);
  if (dash.length) ctx.setLineDash(dash);
  strokeCrispSegment(ctx, x1, y1, x2, y2, lw, dpr, 0);
  ctx.restore();
}

/**
 * ECMA-376 §17.3.1.7 — paragraph-border merge context for a run of consecutive
 * identically-bordered paragraphs. Word draws ONE box around such a run:
 *   - the `top` edge only on the FIRST paragraph,
 *   - the `bottom` edge only on the LAST paragraph,
 *   - the `<w:between>` edge (if any) at every INNER join,
 *   - `left`/`right` always (they form the box sides).
 * The paint loops (renderBodyElements / renderParaList) detect adjacency and
 * pass `suppressTop` when a same-border paragraph precedes this one, and
 * `suppressBottom` when one follows. When `suppressTop` is set the `between`
 * edge (if defined) is drawn at the top join instead of the `top` edge.
 */
interface ParaBorderMerge {
  /** A same-border paragraph is adjacent above ⇒ don't draw this `top` edge
   *  (draw `between` at the top join instead, when defined). */
  suppressTop?: boolean;
  /** A same-border paragraph is adjacent below ⇒ don't draw this `bottom` edge. */
  suppressBottom?: boolean;
}

/** The exact PAINTED height (px) of the lines a {@link renderParagraph} draw pass
 *  puts on this page, computed by replaying the per-line advancement the draw loop
 *  performs — WITHOUT drawing. The shading rect must match the paragraph border's
 *  height, and the border height is `state.y − textAreaTopY` measured AFTER the
 *  loop; but shading is the BACKGROUND and must be filled BEFORE the loop (text
 *  paints on top). So we cannot read the post-loop `state.y` for the fill — we
 *  re-derive it from the same inputs the loop uses.
 *
 *  The loop, for each line `li` in `[sliceStart, paintEnd)`, does exactly:
 *    if (line.topY !== undefined && line.topY > y) y = line.topY;  // float clearance
 *    y += lineHForLine(line);                                       // line box advance
 *  starting from `y = textAreaTopY`. Replaying it here yields H === the loop's final
 *  `state.y − textAreaTopY` BY CONSTRUCTION (same height source), so the shading
 *  meets the bottom border in every case:
 *   - normal (no float/slice): H === Σ lineHForLine over all lines (== the old naive
 *     `totalTextH`, so no regression);
 *   - float clearance: a line whose `topY` jumps past the natural flow grows H to
 *     match the border (previously the naive sum stopped short);
 *   - page-sliced paragraph: only `[sliceStart, paintEnd)` is summed, so H no longer
 *     overfills to the full-paragraph height past the slice's bottom border.
 *  `lineHForLine` is the paragraph-scope resolver (ruby/docGrid/lineSpacing) the
 *  loop already uses; passing it as a callback keeps this pure and testable. */
export function paintedParagraphHeight<L extends { topY?: number }>(
  lines: readonly L[],
  sliceStart: number,
  paintEnd: number,
  textAreaTopY: number,
  lineHForLine: (line: L) => number,
): number {
  let y = textAreaTopY;
  for (let li = sliceStart; li < paintEnd; li++) {
    const line = lines[li];
    if (line.topY !== undefined && line.topY > y) y = line.topY;
    y += lineHForLine(line);
  }
  return y - textAreaTopY;
}

/** ECMA-376 §17.3.1.31 — paragraph shading fills the border BOX, not just the
 *  text extent. §17.3.1.31 itself only says the shading sets the paragraph's
 *  background color and is SILENT on border geometry; the fill-to-border is
 *  observed Word behavior: Word fills the border box, and §17.3.1.7 places each
 *  border's `w:space` OUTSIDE the text box (applied by {@link drawParaBorders}), so
 *  the shading reaches those borders. Return the content box grown by each PRESENT
 *  border's space, using the SAME per-edge conditions as drawParaBorders so the
 *  fill meets the border exactly. Without a bordered edge (or no borders at all)
 *  that edge is not extended. (sample-11: a right border with `space=4` left the
 *  gray box detached from its border because the fill stopped `space` short of it.)
 *  Exported for unit testing the per-edge extension. */
export function paraShadingRect(
  x: number, y: number, w: number, h: number,
  borders: ParagraphBorders | null | undefined,
  merge: ParaBorderMerge | undefined,
  scale: number,
): { x: number; y: number; w: number; h: number } {
  if (!borders) return { x, y, w, h };
  const sp = (edge: ParaBorderEdge | null): number =>
    edge && edge.style !== 'none' ? (edge.space ?? 0) * scale : 0;
  const topEdge = merge?.suppressTop ? borders.between : borders.top;
  const l = sp(borders.left);
  const r = sp(borders.right);
  const t = sp(topEdge);
  const b = merge?.suppressBottom ? 0 : sp(borders.bottom);
  return { x: x - l, y: y - t, w: w + l + r, h: h + t + b };
}

function drawParaBorders(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
  borders: ParagraphBorders,
  scale: number,
  dpr = 1,
  merge?: ParaBorderMerge,
): void {
  const drawEdge = (edge: ParaBorderEdge | null, x1: number, y1: number, x2: number, y2: number) => {
    if (!edge || edge.style === 'none') return;
    const spec: BorderSpec = { width: edge.width, color: edge.color, style: edge.style };
    drawBorderLine(ctx, x1, y1, x2, y2, spec, scale, dpr);
  };
  const sp = (edge: ParaBorderEdge | null) => (edge?.space ?? 0) * scale;
  // §17.3.1.7 top edge: on a non-first paragraph of a shared run, the `top` edge
  // gives way to the `between` edge drawn at the join (nothing when `between` is
  // absent — the box has no internal rules).
  const topEdge = merge?.suppressTop ? borders.between : borders.top;
  drawEdge(topEdge, x, y - sp(topEdge), x + w, y - sp(topEdge));
  // The `bottom` edge is skipped entirely when a same-border paragraph follows
  // (the box continues into it; its own join is handled by that paragraph's
  // suppressed-top `between`).
  if (!merge?.suppressBottom) {
    drawEdge(borders.bottom, x, y + h + sp(borders.bottom), x + w, y + h + sp(borders.bottom));
  }
  drawEdge(borders.left,   x - sp(borders.left), y,        x - sp(borders.left), y + h);
  drawEdge(borders.right,  x + w + sp(borders.right), y,   x + w + sp(borders.right), y + h);
}

// ===== Utilities =====

/** ECMA-376 §17.3.1.7 — two paragraph-border definitions "match" (and so
 *  consecutive paragraphs carrying them merge into a single bordered box) iff
 *  ALL FIVE edges (top/bottom/left/right/between) are pairwise identical in
 *  style, color, size, and space. A `null`/absent edge equals another absent
 *  edge but differs from any present edge. Two paragraphs with NO borders are
 *  not a bordered run (we never merge unbordered paragraphs), so the caller
 *  gates on "both have a non-empty borders object" before calling this.
 *
 *  A `none`-style edge paints nothing, so for matching it is EQUIVALENT to an
 *  absent edge: an explicitly-cleared edge (`<w:bottom w:val="nil"/>`, parsed as a
 *  present `style:"none"` edge so it can override an inherited one — nil/none are
 *  normalized to "none" by parse_edge) and an omitted edge have the same effective
 *  border, and §17.3.1.7 compares effective borders. Normalize both to `null`
 *  before comparing — matching `drawEdge`/`hasAnyBorderEdge`, which also key on
 *  `style === 'none'`. */
function sameParaEdge(a: ParaBorderEdge | null, b: ParaBorderEdge | null): boolean {
  const eff = (e: ParaBorderEdge | null): ParaBorderEdge | null =>
    e == null || e.style === 'none' ? null : e;
  const ea = eff(a);
  const eb = eff(b);
  if (ea == null || eb == null) return ea == null && eb == null;
  return (
    ea.style === eb.style &&
    ea.width === eb.width &&
    (ea.space ?? 0) === (eb.space ?? 0) &&
    (ea.color ?? null) === (eb.color ?? null)
  );
}

function sameParaBorders(
  a: ParagraphBorders | null | undefined,
  b: ParagraphBorders | null | undefined,
): boolean {
  if (a == null || b == null) return false; // unbordered paragraphs never merge
  return (
    sameParaEdge(a.top, b.top) &&
    sameParaEdge(a.bottom, b.bottom) &&
    sameParaEdge(a.left, b.left) &&
    sameParaEdge(a.right, b.right) &&
    sameParaEdge(a.between, b.between)
  );
}

/** True when `borders` defines at least one visible edge (so a paragraph
 *  carrying it actually paints a box). A borders object whose every edge is
 *  null/`none` is treated as "no border" for merge purposes. */
function hasAnyBorderEdge(b: ParagraphBorders | null | undefined): boolean {
  if (!b) return false;
  const live = (e: ParaBorderEdge | null) => e != null && e.style !== 'none';
  return live(b.top) || live(b.bottom) || live(b.left) || live(b.right) || live(b.between);
}

/** ECMA-376 §17.3.1.7 — two paragraphs form (part of) the same bordered box iff
 *  both carry a visible paragraph border AND their five border edges match
 *  exactly. The caller is responsible for the ADJACENCY half of the rule (same
 *  column/flow, no page/column break or non-paragraph element between); this
 *  helper covers only the "identical border definition" half. A `framePr`
 *  (out-of-flow §17.3.1.11) paragraph can never share a run. */
function parasShareBorderBox(a: DocParagraph | null, b: DocParagraph | null): boolean {
  if (!a || !b) return false;
  if (a.framePr || b.framePr) return false;
  if (!hasAnyBorderEdge(a.borders) || !hasAnyBorderEdge(b.borders)) return false;
  return sameParaBorders(a.borders, b.borders);
}

/** ECMA-376 §17.3.2.4 — two `<w:bdr>` borders belong to the same run-border
 *  group iff their attribute sets are identical. We compare the attributes the
 *  model carries (style/sz/space/color); themeColor/themeTint/shadow/frame are
 *  not modelled, so identical themed borders that differ only in unmodelled
 *  attributes still group (acceptable — the painted frame is identical anyway). */
function runBordersEqual(a: DocxRunBorder, b: DocxRunBorder): boolean {
  return (
    a.style === b.style &&
    a.width === b.width &&
    (a.space ?? 0) === (b.space ?? 0) &&
    (a.color ?? null) === (b.color ?? null)
  );
}

function calcEffectiveFontPx(s: LayoutTextSeg, scale: number): number {
  // ECMA-376 §17.3.2.33: small-caps small letters render "in a font size TWO
  // POINTS SMALLER than the actual font size" (subtractive, not a ratio), floored
  // at the smallest renderable size when 2pt smaller is not possible. Applied in
  // pt before scaling. (At ~10pt this is ≈0.8×, but it diverges at larger sizes —
  // a 20pt heading's caps are 18pt, not 16pt.)
  const pt = s.smallCaps ? Math.max(s.fontSize - 2, 1) : s.fontSize;
  let size = pt * scale;
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

/** Resolve the list-marker glyph's font family (ECMA-376 §17.3.2.26 + §17.9.6).
 *  The marker is drawn/measured as a single `fillText`/`measureText`, so it must
 *  be one family. Pick it per the marker's leading code point, exactly like the
 *  body's per-character split and {@link shapeTokenFamily}: a CJK marker (e.g. an
 *  ideographic bullet) → the eastAsia axis, anything else (a decimal "1", roman
 *  "i", letter, or "•") → the ascii axis. Realistic markers are single-script, so
 *  the leading code point classifies the whole glyph string. eastAsia falls back
 *  to ascii when absent (older parser output / no eastAsia font). The font CLASS
 *  (serif/sans) is then resolved by `fontFamilyClasses` (fontTable §17.8.3.10),
 *  so e.g. a serif ascii (Times) number renders serif even when the heading's
 *  eastAsia axis is a Gothic (sans). */
function markerFontFamily(num: NumberingInfo): string | null {
  const cp = num.text.codePointAt(0) ?? 0;
  const ascii = num.fontFamily ?? null;
  return isCjkBreakChar(cp) ? (num.fontFamilyEastAsia ?? ascii) : ascii;
}

/** Marker glyph as it should be drawn/measured. Symbol/Wingdings markers
 *  (§17.9.x `w:lvlText` + §17.3.2.26 `w:rFonts`) store the glyph as the FONT's
 *  own code point (e.g. Symbol U+F0B7 = "•", Wingdings U+F0A7 = "▪"). Those
 *  private-encoding code points render as tofu in any fallback face, so we
 *  normalize them to the Unicode equivalent up front — keyed on the marker's
 *  requested ascii family, not on the sample. Non-symbol markers (decimals,
 *  roman, CJK bullets) pass through unchanged. */
function markerDisplayText(num: NumberingInfo): string {
  return symbolFontToUnicode(num.text, num.fontFamily ?? null);
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
      : // JP / stray-CJK sans faces: historical system-font hints, then the Noto
        // CJK siblings so a CJK glyph still resolves on hosts lacking them.
        ['Noto Sans JP', 'Hiragino Sans', 'Meiryo', ...cjkFallbackChain('jp', 'sans').slice(1)];
  // A Latin (non-CJK) sans font must fall back to a LATIN sans for its
  // letters/digits — otherwise the browser grabs them from a Japanese Gothic
  // (wider, CJK-tuned Latin), widening Latin runs. Lead with Latin sans faces;
  // the CJK gothic faces follow for any stray CJK glyph. (Mirrors serifTail.)
  if (cjk == null) {
    return `${quoteAll([...NON_CJK_SANS_FALLBACKS, 'Arial', 'Helvetica', 'Liberation Sans', ...cjkPart, ...ARABIC_TAIL_SANS])}, sans-serif`;
  }
  return `${quoteAll([...cjkPart, ...ARABIC_TAIL_SANS, ...NON_CJK_SANS_FALLBACKS])}, sans-serif`;
}

/** Serif counterpart of {@link sansTail}. */
function serifTail(cjk: ReturnType<typeof classifyCjkFont>): string {
  const cjkPart =
    cjk && cjk !== 'jp'
      ? cjkFallbackChain(cjk, 'serif')
      : // JP / stray-CJK serif faces: historical mincho system hints, then Noto
        // serif CJK siblings.
        [
          'Yu Mincho', 'YuMincho', 'Hiragino Mincho ProN', 'MS Mincho',
          'Noto Serif JP', ...cjkFallbackChain('jp', 'serif').slice(1),
        ];
  // A Latin (non-CJK) serif font (e.g. Century) must fall back to a LATIN serif
  // for its letters/digits. If the CJK mincho faces lead, the browser's
  // per-glyph fallback grabs Latin glyphs from a Japanese Mincho (e.g. Hiragino
  // Mincho ProN on macOS) whose Latin is ~15-18% wider, widening every Latin
  // run and forcing spurious line wraps. Lead with Latin serif faces; the CJK
  // mincho faces follow so a stray CJK glyph in a Latin-font run still resolves.
  if (cjk == null) {
    return `${quoteAll([...NON_CJK_SERIF_FALLBACKS, 'Times New Roman', 'Cambria', 'Liberation Serif', ...cjkPart, ...ARABIC_TAIL_SANS])}, serif`;
  }
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
/**
 * Per-document memo for {@link normalizeFontFamily}. The regex/classifier work
 * inside is a pure function of `(family, fontFamilyClasses)`, and
 * `fontFamilyClasses` is a stable per-document object (RenderState threads
 * `doc.fontFamilyClasses` — one identity per render). Keying the outer WeakMap on
 * that object identity gives per-doc caching with zero call-site churn (both
 * callers already pass `fontFamilyClasses`) and no leak: the inner
 * family→result Map is collected with the document's classes object. Same idiom
 * as `sheetAxisCache` / `mathRenders`. Chosen over threading an explicit cache
 * param through buildFont because both call sites already carry the classes
 * object, so identity-keying needs no signature changes anywhere.
 */
const fontFamilyNormalizeCache = new WeakMap<Record<string, string>, Map<string, string>>();

export function normalizeFontFamily(
  family: string | null,
  fontFamilyClasses: Record<string, string> = {},
): string {
  const perDoc =
    fontFamilyNormalizeCache.get(fontFamilyClasses) ??
    (() => {
      const m = new Map<string, string>();
      fontFamilyNormalizeCache.set(fontFamilyClasses, m);
      return m;
    })();
  // `family` may be null; use a distinct sentinel key so a null lookup never
  // collides with a real family named "null".
  const key = family ?? '\0null';
  const cached = perDoc.get(key);
  if (cached !== undefined) return cached;
  const result = normalizeFontFamilyUncached(family, fontFamilyClasses);
  perDoc.set(key, result);
  return result;
}

function normalizeFontFamilyUncached(
  family: string | null,
  fontFamilyClasses: Record<string, string>,
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

  // 2) Name-pattern fallback for fonts absent from fontTable or classified
  //    "auto". The serif/sans/mono DECISION is the shared core classifier
  //    (`classifyFontGeneric`, §17.8.3.10-aligned name heuristic) that pptx and
  //    xlsx also route through — so all three renderers agree on the generic
  //    class. docx keeps its own richer fallback-chain construction (Latin-first
  //    ordering + per-language CJK chains + Arabic tail + JP system hints) below;
  //    only the regex-based decision is delegated here. Core's serif token set
  //    is a verified superset of docx's former serif tokens (it additionally
  //    detects e.g. Century/Palatino/Didot as serif and Consolas/Courier/等幅 as
  //    mono on the name path), so no prior serif/sans coverage is lost.
  const generic = classifyFontGeneric(family);
  if (generic === 'serif') {
    return `${head}, ${serifTail(cjk)}`;
  }
  if (generic === 'mono') {
    // Mirror the fontTable `modern` branch's monospace fallback. NEW for the
    // name path: core now detects consolas/courier/等幅 etc. as mono.
    return `${head}, "Courier New", monospace`;
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

/** ECMA-376 §17.9.20 picture-bullet marker size in pt. The size comes from the
 *  `<w:numPicBullet>` drawing's own extent (parsed into
 *  `picBulletWidthPt`/`picBulletHeightPt`). The spec defines NO fallback
 *  dimension, so when the extent is absent the marker is sized to the paragraph's
 *  resolved marker font size — one source of truth shared by the collect side
 *  (WMF raster sharpness) and the draw site (the drawImage box), keeping them in
 *  lock-step. Replaces the former mismatched `?? 0` (collect) / `?? 9` (draw)
 *  defaults; `9` was a magic pt value not present in §17.9.20. */
function picBulletSizePt(num: NumberingInfo, para: DocParagraph): { w: number; h: number } {
  const fallback = getDefaultFontSize(para);
  return {
    w: num.picBulletWidthPt ?? fallback,
    h: num.picBulletHeightPt ?? fallback,
  };
}

/** Width (px) of the paragraph-mark "em" — the smallest horizontal gap a
 *  float-bordered empty / anchor-only paragraph-mark line is allowed to sit in
 *  before it flows below the float band. Single source of truth shared by the
 *  empty-mark float-window probe (resolveEmptyMarkTop) and the layoutLines
 *  empty-line fallback, so the two paths agree on what "one em of the mark
 *  font" means. HEURISTIC threshold (see resolveEmptyMarkTop), not spec. */
function paragraphMarkEmPx(para: DocParagraph, scale: number): number {
  return getDefaultFontSize(para) * scale;
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
  /** ECMA-376 §17.6.5 `<w:docGrid w:charSpace>` divided by 4096 — the per-EA-
   *  glyph character-grid delta = charSpace/4096 in FLAT POINTS (independent of
   *  font size), added to the measured glyph advance (≈1em for full-width EA
   *  glyphs). Negative tightens. `null`/`undefined` when the section declares no
   *  charSpace; the character grid is then inactive even if `type` is
   *  linesAndChars/snapToChars. See {@link gridCharDeltaPx}. */
  charSpacePt?: number | null;
}

// ───────────────────────────────────────────────────────────────────────────
// ECMA-376 §17.6.5 docGrid CHARACTER grid (字詰め). When the section's docGrid
// `type` is "linesAndChars" or "snapToChars" AND a `charSpace` is declared,
// every full-width East-Asian glyph gains a fixed per-EA-glyph spacing delta
//   Δpt = charSpace / 4096   in FLAT POINTS (NEGATIVE = tighter)
// that is INDEPENDENT of font size — it is added to the glyph's MEASURED advance
// (≈1em for full-width EA glyphs), NOT scaled by it. (`gridCharDeltaPx` returns
// exactly `charSpacePt * scale` = charSpace/4096 pt in px; it does not multiply
// by the font size.) Latin / digits are NOT snapped (they keep their natural
// advance), so the grid delta applies only to EA code points.
//
// ── The single advance model (measure == draw) ──────────────────────────────
// To make line-break MEASUREMENT and the draw ADVANCE provably identical, the
// grid delta enters in exactly ONE way: as a per-code-point spacing on a
// PURE-EA segment. `gridSegDeltaPx` returns the total delta a segment's box
// gains (`len × Δpx` for a pure-EA segment, else 0 — mixed/Latin segments get
// no grid effect, sidestepping any contextual-metric or justification drift),
// and `gridWidth` adds it to the natural `measureText` width. BOTH the layout's
// `measuredWidth` and every draw path derive the segment's advance from this
// SAME quantity:
//   • non-justified draw walks the glyphs via `justifiedPiecePositions(cps,
//     [1..n-1], perGap=0, measure, letterSpacingPx=Δ)`, whose final glyph lands
//     at `measure(whole) + n·Δ` = the box edge;
//   • justified draw reuses the EXISTING `justifiedPiecePositions` path with the
//     same `letterSpacingPx = Δ`, so its box edge is `measure(whole) + n·Δ +
//     nGaps·perGap` = `measuredWidth + internalStretch`.
// Because both come from `measure(prefix) + (cps before)·Δ`, draw never diverges
// from `measuredWidth` by construction — there is no separate per-glyph sum to
// drift against the whole-string measure (約物半角 contextual collapse stays
// honoured). See packages/core/src/text/justify-positions.ts.

/** Per-EA-glyph character-grid delta in px for a paragraph's grid, or 0 when the
 *  CHARACTER grid is inactive. Active only for docGrid type ∈ {linesAndChars,
 *  snapToChars} with a declared charSpace (ECMA-376 §17.6.5). The line grid
 *  ("lines") and a missing charSpace leave EA glyphs at natural advance. */
function gridCharDeltaPx(grid: DocGridCtx | undefined, scale: number): number {
  if (!grid || grid.charSpacePt == null) return 0;
  if (grid.type !== 'linesAndChars' && grid.type !== 'snapToChars') return 0;
  return grid.charSpacePt * scale;
}

/** Count of East-Asian (full-width) code points in `text` — the glyphs the
 *  character grid snaps to cells. Uses the same {@link EAST_ASIAN_RE} content
 *  predicate as docGrid line-cell rounding (no font-name heuristic). */
function eaGlyphCount(text: string): number {
  let n = 0;
  for (const ch of text) if (EAST_ASIAN_RE.test(ch)) n++;
  return n;
}

/** The total character-grid delta (px) a segment's advance gains under an active
 *  character grid. Applied ONLY to a PURE East-Asian segment (every code point
 *  is EA): then `len × deltaPx` cells the whole run, and a uniform per-cp
 *  letter-spacing of `deltaPx` reproduces it exactly on both draw paths. A mixed
 *  or pure-Latin segment returns 0 — Latin is never snapped (§17.6.5), and
 *  skipping mixed segments avoids the per-cp-vs-whole-string and justification
 *  drift that would break measure==draw. `deltaPx===0` (grid inactive) ⇒ 0. */
function gridSegDeltaPx(text: string, deltaPx: number): number {
  if (deltaPx === 0 || text.length === 0) return 0;
  const cps = [...text];
  return eaGlyphCount(text) === cps.length ? cps.length * deltaPx : 0;
}

/** Single source of truth for a text segment's laid-out advance: the natural
 *  `measureText` width plus the character-grid delta (0 unless an active grid
 *  AND a pure-EA segment). EVERY line-break / advance measurement and every draw
 *  path must derive the segment advance from this, so they cannot diverge. */
function gridWidth(naturalWidthPx: number, text: string, deltaPx: number): number {
  return naturalWidthPx + gridSegDeltaPx(text, deltaPx);
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
  eastAsian = false,
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
  // A single-spaced line on a docGrid snaps to whole grid CELLS in East Asian
  // text: Word rounds its height UP to the next pitch multiple, so a line taller
  // than one pitch reserves two (or more). sample-9 (pdftotext -bbox of the Word
  // PDF): a 20pt CJK title on a 20pt pitch is 40px = 2 cells; an 11pt body is
  // 20px = 1 cell. A Latin-only line is NOT cell-rounded — it keeps its natural
  // height above a one-cell floor (demo/sample-1: an 18pt heading on an 18pt
  // pitch stays ~20.7px, not 36). ECMA-376 Part 1 only defines the natural ≤
  // pitch case (§17.6.5 / §17.3.1.32); the East-Asian cell rounding for taller
  // lines is Word runtime behaviour, so it is gated on the line's script.
  const gridSingleCell = (): number => eastAsian
    ? Math.max(pitchPx, Math.ceil(glyphNatural / pitchPx) * pitchPx)
    : Math.max(glyphNatural, pitchPx);
  const inheritedOnly = ls !== null && ls.explicit !== true;
  if (!ls) {
    // No explicit spacing → single line. Use the intended single-line height
    // (`natural`) off-grid; on-grid, snap per gridSingleCell.
    return hasGrid ? gridSingleCell() : natural;
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
    return hasGrid ? gridSingleCell() : natural;
  }
  if (ls.rule === 'auto') {
    if (hasGrid) {
      if (inheritedOnly) return gridSingleCell();
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

/** Corrected single-line ascent/descent (px) from an ALREADY-measured
 *  `TextMetrics`: the Canvas `fontBoundingBox` (with the synthetic 0.8/0.2-em
 *  fallback when the engine reports none), rescaled to the document font's design
 *  line box via {@link correctLineMetrics}. The single source of truth for "how
 *  tall is one line of `family`", shared by the text-line path (layoutLines) and
 *  the empty paragraph-mark path (paragraphMarkLineHeight) so the two cannot
 *  drift — that drift (the empty path skipping `correctLineMetrics`) was the
 *  empty-paragraph under-measure bug (§17.3.1.29 / §17.3.1.33). `fallbackEmPx`
 *  sizes the synthetic box (the run's full size); `correctionEmPx` is the design
 *  size handed to `correctLineMetrics` — they differ only for smallCaps/vertAlign
 *  runs (where the text path keeps the full-size fallback) and coincide for a
 *  plain paragraph-mark line. The hhea single-line FLOOR for tabled fonts is
 *  applied separately by lineBoxHeight via {@link intendedSingleLinePx}. */
function correctedLineMetrics(
  m: TextMetrics,
  family: string | null | undefined,
  fallbackEmPx: number,
  correctionEmPx: number,
): { ascent: number; descent: number } {
  const rawAsc = m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent ?? fallbackEmPx * 0.8;
  const rawDesc = m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent ?? fallbackEmPx * 0.2;
  return correctLineMetrics(family, correctionEmPx, rawAsc, rawDesc);
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
  eastAsian = false,
  ctx?: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  fontFamilyClasses: Record<string, string> = {},
): number {
  const fs = getDefaultFontSize(para);
  const family = getDefaultFontFamily(para);
  let asc: number;
  let desc: number;
  if (ctx) {
    // ECMA-376 §17.3.1.29 / §17.3.1.33: an empty paragraph's mark line reserves
    // the mark font's REAL single-line height — the SAME fontBoundingBox a text
    // line of that font and size uses (layoutLines), so an empty paragraph is
    // exactly as tall as a one-character paragraph of the same run properties.
    // The synthetic 0.8/0.2 ≈ 1em box under-measured every empty paragraph
    // whenever the (often substituted) font's real box exceeds 1em — a Latin
    // fallback reports ~1.15em — so a run of empty "spacer" paragraphs fell
    // short and the following content rose into a preceding float's wrap band
    // (sample-12: the figure caption wrapped beside the image instead of below
    // it). East Asian documents probe an EA glyph so docGrid cell rounding
    // (lineBoxHeight) reserves whole cells (a 20pt mark on a 20pt pitch → 2
    // cells); others probe a Latin glyph. fontBoundingBox is reported per
    // resolved face (not per glyph), so the probe choice does not change the box
    // for a face that contains it — and the probe is script-matched, so the mark
    // font does. correctedLineMetrics rescales a substituted font to the document
    // font's design box, identical to the text path; the hhea single-line floor
    // (intendedSingleLinePx, via emptyIntendedSinglePx below) then raises tabled
    // fonts — Latin included — to Word's line height.
    const prevFont = ctx.font;
    ctx.font = buildFont(false, false, fs * scale, family, fontFamilyClasses);
    const m = ctx.measureText(eastAsian ? 'あ' : 'x');
    ctx.font = prevFont;
    // A mark line carries no smallCaps/vertAlign, so fallback == correction size.
    ({ ascent: asc, descent: desc } = correctedLineMetrics(m, family, fs * scale, fs * scale));
  } else {
    ({ asc, desc } = emptyLineNaturalPx(fs, scale));
  }
  return lineBoxHeight(para.lineSpacing, asc, desc, scale, grid, paraHasRuby, emptyIntendedSinglePx(para, scale), eastAsian);
}
