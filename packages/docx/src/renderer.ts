import type {
  DocxDocumentModel, BodyElement, PaginatedBodyElement, DocParagraph, DocTable, DocTableRow, DocTableCell, CellElement,
  DocRun, DocxTextRun, ImageRun, ShapeRun, ShapeTextRun, FieldRun, HeaderFooter, LineSpacing, BorderSpec, TableBorders, CellBorders,
  TabStop, ParagraphBorders, ParaBorderEdge, DocxRunBorder, SectionProps, DocNote, NumberingInfo, ColumnGeom, FramePr, TblpPr,
} from './types';
import type { ArrowEnd, Stroke } from '@silurus/ooxml-core';
import {
  buildCustomPath,
  buildShapePath,
  renderPresetShape,
  hasPreset,
  getCachedSvgImageByPath,
  hexToRgba,
  autoContrastColor,
  resolveFill,
  applyStroke,
  drawArrowHead,
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
  decodeRasterOrMetafile,
  symbolFontToUnicode,
  isSymbolFontFamily,
  symbolTextToUnicodeSegments,
  docxBorderDashArray,
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
}

/** Returns a stable map key for an (imagePath, colorReplaceFrom) pair. */
function imageKey(imagePath: string, colorReplaceFrom?: string): string {
  return colorReplaceFrom ? `${imagePath}|clr:${colorReplaceFrom}` : imagePath;
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
          widthPt: img.widthPt ?? 0,
          heightPt: img.heightPt ?? 0,
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
 * The raster/metafile path delegates to the shared
 * {@link decodeRasterOrMetafile} (the one decoder docx/pptx/xlsx now share):
 * browsers can't decode WMF/EMF via `createImageBitmap`, so it content-sniffs
 * the bytes first (extension/MIME are unreliable — sample-10's chart is a
 * standard WMF mislabeled `.emf`), rasterizing a WMF via the minimal player at a
 * size derived from `widthPt`/`heightPt`, returning `null` for a true EMF (or a
 * geometry-less metafile), else `createImageBitmap`. A `null` result throws so
 * `preloadImages` drops the image (the existing "missing image" behavior, no
 * crash).
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
  const blob = await fetchImage(imagePath, mimeType);
  const bmp = await decodeRasterOrMetafile(blob, {
    widthPt,
    heightPt,
    suppressBoundaryFrame: true,
  });
  if (!bmp) throw new Error(`${imagePath} produced no drawable output`);
  return colorReplaceFrom ? applyColorReplacement(bmp, colorReplaceFrom) : bmp;
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
      try {
        let img: DecodedImage;
        if (pair.svgImagePath != null) {
          // Prefer the vector original (Microsoft `asvg:svgBlip` extension);
          // fall back to the raster on any SVG decode failure. docx images have
          // no srcRect crop, so the vector is always preferred when present.
          try {
            img = await getCachedSvgImageByPath(pair.svgImagePath, fetch);
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
  const tokenHost = canvas as unknown as { __docxRenderToken?: number };
  const myToken = (tokenHost.__docxRenderToken = (tokenHost.__docxRenderToken ?? 0) + 1);
  const superseded = () => tokenHost.__docxRenderToken !== myToken;

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
  const baseState: RenderState = {
    ctx,
    scale,
    dpr,
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
    floatParaSeq: 0,
    docGrid: {
      type: sec.docGridType ?? null,
      linePitchPt: sec.docGridLinePitch ?? null,
      charSpacePt: sec.docGridCharSpace != null ? sec.docGridCharSpace / 4096 : null,
    },
    docEastAsian: docEA,
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

  // Body. ECMA-376 §17.6.4: lay out body text in EACH section's newspaper columns
  // (per-section columns). `columns` is the body-level (final) section's geometry,
  // used as the fallback for elements that carry no per-section `colGeom` (single-
  // section docs, where it equals the whole-body geometry — unchanged path).
  const columns = computeColumns(sec);
  const bodyState: RenderState = { ...baseState, y: sec.marginTop * scale };
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
  renderBodyElements(elements, bodyState, columns);

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
  docEastAsian: boolean,
): { total: number; trailingSpaceAfter: number } {
  const measure = buildMeasureState(ctx, sec, fontFamilyClasses, kinsoku, docEastAsian);
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
    const m = measureNoteBlockForDraw(note, baseState.ctx, sec, baseState.fontFamilyClasses, baseState.kinsoku, baseState.docEastAsian);
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
  const maxY = cssHeight - sec.marginBottom * scale;
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

export function computePages(
  body: BodyElement[],
  section: SectionProps,
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  fontFamilyClasses: Record<string, string> = {},
  kinsoku: KinsokuRules = DEFAULT_KINSOKU_RULES,
  footnotes: DocNote[] = [],
): PaginatedBodyElement[][] {
  const fullContentH = section.pageHeight - section.marginTop - section.marginBottom;
  const measureState = buildMeasureState(ctx, section, fontFamilyClasses, kinsoku, documentHasEastAsian(body));
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
        // computeColumns reads only `section.columns` + the page geometry (which
        // is constant across the body), so resolve this section's cols by
        // swapping in its ColumnsSpec.
        return computeColumns({ ...section, columns: e.columns ?? null });
      }
    }
    return computeColumns(section);
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
  measureState.y = section.marginTop;
  measureState.floats = [];
  measureState.floatParaSeq = 0;
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
      colIndex = 0;
      prevPara = null;
      prevSpaceAfter = 0;
      measureState.y = section.marginTop;
      // Floats are PAGE-scoped (ECMA-376 §20.4.2.x): a new page starts with a
      // clean float set. Columns of the SAME page share floats (see nextColumn).
      measureState.floats = [];
      measureState.floatParaSeq = 0;
      startPageBookkeeping();
    }
  };
  // Advance to the next newspaper column of the CURRENT page: reset the vertical
  // cursor to the column top but KEEP the page's floats (they are page-scoped, so
  // a full-width wrapTopAndBottom band still pushes down every column's first
  // line, and a square float keeps constraining the columns its x-range covers).
  // No new page is pushed and footnote reserve is untouched (same page).
  const nextColumn = () => {
    colIndex++;
    y = 0;
    prevPara = null;
    prevSpaceAfter = 0;
    measureState.y = section.marginTop;
  };
  // Overflow handler shared by element placement and paragraph/table splitting:
  // move to the next column if one remains on this page, otherwise to a new page.
  const nextColumnOrPage = () => {
    if (colIndex < columns.length - 1) nextColumn();
    else newPage();
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
      if (run.type === 'image') {
        const img = run as unknown as ImageRun;
        if (!img.anchor || !img.anchorYFromPara) continue;
        // Wrap floats anchor after spaceBefore (registerAnchorFloats uses
        // state.y post-spaceBefore); non-wrap floats anchor at the paragraph's
        // pre-spaceBefore top (renderAnchorImages uses paragraphStartY). Mirror
        // each so the estimate matches the draw position exactly.
        const anchorBase = isWrapFloat(img.wrapMode) ? spaceBeforePt : 0;
        const bottom = anchorBase + (img.anchorYPt ?? 0) + img.heightPt;
        if (bottom > maxBottom) maxBottom = bottom;
      } else if (run.type === 'shape') {
        // An anchored shape with positionV relativeFrom="paragraph"/"line" is
        // kept on its anchor's page the same way an image is. Mirror the image
        // formula exactly — bottom = anchorBase + anchorYPt + height — but take
        // the height from resolveShapeBox so sizeRelV / wgp-group scaling is
        // honored. measureState is scale 1 (pt), so box.h is already in pt.
        // (paragraphTopPx is passed through for shapes whose height depends on a
        // paragraph/line container via sizeRelV; it does not affect the height
        // for the common static-extent case.)
        const shp = run as unknown as ShapeRun;
        if (!shp.anchorYFromPara) continue;
        const anchorBase = isWrapFloat(shp.wrapMode) ? spaceBeforePt : 0;
        const box = resolveShapeBox(shp, measureState, measureState.y + anchorBase);
        if (box.h <= 0) continue;
        const bottom = anchorBase + (shp.anchorYPt ?? 0) + box.h;
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
    pages[pages.length - 1].push(el);
  };

  for (let i = 0; i < body.length; i++) {
    const el = body[i];
    if (el.type === 'columnBreak') {
      // ECMA-376 §17.3.1.20 <w:br w:type="column"/>: force the next column (or a
      // new page's first column when already in the last column — newPage() no-ops
      // on an empty page, so a column break in the last column of an as-yet-empty
      // page simply stays put).
      nextColumnOrPage();
      continue;
    }
    if (el.type === 'pageBreak') {
      pages.push([]);
      y = 0;
      prevPara = null;
      prevSpaceAfter = 0;
      measureState.y = section.marginTop;
      measureState.floats = [];
      measureState.floatParaSeq = 0;
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
      if (el.kind === 'continuous') {
        // ECMA-376 §17.18.79 "continuous": NO page break — the next section's
        // content continues on the SAME page at the CURRENT vertical position
        // (y is intentionally NOT reset), just in the new section's column 0.
        // (sample-5's continuous break is 1-col → 1-col, so the two single-column
        // sections simply stack; full continuous column-balancing for a column-
        // count change is out of scope per the task.) prevPara is cleared so the
        // first paragraph of the new section doesn't collapse spacing against the
        // last paragraph of the previous one. Floats / footnote reserve stay
        // (page-scoped). measureState.y already tracks the current y.
        prevPara = null;
        prevSpaceAfter = 0;
      } else {
        // nextPage (default) / oddPage / evenPage: start a new page (mirrors the
        // pageBreak path, including parity padding). A new page already resets
        // colIndex to 0 and clears page-scoped floats.
        pages.push([]);
        y = 0;
        prevPara = null;
        prevSpaceAfter = 0;
        measureState.y = section.marginTop;
        measureState.floats = [];
        measureState.floatParaSeq = 0;
        startPageBookkeeping();
        if (el.kind === 'oddPage' && pages.length % 2 === 0) {
          pages.push([]);
          startPageBookkeeping();
        } else if (el.kind === 'evenPage' && pages.length % 2 === 1) {
          pages.push([]);
          startPageBookkeeping();
        }
      }
      continue;
    }
    if (el.type === 'paragraph') {
      const para = el as unknown as DocParagraph;
      if (para.pageBreakBefore) newPage();

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
      // Snapshot the float set + paraSeq BEFORE this paragraph registers, so a
      // relocation to the NEXT COLUMN (which keeps page floats) can roll back this
      // paragraph's own floats and re-register them at the new column position
      // without leaving stale duplicates. (A relocation to a new PAGE clears
      // floats wholesale via newPage(), so this snapshot is unused there.)
      const floatsBefore = measureState.floats.length;
      const floatSeqBefore = measureState.floatParaSeq;
      const paragraphAnchorY = measureState.y + effectiveBefore;
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
      const floatBottomOff = anchoredFloatBottomOffset(para, effectiveBefore);
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
      if (breakForFloat || (overflowsHere && keepIntact && needed <= effContentH())) {
        const pagesBeforeRelocate = pages.length;
        nextColumnOrPage();
        const movedToNewPage = pages.length > pagesBeforeRelocate;
        if (movedToNewPage) {
          // newPage() cleared measureState.floats AND reset floatParaSeq to 0, so
          // this is a REPLACE of the earlier register at the top of the loop (whose
          // floats were just discarded): this paragraph is now the first registrant
          // on the fresh page and gets paraId 0 — matching the renderer, which
          // re-registers from a fresh per-page state.
          registerAnchorFloats(para, measureState, measureState.y + effectiveBefore);
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
          registerAnchorFloats(para, measureState, measureState.y + effectiveBefore);
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
      const remainingH = pageContentH - y;
      const splittable = !para.keepLines || h > pageContentH;
      if (fitHeight > remainingH && splittable) {
        const placed = splitParagraphAcrossPages(
          measureState, para, colW(), suppressBefore, colX(),
          y, pageContentH, pages,
          // Overflow during the split advances to the next column first, then a
          // new page (newspaper fill). Each slice is tagged with the column it
          // landed in via the colIndex thunk, plus this section's column geometry
          // (constant — a paragraph never spans a section boundary).
          () => { nextColumnOrPage(); },
          () => colIndex,
          columns,
        );
        // After splitting, `y` is the bottom of the last slice in the
        // current column (continues for the LAST slice; intermediate slices
        // filled their column/page exactly, so the break callback ran between
        // them).
        y = placed.endY;
        measureState.y = section.marginTop + placed.endY;
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
          () => { nextColumnOrPage(); },
          () => colIndex,
          columns,
        );
        y = endY;
        measureState.y = section.marginTop + endY;
      } else {
        if (y + h > tableContentH) nextColumnOrPage();
        pushTagged(el as PaginatedBodyElement);
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
  docEastAsian = false,
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
    marginTop: section.marginTop,
    marginBottom: section.marginBottom,
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
    const lines = layoutLines(state.ctx, segs, paraW, para.indentFirst, 1, para.tabStops, wrapCtx, state.fontFamilyClasses, indLeft, state.kinsoku, gridCharDeltaPx(grid, 1));
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
  newPage: () => void,
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
): { endY: number } {
  const stamp = (el: PaginatedBodyElement): PaginatedBodyElement => {
    if (tagColIndex) el.colIndex = tagColIndex();
    if (colGeom) el.colGeom = colGeom;
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
    if (initialY > 0 && initialY + markH - para.spaceAfter > contentH) {
      newPage();
      top = 0;
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
  const lines = layoutLines(measureState.ctx, segs, paraW, para.indentFirst, 1, para.tabStops, wrapCtx, measureState.fontFamilyClasses, indLeft, measureState.kinsoku, gridCharDeltaPx(paraGrid(para, measureState), 1));
  if (lines.length === 0) {
    // Anchor-only paragraph: no inline lines, but the paragraph mark still
    // occupies one (possibly relocated) line (§17.3.1.29).
    return placeMarkOnly();
  }
  const paraHasRuby = paragraphHasRuby(para);

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
        newPage();
        cursorY = 0;
        isFirstSliceOnPage = true;
        continue;
      }
    }
    const isFinalSlice = lastFitting === lines.length;
    if (isFinalSlice) usedH += spaceAfter;
    pages[pages.length - 1].push(stamp({
      ...(para as object),
      type: 'paragraph',
      lineSlice: { start: firstFitting, end: lastFitting },
    } as PaginatedBodyElement));
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
 *  ECMA-376 §17.4.80, §17.4.85) via the shared {@link resolveTableRowHeights}
 *  skeleton. Works in pt (scale 1); the cell measurer is the paginator's
 *  float-aware `estimateParagraphHeight` cursor-walk. */
function computeTableRowHeights(state: RenderState, table: DocTable, contentWPt: number): number[] {
  const colWidths = resolveColumnWidths(table, contentWPt, state);
  return resolveTableRowHeights(table, colWidths, 1, (cell, cellW) => {
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
    return ch;
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
  /** ECMA-376 §17.6.4 — current newspaper column index, read AFTER each
   *  `newPage()`. When provided, each table slice is tagged with its column.
   *  Omitted (single-column / direct unit tests) ⇒ no tag. */
  tagColIndex?: () => number,
  /** ECMA-376 §17.6.4 — the current SECTION's column geometry (constant across
   *  the split; a table is never split across a section boundary). Stamped on
   *  each slice so the renderer resolves its column against the right section. */
  colGeom?: ColumnGeom[],
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
    const sliceEl = { ...table, type: 'table', rows: sliceRows } as PaginatedBodyElement;
    if (tagColIndex) sliceEl.colIndex = tagColIndex();
    if (colGeom) sliceEl.colGeom = colGeom;
    pages[pages.length - 1].push(sliceEl);

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
  const topY = sec.marginTop * scale;
  const botY = (sec.pageHeight - sec.marginBottom) * scale;
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
): void {
  let prevPara: DocParagraph | null = null;
  let prevSpaceAfter = 0;
  // The (geometry, column index) the flow `state` is currently set to. `activeCol`
  // starts at -1 so the first element always seeds the column. `activeGeom` tracks
  // the SECTION whose columns are in effect (per-section newspaper columns,
  // §17.6.4): elements carry their own section's geometry in `colGeom`, so two
  // sections sharing a page (a continuous break) each resolve their `colIndex`
  // against the right widths.
  let activeCol = -1;
  let activeGeom: ColumnGeom[] | undefined;
  for (const el of elements) {
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
        state.y = state.marginTop * state.scale;
      }
      prevPara = null;
      prevSpaceAfter = 0;
      activeCol = elCol;
      activeGeom = cols;
    } else if (!multiCol && cols && cols.length === 1 && cols !== activeGeom) {
      // A single-column SECTION on a multi-section page (e.g. sample-5 sections
      // 1–4): set the full-width content band for this section. Continue at the
      // current y (continuous) — only the page-level fresh state / a page break
      // resets to the top, both handled by the caller.
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
  },
): void {
  const { ctx, scale, dryRun } = state;
  const { grid, paraHasRuby, contentX, indLeft, paraW, textAreaTopY,
    paragraphStartY, markTop, totalLines, lineSlice } = markCtx;
  // Displacement applied by the float-flow (0 when the mark fits where it is).
  const flowShift = Math.max(0, markTop - textAreaTopY);
  if (markTop > state.y) state.y = markTop;
  const markRectTop = state.y;
  const emptyH = paragraphMarkLineHeight(para, scale, grid, paraHasRuby, state.docEastAsian, ctx, state.fontFamilyClasses);
  if (para.shading && !dryRun) {
    ctx.fillStyle = `#${para.shading}`;
    ctx.fillRect(contentX + indLeft, markRectTop, paraW, emptyH);
  }
  state.y += emptyH;
  if (para.borders && !dryRun) {
    drawParaBorders(ctx, contentX + indLeft, markRectTop, paraW, emptyH, para.borders, scale, state.dpr);
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
): void {
  const { ctx, scale, contentX, contentW, defaultColor, dryRun, fontFamilyClasses } = state;
  // Capture Y before spaceBefore — used for paragraph-relative anchor image positioning
  const paragraphStartY = state.y;

  if (!suppressSpaceBefore) state.y += para.spaceBefore * scale;

  // Register anchor floats from this paragraph (must happen after spaceBefore so
  // that paragraph-relative Y resolves against the textAreaTop, matching Word).
  // Skipped for the frame-draw recursion: a frame paragraph's wrap exclusion is
  // its own FloatRect (renderFrameParagraph), not an anchor image/shape float.
  if (!inFrame) registerAnchorFloats(para, state, state.y);

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
    if (suff !== 'tab') {
      // Body-offset width: the picture bullet's own width if present, else the
      // measured glyph width (§17.3.2.26 + §17.9.6) with the marker's RESOLVED
      // font — the width must match the draw below so the body offset is exact
      // for a serif (Times) vs sans (Gothic) marker alike.
      let markerW: number;
      if (picBullet) {
        markerW = picBullet.w;
      } else {
        ctx.font = buildFont(false, false, getDefaultFontSize(para) * scale, markerFontFamily(para.numbering), fontFamilyClasses);
        markerW = ctx.measureText(markerDisplayText(para.numbering)).width;
      }
      const spaceW = suff === 'space' ? ctx.measureText(' ').width : 0;
      // marker sits at firstLineX (= paraX + indFirst); body starts at its end.
      numBodyOffset = indFirst + markerW + spaceW;
    }
  }
  // True when the paragraph has any marker to draw (text glyph OR picture bullet).
  const hasMarker = numMarker !== '' || picBullet !== null;

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
      markTop: resolveEmptyMarkTop(), totalLines: 0, lineSlice: undefined,
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
  const lines = layoutLines(ctx, segments, paraW, firstLineIndent, scale, para.tabStops, wrapCtx, state.fontFamilyClasses, indLeft, state.kinsoku, gridCharDeltaPx(grid, scale));

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
      markTop: resolveEmptyMarkTop(), totalLines: lines.length, lineSlice,
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
  // ECMA-376 §17.6.5 character-grid delta (px per EA glyph) for the DRAW pass —
  // the SAME value layoutLines folded into measuredWidth. A pure-EA segment is
  // drawn so its glyphs occupy exactly `measuredWidth` (= natural + len·Δ): the
  // draw uses `justifiedPiecePositions(..., letterSpacingPx = Δ)`, whose final
  // glyph lands on the box edge, so the painted advance equals measuredWidth by
  // construction. See the gridCharDeltaPx / gridSegDeltaPx header.
  const drawGridDeltaPx = gridCharDeltaPx(grid, scale);
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
        const left = baseRtl ? x + lineWidth + numTab - w : lineLeft + indFirst;
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
          // float xOffset, so the marker tracks the body that hangs off it). The
          // body was advanced past the marker above (numBodyOffset).
          ctx.fillText(markerDisplayText(para.numbering!), lineLeft + indFirst, baseline);
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
        lastDrawnSi,
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
      // drawing. Decorations span the stretched width so a highlight / underline /
      // ruby covers the widened glyphs. trailingGap (the inter-segment boundary)
      // is added after the segment below.
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

        // Width spanned by the glyphs after justification, for box-style
        // decorations (highlight) and ruby centring / onTextRun reporting.
        const spanW = s.measuredWidth + internalStretch;

        // Glyph box used by every run-level box decoration (highlight fill,
        // §17.3.2.32 shading fill, §17.3.2.4 border): same vertical extent of
        // ~0.85em above the baseline to ~0.25em below it. Computed once so the
        // three decorations stay byte-identical (no duplicated 0.85 / 1.1).
        const boxTop = baseline + yOffset - effSizePx * 0.85;
        const boxHeight = effSizePx * 1.1;

        if (s.highlight) {
          ctx.fillStyle = HIGHLIGHT_COLORS[s.highlight] ?? '#FFFF00';
          ctx.fillRect(x, boxTop, spanW, boxHeight);
        }

        // ECMA-376 §17.3.2.32 run shading fill (`<w:shd w:fill>`): a solid
        // background rect behind the glyphs. Used for inverse video (black fill
        // + automatic = white text). Same rect geometry as the highlight box.
        if (s.background) {
          ctx.fillStyle = `#${s.background}`;
          ctx.fillRect(x, boxTop, spanW, boxHeight);
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
            borderGroup.right = x + spanW;
            borderGroup.top = Math.min(borderGroup.top, segTop);
            borderGroup.bottom = Math.max(borderGroup.bottom, segBottom);
          } else {
            flushBorderGroup();
            borderGroup = {
              border: activeBorder,
              left: x, right: x + spanW, top: segTop, bottom: segBottom,
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
          const justGaps = stretch?.splitBefore ?? [];
          let g = 0; // justification gaps strictly before the current glyph
          for (let i = 0; i < cps.length; i++) {
            while (g < justGaps.length && justGaps[g] <= i) g++;
            const prefix = cps.slice(0, i).join('');
            const dx = ctx.measureText(prefix).width + i * drawGridDeltaPx + g * distPerGap;
            ctx.fillText(cps[i], x + dx, baseline + yOffset);
          }
        } else if (stretch && stretch.splitBefore.length > 0) {
          // ECMA-376 §17.18.44 `both`/`distribute` inter-CJK justification pitch.
          // Anchor each sliced piece to the WHOLE-string cumulative advance plus
          // the accumulated pitch, instead of summing the isolated pieces'
          // advances. That sum drifts wider than the segment's box and would paint
          // the next run over this segment's tail (most visible at a CJK→Latin
          // boundary). See `@silurus/ooxml-core` → text/justify-positions.ts.
          const cps = [...s.text]; // code points (handles surrogate pairs)
          const measure = (str: string): number => ctx.measureText(str).width;
          for (const { text: piece, dx } of justifiedPiecePositions(
            cps,
            stretch.splitBefore,
            distPerGap,
            measure,
          )) {
            ctx.fillText(piece, x + dx, baseline + yOffset);
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
        // Underline / strike run the full stretched glyph span. Use the
        // segment's box (measuredWidth, which already folds in the §17.6.5
        // character-grid delta) plus the internal justification pitch, so the
        // decoration matches the drawn advance instead of re-measuring the
        // natural width (which would ignore the grid on a packed EA run).
        const textW = s.measuredWidth + internalStretch;

        const isInsertion = revActive && s.revision?.kind === 'insertion';
        const isDeletion = revActive && s.revision?.kind === 'deletion';

        if (s.underline || isInsertion) {
          ctx.strokeStyle = lineColor;
          ctx.lineWidth = lineW;
          const uyRaw = baseline + yOffset + effSizePx * 0.12;
          const uy = uyRaw + crispOffset(uyRaw, lineW, state.dpr);
          ctx.beginPath(); ctx.moveTo(x, uy); ctx.lineTo(x + textW, uy); ctx.stroke();
        }

        if (s.strikethrough || isDeletion) {
          ctx.strokeStyle = lineColor;
          ctx.lineWidth = lineW;
          const syRaw = baseline + yOffset - effSizePx * 0.3;
          const sy = syRaw + crispOffset(syRaw, lineW, state.dpr);
          ctx.beginPath(); ctx.moveTo(x, sy); ctx.lineTo(x + textW, sy); ctx.stroke();
        }

        if (s.doubleStrikethrough) {
          ctx.strokeStyle = lineColor;
          ctx.lineWidth = lineW;
          const sy1Raw = baseline + yOffset - effSizePx * 0.35;
          const sy2Raw = baseline + yOffset - effSizePx * 0.22;
          const sy1 = sy1Raw + crispOffset(sy1Raw, lineW, state.dpr);
          const sy2 = sy2Raw + crispOffset(sy2Raw, lineW, state.dpr);
          ctx.beginPath(); ctx.moveTo(x, sy1); ctx.lineTo(x + textW, sy1); ctx.stroke();
          ctx.beginPath(); ctx.moveTo(x, sy2); ctx.lineTo(x + textW, sy2); ctx.stroke();
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

  if (para.borders && !dryRun) {
    const textH = state.y - textAreaTopY;
    drawParaBorders(ctx, contentX + indLeft, textAreaTopY, paraW, textH, para.borders, scale, state.dpr);
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
        smallCaps: base.smallCaps ?? false,
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
  // ECMA-376 §17.6.5 docGrid CHARACTER grid: per-EA-glyph cell delta in px (0
  // when inactive). Folded into every advance via `gridWidth` so line breaking
  // packs the grid's char count per line; the draw paths add the SAME delta.
  gridDeltaPx = 0,
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
      const intended = intendedSingleLinePx(ts.fontFamily, effectiveFontPx(ts));
      if (intended > lineIntendedSingle) lineIntendedSingle = intended;
    }
  };

  const effectiveFontPx = (s: LayoutTextSeg): number => calcEffectiveFontPx(s, scale);

  const measureText = (s: LayoutTextSeg): TextMetrics => {
    ctx.font = buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses);
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
    ctx.font = buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily, fontFamilyClasses);
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
    // Subtract the GRID width of the trimmed text (not the natural width) so the
    // grid delta on EA glyphs cancels and trailingSpaceW is the bare space
    // advance — keeping `w` and `wForFit` on the one advance model.
    const trailingSpaceW = s.text.endsWith(' ')
      ? w - gridWidth(ctx.measureText(trimmed).width, trimmed, gridDeltaPx)
      : 0;
    const wForFit = w - trailingSpaceW;
    const shrinkBudget = lineTotalTrailingW * SPACE_SHRINK_RATIO;

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
    state.ctx.drawImage(bmp, pageX, pageY, w, h);
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
): string[] {
  if (!text) return [''];
  if (maxWidth <= 0) return [text];
  const tokens = tokenizeShapeText(text);

  const lines: string[] = [];
  let cur = '';
  for (const tok of tokens) {
    if (cur !== '' && ctx.measureText(cur + tok).width > maxWidth) {
      lines.push(cur.replace(/\s+$/, ''));
      cur = tok.replace(/^\s+/, ''); // a wrapped line never starts with a space
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
  for (const tok of tokens) {
    if (cur.length > 0 && curW + tok.width > maxWidth) {
      lines.push(cur);
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

  // First pass: lay out each block. Text blocks WRAP to the inner width
  // (ECMA-376 §21.1.2.1.1) — a long title/abstract that exceeds the box width
  // breaks onto multiple lines instead of overflowing the page; image blocks
  // reserve their fitted height. The computed layout drives both vertical
  // anchoring (totalH) and the draw pass (no re-wrapping).
  type BlockLayout =
    | { kind: 'image'; fitW: number; fitH: number }
    | { kind: 'text'; lines: string[]; lineH: number }
    | { kind: 'rich'; lines: RichToken[][]; lineHeights: number[] };
  const layouts: BlockLayout[] = blocks.map((b) => {
    if (b.imagePath) {
      const { w: fitW, h: fitH } = fitShapeImage(b.imageWidthPt ?? 0, b.imageHeightPt ?? 0, innerW, scale);
      return { kind: 'image', fitW, fitH };
    }
    // Rich path: a paragraph with explicit per-run formatting lays out as mixed
    // fonts. Each line's height is the tallest run on it × 1.2 (ECMA-376 line
    // box ≈ largest font on the line).
    if (b.runs && b.runs.length > 0) {
      const lines = wrapShapeRuns(ctx, b.runs, innerW, scale, fontFamilyClasses);
      const lineHeights = lines.map((toks) => {
        const maxPt = toks.reduce((m, t) => Math.max(m, t.run.fontSizePt), 0);
        return (maxPt > 0 ? maxPt : b.fontSizePt) * scale * 1.2;
      });
      return { kind: 'rich', lines, lineHeights };
    }
    const fontPx = b.fontSizePt * scale;
    ctx.font = buildFont(b.bold ?? false, b.italic ?? false, fontPx, b.fontFamily ?? null, fontFamilyClasses);
    return { kind: 'text', lines: wrapShapeText(ctx, b.text, innerW), lineH: fontPx * 1.2 };
  });
  const blockHeight = (l: BlockLayout): number => {
    if (l.kind === 'image') return l.fitH;
    if (l.kind === 'rich') return l.lineHeights.reduce((s, h) => s + h, 0);
    return l.lines.length * l.lineH;
  };
  const totalH = layouts.reduce((s, l) => s + blockHeight(l), 0);

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
    const layout = layouts[i];

    if (layout.kind === 'image') {
      // Inline image inside the text box. Fit to inner width, place
      // horizontally per the paragraph alignment (figures default to centered),
      // and advance by the reserved height regardless of whether a bitmap is
      // present (a missing decode must not shift the rest of the layout).
      const { fitW, fitH } = layout;
      const bmp = block.imagePath ? images.get(imageKey(block.imagePath)) : undefined;
      if (bmp) {
        let drawX = innerX + Math.max(0, (innerW - fitW) / 2); // default: centered
        if (block.alignment === 'left' || block.alignment === 'both') {
          drawX = innerX;
        } else if (block.alignment === 'right') {
          drawX = innerX + Math.max(0, innerW - fitW);
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
      for (let li = 0; li < layout.lines.length; li++) {
        const lineToks = layout.lines[li];
        const lineH = layout.lineHeights[li];
        const lineW = lineToks.reduce((s, t) => s + t.width, 0);
        // Base direction (first-strong) from the line's own text, matching the
        // single-format path's per-block resolution but resolved per line.
        const lineText = lineToks.map((t) => t.text).join('');
        const baseRtl = resolveBaseDirection(undefined, lineText) === 'rtl';
        ctx.direction = baseRtl ? 'rtl' : 'ltr';
        const edge = edgeFor(baseRtl);
        let tx = innerX;
        if (edge === 'center') {
          tx = innerX + Math.max(0, (innerW - lineW) / 2);
        } else if (edge === 'right') {
          tx = innerX + Math.max(0, innerW - lineW);
        }
        // Baseline uses the tallest font on the line (lineH / 1.2 × 0.85).
        const lineMaxFontPx = lineH / 1.2;
        const baseline = cursorY + lineMaxFontPx * 0.85;
        for (const tok of lineToks) {
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
    for (const line of layout.lines) {
      const m = ctx.measureText(line);
      let tx = innerX;
      if (edge === 'center') {
        tx = innerX + Math.max(0, (innerW - m.width) / 2);
      } else if (edge === 'right') {
        tx = innerX + Math.max(0, innerW - m.width);
      }
      // Baseline = line top + ascent (approx 0.85 of font size for default fonts).
      const baseline = cursorY + fontPx * 0.85;
      ctx.fillText(line, tx, baseline);
      cursorY += layout.lineH;
    }
  }
  ctx.direction = 'ltr'; // reset for subsequent draws
}

/**
 * Resolve an anchor image's page-space box origin and dist* padding (px), shared
 * by registerAnchorFloats (wrap floats) and renderAnchorImages (wrapNone images).
 *
 * X: margin-relative offsets add section.marginLeft (ECMA-376 §20.4.3.4
 * relativeFrom="margin"); otherwise anchorXPt is already page-absolute.
 * Y: paragraph-relative offsets add `paraBaseY`; otherwise page-absolute. The
 * caller supplies `paraBaseY` because the two consumers anchor against different
 * paragraph references — wrap floats use the post-spaceBefore textAreaTop, while
 * wrapNone images use the pre-spaceBefore paragraph top (see the
 * anchoredFloatBottomOffset note in the paginator). This is the box origin BEFORE
 * any overlap displacement; resolveFloatOverlap runs on top of it for floats.
 */
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
  // box size as the align size. ImageRun carries neither a raw relativeFrom
  // string nor pctPos/sizeRel, so those args are null and the boolean
  // anchorXFromMargin / anchorYFromPara hints pick page-vs-margin containers
  // (xContainer/yContainer). When align is absent, resolveAnchorX/Y fall back
  // to the same offset path this function used previously.
  const x = resolveAnchorX(
    img.anchorXAlign, img.anchorXFromMargin ?? false, img.anchorXPt ?? 0, w, state,
    null, null, null,
  );
  const y = resolveAnchorY(
    img.anchorYAlign, img.anchorYFromPara ?? false, img.anchorYPt ?? 0, h, paraBaseY, state,
    null, null, null,
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

/** Register floats from a paragraph's anchor images and shapes. Anchor images
 *  are drawn immediately; anchor shapes are NOT drawn here (renderAnchorShape
 *  paints them separately) — we only reserve their float-exclusion band so body
 *  text wraps around them (ECMA-376 §20.4.2.16/.17), exactly like images. */
function registerAnchorFloats(para: DocParagraph, state: RenderState, paragraphAnchorY: number): void {
  // One id per registerAnchorFloats call ⇒ one id per paragraph. Floats sharing
  // a paraId (e.g. two side-by-side photos in one paragraph) never displace each
  // other; floats from different paragraphs do (de-facto overlap avoidance).
  const paraId = state.floatParaSeq++;
  for (const run of para.runs) {
    if (run.type === 'image') {
      registerImageFloat(run as unknown as ImageRun, state, paragraphAnchorY, paraId);
    } else if (run.type === 'shape') {
      registerShapeFloat(run as unknown as ShapeRun, state, paragraphAnchorY, paraId);
    }
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

  // Wrap floats anchor against the post-spaceBefore textAreaTop (paragraphAnchorY).
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
    if (bmp) state.ctx.drawImage(bmp, rect.imageX, rect.imageY, rect.imageW, rect.imageH);
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
  // wrap floats anchor against the post-spaceBefore textAreaTop, identical to
  // the image path (resolveAnchorBox uses paragraphAnchorY there).
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
 *  paint pass's `measureCellElementHeight`. Shared by the per-row skeleton (via
 *  computeTableLayout) and the exported {@link calculateRowHeight}. */
function measureCellContentHeightPx(
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
        if (!dryRun) renderCell(cell, table, leadX, y, cellW, drawH, state, edges, mirror, clipExact);
        else measureCellContent(cell, table, cellW, scale, state);
      }

      x += cellW;
      ci += span;
    }

    y += rowH;
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
  const lines = layoutLines(state.ctx, segs, maxWidth, 0, scale, para.tabStops, undefined, state.fontFamilyClasses, 0, state.kinsoku, gridCharDeltaPx(grid, scale));
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
  edges: CellEdgeFlags,
  mirror = false,
  clipExact = false,
): void {
  const { ctx, scale } = state;

  if (cell.background) {
    ctx.fillStyle = `#${cell.background}`;
    ctx.fillRect(x, y, w, h);
  }

  drawCellBorders(ctx, x, y, w, h, cell.borders, table.borders, scale, edges, mirror, state.dpr);

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
    // ECMA-376 §17.18.2 ST_Border "double": a double line around the object.
    // The standard does not normatively define the rail/gap geometry, so this
    // matches Word's rendering — three equal bands across the nominal width
    // (line / gap / line), each rail = sz/3, gap = sz/3, the pair centered on
    // the edge. (Word's own UI labels `sz` as the total double-border width.)
    const railW = Math.max(0.5, lw / 3);
    const offset = (lw - railW) / 2; // rail centers sit ±offset from the edge
    strokeCrispSegment(ctx, x1, y1, x2, y2, railW, dpr, -offset);
    strokeCrispSegment(ctx, x1, y1, x2, y2, railW, dpr, offset);
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

function drawParaBorders(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
  borders: ParagraphBorders,
  scale: number,
  dpr = 1,
): void {
  const drawEdge = (edge: ParaBorderEdge | null, x1: number, y1: number, x2: number, y2: number) => {
    if (!edge || edge.style === 'none') return;
    const spec: BorderSpec = { width: edge.width, color: edge.color, style: edge.style };
    drawBorderLine(ctx, x1, y1, x2, y2, spec, scale, dpr);
  };
  const sp = (edge: ParaBorderEdge | null) => (edge?.space ?? 0) * scale;
  drawEdge(borders.top,    x, y - sp(borders.top),         x + w, y - sp(borders.top));
  drawEdge(borders.bottom, x, y + h + sp(borders.bottom),  x + w, y + h + sp(borders.bottom));
  drawEdge(borders.left,   x - sp(borders.left), y,        x - sp(borders.left), y + h);
  drawEdge(borders.right,  x + w + sp(borders.right), y,   x + w + sp(borders.right), y + h);
}

// ===== Utilities =====

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
  let asc: number;
  let desc: number;
  if (eastAsian && ctx) {
    // East Asian document: the empty paragraph-mark line follows the real East
    // Asian font box so docGrid cell rounding (lineBoxHeight) reserves whole
    // cells — a 20pt mark on a 20pt pitch → 2 cells (40px), matching Word's title
    // pages. The synthetic 0.8/0.2 ≈ 1em fallback measures exactly one pitch and
    // would round down to a single cell. fontBoundingBox is font-wide, so any
    // East Asian glyph probes it.
    const prevFont = ctx.font;
    ctx.font = buildFont(false, false, fs * scale, getDefaultFontFamily(para), fontFamilyClasses);
    const m = ctx.measureText('あ');
    asc = m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent ?? fs * scale * 0.8;
    desc = m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent ?? fs * scale * 0.2;
    ctx.font = prevFont;
  } else {
    ({ asc, desc } = emptyLineNaturalPx(fs, scale));
  }
  return lineBoxHeight(para.lineSpacing, asc, desc, scale, grid, paraHasRuby, emptyIntendedSinglePx(para, scale), eastAsian);
}
