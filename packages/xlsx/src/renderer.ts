import type {
  Worksheet, Styles, Cell, CellValue, Font, Fill, Border, BorderEdge, CellXf,
  ViewportRange, RenderViewportOptions, TextRunInfo,
  CfRule, CellRange, CfStop, CfValue, Dxf, Hyperlink, DefinedName,
  Run, ChartData, GradientFillSpec, ShapeInfo, SlicerItem,
} from './types.js';
import { renderChart, renderSparkline, type ChartModel, type SparklineModel } from '@silurus/ooxml-core';
import { evalFormulaToBool, todaySerial, nowSerial } from './formula.js';
import { formatCellValue } from './number-format.js';

// Default font stack. Calibri is the workbook default font in Excel; on
// systems without Office (macOS / Linux) the browser would otherwise fall
// back to Arial / Helvetica, which is meaningfully wider than Calibri at
// every weight/size combination. Carlito is the Google-released, metric-
// compatible Calibri clone (same advance widths and ascender / descender
// metrics) and is loaded opt-in by `XlsxWorkbook.load({ useGoogleFonts:
// true })`. Listing it in the cascade means: Calibri (Windows / Office)
// → Carlito (loaded webfont) → Arial → sans-serif. Caladea is the same
// for Cambria.
const DEFAULT_FONT_FAMILY = '"Calibri", "Carlito", "Cambria", "Caladea", Arial, sans-serif';
const DEFAULT_FONT_SIZE = 11;
// Fallback Max Digit Width of the Normal-style font when the workbook's
// default font isn't known. Calibri 11 pt at 96 DPI ≈ 8 px (Canvas2D
// measurement), matching the EMU offsets Excel 365 writes into
// <xdr:twoCellAnchor>. ECMA-376 §18.3.1.13 defines MDW as the maximum
// rendered width among the digits 0-9 in the workbook's Normal-style font,
// so the spec-correct value depends on which font and point size that style
// resolves to (e.g. Meiryo UI 10 pt yields MDW ≈ 6 px).
const MDW_FALLBACK = 8;
/** Standard pt → CSS px conversion at 96 DPI. ECMA-376 §18.4.11 (font sz),
 * §18.8.5 (border width margins, etc.) all express their dimensions in
 * points. Multiply by this constant to obtain the display pixel value. */
const PT_TO_PX = 4 / 3;

export const HEADER_W = 50;
export const HEADER_H = 22;

// Thin line drawn between frozen and scrollable areas
const FREEZE_LINE_COLOR = '#7a7a7a';

/** Cache of Max Digit Width per "family:sizePt" key. The Canvas2D
 *  `measureText` call is cheap but not free; column-width conversion is
 *  invoked many times per render so we memoize. */
const mdwCache = new Map<string, number>();

/** Excel's empirical MDW overrides for fonts that the host might not have
 *  installed (e.g. Meiryo UI on macOS, where Canvas2D falls back to a
 *  narrower sans-serif and undermeasures the digits — verified against
 *  private/sample-10's column widths in Excel: 21.125 chars = 169 px,
 *  which requires MDW=8). Without these the rendered column widths drift
 *  smaller than Excel's, which then offsets every drawing anchor inside
 *  the sheet (sample-10's H7 sun ended up one cell to the right of where
 *  Excel renders it).
 *
 *  Only listed when the Canvas2D fallback measurement diverges from
 *  Excel's actual MDW; other fonts (e.g. Yu Gothic 12 pt where Canvas
 *  happens to land on 8) continue to use the measurement. */
const MDW_TABLE: Record<string, Record<number, number>> = {
  'meiryo ui':       { 10: 8, 11: 8 },
  'meiryo':          { 10: 8, 11: 8 },
};

/** Measure the Max Digit Width (ECMA-376 §18.3.1.13) for an arbitrary font
 *  using Canvas2D. The maximum of `measureText('0'..'9').width` is taken,
 *  rounded to the nearest pixel to match Excel's storage of integer pixel
 *  widths in `<col>` width values.
 *
 *  When the requested font isn't installed on the host (e.g. Meiryo UI on
 *  macOS), Canvas2D silently falls back to a narrower sans-serif face and
 *  the digit width comes back ~1 px too small. That offset cascades into
 *  the column widths, shifting every drawing anchor inside the sheet
 *  relative to where Excel placed it (private/sample-10 H7 sun ended up
 *  one cell to the right of where Excel renders it). So we consult a small
 *  lookup table of Excel's documented MDW values first and only fall back
 *  to measurement for unknown faces. */
export function computeMdw(family: string, sizePt: number): number {
  const key = `${family}:${sizePt}`;
  const cached = mdwCache.get(key);
  if (cached !== undefined) return cached;
  const tableHit = MDW_TABLE[family.toLowerCase()]?.[Math.round(sizePt)];
  if (tableHit !== undefined) {
    mdwCache.set(key, tableHit);
    return tableHit;
  }
  const sizePx = sizePt * PT_TO_PX;
  // Off-DOM canvas: avoids touching the document tree from background calls.
  const canvas = (typeof OffscreenCanvas !== 'undefined')
    ? new OffscreenCanvas(1, 1)
    : (typeof document !== 'undefined' ? document.createElement('canvas') : null);
  if (!canvas) return MDW_FALLBACK;
  const ctx = canvas.getContext('2d');
  if (!ctx) return MDW_FALLBACK;
  // Quote the family so multi-word names like "Meiryo UI" parse as one face.
  ctx.font = `${sizePx}px "${family}", ${DEFAULT_FONT_FAMILY}`;
  let mdw = 0;
  for (const d of '0123456789') {
    const w = ctx.measureText(d).width;
    if (w > mdw) mdw = w;
  }
  const out = Math.round(mdw) || MDW_FALLBACK;
  mdwCache.set(key, out);
  return out;
}

/** Resolve the Max Digit Width for a worksheet's Normal-style font. Falls
 *  back to the Calibri 11 pt baseline (~8 px) when the parser couldn't
 *  determine the workbook's default font. */
export function getMdwForWorksheet(ws: { defaultFontFamily?: string; defaultFontSize?: number }): number {
  if (!ws.defaultFontFamily || !ws.defaultFontSize) return MDW_FALLBACK;
  return computeMdw(ws.defaultFontFamily, ws.defaultFontSize);
}

export function colWidthToPx(w: number, mdw: number = MDW_FALLBACK): number {
  return Math.trunc(((256 * w + 128 / mdw) / 256) * mdw);
}

/** Convert a row height value from the parser into CSS pixels.
 *
 * ECMA-376 §18.3.1.73 (`<row ht>`) and §18.3.1.81 (`sheetFormatPr@defaultRowHeight`)
 * both specify the value in points. Convert pt → CSS px at 96 DPI (×4/3)
 * to match what Excel actually displays. The parser keeps both per-row
 * heights and the intrinsic default in points so this single conversion
 * applies to either source. */
export function rowHeightToPx(h: number): number {
  return Math.round(h * PT_TO_PX);
}

function hexToRgba(hex: string, alpha = 1): string {
  const h = hex.replace('#', '');
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return alpha === 1 ? `rgb(${r},${g},${b})` : `rgba(${r},${g},${b},${alpha})`;
}

/**
 * Fill a data-bar rectangle. Excel 2010+ dataBars default to a horizontal
 * gradient (`x14:dataBar@gradient="1"`): solid color on the left, fading
 * to an ~85%-tinted-to-white version on the right. We render with alpha
 * stops rather than literally mixing toward white so underlying cell
 * background (including zebra-striping or fills) shows through. With
 * `gradient="0"` the bar is drawn as a flat solid color.
 */
function fillDataBar(
  ctx: CanvasRenderingContext2D,
  color: string,
  x: number, y: number, w: number, h: number,
  gradient: boolean,
): void {
  if (w <= 0 || h <= 0) return;
  if (gradient) {
    const grad = ctx.createLinearGradient(x, y, x + w, y);
    grad.addColorStop(0, hexToRgba(color, 0.85));
    grad.addColorStop(1, hexToRgba(color, 0.15));
    ctx.fillStyle = grad;
  } else {
    ctx.fillStyle = hexToRgba(color);
  }
  ctx.fillRect(x, y, w, h);
}

/**
 * Fractional fg coverage for an ECMA-376 ST_PatternType (§18.8.22). Values are
 * derived from the spec's verbal descriptions — gray125 = 12.5% fg on bg,
 * gray0625 = 6.25%, mediumGray ≈ 50%, darkGray ≈ 75%, lightGray ≈ 25%. Hatch
 * variants (darkHorizontal etc.) approximate their average ink density.
 * Unknown values default to 1 so they render as solid fg.
 */
function patternCoverage(pt: string): number {
  switch (pt) {
    case 'solid':         return 1;
    case 'darkGray':      return 0.75;
    case 'mediumGray':    return 0.50;
    case 'lightGray':     return 0.25;
    case 'gray125':       return 0.125;
    case 'gray0625':      return 0.0625;
    // Directional hatches: visual ink ratio is roughly 50/50 at default cell
    // sizes; without a true hatch tile we blend at 50% so the cell reads as a
    // middle-tone fill rather than a solid fg block.
    case 'darkHorizontal':
    case 'darkVertical':
    case 'darkDown':
    case 'darkUp':
    case 'darkGrid':
    case 'darkTrellis':   return 0.5;
    case 'lightHorizontal':
    case 'lightVertical':
    case 'lightDown':
    case 'lightUp':
    case 'lightGrid':
    case 'lightTrellis':  return 0.25;
    default:              return 1;
  }
}

/**
 * Cache of (pattern, fg, bg, ctx scale) → CanvasPattern. Keyed by a compound
 * string so the same pattern across thousands of cells only builds the tile
 * once *for a given destination scale*. The ctx scale is part of the key
 * because pat.setTransform() pre-bakes the inverse-scale transform into the
 * pattern, and we don't want a Storybook scale=1.5 cell to reuse a cached
 * pattern from a previous scale=2 render.
 */
const PATTERN_CACHE = new Map<string, CanvasPattern | null>();

/**
 * Bitmaps for ECMA-376 ST_PatternType (§18.8.22). Each entry is a square tile
 * — the row count (rows.length) is the tile size in pixels, and bit
 * (size - 1) of each row is the leftmost pixel. A `1` bit places fg on the
 * tile, `0` leaves bg.
 *
 * Mixing tile sizes lets us pick the smallest square that exactly fits each
 * pattern's natural pitch: gray and diagonal families fit in 8×8, but the
 * horizontal / vertical / grid families need a 12×12 tile so dark and light
 * variants can share the same line pitch (3 px) while differing only in
 * stroke thickness (2 px vs 1 px). At 8×8 this would have forced a 4-px
 * pitch on dark vs 2-px on light — visibly different line counts instead of
 * the matched-count, different-thickness pair Excel renders.
 *
 * Values follow the long-standing Office pattern set; the spec only names
 * the patterns, never their pixel geometry.
 *
 * Coverage targets:
 *   gray125    12.5%   sparse dot  (1 dot per 8 pixels)
 *   gray0625    6.25%   very sparse (1 dot per 16 pixels)
 *   lightGray  25%      checker dot (1 in 4)
 *   mediumGray 50%      full checker
 *   darkGray   75%      inverse of lightGray
 */
const PATTERN_BITMAPS: Record<string, number[]> = {
  // ── 8×8 — gray-family dot patterns ────────────────────────────────────
  // Constructed as a single dot motif scaled up tier-by-tier — each pattern
  // is the previous one with extra dots interleaved at the same 4-row pitch.
  // gray0625 is the reference seed (4 dots / 64 ≈ 6%) and the cascade roughly
  // doubles density at each step, so all five tiers read as the same stipple
  // texture at progressively higher coverage. Some moiré at intermediate
  // zooms is acceptable per user preference; the visual continuity wins.
  gray0625:   [0b10000000, 0b00000000, 0b00001000, 0b00000000, 0b10000000, 0b00000000, 0b00001000, 0b00000000], // ≈ 6%
  gray125:    [0b10001000, 0b00000000, 0b00100010, 0b00000000, 0b10001000, 0b00000000, 0b00100010, 0b00000000], // ≈ 12%
  lightGray:  [0b10101010, 0b00000000, 0b01010101, 0b00000000, 0b10101010, 0b00000000, 0b01010101, 0b00000000], // ≈ 25%
  mediumGray: [0b10101010, 0b01010101, 0b10101010, 0b01010101, 0b10101010, 0b01010101, 0b10101010, 0b01010101], // ≈ 50%
  // 75% — distribute the empty pixels evenly across rows so the cell reads
  // as a solid darker grey instead of alternating bands.
  darkGray:   [0b01110111, 0b11011101, 0b01110111, 0b11011101, 0b01110111, 0b11011101, 0b01110111, 0b11011101], // ≈ 75%

  // ── 12×12 — horizontal / vertical (matched line count) ───────────────
  // dark* and light* share the same 4-line-per-tile count and 3-px pitch;
  // they differ only in stroke thickness — dark uses a 2-px bar and a
  // 1-px gap, light uses a 1-px line and a 2-px gap. Per user feedback
  // the visual rule is "B19 has the same number of horizontal lines as
  // B18, just thinner", so the matched-count construction is restored
  // here. The earlier "more lines in light" 8×8 attempt was Excel-wrong.
  darkHorizontal: [
    0b111111111111, 0b111111111111, 0b000000000000,
    0b111111111111, 0b111111111111, 0b000000000000,
    0b111111111111, 0b111111111111, 0b000000000000,
    0b111111111111, 0b111111111111, 0b000000000000,
  ],
  lightHorizontal: [
    0b111111111111, 0b000000000000, 0b000000000000,
    0b111111111111, 0b000000000000, 0b000000000000,
    0b111111111111, 0b000000000000, 0b000000000000,
    0b111111111111, 0b000000000000, 0b000000000000,
  ],
  // darkVertical: 2-col bars at cols 0,1 / 3,4 / 6,7 / 9,10
  //   bits set per row: 11,10 + 8,7 + 5,4 + 2,1 = 0b110110110110 = 0xDB6
  darkVertical:   Array(12).fill(0xDB6),
  // lightVertical: 1-col lines at cols 0 / 3 / 6 / 9
  //   bits set per row: 11 + 8 + 5 + 2 = 0b100100100100 = 0x924
  lightVertical:  Array(12).fill(0x924),

  // darkGrid: Excel renders this as a fine 2×2 checkerboard, not a thick
  // horizontal+vertical lattice. Each black/white "cell" is 2×2 source
  // pixels, so the cell reads as a clear "市松模様" rather than a solid
  // 50% gray (the 1×1 checkerboard mediumGray we already ship).
  darkGrid:  [0b11001100, 0b11001100, 0b00110011, 0b00110011, 0b11001100, 0b11001100, 0b00110011, 0b00110011],
  // lightGrid: a sparse grid with 4-px pitch — horizontal lines at rows
  // 0 and 4, vertical lines at cols 0 and 4. Grid cells (white squares
  // between the lines) are 3×3 source pixels, matching Excel's rendering
  // of lightGrid where the lattice is clearly larger than the dense
  // version and a true "格子" pattern is recognisable.
  lightGrid: [0b11111111, 0b10001000, 0b10001000, 0b10001000, 0b11111111, 0b10001000, 0b10001000, 0b10001000],

  // ── 8×8 — diagonals & trellis ─────────────────────────────────────────
  // darkDown / darkUp: 2-px-wide diagonal stripes every 4 cells.
  // lightDown / lightUp: 1-px-wide diagonal stripes at the same 4-cell pitch.
  darkDown:   [0b11001100, 0b01100110, 0b00110011, 0b10011001, 0b11001100, 0b01100110, 0b00110011, 0b10011001],
  lightDown:  [0b10001000, 0b01000100, 0b00100010, 0b00010001, 0b10001000, 0b01000100, 0b00100010, 0b00010001],
  darkUp:     [0b00110011, 0b01100110, 0b11001100, 0b10011001, 0b00110011, 0b01100110, 0b11001100, 0b10011001],
  lightUp:    [0b00010001, 0b00100010, 0b01000100, 0b10001000, 0b00010001, 0b00100010, 0b01000100, 0b10001000],
  // darkTrellis = darkDown | darkUp.
  darkTrellis:  [0b11111111, 0b01100110, 0b11111111, 0b10011001, 0b11111111, 0b01100110, 0b11111111, 0b10011001],
  // lightTrellis = lightDown | lightUp.
  lightTrellis: [0b10011001, 0b01100110, 0b01100110, 0b10011001, 0b10011001, 0b01100110, 0b01100110, 0b10011001],
};

/**
 * Build a repeating 8x8 tile for an ECMA-376 preset pattern. Returns null
 * for unknown pattern names so the caller can fall back to a flat colour.
 */
function hatchPattern(
  ctx: CanvasRenderingContext2D,
  pt: string,
  fgHex: string,
  bgHex: string,
): CanvasPattern | null {
  // Read the destination context's effective scale so the offscreen tile is
  // sized at the destination's device-pixel resolution. Drawing the source
  // pre-scaled keeps each "1 bit" at exactly one destination *device* pixel
  // when the rendering context is doing ctx.scale(dpr * cs) for hi-DPI /
  // Storybook zoom — without this the crisp 1-px source bits get linearly
  // resampled to 1.5–2 device px and either chunky (integer scale) or
  // blurred-out-of-existence (non-integer scale, the prior pat.setTransform
  // approach lost lightHorizontal entirely at scale=1.5).
  const t = ctx.getTransform();
  const sx = Math.max(1, Math.round(Math.hypot(t.a, t.b)));
  const sy = Math.max(1, Math.round(Math.hypot(t.c, t.d)));
  const key = `${pt}|${fgHex}|${bgHex}|${sx}|${sy}`;
  if (PATTERN_CACHE.has(key)) return PATTERN_CACHE.get(key)!;

  const rows = PATTERN_BITMAPS[pt];
  if (!rows) {
    PATTERN_CACHE.set(key, null);
    return null;
  }

  // Square tile — size inferred from the row count. Bit (size-1) is
  // leftmost so the binary literals read left-to-right at any tile width.
  // The offscreen canvas is sized at sx × tileSize device pixels so each
  // "1 bit" becomes a sx × sy rect when the destination context's scale
  // multiplies it back. Result: the source bit lands on exactly one
  // destination device pixel regardless of the user's CSS zoom.
  const tile = rows.length;
  const off = document.createElement('canvas');
  off.width = tile;
  off.height = tile;
  const octx = off.getContext('2d');
  if (!octx) { PATTERN_CACHE.set(key, null); return null; }

  octx.fillStyle = hexToRgba(bgHex);
  octx.fillRect(0, 0, tile, tile);
  octx.fillStyle = hexToRgba(fgHex);
  for (let y = 0; y < tile; y++) {
    const row = rows[y];
    for (let x = 0; x < tile; x++) {
      if (row & (1 << (tile - 1 - x))) octx.fillRect(x, y, 1, 1);
    }
  }

  const pat = ctx.createPattern(off, 'repeat');
  // Pre-bake an inverse-scale matrix only when the scale rounds to an
  // integer (1, 2, 3, …). Fractional inverses (e.g. 1/1.5 ≈ 0.67) trigger
  // the canvas's bilinear pattern resampler and smear 1-px features into
  // a uniform half-tone — which is exactly what wiped out lightHorizontal
  // at Storybook scale=1.5. For non-integer ctx scales we leave the tile
  // unscaled and accept ~1.5-device-px bits (slightly chunky but crisp).
  if (pat && typeof DOMMatrix !== 'undefined' && (sx >= 2 || sy >= 2)) {
    const m = new DOMMatrix();
    m.scaleSelf(1 / sx, 1 / sy);
    pat.setTransform(m);
  }
  PATTERN_CACHE.set(key, pat);
  return pat;
}

/**
 * Paint a cell background according to its <patternFill> / <gradientFill>.
 *
 * ECMA-376 §18.8.20 / §18.8.22 specifies fg/bg defaults: when the colour
 * children are absent, fg defaults to the system foreground (black) and bg
 * to the system background (white). Without those defaults the directional
 * hatches that Excel emits with no explicit colours (`<patternFill
 * patternType="darkHorizontal"/>` etc.) would render as nothing because
 * the prior gate required a non-null fgColor.
 *
 * Returns true when the cell was painted so the caller can short-circuit
 * its tableStyle / banded fallbacks.
 */
function paintCellPatternFill(
  ctx: CanvasRenderingContext2D,
  fill: Fill,
  x: number, y: number, w: number, h: number,
): boolean {
  if (fill.gradient && fill.gradient.stops.length > 0) {
    ctx.fillStyle = buildGradientFill(ctx, fill.gradient, x, y, w, h);
    ctx.fillRect(x, y, w, h);
    return true;
  }
  const pt = fill.patternType;
  if (!pt || pt === 'none') return false;
  const fg = fill.fgColor ?? '000000';
  const bg = fill.bgColor ?? 'FFFFFF';
  if (pt === 'solid') {
    ctx.fillStyle = hexToRgba(fg);
    ctx.fillRect(x, y, w, h);
    return true;
  }
  const hatch = hatchPattern(ctx, pt, fg, bg);
  if (hatch) {
    ctx.fillStyle = hatch;
  } else {
    const coverage = patternCoverage(pt);
    ctx.fillStyle = coverage >= 1
      ? hexToRgba(fg)
      : blendHex(fg, bg, coverage);
  }
  ctx.fillRect(x, y, w, h);
  return true;
}

/**
 * Build a Canvas gradient object for an xlsx `<gradientFill>`. Linear uses
 * the degree attribute (0° = left→right, 90° = top→bottom). Path gradients
 * radiate from a rectangular inner bounds defined by left/right/top/bottom
 * as fractions of the cell.
 */
function buildGradientFill(
  ctx: CanvasRenderingContext2D,
  g: GradientFillSpec,
  x: number, y: number, w: number, h: number,
): CanvasGradient {
  let grad: CanvasGradient;
  if (g.gradientType === 'path') {
    // Use the inner rectangle's center as the radial origin; radius spans to
    // the farthest cell corner so stop=1 always reaches a cell edge.
    const cxg = x + w * (g.left + (1 - g.right - g.left) / 2);
    const cyg = y + h * (g.top + (1 - g.bottom - g.top) / 2);
    const r = Math.hypot(Math.max(cxg - x, x + w - cxg), Math.max(cyg - y, y + h - cyg));
    grad = ctx.createRadialGradient(cxg, cyg, 0, cxg, cyg, r);
  } else {
    // Linear: rotate around the cell's center and extend to the bounds.
    const rad = (g.degree * Math.PI) / 180;
    const cxg = x + w / 2;
    const cyg = y + h / 2;
    const ext = (Math.abs(Math.cos(rad)) * w + Math.abs(Math.sin(rad)) * h) / 2;
    grad = ctx.createLinearGradient(
      cxg - Math.cos(rad) * ext, cyg - Math.sin(rad) * ext,
      cxg + Math.cos(rad) * ext, cyg + Math.sin(rad) * ext,
    );
  }
  for (const stop of g.stops) {
    const pos = Math.min(1, Math.max(0, stop.position));
    grad.addColorStop(pos, hexToRgba(stop.color));
  }
  return grad;
}

/**
 * Parse an A1-style cell reference ("A1", "B12", "AA3") to 1-based row/col.
 * Returns null when the input doesn't match the expected shape (parser-side
 * data is trusted, but we still guard against malformed refs).
 */
function parseA1Ref(ref: string): { row: number; col: number } | null {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!m) return null;
  const colLetters = m[1];
  const row = parseInt(m[2], 10);
  let col = 0;
  for (let i = 0; i < colLetters.length; i++) {
    col = col * 26 + (colLetters.charCodeAt(i) - 64);
  }
  return { row, col };
}

/**
 * Draw Excel's comment marker — a small filled triangle in the top-right
 * corner of the cell — coloured like Excel's default red indicator. Scales
 * with cell size but is clamped so it stays legible at small zoom.
 */
function drawCommentMarker(
  ctx: CanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
): void {
  const size = Math.max(4, Math.min(8, Math.min(w, h) * 0.18));
  ctx.save();
  ctx.fillStyle = '#D40000';
  ctx.beginPath();
  ctx.moveTo(x + w - size, y);
  ctx.lineTo(x + w, y);
  ctx.lineTo(x + w, y + size);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
}

/** Linear-interpolate two #RRGGBB values in RGB space at coverage fg weight. */
function blendHex(fgHex: string, bgHex: string, fgCoverage: number): string {
  const fh = fgHex.replace('#', '');
  const bh = bgHex.replace('#', '');
  const fr = parseInt(fh.slice(0, 2), 16);
  const fg = parseInt(fh.slice(2, 4), 16);
  const fb = parseInt(fh.slice(4, 6), 16);
  const br = parseInt(bh.slice(0, 2), 16);
  const bg = parseInt(bh.slice(2, 4), 16);
  const bb = parseInt(bh.slice(4, 6), 16);
  const c = Math.min(1, Math.max(0, fgCoverage));
  const r = Math.round(fr * c + br * (1 - c));
  const g = Math.round(fg * c + bg * (1 - c));
  const b = Math.round(fb * c + bb * (1 - c));
  return `rgb(${r},${g},${b})`;
}

function buildFont(font: Font, cs = 1): string {
  const style = font.italic ? 'italic ' : '';
  const weight = font.bold ? 'bold ' : '';
  const sizePx = Math.max(1, Math.round(font.size * PT_TO_PX * cs));
  const family = font.name ? `"${font.name}", ${DEFAULT_FONT_FAMILY}` : DEFAULT_FONT_FAMILY;
  return `${style}${weight}${sizePx}px ${family}`;
}

/**
 * Stroke a single or double horizontal text-decoration line
 * (underline / strikethrough). ECMA-376 §18.4.13 ST_UnderlineValues
 * "double" / "doubleAccounting" both render as two parallel lines with a
 * small gap. The two strokes straddle the requested baseline ±1px so the
 * pair reads as a ~3px-thick rule and the visual centre matches Excel.
 */
function drawTextDecoLine(
  ctx: CanvasRenderingContext2D,
  x1: number, x2: number, y: number,
  color: string, double: boolean,
): void {
  ctx.save();
  ctx.strokeStyle = color;
  ctx.lineWidth = 0.5;
  ctx.beginPath();
  if (double) {
    ctx.moveTo(x1, y - 1); ctx.lineTo(x2, y - 1);
    ctx.moveTo(x1, y + 1); ctx.lineTo(x2, y + 1);
  } else {
    ctx.moveTo(x1, y); ctx.lineTo(x2, y);
  }
  ctx.stroke();
  ctx.restore();
}

/**
 * Resolve a Run's font against a base Font. Per ECMA-376, a run's <rPr>
 * completely specifies bold/italic/underline/strike for that run, while
 * size/color/name fall back to the base when omitted. A run with no
 * <rPr> (run.font undefined) inherits the base entirely.
 */
function applyRunFont(base: Font, run: Run): Font {
  const rf = run.font;
  if (!rf) return base;
  return {
    bold: rf.bold,
    italic: rf.italic,
    underline: rf.underline,
    underlineStyle: rf.underlineStyle,
    strike: rf.strike,
    size: rf.size ?? base.size,
    color: rf.color ?? base.color,
    name: rf.name ?? base.name,
    vertAlign: rf.vertAlign,
  };
}

function resolveXf(styles: Styles, styleIndex: number): { font: Font; fill: Fill; border: Border; xf: CellXf } {
  const xf: CellXf = styles.cellXfs[styleIndex] ?? styles.cellXfs[0] ?? {
    fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, alignH: null, alignV: null, wrapText: false,
  };
  const font: Font = styles.fonts[xf.fontId] ?? { bold: false, italic: false, underline: false, strike: false, size: DEFAULT_FONT_SIZE, color: null, name: null };
  const fill: Fill = styles.fills[xf.fillId] ?? { patternType: 'none', fgColor: null, bgColor: null };
  const border: Border = styles.borders[xf.borderId] ?? { left: null, right: null, top: null, bottom: null };
  return { font, fill, border, xf };
}

function wrapTextLines(ctx: CanvasRenderingContext2D, text: string, maxWidth: number): string[] {
  const lines: string[] = [];
  // Hard line breaks (\n from Alt+Enter) always split regardless of wrapText.
  for (const paragraph of text.split('\n')) {
    lines.push(...wrapParagraphLines(ctx, paragraph, maxWidth));
  }
  return lines;
}

/** Codepoints in the CJK ranges get broken per-character (Excel / JIS X 4051
 *  behaviour). This mirrors the tokenizer in `layoutRichTextLines`. */
function isCJKCodePoint(cp: number): boolean {
  return (cp >= 0x3000 && cp <= 0x9FFF)  // CJK punctuation + CJK Unified Ideographs
      || (cp >= 0xF900 && cp <= 0xFAFF)  // CJK Compatibility Ideographs
      || (cp >= 0xAC00 && cp <= 0xD7AF)  // Hangul Syllables
      || (cp >= 0xFF00 && cp <= 0xFFEF); // Halfwidth/Fullwidth
}

/** Word-wrap a single paragraph (no embedded \n). Unlike a naive
 *  `split(' ')`, CJK characters are treated as individual break opportunities
 *  so that Japanese headings like "夏休みアクティビティ カレンダー 2026"
 *  actually wrap inside a merged cell. ECMA-376 doesn't spec the break
 *  algorithm but this matches what Excel renders on the same input. */
function wrapParagraphLines(ctx: CanvasRenderingContext2D, paragraph: string, maxWidth: number): string[] {
  const lines: string[] = [];
  // Tokenise: runs of non-space non-CJK, single ASCII-space runs, individual
  // CJK characters. Then greedy-fit each token onto the current line.
  const tokens: string[] = [];
  let i = 0;
  while (i < paragraph.length) {
    const ch = paragraph[i];
    const cp = ch.codePointAt(0) ?? 0;
    if (isCJKCodePoint(cp)) {
      tokens.push(ch);
      i += cp > 0xFFFF ? 2 : 1;
    } else if (ch === ' ') {
      let j = i;
      while (j < paragraph.length && paragraph[j] === ' ') j++;
      tokens.push(paragraph.slice(i, j));
      i = j;
    } else {
      let j = i;
      while (j < paragraph.length) {
        const c = paragraph[j];
        const p = c.codePointAt(0) ?? 0;
        if (c === ' ' || isCJKCodePoint(p)) break;
        j += p > 0xFFFF ? 2 : 1;
      }
      tokens.push(paragraph.slice(i, j));
      i = j;
    }
  }
  let current = '';
  for (const tok of tokens) {
    if (current === '') { current = tok; continue; }
    const candidate = current + tok;
    if (ctx.measureText(candidate).width <= maxWidth) {
      current = candidate;
    } else {
      // Token doesn't fit at the end of the current line — break here.
      // Leading spaces at the start of the next line are dropped (matches
      // Excel: wrapped-continuation lines don't preserve the space that
      // caused the break).
      lines.push(current);
      current = tok.replace(/^ +/, '');
      if (current === '') current = tok; // all-space token (preserve width on its own line)
    }
  }
  lines.push(current);
  return lines;
}

interface RichSeg {
  text: string;
  font: Font;
  width: number; // px
}

interface RichLine {
  segments: RichSeg[];
  maxFontSize: number; // pt (line-height source)
}

/**
 * Layout rich text runs into wrapped lines. Each run is split into words (and
 * CJK characters for granular wrapping). Per-run font is preserved so measurement
 * and drawing use the correct font.
 *
 * Follows ECMA-376 §18.3.1.53 (w:r) semantics: runs are inline and share the
 * paragraph width. wrapText breaks at word boundaries (ASCII spaces) and at any
 * CJK code point boundary.
 */
function layoutRichTextLines(
  ctx: CanvasRenderingContext2D,
  runs: Run[],
  baseFont: Font,
  cs: number,
  maxWidth: number,
): RichLine[] {
  const lines: RichLine[] = [];
  let cur: RichSeg[] = [];
  let curW = 0;
  let curMaxSize = 0;

  const flush = () => {
    if (cur.length === 0) return;
    lines.push({ segments: cur, maxFontSize: curMaxSize });
    cur = []; curW = 0; curMaxSize = 0;
  };

  const push = (text: string, font: Font) => {
    if (!text) return;
    ctx.font = buildFont(font, cs);
    const w = ctx.measureText(text).width;
    if (cur.length > 0 && curW + w > maxWidth) flush();
    cur.push({ text, font, width: w });
    curW += w;
    if (font.size > curMaxSize) curMaxSize = font.size;
  };

  const isCJK = (cp: number) =>
    (cp >= 0x3000 && cp <= 0x9FFF) ||
    (cp >= 0xF900 && cp <= 0xFAFF) ||
    (cp >= 0xAC00 && cp <= 0xD7AF) ||
    (cp >= 0xFF00 && cp <= 0xFFEF);

  for (const run of runs) {
    const font = applyRunFont(baseFont, run);
    // Tokenize: runs of non-space latin, spaces, or individual CJK chars
    const tokens: string[] = [];
    let i = 0;
    while (i < run.text.length) {
      const ch = run.text[i];
      const cp = ch.codePointAt(0) ?? 0;
      if (cp === 0x000A) {
        // Explicit newline: force break
        tokens.push('\n'); i += 1;
      } else if (isCJK(cp)) {
        tokens.push(ch);
        i += cp > 0xFFFF ? 2 : 1;
      } else if (ch === ' ') {
        let j = i;
        while (j < run.text.length && run.text[j] === ' ') j++;
        tokens.push(run.text.slice(i, j));
        i = j;
      } else {
        let j = i;
        while (j < run.text.length) {
          const c = run.text[j];
          const p = c.codePointAt(0) ?? 0;
          if (c === ' ' || c === '\n' || isCJK(p)) break;
          j += p > 0xFFFF ? 2 : 1;
        }
        tokens.push(run.text.slice(i, j));
        i = j;
      }
    }
    for (const tok of tokens) {
      if (tok === '\n') flush();
      else push(tok, font);
    }
  }
  flush();
  return lines;
}

function colToLetter(col: number): string {
  let result = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    col = Math.floor((col - 1) / 26);
  }
  return result;
}

// ────────────────────────────────────────────────────────────────
// Conditional formatting
// ────────────────────────────────────────────────────────────────
interface CompiledCfRule {
  rule: CfRule;
  sqref: CellRange[];
  scaleMin?: number;
  scaleMax?: number;
  scaleStops?: number[];
  barMin?: number;
  barMax?: number;
  top10Threshold?: number;
  top10IsTop?: boolean;
  avgValue?: number;
  avgIsAbove?: boolean;
  iconThresholds?: number[];
}

interface CfContext {
  compiled: CompiledCfRule[];
  worksheet: Worksheet;
  cellIndex: Map<string, Cell>;
  definedNames: Map<string, DefinedName>;
}

interface CfResult {
  fill?: Fill;
  fontColor?: string;
  fontBold?: boolean;
  fontItalic?: boolean;
  fontUnderline?: boolean;
  fontStrike?: boolean;
  /** Number format override from a matched CF dxf. Higher-priority rules win
   *  (first match through the rule list). Falls back to the cell's own style
   *  numFmt if unset. */
  numFmt?: { numFmtId: number; formatCode: string | null };
  dataBar?: { color: string; ratio: number; gradient: boolean };
  iconSet?: { name: string; index: number };
  /** Per-edge borders from matched CF rules (merged on top of the cell's base
   *  border). Mostly used by `expression` rules whose dxf only sets borders,
   *  e.g. highlighting today's column in a Gantt chart. */
  border?: Border;
}

function rangeContains(ranges: CellRange[], row: number, col: number): boolean {
  for (const r of ranges) {
    if (row >= r.top && row <= r.bottom && col >= r.left && col <= r.right) return true;
  }
  return false;
}

function cellNumericValue(cell: Cell | undefined): number | null {
  if (!cell) return null;
  if (cell.value.type === 'number') return cell.value.number;
  return null;
}

function cellTextValue(cell: Cell | undefined): string | null {
  if (!cell) return null;
  if (cell.value.type === 'text') return cell.value.text;
  return null;
}

function collectNumericValuesInRanges(worksheet: Worksheet, ranges: CellRange[]): number[] {
  const out: number[] = [];
  for (const row of worksheet.rows) {
    for (const c of row.cells) {
      if (c.value.type !== 'number') continue;
      if (rangeContains(ranges, c.row, c.col)) out.push(c.value.number);
    }
  }
  return out;
}

function resolveCfvoValue(cfv: CfValue | CfStop, samples: number[]): number {
  const minv = samples.length ? Math.min(...samples) : 0;
  const maxv = samples.length ? Math.max(...samples) : 0;
  const n = cfv.value != null ? parseFloat(cfv.value) : NaN;
  switch (cfv.kind) {
    case 'min': return minv;
    case 'max': return maxv;
    case 'num': return isNaN(n) ? 0 : n;
    case 'percent': {
      const p = isNaN(n) ? 50 : n;
      return minv + (maxv - minv) * (p / 100);
    }
    case 'percentile': {
      if (!samples.length) return 0;
      const sorted = [...samples].sort((a, b) => a - b);
      const p = (isNaN(n) ? 50 : n) / 100;
      const idx = Math.max(0, Math.min(sorted.length - 1, Math.round(p * (sorted.length - 1))));
      return sorted[idx];
    }
    default: return isNaN(n) ? 0 : n;
  }
}

function compileCf(worksheet: Worksheet): CfContext {
  const compiled: CompiledCfRule[] = [];
  const cellIndex = new Map<string, Cell>();
  for (const row of worksheet.rows) {
    for (const c of row.cells) {
      cellIndex.set(`${c.row}:${c.col}`, c);
    }
  }
  const definedNames = new Map<string, DefinedName>();
  for (const dn of worksheet.definedNames ?? []) {
    definedNames.set(dn.name, dn);
  }
  for (const cf of worksheet.conditionalFormats ?? []) {
    const samples = collectNumericValuesInRanges(worksheet, cf.sqref);
    for (const rule of cf.rules) {
      const entry: CompiledCfRule = { rule, sqref: cf.sqref };
      if (rule.type === 'colorScale') {
        entry.scaleStops = rule.stops.map(s => resolveCfvoValue(s, samples));
      } else if (rule.type === 'dataBar') {
        entry.barMin = resolveCfvoValue(rule.min, samples);
        entry.barMax = resolveCfvoValue(rule.max, samples);
      } else if (rule.type === 'top10') {
        const sorted = [...samples].sort((a, b) => a - b);
        const n = sorted.length;
        if (n > 0) {
          const rank = Math.min(rule.rank, n);
          if (rule.percent) {
            const p = rule.top ? (1 - rank / 100) : (rank / 100);
            const idx = Math.max(0, Math.min(n - 1, Math.round(p * (n - 1))));
            entry.top10Threshold = sorted[idx];
          } else {
            entry.top10Threshold = rule.top ? sorted[Math.max(0, n - rank)] : sorted[Math.min(n - 1, rank - 1)];
          }
          entry.top10IsTop = rule.top;
        }
      } else if (rule.type === 'aboveAverage') {
        if (samples.length > 0) {
          entry.avgValue = samples.reduce((a, b) => a + b, 0) / samples.length;
          entry.avgIsAbove = rule.aboveAverage;
        }
      } else if (rule.type === 'iconSet') {
        entry.iconThresholds = rule.cfvos.map(cfv => resolveCfvoValue(cfv, samples));
      }
      compiled.push(entry);
    }
  }
  // Excel evaluates CF rules in ascending priority (lowest number = highest
  // priority first). For each property (fill/fontColor/border/…) the first
  // matching rule wins, and `stopIfTrue` on a matching rule skips all later
  // rules. Match that here by iterating asc and only setting properties that
  // are still unset.
  compiled.sort((a, b) => {
    const pa = (a.rule as { priority: number }).priority ?? 0;
    const pb = (b.rule as { priority: number }).priority ?? 0;
    return pa - pb;
  });
  return { compiled, worksheet, cellIndex, definedNames };
}

function cellIsMatch(num: number, operator: string, args: number[]): boolean {
  switch (operator) {
    case 'greaterThan': return num > (args[0] ?? 0);
    case 'greaterThanOrEqual': return num >= (args[0] ?? 0);
    case 'lessThan': return num < (args[0] ?? 0);
    case 'lessThanOrEqual': return num <= (args[0] ?? 0);
    case 'equal': return num === (args[0] ?? 0);
    case 'notEqual': return num !== (args[0] ?? 0);
    case 'between': return num >= (args[0] ?? 0) && num <= (args[1] ?? 0);
    case 'notBetween': return num < (args[0] ?? 0) || num > (args[1] ?? 0);
    default: return false;
  }
}

function parseCellIsFormula(f: string): { text?: string; num?: number } {
  const t = f.trim();
  if (t.length >= 2 && t.startsWith('"') && t.endsWith('"')) {
    return { text: t.slice(1, -1).replace(/""/g, '"') };
  }
  const n = parseFloat(t);
  if (!isNaN(n)) return { num: n };
  return { text: t };
}

function cellIsTextMatch(text: string, operator: string, args: string[]): boolean {
  const a = args[0] ?? '';
  const b = args[1] ?? '';
  const ci = (s: string) => s.toLowerCase();
  switch (operator) {
    case 'equal':         return ci(text) === ci(a);
    case 'notEqual':      return ci(text) !== ci(a);
    case 'containsText':  return ci(text).includes(ci(a));
    case 'notContains':   return !ci(text).includes(ci(a));
    case 'beginsWith':    return ci(text).startsWith(ci(a));
    case 'endsWith':      return ci(text).endsWith(ci(a));
    case 'between':       return ci(text) >= ci(a) && ci(text) <= ci(b);
    case 'notBetween':    return ci(text) <  ci(a) || ci(text) >  ci(b);
    default: return false;
  }
}

function interpolateHex(a: string, b: string, t: number): string {
  const pa = a.replace('#', '');
  const pb = b.replace('#', '');
  const ar = parseInt(pa.slice(0, 2), 16), ag = parseInt(pa.slice(2, 4), 16), ab = parseInt(pa.slice(4, 6), 16);
  const br = parseInt(pb.slice(0, 2), 16), bg = parseInt(pb.slice(2, 4), 16), bb = parseInt(pb.slice(4, 6), 16);
  const r = Math.round(ar + (br - ar) * t);
  const g = Math.round(ag + (bg - ag) * t);
  const bl = Math.round(ab + (bb - ab) * t);
  return `#${r.toString(16).padStart(2, '0').toUpperCase()}${g.toString(16).padStart(2, '0').toUpperCase()}${bl.toString(16).padStart(2, '0').toUpperCase()}`;
}

function colorScaleAt(num: number, stops: CfStop[], stopValues: number[]): string {
  if (!stops.length) return '#FFFFFF';
  if (num <= stopValues[0]) return stops[0].color;
  if (num >= stopValues[stopValues.length - 1]) return stops[stops.length - 1].color;
  for (let i = 1; i < stopValues.length; i++) {
    if (num <= stopValues[i]) {
      const lo = stopValues[i - 1];
      const hi = stopValues[i];
      const t = hi === lo ? 0 : (num - lo) / (hi - lo);
      return interpolateHex(stops[i - 1].color, stops[i].color, t);
    }
  }
  return stops[stops.length - 1].color;
}

function applyDxfToResult(result: CfResult, dxf: Dxf | null | undefined): void {
  if (!dxf) return;
  // First-match-wins (higher priority) for each property. See compileCf.
  // Per ECMA-376 §18.3.1.11, a `<dxf>` is a *differential* format: any child
  // element it contains is an override of the base cell format. So the mere
  // presence of `dxf.fill` means "replace the base fill with this", whatever
  // its patternType / color — including `patternType="none"` (explicit clear)
  // and gradient fills. The paint-site guard (`patternType !== 'none' &&
  // fgColor`) handles whether the result actually paints a color or leaves
  // the cell transparent, so this override stays spec-faithful without
  // second-guessing the fill's shape here.
  if (dxf.fill && !result.fill) result.fill = dxf.fill;
  if (dxf.font?.color && result.fontColor == null) result.fontColor = dxf.font.color;
  if (dxf.font?.bold && result.fontBold == null) result.fontBold = true;
  if (dxf.font?.italic && result.fontItalic == null) result.fontItalic = true;
  if (dxf.font?.underline && result.fontUnderline == null) result.fontUnderline = true;
  if (dxf.font?.strike && result.fontStrike == null) result.fontStrike = true;
  if (dxf.numFmt && result.numFmt == null) {
    result.numFmt = {
      numFmtId: dxf.numFmt.numFmtId,
      formatCode: dxf.numFmt.formatCode || null,
    };
  }
  if (dxf.border) {
    // Merge per-edge — higher-priority edges stay; lower-priority edges fill
    // in unset ones. dxf `border` typically sets only the edges the rule
    // cares about (e.g. left+right for a "today" column marker).
    const existing = result.border ?? {} as Border;
    const merged: Border = {
      left:         existing.left         ?? dxf.border.left,
      right:        existing.right        ?? dxf.border.right,
      top:          existing.top          ?? dxf.border.top,
      bottom:       existing.bottom       ?? dxf.border.bottom,
      diagonalUp:   existing.diagonalUp   ?? dxf.border.diagonalUp,
      diagonalDown: existing.diagonalDown ?? dxf.border.diagonalDown,
    };
    result.border = merged;
  }
}

function evaluateCf(cell: Cell | undefined, row: number, col: number, cfCtx: CfContext, dxfs: Dxf[]): CfResult {
  const result: CfResult = {};
  if (!cfCtx.compiled.length) return result;
  for (const entry of cfCtx.compiled) {
    if (!rangeContains(entry.sqref, row, col)) continue;
    const rule = entry.rule;
    const numVal = cellNumericValue(cell);

    if (rule.type === 'expression') {
      const anchor = entry.sqref[0];
      if (!anchor) continue;
      const matched = evalFormulaToBool(rule.formula, {
        row, col,
        anchorRow: anchor.top, anchorCol: anchor.left,
        cellIndex: cfCtx.cellIndex,
        definedNames: cfCtx.definedNames,
        depth: 0,
      });
      if (matched) {
        applyDxfToResult(result, rule.dxfId != null ? dxfs[rule.dxfId] : null);
        if (rule.stopIfTrue) break;
      }
      continue;
    }

    if (rule.type === 'cellIs') {
      const parsedArgs = rule.formulas.map(parseCellIsFormula);
      const textVal = cellTextValue(cell);
      let matched = false;
      if (numVal != null && parsedArgs.every(a => a.num != null)) {
        matched = cellIsMatch(numVal, rule.operator, parsedArgs.map(a => a.num!));
      } else if (textVal != null && parsedArgs.every(a => a.text != null)) {
        matched = cellIsTextMatch(textVal, rule.operator, parsedArgs.map(a => a.text!));
      }
      if (matched) {
        applyDxfToResult(result, rule.dxfId != null ? dxfs[rule.dxfId] : null);
      }
    } else if (rule.type === 'top10') {
      if (numVal == null || entry.top10Threshold == null) continue;
      const matches = entry.top10IsTop ? numVal >= entry.top10Threshold : numVal <= entry.top10Threshold;
      if (matches) applyDxfToResult(result, rule.dxfId != null ? dxfs[rule.dxfId] : null);
    } else if (rule.type === 'aboveAverage') {
      if (numVal == null || entry.avgValue == null) continue;
      const matches = entry.avgIsAbove ? numVal > entry.avgValue : numVal < entry.avgValue;
      if (matches) applyDxfToResult(result, rule.dxfId != null ? dxfs[rule.dxfId] : null);
    } else if (rule.type === 'iconSet') {
      if (numVal == null || !entry.iconThresholds?.length) continue;
      const thresholds = entry.iconThresholds;
      const n = thresholds.length;
      let iconIdx = 0;
      for (let i = 1; i < n; i++) {
        if (numVal >= thresholds[i]) iconIdx = i;
      }
      if (rule.reverse) iconIdx = n - 1 - iconIdx;
      // Custom iconSets (Excel 2010+ x14 extension) override per-threshold icons.
      if (rule.customIcons && rule.customIcons[iconIdx]) {
        const ci = rule.customIcons[iconIdx];
        if (ci.iconSet !== 'NoIcons') {
          result.iconSet = { name: ci.iconSet, index: ci.iconId };
        }
      } else {
        result.iconSet = { name: rule.iconSet, index: iconIdx };
      }
    } else if (rule.type === 'colorScale') {
      if (numVal == null || !entry.scaleStops) continue;
      if (result.fill) continue;
      const color = colorScaleAt(numVal, rule.stops, entry.scaleStops);
      result.fill = { patternType: 'solid', fgColor: color, bgColor: color };
    } else if (rule.type === 'dataBar') {
      if (numVal == null || entry.barMin == null || entry.barMax == null) continue;
      if (result.dataBar) continue;
      const range = entry.barMax - entry.barMin;
      const ratio = range === 0 ? 0 : Math.max(0, Math.min(1, (numVal - entry.barMin) / range));
      result.dataBar = { color: rule.color, ratio, gradient: rule.gradient };
    }
  }
  return result;
}


// ────────────────────────────────────────────────────────────────
// Shared state for a single renderViewport call
// ────────────────────────────────────────────────────────────────
interface RenderContext {
  worksheet: Worksheet;
  styles: Styles;
  cellMap: Map<string, Cell>;
  mergeAnchorMap: Map<string, { totalW: number; totalH: number; right: number; bottom: number }>;
  mergeSkipSet: Set<string>;
  cfContext: CfContext;
  colWidths: number[];
  rowHeights: number[];
  frozenColWidths: number[];
  frozenRowHeights: number[];
  frozenW: number;
  frozenH: number;
  startRow: number;
  startCol: number;
  cs: number;
  dpr: number;
  autoFilterCells: Set<string>;
  hyperlinkMap: Map<string, string>;
  /** row:col keys for cells that carry a comment; renderer draws a small
   *  red triangle in the top-right corner (ECMA-376 §18.7.3 commentList). */
  commentCells: Set<string>;
  /** row:col → table-style overlay (bold header, banded rows, borders). */
  tableStyleMap: Map<string, TableCellStyle>;
  /** row:col → render-ready SparklineModel for cells that host an
   *  `x14:sparkline`. Built once at viewport start by flattening the
   *  parser's SparklineGroup + per-cell Sparkline pair. */
  sparklineMap: Map<string, SparklineModel>;
  /** Max Digit Width resolved for the worksheet's Normal-style font
   *  (ECMA-376 §18.3.1.13). Used by `colWidthToPx` to convert character-
   *  unit column widths into pixels. */
  mdw: number;
  onTextRun?: (info: TextRunInfo) => void;
}

// ────────────────────────────────────────────────────────────────
// Icon Set drawing
// ────────────────────────────────────────────────────────────────
const ICON_COLORS_3 = ['#FF0000', '#FFFF00', '#00B050'];
const ICON_COLORS_4 = ['#FF0000', '#FF6600', '#FFFF00', '#00B050'];
const ICON_COLORS_5 = ['#FF0000', '#FF6600', '#FFFF00', '#92D050', '#00B050'];

function drawCfIcon(ctx: CanvasRenderingContext2D, name: string, index: number, x: number, y: number, sz: number): void {
  if (name === 'NoIcons') return;
  const safeName = name || '3TrafficLights1';
  const nIcons = parseInt(safeName[0]) || 3;
  const palette = nIcons === 5 ? ICON_COLORS_5 : nIcons === 4 ? ICON_COLORS_4 : ICON_COLORS_3;
  const color = palette[Math.max(0, Math.min(index, palette.length - 1))];
  ctx.save();
  ctx.fillStyle = color;
  if (safeName.includes('Arrow')) {
    const half = sz / 2;
    ctx.beginPath();
    if (index === nIcons - 1) {
      ctx.moveTo(x + half, y); ctx.lineTo(x + sz, y + sz); ctx.lineTo(x, y + sz);
    } else if (index === 0) {
      ctx.moveTo(x, y); ctx.lineTo(x + sz, y); ctx.lineTo(x + half, y + sz);
    } else {
      ctx.moveTo(x, y + sz * 0.3); ctx.lineTo(x + sz, y + half); ctx.lineTo(x, y + sz * 0.7);
    }
    ctx.closePath();
    ctx.fill();
  } else if (safeName.includes('Flag')) {
    ctx.beginPath();
    ctx.moveTo(x, y); ctx.lineTo(x + sz, y); ctx.lineTo(x, y + sz);
    ctx.closePath();
    ctx.fill();
  } else {
    ctx.beginPath();
    ctx.arc(x + sz / 2, y + sz / 2, sz / 2, 0, Math.PI * 2);
    ctx.fill();
  }
  ctx.restore();
}

function drawAutoFilterArrow(ctx: CanvasRenderingContext2D, cx: number, cy: number, cw: number, ch: number): void {
  const sz = Math.max(6, Math.round(Math.min(cw, ch) * 0.45));
  const x = cx + cw - sz - 1;
  const y = cy + ch - sz - 1;
  ctx.save();
  ctx.fillStyle = '#D0D0D0';
  ctx.fillRect(x, y, sz, sz);
  ctx.fillStyle = '#444444';
  const tri = sz * 0.55;
  const tx = x + (sz - tri) / 2;
  const ty = y + (sz - tri * 0.5) / 2;
  ctx.beginPath();
  ctx.moveTo(tx, ty);
  ctx.lineTo(tx + tri, ty);
  ctx.lineTo(tx + tri / 2, ty + tri * 0.5);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
}

// ────────────────────────────────────────────────────────────────
// Excel Table style overlays (ECMA-376 §18.5)
// ────────────────────────────────────────────────────────────────
// We don't ship the full built-in table-style catalog — instead we derive a
// single "accent" color from the style name and overlay bold header + banded
// fills + horizontal rules so that `TableStyle*` files render with visible
// structure rather than as blank ranges.
export interface TableCellStyle {
  accent: string;
  isHeader: boolean;
  isTotals: boolean;
  /** `true` when this is a banded data row that should get the stripe fill. */
  isBanded: boolean;
  isFirstCol: boolean;
  isLastCol: boolean;
  isTopEdge: boolean;
  isBottomEdge: boolean;
  /** Dxf for the whole-table element of a custom `<tableStyle>`
   *  (ECMA-376 §18.8.40). Border/fill apply to every cell as a base layer. */
  wholeTableDxf?: number;
  /** Dxf for the header-row element of a custom `<tableStyle>`. Provides
   *  header fill, font color/weight, and vertical separators. */
  headerRowDxf?: number;
}

function buildTableStyleMap(worksheet: Worksheet): Map<string, TableCellStyle> {
  const map = new Map<string, TableCellStyle>();
  for (const t of worksheet.tables ?? []) {
    // ECMA-376 §18.5.1.4: empty styleName means the table has "None" style —
    // no visual table formatting overlay should be applied. Cell xf borders
    // and fills are still rendered as defined (per §18.8.45); only the
    // table-style overlay (banded fills, separator rules, etc.) is skipped.
    if (!t.styleName) continue;
    const { top, bottom, left, right } = t.range;
    const accent = t.accentColor || '#808080';
    const hdr = Math.max(0, t.headerRowCount ?? 1);
    const tot = Math.max(0, t.totalsRowCount ?? 0);
    const headerEnd = top + hdr - 1;
    const totalsStart = bottom - tot + 1;
    for (let r = top; r <= bottom; r++) {
      const isHeader = hdr > 0 && r <= headerEnd;
      const isTotals = tot > 0 && r >= totalsStart;
      const dataIdx = (!isHeader && !isTotals) ? (r - headerEnd - 1) : -1;
      for (let c = left; c <= right; c++) {
        map.set(`${r}:${c}`, {
          accent,
          isHeader,
          isTotals,
          isBanded: t.showRowStripes && dataIdx >= 0 && dataIdx % 2 === 1,
          isFirstCol: t.showFirstColumn && c === left,
          isLastCol: t.showLastColumn && c === right,
          isTopEdge: r === top,
          isBottomEdge: r === bottom,
          wholeTableDxf: t.wholeTableDxf,
          headerRowDxf: t.headerRowDxf,
        });
      }
    }
  }
  return map;
}

/** Flatten the worksheet's parsed `sparklineGroups` into a per-cell render
 *  model. Each Sparkline inherits its group's formatting; min/max are
 *  computed from the values when the group's `*AxisType` is `individual`,
 *  shared across the group when `group`, or taken from `manualMin/Max` when
 *  `custom`. The renderer can then look up `row:col` and call core's
 *  `renderSparkline` without further work. */
function buildSparklineMap(worksheet: Worksheet): Map<string, SparklineModel> {
  const map = new Map<string, SparklineModel>();
  for (const g of worksheet.sparklineGroups ?? []) {
    // Group-wide min/max if needed.
    let groupMin = Infinity, groupMax = -Infinity;
    if (g.minAxisType === 'group' || g.maxAxisType === 'group') {
      for (const sl of g.sparklines) {
        for (const v of sl.values) {
          if (typeof v === 'number') {
            if (v < groupMin) groupMin = v;
            if (v > groupMax) groupMax = v;
          }
        }
      }
      if (!isFinite(groupMin) || !isFinite(groupMax)) {
        groupMin = 0; groupMax = 1;
      }
    }
    for (const sl of g.sparklines) {
      const numeric = sl.values.filter((v): v is number => typeof v === 'number');
      const indMin = numeric.length ? Math.min(...numeric) : 0;
      const indMax = numeric.length ? Math.max(...numeric) : 1;
      const min = g.minAxisType === 'custom' && typeof g.manualMin === 'number'
        ? g.manualMin
        : g.minAxisType === 'group' ? groupMin : indMin;
      const max = g.maxAxisType === 'custom' && typeof g.manualMax === 'number'
        ? g.manualMax
        : g.maxAxisType === 'group' ? groupMax : indMax;
      map.set(`${sl.row}:${sl.col}`, {
        kind: g.kind,
        values: sl.values,
        min,
        max,
        displayEmptyCellsAs: (g.displayEmptyCellsAs === 'zero' || g.displayEmptyCellsAs === 'span')
          ? g.displayEmptyCellsAs
          : 'gap',
        displayXAxis: g.displayXAxis,
        lineWeight: g.lineWeight,
        markers: g.markers,
        high: g.high,
        low: g.low,
        first: g.first,
        last: g.last,
        negative: g.negative,
        colorSeries: g.colorSeries,
        colorNegative: g.colorNegative,
        colorAxis: g.colorAxis,
        colorMarkers: g.colorMarkers,
        colorFirst: g.colorFirst,
        colorLast: g.colorLast,
        colorHigh: g.colorHigh,
        colorLow: g.colorLow,
      });
    }
  }
  return map;
}

function stripeColorFor(accent: string): string {
  // Light tint of the accent — mimics TableStyleLight* banded rows.
  const hex = accent.replace('#', '');
  if (hex.length < 6) return '#F2F2F2';
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const mix = (ch: number) => Math.round(ch * 0.2 + 255 * 0.8);
  const toHex = (v: number) => v.toString(16).padStart(2, '0').toUpperCase();
  return `#${toHex(mix(r))}${toHex(mix(g))}${toHex(mix(b))}`;
}

// ────────────────────────────────────────────────────────────────
// Render one rectangular region of cells
// ────────────────────────────────────────────────────────────────
function renderQuadrant(
  ctx: CanvasRenderingContext2D,
  rc: RenderContext,
  startRow: number, startCol: number,
  colWidths: number[], rowHeights: number[],
  pixOffsetX: number, pixOffsetY: number,
  originX: number, originY: number,
  clipX: number, clipY: number, clipW: number, clipH: number,
): void {
  if (clipW <= 0 || clipH <= 0) return;

  const { styles, cellMap, mergeAnchorMap, mergeSkipSet, cfContext, cs, dpr } = rc;
  const numCols = colWidths.length;
  const numRows = rowHeights.length;

  // Canvas x for each column
  const colXs: number[] = [];
  let x = -pixOffsetX;
  for (let ci = 0; ci < numCols; ci++) {
    colXs.push(x);
    x += colWidths[ci];
  }

  // Canvas y for each row
  const rowYs: number[] = [];
  let y = -pixOffsetY;
  for (let ri = 0; ri < numRows; ri++) {
    rowYs.push(y);
    y += rowHeights[ri];
  }

  ctx.save();
  ctx.beginPath();
  ctx.rect(clipX, clipY, clipW, clipH);
  ctx.clip();

  // Deferred text drawing. Excel renders cell text on top of *all* cell
  // backgrounds, so a left-aligned overflow into an adjacent empty cell with
  // a white fill stays visible. If we drew fill+text per cell in one pass,
  // the next cell's fill would overpaint the previous cell's overflow.
  const textTasks: Array<() => void> = [];

  // Pre-pass: merge cells whose anchor lies outside this viewport quadrant but whose
  // span overlaps it (e.g. scrolled past the anchor row/col, or the anchor is in a
  // frozen quadrant while we are rendering the scrollable quadrant).
  for (const mc of rc.worksheet.mergeCells ?? []) {
    const aRow = mc.top, aCol = mc.left;
    // If anchor is within the main loop range, skip — handled normally below.
    if (aRow >= startRow && aRow < startRow + numRows &&
        aCol >= startCol && aCol < startCol + numCols) continue;
    // Skip if merge span has no overlap with this viewport.
    if (mc.bottom < startRow || mc.top >= startRow + numRows) continue;
    if (mc.right  < startCol || mc.left >= startCol + numCols) continue;

    const info = rc.mergeAnchorMap.get(`${aRow}:${aCol}`);
    if (!info) continue;

    // Canvas X of anchor col (may be negative = off-screen to the left).
    let aCx: number;
    if (aCol >= startCol) {
      aCx = originX + colXs[aCol - startCol];
    } else {
      let dx = 0;
      for (let c = aCol; c < startCol; c++) {
        dx += Math.round(colWidthToPx(rc.worksheet.colWidths[c] ?? rc.worksheet.defaultColWidth, rc.mdw) * cs);
      }
      aCx = originX - pixOffsetX - dx;
    }
    // Canvas Y of anchor row (may be negative = off-screen above).
    let aCy: number;
    if (aRow >= startRow) {
      aCy = originY + rowYs[aRow - startRow];
    } else {
      let dy = 0;
      for (let r = aRow; r < startRow; r++) {
        dy += Math.round(rowHeightToPx(rc.worksheet.rowHeights[r] ?? rc.worksheet.defaultRowHeight) * cs);
      }
      aCy = originY - pixOffsetY - dy;
    }

    const cW = info.totalW, cH = info.totalH;
    const key = `${aRow}:${aCol}`;
    const cell = rc.cellMap.get(key);
    const { font, fill, border, xf } = resolveXf(styles, cell?.styleIndex ?? 0);
    const cf = evaluateCf(cell, aRow, aCol, cfContext, styles.dxfs ?? []);
    const effectiveFill = cf.fill ?? fill;

    paintCellPatternFill(ctx, effectiveFill, aCx, aCy, cW, cH);
    if (cf.dataBar && cf.dataBar.ratio > 0) {
      const bInset = 2;
      const bW = Math.max(0, (cW - bInset * 2) * cf.dataBar.ratio);
      fillDataBar(ctx, cf.dataBar.color, aCx + bInset, aCy + bInset, bW, cH - bInset * 2, cf.dataBar.gradient);
    }
    const mergedBorder = resolveMergeBorder(border, aRow, aCol, info.right, info.bottom, rc.cellMap, styles);
    renderBorder(ctx, mergeBorders(mergedBorder, cf.border), aCx, aCy, cW, cH);

    if (!cell) continue;
    const text = formatCellValue(cell, styles, cf.numFmt);
    if (!text || (text === '0' && rc.worksheet.showZeros === false)) continue;

    const effectiveBold = font.bold || !!cf.fontBold;
    const effectiveItalic = font.italic || !!cf.fontItalic;
    const effectiveUnderline = font.underline || !!cf.fontUnderline;
    const effectiveStrike = font.strike || !!cf.fontStrike;
    const fontForDraw: Font = (
      effectiveBold !== font.bold || effectiveItalic !== font.italic ||
      effectiveUnderline !== font.underline || effectiveStrike !== font.strike
    ) ? { ...font, bold: effectiveBold, italic: effectiveItalic, underline: effectiveUnderline, strike: effectiveStrike }
      : font;
    ctx.font = buildFont(fontForDraw, cs);
    const hyperlinkUrl = rc.hyperlinkMap.get(key);
    const textColor = hyperlinkUrl ? '#0563C1' : (cf.fontColor ?? font.color);
    ctx.fillStyle = textColor ? hexToRgba(textColor) : '#000000';

    const paddingX = 3, paddingY = 2;
    const isNumeric = cell.value.type === 'number';
    const alignH = xf.alignH ?? (isNumeric ? 'right' : 'left');
    const alignV = xf.alignV ?? 'bottom';
    const indentPx = xf.indent ? Math.round(xf.indent * font.size * PT_TO_PX * 0.5) : 0;
    const leftPad = paddingX + (alignH === 'left' || !xf.alignH ? indentPx : 0);

    ctx.save();
    ctx.beginPath();
    ctx.rect(aCx, aCy, cW, cH);
    ctx.clip();

    let textX: number;
    if (alignH === 'right') { textX = aCx + cW - paddingX; ctx.textAlign = 'right'; }
    else if (alignH === 'center') { textX = aCx + cW / 2; ctx.textAlign = 'center'; }
    else { textX = aCx + leftPad; ctx.textAlign = 'left'; }

    let textY: number;
    if (alignV === 'top') { ctx.textBaseline = 'top'; textY = aCy + paddingY; }
    else if (alignV === 'center') { ctx.textBaseline = 'middle'; textY = aCy + cH / 2; }
    else { ctx.textBaseline = 'bottom'; textY = aCy + cH - paddingY; }

    ctx.fillText(text, textX, textY);
    ctx.restore();
  }

  for (let ri = 0; ri < numRows; ri++) {
    const rowIndex = startRow + ri;
    const cy = originY + rowYs[ri];
    const ch = rowHeights[ri];
    if (cy + ch <= clipY || cy >= clipY + clipH) continue;

    // Pre-compute centerContinuous ranges in this row. ECMA-376 §18.18.40:
    // a contiguous run of cells whose alignment is `centerContinuous` is
    // treated as a single visual span — Excel hides the default gridlines
    // *and* the explicit cell borders inside the run, so the run reads as
    // one merged-looking span with only the outer perimeter visible.
    //
    // A run is bounded by either (a) a non-centerContinuous cell, or
    // (b) another centerContinuous cell that itself carries a value —
    // the spec says spanned cells reference the same style id while only
    // the anchor holds the displayed value. Two anchors in a row therefore
    // mean two adjacent runs, with the border between them visible.
    const suppressRightGridCol = new Set<number>();
    const suppressLeftGridCol = new Set<number>();
    let runStart = -1;
    const closeRun = (endExclusive: number) => {
      if (runStart >= 0 && endExclusive - runStart >= 2) {
        for (let k = runStart; k < endExclusive - 1; k++) suppressRightGridCol.add(k);
        for (let k = runStart + 1; k < endExclusive; k++) suppressLeftGridCol.add(k);
      }
      runStart = -1;
    };
    for (let ci = 0; ci <= numCols; ci++) {
      let isCC = false;
      let hasValue = false;
      if (ci < numCols) {
        const ckey = `${rowIndex}:${startCol + ci}`;
        if (!mergeSkipSet.has(ckey) && !mergeAnchorMap.has(ckey)) {
          const c = cellMap.get(ckey);
          const cXf = resolveXf(styles, c?.styleIndex ?? 0).xf;
          isCC = cXf.alignH === 'centerContinuous';
          hasValue = !!(c && c.value && c.value.type !== 'empty');
        }
      }
      if (!isCC) {
        closeRun(ci);
      } else if (hasValue && runStart >= 0 && ci > runStart) {
        // New anchor: close previous run [runStart, ci) and start a fresh one.
        closeRun(ci);
        runStart = ci;
      } else if (runStart < 0) {
        runStart = ci;
      }
    }

    for (let ci = 0; ci < numCols; ci++) {
      const colIndex = startCol + ci;
      const cx = originX + colXs[ci];
      const cw = colWidths[ci];
      if (cx + cw <= clipX || cx >= clipX + clipW) continue;

      const key = `${rowIndex}:${colIndex}`;
      if (mergeSkipSet.has(key)) continue;

      const mergeInfo = mergeAnchorMap.get(key);
      const cellW = mergeInfo ? mergeInfo.totalW : cw;
      const cellH = mergeInfo ? mergeInfo.totalH : ch;

      const cell = cellMap.get(key);
      const styleIndex = cell?.styleIndex ?? 0;
      const { font, fill, border, xf } = resolveXf(styles, styleIndex);
      const cf = evaluateCf(cell, rowIndex, colIndex, cfContext, styles.dxfs ?? []);
      const effectiveFill = cf.fill ?? fill;
      const tableStyle = rc.tableStyleMap.get(key);
      // Custom `<tableStyle>` dxfs (ECMA-376 §18.8.40). When present, they
      // drive header fill / font color and inter-row borders instead of the
      // built-in accent fallback.
      const tsDxfWhole = (tableStyle?.wholeTableDxf != null)
        ? (styles.dxfs ?? [])[tableStyle.wholeTableDxf] : undefined;
      const tsDxfHeader = (tableStyle?.headerRowDxf != null)
        ? (styles.dxfs ?? [])[tableStyle.headerRowDxf] : undefined;

      // Background fill (base or CF override). ECMA-376 §18.8.22 ST_PatternType.
      // - solid/gray*: blend fgColor with bgColor at the pattern's fg coverage.
      // - directional hatches (dark/light Horizontal/Vertical/Down/Up/Grid/
      //   Trellis): render via a small repeating tile using createPattern so
      //   the hatch actually shows, rather than approximating as a blend.
      if (paintCellPatternFill(ctx, effectiveFill, cx, cy, cellW, cellH)) {
        // own fill painted; tableStyle fallbacks intentionally skipped
      } else if (tableStyle && tableStyle.isHeader && tsDxfHeader?.fill?.fgColor) {
        ctx.fillStyle = hexToRgba(tsDxfHeader.fill.fgColor);
        ctx.fillRect(cx, cy, cellW, cellH);
      } else if (tableStyle && !tableStyle.isHeader && !tableStyle.isTotals && tsDxfWhole?.fill?.fgColor) {
        ctx.fillStyle = hexToRgba(tsDxfWhole.fill.fgColor);
        ctx.fillRect(cx, cy, cellW, cellH);
      } else if (tableStyle && tableStyle.isBanded) {
        ctx.fillStyle = stripeColorFor(tableStyle.accent);
        ctx.fillRect(cx, cy, cellW, cellH);
      }

      // Comment indicator triangle — drawn above fill but below borders so
      // borders still read cleanly around the cell edge.
      if (rc.commentCells.has(key)) {
        drawCommentMarker(ctx, cx, cy, cellW, cellH);
      }

      // DataBar (drawn inside the cell, left-anchored). Excel 2010+ renders
      // these with a horizontal gradient by default.
      if (cf.dataBar && cf.dataBar.ratio > 0) {
        const barInset = 2;
        const barW = Math.max(0, (cellW - barInset * 2) * cf.dataBar.ratio);
        fillDataBar(ctx, cf.dataBar.color, cx + barInset, cy + barInset, barW, cellH - barInset * 2, cf.dataBar.gradient);
      }

      // Sparkline (Office 2010 x14:sparklineGroup). Drawn after the cell
      // background but before borders / text so borders frame the sparkline
      // and any cell text overlays it (matches Excel's z-order, and lets
      // a label like "Trend" share the same cell as the sparkline).
      const sparkModel = rc.sparklineMap.get(key);
      if (sparkModel) {
        renderSparkline(ctx, { x: cx, y: cy, w: cellW, h: cellH }, sparkModel);
      }

      // Grid lines – draw only right + bottom edges once per cell (avoids double-drawing at
      // shared cell boundaries). Half-device-pixel offset (0.5/dpr) aligns each line to the
      // device pixel grid so we get a crisp 1-device-pixel result.
      // Skipped when the sheet has `<sheetView showGridLines="0">` (View →
      // Gridlines unchecked; ECMA-376 §18.3.1.83).
      if (rc.worksheet.showGridlines !== false) {
        const hp = 0.5 / dpr;
        ctx.strokeStyle = '#d0d0d0';
        ctx.lineWidth = 0.5;
        ctx.beginPath();
        if (!suppressRightGridCol.has(ci)) {
          ctx.moveTo(cx + cellW + hp, cy);        // right edge
          ctx.lineTo(cx + cellW + hp, cy + cellH);
        }
        ctx.moveTo(cx, cy + cellH + hp);           // bottom edge
        ctx.lineTo(cx + cellW, cy + cellH + hp);
        if (ri === 0) {                            // top edge for first row
          ctx.moveTo(cx, cy + hp);
          ctx.lineTo(cx + cellW, cy + hp);
        }
        if (ci === 0) {                            // left edge for first column
          ctx.moveTo(cx + hp, cy);
          ctx.lineTo(cx + hp, cy + cellH);
        }
        ctx.stroke();
      }

      // Cell borders (base + any CF borders overlaid via per-edge merge).
      // For merged anchors, combine the anchor's border with the right/bottom
      // edges from the constituent cells on those edges (ECMA-376 §18.3.1.55).
      const baseBorder = mergeInfo
        ? resolveMergeBorder(border, rowIndex, colIndex, mergeInfo.right, mergeInfo.bottom, cellMap, styles)
        : border;
      let mergedBorder = mergeBorders(baseBorder, cf.border);
      // centerContinuous: hide internal vertical borders so the run reads as
      // one visual span (matches Excel — see precompute block above).
      if (suppressRightGridCol.has(ci) || suppressLeftGridCol.has(ci)) {
        mergedBorder = {
          ...mergedBorder,
          left: suppressLeftGridCol.has(ci) ? null : mergedBorder.left,
          right: suppressRightGridCol.has(ci) ? null : mergedBorder.right,
        };
      }
      // Inherit the cell-above's bottom edge as our top, and the cell-left's
      // right edge as our left. Two adjacent cells share an edge along the row
      // / column boundary; the upper cell drew its bottom during its own
      // iteration, but our cell's fill (drawn just above) over-paints the
      // half of that line lying inside our cell. Re-drawing it as our top
      // (after our fill) restores the boundary line.
      //
      // When *both* cells define an edge, Excel renders the stronger style at
      // a conflict (e.g. a medium bottom is not erased by a thin top below
      // it). `pickStrongerEdge` returns the higher-precedence edge so the
      // inherit picks the visually dominant style instead of always favouring
      // the lower cell's own.
      const aboveCell = cellMap.get(`${rowIndex - 1}:${colIndex}`);
      const aboveBottom = aboveCell
        ? resolveXf(styles, aboveCell.styleIndex).border.bottom
        : null;
      let invertedTop = false;
      if (aboveBottom?.style) {
        const before = mergedBorder.top;
        const picked = pickStrongerEdge(before, aboveBottom);
        mergedBorder = { ...mergedBorder, top: picked };
        // We only invert the double-border drawing when the top edge ends up
        // being a *redraw* of the upper cell's bottom (rather than a fresh
        // line for our own xf.top). That is true when the picked edge came
        // from the neighbour — i.e. our own top was unset or weaker.
        invertedTop = picked === aboveBottom && before !== aboveBottom;
      }
      let invertedLeft = false;
      if (!suppressLeftGridCol.has(ci)) {
        // Skip the inherit when the left edge was deliberately suppressed for
        // a centerContinuous run (ECMA-376 §18.18.40) — otherwise the
        // neighbour's xf.right re-introduces the internal vertical that we
        // just hid.
        const leftCell = cellMap.get(`${rowIndex}:${colIndex - 1}`);
        const leftRight = leftCell
          ? resolveXf(styles, leftCell.styleIndex).border.right
          : null;
        if (leftRight?.style) {
          const before = mergedBorder.left;
          const picked = pickStrongerEdge(before, leftRight);
          mergedBorder = { ...mergedBorder, left: picked };
          invertedLeft = picked === leftRight && before !== leftRight;
        }
      }
      renderBorder(ctx, mergedBorder, cx, cy, cellW, cellH, invertedTop, invertedLeft);

      // Excel Table style overlay: thin horizontal rules between rows and a
      // thicker bottom edge under the header row (ECMA-376 §18.5). Drawn on
      // top of cell borders so an empty-border data cell still shows table
      // structure. None-style tables produce no entry in `tableStyleMap`
      // (see `buildTableStyleMap`), so this block is naturally skipped.
      if (tableStyle) {
        const horiz = tsDxfWhole?.border?.horizontal;
        const vert  = tsDxfWhole?.border?.vertical;
        const wtTop = tsDxfWhole?.border?.top;
        const wtBot = tsDxfWhole?.border?.bottom;
        const wtLeft = tsDxfWhole?.border?.left;
        const wtRight = tsDxfWhole?.border?.right;
        const hdrBot = tsDxfHeader?.border?.bottom;
        const hdrTop = tsDxfHeader?.border?.top;
        const hasDxfBorder = !!(horiz || vert || wtTop || wtBot || wtLeft || wtRight || hdrBot || hdrTop);
        if (hasDxfBorder) {
          const overlay: Border = { left: null, right: null, top: null, bottom: null };
          if (tableStyle.isTopEdge) overlay.top = wtTop ?? null;
          else if (horiz) overlay.top = horiz;
          if (tableStyle.isHeader && hdrBot) overlay.bottom = hdrBot;
          else if (tableStyle.isBottomEdge) overlay.bottom = wtBot ?? null;
          else if (horiz) overlay.bottom = horiz;
          if (tableStyle.isFirstCol || colIndex === 0) overlay.left = wtLeft ?? null;
          if (tableStyle.isLastCol) overlay.right = wtRight ?? null;
          // Outer table left/right edges
          renderBorder(ctx, overlay, cx, cy, cellW, cellH);
        } else {
          const hp = 0.5 / dpr;
          ctx.strokeStyle = tableStyle.accent;
          ctx.lineWidth = tableStyle.isHeader ? 1.5 : 1;
          ctx.beginPath();
          ctx.moveTo(cx, cy + cellH - hp);
          ctx.lineTo(cx + cellW, cy + cellH - hp);
          if (tableStyle.isTopEdge) {
            ctx.moveTo(cx, cy + hp);
            ctx.lineTo(cx + cellW, cy + hp);
          }
          ctx.stroke();
        }
      }

      // AutoFilter dropdown indicator
      if (rc.autoFilterCells.has(key)) {
        drawAutoFilterArrow(ctx, cx, cy, cw, cellH);
      }

      if (!cell) continue;
      const text = formatCellValue(cell, styles, cf.numFmt);
      if (!text || (text === '0' && rc.worksheet.showZeros === false)) continue;

      textTasks.push(() => {
      const tableBold = !!(tableStyle && (tableStyle.isHeader || tableStyle.isTotals));
      const effectiveBold = font.bold || !!cf.fontBold || tableBold;
      const effectiveItalic = font.italic || !!cf.fontItalic;
      const effectiveUnderline = font.underline || !!cf.fontUnderline;
      const effectiveStrike = font.strike || !!cf.fontStrike;
      const fontForDraw: Font = (
        effectiveBold !== font.bold || effectiveItalic !== font.italic ||
        effectiveUnderline !== font.underline || effectiveStrike !== font.strike
      )
        ? { ...font, bold: effectiveBold, italic: effectiveItalic, underline: effectiveUnderline, strike: effectiveStrike }
        : font;
      ctx.font = buildFont(fontForDraw, cs);
      const hyperlinkUrl = rc.hyperlinkMap.get(key);
      // Custom table-style header dxfs can override font color (ECMA-376 §18.8.40).
      const tableFontColor =
        (tableStyle?.isHeader && tsDxfHeader?.font?.color) ? tsDxfHeader.font.color :
        (tableStyle && !tableStyle.isHeader && !tableStyle.isTotals && tsDxfWhole?.font?.color) ? tsDxfWhole.font.color :
        null;
      const textColor = hyperlinkUrl
        ? '#0563C1'
        : (cf.fontColor ?? tableFontColor ?? font.color);
      ctx.fillStyle = textColor ? hexToRgba(textColor) : '#000000';

      const paddingX = 3;
      const paddingY = 2;
      const isNumeric = cell.value.type === 'number';
      const alignH = xf.alignH ?? (isNumeric ? 'right' : 'left');
      const alignV = xf.alignV ?? 'bottom';
      // Indent: each level ≈ one character width (ECMA-376 §18.8.44)
      const indentPx = xf.indent ? Math.round(xf.indent * font.size * PT_TO_PX * 0.5) : 0;
      // IconSet: reserve space on the left for the icon
      const iconSz = cf.iconSet ? Math.max(8, Math.round(Math.min(cellW, cellH) * 0.55)) : 0;
      const iconPad = iconSz > 0 ? iconSz + 4 : 0;
      const leftPad = paddingX + (alignH === 'left' || !xf.alignH ? indentPx : 0) + iconPad;

      // ECMA-376 §18.18.40 ST_HorizontalAlignment value `centerContinuous`:
      // text is centered across the **selection range** — the leftmost cell
      // with content + adjacent empty cells to the right that also carry
      // `centerContinuous`. Cells stay independent (unlike merge), but the
      // centering uses the combined width. Walk right collecting empty
      // centerContinuous neighbours so we know how wide to center over.
      let centerContinuousW = cellW;
      let centerContinuousX = cx;
      let centerContinuousLastCi = ci;
      if (alignH === 'centerContinuous' && !mergeInfo) {
        for (let oci = ci + 1; oci < numCols; oci++) {
          const adjKey = `${rowIndex}:${startCol + oci}`;
          if (mergeSkipSet.has(adjKey) || mergeAnchorMap.has(adjKey)) break;
          const adjCell = cellMap.get(adjKey);
          if (adjCell && adjCell.value.type !== 'empty') break;
          const adjStyleIndex = adjCell?.styleIndex ?? 0;
          const adjXf = resolveXf(styles, adjStyleIndex).xf;
          if (adjXf.alignH !== 'centerContinuous') break;
          centerContinuousW += colWidths[oci];
          centerContinuousLastCi = oci;
        }
      }

      // Text overflow into adjacent empty cells (ECMA-376 §18.3.1.4 "spans"
      // — Excel behavior when `wrapText=false`). Left-aligned text flows
      // rightward, right-aligned flows leftward, and centered splits evenly.
      // We only extend the clip rect; the text itself is still drawn once.
      // Stops at merge-cell boundaries, non-empty cells, and iconSet-left
      // overrun (since an icon sits inside this cell's left padding).
      let drawX = alignH === 'centerContinuous' ? centerContinuousX : cx;
      let drawW = alignH === 'centerContinuous' ? centerContinuousW : cellW;
      // Excel only overflows text into adjacent empty cells; numeric values
      // that don't fit are rendered as "####" (they never spill). Cells
      // containing hard line breaks render multi-line in place, so they don't
      // overflow either. centerContinuous overflows symmetrically across the
      // selection range — text stays centered on the original range, but the
      // clip rect extends into adjacent empty cells when the text is wider.
      const hasHardBreak = text.includes('\n');
      if (!mergeInfo && !xf.wrapText && !xf.textRotation && !isNumeric && !hasHardBreak) {
        const textW = ctx.measureText(text).width;
        const isCenterCont = alignH === 'centerContinuous';
        const textPx = isCenterCont ? textW + paddingX * 2 : textW + leftPad + paddingX;
        const containerW = isCenterCont ? centerContinuousW : cellW;
        if (textPx > containerW) {
          const overflow = textPx - containerW;
          let extendRight = 0;
          let extendLeft = 0;
          if (alignH === 'right') {
            extendLeft = overflow;
          } else if (alignH === 'center' || isCenterCont) {
            extendLeft = overflow / 2;
            extendRight = overflow / 2;
          } else {
            extendRight = overflow;
          }
          if (extendRight > 0) {
            let budget = extendRight;
            const startOci = isCenterCont ? centerContinuousLastCi + 1 : ci + 1;
            for (let oci = startOci; oci < numCols && budget > 0; oci++) {
              const adjKey = `${rowIndex}:${startCol + oci}`;
              if (mergeSkipSet.has(adjKey) || mergeAnchorMap.has(adjKey)) break;
              const adjCell = cellMap.get(adjKey);
              if (adjCell && adjCell.value.type !== 'empty') break;
              drawW += colWidths[oci];
              budget -= colWidths[oci];
            }
          }
          if (extendLeft > 0) {
            let budget = extendLeft;
            for (let oci = ci - 1; oci >= 0 && budget > 0; oci--) {
              const adjKey = `${rowIndex}:${startCol + oci}`;
              if (mergeSkipSet.has(adjKey) || mergeAnchorMap.has(adjKey)) break;
              const adjCell = cellMap.get(adjKey);
              if (adjCell && adjCell.value.type !== 'empty') break;
              drawX -= colWidths[oci];
              drawW += colWidths[oci];
              budget -= colWidths[oci];
            }
          }
        }
      }

      // ECMA-376 §18.18.40 horizontal alignment:
      // - `fill`: repeat the cell content until the cell fills horizontally.
      //   We resolve to the repeated string here; layout is then standard
      //   left-aligned.
      // - `distributed`: distribute characters evenly across the cell width
      //   so the first char hugs the left edge and the last hugs the right.
      //   Implemented via Canvas `letterSpacing` (a string CSS length).
      // - `justify`: full-line justification. Multi-line distribution would
      //   need the wrap engine; for the single-line case we fall back to
      //   `distributed` if the text is short, else left-aligned (matches
      //   what Excel does when justify cannot expand a one-line string).
      // Anything we don't recognise falls through to left-align (general).
      let drawText = text;
      let letterSpacingPx = 0;
      if (alignH === 'fill' && !isNumeric && text.length > 0) {
        const innerW = Math.max(1, cellW - paddingX * 2);
        const oneW = ctx.measureText(text).width;
        if (oneW > 0 && oneW < innerW) {
          const reps = Math.max(1, Math.floor(innerW / oneW));
          drawText = text.repeat(reps);
        }
      }
      if (alignH === 'distributed' || (alignH === 'justify' && !xf.wrapText && !hasHardBreak)) {
        const innerW = Math.max(1, cellW - paddingX * 2);
        const naturalW = ctx.measureText(drawText).width;
        const gaps = Math.max(1, [...drawText].length - 1);
        if (naturalW < innerW) {
          letterSpacingPx = Math.max(0, (innerW - naturalW) / gaps);
        }
      }

      let textX: number;
      let textAlign: CanvasTextAlign;
      if (alignH === 'right') {
        textX = cx + cellW - paddingX;
        textAlign = 'right';
      } else if (alignH === 'center') {
        textX = cx + cellW / 2;
        textAlign = 'center';
      } else if (alignH === 'centerContinuous') {
        // Center on the original selection range; drawX/drawW already covers
        // the range (and any overflow extension into empty neighbours).
        textX = centerContinuousX + centerContinuousW / 2;
        textAlign = 'center';
      } else if (alignH === 'distributed' || (alignH === 'justify' && !xf.wrapText && !hasHardBreak)) {
        // Distribute characters evenly: first hugs left edge, last hugs right.
        textX = cx + paddingX;
        textAlign = 'left';
      } else {
        textX = cx + leftPad;
        textAlign = 'left';
      }

      const rotation = xf.textRotation ?? 0;
      const isStacked = rotation === 255;
      const isRotated = rotation > 0 && rotation !== 255;

      // Draw icon set icon (before text clip block)
      if (cf.iconSet && iconSz > 0) {
        ctx.save();
        ctx.beginPath();
        ctx.rect(cx, cy, cellW, cellH);
        ctx.clip();
        drawCfIcon(ctx, cf.iconSet.name, cf.iconSet.index, cx + 2, cy + (cellH - iconSz) / 2, iconSz);
        ctx.restore();
      }

      ctx.save();
      ctx.beginPath();
      ctx.rect(drawX, cy, drawW, cellH);
      ctx.clip();

      // Stacked text (textRotation=255): draw each character on its own line
      if (isStacked) {
        const charH = Math.round(font.size * PT_TO_PX * 1.1);
        const totalH = text.length * charH;
        let charY = alignV === 'top' ? cy + paddingY
          : alignV === 'center' ? cy + (cellH - totalH) / 2
          : cy + cellH - totalH - paddingY;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'top';
        for (const ch of text) {
          ctx.fillText(ch, cx + cellW / 2, charY);
          charY += charH;
        }
        ctx.restore();
        return;
      }

      // Rotated text: translate to cell center, rotate, draw, restore
      if (isRotated) {
        const angleRad = rotation <= 90
          ? -(rotation * Math.PI / 180)
          : ((rotation - 90) * Math.PI / 180);
        ctx.translate(cx + cellW / 2, cy + cellH / 2);
        ctx.rotate(angleRad);
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(text, 0, 0);
        ctx.restore();
        return;
      }

      // shrinkToFit: scale context horizontally if text is wider than cell
      if (xf.shrinkToFit) {
        const textW = ctx.measureText(text).width;
        const availW = cellW - leftPad - paddingX;
        if (textW > availW && textW > 0) {
          const scale = availW / textW;
          const pivotX = alignH === 'right' ? cx + cellW - paddingX
            : alignH === 'center' ? cx + cellW / 2
            : cx + leftPad;
          ctx.transform(scale, 0, 0, 1, pivotX * (1 - scale), 0);
        }
      }

      ctx.textAlign = textAlign;
      // ECMA-376 §18.18.39 distributed / §18.18.40 readingOrder. Canvas
      // `letterSpacing` (CSS length string, supported in modern browsers)
      // implements distributed by stretching the gap between glyphs;
      // `direction` flips the writing direction for RTL languages.
      if (letterSpacingPx > 0) {
        try { (ctx as CanvasRenderingContext2D & { letterSpacing: string }).letterSpacing = `${letterSpacingPx}px`; } catch { /* ignore */ }
      }
      if (xf.readingOrder === 2) {
        try { (ctx as CanvasRenderingContext2D & { direction: 'ltr' | 'rtl' }).direction = 'rtl'; } catch { /* ignore */ }
      } else if (xf.readingOrder === 1) {
        try { (ctx as CanvasRenderingContext2D & { direction: 'ltr' | 'rtl' }).direction = 'ltr'; } catch { /* ignore */ }
      }

      // Rich text: draw each run with its own font. Only supported for the
      // non-wrap path (wrap with mixed fonts is significantly more complex).
      const runs = cell.value.type === 'text' ? cell.value.runs : undefined;
      const hasRichText = runs && runs.length > 0;

      if (xf.wrapText && hasRichText) {
        // Rich text with wrapping: per-run fonts, break on spaces and CJK boundaries
        const wrapW = cellW - leftPad - paddingX;
        const rLines = layoutRichTextLines(ctx, runs, fontForDraw, cs, wrapW);
        const totalH = rLines.reduce((s, l) => s + Math.round(l.maxFontSize * PT_TO_PX * 1.2), 0);
        let yy: number;
        if (alignV === 'top') yy = cy + paddingY;
        else if (alignV === 'center') yy = cy + (cellH - totalH) / 2;
        else yy = cy + cellH - totalH - paddingY;
        ctx.textAlign = 'left';
        ctx.textBaseline = 'top';
        for (const line of rLines) {
          const lineH = Math.round(line.maxFontSize * PT_TO_PX * 1.2);
          const totalW = line.segments.reduce((s, seg) => s + seg.width, 0);
          let xx: number;
          if (alignH === 'right') xx = cx + cellW - paddingX - totalW;
          else if (alignH === 'center') xx = cx + cellW / 2 - totalW / 2;
          else xx = cx + leftPad;
          for (const seg of line.segments) {
            ctx.font = buildFont(seg.font, cs);
            const segColor = cf.fontColor ?? seg.font.color;
            ctx.fillStyle = segColor ? hexToRgba(segColor) : '#000000';
            ctx.fillText(seg.text, xx, yy);
            const rSizePx = Math.round(seg.font.size * PT_TO_PX);
            if (seg.font.underline) {
              const stroke = segColor ? hexToRgba(segColor) : '#000000';
              const dbl = seg.font.underlineStyle === 'double' || seg.font.underlineStyle === 'doubleAccounting';
              drawTextDecoLine(ctx, xx, xx + seg.width, yy + rSizePx + 1, stroke, dbl);
            }
            if (seg.font.strike) {
              ctx.save();
              ctx.strokeStyle = segColor ? hexToRgba(segColor) : '#000000';
              ctx.lineWidth = 0.5;
              const sy2 = yy + Math.round(rSizePx * 0.5);
              ctx.beginPath(); ctx.moveTo(xx, sy2); ctx.lineTo(xx + seg.width, sy2); ctx.stroke();
              ctx.restore();
            }
            xx += seg.width;
          }
          yy += lineH;
        }
      } else if (xf.wrapText) {
        const lines = wrapTextLines(ctx, text, cellW - leftPad - paddingX);
        const lineH = Math.round(font.size * PT_TO_PX * 1.2);
        const totalTextH = lines.length * lineH;
        let startY: number;
        if (alignV === 'top') { startY = cy + paddingY; ctx.textBaseline = 'top'; }
        else if (alignV === 'center') { startY = cy + (cellH - totalTextH) / 2; ctx.textBaseline = 'top'; }
        else { startY = cy + cellH - totalTextH - paddingY; ctx.textBaseline = 'top'; }
        for (let li = 0; li < lines.length; li++) {
          ctx.fillText(lines[li], textX, startY + li * lineH);
        }
      } else if (hasRichText) {
        // Per-run drawing: compute font for each run, measure widths, draw LTR.
        // Layout uses the run's *base* font size (line height & x-position
        // budget); super/subscript runs are rendered at ~65% size with a
        // baseline shift, matching how Excel paints them. ECMA-376 §18.4.6
        // (ST_VerticalAlignRun) leaves the exact ratio implementation-defined.
        const baseRunFonts = runs.map(r => applyRunFont(fontForDraw, r));
        const runVAlign = runs.map(r => r.font?.vertAlign);
        const drawRunFonts = baseRunFonts.map((f, i) => {
          if (runVAlign[i] === 'superscript' || runVAlign[i] === 'subscript') {
            return { ...f, size: f.size * 0.65 };
          }
          return f;
        });
        const runWidths: number[] = runs.map((r, i) => {
          ctx.font = buildFont(drawRunFonts[i], cs);
          return ctx.measureText(r.text).width;
        });
        const totalWidth = runWidths.reduce((a, b) => a + b, 0);
        let startX: number;
        if (alignH === 'right') startX = cx + cellW - paddingX - totalWidth;
        else if (alignH === 'center') startX = cx + cellW / 2 - totalWidth / 2;
        else startX = cx + leftPad;
        // Use left alignment since we position each run ourselves
        ctx.textAlign = 'left';
        let textY: number;
        if (alignV === 'top') { ctx.textBaseline = 'top'; textY = cy + paddingY; }
        else if (alignV === 'center') { ctx.textBaseline = 'middle'; textY = cy + cellH / 2; }
        else { ctx.textBaseline = 'bottom'; textY = cy + cellH - paddingY; }
        let runX = startX;
        for (let i = 0; i < runs.length; i++) {
          const rf = drawRunFonts[i];
          const baseRf = baseRunFonts[i];
          ctx.font = buildFont(rf, cs);
          const runColor = cf.fontColor ?? rf.color;
          ctx.fillStyle = runColor ? hexToRgba(runColor) : '#000000';
          // Baseline shift for super/subscript. With textBaseline 'bottom'
          // (the typical case) shift up for super and slightly down for sub
          // so each run sits at the right vertical band relative to the line.
          const baseSizePx = Math.round(baseRf.size * PT_TO_PX);
          let yShift = 0;
          if (runVAlign[i] === 'superscript') yShift = -Math.round(baseSizePx * 0.35);
          else if (runVAlign[i] === 'subscript') yShift = Math.round(baseSizePx * 0.10);
          ctx.fillText(runs[i].text, runX, textY + yShift);
          const rSizePx = Math.round(rf.size * PT_TO_PX);
          if (rf.underline) {
            const uyBase = alignV === 'top'
              ? cy + paddingY + rSizePx + 1
              : alignV === 'center'
                ? cy + cellH / 2 + Math.round(rSizePx * 0.55)
                : cy + cellH - paddingY + 1;
            const uy = uyBase + yShift;
            const stroke = runColor ? hexToRgba(runColor) : '#000000';
            drawTextDecoLine(ctx, runX, runX + runWidths[i], uy, stroke, rf.underlineStyle === 'double' || rf.underlineStyle === 'doubleAccounting');
          }
          if (rf.strike) {
            const syBase = alignV === 'top'
              ? cy + paddingY + Math.round(rSizePx * 0.5)
              : alignV === 'center'
                ? cy + cellH / 2
                : cy + cellH - paddingY - Math.round(rSizePx * 0.35);
            const sy = syBase + yShift;
            ctx.save();
            ctx.strokeStyle = runColor ? hexToRgba(runColor) : '#000000';
            ctx.lineWidth = 0.5;
            ctx.beginPath(); ctx.moveTo(runX, sy); ctx.lineTo(runX + runWidths[i], sy); ctx.stroke();
            ctx.restore();
          }
          runX += runWidths[i];
        }
      } else {
        // ECMA-376 §18.4.6 — cell-level super/subscript: render the glyphs at
        // ~65% size, shifted off the baseline so the cell still reads at the
        // right vertical band. Excel uses these defaults (size ratio is
        // implementation-defined; ratios match Office's visual output).
        const cellVertAlign = fontForDraw.vertAlign;
        const baseSizePxOrig = Math.round(font.size * PT_TO_PX);
        let vaYShift = 0;
        if (cellVertAlign === 'superscript') vaYShift = -Math.round(baseSizePxOrig * 0.35);
        else if (cellVertAlign === 'subscript') vaYShift = Math.round(baseSizePxOrig * 0.10);
        const drawFont: Font = cellVertAlign
          ? { ...fontForDraw, size: fontForDraw.size * 0.65 }
          : fontForDraw;
        if (cellVertAlign) {
          ctx.font = buildFont(drawFont, cs);
        }

        // Measure once for both underline and strike
        let overlayMetrics: TextMetrics | null = null;
        const measureOverlay = () => overlayMetrics ??= ctx.measureText(text);
        const overlayX = () => {
          const tW = Math.min(measureOverlay().width, drawW - leftPad - paddingX);
          return {
            x: alignH === 'right' ? cx + cellW - paddingX - tW
              : alignH === 'center' ? cx + cellW / 2 - tW / 2
              : cx + leftPad,
            width: tW,
          };
        };
        const sizePx = Math.round(drawFont.size * PT_TO_PX);

        if (fontForDraw.underline || hyperlinkUrl) {
          const { x: ux, width: tW } = overlayX();
          const uy = (alignV === 'top'
            ? cy + paddingY + sizePx + 1
            : alignV === 'center'
              ? cy + cellH / 2 + Math.round(sizePx * 0.55)
              : cy + cellH - paddingY + 1) + vaYShift;
          const stroke = hyperlinkUrl ? '#0563C1' : (textColor ? hexToRgba(textColor) : '#000000');
          const dbl = fontForDraw.underlineStyle === 'double' || fontForDraw.underlineStyle === 'doubleAccounting';
          drawTextDecoLine(ctx, ux, ux + tW, uy, stroke, dbl);
        }
        if (fontForDraw.strike) {
          const { x: sx, width: tW } = overlayX();
          // Strike line sits roughly at the x-height mid-line (~45% up from baseline)
          const sy = (alignV === 'top'
            ? cy + paddingY + Math.round(sizePx * 0.5)
            : alignV === 'center'
              ? cy + cellH / 2
              : cy + cellH - paddingY - Math.round(sizePx * 0.35)) + vaYShift;
          ctx.save();
          ctx.strokeStyle = textColor ? hexToRgba(textColor) : '#000000';
          ctx.lineWidth = 0.5;
          ctx.beginPath(); ctx.moveTo(sx, sy); ctx.lineTo(sx + tW, sy); ctx.stroke();
          ctx.restore();
        }
        // Hard line breaks (\n from Alt+Enter) render as multiple lines even
        // when wrapText is false — this matches Excel's behavior.
        if (text.includes('\n')) {
          const lines = text.split('\n');
          const lineH = Math.round(font.size * PT_TO_PX * 1.2);
          const totalTextH = lines.length * lineH;
          let startY: number;
          if (alignV === 'top') { startY = cy + paddingY; ctx.textBaseline = 'top'; }
          else if (alignV === 'center') { startY = cy + (cellH - totalTextH) / 2; ctx.textBaseline = 'top'; }
          else { startY = cy + cellH - totalTextH - paddingY; ctx.textBaseline = 'top'; }
          for (let li = 0; li < lines.length; li++) {
            ctx.fillText(lines[li], textX, startY + li * lineH + vaYShift);
          }
        } else {
          let textY: number;
          if (alignV === 'top') { ctx.textBaseline = 'top'; textY = cy + paddingY; }
          else if (alignV === 'center') { ctx.textBaseline = 'middle'; textY = cy + cellH / 2; }
          else { ctx.textBaseline = 'bottom'; textY = cy + cellH - paddingY; }
          ctx.fillText(drawText, textX, textY + vaYShift);
        }
      }

      ctx.restore();

      if (text && rc.onTextRun) {
        rc.onTextRun({ text, x: cx, y: cy, width: cellW, height: cellH, row: rowIndex, col: colIndex });
      }
      });
    }
  }

  for (const task of textTasks) task();

  ctx.restore();
}

// ────────────────────────────────────────────────────────────────
// Main render function
// ────────────────────────────────────────────────────────────────
export function renderViewport(
  ctx: CanvasRenderingContext2D,
  worksheet: Worksheet,
  styles: Styles,
  viewport: ViewportRange,
  opts: RenderViewportOptions = {},
): void {
  const dpr = opts.dpr ?? 1;
  const cs = opts.cellScale ?? 1;
  // Resolve MDW once per render — workbook-wide value derived from the
  // Normal-style font (ECMA-376 §18.3.1.13). Cached internally per
  // (family, sizePt) so repeated renders are O(1).
  const mdw = getMdwForWorksheet(worksheet);
  const canvasW = ctx.canvas.width / dpr;
  const canvasH = ctx.canvas.height / dpr;

  ctx.clearRect(0, 0, canvasW, canvasH);
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvasW, canvasH);

  // Scaled pixel helper: apply cellScale to all cell/header dimensions
  const sp = (px: number) => Math.round(px * cs);
  const hw = sp(HEADER_W);  // scaled header column width
  const hh = sp(HEADER_H);  // scaled header row height

  const { row: startRow, col: startCol, rows: numRows, cols: numCols } = viewport;
  const scrollOffsetX = (opts.scrollOffsetX ?? 0) * cs;
  const scrollOffsetY = (opts.scrollOffsetY ?? 0) * cs;
  const freezeRows = opts.freezeRows ?? 0;
  const freezeCols = opts.freezeCols ?? 0;

  // ── Compute frozen area pixel sizes (scaled) ─────────────────
  const frozenColWidths: number[] = [];
  for (let c = 1; c <= freezeCols; c++) {
    frozenColWidths.push(sp(colWidthToPx(worksheet.colWidths[c] ?? worksheet.defaultColWidth, mdw)));
  }
  const frozenRowHeights: number[] = [];
  for (let r = 1; r <= freezeRows; r++) {
    frozenRowHeights.push(sp(rowHeightToPx(worksheet.rowHeights[r] ?? worksheet.defaultRowHeight)));
  }
  const frozenW = frozenColWidths.reduce((s, w) => s + w, 0);
  const frozenH = frozenRowHeights.reduce((s, h) => s + h, 0);

  // ── Scrollable col/row pixel widths (scaled) ─────────────────
  const scrollColWidths: number[] = [];
  for (let c = startCol; c < startCol + numCols; c++) {
    scrollColWidths.push(sp(colWidthToPx(worksheet.colWidths[c] ?? worksheet.defaultColWidth, mdw)));
  }
  const scrollRowHeights: number[] = [];
  for (let r = startRow; r < startRow + numRows; r++) {
    scrollRowHeights.push(sp(rowHeightToPx(worksheet.rowHeights[r] ?? worksheet.defaultRowHeight)));
  }

  // ── Build cell & merge lookup ────────────────────────────────
  const cellMap = new Map<string, Cell>();
  for (const row of worksheet.rows) {
    for (const cell of row.cells) {
      cellMap.set(`${cell.row}:${cell.col}`, cell);
    }
  }

  const mergeAnchorMap = new Map<string, { totalW: number; totalH: number; right: number; bottom: number }>();
  const mergeSkipSet = new Set<string>();
  for (const mc of worksheet.mergeCells ?? []) {
    let totalW = 0;
    for (let c = mc.left; c <= mc.right; c++) {
      totalW += sp(colWidthToPx(worksheet.colWidths[c] ?? worksheet.defaultColWidth, mdw));
    }
    let totalH = 0;
    for (let r = mc.top; r <= mc.bottom; r++) {
      totalH += sp(rowHeightToPx(worksheet.rowHeights[r] ?? worksheet.defaultRowHeight));
    }
    mergeAnchorMap.set(`${mc.top}:${mc.left}`, { totalW, totalH, right: mc.right, bottom: mc.bottom });
    for (let r = mc.top; r <= mc.bottom; r++) {
      for (let c = mc.left; c <= mc.right; c++) {
        if (r === mc.top && c === mc.left) continue;
        mergeSkipSet.add(`${r}:${c}`);
      }
    }
  }

  const cfContext = compileCf(worksheet);

  // Build autoFilter indicator cell set
  const autoFilterCells = new Set<string>();
  if (worksheet.autoFilter) {
    const af = worksheet.autoFilter;
    for (let c = af.left; c <= af.right; c++) {
      autoFilterCells.add(`${af.top}:${c}`);
    }
  }

  // Build hyperlink lookup map
  const hyperlinkMap = new Map<string, string>();
  for (const hl of worksheet.hyperlinks ?? []) {
    if (hl.url) hyperlinkMap.set(`${hl.row}:${hl.col}`, hl.url);
  }

  // Build commented-cell lookup. worksheet.commentRefs are A1-style refs
  // ("A1", "B12", "AA3") so we convert each to "row:col" and stash in a Set
  // for O(1) membership checks in the cell loop.
  const commentCells = new Set<string>();
  for (const ref of worksheet.commentRefs ?? []) {
    const parsed = parseA1Ref(ref);
    if (parsed) commentCells.add(`${parsed.row}:${parsed.col}`);
  }

  const tableStyleMap = buildTableStyleMap(worksheet);
  const sparklineMap = buildSparklineMap(worksheet);

  const rc: RenderContext = {
    worksheet, styles, cellMap, mergeAnchorMap, mergeSkipSet, cfContext,
    colWidths: scrollColWidths,
    rowHeights: scrollRowHeights,
    frozenColWidths, frozenRowHeights,
    frozenW, frozenH,
    startRow, startCol,
    cs,
    dpr,
    autoFilterCells,
    hyperlinkMap,
    commentCells,
    tableStyleMap,
    sparklineMap,
    mdw,
    onTextRun: opts.onTextRun,
  };

  // Canvas areas for each quadrant
  const cellAreaX = hw;
  const cellAreaY = hh;
  const scrollAreaX = cellAreaX + frozenW;
  const scrollAreaY = cellAreaY + frozenH;
  const scrollAreaW = Math.max(0, canvasW - scrollAreaX);
  const scrollAreaH = Math.max(0, canvasH - scrollAreaY);

  // ── Q1: frozen rows × frozen cols ───────────────────────────
  if (freezeRows > 0 && freezeCols > 0) {
    renderQuadrant(ctx, rc,
      1, 1, frozenColWidths, frozenRowHeights,
      0, 0,
      cellAreaX, cellAreaY,
      cellAreaX, cellAreaY, frozenW, frozenH,
    );
  }

  // ── Q2: frozen rows × scrollable cols ───────────────────────
  if (freezeRows > 0) {
    renderQuadrant(ctx, rc,
      1, startCol, scrollColWidths, frozenRowHeights,
      scrollOffsetX, 0,
      scrollAreaX, cellAreaY,
      scrollAreaX, cellAreaY, scrollAreaW, frozenH,
    );
  }

  // ── Q3: scrollable rows × frozen cols ───────────────────────
  if (freezeCols > 0) {
    renderQuadrant(ctx, rc,
      startRow, 1, frozenColWidths, scrollRowHeights,
      0, scrollOffsetY,
      cellAreaX, scrollAreaY,
      cellAreaX, scrollAreaY, frozenW, scrollAreaH,
    );
  }

  // ── Q4: scrollable rows × scrollable cols (main area) ───────
  renderQuadrant(ctx, rc,
    startRow, startCol, scrollColWidths, scrollRowHeights,
    scrollOffsetX, scrollOffsetY,
    scrollAreaX, scrollAreaY,
    scrollAreaX, scrollAreaY, scrollAreaW, scrollAreaH,
  );

  // ── Anchored images (clipped to scrollable area) ─────────────
  if (worksheet.images && worksheet.images.length > 0 && opts.loadedImages) {
    renderImages(
      ctx, worksheet, opts.loadedImages, cs,
      startRow, startCol,
      scrollOffsetX, scrollOffsetY,
      scrollAreaX, scrollAreaY,
      scrollAreaW, scrollAreaH,
    );
  }

  // ── Anchored shape groups (custom geometry, incl. embedded images) ────
  if (worksheet.shapeGroups && worksheet.shapeGroups.length > 0) {
    renderShapeGroups(
      ctx, worksheet, cs,
      startRow, startCol,
      scrollOffsetX, scrollOffsetY,
      scrollAreaX, scrollAreaY,
      scrollAreaW, scrollAreaH,
      opts.loadedImages,
    );
  }

  // ── Anchored charts (clipped to scrollable area) ──────────────
  if (worksheet.charts && worksheet.charts.length > 0) {
    renderCharts(
      ctx, worksheet, cs,
      startRow, startCol,
      scrollOffsetX, scrollOffsetY,
      scrollAreaX, scrollAreaY,
      scrollAreaW, scrollAreaH,
    );
  }

  // ── Anchored slicers (Office 2010+ pivot/table filter buttons) ──
  if (worksheet.slicers && worksheet.slicers.length > 0) {
    renderSlicers(
      ctx, worksheet, cs,
      startRow, startCol,
      scrollOffsetX, scrollOffsetY,
      scrollAreaX, scrollAreaY,
      scrollAreaW, scrollAreaH,
    );
  }

  // ── Row/col headers (drawn last, always on top) ──────────────
  renderHeaders(ctx, canvasW, canvasH,
    startRow, startCol, numRows, numCols,
    scrollColWidths, scrollRowHeights,
    scrollOffsetX, scrollOffsetY,
    frozenColWidths, frozenRowHeights,
    frozenW, frozenH,
    hw, hh, cs, dpr,
    opts.selectedRowRange ?? null,
    opts.selectedColRange ?? null,
  );

  // ── Freeze pane separator lines ──────────────────────────────
  if (freezeRows > 0) {
    ctx.save();
    ctx.strokeStyle = FREEZE_LINE_COLOR;
    ctx.lineWidth = 0.5;
    ctx.beginPath();
    ctx.moveTo(hw, scrollAreaY + 0.5);
    ctx.lineTo(canvasW, scrollAreaY + 0.5);
    ctx.stroke();
    ctx.restore();
  }
  if (freezeCols > 0) {
    ctx.save();
    ctx.strokeStyle = FREEZE_LINE_COLOR;
    ctx.lineWidth = 0.5;
    ctx.beginPath();
    ctx.moveTo(scrollAreaX + 0.5, hh);
    ctx.lineTo(scrollAreaX + 0.5, canvasH);
    ctx.stroke();
    ctx.restore();
  }
}

// ────────────────────────────────────────────────────────────────
// Headers
// ────────────────────────────────────────────────────────────────
function renderHeaders(
  ctx: CanvasRenderingContext2D,
  canvasW: number, canvasH: number,
  startRow: number, startCol: number,
  numRows: number, numCols: number,
  scrollColWidths: number[], scrollRowHeights: number[],
  scrollOffsetX: number, scrollOffsetY: number,
  frozenColWidths: number[], frozenRowHeights: number[],
  frozenW: number, frozenH: number,
  hw: number, hh: number, cs: number, dpr: number,
  selectedRowRange: { start: number; end: number; strong: boolean } | null,
  selectedColRange: { start: number; end: number; strong: boolean } | null,
): void {
  const HEADER_BG = '#f8f9fa';
  const HEADER_BG_SUBTLE = '#e8eaed';
  const HEADER_BG_STRONG = '#caddf6';
  const HEADER_BORDER = '#c8ccd0';
  const HEADER_BORDER_STRONG = '#5b9bd5';
  const HEADER_TEXT = '#444';

  const colBg = (col: number): string => {
    if (!selectedColRange || col < selectedColRange.start || col > selectedColRange.end) return HEADER_BG;
    return selectedColRange.strong ? HEADER_BG_STRONG : HEADER_BG_SUBTLE;
  };
  const colBorder = (col: number): string => {
    if (!selectedColRange || col < selectedColRange.start || col > selectedColRange.end) return HEADER_BORDER;
    return selectedColRange.strong ? HEADER_BORDER_STRONG : HEADER_BORDER;
  };
  const rowBg = (row: number): string => {
    if (!selectedRowRange || row < selectedRowRange.start || row > selectedRowRange.end) return HEADER_BG;
    return selectedRowRange.strong ? HEADER_BG_STRONG : HEADER_BG_SUBTLE;
  };
  const rowBorder = (row: number): string => {
    if (!selectedRowRange || row < selectedRowRange.start || row > selectedRowRange.end) return HEADER_BORDER;
    return selectedRowRange.strong ? HEADER_BORDER_STRONG : HEADER_BORDER;
  };
  const headerFontSize = Math.max(1, Math.round(11 * cs));
  const HEADER_FONT = `${headerFontSize}px ${DEFAULT_FONT_FAMILY}`;
  const scrollAreaX = hw + frozenW;
  const scrollAreaY = hh + frozenH;
  const hp = 0.5 / dpr;  // half device-pixel offset for 1dp crisp lines

  // Corner – draw all 4 edges (standalone box)
  ctx.fillStyle = HEADER_BG;
  ctx.fillRect(0, 0, hw, hh);
  ctx.strokeStyle = HEADER_BORDER;
  ctx.lineWidth = 0.5;
  ctx.beginPath();
  ctx.moveTo(hp, 0); ctx.lineTo(hp, hh);          // left  (outward — canvas edge)
  ctx.moveTo(0, hp); ctx.lineTo(hw, hp);            // top   (outward — canvas edge)
  ctx.moveTo(hw - hp, 0); ctx.lineTo(hw - hp, hh);  // right  (inset — aligns with row-header right)
  ctx.moveTo(0, hh - hp); ctx.lineTo(hw, hh - hp);  // bottom (inset — aligns with col-header bottom)
  ctx.stroke();

  ctx.font = HEADER_FONT;
  ctx.fillStyle = HEADER_TEXT;

  // Helper: draw one column header cell.
  // Borders are drawn INSET (-hp) so that the next cell's fillRect (which starts at cx+cw)
  // never overwrites the current cell's right/bottom border line.
  const drawColHeader = (col: number, cx: number, cw: number) => {
    ctx.fillStyle = colBg(col);
    ctx.fillRect(cx, 0, cw, hh);
    ctx.strokeStyle = colBorder(col);
    ctx.lineWidth = 0.5;
    ctx.beginPath();
    ctx.moveTo(cx + cw - hp, 0);     ctx.lineTo(cx + cw - hp, hh);  // right (inset)
    ctx.moveTo(cx, hh - hp);          ctx.lineTo(cx + cw, hh - hp);  // bottom (inset)
    ctx.moveTo(cx, hp);               ctx.lineTo(cx + cw, hp);        // top
    ctx.stroke();
    ctx.fillStyle = HEADER_TEXT;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(colToLetter(col), cx + cw / 2, hh / 2);
  };

  // Helper: draw one row header cell.
  // Borders drawn inset so adjacent cell's fill never overwrites them.
  const drawRowHeader = (row: number, cy: number, ch: number) => {
    ctx.fillStyle = rowBg(row);
    ctx.fillRect(0, cy, hw, ch);
    ctx.strokeStyle = rowBorder(row);
    ctx.lineWidth = 0.5;
    ctx.beginPath();
    ctx.moveTo(hw - hp, cy);  ctx.lineTo(hw - hp, cy + ch);   // right (inset)
    ctx.moveTo(0, cy + ch - hp); ctx.lineTo(hw, cy + ch - hp); // bottom (inset)
    ctx.moveTo(hp, cy);       ctx.lineTo(hp, cy + ch);          // left
    ctx.stroke();
    ctx.fillStyle = HEADER_TEXT;
    ctx.textAlign = 'right';
    ctx.textBaseline = 'middle';
    ctx.fillText(String(row), hw - Math.max(2, Math.round(4 * cs)), cy + ch / 2);
  };

  // Frozen col headers (no h-scroll, fixed positions)
  if (frozenColWidths.length > 0) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(hw, 0, frozenW, hh);
    ctx.clip();
    let cx = hw;
    for (let ci = 0; ci < frozenColWidths.length; ci++) {
      drawColHeader(ci + 1, cx, frozenColWidths[ci]);
      cx += frozenColWidths[ci];
    }
    ctx.restore();
  }

  // Scrollable col headers
  ctx.save();
  ctx.beginPath();
  ctx.rect(scrollAreaX, 0, canvasW - scrollAreaX, hh);
  ctx.clip();
  let cx = scrollAreaX - scrollOffsetX;
  for (let ci = 0; ci < scrollColWidths.length; ci++) {
    const cw = scrollColWidths[ci];
    if (cx + cw > scrollAreaX && cx < canvasW) {
      drawColHeader(startCol + ci, cx, cw);
    }
    cx += cw;
  }
  ctx.restore();

  // Frozen row headers (no v-scroll)
  if (frozenRowHeights.length > 0) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, hh, hw, frozenH);
    ctx.clip();
    let cy = hh;
    for (let ri = 0; ri < frozenRowHeights.length; ri++) {
      drawRowHeader(ri + 1, cy, frozenRowHeights[ri]);
      cy += frozenRowHeights[ri];
    }
    ctx.restore();
  }

  // Scrollable row headers
  ctx.save();
  ctx.beginPath();
  ctx.rect(0, scrollAreaY, hw, canvasH - scrollAreaY);
  ctx.clip();
  let cy = scrollAreaY - scrollOffsetY;
  for (let ri = 0; ri < scrollRowHeights.length; ri++) {
    const ch = scrollRowHeights[ri];
    if (cy + ch > scrollAreaY && cy < canvasH) {
      drawRowHeader(startRow + ri, cy, ch);
    }
    cy += ch;
  }
  ctx.restore();

}

// ────────────────────────────────────────────────────────────────
// Image anchors  (ECMA-376 §20.5, <xdr:twoCellAnchor>)
// ────────────────────────────────────────────────────────────────
const EMU_PER_PX = 9525;

/** Sum scaled column widths for cols 1..n-1 (sheet-space X of col n in scaled px). */
function sheetXForCol(
  ws: Worksheet,
  col1: number, // 1-indexed column number
  cs: number,
): number {
  const mdw = getMdwForWorksheet(ws);
  let x = 0;
  for (let c = 1; c < col1; c++) {
    x += Math.round(colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth, mdw) * cs);
  }
  return x;
}

/** Sum scaled row heights for rows 1..n-1 (sheet-space Y of row n in scaled px). */
function sheetYForRow(
  ws: Worksheet,
  row1: number, // 1-indexed row number
  cs: number,
): number {
  let y = 0;
  for (let r = 1; r < row1; r++) {
    y += Math.round(rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight) * cs);
  }
  return y;
}

function renderImages(
  ctx: CanvasRenderingContext2D,
  ws: Worksheet,
  loadedImages: Map<string, HTMLImageElement>,
  cs: number,
  startRow: number,
  startCol: number,
  scrollOffsetX: number,
  scrollOffsetY: number,
  scrollAreaX: number,
  scrollAreaY: number,
  scrollAreaW: number,
  scrollAreaH: number,
): void {
  if (scrollAreaW <= 0 || scrollAreaH <= 0) return;

  // Sheet-space origin of the current scroll viewport's first visible cell
  const scrollOriginSheetX = sheetXForCol(ws, startCol, cs);
  const scrollOriginSheetY = sheetYForRow(ws, startRow, cs);

  ctx.save();
  ctx.beginPath();
  ctx.rect(scrollAreaX, scrollAreaY, scrollAreaW, scrollAreaH);
  ctx.clip();

  for (const anchor of ws.images) {
    const img = loadedImages.get(anchor.dataUrl);
    if (!img) continue;

    // xdr col/row are 0-indexed; our widths map is 1-indexed.
    const fromCol1 = anchor.fromCol + 1;
    const fromRow1 = anchor.fromRow + 1;

    // Image sheet-space top-left (always derived from the `from` anchor)
    const imgSheetX1 = sheetXForCol(ws, fromCol1, cs) + (anchor.fromColOff * cs) / EMU_PER_PX;
    const imgSheetY1 = sheetYForRow(ws, fromRow1, cs) + (anchor.fromRowOff * cs) / EMU_PER_PX;

    // ECMA-376 §20.5.2.33 + "Move but don't size with cells": when the
    // anchor was saved with editAs="oneCell" Excel preserves the picture's
    // saved EMU size (<xdr:spPr><a:xfrm><a:ext>) regardless of cell
    // resizing, and the to anchor is only updated to track that fixed size.
    // Use the native ext directly so the rendered image matches Excel even
    // when our column-width / row-height computation diverges slightly from
    // Excel's (e.g. row ht is stored as px in this viewer but Excel applies
    // pt→px for some files). Falls back to the from/to-derived rect for
    // editAs="twoCell" (default, image resizes with cells) and absolute
    // anchors, or when the parser couldn't capture the native ext.
    let imgW: number, imgH: number;
    if (anchor.editAs === 'oneCell' && anchor.nativeExtCx > 0 && anchor.nativeExtCy > 0) {
      imgW = (anchor.nativeExtCx * cs) / EMU_PER_PX;
      imgH = (anchor.nativeExtCy * cs) / EMU_PER_PX;
    } else {
      const toCol1 = anchor.toCol + 1;
      const toRow1 = anchor.toRow + 1;
      const imgSheetX2 = sheetXForCol(ws, toCol1, cs) + (anchor.toColOff * cs) / EMU_PER_PX;
      const imgSheetY2 = sheetYForRow(ws, toRow1, cs) + (anchor.toRowOff * cs) / EMU_PER_PX;
      imgW = imgSheetX2 - imgSheetX1;
      imgH = imgSheetY2 - imgSheetY1;
    }
    if (imgW <= 0 || imgH <= 0) continue;

    // Translate to canvas coordinates of the scrollable viewport
    const canvasX = scrollAreaX + (imgSheetX1 - scrollOriginSheetX) - scrollOffsetX;
    const canvasY = scrollAreaY + (imgSheetY1 - scrollOriginSheetY) - scrollOffsetY;

    // Early out when entirely off-screen
    if (canvasX + imgW < scrollAreaX || canvasX > scrollAreaX + scrollAreaW) continue;
    if (canvasY + imgH < scrollAreaY || canvasY > scrollAreaY + scrollAreaH) continue;

    ctx.drawImage(img, canvasX, canvasY, imgW, imgH);
  }

  ctx.restore();
}

function renderShapeGroups(
  ctx: CanvasRenderingContext2D,
  ws: Worksheet,
  cs: number,
  startRow: number,
  startCol: number,
  scrollOffsetX: number,
  scrollOffsetY: number,
  scrollAreaX: number,
  scrollAreaY: number,
  scrollAreaW: number,
  scrollAreaH: number,
  loadedImages?: Map<string, HTMLImageElement>,
): void {
  if (scrollAreaW <= 0 || scrollAreaH <= 0) return;
  const anchors = ws.shapeGroups;
  if (!anchors || anchors.length === 0) return;

  const scrollOriginSheetX = sheetXForCol(ws, startCol, cs);
  const scrollOriginSheetY = sheetYForRow(ws, startRow, cs);

  ctx.save();
  ctx.beginPath();
  ctx.rect(scrollAreaX, scrollAreaY, scrollAreaW, scrollAreaH);
  ctx.clip();

  for (const anchor of anchors) {
    const fromCol1 = anchor.fromCol + 1;
    const fromRow1 = anchor.fromRow + 1;

    const x1 = sheetXForCol(ws, fromCol1, cs) + (anchor.fromColOff * cs) / EMU_PER_PX;
    const y1 = sheetYForRow(ws, fromRow1, cs) + (anchor.fromRowOff * cs) / EMU_PER_PX;

    // editAs="oneCell" preserves the group's saved grpSpPr/xfrm/ext EMU
    // size regardless of cell resizing (ECMA-376 §20.5.2.33). See
    // renderImages for the same handling on stand-alone <xdr:pic>.
    let w: number, h: number;
    if (anchor.editAs === 'oneCell' && anchor.nativeExtCx > 0 && anchor.nativeExtCy > 0) {
      w = (anchor.nativeExtCx * cs) / EMU_PER_PX;
      h = (anchor.nativeExtCy * cs) / EMU_PER_PX;
    } else {
      const toCol1 = anchor.toCol + 1;
      const toRow1 = anchor.toRow + 1;
      const x2 = sheetXForCol(ws, toCol1, cs) + (anchor.toColOff * cs) / EMU_PER_PX;
      const y2 = sheetYForRow(ws, toRow1, cs) + (anchor.toRowOff * cs) / EMU_PER_PX;
      w = x2 - x1;
      h = y2 - y1;
    }
    if (w <= 0 || h <= 0) continue;

    const canvasX = scrollAreaX + (x1 - scrollOriginSheetX) - scrollOffsetX;
    const canvasY = scrollAreaY + (y1 - scrollOriginSheetY) - scrollOffsetY;

    if (canvasX + w < scrollAreaX || canvasX > scrollAreaX + scrollAreaW) continue;
    if (canvasY + h < scrollAreaY || canvasY > scrollAreaY + scrollAreaH) continue;

    for (const shape of anchor.shapes) {
      const sx = canvasX + shape.x * w;
      const sy = canvasY + shape.y * h;
      const sw = shape.w * w;
      const sh = shape.h * h;
      if (sw <= 0 || sh <= 0) continue;
      drawShape(ctx, shape, sx, sy, sw, sh, loadedImages);
    }
  }

  ctx.restore();
}

function drawShape(
  ctx: CanvasRenderingContext2D,
  shape: ShapeInfo,
  sx: number, sy: number, sw: number, sh: number,
  loadedImages?: Map<string, HTMLImageElement>,
): void {
  ctx.save();
  if (shape.rot !== 0) {
    ctx.translate(sx + sw / 2, sy + sh / 2);
    ctx.rotate((shape.rot * Math.PI) / 180);
    ctx.translate(-sw / 2, -sh / 2);
  } else {
    ctx.translate(sx, sy);
  }

  if (shape.geom.type === 'custom') {
    for (const path of shape.geom.paths) {
      if (path.w <= 0 || path.h <= 0) continue;
      const kx = sw / path.w;
      const ky = sh / path.h;
      ctx.beginPath();
      // Track pen position for arcTo center computation.
      let penX = 0, penY = 0;
      // Track subpath start for close lineTo.
      let subX = 0, subY = 0;
      for (const cmd of path.commands) {
        switch (cmd.op) {
          case 'moveTo': {
            const px = cmd.x * kx, py = cmd.y * ky;
            ctx.moveTo(px, py);
            penX = subX = px; penY = subY = py;
            break;
          }
          case 'lineTo': {
            const px = cmd.x * kx, py = cmd.y * ky;
            ctx.lineTo(px, py);
            penX = px; penY = py;
            break;
          }
          case 'cubicBezTo': {
            const ex = cmd.x3 * kx, ey = cmd.y3 * ky;
            ctx.bezierCurveTo(
              cmd.x1 * kx, cmd.y1 * ky,
              cmd.x2 * kx, cmd.y2 * ky,
              ex, ey,
            );
            penX = ex; penY = ey;
            break;
          }
          case 'quadBezTo': {
            const ex = cmd.x2 * kx, ey = cmd.y2 * ky;
            ctx.quadraticCurveTo(cmd.x1 * kx, cmd.y1 * ky, ex, ey);
            penX = ex; penY = ey;
            break;
          }
          case 'arcTo': {
            // ECMA-376 §20.1.9.3: pen lies on ellipse at stAng;
            // derive center from pen + stAng, then sweep swAng.
            const rx = cmd.wr * kx, ry = cmd.hr * ky;
            if (rx <= 0 || ry <= 0) break;
            const stRad = (cmd.stAng / 60000) * (Math.PI / 180);
            const swRad = (cmd.swAng / 60000) * (Math.PI / 180);
            const cx = penX - Math.cos(stRad) * rx;
            const cy = penY - Math.sin(stRad) * ry;
            const endRad = stRad + swRad;
            ctx.ellipse(cx, cy, rx, ry, 0, stRad, endRad, swRad < 0);
            penX = cx + Math.cos(endRad) * rx;
            penY = cy + Math.sin(endRad) * ry;
            break;
          }
          case 'close':
            ctx.closePath();
            penX = subX; penY = subY;
            break;
        }
      }
      fillAndStroke(ctx, shape);
    }
  } else if (shape.geom.type === 'preset') {
    ctx.beginPath();
    switch (shape.geom.name) {
      case 'ellipse':
      case 'roundRect': {
        const rx = sw / 2, ry = sh / 2;
        ctx.ellipse(rx, ry, rx, ry, 0, 0, Math.PI * 2);
        break;
      }
      default:
        ctx.rect(0, 0, sw, sh);
    }
    fillAndStroke(ctx, shape);
  } else if (shape.geom.type === 'image') {
    // Image leaf inside a group (e.g. a sun-emoji clip-art nested in the
    // calendar header). The caller pre-decodes every data URL seen in
    // `ws.shapeGroups[*].shapes[*].geom` via XlsxWorkbook.renderViewport,
    // so we should normally have it in `loadedImages`. If not, fall back
    // to a silent skip — drawing an empty rect would look worse.
    const img = loadedImages?.get(shape.geom.dataUrl);
    if (img) {
      ctx.drawImage(img, 0, 0, sw, sh);
    }
  }
  // Shape text body (ECMA-376 §20.5.2.34 `<xdr:txBody>`). Drawn after
  // fill/stroke so it sits on top of the shape's background.
  if (shape.text) {
    drawShapeText(ctx, shape.text, sw, sh);
  }
  ctx.restore();
}

/**
 * Render a shape's `<xdr:txBody>` content into the local (already-translated)
 * coordinate system spanning [0..sw, 0..sh].
 *
 * Handles only the subset that real-world Excel files exercise heavily —
 * single column of paragraphs, per-run bold/italic/size/color/font, paragraph
 * align (`l`/`ctr`/`r`), body anchor (`t`/`ctr`/`b`). Text wrapping uses
 * canvas measurements when `bodyPr@wrap="square"` (the default), and the
 * inset is the OOXML default (`lIns=91440 EMU` ≈ 7.2 pt on each side, plus
 * `tIns=45720 EMU` ≈ 3.6 pt top/bottom). We approximate inset as a fixed
 * 7 px / 4 px since `bodyPr@*Ins` is rarely overridden in practice.
 */
function drawShapeText(
  ctx: CanvasRenderingContext2D,
  txt: import('./types.js').ShapeText,
  sw: number, sh: number,
): void {
  if (sw <= 0 || sh <= 0 || txt.paragraphs.length === 0) return;
  const padX = 7;
  const padY = 4;
  const innerW = Math.max(0, sw - padX * 2);
  const innerH = Math.max(0, sh - padY * 2);
  if (innerW <= 0 || innerH <= 0) return;

  type Line = { runs: { text: string; run: import('./types.js').ShapeTextRun }[]; align: string; height: number; ascent: number };

  const runFont = (run: import('./types.js').ShapeTextRun): { font: string; size: number } => {
    const size = run.size > 0 ? run.size : DEFAULT_FONT_SIZE;
    const px = size * PT_TO_PX;
    const family = run.fontFace ? `"${run.fontFace}", ${DEFAULT_FONT_FAMILY}` : DEFAULT_FONT_FAMILY;
    const weight = run.bold ? 'bold ' : '';
    const italic = run.italic ? 'italic ' : '';
    return { font: `${italic}${weight}${px}px ${family}`, size: px };
  };

  // Wrap each paragraph into lines.
  const wrap = txt.wrap !== 'none';
  const lines: Line[] = [];
  for (const p of txt.paragraphs) {
    const align = p.align || 'l';
    let lineRuns: Line['runs'] = [];
    let lineW = 0;
    let lineHeight = 0;
    let lineAscent = 0;
    const flushLine = () => {
      lines.push({ runs: lineRuns, align, height: lineHeight, ascent: lineAscent });
      lineRuns = [];
      lineW = 0;
      lineHeight = 0;
      lineAscent = 0;
    };
    for (const run of p.runs) {
      const { font, size: pxSize } = runFont(run);
      lineHeight = Math.max(lineHeight, pxSize * 1.2);
      lineAscent = Math.max(lineAscent, pxSize * 0.85);
      ctx.font = font;
      // Split on explicit newlines (from <a:br/>) first.
      const segments = run.text.split('\n');
      for (let s = 0; s < segments.length; s++) {
        if (s > 0) flushLine();
        const seg = segments[s];
        if (!seg) continue;
        if (!wrap) {
          lineRuns.push({ text: seg, run });
          lineW += ctx.measureText(seg).width;
          continue;
        }
        // Greedy word-wrap. Treat each character as a possible break point
        // for CJK; for spaces, prefer breaking on space boundaries. We do a
        // simple character-level greedy wrap which works adequately for
        // both Latin and CJK.
        let buf = '';
        for (const ch of seg) {
          const candidate = buf + ch;
          const cw = ctx.measureText(candidate).width;
          if (lineW + cw > innerW && (buf.length > 0 || lineRuns.length > 0)) {
            if (buf) {
              lineRuns.push({ text: buf, run });
              lineW += ctx.measureText(buf).width;
            }
            flushLine();
            buf = ch;
            ctx.font = font;
            lineHeight = Math.max(lineHeight, pxSize * 1.2);
            lineAscent = Math.max(lineAscent, pxSize * 0.85);
          } else {
            buf = candidate;
          }
        }
        if (buf) {
          lineRuns.push({ text: buf, run });
          lineW += ctx.measureText(buf).width;
        }
      }
    }
    flushLine();
  }

  // Total text block height
  const blockH = lines.reduce((s, l) => s + l.height, 0);

  // Vertical anchor — ECMA-376 §20.1.7.2 <a:bodyPr anchor>.
  // For 'ctr' we intentionally skip Math.max(0,...) so the block stays
  // visually centered even when blockH exceeds innerH (text clips at edge).
  let y0 = padY;
  if (txt.anchor === 'ctr') y0 = padY + (innerH - blockH) / 2;
  else if (txt.anchor === 'b') y0 = padY + Math.max(0, innerH - blockH);

  // Use 'middle' baseline: fillText(x, y) draws with the em-box midpoint at y.
  // This gives clean per-line centering without manual ascent bookkeeping.
  ctx.textBaseline = 'middle';
  let lineTop = y0;
  for (const line of lines) {
    const drawY = lineTop + line.height / 2;
    // Compute total line width to align horizontally
    let totalW = 0;
    for (const r of line.runs) {
      const { font } = runFont(r.run);
      ctx.font = font;
      totalW += ctx.measureText(r.text).width;
    }
    let x = padX;
    if (line.align === 'ctr') x = padX + Math.max(0, (innerW - totalW) / 2);
    else if (line.align === 'r') x = padX + Math.max(0, innerW - totalW);
    for (const r of line.runs) {
      const { font } = runFont(r.run);
      ctx.font = font;
      ctx.fillStyle = r.run.color ?? '#000000';
      ctx.fillText(r.text, x, drawY);
      x += ctx.measureText(r.text).width;
    }
    lineTop += line.height;
  }
}

function fillAndStroke(ctx: CanvasRenderingContext2D, shape: ShapeInfo): void {
  if (shape.fillColor) {
    ctx.fillStyle = shape.fillColor;
    ctx.fill();
  }
  if (shape.strokeColor && shape.strokeWidth > 0) {
    ctx.strokeStyle = shape.strokeColor;
    ctx.lineWidth = Math.max(0.5, shape.strokeWidth / EMU_PER_PX);
    ctx.stroke();
  }
}

// ────────────────────────────────────────────────────────────────
// Border drawing
// ────────────────────────────────────────────────────────────────
/**
 * Overlay any CF-rule border edges on top of the cell's base border. CF
 * borders win per-edge where they set a style (e.g. a red left+right for a
 * "today" column marker replaces the underlying edge only, leaving top/bottom
 * from the base style intact).
 */
/**
 * Resolve the outer border of a merged range. Excel keeps each constituent
 * cell's own style; the merged rectangle's outer edges come from the cells
 * on those edges (e.g. the right edge from the rightmost column's `right`),
 * not from the anchor cell alone. Without this the right border of an
 * `E2:F2` merge — stored on F2 — goes missing.
 */
function resolveMergeBorder(
  anchorBorder: Border,
  anchorRow: number,
  anchorCol: number,
  rightCol: number,
  bottomRow: number,
  cellMap: Map<string, Cell>,
  styles: Styles,
): Border {
  if (rightCol === anchorCol && bottomRow === anchorRow) return anchorBorder;
  const edgeBorder = (r: number, c: number): Border | null => {
    if (r === anchorRow && c === anchorCol) return null;
    const cell = cellMap.get(`${r}:${c}`);
    if (!cell) return null;
    return resolveXf(styles, cell.styleIndex).border;
  };
  const rightB  = edgeBorder(anchorRow, rightCol);
  const bottomB = edgeBorder(bottomRow, anchorCol);
  const cornerB = edgeBorder(bottomRow, rightCol);
  const pick = (primary: BorderEdge | null | undefined, ...rest: Array<BorderEdge | null | undefined>): BorderEdge | null => {
    if (primary?.style) return primary;
    for (const r of rest) if (r?.style) return r;
    return primary ?? null;
  };
  return {
    left:         anchorBorder.left,
    top:          anchorBorder.top,
    right:        pick(rightB?.right,   cornerB?.right,   anchorBorder.right),
    bottom:       pick(bottomB?.bottom, cornerB?.bottom,  anchorBorder.bottom),
    diagonalUp:   anchorBorder.diagonalUp ?? null,
    diagonalDown: anchorBorder.diagonalDown ?? null,
  };
}

function mergeBorders(base: Border, overlay: Border | undefined): Border {
  if (!overlay) return base;
  const pick = (a: BorderEdge | null | undefined, b: BorderEdge | null | undefined): BorderEdge | null =>
    (b && b.style) ? b : (a ?? null);
  return {
    left:         pick(base.left,         overlay.left),
    right:        pick(base.right,        overlay.right),
    top:          pick(base.top,          overlay.top),
    bottom:       pick(base.bottom,       overlay.bottom),
    diagonalUp:   pick(base.diagonalUp,   overlay.diagonalUp),
    diagonalDown: pick(base.diagonalDown, overlay.diagonalDown),
  };
}

function renderBorder(
  ctx: CanvasRenderingContext2D,
  border: Border,
  x: number, y: number, w: number, h: number,
  /** When set, the cell's top edge was inherited from the cell above's bottom
   *  to redraw the part over-painted by this cell's fill. Double-border
   *  rendering uses inverted "outer / inner" extensions in that case so the
   *  line that the upper cell drew as its bottom *outer* (extended past the
   *  corner, at y + 1 from this cell's perspective) is the one we extend
   *  here too — otherwise the inherited redraw shortens that line and
   *  leaves a 1-px gap at every outer corner of the upper cell's double box.
   *  Same idea for `invertedLeft` (inherited from the left cell's right). */
  invertedTop = false,
  invertedLeft = false,
): void {
  type EdgeRef = {
    edge: BorderEdge | null | undefined;
    x1: number; y1: number; x2: number; y2: number;
    /** 'h' = horizontal (top/bottom), 'v' = vertical (left/right),
     *  'd' = diagonal (no doubling support). */
    kind: 'h' | 'v' | 'd';
  };
  const edges: EdgeRef[] = [
    { edge: border.top,         x1: x,     y1: y,     x2: x + w, y2: y,     kind: 'h' },
    { edge: border.bottom,      x1: x,     y1: y + h, x2: x + w, y2: y + h, kind: 'h' },
    { edge: border.left,        x1: x,     y1: y,     x2: x,     y2: y + h, kind: 'v' },
    { edge: border.right,       x1: x + w, y1: y,     x2: x + w, y2: y + h, kind: 'v' },
    { edge: border.diagonalUp,  x1: x,     y1: y + h, x2: x + w, y2: y,     kind: 'd' },
    { edge: border.diagonalDown,x1: x,     y1: y,     x2: x + w, y2: y + h, kind: 'd' },
  ];
  for (const { edge, x1, y1, x2, y2, kind } of edges) {
    if (!edge || !edge.style || edge.style === 'none') continue;
    const color = edge.color ? hexToRgba(edge.color) : '#000000';
    // ECMA-376 §18.18.3 ST_BorderStyle "double": two parallel thin lines
    // with a small gap. Drawn as a 1-px line on either side of the cell
    // edge so the pair reads as a ~3-px-thick double rule, matching Excel.
    // The outer line extends past the cell corners by `off` so adjacent
    // doubled edges close cleanly; the inner line is *shortened* by `off`
    // on each end so the two perpendicular inner lines meet exactly at the
    // inner corner (without the trim they cross past the corner forming a
    // small "+" overhang). Diagonals are handled separately further down.
    if (edge.style === 'double' && kind === 'd') {
      // Two parallel diagonal lines offset perpendicular to the diagonal
      // direction. Excel renders <diagonal style="double"/> the same way it
      // does horizontal/vertical doubles — two thin lines with a small gap.
      ctx.strokeStyle = color;
      ctx.lineWidth = 1;
      ctx.setLineDash([]);
      const dx = x2 - x1;
      const dy = y2 - y1;
      const len = Math.hypot(dx, dy);
      // Perpendicular unit vector. ±off shifts the line ~1px to either side.
      const off = 1;
      const px = (-dy / len) * off;
      const py = (dx / len) * off;
      ctx.beginPath();
      ctx.moveTo(x1 + px, y1 + py); ctx.lineTo(x2 + px, y2 + py);
      ctx.moveTo(x1 - px, y1 - py); ctx.lineTo(x2 - px, y2 - py);
      ctx.stroke();
      continue;
    }
    if (edge.style === 'double' && kind !== 'd') {
      ctx.strokeStyle = color;
      ctx.lineWidth = 1;
      ctx.setLineDash([]);
      const off = 1;
      ctx.beginPath();
      if (kind === 'h') {
        const isTop = y1 === y;
        // For an inherited top, the line at y - off is the upper cell's
        // bottom *inner* (which survives our fill, sitting above the fill
        // band) and the line at y + off is the upper cell's bottom *outer*
        // (which our fill erased and which we are restoring). Swap which
        // side gets the extension so the restored line is the extended one.
        const swap = isTop && invertedTop;
        const outerY = isTop ? (swap ? y + off : y - off) : y + h + off;
        const innerY = isTop ? (swap ? y - off : y + off) : y + h - off;
        ctx.moveTo(x - off, outerY);   ctx.lineTo(x + w + off, outerY);
        ctx.moveTo(x + off, innerY);   ctx.lineTo(x + w - off, innerY);
      } else {
        const isLeft = x1 === x;
        const swap = isLeft && invertedLeft;
        const outerX = isLeft ? (swap ? x + off : x - off) : x + w + off;
        const innerX = isLeft ? (swap ? x - off : x + off) : x + w - off;
        ctx.moveTo(outerX, y - off);   ctx.lineTo(outerX, y + h + off);
        ctx.moveTo(innerX, y + off);   ctx.lineTo(innerX, y + h - off);
      }
      ctx.stroke();
      continue;
    }
    ctx.beginPath();
    ctx.strokeStyle = color;
    ctx.lineWidth = borderStyleWidth(edge.style);
    const dash = borderStyleDash(edge.style);
    ctx.setLineDash(dash);
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();
    ctx.setLineDash([]);
  }
}

function borderStyleWidth(style: string): number {
  // Widths follow Excel's pt convention so medium and thick read as visibly
  // distinct from thin (and from each other) — thin=1pt, medium=2pt, thick=3pt.
  // Earlier values (1.5 / 2) compressed medium and thick to roughly the same
  // 2-row antialiased band, which made e.g. sample-27 row 9 appear to have
  // top (medium) and bottom (thick) of equal weight.
  switch (style) {
    case 'thick': return 3;
    case 'medium': case 'mediumDashed': case 'mediumDashDot': case 'mediumDashDotDot': case 'slantDashDot': return 2;
    case 'hair': return 0.5;
    default: return 1;
  }
}

function borderStyleDash(style: string): number[] {
  switch (style) {
    // ECMA-376 §18.18.3 ST_BorderStyle "hair" — Excel renders this as a very
    // fine dashed line (the finest dashing in the border style picker). A 1-px
    // on / 1-px off pattern at the hair lineWidth (0.5) reproduces that look;
    // without a dash pattern, hair would read as a faint solid line.
    case 'hair': return [1, 1];
    case 'dashed': case 'mediumDashed': return [4, 3];
    case 'dotted': return [2, 2];
    case 'dashDot': case 'mediumDashDot': return [4, 2, 1, 2];
    case 'dashDotDot': case 'mediumDashDotDot': return [4, 2, 1, 2, 1, 2];
    case 'slantDashDot': return [5, 3, 1, 3];
    default: return [];
  }
}

/**
 * Visual precedence per ECMA-376 ST_BorderStyle. Higher = stronger / more
 * visually prominent. Excel's behaviour at a shared edge between two adjacent
 * cells that both define a border is to render the stronger one — this lets
 * the renderer pick a single edge to draw instead of stacking two strokes
 * (which compounds antialiasing and visually thickens the line).
 */
function borderPrecedence(style: string | null | undefined): number {
  switch (style) {
    case 'double': return 13;
    case 'thick': return 12;
    case 'medium': return 11;
    case 'mediumDashed': return 10;
    case 'mediumDashDot': return 9;
    case 'slantDashDot': return 8;
    case 'mediumDashDotDot': return 7;
    case 'thin': return 6;
    case 'dashed': return 5;
    case 'dashDot': return 4;
    case 'dashDotDot': return 3;
    case 'dotted': return 2;
    case 'hair': return 1;
    default: return 0;
  }
}

function pickStrongerEdge(
  a: BorderEdge | null | undefined,
  b: BorderEdge | null | undefined,
): BorderEdge | null {
  const aP = borderPrecedence(a?.style);
  const bP = borderPrecedence(b?.style);
  if (aP === 0 && bP === 0) return null;
  if (aP >= bP) return a ?? null;
  return b ?? null;
}

// ═══════════════════════════════════════════════════════════════════════════
// Chart rendering — delegated to @silurus/ooxml-core's unified renderer.
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Normalize the XLSX parser's raw `ChartData` (chartType="bar" + barDir + grouping)
 * into the canonical `ChartModel.chartType` vocabulary expected by core.
 */
function canonicalChartType(chart: ChartData): string {
  const t = chart.chartType;
  const g = chart.grouping;
  if (t === 'bar') {
    const isH = chart.barDir === 'bar';
    if (g === 'stacked')        return isH ? 'stackedBarH'    : 'stackedBar';
    if (g === 'percentStacked') return isH ? 'stackedBarHPct' : 'stackedBarPct';
    return isH ? 'clusteredBarH' : 'clusteredBar';
  }
  if (t === 'line') {
    if (g === 'stacked')        return 'stackedLine';
    if (g === 'percentStacked') return 'stackedLinePct';
    return 'line';
  }
  if (t === 'area') {
    if (g === 'stacked')        return 'stackedArea';
    if (g === 'percentStacked') return 'stackedAreaPct';
    return 'area';
  }
  return t;
}

function adaptChartData(chart: ChartData): ChartModel {
  return {
    chartType: canonicalChartType(chart),
    title: chart.title,
    categories: chart.categories,
    catAxisFormatCode: chart.catAxisFormatCode ?? null,
    catAxisMin: chart.catAxisMin ?? null,
    catAxisMax: chart.catAxisMax ?? null,
    titleFontBold: chart.titleFontBold ?? null,
    catAxisFontBold: chart.catAxisFontBold ?? null,
    valAxisFontBold: chart.valAxisFontBold ?? null,
    catAxisCrosses: chart.catAxisCrosses ?? null,
    catAxisCrossesAt: chart.catAxisCrossesAt ?? null,
    valAxisCrosses: chart.valAxisCrosses ?? null,
    valAxisCrossesAt: chart.valAxisCrossesAt ?? null,
    catAxisLineColor: chart.catAxisLineColor ?? null,
    catAxisLineWidthEmu: chart.catAxisLineWidthEmu ?? null,
    valAxisLineColor: chart.valAxisLineColor ?? null,
    valAxisLineWidthEmu: chart.valAxisLineWidthEmu ?? null,
    series: chart.series.map(s => ({
      name: s.name,
      color: s.color ?? null,
      values: s.values,
      seriesType: s.seriesType ?? null,
      categories: s.categories.length > 0 ? s.categories : null,
      showMarker: s.showMarker ?? null,
      valFormatCode: s.valFormatCode ?? null,
      markerSymbol: s.markerSymbol ?? null,
      markerSize: s.markerSize ?? null,
      markerFill: s.markerFill ?? null,
      markerLine: s.markerLine ?? null,
      dataPointOverrides: s.dataPointOverrides ?? null,
      dataLabelOverrides: s.dataLabelOverrides ?? null,
      seriesDataLabels: s.seriesDataLabels ?? null,
      errBars: s.errBars ?? null,
    })),
    showDataLabels: chart.showDataLabels ?? false,
    valMin: chart.valAxisMin ?? null,
    valMax: chart.valAxisMax ?? null,
    catAxisTitle: chart.catAxisTitle ?? null,
    valAxisTitle: chart.valAxisTitle ?? null,
    catAxisHidden: chart.catAxisHidden ?? false,
    valAxisHidden: chart.valAxisHidden ?? false,
    catAxisLineHidden: chart.catAxisLineHidden ?? false,
    valAxisLineHidden: chart.valAxisLineHidden ?? false,
    plotAreaBg: null,
    // `<c:chartSpace><c:spPr>` resolution: when the spPr element was present
    // we honor whatever it said (solid hex or `<a:noFill/>` → null =
    // transparent). When spPr was absent the file is relying on the Excel
    // default, which is an opaque white chart area — keep that so legacy
    // charts still get their familiar frame.
    chartBg: chart.hasChartSpPr ? (chart.chartBg ?? null) : 'FFFFFF',
    legendManualLayout: chart.legendManualLayout ?? null,
    // <c:legend> is the authoritative signal: present → show, absent → hide.
    // A single-series bar chart in Excel typically omits <c:legend>, so we
    // must honor that rather than deriving from series count.
    showLegend: chart.showLegend ?? false,
    legendPos: chart.legendPos ?? null,
    catAxisCrossBetween: 'between',
    // Default `out` per ECMA-376 §21.2.2.49 ST_TickMark when the spec
    // didn't say. (We previously hard-coded `cross` which made every
    // chart pretend it had crossing ticks even when the file said
    // none / out.)
    valAxisMajorTickMark: chart.valAxisMajorTickMark ?? 'out',
    catAxisMajorTickMark: chart.catAxisMajorTickMark ?? 'out',
    valAxisMinorTickMark: chart.valAxisMinorTickMark ?? null,
    catAxisMinorTickMark: chart.catAxisMinorTickMark ?? null,
    titleFontSizeHpt: chart.titleFontSizeHpt ?? null,
    titleFontColor: chart.titleFontColor ?? null,
    titleFontFace: chart.titleFontFace ?? null,
    catAxisFontSizeHpt: chart.catAxisFontSizeHpt ?? null,
    valAxisFontSizeHpt: chart.valAxisFontSizeHpt ?? null,
    dataLabelFontSizeHpt: null,
    subtotalIndices: [],
    valAxisFormatCode: chart.valAxisFormatCode ?? null,
    barGapWidth: chart.barGapWidth ?? null,
    barOverlap: chart.barOverlap ?? null,
    dataLabelPosition: chart.dataLabelPosition ?? null,
    dataLabelFontColor: chart.dataLabelFontColor ?? null,
    dataLabelFormatCode: chart.dataLabelFormatCode ?? null,
    titleManualLayout: chart.titleManualLayout ?? null,
    plotAreaManualLayout: chart.plotAreaManualLayout ?? null,
    radarStyle: chart.radarStyle ?? null,
  };
}

// ── renderCharts ────────────────────────────────────────────────────────────

function renderCharts(
  ctx: CanvasRenderingContext2D,
  ws: Worksheet,
  cs: number,
  startRow: number,
  startCol: number,
  scrollOffsetX: number,
  scrollOffsetY: number,
  scrollAreaX: number,
  scrollAreaY: number,
  scrollAreaW: number,
  scrollAreaH: number,
): void {
  if (scrollAreaW <= 0 || scrollAreaH <= 0) return;

  const scrollOriginSheetX = sheetXForCol(ws, startCol, cs);
  const scrollOriginSheetY = sheetYForRow(ws, startRow, cs);

  for (const anchor of ws.charts) {
    const fromCol1 = anchor.fromCol + 1;
    const fromRow1 = anchor.fromRow + 1;
    const toCol1   = anchor.toCol   + 1;
    const toRow1   = anchor.toRow   + 1;

    const shX1 = sheetXForCol(ws, fromCol1, cs) + (anchor.fromColOff * cs) / EMU_PER_PX;
    const shY1 = sheetYForRow(ws, fromRow1, cs) + (anchor.fromRowOff * cs) / EMU_PER_PX;
    const shX2 = sheetXForCol(ws, toCol1,   cs) + (anchor.toColOff   * cs) / EMU_PER_PX;
    const shY2 = sheetYForRow(ws, toRow1,   cs) + (anchor.toRowOff   * cs) / EMU_PER_PX;

    const cw = shX2 - shX1;
    const ch = shY2 - shY1;
    if (cw <= 0 || ch <= 0) continue;

    const cx = scrollAreaX + (shX1 - scrollOriginSheetX) - scrollOffsetX;
    const cy = scrollAreaY + (shY1 - scrollOriginSheetY) - scrollOffsetY;

    if (cx + cw < scrollAreaX || cx > scrollAreaX + scrollAreaW) continue;
    if (cy + ch < scrollAreaY || cy > scrollAreaY + scrollAreaH) continue;

    ctx.save();
    ctx.beginPath();
    ctx.rect(scrollAreaX, scrollAreaY, scrollAreaW, scrollAreaH);
    ctx.clip();

    // XLSX natural rendering is device-px at 96 DPI where 1pt = 4/3 px. Scale
    // that by `cs` so OOXML-specified font sizes (title/axes) scale with zoom.
    const ptToPx = (4 / 3) * cs;
    renderChart(ctx, adaptChartData(anchor.chart), { x: cx, y: cy, w: cw, h: ch }, ptToPx);
    ctx.restore();
  }
}

// ── renderSlicers ───────────────────────────────────────────────────────────
//
// Office 2010+ pivot / table slicer. We don't own a slicer engine, so this
// renders a static button bank: header with the slicer caption, then one
// button per saved item using the selection flags from the slicerCache. The
// visual language (pale blue outline, white "selected" buttons on a darker
// background, gray "deselected" buttons) intentionally mirrors Excel's
// default slicer style — the workbook may ship a custom `slicerStyle` but
// rendering that is deferred (the built-in look is already recognisable).

const SLICER_HEADER_FONT = '600 12px "Meiryo UI", "Segoe UI", sans-serif';
const SLICER_ITEM_FONT   = '11px "Meiryo UI", "Segoe UI", sans-serif';
const SLICER_BG           = '#FFFFFF';
const SLICER_BORDER       = '#BFBFBF';
const SLICER_HEADER_BG    = '#F2F2F2';
const SLICER_HEADER_FG    = '#404040';
const SLICER_ITEM_SEL_BG  = '#FFFFFF';
const SLICER_ITEM_SEL_FG  = '#000000';
const SLICER_ITEM_SEL_BD  = '#A5A5A5';
const SLICER_ITEM_OFF_BG  = '#E7E6E6';
const SLICER_ITEM_OFF_FG  = '#A6A6A6';
const SLICER_ITEM_OFF_BD  = '#C6C6C6';

function renderSlicers(
  ctx: CanvasRenderingContext2D,
  ws: Worksheet,
  cs: number,
  startRow: number,
  startCol: number,
  scrollOffsetX: number,
  scrollOffsetY: number,
  scrollAreaX: number,
  scrollAreaY: number,
  scrollAreaW: number,
  scrollAreaH: number,
): void {
  if (scrollAreaW <= 0 || scrollAreaH <= 0) return;
  const slicers = ws.slicers;
  if (!slicers) return;

  const scrollOriginSheetX = sheetXForCol(ws, startCol, cs);
  const scrollOriginSheetY = sheetYForRow(ws, startRow, cs);

  for (const anchor of slicers) {
    const fromCol1 = anchor.fromCol + 1;
    const fromRow1 = anchor.fromRow + 1;
    const toCol1   = anchor.toCol   + 1;
    const toRow1   = anchor.toRow   + 1;

    const shX1 = sheetXForCol(ws, fromCol1, cs) + (anchor.fromColOff * cs) / EMU_PER_PX;
    const shY1 = sheetYForRow(ws, fromRow1, cs) + (anchor.fromRowOff * cs) / EMU_PER_PX;
    const shX2 = sheetXForCol(ws, toCol1,   cs) + (anchor.toColOff   * cs) / EMU_PER_PX;
    const shY2 = sheetYForRow(ws, toRow1,   cs) + (anchor.toRowOff   * cs) / EMU_PER_PX;

    const w = shX2 - shX1;
    const h = shY2 - shY1;
    if (w <= 0 || h <= 0) continue;

    const x = scrollAreaX + (shX1 - scrollOriginSheetX) - scrollOffsetX;
    const y = scrollAreaY + (shY1 - scrollOriginSheetY) - scrollOffsetY;

    if (x + w < scrollAreaX || x > scrollAreaX + scrollAreaW) continue;
    if (y + h < scrollAreaY || y > scrollAreaY + scrollAreaH) continue;

    ctx.save();
    ctx.beginPath();
    ctx.rect(scrollAreaX, scrollAreaY, scrollAreaW, scrollAreaH);
    ctx.clip();

    drawSlicerFrame(ctx, anchor.caption, anchor.items, x, y, w, h, cs);

    ctx.restore();
  }
}

function drawSlicerFrame(
  ctx: CanvasRenderingContext2D,
  caption: string,
  items: SlicerItem[],
  x: number,
  y: number,
  w: number,
  h: number,
  cs: number,
): void {
  // Outer frame (white with a soft gray hairline).
  ctx.fillStyle = SLICER_BG;
  ctx.fillRect(x, y, w, h);
  ctx.strokeStyle = SLICER_BORDER;
  ctx.lineWidth = 1;
  ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);

  // Header band with caption.
  const headerH = Math.max(20 * cs, 14);
  ctx.fillStyle = SLICER_HEADER_BG;
  ctx.fillRect(x + 1, y + 1, w - 2, headerH);
  ctx.fillStyle = SLICER_HEADER_FG;
  ctx.font = scaleFont(SLICER_HEADER_FONT, cs);
  ctx.textBaseline = 'middle';
  ctx.textAlign = 'left';
  const headerPad = 6 * cs;
  drawClippedText(ctx, caption, x + headerPad, y + headerH / 2 + 1, w - 2 * headerPad);

  // Item buttons. Items expand to fill the available height (up to a
  // minimum button size) and are clipped by the slicer rect. We don't
  // implement scroll arrows because this renderer is non-interactive.
  if (items.length === 0) return;
  const gap = Math.max(1, Math.round(2 * cs));
  const innerPad = 4 * cs;
  const listX = x + innerPad;
  const listY = y + headerH + innerPad;
  const listW = w - 2 * innerPad;
  const listH = h - headerH - 2 * innerPad;
  if (listW <= 0 || listH <= 0) return;

  // Prefer Excel's rough row height (~20 sheet-px) but compress if the
  // slicer is shallow so at least the first items fit.
  const preferredItemH = Math.max(18 * cs, 16);
  const maxVisibleByH = Math.max(1, Math.floor((listH + gap) / (preferredItemH + gap)));
  const visible = Math.min(items.length, maxVisibleByH);
  const itemH = Math.min(preferredItemH, (listH - gap * (visible - 1)) / visible);
  if (itemH <= 0) return;

  ctx.font = scaleFont(SLICER_ITEM_FONT, cs);
  const itemPad = 8 * cs;
  for (let i = 0; i < visible; i++) {
    const item = items[i];
    const iy = listY + i * (itemH + gap);
    const selected = item.selected;
    ctx.fillStyle = selected ? SLICER_ITEM_SEL_BG : SLICER_ITEM_OFF_BG;
    ctx.fillRect(listX, iy, listW, itemH);
    ctx.strokeStyle = selected ? SLICER_ITEM_SEL_BD : SLICER_ITEM_OFF_BD;
    ctx.lineWidth = 1;
    ctx.strokeRect(listX + 0.5, iy + 0.5, listW - 1, itemH - 1);
    ctx.fillStyle = selected ? SLICER_ITEM_SEL_FG : SLICER_ITEM_OFF_FG;
    drawClippedText(ctx, item.name, listX + itemPad, iy + itemH / 2 + 1, listW - 2 * itemPad);
  }
}

function scaleFont(css: string, cs: number): string {
  // Re-scale the leading `<size>px` token by `cs`. Safe fallback: leave the
  // string as-is so the slicer remains readable when parsing fails.
  return css.replace(/(\d+(?:\.\d+)?)px/, (_, n) => `${Math.round(Number(n) * cs)}px`);
}

function drawClippedText(
  ctx: CanvasRenderingContext2D,
  text: string,
  x: number,
  y: number,
  maxWidth: number,
): void {
  if (maxWidth <= 0) return;
  let s = text;
  if (ctx.measureText(s).width > maxWidth) {
    const ellipsis = '…';
    while (s.length > 0 && ctx.measureText(s + ellipsis).width > maxWidth) {
      s = s.slice(0, -1);
    }
    s = s.length > 0 ? s + ellipsis : '';
  }
  ctx.fillText(s, x, y);
}
