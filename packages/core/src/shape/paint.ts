import type { Fill, PatternFill, Stroke } from '../types/common';
import { buildPatternBitmap } from './pattern-bitmaps';

/**
 * Convert a 6- or 8-char hex colour to a CSS `rgba()` string.
 * 8-char hex encodes alpha in the last two chars (RRGGBBAA).
 * `alpha` applies to 6-char hex; ignored for 8-char.
 * A leading `#` is tolerated (`#RRGGBB` and `RRGGBB` both work).
 */
export function hexToRgba(hex: string, alpha = 1): string {
  const h = hex.charCodeAt(0) === 35 /* '#' */ ? hex.slice(1) : hex;
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  const a = h.length >= 8 ? parseInt(h.slice(6, 8), 16) / 255 : alpha;
  return `rgba(${r},${g},${b},${a})`;
}

/**
 * Rec.601 perceptual luma (`0.299·R + 0.587·G + 0.114·B`) of a colour, on the
 * 0–255 scale. Accepts a 6- or 8-char hex; a leading `#` is tolerated and the
 * alpha byte (if present) is ignored, matching {@link hexToRgba}'s hex
 * normalisation.
 */
export function relativeLuma(hex: string): number {
  const h = hex.charCodeAt(0) === 35 /* '#' */ ? hex.slice(1) : hex;
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return 0.299 * r + 0.587 * g + 0.114 * b;
}

/**
 * Pick black or white text for legibility against a background colour. The
 * mid-gray threshold (128) on the Rec.601 luma splits light vs dark: a dark
 * background ⇒ white text, otherwise black. `bgHex=null` (no background ⇒ page
 * white) ⇒ black text. The black/white pick is implementation-defined — no
 * normative algorithm exists (ECMA-376 §17.3.2.6 `w:color="auto"` only says the
 * consumer chooses "an appropriate color based on the background").
 */
export function autoContrastColor(bgHex: string | null): '#000000' | '#FFFFFF' {
  if (!bgHex) return '#000000';
  return relativeLuma(bgHex) < 128 ? '#FFFFFF' : '#000000';
}

/**
 * Resolve a Fill to a CanvasRenderingContext2D-compatible paint.
 * Gradients require pixel bounds (x, y, w, h) to construct the CanvasGradient.
 * Returns null for noFill.
 */
export function resolveFill(
  fill: Fill | null,
  ctx: CanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
): string | CanvasGradient | CanvasPattern | null {
  if (!fill || fill.fillType === 'none') return null;
  if (fill.fillType === 'solid') return hexToRgba(fill.color);
  if (fill.fillType === 'pattern') {
    return resolvePatternFill(fill, ctx);
  }
  if (fill.fillType === 'gradient') {
    const stops = fill.stops;
    if (stops.length === 0) return null;
    if (stops.length === 1) return hexToRgba(stops[0].color);

    let gradient: CanvasGradient;
    if (fill.gradType === 'radial') {
      const cx = x + w / 2;
      const cy = y + h / 2;
      const r = Math.sqrt(w * w + h * h) / 2;
      gradient = ctx.createRadialGradient(cx, cy, 0, cx, cy, r);
    } else {
      const rad = (fill.angle * Math.PI) / 180;
      const cx = x + w / 2;
      const cy = y + h / 2;
      const gradLen = (Math.abs(Math.cos(rad)) * w + Math.abs(Math.sin(rad)) * h) / 2;
      gradient = ctx.createLinearGradient(
        cx - Math.cos(rad) * gradLen, cy - Math.sin(rad) * gradLen,
        cx + Math.cos(rad) * gradLen, cy + Math.sin(rad) * gradLen,
      );
    }
    for (const stop of stops) {
      gradient.addColorStop(Math.min(1, Math.max(0, stop.position)), hexToRgba(stop.color));
    }
    return gradient;
  }
  return null;
}

/**
 * Build a tiling CanvasPattern for an OOXML preset pattern fill.
 * Falls back to the foreground colour string when the preset name is unknown
 * or the OffscreenCanvas / Canvas environment cannot create a pattern.
 *
 * Cached per (preset, fg, bg) tuple — patterns are immutable bitmaps so the
 * same backing canvas can be reused across many shapes.
 */
const patternCache = new WeakMap<CanvasRenderingContext2D, Map<string, CanvasPattern>>();

function resolvePatternFill(
  fill: PatternFill,
  ctx: CanvasRenderingContext2D,
): CanvasPattern | string {
  const key = `${fill.preset}|${fill.fg}|${fill.bg}`;
  let perCtx = patternCache.get(ctx);
  if (!perCtx) {
    perCtx = new Map();
    patternCache.set(ctx, perCtx);
  }
  const cached = perCtx.get(key);
  if (cached) return cached;

  const bitmap = buildPatternBitmap(fill.preset, fill.fg, fill.bg);
  if (!bitmap) return hexToRgba(fill.fg);
  const pat = ctx.createPattern(bitmap, 'repeat');
  if (!pat) return hexToRgba(fill.fg);
  perCtx.set(key, pat);
  return pat;
}

const DASH_PATTERNS: Record<string, number[]> = {
  dash:         [6, 3],
  dot:          [1.5, 3],
  dashDot:      [6, 3, 1.5, 3],
  lgDash:       [10, 4],
  lgDashDot:    [10, 4, 1.5, 4],
  lgDashDotDot: [10, 4, 1.5, 4, 1.5, 4],
  sysDash:      [4, 2],
  sysDot:       [1, 2],
  sysDashDot:   [4, 2, 1, 2],
};

/**
 * Apply a Stroke to ctx. `emuPerPx` converts stroke width from EMU to px
 * (e.g. scale factor from pptx's emuToPx).
 */
export function applyStroke(
  ctx: CanvasRenderingContext2D,
  stroke: Stroke | null,
  emuPerPx: number,
): void {
  if (!stroke) {
    ctx.strokeStyle = 'transparent';
    ctx.lineWidth = 0;
    ctx.setLineDash([]);
    return;
  }
  ctx.strokeStyle = hexToRgba(stroke.color);
  const lw = Math.max(0.5, stroke.width * emuPerPx);
  ctx.lineWidth = lw;
  const pat = stroke.dashStyle ? DASH_PATTERNS[stroke.dashStyle] : null;
  ctx.setLineDash(pat ? pat.map((v) => v * lw) : []);
}
