/**
 * Vertical (tbRl) prolonged sound mark `ー` reflection probe — ECMA-376 §17.6.20 +
 * UAX#50 §5 Tr fallback (core `verticalTrMirrorFallback`).
 *
 * `ー` (U+30FC) is vo=Tr with NO Unicode vertical presentation form, so it takes the
 * ROTATE fallback. But its font-DESIGNED vertical glyph is the HORIZONTAL REFLECTION
 * of the +90° page rotation, NOT the rotation itself — a documented Japanese
 * typographic convention (the 起筆/uroko and the stroke curvature flip left↔right
 * between the horizontal and vertical orientations). A plain rotation therefore paints
 * a LEFT-RIGHT MIRROR of Word.
 *
 * Word ground truth (sample-47 PDF, Yu Mincho via macOS Quartz) and the font's own
 * `vert` glyph (Hiragino Mincho ProN cid07891) agree: the vertical `ー` has its thick
 * HEAD at the TOP bulging to the RIGHT, tapering to a blunt bottom-left. The
 * un-reflected +90° rotation bulges LEFT. `drawVerticalRun` reflects the mark via
 * `scale(1, -1)` (the on-screen horizontal mirror in the +90° page frame) to match.
 *
 * This probe renders `話ーー話` through the REAL `drawVerticalRun` and asserts the
 * mark's HEAD (top band) ink leans to the physical RIGHT of the glyph's own vertical
 * axis — the Word topology — which a sign regression (dropping the reflection, or
 * reflecting the wrong axis) would flip. Measured relative to the glyph's OWN ink
 * bbox so it is font-metric independent; the separation from the mirror is ~0.15 of
 * the glyph width (robust, not a sub-pixel margin).
 *
 * CI-safe: gated on skia (devDependency) + Hiragino (macOS dev host).
 */
import { describe, it, expect } from 'vitest';
import { existsSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { loadSkiaForTests, importForTests } from './test-imports';

const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas, FontLibrary } = (skia ?? {}) as Skia;
const MINCHO = '/System/Library/Fonts/ヒラギノ明朝 ProN.ttc';
const haveFont = existsSync(MINCHO);
const vtMod = await importForTests(
  () => import('../../docx/src/vertical-text.ts'),
  'packages/docx/src/vertical-text.ts',
);
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

function ink(d: Uint8ClampedArray, i: number): number {
  return 255 - (0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2]);
}

describe.skipIf(!skia || !vtMod || !haveFont)('docx vertical prolonged sound mark ー reflection (§17.6.20 / UAX#50)', () => {
  const fontPx = 120;
  const W = 400, H = 700, baseline = 220, logX = 40, cellW = fontPx;
  const centerline = W - baseline; // column centreline in physical x

  /** Render "話<mid>話" via drawVerticalRun and return the pixel buffer. */
  function render(mid: string): Uint8ClampedArray {
    const { drawVerticalRun } = vtMod as { drawVerticalRun: (...a: Any[]) => void };
    for (const fam of ['Yu Mincho', 'MS Mincho', 'Hiragino Mincho ProN']) FontLibrary.use(fam, [MINCHO]);
    const canvas = new Canvas(W, H);
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, H);
    ctx.fillStyle = '#000';
    ctx.font = `${fontPx}px "MS Mincho", serif`;
    ctx.save();
    ctx.translate(W, 0); ctx.rotate(Math.PI / 2);
    drawVerticalRun(ctx, `話${mid}話`, logX, baseline, fontPx, 0);
    ctx.restore();
    return ctx.getImageData(0, 0, W, H).data;
  }

  /**
   * The middle cell (2nd glyph) occupies the along-column band physical y in
   * [logX+cellW, logX+2cellW]. Return the ink bbox there plus the cross-axis (physical
   * x) centroid of the TOP third of that ink — where the mark's head sits — normalised
   * to the glyph's own ink-bbox width (0 = left edge, 1 = right edge). Also the whole-
   * glyph cross centroid for the column-centring sanity check.
   */
  function head(data: Uint8ClampedArray): { topNorm: number; wholeCx: number; x0: number; x1: number } {
    const py0 = logX + cellW, py1 = logX + 2 * cellW;
    let x0 = W, x1 = 0, sxAll = 0, swAll = 0;
    for (let py = py0; py < py1; py++) for (let px = 0; px < W; px++) {
      const w = ink(data, (py * W + px) * 4);
      if (w > 40) { if (px < x0) x0 = px; if (px > x1) x1 = px; sxAll += px * w; swAll += w; }
    }
    const topEnd = py0 + (py1 - py0) / 3;
    let sx = 0, sw = 0;
    for (let py = py0; py < topEnd; py++) for (let px = 0; px < W; px++) {
      const w = ink(data, (py * W + px) * 4);
      if (w > 40) { sx += px * w; sw += w; }
    }
    const topCx = sw > 0 ? sx / sw : NaN;
    return { topNorm: (topCx - x0) / (x1 - x0), wholeCx: swAll > 0 ? sxAll / swAll : NaN, x0, x1 };
  }

  /** Robust physical cross-axis drift over every ink-bearing along-column row. */
  function strokeAngleDeg(data: Uint8ClampedArray): number {
    const py0 = logX + cellW, py1 = logX + 2 * cellW;
    const centroids: Array<{ along: number; cross: number }> = [];
    for (let py = py0; py < py1; py++) {
      let sx = 0, sw = 0;
      for (let px = 0; px < W; px++) {
        const w = ink(data, (py * W + px) * 4);
        if (w > 0) { sx += px * w; sw += w; }
      }
      if (sw > 0) centroids.push({ along: py, cross: sx / sw });
    }
    const slopes: number[] = [];
    for (let i = 0; i < centroids.length; i++) {
      for (let j = i + 1; j < centroids.length; j++) {
        slopes.push(
          (centroids[j].cross - centroids[i].cross) /
          (centroids[j].along - centroids[i].along),
        );
      }
    }
    slopes.sort((a, b) => a - b);
    const middle = Math.floor(slopes.length / 2);
    const slope = slopes.length % 2 === 0
      ? (slopes[middle - 1] + slopes[middle]) / 2
      : slopes[middle];
    return Math.atan(slope) * 180 / Math.PI;
  }

  it('straightens the non-vert fallback ー stroke to within 0.8 degrees', () => {
    const angle = strokeAngleDeg(render('ー'));
    expect(Math.abs(angle)).toBeLessThanOrEqual(0.8);
  });

  it.each(['〜', '～'])('keeps the fallback %s waveform byte-identical', (mark) => {
    const digest = createHash('sha256').update(render(mark)).digest('hex');
    expect(digest).toBe('c3c484be6ea34f4283a70bf4aef9b07e62efdd0c2141772074f581e16842c60b');
  });

  it('ー head leans to the physical RIGHT of its own axis (Word topology), the reflected form', () => {
    const m = head(render('ー'));
    // Word / the font `vert` glyph put the head's bulge on the RIGHT: the top-band ink
    // centroid sits in the RIGHT half of the mark's own width. A dropped/flipped
    // reflection would land it in the LEFT half (~0.4). Threshold 0.5 with the measured
    // ~0.57 (vs ~0.43 mirror) leaves a robust, non-sub-pixel margin.
    expect(m.topNorm).toBeGreaterThan(0.5);
    // Sanity: the mark still centres on the column (the reflection is about the cell
    // centre, so the advance/centring is unchanged) — within 0.1em of the centreline.
    expect(Math.abs(m.wholeCx - centerline)).toBeLessThanOrEqual(0.1 * fontPx);
    // The shear grows only the cross-axis bbox and the result remains inside its
    // one-em column band, so an authored textbox clip does not cut the stroke.
    expect(m.x0).toBeGreaterThan(centerline - fontPx / 2);
    expect(m.x1).toBeLessThan(centerline + fontPx / 2);
  });
});
