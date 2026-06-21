/**
 * Device-pixel crispness probe for the pptx axis-aligned thin-line fix.
 *
 * Background: on a DPR=1 monitor pptx table cell borders (and text underline /
 * strike-through) rendered blurry (~2 device rows at ~50% ink each) instead of
 * PowerPoint's crisp 1px. Root cause: `renderSlide` applies `ctx.scale(dpr,dpr)`,
 * so drawing is in logical px; a `lineWidth=1` stroke at an INTEGER y has its
 * span `[y-0.5, y+0.5]` in device space → it straddles two device rows (each
 * ~50% ink → antialiased blur). The fix (renderer.ts `renderTable` /
 * `drawUnderline` / strike-through) adds a half-device-pixel offset
 * `crispOffset(lw, dpr) = 0.5/dpr` ONLY when the device-pixel width is odd,
 * centering odd-width strokes on a single device row (crisp).
 *
 * This test MEASURES device pixels (not eyeballs). It builds a synthetic
 * Presentation in memory (no .pptx file, no WASM parse) containing a single
 * table cell carrying a thin PURE-BLACK bottom border (1 logical px) on the
 * default white background, renders it through `renderSlideNode` at dpr=1 and
 * dpr=2, and reads the vertical luminance profile across that border. Pure black
 * on white is unambiguous, and the cell's geometry is chosen so the border's
 * bottom edge lands on an INTEGER device-y at dpr=1 (the worst case: an integer
 * coordinate straddles two rows without the fix).
 *
 * This committed test asserts the FIXED state (crisp single near-black device
 * row at dpr=1; clean 2-device-row band at dpr=2). It fails against the pre-fix
 * renderer, so it is a genuine regression guard — not a tautology.
 *
 * CI-safe: skia-canvas ships a native binding CI omits, so the suite is gated
 * with `describe.skipIf(!skia)`. The render helper (`./render.ts`) and the pptx
 * renderer it dynamically imports do NOT statically import WASM, but to stay
 * future-proof the render module is itself loaded via a caught dynamic import and
 * gated too.
 */
import { describe, it, expect } from 'vitest';
import type { Presentation, Slide, TableElement, Stroke } from '@silurus/ooxml-pptx';

const skia = await import('skia-canvas').catch(() => null);
type Skia = typeof import('skia-canvas');
const { Canvas } = (skia ?? {}) as Skia;

// `./render.ts` does not statically import WASM today, but load it via a caught
// dynamic import (and gate on it) so the suite never fails at collection if that
// ever changes.
const renderMod = await import('./render.ts').catch(() => null);

const EMU_PER_PX = 9525; // 96 dpi → 1 logical px

// Slide / table geometry, all in EMU so logical px == EMU / EMU_PER_PX when the
// render width matches the slide's px width.
const SLIDE_W_PX = 960;
const SLIDE_H_PX = 540;
const TABLE_X_PX = 100;
const TABLE_Y_PX = 100;
const COL_W_PX = 200;
const ROW_H_PX = 80;
// Bottom edge of the single cell, in logical px: integer 180. At dpr=1 the
// device row is 180 (worst case — straddles two rows without the fix); at dpr=2
// it is 360 (even device width → clean 2-row band, no over-correction).
const BORDER_Y_PX = TABLE_Y_PX + ROW_H_PX; // 180

function lum(r: number, g: number, b: number): number {
  return 0.299 * r + 0.587 * g + 0.114 * b;
}

/** A single-cell table whose only inked edge is a thin pure-black bottom
 *  border (1 logical px), on the default white slide background. */
function buildTableSlide(): Presentation {
  const blackThin: Stroke = { color: '#000000', width: EMU_PER_PX }; // 1 logical px
  const cell = {
    textBody: null,
    fill: null,
    borderL: null,
    borderR: null,
    borderT: null,
    borderB: blackThin,
    gridSpan: 1,
    rowSpan: 1,
    hMerge: false,
    vMerge: false,
  };
  const table: TableElement = {
    type: 'table',
    x: TABLE_X_PX * EMU_PER_PX,
    y: TABLE_Y_PX * EMU_PER_PX,
    width: COL_W_PX * EMU_PER_PX,
    height: ROW_H_PX * EMU_PER_PX,
    cols: [COL_W_PX * EMU_PER_PX],
    rows: [{ height: ROW_H_PX * EMU_PER_PX, cells: [cell] }],
  };
  const slide: Slide = {
    index: 0,
    slideNumber: 1,
    background: null, // → white
    elements: [table],
  };
  return {
    slideWidth: SLIDE_W_PX * EMU_PER_PX,
    slideHeight: SLIDE_H_PX * EMU_PER_PX,
    slides: [slide],
    defaultTextColor: null,
    majorFont: null,
    minorFont: null,
  };
}

async function renderTableSlide(
  dpr: number,
): Promise<{ data: Uint8ClampedArray; w: number; h: number; canvas: InstanceType<typeof Canvas> }> {
  if (!renderMod) throw new Error('render module unavailable');
  const { renderSlideNode } = renderMod;
  const canvas = new Canvas(Math.round(SLIDE_W_PX * dpr), Math.round(SLIDE_H_PX * dpr));
  await renderSlideNode(
    canvas as unknown as Parameters<typeof renderSlideNode>[0],
    buildTableSlide(),
    0,
    { width: SLIDE_W_PX, dpr },
  );
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
  const w = canvas.width;
  const h = canvas.height;
  const img = ctx.getImageData(0, 0, w, h);
  return { data: img.data, w, h, canvas };
}

/** Locate the injected thin black bottom border: the only dark, near-neutral
 *  horizontal run isolated in white (white a few px above AND below). Works for
 *  BOTH the crisp AFTER case (one pure-black row, L≈0) and the blurry BEFORE
 *  case (two adjacent mid-grey rows, L≈128 each, ink split 50/50). */
function findBlackHLine(
  data: Uint8ClampedArray,
  w: number,
  h: number,
  dpr: number,
): { x: number; y: number; runLen: number } {
  const isDarkNeutral = (i: number): boolean => {
    const r = data[i];
    const g = data[i + 1];
    const b = data[i + 2];
    const L = lum(r, g, b);
    const spread = Math.max(r, g, b) - Math.min(r, g, b);
    return L < 170 && spread < 24;
  };
  const isLight = (i: number): boolean =>
    lum(data[i], data[i + 1], data[i + 2]) > 200;
  const minRun = Math.round(120 * dpr); // 200-px-wide cell → ~200*dpr px run
  const gap = Math.round(4 * dpr);
  let best = { x: -1, y: -1, runLen: 0 };
  for (let y = gap; y < h - gap; y++) {
    let cur = 0;
    let curStart = -1;
    for (let x = 0; x <= w; x++) {
      const i = (y * w + x) * 4;
      const dark = x < w && isDarkNeutral(i);
      if (dark) {
        if (curStart < 0) curStart = x;
        cur++;
      } else {
        if (cur >= minRun) {
          const mx = curStart + Math.floor(cur / 2);
          const aboveWhite = isLight(((y - gap) * w + mx) * 4);
          const belowWhite = isLight(((y + gap) * w + mx) * 4);
          if (aboveWhite && belowWhite && cur > best.runLen) {
            best = { x: mx, y, runLen: cur };
          }
        }
        cur = 0;
        curStart = -1;
      }
    }
  }
  return best;
}

function vProfile(
  data: Uint8ClampedArray,
  w: number,
  h: number,
  x: number,
  y: number,
  span = 3,
): { ys: number[]; L: number[] } {
  const ys: number[] = [];
  const L: number[] = [];
  for (let dy = -span; dy <= span; dy++) {
    const yy = y + dy;
    ys.push(yy);
    if (yy < 0 || yy >= h) {
      L.push(NaN);
      continue;
    }
    const i = (yy * w + x) * 4;
    L.push(lum(data[i], data[i + 1], data[i + 2]));
  }
  return { ys, L };
}

/** Recenter on the darkest row within ±r so a 2-row blurry band reports
 *  symmetrically around its center of mass. */
function darkestNear(
  data: Uint8ClampedArray,
  w: number,
  x: number,
  y: number,
  r = 2,
): number {
  let cy = y;
  let best = Infinity;
  for (let dy = -r; dy <= r; dy++) {
    const i = ((y + dy) * w + x) * 4;
    const L = lum(data[i], data[i + 1], data[i + 2]);
    if (L < best) {
      best = L;
      cy = y + dy;
    }
  }
  return cy;
}

describe.skipIf(!skia || !renderMod)('pptx line crispness (device-pixel probe)', () => {
  it('thin black table border at dpr=1 collapses to one near-black device row', async () => {
    const { data, w, h } = await renderTableSlide(1);

    const hit = findBlackHLine(data, w, h, 1);
    expect(hit.x).toBeGreaterThanOrEqual(0);

    const cy = darkestNear(data, w, hit.x, hit.y, 2);
    const { ys, L } = vProfile(data, w, h, hit.x, cy, 3);
    const finite = L.filter((v) => !Number.isNaN(v));
    const minLum = Math.min(...finite);
    const darkRowCount = finite.filter((v) => v < 160).length;

    // eslint-disable-next-line no-console
    console.log(
      `\n[PROBE dpr=1] border @ (x=${hit.x}, y=${cy}) runLen=${hit.runLen}\n` +
        ys.map((yy, k) => `  y=${yy} L=${L[k].toFixed(1)}`).join('\n') +
        `\n  minLum=${minLum.toFixed(1)} darkRowCount(<160)=${darkRowCount}`,
    );

    // A thin (1 device px) black border must collapse to one near-black row.
    expect(minLum).toBeLessThan(80);
    expect(darkRowCount).toBe(1);
  });

  it('dpr=2 sanity: thin border = even device width → clean 2-row band (no over-correction)', async () => {
    const { data, w, h } = await renderTableSlide(2);

    const hit = findBlackHLine(data, w, h, 2);
    expect(hit.x).toBeGreaterThanOrEqual(0);

    const cy = darkestNear(data, w, hit.x, hit.y, 3);
    const { ys, L } = vProfile(data, w, h, hit.x, cy, 4);
    const finite = L.filter((v) => !Number.isNaN(v));
    const darkRowCount = finite.filter((v) => v < 160).length;

    // eslint-disable-next-line no-console
    console.log(
      `\n[PROBE dpr=2] border @ (x=${hit.x}, y=${cy}) runLen=${hit.runLen}\n` +
        ys.map((yy, k) => `  y=${yy} L=${L[k].toFixed(1)}`).join('\n') +
        `\n  darkRowCount(<160)=${darkRowCount}`,
    );

    // deviceW=2 (even) → no offset → a clean 2-device-row band, NOT 3.
    expect(darkRowCount).toBeLessThanOrEqual(2);
  });
});
