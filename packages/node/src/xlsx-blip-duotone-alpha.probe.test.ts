/**
 * End-to-end device-pixel probe for xlsx blip effects (issue #880):
 *  - `<a:duotone>` recolour (ECMA-376 §20.1.8.23), and
 *  - `<a:alphaModFix>` opacity (§20.1.8.6).
 *
 * Real file (sample-9.xlsx, "Gift budget and tracker"): a near-white opaque
 * photo is shown by Excel as PINK and SEMI-TRANSPARENT — the pink comes from a
 * duotone (`black` ↔ light-pink `FFF3F4`, resolved from `DAB6BA` + lumMod/lumOff/
 * tint/satMod), and the translucency from `alphaModFix amt="70000"` (0.70), which
 * blends the picture with the coloured cells beneath. We previously drew the raw
 * opaque white PNG.
 *
 * This probe is FIXTURE-FREE (CI-safe — sample-9 is private/redistribution-
 * restricted): it builds a synthetic worksheet whose single picture is a solid
 * near-white 4×4 PNG carrying `alpha: 0.7` + `duotone {clr1:000000,clr2:FFF3F4}`,
 * over a fully-coloured (pure blue) sheet region. It renders twice through the
 * SAME orchestrator path the browser/worker use, once WITH the effects and once
 * with a raw (effect-stripped) copy, and asserts, at real device pixels:
 *   1. duotone recolours the neutral white photo to a pink cast (R > G ≈ B); and
 *   2. alphaModFix < 1 makes the picture composite over the blue cell beneath
 *      (the drawn pixel is pulled toward blue vs. the opaque control).
 * Gated on skia like the sibling render probes.
 */
import { describe, it, expect } from 'vitest';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';
import { installImageBitmapShim, installOffscreenCanvasShim } from './render.ts';
import type { NodeCanvasFactory } from './render.ts';
import type { Worksheet, Styles, ViewportRange, ImageAnchor } from '@silurus/ooxml-xlsx';
import { loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas, loadImage } = (skia ?? {}) as Skia;

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) =>
    new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: ((buf: ArrayBuffer | Uint8Array | Buffer) =>
    loadImage(Buffer.from(buf as Uint8Array))) as unknown as NodeCanvasFactory['loadImage'],
};

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '../../..');
const ORCH_PATH = resolve(ROOT, 'packages/xlsx/src/render-orchestrator.ts');

/** A minimal 4×4 solid opaque PNG of colour (r,g,b), base-64 → bytes. Built with
 *  skia so we don't hand-encode PNG. Returned as a Blob for `fetchImage`. */
function solidPngBytes(r: number, g: number, b: number): Uint8Array {
  const c = new Canvas(4, 4);
  const ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D;
  ctx.fillStyle = `rgb(${r},${g},${b})`;
  ctx.fillRect(0, 0, 4, 4);
  // skia-canvas: `toBufferSync('png')` (sync) or the async `png` getter.
  const buf = (c as unknown as { toBufferSync(fmt: string): Buffer }).toBufferSync('png');
  return new Uint8Array(buf);
}

/** A sheet with one twoCellAnchor picture over a fully blue-filled region. */
function buildSynthetic(): { ws: Worksheet; styles: Styles } {
  const blueFill = { patternType: 'solid', fgColor: '#0000FF', bgColor: null } as const;
  const styles = {
    fonts: [{ bold: false, italic: false, underline: false, strike: false, size: 11, color: null, name: 'Calibri' }],
    fills: [{ patternType: 'none', fgColor: null, bgColor: null }, blueFill],
    borders: [{ left: null, right: null, top: null, bottom: null }],
    cellXfs: [
      { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, alignH: null, alignV: null, wrapText: false },
      { fontId: 0, fillId: 1, borderId: 0, numFmtId: 0, alignH: null, alignV: null, wrapText: false },
    ],
    numFmts: [],
    dxfs: [],
  } as unknown as Styles;

  // Fill a block of cells (cols 1..6, rows 1..12) with the blue fill so the
  // picture sits over solid colour (so an alpha blend is visible).
  const rows = [];
  for (let r = 1; r <= 12; r++) {
    const cells = [];
    for (let col = 1; col <= 6; col++) {
      cells.push({ col, row: r, value: { type: 'empty' as const }, styleIndex: 1 });
    }
    rows.push({ index: r, height: null, cells });
  }

  const image: ImageAnchor = {
    fromCol: 1, fromColOff: 0, fromRow: 2, fromRowOff: 0,
    toCol: 5, toColOff: 0, toRow: 10, toRowOff: 0,
    nativeExtCx: 0, nativeExtCy: 0,
    imagePath: 'xl/media/image1.png',
    mimeType: 'image/png',
    alpha: 0.7,
    duotone: { clr1: '000000', clr2: 'FFF3F4' },
  };

  const ws = {
    name: 'S',
    rows,
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 12,
    defaultRowHeight: 20,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [image],
    charts: [],
    showGridlines: false,
    defaultFontFamily: 'Calibri',
    defaultFontSize: 11,
  } as unknown as Worksheet;

  return { ws, styles };
}

const W = 500;
const H = 400;

async function render(effects: 'full' | 'duotone-opaque' | 'raw'): Promise<{ data: Uint8ClampedArray; w: number; h: number }> {
  const { ws, styles } = buildSynthetic();
  // A mid-grey opaque source PNG. sample-9's photo corners are near-white, but a
  // mid grey lands the duotone at a mid-pink (leaving clear channel headroom so
  // the alpha blend over the blue cell is measurable — with a near-white source
  // the pink endpoint FFF3F4 is itself so close to white that the blend delta is
  // tiny). The recolour/blend LOGIC is identical for any luminance.
  const pngBytes = solidPngBytes(150, 150, 150);
  const fetchImage = async (): Promise<Blob> =>
    new Blob([pngBytes as unknown as BlobPart], { type: 'image/png' });

  // Clone + optionally strip effects for the controls.
  const ws2 = JSON.parse(JSON.stringify(ws)) as Worksheet;
  const im = (ws2.images as ImageAnchor[])[0];
  if (effects === 'duotone-opaque') delete im.alpha;
  if (effects === 'raw') { delete im.alpha; delete im.duotone; }

  const { renderWorksheetViewport } = (await import(ORCH_PATH)) as {
    renderWorksheetViewport: (
      deps: { ws: Worksheet; styles: Styles; imageCache: Map<string, unknown> },
      target: unknown,
      viewport: ViewportRange,
      opts: { dpr: number; width: number; height: number; fetchImage: () => Promise<Blob> },
    ) => Promise<void>;
  };
  const canvas = new Canvas(W, H);
  const restoreImg = installImageBitmapShim(factory);
  const restoreOff = installOffscreenCanvasShim(factory);
  try {
    await renderWorksheetViewport(
      { ws: ws2, styles, imageCache: new Map() },
      canvas,
      { row: 1, col: 1, rows: 14, cols: 8 } as ViewportRange,
      { dpr: 1, width: W, height: H, fetchImage },
    );
  } finally {
    restoreOff();
    restoreImg();
  }
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
  const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
  return { data: img.data, w: canvas.width, h: canvas.height };
}

/** Mean RGB over a rectangle. */
function mean(data: Uint8ClampedArray, w: number, x0: number, y0: number, x1: number, y1: number) {
  let r = 0, g = 0, b = 0, n = 0;
  for (let y = y0; y < y1; y++) {
    for (let x = x0; x < x1; x++) {
      const i = (y * w + x) * 4;
      r += data[i]; g += data[i + 1]; b += data[i + 2]; n++;
    }
  }
  return { r: r / n, g: g / n, b: b / n };
}

describe.skipIf(!skia)('xlsx blip duotone + alphaModFix (issue #880)', () => {
  it('recolours a white photo to pink (duotone) and composites it over the cell (alphaModFix)', async () => {
    const raw = await render('raw');
    const duoOpaque = await render('duotone-opaque');
    const full = await render('full');

    // Sample a patch well inside the picture rect (cols 1..5 × rows 2..10 in a
    // 12-col-wide 14-row viewport at ~ (colStart + 1.5 cols, rowStart + 3 rows)).
    // A 40×40 patch near the picture centre is safely interior.
    const cx = Math.round(W * 0.35);
    const cy = Math.round(H * 0.4);
    const box = [cx - 20, cy - 20, cx + 20, cy + 20] as const;
    const pRaw = mean(raw.data, raw.w, ...box);
    const pDuo = mean(duoOpaque.data, duoOpaque.w, ...box);
    const pFull = mean(full.data, full.w, ...box);

    // 1. Raw white photo is neutral: R ≈ G ≈ B (no colour cast).
    expect(Math.abs(pRaw.r - pRaw.g)).toBeLessThan(6);
    expect(Math.abs(pRaw.r - pRaw.b)).toBeLessThan(6);
    // And it is NOT blue-dominant (the opaque photo hides the blue cell).
    expect(pRaw.b).toBeLessThan(pRaw.r + 40);

    // 2. Duotone recolours toward light pink FFF3F4: R > G and R > B by a clear
    //    margin (the pink signature). The mid-grey source lands mid-ramp.
    expect(pDuo.r).toBeGreaterThan(pDuo.g + 3);
    expect(pDuo.r).toBeGreaterThan(pDuo.b + 3);
    // And the recolour actually happened (not left as neutral grey): the
    // green/blue channels dropped below the raw grey.
    expect(pDuo.g).toBeLessThan(pRaw.g - 2);

    // 3. alphaModFix (0.7) composites over the pure-blue cell: the full render's
    //    blue channel is meaningfully higher than the opaque duotone control
    //    (blue bleeds through), proving the alpha blend is applied.
    expect(pFull.b).toBeGreaterThan(pDuo.b + 15);
    // The pink cast survives the blend: red still exceeds green (the blend only
    // adds blue, so R>G from the duotone is preserved).
    expect(pFull.r).toBeGreaterThan(pFull.g);
  });
});
