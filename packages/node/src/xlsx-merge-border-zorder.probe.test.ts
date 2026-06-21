/**
 * Device-pixel probe for the merged-cell LEFT-border z-order fix.
 *
 * Bug (sample-30 B14:B23): the vertical merge's LEFT border rendered thinner /
 * greyer below the anchor row, while the anchor row's slice and the merge's
 * RIGHT border stayed crisp black. Root cause: the renderer draws cell
 * gridlines (#d0d0d0) interleaved per-cell at each cell's right+bottom edge,
 * then draws cell borders on top — an ordering that only works because, within
 * a row, the left neighbour is visited before the cell. A merged anchor draws
 * its FULL-height border once (during the anchor's row), but the cells in the
 * column to the LEFT of the covered rows are visited in LATER rows and stroke
 * their right gridline over the already-drawn black left border, eating one
 * device column of it. The merge's RIGHT edge is never overpainted (no cell
 * draws a gridline on that boundary in the covered rows), so it stays a clean
 * control.
 *
 * The fix defers merged-anchor borders to a pass that runs after the whole
 * grid loop (after every gridline), so the merge perimeter always sits above
 * the gridlines — matching Excel, where an explicit cell border replaces the
 * gridline on that edge.
 *
 * This probe builds a SYNTHETIC sheet (no private data → CI-safe, though gated
 * on skia like the sibling border-crisp probe): a single 1-col × 6-row merge
 * with a thin pure-black box border, an empty column to its left, and gridlines
 * on. It renders at dpr=2 and asserts the LEFT edge of a COVERED row carries the
 * same amount of black ink as the merge's never-overpainted RIGHT edge. Before
 * the fix the covered-row left edge is thinner → the assertion fails, so this is
 * a genuine regression guard, not a tautology.
 */
import { describe, it, expect } from 'vitest';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';
import { installImageBitmapShim, installOffscreenCanvasShim } from './render.ts';
import type { NodeCanvasFactory } from './render.ts';
import type {
  Worksheet,
  Styles,
  Border,
  BorderEdge,
  CellFont,
  CellFill,
  CellXf,
  Row,
  ViewportRange,
} from '@silurus/ooxml-xlsx';

const skia = await import('skia-canvas').catch(() => null);
type Skia = typeof import('skia-canvas');
const { Canvas } = (skia ?? {}) as Skia;

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) =>
    new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (() => {
    throw new Error('loadImage not needed');
  }) as unknown as NodeCanvasFactory['loadImage'],
};

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '../../..');
const ORCH_PATH = resolve(ROOT, 'packages/xlsx/src/render-orchestrator.ts');

const HEADER_W = 50;
const HEADER_H = 22;

const MERGE_TOP = 3;
const MERGE_BOTTOM = 8; // 6-row vertical merge
const MERGE_COL = 3; // column C; column B (2) to its left stays empty

function lum(r: number, g: number, b: number): number {
  return 0.299 * r + 0.587 * g + 0.114 * b;
}

/** A thin pure-black box border on all four edges. */
function buildSynthetic(): { ws: Worksheet; styles: Styles } {
  const thin: BorderEdge = { style: 'thin', color: '#000000' };
  const font: CellFont = {
    bold: false, italic: false, underline: false, strike: false,
    size: 11, color: null, name: 'Calibri',
  };
  const noFill: CellFill = { patternType: 'none', fgColor: null, bgColor: null };
  const emptyBorder: Border = { left: null, right: null, top: null, bottom: null };
  const boxBorder: Border = { left: thin, right: thin, top: thin, bottom: thin };
  const xf0: CellXf = {
    fontId: 0, fillId: 0, borderId: 0, numFmtId: 0,
    alignH: null, alignV: null, wrapText: false,
  };
  const xfBox: CellXf = { ...xf0, borderId: 1 };
  const styles: Styles = {
    fonts: [font],
    fills: [noFill],
    borders: [emptyBorder, boxBorder],
    cellXfs: [xf0, xfBox],
    numFmts: [],
    dxfs: [],
  };

  // The covered cells carry the same border style as the anchor (mirrors real
  // files like sample-30, where every cell of the merge column references the
  // bordered xf).
  const rows: Row[] = [];
  for (let r = MERGE_TOP; r <= MERGE_BOTTOM; r++) {
    rows.push({
      index: r,
      height: null,
      cells: [
        { col: MERGE_COL, row: r, colRef: `${MERGE_COL}`, value: { type: 'empty' }, styleIndex: 1 },
      ],
    });
  }

  const ws: Worksheet = {
    name: 'S',
    rows,
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 10,
    defaultRowHeight: 18,
    mergeCells: [{ top: MERGE_TOP, left: MERGE_COL, bottom: MERGE_BOTTOM, right: MERGE_COL }],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts: [],
    showGridlines: true,
    defaultFontFamily: 'Calibri',
    defaultFontSize: 11,
  };
  return { ws, styles };
}

const W = 360;
const H = 280;

async function render(dpr: number): Promise<{ data: Uint8ClampedArray; w: number; h: number }> {
  const { ws, styles } = buildSynthetic();
  const { renderWorksheetViewport } = (await import(ORCH_PATH)) as {
    renderWorksheetViewport: (
      deps: { ws: Worksheet; styles: Styles; imageCache: Map<string, unknown> },
      target: unknown,
      viewport: ViewportRange,
      opts: { dpr: number; width: number; height: number },
    ) => Promise<void>;
  };
  const canvas = new Canvas(W * dpr, H * dpr);
  const restoreImg = installImageBitmapShim(factory);
  const restoreOff = installOffscreenCanvasShim(factory);
  try {
    await renderWorksheetViewport(
      { ws, styles, imageCache: new Map() },
      canvas,
      { row: 1, col: 1, rows: 14, cols: 8 } as ViewportRange,
      { dpr, width: W, height: H },
    );
  } finally {
    restoreOff();
    restoreImg();
  }
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
  const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
  return { data: img.data, w: canvas.width, h: canvas.height };
}

/** Pure-black, near-neutral pixel (excludes the grey #d0d0d0 gridline at L≈208
 *  and any grey header chrome). */
function isBlack(data: Uint8ClampedArray, i: number): boolean {
  const r = data[i], g = data[i + 1], b = data[i + 2];
  const spread = Math.max(r, g, b) - Math.min(r, g, b);
  return lum(r, g, b) < 60 && spread < 24;
}

/** Bounding box of all near-black pixels inside the cell area (excludes the
 *  header strips so header label glyphs can't pollute it). */
function blackBBox(data: Uint8ClampedArray, w: number, h: number, dpr: number) {
  const x0 = Math.round((HEADER_W + 2) * dpr);
  const y0 = Math.round((HEADER_H + 2) * dpr);
  let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
  for (let y = y0; y < h; y++) {
    for (let x = x0; x < w; x++) {
      if (isBlack(data, (y * w + x) * 4)) {
        if (x < minX) minX = x;
        if (x > maxX) maxX = x;
        if (y < minY) minY = y;
        if (y > maxY) maxY = y;
      }
    }
  }
  return { minX, maxX, minY, maxY };
}

/** Count near-black device columns within a horizontal window around `cx`,
 *  taking the max over a small vertical neighbourhood of `y` (robust to AA). */
function blackColsAt(data: Uint8ClampedArray, w: number, cx: number, y: number, half = 3): number {
  let best = 0;
  for (let dy = -1; dy <= 1; dy++) {
    let n = 0;
    for (let dx = -half; dx <= half; dx++) {
      if (isBlack(data, ((y + dy) * w + (cx + dx)) * 4)) n++;
    }
    if (n > best) best = n;
  }
  return best;
}

describe.skipIf(!skia)('xlsx merged-cell left border (z-order probe)', () => {
  it('covered-row left edge carries the same black ink as the merge right edge', async () => {
    const dpr = 2;
    const { data, w, h } = await render(dpr);
    const bb = blackBBox(data, w, h, dpr);
    expect(Number.isFinite(bb.minX)).toBe(true);

    const leftX = bb.minX;
    const rightX = bb.maxX;
    // Anchor row sits just below the merge's top border; a covered row sits well
    // into the lower half of the span.
    const anchorY = bb.minY + Math.round((bb.maxY - bb.minY) * 0.12);
    const coveredY = bb.minY + Math.round((bb.maxY - bb.minY) * 0.62);

    const leftAnchor = blackColsAt(data, w, leftX, anchorY);
    const rightAnchor = blackColsAt(data, w, rightX, anchorY);
    const leftCovered = blackColsAt(data, w, leftX, coveredY);
    const rightCovered = blackColsAt(data, w, rightX, coveredY);

    // eslint-disable-next-line no-console
    console.log(
      `[MERGE-ZORDER dpr=${dpr}] bbox x[${bb.minX}..${bb.maxX}] y[${bb.minY}..${bb.maxY}]\n` +
        `  anchorY=${anchorY}: leftBlackCols=${leftAnchor} rightBlackCols=${rightAnchor}\n` +
        `  coveredY=${coveredY}: leftBlackCols=${leftCovered} rightBlackCols=${rightCovered}`,
    );

    // The right edge is the never-overpainted control. The covered-row LEFT edge
    // must match it (the merge's black left border owns that boundary, the
    // gridline must not eat into it).
    expect(rightCovered).toBeGreaterThan(0);
    expect(leftCovered).toBe(rightCovered);
  });
});
