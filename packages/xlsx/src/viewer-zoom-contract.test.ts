import { describe, it, expect, afterEach, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import { installDom, makeContainer } from './viewer-destroy-test-dom.js';
import { HEADER_W, HEADER_H, colWidthToPx, rowHeightToPx, getMdwForWorksheet } from './renderer.js';
import type { Worksheet } from './types.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * IX9 — the XlsxViewer's slice of the shared {@link import('@silurus/ooxml-core').ZoomableViewer}
 * contract. XlsxViewer already had `setScale` + a slider + Ctrl-wheel zoom; IX9
 * adds `getScale` / `zoomIn` / `zoomOut` / `fitWidth` / `fitPage` and the
 * `onScaleChange` notification, WITHOUT changing the existing slider / setScale
 * clamp-and-snap behaviour (non-regression is pinned below).
 *
 * The viewer is driven through its real API on the hand-rolled fake DOM (the repo
 * has no jsdom). A worksheet is injected into the private `currentWorksheet` field
 * — the same technique the hit-test suite uses — so scale + fit run against real
 * geometry. `renderCurrentSheet` early-returns on the fake DOM's 0-sized canvas,
 * so these tests observe scale STATE and the callback, not pixels.
 */

/** A small used range: 3 custom columns + defaults, 2 custom rows. */
function makeSheet(): Worksheet {
  return {
    name: 'Sheet1',
    rows: [{ index: 5, cells: [{ row: 5, col: 8, value: 'x' }] }],
    colWidths: { 1: 12, 2: 20, 3: 30 },
    rowHeights: { 1: 25, 2: 40 },
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts: [],
  } as unknown as Worksheet;
}

interface Priv {
  currentWorksheet: Worksheet | null;
  canvasArea: { clientWidth: number; clientHeight: number };
}

/** Construct a viewer, inject `ws`, and size the canvas area so fit math has a
 *  laid-out container (the fake DOM defaults geometry to 0). */
function mount(
  ws: Worksheet | null,
  opts: ConstructorParameters<typeof XlsxViewer>[1] = {},
  container = { cw: 400, ch: 300 },
): { v: XlsxViewer; priv: Priv } {
  const v = new XlsxViewer(makeContainer() as unknown as HTMLElement, opts);
  const priv = v as unknown as Priv;
  priv.currentWorksheet = ws;
  priv.canvasArea.clientWidth = container.cw;
  priv.canvasArea.clientHeight = container.ch;
  return { v, priv };
}

/** Natural (cs=1) used-range extent oracle, mirroring the viewer's private
 *  `_naturalContentExtent` (header + used cols/rows, no scroll headroom). */
function naturalExtent(ws: Worksheet): { width: number; height: number } {
  const mdw = getMdwForWorksheet(ws);
  let maxRow = Math.max(50, ws.freezeRows ?? 0);
  let maxCol = Math.max(26, ws.freezeCols ?? 0);
  for (const row of ws.rows) {
    if (row.index > maxRow) maxRow = row.index;
    for (const cell of row.cells) if (cell.col > maxCol) maxCol = cell.col;
  }
  let width = HEADER_W;
  for (let c = 1; c <= maxCol; c++) width += colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth, mdw);
  let height = HEADER_H;
  for (let r = 1; r <= maxRow; r++) height += rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight);
  return { width, height };
}

describe('XlsxViewer IX9 zoom contract', () => {
  it('getScale() is 1 (100%) by default and reflects the cellScale option', () => {
    installDom();
    expect(mount(makeSheet()).v.getScale()).toBe(1);
    expect(mount(makeSheet(), { cellScale: 1.5 }).v.getScale()).toBe(1.5);
  });

  it('setScale fires onScaleChange with the snapped factor exactly once', () => {
    installDom();
    const onScaleChange = vi.fn();
    const { v } = mount(makeSheet(), { onScaleChange });
    v.setScale(1.5);
    expect(v.getScale()).toBe(1.5);
    expect(onScaleChange).toHaveBeenCalledTimes(1);
    expect(onScaleChange).toHaveBeenCalledWith(1.5);
  });

  it('setScale does NOT fire onScaleChange when the scale is unchanged', () => {
    installDom();
    const onScaleChange = vi.fn();
    const { v } = mount(makeSheet(), { onScaleChange, cellScale: 2 });
    v.setScale(2); // already 200%
    expect(onScaleChange).not.toHaveBeenCalled();
  });

  it('setScale clamps to [zoomMin, zoomMax]', () => {
    installDom();
    const { v } = mount(makeSheet(), { zoomMin: 0.5, zoomMax: 2 });
    v.setScale(10);
    expect(v.getScale()).toBe(2);
    v.setScale(0.01);
    expect(v.getScale()).toBe(0.5);
  });

  it('zoomIn / zoomOut walk the shared ladder', () => {
    installDom();
    const { v } = mount(makeSheet());
    expect(v.getScale()).toBe(1);
    v.zoomIn();
    expect(v.getScale()).toBe(1.1); // 100% → next ladder rung
    v.zoomIn();
    expect(v.getScale()).toBe(1.25);
    v.zoomOut();
    expect(v.getScale()).toBe(1.1);
    v.zoomOut();
    expect(v.getScale()).toBe(1);
  });

  it('zoomIn from an off-ladder (wheel-zoomed) scale snaps onto the ladder', () => {
    installDom();
    const { v } = mount(makeSheet(), { cellScale: 1.03 });
    v.zoomIn();
    expect(v.getScale()).toBe(1.1);
  });

  it('fitWidth sets the scale that spans the used-range width in the container', () => {
    installDom();
    const ws = makeSheet();
    const { width } = naturalExtent(ws);
    const cw = 400;
    const { v } = mount(ws, {}, { cw, ch: 300 });
    v.fitWidth();
    // fitScale = cw / width, then snapped to whole percent by setScale.
    const expected = Math.round((cw / width) * 100) / 100;
    expect(v.getScale()).toBe(expected);
  });

  it('fitPage takes the tighter of the width/height fit', () => {
    installDom();
    const ws = makeSheet();
    const { width, height } = naturalExtent(ws);
    const cw = 400;
    const ch = 300;
    const { v } = mount(ws, {}, { cw, ch });
    v.fitPage();
    const raw = Math.min(cw / width, ch / height);
    const expected = Math.round(raw * 100) / 100;
    expect(v.getScale()).toBe(expected);
    // fitPage must never exceed fitWidth (height can only tighten).
    const wRaw = cw / width;
    expect(raw).toBeLessThanOrEqual(wRaw);
  });

  it('fitWidth is a no-op (defers) with no sheet or an unlaid-out container', () => {
    installDom();
    const onScaleChange = vi.fn();
    // No worksheet.
    const a = mount(null, { onScaleChange });
    a.v.fitWidth();
    expect(a.v.getScale()).toBe(1);
    // Zero-width container ⇒ fitScale returns 0 ⇒ defer.
    const b = mount(makeSheet(), { onScaleChange }, { cw: 0, ch: 0 });
    b.v.fitWidth();
    expect(b.v.getScale()).toBe(1);
    expect(onScaleChange).not.toHaveBeenCalled();
  });

  // IX9 F1 — family-unified pre-load setScale semantics (pinned across all five
  // viewers): a setScale before load is LATCHED and applied to the first render
  // (cellScale is read by every subsequent sheet render).
  it('setScale before load/layout is latched and applied once established (IX9 F1)', () => {
    installDom();
    // No worksheet loaded yet — setScale must latch (renderCurrentSheet no-ops).
    const v = new XlsxViewer(makeContainer() as unknown as HTMLElement, {});
    v.setScale(1.5);
    expect(v.getScale()).toBe(1.5); // latched; the first showSheet renders at it
  });

  it('a pre-load setScale latch is clamped to [zoomMin, zoomMax] (IX9 F1)', () => {
    installDom();
    const v = new XlsxViewer(makeContainer() as unknown as HTMLElement, { zoomMin: 0.5, zoomMax: 3 });
    v.setScale(100);
    expect(v.getScale()).toBe(3); // latched pre-clamped
  });
});

/**
 * Non-regression: the pre-IX9 slider position ↔ scale mapping (PR #315's
 * "100% dead-center piecewise-linear" behaviour) is unchanged. These call the
 * private helpers directly so a future contract change can't silently alter the
 * slider feel.
 */
describe('XlsxViewer zoom slider mapping (pre-IX9 non-regression)', () => {
  it('slider position 50 maps to 100% for any bounds', () => {
    installDom();
    const { v } = mount(makeSheet());
    const priv = v as unknown as {
      zoomPosToScale(p: number, min: number, max: number): number;
      zoomScaleToPos(s: number, min: number, max: number): number;
    };
    expect(priv.zoomPosToScale(50, 0.1, 4)).toBeCloseTo(1, 10);
    expect(priv.zoomScaleToPos(1, 0.1, 4)).toBeCloseTo(50, 10);
    // Each half is its own linear segment.
    expect(priv.zoomPosToScale(0, 0.1, 4)).toBeCloseTo(0.1, 10);
    expect(priv.zoomPosToScale(100, 0.1, 4)).toBeCloseTo(4, 10);
  });
});
