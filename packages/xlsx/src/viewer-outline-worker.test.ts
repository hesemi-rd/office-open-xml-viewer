import { describe, it, expect, afterEach, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import { installDom, makeContainer, type FakeEl } from './viewer-destroy-test-dom.js';
import { applySizeOverrides, type WireSizeOverrides } from './worker-protocol.js';
import type { Worksheet } from './types.js';
import type { OutlineLayout } from './outline.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * Worker-mode model-sync for view-only size mutations.
 *
 * The render worker draws from its own worker-local parsed-sheet cache
 * (render-worker.ts parseSheetLocally), so a main-thread Worksheet mutation —
 * outline collapse/expand mapping bands to the size-0 hidden encoding, or
 * drag-to-resize (#567) — never reaches the worker on its own: the gutter and
 * overlays would update while the grid bitmap kept the file's original sizes
 * (rows stayed hidden after an expand click). These tests pin the override
 * channel that closes the gap:
 *
 * 1. every worker `renderViewport` request issued after a mutation carries
 *    `opts.sizeOverrides` describing the touched bands' CURRENT model sizes, and
 * 2. applying that wire payload to a pristine copy of the sheet (exactly what
 *    the worker does before drawing) converges its size maps to the main-thread
 *    model — the main↔worker equivalence that makes both modes lay out the
 *    same grid.
 */

/** The synthetic outline fixture's row model: rows 2-9 grouped (3 nested
 *  levels), detail rows 4-7 collapsed-hidden (height 0), row 8 the collapsed
 *  level-2 summary. Mirrors outline-fixture.xlsx. */
function outlineWorksheet(): Worksheet {
  const row = (
    index: number,
    outlineLevel = 0,
    opts: { collapsed?: boolean; hidden?: boolean } = {},
  ) => ({ index, height: opts.hidden ? 0 : null, cells: [], outlineLevel, ...opts });
  return {
    name: 'Outline',
    rows: [
      row(1),
      row(2, 1),
      row(3, 2),
      row(4, 3, { hidden: true }),
      row(5, 3, { hidden: true }),
      row(6, 3, { hidden: true }),
      row(7, 3, { hidden: true }),
      row(8, 2, { collapsed: true }),
      row(9, 1),
      row(10),
    ],
    colWidths: {},
    rowHeights: { 4: 0, 5: 0, 6: 0, 7: 0 },
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    charts: [],
    images: [],
    shapeGroups: [],
    outlinePr: { summaryBelow: true, summaryRight: true },
  } as unknown as Worksheet;
}

/** Private-surface shape of the viewer these tests drive. */
interface ViewerPriv {
  wb: unknown;
  currentWorksheet: Worksheet;
  currentSheet: number;
  canvasArea: FakeEl;
  rowOutline: OutlineLayout | null;
  resizeDrag: { kind: 'col' | 'row'; index: number; originScaled: number; mdw: number } | null;
  buildOutline(ws: Worksheet): void;
  applyGroupToggle(group: OutlineLayout['groups'][number], axis: 'row' | 'col'): void;
  applyResize(clientX: number, clientY: number): void;
  renderCurrentSheet(): Promise<void>;
}

/** Worker-mode viewer over the outline worksheet, with a fake workbook whose
 *  renderViewportToBitmap records every request (never resolves — only the
 *  synchronously-sent request matters here). */
function buildWorker() {
  installDom();
  const container = makeContainer();
  const v = new XlsxViewer(container as unknown as HTMLElement, { mode: 'worker' });
  const renderViewportToBitmap = vi.fn(() => new Promise<never>(() => undefined));
  const fakeWb = {
    renderViewportToBitmap,
    sheetNames: ['Outline'],
    sheetCount: 1,
    destroy: vi.fn(),
  };
  const priv = v as unknown as ViewerPriv;
  priv.wb = fakeWb;
  priv.currentWorksheet = outlineWorksheet();
  priv.currentSheet = 0;
  priv.canvasArea.clientWidth = 800;
  priv.canvasArea.clientHeight = 600;
  priv.buildOutline(priv.currentWorksheet);
  return { v, priv, renderViewportToBitmap };
}

/** The `sizeOverrides` of the most recent render request, if any. */
function lastOverrides(fn: ReturnType<typeof vi.fn>): WireSizeOverrides | undefined {
  const call = fn.mock.calls.at(-1) as unknown[] | undefined;
  return (call?.[2] as { sizeOverrides?: WireSizeOverrides } | undefined)?.sizeOverrides;
}

describe('worker-mode outline collapse/expand reaches the grid bitmap', () => {
  it('expanding the collapsed group sends row overrides that reveal rows 4-7', () => {
    const { priv, renderViewportToBitmap } = buildWorker();
    const l3 = priv.rowOutline?.groups.find((g) => g.level === 3);
    expect(l3?.collapsed).toBe(true);

    // Expand (the sync-render fallback fires renderCurrentSheet immediately).
    priv.applyGroupToggle(l3 as OutlineLayout['groups'][number], 'row');
    expect(renderViewportToBitmap).toHaveBeenCalled();
    const o = lastOverrides(renderViewportToBitmap);
    // Rows 4-7 were revealed: the model has NO entry (default height) ⇒ null.
    expect(o?.rows).toMatchObject({ 4: null, 5: null, 6: null, 7: null });

    // Worker-side application converges a pristine sheet to the main model —
    // the property that makes the worker's next bitmap actually re-lay the
    // revealed rows (main↔worker equivalence at the model level).
    const workerSheet = outlineWorksheet();
    applySizeOverrides(workerSheet, o);
    expect(workerSheet.rowHeights).toEqual(priv.currentWorksheet.rowHeights);
    expect(workerSheet.rowHeights[4]).toBeUndefined(); // was 0 (hidden) pre-override
  });

  it('re-collapsing sends rows back as 0-height overrides', () => {
    const { priv, renderViewportToBitmap } = buildWorker();
    const expand = priv.rowOutline?.groups.find((g) => g.level === 3);
    priv.applyGroupToggle(expand as OutlineLayout['groups'][number], 'row');
    // The layout was rebuilt after the expand — fetch the group's new object.
    const collapse = priv.rowOutline?.groups.find((g) => g.level === 3);
    expect(collapse?.collapsed).toBe(false);
    priv.applyGroupToggle(collapse as OutlineLayout['groups'][number], 'row');

    const o = lastOverrides(renderViewportToBitmap);
    expect(o?.rows).toMatchObject({ 4: 0, 5: 0, 6: 0, 7: 0 });

    const workerSheet = outlineWorksheet();
    applySizeOverrides(workerSheet, o);
    expect(workerSheet.rowHeights).toEqual(priv.currentWorksheet.rowHeights);
  });
});

describe('worker-mode drag-to-resize reaches the grid bitmap (#567 hole)', () => {
  it('a column resize sends the new width as a col override', () => {
    const { priv, renderViewportToBitmap } = buildWorker();
    priv.resizeDrag = { kind: 'col', index: 2, originScaled: 0, mdw: 7 };
    priv.applyResize(100, 0); // drag column B's right border to x=100

    const o = lastOverrides(renderViewportToBitmap);
    const newWidth = priv.currentWorksheet.colWidths[2];
    expect(newWidth).toBeGreaterThan(0);
    expect(o?.cols?.[2]).toBe(newWidth);

    const workerSheet = outlineWorksheet();
    applySizeOverrides(workerSheet, o);
    expect(workerSheet.colWidths[2]).toBe(newWidth);
  });

  it('a row resize sends the new height as a row override', () => {
    const { priv, renderViewportToBitmap } = buildWorker();
    priv.resizeDrag = { kind: 'row', index: 9, originScaled: 0, mdw: 7 };
    priv.applyResize(0, 60);

    const o = lastOverrides(renderViewportToBitmap);
    const newHeight = priv.currentWorksheet.rowHeights[9];
    expect(newHeight).toBeGreaterThan(0);
    expect(o?.rows?.[9]).toBe(newHeight);
  });
});

describe('applySizeOverrides wire semantics', () => {
  it('sets numeric values and deletes null entries on both axes', () => {
    const ws = outlineWorksheet();
    ws.colWidths[3] = 20;
    applySizeOverrides(ws, { rows: { 4: null, 9: 30 }, cols: { 3: null, 5: 12 } });
    expect(ws.rowHeights[4]).toBeUndefined();
    expect(ws.rowHeights[9]).toBe(30);
    expect(ws.colWidths[3]).toBeUndefined();
    expect(ws.colWidths[5]).toBe(12);
  });

  it('is idempotent (re-applying the same map is a no-op) and undefined-safe', () => {
    const ws = outlineWorksheet();
    const o: WireSizeOverrides = { rows: { 4: null, 5: 22 } };
    applySizeOverrides(ws, o);
    const snapshot = JSON.stringify(ws.rowHeights);
    applySizeOverrides(ws, o);
    expect(JSON.stringify(ws.rowHeights)).toBe(snapshot);
    applySizeOverrides(ws, undefined); // absent overrides: no mutation
    expect(JSON.stringify(ws.rowHeights)).toBe(snapshot);
  });
});

describe('override plumbing stays silent when nothing was mutated', () => {
  it('a plain render request carries NO sizeOverrides key', async () => {
    const { priv, renderViewportToBitmap } = buildWorker();
    void priv.renderCurrentSheet();
    expect(renderViewportToBitmap).toHaveBeenCalledTimes(1);
    const call = renderViewportToBitmap.mock.calls[0] as unknown[] | undefined;
    const opts = call?.[2] as Record<string, unknown> | undefined;
    expect(opts).toBeDefined();
    expect(opts !== undefined && 'sizeOverrides' in opts).toBe(false);
  });
});
