import { describe, it, expect, afterEach, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import { installDom, makeContainer, type FakeEl } from './viewer-destroy-test-dom.js';
import type { Worksheet } from './types.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * Outline gutter DOM parity (hotfix for the layouts-smoke regression).
 *
 * The three gutter canvases must exist in the DOM only while the shown sheet
 * actually has an outline. `display:none` is NOT enough: consumers that count
 * or index `<canvas>` elements see hidden nodes — the layouts smoke's
 * `page.locator('canvas').count()` broke on outline-free demo/sample-1
 * (expected N+1 canvases, got N+4) when the gutters were attached
 * unconditionally. An outline-free sheet must present the exact pre-outline
 * element set: one grid canvas, nothing else.
 */

function countCanvases(el: FakeEl): number {
  let n = el.tag === 'canvas' ? 1 : 0;
  for (const c of el.children) n += countCanvases(c);
  return n;
}

function worksheet(withOutline: boolean): Worksheet {
  return {
    name: 'S',
    rows: withOutline
      ? [
          { index: 1, height: null, cells: [] },
          { index: 2, height: null, cells: [], outlineLevel: 1 },
          { index: 3, height: null, cells: [], outlineLevel: 1 },
        ]
      : [{ index: 1, height: null, cells: [] }],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    charts: [],
    images: [],
    shapeGroups: [],
  } as unknown as Worksheet;
}

interface ViewerPriv {
  currentWorksheet: Worksheet;
  buildOutline(ws: Worksheet): void;
  layoutGutters(): void;
}

describe('outline gutter canvases are attached only while an outline exists', () => {
  it('a freshly constructed viewer has exactly ONE canvas (the grid)', () => {
    installDom();
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement);
    expect(countCanvases(container as unknown as FakeEl)).toBe(1);
    v.destroy();
  });

  it('gutter canvases appear with an outline sheet and disappear again on an outline-free sheet', () => {
    installDom();
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement);
    const priv = v as unknown as ViewerPriv;

    // Outline sheet: grid + row/col/corner gutters = 4 canvases.
    priv.currentWorksheet = worksheet(true);
    priv.buildOutline(priv.currentWorksheet);
    priv.layoutGutters();
    expect(countCanvases(container as unknown as FakeEl)).toBe(4);

    // Back to an outline-free sheet: exact pre-outline parity (1 canvas).
    priv.currentWorksheet = worksheet(false);
    priv.buildOutline(priv.currentWorksheet);
    priv.layoutGutters();
    expect(countCanvases(container as unknown as FakeEl)).toBe(1);

    // And an outline sheet again: the cached elements re-attach (4).
    priv.currentWorksheet = worksheet(true);
    priv.buildOutline(priv.currentWorksheet);
    priv.layoutGutters();
    expect(countCanvases(container as unknown as FakeEl)).toBe(4);
    v.destroy();
  });
});
