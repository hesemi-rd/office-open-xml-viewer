import { describe, it, expect, afterEach, vi } from 'vitest';
import { XlsxViewer, type XlsxViewerOptions } from './viewer.js';
import { installDom, makeContainer, type FakeEl } from './viewer-destroy-test-dom.js';
import type { Worksheet } from './types.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

function emptyWorksheet(): Worksheet {
  return {
    name: 'Sheet1',
    rows: [],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 64,
    defaultRowHeight: 20,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    charts: [],
    images: [],
    shapeGroups: [],
  } as unknown as Worksheet;
}

/** A worker-mode viewer whose `renderViewportToBitmap` returns a promise the test
 *  rejects manually, so the render is in flight across `destroy()`. */
function buildDeferredReject(opts: XlsxViewerOptions = {}) {
  installDom();
  const container = makeContainer();
  const v = new XlsxViewer(container as unknown as HTMLElement, { mode: 'worker', ...opts });

  let reject!: (e: Error) => void;
  const bitmap = new Promise((_res, rej) => {
    reject = rej;
  });
  const renderViewportToBitmap = vi.fn(() => bitmap);
  const fakeWb = { renderViewportToBitmap, sheetNames: ['Sheet1'], sheetCount: 1, destroy: vi.fn() };

  const priv = v as unknown as {
    wb: unknown;
    currentWorksheet: Worksheet;
    currentSheet: number;
    canvasArea: FakeEl;
    renderCurrentSheet: () => Promise<void>;
  };
  priv.wb = fakeWb;
  priv.currentWorksheet = emptyWorksheet();
  priv.currentSheet = 0;
  priv.canvasArea.clientWidth = 800;
  priv.canvasArea.clientHeight = 600;

  return { v, render: () => priv.renderCurrentSheet(), reject: (e: Error) => reject(e) };
}

/**
 * Destroy-during-render guard: a render dispatched before `destroy()` can reject
 * AFTER teardown (the worker/WASM resolve late). The `_destroyed` flag (set as the
 * first line of `destroy()`) makes `_reportRenderError` swallow such a rejection
 * so it never fires `onError` / `console.error` on a dead viewer — parity with the
 * scroll viewers' existing `_destroyed` guard.
 */
describe('XlsxViewer — destroy-during-render guard', () => {
  it('swallows a render rejection that lands after destroy() (no onError)', async () => {
    const onError = vi.fn();
    const { v, render, reject } = buildDeferredReject({ onError });

    const p = render(); // dispatch the worker render; it's now in flight
    v.destroy(); // tear down while in flight
    reject(new Error('late render boom'));
    await p;

    expect(onError).not.toHaveBeenCalled();
  });

  it('swallows a render rejection after destroy() with no onError (no console.error)', async () => {
    const spy = vi.spyOn(console, 'error').mockImplementation(() => {});
    const { v, render, reject } = buildDeferredReject();

    const p = render();
    v.destroy();
    reject(new Error('late render boom'));
    await p;

    expect(spy).not.toHaveBeenCalled();
  });
});
