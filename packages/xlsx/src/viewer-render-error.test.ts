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

/** A worker-mode viewer whose `renderViewportToBitmap` REJECTS, wired so
 *  `renderCurrentSheet` reaches the worker branch. */
function buildRejecting(opts: XlsxViewerOptions = {}) {
  installDom();
  const container = makeContainer();
  const v = new XlsxViewer(container as unknown as HTMLElement, { mode: 'worker', ...opts });

  const renderViewportToBitmap = vi.fn(() => Promise.reject(new Error('render boom')));
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

  return { v, render: () => priv.renderCurrentSheet() };
}

/**
 * PD14: a failed sheet render must follow the same contract as the scroll viewer
 * — invoke `onError` if provided, else `console.error` (never silent), and never
 * surface as an unhandled promise rejection. Before the fix `renderCurrentSheet`
 * had no try/catch and was called `void`-style from scroll/resize handlers, so a
 * render rejection became an unhandled rejection.
 */
describe('XlsxViewer render error contract (PD14)', () => {
  it('routes a render failure to onError and resolves (no rejection)', async () => {
    const onError = vi.fn();
    const { render } = buildRejecting({ onError });
    await expect(render()).resolves.toBeUndefined();
    expect(onError).toHaveBeenCalledTimes(1);
    expect(onError.mock.calls[0][0]).toBeInstanceOf(Error);
    expect((onError.mock.calls[0][0] as Error).message).toContain('boom');
  });

  it('falls back to console.error when no onError is provided (never silent)', async () => {
    const spy = vi.spyOn(console, 'error').mockImplementation(() => {});
    const { render } = buildRejecting();
    await expect(render()).resolves.toBeUndefined();
    expect(spy).toHaveBeenCalledTimes(1);
  });
});
