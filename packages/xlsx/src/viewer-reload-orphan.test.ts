import { describe, it, expect, afterEach, vi } from 'vitest';
import { XlsxViewer, type XlsxViewerOptions } from './viewer.js';
import { XlsxWorkbook } from './workbook.js';
import { installDom, makeContainer } from './viewer-destroy-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/** A minimal XlsxWorkbook stand-in covering exactly the surface the viewer's
 *  load() → buildTabs() path touches, plus a `destroy` spy so SC20 can assert the
 *  previous workbook (its worker + WASM) is torn down on a re-load. */
function fakeWorkbook() {
  const destroy = vi.fn();
  const wb = {
    sheetNames: ['Sheet1'],
    tabColors: {} as Record<number, string>,
    destroy,
    getWorksheet: vi.fn().mockResolvedValue(undefined),
  };
  return { wb: wb as unknown as XlsxWorkbook, destroy };
}

/**
 * SC20: a second `load()` must not orphan the previous workbook (its worker +
 * pinned WASM). The fix is an atomic swap — load the new workbook first, destroy
 * the old one only on success.
 */
describe('XlsxViewer.load() — no orphaned workbook on re-load (SC20)', () => {
  function build(opts: XlsxViewerOptions = {}) {
    installDom();
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement, opts);
    // Isolate SC20 from the sheet-render path: showSheet needs a full worksheet
    // model to lay out, which is out of scope here. The engine-swap happens in
    // load() BEFORE showSheet, so a resolved no-op keeps this test on the leak.
    vi.spyOn(
      v as unknown as { showSheet: (i: number) => Promise<void> },
      'showSheet',
    ).mockResolvedValue(undefined);
    return { v };
  }

  it('destroys the previous workbook when load() is called again', async () => {
    const { v } = build();
    const a = fakeWorkbook();
    const b = fakeWorkbook();
    const loadSpy = vi
      .spyOn(XlsxWorkbook, 'load')
      .mockResolvedValueOnce(a.wb)
      .mockResolvedValueOnce(b.wb);

    await v.load('one.xlsx');
    expect(a.destroy).not.toHaveBeenCalled();

    await v.load('two.xlsx');
    expect(loadSpy).toHaveBeenCalledTimes(2);
    expect(a.destroy).toHaveBeenCalledTimes(1);
    expect(b.destroy).not.toHaveBeenCalled();

    v.destroy();
    expect(b.destroy).toHaveBeenCalledTimes(1);
  });

  it('keeps the current workbook when the re-load fails (atomic swap)', async () => {
    const onError = vi.fn();
    const { v } = build({ onError });
    const a = fakeWorkbook();
    vi.spyOn(XlsxWorkbook, 'load')
      .mockResolvedValueOnce(a.wb)
      .mockRejectedValueOnce(new Error('boom'));

    await v.load('one.xlsx');

    await v.load('bad.xlsx');
    expect(onError).toHaveBeenCalledTimes(1);
    expect(a.destroy).not.toHaveBeenCalled();

    v.destroy();
    expect(a.destroy).toHaveBeenCalledTimes(1);
  });
});
