import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxViewer } from './viewer.js';
import { DocxDocument } from './document.js';
import { installDom, makeEl, FakeDocxEngine } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const A4 = [{ widthPt: 595, heightPt: 842 }];

/** A FakeDocxEngine whose `renderPage` throws, to exercise the render-error path
 *  (the base fake resolves; we override just renderPage). */
function throwingEngine(): FakeDocxEngine {
  const e = new FakeDocxEngine(2, A4);
  (e as unknown as { renderPage: () => Promise<void> }).renderPage = () => {
    throw new Error('render boom');
  };
  return e;
}

/**
 * PD14: a failed page render must follow the scroll-viewer contract — invoke
 * `onError` if provided, else `console.error` (never silent), and never
 * propagate (which from a `void`-style `nextPage()` would be an unhandled
 * rejection). Before the fix `_render` had no try/catch and rethrew to its
 * caller.
 */
describe('DocxViewer render error contract (PD14)', () => {
  function mount() {
    installDom();
    return { canvas: makeEl('canvas') };
  }

  it('routes a render failure to onError during load() and resolves', async () => {
    const { canvas } = mount();
    vi.spyOn(DocxDocument, 'load').mockResolvedValue(throwingEngine().asDoc());
    const onError = vi.fn();
    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement, { onError });
    await expect(v.load('x.docx')).resolves.toBeUndefined();
    expect(onError).toHaveBeenCalledTimes(1);
    expect((onError.mock.calls[0][0] as Error).message).toContain('boom');
    v.destroy();
  });

  it('routes a render failure to onError during goToPage() (no rejection)', async () => {
    const { canvas } = mount();
    const good = new FakeDocxEngine(2, A4);
    vi.spyOn(DocxDocument, 'load').mockResolvedValue(good.asDoc());
    const onError = vi.fn();
    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement, { onError });
    await v.load('x.docx');
    expect(onError).not.toHaveBeenCalled(); // clean first render

    // Now make subsequent renders throw and navigate.
    (good as unknown as { renderPage: () => Promise<void> }).renderPage = () => {
      throw new Error('nav boom');
    };
    await expect(v.nextPage()).resolves.toBeUndefined();
    expect(onError).toHaveBeenCalledTimes(1);
    expect((onError.mock.calls[0][0] as Error).message).toContain('nav boom');
    v.destroy();
  });

  it('falls back to console.error when no onError is provided (never silent)', async () => {
    const { canvas } = mount();
    const spy = vi.spyOn(console, 'error').mockImplementation(() => {});
    vi.spyOn(DocxDocument, 'load').mockResolvedValue(throwingEngine().asDoc());
    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement);
    await expect(v.load('x.docx')).resolves.toBeUndefined();
    expect(spy).toHaveBeenCalledTimes(1);
    v.destroy();
  });
});
