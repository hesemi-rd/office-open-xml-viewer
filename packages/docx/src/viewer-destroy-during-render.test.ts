import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxViewer } from './viewer.js';
import { DocxDocument } from './document.js';
import { installDom, makeEl, FakeDocxEngine } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const A4 = [{ widthPt: 595, heightPt: 842 }];

/** Drain the microtask queue so an awaited chain (mocked load → deferred render)
 *  reaches its next suspension point before the test inspects recorded calls. */
async function flushMicrotasks(): Promise<void> {
  for (let i = 0; i < 4; i++) await Promise.resolve();
}

/**
 * Destroy-during-render guard: a render dispatched before `destroy()` can reject
 * AFTER teardown (the worker/WASM resolve late). The `_destroyed` flag (set as the
 * first line of `destroy()`) makes `_reportRenderError` swallow such a rejection
 * so it never fires `onError` / `console.error` on a dead viewer — parity with the
 * scroll viewers' existing `_destroyed` guard.
 */
describe('DocxViewer — destroy-during-render guard', () => {
  function mount() {
    installDom();
    return { canvas: makeEl('canvas') };
  }

  it('swallows a render rejection that lands after destroy() (no onError)', async () => {
    const { canvas } = mount();
    // deferred=true: renderPage returns a promise the test rejects manually, so
    // the first-page render is still in flight across destroy().
    const engine = new FakeDocxEngine(2, A4, 'main', true);
    vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
    const onError = vi.fn();
    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement, { onError });

    // load() awaits the first render, which is deferred — kick it off but don't
    // await, so the render call is recorded and pending. Flush the microtasks that
    // resolve the (mocked) DocxDocument.load and reach the deferred renderPage.
    const loadPromise = v.load('x.docx');
    await flushMicrotasks();
    // The initial render dispatched by load() is the last recorded renderPage call.
    const call = engine.renderCalls[engine.renderCalls.length - 1];
    expect(call).toBeDefined();

    // Tear down while the render is in flight, THEN let it reject.
    v.destroy();
    call.reject(new Error('late render boom'));
    await loadPromise;
    // Give the rejection a microtask to propagate into _render's catch.
    await Promise.resolve();

    expect(onError).not.toHaveBeenCalled();
  });

  it('swallows a render rejection after destroy() with no onError (no console.error)', async () => {
    const { canvas } = mount();
    const spy = vi.spyOn(console, 'error').mockImplementation(() => {});
    const engine = new FakeDocxEngine(2, A4, 'main', true);
    vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement);

    const loadPromise = v.load('x.docx');
    await flushMicrotasks();
    const call = engine.renderCalls[engine.renderCalls.length - 1];
    v.destroy();
    call.reject(new Error('late render boom'));
    await loadPromise;
    await Promise.resolve();

    expect(spy).not.toHaveBeenCalled();
  });
});
