import { describe, it, expect, afterEach, vi } from 'vitest';
import { PptxViewer } from './viewer.js';
import { PptxPresentation } from './presentation.js';
import { installDom, makeEl, FakePptxEngine } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const SLIDE_W_EMU = 9144000;
const SLIDE_H_EMU = 6858000;

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
describe('PptxViewer — destroy-during-render guard', () => {
  function mount() {
    installDom();
    return { canvas: makeEl('canvas') };
  }

  it('swallows a render rejection that lands after destroy() (no onError)', async () => {
    const { canvas } = mount();
    // deferred=true: renderSlide returns a promise the test rejects manually, so
    // the first-slide render is still in flight across destroy().
    const engine = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU, 'main', true);
    vi.spyOn(PptxPresentation, 'load').mockResolvedValue(engine.asPres());
    const onError = vi.fn();
    const v = new PptxViewer(canvas as unknown as HTMLCanvasElement, { onError });

    const loadPromise = v.load('x.pptx');
    await flushMicrotasks();
    const call = engine.renderCalls[engine.renderCalls.length - 1];
    expect(call).toBeDefined();

    v.destroy();
    call.reject(new Error('late render boom'));
    await loadPromise;
    await Promise.resolve();

    expect(onError).not.toHaveBeenCalled();
  });

  it('swallows a render rejection after destroy() with no onError (no console.error)', async () => {
    const { canvas } = mount();
    const spy = vi.spyOn(console, 'error').mockImplementation(() => {});
    const engine = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU, 'main', true);
    vi.spyOn(PptxPresentation, 'load').mockResolvedValue(engine.asPres());
    const v = new PptxViewer(canvas as unknown as HTMLCanvasElement);

    const loadPromise = v.load('x.pptx');
    await flushMicrotasks();
    const call = engine.renderCalls[engine.renderCalls.length - 1];
    v.destroy();
    call.reject(new Error('late render boom'));
    await loadPromise;
    await Promise.resolve();

    expect(spy).not.toHaveBeenCalled();
  });
});
