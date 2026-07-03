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

/** A FakePptxEngine whose `renderSlide` throws, to exercise the render-error
 *  path (the base fake resolves; override just renderSlide). */
function throwingEngine(): FakePptxEngine {
  const e = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
  (e as unknown as { renderSlide: () => Promise<void> }).renderSlide = () => {
    throw new Error('render boom');
  };
  return e;
}

/**
 * PD14: PptxViewer already routed render failures to `onError`; this pins that,
 * and that it now also `console.error`s when no `onError` is given (never
 * silent), so all three single-canvas viewers share one contract.
 */
describe('PptxViewer render error contract (PD14)', () => {
  function mount() {
    installDom();
    return { canvas: makeEl('canvas') };
  }

  it('routes a render failure to onError during load() and resolves', async () => {
    const { canvas } = mount();
    vi.spyOn(PptxPresentation, 'load').mockResolvedValue(throwingEngine().asPres());
    const onError = vi.fn();
    const v = new PptxViewer(canvas as unknown as HTMLCanvasElement, { onError });
    await expect(v.load('x.pptx')).resolves.toBeUndefined();
    expect(onError).toHaveBeenCalledTimes(1);
    expect((onError.mock.calls[0][0] as Error).message).toContain('boom');
    v.destroy();
  });

  it('routes a render failure to onError during goToSlide() (no rejection)', async () => {
    const { canvas } = mount();
    const good = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    vi.spyOn(PptxPresentation, 'load').mockResolvedValue(good.asPres());
    const onError = vi.fn();
    const v = new PptxViewer(canvas as unknown as HTMLCanvasElement, { onError });
    await v.load('x.pptx');
    expect(onError).not.toHaveBeenCalled();

    (good as unknown as { renderSlide: () => Promise<void> }).renderSlide = () => {
      throw new Error('nav boom');
    };
    await expect(v.nextSlide()).resolves.toBeUndefined();
    expect(onError).toHaveBeenCalledTimes(1);
    expect((onError.mock.calls[0][0] as Error).message).toContain('nav boom');
    v.destroy();
  });

  it('falls back to console.error when no onError is provided (never silent)', async () => {
    const { canvas } = mount();
    const spy = vi.spyOn(console, 'error').mockImplementation(() => {});
    vi.spyOn(PptxPresentation, 'load').mockResolvedValue(throwingEngine().asPres());
    const v = new PptxViewer(canvas as unknown as HTMLCanvasElement);
    await expect(v.load('x.pptx')).resolves.toBeUndefined();
    expect(spy).toHaveBeenCalledTimes(1);
    v.destroy();
  });
});
