import { describe, it, expect, afterEach, vi } from 'vitest';
import { PptxScrollViewer } from './scroll-viewer.js';
import { installDom, makeContainer, FakePptxEngine, type FakeEl } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * IX9 — PptxScrollViewer's slice of the shared
 * {@link import('@silurus/ooxml-core').ZoomableViewer} contract. Mirrors the docx
 * scroll-viewer contract test: the pre-existing ABSOLUTE `setScale`, Ctrl-wheel
 * zoom and resize re-fit stay verbatim; IX9 adds getScale / zoomIn / zoomOut /
 * fitWidth / fitPage + onScaleChange over the same absolute `_scale` (a slide
 * draws at `slideWidth/EMU_PER_PX × _scale`).
 */
const SLIDE_W_EMU = 9525 * 200; // natural 200 px
const SLIDE_H_EMU = 9525 * 120; // natural 120 px

function setup(opts: Record<string, unknown> = {}, host = { w: 200, h: 400 }) {
  installDom();
  const container = makeContainer(host.w, host.h);
  const engine = new FakePptxEngine(20, SLIDE_W_EMU, SLIDE_H_EMU);
  const v = new PptxScrollViewer(container as unknown as HTMLElement, {
    presentation: engine.asPres(),
    gap: 10,
    paddingTop: 0,
    paddingBottom: 0,
    paddingLeft: 0,
    paddingRight: 0,
    zoomMin: 0.1,
    zoomMax: 4,
    ...opts,
  });
  const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
  scrollHost.clientHeight = host.h;
  scrollHost.clientWidth = host.w;
  v.relayout();
  return { v, scrollHost, engine, container };
}

describe('PptxScrollViewer IX9 zoom contract', () => {
  it('getScale() returns the absolute factor (the base fit after load)', () => {
    const { v } = setup();
    // base = 200 / 200 = 1.0
    expect(v.getScale()).toBeCloseTo(1.0, 6);
    expect(v.getScale()).toBeCloseTo(v.scaleForTest(), 10);
    v.destroy();
  });

  it('getScale() is 1 before a scale is established (no width yet)', () => {
    installDom();
    const engine = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(makeContainer(0, 0) as unknown as HTMLElement, {
      presentation: engine.asPres(),
    });
    expect(v.getScale()).toBe(1);
    v.destroy();
  });

  it('setScale fires onScaleChange with the new factor on a change only', () => {
    const onScaleChange = vi.fn();
    const { v } = setup({ onScaleChange });
    v.setScale(2);
    expect(v.getScale()).toBeCloseTo(2, 6);
    expect(onScaleChange).toHaveBeenCalledTimes(1);
    expect(onScaleChange).toHaveBeenCalledWith(2);
    onScaleChange.mockClear();
    v.setScale(2); // unchanged
    expect(onScaleChange).not.toHaveBeenCalled();
    v.destroy();
  });

  it('zoomIn / zoomOut walk the shared ladder from the base 1.0', () => {
    const { v } = setup();
    expect(v.getScale()).toBeCloseTo(1.0, 6);
    v.zoomIn();
    expect(v.getScale()).toBeCloseTo(1.1, 6);
    v.zoomIn();
    expect(v.getScale()).toBeCloseTo(1.25, 6);
    v.zoomOut();
    expect(v.getScale()).toBeCloseTo(1.1, 6);
    v.destroy();
  });

  it('fitWidth restores the width-fit base after a zoom', () => {
    const { v } = setup();
    v.setScale(4);
    expect(v.getScale()).toBeCloseTo(4, 6);
    v.fitWidth();
    expect(v.getScale()).toBeCloseTo(1.0, 6);
    v.destroy();
  });

  it('fitPage takes the tighter of width/height fit', () => {
    // Container 200 wide × 60 tall. Natural slide 200 × 120.
    // widthfit = 200/200 = 1.0; heightfit = 60/120 = 0.5 ⇒ page-fit 0.5.
    const { v } = setup({}, { w: 200, h: 60 });
    v.fitPage();
    expect(v.getScale()).toBeCloseTo(0.5, 5);
    v.destroy();
  });

  it('the wheel-zoom path also notifies through onScaleChange', () => {
    const onScaleChange = vi.fn();
    const { v, scrollHost } = setup({ onScaleChange });
    scrollHost.dispatch('wheel', { ctrlKey: true, deltaY: -50, preventDefault() {} });
    expect(onScaleChange).toHaveBeenCalledTimes(1);
    expect(onScaleChange.mock.calls[0][0]).toBeGreaterThan(1.0);
    v.destroy();
  });

  // IX9 F1 — family-unified pre-load setScale semantics: a setScale before the
  // layout establishes must be LATCHED and applied once it does (the
  // single-canvas viewers honour a pre-load setScale on their first render; the
  // scroll viewers used to silently drop it).
  it('setScale before load/layout is latched and applied once established (IX9 F1)', () => {
    installDom();
    const container = makeContainer(0, 0); // zero-width ⇒ fit deferred, nothing established
    const engine = new FakePptxEngine(20, SLIDE_W_EMU, SLIDE_H_EMU);
    const onScaleChange = vi.fn();
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
      paddingTop: 0,
      paddingBottom: 0,
      paddingLeft: 0,
      paddingRight: 0,
      onScaleChange,
    });
    v.setScale(2); // pre-establishment: latched, not dropped
    expect(v.getScale()).toBe(2); // getScale reports the pending factor
    expect(onScaleChange).not.toHaveBeenCalled(); // fires at APPLICATION time
    // Container gains width ⇒ the next relayout establishes and applies the latch.
    container.clientWidth = 200;
    container.clientHeight = 400;
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientWidth = 200;
    scrollHost.clientHeight = 400;
    v.relayout();
    expect(v.scaleForTest()).toBeCloseTo(2, 6); // applied over the 1.0 base fit
    expect(v.getScale()).toBeCloseTo(2, 6);
    expect(onScaleChange).toHaveBeenCalledTimes(1);
    expect(onScaleChange).toHaveBeenCalledWith(2);
    // The base fit itself is still the true base (resize re-fit multiplier intact).
    expect(v.baseScaleForTest()).toBeCloseTo(1.0, 6);
    v.destroy();
  });

  it('a pre-establishment setScale latch is clamped to [zoomMin, zoomMax] (IX9 F1)', () => {
    installDom();
    const engine = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = new PptxScrollViewer(makeContainer(0, 0) as unknown as HTMLElement, {
      presentation: engine.asPres(),
      zoomMin: 0.5,
      zoomMax: 3,
    });
    v.setScale(100);
    expect(v.getScale()).toBe(3); // latched pre-clamped
    v.destroy();
  });
});
