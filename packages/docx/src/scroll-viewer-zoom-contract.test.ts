import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxScrollViewer } from './scroll-viewer.js';
import { installDom, makeContainer, FakeDocxEngine, type FakeEl } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * IX9 — DocxScrollViewer's slice of the shared
 * {@link import('@silurus/ooxml-core').ZoomableViewer} contract. The viewer
 * already had an ABSOLUTE `setScale(scale)` (1 = 100%, the base fit ≠ 1),
 * Ctrl-wheel zoom, and a container-resize re-fit; IX9 keeps that verbatim and
 * layers on `getScale` / `zoomIn` / `zoomOut` / `fitWidth` / `fitPage` and the
 * `onScaleChange` notification. The absolute `_scale` IS the contract's
 * user-facing factor (a page draws at `widthPt × PT_TO_PX × _scale`), so the new
 * methods operate directly on it.
 */

/** A page 100pt × 200pt (natural CSS 133.33 × 266.67 px at 100%). */
const PAGE = { widthPt: 100, heightPt: 200 };

function setup(opts: Record<string, unknown> = {}, host = { w: 200, h: 400 }) {
  installDom();
  const container = makeContainer(host.w, host.h);
  const engine = new FakeDocxEngine(5, [PAGE]);
  const v = new DocxScrollViewer(container as unknown as HTMLElement, {
    document: engine.asDoc(),
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

describe('DocxScrollViewer IX9 zoom contract', () => {
  it('getScale() returns the absolute factor (the base fit after load)', () => {
    const { v } = setup();
    // base = 200 / (100 × 4/3) = 1.5
    expect(v.getScale()).toBeCloseTo(1.5, 5);
    expect(v.getScale()).toBeCloseTo(v.scaleForTest(), 10);
    v.destroy();
  });

  it('getScale() is 1 before a scale is established (no width yet)', () => {
    installDom();
    const engine = new FakeDocxEngine(3, [PAGE]);
    const v = new DocxScrollViewer(makeContainer(0, 0) as unknown as HTMLElement, {
      document: engine.asDoc(),
    });
    expect(v.getScale()).toBe(1);
    v.destroy();
  });

  it('setScale fires onScaleChange with the new factor on a change only', () => {
    const onScaleChange = vi.fn();
    const { v } = setup({ onScaleChange });
    v.setScale(2);
    expect(v.getScale()).toBeCloseTo(2, 5);
    expect(onScaleChange).toHaveBeenCalledTimes(1);
    expect(onScaleChange).toHaveBeenCalledWith(2);
    onScaleChange.mockClear();
    v.setScale(2); // unchanged
    expect(onScaleChange).not.toHaveBeenCalled();
    v.destroy();
  });

  it('zoomIn / zoomOut walk the shared ladder (off-base start snaps on)', () => {
    const { v } = setup();
    // Start at base 1.5 (an off-ladder value). First zoomIn snaps to 1.75.
    v.zoomIn();
    expect(v.getScale()).toBeCloseTo(1.75, 5);
    v.zoomIn();
    expect(v.getScale()).toBeCloseTo(2, 5);
    v.zoomOut();
    expect(v.getScale()).toBeCloseTo(1.75, 5);
    v.zoomOut();
    expect(v.getScale()).toBeCloseTo(1.5, 5); // ladder rung between 1.5 base
    v.destroy();
  });

  it('fitWidth restores the width-fit base after a zoom', () => {
    const { v } = setup();
    v.setScale(4); // zoom right in
    expect(v.getScale()).toBeCloseTo(4, 5);
    v.fitWidth();
    // Back to the width-fit base = 200 / (100 × 4/3) = 1.5.
    expect(v.getScale()).toBeCloseTo(1.5, 5);
    v.destroy();
  });

  it('fitPage takes the tighter of width/height fit', () => {
    // Container 200 wide × 200 tall. Natural page 133.33 × 266.67.
    // widthfit = 200/133.33 = 1.5; heightfit = 200/266.67 = 0.75 ⇒ page-fit 0.75.
    const { v } = setup({}, { w: 200, h: 200 });
    v.fitPage();
    expect(v.getScale()).toBeCloseTo(0.75, 4);
    v.destroy();
  });

  it('the wheel-zoom path also notifies through onScaleChange', () => {
    const onScaleChange = vi.fn();
    const { v, scrollHost } = setup({ onScaleChange });
    // Ctrl+wheel up (deltaY < 0 ⇒ zoom in). The handler routes through setScale.
    scrollHost.dispatch('wheel', { ctrlKey: true, deltaY: -50, preventDefault() {} });
    expect(onScaleChange).toHaveBeenCalledTimes(1);
    expect(onScaleChange.mock.calls[0][0]).toBeGreaterThan(1.5);
    v.destroy();
  });
});
