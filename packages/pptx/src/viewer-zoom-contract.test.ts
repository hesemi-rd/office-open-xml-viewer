import { describe, it, expect, afterEach, vi } from 'vitest';
import { PptxViewer } from './viewer.js';
import { PptxPresentation } from './presentation.js';
import { EMU_PER_PX } from '@silurus/ooxml-core';
import { installDom, makeEl, FakePptxEngine, type FakeEl } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const SLIDE_W_EMU = 9144000; // 960 CSS px at 100%
const SLIDE_H_EMU = 6858000; // 720 CSS px at 100%
const NATURAL_W = SLIDE_W_EMU / EMU_PER_PX; // 960
const NATURAL_H = SLIDE_H_EMU / EMU_PER_PX; // 720

/** Mount a PptxViewer over a FakePptxEngine, optionally with a laid-out DOM
 *  parent so fit has a container to measure. */
async function mount(opts: Record<string, unknown> = {}, containerSize?: { w: number; h: number }) {
  installDom();
  const canvas = makeEl('canvas');
  if (containerSize) {
    const parent = makeEl('div');
    parent.clientWidth = containerSize.w;
    parent.clientHeight = containerSize.h;
    parent.appendChild(canvas);
  }
  const engine = new FakePptxEngine(4, SLIDE_W_EMU, SLIDE_H_EMU);
  vi.spyOn(PptxPresentation, 'load').mockResolvedValue(engine.asPres());
  const v = new PptxViewer(canvas as unknown as HTMLCanvasElement, opts);
  await v.load('x.pptx');
  return { v, engine, canvas: canvas as FakeEl };
}

function lastRenderWidth(engine: FakePptxEngine): number | undefined {
  const w = engine.renderCalls;
  return w[w.length - 1]?.width;
}

describe('PptxViewer IX9 zoom contract — byte-identical default', () => {
  it('renders at opts.width verbatim when no zoom method is called', async () => {
    const { engine } = await mount({ width: 1200 });
    expect(lastRenderWidth(engine)).toBe(1200);
  });

  it('renders at the offsetWidth||960 fallback when opts.width is unset', async () => {
    const { engine } = await mount(); // fake canvas.offsetWidth = 0 ⇒ 960
    expect(lastRenderWidth(engine)).toBe(960);
  });

  it('getScale() reflects the effective render width before any zoom call', async () => {
    const { v } = await mount({ width: NATURAL_W * 2 }); // 1920
    expect(v.getScale()).toBeCloseTo(2, 6);
  });

  it('getScale() is 1 at the natural default width', async () => {
    const { v } = await mount(); // 960 == natural
    expect(v.getScale()).toBeCloseTo(1, 6);
  });
});

describe('PptxViewer IX9 zoom contract — setScale / steppers', () => {
  it('setScale renders at naturalWidth × scale and fires onScaleChange', async () => {
    const onScaleChange = vi.fn();
    const { v, engine } = await mount({ onScaleChange });
    await v.setScale(1.5);
    expect(v.getScale()).toBe(1.5);
    expect(lastRenderWidth(engine)).toBe(Math.round(NATURAL_W * 1.5));
    expect(onScaleChange).toHaveBeenCalledTimes(1);
    expect(onScaleChange).toHaveBeenCalledWith(1.5);
  });

  it('setScale clamps to [zoomMin, zoomMax]', async () => {
    const { v } = await mount({ zoomMin: 0.5, zoomMax: 2 });
    await v.setScale(10);
    expect(v.getScale()).toBe(2);
    await v.setScale(0.01);
    expect(v.getScale()).toBe(0.5);
  });

  it('setScale does not fire onScaleChange when the factor is unchanged', async () => {
    const onScaleChange = vi.fn();
    const { v } = await mount({ onScaleChange });
    await v.setScale(1.5);
    onScaleChange.mockClear();
    await v.setScale(1.5);
    expect(onScaleChange).not.toHaveBeenCalled();
  });

  it('zoomIn / zoomOut walk the shared ladder from 100%', async () => {
    const { v } = await mount();
    expect(v.getScale()).toBeCloseTo(1, 6);
    await v.zoomIn();
    expect(v.getScale()).toBe(1.1);
    await v.zoomIn();
    expect(v.getScale()).toBe(1.25);
    await v.zoomOut();
    expect(v.getScale()).toBe(1.1);
  });
});

describe('PptxViewer IX9 zoom contract — fitWidth / fitPage', () => {
  it('fitWidth sets the scale that spans the slide width in the container', async () => {
    const { v } = await mount({}, { w: NATURAL_W / 2, h: 10000 });
    await v.fitWidth();
    expect(v.getScale()).toBeCloseTo(0.5, 6);
  });

  it('fitPage takes the tighter of width/height', async () => {
    // width fits at 1.0, height fits at 0.5 ⇒ min = 0.5.
    const { v } = await mount({}, { w: NATURAL_W, h: NATURAL_H / 2 });
    await v.fitPage();
    expect(v.getScale()).toBeCloseTo(0.5, 6);
  });

  it('fitWidth defers (no-op) with no container to measure', async () => {
    const onScaleChange = vi.fn();
    const { v } = await mount({ onScaleChange }); // detached canvas
    await v.fitWidth();
    expect(v.getScale()).toBeCloseTo(1, 6);
    expect(onScaleChange).not.toHaveBeenCalled();
  });
});
