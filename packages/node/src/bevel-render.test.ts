import { describe, it, expect } from 'vitest';
import { Canvas, loadImage } from 'skia-canvas';
import { renderSlideNode } from './render';
import type { NodeCanvasFactory } from './render';
import type { Presentation, Slide, PictureElement } from '@silurus/ooxml-pptx';

/**
 * Regression test for the Node OffscreenCanvas shim.
 *
 * `packages/core/src/shape/effects.ts`'s `createAuxCanvas` needs an
 * `OffscreenCanvas` (or DOM `document`) to allocate the auxiliary canvas that
 * the pptx renderer's `paintBeveledFlat` uses to bake sp3d bevel-lip shading.
 * Under Node neither exists, so without the shim the renderer silently skips
 * the bevel and blits the flat image. This test renders the SAME synthetic
 * slide (a solid picture carrying an sp3d `bevelT`) twice — once WITH the shim
 * (factory passed → bevel shading active) and once WITHOUT (no factory → flat
 * fallback) — and asserts the rim band differs. If the shim regresses, both
 * renders match and this fails.
 */

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) => new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (buf) => loadImage(buf as Buffer) as unknown as ReturnType<NodeCanvasFactory['loadImage']>,
};

const g = globalThis as unknown as { createImageBitmap?: unknown; OffscreenCanvas?: unknown };

/** Build a small flat-colour PNG as a data URL using skia-canvas. */
async function flatPngDataUrl(w: number, h: number, color: string): Promise<string> {
  const c = new Canvas(w, h);
  const ctx = c.getContext('2d');
  ctx.fillStyle = color;
  ctx.fillRect(0, 0, w, h);
  const buf = await c.toBuffer('png');
  return `data:image/png;base64,${buf.toString('base64')}`;
}

const EMU_PER_PX = 9525; // 96 dpi

function buildBeveledSlide(dataUrl: string): Presentation {
  // A 300x300 px picture centred on a 400x400 px slide, with a chunky top
  // bevel so the lit/shadowed rim is unmistakable.
  const px = (n: number) => Math.round(n * EMU_PER_PX);
  const pic: PictureElement = {
    type: 'picture',
    x: px(50),
    y: px(50),
    width: px(300),
    height: px(300),
    rotation: 0,
    flipH: false,
    flipV: false,
    dataUrl,
    stroke: null,
    sp3d: {
      prstMaterial: 'warmMatte',
      // 360000 EMU ≈ 38 px bevel at 1:1; circle lip → strong gradient.
      bevelT: { w: 360000, h: 360000, prst: 'circle' },
    },
  };
  const slide: Slide = {
    index: 0,
    slideNumber: 1,
    background: null,
    elements: [pic],
  };
  return {
    slideWidth: px(400),
    slideHeight: px(400),
    slides: [slide],
    defaultTextColor: null,
    majorFont: null,
    minorFont: null,
  };
}

async function renderToRgba(
  presentation: Presentation,
  withShim: boolean,
): Promise<{ data: Uint8ClampedArray; width: number; height: number }> {
  const width = 400;
  const dpr = 1;
  const canvas = new Canvas(width * dpr, 400 * dpr);
  // createImageBitmap is always needed to decode the picture's dataUrl.
  const g2 = globalThis as unknown as { createImageBitmap?: unknown };
  const prevBitmap = g2.createImageBitmap;
  g2.createImageBitmap = async (source: Blob | ArrayBuffer | Uint8Array) => {
    let buf: ArrayBuffer | Uint8Array;
    if (source instanceof Uint8Array || source instanceof ArrayBuffer) buf = source;
    else buf = await (source as Blob).arrayBuffer();
    return loadImage(Buffer.from(buf as ArrayBuffer));
  };
  try {
    await renderSlideNode(
      canvas as unknown as Parameters<typeof renderSlideNode>[0],
      presentation,
      0,
      { width, dpr, ...(withShim ? { factory } : {}) },
    );
  } finally {
    g2.createImageBitmap = prevBitmap as typeof globalThis.createImageBitmap;
  }
  const ctx = canvas.getContext('2d');
  const img = ctx.getImageData(0, 0, width * dpr, 400 * dpr);
  return { data: img.data as unknown as Uint8ClampedArray, width: width * dpr, height: 400 * dpr };
}

describe('node bevel rendering (OffscreenCanvas shim)', () => {
  it('shim activates sp3d bevel shading — rim pixels differ from the flat fallback', async () => {
    // Guard: no ambient OffscreenCanvas in the test runtime, so the only way the
    // bevel can appear is via the shim installed by renderSlideNode(factory).
    expect(typeof g.OffscreenCanvas).toBe('undefined');

    const dataUrl = await flatPngDataUrl(8, 8, '#808080'); // mid-grey face
    const presentation = buildBeveledSlide(dataUrl);

    const flat = await renderToRgba(presentation, /* withShim */ false);
    const beveled = await renderToRgba(presentation, /* withShim */ true);

    // Shim must not leak out of renderSlideNode.
    expect(typeof g.OffscreenCanvas).toBe('undefined');

    // The picture occupies px (50,50)..(350,350) at dpr=1. Sample a horizontal
    // band a few px inside the top edge — that's where the top-bevel lip shades
    // brightest/darkest. Count how many pixels differ from the flat render.
    const { data: fd, width } = flat;
    const { data: bd } = beveled;
    const yBand = 56; // ~6 px below the top edge of the picture
    let diffCount = 0;
    let maxDelta = 0;
    for (let x = 60; x < 340; x++) {
      const i = (yBand * width + x) * 4;
      const dr = Math.abs(fd[i] - bd[i]);
      const dgc = Math.abs(fd[i + 1] - bd[i + 1]);
      const db = Math.abs(fd[i + 2] - bd[i + 2]);
      const delta = Math.max(dr, dgc, db);
      if (delta > 8) diffCount++;
      if (delta > maxDelta) maxDelta = delta;
    }

    // Without the shim the flat render fills this band with the flat grey; with
    // the shim the bevel lip lightens/darkens it. Expect a substantial run of
    // differing pixels and a clear luminance swing.
    expect(diffCount).toBeGreaterThan(40);
    expect(maxDelta).toBeGreaterThan(20);
  });

  it('without a factory the render is the flat fallback (no aux canvas, no throw)', async () => {
    const dataUrl = await flatPngDataUrl(8, 8, '#808080');
    const presentation = buildBeveledSlide(dataUrl);
    // Two flat renders must be identical (deterministic, no bevel).
    const a = await renderToRgba(presentation, false);
    const b = await renderToRgba(presentation, false);
    let diff = 0;
    for (let i = 0; i < a.data.length; i++) if (a.data[i] !== b.data[i]) diff++;
    expect(diff).toBe(0);
  });
});
