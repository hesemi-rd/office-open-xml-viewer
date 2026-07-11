import { describe, it, expect } from 'vitest';
import { renderSlideNode } from './render';
import type { NodeCanvasFactory } from './render';
import type { Presentation, Slide, PictureElement, ImageFill } from '@silurus/ooxml-pptx';
import { loadSkiaForTests } from './test-imports';

// End-to-end pixel check for the DrawingML `<a:duotone>` recolour (§20.1.8.23)
// on a pptx picture. A parser round-trip is NOT proof the effect draws, so this
// renders a synthetic slide whose picture carries a duotone through the real
// node render path (skia + the OffscreenCanvas / createImageBitmap shims) and
// measures the drawn pixels against the expected luminance-ramp interpolation.
const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas, loadImage } = (skia ?? {}) as Skia;

const factory: NodeCanvasFactory | null =
  Canvas && loadImage
    ? {
        createCanvas: (w, h) =>
          new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
        loadImage: (buf) =>
          loadImage(buf as Buffer) as unknown as ReturnType<NodeCanvasFactory['loadImage']>,
      }
    : null;

async function flatPngBytes(w: number, h: number, color: string): Promise<Uint8Array> {
  const c = new Canvas(w, h);
  const ctx = c.getContext('2d');
  ctx.fillStyle = color;
  ctx.fillRect(0, 0, w, h);
  return new Uint8Array(await c.toBuffer('png'));
}

const EMU_PER_PX = 9525; // 96 dpi
const PIC_IMAGE_PATH = 'ppt/media/image1.png';

/** A 300×300 px picture centred on a 400×400 px slide, carrying (or not) a
 *  `<a:duotone>` recolour. */
function buildDuotoneSlide(duotone?: { clr1: string; clr2: string }): Presentation {
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
    imagePath: PIC_IMAGE_PATH,
    mimeType: 'image/png',
    stroke: null,
    ...(duotone ? { duotone } : {}),
  };
  const slide: Slide = { index: 0, slideNumber: 1, background: null, elements: [pic] };
  return {
    slideWidth: px(400),
    slideHeight: px(400),
    slides: [slide],
    defaultTextColor: null,
    majorFont: null,
    minorFont: null,
  };
}

/** A slide whose whole-slide BACKGROUND is a stretched picture fill
 *  (`Fill::Image`, §20.1.8.14), carrying (or not) a `<a:duotone>` recolour —
 *  the latent path wired by issue #889. No elements: the background IS the ink. */
function buildDuotoneBackgroundSlide(duotone?: { clr1: string; clr2: string }): Presentation {
  const px = (n: number) => Math.round(n * EMU_PER_PX);
  const background: ImageFill = {
    fillType: 'image',
    imagePath: PIC_IMAGE_PATH,
    mimeType: 'image/png',
    ...(duotone ? { duotone } : {}),
  };
  const slide: Slide = { index: 0, slideNumber: 1, background, elements: [] };
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
  pngBytes: Uint8Array,
): Promise<{ data: Uint8ClampedArray; width: number }> {
  const width = 400;
  const canvas = new Canvas(width, 400);
  const g2 = globalThis as unknown as { createImageBitmap?: unknown };
  const prevBitmap = g2.createImageBitmap;
  g2.createImageBitmap = async (source: Blob | ArrayBuffer | Uint8Array | { toBuffer?: unknown }) => {
    // The base decode passes a Blob; applyDuotone passes the offscreen (skia)
    // canvas surface — loadImage handles a canvas by rasterizing it.
    if (source && typeof (source as { toBuffer?: unknown }).toBuffer === 'function') {
      return source; // a skia Canvas is already a valid drawImage source
    }
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
      {
        width,
        dpr: 1,
        fetchImage: async (_path: string, mime: string) =>
          new Blob([pngBytes as BlobPart], { type: mime }),
        ...(factory ? { factory } : {}),
      },
    );
  } finally {
    g2.createImageBitmap = prevBitmap as typeof globalThis.createImageBitmap;
  }
  const ctx = canvas.getContext('2d');
  const img = ctx.getImageData(0, 0, width, 400);
  return { data: img.data as unknown as Uint8ClampedArray, width };
}

/** Rec. 601 luminance (0..1) of an sRGB byte triple — the exact ramp factor the
 *  core duotone transform uses. */
function luminance601(r: number, g: number, b: number): number {
  return (0.299 * r + 0.587 * g + 0.114 * b) / 255;
}

describe.skipIf(!skia)('node duotone picture rendering (§20.1.8.23)', () => {
  it('remaps a mid-grey picture along the clr1→clr2 luminance ramp', async () => {
    // Uniform mid-grey source: luminance ≈ 0.502 → interpolates ~half-way from
    // clr1 (black) to clr2 (a light pink).
    const src = [128, 128, 128];
    const t = luminance601(src[0], src[1], src[2]);
    const clr1 = [0x00, 0x00, 0x00]; // black
    const clr2 = [0xda, 0xb6, 0xba]; // light pink
    const expected = clr1.map((d, i) => Math.round(d + (clr2[i] - d) * t));

    const pngBytes = await flatPngBytes(16, 16, '#808080');
    const plain = await renderToRgba(buildDuotoneSlide(), pngBytes);
    const duo = await renderToRgba(
      buildDuotoneSlide({ clr1: '000000', clr2: 'DAB6BA' }),
      pngBytes,
    );

    // Sample the picture centre (px 200,200 on the 400×400 slide).
    const { data: pd, width } = plain;
    const { data: dd } = duo;
    const i = (200 * width + 200) * 4;

    // Plain render shows the untouched grey.
    expect(Math.abs(pd[i] - 128)).toBeLessThan(6);
    expect(Math.abs(pd[i + 1] - 128)).toBeLessThan(6);
    expect(Math.abs(pd[i + 2] - 128)).toBeLessThan(6);

    // Duotone render shows the ramp interpolation: pink (R > G ≈ B), each channel
    // within a small tolerance of the expected duotone value.
    expect(Math.abs(dd[i] - expected[0])).toBeLessThan(10);
    expect(Math.abs(dd[i + 1] - expected[1])).toBeLessThan(10);
    expect(Math.abs(dd[i + 2] - expected[2])).toBeLessThan(10);
    // Sanity: the recolour actually happened (moved away from neutral grey) and
    // is a pink (red channel dominant), matching the clr2 endpoint hue.
    expect(dd[i]).toBeGreaterThan(dd[i + 1]);
    expect(dd[i]).toBeGreaterThan(dd[i + 2]);
    expect(Math.abs(dd[i] - 128)).toBeGreaterThan(15);
  });

  it('remaps a picture-FILL background along the luminance ramp (#889)', async () => {
    // Same ramp math as the picture case, but the duotone rides a shape/background
    // picture fill (Fill::Image) rather than a <p:pic>. This is the latent path
    // #889 wires: before the fix the background decoded through the plain cache
    // and stayed neutral grey.
    const src = [128, 128, 128];
    const t = luminance601(src[0], src[1], src[2]);
    const clr1 = [0x00, 0x00, 0x00]; // black
    const clr2 = [0xda, 0xb6, 0xba]; // light pink
    const expected = clr1.map((d, i) => Math.round(d + (clr2[i] - d) * t));

    const pngBytes = await flatPngBytes(16, 16, '#808080');
    const plain = await renderToRgba(buildDuotoneBackgroundSlide(), pngBytes);
    const duo = await renderToRgba(
      buildDuotoneBackgroundSlide({ clr1: '000000', clr2: 'DAB6BA' }),
      pngBytes,
    );

    // Sample the slide centre (px 200,200) — covered by the stretched background.
    const { data: pd, width } = plain;
    const { data: dd } = duo;
    const i = (200 * width + 200) * 4;

    // Plain background shows the untouched grey.
    expect(Math.abs(pd[i] - 128)).toBeLessThan(6);
    expect(Math.abs(pd[i + 1] - 128)).toBeLessThan(6);
    expect(Math.abs(pd[i + 2] - 128)).toBeLessThan(6);

    // Duotone background shows the ramp interpolation: a red-dominant pink within
    // tolerance of the expected duotone value.
    expect(Math.abs(dd[i] - expected[0])).toBeLessThan(10);
    expect(Math.abs(dd[i + 1] - expected[1])).toBeLessThan(10);
    expect(Math.abs(dd[i + 2] - expected[2])).toBeLessThan(10);
    expect(dd[i]).toBeGreaterThan(dd[i + 1]);
    expect(dd[i]).toBeGreaterThan(dd[i + 2]);
    expect(Math.abs(dd[i] - 128)).toBeGreaterThan(15);
  });
});
