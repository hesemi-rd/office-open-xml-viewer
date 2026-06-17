import { describe, it, expect } from 'vitest';
import { crc32 } from 'node:zlib';
import {
  makeSourceBufferFetchImage,
  renderSlideNode,
  installImageBitmapShim,
} from './render';
import type { NodeCanvasFactory } from './render';
import type { Presentation, Slide, PictureElement } from '@silurus/ooxml-pptx';

// skia-canvas ships a native binding CI omits; load dynamically so the
// canvas-backed block skips cleanly when absent (mirrors render.test.ts).
const skia = await import('skia-canvas').catch(() => null);
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

/**
 * Build a minimal, valid ZIP archive in memory holding a single STORED
 * (uncompressed) entry. The Rust `zip` crate that backs `extract_image` reads
 * this happily, so we avoid pulling in a zip dependency just for the test.
 *
 * Layout per the ZIP spec (PKZIP APPNOTE): one local file header + data, then a
 * central directory record, then the end-of-central-directory record.
 */
function makeZipWithEntry(name: string, data: Uint8Array): Uint8Array {
  const enc = new TextEncoder();
  const nameBytes = enc.encode(name);
  const crc = crc32(Buffer.from(data)) >>> 0;
  const size = data.length;

  // Local file header (30 bytes fixed + name).
  const local = new Uint8Array(30 + nameBytes.length + size);
  const lv = new DataView(local.buffer);
  lv.setUint32(0, 0x04034b50, true); // local file header signature
  lv.setUint16(4, 20, true); // version needed
  lv.setUint16(6, 0, true); // flags
  lv.setUint16(8, 0, true); // compression method: 0 = stored
  lv.setUint16(10, 0, true); // mod time
  lv.setUint16(12, 0, true); // mod date
  lv.setUint32(14, crc, true); // crc-32
  lv.setUint32(18, size, true); // compressed size
  lv.setUint32(22, size, true); // uncompressed size
  lv.setUint16(26, nameBytes.length, true); // file name length
  lv.setUint16(28, 0, true); // extra field length
  local.set(nameBytes, 30);
  local.set(data, 30 + nameBytes.length);

  // Central directory header (46 bytes fixed + name).
  const central = new Uint8Array(46 + nameBytes.length);
  const cv = new DataView(central.buffer);
  cv.setUint32(0, 0x02014b50, true); // central file header signature
  cv.setUint16(4, 20, true); // version made by
  cv.setUint16(6, 20, true); // version needed
  cv.setUint16(8, 0, true); // flags
  cv.setUint16(10, 0, true); // compression method
  cv.setUint16(12, 0, true); // mod time
  cv.setUint16(14, 0, true); // mod date
  cv.setUint32(16, crc, true); // crc-32
  cv.setUint32(20, size, true); // compressed size
  cv.setUint32(24, size, true); // uncompressed size
  cv.setUint16(28, nameBytes.length, true); // file name length
  cv.setUint16(30, 0, true); // extra field length
  cv.setUint16(32, 0, true); // comment length
  cv.setUint16(34, 0, true); // disk number start
  cv.setUint16(36, 0, true); // internal attrs
  cv.setUint32(38, 0, true); // external attrs
  cv.setUint32(42, 0, true); // local header offset
  central.set(nameBytes, 46);

  // End of central directory record (22 bytes).
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true); // EOCD signature
  ev.setUint16(4, 0, true); // disk number
  ev.setUint16(6, 0, true); // disk with central dir
  ev.setUint16(8, 1, true); // entries on this disk
  ev.setUint16(10, 1, true); // total entries
  ev.setUint32(12, central.length, true); // central dir size
  ev.setUint32(16, local.length, true); // central dir offset
  ev.setUint16(20, 0, true); // comment length

  const out = new Uint8Array(local.length + central.length + eocd.length);
  out.set(local, 0);
  out.set(central, local.length);
  out.set(eocd, local.length + central.length);
  return out;
}

describe('makeSourceBufferFetchImage', () => {
  it('returns the real entry bytes from the source buffer (not an empty Blob)', async () => {
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4]);
    const zip = makeZipWithEntry('ppt/media/image1.png', png);

    const fetchImage = makeSourceBufferFetchImage(zip);
    const blob = await fetchImage('ppt/media/image1.png', 'image/png');

    expect(blob.type).toBe('image/png');
    expect(blob.size).toBe(png.length);
    const got = new Uint8Array(await blob.arrayBuffer());
    expect(Array.from(got)).toEqual(Array.from(png));
  });

  it('accepts an ArrayBuffer source buffer', async () => {
    const bytes = new Uint8Array([10, 20, 30]);
    const zip = makeZipWithEntry('word/media/i.bin', bytes);
    // Copy into a standalone ArrayBuffer to confirm both Uint8Array and
    // ArrayBuffer sources work.
    const ab = new ArrayBuffer(zip.byteLength);
    new Uint8Array(ab).set(zip);

    const fetchImage = makeSourceBufferFetchImage(ab);
    const blob = await fetchImage('word/media/i.bin', 'application/octet-stream');

    expect(blob.size).toBe(bytes.length);
    expect(Array.from(new Uint8Array(await blob.arrayBuffer()))).toEqual([10, 20, 30]);
  });
});

const EMU_PER_PX = 9525; // 96 dpi
const PIC_IMAGE_PATH = 'ppt/media/image1.png';

function buildPictureSlide(): Presentation {
  const px = (n: number) => Math.round(n * EMU_PER_PX);
  const pic: PictureElement = {
    type: 'picture',
    x: px(50),
    y: px(50),
    width: px(100),
    height: px(100),
    rotation: 0,
    flipH: false,
    flipV: false,
    imagePath: PIC_IMAGE_PATH,
    mimeType: 'image/png',
    stroke: null,
  };
  const slide: Slide = { index: 0, slideNumber: 1, background: null, elements: [pic] };
  return {
    slideWidth: px(200),
    slideHeight: px(200),
    slides: [slide],
    defaultTextColor: null,
    majorFont: null,
    minorFont: null,
  };
}

// Full-path coverage: `renderSlideNode` given only `sourceBuffer` (no explicit
// fetchImage) must read the picture's bytes out of the archive via WASM
// `extract_image` and paint them — proving the source-buffer wiring end to end.
describe.skipIf(!skia)('renderSlideNode sourceBuffer wiring', () => {
  it('paints an embedded picture by extracting its bytes from sourceBuffer', async () => {
    // A solid-red 8x8 PNG produced by skia, stored into a one-entry zip that
    // plays the role of the source .pptx archive.
    const c = new Canvas(8, 8);
    const cctx = c.getContext('2d');
    cctx.fillStyle = '#ff0000';
    cctx.fillRect(0, 0, 8, 8);
    const png = new Uint8Array(await c.toBuffer('png'));
    const sourceBuffer = makeZipWithEntry(PIC_IMAGE_PATH, png);

    const width = 200;
    const dpr = 1;
    const canvas = new Canvas(width * dpr, 200 * dpr);
    const restoreBitmap = installImageBitmapShim(factory as NodeCanvasFactory);
    try {
      await renderSlideNode(
        canvas as unknown as Parameters<typeof renderSlideNode>[0],
        buildPictureSlide(),
        0,
        { width, dpr, sourceBuffer }, // no explicit fetchImage — drive it from the buffer
      );
    } finally {
      restoreBitmap();
    }

    // The picture occupies px (50,50)..(150,150). Sample its centre: it must be
    // the red we encoded, not transparent (which is what an empty-Blob fetch
    // would have produced).
    const ctx = canvas.getContext('2d');
    const px = ctx.getImageData(100, 100, 1, 1).data;
    expect(px[0]).toBeGreaterThan(200); // R
    expect(px[1]).toBeLessThan(60); // G
    expect(px[2]).toBeLessThan(60); // B
    expect(px[3]).toBeGreaterThan(200); // A (opaque — image painted)
  });

  it('without sourceBuffer or fetchImage the picture draws nothing (transparent)', async () => {
    const width = 200;
    const dpr = 1;
    const canvas = new Canvas(width * dpr, 200 * dpr);
    const restoreBitmap = installImageBitmapShim(factory as NodeCanvasFactory);
    try {
      await renderSlideNode(
        canvas as unknown as Parameters<typeof renderSlideNode>[0],
        buildPictureSlide(),
        0,
        { width, dpr }, // additive: neither sourceBuffer nor fetchImage
      );
    } finally {
      restoreBitmap();
    }
    const ctx = canvas.getContext('2d');
    const px = ctx.getImageData(100, 100, 1, 1).data;
    expect(px[3]).toBe(0); // fully transparent — prior empty-Blob behavior preserved
  });
});
