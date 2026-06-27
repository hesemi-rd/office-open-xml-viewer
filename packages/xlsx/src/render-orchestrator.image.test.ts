import { describe, it, expect, vi, afterEach } from 'vitest';
import { prefetchImages, decodeImageSource } from './render-orchestrator';
import type { Worksheet } from './types';

/**
 * The render orchestrator decodes embedded images lazily by zip path:
 * `decodeImageSource(imagePath, mimeType, svgImagePath?, fetchImage)` returns a
 * `CanvasImageSource` (an ImageBitmap for raster via `createImageBitmap`, an
 * HTMLImageElement for SVG via core's `getCachedSvgImageByPath`).
 * `prefetchImages` collects every image path from BOTH `ws.images` (top-level
 * `twoCellAnchor` pictures) AND the image leaves inside `ws.shapeGroups`,
 * fetches each once, and stores the decoded source in the shared cache keyed by
 * `imagePath`. Mirrors the pptx/docx renderer decode swap.
 */

// A minimal stand-in for a decoded raster bitmap.
class FakeBitmap {
  constructor(public readonly tag: string) {}
}

/** Build a minimal standard (non-placeable) WMF that draws one polyline, so the
 *  shared core player produces non-empty geometry (→ a non-null bitmap). */
function buildMinimalWmf(): Uint8Array {
  const b: number[] = [];
  const u16 = (v: number) => b.push(v & 0xff, (v >>> 8) & 0xff);
  const i16 = (v: number) => u16(v & 0xffff);
  const u32 = (v: number) => b.push(v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff);
  u16(1); u16(9); u16(0x0300); u32(0); u16(8); u32(0); u16(0); // 18-byte header
  const rec = (fn: number, params: number[]) => { u32(3 + params.length); u16(fn); for (const p of params) i16(p); };
  rec(0x020b, [0, 0]);             // SETWINDOWORG
  rec(0x020c, [100, 100]);         // SETWINDOWEXT
  rec(0x02fa, [0, 1, 0, 0, 0]);    // CREATEPENINDIRECT (color as low/high words)
  rec(0x012d, [0]);                // SELECTOBJECT
  rec(0x0325, [2, 0, 0, 50, 50]);  // POLYLINE
  u32(3); u16(0x0000);             // EOF
  return new Uint8Array(b);
}

/** Build a true EMF (ENHMETAHEADER) header so isEmf detects it. */
function buildEmfHeader(): Uint8Array {
  const buf = new Uint8Array(48);
  const dv = new DataView(buf.buffer);
  dv.setUint32(0, 1, true); // EMR_HEADER iType
  dv.setUint32(40, 0x464d4520, true); // " EMF"
  return buf;
}

/** Stub OffscreenCanvas (the WMF player's target) for the node test env. */
function stubOffscreenCanvas(): void {
  vi.stubGlobal(
    'OffscreenCanvas',
    class {
      width: number;
      height: number;
      constructor(w: number, h: number) { this.width = w; this.height = h; }
      getContext() {
        return {
          fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
          lineJoin: 'miter', lineCap: 'butt',
          save() {}, restore() {}, beginPath() {}, closePath() {},
          moveTo() {}, lineTo() {}, rect() {}, stroke() {}, fill() {},
        };
      }
    },
  );
}

/** Build a Worksheet with one top-level image and one group-leaf image, each at
 *  a distinct zip path, plus enough required fields to satisfy the type. */
function worksheetWithImages(): Worksheet {
  return {
    name: 'Sheet1',
    rows: [],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 64,
    defaultRowHeight: 20,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    charts: [],
    images: [
      {
        fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
        toCol: 2, toColOff: 0, toRow: 2, toRowOff: 0,
        nativeExtCx: 0, nativeExtCy: 0,
        imagePath: 'xl/media/image1.png',
        mimeType: 'image/png',
      },
    ],
    shapeGroups: [
      {
        fromCol: 3, fromColOff: 0, fromRow: 3, fromRowOff: 0,
        toCol: 5, toColOff: 0, toRow: 5, toRowOff: 0,
        nativeExtCx: 0, nativeExtCy: 0,
        shapes: [
          {
            x: 0, y: 0, w: 1, h: 1, rot: 0, strokeWidth: 0,
            geom: {
              type: 'image',
              imagePath: 'xl/media/image2.png',
              mimeType: 'image/png',
            },
          },
        ],
      },
    ],
  } as Worksheet;
}

describe('render-orchestrator image decode (lazy bytes)', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('prefetchImages collects BOTH ws.images and group-leaf images, keyed by imagePath, decoded once each', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async (blob: Blob) => new FakeBitmap(blob.type)));
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }),
    );
    const ws = worksheetWithImages();
    const cache = new Map<string, CanvasImageSource>();

    await prefetchImages(ws, cache, fetchImage);

    // Both paths decoded and cached under their zip path (not a data URL).
    expect(cache.has('xl/media/image1.png')).toBe(true);
    expect(cache.has('xl/media/image2.png')).toBe(true);
    expect(cache.size).toBe(2);
    // Each path fetched exactly once.
    expect(fetchImage).toHaveBeenCalledTimes(2);
    expect(fetchImage).toHaveBeenCalledWith('xl/media/image1.png', 'image/png');
    expect(fetchImage).toHaveBeenCalledWith('xl/media/image2.png', 'image/png');
  });

  it('prefetchImages skips already-cached paths (no re-fetch)', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async (blob: Blob) => new FakeBitmap(blob.type)));
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      new Blob([new TextEncoder().encode(path)], { type: mime }),
    );
    const ws = worksheetWithImages();
    const cache = new Map<string, CanvasImageSource>();
    cache.set('xl/media/image1.png', new FakeBitmap('preexisting') as unknown as CanvasImageSource);

    await prefetchImages(ws, cache, fetchImage);

    // image1 was already cached → only image2 is fetched.
    expect(fetchImage).toHaveBeenCalledTimes(1);
    expect(fetchImage).toHaveBeenCalledWith('xl/media/image2.png', 'image/png');
  });

  it('prefetchImages is a no-op when fetchImage is absent (cache stays empty)', async () => {
    const ws = worksheetWithImages();
    const cache = new Map<string, CanvasImageSource>();
    await prefetchImages(ws, cache, undefined);
    expect(cache.size).toBe(0);
  });

  it('decodeImageSource decodes raster via createImageBitmap from fetched bytes', async () => {
    const bmp = new FakeBitmap('image/png');
    const createImageBitmap = vi.fn(async () => bmp);
    vi.stubGlobal('createImageBitmap', createImageBitmap);
    const fetchImage = vi.fn(async (_path: string, mime: string) => new Blob(['X'], { type: mime }));

    const src = await decodeImageSource('xl/media/image1.png', 'image/png', undefined, fetchImage);

    expect(src).toBe(bmp);
    expect(fetchImage).toHaveBeenCalledWith('xl/media/image1.png', 'image/png');
    expect(createImageBitmap).toHaveBeenCalledTimes(1);
  });

  it('decodeImageSource forces the raster (not the SVG vector) when the picture is cropped', async () => {
    // A cropped picture (a non-null `srcRect`) with an svgBlip vector original
    // must decode the RASTER fallback: the renderer's `<a:srcRect>` crop math
    // needs the bitmap's native pixel grid, which an SVG element lacks. So even
    // with svgImagePath present, createImageBitmap (raster) is the path taken.
    const bmp = new FakeBitmap('image/png');
    const createImageBitmap = vi.fn(async () => bmp);
    vi.stubGlobal('createImageBitmap', createImageBitmap);
    const fetchImage = vi.fn(async (_p: string, mime: string) => new Blob(['X'], { type: mime }));

    const src = await decodeImageSource(
      'xl/media/image1.png',
      'image/png',
      'xl/media/image1.svg', // svgBlip present …
      fetchImage,
      0,
      0,
      { l: 0.1, t: 0, r: 0.1, b: 0 }, // … but the picture is cropped → raster wins
    );

    expect(src).toBe(bmp);
    expect(createImageBitmap).toHaveBeenCalledTimes(1);
    expect(fetchImage).toHaveBeenCalledWith('xl/media/image1.png', 'image/png');
    // The SVG part is never fetched when a crop forces the raster path.
    expect(fetchImage).not.toHaveBeenCalledWith('xl/media/image1.svg', expect.anything());
  });

  it('decodeImageSource rasterizes a WMF blip (no throw) instead of vanishing', async () => {
    // A WMF blob used to throw in createImageBitmap; now decodeImageSource routes
    // through core's decodeRasterOrMetafile, which sniffs + rasterizes it.
    stubOffscreenCanvas();
    vi.stubGlobal('createImageBitmap', vi.fn(async (s: { width: number; height: number }) =>
      ({ width: s.width, height: s.height, close() {} }) as unknown as ImageBitmap));
    const fetchImage = vi.fn(async (_p: string, _m: string) => new Blob([buildMinimalWmf() as BlobPart], { type: 'image/wmf' }));

    const src = await decodeImageSource('xl/media/chart1.wmf', 'image/wmf', undefined, fetchImage, 100, 100);

    expect(src).not.toBeNull();
    expect((src as ImageBitmap).width).toBe(200); // wmfRasterTarget(100,100) → 200×200
  });

  it('decodeImageSource returns null for an unsupported metafile (true EMF), not a throw', async () => {
    const cib = vi.fn(async () => ({ width: 1, height: 1, close() {} }) as unknown as ImageBitmap);
    vi.stubGlobal('createImageBitmap', cib);
    const fetchImage = vi.fn(async (_p: string, _m: string) => new Blob([buildEmfHeader() as BlobPart], { type: 'image/emf' }));

    const src = await decodeImageSource('xl/media/diagram.emf', 'image/emf', undefined, fetchImage, 100, 100);

    expect(src).toBeNull();
    expect(cib).not.toHaveBeenCalled(); // EMF branch never touches createImageBitmap
  });

  it('prefetchImages caches an EMF decode as null (sniffed once) — renderer skips a null source', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async () => ({ width: 1, height: 1, close() {} }) as unknown as ImageBitmap));
    const ws = worksheetWithImages();
    // Point the top-level image at an EMF; the group leaf stays a PNG.
    ws.images[0].imagePath = 'xl/media/image1.emf';
    ws.images[0].mimeType = 'image/emf';
    const fetchImage = vi.fn(async (path: string, mime: string) =>
      path.endsWith('.emf')
        ? new Blob([buildEmfHeader() as BlobPart], { type: mime })
        : new Blob([new TextEncoder().encode(path)], { type: mime }),
    );
    const cache = new Map<string, CanvasImageSource | null>();

    await prefetchImages(ws, cache, fetchImage);

    // EMF decodes to null but is CACHED as null (matching pptx's getCachedBitmap):
    // has() short-circuits the per-render prefetch so it isn't re-fetched every
    // frame, and the renderer skips a null (falsy) source.
    expect(cache.has('xl/media/image1.emf')).toBe(true);
    expect(cache.get('xl/media/image1.emf')).toBeNull();
    expect(cache.has('xl/media/image2.png')).toBe(true);
    expect(cache.size).toBe(2);

    // A second prefetch must NOT re-fetch the now-cached (null) EMF — the whole
    // point of caching the null: the unsupported blip is sniffed exactly once.
    const callsAfterFirst = fetchImage.mock.calls.length;
    await prefetchImages(ws, cache, fetchImage);
    expect(fetchImage.mock.calls.length).toBe(callsAfterFirst);
  });
});
