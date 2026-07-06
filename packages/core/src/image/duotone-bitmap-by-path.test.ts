import { describe, it, expect, vi, afterEach } from 'vitest';
import {
  getCachedDuotoneBitmapByPath,
  duotoneCacheKey,
  dropDuotoneBitmapCache,
} from './duotone-bitmap-by-path';
import { dropBitmapCacheByPath } from './bitmap-image-by-path';
import type { OffscreenFactory } from './duotone';

/**
 * The core second-layer duotone cache: decode the base blip once (shared
 * path-keyed cache), then run the `<a:duotone>` recolour once per (path +
 * colours). Shared by the docx and pptx renderers. The recolour reads the base
 * bitmap's pixels via getImageData → transform → putImageData → a NEW bitmap, so
 * we inject an offscreen factory + stub createImageBitmap to exercise the path
 * without a real canvas.
 */
describe('getCachedDuotoneBitmapByPath', () => {
  afterEach(() => vi.unstubAllGlobals());

  /** An offscreen surface whose getImageData returns a near-white pixel buffer;
   *  putImageData records the recoloured bytes so a test can confirm the ramp ran. */
  function recordingFactory(record: { out?: Uint8ClampedArray }): OffscreenFactory {
    return ((w: number, h: number) => ({
      width: w,
      height: h,
      getContext() {
        return {
          drawImage() {},
          getImageData(_sx: number, _sy: number, sw: number, sh: number) {
            // One near-white opaque pixel (luminance ≈ 1 → maps to clr2).
            const data = new Uint8ClampedArray(sw * sh * 4).fill(255);
            return { data, width: sw, height: sh } as unknown as ImageData;
          },
          putImageData(img: ImageData) {
            record.out = img.data;
          },
        };
      },
    })) as unknown as OffscreenFactory;
  }

  it('decodes the base once and recolours once per colour pair, caching both', async () => {
    const path = 'ppt/media/duo-cachehit-a.png';
    const baseBitmap = { width: 4, height: 4, close() {} } as unknown as ImageBitmap;
    const recoloured = { width: 4, height: 4, tag: 'duo', close() {} } as unknown as ImageBitmap;
    const cib = vi.fn(async (src: unknown) => {
      // The base decode passes a Blob; applyDuotone passes the offscreen surface.
      return src instanceof Blob ? baseBitmap : recoloured;
    });
    vi.stubGlobal('createImageBitmap', cib);

    const fetchImage = vi.fn(
      async (_p: string, mime: string) => new Blob([new Uint8Array([1, 2, 3])], { type: mime }),
    );
    const duotone = { clr1: '000000', clr2: 'DAB6BA' };
    const record: { out?: Uint8ClampedArray } = {};
    const opts = { offscreenFactory: recordingFactory(record) };

    const first = await getCachedDuotoneBitmapByPath(path, 'image/png', duotone, fetchImage, opts);
    const second = await getCachedDuotoneBitmapByPath(path, 'image/png', duotone, fetchImage, opts);

    // Both calls return the SAME recoloured bitmap (memoized), the base blip was
    // fetched + decoded once, and the recolour ran once.
    expect(first).toBe(recoloured);
    expect(second).toBe(recoloured);
    expect(fetchImage).toHaveBeenCalledTimes(1);
    // createImageBitmap: once for the base decode + once for the recolour = 2.
    expect(cib).toHaveBeenCalledTimes(2);
    // The recolour mapped a near-white pixel toward clr2 (DAB6BA): R≈0xDA, not 0.
    expect(record.out?.[0]).toBeGreaterThan(200);

    dropDuotoneBitmapCache(fetchImage);
    dropBitmapCacheByPath(fetchImage);
  });

  it('passes through to the base cache (no recolour) when duotone is null', async () => {
    const path = 'ppt/media/duo-passthrough-b.png';
    const baseBitmap = { width: 2, height: 2, close() {} } as unknown as ImageBitmap;
    const cib = vi.fn(async () => baseBitmap);
    vi.stubGlobal('createImageBitmap', cib);
    const fetchImage = vi.fn(
      async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    );

    const out = await getCachedDuotoneBitmapByPath(path, 'image/png', null, fetchImage, {});

    // Returns the base bitmap directly; no second recolour decode.
    expect(out).toBe(baseBitmap);
    expect(cib).toHaveBeenCalledTimes(1);

    dropBitmapCacheByPath(fetchImage);
  });

  it('keys separate colour pairs independently (same path, two duotones)', async () => {
    const path = 'ppt/media/duo-two-c.png';
    const baseBitmap = { width: 2, height: 2, close() {} } as unknown as ImageBitmap;
    let n = 0;
    const cib = vi.fn(async (src: unknown) =>
      src instanceof Blob
        ? baseBitmap
        : ({ width: 2, height: 2, tag: `duo${n++}`, close() {} } as unknown as ImageBitmap),
    );
    vi.stubGlobal('createImageBitmap', cib);
    const fetchImage = vi.fn(
      async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    );
    const record: { out?: Uint8ClampedArray } = {};
    const opts = { offscreenFactory: recordingFactory(record) };

    const a = await getCachedDuotoneBitmapByPath(path, 'image/png', { clr1: '000000', clr2: 'FF0000' }, fetchImage, opts);
    const b = await getCachedDuotoneBitmapByPath(path, 'image/png', { clr1: '000000', clr2: '00FF00' }, fetchImage, opts);

    // Two distinct recolour results, but the base was decoded only once.
    expect(a).not.toBe(b);
    expect(fetchImage).toHaveBeenCalledTimes(1);

    dropDuotoneBitmapCache(fetchImage);
    dropBitmapCacheByPath(fetchImage);
  });

  it('duotoneCacheKey suffixes the path with both colours only when a duotone is set', () => {
    expect(duotoneCacheKey('word/media/image1.png')).toBe('word/media/image1.png');
    expect(duotoneCacheKey('word/media/image1.png', null)).toBe('word/media/image1.png');
    expect(duotoneCacheKey('word/media/image1.png', { clr1: '000000', clr2: 'DAB6BA' })).toBe(
      'word/media/image1.png|duo:000000:DAB6BA',
    );
  });
});
