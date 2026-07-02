import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import {
  getCachedBitmapByPath,
  peekCachedBitmapByPath,
  dropBitmapCacheByPath,
} from './bitmap-image-by-path';

/** Build a minimal standard (non-placeable) WMF that draws one polyline, so the
 *  shared player produces non-empty geometry (→ a non-null bitmap). Mirrors the
 *  byte layout exercised in wmf.test.ts. */
function buildMinimalWmf(): Uint8Array {
  const b: number[] = [];
  const u16 = (v: number) => b.push(v & 0xff, (v >>> 8) & 0xff);
  const i16 = (v: number) => u16(v & 0xffff);
  const u32 = (v: number) => b.push(v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff);
  // 18-byte standard header (type=1, headerSize=9 words).
  u16(1); u16(9); u16(0x0300); u32(0); u16(8); u32(0); u16(0);
  const rec = (fn: number, params: number[]) => { u32(3 + params.length); u16(fn); for (const p of params) i16(p); };
  rec(0x020b, [0, 0]);                       // SETWINDOWORG (y,x)
  rec(0x020c, [100, 100]);                   // SETWINDOWEXT (y,x)
  rec(0x02fa, [0, 1, 0, 0, 0]);              // CREATEPENINDIRECT
  rec(0x012d, [0]);                          // SELECTOBJECT idx 0
  rec(0x0325, [2, 0, 0, 50, 50]);            // POLYLINE 2 pts (0,0)-(50,50)
  u32(3); u16(0x0000);                       // EOF
  return new Uint8Array(b);
}

/**
 * The path-keyed decoded-bitmap cache (sibling of getCachedSvgImageByPath):
 * pulls bytes via the injected `fetchImage(path, mime)` and caches the decode by
 * zip path, namespaced per document by the `fetchImage` closure identity.
 */
describe('getCachedBitmapByPath', () => {
  beforeEach(() => {
    // `createImageBitmap` doesn't exist in the node test env; stub it to a
    // sentinel with the .close() the LRU eviction / drop calls.
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (_blob: Blob) => ({ width: 1, height: 1, close: () => {} }) as unknown as ImageBitmap),
    );
  });
  afterEach(() => vi.unstubAllGlobals());

  it('decodes by path via fetchImage and caches across draws (single fetch, single decode)', async () => {
    const fetchImage = vi.fn(async (_path: string, mime: string) => new Blob([new Uint8Array([1, 2, 3])], { type: mime }));
    // Unique path so the module-level LRU isn't pre-warmed by another test.
    const path = 'word/media/cachehit-a.png';

    const first = await getCachedBitmapByPath(path, 'image/png', fetchImage);
    const second = await getCachedBitmapByPath(path, 'image/png', fetchImage);

    expect(first).toBe(second);
    expect(fetchImage).toHaveBeenCalledTimes(1);
    expect(fetchImage).toHaveBeenCalledWith(path, 'image/png');
    expect(globalThis.createImageBitmap as ReturnType<typeof vi.fn>).toHaveBeenCalledTimes(1);
  });

  it('namespaces the cache by fetchImage — two documents sharing a zip path decode independently', async () => {
    // Different files reuse the SAME internal paths (…/media/image1.png). Opening
    // document B after document A must NOT paint A's bytes for B: the cache is
    // scoped per byte source, not by path alone.
    const path = 'word/media/image1.png';
    const fetchA = vi.fn(async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }));
    const fetchB = vi.fn(async (_p: string, mime: string) => new Blob([new Uint8Array([2])], { type: mime }));
    await getCachedBitmapByPath(path, 'image/png', fetchA);
    await getCachedBitmapByPath(path, 'image/png', fetchB);
    expect(fetchA).toHaveBeenCalledTimes(1);
    expect(fetchB).toHaveBeenCalledTimes(1);
    expect(globalThis.createImageBitmap as ReturnType<typeof vi.fn>).toHaveBeenCalledTimes(2);
    // Within one document the path still dedupes across draws.
    await getCachedBitmapByPath(path, 'image/png', fetchA);
    expect(fetchA).toHaveBeenCalledTimes(1);
  });

  it('peek returns undefined until the decode resolves, then the warmed bitmap (sync bullet contract)', async () => {
    let release!: () => void;
    const gate = new Promise<void>((r) => { release = r; });
    const fetchImage = vi.fn(async (_p: string, mime: string) => {
      await gate; // hold the fetch open so the decode hasn't settled yet
      return new Blob([new Uint8Array([7])], { type: mime });
    });
    const path = 'word/media/peek.png';
    const p = getCachedBitmapByPath(path, 'image/png', fetchImage);
    // Not warmed yet → the synchronous peek must see nothing (bullet skips).
    expect(peekCachedBitmapByPath(path, fetchImage)).toBeUndefined();
    release();
    await p;
    // After the warm pass awaited the decode, the peek sees the settled bitmap.
    const bmp = peekCachedBitmapByPath(path, fetchImage);
    expect(bmp).not.toBeUndefined();
    expect(bmp).not.toBeNull();
  });

  it('a WMF blip rasterizes through the shared player (opts.widthPt/heightPt size the raster)', async () => {
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
    vi.stubGlobal('createImageBitmap', vi.fn(async (src: { width: number; height: number }) =>
      ({ width: src.width, height: src.height, close() {} }) as unknown as ImageBitmap));

    const wmf = buildMinimalWmf();
    const fetchImage = vi.fn(async (_p: string, _m: string) => new Blob([wmf as BlobPart], { type: 'image/wmf' }));

    const bmp = await getCachedBitmapByPath('word/media/wmf.wmf', 'image/wmf', fetchImage, { widthPt: 100, heightPt: 100 });
    expect(bmp).not.toBeNull();
    expect(bmp?.width).toBe(200); // wmfRasterTarget(100,100) → 200×200
  });

  it('a true EMF blip is cached as null (skipped, not crashed)', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async () =>
      ({ width: 1, height: 1, close() {} }) as unknown as ImageBitmap));
    // ENHMETAHEADER: u32@0=1 (EMR_HEADER), u32@40=0x464D4520 (" EMF").
    const emf = new Uint8Array(48);
    const dv = new DataView(emf.buffer);
    dv.setUint32(0, 1, true);
    dv.setUint32(40, 0x464d4520, true);
    const fetchImage = vi.fn(async (_p: string, _m: string) => new Blob([emf as BlobPart], { type: 'image/emf' }));

    const bmp = await getCachedBitmapByPath('word/media/diagram.emf', 'image/emf', fetchImage, { widthPt: 100, heightPt: 100 });
    expect(bmp).toBeNull();
    expect(globalThis.createImageBitmap as ReturnType<typeof vi.fn>).not.toHaveBeenCalled();
    // The null is cached — a second draw does not re-fetch/re-sniff.
    await getCachedBitmapByPath('word/media/diagram.emf', 'image/emf', fetchImage, { widthPt: 100, heightPt: 100 });
    expect(fetchImage).toHaveBeenCalledTimes(1);
  });

  it('self-evicts on a failed decode (no poisoned cache) and retries on the next call', async () => {
    let calls = 0;
    const fetchImage = vi.fn(async (_p: string, _m: string) => {
      calls++;
      throw new Error('byte source unavailable');
    });
    const path = 'word/media/fail.png';
    await expect(getCachedBitmapByPath(path, 'image/png', fetchImage)).rejects.toThrow();
    // Second call must RETRY (cache self-evicted), not return a cached rejection.
    await expect(getCachedBitmapByPath(path, 'image/png', fetchImage)).rejects.toThrow();
    expect(calls).toBe(2);
    // The failed entry left nothing warm for the sync peek.
    expect(peekCachedBitmapByPath(path, fetchImage)).toBeUndefined();
  });

  it('evicts the LRU-oldest past the cap and closes its GPU backing', async () => {
    const closed: string[] = [];
    // Each decode returns a bitmap tagged with the blob's first byte so we can
    // see which one gets closed.
    vi.stubGlobal('createImageBitmap', vi.fn(async (blob: Blob) => {
      const tag = new Uint8Array(await blob.arrayBuffer())[0];
      return { width: 1, height: 1, close: () => closed.push(`b${tag}`) } as unknown as ImageBitmap;
    }));
    // A dedicated fetchImage → dedicated (empty) cache, so the 256 cap is reached
    // deterministically by this test alone.
    const fetchImage = vi.fn(async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }));
    // Fill to the cap (256 distinct paths), then one more to force one eviction.
    for (let i = 0; i < 256; i++) {
      await getCachedBitmapByPath(`word/media/lru-${i}.png`, 'image/png', fetchImage);
    }
    expect(closed.length).toBe(0); // nothing evicted yet at the cap
    await getCachedBitmapByPath('word/media/lru-256.png', 'image/png', fetchImage);
    // Let the eviction close-through-promise microtask run.
    await Promise.resolve();
    await Promise.resolve();
    expect(closed.length).toBe(1); // the oldest (lru-0) was closed
  });

  it('dropBitmapCacheByPath closes a document\'s bitmaps and lets it re-decode', async () => {
    const closes: number[] = [];
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (_blob: Blob) => ({ width: 1, height: 1, close: () => closes.push(1) }) as unknown as ImageBitmap),
    );
    const fetchImage = vi.fn(async (_p: string, mime: string) => new Blob([new Uint8Array([3])], { type: mime }));
    await getCachedBitmapByPath('word/media/drop-a.png', 'image/png', fetchImage);
    await getCachedBitmapByPath('word/media/drop-b.png', 'image/png', fetchImage);
    dropBitmapCacheByPath(fetchImage); // e.g. on Document.destroy()
    await Promise.resolve();
    expect(closes.length).toBe(2);
    await getCachedBitmapByPath('word/media/drop-a.png', 'image/png', fetchImage);
    expect(fetchImage).toHaveBeenCalledTimes(3); // cache cleared → fresh decode
  });
});
