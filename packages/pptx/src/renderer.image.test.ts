import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { getCachedBitmap, dropImageBitmapCache } from './renderer';

/**
 * Raster blips decode through `getCachedBitmap`, which now keys its LRU by zip
 * path and pulls bytes via the injected `fetchImage(path, mime)` (twin of the
 * audio/video `fetchMedia` path) instead of `fetch`-ing an inlined data URL.
 * Two draws of the same path must share one fetch + one decode.
 */
describe('getCachedBitmap (lazy image bytes)', () => {
  beforeEach(() => {
    // `createImageBitmap` doesn't exist in the node test env; stub it to a
    // sentinel with the .close() the LRU eviction calls.
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (_blob: Blob) => ({ width: 1, height: 1, close: () => {} }) as unknown as ImageBitmap),
    );
  });
  afterEach(() => vi.unstubAllGlobals());

  it('decodes by path via fetchImage and caches across draws (single fetch)', async () => {
    const fetchImage = vi.fn(async (_path: string, mime: string) => new Blob([new Uint8Array([1, 2, 3])], { type: mime }));
    // Unique path so the module-level LRU isn't pre-warmed by another test.
    const path = 'ppt/media/getcachedbitmap-a.png';

    const first = await getCachedBitmap(path, 'image/png', fetchImage);
    const second = await getCachedBitmap(path, 'image/png', fetchImage);

    expect(first).toBe(second); // same cached promise result
    expect(fetchImage).toHaveBeenCalledTimes(1);
    expect(fetchImage).toHaveBeenCalledWith(path, 'image/png');
    expect((globalThis.createImageBitmap as ReturnType<typeof vi.fn>)).toHaveBeenCalledTimes(1);
  });

  it('passes the blip MIME through to fetchImage', async () => {
    const fetchImage = vi.fn(async (_path: string, mime: string) => new Blob([new Uint8Array([9])], { type: mime }));
    const path = 'ppt/media/getcachedbitmap-b.jpeg';
    await getCachedBitmap(path, 'image/jpeg', fetchImage);
    expect(fetchImage).toHaveBeenCalledWith(path, 'image/jpeg');
  });

  it('namespaces the cache by fetchImage — two decks sharing a zip path decode independently', async () => {
    // Different .pptx files reuse the SAME internal paths (ppt/media/image1.png).
    // Opening deck B after deck A must NOT paint deck A's bytes for deck B: the
    // decoded-bitmap cache has to be scoped per byte source, not by path alone.
    const path = 'ppt/media/image1.png';
    const fetchA = vi.fn(async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }));
    const fetchB = vi.fn(async (_p: string, mime: string) => new Blob([new Uint8Array([2])], { type: mime }));
    await getCachedBitmap(path, 'image/png', fetchA);
    await getCachedBitmap(path, 'image/png', fetchB);
    expect(fetchA).toHaveBeenCalledTimes(1);
    expect(fetchB).toHaveBeenCalledTimes(1); // deck B must consult its own source
    expect((globalThis.createImageBitmap as ReturnType<typeof vi.fn>)).toHaveBeenCalledTimes(2);
    // Within one deck the path still dedupes across draws.
    await getCachedBitmap(path, 'image/png', fetchA);
    expect(fetchA).toHaveBeenCalledTimes(1);
  });

  it('dropImageBitmapCache closes a deck\'s bitmaps and lets it re-decode', async () => {
    const closes: number[] = [];
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (_blob: Blob) => ({ width: 1, height: 1, close: () => closes.push(1) }) as unknown as ImageBitmap),
    );
    const fetchImage = vi.fn(async (_p: string, mime: string) => new Blob([new Uint8Array([3])], { type: mime }));
    await getCachedBitmap('ppt/media/drop-a.png', 'image/png', fetchImage);
    await getCachedBitmap('ppt/media/drop-b.png', 'image/png', fetchImage);
    dropImageBitmapCache(fetchImage); // e.g. on Presentation.destroy()
    // Both decoded bitmaps were closed to release their GPU backing.
    await Promise.resolve();
    expect(closes.length).toBe(2);
    await getCachedBitmap('ppt/media/drop-a.png', 'image/png', fetchImage);
    expect(fetchImage).toHaveBeenCalledTimes(3); // cache cleared → fresh decode
  });
});
