import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { getCachedBitmap, dropImageBitmapCache } from './renderer';

/** Build a minimal standard (non-placeable) WMF that draws one polyline, so the
 *  shared player produces non-empty geometry (→ a non-null bitmap). Mirrors the
 *  byte layout exercised in core's wmf.test.ts. */
function buildMinimalWmf(): Uint8Array {
  const b: number[] = [];
  const u16 = (v: number) => b.push(v & 0xff, (v >>> 8) & 0xff);
  const i16 = (v: number) => u16(v & 0xffff);
  const u32 = (v: number) => b.push(v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff);
  // 18-byte standard header (type=1, headerSize=9 words).
  u16(1); u16(9); u16(0x0300); u32(0); u16(8); u32(0); u16(0);
  // record(fn, paramWords...) — sizeWords includes the 3-word size+fn header.
  const rec = (fn: number, params: number[]) => { u32(3 + params.length); u16(fn); for (const p of params) i16(p); };
  rec(0x020b, [0, 0]);                       // SETWINDOWORG (y,x)
  rec(0x020c, [100, 100]);                   // SETWINDOWEXT (y,x)
  // CREATEPENINDIRECT: style=0, widthX=1, widthY=0, color u32 (low word, high
  // word) → 5 param words.
  rec(0x02fa, [0, 1, 0, 0, 0]);
  rec(0x012d, [0]);                          // SELECTOBJECT idx 0
  rec(0x0325, [2, 0, 0, 50, 50]);            // POLYLINE 2 pts (0,0)-(50,50)
  u32(3); u16(0x0000);                       // EOF
  return new Uint8Array(b);
}

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

  it('a WMF blip no longer throws — it rasterizes through the shared player to a bitmap', async () => {
    // Stub OffscreenCanvas (the WMF player target) + createImageBitmap, which
    // the player calls on the canvas. createImageBitmap on a *WMF blob* used to
    // throw and vanish; now getCachedBitmap routes through decodeRasterOrMetafile
    // which content-sniffs the WMF and rasterizes it instead.
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

    // A minimal standard WMF: header + window + a black pen + a polyline so the
    // player actually draws (non-empty → non-null bitmap).
    const wmf = buildMinimalWmf();
    const fetchImage = vi.fn(async (_p: string, _m: string) => new Blob([wmf as BlobPart], { type: 'image/wmf' }));

    const bmp = await getCachedBitmap('ppt/media/wmf-chart.wmf', 'image/wmf', fetchImage, 100, 100);
    expect(bmp).not.toBeNull();
    expect(bmp?.width).toBe(200); // wmfRasterTarget(100,100) → 200×200
  });

  it('a true EMF blip is skipped (null), not crashed on', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async () =>
      ({ width: 1, height: 1, close() {} }) as unknown as ImageBitmap));
    // ENHMETAHEADER: u32@0=1 (EMR_HEADER), u32@40=0x464D4520 (" EMF").
    const emf = new Uint8Array(48);
    const dv = new DataView(emf.buffer);
    dv.setUint32(0, 1, true);
    dv.setUint32(40, 0x464d4520, true);
    const fetchImage = vi.fn(async (_p: string, _m: string) => new Blob([emf as BlobPart], { type: 'image/emf' }));

    const bmp = await getCachedBitmap('ppt/media/diagram.emf', 'image/emf', fetchImage, 100, 100);
    expect(bmp).toBeNull(); // skipped gracefully — the draw site guards null
    expect(globalThis.createImageBitmap as ReturnType<typeof vi.fn>).not.toHaveBeenCalled();
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
