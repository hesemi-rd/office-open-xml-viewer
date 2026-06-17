import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { decodeRaster, preloadImages } from './renderer';
import type { DocxDocumentModel } from './types';

/**
 * docx raster blips decode through `fetchImage(path, mime)` (twin of pptx's
 * lazy-bytes path) instead of `fetch`-ing an inlined data URL. `preloadImages`
 * keys the decoded-image map by `imageKey(imagePath, colorReplaceFrom)` and must
 * decode each distinct key exactly once. SVG vector-優先 + color-replacement
 * behavior is unchanged; this test pins the raster + keying contract.
 */
describe('docx lazy image bytes', () => {
  beforeEach(() => {
    // `createImageBitmap` doesn't exist in the node test env; stub it to a
    // sentinel image bitmap.
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (_src: unknown) => ({ width: 2, height: 2, close: () => {} }) as unknown as ImageBitmap),
    );
  });
  afterEach(() => vi.unstubAllGlobals());

  it('decodeRaster pulls bytes by path via fetchImage and passes the MIME through', async () => {
    const fetchImage = vi.fn(
      async (_path: string, mime: string) => new Blob([new Uint8Array([1, 2, 3])], { type: mime }),
    );
    const bmp = await decodeRaster('word/media/image1.png', 'image/png', undefined, fetchImage);
    expect(bmp).toBeTruthy();
    expect(fetchImage).toHaveBeenCalledTimes(1);
    expect(fetchImage).toHaveBeenCalledWith('word/media/image1.png', 'image/png');
    expect((globalThis.createImageBitmap as ReturnType<typeof vi.fn>)).toHaveBeenCalledTimes(1);
  });

  it('preloadImages keys by imagePath and decodes each distinct key exactly once', async () => {
    const fetchImage = vi.fn(
      async (path: string, mime: string) => new Blob([new Uint8Array([path.length])], { type: mime }),
    );
    const imgRun = (imagePath: string) => ({
      type: 'image',
      imagePath,
      mimeType: 'image/png',
      widthPt: 10,
      heightPt: 10,
    });
    // image1 is referenced twice (must collapse to ONE decode); image2 once.
    const doc = {
      body: [
        { type: 'paragraph', runs: [imgRun('word/media/image1.png')] },
        { type: 'paragraph', runs: [imgRun('word/media/image1.png')] }, // dup → same key
        { type: 'paragraph', runs: [imgRun('word/media/image2.png')] }, // distinct key
      ],
      headers: {},
      footers: {},
    } as unknown as DocxDocumentModel;

    const map = await preloadImages(doc, fetchImage);

    // Two distinct keys (the raster path itself, no colorReplaceFrom suffix).
    expect(map.has('word/media/image1.png')).toBe(true);
    expect(map.has('word/media/image2.png')).toBe(true);
    expect(map.size).toBe(2);
    // The duplicate reference must NOT trigger a second fetch/decode for its key:
    // one fetch per distinct path = 2 total (not 3).
    expect(fetchImage).toHaveBeenCalledTimes(2);
    expect((globalThis.createImageBitmap as ReturnType<typeof vi.fn>)).toHaveBeenCalledTimes(2);
  });

  it('imageKey appends a colorReplaceFrom suffix so a recoloured ref is a distinct key', () => {
    // The colorReplaceFrom variant produces a distinct cache key even for the
    // same path. (The decode of that variant exercises canvas APIs unavailable
    // in the node test env, so the keying — not the recolour pixels — is pinned
    // here; the recolour pass itself is unchanged from before this refactor.)
    const fetchImage = vi.fn(
      async (_path: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    );
    const plain = { type: 'image', imagePath: 'word/media/image1.png', mimeType: 'image/png', widthPt: 1, heightPt: 1 };
    const recoloured = { ...plain, colorReplaceFrom: 'FFFFFF' };
    const doc = {
      body: [
        { type: 'paragraph', runs: [plain] },
        { type: 'paragraph', runs: [recoloured] },
      ],
      headers: {},
      footers: {},
    } as unknown as DocxDocumentModel;
    // Both refs are collected as DISTINCT pairs → two fetch attempts by key.
    return preloadImages(doc, fetchImage).then(() => {
      expect(fetchImage).toHaveBeenCalledTimes(2);
      // The plain key always decodes; the recolour key shares the same fetch
      // path but a distinct cache key (proven by the 2 fetches above).
      expect(fetchImage).toHaveBeenNthCalledWith(1, 'word/media/image1.png', 'image/png');
      expect(fetchImage).toHaveBeenNthCalledWith(2, 'word/media/image1.png', 'image/png');
    });
  });

  it('preloadImages with no fetchImage yields an empty map (no byte source)', async () => {
    const doc = {
      body: [
        {
          type: 'paragraph',
          runs: [{ type: 'image', imagePath: 'word/media/image1.png', mimeType: 'image/png', widthPt: 10, heightPt: 10 }],
        },
      ],
      headers: {},
      footers: {},
    } as unknown as DocxDocumentModel;
    const map = await preloadImages(doc, undefined);
    expect(map.size).toBe(0);
  });
});
