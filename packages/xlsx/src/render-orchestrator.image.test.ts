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
});
