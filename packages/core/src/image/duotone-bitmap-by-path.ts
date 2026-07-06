// Second-layer cache for the DrawingML `<a:duotone>` recolour (ECMA-376
// §20.1.8.23) of a raster/metafile blip, shared by the docx and pptx renderers
// (xlsx keeps its own worksheet-scoped `loadedImages` map with the same keying).
//
// The base, colour-FREE bitmap comes from the path-keyed
// `getCachedBitmapByPath` — shared across every reference to a path and
// reclaimed with the document. The duotone pixel pass (getImageData →
// luminance remap → putImageData → ImageBitmap, expensive) then runs ONCE per
// (imagePath, clr1, clr2) triple and its ImageBitmap is memoized here, so
// revisiting a page/slide re-runs NEITHER the decode NOR the recolour. This is
// the exact two-layer shape docx already used for its `a:clrChange`
// (colorReplaced) cache, lifted to core so pptx reuses it verbatim.
//
// Keyed FIRST by the document's `fetchImage` closure (one stable identity per
// document/deck), then by `duotoneCacheKey(imagePath, duotone)` — mirroring the
// core base cache's per-document namespacing so two documents sharing a zip path
// + duotone don't cross-contaminate, and the whole map is reclaimed with the
// document. The stored value is an ImageBitmap (a fresh OffscreenCanvas raster),
// so on destroy it must be closed (see `dropDuotoneBitmapCache`), the same
// GPU-lifecycle discipline the base cache follows through its promise.

import { getCachedBitmapByPath, type CachedBitmapOptions } from './bitmap-image-by-path';
import { applyDuotone, type Duotone, type OffscreenFactory } from './duotone';
import { imageNaturalSize } from './crop';

type FetchImage = (path: string, mime: string) => Promise<Blob>;

/** Cache key for a decoded bitmap that may carry a `<a:duotone>` recolour. A
 *  plain picture is keyed by its zip `imagePath`; a duotone picture is keyed by
 *  the path PLUS both resolved endpoint colours, so the recoloured bitmap is
 *  cached and looked up separately from the raw blip. Callers compute this both
 *  when warming the cache and when drawing, so the two agree without sharing a
 *  cache reference. Mirrors xlsx's `imageCacheKey` and docx's former
 *  `imageKey(path, colorReplaceFrom)`. */
export function duotoneCacheKey(imagePath: string, duotone?: Duotone | null): string {
  return duotone ? `${imagePath}|duo:${duotone.clr1}:${duotone.clr2}` : imagePath;
}

const duotoneByFetch = new WeakMap<FetchImage, Map<string, Promise<ImageBitmap | null>>>();

function duotoneCacheFor(fetchImage: FetchImage): Map<string, Promise<ImageBitmap | null>> {
  let cache = duotoneByFetch.get(fetchImage);
  if (!cache) {
    cache = new Map();
    duotoneByFetch.set(fetchImage, cache);
  }
  return cache;
}

/**
 * Decode a raster/metafile blip at `imagePath` and, when `duotone` is set,
 * recolour it along the `clr1`→`clr2` luminance ramp (§20.1.8.23), returning a
 * drawable source cached per document then by `duotoneCacheKey`.
 *
 * With NO duotone this is a thin pass-through to {@link getCachedBitmapByPath}
 * (the shared base cache) — no second-layer entry is created, so a non-duotone
 * picture behaves byte-for-byte as before. With a duotone the base bitmap is
 * decoded (and cached) once, then the recolour runs once per colour pair and is
 * memoized here.
 *
 * The recolour needs a readable pixel grid, so it only applies to a decoded
 * raster/metafile bitmap; a `null` base (an unsupported metafile — true EMF /
 * geometry-less WMF) propagates as `null` and the draw site skips it. When the
 * offscreen pixel pipeline is unavailable, {@link applyDuotone} returns the base
 * unchanged, so the picture still draws (just without the recolour).
 *
 * @param opts.offscreenFactory optional surface factory for environments without
 *   a global `OffscreenCanvas` (node); forwarded to {@link applyDuotone}.
 */
export async function getCachedDuotoneBitmapByPath(
  imagePath: string,
  mimeType: string,
  duotone: Duotone | null | undefined,
  fetchImage: FetchImage,
  opts: CachedBitmapOptions & { offscreenFactory?: OffscreenFactory } = {},
): Promise<ImageBitmap | null> {
  const { offscreenFactory, ...bitmapOpts } = opts;
  // Base, colour-free bitmap from the shared path-keyed cache.
  const base = await getCachedBitmapByPath(imagePath, mimeType, fetchImage, bitmapOpts);
  // No duotone → return the base directly (no second-layer entry). A `null`
  // (unsupported metafile) propagates unchanged.
  if (!duotone || !base) return base;
  const cache = duotoneCacheFor(fetchImage);
  const key = duotoneCacheKey(imagePath, duotone);
  let hit = cache.get(key);
  if (!hit) {
    // The recolour reads the SHARED base bitmap and produces a fresh independent
    // raster, so the base is never mutated and stays reusable for other draws.
    hit = (async () => {
      const { w, h } = imageNaturalSize(base);
      if (w <= 0 || h <= 0) return base;
      const recoloured = await applyDuotone(base, duotone, { width: w, height: h, offscreenFactory });
      // `applyDuotone` returns a CanvasImageSource; when the pixel pipeline ran
      // it is a fresh ImageBitmap, otherwise it is the (unchanged) base bitmap.
      return recoloured as ImageBitmap;
    })();
    // Don't poison the cache if the recolour pass rejects; let the next call retry.
    hit.catch(() => cache.delete(key));
    cache.set(key, hit);
  }
  return hit;
}

/**
 * Close every duotone-recoloured ImageBitmap for one document's `fetchImage` and
 * forget the document. Call from the owning viewer's `destroy()` alongside
 * `dropBitmapCacheByPath` (base bitmaps) so both caches release their GPU
 * backing promptly. A no-op when the document decoded no duotone picture.
 *
 * Almost every stored value is a fresh recolour raster owned ONLY by this cache,
 * so it must be closed here. The rare exception is a pass-through (the pixel
 * pipeline was unavailable, so `applyDuotone` returned the unchanged base
 * bitmap): that bitmap is also owned by the base cache and closed by
 * `dropBitmapCacheByPath`, but `ImageBitmap.close()` is idempotent, so closing
 * it a second time here is harmless. Closing through the promise (never a raw
 * reference) means a still-in-flight recolour is closed only once it resolves.
 */
export function dropDuotoneBitmapCache(fetchImage: FetchImage): void {
  const cache = duotoneByFetch.get(fetchImage);
  if (!cache) return;
  for (const p of cache.values()) p.then((b) => b?.close()).catch(() => {});
  cache.clear();
  duotoneByFetch.delete(fetchImage);
}
