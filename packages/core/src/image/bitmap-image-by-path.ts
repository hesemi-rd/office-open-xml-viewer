// Decoded-bitmap cache for raster / metafile blips shared by the docx, pptx and
// xlsx renderers, for the lazy byte-on-demand image pipeline. The sibling of
// `svg-image-by-path.ts`: same path-keyed, per-document (per-`fetchImage`)
// shape, but the drawable here is an `ImageBitmap` (GPU-backed) decoded via the
// shared `decodeRasterOrMetafile` rather than an `<img>`. Decoding an inlined
// base64 image to an ImageBitmap is expensive, and the same picture is otherwise
// re-decoded on every render (each scroll / resize / interaction / page revisit);
// this caches the decode — the Promise, so concurrent first-renders dedupe — and
// reuses it. Bounded LRU per document.
//
// The cache is keyed FIRST by the document's `fetchImage` closure, then by the
// embedded zip path. Different documents reuse the same internal paths
// (ppt/media/image1.png, word/media/image1.png, xl/media/image1.png), so a
// module-global path→bitmap map would paint document A's image for document B's
// identically-named blip when both are open on the main thread. Keying by
// `fetchImage` (one stable closure per document/deck/workbook instance) scopes
// the cache per byte source; the outer WeakMap also lets a document's whole
// bitmap cache be reclaimed with the document.
//
// GPU-lifecycle discipline (learned in PR #658): three `.catch` sites keep a
// failed or evicted decode from leaking or crashing —
//   1. the RECORD side-chain (`void promise.then(...).catch(() => {})`) that
//      copies the resolved bitmap onto the entry for the synchronous peek path
//      must swallow a decode rejection, or an empty/undecodable blob surfaces as
//      an UNHANDLED rejection (the real caller still sees it via the returned
//      promise);
//   2. the SELF-EVICT (`promise.catch(() => cache.delete(...))`) that removes a
//      transiently-failed entry so the next call retries fresh rather than
//      serving a poisoned rejection;
//   3. the EVICTION close (`oldest?.promise.then((b) => b?.close()).catch(...)`)
//      that releases the GPU backing of an LRU-evicted bitmap THROUGH its promise
//      — never by holding a raw bitmap reference — so a still-in-flight decode is
//      closed only once it resolves, and a draw already in progress is never
//      handed a closed bitmap.
// `dropBitmapCacheByPath` closes every live bitmap the same way (through the
// promise), for prompt release on the owning viewer's `destroy()`.

import { decodeRasterOrMetafile } from './wmf';

type FetchImage = (path: string, mime: string) => Promise<Blob>;

const IMAGE_BITMAP_CACHE_MAX = 256;

// Each entry pairs the in-flight/settled decode promise with its resolved bitmap.
// `bitmap` is populated once the promise resolves (see getCachedBitmapByPath),
// giving the synchronous draw sites (picture bullets, §21.1.2.4.2) a settled
// value to read via peekCachedBitmapByPath without awaiting — no separate
// parallel cache to keep in sync, so eviction/teardown only ever drop the whole
// entry.
//
// The decode can resolve to `null` for a metafile we can't rasterize (a true
// EMF, or a WMF with no drawable geometry); the null is cached (avoiding a
// re-fetch+re-sniff every frame) and the draw sites skip a null bitmap.
type BitmapCacheEntry = { promise: Promise<ImageBitmap | null>; bitmap?: ImageBitmap | null };

const bitmapCacheByFetch = new WeakMap<FetchImage, Map<string, BitmapCacheEntry>>();

function bitmapCacheFor(fetchImage: FetchImage): Map<string, BitmapCacheEntry> {
  let cache = bitmapCacheByFetch.get(fetchImage);
  if (!cache) {
    cache = new Map();
    bitmapCacheByFetch.set(fetchImage, cache);
  }
  return cache;
}

/** Options for {@link getCachedBitmapByPath}. `widthPt`/`heightPt` are the
 *  picture's intended draw size in points and only affect metafile raster
 *  sharpness (a raster blip ignores them), so the path alone keys the cache.
 *  `suppressBoundaryFrame` is the docx-only WMF window/device-boundary edge
 *  suppression (spec-clean default OFF; pptx/xlsx leave it unset). */
export interface CachedBitmapOptions {
  /** Intended draw width in points; sizes any metafile raster target. */
  widthPt?: number;
  /** Intended draw height in points; see `widthPt`. */
  heightPt?: number;
  /** Enable the docx cosmetic window/device-frame suppression heuristic. Default
   *  false = spec-clean. Only docx opts in. */
  suppressBoundaryFrame?: boolean;
}

/**
 * Decode a raster-or-metafile blip at `imagePath` to an `ImageBitmap`, cached per
 * document (keyed by `fetchImage`) then by path. The bytes are fetched lazily
 * through `fetchImage(imagePath, mimeType)` (twin of the audio/video `fetchMedia`
 * path) rather than `fetch`-ing an inlined data URL. The returned bitmap is
 * drawable with `ctx.drawImage`.
 *
 * Decoding goes through core's {@link decodeRasterOrMetafile}, which content-
 * sniffs the bytes: a WMF (which `createImageBitmap` can't decode) is rasterized
 * by the shared minimal player at a size derived from `widthPt`/`heightPt`; a
 * true EMF (or a WMF with no geometry) resolves to `null` so the draw site skips
 * the picture instead of crashing — the `null` is cached too, so the draw skips
 * it without a re-fetch+re-sniff every frame.
 *
 * LRU(256); evicted bitmaps are `.close()`d (through their promise) to release
 * their GPU backing. The returned promise rejects (and the entry self-evicts, so
 * the next call retries fresh) only on a transient fetch/decode failure —
 * callers should fall back to a raster/skip representation on rejection.
 */
export function getCachedBitmapByPath(
  imagePath: string,
  mimeType: string,
  fetchImage: FetchImage,
  opts: CachedBitmapOptions = {},
): Promise<ImageBitmap | null> {
  const { widthPt = 0, heightPt = 0, suppressBoundaryFrame = false } = opts;
  const cache = bitmapCacheFor(fetchImage);
  const existing = cache.get(imagePath);
  if (existing) {
    // Refresh LRU position.
    cache.delete(imagePath);
    cache.set(imagePath, existing);
    return existing.promise;
  }
  const promise = fetchImage(imagePath, mimeType).then((b) =>
    decodeRasterOrMetafile(b, { widthPt, heightPt, suppressBoundaryFrame }),
  );
  const entry: BitmapCacheEntry = { promise };
  // Record the resolved bitmap on the entry so the synchronous bullet draw
  // (peekCachedBitmapByPath) can read it after the warm pass awaits this promise.
  // A `null` (unsupported metafile) is recorded too, so the draw skips it. The
  // `.catch(() => {})` swallows a decode rejection on THIS side-chain (the real
  // caller still sees it via the returned `promise`): without it a failed decode
  // — e.g. an empty/undecodable blob — would surface as an unhandled rejection.
  // On failure `entry.bitmap` simply stays undefined (treated as "not ready").
  void promise
    .then((bmp) => {
      entry.bitmap = bmp;
    })
    .catch(() => {});
  // Don't poison the cache on a transient decode failure.
  promise.catch(() => cache.delete(imagePath));
  cache.set(imagePath, entry);
  if (cache.size > IMAGE_BITMAP_CACHE_MAX) {
    const oldestKey = cache.keys().next().value as string;
    const oldest = cache.get(oldestKey);
    cache.delete(oldestKey);
    oldest?.promise.then((b) => b?.close()).catch(() => {});
  }
  return promise;
}

/**
 * Synchronously return a blip's decoded bitmap if its decode has already
 * resolved (warmed by {@link getCachedBitmapByPath}), else `undefined`. Used by
 * the synchronous text-body draw to paint picture bullets (`<a:buBlip>`,
 * §21.1.2.4.2) without awaiting. A still-loading image has no `bitmap` on its
 * entry yet, so it's simply skipped.
 */
export function peekCachedBitmapByPath(
  imagePath: string,
  fetchImage: FetchImage,
): ImageBitmap | null | undefined {
  return bitmapCacheByFetch.get(fetchImage)?.get(imagePath)?.bitmap;
}

/**
 * Close every decoded bitmap for one document's `fetchImage` and forget the
 * document. Call from the owning viewer's `destroy()` so GPU-backed ImageBitmaps
 * are released promptly rather than waiting for GC. A no-op when the document
 * decoded no raster blips.
 */
export function dropBitmapCacheByPath(fetchImage: FetchImage): void {
  const cache = bitmapCacheByFetch.get(fetchImage);
  if (!cache) return;
  for (const entry of cache.values()) entry.promise.then((b) => b?.close()).catch(() => {});
  cache.clear();
  bitmapCacheByFetch.delete(fetchImage);
}
