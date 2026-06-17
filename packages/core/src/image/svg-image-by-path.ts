// Decoder for embedded SVG images (Microsoft's `asvg:svgBlip` extension,
// MS-ODRAWXML) used by the docx, pptx and xlsx renderers, for the lazy,
// byte-on-demand image pipeline. The bytes are pulled lazily via a caller-
// supplied `fetchImage(path, mimeType)` — mirroring the pptx audio/video
// extraction pattern. The fetched bytes are wrapped in an object URL that this
// module owns and revokes on drop (unlike a `data:` URL, an object URL is a
// live handle that leaks if never released).
//
// The cache is keyed FIRST by the document's `fetchImage` closure, then by the
// embedded zip path. Two different documents reuse the same internal paths
// (e.g. both have "word/media/image1.svg"), so the path alone is NOT a unique
// key — the byte source is. Keying by `fetchImage` (one stable closure per
// document instance) keeps document A's decoded SVG from being served for
// document B's identically-named blip. The outer WeakMap also lets a document's
// whole image cache be reclaimed with the document.
//
// SVG is decoded via an `<img>` element rather than `createImageBitmap`,
// because `createImageBitmap` cannot rasterize SVG in every browser. The Promise
// is cached so concurrent first-renders dedupe; an HTMLImageElement holds no GPU
// resource, so a drop only needs to forget the entry and revoke its object URL.
// Bounded LRU per document.

type FetchImage = (path: string, mimeType: string) => Promise<Blob>;

/** Per-document decode state: the LRU of decoded `<img>` promises and the
 *  object URL owned for each path (kept in lockstep so drop revokes both). */
interface DocCache {
  imgs: Map<string, Promise<HTMLImageElement>>;
  urls: Map<string, string>;
}

const byFetch = new WeakMap<FetchImage, DocCache>();
const MAX = 256;

function docCacheFor(fetchImage: FetchImage): DocCache {
  let dc = byFetch.get(fetchImage);
  if (!dc) {
    dc = { imgs: new Map(), urls: new Map() };
    byFetch.set(fetchImage, dc);
  }
  return dc;
}

// Drop a path within one document: forget the cached promise and revoke its
// object URL together, so the URL lifecycle has a single owner whether the drop
// is by LRU eviction or by decode failure. Co-locating forget+revoke is what
// keeps a failed decode from leaking its handle.
function dropEntry(dc: DocCache, path: string): void {
  dc.imgs.delete(path);
  const url = dc.urls.get(path);
  if (url) {
    URL.revokeObjectURL(url);
    dc.urls.delete(path);
  }
}

/**
 * Decode the SVG at `svgImagePath` to an `HTMLImageElement`, cached per
 * document (keyed by `fetchImage`) then by path. The bytes are fetched lazily
 * through `fetchImage(path, mimeType)`; the resulting object URL is owned by
 * this module and revoked when the entry is dropped. The returned image is
 * drawable with `ctx.drawImage` exactly like an ImageBitmap. Rejects (and
 * self-evicts, so the next call retries fresh) if the SVG fails to load —
 * callers should fall back to a raster representation on rejection.
 */
export function getCachedSvgImageByPath(
  svgImagePath: string,
  fetchImage: FetchImage,
): Promise<HTMLImageElement> {
  const dc = docCacheFor(fetchImage);
  const hit = dc.imgs.get(svgImagePath);
  if (hit) {
    // Refresh LRU position.
    dc.imgs.delete(svgImagePath);
    dc.imgs.set(svgImagePath, hit);
    return hit;
  }
  const p = (async () => {
    const blob = await fetchImage(svgImagePath, 'image/svg+xml');
    const url = URL.createObjectURL(blob);
    dc.urls.set(svgImagePath, url);
    const img = new Image();
    await new Promise<void>((resolve, reject) => {
      img.onload = () => {
        // decode() guarantees the bitmap is ready before the synchronous paint
        // pass, so a draw never hits an undecoded image. Fall back to the load
        // event if decode() is unavailable / rejects.
        if (typeof img.decode === 'function') {
          img
            .decode()
            .then(() => resolve())
            .catch(() => resolve());
        } else {
          resolve();
        }
      };
      img.onerror = () => reject(new Error(`svg load failed: ${svgImagePath}`));
      img.src = url;
    });
    return img;
  })();
  // Don't poison the cache on a transient fetch/decode failure; revoke the URL too.
  p.catch(() => dropEntry(dc, svgImagePath));
  dc.imgs.set(svgImagePath, p);
  if (dc.imgs.size > MAX) {
    const oldest = dc.imgs.keys().next().value as string;
    dropEntry(dc, oldest);
  }
  return p;
}

/**
 * Release every decoded SVG and its object URL for one document's `fetchImage`,
 * then forget the document. Call from the owning viewer's `destroy()` so the
 * live object URLs are revoked promptly rather than waiting for GC. A no-op when
 * the document never decoded an SVG.
 */
export function dropSvgImageCache(fetchImage: FetchImage): void {
  const dc = byFetch.get(fetchImage);
  if (!dc) return;
  for (const url of dc.urls.values()) URL.revokeObjectURL(url);
  dc.urls.clear();
  dc.imgs.clear();
  byFetch.delete(fetchImage);
}
