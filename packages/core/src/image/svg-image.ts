// Shared decoder for embedded SVG images (Microsoft's `asvg:svgBlip` extension,
// MS-ODRAWXML) used by the docx, pptx and xlsx renderers. The parsers surface
// the vector original as a `data:image/svg+xml;base64,…` URL on the element's
// `svgDataUrl`; this decodes it to a drawable `HTMLImageElement`.
//
// SVG is decoded via an `<img>` element rather than `createImageBitmap`, because
// `createImageBitmap` cannot rasterize SVG in every browser. The same picture is
// re-drawn on every render (scroll / resize / interaction), so the decoded image
// is cached — the Promise, so concurrent first-renders dedupe. Unlike an
// ImageBitmap, an HTMLImageElement holds no GPU resource that must be released,
// so eviction just drops the entry (nothing to `.close()`). Bounded FIFO.

const SVG_IMAGE_CACHE_MAX = 256;
const svgImageCache = new Map<string, Promise<HTMLImageElement>>();

/**
 * Decode a `data:image/svg+xml;base64,…` URL to an `HTMLImageElement`, cached by
 * URL. The returned image is drawable with `ctx.drawImage` exactly like an
 * ImageBitmap; both expose numeric `.width`/`.height`. Rejects if the SVG fails
 * to load — callers should fall back to the raster `dataUrl` on rejection.
 */
export function getCachedSvgImage(dataUrl: string): Promise<HTMLImageElement> {
  const existing = svgImageCache.get(dataUrl);
  if (existing) {
    // Refresh LRU position.
    svgImageCache.delete(dataUrl);
    svgImageCache.set(dataUrl, existing);
    return existing;
  }
  const p = new Promise<HTMLImageElement>((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      // `decode()` guarantees the bitmap is ready before the first draw, so the
      // synchronous paint pass never draws an undecoded image. Fall back to the
      // already-fired load event if decode() is unavailable / rejects.
      if (typeof img.decode === 'function') {
        img
          .decode()
          .then(() => resolve(img))
          .catch(() => resolve(img));
      } else {
        resolve(img);
      }
    };
    img.onerror = () => reject(new Error('SVG image failed to load'));
    img.src = dataUrl;
  });
  // Don't poison the cache on a transient decode failure.
  p.catch(() => svgImageCache.delete(dataUrl));
  svgImageCache.set(dataUrl, p);
  if (svgImageCache.size > SVG_IMAGE_CACHE_MAX) {
    // HTMLImageElement holds no GPU handle — just drop the oldest entry.
    const oldestKey = svgImageCache.keys().next().value as string;
    svgImageCache.delete(oldestKey);
  }
  return p;
}
