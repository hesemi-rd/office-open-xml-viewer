// Decoder for embedded SVG images (Microsoft's `asvg:svgBlip` extension,
// MS-ODRAWXML) used by the docx, pptx and xlsx renderers, for the lazy,
// byte-on-demand image pipeline. The cache keys on the embedded zip path
// (e.g. "word/media/image1.svg") and pulls the bytes lazily via a caller-
// supplied `fetchImage(path, mimeType)` — mirroring the pptx audio/video
// extraction pattern. The fetched bytes are wrapped in an object URL that this
// module owns and revokes on drop (unlike a `data:` URL, an object URL is a
// live handle that leaks if never released).
//
// SVG is decoded via an `<img>` element rather than `createImageBitmap`,
// because `createImageBitmap` cannot rasterize SVG in every browser. The Promise
// is cached so concurrent first-renders dedupe; an HTMLImageElement holds no GPU
// resource, so a drop only needs to forget the entry and revoke its object URL.
// Bounded LRU.

const svgByPathCache = new Map<string, Promise<HTMLImageElement>>();
const urlByPath = new Map<string, string>(); // object URL owned per path
const MAX = 256;

// Drop a path: forget the cached promise and revoke its object URL together, so
// the URL lifecycle has a single owner whether the drop is by LRU eviction or by
// decode failure. Co-locating forget+revoke is what keeps a failed decode from
// leaking its handle.
function dropEntry(path: string): void {
  svgByPathCache.delete(path);
  const url = urlByPath.get(path);
  if (url) {
    URL.revokeObjectURL(url);
    urlByPath.delete(path);
  }
}

/**
 * Decode the SVG at `svgImagePath` to an `HTMLImageElement`, cached by path.
 * The bytes are fetched lazily through `fetchImage(path, mimeType)`; the
 * resulting object URL is owned by this module and revoked when the entry is
 * dropped. The returned image is drawable with `ctx.drawImage` exactly like an
 * ImageBitmap. Rejects (and self-evicts, so the next call retries fresh) if the
 * SVG fails to load — callers should fall back to a raster representation on
 * rejection.
 */
export function getCachedSvgImageByPath(
  svgImagePath: string,
  fetchImage: (path: string, mimeType: string) => Promise<Blob>,
): Promise<HTMLImageElement> {
  const hit = svgByPathCache.get(svgImagePath);
  if (hit) {
    // Refresh LRU position.
    svgByPathCache.delete(svgImagePath);
    svgByPathCache.set(svgImagePath, hit);
    return hit;
  }
  const p = (async () => {
    const blob = await fetchImage(svgImagePath, 'image/svg+xml');
    const url = URL.createObjectURL(blob);
    urlByPath.set(svgImagePath, url);
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
  p.catch(() => dropEntry(svgImagePath));
  svgByPathCache.set(svgImagePath, p);
  if (svgByPathCache.size > MAX) {
    const oldest = svgByPathCache.keys().next().value as string;
    dropEntry(oldest);
  }
  return p;
}
