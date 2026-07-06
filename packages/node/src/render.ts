/**
 * Server-side rendering helpers. These adapt the browser-bound canvas
 * renderers in `@silurus/ooxml-{pptx,docx,xlsx}` to a user-supplied
 * Node canvas implementation (e.g. `skia-canvas`).
 *
 * The original renderers reach for a few browser-only globals:
 *   - `createImageBitmap` — used to paint embedded raster pictures
 *   - `new Image()` — used by xlsx image anchors
 *   - `document.fonts.add(new FontFace(...))` — used by the Google-Fonts loader
 *
 * We provide minimal Node shims for the first two (the third stays opt-in
 * via `useGoogleFonts: false`). The user passes in a canvas factory so the
 * package itself does not pin a particular Node canvas implementation;
 * `skia-canvas` is recommended in the README.
 */

import type { Presentation } from '@silurus/ooxml-pptx';

/** A subset of the Node-canvas API that the renderers actually need. The
 *  `skia-canvas` `Canvas` (and `@napi-rs/canvas`'s `Canvas`) both satisfy
 *  this — they expose the same `getContext('2d')` shape as the browser. */
export interface NodeCanvasLike {
  width: number;
  height: number;
  getContext(kind: '2d'): CanvasRenderingContext2D;
  /** Encode to PNG bytes. skia-canvas: `canvas.png` async getter or
   *  `toBuffer('png')`. @napi-rs/canvas: `toBuffer('image/png')`. */
  toBuffer?(format?: string): Buffer | Promise<Buffer>;
}

export interface NodeImageLike {
  width: number;
  height: number;
}

export interface NodeCanvasFactory {
  /** Create a blank canvas of the given pixel size. */
  createCanvas(width: number, height: number): NodeCanvasLike;
  /** Decode a buffer (PNG/JPEG/etc.) into something the canvas can `drawImage`. */
  loadImage(buffer: ArrayBuffer | Uint8Array | Buffer): Promise<NodeImageLike>;
}

/**
 * Polyfill `globalThis.OffscreenCanvas` so the shared rendering primitives can
 * allocate auxiliary canvases under Node.
 *
 * `packages/core/src/shape/effects.ts`'s `createAuxCanvas` probes
 * `typeof OffscreenCanvas !== 'undefined'` and otherwise `document`. Under Node
 * neither exists, so it returns `null` and the pptx renderer's beveled-flat
 * path, scene3d projection, and the inner-shadow / soft-edge / reflection
 * effects all *silently* degrade to flat output (no rim shading, no 3D warp,
 * no blur). This shim makes `new OffscreenCanvas(w, h)` allocate a real
 * skia-canvas (via the user's `factory.createCanvas`), so those paths light up
 * server-side exactly as they do in the browser.
 *
 * `new OffscreenCanvas(w, h)` simply *returns* a backing canvas from
 * `factory.createCanvas` (a class constructor may return a different object).
 * Returning the real canvas — rather than a wrapper — matters: the effect
 * helpers pass the allocated canvas straight to `ctx.drawImage(aux, …)`, and
 * skia-canvas's `drawImage` only accepts a real `Image`/`Canvas`, rejecting a
 * forwarding wrapper. The backing canvas already exposes everything
 * `createAuxCanvas` consumers touch (getContext, getImageData/putImageData,
 * drawImage as a source, and `ctx.filter = 'blur(Npx)'`); skia-canvas supports
 * all of them, including the `filter` blur used by the effect helpers.
 *
 * If `globalThis.OffscreenCanvas` is already defined (e.g. on a real DOM/worker
 * runtime, or Node ≥ a future version that ships it) the existing value is left
 * untouched. The returned function restores the global to its pre-call value.
 */
export function installOffscreenCanvasShim(factory: NodeCanvasFactory): () => void {
  const g = globalThis as unknown as { OffscreenCanvas?: unknown };
  const prev = g.OffscreenCanvas;
  const hadOwn = Object.prototype.hasOwnProperty.call(globalThis, 'OffscreenCanvas');

  // Respect a pre-existing implementation — never overwrite a real one.
  if (typeof prev !== 'undefined') {
    return () => {
      /* nothing to restore: we never touched the global */
    };
  }

  class OffscreenCanvasShim {
    constructor(width: number, height: number) {
      // Return the backing canvas itself (constructors may return another
      // object). This keeps `aux instanceof <skia Canvas>` true so skia's
      // `drawImage` accepts it as an image source.
      return factory.createCanvas(width, height) as unknown as OffscreenCanvasShim;
    }
  }

  g.OffscreenCanvas = OffscreenCanvasShim as unknown;

  return () => {
    if (hadOwn) {
      g.OffscreenCanvas = prev;
    } else {
      delete g.OffscreenCanvas;
    }
  };
}

/** Polyfill `globalThis.createImageBitmap` so the existing renderers can
 *  decode raster pictures. Wires it to the user's `loadImage`. Returns the
 *  previous global (if any) so the caller can restore it. */
export function installImageBitmapShim(factory: NodeCanvasFactory): () => void {
  const g = globalThis as unknown as { createImageBitmap?: unknown };
  const prev = g.createImageBitmap;
  // The source may be raw bytes (the raster-decode path) OR a canvas-like
  // surface with a 2D context (core's `applyDuotone`, §20.1.8.23, which bakes a
  // recoloured offscreen surface back into an "ImageBitmap"). Widen the param so
  // the canvas branch is a real member and needs no double-cast.
  type CanvasLike = { getContext(id: '2d'): unknown };
  g.createImageBitmap = async (source: Blob | ArrayBuffer | Uint8Array | CanvasLike) => {
    // A canvas-like source is already a drawable image source in node-canvas —
    // return it directly. The surface IS the skia Canvas the factory made, so no
    // byte round-trip is needed (and skia has no `createImageBitmap(canvas)`).
    if (source && typeof (source as CanvasLike).getContext === 'function') {
      return source as CanvasLike;
    }
    let buf: ArrayBuffer | Uint8Array | Buffer;
    if (source instanceof Uint8Array || source instanceof ArrayBuffer) {
      buf = source;
    } else if (typeof (source as Blob).arrayBuffer === 'function') {
      buf = await (source as Blob).arrayBuffer();
    } else {
      throw new Error('createImageBitmap shim: unsupported source type');
    }
    return factory.loadImage(buf);
  };
  return () => { g.createImageBitmap = prev as typeof globalThis.createImageBitmap; };
}

/**
 * Build a `fetchImage` that reads embedded image bytes straight out of the
 * original `.pptx` archive via the WASM `extract_image` export — the Node twin
 * of the browser worker's in-worker `getImage` closure (render-worker.ts). The
 * lazy image pipeline carries only zip paths on pictures/blip fills, so the
 * source archive bytes are the byte source server-side; no base64 is ever
 * inlined. Mime travels on the element, so the renderer supplies it.
 *
 * `maxZipEntryBytes` mirrors the worker's per-entry guard and is optional
 * (no cap when omitted).
 */
export function makeSourceBufferFetchImage(
  sourceBuffer: ArrayBuffer | Uint8Array,
  maxZipEntryBytes?: number,
): (path: string, mimeType: string) => Promise<Blob> {
  return async (path: string, mimeType: string): Promise<Blob> => {
    // Dynamic import keeps the WASM parser binding out of this module's static
    // import graph: render.ts must load under Node without the git-ignored,
    // build-on-demand WASM artifacts present (CI runs `pnpm test` before
    // `pnpm build:wasm`). Mirrors the dynamic renderer import in renderSlideNode.
    const { extractImage } = await import('./pptx.ts');
    const bytes = extractImage(sourceBuffer, path, maxZipEntryBytes);
    // `.slice()` detaches from the WASM linear memory so the Blob owns a stable
    // copy (the WASM heap can be reused by the next call).
    return new Blob([new Uint8Array(bytes).slice() as BlobPart], { type: mimeType });
  };
}

/** Skeleton: render a single slide into a user-supplied Node canvas. The
 *  caller must:
 *   - have called `parsePptx(buffer)` to obtain `presentation`
 *   - install `createImageBitmap` shim via {@link installImageBitmapShim}
 *   - load fonts they want available into the canvas implementation's font
 *     registry (e.g. `Font.use(...)` for skia-canvas) BEFORE calling render
 *
 *  Returns the canvas; encode to PNG with `canvas.toBuffer('png')`.
 *
 *  Note: the underlying browser renderer is `async` and imports Vite-only
 *  worker assets at the top of `presentation.ts`. The Node path bypasses
 *  `PptxPresentation` and `worker.ts` entirely and calls the pure
 *  `renderSlide` function from `@silurus/ooxml-pptx`.
 *
 *  Bevel / scene3d / effects: pass `opts.factory` (the same canvas factory you
 *  used for {@link installImageBitmapShim}) so this function can install the
 *  {@link installOffscreenCanvasShim} for the duration of the render. Without a
 *  factory there is no way to allocate auxiliary canvases, and the renderer's
 *  beveled-flat path, scene3d projection, and inner-shadow / soft-edge /
 *  reflection effects silently fall back to flat output. The shim is restored
 *  before this function returns. */
export async function renderSlideNode(
  canvas: NodeCanvasLike,
  presentation: Presentation,
  slideIndex: number,
  opts: {
    width?: number;
    dpr?: number;
    factory?: NodeCanvasFactory;
    /**
     * The original `.pptx` archive bytes. When supplied (and no explicit
     * `fetchImage` is given), embedded images are painted by reading their bytes
     * straight out of this buffer via the WASM `extract_image` export — the Node
     * twin of the browser worker's in-worker image loader. Pictures and blip
     * fills carry only zip paths now (no inlined base64), so this is the byte
     * source server-side. Additive: omit it (and `fetchImage`) to keep the prior
     * behavior where pictures simply draw nothing.
     */
    sourceBuffer?: ArrayBuffer | Uint8Array;
    /**
     * Optional per-zip-entry byte cap forwarded to `extract_image`, mirroring the
     * browser worker's guard. Only consulted when `sourceBuffer` drives the
     * default `fetchImage`. No cap when omitted.
     */
    maxZipEntryBytes?: number;
    /**
     * Lazily resolve an embedded image (by zip path + MIME) to a Blob. Pictures
     * and blip fills carry only zip paths now (no inlined base64). Supplying
     * `sourceBuffer` builds this automatically; pass an explicit `fetchImage`
     * only to override that (e.g. a custom byte source). When neither is given,
     * defaults to an empty-Blob fetcher (images decode to nothing), matching the
     * media placeholder.
     */
    fetchImage?: (path: string, mimeType: string) => Promise<Blob>;
  } = {},
): Promise<void> {
  // Direct import of the pure renderer module — avoids `presentation.ts`
  // and `viewer.ts`, both of which pull Vite-specific worker / asset
  // imports that don't resolve under Node.
  const { renderSlide } = (await import('../../pptx/src/renderer.js')) as unknown as {
    renderSlide: (
      canvas: HTMLCanvasElement,
      slide: Presentation['slides'][number],
      slideWidth: number,
      slideHeight: number,
      opts: Record<string, unknown>,
    ) => Promise<HTMLCanvasElement>;
  };
  const slide = presentation.slides[slideIndex];
  if (!slide) throw new Error(`Slide index ${slideIndex} out of range`);
  const width = opts.width ?? 960;
  const dpr = opts.dpr ?? 2;
  // Light up bevel/scene3d/effects auxiliary-canvas allocation for the render,
  // then restore the global. No-op restore if a factory was not supplied or an
  // OffscreenCanvas already exists.
  const restoreOffscreen = opts.factory
    ? installOffscreenCanvasShim(opts.factory)
    : () => {};
  // Resolve the image byte source: an explicit `fetchImage` wins; otherwise, if
  // the caller handed us the source archive, read image bytes from it via
  // `extract_image`; otherwise fall back to the empty-Blob default (pictures
  // draw nothing) so existing callers without images keep working.
  const fetchImage =
    opts.fetchImage ??
    (opts.sourceBuffer
      ? makeSourceBufferFetchImage(opts.sourceBuffer, opts.maxZipEntryBytes)
      : async () => new Blob([]));
  try {
    await renderSlide(
      canvas as unknown as HTMLCanvasElement,
      slide,
      presentation.slideWidth,
      presentation.slideHeight,
      {
        width,
        dpr,
        defaultTextColor: presentation.defaultTextColor,
        majorFont: presentation.majorFont,
        minorFont: presentation.minorFont,
        hlinkColor: presentation.hlinkColor ?? null,
        // Node-side renderers don't run media playback, so an empty fetcher
        // is fine for posters.
        fetchMedia: async () => new Blob([]),
        // Pictures/blip fills now carry zip paths; bytes come from `sourceBuffer`
        // (via extract_image), an explicit `fetchImage`, or — when neither is
        // given — an empty Blob so text/shape-only renders still work.
        fetchImage,
        skipMediaControls: true,
      },
    );
  } finally {
    restoreOffscreen();
  }
}
