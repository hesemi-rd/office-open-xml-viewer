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

/** Polyfill `globalThis.createImageBitmap` so the existing renderers can
 *  decode raster pictures. Wires it to the user's `loadImage`. Returns the
 *  previous global (if any) so the caller can restore it. */
export function installImageBitmapShim(factory: NodeCanvasFactory): () => void {
  const g = globalThis as unknown as { createImageBitmap?: unknown };
  const prev = g.createImageBitmap;
  g.createImageBitmap = async (source: Blob | ArrayBuffer | Uint8Array) => {
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
 *  `renderSlide` function from `@silurus/ooxml-pptx`. */
export async function renderSlideNode(
  canvas: NodeCanvasLike,
  presentation: Presentation,
  slideIndex: number,
  opts: { width?: number; dpr?: number } = {},
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
      // is fine — picture elements with `dataUrl` are handled by the
      // existing path through createImageBitmap.
      fetchMedia: async () => new Blob([]),
      skipMediaControls: true,
    },
  );
}
