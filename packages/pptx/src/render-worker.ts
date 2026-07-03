/**
 * Render-capable worker entry: a superset of `worker.ts` that keeps the parsed
 * Presentation worker-side and renders slides into an OffscreenCanvas,
 * replying with transferable ImageBitmaps. Used by
 * `PptxPresentation.load(src, { mode: 'worker' })`; the slim parse-only
 * `worker.ts` stays untouched so main-mode users pay no bundle growth.
 *
 * Single-document contract: the proxy issues one `parse` and then renders.
 * A re-parse while a render is suspended mid-await would let that render
 * resume against the new model — callers must not interleave them.
 */
import type { MediaElement, Presentation } from './types';
import init, { PptxArchive } from './wasm/pptx_parser.js';
import { renderSlide } from './renderer';
import { selectNotes } from './notes';
import { findMimeTypeForPath } from './media-mime';
import { PPTX_GOOGLE_FONTS, pptxFontPreloadNames } from './google-fonts';
import { preloadGoogleFonts, decodeDataUrl } from '@silurus/ooxml-core';
import type { RenderWorkerRequest, RenderWorkerResponse, PresentationMeta } from './worker-protocol';

let initPromise: Promise<unknown> | null = null;
let pres: Presentation | null = null;
// A `PptxArchive` handle over the opened zip. `new PptxArchive(bytes, max)`
// copies the file into WASM ONCE and opens it ONCE; the in-worker
// `getMedia` / `getImage` closures then read bytes by zip path straight from the
// retained archive (no transfer, no re-open, no JS-side buffer kept alive).
// Freed + replaced on a re-parse.
let archive: PptxArchive | null = null;
/** Settled before any render when `useGoogleFonts` was requested. */
let fontsLoaded: Promise<void> = Promise.resolve();
const mediaCache = new Map<string, Promise<Blob>>();
const imageCache = new Map<string, Promise<Blob>>();

const post = (msg: RenderWorkerResponse, transfer?: Transferable[]) =>
  (self.postMessage as (m: unknown, t?: Transferable[]) => void)(msg, transfer);

/** Free the current handle (if any) and null it out — double-free / UAF guard. */
function disposeArchive(): void {
  if (archive) {
    archive.free();
    archive = null;
  }
}

function getMedia(path: string): Promise<Blob> {
  const hit = mediaCache.get(path);
  if (hit) return hit;
  const p = (async () => {
    if (!archive || !pres) throw new Error('No pptx loaded');
    const bytes = archive.extract_media(path);
    return new Blob([new Uint8Array(bytes).slice()], { type: findMimeTypeForPath(pres, path) });
  })();
  mediaCache.set(path, p);
  return p;
}

/** In-worker image-byte loader (twin of {@link getMedia}). The renderer's
 *  `fetchImage` routes here in worker mode, so image bytes are decoded straight
 *  from the retained archive with no main-thread round-trip. Mime travels on the
 *  element, so the caller supplies it. */
function getImage(path: string, mimeType: string): Promise<Blob> {
  const hit = imageCache.get(path);
  if (hit) return hit;
  const p = (async () => {
    if (!archive) throw new Error('No pptx loaded');
    const bytes = archive.extract_image(path);
    return new Blob([new Uint8Array(bytes).slice()], { type: mimeType });
  })();
  imageCache.set(path, p);
  return p;
}

self.onmessage = async (e: MessageEvent<RenderWorkerRequest>) => {
  const req = e.data;

  if (req.kind === 'init') {
    // Retain the init promise (docx/xlsx pattern) rather than a `ready` flag +
    // handshake. Every request below `await`s it, so a REJECTED init rejects the
    // request (the outer catch posts an `error` response the bridge turns into a
    // rejected `load()` / render), never a silent hang on a main-side wait.
    initPromise = init(decodeDataUrl(req.wasmUrl) ?? req.wasmUrl);
    return;
  }

  try {
    await initPromise;
    if (req.kind === 'parse') {
      // Cached blobs belong to the previous document; serving them after a
      // re-parse would silently return the wrong file's media/image.
      mediaCache.clear();
      imageCache.clear();
      const max =
        typeof req.maxZipEntryBytes === 'number' && req.maxZipEntryBytes > 0
          ? BigInt(req.maxZipEntryBytes)
          : undefined;
      // Constructing the handle copies `req.buffer` into WASM; the worker then
      // holds no reference to those bytes (memory is not doubled). Replace any
      // prior handle first so a re-parse frees the old archive.
      disposeArchive();
      const bytes = new Uint8Array(req.buffer);
      archive = new PptxArchive(bytes, max);
      // `parse()` returns UTF-8 JSON bytes (Result<Vec<u8>, JsValue>). Render mode
      // consumes the model in-worker, so decode + parse here (one decode, no
      // passthrough).
      pres = JSON.parse(new TextDecoder().decode(archive.parse())) as Presentation;
      if (req.useGoogleFonts) {
        // Kick the preload now so it overlaps with main-thread work; renders
        // await `fontsLoaded` so text never rasterizes with fallback metrics.
        fontsLoaded = preloadGoogleFonts(
          pptxFontPreloadNames(pres),
          PPTX_GOOGLE_FONTS,
        );
      }
      const meta: PresentationMeta = {
        slideCount: pres.slides.length,
        slideWidth: pres.slideWidth,
        slideHeight: pres.slideHeight,
        majorFont: pres.majorFont ?? null,
        minorFont: pres.minorFont ?? null,
        notes: pres.slides.map((_, i) => selectNotes(pres!.slides, i)),
        mediaElements: pres.slides.map((s) =>
          s.elements.filter((el): el is MediaElement => el.type === 'media')),
        hidden: pres.slides.map((s) => s.hidden ?? false),
      };
      post({ kind: 'parsedMeta', id: req.id, meta });
      return;
    }

    if (req.kind === 'renderSlide') {
      if (!pres) throw new Error('No pptx loaded');
      const slide = pres.slides[req.slideIndex];
      if (!slide) throw new Error(`Slide index ${req.slideIndex} out of range (count: ${pres.slides.length})`);
      await fontsLoaded;
      const canvas = new OffscreenCanvas(1, 1); // renderSlide resizes it
      await renderSlide(canvas, slide, pres.slideWidth, pres.slideHeight, {
        width: req.width,
        dpr: req.dpr,
        defaultTextColor: pres.defaultTextColor,
        majorFont: pres.majorFont,
        minorFont: pres.minorFont,
        hlinkColor: pres.hlinkColor ?? null,
        fetchMedia: getMedia,
        fetchImage: getImage,
        skipMediaControls: req.skipMediaControls,
        dim: req.dim,
        // math intentionally omitted: MathJax needs a DOM <script>; worker
        // mode skips equations (documented in the design spec).
      });
      const bitmap = canvas.transferToImageBitmap();
      post({ kind: 'slideRendered', id: req.id, bitmap }, [bitmap]);
      return;
    }

    if (req.kind === 'extractMedia') {
      const blob = await getMedia(req.path);
      const bytes = await blob.arrayBuffer();
      post({ kind: 'mediaExtracted', id: req.id, bytes }, [bytes]);
      return;
    }

    if (req.kind === 'extractImage') {
      // Worker render mode decodes images in-worker via the getImage closure;
      // this message arm exists only for protocol parity with worker.ts. Raw
      // bytes are read straight from the retained archive (no mime needed for a
      // byte transfer).
      if (!archive) throw new Error('No pptx loaded');
      const raw = archive.extract_image(req.path);
      const bytes = new Uint8Array(raw).slice().buffer;
      post({ kind: 'imageExtracted', id: req.id, bytes }, [bytes]);
      return;
    }

    if (req.kind === 'toMarkdown') {
      // Project the retained archive to markdown, straight from the handle the
      // worker already holds (same source as worker.ts's parse-mode arm).
      if (!archive) throw new Error('No pptx loaded');
      const markdown = archive.to_markdown();
      post({ kind: 'markdownRendered', id: req.id, markdown });
      return;
    }
  } catch (err) {
    if ('id' in req) {
      post({ kind: 'error', id: req.id, message: err instanceof Error ? err.message : String(err) });
    }
  }
};
