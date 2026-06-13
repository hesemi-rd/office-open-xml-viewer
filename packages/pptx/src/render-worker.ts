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
import init, { parse_pptx, extract_media } from './wasm/pptx_parser.js';
import { renderSlide } from './renderer';
import { selectNotes } from './notes';
import { findMimeTypeForPath } from './media-mime';
import { PPTX_GOOGLE_FONTS, pptxFontPreloadNames } from './google-fonts';
import { preloadGoogleFonts, decodeDataUrl } from '@silurus/ooxml-core';
import type { RenderWorkerRequest, RenderWorkerResponse, PresentationMeta } from './worker-protocol';

let ready = false;
let pres: Presentation | null = null;
let currentBuffer: Uint8Array | null = null;
let currentMaxZipEntryBytes: bigint | undefined;
/** Settled before any render when `useGoogleFonts` was requested. */
let fontsLoaded: Promise<void> = Promise.resolve();
const mediaCache = new Map<string, Promise<Blob>>();

const post = (msg: RenderWorkerResponse, transfer?: Transferable[]) =>
  (self.postMessage as (m: unknown, t?: Transferable[]) => void)(msg, transfer);

async function initWasm(wasmUrl: string) {
  await init(decodeDataUrl(wasmUrl) ?? wasmUrl);
  ready = true;
  post({ kind: 'ready' });
}

function getMedia(path: string): Promise<Blob> {
  const hit = mediaCache.get(path);
  if (hit) return hit;
  const p = (async () => {
    if (!currentBuffer || !pres) throw new Error('No pptx loaded');
    const bytes = extract_media(currentBuffer, path, currentMaxZipEntryBytes);
    return new Blob([new Uint8Array(bytes).slice()], { type: findMimeTypeForPath(pres, path) });
  })();
  mediaCache.set(path, p);
  return p;
}

self.onmessage = async (e: MessageEvent<RenderWorkerRequest>) => {
  const req = e.data;

  if (req.kind === 'init') {
    initWasm(req.wasmUrl).catch((err) => {
      console.error('[pptx-render-worker] WASM init failed:', err);
    });
    return;
  }

  try {
    if (req.kind === 'parse') {
      if (!ready) throw new Error('WASM not initialized');
      // Cached blobs belong to the previous document; serving them after a
      // re-parse would silently return the wrong file's media.
      mediaCache.clear();
      currentBuffer = new Uint8Array(req.buffer);
      currentMaxZipEntryBytes =
        typeof req.maxZipEntryBytes === 'number' && req.maxZipEntryBytes > 0
          ? BigInt(req.maxZipEntryBytes)
          : undefined;
      pres = JSON.parse(parse_pptx(currentBuffer, currentMaxZipEntryBytes)) as Presentation;
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
        skipMediaControls: req.skipMediaControls,
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
  } catch (err) {
    if ('id' in req) {
      post({ kind: 'error', id: req.id, message: err instanceof Error ? err.message : String(err) });
    }
  }
};
