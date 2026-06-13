import type { MediaElement, Presentation, WorkerRequest, WorkerResponse } from './types';
import { renderSlide, type TextRunCallback } from './renderer';
import { createPresentationHandle, type PresentationHandle } from './presentation-handle';
import { selectNotes } from './notes';
import {
  preloadGoogleFonts,
  WorkerBridge,
  defaultDpr,
  isHTMLCanvas,
  type LoadOptions as CoreLoadOptions,
  type MathRenderer,
} from '@silurus/ooxml-core';
import { PPTX_GOOGLE_FONTS, pptxFontPreloadNames } from './google-fonts';
import { findMimeTypeForPath } from './media-mime';
import type {
  PresentationMeta,
  RenderWorkerRequest,
  RenderWorkerResponse,
} from './worker-protocol';
import InlineWorker from './worker.ts?worker&inline';
import wasmAssetUrl from './wasm/pptx_parser_bg.wasm?url';

/** Options for {@link PptxPresentation.load}. */
export type LoadOptions = CoreLoadOptions & {
  /**
   * 'main' (default): parse in a worker, render on the main thread (current
   * behaviour). 'worker': parse AND render inside the worker; use
   * {@link PptxPresentation.renderSlideToBitmap} and paint the returned
   * ImageBitmap via an `ImageBitmapRenderingContext`. Requires OffscreenCanvas.
   */
  mode?: 'main' | 'worker';
};

/** Options for {@link PptxPresentation.renderSlideToBitmap}. */
export interface RenderSlideToBitmapOptions {
  /** Slide width in CSS pixels. Defaults to 960. */
  width?: number;
  /** Device pixel ratio. Defaults to window.devicePixelRatio (workers have none). */
  dpr?: number;
  /**
   * Skip the static media play-badge so a live overlay can draw its own
   * controls. Used internally by {@link PptxPresentation.presentSlide}.
   * @internal
   */
  skipMediaControls?: boolean;
}

/** Options for rendering a single slide onto a canvas. */
export interface RenderSlideOptions {
  /** Display width in CSS pixels. Defaults to canvas.offsetWidth or 960. */
  width?: number;
  /** Device pixel ratio. Defaults to window.devicePixelRatio or 1. */
  dpr?: number;
  /** Called for each rendered text segment. Used to build a transparent text selection overlay. */
  onTextRun?: TextRunCallback;
  /**
   * Skip drawing the play badge overlay on media elements. Used internally by
   * {@link PptxPresentation.presentSlide} so its interactive handle can draw
   * its own play/pause chrome without duplication.
   */
  skipMediaControls?: boolean;
}

/**
 * Headless PPTX rendering engine.
 *
 * Parses `.pptx` archives in a background worker (WASM) but renders slides
 * synchronously on the main thread, so the canvas shares the document's
 * `FontFaceSet` — avoiding subtle wrap differences between system fallback
 * fonts and theme-declared webfonts (e.g. Nunito Sans).
 *
 * Construct via the static `load` factory. A single instance can drive any
 * number of canvases (scroll view, thumbnail grid, master-detail, etc.).
 *
 * @example
 * const pres = await PptxPresentation.load(buffer);
 * await pres.renderSlide(canvas, 0, { width: 960 });
 */
export class PptxPresentation {
  private readonly _worker: Worker;
  private readonly _bridge: WorkerBridge<WorkerResponse | RenderWorkerResponse>;
  private _mode: 'main' | 'worker' = 'main';
  private _presentation: Presentation | null = null;
  private _meta: PresentationMeta | null = null;
  private _mediaCache = new Map<string, Promise<Blob>>();
  private _workerReady = false;
  private _workerReadyCallbacks: Array<() => void> = [];
  /** Opt-in OMML equation engine, injected once at {@link load}. Every
   *  `renderSlide` / `presentSlide` reuses it — equations render when present,
   *  and are skipped (engine tree-shaken) when omitted. */
  private _math: MathRenderer | undefined;

  private constructor(worker: Worker, mode: 'main' | 'worker') {
    this._worker = worker;
    this._mode = mode;
    this._bridge = new WorkerBridge<WorkerResponse | RenderWorkerResponse>(this._worker, {
      // The init `ready` handshake carries no id; everything else does.
      correlate: (msg) => ('id' in msg ? msg.id : undefined),
      toError: (msg) => (msg.kind === 'error' ? msg.message : undefined),
      onUnsolicited: (msg) => {
        if (msg.kind === 'ready') {
          this._workerReady = true;
          for (const cb of this._workerReadyCallbacks) cb();
          this._workerReadyCallbacks = [];
        }
      },
    });
    const wasmUrl = new URL(wasmAssetUrl, location.href).href;
    this._bridge.post({ kind: 'init', wasmUrl } satisfies WorkerRequest);
  }

  /** Parse a PPTX from URL or ArrayBuffer. */
  static async load(
    source: string | ArrayBuffer,
    opts: LoadOptions = {},
  ): Promise<PptxPresentation> {
    const mode = opts.mode ?? 'main';
    if (mode === 'worker' && (typeof Worker === 'undefined' || typeof OffscreenCanvas === 'undefined')) {
      throw new Error("mode: 'worker' requires Worker and OffscreenCanvas support");
    }
    // The render worker is reachable only through this dynamic import, so
    // main-mode bundles never pull in its (renderer-bearing) chunk.
    const worker =
      mode === 'worker'
        ? (await import('./render-worker-host')).createRenderWorker()
        : new InlineWorker();
    const pres = new PptxPresentation(worker, mode);
    let buffer: ArrayBuffer;
    if (typeof source === 'string') {
      const res = await fetch(source);
      if (!res.ok) throw new Error(`Failed to fetch: ${res.status} ${res.statusText}`);
      buffer = await res.arrayBuffer();
    } else {
      buffer = source;
    }
    if (opts.math && mode === 'worker') {
      console.warn(
        "[ooxml] the math engine is unavailable in mode: 'worker'; equations will be skipped. Use mode: 'main' for documents with equations.",
      );
    }
    pres._math = mode === 'worker' ? undefined : opts.math;
    await pres._parse(
      buffer,
      opts.maxZipEntryBytes,
      mode === 'worker' ? !!opts.useGoogleFonts : false,
    );
    if (mode === 'main' && opts.useGoogleFonts && pres._presentation) {
      const parsed = pres._presentation;
      await preloadGoogleFonts(
        pptxFontPreloadNames(parsed.majorFont, parsed.minorFont),
        PPTX_GOOGLE_FONTS,
      );
    }
    return pres;
  }

  private _waitForWorker(): Promise<void> {
    if (this._workerReady) return Promise.resolve();
    return new Promise((resolve) => this._workerReadyCallbacks.push(resolve));
  }

  private async _parse(
    buffer: ArrayBuffer,
    maxZipEntryBytes?: number,
    useGoogleFonts = false,
  ): Promise<void> {
    await this._waitForWorker();
    const res = await this._bridge.request(
      (id) =>
        this._mode === 'worker'
          ? ({ kind: 'parse', id, buffer, maxZipEntryBytes, useGoogleFonts } satisfies RenderWorkerRequest)
          : ({ kind: 'parse', id, buffer, maxZipEntryBytes } satisfies WorkerRequest),
      [buffer],
    );
    if (this._mode === 'worker') {
      this._meta = (res as Extract<RenderWorkerResponse, { kind: 'parsedMeta' }>).meta;
    } else {
      this._presentation = (res as Extract<WorkerResponse, { kind: 'parsed' }>).presentation;
    }
  }

  /** Total number of slides in the loaded presentation. */
  get slideCount(): number { return this._presentation?.slides.length ?? this._meta?.slideCount ?? 0; }

  /** Slide width in EMU. */
  get slideWidth(): number { return this._presentation?.slideWidth ?? this._meta?.slideWidth ?? 0; }

  /** Slide height in EMU. */
  get slideHeight(): number { return this._presentation?.slideHeight ?? this._meta?.slideHeight ?? 0; }

  /**
   * Speaker-notes text for a slide (`ppt/notesSlides/notesSlideN.xml`,
   * ECMA-376 §13.3.5 — Notes Slide). Returns the notes-body text as a single
   * string (paragraphs joined with `\n`), or `null` when the slide has no
   * notes part. The notes are parsed at {@link load} time, so this is a
   * synchronous lookup.
   *
   * `slideIndex` is 0-based. Unlike navigation methods it is *not* clamped:
   * an out-of-range or non-integer index returns `null` rather than the notes
   * of the nearest slide (so a tool iterating by index gets an honest "no
   * notes" instead of a duplicated neighbour).
   *
   * @example
   * const pres = await PptxPresentation.load(buffer);
   * for (let i = 0; i < pres.slideCount; i++) {
   *   const notes = pres.getNotes(i);
   *   if (notes) console.log(`Slide ${i + 1} notes:`, notes);
   * }
   */
  getNotes(slideIndex: number): string | null {
    if (this._meta) {
      // Worker mode: the model lives in the worker, so honour the same
      // non-clamped contract against the per-slide notes array.
      return Number.isInteger(slideIndex) ? (this._meta.notes[slideIndex] ?? null) : null;
    }
    return selectNotes(this._presentation?.slides ?? [], slideIndex);
  }

  /** Render a slide onto the given canvas. */
  async renderSlide(
    canvas: HTMLCanvasElement | OffscreenCanvas,
    slideIndex: number,
    opts: RenderSlideOptions = {},
  ): Promise<void> {
    if (this._mode === 'worker') {
      throw new Error(
        "renderSlide(canvas) is unavailable in mode: 'worker'; use renderSlideToBitmap() and paint it via an ImageBitmapRenderingContext",
      );
    }
    if (!this._presentation) throw new Error('Presentation not loaded');
    const slide = this._presentation.slides[slideIndex];
    if (!slide) throw new Error(`Slide index ${slideIndex} out of range (count: ${this.slideCount})`);
    const dpr = opts.dpr ?? defaultDpr();
    const width = opts.width ?? ((isHTMLCanvas(canvas) ? canvas.offsetWidth : 0) || 960);
    await renderSlide(
      canvas,
      slide,
      this._presentation.slideWidth,
      this._presentation.slideHeight,
      {
        width,
        dpr,
        defaultTextColor: this._presentation.defaultTextColor,
        majorFont: this._presentation.majorFont,
        minorFont: this._presentation.minorFont,
        hlinkColor: this._presentation.hlinkColor ?? null,
        fetchMedia: (path) => this.getMedia(path),
        skipMediaControls: opts.skipMediaControls,
        math: this._math,
      },
      opts.onTextRun,
    );
  }

  /**
   * Render a slide and return it as an ImageBitmap. Works in both modes; in
   * worker mode the entire render runs off the main thread. Paint with:
   * `canvas.getContext('bitmaprenderer').transferFromImageBitmap(bitmap)`.
   *
   * The returned ImageBitmap is owned by the caller: pass it to
   * `transferFromImageBitmap` (which consumes it) or call `bitmap.close()`
   * when done, or its backing memory is held until GC.
   */
  async renderSlideToBitmap(
    slideIndex: number,
    opts: RenderSlideToBitmapOptions = {},
  ): Promise<ImageBitmap> {
    const width = opts.width ?? 960;
    const dpr = opts.dpr ?? defaultDpr();
    if (this._mode === 'worker') {
      if (!Number.isInteger(slideIndex) || slideIndex < 0 || slideIndex >= this.slideCount) {
        throw new Error(`Slide index ${slideIndex} out of range (count: ${this.slideCount})`);
      }
      const res = await this._bridge.request(
        (id) => ({ kind: 'renderSlide', id, slideIndex, width, dpr, skipMediaControls: opts.skipMediaControls }) satisfies RenderWorkerRequest,
      );
      return (res as Extract<RenderWorkerResponse, { kind: 'slideRendered' }>).bitmap;
    }
    const off = new OffscreenCanvas(1, 1);
    await this.renderSlide(off, slideIndex, { width, dpr, skipMediaControls: opts.skipMediaControls });
    return off.transferToImageBitmap();
  }

  /**
   * Extract raw media bytes for a zip path referenced by {@link MediaElement}.
   * Results are cached by path for the lifetime of this instance.
   */
  async getMedia(mediaPath: string): Promise<Blob> {
    const hit = this._mediaCache.get(mediaPath);
    if (hit) return hit;
    // Worker mode has no main-thread model, so the mime lookup is skipped and
    // the Blob carries an empty type. That is fine: presentation-handle.ts
    // re-types blobs from MediaElement.mimeType when it builds media controls.
    const mimeType = this._findMimeTypeForPath(mediaPath);
    const p = (async () => {
      await this._waitForWorker();
      const res = await this._bridge.request(
        (id) => ({ kind: 'extractMedia', id, path: mediaPath }) satisfies WorkerRequest,
      );
      const bytes = (res as Extract<WorkerResponse, { kind: 'mediaExtracted' }>).bytes;
      return new Blob([bytes], { type: mimeType });
    })();
    this._mediaCache.set(mediaPath, p);
    return p;
  }

  private _findMimeTypeForPath(mediaPath: string): string {
    if (!this._presentation) return '';
    return findMimeTypeForPath(this._presentation, mediaPath);
  }

  /**
   * Render a slide and attach canvas-native playback controls for any
   * embedded audio/video. Returns a {@link PresentationHandle} that owns the
   * RAF loop, media elements, and object URLs. Unlike {@link renderSlide}, this
   * method is stateful — always call `handle.destroy()` when leaving the slide.
   */
  async presentSlide(
    canvas: HTMLCanvasElement,
    slideIndex: number,
    opts: RenderSlideOptions = {},
  ): Promise<PresentationHandle> {
    if (this._mode === 'main' && !this._presentation) {
      throw new Error('Presentation not loaded');
    }
    if (!Number.isInteger(slideIndex) || slideIndex < 0 || slideIndex >= this.slideCount) {
      throw new Error(`Slide index ${slideIndex} out of range (count: ${this.slideCount})`);
    }
    const dpr = opts.dpr ?? defaultDpr();
    const width = opts.width ?? (canvas.offsetWidth || 960);

    if (this._mode === 'worker' && opts.onTextRun) {
      // The callback can't cross the worker boundary.
      console.warn(
        "[ooxml] onTextRun is unavailable in mode: 'worker'; the text selection overlay will be empty for this slide.",
      );
    }

    const drawBase =
      this._mode === 'worker'
        ? async () => {
            // Whole slide rendered off-thread; the handle snapshots this paint
            // into its own base copy, so the bitmap can be closed right after.
            const bmp = await this.renderSlideToBitmap(slideIndex, { width, dpr, skipMediaControls: true });
            canvas.width = bmp.width;
            canvas.height = bmp.height;
            // Set only the CSS width and let height follow the intrinsic aspect
            // ratio — mirrors the main renderer (renderer.ts), which avoids an
            // explicit style.height that could fight the ratio.
            canvas.style.width = `${Math.round(bmp.width / dpr)}px`;
            if (!canvas.style.display) canvas.style.display = 'block';
            const ctx = canvas.getContext('2d');
            if (!ctx) throw new Error('2D context not available');
            ctx.drawImage(bmp, 0, 0);
            bmp.close();
          }
        : () =>
            this.renderSlide(canvas, slideIndex, {
              width,
              dpr,
              skipMediaControls: true,
              onTextRun: opts.onTextRun,
            });

    const mediaElements =
      this._mode === 'worker'
        ? (this._meta?.mediaElements[slideIndex] ?? [])
        : (this._presentation as Presentation).slides[slideIndex].elements.filter(
            (el): el is MediaElement => el.type === 'media',
          );

    return createPresentationHandle(canvas, mediaElements, {
      width,
      dpr,
      slideWidthEmu: this.slideWidth,
      fetchMedia: (path) => this.getMedia(path),
      drawBase,
    });
  }

  /** Terminate the worker and release all resources. */
  destroy(): void {
    this._bridge.terminate();
    this._presentation = null;
    this._meta = null;
    this._mediaCache.clear();
  }
}
