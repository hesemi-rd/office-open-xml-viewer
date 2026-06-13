// API reference data for the per-format pages. Hand-extracted from the real
// source (viewer.ts / presentation.ts / document.ts + the shared RenderOptions
// / LoadOptions). Keep in sync when the public types change.

export interface ApiOption {
  name: string;
  type: string;
  def?: string;
  desc: string;
}
export interface ApiMethod {
  sig: string;
  desc: string;
}
export interface ApiClass {
  name: string;
  ctor: string;
  note?: string;
  options?: ApiOption[];
  methods: ApiMethod[];
}

const ZIP = { name: 'maxZipEntryBytes', type: 'number', def: '512 MiB', desc: 'Per-entry ZIP decompression cap (zip-bomb guard). Lower it for untrusted input.' };
const GFONTS = { name: 'useGoogleFonts', type: 'boolean', def: 'false', desc: 'Load metric-compatible webfonts and non-Latin script fallbacks (Noto Arabic / CJK KR·SC·TC·JP / Cyrillic / Hebrew / Thai / Devanagari) from Google Fonts so layout matches Office and non-Latin text never falls back to tofu. Off by default for privacy.' };
const DPR = { name: 'dpr', type: 'number', def: 'devicePixelRatio', desc: 'Device pixel ratio for the backing store (crispness on HiDPI).' };
const MATH = { name: 'math', type: 'MathRenderer', def: 'undefined', desc: 'Opt-in OMML equation engine (MathJax + STIX Two Math, ~4 MB). Import it from the separate @silurus/ooxml/math entry — `import { math } from "@silurus/ooxml/math"` — and pass it to render equations. Omit it and equations are skipped — the engine tree-shakes away entirely.' };
const MODE = { name: 'mode', type: "'main' | 'worker'", def: "'main'", desc: "'main' parses in a worker and renders on the main thread (default). 'worker' parses AND renders entirely inside the worker; the main thread only paints the ImageBitmap returned by the render*ToBitmap method via a `bitmaprenderer` context. Requires Worker + OffscreenCanvas. The canvas-target render methods are unavailable in 'worker' mode, and equations require 'main'. Trade-off: each frame is transferred from the worker as an ImageBitmap, so a single render can be marginally slower than 'main' — the win is that the main thread never blocks." };
const VIEWER_MODE = { name: 'mode', type: "'main' | 'worker'", def: "'main'", desc: "'main' renders on the main thread (default). 'worker' renders the whole viewer off the main thread — every frame is produced in a Web Worker and painted via a `bitmaprenderer` context — so document rendering never blocks the UI. Scroll, sheet tabs, zoom and (xlsx) cell selection are unchanged. Requires Worker + OffscreenCanvas. The pptx/docx text-selection overlay is unavailable in 'worker' mode (onTextRun can't cross the worker boundary), and equations require 'main'. Trade-off: each frame crosses the worker boundary as an ImageBitmap, so an individual render can be marginally slower than 'main' — the win is a responsive main thread, not raw render speed." };

export const apiReference: Record<'docx' | 'xlsx' | 'pptx', ApiClass[]> = {
  pptx: [
    {
      name: 'PptxViewer',
      ctor: 'new PptxViewer(canvas: HTMLCanvasElement, options?: PptxViewerOptions)',
      note: 'Opinionated single-canvas viewer. Hand it a <canvas>; it owns parsing, rendering and the current slide.',
      options: [
        { name: 'width', type: 'number', def: '960', desc: 'Canvas CSS width in px; height is derived from the slide aspect ratio.' },
        DPR,
        GFONTS,
        { name: 'enableTextSelection', type: 'boolean', def: 'false', desc: 'Overlay a transparent text layer so users can select & copy slide text.' },
        { name: 'enableMediaPlayback', type: 'boolean', def: 'false', desc: 'Make embedded audio/video interactive (the viewer draws its own play chrome).' },
        ZIP,
        MATH,
        VIEWER_MODE,
        { name: 'onSlideChange', type: '(index: number, total: number) => void', desc: 'Called after a slide finishes rendering.' },
        { name: 'onError', type: '(err: Error) => void', desc: 'Called on parse or render errors.' },
      ],
      methods: [
        { sig: 'load(source: string | ArrayBuffer): Promise<void>', desc: 'Load from a URL or ArrayBuffer and render the first slide.' },
        { sig: 'goToSlide(index: number): Promise<void>', desc: 'Render a specific slide (0-indexed, clamped).' },
        { sig: 'nextSlide(): Promise<void>', desc: 'Advance one slide.' },
        { sig: 'prevSlide(): Promise<void>', desc: 'Go back one slide.' },
        { sig: 'get slideIndex(): number', desc: 'Current slide index.' },
        { sig: 'get slideCount(): number', desc: 'Total slides (0 until loaded).' },
        { sig: 'getNotes(slideIndex: number): string | null', desc: 'Speaker-notes text for a slide (0-based); null when the slide has no notes part or the index is out of range.' },
        { sig: 'get canvasElement(): HTMLCanvasElement', desc: 'The underlying canvas.' },
        { sig: 'destroy(): void', desc: 'Tear down the worker and release resources.' },
      ],
    },
    {
      name: 'PptxPresentation',
      ctor: 'await PptxPresentation.load(source, options?)',
      note: 'Headless engine — parse once, render any slide into any canvas you supply (scroll views, thumbnail grids, master–detail).',
      options: [GFONTS, ZIP, MATH, MODE],
      methods: [
        { sig: 'static load(source, options?): Promise<PptxPresentation>', desc: 'Parse a deck from a URL or ArrayBuffer.' },
        { sig: 'get slideCount(): number', desc: 'Total slides.' },
        { sig: 'renderSlide(canvas, index, opts?: { width?, dpr? }): Promise<void>', desc: 'Render one slide into the given canvas at the given width. Equations render when a `math` engine was passed to `load`. Unavailable in `mode: "worker"` — use renderSlideToBitmap.' },
        { sig: 'renderSlideToBitmap(index, opts?: { width?, dpr? }): Promise<ImageBitmap>', desc: 'Render one slide and return it as an ImageBitmap (both modes; in worker mode the render runs off the main thread). Equations are skipped in `mode: "worker"` (they require `mode: "main"`). The bitmap is caller-owned: pass it to `transferFromImageBitmap` (which consumes it) or call `bitmap.close()`.' },
        { sig: 'presentSlide(canvas, index, opts?: { width?, dpr?, onTextRun? }): Promise<PresentationHandle>', desc: 'Render a slide and attach canvas-native audio/video playback, returning a handle with play() / pause() / destroy(). Works in both modes — in `mode: "worker"` the base slide is rendered off the main thread and the video overlay is composited on the main thread; `onTextRun` is unavailable there (it cannot cross the worker boundary).' },
        { sig: 'getNotes(slideIndex: number): string | null', desc: 'Speaker-notes text for a slide (0-based; ECMA-376 §13.3.5). Returns null when the slide has no notes part or the index is out of range.' },
        { sig: 'destroy(): void', desc: 'Release the worker.' },
      ],
    },
  ],

  docx: [
    {
      name: 'DocxViewer',
      ctor: 'new DocxViewer(canvas: HTMLCanvasElement, options?: DocxViewerOptions)',
      note: 'Single-canvas viewer that paginates the document and tracks the current page.',
      options: [
        { name: 'width', type: 'number', desc: 'Canvas CSS width in px; height is auto-computed from the page aspect ratio.' },
        DPR,
        GFONTS,
        { name: 'enableTextSelection', type: 'boolean', def: 'false', desc: 'Overlay a transparent text layer for native selection & copy.' },
        { name: 'showTrackChanges', type: 'boolean', desc: 'Render tracked insertions/deletions with author colours.' },
        ZIP,
        MATH,
        VIEWER_MODE,
        { name: 'onPageChange', type: '(index: number, total: number) => void', desc: 'Called after a page finishes rendering.' },
        { name: 'onError', type: '(err: Error) => void', desc: 'Called on parse or render errors.' },
      ],
      methods: [
        { sig: 'load(source: string | ArrayBuffer): Promise<void>', desc: 'Load from a URL or ArrayBuffer and render the first page.' },
        { sig: 'goToPage(index: number): Promise<void>', desc: 'Render a specific page (0-indexed, clamped).' },
        { sig: 'nextPage(): Promise<void>', desc: 'Advance one page.' },
        { sig: 'prevPage(): Promise<void>', desc: 'Go back one page.' },
        { sig: 'get pageCount(): number', desc: 'Total pages (0 until loaded).' },
        { sig: 'get currentPage(): number', desc: 'Current page index.' },
        { sig: 'get canvasElement(): HTMLCanvasElement', desc: 'The underlying canvas.' },
        { sig: 'destroy(): void', desc: 'Tear down the worker and release resources.' },
      ],
    },
    {
      name: 'DocxDocument',
      ctor: 'await DocxDocument.load(source, options?)',
      note: 'Headless engine — render any page into any canvas you supply.',
      options: [GFONTS, ZIP, MATH, MODE],
      methods: [
        { sig: 'static load(source, options?): Promise<DocxDocument>', desc: 'Parse a document from a URL or ArrayBuffer.' },
        { sig: 'get pageCount(): number', desc: 'Total pages.' },
        { sig: 'renderPage(canvas, index, opts?: { width?, dpr?, showTrackChanges? }): Promise<void>', desc: 'Render one page into the given canvas. Unavailable in `mode: "worker"` — use renderPageToBitmap.' },
        { sig: 'renderPageToBitmap(index, opts?: { width?, dpr?, showTrackChanges? }): Promise<ImageBitmap>', desc: 'Render one page and return it as an ImageBitmap (both modes; in worker mode the render runs off the main thread). Equations are skipped in `mode: "worker"` (they require `mode: "main"`). The bitmap is caller-owned: pass it to `transferFromImageBitmap` (which consumes it) or call `bitmap.close()`.' },
      ],
    },
  ],

  xlsx: [
    {
      name: 'XlsxViewer',
      ctor: 'new XlsxViewer(container: HTMLElement, options?: XlsxViewerOptions)',
      note: 'Full workbook viewer. Takes a container <div> (not a canvas) — it manages its own canvas, sheet-tab bar and zoom slider.',
      options: [
        { name: 'cellScale', type: 'number', def: '1', desc: 'Scale factor for cell/header dimensions (0.5 = half size).' },
        { name: 'showZoomSlider', type: 'boolean', def: 'true', desc: 'Show the Excel-style zoom slider at the end of the tab bar.' },
        { name: 'zoomMin / zoomMax', type: 'number', def: '0.1 / 4', desc: 'Zoom slider bounds as scale factors (10%–400%).' },
        GFONTS,
        ZIP,
        MATH,
        VIEWER_MODE,
        { name: 'onReady', type: '(sheetNames: string[]) => void', desc: 'Called once the workbook is parsed.' },
        { name: 'onSheetChange', type: '(index: number, total: number) => void', desc: 'Called when the active sheet changes; `total` is the sheet count. Read the name via `sheetNames[index]`.' },
        { name: 'onSelectionChange', type: '(sel: CellRange | null) => void', desc: 'Called when the selected range changes; null clears it.' },
        { name: 'onError', type: '(err: Error) => void', desc: 'Called on parse or render errors.' },
      ],
      methods: [
        { sig: 'load(source: string | ArrayBuffer): Promise<void>', desc: 'Load a workbook from a URL or ArrayBuffer and render the first sheet.' },
        { sig: 'goToSheet(index: number): Promise<void>', desc: 'Show a specific sheet (0-indexed, clamped).' },
        { sig: 'nextSheet(): Promise<void>', desc: 'Advance one sheet.' },
        { sig: 'prevSheet(): Promise<void>', desc: 'Go back one sheet.' },
        { sig: 'get sheetIndex(): number', desc: 'Current sheet index.' },
        { sig: 'get sheetCount(): number', desc: 'Total sheets (0 until loaded).' },
        { sig: 'get sheetNames(): string[]', desc: 'Names of all sheets.' },
        { sig: 'get selection(): CellRange | null', desc: 'The current selected range.' },
        { sig: 'getCellAt(clientX: number, clientY: number): CellAddress | null', desc: 'Hit-test a viewport coordinate to a cell address.' },
        { sig: 'get canvasElement(): HTMLCanvasElement', desc: 'The underlying canvas the grid is drawn on.' },
        { sig: 'destroy(): void', desc: 'Tear down the worker and release resources.' },
      ],
    },
    {
      name: 'XlsxWorkbook',
      ctor: 'await XlsxWorkbook.load(source, options?)',
      note: 'Headless engine — parse once, render any sheet viewport into any canvas you supply.',
      options: [GFONTS, ZIP, MATH, MODE],
      methods: [
        { sig: 'static load(source, options?): Promise<XlsxWorkbook>', desc: 'Parse a workbook from a URL or ArrayBuffer.' },
        { sig: 'get sheetNames(): string[]', desc: 'Names of all sheets.' },
        { sig: 'get sheetCount(): number', desc: 'Total sheets.' },
        { sig: 'renderViewport(canvas, sheetIndex, viewport, opts?: { width?, height?, dpr?, cellScale? }): Promise<void>', desc: 'Render a row/col window of a sheet into the given canvas. Equations in shapes render when a `math` engine was passed to `load`. Unavailable in `mode: "worker"` — use renderViewportToBitmap.' },
        { sig: 'renderViewportToBitmap(sheetIndex, viewport, opts: { width, height, dpr?, cellScale? }): Promise<ImageBitmap>', desc: 'Render a sheet viewport and return it as an ImageBitmap (both modes; in worker mode the render runs off the main thread). `width` and `height` are required — a worker has no DOM element to measure. Equations in shapes are skipped in `mode: "worker"` (they require `mode: "main"`). The bitmap is caller-owned: pass it to `transferFromImageBitmap` (which consumes it) or call `bitmap.close()`.' },
        { sig: 'resolveValidationList(sheetIndex, formula1): Promise<ResolvedList>', desc: 'Resolve a list-type data-validation `formula1` (ECMA-376 §18.3.1.32) into the allowed values to display — inline quoted list, a range reference (each cell’s display string), or `{ kind: \'formula\' }` for named ranges. Read-only.' },
        { sig: 'destroy(): void', desc: 'Release the worker.' },
      ],
    },
  ],
};
