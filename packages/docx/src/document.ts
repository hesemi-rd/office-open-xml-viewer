import InlineWorker from './worker.ts?worker&inline';
import wasmAssetUrl from './wasm/docx_parser_bg.wasm?url';
import {
  preloadGoogleFonts,
  WorkerBridge,
  type FontPreloadEntry,
  type LoadOptions as CoreLoadOptions,
  type MathRenderer,
} from '@silurus/ooxml-core';
import type { PaginatedBodyElement, DocxDocumentModel, RenderPageOptions, WorkerResponse, DocComment, DocNote } from './types';
import { computePages, renderDocumentToCanvas, documentHasMath, prepareMathRuns, resolveKinsokuRules } from './renderer';

/** Theme-referenced typefaces commonly used by DOCX templates. Mirrors the
 *  PPTX map — these are the well-known free webfont alternatives Microsoft
 *  Office templates pull from. Substitutes that diverge from the requested
 *  family name (Calibri → Carlito, Cambria → Caladea) include
 *  `loadFamily` so the FontFaceSet load is driven against the substitute. */
const NOTO_NASKH_ARABIC_URL =
  'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap';
const NOTO_SANS_ARABIC_URL =
  'https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;700&display=swap';

const DOCX_GOOGLE_FONTS: Record<string, FontPreloadEntry> = {
  'calibri':           { url: 'https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Carlito' },
  'cambria':           { url: 'https://fonts.googleapis.com/css2?family=Caladea:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Caladea' },
  'nunito sans':       { url: 'https://fonts.googleapis.com/css2?family=Nunito+Sans:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'nunito':            { url: 'https://fonts.googleapis.com/css2?family=Nunito:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'open sans':         { url: 'https://fonts.googleapis.com/css2?family=Open+Sans:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'roboto':            { url: 'https://fonts.googleapis.com/css2?family=Roboto:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'lato':              { url: 'https://fonts.googleapis.com/css2?family=Lato:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'montserrat':        { url: 'https://fonts.googleapis.com/css2?family=Montserrat:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'poppins':           { url: 'https://fonts.googleapis.com/css2?family=Poppins:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'raleway':           { url: 'https://fonts.googleapis.com/css2?family=Raleway:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'playfair display':  { url: 'https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  // Common Arabic-script faces that hosts rarely ship. Map them to Noto
  // substitutes so RTL documents (e.g. sample-7, which requests Sakkal Majalla
  // / Univers Next Arabic) render with a real web font instead of an oversized
  // OS fallback. "Naskh" covers traditional serif-like Arabic faces; "Sans"
  // covers the modern geometric ones.
  'sakkal majalla':      { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'traditional arabic':  { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'simplified arabic':   { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'arabic typesetting':  { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'univers next arabic': { url: NOTO_SANS_ARABIC_URL, loadFamily: 'Noto Sans Arabic' },
  // Self-referencing entries so the generic Arabic fallback fonts (appended to
  // the renderer's font chain) are themselves loaded whenever useGoogleFonts
  // is enabled — see `load`, which always queues these names.
  'noto naskh arabic':   { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'noto sans arabic':    { url: NOTO_SANS_ARABIC_URL, loadFamily: 'Noto Sans Arabic' },
};

/** Options for {@link DocxDocument.load}. Extends the shared load-options type
 *  from `@silurus/ooxml-core` (`useGoogleFonts`, `maxZipEntryBytes`) with the
 *  opt-in math engine. */
export interface LoadOptions extends CoreLoadOptions {
  /**
   * Opt-in OMML equation engine. Import it from the separate `@silurus/ooxml/math`
   * entry and pass it in: `import { math } from '@silurus/ooxml/math'`. When
   * omitted, equations are skipped and the ~3 MB engine never enters the bundle.
   */
  math?: MathRenderer;
}

export class DocxDocument {
  private _document: DocxDocumentModel | null = null;
  private _pages: PaginatedBodyElement[][] | null = null;
  private _worker: Worker;
  private _bridge: WorkerBridge<WorkerResponse>;

  private constructor() {
    this._worker = new InlineWorker();
    this._bridge = new WorkerBridge<WorkerResponse>(this._worker, {
      correlate: (res) => res.id,
      toError: (res) => (res.type === 'error' ? res.message : undefined),
    });
    const wasmUrl = new URL(wasmAssetUrl, location.href).href;
    this._bridge.post({ type: 'init', wasmUrl });
  }

  static async load(source: string | ArrayBuffer, opts: LoadOptions = {}): Promise<DocxDocument> {
    const doc = new DocxDocument();
    let buffer: ArrayBuffer;
    if (typeof source === 'string') {
      const res = await fetch(source);
      if (!res.ok) throw new Error(`Failed to fetch: ${res.status} ${res.statusText}`);
      buffer = await res.arrayBuffer();
    } else {
      buffer = source;
    }
    await doc._parse(buffer, opts.maxZipEntryBytes);
    if (opts.useGoogleFonts) {
      // Always load the generic Arabic fallbacks so any Arabic-script run gets
      // a real web font even when its named family is unmapped (the renderer's
      // font fallback chains end with these two Noto faces).
      await preloadGoogleFonts(
        [
          doc._document?.majorFont,
          doc._document?.minorFont,
          'Noto Naskh Arabic',
          'Noto Sans Arabic',
        ],
        DOCX_GOOGLE_FONTS,
      );
    }
    // Equations are converted + rasterized before pagination (which reads their
    // extents synchronously). Requires the opt-in `math` engine; without it,
    // equations are skipped (and the engine asset is never bundled).
    if (opts.math && doc._document && documentHasMath(doc._document.body)) {
      await prepareMathRuns(doc._document.body, opts.math);
    }
    return doc;
  }

  private async _parse(buffer: ArrayBuffer, maxZipEntryBytes?: number): Promise<void> {
    const res = await this._bridge.request(
      (id) => ({ type: 'parse', id, data: buffer, maxZipEntryBytes }),
      [buffer],
    );
    this._document = (res as Extract<WorkerResponse, { type: 'parsed' }>).document;
  }

  destroy(): void {
    this._bridge.terminate();
  }

  get pageCount(): number {
    if (!this._document) return 0;
    return this._getPages().length;
  }

  get document(): DocxDocumentModel {
    if (!this._document) throw new Error('Document not loaded');
    return this._document;
  }

  /**
   * ECMA-376 §17.13.4 — the document's comments (`word/comments.xml`), each with
   * id / author / initials / date / plain-text body. Comments are a data-only
   * API: they are NOT drawn on the page (Word renders them in a margin pane /
   * balloons, which this viewer does not reproduce). Use this to build a review
   * panel, export an annotation list, etc. Returns `[]` when the document has no
   * comments part. The same data is also reachable via `document.comments`.
   */
  get comments(): DocComment[] {
    return this._document?.comments ?? [];
  }

  /**
   * ECMA-376 §17.11.10 — the document's footnotes (`word/footnotes.xml`),
   * excluding the reserved separator entries. Each note carries its `id` and
   * block-level `content`; use {@link noteText} for the plain-text body. These
   * ARE drawn at the bottom of the page that holds their reference; this getter
   * additionally exposes them as data. Returns `[]` when absent.
   */
  get footnotes(): DocNote[] {
    return this._document?.footnotes ?? [];
  }

  /**
   * ECMA-376 §17.11.4 — the document's endnotes (`word/endnotes.xml`). Same
   * shape as {@link footnotes}; rendered at the end of the document. Returns
   * `[]` when absent.
   */
  get endnotes(): DocNote[] {
    return this._document?.endnotes ?? [];
  }

  private _getPages(): PaginatedBodyElement[][] {
    if (this._pages) return this._pages;
    if (!this._document) return [];
    const measure = new OffscreenCanvas(1, 1);
    const ctx = measure.getContext('2d');
    if (!ctx) {
      this._pages = [this._document.body];
      return this._pages;
    }
    // Pagination must use the same fontFamilyClasses + kinsoku rules as the
    // render path, otherwise line-break decisions (and thus page breaks)
    // diverge between measurement and paint. ECMA-376 §17.15.1.58–.60.
    this._pages = computePages(
      this._document.body,
      this._document.section,
      ctx,
      this._document.fontFamilyClasses ?? {},
      resolveKinsokuRules(this._document.settings),
      this._document.footnotes ?? [],
    );
    return this._pages;
  }

  renderPage(
    target: HTMLCanvasElement | OffscreenCanvas,
    pageIndex: number,
    opts: RenderPageOptions = {},
  ): Promise<void> {
    if (!this._document) throw new Error('Document not loaded');
    const pages = this._getPages();
    return renderDocumentToCanvas(this._document, target, pageIndex, {
      ...opts,
      totalPages: pages.length,
      prebuiltPages: pages,
    });
  }
}
