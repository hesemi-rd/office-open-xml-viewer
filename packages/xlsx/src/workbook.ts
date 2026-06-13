import InlineWorker from './worker.ts?worker&inline';
import wasmAssetUrl from './wasm/xlsx_parser_bg.wasm?url';
import {
  preloadGoogleFonts,
  WorkerBridge,
  type FontPreloadEntry,
  type LoadOptions as CoreLoadOptions,
  type MathRenderer,
} from '@silurus/ooxml-core';
import type { ParsedWorkbook, Worksheet, ViewportRange, RenderViewportOptions, WorkerResponse, Cell } from './types.js';
import { renderWorksheetViewport } from './render-orchestrator.js';
import { formatCellValue } from './number-format.js';
import {
  parseListFormula,
  resolveListValues,
  type ResolvedList,
} from './validation-list.js';

/** Office font name → metric-compatible Google Fonts substitute. These are
 *  the well-known pairings Microsoft and Google both publish and ship on
 *  Linux distributions: Calibri → Carlito, Cambria → Caladea (same advance
 *  widths and ascender / descender). Loading the substitute on a system
 *  that lacks the Office face keeps text width measurements close to
 *  Excel's. The substitute font-family differs from the requested name, so
 *  `loadFamily` redirects FontFaceSet loading appropriately. */
const NOTO_NASKH_ARABIC_URL =
  'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap';
const NOTO_SANS_ARABIC_URL =
  'https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;700&display=swap';

const XLSX_GOOGLE_FONTS: Record<string, FontPreloadEntry> = {
  'calibri': {
    url: 'https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&display=swap',
    loadFamily: 'Carlito',
  },
  'cambria': {
    url: 'https://fonts.googleapis.com/css2?family=Caladea:ital,wght@0,400;0,700;1,400;1,700&display=swap',
    loadFamily: 'Caladea',
  },
  // Common Arabic-script faces that hosts rarely ship. Map them to Noto
  // substitutes so RTL workbooks (e.g. the LibreOffice-authored sample-29,
  // which requests Sakkal Majalla / Univers Next Arabic) render with a real
  // web font instead of an oversized OS fallback. "Naskh" covers traditional
  // serif-like Arabic faces; "Sans" covers the modern geometric ones.
  'sakkal majalla': { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'traditional arabic': { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'simplified arabic': { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'arabic typesetting': { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'univers next arabic': { url: NOTO_SANS_ARABIC_URL, loadFamily: 'Noto Sans Arabic' },
  // Self-referencing entries so the generic Arabic fallback fonts (appended to
  // the renderer's font chain) are themselves loaded whenever useGoogleFonts
  // is enabled — see `_load`, which always queues these names.
  'noto naskh arabic': { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'noto sans arabic': { url: NOTO_SANS_ARABIC_URL, loadFamily: 'Noto Sans Arabic' },
};

/** Options for {@link XlsxWorkbook.load}. The shared load-options type from
 *  `@silurus/ooxml-core` (`useGoogleFonts`, `maxZipEntryBytes`). */
export type LoadOptions = CoreLoadOptions;

export class XlsxWorkbook {
  private worker: Worker;
  private bridge: WorkerBridge<WorkerResponse>;
  private parsedWorkbook: ParsedWorkbook | null = null;
  private sheetCache = new Map<number, Worksheet>();
  /** Cache of decoded image bitmaps keyed by their data URL. Shared across sheets. */
  private imageCache = new Map<string, CanvasImageSource>();
  private rawData: ArrayBuffer | null = null;
  private maxZipEntryBytes: number | undefined;
  /** Opt-in OMML equation engine, injected once at {@link load}. Every
   *  `renderViewport` call reuses it — equations in shapes render when present,
   *  and are skipped (engine tree-shaken) when omitted. */
  private math: MathRenderer | undefined;

  private constructor() {
    this.worker = new InlineWorker();
    this.bridge = new WorkerBridge<WorkerResponse>(this.worker, {
      correlate: (res) => res.id,
      toError: (res) => (res.type === 'error' ? res.message : undefined),
    });
    const wasmUrl = new URL(wasmAssetUrl, location.href).href;
    this.bridge.post({ type: 'init', wasmUrl });
  }

  /** Parse an XLSX from a URL or ArrayBuffer. */
  static async load(source: string | ArrayBuffer, opts: LoadOptions = {}): Promise<XlsxWorkbook> {
    const wb = new XlsxWorkbook();
    await wb._load(source, opts);
    return wb;
  }

  private async _load(source: string | ArrayBuffer, opts: LoadOptions = {}): Promise<void> {
    let data: ArrayBuffer;
    if (typeof source === 'string') {
      const res = await fetch(source);
      if (!res.ok) throw new Error(`Failed to fetch: ${res.status} ${res.statusText}`);
      data = await res.arrayBuffer();
    } else {
      data = source;
    }
    this.rawData = data;
    this.maxZipEntryBytes = opts.maxZipEntryBytes;
    this.math = opts.math;
    const parsed = await this.bridge.request((id) => ({
      type: 'parse',
      id,
      data: data.slice(0),
      maxZipEntryBytes: this.maxZipEntryBytes,
    }));
    this.parsedWorkbook = (parsed as Extract<WorkerResponse, { type: 'parsed' }>).workbook;
    if (opts.useGoogleFonts) {
      // Walk every styled font in the workbook and queue Google Fonts
      // substitutes for any Office faces (Calibri → Carlito, Cambria →
      // Caladea). Documents that use only system fonts produce zero
      // network requests.
      const names = new Set<string>();
      for (const f of this.parsedWorkbook.styles?.fonts ?? []) {
        if (f.name) names.add(f.name);
      }
      // Always load the generic Arabic fallbacks so any Arabic-script cell
      // gets a real web font even when its named family is unmapped (the
      // renderer's DEFAULT_FONT_FAMILY chain ends with these two Noto faces).
      names.add('Noto Naskh Arabic');
      names.add('Noto Sans Arabic');
      await preloadGoogleFonts(names, XLSX_GOOGLE_FONTS);
    }
  }

  get sheetNames(): string[] {
    return this.parsedWorkbook?.workbook.sheets.map((s) => s.name) ?? [];
  }

  get sheetCount(): number {
    return this.parsedWorkbook?.workbook.sheets.length ?? 0;
  }

  /** Per-sheet tab colors (`#RRGGBB`) parallel to {@link sheetNames}.
   *  `null` for sheets that declare no tab color. */
  get tabColors(): (string | null)[] {
    return this.parsedWorkbook?.workbook.sheets.map((s) => s.tabColor ?? null) ?? [];
  }

  async getWorksheet(sheetIndex: number): Promise<Worksheet> {
    const cached = this.sheetCache.get(sheetIndex);
    if (cached) return cached;
    if (!this.parsedWorkbook || !this.rawData) {
      throw new Error('Workbook not loaded');
    }
    const rawData = this.rawData;
    const sheetMeta = this.parsedWorkbook.workbook.sheets[sheetIndex];
    if (!sheetMeta) throw new Error(`Sheet index ${sheetIndex} out of range`);

    const res = await this.bridge.request((id) => ({
      type: 'parseSheet',
      id,
      data: rawData.slice(0),
      sheetIndex,
      sheetName: sheetMeta.name,
      maxZipEntryBytes: this.maxZipEntryBytes,
    }));
    const ws = (res as Extract<WorkerResponse, { type: 'parsedSheet' }>).worksheet;
    this.sheetCache.set(sheetIndex, ws);
    return ws;
  }

  /**
   * Resolve a `list`-type data-validation `formula1` (ECMA-376 §18.3.1.32) into
   * the set of allowed values to display, evaluated relative to `sheetIndex`
   * (the sheet that owns the validation, used to resolve unqualified ranges):
   *
   * - Inline quoted list `"A,B,C"`        → the literal values.
   * - Range ref `$B$2:$B$5`               → each non-empty cell's *display
   *   string* (the same formatted text the grid shows, via {@link formatCellValue}),
   *   walked row-major. `Sheet2!$A$1:$A$9` resolves against the named sheet
   *   (lazily parsed via {@link getWorksheet}, hence async).
   * - Named range / complex formula       → `{ kind: 'formula' }` carrying the
   *   raw text so the caller can disclose it rather than blanking it.
   *
   * Read-only: this only reads cell values for display; it never writes.
   */
  async resolveValidationList(
    sheetIndex: number,
    formula1: string | undefined,
  ): Promise<ResolvedList> {
    if (!this.parsedWorkbook) throw new Error('Workbook not loaded');
    const parsed = parseListFormula(formula1);
    if (parsed.kind !== 'range') {
      // Inline / unresolved need no cell lookup.
      return resolveListValues(parsed, () => null);
    }

    // Pick the target sheet: the qualifier name (case-insensitive) or, when the
    // range is unqualified, the sheet that owns the validation.
    let targetIndex = sheetIndex;
    if (parsed.sheet) {
      const names = this.sheetNames;
      const found = names.findIndex(
        (n) => n.toLowerCase() === parsed.sheet?.toLowerCase(),
      );
      // Unknown sheet name (e.g. an external reference) → cannot expand;
      // surface the formula instead of silently dropping it.
      if (found < 0) return { kind: 'formula', formula: formula1 ?? '' };
      targetIndex = found;
    }

    const ws = await this.getWorksheet(targetIndex);
    const styles = this.parsedWorkbook.styles;
    // Index the target sheet's cells by "row:col" for O(1) lookup during the
    // row-major walk in resolveListValues.
    const byRC = new Map<string, Cell>();
    for (const r of ws.rows) {
      for (const c of r.cells) byRC.set(`${c.row}:${c.col}`, c);
    }

    return resolveListValues(parsed, (row, col) => {
      const cell = byRC.get(`${row}:${col}`);
      if (!cell) return null;
      return formatCellValue(cell, styles);
    });
  }

  async renderViewport(
    target: HTMLCanvasElement | OffscreenCanvas,
    sheetIndex: number,
    viewport: ViewportRange,
    opts: RenderViewportOptions = {},
  ): Promise<void> {
    if (!this.parsedWorkbook) throw new Error('Workbook not loaded');
    // Hot path: during scroll the worksheet is already cached. Skip the await
    // to keep the whole render in a single synchronous task so the browser
    // doesn't paint between the canvas clear and the draw.
    const ws = this.sheetCache.get(sheetIndex) ?? await this.getWorksheet(sheetIndex);
    return renderWorksheetViewport(
      { ws, styles: this.parsedWorkbook.styles, imageCache: this.imageCache, math: this.math },
      target,
      viewport,
      opts,
    );
  }

  destroy(): void {
    this.bridge.terminate();
  }
}
