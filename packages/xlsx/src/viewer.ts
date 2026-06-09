import { XlsxWorkbook } from './workbook.js';
import type { ViewportRange, Worksheet } from './types.js';
import type { MathRenderer } from '@silurus/ooxml-core';
import { HEADER_W, HEADER_H, colWidthToPx, rowHeightToPx, getMdwForWorksheet } from './renderer.js';

const TAB_BAR_H = 30;
// Gap between adjacent sheet tabs. The first tab also gets this much leading
// space so it is offset from the row-header boundary by the same margin that
// separates tabs from each other.
const TAB_GAP = 1;

export interface XlsxViewerOptions {
  /** Scale factor for cell/header dimensions (default 1). 0.5 = half size. */
  cellScale?: number;
  /** Show the Excel-style zoom slider at the right end of the sheet-tab bar.
   *  Default `true`. Set `false` to hide it (e.g. when the host supplies its
   *  own zoom control). */
  showZoomSlider?: boolean;
  /** Lower/upper bounds for the zoom slider as scale factors. Default 0.1–4
   *  (10%–400%, matching Excel's zoom range). */
  zoomMin?: number;
  zoomMax?: number;
  onReady?: (sheetNames: string[]) => void;
  onSheetChange?: (index: number, name: string) => void;
  onError?: (err: Error) => void;
  /** Called when the selected cell range changes. null means no selection. */
  onSelectionChange?: (selection: CellRange | null) => void;
  /**
   * Opt in to Google-Fonts-hosted, metric-compatible substitutes for the
   * Office default fonts (Carlito for Calibri, Caladea for Cambria) so
   * column layouts match Excel on systems without Office installed.
   * Default `false`. See `XlsxWorkbook.LoadOptions.useGoogleFonts` for the
   * privacy implications.
   */
  useGoogleFonts?: boolean;
  /**
   * Override the per-entry ZIP decompression cap (bytes) used by the
   * zip-bomb guard in the Rust parser. Defaults to 512 MiB. Zero / negative
   * values fall back to the default.
   */
  maxZipEntryBytes?: number;
  /**
   * Opt-in OMML equation engine for rendering math in shapes/text boxes.
   * Import it from the separate `@silurus/ooxml/math` entry and pass it in
   * (`import { math } from '@silurus/ooxml/math'`). When omitted, equations are
   * skipped and the ~3 MB engine never enters the bundle (tree-shaken). Same
   * dependency-injection contract as the docx/pptx viewers.
   */
  math?: MathRenderer;
}

export interface CellAddress {
  row: number;
  col: number;
}

export type SelectionMode = 'cells' | 'rows' | 'cols' | 'all';

export interface CellRange {
  anchor: CellAddress;
  active: CellAddress;
  mode: SelectionMode;
}

/**
 * Cumulative-offset axis (columns or rows) with O(log n) lookup instead of the
 * O(n) linear scan that previously walked from the first cell to the scroll
 * position on every frame / click (up to ~1M rows). Sizes are sparse
 * (`Record<index, size>`); most cells use the default, so the prefix sum is
 * built only over the custom entries and a binary search resolves both
 * directions. All offsets are in *logical* pixels.
 */
class AxisMetrics {
  private readonly idxs: number[];      // sorted custom (1-based) indices
  private readonly cumDelta: number[];  // prefix sum of (customPx - defaultPx)
  constructor(
    customs: Record<number, number>,
    private readonly defaultPx: number,
    toPx: (raw: number) => number,
    private readonly maxIndex: number,
  ) {
    this.idxs = Object.keys(customs)
      .map(Number)
      .filter((n) => n >= 1 && n <= maxIndex)
      .sort((a, b) => a - b);
    this.cumDelta = new Array(this.idxs.length);
    let acc = 0;
    for (let k = 0; k < this.idxs.length; k++) {
      acc += toPx(customs[this.idxs[k]]) - defaultPx;
      this.cumDelta[k] = acc;
    }
  }

  /** Σ (customPx − defaultPx) for custom indices strictly below `index`. */
  private deltaBefore(index: number): number {
    let lo = 0, hi = this.idxs.length;
    while (lo < hi) {
      const m = (lo + hi) >> 1;
      if (this.idxs[m] < index) lo = m + 1; else hi = m;
    }
    return lo === 0 ? 0 : this.cumDelta[lo - 1];
  }

  /** Logical-px offset to the START of `index` (1-based). */
  offsetOf(index: number): number {
    return (index - 1) * this.defaultPx + this.deltaBefore(index);
  }

  /** Index whose span contains absolute logical-px `offset`, plus the partial
   *  scroll into it. Mirrors the old linear search exactly. */
  indexAt(offset: number): { index: number; partial: number } {
    if (offset <= 0) return { index: 1, partial: 0 };
    let lo = 1, hi = this.maxIndex;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.offsetOf(mid) <= offset) lo = mid; else hi = mid - 1;
    }
    return { index: lo, partial: offset - this.offsetOf(lo) };
  }
}

interface SheetAxes { col: AxisMetrics; row: AxisMetrics; }
const sheetAxisCache = new WeakMap<Worksheet, SheetAxes>();

/** Per-Worksheet column/row axis metrics (memoized; the workbook keeps one
 *  Worksheet object per sheet so this hits across frames). */
function getSheetAxes(ws: Worksheet, mdw: number): SheetAxes {
  const cached = sheetAxisCache.get(ws);
  if (cached) return cached;
  const axes: SheetAxes = {
    col: new AxisMetrics(ws.colWidths, colWidthToPx(ws.defaultColWidth, mdw), (raw) => colWidthToPx(raw, mdw), 16384),
    row: new AxisMetrics(ws.rowHeights, rowHeightToPx(ws.defaultRowHeight), (raw) => rowHeightToPx(raw), 1048576),
  };
  sheetAxisCache.set(ws, axes);
  return axes;
}

export class XlsxViewer {
  private wb: XlsxWorkbook | null = null;
  private canvas: HTMLCanvasElement;
  private canvasArea: HTMLDivElement;
  private scrollHost: HTMLDivElement;
  private spacer: HTMLDivElement;
  private tabBar: HTMLDivElement;
  private tabStrip: HTMLDivElement;
  private navPrev: HTMLButtonElement;
  private navNext: HTMLButtonElement;
  private navGroup!: HTMLDivElement;
  private tabs: HTMLButtonElement[] = [];
  /** Per-tab colors parallel to `tabs`, from `<sheetPr><tabColor>`. */
  private tabColors: (string | null)[] = [];
  private zoomSlider: HTMLInputElement | null = null;
  private zoomLabel: HTMLSpanElement | null = null;
  private currentSheet = 0;
  private currentWorksheet: Worksheet | null = null;
  private opts: XlsxViewerOptions;
  private resizeObserver: ResizeObserver | null = null;

  // Selection state
  private anchorCell: CellAddress | null = null;
  private activeCell: CellAddress | null = null;
  private selectionMode: SelectionMode = 'cells';
  private isSelecting = false;
  private selectionOverlay: HTMLDivElement;
  private keydownHandler: ((e: KeyboardEvent) => void) | null = null;
  // Pending touch/pen tap: only commit selection on pointerup if movement stays within tap threshold,
  // so swipe-to-scroll on mobile doesn't change the selected cell.
  private pendingTap: { x: number; y: number; shiftKey: boolean; pointerId: number } | null = null;

  constructor(container: HTMLElement, opts: XlsxViewerOptions = {}) {
    this.opts = opts;

    const wrapper = document.createElement('div');
    wrapper.style.cssText =
      `position:relative;width:100%;height:100%;` +
      `border:1px solid #c8ccd0;background:#fff;box-sizing:border-box;font-family:sans-serif;display:flex;flex-direction:column;`;

    this.canvasArea = document.createElement('div');
    this.canvasArea.style.cssText = `position:relative;flex:1;min-height:0;overflow:hidden;`;

    this.canvas = document.createElement('canvas');
    this.canvas.style.cssText = `position:absolute;top:0;left:0;z-index:0;display:block;`;

    // Selection overlay: sits above canvas, below scrollHost (z-index 0.5 via fractional z not possible,
    // use pointer-events:none so scrollHost still receives events)
    this.selectionOverlay = document.createElement('div');
    this.selectionOverlay.style.cssText =
      `position:absolute;top:0;left:0;z-index:1;pointer-events:none;overflow:hidden;width:100%;height:100%;`;

    this.scrollHost = document.createElement('div');
    this.scrollHost.style.cssText = `position:absolute;inset:0;overflow:auto;z-index:2;background:transparent;`;

    this.spacer = document.createElement('div');
    this.spacer.style.cssText = `position:absolute;top:0;left:0;pointer-events:none;`;
    this.scrollHost.appendChild(this.spacer);

    this.canvasArea.appendChild(this.canvas);
    this.canvasArea.appendChild(this.selectionOverlay);
    this.canvasArea.appendChild(this.scrollHost);

    const headerW = Math.round(HEADER_W * (this.opts.cellScale ?? 1));

    this.tabBar = document.createElement('div');
    this.tabBar.style.cssText =
      `display:flex;align-items:flex-end;height:${TAB_BAR_H}px;flex-shrink:0;` +
      `background:#f0f0f0;border-top:1px solid #c8ccd0;`;

    // Excel-style scroll buttons. They scroll the tab strip; they do NOT change
    // the active sheet. Disabled (greyed) at the ends / when there is no overflow.
    this.navPrev = this.makeNavButton('◀', 'Scroll tabs left', () => this.scrollTabs(-1));
    this.navNext = this.makeNavButton('▶', 'Scroll tabs right', () => this.scrollTabs(1));
    this.navPrev.dataset.xlsxTabNav = 'prev';
    this.navNext.dataset.xlsxTabNav = 'next';

    // The two buttons together span the row-header width so the tab strip starts
    // at the same x as the data columns (which begin after the HEADER_W header).
    const navGroup = document.createElement('div');
    navGroup.style.cssText =
      `display:flex;flex-shrink:0;width:${headerW}px;height:100%;`;
    navGroup.appendChild(this.navPrev);
    navGroup.appendChild(this.navNext);
    this.navGroup = navGroup;

    // The scrollable strip that actually holds the sheet tabs. position:relative
    // so each tab's offsetLeft is measured against the strip's scroll content.
    this.tabStrip = document.createElement('div');
    // margin-left (not padding) keeps the leading gap OUTSIDE the scroll content
    // so each tab's offsetLeft / scrollLeft math is unaffected and scrolling
    // still returns to exactly 0.
    this.tabStrip.style.cssText =
      `position:relative;display:flex;align-items:flex-end;flex:1;min-width:0;height:100%;` +
      `margin-left:${TAB_GAP}px;overflow-x:auto;overflow-y:hidden;gap:${TAB_GAP}px;scrollbar-width:none;`;
    this.tabStrip.classList.add('xlsx-tab-strip');
    const style = document.createElement('style');
    style.textContent =
      `.xlsx-tab-strip::-webkit-scrollbar{display:none}` +
      `.xlsx-tab-nav{background:transparent;transition:background 0.1s;}` +
      `.xlsx-tab-nav:hover{background:rgba(0,0,0,0.08);}` +
      // Excel-status-bar zoom slider: a thin uniform gray track (no colored
      // fill on either side of the thumb) with a small round gray handle.
      `.xlsx-zoom-slider{-webkit-appearance:none;appearance:none;background:transparent;height:15px;margin:0;}` +
      `.xlsx-zoom-slider::-webkit-slider-runnable-track{height:4px;background:#c4c4c4;border-radius:2px;}` +
      `.xlsx-zoom-slider::-webkit-slider-thumb{-webkit-appearance:none;appearance:none;width:12px;height:12px;margin-top:-4px;border-radius:50%;background:#808080;cursor:pointer;}` +
      `.xlsx-zoom-slider:hover::-webkit-slider-thumb{background:#5f5f5f;}` +
      `.xlsx-zoom-slider::-moz-range-track{height:4px;background:#c4c4c4;border-radius:2px;}` +
      `.xlsx-zoom-slider::-moz-range-thumb{width:12px;height:12px;border:none;border-radius:50%;background:#808080;cursor:pointer;}`;
    document.head.appendChild(style);
    this.tabStrip.addEventListener('scroll', () => this.updateNavButtons());

    this.tabBar.appendChild(navGroup);
    this.tabBar.appendChild(this.tabStrip);
    if (this.opts.showZoomSlider !== false) {
      this.tabBar.appendChild(this.buildZoomControl());
    }

    wrapper.appendChild(this.canvasArea);
    wrapper.appendChild(this.tabBar);
    container.appendChild(wrapper);

    this.scrollHost.addEventListener('scroll', () => {
      this.renderCurrentSheet();
      this.updateSelectionOverlay();
    });

    // Re-render whenever the canvas area changes size
    this.resizeObserver = new ResizeObserver(() => {
      this.renderCurrentSheet();
      this.updateSelectionOverlay();
      this.updateNavButtons();
    });
    this.resizeObserver.observe(this.canvasArea);

    this.setupSelectionEvents();
  }

  /**
   * Load an XLSX from URL or ArrayBuffer and render the first sheet.
   *
   * Error contract (shared by all three viewers): on failure, if an `onError`
   * callback was provided it is invoked and `load` resolves normally; if not,
   * the error is rethrown so it is never silently swallowed.
   */
  async load(source: string | ArrayBuffer): Promise<void> {
    try {
      this.wb = await XlsxWorkbook.load(source, {
        useGoogleFonts: this.opts.useGoogleFonts,
        maxZipEntryBytes: this.opts.maxZipEntryBytes,
        math: this.opts.math,
      });
      this.buildTabs();
      this.opts.onReady?.(this.wb.sheetNames);
      await this.showSheet(0);
    } catch (err) {
      const e = err instanceof Error ? err : new Error(String(err));
      if (this.opts.onError) {
        this.opts.onError(e);
        return;
      }
      throw e;
    }
  }

  /** The loaded workbook, or throws if {@link load} has not completed. */
  private get workbook(): XlsxWorkbook {
    if (!this.wb) throw new Error('Workbook not loaded');
    return this.wb;
  }

  async showSheet(index: number): Promise<void> {
    this.currentSheet = index;
    this.scrollHost.scrollLeft = 0;
    this.scrollHost.scrollTop = 0;
    this.anchorCell = null;
    this.activeCell = null;
    this.selectionMode = 'cells';
    this.updateSelectionOverlay();
    this.updateTabActive(index);
    this.currentWorksheet = await this.workbook.getWorksheet(index);
    this.updateSpacerSize(this.currentWorksheet);
    await this.renderCurrentSheet();
    this.opts.onSheetChange?.(index, this.workbook.sheetNames[index] ?? '');
  }

  /** 0-based index of the currently displayed sheet. */
  get sheetIndex(): number {
    return this.currentSheet;
  }

  /** Total number of sheets in the loaded workbook. */
  get sheetCount(): number {
    return this.wb?.sheetCount ?? 0;
  }

  /**
   * Navigate to a sheet by index, clamped to range. Canonical navigation verb
   * matching {@link PptxViewer.goToSlide} / {@link DocxViewer.goToPage};
   * {@link showSheet} is the lower-level form that assumes a valid index.
   */
  async goToSheet(index: number): Promise<void> {
    if (this.sheetCount === 0) return;
    await this.showSheet(Math.max(0, Math.min(index, this.sheetCount - 1)));
  }

  async nextSheet(): Promise<void> {
    await this.goToSheet(this.currentSheet + 1);
  }

  async prevSheet(): Promise<void> {
    await this.goToSheet(this.currentSheet - 1);
  }

  /** Returns the cell at canvas-client coordinates, or null if outside the cell grid. */
  getCellAt(clientX: number, clientY: number): CellAddress | null {
    const ws = this.currentWorksheet;
    if (!ws) return null;
    const cs = this.opts.cellScale ?? 1;

    const rect = this.canvasArea.getBoundingClientRect();
    const lx = (clientX - rect.left) / cs;
    const ly = (clientY - rect.top) / cs;

    if (lx < HEADER_W || ly < HEADER_H) return null;

    const innerX = lx - HEADER_W;
    const innerY = ly - HEADER_H;

    const freezeRows = ws.freezeRows ?? 0;
    const freezeCols = ws.freezeCols ?? 0;

    // Compute frozen pixel dimensions (unscaled)
    let frozenH = 0;
    const frozenRowH: number[] = [];
    for (let r = 1; r <= freezeRows; r++) {
      const h = rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight);
      frozenRowH.push(h);
      frozenH += h;
    }
    let frozenW = 0;
    const frozenColW: number[] = [];
    for (let c = 1; c <= freezeCols; c++) {
      const w = colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth, getMdwForWorksheet(ws));
      frozenColW.push(w);
      frozenW += w;
    }

    // Find row
    let row: number;
    if (innerY < frozenH) {
      row = -1;
      let acc = 0;
      for (let r = 0; r < freezeRows; r++) {
        acc += frozenRowH[r];
        if (innerY < acc) { row = r + 1; break; }
      }
      if (row === -1) return null;
    } else {
      const contentY = innerY - frozenH + this.scrollHost.scrollTop / cs;
      row = -1;
      let acc = 0;
      for (let r = freezeRows + 1; r <= 1048576; r++) {
        acc += rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight);
        if (contentY < acc) { row = r; break; }
      }
      if (row === -1) return null;
    }

    // Find col
    let col: number;
    if (innerX < frozenW) {
      col = -1;
      let acc = 0;
      for (let c = 0; c < freezeCols; c++) {
        acc += frozenColW[c];
        if (innerX < acc) { col = c + 1; break; }
      }
      if (col === -1) return null;
    } else {
      const contentX = innerX - frozenW + this.scrollHost.scrollLeft / cs;
      col = -1;
      let acc = 0;
      for (let c = freezeCols + 1; c <= 16384; c++) {
        acc += colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth, getMdwForWorksheet(ws));
        if (contentX < acc) { col = c; break; }
      }
      if (col === -1) return null;
    }

    return { row, col };
  }

  /** Returns the CSS-pixel rect of a cell within canvasArea, or null if not
   *  computable. Mirrors the renderer's per-cell rounding (Math.round(px * cs))
   *  so the selection overlay sits exactly on the canvas's drawn cell borders;
   *  multiplying logical accumulators by `cs` once at the end (the previous
   *  approach) drifted by up to 1 px per cell at non-integer scales.
   */
  private getCellRect(row: number, col: number): { x: number; y: number; w: number; h: number } | null {
    const ws = this.currentWorksheet;
    if (!ws) return null;
    const cs = this.opts.cellScale ?? 1;
    const mdw = getMdwForWorksheet(ws);
    const sp = (px: number) => Math.round(px * cs);
    const colW = (c: number) => sp(colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth, mdw));
    const rowH = (r: number) => sp(rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight));

    const freezeRows = ws.freezeRows ?? 0;
    const freezeCols = ws.freezeCols ?? 0;

    // Compute x. The renderer draws the scrollable area starting at
    // `scrollAreaX = sp(HEADER_W) + Σ sp(frozenWidth)` and offsets each
    // cell by `-(offsetX * cs)` where offsetX is the *logical* partial
    // visibility of the leftmost visible column.
    let x: number;
    if (col <= freezeCols) {
      let acc = sp(HEADER_W);
      for (let c = 1; c < col; c++) acc += colW(c);
      x = acc;
    } else {
      let frozenW = 0;
      for (let c = 1; c <= freezeCols; c++) frozenW += colW(c);
      const scrollAreaX = sp(HEADER_W) + frozenW;

      // Mirror renderCurrentSheet's startCol / offsetX search (binary search).
      const logicalScrollX = this.scrollHost.scrollLeft / cs;
      const colAxis = getSheetAxes(ws, mdw).col;
      const { index: startCol, partial: offsetX } =
        colAxis.indexAt(logicalScrollX + colAxis.offsetOf(freezeCols + 1));

      let acc = scrollAreaX - offsetX * cs;
      if (col >= startCol) {
        for (let c = startCol; c < col; c++) acc += colW(c);
      } else {
        // Cell scrolled off to the left of startCol — subtract instead.
        for (let c = col; c < startCol; c++) acc -= colW(c);
      }
      x = acc;
    }

    // Compute y, same logic as x.
    let y: number;
    if (row <= freezeRows) {
      let acc = sp(HEADER_H);
      for (let r = 1; r < row; r++) acc += rowH(r);
      y = acc;
    } else {
      let frozenH = 0;
      for (let r = 1; r <= freezeRows; r++) frozenH += rowH(r);
      const scrollAreaY = sp(HEADER_H) + frozenH;

      const logicalScrollY = this.scrollHost.scrollTop / cs;
      const rowAxis = getSheetAxes(ws, mdw).row;
      const { index: startRow, partial: offsetY } =
        rowAxis.indexAt(logicalScrollY + rowAxis.offsetOf(freezeRows + 1));

      let acc = scrollAreaY - offsetY * cs;
      if (row >= startRow) {
        for (let r = startRow; r < row; r++) acc += rowH(r);
      } else {
        for (let r = row; r < startRow; r++) acc -= rowH(r);
      }
      y = acc;
    }

    return { x, y, w: colW(col), h: rowH(row) };
  }

  /** Returns the current selection, including mode. */
  get selection(): CellRange | null {
    if (!this.anchorCell || !this.activeCell) return null;
    return { anchor: this.anchorCell, active: this.activeCell, mode: this.selectionMode };
  }

  /**
   * Returns what the header area contains at the given client coordinates.
   * Returns null when the point is in the cell grid (not a header).
   */
  private getHeaderHit(
    clientX: number,
    clientY: number,
  ): { kind: 'corner' } | { kind: 'row'; row: number } | { kind: 'col'; col: number } | null {
    const ws = this.currentWorksheet;
    if (!ws) return null;
    const cs = this.opts.cellScale ?? 1;
    const rect = this.canvasArea.getBoundingClientRect();
    const lx = (clientX - rect.left) / cs;
    const ly = (clientY - rect.top) / cs;

    const inRowHeader = lx < HEADER_W;
    const inColHeader = ly < HEADER_H;
    if (!inRowHeader && !inColHeader) return null;
    if (inRowHeader && inColHeader) return { kind: 'corner' };

    const freezeRows = ws.freezeRows ?? 0;
    const freezeCols = ws.freezeCols ?? 0;

    if (inRowHeader) {
      // Determine which row was clicked
      const innerY = ly - HEADER_H;
      if (innerY < 0) return { kind: 'corner' };
      let frozenH = 0;
      const frozenRowH: number[] = [];
      for (let r = 1; r <= freezeRows; r++) {
        const h = rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight);
        frozenRowH.push(h); frozenH += h;
      }
      if (innerY < frozenH) {
        let acc = 0;
        for (let r = 0; r < freezeRows; r++) {
          acc += frozenRowH[r];
          if (innerY < acc) return { kind: 'row', row: r + 1 };
        }
        return null;
      }
      const contentY = innerY - frozenH + this.scrollHost.scrollTop / cs;
      let acc = 0;
      for (let r = freezeRows + 1; r <= 1048576; r++) {
        acc += rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight);
        if (contentY < acc) return { kind: 'row', row: r };
      }
      return null;
    }

    // inColHeader
    const innerX = lx - HEADER_W;
    if (innerX < 0) return { kind: 'corner' };
    let frozenW = 0;
    const frozenColW: number[] = [];
    for (let c = 1; c <= freezeCols; c++) {
      const w = colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth, getMdwForWorksheet(ws));
      frozenColW.push(w); frozenW += w;
    }
    if (innerX < frozenW) {
      let acc = 0;
      for (let c = 0; c < freezeCols; c++) {
        acc += frozenColW[c];
        if (innerX < acc) return { kind: 'col', col: c + 1 };
      }
      return null;
    }
    const contentX = innerX - frozenW + this.scrollHost.scrollLeft / cs;
    let acc = 0;
    for (let c = freezeCols + 1; c <= 16384; c++) {
      acc += colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth, getMdwForWorksheet(ws));
      if (contentX < acc) return { kind: 'col', col: c };
    }
    return null;
  }

  /** Copy the selected cell range as tab-separated text to the clipboard. */
  private copySelection(): void {
    const ws = this.currentWorksheet;
    if (!ws || !this.anchorCell || !this.activeCell) return;

    // Determine actual data extent for rows/cols/all modes
    let maxRow = 1, maxCol = 1;
    for (const row of ws.rows) {
      if (row.index > maxRow) maxRow = row.index;
      for (const cell of row.cells) {
        if (cell.col > maxCol) maxCol = cell.col;
      }
    }

    let r1: number, r2: number, c1: number, c2: number;
    if (this.selectionMode === 'all') {
      r1 = 1; r2 = maxRow; c1 = 1; c2 = maxCol;
    } else if (this.selectionMode === 'rows') {
      r1 = Math.min(this.anchorCell.row, this.activeCell.row);
      r2 = Math.max(this.anchorCell.row, this.activeCell.row);
      c1 = 1; c2 = maxCol;
    } else if (this.selectionMode === 'cols') {
      c1 = Math.min(this.anchorCell.col, this.activeCell.col);
      c2 = Math.max(this.anchorCell.col, this.activeCell.col);
      r1 = 1; r2 = maxRow;
    } else {
      r1 = Math.min(this.anchorCell.row, this.activeCell.row);
      r2 = Math.max(this.anchorCell.row, this.activeCell.row);
      c1 = Math.min(this.anchorCell.col, this.activeCell.col);
      c2 = Math.max(this.anchorCell.col, this.activeCell.col);
    }

    const cellMap = new Map<string, string>();
    for (const row of ws.rows) {
      if (row.index < r1 || row.index > r2) continue;
      for (const cell of row.cells) {
        if (cell.col < c1 || cell.col > c2) continue;
        const v = cell.value;
        let text = '';
        if (v.type === 'text') text = v.runs ? v.runs.map((r) => r.text).join('') : v.text;
        else if (v.type === 'number') text = String(v.number);
        else if (v.type === 'bool') text = v.bool ? 'TRUE' : 'FALSE';
        else if (v.type === 'error') text = v.error;
        if (text) cellMap.set(`${row.index}:${cell.col}`, text);
      }
    }

    const lines: string[] = [];
    for (let r = r1; r <= r2; r++) {
      const cols: string[] = [];
      for (let c = c1; c <= c2; c++) cols.push(cellMap.get(`${r}:${c}`) ?? '');
      lines.push(cols.join('\t'));
    }
    navigator.clipboard.writeText(lines.join('\n')).catch(() => undefined);
  }

  private updateSelectionOverlay(): void {
    this.selectionOverlay.innerHTML = '';
    if (!this.anchorCell || !this.activeCell) return;

    const cs = this.opts.cellScale ?? 1;
    const ws = this.currentWorksheet;
    const freezeRows = ws?.freezeRows ?? 0;
    const freezeCols = ws?.freezeCols ?? 0;
    // Same per-cell rounding as getCellRect / the renderer, so clamp
    // boundaries land on the canvas's actual pixel edges.
    const sp = (px: number) => Math.round(px * cs);
    const headerW = sp(HEADER_W);
    const headerH = sp(HEADER_H);

    let frozenH = 0;
    if (ws) for (let r = 1; r <= freezeRows; r++) frozenH += sp(rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight));
    let frozenW = 0;
    if (ws) for (let c = 1; c <= freezeCols; c++) frozenW += sp(colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth, getMdwForWorksheet(ws)));

    let x: number, y: number, w: number, h: number;
    let selR1 = 1, selC1 = 1;

    if (this.selectionMode === 'all') {
      x = headerW;
      y = headerH;
      w = this.canvasArea.clientWidth - headerW;
      h = this.canvasArea.clientHeight - headerH;
    } else if (this.selectionMode === 'rows') {
      selR1 = Math.min(this.anchorCell.row, this.activeCell.row);
      const r2 = Math.max(this.anchorCell.row, this.activeCell.row);
      const top = this.getCellRect(selR1, 1);
      const bot = this.getCellRect(r2, 1);
      if (!top || !bot) return;
      x = headerW;
      y = top.y;
      w = this.canvasArea.clientWidth - headerW;
      h = bot.y + bot.h - top.y;
    } else if (this.selectionMode === 'cols') {
      selC1 = Math.min(this.anchorCell.col, this.activeCell.col);
      const c2 = Math.max(this.anchorCell.col, this.activeCell.col);
      const left = this.getCellRect(1, selC1);
      const right = this.getCellRect(1, c2);
      if (!left || !right) return;
      x = left.x;
      y = headerH;
      w = right.x + right.w - left.x;
      h = this.canvasArea.clientHeight - headerH;
    } else {
      selR1 = Math.min(this.anchorCell.row, this.activeCell.row);
      const r2 = Math.max(this.anchorCell.row, this.activeCell.row);
      selC1 = Math.min(this.anchorCell.col, this.activeCell.col);
      const c2 = Math.max(this.anchorCell.col, this.activeCell.col);
      const tl = this.getCellRect(selR1, selC1);
      const br = this.getCellRect(r2, c2);
      if (!tl || !br) return;
      x = tl.x; y = tl.y;
      w = br.x + br.w - tl.x;
      h = br.y + br.h - tl.y;
    }

    // Clamp to header boundaries so the overlay never overlaps fixed headers.
    if (x < headerW) { w -= headerW - x; x = headerW; }
    if (y < headerH) { h -= headerH - y; y = headerH; }

    // Clamp scrollable-region selections at the frozen pane boundary.
    // Frozen cells legitimately live inside the frozen area; scrollable cells
    // that have scrolled behind the frozen area must be clipped there instead.
    const frozenBoundX = headerW + frozenW;
    const frozenBoundY = headerH + frozenH;
    if (selC1 > freezeCols && x < frozenBoundX) { w -= frozenBoundX - x; x = frozenBoundX; }
    if (selR1 > freezeRows && y < frozenBoundY) { h -= frozenBoundY - y; y = frozenBoundY; }

    if (w <= 0 || h <= 0) return;
    const box = document.createElement('div');
    box.style.cssText =
      `position:absolute;` +
      `left:${x}px;top:${y}px;width:${w}px;height:${h}px;` +
      `box-sizing:border-box;border:2px solid #1a73e8;` +
      `background:rgba(26,115,232,0.08);pointer-events:none;`;
    this.selectionOverlay.appendChild(box);
  }

  private applyPointerSelection(clientX: number, clientY: number, shiftKey: boolean, pointerId: number, allowDrag: boolean): void {
    const headerHit = this.getHeaderHit(clientX, clientY);

    if (headerHit) {
      if (headerHit.kind === 'corner') {
        // Select all — no drag extension needed
        this.selectionMode = 'all';
        this.anchorCell = { row: 1, col: 1 };
        this.activeCell = { row: 1, col: 1 };
        this.isSelecting = false;
      } else if (headerHit.kind === 'row') {
        if (shiftKey && this.anchorCell && this.selectionMode === 'rows') {
          this.activeCell = { row: headerHit.row, col: 1 };
        } else {
          this.selectionMode = 'rows';
          this.anchorCell = { row: headerHit.row, col: 1 };
          this.activeCell = { row: headerHit.row, col: 1 };
          if (allowDrag) {
            this.isSelecting = true;
            this.scrollHost.setPointerCapture(pointerId);
          }
        }
      } else {
        if (shiftKey && this.anchorCell && this.selectionMode === 'cols') {
          this.activeCell = { row: 1, col: headerHit.col };
        } else {
          this.selectionMode = 'cols';
          this.anchorCell = { row: 1, col: headerHit.col };
          this.activeCell = { row: 1, col: headerHit.col };
          if (allowDrag) {
            this.isSelecting = true;
            this.scrollHost.setPointerCapture(pointerId);
          }
        }
      }
      this.updateSelectionOverlay();
      void this.renderCurrentSheet();
      this.opts.onSelectionChange?.(this.selection);
      return;
    }

    const cell = this.getCellAt(clientX, clientY);
    if (!cell) return;

    if (shiftKey && this.anchorCell && this.selectionMode === 'cells') {
      this.activeCell = cell;
    } else {
      this.selectionMode = 'cells';
      this.anchorCell = cell;
      this.activeCell = cell;
    }
    if (allowDrag) {
      this.isSelecting = true;
      this.scrollHost.setPointerCapture(pointerId);
    }
    this.updateSelectionOverlay();
    void this.renderCurrentSheet();
    this.opts.onSelectionChange?.(this.selection);
  }

  private setupSelectionEvents(): void {
    // Distance (CSS px) beyond which a touch/pen pointerdown→pointerup is treated as a swipe (scroll), not a tap.
    const TAP_SLOP = 8;

    this.scrollHost.addEventListener('pointerdown', (e: PointerEvent) => {
      if (e.button !== 0) return;

      // Touch / pen: defer selection until pointerup so swipe-to-scroll doesn't change the cell.
      // Mouse: select immediately to preserve drag-to-extend behavior.
      if (e.pointerType !== 'mouse') {
        this.pendingTap = { x: e.clientX, y: e.clientY, shiftKey: e.shiftKey, pointerId: e.pointerId };
        return;
      }

      this.applyPointerSelection(e.clientX, e.clientY, e.shiftKey, e.pointerId, true);
    });

    this.scrollHost.addEventListener('pointermove', (e: PointerEvent) => {
      // Cancel a pending tap once the pointer moves beyond the slop — the user is scrolling.
      if (this.pendingTap && this.pendingTap.pointerId === e.pointerId) {
        const dx = e.clientX - this.pendingTap.x;
        const dy = e.clientY - this.pendingTap.y;
        if (dx * dx + dy * dy > TAP_SLOP * TAP_SLOP) {
          this.pendingTap = null;
        }
      }

      if (!this.isSelecting) return;

      if (this.selectionMode === 'rows') {
        const hit = this.getHeaderHit(e.clientX, e.clientY);
        const row = hit?.kind === 'row' ? hit.row : this.getCellAt(e.clientX, e.clientY)?.row;
        if (!row || row === this.activeCell?.row) return;
        this.activeCell = { row, col: 1 };
      } else if (this.selectionMode === 'cols') {
        const hit = this.getHeaderHit(e.clientX, e.clientY);
        const col = hit?.kind === 'col' ? hit.col : this.getCellAt(e.clientX, e.clientY)?.col;
        if (!col || col === this.activeCell?.col) return;
        this.activeCell = { row: 1, col };
      } else {
        const cell = this.getCellAt(e.clientX, e.clientY);
        if (!cell || (cell.row === this.activeCell?.row && cell.col === this.activeCell?.col)) return;
        this.activeCell = cell;
      }

      this.updateSelectionOverlay();
      void this.renderCurrentSheet();
      this.opts.onSelectionChange?.(this.selection);
    });

    this.scrollHost.addEventListener('pointerup', (e: PointerEvent) => {
      if (this.pendingTap && this.pendingTap.pointerId === e.pointerId) {
        const dx = e.clientX - this.pendingTap.x;
        const dy = e.clientY - this.pendingTap.y;
        if (dx * dx + dy * dy <= TAP_SLOP * TAP_SLOP) {
          this.applyPointerSelection(e.clientX, e.clientY, this.pendingTap.shiftKey, e.pointerId, false);
        }
        this.pendingTap = null;
      }
      this.isSelecting = false;
    });

    this.scrollHost.addEventListener('pointercancel', (e: PointerEvent) => {
      if (this.pendingTap && this.pendingTap.pointerId === e.pointerId) {
        this.pendingTap = null;
      }
      this.isSelecting = false;
    });

    this.keydownHandler = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
        this.copySelection();
      }
    };
    document.addEventListener('keydown', this.keydownHandler);
  }

  private buildTabs(): void {
    this.tabStrip.innerHTML = '';
    this.tabs = [];
    this.tabColors = this.workbook.tabColors;
    this.workbook.sheetNames.forEach((name, i) => {
      const btn = document.createElement('button');
      btn.textContent = name;
      btn.title = name;
      btn.style.cssText = this.tabStyle(false, this.tabColors[i]);
      btn.addEventListener('click', () => this.showSheet(i));
      this.tabStrip.appendChild(btn);
      this.tabs.push(btn);
    });
    this.updateNavButtons();
  }

  private makeNavButton(glyph: string, label: string, onClick: () => void): HTMLButtonElement {
    const btn = document.createElement('button');
    btn.textContent = glyph;
    btn.setAttribute('aria-label', label);
    btn.title = label;
    btn.classList.add('xlsx-tab-nav');
    btn.style.cssText = this.navButtonStyle(false);
    btn.addEventListener('click', onClick);
    return btn;
  }

  private navButtonStyle(disabled: boolean): string {
    // Plain triangle icons — no border / tab chrome. The background (incl. the
    // hover tint) lives in the injected `.xlsx-tab-nav` stylesheet so the inline
    // style does not shadow the `:hover` rule.
    const base =
      `flex:1;height:100%;padding:0;` +
      `display:flex;align-items:center;justify-content:center;` +
      `border:none;color:#666;font-size:9px;line-height:1;` +
      `box-sizing:border-box;outline:none;`;
    return disabled
      ? base + `opacity:0.3;cursor:default;pointer-events:none;`
      : base + `cursor:pointer;`;
  }

  private scrollTabs(dir: -1 | 1): void {
    const strip = this.tabStrip;
    const viewLeft = strip.scrollLeft;
    const viewRight = viewLeft + strip.clientWidth;
    let target: number | null = null;
    if (dir === 1) {
      // First tab clipped on the right; align its right edge to the viewport.
      for (const tab of this.tabs) {
        const right = tab.offsetLeft + tab.offsetWidth;
        if (right > viewRight + 1) {
          target = right - strip.clientWidth;
          break;
        }
      }
    } else {
      // Last tab clipped on the left; align its left edge to the viewport.
      for (let i = this.tabs.length - 1; i >= 0; i--) {
        const left = this.tabs[i].offsetLeft;
        if (left < viewLeft - 1) {
          target = left;
          break;
        }
      }
    }
    if (target !== null) {
      // Instant (not smooth) so the disabled state is consistent the moment the
      // click resolves — keeps the interaction deterministic to drive/test.
      strip.scrollLeft = Math.max(0, target);
    }
    this.updateNavButtons();
  }

  private updateNavButtons(): void {
    const strip = this.tabStrip;
    const atStart = strip.scrollLeft <= 0;
    const atEnd = strip.scrollLeft + strip.clientWidth >= strip.scrollWidth - 1;
    // No overflow => scrollWidth ≈ clientWidth => both ends true => both disabled.
    this.navPrev.style.cssText = this.navButtonStyle(atStart);
    this.navNext.style.cssText = this.navButtonStyle(atEnd);
  }

  private updateTabActive(index: number): void {
    this.tabs.forEach((btn, i) => {
      btn.style.cssText = this.tabStyle(i === index, this.tabColors[i]);
    });
    // Keep the active tab visible by scrolling the tab strip HORIZONTALLY only.
    // `scrollIntoView` walks every scrollable ancestor, so it also scrolls the
    // page vertically — on first load that jumped the whole page down to the
    // tab bar (the active sheet is set during load). Adjust the strip's
    // scrollLeft directly so the page never moves.
    const tab = this.tabs[index];
    if (tab) {
      const strip = this.tabStrip;
      const tabRect = tab.getBoundingClientRect();
      const stripRect = strip.getBoundingClientRect();
      if (tabRect.left < stripRect.left) {
        strip.scrollLeft -= stripRect.left - tabRect.left;
      } else if (tabRect.right > stripRect.right) {
        strip.scrollLeft += tabRect.right - stripRect.right;
      }
    }
  }

  private tabStyle(active: boolean, tabColor?: string | null): string {
    // Active tab renders taller than inactive so the selected sheet draws the
    // eye. Tabs align to flex-end, so shorter inactive tabs sit lower and the
    // active tab sticks up. Font size also bumps a hair on active.
    const activeH = TAB_BAR_H - 2;
    const inactiveH = TAB_BAR_H - 5;
    const base =
      `display:inline-block;flex:none;padding:0 14px;position:relative;` +
      `border:1px solid #c8ccd0;border-bottom:none;border-radius:3px 3px 0 0;` +
      `cursor:pointer;white-space:nowrap;max-width:160px;overflow:hidden;text-overflow:ellipsis;` +
      `outline:none;box-sizing:border-box;`;
    // `<sheetPr><tabColor>` renders as a color bar along the tab's bottom edge
    // (Excel's "sheet tab color" treatment), drawn as an inset bottom shadow so
    // it doesn't fight the tab's own border/background. The active tab keeps a
    // thinner bar since its bottom merges into the white sheet body.
    const bar = tabColor
      ? `box-shadow:inset 0 -${active ? 2 : 3}px 0 0 ${tabColor};`
      : '';
    return active
      ? base +
        `height:${activeH}px;font-size:13px;` +
        `background:#fff;color:#000;border-bottom:1px solid #fff;font-weight:600;top:1px;` +
        bar
      : base +
        `height:${inactiveH}px;font-size:11px;` +
        `background:#e0e0e0;color:#555;` +
        bar;
  }

  /** Excel-style zoom control pinned to the right end of the tab bar:
   *  `−  [────slider────]  +  100%`. Live-updates the cell scale on input. */
  private buildZoomControl(): HTMLDivElement {
    const zoomMin = this.opts.zoomMin ?? 0.1;
    const zoomMax = this.opts.zoomMax ?? 4;
    const cur = this.opts.cellScale ?? 1;

    const wrap = document.createElement('div');
    wrap.style.cssText =
      `display:flex;align-items:center;flex-shrink:0;gap:2px;` +
      `padding:0 10px;height:100%;color:#555;font-size:12px;user-select:none;`;

    const mkBtn = (glyph: string, label: string, delta: number): HTMLButtonElement => {
      const b = document.createElement('button');
      b.textContent = glyph;
      b.setAttribute('aria-label', label);
      b.title = label;
      b.style.cssText =
        `width:18px;height:18px;padding:0;border:none;background:transparent;` +
        `color:#555;font-size:14px;line-height:1;cursor:pointer;border-radius:3px;`;
      b.addEventListener('click', () => this.setScale((this.opts.cellScale ?? 1) + delta));
      return b;
    };

    // The slider works in "position" units [0,100]; 50 is dead-center and maps
    // to 100% so each half is its own linear segment (zoomMin→1 on the left,
    // 1→zoomMax on the right), mirroring Excel's status-bar zoom where 100% sits
    // in the middle even though the range (10%–400%) is asymmetric.
    const slider = document.createElement('input');
    slider.type = 'range';
    slider.min = '0';
    slider.max = '100';
    slider.step = 'any';
    slider.value = String(this.zoomScaleToPos(cur, zoomMin, zoomMax));
    slider.setAttribute('aria-label', 'Zoom');
    slider.title = 'Zoom';
    slider.classList.add('xlsx-zoom-slider');
    slider.style.cssText = `width:90px;cursor:pointer;`;
    slider.addEventListener('input', () =>
      this.setScale(this.zoomPosToScale(Number(slider.value), zoomMin, zoomMax)),
    );

    const label = document.createElement('span');
    label.textContent = `${Math.round(cur * 100)}%`;
    label.style.cssText = `min-width:42px;margin-left:6px;text-align:right;font-variant-numeric:tabular-nums;`;

    wrap.appendChild(mkBtn('−', 'Zoom out', -0.1));
    wrap.appendChild(slider);
    wrap.appendChild(mkBtn('+', 'Zoom in', 0.1));
    wrap.appendChild(label);

    this.zoomSlider = slider;
    this.zoomLabel = label;
    return wrap;
  }

  /** Map a slider position [0,100] to a scale factor. 50 → 1.0 (100%), with a
   *  separate linear segment on each side so the center is always 100%. */
  private zoomPosToScale(pos: number, min: number, max: number): number {
    return pos <= 50
      ? min + (pos / 50) * (1 - min)
      : 1 + ((pos - 50) / 50) * (max - 1);
  }

  /** Inverse of {@link zoomPosToScale}: scale factor → slider position [0,100]. */
  private zoomScaleToPos(scale: number, min: number, max: number): number {
    const clamped = Math.min(max, Math.max(min, scale));
    return clamped <= 1
      ? ((clamped - min) / (1 - min)) * 50
      : 50 + ((clamped - 1) / (max - 1)) * 50;
  }

  /** Set the cell/header scale and re-lay-out the current sheet. Clamped to the
   *  zoom bounds; keeps the slider thumb, percentage label and the row-header-
   *  aligned tab-nav width in sync. */
  setScale(scale: number): void {
    const zoomMin = this.opts.zoomMin ?? 0.1;
    const zoomMax = this.opts.zoomMax ?? 4;
    // Snap to whole percent so the label and cellScale stay tidy.
    const pct = Math.min(
      Math.round(zoomMax * 100),
      Math.max(Math.round(zoomMin * 100), Math.round(scale * 100)),
    );
    const next = pct / 100;
    if (next === (this.opts.cellScale ?? 1)) return;
    this.opts.cellScale = next;

    if (this.zoomSlider) this.zoomSlider.value = String(this.zoomScaleToPos(next, zoomMin, zoomMax));
    if (this.zoomLabel) this.zoomLabel.textContent = `${pct}%`;
    // The tab-nav block spans the row-header width, which scales with the cells.
    this.navGroup.style.width = `${Math.round(HEADER_W * next)}px`;

    if (this.currentWorksheet) this.updateSpacerSize(this.currentWorksheet);
    void this.renderCurrentSheet();
    this.updateSelectionOverlay();
    this.updateNavButtons();
  }

  private updateSpacerSize(ws: Worksheet): void {
    const cs = this.opts.cellScale ?? 1;
    const mdw = getMdwForWorksheet(ws);
    // Match getCellRect / the renderer: round each cell's scaled width
    // independently so the scrollbar's max scroll lines up with the
    // canvas's furthest-right / furthest-down drawn cell edge.
    const sp = (px: number) => Math.round(px * cs);
    const freezeRows = ws.freezeRows ?? 0;
    const freezeCols = ws.freezeCols ?? 0;

    // Find actual scrollable data extent
    let maxRow = Math.max(50, freezeRows);
    let maxCol = Math.max(26, freezeCols);
    for (const row of ws.rows) {
      if (row.index > maxRow) maxRow = row.index;
      for (const cell of row.cells) {
        if (cell.col > maxCol) maxCol = cell.col;
      }
    }
    maxRow += 30;
    maxCol += 10;

    // Spacer = sp(header) + Σ sp(width) for every visible col, same for rows.
    let totalW = sp(HEADER_W);
    for (let c = 1; c <= maxCol; c++) {
      totalW += sp(colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth, mdw));
    }
    let totalH = sp(HEADER_H);
    for (let r = 1; r <= maxRow; r++) {
      totalH += sp(rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight));
    }

    this.spacer.style.width = `${totalW}px`;
    this.spacer.style.height = `${totalH}px`;
  }

  private async renderCurrentSheet(): Promise<void> {
    if (!this.currentWorksheet) return;
    const ws = this.currentWorksheet;
    const w = this.canvasArea.clientWidth;
    const h = this.canvasArea.clientHeight;
    if (w <= 0 || h <= 0) return;

    const cs = this.opts.cellScale ?? 1;
    const dpr = window.devicePixelRatio ?? 1;

    const freezeRows = ws.freezeRows ?? 0;
    const freezeCols = ws.freezeCols ?? 0;

    // Compute frozen area in logical (unscaled) pixels
    let frozenW = 0;
    for (let c = 1; c <= freezeCols; c++) {
      frozenW += colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth, getMdwForWorksheet(ws));
    }
    let frozenH = 0;
    for (let r = 1; r <= freezeRows; r++) {
      frozenH += rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight);
    }

    // DOM scrollLeft/scrollTop are in scaled (physical) CSS pixels.
    // Convert to logical pixels for cell-finding by dividing by cs.
    const logicalScrollX = this.scrollHost.scrollLeft / cs;
    const logicalScrollY = this.scrollHost.scrollTop / cs;

    // Find startCol / startRow in logical pixel space (binary search over the
    // per-sheet cumulative-offset axes instead of an O(n) walk from cell 1).
    const axes = getSheetAxes(ws, getMdwForWorksheet(ws));
    const { index: startCol, partial: offsetX } =
      axes.col.indexAt(logicalScrollX + axes.col.offsetOf(freezeCols + 1));
    const { index: startRow, partial: offsetY } =
      axes.row.indexAt(logicalScrollY + axes.row.offsetOf(freezeRows + 1));

    // Effective scrollable area in logical pixels (canvas / cs - headers - frozen)
    const cellW = w / cs - HEADER_W - frozenW;
    const cellH = h / cs - HEADER_H - frozenH;

    // Compute exact number of visible columns by walking actual widths (+ 2 buffer)
    let cols = 0;
    { let xAcc = -offsetX; let c = startCol;
      while (xAcc < cellW + offsetX && c <= 16384) {
        xAcc += colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth, getMdwForWorksheet(ws)); cols++; c++;
      }
      cols += 2;
    }
    // Compute exact number of visible rows by walking actual heights (+ 2 buffer)
    let rows = 0;
    { let yAcc = -offsetY; let r = startRow;
      while (yAcc < cellH + offsetY && r <= 1048576) {
        yAcc += rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight); rows++; r++;
      }
      rows += 2;
    }

    const viewport: ViewportRange = { row: startRow, col: startCol, rows, cols };

    const { selectedRowRange, selectedColRange } = this.computeHeaderHighlight();

    await this.workbook.renderViewport(this.canvas, this.currentSheet, viewport, {
      width: w,
      height: h,
      dpr,
      cellScale: cs,
      scrollOffsetX: offsetX,
      scrollOffsetY: offsetY,
      freezeRows,
      freezeCols,
      selectedRowRange,
      selectedColRange,
    });
  }

  private computeHeaderHighlight(): {
    selectedRowRange: { start: number; end: number; strong: boolean } | null;
    selectedColRange: { start: number; end: number; strong: boolean } | null;
  } {
    if (!this.anchorCell || !this.activeCell) {
      return { selectedRowRange: null, selectedColRange: null };
    }
    const ALL = Number.MAX_SAFE_INTEGER;
    const r1 = Math.min(this.anchorCell.row, this.activeCell.row);
    const r2 = Math.max(this.anchorCell.row, this.activeCell.row);
    const c1 = Math.min(this.anchorCell.col, this.activeCell.col);
    const c2 = Math.max(this.anchorCell.col, this.activeCell.col);
    switch (this.selectionMode) {
      case 'cells':
        return {
          selectedRowRange: { start: r1, end: r2, strong: false },
          selectedColRange: { start: c1, end: c2, strong: false },
        };
      case 'rows':
        return {
          selectedRowRange: { start: r1, end: r2, strong: true },
          selectedColRange: { start: 1, end: ALL, strong: false },
        };
      case 'cols':
        return {
          selectedRowRange: { start: 1, end: ALL, strong: false },
          selectedColRange: { start: c1, end: c2, strong: true },
        };
      case 'all':
        return {
          selectedRowRange: { start: 1, end: ALL, strong: true },
          selectedColRange: { start: 1, end: ALL, strong: true },
        };
    }
  }

  get sheetNames(): string[] {
    return this.wb?.sheetNames ?? [];
  }

  /** The underlying <canvas> element the grid is drawn on. */
  get canvasElement(): HTMLCanvasElement {
    return this.canvas;
  }

  destroy(): void {
    this.resizeObserver?.disconnect();
    if (this.keydownHandler) {
      document.removeEventListener('keydown', this.keydownHandler);
    }
    this.wb?.destroy();
  }
}
