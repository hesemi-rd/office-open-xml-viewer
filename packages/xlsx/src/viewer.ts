import { XlsxWorkbook } from './workbook.js';
import type { ViewportRange, Worksheet } from './types.js';
import { HEADER_W, HEADER_H, colWidthToPx, rowHeightToPx, getMdwForWorksheet } from './renderer.js';

const TAB_BAR_H = 30;

export interface XlsxViewerOptions {
  /** Scale factor for cell/header dimensions (default 1). 0.5 = half size. */
  cellScale?: number;
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

export class XlsxViewer {
  private wb: XlsxWorkbook;
  private canvas: HTMLCanvasElement;
  private canvasArea: HTMLDivElement;
  private scrollHost: HTMLDivElement;
  private spacer: HTMLDivElement;
  private tabBar: HTMLDivElement;
  private tabStrip: HTMLDivElement;
  private navPrev: HTMLButtonElement;
  private navNext: HTMLButtonElement;
  private tabs: HTMLButtonElement[] = [];
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
    this.wb = new XlsxWorkbook();

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

    // The scrollable strip that actually holds the sheet tabs. position:relative
    // so each tab's offsetLeft is measured against the strip's scroll content.
    this.tabStrip = document.createElement('div');
    this.tabStrip.style.cssText =
      `position:relative;display:flex;align-items:flex-end;flex:1;min-width:0;height:100%;` +
      `overflow-x:auto;overflow-y:hidden;gap:1px;scrollbar-width:none;`;
    this.tabStrip.classList.add('xlsx-tab-strip');
    const style = document.createElement('style');
    style.textContent =
      `.xlsx-tab-strip::-webkit-scrollbar{display:none}` +
      `.xlsx-tab-nav{background:transparent;transition:background 0.1s;}` +
      `.xlsx-tab-nav:hover{background:rgba(0,0,0,0.08);}`;
    document.head.appendChild(style);
    this.tabStrip.addEventListener('scroll', () => this.updateNavButtons());

    this.tabBar.appendChild(navGroup);
    this.tabBar.appendChild(this.tabStrip);

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

  async load(source: string | ArrayBuffer): Promise<void> {
    try {
      await this.wb.load(source, {
        useGoogleFonts: this.opts.useGoogleFonts,
        maxZipEntryBytes: this.opts.maxZipEntryBytes,
      });
      this.buildTabs();
      this.opts.onReady?.(this.wb.sheetNames);
      await this.showSheet(0);
    } catch (err) {
      this.opts.onError?.(err instanceof Error ? err : new Error(String(err)));
    }
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
    this.currentWorksheet = await this.wb.getWorksheet(index);
    this.updateSpacerSize(this.currentWorksheet);
    await this.renderCurrentSheet();
    this.opts.onSheetChange?.(index, this.wb.sheetNames[index] ?? '');
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

      // Mirror renderCurrentSheet's startCol / offsetX search.
      const logicalScrollX = this.scrollHost.scrollLeft / cs;
      let startCol = freezeCols + 1;
      let xAcc = 0;
      let offsetX = 0;
      while (startCol <= 16384) {
        const cw = colWidthToPx(ws.colWidths[startCol] ?? ws.defaultColWidth, mdw);
        if (xAcc + cw > logicalScrollX) { offsetX = logicalScrollX - xAcc; break; }
        xAcc += cw;
        startCol++;
      }

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
      let startRow = freezeRows + 1;
      let yAcc = 0;
      let offsetY = 0;
      while (startRow <= 1048576) {
        const rh = rowHeightToPx(ws.rowHeights[startRow] ?? ws.defaultRowHeight);
        if (yAcc + rh > logicalScrollY) { offsetY = logicalScrollY - yAcc; break; }
        yAcc += rh;
        startRow++;
      }

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
    this.wb.sheetNames.forEach((name, i) => {
      const btn = document.createElement('button');
      btn.textContent = name;
      btn.title = name;
      btn.style.cssText = this.tabStyle(false);
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
      btn.style.cssText = this.tabStyle(i === index);
    });
    this.tabs[index]?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  private tabStyle(active: boolean): string {
    // Active tab renders taller than inactive so the selected sheet draws the
    // eye. Tabs align to flex-end, so shorter inactive tabs sit lower and the
    // active tab sticks up. Font size also bumps a hair on active.
    const activeH = TAB_BAR_H - 2;
    const inactiveH = TAB_BAR_H - 5;
    const base =
      `display:inline-block;flex:none;padding:0 14px;` +
      `border:1px solid #c8ccd0;border-bottom:none;border-radius:3px 3px 0 0;` +
      `cursor:pointer;white-space:nowrap;max-width:160px;overflow:hidden;text-overflow:ellipsis;` +
      `outline:none;box-sizing:border-box;`;
    return active
      ? base +
        `height:${activeH}px;font-size:13px;` +
        `background:#fff;color:#000;border-bottom:1px solid #fff;font-weight:600;position:relative;top:1px;`
      : base +
        `height:${inactiveH}px;font-size:11px;` +
        `background:#e0e0e0;color:#555;`;
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

    // Find startCol in logical pixel space
    let startCol = freezeCols + 1;
    let xAcc = 0;
    let offsetX = 0;
    while (true) {
      const cw = colWidthToPx(ws.colWidths[startCol] ?? ws.defaultColWidth, getMdwForWorksheet(ws));
      if (xAcc + cw > logicalScrollX) { offsetX = logicalScrollX - xAcc; break; }
      xAcc += cw;
      startCol++;
      if (startCol > 16384) break;
    }

    // Find startRow in logical pixel space
    let startRow = freezeRows + 1;
    let yAcc = 0;
    let offsetY = 0;
    while (true) {
      const rh = rowHeightToPx(ws.rowHeights[startRow] ?? ws.defaultRowHeight);
      if (yAcc + rh > logicalScrollY) { offsetY = logicalScrollY - yAcc; break; }
      yAcc += rh;
      startRow++;
      if (startRow > 1048576) break;
    }

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

    await this.wb.renderViewport(this.canvas, this.currentSheet, viewport, {
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
    return this.wb.sheetNames;
  }

  destroy(): void {
    this.resizeObserver?.disconnect();
    if (this.keydownHandler) {
      document.removeEventListener('keydown', this.keydownHandler);
    }
    this.wb.destroy();
  }
}
