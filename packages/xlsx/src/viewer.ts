import { XlsxWorkbook } from './workbook.js';
import type { ViewportRange, Worksheet, XlsxComment } from './types.js';
import type { LoadOptions } from '@silurus/ooxml-core';
import { nextVisibleIndex, resolveVisibleIndex, countVisible, zoomStepScale } from '@silurus/ooxml-core';
import { HEADER_W, HEADER_H, colWidthToPx, rowHeightToPx, pxToColWidth, pxToRowHeight, getMdwForWorksheet, rtlMirrorX } from './renderer.js';
import { findListValidationAt } from './data-validation.js';
import { parseA1 } from './a1.js';
import { computeCommentPopupPosition } from './comment-popup.js';
import {
  computeValidationPanelPosition,
  type ResolvedList,
} from './validation-list.js';

// Re-exported for the existing xlsx zoom tests (resize-zoom.test.ts imports it
// from this module) and any consumer that referenced it here before it moved to
// @silurus/ooxml-core. The single source of truth is core (design §5.2).
export { zoomStepScale } from '@silurus/ooxml-core';

/** Delay (ms) before a hovered comment popup appears. A short hover dwell
 *  prevents the popup from flickering while the cursor sweeps across many
 *  commented cells; ~150ms is the common tooltip-show threshold (responsive yet
 *  long enough to suppress transient passes). Excel itself uses a comparable
 *  short hover delay before showing a note. */
const COMMENT_POPUP_DELAY_MS = 150;
/** Max width of the comment popup body (CSS px). */
const COMMENT_POPUP_MAX_W = 280;
/** Max height before the body scrolls/clips (CSS px). */
const COMMENT_POPUP_MAX_H = 200;

/** Max width of the list-validation dropdown panel (CSS px). */
const VALIDATION_PANEL_MAX_W = 240;
/** Max height before the value list scrolls (CSS px). */
const VALIDATION_PANEL_MAX_H = 200;

const TAB_BAR_H = 30;
// Gap between adjacent sheet tabs. The first tab also gets this much leading
// space so it is offset from the row-header boundary by the same margin that
// separates tabs from each other.
const TAB_GAP = 1;

/** How {@link XlsxViewer} presents hidden sheets (`<sheet state>`, §18.2.19). */
export type HiddenSheetMode = 'show' | 'skip' | 'dim';

/** `'dim'`-mode tab opacity: hidden/veryHidden tabs are greyed but selectable.
 *  A UI-presentation default (ECMA-376 defines no hidden-tab rendering); mirrors
 *  the named pptx `DEFAULT_HIDDEN_DIM` constant. */
const HIDDEN_TAB_DIM_OPACITY = 0.45;

/** Marker attribute on the single injected viewer stylesheet, so the module-
 *  level injector is idempotent and destroy() can leave it in place. */
const VIEWER_STYLE_ATTR = 'data-xlsx-viewer-styles';

/** Class-constant CSS shared by every XlsxViewer: it styles pseudo-elements
 *  (scrollbar, slider track/thumb) that inline `element.style` cannot reach, so
 *  it must live in a stylesheet rather than on the elements. */
const VIEWER_STYLE_CSS =
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

/**
 * Inject the shared viewer stylesheet into `document.head` exactly once for the
 * whole module, keyed by the {@link VIEWER_STYLE_ATTR} marker. Earlier this ran
 * per-instance, so every mount/unmount cycle leaked another `<style>` into the
 * head (unbounded growth). It is deliberately NEVER removed on destroy: the CSS
 * is a class constant that any still-live viewer may depend on, and a single
 * leftover `<style>` after the last teardown is harmless (a fixed, bounded cost,
 * not a per-instance leak).
 */
function ensureViewerStyleInjected(): void {
  if (typeof document === 'undefined' || !document.head) return;
  if (document.head.querySelector(`style[${VIEWER_STYLE_ATTR}]`)) return;
  const style = document.createElement('style');
  style.setAttribute(VIEWER_STYLE_ATTR, '');
  style.textContent = VIEWER_STYLE_CSS;
  document.head.appendChild(style);
}

export interface XlsxViewerOptions extends LoadOptions {
  /** Scale factor for cell/header dimensions (default 1). 0.5 = half size. */
  cellScale?: number;
  /**
   * Enable drag-to-resize of column widths / row heights by dragging header
   * borders. Resizing only changes the on-screen view — it never modifies the
   * loaded file. Default: true.
   */
  resizable?: boolean;
  /** Show the Excel-style zoom slider at the right end of the sheet-tab bar.
   *  Default `true`. Set `false` to hide it (e.g. when the host supplies its
   *  own zoom control). */
  showZoomSlider?: boolean;
  /** Lower/upper bounds for the zoom slider as scale factors. Default 0.1–4
   *  (10%–400%, matching Excel's zoom range). */
  zoomMin?: number;
  zoomMax?: number;
  onReady?: (sheetNames: string[]) => void;
  /**
   * Called when the active sheet changes, with the new sheet's zero-based
   * `index` and the `total` number of sheets in the workbook. This mirrors the
   * docx `onPageChange` and pptx `onSlideChange` contracts so all three viewers
   * share one callback shape. To get the sheet *name*, look it up by index from
   * `viewer.sheetNames[index]` (or the `sheetNames` array delivered to
   * `onReady`).
   */
  onSheetChange?: (index: number, total: number) => void;
  onError?: (err: Error) => void;
  /** Called when the selected cell range changes. null means no selection. */
  onSelectionChange?: (selection: CellRange | null) => void;
  /**
   * Color of the cell-selection highlight. A single CSS color drives both the
   * selection rectangle's border (drawn in this color) and its fill (the same
   * color made translucent — see {@link selectionOverlayStyle}), so callers pick
   * one accent color instead of a separate border + background. Any CSS color
   * string works (`#1a73e8`, `rgb(...)`, `tomato`, …). Default `#1a73e8`
   * (Google blue), matching the historical look. Can also be changed at runtime
   * via {@link XlsxViewer.setSelectionColor}.
   */
  selectionColor?: string;
  /**
   * `'main'` (default): parse in a worker, render on the main thread. `'worker'`:
   * parse AND render entirely inside the worker and paint the returned
   * ImageBitmap onto the viewer's canvas, so document rendering never blocks the
   * UI thread. All interaction (scroll, sheet tabs, frozen panes, zoom, cell
   * selection) is unchanged. Requires `Worker` + `OffscreenCanvas`. Equations
   * require `'main'` (the math engine cannot cross the worker boundary).
   */
  mode?: 'main' | 'worker';
  /**
   * How hidden / veryHidden sheets (`<sheet state>`, ECMA-376 §18.2.19) are
   * presented:
   * - `'show'` (default): every sheet gets a tab — current behavior.
   * - `'skip'`: hidden/veryHidden sheets get no tab and are jumped over by
   *   `nextSheet`/`prevSheet` and initial load; absolute indices are unchanged,
   *   and an explicit `goToSheet(i)` to a hidden sheet is still honored.
   * - `'dim'`: hidden/veryHidden tabs are shown greyed but stay selectable.
   *
   * Named to match the {@link XlsxViewer.hiddenSheetMode} getter and
   * {@link XlsxViewer.setHiddenSheetMode} setter. Mirrors pptx `hiddenSlideMode`.
   */
  hiddenSheetMode?: HiddenSheetMode;
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

  /**
   * Resolve the 1-based index a hit-test lands on in the *scrollable* region,
   * given `content` = the logical-px distance from the START of `firstScrollable`
   * (i.e. `innerPos - frozenExtent + logicalScroll`). Returns `null` when the
   * point falls at or past the end of the very last cell (`maxIndex`), exactly
   * like the previous linear scan that walked `r = firstScrollable … maxIndex`
   * accumulating sizes and returned `null` if none satisfied `content < acc`.
   *
   * O(log n) via {@link indexAt}: shifting `content` by `offsetOf(firstScrollable)`
   * expresses it in the absolute (index-1-origin) coordinate `indexAt` uses. The
   * only place the linear scan and `indexAt` disagree is past the last cell —
   * `indexAt` clamps to `maxIndex`, the scan returned `null` — so we reproduce the
   * null by testing whether the absolute offset reaches the end of `maxIndex`.
   */
  scrollableIndexAt(content: number, firstScrollable: number): number | null {
    const absOffset = content + this.offsetOf(firstScrollable);
    // Past the end of the last cell ⇒ the old scan found no `content < acc`.
    if (absOffset >= this.offsetOf(this.maxIndex) + this.sizeOf(this.maxIndex)) {
      return null;
    }
    return this.indexAt(absOffset).index;
  }

  /** Logical-px span of `index` (1-based): its custom size if any, else default. */
  private sizeOf(index: number): number {
    return this.offsetOf(index + 1) - this.offsetOf(index);
  }
}

/** Default cell-selection accent (Google blue), used when no `selectionColor`
 *  option is supplied. */
const DEFAULT_SELECTION_COLOR = '#1a73e8';

/** Half-width (CSS px) of the grab zone around a header border for
 *  drag-to-resize (issue #567), and the minimum size a column/row can be
 *  dragged to (logical px) so a collapsed band keeps a grabbable border. */
const RESIZE_GRAB_PX = 4;
const RESIZE_MIN_PX = 5;

/**
 * Pure hit predicate for drag-to-resize (issue #567): given a pointer
 * coordinate `pt` (in the header-strip's CSS-px axis — already RTL-un-mirrored
 * by the caller) and the candidate band trailing edges `edges`, return the band
 * index whose edge is within `grabPx` of `pt`, or `null` if none qualifies.
 *
 * `edges` is the candidate list the caller builds — for the band the pointer is
 * over (`hit`) Excel lets you resize the band whose *trailing* border you grab,
 * so the caller passes both `hit - 1` and `hit` (the neighbour-to-the-far-side
 * and the band itself); the first edge within the grab zone wins, in the order
 * given. An edge that sits at or under the header strip (`edge <= headerExtent`,
 * i.e. scrolled behind the frozen corner) is rejected — you can't grab a border
 * hidden under the header. Kept pure (no DOM, no `this`) so the off-by-one
 * geometry — exact-on-edge, within-grab, just-outside, `[hit-1, hit]` neighbour
 * selection, header rejection — is unit-testable. {@link XlsxViewer.getResizeTarget}
 * does the DOM/geometry and calls this.
 */
export function resizeHitIndex(
  pt: number,
  edges: { index: number; edge: number }[],
  grabPx: number,
  headerExtent: number,
): number | null {
  for (const { index, edge } of edges) {
    if (edge <= headerExtent) continue; // scrolled behind the header strip
    if (Math.abs(pt - edge) <= grabPx) return index;
  }
  return null;
}

/**
 * Derive the selection rectangle's `border` and `background` CSS from a single
 * accent color: the border is the color verbatim and the fill is the same color
 * at 8% opacity via `color-mix`, so any CSS color string (`#rgb`, `rgb(...)`,
 * named) yields a matching translucent fill without the caller computing an
 * rgba. For the default `#1a73e8` this reproduces the historical
 * `rgba(26,115,232,0.08)` fill.
 */
export function selectionOverlayStyle(color: string): { border: string; background: string } {
  return {
    border: `2px solid ${color}`,
    background: `color-mix(in srgb, ${color} 8%, transparent)`,
  };
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
  /** The single subtree root the constructor appended to the caller's
   *  container. destroy() removes it to return the container to its original
   *  (empty) state. */
  private wrapper!: HTMLDivElement;
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
  private _hiddenSheetMode: HiddenSheetMode;
  private currentWorksheet: Worksheet | null = null;
  private opts: XlsxViewerOptions;
  /** 'main' renders on this thread; 'worker' paints worker-produced bitmaps. */
  private readonly _mode: 'main' | 'worker';
  /** The canvas's bitmaprenderer context, used only in worker mode. A canvas
   *  holds one context type for its lifetime, so this is obtained once and the
   *  main-mode 2d render path is never used on the same canvas. */
  private _bitmapCtx: ImageBitmapRenderingContext | null = null;
  private resizeObserver: ResizeObserver | null = null;
  /**
   * Pending `requestAnimationFrame` handle for a coalesced re-render, or `null`
   * when none is scheduled. High-frequency event-driven repaints (scroll, live
   * resize drag, selection drag, container resize) route through
   * {@link scheduleRender} so at most one render runs per animation frame: a
   * burst of scroll events within a single frame collapses to one draw at the
   * frame's latest scroll position (`renderCurrentSheet` reads the live scroll
   * offset, so "latest wins" needs no stored position). Explicit API calls
   * (`showSheet`/`goToSheet`, `select`, `setScale`) stay synchronous — they must
   * paint immediately, not a frame later. `destroy()` cancels any pending frame.
   */
  private _rafId: number | null = null;
  /**
   * Start-anchored horizontal scroll position (the {@link effectiveScrollLeft}
   * value last produced by a real user scroll or a programmatic reset), kept
   * as the source of truth across container size changes. The native
   * `scrollLeft` cannot serve that role for RTL sheets (ECMA-376 §18.3.1.87):
   * it is the *inverse* of the start-anchored offset, and the browser clamps
   * any assignment to 0 while the host is unlaid-out (`display:none` mount —
   * e.g. a host revealed only after `load()` resolves), which would otherwise
   * strand the view at the sheet's far end once the host gains its real size.
   */
  private effectiveH = 0;

  // Selection state
  private anchorCell: CellAddress | null = null;
  private activeCell: CellAddress | null = null;
  private selectionMode: SelectionMode = 'cells';
  private isSelecting = false;
  private selectionOverlay: HTMLDivElement;
  private keydownHandler: ((e: KeyboardEvent) => void) | null = null;
  // Deferred selection press: committed on pointerup only if the pointer
  // neither moved beyond the tap threshold nor caused a scroll. Used for
  // touch/pen (swipe-to-scroll must not change the cell) and for mouse
  // presses inside the overlay-scrollbar band (a thumb drag must not select
  // the cell underneath).
  private pendingTap: { x: number; y: number; shiftKey: boolean; pointerId: number } | null = null;
  // In-flight column/row resize drag (issue #567). `originScaled` is the fixed
  // LTR edge the resized band grows from (left edge for a column, top for a row)
  // in canvasArea CSS px; `mdw` is captured once so the live px→model-unit
  // conversion is stable across the drag. A resize is a *view-only* adjustment:
  // it mutates the in-memory worksheet's colWidths/rowHeights, never the file.
  private resizeDrag:
    | { kind: 'col' | 'row'; index: number; originScaled: number; mdw: number; pointerId: number }
    | null = null;

  // ─── Comment hover popup (Excel-style note) ───────────────────────────────
  /** DOM overlay element that shows the hovered cell's comment. Lives in
   *  canvasArea above the scrollHost; `pointer-events:none` so it never blocks
   *  cell interaction. */
  private commentPopup: HTMLDivElement;
  /** `"row:col"` → comment for the current sheet, rebuilt on every showSheet. */
  private commentMap = new Map<string, XlsxComment>();
  /** `"row:col"` of the cell whose popup is currently shown (or pending), so a
   *  pointermove within the same cell doesn't restart the show timer. */
  private commentPopupKey: string | null = null;
  /** Pending show timer (see {@link COMMENT_POPUP_DELAY_MS}). */
  private commentPopupTimer: ReturnType<typeof setTimeout> | null = null;

  // ─── List data-validation dropdown panel (display-only) ───────────────────
  /** DOM overlay listing a list-validated cell's allowed values. Lives in
   *  canvasArea above the scrollHost; unlike the comment popup this is a click
   *  target (`pointer-events:auto`). Read-only: hovering an item highlights it
   *  but selecting does NOT change the cell. */
  private validationPanel: HTMLDivElement;
  /** `"row:col"` of the cell whose panel is currently open, or null. Lets a
   *  re-click on the same arrow toggle the panel closed. */
  private validationPanelKey: string | null = null;
  /** Screen rect (canvasArea CSS px) of the dropdown arrow button last drawn by
   *  {@link maybeDrawValidationDropdown}, so pointerdown can hit-test it. Null
   *  when no arrow is currently visible. */
  private validationArrowRect: { x: number; y: number; w: number; h: number } | null = null;
  /** Document-level pointerdown listener that closes the panel on an outside
   *  click; installed only while the panel is open. */
  private validationOutsideHandler: ((e: PointerEvent) => void) | null = null;

  constructor(container: HTMLElement, opts: XlsxViewerOptions = {}) {
    this.opts = opts;
    this._mode = opts.mode ?? 'main';
    this._hiddenSheetMode = opts.hiddenSheetMode ?? 'show';

    this.wrapper = document.createElement('div');
    this.wrapper.style.cssText =
      `position:relative;width:100%;height:100%;` +
      `border:1px solid #c8ccd0;background:#fff;box-sizing:border-box;font-family:sans-serif;display:flex;flex-direction:column;`;

    this.canvasArea = document.createElement('div');
    this.canvasArea.style.cssText = `position:relative;flex:1;min-height:0;overflow:hidden;`;

    this.canvas = document.createElement('canvas');
    this.canvas.style.cssText = `position:absolute;top:0;left:0;z-index:0;display:block;`;
    // Worker mode paints worker-produced bitmaps; grab the bitmaprenderer
    // context once (a canvas can hold only one context type for its lifetime).
    if (this._mode === 'worker') {
      this._bitmapCtx = this.canvas.getContext('bitmaprenderer');
    }

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

    // Comment hover popup. z-index 3 sits above the scrollHost (z-index 2) so
    // it is visible over the grid; pointer-events:none keeps cell interaction
    // (selection, scroll) working through it. The zoom slider lives in a
    // sibling tab-bar flex child, so there is no stacking conflict.
    this.commentPopup = document.createElement('div');
    this.commentPopup.style.cssText =
      `position:absolute;z-index:3;pointer-events:none;display:none;` +
      `max-width:${COMMENT_POPUP_MAX_W}px;max-height:${COMMENT_POPUP_MAX_H}px;overflow:hidden;` +
      `box-sizing:border-box;padding:6px 8px;` +
      `background:#fffbcc;border:1px solid #b8b8a0;` +
      `box-shadow:1px 2px 5px rgba(0,0,0,0.25);` +
      `font:12px/1.4 sans-serif;color:#222;white-space:pre-wrap;word-break:break-word;`;

    // List-validation dropdown panel. z-index 4 sits above the comment popup
    // and scrollHost; pointer-events:auto because it IS a click target (the
    // user opens it by clicking the arrow and scrolls inside it). The wheel
    // handler below keeps that scroll from leaking to the grid.
    this.validationPanel = document.createElement('div');
    this.validationPanel.setAttribute('data-xlsx-validation-panel', '');
    this.validationPanel.style.cssText =
      `position:absolute;z-index:4;pointer-events:auto;display:none;` +
      `min-width:80px;max-width:${VALIDATION_PANEL_MAX_W}px;max-height:${VALIDATION_PANEL_MAX_H}px;overflow-y:auto;` +
      `box-sizing:border-box;background:#fff;border:1px solid #7f7f7f;` +
      `box-shadow:1px 2px 5px rgba(0,0,0,0.25);` +
      `font:12px/1.4 sans-serif;color:#222;`;
    // Keep a wheel inside the panel from scrolling the grid behind it. The panel
    // itself still scrolls (default action) up to its own bounds.
    this.validationPanel.addEventListener('wheel', (e) => e.stopPropagation());

    this.canvasArea.appendChild(this.canvas);
    this.canvasArea.appendChild(this.selectionOverlay);
    this.canvasArea.appendChild(this.scrollHost);
    this.canvasArea.appendChild(this.commentPopup);
    this.canvasArea.appendChild(this.validationPanel);

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
    // Inject the shared viewer stylesheet once per module (idempotent). Formerly
    // a per-instance `<style>` was appended to <head> on every construction,
    // leaking one node per mount/unmount cycle.
    ensureViewerStyleInjected();
    this.tabStrip.addEventListener('scroll', () => this.updateNavButtons());

    this.tabBar.appendChild(navGroup);
    this.tabBar.appendChild(this.tabStrip);
    if (this.opts.showZoomSlider !== false) {
      this.tabBar.appendChild(this.buildZoomControl());
    }

    this.wrapper.appendChild(this.canvasArea);
    this.wrapper.appendChild(this.tabBar);
    container.appendChild(this.wrapper);

    this.scrollHost.addEventListener('scroll', () => {
      // Any scroll cancels a deferred tap: the press that started it was a
      // scrollbar-thumb drag (overlay scrollbars) or a touch swipe, not a
      // cell click.
      this.pendingTap = null;
      // A comment popup is anchored to a cell's on-screen rect, which moves
      // under the cursor while scrolling — hide it (Excel does the same).
      this.hideCommentPopup();
      // The validation panel is anchored to the cell too; Excel closes its
      // dropdown on scroll, so do the same.
      this.hideValidationPanel();
      // Track the start-anchored position, but only while the host is laid
      // out: a hidden host reports clientWidth 0 and fires bogus scroll
      // events when the browser clamps scrollLeft, which must not overwrite
      // the last real position.
      if (this.scrollHost.clientWidth > 0) {
        this.effectiveH = this.effectiveScrollLeft;
      }
      // Coalesce into the next frame: a scroll gesture fires many events per
      // frame, and the previous synchronous redraw ran the full render on each
      // one. The overlay update is cheap DOM geometry (no canvas paint) and must
      // track the scroll immediately, so it stays synchronous.
      this.scheduleRender();
      this.updateSelectionOverlay();
    });

    // Re-render whenever the canvas area changes size. Re-anchor first: a
    // size change shifts maxScrollLeft, and for RTL sheets the native
    // scrollLeft must be re-derived from the start-anchored position or the
    // view drifts (or, after a hidden mount, stays stranded at the far end).
    this.resizeObserver = new ResizeObserver(() => {
      this.reanchorHorizontalScroll();
      // Container resizes can burst (a live window/pane drag); coalesce the
      // canvas paint into one frame. The re-anchor, overlay and nav updates are
      // cheap and must reflect the new size at once, so they stay synchronous.
      this.scheduleRender();
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
        mode: this._mode,
      });
      this.buildTabs();
      this.opts.onReady?.(this.wb.sheetNames);
      await this.showSheet(this._initialSheet());
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
    this.scrollHost.scrollTop = 0;
    this.anchorCell = null;
    this.activeCell = null;
    this.selectionMode = 'cells';
    this.hideCommentPopup();
    this.hideValidationPanel();
    this.updateSelectionOverlay();
    this.updateTabActive(index);
    this.currentWorksheet = await this.workbook.getWorksheet(index);
    this.buildCommentMap(this.currentWorksheet);
    this.updateSpacerSize(this.currentWorksheet);
    // Reset the horizontal scroll origin to the natural START of the sheet.
    // For RTL sheets the start column (col A) lives at the RIGHT, which means
    // the native scrollbar thumb must sit at its right end (max scrollLeft);
    // for LTR sheets the start is scrollLeft=0. updateSpacerSize must run first
    // so scrollWidth reflects the new sheet before we read the max offset.
    this.resetHorizontalScroll();
    await this.renderCurrentSheet();
    this.opts.onSheetChange?.(index, this.workbook.sheetNames.length);
  }

  /** True when the current sheet's grid is laid out right-to-left. */
  private get isRtl(): boolean {
    return this.currentWorksheet?.rightToLeft === true;
  }

  /** Maximum horizontal scroll offset the native scroll host allows (≥ 0). */
  private get maxScrollLeft(): number {
    return Math.max(0, this.scrollHost.scrollWidth - this.scrollHost.clientWidth);
  }

  /**
   * The logical horizontal scroll position used to find the start-of-sheet
   * (col A) edge, in *scaled* CSS pixels — the same unit as
   * `scrollHost.scrollLeft`. The renderer always lays the grid out LTR and then
   * mirrors it (ECMA-376 §18.3.1.87), so the viewer must hand it a position
   * where 0 = the START of the sheet (col A) and increasing values reveal later
   * columns.
   *
   * For LTR that is exactly the native `scrollLeft`. For RTL the sheet starts at
   * the RIGHT, so the native scrollbar runs the opposite way: thumb fully right
   * (`scrollLeft = maxScrollLeft`) is the start, thumb left is the far columns.
   * Inverting here makes wheel/trackpad follow the finger and aligns the
   * thumb↔page mapping with Excel, without depending on browser-specific RTL
   * `scrollLeft` sign conventions.
   */
  private get effectiveScrollLeft(): number {
    const raw = this.scrollHost.scrollLeft;
    return this.isRtl ? this.maxScrollLeft - raw : raw;
  }

  /**
   * Map between the logical-LTR x used by all the cell-geometry math and the
   * on-screen (canvasArea CSS-pixel) x, applying the RTL mirror (ECMA-376
   * §18.3.1.87) via the same {@link rtlMirrorX} the renderer uses. For LTR this
   * is the identity. The mirror is an involution, so this one method serves
   * both cell→px (overlay draw, `w` = cell width) and px→cell (pointer
   * hit-testing, `w` = 0 for a point) — guaranteeing the overlay sits exactly
   * where the cell is drawn and a click resolves to that same cell at every
   * scroll offset. `canvasArea.clientWidth` equals the renderer's `canvasW`.
   */
  private screenX(logicalX: number, w: number): number {
    return this.isRtl ? rtlMirrorX(logicalX, w, this.canvasArea.clientWidth) : logicalX;
  }

  /** Park the scrollbar at the sheet's natural start: scrollLeft=0 for LTR,
   *  the right end for RTL (so col A shows first). */
  private resetHorizontalScroll(): void {
    this.effectiveH = 0;
    this.scrollHost.scrollLeft = this.isRtl ? this.maxScrollLeft : 0;
  }

  /** Re-derive the native scrollLeft from the tracked start-anchored
   *  position after the scroll host's size changes. Only RTL needs this:
   *  for LTR the native scrollLeft *is* start-anchored and the browser
   *  already clamps it sensibly on resize. */
  private reanchorHorizontalScroll(): void {
    if (!this.isRtl || this.scrollHost.clientWidth === 0) return;
    const want = Math.max(0, this.maxScrollLeft - this.effectiveH);
    if (Math.abs(this.scrollHost.scrollLeft - want) > 1) {
      this.scrollHost.scrollLeft = want;
    }
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
    await this.goToSheet(this._stepSheet(1));
  }

  async prevSheet(): Promise<void> {
    await this.goToSheet(this._stepSheet(-1));
  }

  /** Next sheet index for sequential nav: skip mode jumps over hidden sheets. */
  private _stepSheet(dir: 1 | -1): number {
    if (this._hiddenSheetMode === 'skip' && this.wb) {
      return nextVisibleIndex(this.currentSheet, dir, (i) => this.wb!.isHidden(i), this.sheetCount);
    }
    return this.currentSheet + dir;
  }

  /** Initial sheet for load() / entering skip mode: land on a visible sheet. */
  private _initialSheet(): number {
    if (this._hiddenSheetMode === 'skip' && this.wb) {
      return resolveVisibleIndex(0, (i) => this.wb!.isHidden(i), this.sheetCount);
    }
    return 0;
  }

  /** Returns the cell at canvas-client coordinates, or null if outside the cell grid. */
  getCellAt(clientX: number, clientY: number): CellAddress | null {
    const ws = this.currentWorksheet;
    if (!ws) return null;
    const cs = this.opts.cellScale ?? 1;

    const rect = this.canvasArea.getBoundingClientRect();
    // Un-mirror the screen x into the logical-LTR layout the geometry below
    // assumes (header on the left). screenX is an involution, so applying it to
    // a screen point recovers the logical point; w = 0 for a point. Done in
    // scaled CSS px (canvasArea space) before converting to logical px.
    const lx = this.screenX(clientX - rect.left, 0) / cs;
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
      const axes = getSheetAxes(ws, getMdwForWorksheet(ws));
      const r = axes.row.scrollableIndexAt(contentY, freezeRows + 1);
      if (r === null) return null;
      row = r;
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
      const contentX = innerX - frozenW + this.effectiveScrollLeft / cs;
      const axes = getSheetAxes(ws, getMdwForWorksheet(ws));
      const c = axes.col.scrollableIndexAt(contentX, freezeCols + 1);
      if (c === null) return null;
      col = c;
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
      const logicalScrollX = this.effectiveScrollLeft / cs;
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
   * Programmatically select a single cell by A1 reference (e.g. `"B2"`), as if
   * the user had clicked it: updates the active/anchor cell, redraws the
   * selection overlay (including any list-validation dropdown arrow), and fires
   * `onSelectionChange`. A no-op for malformed refs. Closes any open validation
   * panel, matching the click path.
   */
  select(ref: string): void {
    const p = parseA1(ref);
    if (!p) return;
    this.hideValidationPanel();
    this.selectionMode = 'cells';
    this.anchorCell = { row: p.row, col: p.col };
    this.activeCell = { row: p.row, col: p.col };
    this.updateSelectionOverlay();
    void this.renderCurrentSheet();
    this.opts.onSelectionChange?.(this.selection);
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
    // Same RTL un-mirror as getCellAt: map the screen x back to the logical-LTR
    // layout (row header on the left) before the header math below.
    const lx = this.screenX(clientX - rect.left, 0) / cs;
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
      const axes = getSheetAxes(ws, getMdwForWorksheet(ws));
      const r = axes.row.scrollableIndexAt(contentY, freezeRows + 1);
      return r === null ? null : { kind: 'row', row: r };
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
    const contentX = innerX - frozenW + this.effectiveScrollLeft / cs;
    const axes = getSheetAxes(ws, getMdwForWorksheet(ws));
    const c = axes.col.scrollableIndexAt(contentX, freezeCols + 1);
    return c === null ? null : { kind: 'col', col: c };
  }

  /**
   * If the pointer sits on a column/row-header border (within {@link
   * RESIZE_GRAB_PX}), return the resize target: which index to resize and the
   * fixed LTR edge it grows from (in canvasArea CSS px). Excel resizes the band
   * whose *trailing* border you grab — the column to the left of a vertical
   * border, the row above a horizontal one — so both that band and its
   * neighbour-to-the-far-side are checked. Geometry comes straight from {@link
   * getCellRect}, so the grab line always coincides with the drawn border at any
   * scroll offset / zoom / RTL. Returns null off the header borders.
   */
  private getResizeTarget(
    clientX: number,
    clientY: number,
  ): { kind: 'col' | 'row'; index: number; originScaled: number; mdw: number } | null {
    const ws = this.currentWorksheet;
    if (!ws) return null;
    const cs = this.opts.cellScale ?? 1;
    const rect = this.canvasArea.getBoundingClientRect();
    // Un-mirror the screen x to the logical-LTR space getCellRect draws in (the
    // same transform getHeaderHit uses), so the comparison holds for RTL sheets.
    const ptX = this.screenX(clientX - rect.left, 0);
    const ptY = clientY - rect.top;
    const headerW = Math.round(HEADER_W * cs);
    const headerH = Math.round(HEADER_H * cs);
    const mdw = getMdwForWorksheet(ws);

    // Column borders live in the column-header strip, right of the corner.
    if (ptY <= headerH && ptX > headerW) {
      const hit = this.getHeaderHit(clientX, clientY);
      if (hit?.kind !== 'col') return null;
      const origins = new Map<number, number>(); // index -> fixed LTR origin edge
      const edges: { index: number; edge: number }[] = [];
      for (const c of [hit.col - 1, hit.col]) {
        if (c < 1) continue;
        const r = this.getCellRect(1, c); // x is independent of the row
        if (!r) continue;
        origins.set(c, r.x);
        edges.push({ index: c, edge: r.x + r.w }); // trailing (right) border
      }
      const index = resizeHitIndex(ptX, edges, RESIZE_GRAB_PX, headerW);
      if (index === null) return null;
      return { kind: 'col', index, originScaled: origins.get(index) as number, mdw };
    }

    // Row borders live in the row-header strip, below the corner.
    if (ptX <= headerW && ptY > headerH) {
      const hit = this.getHeaderHit(clientX, clientY);
      if (hit?.kind !== 'row') return null;
      const origins = new Map<number, number>(); // index -> fixed LTR origin edge
      const edges: { index: number; edge: number }[] = [];
      for (const rIdx of [hit.row - 1, hit.row]) {
        if (rIdx < 1) continue;
        const r = this.getCellRect(rIdx, 1); // y is independent of the column
        if (!r) continue;
        origins.set(rIdx, r.y);
        edges.push({ index: rIdx, edge: r.y + r.h }); // trailing (bottom) border
      }
      const index = resizeHitIndex(ptY, edges, RESIZE_GRAB_PX, headerH);
      if (index === null) return null;
      return { kind: 'row', index, originScaled: origins.get(index) as number, mdw };
    }

    return null;
  }

  /**
   * Apply a live resize drag: size the band from its fixed origin edge to the
   * current pointer, clamp to {@link RESIZE_MIN_PX}, and write the result back
   * into the in-memory worksheet model in its native unit (Excel column widths /
   * points). This is a *view-only* mutation — the file is never written. The
   * memoized axis cache for this sheet is invalidated so every geometry read
   * (spacer, hit-test, overlay, renderer) sees the new size on the next frame.
   */
  private applyResize(clientX: number, clientY: number): void {
    const drag = this.resizeDrag;
    const ws = this.currentWorksheet;
    if (!drag || !ws) return;
    const cs = this.opts.cellScale ?? 1;
    const rect = this.canvasArea.getBoundingClientRect();

    if (drag.kind === 'col') {
      const ptX = this.screenX(clientX - rect.left, 0);
      const sizePx = Math.max(RESIZE_MIN_PX, Math.round((ptX - drag.originScaled) / cs));
      ws.colWidths[drag.index] = pxToColWidth(sizePx, drag.mdw);
    } else {
      const ptY = clientY - rect.top;
      const sizePx = Math.max(RESIZE_MIN_PX, Math.round((ptY - drag.originScaled) / cs));
      ws.rowHeights[drag.index] = pxToRowHeight(sizePx);
    }

    sheetAxisCache.delete(ws); // sizes changed → rebuild the cumulative-offset axes
    this.updateSpacerSize(ws);
    this.updateSelectionOverlay();
    // Live resize drag fires per pointermove; coalesce the canvas repaint into
    // one frame. The spacer (scrollbar extent) and overlay updates are cheap DOM
    // writes that must track the drag immediately, so they stay synchronous.
    this.scheduleRender();
  }

  /**
   * Change the cell-selection highlight color at runtime (see {@link
   * XlsxViewerOptions.selectionColor}). The border takes the color as-is and the
   * fill becomes a translucent shade of it; the current selection repaints
   * immediately.
   */
  setSelectionColor(color: string): void {
    this.opts.selectionColor = color;
    this.updateSelectionOverlay();
  }

  /**
   * Switch the hidden-sheet mode at runtime: restyle the tabs and re-render.
   * Entering `'skip'` while on a hidden sheet advances to the nearest visible.
   */
  async setHiddenSheetMode(mode: HiddenSheetMode): Promise<void> {
    this._hiddenSheetMode = mode;
    this.buildTabs();
    if (mode === 'skip' && this.wb && this.wb.isHidden(this.currentSheet)) {
      await this.showSheet(
        resolveVisibleIndex(this.currentSheet, (i) => this.wb!.isHidden(i), this.sheetCount),
      );
    } else {
      this.updateTabActive(this.currentSheet);
    }
  }

  /** The current hidden-sheet mode. */
  get hiddenSheetMode(): HiddenSheetMode { return this._hiddenSheetMode; }

  /** Number of non-hidden sheets (absolute `sheetCount` is unchanged). */
  get visibleSheetCount(): number {
    if (!this.wb) return 0;
    const wb = this.wb;
    return countVisible((i) => wb.isHidden(i), this.sheetCount);
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

    // The rect above is in the logical-LTR layout (header on the left), the
    // same space getCellRect / the all|rows|cols branches use. For an RTL sheet
    // the renderer mirrors every cell about canvasW (ECMA-376 §18.3.1.87), so
    // mirror the final [x, x+w] band once with the same transform the renderer
    // uses. Doing it here — after the LTR clamps — keeps the header/frozen
    // clamping correct and guarantees the overlay lands exactly on the drawn
    // cell at every scroll offset (cell→px uses the same map as px→cell above).
    const screenLeft = this.screenX(x, w);

    const { border, background } = selectionOverlayStyle(
      this.opts.selectionColor ?? DEFAULT_SELECTION_COLOR,
    );
    const box = document.createElement('div');
    box.style.cssText =
      `position:absolute;` +
      `left:${screenLeft}px;top:${y}px;width:${w}px;height:${h}px;` +
      `box-sizing:border-box;border:${border};` +
      `background:${background};pointer-events:none;`;
    this.selectionOverlay.appendChild(box);

    // List data-validation dropdown arrow (ECMA-376 §18.3.1.33). Excel shows an
    // in-cell dropdown button only while the cell is *selected* and only for
    // `list`-type rules — so it is drawn here (selection overlay) rather than in
    // the canvas renderer. The button itself is non-interactive
    // (pointer-events:none); clicks are hit-tested against its rect in the
    // pointerdown handler, which opens a panel listing the allowed values
    // (display only — picking a value never changes the cell).
    this.maybeDrawValidationDropdown();
  }

  /** Draw the Excel list-validation dropdown button just outside the
   *  bottom-right corner of the *active* cell when that cell is covered by a
   *  `list` data-validation rule. Anchored to the single active cell (not the
   *  whole range) to mirror Excel, which attaches the button to the active
   *  cell of the selection. */
  private maybeDrawValidationDropdown(): void {
    // The overlay is rebuilt on every selection / scroll change, so the
    // arrow's hit-test rect is recomputed here each time (cleared when no arrow
    // is currently shown).
    this.validationArrowRect = null;
    if (this.selectionMode !== 'cells') return;
    const ws = this.currentWorksheet;
    const active = this.activeCell;
    if (!ws || !active) return;
    const dv = findListValidationAt(ws.dataValidations, active.row, active.col);
    if (!dv) return;

    const rect = this.getCellRect(active.row, active.col);
    if (!rect) return;

    // Excel's dropdown button is a fixed square sized to the cell height,
    // clamped to a sensible range so it stays usable at small zoom and doesn't
    // dominate tall rows. The arrow glyph is centered inside.
    const cs = this.opts.cellScale ?? 1;
    const headerW = Math.round(HEADER_W * cs);
    const headerH = Math.round(HEADER_H * cs);
    const side = Math.max(14, Math.min(rect.h, 22 * cs));
    // Button sits flush to the right of the cell, top-aligned with it.
    const btnLogicalX = rect.x + rect.w;
    const btnY = rect.y;
    // Cull when the active cell (hence its button) is scrolled behind the
    // fixed headers.
    if (btnLogicalX + side <= headerW || btnY + side <= headerH) return;

    const screenLeft = this.screenX(btnLogicalX, side);

    const btn = document.createElement('div');
    btn.setAttribute('data-xlsx-validation-dropdown', '');
    btn.style.cssText =
      `position:absolute;` +
      `left:${screenLeft}px;top:${btnY}px;width:${side}px;height:${side}px;` +
      `box-sizing:border-box;display:flex;align-items:center;justify-content:center;` +
      // Match Excel's grey button chrome; non-interactive (display only).
      `background:#f0f0f0;border:1px solid #7f7f7f;pointer-events:none;`;
    const arrow = Math.max(4, Math.round(side * 0.42));
    btn.innerHTML =
      `<svg width="${arrow}" height="${arrow}" viewBox="0 0 10 6" aria-hidden="true">` +
      `<path d="M0 0 L10 0 L5 6 Z" fill="#333"/></svg>`;
    this.selectionOverlay.appendChild(btn);

    // Record the arrow's on-screen rect (canvasArea space) for pointer
    // hit-testing. The button element has pointer-events:none, so clicks fall
    // through to the scrollHost where the pointerdown handler tests this rect.
    this.validationArrowRect = { x: screenLeft, y: btnY, w: side, h: side };

    // Keep an already-open panel glued to the arrow as the grid scrolls. If the
    // active cell's validation differs from the open panel (selection moved),
    // close it instead.
    if (this.validationPanel.style.display !== 'none') {
      if (this.validationPanelKey === `${active.row}:${active.col}`) {
        this.positionValidationPanel();
      } else {
        this.hideValidationPanel();
      }
    }
  }

  // ─── List data-validation dropdown panel (display-only) ───────────────────

  /** Toggle the dropdown panel for the active cell's list validation. Called
   *  from pointerdown when the arrow rect is hit. Re-clicking the same arrow
   *  closes it. */
  private toggleValidationPanel(): void {
    const ws = this.currentWorksheet;
    const active = this.activeCell;
    if (!ws || !active) return;
    const key = `${active.row}:${active.col}`;
    if (this.validationPanelKey === key && this.validationPanel.style.display !== 'none') {
      this.hideValidationPanel();
      return;
    }
    const dv = findListValidationAt(ws.dataValidations, active.row, active.col);
    if (!dv) return;
    void this.openValidationPanel(active, dv.formula1);
  }

  /** Resolve the allowed values for `formula1` (relative to the current sheet)
   *  and render them in the panel anchored below the active cell. Async because
   *  cross-sheet range references may need a lazily-parsed worksheet. */
  private async openValidationPanel(cell: CellAddress, formula1: string | undefined): Promise<void> {
    let resolved: ResolvedList;
    try {
      resolved = await this.workbook.resolveValidationList(this.currentSheet, formula1);
    } catch {
      // A resolution failure (e.g. a missing sheet) must not break the viewer;
      // fall back to disclosing the raw formula.
      resolved = { kind: 'formula', formula: formula1 ?? '' };
    }
    // The selection may have moved while awaiting — bail if so.
    const active = this.activeCell;
    if (!active || active.row !== cell.row || active.col !== cell.col) return;

    this.validationPanelKey = `${cell.row}:${cell.col}`;
    this.renderValidationPanel(resolved);
    this.positionValidationPanel();
    this.installValidationOutsideHandler();
  }

  /** Build the panel's children. Uses textContent throughout (no HTML injection
   *  from cell values). Items highlight on hover but are NOT selectable —
   *  this is a read-only viewer, so clicking a value must not change the cell. */
  private renderValidationPanel(resolved: ResolvedList): void {
    const panel = this.validationPanel;
    panel.textContent = '';
    if (resolved.kind === 'formula' || resolved.values.length === 0) {
      // Unresolved operand (named range / complex formula) or an empty range:
      // disclose the formula / a placeholder rather than showing a blank box.
      const note = document.createElement('div');
      note.style.cssText = 'padding:4px 8px;color:#666;font-style:italic;white-space:pre-wrap;word-break:break-word;';
      note.textContent =
        resolved.kind === 'formula'
          ? (resolved.formula ? `= ${resolved.formula}` : '(no list)')
          : '(empty list)';
      panel.appendChild(note);
      return;
    }
    for (const value of resolved.values) {
      const item = document.createElement('div');
      item.setAttribute('data-xlsx-validation-item', '');
      item.style.cssText = 'padding:3px 8px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;cursor:default;';
      item.textContent = value;
      // Hover highlight only — no click/select (read-only viewer).
      item.addEventListener('pointerenter', () => {
        item.style.background = '#cfe3ff';
      });
      item.addEventListener('pointerleave', () => {
        item.style.background = '';
      });
      panel.appendChild(item);
    }
  }

  /** Position the (already-populated, visible-or-becoming-visible) panel below
   *  the dropdown arrow / active cell using the pure geometry calculator. */
  private positionValidationPanel(): void {
    const active = this.activeCell;
    if (!active) return;
    const rect = this.getCellRect(active.row, active.col);
    if (!rect) return;
    const screenLeft = this.screenX(rect.x, rect.w);
    // Make it measurable off-screen first so offsetWidth/Height reflect content.
    this.validationPanel.style.left = '-9999px';
    this.validationPanel.style.top = '-9999px';
    this.validationPanel.style.display = 'block';
    const pos = computeValidationPanelPosition({
      cell: { x: screenLeft, y: rect.y, w: rect.w, h: rect.h },
      panel: { w: this.validationPanel.offsetWidth, h: this.validationPanel.offsetHeight },
      viewport: { w: this.canvasArea.clientWidth, h: this.canvasArea.clientHeight },
      rtl: this.isRtl,
    });
    this.validationPanel.style.left = `${pos.left}px`;
    this.validationPanel.style.top = `${pos.top}px`;
  }

  /** Install a document-level pointerdown listener that closes the panel on a
   *  click outside it (and outside the arrow, which toggles via its own path).
   *  Removed by {@link hideValidationPanel}. */
  private installValidationOutsideHandler(): void {
    if (this.validationOutsideHandler) return;
    this.validationOutsideHandler = (e: PointerEvent) => {
      const target = e.target as Node | null;
      if (target && this.validationPanel.contains(target)) return; // inside panel
      // A click on the arrow is handled by the scrollHost pointerdown (toggle);
      // don't double-handle it here. Detect by hit-testing the arrow rect.
      const rect = this.canvasArea.getBoundingClientRect();
      const ax = e.clientX - rect.left;
      const ay = e.clientY - rect.top;
      const ar = this.validationArrowRect;
      if (ar && ax >= ar.x && ax <= ar.x + ar.w && ay >= ar.y && ay <= ar.y + ar.h) {
        return;
      }
      this.hideValidationPanel();
    };
    // Capture phase so we see the click before it mutates selection.
    document.addEventListener('pointerdown', this.validationOutsideHandler, true);
  }

  /** Hide the panel and detach its outside-click listener. Called on re-click,
   *  outside click, Esc, scroll, selection change, sheet switch and destroy. */
  private hideValidationPanel(): void {
    this.validationPanel.style.display = 'none';
    this.validationPanelKey = null;
    if (this.validationOutsideHandler) {
      document.removeEventListener('pointerdown', this.validationOutsideHandler, true);
      this.validationOutsideHandler = null;
    }
  }

  // ─── Comment hover popup ──────────────────────────────────────────────────

  /** Build the `"row:col"` → comment index for the given sheet. Parses each
   *  `XlsxComment.cellRef` with the shared {@link parseA1}; later refs win on a
   *  collision (Excel allows at most one note per cell, so this is moot in
   *  practice). */
  private buildCommentMap(ws: Worksheet): void {
    this.commentMap = new Map();
    for (const c of ws.comments ?? []) {
      const p = parseA1(c.cellRef);
      if (p) this.commentMap.set(`${p.row}:${p.col}`, c);
    }
  }

  /** Show the popup for the comment on `cell` after the hover dwell, anchored to
   *  the cell's current on-screen rect. No-op when the cell carries no comment.
   *  Re-hovering the same cell does not restart the timer. */
  private scheduleCommentPopup(cell: CellAddress): void {
    const key = `${cell.row}:${cell.col}`;
    const comment = this.commentMap.get(key);
    if (!comment) {
      this.hideCommentPopup();
      return;
    }
    if (this.commentPopupKey === key) return; // already shown / pending here
    this.hideCommentPopup();
    this.commentPopupKey = key;
    this.commentPopupTimer = setTimeout(() => {
      this.commentPopupTimer = null;
      this.renderCommentPopup(cell, comment);
    }, COMMENT_POPUP_DELAY_MS);
  }

  /** Immediately render the popup for `comment` anchored to `cell` (used by the
   *  hover-dwell timer and by touch selection, which has no hover). */
  private renderCommentPopup(cell: CellAddress, comment: XlsxComment): void {
    const rect = this.getCellRect(cell.row, cell.col);
    if (!rect) return;

    // Author on its own bold line (when present), then the body text with
    // newlines preserved. textContent escapes everything — no HTML injection
    // from comment text.
    this.commentPopup.textContent = '';
    if (comment.author) {
      const authorEl = document.createElement('div');
      authorEl.style.cssText = 'font-weight:bold;margin-bottom:2px;';
      authorEl.textContent = comment.author;
      this.commentPopup.appendChild(authorEl);
    }
    const bodyEl = document.createElement('div');
    bodyEl.textContent = comment.text;
    this.commentPopup.appendChild(bodyEl);

    // Anchor to the cell's *screen* rect (RTL already mirrored by screenX), then
    // run the pure position calc against the popup's measured size. Make it
    // visible (off-screen) first so offsetWidth/Height reflect the wrapped text.
    const screenLeft = this.screenX(rect.x, rect.w);
    this.commentPopup.style.left = '-9999px';
    this.commentPopup.style.top = '-9999px';
    this.commentPopup.style.display = 'block';

    const pos = computeCommentPopupPosition({
      cell: { x: screenLeft, y: rect.y, w: rect.w, h: rect.h },
      popup: { w: this.commentPopup.offsetWidth, h: this.commentPopup.offsetHeight },
      viewport: { w: this.canvasArea.clientWidth, h: this.canvasArea.clientHeight },
      rtl: this.isRtl,
    });
    this.commentPopup.style.left = `${pos.left}px`;
    this.commentPopup.style.top = `${pos.top}px`;
  }

  /** Hide the popup and cancel any pending show. Called on cell-out, scroll,
   *  sheet switch and destroy. */
  private hideCommentPopup(): void {
    if (this.commentPopupTimer !== null) {
      clearTimeout(this.commentPopupTimer);
      this.commentPopupTimer = null;
    }
    this.commentPopupKey = null;
    this.commentPopup.style.display = 'none';
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

      // Drag-to-resize a column/row from its header border (issue #567). Checked
      // before selection so grabbing the border never moves the cell selection.
      // Gated by the `resizable` option (default true); when off, a header-border
      // press falls through to normal selection behavior.
      const resize = (this.opts.resizable ?? true)
        ? this.getResizeTarget(e.clientX, e.clientY)
        : null;
      if (resize) {
        e.preventDefault();
        this.resizeDrag = { ...resize, pointerId: e.pointerId };
        this.scrollHost.setPointerCapture(e.pointerId);
        this.hideCommentPopup();
        return;
      }

      // List-validation dropdown arrow: if the press lands on the (display-only)
      // arrow button drawn on the active cell, toggle the value panel instead of
      // re-selecting the cell. The arrow's rect is in canvasArea space, so map
      // the client point through canvasArea's box.
      const ar = this.validationArrowRect;
      if (ar) {
        const areaRect = this.canvasArea.getBoundingClientRect();
        const ax = e.clientX - areaRect.left;
        const ay = e.clientY - areaRect.top;
        if (ax >= ar.x && ax <= ar.x + ar.w && ay >= ar.y && ay <= ar.y + ar.h) {
          e.preventDefault();
          this.toggleValidationPanel();
          return;
        }
      }

      // A pointerdown on the native scrollbar must not move the cell
      // selection — dragging the thumb would otherwise select whatever cell
      // sits underneath it. Two scrollbar styles need different handling:
      // classic scrollbars reserve layout space, so the press lands in the
      // band between the content box (clientWidth/Height) and the border-box
      // edge and can be rejected exactly; OS overlay scrollbars (macOS
      // "show when scrolling") float over the content without affecting
      // client sizes, so a press near a scrollable edge is geometrically
      // indistinguishable from a cell click. For that case we defer the
      // selection to pointerup via the pendingTap path and cancel it when a
      // scroll event arrives first (the press was a thumb drag). A plain
      // click in the band still selects the cell on release.
      const hostRect = this.scrollHost.getBoundingClientRect();
      const localX = e.clientX - hostRect.left - this.scrollHost.clientLeft;
      const localY = e.clientY - hostRect.top - this.scrollHost.clientTop;
      if (localX >= this.scrollHost.clientWidth || localY >= this.scrollHost.clientHeight) {
        return; // classic scrollbar gutter
      }
      // Overlay scrollbar hit band (~15 CSS px on macOS / Windows 11).
      const OVERLAY_SCROLLBAR_BAND = 16;
      const inOverlayBand =
        (this.scrollHost.scrollWidth > this.scrollHost.clientWidth &&
          this.scrollHost.clientHeight - localY <= OVERLAY_SCROLLBAR_BAND) ||
        (this.scrollHost.scrollHeight > this.scrollHost.clientHeight &&
          this.scrollHost.clientWidth - localX <= OVERLAY_SCROLLBAR_BAND);

      // Touch / pen: defer selection until pointerup so swipe-to-scroll doesn't change the cell.
      // Mouse: select immediately to preserve drag-to-extend behavior.
      if (e.pointerType !== 'mouse' || inOverlayBand) {
        this.pendingTap = { x: e.clientX, y: e.clientY, shiftKey: e.shiftKey, pointerId: e.pointerId };
        return;
      }

      this.applyPointerSelection(e.clientX, e.clientY, e.shiftKey, e.pointerId, true);
    });

    this.scrollHost.addEventListener('pointermove', (e: PointerEvent) => {
      // Live column/row resize takes priority over every other pointer behavior.
      if (this.resizeDrag && this.resizeDrag.pointerId === e.pointerId) {
        e.preventDefault();
        this.applyResize(e.clientX, e.clientY);
        return;
      }

      // Resize-handle affordance: show the col/row-resize cursor when hovering a
      // header border (mouse only — touch/pen have no hover). Skipped mid-select
      // and when the `resizable` option (default true) is off, so no resize
      // cursor is shown when drag-resize is disabled.
      if (e.pointerType === 'mouse' && !this.isSelecting && (this.opts.resizable ?? true)) {
        const rt = this.getResizeTarget(e.clientX, e.clientY);
        this.scrollHost.style.cursor = rt ? (rt.kind === 'col' ? 'col-resize' : 'row-resize') : '';
        if (rt) {
          this.hideCommentPopup();
          return;
        }
      }

      // Cancel a pending tap once the pointer moves beyond the slop — the user is scrolling.
      if (this.pendingTap && this.pendingTap.pointerId === e.pointerId) {
        const dx = e.clientX - this.pendingTap.x;
        const dy = e.clientY - this.pendingTap.y;
        if (dx * dx + dy * dy > TAP_SLOP * TAP_SLOP) {
          this.pendingTap = null;
        }
      }

      // Comment hover popup (mouse only — touch/pen have no hover, so they get
      // the popup on selection instead, below). Suppressed while drag-selecting
      // so the popup doesn't fight the selection rect. A header hover hides it.
      if (e.pointerType === 'mouse' && !this.isSelecting) {
        const hovered = this.getCellAt(e.clientX, e.clientY);
        if (hovered) this.scheduleCommentPopup(hovered);
        else this.hideCommentPopup();
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
      // Drag-select fires per pointermove; coalesce the canvas repaint (the
      // header-highlight bands the renderer draws) into one frame. The overlay
      // rect and the selection-change callback stay synchronous.
      this.scheduleRender();
      this.opts.onSelectionChange?.(this.selection);
    });

    this.scrollHost.addEventListener('pointerup', (e: PointerEvent) => {
      if (this.resizeDrag && this.resizeDrag.pointerId === e.pointerId) {
        this.scrollHost.releasePointerCapture(e.pointerId);
        this.resizeDrag = null;
        return;
      }
      if (this.pendingTap && this.pendingTap.pointerId === e.pointerId) {
        const dx = e.clientX - this.pendingTap.x;
        const dy = e.clientY - this.pendingTap.y;
        if (dx * dx + dy * dy <= TAP_SLOP * TAP_SLOP) {
          this.applyPointerSelection(e.clientX, e.clientY, this.pendingTap.shiftKey, e.pointerId, false);
          // Touch / pen have no hover, so surface the comment popup on a tap
          // (the active cell after the selection commit). Mouse uses hover.
          if (e.pointerType !== 'mouse' && this.activeCell) {
            const key = `${this.activeCell.row}:${this.activeCell.col}`;
            const comment = this.commentMap.get(key);
            if (comment) {
              this.hideCommentPopup();
              this.renderCommentPopup(this.activeCell, comment);
            } else {
              this.hideCommentPopup();
            }
          }
        }
        this.pendingTap = null;
      }
      this.isSelecting = false;
    });

    this.scrollHost.addEventListener('pointercancel', (e: PointerEvent) => {
      if (this.resizeDrag && this.resizeDrag.pointerId === e.pointerId) {
        this.resizeDrag = null;
      }
      if (this.pendingTap && this.pendingTap.pointerId === e.pointerId) {
        this.pendingTap = null;
      }
      this.isSelecting = false;
    });

    // Ctrl/⌘ + mouse wheel (and trackpad pinch, which the browser reports as a
    // ctrl-wheel) zooms the grid, matching Excel. preventDefault stops the
    // browser's own page zoom. A plain wheel still scrolls the grid natively.
    // The step is exponential in deltaY (see zoomStepScale) so a trackpad
    // pinch — a high-frequency stream of small-deltaY events — does not zoom
    // away; the total zoom tracks the gesture distance, not the event count.
    this.scrollHost.addEventListener(
      'wheel',
      (e: WheelEvent) => {
        if (!(e.ctrlKey || e.metaKey)) return;
        e.preventDefault();
        if (e.deltaY === 0) return;
        this.setScale(zoomStepScale(this.opts.cellScale ?? 1, e.deltaY));
      },
      { passive: false },
    );

    // Hide the comment popup when the cursor leaves the grid entirely.
    this.scrollHost.addEventListener('pointerleave', () => this.hideCommentPopup());

    this.keydownHandler = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
        this.copySelection();
      } else if (e.key === 'Escape' && this.validationPanel.style.display !== 'none') {
        this.hideValidationPanel();
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
      btn.style.cssText = this.tabCss(i, false);
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
      btn.style.cssText = this.tabCss(i, i === index);
    });
    // Keep the active tab visible by scrolling the tab strip HORIZONTALLY only.
    // `scrollIntoView` walks every scrollable ancestor, so it also scrolls the
    // page vertically — on first load that jumped the whole page down to the
    // tab bar (the active sheet is set during load). Adjust the strip's
    // scrollLeft directly so the page never moves.
    // `offsetParent === null` for a `display:none` tab (a hidden sheet reached
    // by an explicit goToSheet in 'skip' mode). Its getBoundingClientRect is all
    // zeros, which would spuriously scroll the strip — skip the scroll for it.
    const tab = this.tabs[index];
    if (tab && tab.offsetParent !== null) {
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

  /**
   * Full inline style for the tab of sheet `i`, honoring the hidden-sheet mode:
   * `'skip'` hides the tab of a hidden/veryHidden sheet (`display:none`); `'dim'`
   * greys it but leaves it clickable; `'show'` styles every tab normally. Used
   * by both buildTabs and updateTabActive so navigation never wipes the styling.
   */
  private tabCss(i: number, active: boolean): string {
    let css = this.tabStyle(active, this.tabColors[i]);
    if (this._hiddenSheetMode !== 'show' && this.wb?.isHidden(i)) {
      css += this._hiddenSheetMode === 'skip' ? 'display:none;' : `opacity:${HIDDEN_TAB_DIM_OPACITY};`;
    }
    return css;
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

    if (this.currentWorksheet) {
      // Preserve the START-anchored effective scroll position across the zoom.
      // The spacer (scrollWidth) is re-sized below, which changes maxScrollLeft;
      // for RTL the native scrollLeft is the inverse of the effective position,
      // so we must re-derive scrollLeft from the preserved effective value or
      // the view would jump toward the start on every zoom step.
      const prevEffective = this.effectiveScrollLeft;
      this.updateSpacerSize(this.currentWorksheet);
      this.effectiveH = prevEffective;
      if (this.isRtl) {
        this.scrollHost.scrollLeft = Math.max(0, this.maxScrollLeft - prevEffective);
      }
    }
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

  /**
   * Coalesce a re-render into the next animation frame. Called from the
   * high-frequency event-driven paths (scroll, live column/row resize, drag-
   * selection, container resize); a burst of these within one frame schedules a
   * single {@link renderCurrentSheet}, avoiding the previous behavior where every
   * scroll event forced its own synchronous full redraw. Already-scheduled frames
   * are not re-scheduled — the one pending render reads the live scroll/scale
   * state when it runs, so the most recent position always wins without threading
   * a coordinate through. Falls back to a synchronous render when
   * `requestAnimationFrame` is unavailable (e.g. a non-DOM host), preserving the
   * old semantics there.
   */
  private scheduleRender(): void {
    if (this._rafId !== null) return;
    if (typeof requestAnimationFrame !== 'function') {
      void this.renderCurrentSheet();
      return;
    }
    this._rafId = requestAnimationFrame(() => {
      this._rafId = null;
      void this.renderCurrentSheet();
    });
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
    // Convert to logical pixels for cell-finding by dividing by cs. For RTL
    // sheets effectiveScrollLeft inverts the native scrollLeft so that 0 = col A
    // at the (mirrored) right edge — see the getter for the rationale.
    const logicalScrollX = this.effectiveScrollLeft / cs;
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

    const renderOpts = {
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
    };

    if (this._mode === 'worker') {
      // Render the viewport off the main thread and paint the returned bitmap.
      // The selection overlay (geometry-based, from getCellRect) is unaffected.
      const bmp = await this.workbook.renderViewportToBitmap(this.currentSheet, viewport, renderOpts);
      // Resize the canvas only when the bitmap dimensions actually change.
      // Re-assigning canvas.width/height re-allocates the GPU backing store even
      // when the value is identical, which on a steady scroll stream (same size
      // every frame) is a wasted allocation per frame (improvement plan C4).
      // transferFromImageBitmap replaces the whole canvas, so the resize's
      // implicit clear is not relied upon; skipping the no-op resize is safe.
      if (this.canvas.width !== bmp.width) this.canvas.width = bmp.width;
      if (this.canvas.height !== bmp.height) this.canvas.height = bmp.height;
      const cssW = `${w}px`;
      const cssH = `${h}px`;
      if (this.canvas.style.width !== cssW) this.canvas.style.width = cssW;
      if (this.canvas.style.height !== cssH) this.canvas.style.height = cssH;
      this._bitmapCtx?.transferFromImageBitmap(bmp);
    } else {
      await this.workbook.renderViewport(this.canvas, this.currentSheet, viewport, renderOpts);
    }
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

  /**
   * Tear down the viewer and release resources.
   *
   * The caller's container is returned to the state it had before construction
   * (empty): the entire wrapper subtree the constructor appended is removed.
   * All document-level listeners are detached — the keydown handler here, and
   * the validation-panel outside-click handler via {@link hideValidationPanel}.
   * Listeners on elements inside the wrapper (scrollHost, tabs, …) need no
   * explicit removal: removing the subtree makes them unreachable and eligible
   * for GC. Safe to call more than once.
   *
   * NOTE: the shared `<style>` in `document.head` is intentionally NOT removed —
   * it is a class constant that any still-live viewer may depend on, and one
   * leftover sheet is a bounded, harmless cost (see {@link ensureViewerStyleInjected}).
   */
  destroy(): void {
    this.resizeObserver?.disconnect();
    // Cancel any coalesced render still queued for the next frame so it can't
    // fire against a torn-down viewer (matches the destroy-completeness flow:
    // no scheduled work outlives destroy()).
    if (this._rafId !== null && typeof cancelAnimationFrame === 'function') {
      cancelAnimationFrame(this._rafId);
      this._rafId = null;
    }
    this.hideCommentPopup();
    this.hideValidationPanel();
    if (this.keydownHandler) {
      document.removeEventListener('keydown', this.keydownHandler);
    }
    this.wb?.destroy();
    // Remove the whole UI subtree so the container is empty again. This also
    // detaches every listener bound to elements within it (scrollHost pointer/
    // wheel handlers, tab clicks, zoom slider) without per-element cleanup.
    this.wrapper.remove();
  }
}
