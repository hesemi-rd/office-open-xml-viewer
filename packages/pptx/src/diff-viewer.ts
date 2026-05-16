import type { Presentation } from './types';
import { PptxPresentation } from './presentation';

/** Structural shape mirroring `@silurus/ooxml-diff` — duplicated here so the
 *  pptx package stays free of a runtime dependency on the diff package while
 *  remaining type-compatible with anything it returns. */
interface BBox { x: number; y: number; width: number; height: number; }
type ChangeLocation =
  | { kind: 'slide'; slideIndex: number; bbox?: BBox }
  | { kind: 'page'; pageIndex: number; bbox?: BBox }
  | { kind: 'paragraph'; paragraphIndex: number }
  | { kind: 'table'; row: number; col: number }
  | { kind: 'cell'; sheetName: string; row: number; col: number }
  | { kind: 'sheet'; sheetName: string };
type DiffChange = {
  op: 'add' | 'remove' | 'modify';
  path: string;
  kind: string;
  before?: unknown;
  after?: unknown;
  location?: ChangeLocation;
};
type DiffResult = {
  format: 'pptx' | 'docx' | 'xlsx';
  changes: DiffChange[];
};
type DiffFn = (a: Presentation, b: Presentation) => DiffResult;

export interface PptxDiffViewerOptions {
  /** Display width in CSS pixels for each side. Defaults to canvas.offsetWidth or 600. */
  width?: number;
  /** Device pixel ratio. Defaults to window.devicePixelRatio or 1. */
  dpr?: number;
  /** Use Google Fonts for theme webfonts (privacy implications — see PptxPresentation.load). */
  useGoogleFonts?: boolean;
  /** Called when the active slide changes. */
  onSlideChange?: (index: number, total: number) => void;
  /** Called when the diff completes after both sources load. */
  onDiff?: (result: DiffResult) => void;
  /** Overlay colours. */
  colors?: {
    add?: string;     // default '#22c55e' (green)
    remove?: string;  // default '#ef4444' (red)
    modify?: string;  // default '#f59e0b' (amber)
  };
}

/**
 * Side-by-side PPTX diff viewer.
 *
 * Renders two presentations into two canvases, computes a structural diff, and
 * overlays per-change bounding-box highlights so the eye can spot what moved,
 * changed text, or disappeared.
 *
 * Inject the diff implementation via `setDiffFn` (this keeps the pptx package
 * independent of `@silurus/ooxml-diff` at build time).
 */
export class PptxDiffViewer {
  private readonly left: HTMLCanvasElement;
  private readonly right: HTMLCanvasElement;
  private readonly leftOverlay: HTMLCanvasElement;
  private readonly rightOverlay: HTMLCanvasElement;
  private engineLeft: PptxPresentation | null = null;
  private engineRight: PptxPresentation | null = null;
  private readonly opts: PptxDiffViewerOptions;
  private diffFn: DiffFn | null = null;
  private result: DiffResult | null = null;
  private currentSlide = 0;

  constructor(leftCanvas: HTMLCanvasElement, rightCanvas: HTMLCanvasElement, opts: PptxDiffViewerOptions = {}) {
    this.opts = opts;
    this.left = leftCanvas;
    this.right = rightCanvas;

    this.leftOverlay = this._wrapWithOverlay(leftCanvas);
    this.rightOverlay = this._wrapWithOverlay(rightCanvas);
  }

  /** Inject the diff implementation. Pass `diffPptx` from `@silurus/ooxml-diff`. */
  setDiffFn(fn: (a: Presentation, b: Presentation) => unknown): void {
    this.diffFn = fn as DiffFn;
  }

  /** Load both before / after presentations and render the first slide. */
  async load(before: string | ArrayBuffer, after: string | ArrayBuffer): Promise<void> {
    this.engineLeft = await PptxPresentation.load(before, { useGoogleFonts: this.opts.useGoogleFonts });
    this.engineRight = await PptxPresentation.load(after, { useGoogleFonts: this.opts.useGoogleFonts });
    this._computeDiff();
    this.currentSlide = 0;
    await this._renderBoth();
  }

  get slideCount(): number {
    return Math.max(this.engineLeft?.slideCount ?? 0, this.engineRight?.slideCount ?? 0);
  }

  get slideIndex(): number { return this.currentSlide; }

  get diffResult(): DiffResult | null { return this.result; }

  async goToSlide(index: number): Promise<void> {
    if (this.slideCount === 0) return;
    this.currentSlide = Math.max(0, Math.min(index, this.slideCount - 1));
    await this._renderBoth();
  }

  async nextSlide(): Promise<void> { await this.goToSlide(this.currentSlide + 1); }
  async prevSlide(): Promise<void> { await this.goToSlide(this.currentSlide - 1); }

  destroy(): void {
    this.engineLeft?.destroy();
    this.engineRight?.destroy();
    this.leftOverlay.parentElement?.remove();
    this.rightOverlay.parentElement?.remove();
  }

  private _wrapWithOverlay(canvas: HTMLCanvasElement): HTMLCanvasElement {
    const parent = canvas.parentElement;
    const wrapper = document.createElement('div');
    wrapper.style.cssText = 'position:relative;display:inline-block;vertical-align:top;';
    if (!canvas.style.display) canvas.style.display = 'block';
    if (parent) parent.insertBefore(wrapper, canvas);
    wrapper.appendChild(canvas);
    const overlay = document.createElement('canvas');
    overlay.style.cssText = 'position:absolute;top:0;left:0;width:100%;height:100%;pointer-events:none;';
    wrapper.appendChild(overlay);
    return overlay;
  }

  private _computeDiff(): void {
    if (!this.engineLeft || !this.engineRight || !this.diffFn) {
      this.result = null;
      return;
    }
    const a = this.engineLeft.presentation;
    const b = this.engineRight.presentation;
    if (!a || !b) return;
    this.result = this.diffFn(a, b);
    this.opts.onDiff?.(this.result);
  }

  private async _renderBoth(): Promise<void> {
    const w = this.opts.width ?? this.left.offsetWidth ?? 600;
    const dpr = this.opts.dpr ?? window.devicePixelRatio ?? 1;

    if (this.engineLeft && this.currentSlide < this.engineLeft.slideCount) {
      await this.engineLeft.renderSlide(this.left, this.currentSlide, { width: w, dpr });
      this._paintOverlay(this.leftOverlay, this.engineLeft, 'remove');
    } else {
      const ctx = this.left.getContext('2d');
      if (ctx) ctx.clearRect(0, 0, this.left.width, this.left.height);
      this._clearOverlay(this.leftOverlay);
    }

    if (this.engineRight && this.currentSlide < this.engineRight.slideCount) {
      await this.engineRight.renderSlide(this.right, this.currentSlide, { width: w, dpr });
      this._paintOverlay(this.rightOverlay, this.engineRight, 'add');
    } else {
      const ctx = this.right.getContext('2d');
      if (ctx) ctx.clearRect(0, 0, this.right.width, this.right.height);
      this._clearOverlay(this.rightOverlay);
    }

    this.opts.onSlideChange?.(this.currentSlide, this.slideCount);
  }

  private _paintOverlay(overlay: HTMLCanvasElement, engine: PptxPresentation, side: 'add' | 'remove'): void {
    const presentation = engine.presentation;
    if (!presentation || !this.result) {
      this._clearOverlay(overlay);
      return;
    }
    const slideWidthEMU = presentation.slideWidth;
    const slideHeightEMU = presentation.slideHeight;

    const cssWidth = this.opts.width ?? overlay.parentElement?.clientWidth ?? 600;
    const cssHeight = Math.round(slideHeightEMU * (cssWidth / slideWidthEMU));
    const dpr = this.opts.dpr ?? window.devicePixelRatio ?? 1;
    overlay.width = Math.round(cssWidth * dpr);
    overlay.height = Math.round(cssHeight * dpr);
    overlay.style.width = `${cssWidth}px`;
    overlay.style.height = `${cssHeight}px`;
    const ctx = overlay.getContext('2d');
    if (!ctx) return;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, cssWidth, cssHeight);

    const colors = {
      add: this.opts.colors?.add ?? '#22c55e',
      remove: this.opts.colors?.remove ?? '#ef4444',
      modify: this.opts.colors?.modify ?? '#f59e0b',
    };

    const scale = cssWidth / slideWidthEMU;
    for (const c of this.result.changes) {
      if (c.location?.kind !== 'slide') continue;
      if (c.location.slideIndex !== this.currentSlide) continue;
      if (!c.location.bbox) continue;
      // Show "remove" boxes only on the left, "add" only on the right; modify on both.
      if (c.op === 'add' && side !== 'add') continue;
      if (c.op === 'remove' && side !== 'remove') continue;
      const colour = c.op === 'add' ? colors.add : c.op === 'remove' ? colors.remove : colors.modify;
      const { x, y, width, height } = c.location.bbox;
      ctx.save();
      ctx.lineWidth = 2;
      ctx.strokeStyle = colour;
      ctx.fillStyle = hexToRgba(colour, 0.18);
      ctx.fillRect(x * scale, y * scale, width * scale, height * scale);
      ctx.strokeRect(x * scale, y * scale, width * scale, height * scale);
      ctx.restore();
    }
  }

  private _clearOverlay(overlay: HTMLCanvasElement): void {
    const ctx = overlay.getContext('2d');
    if (!ctx) return;
    ctx.clearRect(0, 0, overlay.width, overlay.height);
  }
}

function hexToRgba(hex: string, alpha: number): string {
  const m = hex.match(/^#?([0-9a-fA-F]{6})$/);
  if (!m) return hex;
  const v = parseInt(m[1], 16);
  const r = (v >> 16) & 0xff;
  const g = (v >> 8) & 0xff;
  const b = v & 0xff;
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}
