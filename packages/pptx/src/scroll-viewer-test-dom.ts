import { vi } from 'vitest';
import type { PptxPresentation, RenderSlideOptions, RenderSlideToBitmapOptions } from './presentation';
import type { PptxTextRunInfo } from './renderer';

/** A recording fake DOM element. Extends the `FakeEl` pattern from
 *  text-layer.test.ts with the surface PptxScrollViewer touches: scrollTop,
 *  clientHeight/Width, event listeners, removeChild/remove, canvas ctx. */
export interface FakeEl {
  tag: string;
  textContent: string;
  innerHTML: string;
  style: Record<string, string> & { cssText: string };
  children: FakeEl[];
  parentElement: FakeEl | null;
  // canvas-only
  width: number;
  height: number;
  // scrollHost-only (settable by the test to drive the virtualization loop)
  scrollTop: number;
  clientHeight: number;
  clientWidth: number;
  _listeners: Map<string, Array<(e: unknown) => void>>;
  _bitmapCtx?: { transferFromImageBitmap: (b: unknown) => void; lastBitmap: unknown };
  appendChild(c: FakeEl): FakeEl;
  removeChild(c: FakeEl): FakeEl;
  remove(): void;
  insertBefore(n: FakeEl, ref: FakeEl | null): FakeEl;
  addEventListener(type: string, fn: (e: unknown) => void, opts?: unknown): void;
  removeEventListener(type: string, fn: (e: unknown) => void): void;
  getContext(kind: string): unknown;
  getBoundingClientRect(): { top: number; left: number; width: number; height: number };
  /** test-only: fire a recorded listener */
  dispatch(type: string, event?: unknown): void;
}

export function makeEl(tag: string): FakeEl {
  const style: Record<string, string> = {};
  const el: FakeEl = {
    tag,
    textContent: '',
    innerHTML: '',
    width: 0,
    height: 0,
    scrollTop: 0,
    clientHeight: 0,
    clientWidth: 0,
    children: [],
    parentElement: null,
    _listeners: new Map(),
    style: new Proxy(style as Record<string, string> & { cssText: string }, {
      set(target, prop: string, value: string) {
        if (prop === 'cssText') {
          for (const decl of value.split(';')) {
            const idx = decl.indexOf(':');
            if (idx > 0) target[decl.slice(0, idx).trim()] = decl.slice(idx + 1).trim();
          }
          target.cssText = value;
        } else {
          target[prop] = value;
        }
        return true;
      },
      get(target, prop: string) {
        return target[prop] ?? '';
      },
    }),
    appendChild(c: FakeEl) {
      c.parentElement = this;
      this.children.push(c);
      return c;
    },
    removeChild(c: FakeEl) {
      const i = this.children.indexOf(c);
      if (i >= 0) this.children.splice(i, 1);
      c.parentElement = null;
      return c;
    },
    remove() {
      this.parentElement?.removeChild(this);
    },
    insertBefore(n: FakeEl, ref: FakeEl | null) {
      n.parentElement = this;
      const i = ref ? this.children.indexOf(ref) : -1;
      if (i >= 0) this.children.splice(i, 0, n);
      else this.children.push(n);
      return n;
    },
    addEventListener(type: string, fn: (e: unknown) => void) {
      const arr = this._listeners.get(type) ?? [];
      arr.push(fn);
      this._listeners.set(type, arr);
    },
    removeEventListener(type: string, fn: (e: unknown) => void) {
      const arr = this._listeners.get(type);
      if (arr) this._listeners.set(type, arr.filter((f) => f !== fn));
    },
    getContext(kind: string) {
      if (kind === 'bitmaprenderer') {
        this._bitmapCtx = {
          lastBitmap: null,
          transferFromImageBitmap(b: unknown) {
            this.lastBitmap = b;
          },
        };
        return this._bitmapCtx;
      }
      // A minimal recording 2d context is never actually used by the viewer
      // (renderSlide owns the ctx); return a no-op object for safety.
      return {};
    },
    getBoundingClientRect() {
      return { top: 0, left: 0, width: this.clientWidth, height: this.clientHeight };
    },
    dispatch(type: string, event: unknown = {}) {
      for (const fn of this._listeners.get(type) ?? []) fn(event);
    },
  };
  // Mirror the DOM: assigning `innerHTML = ''` detaches all children (both
  // buildPptxTextLayer's leading clear and the viewer's overlay-teardown clear on
  // recycle rely on this). A non-empty assignment only records the string (the
  // viewer/text-layer never set non-empty innerHTML, so no parsing is needed).
  let _html = '';
  Object.defineProperty(el, 'innerHTML', {
    get() {
      return _html;
    },
    set(value: string) {
      _html = value;
      if (value === '') {
        for (const c of el.children) c.parentElement = null;
        el.children.length = 0;
      }
    },
    enumerable: true,
    configurable: true,
  });
  return el;
}

/** A container FakeEl with a nonzero clientWidth so `width` defaults resolve. */
export function makeContainer(clientWidth = 800, clientHeight = 600): FakeEl {
  const c = makeEl('div');
  c.clientWidth = clientWidth;
  c.clientHeight = clientHeight;
  return c;
}

/** Install a recording document + window + ResizeObserver into globals.
 *  Returns the last-constructed ResizeObserver callback so a test can fire a
 *  synthetic resize. Call `vi.unstubAllGlobals()` in afterEach. */
export function installDom(): { resizeCb: () => (() => void) | undefined } {
  let lastResizeCb: (() => void) | undefined;
  vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
  vi.stubGlobal('window', { devicePixelRatio: 1 });
  vi.stubGlobal(
    'ResizeObserver',
    class {
      constructor(cb: () => void) {
        lastResizeCb = cb;
      }
      observe(): void {}
      disconnect(): void {}
    },
  );
  return { resizeCb: () => lastResizeCb };
}

export interface RenderCall {
  slide: number;
  /** The per-call `width` (px) the viewer passed — asserted by T3 to confirm
   *  each slide gets its OWN px width (uniform px width from the deck-wide scale,
   *  §7). pptx slides are uniform, so every value equals the same px width, but
   *  the per-call contract still passes a width per slide. */
  width?: number;
  resolve: () => void;
  reject: (e: Error) => void;
}

/** A fake PptxPresentation covering exactly the surface PptxScrollViewer consumes.
 *  Slide size is UNIFORM (slideWidth/slideHeight in EMU), unlike docx's per-index
 *  pageSize. Deferred mode: renderSlide / renderSlideToBitmap return promises the
 *  test resolves via the recorded RenderCall, so coalescing / stale-drop is
 *  observable. */
export class FakePptxEngine {
  destroyed = false;
  renderCalls: RenderCall[] = [];
  bitmapCalls: RenderCall[] = [];
  createdBitmaps: Array<{ width: number; height: number; close: ReturnType<typeof vi.fn> }> = [];
  /** Runs fed to `renderSlide`'s `onTextRun` (main mode only) so a test can drive
   *  the viewer's per-slot text overlay (buildPptxTextLayer) without a real
   *  render. Empty/undefined ⇒ no runs emitted. */
  feedTextRuns?: PptxTextRunInfo[];
  constructor(
    private _slideCount: number,
    public readonly slideWidth: number, // EMU, deck-wide (uniform)
    public readonly slideHeight: number, // EMU, deck-wide (uniform)
    private _mode: 'main' | 'worker' = 'main',
    private deferred = false,
  ) {}
  get slideCount(): number {
    return this._slideCount;
  }
  /** Mirrors the real `PptxPresentation.mode` fact (presentation.ts) — the exact
   *  fact the viewer constructor reads to decide the render path (main ⇒
   *  renderSlide, worker ⇒ renderSlideToBitmap). Design §11: no probing / no
   *  silent mis-pathing. */
  get mode(): 'main' | 'worker' {
    return this._mode;
  }
  renderSlide(_canvas: unknown, slide: number, opts?: RenderSlideOptions): Promise<void> {
    // Mirror the real renderer's backing-store sizing so tests can exercise the
    // dpr≠1 path: the real PptxPresentation.renderSlide sets `canvas.width =
    // round(cssWidth × dpr)` (and height to keep the slide aspect). A retina (dpr 2)
    // render leaves canvas.width at 2× the CSS box — the overlay must NOT copy it.
    if (this._mode === 'main' && opts?.width && opts.width > 0) {
      const dpr = opts.dpr ?? 1;
      const canvas = _canvas as { width: number; height: number };
      canvas.width = Math.round(opts.width * dpr);
      canvas.height = this.slideWidth > 0 ? Math.round((canvas.width * this.slideHeight) / this.slideWidth) : 0;
    }
    if (this._mode === 'worker') {
      // Record the attempt BEFORE throwing so a test can assert the viewer never
      // routes a worker-mode engine through renderSlide (and detect wrong-path
      // routing). The real PptxPresentation.renderSlide THROWS synchronously in
      // worker mode — mirror that exactly (do NOT return Promise.reject), reusing
      // the real error text (presentation.ts).
      this.renderCalls.push({ slide, resolve: () => {}, reject: () => {} });
      throw new Error(
        "renderSlide(canvas) is unavailable in mode: 'worker'; use renderSlideToBitmap() and paint it via an ImageBitmapRenderingContext",
      );
    }
    // Emit any fed runs to the viewer's internal onTextRun so it can build the
    // per-slot overlay (mirrors the real renderer emitting run geometry).
    for (const r of this.feedTextRuns ?? []) opts?.onTextRun?.(r);
    return new Promise<void>((resolve, reject) => {
      const call: RenderCall = { slide, width: opts?.width, resolve: () => resolve(), reject };
      this.renderCalls.push(call);
      if (!this.deferred) resolve();
    });
  }
  renderSlideToBitmap(slide: number, opts?: RenderSlideToBitmapOptions): Promise<ImageBitmap> {
    const w = opts?.width ?? 960;
    const h = this.slideWidth > 0 ? Math.round((w * this.slideHeight) / this.slideWidth) : 0;
    const bmp = { width: w, height: h, close: vi.fn() };
    this.createdBitmaps.push(bmp);
    return new Promise<ImageBitmap>((resolve, reject) => {
      const call: RenderCall = {
        slide,
        width: opts?.width,
        resolve: () => resolve(bmp as unknown as ImageBitmap),
        reject,
      };
      this.bitmapCalls.push(call);
      if (!this.deferred) resolve(bmp as unknown as ImageBitmap);
    });
  }
  /** The per-call `width` (px) recorded for every renderSlide call, in call order.
   *  T3 asserts each mounted slide received its OWN px width. */
  renderSlideWidths(): number[] {
    return this.renderCalls.map((c) => c.width ?? NaN);
  }
  destroy(): void {
    this.destroyed = true;
  }
  asPres(): PptxPresentation {
    return this as unknown as PptxPresentation;
  }
}
