import { vi } from 'vitest';
import type { DocxDocument } from './document';
import type { RenderPageOptions } from './types';
import type { WireRenderPageOptions } from './worker-protocol';

/** A recording fake DOM element. Extends the `FakeEl` pattern from
 *  text-layer.test.ts with the surface DocxScrollViewer touches: scrollTop,
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
      // (renderPage owns the ctx); return a no-op object for safety.
      return {};
    },
    getBoundingClientRect() {
      return { top: 0, left: 0, width: this.clientWidth, height: this.clientHeight };
    },
    dispatch(type: string, event: unknown = {}) {
      for (const fn of this._listeners.get(type) ?? []) fn(event);
    },
  };
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
  page: number;
  /** The per-call `width` (px) the viewer passed — asserted by T3 to confirm
   *  each page gets its OWN px width (uniform px-per-pt scale, §7). */
  width?: number;
  resolve: () => void;
  reject: (e: Error) => void;
}

/** A fake DocxDocument covering exactly the surface DocxScrollViewer consumes.
 *  Deferred mode: renderPage / renderPageToBitmap return promises the test
 *  resolves via the recorded RenderCall, so coalescing / stale-drop is
 *  observable. */
export class FakeDocxEngine {
  destroyed = false;
  renderCalls: RenderCall[] = [];
  bitmapCalls: RenderCall[] = [];
  createdBitmaps: Array<{ width: number; height: number; close: ReturnType<typeof vi.fn> }> = [];
  constructor(
    private _pageCount: number,
    // Uniform-page convention: a single-element `_sizes` array means EVERY page
    // is that size (pageSize clamps the index). T2 variable-height tests pass a
    // full per-page array instead.
    private _sizes: Array<{ widthPt: number; heightPt: number }>,
    private _mode: 'main' | 'worker' = 'main',
    private deferred = false,
  ) {}
  get pageCount(): number {
    return this._pageCount;
  }
  /** Mirrors the real `DocxDocument.mode` fact (document.ts) — the exact fact the
   *  viewer constructor reads to decide the render path (main ⇒ renderPage,
   *  worker ⇒ renderPageToBitmap). Design §11: no probing / no silent mis-pathing. */
  get mode(): 'main' | 'worker' {
    return this._mode;
  }
  pageSize(i: number): { widthPt: number; heightPt: number } {
    const clamped = Math.max(0, Math.min(i, this._sizes.length - 1));
    const s = this._sizes[clamped] ?? { widthPt: 0, heightPt: 0 };
    return { widthPt: s.widthPt, heightPt: s.heightPt };
  }
  renderPage(_canvas: unknown, page: number, opts?: RenderPageOptions): Promise<void> {
    if (this._mode === 'worker') {
      // Record the attempt BEFORE throwing so a test can assert the viewer never
      // routes a worker-mode engine through renderPage (and detect wrong-path
      // routing). The real DocxDocument.renderPage is a non-async method that
      // THROWS synchronously in worker mode — mirror that exactly (do NOT return
      // Promise.reject), reusing the real error text (document.ts).
      this.renderCalls.push({ page, resolve: () => {}, reject: () => {} });
      throw new Error(
        "renderPage(canvas) is unavailable in mode: 'worker'; use renderPageToBitmap() and paint it via an ImageBitmapRenderingContext",
      );
    }
    return new Promise<void>((resolve, reject) => {
      const call: RenderCall = { page, width: opts?.width, resolve: () => resolve(), reject };
      this.renderCalls.push(call);
      if (!this.deferred) resolve();
    });
  }
  renderPageToBitmap(page: number, opts?: WireRenderPageOptions): Promise<ImageBitmap> {
    const size = this.pageSize(page);
    const bmp = { width: Math.round(size.widthPt), height: Math.round(size.heightPt), close: vi.fn() };
    this.createdBitmaps.push(bmp);
    return new Promise<ImageBitmap>((resolve, reject) => {
      const call: RenderCall = {
        page,
        width: opts?.width,
        resolve: () => resolve(bmp as unknown as ImageBitmap),
        reject,
      };
      this.bitmapCalls.push(call);
      if (!this.deferred) resolve(bmp as unknown as ImageBitmap);
    });
  }
  /** The per-call `width` (px) recorded for every renderPage call, in call order.
   *  T3 asserts each mounted page received its OWN px width. */
  renderPageWidths(): number[] {
    return this.renderCalls.map((c) => c.width ?? NaN);
  }
  destroy(): void {
    this.destroyed = true;
  }
  asDoc(): DocxDocument {
    return this as unknown as DocxDocument;
  }
}
