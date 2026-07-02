import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxScrollViewer } from './scroll-viewer.js';
import { DocxDocument } from './document.js';
import { installDom, makeContainer, FakeDocxEngine, type FakeEl } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe('DocxScrollViewer — skeleton (T1)', () => {
  it('builds the wrapper → scrollHost → spacer DOM inside the container', () => {
    installDom();
    const container = makeContainer();
    const engine = new FakeDocxEngine(3, [{ widthPt: 612, heightPt: 792 }]);
    const v = new DocxScrollViewer(container as unknown as HTMLElement, { document: engine.asDoc() });
    // container → wrapper
    const wrapper = container.children[0];
    expect(wrapper.tag).toBe('div');
    expect(wrapper.style.position).toBe('relative');
    // wrapper → scrollHost
    const scrollHost = wrapper.children[0];
    expect(scrollHost.style.overflow).toBe('auto');
    // scrollHost → spacer
    const spacer = scrollHost.children[0];
    expect(spacer.style.position).toBe('absolute');
    v.destroy();
  });

  it('exposes pageCount from the injected engine', () => {
    installDom();
    const engine = new FakeDocxEngine(5, [{ widthPt: 612, heightPt: 792 }]);
    const v = new DocxScrollViewer(makeContainer() as unknown as HTMLElement, { document: engine.asDoc() });
    expect(v.pageCount).toBe(5);
    v.destroy();
  });

  it('load() is unsupported when an engine is injected', async () => {
    installDom();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }]);
    const v = new DocxScrollViewer(makeContainer() as unknown as HTMLElement, { document: engine.asDoc() });
    await expect(v.load('x.docx')).rejects.toThrow(/injected/i);
    v.destroy();
  });

  it('destroy() removes the DOM and does NOT destroy an injected engine', () => {
    installDom();
    const container = makeContainer();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }]);
    const v = new DocxScrollViewer(container as unknown as HTMLElement, { document: engine.asDoc() });
    expect(container.children.length).toBe(1); // wrapper mounted
    v.destroy();
    expect(container.children.length).toBe(0); // wrapper removed
    expect(engine.destroyed).toBe(false); // injected engine preserved (caller owns it)
  });

  it('pageCount is 0 before load resolves (no injected engine)', () => {
    installDom();
    const v = new DocxScrollViewer(makeContainer() as unknown as HTMLElement, {});
    expect(v.pageCount).toBe(0);
    v.destroy();
  });

  // O1 (design §11): an injected engine's own `mode` is authoritative. An
  // EXPLICITLY conflicting opts.mode is a mis-configuration rejected at
  // construction; a matching or absent opts.mode constructs fine.
  it('throws when opts.mode conflicts with an injected worker-mode engine', () => {
    installDom();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }], 'worker');
    expect(
      () =>
        new DocxScrollViewer(makeContainer() as unknown as HTMLElement, {
          document: engine.asDoc(),
          mode: 'main',
        }),
    ).toThrow(/mode/i);
  });

  it('does NOT throw when opts.mode matches an injected worker-mode engine', () => {
    installDom();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }], 'worker');
    const v = new DocxScrollViewer(makeContainer() as unknown as HTMLElement, {
      document: engine.asDoc(),
      mode: 'worker',
    });
    expect(v.pageCount).toBe(1);
    v.destroy();
    // Injected engine is caller-owned even in the worker case: destroy() leaves it intact.
    expect(engine.destroyed).toBe(false);
  });

  it('constructs a default-main injected engine with absent opts.mode (load still rejects; destroy preserves engine)', async () => {
    installDom();
    // Default mode is 'main'; opts.mode is absent ⇒ no conflict, resolved path is main.
    const engine = new FakeDocxEngine(2, [{ widthPt: 612, heightPt: 792 }]);
    const v = new DocxScrollViewer(makeContainer() as unknown as HTMLElement, {
      document: engine.asDoc(),
    });
    expect(v.pageCount).toBe(2);
    await expect(v.load('x.docx')).rejects.toThrow(/injected/i);
    v.destroy();
    // Injected engine is caller-owned: destroy() must not tear it down.
    expect(engine.destroyed).toBe(false);
  });
});

describe('DocxScrollViewer — layout + virtualization (T2)', () => {
  // Uniform 100pt×200pt pages, dpr:1, PT_TO_PX applied in the viewer. To keep
  // the assertions in pt-independent terms we drive scrollTop/clientHeight in the
  // SAME px units the viewer computes heights in. Use pageSize heightPt directly
  // and set base scale by mapping the FIRST page's width to the container width
  // (width:undefined ⇒ fit the container).
  function setup(pageCount: number, opts = {}) {
    const dom = installDom();
    const container = makeContainer(200, 400); // clientWidth 200, clientHeight 400
    const engine = new FakeDocxEngine(
      pageCount,
      // Full per-page array (fixed harness treats a single-element array as
      // uniform, but T2 passes the full array explicitly).
      Array.from({ length: pageCount }, () => ({ widthPt: 100, heightPt: 200 })),
    );
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
      overscan: 1,
      width: undefined, // fit container width (200) → base scale below
      ...opts,
    });
    const wrapper = container.children[0] as FakeEl;
    const scrollHost = wrapper.children[0] as FakeEl;
    const spacer = scrollHost.children[0] as FakeEl;
    // Drive layout: container width 200, page width 100pt → base scale maps
    // 100pt*PT_TO_PX to 200px. Provide the viewport height.
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    return { dom, container, engine, v, wrapper, scrollHost, spacer };
  }

  it('sizes the spacer to computeVisibleRange.totalHeight and mounts only the visible window', () => {
    const { v, scrollHost, spacer } = setup(10);
    v.relayout(); // T2 exposes an explicit relayout() the viewer calls after load/resize
    // Spacer height > 0 and equals Σ page-heights + (n-1)*gap in px.
    expect(parseFloat(spacer.style.height)).toBeGreaterThan(0);
    // At scrollTop 0 with a 400px viewport and pages ~ (100pt fit to 200px wide →
    // 200pt page becomes 400px tall) only page 0 is fully visible; overscan 1
    // mounts page 1 too. Assert the mounted slot count is small and bounded.
    const mounted = scrollHost.children.filter((c) => c !== spacer);
    expect(mounted.length).toBeGreaterThanOrEqual(1);
    expect(mounted.length).toBeLessThanOrEqual(3); // topIndex 0 .. lastVisible+overscan
  });

  it('spacer height is exact for variable page heights + gap', () => {
    const dom = installDom();
    const container = makeContainer(200, 400);
    // Variable heights: widths equal so base scale is one value; heights differ.
    const sizes = [
      { widthPt: 100, heightPt: 200 },
      { widthPt: 100, heightPt: 300 },
      { widthPt: 100, heightPt: 150 },
    ];
    const engine = new FakeDocxEngine(3, sizes);
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
      overscan: 1,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    const spacer = scrollHost.children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    // base scale = 200 / (100 * PT_TO_PX) = 200 / (400/3) = 1.5.
    const PT_TO_PX = 4 / 3;
    const scale = 200 / (100 * PT_TO_PX);
    const heights = sizes.map((s) => s.heightPt * PT_TO_PX * scale);
    const expected = heights.reduce((a, b) => a + b, 0) + (sizes.length - 1) * 10;
    expect(parseFloat(spacer.style.height)).toBeCloseTo(expected, 3);
    v.destroy();
    void dom;
  });

  it('recycles slots on scroll without unbounded canvas growth (pool reuse)', () => {
    const { v, scrollHost } = setup(50);
    v.relayout();
    const initialMount = scrollHost.children.length;
    // Scroll far down and fire the scroll listener repeatedly.
    for (let top = 0; top <= 4000; top += 400) {
      scrollHost.scrollTop = top;
      scrollHost.dispatch('scroll');
    }
    // The DOM child count (spacer + mounted slots) must stay bounded — the pool
    // reuses slots rather than appending a new canvas per page.
    expect(scrollHost.children.length).toBeLessThanOrEqual(initialMount + 2);
  });

  it('scrolling far then back reuses pooled slot wrappers (bounded distinct allocations)', () => {
    const { v, scrollHost } = setup(50);
    v.relayout();
    // Track EVERY distinct wrapper element the viewer ever appends. If slots are
    // pooled (recycled), scrolling across the whole document then back reuses a
    // small fixed set of wrappers rather than allocating one per visited page.
    const seen = new Set<FakeEl>();
    const collect = () => {
      for (const c of scrollHost.children) if (c.tag === 'div') seen.add(c);
    };
    collect();
    // Sweep deep into the document and back to the top.
    for (const top of [8000, 16000, 8000, 0]) {
      scrollHost.scrollTop = top;
      scrollHost.dispatch('scroll');
      collect();
    }
    // 50 pages were visited across the sweep; a per-page allocator would have
    // created dozens of wrappers. The pool must keep the distinct-wrapper count
    // tiny: spacer(1) + at most a couple of window-sized generations of slots.
    // Window here is ~2-4 slots; allow generous headroom but far below 50.
    expect(seen.size).toBeLessThanOrEqual(8);
    // Coming back to the top mounts page 0 again, drawn from the pool.
    expect(v.mountedPageIndicesForTest()).toContain(0);
  });

  it('mounts the correct window for a mid-document scrollTop', () => {
    const { v, scrollHost } = setup(20);
    v.relayout();
    // Jump to a scrollTop deep in the doc; the mounted slots must include the
    // page under the viewport top (topVisiblePage) and its overscan neighbours.
    scrollHost.scrollTop = 2000;
    scrollHost.dispatch('scroll');
    const top = v.topVisiblePage;
    const mountedPages = v.mountedPageIndicesForTest(); // T2 test hook
    expect(mountedPages).toContain(top);
    expect(Math.max(...mountedPages) - Math.min(...mountedPages)).toBeLessThanOrEqual(4);
  });
});

describe('DocxScrollViewer — rendering (T3)', () => {
  it("main mode: calls renderPage once per mounted slot with that page's px width", () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(10, [{ widthPt: 100, heightPt: 200 }]);
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    // Every mounted page got exactly one renderPage call, no more no less.
    const mounted = v.mountedPageIndicesForTest().sort((a, b) => a - b);
    const rendered = engine.renderCalls.map((c) => c.page).sort((a, b) => a - b);
    expect(rendered).toEqual(mounted);
    // Each renderPage call carried THIS page's own px width (uniform px-per-pt
    // scale, §7). base scale = 200/(100*PT_TO_PX); page width px = 200 for all.
    const widths = engine.renderPageWidths();
    for (const w of widths) expect(w).toBeCloseTo(200, 3);
    v.destroy();
  });

  it('main mode: passes each page its OWN px width for MIXED page sizes (uniform px-per-pt scale, §7)', () => {
    installDom();
    const container = makeContainer(200, 400);
    // Mixed physical widths: page 0 is 100pt wide, page 1 is 200pt wide. The
    // uniform document scale is fixed by fitting the FIRST page to the container
    // (base scale = 200 / (100 * PT_TO_PX) = 1.5), so page 1 — twice as wide —
    // must render at TWICE the px width, not the same fit-width. A vacuous impl
    // that hands every page a constant fit-width would give [200, 200].
    const engine = new FakeDocxEngine(2, [
      { widthPt: 100, heightPt: 100 },
      { widthPt: 200, heightPt: 100 },
    ]);
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    // Both pages mount (page 0 height px = 100*PT_TO_PX*1.5 = 200; page 1 sits at
    // top 210 and intersects the 400px viewport), so both get exactly one render.
    const mounted = v.mountedPageIndicesForTest().sort((a, b) => a - b);
    expect(mounted).toEqual([0, 1]);
    // Per-page px widths in call order: page 0 → 200, page 1 → 400.
    const widths = engine.renderCalls
      .slice()
      .sort((a, b) => a.page - b.page)
      .map((c) => c.width ?? NaN);
    expect(widths[0]).toBeCloseTo(200, 3);
    expect(widths[1]).toBeCloseTo(400, 3);
    v.destroy();
  });

  it('does not re-render a mounted slot for the same page on a no-op scroll', () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(10, [{ widthPt: 100, heightPt: 200 }]);
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    const before = engine.renderCalls.length;
    scrollHost.scrollTop = 0; // unchanged window
    scrollHost.dispatch('scroll');
    expect(engine.renderCalls.length).toBe(before); // no duplicate renders
    v.destroy();
  });

  it('worker mode: never calls renderPage — routes every slot through renderPageToBitmap', async () => {
    installDom();
    const container = makeContainer(200, 400);
    // mode:'worker' — the real DocxDocument.renderPage THROWS synchronously in
    // worker mode; a viewer that mis-routed would blow up (and renderCalls would
    // record the attempt). The direct _mode routing must never touch renderPage.
    const engine = new FakeDocxEngine(10, [{ widthPt: 100, heightPt: 200 }], 'worker');
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    await Promise.resolve();
    await Promise.resolve();
    expect(engine.renderCalls).toHaveLength(0); // renderPage never touched
    // Every mounted page dispatched exactly one bitmap render.
    const mounted = v.mountedPageIndicesForTest().sort((a, b) => a - b);
    const bitmapPages = [...new Set(engine.bitmapCalls.map((c) => c.page))].sort((a, b) => a - b);
    expect(bitmapPages).toEqual(mounted);
    v.destroy();
  });

  it('worker mode: paints a resolved bitmap into the slot canvas (transfer)', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(10, [{ widthPt: 100, heightPt: 200 }], 'worker');
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    await Promise.resolve();
    await Promise.resolve();
    // A bitmap was produced and transferred into the mounted slot's canvas.
    expect(engine.createdBitmaps.length).toBeGreaterThan(0);
    const slotWrapper = scrollHost.children.find((c) =>
      c.children.some((k) => k.tag === 'canvas'),
    ) as FakeEl;
    const canvas = slotWrapper.children.find((k) => k.tag === 'canvas') as FakeEl;
    // The bitmaprenderer ctx received the bitmap (fake records lastBitmap).
    expect(canvas._bitmapCtx?.lastBitmap).toBe(engine.createdBitmaps[0]);
    v.destroy();
  });

  it('worker mode: closes the ImageBitmap on recycle (deferred — render in flight when the slot recycles)', async () => {
    // Bitmap-close observability path (plan open question): the primary path
    // (deterministic slot.bitmap hold under synchronous resolve) does NOT hold —
    // in non-deferred mode transferFromImageBitmap consumes the bitmap and nulls
    // slot.bitmap synchronously BEFORE any scroll-away, so a later recycle has
    // nothing to close. We use the DOCUMENTED FALLBACK: a deferred fake so the
    // render is genuinely in flight when the slot recycles, and the on-resolution
    // stale-check closes the orphan (design §11).
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(50, [{ widthPt: 100, heightPt: 200 }], 'worker', true);
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    const firstBatch = engine.bitmapCalls.length;
    expect(firstBatch).toBeGreaterThan(0);
    // Scroll far away so the first pages' slots recycle while their renders are
    // still in flight (deferred → not yet resolved).
    scrollHost.scrollTop = 12000;
    scrollHost.dispatch('scroll');
    // Now resolve the early (now-stale) renders — the viewer must close them.
    for (const call of engine.bitmapCalls.slice(0, firstBatch)) call.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(engine.createdBitmaps.some((b) => b.close.mock.calls.length > 0)).toBe(true);
    v.destroy();
  });

  it('worker mode: drops a stale in-flight render — no paint + bitmap closed when the slot moved on', async () => {
    installDom();
    const container = makeContainer(200, 400);
    // Deferred: renderPageToBitmap resolves only when the test calls resolve(),
    // so we can scroll a slot's page out of the window BEFORE the bitmap arrives.
    const engine = new FakeDocxEngine(50, [{ widthPt: 100, heightPt: 200 }], 'worker', true);
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    // Page 0's bitmap is dispatched but NOT yet resolved.
    const page0 = engine.bitmapCalls.find((c) => c.page === 0);
    expect(page0).toBeDefined();
    // Scroll far away so page 0's slot recycles while its render is in flight.
    scrollHost.scrollTop = 12000;
    scrollHost.dispatch('scroll');
    expect(v.mountedPageIndicesForTest()).not.toContain(0);
    // Now the stale render for page 0 resolves — the viewer must NOT paint it and
    // must close the orphaned bitmap.
    const bmp0 = engine.createdBitmaps[engine.bitmapCalls.indexOf(page0!)];
    page0!.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(bmp0.close.mock.calls.length).toBeGreaterThan(0); // orphan closed
    v.destroy();
  });

  it('worker mode: a page that recycles then re-mounts while in flight still gets a fresh render', async () => {
    // Coalescing keys on page index; a naive Set<number> would swallow the
    // re-mounted slot's render (the stale resolve clears in-flight AFTER the
    // remount coalesced away → the new slot never paints). The viewer must
    // re-dispatch the live slot's render so the page never stays blank.
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(50, [{ widthPt: 100, heightPt: 200 }], 'worker', true);
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    const page0Initial = engine.bitmapCalls.find((c) => c.page === 0);
    expect(page0Initial).toBeDefined();
    const dispatchesForPage0 = () => engine.bitmapCalls.filter((c) => c.page === 0).length;
    expect(dispatchesForPage0()).toBe(1);
    // Scroll away (page 0 recycles, render still in flight) then back to the top
    // (page 0 re-mounts on a fresh slot) — all while the initial render is deferred.
    scrollHost.scrollTop = 12000;
    scrollHost.dispatch('scroll');
    expect(v.mountedPageIndicesForTest()).not.toContain(0);
    scrollHost.scrollTop = 0;
    scrollHost.dispatch('scroll');
    expect(v.mountedPageIndicesForTest()).toContain(0);
    // Coalescing pin: page 0's render is STILL in flight (initial deferred call
    // not yet resolved), so the re-mount must NOT dispatch a second render for
    // page 0 — the in-flight guard swallows it. Exactly one dispatch so far.
    expect(dispatchesForPage0()).toBe(1);
    // The initial (now stale) render resolves — the orphan is dropped and the
    // re-mounted slot's render is (re-)dispatched.
    page0Initial!.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(dispatchesForPage0()).toBeGreaterThanOrEqual(2); // fresh render issued
    // Resolve the fresh render and confirm it paints into the live slot.
    for (const c of engine.bitmapCalls.filter((c) => c.page === 0)) c.resolve();
    await Promise.resolve();
    await Promise.resolve();
    const slot0 = scrollHost.children.find(
      (c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas') && c.style.top === '0px',
    ) as FakeEl | undefined;
    expect(slot0).toBeDefined();
    const canvas0 = slot0!.children.find((k) => k.tag === 'canvas') as FakeEl;
    expect(canvas0._bitmapCtx?.lastBitmap).toBeTruthy(); // painted, not blank
    v.destroy();
  });

  it('worker mode: a PLAIN render rejection (slot still live, epoch unchanged) does NOT re-dispatch — no retry storm', async () => {
    // B1 regression: the finally re-dispatch must gate on STALENESS, not merely
    // `!painted`. When `renderPageToBitmap` rejects while the slot is still live
    // and the epoch is unchanged, `painted` is false and `live === slot`, but
    // NEITHER staleness test fires, so we must NOT re-dispatch. A `!painted`-only
    // gate would loop reject → re-dispatch → reject … unbounded (empirically 1→2
    // →3→4 with onError every round). The onError contract leaves the page blank.
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(50, [{ widthPt: 100, heightPt: 200 }], 'worker', true);
    const onError = vi.fn();
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
      overscan: 1,
      onError,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();

    const page0 = engine.bitmapCalls.find((c) => c.page === 0);
    expect(page0).toBeDefined();
    expect(v.mountedPageIndicesForTest()).toContain(0); // slot still live
    const dispatchesForPage0 = () => engine.bitmapCalls.filter((c) => c.page === 0).length;
    expect(dispatchesForPage0()).toBe(1);

    // Reject the render with the slot STILL LIVE and no scale change (epoch fixed).
    page0!.reject(new Error('worker render failed'));
    // Flush microtasks generously — a retry storm would keep queuing dispatches.
    for (let k = 0; k < 8; k++) await Promise.resolve();

    // Dispatch count for page 0 stayed at 1 — the storm is gone.
    expect(dispatchesForPage0()).toBe(1);
    // onError fired exactly once (the single failure), never again.
    expect(onError).toHaveBeenCalledTimes(1);
    // Still no further dispatches after more microtask flushing.
    for (let k = 0; k < 8; k++) await Promise.resolve();
    expect(dispatchesForPage0()).toBe(1);
    expect(onError).toHaveBeenCalledTimes(1);
    v.destroy();
  });

  it('worker mode: destroy() mid-flight closes the resolving bitmap and does not fire onError post-destroy', async () => {
    installDom();
    const container = makeContainer(200, 400);
    // Deferred so a render is genuinely in flight when we destroy the viewer.
    const engine = new FakeDocxEngine(50, [{ widthPt: 100, heightPt: 200 }], 'worker', true);
    const onError = vi.fn();
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
      onError,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    const page0 = engine.bitmapCalls.find((c) => c.page === 0);
    expect(page0).toBeDefined();
    // Tear down while page 0's render is still in flight. destroy() recycles the
    // slot (no bitmap held yet — none received), so the on-resolution stale-check
    // must close the orphan; the identity guard fails (slot no longer live).
    v.destroy();
    const bmp0 = engine.createdBitmaps[engine.bitmapCalls.indexOf(page0!)];
    page0!.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(bmp0.close.mock.calls.length).toBeGreaterThan(0); // orphan closed, no leak
    expect(onError).not.toHaveBeenCalled(); // no error surfaced after teardown
  });
});

describe('DocxScrollViewer — zoom (T4)', () => {
  // PT_TO_PX = 4/3. Uniform 100pt×200pt pages, container 200×400, width:undefined
  // ⇒ base fit maps page-0 width (100pt) to 200px:
  //   base = 200 / (100 * 4/3) = 1.5
  //   page height px = 200 * 4/3 * base = 400
  //   offset[i] = i * (400 + gap) = i * 410
  const PT_TO_PX = 4 / 3;

  function setup(pageCount = 20, opts = {}) {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(
      pageCount,
      Array.from({ length: pageCount }, () => ({ widthPt: 100, heightPt: 200 })),
    );
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
      overscan: 1,
      zoomMin: 0.5,
      zoomMax: 3,
      ...opts,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    return { v, scrollHost, engine, container };
  }

  it('base fit scale maps the first page width to the container width', () => {
    const { v } = setup();
    // base = 200 / (100 * 4/3) = 1.5
    expect(v.scaleForTest()).toBeCloseTo(1.5, 5);
    expect(v.baseScaleForTest()).toBeCloseTo(1.5, 5);
    v.destroy();
  });

  it('re-anchors so the page under the viewport top stays fixed across a zoom (intraFrac 0)', () => {
    const { v, scrollHost } = setup();
    // page 3 top offset at base (1.5): 3 * (400 + 10) = 1230
    scrollHost.scrollTop = 1230;
    scrollHost.dispatch('scroll');
    expect(v.topVisiblePage).toBe(3);
    const cur = v.scaleForTest(); // 1.5
    v.setScale(cur * 2); // 3.0 (== zoomMax, no clamp)
    expect(v.scaleForTest()).toBeCloseTo(3, 5);
    // page 3 stays the top page; new offset = 3 * (800 + 10) = 2430
    expect(v.topVisiblePage).toBe(3);
    expect(Math.abs(scrollHost.scrollTop - 2430)).toBeLessThan(2);
    v.destroy();
  });

  it('re-anchors preserving the intra-page fraction (intraFrac ≠ 0)', () => {
    const { v, scrollHost } = setup();
    // Scroll so HALF of page 3 (200 of its 400px) has passed above the viewport
    // top: scrollTop = offset[3] + 0.5*400 = 1230 + 200 = 1430 → intraFrac 0.5.
    scrollHost.scrollTop = 1430;
    scrollHost.dispatch('scroll');
    expect(v.topVisiblePage).toBe(3);
    v.setScale(v.scaleForTest() * 2); // 1.5 → 3.0; heights' = 800
    // newScrollTop = offset'[3] + 0.5 * 800 = 2430 + 400 = 2830
    expect(Math.abs(scrollHost.scrollTop - 2830)).toBeLessThan(2);
    // Page 3 is still under the viewport top: offset'[3]=2430 <= 2830 < 3240.
    expect(v.topVisiblePage).toBe(3);
    v.destroy();
  });

  it('clamps setScale to the absolute [zoomMin, zoomMax] px-per-pt bounds', () => {
    const { v } = setup();
    v.setScale(100); // above zoomMax 3 (absolute, NOT a multiple of base)
    expect(v.scaleForTest()).toBeCloseTo(3, 5);
    v.setScale(0.01); // below zoomMin 0.5
    expect(v.scaleForTest()).toBeCloseTo(0.5, 5);
    v.destroy();
  });

  it('setScale defaults clamp to [0.1, 4] when zoomMin/zoomMax are unset', () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(5, [{ widthPt: 100, heightPt: 200 }]);
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    v.setScale(999);
    expect(v.scaleForTest()).toBeCloseTo(4, 5); // default zoomMax
    v.setScale(0.0001);
    expect(v.scaleForTest()).toBeCloseTo(0.1, 5); // default zoomMin
    v.destroy();
  });

  it('Ctrl+wheel zooms in and calls preventDefault; bare wheel does not zoom or preventDefault', () => {
    const { v, scrollHost } = setup();
    const before = v.scaleForTest();

    // Bare wheel: no zoom, native scroll (preventDefault NOT called).
    const barePrevent = vi.fn();
    scrollHost.dispatch('wheel', { deltaY: -100, ctrlKey: false, metaKey: false, preventDefault: barePrevent });
    expect(v.scaleForTest()).toBe(before);
    expect(barePrevent).not.toHaveBeenCalled();

    // Ctrl+wheel (deltaY<0 = zoom in): scale increases, preventDefault called.
    const ctrlPrevent = vi.fn();
    scrollHost.dispatch('wheel', { deltaY: -100, ctrlKey: true, metaKey: false, preventDefault: ctrlPrevent });
    expect(v.scaleForTest()).toBeGreaterThan(before);
    expect(ctrlPrevent).toHaveBeenCalledTimes(1);
    v.destroy();
  });

  it('Cmd(meta)+wheel also zooms', () => {
    const { v, scrollHost } = setup();
    const before = v.scaleForTest();
    scrollHost.dispatch('wheel', { deltaY: -100, ctrlKey: false, metaKey: true, preventDefault() {} });
    expect(v.scaleForTest()).toBeGreaterThan(before);
    v.destroy();
  });

  it('positive deltaY (ctrl+wheel down) zooms out', () => {
    const { v, scrollHost } = setup();
    const before = v.scaleForTest();
    scrollHost.dispatch('wheel', { deltaY: 100, ctrlKey: true, metaKey: false, preventDefault() {} });
    expect(v.scaleForTest()).toBeLessThan(before);
    v.destroy();
  });

  it('enableZoom:false installs no wheel handler — ctrl+wheel is inert', () => {
    const { v, scrollHost } = setup(20, { enableZoom: false });
    const before = v.scaleForTest();
    // No wheel listener was registered at all.
    expect(scrollHost._listeners.has('wheel')).toBe(false);
    const prevent = vi.fn();
    scrollHost.dispatch('wheel', { deltaY: -100, ctrlKey: true, preventDefault: prevent });
    expect(v.scaleForTest()).toBe(before);
    expect(prevent).not.toHaveBeenCalled();
    v.destroy();
  });

  it('a wheel with deltaY 0 is a no-op even with ctrl held', () => {
    const { v, scrollHost } = setup();
    const before = v.scaleForTest();
    const prevent = vi.fn();
    scrollHost.dispatch('wheel', { deltaY: 0, ctrlKey: true, preventDefault: prevent });
    expect(v.scaleForTest()).toBe(before);
    v.destroy();
  });

  // ⚠ MANDATORY render-epoch test (T3 review finding I-3). Worker path, deferred:
  // dispatch at scale A, setScale(B) mid-flight, resolve — the old-scale bitmap is
  // closed, NOT painted, and a fresh dispatch at scale B repaints the slot.
  it('render epoch: a bitmap dispatched at the old scale is dropped (closed, not painted) after setScale, and re-dispatched at the new scale', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(20, [{ widthPt: 100, heightPt: 200 }], 'worker', true);
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
      overscan: 1,
      zoomMin: 0.5,
      zoomMax: 3,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();

    // Page 0's bitmap is dispatched at the base (old) scale but NOT yet resolved.
    const page0Old = engine.bitmapCalls.find((c) => c.page === 0);
    expect(page0Old).toBeDefined();
    const oldDispatchCount = engine.bitmapCalls.filter((c) => c.page === 0).length;
    expect(oldDispatchCount).toBe(1);
    const oldWidth = page0Old!.width;

    // Zoom mid-flight: bumps the render epoch and re-mounts. Page 0's re-mount is
    // coalesced away by the in-flight guard (same index still in flight).
    v.setScale(v.scaleForTest() * 2);
    expect(v.mountedPageIndicesForTest()).toContain(0); // page 0 still visible at top
    expect(engine.bitmapCalls.filter((c) => c.page === 0).length).toBe(1); // no double-dispatch yet

    // The OLD-scale render resolves. Epoch moved ⇒ STALE: close, do not paint,
    // then re-dispatch page 0 at the NEW scale.
    const bmpOld = engine.createdBitmaps[engine.bitmapCalls.indexOf(page0Old!)];
    page0Old!.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(bmpOld.close.mock.calls.length).toBeGreaterThan(0); // old bitmap closed
    // A fresh dispatch for page 0 was issued at the new scale.
    const page0Calls = engine.bitmapCalls.filter((c) => c.page === 0);
    expect(page0Calls.length).toBeGreaterThanOrEqual(2);
    const freshCall = page0Calls[page0Calls.length - 1];
    // The fresh dispatch's width reflects the NEW scale (double the old width).
    expect(freshCall.width).toBeGreaterThan(oldWidth ?? 0);

    // The stale old bitmap was NOT transferred; only the fresh one paints.
    freshCall.resolve();
    await Promise.resolve();
    await Promise.resolve();
    const slot0 = scrollHost.children.find(
      (c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas') && c.style.top === '0px',
    ) as FakeEl | undefined;
    expect(slot0).toBeDefined();
    const canvas0 = slot0!.children.find((k) => k.tag === 'canvas') as FakeEl;
    expect(canvas0._bitmapCtx?.lastBitmap).not.toBe(bmpOld); // never painted the stale bitmap
    expect(canvas0._bitmapCtx?.lastBitmap).toBeTruthy(); // painted the fresh one
    v.destroy();
  });

  // B1: epoch-then-reject must stay BOUNDED. setScale bumps the epoch mid-flight,
  // so the OLD dispatch is stale on resolution; whether it resolves or REJECTS,
  // the finally must issue exactly ONE fresh dispatch at the new epoch. The fresh
  // dispatch captures the new epoch, so if IT later rejects (same epoch, still
  // live) nothing re-dispatches — the retry is bounded to a single fresh attempt.
  it('render epoch: rejecting the OLD-scale dispatch after setScale still yields exactly ONE fresh dispatch at the new scale (bounded)', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(20, [{ widthPt: 100, heightPt: 200 }], 'worker', true);
    const onError = vi.fn();
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
      overscan: 1,
      zoomMin: 0.5,
      zoomMax: 3,
      onError,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();

    const page0Old = engine.bitmapCalls.find((c) => c.page === 0);
    expect(page0Old).toBeDefined();
    const dispatchesForPage0 = () => engine.bitmapCalls.filter((c) => c.page === 0).length;
    expect(dispatchesForPage0()).toBe(1);
    const oldWidth = page0Old!.width;

    // Zoom mid-flight: bumps the epoch. Page 0's re-mount is coalesced away.
    v.setScale(v.scaleForTest() * 2);
    expect(v.mountedPageIndicesForTest()).toContain(0);
    expect(dispatchesForPage0()).toBe(1);

    // REJECT the old-scale dispatch. Epoch moved ⇒ stale ⇒ re-dispatch (not a plain
    // failure retry). Exactly ONE fresh dispatch at the new epoch.
    page0Old!.reject(new Error('old-scale render failed'));
    for (let k = 0; k < 8; k++) await Promise.resolve();
    const afterReject = engine.bitmapCalls.filter((c) => c.page === 0);
    expect(afterReject.length).toBe(2); // one fresh dispatch, no storm
    const freshCall = afterReject[afterReject.length - 1];
    expect(freshCall.width).toBeGreaterThan(oldWidth ?? 0); // at the NEW (larger) scale

    // The fresh dispatch then SUCCEEDS and paints — the page is not left blank.
    freshCall.resolve();
    for (let k = 0; k < 8; k++) await Promise.resolve();
    expect(dispatchesForPage0()).toBe(2); // still exactly two; success does not re-dispatch
    const slot0 = scrollHost.children.find(
      (c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas') && c.style.top === '0px',
    ) as FakeEl | undefined;
    expect(slot0).toBeDefined();
    const canvas0 = slot0!.children.find((k) => k.tag === 'canvas') as FakeEl;
    expect(canvas0._bitmapCtx?.lastBitmap).toBeTruthy(); // painted the fresh bitmap
    v.destroy();
  });

  it('setScale to the SAME scale is a no-op (no re-anchor, no epoch bump churn)', () => {
    const { v, scrollHost } = setup();
    scrollHost.scrollTop = 1230;
    scrollHost.dispatch('scroll');
    const topBefore = scrollHost.scrollTop;
    v.setScale(v.scaleForTest()); // identical scale
    expect(scrollHost.scrollTop).toBe(topBefore); // unchanged
    v.destroy();
  });
});

describe('DocxScrollViewer — self-load path (T7 story)', () => {
  it('load(url) lays out and mounts the first window (relayout wired into load)', async () => {
    installDom();
    const container = makeContainer(200, 400);
    // Mock the static loader so load() resolves to a fake engine WITHOUT touching
    // a real Worker / WASM. The viewer must call relayout() after assignment so a
    // self-loaded (non-injected) viewer is not left blank (I-2).
    const engine = new FakeDocxEngine(10, [{ widthPt: 100, heightPt: 200 }]);
    const loadSpy = vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
    const v = new DocxScrollViewer(container as unknown as HTMLElement, { gap: 10 });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    await v.load('sample.docx');
    expect(loadSpy).toHaveBeenCalledTimes(1);
    // Layout happened: slots mounted and the spacer was sized.
    expect(v.mountedPageIndicesForTest().length).toBeGreaterThan(0);
    const spacer = scrollHost.children[0] as FakeEl;
    expect(parseFloat(spacer.style.height)).toBeGreaterThan(0);
    // The mounted pages were actually rendered (main mode → renderPage).
    expect(engine.renderCalls.length).toBeGreaterThan(0);
    v.destroy();
  });
});

describe('DocxScrollViewer — text selection (T5)', () => {
  const RUN = { text: 'Hi', x: 1, y: 2, w: 10, h: 12, fontSize: 12, font: '12px serif' };

  function findSlotWrapper(scrollHost: FakeEl): FakeEl {
    return scrollHost.children.find((c) => c.children.some((k) => k.tag === 'canvas')) as FakeEl;
  }
  function findTextLayer(slotWrapper: FakeEl): FakeEl {
    return slotWrapper.children.find((c) => c.tag === 'div') as FakeEl;
  }

  it('main mode: builds an overlay span per run for each visible slot', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(3, [{ widthPt: 100, heightPt: 200 }]);
    engine.feedTextRuns = [RUN];
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      enableTextSelection: true,
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    await Promise.resolve();
    await Promise.resolve();
    // Every mounted slot got its overlay built from that page's render.
    const mounted = v.mountedPageIndicesForTest();
    expect(mounted.length).toBeGreaterThan(0);
    for (const wrapper of scrollHost.children.filter((c) => c.children.some((k) => k.tag === 'canvas'))) {
      const textLayer = findTextLayer(wrapper);
      expect(textLayer.children.length).toBe(1);
      expect(textLayer.children[0].textContent).toBe('Hi');
    }
    v.destroy();
  });

  it('main mode: sizes the overlay to the slot canvas CSS size', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(1, [{ widthPt: 100, heightPt: 200 }]);
    engine.feedTextRuns = [RUN];
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      enableTextSelection: true,
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    await Promise.resolve();
    await Promise.resolve();
    const wrapper = findSlotWrapper(scrollHost);
    const canvas = wrapper.children.find((c) => c.tag === 'canvas') as FakeEl;
    const textLayer = findTextLayer(wrapper);
    // buildDocxTextLayer sizes the layer to the canvas CSS width/height. The
    // main renderPage owns the canvas so its CSS size falls back to
    // `${canvas.width}px` — either way the overlay must match the canvas.
    const expectW = canvas.style.width || `${canvas.width}px`;
    const expectH = canvas.style.height || `${canvas.height}px`;
    expect(textLayer.style.width).toBe(expectW);
    expect(textLayer.style.height).toBe(expectH);
    v.destroy();
  });

  it('main mode: recycling a slot clears its overlay so the free pool holds no stale spans', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(50, [{ widthPt: 100, heightPt: 200 }]);
    engine.feedTextRuns = [RUN];
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      enableTextSelection: true,
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    // The overlay is built in the renderPage .then() microtask.
    await Promise.resolve();
    await Promise.resolve();
    // A page-0 slot exists with a built overlay.
    const wrapper = findSlotWrapper(scrollHost);
    const textLayer = findTextLayer(wrapper);
    expect(textLayer.children.length).toBe(1);
    // Scroll far away so page 0 recycles into the free pool.
    scrollHost.scrollTop = 8000;
    scrollHost.dispatch('scroll');
    // The recycled slot's overlay was cleared (no stale spans linger in the pool).
    expect(textLayer.children.length).toBe(0);
    v.destroy();
  });

  it('worker mode: warns once and builds no overlay spans', async () => {
    installDom();
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(3, [{ widthPt: 100, heightPt: 200 }], 'worker');
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      enableTextSelection: true,
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    await Promise.resolve();
    await Promise.resolve();
    // Warned exactly once with the same wording as DocxViewer, across every
    // mounted slot's render.
    expect(warn).toHaveBeenCalledTimes(1);
    expect(warn.mock.calls[0][0]).toMatch(/text selection is unavailable in mode: 'worker'/);
    // No renderPage (worker path only) and no overlay spans anywhere.
    expect(engine.renderCalls.length).toBe(0);
    for (const wrapper of scrollHost.children.filter((c) => c.children.some((k) => k.tag === 'canvas'))) {
      const textLayer = wrapper.children.find((c) => c.tag === 'div') as FakeEl;
      expect(textLayer.children.length).toBe(0);
    }
    warn.mockRestore();
    v.destroy();
  });

  it('stale main render (epoch moved by setScale mid-flight) does NOT rebuild the overlay', async () => {
    installDom();
    const container = makeContainer(200, 400);
    // Deferred main mode: the test resolves renderPage manually so setScale can
    // move the epoch WHILE page 0's first render is in flight.
    const engine = new FakeDocxEngine(20, [{ widthPt: 100, heightPt: 200 }], 'main', true);
    engine.feedTextRuns = [RUN];
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      enableTextSelection: true,
      gap: 10,
      zoomMin: 0.1,
      zoomMax: 4,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    // Grab the first (stale) page-0 render call before resolving it. Capture the
    // slot wrapper it targets so we can inspect its overlay afterwards.
    const staleWrapper = scrollHost.children.find((c) => c.children.some((k) => k.tag === 'canvas')) as FakeEl;
    const staleLayer = staleWrapper.children.find((c) => c.tag === 'div') as FakeEl;
    const staleCall = engine.renderCalls.find((c) => c.page === 0)!;
    // Zoom mid-flight: bumps the epoch and force-re-mounts every slot (which
    // re-dispatches a fresh page-0 render at the new scale).
    v.setScale(v.scaleForTest() * 2);
    // Now resolve the STALE (old-epoch) render. Its .then must bail on the epoch
    // guard and NOT build the overlay.
    staleCall.resolve();
    await Promise.resolve();
    await Promise.resolve();
    // The stale render built nothing (epoch guard). The fresh render is still
    // deferred (not resolved), so its overlay isn't built either — the stale
    // slot/layer holds zero spans from the superseded render.
    expect(staleLayer.children.length).toBe(0);
    v.destroy();
  });
});

describe('DocxScrollViewer — navigation, resize, empty (T6)', () => {
  const PT_TO_PX = 4 / 3;
  // container fit width 200, page widthPt 100 → base scale = 200/(100*PT_TO_PX)=1.5.
  // page height px = heightPt * PT_TO_PX * scale = 200 * (4/3) * 1.5 = 400.
  // (The plan text mislabelled this as 200px; the viewer fits WIDTH, which scales
  //  height too — verified against the T2/T4 blocks' identical geometry.)
  const BASE = 200 / (100 * PT_TO_PX);
  const PAGE_H = 200 * PT_TO_PX * BASE; // 400
  const GAP = 10;
  const STRIDE = PAGE_H + GAP; // 410 — the top-to-top distance between pages

  function setup(pageCount = 20, opts = {}) {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(
      pageCount,
      Array.from({ length: pageCount }, () => ({ widthPt: 100, heightPt: 200 })),
    );
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: GAP,
      ...opts,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    return { v, scrollHost, container, engine };
  }

  it('scrollToPage sets scrollTop to the page offset and clamps out-of-range', () => {
    const { v, scrollHost } = setup();
    // Pick a clientHeight that makes maxTop STRICTLY LESS than the last page's top
    // offset, so the `Math.min(maxTop, target)` clamp is actually exercised. With the
    // default clientHeight 400, offsets[19] === maxTop and the clamp is a degenerate
    // no-op. totalHeight = 8190, offsets[19] = 19*STRIDE = 7790.
    scrollHost.clientHeight = 500; // maxTop = 8190 − 500 = 7690 < offsets[19] 7790
    v.scrollToPage(3);
    expect(Math.abs(scrollHost.scrollTop - 3 * STRIDE)).toBeLessThan(2);
    v.scrollToPage(999); // clamps to last page index (19)
    // Page 19's top offset (7790) exceeds the max scroll top (7690), so scrollToPage
    // pins to the max (totalHeight − viewportHeight). This asserts the Math.min clamp
    // fired: scrollTop is maxTop (7690), STRICTLY below offsets[19] (7790).
    const totalHeight = 20 * PAGE_H + 19 * GAP; // Σheights + Σgaps = 8190
    const maxTop = totalHeight - 500; // − viewport height = 7690
    const lastOffset = 19 * STRIDE; // 7790
    expect(maxTop).toBeLessThan(lastOffset); // precondition: clamp is exercised
    expect(Math.abs(scrollHost.scrollTop - maxTop)).toBeLessThan(2);
    v.scrollToPage(-5); // clamps to 0
    expect(scrollHost.scrollTop).toBe(0);
    v.destroy();
  });

  it('scrollToPage: viewport taller than total content ⇒ scrollTop pinned to 0 (negative maxTop guard)', () => {
    // A 2-page document is shorter than a very tall viewport, so
    // totalHeight − clientHeight is NEGATIVE; the `Math.max(0, …)` maxTop guard must
    // pin scrollTop to 0 rather than a negative top.
    const { v, scrollHost } = setup(2);
    scrollHost.clientHeight = 5000; // >> totalHeight (2*400 + 10 = 810)
    v.scrollToPage(1); // last page; its offset (410) is above maxTop (0)
    expect(scrollHost.scrollTop).toBe(0);
    v.destroy();
  });

  it('onVisiblePageChange fires only when topIndex changes', () => {
    const changes: number[] = [];
    const { v, scrollHost } = setup(20, { onVisiblePageChange: (i: number) => changes.push(i) });
    scrollHost.scrollTop = STRIDE; // page 1 top
    scrollHost.dispatch('scroll');
    scrollHost.scrollTop = STRIDE + 10; // still within page 1
    scrollHost.dispatch('scroll');
    scrollHost.scrollTop = 3 * STRIDE; // page 3 top
    scrollHost.dispatch('scroll');
    // Deduped: 0 (initial) → 1 → 3, no repeat for the STRIDE+10 no-op.
    expect(changes).toEqual([0, 1, 3]);
    v.destroy();
  });

  it('scrollToPage does not re-fire onVisiblePageChange for the same top page', () => {
    const changes: number[] = [];
    const { v } = setup(20, { onVisiblePageChange: (i: number) => changes.push(i) });
    // Initial mount fired 0. scrollToPage(0) lands on the same top page — no new fire.
    v.scrollToPage(0);
    expect(changes).toEqual([0]);
    // Navigate to page 5 → one fire; navigate there again → no duplicate.
    v.scrollToPage(5);
    v.scrollToPage(5);
    expect(changes).toEqual([0, 5]);
    v.destroy();
  });

  it('ResizeObserver re-fit preserves the zoom multiplier', () => {
    // zoomMax headroom: base 1.5, 2× zoom → 3.0; after the width doubles the new
    // base is 3.0 so the preserved scale is 6.0. zoomMin/zoomMax are ABSOLUTE
    // px-per-pt bounds (design §8.1), so a low zoomMax would legitimately CLAMP
    // the re-fit and break preservation. Give the multiplier room to survive.
    const { v, container } = setup(20, { zoomMax: 10 });
    v.setScale(v.baseScaleForTest() * 2); // 2× zoom
    const zoomMultiplier = v.scaleForTest() / v.baseScaleForTest();
    // Container widens; the observer callback re-fits base and re-applies the mult.
    container.clientWidth = 400;
    (container.children[0] as FakeEl).children[0].clientWidth = 400;
    v.resizeForTest(); // fires the observed resize path
    expect(v.scaleForTest() / v.baseScaleForTest()).toBeCloseTo(zoomMultiplier, 5);
    v.destroy();
  });

  it('ResizeObserver re-fit bumps the render epoch (resize is an epoch event)', () => {
    const { v, container } = setup();
    const before = v.renderEpochForTest();
    container.clientWidth = 400;
    (container.children[0] as FakeEl).children[0].clientWidth = 400;
    v.resizeForTest();
    // A width change re-fits the base scale (routed through setScale), which bumps
    // the epoch so any bitmap in flight at the old scale is treated as stale.
    expect(v.renderEpochForTest()).toBeGreaterThan(before);
    v.destroy();
  });

  it('ResizeObserver re-fit with an UNCHANGED width does not re-fit or bump the epoch', () => {
    const { v } = setup();
    const scaleBefore = v.scaleForTest();
    const epochBefore = v.renderEpochForTest();
    v.resizeForTest(); // same width — no-op re-fit
    expect(v.scaleForTest()).toBe(scaleBefore);
    expect(v.renderEpochForTest()).toBe(epochBefore);
    v.destroy();
  });

  it('height-only resize (width unchanged) re-mounts the newly-revealed window without a scroll, no epoch bump, no callback', () => {
    // F1 regression: a height-only grow used to early-return before `_mountVisible`,
    // leaving the newly-revealed rows blank until the user scrolled. At scrollTop 0
    // the initially-mounted window is [0,1] (viewport 400 = one page; +1 overscan).
    // Growing the viewport to 1600 must mount [0..4] purely from the resize — no
    // scroll event, no epoch bump (geometry/scale unchanged so cached canvases stay
    // valid), and no onVisiblePageChange fire (topIndex stays 0 at scrollTop 0).
    const changes: number[] = [];
    const { v, scrollHost } = setup(20, { onVisiblePageChange: (i: number) => changes.push(i) });
    const mountedBefore = v.mountedPageIndicesForTest().slice().sort((a, b) => a - b);
    expect(mountedBefore).toEqual([0, 1]);
    expect(changes).toEqual([0]); // initial mount fired 0
    const epochBefore = v.renderEpochForTest();
    const scaleBefore = v.scaleForTest();

    // Grow HEIGHT only; width (and thus the fit-width base scale) is unchanged.
    scrollHost.clientHeight = 1600;
    v.resizeForTest(); // NO scroll event dispatched

    const mountedAfter = v.mountedPageIndicesForTest().slice().sort((a, b) => a - b);
    // The revealed window grew: bottom edge 1600 now intersects pages 0..3 (+1
    // overscan ⇒ end 4). The mounted span strictly grows.
    expect(mountedAfter).toEqual([0, 1, 2, 3, 4]);
    expect(mountedAfter.length).toBeGreaterThan(mountedBefore.length);
    // No epoch bump — nothing was re-scaled, so no in-flight render is stale.
    expect(v.renderEpochForTest()).toBe(epochBefore);
    expect(v.scaleForTest()).toBe(scaleBefore);
    // topIndex is still 0 at scrollTop 0 ⇒ callback did NOT fire again.
    expect(changes).toEqual([0]);
    v.destroy();
  });

  it('clamped resize (zoomMax clamp + width & height grow) still refreshes the revealed window', () => {
    // F1 clamped-path variant: pin zoomMax to the current base so a width grow's
    // re-fit (newBase × mult) clamps back to the SAME scale, making setScale a no-op
    // (which skips its own force-re-render mount). The post-setScale `_mountVisible`
    // must still mount the rows the taller viewport revealed. base = 1.5, no user
    // zoom ⇒ mult 1; after width→400 newBase 3.0 clamps to zoomMax 1.5 (== _scale).
    const changes: number[] = [];
    const { v, scrollHost, container } = setup(20, {
      zoomMax: BASE, // 1.5 — clamps the re-fit back to the current scale
      onVisiblePageChange: (i: number) => changes.push(i),
    });
    expect(v.scaleForTest()).toBe(BASE);
    const mountedBefore = v.mountedPageIndicesForTest().slice().sort((a, b) => a - b);
    expect(mountedBefore).toEqual([0, 1]);
    const epochBefore = v.renderEpochForTest();

    // Grow BOTH width and height. The width grow triggers the re-fit branch, but the
    // clamp pins the scale unchanged ⇒ setScale no-ops. Heights are therefore
    // unchanged (scale unchanged), so the taller viewport reveals more of the SAME
    // geometry: mounted grows to [0..4].
    container.clientWidth = 400;
    (container.children[0] as FakeEl).children[0].clientWidth = 400;
    scrollHost.clientWidth = 400;
    scrollHost.clientHeight = 1600;
    v.resizeForTest();

    const mountedAfter = v.mountedPageIndicesForTest().slice().sort((a, b) => a - b);
    expect(mountedAfter).toEqual([0, 1, 2, 3, 4]);
    expect(mountedAfter.length).toBeGreaterThan(mountedBefore.length);
    // Scale stayed pinned at the clamp, so setScale no-oped ⇒ no epoch bump.
    expect(v.scaleForTest()).toBe(BASE);
    expect(v.renderEpochForTest()).toBe(epochBefore);
    // topIndex still 0 ⇒ no extra callback.
    expect(changes).toEqual([0]);
    v.destroy();
  });

  it('zero-width container defers, then lays out on first non-zero resize', () => {
    installDom();
    const container = makeContainer(0, 0); // unlaid-out
    const engine = new FakeDocxEngine(5, Array.from({ length: 5 }, () => ({ widthPt: 100, heightPt: 200 })));
    const v = new DocxScrollViewer(container as unknown as HTMLElement, { document: engine.asDoc() });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    expect(v.mountedPageIndicesForTest().length).toBe(0); // deferred
    container.clientWidth = 300;
    scrollHost.clientWidth = 300;
    scrollHost.clientHeight = 400;
    v.resizeForTest();
    expect(v.mountedPageIndicesForTest().length).toBeGreaterThan(0);
    v.destroy();
  });

  it('empty document: pageCount 0 ⇒ spacer 0, no slots, scrollToPage no-op', () => {
    installDom();
    const container = makeContainer(300, 400);
    const engine = new FakeDocxEngine(0, []);
    const v = new DocxScrollViewer(container as unknown as HTMLElement, { document: engine.asDoc() });
    v.relayout();
    const spacer = (container.children[0] as FakeEl).children[0].children[0] as FakeEl;
    expect(parseFloat(spacer.style.height || '0')).toBe(0);
    expect(v.mountedPageIndicesForTest().length).toBe(0);
    v.scrollToPage(0); // no throw
    v.destroy();
  });

  it('empty document: resize does not crash and fires no callback', () => {
    installDom();
    const container = makeContainer(0, 0);
    const engine = new FakeDocxEngine(0, []);
    const changes: number[] = [];
    const v = new DocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      onVisiblePageChange: (i: number) => changes.push(i),
    });
    container.clientWidth = 300;
    (container.children[0] as FakeEl).children[0].clientWidth = 300;
    v.resizeForTest(); // no pages ⇒ no-op
    expect(v.mountedPageIndicesForTest().length).toBe(0);
    expect(changes).toEqual([]);
    v.destroy();
  });

  it('destroy() disconnects the ResizeObserver', () => {
    installDom();
    let disconnected = 0;
    class SpyRO {
      cb: () => void;
      constructor(cb: () => void) {
        this.cb = cb;
      }
      observe(): void {}
      disconnect(): void {
        disconnected++;
      }
    }
    vi.stubGlobal('ResizeObserver', SpyRO);
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(3, [{ widthPt: 100, heightPt: 200 }]);
    const v = new DocxScrollViewer(container as unknown as HTMLElement, { document: engine.asDoc() });
    v.destroy();
    expect(disconnected).toBe(1);
  });
});
