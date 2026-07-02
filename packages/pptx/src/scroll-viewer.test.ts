import { describe, it, expect, afterEach, vi } from 'vitest';
import { PptxScrollViewer } from './scroll-viewer.js';
import { installDom, makeContainer, FakePptxEngine, type FakeEl } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe('PptxScrollViewer — skeleton (T1)', () => {
  it('builds the wrapper → scrollHost → spacer DOM inside the container', () => {
    installDom();
    const container = makeContainer();
    const engine = new FakePptxEngine(3, 1000, 600);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, { presentation: engine.asPres() });
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

  it('exposes slideCount from the injected engine', () => {
    installDom();
    const engine = new FakePptxEngine(5, 1000, 600);
    const v = new PptxScrollViewer(makeContainer() as unknown as HTMLElement, { presentation: engine.asPres() });
    expect(v.slideCount).toBe(5);
    v.destroy();
  });

  it('load() is unsupported when an engine is injected', async () => {
    installDom();
    const engine = new FakePptxEngine(1, 1000, 600);
    const v = new PptxScrollViewer(makeContainer() as unknown as HTMLElement, { presentation: engine.asPres() });
    await expect(v.load('x.pptx')).rejects.toThrow(/injected/i);
    v.destroy();
  });

  it('destroy() removes the DOM and does NOT destroy an injected engine', () => {
    installDom();
    const container = makeContainer();
    const engine = new FakePptxEngine(1, 1000, 600);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, { presentation: engine.asPres() });
    expect(container.children.length).toBe(1); // wrapper mounted
    v.destroy();
    expect(container.children.length).toBe(0); // wrapper removed
    expect(engine.destroyed).toBe(false); // injected engine preserved (caller owns it)
  });

  it('slideCount is 0 before load resolves (no injected engine)', () => {
    installDom();
    const v = new PptxScrollViewer(makeContainer() as unknown as HTMLElement, {});
    expect(v.slideCount).toBe(0);
    v.destroy();
  });

  // O1 (design §11): an injected engine's own `mode` is authoritative. An
  // EXPLICITLY conflicting opts.mode is a mis-configuration rejected at
  // construction; a matching or absent opts.mode constructs fine.
  it('throws when opts.mode conflicts with an injected worker-mode engine', () => {
    installDom();
    const engine = new FakePptxEngine(1, 1000, 600, 'worker');
    expect(
      () =>
        new PptxScrollViewer(makeContainer() as unknown as HTMLElement, {
          presentation: engine.asPres(),
          mode: 'main',
        }),
    ).toThrow(/mode/i);
  });

  it('does NOT throw when opts.mode matches an injected worker-mode engine', () => {
    installDom();
    const engine = new FakePptxEngine(1, 1000, 600, 'worker');
    const v = new PptxScrollViewer(makeContainer() as unknown as HTMLElement, {
      presentation: engine.asPres(),
      mode: 'worker',
    });
    expect(v.slideCount).toBe(1);
    v.destroy();
    // Injected engine is caller-owned even in the worker case: destroy() leaves it intact.
    expect(engine.destroyed).toBe(false);
  });

  it('constructs a default-main injected engine with absent opts.mode (load still rejects; destroy preserves engine)', async () => {
    installDom();
    // Default mode is 'main'; opts.mode is absent ⇒ no conflict, resolved path is main.
    const engine = new FakePptxEngine(2, 1000, 600);
    const v = new PptxScrollViewer(makeContainer() as unknown as HTMLElement, {
      presentation: engine.asPres(),
    });
    expect(v.slideCount).toBe(2);
    await expect(v.load('x.pptx')).rejects.toThrow(/injected/i);
    v.destroy();
    // Injected engine is caller-owned: destroy() must not tear it down.
    expect(engine.destroyed).toBe(false);
  });
});

describe('PptxScrollViewer — layout + virtualization (T2)', () => {
  // Uniform 1000×600 EMU slides; `_scale` is px-per-EMU (no PT_TO_PX). Fit the
  // slide width to the container: base scale = clientWidth / slideWidth. With
  // clientWidth 200 → base 0.2 ⇒ each slide is 200px wide, 120px tall.
  function setup(slideCount: number, opts = {}) {
    const dom = installDom();
    const container = makeContainer(200, 400); // clientWidth 200, clientHeight 400
    const engine = new FakePptxEngine(slideCount, 1000, 600);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
      overscan: 1,
      width: undefined, // fit container width (200) → base scale below
      ...opts,
    });
    const wrapper = container.children[0] as FakeEl;
    const scrollHost = wrapper.children[0] as FakeEl;
    const spacer = scrollHost.children[0] as FakeEl;
    // Drive layout: container width 200, slide width 1000 EMU → base scale maps
    // 1000 EMU to 200px (0.2 px/EMU). Provide the viewport height.
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    return { dom, container, engine, v, wrapper, scrollHost, spacer };
  }

  it('sizes the spacer to computeVisibleRange.totalHeight and mounts only the visible window', () => {
    const { v, scrollHost, spacer } = setup(10);
    v.relayout(); // T2 exposes an explicit relayout() the viewer calls after load/resize
    // Spacer height > 0 and equals n * slideHeightPx + (n-1)*gap in px.
    expect(parseFloat(spacer.style.height)).toBeGreaterThan(0);
    // At scrollTop 0 with a 400px viewport and 120px-tall slides, several slides
    // fit; overscan 1 adds one more. Assert the mounted slot count is bounded.
    const mounted = scrollHost.children.filter((c) => c !== spacer);
    expect(mounted.length).toBeGreaterThanOrEqual(1);
    // Viewport 400 / stride 130 ⇒ ~4 slides visible + overscan; comfortably bounded.
    expect(mounted.length).toBeLessThanOrEqual(6);
    v.destroy();
  });

  it('spacer height is exact for uniform slide heights + gap', () => {
    const dom = installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(3, 1000, 600);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
      overscan: 1,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    const spacer = scrollHost.children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    // base scale = 200 / 1000 = 0.2 px/EMU ⇒ slide height px = 600 * 0.2 = 120.
    const scale = 200 / 1000;
    const slideH = 600 * scale; // 120
    const n = 3;
    const expected = n * slideH + (n - 1) * 10;
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
    // reuses slots rather than appending a new canvas per slide.
    expect(scrollHost.children.length).toBeLessThanOrEqual(initialMount + 2);
  });

  it('scrolling far then back reuses pooled slot wrappers (bounded distinct allocations)', () => {
    const { v, scrollHost } = setup(50);
    v.relayout();
    // Track EVERY distinct wrapper element the viewer ever appends. If slots are
    // pooled (recycled), scrolling across the whole deck then back reuses a small
    // fixed set of wrappers rather than allocating one per visited slide.
    const seen = new Set<FakeEl>();
    const collect = () => {
      for (const c of scrollHost.children) if (c.tag === 'div') seen.add(c);
    };
    collect();
    // Sweep deep into the deck and back to the top.
    for (const top of [3000, 6000, 3000, 0]) {
      scrollHost.scrollTop = top;
      scrollHost.dispatch('scroll');
      collect();
    }
    // 50 slides visited across the sweep; a per-slide allocator would have created
    // dozens of wrappers. The pool must keep the distinct-wrapper count tiny:
    // spacer(1) + at most a couple of window-sized generations of slots.
    expect(seen.size).toBeLessThanOrEqual(12);
    // Coming back to the top mounts slide 0 again, drawn from the pool.
    expect(v.mountedSlideIndicesForTest()).toContain(0);
  });

  it('mounts the correct window for a mid-deck scrollTop', () => {
    const { v, scrollHost } = setup(20);
    v.relayout();
    // Jump to a scrollTop deep in the deck; the mounted slots must include the
    // slide under the viewport top (topVisibleSlide) and its overscan neighbours.
    scrollHost.scrollTop = 1300;
    scrollHost.dispatch('scroll');
    const top = v.topVisibleSlide;
    const mountedSlides = v.mountedSlideIndicesForTest(); // T2 test hook
    expect(mountedSlides).toContain(top);
    expect(Math.max(...mountedSlides) - Math.min(...mountedSlides)).toBeLessThanOrEqual(6);
  });
});

describe('PptxScrollViewer — rendering (T3)', () => {
  it("main mode: calls renderSlide once per mounted slot with that slide's px width", () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(10, 1000, 600);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    // Every mounted slide got exactly one renderSlide call, no more no less.
    const mounted = v.mountedSlideIndicesForTest().sort((a, b) => a - b);
    const rendered = engine.renderCalls.map((c) => c.slide).sort((a, b) => a - b);
    expect(rendered).toEqual(mounted);
    // Each renderSlide call carried the per-call px width. pptx slides are uniform,
    // so every width equals the same fit-width (base scale = 200/1000 = 0.2 px/EMU;
    // slide width px = 1000 * 0.2 = 200), but the per-call width contract still
    // passes a width per slide (mirrors docx's per-page width assertion, §7).
    const widths = engine.renderSlideWidths();
    expect(widths.length).toBe(mounted.length);
    for (const w of widths) expect(w).toBeCloseTo(200, 3);
    v.destroy();
  });

  it('main mode: passes the uniform px width per slide (per-call width contract, §7)', () => {
    installDom();
    const container = makeContainer(200, 400);
    // pptx slide size is UNIFORM deck-wide, so unlike docx (mixed page sizes) every
    // slide renders at the SAME px width. The per-call width contract still holds:
    // each mounted slide receives its own width argument, equal to the uniform px.
    const engine = new FakePptxEngine(2, 1000, 600);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    // Both slides mount (slide 0 height px = 600*0.2 = 120; slide 1 top 130 within
    // the 400px viewport), so both get exactly one render.
    const mounted = v.mountedSlideIndicesForTest().sort((a, b) => a - b);
    expect(mounted).toEqual([0, 1]);
    // Per-slide px widths in call order: uniform 200 each.
    const widths = engine.renderCalls
      .slice()
      .sort((a, b) => a.slide - b.slide)
      .map((c) => c.width ?? NaN);
    expect(widths[0]).toBeCloseTo(200, 3);
    expect(widths[1]).toBeCloseTo(200, 3);
    v.destroy();
  });

  it('does not re-render a mounted slot for the same slide on a no-op scroll', () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(10, 1000, 600);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
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

  it('worker mode: never calls renderSlide — routes every slot through renderSlideToBitmap', async () => {
    installDom();
    const container = makeContainer(200, 400);
    // mode:'worker' — the real PptxPresentation.renderSlide THROWS synchronously in
    // worker mode; a viewer that mis-routed would blow up (and renderCalls would
    // record the attempt). The direct _mode routing must never touch renderSlide.
    const engine = new FakePptxEngine(10, 1000, 600, 'worker');
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    await Promise.resolve();
    await Promise.resolve();
    expect(engine.renderCalls).toHaveLength(0); // renderSlide never touched
    // Every mounted slide dispatched exactly one bitmap render.
    const mounted = v.mountedSlideIndicesForTest().sort((a, b) => a - b);
    const bitmapSlides = [...new Set(engine.bitmapCalls.map((c) => c.slide))].sort((a, b) => a - b);
    expect(bitmapSlides).toEqual(mounted);
    v.destroy();
  });

  it('worker mode: paints a resolved bitmap into the slot canvas (transfer)', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(10, 1000, 600, 'worker');
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
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
    const slotWrapper = scrollHost.children.find((c) => c.children.some((k) => k.tag === 'canvas')) as FakeEl;
    const canvas = slotWrapper.children.find((k) => k.tag === 'canvas') as FakeEl;
    // The bitmaprenderer ctx received the bitmap (fake records lastBitmap).
    expect(canvas._bitmapCtx?.lastBitmap).toBe(engine.createdBitmaps[0]);
    v.destroy();
  });

  it('worker mode: closes the ImageBitmap on recycle (deferred — render in flight when the slot recycles)', async () => {
    // Bitmap-close observability path: the primary path (deterministic slot.bitmap
    // hold under synchronous resolve) does NOT hold — in non-deferred mode
    // transferFromImageBitmap consumes the bitmap and nulls slot.bitmap
    // synchronously BEFORE any scroll-away, so a later recycle has nothing to
    // close. We use the DOCUMENTED FALLBACK: a deferred fake so the render is
    // genuinely in flight when the slot recycles, and the on-resolution
    // stale-check closes the orphan (design §11).
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(50, 1000, 600, 'worker', true);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    const firstBatch = engine.bitmapCalls.length;
    expect(firstBatch).toBeGreaterThan(0);
    // Scroll far away so the first slides' slots recycle while their renders are
    // still in flight (deferred → not yet resolved).
    scrollHost.scrollTop = 6000;
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
    // Deferred: renderSlideToBitmap resolves only when the test calls resolve(),
    // so we can scroll a slot's slide out of the window BEFORE the bitmap arrives.
    const engine = new FakePptxEngine(50, 1000, 600, 'worker', true);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    // Slide 0's bitmap is dispatched but NOT yet resolved.
    const slide0 = engine.bitmapCalls.find((c) => c.slide === 0);
    expect(slide0).toBeDefined();
    // Scroll far away so slide 0's slot recycles while its render is in flight.
    scrollHost.scrollTop = 6000;
    scrollHost.dispatch('scroll');
    expect(v.mountedSlideIndicesForTest()).not.toContain(0);
    // Now the stale render for slide 0 resolves — the viewer must NOT paint it and
    // must close the orphaned bitmap.
    const bmp0 = engine.createdBitmaps[engine.bitmapCalls.indexOf(slide0!)];
    slide0!.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(bmp0.close.mock.calls.length).toBeGreaterThan(0); // orphan closed
    v.destroy();
  });

  it('worker mode: a slide that recycles then re-mounts while in flight still gets a fresh render', async () => {
    // Coalescing keys on slide index; a naive Set<number> would swallow the
    // re-mounted slot's render (the stale resolve clears in-flight AFTER the
    // remount coalesced away → the new slot never paints). The viewer must
    // re-dispatch the live slot's render so the slide never stays blank.
    installDom();
    const container = makeContainer(200, 400);
    // Tall slides (1000×2000 EMU ⇒ 200×400px at base scale 0.2, stride 410) so the
    // visible window is exactly one slide + overscan = [0,1], mirroring the docx
    // recycle/re-mount pool dynamics: the re-mounted slide-0 slot is a DIFFERENT
    // pooled object than the in-flight one, so `live !== slot` triggers the
    // re-dispatch. (A short-slide window packs several slots, whose LIFO reuse can
    // hand slide 0 back its own in-flight slot — an epoch-only case not under test.)
    const engine = new FakePptxEngine(50, 1000, 2000, 'worker', true);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    const slide0Initial = engine.bitmapCalls.find((c) => c.slide === 0);
    expect(slide0Initial).toBeDefined();
    const dispatchesForSlide0 = () => engine.bitmapCalls.filter((c) => c.slide === 0).length;
    expect(dispatchesForSlide0()).toBe(1);
    // Scroll away (slide 0 recycles, render still in flight) then back to the top
    // (slide 0 re-mounts on a fresh slot) — all while the initial render is deferred.
    scrollHost.scrollTop = 12000;
    scrollHost.dispatch('scroll');
    expect(v.mountedSlideIndicesForTest()).not.toContain(0);
    scrollHost.scrollTop = 0;
    scrollHost.dispatch('scroll');
    expect(v.mountedSlideIndicesForTest()).toContain(0);
    // Coalescing pin: slide 0's render is STILL in flight (initial deferred call not
    // yet resolved), so the re-mount must NOT dispatch a second render for slide 0
    // — the in-flight guard swallows it. Exactly one dispatch so far.
    expect(dispatchesForSlide0()).toBe(1);
    // The initial (now stale) render resolves — the orphan is dropped and the
    // re-mounted slot's render is (re-)dispatched.
    slide0Initial!.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(dispatchesForSlide0()).toBeGreaterThanOrEqual(2); // fresh render issued
    // Resolve the fresh render and confirm it paints into the live slot.
    for (const c of engine.bitmapCalls.filter((c) => c.slide === 0)) c.resolve();
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
    // `!painted`. When `renderSlideToBitmap` rejects while the slot is still live
    // and the epoch is unchanged, `painted` is false and `live === slot`, but
    // NEITHER staleness test fires, so we must NOT re-dispatch. A `!painted`-only
    // gate would loop reject → re-dispatch → reject … unbounded (empirically 1→2
    // →3→4 with onError every round). The onError contract leaves the slide blank.
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(50, 1000, 600, 'worker', true);
    const onError = vi.fn();
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
      overscan: 1,
      onError,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();

    const slide0 = engine.bitmapCalls.find((c) => c.slide === 0);
    expect(slide0).toBeDefined();
    expect(v.mountedSlideIndicesForTest()).toContain(0); // slot still live
    const dispatchesForSlide0 = () => engine.bitmapCalls.filter((c) => c.slide === 0).length;
    expect(dispatchesForSlide0()).toBe(1);

    // Reject the render with the slot STILL LIVE and no scale change (epoch fixed).
    slide0!.reject(new Error('worker render failed'));
    // Flush microtasks generously — a retry storm would keep queuing dispatches.
    for (let k = 0; k < 8; k++) await Promise.resolve();

    // Dispatch count for slide 0 stayed at 1 — the storm is gone.
    expect(dispatchesForSlide0()).toBe(1);
    // onError fired exactly once (the single failure), never again.
    expect(onError).toHaveBeenCalledTimes(1);
    // Still no further dispatches after more microtask flushing.
    for (let k = 0; k < 8; k++) await Promise.resolve();
    expect(dispatchesForSlide0()).toBe(1);
    expect(onError).toHaveBeenCalledTimes(1);
    v.destroy();
  });

  it('worker mode: destroy() mid-flight closes the resolving bitmap and does not fire onError post-destroy', async () => {
    installDom();
    const container = makeContainer(200, 400);
    // Deferred so a render is genuinely in flight when we destroy the viewer.
    const engine = new FakePptxEngine(50, 1000, 600, 'worker', true);
    const onError = vi.fn();
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
      onError,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    const slide0 = engine.bitmapCalls.find((c) => c.slide === 0);
    expect(slide0).toBeDefined();
    // Tear down while slide 0's render is still in flight. destroy() recycles the
    // slot (no bitmap held yet — none received), so the on-resolution stale-check
    // must close the orphan; the identity guard fails (slot no longer live).
    v.destroy();
    const bmp0 = engine.createdBitmaps[engine.bitmapCalls.indexOf(slide0!)];
    slide0!.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(bmp0.close.mock.calls.length).toBeGreaterThan(0); // orphan closed, no leak
    expect(onError).not.toHaveBeenCalled(); // no error surfaced after teardown
  });

  // ⚠ MANDATORY render-epoch test (T3 review finding I-3). Worker path, deferred:
  // dispatch at scale A, setScale(B) mid-flight, resolve — the old-scale bitmap is
  // closed, NOT painted, and a fresh dispatch at scale B repaints the slot.
  it('render epoch: a bitmap dispatched at the old scale is dropped (closed, not painted) after setScale, and re-dispatched at the new scale', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(20, 1000, 600, 'worker', true);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
      overscan: 1,
      zoomMin: 0.05,
      zoomMax: 3,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();

    // Slide 0's bitmap is dispatched at the base (old) scale but NOT yet resolved.
    const slide0Old = engine.bitmapCalls.find((c) => c.slide === 0);
    expect(slide0Old).toBeDefined();
    const oldDispatchCount = engine.bitmapCalls.filter((c) => c.slide === 0).length;
    expect(oldDispatchCount).toBe(1);
    const oldWidth = slide0Old!.width;

    // Zoom mid-flight: bumps the render epoch and re-mounts. Slide 0's re-mount is
    // coalesced away by the in-flight guard (same index still in flight).
    v.setScale(v.scaleForTest() * 2);
    expect(v.mountedSlideIndicesForTest()).toContain(0); // slide 0 still visible at top
    expect(engine.bitmapCalls.filter((c) => c.slide === 0).length).toBe(1); // no double-dispatch yet

    // The OLD-scale render resolves. Epoch moved ⇒ STALE: close, do not paint,
    // then re-dispatch slide 0 at the NEW scale.
    const bmpOld = engine.createdBitmaps[engine.bitmapCalls.indexOf(slide0Old!)];
    slide0Old!.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(bmpOld.close.mock.calls.length).toBeGreaterThan(0); // old bitmap closed
    // A fresh dispatch for slide 0 was issued at the new scale.
    const slide0Calls = engine.bitmapCalls.filter((c) => c.slide === 0);
    expect(slide0Calls.length).toBeGreaterThanOrEqual(2);
    const freshCall = slide0Calls[slide0Calls.length - 1];
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
    const engine = new FakePptxEngine(20, 1000, 600, 'worker', true);
    const onError = vi.fn();
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
      overscan: 1,
      zoomMin: 0.05,
      zoomMax: 3,
      onError,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();

    const slide0Old = engine.bitmapCalls.find((c) => c.slide === 0);
    expect(slide0Old).toBeDefined();
    const dispatchesForSlide0 = () => engine.bitmapCalls.filter((c) => c.slide === 0).length;
    expect(dispatchesForSlide0()).toBe(1);
    const oldWidth = slide0Old!.width;

    // Zoom mid-flight: bumps the epoch. Slide 0's re-mount is coalesced away.
    v.setScale(v.scaleForTest() * 2);
    expect(v.mountedSlideIndicesForTest()).toContain(0);
    expect(dispatchesForSlide0()).toBe(1);

    // REJECT the old-scale dispatch. Epoch moved ⇒ stale ⇒ re-dispatch (not a plain
    // failure retry). Exactly ONE fresh dispatch at the new epoch.
    slide0Old!.reject(new Error('old-scale render failed'));
    for (let k = 0; k < 8; k++) await Promise.resolve();
    const afterReject = engine.bitmapCalls.filter((c) => c.slide === 0);
    expect(afterReject.length).toBe(2); // one fresh dispatch, no storm
    const freshCall = afterReject[afterReject.length - 1];
    expect(freshCall.width).toBeGreaterThan(oldWidth ?? 0); // at the NEW (larger) scale

    // The fresh dispatch then SUCCEEDS and paints — the slide is not left blank.
    freshCall.resolve();
    for (let k = 0; k < 8; k++) await Promise.resolve();
    expect(dispatchesForSlide0()).toBe(2); // still exactly two; success does not re-dispatch
    const slot0 = scrollHost.children.find(
      (c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas') && c.style.top === '0px',
    ) as FakeEl | undefined;
    expect(slot0).toBeDefined();
    const canvas0 = slot0!.children.find((k) => k.tag === 'canvas') as FakeEl;
    expect(canvas0._bitmapCtx?.lastBitmap).toBeTruthy(); // painted the fresh bitmap
    v.destroy();
  });
});

describe('PptxScrollViewer — zoom (T4)', () => {
  // Uniform 1000×600 EMU slides; `_scale` is px-per-EMU (no PT_TO_PX). Container
  // 200×400, width:undefined ⇒ base fit maps the slide width (1000 EMU) to 200px:
  //   base = 200 / 1000 = 0.2 px/EMU
  //   slide height px = 600 * 0.2 = 120
  //   offset[i] = i * (120 + gap) = i * 130
  // Zoom ×2 ⇒ scale 0.4, height 240, stride 250. The two MANDATORY render-epoch
  // tests (resolve + reject) live in the T3 rendering block above (already covered
  // by WS4b-1), so they are not repeated here.
  function setup(slideCount = 20, opts = {}) {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(slideCount, 1000, 600);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
      overscan: 1,
      // Give the base fit (0.2) headroom: base ×2 = 0.4 stays within these bounds
      // so the re-anchor tests exercise the zoom path, not the clamp.
      zoomMin: 0.05,
      zoomMax: 3,
      ...opts,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    return { v, scrollHost, engine, container };
  }

  it('base fit scale maps the slide width to the container width', () => {
    const { v } = setup();
    // base = 200 / 1000 = 0.2
    expect(v.scaleForTest()).toBeCloseTo(0.2, 6);
    expect(v.baseScaleForTest()).toBeCloseTo(0.2, 6);
    v.destroy();
  });

  it('re-anchors so the slide under the viewport top stays fixed across a zoom (intraFrac 0)', () => {
    const { v, scrollHost } = setup();
    // slide 3 top offset at base (0.2): 3 * (120 + 10) = 390
    scrollHost.scrollTop = 390;
    scrollHost.dispatch('scroll');
    expect(v.topVisibleSlide).toBe(3);
    const cur = v.scaleForTest(); // 0.2
    v.setScale(cur * 2); // 0.4 (within [0.05, 3], no clamp)
    expect(v.scaleForTest()).toBeCloseTo(0.4, 6);
    // slide 3 stays the top slide; new offset = 3 * (240 + 10) = 750
    expect(v.topVisibleSlide).toBe(3);
    expect(Math.abs(scrollHost.scrollTop - 750)).toBeLessThan(2);
    v.destroy();
  });

  it('re-anchors preserving the intra-slide fraction (intraFrac ≠ 0)', () => {
    const { v, scrollHost } = setup();
    // Scroll so HALF of slide 3 (60 of its 120px) has passed above the viewport
    // top: scrollTop = offset[3] + 0.5*120 = 390 + 60 = 450 → intraFrac 0.5.
    scrollHost.scrollTop = 450;
    scrollHost.dispatch('scroll');
    expect(v.topVisibleSlide).toBe(3);
    v.setScale(v.scaleForTest() * 2); // 0.2 → 0.4; heights' = 240
    // newScrollTop = offset'[3] + 0.5 * 240 = 750 + 120 = 870
    expect(Math.abs(scrollHost.scrollTop - 870)).toBeLessThan(2);
    // Slide 3 is still under the viewport top: offset'[3]=750 <= 870 < 1000.
    expect(v.topVisibleSlide).toBe(3);
    v.destroy();
  });

  it('clamps setScale to the absolute [zoomMin, zoomMax] px-per-EMU bounds', () => {
    const { v } = setup();
    v.setScale(100); // above zoomMax 3 (absolute, NOT a multiple of base)
    expect(v.scaleForTest()).toBeCloseTo(3, 6);
    v.setScale(0.001); // below zoomMin 0.05
    expect(v.scaleForTest()).toBeCloseTo(0.05, 6);
    v.destroy();
  });

  it('setScale defaults clamp to [0.1, 4] when zoomMin/zoomMax are unset', () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(5, 1000, 600);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
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

  it('setScale to the SAME scale is a no-op (no re-anchor, no epoch bump churn)', () => {
    const { v, scrollHost } = setup();
    scrollHost.scrollTop = 390;
    scrollHost.dispatch('scroll');
    const topBefore = scrollHost.scrollTop;
    const epochBefore = v.renderEpochForTest();
    v.setScale(v.scaleForTest()); // identical scale
    expect(scrollHost.scrollTop).toBe(topBefore); // unchanged
    expect(v.renderEpochForTest()).toBe(epochBefore); // no epoch bump
    v.destroy();
  });
});

describe('PptxScrollViewer — text selection (T5)', () => {
  // A minimal PptxTextRunInfo: richer than docx's flat run — it carries per-shape
  // frame geometry (shapeX/Y/W/H, rotation) so buildPptxTextLayer can group runs
  // into a positioned shape <div> and nest the run <span> inside it.
  const RUN = {
    text: 'Hi',
    inShapeX: 1,
    inShapeY: 2,
    w: 10,
    h: 12,
    fontSize: 12,
    font: '12px serif',
    shapeX: 0,
    shapeY: 0,
    shapeW: 100,
    shapeH: 40,
    rotation: 0,
  };

  function findSlotWrapper(scrollHost: FakeEl): FakeEl {
    return scrollHost.children.find((c) => c.children.some((k) => k.tag === 'canvas')) as FakeEl;
  }
  function findTextLayer(slotWrapper: FakeEl): FakeEl {
    return slotWrapper.children.find((c) => c.tag === 'div') as FakeEl;
  }

  it('main mode: builds a shape div with a nested run span for each visible slot', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(3, 1000, 600);
    engine.feedTextRuns = [RUN];
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      enableTextSelection: true,
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    await Promise.resolve();
    await Promise.resolve();
    // Every mounted slot got its overlay built from that slide's render.
    const mounted = v.mountedSlideIndicesForTest();
    expect(mounted.length).toBeGreaterThan(0);
    for (const wrapper of scrollHost.children.filter((c) => c.children.some((k) => k.tag === 'canvas'))) {
      const textLayer = findTextLayer(wrapper);
      // pptx nests spans INSIDE a per-shape <div> (one level deeper than docx):
      // textLayer → shape <div> → run <span>.
      expect(textLayer.children.length).toBe(1);
      const shapeDiv = textLayer.children[0] as FakeEl;
      expect(shapeDiv.tag).toBe('div');
      expect(shapeDiv.children.length).toBe(1);
      expect(shapeDiv.children[0].tag).toBe('span');
      expect(shapeDiv.children[0].textContent).toBe('Hi');
    }
    v.destroy();
  });

  it('main mode: a rotated shape gets a CSS rotate() on its shape div', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(1, 1000, 600);
    // Rotated shape frame: buildPptxTextLayer applies transform:rotate(totalRot).
    engine.feedTextRuns = [{ ...RUN, rotation: 30 }];
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      enableTextSelection: true,
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    await Promise.resolve();
    await Promise.resolve();
    const wrapper = findSlotWrapper(scrollHost);
    const textLayer = findTextLayer(wrapper);
    const shapeDiv = textLayer.children[0] as FakeEl;
    expect(shapeDiv.tag).toBe('div');
    expect(shapeDiv.style.transform).toBe('rotate(30deg)');
    // The run span still nests inside the rotated shape div.
    expect(shapeDiv.children[0].textContent).toBe('Hi');
    v.destroy();
  });

  it('main mode: sizes the overlay to the slot canvas size (number args)', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(1, 1000, 600);
    engine.feedTextRuns = [RUN];
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
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
    // buildPptxTextLayer takes NUMBERS and writes `${n}px`. The viewer passes
    // canvas.width || round(widthPx) and canvas.height || round(slideHeightPx).
    // base scale = 200/1000 = 0.2 ⇒ widthPx 200, slideHeightPx 120.
    const expectW = canvas.width || 200;
    const expectH = canvas.height || 120;
    expect(textLayer.style.width).toBe(`${expectW}px`);
    expect(textLayer.style.height).toBe(`${expectH}px`);
    v.destroy();
  });

  it('main mode: recycling a slot clears its overlay so the free pool holds no stale spans', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(50, 1000, 600);
    engine.feedTextRuns = [RUN];
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      enableTextSelection: true,
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    // The overlay is built in the renderSlide .then() microtask.
    await Promise.resolve();
    await Promise.resolve();
    // A slide-0 slot exists with a built overlay (one shape div).
    const wrapper = findSlotWrapper(scrollHost);
    const textLayer = findTextLayer(wrapper);
    expect(textLayer.children.length).toBe(1);
    // Scroll far away so slide 0 recycles into the free pool.
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
    const engine = new FakePptxEngine(3, 1000, 600, 'worker');
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      enableTextSelection: true,
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    await Promise.resolve();
    await Promise.resolve();
    // Warned exactly once with the same wording as PptxViewer, across every
    // mounted slot's render.
    expect(warn).toHaveBeenCalledTimes(1);
    expect(warn.mock.calls[0][0]).toMatch(/text selection is unavailable in mode: 'worker'/);
    // No renderSlide (worker path only) and no overlay children anywhere.
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
    // Deferred main mode: the test resolves renderSlide manually so setScale can
    // move the epoch WHILE slide 0's first render is in flight.
    const engine = new FakePptxEngine(20, 1000, 600, 'main', true);
    engine.feedTextRuns = [RUN];
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      enableTextSelection: true,
      gap: 10,
      zoomMin: 0.05,
      zoomMax: 4,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    // Grab the first (stale) slide-0 render call before resolving it. Capture the
    // slot wrapper it targets so we can inspect its overlay afterwards.
    const staleWrapper = scrollHost.children.find((c) => c.children.some((k) => k.tag === 'canvas')) as FakeEl;
    const staleLayer = staleWrapper.children.find((c) => c.tag === 'div') as FakeEl;
    const staleCall = engine.renderCalls.find((c) => c.slide === 0)!;
    // Zoom mid-flight: bumps the epoch and force-re-mounts every slot (which
    // re-dispatches a fresh slide-0 render at the new scale).
    v.setScale(v.scaleForTest() * 2);
    // Now resolve the STALE (old-epoch) render. Its .then must bail on the epoch
    // guard and NOT build the overlay.
    staleCall.resolve();
    await Promise.resolve();
    await Promise.resolve();
    // The stale render built nothing (epoch guard). The fresh render is still
    // deferred (not resolved), so its overlay isn't built either — the stale
    // slot/layer holds zero children from the superseded render.
    expect(staleLayer.children.length).toBe(0);
    v.destroy();
  });
});

describe('PptxScrollViewer — navigation, resize, empty (T6)', () => {
  // container fit width 200, slide 1000×600 EMU → base scale = 200/1000 = 0.2.
  // slide height px = 600 * 0.2 = 120. gap 10 ⇒ stride 130.
  const BASE = 200 / 1000; // 0.2
  const SLIDE_H = 600 * BASE; // 120
  const GAP = 10;
  const STRIDE = SLIDE_H + GAP; // 130 — the top-to-top distance between slides

  function setup(slideCount = 20, opts = {}) {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakePptxEngine(slideCount, 1000, 600);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: GAP,
      ...opts,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    return { v, scrollHost, container, engine };
  }

  it('scrollToSlide sets scrollTop to the slide offset and clamps out-of-range', () => {
    const { v, scrollHost } = setup();
    // Pick a clientHeight that makes maxTop STRICTLY LESS than the last slide's top
    // offset, so the `Math.min(maxTop, target)` clamp is actually exercised.
    // totalHeight = 20*120 + 19*10 = 2590; offsets[19] = 19*130 = 2470.
    scrollHost.clientHeight = 200; // maxTop = 2590 − 200 = 2390 < offsets[19] 2470
    v.scrollToSlide(3);
    expect(Math.abs(scrollHost.scrollTop - 3 * STRIDE)).toBeLessThan(2);
    v.scrollToSlide(999); // clamps to last slide index (19)
    const totalHeight = 20 * SLIDE_H + 19 * GAP; // 2590
    const maxTop = totalHeight - 200; // 2390
    const lastOffset = 19 * STRIDE; // 2470
    expect(maxTop).toBeLessThan(lastOffset); // precondition: clamp is exercised
    expect(Math.abs(scrollHost.scrollTop - maxTop)).toBeLessThan(2);
    v.scrollToSlide(-5); // clamps to 0
    expect(scrollHost.scrollTop).toBe(0);
    v.destroy();
  });

  it('scrollToSlide: viewport taller than total content ⇒ scrollTop pinned to 0 (negative maxTop guard)', () => {
    // A 2-slide deck is shorter than a very tall viewport, so
    // totalHeight − clientHeight is NEGATIVE; the `Math.max(0, …)` maxTop guard must
    // pin scrollTop to 0 rather than a negative top.
    const { v, scrollHost } = setup(2);
    scrollHost.clientHeight = 5000; // >> totalHeight (2*120 + 10 = 250)
    v.scrollToSlide(1); // last slide; its offset (130) is above maxTop (0)
    expect(scrollHost.scrollTop).toBe(0);
    v.destroy();
  });

  it('onVisibleSlideChange fires only when topIndex changes', () => {
    const changes: number[] = [];
    const { v, scrollHost } = setup(20, { onVisibleSlideChange: (i: number) => changes.push(i) });
    scrollHost.scrollTop = STRIDE; // slide 1 top
    scrollHost.dispatch('scroll');
    scrollHost.scrollTop = STRIDE + 10; // still within slide 1
    scrollHost.dispatch('scroll');
    scrollHost.scrollTop = 3 * STRIDE; // slide 3 top
    scrollHost.dispatch('scroll');
    // Deduped: 0 (initial) → 1 → 3, no repeat for the STRIDE+10 no-op.
    expect(changes).toEqual([0, 1, 3]);
    v.destroy();
  });

  it('scrollToSlide does not re-fire onVisibleSlideChange for the same top slide', () => {
    const changes: number[] = [];
    const { v } = setup(20, { onVisibleSlideChange: (i: number) => changes.push(i) });
    // Initial mount fired 0. scrollToSlide(0) lands on the same top slide — no new fire.
    v.scrollToSlide(0);
    expect(changes).toEqual([0]);
    // Navigate to slide 5 → one fire; navigate there again → no duplicate.
    v.scrollToSlide(5);
    v.scrollToSlide(5);
    expect(changes).toEqual([0, 5]);
    v.destroy();
  });

  it('ResizeObserver re-fit preserves the zoom multiplier', () => {
    // zoomMax headroom: base 0.2, 2× zoom → 0.4; after the width doubles the new
    // base is 0.4 so the preserved scale is 0.8. zoomMin/zoomMax are ABSOLUTE
    // px-per-EMU bounds (design §8.2), so a low zoomMax would legitimately CLAMP
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
    // the initially-mounted window is [0..3] (viewport 400 / stride 130 ≈ 3 slides;
    // +1 overscan). Growing the viewport must mount more purely from the resize —
    // no scroll event, no epoch bump (geometry/scale unchanged so cached canvases
    // stay valid), and no onVisibleSlideChange fire (topIndex stays 0).
    const changes: number[] = [];
    const { v, scrollHost } = setup(20, { onVisibleSlideChange: (i: number) => changes.push(i) });
    const mountedBefore = v.mountedSlideIndicesForTest().slice().sort((a, b) => a - b);
    expect(mountedBefore.length).toBeGreaterThan(0);
    expect(changes).toEqual([0]); // initial mount fired 0
    const epochBefore = v.renderEpochForTest();
    const scaleBefore = v.scaleForTest();

    // Grow HEIGHT only; width (and thus the fit-width base scale) is unchanged.
    scrollHost.clientHeight = 1600;
    v.resizeForTest(); // NO scroll event dispatched

    const mountedAfter = v.mountedSlideIndicesForTest().slice().sort((a, b) => a - b);
    // The revealed window strictly grows (taller viewport intersects more slides).
    expect(mountedAfter.length).toBeGreaterThan(mountedBefore.length);
    expect(Math.max(...mountedAfter)).toBeGreaterThan(Math.max(...mountedBefore));
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
    // must still mount the rows the taller viewport revealed. base = 0.2, no user
    // zoom ⇒ mult 1; after width→400 newBase 0.4 clamps to zoomMax 0.2 (== _scale).
    const changes: number[] = [];
    const { v, scrollHost, container } = setup(20, {
      zoomMax: BASE, // 0.2 — clamps the re-fit back to the current scale
      onVisibleSlideChange: (i: number) => changes.push(i),
    });
    expect(v.scaleForTest()).toBe(BASE);
    const mountedBefore = v.mountedSlideIndicesForTest().slice().sort((a, b) => a - b);
    expect(mountedBefore.length).toBeGreaterThan(0);
    const epochBefore = v.renderEpochForTest();

    // Grow BOTH width and height. The width grow triggers the re-fit branch, but the
    // clamp pins the scale unchanged ⇒ setScale no-ops. Heights are therefore
    // unchanged (scale unchanged), so the taller viewport reveals more of the SAME
    // geometry.
    container.clientWidth = 400;
    (container.children[0] as FakeEl).children[0].clientWidth = 400;
    scrollHost.clientWidth = 400;
    scrollHost.clientHeight = 1600;
    v.resizeForTest();

    const mountedAfter = v.mountedSlideIndicesForTest().slice().sort((a, b) => a - b);
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
    const engine = new FakePptxEngine(5, 1000, 600);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, { presentation: engine.asPres() });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    expect(v.mountedSlideIndicesForTest().length).toBe(0); // deferred
    container.clientWidth = 300;
    scrollHost.clientWidth = 300;
    scrollHost.clientHeight = 400;
    v.resizeForTest();
    expect(v.mountedSlideIndicesForTest().length).toBeGreaterThan(0);
    v.destroy();
  });

  it('empty deck: slideCount 0 ⇒ spacer 0, no slots, scrollToSlide no-op', () => {
    installDom();
    const container = makeContainer(300, 400);
    const engine = new FakePptxEngine(0, 1000, 600);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, { presentation: engine.asPres() });
    v.relayout();
    const spacer = (container.children[0] as FakeEl).children[0].children[0] as FakeEl;
    expect(parseFloat(spacer.style.height || '0')).toBe(0);
    expect(v.mountedSlideIndicesForTest().length).toBe(0);
    v.scrollToSlide(0); // no throw
    v.destroy();
  });

  it('empty deck: resize does not crash and fires no callback', () => {
    installDom();
    const container = makeContainer(0, 0);
    const engine = new FakePptxEngine(0, 1000, 600);
    const changes: number[] = [];
    const v = new PptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      onVisibleSlideChange: (i: number) => changes.push(i),
    });
    container.clientWidth = 300;
    (container.children[0] as FakeEl).children[0].clientWidth = 300;
    v.resizeForTest(); // no slides ⇒ no-op
    expect(v.mountedSlideIndicesForTest().length).toBe(0);
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
    const engine = new FakePptxEngine(3, 1000, 600);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, { presentation: engine.asPres() });
    v.destroy();
    expect(disconnected).toBe(1);
  });
});
