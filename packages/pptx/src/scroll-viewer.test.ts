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
