import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxScrollViewer } from './scroll-viewer.js';
import { installDom, makeContainer, FakeDocxEngine, type FakeEl } from './scroll-viewer-test-dom.js';

afterEach(() => vi.unstubAllGlobals());

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
