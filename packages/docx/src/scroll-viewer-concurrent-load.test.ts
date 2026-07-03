import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxScrollViewer } from './scroll-viewer.js';
import { DocxDocument } from './document.js';
import { installDom, makeContainer, FakeDocxEngine, type FakeEl } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const SIZE = [{ widthPt: 100, heightPt: 200 }];

/**
 * Concurrent-load latch for the SELF-LOADED scroll viewer (composes with SC20's
 * success-after-swap). Two overlapping `load(A)`/`load(B)` calls race the WASM
 * parse / worker init; the stale one resolving LAST must NOT win the swap — it
 * destroys its own just-loaded engine and leaves the winner (its engine +
 * recycle/relayout post-load work) untouched. An injected engine can never orphan
 * (load() throws up-front there), so this only covers the self-loading path.
 */
describe('DocxScrollViewer.load() — concurrent-load latch', () => {
  function build() {
    installDom();
    const container = makeContainer(200, 400);
    const v = new DocxScrollViewer(container as unknown as HTMLElement, { gap: 10 });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    return { v };
  }

  function deferredLoad(engine: FakeDocxEngine): { resolve: () => void; promise: Promise<DocxDocument> } {
    let resolve!: () => void;
    const promise = new Promise<DocxDocument>((r) => {
      resolve = () => r(engine.asDoc());
    });
    return { resolve, promise };
  }

  it('the later-started load winning first leaves the stale load a no-op (its engine destroyed)', async () => {
    const { v } = build();
    const a = new FakeDocxEngine(4, SIZE);
    const b = new FakeDocxEngine(4, SIZE);
    const da = deferredLoad(a);
    const db = deferredLoad(b);
    vi.spyOn(DocxDocument, 'load')
      .mockImplementationOnce(() => da.promise)
      .mockImplementationOnce(() => db.promise);

    const pa = v.load('a.docx'); // gen 1
    const pb = v.load('b.docx'); // gen 2 — supersedes A

    db.resolve();
    await pb;
    expect(b.destroyed).toBe(false);
    expect(a.destroyed).toBe(false);

    da.resolve();
    await pa;
    expect(a.destroyed).toBe(true); // loser's engine cleaned up (no leak)
    expect(b.destroyed).toBe(false); // winner untouched — still current
    expect(v.pageCount).toBe(4);

    v.destroy();
    expect(b.destroyed).toBe(true);
    expect(a.destroyed).toBe(true);
  });

  it('resolving in start order (A then B) behaves like today — B wins normally', async () => {
    const { v } = build();
    const a = new FakeDocxEngine(4, SIZE);
    const b = new FakeDocxEngine(4, SIZE);
    const da = deferredLoad(a);
    const db = deferredLoad(b);
    vi.spyOn(DocxDocument, 'load')
      .mockImplementationOnce(() => da.promise)
      .mockImplementationOnce(() => db.promise);

    const pa = v.load('a.docx');
    const pb = v.load('b.docx');

    da.resolve();
    await pa;
    expect(a.destroyed).toBe(true); // superseded loser cleaned up

    db.resolve();
    await pb;
    expect(b.destroyed).toBe(false);
    expect(v.pageCount).toBe(4);

    v.destroy();
    expect(b.destroyed).toBe(true);
  });
});
