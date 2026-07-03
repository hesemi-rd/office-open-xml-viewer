import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxViewer } from './viewer.js';
import { DocxDocument } from './document.js';
import { installDom, makeEl, FakeDocxEngine } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const A4 = [{ widthPt: 595, heightPt: 842 }];

/**
 * Concurrent-load latch (composes with SC20's success-after-swap): if a caller
 * fires `load(A)` and, before it resolves, `load(B)`, both loads race the WASM
 * parse / worker init concurrently. Whichever resolves LAST must NOT win the swap
 * when it is the stale one — the loser's freshly-loaded engine (never installed,
 * or installed then overwritten) must be destroyed, not leaked, and the winner's
 * engine must stay live and untouched. A generation token (`_loadGen`) closes it.
 */
describe('DocxViewer.load() — concurrent-load latch', () => {
  function mount() {
    installDom();
    return { canvas: makeEl('canvas') };
  }

  /** A load whose resolution the test controls, so two loads can overlap and
   *  resolve in a chosen order. */
  function deferredLoad(engine: FakeDocxEngine): { resolve: () => void; promise: Promise<DocxDocument> } {
    let resolve!: () => void;
    const promise = new Promise<DocxDocument>((r) => {
      resolve = () => r(engine.asDoc());
    });
    return { resolve, promise };
  }

  it('the later-started load winning first leaves the stale load a no-op (its engine destroyed)', async () => {
    const { canvas } = mount();
    const a = new FakeDocxEngine(2, A4);
    const b = new FakeDocxEngine(2, A4);
    const da = deferredLoad(a);
    const db = deferredLoad(b);
    vi.spyOn(DocxDocument, 'load')
      .mockImplementationOnce(() => da.promise)
      .mockImplementationOnce(() => db.promise);

    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement);
    const pa = v.load('a.docx'); // gen 1
    const pb = v.load('b.docx'); // gen 2 — supersedes A

    // B (the later-started load) resolves FIRST and installs normally via SC20.
    db.resolve();
    await pb;
    expect(b.destroyed).toBe(false); // winner is live
    expect(a.destroyed).toBe(false); // A's engine not even loaded yet

    // A resolves LATE. It lost the race (gen 1 ≠ live gen 2): it must destroy its
    // OWN engine and NOT touch the installed winner B.
    da.resolve();
    await pa;
    expect(a.destroyed).toBe(true); // loser's engine cleaned up (no leak)
    expect(b.destroyed).toBe(false); // winner untouched — still current

    v.destroy();
    expect(b.destroyed).toBe(true); // only B is torn down by destroy()
    expect(a.destroyed).toBe(true); // and it was closed exactly once (still true)
  });

  it('resolving in start order (A then B) behaves like today — B wins normally', async () => {
    const { canvas } = mount();
    const a = new FakeDocxEngine(2, A4);
    const b = new FakeDocxEngine(2, A4);
    const da = deferredLoad(a);
    const db = deferredLoad(b);
    vi.spyOn(DocxDocument, 'load')
      .mockImplementationOnce(() => da.promise)
      .mockImplementationOnce(() => db.promise);

    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement);
    const pa = v.load('a.docx'); // gen 1
    const pb = v.load('b.docx'); // gen 2

    // A resolves first but is already superseded: it installs nothing and destroys
    // its own engine. B resolves next and wins the swap (SC20 unaffected).
    da.resolve();
    await pa;
    expect(a.destroyed).toBe(true); // superseded loser cleaned up

    db.resolve();
    await pb;
    expect(b.destroyed).toBe(false); // winner is current
    expect(v.pageCount).toBe(2);

    v.destroy();
    expect(b.destroyed).toBe(true);
  });

  it('does not double-destroy or leak when only one load runs (regression guard)', async () => {
    const { canvas } = mount();
    const only = new FakeDocxEngine(2, A4);
    vi.spyOn(DocxDocument, 'load').mockResolvedValue(only.asDoc());
    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement);
    await v.load('one.docx');
    expect(only.destroyed).toBe(false); // installed, never superseded
    v.destroy();
    expect(only.destroyed).toBe(true); // destroyed exactly once
  });
});
