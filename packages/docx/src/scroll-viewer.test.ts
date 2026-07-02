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
});
