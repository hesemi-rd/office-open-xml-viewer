import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxViewer } from './viewer.js';
import { DocxScrollViewer } from './scroll-viewer.js';
import { DocxDocument } from './document.js';
import {
  installDom,
  makeEl,
  makeContainer,
  FakeDocxEngine,
  type FakeEl,
} from './scroll-viewer-test-dom.js';
import type { DocxTextRunInfo } from './renderer';
import type { HyperlinkTarget } from '@silurus/ooxml-core';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * IX1 `enableHyperlinks` — a viewer-level policy switch (default `true`). When
 * `false`, the hyperlink machinery is not wired AT ALL: the overlay's clickable
 * span is never installed, so there is no pointer cursor, no title tooltip, no
 * click listener, no default navigation, and `onHyperlinkClick` is never called.
 *
 * The single gate is `_hyperlinkHandler()`, which returns `undefined` when the
 * option is `false`. `buildDocxTextLayer` already treats a missing handler as
 * "render link runs exactly like plain runs", so passing `undefined` disables the
 * whole feature at its source rather than sprinkling `if` checks through the
 * overlay. These tests pin the gate on the single-canvas viewer and the resulting
 * inert overlay end-to-end on the scroll viewer.
 */

const PAGE = [{ widthPt: 595, heightPt: 842 }];

async function mountViewer(opts: Record<string, unknown> = {}) {
  installDom();
  const canvas = makeEl('canvas');
  const engine = new FakeDocxEngine(3, PAGE);
  vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
  const v = new DocxViewer(canvas as unknown as HTMLCanvasElement, opts);
  await v.load('x.docx');
  const handler = (
    v as unknown as {
      _hyperlinkHandler: () => ((t: HyperlinkTarget) => void) | undefined;
    }
  )._hyperlinkHandler();
  return { v, engine, handler };
}

describe('DocxViewer — enableHyperlinks option', () => {
  it('installs no hyperlink handler when enableHyperlinks is false', async () => {
    const { v, handler } = await mountViewer({ enableHyperlinks: false });
    expect(handler).toBeUndefined();
    v.destroy();
  });

  it('suppresses even a supplied onHyperlinkClick when enableHyperlinks is false', async () => {
    const onHyperlinkClick = vi.fn<(t: HyperlinkTarget) => void>();
    const { v, handler } = await mountViewer({ enableHyperlinks: false, onHyperlinkClick });
    expect(handler).toBeUndefined();
    expect(onHyperlinkClick).not.toHaveBeenCalled();
    v.destroy();
  });

  it('wires the default handler when the option is omitted (default true)', async () => {
    const { v, handler } = await mountViewer();
    expect(typeof handler).toBe('function');
    // The default external branch still opens a new tab (unchanged behaviour).
    const winOpen = vi.fn();
    (globalThis as unknown as { window: { open: unknown } }).window.open = winOpen;
    handler?.({ kind: 'external', url: 'https://example.com/' });
    expect(winOpen).toHaveBeenCalledWith('https://example.com/', '_blank', 'noopener,noreferrer');
    v.destroy();
  });

  it('wires the handler when enableHyperlinks is explicitly true', async () => {
    const { v, handler } = await mountViewer({ enableHyperlinks: true });
    expect(typeof handler).toBe('function');
    v.destroy();
  });
});

// ---------------------------------------------------------------------------
// DocxScrollViewer — end-to-end overlay behaviour.
// ---------------------------------------------------------------------------

function linkRun(hyperlink: HyperlinkTarget): DocxTextRunInfo {
  return { text: 'go', x: 1, y: 2, w: 10, h: 12, fontSize: 12, font: '12px serif', hyperlink };
}

async function setupScroll(hyperlink: HyperlinkTarget, opts: Record<string, unknown> = {}) {
  installDom();
  const container = makeContainer(200, 400);
  const engine = new FakeDocxEngine(3, [
    { widthPt: 100, heightPt: 200 },
    { widthPt: 100, heightPt: 200 },
    { widthPt: 100, heightPt: 200 },
  ]);
  engine.feedTextRuns = [linkRun(hyperlink)];
  const v = new DocxScrollViewer(container as unknown as HTMLElement, {
    document: engine.asDoc(),
    gap: 10,
    overscan: 1,
    paddingLeft: 0,
    paddingRight: 0,
    enableTextSelection: true,
    ...opts,
  });
  const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
  scrollHost.clientHeight = 400;
  scrollHost.clientWidth = 200;
  v.relayout();
  await Promise.resolve();
  await Promise.resolve();

  /** The first overlay span for the fed run, in any mounted slot (or null). */
  function linkSpan(): FakeEl | null {
    for (const slot of scrollHost.children) {
      const layer = slot.children.find((k) => k.tag === 'div') as FakeEl | undefined;
      const span = layer?.children[0] as FakeEl | undefined;
      if (span) return span;
    }
    return null;
  }

  return { v, engine, scrollHost, linkSpan };
}

describe('DocxScrollViewer — enableHyperlinks option', () => {
  it('renders the link run as an inert (text-cursor, no click) span when false', async () => {
    const { v, linkSpan } = await setupScroll(
      { kind: 'external', url: 'https://example.com/' },
      { enableHyperlinks: false },
    );
    const span = linkSpan();
    expect(span).not.toBeNull();
    // No click affordance: text cursor, and no click listener was installed.
    expect(span?.style.cursor).toBe('text');
    expect(span?._listeners.has('click')).toBe(false);
    v.destroy();
  });

  it('does not fire a supplied onHyperlinkClick when false (no listener at all)', async () => {
    const onHyperlinkClick = vi.fn<(t: HyperlinkTarget) => void>();
    const { v, linkSpan } = await setupScroll(
      { kind: 'internal', ref: 'Heading2' },
      { enableHyperlinks: false, onHyperlinkClick },
    );
    const span = linkSpan();
    expect(span?._listeners.has('click')).toBe(false);
    // Even if something dispatched a click, there is no listener to invoke it.
    (span as unknown as { dispatch: (t: string, e?: unknown) => void } | null)?.dispatch('click', {
      preventDefault() {},
    });
    expect(onHyperlinkClick).not.toHaveBeenCalled();
    v.destroy();
  });

  it('renders a clickable (pointer, click listener) span by default', async () => {
    const { v, linkSpan } = await setupScroll({ kind: 'external', url: 'https://example.com/' });
    const span = linkSpan();
    expect(span?.style.cursor).toBe('pointer');
    expect(span?._listeners.has('click')).toBe(true);
    v.destroy();
  });
});
