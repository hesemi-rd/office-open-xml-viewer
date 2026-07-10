import { describe, it, expect, afterEach, vi } from 'vitest';
import { PptxViewer } from './viewer.js';
import { PptxScrollViewer } from './scroll-viewer.js';
import { PptxPresentation } from './presentation.js';
import {
  installDom,
  makeEl,
  makeContainer,
  FakePptxEngine,
  type FakeEl,
} from './scroll-viewer-test-dom.js';
import type { PptxTextRunInfo } from './renderer';
import type { HyperlinkTarget } from '@silurus/ooxml-core';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * IX1 `enableHyperlinks` — a viewer-level policy switch (default `true`). When
 * `false`, the hyperlink machinery is not wired AT ALL: the overlay's clickable
 * span is never installed (no pointer cursor, no title, no click listener, no
 * default navigation) and `onHyperlinkClick` is never called.
 *
 * The single gate is `_hyperlinkHandler()`, which returns `undefined` when the
 * option is `false`; `buildPptxTextLayer` already renders link runs like plain
 * runs when no handler is supplied. These tests pin the gate on the single-canvas
 * viewer and the inert overlay end-to-end on the scroll viewer.
 */

const SLIDE_W_EMU = 9525 * 200;
const SLIDE_H_EMU = 9525 * 120;

async function mountViewer(opts: Record<string, unknown> = {}) {
  installDom();
  const canvas = makeEl('canvas');
  const engine = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
  vi.spyOn(PptxPresentation, 'load').mockResolvedValue(engine.asPres());
  const v = new PptxViewer(canvas as unknown as HTMLCanvasElement, opts);
  await v.load('x.pptx');
  const handler = (
    v as unknown as {
      _hyperlinkHandler: () => ((t: HyperlinkTarget) => void) | undefined;
    }
  )._hyperlinkHandler();
  return { v, engine, handler };
}

describe('PptxViewer — enableHyperlinks option', () => {
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
// PptxScrollViewer — end-to-end overlay behaviour.
// ---------------------------------------------------------------------------

function linkRun(hyperlink: HyperlinkTarget): PptxTextRunInfo {
  return {
    text: 'go',
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
    hyperlink,
  };
}

async function setupScroll(hyperlink: HyperlinkTarget, opts: Record<string, unknown> = {}) {
  installDom();
  const container = makeContainer(200, 400);
  const engine = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
  engine.feedTextRuns = [linkRun(hyperlink)];
  const v = new PptxScrollViewer(container as unknown as HTMLElement, {
    presentation: engine.asPres(),
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

  /** The first overlay span for the fed run, in any mounted slot (or null). The
   *  pptx overlay nests spans one shape-div deep inside the text-layer div. */
  function linkSpan(): FakeEl | null {
    for (const slot of scrollHost.children) {
      const layer = slot.children.find((k) => k.tag === 'div') as FakeEl | undefined;
      const shapeDiv = layer?.children[0] as FakeEl | undefined;
      const span = shapeDiv?.children[0] as FakeEl | undefined;
      if (span) return span;
    }
    return null;
  }

  return { v, engine, scrollHost, linkSpan };
}

describe('PptxScrollViewer — enableHyperlinks option', () => {
  it('renders the link run as an inert (text-cursor, no click) span when false', async () => {
    const { v, linkSpan } = await setupScroll(
      { kind: 'external', url: 'https://example.com/' },
      { enableHyperlinks: false },
    );
    const span = linkSpan();
    expect(span).not.toBeNull();
    expect(span?.style.cursor).toBe('text');
    expect(span?._listeners.has('click')).toBe(false);
    v.destroy();
  });

  it('does not fire a supplied onHyperlinkClick when false (no listener at all)', async () => {
    const onHyperlinkClick = vi.fn<(t: HyperlinkTarget) => void>();
    const { v, linkSpan } = await setupScroll(
      { kind: 'internal', ref: '../slides/slide3.xml' },
      { enableHyperlinks: false, onHyperlinkClick },
    );
    const span = linkSpan();
    expect(span?._listeners.has('click')).toBe(false);
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
