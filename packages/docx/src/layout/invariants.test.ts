import { describe, expect, it } from 'vitest';
import { layoutFlowBlocks } from './flow.js';
import { assertDocumentLayout, layoutFingerprint } from './invariants.js';
import type {
  BlockLayoutAlgorithms,
  DocumentLayout,
  DrawingLayout,
  LayoutRect,
  LayoutServices,
  PaintNode,
  SourceRef,
} from './types.js';
import type { SectionLayoutContext } from '../layout-context.js';

const source = (index: number): SourceRef => ({
  story: 'body',
  storyInstance: 'body',
  path: [index],
});

const rect = (xPt: number, yPt: number, widthPt: number, heightPt: number): LayoutRect => ({
  xPt,
  yPt,
  widthPt,
  heightPt,
});

function drawing(
  id: string,
  flowBounds: LayoutRect,
  options: Partial<Pick<DrawingLayout, 'inkBounds' | 'clipBounds' | 'ordinaryFlow'>> = {},
): DrawingLayout {
  return {
    kind: 'drawing',
    id,
    source: source(Number(id.replace(/\D/g, '')) || 0),
    flowBounds,
    inkBounds: options.inkBounds ?? flowBounds,
    clipBounds: options.clipBounds,
    advancePt: flowBounds.heightPt,
    ordinaryFlow: options.ordinaryFlow ?? true,
    commands: [],
  };
}

function documentWith(
  nodes: readonly PaintNode[],
  diagnostics: DocumentLayout['diagnostics'] = [],
): DocumentLayout {
  return {
    pages: [{
      pageIndex: 0,
      geometry: {
        ...rect(0, 0, 612, 792),
        contentTopPt: 72,
        contentBottomPt: 720,
      },
      section: {} as SectionLayoutContext,
      layers: {
        paintOrder: nodes.map((node) => ({ layer: 'body' as const, nodeId: node.id })),
        background: [],
        behindText: [],
        header: [],
        body: nodes,
        notes: [],
        front: [],
        footer: [],
      },
      readingOrder: nodes.map((node) => node.id),
    }],
    diagnostics,
  };
}

describe('assertDocumentLayout', () => {
  it('rejects overlapping ordinary flow allocations', () => {
    const layout = documentWith([
      drawing('n1', rect(72, 100, 200, 30)),
      drawing('n2', rect(72, 120, 200, 30)),
    ]);

    expect(() => assertDocumentLayout(layout)).toThrow(/FLOW_OVERLAP/);
  });

  it('rejects ordinary flow that enters the bottom margin', () => {
    const layout = documentWith([drawing('n1', rect(72, 710, 200, 20))]);

    expect(() => assertDocumentLayout(layout)).toThrow(/BOTTOM_MARGIN_INVASION/);
  });

  it('allows floating overlap, negative-spacing ink, and clipped overhang', () => {
    const ordinary = drawing('n1', rect(72, 100, 200, 30), {
      inkBounds: rect(72, 92, 200, 38),
    });
    const floating = drawing('n2', rect(72, 110, 200, 30), {
      ordinaryFlow: false,
    });
    const clipped = drawing('n3', rect(72, 200, 200, 30), {
      inkBounds: rect(60, 190, 240, 600),
      clipBounds: rect(72, 200, 200, 30),
    });

    expect(() => assertDocumentLayout(documentWith([ordinary, floating, clipped]))).not.toThrow();
  });

  it('rejects non-finite retained geometry', () => {
    const layout = documentWith([drawing('n1', rect(Number.NaN, 100, 200, 30))]);

    expect(() => assertDocumentLayout(layout)).toThrow(/INVALID_GEOMETRY/);
  });
});

describe('layoutFingerprint', () => {
  it('normalizes geometry and excludes diagnostic prose while retaining diagnostic identity', () => {
    const first = documentWith(
      [drawing('n1', rect(72.0000001, 100, 200, 30))],
      [{ code: 'UNSUPPORTED_FEATURE', severity: 'warning', message: 'first prose' }],
    );
    const second = documentWith(
      [drawing('n1', rect(72.0000002, 100, 200, 30))],
      [{ code: 'UNSUPPORTED_FEATURE', severity: 'warning', message: 'different prose' }],
    );
    const changedCode = documentWith(
      [drawing('n1', rect(72.0000002, 100, 200, 30))],
      [{ code: 'NON_CONVERGENCE', severity: 'warning', message: 'different prose' }],
    );

    expect(layoutFingerprint(first)).toBe(layoutFingerprint(second));
    expect(layoutFingerprint(first)).not.toBe(layoutFingerprint(changedCode));
  });
});

describe('layoutFlowBlocks', () => {
  it('dispatches paragraph and table blocks through one injected coordinator', () => {
    const calls: string[] = [];
    const algorithms: BlockLayoutAlgorithms = {
      layoutParagraph(input) {
        calls.push(`paragraph:${input.source.path.join('.')}`);
        return { ...drawing('p1', rect(10, 20, 100, 12)), kind: 'paragraph' };
      },
      layoutTable(input) {
        calls.push(`table:${input.source.path.join('.')}`);
        return { ...drawing('t2', rect(10, 32, 100, 18)), kind: 'table' };
      },
    };
    const services: LayoutServices = {
      text: { fingerprint: 'text' },
      images: { fingerprint: 'images' },
      math: { fingerprint: 'math' },
    };

    const result = layoutFlowBlocks({
      source: source(0),
      container: { bounds: rect(10, 20, 100, 200) },
      blocks: [
        { kind: 'paragraph', source: source(1) },
        { kind: 'table', source: source(2) },
      ],
    }, services, algorithms);

    expect(calls).toEqual(['paragraph:1', 'table:2']);
    expect(result.blocks.map((block) => block.id)).toEqual(['p1', 't2']);
    expect(result.advancePt).toBe(30);
    expect(result.flowBounds).toEqual(rect(10, 20, 100, 30));
  });
});
