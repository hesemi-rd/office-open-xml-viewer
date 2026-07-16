import { describe, expect, it } from 'vitest';
import { layoutFlowBlocks } from './flow.js';
import { assertDocumentLayout, layoutFingerprint } from './invariants.js';
import { LayoutInvariantError } from './diagnostics.js';
import type {
  BlockLayoutAlgorithms,
  DocumentLayout,
  DrawingLayout,
  FlowDomain,
  LayoutRect,
  LayoutServices,
  PaintNode,
  SourceRef,
  TableEdgeInputs,
  TableLayoutInput,
} from './types.js';
import type { SectionLayoutContext } from '../layout-context.js';
import { createCanvasFontRoute } from '@silurus/ooxml-core';

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

const noTableBorders: TableEdgeInputs = {
  top: null, right: null, bottom: null, left: null, insideH: null, insideV: null,
};

function tableInput(index: number): TableLayoutInput {
  return {
    kind: 'table', id: `table-input-${index}`, source: source(index),
    flowDomainId: 'body', ordinaryFlow: true,
    alignment: 'left', indentPt: 0, bidiVisual: false,
    columnWidthsPt: [], borders: noTableBorders, rows: [],
  };
}

function drawing(
  id: string,
  flowBounds: LayoutRect,
  options: Partial<Pick<DrawingLayout, 'inkBounds' | 'clipBounds' | 'ordinaryFlow' | 'flowDomainId'>> = {},
): DrawingLayout {
  return {
    kind: 'drawing',
    id,
    source: source(Number(id.replace(/\D/g, '')) || 0),
    flowBounds,
    inkBounds: options.inkBounds ?? flowBounds,
    ...(options.clipBounds ? { clipBounds: options.clipBounds } : {}),
    advancePt: flowBounds.heightPt,
    ordinaryFlow: options.ordinaryFlow ?? true,
    flowDomainId: options.flowDomainId ?? 'body',
    commands: [],
  };
}

const bodyDomain: FlowDomain = {
  id: 'body',
  kind: 'body',
  logicalBounds: rect(72, 72, 468, 648),
  physicalBounds: rect(72, 72, 468, 648),
};

const horizontalSection: SectionLayoutContext = {
  geometry: {
    pageWidth: 612, pageHeight: 792,
    marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
    headerDistance: 36, footerDistance: 36,
  },
  columns: [{ xPt: 72, wPt: 468 }],
  grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
  textDirection: 'lrTb',
  verticalAlignment: 'top',
};

const verticalSection: SectionLayoutContext = {
  ...horizontalSection,
  geometry: { ...horizontalSection.geometry, pageWidth: 792, pageHeight: 612 },
  columns: [{ xPt: 72, wPt: 648 }],
  textDirection: 'tbRl',
};

function regionLayout(): DocumentLayout {
  const base = documentWith([]);
  return {
    ...base,
    pages: [{
      ...base.pages[0]!,
      section: horizontalSection,
      sectionOccurrenceId: 'section:body',
      flowDomains: [{
        id: 'body', kind: 'body',
        logicalBounds: rect(72, 72, 468, 648),
        physicalBounds: rect(72, 72, 468, 648),
      }],
      sectionRegions: [{
        id: 'region:body', sectionOccurrenceId: 'section:body',
        coordinateSpace: {
          writingMode: 'horizontal-tb',
          logicalToPhysical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          physicalToLogical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
        },
        blockStartPt: 72, blockEndPt: 720,
        flowDomainIds: ['body'], section: horizontalSection,
      }],
    }],
  };
}

function twoRegionLayout(
  secondSectionOccurrenceId: string,
  secondWritingMode: 'horizontal-tb' | 'vertical-rl' = 'horizontal-tb',
): DocumentLayout {
  const base = regionLayout();
  const firstDomain = {
    ...base.pages[0]!.flowDomains[0]!,
    logicalBounds: rect(72, 72, 468, 228),
    physicalBounds: rect(72, 72, 468, 228),
  };
  const secondIsVertical = secondWritingMode === 'vertical-rl';
  const secondSection = secondIsVertical ? verticalSection : horizontalSection;
  const secondDomain: FlowDomain = secondIsVertical
    ? {
        id: 'body:second', kind: 'body',
        logicalBounds: rect(72, 300, 648, 240),
        physicalBounds: rect(72, 72, 240, 648),
      }
    : {
        id: 'body:second', kind: 'body',
        logicalBounds: rect(72, 300, 468, 420),
        physicalBounds: rect(72, 300, 468, 420),
      };
  return {
    ...base,
    pages: [{
      ...base.pages[0]!,
      flowDomains: [firstDomain, secondDomain],
      sectionRegions: [
        {
          ...base.pages[0]!.sectionRegions![0]!,
          blockEndPt: 300,
          flowDomainIds: [firstDomain.id],
        },
        {
          id: 'region:second',
          sectionOccurrenceId: secondSectionOccurrenceId,
          coordinateSpace: secondIsVertical
            ? {
                writingMode: 'vertical-rl',
                logicalToPhysical: { a: 0, b: 1, c: -1, d: 0, e: 612, f: 0 },
                physicalToLogical: { a: 0, b: -1, c: 1, d: 0, e: 0, f: 612 },
              }
            : {
                writingMode: 'horizontal-tb',
                logicalToPhysical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
                physicalToLogical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
              },
          blockStartPt: 300,
          blockEndPt: secondIsVertical ? 540 : 720,
          flowDomainIds: [secondDomain.id],
          section: secondSection,
        },
      ],
    }],
  };
}

function expectInvariantCode(
  code: LayoutInvariantError['code'],
  run: () => void,
): void {
  try {
    run();
    throw new Error(`expected ${code}`);
  } catch (error) {
    expect(error).toBeInstanceOf(LayoutInvariantError);
    expect((error as LayoutInvariantError).code).toBe(code);
  }
}

function expectInvalidGeometry(run: () => void): void {
  expectInvariantCode('INVALID_GEOMETRY', run);
}

function serviceStubs(): LayoutServices {
  return {
    text: {
      fingerprint: 'text',
      localMetrics: {},
      resolve: () => ({
        requestedFamily: 'sans-serif', resolvedFamily: 'sans-serif',
        route: createCanvasFontRoute('sans-serif', 'generic'),
        source: 'generic', weight: 400, style: 'normal', diagnostics: [], genericFamily: 'sans-serif',
      }),
      shape: () => ({ advancePt: 0, ascentPt: 0, descentPt: 0, spans: [], graphemeBoundaries: [0], diagnostics: [] }),
    },
    images: {
      fingerprint: 'images',
      resolve: () => ({ widthPt: 0, heightPt: 0, mimeType: 'application/octet-stream' }),
    },
    math: {
      fingerprint: 'math',
      resolve: () => ({ resourceKey: 'math', widthEm: 0, ascentEm: 0, descentEm: 0, diagnostics: [] }),
    },
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
      flowDomains: [bodyDomain],
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

  it('validates ordinary flow only against siblings in the same story container', () => {
    const body = drawing('body-1', rect(72, 690, 200, 20));
    const footer = {
      ...drawing('footer-1', rect(72, 738, 200, 20), { flowDomainId: 'footer:default' }),
      source: { story: 'footer' as const, storyInstance: 'default', path: [0] },
    };
    const base = documentWith([]);
    const layout: DocumentLayout = {
      ...base,
      pages: [{
        ...base.pages[0]!,
        flowDomains: [
          bodyDomain,
          {
            id: 'footer:default', kind: 'footer',
            logicalBounds: rect(72, 730, 468, 40), physicalBounds: rect(72, 730, 468, 40),
          },
        ],
        layers: {
          ...base.pages[0]!.layers,
          paintOrder: [
            { layer: 'body', nodeId: body.id },
            { layer: 'footer', nodeId: footer.id },
          ],
          body: [body],
          footer: [footer],
        },
        readingOrder: [body.id, footer.id],
      }],
    };

    expect(() => assertDocumentLayout(layout)).not.toThrow();
  });

  it('rejects overlap within one domain but permits the same geometry in independent cells', () => {
    const first = drawing('cell-1', rect(100, 100, 80, 20), { flowDomainId: 'cell:1' });
    const second = drawing('cell-2', rect(100, 100, 80, 20), { flowDomainId: 'cell:2' });
    const base = documentWith([first, second]);
    const layout: DocumentLayout = {
      ...base,
      pages: [{
        ...base.pages[0]!,
        flowDomains: [
          {
            id: 'cell:1', kind: 'tableCell',
            logicalBounds: rect(90, 90, 100, 40), physicalBounds: rect(90, 90, 100, 40),
          },
          {
            id: 'cell:2', kind: 'tableCell',
            logicalBounds: rect(90, 90, 100, 40), physicalBounds: rect(90, 90, 100, 40),
          },
        ],
      }],
    };

    expect(() => assertDocumentLayout(layout)).not.toThrow();
  });

  it('rejects missing domains and invalid paint or reading-order references', () => {
    const node = drawing('n1', rect(72, 100, 200, 30), { flowDomainId: 'missing' });
    expect(() => assertDocumentLayout(documentWith([node]))).toThrow(/INVALID_REFERENCE/);

    const base = documentWith([drawing('n1', rect(72, 100, 200, 30))]);
    const badPaint: DocumentLayout = {
      ...base,
      pages: [{
        ...base.pages[0]!,
        layers: { ...base.pages[0]!.layers, paintOrder: [{ layer: 'body', nodeId: 'unknown' }] },
      }],
    };
    expect(() => assertDocumentLayout(badPaint)).toThrow(/INVALID_REFERENCE/);

    const badReading: DocumentLayout = {
      ...base,
      pages: [{ ...base.pages[0]!, readingOrder: ['unknown'] }],
    };
    expect(() => assertDocumentLayout(badReading)).toThrow(/INVALID_REFERENCE/);
  });

  it('rejects duplicate node IDs and duplicate paint entries', () => {
    expect(() => assertDocumentLayout(documentWith([
      drawing('n1', rect(72, 100, 200, 20)),
      drawing('n1', rect(72, 130, 200, 20)),
    ]))).toThrow(/INVALID_REFERENCE/);

    const base = documentWith([drawing('n1', rect(72, 100, 200, 20))]);
    const duplicatePaint: DocumentLayout = {
      ...base,
      pages: [{
        ...base.pages[0]!,
        layers: {
          ...base.pages[0]!.layers,
          paintOrder: [
            { layer: 'body', nodeId: 'n1' },
            { layer: 'body', nodeId: 'n1' },
          ],
        },
      }],
    };
    expect(() => assertDocumentLayout(duplicatePaint)).toThrow(/INVALID_REFERENCE/);
  });

  it('rejects non-finite retained geometry', () => {
    const layout = documentWith([drawing('n1', rect(Number.NaN, 100, 200, 30))]);

    expect(() => assertDocumentLayout(layout)).toThrow(/INVALID_GEOMETRY/);
  });

  it('requires non-negative sequential page identity', () => {
    const base = documentWith([]);
    const negative: DocumentLayout = {
      ...base,
      pages: [{ ...base.pages[0]!, pageIndex: -1 }],
    };
    const skipped: DocumentLayout = {
      ...base,
      pages: [{ ...base.pages[0]!, pageIndex: 2 }],
    };
    const duplicate: DocumentLayout = {
      ...base,
      pages: [base.pages[0]!, { ...base.pages[0]! }],
    };

    expect(() => assertDocumentLayout(negative)).toThrow(/page index/);
    expect(() => assertDocumentLayout(skipped)).toThrow(/page index/);
    expect(() => assertDocumentLayout(duplicate)).toThrow(/page index/);
  });

  it('requires ordered effective page edges within the physical page and permits equality', () => {
    const base = documentWith([]);
    const withEdges = (contentTopPt: number, contentBottomPt: number): DocumentLayout => ({
      ...base,
      pages: [{
        ...base.pages[0]!,
        geometry: { ...base.pages[0]!.geometry, contentTopPt, contentBottomPt },
      }],
    });

    expect(() => assertDocumentLayout(withEdges(-1, 720))).toThrow(/effective page edges/);
    expect(() => assertDocumentLayout(withEdges(72, 793))).toThrow(/effective page edges/);
    expect(() => assertDocumentLayout(withEdges(721, 720))).toThrow(/effective page edges/);
    expect(() => assertDocumentLayout(withEdges(0, 0))).not.toThrow();
    expect(() => assertDocumentLayout(withEdges(792, 792))).not.toThrow();
  });

  it('requires every body flow domain to belong to exactly one page-local section region', () => {
    const base = documentWith([drawing('n1', rect(72, 100, 200, 30))]);
    const noColumnSection = { ...horizontalSection, columns: [] };
    const layout = {
      ...base,
      pages: [{
        ...base.pages[0]!,
        section: noColumnSection,
        sectionRegions: [{
          id: 'section-region:0',
          sectionOccurrenceId: 'section:0',
          coordinateSpace: {
            writingMode: 'horizontal-tb',
            logicalToPhysical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
            physicalToLogical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          },
          blockStartPt: 72,
          blockEndPt: 720,
          flowDomainIds: [],
          section: noColumnSection,
        }],
      }],
    } as DocumentLayout;

    expect(() => assertDocumentLayout(layout)).toThrow(/section region ownership/);
  });

  it('accepts logical vertical flow even when its upright physical rectangle has exchanged axes', () => {
    const node = drawing('vertical', rect(72, 100, 648, 40), { flowDomainId: 'vertical-body' });
    const base = documentWith([node]);
    const strictVerticalSection = { ...verticalSection, textDirection: 'rlV' };
    const layout: DocumentLayout = {
      ...base,
      pages: [{
        ...base.pages[0]!,
        section: strictVerticalSection,
        sectionOccurrenceId: 'section:vertical',
        flowDomains: [{
          id: 'vertical-body', kind: 'body',
          logicalBounds: rect(72, 72, 648, 468),
          physicalBounds: rect(72, 72, 468, 648),
        }],
        sectionRegions: [{
          id: 'region:vertical', sectionOccurrenceId: 'section:vertical',
          coordinateSpace: {
            writingMode: 'vertical-rl',
            logicalToPhysical: { a: 0, b: 1, c: -1, d: 0, e: 612, f: 0 },
            physicalToLogical: { a: 0, b: -1, c: 1, d: 0, e: 0, f: 612 },
          },
          blockStartPt: 72, blockEndPt: 540,
          flowDomainIds: ['vertical-body'], section: strictVerticalSection,
        }],
      }],
    };

    expect(() => assertDocumentLayout(layout)).not.toThrow();
  });

  it('rejects genuine logical vertical block overflow regardless of physical page containment', () => {
    const node = drawing('vertical-overflow', rect(72, 530, 100, 20), {
      flowDomainId: 'vertical-body',
    });
    const base = documentWith([node]);
    const layout: DocumentLayout = {
      ...base,
      pages: [{
        ...base.pages[0]!, section: verticalSection, sectionOccurrenceId: 'section:vertical',
        flowDomains: [{
          id: 'vertical-body', kind: 'body',
          logicalBounds: rect(72, 72, 648, 468), physicalBounds: rect(72, 72, 468, 648),
        }],
        sectionRegions: [{
          id: 'region:vertical', sectionOccurrenceId: 'section:vertical',
          coordinateSpace: {
            writingMode: 'vertical-rl',
            logicalToPhysical: { a: 0, b: 1, c: -1, d: 0, e: 612, f: 0 },
            physicalToLogical: { a: 0, b: -1, c: 1, d: 0, e: 0, f: 612 },
          },
          blockStartPt: 72, blockEndPt: 540,
          flowDomainIds: ['vertical-body'], section: verticalSection,
        }],
      }],
    };

    expect(() => assertDocumentLayout(layout))
      .toThrow(/BOTTOM_MARGIN_INVASION|FLOW_DOMAIN_INVASION/);
  });

  it('rejects a retained physical domain that is not derived from its region transform', () => {
    const base = documentWith([]);
    const layout: DocumentLayout = {
      ...base,
      pages: [{
        ...base.pages[0]!, section: verticalSection, sectionOccurrenceId: 'section:vertical',
        flowDomains: [{
          id: 'vertical-body', kind: 'body',
          logicalBounds: rect(72, 72, 648, 468), physicalBounds: rect(0, 0, 1, 1),
        }],
        sectionRegions: [{
          id: 'region:vertical', sectionOccurrenceId: 'section:vertical',
          coordinateSpace: {
            writingMode: 'vertical-rl',
            logicalToPhysical: { a: 0, b: 1, c: -1, d: 0, e: 612, f: 0 },
            physicalToLogical: { a: 0, b: -1, c: 1, d: 0, e: 0, f: 612 },
          },
          blockStartPt: 72, blockEndPt: 540,
          flowDomainIds: ['vertical-body'], section: verticalSection,
        }],
      }],
    };

    expect(() => assertDocumentLayout(layout)).toThrow(/INVALID_GEOMETRY/);
  });

  it('requires equal dual-space bounds on transitional pages without regions', () => {
    const base = documentWith([]);
    const unequal: DocumentLayout = {
      ...base,
      pages: [{
        ...base.pages[0]!,
        flowDomains: [{
          ...bodyDomain,
          physicalBounds: rect(0, 0, 468, 648),
        }],
      }],
    };

    expect(() => assertDocumentLayout(unequal)).toThrow(/INVALID_GEOMETRY/);
    expect(() => assertDocumentLayout(base)).not.toThrow();
  });

  it('requires page-start ownership and the exact supported retained transform', () => {
    const base = documentWith([]);
    const region = {
      id: 'region:body', sectionOccurrenceId: 'section:first',
      coordinateSpace: {
        writingMode: 'horizontal-tb' as const,
        logicalToPhysical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
        physicalToLogical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
      },
      blockStartPt: 72, blockEndPt: 720,
      flowDomainIds: ['body'], section: horizontalSection,
    };
    const page = {
      ...base.pages[0]!, section: horizontalSection,
      sectionOccurrenceId: 'section:first', sectionRegions: [region],
    };

    expect(() => assertDocumentLayout({ ...base, pages: [page] })).not.toThrow();
    expect(() => assertDocumentLayout({
      ...base, pages: [{ ...page, sectionOccurrenceId: 'section:other' }],
    })).toThrow(/page-start section occurrence/);
    expect(() => assertDocumentLayout({
      ...base,
      pages: [{
        ...page,
        sectionRegions: [{
          ...region,
          coordinateSpace: {
            ...region.coordinateSpace,
            logicalToPhysical: { ...region.coordinateSpace.logicalToPhysical, e: 1 },
          },
        }],
      }],
    })).toThrow(/invalid coordinate transform/);
  });

  it('rejects clone data whose region direction or columns contradict its section', () => {
    const valid = regionLayout();
    expect(() => assertDocumentLayout(valid)).not.toThrow();
    expectInvalidGeometry(() => assertDocumentLayout({
      ...valid,
      pages: [{
        ...valid.pages[0]!,
        sectionRegions: [{
          ...valid.pages[0]!.sectionRegions![0]!,
          coordinateSpace: {
            ...valid.pages[0]!.sectionRegions![0]!.coordinateSpace,
            writingMode: 'vertical-rl',
          },
        }],
      }],
    }));
    expectInvalidGeometry(() => assertDocumentLayout({
      ...valid,
      pages: [{
        ...valid.pages[0]!,
        sectionRegions: [{
          ...valid.pages[0]!.sectionRegions![0]!,
          section: { ...horizontalSection, columns: [{ xPt: 73, wPt: 468 }] },
        }],
      }],
    }));
  });

  it('rejects duplicate section occurrence identities in cloned page regions', () => {
    expectInvariantCode(
      'INVALID_REFERENCE',
      () => assertDocumentLayout(twoRegionLayout('section:body')),
    );
  });

  it('rejects cloned page regions that mix logical coordinate systems', () => {
    expectInvalidGeometry(() => assertDocumentLayout(
      twoRegionLayout('section:second', 'vertical-rl'),
    ));
  });

  it('rejects page-start section facts that contradict the first region clone', () => {
    const valid = regionLayout();
    const mismatches: SectionLayoutContext[] = [
      { ...horizontalSection, geometry: { ...horizontalSection.geometry, marginLeft: 73 } },
      { ...horizontalSection, columns: [{ xPt: 73, wPt: 468 }] },
      { ...horizontalSection, textDirection: 'lrTbV' },
      { ...horizontalSection, grid: { ...horizontalSection.grid, kind: 'lines' } },
      { ...horizontalSection, verticalAlignment: 'center' },
      {
        ...horizontalSection,
        lineNumbering: { start: 1, countBy: 1, restart: 'newPage' },
      },
    ];
    for (const section of mismatches) {
      expectInvalidGeometry(() => assertDocumentLayout({
        ...valid, pages: [{ ...valid.pages[0]!, section }],
      }));
    }
  });

  it('rejects invalid region intervals and owned logical domain geometry', () => {
    const valid = regionLayout();
    const region = valid.pages[0]!.sectionRegions![0]!;
    const domain = valid.pages[0]!.flowDomains[0]!;
    const invalidRegions = [
      { ...region, blockStartPt: -1 },
      { ...region, blockEndPt: 793 },
      { ...region, blockEndPt: region.blockStartPt },
    ];
    for (const invalidRegion of invalidRegions) {
      expectInvalidGeometry(() => assertDocumentLayout({
        ...valid,
        pages: [{ ...valid.pages[0]!, sectionRegions: [invalidRegion] }],
      }));
    }

    const invalidDomains = [
      { ...domain, logicalBounds: rect(72, 71, 468, 649) },
      { ...domain, logicalBounds: rect(72, 72, 0, 648) },
      { ...domain, logicalBounds: rect(72, 72, 468, 0) },
      { ...domain, logicalBounds: rect(71, 72, 469, 648) },
      { ...domain, physicalBounds: rect(600, 72, 20, 648) },
    ];
    for (const invalidDomain of invalidDomains) {
      expectInvalidGeometry(() => assertDocumentLayout({
        ...valid,
        pages: [{ ...valid.pages[0]!, flowDomains: [invalidDomain] }],
      }));
    }
  });

  it('rejects overlapping and out-of-page inline domains', () => {
    const valid = regionLayout();
    const region = valid.pages[0]!.sectionRegions![0]!;
    const domains: FlowDomain[] = [
      {
        id: 'body:first', kind: 'body',
        logicalBounds: rect(72, 72, 250, 648), physicalBounds: rect(72, 72, 250, 648),
      },
      {
        id: 'body:second', kind: 'body',
        logicalBounds: rect(300, 72, 240, 648), physicalBounds: rect(300, 72, 240, 648),
      },
    ];
    const section = { ...horizontalSection, columns: [{ xPt: 72, wPt: 250 }, { xPt: 300, wPt: 240 }] };
    expectInvalidGeometry(() => assertDocumentLayout({
      ...valid,
      pages: [{
        ...valid.pages[0]!, section,
        flowDomains: domains,
        sectionRegions: [{ ...region, section, flowDomainIds: domains.map(({ id }) => id) }],
      }],
    }));

    const outsideSection = {
      ...horizontalSection,
      columns: [{ xPt: 500, wPt: 200 }],
    };
    const outsideDomain: FlowDomain = {
      id: 'body:outside', kind: 'body',
      logicalBounds: rect(500, 72, 200, 648),
      physicalBounds: rect(500, 72, 200, 648),
    };
    expectInvalidGeometry(() => assertDocumentLayout({
      ...valid,
      pages: [{
        ...valid.pages[0]!, section: outsideSection,
        flowDomains: [outsideDomain],
        sectionRegions: [{
          ...region, section: outsideSection, flowDomainIds: [outsideDomain.id],
        }],
      }],
    }));
  });

  it('rejects exact derived physical bounds that leave the retained physical page box', () => {
    const valid = regionLayout();
    const page = valid.pages[0]!;
    const region = page.sectionRegions![0]!;
    const section = {
      ...horizontalSection,
      columns: [{ xPt: 0, wPt: 468 }],
    };
    const domain: FlowDomain = {
      id: 'body:physical-overflow', kind: 'body',
      logicalBounds: rect(0, 72, 468, 648),
      physicalBounds: rect(0, 72, 468, 648),
    };
    expectInvalidGeometry(() => assertDocumentLayout({
      ...valid,
      pages: [{
        ...page,
        geometry: { ...page.geometry, xPt: 10 },
        section,
        flowDomains: [domain],
        sectionRegions: [{ ...region, section, flowDomainIds: [domain.id] }],
      }],
    }));
  });

  it('normalizes malformed coordinate clones and invalid page extents to INVALID_GEOMETRY', () => {
    const valid = regionLayout();
    const region = valid.pages[0]!.sectionRegions![0]!;
    const malformedCoordinateSpaces: unknown[] = [
      undefined,
      {},
      { writingMode: 'diagonal', logicalToPhysical: {}, physicalToLogical: {} },
      { writingMode: 'horizontal-tb', logicalToPhysical: undefined, physicalToLogical: {} },
    ];
    for (const coordinateSpace of malformedCoordinateSpaces) {
      expectInvalidGeometry(() => assertDocumentLayout({
        ...valid,
        pages: [{
          ...valid.pages[0]!,
          sectionRegions: [{ ...region, coordinateSpace }],
        }],
      } as DocumentLayout));
    }
    for (const [widthPt, heightPt] of [[0, 792], [-1, 792], [612, 0]]) {
      expectInvalidGeometry(() => assertDocumentLayout({
        ...valid,
        pages: [{
          ...valid.pages[0]!,
          geometry: { ...valid.pages[0]!.geometry, widthPt, heightPt },
        }],
      }));
    }
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
      layoutParagraph(input, placement) {
        calls.push(`paragraph:${input.source.path.join('.')}:${placement.cursor.yPt}`);
        const layout = {
          ...drawing('p1', rect(10, placement.cursor.yPt, 100, 12)),
          kind: 'paragraph' as const,
          spacing: { beforePt: 0, afterPt: 0 }, contextualSpacing: false,
          lines: [], borders: [], resources: [], drawings: [], textBoxes: [], events: [], exclusions: [],
        };
        return { layout, nextCursor: { xPt: 10, yPt: placement.cursor.yPt + 12 } };
      },
      layoutTable(input, placement) {
        calls.push(`table:${input.source.path.join('.')}:${placement.cursor.yPt}`);
        const layout = {
          ...drawing('t2', rect(10, placement.cursor.yPt, 100, 18)),
          kind: 'table' as const,
          columnWidthsPt: [], rows: [], borders: [],
        };
        return { layout, nextCursor: { xPt: 10, yPt: placement.cursor.yPt + 18 } };
      },
    };
    const services = serviceStubs();

    const result = layoutFlowBlocks({
      source: source(0),
      container: { id: 'body', kind: 'body', bounds: rect(10, 20, 100, 200) },
      cursor: { xPt: 10, yPt: 20 },
      blocks: [
        { kind: 'paragraph', source: source(1) },
        tableInput(2),
      ],
    }, services, algorithms);

    expect(calls).toEqual(['paragraph:1:20', 'table:2:32']);
    expect(result.blocks.map((block) => block.id)).toEqual(['p1', 't2']);
    expect(result.advancePt).toBe(30);
    expect(result.flowBounds).toEqual(rect(10, 20, 100, 30));
    expect(result.nextCursor).toEqual({ xPt: 10, yPt: 50 });
  });

  it('rejects a block assigned outside its enclosing flow domain', () => {
    const algorithms: BlockLayoutAlgorithms = {
      layoutParagraph(_input, placement) {
        const layout = {
          ...drawing('p1', rect(10, placement.cursor.yPt, 100, 12), { flowDomainId: 'other' }),
          kind: 'paragraph' as const,
          spacing: { beforePt: 0, afterPt: 0 }, contextualSpacing: false,
          lines: [], borders: [], resources: [], drawings: [], textBoxes: [], events: [], exclusions: [],
        };
        return { layout, nextCursor: { xPt: 10, yPt: placement.cursor.yPt + 12 } };
      },
      layoutTable() {
        throw new Error('not used');
      },
    };
    const services = serviceStubs();

    expect(() => layoutFlowBlocks({
      source: source(0),
      container: { id: 'cell:1', kind: 'tableCell', bounds: rect(10, 20, 100, 200) },
      cursor: { xPt: 10, yPt: 20 },
      blocks: [{ kind: 'paragraph', source: source(1) }],
    }, services, algorithms)).toThrow(/INVALID_REFERENCE/);
  });

  it('rejects invalid containers and initial cursors before dispatch', () => {
    const unused: BlockLayoutAlgorithms = {
      layoutParagraph() { throw new Error('not used'); },
      layoutTable() { throw new Error('not used'); },
    };
    const services = serviceStubs();
    const base = {
      source: source(0),
      blocks: [],
      container: { id: 'body', kind: 'body' as const, bounds: rect(10, 20, 100, 200) },
      cursor: { xPt: 10, yPt: 20 },
    };

    expect(() => layoutFlowBlocks({
      ...base,
      container: { ...base.container, bounds: rect(10, 20, Number.NaN, 200) },
    }, services, unused)).toThrow(/INVALID_GEOMETRY/);
    expect(() => layoutFlowBlocks({
      ...base,
      cursor: { xPt: 10, yPt: 19 },
    }, services, unused)).toThrow(/INVALID_GEOMETRY/);
    expect(() => layoutFlowBlocks({
      ...base,
      cursor: { xPt: 111, yPt: 20 },
    }, services, unused)).toThrow(/INVALID_GEOMETRY/);
  });
});
