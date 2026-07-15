import { describe, expect, it } from 'vitest';
import type {
  ParagraphLayout,
  ResolvedBorderSegment,
  TableCellLayout,
  TableLayout,
  TableRowLayout,
} from '../layout/types.js';
import type { CanvasPaintContext, CanvasPaintResourcePainter, PaintCanvas2D } from './types.js';

const resources: CanvasPaintResourcePainter = {
  paint(resourceKey, kind): never {
    throw new Error(`Unexpected ${kind} resource: ${resourceKey}`);
  },
};

function paragraph(): ParagraphLayout {
  return {
    kind: 'paragraph', id: 'paragraph-0',
    source: { story: 'body', storyInstance: 'body', path: [0, 0, 0] },
    flowDomainId: 'body', ordinaryFlow: true,
    flowBounds: { xPt: 12, yPt: 22, widthPt: 30, heightPt: 12 },
    inkBounds: { xPt: 12, yPt: 22, widthPt: 20, heightPt: 10 },
    advancePt: 12, spacing: { beforePt: 0, afterPt: 0 }, contextualSpacing: false,
    lines: [{
      range: { start: 0, end: 5 },
      bounds: { xPt: 12, yPt: 22, widthPt: 20, heightPt: 10 },
      baselinePt: 30, advancePt: 12,
      placements: [{
        kind: 'text', text: 'child', range: { start: 0, end: 5 },
        origin: { xPt: 12, yPt: 30 },
        bounds: { xPt: 12, yPt: 22, widthPt: 20, heightPt: 10 },
        advancePt: 20,
        clusters: [{ range: { start: 0, end: 5 }, offset: { xPt: 0, yPt: 0 }, advancePt: 20 }],
        paintOps: [{
          text: 'child', range: { start: 0, end: 5 }, offset: { xPt: 0, yPt: 0 },
          letterSpacingPt: 0, scaleX: 1, direction: 'ltr', kerning: 'auto',
          writingMode: 'horizontal-tb',
        }],
        color: { kind: 'explicit', color: '#112233' },
        fontRoute: { familyList: '"Test Sans"', scope: 'native', fingerprint: 'test-sans' },
        fontSizePt: 10, fontWeight: 400, fontStyle: 'normal', direction: 'ltr',
        decorations: [],
      }],
    }],
    borders: [], resources: [], drawings: [], textBoxes: [], events: [], exclusions: [],
  };
}

function tableLayout(): TableLayout {
  const child = paragraph();
  const segment = {
    edge: 'left',
    from: { xPt: 50, yPt: 20 }, to: { xPt: 50, yPt: 36 },
    color: '#445566', widthPt: 1, authoredStyle: 'single', style: 'solid',
  } as ResolvedBorderSegment;
  const cells = [
    {
      kind: 'table-cell', id: 'cell-0',
      source: { story: 'body', storyInstance: 'body', path: [0, 0, 0] },
      flowDomainId: 'body', ordinaryFlow: true,
      flowBounds: { xPt: 10, yPt: 20, widthPt: 40, heightPt: 16 },
      inkBounds: { xPt: 10, yPt: 20, widthPt: 40, heightPt: 16 },
      contentBounds: { xPt: 12, yPt: 22, widthPt: 36, heightPt: 12 },
      advancePt: 16, verticalMerge: 'none', vAlign: 'top',
      background: { color: '#abcdef' },
      blocks: [{ layout: child, offsetPt: 2, advancePt: 12 }],
    },
    {
      kind: 'table-cell', id: 'cell-1',
      source: { story: 'body', storyInstance: 'body', path: [0, 0, 1] },
      flowDomainId: 'body', ordinaryFlow: true,
      flowBounds: { xPt: 50, yPt: 20, widthPt: 40, heightPt: 16 },
      inkBounds: { xPt: 50, yPt: 20, widthPt: 40, heightPt: 16 },
      contentBounds: { xPt: 50, yPt: 20, widthPt: 40, heightPt: 16 },
      advancePt: 16, verticalMerge: 'none', vAlign: 'top', blocks: [],
    },
  ] as unknown as readonly TableCellLayout[];
  const rows = [{
    kind: 'table-row', id: 'row-0',
    source: { story: 'body', storyInstance: 'body', path: [0, 0] },
    flowDomainId: 'body', ordinaryFlow: true,
    flowBounds: { xPt: 10, yPt: 20, widthPt: 80, heightPt: 16 },
    inkBounds: { xPt: 10, yPt: 20, widthPt: 80, heightPt: 16 },
    advancePt: 16, cells,
  }] as unknown as readonly TableRowLayout[];
  return {
    kind: 'table', id: 'table-0',
    source: { story: 'body', storyInstance: 'body', path: [0] },
    flowDomainId: 'body', ordinaryFlow: true,
    flowBounds: { xPt: 10, yPt: 20, widthPt: 80, heightPt: 16 },
    inkBounds: { xPt: 10, yPt: 20, widthPt: 80, heightPt: 16 },
    advancePt: 16, columnWidthsPt: [40, 40], rows, borders: [segment],
  } as TableLayout;
}

describe('paintTableLayout', () => {
  it('paints only retained backgrounds, child layouts, and resolved borders without measuring', async () => {
    const operations: unknown[] = [];
    const target = {
      globalAlpha: 1, fillStyle: '', strokeStyle: '', lineWidth: 1,
      font: '', textAlign: 'left' as CanvasTextAlign,
      textBaseline: 'alphabetic' as CanvasTextBaseline,
      direction: 'ltr' as CanvasDirection, letterSpacing: '0px',
      fontKerning: 'auto' as CanvasFontKerning,
      save() { operations.push('save'); }, restore() { operations.push('restore'); },
      beginPath() { operations.push('beginPath'); },
      rect() {}, clip() {}, translate() {}, rotate() {}, scale() {},
      fillRect(x: number, y: number, width: number, height: number) {
        operations.push(['fillRect', x, y, width, height, this.fillStyle]);
      },
      strokeRect() {}, setLineDash(value: number[]) { operations.push(['dash', value]); },
      moveTo(x: number, y: number) { operations.push(['moveTo', x, y]); },
      lineTo(x: number, y: number) { operations.push(['lineTo', x, y]); },
      stroke() { operations.push(['stroke', this.strokeStyle, this.lineWidth]); },
      fill() {}, drawImage() {},
      fillText(text: string, x: number, y: number) { operations.push(['fillText', text, x, y]); },
    };
    const ctx = new Proxy(target, {
      get(object, property, receiver) {
        if (property === 'measureText') throw new Error('table paint must not measure text');
        return Reflect.get(object, property, receiver);
      },
    }) as unknown as PaintCanvas2D;
    const context: CanvasPaintContext = { ctx, scale: 1, dpr: 1, resources };
    const node = tableLayout();
    const before = JSON.stringify(node);
    const { paintTableLayout } = await import('./canvas-table.js');

    expect(() => paintTableLayout(node, context)).not.toThrow();

    expect(operations).toContainEqual(['fillRect', 10, 20, 40, 16, '#abcdef']);
    expect(operations).toContainEqual(['fillText', 'child', 12, 30]);
    // The shared retained-border painter applies the same odd-device-pixel
    // crisp offset used by paragraph and run borders.
    expect(operations).toContainEqual(['moveTo', 50.5, 20]);
    expect(operations).toContainEqual(['lineTo', 50.5, 36]);
    expect(operations.filter((operation) =>
      Array.isArray(operation) && operation[0] === 'stroke')).toHaveLength(1);
    expect(JSON.stringify(node)).toBe(before);
  });

  it('reports text-run bounds in final CSS pixels after placed table transforms', async () => {
    const runs: Array<Readonly<{
      text: string;
      x: number;
      y: number;
      w: number;
      h: number;
      fontSize: number;
    }>> = [];
    const ctx = {
      globalAlpha: 1, fillStyle: '', strokeStyle: '', lineWidth: 1,
      font: '', textAlign: 'left' as CanvasTextAlign,
      textBaseline: 'alphabetic' as CanvasTextBaseline,
      direction: 'ltr' as CanvasDirection, letterSpacing: '0px',
      fontKerning: 'auto' as CanvasFontKerning,
      save() {}, restore() {}, beginPath() {}, rect() {}, clip() {},
      translate() {}, rotate() {}, scale() {}, fillRect() {}, strokeRect() {},
      setLineDash() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {}, drawImage() {},
      fillText() {},
    } as unknown as PaintCanvas2D;
    const { paintPlacedTableLayout } = await import('./canvas-table.js');

    paintPlacedTableLayout(tableLayout(), { xPt: 110, yPt: 220 }, {
      ctx,
      scale: 2,
      dpr: 1,
      resources,
      onTextRun: (run) => runs.push(run),
    });

    expect(runs).toEqual([
      expect.objectContaining({
        text: 'child',
        x: 224,
        y: 444,
        w: 40,
        h: 20,
        fontSize: 20,
      }),
    ]);
  });

  it('preserves a nested table alignment offset inside the outer cell content band', async () => {
    const outer = tableLayout();
    const nestedParagraph = paragraph();
    const nested: TableLayout = {
      kind: 'table', id: 'nested-table',
      source: { story: 'body', storyInstance: 'body', path: [0, 0, 0, 0] },
      flowDomainId: 'nested', ordinaryFlow: true,
      // The nested table is centered 20pt into its own available content band.
      flowBounds: { xPt: 20, yPt: 0, widthPt: 40, heightPt: 12 },
      inkBounds: { xPt: 20, yPt: 0, widthPt: 40, heightPt: 12 },
      advancePt: 12, columnWidthsPt: [40], borders: [],
      rows: [{
        kind: 'table-row', id: 'nested-row',
        source: { story: 'body', storyInstance: 'body', path: [0, 0, 0, 0, 0] },
        flowDomainId: 'nested', ordinaryFlow: true,
        flowBounds: { xPt: 20, yPt: 0, widthPt: 40, heightPt: 12 },
        inkBounds: { xPt: 20, yPt: 0, widthPt: 40, heightPt: 12 },
        advancePt: 12, heightPt: 12, contentHeightPt: 12,
        cells: [{
          kind: 'table-cell', id: 'nested-cell',
          source: { story: 'body', storyInstance: 'body', path: [0, 0, 0, 0, 0, 0] },
          flowDomainId: 'nested', ordinaryFlow: true,
          flowBounds: { xPt: 20, yPt: 0, widthPt: 40, heightPt: 12 },
          inkBounds: { xPt: 20, yPt: 0, widthPt: 40, heightPt: 12 },
          contentBounds: { xPt: 22, yPt: 0, widthPt: 36, heightPt: 12 },
          advancePt: 12, verticalMerge: 'none', vAlign: 'top',
          blocks: [{ layout: nestedParagraph, offsetPt: 0, advancePt: 12 }],
        }],
      }],
    };
    const firstCell = outer.rows[0]!.cells[0]!;
    const withNested: TableLayout = {
      ...outer,
      rows: [{
        ...outer.rows[0]!,
        cells: [{
          ...firstCell,
          blocks: [{ layout: nested, offsetPt: 0, advancePt: 12 }],
        }, ...outer.rows[0]!.cells.slice(1)],
      }],
    };
    const runs: Array<{ x: number }> = [];
    const ctx = {
      globalAlpha: 1, fillStyle: '', strokeStyle: '', lineWidth: 1,
      font: '', textAlign: 'left' as CanvasTextAlign,
      textBaseline: 'alphabetic' as CanvasTextBaseline,
      direction: 'ltr' as CanvasDirection, letterSpacing: '0px',
      fontKerning: 'auto' as CanvasFontKerning,
      save() {}, restore() {}, beginPath() {}, rect() {}, clip() {},
      translate() {}, rotate() {}, scale() {}, fillRect() {}, strokeRect() {},
      setLineDash() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {}, drawImage() {},
      fillText() {},
    } as unknown as PaintCanvas2D;
    const { paintTableLayout } = await import('./canvas-table.js');

    paintTableLayout(withNested, {
      ctx, scale: 1, dpr: 1, resources,
      onTextRun: (run) => runs.push({ x: run.x }),
    });

    // outer content x=12 + nested-local paragraph x=22. The nested table's
    // 20pt centered origin is already retained by its child geometry.
    expect(runs).toEqual([{ x: 34 }]);
  });
});
