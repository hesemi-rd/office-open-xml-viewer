import { beforeAll, describe, expect, it } from 'vitest';
import {
  bodyFragmentFor,
  paginateDocument,
} from '../renderer.js';
import type {
  BodyElement,
  CellElement,
  DocParagraph,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  PaginatedBodyElement,
  SectionProps,
} from '../types.js';
import type { TableFragmentLayout } from './table-pagination.js';

function makeStubCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const context = {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText(text: string) {
      const fontSize = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...text].length * fontSize * 0.5,
        fontBoundingBoxAscent: fontSize * 0.8,
        fontBoundingBoxDescent: fontSize * 0.2,
        actualBoundingBoxAscent: fontSize * 0.8,
        actualBoundingBoxDescent: fontSize * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {}, createLinearGradient() {
      return { addColorStop() {} };
    },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000000', strokeStyle: '#000000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  };
  (context as unknown as { canvas: unknown }).canvas = { width: 2000, height: 2000 };
  return context as unknown as CanvasRenderingContext2D;
}

beforeAll(() => {
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    getContext() { return makeStubCtx(); }
  };
});

function emptyBorders() {
  return { top: null, right: null, bottom: null, left: null, insideH: null, insideV: null };
}

function singleBorder(width: number) {
  return { style: 'single', width, color: '#000000' } as const;
}

function paragraph(text: string, fontSize = 10): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: text.length === 0 ? [] : [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize, color: null, fontFamily: 'Times New Roman',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null,
      hyperlink: null,
    } as DocParagraph['runs'][number]],
    defaultFontSize: fontSize, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

function cell(content: CellElement[], overrides: Partial<DocTableCell> = {}): DocTableCell {
  return {
    content, colSpan: 1, vMerge: null, borders: emptyBorders(),
    background: null, vAlign: 'top', widthPt: null,
    ...overrides,
  } as unknown as DocTableCell;
}

function textCell(text: string): DocTableCell {
  return cell([{ type: 'paragraph', ...paragraph(text) } as unknown as CellElement]);
}

function row(cells: DocTableCell[], overrides: Partial<DocTableRow> = {}): DocTableRow {
  return {
    cells, rowHeight: null, rowHeightRule: 'auto', isHeader: false,
    ...overrides,
  } as unknown as DocTableRow;
}

function table(
  rows: DocTableRow[],
  columnWidthsPt: number[],
  overrides: Partial<DocTable> = {},
): DocTable {
  return {
    type: 'table', rows, colWidths: columnWidthsPt, borders: emptyBorders(),
    cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
    jc: 'left', layout: 'fixed',
    ...overrides,
  } as unknown as DocTable;
}

function documentModel(
  body: BodyElement[],
  pageHeight: number,
  pageWidth = 200,
): DocxDocumentModel {
  return {
    section: {
      pageWidth, pageHeight,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
      sectionStart: 'nextPage', columns: null,
    } as SectionProps,
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
    footnotes: [],
  } as unknown as DocxDocumentModel;
}

function retainedTopLevelTables(pages: PaginatedBodyElement[][]): TableFragmentLayout[] {
  return pages.flatMap((page) => page.flatMap((element) => {
    if (element.type !== 'table') return [];
    const retained = bodyFragmentFor(element)?.fragment;
    return retained?.kind === 'table' && 'flowBounds' in retained
      ? [retained as TableFragmentLayout]
      : [];
  }));
}

function expectDeeplyFrozen(value: unknown, seen = new WeakSet<object>()): void {
  if (value === null || typeof value !== 'object' || seen.has(value)) return;
  seen.add(value);
  expect(Object.isFrozen(value)).toBe(true);
  for (const child of Object.values(value)) expectDeeplyFrozen(child, seen);
}

describe('retained table pagination contracts', () => {
  it('retains nested page-split geometry as clone-safe immutable parser-independent data', () => {
    const nested = table(
      Array.from({ length: 4 }, (_unused, index) => row(
        [textCell(`nested ${index}`)],
        { rowHeight: 30, rowHeightRule: 'exact' },
      )),
      [80],
    );
    const sourceCell = cell([{ type: 'table', ...nested } as unknown as CellElement], {
      background: '112233',
    });
    const source = table([row([sourceCell])], [120]);
    const pages = paginateDocument(documentModel([source as unknown as BodyElement], 80));
    const nestedFragments = retainedTopLevelTables(pages).flatMap((fragment) =>
      fragment.rows.flatMap((fragmentRow) => fragmentRow.cells.flatMap((fragmentCell) =>
        fragmentCell.blocks.flatMap((block) => block.layout.kind === 'table'
          ? [block.layout as TableFragmentLayout]
          : []),
      )),
    );

    expect(nestedFragments.length).toBeGreaterThan(1);
    for (const fragment of nestedFragments) {
      expect(() => structuredClone(fragment)).not.toThrow();
      expectDeeplyFrozen(fragment);
    }
    const fingerprint = JSON.stringify(nestedFragments);

    source.jc = 'right';
    sourceCell.background = 'ffffff';
    nested.rows[0]!.cells[0]!.vAlign = 'bottom';
    nested.rows[1]!.rowHeight = 99;

    expect(JSON.stringify(nestedFragments)).toBe(fingerprint);
  });

  it('isolates placements and widths across independent pagination runs', () => {
    const source = table([row([textCell('independent')])], [200], { widthPct: 5000 });
    const body = [source as unknown as BodyElement];
    const wideModel = documentModel(body, 100, 200);
    const narrowModel = {
      ...wideModel,
      section: { ...wideModel.section, pageWidth: 150 },
    } as DocxDocumentModel;

    const [wideElement] = paginateDocument(wideModel)[0] ?? [];
    if (!wideElement) throw new Error('expected wide table placement');
    const wide = bodyFragmentFor(wideElement);
    const wideFingerprint = JSON.stringify(wide);

    const [narrowElement] = paginateDocument(narrowModel)[0] ?? [];
    if (!narrowElement) throw new Error('expected narrow table placement');
    const narrow = bodyFragmentFor(narrowElement);

    expect(wide?.fragment.kind).toBe('table');
    expect(narrow?.fragment.kind).toBe('table');
    expect(wide?.widthPt).toBeCloseTo(180, 6);
    expect(narrow?.widthPt).toBeCloseTo(130, 6);
    expect(wide?.xPt).toBeCloseTo(10, 6);
    expect(narrow?.xPt).toBeCloseTo(10, 6);
    expect(wide?.fragment).not.toBe(narrow?.fragment);
    expect(JSON.stringify(wide)).toBe(wideFingerprint);
    expect(source).not.toHaveProperty('tableColWidthsPt');
    expect(source).not.toHaveProperty('tableRowHeightsPt');
  });

  it('resolves page-local outer border ink for every auto-height slice', () => {
    const outer = singleBorder(4);
    const inside = singleBorder(1);
    const source = table(
      Array.from({ length: 3 }, (_unused, index) => row(
        [textCell(`row ${index}`)],
        { rowHeight: 20, rowHeightRule: 'auto' },
      )),
      [120],
      {
        borders: {
          top: outer, bottom: outer, left: null, right: null,
          insideH: inside, insideV: null,
        },
      },
    );

    const fragments = retainedTopLevelTables(
      paginateDocument(documentModel([source as unknown as BodyElement], 50)),
    );

    expect(fragments).toHaveLength(3);
    for (const fragment of fragments) {
      expect(fragment.rows).toHaveLength(1);
      expect(fragment.rows[0]?.heightPt).toBe(20);
      expect(fragment.advancePt).toBe(20);
      expect(fragment.borders.filter((border) => border.edge === 'top')).toEqual([
        expect.objectContaining({ widthPt: 4 }),
      ]);
      expect(fragment.borders.filter((border) => border.edge === 'bottom')).toEqual([
        expect.objectContaining({ widthPt: 4 }),
      ]);
      expect(fragment.inkBounds.yPt).toBeCloseTo(fragment.flowBounds.yPt - 2, 6);
      expect(fragment.inkBounds.heightPt).toBeCloseTo(fragment.flowBounds.heightPt + 4, 6);
    }
  });

  it('keeps thick collapsed outer border ink outside mixed row-track allocation', () => {
    const outer = singleBorder(12);
    const source = table(
      [
        row([cell([])], { rowHeight: 10, rowHeightRule: 'auto' }),
        row([cell([])], { rowHeight: 1, rowHeightRule: 'exact' }),
        row([cell([])], { rowHeight: 10, rowHeightRule: 'auto' }),
        row([cell([])], { rowHeight: 1, rowHeightRule: 'exact' }),
        row([cell([])], { rowHeight: 10, rowHeightRule: 'auto' }),
      ],
      [120],
      {
        borders: {
          top: outer, bottom: outer, left: null, right: null,
          insideH: null, insideV: null,
        },
      },
    );

    const fragments = retainedTopLevelTables(
      paginateDocument(documentModel([source as unknown as BodyElement], 52)),
    );

    expect(fragments).toHaveLength(1);
    const [fragment] = fragments;
    expect(fragment?.rows.map((fragmentRow) => fragmentRow.heightPt)).toEqual([10, 1, 10, 1, 10]);
    expect(fragment?.advancePt).toBe(32);
    expect(fragment?.flowBounds.heightPt).toBe(32);
    expect(fragment?.inkBounds.yPt).toBeCloseTo((fragment?.flowBounds.yPt ?? 0) - 6, 6);
    expect(fragment?.inkBounds.heightPt).toBe(44);
  });

  it('re-resolves repeated atLeast header boundaries and row tracks on every slice', () => {
    const outer = singleBorder(4);
    const inside = singleBorder(1);
    const source = table(
      [
        row([textCell('header')], {
          isHeader: true, rowHeight: 20, rowHeightRule: 'atLeast',
        }),
        ...Array.from({ length: 3 }, (_unused, index) => row(
          [textCell(`body ${index}`)],
          { rowHeight: 20, rowHeightRule: 'atLeast' },
        )),
      ],
      [120],
      {
        borders: {
          top: outer, bottom: outer, left: null, right: null,
          insideH: inside, insideV: null,
        },
      },
    );

    const fragments = retainedTopLevelTables(
      paginateDocument(documentModel([source as unknown as BodyElement], 70)),
    );

    expect(fragments).toHaveLength(3);
    for (const [index, fragment] of fragments.entries()) {
      expect(fragment.rows.map((fragmentRow) => fragmentRow.heightPt)).toEqual([20, 20]);
      expect(fragment.rows[0]?.ownership).toBe(index === 0 ? 'source' : 'repeated-header');
      expect(fragment.borders.filter((border) => border.edge === 'top')).toEqual([
        expect.objectContaining({ widthPt: 4 }),
      ]);
      expect(fragment.borders.filter((border) => border.edge === 'between')).toEqual([
        expect.objectContaining({ widthPt: 1 }),
      ]);
      expect(fragment.borders.filter((border) => border.edge === 'bottom')).toEqual([
        expect.objectContaining({ widthPt: 4 }),
      ]);
    }
  });

  it('fits body rows against the repeated header interior boundary, not an outer bottom', () => {
    const outer = singleBorder(4);
    const inside = singleBorder(1);
    const source = table(
      [
        row([textCell('header')], {
          isHeader: true, rowHeight: 20, rowHeightRule: 'atLeast',
        }),
        ...Array.from({ length: 4 }, (_unused, index) => row(
          [textCell(`body ${index}`)],
          { rowHeight: 20, rowHeightRule: 'exact' },
        )),
      ],
      [120],
      {
        borders: {
          top: outer, bottom: outer, left: null, right: null,
          insideH: inside, insideV: null,
        },
      },
    );

    const fragments = retainedTopLevelTables(
      paginateDocument(documentModel([source as unknown as BodyElement], 80)),
    );

    expect(fragments).toHaveLength(2);
    expect(fragments.map((fragment) => fragment.rows.length)).toEqual([3, 3]);
    expect(fragments.map((fragment) => fragment.advancePt)).toEqual([60, 60]);
    for (const fragment of fragments) {
      expect(fragment.rows.map((fragmentRow) => fragmentRow.logicalRowIndex)).toEqual(
        fragment === fragments[0] ? [0, 1, 2] : [0, 3, 4],
      );
      expect(fragment.borders.filter((border) => border.edge === 'between')).toHaveLength(2);
      expect(fragment.borders.filter((border) => border.edge === 'bottom')).toEqual([
        expect.objectContaining({ widthPt: 4 }),
      ]);
    }
  });

  it('retains split-row content advance and each piece own page-local boundary ink', () => {
    const outer = singleBorder(4);
    const inside = singleBorder(1);
    const source = table(
      [row([cell([{
        type: 'paragraph',
        ...paragraph('あ'.repeat(400)),
      } as unknown as CellElement])])],
      [120],
      {
        borders: {
          top: outer, bottom: outer, left: null, right: null,
          insideH: inside, insideV: null,
        },
      },
    );

    const fragments = retainedTopLevelTables(
      paginateDocument(documentModel([source as unknown as BodyElement], 80)),
    );

    expect(fragments.length).toBeGreaterThan(1);
    for (const [fragmentIndex, fragment] of fragments.entries()) {
      const fragmentRow = fragment.rows[0];
      expect(fragmentRow?.logicalRowIndex).toBe(0);
      expect(fragmentRow?.fragmentIndex).toBe(fragmentIndex);
      expect(fragment.advancePt).toBeCloseTo(fragmentRow?.contentHeightPt ?? 0, 8);
      expect(fragment.flowBounds.heightPt).toBeCloseTo(fragment.advancePt, 8);
      expect(fragment.borders.filter((border) => border.edge === 'top')).toEqual([
        expect.objectContaining({ widthPt: 4 }),
      ]);
      expect(fragment.borders.filter((border) => border.edge === 'bottom')).toEqual([
        expect.objectContaining({ widthPt: 4 }),
      ]);
      expect(fragment.inkBounds.yPt).toBeCloseTo(fragment.flowBounds.yPt - 2, 8);
      expect(fragment.inkBounds.heightPt).toBeCloseTo(fragment.advancePt + 4, 8);
    }
    const lineRanges = fragments.map((fragment) => {
      const range = fragment.rows[0]?.cells[0]?.contentRanges.find(
        (candidate) => candidate.kind === 'paragraph',
      );
      if (range?.kind !== 'paragraph') throw new Error('expected split paragraph ownership');
      return range;
    });
    expect(lineRanges[0]?.lineStart).toBe(0);
    for (let index = 1; index < lineRanges.length; index += 1) {
      expect(lineRanges[index]?.lineStart).toBe(lineRanges[index - 1]?.lineEnd);
    }
    expect(fragments.reduce((sum, fragment) => sum + fragment.advancePt, 0)).toBeCloseTo(
      fragments.reduce((sum, fragment) => sum + (fragment.rows[0]?.contentHeightPt ?? 0), 0),
      8,
    );
  });
});
