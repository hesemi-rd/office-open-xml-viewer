import { describe, it, expect, beforeAll } from 'vitest';
import { layoutDocument } from './document-layout.js';
import { bodyFragmentFor, computePages } from './renderer.js';
import { buildTableFragment } from './table-fragments.js';
import * as layoutFragments from './layout-fragments.js';
import {
  paragraphFragmentAdvancePt,
  tableFragmentHeightPt,
  flowFragmentAdvancePt,
  type TableFragment,
  type RowFragment,
  type CellFragment,
  type ParagraphFragment,
  type FlowFragment,
  type PlacedFragment,
} from './layout-fragments.js';
import type {
  BodyElement,
  CellElement,
  DocParagraph,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  SectionProps,
} from './types';

// ─────────────────────────────────────────────────────────────────────────────
// PR 6 Task 15 — table layout fragments.
//
// A table fragments into `TableFragment` → `RowFragment` → `CellFragment` →
// recursive `FlowFragment` (paragraph fragments + nested-table fragments), so body
// and table flow share one immutable fragment result (design §"Measured Fragment
// Model"). This suite pins:
//   • the pure `buildTableFragment` contract (row/cell recursion, vMerge roles,
//     repeated-header marking, continuation flags, source provenance, freezing);
//   • `layoutDocument` producing table fragments on the body pages, reusing the
//     paginator's real cell measurement (ordinary row breaks, auto-height splits,
//     repeated headers, nested tables, vertical merges, per-cell continuation).
// ─────────────────────────────────────────────────────────────────────────────

/** OffscreenCanvas polyfill with a linear glyph metric (width = fontPx * 0.5) so
 *  `layoutDocument`'s scale-1 measurement is deterministic in node. Mirrors
 *  document-layout.test.ts. */
function makeStubCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      const per = p * 0.5;
      return {
        width: [...s].length * per,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  };
  (ctx as unknown as { canvas: unknown }).canvas = { width: 2000, height: 2000 };
  return ctx as unknown as CanvasRenderingContext2D;
}

beforeAll(() => {
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    getContext() { return makeStubCtx(); }
  };
});

// ---- Model builders -----------------------------------------------------------

function emptyBorders() {
  return { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };
}

function para(text: string, over: Partial<DocParagraph> = {}): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: text === '' ? [] : [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'Times New Roman',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
    } as DocParagraph['runs'][number]],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
    ...over,
  } as unknown as DocParagraph;
}

function cell(content: CellElement[], over: Partial<DocTableCell> = {}): DocTableCell {
  return {
    content, colSpan: 1, vMerge: null, borders: emptyBorders(),
    background: null, vAlign: 'top', widthPt: null,
    ...over,
  } as unknown as DocTableCell;
}

/** Cell holding one text paragraph. */
function textCell(text: string, over: Partial<DocTableCell> = {}): DocTableCell {
  return cell([{ type: 'paragraph', ...para(text) } as unknown as CellElement], over);
}

function row(cells: DocTableCell[], over: Partial<DocTableRow> = {}): DocTableRow {
  return {
    cells, rowHeight: null, rowHeightRule: 'auto', isHeader: false,
    ...over,
  } as unknown as DocTableRow;
}

function table(rows: DocTableRow[], colWidths: number[], over: Partial<DocTable> = {}): DocTable {
  return {
    type: 'table',
    colWidths, rows, borders: emptyBorders(),
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left', layout: 'fixed',
    ...over,
  } as unknown as DocTable;
}

function doc(body: BodyElement[], pageHeight = 400): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 200, pageHeight,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage', columns: null,
  } as SectionProps;
  return {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
    footnotes: [],
  } as unknown as DocxDocumentModel;
}

/** Every placed table fragment on every page, in document order. */
function allTables(model: DocxDocumentModel): { placed: PlacedFragment; table: TableFragment }[] {
  const layout = layoutDocument(model);
  const out: { placed: PlacedFragment; table: TableFragment }[] = [];
  for (const page of layout.pages) {
    for (const placed of page.fragments) {
      if (placed.fragment.kind === 'table') out.push({ placed, table: placed.fragment });
    }
  }
  return out;
}

function firstParagraphBlock(cf: CellFragment): ParagraphFragment {
  const block = cf.blocks[0];
  if (!block || block.kind !== 'paragraph') throw new Error('expected a paragraph block');
  return block;
}

// ─────────────────────────────────────────────────────────────────────────────
// buildTableFragment — pure contract
// ─────────────────────────────────────────────────────────────────────────────

describe('buildTableFragment — pure recursion contract', () => {
  /** A recognizable stub paragraph fragment so blocks are identity-checkable
   *  without invoking real measurement. */
  function stubBlock(tag: string): FlowFragment {
    return { kind: 'paragraph', source: { tag } as unknown as DocParagraph } as unknown as FlowFragment;
  }

  it('produces one RowFragment per row and one CellFragment per cell, in order', () => {
    const t = table(
      [row([textCell('a'), textCell('b')]), row([textCell('c'), textCell('d')])],
      [30, 30],
    );
    const calls: number[] = [];
    const frag = buildTableFragment({
      table: t,
      columnWidthsPt: [30, 30],
      rowHeightsPt: [12, 18],
      continuesFromPreviousPage: false,
      continuesOnNextPage: false,
      repeatedHeaderRowCount: 0,
      buildCellBlocks: (_c, w) => { calls.push(w); return [stubBlock('x')]; },
    });
    expect(frag.kind).toBe('table');
    expect(frag.source).toBe(t);
    expect(frag.columnWidthsPt).toEqual([30, 30]);
    expect(frag.rows).toHaveLength(2);
    expect(frag.rows[0].cells).toHaveLength(2);
    expect(frag.rows[0].source).toBe(t.rows[0]);
    expect(frag.rows[0].sourceRowIndex).toBe(0);
    expect(frag.rows[1].sourceRowIndex).toBe(1);
    // each cell measured at its single-column width
    expect(calls).toEqual([30, 30, 30, 30]);
  });

  it('records per-row heightPt and the summed table height', () => {
    const t = table([row([textCell('a')]), row([textCell('b')])], [40]);
    const frag = buildTableFragment({
      table: t, columnWidthsPt: [40], rowHeightsPt: [15, 25],
      continuesFromPreviousPage: false, continuesOnNextPage: false, repeatedHeaderRowCount: 0,
      buildCellBlocks: () => [],
    });
    expect(frag.rows.map((r) => r.heightPt)).toEqual([15, 25]);
    expect(frag.rows.map((r) => r.cells[0].boxHeightPt)).toEqual([15, 25]);
    expect(tableFragmentHeightPt(frag)).toBe(40);
    expect(flowFragmentAdvancePt(frag)).toBe(40);
  });

  it('sums only fragment-owned ranged paragraph advances and nested-table heights', () => {
    const rangedParagraph = {
      kind: 'paragraph',
      source: para('ignored'),
      measured: {
        markOnly: false,
        lines: [
          { topYPt: 0, advancePt: 10 },
          { topYPt: 12, advancePt: 10 },
          { topYPt: 25, advancePt: 10 },
        ],
      },
      lineStart: 1,
      lineEnd: 3,
      leadingSpacePt: 2,
      trailingSpacePt: 4,
    } as unknown as ParagraphFragment;
    const nestedTable = {
      kind: 'table',
      source: table([], []),
      columnWidthsPt: [],
      rows: [
        { heightPt: 7 },
        { heightPt: 8 },
      ],
      continuesFromPreviousPage: false,
      continuesOnNextPage: false,
    } as unknown as TableFragment;
    const fragment = {
      source: textCell('ignored'),
      blocks: [rangedParagraph, nestedTable],
      verticalMerge: 'none',
      boxHeightPt: 50,
    } as CellFragment;
    const contentHeight = (
      layoutFragments as unknown as {
        cellFragmentContentHeightPt?: (cellFragment: CellFragment) => number;
      }
    ).cellFragmentContentHeightPt;

    expect(contentHeight).toBeTypeOf('function');
    if (!contentHeight) return;
    expect(contentHeight(fragment)).toBe(
      paragraphFragmentAdvancePt(rangedParagraph) + tableFragmentHeightPt(nestedTable),
    );
    expect(contentHeight(fragment)).toBe(44);
  });

  it('sums the spanned column widths for a gridSpan cell', () => {
    const t = table([row([textCell('wide', { colSpan: 2 }), textCell('c')])], [20, 20, 30]);
    const widths: number[] = [];
    buildTableFragment({
      table: t, columnWidthsPt: [20, 20, 30], rowHeightsPt: [12],
      continuesFromPreviousPage: false, continuesOnNextPage: false, repeatedHeaderRowCount: 0,
      buildCellBlocks: (_c, w) => { widths.push(w); return []; },
    });
    expect(widths).toEqual([40, 30]); // spanned 20+20, then 30
  });

  it('§17.4.15: builds cell blocks from columns after gridBefore', () => {
    const t = table(
      [row([textCell('middle')], { gridBefore: 1, gridAfter: 1 } as Partial<DocTableRow>)],
      [20, 40, 60],
    );
    const widths: number[] = [];
    buildTableFragment({
      table: t, columnWidthsPt: [20, 40, 60], rowHeightsPt: [12],
      continuesFromPreviousPage: false, continuesOnNextPage: false, repeatedHeaderRowCount: 0,
      buildCellBlocks: (_c, w) => { widths.push(w); return []; },
    });
    expect(widths).toEqual([40]);
  });

  it('classifies vMerge roles and renders no content for a continue cell', () => {
    const t = table(
      [
        row([textCell('r', { vMerge: true }), textCell('a')]),
        row([textCell('', { vMerge: false }), textCell('b')]),
      ],
      [30, 30],
    );
    const measured: string[] = [];
    const frag = buildTableFragment({
      table: t, columnWidthsPt: [30, 30], rowHeightsPt: [12, 12],
      continuesFromPreviousPage: false, continuesOnNextPage: false, repeatedHeaderRowCount: 0,
      buildCellBlocks: (c, _w) => {
        const p = c.content[0] as unknown as { runs?: { text: string }[] };
        measured.push(p.runs?.[0]?.text ?? '');
        return [stubBlock('m')];
      },
    });
    expect(frag.rows[0].cells[0].verticalMerge).toBe('restart');
    expect(frag.rows[0].cells[1].verticalMerge).toBe('none');
    expect(frag.rows[1].cells[0].verticalMerge).toBe('continue');
    expect(frag.rows[1].cells[0].blocks).toEqual([]); // continue: no content
    // the continue cell was never measured
    expect(measured).toEqual(['r', 'a', 'b']);
  });

  it('marks repeated header rows and maps slice rows back to source indices', () => {
    const t = table(
      [row([textCell('H')], { isHeader: true }), row([textCell('x')]), row([textCell('y')])],
      [40],
    );
    const frag = buildTableFragment({
      table: t, columnWidthsPt: [40], rowHeightsPt: [10, 12, 12],
      continuesFromPreviousPage: true, continuesOnNextPage: false,
      repeatedHeaderRowCount: 1,
      sourceRowIndexOf: (i) => (i === 0 ? 0 : i + 3),
      buildCellBlocks: () => [],
    });
    expect(frag.continuesFromPreviousPage).toBe(true);
    expect(frag.rows[0].repeatedHeader).toBe(true);
    expect(frag.rows[1].repeatedHeader).toBe(false);
    expect(frag.rows.map((r) => r.sourceRowIndex)).toEqual([0, 4, 5]);
  });

  it('deep-freezes the fragment, its rows, cells, and column widths (M-3)', () => {
    const t = table([row([textCell('a')])], [40]);
    const frag = buildTableFragment({
      table: t, columnWidthsPt: [40], rowHeightsPt: [12],
      continuesFromPreviousPage: false, continuesOnNextPage: false, repeatedHeaderRowCount: 0,
      buildCellBlocks: () => [],
    });
    expect(Object.isFrozen(frag)).toBe(true);
    expect(Object.isFrozen(frag.rows)).toBe(true);
    expect(Object.isFrozen(frag.rows[0])).toBe(true);
    expect(Object.isFrozen(frag.rows[0].cells)).toBe(true);
    expect(Object.isFrozen(frag.rows[0].cells[0])).toBe(true);
    expect(Object.isFrozen(frag.columnWidthsPt)).toBe(true);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// layoutDocument — table fragments on body pages (real measurement)
// ─────────────────────────────────────────────────────────────────────────────

describe('layoutDocument — table fragments', () => {
  it('emits a placed table fragment from a clone that preserves parsed row identity', () => {
    const t = table([row([textCell('a'), textCell('b')])], [80, 80]);
    const tables = allTables(doc([t as unknown as BodyElement]));
    expect(tables).toHaveLength(1);
    expect(tables[0].table.source).not.toBe(t);
    expect(tables[0].table.source.rows).toBe(t.rows);
    expect(tables[0].table.rows).toHaveLength(1);
    expect(tables[0].table.continuesFromPreviousPage).toBe(false);
    expect(tables[0].table.continuesOnNextPage).toBe(false);
    // column widths sum to the table width the paginator resolved (≤ content band)
    const sum = tables[0].table.columnWidthsPt.reduce((s, w) => s + w, 0);
    expect(sum).toBeGreaterThan(0);
    expect(sum).toBeLessThanOrEqual(180 + 1e-6);
  });

  it('keeps fragment state OFF the parsed table (no fragment fields on DocTable)', () => {
    const t = table([row([textCell('a')])], [120]);
    const keysBeforePagination = Object.keys(t);
    allTables(doc([t as unknown as BodyElement]));
    expect(Object.keys(t)).toEqual(keysBeforePagination);
    expect('kind' in (t as unknown as Record<string, unknown>)).toBe(false);
    expect('columnWidthsPt' in (t as unknown as Record<string, unknown>)).toBe(false);
  });

  it('keeps table fragment side-table entries isolated across pagination runs', () => {
    const t = table([row([textCell('a')])], [200]);
    const body = [t as unknown as BodyElement];
    const wideSection = doc(body).section;
    const narrowSection = { ...wideSection, pageWidth: 150 };

    const widePages = computePages(body, wideSection, makeStubCtx());
    const wideElement = widePages.flat().find((element) => element.type === 'table');
    expect(wideElement).toBeDefined();

    const narrowPages = computePages(body, narrowSection, makeStubCtx());
    const narrowElement = narrowPages.flat().find((element) => element.type === 'table');
    expect(narrowElement).toBeDefined();

    expect(bodyFragmentFor(wideElement as NonNullable<typeof wideElement>)?.widthPt)
      .toBeCloseTo(180, 6);
    expect(bodyFragmentFor(narrowElement as NonNullable<typeof narrowElement>)?.widthPt)
      .toBeCloseTo(130, 6);
    expect(wideElement).not.toBe(narrowElement);
  });

  it('builds each cell paragraph as a ParagraphFragment referencing the source cell paragraph', () => {
    const cellPara = para('hello');
    const c = cell([{ type: 'paragraph', ...cellPara } as unknown as CellElement]);
    const t = table([row([c])], [120]);
    const { table: tf } = allTables(doc([t as unknown as BodyElement]))[0];
    const block = firstParagraphBlock(tf.rows[0].cells[0]);
    // the block's source is the cell's paragraph content element (same runs)
    expect(block.kind).toBe('paragraph');
    expect(block.source.runs[0]).toBeDefined();
    expect((block.source.runs[0] as { text: string }).text).toBe('hello');
  });

  it('splits a tall table across pages into continuation fragments (auto rows)', () => {
    const rows = Array.from({ length: 12 }, (_v, i) => row([textCell(`row ${i}`)]));
    const t = table(rows, [120]);
    // page height 120 with 10pt margins → ~100pt body; each auto row ~ MIN 10pt+
    const tables = allTables(doc([t as unknown as BodyElement], 120));
    expect(tables.length).toBeGreaterThanOrEqual(2);
    expect(tables[0].table.continuesOnNextPage).toBe(true);
    expect(tables[0].table.continuesFromPreviousPage).toBe(false);
    const last = tables[tables.length - 1].table;
    expect(last.continuesOnNextPage).toBe(false);
    expect(last.continuesFromPreviousPage).toBe(true);
    // every source row appears exactly once across the slices, in order
    const seen = tables.flatMap((x) => x.table.rows.map((r) => r.sourceRowIndex));
    expect(seen).toEqual(rows.map((_v, i) => i));
  });

  it('keeps original row indices when a row taller than the page is split into pieces', () => {
    const tallCell = cell([
      { type: 'paragraph', ...para('あ'.repeat(400)) } as unknown as CellElement,
    ]);
    const t = table(
      [row([tallCell]), row([textCell('following row')])],
      [120],
    );
    const tables = allTables(doc([t as unknown as BodyElement], 120));
    const fragmentRows = tables.flatMap(({ table: fragment }) => fragment.rows);

    // Sanity: splitRowsTallerThanPage expanded the first parsed row into pieces.
    expect(fragmentRows.length).toBeGreaterThan(t.rows.length);
    expect(fragmentRows.at(-1)?.source).toBe(t.rows[1]);
    expect(fragmentRows.slice(0, -1).map((fragment) => fragment.sourceRowIndex))
      .toEqual(Array.from({ length: fragmentRows.length - 1 }, () => 0));
    expect(fragmentRows.at(-1)?.sourceRowIndex).toBe(1);
  });

  it('repeats a leading header row on every continuation slice (§17.4.78)', () => {
    const bodyRows = Array.from({ length: 12 }, (_v, i) => row([textCell(`b${i}`)]));
    const rows = [row([textCell('HEADER')], { isHeader: true }), ...bodyRows];
    const t = table(rows, [120]);
    const tables = allTables(doc([t as unknown as BodyElement], 120));
    expect(tables.length).toBeGreaterThanOrEqual(2);
    // continuation slices lead with a repeated-header RowFragment referencing the source header row
    for (let i = 1; i < tables.length; i++) {
      const head = tables[i].table.rows[0];
      expect(head.repeatedHeader).toBe(true);
      expect(head.sourceRowIndex).toBe(0);
    }
    // the first slice's header is NOT marked as repeated
    expect(tables[0].table.rows[0].repeatedHeader).toBe(false);
  });

  it('fragments a nested table as a TableFragment cell block', () => {
    const inner = table([row([textCell('inner')])], [80]);
    const outerCell = cell([{ type: 'table', ...inner } as unknown as CellElement]);
    const t = table([row([outerCell])], [120]);
    const { table: tf } = allTables(doc([t as unknown as BodyElement]))[0];
    const block = tf.rows[0].cells[0].blocks[0];
    expect(block).toBeDefined();
    expect(block.kind).toBe('table');
    if (block.kind === 'table') {
      expect(block.rows).toHaveLength(1);
      expect(firstParagraphBlock(block.rows[0].cells[0]).source.runs[0]).toBeDefined();
    }
  });

  it('marks continuation on nested table fragments split at inner row boundaries', () => {
    const inner = table(
      Array.from({ length: 4 }, (_unused, index) =>
        row([textCell(`inner ${index}`)], { rowHeight: 30, rowHeightRule: 'exact' }),
      ),
      [80],
    );
    const outerCell = cell([{ type: 'table', ...inner } as unknown as CellElement]);
    const outer = table([row([outerCell])], [120]);
    const outerFragments = allTables(doc([outer as unknown as BodyElement], 120));
    const nestedFragments: TableFragment[] = [];
    for (const { table: outerFragment } of outerFragments) {
      for (const outerRow of outerFragment.rows) {
        for (const cellFragment of outerRow.cells) {
          for (const block of cellFragment.blocks) {
            if (block.kind === 'table') nestedFragments.push(block);
          }
        }
      }
    }

    expect(nestedFragments.length).toBeGreaterThan(1);
    expect(nestedFragments[0].continuesFromPreviousPage).toBe(false);
    expect(nestedFragments[0].continuesOnNextPage).toBe(true);
    const last = nestedFragments.at(-1) as TableFragment;
    expect(last.continuesFromPreviousPage).toBe(true);
    expect(last.continuesOnNextPage).toBe(false);
  });

  it('carries vertical-merge roles across the merged span', () => {
    const t = table(
      [
        row([textCell('m', { vMerge: true }), textCell('a')]),
        row([textCell('', { vMerge: false }), textCell('b')]),
      ],
      [60, 60],
    );
    const { table: tf } = allTables(doc([t as unknown as BodyElement]))[0];
    expect(tf.rows[0].cells[0].verticalMerge).toBe('restart');
    expect(tf.rows[1].cells[0].verticalMerge).toBe('continue');
    expect(tf.rows[1].cells[0].blocks).toEqual([]);
  });

  it('INVARIANT: placed.heightPt equals the summed row heights of the fragment', () => {
    const rows = Array.from({ length: 3 }, (_v, i) => row([textCell(`r${i}`)]));
    const t = table(rows, [120]);
    for (const { placed, table: tf } of allTables(doc([t as unknown as BodyElement]))) {
      expect(placed.heightPt).toBeCloseTo(tableFragmentHeightPt(tf), 6);
    }
  });
});
