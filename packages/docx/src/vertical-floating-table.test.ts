import { describe, expect, it } from 'vitest';
import {
  bodyFragmentFor,
  computePages,
  renderDocumentToCanvas,
  __test_verticalLayoutSection,
  type DocxTextRunInfo,
} from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  SectionProps,
} from './types.js';
import type { TableFragmentLayout } from './layout/table-pagination.js';

type PaintEvent =
  | Readonly<{ kind: 'text'; text: string }>
  | Readonly<{ kind: 'stroke'; color: string }>;

function recordingCanvas(): {
  canvas: HTMLCanvasElement;
  paintEvents: PaintEvent[];
  measureCalls: () => number;
} {
  let font = '10px serif';
  let measures = 0;
  const paintEvents: PaintEvent[] = [];
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText(text: string) {
      measures += 1;
      const px = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...text].length * px,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {
      paintEvents.push({ kind: 'stroke', color: String(ctx.strokeStyle) });
    }, fill() {}, fillRect() {}, strokeRect() {},
    rect() {}, clip() {}, scale() {}, translate() {}, rotate() {}, setTransform() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign,
    direction: 'ltr' as CanvasDirection,
    globalAlpha: 1,
    lineCap: 'butt' as CanvasLineCap,
    lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0,
    height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return {
    canvas: canvas as unknown as HTMLCanvasElement,
    paintEvents,
    measureCalls: () => measures,
  };
}

function emptyBorders() {
  return { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };
}

function paragraph(text: string): DocParagraph {
  return {
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text,
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null,
      fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 10,
    defaultFontFamily: 'Times New Roman',
    widowControl: false,
  } as unknown as DocParagraph;
}

function floatingTable(rows: readonly Readonly<{ text: string; heightPt: number }>[]): DocTable {
  return {
    colWidths: [60],
    rows: rows.map(({ text, heightPt }): DocTableRow => ({
      cells: [{
        content: [{ type: 'paragraph', ...paragraph(text) }],
        colSpan: 1,
        vMerge: null,
        borders: emptyBorders(),
        background: null,
        vAlign: 'top',
        widthPt: 60,
        marginTop: 0, marginBottom: 0, marginLeft: 0, marginRight: 0,
      } as unknown as DocTableCell],
      rowHeight: heightPt,
      rowHeightRule: 'exact',
      isHeader: false,
    } as unknown as DocTableRow)),
    borders: emptyBorders(),
    cellMarginTop: 0,
    cellMarginBottom: 0,
    cellMarginLeft: 0,
    cellMarginRight: 0,
    jc: 'left',
    tblpPr: {
      leftFromText: 0,
      rightFromText: 0,
      topFromText: 0,
      bottomFromText: 0,
      horzAnchor: 'text',
      horzSpecified: true,
      vertAnchor: 'text',
      tblpX: 0,
      tblpY: 0,
    },
    overlap: 'never',
  } as unknown as DocTable;
}

const PHYSICAL_SECTION = {
  pageWidth: 200,
  pageHeight: 300,
  marginTop: 20,
  marginRight: 30,
  marginBottom: 40,
  marginLeft: 24,
  headerDistance: 0,
  footerDistance: 0,
  titlePage: false,
  evenAndOddHeaders: false,
  textDirection: 'tbRl',
} as unknown as SectionProps;

function documentWith(table: DocTable): DocxDocumentModel {
  return {
    section: PHYSICAL_SECTION,
    body: [{ type: 'table', ...table } as unknown as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

describe('vertical outer floating tables retain the canonical logical paint domain', () => {
  it('paints a fitting outer tblpPr table once without physical-domain finalization', async () => {
    const doc = documentWith(floatingTable([{ text: 'OUTER-FLOAT', heightPt: 30 }]));
    const measure = recordingCanvas();
    const pages = computePages(
      doc.body,
      __test_verticalLayoutSection(PHYSICAL_SECTION),
      measure.canvas.getContext('2d') as CanvasRenderingContext2D,
      doc.fontFamilyClasses,
    );
    const table = pages[0]?.find((element) => element.type === 'table');
    const retained = table ? bodyFragmentFor(table) : undefined;
    const coordinateSpace = retained?.fragment.kind === 'table'
      && 'floatingTableCoordinateSpace' in retained.fragment
      ? (retained.fragment as TableFragmentLayout).floatingTableCoordinateSpace
      : undefined;

    expect(pages).toHaveLength(1);

    const paint = recordingCanvas();
    const runs: DocxTextRunInfo[] = [];
    await expect(renderDocumentToCanvas(doc, paint.canvas, 0, {
      dpr: 1,
      width: PHYSICAL_SECTION.pageWidth,
      prebuiltPages: pages,
      onTextRun: (run) => runs.push(run),
    })).resolves.toBeUndefined();
    expect(coordinateSpace).not.toBe('upright-physical-page-points');
    expect(paint.measureCalls()).toBe(0);
    expect(runs.filter((run) => run.text === 'OUTER-FLOAT')).toHaveLength(1);
  });

  it('finalizes a fitting nested float before registering and painting its outer tblpPr parent', async () => {
    const outer = floatingTable([{ text: 'unused', heightPt: 90 }]);
    const parentBorder = { style: 'single', width: 1, color: 'ff00ff' };
    outer.borders = {
      top: parentBorder,
      bottom: parentBorder,
      left: parentBorder,
      right: parentBorder,
      insideH: null,
      insideV: null,
    };
    const nestedTable = floatingTable([{ text: 'FITTING-NESTED', heightPt: 20 }]);
    nestedTable.tblpPr = {
      ...nestedTable.tblpPr!,
      horzAnchor: 'page',
      vertAnchor: 'page',
      tblpX: 20,
      tblpY: 30,
    };
    outer.rows[0]!.cells[0]!.content = [
      { type: 'table', ...nestedTable },
      { type: 'paragraph', ...paragraph('FITTING-ANCHOR') },
    ] as unknown as DocTableCell['content'];
    const doc = documentWith(outer);
    const measure = recordingCanvas();
    const pages = computePages(
      doc.body,
      __test_verticalLayoutSection(PHYSICAL_SECTION),
      measure.canvas.getContext('2d') as CanvasRenderingContext2D,
      doc.fontFamilyClasses,
    );
    const table = pages[0]?.find((element) => element.type === 'table');
    const retained = table ? bodyFragmentFor(table) : undefined;
    if (retained?.fragment.kind !== 'table'
      || !('resolvedFloatingTables' in retained.fragment)) {
      throw new Error('expected a fitting logical outer fragment with a selected nested float');
    }
    const fragment = retained.fragment as TableFragmentLayout;
    const [nested] = fragment.resolvedFloatingTables;
    if (!nested) throw new Error('expected a retained fitting nested float');
    const anchor = fragment.rows[0]?.cells[0]?.blocks.find(
      (block) => block.layout.kind === 'paragraph',
    );
    const expectedLocalExclusion = {
      xPt: nested.exclusionBounds.xPt - nested.source.anchorBounds.xPt,
      yPt: nested.exclusionBounds.yPt - nested.source.anchorBounds.yPt,
      widthPt: nested.exclusionBounds.widthPt,
      heightPt: nested.exclusionBounds.heightPt,
    };

    expect(pages).toHaveLength(1);
    expect(fragment.floatingTableCoordinateSpace).toBe('logical-page-points');
    expect(fragment.resolvedFloatingTables).toHaveLength(1);
    expect(nested.bounds).toEqual({
      xPt: 20,
      yPt: 30,
      widthPt: nested.child.columnWidthsPt.reduce((sum, width) => sum + width, 0),
      heightPt: 20,
    });
    expect(anchor?.layout.kind === 'paragraph' ? anchor.layout.exclusions : [])
      .toContainEqual(expect.objectContaining({ bounds: expectedLocalExclusion }));

    const paint = recordingCanvas();
    const runs: DocxTextRunInfo[] = [];
    await expect(renderDocumentToCanvas(doc, paint.canvas, 0, {
      dpr: 1,
      width: PHYSICAL_SECTION.pageWidth,
      prebuiltPages: pages,
      onTextRun: (run) => {
        runs.push(run);
        paint.paintEvents.push({ kind: 'text', text: run.text });
      },
    })).resolves.toBeUndefined();
    expect(paint.measureCalls()).toBe(0);
    expect(runs.filter((run) => run.text === 'FITTING-NESTED')).toHaveLength(1);
    const childPaint = paint.paintEvents.findIndex(
      (event) => event.kind === 'text' && event.text === 'FITTING-NESTED',
    );
    const parentBorderPaint = paint.paintEvents.findIndex(
      (event) => event.kind === 'stroke' && event.color === '#ff00ff',
    );
    expect(childPaint).toBeGreaterThanOrEqual(0);
    expect(parentBorderPaint).toBeGreaterThan(childPaint);
  });

  it('keeps every split outer tblpPr slice in the logical domain and paints each row once', async () => {
    const doc = documentWith(floatingTable([
      { text: 'OUTER-SLICE-1', heightPt: 90 },
      { text: 'OUTER-SLICE-2', heightPt: 90 },
    ]));
    const measure = recordingCanvas();
    const pages = computePages(
      doc.body,
      __test_verticalLayoutSection(PHYSICAL_SECTION),
      measure.canvas.getContext('2d') as CanvasRenderingContext2D,
      doc.fontFamilyClasses,
    );
    const retainedSlices = pages.map((page) => {
      const table = page.find((element) => element.type === 'table');
      const retained = table ? bodyFragmentFor(table) : undefined;
      if (retained?.fragment.kind !== 'table'
        || !('floatingTableCoordinateSpace' in retained.fragment)) {
        throw new Error('expected a retained split outer float slice');
      }
      return retained.fragment as TableFragmentLayout;
    });

    expect(pages).toHaveLength(2);
    expect(retainedSlices.map((fragment) => fragment.floatingTableCoordinateSpace))
      .toEqual(['logical-page-points', 'logical-page-points']);

    const runs: DocxTextRunInfo[] = [];
    for (let pageIndex = 0; pageIndex < pages.length; pageIndex += 1) {
      const paint = recordingCanvas();
      await expect(renderDocumentToCanvas(doc, paint.canvas, pageIndex, {
        dpr: 1,
        width: PHYSICAL_SECTION.pageWidth,
        prebuiltPages: pages,
        onTextRun: (run) => runs.push(run),
      })).resolves.toBeUndefined();
      expect(paint.measureCalls()).toBe(0);
    }
    expect(runs.filter((run) => run.text === 'OUTER-SLICE-1')).toHaveLength(1);
    expect(runs.filter((run) => run.text === 'OUTER-SLICE-2')).toHaveLength(1);
  });

  it('commits a selected nested float before registering its split outer slice', async () => {
    const outer = floatingTable([
      { text: 'unused', heightPt: 90 },
      { text: 'OUTER-TAIL', heightPt: 90 },
    ]);
    const nestedTable = floatingTable([{ text: 'NESTED-FLOAT', heightPt: 20 }]);
    nestedTable.tblpPr = {
      ...nestedTable.tblpPr!,
      horzAnchor: 'page',
      vertAnchor: 'page',
      tblpX: 20,
      tblpY: 30,
    };
    outer.rows[0]!.cells[0]!.content = [
      { type: 'table', ...nestedTable },
      { type: 'paragraph', ...paragraph('NESTED-ANCHOR') },
    ] as unknown as DocTableCell['content'];
    const doc = documentWith(outer);
    const measure = recordingCanvas();
    const pages = computePages(
      doc.body,
      __test_verticalLayoutSection(PHYSICAL_SECTION),
      measure.canvas.getContext('2d') as CanvasRenderingContext2D,
      doc.fontFamilyClasses,
    );
    const firstTable = pages[0]?.find((element) => element.type === 'table');
    const retained = firstTable ? bodyFragmentFor(firstTable) : undefined;
    if (retained?.fragment.kind !== 'table'
      || !('resolvedFloatingTables' in retained.fragment)) {
      throw new Error('expected the first logical outer slice with a selected nested float');
    }
    const fragment = retained.fragment as TableFragmentLayout;
    const [nested] = fragment.resolvedFloatingTables;

    expect(pages).toHaveLength(2);
    expect(fragment.floatingTableCoordinateSpace).toBe('logical-page-points');
    expect(fragment.resolvedFloatingTables).toHaveLength(1);
    expect(nested?.bounds).toEqual({
      xPt: 20,
      yPt: 30,
      widthPt: nested.child.columnWidthsPt.reduce((sum, width) => sum + width, 0),
      heightPt: 20,
    });

    const runs: DocxTextRunInfo[] = [];
    for (let pageIndex = 0; pageIndex < pages.length; pageIndex += 1) {
      const paint = recordingCanvas();
      await expect(renderDocumentToCanvas(doc, paint.canvas, pageIndex, {
        dpr: 1,
        width: PHYSICAL_SECTION.pageWidth,
        prebuiltPages: pages,
        onTextRun: (run) => runs.push(run),
      })).resolves.toBeUndefined();
      expect(paint.measureCalls()).toBe(0);
    }
    expect(runs.filter((run) => run.text === 'NESTED-FLOAT')).toHaveLength(1);
  });
});
