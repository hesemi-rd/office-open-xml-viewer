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

function freshPageSensitiveSplitDocument(): DocxDocumentModel {
  const blocker = floatingTable([{ text: 'B', heightPt: 20 }]);
  blocker.layout = 'fixed';
  blocker.colWidths = [40];
  blocker.rows[0]!.cells[0]!.widthPt = 40;
  blocker.tblpPr = {
    ...blocker.tblpPr!,
    horzAnchor: 'page',
    vertAnchor: 'page',
    tblpX: 40,
    tblpY: 110,
  };
  const outer = floatingTable([
    { text: 'unused', heightPt: 20 },
    { text: 'TAIL', heightPt: 20 },
  ]);
  outer.rows[0]!.rowHeight = null;
  outer.rows[0]!.rowHeightRule = 'auto';
  const nested = floatingTable([{ text: 'N', heightPt: 70 }]);
  nested.layout = 'fixed';
  nested.colWidths = [60];
  nested.rows[0]!.cells[0]!.widthPt = 60;
  nested.tblpPr = {
    ...nested.tblpPr!,
    horzAnchor: 'page',
    vertAnchor: 'text',
    tblpX: 100,
    tblpY: 0,
  };
  outer.rows[0]!.cells[0]!.content = [
    { type: 'table', ...nested },
    { type: 'paragraph', ...paragraph('A') },
  ] as unknown as DocTableCell['content'];
  const doc = documentWith(blocker);
  doc.body.push({ type: 'table', ...outer } as unknown as BodyElement);
  return doc;
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
    outer.tblpPr = {
      ...outer.tblpPr!,
      leftFromText: 3,
      rightFromText: 5,
      topFromText: 7,
      bottomFromText: 11,
    };
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
    const external = floatingTable([{ text: 'AFTER-EXTERNAL', heightPt: 20 }]);
    doc.body.push({ type: 'table', ...external } as unknown as BodyElement);
    doc.body.push({ type: 'paragraph', ...paragraph('AFTER-FITTING') } as unknown as BodyElement);
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
    const followingElement = pages[0]?.find(
      (element) => element.type === 'paragraph' && element.runs?.some(
        (run) => run.type === 'text' && run.text === 'AFTER-FITTING',
      ),
    );
    const following = followingElement ? bodyFragmentFor(followingElement) : undefined;
    if (following?.fragment.kind !== 'paragraph') {
      throw new Error('expected a following fitting-parent paragraph fragment');
    }
    const expectedParentExclusion = {
      xPt: retained.xPt - outer.tblpPr.leftFromText,
      yPt: retained.yPt - outer.tblpPr.topFromText,
      widthPt: retained.widthPt + outer.tblpPr.leftFromText + outer.tblpPr.rightFromText,
      heightPt: retained.heightPt + outer.tblpPr.topFromText + outer.tblpPr.bottomFromText,
    };
    const externalElement = pages[0]?.filter((element) => element.type === 'table')[1];
    const retainedExternal = externalElement ? bodyFragmentFor(externalElement) : undefined;
    if (retainedExternal?.fragment.kind !== 'table') {
      throw new Error('expected a following external floating table fragment');
    }

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
    expect(following.fragment.exclusions.map((exclusion) => exclusion.id)).toEqual([
      nested.occurrenceId,
      'body:float:1',
      'body:float:2',
    ]);
    expect(following.fragment.exclusions[1]?.bounds).toEqual(expectedParentExclusion);
    expect(retainedExternal.yPt).toBeGreaterThanOrEqual(
      expectedParentExclusion.yPt + expectedParentExclusion.heightPt,
    );
    expect(following.fragment.exclusions[2]?.bounds).toEqual({
      xPt: retainedExternal.xPt,
      yPt: retainedExternal.yPt,
      widthPt: retainedExternal.widthPt,
      heightPt: retainedExternal.heightPt,
    });

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

  it('derives an auto-height parent exclusion from the child-reflowed fitting fragment', () => {
    const outer = floatingTable([{ text: 'unused', heightPt: 20 }]);
    outer.layout = 'fixed';
    outer.rows[0]!.rowHeight = null;
    outer.rows[0]!.rowHeightRule = 'auto';
    const nestedTable = floatingTable([{ text: 'N', heightPt: 20 }]);
    nestedTable.layout = 'fixed';
    nestedTable.colWidths = [40];
    nestedTable.rows[0]!.cells[0]!.widthPt = 40;
    nestedTable.tblpPr = {
      ...nestedTable.tblpPr!,
      horzAnchor: 'page',
      vertAnchor: 'page',
      tblpX: 20,
      tblpY: 30,
    };
    outer.rows[0]!.cells[0]!.content = [
      { type: 'table', ...nestedTable },
      { type: 'paragraph', ...paragraph('x'.repeat(20)) },
    ] as unknown as DocTableCell['content'];
    const doc = documentWith(outer);
    doc.body.push({ type: 'paragraph', ...paragraph('FOLLOWING') } as unknown as BodyElement);
    const measure = recordingCanvas();
    const pages = computePages(
      doc.body,
      __test_verticalLayoutSection(PHYSICAL_SECTION),
      measure.canvas.getContext('2d') as CanvasRenderingContext2D,
      doc.fontFamilyClasses,
    );
    const tableElement = pages.flat().find((element) => element.type === 'table');
    const followingElement = pages.flat().find((element) => element.type === 'paragraph');
    const retainedTable = tableElement ? bodyFragmentFor(tableElement) : undefined;
    const retainedFollowing = followingElement ? bodyFragmentFor(followingElement) : undefined;
    if (retainedTable?.fragment.kind !== 'table'
      || !('resolvedFloatingTables' in retainedTable.fragment)
      || retainedFollowing?.fragment.kind !== 'paragraph') {
      throw new Error('expected retained auto-height outer table and following paragraph');
    }
    const parentExclusion = retainedFollowing.fragment.exclusions.find(
      (exclusion) => Math.abs(exclusion.bounds.widthPt - retainedTable.widthPt) < 0.001,
    );
    if (!parentExclusion) throw new Error('expected the registered parent exclusion');

    expect(retainedTable.fragment.resolvedFloatingTables).toHaveLength(1);
    expect(parentExclusion.bounds.heightPt).toBeCloseTo(retainedTable.fragment.advancePt, 6);
  });

  it('finalizes a deferred fitting outer float with its destination page identity', async () => {
    const blocker = floatingTable([{ text: 'BLOCKER', heightPt: 30 }]);
    blocker.tblpPr = {
      ...blocker.tblpPr!,
      horzAnchor: 'page',
      vertAnchor: 'page',
      tblpX: 20,
      tblpY: 30,
    };
    const relocated = floatingTable([{ text: 'unused', heightPt: 70 }]);
    relocated.tblpPr = { ...blocker.tblpPr! };
    const nestedTable = floatingTable([{ text: 'LATER-NESTED', heightPt: 20 }]);
    nestedTable.tblpPr = {
      ...nestedTable.tblpPr!,
      horzAnchor: 'page',
      vertAnchor: 'page',
      tblpX: 100,
      tblpY: 30,
    };
    const pageAnchor = paragraph('');
    (pageAnchor.runs as unknown[]).push({
      type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '?',
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', background: null,
    });
    relocated.rows[0]!.cells[0]!.content = [
      { type: 'table', ...nestedTable },
      { type: 'paragraph', ...pageAnchor },
    ] as unknown as DocTableCell['content'];
    const doc = documentWith(blocker);
    doc.body.push({ type: 'table', ...relocated } as unknown as BodyElement);
    const measure = recordingCanvas();
    const pages = computePages(
      doc.body,
      __test_verticalLayoutSection(PHYSICAL_SECTION),
      measure.canvas.getContext('2d') as CanvasRenderingContext2D,
      doc.fontFamilyClasses,
    );
    const destinationTable = pages[1]?.find((element) => element.type === 'table');
    const retained = destinationTable ? bodyFragmentFor(destinationTable) : undefined;
    if (retained?.fragment.kind !== 'table'
      || !('resolvedFloatingTables' in retained.fragment)) {
      throw new Error('expected the deferred fitting outer fragment');
    }
    const fragment = retained.fragment as TableFragmentLayout;
    const [nested] = fragment.resolvedFloatingTables;

    expect(pages).toHaveLength(2);
    expect(fragment.floatingTableCoordinateSpace).toBe('logical-page-points');
    expect(nested?.source.physicalPageIndex).toBe(1);
    expect(nested?.source.displayPageNumber).toBe(2);
    expect(nested?.occurrenceId).toContain(':fitting-outer:1:');
    expect(nested?.bounds).toEqual({
      xPt: 100,
      yPt: 30,
      widthPt: nested.child.columnWidthsPt.reduce((sum, width) => sum + width, 0),
      heightPt: 20,
    });

    const paint = recordingCanvas();
    const runs: DocxTextRunInfo[] = [];
    await renderDocumentToCanvas(doc, paint.canvas, 1, {
      dpr: 1,
      width: PHYSICAL_SECTION.pageWidth,
      prebuiltPages: pages,
      onTextRun: (run) => runs.push(run),
    });
    expect(paint.measureCalls()).toBe(0);
    expect(runs.some((run) => run.text === '2')).toBe(true);
    expect(runs.filter((run) => run.text === 'LATER-NESTED')).toHaveLength(1);
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

  it('stamps later-page split slices with their live page ownership without a PAGE map', () => {
    const doc = documentWith(floatingTable([
      { text: 'LATER-SLICE-1', heightPt: 90 },
      { text: 'LATER-SLICE-2', heightPt: 90 },
    ]));
    doc.body.unshift(
      { type: 'pageBreak' } as BodyElement,
      { type: 'pageBreak' } as BodyElement,
    );
    const measure = recordingCanvas();
    const pages = computePages(
      doc.body,
      __test_verticalLayoutSection(PHYSICAL_SECTION),
      measure.canvas.getContext('2d') as CanvasRenderingContext2D,
      doc.fontFamilyClasses,
    );
    const retainedSlices = pages.slice(2).map((page) => {
      const table = page.find((element) => element.type === 'table');
      const retained = table ? bodyFragmentFor(table) : undefined;
      if (retained?.fragment.kind !== 'table') {
        throw new Error('expected a retained later-page split float slice');
      }
      return retained.fragment as TableFragmentLayout;
    });

    expect(pages).toHaveLength(4);
    expect(retainedSlices.map((fragment) => fragment.rows[0]?.physicalPageIndex))
      .toEqual([2, 3]);
    expect(retainedSlices.map((fragment) => fragment.rows[0]?.displayPageNumber))
      .toEqual([3, 4]);
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
    outer.tblpPr = {
      ...outer.tblpPr!,
      leftFromText: 3,
      rightFromText: 5,
      topFromText: 7,
      bottomFromText: 11,
    };
    outer.rows[1]!.cells[0]!.content = [
      { type: 'table', ...nestedTable },
      { type: 'paragraph', ...paragraph('NESTED-ANCHOR') },
    ] as unknown as DocTableCell['content'];
    const doc = documentWith(outer);
    doc.body.push({ type: 'paragraph', ...paragraph('AFTER-SPLIT') } as unknown as BodyElement);
    const measure = recordingCanvas();
    const pages = computePages(
      doc.body,
      __test_verticalLayoutSection(PHYSICAL_SECTION),
      measure.canvas.getContext('2d') as CanvasRenderingContext2D,
      doc.fontFamilyClasses,
    );
    const finalTable = pages[1]?.find((element) => element.type === 'table');
    const retained = finalTable ? bodyFragmentFor(finalTable) : undefined;
    if (retained?.fragment.kind !== 'table'
      || !('resolvedFloatingTables' in retained.fragment)) {
      throw new Error('expected the final logical outer slice with a selected nested float');
    }
    const fragment = retained.fragment as TableFragmentLayout;
    const [nested] = fragment.resolvedFloatingTables;
    if (!nested) throw new Error('expected a retained split nested float');
    const followingElement = pages[1]?.find(
      (element) => element.type === 'paragraph' && element.runs?.some(
        (run) => run.type === 'text' && run.text === 'AFTER-SPLIT',
      ),
    );
    const following = followingElement ? bodyFragmentFor(followingElement) : undefined;
    if (following?.fragment.kind !== 'paragraph') {
      throw new Error('expected a following split-parent paragraph fragment');
    }
    const expectedParentExclusion = {
      xPt: retained.xPt - outer.tblpPr.leftFromText,
      yPt: retained.yPt - outer.tblpPr.topFromText,
      widthPt: retained.widthPt + outer.tblpPr.leftFromText + outer.tblpPr.rightFromText,
      heightPt: retained.heightPt + outer.tblpPr.topFromText + outer.tblpPr.bottomFromText,
    };

    expect(pages).toHaveLength(2);
    expect(fragment.floatingTableCoordinateSpace).toBe('logical-page-points');
    expect(fragment.resolvedFloatingTables).toHaveLength(1);
    expect(nested?.bounds).toEqual({
      xPt: 20,
      yPt: 30,
      widthPt: nested.child.columnWidthsPt.reduce((sum, width) => sum + width, 0),
      heightPt: 20,
    });
    expect(following.fragment.exclusions.map((exclusion) => exclusion.id)).toEqual([
      nested.occurrenceId,
      'body:float:1',
    ]);
    expect(following.fragment.exclusions[1]?.bounds).toEqual(expectedParentExclusion);

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

  it('reflows a split child against the accepted slice box after an external blocker', async () => {
    const blocker = floatingTable([{ text: 'B', heightPt: 20 }]);
    blocker.layout = 'fixed';
    blocker.colWidths = [40];
    blocker.rows[0]!.cells[0]!.widthPt = 40;
    blocker.tblpPr = {
      ...blocker.tblpPr!,
      horzAnchor: 'page',
      vertAnchor: 'page',
      tblpX: 40,
      tblpY: 170,
    };

    const outer = floatingTable([
      { text: 'unused', heightPt: 140 },
      { text: 'OUTER-TAIL', heightPt: 20 },
    ]);
    const nestedTable = floatingTable([{ text: 'C', heightPt: 20 }]);
    nestedTable.layout = 'fixed';
    nestedTable.colWidths = [20];
    nestedTable.rows[0]!.cells[0]!.widthPt = 20;
    nestedTable.tblpPr = {
      ...nestedTable.tblpPr!,
      horzAnchor: 'text',
      vertAnchor: 'page',
      tblpX: 0,
      tblpY: 0,
    };
    outer.rows[0]!.cells[0]!.content = [
      { type: 'table', ...nestedTable },
      { type: 'paragraph', ...paragraph('CHILD-ANCHOR') },
    ] as unknown as DocTableCell['content'];

    const doc = documentWith(blocker);
    doc.body.push({ type: 'table', ...outer } as unknown as BodyElement);
    doc.body.push({ type: 'paragraph', ...paragraph('AFTER-BLOCKED-SPLIT') } as unknown as BodyElement);
    const measure = recordingCanvas();
    const pages = computePages(
      doc.body,
      __test_verticalLayoutSection(PHYSICAL_SECTION),
      measure.canvas.getContext('2d') as CanvasRenderingContext2D,
      doc.fontFamilyClasses,
    );
    const firstOuterElement = pages[0]?.filter((element) => element.type === 'table')[1];
    const blockerElement = pages[0]?.filter((element) => element.type === 'table')[0];
    const retainedOuter = firstOuterElement ? bodyFragmentFor(firstOuterElement) : undefined;
    const retainedBlocker = blockerElement ? bodyFragmentFor(blockerElement) : undefined;
    if (retainedOuter?.fragment.kind !== 'table'
      || !('resolvedFloatingTables' in retainedOuter.fragment)
      || retainedBlocker?.fragment.kind !== 'table') {
      throw new Error('expected a first split slice with a selected text-anchored child');
    }
    const fragment = retainedOuter.fragment as TableFragmentLayout;
    const [nested] = fragment.resolvedFloatingTables;
    if (!nested) throw new Error('expected the selected text-anchored child');
    expect(pages).toHaveLength(2);
    expect(retainedOuter.xPt).toBeCloseTo(20, 6);
    expect(retainedOuter.yPt).toBeCloseTo(30, 6);
    expect(retainedOuter.heightPt).toBeCloseTo(140, 6);
    expect(retainedOuter.yPt + retainedOuter.heightPt).toBeCloseTo(retainedBlocker.yPt, 6);
    expect(retainedOuter.yPt + 160).toBeGreaterThan(retainedBlocker.yPt);
    expect(nested.bounds.xPt).toBeCloseTo(retainedOuter.xPt, 6);
    expect(nested.source.anchorBounds.xPt).toBeCloseTo(retainedOuter.xPt, 6);

    const paint = recordingCanvas();
    const runs: DocxTextRunInfo[] = [];
    await expect(renderDocumentToCanvas(doc, paint.canvas, 0, {
      dpr: 1,
      width: PHYSICAL_SECTION.pageWidth,
      prebuiltPages: pages,
      onTextRun: (run) => runs.push(run),
    })).resolves.toBeUndefined();
    expect(paint.measureCalls()).toBe(0);
    expect(runs.filter((run) => run.text === 'C')).toHaveLength(1);
  });

  it('keeps a provisionally deferred float on the current page after slice reseating', () => {
    const doc = freshPageSensitiveSplitDocument();
    const measure = recordingCanvas();
    const pages = computePages(
      doc.body,
      __test_verticalLayoutSection(PHYSICAL_SECTION),
      measure.canvas.getContext('2d') as CanvasRenderingContext2D,
      doc.fontFamilyClasses,
    );
    const outerFragments = pages.flatMap((page, pageIndex) => page.flatMap((element) => {
      if (element.type !== 'table') return [];
      const retained = bodyFragmentFor(element);
      if (retained?.fragment.kind !== 'table' || retained.fragment.source.path[0] !== 1) return [];
      return [{ pageIndex, fragment: retained.fragment as TableFragmentLayout }];
    }));

    expect(pages).toHaveLength(1);
    expect(outerFragments).toHaveLength(1);
    expect(outerFragments[0]?.pageIndex).toBe(0);
    expect(outerFragments[0]?.fragment.rows.map((row) => row.logicalRowIndex)).toEqual([0, 1]);
  });

  it('advances exactly once only after a true fresh-page outcome is validated', () => {
    const doc = freshPageSensitiveSplitDocument();
    const outer = doc.body[1] as unknown as DocTable;
    outer.tblpPr = { ...outer.tblpPr!, tblpY: 80 };
    const measure = recordingCanvas();
    const pages = computePages(
      doc.body,
      __test_verticalLayoutSection(PHYSICAL_SECTION),
      measure.canvas.getContext('2d') as CanvasRenderingContext2D,
      doc.fontFamilyClasses,
    );
    const outerPlacements = pages.flatMap((page, pageIndex) => page.flatMap((element) => {
      if (element.type !== 'table') return [];
      const retained = bodyFragmentFor(element);
      if (retained?.fragment.kind !== 'table' || retained.fragment.source.path[0] !== 1) return [];
      return [{ pageIndex, fragment: retained.fragment as TableFragmentLayout }];
    }));

    expect(pages).toHaveLength(2);
    expect(outerPlacements).toHaveLength(1);
    expect(outerPlacements[0]?.pageIndex).toBe(1);
    expect(outerPlacements[0]?.fragment.rows.map((row) => row.logicalRowIndex)).toEqual([0, 1]);
    expect(pages[0]?.some((element) => element.type === 'table'
      && bodyFragmentFor(element)?.fragment.source.path[0] === 1)).toBe(false);
  });

  it('converges split parent frames through production pagination', () => {
    const doc = documentWith(floatingTable([
      { text: 'FIRST', heightPt: 100 },
      { text: 'SECOND', heightPt: 100 },
    ]));
    const table = doc.body[0] as unknown as DocTable;
    table.tblpPr = { ...table.tblpPr!, tblpY: 100 };
    const measure = recordingCanvas();
    const pages = computePages(
      doc.body,
      __test_verticalLayoutSection(PHYSICAL_SECTION),
      measure.canvas.getContext('2d') as CanvasRenderingContext2D,
      doc.fontFamilyClasses,
    );
    const rowSlices = pages.flatMap((page) => page.flatMap((element) => {
      if (element.type !== 'table') return [];
      const retained = bodyFragmentFor(element);
      if (retained?.fragment.kind !== 'table') return [];
      return [(retained.fragment as TableFragmentLayout).rows.map((row) => row.logicalRowIndex)];
    }));

    expect(pages).toHaveLength(2);
    expect(rowSlices).toEqual([[0], [1]]);
  });
});
