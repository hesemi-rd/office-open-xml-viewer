import { describe, it, expect, beforeAll } from 'vitest';
import { createLayoutServices, renderDocumentToCanvas, paginateDocument } from './renderer.js';
import type {
  BodyElement,
  CellElement,
  DocParagraph,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  FramePr,
  PaginatedBodyElement,
  SectionProps,
} from './types';
import { bodyFragmentFor } from './renderer.js';
import type { TableFragmentLayout } from './layout/table-pagination.js';

// ─────────────────────────────────────────────────────────────────────────────
// PR 5 Task 13 — body fragment paint purity.
//
// A migrated body paragraph paints from its stored measured fragment
// through the renderer-owned retained adapter. At paint scale 1, stored point
// geometry needs no
// rescale, so the paint pass must draw the paragraph's lines WITHOUT calling
// measureText at all — no line layout, no segment measurement, no remeasurement.
//
// This is proved end-to-end through the real production flow: pages are paginated
// with a normal OffscreenCanvas metric, then painted at scale 1 onto a canvas whose
// measureText THROWS. If any part of the migrated paragraph paint tried to measure,
// the render would throw; instead it completes and draws the paragraph text.
// ─────────────────────────────────────────────────────────────────────────────

interface Call { text: string; x: number; y: number; }

/** Pagination-side canvas with a normal linear glyph metric. */
function makeMeasuringCtx(): CanvasRenderingContext2D {
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
  return ctx as unknown as CanvasRenderingContext2D;
}

beforeAll(() => {
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    getContext() { return makeMeasuringCtx(); }
  };
});

/** Paint-side recording canvas whose measureText THROWS — any measurement during
 *  paint fails the test loudly. Records every text draw for the content assertion. */
function makeThrowingPaintCanvas(): { canvas: HTMLCanvasElement; calls: Call[]; measured: () => number } {
  let font = '10px serif';
  const calls: Call[] = [];
  let measured = 0;
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (_s: string) => {
      measured++;
      throw new Error('measureText must not be called during fragment paint');
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(s: string, x: number, y: number) { calls.push({ text: s, x, y }); },
    strokeText(s: string, x: number, y: number) { calls.push({ text: s, x, y }); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls, measured: () => measured };
}

function makeMeasuringPaintCanvas(): { canvas: HTMLCanvasElement; calls: Call[] } {
  let font = '10px serif';
  const calls: Call[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (text: string) => {
      const size = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...text].length * size * 0.5,
        fontBoundingBoxAscent: size * 0.8, fontBoundingBoxDescent: size * 0.2,
        actualBoundingBoxAscent: size * 0.8, actualBoundingBoxDescent: size * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(text: string, x: number, y: number) { calls.push({ text, x, y }); },
    strokeText(text: string, x: number, y: number) { calls.push({ text, x, y }); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls };
}

function para(text: string, over: Partial<DocParagraph> = {}): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'Times New Roman',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
    } as DocParagraph['runs'][number]],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
    ...over,
  } as unknown as DocParagraph;
}

function frame(over: Partial<FramePr> = {}): FramePr {
  return {
    dropCap: 'none', lines: 1, wrap: 'around',
    hAnchor: 'text', vAnchor: 'text', hRule: 'auto',
    hSpace: 0, vSpace: 0,
    ...over,
  };
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

// ---- Table builders (PR 6 Task 16) --------------------------------------------
function eb() {
  return { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };
}
function tcell(content: CellElement[], over: Partial<DocTableCell> = {}): DocTableCell {
  return {
    content, colSpan: 1, vMerge: null, borders: eb(),
    background: null, vAlign: 'top', widthPt: null, ...over,
  } as unknown as DocTableCell;
}
function textCell(text: string, over: Partial<DocTableCell> = {}): DocTableCell {
  return tcell([{ type: 'paragraph', ...para(text) } as unknown as CellElement], over);
}
function trow(cells: DocTableCell[], over: Partial<DocTableRow> = {}): DocTableRow {
  return { cells, rowHeight: null, rowHeightRule: 'auto', isHeader: false, ...over } as unknown as DocTableRow;
}
function tbl(rows: DocTableRow[], colWidths: number[], over: Partial<DocTable> = {}): DocTable {
  return {
    type: 'table', colWidths, rows, borders: eb(),
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left', layout: 'fixed', ...over,
  } as unknown as DocTable;
}

describe('table fragment paint purity (PR 6 Task 16)', () => {
  it('paints a table with a NESTED table + vMerge at scale 1 without calling measureText', async () => {
    // Both outer and nested tables must be present in the retained layout tree;
    // paint traverses those nodes without reacquiring paragraph metrics.
    const inner = tbl([trow([textCell('inner one'), textCell('inner two')])], [40, 40]);
    const outer = tbl(
      [
        trow([textCell('a', { vMerge: true }), textCell('top')]),
        trow([textCell('', { vMerge: false }), tcell([{ type: 'table', ...inner } as unknown as CellElement])]),
      ],
      [60, 100],
    );
    const model = doc([outer as unknown as BodyElement]);
    const pages = paginateDocument(model);
    const placed = bodyFragmentFor(pages[0]![0]!);
    expect(placed?.fragment.kind).toBe('table');
    if (placed?.fragment.kind !== 'table' || !('flowBounds' in placed.fragment)) {
      throw new Error('expected retained TableLayout/TableFragmentLayout');
    }
    const nestedLayouts = placed.fragment.rows.flatMap((tableRow) => tableRow.cells.flatMap((tableCell) =>
      tableCell.blocks.map((block) => block.layout).filter((layout) => layout.kind === 'table')));
    expect(nestedLayouts).toHaveLength(1);
    expect(nestedLayouts[0]!.rows.length).toBeGreaterThan(0);
    const paint = makeThrowingPaintCanvas();
    await expect(
      renderDocumentToCanvas(model, paint.canvas, 0, { dpr: 1, width: 200, prebuiltPages: pages }),
    ).resolves.not.toThrow();
    expect(paint.measured()).toBe(0);
    // Non-vacuity: outer + inner cell text were drawn.
    expect(paint.calls.some((c) => c.text.includes('top'))).toBe(true);
    expect(paint.calls.some((c) => c.text.includes('inner'))).toBe(true);
  });

  it('paints a page-split table with a repeated header at scale 1, measure-free', async () => {
    const bodyRows = Array.from({ length: 12 }, (_v, i) => trow([textCell(`row ${i}`)]));
    const rows = [trow([textCell('HEADER')], { isHeader: true }), ...bodyRows];
    const model = doc([tbl(rows, [120]) as unknown as BodyElement], 120);
    const pages = paginateDocument(model);
    expect(pages.length).toBeGreaterThan(1);
    const retained = pages.map((page) => {
      const tableElement = page.find((element) => element.type === 'table');
      expect(tableElement).toBeDefined();
      const placed = bodyFragmentFor(tableElement!);
      if (placed?.fragment.kind !== 'table' || !('flowBounds' in placed.fragment)) {
        throw new Error('expected retained TableFragmentLayout');
      }
      return placed.fragment as TableFragmentLayout;
    });
    for (const fragment of retained.slice(1)) {
      expect(fragment.rows[0]?.ownership).toBe('repeated-header');
      expect(fragment.rows[0]?.logicalRowIndex).toBe(0);
    }
    for (let p = 0; p < pages.length; p++) {
      const paint = makeThrowingPaintCanvas();
      await expect(
        renderDocumentToCanvas(model, paint.canvas, p, { dpr: 1, width: 200, prebuiltPages: pages }),
      ).resolves.not.toThrow();
      expect(paint.measured()).toBe(0);
      expect(paint.calls.length).toBeGreaterThan(0);
    }
    // The header text is repeated on the continuation page.
    const paint2 = makeThrowingPaintCanvas();
    await renderDocumentToCanvas(model, paint2.canvas, 1, { dpr: 1, width: 200, prebuiltPages: pages });
    expect(paint2.calls.some((c) => c.text.includes('HEADER'))).toBe(true);
  });
});

describe('fragment paint purity (PR 5 Task 13)', () => {
  it('keeps header/footer frame paragraphs on the B1 legacy story painter', async () => {
    const model = doc([para('body') as unknown as BodyElement]);
    model.headers.default = {
      body: [{
        type: 'paragraph',
        ...para('legacy header frame', { framePr: frame({ w: 80, hRule: 'auto' }) }),
      } as unknown as BodyElement],
    };
    const pages = paginateDocument(model);
    const paint = makeMeasuringPaintCanvas();

    await renderDocumentToCanvas(model, paint.canvas, 0, {
      dpr: 1, width: 200, prebuiltPages: pages,
    });

    expect(paint.calls.map((call) => call.text).join('')).toContain('legacy header frame');
  });

  it('paints a body text frame from retained geometry without measuring', async () => {
    const framed = para('retained frame text', {
      framePr: frame({ w: 50 }),
    });
    const model = doc([
      framed as unknown as BodyElement,
      para('anchor paragraph') as unknown as BodyElement,
    ]);
    const pages = paginateDocument(model);
    const placed = bodyFragmentFor(pages[0]![0]!);
    expect(placed?.fragment.kind).toBe('paragraph');
    expect(placed?.heightPt).toBe(0);
    if (placed?.fragment.kind !== 'paragraph') throw new Error('expected frame paragraph layout');
    expect(placed.fragment.ordinaryFlow).toBe(false);
    expect(placed.fragment.lines.length).toBeGreaterThan(0);
    expect(placed.fragment.lines.flatMap((line) => line.placements)).not.toHaveLength(0);

    const retainedLines = placed.fragment.lines;
    const partition = retainedLines.map((line) => line.range);
    for (const width of [200, 400]) {
      const paint = makeThrowingPaintCanvas();
      await expect(
        renderDocumentToCanvas(model, paint.canvas, 0, { dpr: 1, width, prebuiltPages: pages }),
      ).resolves.not.toThrow();
      expect(paint.measured()).toBe(0);
      expect(paint.calls.map((call) => call.text).join('')).toContain('retained frame');
      expect(placed.fragment.lines).toBe(retainedLines);
      expect(placed.fragment.lines.map((line) => line.range)).toEqual(partition);
    }
  });

  it('prepares frame metadata even when callers provide custom layout services', () => {
    const model = doc([
      para('custom services frame', { framePr: frame({ w: 50 }) }) as unknown as BodyElement,
      para('anchor') as unknown as BodyElement,
    ]);
    const base = createLayoutServices(model);
    const custom = { text: base.text, images: base.images, math: base.math };
    const pages = paginateDocument(model, custom);
    const placed = bodyFragmentFor(pages[0]![0]!);

    expect(placed?.fragment.kind).toBe('paragraph');
    if (placed?.fragment.kind !== 'paragraph') throw new Error('expected frame paragraph layout');
    expect(placed.fragment.ordinaryFlow).toBe(false);
  });

  it('retains identical adjacent framePr paragraphs as one stacked frame and one exclusion', () => {
    const shared = frame({ w: 50 });
    const model = doc([
      para('frame first', { framePr: { ...shared } }) as unknown as BodyElement,
      para('frame second', { framePr: { ...shared } }) as unknown as BodyElement,
      para('anchor') as unknown as BodyElement,
    ]);
    const pages = paginateDocument(model);
    const first = bodyFragmentFor(pages[0]![0]!);
    const second = bodyFragmentFor(pages[0]![1]!);
    const anchor = bodyFragmentFor(pages[0]![2]!);
    expect(first?.fragment.kind).toBe('paragraph');
    expect(second?.fragment.kind).toBe('paragraph');
    expect(first?.heightPt).toBe(0);
    expect(second?.heightPt).toBe(0);
    if (first?.fragment.kind !== 'paragraph' || second?.fragment.kind !== 'paragraph') {
      throw new Error('expected frame paragraph layouts');
    }
    expect(second!.yPt).toBeGreaterThanOrEqual(first!.yPt + first!.fragment.advancePt);
    expect(anchor?.fragment.kind).toBe('paragraph');
    if (anchor?.fragment.kind !== 'paragraph') throw new Error('expected anchor paragraph layout');
    expect(anchor.fragment.exclusions.filter((exclusion) => exclusion.id.includes('frame'))).toHaveLength(1);
  });

  it('uses final-width reflow to determine an automatic frame height', () => {
    const model = doc([
      para('abcdefghij', { framePr: frame({ w: 20, hRule: 'auto' }) }) as unknown as BodyElement,
      para('anchor') as unknown as BodyElement,
    ]);
    const pages = paginateDocument(model);
    const framed = bodyFragmentFor(pages[0]![0]!);
    const anchor = bodyFragmentFor(pages[0]![1]!);
    expect(framed?.fragment.kind).toBe('paragraph');
    if (framed?.fragment.kind !== 'paragraph') throw new Error('expected frame paragraph layout');
    expect(framed.fragment.lines.length).toBeGreaterThan(1);
    expect(anchor?.fragment.kind).toBe('paragraph');
    if (anchor?.fragment.kind !== 'paragraph') throw new Error('expected anchor paragraph layout');
    expect(anchor.fragment.exclusions[0]?.bounds.heightPt)
      .toBeCloseTo(framed.fragment.advancePt, 6);
  });

  it('uses the larger of authored and final-content height for hRule=atLeast', () => {
    const contentDriven = doc([
      para('abcdefghij', { framePr: frame({ w: 20, hRule: 'atLeast', h: 5 }) }) as unknown as BodyElement,
      para('anchor one') as unknown as BodyElement,
    ]);
    const authoredDriven = doc([
      para('x', { framePr: frame({ w: 20, hRule: 'atLeast', h: 80 }) }) as unknown as BodyElement,
      para('anchor two') as unknown as BodyElement,
    ]);
    const contentPages = paginateDocument(contentDriven);
    const authoredPages = paginateDocument(authoredDriven);
    const contentFrame = bodyFragmentFor(contentPages[0]![0]!)!;
    const contentAnchor = bodyFragmentFor(contentPages[0]![1]!)!;
    const authoredAnchor = bodyFragmentFor(authoredPages[0]![1]!)!;
    if (contentFrame.fragment.kind !== 'paragraph'
      || contentAnchor.fragment.kind !== 'paragraph'
      || authoredAnchor.fragment.kind !== 'paragraph') throw new Error('expected paragraph layouts');

    expect(contentAnchor.fragment.exclusions[0]?.bounds.heightPt)
      .toBeCloseTo(contentFrame.fragment.advancePt, 6);
    expect(contentAnchor.fragment.exclusions[0]?.bounds.heightPt).toBeGreaterThan(5);
    expect(authoredAnchor.fragment.exclusions[0]?.bounds.heightPt).toBeCloseTo(80, 6);
  });

  it('retains one authored outer clip on every member of an hRule=exact frame group', () => {
    const shared = frame({ w: 30, hRule: 'exact', h: 25 });
    const model = doc([
      para('first frame member', { framePr: { ...shared } }) as unknown as BodyElement,
      para('second frame member', { framePr: { ...shared } }) as unknown as BodyElement,
      para('anchor') as unknown as BodyElement,
    ]);
    const pages = paginateDocument(model);
    const first = bodyFragmentFor(pages[0]![0]!)!;
    const second = bodyFragmentFor(pages[0]![1]!)!;
    const anchor = bodyFragmentFor(pages[0]![2]!)!;
    if (first.fragment.kind !== 'paragraph'
      || second.fragment.kind !== 'paragraph'
      || anchor.fragment.kind !== 'paragraph') throw new Error('expected paragraph layouts');

    expect(first.fragment.clipBounds).toEqual(second.fragment.clipBounds);
    expect(first.fragment.clipBounds).toEqual({
      xPt: first.fragment.clipBounds?.xPt,
      yPt: first.fragment.clipBounds?.yPt,
      widthPt: 30,
      heightPt: 25,
    });
    expect(anchor.fragment.exclusions[0]?.bounds.heightPt).toBeCloseTo(25, 6);
  });

  it('folds contextual spacing once across a three-paragraph frame group', () => {
    const shared = frame({ w: 50 });
    const model = doc([
      para('one', { framePr: { ...shared }, styleId: 's', spaceBefore: 2, spaceAfter: 6 }) as unknown as BodyElement,
      para('two', { framePr: { ...shared }, styleId: 's', contextualSpacing: true, spaceBefore: 4, spaceAfter: 5 }) as unknown as BodyElement,
      para('three', { framePr: { ...shared }, styleId: 's', spaceBefore: 8, spaceAfter: 3 }) as unknown as BodyElement,
      para('anchor') as unknown as BodyElement,
    ]);
    const pages = paginateDocument(model);
    const first = bodyFragmentFor(pages[0]![0]!)!;
    const second = bodyFragmentFor(pages[0]![1]!)!;
    const third = bodyFragmentFor(pages[0]![2]!)!;
    const anchor = bodyFragmentFor(pages[0]![3]!)!;
    if (
      first.fragment.kind !== 'paragraph'
      || second.fragment.kind !== 'paragraph'
      || third.fragment.kind !== 'paragraph'
    ) throw new Error('expected frame paragraph layouts');
    expect(second.yPt - first.yPt).toBeCloseTo(first.fragment.advancePt, 6);
    // p2 after=5 is replaced by the contextual gap 3 before p3.
    expect(third.yPt - second.yPt).toBeCloseTo(second.fragment.advancePt - 2, 6);
    if (anchor.fragment.kind !== 'paragraph') throw new Error('expected anchor paragraph layout');
    expect(anchor.fragment.exclusions[0]?.bounds.heightPt).toBeCloseTo(
      third.yPt + third.fragment.advancePt - (first.yPt - 2),
      6,
    );
  });

  it('merges paragraph borders inside one frame and reserves the final bottom edge once', () => {
    const edge = { style: 'single', color: '000000', width: 2, space: 3 };
    const borders = { top: edge, right: edge, bottom: edge, left: edge, between: null };
    const shared = frame({ w: 50 });
    const model = doc([
      para('one', { framePr: { ...shared }, borders, spaceBefore: 0, spaceAfter: 0 }) as unknown as BodyElement,
      para('two', { framePr: { ...shared }, borders, spaceBefore: 0, spaceAfter: 0 }) as unknown as BodyElement,
      para('anchor') as unknown as BodyElement,
    ]);
    const pages = paginateDocument(model);
    const first = bodyFragmentFor(pages[0]![0]!)!;
    const second = bodyFragmentFor(pages[0]![1]!)!;
    const anchor = bodyFragmentFor(pages[0]![2]!)!;
    if (first.fragment.kind !== 'paragraph' || second.fragment.kind !== 'paragraph') {
      throw new Error('expected frame paragraph layouts');
    }
    expect(first.fragment.borders.some((border) => border.edge === 'bottom')).toBe(false);
    expect(second.fragment.borders.some((border) => border.edge === 'bottom')).toBe(true);
    expect(second.yPt).toBeCloseTo(first.yPt + first.fragment.advancePt, 6);
    if (anchor.fragment.kind !== 'paragraph') throw new Error('expected anchor paragraph layout');
    expect(anchor.fragment.exclusions[0]?.bounds.heightPt).toBeCloseTo(
      second.yPt + second.fragment.advancePt - first.yPt,
      6,
    );
  });

  it('paints a premeasured body paragraph at scale 1 without ever calling measureText', async () => {
    const model = doc([para('hello world one two three') as unknown as BodyElement]);
    const pages = paginateDocument(model); // measured with the normal OffscreenCanvas
    const paint = makeThrowingPaintCanvas();

    // Paint scale 1 (render width == page width). A migrated paragraph draws its
    // stored fragment lines; nothing measures.
    await expect(
      renderDocumentToCanvas(model, paint.canvas, 0, { dpr: 1, width: 200, prebuiltPages: pages }),
    ).resolves.not.toThrow();

    expect(paint.measured()).toBe(0);
    // Non-vacuity: the paragraph's words were actually drawn.
    const drewText = paint.calls.some((c) => c.text.includes('hello'));
    expect(drewText).toBe(true);
  });

  it('paints a paragraph that SPLITS across pages from fragments, still measure-free', async () => {
    // A long paragraph over a short page splits; each continuation slice paints from
    // the shared measured fragment window without remeasuring.
    const long = Array.from({ length: 120 }, () => 'w').join(' ');
    const model = doc([para(long) as unknown as BodyElement], 60);
    const pages = paginateDocument(model);
    expect(pages.length).toBeGreaterThan(1);
    const split = pages.some((pg) => pg.some((el) => (el as PaginatedBodyElement).lineSlice));
    expect(split).toBe(true);

    for (let p = 0; p < pages.length; p++) {
      const paint = makeThrowingPaintCanvas();
      await expect(
        renderDocumentToCanvas(model, paint.canvas, p, { dpr: 1, width: 200, prebuiltPages: pages }),
      ).resolves.not.toThrow();
      expect(paint.measured()).toBe(0);
      expect(paint.calls.length).toBeGreaterThan(0);
    }
  });
});
