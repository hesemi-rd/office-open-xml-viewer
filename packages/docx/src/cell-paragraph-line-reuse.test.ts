import { describe, it, expect } from 'vitest';
import {
  renderDocumentToCanvas,
  paginateDocument,
  createLayoutServices,
  __test_setFragmentPaintEnabled,
  __test_setLineReuseEnabled,
} from './renderer.js';
import { testFontSnapshot } from './layout/test-font-snapshot.js';
import type { BodyElement, DocParagraph, DocTable, DocTableRow, DocTableCell, DocxDocumentModel, SectionProps, PaginatedBodyElement } from './types';

// ─────────────────────────────────────────────────────────────────────────────
// Phase 4-1 B2 T2 — TABLE-CELL paragraph line reuse (compute-once).
//
// The paginator lays every cell paragraph out at scale 1 while sizing table rows
// (computeTablePtLayout → resolveTableRowHeights → measureCellContentHeightPx →
// measureParaHeight). That measure now STAMPS the scale-1 lines + layout inputs
// onto the cell paragraph, and `renderParagraph`'s existing reuse gate consumes
// them at paint — so cell paragraphs no longer re-run layoutLines twice
// (measure + paint), exactly as body paragraphs already did (Stage 1/2).
//
// These tests pin, using the SAME cross-context flow the public renderPage uses
// (paginate ctx from OffscreenCanvas(1,1) ≠ paint ctx), that:
//   (a) the reuse actually FIRES for cell paragraphs (paint makes strictly fewer
//       measureText calls with reuse ON than OFF) AND is byte-identical (same
//       paint stream);
//   (b) a cell whose paint width no longer matches the stamp (a column re-resolved
//       to a different width) RECOMPUTES — the self-verifying gate rejects the
//       stale stamp — while still painting an identical stream to the recompute;
//   (c) the wrap PARTITION is zoom-invariant: the same table painted at scale 1
//       and at scale 0.75 breaks each cell paragraph at the same points (line
//       count + relative segment order), so a column shrunk by ×scale partitions
//       identically (Stage-2 contract);
//   (d) NESTED tables (a cell containing a table) recurse through the same measure
//       path, so their inner cell paragraphs are stamped and reused too.
// ─────────────────────────────────────────────────────────────────────────────

interface Call { op: 'fill' | 'stroke' | 'img'; text: string; x: number; y: number; font: string; }

/** Linear-metric measure stub for the paginate ctx (node lacks OffscreenCanvas).
 *  Same advance law as the recording paint ctx so paginate == paint measures. */
function makeMeasureStubCtx(): CanvasRenderingContext2D {
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
(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
  getContext() { return makeMeasureStubCtx(); }
};

/** A recording paint ctx with a font-size-LINEAR advance, so scale-1 and paint
 *  lines are bit-identical (this suite is about the reuse mechanism, not hinting).
 *  Records every text/image draw + counts measureText calls (the proxy for the
 *  layoutLines wrap work the stamp eliminates). */
function makeRecordingCanvas(): { canvas: HTMLCanvasElement; calls: Call[]; measures: () => number } {
  let font = '10px serif';
  const calls: Call[] = [];
  let measures = 0;
  let transform = { scaleX: 1, scaleY: 1, translateX: 0, translateY: 0 };
  const stack: typeof transform[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      measures++;
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      const per = p * 0.5;
      return {
        width: [...s].length * per,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() { stack.push({ ...transform }); },
    restore() { transform = stack.pop() ?? transform; },
    beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {},
    scale(x: number, y: number) {
      transform.scaleX *= x;
      transform.scaleY *= y;
    },
    translate(x: number, y: number) {
      transform.translateX += transform.scaleX * x;
      transform.translateY += transform.scaleY * y;
    },
    rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage(_img: unknown, x: number, y: number) {
      calls.push({
        op: 'img', text: '',
        x: transform.translateX + transform.scaleX * x,
        y: transform.translateY + transform.scaleY * y,
        font,
      });
    },
    fillText(s: string, x: number, y: number) {
      calls.push({
        op: 'fill', text: s,
        x: transform.translateX + transform.scaleX * x,
        y: transform.translateY + transform.scaleY * y,
        font,
      });
    },
    strokeText(s: string, x: number, y: number) {
      calls.push({
        op: 'stroke', text: s,
        x: transform.translateX + transform.scaleX * x,
        y: transform.translateY + transform.scaleY * y,
        font,
      });
    },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls, measures: () => measures };
}

// ---- Model builders (mirror table-layout-reuse.test / layout-lines-reuse) ------
function textRun(text: string): DocParagraph['runs'][number] {
  return {
    type: 'text', text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'Times New Roman', fontFamilyEastAsia: '',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  } as DocParagraph['runs'][number];
}
function para(text: string, over: Partial<DocParagraph> = {}): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [textRun(text)],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
    ...over,
  } as unknown as DocParagraph;
}
function emptyBorders() {
  return { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };
}
function cellOf(content: DocParagraph[] | DocTable[], widthPt = 120): DocTableCell {
  return {
    content: content.map((c) => ((c as { type?: string }).type === 'table' ? c : { type: 'paragraph', ...(c as DocParagraph) })),
    colSpan: 1, vMerge: null, borders: emptyBorders(), background: null, vAlign: 'top', widthPt,
  } as unknown as DocTableCell;
}
function cell(text: string, widthPt = 120): DocTableCell { return cellOf([para(text)], widthPt); }
function row(cells: DocTableCell[]): DocTableRow {
  return { cells, rowHeight: null, rowHeightRule: 'auto', isHeader: false } as unknown as DocTableRow;
}
function tableOf(rows: DocTableRow[], colWidths: number[]): DocTable {
  return {
    type: 'table', colWidths, rows, borders: emptyBorders(),
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 2, cellMarginRight: 2,
    jc: 'left', layout: 'fixed',
  } as unknown as DocTable;
}
/** A table with `nRows` rows × 2 cols of multi-word wrapping content. */
function wrapTable(nRows: number): DocTable {
  const rows: DocTableRow[] = [];
  for (let r = 0; r < nRows; r++) {
    rows.push(row([
      cell(`row ${r} left ` + Array.from({ length: 7 }, (_, i) => `wa${i}`).join(' '), 120),
      cell(`row ${r} right ` + Array.from({ length: 7 }, (_, i) => `wb${i}`).join(' '), 160),
    ]));
  }
  return tableOf(rows, [120, 160]);
}
function doc(body: BodyElement[], pageHeight = 200): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 300, pageHeight,
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

/** Render every page at `width` (paint scale = width / pageWidth), returning the
 *  concatenated paint stream per page + the total paint-time measureText count. */
async function renderAll(model: DocxDocumentModel, pages: PaginatedBodyElement[][], width: number): Promise<{ perPage: Call[][]; measures: number }> {
  const perPage: Call[][] = [];
  let measures = 0;
  for (let p = 0; p < pages.length; p++) {
    const rec = makeRecordingCanvas();
    const services = createLayoutServices(model, {
      localMetrics: testFontSnapshot([{ family: 'Times New Roman' }]), measureContext: rec.canvas.getContext('2d'),
    });
    await renderDocumentToCanvas(model, rec.canvas, p, { dpr: 1, width, prebuiltPages: pages, layoutServices: services });
    perPage.push(rec.calls);
    measures += rec.measures();
  }
  return { perPage, measures };
}

/** Paginate once, then paint with reuse OFF and ON at `width`; assert byte-identical
 *  streams and report the measure counts. */
async function reuseVsRecompute(model: DocxDocumentModel, width = 300): Promise<{ pages: number; drawn: number; on: number; off: number; streams: Call[][] }> {
  const prevFragmentPaint = __test_setFragmentPaintEnabled(false);
  try {
    const pages = paginateDocument(model, createLayoutServices(model, { localMetrics: testFontSnapshot([{ family: 'Times New Roman' }]) }));
    const prev = __test_setLineReuseEnabled(false);
    let off: { perPage: Call[][]; measures: number };
    try { off = await renderAll(model, pages, width); } finally { __test_setLineReuseEnabled(prev); }
    const on = await renderAll(model, pages, width);
    expect(on.perPage.length).toBe(off.perPage.length);
    let drawn = 0;
    for (let p = 0; p < on.perPage.length; p++) {
      expect(on.perPage[p]).toEqual(off.perPage[p]); // exact stream identity
      drawn += on.perPage[p].filter((c) => c.op !== 'img').length;
    }
    return { pages: pages.length, drawn, on: on.measures, off: off.measures, streams: on.perPage };
  } finally {
    __test_setFragmentPaintEnabled(prevFragmentPaint);
  }
}

function tableMeasurementGeometry() {
  const model = doc([wrapTable(8) as unknown as BodyElement]);
  const pages = paginateDocument(model, createLayoutServices(model, { localMetrics: testFontSnapshot([{ family: 'Times New Roman' }]) }));
  const tables = pages.flatMap((page) => page.filter((el) => el.type === 'table'));
  return {
    pageCount: pages.length,
    rowHeightsPt: tables.flatMap((table) => table.tableRowHeightsPt ?? []),
    cellLineCounts: tables.flatMap((table) =>
      (table as unknown as DocTable).rows.flatMap((tableRow) =>
        tableRow.cells.flatMap((tableCell) => tableCell.content.map((element) =>
          (element as unknown as { layoutLines?: unknown[] }).layoutLines?.length ?? null,
        )),
      ),
    ),
  };
}

describe('table-cell paragraph line reuse — B2 T2', () => {
  it('preserves table-cell heights and reusable line counts', () => {
    const geometry = tableMeasurementGeometry();

    expect(geometry.pageCount).toBe(2);
    expect(geometry.rowHeightsPt).toEqual([
      22.998046875,
      22.998046875,
      22.998046875,
      22.998046875,
      22.998046875,
      22.998046875,
      22.998046875,
      11.4990234375,
      11.4990234375,
    ]);
    expect(geometry.cellLineCounts).toEqual(Array.from({ length: 18 }, () => 2));
  });

  it('(a) cell paragraphs reuse the paginator stamp: fewer paint measures, identical stream', async () => {
    const r = await reuseVsRecompute(doc([wrapTable(16) as unknown as BodyElement]));
    expect(r.pages).toBeGreaterThan(1); // table split across pages
    expect(r.drawn).toBeGreaterThan(0); // really painted cell text
    // Reuse fired for the cell paragraphs: paint skipped the scale-1 wrap-decision
    // measureText storm it would otherwise re-run per cell.
    expect(r.on).toBeLessThan(r.off);
  });

  it('(b) a stale-width stamp is rejected (recompute), still identical to a fresh recompute', async () => {
    // Paginate, then CORRUPT every cell paragraph's stamped paraW so it no longer
    // matches the paint band. The self-verifying gate must reject each stamp and
    // recompute the partition — measures rise back to the recompute count, and the
    // paint stream stays identical to a from-scratch recompute.
    const model = doc([wrapTable(16) as unknown as BodyElement]);
    const pages = paginateDocument(model, createLayoutServices(model, { localMetrics: testFontSnapshot([{ family: 'Times New Roman' }]) }));

    const poison = (pg: PaginatedBodyElement[][]) => {
      for (const page of pg) for (const el of page) {
        if ((el as { type?: string }).type !== 'table') continue;
        const tbl = el as unknown as DocTable;
        for (const rw of tbl.rows) for (const c of rw.cells) for (const ce of c.content) {
          const st = ce as unknown as { layoutLinesInputs?: { paraW: number } };
          if (st.layoutLinesInputs) st.layoutLinesInputs.paraW = -999; // impossible width
        }
      }
    };
    poison(pages);

    const prev = __test_setLineReuseEnabled(false);
    let off: { perPage: Call[][]; measures: number };
    try { off = await renderAll(model, pages, 300); } finally { __test_setLineReuseEnabled(prev); }
    const on = await renderAll(model, pages, 300);

    // Gate rejected the poisoned stamp ⇒ ON recomputed exactly like OFF.
    expect(on.measures).toBe(off.measures);
    for (let p = 0; p < on.perPage.length; p++) expect(on.perPage[p]).toEqual(off.perPage[p]);
  });

  it('(c) wrap partition is zoom-invariant: scale 1 and scale 0.75 break each cell paragraph identically', async () => {
    const model = doc([wrapTable(10) as unknown as BodyElement]);
    const pages = paginateDocument(model, createLayoutServices(model, { localMetrics: testFontSnapshot([{ family: 'Times New Roman' }]) }));
    // Same paint text at two scales; the fillText SEQUENCE (line partition) must be
    // identical — only the x/y coordinates scale. Compare the per-line text runs.
    const at1 = await renderAll(model, pages, 300);   // scale 1
    const at075 = await renderAll(model, pages, 225);  // scale 0.75 (225 = 300*0.75)
    expect(at1.perPage.length).toBe(at075.perPage.length);
    for (let p = 0; p < at1.perPage.length; p++) {
      const text1 = at1.perPage[p].filter((c) => c.op !== 'img').map((c) => c.text);
      const text075 = at075.perPage[p].filter((c) => c.op !== 'img').map((c) => c.text);
      // Identical drawn-text sequence ⇒ identical wrap partition at both zooms.
      expect(text075).toEqual(text1);
    }
    // And the x positions scale by 0.75 (partition reused, geometry rehydrated).
    const fills1 = at1.perPage[0].filter((c) => c.op === 'fill');
    const fills075 = at075.perPage[0].filter((c) => c.op === 'fill');
    expect(fills075.length).toBe(fills1.length);
    for (let i = 0; i < fills1.length; i++) {
      expect(fills075[i].x).toBeCloseTo(fills1[i].x * 0.75, 6);
    }
  });

  it('(d) nested table: inner-cell paragraphs are stamped and reused too', async () => {
    // Outer cell contains a nested table whose own cells wrap. The nested table's
    // cell paragraphs are measured at scale 1 through the SAME recursion
    // (measureCellElementHeight → estimateTableHeight → measureParaHeight), so they
    // get stamped and must reuse at paint.
    const inner = tableOf([
      row([cell('inner ' + Array.from({ length: 6 }, (_, i) => `x${i}`).join(' '), 100)]),
      row([cell('more ' + Array.from({ length: 6 }, (_, i) => `y${i}`).join(' '), 100)]),
    ], [100]);
    const outer = tableOf([
      row([cellOf([inner] as unknown as DocTable[], 140), cell('side ' + Array.from({ length: 6 }, (_, i) => `z${i}`).join(' '), 140)]),
    ], [140, 140]);
    const r = await reuseVsRecompute(doc([outer as unknown as BodyElement]));
    expect(r.drawn).toBeGreaterThan(0);
    // Reuse fired somewhere in the nested structure (inner + outer cell paragraphs).
    expect(r.on).toBeLessThan(r.off);
    // The inner table's text was actually drawn (nested content painted).
    const drewInner = r.streams.some((page) => page.some((c) => c.text.startsWith('inner') || c.text.startsWith('x')));
    expect(drewInner).toBe(true);
  });

  it('(e) same page painted twice is identical (shared stamped array never mutated by draw)', async () => {
    const model = doc([wrapTable(16) as unknown as BodyElement]);
    const pages = paginateDocument(model, createLayoutServices(model, { localMetrics: testFontSnapshot([{ family: 'Times New Roman' }]) }));
    const first = await renderAll(model, pages, 300);
    const second = await renderAll(model, pages, 300);
    for (let p = 0; p < first.perPage.length; p++) expect(second.perPage[p]).toEqual(first.perPage[p]);
  });
});
