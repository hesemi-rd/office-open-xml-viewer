import { describe, it, expect } from 'vitest';
import { computePages } from './renderer.js';
import type {
  BodyElement, DocParagraph, DocTable, DocTableRow, DocTableCell,
  DocxTextRun, SectionProps, TblpPr, PaginatedBodyElement,
} from './types';

// ECMA-376 §17.4.57 — a FLOATING table (`<w:tblpPr>`) is positioned absolutely
// (out of flow) and is NOT confined to the text column: Word keeps its declared
// `<w:tblW>`/`<w:tblGrid>` width even when that exceeds the column band, letting
// the box extend into the page margins. sample-28's three page-anchored forms:
//   • tblLayout=fixed, grid 10475 twips = 523.75pt — Word renders 524pt on a
//     451.35pt text band (PDF p.15: box [35, 559] centered on the margin band);
//   • autofit + preferred tblW=10440 dxa, grid = 522pt — Word renders 522pt
//     (PDF p.17/18: box [36, 558]).
// We previously scaled BOTH down to the 451.35pt band (the block-table overflow
// caps in resolveColumnWidths), so every form rendered ~72pt too narrow. For a
// floating table the overflow cap is the PAGE width (the physical constraint),
// not the column band; block tables keep the band cap unchanged.
//
// Deterministic stub canvas (glyph advance = charCount × fontPx). Copied from
// float-table-page-fit.test.ts.
function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, letterSpacing: '0px',
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

// pageWidth 600 / margins 75 ⇒ column band 450; page cap 600.
function section(): SectionProps {
  return {
    pageWidth: 600, pageHeight: 400,
    marginTop: 20, marginRight: 75, marginBottom: 20, marginLeft: 75,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
  } as SectionProps;
}

function textRun(text: string): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'NotInMetrics', isLink: false,
    background: null, vertAlign: null, hyperlink: null,
  } as unknown as DocxTextRun;
}

function cellPara(text: string): unknown {
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{ type: 'text', ...textRun(text) }],
    defaultFontSize: 10, defaultFontFamily: 'NotInMetrics', widowControl: false,
  };
}

function tblp(): TblpPr {
  return {
    leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
    horzAnchor: 'margin', horzSpecified: true, tblpXSpec: 'center',
    vertAnchor: 'page', tblpX: 0, tblpY: 30,
  } as TblpPr;
}

/** A one-row table over `colWidthsPt` (pt). `layout`/`widthPt`/`tblpPr` control
 *  the branch under test. Short cell text keeps min-content below every grid col. */
function table(
  colWidthsPt: number[],
  opts: { layout?: 'fixed'; widthPt?: number; float?: boolean },
): BodyElement {
  const cells: DocTableCell[] = colWidthsPt.map((_, i) => ({
    content: [cellPara(`c${i}`)] as unknown as DocTableCell['content'],
    colSpan: 1, vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background: null, vAlign: 'top', widthPt: null,
  } as DocTableCell));
  const rows: DocTableRow[] = [
    { cells, rowHeight: 20, rowHeightRule: 'exact', isHeader: false } as unknown as DocTableRow,
  ];
  const t: DocTable = {
    colWidths: colWidthsPt, rows,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
    ...(opts.layout ? { layout: opts.layout } : {}),
    ...(opts.widthPt != null ? { widthPt: opts.widthPt } : {}),
    ...(opts.float ? { tblpPr: tblp() } : {}),
  } as unknown as DocTable;
  return { type: 'table', ...t } as unknown as BodyElement;
}

const isTable = (e: PaginatedBodyElement): boolean => e.type === 'table';
const stampedWidth = (pages: PaginatedBodyElement[][]): number => {
  const el = pages.flat().find(isTable);
  expect(el).toBeDefined();
  return ((el as PaginatedBodyElement).tableColWidthsPt ?? []).reduce((s, w) => s + w, 0);
};

describe('floating-table width cap (§17.4.57) — page width, not the column band', () => {
  it('keeps a FIXED-layout floating table at its full tblGrid width past the column band (sample-28 p.15)', () => {
    // Grid 520 > band 450, < page 600 ⇒ stays 520 (was scaled to 450).
    const pages = computePages(
      [table([200, 120, 200], { layout: 'fixed', float: true })],
      section(), makeCtx(),
    );
    expect(stampedWidth(pages)).toBeCloseTo(520, 1);
  });

  it('keeps an AUTOFIT floating table with a preferred tblW at its full grid width (sample-28 p.17)', () => {
    // Autofit + tblW=dxa (grid trusted): grid 520 > band 450 ⇒ stays 520.
    const pages = computePages(
      [table([200, 120, 200], { widthPt: 520, float: true })],
      section(), makeCtx(),
    );
    expect(stampedWidth(pages)).toBeCloseTo(520, 1);
  });

  it('clamps a floating table wider than the PAGE to the page width', () => {
    // Grid 700 > page 600 ⇒ scaled to 600 (the float cannot exceed the paper).
    const pages = computePages(
      [table([300, 100, 300], { layout: 'fixed', float: true })],
      section(), makeCtx(),
    );
    expect(stampedWidth(pages)).toBeCloseTo(600, 1);
  });

  it('still scales a NON-floating fixed table down to the column band (block-table cap unchanged)', () => {
    const pages = computePages(
      [table([200, 120, 200], { layout: 'fixed' })],
      section(), makeCtx(),
    );
    expect(stampedWidth(pages)).toBeCloseTo(450, 1);
  });

  it('leaves a floating table narrower than the band untouched (cap never bites)', () => {
    const pages = computePages(
      [table([100, 100], { layout: 'fixed', float: true })],
      section(), makeCtx(),
    );
    expect(stampedWidth(pages)).toBeCloseTo(200, 1);
  });
});
