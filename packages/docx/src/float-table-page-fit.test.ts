import { describe, it, expect } from 'vitest';
import { computePages } from './renderer.js';
import type {
  BodyElement,
  CellElement,
  DocParagraph,
  DocTable,
  DocTableRow,
  DocxTextRun,
  SectionProps,
  TblpPr,
  PaginatedBodyElement,
} from './types';

// Unit tests for the page-fit pagination of a page-overflowing FLOATING table
// (ECMA-376 §17.4.57 `<w:tblpPr>`).
//
// HISTORY: PR #691 shipped "relocate the whole undivided floating table to the
// next page". That was measured to be WRONG against Word: the Word-exported PDFs
// of private/sample-18 + sample-21 (pdftotext bbox — see issue #674's reopening
// comment) show Word SPLITS a page-overflowing vertAnchor="text" floating table
// ROW BY ROW like a block table, spilling the remainder onto continuation pages,
// and flows the anchor paragraph beside the FINAL continuation band from that
// page's body TOP. The old whole-table-relocation tests are therefore replaced
// here with row-split assertions. A page/margin-anchored floating table is still
// NOT split — its absolute in-page y is instead clamped up into its container by
// computeFloatTableBox (geometry; see float-table-geometry.test.ts), so it stays a
// single element on its page.
//
// The stub canvas mirrors frame-keep-with-anchor.test.ts / pagination.test.ts:
// glyph advance = charCount × fontPx and the font box = 0.8/0.2 em, so a single
// line is exactly fontPx tall. Table row heights are pinned with
// rowHeightRule="exact" so each row's height is deterministic (independent of the
// stub's cell measurement), the table analogue of the frame's hRule="exact"/h.

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
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, textAlign: 'left' as CanvasTextAlign,
    direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

// pageWidth 200 / pageHeight 140, margins 20 ⇒ content band 160×100 (bodyTop 20).
function section(overrides: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 200, pageHeight: 140,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    ...overrides,
  };
}

type DocRun = DocParagraph['runs'][number];

function textRun(text: string, fontSize: number): DocRun {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  };
  return { type: 'text', ...run } as DocRun;
}

function para(opts: { text?: string; fontSize?: number } = {}): BodyElement {
  const fontSize = opts.fontSize ?? 20;
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: opts.text ? [textRun(opts.text, fontSize)] : [],
    defaultFontSize: fontSize, defaultFontFamily: 'NotInMetrics',
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

// Full TblpPr with the spec defaults; callers override only the axis under test.
// The default vertAnchor is 'page' (matching the Rust parser's fill for an
// absent w:vertAnchor), so the vAnchor gate is exercised by overriding it.
function tblp(over: Partial<TblpPr> = {}): TblpPr {
  return {
    leftFromText: 0,
    rightFromText: 0,
    topFromText: 0,
    bottomFromText: 0,
    horzAnchor: 'text',
    horzSpecified: true,
    vertAnchor: 'page',
    tblpX: 0,
    tblpY: 0,
    ...over,
  };
}

/** A single-cell row of the given exact pt height (`rowHeightRule="exact"` short-
 *  circuits content measurement, so the row height == `heightPt` regardless of the
 *  stub canvas — the table analogue of the frame's hRule="exact"/h). The optional
 *  `label` cell text lets a test read which rows landed on which page. */
function row(heightPt: number, label = ''): DocTableRow {
  return {
    cells: [
      {
        content: label ? [para({ text: label }) as unknown as CellElement] : [],
        colSpan: 1,
        vMerge: null,
        borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
        background: null,
        vAlign: 'top',
        widthPt: null,
      },
    ],
    rowHeight: heightPt,
    rowHeightRule: 'exact',
    isHeader: false,
  };
}

/** A floating table (`w:tblpPr`) of `n` exact-height rows, each `rowHPt` tall and
 *  labelled `r1`, `r2`, … so a test can read the per-page row split. Its total
 *  height is `n × rowHPt`, deterministic under the stub. */
function floatTableRows(tp: TblpPr, n: number, rowHPt: number): BodyElement {
  const rows: DocTableRow[] = [];
  for (let i = 1; i <= n; i++) rows.push(row(rowHPt, `r${i}`));
  const t: DocTable = {
    colWidths: [80],
    rows,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
    tblpPr: tp,
  };
  return { type: 'table', ...t } as unknown as BodyElement;
}

/** A floating table (`w:tblpPr`) of total height `tableHPt`, laid out as one
 *  exact-height row so its measured extent is deterministic. */
function floatTable(tp: TblpPr, tableHPt: number): BodyElement {
  return floatTableRows(tp, 1, tableHPt);
}

/** True when an element is a floating table (identified by its tblpPr). */
const isFloatTable = (el: PaginatedBodyElement): boolean =>
  el.type === 'table' && (el as unknown as DocTable).tblpPr != null;

/** True when a page holds a floating table (any slice). */
const hasFloatTable = (page: PaginatedBodyElement[]): boolean => page.some(isFloatTable);

/** The row labels (r1, r2, …) of the floating-table slice(s) on a page, in order. */
const floatRowsOn = (page: PaginatedBodyElement[]): string[] => {
  const out: string[] = [];
  for (const el of page) {
    if (!isFloatTable(el)) continue;
    for (const r of (el as unknown as DocTable).rows) {
      const c0 = r.cells[0]?.content?.[0] as unknown as DocParagraph | undefined;
      const t = c0?.runs?.filter((x) => x.type === 'text').map((x) => (x as DocxTextRun).text).join('') ?? '';
      out.push(t);
    }
  }
  return out;
};

/** Text of a paragraph element (joins its text runs). */
const textOf = (el: PaginatedBodyElement): string =>
  el.type === 'paragraph'
    ? (el as unknown as DocParagraph).runs
        .filter((r) => r.type === 'text')
        .map((r) => (r as DocxTextRun).text)
        .join('')
    : '';

/** True when a page holds the anchor paragraph (matched by its text). */
const hasAnchorText = (page: PaginatedBodyElement[], text: string): boolean =>
  page.some((el) => textOf(el) === text);

/** The newspaper column an element landed in. */
const colOf = (el: PaginatedBodyElement): number | undefined => el.colIndex;

/** Find the (first) floating-table element on a page. */
const floatTableEl = (page: PaginatedBodyElement[]): PaginatedBodyElement | undefined =>
  page.find(isFloatTable);

describe('computePages — floating-table page-fit / row-split (§17.4.57, Word ground truth)', () => {
  // Content band 160×100, bodyTop 20. A vertAnchor="text" table's first slice sits
  // at its in-flow anchor (y=20N after N leading 20pt lines); it overflows once
  // 20N + tableH > 100 and is then SPLIT row-by-row, the remainder continuing at
  // the next page's body top.

  it('(a) splits a text-anchored floating table across pages, greedily filling page 1 from the anchor', () => {
    // 2 leading 20pt lines (y 20→40), then a 5-row table (each 20pt). Anchor at
    // y=40 ⇒ remaining band to the content bottom (100) is 60pt ⇒ 3 rows (r1–r3)
    // fit on page 1; r4–r5 continue at page 2's body top. The trailing anchor text
    // follows onto page 2 (kept with the final band).
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      floatTableRows(tblp({ vertAnchor: 'text', tblpY: 0 }), 5, 20),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(2);
    // Page 1: the two leading lines + a slice holding exactly the rows that fit.
    expect(floatRowsOn(pages[0])).toEqual(['r1', 'r2', 'r3']);
    // Page 2: the continuation slice with the remaining rows.
    expect(floatRowsOn(pages[1])).toEqual(['r4', 'r5']);
    // Every row appears exactly once across the split (no duplication, no loss).
    expect([...floatRowsOn(pages[0]), ...floatRowsOn(pages[1])]).toEqual([
      'r1', 'r2', 'r3', 'r4', 'r5',
    ]);
  });

  it('(b) splits a tall floating table across pages until every row is placed (sample-21 shape)', () => {
    // A single leading 20pt line (anchor at y=20) then an 8-row table (20pt each ⇒
    // 160pt total, > the 100pt content area, so it needs 2 pages). Page 1's band
    // runs from the anchor (y=20) to the bottom (100) = 80pt ⇒ 4 rows (r1–r4);
    // page 2 (fresh, full 100pt band) takes the remaining 4 (r5–r8). This is the
    // reduced analogue of sample-21 (800pt/32 rows → r1–r23 then r24–r32).
    const body = [
      para({ text: 'a' }),
      floatTableRows(tblp({ vertAnchor: 'text', tblpY: 0 }), 8, 20),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(2);
    expect(floatRowsOn(pages[0])).toEqual(['r1', 'r2', 'r3', 'r4']);
    expect(floatRowsOn(pages[1])).toEqual(['r5', 'r6', 'r7', 'r8']);
  });

  it('(c) flows the trailing anchor paragraph from the TERMINAL page body top, not below the band', () => {
    // 2 leading lines (anchor at y=40), 5-row 20pt table split (r1–r3 | r4–r5).
    // The anchor paragraph must land on page 2 (the terminal continuation page) —
    // Word flows it beside the final band from that page's body top, so it never
    // stays on page 1 and never starts below the band.
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      floatTableRows(tblp({ vertAnchor: 'text', tblpY: 0 }), 5, 20),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(hasAnchorText(pages[0], 'anchor')).toBe(false);
    expect(hasAnchorText(pages[1], 'anchor')).toBe(true);
    // The anchor paragraph is the FIRST paragraph on page 2 after the continuation
    // slice (the slice is pushed before the anchor so the wrap band is registered
    // first) — i.e. it starts at the body top beside the band, not after it.
    const page2Types = pages[1].map((el) => (isFloatTable(el) ? 'slice' : textOf(el)));
    expect(page2Types).toEqual(['slice', 'anchor']);
  });

  it('(a-multicol) splits a text-anchored floating table into the NEXT COLUMN in a multi-column section', () => {
    // 2 equal columns: colW = (160-20)/2 = 70; content height 100. Column 0 gets 2
    // leading lines (y 20→40); a 5-row 20pt table then overflows column 0 from the
    // anchor (40 + 60 = 100 exactly for 3 rows, r4–r5 spill). With a column still
    // available on the page, the remainder continues in COLUMN 1 — not a new page.
    const twoCol = section({ columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] } });
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      floatTableRows(tblp({ vertAnchor: 'text', horzAnchor: 'text', tblpY: 0 }), 5, 20),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, twoCol, makeCtx());
    // Still a single page (column 1 absorbed the remaining rows + the anchor).
    expect(pages.length).toBe(1);
    // The first slice sits in column 0, the continuation slice in column 1.
    const slices = pages[0].filter(isFloatTable);
    expect(slices.length).toBe(2);
    expect(colOf(slices[0])).toBe(0);
    expect(colOf(slices[1])).toBe(1);
    // The trailing anchor text follows the final band into column 1.
    const anchor = pages[0].find((el) => textOf(el) === 'anchor');
    expect(anchor).toBeDefined();
    expect(colOf(anchor as PaginatedBodyElement)).toBe(1);
  });

  it('(f) keeps a text-anchored floating table as ONE element when every row fits (no split)', () => {
    // Only 1 leading line (anchor at y=20). A 3-row 20pt table (60pt) fits within
    // [20,100] ⇒ no split. Everything stays on page 1 as a single float element
    // (sample-11 shape: a small vertAnchor="text" float near the page top must not
    // be divided or relocated).
    const body = [
      para({ text: 'a' }),
      floatTableRows(tblp({ vertAnchor: 'text', tblpY: 1 }), 3, 20),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(1);
    expect(floatRowsOn(pages[0])).toEqual(['r1', 'r2', 'r3']);
    // Exactly ONE floating-table element (not sliced).
    const floatCount = pages.reduce((s, p) => s + p.filter(isFloatTable).length, 0);
    expect(floatCount).toBe(1);
    expect(hasAnchorText(pages[0], 'anchor')).toBe(true);
  });

  it('moves the whole table to the next page when not even the first row fits the remaining band', () => {
    // 4 leading lines (anchor at y=80), then a 3-row 30pt table. The remaining band
    // (100−80 = 20pt) cannot hold even the first row (30pt), and a fuller band is
    // one page away ⇒ the WHOLE table moves to page 2 first (no page-1 slice), then
    // fills there (30+30 = 60 ≤ 100 ⇒ all 3 rows: 30·3 = 90 ≤ 100).
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      para({ text: 'c' }),
      para({ text: 'd' }),
      floatTableRows(tblp({ vertAnchor: 'text', tblpY: 0 }), 3, 30),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(2);
    // No slice on page 1 (the four leading lines only).
    expect(hasFloatTable(pages[0])).toBe(false);
    expect(pages[0].map(textOf)).toEqual(['a', 'b', 'c', 'd']);
    // All rows land on page 2 (a fresh full band fits them).
    expect(floatRowsOn(pages[1])).toEqual(['r1', 'r2', 'r3']);
    expect(hasAnchorText(pages[1], 'anchor')).toBe(true);
  });

  it('terminates (no loop) and places a single over-tall row on the page it best fits', () => {
    // 3 leading lines (anchor at y=60), then a 1-row table 150pt tall — taller than
    // the whole 100pt content area, so it can never fit any band. The forward-
    // progress guarantee places it (overflowing) rather than looping. Because the
    // remaining page-1 band (40pt) can't hold it AND a fuller band is one page
    // away, it moves to page 2 first, then is placed there (over-tall rows are not
    // sub-divided — a floating table splits by ROW, and one row is atomic).
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      para({ text: 'c' }),
      floatTableRows(tblp({ vertAnchor: 'text', tblpY: 0 }), 1, 150),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, section(), makeCtx());
    // Terminates with a bounded page count; the over-tall single row is placed once.
    expect(pages.length).toBe(2);
    expect(floatRowsOn(pages[1])).toEqual(['r1']);
    const floatCount = pages.reduce((s, p) => s + p.filter(isFloatTable).length, 0);
    expect(floatCount).toBe(1);
  });

  it('(d) does NOT split a page-anchored floating table (absolute y; clamped in place by geometry)', () => {
    // vertAnchor="page", tblpY=90: the table is pinned at page-y 90 with total
    // height 60 ⇒ it would reach y=150, past the 140 page edge. An absolute page
    // position is the SAME on any page, so it is NOT split or relocated — it stays
    // one element on page 1 (computeFloatTableBox clamps the box UP into the page;
    // see float-table-geometry.test.ts for the clamped y).
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      para({ text: 'c' }),
      floatTableRows(tblp({ vertAnchor: 'page', tblpY: 90 }), 2, 30),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(1);
    expect(hasFloatTable(pages[0])).toBe(true);
    // Not sliced — both rows are on the single element.
    expect(floatRowsOn(pages[0])).toEqual(['r1', 'r2']);
    const floatCount = pages.reduce((s, p) => s + p.filter(isFloatTable).length, 0);
    expect(floatCount).toBe(1);
    expect(hasAnchorText(pages[0], 'anchor')).toBe(true);
  });

  it('(e) does NOT split a margin-anchored floating table (absolute y; clamped in place by geometry)', () => {
    // vertAnchor="margin", tblpY=70: the table is at margin-top(20)+70 = 90 with
    // total height 60 ⇒ bottom 150, past the bottom margin. Absolute ⇒ left as one
    // element (mirrors vertAnchor="page"); the box clamps up into the margin band.
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      para({ text: 'c' }),
      floatTableRows(tblp({ vertAnchor: 'margin', tblpY: 70 }), 2, 30),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(1);
    expect(hasFloatTable(pages[0])).toBe(true);
    expect(floatRowsOn(pages[0])).toEqual(['r1', 'r2']);
    const floatCount = pages.reduce((s, p) => s + p.filter(isFloatTable).length, 0);
    expect(floatCount).toBe(1);
  });

  it('registers each slice as its own floating table so both pages report a float element', () => {
    // Sanity: a split leaves a floating-table element on BOTH pages (one slice
    // each), never a single element straddling the boundary.
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      floatTableRows(tblp({ vertAnchor: 'text', tblpY: 0 }), 5, 20),
      para({ text: 'anchor' }),
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(hasFloatTable(pages[0])).toBe(true);
    expect(hasFloatTable(pages[1])).toBe(true);
    // Two slices total (one per page).
    const floatCount = pages.reduce((s, p) => s + p.filter(isFloatTable).length, 0);
    expect(floatCount).toBe(2);
    // Each slice still carries a tblpPr so the paint pass diverts to renderFloatTable.
    for (const p of pages) {
      const el = floatTableEl(p);
      if (el) expect((el as unknown as DocTable).tblpPr).toBeDefined();
    }
  });
});
