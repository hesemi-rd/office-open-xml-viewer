import { describe, it, expect } from 'vitest';
import { calculateRowHeight } from './renderer.js';
import type { DocTable, DocTableRow, DocTableCell, DocParagraph } from './types.js';

// The render/measure state type, taken from calculateRowHeight's signature so
// the test does not depend on RenderState being exported (mirrors
// table-row-height.test's EMPTY_STATE pattern).
type MeasureState = Parameters<typeof calculateRowHeight>[4];

// ECMA-376 §17.3.1.12 (`<w:ind>`) + §17.4.80 (auto row height): a table cell's
// auto row height is the tallest cell's measured CONTENT height. The content is
// laid out inside `cellWidth − cell margins − the paragraph's own left/right
// indent`, with the FIRST line further inset by `<w:ind w:firstLine>`. The paint
// path (`renderParagraph`) and the paginator (`estimateParagraphHeight`) both
// apply those indents; the per-row measurer (`measureParaHeight`) must too, or a
// cell whose paragraph carries a first-line indent that forces a wrap is sized
// for FEWER lines than it paints and overflows into the next row.
//
// sample-11.docx exposes this: the "City or Town" header cell and the
// "Cedar University" data cells carry `w:firstLine=432` (21.6 pt). At their
// column width the text wraps to two lines when the indent is honored, but the
// (pre-fix) measurer ignored the indent, sized the row for one line, and the
// second line ("Town" / "University") bled into the row below.
//
// This is a DIFFERENTIAL test: two cells identical except for the first-line
// indent. The recording ctx measures every glyph at `fontSize` px, so the
// wrap is deterministic. The indented cell MUST be taller (it wraps to two
// lines); the un-indented cell stays one line. Pre-fix both measured one line
// (equal heights) — so this fails Red and passes Green.

function makeMeasureState(): MeasureState {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() {
      return font;
    },
    set font(v: string) {
      font = v;
    },
    letterSpacing: '0px',
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
    save() {},
    restore() {},
    measureText2() {},
  };
  return {
    ctx: ctx as unknown as CanvasRenderingContext2D,
    scale: 1,
    fontFamilyClasses: {},
    docGrid: { type: null, linePitchPt: null, charSpacePt: null },
    docEastAsian: false,
  } as unknown as MeasureState;
}

function para(text: string, indentFirst: number): DocParagraph {
  return {
    alignment: 'left',
    indentLeft: 0,
    indentRight: 0,
    indentFirst,
    spaceBefore: 0,
    spaceAfter: 0,
    lineSpacing: null,
    numbering: null,
    tabStops: [],
    runs: [
      {
        type: 'text',
        text,
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
        fontSize: 10,
        color: null,
        fontFamily: 'Times New Roman',
        fontFamilyEastAsia: 'Times New Roman',
        isLink: false,
        background: null,
        vertAlign: null,
        hyperlink: null,
      },
    ],
    defaultFontSize: 10,
    defaultFontFamily: 'Times New Roman',
    widowControl: false,
  } as unknown as DocParagraph;
}

function cellWith(p: DocParagraph): DocTableCell {
  return {
    content: [{ type: 'paragraph', ...p }],
    colSpan: 1,
    vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background: null,
    vAlign: 'top',
    widthPt: 120,
    marginTop: 0,
    marginBottom: 0,
    marginLeft: 0,
    marginRight: 0,
  } as unknown as DocTableCell;
}

function rowWith(p: DocParagraph): DocTableRow {
  return {
    cells: [cellWith(p)],
    rowHeight: null,
    rowHeightRule: 'auto',
    isHeader: false,
  } as unknown as DocTableRow;
}

function table(): DocTable {
  return {
    colWidths: [120],
    rows: [],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0,
    cellMarginBottom: 0,
    cellMarginLeft: 0,
    cellMarginRight: 0,
    jc: 'left',
  } as unknown as DocTable;
}

describe('calculateRowHeight — paragraph indent affects wrap (§17.3.1.12 / §17.4.80)', () => {
  const COLS = [120];
  const SCALE = 1;
  const t = table();
  // "WORD ALPHA" = 10 glyphs × 10 px = 100 px. Fits the 120 px column on one
  // line with no indent. With a 60 px first-line indent the first line only has
  // 60 px, so it breaks after "WORD" (40 px) and "ALPHA" wraps to a second line.
  const TEXT = 'WORD ALPHA';

  it('un-indented cell stays one line', () => {
    const h = calculateRowHeight(rowWith(para(TEXT, 0)), t, COLS, SCALE, makeMeasureState());
    // One 10 px line box (+ its descent). Comfortably below two lines.
    expect(h).toBeLessThan(16);
  });

  it('first-line-indent cell wraps to two lines and the row grows to fit', () => {
    const h1 = calculateRowHeight(rowWith(para(TEXT, 0)), t, COLS, SCALE, makeMeasureState());
    const h2 = calculateRowHeight(rowWith(para(TEXT, 60)), t, COLS, SCALE, makeMeasureState());
    // The indented cell wraps to a second line, so its row must be ~2× taller.
    expect(h2).toBeGreaterThan(h1 * 1.8);
  });
});
