import { describe, expect, it } from 'vitest';
import {
  resolveSectionLayoutContext,
  toLegacyDocGridContext,
  type DocumentLayoutSettings,
} from './layout-context.js';
import { calculateRowHeight } from './renderer.js';
import type {
  DocParagraph,
  DocTable,
  DocTableCell,
  DocTableRow,
} from './types.js';

type MeasureState = Parameters<typeof calculateRowHeight>[4];

function makeMeasureState(
  adjustLineHeightInTable: boolean,
  docGridType: 'lines' | 'snapToChars' = 'lines',
): MeasureState {
  let font = '10px serif';
  // The mock glyph box is a flat 1.0×em (ascent 0.8 + descent 0.2 of the
  // CURRENT font size). No run family is in the core metric table, so every
  // height below is exactly the tallest run's synthetic box — any growth
  // beyond it must come from the grid cell rounding under test.
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() {
      return font;
    },
    set font(value: string) {
      font = value;
    },
    letterSpacing: '0px',
    measureText: (text: string) => ({
      width: [...text].length * 10,
      fontBoundingBoxAscent: px() * 0.8,
      fontBoundingBoxDescent: px() * 0.2,
      actualBoundingBoxAscent: px() * 0.8,
      actualBoundingBoxDescent: px() * 0.2,
    }) as TextMetrics,
    save() {},
    restore() {},
  };
  const kinsoku = {
    enabled: false,
    lineStartForbidden: new Set<number>(),
    lineEndForbidden: new Set<number>(),
  };
  const layoutSettings: DocumentLayoutSettings = {
    kinsoku,
    defaultTabPt: 36,
    documentHasEastAsianText: true,
    compat: {
      adjustLineHeightInTable,
      useFeLayout: false,
      balanceSingleByteDoubleByteWidth: false,
    },
  };
  const sectionLayout = resolveSectionLayoutContext(layoutSettings, {
    pageWidth: 200,
    pageHeight: 300,
    marginTop: 20,
    marginRight: 20,
    marginBottom: 20,
    marginLeft: 20,
    headerDistance: 10,
    footerDistance: 10,
    titlePage: false,
    evenAndOddHeaders: false,
    docGridType,
    docGridLinePitch: 20,
  });

  return {
    ctx: ctx as unknown as CanvasRenderingContext2D,
    scale: 1,
    fontFamilyClasses: {},
    docGrid: toLegacyDocGridContext(sectionLayout),
    layoutSettings,
    sectionLayout,
    docEastAsian: true,
    kinsoku,
    defaultTabPt: 36,
  } as unknown as MeasureState;
}

function paragraph(): DocParagraph {
  return {
    alignment: 'left',
    indentLeft: 0,
    indentRight: 0,
    indentFirst: 0,
    spaceBefore: 0,
    spaceAfter: 0,
    lineSpacing: null,
    numbering: null,
    tabStops: [],
    runs: [{
      type: 'text',
      text: 'あ',
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      fontSize: 10,
      color: null,
      fontFamily: 'serif',
      fontFamilyEastAsia: 'serif',
      isLink: false,
      background: null,
      vertAlign: null,
      hyperlink: null,
    }],
    defaultFontSize: 10,
    defaultFontFamily: 'serif',
    widowControl: false,
  } as unknown as DocParagraph;
}

/** A paragraph with caller-supplied runs (fontSize in pt; text runs get the
 *  same inert defaults as {@link paragraph}). */
function paragraphWithRuns(runs: Record<string, unknown>[]): DocParagraph {
  const base = paragraph() as unknown as { runs: unknown[] };
  base.runs = runs.map((r) => (
    r.type === 'break'
      ? r
      : {
          type: 'text',
          bold: false, italic: false, underline: false, strikethrough: false,
          color: null, fontFamily: 'serif', fontFamilyEastAsia: 'serif',
          isLink: false, background: null, vertAlign: null, hyperlink: null,
          ...r,
        }
  ));
  return base as unknown as DocParagraph;
}

function cellWith(para: DocParagraph): DocTableCell {
  const c = cell() as unknown as { content: unknown[] };
  c.content = [{ type: 'paragraph', ...para }];
  return c as unknown as DocTableCell;
}

function rowWith(para: DocParagraph): DocTableRow {
  const r = row() as unknown as { cells: unknown[] };
  r.cells = [cellWith(para)];
  return r as unknown as DocTableRow;
}

function cell(): DocTableCell {
  return {
    content: [{ type: 'paragraph', ...paragraph() }],
    colSpan: 1,
    vMerge: null,
    borders: {
      top: null,
      bottom: null,
      left: null,
      right: null,
      insideH: null,
      insideV: null,
    },
    background: null,
    vAlign: 'top',
    widthPt: 100,
    marginTop: 0,
    marginBottom: 0,
    marginLeft: 0,
    marginRight: 0,
  } as unknown as DocTableCell;
}

function row(): DocTableRow {
  return {
    cells: [cell()],
    rowHeight: null,
    rowHeightRule: 'auto',
    isHeader: false,
  } as unknown as DocTableRow;
}

function table(): DocTable {
  return {
    colWidths: [100],
    rows: [],
    borders: {
      top: null,
      bottom: null,
      left: null,
      right: null,
      insideH: null,
      insideV: null,
    },
    cellMarginTop: 0,
    cellMarginBottom: 0,
    cellMarginLeft: 0,
    cellMarginRight: 0,
    jc: 'left',
  } as unknown as DocTable;
}

describe('table-cell line grid compatibility', () => {
  it('keeps cell text at natural height when compatibility is disabled', () => {
    expect(calculateRowHeight(row(), table(), [100], 1, makeMeasureState(false))).toBe(10);
  });

  it('applies the section line pitch when compatibility is enabled', () => {
    expect(calculateRowHeight(row(), table(), [100], 1, makeMeasureState(true))).toBe(20);
  });

  it('gates the line axis of snapToChars by the same compatibility setting', () => {
    expect(calculateRowHeight(
      row(),
      table(),
      [100],
      1,
      makeMeasureState(false, 'snapToChars'),
    )).toBe(10);
    expect(calculateRowHeight(
      row(),
      table(),
      [100],
      1,
      makeMeasureState(true, 'snapToChars'),
    )).toBe(20);
  });
});

// The docGrid line-cell count (docGridLineCells) is derived from the line's
// resolved single-line height — the TALLEST run's box governs the line
// (ECMA-376 §17.3.1.33; the grid reserves the whole cells that CONTAIN that
// box, §17.6.5 / issue #1013 sample-58 adjudication). These integration cases
// pin two call-path properties the pure lineBoxHeight tests cannot see: which
// line height the layout hands over, and which script gate each LINE uses. The
// grid is active in a cell via adjustLineHeightInTable (§17.15.3.1); pitch =
// 20 pt; the mock font box is a flat 1.0×em, so every height difference below
// comes from the line-height / script routing alone.
describe('docGrid line-cell integration through the cell measure path', () => {
  it('does not whole-cell round tall East Asian design metrics for explicit atLeast', () => {
    const para = paragraphWithRuns([{
      text: 'あ',
      fontSize: 14,
      fontFamily: 'Meiryo',
      fontFamilyEastAsia: 'Meiryo',
    }]);
    para.lineSpacing = { value: 0, rule: 'atLeast', explicit: true };

    // Meiryo's established design line is 3269/2048 em. At 14pt it is
    // 22.35pt, taller than the 20pt table grid pitch. Explicit atLeast expands
    // to that content height; automatic grid layout would round it to 40pt.
    expect(calculateRowHeight(rowWith(para), table(), [100], 1, makeMeasureState(true)))
      .toBeCloseTo(14 * 3269 / 2048, 12);
  });

  it('a manual line break in a SMALLER run does not shrink the line height (tallest governs)', () => {
    // Line 1: 'あ' at 24 pt + 'い' at 10 pt, terminated by a <w:br> whose
    // nearby size resolves to 10 pt (§17.3.3.1; findNearbyFontSize looks at the
    // preceding run). Line 2: 'う' at 10 pt. The break must NOT overwrite the
    // line's tallest box (24 px → ceil(24/20) = 2 cells = 40) with its own
    // 10 pt box (1 cell = 20).
    const para = paragraphWithRuns([
      { text: 'あ', fontSize: 24 },
      { text: 'い', fontSize: 10 },
      { type: 'break', breakType: 'line' },
      { text: 'う', fontSize: 10 },
    ]);
    expect(calculateRowHeight(rowWith(para), table(), [100], 1, makeMeasureState(true)))
      .toBe(60); // 40 (2 cells) + 20 (1 cell)
  });

  it('the East Asian cell rounding is gated per LINE, not per paragraph', () => {
    // Line 1: CJK 10 pt → 1 cell (20). Line 2: Latin-only 'Hello' at 22 pt —
    // Word does not cell-round Latin lines; they keep their natural height
    // above a one-cell floor (mock natural = 22 px > floor 20 → 22), NOT the
    // cell count ceil(22/20) = 2 cells = 40 that a paragraph-level East Asian
    // flag would apply.
    const para = paragraphWithRuns([
      { text: 'あ', fontSize: 10 },
      { type: 'break', breakType: 'line' },
      { text: 'Hello', fontSize: 22 },
    ]);
    expect(calculateRowHeight(rowWith(para), table(), [100], 1, makeMeasureState(true)))
      .toBe(42); // 20 (CJK, 1 cell) + 22 (Latin, natural above the one-cell floor)
  });
});
