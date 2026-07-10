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
      fontBoundingBoxAscent: 8,
      fontBoundingBoxDescent: 2,
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
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
