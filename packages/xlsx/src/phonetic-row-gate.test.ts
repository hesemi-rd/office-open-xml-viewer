import { describe, it, expect } from 'vitest';
import { renderViewport } from './renderer.js';
import type { Cell, PhoneticRun, Styles, Worksheet } from './types.js';

/**
 * Row-level furigana gate (ECMA-376 §18.3.1.73 `<row ph>`, issue #814).
 *
 * The parser resolves each cell's `showPhonetic` as `cell/@ph ?? row/@ph ??
 * false`, so a cell that inherits an enabled row arrives here with
 * `showPhonetic: true`, and a cell that overrode the row with `ph="0"` arrives
 * with `showPhonetic: false` (omitted on the wire). This render test drives the
 * real viewport draw and asserts the furigana band is stamped only for the
 * cell whose RESOLVED value is true — i.e. the row-level toggle reaches the
 * renderer gate exactly like the per-cell one, through the same `showPhonetic`
 * field. (The resolution itself is unit-tested in the parser's Rust tests.)
 */

const READING: PhoneticRun[] = [{ sb: 0, eb: 1, text: 'カ' }];

/** A text cell carrying a furigana run, with `showPhonetic` set as the parser
 *  would have resolved it (true = row-inherited or cell opt-in, false = hidden). */
function phoneticCell(col: number, showPhonetic: boolean): Cell {
  return {
    col,
    row: 1,
    value: { type: 'text', text: '課', phoneticRuns: READING },
    styleIndex: 0,
    showPhonetic,
  } as Cell;
}

const BASE_FONT = {
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  size: 11,
  color: null,
  name: null,
};

const STYLES: Styles = {
  fonts: [BASE_FONT],
  fills: [],
  borders: [],
  cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 } as Styles['cellXfs'][number]],
  numFmts: [],
  dxfs: [],
};

/** Worksheet with a single row: A1 resolved show=true, B1 resolved show=false. */
function sheet(aShow: boolean, bShow: boolean): Worksheet {
  return {
    name: 'Sheet1',
    rows: [
      {
        index: 1,
        height: null,
        cells: [phoneticCell(1, aShow), phoneticCell(2, bShow)],
      },
    ],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts: [],
    defaultFontFamily: 'Calibri',
    defaultFontSize: 11,
  } as Worksheet;
}

/** Recording ctx that captures every fillText string and exposes a canvas. */
function recordingCtx(width = 400, height = 200) {
  const texts: string[] = [];
  let font = '11px sans-serif';
  const ctx: Record<string, unknown> = {
    canvas: { width, height },
    get font() {
      return font;
    },
    set font(v: string) {
      font = v;
    },
    fillStyle: '#000',
    strokeStyle: '#000',
    lineWidth: 1,
    textBaseline: 'alphabetic',
    textAlign: 'left',
    letterSpacing: '0px',
    direction: 'ltr',
    globalAlpha: 1,
    measureText: (s: string) => ({ width: [...s].length * 8 }),
    fillText: (t: string) => {
      texts.push(t);
    },
    strokeText: () => {},
    fillRect: () => {},
    strokeRect: () => {},
    clearRect: () => {},
    beginPath: () => {},
    closePath: () => {},
    moveTo: () => {},
    lineTo: () => {},
    rect: () => {},
    arc: () => {},
    fill: () => {},
    stroke: () => {},
    clip: () => {},
    save: () => {},
    restore: () => {},
    translate: () => {},
    rotate: () => {},
    scale: () => {},
    setLineDash: () => {},
    setTransform: () => {},
    createLinearGradient: () => ({ addColorStop: () => {} }),
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, texts };
}

const VIEWPORT = { row: 1, col: 1, rows: 1, cols: 2 };

describe('row-level furigana gate (§18.3.1.73)', () => {
  it('draws the reading for a cell whose resolved showPhonetic is true', () => {
    const { ctx, texts } = recordingCtx();
    renderViewport(ctx, sheet(true, false), STYLES, VIEWPORT);
    // The furigana reading "カ" is stamped exactly once — for A1 (row-inherited
    // true), NOT for B1 (cell ph="0" override → false).
    expect(texts.filter((t) => t === 'カ').length).toBe(1);
    // Both base texts still draw.
    expect(texts.filter((t) => t === '課').length).toBe(2);
  });

  it('skips the reading when every cell resolved to false', () => {
    const { ctx, texts } = recordingCtx();
    renderViewport(ctx, sheet(false, false), STYLES, VIEWPORT);
    expect(texts.some((t) => t === 'カ')).toBe(false);
    expect(texts.filter((t) => t === '課').length).toBe(2);
  });

  it('draws the reading for every cell when the whole row resolved to true', () => {
    const { ctx, texts } = recordingCtx();
    renderViewport(ctx, sheet(true, true), STYLES, VIEWPORT);
    expect(texts.filter((t) => t === 'カ').length).toBe(2);
  });
});
