import { describe, it, expect } from 'vitest';
import type { Stroke } from '@silurus/ooxml-core';
import { renderTable } from './renderer.js';
import type { TableCell, TableElement, TableRow } from './types';

// DrawingML `<a:tbl>` — end-to-end render: an interior gridline SHARED by two
// adjacent cells is drawn exactly ONCE (not once per cell), and when the two
// cells DISAGREE the drawn line is the conflict WINNER, not the later-painted
// cell. The spec is SILENT for PresentationML (no §17.4.66 analog), so this is
// the pptx leg / structural mirror of docx PR #811.
//
// We record each stroked segment (endpoints + width + colour) from a fake 2D
// context (the pptx renderer strokes borders via beginPath/moveTo/lineTo/stroke
// after applyStroke sets strokeStyle + lineWidth), then read the shared line.

interface StrokeSeg { x1: number; y1: number; x2: number; y2: number; width: number; color: string; }

function makeRecordingCtx(): { ctx: CanvasRenderingContext2D; strokes: StrokeSeg[] } {
  const strokes: StrokeSeg[] = [];
  let cur = { x: 0, y: 0 };
  let pending: { x1: number; y1: number; x2: number; y2: number } | null = null;
  let lineWidth = 1;
  let strokeStyle = '#000000';
  const ctx = {
    canvas: { width: 1000, height: 1000 },
    get lineWidth() { return lineWidth; },
    set lineWidth(v: number) { lineWidth = v; },
    get strokeStyle() { return strokeStyle as string; },
    set strokeStyle(v: string) { strokeStyle = v; },
    fillStyle: '#000000',
    lineCap: 'butt' as CanvasLineCap,
    lineJoin: 'miter' as CanvasLineJoin,
    globalAlpha: 1,
    save() {}, restore() {},
    beginPath() { pending = null; }, closePath() {},
    moveTo(x: number, y: number) { cur = { x, y }; },
    lineTo(x: number, y: number) { pending = { x1: cur.x, y1: cur.y, x2: x, y2: y }; cur = { x, y }; },
    stroke() {
      if (pending) strokes.push({ ...pending, width: lineWidth, color: String(strokeStyle) });
    },
    fill() {}, fillRect() {}, strokeRect() {}, clearRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setTransform() {}, transform() {}, resetTransform() {},
    setLineDash() {}, getLineDash() { return []; },
    drawImage() {}, arc() {}, arcTo() {}, ellipse() {},
    quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    createPattern() { return null; },
    measureText: () => ({ width: 0, fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2 } as TextMetrics),
    fillText() {}, strokeText() {},
    font: '10px sans-serif',
    textAlign: 'left' as CanvasTextAlign,
    textBaseline: 'alphabetic' as CanvasTextBaseline,
    direction: 'ltr' as CanvasDirection,
    letterSpacing: '0px',
    globalCompositeOperation: 'source-over' as GlobalCompositeOperation,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, strokes };
}

const EMU = 12700; // 1 pt

// A stroke helper. width defaults to 1pt; colour is 6-hex without '#'.
const ln = (over: Partial<Stroke> = {}): Stroke => ({ color: '000000', width: EMU, ...over });

function cell(over: Partial<TableCell> = {}): TableCell {
  return {
    textBody: null, fill: null,
    borderL: null, borderR: null, borderT: null, borderB: null,
    diagonalTL: null, diagonalTR: null,
    gridSpan: 1, rowSpan: 1, hMerge: false, vMerge: false,
    ...over,
  } as TableCell;
}

function tableOf(rows: TableCell[][], cols: number[]): TableElement {
  const rowH = 20 * EMU;
  const tableRows: TableRow[] = rows.map((cells) => ({ height: rowH, cells }));
  return {
    type: 'table', x: 0, y: 0,
    width: cols.reduce((a, b) => a + b, 0), height: rowH * rows.length,
    rotation: 0, flipH: false, flipV: false,
    cols, rows: tableRows,
  };
}

// scale=1 ⇒ emuToPx is identity; a 60pt (60*EMU) column ⇒ 60px wide. Two 60pt
// columns put the shared vertical gridline at x=60. Rows are 20pt ⇒ shared
// horizontal line at y=20.
const SCALE = 1 / EMU; // makes 1 EMU → (1/EMU)px; so 60*EMU → 60px, widths in EMU → px
const COL = 60 * EMU;

function render(t: TableElement): StrokeSeg[] {
  const { ctx, strokes } = makeRecordingCtx();
  renderTable(ctx, t, SCALE, undefined, { themeMajorFont: null, themeMinorFont: null, dpr: 1 });
  return strokes;
}

function verticalAt(strokes: StrokeSeg[], x: number): StrokeSeg[] {
  return strokes.filter((s) => Math.abs(s.x1 - s.x2) < 0.5 && Math.abs(s.x1 - x) <= 1);
}
function horizontalAt(strokes: StrokeSeg[], y: number): StrokeSeg[] {
  return strokes.filter((s) => Math.abs(s.y1 - s.y2) < 0.5 && Math.abs(s.y1 - y) <= 1);
}

// applyStroke sets ctx.strokeStyle via hexToRgba, so a recorded segment colour is
// `rgba(r,g,b,a)`. Convert a 6-hex authoring colour to that form for comparison.
function rgba(hex: string): string {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  return `rgba(${r},${g},${b},1)`;
}

describe('DrawingML <a:tbl> — shared interior gridline drawn once (spec-silent)', () => {
  it('shared VERTICAL gridline is drawn exactly ONCE (not once per cell)', () => {
    // Left cell right = 1pt; right cell left = 1pt (same). Previously TWO strokes
    // at x=60 (one per cell); now exactly one.
    const strokes = render(tableOf([
      [cell({ borderR: ln() }), cell({ borderL: ln() })],
    ], [COL, COL]));
    expect(verticalAt(strokes, 60)).toHaveLength(1);
  });

  it('wider line wins the shared gridline (rule #1), ignoring paint order', () => {
    // Left cell right = 1pt red; right cell left = 3pt blue ⇒ blue (wider) wins.
    const strokes = render(tableOf([
      [cell({ borderR: ln({ width: EMU, color: 'ff0000' }) }),
       cell({ borderL: ln({ width: 3 * EMU, color: '0000ff' }) })],
    ], [COL, COL]));
    const shared = verticalAt(strokes, 60);
    expect(shared).toHaveLength(1);
    expect(shared[0].color).toBe(rgba('0000ff')); // wider blue wins
    expect(Math.abs(shared[0].width - 3)).toBeLessThan(0.6); // winner's 3pt width
  });

  it('null on one side leaves the OTHER side visible (rule #0)', () => {
    // Left cell right = null (no line); right cell left = 1pt. The real line shows.
    const strokes = render(tableOf([
      [cell({ borderR: null }), cell({ borderL: ln({ color: '112233' }) })],
    ], [COL, COL]));
    const shared = verticalAt(strokes, 60);
    expect(shared).toHaveLength(1);
    expect(shared[0].color).toBe(rgba('112233'));
  });

  it('null on BOTH sides suppresses the shared gridline entirely', () => {
    const strokes = render(tableOf([
      [cell({ borderR: null }), cell({ borderL: null })],
    ], [COL, COL]));
    expect(verticalAt(strokes, 60)).toHaveLength(0);
  });

  it('shared HORIZONTAL gridline resolves between stacked rows (wider wins)', () => {
    // Row0 cell bottom = 1pt red; row1 cell top = 3pt green ⇒ green wins. Shared
    // line at y=20 (rows 20pt tall).
    const strokes = render(tableOf([
      [cell({ borderB: ln({ width: EMU, color: 'ff0000' }) })],
      [cell({ borderT: ln({ width: 3 * EMU, color: '00ff00' }) })],
    ], [COL]));
    const shared = horizontalAt(strokes, 20);
    expect(shared).toHaveLength(1);
    expect(shared[0].color).toBe(rgba('00ff00')); // wider green wins
  });

  it('outer table edges are unaffected (non-regression: each drawn once)', () => {
    // Single-cell table with all four outer borders → 4 edges, each once.
    const strokes = render(tableOf([
      [cell({ borderT: ln(), borderB: ln(), borderL: ln(), borderR: ln() })],
    ], [COL]));
    expect(verticalAt(strokes, 0)).toHaveLength(1);   // outer left
    expect(verticalAt(strokes, 60)).toHaveLength(1);  // outer right
    expect(horizontalAt(strokes, 0)).toHaveLength(1); // outer top
    expect(horizontalAt(strokes, 20)).toHaveLength(1);// outer bottom
  });

  it('gridSpan: a horizontally-merged cell spans the shared vertical line region', () => {
    // Row0: one cell spanning both columns (gridSpan=2) + hMerge continuation.
    // Row1: two separate cells sharing a vertical line at x=60.
    // The spanning cell has NO interior vertical line inside it (x=60 within the
    // merged cell must not be drawn from row0), and row1's shared line at x=60 is
    // drawn once. Also the horizontal line between the rows at y=20 under x∈[0,60]
    // and x∈[60,120] must not double up where row0's single wide cell meets the
    // two row1 cells.
    const strokes = render(tableOf([
      [cell({ gridSpan: 2, borderT: ln(), borderB: ln() }), cell({ hMerge: true })],
      [cell({ borderR: ln(), borderB: ln() }), cell({ borderL: ln(), borderB: ln() })],
    ], [COL, COL]));
    // Row1 interior vertical shared line at x=60 drawn exactly once.
    expect(verticalAt(strokes, 60)).toHaveLength(1);
  });

  it('rtl (§21.1.3.13): shared vertical gridline still drawn once with the winner', () => {
    // A right-to-left table places logical column 0 at the physical RIGHT. Two
    // cells, logical [a, b]: a is physically RIGHT (x=60..120), b physically LEFT
    // (x=0..60), so the shared interior line is at x=60. At that line a's PHYSICAL-
    // left edge is its logical-RIGHT border (a.borderR), and b's PHYSICAL-right
    // edge is its logical-LEFT border (b.borderL) — these face each other.
    // a.borderR=1pt vs b.borderL=3pt ⇒ the wider (b's) wins, drawn once.
    const t = tableOf([
      [cell({ borderR: ln({ width: EMU, color: 'ff0000' }) }),
       cell({ borderL: ln({ width: 3 * EMU, color: '0000ff' }) })],
    ], [COL, COL]);
    t.rtl = true;
    const strokes = render(t);
    const shared = verticalAt(strokes, 60);
    expect(shared).toHaveLength(1);
    expect(shared[0].color).toBe(rgba('0000ff')); // wider (b's borderR) wins
    // Outer physical edges (x=0 = b's physical left / logical right side of grid,
    // x=120 = a's physical right) carry no authored border here, so none drawn.
    expect(verticalAt(strokes, 0)).toHaveLength(0);
    expect(verticalAt(strokes, 120)).toHaveLength(0);
  });

  it('rowSpan: a vertically-merged cell spans the shared horizontal line region', () => {
    // Col0 row0 rowSpan=2 (+ vMerge continuation in row1); col1 has two cells that
    // share a horizontal line at y=20. The merged cell must not draw an interior
    // horizontal line at y=20 within its own footprint.
    const strokes = render(tableOf([
      [cell({ rowSpan: 2, borderR: ln() }), cell({ borderB: ln() })],
      [cell({ vMerge: true }), cell({ borderT: ln() })],
    ], [COL, COL]));
    // Col1's shared horizontal line at y=20 (x in [60,120]) drawn exactly once.
    const shared = horizontalAt(strokes, 20).filter((s) => Math.min(s.x1, s.x2) >= 59);
    expect(shared).toHaveLength(1);
  });
  it('#824: a gridSpan cell resolves EACH below sub-segment against its own neighbour', () => {
    // A cell spanning both columns (gridSpan=2, no bottom border) faces TWO below
    // cells with DIFFERENT top borders. The shared horizontal edge must be split at
    // the column boundary: left half → the left neighbour's border, right half →
    // the right neighbour's. Before the fix only the span-origin (left) neighbour
    // was consulted and the right half's border was dropped.
    const strokes = render(tableOf([
      [cell({ gridSpan: 2 }), cell({ hMerge: true })],
      [cell({ borderT: ln({ width: EMU, color: '111111' }) }),
       cell({ borderT: ln({ width: 3 * EMU, color: '222222' }) })],
    ], [COL, COL]));
    const seg = horizontalAt(strokes, 20);
    const left = seg.filter((s) => Math.min(s.x1, s.x2) < 30);
    const right = seg.filter((s) => Math.max(s.x1, s.x2) > 90);
    expect(left.some((s) => s.color === rgba('111111'))).toBe(true);
    expect(right.some((s) => s.color === rgba('222222'))).toBe(true);
  });

  it('#824: a rowSpan cell resolves EACH right sub-segment against its own neighbour', () => {
    // A cell spanning both rows (rowSpan=2, no right border) faces TWO right
    // neighbours with DIFFERENT left borders. Its shared vertical edge must be split
    // at the row boundary: top half → the upper neighbour's border, bottom half →
    // the lower neighbour's. Before the fix only the span-origin (upper) neighbour
    // was consulted and the bottom half's border was dropped.
    const strokes = render(tableOf([
      [cell({ rowSpan: 2 }), cell({ borderL: ln({ width: EMU, color: '111111' }) })],
      [cell({ vMerge: true }), cell({ borderL: ln({ width: 3 * EMU, color: '222222' }) })],
    ], [COL, COL]));
    const seg = verticalAt(strokes, 60);
    const top = seg.filter((s) => Math.min(s.y1, s.y2) < 10);
    const bot = seg.filter((s) => Math.max(s.y1, s.y2) > 30);
    expect(top.some((s) => s.color === rgba('111111'))).toBe(true);
    expect(bot.some((s) => s.color === rgba('222222'))).toBe(true);
  });

});
