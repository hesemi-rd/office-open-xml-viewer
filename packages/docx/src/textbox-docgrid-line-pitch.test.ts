import { describe, it, expect } from 'vitest';
import { measureShapeTextAutoFitHeight } from './renderer.js';
import { acquireAndPaintShapeTextBox } from './retained-shape-textbox.test-support.js';
import { shapeRenderState } from './line-layout.js';
import type { RenderState } from './renderer.js';
import type { ShapeRun, ShapeText } from './types';

// ECMA-376 §17.6.5 `<w:docGrid w:type="lines" w:linePitch>` — a document with a
// line grid snaps EACH text line to the grid pitch (auto / single spacing). The
// BODY renderer already routes this through `lineBoxHeight`'s grid argument, but
// text INSIDE a DrawingML text box (`<wps:txbx><w:txbxContent>`) was measured
// with `grid=undefined`, so its lines stayed at the font's natural single-line
// height instead of the grid pitch. In a Japanese template (pitch = 360 twips =
// 18 pt) a 10 pt run's natural Yu Mincho line box (~1.60 em ≈ 16 pt) rendered
// ~2 pt SHORT per line, tightening every text-box line and — under the bodyPr
// `anchor="ctr"` — leaving a large empty band above the (now-shorter) content.
//
// Word applies the same section line grid to text-box content: the reference PDF
// of the reported private fixture puts every text-box line on an exact 18 pt
// pitch (measured by pdftotext -bbox / pdftoppm), matching the body. These tests
// lock the text-box path onto the grid via the SAME `lineBoxHeight` cell-snap the
// body uses, and prove it is inert when the section declares no line grid.

interface FillTextEvent { text: string; x: number; y: number }

/** A recording 2D context whose substituted-font natural box is a flat 1.0×em
 *  (ascent 0.8 + descent 0.2). No family is in the metric table, so any line-box
 *  growth beyond 1.0×em must come from the grid snap under test. */
function makeRecordingCanvas(): { ctx: CanvasRenderingContext2D; fillTexts: FillTextEvent[] } {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const fillTexts: FillTextEvent[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setLineDash() {}, drawImage() {}, clearRect() {},
    arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    fillRect() {}, strokeRect() {},
    fillText(text: string, x: number, y: number) { fillTexts.push({ text, x, y }); },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, fillTexts };
}

/** A no-fill/no-line 10 pt text box whose single EA block wraps into ≥2 lines in
 *  a narrow box (mock measureText is chars × px). The text is CJK so the line is
 *  classified East Asian for docGrid cell rounding (EAST_ASIAN_RE). The family
 *  is deliberately NOT in the core metric table (游明朝 is, since issue #1013)
 *  so the natural box stays the mock's flat 1.0×em — the grid snap alone
 *  accounts for any growth. */
function eaTextbox(): ShapeRun {
  const text = 'あいうえおかきくけこさしすせそたちつてと';
  const block: ShapeText = {
    text,
    fontSizePt: 10,
    alignment: 'left',
    runs: [{ text, fontSizePt: 10, fontFamily: 'テスト明朝', fontFamilyEastAsia: 'テスト明朝' }],
  } as ShapeText;
  return {
    type: 'shape',
    zOrder: 0, subpaths: [], presetGeometry: 'rect',
    fill: null, stroke: null,
    behindDoc: false, wrapMode: 'none',
    widthPt: 60, heightPt: 400,
    textBlocks: [block], textAnchor: 't',
    textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
  } as unknown as ShapeRun;
}

/** State carrying the section docGrid (like the production render threads). */
function stateWithGrid(
  ctx: CanvasRenderingContext2D,
  grid: { type: string | null; linePitchPt: number | null } | undefined,
): RenderState {
  const base = shapeRenderState(ctx, 1, {}, new Map());
  return { ...base, docGrid: grid } as unknown as RenderState;
}

/** Vertical delta between the first two DISTINCT baseline y values = the first
 *  wrapped line's line-box height. */
function firstLineHeight(fillTexts: FillTextEvent[]): number {
  const ys = [...new Set(fillTexts.map((f) => f.y))].sort((a, b) => a - b);
  expect(ys.length).toBeGreaterThanOrEqual(2);
  return ys[1] - ys[0];
}

/** A text box whose every run carries a ruby annotation (§17.3.3.25), base
 *  18 pt on an 18 pt pitch. Three 3-char ruby runs — the mock measureText is
 *  chars × px, so each 3 × 18 pt run exactly fills the 60 pt width and lands on
 *  its own line, and every line carries a ruby-bearing segment (layoutLines
 *  attaches the annotation to a run's first segment). */
function rubyTextbox(): ShapeRun {
  const mkRun = (text: string) => ({
    text, fontSizePt: 18, fontFamily: '游明朝', fontFamilyEastAsia: '游明朝',
    ruby: { text: 'るび', fontSizePt: 9 },
  });
  const block: ShapeText = {
    text: 'あいうかきくさしす',
    fontSizePt: 18,
    alignment: 'left',
    runs: [mkRun('あいう'), mkRun('かきく'), mkRun('さしす')],
  } as ShapeText;
  return {
    type: 'shape',
    zOrder: 0, subpaths: [], presetGeometry: 'rect',
    fill: null, stroke: null,
    behindDoc: false, wrapMode: 'none',
    widthPt: 60, heightPt: 400,
    textBlocks: [block], textAnchor: 't',
    textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
  } as unknown as ShapeRun;
}

describe('text-box lines snap to the section docGrid line pitch (ECMA-376 §17.6.5)', () => {
  const PITCH = 18;      // 360 twips
  const NATURAL = 10;    // mock 1.0×em at 10 pt

  it('snaps each EA line to the grid pitch (was natural, too tight)', () => {
    const { ctx, fillTexts } = makeRecordingCanvas();
    acquireAndPaintShapeTextBox(eaTextbox(), 0, 0, 60, 400, ctx, 1, {}, new Map(),
      stateWithGrid(ctx, { type: 'lines', linePitchPt: PITCH }));
    // Each line occupies exactly one grid cell (18 pt), NOT the 10 pt natural box.
    expect(firstLineHeight(fillTexts)).toBeCloseTo(PITCH, 3);
  });

  it('is inert when the section declares no line grid (natural spacing preserved)', () => {
    const { ctx, fillTexts } = makeRecordingCanvas();
    acquireAndPaintShapeTextBox(eaTextbox(), 0, 0, 60, 400, ctx, 1, {}, new Map(),
      stateWithGrid(ctx, { type: 'default', linePitchPt: null }));
    expect(firstLineHeight(fillTexts)).toBeCloseTo(NATURAL, 3);
  });

  it('measureShapeTextAutoFitHeight totals the grid-snapped line heights', () => {
    const { ctx } = makeRecordingCanvas();
    const shape = eaTextbox();
    const gridState = stateWithGrid(ctx, { type: 'lines', linePitchPt: PITCH });
    const flatState = stateWithGrid(ctx, { type: 'default', linePitchPt: null });
    const hGrid = measureShapeTextAutoFitHeight(shape, 60, ctx, 1, {}, new Map(), gridState);
    const hFlat = measureShapeTextAutoFitHeight(shape, 60, ctx, 1, {}, new Map(), flatState);
    // The grid total is a whole number of 18 pt cells; the flat total is the same
    // line count at 10 pt — so the grid path is strictly taller by 8 pt / line.
    const lineCount = Math.round(hFlat / NATURAL);
    expect(lineCount).toBeGreaterThanOrEqual(2);
    expect(hGrid).toBeCloseTo(lineCount * PITCH, 3);
    expect(hGrid).toBeGreaterThan(hFlat + 0.5);
  });

  // §17.3.3.25 ruby in a text box flows through the SAME shared line engine as
  // body ruby, so a ruby line must take lineBoxHeight's ruby branch (measured
  // glyph box), NOT the plain-EA design-height cell count. The ruby box keeps
  // 游明朝 (tabled since issue #1013): its 18 pt design line is 25.79 pt, so
  // the plain-EA rule would round to ceil(25.79/18) = 2 cells = 36 pt — while
  // the ruby branch keeps the measured 1.0×em glyph box, exactly one 18 pt
  // cell. `line.hasRuby` (built by layoutLines from the run's ruby annotation)
  // must reach lineBoxHeight in BOTH text-box paths.
  it('keeps a ruby line on its measured glyph box (1 cell), not the design cell count', () => {
    const { ctx, fillTexts } = makeRecordingCanvas();
    acquireAndPaintShapeTextBox(rubyTextbox(), 0, 0, 60, 400, ctx, 1, {}, new Map(),
      stateWithGrid(ctx, { type: 'lines', linePitchPt: PITCH }));
    // Ruby annotations draw at their own y; measure the BASE lines only (18 px
    // mock font). Base baselines are 18 pt apart (1 cell), not 36 (2 cells).
    const baseYs = [...new Set(
      fillTexts.filter((f) => !/[るび]/.test(f.text)).map((f) => f.y),
    )].sort((a, b) => a - b);
    expect(baseYs.length).toBeGreaterThanOrEqual(2);
    expect(baseYs[1] - baseYs[0]).toBeCloseTo(PITCH, 3);
  });

  it('measureShapeTextAutoFitHeight totals ruby lines at the measured glyph box (1 cell each)', () => {
    const { ctx } = makeRecordingCanvas();
    const shape = rubyTextbox();
    const gridState = stateWithGrid(ctx, { type: 'lines', linePitchPt: PITCH });
    const hGrid = measureShapeTextAutoFitHeight(shape, 60, ctx, 1, {}, new Map(), gridState);
    // rubyTextbox lands each of its three runs on its own line (3 × 18 pt runs
    // exactly fill the 60 pt width). Each ruby line is ONE 18 pt cell (its
    // measured 1.0×em glyph box), not the 2 cells the 游明朝 design height
    // (25.79 pt) would claim through the plain-EA rule.
    expect(hGrid).toBeCloseTo(3 * PITCH, 3);
  });
});
