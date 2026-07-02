import { describe, it, expect } from 'vitest';
import { renderShapeText } from './renderer.js';
import type { ShapeRun, ShapeText } from './types';

// ECMA-376 §17.3.1.33 / §17.3.2.26 — a text-box (txbxContent) run's single-line
// box must be floored to the DESIGN line height of whichever declared face
// renders its glyphs. The common Japanese encoding sets a substituted CJK face
// (Meiryo, win ratio 3269/2048 = 1.5962×em) ONLY on `<w:rFonts w:eastAsia>`
// while `<w:rFonts w:ascii>` stays an UNTABLED Latin default. Before the fix the
// shape/textbox measure pass floored on the ascii face alone
// (`intendedSingleLinePx(untabledAscii) = 0`), so the line box stayed flat at the
// substituted font's natural box (this mock: 1.0×em) instead of growing to
// Meiryo's 1.5962×em. This mirrors the xlsx shape-text floor (PR #646) and the
// docx BODY per-eastAsia-segment floor.
//
// The floor is asserted on the LINE-BOX height, which in the draw pass equals
// the vertical delta between consecutive wrapped lines' baselines
// (`cursorY += lineH; baseline = cursorY + baselineOffset`). Two wrapped lines
// with the same run give one such delta = the first line's lineH.

interface FillTextEvent { text: string; x: number; y: number }

/** A recording 2D-context whose substituted-font natural box is a flat 1.0×em
 *  (ascent 0.8 + descent 0.2). No family is in the metric table via measureText,
 *  so any design-line growth must come from `intendedSingleLinePx` (the floor
 *  under test), not from the mock's own metrics. */
function makeRecordingCanvas(): { ctx: CanvasRenderingContext2D; fillTexts: FillTextEvent[] } {
  let font = '11px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '11');
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

/** A no-fill/no-line text box holding one rich block whose single run has the
 *  given ascii + eastAsia faces. The run text is wide enough to wrap into ≥2
 *  lines in a narrow box (the mock's measureText is chars × px). */
function textboxWith(fontFamily: string | null, fontFamilyEastAsia: string | null | undefined): ShapeRun {
  const block: ShapeText = {
    text: 'aa bb cc dd ee ff gg hh',
    fontSizePt: 20,
    alignment: 'left',
    runs: [
      { text: 'aa bb cc dd ee ff gg hh', fontSizePt: 20, fontFamily, fontFamilyEastAsia },
    ],
  } as ShapeText;
  return {
    type: 'shape',
    zOrder: 0, subpaths: [], presetGeometry: 'rect',
    fill: null, stroke: null,
    behindDoc: false,
    wrapMode: 'none',
    widthPt: 120, heightPt: 400,
    textBlocks: [block], textAnchor: 't',
    textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
  } as unknown as ShapeRun;
}

/** Vertical delta between the first two DISTINCT baseline y values among the
 *  recorded fillText events = the first wrapped line's line-box height. */
function firstLineHeight(fillTexts: FillTextEvent[]): number {
  const ys = [...new Set(fillTexts.map((f) => f.y))].sort((a, b) => a - b);
  expect(ys.length).toBeGreaterThanOrEqual(2);
  return ys[1] - ys[0];
}

describe('textbox line-box floors on the eastAsia face (ECMA-376 §17.3.2.26)', () => {
  const scale = 1;
  const emPx = 20 * scale; // fontSizePt=20, scale=1
  const MEIRYO_RATIO = 3269 / 2048; // 1.5962…, WIN_METRICS Meiryo win sum
  const NATURAL_RATIO = 1.0;        // mock substituted-font box (0.8 + 0.2)

  it('grows the line box to Meiryo when Meiryo is only on the eastAsia axis (untabled ascii)', () => {
    const { ctx, fillTexts } = makeRecordingCanvas();
    // Untabled ascii ('Calibri' is NOT in WIN_METRICS) + Meiryo on eastAsia.
    renderShapeText(textboxWith('Calibri', 'Meiryo'), 0, 0, 120, 400, ctx, scale);
    const lineH = firstLineHeight(fillTexts);
    // Floored to Meiryo's design line, NOT the flat 1.0×em natural box.
    expect(lineH).toBeCloseTo(MEIRYO_RATIO * emPx, 3);
    expect(lineH).toBeGreaterThan(NATURAL_RATIO * emPx + 0.5);
  });

  it('stays flat for an untabled ascii run with NO eastAsia face (zero regression)', () => {
    const { ctx, fillTexts } = makeRecordingCanvas();
    // Untabled ascii, no eastAsia axis: intendedSingleLinePx returns 0 for both,
    // so the line box keeps the substituted-font natural box (1.0×em) — proving
    // the change is a FLOOR, not a replace.
    renderShapeText(textboxWith('Calibri', undefined), 0, 0, 120, 400, ctx, scale);
    const lineH = firstLineHeight(fillTexts);
    expect(lineH).toBeCloseTo(NATURAL_RATIO * emPx, 3);
  });
});
