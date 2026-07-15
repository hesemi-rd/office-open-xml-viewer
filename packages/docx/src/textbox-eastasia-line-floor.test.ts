import { describe, it, expect } from 'vitest';
import { acquireAndPaintShapeTextBox } from './retained-shape-textbox.test-support.js';
import type { ShapeRun, ShapeText, ShapeTextRun } from './types';

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

/** The FIRST (topmost) recorded baseline y. With textAnchor='t' and zero insets
 *  the first line's box top is 0, so this equals that line's baselineOffset
 *  (= half-leading above the glyph box + the ascent). */
function firstBaselineY(fillTexts: FillTextEvent[]): number {
  const ys = [...new Set(fillTexts.map((f) => f.y))].sort((a, b) => a - b);
  expect(ys.length).toBeGreaterThanOrEqual(1);
  return ys[0];
}

/** A no-fill/no-line text box holding one rich block with the given runs. */
function textboxWithRuns(runs: ShapeTextRun[]): ShapeRun {
  const block: ShapeText = {
    text: runs.map((r) => r.text).join(''),
    fontSizePt: runs[0]?.fontSizePt ?? 20,
    alignment: 'left',
    runs,
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

describe('textbox line-box floors on the eastAsia face (ECMA-376 §17.3.2.26)', () => {
  const scale = 1;
  const emPx = 20 * scale; // fontSizePt=20, scale=1
  const MEIRYO_RATIO = 3269 / 2048; // 1.5962…, WIN_METRICS Meiryo win sum
  const NATURAL_RATIO = 1.0;        // mock substituted-font box (0.8 + 0.2)

  it('grows the line box to Meiryo when Meiryo is only on the eastAsia axis (untabled ascii)', () => {
    const { ctx, fillTexts } = makeRecordingCanvas();
    // Untabled ascii ('Calibri' is NOT in WIN_METRICS) + Meiryo on eastAsia.
    acquireAndPaintShapeTextBox(textboxWith('Calibri', 'Meiryo'), 0, 0, 120, 400, ctx, scale);
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
    acquireAndPaintShapeTextBox(textboxWith('Calibri', undefined), 0, 0, 120, 400, ctx, scale);
    const lineH = firstLineHeight(fillTexts);
    expect(lineH).toBeCloseTo(NATURAL_RATIO * emPx, 3);
    // Baseline is ALSO byte-for-byte unchanged: with no floor the glyph box fills
    // the line box (lineH == glyphNatural), so half-leading is 0 and the baseline
    // reduces to c.ascent (= 0.8×em) exactly as before Fix 1 — proving the
    // centering + rendering-face change is inert on an all-untabled line.
    expect(firstBaselineY(fillTexts)).toBeCloseTo(0.8 * emPx, 3);
  });

  // Fix 1 — the glyph ink is CENTERED in the (Meiryo-)inflated line box (real
  // half-leading above the baseline), not top-pinned. Before the fix,
  // baselineOffset folded `intended` into `natural`, so it collapsed to c.ascent
  // (the ink sat at the box top). Body path: baseline = top + (lineH −
  // glyphNatural)/2 + ascent, glyphNatural = ascent+descent NOT floor-inflated.
  it('centers the CJK glyph box in the inflated line box (half-leading, not top-pinned)', () => {
    const { ctx, fillTexts } = makeRecordingCanvas();
    acquireAndPaintShapeTextBox(textboxWith('Calibri', 'Meiryo'), 0, 0, 120, 400, ctx, scale);
    const baselineY = firstBaselineY(fillTexts); // = first line's baselineOffset
    const ascentPx = 0.8 * emPx;                 // mock glyph ascent (0.8×em)
    const glyphNaturalPx = 1.0 * emPx;           // mock glyph box (0.8 + 0.2)
    const lineHPx = MEIRYO_RATIO * emPx;         // floored line box (Fix #648)
    const expectedHalfLeading = (lineHPx - glyphNaturalPx) / 2;
    // Centered: baseline = half-leading + ascent, strictly BELOW the top-pinned
    // c.ascent (the pre-fix value). Half-leading is real and positive.
    expect(expectedHalfLeading).toBeGreaterThan(0.5);
    expect(baselineY).toBeCloseTo(expectedHalfLeading + ascentPx, 3);
    expect(baselineY).toBeGreaterThan(ascentPx + 0.5); // NOT top-pinned (== ascent)
  });

  // Fix 2 — the design-line floor is the MAX over ALL runs on the line, not just
  // the TALLEST run's faces. Here the tallest run (ties → earliest) is an
  // untabled-ascii run; a later EQUAL-size Meiryo-eastAsia run shares line 1.
  // The tallest-only code floored on the untabled run (0) and left line 1 flat;
  // the all-runs max floors line 1 to Meiryo's design line. Mirrors the body's
  // per-segment lineIntendedSingle max.
  it('floors a mixed line to Meiryo when a non-tallest run carries it (all-runs max)', () => {
    const { ctx, fillTexts } = makeRecordingCanvas();
    // 20px/char in the mock, box 120px ⇒ 6 chars/line. Line 1 packs the untabled
    // ascii word 'ab ' (3) + 'あいう' (3) = 6 chars; 'えお' wraps to line 2. Both
    // runs are size 20, so the earliest (untabled ascii) is the "tallest".
    const runs: ShapeTextRun[] = [
      { text: 'ab ', fontSizePt: 20, fontFamily: 'Calibri' }, // untabled, tallest (tie → first)
      { text: 'あいうえお', fontSizePt: 20, fontFamily: null, fontFamilyEastAsia: 'Meiryo' },
    ];
    acquireAndPaintShapeTextBox(textboxWithRuns(runs), 0, 0, 120, 400, ctx, scale);
    const lineH = firstLineHeight(fillTexts);
    // The Meiryo run on line 1 raises the box despite not being the tallest run.
    expect(lineH).toBeCloseTo(MEIRYO_RATIO * emPx, 3);
    expect(lineH).toBeGreaterThan(NATURAL_RATIO * emPx + 0.5);
  });
});
