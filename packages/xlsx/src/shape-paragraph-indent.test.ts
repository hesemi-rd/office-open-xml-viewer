import { describe, it, expect } from 'vitest';
import { drawShapeText } from './renderer.js';
import type { ShapeParagraph, ShapeText, ShapeTextRun } from './types.js';

// ECMA-376 §21.1.2.2.7 (`CT_TextParagraphProperties`): a shape paragraph's
// direct `<a:pPr>` indent attributes — `marL` (left margin), `marR` (right
// margin), `indent` (first-line indent, negative = hanging) — shift the text's
// horizontal layout box. Previously `drawShapeText` dropped them entirely, so
// indented paragraphs drew flush-left at `padX`. The renderer now mirrors the
// pptx consumption: leftInset = marL/EMU_PER_PX*cs (+ first-line indent on the
// first line), and the per-line alignment region width = paraW. All units are
// EMU on the model and converted with EMU_PER_PX = 9525, scaled by `cs`.

const EMU_PER_PX = 9525;

interface FillTextCall {
  text: string;
  x: number;
  y: number;
}

function makeRecordingCtx(): { ctx: CanvasRenderingContext2D; calls: FillTextCall[] } {
  let font = '11px sans-serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '11');
  const calls: FillTextCall[] = [];
  const ctx = {
    get font() {
      return font;
    },
    set font(v: string) {
      font = v;
    },
    measureText: (s: string) => ({ width: [...s].length * px() }) as TextMetrics,
    fillText(text: string, x: number, y: number) {
      calls.push({ text, x, y });
    },
    drawImage() {},
    fillStyle: '#000' as string | CanvasGradient | CanvasPattern,
    textBaseline: 'alphabetic' as CanvasTextBaseline,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, calls };
}

function textRun(text: string, size = 11): ShapeTextRun {
  return { type: 'text', text, bold: false, italic: false, size };
}

function shapeOf(paragraphs: ShapeParagraph[]): ShapeText {
  // wrap:none keeps each paragraph on a single line so the first fillText x is
  // the line's left edge (no wrapping interference).
  return { anchor: 't', wrap: 'none', paragraphs };
}

function drawFirstX(p: ShapeParagraph): number {
  const { ctx, calls } = makeRecordingCtx();
  drawShapeText(ctx, shapeOf([p]), 300, 300, 1);
  expect(calls.length).toBeGreaterThan(0);
  return calls[0].x;
}

describe('shape paragraph indent attributes (§21.1.2.2.7 marL/marR/indent)', () => {
  it('a left-aligned paragraph with marL draws shifted right by marL/EMU_PER_PX*cs', () => {
    const baseline: ShapeParagraph = { align: 'l', runs: [textRun('X')] };
    // marL = 457200 EMU = 0.5 inch = 48 px at EMU_PER_PX=9525, cs=1.
    const indented: ShapeParagraph = { align: 'l', marL: 457200, runs: [textRun('X')] };

    const baseX = drawFirstX(baseline);
    const indentedX = drawFirstX(indented);

    const expectedShift = (457200 / EMU_PER_PX) * 1; // 48 px
    expect(expectedShift).toBeCloseTo(48, 5);
    expect(indentedX - baseX).toBeCloseTo(expectedShift, 5);
  });

  it('a paragraph with no indent attrs is byte-identical (same x) to the baseline', () => {
    // Invariant: when marL/marR/indent are all absent, leftInset=0 and
    // availW=innerW, so the draw x is exactly padX (the pre-change value).
    const noAttrs: ShapeParagraph = { align: 'l', runs: [textRun('X')] };
    const padX = 7 * 1; // padX = 7 * cs in drawShapeText
    expect(drawFirstX(noAttrs)).toBeCloseTo(padX, 5);
  });

  it('first-line indent shifts the first line right by indent on top of marL', () => {
    const marLpx = (457200 / EMU_PER_PX) * 1; // 48 px
    const indentPx = (228600 / EMU_PER_PX) * 1; // 24 px
    const p: ShapeParagraph = {
      align: 'l',
      marL: 457200,
      indent: 228600,
      runs: [textRun('X')],
    };
    const padX = 7 * 1;
    expect(drawFirstX(p)).toBeCloseTo(padX + marLpx + indentPx, 5);
  });
});

// The cases above only exercise wrap:'none' + left-align (first line, left
// branch). These cover the genuinely new per-line machinery: the first-vs-
// continuation `leftInset`/`availW` selection, the wrap budget = paraW (− the
// first-line indent on line 1), `ctr` alignment within the indented region, the
// hanging-indent clamp, and marR narrowing the wrap width.
describe('shape paragraph indent — wrapping, alignment region, hanging, marR', () => {
  const PADX = 7; // padX = 7 * cs (cs=1)
  // All fillText calls for a single paragraph at a given wrap mode (sw=sh=300, cs=1).
  function callsFor(p: ShapeParagraph, wrap: 'none' | 'normal'): FillTextCall[] {
    const { ctx, calls } = makeRecordingCtx();
    drawShapeText(ctx, { anchor: 't', wrap, paragraphs: [p] }, 300, 300, 1);
    return calls;
  }

  it('wrapping: the first line carries the indent, continuation lines carry only marL', () => {
    // indent = 457200 EMU = 48 px shrinks ONLY the first line's wrap budget.
    const indentPx = 457200 / EMU_PER_PX; // 48
    const text = 'X'.repeat(40);
    const without = callsFor({ align: 'l', runs: [textRun(text)] }, 'normal');
    const withInd = callsFor({ align: 'l', indent: 457200, runs: [textRun(text)] }, 'normal');
    expect(withInd.length).toBeGreaterThanOrEqual(2);
    expect(withInd[0].x).toBeCloseTo(PADX + indentPx, 5); // first line shifted by indent
    expect(withInd[1].x).toBeCloseTo(PADX, 5); // continuation: no first-line indent
    // The narrower first-line budget wraps it earlier than the un-indented first line.
    expect(withInd[0].text.length).toBeLessThan(without[0].text.length);
  });

  it('center alignment is within the indented region (marL shifts the center by marL/2)', () => {
    // Centering within paraW (= innerW − marL) instead of full innerW shifts a
    // centered line right by exactly marL/2 (independent of the text width).
    const marLpx = 457200 / EMU_PER_PX; // 48
    const base = callsFor({ align: 'ctr', runs: [textRun('XX')] }, 'none')[0].x;
    const shifted = callsFor({ align: 'ctr', marL: 457200, runs: [textRun('XX')] }, 'none')[0].x;
    expect(shifted - base).toBeCloseTo(marLpx / 2, 5);
  });

  it('a hanging (negative) indent is clamped to 0 — the first line stays at marL', () => {
    // marL=48px, indent=−24px ⇒ firstLineIndent=max(0,−24)=0, so x = padX + marL.
    const x = callsFor({ align: 'l', marL: 457200, indent: -228600, runs: [textRun('X')] }, 'none')[0].x;
    expect(x).toBeCloseTo(PADX + 457200 / EMU_PER_PX, 5);
  });

  it('marR reduces the wrap width (forces an earlier wrap)', () => {
    const text = 'X'.repeat(12);
    const noMarR = callsFor({ align: 'l', runs: [textRun(text)] }, 'normal');
    const bigMarR = callsFor({ align: 'l', marR: 2286000, runs: [textRun(text)] }, 'normal'); // 240px
    expect(noMarR.length).toBe(1); // 12 chars fit the full box on one line
    expect(bigMarR.length).toBeGreaterThan(1); // the narrowed region forces a wrap
  });
});
