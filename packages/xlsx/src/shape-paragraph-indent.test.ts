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
