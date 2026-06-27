import { describe, it, expect } from 'vitest';
import { drawShapeText } from './renderer.js';
import type { ShapeParagraph, ShapeText, ShapeTextRun } from './types.js';

// ECMA-376 §21.1.2.1 (text body) / §21.1.2.2 (paragraph): an EMPTY paragraph in
// a shape's text body — and a blank line produced by a standalone/trailing
// `<a:br>` — still occupies ONE single-line height, exactly as tall as a
// one-character line of the same font would be. `drawShapeText` only raised the
// per-line height when it saw a text/math run (`lineHeight = max(lineHeight,
// pxSize * 1.2)`), so a line with no segment reserved ZERO height: the text
// block under-measured and vertical anchoring (`anchor` 'ctr'/'b') and total
// height drifted whenever the body contained blank lines. This is the xlsx
// analog of the docx fix in PR #582 (empty paragraph-mark lines).
//
// The shape path derives line height analytically (font size × 1.2), not from
// font metrics, so the bug is visible without an asymmetric font stub: a fixed
// empty line should contribute the SAME 1.2-em height as a text line of the
// paragraph's default font size.

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

/** A paragraph with one text run, or — when `text` is `null` — an EMPTY one. */
function para(text: string | null): ShapeParagraph {
  return { align: 'l', runs: text === null ? [] : [textRun(text)] };
}

function shapeOf(paragraphs: ShapeParagraph[]): ShapeText {
  // Top-anchored + wrap:none isolates the line-height contribution: each line's
  // baseline is the cumulative sum of the heights of the lines above it, so the
  // A→B gap reports exactly how much the middle paragraph reserved.
  return { anchor: 't', wrap: 'none', paragraphs };
}

function baselines(paragraphs: ShapeParagraph[]): FillTextCall[] {
  const { ctx, calls } = makeRecordingCtx();
  drawShapeText(ctx, shapeOf(paragraphs), 300, 300, 1);
  return calls;
}

describe('empty paragraph line height in shape text bodies (§21.1.2.1 / §21.1.2.2)', () => {
  it('an empty paragraph reserves the SAME single-line height as a text paragraph', () => {
    // Reference: [A, M, B] — three single-line text paragraphs. The A→B baseline
    // gap spans one full intervening line (M).
    const ref = baselines([para('A'), para('M'), para('B')]);
    const refA = ref.find((c) => c.text === 'A');
    const refB = ref.find((c) => c.text === 'B');
    expect(refA).toBeDefined();
    expect(refB).toBeDefined();
    const refGap = refB!.y - refA!.y;

    // Subject: [A, <empty>, B] — the middle paragraph has no runs. Its line must
    // reserve the same single-line height, so the A→B gap is identical.
    const subj = baselines([para('A'), para(null), para('B')]);
    const subjA = subj.find((c) => c.text === 'A');
    const subjB = subj.find((c) => c.text === 'B');
    expect(subjA).toBeDefined();
    expect(subjB).toBeDefined();
    const subjGap = subjB!.y - subjA!.y;

    // Before the fix the empty line reserved 0 px, so subjGap was short by one
    // full line height (11 pt × 1.2) versus the reference.
    expect(subjGap).toBeCloseTo(refGap, 5);
  });
});
