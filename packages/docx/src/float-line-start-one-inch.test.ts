import { describe, it, expect } from 'vitest';
import {
  resolveLineFloatWindow,
  wordMinLineStartPx,
  WORD_MIN_LINE_START_PT,
  LINE_START_GAP_EPS_PT,
  type FloatRect,
} from './float-layout.js';
import {
  layoutLines,
  type LayoutSeg,
  type LayoutTextSeg,
  type WrapLayoutCtx,
} from './line-layout.js';

// ─────────────────────────────────────────────────────────────────────────────
// Word's measured minimum line-start rule beside a float (issue #676).
//
// GROUND TRUTH (fixtures private/sample-19/20/22, Word-exported PDF, pdftotext
// bbox): Word starts a CONTENT line beside a float ONLY when the free horizontal
// gap is ≥ 72pt (= 1 inch = 1440 twips), and always when it is. For a content
// line the threshold is:
//   - text-independent (a short-token line and a long-word line switch at the
//     same width),
//   - font-size-independent (8/12/24pt switch at 72pt),
//   - line-spacing-independent (single/1.5/double switch at 72pt),
// i.e. an ABSOLUTE width, not an em- or line-height-proportional quantity. At a
// gap of 70pt the line flows below the band; at 72pt it sits beside. A first
// word that overruns the ≥1-inch gap is force-broken there (Word's "AFTE"/"R-10"
// wrap), not refused.
//
// SCOPE — this 1-inch rule is the CONTENT-line threshold. A literally-empty /
// anchor-only paragraph's pilcrow uses the NARROWER pilcrow-em threshold
// (paragraphMarkEmPx via resolveEmptyMarkTop / flowMarkLine): Word keeps such a
// mark beside a float down to a sub-inch gap and drops it below only for a
// full-width band (sample-9 p.4 + sample-12 p.2; the #676 change wrongly
// applied 1 inch to empty marks). The (c) case below exercises the layoutLines
// EMPTY-CONTENT-line path (a paragraph carrying an empty text segment), which is
// a content line and keeps the 1-inch rule.
//
// The boundary is INCLUSIVE at 1 inch, but a frame authored so the gap is
// exactly 1 inch computes as content-width − frame-width slightly under 72
// (sample-22 p.7 → 71.963716pt in this renderer). The callers therefore pass
// `wordMinLineStartPx(scale)` = (72 − LINE_START_GAP_EPS_PT) × scale, a half-twip
// rounding tolerance, so the effective threshold is 71.95pt at scale 1: 70/71.9pt
// stay below, 71.95pt and up (incl. the 71.96pt computed for a 72.0pt frame) go
// beside. See issue #676.
//
// This file pins the pure geometry gate (resolveLineFloatWindow) and the
// layoutLines integration that consumes it. It replaced the first-atomic-token-
// width probe (the former requiredLineWidth) for content lines with this single
// grounded 1-inch rule. The literally-empty-paragraph mark path
// (resolveEmptyMarkTop / flowMarkLine) keeps its own pilcrow-em threshold — see
// the SCOPE note above.
// ─────────────────────────────────────────────────────────────────────────────

/** A LEFT-anchored square float band occupying [0, floatRightPx] horizontally on
 *  rows [0, floatBottomPx). paraX is 0, so with a column of `colWpx` the free
 *  RIGHT gap is `colWpx - floatRightPx`. */
function leftBand(floatRightPx: number, floatBottomPx: number): FloatRect {
  return {
    kind: 'shape', mode: 'square', imageKey: 'x',
    imageX: 0, imageY: 0, imageW: floatRightPx, imageH: floatBottomPx,
    xLeft: 0, xRight: floatRightPx, yTop: 0, yBottom: floatBottomPx,
    side: 'bothSides', distLeft: 0, distRight: 0, distTop: 0, distBottom: 0,
    paraId: 1, drawn: false,
  } as FloatRect;
}

/** Query resolveLineFloatWindow with a given free-gap width (px) at a given
 *  scale, passing EXACTLY what the docx renderer passes for a line-start probe
 *  (`wordMinLineStartPx(scale)`). Returns whether the line was placed BESIDE the
 *  band (topY 0 with a non-zero xOffset) or FLOWED BELOW it (topY advanced past
 *  the band bottom). */
function placeLine(gapPx: number, scale: number): { beside: boolean; topY: number; xOffset: number } {
  const colW = 1000 * scale;
  const floatBottom = 50 * scale;
  const floatRight = colW - gapPx; // leave exactly `gapPx` of free gap on the right
  const win = resolveLineFloatWindow(
    0, wordMinLineStartPx(scale), 10 * scale, 0, colW, [leftBand(floatRight, floatBottom)],
  );
  const beside = win.topY === 0 && win.xOffset > 0;
  return { beside, topY: win.topY, xOffset: win.xOffset };
}

describe('resolveLineFloatWindow — Word 1-inch line-start gate (issue #676)', () => {
  it('the grounded constant is exactly 1 inch (72pt) with a half-twip tolerance', () => {
    expect(WORD_MIN_LINE_START_PT).toBe(72);
    expect(LINE_START_GAP_EPS_PT).toBe(0.05); // half a twip (1 twip = 1/20 pt)
    expect(wordMinLineStartPx(1)).toBeCloseTo(71.95, 10);
    expect(wordMinLineStartPx(2)).toBeCloseTo(143.9, 10);
  });

  it('(a) a 71.9pt gap flows the line BELOW the band (clear of the tolerance band)', () => {
    // 71.9 < 71.95 effective threshold → below. 71.9pt is the largest "below"
    // probe that stays outside the half-twip tolerance (a genuinely sub-inch gap).
    const r = placeLine(71.9, 1);
    expect(r.beside).toBe(false);
    expect(r.topY).toBe(50); // pushed to the band bottom
  });

  it('(b) a 72.0pt gap places the line BESIDE the band (exactly 1 inch)', () => {
    const r = placeLine(72.0, 1);
    expect(r.beside).toBe(true);
    expect(r.topY).toBe(0);
    expect(r.xOffset).toBeGreaterThan(0);
  });

  it('(b) sample-22 p.7: a gap computed at 71.9637pt (a 72.0pt frame) is BESIDE', () => {
    // The exact value this renderer computes for the gap=72.0pt frame — the
    // tolerance exists precisely so this lands beside, matching Word's PDF.
    expect(placeLine(71.963716, 1).beside).toBe(true);
  });

  it('a 70pt gap is below, a 74pt gap is beside (the sample-22 bracket)', () => {
    expect(placeLine(70, 1).beside).toBe(false);
    expect(placeLine(74, 1).beside).toBe(true);
  });

  it('(e) the 70/72pt boundary is identical in PT space at scale 0.75', () => {
    const s = 0.75;
    // 70pt and 72pt gaps expressed in px at this scale must still switch across
    // the 1-inch boundary (requiredWidth is wordMinLineStartPx(scale), so the
    // decision is taken in pt space and is scale-invariant).
    expect(placeLine(70 * s, s).beside).toBe(false);
    expect(placeLine(72 * s, s).beside).toBe(true);
  });

  it('(e) the boundary is identical across scales (absolute pt width)', () => {
    for (const s of [1, 2, 0.5, 1.5, 0.75, 3]) {
      expect(placeLine(70 * s, s).beside).toBe(false); // 70pt < 1 inch → below
      expect(placeLine(72 * s, s).beside).toBe(true);  // 72pt = 1 inch → beside
    }
  });

  it('is a pure width gate: no content input, so font size / empty-vs-filled cannot matter', () => {
    // resolveLineFloatWindow takes only a numeric requiredWidth — there is no
    // content input at all. The gate therefore cannot depend on font size or the
    // line being empty vs. filled; every caller resolves to wordMinLineStartPx.
    // (c)+(d) parity is enforced structurally by the single call site.
    expect(placeLine(71.9, 1).beside).toBe(false);
    expect(placeLine(72.0, 1).beside).toBe(true);
  });
});

// ── layoutLines integration ──────────────────────────────────────────────────
// Linear mock canvas: glyph advance = perPx · px · chars; ascent/descent 0.8/0.2
// em. Perfectly scale-linear so the wrap ALGORITHM is isolated from font hinting.
function makeLinearCtx(perPx = 0.5): CanvasRenderingContext2D {
  let font = '10px serif';
  const pxOf = (): number => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = pxOf();
      const per = p * perPx;
      return {
        width: [...s].length * per,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

function textSeg(text: string, fontSize = 10, extra: Partial<LayoutTextSeg> = {}): LayoutSeg {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'Times New Roman', vertAlign: null,
    measuredWidth: 0, ...extra,
  } as LayoutSeg;
}

function wrapCtx(floats: FloatRect[]): WrapLayoutCtx {
  return {
    startPageY: 0,
    paraX: 0,
    floats,
    lineBoxH: (asc: number, desc: number) => asc + desc,
    pageH: 100000,
  } as WrapLayoutCtx;
}

/** Was the FIRST line placed beside the band (topY 0, xOffset > 0) or flowed
 *  below it (topY past the band bottom)? */
function firstLinePlacement(lines: ReturnType<typeof layoutLines>): 'beside' | 'below' {
  const l = lines[0];
  return l.topY === 0 && l.xOffset > 0 ? 'beside' : 'below';
}

describe('layoutLines — 1-inch line-start rule end to end (issue #676)', () => {
  const scale = 1;
  const colW = 1000;
  const floatBottom = 50;

  // A gap just under 1 inch (70px) and just over (72px) at scale 1.
  const bandFor = (gapPx: number) => [leftBand(colW - gapPx, floatBottom)];

  it('(c) an empty-content CONTENT line flows below a sub-inch gap and beside a ≥1-inch gap', () => {
    // A content paragraph whose sole segment is empty text — the layoutLines
    // content-line path, which keeps the 1-inch rule. (A literally-empty
    // paragraph with NO runs is placed by resolveEmptyMarkTop against the
    // narrower pilcrow-em threshold instead — see SCOPE note.)
    const emptyBelow = layoutLines(makeLinearCtx(), [textSeg('', 10)], colW, 0, scale, [], wrapCtx(bandFor(70)), {}, 0);
    const emptyBeside = layoutLines(makeLinearCtx(), [textSeg('', 10)], colW, 0, scale, [], wrapCtx(bandFor(72)), {}, 0);
    expect(firstLinePlacement(emptyBelow)).toBe('below');
    expect(firstLinePlacement(emptyBeside)).toBe('beside');
  });

  it('(c) a text line makes the SAME below/beside decision as the empty line', () => {
    const textBelow = layoutLines(makeLinearCtx(), [textSeg('hi', 10)], colW, 0, scale, [], wrapCtx(bandFor(70)), {}, 0);
    const textBeside = layoutLines(makeLinearCtx(), [textSeg('hi', 10)], colW, 0, scale, [], wrapCtx(bandFor(72)), {}, 0);
    expect(firstLinePlacement(textBelow)).toBe('below');
    expect(firstLinePlacement(textBeside)).toBe('beside');
  });

  it('(d) the below/beside decision is font-size-independent (8pt vs 24pt agree)', () => {
    for (const fs of [8, 24]) {
      const below = layoutLines(makeLinearCtx(), [textSeg('X', fs)], colW, 0, scale, [], wrapCtx(bandFor(70)), {}, 0);
      const beside = layoutLines(makeLinearCtx(), [textSeg('X', fs)], colW, 0, scale, [], wrapCtx(bandFor(72)), {}, 0);
      expect(firstLinePlacement(below)).toBe('below');
      expect(firstLinePlacement(beside)).toBe('beside');
    }
  });

  it('(d) a SHORT token no longer wedges into a sub-inch gap it would have fit', () => {
    // "X" at 10pt is 5px wide — under the old 1-em (10px) probe it might have
    // been rejected, but a longer prior-behaviour concern was a short token
    // fitting a sub-inch sliver. With the 1-inch rule, a 5px-wide token in a
    // 30px gap (well under 1 inch) is sent below, matching Word.
    const lines = layoutLines(makeLinearCtx(), [textSeg('X', 10)], colW, 0, scale, [], wrapCtx(bandFor(30)), {}, 0);
    expect(firstLinePlacement(lines)).toBe('below');
  });

  it('force-wrap: a word wider than a ≥1-inch gap is CHAR-BROKEN in the gap (Word "AFTE"/"R-10")', () => {
    // Gap = 72px (exactly 1 inch). Word "AFTERTENAFTERTEN" = 16 chars × 5px = 80px,
    // wider than the 72px gap. The line IS started beside the band (gap ≥ 1 inch)
    // and the word is force-broken to fit — it is NOT sent below.
    const lines = layoutLines(
      makeLinearCtx(), [textSeg('AFTERTENAFTERTEN', 10)], colW, 0, scale, [], wrapCtx(bandFor(72)), {}, 0,
    );
    // First line beside the band, holding as many chars as fit the 72px gap.
    expect(firstLinePlacement(lines)).toBe('beside');
    expect(lines.length).toBeGreaterThan(1); // the word was split across lines
    const firstText = (lines[0].segments[0] as LayoutTextSeg).text;
    const secondText = (lines[1].segments[0] as LayoutTextSeg).text;
    // 72px gap / 5px per char = 14 chars fit; the split preserves the whole word.
    expect(firstText.length).toBeGreaterThan(0);
    expect(firstText.length).toBeLessThan('AFTERTENAFTERTEN'.length);
    expect(firstText + secondText).toBe('AFTERTENAFTERTEN');
    // The line sat in the gap (xOffset at the band's right edge), not full width.
    expect(lines[0].xOffset).toBeGreaterThan(0);
    expect(lines[0].availWidth).toBeLessThanOrEqual(72 + 1e-6);
  });

  it('a word narrower than the ≥1-inch gap sits beside the band without splitting', () => {
    // Gap = 200px. "AFTER" = 5 chars × 5px = 25px < 200px → sits beside, no split.
    const lines = layoutLines(makeLinearCtx(), [textSeg('AFTER', 10)], colW, 0, scale, [], wrapCtx(bandFor(200)), {}, 0);
    expect(firstLinePlacement(lines)).toBe('beside');
    expect((lines[0].segments[0] as LayoutTextSeg).text).toBe('AFTER');
  });
});
