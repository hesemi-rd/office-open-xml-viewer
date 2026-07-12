/**
 * Producer-level pins for the per-line docGrid grid-count height
 * (`LayoutLine.gridCountSingle`) computed by {@link layoutLines} — the value
 * that feeds §17.6.5 East Asian cell counting in {@link lineBoxHeight}.
 *
 * The rule (see addToLine): the max over segments of each run's Word-faithful
 * single-line height, where a tabled EA run counts from its DESIGN height, an
 * untabled EA run from its measured box, and a Latin run does NOT contribute
 * (it is not cell-rounded). This is what keeps a substituted face's over-tall
 * box from inflating the cell count (sample-52) while still letting a genuinely
 * tall untabled East Asian run claim its cells (independent-review mixed case).
 *
 * A linear stub context makes the metrics analytic: measured box = fontSize
 * (0.8 + 0.2 em); Yu Mincho's tabled EA design height = 1.43267 × em.
 */
import { describe, expect, it } from 'vitest';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import { layoutLines, type LayoutSeg, type LayoutTextSeg } from './line-layout.js';

function linearCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = (): number => Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (t: string) => ({
      width: [...t].length * px() * 0.5,
      fontBoundingBoxAscent: px() * 0.8,
      fontBoundingBoxDescent: px() * 0.2,
      actualBoundingBoxAscent: px() * 0.8,
      actualBoundingBoxDescent: px() * 0.2,
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}

function seg(text: string, fontFamily: string, fontSize: number): LayoutTextSeg {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily, vertAlign: null, measuredWidth: 0,
  };
}

function layout(segs: LayoutTextSeg[], width = 400) {
  return layoutLines(
    linearCtx(), segs.map((s) => ({ ...s })) as LayoutSeg[], width, 0, 1,
    undefined, undefined, {}, 0, DEFAULT_KINSOKU_RULES, 0, 36, width, false, false, false,
  );
}

// Yu Mincho tabled EA design single-line height per pt: 1.3 × hhea box = 1.43267 em.
const YU = (2257 * 1.3) / 2048;

describe('layoutLines — per-line gridCountSingle (§17.6.5 docGrid cell height)', () => {
  it('a tabled EA line counts from its DESIGN height, not the (larger) measured box', () => {
    // 12pt Yu Mincho CJK: measured box 12px (linear stub) but design 17.19px.
    const [line] = layout([seg('あいうえお', 'Yu Mincho', 12)]);
    expect(line.gridCountSingle).toBeCloseTo(12 * YU, 4);
    expect(line.gridCountSingle).toBeGreaterThan(12); // above the raw box
  });

  it('EXCLUDES a Latin run on a mixed CJK+Latin line (sample-52 shape)', () => {
    // sample-52's lines mix ASCII and CJK ("TABLE 1 = …（…）"). The ASCII run in
    // Yu Mincho has no design height (eaOnly ⇒ intendedSingle 0), and its
    // SUBSTITUTED Canvas box overshoots the real Latin height — Word measured
    // sample-52's exact table at x0=417 == the CJK-design (1-cell) layout, so
    // the ASCII box must NOT drive the count. This test uses a deliberately
    // TALL 20pt Latin run (box 20px > the 12pt CJK design 17.19px) to
    // DISCRIMINATE the exclusion: were it counted the result would be 20; the
    // §17.6.5 "Latin is not cell-rounded" rule yields the CJK design alone.
    // (The tall-Latin mixed line has no Word ground truth of its own — this
    // characterizes the deliberate rule, consistent with the Latin-only case.)
    const [line] = layout([seg('TABLE 1 = ', 'Yu Mincho', 20), seg('固定幅（両軸拘束）。', 'Yu Mincho', 12)]);
    expect(line.gridCountSingle).toBeCloseTo(12 * YU, 4);
    expect(line.gridCountSingle).toBeLessThan(20); // the 20pt Latin box did NOT count
  });

  it('counts a genuinely TALL untabled EA run over a small tabled run (mixed case)', () => {
    // Small 12pt Yu Mincho (design 17.19) + a large 24pt untabled CJK face
    // (measured box 24). The untabled run is taller and East Asian, so it must
    // claim its extent: max(17.19, 24) = 24. Guards the design-height rule from
    // under-counting a tall untabled run (independent review, Codex gpt-5.6-sol).
    const [line] = layout([seg('小', 'Yu Mincho', 12), seg('大きい', 'PMingLiU', 24)]);
    expect(line.gridCountSingle).toBeCloseTo(24, 4);
  });

  it('a pure-Latin line is NOT cell-rounded — falls back to the box (unused off the EA path)', () => {
    // No EA segment contributes, so gridCountSingle falls back to the line box
    // (== old `natural`). The line is non-EA, so lineBoxHeight never consults it.
    const [line] = layout([seg('hello world', 'Times New Roman', 12)]);
    expect(line.eastAsian ?? false).toBe(false);
    expect(line.gridCountSingle).toBeCloseTo(line.ascent + line.descent, 4);
  });
});
