import { describe, it, expect } from 'vitest';
import { distributeLineSlack, isCJKCodePoint } from './text-distribute.js';
import type { DistributeSeg, DistributeResult } from './text-distribute.js';

// ECMA-376 §17.18.44 ST_Jc: `both` and `distribute` fill a line to the margin by
// adding equal slack at every gap opportunity. The clause "both affects only
// inter-word spacing, not inter-character within a word" is LATIN-scoped — in CJK
// each ideograph is its own word, so inter-CJK boundaries are inter-word gaps.
// The Word PDF for private/sample-9 confirms it: a wrapped pure-CJK `both` line
// with zero ASCII spaces still fills to the right margin (xMax 523.0pt via
// pdftotext -bbox), which is only possible by widening inter-CJK boundaries.
//
// distributeLineSlack is the pure kernel: given a line's segments (logical order),
// its slack, and the fixed-edge indices the renderer already computes, it returns
// per-segment split points + a trailing-gap flag and the per-gap px.

const seg = (text?: string): DistributeSeg => ({ text });

/** Total px the kernel will add across the whole line = Σ over segments of
 *  (internal gaps + trailing gap) × perGap. Must equal `slack` so the final
 *  glyph lands on the right margin. */
function totalStretch(r: DistributeResult): number {
  let gaps = 0;
  for (const s of r.perSeg.values()) {
    gaps += s.splitBefore.length + (s.trailingGap ? 1 : 0);
  }
  return gaps * r.perGap;
}

describe('isCJKCodePoint', () => {
  it('classifies CJK ideographs, kana, hangul, fullwidth as CJK', () => {
    expect(isCJKCodePoint('観'.codePointAt(0)!)).toBe(true); // CJK Unified
    expect(isCJKCodePoint('す'.codePointAt(0)!)).toBe(true); // Hiragana
    expect(isCJKCodePoint('カ'.codePointAt(0)!)).toBe(true); // Katakana
    expect(isCJKCodePoint('가'.codePointAt(0)!)).toBe(true); // Hangul syllable
    expect(isCJKCodePoint('、'.codePointAt(0)!)).toBe(true); // CJK punctuation
    expect(isCJKCodePoint('Ａ'.codePointAt(0)!)).toBe(true); // Fullwidth Latin A
  });
  it('classifies Latin letters, digits, ASCII space as non-CJK', () => {
    expect(isCJKCodePoint('A'.codePointAt(0)!)).toBe(false);
    expect(isCJKCodePoint('z'.codePointAt(0)!)).toBe(false);
    expect(isCJKCodePoint('5'.codePointAt(0)!)).toBe(false);
    expect(isCJKCodePoint(' '.codePointAt(0)!)).toBe(false);
  });
});

describe('distributeLineSlack — pure CJK (the bug)', () => {
  // A pure-CJK phrase is ONE segment with no ASCII spaces. The old space-only
  // model produced zero stretch; the kernel must open every inter-CJK boundary.
  it('opens an inter-CJK gap at every interior boundary of a single CJK segment', () => {
    const segs = [seg('観察することで')]; // 7 code points → 6 interior boundaries
    const r = distributeLineSlack(segs, 60, 0, 1);
    expect(r).not.toBeNull();
    const s = r!.perSeg.get(0)!;
    // 6 boundaries, all internal (single segment, no trailing gap since the final
    // segment is excluded — here there IS no later segment, so the segment is the
    // only eligible one and its last boundary is interior).
    expect(s.splitBefore).toEqual([1, 2, 3, 4, 5, 6]);
    expect(s.trailingGap).toBe(false);
    expect(r!.perGap).toBeCloseTo(60 / 6, 6);
    expect(totalStretch(r!)).toBeCloseTo(60, 6);
  });

  it('returns null for a single CJK glyph (no interior boundary)', () => {
    expect(distributeLineSlack([seg('観')], 40, 0, 1)).toBeNull();
  });
});

describe('distributeLineSlack — Latin justify is unchanged (regression guard)', () => {
  it('stretches only inter-word ASCII spaces, never inside a Latin word', () => {
    // Two tokens "the " and "quick". The only gap is the space after "the " —
    // its left side (seg 0) is eligible and its right side is the final segment.
    // No split lands inside "the" or "quick".
    const segs = [seg('the '), seg('quick')];
    const r = distributeLineSlack(segs, 20, 0, 1);
    expect(r).not.toBeNull();
    const s = r!.perSeg.get(0)!;
    expect(s.splitBefore).toEqual([]); // no intra-word split
    expect(s.trailingGap).toBe(true); // the inter-word space after "the "
    expect(r!.perSeg.has(1)).toBe(false); // final segment opens no gap
    expect(r!.perGap).toBeCloseTo(20, 6);
    expect(totalStretch(r!)).toBeCloseTo(20, 6);
  });

  it('distributes across multiple inter-word spaces equally', () => {
    // Three tokens; "the " and "lazy " are eligible (final "dog" excluded). Two
    // trailing spaces → two gaps.
    const segs = [seg('the '), seg('lazy '), seg('dog')];
    const r = distributeLineSlack(segs, 30, 0, 2);
    expect(r).not.toBeNull();
    expect(r!.perGap).toBeCloseTo(15, 6);
    expect(r!.perSeg.get(0)!.trailingGap).toBe(true);
    expect(r!.perSeg.get(1)!.trailingGap).toBe(true);
    expect(totalStretch(r!)).toBeCloseTo(30, 6);
  });

  it('never splits a long single Latin word (no gaps → null)', () => {
    expect(distributeLineSlack([seg('supercalifragilistic')], 50, 0, 1)).toBeNull();
  });
});

describe('distributeLineSlack — mixed Latin + CJK', () => {
  // Mirrors private/sample-9 para#16: Latin word then CJK, no spaces inside CJK.
  it('opens gaps at inter-CJK boundaries and at the Latin/CJK boundary', () => {
    // One segment "BZ反応" — 4 code points: B Z 反 応.
    //   B|Z : both Latin → NOT a gap.
    //   Z|反 : right side CJK → gap (split before index 3, the 反).
    //   反|応 : both CJK → gap (split before index 4? no — that's interior, before 応).
    const segs = [seg('BZ反応')];
    const r = distributeLineSlack(segs, 24, 0, 1);
    expect(r).not.toBeNull();
    const s = r!.perSeg.get(0)!;
    // Gaps after Z (offset 1 → split before 2) and after 反 (offset 2 → split before 3).
    expect(s.splitBefore).toEqual([2, 3]);
    expect(s.trailingGap).toBe(false);
    expect(r!.perGap).toBeCloseTo(12, 6);
  });

  it('does not open a gap between two Latin letters even next to CJK', () => {
    // "反BZ応" code points: 反(0) B(1) Z(2) 応(3).
    //   反|B → CJK left → gap after offset 0 → split before index 1.
    //   B|Z → Latin-Latin → no gap.
    //   Z|応 → CJK right → gap after offset 2 → split before index 3.
    const segs = [seg('反BZ応')];
    const r = distributeLineSlack(segs, 20, 0, 1);
    const s = r!.perSeg.get(0)!;
    expect(s.splitBefore).toEqual([1, 3]); // none after B (the Latin-Latin pair)
  });
});

describe('distributeLineSlack — leading 字下げ indent stays fixed', () => {
  it('does not stretch a leading whitespace segment before firstContentSi', () => {
    // seg 0 = ideographic-space indent (fixed), seg 1 = CJK content, seg 2 = final.
    const segs = [seg('　'), seg('観察する'), seg('結果')];
    // firstContentSi = 1 (indent skipped), lastDrawnSi = 2 (final opens no gap).
    const r = distributeLineSlack(segs, 30, 1, 2);
    expect(r).not.toBeNull();
    expect(r!.perSeg.has(0)).toBe(false); // indent segment never stretched
    expect(r!.perSeg.has(2)).toBe(false); // final segment opens no gap
    const s = r!.perSeg.get(1)!;
    // "観察する" = 4 cp: 観|察, 察|す, す|る are internal; る|結 is the boundary
    // into the final segment (trailing). 4 gaps total.
    expect(s.splitBefore).toEqual([1, 2, 3]);
    expect(s.trailingGap).toBe(true);
    expect(r!.perGap).toBeCloseTo(7.5, 6); // 30 / 4
    expect(totalStretch(r!)).toBeCloseTo(30, 6);
  });

  it('skips a leading ideographic space at the START of the content segment', () => {
    // sample-9 para#16 shape: a U+3000 indent fused into the content segment.
    // The leading 　 must NOT open a gap; the first stretchable boundary is
    // inside the CJK content after it.
    const segs = [seg('　観察'), seg('結果')];
    const r = distributeLineSlack(segs, 16, 0, 1);
    expect(r).not.toBeNull();
    const s = r!.perSeg.get(0)!;
    // code points: 　(0) 観(1) 察(2). Leading 　 is trimmed from the content
    // span, so no gap after offset 0; gap 観|察 (after offset 1 → split before 2)
    // and 察|結 (after offset 2 = last → trailing into final).
    expect(s.splitBefore).toEqual([2]);
    expect(s.trailingGap).toBe(true);
    expect(totalStretch(r!)).toBeCloseTo(16, 6);
  });

  it('treats an interior ideographic space as a single inter-word gap', () => {
    // "観　察" → 観(0) 　(1) 察(2). The 　 is one inter-word gap (after offset 1),
    // NOT two CJK boundaries around it. 観|　 is a boundary into whitespace
    // (counted by the space), so only one gap.
    const segs = [seg('観　察結')];
    const r = distributeLineSlack(segs, 10, 0, 1);
    expect(r).not.toBeNull();
    const s = r!.perSeg.get(0)!;
    // gaps: after 　(offset 1 → split before 2, the inter-word gap) and 察|結
    // (offset 2 → split before 3). 観|　 is NOT a gap.
    expect(s.splitBefore).toEqual([2, 3]);
    expect(r!.perGap).toBeCloseTo(5, 6); // 10 / 2
  });
});

describe('distributeLineSlack — final segment opens no gap, Σgaps == slack', () => {
  it('opens no gap WITHIN the final segment but does widen the boundary into it', () => {
    // [観察][結果], justify. The 4 CJK chars spread as 観 _ 察 _ 結 _ 果: three
    // equal gaps, 果 on the margin. Two of those gaps belong to seg 0 (観|察
    // internal, 察|結 trailing-into-final); none belong to seg 1.
    const segs = [seg('観察'), seg('結果')];
    const r = distributeLineSlack(segs, 18, 0, 1);
    expect(r).not.toBeNull();
    expect(r!.perSeg.has(1)).toBe(false); // final segment opens no gap, never split
    const s = r!.perSeg.get(0)!;
    expect(s.splitBefore).toEqual([1]); // 観|察 internal
    expect(s.trailingGap).toBe(true); // 察|結 boundary into the final segment
    expect(r!.perGap).toBeCloseTo(9, 6); // 18 / 2 gaps
    expect(totalStretch(r!)).toBeCloseTo(18, 6); // 果 still reaches the margin
  });

  it('Σgaps equals slack for a mixed multi-segment line', () => {
    const segs = [seg('図 '), seg('観察する'), seg('こと')];
    const r = distributeLineSlack(segs, 40, 0, 2);
    expect(r).not.toBeNull();
    expect(totalStretch(r!)).toBeCloseTo(40, 6);
  });
});

describe('distributeLineSlack — negative slack (compression)', () => {
  it('compresses gaps with a negative perGap, clamped to minPerGap', () => {
    const segs = [seg('観察することで')]; // 6 interior gaps
    // slack -60 → perGap -10, but clamp at -4 → perGap -4.
    const r = distributeLineSlack(segs, -60, 0, 1, -4);
    expect(r).not.toBeNull();
    expect(r!.perGap).toBe(-4);
  });

  it('applies an unclamped negative perGap when within the cap', () => {
    const segs = [seg('観察することで')];
    const r = distributeLineSlack(segs, -12, 0, 1, -10);
    expect(r!.perGap).toBeCloseTo(-2, 6); // -12/6 = -2, above the -10 floor
  });

  it('returns null when |slack| is below the 0.5px noise floor', () => {
    expect(distributeLineSlack([seg('観察')], 0.3, 0, 1)).toBeNull();
    expect(distributeLineSlack([seg('観察')], -0.4, 0, 1)).toBeNull();
  });

  it('with includeCJK=false, compresses only spaces and leaves CJK boundaries alone', () => {
    // Mixed line, negative slack: only the inter-word space gap participates; the
    // inter-CJK boundary is NOT compressed (no glyph overlap). This is the
    // renderer's compression path (slack < 0 → includeCJK=false).
    const segs = [seg('図 '), seg('観察'), seg('end')];
    const r = distributeLineSlack(segs, -10, 0, 2, -Infinity, false);
    expect(r).not.toBeNull();
    // Only the space after "図 " (trailing gap of seg 0). 観|察 NOT opened.
    expect(r!.perSeg.get(0)!.trailingGap).toBe(true);
    expect(r!.perSeg.has(1)).toBe(false);
    expect(r!.perGap).toBeCloseTo(-10, 6); // one gap absorbs all the (negative) slack
  });

  it('with includeCJK=false, returns null for a space-free CJK line', () => {
    // No spaces → no compressible gap → null (the renderer then leaves the line
    // natural, overflowing by the small canvas-vs-Word bias, as before for CJK).
    expect(distributeLineSlack([seg('観察することで')], -30, 0, 1, -Infinity, false)).toBeNull();
  });
});

describe('distributeLineSlack — non-text atoms', () => {
  it('opens a CJK gap against an inline atom edge but not inside it', () => {
    // seg 0 = image atom, seg 1 = CJK, seg 2 = final. atom|観 → CJK right → gap.
    const segs: DistributeSeg[] = [{}, seg('観察'), seg('end')];
    const r = distributeLineSlack(segs, 30, 0, 2);
    expect(r).not.toBeNull();
    // atom gets a trailing gap (boundary atom|観), CJK seg gets one interior gap (観|察).
    expect(r!.perSeg.get(0)!.trailingGap).toBe(true);
    expect(r!.perSeg.get(0)!.splitBefore).toEqual([]);
    expect(r!.perSeg.get(1)!.splitBefore).toEqual([1]);
    expect(totalStretch(r!)).toBeCloseTo(30, 6);
  });
});
