import { describe, it, expect } from 'vitest';
import { distributeLineSlack } from './text-distribute.js';
import type { DistributeSeg, DistributeResult } from './text-distribute.js';

// text-distribute.ts is now a thin POSITIONAL adapter over the shared kernel
// `@silurus/ooxml-core` → distributeLineSlack. The exhaustive gap-geometry tests
// live with the kernel (packages/core/src/text/line-distribute.test.ts). These
// tests pin the docx adapter contract the renderer depends on: the positional
// argument order, the includeCJK → compression mapping, the leading-indent skip,
// the bidi exact-match exclusion, and the §17.18.44 pure-CJK fill — i.e. that the
// docx-specific wiring (whitespace = U+0020/U+3000, includeCJK = slack > 0) is
// preserved across the extraction.

const seg = (text?: string): DistributeSeg => ({ text });

/** Σ over segments of (internal gaps + trailing gap) × perGap; must equal slack. */
function totalStretch(r: DistributeResult): number {
  let gaps = 0;
  for (const s of r.perSeg.values()) {
    gaps += s.splitBefore.length + (s.trailingGap ? 1 : 0);
  }
  return gaps * r.perGap;
}

describe('text-distribute adapter — positional signature', () => {
  it('forwards (segs, slack, firstContentSi, lastDrawnSi) positionally', () => {
    // [indent][観察する][結果]: firstContentSi=1 skips the indent, lastDrawnSi=2 is
    // the final segment (opens no gap). Verifies the positional args reach the
    // kernel in the right slots.
    const segs = [seg('　'), seg('観察する'), seg('結果')];
    const r = distributeLineSlack(segs, 30, 1, 2);
    expect(r).not.toBeNull();
    expect(r!.perSeg.has(0)).toBe(false); // indent fixed
    expect(r!.perSeg.has(2)).toBe(false); // final segment opens no gap
    const s = r!.perSeg.get(1)!;
    expect(s.splitBefore).toEqual([1, 2, 3]);
    expect(s.trailingGap).toBe(true);
    expect(r!.perGap).toBeCloseTo(7.5, 6);
    expect(totalStretch(r!)).toBeCloseTo(30, 6);
  });

  it('§17.18.44: fills a wrapped pure-CJK line via inter-CJK pitch (expansion default)', () => {
    // A pure-CJK phrase is ONE segment with no ASCII spaces; the default
    // (includeCJK omitted = true) opens every interior boundary so the line fills.
    const segs = [seg('観察することで')]; // 7 cp → 6 interior gaps
    const r = distributeLineSlack(segs, 60, 0, 1);
    expect(r).not.toBeNull();
    expect(r!.perSeg.get(0)!.splitBefore).toEqual([1, 2, 3, 4, 5, 6]);
    expect(r!.perGap).toBeCloseTo(10, 6);
  });
});

describe('text-distribute adapter — includeCJK compression mapping', () => {
  it('includeCJK=false compresses spaces only, never overlaps ideographs', () => {
    // The docx compression path (slack < 0 → includeCJK = false). Only the inter-
    // word space participates; the inter-CJK boundary is left alone.
    const segs = [seg('図 '), seg('観察'), seg('end')];
    const r = distributeLineSlack(segs, -10, 0, 2, -Infinity, false);
    expect(r).not.toBeNull();
    expect(r!.perSeg.get(0)!.trailingGap).toBe(true); // the space after "図 "
    expect(r!.perSeg.has(1)).toBe(false); // 観|察 NOT opened
    expect(r!.perGap).toBeCloseTo(-10, 6);
  });

  it('includeCJK=false returns null for a space-free CJK line (nothing to compress)', () => {
    expect(distributeLineSlack([seg('観察することで')], -30, 0, 1, -Infinity, false)).toBeNull();
  });

  it('clamps a compressing perGap to minPerGap', () => {
    const segs = [seg('観察することで')]; // 6 interior gaps, expansion
    const r = distributeLineSlack(segs, -60, 0, 1, -4);
    expect(r!.perGap).toBe(-4); // -60/6 = -10, clamped up to the -4 floor
  });
});

describe('text-distribute adapter — bidi exact-match exclusion (#483)', () => {
  it('justifies a pure-RTL line where the visually-last segment is the logical first', () => {
    // lastDrawnSi = logical-first (0). The exclusion is an EXACT match, not `>=`,
    // so segs 1 and 2 still distribute; `>=` would have skipped the whole line.
    const segs = [seg('a '), seg('b '), seg('c')];
    const r = distributeLineSlack(segs, 30, 0, 0);
    expect(r).not.toBeNull();
    expect(r!.perSeg.has(0)).toBe(false);
    expect(r!.perSeg.get(1)!.trailingGap).toBe(true);
    expect(totalStretch(r!)).toBeCloseTo(30, 6);
  });
});

describe('text-distribute adapter — noise floor', () => {
  it('returns null when |slack| is below 0.5px', () => {
    expect(distributeLineSlack([seg('観察')], 0.3, 0, 1)).toBeNull();
    expect(distributeLineSlack([seg('観察')], -0.4, 0, 1)).toBeNull();
  });
});
