import { describe, it, expect } from 'vitest';
import { distributeLineSlack } from './line-distribute.js';
import type { DistributeSeg, DistributeResult } from './line-distribute.js';

// distributeLineSlack is the shared "fill the line" kernel behind WordprocessingML
// §17.18.44 (`both`/`distribute`) and DrawingML §20.1.10.59 (`just`/`dist`). Given
// a line's segments (logical order) and its slack, it returns per-segment split
// points + a trailing-gap flag and the per-gap px. Both gap families participate:
// inter-word whitespace AND inter-CJK boundaries (either side a CJK / ideographic
// glyph). The Word PDF for private/sample-9 pins the CJK case: a wrapped pure-CJK
// `both` line with zero ASCII spaces still fills to the right margin (xMax 523.0pt
// via pdftotext -bbox), only possible by widening inter-CJK boundaries.
//
// These tests exercise the kernel directly with its default (WordprocessingML)
// predicates AND with the injected predicates the pptx adapter supplies.

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

describe('distributeLineSlack — pure CJK (the bug)', () => {
  // A pure-CJK phrase is ONE segment with no ASCII spaces. A space-only model
  // produces zero stretch; the kernel must open every inter-CJK boundary.
  it('opens an inter-CJK gap at every interior boundary of a single CJK segment', () => {
    const segs = [seg('観察することで')]; // 7 code points → 6 interior boundaries
    const r = distributeLineSlack(segs, 60, { firstContentSi: 0, lastDrawnSi: 1 });
    expect(r).not.toBeNull();
    const s = r!.perSeg.get(0)!;
    expect(s.splitBefore).toEqual([1, 2, 3, 4, 5, 6]);
    expect(s.trailingGap).toBe(false);
    expect(r!.perGap).toBeCloseTo(60 / 6, 6);
    expect(totalStretch(r!)).toBeCloseTo(60, 6);
  });

  it('returns null for a single CJK glyph (no interior boundary)', () => {
    expect(distributeLineSlack([seg('観')], 40, { firstContentSi: 0, lastDrawnSi: 1 })).toBeNull();
  });
});

describe('distributeLineSlack — default options', () => {
  // firstContentSi defaults to 0, lastDrawnSi to segments.length-1, isGapChar to
  // core.isCjkBreakChar, isWhitespace to U+0020||U+3000. Omitting opts entirely
  // must behave like the explicit WordprocessingML-expansion form.
  it('omitting opts uses 0 / length-1 / CJK / ASCII+ideographic-space defaults', () => {
    const segs = [seg('観察'), seg('結果')];
    const withDefaults = distributeLineSlack(segs, 18);
    const explicit = distributeLineSlack(segs, 18, {
      firstContentSi: 0,
      lastDrawnSi: 1,
    });
    expect(withDefaults).not.toBeNull();
    expect(explicit).not.toBeNull();
    expect(withDefaults!.perGap).toBeCloseTo(explicit!.perGap, 9);
    expect(withDefaults!.perSeg.get(0)).toEqual(explicit!.perSeg.get(0));
    expect(withDefaults!.perSeg.has(1)).toBe(false);
  });
});

describe('distributeLineSlack — Latin justify is unchanged (regression guard)', () => {
  it('stretches only inter-word ASCII spaces, never inside a Latin word', () => {
    const segs = [seg('the '), seg('quick')];
    const r = distributeLineSlack(segs, 20, { firstContentSi: 0, lastDrawnSi: 1 });
    expect(r).not.toBeNull();
    const s = r!.perSeg.get(0)!;
    expect(s.splitBefore).toEqual([]); // no intra-word split
    expect(s.trailingGap).toBe(true); // the inter-word space after "the "
    expect(r!.perSeg.has(1)).toBe(false); // final segment opens no gap
    expect(r!.perGap).toBeCloseTo(20, 6);
    expect(totalStretch(r!)).toBeCloseTo(20, 6);
  });

  it('distributes across multiple inter-word spaces equally', () => {
    const segs = [seg('the '), seg('lazy '), seg('dog')];
    const r = distributeLineSlack(segs, 30, { firstContentSi: 0, lastDrawnSi: 2 });
    expect(r).not.toBeNull();
    expect(r!.perGap).toBeCloseTo(15, 6);
    expect(r!.perSeg.get(0)!.trailingGap).toBe(true);
    expect(r!.perSeg.get(1)!.trailingGap).toBe(true);
    expect(totalStretch(r!)).toBeCloseTo(30, 6);
  });

  it('never splits a long single Latin word (no gaps → null)', () => {
    expect(
      distributeLineSlack([seg('supercalifragilistic')], 50, { firstContentSi: 0, lastDrawnSi: 1 }),
    ).toBeNull();
  });
});

describe('distributeLineSlack — mixed Latin + CJK', () => {
  it('opens gaps at inter-CJK boundaries and at the Latin/CJK boundary', () => {
    // "BZ反応" — 4 code points: B Z 反 応. B|Z Latin → no gap; Z|反 (CJK right) and
    // 反|応 (both CJK) → gaps. A single content segment uses lastDrawnSi past the
    // end (1) so the segment is NOT the "visually-last" one and may open gaps.
    const segs = [seg('BZ反応')];
    const r = distributeLineSlack(segs, 24, { firstContentSi: 0, lastDrawnSi: 1 });
    expect(r).not.toBeNull();
    const s = r!.perSeg.get(0)!;
    expect(s.splitBefore).toEqual([2, 3]);
    expect(s.trailingGap).toBe(false);
    expect(r!.perGap).toBeCloseTo(12, 6);
  });

  it('does not open a gap between two Latin letters even next to CJK', () => {
    // "反BZ応": 反(0) B(1) Z(2) 応(3). 反|B and Z|応 are gaps; B|Z is not.
    const segs = [seg('反BZ応')];
    const r = distributeLineSlack(segs, 20, { firstContentSi: 0, lastDrawnSi: 1 });
    const s = r!.perSeg.get(0)!;
    expect(s.splitBefore).toEqual([1, 3]);
  });
});

describe('distributeLineSlack — leading 字下げ indent stays fixed', () => {
  it('does not stretch a leading whitespace segment before firstContentSi', () => {
    const segs = [seg('　'), seg('観察する'), seg('結果')];
    const r = distributeLineSlack(segs, 30, { firstContentSi: 1, lastDrawnSi: 2 });
    expect(r).not.toBeNull();
    expect(r!.perSeg.has(0)).toBe(false); // indent segment never stretched
    expect(r!.perSeg.has(2)).toBe(false); // final segment opens no gap
    const s = r!.perSeg.get(1)!;
    expect(s.splitBefore).toEqual([1, 2, 3]);
    expect(s.trailingGap).toBe(true);
    expect(r!.perGap).toBeCloseTo(7.5, 6); // 30 / 4
    expect(totalStretch(r!)).toBeCloseTo(30, 6);
  });

  it('skips a leading ideographic space at the START of the content segment', () => {
    const segs = [seg('　観察'), seg('結果')];
    const r = distributeLineSlack(segs, 16, { firstContentSi: 0, lastDrawnSi: 1 });
    expect(r).not.toBeNull();
    const s = r!.perSeg.get(0)!;
    expect(s.splitBefore).toEqual([2]);
    expect(s.trailingGap).toBe(true);
    expect(totalStretch(r!)).toBeCloseTo(16, 6);
  });

  it('treats an interior ideographic space as a single inter-word gap', () => {
    const segs = [seg('観　察結')];
    const r = distributeLineSlack(segs, 10, { firstContentSi: 0, lastDrawnSi: 1 });
    expect(r).not.toBeNull();
    const s = r!.perSeg.get(0)!;
    expect(s.splitBefore).toEqual([2, 3]);
    expect(r!.perGap).toBeCloseTo(5, 6); // 10 / 2
  });
});

describe('distributeLineSlack — final segment opens no gap, Σgaps == slack', () => {
  it('opens no gap WITHIN the final segment but does widen the boundary into it', () => {
    const segs = [seg('観察'), seg('結果')];
    const r = distributeLineSlack(segs, 18, { firstContentSi: 0, lastDrawnSi: 1 });
    expect(r).not.toBeNull();
    expect(r!.perSeg.has(1)).toBe(false);
    const s = r!.perSeg.get(0)!;
    expect(s.splitBefore).toEqual([1]); // 観|察 internal
    expect(s.trailingGap).toBe(true); // 察|結 boundary into the final segment
    expect(r!.perGap).toBeCloseTo(9, 6); // 18 / 2 gaps
    expect(totalStretch(r!)).toBeCloseTo(18, 6);
  });

  it('Σgaps equals slack for a mixed multi-segment line', () => {
    const segs = [seg('図 '), seg('観察する'), seg('こと')];
    const r = distributeLineSlack(segs, 40, { firstContentSi: 0, lastDrawnSi: 2 });
    expect(r).not.toBeNull();
    expect(totalStretch(r!)).toBeCloseTo(40, 6);
  });
});

describe('distributeLineSlack — negative slack (compression)', () => {
  it('compresses gaps with a negative perGap, clamped to minPerGap', () => {
    const segs = [seg('観察することで')]; // 6 interior gaps
    const r = distributeLineSlack(segs, -60, { firstContentSi: 0, lastDrawnSi: 1, minPerGap: -4 });
    expect(r).not.toBeNull();
    expect(r!.perGap).toBe(-4);
  });

  it('applies an unclamped negative perGap when within the cap', () => {
    const segs = [seg('観察することで')];
    const r = distributeLineSlack(segs, -12, { firstContentSi: 0, lastDrawnSi: 1, minPerGap: -10 });
    expect(r!.perGap).toBeCloseTo(-2, 6); // -12/6 = -2, above the -10 floor
  });

  it('returns null when |slack| is below the 0.5px noise floor', () => {
    expect(distributeLineSlack([seg('観察')], 0.3, { firstContentSi: 0, lastDrawnSi: 1 })).toBeNull();
    expect(distributeLineSlack([seg('観察')], -0.4, { firstContentSi: 0, lastDrawnSi: 1 })).toBeNull();
  });

  it('with isGapChar=()=>false, compresses only spaces and leaves CJK boundaries alone', () => {
    // The docx COMPRESSION path: slack < 0 → only spaces stretch, no glyph overlap.
    const segs = [seg('図 '), seg('観察'), seg('end')];
    const r = distributeLineSlack(segs, -10, {
      firstContentSi: 0,
      lastDrawnSi: 2,
      isGapChar: () => false,
    });
    expect(r).not.toBeNull();
    expect(r!.perSeg.get(0)!.trailingGap).toBe(true);
    expect(r!.perSeg.has(1)).toBe(false);
    expect(r!.perGap).toBeCloseTo(-10, 6); // one gap absorbs all the (negative) slack
  });

  it('with isGapChar=()=>false, returns null for a space-free CJK line', () => {
    expect(
      distributeLineSlack([seg('観察することで')], -30, {
        firstContentSi: 0,
        lastDrawnSi: 1,
        isGapChar: () => false,
      }),
    ).toBeNull();
  });
});

describe('distributeLineSlack — non-text atoms', () => {
  it('opens a CJK gap against an inline atom edge but not inside it', () => {
    const segs: DistributeSeg[] = [{}, seg('観察'), seg('end')];
    const r = distributeLineSlack(segs, 30, { firstContentSi: 0, lastDrawnSi: 2 });
    expect(r).not.toBeNull();
    expect(r!.perSeg.get(0)!.trailingGap).toBe(true);
    expect(r!.perSeg.get(0)!.splitBefore).toEqual([]);
    expect(r!.perSeg.get(1)!.splitBefore).toEqual([1]);
    expect(totalStretch(r!)).toBeCloseTo(30, 6);
  });
});

describe('distributeLineSlack — bidi (visually-last segment ≠ logical-last)', () => {
  // lastDrawnSi = the VISUALLY-last segment's LOGICAL index. The exclusion must be
  // an EXACT match, not `>=`: for a pure-RTL line lastDrawnSi is the logical-FIRST
  // segment (0), so `>=` would suppress every segment → total 0 → null. #483.
  it('justifies when the fixed (visually-last) segment is the logical first (RTL)', () => {
    const segs = [seg('a '), seg('b '), seg('c')];
    const r = distributeLineSlack(segs, 30, { firstContentSi: 0, lastDrawnSi: 0 });
    expect(r).not.toBeNull(); // was null under the `>=` regression
    expect(r!.perSeg.has(0)).toBe(false); // the visually-last segment opens no gap
    expect(r!.perSeg.get(1)!.trailingGap).toBe(true);
    expect(totalStretch(r!)).toBeCloseTo(30, 6);
  });
});

describe('distributeLineSlack — injected whitespace predicate (pptx parity)', () => {
  // PowerPoint treats every JS `\s` char as inter-word whitespace, wider than the
  // WordprocessingML default (U+0020 || U+3000). The kernel must classify by the
  // injected predicate, so e.g. a TAB becomes an inter-word gap under pptx but a
  // CJK-or-nothing boundary under the default.
  const pptxWs = (cp: number): boolean => /\s/.test(String.fromCodePoint(cp));

  it('a TAB is an inter-word gap under the pptx predicate', () => {
    // "a\tb" + "c" → a(off0) \t(off1) b(off2) | c. Under pptxWs the tab is the
    // only whitespace → exactly one gap, after it (interior: off1 → split before
    // off2). The b|c boundary is Latin/Latin and NOT whitespace, so it opens no
    // gap; the only stretch is the tab.
    const segs = [seg('a\tb'), seg('c')];
    const r = distributeLineSlack(segs, 12, {
      firstContentSi: 0,
      lastDrawnSi: 1,
      isWhitespace: pptxWs,
    });
    expect(r).not.toBeNull();
    const s = r!.perSeg.get(0)!;
    expect(s.splitBefore).toEqual([2]); // split before b (after the tab)
    expect(s.trailingGap).toBe(false); // b|c is Latin/Latin → not a gap
    expect(r!.perGap).toBeCloseTo(12, 6); // single gap absorbs all the slack
    expect(r!.perSeg.has(1)).toBe(false);
  });

  it('the default whitespace predicate does NOT treat a TAB as a gap', () => {
    // Same input under defaults: \t is not whitespace and not CJK, so a|\t, \t|b
    // and b|c are all Latin/control boundaries → no gap anywhere → null. (Word
    // never sees a bare \t as a justify opportunity; pptx does.)
    const segs = [seg('a\tb'), seg('c')];
    const r = distributeLineSlack(segs, 12, { firstContentSi: 0, lastDrawnSi: 1 });
    expect(r).toBeNull();
  });
});
