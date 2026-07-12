import { describe, expect, it } from 'vitest';

import { shapeParagraphGapBefore } from './renderer';
import type { ShapeText } from './types';

/**
 * ECMA-376 §17.3.1.9 `<w:contextualSpacing>` — Word-adjudicated PER-SIDE
 * semantics (issue #1015, fixture sample-57 ground truth).
 *
 * The inter-paragraph gap decomposes as
 *   gap = prevContrib + currContrib
 *   prevContrib = prev.spaceAfter
 *   currContrib = max(curr.spaceBefore − prev.spaceAfter, 0)
 * (summing to max(after, before) — the §17.3.1.33 collapse). A paragraph whose
 * toggle matches a same-style neighbour drops ITS OWN contribution only:
 *   - prev toggles → gap = max(before − after, 0)   (spec's worked example: 10/12 → 2)
 *   - curr toggles → gap = after                    (Word measured 10, NOT the
 *     spec-literal "subtract own before from net" which would give 0)
 *   - both toggle  → gap = 0
 * Word applies this identically in body, table cell, and text box (sample-57
 * measured parity), so this text-box helper carries the same table as the body
 * and cell paths.
 */
describe('shapeParagraphGapBefore — §17.3.1.9 contextualSpacing per-side semantics', () => {
  const block = (over: Partial<ShapeText>): ShapeText =>
    ({ text: 'x', fontSizePt: 12, alignment: 'left', ...over }) as ShapeText;

  const pair = (
    prev: Partial<ShapeText>,
    curr: Partial<ShapeText>,
    afterPt: number,
    beforePt: number,
  ): number =>
    shapeParagraphGapBefore(
      [block(prev), block(curr)],
      1,
      [0, beforePt],
      [afterPt, 0],
    );

  it('reserves only the first block own spaceBefore at i=0 (no previous block)', () => {
    const blocks = [block({ spaceBefore: 6 }), block({})];
    expect(shapeParagraphGapBefore(blocks, 0, [6, 4], [0, 0])).toBe(6);
  });

  // ── sample-57 adjudication table (after=10 / before=12 unless noted) ──────
  it('case 1 — PREV-only toggle: gap = max(before − after, 0) = 2 (spec worked example)', () => {
    expect(
      pair(
        { styleId: 'CtxPair', contextualSpacing: true },
        { styleId: 'CtxPair' },
        10,
        12,
      ),
    ).toBe(2);
  });

  it('case 2 — CURR-only toggle: gap = prev.spaceAfter = 10 (Word; not the literal net-minus-before 0)', () => {
    expect(
      pair(
        { styleId: 'CtxPair' },
        { styleId: 'CtxPair', contextualSpacing: true },
        10,
        12,
      ),
    ).toBe(10);
  });

  it('case 3a — both toggle: gap = 0', () => {
    expect(
      pair(
        { styleId: 'CtxPair', contextualSpacing: true },
        { styleId: 'CtxPair', contextualSpacing: true },
        10,
        12,
      ),
    ).toBe(0);
  });

  it('case 3b — both toggle, asymmetric 4/12: still 0 (no negative residue)', () => {
    expect(
      pair(
        { styleId: 'CtxPair', contextualSpacing: true },
        { styleId: 'CtxPair', contextualSpacing: true },
        4,
        12,
      ),
    ).toBe(0);
  });

  it('case 4 — both toggle but DIFFERENT styles: no suppression, gap = max = 12', () => {
    expect(
      pair(
        { styleId: 'CtxPair', contextualSpacing: true },
        { styleId: 'CtxOther', contextualSpacing: true },
        10,
        12,
      ),
    ).toBe(12);
  });

  it('case 5 — no toggle: gap collapses to max(after, before) = 12', () => {
    expect(pair({ styleId: 'CtxPair' }, { styleId: 'CtxPair' }, 10, 12)).toBe(12);
  });

  // ── direction corners beyond the fixture values ───────────────────────────
  it('PREV-only toggle with after ≥ before floors at 0 (8/4 → max(4−8,0))', () => {
    expect(
      pair(
        { styleId: 'ListParagraph', contextualSpacing: true },
        { styleId: 'ListParagraph' },
        8,
        4,
      ),
    ).toBe(0);
  });

  it('CURR-only toggle with after ≥ before keeps prev.spaceAfter (8/4 → 8)', () => {
    expect(
      pair(
        { styleId: 'ListParagraph' },
        { styleId: 'ListParagraph', contextualSpacing: true },
        8,
        4,
      ),
    ).toBe(8);
  });

  it('keeps the collapsed gap when the shared style id is missing on either block', () => {
    expect(
      pair({ contextualSpacing: true }, { contextualSpacing: true }, 8, 4),
    ).toBe(8);
  });
});
