import { describe, it, expect } from 'vitest';
import { computePageNumbering } from './page-numbering';
import type { PaginatedBodyElement, PageNumType } from './types';

// Build one physical page carrying a section identity + its pgNumType. The
// paginator stamps `sectionHF` (a SINGLE object reference SHARED across all pages
// of one section — object identity is how a section boundary is detected) and
// `sectionPageNumType` on each element; the numbering layer reads only those, so a
// single stub element per page is sufficient. `sectionId` must therefore be the
// SAME object for every page of the same section (mirroring `currentSectionHF`).
function page(sectionId: object, pgNum: PageNumType | null): PaginatedBodyElement[] {
  const el = {
    type: 'paragraph',
    // Use `sectionId` itself as the identity object the numbering layer compares.
    sectionHF: sectionId as unknown,
    sectionPageNumType: pgNum,
  } as unknown as PaginatedBodyElement;
  return [el];
}

describe('computePageNumbering — ECMA-376 §17.6.12', () => {
  it('single section without pgNumType numbers 1..N in decimal (unchanged behaviour)', () => {
    const s = {};
    const pages = [page(s, null), page(s, null), page(s, null)];
    expect(computePageNumbering(pages)).toEqual([
      { displayNumber: 1, format: 'decimal' },
      { displayNumber: 2, format: 'decimal' },
      { displayNumber: 3, format: 'decimal' },
    ]);
  });

  it('start=1 on the first section is the identity (matches physical numbers)', () => {
    const s = {};
    const pages = [page(s, { start: 1 }), page(s, null), page(s, null)];
    expect(computePageNumbering(pages).map((n) => n.displayNumber)).toEqual([1, 2, 3]);
  });

  it('start=0 offsets the whole document (physical page 1 shows 0)', () => {
    const s = {};
    const pages = [page(s, { start: 0 }), page(s, null), page(s, null)];
    expect(computePageNumbering(pages).map((n) => n.displayNumber)).toEqual([0, 1, 2]);
  });

  it('start=25 offsets numbering to begin at 25', () => {
    const s = {};
    const pages = [page(s, { start: 25 }), page(s, null)];
    expect(computePageNumbering(pages).map((n) => n.displayNumber)).toEqual([25, 26]);
  });

  it('restarts the counter when a NEW section (page break) declares w:start', () => {
    // Front matter (2 pages, lowerRoman, start=1) then body (restart decimal from 1).
    // Every page of a section carries that section's pgNumType (the paginator stamps
    // `currentSectionPageNumType` on EVERY element, not only the section's first).
    const front = {};
    const body = {};
    const frontNum: PageNumType = { start: 1, fmt: 'lowerRoman' };
    const bodyNum: PageNumType = { start: 1, fmt: 'decimal' };
    const pages = [
      page(front, frontNum),
      page(front, frontNum),
      page(body, bodyNum),
      page(body, bodyNum),
    ];
    expect(computePageNumbering(pages)).toEqual([
      { displayNumber: 1, format: 'lowerRoman' },
      { displayNumber: 2, format: 'lowerRoman' },
      { displayNumber: 1, format: 'decimal' },
      { displayNumber: 2, format: 'decimal' },
    ]);
  });

  it('a new section WITHOUT w:start continues numbering from the previous section', () => {
    const s1 = {};
    const s2 = {};
    const pages = [page(s1, { start: 5 }), page(s1, null), page(s2, null), page(s2, null)];
    // 5, 6 in s1; s2 has no start so it continues 7, 8.
    expect(computePageNumbering(pages).map((n) => n.displayNumber)).toEqual([5, 6, 7, 8]);
  });

  it('applies each section format independently; absent fmt is decimal (not inherited)', () => {
    const s1 = {};
    const s2 = {};
    const s3 = {};
    const pages = [
      page(s1, { fmt: 'lowerRoman' }),
      page(s2, { start: 1 }), // no fmt ⇒ decimal, not roman
      page(s3, { fmt: 'upperLetter', start: 1 }),
    ];
    expect(computePageNumbering(pages)).toEqual([
      { displayNumber: 1, format: 'lowerRoman' },
      { displayNumber: 1, format: 'decimal' },
      { displayNumber: 1, format: 'upperLetter' },
    ]);
  });

  it('does NOT restart mid-page: a continuous section sharing a page keeps its number', () => {
    // A continuous section starts mid-page, so the SHARED page's TOP still belongs to
    // the preceding section — the displayed number is the preceding section's, never
    // the continuous section's `w:start`. (If the continuous section then SPILLS onto
    // a later page it OWNS, its start surfaces there, anchored to this shared page as
    // its first-appearance page — see the sample-27 [1, 51, 52] case in the module
    // header of page-numbering.ts and page-number-field-render.test.ts. This stub's
    // pages carry a single element each, so it models only the shared-page rule.)
    const s1 = {};
    const pages = [
      page(s1, null), // page 1: top belongs to s1 (even if s2 continues below it)
      page(s1, null),
    ];
    expect(computePageNumbering(pages).map((n) => n.displayNumber)).toEqual([1, 2]);
  });

  it('handles an empty page (no stamped element) as a decimal continuation', () => {
    const s = {};
    const pages = [page(s, { start: 3 }), [] as PaginatedBodyElement[], page(s, null)];
    // page 2 empty ⇒ continues 4; page 3 continues 5.
    expect(computePageNumbering(pages).map((n) => n.displayNumber)).toEqual([3, 4, 5]);
  });

  // Non-regression: a SIMPLIFIED sample-13-like shape — one section with start=1
  // owning every page top. Word's PDF for sample-13 shows sequential 1,2,3,4,5 and
  // this stub asserts the identity case (start=1 on physical page 1 ⇒ 1..N).
  //
  // NOTE — this stub is NOT how real sample-13 paginates. Probed on the real file
  // (browser pagination), its `start=2` continuous section begins exactly AT a page
  // boundary: its content first appears on physical page 2, the SAME page whose top
  // it owns (it does not share page 1). Its restart therefore fires with anchor
  // offset 0 and shows plain start=2 — which coincides with the natural continuation
  // (1+1) — so numbering stays sequential. The real-file behaviour is covered
  // end-to-end by tests/visual/page-number.spec.ts (renders sample-13 and asserts
  // 1..5); the SPILLOVER case where a continuous restart's first-appearance page
  // PRECEDES the first page it owns is pinned by sample-27 there and by the
  // deterministic continuous-spillover cases in page-number-field-render.test.ts
  // (#804). This stub pins only the start=1-identity arithmetic.
  it('reproduces sample-13: nextPage start=1 + continuous start=2 stays sequential', () => {
    const firstSection = {};
    const num: PageNumType = { start: 1 };
    const pages = [
      page(firstSection, num),
      page(firstSection, num),
      page(firstSection, num),
      page(firstSection, num),
      page(firstSection, num),
    ];
    // start=1 fires on physical page 1 (the identity) ⇒ 1..5.
    expect(computePageNumbering(pages).map((n) => n.displayNumber)).toEqual([1, 2, 3, 4, 5]);
  });
});
