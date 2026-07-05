// ECMA-376 §17.6.12 `<w:pgNumType>` — resolve the DISPLAYED page number (and its
// number format) for every PHYSICAL page, honoring per-section restart (`w:start`)
// and re-formatting (`w:fmt`). This is the layer that separates the "physical page
// index" (0..N-1, produced by the paginator) from the number a PAGE field shows.
//
// The paginator stamps each element with `sectionPageNumType` — the page-numbering
// settings of the section that element belongs to (from the upcoming
// `SectionBreak.pageNumType`, or the body-level `section.pageNumType`). We read the
// FIRST element of each physical page to learn which section OWNS the top of that
// page and whether that section carries a restart.
//
// ── §17.6.12 semantics ──────────────────────────────────────────────────────
//   • `start` — "the page number that appears on the first page of the section."
//     So when a NEW section begins at the top of a physical page and that section
//     declares `w:start`, the running counter RESETS to `start` on that page.
//     "If this value is omitted, numbering continues from the highest page number
//     in the previous section" — i.e. keep incrementing.
//   • `fmt` — "the number format that shall be used for all page numbering in this
//     section." Absent ⇒ decimal (the default). Unlike `start`, `fmt` has no
//     continuation clause, so each section's format is independent (absent ⇒
//     decimal, NOT inherited from the previous section).
//
// ── §17.6.12 continuous restart semantics (issue #804, Word-confirmed) ────────
// A CONTINUOUS section break does not open a new physical page: the section begins
// mid-page, below the section it follows, and SHARES that page. §17.6.12 anchors a
// section's number series to the FIRST page its content appears on — and a
// continuous section's content first appears on the SHARED page, even though that
// page's TOP is still owned (and displayed) by the preceding section. So the series
// value equals `start` on the shared page and increments by 1 per physical page as
// the section spills forward. The DISPLAYED number of a page is the series value of
// whichever section OWNS that page's top.
//
// GROUND TRUTH — sample-27 (public/private): section 1 (p.1 upper) → continuous
// break + `w:pgNumType w:start="50"` → section 2 (shares p.1, spills to p.2/p.3),
// footer PAGE. Rendered in real Word the footers are [Page 1, Page 51, Page 52]:
//   • p.1 top is owned by section 1 ⇒ displays 1 (section 2's series is 50 here,
//     but section 2 does not own the top, so 50 is never shown);
//   • p.2 top is owned by section 2 ⇒ displays 51 (= 50 + (p.2 − p.1));
//   • p.3 top is owned by section 2 ⇒ displays 52.
// A pre-#804 implementation restarted the counter at the SPILLOVER page (the first
// page the section OWNS a top) and produced [1, 50, 51] — one short, because it did
// not count the shared page as the section's first page. `computePageNumbering` now
// anchors each restart to the section's first-APPEARANCE page (firstAppearanceBySection),
// so the spillover page shows `start + (thisPage − firstAppearancePage)`.
//
// This is consistent with real sample-13 ([1, 2, 3, 4, 5]): probed on the real file
// (browser pagination), its `w:start="2"` continuous section begins exactly AT a
// page boundary — its content first appears on physical page 2, the SAME page whose
// top it owns (it does not share page 1). The anchor offset is therefore 0 and its
// restart shows plain `start` = 2, which coincides with the natural continuation
// (1+1), so numbering stays sequential. (The related shape — a continuous restart
// that OWNS no page top at all, a mid-page island — never surfaces its start; no
// real sample has that shape, so it is pinned deterministically in
// page-number-field-render.test.ts.)

import type { PaginatedBodyElement, PageNumType } from './types';
import type { NumberFormat } from '@silurus/ooxml-core';

/** The displayed page number + its format for one physical page. */
export interface PageNumber {
  /** The number a PAGE field shows on this page (after §17.6.12 restart). */
  displayNumber: number;
  /** The ST_NumberFormat the section governing this page's TOP declares (§17.18.59);
   *  `decimal` when the section omits `w:fmt`. A PAGE field may still override this
   *  with its own `\*` switch (that is applied at field-resolution time, not here). */
  format: NumberFormat;
}

/** The page-numbering settings governing the TOP of physical page `pageIndex`,
 *  read from the first stamped element (the section that owns the page start).
 *  `null` when the page is empty or the section carries no `<w:pgNumType>`. */
function pageTopSettings(page: PaginatedBodyElement[] | undefined): PageNumType | null {
  return page?.[0]?.sectionPageNumType ?? null;
}

/** Identity of the section owning a page's top, used to detect a NEW section at a
 *  page boundary. We reuse the same `sectionHF` object identity the paginator
 *  stamps (a fresh object per section in the pagination pass), mirroring
 *  `resolvePageSection`'s `isFirstPageOfSection` test. The final (body-level)
 *  section is stamped as `undefined` (no upcoming `SectionBreak` marker) — a single
 *  well-defined identity, since a document has exactly one final section. */
function pageTopSectionId(page: PaginatedBodyElement[] | undefined): unknown {
  return page?.[0]?.sectionHF;
}

/**
 * Index the FIRST physical page each section's content APPEARS on — including a page
 * it merely SHARES (a continuous section starting mid-page, below the section it
 * follows), which it does not OWN. §17.6.12 anchors a section's number series to the
 * first page its content appears on, so a continuous restart must count that shared
 * page as the section's first page (issue #804 — see the module header).
 *
 * Keyed by the `sectionHF` identity the paginator stamps (fresh per section;
 * `undefined` for the final section). A section's pages are contiguous in flow order,
 * so the first appearance is a lower bound for every page it later owns.
 */
function firstAppearanceBySection(pages: PaginatedBodyElement[][]): Map<unknown, number> {
  const first = new Map<unknown, number>();
  for (let p = 0; p < pages.length; p++) {
    for (const el of pages[p]) {
      const id = el.sectionHF;
      if (!first.has(id)) first.set(id, p);
    }
  }
  return first;
}

/**
 * Compute the displayed page number + format for every physical page.
 *
 * @param pages the paginated body (one array of stamped elements per physical page)
 * @returns a per-physical-page array of {@link PageNumber}
 */
export function computePageNumbering(pages: PaginatedBodyElement[][]): PageNumber[] {
  const out: PageNumber[] = [];
  // §17.6.12 (#804) — a continuous section's `w:start` counts from the FIRST page its
  // content appears on (the shared page it starts mid-way down), not from the page it
  // first OWNS. Pre-index each section's first-appearance page so a restart that
  // surfaces only after a spillover still anchors to the shared page.
  const firstAppearance = firstAppearanceBySection(pages);
  let counter = 0; // the previous page's display number (0 before the first page)
  for (let p = 0; p < pages.length; p++) {
    const settings = pageTopSettings(pages[p]);
    const fmt = (settings?.fmt ?? 'decimal') as NumberFormat;

    // A page starts a NEW section when its owning section differs from the
    // previous page's (or it is the very first page). Only a NEW section may
    // restart the counter — a section continuing across a page break keeps
    // incrementing.
    const startsNewSection = p === 0 || pageTopSectionId(pages[p]) !== pageTopSectionId(pages[p - 1]);

    if (startsNewSection && settings?.start != null) {
      // §17.6.12 `w:start` — the number shown on the first page of the section.
      // The series began at `start` on the section's FIRST-APPEARANCE page (which
      // may be a page it merely shares, above the page it now owns — #804). So this
      // OWNED page shows `start + (thisPage − firstAppearancePage)`. When the section
      // begins at a genuine page top (a next-page break, or a continuous break whose
      // shared page IS the page it owns) the offset is 0 and this is just `start`.
      const anchor = firstAppearance.get(pageTopSectionId(pages[p])) ?? p;
      counter = settings.start + (p - anchor);
    } else {
      // §17.6.12 — otherwise numbering continues from the previous page.
      counter = counter + 1;
    }
    out.push({ displayNumber: counter, format: fmt });
  }
  return out;
}
