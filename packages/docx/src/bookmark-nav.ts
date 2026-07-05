import type { PaginatedBodyElement, BodyElement, CellElement } from './types';

/**
 * DOCX internal-navigation resolution (IX-nav): a `bookmarkName → pageIndex` map.
 *
 * A `<w:hyperlink w:anchor="X">` (ECMA-376 §17.16.23) is an internal link that
 * jumps to the `<w:bookmarkStart w:name="X">` destination (§17.13.6.2). The
 * parser records, on each `DocParagraph`, the names of the bookmarks that start
 * within it (`DocParagraph.bookmarks`). This module scans the *paginated* pages
 * and records, for every bookmark name, the 0-based index of the FIRST page that
 * holds the paragraph carrying it — so an internal-link click can render / scroll
 * to the destination page.
 *
 * "First page" matters because the paginator can split a long paragraph across
 * pages (`lineSlice`); a bookmark on such a paragraph resolves to the page where
 * the paragraph *begins*, which is where the anchored content the user expects to
 * see starts. Bookmarks nested in table cells are honored too (a bookmark on a
 * heading inside a table still resolves), by walking each cell's paragraphs.
 *
 * Pure over the already-paginated pages — no DOM, no re-layout — so it is cheap
 * to build once per parse and trivially unit-testable from synthetic pages.
 */

/** Collect the bookmark names carried by a single body element (a paragraph, or
 *  the paragraphs nested in a table's cells). Tables can hold bookmarks on their
 *  cell paragraphs, so recurse one level into rows/cells. */
function collectElementBookmarks(
  el: BodyElement | CellElement,
  out: (name: string) => void,
): void {
  if (el.type === 'paragraph') {
    for (const name of el.bookmarks ?? []) out(name);
    return;
  }
  if (el.type === 'table') {
    for (const row of el.rows) {
      for (const cell of row.cells) {
        for (const cellEl of cell.content) {
          // A cell's content is paragraphs / nested tables (CellElement), so
          // recurse — a bookmark inside a nested table still resolves.
          collectElementBookmarks(cellEl, out);
        }
      }
    }
  }
}

/**
 * Build a `bookmarkName → 0-based page index` map from the paginated pages. For
 * each bookmark name the FIRST page carrying its paragraph wins (a name repeated
 * across pages resolves to the earliest, matching a top-to-bottom document scan).
 *
 * @param pages the paginated body elements, one array per page (as produced by
 *              the docx paginator).
 */
export function buildBookmarkPageMap(
  pages: readonly PaginatedBodyElement[][],
): Map<string, number> {
  const map = new Map<string, number>();
  for (let pageIndex = 0; pageIndex < pages.length; pageIndex++) {
    for (const el of pages[pageIndex]) {
      collectElementBookmarks(el, (name) => {
        if (name !== '' && !map.has(name)) map.set(name, pageIndex);
      });
    }
  }
  return map;
}

/**
 * Resolve a bookmark name (a `<w:hyperlink w:anchor>` internal target) to its
 * 0-based destination page index, or `undefined` when no bookmark of that name
 * exists in the document. Thin lookup over {@link buildBookmarkPageMap}'s result;
 * kept as a named function so the viewer's intent (`resolve anchor → page`) reads
 * clearly at the call site.
 */
export function resolveBookmarkPage(
  map: Map<string, number>,
  bookmarkName: string,
): number | undefined {
  return map.get(bookmarkName);
}
