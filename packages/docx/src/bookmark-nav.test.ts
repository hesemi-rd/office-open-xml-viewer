import { describe, it, expect } from 'vitest';
import type { PaginatedBodyElement } from './types';
import { buildBookmarkPageMap, resolveBookmarkPage } from './bookmark-nav';

/** Minimal paginated paragraph carrying the given bookmark names. Only the
 *  fields the map builder reads (`type`, `bookmarks`) matter; the rest are
 *  filled to satisfy the type without affecting resolution. */
function para(bookmarks: string[]): PaginatedBodyElement {
  return {
    type: 'paragraph',
    alignment: 'left',
    indentLeft: 0,
    indentRight: 0,
    indentFirst: 0,
    spaceBefore: 0,
    spaceAfter: 0,
    lineSpacing: null,
    numbering: null,
    tabStops: [],
    runs: [],
    bookmarks,
  } as unknown as PaginatedBodyElement;
}

/** A paragraph with no bookmarks (page filler). */
function plain(): PaginatedBodyElement {
  return para([]);
}

describe('buildBookmarkPageMap', () => {
  it('maps a bookmark to the 0-based index of the page carrying it', () => {
    const pages: PaginatedBodyElement[][] = [
      [plain(), para(['_Toc_intro'])], // page 0
      [para(['_Toc_methods']), plain()], // page 1
      [para(['_Toc_results'])], // page 2
    ];
    const map = buildBookmarkPageMap(pages);
    expect(map.get('_Toc_intro')).toBe(0);
    expect(map.get('_Toc_methods')).toBe(1);
    expect(map.get('_Toc_results')).toBe(2);
  });

  it('maps multiple bookmarks that share a paragraph to the same page', () => {
    const pages: PaginatedBodyElement[][] = [[plain()], [para(['a', 'b', 'c'])]];
    const map = buildBookmarkPageMap(pages);
    expect(map.get('a')).toBe(1);
    expect(map.get('b')).toBe(1);
    expect(map.get('c')).toBe(1);
  });

  it('resolves a name repeated across pages to the FIRST (earliest) page', () => {
    // A paragraph split across a page break carries its bookmark on both slices;
    // the destination is where the paragraph begins.
    const pages: PaginatedBodyElement[][] = [
      [para(['dup'])], // page 0 — earliest wins
      [para(['dup'])], // page 1
    ];
    expect(buildBookmarkPageMap(pages).get('dup')).toBe(0);
  });

  it('returns undefined for a name that appears in no paragraph', () => {
    const map = buildBookmarkPageMap([[para(['known'])]]);
    expect(map.get('missing')).toBeUndefined();
  });

  it('honors bookmarks nested in a table cell paragraph', () => {
    const table = {
      type: 'table',
      rows: [
        {
          cells: [{ content: [para(['cellmark'])] }],
        },
      ],
    } as unknown as PaginatedBodyElement;
    const pages: PaginatedBodyElement[][] = [[plain()], [table]];
    expect(buildBookmarkPageMap(pages).get('cellmark')).toBe(1);
  });

  it('ignores empty-string bookmark names', () => {
    const map = buildBookmarkPageMap([[para([''])]]);
    expect(map.size).toBe(0);
  });
});

describe('resolveBookmarkPage', () => {
  it('is a thin lookup over the built map', () => {
    const map = buildBookmarkPageMap([[plain()], [para(['x'])]]);
    expect(resolveBookmarkPage(map, 'x')).toBe(1);
    expect(resolveBookmarkPage(map, 'nope')).toBeUndefined();
  });
});
