import { describe, it, expect } from 'vitest';
import { sumCellContentHeight } from './renderer.js';
import type { CellElement, DocParagraph } from './types.js';

// `sumCellContentHeight` is the shared measure-side spacing-collapse helper
// that mirrors `renderCellContent`'s paint-side spacing rules:
//   * ECMA-376 §17.3.1.33 contextualSpacing (same styleId siblings suppress
//     spaceBefore between them).
//   * Adjacent-paragraph spacing OVERLAP: the gap between two paragraphs is
//     max(prevSpaceAfter, currSpaceBefore), not their sum.
// Both rules were applied in the paint loop (`renderCellContent`) but the
// measurement paths (`measureCellContentHeightPx`, `computeTableRowHeights`)
// previously summed `spaceBefore + spaceAfter` per element unconditionally, so a
// cell measured taller than it painted — leaving an unexplained gap below a
// nested table whose trailing empty paragraph carried `w:spaceBefore`.

function para(opts: Partial<DocParagraph> = {}): CellElement {
  return {
    type: 'paragraph',
    runs: [],
    spaceBefore: 0,
    spaceAfter: 0,
    contextualSpacing: false,
    styleId: null,
    ...opts,
  } as unknown as CellElement;
}

function tbl(): CellElement {
  return { type: 'table' } as unknown as CellElement;
}

// In each case the per-element height returns the paragraph's "intrinsic + full
// spacing" measurement — i.e. exactly what measureCellElementHeight and
// estimateParagraphHeight return without contextual or overlap collapse.
const fullHeight = (intrinsic: number) => (ce: CellElement): number => {
  if (ce.type === 'paragraph') {
    const p = ce as unknown as DocParagraph;
    return intrinsic + p.spaceBefore + p.spaceAfter;
  }
  // Nested table — caller supplies its own height (10 in these tests).
  return 10;
};

describe('sumCellContentHeight — spacing collapse mirrors renderCellContent', () => {
  it('a single paragraph contributes intrinsic + spaceBefore + spaceAfter', () => {
    const content = [para({ spaceBefore: 6, spaceAfter: 8 })];
    expect(sumCellContentHeight(content, fullHeight(20), 1)).toBe(20 + 6 + 8);
  });

  it('between two paragraphs the gap is max(prevAfter, currBefore), not the sum', () => {
    const content = [
      para({ spaceAfter: 12 }),
      para({ spaceBefore: 12 }),
    ];
    // Without collapse: 20 + 12 + 20 + 12 = 64.
    // With  collapse: gap = max(12,12) = 12, total = 20 + 12 + 20 = 52.
    expect(sumCellContentHeight(content, fullHeight(20), 1)).toBe(52);
  });

  it('asymmetric spacing collapses to the larger of the two', () => {
    const content = [
      para({ spaceAfter: 4 }),
      para({ spaceBefore: 10 }),
    ];
    // gap = max(4, 10) = 10, total = 20 + 10 + 20 = 50.
    expect(sumCellContentHeight(content, fullHeight(20), 1)).toBe(50);
  });

  it('contextualSpacing across siblings of the same style suppresses spaceBefore (§17.3.1.33)', () => {
    const content = [
      para({ styleId: 'ListBullet', contextualSpacing: true, spaceAfter: 0 }),
      para({ styleId: 'ListBullet', contextualSpacing: true, spaceBefore: 8 }),
    ];
    // 2nd paragraph's spaceBefore is suppressed: 20 + 0 + 20 + 0 = 40
    // (no spaceBefore for the suppressed paragraph, no overlap to subtract).
    expect(sumCellContentHeight(content, fullHeight(20), 1)).toBe(40);
  });

  it('contextualSpacing does NOT apply across different styles', () => {
    const content = [
      para({ styleId: 'ListBullet', contextualSpacing: true, spaceAfter: 4 }),
      para({ styleId: 'Body', contextualSpacing: true, spaceBefore: 10 }),
    ];
    // Different styleIds → no §17.3.1.33 suppression; falls back to overlap.
    // gap = max(4, 10) = 10, total = 20 + 10 + 20 = 50.
    expect(sumCellContentHeight(content, fullHeight(20), 1)).toBe(50);
  });

  it('a nested table resets the prev-paragraph context (no overlap across the table)', () => {
    const content = [
      para({ spaceAfter: 12 }),
      tbl(),                          // intrinsic 10, ends spacing context
      para({ spaceBefore: 12 }),
    ];
    // 1st para 20 + 12 (its spaceAfter is NOT collapsed against the table —
    // table reset is the paint-pass model)
    // + table 10
    // + 2nd para 20 + 12 (full spaceBefore, no prev paragraph to overlap with)
    // = 74.
    expect(sumCellContentHeight(content, fullHeight(20), 1)).toBe(74);
  });

  it('scales spacing in px units (spaceScale = device scale)', () => {
    const content = [
      para({ spaceAfter: 12 }),
      para({ spaceBefore: 12 }),
    ];
    // perElementHeight returns px directly (intrinsic 40 + spacing × 2);
    // sumCellContentHeight subtracts the overlap in matching px (12 × 2 = 24).
    const perEl = (ce: CellElement): number => {
      if (ce.type === 'paragraph') {
        const p = ce as unknown as DocParagraph;
        return 40 + (p.spaceBefore + p.spaceAfter) * 2;
      }
      return 20;
    };
    // Without collapse: 40 + 12*2 + 40 + 12*2 = 128.
    // With  collapse: 40 + 12*2 + 40 = 104.
    expect(sumCellContentHeight(content, perEl, 2)).toBe(104);
  });
});
