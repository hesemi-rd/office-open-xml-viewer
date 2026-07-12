import { describe, expect, it } from 'vitest';

import { shapeParagraphGapBefore } from './renderer';
import type { ShapeText } from './types';

/**
 * ECMA-376 §17.3.1.9 `<w:contextualSpacing>` inside a text box (`<wps:txbx>`).
 *
 * The vertical gap reserved ABOVE text-box paragraph `i` is normally
 * max(prev.spaceAfter, this.spaceBefore) (collapse, not sum). When two ADJACENT
 * paragraphs share a style id and BOTH set contextualSpacing, Word drops the
 * whole gap — the same rule the body renderer applies via `contextualSuppressed`.
 * Before the fix the text-box path lacked this, so a `<w:contextualSpacing/>`
 * ListParagraph list inherited the docDefault `after=160` (8 pt) gap that
 * inflated its line pitch and clipped the trailing line (sample-32).
 */
describe('shapeParagraphGapBefore — §17.3.1.9 contextualSpacing in a text box', () => {
  const block = (over: Partial<ShapeText>): ShapeText =>
    ({ text: 'x', fontSizePt: 12, alignment: 'left', ...over }) as ShapeText;

  // 8 pt after / 4 pt before, both scaled ×1 for the table.
  const spBefore = [0, 4];
  const spAfter = [8, 0];

  it('reserves only the first block own spaceBefore at i=0 (no previous block)', () => {
    const blocks = [block({ spaceBefore: 6 }), block({})];
    expect(shapeParagraphGapBefore(blocks, 0, [6, 4], [0, 0])).toBe(6);
  });

  it('drops the gap when both blocks share a style id AND both set contextualSpacing', () => {
    const blocks = [
      block({ styleId: 'ListParagraph', contextualSpacing: true }),
      block({ styleId: 'ListParagraph', contextualSpacing: true }),
    ];
    expect(shapeParagraphGapBefore(blocks, 1, spBefore, spAfter)).toBe(0);
  });

  it('keeps the collapsed gap when the two blocks have DIFFERENT style ids', () => {
    const blocks = [
      block({ styleId: 'ListParagraph', contextualSpacing: true }),
      block({ styleId: 'Body', contextualSpacing: true }),
    ];
    // max(spBefore[1]=4, spAfter[0]=8) = 8.
    expect(shapeParagraphGapBefore(blocks, 1, spBefore, spAfter)).toBe(8);
  });

  it('keeps the gap when only ONE of the two same-style blocks sets contextualSpacing', () => {
    const blocks = [
      block({ styleId: 'ListParagraph', contextualSpacing: true }),
      block({ styleId: 'ListParagraph', contextualSpacing: false }),
    ];
    expect(shapeParagraphGapBefore(blocks, 1, spBefore, spAfter)).toBe(8);
  });

  it('keeps the gap when the shared style id is missing on either block', () => {
    const blocks = [
      block({ contextualSpacing: true }),
      block({ contextualSpacing: true }),
    ];
    expect(shapeParagraphGapBefore(blocks, 1, spBefore, spAfter)).toBe(8);
  });
});
