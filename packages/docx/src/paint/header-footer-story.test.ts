import { describe, expect, it, vi } from 'vitest';
import { createHeaderFooterStoryPainter } from './header-footer-story.js';

interface TestElement {
  readonly type: 'paragraph' | 'table';
  readonly id: string;
  readonly frame?: boolean;
}

describe('header/footer story frame compatibility', () => {
  it('keeps frame paragraphs on the legacy story callback without ordinary paragraph paint', () => {
    const paintFrameParagraph = vi.fn();
    const paintParagraph = vi.fn();
    const painter = createHeaderFooterStoryPainter({
      preRegisterPageFloats: vi.fn(),
      paragraphOf: (element: TestElement) => element,
      tableOf: (element: TestElement) => element,
      hasFrame: (paragraph: TestElement) => paragraph.frame === true,
      frameAnchorLineHeight: () => 14,
      paintFrameParagraph,
      spaceBefore: () => 0,
      spaceAfter: () => 0,
      bordersOf: () => undefined,
      contextualSpacing: () => ({ suppressBefore: false, overlap: 0 }),
      hasBorder: () => false,
      sharesBorder: () => false,
      paintParagraph,
      paintTable: vi.fn(),
      tableResetsParagraphFlow: () => false,
    });
    const frame = { type: 'paragraph', id: 'frame', frame: true } as const;
    const base = { y: 0, scale: 1 };

    expect(painter({ body: [frame] }, 20, base)).toBe(20);
    expect(paintFrameParagraph).toHaveBeenCalledOnce();
    expect(paintFrameParagraph).toHaveBeenCalledWith(frame, expect.objectContaining({ y: 20 }), 14);
    expect(paintParagraph).not.toHaveBeenCalled();
  });
});
