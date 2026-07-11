import { describe, expect, it } from 'vitest';
import type { TextRunData } from '@silurus/ooxml-core';
import { layoutParagraph } from './renderer.js';
import type { Paragraph } from './types.js';

function mockCtx(): CanvasRenderingContext2D {
  let font = '';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (text: string) => ({ width: [...text].length * 10 }) as TextMetrics,
    fillRect() {},
    fillText() {},
    fillStyle: '',
    strokeStyle: '',
  } as unknown as CanvasRenderingContext2D;
}

function run(text: string): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: 20, color: '000000', fontFamily: 'Arial',
  };
}

function paragraph(runs: TextRunData[]): Paragraph {
  return {
    alignment: 'l', marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null,
    defBold: null, defItalic: null, defFontFamily: null, tabStops: [],
    eaLnBrk: true, runs,
  } as Paragraph;
}

function lineText(line: { segments: { text: string }[] }): string {
  return line.segments.map((segment) => segment.text).join('');
}

describe('PPTX UAX #14 LB28 run-boundary glue', () => {
  it('moves a separate-run less-than sign with the following Arabic word', () => {
    // 7 cells: "wwww " + "<" fits, but adding the following four-letter word
    // does not. The preceding space is the legal opportunity.
    const lines = layoutParagraph(
      mockCtx(),
      paragraph([run('wwww '), run('<'), run('شيء')]),
      70, 20, '000000', 1, 0,
    );
    const texts = lines.map(lineText);
    const bracketLine = texts.find((text) => text.includes('<'));

    expect(bracketLine).toBeDefined();
    expect(bracketLine).toContain('شيء');
  });

  it('does NOT retract a preceding Thai (SEA) word across an AL run boundary', () => {
    // "ภาษา" is Thai (Line_Break SA → AL by LB1). The SEA dictionary tailoring
    // exposes a legal break at the ภาษา|< seam, so LB28 must not glue it:
    // "wwww ภาษา" (90px) stays on line 1 and "<a" wraps to line 2. Before the SEA
    // preceding-side guard, LB28 dragged "ภาษา" down onto the "<a" line.
    const lines = layoutParagraph(
      mockCtx(),
      paragraph([run('wwww '), run('ภาษา'), run('<a')]),
      100, 20, '000000', 1, 0,
    );
    const texts = lines.map(lineText);
    const bracketLine = texts.find((text) => text.includes('<'));

    expect(bracketLine).toBeDefined();
    expect(bracketLine).not.toContain('ภาษา');
    expect(texts.find((text) => text.includes('ภาษา'))).toContain('wwww');
  });
});
