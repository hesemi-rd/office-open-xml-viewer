import { describe, expect, it } from 'vitest';
import { layoutRichTextLines } from './renderer.js';
import type { CellFont, Run } from './types.js';

const BASE: CellFont = {
  bold: false, italic: false, underline: false, strike: false,
  size: 11, color: null, name: null,
};

function mockCtx(): CanvasRenderingContext2D {
  let font = '11px sans-serif';
  const px = () => Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '11');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (text: string) => ({ width: [...text].length * px() }) as TextMetrics,
    fillText() {},
    save() {},
    restore() {},
    fillStyle: '#000',
    strokeStyle: '#000',
    lineWidth: 1,
    textAlign: 'left',
    direction: 'ltr',
  } as unknown as CanvasRenderingContext2D;
}

describe('XLSX rich-text UAX #14 LB28 run-boundary glue', () => {
  it('moves a separate-run less-than sign with the following Arabic word', () => {
    // The draw font is 11pt = 14.67px. Seven cells fit: "wwww " + "<" fits,
    // while the following Arabic word overflows. The preceding space is legal.
    const lines = layoutRichTextLines(
      mockCtx(),
      [{ text: 'wwww ' }, { text: '<' }, { text: 'شيء' }] as Run[],
      BASE,
      1,
      7 * (11 * 96 / 72),
    );
    const texts = lines.map((line) => line.segments.map((segment) => segment.text).join(''));
    const bracketLine = texts.find((text) => text.includes('<'));

    expect(bracketLine).toBeDefined();
    expect(bracketLine).toContain('شيء');
  });

  it('moves a separate-run currency prefix with the following digits (LB25 PR × NU)', () => {
    // Seven cells fit: "wwww " + "$" fits, "100" overflows. LB25 (PR × NU)
    // forbids the $|1 seam, so "$100" wraps down as one unit.
    const lines = layoutRichTextLines(
      mockCtx(),
      [{ text: 'wwww ' }, { text: '$' }, { text: '100' }] as Run[],
      BASE,
      1,
      7 * (11 * 96 / 72),
    );
    const texts = lines.map((line) => line.segments.map((segment) => segment.text).join(''));
    const digitLine = texts.find((text) => text.includes('100'));

    expect(digitLine).toBeDefined();
    expect(digitLine).toContain('$100');
  });

  it('never orphans a separate-run opening bracket before digits (LB14 OP × NU)', () => {
    // Seven cells fit: "wwww " + "(" fits, "2026" overflows. LB14 (OP SP* ×)
    // forbids any break after "(", so "(2026" wraps down as one unit.
    const lines = layoutRichTextLines(
      mockCtx(),
      [{ text: 'wwww ' }, { text: '(' }, { text: '2026' }] as Run[],
      BASE,
      1,
      7 * (11 * 96 / 72),
    );
    const texts = lines.map((line) => line.segments.map((segment) => segment.text).join(''));
    const digitLine = texts.find((text) => text.includes('2026'));

    expect(digitLine).toBeDefined();
    expect(digitLine).toContain('(2026');
  });

  it('does NOT retract a preceding Thai (SEA) segment across an AL run boundary', () => {
    // "ภาษา" is Thai (Line_Break SA, resolved to AL by LB1). SEA dictionary
    // tailoring exposes a legal break at the ภาษา|< seam, so LB28 must NOT glue
    // it: "wwww ภาษา" stays on line 1 and "<a" wraps to line 2. Before the SEA
    // preceding-side guard, LB28 dragged "ภาษา" down onto the "<a" line.
    const lines = layoutRichTextLines(
      mockCtx(),
      [{ text: 'wwww ' }, { text: 'ภาษา' }, { text: '<a' }] as Run[],
      BASE,
      1,
      10 * (11 * 96 / 72),
    );
    const texts = lines.map((line) => line.segments.map((segment) => segment.text).join(''));
    const bracketLine = texts.find((text) => text.includes('<'));

    expect(bracketLine).toBeDefined();
    // The Thai word must not have been pulled onto the "<a" line.
    expect(bracketLine).not.toContain('ภาษา');
    // It stays on the first line with "wwww".
    expect(texts.find((text) => text.includes('ภาษา'))).toContain('wwww');
  });
});
