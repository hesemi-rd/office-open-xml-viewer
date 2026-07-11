import { describe, it, expect } from 'vitest';
import { seaWordBreakOffsets } from '@silurus/ooxml-core';
import { layoutParagraph } from './renderer.js';
import type { Paragraph } from './types';
import type { TextRunData } from '@silurus/ooxml-core';

// Issue #797 — dictionary line breaking for Thai/Lao/Khmer in pptx. A spaceless
// Thai run must wrap at segmenter word boundaries (never mid-word), and each
// line's SEA text must stay ONE segment (measure==paint — pptx sums per-push
// widths into the wrap accumulator). char = 10px in the mock, so widths are
// deterministic; word boundaries come from the platform ICU dictionary.

function mockCtx() {
  let font = '';
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    measureText: (s: string) => ({ width: s.length * 10 }),
    fillRect() {}, fillText() {},
    fillStyle: '', strokeStyle: '',
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

function run(text: string, over: Partial<TextRunData> = {}): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: 20, color: '000000', fontFamily: 'Arial', ...over,
  };
}

function para(runs: TextRunData[]): Paragraph {
  return {
    alignment: 'l', marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null, defBold: null,
    defItalic: null, defFontFamily: null, tabStops: [], eaLnBrk: true, runs,
  } as Paragraph;
}

const lineText = (line: { segments: { text: string }[] }): string =>
  line.segments.map((s) => s.text).join('');

function breakOffsets(texts: string[]): number[] {
  const out: number[] = [];
  let acc = 0;
  for (let i = 0; i < texts.length - 1; i++) { acc += texts[i].length; out.push(acc); }
  return out;
}

describe('pptx SEA (Thai) dictionary line breaking', () => {
  const thai = 'ภาษาไทยเป็นภาษาที่สวยงามมาก';
  const wordStarts = new Set(seaWordBreakOffsets(thai));

  it('breaks only at dictionary word boundaries and preserves the text', () => {
    const lines = layoutParagraph(mockCtx(), para([run(thai)]), 100, 20, '000000', 1, 0);
    const texts = lines.map(lineText);
    expect(texts.length).toBeGreaterThan(1);
    expect(texts.join('')).toBe(thai);
    for (const b of breakOffsets(texts)) expect(wordStarts.has(b)).toBe(true);
  });

  it('keeps each wrapped line SEA text as a single segment (measure==paint)', () => {
    const lines = layoutParagraph(mockCtx(), para([run(thai)]), 100, 20, '000000', 1, 0);
    for (const ln of lines) {
      const textSegs = ln.segments.filter((s) => s.text.length > 0);
      // At most one text segment per line (all same-font Thai merges / is one push).
      expect(textSegs.length).toBeLessThanOrEqual(1);
    }
  });

  it('makes progress and preserves text when the shape is narrower than one word', () => {
    const lines = layoutParagraph(mockCtx(), para([run(thai)]), 15, 20, '000000', 1, 0);
    const texts = lines.map(lineText);
    expect(texts.join('')).toBe(thai);
    expect(texts.every((t) => t.length > 0)).toBe(true);
  });

  it('lays a Thai run that fits on one line without fragmenting', () => {
    const lines = layoutParagraph(mockCtx(), para([run(thai)]), 400, 20, '000000', 1, 0);
    expect(lines).toHaveLength(1);
    const textSegs = lines[0].segments.filter((s) => s.text.length > 0);
    expect(textSegs).toHaveLength(1);
    expect(textSegs[0].text).toBe(thai);
  });
});
