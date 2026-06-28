import { describe, it, expect } from 'vitest';
import { layoutParagraph } from './renderer.js';
import type { Paragraph } from './types';
import type { TextRunData } from '@silurus/ooxml-core';

// UAX#14 LB13: a closing / mid-punctuation char (comma, period, ;:!?)]} — class
// IS/CL/CP/EX) has NO break opportunity before it, so it may never BEGIN a line.
// When a word and a trailing comma live in SEPARATE runs (no whitespace between
// them), pptx's per-run tokeniser yields "system" + "," as adjacent tokens; the
// greedy wrap must keep them together (move "system," down as a unit), never
// orphaning "," at the next line's head. Mirrors the docx fix.

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

function lineTexts(line: { segments: { text: string }[] }): string {
  return line.segments.map((s) => s.text).join('');
}

describe('pptx Latin line-start-forbidden wrap (UAX#14 LB13)', () => {
  it('never orphans a comma at the start of a line when it lives in a separate run', () => {
    // char = 10px (mock). maxWidth=165 puts the break in the band where "system"
    // alone fits at line 1's end but "system," (comma glued) does not.
    const lines = layoutParagraph(
      mockCtx(), para([run('aaaa bbbb system'), run(', cc')]), 165, 20, '000000', 1, 0,
    );
    // No line may BEGIN with a comma.
    for (const ln of lines) {
      const first = ln.segments.find((s) => s.text.trim().length > 0);
      if (first) expect(first.text.startsWith(','), `line "${lineTexts(ln)}" starts with comma`).toBe(false);
    }
    // "system" must stay intact (not torn into "syste" + "m,").
    const allText = lines.map(lineTexts).join('|');
    expect(allText.includes('system'), `"system" intact in: ${allText}`).toBe(true);
  });

  it('never orphans a comma even when the word is split across runs by a font change', () => {
    // "system" is split into "sys" (Arial) + "tem" (Courier) by a formatting
    // change, then a comma in a third run. The comma must still never lead a
    // line; the word may split at the format seam ("…sys" / "tem, cc"), matching
    // docx/xlsx, but the non-starter rule (LB13) must still hold.
    const lines = layoutParagraph(
      mockCtx(),
      para([run('aaaa bbbb sys'), run('tem', { fontFamily: 'Courier' }), run(', cc')]),
      165, 20, '000000', 1, 0,
    );
    for (const ln of lines) {
      const first = ln.segments.find((s) => s.text.trim().length > 0);
      if (first) expect(first.text.startsWith(','), `line "${lineTexts(ln)}" starts with comma`).toBe(false);
    }
  });
});
