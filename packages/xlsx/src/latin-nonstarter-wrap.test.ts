import { describe, it, expect } from 'vitest';
import { layoutRichTextLines } from './renderer.js';
import type { CellFont, Run } from './types.js';

// UAX#14 LB13: a closing / mid-punctuation char (comma, period, ;:!?)]} — class
// IS/CL/CP/EX) has NO break opportunity before it, so it may never BEGIN a line.
// In a WRAPPED rich-text cell, a word and a trailing comma can live in SEPARATE
// runs (no whitespace between). The per-run tokeniser yields "system" + "," as
// adjacent tokens; the wrap must keep them together (move "system," down as a
// unit). The existing kinsoku retract is CHARACTER-level, so it would instead
// tear the word ("syste" + "m,") — also wrong. Mirrors the docx fix.

const BASE: CellFont = {
  bold: false, italic: false, underline: false, strike: false,
  size: 11, color: null, name: null,
};

function makeCtx(): CanvasRenderingContext2D {
  let font = '11px sans-serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '11');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    measureText: (s: string) => ({ width: [...s].length * px() }) as TextMetrics,
    fillText() {}, save() {}, restore() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as 'ltr' | 'rtl',
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

describe('xlsx Latin line-start-forbidden wrap (UAX#14 LB13)', () => {
  it('keeps "system," together — never orphans the comma nor tears the word', () => {
    // Mock metrics = 15px/char (round(11pt × 96/72)). maxWidth=247 lands in the
    // band [240,255) where "aaaa bbbb system" fits on line 1 but the comma
    // overflows: the char-level kinsoku retract used to tear the word into
    // "aaaa bbbb syste" + "m," (verified empirically). The fix moves the whole
    // word + comma down together → "aaaa bbbb " / "system, cc".
    const lines = layoutRichTextLines(
      makeCtx(), [{ text: 'aaaa bbbb system' }, { text: ', cc' }] as Run[], BASE, 1, 247,
    );
    const lineTexts = lines.map((l) => l.segments.map((s) => s.text).join(''));
    // No line may BEGIN with a comma.
    for (const lt of lineTexts) {
      expect(/^\s*,/.test(lt), `line "${lt}" starts with comma`).toBe(false);
    }
    // "system" must stay intact (not torn into "syste" + "m,").
    expect(lineTexts.some((lt) => lt.includes('system')), `"system" intact in: ${lineTexts.join('|')}`).toBe(true);
  });
});
