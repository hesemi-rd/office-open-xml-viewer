import { describe, it, expect } from 'vitest';
import { tateChuYokoOverlayScale } from './tate-chu-yoko-overlay.js';
import type { DocxTextRunInfo } from './renderer';

// ECMA-376 §17.3.2.10 縦中横 overlay clamp (#836) — the shared factor used by both
// the find-highlight and the text-selection overlays.

function run(partial: Partial<DocxTextRunInfo>): DocxTextRunInfo {
  return { text: 'X', x: 0, y: 0, w: 10, h: 12, fontSize: 12, font: '12px serif', ...partial };
}

describe('tateChuYokoOverlayScale', () => {
  it('returns 1 for a run that is not eastAsianVert (no clamp)', () => {
    const r = run({ text: '２９', w: 7 });
    expect(tateChuYokoOverlayScale(r, (s) => [...s].length * 7)).toBe(1);
  });

  it('returns run.w / naturalWidth for a 縦中横 run (compresses to the cell)', () => {
    // "２９" natural 14px, cell 7px ⇒ 0.5.
    const r = run({ text: '２９', w: 7, eastAsianVert: true });
    expect(tateChuYokoOverlayScale(r, (s) => [...s].length * 7)).toBeCloseTo(0.5, 10);
  });

  it('folds in w:w: a further-narrowed cell yields a smaller factor', () => {
    // With w:w compression the reported cell can be < one em; the factor tracks
    // whatever the drawn cell is (natural 14px, cell 5px ⇒ 5/14).
    const r = run({ text: '２９', w: 5, eastAsianVert: true });
    expect(tateChuYokoOverlayScale(r, (s) => [...s].length * 7)).toBeCloseTo(5 / 14, 10);
  });

  it('never expands: a cell wider than the natural glyphs keeps scale 1', () => {
    const r = run({ text: '２', w: 20, eastAsianVert: true }); // natural 7, cell 20
    expect(tateChuYokoOverlayScale(r, (s) => [...s].length * 7)).toBe(1);
  });

  it('degrades to 1 when the measurer reports a degenerate (zero) width', () => {
    const r = run({ text: '２９', w: 7, eastAsianVert: true });
    expect(tateChuYokoOverlayScale(r, () => 0)).toBe(1);
  });
});
