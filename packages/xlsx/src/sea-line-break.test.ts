import { describe, it, expect } from 'vitest';
import { seaWordBreakOffsets } from '@silurus/ooxml-core';
import { wrapParagraphLines, layoutRichTextLines } from './renderer.js';
import type { CellFont, Run } from './types.js';

// Issue #797 — dictionary line breaking for Thai/Lao/Khmer in xlsx. Both the
// plain path (`wrapParagraphLines`, lines re-concatenated into one drawn string)
// and the rich path (`layoutRichTextLines`, one fillText per segment) must wrap a
// spaceless SEA run at segmenter word boundaries without tearing a word, and the
// rich path must keep each line's SEA text as ONE segment (measure==paint).
// char = 10px in the stub; word boundaries come from the platform ICU dictionary.

const ctx = {
  font: '',
  measureText: (s: string) => ({ width: [...s].length * 10 }),
} as unknown as CanvasRenderingContext2D;

const baseFont: CellFont = {
  bold: false, italic: false, underline: false, strike: false, size: 11, color: null, name: null,
};

const richText = (lines: ReturnType<typeof layoutRichTextLines>): string[] =>
  lines.map((l) => l.segments.map((s) => s.text).join(''));

function breakOffsets(texts: string[]): number[] {
  const out: number[] = [];
  let acc = 0;
  for (let i = 0; i < texts.length - 1; i++) { acc += texts[i].length; out.push(acc); }
  return out;
}

const thai = 'ภาษาไทยเป็นภาษาที่สวยงามมาก';
const wordStarts = new Set(seaWordBreakOffsets(thai));

describe('xlsx wrapParagraphLines — SEA (Thai) dictionary breaking', () => {
  it('breaks only at dictionary word boundaries and preserves the text', () => {
    const out = wrapParagraphLines(ctx, thai, 100); // 10 cp per line
    expect(out.length).toBeGreaterThan(1);
    expect(out.join('')).toBe(thai);
    for (const b of breakOffsets(out)) expect(wordStarts.has(b)).toBe(true);
  });

  it('keeps a Thai run that fits on one line intact', () => {
    expect(wrapParagraphLines(ctx, thai, 400)).toEqual([thai]);
  });

  it('does not change non-SEA wrapping (plain CJK parity)', () => {
    expect(wrapParagraphLines(ctx, 'あいうえお', 30)).toEqual(['あいう', 'えお']);
  });
});

describe('xlsx layoutRichTextLines — SEA (Thai) dictionary breaking', () => {
  it('breaks only at word boundaries, preserves text, one segment per wrapped line', () => {
    const lines = layoutRichTextLines(ctx, [{ text: thai }] as Run[], baseFont, 1, 100);
    const texts = richText(lines);
    expect(texts.length).toBeGreaterThan(1);
    expect(texts.join('')).toBe(thai);
    for (const b of breakOffsets(texts)) expect(wordStarts.has(b)).toBe(true);
    // Each line is a single contiguous draw (measure==paint), not per-word segments.
    for (const l of lines) expect(l.segments.length).toBeLessThanOrEqual(1);
  });

  it('lays a Thai run that fits on one line as a single segment', () => {
    const lines = layoutRichTextLines(ctx, [{ text: thai }] as Run[], baseFont, 1, 400);
    expect(lines).toHaveLength(1);
    expect(lines[0].segments).toHaveLength(1);
    expect(lines[0].segments[0].text).toBe(thai);
  });

  it('makes progress and preserves text when the cell is narrower than one word', () => {
    const lines = layoutRichTextLines(ctx, [{ text: thai }] as Run[], baseFont, 1, 15);
    const texts = richText(lines);
    expect(texts.join('')).toBe(thai);
    expect(texts.every((t) => t.length > 0)).toBe(true);
  });

  it('does not change non-SEA wrapping (plain CJK parity)', () => {
    const lines = layoutRichTextLines(ctx, [{ text: 'あいうえお' }] as Run[], baseFont, 1, 30);
    expect(richText(lines)).toEqual(['あいう', 'えお']);
  });
});
