import { describe, expect, it } from 'vitest';
import { DEFAULT_KINSOKU_RULES, seaWordBreakOffsets } from '@silurus/ooxml-core';
import { layoutLines, type LayoutLine, type LayoutSeg, type LayoutTextSeg } from './line-layout.js';

// Issue #797 — dictionary-based line breaking for Thai/Lao/Khmer (no inter-word
// spaces). The docx wrap loop must break a SEA run only at a segmenter word
// boundary, never mid-word, and never lose/duplicate text. The stub metric is
// 5px per code point (see makeLinearCtx), so break points are deterministic w.r.t.
// widths; the WORD boundaries come from the platform ICU dictionary, and we
// assert the PROPERTY (every break ∈ seaWordBreakOffsets) rather than exact
// offsets so the test is robust to ICU version drift.

function makeLinearCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const fontSize = (): number => Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (text: string) => {
      const size = fontSize();
      return {
        width: [...text].length * size * 0.5,
        fontBoundingBoxAscent: size * 0.8,
        fontBoundingBoxDescent: size * 0.2,
        actualBoundingBoxAscent: size * 0.8,
        actualBoundingBoxDescent: size * 0.2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

function textSeg(text: string): LayoutTextSeg {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'T', vertAlign: null, measuredWidth: 0,
  } as unknown as LayoutTextSeg;
}

function lay(segs: LayoutSeg[], width: number): LayoutLine[] {
  return layoutLines(
    makeLinearCtx(), segs, width, 0, 1, [], undefined, {}, 0,
    DEFAULT_KINSOKU_RULES, 0, 36, width, false,
  );
}

const lineTexts = (lines: LayoutLine[]): string[] =>
  lines.map((l) => l.segments.filter((s): s is LayoutTextSeg => 'text' in s).map((s) => s.text).join(''));

/** Cumulative UTF-16 offset at the END of each non-final line — i.e. every point
 *  where the paragraph was broken. */
function breakOffsets(texts: string[]): number[] {
  const out: number[] = [];
  let acc = 0;
  for (let i = 0; i < texts.length - 1; i++) {
    acc += texts[i].length;
    out.push(acc);
  }
  return out;
}

describe('SEA (Thai) dictionary line breaking in docx layoutLines', () => {
  const thai = 'ภาษาไทยเป็นภาษาที่สวยงามมาก'; // one spaceless Thai run
  const wordStarts = new Set(seaWordBreakOffsets(thai)); // legal break-before offsets

  it('has dictionary break opportunities to work with', () => {
    expect(wordStarts.size).toBeGreaterThan(1);
  });

  it('breaks only at dictionary word boundaries and preserves the text', () => {
    const lines = lay([textSeg(thai)], 50); // 10 code points fit per line
    const texts = lineTexts(lines);
    expect(texts.length).toBeGreaterThan(1); // it actually wrapped
    expect(texts.join('')).toBe(thai); // nothing lost or duplicated
    for (const b of breakOffsets(texts)) {
      expect(wordStarts.has(b)).toBe(true); // never mid-word
    }
  });

  it('packs as many whole words as fit before each break', () => {
    // Width 50 = 10 code points. First line takes the widest word run ≤ 10 cp.
    const lines = lay([textSeg(thai)], 50);
    const texts = lineTexts(lines);
    for (const t of texts.slice(0, -1)) {
      expect([...t].length).toBeLessThanOrEqual(10);
    }
    // And it packed more than a single (short) first word onto line 1.
    expect([...texts[0]].length).toBeGreaterThan(4);
  });

  it('makes progress and preserves text when the band is narrower than one word (emergency split)', () => {
    // Width 8 < first word (ภาษา = 4 cp = 20px): grapheme-safe emergency split.
    const lines = lay([textSeg(thai)], 8);
    const texts = lineTexts(lines);
    expect(texts.join('')).toBe(thai);
    expect(texts.every((t) => t.length > 0)).toBe(true); // no empty line / infinite loop
  });

  it('keeps a Thai run that fits on one line as a single contiguous segment (measure==paint)', () => {
    const lines = lay([textSeg(thai)], 200); // 27 cp = 135px < 200 — fits
    expect(lines).toHaveLength(1);
    const segs = lines[0].segments.filter((s): s is LayoutTextSeg => 'text' in s);
    expect(segs).toHaveLength(1); // ONE draw, not fragmented per word
    expect(segs[0].text).toBe(thai);
  });

  it('splits a single over-long SEA word (no dictionary boundary) grapheme-safely', () => {
    const word = 'ภาษา'; // one Thai dictionary word → seaWordBreakOffsets = []
    expect(seaWordBreakOffsets(word)).toEqual([]); // precondition: no interior break
    const lines = lay([textSeg(word)], 8); // 8px < 4 cp × 5px = 20px
    const texts = lineTexts(lines);
    expect(texts.join('')).toBe(word); // text preserved
    expect(texts.every((t) => t.length > 0)).toBe(true); // progress, no infinite loop
  });

  it('does not alter non-SEA (Latin) wrapping', () => {
    // Two Latin words; 'hello ' fit-width collapses trailing space. Byte-identical
    // to the pre-existing behavior (no SEA code path entered).
    const lines = lay([textSeg('hello '), textSeg('world')], 30);
    expect(lineTexts(lines).join('')).toContain('hello');
    expect(lineTexts(lines).join('')).toContain('world');
  });
});
