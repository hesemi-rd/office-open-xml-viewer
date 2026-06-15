import { describe, it, expect } from 'vitest';
import { computePages } from './renderer.js';
import type { BodyElement, DocParagraph, DocxTextRun, SectionProps, PaginatedBodyElement } from './types';

// Unit tests for computePages pagination behaviour that the renderer-path VRT
// (local-only, private samples) cannot guard in CI. A deterministic stub canvas
// makes line wrapping and line heights predictable: glyph advance = charCount ×
// fontPx, and the font box = 0.8/0.2 em (so a single line is exactly fontPx tall
// with no spacing/grid). CJK characters break between any two glyphs, so a run of
// N of them wraps to ceil(N / charsPerLine) lines.

function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, textAlign: 'left' as CanvasTextAlign,
    direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

function section(overrides: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 200, pageHeight: 140,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    ...overrides,
  };
}

function textRun(text: string, fontSize: number): DocRun {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  };
  return { type: 'text', ...run } as DocRun;
}

type DocRun = DocParagraph['runs'][number];

function para(opts: { text?: string; fontSize?: number; widowControl?: boolean } = {}): BodyElement {
  const fontSize = opts.fontSize ?? 20;
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: opts.text ? [textRun(opts.text, fontSize)] : [],
    defaultFontSize: fontSize, defaultFontFamily: 'NotInMetrics',
    widowControl: opts.widowControl,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

const sliceOf = (el: PaginatedBodyElement) =>
  (el as { lineSlice?: { start: number; end: number } }).lineSlice;

describe('computePages — empty-paragraph relocation (C2: §17.3.1.29)', () => {
  it('moves an unsplittable mark-only paragraph to the next page instead of overflowing the bottom margin', () => {
    // content height = 140 - 40 = 100; each empty mark = 20px → exactly 5 per page.
    const body = Array.from({ length: 7 }, () => para()); // 7 empty paragraphs
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(2);
    expect(pages[0].length).toBe(5); // page 1 fills exactly
    expect(pages[1].length).toBe(2); // overflow relocated, NOT clipped onto page 1
    // no page holds more than its 5-line capacity (would mean an overflow)
    for (const p of pages) expect(p.length).toBeLessThanOrEqual(5);
  });
});

describe('computePages — line-boundary splitting + widowControl (C1: §17.3.1.44)', () => {
  // contentW = 160, glyph advance = fontPx; at 20px → 8 chars/line. 48 chars → 6 lines.
  // content height = 100 → 5 lines (100px) fit per page.
  const sixLineText = 'あ'.repeat(48);

  it('avoids a widow: a single trailing line is not stranded on the next page (default widowControl on)', () => {
    const pages = computePages([para({ text: sixLineText })], section(), makeCtx());
    expect(pages.length).toBe(2);
    // Greedy fit is 5 lines on page 1; widowControl pulls one down so ≥2 carry over.
    expect(sliceOf(pages[0][0])).toEqual({ start: 0, end: 4 });
    expect(sliceOf(pages[1][0])).toEqual({ start: 4, end: 6 });
  });

  it('honors w:widowControl="off": the trailing single line is allowed (matches sample-9)', () => {
    const pages = computePages([para({ text: sixLineText, widowControl: false })], section(), makeCtx());
    expect(pages.length).toBe(2);
    expect(sliceOf(pages[0][0])).toEqual({ start: 0, end: 5 }); // greedy 5 lines
    expect(sliceOf(pages[1][0])).toEqual({ start: 5, end: 6 }); // lone widow line allowed
  });
});
