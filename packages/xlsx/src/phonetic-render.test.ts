import { describe, it, expect } from 'vitest';
import { drawPhoneticBand } from './renderer.js';
import type { CellFont, PhoneticProperties, PhoneticRun, Styles } from './types.js';

// Renderer-level checks for the furigana band draw (ECMA-376 §18.4.6 / §18.4.3).
// `drawPhoneticBand` is the single entry point; the in-viewport draw path calls
// it after the base text. Here we drive it directly with a recording context so
// we can assert the reading strings, their x (band placement over the base
// glyphs), the reading font size (from `phoneticPr.fontId`), and the alignment
// spread — all without a real Canvas.

interface FillTextCall {
  text: string;
  x: number;
  y: number;
  font: string;
  letterSpacing: string;
}

/** A recording context. Every code point is `px()` wide in whatever font is
 *  active (font px parsed from the CSS font string), so base-font vs reading-
 *  font advances differ by their size, exactly like a real measurer. */
function makeRecordingCtx() {
  let font = '11px sans-serif';
  let letterSpacing = '0px';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '11');
  const calls: FillTextCall[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(v: string) { letterSpacing = v; },
    textBaseline: 'alphabetic' as CanvasTextBaseline,
    textAlign: 'left' as CanvasTextAlign,
    fillStyle: '#000' as string,
    measureText: (s: string) => ({ width: [...s].length * px() }) as TextMetrics,
    fillText(text: string, x: number, y: number) {
      calls.push({ text, x, y, font, letterSpacing });
    },
    save() {},
    restore() {},
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, calls };
}

const BASE_FONT: CellFont = { bold: false, italic: false, underline: false, strike: false, size: 11, color: null, name: null };
const READING_FONT: CellFont = { bold: false, italic: false, underline: false, strike: false, size: 6, color: null, name: null };

/** styles.fonts[0] = base 11pt, styles.fonts[2] = reading 6pt (mirrors the real
 *  sample where phoneticPr fontId=2 is a small font). */
const STYLES: Styles = {
  fonts: [BASE_FONT, BASE_FONT, READING_FONT],
  fills: [],
  borders: [],
  cellXfs: [],
  numFmts: [],
  dxfs: [],
};

const RUNS: PhoneticRun[] = [
  { sb: 0, eb: 1, text: 'カ' },
  { sb: 1, eb: 2, text: 'チョウ' },
];

function draw(pr: PhoneticProperties | undefined, runs = RUNS, baseText = '課長', baseLeftX = 100) {
  const { ctx, calls } = makeRecordingCtx();
  // baseFontStr must be the 11px base font so base-span advances are 11/cp.
  drawPhoneticBand(ctx, runs, pr, baseText, '11px sans-serif', STYLES, baseLeftX, 0, 1, '#000000');
  return calls;
}

// Reading font (fontId=2) is 6pt → round(6 × 4/3) = 8px. Base font here is the
// literal '11px sans-serif' string we pass as baseFontStr, so base-span advances
// are 11px/code-point. Font 0 (fallback) is 11pt → round(11 × 4/3) = 15px.
const READING_PX = '8px';
const FALLBACK_PX = '15px';

describe('drawPhoneticBand (§18.4.6 / §18.4.3)', () => {
  it('draws each reading string in the phoneticPr.fontId font (6pt → 8px), not the base', () => {
    const calls = draw({ fontId: 2, alignment: 'left' });
    expect(calls.map((c) => c.text)).toEqual(['カ', 'チョウ']);
    expect(calls.every((c) => c.font.includes(READING_PX))).toBe(true);
  });

  it('positions each reading over its base span (left alignment): base char is 11px wide', () => {
    const calls = draw({ fontId: 2, alignment: 'left' });
    // base char0 span starts at baseLeftX=100; char1 at 100+11=111.
    expect(calls[0].x).toBe(100);
    expect(calls[1].x).toBe(111);
  });

  it('center alignment centres the natural reading width over the base span', () => {
    const calls = draw({ fontId: 2, alignment: 'center' });
    // char0 base span [100,111) width 11; reading "カ" is 1 cp × 8px = 8.
    // centred x = 100 + (11 - 8)/2 = 101.5
    expect(calls[0].x).toBeCloseTo(101.5, 5);
  });

  it('distributed alignment sets letterSpacing to spread the reading across the span', () => {
    // Use a reading whose natural width < span so spreading applies.
    const runs: PhoneticRun[] = [{ sb: 0, eb: 2, text: 'ヤマ' }]; // 2 cps, base span = 2×11 = 22
    const calls = draw({ fontId: 2, alignment: 'distributed' }, runs, '山川');
    // natural reading width = 2 × 8 = 16; span 22; gap count 1 → extra = (22-16)/1 = 6.
    expect(calls[0].letterSpacing).toBe('6px');
    expect(calls[0].text).toBe('ヤマ');
  });

  it('defaults alignment to left when phoneticPr is absent (schema default)', () => {
    const calls = draw(undefined);
    // Falls back to fontId 0 (11pt → 15px) and left alignment.
    expect(calls[0].x).toBe(100);
    expect(calls[1].x).toBe(111);
    expect(calls.every((c) => c.font.includes(FALLBACK_PX))).toBe(true);
  });

  it('falls back to font 0 when fontId is out of bounds (§18.4.3)', () => {
    const calls = draw({ fontId: 99, alignment: 'left' });
    // fonts[99] undefined → fonts[0] (11pt → 15px).
    expect(calls.every((c) => c.font.includes(FALLBACK_PX))).toBe(true);
  });

  it('draws nothing for an empty run list', () => {
    const calls = draw({ fontId: 2 }, []);
    expect(calls).toEqual([]);
  });

  it('noControl lays readings out sequentially from the text start (not per word)', () => {
    const calls = draw({ fontId: 2, alignment: 'noControl' });
    // Reading widths (8px/cp): カ = 8, チョウ = 24. Sequential from baseLeftX=100.
    expect(calls[0].x).toBe(100);
    expect(calls[1].x).toBe(108);
  });
});
