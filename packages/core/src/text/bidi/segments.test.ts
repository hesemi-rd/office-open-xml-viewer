import { describe, it, expect } from 'vitest';
import { toVisualSegments, resolveBaseDirection } from './segments.js';
import type { StyledRun } from './types.js';

const mk = (text: string, over: Partial<StyledRun> = {}): StyledRun => ({
  text,
  fontFamily: 'Arial',
  bold: false,
  italic: false,
  fontSizePx: 16,
  meta: undefined,
  ...over,
});

describe('toVisualSegments', () => {
  it('passes pure LTR text through as one segment', () => {
    const segs = toVisualSegments([mk('Hello')], 'ltr');
    expect(segs).toHaveLength(1);
    expect(segs[0].text).toBe('Hello');
    expect(segs[0].isRTL).toBe(false);
    expect(segs[0].parts).toHaveLength(1);
  });

  it('marks pure RTL text as one RTL segment', () => {
    const segs = toVisualSegments([mk('אבג')], 'rtl');
    expect(segs).toHaveLength(1);
    expect(segs[0].text).toBe('אבג');
    expect(segs[0].isRTL).toBe(true);
  });

  it('orders mixed bidi runs visually and preserves logical text', () => {
    // Logical: א AB ב  (Hebrew, Latin, Hebrew) in an RTL paragraph.
    const segs = toVisualSegments([mk('א'), mk('AB'), mk('ב')], 'rtl');
    // Visual (left-to-right): ב | AB | א
    expect(segs.map((s) => s.text)).toEqual(['ב', 'AB', 'א']);
    expect(segs.map((s) => s.isRTL)).toEqual([true, false, true]);
    // Reconstructing in logical order (by logicalStart) reproduces the input.
    const logical = [...segs].sort((a, b) => a.logicalStart - b.logicalStart);
    expect(logical.map((s) => s.text).join('')).toBe('אABב');
  });

  it('keeps a word split only by color in ONE segment (joining preserved)', () => {
    const segs = toVisualSegments(
      [mk('مر', { meta: 'red' }), mk('حبا', { meta: 'blue' })],
      'rtl',
    );
    expect(segs).toHaveLength(1);
    expect(segs[0].text).toBe('مرحبا');
    expect(segs[0].parts).toHaveLength(2);
    expect(segs[0].parts.map((p) => p.run.meta)).toEqual(['red', 'blue']);
  });

  it('retains X9-removed join controls (ZWJ) in the drawn text', () => {
    // BN-class characters (ZWJ etc.) are removed by rule X9 for LEVEL purposes
    // but must stay in the text passed to the shaper (emoji ZWJ sequences,
    // Arabic join control). Regression: they were dropped and split the atom.
    const family = '\u{1F469}\u200D\u{1F469}\u200D\u{1F467}'; // woman-woman-girl ZWJ sequence
    const segs = toVisualSegments([mk(family)], 'ltr');
    expect(segs).toHaveLength(1);
    expect(segs[0].text).toBe(family);
  });

  it('splits a word at a shape-affecting (font-size) change', () => {
    const segs = toVisualSegments(
      [mk('مر', { fontSizePx: 16 }), mk('حبا', { fontSizePx: 20 })],
      'rtl',
    );
    expect(segs).toHaveLength(2);
    // RTL: the later run ('حبا') is visually to the left.
    expect(segs.map((s) => s.text)).toEqual(['حبا', 'مر']);
  });
});

describe('resolveBaseDirection', () => {
  it('honors an explicit flag', () => {
    expect(resolveBaseDirection(true, 'Hello')).toBe('rtl');
    expect(resolveBaseDirection(false, 'مرحبا')).toBe('ltr');
  });
  it('falls back to UAX#9 first-strong when auto/undefined', () => {
    expect(resolveBaseDirection(undefined, 'مرحبا')).toBe('rtl');
    expect(resolveBaseDirection('auto', 'Hello')).toBe('ltr');
    expect(resolveBaseDirection(undefined, '  123 مرحبا')).toBe('rtl'); // first strong is Arabic
    expect(resolveBaseDirection(undefined, '123')).toBe('ltr'); // no strong -> LTR
  });
});
