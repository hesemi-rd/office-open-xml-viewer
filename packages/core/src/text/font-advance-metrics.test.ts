import { describe, expect, it } from 'vitest';
import {
  fontAdvanceBiasEm,
  fontScriptAdvanceScale,
  splitFontAdvanceRuns,
} from './font-advance-metrics.js';

describe('fontScriptAdvanceScale', () => {
  it('returns Meiryo UI hmtx advance ratios by script class', () => {
    expect(fontScriptAdvanceScale('Meiryo UI', 'ひ')).toBeCloseTo(0.7775, 10);
    expect(fontScriptAdvanceScale('Meiryo UI', 'カ')).toBeCloseTo(0.7438, 10);
    expect(fontScriptAdvanceScale('Meiryo UI', '。')).toBeCloseTo(0.7214, 10);
    expect(fontScriptAdvanceScale('Meiryo UI', '漢')).toBe(1);
    expect(fontScriptAdvanceScale('Meiryo UI', 'Ａ')).toBe(1);
  });

  it('does not condense plain Meiryo or untabled families', () => {
    expect(fontScriptAdvanceScale('Meiryo', 'ひ')).toBe(1);
    expect(fontScriptAdvanceScale('serif', 'ひ')).toBe(1);
    expect(fontScriptAdvanceScale(null, 'ひ')).toBe(1);
  });
});

describe('splitFontAdvanceRuns', () => {
  it('splits a tabled face into homogeneous scale runs', () => {
    expect(splitFontAdvanceRuns('Meiryo UI', '漢ひらカナ。Ａ')).toEqual([
      { text: '漢', scale: 1 },
      { text: 'ひら', scale: 0.7775 },
      { text: 'カナ', scale: 0.7438 },
      { text: '。', scale: 0.7214 },
      { text: 'Ａ', scale: 1 },
    ]);
  });

  it('keeps untabled text byte-stable as one no-op run', () => {
    expect(splitFontAdvanceRuns('Yu Gothic', '漢ひらカナ。Ａ')).toEqual([
      { text: '漢ひらカナ。Ａ', scale: 1 },
    ]);
  });
});

describe('fontAdvanceBiasEm', () => {
  it('returns the Chromium-vs-Word Georgia bias and zero for near-real/unknown faces', () => {
    expect(fontAdvanceBiasEm('Georgia')).toBe(0.0105);
    expect(fontAdvanceBiasEm('  "georgia" ')).toBe(0.0105);
    expect(fontAdvanceBiasEm('Times New Roman')).toBe(0);
    expect(fontAdvanceBiasEm('serif')).toBe(0);
    expect(fontAdvanceBiasEm(undefined)).toBe(0);
  });
});
