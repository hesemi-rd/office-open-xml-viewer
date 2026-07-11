import { describe, expect, it } from 'vitest';
import { fontAdvanceBiasEm } from './font-advance-metrics.js';

describe('fontAdvanceBiasEm', () => {
  it('returns the Chromium-vs-Word Georgia bias and zero for near-real/unknown faces', () => {
    expect(fontAdvanceBiasEm('Georgia')).toBe(0.0105);
    expect(fontAdvanceBiasEm('  "georgia" ')).toBe(0.0105);
    expect(fontAdvanceBiasEm('Times New Roman')).toBe(0);
    expect(fontAdvanceBiasEm('serif')).toBe(0);
    expect(fontAdvanceBiasEm(undefined)).toBe(0);
  });
});
