import { describe, it, expect } from 'vitest';
import { normalizeFontFamily } from './renderer.js';

describe('normalizeFontFamily — Arabic substitute fonts', () => {
  it('puts the Arabic substitute first so Latin/digits resolve from the same family as Arabic', () => {
    // Sakkal Majalla is family="auto" in fontTable; the run carries both
    // Arabic glyphs and Latin/digits. The Arabic substitute must lead the chain
    // (before any CJK sans face) so Latin/digits don't leak to Noto Sans JP.
    const chain = normalizeFontFamily('Sakkal Majalla');
    expect(chain.startsWith('"Sakkal Majalla", "Noto Naskh Arabic"')).toBe(true);
    // No CJK sans face before the Arabic substitute.
    const naskhIdx = chain.indexOf('Noto Naskh Arabic');
    const cjkIdx = chain.indexOf('Noto Sans JP');
    expect(naskhIdx).toBeGreaterThan(0);
    expect(naskhIdx).toBeLessThan(cjkIdx);
  });

  it('routes traditional Naskh faces to a serif Latin companion', () => {
    // Word's PDF export of sample-7 renders Sakkal Majalla's Latin with serifs,
    // so a serif Latin generic precedes the sans generics.
    const chain = normalizeFontFamily('Traditional Arabic');
    expect(chain).toContain('"Noto Serif"');
    expect(chain.endsWith('serif')).toBe(true);
    expect(chain.indexOf('Noto Serif')).toBeLessThan(chain.indexOf('Noto Sans JP'));
  });

  it('keys off the family, not a hardcoded string — case-insensitive', () => {
    expect(normalizeFontFamily('sakkal majalla').startsWith('"sakkal majalla", "Noto Naskh Arabic"')).toBe(true);
  });

  it('routes geometric Arabic faces (Univers Next Arabic) to a sans chain', () => {
    const chain = normalizeFontFamily('Univers Next Arabic');
    expect(chain.startsWith('"Univers Next Arabic", "Noto Sans Arabic"')).toBe(true);
    expect(chain.endsWith('sans-serif')).toBe(true);
  });

  it('leaves pure-Latin fonts unchanged (sans default chain)', () => {
    expect(normalizeFontFamily('Arial')).toBe(
      '"Arial", "Noto Sans JP", "Hiragino Sans", "Meiryo", "Noto Naskh Arabic", "Noto Sans Arabic", sans-serif',
    );
  });

  it('leaves serif Latin fonts unchanged', () => {
    expect(normalizeFontFamily('Times New Roman')).toBe(
      '"Times New Roman", "Yu Mincho", "YuMincho", "Hiragino Mincho ProN", "MS Mincho", "Noto Serif JP", "Noto Serif", "Noto Naskh Arabic", "Noto Sans Arabic", serif',
    );
  });
});
