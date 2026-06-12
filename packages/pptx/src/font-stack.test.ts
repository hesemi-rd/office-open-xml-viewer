import { describe, it, expect } from 'vitest';
import { cssFontStack } from './renderer.js';

describe('cssFontStack — Arabic faces keep the Arabic chain (regression)', () => {
  it('leads with the Arabic Noto fallbacks for an Arabic-script face', () => {
    // OFFICE_FONT_SUBSTITUTE maps Sakkal Majalla → Noto Naskh Arabic.
    const stack = cssFontStack('Sakkal Majalla');
    expect(stack.startsWith('"Sakkal Majalla", "Noto Naskh Arabic"')).toBe(true);
    expect(stack).toContain('"Noto Sans Arabic"');
    // No CJK / non-CJK script tail injected before the generic for Arabic.
    expect(stack).not.toContain('Noto Sans KR');
    expect(stack).not.toContain('Noto Sans Thai');
  });
});

describe('cssFontStack — CJK language-specific Noto ordering', () => {
  it('Korean sans (Malgun Gothic) → Noto Sans KR leads the CJK tail', () => {
    const stack = cssFontStack('Malgun Gothic');
    expect(stack).toContain('"Noto Sans KR"');
    expect(stack.indexOf('Noto Sans KR')).toBeLessThan(stack.indexOf('Noto Sans JP'));
    expect(stack.endsWith('sans-serif')).toBe(true);
  });

  it('Simplified Chinese serif (SimSun) → Noto Serif SC leads', () => {
    const stack = cssFontStack('SimSun');
    expect(stack).toContain('"Noto Serif SC"');
    expect(stack.indexOf('Noto Serif SC')).toBeLessThan(stack.indexOf('Noto Serif JP'));
    expect(stack.endsWith('serif')).toBe(true);
  });

  it('Simplified Chinese sans (Microsoft YaHei) → Noto Sans SC leads', () => {
    const stack = cssFontStack('Microsoft YaHei');
    expect(stack).toContain('"Noto Sans SC"');
    expect(stack.indexOf('Noto Sans SC')).toBeLessThan(stack.indexOf('Noto Sans JP'));
  });

  it('Traditional Chinese (Microsoft JhengHei) → Noto Sans TC leads', () => {
    const stack = cssFontStack('Microsoft JhengHei');
    expect(stack).toContain('"Noto Sans TC"');
    expect(stack.indexOf('Noto Sans TC')).toBeLessThan(stack.indexOf('Noto Sans SC'));
  });

  it('Japanese faces stay on Noto JP (regression — Yu Gothic, Meiryo)', () => {
    expect(cssFontStack('Yu Gothic')).toContain('"Noto Sans JP"');
    expect(cssFontStack('Meiryo')).toContain('"Noto Sans JP"');
  });
});

describe('cssFontStack — non-CJK scripts appended to Latin faces', () => {
  it('adds Hebrew / Thai / Devanagari Notos to a plain Latin sans face', () => {
    const stack = cssFontStack('Arial');
    expect(stack).toContain('"Noto Sans Hebrew"');
    expect(stack).toContain('"Noto Sans Thai"');
    expect(stack).toContain('"Noto Sans Devanagari"');
    expect(stack).toContain('"Noto Sans"'); // Cyrillic coverage
    expect(stack.endsWith('sans-serif')).toBe(true);
  });

  it('adds Hebrew serif Noto to a serif face', () => {
    const stack = cssFontStack('Times New Roman');
    expect(stack).toContain('"Noto Serif Hebrew"');
    expect(stack).toContain('"Noto Serif"');
    expect(stack.endsWith('serif')).toBe(true);
  });
});
