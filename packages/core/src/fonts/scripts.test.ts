import { describe, it, expect } from 'vitest';
import {
  classifyCjkFont,
  cjkFallbackChain,
  NON_CJK_SANS_FALLBACKS,
  NON_CJK_SERIF_FALLBACKS,
  SCRIPT_GOOGLE_FONTS,
  SCRIPT_PRELOAD_NAMES,
} from './scripts.js';

describe('classifyCjkFont — Office font name → CJK language', () => {
  it('classifies Korean faces (Malgun Gothic, Batang, Gulim, Dotum, 돋움)', () => {
    expect(classifyCjkFont('Malgun Gothic')).toBe('kr');
    expect(classifyCjkFont('Batang')).toBe('kr');
    expect(classifyCjkFont('Gulim')).toBe('kr');
    expect(classifyCjkFont('Dotum')).toBe('kr');
    expect(classifyCjkFont('돋움')).toBe('kr');
    expect(classifyCjkFont('맑은 고딕')).toBe('kr');
  });

  it('classifies Simplified Chinese faces (SimSun, 宋体, Microsoft YaHei, DengXian, FangSong, KaiTi)', () => {
    expect(classifyCjkFont('SimSun')).toBe('sc');
    expect(classifyCjkFont('宋体')).toBe('sc');
    expect(classifyCjkFont('Microsoft YaHei')).toBe('sc');
    expect(classifyCjkFont('微软雅黑')).toBe('sc');
    expect(classifyCjkFont('DengXian')).toBe('sc');
    expect(classifyCjkFont('等线')).toBe('sc');
    expect(classifyCjkFont('FangSong')).toBe('sc');
    expect(classifyCjkFont('KaiTi')).toBe('sc');
    expect(classifyCjkFont('SimHei')).toBe('sc');
  });

  it('classifies Traditional Chinese faces (PMingLiU, 新細明體, Microsoft JhengHei, MingLiU, DFKai-SB)', () => {
    expect(classifyCjkFont('PMingLiU')).toBe('tc');
    expect(classifyCjkFont('新細明體')).toBe('tc');
    expect(classifyCjkFont('Microsoft JhengHei')).toBe('tc');
    expect(classifyCjkFont('微軟正黑體')).toBe('tc');
    expect(classifyCjkFont('MingLiU')).toBe('tc');
    expect(classifyCjkFont('DFKai-SB')).toBe('tc');
    expect(classifyCjkFont('標楷體')).toBe('tc');
  });

  it('classifies Japanese faces (Yu Gothic, Meiryo, MS Mincho, Hiragino) as jp', () => {
    expect(classifyCjkFont('Yu Gothic')).toBe('jp');
    expect(classifyCjkFont('游ゴシック')).toBe('jp');
    expect(classifyCjkFont('Meiryo')).toBe('jp');
    expect(classifyCjkFont('MS Mincho')).toBe('jp');
    expect(classifyCjkFont('ＭＳ 明朝')).toBe('jp');
    expect(classifyCjkFont('Hiragino Sans')).toBe('jp');
    expect(classifyCjkFont('ヒラギノ角ゴ')).toBe('jp');
  });

  it('returns null for non-CJK faces', () => {
    expect(classifyCjkFont('Arial')).toBeNull();
    expect(classifyCjkFont('Times New Roman')).toBeNull();
    expect(classifyCjkFont('Sakkal Majalla')).toBeNull();
    expect(classifyCjkFont('')).toBeNull();
  });

  it('is case-insensitive for Latin transliterations', () => {
    expect(classifyCjkFont('simsun')).toBe('sc');
    expect(classifyCjkFont('MALGUN GOTHIC')).toBe('kr');
  });
});

describe('cjkFallbackChain — language-specific Noto CJK ordering', () => {
  it('places the matching Noto CJK first (sans)', () => {
    expect(cjkFallbackChain('kr', 'sans')[0]).toBe('Noto Sans KR');
    expect(cjkFallbackChain('sc', 'sans')[0]).toBe('Noto Sans SC');
    expect(cjkFallbackChain('tc', 'sans')[0]).toBe('Noto Sans TC');
    expect(cjkFallbackChain('jp', 'sans')[0]).toBe('Noto Sans JP');
  });

  it('places the matching Noto CJK first (serif)', () => {
    expect(cjkFallbackChain('kr', 'serif')[0]).toBe('Noto Serif KR');
    expect(cjkFallbackChain('sc', 'serif')[0]).toBe('Noto Serif SC');
    expect(cjkFallbackChain('tc', 'serif')[0]).toBe('Noto Serif TC');
    expect(cjkFallbackChain('jp', 'serif')[0]).toBe('Noto Serif JP');
  });

  it('includes the other CJK languages after the matching one (so shared Han still resolves)', () => {
    const kr = cjkFallbackChain('kr', 'sans');
    expect(kr).toContain('Noto Sans JP');
    expect(kr).toContain('Noto Sans SC');
    expect(kr.indexOf('Noto Sans KR')).toBeLessThan(kr.indexOf('Noto Sans JP'));
  });
});

describe('non-CJK fallback constants', () => {
  it('sans chain covers Cyrillic (Noto Sans), Thai, Devanagari, Hebrew', () => {
    expect(NON_CJK_SANS_FALLBACKS).toContain('Noto Sans');
    expect(NON_CJK_SANS_FALLBACKS).toContain('Noto Sans Thai');
    expect(NON_CJK_SANS_FALLBACKS).toContain('Noto Sans Devanagari');
    expect(NON_CJK_SANS_FALLBACKS).toContain('Noto Sans Hebrew');
  });

  it('serif chain covers Cyrillic (Noto Serif) and Hebrew', () => {
    expect(NON_CJK_SERIF_FALLBACKS).toContain('Noto Serif');
    expect(NON_CJK_SERIF_FALLBACKS).toContain('Noto Serif Hebrew');
  });
});

describe('SCRIPT_GOOGLE_FONTS / SCRIPT_PRELOAD_NAMES', () => {
  it('maps every script Noto family to a Google Fonts URL', () => {
    for (const name of SCRIPT_PRELOAD_NAMES) {
      const entry = SCRIPT_GOOGLE_FONTS[name.toLowerCase()];
      expect(entry, `missing Google Fonts entry for ${name}`).toBeDefined();
      expect(entry.url).toMatch(/fonts\.googleapis\.com/);
    }
  });

  it('includes CJK, Cyrillic-covering, Thai, Devanagari, Hebrew families', () => {
    expect(SCRIPT_GOOGLE_FONTS['noto sans kr']).toBeDefined();
    expect(SCRIPT_GOOGLE_FONTS['noto serif sc']).toBeDefined();
    expect(SCRIPT_GOOGLE_FONTS['noto sans thai']).toBeDefined();
    expect(SCRIPT_GOOGLE_FONTS['noto sans devanagari']).toBeDefined();
    expect(SCRIPT_GOOGLE_FONTS['noto sans hebrew']).toBeDefined();
    expect(SCRIPT_GOOGLE_FONTS['noto serif hebrew']).toBeDefined();
  });
});
