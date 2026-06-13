import { describe, it, expect } from 'vitest';
import {
  classifyCjkFont,
  cjkFallbackChain,
  NON_CJK_SANS_FALLBACKS,
  NON_CJK_SERIF_FALLBACKS,
  SCRIPT_GOOGLE_FONTS,
  SCRIPT_PRELOAD_NAMES,
  scriptPreloadNamesForText,
} from './scripts.js';

/** Lower-case every name and assert it is resolvable in the preload URL map. */
function assertResolvable(names: string[]): void {
  for (const n of names) {
    expect(
      SCRIPT_GOOGLE_FONTS[n.toLowerCase()],
      `${n} missing from SCRIPT_GOOGLE_FONTS`,
    ).toBeDefined();
  }
}

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

describe('scriptPreloadNamesForText — script-aware preload set from document text', () => {
  it('returns nothing for empty / no text', () => {
    expect(scriptPreloadNamesForText([], null)).toEqual([]);
    expect(scriptPreloadNamesForText([''], null)).toEqual([]);
  });

  it('returns ZERO CJK/script names for pure Latin (ASCII + Latin-1)', () => {
    const out = scriptPreloadNamesForText(['Hello, World!', 'café résumé'], null);
    expect(out).toEqual([]);
  });

  it('detects Japanese (Hiragana/Katakana) → Noto Sans/Serif JP', () => {
    const out = scriptPreloadNamesForText(['こんにちは カタカナ'], null);
    expect(out).toEqual(['Noto Sans JP', 'Noto Serif JP']);
    assertResolvable(out);
  });

  it('detects Korean (Hangul) → Noto Sans/Serif KR', () => {
    const out = scriptPreloadNamesForText(['안녕하세요'], null);
    expect(out).toEqual(['Noto Sans KR', 'Noto Serif KR']);
    assertResolvable(out);
  });

  it('Han only with cjkLang hint sc → uses the hint', () => {
    const out = scriptPreloadNamesForText(['汉字测试'], 'sc');
    expect(out).toEqual(['Noto Sans SC', 'Noto Serif SC']);
  });

  it('Han only with no hint → defaults to jp', () => {
    const out = scriptPreloadNamesForText(['漢字'], null);
    expect(out).toEqual(['Noto Sans JP', 'Noto Serif JP']);
  });

  it('Han + Hangul → Hangul sets kr; Han resolves to the kr face (no JP emitted)', () => {
    const out = scriptPreloadNamesForText(['한국어 漢字'], 'jp');
    expect(out).toEqual(['Noto Sans KR', 'Noto Serif KR']);
  });

  it('Han + Kana → Kana sets jp even with sc hint', () => {
    const out = scriptPreloadNamesForText(['ひらがな 漢字'], 'sc');
    expect(out).toEqual(['Noto Sans JP', 'Noto Serif JP']);
  });

  it('Hangul + Kana → both KR and JP faces', () => {
    const out = scriptPreloadNamesForText(['한국 ひらがな'], null);
    expect(out).toContain('Noto Sans KR');
    expect(out).toContain('Noto Serif KR');
    expect(out).toContain('Noto Sans JP');
    expect(out).toContain('Noto Serif JP');
    assertResolvable(out);
  });

  it('detects Arabic → Noto Naskh Arabic + Noto Sans Arabic', () => {
    const out = scriptPreloadNamesForText(['مرحبا بالعالم'], null);
    expect(out).toContain('Noto Naskh Arabic');
    expect(out).toContain('Noto Sans Arabic');
  });

  it('detects Thai → Noto Sans Thai', () => {
    const out = scriptPreloadNamesForText(['สวัสดี'], null);
    expect(out).toEqual(['Noto Sans Thai']);
    assertResolvable(out);
  });

  it('detects Hebrew → Noto Sans/Serif Hebrew', () => {
    const out = scriptPreloadNamesForText(['שלום עולם'], null);
    expect(out).toContain('Noto Sans Hebrew');
    expect(out).toContain('Noto Serif Hebrew');
    assertResolvable(out);
  });

  it('detects Devanagari → Noto Sans Devanagari', () => {
    const out = scriptPreloadNamesForText(['नमस्ते'], null);
    expect(out).toEqual(['Noto Sans Devanagari']);
    assertResolvable(out);
  });

  it('detects Cyrillic → Noto Sans/Serif (which cover Cyrillic)', () => {
    const out = scriptPreloadNamesForText(['Привет мир'], null);
    expect(out).toEqual(['Noto Sans', 'Noto Serif']);
    assertResolvable(out);
  });

  it('detects Greek → Noto Sans/Serif', () => {
    const out = scriptPreloadNamesForText(['Γειά σου'], null);
    expect(out).toEqual(['Noto Sans', 'Noto Serif']);
  });

  it('mixed CJK (Japanese) + Arabic + Cyrillic', () => {
    const out = scriptPreloadNamesForText(['日本語 العربية Кириллица'], null);
    expect(out).toContain('Noto Sans JP');
    expect(out).toContain('Noto Serif JP');
    expect(out).toContain('Noto Naskh Arabic');
    expect(out).toContain('Noto Sans Arabic');
    expect(out).toContain('Noto Sans');
    expect(out).toContain('Noto Serif');
  });

  it('handles astral Han (U+20000+) via codePointAt', () => {
    const out = scriptPreloadNamesForText(['\u{20089}'], null);
    expect(out).toEqual(['Noto Sans JP', 'Noto Serif JP']);
  });

  it('is deterministic: same text yields the same set regardless of chunking', () => {
    const a = scriptPreloadNamesForText(['日本語 العربية'], null);
    const b = scriptPreloadNamesForText(['日本', '語 ', 'العربية'], null);
    expect([...a].sort()).toEqual([...b].sort());
  });

  it('every returned name is resolvable in SCRIPT_GOOGLE_FONTS (CJK + Cyrillic + scripts)', () => {
    const out = scriptPreloadNamesForText(['漢字 Привет สวัสดี नमस्ते שלום'], null);
    assertResolvable(out);
  });
});
