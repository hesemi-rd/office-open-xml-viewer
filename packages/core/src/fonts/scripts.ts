/**
 * Shared script → Noto fallback definitions used by the docx / pptx / xlsx
 * renderers and their Google-Fonts preload maps.
 *
 * The library ships NO font binaries; these Noto families are loaded on demand
 * from the Google Fonts CDN only when the caller opts in with
 * `useGoogleFonts: true`. The renderers append the relevant Noto family names
 * to the CSS font stack so the browser's per-glyph fallback resolves CJK /
 * Cyrillic / Thai / Devanagari / Hebrew glyphs to a real web font instead of an
 * oversized OS face or tofu.
 *
 * ## Why CJK is handled differently from the other scripts
 *
 * KR / SC / TC / JP share the Han (漢字) block but render those shared
 * codepoints with DIFFERENT glyph shapes (e.g. 直/海/骨 differ between
 * Japanese, Simplified Chinese, Traditional Chinese, Korean conventions).
 * Canvas text falls back PER GLYPH down the font-family chain, so whichever
 * Noto CJK comes first in the chain decides the shape of every shared Han
 * glyph. We therefore cannot just append all four — we must put the Noto CJK
 * matching the document's CJK language FIRST. {@link classifyCjkFont} derives
 * that language from the requested Office font name (and CJK-only Unicode
 * markers in the name). The other three CJK Notos still follow, so a stray
 * codepoint missing from the primary face (rare) resolves to a sibling rather
 * than tofu.
 *
 * The non-CJK scripts (Cyrillic, Thai, Devanagari, Hebrew) do NOT share glyphs
 * with Latin or with each other, so the browser's per-glyph fallback picks the
 * right face no matter the order. They are appended unconditionally to the
 * generic sans / serif tails ({@link NON_CJK_SANS_FALLBACKS} /
 * {@link NON_CJK_SERIF_FALLBACKS}).
 *
 * Hebrew is RTL; only the FONT is supplied here — bidi/visual ordering is
 * handled by the existing bidi line logic in each package.
 */

export type CjkLang = 'kr' | 'sc' | 'tc' | 'jp';
export type FontVariant = 'sans' | 'serif';

/**
 * Classify a requested Office font name into a CJK language, or `null` when it
 * is not a known CJK face. Based on the well-known default East-Asian fonts
 * Office ships per locale (names verified against the Windows/Office font set):
 *
 * - Korean  : Malgun Gothic (맑은 고딕), Batang/Batangche, Gulim/Gulimche,
 *             Dotum/Dotumche (돋움), Gungsuh, New Gulim. Plus the Hangul
 *             Jamo/Syllables marker in the name.
 * - Simplified Chinese : SimSun (宋体), NSimSun, SimHei (黑体), Microsoft YaHei
 *             (微软雅黑), DengXian (等线), FangSong (仿宋), KaiTi (楷体), STSong,
 *             STKaiti, STFangsong, STHeiti, STXihei, LiSu, YouYuan.
 * - Traditional Chinese : PMingLiU/MingLiU (細明體/新細明體), Microsoft JhengHei
 *             (微軟正黑體), DFKai-SB (標楷體), MingLiU_HKSCS, Kaiti TC.
 * - Japanese : Yu Gothic/Mincho (游ゴシック/游明朝), Meiryo (メイリオ),
 *             MS Gothic/Mincho (ＭＳ ゴシック/明朝), Hiragino (ヒラギノ),
 *             plus the Hiragana/Katakana marker in the name.
 *
 * Detection is name + Unicode-marker based (no per-sample tuning). Where Latin
 * transliterations collide we disambiguate on the more specific token (e.g.
 * "Microsoft JhengHei" → TC vs "Microsoft YaHei" → SC; "MingLiU" → TC).
 */
export function classifyCjkFont(family: string | null | undefined): CjkLang | null {
  if (!family) return null;
  const l = family.toLowerCase();

  // --- Unicode-script markers in the name (script-exclusive ranges) ---
  // Hangul (Jamo U+1100–11FF, Compatibility Jamo U+3130–318F, Syllables
  // U+AC00–D7AF) → Korean. 돋움 / 맑은 고딕 etc.
  if (/[ᄀ-ᇿ㄰-㆏가-힯]/.test(family)) return 'kr';
  // Hiragana (U+3040–309F) / Katakana (U+30A0–30FF) → Japanese. メイリオ etc.
  if (/[぀-ヿ]/.test(family)) return 'jp';

  // --- Latin transliterations + Han names. Order matters: TC tokens that
  //     contain an SC-looking substring are matched first. ---

  // Traditional Chinese (check before SC: "jhenghei" must not be caught as a
  // generic "hei", and MingLiU family before any "ming" heuristics).
  if (
    /jhenghei|微軟正黑|新細明|細明|pmingliu|mingliu|dfkai|標楷|華康|cns11643|kaiti tc|ming\s*liu/.test(l) ||
    /新細明體|細明體|標楷體|微軟正黑體|華康/.test(family)
  ) {
    return 'tc';
  }

  // Simplified Chinese
  if (
    /simsun|nsimsun|simhei|simkai|simfang|yahei|dengxian|fangsong|kaiti|youyuan|lisu|stsong|stkaiti|stfangsong|stheiti|stxihei|stzhongsong|songti sc|heiti sc|微软雅黑/.test(l) ||
    /宋体|黑体|楷体|仿宋|等线|微软雅黑|隶书|幼圆/.test(family)
  ) {
    return 'sc';
  }

  // Korean (Latin transliterations)
  if (/malgun|batang|gulim|dotum|gungsuh|nanum|new gulim|hancom|hy(gothic|graphic|namu)?/.test(l)) {
    return 'kr';
  }

  // Japanese (Latin transliterations). Yu/MS Gothic+Mincho, Meiryo, Hiragino.
  if (
    /\bmeiryo\b|\byu\s*(gothic|mincho)\b|yugothic|yumincho|hiragino|\bms\s*(gothic|mincho|pgothic|pmincho|ui\s*gothic)\b|\bms[pg]?(gothic|mincho)\b|ipa(ex)?(gothic|mincho)|noto\s+(sans|serif)\s+jp|游ゴシック|游明朝|ＭＳ|メイリオ|ヒラギノ/.test(l) ||
    /游ゴシック|游明朝|ＭＳ ゴシック|ＭＳ 明朝|ＭＳ Ｐゴシック|メイリオ|ヒラギノ/.test(family)
  ) {
    return 'jp';
  }

  return null;
}

/**
 * Ordered Noto CJK family names for a given language, primary face first so the
 * browser resolves shared Han glyphs to that language's shapes. The other three
 * follow as a last-resort so a codepoint absent from the primary face does not
 * fall to tofu.
 */
export function cjkFallbackChain(lang: CjkLang, variant: FontVariant): string[] {
  const prefix = variant === 'serif' ? 'Noto Serif' : 'Noto Sans';
  const order: Record<CjkLang, CjkLang[]> = {
    // Primary language first; the rest in a stable order. The trailing list is
    // a tofu safety net only — shared Han glyphs are already taken by the head.
    jp: ['jp', 'sc', 'tc', 'kr'],
    sc: ['sc', 'tc', 'jp', 'kr'],
    tc: ['tc', 'sc', 'jp', 'kr'],
    kr: ['kr', 'jp', 'sc', 'tc'],
  };
  const suffix: Record<CjkLang, string> = { kr: 'KR', sc: 'SC', tc: 'TC', jp: 'JP' };
  return order[lang].map((x) => `${prefix} ${suffix[x]}`);
}

/**
 * Noto faces appended to a SANS generic tail so non-CJK, non-Latin scripts
 * resolve to a real web font via per-glyph fallback. Cyrillic is covered by the
 * un-suffixed "Noto Sans"; Thai / Devanagari / Hebrew need their script-specific
 * Noto family because "Noto Sans" does not embed those scripts.
 *
 * No glyph-shape collision with Latin or with each other, so order is
 * immaterial and these are safe to append unconditionally.
 */
export const NON_CJK_SANS_FALLBACKS = [
  'Noto Sans',          // Latin + Cyrillic + Greek
  'Noto Sans Hebrew',   // Hebrew (RTL; bidi handled separately)
  'Noto Sans Thai',
  'Noto Sans Devanagari',
] as const;

/** Serif counterpart of {@link NON_CJK_SANS_FALLBACKS}. Devanagari/Thai ship
 *  primarily as sans in Office serif contexts; Cyrillic + Hebrew get serif
 *  Noto faces. (Noto Serif Thai/Devanagari are omitted to keep the serif tail
 *  small — sans coverage above still prevents tofu for those scripts.) */
export const NON_CJK_SERIF_FALLBACKS = [
  'Noto Serif',         // Latin + Cyrillic + Greek
  'Noto Serif Hebrew',  // Hebrew (RTL; bidi handled separately)
] as const;

// ---- Google Fonts preload definitions (script Noto families) ----

const w = (q: string) =>
  `https://fonts.googleapis.com/css2?family=${q}:wght@400;700&display=swap`;

/**
 * Lower-cased Noto family name → Google Fonts CSS URL for every script face the
 * renderers may reference. Merged into each package's `*_GOOGLE_FONTS` map so a
 * single source of truth defines the URLs. `loadFamily` is omitted because
 * Google Fonts serves these under the same family name we request.
 */
export const SCRIPT_GOOGLE_FONTS: Record<string, { url: string; loadFamily?: string }> = {
  // CJK — sans + serif per language.
  'noto sans kr': { url: w('Noto+Sans+KR') },
  'noto sans sc': { url: w('Noto+Sans+SC') },
  'noto sans tc': { url: w('Noto+Sans+TC') },
  'noto sans jp': { url: w('Noto+Sans+JP') },
  'noto serif kr': { url: w('Noto+Serif+KR') },
  'noto serif sc': { url: w('Noto+Serif+SC') },
  'noto serif tc': { url: w('Noto+Serif+TC') },
  'noto serif jp': { url: w('Noto+Serif+JP') },
  // Latin/Cyrillic/Greek base (also the non-CJK sans/serif tail head).
  'noto sans': { url: w('Noto+Sans') },
  'noto serif': { url: w('Noto+Serif') },
  // Indic / SE-Asian.
  'noto sans devanagari': { url: w('Noto+Sans+Devanagari') },
  'noto sans thai': { url: w('Noto+Sans+Thai') },
  // Hebrew (RTL).
  'noto sans hebrew': { url: w('Noto+Sans+Hebrew') },
  'noto serif hebrew': { url: w('Noto+Serif+Hebrew') },
};

/**
 * Family names always queued for preload (self-referencing) when
 * `useGoogleFonts` is on, so every Noto face the renderer might append to a
 * font stack is actually loaded even when no document font maps to it by name.
 * Mirrors the existing Arabic self-referencing entries.
 */
export const SCRIPT_PRELOAD_NAMES: string[] = [
  'Noto Sans KR', 'Noto Sans SC', 'Noto Sans TC', 'Noto Sans JP',
  'Noto Serif KR', 'Noto Serif SC', 'Noto Serif TC', 'Noto Serif JP',
  'Noto Sans', 'Noto Serif',
  'Noto Sans Devanagari', 'Noto Sans Thai',
  'Noto Sans Hebrew', 'Noto Serif Hebrew',
];
