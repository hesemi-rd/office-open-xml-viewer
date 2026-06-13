import {
  classifyCjkFont,
  scriptPreloadNamesForText,
  SCRIPT_GOOGLE_FONTS,
  type CjkLang,
  type FontPreloadEntry,
} from '@silurus/ooxml-core';
import type { ParsedWorkbook } from './types.js';

/** Office font name → metric-compatible Google Fonts substitute. These are
 *  the well-known pairings Microsoft and Google both publish and ship on
 *  Linux distributions: Calibri → Carlito, Cambria → Caladea (same advance
 *  widths and ascender / descender). Loading the substitute on a system
 *  that lacks the Office face keeps text width measurements close to
 *  Excel's. The substitute font-family differs from the requested name, so
 *  `loadFamily` redirects FontFaceSet loading appropriately. */
const NOTO_NASKH_ARABIC_URL =
  'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap';
const NOTO_SANS_ARABIC_URL =
  'https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;700&display=swap';

export const XLSX_GOOGLE_FONTS: Record<string, FontPreloadEntry> = {
  'calibri': {
    url: 'https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&display=swap',
    loadFamily: 'Carlito',
  },
  'cambria': {
    url: 'https://fonts.googleapis.com/css2?family=Caladea:ital,wght@0,400;0,700;1,400;1,700&display=swap',
    loadFamily: 'Caladea',
  },
  // Common Arabic-script faces that hosts rarely ship. Map them to Noto
  // substitutes so RTL workbooks (e.g. the LibreOffice-authored sample-29,
  // which requests Sakkal Majalla / Univers Next Arabic) render with a real
  // web font instead of an oversized OS fallback. "Naskh" covers traditional
  // serif-like Arabic faces; "Sans" covers the modern geometric ones.
  'sakkal majalla': { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'traditional arabic': { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'simplified arabic': { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'arabic typesetting': { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'univers next arabic': { url: NOTO_SANS_ARABIC_URL, loadFamily: 'Noto Sans Arabic' },
  // Self-referencing entries so the generic Arabic fallback fonts (appended to
  // the renderer's font chain) are themselves loaded whenever useGoogleFonts
  // is enabled — see `_load`, which always queues these names.
  'noto naskh arabic': { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'noto sans arabic': { url: NOTO_SANS_ARABIC_URL, loadFamily: 'Noto Sans Arabic' },
  // CJK (KR/SC/TC/JP), Cyrillic (Noto Sans/Serif), Thai, Devanagari and Hebrew
  // Noto faces, shared with docx/pptx via @silurus/ooxml-core. The renderer
  // chooses the CJK Noto per cell from the cell's font name; non-CJK scripts are
  // appended to the default chain. Loaded only when useGoogleFonts is on.
  ...SCRIPT_GOOGLE_FONTS,
};

/** Yield every textual cell value carried by the parsed workbook: the shared
 *  string table (`text` plus rich-text `runs[].text`). This is the bulk of a
 *  workbook's painted text and is present in BOTH main and worker at parse time
 *  (sheets parse lazily, but the shared string table is workbook-level).
 *  Numbers / dates carry no script-specific glyphs, so they are irrelevant. */
function* xlsxTextRuns(wb: ParsedWorkbook | undefined): Generator<string> {
  for (const s of wb?.sharedStrings ?? []) {
    if (s.runs && s.runs.length > 0) {
      for (const r of s.runs) yield r.text;
    } else {
      yield s.text;
    }
  }
}

/**
 * The font-family names to preload for a workbook: every styled cell font, plus
 * only the script-fallback Noto faces whose script the workbook's TEXT actually
 * contains ({@link scriptPreloadNamesForText}). Office faces map to
 * metric-compatible substitutes (Calibri → Carlito, Cambria → Caladea); the
 * renderer's default chain still ends with the full Noto set, but eagerly
 * fetching the multi-MB CJK families for a workbook that has no CJK glyphs would
 * block first paint for nothing; an un-preloaded face loads lazily if it ever
 * proves needed. A workbook using only system fonts (no map entries) still
 * produces zero network requests.
 *
 * Single source of truth shared by the main-thread `_load()` and the render
 * worker. Both derive the set from the SAME parsed {@link ParsedWorkbook}, so
 * both modes preload an identical set — worker/main rendering must stay
 * pixel-equivalent.
 */
export function xlsxFontPreloadNames(wb: ParsedWorkbook | undefined): Set<string> {
  const names = new Set<string>();
  let cjkLang: CjkLang | null = null;
  for (const f of wb?.styles?.fonts ?? []) {
    if (f.name) {
      names.add(f.name);
      cjkLang ??= classifyCjkFont(f.name);
    }
  }
  for (const n of scriptPreloadNamesForText(xlsxTextRuns(wb), cjkLang)) {
    names.add(n);
  }
  return names;
}
