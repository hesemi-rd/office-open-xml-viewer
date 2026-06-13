import { SCRIPT_GOOGLE_FONTS, SCRIPT_PRELOAD_NAMES, type FontPreloadEntry } from '@silurus/ooxml-core';
import type { Styles } from './types.js';

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

/**
 * The font-family names to preload for a workbook: every styled cell font, the
 * generic Arabic fallbacks, and every script Noto face. Office faces map to
 * metric-compatible substitutes (Calibri → Carlito, Cambria → Caladea); the
 * renderer's default chain ends with the Noto faces, so they must be loaded for
 * any Arabic/CJK/Cyrillic/Thai/Devanagari/Hebrew glyph that falls through to
 * resolve to a real web font instead of an OS face or tofu. A workbook using
 * only system fonts (no map entries) still produces zero network requests.
 *
 * Single source of truth shared by the main-thread `_load()` and the render
 * worker, so both modes preload an identical set — worker/main rendering must
 * stay pixel-equivalent.
 */
export function xlsxFontPreloadNames(styles: Styles | undefined): Set<string> {
  const names = new Set<string>();
  for (const f of styles?.fonts ?? []) {
    if (f.name) names.add(f.name);
  }
  names.add('Noto Naskh Arabic');
  names.add('Noto Sans Arabic');
  for (const n of SCRIPT_PRELOAD_NAMES) names.add(n);
  return names;
}
