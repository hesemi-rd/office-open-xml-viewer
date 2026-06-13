import type { FontPreloadEntry } from '@silurus/ooxml-core';

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
};
