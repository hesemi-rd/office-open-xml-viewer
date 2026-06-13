import { SCRIPT_GOOGLE_FONTS, SCRIPT_PRELOAD_NAMES, type FontPreloadEntry } from '@silurus/ooxml-core';

/** Theme-referenced typefaces commonly used by PPTX templates. Keys are
 *  lower-cased family names. Entries that substitute a metric-compatible
 *  family (Calibri → Carlito, Cambria → Caladea) include `loadFamily` so the
 *  FontFaceSet load is driven against the substitute; the renderer puts the
 *  substitute into the canvas font stack so missing Office fonts degrade to a
 *  same-width webfont instead of a wider system serif/sans. The remaining
 *  entries omit `loadFamily` because Google Fonts ships the same family name. */
const NOTO_NASKH_ARABIC_URL =
  'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap';
const NOTO_SANS_ARABIC_URL =
  'https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;700&display=swap';

export const PPTX_GOOGLE_FONTS: Record<string, FontPreloadEntry> = {
  'calibri':           { url: 'https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Carlito' },
  'calibri light':     { url: 'https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Carlito' },
  'cambria':           { url: 'https://fonts.googleapis.com/css2?family=Caladea:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Caladea' },
  'cambria math':      { url: 'https://fonts.googleapis.com/css2?family=Caladea:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Caladea' },
  'nunito sans':       { url: 'https://fonts.googleapis.com/css2?family=Nunito+Sans:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'nunito':            { url: 'https://fonts.googleapis.com/css2?family=Nunito:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'open sans':         { url: 'https://fonts.googleapis.com/css2?family=Open+Sans:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'roboto':            { url: 'https://fonts.googleapis.com/css2?family=Roboto:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'lato':              { url: 'https://fonts.googleapis.com/css2?family=Lato:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'montserrat':        { url: 'https://fonts.googleapis.com/css2?family=Montserrat:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'poppins':           { url: 'https://fonts.googleapis.com/css2?family=Poppins:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'raleway':           { url: 'https://fonts.googleapis.com/css2?family=Raleway:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'playfair display':  { url: 'https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  // Common Arabic-script faces that hosts rarely ship. Map them to Noto
  // substitutes so RTL slides (e.g. sample-10, which requests Sakkal Majalla /
  // Univers Next Arabic) render with a real web font instead of an oversized
  // OS fallback. "Naskh" covers traditional serif-like Arabic faces; "Sans"
  // covers the modern geometric ones.
  'sakkal majalla':      { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'traditional arabic':  { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'simplified arabic':   { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'arabic typesetting':  { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'univers next arabic': { url: NOTO_SANS_ARABIC_URL, loadFamily: 'Noto Sans Arabic' },
  // Self-referencing entries so the generic Arabic fallback fonts (appended to
  // the renderer's font stack) are themselves loaded whenever useGoogleFonts is
  // enabled — see `load`, which always queues these names.
  'noto naskh arabic':   { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'noto sans arabic':    { url: NOTO_SANS_ARABIC_URL, loadFamily: 'Noto Sans Arabic' },
  // CJK (KR/SC/TC/JP), Cyrillic (Noto Sans/Serif), Thai, Devanagari and Hebrew
  // Noto faces, shared with docx/xlsx via @silurus/ooxml-core. The renderer
  // appends these to the canvas font stack (CJK ordered by document language);
  // loaded only when useGoogleFonts is on — no binaries ship in the bundle.
  ...SCRIPT_GOOGLE_FONTS,
};

/**
 * The font-family names to preload for a presentation: the theme major/minor
 * fonts, the generic Arabic fallbacks, and every script Noto face. The renderer's
 * canvas font stack ends with those fallbacks, so they must be loaded for any
 * Arabic/CJK/Cyrillic/Thai/Devanagari/Hebrew glyph that falls through the stack
 * to resolve to a real web font instead of an OS face or tofu.
 *
 * Single source of truth shared by the main-thread `load()` and the render
 * worker, so both modes preload an identical set — worker/main rendering must
 * stay pixel-equivalent.
 */
export function pptxFontPreloadNames(
  majorFont: string | null | undefined,
  minorFont: string | null | undefined,
): (string | null | undefined)[] {
  return [majorFont, minorFont, 'Noto Naskh Arabic', 'Noto Sans Arabic', ...SCRIPT_PRELOAD_NAMES];
}
