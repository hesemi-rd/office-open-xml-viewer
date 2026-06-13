import {
  classifyCjkFont,
  scriptPreloadNamesForText,
  SCRIPT_GOOGLE_FONTS,
  type FontPreloadEntry,
  type TextBody,
} from '@silurus/ooxml-core';
import type { Presentation, SlideElement } from './types';

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

/** Yield every painted text string in a text body (paragraph runs). */
function* textBodyRuns(body: TextBody | null | undefined): Generator<string> {
  for (const p of body?.paragraphs ?? []) {
    for (const r of p.runs) {
      if (r.type === 'text') yield r.text;
    }
  }
}

/** Yield every rendered text string in the presentation: shape text, table
 *  cell text and chart labels across all slides. Speaker notes and comments are
 *  not painted on the slide, so they are excluded (the renderer ignores them). */
function* pptxTextRuns(pres: Presentation): Generator<string> {
  for (const slide of pres.slides) {
    for (const el of slide.elements as SlideElement[]) {
      if (el.type === 'shape') {
        yield* textBodyRuns(el.textBody);
      } else if (el.type === 'table') {
        for (const row of el.rows) {
          for (const cell of row.cells) yield* textBodyRuns(cell.textBody);
        }
      } else if (el.type === 'chart') {
        if (el.title) yield el.title;
        for (const c of el.categories) yield c;
        for (const s of el.series) if (s.name) yield s.name;
      }
    }
  }
}

/**
 * The font-family names to preload for a presentation: the theme major/minor
 * fonts, plus only the script-fallback Noto faces whose script the slide TEXT
 * actually contains ({@link scriptPreloadNamesForText}). The renderer's canvas
 * font stack still ends with the full Noto set, but eagerly fetching the
 * multi-MB CJK families for a deck with no CJK glyphs would block first paint
 * for nothing; an un-preloaded face loads lazily if it ever proves needed.
 *
 * Single source of truth shared by the main-thread `load()` and the render
 * worker. Both derive the set from the SAME parsed {@link Presentation}, so both
 * modes preload an identical set — worker/main rendering must stay
 * pixel-equivalent.
 */
export function pptxFontPreloadNames(
  pres: Presentation,
): (string | null | undefined)[] {
  const cjkLang =
    classifyCjkFont(pres.majorFont) ?? classifyCjkFont(pres.minorFont) ?? null;
  return [
    pres.majorFont,
    pres.minorFont,
    ...scriptPreloadNamesForText(pptxTextRuns(pres), cjkLang),
  ];
}
