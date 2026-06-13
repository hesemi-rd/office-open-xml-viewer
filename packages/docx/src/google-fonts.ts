import {
  classifyCjkFont,
  scriptPreloadNamesForText,
  SCRIPT_GOOGLE_FONTS,
  type FontPreloadEntry,
} from '@silurus/ooxml-core';
import type {
  BodyElement,
  DocParagraph,
  DocTable,
  DocxDocumentModel,
} from './types.js';

/** Theme-referenced typefaces commonly used by DOCX templates. Mirrors the
 *  PPTX map — these are the well-known free webfont alternatives Microsoft
 *  Office templates pull from. Substitutes that diverge from the requested
 *  family name (Calibri → Carlito, Cambria → Caladea) include
 *  `loadFamily` so the FontFaceSet load is driven against the substitute. */
const NOTO_NASKH_ARABIC_URL =
  'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap';
const NOTO_SANS_ARABIC_URL =
  'https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;700&display=swap';

export const DOCX_GOOGLE_FONTS: Record<string, FontPreloadEntry> = {
  'calibri':           { url: 'https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Carlito' },
  'cambria':           { url: 'https://fonts.googleapis.com/css2?family=Caladea:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Caladea' },
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
  // substitutes so RTL documents (e.g. sample-7, which requests Sakkal Majalla
  // / Univers Next Arabic) render with a real web font instead of an oversized
  // OS fallback. "Naskh" covers traditional serif-like Arabic faces; "Sans"
  // covers the modern geometric ones.
  'sakkal majalla':      { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'traditional arabic':  { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'simplified arabic':   { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'arabic typesetting':  { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'univers next arabic': { url: NOTO_SANS_ARABIC_URL, loadFamily: 'Noto Sans Arabic' },
  // Self-referencing entries so the generic Arabic fallback fonts (appended to
  // the renderer's font chain) are themselves loaded whenever useGoogleFonts
  // is enabled — see `load`, which always queues these names.
  'noto naskh arabic':   { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'noto sans arabic':    { url: NOTO_SANS_ARABIC_URL, loadFamily: 'Noto Sans Arabic' },
  // CJK (KR/SC/TC/JP), Cyrillic (Noto Sans/Serif), Thai, Devanagari and Hebrew
  // Noto faces, shared with pptx/xlsx via @silurus/ooxml-core. The renderer
  // appends these to the font chain (CJK ordered by document language); they are
  // loaded only when useGoogleFonts is on — no binaries ship in the bundle.
  ...SCRIPT_GOOGLE_FONTS,
};

/** Yield every rendered text string in the document model so the preload step
 *  can detect which scripts are actually present. Walks the body, headers /
 *  footers, footnotes / endnotes and nested tables (text-only). Comments are
 *  not painted on the page, so they are excluded. */
function* docxTextRuns(doc: DocxDocumentModel): Generator<string> {
  const fromRuns = function* (runs: DocParagraph['runs']): Generator<string> {
    for (const r of runs) {
      if (r.type === 'text') yield r.text;
      else if (r.type === 'field') yield r.fallbackText;
      else if (r.type === 'shape') {
        for (const t of r.textBlocks ?? []) yield t.text;
      }
    }
  };
  const walk = function* (el: BodyElement): Generator<string> {
    if ('runs' in el) yield* fromRuns((el as DocParagraph).runs);
    if ('rows' in el) {
      for (const row of (el as DocTable).rows) {
        for (const cell of row.cells) {
          for (const child of cell.content) yield* walk(child as BodyElement);
        }
      }
    }
  };
  const walkBody = function* (body: BodyElement[] | undefined): Generator<string> {
    for (const el of body ?? []) yield* walk(el);
  };

  yield* walkBody(doc.body);
  for (const hf of [doc.headers, doc.footers]) {
    for (const which of [hf?.default, hf?.first, hf?.even]) {
      yield* walkBody(which?.body);
    }
  }
  for (const note of [...(doc.footnotes ?? []), ...(doc.endnotes ?? [])]) {
    yield* walkBody(note.content);
  }
}

/**
 * The font-family names to preload for a document: the theme major/minor fonts,
 * plus only the script-fallback Noto faces whose script the document's TEXT
 * actually contains ({@link scriptPreloadNamesForText}). The renderer's font
 * fallback chains still END with the full Noto set, but eagerly fetching the
 * multi-MB CJK families for a document that has no CJK glyphs would block first
 * paint for nothing; an un-preloaded face loads lazily if it ever proves needed.
 *
 * Single source of truth shared by the main-thread `load()` and the render
 * worker. Both derive the set from the SAME parsed {@link DocxDocumentModel}, so
 * they preload an identical set — worker/main rendering must stay
 * pixel-equivalent. (Fonts must also be loaded before pagination, which measures
 * text; both callers await this before paginating.)
 */
export function docxFontPreloadNames(
  doc: DocxDocumentModel,
): (string | null | undefined)[] {
  const cjkLang =
    classifyCjkFont(doc.majorFont) ?? classifyCjkFont(doc.minorFont) ?? null;
  return [
    doc.majorFont,
    doc.minorFont,
    ...scriptPreloadNamesForText(docxTextRuns(doc), cjkLang),
  ];
}
