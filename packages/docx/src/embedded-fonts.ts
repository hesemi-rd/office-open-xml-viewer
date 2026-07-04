import { registerEmbeddedFonts, type EmbeddedFontFace } from '@silurus/ooxml-core';
import type { DocxDocumentModel, EmbeddedFontRef } from './types';

/**
 * Register a document's embedded fonts (ECMA-376 §17.8.3.3-.6) into the active
 * FontFaceSet so the renderer measures and paints text with the authored
 * typeface instead of a substitute.
 *
 * `doc.embeddedFonts` names the obfuscated `.odttf` parts + their `w:fontKey`
 * GUIDs; the bytes are fetched by zip path through `fetchFontBytes` (the docx
 * archive extracts any part, not just images), de-obfuscated per §17.8.1 by
 * {@link registerEmbeddedFonts}, and added to the set under the exact document
 * font name. Each `<w:embed*>` style slot becomes one CSS weight/style pair:
 * bold / boldItalic ⇒ `weight: 'bold'`; italic / boldItalic ⇒ `style: 'italic'`.
 *
 * MUST run before pagination (which measures text). No-ops when the document
 * embeds no fonts. Individual part fetches are concurrent; a rejected fetch
 * skips only that face (the rest still register) so one missing part never
 * aborts the whole document.
 */
export async function loadEmbeddedFonts(
  doc: DocxDocumentModel,
  fetchFontBytes: (partPath: string) => Promise<Uint8Array>,
): Promise<void> {
  const refs = doc.embeddedFonts;
  if (!refs || refs.length === 0) return;

  const faces = await Promise.all(
    refs.map(async (ref): Promise<EmbeddedFontFace | null> => {
      try {
        const bytes = await fetchFontBytes(ref.partPath);
        return {
          family: ref.fontName,
          bytes,
          odttf: ref.partPath.toLowerCase().endsWith('.odttf'),
          fontKey: ref.fontKey,
          weight: weightForStyle(ref.style),
          style: styleForStyle(ref.style),
        };
      } catch {
        // A missing / unreadable part: skip this face, keep the rest.
        return null;
      }
    }),
  );

  const loadable = faces.filter((f): f is EmbeddedFontFace => f !== null);
  if (loadable.length === 0) return;
  await registerEmbeddedFonts(loadable);
}

/** bold / boldItalic slots ⇒ CSS `font-weight: bold`; otherwise `normal`. */
function weightForStyle(style: EmbeddedFontRef['style']): 'normal' | 'bold' {
  return style === 'bold' || style === 'boldItalic' ? 'bold' : 'normal';
}

/** italic / boldItalic slots ⇒ CSS `font-style: italic`; otherwise `normal`. */
function styleForStyle(style: EmbeddedFontRef['style']): 'normal' | 'italic' {
  return style === 'italic' || style === 'boldItalic' ? 'italic' : 'normal';
}
