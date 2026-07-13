import {
  loadLocalFontMetrics,
  normalizeLocalFontMetricFamily,
  type LoadedLocalFontMetrics,
  type LocalFontMetricRequest,
} from '@silurus/ooxml-core';
import type { DocxDocumentModel } from './types.js';
import { docxRenderedFontFamilies } from './document-content.js';

/** Word's Far-East single-line height for Meiryo is 1.3 × the selected face's
 * hhea design box. This is Office compatibility behavior, not an ECMA-376
 * formula: it is measured from Word output and already underpins the static
 * fallback in core. Resolving the exact local face makes the rule version-safe
 * (Meiryo 6.30 has a different hhea box from older builds) without encoding a
 * version-specific number. Meiryo UI is deliberately excluded: it is a distinct
 * family with different metrics and has no equivalent Word measurement here. */
export function docxLocalMetricRequests(doc: DocxDocumentModel): LocalFontMetricRequest[] {
  // ECMA-376 §17.8.3.3: an embedded regular face is the document's authored
  // normal face for that family. Registering a terminal-local alias afterwards
  // would replace it during normal-run layout, so embedded families never enter
  // local probing. Bold/italic-only embeddings do not suppress a missing normal
  // face because they cannot satisfy the normal style slot.
  const embeddedRegularFamilies = new Set(
    (doc.embeddedFonts ?? [])
      .filter((font) => font.style === 'regular')
      .map((font) => normalizeLocalFontMetricFamily(font.fontName)),
  );
  const names = new Set<string>([
    ...Object.keys(doc.fontFamilyClasses ?? {}),
    ...(doc.majorFont ? [doc.majorFont] : []),
    ...(doc.minorFont ? [doc.minorFont] : []),
    ...docxRenderedFontFamilies(doc),
  ]);
  const requests: LocalFontMetricRequest[] = [];
  for (const family of names) {
    const normalized = normalizeLocalFontMetricFamily(family);
    if (embeddedRegularFamilies.has(normalized)) continue;
    const isMeiryo = normalized === 'meiryo' || family.trim() === 'メイリオ';
    if (!isMeiryo) continue;
    requests.push({
      family,
      localNames: ['Meiryo'],
      lineHeightMultiplier: 1.3,
    });
  }
  return requests;
}

export function loadDocxLocalFontMetrics(
  doc: DocxDocumentModel,
): Promise<LoadedLocalFontMetrics> {
  return loadLocalFontMetrics(docxLocalMetricRequests(doc));
}
