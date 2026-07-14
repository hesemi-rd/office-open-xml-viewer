import {
  loadLocalFontMetrics,
  normalizeLocalFontMetricFamily,
  type LoadedLocalFontMetrics,
  type LocalFontMetricRequest,
} from '@silurus/ooxml-core';
import type { DocxDocumentModel } from './types.js';
import { docxRenderedTextUsages } from './document-content.js';

/** Word's Far-East single-line height for Meiryo is 1.3 × the selected face's
 * hhea design box. This is Office compatibility behavior, not an ECMA-376
 * formula: it is measured from Word output and already underpins the static
 * fallback in core. Resolving the exact local face makes the rule version-safe
 * (Meiryo 6.30 has a different hhea box from older builds) without encoding a
 * version-specific number. Meiryo UI is deliberately excluded: it is a distinct
 * family with different metrics and has no equivalent Word measurement here. */
export function docxLocalMetricRequests(
  doc: DocxDocumentModel,
  includeGeometryText = false,
): LocalFontMetricRequest[] {
  // ECMA-376 §17.8.3.3: an embedded face is authoritative for its exact
  // family/weight/style tuple. Other tuples may still use a positively probed
  // local regular face through the core loader's isolated synthesis alias.
  const embeddedTuples = new Set(
    (doc.embeddedFonts ?? [])
      .map((font) => {
        const weight = font.style === 'bold' || font.style === 'boldItalic' ? 700 : 400;
        const style = font.style === 'italic' || font.style === 'boldItalic' ? 'italic' : 'normal';
        return `${normalizeLocalFontMetricFamily(font.fontName)}:${weight}:${style}`;
      }),
  );
  const candidates: Array<{
    request: LocalFontMetricRequest;
    geometryTexts: Set<string>;
  }> = [];
  const byTuple = new Map<string, (typeof candidates)[number]>();
  const addCandidate = (
    familyValue: string | null | undefined,
    weight: number,
    style: 'normal' | 'italic',
    geometryText: string,
  ): void => {
    const family = familyValue?.trim();
    if (!family) return;
    const normalized = normalizeLocalFontMetricFamily(family);
    const key = `${normalized}:${weight}:${style}`;
    if (embeddedTuples.has(key)) return;
    const existing = byTuple.get(key);
    if (existing) {
      if (geometryText) existing.geometryTexts.add(geometryText);
      return;
    }
    const isMeiryo = normalized === 'meiryo' || family === 'メイリオ';
    const candidate = {
      request: {
        family,
        localNames: [isMeiryo ? 'Meiryo' : family],
        ...(isMeiryo ? { lineHeightMultiplier: 1.3 } : {}),
        ...(weight === 400 ? {} : { weight }),
        ...(style === 'normal' ? {} : { style }),
      },
      geometryTexts: new Set(geometryText ? [geometryText] : []),
    };
    candidates.push(candidate);
    byTuple.set(key, candidate);
  };
  for (const usage of docxRenderedTextUsages(doc)) {
    const weight = usage.bold ? 700 : 400;
    const style = usage.italic ? 'italic' as const : 'normal' as const;
    for (const authoredFamily of usage.fontFamilies) {
      addCandidate(authoredFamily, weight, style, usage.text);
    }
  }
  return candidates.map(({ request, geometryTexts }) => ({
    ...request,
    ...(includeGeometryText && geometryTexts.size > 0
      ? { geometryProbeTexts: [...geometryTexts] }
      : {}),
  }));
}

export function loadDocxLocalFontMetrics(
  doc: DocxDocumentModel,
): Promise<LoadedLocalFontMetrics> {
  return loadLocalFontMetrics(docxLocalMetricRequests(doc, true));
}
