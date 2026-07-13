import {
  loadLocalFontMetrics,
  type LoadedLocalFontMetrics,
  type LocalFontMetricRequest,
} from '@silurus/ooxml-core';
import type { DocxDocumentModel } from './types.js';

/** Word's Far-East single-line height for Meiryo is 1.3 × the selected face's
 * hhea design box. This is Office compatibility behavior, not an ECMA-376
 * formula: it is measured from Word output and already underpins the static
 * fallback in core. Resolving the exact local face makes the rule version-safe
 * (Meiryo 6.30 has a different hhea box from older builds) without encoding a
 * version-specific number. Meiryo UI is deliberately excluded: it is a distinct
 * family with different metrics and has no equivalent Word measurement here. */
export function docxLocalMetricRequests(doc: DocxDocumentModel): LocalFontMetricRequest[] {
  const names = new Set<string>([
    ...Object.keys(doc.fontFamilyClasses ?? {}),
    ...(doc.majorFont ? [doc.majorFont] : []),
    ...(doc.minorFont ? [doc.minorFont] : []),
  ]);
  const requests: LocalFontMetricRequest[] = [];
  for (const family of names) {
    const normalized = family.trim().toLowerCase();
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
