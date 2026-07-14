import type { ResolvedLocalFontMetric } from '@silurus/ooxml-core';
import { docxRenderedFontFamilies } from '../document-content.js';
import type { DocxDocumentModel } from '../types.js';

/** Explicit deterministic authored-face inventory for renderer unit fixtures.
 * Real document opens populate this map only through positive local()/embedded/
 * web-font loading. Tests use synthetic Canvas contexts, so they must state the
 * faces those contexts model instead of relying on host-local resolution. */
export function testFontSnapshot(
  doc: DocxDocumentModel,
): Record<string, ResolvedLocalFontMetric> {
  const metrics: Record<string, ResolvedLocalFontMetric> = {};
  const families = new Set([
    ...Object.keys(doc.fontFamilyClasses ?? {}),
    ...docxRenderedFontFamilies(doc),
  ]);
  for (const family of families) {
    const normalized = family.trim().toLowerCase();
    if (!normalized) continue;
    for (const [weight, style] of [
      [400, 'normal'], [700, 'normal'], [400, 'italic'], [700, 'italic'],
    ] as const) {
      const key = weight === 400 && style === 'normal'
        ? normalized
        : `${normalized}:${weight}:${style}`;
      metrics[key] = {
        family,
        requestedFamily: family,
        weight,
        style,
        geometrySignature: 'synthetic-test-context',
      };
    }
  }
  return metrics;
}
