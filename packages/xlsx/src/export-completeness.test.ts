import { describe, it, expect } from 'vitest';
import { findMissingExportsFromUrl, formatMissing } from '@silurus/ooxml-core/testing';

/**
 * v1.0 API freeze guard: every public type reachable from the xlsx barrel must
 * itself be exported. See `@silurus/ooxml-core/testing` for the algorithm.
 */
describe('xlsx public API export completeness', () => {
  it('exports every in-package type reachable from index.ts', () => {
    const missing = findMissingExportsFromUrl(import.meta.url);
    expect(missing, formatMissing(missing)).toEqual([]);
  });
});
