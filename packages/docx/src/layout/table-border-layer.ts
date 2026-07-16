import type { TableBorderInput } from './types.js';

/** Resolve one authored border layer before shared-edge conflict resolution.
 * `none` is an omitted layer while `nil` is an authored suppression. Keeping
 * that distinction here prevents callers that merge tblPrEx/table inputs from
 * accidentally turning an OOXML suppression into a fallback. */
export function firstAuthoredTableBorder(
  ...borders: readonly (TableBorderInput | null)[]
): TableBorderInput | null {
  for (const border of borders) {
    // [MS-OI29500] 2.1.169: nil remains specified and suppresses the edge;
    // none allows the next applicable border layer to participate.
    if (border && border.authoredStyle !== 'none') return border;
  }
  return null;
}
