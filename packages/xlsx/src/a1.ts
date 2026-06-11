/**
 * Shared A1-style cell-reference parser (ECMA-376 §18.3.1.95 `ST_CellRef`).
 * Used by the renderer (comment-indicator placement), data-validation sqref
 * matching, and the comment hover popup. `$` absolute markers are stripped so
 * both `"H6"` and `"$H$6"` parse. Returns null on malformed input — parser-side
 * data is trusted, but callers still guard against junk.
 */
export function parseA1(ref: string): { row: number; col: number } | null {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(ref.trim());
  if (!m) return null;
  const letters = m[1];
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return { row: parseInt(m[2], 10), col };
}
