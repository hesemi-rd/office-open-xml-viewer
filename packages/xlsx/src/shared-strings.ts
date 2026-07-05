import type { CellValue, SharedString, Worksheet } from './types.js';

/**
 * Resolve every `{ type: 'shared', si }` cell in `ws` to a concrete
 * `{ type: 'text', text, runs? }` by looking `si` up in the workbook
 * `sharedStrings` table (ECMA-376 §18.4.8). Mutates cells in place and returns
 * `ws` for chaining. Out-of-range / missing `si` resolves to empty text —
 * matching the parser's historical fallback. Idempotent: a `Worksheet` with no
 * `shared` cells is returned unchanged.
 *
 * This keeps the dedup win on the wire (each shared string ships ONCE in the
 * workbook) while every downstream consumer — renderer, formula engine, number
 * formatter, markdown — still sees fully-resolved cell text.
 */
export function resolveSharedStrings(ws: Worksheet, sharedStrings: SharedString[]): Worksheet {
  for (const row of ws.rows) {
    for (const cell of row.cells) {
      const v = cell.value;
      if (v.type === 'shared') {
        const ss = sharedStrings[v.si];
        if (ss) {
          // Carry the String Item's furigana (§18.4.6 rPh / §18.4.3
          // phoneticPr) onto the resolved text cell so the renderer can draw
          // the phonetic band. Only meaningful when the CELL opted in with
          // `ph="1"` (checked at draw time), but the reading always rides with
          // the resolved value.
          const resolved: CellValue = { type: 'text', text: ss.text };
          if (ss.runs !== undefined) resolved.runs = ss.runs;
          if (ss.phoneticRuns !== undefined) resolved.phoneticRuns = ss.phoneticRuns;
          if (ss.phoneticPr !== undefined) resolved.phoneticPr = ss.phoneticPr;
          cell.value = resolved;
        } else {
          cell.value = { type: 'text', text: '' };
        }
      }
    }
  }
  return ws;
}
