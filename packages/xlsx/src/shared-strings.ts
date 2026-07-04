import type { SharedString, Worksheet } from './types.js';

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
        cell.value = ss
          ? ss.runs !== undefined
            ? { type: 'text', text: ss.text, runs: ss.runs }
            : { type: 'text', text: ss.text }
          : { type: 'text', text: '' };
      }
    }
  }
  return ws;
}
