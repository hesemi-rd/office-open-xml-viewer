// Minimal sfnt (OpenType) parser for math rendering. Reads only what layout needs:
// unitsPerEm (head), char->glyphId (cmap fmt 4/12), glyphId->advance (hmtx/hhea),
// global ascent/descent (hhea). Outline extraction and the MATH table live elsewhere.
// All distances are in font units; callers scale by fontSizePx / unitsPerEm.

export interface MathFont {
  unitsPerEm: number;
  /** font units, positive up */
  ascent: number;
  /** font units, positive magnitude (below baseline) */
  descent: number;
  /** Unicode code point -> glyph id (0 = .notdef / unmapped). */
  glyphForChar(cp: number): number;
  /** glyph id -> horizontal advance, font units. */
  advance(gid: number): number;
  /** sfnt table offset for `tag`, or -1 if absent. */
  tableOffset(tag: string): number;
  /** retained for MATH table parsing and future outline use. */
  buffer: ArrayBuffer;
}

function tableDirectory(dv: DataView): Map<string, number> {
  const num = dv.getUint16(4);
  const map = new Map<string, number>();
  let off = 12;
  for (let i = 0; i < num; i++) {
    const tag = String.fromCharCode(
      dv.getUint8(off),
      dv.getUint8(off + 1),
      dv.getUint8(off + 2),
      dv.getUint8(off + 3),
    );
    map.set(tag, dv.getUint32(off + 8));
    off += 16;
  }
  return map;
}

function readSubtable(dv: DataView, off: number): Map<number, number> {
  const fmt = dv.getUint16(off);
  const m = new Map<number, number>();
  if (fmt === 4) {
    const segX2 = dv.getUint16(off + 6);
    const segCount = segX2 / 2;
    const endO = off + 14;
    const startO = endO + segX2 + 2;
    const deltaO = startO + segX2;
    const rangeO = deltaO + segX2;
    for (let s = 0; s < segCount; s++) {
      const end = dv.getUint16(endO + s * 2);
      const start = dv.getUint16(startO + s * 2);
      const delta = dv.getUint16(deltaO + s * 2);
      const ro = dv.getUint16(rangeO + s * 2);
      for (let c = start; c <= end && c !== 0xffff; c++) {
        let g: number;
        if (ro === 0) {
          g = (c + delta) & 0xffff;
        } else {
          const gi = dv.getUint16(rangeO + s * 2 + ro + (c - start) * 2);
          g = gi === 0 ? 0 : (gi + delta) & 0xffff;
        }
        if (g) m.set(c, g);
      }
    }
  } else if (fmt === 12) {
    const nGroups = dv.getUint32(off + 12);
    let g = off + 16;
    for (let i = 0; i < nGroups; i++) {
      const sc = dv.getUint32(g);
      const ec = dv.getUint32(g + 4);
      const sg = dv.getUint32(g + 8);
      for (let c = sc; c <= ec; c++) m.set(c, sg + (c - sc));
      g += 12;
    }
  }
  return m;
}

function parseCmap(dv: DataView, base: number): Map<number, number> {
  const numTables = dv.getUint16(base + 2);
  let best = -1;
  for (let i = 0; i < numTables; i++) {
    const rec = base + 4 + i * 8;
    const plat = dv.getUint16(rec);
    const enc = dv.getUint16(rec + 2);
    const sub = base + dv.getUint32(rec + 4);
    // Prefer Unicode full-repertoire / BMP subtables.
    if ((plat === 3 && enc === 10) || (plat === 0 && enc >= 4)) {
      return readSubtable(dv, sub);
    }
    if ((plat === 3 && enc === 1) || plat === 0) best = sub;
  }
  return best >= 0 ? readSubtable(dv, best) : new Map();
}

/**
 * URL of the bundled default math font (Latin Modern Math, OFL). Resolved relative
 * to this module so bundlers (Vite/Rollup) fingerprint and serve the asset, and Node
 * gets a `file://` URL. Callers fetch this and pass the bytes to {@link parseMathFont}.
 */
export const defaultMathFontUrl: string = new URL(
  '../../assets/LatinModernMath.otf',
  import.meta.url,
).href;

/** Family name under which the default math font is registered for `ctx.fillText`. */
export const DEFAULT_MATH_FONT_FAMILY = 'LatinModernMath';

export function parseMathFont(buffer: ArrayBuffer): MathFont {
  const dv = new DataView(buffer);
  const dir = tableDirectory(dv);
  const head = dir.get('head');
  const hhea = dir.get('hhea');
  const hmtx = dir.get('hmtx');
  const maxp = dir.get('maxp');
  const cmapOff = dir.get('cmap');
  if (head == null || hhea == null || hmtx == null || maxp == null || cmapOff == null) {
    throw new Error('math font missing required sfnt tables');
  }
  const unitsPerEm = dv.getUint16(head + 18);
  const ascent = dv.getInt16(hhea + 4);
  const descent = -dv.getInt16(hhea + 6);
  const numHMetrics = dv.getUint16(hhea + 34);
  const numGlyphs = dv.getUint16(maxp + 4);
  const cmap = parseCmap(dv, cmapOff);

  const advance = (gid: number): number => {
    if (gid < 0 || gid >= numGlyphs) return 0;
    const i = gid < numHMetrics ? gid : numHMetrics - 1;
    return dv.getUint16(hmtx + i * 4);
  };

  return {
    unitsPerEm,
    ascent,
    descent,
    buffer,
    glyphForChar: (cp) => cmap.get(cp) ?? 0,
    advance,
    tableOffset: (tag) => dir.get(tag) ?? -1,
  };
}
