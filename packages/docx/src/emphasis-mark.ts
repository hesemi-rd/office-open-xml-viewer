// ECMA-376 §17.3.2.12 `<w:em>` / §17.18.24 ST_Em — emphasis marks (圏点 / boten).
//
// A run may carry a single emphasis-mark style that is stamped on EVERY
// non-space character of the run (§17.18.24: "applied to each non-space
// character in a run"). In horizontal writing the mark sits centred above the
// character (below for `underDot`). The glyph used is implementation-dependent
// (§17.18.24); we follow the JIS X 4051 圏点 convention and Word's rendering:
//
//   dot      → a small filled disc centred above the glyph
//   circle   → a small hollow (stroked) circle above the glyph
//   comma    → a filled "sesame" teardrop (bōten «﹅»-like) above the glyph
//   underDot → a small filled disc centred BELOW the glyph
//
// The marks are drawn per glyph AFTER the run's text and never change the glyph
// advance (they are an overlay, exactly like the ruby / highlight decorations),
// so layout metrics are untouched. The one metric consequence — a tall mark can
// reach into the previous line's descent region when the line box is tight — is
// deliberately NOT compensated here: Word widens line spacing for emphasised
// text, but the exact reservation is not specified, so we render the mark in the
// existing line box and accept possible overlap on very tight leading. See the
// module test + the renderer call site for the rationale.

import type { EmphasisMark } from './types';

/** One emphasis mark to paint, positioned by its horizontal CENTRE (device px,
 *  absolute — i.e. already offset by the segment pen `x`). */
export interface EmphasisMarkPlacement {
  /** Absolute device-x of the mark centre. */
  centerX: number;
}

/**
 * Compute the centre-x of the emphasis mark for every NON-SPACE code point of
 * `text`, in the same left-to-right pen order the glyphs are drawn.
 *
 * The centre of glyph *i* is the midpoint of its drawn advance:
 *   left_i  = penX + measure(prefix_i)          + i·pitch
 *   right_i = penX + measure(prefix_{i+1})       + (i+1)·pitch
 *   centre  = (left_i + right_i) / 2
 *
 * `measure(str)` must return the contextual advance width of `str` in the
 * current font (the caller passes `ctx.measureText(str).width`), so the mark
 * tracks the browser's contextual CJK metrics (約物半角 half-width collapse etc.)
 * exactly as the glyph draw does. `pitch` is the uniform per-glyph-boundary
 * spacing the run was drawn with (docGrid cell delta or justification distribute
 * pitch, 0 for the common path); it is added between glyphs so the mark stays
 * centred under justification/grid stretch. Spaces (and other whitespace) get no
 * mark (§17.18.24) but still advance the cumulative measure.
 */
export function emphasisMarkCenters(
  text: string,
  measure: (s: string) => number,
  penX: number,
  pitch: number,
): EmphasisMarkPlacement[] {
  const cps = [...text]; // code points — surrogate pairs stay intact
  const out: EmphasisMarkPlacement[] = [];
  let prefix = '';
  for (let i = 0; i < cps.length; i++) {
    const cp = cps[i];
    const left = penX + measure(prefix) + i * pitch;
    const nextPrefix = prefix + cp;
    const right = penX + measure(nextPrefix) + (i + 1) * pitch;
    prefix = nextPrefix;
    // §17.18.24: no mark on space characters. Test the whole code point so a
    // combining space or an ideographic space (U+3000) is also skipped.
    if (!/\s/u.test(cp)) {
      out.push({ centerX: (left + right) / 2 });
    }
  }
  return out;
}

/** Geometry of an emphasis mark, in device px, relative to the glyph box.
 *  `radius` is the mark's radius; `above` = draw centred above the glyph top
 *  (false ⇒ below the glyph bottom, for `underDot`). */
export interface EmphasisMarkGeometry {
  /** 'dot' | 'comma' | 'circle' — the shape to stamp. */
  shape: 'dot' | 'comma' | 'circle';
  /** Mark radius (device px). */
  radius: number;
  /** True ⇒ mark sits above the glyphs; false ⇒ below (underDot). */
  above: boolean;
}

/**
 * Resolve an {@link EmphasisMark} value to its drawing geometry for a run of the
 * given effective font size (device px).
 *
 * The mark radius is ~7% of the em (⌀ ≈ 0.14 em). JIS X 4051 圏点 are drawn at
 * roughly a sixth of the character body; a disc of ⌀ 0.14 em reads as a clear
 * boten without crowding the line above. `underDot` is the only mark placed
 * below the glyphs (§17.18.24). `comma` reuses the disc radius but is drawn as a
 * teardrop by the renderer.
 */
export function emphasisMarkGeometry(
  mark: EmphasisMark,
  effSizePx: number,
): EmphasisMarkGeometry {
  const radius = effSizePx * 0.07;
  switch (mark) {
    case 'circle':
      return { shape: 'circle', radius, above: true };
    case 'comma':
      return { shape: 'comma', radius, above: true };
    case 'underDot':
      return { shape: 'dot', radius, above: false };
    case 'dot':
    default:
      return { shape: 'dot', radius, above: true };
  }
}
