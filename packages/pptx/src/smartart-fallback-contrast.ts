// Contrast-aware default text colour for the synthetic SmartArt fallback
// (issue #805).
//
// When a pptx stores no prebaked SmartArt drawing part, the Rust parser
// (`smartart_fallback.rs`) synthesizes a bulleted-list shape from the diagram
// data model (`ppt/diagrams/dataN.xml`). That data model rarely carries run
// colours — the real colours live in `colors*.xml` / `quickStyle*.xml`, which
// the fallback deliberately discards — so the runs resolve to `color: null`
// and fall to the renderer's theme default (dk1, dark). On a dark slide
// background the whole list is invisible (dark-on-dark).
//
// This module derives a legible default for THAT SYNTHETIC SHAPE ONLY, from
// the slide-background luminance: light background → keep the dark theme
// default (unchanged output), dark background → white. This is a UX choice
// for an already-synthetic representation — ECMA-376 defines no text colour
// for a diagram whose drawing part is absent — closest in spirit to
// PowerPoint's own dk1/lt1 text autocolour. Explicit run colours (and a
// style-derived `defaultTextColor`) always take precedence and are never
// touched.

import { hex6ToRgb, luminance601 } from '@silurus/ooxml-core';
import type { Fill } from '@silurus/ooxml-core';

/**
 * Rec. 601 luma (0–1) of a 6- or 8-char hex colour (no leading `#`). An
 * 8-char `RRGGBBAA` value is composited over white first — the canvas base
 * every render paints behind a translucent background — so a mostly
 * transparent dark colour reads as the light surface it actually shows as.
 * Returns null for malformed input.
 */
function hexLuma(hex: string): number | null {
  const rgb = hex6ToRgb(hex.length === 8 ? hex.slice(0, 6) : hex);
  if (!rgb) return null;
  const base = luminance601(rgb[0], rgb[1], rgb[2]);
  if (hex.length !== 8) return base;
  const a = Number.parseInt(hex.slice(6, 8), 16);
  if (Number.isNaN(a)) return null;
  const alpha = a / 255;
  return alpha * base + (1 - alpha); // over the white base (luma 1)
}

/**
 * Overall luminance (0–1) of a slide background fill, or null when it cannot
 * be derived statically:
 *
 * - solid → the colour's luma;
 * - gradient → the piecewise-linear integral of stop luma over [0,1] (the
 *   average brightness of the ramp; end stops extend to the edges exactly as
 *   canvas gradients clamp them);
 * - image / pattern / none / absent → null. An image's brightness would need
 *   a decode, a preset pattern's fg/bg mix depends on the preset bitmap, and
 *   a missing background shows the host page. Unknown means "change nothing".
 */
export function backgroundLuminance(fill: Fill | null): number | null {
  if (!fill) return null;
  if (fill.fillType === 'solid') return hexLuma(fill.color);
  if (fill.fillType === 'gradient') {
    const stops = fill.stops
      .map((s) => ({ p: Math.min(1, Math.max(0, s.position)), l: hexLuma(s.color) }))
      .filter((s): s is { p: number; l: number } => s.l !== null)
      .sort((a, b) => a.p - b.p);
    if (stops.length === 0) return null;
    const first = stops[0];
    const last = stops[stops.length - 1];
    let area = first.l * first.p + last.l * (1 - last.p);
    for (let i = 0; i + 1 < stops.length; i++) {
      area += ((stops[i].l + stops[i + 1].l) / 2) * (stops[i + 1].p - stops[i].p);
    }
    return area;
  }
  return null;
}

/**
 * True for the shape the Rust SmartArt fallback synthesizes
 * (`smartart_fallback.rs` `text_list_shape`): it is the only shape the parser
 * emits with `name: "SmartArt"` and NO `id` — a real shape's `<p:cNvPr>`
 * carries a schema-required `id` attribute, so a file-authored shape that
 * happens to be named "SmartArt" still has one.
 */
export function isSmartArtFallbackShape(el: { name?: string; id?: string }): boolean {
  return el.name === 'SmartArt' && el.id === undefined;
}

/**
 * The default text colour the SmartArt fallback shape's null-colour runs
 * should use on this slide, or null to keep the ordinary theme default.
 *
 * The 0.5 threshold is the Rec. 601 luma midpoint: the crossover at which
 * black and white text are equidistant from the background — below it, white
 * text has the larger luma separation. A theme default that is itself light
 * (≥ 0.5, e.g. a dark-designed theme whose dk1 maps light) is already legible
 * on a dark background and is kept, so this only ever REPLACES an illegible
 * dark-on-dark default, never a working one.
 */
export function smartArtFallbackTextColor(
  background: Fill | null,
  themeDefaultColor: string,
): string | null {
  const bg = backgroundLuminance(background);
  if (bg === null || bg >= 0.5) return null;
  const themeLuma = hexLuma(themeDefaultColor.replace(/^#/, ''));
  if (themeLuma !== null && themeLuma >= 0.5) return null;
  // lt1's spec default and the maximum-contrast choice on a dark background.
  return '#FFFFFF';
}
