// Entry for the pre-bundled MathJax v4 + STIX Two Math converter.
//
// esbuild bundles this (and only this) into `assets/mathjax-stix2.js`, an
// opaque, tree-shaken, minified IIFE (~3 MB). The renderer loads that asset
// lazily, so it never bloats non-math viewers and is never re-bundled by the
// consuming app's bundler. The STIX2 font is baked in statically: the
// math-relevant glyph ranges below are imported up-front so `dynamicSetup`
// marks them loaded → DOM-free, zero network, zero cross-origin (no on-demand
// range fetches). Non-math ranges (cyrillic, phonetics, dingbats, accented
// Latin variants) are intentionally omitted to keep the bundle small.
import { mathjax } from '@mathjax/src/mjs/mathjax.js';
import { MathML } from '@mathjax/src/mjs/input/mathml.js';
import { SVG } from '@mathjax/src/mjs/output/svg.js';
import { liteAdaptor } from '@mathjax/src/mjs/adaptors/liteAdaptor.js';
import { RegisterHTMLHandler } from '@mathjax/src/mjs/handlers/html.js';
import { MathJaxStix2Font } from '@mathjax/mathjax-stix2-font/mjs/svg.js';
import '@mathjax/mathjax-stix2-font/mjs/svg/dynamic/accents-other.js';
import '@mathjax/mathjax-stix2-font/mjs/svg/dynamic/accents.js';
import '@mathjax/mathjax-stix2-font/mjs/svg/dynamic/arrows.js';
import '@mathjax/mathjax-stix2-font/mjs/svg/dynamic/calligraphic.js';
import '@mathjax/mathjax-stix2-font/mjs/svg/dynamic/double-struck.js';
import '@mathjax/mathjax-stix2-font/mjs/svg/dynamic/enclosed.js';
import '@mathjax/mathjax-stix2-font/mjs/svg/dynamic/fraktur.js';
import '@mathjax/mathjax-stix2-font/mjs/svg/dynamic/greek.js';
import '@mathjax/mathjax-stix2-font/mjs/svg/dynamic/math.js';
import '@mathjax/mathjax-stix2-font/mjs/svg/dynamic/monospace.js';
import '@mathjax/mathjax-stix2-font/mjs/svg/dynamic/sans-serif.js';
import '@mathjax/mathjax-stix2-font/mjs/svg/dynamic/script.js';
import '@mathjax/mathjax-stix2-font/mjs/svg/dynamic/shapes.js';
import '@mathjax/mathjax-stix2-font/mjs/svg/dynamic/stretchy.js';
import '@mathjax/mathjax-stix2-font/mjs/svg/dynamic/symbols-other.js';
import '@mathjax/mathjax-stix2-font/mjs/svg/dynamic/symbols.js';
import '@mathjax/mathjax-stix2-font/mjs/svg/dynamic/upright.js';
import '@mathjax/mathjax-stix2-font/mjs/svg/dynamic/variants.js';

const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);
// `linebreaks.inline:false` + `displayOverflow:'overflow'`: never auto-break an
// equation into multiple sibling <svg>s (MathJax v4 separates them with
// <mjx-break>, which our single-<svg> extraction / <img> rasterization can't
// consume — a long equation would otherwise silently fail to load).
const svgJax = new SVG({
  fontData: MathJaxStix2Font,
  fontCache: 'none',
  displayOverflow: 'overflow',
  linebreaks: { inline: false },
});
const doc = mathjax.document('', { InputJax: new MathML(), OutputJax: svgJax });

// Force-load every dynamic glyph range into the font instance up-front.
//
// `dynamicSetup` (run by the range imports above) only stores a `setup(font)`
// closure on each `dynamicFiles` entry; the glyphs aren't defined until that
// closure runs (normally triggered by an on-demand, async, cross-origin fetch).
// Since the ranges are already bundled, we run the closures synchronously here
// against the output's font instance. This makes EVERY styled glyph (script,
// fraktur, double-struck, calligraphic, sans-serif, stretchy bars, …) render
// as a `<path>` immediately — no `<text>` fallback (wrong font/metrics), no
// "retry" exception, and no network request.
const dynamicFiles = svgJax.font.constructor.dynamicFiles;
for (const name of Object.keys(dynamicFiles)) {
  try {
    dynamicFiles[name].setup(svgJax.font);
  } catch {
    /* a range with no real setup (shouldn't happen — all are imported) */
  }
}

// Exposed on globalThis so the lazily-injected script hands the API back to the
// renderer. Returns a standalone `<svg>…</svg>` string (currentColor fill, a
// `0 -minY w h` viewBox in 1000-units/em).
globalThis.__ooxmlStix2 = {
  mathml2svg(mathml) {
    const node = doc.convert(mathml, { display: true });
    const html = adaptor.outerHTML(node);
    // Use the FIRST <svg and the LAST </svg>: stretchy glyphs (overlines, norm
    // bars, big operators) emit a NESTED <svg>, so matching the first </svg>
    // would truncate the markup into a broken (img-unloadable) fragment.
    const s = html.indexOf('<svg');
    const e = html.lastIndexOf('</svg>');
    return s >= 0 && e >= 0 ? html.slice(s, e + 6) : html;
  },
};
