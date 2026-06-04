// MathJax turns MathML into an SVG string that the consumer rasterizes onto the canvas.
//
// The MathJax v3 `mml-svg` component is vendored in this package (../../assets/mathjax),
// is self-contained (glyph outlines inlined, ZERO runtime network), and is loaded lazily
// only when a document actually contains equations. Browser-only (it needs a DOM). The
// URL is overridable via `setMathJaxUrl` for self-hosting / a pinned CDN copy.

/* eslint-disable @typescript-eslint/no-explicit-any */

// The MathJax v3 component bundled with this package (Apache-2.0) is resolved to a
// same-origin asset URL by the consumer's bundler (Vite/webpack/rollup all handle
// `new URL(asset, import.meta.url)`). This makes math a built-in feature with no
// cross-origin request. `setMathJaxUrl` overrides it (e.g. a CDN or pinned copy).
//
// IMPORTANT: the `new URL(...)` lives inside `resolveMathJaxUrl()` rather than at module
// scope so that bundles which never call into the math pipeline (the pptx/xlsx viewers)
// tree-shake the whole module out and don't inline the ~2.5MB MathJax asset.
let mathJaxUrlOverride: string | null = null;
let mjPromise: Promise<any> | null = null;

function resolveMathJaxUrl(): string {
  if (mathJaxUrlOverride) return mathJaxUrlOverride;
  return new URL('../../assets/mathjax/mml-svg.js', import.meta.url).href;
}

/** Override the MathJax script URL (e.g. a CDN or self-hosted copy). Call before load. */
export function setMathJaxUrl(url: string): void {
  mathJaxUrlOverride = url;
}

function loadScript(src: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = src;
    s.async = true;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error(`Failed to load MathJax from ${src}`));
    document.head.appendChild(s);
  });
}

async function ensureMathJax(): Promise<any> {
  if (mjPromise) return mjPromise;
  mjPromise = (async () => {
    const w = globalThis as any;
    if (w.MathJax?.startup?.promise) {
      await w.MathJax.startup.promise;
      return w.MathJax;
    }
    if (typeof document === 'undefined') {
      throw new Error('MathJax rendering requires a DOM (browser environment)');
    }
    // Configure before the component script runs: don't auto-typeset the page, and
    // inline glyph outlines (fontCache:'none') so each SVG is self-contained.
    w.MathJax = {
      ...(w.MathJax || {}),
      startup: { ...(w.MathJax?.startup || {}), typeset: false },
      svg: { ...(w.MathJax?.svg || {}), fontCache: 'none' },
      // No a11y menu / speech-rule worker — keeps it fully offline (no extra fetches).
      options: { ...(w.MathJax?.options || {}), enableMenu: false },
    };
    await loadScript(resolveMathJaxUrl());
    await w.MathJax.startup.promise;
    return w.MathJax;
  })();
  return mjPromise;
}

/** Preload MathJax. Call once before rendering equations. */
export async function loadMathJax(): Promise<void> {
  await ensureMathJax();
}

export interface MathSvg {
  /** standalone `<svg>…</svg>` markup. */
  svg: string;
  /** extents in em (MathJax SVG uses 1em = 1000 viewBox units). */
  widthEm: number;
  ascentEm: number;
  descentEm: number;
}

const UNITS_PER_EM = 1000;

/** Parse the MathJax SVG viewBox into baseline-relative em extents. */
export function svgExtents(svg: string): { widthEm: number; ascentEm: number; descentEm: number } {
  const m = /viewBox="([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+)"/.exec(svg);
  if (!m) return { widthEm: 0, ascentEm: 0, descentEm: 0 };
  const minY = parseFloat(m[2]);
  const w = parseFloat(m[3]);
  const h = parseFloat(m[4]);
  // The output's top <g> applies scale(1,-1): content rises to -minY above the
  // baseline and falls to (minY + h) below it.
  return {
    widthEm: w / UNITS_PER_EM,
    ascentEm: -minY / UNITS_PER_EM,
    descentEm: (minY + h) / UNITS_PER_EM,
  };
}

/** Convert a MathML string to an SVG + its baseline-relative extents. */
export async function mathMLToSvg(mathml: string): Promise<MathSvg> {
  const MathJax = await ensureMathJax();
  const container = MathJax.mathml2svg(mathml, { display: true });
  const svgEl: Element | null =
    typeof container.querySelector === 'function' ? container.querySelector('svg') : null;
  const svg: string = svgEl ? svgEl.outerHTML : MathJax.startup.adaptor.outerHTML(container);
  return { svg, ...svgExtents(svg) };
}

/** Replace MathJax's `currentColor` placeholders with an explicit color (for raster). */
export function recolorSvg(svg: string, color: string): string {
  return svg.replace(/currentColor/g, color);
}
