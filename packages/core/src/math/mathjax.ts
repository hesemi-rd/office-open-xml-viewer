// MathML → SVG via a pre-bundled MathJax v4 + STIX Two Math converter.
//
// The converter (engine + statically-baked STIX2 font, all glyph ranges) is
// built by `build/build-mathjax.mjs` into the opaque asset
// `assets/mathjax-stix2.js`. We load that asset *lazily* (only when a document
// actually contains equations): the `new URL(...)` lives inside a function so
// non-math viewers (xlsx) tree-shake it out, and keeping it pre-bundled stops
// the consuming app's bundler from re-bundling — and over-including — the
// MathJax source. The asset is self-contained: DOM-free internally, zero
// network, zero cross-origin requests. It exposes `globalThis.__ooxmlStix2`.

/* eslint-disable @typescript-eslint/no-explicit-any */

interface Stix2Engine {
  /** MathML string → standalone `<svg>…</svg>` (currentColor fill, viewBox in 1000-units/em). */
  mathml2svg(mathml: string): string;
}

let enginePromise: Promise<Stix2Engine> | null = null;

function resolveAssetUrl(): string {
  return new URL('../../assets/mathjax-stix2.js', import.meta.url).href;
}

function loadScript(src: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = src;
    s.async = true;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error(`Failed to load math engine from ${src}`));
    document.head.appendChild(s);
  });
}

function ensureEngine(): Promise<Stix2Engine> {
  if (enginePromise) return enginePromise;
  enginePromise = (async () => {
    const existing = (globalThis as any).__ooxmlStix2 as Stix2Engine | undefined;
    if (existing) return existing;
    if (typeof document === 'undefined') {
      throw new Error('Math rendering requires a DOM (browser environment)');
    }
    await loadScript(resolveAssetUrl());
    const engine = (globalThis as any).__ooxmlStix2 as Stix2Engine | undefined;
    if (!engine) throw new Error('Math engine failed to initialize');
    return engine;
  })();
  return enginePromise;
}

/** Preload the math engine. Call once before rendering equations. */
export async function loadMathJax(): Promise<void> {
  await ensureEngine();
}

export interface MathSvg {
  /** standalone `<svg>…</svg>` markup. */
  svg: string;
  /** extents in em (the SVG viewBox uses 1em = 1000 units). */
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

/** Convert a MathML string to a standalone SVG + its baseline-relative extents. */
export async function mathMLToSvg(mathml: string): Promise<MathSvg> {
  const engine = await ensureEngine();
  const svg = engine.mathml2svg(mathml);
  return { svg, ...svgExtents(svg) };
}

/** Replace MathJax's `currentColor` placeholders with an explicit color (for raster). */
export function recolorSvg(svg: string, color: string): string {
  return svg.replace(/currentColor/g, color);
}
