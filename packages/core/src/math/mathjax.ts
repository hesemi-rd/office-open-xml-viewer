// MathJax (DOM-free, via liteAdaptor) turns MathML into an SVG string. The SVG is
// rasterized to the canvas by the consumer. MathJax is loaded lazily so it only ships
// to pages that actually render equations.

/* eslint-disable @typescript-eslint/no-explicit-any */
interface MathJaxInstance {
  adaptor: any;
  doc: any;
}

let mjPromise: Promise<MathJaxInstance> | null = null;

async function getMathJax(): Promise<MathJaxInstance> {
  if (!mjPromise) {
    mjPromise = (async () => {
      const [{ mathjax }, { MathML }, { SVG }, { liteAdaptor }, { RegisterHTMLHandler }] =
        await Promise.all([
          import('mathjax-full/js/mathjax.js'),
          import('mathjax-full/js/input/mathml.js'),
          import('mathjax-full/js/output/svg.js'),
          import('mathjax-full/js/adaptors/liteAdaptor.js'),
          import('mathjax-full/js/handlers/html.js'),
        ]);
      const adaptor = liteAdaptor();
      RegisterHTMLHandler(adaptor);
      // fontCache: 'none' inlines glyph outlines as <path> (no <use>/<defs>), which
      // keeps each SVG self-contained and rasterizable as a standalone image.
      const doc = mathjax.document('', {
        InputJax: new MathML(),
        OutputJax: new SVG({ fontCache: 'none' }),
      });
      return { adaptor, doc };
    })();
  }
  return mjPromise;
}

/** Preload MathJax. Call once before rendering equations. */
export async function loadMathJax(): Promise<void> {
  await getMathJax();
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
  const { adaptor, doc } = await getMathJax();
  const container = doc.convert(mathml, { display: true });
  const svgNode = adaptor.firstChild(container);
  const svg: string = adaptor.outerHTML(svgNode);
  return { svg, ...svgExtents(svg) };
}

/** Replace MathJax's `currentColor` placeholders with an explicit color (for raster). */
export function recolorSvg(svg: string, color: string): string {
  return svg.replace(/currentColor/g, color);
}
