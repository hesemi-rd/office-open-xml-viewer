import { describe, it, expect } from 'vitest';
import { createLayoutServices, paginateDocument, renderDocumentToCanvas } from './renderer.js';
import { testFontSnapshot } from './layout/test-font-snapshot.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types';

// ECMA-376 §17.6.4 (newspaper columns) + the renderer's scale-independent
// pagination contract: paginateWithHeaderFooterReserve lays paragraphs out at
// scale 1 (pt space) so a page assignment is width-independent and cacheable
// across every render width (opts.prebuiltPages). Painting re-runs layoutLines
// at the actual render scale, and Canvas ctx.measureText is NOT perfectly
// scale-invariant (font hinting / sub-pixel glyph advances differ between the
// pt-size and the device-size run). A long, narrow paragraph can therefore wrap
// to FEWER lines when painted (larger scale) than the slice indices — computed
// at scale 1 — assume. The paginator stamps a continuation slice
// `lineSlice = { start, end }`; the paint pass must paint only the lines it
// actually has (lines.length) and never index `lines[end-1]` when that line is
// a phantom that existed only in the scale-1 measurement. Pre-fix the paint loop
// ran to `sliceEnd` unconditionally and threw
// "Cannot read properties of undefined (reading 'topY')" on `lines[i].topY`.
//
// This reproduces the sample-16 page-2 crash with a synthetic document — no
// dependency on the (gitignored) private sample. The non-linear measureText mock
// makes glyphs proportionally NARROWER at a larger font size, so the scale-2
// paint pass fits more characters per line and wraps the long paragraph to fewer
// lines than the scale-1 pagination — exactly the real font-hinting direction.

interface Call { text: string; x: number; y: number; }

/** Recording canvas whose glyph width is SUB-LINEAR in the font px size: each
 *  character is `px * (0.5 - SHRINK * px)` wide. At a larger render scale the
 *  per-glyph width grows slower than the box, so MORE characters fit per line
 *  and a long paragraph wraps to fewer lines than at scale 1 — the same
 *  paginate-vs-paint divergence real fonts produce through hinting. */
function makeNonLinearCanvas(): { canvas: HTMLCanvasElement; calls: Call[] } {
  const SHRINK = 0.002; // per-px narrowing; tuned so scale 1 vs 2 differ by lines
  let font = '10px serif';
  const calls: Call[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      const perChar = Math.max(0.05, p * (0.5 - SHRINK * p));
      return {
        width: [...s].length * perChar,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(s: string, x: number, y: number) { calls.push({ text: s, x, y }); },
    strokeText(s: string, x: number, y: number) { calls.push({ text: s, x, y }); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls };
}

(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
  getContext() { return makeNonLinearCanvas().canvas.getContext('2d'); }
};

function longPara(text: string): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'Times New Roman',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
    } as DocParagraph['runs'][number]],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

function doc(body: BodyElement[], pageHeight: number): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 200, pageHeight,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage', columns: null,
  } as SectionProps;
  return {
    section,
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
    footnotes: [],
  } as unknown as DocxDocumentModel;
}

async function paintedParagraphGeometry() {
  const paragraph = longPara(Array.from({ length: 180 }, () => 'w').join(' '));
  paragraph.spaceBefore = 6;
  paragraph.spaceAfter = 4;
  const model = doc([paragraph as unknown as BodyElement], 80);
  const services = createLayoutServices(model, { localMetrics: testFontSnapshot([{ family: 'Times New Roman' }]) });
  const pages = paginateDocument(model, services);
  const paintedPages: Array<{ lineCount: number; topYPx: number | null }> = [];
  for (let pageIndex = 0; pageIndex < pages.length; pageIndex++) {
    const { canvas, calls } = makeNonLinearCanvas();
    await renderDocumentToCanvas(model, canvas, pageIndex, {
      dpr: 1,
      width: 200,
      prebuiltPages: pages,
      layoutServices: services,
    });
    const textCalls = calls.filter((call) => call.text.includes('w'));
    paintedPages.push({
      lineCount: textCalls.length,
      topYPx: textCalls.length > 0 ? textCalls[0].y : null,
    });
  }
  return { pageCount: pages.length, paintedPages };
}

describe('paginate/paint line-count divergence — paint never indexes a phantom line (ECMA-376 §17.6.4)', () => {
  // A long paragraph of single-letter "words" (each followed by a space so the
  // line breaker has wrap opportunities). Narrow page + short page height force
  // it to wrap to many lines and split across multiple pages, so a later page
  // carries a continuation slice whose `end` is the paragraph's full scale-1
  // line count.
  const text = Array.from({ length: 400 }, () => 'w').join(' ');
  const body = (): BodyElement[] => [longPara(text) as unknown as BodyElement];

  it('preserves page count, painted line counts, and continuation top positions', async () => {
    const geometry = await paintedParagraphGeometry();

    expect(geometry).toEqual({
      pageCount: 2,
      paintedPages: [
        { lineCount: 84, topYPx: 24.74951171875 },
        { lineCount: 96, topYPx: 18.74951171875 },
      ],
    });
  });

  it('renders the final continuation slice without throwing when the paint scale wraps to fewer lines', async () => {
    // Render at width 400 over a 200pt page → scale 2. The non-linear mock fits
    // more characters per line at scale 2 than the scale-1 paginator did, so the
    // paint pass produces fewer lines and the paginator's final-slice `end`
    // overruns the paint `lines` array. The LAST page carries that slice.
    const pageHeight = 80; // short page → the paragraph spans several pages

    // Discover how many pages the paragraph spans by rendering page 0 first; then
    // render every page and assert none throws and content is painted on each.
    const pageCount = 8; // upper bound; out-of-range indices clamp to page 0 / []
    let totalLines = 0;
    let threw: unknown = null;
    for (let p = 0; p < pageCount; p++) {
      const { canvas, calls } = makeNonLinearCanvas();
      try {
        const model = doc(body(), pageHeight);
        await renderDocumentToCanvas(model, canvas, p, {
          dpr: 1, width: 400,
          layoutServices: createLayoutServices(model, {
            localMetrics: testFontSnapshot([{ family: 'Times New Roman' }]), measureContext: canvas.getContext('2d'),
          }),
        });
      } catch (e) {
        threw = e;
        break;
      }
      totalLines += calls.filter((c) => c.text.includes('w')).length;
    }

    // INVARIANT: no page throws. Pre-fix the final continuation slice indexed a
    // phantom line and threw the "reading 'topY'" TypeError.
    expect(threw).toBeNull();

    // NON-TRIVIALITY: the document actually painted content across pages (the
    // long paragraph really did wrap and split — otherwise the invariant above
    // would be vacuous).
    expect(totalLines).toBeGreaterThan(0);
  });
});
