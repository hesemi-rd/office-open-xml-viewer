import { describe, it, expect, beforeAll } from 'vitest';
import { layoutDocument } from './document-layout.js';
import { __test_setFitMeasureReuseEnabled } from './renderer.js';
import type { ParagraphFragment } from './layout-fragments.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types';

// ─────────────────────────────────────────────────────────────────────────────
// M-1 — avoid the double measurement per non-split body paragraph.
//
// The paginator measures each body paragraph once for the fit decision
// (measureBodyParagraphAtCursor) and, before M-1, measured it AGAIN at its final
// placement to build the fragment. When a paragraph is not relocated after the
// estimate its final placement is identical, so the fragment now REUSES the fit
// measurement. This suite pins non-vacuity — pagination makes strictly fewer
// measureText calls with the reuse ON than OFF — and equivalence: the fragment line
// partitions and placements are identical either way (keyed on placement equality,
// so correctness never depends on the optimization).
// ─────────────────────────────────────────────────────────────────────────────

let measureCount = 0;

/** OffscreenCanvas polyfill whose measureText increments a shared counter (linear
 *  glyph metric = fontPx * 0.5, matching the other renderer suites). */
function makeCountingCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      measureCount++;
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      const per = p * 0.5;
      return {
        width: [...s].length * per,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

beforeAll(() => {
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    getContext() { return makeCountingCtx(); }
  };
});

function para(text: string, over: Partial<DocParagraph> = {}): DocParagraph {
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
    ...over,
  } as unknown as DocParagraph;
}

function doc(body: BodyElement[], pageHeight = 400): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 200, pageHeight,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage', columns: null,
  } as SectionProps;
  return {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
    footnotes: [],
  } as unknown as DocxDocumentModel;
}

/** A serialisable projection of a layout's fragment partitions + placements — enough
 *  to prove the reuse produced the SAME layout as a fresh measurement. */
function projection(model: DocxDocumentModel) {
  return layoutDocument(model).pages.map((page) =>
    page.fragments.map((placed) => {
      const f = placed.fragment as ParagraphFragment;
      return {
        lineStart: f.lineStart,
        lineEnd: f.lineEnd,
        lines: f.measured.lines.length,
        availW: f.measured.placement.availableWidthPt,
        startY: f.measured.placement.startYPt,
        heightPt: placed.heightPt,
        advances: f.measured.lines.map((l) => l.advancePt),
      };
    }),
  );
}

describe('M-1 fit-check measurement reuse for non-split body fragments', () => {
  it('reuses the fit measurement (fewer measureText calls) and produces the identical layout', () => {
    // Several short, non-splitting, float-free paragraphs on a tall page — each is
    // measured for the fit decision and then attached at the SAME placement, so the
    // reuse fires for every one of them.
    const model = doc([
      para('alpha beta gamma delta') as unknown as BodyElement,
      para('one two three four five six') as unknown as BodyElement,
      para('lorem ipsum dolor sit amet consectetur') as unknown as BodyElement,
      para('the quick brown fox jumps over') as unknown as BodyElement,
    ]);

    __test_setFitMeasureReuseEnabled(true);
    measureCount = 0;
    const onProjection = projection(model);
    const onCount = measureCount;

    __test_setFitMeasureReuseEnabled(false);
    measureCount = 0;
    const offProjection = projection(model);
    const offCount = measureCount;

    __test_setFitMeasureReuseEnabled(true); // restore production default

    // Non-vacuity: the reuse genuinely avoided a second measurement of every
    // non-relocated paragraph, so pagination measured strictly less text.
    expect(onCount).toBeGreaterThan(0);
    expect(onCount).toBeLessThan(offCount);

    // Equivalence: reusing the fit measurement yields the byte-same fragment layout as
    // measuring twice — correctness does not depend on the optimization.
    expect(onProjection).toEqual(offProjection);
  });

  it('does not affect a RELOCATED paragraph (placement changed → remeasured, layout still correct)', () => {
    // A tall paragraph with keepLines that must move to a fresh page: its fit estimate
    // is taken at a mid-page cursor, then the paragraph relocates, so the placement no
    // longer matches and the fragment remeasures. Toggling the reuse must not change
    // the resulting layout.
    const filler = para(Array.from({ length: 20 }, () => 'w').join(' '));
    const keep = para(Array.from({ length: 20 }, () => 'w').join(' '), { keepLines: true });
    const model = doc([filler as unknown as BodyElement, keep as unknown as BodyElement], 80);

    __test_setFitMeasureReuseEnabled(true);
    const on = projection(model);
    __test_setFitMeasureReuseEnabled(false);
    const off = projection(model);
    __test_setFitMeasureReuseEnabled(true);

    expect(on).toEqual(off);
  });
});
