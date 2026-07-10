import { describe, it, expect } from 'vitest';
import {
  renderDocumentToCanvas,
  paginateDocument,
  __test_setLineReuseEnabled,
  __test_setFragmentPaintEnabled,
} from './renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps, PaginatedBodyElement } from './types';

// ─────────────────────────────────────────────────────────────────────────────
// Body paint byte-identity — the compute-once line reuse (Phase 4-1 B2 Stage 1)
// AND the PR 5 body fragment paint.
//
// A migrated body paragraph now paints from its stored measured fragment
// (fragment-paint.ts → renderBodyParagraphLines): the fragment's scale-1 line
// partition is rescaled to the paint scale and drawn, with NO re-run of layoutLines.
// Marker / float / state-sensitive paragraphs stay on the legacy renderParagraph
// acquisition (its own scale-1 reuse gate). This suite pins that both mechanisms are
// behaviour-PRESERVING: rendering the exact same page three ways —
//   (1) production      : fragment paint ON,  reuse ON
//   (2) legacy reuse    : fragment paint OFF, reuse ON
//   (3) legacy recompute: fragment paint OFF, reuse OFF
// must emit a byte-identical paint call stream (every fillText / strokeText /
// drawImage with identical text, x, y and font). It also pins NON-VACUITY: fragment
// paint (or, for non-migrated paragraphs, the legacy reuse gate) actually avoided
// re-laying-out the paragraph, so production makes FEWER measureText calls than the
// legacy recompute — except where the paragraph is legitimately excluded from both
// fast paths (a numbered list's firstLineIndent, a NUMPAGES field), where the counts
// are equal.
//
// The pages are built with `paginateDocument` (a fresh OffscreenCanvas(1,1)) and
// handed to `renderDocumentToCanvas` via `prebuiltPages` — the SAME cross-context
// flow the public `DocxDocument.renderPage` uses. The render width equals the page
// width so the paint scale is exactly 1 (fragment rescale is then a no-op → a
// migrated paragraph paints with zero measureText calls).
// ─────────────────────────────────────────────────────────────────────────────

interface Call { op: 'fill' | 'stroke' | 'img'; text: string; x: number; y: number; font: string; }

/** `paginateDocument` builds its measure ctx from `new OffscreenCanvas(1,1)`,
 *  which the node test env lacks. Polyfill it with the SAME linear metric the
 *  recording paint canvas uses (glyph width = fontPx·0.5) so the paginate ctx and
 *  the paint ctx measure identically — the condition under which Stage 1
 *  guarantees byte-identical reuse. (Real cross-context metric parity is a browser
 *  Canvas invariant, exercised end-to-end by `pnpm vrt`.) */
function makeMeasureStubCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
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
(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
  getContext() { return makeMeasureStubCtx(); }
};

/** A recording context whose glyph advance is LINEAR in the font px size, so the
 *  scale-1 paginate lines and the scale-1 paint lines are bit-identical (this
 *  test is about the reuse mechanism, not font hinting). It records every text
 *  and image draw with position + font for an exact stream comparison. */
function makeRecordingCanvas(): { canvas: HTMLCanvasElement; calls: Call[]; measures: () => number } {
  let font = '10px serif';
  const calls: Call[] = [];
  let measures = 0;
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      measures++;
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      const per = p * 0.5;
      return {
        width: [...s].length * per,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage(_img: unknown, x: number, y: number) { calls.push({ op: 'img', text: '', x, y, font }); },
    fillText(s: string, x: number, y: number) { calls.push({ op: 'fill', text: s, x, y, font }); },
    strokeText(s: string, x: number, y: number) { calls.push({ op: 'stroke', text: s, x, y, font }); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls, measures: () => measures };
}

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

function doc(body: BodyElement[], pageHeight = 60, settings?: Record<string, unknown>): DocxDocumentModel {
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
    ...(settings ? { settings } : {}),
  } as unknown as DocxDocumentModel;
}

/** Render every page of `model` at scale 1 (width == pageWidth) via the shared
 *  prebuilt pages, returning the concatenated paint stream per page. */
async function renderAllPages(model: DocxDocumentModel, pages: PaginatedBodyElement[][]): Promise<{ perPage: Call[][]; measures: number }> {
  const perPage: Call[][] = [];
  let measures = 0;
  for (let p = 0; p < pages.length; p++) {
    const rec = makeRecordingCanvas();
    await renderDocumentToCanvas(model, rec.canvas, p, { dpr: 1, width: 200, prebuiltPages: pages });
    perPage.push(rec.calls);
    measures += rec.measures();
  }
  return { perPage, measures };
}

/** Render every page under an explicit (fragmentPaint, reuse) configuration, then
 *  restore the previous flags. */
async function renderVariant(
  model: DocxDocumentModel,
  pages: PaginatedBodyElement[][],
  cfg: { fragmentPaint: boolean; reuse: boolean },
): Promise<{ perPage: Call[][]; measures: number }> {
  const prevFragment = __test_setFragmentPaintEnabled(cfg.fragmentPaint);
  const prevReuse = __test_setLineReuseEnabled(cfg.reuse);
  try {
    return await renderAllPages(model, pages);
  } finally {
    __test_setLineReuseEnabled(prevReuse);
    __test_setFragmentPaintEnabled(prevFragment);
  }
}

/** Assert the production paint (fragment ON, reuse ON) is byte-identical to the
 *  legacy paint (fragment OFF) both with reuse ON and with reuse OFF, on every page.
 *  Reports measureText counts so the caller can pin non-vacuity (a fast path really
 *  fired, or was legitimately rejected). */
async function assertPaintIdentical(model: DocxDocumentModel): Promise<{ pages: number; drawn: number; split: boolean; measuresProduction: number; measuresRecompute: number; streams: Call[][] }> {
  const pages = paginateDocument(model);
  // Sanity: this document actually split a paragraph, so continuation slices exist.
  const split = pages.some((pg) => pg.some((el) => (el as PaginatedBodyElement).lineSlice));

  const production = await renderVariant(model, pages, { fragmentPaint: true, reuse: true });
  const legacyReuse = await renderVariant(model, pages, { fragmentPaint: false, reuse: true });
  const legacyRecompute = await renderVariant(model, pages, { fragmentPaint: false, reuse: false });

  expect(production.perPage.length).toBe(legacyReuse.perPage.length);
  expect(production.perPage.length).toBe(legacyRecompute.perPage.length);
  let drawn = 0;
  for (let p = 0; p < production.perPage.length; p++) {
    // Exact stream identity — same ops, text, positions, fonts, in the same order —
    // across fragment paint AND both legacy variants.
    expect(production.perPage[p]).toEqual(legacyReuse.perPage[p]);
    expect(production.perPage[p]).toEqual(legacyRecompute.perPage[p]);
    drawn += production.perPage[p].filter((c) => c.op !== 'img').length;
  }
  return {
    pages: pages.length,
    drawn,
    split,
    measuresProduction: production.measures,
    measuresRecompute: legacyRecompute.measures,
    streams: production.perPage,
  };
}

describe('body paint byte-identity — fragment paint and compute-once line reuse', () => {
  it('long single-column paragraph that splits across pages: fragment paint === legacy', async () => {
    const text = Array.from({ length: 120 }, () => 'w').join(' ');
    const r = await assertPaintIdentical(doc([para(text) as unknown as BodyElement]));
    expect(r.pages).toBeGreaterThan(1); // really split
    expect(r.split).toBe(true);         // continuation slices present
    expect(r.drawn).toBeGreaterThan(0); // really painted
    // Non-vacuity: fragment paint skipped the paragraph re-layout, so production
    // made strictly fewer measureText calls than the legacy recompute.
    expect(r.measuresProduction).toBeLessThan(r.measuresRecompute);
  });

  it('justified paragraph (both): slack distribution over fragment segments is identical', async () => {
    const text = Array.from({ length: 120 }, (_, i) => (i % 3 === 0 ? 'lorem' : 'ipsum')).join(' ');
    const r = await assertPaintIdentical(doc([para(text, { alignment: 'both' }) as unknown as BodyElement]));
    expect(r.pages).toBeGreaterThan(1);
    expect(r.drawn).toBeGreaterThan(0);
    expect(r.measuresProduction).toBeLessThan(r.measuresRecompute); // fragment paint fired
  });

  it('CJK paragraph (per-glyph wrap) that splits: fragment paint === legacy', async () => {
    const text = 'あ'.repeat(200);
    const r = await assertPaintIdentical(doc([para(text) as unknown as BodyElement]));
    expect(r.pages).toBeGreaterThan(1);
    expect(r.drawn).toBeGreaterThan(0);
    expect(r.measuresProduction).toBeLessThan(r.measuresRecompute); // fragment paint fired
  });

  it('numbered list that splits: excluded from fragment paint (firstLineIndent) yet paint is identical', async () => {
    // A numbered paragraph: the placement-aware measurement lays out with
    // para.indentFirst, but paint positions the body at numBodyOffset — so the
    // fragment's scale-1 lines would NOT reproduce the paint. isFragmentPaintable
    // excludes it (numbering != null) and the legacy reuse gate also rejects it, so
    // production falls back to the recompute path — which must still be identical to
    // itself. This is the control proving the exclusion is safe.
    const numbering = { numId: 1, level: 0, format: 'decimal', text: '1.',
      indentLeft: 36, tab: 36, suff: 'tab', jc: 'left' } as unknown as DocParagraph['numbering'];
    const text = Array.from({ length: 120 }, () => 'w').join(' ');
    const p = para(text, {
      numbering, indentLeft: 36, indentFirst: -18,
    });
    const r = await assertPaintIdentical(doc([p as unknown as BodyElement]));
    expect(r.pages).toBeGreaterThan(1);
    expect(r.drawn).toBeGreaterThan(0);
    // Neither fast path fired (fragment paint excluded, reuse rejected), so
    // production and the recompute path made the same number of measures.
    expect(r.measuresProduction).toBe(r.measuresRecompute);
  });

  it('NUMPAGES field in a splitting paragraph: excluded from fragment paint — field text resolves against the real page context', async () => {
    // resolveFieldText is paint-state-dependent: numPages → state.totalPages,
    // which is 1 in the paginator's measure state but the real count at paint.
    // The fragment's line segments freeze the stale "1", so paragraphSegsStateSensitive
    // excludes such paragraphs from BOTH the fragment path and the reuse stamp —
    // they recompute their segments against the real page context.
    const text = Array.from({ length: 120 }, () => 'w').join(' ');
    const p = para(text);
    (p.runs as unknown[]).push({
      type: 'field', fieldType: 'numPages', instruction: 'NUMPAGES', fallbackText: '?',
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', background: null,
    });
    const r = await assertPaintIdentical(doc([p as unknown as BodyElement]));
    expect(r.pages).toBeGreaterThan(1);
    expect(r.split).toBe(true);
    // No fast path: production and recompute made the same number of measures.
    expect(r.measuresProduction).toBe(r.measuresRecompute);
    // And the CURRENT total page count was drawn (not the measure-time "1" —
    // with pages > 1 the real count is distinguishable from the stale value).
    const drewTotal = r.streams.some((page) => page.some((c) => c.text === String(r.pages)));
    expect(drewTotal).toBe(true);
  });

  it('custom kinsoku settings: fragment paint stays identical across fresh-but-value-equal rule objects', async () => {
    // The prebuiltPages production path resolves resolveKinsokuRules(doc.settings)
    // TWICE — once in paginateDocument, once in renderDocumentToCanvas — building
    // fresh Set objects per call. The fragment's lines were laid out under the
    // paginate-time rules; the paint stays byte-identical.
    const text = Array.from({ length: 120 }, () => 'w').join(' ');
    const model = doc([para(text) as unknown as BodyElement], 60,
      { kinsoku: true, noLineBreaksBefore: '、。！' , noLineBreaksAfter: '（「' });
    const r = await assertPaintIdentical(model);
    expect(r.pages).toBeGreaterThan(1);
    expect(r.measuresProduction).toBeLessThan(r.measuresRecompute); // fragment paint fired
  });

  it('same page rendered twice is identical (the shared measured line array is never mutated by paint)', async () => {
    const text = Array.from({ length: 120 }, () => 'w').join(' ');
    const model = doc([para(text) as unknown as BodyElement]);
    const pages = paginateDocument(model);
    const first = await renderAllPages(model, pages);
    const second = await renderAllPages(model, pages);
    for (let p = 0; p < first.perPage.length; p++) expect(second.perPage[p]).toEqual(first.perPage[p]);
  });
});
