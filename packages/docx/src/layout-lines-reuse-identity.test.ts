import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas, paginateDocument, __test_setLineReuseEnabled } from './renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps, PaginatedBodyElement } from './types';

// ─────────────────────────────────────────────────────────────────────────────
// Phase 4-1 B2 Stage 1 — pixel-identity of the compute-once line reuse.
//
// renderParagraph reuses the paginator's stamped scale-1 lines instead of
// re-running layoutLines (see the reuse gate in renderer.ts). This suite pins
// that the reuse is behaviour-PRESERVING: rendering the exact same page with
// reuse ON and with reuse OFF must emit a byte-identical paint call stream
// (every fillText / strokeText / drawImage with identical text, x, y and font).
//
// The pages are built with `paginateDocument` (a fresh OffscreenCanvas(1,1)) and
// handed to `renderDocumentToCanvas` via `prebuiltPages` — the SAME cross-context
// flow the public `DocxDocument.renderPage` uses, so the test exercises the real
// production reuse path (paginate ctx ≠ paint ctx), not a same-ctx shortcut. The
// render width equals the page width so the paint scale is exactly 1 and the
// reuse gate actually fires.
//
// Coverage: a long single-column paragraph that splits across pages (reuse fires),
// a justified paragraph (per-line slack distribution reads the reused segments),
// a CJK paragraph (per-glyph wrap), and a NUMBERED list that splits (reuse must
// NOT fire — the firstLineIndent gate rejects it — yet the recompute path is
// still identical to itself). Also pins that painting the same page twice is
// identical (the shared stamped array is never mutated by the draw path).
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

/** Assert reuse ON and reuse OFF paint an identical stream on every page, and
 *  report how many measureText calls each variant made so the caller can pin
 *  whether the reuse actually fired (fewer measures) or the gate rejected it
 *  (equal measures). */
async function assertReuseIdentical(model: DocxDocumentModel): Promise<{ pages: number; drawn: number; split: boolean; measuresOn: number; measuresOff: number; streams: Call[][] }> {
  const pages = paginateDocument(model);
  // Sanity: this document actually split a paragraph, so stamped lines exist.
  const split = pages.some((pg) => pg.some((el) => (el as PaginatedBodyElement).lineSlice));

  const prev = __test_setLineReuseEnabled(false);
  let off: { perPage: Call[][]; measures: number };
  try { off = await renderAllPages(model, pages); } finally { __test_setLineReuseEnabled(prev); }

  const on = await renderAllPages(model, pages);

  expect(on.perPage.length).toBe(off.perPage.length);
  let drawn = 0;
  for (let p = 0; p < on.perPage.length; p++) {
    // Exact stream identity — same ops, text, positions, fonts, in the same order.
    expect(on.perPage[p]).toEqual(off.perPage[p]);
    drawn += on.perPage[p].filter((c) => c.op !== 'img').length;
  }
  return { pages: pages.length, drawn, split, measuresOn: on.measures, measuresOff: off.measures, streams: on.perPage };
}

describe('compute-once line reuse — pixel identity (Phase 4-1 B2 Stage 1)', () => {
  it('long single-column paragraph that splits across pages: reuse ON === reuse OFF', async () => {
    const text = Array.from({ length: 120 }, () => 'w').join(' ');
    const r = await assertReuseIdentical(doc([para(text) as unknown as BodyElement]));
    expect(r.pages).toBeGreaterThan(1); // really split
    expect(r.split).toBe(true);         // stamped lines present
    expect(r.drawn).toBeGreaterThan(0); // really painted
    // The reuse actually fired: the paint pass skipped its own wrap-loop
    // measureText calls, so ON made strictly fewer than OFF.
    expect(r.measuresOn).toBeLessThan(r.measuresOff);
  });

  it('justified paragraph (both): slack distribution over reused segments is identical', async () => {
    const text = Array.from({ length: 120 }, (_, i) => (i % 3 === 0 ? 'lorem' : 'ipsum')).join(' ');
    const r = await assertReuseIdentical(doc([para(text, { alignment: 'both' }) as unknown as BodyElement]));
    expect(r.pages).toBeGreaterThan(1);
    expect(r.drawn).toBeGreaterThan(0);
    expect(r.measuresOn).toBeLessThan(r.measuresOff); // reuse fired
  });

  it('CJK paragraph (per-glyph wrap) that splits: reuse ON === reuse OFF', async () => {
    const text = 'あ'.repeat(200);
    const r = await assertReuseIdentical(doc([para(text) as unknown as BodyElement]));
    expect(r.pages).toBeGreaterThan(1);
    expect(r.drawn).toBeGreaterThan(0);
    expect(r.measuresOn).toBeLessThan(r.measuresOff); // reuse fired
  });

  it('numbered list that splits: reuse gate rejects it (firstLineIndent) yet paint is identical', async () => {
    // A numbered paragraph: measure lays out with para.indentFirst, paint with
    // numBodyOffset — the gate's firstIndent check fails, so reuse does NOT fire.
    // The recompute path must still be self-identical (this is the control that
    // proves the gate's rejection is safe and that toggling reuse is a no-op here).
    const numbering = { numId: 1, level: 0, format: 'decimal', text: '1.',
      indentLeft: 36, tab: 36, suff: 'tab', jc: 'left' } as unknown as DocParagraph['numbering'];
    const text = Array.from({ length: 120 }, () => 'w').join(' ');
    const p = para(text, {
      numbering, indentLeft: 36, indentFirst: -18,
    });
    const r = await assertReuseIdentical(doc([p as unknown as BodyElement]));
    expect(r.pages).toBeGreaterThan(1);
    expect(r.drawn).toBeGreaterThan(0);
    // The gate REJECTED reuse for the numbered list (firstLineIndent derivation
    // differs between measure and paint), so ON and OFF recomputed identically —
    // the measure counts are equal. This is what makes the gate non-vacuous: it
    // proves the reuse did NOT silently fire on a paragraph whose measure lines
    // would have painted wrong.
    expect(r.measuresOn).toBe(r.measuresOff);
  });

  it('NUMPAGES field in a splitting paragraph: never stamped — field text resolves against the real page context', async () => {
    // resolveFieldText is paint-state-dependent: numPages → state.totalPages,
    // which is 1 in the paginator's measure state but the real count at paint.
    // Stamped measure-time lines would freeze the stale "1" into the drawn text,
    // so paragraphSegsStateSensitive excludes such paragraphs from stamping —
    // they stay on the recompute path (the pre-reuse behaviour).
    const text = Array.from({ length: 120 }, () => 'w').join(' ');
    const p = para(text);
    (p.runs as unknown[]).push({
      type: 'field', fieldType: 'numPages', instruction: 'NUMPAGES', fallbackText: '?',
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', background: null,
    });
    const r = await assertReuseIdentical(doc([p as unknown as BodyElement]));
    expect(r.pages).toBeGreaterThan(1);
    expect(r.split).toBe(true);
    // No stamp → no reuse: ON and OFF recomputed identically.
    expect(r.measuresOn).toBe(r.measuresOff);
    // And the CURRENT total page count was drawn (not the measure-time "1" —
    // with pages > 1 the real count is distinguishable from the stale value).
    const drewTotal = r.streams.some((page) => page.some((c) => c.text === String(r.pages)));
    expect(drewTotal).toBe(true);
  });

  it('custom kinsoku settings: fresh-but-value-equal rule objects still reuse (=== alone would reject)', async () => {
    // The prebuiltPages production path resolves resolveKinsokuRules(doc.settings)
    // TWICE — once in paginateDocument, once in renderDocumentToCanvas — and the
    // resolver builds fresh Set objects per call. With custom settings the rules
    // are non-default on both sides yet reference-distinct; the gate's value
    // equivalence must still let the reuse fire.
    const text = Array.from({ length: 120 }, () => 'w').join(' ');
    const model = doc([para(text) as unknown as BodyElement], 60,
      { kinsoku: true, noLineBreaksBefore: '、。！' , noLineBreaksAfter: '（「' });
    const r = await assertReuseIdentical(model);
    expect(r.pages).toBeGreaterThan(1);
    expect(r.measuresOn).toBeLessThan(r.measuresOff); // reuse fired across fresh rule objects
  });

  it('same page rendered twice is identical (shared stamped array is never mutated)', async () => {
    const text = Array.from({ length: 120 }, () => 'w').join(' ');
    const model = doc([para(text) as unknown as BodyElement]);
    const pages = paginateDocument(model);
    const first = await renderAllPages(model, pages);
    const second = await renderAllPages(model, pages);
    for (let p = 0; p < first.perPage.length; p++) expect(second.perPage[p]).toEqual(first.perPage[p]);
  });
});
