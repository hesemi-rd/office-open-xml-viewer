import { describe, it, expect } from 'vitest';
import {
  renderDocumentToCanvas,
  paginateDocument,
  createLayoutServices,
  bodyFragmentFor,
  __test_setBodyFragment,
  __test_setLineReuseEnabled,
  __test_setFragmentPaintEnabled,
  __test_tableRequiresLegacyPaint,
} from './renderer.js';
import { testFontSnapshot } from './layout/test-font-snapshot.js';
import type {
  BodyElement,
  CellElement,
  DocParagraph,
  DocTable,
  DocxDocumentModel,
  SectionProps,
  PaginatedBodyElement,
} from './types';

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

interface Call {
  op: 'fill' | 'stroke' | 'img';
  text: string;
  x: number;
  y: number;
  font: string;
  scaleX: number;
  scaleY: number;
}

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
  let transform = { scaleX: 1, scaleY: 1, translateX: 0, translateY: 0 };
  const transformStack: typeof transform[] = [];
  const record = (op: Call['op'], text: string, x: number, y: number): void => {
    calls.push({
      op,
      text,
      x: transform.translateX + transform.scaleX * x,
      y: transform.translateY + transform.scaleY * y,
      font,
      scaleX: transform.scaleX,
      scaleY: transform.scaleY,
    });
  };
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
    save() { transformStack.push({ ...transform }); },
    restore() { transform = transformStack.pop() ?? transform; },
    beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {},
    scale(sx: number, sy: number) {
      transform.scaleX *= sx;
      transform.scaleY *= sy;
    },
    translate(x: number, y: number) {
      transform.translateX += transform.scaleX * x;
      transform.translateY += transform.scaleY * y;
    },
    rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage(_img: unknown, ...args: number[]) {
      const x = args.length >= 8 ? args[4] : args[0];
      const y = args.length >= 8 ? args[5] : args[1];
      record('img', '', x ?? 0, y ?? 0);
    },
    fillText(s: string, x: number, y: number) { record('fill', s, x, y); },
    strokeText(s: string, x: number, y: number) { record('stroke', s, x, y); },
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
    const services = createLayoutServices(model, {
      localMetrics: testFontSnapshot(model), measureContext: rec.canvas.getContext('2d'),
    });
    await renderDocumentToCanvas(model, rec.canvas, p, { dpr: 1, width: 200, prebuiltPages: pages, layoutServices: services });
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

/** Render every page at the DEFAULT CSS scale (PT_TO_PX = 4/3, `width` OMITTED so
 *  scale = 4/3 ≠ 1) under an explicit (fragmentPaint, reuse) configuration; restore
 *  the flags after. At this scale the fragment path and the legacy reuse path both
 *  map their stored scale-1 partition through the canonical viewport transform,
 *  so their streams must be byte-identical. */
async function renderVariantScaled(
  model: DocxDocumentModel,
  pages: PaginatedBodyElement[][],
  cfg: { fragmentPaint: boolean; reuse: boolean },
): Promise<Call[][]> {
  const prevFragment = __test_setFragmentPaintEnabled(cfg.fragmentPaint);
  const prevReuse = __test_setLineReuseEnabled(cfg.reuse);
  try {
    const perPage: Call[][] = [];
    for (let p = 0; p < pages.length; p++) {
      const rec = makeRecordingCanvas();
      // No `width` → cssWidth = pageWidth · PT_TO_PX (4/3), so the paint scale is 4/3.
      await renderDocumentToCanvas(model, rec.canvas, p, { dpr: 1, prebuiltPages: pages });
      perPage.push(rec.calls);
    }
    return perPage;
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
  const pages = paginateDocument(model, createLayoutServices(model, { localMetrics: testFontSnapshot(model) }));
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

  it('zoom (default 4/3 scale): fragment paint === legacy stamp-reuse over the same scale-1 partition', async () => {
    // A long paragraph that SPLITS: pagination stamps its scale-1 line partition
    // (stampParagraphLines) AND builds the fragment from the SAME measured result, so
    // the legacy reuse path genuinely rescales the STAMP. At a paint scale ≠ 1 the
    // production fragment path and the legacy reuse path both rescale that one scale-1
    // partition, so their paint streams must be byte-identical (design invariant:
    // "paint scales measured geometry; it does not repeat text layout"). Identity vs
    // the full-RECOMPUTE path is intentionally NOT asserted here — recompute measures
    // at the paint scale (a documented pre-existing property), not this PR's concern.
    const text = Array.from({ length: 120 }, () => 'w').join(' ');
    const model = doc([para(text) as unknown as BodyElement]); // pageHeight 60 → splits
    const pages = paginateDocument(model);
    expect(pages.length).toBeGreaterThan(1);
    expect(pages.some((pg) => pg.some((el) => (el as PaginatedBodyElement).lineSlice))).toBe(true);

    const production = await renderVariantScaled(model, pages, { fragmentPaint: true, reuse: true });
    const reuse = await renderVariantScaled(model, pages, { fragmentPaint: false, reuse: true });
    expect(production.length).toBe(reuse.length);
    for (let p = 0; p < production.length; p++) {
      expect(production[p]).toEqual(reuse[p]);
    }
    // Non-vacuity: the glyphs retain their canonical 10px shaping font and are
    // mapped to the default 4/3 paint scale by the local Canvas transform.
    const scaled = production.flat().some((c) =>
      c.op !== 'img' && c.font.includes('10px') && Math.abs(c.scaleX - 4 / 3) < 1e-9);
    expect(scaled).toBe(true);
    expect(production.some((pg) => pg.length > 0)).toBe(true);
  });

  it('first-line AND hanging indents: fragment paint === legacy (3-way)', async () => {
    const text = Array.from({ length: 120 }, () => 'w').join(' ');
    // First-line indent (positive indentFirst) with left + right indents.
    const firstLine = await assertPaintIdentical(
      doc([para(text, { indentLeft: 24, indentRight: 12, indentFirst: 18 }) as unknown as BodyElement]),
    );
    expect(firstLine.drawn).toBeGreaterThan(0);
    expect(firstLine.measuresProduction).toBeLessThan(firstLine.measuresRecompute); // fragment paint fired
    // Hanging indent (negative first-line) WITHOUT numbering — still fragment-paintable
    // (the numbering exclusion is about numBodyOffset, not a bare hanging indent).
    const hanging = await assertPaintIdentical(
      doc([para(text, { indentLeft: 36, indentFirst: -18 }) as unknown as BodyElement]),
    );
    expect(hanging.drawn).toBeGreaterThan(0);
    expect(hanging.measuresProduction).toBeLessThan(hanging.measuresRecompute);
  });

  it('explicit tab stops + tab runs: fragment paint === legacy (3-way)', async () => {
    // A run with embedded tab characters (\t) and custom tab stops. layoutLines splits
    // on '\t' and advances to the stops (left + right-aligned with a dot leader); the
    // fragment's stored lines carry the tabbed geometry, painted without re-tabbing.
    const cell = Array.from({ length: 6 }, () => 'w').join(' ');
    const tabbed = `${cell}\t${cell}\t${cell}`;
    const tabStops = [
      { pos: 60, alignment: 'left', leader: 'none' },
      { pos: 130, alignment: 'right', leader: 'dot' },
    ] as unknown as DocParagraph['tabStops'];
    const p = para(Array.from({ length: 16 }, () => tabbed).join(' '), { tabStops });
    const r = await assertPaintIdentical(doc([p as unknown as BodyElement]));
    expect(r.drawn).toBeGreaterThan(0);
    expect(r.measuresProduction).toBeLessThan(r.measuresRecompute); // fragment paint fired
  });

  it('two-column section: fragment paint === legacy (3-way)', async () => {
    // ECMA-376 §17.6.4 newspaper columns — the body flows through 2 equal columns.
    // Each column-slice fragment records its column band width; paint sets contentW
    // per column, and the placement guard's width check passes (equal columns).
    const columns = { count: 2, spacePt: 12, equalWidth: true, sep: false, cols: [] } as unknown as SectionProps['columns'];
    const text = Array.from({ length: 200 }, () => 'w').join(' ');
    const model = doc([para(text) as unknown as BodyElement]);
    (model.section as SectionProps).columns = columns;
    const r = await assertPaintIdentical(model);
    expect(r.drawn).toBeGreaterThan(0);
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

  it('stale-placement fragment (a NEWER narrower measurement) → placement guard falls back to legacy, output correct', async () => {
    // Design invariant: "a measurement is valid only for its recorded placement".
    // A fragment measured for one placement must never be painted at another. Here the
    // width-180 prebuiltPages carry the paragraph p (its element stamps say the column
    // is 180 pt wide), but the side table is poisoned with a fragment measured at
    // width 100 (a 3-line partition) whose SOURCE is still p. Painting it would draw
    // the wrong (narrower) line breaks; the placement guard in
    // isFragmentPaintableParagraph detects the fragment's recorded availableWidthPt
    // (100) ≠ the paint column width (180) and falls back to legacy renderParagraph,
    // which reproduces the correct width-180 layout.
    //
    // The stale fragment is INJECTED rather than produced by a second pagination:
    // re-paginating the same p rewrites its colGeom/section stamps too, so the paint
    // width would track the stale fragment and no mismatch would be observable — a
    // faithful reproduction of "a stale side-table entry from a newer re-pagination
    // paints older prebuiltPages" without corrupting the element geometry.
    const p = para(Array.from({ length: 30 }, () => 'w').join(' '));
    const wideModel = doc([p as unknown as BodyElement], 600); // colW 180
    // A narrow pagination of the SAME p yields a width-100 fragment whose source is p.
    const narrowModel = doc([p as unknown as BodyElement], 600);
    (narrowModel.section as SectionProps).pageWidth = 120; // colW 100
    paginateDocument(narrowModel);
    const stale = bodyFragmentFor(paginateDocument(narrowModel)[0][0]);
    if (!stale) throw new Error('expected a narrow fragment to capture');
    if (stale.fragment.kind !== 'paragraph') throw new Error('expected a paragraph fragment');
    expect(stale.fragment.measured.placement.availableWidthPt).toBeCloseTo(100, 6);
    expect(stale.fragment.source).toBe(p); // same source paragraph, different placement
    // Re-paginate wide LAST so p's element stamps + prebuiltPages are the width-180 layout.
    const pagesWide = paginateDocument(wideModel);
    // Poison the side table for p's element with the stale width-100 fragment.
    __test_setBodyFragment(pagesWide[0][0], stale);

    const production = await renderVariant(wideModel, pagesWide, { fragmentPaint: true, reuse: true });
    const legacy = await renderVariant(wideModel, pagesWide, { fragmentPaint: false, reuse: true });
    expect(production.perPage.length).toBe(legacy.perPage.length);
    for (let pg = 0; pg < production.perPage.length; pg++) {
      // Guard falls back to legacy renderParagraph, so production === legacy (the
      // correct width-180 layout). Without the guard, production paints the width-100
      // fragment's 3-line partition and diverges — this is the Red case.
      expect(production.perPage[pg]).toEqual(legacy.perPage[pg]);
    }
    // Non-vacuity: the paragraph wrapped (multiple painted lines), so the stale
    // width-100 fragment would have been a visible divergence had the guard not fired.
    expect(production.perPage[0].filter((c) => c.op !== 'img').length).toBeGreaterThan(1);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// PR 6 Task 16 — TABLE-CELL paint byte-identity. A migrated block table paints its
// cell paragraphs from CellFragment blocks; the per-block gate must exclude the SAME
// divergence classes the PR 5 body gate excludes (marker firstLineIndent /
// state-sensitive fields / in-cell float wrap), falling back to the legacy
// renderParagraph — otherwise the fragment paints a partition the legacy paint would
// never draw. Pinned against sample class: numbered paragraph inside a table cell
// (the §17.9.28 marker's numBodyOffset ≠ para.indentFirst).
// ─────────────────────────────────────────────────────────────────────────────

function cellOf(content: unknown[], widthPt: number): Record<string, unknown> {
  return {
    content, colSpan: 1, vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background: null, vAlign: 'top', widthPt,
  };
}

function tableOf(
  cellContent: unknown[],
  overrides: {
    colWidths?: number[];
    widthPt?: number;
    tblInd?: number;
    jc?: string;
    layout?: string;
  } = {},
): BodyElement {
  const colWidths = overrides.colWidths ?? [180];
  const widthPt = overrides.widthPt ?? colWidths[0] ?? 180;
  return {
    type: 'table',
    colWidths,
    rows: [{ cells: [cellOf(cellContent, widthPt)], rowHeight: null, rowHeightRule: 'auto', isHeader: false }],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: overrides.jc ?? 'left', layout: overrides.layout ?? 'fixed',
    ...(overrides.tblInd == null ? {} : { tblInd: overrides.tblInd }),
  } as unknown as BodyElement;
}

describe('table fragment-paint legacy gate', () => {
  it('excludes only negative leading tblInd tables, including nested tables', () => {
    const negative = tableOf([para('negative')], { tblInd: -30 }) as unknown as DocTable;
    const nestedNegative = tableOf([para('nested')], { tblInd: -12 }) as unknown as DocTable;
    const outer = tableOf([nestedNegative]) as unknown as DocTable;
    const positive = tableOf([para('positive')], { tblInd: 12 }) as unknown as DocTable;
    const centeredNegative = tableOf([para('centered')], {
      tblInd: -30,
      jc: 'center',
    }) as unknown as DocTable;

    expect(__test_tableRequiresLegacyPaint(negative)).toBe(true);
    expect(__test_tableRequiresLegacyPaint(outer)).toBe(true);
    expect(__test_tableRequiresLegacyPaint(positive)).toBe(false);
    expect(__test_tableRequiresLegacyPaint(centeredNegative)).toBe(false);
  });

  it('permits a table containing a sliced cell paragraph', () => {
    const sliced = {
      ...para('sliced'),
      lineSlice: { start: 0, end: 2 },
    } as unknown as CellElement;
    const table = tableOf([sliced]) as unknown as DocTable;

    expect(__test_tableRequiresLegacyPaint(table)).toBe(false);
  });

  it('permits production row slices emitted from a tall cell paragraph', () => {
    const cellPara = para(Array.from({ length: 40 }, () => 'wrap').join(' '));
    const pages = paginateDocument(doc([tableOf([cellPara])], 60));
    const sliceTables = pages
      .flatMap((page) => page)
      .filter((el): el is PaginatedBodyElement & DocTable => el.type === 'table');
    const slicedTables = sliceTables.filter((table) =>
      table.rows.some((row) =>
        row.cells.some((cell) =>
          cell.content.some(
            (ce) =>
              ce.type === 'paragraph' &&
              (ce as CellElement & { lineSlice?: unknown }).lineSlice !== undefined,
          ),
        ),
      ),
    );

    expect(sliceTables.length).toBeGreaterThan(1);
    expect(slicedTables.length).toBeGreaterThan(0);
    for (const table of slicedTables) {
      expect(__test_tableRequiresLegacyPaint(table)).toBe(false);
    }
  });
});

describe('table-cell paint byte-identity — fragment table paint (PR 6)', () => {
  it('negative leading tblInd table: fragment paint falls back to the page-width legacy budget', async () => {
    const cellPara = para(Array.from({ length: 50 }, () => 'wide').join(' '));
    const model = doc([
      tableOf([cellPara], {
        colWidths: [220],
        widthPt: 220,
        tblInd: -30,
        jc: 'left',
        layout: 'fixed',
      }),
    ], 400);

    const r = await assertPaintIdentical(model);
    expect(r.drawn).toBeGreaterThan(0);
  });

  it('numbered paragraph inside a table cell: fragment paint === legacy (marker indent)', async () => {
    // Legacy cell paint recomputes a NUMBERED paragraph with the marker-aware
    // numBodyOffset first-line indent (the stamp gate rejects it); the fragment's
    // stored lines were measured with para.indentFirst. The per-block gate must fall
    // back to legacy renderParagraph for numbered cell paragraphs.
    const numbering = { numId: 1, level: 0, format: 'decimal', text: '1.',
      indentLeft: 36, tab: 36, suff: 'tab', jc: 'left' } as unknown as DocParagraph['numbering'];
    const cellPara = {
      type: 'paragraph',
      ...para(Array.from({ length: 30 }, () => 'w').join(' '), {
        numbering, indentLeft: 36, indentFirst: -18,
      }),
    };
    const model = doc([tableOf([cellPara])], 400);
    const r = await assertPaintIdentical(model);
    expect(r.drawn).toBeGreaterThan(0);
    // Non-vacuity: the marker itself was drawn.
    expect(r.streams.some((page) => page.some((c) => c.text === '1.'))).toBe(true);
  });

  it('NUMPAGES field inside a table cell: fragment paint === legacy (state-sensitive)', async () => {
    // The fragment's segments freeze the measure-time totalPages (1); legacy paint
    // re-resolves the field against the real page context. The per-block gate must
    // fall back for paragraphSegsStateSensitive cell paragraphs.
    const fieldPara = {
      type: 'paragraph',
      ...para('total pages:'),
    } as unknown as DocParagraph;
    (fieldPara.runs as unknown[]).push({
      type: 'field', fieldType: 'numPages', instruction: 'NUMPAGES', fallbackText: '?',
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', background: null,
    });
    // A long body paragraph AFTER the table forces a second page, so the real
    // totalPages (2) differs from the measure-time value (1).
    const filler = para(Array.from({ length: 120 }, () => 'w').join(' '));
    const model = doc([tableOf([fieldPara]), filler as unknown as BodyElement], 100);
    const r = await assertPaintIdentical(model);
    expect(r.pages).toBeGreaterThan(1);
    // The REAL page count was drawn (not the frozen measure-time "1").
    expect(r.streams.some((page) => page.some((c) => c.text === String(r.pages)))).toBe(true);
  });

  it('cell paragraph after an in-cell wrap float: fragment paint === legacy (wrap context)', async () => {
    // Cell paragraph 1 carries an anchored square-wrap image (registered into the
    // cell's isolated float set during paint); paragraph 2's legacy paint re-lays-out
    // around it (wrap context), while its fragment lines were measured with no wrap
    // oracle. The per-block gate must fall back when the cell's float set is non-empty.
    const imgPara = {
      type: 'paragraph',
      ...para(''),
    } as unknown as DocParagraph;
    (imgPara.runs as unknown[]).push({
      type: 'image',
      imagePath: 'word/media/test1.png', mimeType: 'image/png',
      widthPt: 80, heightPt: 40,
      anchor: true, anchorXPt: 100, anchorYPt: 0,
      anchorXFromMargin: false, anchorYFromPara: true,
      wrapMode: 'square', wrapSide: 'bothSides',
      anchorXRelativeFrom: 'column', anchorYRelativeFrom: 'paragraph',
    });
    const wrapPara = {
      type: 'paragraph',
      ...para(Array.from({ length: 40 }, () => 'w').join(' ')),
    };
    const model = doc([tableOf([imgPara, wrapPara])], 400);
    const r = await assertPaintIdentical(model);
    expect(r.drawn).toBeGreaterThan(0);
  });
});

describe('re-wrapped continuation slices (issue #908) — fragment paint parity', () => {
  it.each([
    ['fragment paint', true],
    ['legacy paint', false],
  ])('does not re-apply first-line indent while painting a continuation (%s)', async (_label, fragmentPaint) => {
    const p = para('あ'.repeat(28), { defaultFontSize: 20, indentFirst: 20 });
    (p.runs[0] as { fontSize: number }).fontSize = 20;
    const model = doc([p as unknown as BodyElement], 60);
    (model.section as SectionProps).columns = {
      count: 2, spacePt: 32, equalWidth: false, sep: false,
      cols: [{ widthPt: 100, spacePt: 32 }, { widthPt: 48, spacePt: 0 }],
    } as SectionProps['columns'];

    const pages = paginateDocument(model);
    const continuationPage = pages.findIndex((page) => page.some((el) => {
      const slice = el as PaginatedBodyElement & {
        lineSlice?: { start: number; end: number; continues?: boolean };
      };
      return slice.colIndex === 1 && slice.lineSlice?.continues === true;
    }));
    expect(continuationPage).toBeGreaterThanOrEqual(0);

    const painted = await renderVariant(model, pages, { fragmentPaint, reuse: true });
    const col1Origin = model.section.marginLeft + 100 + 32;
    const lineCalls = painted.perPage[continuationPage]
      .filter((call) => call.op === 'fill' && call.text.includes('あ') && call.x >= col1Origin);
    const callsByY = new Map<number, Call[]>();
    for (const call of lineCalls) {
      const calls = callsByY.get(call.y);
      if (calls) calls.push(call);
      else callsByY.set(call.y, [call]);
    }
    const lineStarts = [...callsByY.entries()]
      .sort(([a], [b]) => a - b)
      .map(([, calls]) => Math.min(...calls.map((call) => call.x)));

    expect(lineStarts.length).toBeGreaterThan(1);
    expect(lineStarts.slice(1).every((x) => x === lineStarts[0])).toBe(true);
  });

  it('draws paragraph anchors ONCE for a re-wrapped continuation (fragment === legacy)', async () => {
    // A §17.6.4 unequal-width split: col0 keeps the wide partition, col1 gets a
    // RE-MEASURED remainder whose fragment covers its WHOLE partition (lineStart 0,
    // lineEnd = length). paintParagraphFragment must not degrade that full-range
    // continuation slice to "no slice": renderParagraph's first-slice-only work
    // (anchor drawing, §17.3.1.7 top border) would run again in the destination
    // column. Legacy paint sees the element's `continues` flag; the fragment path
    // must thread it through. The anchored no-wrap shape's text is the observable:
    // its fillText must appear exactly once per page in BOTH paint modes.
    const anchorShape = {
      type: 'shape',
      widthPt: 40, heightPt: 10,
      anchorXPt: 0, anchorYPt: 0,
      anchorXFromMargin: false, anchorYFromPara: true,
      anchorXRelativeFrom: 'column', anchorYRelativeFrom: 'paragraph',
      zOrder: 0, subpaths: [], presetGeometry: 'rect',
      fill: null, stroke: null,
      // No wrap: the shape registers no float band, so the width-mismatch swap
      // fires (a wrap float would scope the swap out — asserted below).
      wrapMode: 'none', wrapSide: 'bothSides',
      distTop: 0, distBottom: 0, distLeft: 0, distRight: 0,
      textBlocks: [{
        text: 'ANCHORTEXT', fontSizePt: 8, bold: false, italic: false,
        color: '000000', fontFamily: 'Times New Roman', alignment: 'left',
      }],
      textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
    };
    const p = para('あ'.repeat(28), { defaultFontSize: 20 });
    (p.runs as unknown[]).unshift(anchorShape);
    (p.runs[1] as { fontSize: number }).fontSize = 20;
    const model = doc([p as unknown as BodyElement], 60);
    (model.section as SectionProps).columns = {
      count: 2, spacePt: 32, equalWidth: false, sep: false,
      cols: [{ widthPt: 100, spacePt: 32 }, { widthPt: 48, spacePt: 0 }],
    } as SectionProps['columns'];

    const pages = paginateDocument(model);
    // Non-vacuity: the swap really fired — a col-1 slice carries the remainder
    // partition marker (if a float had registered, the swap would be scoped out
    // and this test would be exercising nothing).
    const slices = pages.flat().filter((el) => el.type === 'paragraph') as
      (PaginatedBodyElement & { lineSlice?: { start: number; end: number; continues?: boolean } })[];
    expect(slices.some((s) => s.lineSlice?.continues === true)).toBe(true);

    const production = await renderVariant(model, pages, { fragmentPaint: true, reuse: true });
    const legacy = await renderVariant(model, pages, { fragmentPaint: false, reuse: true });
    expect(production.perPage.length).toBe(legacy.perPage.length);
    for (let pg = 0; pg < production.perPage.length; pg++) {
      // Byte-identical streams — in particular the anchor's text appears the same
      // number of times (once) in both modes.
      const anchorDrawsProduction = production.perPage[pg].filter((c) => c.text === 'ANCHORTEXT').length;
      const anchorDrawsLegacy = legacy.perPage[pg].filter((c) => c.text === 'ANCHORTEXT').length;
      expect(anchorDrawsProduction).toBe(anchorDrawsLegacy);
      expect(anchorDrawsProduction).toBe(1);
      expect(anchorDrawsLegacy).toBe(1);
      expect(production.perPage[pg]).toEqual(legacy.perPage[pg]);
    }
    // The anchor really painted (observable non-vacuity).
    expect(legacy.perPage.flat().some((c) => c.text === 'ANCHORTEXT')).toBe(true);
  });

  it('draws behindDoc paragraph anchors once for a re-wrapped continuation', async () => {
    const anchorShape = {
      type: 'shape',
      widthPt: 40, heightPt: 10,
      anchorXPt: 0, anchorYPt: 0,
      anchorXFromMargin: false, anchorYFromPara: true,
      anchorXRelativeFrom: 'column', anchorYRelativeFrom: 'paragraph',
      zOrder: 0, behindDoc: true,
      subpaths: [], presetGeometry: 'rect',
      fill: null, stroke: null,
      wrapMode: 'none', wrapSide: 'bothSides',
      distTop: 0, distBottom: 0, distLeft: 0, distRight: 0,
      textBlocks: [{
        text: 'ANCHORTEXT', fontSizePt: 8, bold: false, italic: false,
        color: '000000', fontFamily: 'Times New Roman', alignment: 'left',
      }],
      textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
    };
    const p = para('あ'.repeat(28), { defaultFontSize: 20 });
    (p.runs as unknown[]).unshift(anchorShape);
    (p.runs[1] as { fontSize: number }).fontSize = 20;
    const model = doc([p as unknown as BodyElement], 60);
    (model.section as SectionProps).columns = {
      count: 2, spacePt: 32, equalWidth: false, sep: false,
      cols: [{ widthPt: 100, spacePt: 32 }, { widthPt: 48, spacePt: 0 }],
    } as SectionProps['columns'];

    const pages = paginateDocument(model);
    const slices = pages.flat().filter((el) => el.type === 'paragraph') as
      (PaginatedBodyElement & { lineSlice?: { start: number; end: number; continues?: boolean } })[];
    expect(slices.some((s) => s.lineSlice?.continues === true)).toBe(true);

    const production = await renderVariant(model, pages, { fragmentPaint: true, reuse: true });
    const legacy = await renderVariant(model, pages, { fragmentPaint: false, reuse: true });
    expect(production.perPage.length).toBe(legacy.perPage.length);
    for (let pg = 0; pg < production.perPage.length; pg++) {
      const anchorDrawsProduction = production.perPage[pg].filter((c) => c.text === 'ANCHORTEXT').length;
      const anchorDrawsLegacy = legacy.perPage[pg].filter((c) => c.text === 'ANCHORTEXT').length;
      expect.soft(anchorDrawsProduction).toBe(1);
      expect.soft(anchorDrawsLegacy).toBe(1);
    }
  });
});
