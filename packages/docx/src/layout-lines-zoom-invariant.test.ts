import { describe, it, expect } from 'vitest';
import {
  renderDocumentToCanvas,
  paginateDocument,
  __test_setLineReuseEnabled,
  __test_layoutLines as layoutLines,
  type __test_LayoutSeg as LayoutSeg,
} from './renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps, PaginatedBodyElement } from './types';

// ─────────────────────────────────────────────────────────────────────────────
// Phase 4-1 B2 Stage 2 — ZOOM-INVARIANT line breaking.
//
// Word lays text out in the document's own coordinate space and treats the
// display scale as a viewport transform: the line PARTITION (which glyphs land
// on which line) is the same at every zoom. Stage 2 realises this by reusing the
// paginator's scale-1 line PARTITION at ANY paint scale — re-measuring the glyph
// geometry at the paint scale (so the pen tracks the glyphs) but never
// re-deciding the wrap points — instead of re-running layoutLines' break
// decisions at the paint scale, which under a real (hinted) font whose glyph
// advance is NOT proportional to the font px size would shift wrap points as the
// zoom changes (documented in layout-lines-scale-invariance.test.ts).
//
// This suite drives the REAL render path (renderParagraph) through a SUB-LINEAR
// mock canvas — glyph advance shrinks per-px as the size grows, the direction
// real hinting bends — and asserts the drawn per-line text partition is
// IDENTICAL at scale 1 and scale 0.75. A control proves the test is non-vacuous:
// fed the SAME segments, layoutLines' OWN break decisions wrap differently at the
// two scales under this font, so the invariance is a real property of the reuse.
// ─────────────────────────────────────────────────────────────────────────────

/** Sub-linear glyph advance: `px·(perPx − shrink·px)` per code point. NOT
 *  proportional to px, so a scale-1 layout and a scale-s re-layout would wrap
 *  differently — the exact hinting non-linearity Stage 2 must make invisible to
 *  the line partition. The SAME metric backs the paginate ctx (OffscreenCanvas)
 *  and every paint ctx, so scale-1 paginate and scale-1 paint agree. */
function subLinearWidth(text: string, fontPx: number): number {
  const per = Math.max(0.01, fontPx * (0.5 - 0.006 * fontPx));
  return [...text].length * per;
}

function makeMeasureStubCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: subLinearWidth(s, p),
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

interface Draw { text: string; x: number; y: number; }

/** A recording canvas backed by the SAME sub-linear metric. Records every text
 *  draw with its baseline y so the caller can reconstruct the per-line partition
 *  (glyphs sharing a baseline belong to one line). */
function makeRecordingCanvas(): { canvas: HTMLCanvasElement; draws: Draw[] } {
  let font = '10px serif';
  const draws: Draw[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: subLinearWidth(s, p),
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
    fillText(s: string, x: number, y: number) { if (s) draws.push({ text: s, x, y }); },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, draws };
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

// A SHORT page (content height ≈ 40pt ⇒ ~3 lines) so the long paragraph SPLITS
// across pages — the paginator only stamps its scale-1 lines on a split
// (splitParagraphAcrossPages), which is what arms the reuse path.
function doc(body: BodyElement[], pageHeight = 60): DocxDocumentModel {
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

/** Reconstruct the per-line TEXT partition from a paint stream: group draws by
 *  their baseline y (rounded to absorb sub-px), then concatenate each line's text
 *  in draw order. Independent of the absolute y values, so it compares across
 *  scales. */
function linePartition(draws: Draw[]): string[] {
  const byLine = new Map<number, { x: number; text: string }[]>();
  for (const d of draws) {
    const key = Math.round(d.y * 100) / 100;
    if (!byLine.has(key)) byLine.set(key, []);
    byLine.get(key)!.push({ x: d.x, text: d.text });
  }
  return [...byLine.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([, segs]) => segs.sort((a, b) => a.x - b.x).map((s) => s.text).join(''));
}

/** Render every page of `model` at `width` (⇒ paint scale = width/pageWidth) and
 *  return the concatenated per-line text partition across all pages. */
async function partitionAtWidth(model: DocxDocumentModel, pages: PaginatedBodyElement[][], width: number): Promise<string[]> {
  const all: string[] = [];
  for (let p = 0; p < pages.length; p++) {
    const rec = makeRecordingCanvas();
    await renderDocumentToCanvas(model, rec.canvas, p, { dpr: 1, width, prebuiltPages: pages });
    all.push(...linePartition(rec.draws));
  }
  return all;
}

describe('zoom-invariant line breaking (Phase 4-1 B2 Stage 2)', () => {
  it('a long paragraph draws the SAME per-line partition at scale 1 and scale 0.75', async () => {
    // 160 single-letter words → many wrap points; page width 200pt so it wraps
    // and splits across pages. Sub-linear font ⇒ a re-layout at 0.75 would wrap
    // differently, but the scale-1 stamp reuse pins the partition.
    const text = Array.from({ length: 160 }, (_, i) => `w${i % 10}`).join(' ');
    const model = doc([para(text) as unknown as BodyElement]);
    const pages = paginateDocument(model);
    expect(pages.length).toBeGreaterThan(1); // actually split

    const at1 = await partitionAtWidth(model, pages, 200);       // scale 1.0
    const at075 = await partitionAtWidth(model, pages, 150);      // scale 0.75
    expect(at075).toEqual(at1);
    // Non-vacuous: the paragraph really wrapped to multiple lines.
    expect(at1.length).toBeGreaterThan(3);
  });

  it('CJK per-glyph wrap: same partition at scale 1 and scale 0.5', async () => {
    const text = 'あ'.repeat(240);
    const model = doc([para(text) as unknown as BodyElement]);
    const pages = paginateDocument(model);
    expect(pages.length).toBeGreaterThan(1);
    const at1 = await partitionAtWidth(model, pages, 200);   // scale 1.0
    const at05 = await partitionAtWidth(model, pages, 100);  // scale 0.5
    expect(at05).toEqual(at1);
  });

  it('CONTROL: the mock font is genuinely NON-linear — a paint-scale re-layout WOULD wrap differently', () => {
    // Non-vacuity: proves the invariance above is a real property of Stage 2, not
    // an artefact of a metric that happens to be scale-clean. Feeding the SAME
    // segments to layoutLines at scale 1 vs scale 0.75 with this sub-linear
    // advance yields a DIFFERENT line PARTITION (fewer glyphs fit per line at the
    // smaller px size) — exactly the paint-scale re-break Stage 2 removed by
    // reusing the scale-1 stamp. If this ever passes trivially the render-level
    // assertions above would be meaningless.
    const seg = (text: string): LayoutSeg => ({
      text, bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', vertAlign: null,
      measuredWidth: 0,
    } as unknown as LayoutSeg);
    const text = Array.from({ length: 200 }, (_, i) => `w${i % 10}`).join(' ');

    // Same real content width (180pt) at two scales: scale-1 box 180, scale-0.75
    // box 135, first-indent 0. A linear font would give the SAME line count; the
    // sub-linear font does not.
    const a = layoutLines(makeMeasureStubCtx(), [seg(text)], 180, 0, 1);
    const b = layoutLines(makeMeasureStubCtx(), [seg(text)], 135, 0, 0.75);
    expect(a.length).toBeGreaterThan(1);
    // The partitions differ — different line count under the non-linear advance.
    expect(b.length).not.toBe(a.length);
  });

  it('reuse ON matches the scale-1 recompute partition (the stamp IS the scale-1 layout)', async () => {
    // Ties the two: the zoom-invariant partition equals what a fresh scale-1
    // layout produces (reuse OFF at scale 1), so Stage 2 did not silently change
    // the scale-1 answer — only made every other scale agree with it.
    const text = Array.from({ length: 160 }, (_, i) => `w${i % 10}`).join(' ');
    const model = doc([para(text) as unknown as BodyElement]);
    const pages = paginateDocument(model);

    const prev = __test_setLineReuseEnabled(false);
    let recomputeAt1: string[];
    try { recomputeAt1 = await partitionAtWidth(model, pages, 200); } finally { __test_setLineReuseEnabled(prev); }

    const reuseAt075 = await partitionAtWidth(model, pages, 150);
    expect(reuseAt075).toEqual(recomputeAt1);
  });

  it('an ANCHORED image in the paragraph adds no inline advance at any scale', async () => {
    // Anchored images live out of inline flow: layoutLines pins measuredWidth=0
    // and never adds them to line.segments (they are drawn by renderAnchorImages).
    // Characterizes that the reuse/rehydration path preserves this: the partition
    // is zoom-invariant AND every line's first glyph starts at the paragraph
    // origin (∝ scale) — not shifted right by imageWidth·scale, which is what a
    // rescale that conjured an inline advance for the anchor would produce.
    const text = Array.from({ length: 120 }, (_, i) => `w${i % 10}`).join(' ');
    const p = para(text);
    p.runs.unshift({
      type: 'image', imagePath: 'word/media/image1.png', mimeType: 'image/png',
      widthPt: 50, heightPt: 40, anchor: true, anchorXPt: 30, anchorYPt: 20,
    } as unknown as DocParagraph['runs'][number]);
    const model = doc([p as unknown as BodyElement]);
    const pages = paginateDocument(model);
    expect(pages.length).toBeGreaterThan(1);

    const collect = async (width: number): Promise<{ partition: string[]; firstX: number[] }> => {
      const partition: string[] = [];
      const firstX: number[] = [];
      for (let pg = 0; pg < pages.length; pg++) {
        const rec = makeRecordingCanvas();
        await renderDocumentToCanvas(model, rec.canvas, pg, { dpr: 1, width, prebuiltPages: pages });
        const byLine = new Map<number, Draw[]>();
        for (const d of rec.draws) {
          const key = Math.round(d.y * 100) / 100;
          if (!byLine.has(key)) byLine.set(key, []);
          byLine.get(key)!.push(d);
        }
        for (const [, draws] of [...byLine.entries()].sort((a, b) => a[0] - b[0])) {
          draws.sort((a, b) => a.x - b.x);
          partition.push(draws.map((d) => d.text).join(''));
          firstX.push(draws[0].x);
        }
      }
      return { partition, firstX };
    };

    const at1 = await collect(200);   // scale 1.0
    const at075 = await collect(150); // scale 0.75
    expect(at075.partition).toEqual(at1.partition);
    // Line starts scale linearly (page origin ∝ scale): x@0.75 = x@1 × 0.75.
    // A conjured 50pt inline advance would add 37.5px here and fail.
    expect(at075.firstX.length).toBe(at1.firstX.length);
    for (let i = 0; i < at1.firstX.length; i++) {
      expect(at075.firstX[i]).toBeCloseTo(at1.firstX[i] * 0.75, 6);
    }
  });
});
