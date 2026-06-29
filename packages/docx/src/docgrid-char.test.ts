import { describe, it, expect } from 'vitest';
import { computePages, renderDocumentToCanvas, type DocxTextRunInfo } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
  PaginatedBodyElement,
} from './types';

// ECMA-376 §17.6.5 docGrid CHARACTER grid (字詰め). These tests guard the ONE
// thing the feature must never break: the line-break MEASUREMENT and the draw
// ADVANCE use the SAME per-character cell width. A previous attempt corrupted
// the layout (overlapping / scrambled glyphs) because measure and draw diverged.
//
// The stub canvas models the grid arithmetic exactly: a glyph's natural advance
// is `fontPx` (= charCount × fontPx for a string) and the font box is 0.8/0.2 em.
// Under an active character grid with `charSpace` raw value C, the per-EA-glyph
// delta is Δpt = C/4096, Δpx = Δpt × scale; every full-width EA glyph then
// occupies a cell of `fontPx + Δpx`. CJK characters break between any two glyphs.

const FONT_PX = 20; // glyph advance per CJK char in the stub (scale = 1)

/** Recording 2D context: glyph advance = charCount × fontPx, font box 0.8/0.2 em.
 *  Records every fillText so per-glyph draw positions can be asserted. */
function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  fillTextCalls: { text: string; x: number; y: number; letterSpacing: string }[];
} {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const fillTextCalls: { text: string; x: number; y: number; letterSpacing: string }[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    // measureText models the per-glyph natural advance only (no letterSpacing):
    // the grid draw applies its per-cell delta Δ via ctx.letterSpacing, and the
    // renderer measures pieces BEFORE setting letterSpacing, so the stub keeps the
    // (natural) measure==(box) invariant while letterSpacing carries Δ for draw.
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(v: string) { letterSpacing = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText(text: string, x: number, y: number) { fillTextCalls.push({ text, x, y, letterSpacing }); },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0,
    height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, fillTextCalls };
}

function textRun(text: string, fontSize: number, fontFamily = 'NotInMetrics'): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily, isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  };
}

type DocRun = DocParagraph['runs'][number];

function para(text: string, opts: { fontSize?: number; alignment?: string } = {}): BodyElement {
  const fontSize = opts.fontSize ?? FONT_PX;
  const p: DocParagraph = {
    alignment: opts.alignment ?? 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{ type: 'text', ...textRun(text, fontSize) } as DocRun],
    defaultFontSize: fontSize, defaultFontFamily: 'NotInMetrics',
    widowControl: false, // keep greedy split deterministic
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

/** A wide section so a paragraph fits on one line unless the grid changes the
 *  per-char width; tall enough that vertical fit never wraps in these tests. */
function section(overrides: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 200, pageHeight: 400,
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    ...overrides,
  };
}

/** linesAndChars grid with the given raw charSpace and a line pitch. */
function charGrid(charSpace: number): Partial<SectionProps> {
  return { docGridType: 'linesAndChars', docGridLinePitch: 20, docGridCharSpace: charSpace };
}

function doc(body: BodyElement[], sec: SectionProps): DocxDocumentModel {
  return {
    section: sec,
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

const sliceOf = (el: PaginatedBodyElement) =>
  (el as { lineSlice?: { start: number; end: number } }).lineSlice;

/** Lines a paragraph occupies on its (single) page = number of line slices, or 1
 *  when the paragraph fits on one line (no slice tag). For these single-paragraph
 *  fixtures we count the rendered text runs per line instead by re-deriving from
 *  computePages line slices is fragile, so we render and count distinct baselines. */
async function renderRun(
  body: BodyElement[],
  sec: SectionProps,
): Promise<{
  runs: DocxTextRunInfo[];
  fillTextCalls: { text: string; x: number; y: number; letterSpacing: string }[];
}> {
  const { canvas, fillTextCalls } = makeRecordingCanvas();
  const runs: DocxTextRunInfo[] = [];
  await renderDocumentToCanvas(doc(body, sec), canvas, 0, {
    dpr: 1,
    width: sec.pageWidth, // scale = 1 (px per pt)
    onTextRun: (r) => runs.push(r),
  });
  return { runs, fillTextCalls };
}

describe('docGrid character grid — measure==draw invariant (§17.6.5)', () => {
  // THE core anti-corruption guard. For a CJK string under an active char grid,
  // the measured segment box (onTextRun.w) and the painted advance must be
  // derived from the SAME per-char cell width fontPx + Δpx. A no-justify pure-EA
  // segment is now painted as ONE contiguous fillText (contextual shaping ⇒
  // 約物半角 honoured, no bracket overlap) with the per-cell delta carried by
  // ctx.letterSpacing = Δ. The painted box edge = measure(whole) + n·Δ = the
  // measured box, by construction (see justify-positions.ts / grid-bracket-overlap).
  it('contiguous draw with letterSpacing=Δ matches the measured box width exactly', async () => {
    const charSpace = -1161; // sample-10's value
    const deltaPx = charSpace / 4096; // Δpt × scale(=1)
    const cell = FONT_PX + deltaPx; // per-CJK-glyph cell width
    const text = 'あいうえお'; // 5 full-width EA glyphs, no spaces → one segment
    const n = [...text].length;

    const { runs, fillTextCalls } = await renderRun([para(text)], section(charGrid(charSpace)));

    // One reported text run (the whole CJK segment on one line).
    const seg = runs.find((r) => r.text === text);
    expect(seg).toBeDefined();

    // MEASURE: the segment box width is exactly n cells.
    expect(seg!.w).toBeCloseTo(n * cell, 6);

    // DRAW: the pure-EA segment is painted as ONE contiguous fillText (not n
    // isolated per-code-point draws — the previous, bracket-overlapping path).
    const whole = fillTextCalls.filter((c) => c.text === text);
    expect(whole.length).toBe(1);
    // It starts at the segment's measured origin…
    expect(whole[0].x).toBeCloseTo(seg!.x, 6);
    // …and the per-cell grid delta is applied via ctx.letterSpacing, so the
    // browser advances each glyph by measure(glyph)+Δ. With measure(whole) baked
    // into seg.w and letterSpacing=Δ adding n·Δ, the painted box edge equals the
    // measured box edge: the next segment abuts with no overlap and no gap. This
    // is the invariant the previous (corrupting) per-code-point attempt violated.
    expect(whole[0].letterSpacing).toBe(`${deltaPx}px`);
    // No isolated single-code-point EA draw exists for this segment.
    const isolated = fillTextCalls.filter(
      (c) => [...c.text].length === 1 && [...text].includes(c.text),
    );
    expect(isolated.length).toBe(0);
  });

  // A space-free MIXED CJK+Latin token (no U+0020, so `splitTextForLayout` keeps
  // it as one word). The §17.3.2.26 ascii/eastAsia split sub-divides it into
  // per-script segments [あ][A][本]; under an active grid the pure-EA segments
  // get the cell delta while the Latin segment keeps its natural advance. This
  // guards two things at once: (1) measure==draw still holds across the mixed
  // token (each single-font segment's box abuts the next — no overlap), and
  // (2) the EA glyphs ARE gridded even when sandwiching Latin (the behaviour
  // that changed: before the split, the whole mixed token was one non-pure-EA
  // segment and its CJK chars were NOT snapped). It also pins the load-bearing
  // interaction between `splitByEastAsia` (isCjkBreakChar) and the grid's own
  // `EAST_ASIAN_RE` purity test, so a future predicate edit can't silently break it.
  it('mixed CJK+Latin token: EA sub-segments grid, Latin keeps natural width, boxes abut', async () => {
    const charSpace = -1161;
    const cell = FONT_PX + charSpace / 4096; // per-CJK-glyph cell
    const { runs, fillTextCalls } = await renderRun([para('あA本')], section(charGrid(charSpace)));

    const segA1 = runs.find((r) => r.text === 'あ');
    const segLat = runs.find((r) => r.text === 'A');
    const segA2 = runs.find((r) => r.text === '本');
    expect(segA1, 'EA segment あ').toBeDefined();
    expect(segLat, 'Latin segment A').toBeDefined();
    expect(segA2, 'EA segment 本').toBeDefined();

    // MEASURE: EA segments are one grid cell; the Latin segment is its natural
    // advance (NOT gridded).
    expect(segA1!.w).toBeCloseTo(cell, 6);
    expect(segA2!.w).toBeCloseTo(cell, 6);
    expect(segLat!.w).toBeCloseTo(FONT_PX, 6);

    // measure==draw + abutment: each segment's box starts exactly where the prior
    // one ends, so glyphs never overlap and the line width is the sum of cells.
    expect(segLat!.x).toBeCloseTo(segA1!.x + segA1!.w, 6);
    expect(segA2!.x).toBeCloseTo(segLat!.x + segLat!.w, 6);

    // DRAW lands at each segment's measured origin (EA glyphs at their cell start,
    // the Latin glyph at its segment x).
    expect(fillTextCalls.find((c) => c.text === 'あ')!.x).toBeCloseTo(segA1!.x, 6);
    expect(fillTextCalls.find((c) => c.text === 'A')!.x).toBeCloseTo(segLat!.x, 6);
    expect(fillTextCalls.find((c) => c.text === '本')!.x).toBeCloseTo(segA2!.x, 6);
  });

  it('a negative charSpace tightens the box (< natural); positive loosens it', async () => {
    const text = 'あいうえお';
    const n = [...text].length;

    const tight = await renderRun([para(text)], section(charGrid(-1161)));
    const loose = await renderRun([para(text)], section(charGrid(+2048)));
    const none = await renderRun([para(text)], section()); // no grid

    const w = (r: { runs: DocxTextRunInfo[] }) => r.runs.find((x) => x.text === text)!.w;
    expect(w(none)).toBeCloseTo(n * FONT_PX, 6);
    expect(w(tight)).toBeLessThan(w(none));
    expect(w(loose)).toBeGreaterThan(w(none));
    // Exact cell arithmetic both ways.
    expect(w(tight)).toBeCloseTo(n * (FONT_PX + -1161 / 4096), 6);
    expect(w(loose)).toBeCloseTo(n * (FONT_PX + 2048 / 4096), 6);
  });

  it('does NOT snap Latin text (Latin spans cells at natural advance)', async () => {
    const text = 'abcde'; // 5 non-EA glyphs → one segment, never gridded
    const n = [...text].length;
    const { runs, fillTextCalls } = await renderRun([para(text)], section(charGrid(-1161)));
    const seg = runs.find((r) => r.text === text)!;
    // Box is the natural width — the grid delta does not apply to Latin.
    expect(seg.w).toBeCloseTo(n * FONT_PX, 6);
    // Drawn as a single fillText at the natural position (no per-glyph walk).
    const whole = fillTextCalls.filter((c) => c.text === text);
    expect(whole.length).toBe(1);
    expect(whole[0].x).toBeCloseTo(seg.x, 6);
  });

  it('type="lines" (line grid, no char grid) leaves EA glyphs at natural advance', async () => {
    const text = 'あいうえお';
    const n = [...text].length;
    const sec = section({ docGridType: 'lines', docGridLinePitch: 20, docGridCharSpace: -1161 });
    const { runs } = await renderRun([para(text)], sec);
    const seg = runs.find((r) => r.text === text)!;
    // charSpace present but type is "lines" ⇒ the CHARACTER grid is inactive.
    expect(seg.w).toBeCloseTo(n * FONT_PX, 6);
  });

  it('an absent charSpace leaves EA glyphs at natural advance even under linesAndChars', async () => {
    const text = 'あいうえお';
    const n = [...text].length;
    const sec = section({ docGridType: 'linesAndChars', docGridLinePitch: 20 });
    const { runs } = await renderRun([para(text)], sec);
    const seg = runs.find((r) => r.text === text)!;
    expect(seg.w).toBeCloseTo(n * FONT_PX, 6);
  });
});

describe('docGrid character grid — packs more chars per line (§17.6.5)', () => {
  // The pagination payoff: a negative charSpace makes each EA cell narrower, so
  // MORE characters fit per line ⇒ FEWER wrapped lines ⇒ less vertical space
  // (sample-10 fits one page instead of overflowing to two). The page is one
  // line tall (pitch 20, pageHeight 25), so each wrapped line becomes its own
  // page slice and the total line count = Σ slice spans across pages.
  //
  // contentW = 200, fontPx 20 → natural 10 chars/line. charSpace -8192 ⇒ Δ = -2pt
  // ⇒ cell 18 ⇒ 11 chars/line. 22 chars: natural 3 lines, tightened 2 lines.
  const ONE_LINE_PAGE = { pageHeight: 25 };

  const linesOf = (text: string, sec: SectionProps): number => {
    const pages = computePages(
      [para(text)], sec,
      makeRecordingCanvas().canvas.getContext('2d') as CanvasRenderingContext2D,
    );
    let lines = 0;
    for (const page of pages) {
      for (const el of page) {
        const sl = sliceOf(el);
        if (sl) lines += sl.end - sl.start;
        else if (el.type === 'paragraph') lines += 1;
      }
    }
    return lines;
  };

  it('a negative charSpace fits more CJK chars per line than no grid', () => {
    const text = 'あ'.repeat(22);
    const naturalLines = linesOf(text, section(ONE_LINE_PAGE)); // 10/line → 3 lines
    const tightLines = linesOf(text, section({ ...charGrid(-8192), ...ONE_LINE_PAGE })); // 11/line → 2 lines
    expect(naturalLines).toBe(3);
    expect(tightLines).toBe(2);
    expect(tightLines).toBeLessThan(naturalLines);
  });

  it('has NO effect on Latin-only text (cells not applied) — line count unchanged', () => {
    // Space-separated Latin words wrap by word; the grid must not change that.
    const text = 'ab '.repeat(30).trim();
    expect(linesOf(text, section({ ...charGrid(-8192), ...ONE_LINE_PAGE })))
      .toBe(linesOf(text, section(ONE_LINE_PAGE)));
  });

  it('type="lines" does not change CJK line count (char grid inactive)', () => {
    const text = 'あ'.repeat(22);
    const lineGrid = section({ docGridType: 'lines', docGridLinePitch: 20, docGridCharSpace: -8192, ...ONE_LINE_PAGE });
    expect(linesOf(text, lineGrid)).toBe(linesOf(text, section(ONE_LINE_PAGE)));
  });
});
