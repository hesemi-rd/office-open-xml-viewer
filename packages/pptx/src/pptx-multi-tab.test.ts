import { describe, it, expect } from 'vitest';
import { renderTextBody } from './renderer.js';
import type { TextBody, Paragraph } from './types';
import type { TextRunData, TabStop } from '@silurus/ooxml-core';

// ECMA-376 §21.1.2.1.x (a:tabLst) + UAX#9 — issue #916: a paragraph line may hold
// MANY tab-delimited cells (`title\tvalue\tpage`), each resolved against the stop
// grid in the reading frame and reordered per UAX#9 (tabs are Bidi_Class S). The
// old single-slot `LayoutLine.tabStop` dropped every cell but the last, bypassed
// the visual reorder, and never rendered a start-tab gap. These tests pin the
// multi-cell (item 1), reading-frame (LTR + RTL), and start-tab-gap (item 3)
// behaviour; the per-cell mixed-direction reorder (item 2) is pinned by
// bidi-line.test.ts.

const FONT_PX = 20;
const SCALE = 1 / 12700; // emuToPx(emu, SCALE) = emu·SCALE; 1pt → 1px

function mockCtx(): {
  ctx: CanvasRenderingContext2D;
  texts: { text: string; x: number; direction: CanvasDirection }[];
} {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  let fillStyle = '';
  let direction: CanvasDirection = 'ltr';
  const px = (): number => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const texts: { text: string; x: number; direction: CanvasDirection }[] = [];
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
    get fillStyle() { return fillStyle; }, set fillStyle(v: string) { fillStyle = v; },
    get direction() { return direction; }, set direction(v: CanvasDirection) { direction = v; },
    get letterSpacing() { return letterSpacing; }, set letterSpacing(v: string) { letterSpacing = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    fillText: (t: string, x: number) => texts.push({ text: t, x, direction }),
    strokeText: () => {},
    fillRect: () => {}, drawImage: () => {}, save: () => {}, restore: () => {},
    translate: () => {}, rotate: () => {}, scale: () => {}, beginPath: () => {},
    moveTo: () => {}, lineTo: () => {}, stroke: () => {}, clip: () => {}, rect: () => {},
    setLineDash: () => {}, closePath: () => {}, arc: () => {},
    strokeStyle: '#000', lineWidth: 1, lineJoin: 'miter' as CanvasLineJoin,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, texts };
}

function run(text: string, color = '000000'): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: FONT_PX, color, fontFamily: 'Serif',
  } as TextRunData;
}

function bodyWithTabs(
  tabStops: TabStop[],
  runs: TextRunData[],
  opts: { rtl?: boolean; algn?: string; marR?: number; indent?: number } = {},
): TextBody {
  const rtl = opts.rtl ?? false;
  const para: Paragraph = {
    alignment: opts.algn ?? (rtl ? 'r' : 'l'),
    marL: 0, marR: opts.marR ?? 0, indent: opts.indent ?? 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null, defBold: null, defItalic: null,
    defFontFamily: null, tabStops, rtl, runs,
  } as unknown as Paragraph;
  return {
    verticalAnchor: 't', paragraphs: [para], defaultFontSize: FONT_PX,
    defaultBold: null, defaultItalic: null,
    lIns: 0, rIns: 0, tIns: 0, bIns: 0,
    wrap: 'square', vert: 'horz', autoFit: 'none',
  } as unknown as TextBody;
}

type RunInfo = { text: string; inShapeX: number; w: number };

function render(body: TextBody, boxW = 600): {
  texts: { text: string; x: number; direction: CanvasDirection }[];
  runs: RunInfo[];
} {
  const { ctx, texts } = mockCtx();
  const runs: RunInfo[] = [];
  renderTextBody(
    ctx, body, 0, 0, boxW, 400, SCALE,
    null, 0, false, false, '#000000', undefined,
    { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
    (r) => runs.push({ text: r.text, inShapeX: r.inShapeX, w: r.w }),
  );
  return { texts, runs };
}

const px = (n: number) => n * 12700;

describe('pptx multi-cell tab lines (issue #916)', () => {
  // Two right stops at 300 and 560 in a 600px box; content title|value|page.
  const STOPS: TabStop[] = [{ pos: px(300), algn: 'r' }, { pos: px(560), algn: 'r' }];

  it('LTR: renders ALL THREE cells (item-1 fix), each ending on its stop', () => {
    const { runs } = render(bodyWithTabs(STOPS, [run('title\tvalue\tpage')]));
    const title = runs.find((r) => r.text === 'title');
    const value = runs.find((r) => r.text === 'value');
    const page = runs.find((r) => r.text === 'page');
    // Previously only the LAST cell ('page') survived; now all three render.
    expect(title, 'title cell drawn').toBeTruthy();
    expect(value, 'value cell drawn (was dropped by the single-slot model)').toBeTruthy();
    expect(page, 'page cell drawn').toBeTruthy();
    // title at the leading edge; value ends on stop 300; page ends on stop 560.
    expect(title!.inShapeX).toBeCloseTo(0, 6);
    expect(value!.inShapeX + value!.w).toBeCloseTo(300, 6);
    expect(page!.inShapeX + page!.w).toBeCloseTo(560, 6);
  });

  it('RTL: mirrors the cells (leading cell rightmost) with each on its mirrored stop', () => {
    const { runs } = render(bodyWithTabs(STOPS, [run('title\tvalue\tpage')], { rtl: true }));
    const title = runs.find((r) => r.text === 'title')!;
    const value = runs.find((r) => r.text === 'value')!;
    const page = runs.find((r) => r.text === 'page')!;
    expect(title && value && page, 'all three cells drawn').toBeTruthy();
    // Reading order reversed: page (last, leftmost) < value < title (leading, rightmost).
    expect(page.inShapeX).toBeLessThan(value.inShapeX);
    expect(value.inShapeX).toBeLessThan(title.inShapeX);
    // right/end cells: TRAILING (left) edge on the mirrored stop (600 − pos).
    expect(value.inShapeX).toBeCloseTo(600 - 300, 6); // 300
    expect(page.inShapeX).toBeCloseTo(600 - 560, 6); // 40
    // leading cell's right edge sits at the leading (right) text edge.
    expect(title.inShapeX + title.w).toBeCloseTo(600, 6);
  });

  it('materialises the START-tab gap so the following cell is not contiguous (item-3 fix)', () => {
    // 'A' then a start(l) tab at 100, then 'B'. Old model collapsed the gap (B at
    // 20, right after A); now B sits at the stop (100).
    const stops: TabStop[] = [{ pos: px(100), algn: 'l' }];
    const { runs } = render(bodyWithTabs(stops, [run('A\tB')]));
    const a = runs.find((r) => r.text === 'A')!;
    const b = runs.find((r) => r.text === 'B')!;
    expect(a.inShapeX).toBeCloseTo(0, 6);
    expect(b.inShapeX).toBeCloseTo(100, 6); // gap rendered; NOT contiguous at 20
  });

  it('keeps a cell ON its absolute stop when the first line carries a positive indent', () => {
    // §21.1.2.1.x: a stop is an ABSOLUTE distance from the leading text-inset
    // edge — a first-line indent shifts where the CONTENT starts (the reading
    // pen), not where the stop sits. The resolver's start pen must include the
    // draw-side first-line indent (textXOffset), or every gap over-shoots by the
    // indent and the cell slides past its stop.
    const stops: TabStop[] = [{ pos: px(300), algn: 'r' }];
    const { runs } = render(
      bodyWithTabs(stops, [run('title\tvalue')], { indent: px(50) }),
    );
    const title = runs.find((r) => r.text === 'title')!;
    const value = runs.find((r) => r.text === 'value')!;
    expect(title.inShapeX).toBeCloseTo(50, 6); // first line starts at the indent
    expect(value.inShapeX + value.w).toBeCloseTo(300, 6); // cell still ENDS on the stop
  });
});
