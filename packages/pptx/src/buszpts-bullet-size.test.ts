import { describe, it, expect } from 'vitest';
import { renderTextBody } from './renderer.js';
import type { TextBody, Paragraph, Bullet } from './types';
import type { TextRunData } from '@silurus/ooxml-core';

// ECMA-376 §21.1.2.4.10 (buSzPts) — an ABSOLUTE point size for the bullet
// marker. Unlike `buSzPct` (§21.1.2.4.9, a percentage of the run size), a
// `buSzPts` marker is sized in points independent of the run, so the marker
// font must be `sizePts` pt (× scale) rather than `sizePct` × the run size.
//
// The mock ctx records the `font` string active at each fillText so the marker
// glyph's pixel size is recoverable independently of the run font.

const FONT_PX = 20;
// emuToPx(emu) = emu·SCALE; PT_TO_EMU = 12700, so with SCALE = 1/12700 a point
// value maps 1:1 to px (20pt run → 20px, 40pt marker → 40px).
const SCALE = 1 / 12700;

interface Fill { text: string; fontPx: number }

function mockCtx(): { ctx: CanvasRenderingContext2D; fills: Fill[] } {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  let fillStyle = '';
  let direction: CanvasDirection = 'ltr';
  const px = (): number => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const fills: Fill[] = [];
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
    get fillStyle() { return fillStyle; }, set fillStyle(v: string) { fillStyle = v; },
    get direction() { return direction; }, set direction(v: CanvasDirection) { direction = v; },
    get letterSpacing() { return letterSpacing; }, set letterSpacing(v: string) { letterSpacing = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    fillText: (t: string) => fills.push({ text: t, fontPx: px() }),
    strokeText: () => {}, fillRect: () => {}, drawImage: () => {}, save: () => {}, restore: () => {},
    translate: () => {}, rotate: () => {}, scale: () => {}, beginPath: () => {},
    moveTo: () => {}, lineTo: () => {}, stroke: () => {}, clip: () => {}, rect: () => {},
    setLineDash: () => {}, closePath: () => {}, arc: () => {},
    strokeStyle: '#000', lineWidth: 1, lineJoin: 'miter' as CanvasLineJoin,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, fills };
}

function run(text: string): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: FONT_PX, color: '000000', fontFamily: 'Serif',
  } as TextRunData;
}

// marL / indent = a hanging gutter so the marker draws in its gutter.
const MARK_EMU = 342900; // 27 px at this SCALE

// A plain Unicode bullet glyph (U+25C6 ◆) with no symbol font, so
// symbolFontToUnicode passes it through unchanged and the marker fill's text
// equals the char.
const BULLET_CHAR = '◆';

function charBullet(opts: { sizePct?: number | null; sizePts?: number }): Bullet {
  return {
    type: 'char', char: BULLET_CHAR, color: null,
    sizePct: opts.sizePct ?? null, sizePts: opts.sizePts,
    fontFamily: null,
  } as unknown as Bullet;
}

function body(bullet: Bullet): TextBody {
  const para: Paragraph = {
    alignment: 'l',
    marL: MARK_EMU, marR: 0, indent: -MARK_EMU,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet, defFontSize: null, defColor: null, defBold: null, defItalic: null,
    defFontFamily: null, tabStops: [], rtl: false, runs: [run('item')],
  } as unknown as Paragraph;
  return {
    verticalAnchor: 't', paragraphs: [para], defaultFontSize: FONT_PX,
    defaultBold: null, defaultItalic: null,
    lIns: 0, rIns: 0, tIns: 0, bIns: 0,
    wrap: 'square', vert: 'horz', autoFit: 'none',
  } as unknown as TextBody;
}

const BOX_W = 600;

function render(b: TextBody): Fill[] {
  const { ctx, fills } = mockCtx();
  renderTextBody(
    ctx, b, 0, 0, BOX_W, 400, SCALE,
    null, 0, false, false, '#000000', undefined,
    { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
    () => {},
  );
  return fills;
}

// The marker fill is the one whose text is the bullet glyph.
function markerPx(fills: Fill[]): number | undefined {
  return fills.find((f) => f.text === BULLET_CHAR)?.fontPx;
}

describe('§21.1.2.4.10 buSzPts sizes the bullet marker in absolute points', () => {
  it('uses the absolute point size (× scale), NOT the run size', () => {
    // buSzPts = 40pt; run = 20pt. The marker must be 40px (40pt at this scale),
    // independent of the run — not the 20px run baseline.
    const px = markerPx(render(body(charBullet({ sizePts: 40 }))));
    expect(px, 'char marker drawn').toBeDefined();
    expect(px).toBeCloseTo(40, 5);
  });

  it('buSzPts wins over a co-present buSzPct', () => {
    // Absolute size (30pt → 30px) takes precedence over the percentage path
    // (200% × 20pt run = 40px), so the marker is 30px, not 40px.
    const px = markerPx(render(body(charBullet({ sizePts: 30, sizePct: 200 }))));
    expect(px).toBeCloseTo(30, 5);
  });

  it('falls back to buSzPct × run size when no buSzPts is present', () => {
    // Control: 200% of the 20pt run = 40px, via the existing percentage path.
    const px = markerPx(render(body(charBullet({ sizePct: 200 }))));
    expect(px).toBeCloseTo(40, 5);
  });

  it('falls back to the run size when neither buSzPts nor buSzPct is present', () => {
    // Control: the marker inherits the 20pt first-run size (§21.1.2.4.13).
    const px = markerPx(render(body(charBullet({}))));
    expect(px).toBeCloseTo(20, 5);
  });
});
