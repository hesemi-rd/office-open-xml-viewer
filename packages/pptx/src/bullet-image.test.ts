import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { renderTextBody, getCachedBitmap, dropImageBitmapCache } from './renderer.js';
import type { TextBody, Paragraph, BlipBullet } from './types';
import type { TextRunData } from '@silurus/ooxml-core';

/**
 * Picture bullets (`<a:buBlip>`, ECMA-376 §21.1.2.4.2) are drawn inside the
 * synchronous text-body layout. The renderer warms the bitmap cache up front
 * (renderSlide's prefetch pass), then the draw reads the settled bitmap via
 * `peekCachedBitmap` and paints it with `ctx.drawImage`. These tests drive
 * `renderTextBody` directly against a mock 2D context that records `drawImage`,
 * mirroring the mock-ctx approach in text-highlight.test.ts / tabular-text.test.ts.
 *
 * EMU_PER_PT = 12700 and emuToPx(emu, scale) = emu * scale, so scale = 1/12700
 * makes "1pt → 1px": a 20pt run yields a 20px bullet box (× buSzPct).
 */
const SCALE = 1 / 12700;

// A sentinel ImageBitmap the stubbed createImageBitmap returns, so we can assert
// the exact object reaches drawImage. Intentionally NON-square (16×8, ratio 2:1)
// so the tests catch a forced-square draw: PowerPoint scales a picture bullet to
// the text height while preserving the bitmap's aspect ratio (§21.1.2.4.2).
const SENTINEL_W = 16;
const SENTINEL_H = 8;
const SENTINEL_RATIO = SENTINEL_W / SENTINEL_H;
const SENTINEL = {
  width: SENTINEL_W,
  height: SENTINEL_H,
  close: () => {},
} as unknown as ImageBitmap;

function mockCtx() {
  const draws: Array<{ img: unknown; x: number; y: number; w: number; h: number }> = [];
  const texts: Array<{ text: string; x: number; y: number }> = [];
  let fillStyle = '';
  let font = '';
  let direction: CanvasDirection = 'ltr';
  const ctx = {
    get fillStyle() {
      return fillStyle;
    },
    set fillStyle(v: string) {
      fillStyle = v;
    },
    get font() {
      return font;
    },
    set font(v: string) {
      font = v;
    },
    get direction() {
      return direction;
    },
    set direction(v: CanvasDirection) {
      direction = v;
    },
    // Every glyph advances 10px so the line has a measurable, predictable width.
    measureText: (s: string) => ({
      width: s.length * 10,
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
    }),
    fillText: (t: string, x: number, y: number) => texts.push({ text: t, x, y }),
    fillRect: () => {},
    drawImage: (img: unknown, x: number, y: number, w: number, h: number) =>
      draws.push({ img, x, y, w, h }),
    save: () => {},
    restore: () => {},
    translate: () => {},
    rotate: () => {},
    scale: () => {},
    beginPath: () => {},
    moveTo: () => {},
    lineTo: () => {},
    stroke: () => {},
    clip: () => {},
    rect: () => {},
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, draws, texts };
}

function run(text: string, over: Partial<TextRunData> = {}): TextRunData {
  return {
    type: 'text',
    text,
    bold: null,
    italic: null,
    underline: false,
    strikethrough: false,
    fontSize: 20,
    color: '000000',
    fontFamily: 'Arial',
    ...over,
  };
}

function bodyWithBullet(bullet: Paragraph['bullet'], runs: TextRunData[] = [run('Item')]): TextBody {
  const para: Paragraph = {
    alignment: 'l',
    marL: 457200, // a normal hanging-indent list metric so the bullet has a gutter
    marR: 0,
    indent: -457200,
    spaceBefore: null,
    spaceAfter: null,
    spaceLine: null,
    lvl: 0,
    bullet,
    defFontSize: null,
    defColor: null,
    defBold: null,
    defItalic: null,
    defFontFamily: null,
    tabStops: [],
    eaLnBrk: true,
    runs,
  } as Paragraph;
  return {
    verticalAnchor: 't',
    paragraphs: [para],
    defaultFontSize: 20,
    defaultBold: null,
    defaultItalic: null,
    lIns: 91440,
    rIns: 91440,
    tIns: 45720,
    bIns: 45720,
    wrap: 'square',
    vert: 'horz',
    autoFit: 'none',
  };
}

// A picture-bullet variant. Cast through the PPTX Bullet union (the parser emits
// `type: "blip"`; the statically-narrower core Bullet doesn't list it).
function blipBullet(over: Partial<BlipBullet> = {}): Paragraph['bullet'] {
  const b: BlipBullet = {
    type: 'blip',
    imagePath: 'ppt/media/bullet-img.png',
    mimeType: 'image/png',
    sizePct: null,
    ...over,
  };
  return b as unknown as Paragraph['bullet'];
}

type FetchImageFn = (path: string, mime: string) => Promise<Blob>;

describe('renderTextBody — picture bullet (buBlip) draws the bitmap', () => {
  let fetchImage: FetchImageFn;

  beforeEach(() => {
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (_blob: Blob) => SENTINEL),
    );
    fetchImage = vi.fn(
      async (_path: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    ) as FetchImageFn;
  });
  afterEach(() => {
    dropImageBitmapCache(fetchImage);
    vi.unstubAllGlobals();
  });

  it('draws the warmed bullet bitmap sized to the text height with the bitmap aspect ratio, at the bullet gutter', async () => {
    const path = 'ppt/media/bullet-warmed.png';
    // Warm the cache the way renderSlide's prefetch pass does, then await it so
    // the settled bitmap is visible to the synchronous draw.
    await getCachedBitmap(path, 'image/png', fetchImage);

    const { ctx, draws } = mockCtx();
    renderTextBody(
      ctx,
      bodyWithBullet(blipBullet({ imagePath: path })),
      0, 0, 4000, 2000,
      SCALE,
      null, 0, false, false, '#000000', 1,
      { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
      undefined,
      false,
      fetchImage,
    );

    expect(draws).toHaveLength(1);
    const d = draws[0];
    expect(d.img).toBe(SENTINEL);
    // Height = 20pt run × scale(1/12700) × 12700 = 20px (default buSzPct = 100%).
    expect(d.h).toBeCloseTo(20, 6);
    // Width preserves the bitmap aspect ratio (16:8 = 2:1), not forced square.
    expect(d.w).toBeCloseTo(20 * SENTINEL_RATIO, 6);
    expect(d.w).toBeCloseTo(d.h * SENTINEL_RATIO, 6);
  });

  it('scales the bullet by buSzPct (§21.1.2.4.3)', async () => {
    const path = 'ppt/media/bullet-sized.png';
    await getCachedBitmap(path, 'image/png', fetchImage);

    const { ctx, draws } = mockCtx();
    renderTextBody(
      ctx,
      bodyWithBullet(blipBullet({ imagePath: path, sizePct: 50 })),
      0, 0, 4000, 2000,
      SCALE,
      null, 0, false, false, '#000000', 1,
      { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
      undefined,
      false,
      fetchImage,
    );

    expect(draws).toHaveLength(1);
    // Height = 20px text × 50% = 10px; width preserves the 2:1 aspect ratio.
    expect(draws[0].h).toBeCloseTo(10, 6);
    expect(draws[0].w).toBeCloseTo(10 * SENTINEL_RATIO, 6);
  });

  it('draws nothing (no throw) when the bullet image is not yet decoded', () => {
    // No getCachedBitmap warm-up → peekCachedBitmap returns undefined.
    const { ctx, draws } = mockCtx();
    expect(() =>
      renderTextBody(
        ctx,
        bodyWithBullet(blipBullet({ imagePath: 'ppt/media/cold.png' })),
        0, 0, 4000, 2000,
        SCALE,
        null, 0, false, false, '#000000', 1,
        { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
        undefined,
        false,
        fetchImage,
      ),
    ).not.toThrow();
    expect(draws).toHaveLength(0);
  });

  it('does not draw a picture bullet on an empty paragraph', async () => {
    const path = 'ppt/media/bullet-empty.png';
    await getCachedBitmap(path, 'image/png', fetchImage);

    const { ctx, draws } = mockCtx();
    // Empty paragraph (no runs) — PowerPoint draws no marker.
    renderTextBody(
      ctx,
      bodyWithBullet(blipBullet({ imagePath: path }), []),
      0, 0, 4000, 2000,
      SCALE,
      null, 0, false, false, '#000000', 1,
      { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
      undefined,
      false,
      fetchImage,
    );

    expect(draws).toHaveLength(0);
  });

  it('does not call drawImage for a char bullet (regression guard)', async () => {
    // A char bullet must still go through fillText, never drawImage.
    const { ctx, draws } = mockCtx();
    renderTextBody(
      ctx,
      bodyWithBullet({ type: 'char', char: '•', color: null, sizePct: null, fontFamily: 'Arial' }),
      0, 0, 4000, 2000,
      SCALE,
      null, 0, false, false, '#000000', 1,
      { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
      undefined,
      false,
      fetchImage,
    );
    expect(draws).toHaveLength(0);
  });

  it('first-line text does not overlap the picture bullet (hanging indent)', async () => {
    // Regression guard for the first-line hanging-indent overlap. A picture
    // bullet occupies the gutter exactly like a char/autoNum marker, so the
    // first line's hanging indent MUST be suppressed. Before the fix `hasBullet`
    // excluded `blip`, so the first line started at the bullet's x and rendered
    // ON TOP of the image (same class as docx PR #476). The mock records fillText
    // x, which the other tests don't, so this is the case they couldn't catch.
    const path = 'ppt/media/bullet-hang.png';
    await getCachedBitmap(path, 'image/png', fetchImage);

    const { ctx, draws, texts } = mockCtx();
    renderTextBody(
      ctx,
      bodyWithBullet(blipBullet({ imagePath: path }), [run('Item')]),
      0, 0, 4000, 2000,
      SCALE,
      null, 0, false, false, '#000000', 1,
      { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
      undefined,
      false,
      fetchImage,
    );

    expect(draws).toHaveLength(1); // the bullet image was drawn
    const bulletX = draws[0].x;
    expect(texts.length).toBeGreaterThan(0); // the run text was drawn
    const firstTextX = texts[0].x;
    // First-line text starts to the RIGHT of the bullet by the hanging gap
    // (|indent| = 457200 EMU = 36px at this scale), NOT at the bullet's x.
    // Pre-fix this difference was 0 (overlap) → this assertion fails without it.
    expect(firstTextX).toBeGreaterThan(bulletX);
    expect(firstTextX - bulletX).toBeCloseTo(457200 * SCALE, 4);
  });
});
