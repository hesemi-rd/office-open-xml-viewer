/**
 * Headless geometry probe for WordArt text warp (ECMA-376 §20.1.9.19,
 * `<a:prstTxWarp>`).
 *
 * Canvas 2D cannot deform glyph outlines, so the pptx renderer approximates a
 * text warp PER GLYPH: it lays the run out flat, then places each glyph against
 * the preset's envelope (a top + bottom boundary curve). This probe renders a
 * synthetic in-memory Presentation (no .pptx, no WASM) carrying one text-bearing
 * shape, and MEASURES the resulting ink to confirm the glyphs actually follow
 * the envelope rather than sitting on a flat baseline.
 *
 * Two assertions, both purely geometric (no eyeballing):
 *   - textArchUp: the ink forms an up-arch — the topmost inked row at the
 *     horizontal CENTRE of the text is higher (smaller y) than at the left/right
 *     ends. A flat baseline would have the centre no higher than the ends.
 *   - textPlain (control): the ink stays flat — the topmost inked row is roughly
 *     constant across the width. This guards against the warp path accidentally
 *     bending un-warpable ("identity") text.
 *
 * CI-safe: gated on skia-canvas via the shared test helper (skip locally, fail
 * under OOXML_REQUIRE_SKIA=1).
 */
import { describe, it, expect } from 'vitest';
import type { Presentation, Slide, ShapeElement, TextBody, Paragraph } from '@silurus/ooxml-pptx';
import { importForTests, loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas } = (skia ?? {}) as Skia;

const renderMod = await importForTests(() => import('./render.ts'), './render.ts');
const { renderSlideNode } = (renderMod ?? {}) as typeof import('./render.ts');

const EMU_PER_PX = 9525; // 96 dpi
const px = (n: number) => Math.round(n * EMU_PER_PX);

function warpShape(preset: string): ShapeElement {
  const para: Paragraph = {
    alignment: 'ctr',
    marL: 0,
    marR: 0,
    indent: 0,
    spaceBefore: null,
    spaceAfter: null,
    spaceLine: null,
    lvl: 0,
    bullet: { type: 'none' },
    defFontSize: null,
    defColor: null,
    defBold: null,
    defItalic: null,
    defFontFamily: null,
    tabStops: [],
    eaLnBrk: true,
    runs: [
      {
        type: 'text',
        text: 'WARP',
        bold: true,
        italic: null,
        underline: false,
        strikethrough: false,
        fontSize: 40,
        color: '000000',
        fontFamily: 'Arial',
      } as Paragraph['runs'][number],
    ],
  };
  const textBody: TextBody = {
    verticalAnchor: 'ctr',
    paragraphs: [para],
    defaultFontSize: 40,
    defaultBold: null,
    defaultItalic: null,
    lIns: 0,
    rIns: 0,
    tIns: 0,
    bIns: 0,
    wrap: 'none',
    vert: 'horz',
    autoFit: 'none',
    textWarp: { preset },
  } as TextBody;
  return {
    type: 'shape',
    x: px(40),
    y: px(40),
    width: px(320),
    height: px(160),
    rotation: 0,
    flipH: false,
    flipV: false,
    geometry: 'rect',
    fill: null,
    stroke: null,
    textBody,
    defaultTextColor: null,
    custGeom: null,
    adj: null,
    adj2: null,
    adj3: null,
    adj4: null,
    adj5: null,
    adj6: null,
    adj7: null,
    adj8: null,
    shadow: null,
  };
}

function buildSlide(preset: string): Presentation {
  const slide: Slide = {
    index: 0,
    slideNumber: 1,
    background: null,
    elements: [warpShape(preset)],
  };
  return {
    slideWidth: px(400),
    slideHeight: px(240),
    slides: [slide],
    defaultTextColor: null,
    majorFont: null,
    minorFont: null,
  };
}

async function renderInk(preset: string): Promise<{
  topRow: (col: number) => number | null;
  bottomRow: (col: number) => number | null;
  width: number;
  height: number;
}> {
  const width = 400;
  const height = 240;
  const dpr = 1;
  const canvas = new Canvas(width * dpr, height * dpr);
  await renderSlideNode(canvas as unknown as Parameters<typeof renderSlideNode>[0], buildSlide(preset), 0, {
    width,
    dpr,
  });
  const ctx = canvas.getContext('2d');
  const img = ctx.getImageData(0, 0, width * dpr, height * dpr);
  const data = img.data as unknown as Uint8ClampedArray;
  const W = width * dpr;
  const H = height * dpr;
  const inked = (col: number, y: number): boolean => {
    const i = (y * W + col) * 4;
    // Dark ink on the default (transparent/white) background: low luminance
    // AND non-transparent. skia clears to transparent; text is opaque black.
    const a = data[i + 3];
    const lum = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
    return a > 40 && lum < 128;
  };
  // For a given column, the first (topmost) row whose pixel is inked (dark).
  const topRow = (col: number): number | null => {
    for (let y = 0; y < H; y++) if (inked(col, y)) return y;
    return null;
  };
  // …and the last (bottommost) inked row.
  const bottomRow = (col: number): number | null => {
    for (let y = H - 1; y >= 0; y--) if (inked(col, y)) return y;
    return null;
  };
  return { topRow, bottomRow, width: W, height: H };
}

/** Min top-inked row across a column band [c0, c1). */
function minTopInBand(topRow: (c: number) => number | null, c0: number, c1: number): number | null {
  let best: number | null = null;
  for (let c = c0; c < c1; c++) {
    const t = topRow(c);
    if (t != null && (best == null || t < best)) best = t;
  }
  return best;
}

describe.skipIf(!skia)('node WordArt text-warp geometry (prstTxWarp)', () => {
  it('textArchUp bends the ink into an up-arch (centre higher than the ends)', async () => {
    const { topRow, width } = await renderInk('textArchUp');
    // Bands: left third, centre third, right third of the rendered width.
    const third = Math.floor(width / 3);
    const leftTop = minTopInBand(topRow, Math.floor(width * 0.18), Math.floor(width * 0.32));
    const centreTop = minTopInBand(topRow, third, 2 * third);
    const rightTop = minTopInBand(topRow, Math.floor(width * 0.68), Math.floor(width * 0.82));

    // There must be ink in all three bands.
    expect(leftTop).not.toBeNull();
    expect(centreTop).not.toBeNull();
    expect(rightTop).not.toBeNull();

    // The apex (centre) sits meaningfully HIGHER (smaller y) than either end.
    // A flat baseline would make these near-equal; require a real arch rise.
    expect(centreTop!).toBeLessThan(leftTop! - 6);
    expect(centreTop!).toBeLessThan(rightTop! - 6);
  });

  it('textPlain leaves the ink flat (control)', async () => {
    const { topRow, width } = await renderInk('textPlain');
    const third = Math.floor(width / 3);
    const leftTop = minTopInBand(topRow, Math.floor(width * 0.18), Math.floor(width * 0.32));
    const centreTop = minTopInBand(topRow, third, 2 * third);
    const rightTop = minTopInBand(topRow, Math.floor(width * 0.68), Math.floor(width * 0.82));
    expect(leftTop).not.toBeNull();
    expect(centreTop).not.toBeNull();
    expect(rightTop).not.toBeNull();
    // Flat: the three bands' top rows agree within a small tolerance (glyph
    // ascenders vary a little, but there is no arch rise).
    const spread = Math.max(leftTop!, centreTop!, rightTop!) - Math.min(leftTop!, centreTop!, rightTop!);
    expect(spread).toBeLessThan(12);
  });

  it('textInflate stretches glyph ink to fill the envelope gap (G1 guard)', async () => {
    // PowerPoint stretches the flat text's ink rectangle to the shape box
    // before warping — verified against PowerPoint's PDF of the warp fixture
    // (ink height 1.48in in a 1.5in-tall shape). The shape here is 160px tall
    // and textInflate's envelope spans the full box height at its centre, so
    // the centre glyphs' ink must span well over half the box. Before the fix
    // glyphs kept their natural ~38px size (vScale collapsed to ≈1).
    const { topRow, bottomRow, width } = await renderInk('textInflate');
    const boxHpx = 160; // shape height in px at this render scale
    let maxSpan = 0;
    for (let c = Math.floor(width * 0.35); c < Math.floor(width * 0.65); c++) {
      const t = topRow(c);
      const b = bottomRow(c);
      if (t != null && b != null) maxSpan = Math.max(maxSpan, b - t + 1);
    }
    expect(maxSpan).toBeGreaterThan(boxHpx * 0.5);
  });
});
