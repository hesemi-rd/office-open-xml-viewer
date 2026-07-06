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
 * Assertions, all purely geometric (no eyeballing):
 *   - textArchUp: Follow Path (issue #846) — the word keeps its NATURAL width
 *     and follows the arch from the path start (stAng = 180°, the LEFT end)
 *     for only its own arc length. So the ink occupies a compact LEADING
 *     segment of the arch and climbs monotonically toward the apex; the right
 *     side of the box stays empty. (An earlier revision asserted the apex at
 *     the horizontal centre — that presumed the pre-#846 behaviour of
 *     stretching the run across the ENTIRE path, which PowerPoint's PDF of the
 *     warp fixture disproves: measured arch ink is 1.75in ≈ natural width, not
 *     the full 5.4-6.1in path.)
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
  it('textArchUp follows the path at natural width from the start (Follow Path, #846)', async () => {
    // Shape box: x=40, w=320 (canvas px). The arch baseline is the top half of
    // the box ellipse, running clockwise from stAng=180° (LEFT end, mid-height)
    // over the apex (270°) to the right end. Follow Path lays "WARP" along only
    // its natural arc length from that start, so the ink is a compact leading
    // segment climbing the LEFT side of the arch (measured locally: ink columns
    // ≈[30,166] of the 400px canvas, top profile 112→44). The pre-#846 renderer
    // stretched the run across the whole 180° arc (ink ≈[40,360], symmetric
    // about the apex) — each assertion below fails against that behaviour.
    const { topRow, width } = await renderInk('textArchUp');
    const shapeX = 40;
    const shapeW = 320;

    // Locate the inked column range.
    let minC: number | null = null;
    let maxC: number | null = null;
    for (let c = 0; c < width; c++) {
      if (topRow(c) != null) {
        if (minC == null) minC = c;
        maxC = c;
      }
    }
    expect(minC).not.toBeNull();
    expect(maxC).not.toBeNull();

    // Natural width, not full-path stretch: the ink span stays well under the
    // box width (a full-arch distribution spans ~the whole 320px).
    expect(maxC! - minC!).toBeLessThan(shapeW * 0.75);

    // The word STARTS at the path start — the left end of the arch. (Glyph
    // centring lets the first glyph overhang slightly left of the box edge.)
    expect(minC!).toBeLessThan(shapeX + shapeW * 0.25);

    // …and does NOT wrap around toward the far side: the right quarter of the
    // box is empty. Before #846 the run reached the arch's right end.
    const rightQuarter = minTopInBand(
      topRow,
      shapeX + Math.floor(shapeW * 0.75),
      shapeX + shapeW,
    );
    expect(rightQuarter).toBeNull();

    // Within its span the ink CLIMBS toward the apex: the trailing end of the
    // word sits meaningfully higher (smaller y) than its start at mid-height.
    // The pre-#846 symmetric distribution had start and end at the same height.
    const span = maxC! - minC!;
    const startTop = minTopInBand(topRow, minC!, minC! + Math.max(1, Math.floor(span * 0.15)));
    const endTop = minTopInBand(topRow, maxC! - Math.max(1, Math.floor(span * 0.15)), maxC! + 1);
    expect(startTop).not.toBeNull();
    expect(endTop).not.toBeNull();
    expect(endTop!).toBeLessThan(startTop! - 20);
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

  it('textWave1 shears glyphs along the wave slope, not a rigid rotation', async () => {
    // ECMA-376 §20.1.9.19: the envelope map's local Jacobian is a rotate+shear,
    // not a plain rotate+scale — on the wave's rising/falling flanks the glyph
    // em-box skews into a parallelogram (vertical strokes track the vertical
    // edge-gap while the baseline follows the slope). PowerPoint's own PDF of the
    // warp fixture shows the "Wave One" glyphs leaning italic-forward.
    //
    // Pixel signature of a SHEAR (as opposed to a rigid rotation): within a
    // narrow vertical slab, the ink's horizontal CENTROID shifts between the
    // upper and lower parts of the glyph band — the top slides forward on an
    // upslope and back on a downslope. A rigid rotation keeps the slab's ink
    // vertically aligned (top and bottom centroids agree). We sample several
    // slabs across the word and require a MEANINGFUL top↔bottom shift that also
    // CHANGES SIGN across the wave (proof the skew tracks the local slope rather
    // than being a single global tilt).
    const width = 400;
    const height = 240;
    const canvas = new Canvas(width, height);
    await renderSlideNode(
      canvas as unknown as Parameters<typeof renderSlideNode>[0],
      buildSlide('textWave1'),
      0,
      { width, dpr: 1 },
    );
    const ctx = canvas.getContext('2d');
    const data = ctx.getImageData(0, 0, width, height).data as unknown as Uint8ClampedArray;
    const inked = (x: number, y: number): boolean => {
      const i = (y * width + x) * 4;
      const a = data[i + 3];
      const lum = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
      return a > 40 && lum < 128;
    };
    // Ink bbox of the whole warped word.
    let minX = width, maxX = -1, minY = height, maxY = -1;
    for (let y = 0; y < height; y++)
      for (let x = 0; x < width; x++)
        if (inked(x, y)) {
          if (x < minX) minX = x;
          if (x > maxX) maxX = x;
          if (y < minY) minY = y;
          if (y > maxY) maxY = y;
        }
    expect(maxX).toBeGreaterThan(minX);
    expect(maxY).toBeGreaterThan(minY);

    // For a vertical slab [xc-6, xc+6], the mean-x of ink in the TOP third vs the
    // BOTTOM third of the glyph band. Positive shift = top sits right of bottom.
    const slabShift = (xc: number): number | null => {
      const bandTop = minY + (maxY - minY) * 0.15;
      const bandBot = minY + (maxY - minY) * 0.85;
      const third = (bandBot - bandTop) / 3;
      let txSum = 0, tN = 0, bxSum = 0, bN = 0;
      for (let x = Math.max(0, xc - 6); x <= Math.min(width - 1, xc + 6); x++) {
        for (let y = Math.floor(bandTop); y < bandTop + third; y++) if (inked(x, y)) { txSum += x; tN++; }
        for (let y = Math.floor(bandBot - third); y < bandBot; y++) if (inked(x, y)) { bxSum += x; bN++; }
      }
      if (tN < 4 || bN < 4) return null;
      return txSum / tN - bxSum / bN;
    };
    const shifts: number[] = [];
    for (let f = 0.15; f <= 0.85; f += 0.05) {
      const s = slabShift(Math.round(minX + (maxX - minX) * f));
      if (s != null) shifts.push(s);
    }
    expect(shifts.length).toBeGreaterThan(4);
    // The lean is substantial somewhere (a rigid rotate keeps |shift| small since
    // the whole glyph turns together; the shear slides the top several px).
    const maxAbs = Math.max(...shifts.map((s) => Math.abs(s)));
    expect(maxAbs).toBeGreaterThan(4);
    // …and it reverses across the wave (upslope vs downslope) — a single global
    // tilt could not do that.
    expect(Math.max(...shifts)).toBeGreaterThan(2);
    expect(Math.min(...shifts)).toBeLessThan(-2);
  });
});
