import { describe, it, expect } from 'vitest';
import { hexToRgba, relativeLuma, autoContrastColor, applyStroke } from './paint.js';
import type { Stroke } from '../types/common.js';

// hexToRgba is the colour pipeline's exit point: the pptx parser resolves a
// run's a:highlight (§21.1.2.3.4) to either a 6-char opaque hex or, when an
// alpha transform is present, an 8-char RRGGBBAA hex, and the renderer feeds it
// straight to hexToRgba before fillRect. Both shapes must round-trip; this is
// the only conversion the highlight box relies on.
describe('hexToRgba', () => {
  it('converts 6-char hex to opaque rgba', () => {
    expect(hexToRgba('FFFF00')).toBe('rgba(255,255,0,1)');
  });

  it('tolerates a leading # on 6-char hex', () => {
    expect(hexToRgba('#00FF00')).toBe('rgba(0,255,0,1)');
  });

  it('reads alpha from the 8-char RRGGBBAA form', () => {
    // AA = 80 → 128/255 ≈ 0.502 (a translucent marker from <a:alpha>).
    expect(hexToRgba('00FF0080')).toBe(`rgba(0,255,0,${128 / 255})`);
  });

  it('8-char alpha overrides the explicit alpha argument', () => {
    // The trailing AA wins so callers can pass colours uniformly.
    expect(hexToRgba('FF000000', 1)).toBe('rgba(255,0,0,0)');
  });

  it('applies the alpha argument only to 6-char hex', () => {
    expect(hexToRgba('FFFF00', 0.5)).toBe('rgba(255,255,0,0.5)');
  });
});

// relativeLuma is the Rec.601 perceptual luma used by automatic-contrast text
// pickers. It shares hexToRgba's hex normalisation: a leading '#' is tolerated
// and any 8-char alpha byte is ignored.
describe('relativeLuma (Rec.601)', () => {
  it('returns 0 for black and 255 for white', () => {
    expect(relativeLuma('000000')).toBe(0);
    expect(relativeLuma('FFFFFF')).toBeCloseTo(255);
  });

  it('weights the channels 0.299 / 0.587 / 0.114', () => {
    expect(relativeLuma('FF0000')).toBeCloseTo(0.299 * 255);
    expect(relativeLuma('00FF00')).toBeCloseTo(0.587 * 255);
    expect(relativeLuma('0000FF')).toBeCloseTo(0.114 * 255);
  });

  it('tolerates a leading # and ignores the 8-char alpha byte', () => {
    expect(relativeLuma('#00FF00')).toBeCloseTo(0.587 * 255);
    // RRGGBBAA — the AA is ignored, same value as the 6-char form.
    expect(relativeLuma('00FF0080')).toBeCloseTo(0.587 * 255);
  });
});

// autoContrastColor implements `w:color="auto"` (ECMA-376 §17.3.2.6) and the
// generic "pick legible text colour for a background" need shared across
// renderers. The black/white pick is implementation-defined (no normative
// algorithm): luma < 128 ⇒ white text, else black; null background ⇒ black.
describe('autoContrastColor (impl-defined; §17.3.2.6 w:color auto)', () => {
  it('picks white text on a black background (inverse video)', () => {
    expect(autoContrastColor('000000')).toBe('#FFFFFF');
  });

  it('picks black text on a white background', () => {
    expect(autoContrastColor('ffffff')).toBe('#000000');
  });

  it('defaults to black text when there is no background (page white)', () => {
    expect(autoContrastColor(null)).toBe('#000000');
  });

  it('uses Rec.601 luma — mid green is light enough for black text', () => {
    // green #00FF00 has luma 0.587*255 ≈ 150 (> 128) ⇒ black text.
    expect(autoContrastColor('00ff00')).toBe('#000000');
    // dark blue #0000FF has luma 0.114*255 ≈ 29 (< 128) ⇒ white text.
    expect(autoContrastColor('0000ff')).toBe('#FFFFFF');
  });

  it('tolerates a leading # (same normalisation as hexToRgba)', () => {
    expect(autoContrastColor('#000000')).toBe('#FFFFFF');
    expect(autoContrastColor('#ffffff')).toBe('#000000');
  });
});

// applyStroke turns a Stroke's prstDash value (ST_PresetLineDashVal, §20.1.10.49)
// into a setLineDash pattern. A minimal mock records the last setLineDash
// argument so we can assert dashed vs continuous (solid) output.
function makeStrokeMock(): {
  ctx: CanvasRenderingContext2D;
  lastDash: () => number[];
  lastCap: () => CanvasLineCap;
} {
  let dash: number[] = [];
  let cap: CanvasLineCap = 'butt';
  const ctx = {
    strokeStyle: '',
    lineWidth: 0,
    setLineDash(segments: number[]) {
      dash = segments;
    },
    get lineCap() { return cap; },
    set lineCap(value: CanvasLineCap) { cap = value; },
  };
  return {
    ctx: ctx as unknown as CanvasRenderingContext2D,
    lastDash: () => dash,
    lastCap: () => cap,
  };
}

function strokeWith(dashStyle: string): Stroke {
  // width is in EMU; with emuPerPx = 1 the pixel line width lw = max(0.5, 2) = 2.
  return { color: '#000000', width: 2, dashStyle };
}

// The Rust parsers filter out "solid" from prstDash @val at parse time, so
// applyStroke only ever sees the remaining 10 ST_PresetLineDashVal members —
// each of which must produce a NON-empty setLineDash pattern (a table miss
// would silently render them as a solid line, which is the bug this guards).
const PRESET_DASH_STYLES = [
  'dot',
  'dash',
  'lgDash',
  'dashDot',
  'lgDashDot',
  'lgDashDotDot',
  'sysDash',
  'sysDot',
  'sysDashDot',
  'sysDashDotDot',
] as const;

describe('applyStroke dash patterns (§20.1.10.49 ST_PresetLineDashVal)', () => {
  it('renders sysDashDotDot with a non-empty dash pattern (not solid)', () => {
    const { ctx, lastDash } = makeStrokeMock();
    applyStroke(ctx, strokeWith('sysDashDotDot'), 1);
    expect(lastDash().length).toBeGreaterThan(0);
  });

  it('covers every non-solid ST_PresetLineDashVal value with a dash pattern', () => {
    for (const style of PRESET_DASH_STYLES) {
      const { ctx, lastDash } = makeStrokeMock();
      applyStroke(ctx, strokeWith(style), 1);
      expect(
        lastDash().length,
        `dash style "${style}" should not render as solid`,
      ).toBeGreaterThan(0);
    }
  });

  it('renders solid (no dashStyle) and unknown styles as a continuous line', () => {
    for (const stroke of [
      { color: '#000000', width: 2 } as Stroke,
      { color: '#000000', width: 2, dashStyle: 'solid' } as Stroke,
      { color: '#000000', width: 2, dashStyle: 'notARealStyle' } as Stroke,
    ]) {
      const { ctx, lastDash } = makeStrokeMock();
      applyStroke(ctx, stroke, 1);
      expect(lastDash()).toEqual([]);
    }
  });

  it('handles a null stroke as a continuous (empty-dash) line', () => {
    const { ctx, lastDash } = makeStrokeMock();
    applyStroke(ctx, null, 1);
    expect(lastDash()).toEqual([]);
  });

  it('scales the dash pattern by the pixel line width', () => {
    // sysDash relative pattern is [4, 2]; lw = width(2) * emuPerPx(3) = 6.
    const { ctx, lastDash } = makeStrokeMock();
    applyStroke(ctx, strokeWith('sysDash'), 3);
    expect(lastDash()).toEqual([24, 12]);
  });

  it('carries a round VML endcap so a zero-length dash renders as a dot', () => {
    const { ctx, lastDash, lastCap } = makeStrokeMock();
    applyStroke(ctx, { ...strokeWith('0 2'), lineCap: 'round' }, 1);
    expect(lastDash()).toEqual([0, 4]);
    expect(lastCap()).toBe('round');
  });

  it.each([undefined, 'butt' as const])(
    'renders a zero-length VML dash as a fourfold-symmetric dot with %s endcap',
    (lineCap) => {
      const { ctx, lastDash, lastCap } = makeStrokeMock();
      applyStroke(ctx, { ...strokeWith('0 2'), lineCap }, 1);
      expect(lastDash()).toEqual([0, 4]);
      // Canvas paints a zero-length segment with a square cap, but drops it
      // with its `butt` cap. VML requires the zero token to remain visible.
      expect(lastCap()).toBe('square');
    },
  );

  it('keeps a flat cap when only a VML gap is zero-length', () => {
    const { ctx, lastDash, lastCap } = makeStrokeMock();
    applyStroke(ctx, strokeWith('2 0'), 1);
    expect(lastDash()).toEqual([4, 0]);
    expect(lastCap()).toBe('butt');
  });

  it('renders a VML numeric relative dashstyle without falling back to solid', () => {
    const { ctx, lastDash } = makeStrokeMock();
    // VML dashstyle numbers are relative to the stroke width. With lw=2,
    // `1 1` becomes one stroke-width on and one stroke-width off.
    applyStroke(ctx, strokeWith('1 1'), 1);
    expect(lastDash()).toEqual([2, 2]);
  });
});
