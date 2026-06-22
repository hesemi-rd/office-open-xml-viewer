import { describe, it, expect } from 'vitest';
import { hexToRgba, relativeLuma, autoContrastColor } from './paint.js';

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
