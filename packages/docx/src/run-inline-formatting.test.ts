import { describe, it, expect } from 'vitest';
import { autoContrastColor } from './renderer.js';

describe('autoContrastColor — automatic text color (impl-defined; w:color auto §17.3.2.6)', () => {
  it('picks white text on a black background (inverse video)', () => {
    // "Some text in inverse video." — run shading fill #000000 + w:color="auto".
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
});
