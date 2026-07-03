import { describe, it, expect } from 'vitest';
import { docxUnderlineToDrawingML } from './underline-map.js';

// Exhaustive check that every WordprocessingML ST_Underline (§17.18.99) value
// maps to a DrawingML ST_TextUnderlineType (§20.1.10.82) value that
// core.drawUnderline dispatches on. The DrawingML styles core understands are
// the keys of pptx's PPTX_UNDERLINE_RELATIVE plus sng/dbl/heavy/wavy/wavyHeavy/
// wavyDbl (solid / wave families).
const DRAWINGML_UNDERSTOOD = new Set([
  'sng',
  'dbl',
  'heavy',
  'dotted',
  'dottedHeavy',
  'dash',
  'dashHeavy',
  'dashLong',
  'dashLongHeavy',
  'dotDash',
  'dotDashHeavy',
  'dotDotDash',
  'dotDotDashHeavy',
  'wavy',
  'wavyHeavy',
  'wavyDbl',
]);

describe('docxUnderlineToDrawingML (§17.18.99 → §20.1.10.82)', () => {
  it('maps every ST_Underline value to a core-understood DrawingML style', () => {
    // The full §17.18.99 enum, minus "none" (no underline is drawn at all).
    const all = [
      'single',
      'words',
      'double',
      'thick',
      'dotted',
      'dottedHeavy',
      'dash',
      'dashedHeavy',
      'dashLong',
      'dashLongHeavy',
      'dotDash',
      'dashDotHeavy',
      'dotDotDash',
      'dashDotDotHeavy',
      'wave',
      'wavyHeavy',
      'wavyDouble',
    ];
    for (const v of all) {
      const mapped = docxUnderlineToDrawingML(v);
      expect(DRAWINGML_UNDERSTOOD.has(mapped), `${v} → ${mapped} not understood`).toBe(true);
    }
  });

  it('maps the exact-match values verbatim', () => {
    expect(docxUnderlineToDrawingML('double')).toBe('dbl');
    expect(docxUnderlineToDrawingML('dotted')).toBe('dotted');
    expect(docxUnderlineToDrawingML('dottedHeavy')).toBe('dottedHeavy');
    expect(docxUnderlineToDrawingML('dash')).toBe('dash');
    expect(docxUnderlineToDrawingML('dashLong')).toBe('dashLong');
    expect(docxUnderlineToDrawingML('dashLongHeavy')).toBe('dashLongHeavy');
    expect(docxUnderlineToDrawingML('dotDash')).toBe('dotDash');
    expect(docxUnderlineToDrawingML('dotDotDash')).toBe('dotDotDash');
    expect(docxUnderlineToDrawingML('wavyHeavy')).toBe('wavyHeavy');
  });

  it('normalizes the spelling/order differences to DrawingML names', () => {
    expect(docxUnderlineToDrawingML('single')).toBe('sng');
    expect(docxUnderlineToDrawingML('wave')).toBe('wavy');
    expect(docxUnderlineToDrawingML('wavyDouble')).toBe('wavyDbl');
    // dashed→dash spelling
    expect(docxUnderlineToDrawingML('dashedHeavy')).toBe('dashHeavy');
    // dashDot vs dotDash element ORDER differs between the two enums; both
    // denote the same dot-dash rule.
    expect(docxUnderlineToDrawingML('dashDotHeavy')).toBe('dotDashHeavy');
    expect(docxUnderlineToDrawingML('dashDotDotHeavy')).toBe('dotDotDashHeavy');
  });

  it('maps `thick` to the heavy solid rule (no thick primitive in DrawingML)', () => {
    expect(docxUnderlineToDrawingML('thick')).toBe('heavy');
  });

  it('approximates `words` as a plain single rule (word-gap gaps not modelled)', () => {
    // DrawingML has no words-only underline; drawn as continuous single.
    expect(docxUnderlineToDrawingML('words')).toBe('sng');
  });

  it('falls back to single for unknown / undefined input', () => {
    expect(docxUnderlineToDrawingML(undefined)).toBe('sng');
    expect(docxUnderlineToDrawingML('none')).toBe('sng');
    expect(docxUnderlineToDrawingML('bogus')).toBe('sng');
  });
});
