import { describe, it, expect } from 'vitest';
import { splitDigitGroups } from './renderer.js';

// ECMA-376 RTL number handling relies on UAX#9. In an AN-classified Arabic run
// the renderer splits a token into European-digit groups + separators so the
// per-line bidi pass can reorder a date right-to-left (canvas only reorders
// EN-LTR within one fillText). A SINGLE common separator (CS) between two numbers
// joins them into one number (W4) and must NOT split — else a decimal is drawn
// reversed (sample-8: "1234.56" came out "56.1234").
describe('splitDigitGroups — UAX#9 W4 keeps a CS-joined number whole', () => {
  it('keeps a decimal number as one left-to-right group', () => {
    expect(splitDigitGroups('1234.56')).toEqual(['1234.56']);
  });

  it('keeps thousands + decimal grouping whole', () => {
    expect(splitDigitGroups('1,234.56')).toEqual(['1,234.56']);
  });

  it('keeps a time (colon CS) whole', () => {
    expect(splitDigitGroups('12:34')).toEqual(['12:34']);
  });

  it('keeps a slash-joined number whole (CS)', () => {
    expect(splitDigitGroups('3/4')).toEqual(['3/4']);
  });

  it('still splits a hyphen date (ES is not CS for AN digits) for RTL reorder', () => {
    expect(splitDigitGroups('28-02-2026')).toEqual(['28', '-', '02', '-', '2026']);
  });

  it('splits at a CS that is not flanked by digits on both sides', () => {
    // Trailing dot (sentence period after a number): not between two digits.
    expect(splitDigitGroups('1234.')).toEqual(['1234', '.']);
    // Leading separator.
    expect(splitDigitGroups('.5')).toEqual(['.', '5']);
  });

  it('leaves pure non-digit / pure-digit tokens untouched', () => {
    expect(splitDigitGroups('USD')).toEqual(['USD']);
    expect(splitDigitGroups('99')).toEqual(['99']);
    expect(splitDigitGroups('')).toEqual(['']);
  });
});
