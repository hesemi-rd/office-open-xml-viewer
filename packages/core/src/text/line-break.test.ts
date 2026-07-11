import { describe, expect, it } from 'vitest';
import {
  LINE_BREAK_UNICODE_VERSION,
  isUax14NoBreakPair,
  lineBreakClass,
} from './line-break.js';

describe('lineBreakClass', () => {
  it('uses the pinned Unicode 17.0.0 table across its full range', () => {
    expect(LINE_BREAK_UNICODE_VERSION).toBe('17.0.0');
    expect(lineBreakClass(0x0000)).toBe('CM');
    expect(lineBreakClass(0x20000)).toBe('ID');
    expect(lineBreakClass(0x10ffff)).toBe('AL');
  });

  it('classifies the reported alphabetic boundary from real UCD data', () => {
    expect(lineBreakClass(0x003c)).toBe('AL'); // LESS-THAN SIGN
    expect(lineBreakClass(0x0041)).toBe('AL'); // Latin
    expect(lineBreakClass(0x0627)).toBe('AL'); // Arabic ALEF
    expect(lineBreakClass(0x05d0)).toBe('HL'); // Hebrew ALEF
    expect(lineBreakClass(0x0030)).toBe('NU');
  });

  it('applies the default UAX #14 LB1 resolutions', () => {
    expect(lineBreakClass(0x00a7)).toBe('AL'); // raw AI
    expect(lineBreakClass(0x0e01)).toBe('AL'); // raw SA, Lo
    expect(lineBreakClass(0x0e31)).toBe('CM'); // raw SA, Mn
    expect(lineBreakClass(0x3041)).toBe('NS'); // raw CJ
    expect(lineBreakClass(0xd800)).toBe('AL'); // raw SG
    expect(lineBreakClass(0x0378)).toBe('AL'); // raw XX
  });
});

describe('isUax14NoBreakPair', () => {
  const AL = 0x003c;
  const HL = 0x05d0;

  it.each([
    ['AL × AL', AL, 0x0627],
    ['AL × HL', AL, HL],
    ['HL × AL', HL, 0x0041],
    ['HL × HL', HL, 0x05d1],
  ])('implements LB28 for %s', (_label, prev, next) => {
    expect(isUax14NoBreakPair(prev as number, next as number)).toBe(true);
  });

  it.each([
    ['AL × NU', AL, 0x0030],
    ['NU × AL', 0x0030, AL],
    ['NU × NU', 0x0030, 0x0031],
    ['OP × AL', 0x0028, AL],
    ['AL × OP', AL, 0x0028],
    ['CP × AL', 0x0029, AL],
    ['AL × CL', AL, 0x007d],
    ['AL × EX', AL, 0x0021],
    ['AL × IS', AL, 0x002c],
    ['AL × SY', AL, 0x002f],
    ['AL × SP', AL, 0x0020],
    ['SP × AL', 0x0020, AL],
    ['AL × ZW', AL, 0x200b],
    ['ZW × AL', 0x200b, AL],
    ['AL × BA', AL, 0x007c],
    ['BA × AL', 0x007c, AL],
    ['AL × B2', AL, 0x2014],
    ['B2 × AL', 0x2014, AL],
    ['AL × HH', AL, 0x2010],
    ['HH × AL', 0x2010, AL],
    ['AL × HY', AL, 0x002d],
    ['HY × AL', 0x002d, AL],
    ['AL × ID', AL, 0x4e00],
    ['ID × AL', 0x4e00, AL],
    ['AL × CM', AL, 0x0301],
    ['CM × AL', 0x0301, AL],
    ['AL × ZWJ', AL, 0x200d],
    ['ZWJ × AL', 0x200d, AL],
    ['AL × WJ', AL, 0x2060],
    ['WJ × AL', 0x2060, AL],
    ['AL × GL', AL, 0x00a0],
    ['GL × AL', 0x00a0, AL],
    ['AL × QU', AL, 0x0022],
    ['QU × AL', 0x0022, AL],
  ])('returns false for deferred/excluded %s', (_label, prev, next) => {
    expect(isUax14NoBreakPair(prev as number, next as number)).toBe(false);
  });
});
