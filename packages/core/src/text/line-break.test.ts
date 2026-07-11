import { describe, expect, it } from 'vitest';
import {
  LINE_BREAK_UNICODE_VERSION,
  isEastAsianFWH,
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

describe('isEastAsianFWH', () => {
  it('matches East_Asian_Width Fullwidth / Wide / Halfwidth from real UCD data', () => {
    expect(isEastAsianFWH(0xff08)).toBe(true); // （ Fullwidth
    expect(isEastAsianFWH(0x300c)).toBe(true); // 「 Wide
    expect(isEastAsianFWH(0xff62)).toBe(true); // ｢ Halfwidth
    expect(isEastAsianFWH(0x4e00)).toBe(true); // 一 Wide
    expect(isEastAsianFWH(0x0028)).toBe(false); // ( Narrow
    expect(isEastAsianFWH(0x0041)).toBe(false); // A Narrow
    expect(isEastAsianFWH(0x00c9)).toBe(false); // É Ambiguous
    expect(isEastAsianFWH(0x10ffff)).toBe(false);
  });
});

describe('isUax14NoBreakPair', () => {
  const AL = 0x003c; // <
  const LA = 0x0041; // A (AL)
  const HL = 0x05d0;
  const NU = 0x0030; // 0
  const PR = 0x0024; // $
  const PO = 0x0025; // %
  const ID = 0x4e00; // 一
  const EB = 0x1f466; // 👦
  const EM = 0x1f3fb; // skin-tone modifier
  const OP = 0x0028; // (
  const CP = 0x0029; // )

  it.each([
    ['AL × AL', AL, 0x0627],
    ['AL × HL', AL, HL],
    ['HL × AL', HL, LA],
    ['HL × HL', HL, 0x05d1],
  ])('implements LB28 for %s', (_label, prev, next) => {
    expect(isUax14NoBreakPair(prev as number, next as number)).toBe(true);
  });

  // LB14 is `OP SP* ×`: nothing may break directly after an opening bracket,
  // whatever follows (the SP-spanning instances of the rule are out of scope
  // for a pair predicate — adjacent seams never carry the interior spaces).
  it.each([
    ['OP × AL', OP, LA],
    ['OP × NU', 0x005b, NU],
    ['OP × ID', OP, ID],
    ['OP × OP', OP, OP],
    ['OP × QU', OP, 0x0022],
    ['OP × SP', OP, 0x0020],
    ['OP(F) × AL — LB14 has no East_Asian_Width gate', 0xff08, LA],
  ])('implements LB14 for %s', (_label, prev, next) => {
    expect(isUax14NoBreakPair(prev as number, next as number)).toBe(true);
  });

  it.each([
    ['AL × NU', LA, NU],
    ['HL × NU', HL, NU],
    ['NU × AL', NU, LA],
    ['NU × HL', NU, HL],
  ])('implements LB23 for %s', (_label, prev, next) => {
    expect(isUax14NoBreakPair(prev as number, next as number)).toBe(true);
  });

  it.each([
    ['PR × ID', PR, ID],
    ['PR × EB', PR, EB],
    ['PR × EM', PR, EM],
    ['ID × PO', ID, PO],
    ['EB × PO', EB, PO],
    ['EM × PO', EM, PO],
  ])('implements LB23a for %s', (_label, prev, next) => {
    expect(isUax14NoBreakPair(prev as number, next as number)).toBe(true);
  });

  it.each([
    ['PR × AL', PR, LA],
    ['PR × HL', PR, HL],
    ['PO × AL', PO, LA],
    ['PO × HL', PO, HL],
    ['AL × PR', LA, PR],
    ['HL × PR', HL, PR],
    ['AL × PO', LA, PO],
    ['HL × PO', HL, PO],
  ])('implements LB24 for %s', (_label, prev, next) => {
    expect(isUax14NoBreakPair(prev as number, next as number)).toBe(true);
  });

  // LB25: only the zero-repetition instances of the number regex are decidable
  // from one adjacent pair; the CL/CP × PO/PR lines need the left `NU (SY|IS)*`
  // context, so they stay deferred (see the false table).
  it.each([
    ['NU × NU', NU, 0x0031],
    ['NU × PO', NU, PO],
    ['NU × PR', NU, 0x00a5],
    ['PO × NU', PO, NU],
    ['PR × NU', PR, NU],
    ['HY × NU', 0x002d, NU],
    ['IS × NU', 0x002c, NU],
  ])('implements LB25 (adjacency-provable subset) for %s', (_label, prev, next) => {
    expect(isUax14NoBreakPair(prev as number, next as number)).toBe(true);
  });

  it.each([
    ['AL × OP(Na)', LA, OP],
    ['HL × OP(Na)', HL, OP],
    ['NU × OP(Na)', NU, 0x005b],
    ['CP(Na) × AL', CP, LA],
    ['CP(Na) × HL', 0x005d, HL],
    ['CP(Na) × NU', CP, NU],
  ])('implements LB30 for %s', (_label, prev, next) => {
    expect(isUax14NoBreakPair(prev as number, next as number)).toBe(true);
  });

  it.each([
    ['AL × OP(F)', LA, 0xff08],
    ['AL × OP(W)', LA, 0x300c],
    ['AL × OP(H)', LA, 0xff62],
    ['NU × OP(F)', NU, 0xff08],
    ['HL × OP(W)', HL, 0x300c],
  ])('excludes East Asian opening brackets from LB30: %s', (_label, prev, next) => {
    expect(isUax14NoBreakPair(prev as number, next as number)).toBe(false);
  });

  it.each([
    // LB13 (× CL/EX/IS/SY) is still deferred.
    ['AL × CL', AL, 0x007d],
    ['AL × EX', AL, 0x0021],
    ['AL × IS', AL, 0x002c],
    ['AL × SY', AL, 0x002f],
    ['AL × CL(W)', LA, 0x300d],
    ['CL(W) × AL', 0x300d, LA],
    // LB25 lines that need more than one pair of context stay deferred.
    ['SY × NU', 0x002f, NU],
    ['CL × PO', 0x007d, PO],
    ['CP × PO', CP, PO],
    ['CL × PR', 0x007d, PR],
    ['CP × PR', CP, PR],
    ['NU × HY', NU, 0x002d],
    ['NU × IS', NU, 0x002c],
    // LB29 (IS × AL/HL) is out of scope here.
    ['IS × AL', 0x002c, LA],
    ['IS × HL', 0x002c, HL],
    // LB30b (EB × EM) is out of scope here.
    ['EB × EM', EB, EM],
    // Prefix/postfix pairs no rule joins.
    ['PR × PR', PR, PR],
    ['PO × PO', PO, PO],
    ['PO × PR', PO, PR],
    ['PR × PO', PR, PO],
    ['PO × ID', PO, ID],
    ['ID × PR', ID, PR],
    ['NU × ID', NU, ID],
    ['ID × NU', ID, NU],
    // Everything below was already deferred in v1 and stays that way.
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
    ['AL × ID', AL, ID],
    ['ID × AL', ID, AL],
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
