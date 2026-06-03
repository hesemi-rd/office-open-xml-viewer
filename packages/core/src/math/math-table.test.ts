import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parseMathFont } from './font';
import { parseMathConstants, type MathConstants } from './math-table';

let mc: MathConstants;
beforeAll(() => {
  const url = new URL('../../assets/LatinModernMath.otf', import.meta.url);
  const buf = readFileSync(fileURLToPath(url));
  const font = parseMathFont(
    buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
  );
  mc = parseMathConstants(font);
});

describe('parseMathConstants', () => {
  it('reads exact constants from Latin Modern Math', () => {
    // Values dumped from the bundled font (font units, percentages 0..100).
    expect(mc.scriptPercentScaleDown).toBe(70);
    expect(mc.scriptScriptPercentScaleDown).toBe(50);
    expect(mc.axisHeight).toBe(250);
    expect(mc.fractionRuleThickness).toBe(40);
    expect(mc.fractionNumeratorShiftUp).toBe(394);
    expect(mc.fractionDenominatorShiftDown).toBe(345);
    expect(mc.fractionNumeratorGapMin).toBe(40);
    expect(mc.fractionDenominatorGapMin).toBe(40);
    expect(mc.superscriptShiftUp).toBe(363);
    expect(mc.subscriptShiftDown).toBe(247);
    expect(mc.superscriptBottomMin).toBe(108);
    expect(mc.subscriptTopMax).toBe(344);
    expect(mc.spaceAfterScript).toBe(56);
    expect(mc.upperLimitGapMin).toBe(200);
    expect(mc.lowerLimitGapMin).toBe(167);
    expect(mc.mathLeading).toBe(154);
    expect(mc.radicalVerticalGap).toBe(50);
    expect(mc.radicalRuleThickness).toBe(40);
    expect(mc.radicalExtraAscender).toBe(40);
  });
});
