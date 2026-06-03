import type { MathFont } from './font';

// OpenType MATH table -> the MathConstants subset the layout engine needs.
// Spec: https://learn.microsoft.com/typography/opentype/spec/math#mathconstants-table
//
// MATH header: majorVersion(u16) minorVersion(u16) mathConstantsOffset(Offset16)
//              mathGlyphInfoOffset(Offset16) mathVariantsOffset(Offset16).
// MathConstants layout: the first two fields are Int16 percentages, then two UInt16
// min-heights, then a long run of MathValueRecord (Int16 value + UInt16 deviceOffset,
// 4 bytes each). We read only the Int16 value of each record.
//
// All distances are font units (scale by fontSizePx / unitsPerEm). Percentages are 0..100.
export interface MathConstants {
  scriptPercentScaleDown: number;
  scriptScriptPercentScaleDown: number;
  mathLeading: number;
  axisHeight: number;
  accentBaseHeight: number;
  subscriptShiftDown: number;
  subscriptTopMax: number;
  superscriptShiftUp: number;
  superscriptBottomMin: number;
  subSuperscriptGapMin: number;
  spaceAfterScript: number;
  upperLimitGapMin: number;
  lowerLimitGapMin: number;
  fractionNumeratorShiftUp: number;
  fractionDenominatorShiftDown: number;
  fractionNumeratorGapMin: number;
  fractionRuleThickness: number;
  fractionDenominatorGapMin: number;
  radicalVerticalGap: number;
  radicalRuleThickness: number;
  radicalExtraAscender: number;
}

// Byte offsets within the MathConstants table (verified against Latin Modern Math).
const OFF = {
  scriptPercentScaleDown: 0, // Int16
  scriptScriptPercentScaleDown: 2, // Int16
  // 4: delimitedSubFormulaMinHeight (UInt16), 6: displayOperatorMinHeight (UInt16)
  // MathValueRecords begin at byte 8, 4 bytes each; value = Int16 at record start.
  mathLeading: 8,
  axisHeight: 12,
  accentBaseHeight: 16,
  // 20 flattenedAccentBaseHeight
  subscriptShiftDown: 24,
  subscriptTopMax: 28,
  // 32 subscriptBaselineDropMin
  superscriptShiftUp: 36,
  // 40 superscriptShiftUpCramped
  superscriptBottomMin: 44,
  // 48 superscriptBaselineDropMax
  subSuperscriptGapMin: 52,
  // 56 superscriptBottomMaxWithSubscript
  spaceAfterScript: 60,
  upperLimitGapMin: 64,
  // 68 upperLimitBaselineRiseMin
  lowerLimitGapMin: 72,
  // 76 lowerLimitBaselineDropMin, 80..116 stack/stretch-stack records
  fractionNumeratorShiftUp: 120,
  // 124 fractionNumeratorDisplayStyleShiftUp
  fractionDenominatorShiftDown: 128,
  // 132 fractionDenominatorDisplayStyleShiftDown
  fractionNumeratorGapMin: 136,
  // 140 fractionNumDisplayStyleGapMin
  fractionRuleThickness: 144,
  fractionDenominatorGapMin: 148,
  // 152..184: skewed-fraction / overbar / underbar records
  radicalVerticalGap: 188,
  // 192 radicalDisplayStyleVerticalGap
  radicalRuleThickness: 196,
  radicalExtraAscender: 200,
} as const;

export function parseMathConstants(font: MathFont): MathConstants {
  const mathOff = font.tableOffset('MATH');
  if (mathOff < 0) throw new Error('font has no MATH table');
  const dv = new DataView(font.buffer);
  const c = mathOff + dv.getUint16(mathOff + 4); // start of MathConstants
  const v = (p: number) => dv.getInt16(c + p);
  return {
    scriptPercentScaleDown: v(OFF.scriptPercentScaleDown),
    scriptScriptPercentScaleDown: v(OFF.scriptScriptPercentScaleDown),
    mathLeading: v(OFF.mathLeading),
    axisHeight: v(OFF.axisHeight),
    accentBaseHeight: v(OFF.accentBaseHeight),
    subscriptShiftDown: v(OFF.subscriptShiftDown),
    subscriptTopMax: v(OFF.subscriptTopMax),
    superscriptShiftUp: v(OFF.superscriptShiftUp),
    superscriptBottomMin: v(OFF.superscriptBottomMin),
    subSuperscriptGapMin: v(OFF.subSuperscriptGapMin),
    spaceAfterScript: v(OFF.spaceAfterScript),
    upperLimitGapMin: v(OFF.upperLimitGapMin),
    lowerLimitGapMin: v(OFF.lowerLimitGapMin),
    fractionNumeratorShiftUp: v(OFF.fractionNumeratorShiftUp),
    fractionDenominatorShiftDown: v(OFF.fractionDenominatorShiftDown),
    fractionNumeratorGapMin: v(OFF.fractionNumeratorGapMin),
    fractionRuleThickness: v(OFF.fractionRuleThickness),
    fractionDenominatorGapMin: v(OFF.fractionDenominatorGapMin),
    radicalVerticalGap: v(OFF.radicalVerticalGap),
    radicalRuleThickness: v(OFF.radicalRuleThickness),
    radicalExtraAscender: v(OFF.radicalExtraAscender),
  };
}
