import { describe, expect, it } from 'vitest';
import {
  divideExactLengthKey,
  exactLengthKeyFromDecimal,
  exactLengthKeyToNumber,
} from './exact-length.js';

describe('exact layout length', () => {
  it('converts finite subnormal and large rationals without overflowing either BigInt operand', () => {
    const smallest = exactLengthKeyFromDecimal('1e-323');
    const largest = exactLengthKeyFromDecimal('1e308');

    expect(smallest).not.toBeNull();
    expect(largest).not.toBeNull();
    expect(exactLengthKeyToNumber(smallest!)).toBe(1e-323);
    expect(exactLengthKeyToNumber(largest!)).toBe(1e308);
  });

  it('rounds a schema-valid large unsigned twips measure exactly like the numeric solver input', () => {
    const lexicalTwips = '2542686831678384';
    const twips = exactLengthKeyFromDecimal(lexicalTwips);

    expect(twips).not.toBeNull();
    const points = divideExactLengthKey(twips!, 20n);
    expect(exactLengthKeyToNumber(points)).toBe(Number(lexicalTwips) / 20);
  });

  it('rounds normal midpoint ties to the even binary64 significand', () => {
    const evenLowerMidpoint = '9007199254740993/9007199254740992';
    const oddLowerMidpoint = '9007199254740995/9007199254740992';

    expect(exactLengthKeyToNumber(evenLowerMidpoint)).toBe(1);
    expect(exactLengthKeyToNumber(oddLowerMidpoint)).toBe(1.0000000000000004);
  });

  it('rounds the minimum-subnormal midpoint to even zero and values above it upward', () => {
    const twiceMinimumSubnormalDenominator = (1n << 1075n).toString();
    const fourTimesMinimumSubnormalDenominator = (1n << 1076n).toString();

    expect(exactLengthKeyToNumber(`1/${twiceMinimumSubnormalDenominator}`)).toBe(0);
    expect(exactLengthKeyToNumber(`3/${fourTimesMinimumSubnormalDenominator}`))
      .toBe(Number.MIN_VALUE);
  });

  it('changes from the maximum finite binary64 value to Infinity at the overflow midpoint', () => {
    const overflowMidpoint = (1n << 1024n) - (1n << 970n);

    expect(exactLengthKeyToNumber(`${overflowMidpoint - 1n}/1`)).toBe(Number.MAX_VALUE);
    expect(exactLengthKeyToNumber(`${overflowMidpoint}/1`)).toBe(Number.POSITIVE_INFINITY);
  });

  it('canonicalizes millions of redundant zeros without constructing a giant BigInt', () => {
    // If leading/trailing zeros were folded into a BigInt or power of ten this
    // would allocate a ~1e6-digit integer and time out; canonicalization first
    // keeps it O(string length).
    const leadingZeros = `${'0'.repeat(1_000_000)}5`;
    const trailingZeros = `5${'0'.repeat(1_000_000)}`;

    expect(exactLengthKeyFromDecimal(leadingZeros)).toBe('5/1');
    // The trailing-zero magnitude has a normalized exponent of 1e6, over budget.
    expect(exactLengthKeyFromDecimal(trailingZeros)).toBeNull();
  });

  it('strips integer leading zeros and maps -0 to 0', () => {
    expect(exactLengthKeyFromDecimal('007')).toBe('7/1');
    expect(exactLengthKeyFromDecimal('00.500')).toBe('1/2');
    expect(exactLengthKeyFromDecimal('-0')).toBe('0/1');
    expect(exactLengthKeyFromDecimal('-0.000')).toBe('0/1');
  });

  it('preserves subnormal underflow magnitudes as exact keys rather than rejecting them', () => {
    const subnormal = exactLengthKeyFromDecimal('1e-323');
    expect(subnormal).not.toBeNull();
    expect(exactLengthKeyToNumber(subnormal!)).toBe(1e-323);

    // Below the smallest subnormal but within the decimal budget: kept exact,
    // and only rounds to 0 at the binary64 boundary rather than being dropped.
    const belowSubnormal = exactLengthKeyFromDecimal('1e-400');
    expect(belowSubnormal).not.toBeNull();
    expect(exactLengthKeyToNumber(belowSubnormal!)).toBe(0);
  });

  it('accepts values at the significant-digit and exponent budget and rejects beyond it', () => {
    expect(exactLengthKeyFromDecimal(`1${'0'.repeat(766)}3`)).not.toBeNull(); // 768 sig digits
    expect(exactLengthKeyFromDecimal(`1${'2'.repeat(768)}`)).toBeNull(); // 769 sig digits
    expect(exactLengthKeyFromDecimal('1e1100')).not.toBeNull();
    expect(exactLengthKeyFromDecimal('1e1101')).toBeNull();
    expect(exactLengthKeyFromDecimal('1e-1100')).not.toBeNull();
    expect(exactLengthKeyFromDecimal('1e-1101')).toBeNull();
  });
});
