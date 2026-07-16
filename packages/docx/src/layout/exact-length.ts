/** A reduced rational point value (`numerator/denominator`). BigInt is kept
 * inside arithmetic only: clone-safe layout inputs retain this canonical text. */
export type ExactLengthKey = string;

type Rational = Readonly<{ numerator: bigint; denominator: bigint }>;

function gcd(left: bigint, right: bigint): bigint {
  let a = left < 0n ? -left : left;
  let b = right < 0n ? -right : right;
  while (b !== 0n) [a, b] = [b, a % b];
  return a === 0n ? 1n : a;
}

function rational(numerator: bigint, denominator: bigint): Rational {
  if (denominator === 0n) throw new RangeError('Exact length denominator must not be zero');
  const sign = denominator < 0n ? -1n : 1n;
  const divisor = gcd(numerator, denominator);
  return Object.freeze({
    numerator: sign * numerator / divisor,
    denominator: sign * denominator / divisor,
  });
}

// Resource budget bounding the BigInt work an exact key may cost. These limits
// are generous supersets of the binary64 finite range (max ~1.8e308, min
// subnormal ~5e-324, ~17 significant decimal digits), so NO in-range authored
// measure is ever refused for being merely large or precise — only pathological
// lexical forms (millions of digits, |exponent| in the thousands) exceed them.
// Over-budget is therefore "identity unknown", not "out of range": the adapter
// keeps deterministic binary64 geometry while returning a null exact key.
const MAX_SIGNIFICANT_DIGITS = 768;
const MAX_DECIMAL_EXPONENT = 1100;

function parseDecimal(value: string): Rational | null {
  const match = /^([+-]?)(?:(\d+)(?:\.(\d*))?|\.(\d+))(?:[eE]([+-]?\d+))?$/.exec(value);
  if (!match) return null;
  const negative = match[1] === '-';
  const intDigits = match[2] ?? '';
  const fractionDigits = match[3] ?? match[4] ?? '';
  const authoredExponent = Number(match[5] ?? '0');
  if (!Number.isSafeInteger(authoredExponent)) return null;
  // The combined digit string with the implied decimal point after intDigits.
  const digits = `${intDigits}${fractionDigits}`;
  // Losslessly canonicalize BEFORE constructing any BigInt: strip leading zeros
  // (value-neutral) and trailing zeros (folded into the decimal scale), and map
  // -0 to 0. Use two linear index scans rather than a `/0*$/` regex, whose
  // anchored star backtracks in O(n^2) on a long non-zero-terminated string; a
  // lexical value carrying millions of redundant zeros must stay O(n) and never
  // build a giant BigInt or power of ten. `48` is the code point of '0'.
  let first = 0;
  while (first < digits.length && digits.charCodeAt(first) === 48) first += 1;
  if (first === digits.length) return rational(0n, 1n);
  let last = digits.length - 1;
  while (last > first && digits.charCodeAt(last) === 48) last -= 1;
  const mantissa = digits.slice(first, last + 1);
  const trailingZeros = digits.length - 1 - last;
  // Decimal scale of the least-significant retained digit.
  const scale = authoredExponent - fractionDigits.length + trailingZeros;
  // Normalized scientific exponent (position of the leading significant digit).
  const normalizedExponent = scale + mantissa.length - 1;
  if (mantissa.length > MAX_SIGNIFICANT_DIGITS
    || Math.abs(normalizedExponent) > MAX_DECIMAL_EXPONENT) {
    return null;
  }
  let numerator = BigInt(mantissa);
  let denominator = 1n;
  if (scale >= 0) numerator *= 10n ** BigInt(scale);
  else denominator = 10n ** BigInt(-scale);
  if (negative) numerator = -numerator;
  return rational(numerator, denominator);
}

function parseKey(value: ExactLengthKey): Rational {
  const match = /^(-?\d+)\/([1-9]\d*)$/.exec(value);
  if (!match) throw new RangeError(`Invalid exact length key: ${value}`);
  return rational(BigInt(match[1]!), BigInt(match[2]!));
}

function key(value: Rational): ExactLengthKey {
  const reduced = rational(value.numerator, value.denominator);
  return `${reduced.numerator}/${reduced.denominator}`;
}

export function exactLengthKeyFromDecimal(value: string): ExactLengthKey | null {
  const parsed = parseDecimal(value);
  return parsed ? key(parsed) : null;
}

export function exactLengthKeyFromNumber(value: number): ExactLengthKey | null {
  if (!Number.isFinite(value) || value < 0) return null;
  const parsed = parseDecimal(value.toString());
  return parsed ? key(parsed) : null;
}

function binaryExponent(numerator: bigint, denominator: bigint): number {
  let exponent = numerator.toString(2).length - denominator.toString(2).length;
  const belowCandidate = exponent >= 0
    ? numerator < denominator << BigInt(exponent)
    : numerator << BigInt(-exponent) < denominator;
  if (belowCandidate) exponent -= 1;
  return exponent;
}

function roundedBinaryQuotient(
  numerator: bigint,
  denominator: bigint,
  binaryShift: number,
): bigint {
  const scaledNumerator = binaryShift >= 0
    ? numerator << BigInt(binaryShift)
    : numerator;
  const scaledDenominator = binaryShift < 0
    ? denominator << BigInt(-binaryShift)
    : denominator;
  const quotient = scaledNumerator / scaledDenominator;
  const remainder = scaledNumerator % scaledDenominator;
  const comparison = remainder * 2n - scaledDenominator;
  return comparison > 0n || (comparison === 0n && quotient % 2n !== 0n)
    ? quotient + 1n
    : quotient;
}

export function exactLengthKeyToNumber(value: ExactLengthKey): number {
  const parsed = parseKey(value);
  if (parsed.numerator === 0n) return 0;
  const numeratorNegative = parsed.numerator < 0n;
  const numerator = numeratorNegative ? -parsed.numerator : parsed.numerator;
  let exponent = binaryExponent(numerator, parsed.denominator);
  let magnitude: number;
  if (exponent < -1022) {
    const significand = roundedBinaryQuotient(numerator, parsed.denominator, 1074);
    magnitude = Number(significand) * Number.MIN_VALUE;
  } else {
    let significand = roundedBinaryQuotient(
      numerator,
      parsed.denominator,
      52 - exponent,
    );
    // Rounding can carry a 53-bit significand into the next exponent. Handling
    // that carry before Number conversion avoids a second, platform-dependent
    // decimal rounding step at the finite/Infinity boundary as well.
    if (significand === 1n << 53n) {
      significand >>= 1n;
      exponent += 1;
    }
    magnitude = exponent > 1023
      ? Number.POSITIVE_INFINITY
      : Number(significand) * 2 ** (exponent - 52);
  }
  return numeratorNegative ? -magnitude : magnitude;
}

export function addExactLengthKeys(left: ExactLengthKey, right: ExactLengthKey): ExactLengthKey {
  const a = parseKey(left);
  const b = parseKey(right);
  return key(rational(
    a.numerator * b.denominator + b.numerator * a.denominator,
    a.denominator * b.denominator,
  ));
}

export function multiplyExactLengthKeys(left: ExactLengthKey, right: ExactLengthKey): ExactLengthKey {
  const a = parseKey(left);
  const b = parseKey(right);
  return key(rational(
    a.numerator * b.numerator,
    a.denominator * b.denominator,
  ));
}

export function subtractExactLengthKeys(left: ExactLengthKey, right: ExactLengthKey): ExactLengthKey {
  const a = parseKey(left);
  const b = parseKey(right);
  return key(rational(
    a.numerator * b.denominator - b.numerator * a.denominator,
    a.denominator * b.denominator,
  ));
}

export function divideExactLengthKey(value: ExactLengthKey, divisor: bigint): ExactLengthKey {
  if (divisor === 0n) throw new RangeError('Exact length divisor must not be zero');
  const parsed = parseKey(value);
  return key(rational(parsed.numerator, parsed.denominator * divisor));
}

export function compareExactLengthKeys(left: ExactLengthKey, right: ExactLengthKey): number {
  const a = parseKey(left);
  const b = parseKey(right);
  const difference = a.numerator * b.denominator - b.numerator * a.denominator;
  return difference < 0n ? -1 : difference > 0n ? 1 : 0;
}
