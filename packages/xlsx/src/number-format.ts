import type { Cell, CellValue, Styles } from './types.js';
import { todaySerial, nowSerial } from './formula.js';


function cellValueText(value: CellValue): string {
  switch (value.type) {
    case 'empty': return '';
    case 'text': return value.text;
    case 'number': return String(value.number);
    case 'bool': return value.bool ? 'TRUE' : 'FALSE';
    case 'error': return value.error;
  }
}

export function formatCellValue(
  cell: Cell,
  styles: Styles,
  cfNumFmt?: { numFmtId: number; formatCode: string | null } | null,
): string {
  // Resolve the effective format once so both the numeric and text paths
  // honour the same precedence: CF dxf numFmt > style numFmt (§18.8.17).
  const xf = styles.cellXfs[cell.styleIndex ?? 0];
  const styleNumFmtId = xf?.numFmtId ?? 0;
  const styleFmt = styles.numFmts?.find(f => f.numFmtId === styleNumFmtId)?.formatCode ?? null;
  const effectiveFmtId = cfNumFmt?.numFmtId ?? styleNumFmtId;
  const effectiveFmt = cfNumFmt?.formatCode ?? styleFmt;

  // Non-numeric cells still need to honour the 4th format section (text).
  // §18.8.30: format sections are positive;negative;zero;text. An empty text
  // section hides the value (Excel's `;;;` trick used for chart-placeholder
  // cells like D3 in the holiday-budget sample), and `@` substitutes the
  // original text. Cells without a 4-section format pass through unchanged.
  if (cell.value.type !== 'number') {
    const text = cellValueText(cell.value);
    return effectiveFmt ? applyTextSection(text, effectiveFmt) : text;
  }

  // Volatile builtins: TODAY()/NOW() cells have a cached `<v>` from the last
  // save, which the viewer would otherwise show as a stale date. Recompute
  // them against the current system clock at render time.
  const num = recomputeVolatile(cell.formula) ?? cell.value.number;
  return applyFormat(num, effectiveFmtId, effectiveFmt);
}

/**
 * Apply the 4th section (text section) of an Excel number format to a text
 * value. ECMA-376 §18.8.30:
 *   - Fewer than 4 sections → text passes through unchanged (Excel default).
 *   - Empty text section   → the value is hidden.
 *   - `@` in the section   → substituted by the original text.
 *   - Quoted / escaped literals are emitted; `[...]` metadata and
 *     `_`/`*` pad pairs are dropped (same conventions as the numeric path).
 */
function applyTextSection(text: string, formatCode: string): string {
  const sections = formatCode.split(';');
  if (sections.length < 4) return text;
  const section = sections[3];
  if (section === '') return '';
  let out = '';
  let i = 0;
  while (i < section.length) {
    const ch = section[i];
    if (ch === '"') {
      i++;
      while (i < section.length && section[i] !== '"') out += section[i++];
      if (i < section.length) i++;
    } else if (ch === '\\') {
      if (i + 1 < section.length) out += section[i + 1];
      i += 2;
    } else if (ch === '[') {
      while (i < section.length && section[i] !== ']') i++;
      if (i < section.length) i++;
    } else if (ch === '@') {
      out += text;
      i++;
    } else if (ch === '_' || ch === '*') {
      i += 2;
    } else {
      out += ch;
      i++;
    }
  }
  return out;
}

/** If `formula` is a volatile builtin (TODAY/NOW), return the current Excel
 *  serial. Tolerates surrounding whitespace and an optional leading `=`. */
function recomputeVolatile(formula: string | undefined): number | null {
  if (!formula) return null;
  const f = formula.trim().replace(/^=/, '').toUpperCase().replace(/\s+/g, '');
  if (f === 'TODAY()') return todaySerial();
  if (f === 'NOW()') return nowSerial();
  return null;
}

// ────────────────────────────────────────────────────────────────
// Date / time formatting  (ECMA-376 §18.8.30)
// ────────────────────────────────────────────────────────────────

// Built-in numFmtId → format code. IDs 14-22 are the ECMA-376 US-English
// built-ins; IDs 27-31 and 50-58 are East-Asian (Japanese) locale built-ins
// that Office ships pre-assigned when the file was authored in ja-JP. The
// spec lists the codes under §18.8.30 Table "Built-in formats" (the
// locale-dependent block is given without format strings but the de-facto
// codes match the ones that Office writes back when opening and re-saving).
const BUILTIN_DATE_FMT: Record<number, string> = {
  14: 'm/d/yyyy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yyyy h:mm',
  // Japanese locale built-ins (East-Asian Office). Values mirror what
  // Excel ja-JP writes for these IDs.
  27: '[$-411]ge.m.d',
  28: '[$-411]ggge"年"m"月"d"日"',
  29: '[$-411]ggge"年"m"月"d"日"',
  30: 'm/d/yy',
  31: 'yyyy"年"m"月"d"日"',
  50: '[$-411]ge.m.d',
  51: '[$-411]ggge"年"m"月"d"日"',
  52: 'yyyy"年"m"月"',
  53: 'm"月"d"日"',
  54: '[$-411]ggge"年"m"月"d"日"',
  55: 'yyyy"年"m"月"',
  56: 'm"月"d"日"',
  57: '[$-411]ge.m.d',
  58: '[$-411]ggge"年"m"月"d"日"',
};

const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];
const WEEKDAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
/** Japanese short weekday names (aaa format code, e.g. "水"). */
const JP_WEEKDAY_SHORT = ['日', '月', '火', '水', '木', '金', '土'];
/** Japanese long weekday names (aaaa format code, e.g. "水曜日"). */
const JP_WEEKDAY_LONG = ['日曜日', '月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日'];

/** Japanese imperial eras, newest-first. First entry whose `start` is
 *  ≤ the target date wins (ECMA-376 §18.8.30 — g/gg/ggg and e/ee codes). */
const JP_ERAS: Array<{ start: Date; abbr: string; short: string; long: string }> = [
  { start: new Date(Date.UTC(2019, 4,  1)), abbr: 'R', short: '令', long: '令和' },
  { start: new Date(Date.UTC(1989, 0,  8)), abbr: 'H', short: '平', long: '平成' },
  { start: new Date(Date.UTC(1926, 11, 25)), abbr: 'S', short: '昭', long: '昭和' },
  { start: new Date(Date.UTC(1912, 6,  30)), abbr: 'T', short: '大', long: '大正' },
  { start: new Date(Date.UTC(1868, 0,  25)), abbr: 'M', short: '明', long: '明治' },
];

function resolveJpEra(date: Date): { abbr: string; short: string; long: string; year: number } {
  for (const era of JP_ERAS) {
    if (date.getTime() >= era.start.getTime()) {
      return {
        abbr: era.abbr,
        short: era.short,
        long: era.long,
        year: date.getUTCFullYear() - era.start.getUTCFullYear() + 1,
      };
    }
  }
  // Pre-Meiji: fall back to Gregorian year, keep Meiji names as a best effort.
  const last = JP_ERAS[JP_ERAS.length - 1];
  return { abbr: last.abbr, short: last.short, long: last.long, year: date.getUTCFullYear() };
}

/** Convert an Excel date serial to a UTC Date (avoids local-timezone off-by-one errors). */
function excelSerialToUTCDate(serial: number): Date {
  return new Date((serial - 25569) * 86400 * 1000);
}

/**
 * Format an Excel date serial using an ECMA-376 format code.
 * Supports: y/yy/yyy/yyyy, m/mm/mmm/mmmm/mmmmm, d/dd/ddd/dddd,
 *           h/hh, m/mm (minutes when after h), s/ss, AM/PM, A/P,
 *           quoted literals, bracket escapes, _ padding, * fill.
 */
function formatExcelDateCode(serial: number, fmtCode: string): string {
  const date = excelSerialToUTCDate(serial);
  const yr = date.getUTCFullYear();
  const mo = date.getUTCMonth() + 1;   // 1-12
  const dy = date.getUTCDate();
  const wd = date.getUTCDay();          // 0=Sun
  const hr = date.getUTCHours();
  const mi = date.getUTCMinutes();
  const sc = date.getUTCSeconds();

  // Take the first section (positive / no-sign section)
  const section = fmtCode.split(';')[0];
  const hasAmPm = /am\/pm|a\/p/i.test(section);
  let era: ReturnType<typeof resolveJpEra> | null = null;
  const getEra = (): ReturnType<typeof resolveJpEra> => era ?? (era = resolveJpEra(date));

  let result = '';
  let i = 0;
  let prevWasHour = false;

  while (i < section.length) {
    const ch = section[i];

    if (ch === '"') {
      // Quoted string literal
      i++;
      while (i < section.length && section[i] !== '"') result += section[i++];
      if (i < section.length) i++;
      prevWasHour = false;

    } else if (ch === '[') {
      // ECMA-376 §18.8.30: `[h]` / `[m]` / `[s]` are elapsed-time tokens that
      // suppress the h < 24 / m < 60 / s < 60 wrap-around and instead render
      // the full duration. Any other bracket content (locale IDs, colours,
      // conditions) is metadata and skipped.
      const end = section.indexOf(']', i);
      const inner = end > i ? section.slice(i + 1, end) : '';
      const elapsed = inner.match(/^([hms])\1*$/i);
      if (elapsed) {
        const kind = elapsed[1].toLowerCase();
        const sign = serial < 0 ? '-' : '';
        const absSec = Math.floor(Math.abs(serial) * 86400);
        let v: number;
        if      (kind === 'h') v = Math.floor(absSec / 3600);
        else if (kind === 'm') v = Math.floor(absSec / 60);
        else                   v = absSec;
        const padded = inner.length >= 2 ? String(v).padStart(inner.length, '0') : String(v);
        result += sign + padded;
        i = end + 1;
        prevWasHour = kind === 'h';
      } else {
        while (i < section.length && section[i] !== ']') i++;
        if (i < section.length) i++;
      }

    } else if (ch === '_') {
      i += 2; // _ followed by a padding character — skip both

    } else if (ch === '*') {
      i += 2; // * followed by fill character — skip both

    } else if (ch === '\\') {
      if (i + 1 < section.length) result += section[i + 1];
      i += 2;
      prevWasHour = false;

    } else if (ch === 'y' || ch === 'Y') {
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'y') { n++; i++; }
      result += n <= 2 ? String(yr).slice(-2) : String(yr).padStart(4, '0');
      prevWasHour = false;

    } else if (ch === 'm' || ch === 'M') {
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'm') { n++; i++; }
      // Determine month vs minutes:
      //   minutes when immediately after h/hh, OR immediately before :s/:ss
      const rest = section.slice(i).replace(/\[[^\]]*\]/g, '');
      const isMinutes = prevWasHour || /^:s/i.test(rest);
      if (isMinutes) {
        result += n >= 2 ? String(mi).padStart(2, '0') : String(mi);
      } else {
        if      (n === 1) result += String(mo);
        else if (n === 2) result += String(mo).padStart(2, '0');
        else if (n === 3) result += MONTH_NAMES[mo - 1].slice(0, 3);
        else if (n === 4) result += MONTH_NAMES[mo - 1];
        else              result += MONTH_NAMES[mo - 1][0]; // mmmmm = first letter
      }
      prevWasHour = false;

    } else if (ch === 'd' || ch === 'D') {
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'd') { n++; i++; }
      if      (n === 1) result += String(dy);
      else if (n === 2) result += String(dy).padStart(2, '0');
      else if (n === 3) result += WEEKDAY_NAMES[wd].slice(0, 3);
      else              result += WEEKDAY_NAMES[wd];
      prevWasHour = false;

    } else if (ch === 'h' || ch === 'H') {
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'h') { n++; i++; }
      const h = hasAmPm ? (hr % 12 || 12) : hr;
      result += n >= 2 ? String(h).padStart(2, '0') : String(h);
      prevWasHour = true;

    } else if (ch === 's' || ch === 'S') {
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 's') { n++; i++; }
      result += n >= 2 ? String(sc).padStart(2, '0') : String(sc);
      prevWasHour = false;

    } else if (ch === 'g' || ch === 'G') {
      // Japanese era name (ECMA-376 §18.8.30 ja locale):
      //   g   → 'R' / 'H' / 'S' / 'T' / 'M'
      //   gg  → '令' / '平' / '昭' / '大' / '明'
      //   ggg → '令和' / '平成' / '昭和' / '大正' / '明治'
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'g') { n++; i++; }
      const e = getEra();
      if      (n === 1) result += e.abbr;
      else if (n === 2) result += e.short;
      else              result += e.long;
      prevWasHour = false;

    } else if (ch === 'e' || ch === 'E') {
      // Japanese era year: `e` → unpadded, `ee` → 2-digit zero-padded.
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'e') { n++; i++; }
      const y = getEra().year;
      result += n >= 2 ? String(y).padStart(2, '0') : String(y);
      prevWasHour = false;

    } else if (ch === 'r' || ch === 'R') {
      // Some Japanese Excel variants expose `r` / `rr` as era-year aliases.
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'r') { n++; i++; }
      const y = getEra().year;
      result += n >= 2 ? String(y).padStart(2, '0') : String(y);
      prevWasHour = false;

    } else if (ch === 'A' || ch === 'a') {
      const upper = section.slice(i).toUpperCase();
      // Japanese weekday format codes (Excel ja locale). `aaaa` = "水曜日",
      // `aaa` = "水". Checked before AM/PM because those are shorter matches
      // and would otherwise swallow the leading 'a'.
      if (upper.startsWith('AAAA')) {
        result += JP_WEEKDAY_LONG[wd]; i += 4;
      } else if (upper.startsWith('AAA')) {
        result += JP_WEEKDAY_SHORT[wd]; i += 3;
      } else if (upper.startsWith('AM/PM')) {
        result += hr < 12 ? 'AM' : 'PM'; i += 5;
      } else if (upper.startsWith('A/P')) {
        result += hr < 12 ? 'A' : 'P'; i += 3;
      } else {
        result += ch; i++;
      }
      prevWasHour = false;

    } else {
      result += ch;
      i++;
      // Separators (:/-. space) don't reset the hour context for m/mm lookahead
      if (ch !== ':' && ch !== '/' && ch !== '-' && ch !== '.' && ch !== ' ') {
        prevWasHour = false;
      }
    }
  }

  return result;
}

/** Returns true if a custom formatCode is a date/time format. */
function isDateFormatCode(code: string): boolean {
  // Elapsed-time brackets `[h]`, `[m]`, `[s]` (ECMA-376 §18.8.30) are themselves
  // time formats, so detect those *before* stripping bracket content below.
  if (/\[[hms]+\]/i.test(code)) return true;
  // Strip quoted literals and bracket content, then look for unambiguous date specifiers.
  // 'y' = year, 'd' = day — both are unambiguous. 'm' alone is ambiguous (month or minutes).
  const stripped = code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '');
  // y / d are unambiguous date specifiers. `aaa+` is the Japanese-locale
  // weekday code and implies a date format even without y/d (e.g. the
  // bare `aaa` custom format).
  return /[yd]/i.test(stripped) || /a{3,}/i.test(stripped);
}

function applyFormat(num: number, numFmtId: number, formatCode: string | null): string {
  // Built-in date/time numFmtIds (ECMA-376 §18.8.30 table)
  const builtinFmt = BUILTIN_DATE_FMT[numFmtId];
  if (builtinFmt) return formatExcelDateCode(num, builtinFmt);
  if (formatCode) {
    if (isDateFormatCode(formatCode)) return formatExcelDateCode(num, formatCode);
    return applyFormatCode(num, formatCode);
  }
  switch (numFmtId) {
    case 0: return String(num);
    case 1: return Math.round(num).toString();
    case 2: return num.toFixed(2);
    case 3: return formatThousands(num, 0);
    case 4: return formatThousands(num, 2);
    case 9: return Math.round(num * 100) + '%';
    case 10: return (num * 100).toFixed(2) + '%';
    case 11: return num.toExponential(2);
    case 37: case 38: return formatThousands(num, 0);
    case 39: case 40: return formatThousands(num, 2);
    case 49: return String(num);
    default: return String(num);
  }
}

function formatThousands(num: number, decimals: number): string {
  return num.toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}

// (formatExcelDate removed; all date formatting now goes through formatExcelDateCode)

function countDecimalPlaces(fmt: string): number {
  const m = fmt.match(/\.([0#]+)/);
  return m ? m[1].length : 0;
}

/**
 * Split a format section into an ordered list of tokens, preserving the exact
 * literal surroundings (quoted strings, backslash escapes, non-placeholder
 * characters like `$`, `€`, `¥` when unquoted) so they can be reassembled
 * around the formatted number. Drops bracket metadata (`[Red]`, `[>0]`),
 * underscore-pad pairs and `*`-fill pairs, per ECMA-376 §18.8.30.
 */
type FmtToken =
  | { kind: 'lit'; text: string }
  | { kind: 'num' }
  | { kind: 'percent' }
  | { kind: 'sci'; expSign: boolean };

function tokenizeNumberFormat(section: string): { tokens: FmtToken[]; numSpec: string } {
  const tokens: FmtToken[] = [];
  let numSpec = '';
  let numPushed = false;
  let sciPushed = false;
  const pushLit = (s: string) => {
    if (!s) return;
    const last = tokens[tokens.length - 1];
    if (last && last.kind === 'lit') last.text += s;
    else tokens.push({ kind: 'lit', text: s });
  };
  const ensureNum = () => {
    if (!numPushed) { tokens.push({ kind: 'num' }); numPushed = true; }
  };

  let i = 0;
  while (i < section.length) {
    const ch = section[i];
    if (ch === '"') {
      i++;
      let s = '';
      while (i < section.length && section[i] !== '"') s += section[i++];
      if (i < section.length) i++;
      pushLit(s);
    } else if (ch === '\\') {
      if (i + 1 < section.length) pushLit(section[i + 1]);
      i += 2;
    } else if (ch === '[') {
      while (i < section.length && section[i] !== ']') i++;
      if (i < section.length) i++;
    } else if (ch === '_') {
      i += 2;
    } else if (ch === '*') {
      i += 2;
    } else if (ch === '#' || ch === '0' || ch === '?' || ch === '.' || ch === ',') {
      ensureNum();
      numSpec += ch;
      i++;
    } else if (ch === '%') {
      tokens.push({ kind: 'percent' });
      i++;
    } else if ((ch === 'E' || ch === 'e') && (section[i + 1] === '+' || section[i + 1] === '-')) {
      if (!sciPushed) {
        tokens.push({ kind: 'sci', expSign: section[i + 1] === '+' });
        sciPushed = true;
      }
      i += 2;
      while (i < section.length && section[i] === '0') i++;
    } else {
      pushLit(ch);
      i++;
    }
  }
  return { tokens, numSpec };
}

function formatNumberSpec(value: number, numSpec: string): string {
  const hasThousands = numSpec.includes(',') && /[#0]/.test(numSpec);
  const dec = countDecimalPlaces(numSpec);
  if (hasThousands) return formatThousands(value, dec);
  if (numSpec.includes('.')) return value.toFixed(dec);
  if (/[#0?]/.test(numSpec)) return Math.round(value).toString();
  return String(value);
}

function applyFormatCode(num: number, formatCode: string): string {
  const sections = formatCode.split(';');
  // Excel number formats have up to 4 sections: positive;negative;zero;text
  // (§18.8.30). Pick the section matching `num`, falling back to the positive
  // section when the target one is absent.
  let section: string;
  // When a dedicated negative section is present, Excel formats the negative
  // value's *magnitude* in it — the minus is conveyed by the section's own
  // literals (e.g. parentheses). `0;(0)` renders -5 as "(5)", not "(-5)".
  let useMagnitude = false;
  if (num > 0) section = sections[0];
  else if (num < 0) {
    if (sections.length > 1) { section = sections[1]; useMagnitude = true; }
    else section = sections[0];
  }
  else section = sections.length > 2 ? sections[2] : sections[0];
  const { tokens, numSpec } = tokenizeNumberFormat(section);
  const hasPercent = tokens.some(t => t.kind === 'percent');
  const sciTok = tokens.find(t => t.kind === 'sci') as Extract<FmtToken, { kind: 'sci' }> | undefined;

  let value = useMagnitude ? Math.abs(num) : num;
  if (hasPercent) value = value * 100;

  let numberText: string;
  let expText = '';
  if (sciTok) {
    const dec = countDecimalPlaces(numSpec);
    const [mantissa, exp] = value.toExponential(dec).split('e');
    numberText = mantissa;
    const e = parseInt(exp, 10);
    const sign = e < 0 ? '-' : (sciTok.expSign ? '+' : '');
    expText = sign + String(Math.abs(e)).padStart(2, '0');
  } else {
    numberText = formatNumberSpec(value, numSpec);
  }

  let result = '';
  let numberEmitted = false;
  for (const t of tokens) {
    if (t.kind === 'lit') result += t.text;
    else if (t.kind === 'percent') result += '%';
    else if (t.kind === 'num') { result += numberText; numberEmitted = true; }
    else if (t.kind === 'sci') result += 'E' + expText;
  }
  if (!numberEmitted && (numSpec.length > 0 || sciTok)) result += numberText;
  return result;
}
