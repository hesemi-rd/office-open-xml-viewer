// Generates packages/core/src/text/line-break-class.generated.ts from the
// Unicode Character Database (UCD). Run:
//   node packages/core/scripts/gen-line-break-class.mjs
//
// Emits a gap-free run-length table for the UAX #14 Line_Break property after
// applying the default LB1 resolutions. The generated file is data straight
// from the pinned UCD inputs — never hand-edit it.

import { writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const VERSION = '17.0.0';
const BASE = `https://www.unicode.org/Public/${VERSION}/ucd`;
const LINE_BREAK_SOURCE = `${BASE}/LineBreak.txt`;
const GENERAL_CATEGORY_SOURCE = `${BASE}/extracted/DerivedGeneralCategory.txt`;
const OUT = join(
  dirname(fileURLToPath(import.meta.url)),
  '..',
  'src',
  'text',
  'line-break-class.generated.ts',
);
const MAX_CP = 0x110000;

const RAW_CLASS_NAMES = [
  'BK', 'CM', 'CR', 'GL', 'LF', 'NL', 'SP', 'WJ', 'ZW', 'ZWJ',
  'AI', 'AK', 'AL', 'AP', 'AS', 'B2', 'BA', 'BB', 'CB', 'CJ',
  'CL', 'CP', 'EB', 'EM', 'EX', 'H2', 'H3', 'HL', 'HH', 'HY', 'ID',
  'IN', 'IS', 'JL', 'JT', 'JV', 'NS', 'NU', 'OP', 'PO', 'PR',
  'QU', 'RI', 'SA', 'SG', 'SY', 'VF', 'VI', 'XX',
];
const RAW_CLASS_INDEX = Object.fromEntries(RAW_CLASS_NAMES.map((name, index) => [name, index]));
const RESOLVED_AWAY = new Set(['AI', 'CJ', 'SA', 'SG', 'XX']);
const LB_CLASS_NAMES = RAW_CLASS_NAMES.filter((name) => !RESOLVED_AWAY.has(name));
const LB_CLASS_INDEX = Object.fromEntries(LB_CLASS_NAMES.map((name, index) => [name, index]));

async function fetchText(url) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`GET ${url} -> ${response.status}`);
  const text = await response.text();
  const actualVersion = text.match(/^#\s*[^\n]*?-(\d+\.\d+\.\d+)\.txt/m)?.[1];
  if (actualVersion !== VERSION) {
    throw new Error(`${url}: expected UCD ${VERSION}, found ${actualVersion ?? 'unknown'}`);
  }
  return text;
}

function parseRange(text) {
  const match = text.match(/^([0-9A-Fa-f]+)(?:\.\.([0-9A-Fa-f]+))?$/);
  if (!match) throw new Error(`invalid code point range: ${text}`);
  const start = Number.parseInt(match[1], 16);
  const end = match[2] ? Number.parseInt(match[2], 16) : start;
  if (start < 0 || end < start || end >= MAX_CP) {
    throw new Error(`out-of-range code point range: ${text}`);
  }
  return { start, end };
}

function applyRange(target, start, end, value) {
  target.fill(value, start, end + 1);
}

function buildRawLineBreakTable(text) {
  const raw = new Uint8Array(MAX_CP).fill(RAW_CLASS_INDEX.XX);

  for (const sourceLine of text.split('\n')) {
    const missing = sourceLine.match(
      /^#\s*@missing:\s*([0-9A-Fa-f]+(?:\.\.[0-9A-Fa-f]+)?)\s*;\s*([A-Z0-9]+)\b/,
    );
    if (!missing) continue;
    const index = RAW_CLASS_INDEX[missing[2]];
    if (index === undefined) throw new Error(`unknown Line_Break class: ${missing[2]}`);
    const { start, end } = parseRange(missing[1]);
    applyRange(raw, start, end, index);
  }

  for (const sourceLine of text.split('\n')) {
    const line = sourceLine.split('#')[0].trim();
    if (!line) continue;
    const match = line.match(/^([0-9A-Fa-f]+(?:\.\.[0-9A-Fa-f]+)?)\s*;\s*([A-Z0-9]+)$/);
    if (!match) throw new Error(`invalid LineBreak.txt row: ${sourceLine}`);
    const index = RAW_CLASS_INDEX[match[2]];
    if (index === undefined) throw new Error(`unknown Line_Break class: ${match[2]}`);
    const { start, end } = parseRange(match[1]);
    applyRange(raw, start, end, index);
  }

  return raw;
}

function buildCombiningMarkTable(text) {
  const mark = new Uint8Array(MAX_CP);
  for (const sourceLine of text.split('\n')) {
    const line = sourceLine.split('#')[0].trim();
    if (!line) continue;
    const match = line.match(/^([0-9A-Fa-f]+(?:\.\.[0-9A-Fa-f]+)?)\s*;\s*([A-Za-z]+)$/);
    if (!match) throw new Error(`invalid DerivedGeneralCategory.txt row: ${sourceLine}`);
    if (match[2] !== 'Mn' && match[2] !== 'Mc') continue;
    const { start, end } = parseRange(match[1]);
    applyRange(mark, start, end, 1);
  }
  return mark;
}

function resolveClass(rawName, isCombiningMark) {
  switch (rawName) {
    case 'AI':
    case 'SG':
    case 'XX':
      return 'AL';
    case 'SA':
      return isCombiningMark ? 'CM' : 'AL';
    case 'CJ':
      return 'NS';
    default:
      return rawName;
  }
}

function buildResolvedRanges(raw, combiningMarks) {
  const starts = [0];
  const firstName = resolveClass(RAW_CLASS_NAMES[raw[0]], combiningMarks[0] === 1);
  const firstIndex = LB_CLASS_INDEX[firstName];
  if (firstIndex === undefined) throw new Error(`unresolved Line_Break class at U+0000: ${firstName}`);
  const classes = [firstIndex];

  for (let cp = 1; cp < MAX_CP; cp++) {
    const name = resolveClass(RAW_CLASS_NAMES[raw[cp]], combiningMarks[cp] === 1);
    const index = LB_CLASS_INDEX[name];
    if (index === undefined) {
      throw new Error(`unresolved Line_Break class at U+${cp.toString(16).toUpperCase()}: ${name}`);
    }
    if (index !== classes[classes.length - 1]) {
      starts.push(cp);
      classes.push(index);
    }
  }

  if (starts[0] !== 0 || starts.length !== classes.length) {
    throw new Error('generated Line_Break ranges are not gap-free');
  }
  return { starts, classes };
}

function formatArray(values, perLine = 16) {
  const lines = [];
  for (let i = 0; i < values.length; i += perLine) {
    lines.push(`  ${values.slice(i, i + perLine).join(', ')},`);
  }
  return lines.join('\n');
}

async function main() {
  const [lineBreakText, generalCategoryText] = await Promise.all([
    fetchText(LINE_BREAK_SOURCE),
    fetchText(GENERAL_CATEGORY_SOURCE),
  ]);
  const raw = buildRawLineBreakTable(lineBreakText);
  const combiningMarks = buildCombiningMarkTable(generalCategoryText);
  const { starts, classes } = buildResolvedRanges(raw, combiningMarks);

  const body = `// AUTO-GENERATED from the Unicode Character Database (UCD ${VERSION}).
// Sources: ${LINE_BREAK_SOURCE}
//          ${GENERAL_CATEGORY_SOURCE}
// Property: Line_Break after the default UAX #14 LB1 resolutions.
// DO NOT EDIT — regenerate via packages/core/scripts/gen-line-break-class.mjs
/* eslint-disable */

export const LINE_BREAK_UNICODE_VERSION = '${VERSION}';

/** Effective Line_Break class names after the default LB1 resolutions. */
export const LB_CLASS_NAMES = [
  ${LB_CLASS_NAMES.map((name) => `'${name}'`).join(', ')},
] as const;

/** Range starts (sorted, gap-free): range i covers [LB_RANGE_STARTS[i], LB_RANGE_STARTS[i+1]). */
export const LB_RANGE_STARTS: number[] = [
${formatArray(starts)}
];

/** Line_Break class index (into LB_CLASS_NAMES) for each range. */
export const LB_RANGE_CLASS: number[] = [
${formatArray(classes)}
];
`;

  await writeFile(OUT, body, 'utf8');
  console.log(`wrote ${OUT}\n  UCD version: ${VERSION}\n  line-break ranges: ${starts.length}`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
