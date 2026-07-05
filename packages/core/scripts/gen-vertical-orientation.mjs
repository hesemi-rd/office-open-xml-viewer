// Generates packages/core/src/text/vertical-orientation.generated.ts from the
// Unicode Character Database (UCD). Run:
//   node packages/core/scripts/gen-vertical-orientation.mjs
//
// Emits, as a gap-free run-length table, the Vertical_Orientation (vo) property
// for every code point in [0, 0x110000):
//   - VerticalOrientation.txt (UAX #50), including its `@missing` block default
//     (R) for code points not otherwise listed.
//
// UAX #50 (https://www.unicode.org/reports/tr50/) defines vo ∈ {U, R, Tu, Tr}:
//   U  Upright   — same orientation as in the code charts (CJK ideographs, kana,
//                  Hangul, fullwidth forms, …).
//   R  Rotated   — rotated 90° clockwise (Latin, digits, most punctuation). This
//                  is the file-wide default (`@missing: 0000..10FFFF; R`).
//   Tu Transformed, fallback Upright — glyph is substituted with a vertical form
//                  if the font has one, else drawn upright (small kana, 、。, …).
//   Tr Transformed, fallback Rotated — glyph is substituted with a vertical form
//                  if the font has one, else ROTATED (brackets （「」, ー long
//                  vowel mark, quotation marks, …).
//
// Note on unassigned code points: several unassigned ranges default to U (per
// the UAX#50 header — CJK-adjacent blocks, PUA, etc.). The UCD data file already
// materialises those as explicit `... ; U # Cn <reserved>` lines, so building the
// table straight from the data section (plus the `@missing` R default for the
// gaps) reproduces the full property with no special-casing needed here.
//
// The generated file is data straight from the UCD — never hand-edit it.

import { readFile, writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const VERSION = '17.0.0';
const BASE = `https://www.unicode.org/Public/${VERSION}/ucd`;
const SRC = 'VerticalOrientation.txt';
const HERE = dirname(fileURLToPath(import.meta.url));
const LOCAL = join(HERE, SRC);
const OUT = join(HERE, '..', 'src', 'text', 'vertical-orientation.generated.ts');
const MAX_CP = 0x110000;

// Canonical vo value index order. Index is what we store; names map back in
// vertical-orientation.ts.
const VO_NAMES = ['U', 'R', 'Tu', 'Tr'];
const VO_INDEX = Object.fromEntries(VO_NAMES.map((n, i) => [n, i]));

// Prefer the checked-in local copy (so the generator is reproducible offline and
// records the exact bytes we shipped); fall back to fetching the canonical URL.
async function readSource() {
  try {
    return await readFile(LOCAL, 'utf8');
  } catch {
    const url = `${BASE}/${SRC}`;
    const res = await fetch(url);
    if (!res.ok) throw new Error(`GET ${url} -> ${res.status}`);
    return res.text();
  }
}

function voIndex(name) {
  if (!(name in VO_INDEX)) throw new Error(`unknown Vertical_Orientation: ${name}`);
  return VO_INDEX[name];
}

function buildTable(text) {
  // Default everything to the file-wide @missing value (R), then overlay explicit
  // assigned/reserved ranges. Result covers [0, MAX_CP) with no gaps.
  const vo = new Uint8Array(MAX_CP).fill(VO_INDEX.R);

  const apply = (start, end, idx) => {
    for (let cp = start; cp <= end && cp < MAX_CP; cp++) vo[cp] = idx;
  };

  // Honour any `@missing` block defaults in file order (there is one: R).
  for (const raw of text.split('\n')) {
    const missing = raw.match(
      /^#\s*@missing:\s*([0-9A-Fa-f]+)\.\.([0-9A-Fa-f]+)\s*;\s*(\w+)/,
    );
    if (missing) {
      apply(parseInt(missing[1], 16), parseInt(missing[2], 16), voIndex(missing[3]));
    }
  }
  // Overlay explicit ranges. Field 1 is one of U / R / Tu / Tr.
  for (const raw of text.split('\n')) {
    const line = raw.split('#')[0].trim();
    if (!line) continue;
    const m = line.match(/^([0-9A-Fa-f]+)(?:\.\.([0-9A-Fa-f]+))?\s*;\s*(\w+)$/);
    if (!m) continue;
    const start = parseInt(m[1], 16);
    const end = m[2] ? parseInt(m[2], 16) : start;
    apply(start, end, voIndex(m[3]));
  }

  // Run-length compress into gap-free ranges: range i covers [starts[i], starts[i+1]).
  const starts = [0];
  const values = [vo[0]];
  for (let cp = 1; cp < MAX_CP; cp++) {
    if (vo[cp] !== values[values.length - 1]) {
      starts.push(cp);
      values.push(vo[cp]);
    }
  }
  return { starts, values };
}

function fmtArray(nums, perLine = 16) {
  const out = [];
  for (let i = 0; i < nums.length; i += perLine) {
    out.push('  ' + nums.slice(i, i + perLine).join(', ') + ',');
  }
  return out.join('\n');
}

async function main() {
  const text = await readSource();
  // Record the exact UCD file version from its header for the provenance comment.
  const headerVer = text.match(/#\s*VerticalOrientation-([\d.]+)\.txt/)?.[1] ?? VERSION;
  const headerDate = text.match(/#\s*Date:\s*([0-9-]+)/)?.[1] ?? 'unknown';
  const { starts, values } = buildTable(text);

  const body = `// AUTO-GENERATED from the Unicode Character Database (UCD ${headerVer}).
// Source: ${BASE}/${SRC} (VerticalOrientation-${headerVer}.txt, dated ${headerDate}).
// Property: Vertical_Orientation (vo), UAX #50 https://www.unicode.org/reports/tr50/
// DO NOT EDIT — regenerate via packages/core/scripts/gen-vertical-orientation.mjs
/* eslint-disable */

export const UNICODE_VERSION = '${headerVer}';

/** Canonical Vertical_Orientation value order (index into these tables). */
export const VO_NAMES = [
  ${VO_NAMES.map((n) => `'${n}'`).join(', ')},
] as const;

/** Range starts (sorted, gap-free): range i covers [VO_RANGE_STARTS[i], VO_RANGE_STARTS[i+1]). */
export const VO_RANGE_STARTS: number[] = [
${fmtArray(starts)}
];

/** Vertical_Orientation index (into VO_NAMES) for the range beginning at the matching VO_RANGE_STARTS entry. */
export const VO_RANGE_VALUE: number[] = [
${fmtArray(values)}
];
`;

  await writeFile(OUT, body, 'utf8');
  console.log(
    `wrote ${OUT}\n  UCD version: ${headerVer} (${headerDate})\n  vo ranges: ${starts.length}`,
  );
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
