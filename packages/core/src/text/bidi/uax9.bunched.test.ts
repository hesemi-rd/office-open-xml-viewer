import { describe, it, expect, beforeAll } from 'vitest';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { resolveLevels, reorderByLevels, REMOVED } from './uax9/rules.js';
import { bidiClass } from './char-data.js';
import type { BaseDirection, BidiClass } from './types.js';

// Drives the engine with Unicode's class-sequence conformance data (no brackets;
// complements BidiCharacterTest.txt). Each data line lists Bidi_Class names and a
// bitset of paragraph directions; @Levels / @Reorder give the expected output.

const VERSION = '17.0.0';
const URL = `https://www.unicode.org/Public/${VERSION}/ucd/BidiTest.txt`;
const FIX_DIR = join(dirname(fileURLToPath(import.meta.url)), '__fixtures__');
const FIX = join(FIX_DIR, 'BidiTest.txt');

// One representative code point per Bidi_Class (behavior depends only on class).
const REP: Record<string, number> = {
  L: 0x41, R: 0x5d0, AL: 0x627,
  EN: 0x30, ES: 0x2b, ET: 0x24, AN: 0x660, CS: 0x2c, NSM: 0x300, BN: 0xad,
  B: 0x2029, S: 0x09, WS: 0x20, ON: 0x21,
  LRE: 0x202a, LRO: 0x202d, RLE: 0x202b, RLO: 0x202e, PDF: 0x202c,
  LRI: 0x2066, RLI: 0x2067, FSI: 0x2068, PDI: 0x2069,
};
const BASE_BY_BIT: { bit: number; base: BaseDirection }[] = [
  { bit: 1, base: 'auto' },
  { bit: 2, base: 'ltr' },
  { bit: 4, base: 'rtl' },
];

let content: string | null = null;

beforeAll(async () => {
  if (existsSync(FIX)) {
    content = await readFile(FIX, 'utf8');
  } else {
    try {
      const res = await fetch(URL);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      content = await res.text();
      await mkdir(FIX_DIR, { recursive: true });
      await writeFile(FIX, content, 'utf8');
    } catch (e) {
      console.warn(`[bidi bunched] SKIPPED — could not load ${URL}: ${(e as Error).message}`);
      content = null;
    }
  }
}, 60_000);

describe('UAX#9 conformance — BidiTest.txt (class sequences)', () => {
  it('every representative code point round-trips to its Bidi_Class', () => {
    for (const [name, cp] of Object.entries(REP)) {
      expect(bidiClass(cp)).toBe(name as BidiClass);
    }
  });

  it(
    'matches resolved levels and reorder for every line × paragraph direction',
    (ctx) => {
      if (content === null) {
        // Report as SKIPPED, not passed — a silent green here would hide a
        // missing conformance run (e.g. unicode.org unreachable).
        ctx.skip();
        return;
      }

      const lines = content.split('\n');
      let curLevels: string[] = [];
      let curOrder: number[] = [];
      let checked = 0;
      let levelFail = 0;
      let orderFail = 0;
      const samples: string[] = [];

      for (const raw of lines) {
        if (!raw || raw.startsWith('#')) continue;
        if (raw.startsWith('@Levels:')) {
          const v = raw.slice('@Levels:'.length).trim();
          curLevels = v === '' ? [] : v.split(/\s+/);
          continue;
        }
        if (raw.startsWith('@Reorder:')) {
          const v = raw.slice('@Reorder:'.length).trim();
          curOrder = v === '' ? [] : v.split(/\s+/).map(Number);
          continue;
        }
        const semi = raw.indexOf(';');
        if (semi < 0) continue;
        const names = raw.slice(0, semi).trim().split(/\s+/);
        const bitset = parseInt(raw.slice(semi + 1).trim(), 16);
        const cps = names.map((nm) => REP[nm]);

        for (const { bit, base } of BASE_BY_BIT) {
          if (!(bitset & bit)) continue;
          const { levels } = resolveLevels(cps, base);
          const order = reorderByLevels(levels, 0, cps.length);

          let bad = false;
          for (let i = 0; i < curLevels.length; i++) {
            const exp = curLevels[i];
            if (exp === 'x') {
              if (levels[i] !== REMOVED) bad = true;
            } else if (levels[i] !== parseInt(exp, 10)) bad = true;
            if (bad) break;
          }
          if (bad) levelFail++;
          const orderBad =
            order.length !== curOrder.length || order.some((v, i) => v !== curOrder[i]);
          if (orderBad) orderFail++;
          if ((bad || orderBad) && samples.length < 12) {
            samples.push(
              `[${names.join(' ')}] base=${base}\n` +
                `   expLvl=[${curLevels.join(' ')}] gotLvl=[${[...levels].map((l) => (l === REMOVED ? 'x' : l)).join(' ')}]\n` +
                `   expOrd=[${curOrder.join(' ')}] gotOrd=[${order.join(' ')}]`,
            );
          }
          checked++;
        }
      }

      if (levelFail || orderFail) {
        console.error(
          `[bidi bunched] checked=${checked} levelFail=${levelFail} orderFail=${orderFail}\n` +
            samples.join('\n'),
        );
      }
      expect(checked).toBeGreaterThan(490_000);
      expect({ levelFail, orderFail }).toEqual({ levelFail: 0, orderFail: 0 });
    },
    180_000,
  );
});
