import { describe, it, expect, beforeAll } from 'vitest';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { resolveLevels, reorderByLevels, REMOVED } from './uax9/rules.js';
import type { BaseDirection } from './types.js';

// Drives the engine with Unicode's official per-code-point conformance data.
// The 6.6 MB fixture is downloaded + cached on first run (gitignored). If it is
// absent and cannot be downloaded, the test is skipped LOUDLY (no silent pass).

const VERSION = '17.0.0';
const URL = `https://www.unicode.org/Public/${VERSION}/ucd/BidiCharacterTest.txt`;
const FIX_DIR = join(dirname(fileURLToPath(import.meta.url)), '__fixtures__');
const FIX = join(FIX_DIR, 'BidiCharacterTest.txt');

const BASE_BY_FIELD: Record<string, BaseDirection> = { '0': 'ltr', '1': 'rtl', '2': 'auto' };

let content: string | null = null;

beforeAll(async () => {
  if (existsSync(FIX)) {
    content = await readFile(FIX, 'utf8');
    return;
  }
  try {
    const res = await fetch(URL);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    content = await res.text();
    await mkdir(FIX_DIR, { recursive: true });
    await writeFile(FIX, content, 'utf8');
  } catch (e) {
    console.warn(`[bidi conformance] SKIPPED — could not load ${URL}: ${(e as Error).message}`);
    content = null;
  }
}, 60_000);

describe('UAX#9 conformance — BidiCharacterTest.txt', () => {
  it(
    'matches resolved levels and reorder for every line',
    (ctx) => {
      if (content === null) {
        // Report as SKIPPED, not passed — a silent green here would hide a
        // missing conformance run (e.g. unicode.org unreachable).
        ctx.skip();
        return;
      }

      const lines = content.split('\n');
      let checked = 0;
      let levelFail = 0;
      let paraFail = 0;
      let orderFail = 0;
      const samples: string[] = [];

      for (let ln = 0; ln < lines.length; ln++) {
        const raw = lines[ln];
        if (!raw || raw.startsWith('#')) continue;
        const [f0, f1, f2, f3, f4] = raw.split(';');
        if (f4 === undefined) continue;

        const cps = f0.trim().split(/\s+/).map((h) => parseInt(h, 16));
        const base = BASE_BY_FIELD[f1.trim()];
        const expParaLevel = parseInt(f2.trim(), 10);
        const expLevels = f3.trim().split(/\s+/); // entries are numbers or 'x'
        const expOrder = f4.trim() === '' ? [] : f4.trim().split(/\s+/).map(Number);

        const { levels, paragraphLevel } = resolveLevels(cps, base);
        const order = reorderByLevels(levels, 0, cps.length);

        let bad = false;
        if (paragraphLevel !== expParaLevel) {
          paraFail++;
          bad = true;
        }
        for (let i = 0; i < expLevels.length; i++) {
          const exp = expLevels[i];
          if (exp === 'x') {
            if (levels[i] !== REMOVED) {
              levelFail++;
              bad = true;
              break;
            }
          } else if (levels[i] !== parseInt(exp, 10)) {
            levelFail++;
            bad = true;
            break;
          }
        }
        if (order.length !== expOrder.length || order.some((v, i) => v !== expOrder[i])) {
          orderFail++;
          bad = true;
        }
        if (bad && samples.length < 12) {
          samples.push(
            `L${ln + 1}: [${f0.trim()}] dir=${f1} expPara=${expParaLevel} gotPara=${paragraphLevel}\n` +
              `   expLvl=[${expLevels.join(' ')}] gotLvl=[${[...levels].map((l) => (l === REMOVED ? 'x' : l)).join(' ')}]\n` +
              `   expOrд=[${expOrder.join(' ')}] gotOrd=[${order.join(' ')}]`,
          );
        }
        checked++;
      }

      if (levelFail || paraFail || orderFail) {
        console.error(
          `[bidi conformance] checked=${checked} levelFail=${levelFail} paraFail=${paraFail} orderFail=${orderFail}\n` +
            samples.join('\n'),
        );
      }
      // Guard against a silently-truncated fixture: BidiCharacterTest 17.0.0 has
      // 91,707 data lines. Require the bulk of them so a partial download fails.
      expect(checked).toBeGreaterThan(90_000);
      expect({ levelFail, paraFail, orderFail }).toEqual({
        levelFail: 0,
        paraFail: 0,
        orderFail: 0,
      });
    },
    120_000,
  );
});
