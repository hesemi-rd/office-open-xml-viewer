import { priorityKashidaInsertionCandidates } from './arabic-joining.js';

export type KashidaLevel = 'low' | 'medium' | 'high';

export interface KashidaInputSeg {
  text?: string;
}

export interface KashidaInsertion {
  beforeCp: number;
  count: number;
}

export interface KashidaSegmentPlan {
  text: string;
  insertions: KashidaInsertion[];
  advanceDeltaPx: number;
}

export interface KashidaDistribution {
  perSeg: Map<number, KashidaSegmentPlan>;
  appliedPx: number;
  residualPx: number;
}

export type MeasureSegmentAdvance = (si: number, text: string) => number;

const TATWEEL = '\u0640';

function augmentedText(original: string, insertions: ReadonlyMap<number, number>): string {
  const cps = [...original];
  let result = '';
  for (let i = 0; i < cps.length; i++) {
    const count = insertions.get(i) ?? 0;
    if (count > 0) result += TATWEEL.repeat(count);
    result += cps[i];
  }
  return result;
}

/** ECMA-376 §17.18.44 — allocate line slack as true Arabic elongation. */
export function computeKashidaDistribution(
  segments: readonly KashidaInputSeg[],
  slackPx: number,
  level: KashidaLevel,
  measureAdvance: MeasureSegmentAdvance,
): KashidaDistribution | null {
  if (slackPx <= 0.5) return null;

  const candidates: Array<{
    si: number;
    beforeCp: number;
    priority: number;
    textOrder: number;
  }> = [];
  for (let si = 0; si < segments.length; si++) {
    const text = segments[si].text;
    if (text === undefined) continue;
    // A bidi layout segment is a shaping island: font/weight/style boundaries
    // sever Arabic joining (packages/core/src/text/bidi/types.ts), and each
    // segment is painted by its own fillText call. Ranking per segment therefore
    // cannot select a kashida at a join that paint renders as disconnected.
    for (const { beforeCp, priority } of priorityKashidaInsertionCandidates(text)) {
      candidates.push({ si, beforeCp, priority, textOrder: candidates.length });
    }
  }
  if (candidates.length === 0) return null;
  candidates.sort((a, b) => b.priority - a.priority || a.textOrder - b.textOrder);

  const cap = level === 'low' ? 1 : level === 'medium' ? 2 : Infinity;
  const maxRounds = cap;
  const insertionCeiling = candidates.length * 64;
  const countsBySeg = new Map<number, Map<number, number>>();
  const baseAdvance = new Map<number, number>();
  const currentAdvance = new Map<number, number>();
  for (const { si } of candidates) {
    if (currentAdvance.has(si)) continue;
    const advance = measureAdvance(si, segments[si].text!);
    baseAdvance.set(si, advance);
    currentAdvance.set(si, advance);
  }

  let remaining = slackPx;
  let totalInsertions = 0;
  for (
    let round = 0;
    round < maxRounds && remaining > 0.5 && totalInsertions < insertionCeiling;
    round++
  ) {
    let progressed = false;
    for (const { si, beforeCp } of candidates) {
      if (remaining <= 0.5 || totalInsertions >= insertionCeiling) break;
      let counts = countsBySeg.get(si);
      if (!counts) {
        counts = new Map();
        countsBySeg.set(si, counts);
      }
      const previousCount = counts.get(beforeCp) ?? 0;
      if (previousCount >= cap) continue;

      counts.set(beforeCp, previousCount + 1);
      const text = augmentedText(segments[si].text!, counts);
      const nextAdvance = measureAdvance(si, text);
      const delta = nextAdvance - currentAdvance.get(si)!;
      if (delta > 0 && delta <= remaining + 1e-6) {
        currentAdvance.set(si, nextAdvance);
        remaining -= delta;
        totalInsertions++;
        progressed = true;
      } else if (previousCount === 0) {
        counts.delete(beforeCp);
      } else {
        counts.set(beforeCp, previousCount);
      }
    }
    if (!progressed) break;
  }

  const perSeg = new Map<number, KashidaSegmentPlan>();
  for (const [si, counts] of countsBySeg) {
    const insertions = [...counts.entries()]
      .filter(([, count]) => count > 0)
      .sort(([a], [b]) => a - b)
      .map(([beforeCp, count]) => ({ beforeCp, count }));
    if (insertions.length === 0) continue;
    perSeg.set(si, {
      text: augmentedText(segments[si].text!, counts),
      insertions,
      advanceDeltaPx: currentAdvance.get(si)! - baseAdvance.get(si)!,
    });
  }
  if (perSeg.size === 0) return null;

  const appliedPx = [...perSeg.values()].reduce(
    (sum, plan) => sum + plan.advanceDeltaPx,
    0,
  );
  return { perSeg, appliedPx, residualPx: slackPx - appliedPx };
}
