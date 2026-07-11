import { describe, it, expect } from 'vitest';
import {
  computeKashidaDistribution,
  type KashidaLevel,
  type MeasureSegmentAdvance,
} from './kashida-justify.js';

/**
 * Word kashida insertion-priority model (issue #954).
 *
 * Word does NOT distribute tatweel over every eligible Arabic join. It elongates
 * ONE priority join per word, chosen by the documented Uniscribe / Qt
 * SCRIPT_JUSTIFY class ranking (Kashida > Seen > HahDal > Alef > BaRa > Waw >
 * Normal), tie-broken toward the word end. The chosen join set is identical for
 * low/medium/high — only the elongation amount grows.
 *
 * Ground truth: the adjudication fixture PDF (Word export). The same Seen/Sheen/
 * Sad-rich Arabic sentence is set as jc = both / low / medium / high Kashida.
 * Positions were extracted from the PDF (pdftotext -bbox, logical order): Word
 * elongates exactly one join per word, at the offsets asserted below.
 */

const TATWEEL = 'ـ';
const countTatweel = (s: string): number =>
  [...s].filter((c) => c === TATWEEL).length;

// Every code point measures 1px, so an inserted tatweel adds exactly 1px.
const unitMeasure: MeasureSegmentAdvance = (_si, text) => [...text].length;

// The fixture sentence (one repetition). Seen/Sheen/Sad-rich, all connectable.
const SENTENCE =
  'السلسلة تتسلسل في المستشفى بسلاسة وتستمر الكتابة المتصلة بين الحروف';

// Word's chosen priority join per word, expressed as the code-point offset in
// SENTENCE before which the tatweel is inserted. One join per word (10 words):
//   السلسلة  s->l  @5     المستشفى sh->f @24    الكتابة  b->te @47
//   تتسلسل   s->l  @13    بسلاسة   s->te @32    المتصلة  s.->l @54
//   في       f->y  @16    وتستمر   s->t  @37    بين      b->y  @58 (BaRa)
//                                              الحروف   h->r  @64 (HahDal)
const PRIORITY_OFFSETS = [5, 13, 16, 24, 32, 37, 47, 54, 58, 64];

// All eligible interior joins in the same sentence (what a uniform round-robin
// over every join would touch). Word ignores most of these.
const ALL_ELIGIBLE_COUNT = 37;

function insertionOffsets(level: KashidaLevel, slackPx: number): number[] {
  const d = computeKashidaDistribution(
    [{ text: SENTENCE }],
    slackPx,
    level,
    unitMeasure,
  );
  expect(d).not.toBeNull();
  const plan = d!.perSeg.get(0);
  expect(plan).toBeDefined();
  return plan!.insertions.map((ins) => ins.beforeCp).sort((a, b) => a - b);
}

describe('kashida insertion-priority model (issue #954)', () => {
  it('elongates exactly one priority join per word — not every eligible join', () => {
    // Ample slack so every priority join is elongated. Word chooses 10 joins
    // (one per word); a round-robin over all joins would touch far more.
    const offsets = insertionOffsets('high', 400);
    expect(offsets).toEqual(PRIORITY_OFFSETS);
    expect(offsets.length).toBeLessThan(ALL_ELIGIBLE_COUNT);
  });

  it('keeps the same join set across low / medium / high (only elongation grows)', () => {
    const low = insertionOffsets('low', 400);
    const medium = insertionOffsets('medium', 400);
    const high = insertionOffsets('high', 400);
    expect(low).toEqual(PRIORITY_OFFSETS);
    expect(medium).toEqual(PRIORITY_OFFSETS);
    expect(high).toEqual(PRIORITY_OFFSETS);
  });

  it('elongation is monotonic low <= medium <= high', () => {
    const total = (level: KashidaLevel): number => {
      const d = computeKashidaDistribution(
        [{ text: SENTENCE }],
        400,
        level,
        unitMeasure,
      )!;
      return countTatweel(d.perSeg.get(0)!.text);
    };
    const low = total('low');
    const medium = total('medium');
    const high = total('high');
    expect(low).toBeGreaterThanOrEqual(PRIORITY_OFFSETS.length); // >= 1 per join
    expect(low).toBeLessThanOrEqual(medium);
    expect(medium).toBeLessThanOrEqual(high);
    // low is a fixed minimum: with generous slack it does NOT fill everything,
    // so it stays well below high's elongation.
    expect(low).toBeLessThan(high);
  });

  it('picks the Seen join nearest the word end when a word has several', () => {
    // "السلسلة" (a l s l s l te) has two seen->lam joins (cp 3 and cp 5);
    // Word elongates the one nearest the word end (cp 5).
    const offsets = insertionOffsets('high', 400);
    expect(offsets).toContain(5); // second seen->lam, not the first (cp 3)
    expect(offsets).not.toContain(3);
  });

  it('uses the BaRa join (Beh->Yeh) over the final-letter join in بين', () => {
    // "بين" (b y n): eligible joins are b->y (cp 58) and y->n (cp 59).
    // Word elongates b->y (BaRa class), not the join before the final noon.
    const offsets = insertionOffsets('high', 400);
    expect(offsets).toContain(58);
    expect(offsets).not.toContain(59);
  });
});
