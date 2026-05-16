import type {
  Document,
  BodyElement,
  DocParagraph,
  DocRun,
  TextRun,
} from '@silurus/ooxml-docx';
import type { Change, DiffResult } from './types.ts';
import { deepEqual, stableText } from './util/equal.ts';
import { alignSequences } from './util/sequence.ts';

/** Top-level entry. Produce a structural diff between two parsed DOCX documents. */
export function diffDocx(before: Document, after: Document): DiffResult {
  const changes: Change[] = [];

  if (!deepEqual(before.section, after.section)) {
    changes.push({
      op: 'modify',
      path: 'section',
      kind: 'section',
      before: before.section,
      after: after.section,
    });
  }

  diffBody(before.body, after.body, changes);

  if (!deepEqual(before.headers, after.headers)) {
    changes.push({
      op: 'modify',
      path: 'headers',
      kind: 'headers',
      before: before.headers,
      after: after.headers,
    });
  }
  if (!deepEqual(before.footers, after.footers)) {
    changes.push({
      op: 'modify',
      path: 'footers',
      kind: 'footers',
      before: before.footers,
      after: after.footers,
    });
  }

  return { format: 'docx', changes };
}

function diffBody(left: BodyElement[], right: BodyElement[], out: Change[]): void {
  // Align body elements by (type, text-signature). Text-signature for a paragraph
  // is the concatenated run text; for a table it's a stable JSON-ish hash. This
  // lets the LCS distinguish "paragraph 3 was modified" from "a paragraph was
  // inserted between 2 and 3".
  const leftSigs = left.map(elementSignature);
  const rightSigs = right.map(elementSignature);

  const alignment = alignSequences(leftSigs, rightSigs, (a, b) => a === b);
  const matchedLeft = new Set(alignment.matches.map(([i]) => i));
  const matchedRight = new Set(alignment.matches.map(([, j]) => j));

  for (const [i, j] of alignment.matches) {
    diffBodyElement(left[i], right[j], i, j, out);
  }

  // Second pass: for unmatched elements, pair them up by index proximity so
  // small edits register as modify rather than remove+add.
  const removedQueue = alignment.removed.slice();
  const addedQueue = alignment.added.slice();
  while (removedQueue.length > 0 && addedQueue.length > 0) {
    const li = removedQueue.shift()!;
    let bestK = 0;
    let bestD = Math.abs(addedQueue[0] - li);
    for (let k = 1; k < addedQueue.length; k++) {
      const d = Math.abs(addedQueue[k] - li);
      if (d < bestD) {
        bestD = d;
        bestK = k;
      }
    }
    const rj = addedQueue.splice(bestK, 1)[0];
    matchedLeft.add(li);
    matchedRight.add(rj);
    if (left[li].type !== right[rj].type) {
      out.push({
        op: 'modify',
        path: `body[${li}]`,
        kind: 'element-type',
        before: left[li],
        after: right[rj],
        location: { kind: 'paragraph', paragraphIndex: li },
      });
    } else {
      diffBodyElement(left[li], right[rj], li, rj, out);
    }
  }

  for (const i of removedQueue) {
    out.push({
      op: 'remove',
      path: `body[${i}]`,
      kind: 'paragraph',
      before: left[i],
      location: { kind: 'paragraph', paragraphIndex: i },
    });
  }
  for (const j of addedQueue) {
    out.push({
      op: 'add',
      path: `body[${j}]`,
      kind: 'paragraph',
      after: right[j],
      location: { kind: 'paragraph', paragraphIndex: j },
    });
  }
}

function diffBodyElement(
  a: BodyElement,
  b: BodyElement,
  leftIdx: number,
  _rightIdx: number,
  out: Change[],
): void {
  if (a.type !== b.type) {
    out.push({
      op: 'modify',
      path: `body[${leftIdx}]`,
      kind: 'element-type',
      before: a,
      after: b,
      location: { kind: 'paragraph', paragraphIndex: leftIdx },
    });
    return;
  }
  if (a.type === 'paragraph' && b.type === 'paragraph') {
    diffParagraph(a, b, leftIdx, out);
    return;
  }
  if (a.type === 'table' && b.type === 'table') {
    if (!deepEqual(a, b)) {
      out.push({
        op: 'modify',
        path: `body[${leftIdx}]`,
        kind: 'table',
        before: a,
        after: b,
        location: { kind: 'paragraph', paragraphIndex: leftIdx },
      });
    }
    return;
  }
  if (a.type === 'pageBreak' && b.type === 'pageBreak') {
    if (a.parity !== b.parity) {
      out.push({
        op: 'modify',
        path: `body[${leftIdx}]`,
        kind: 'pageBreak',
        before: a,
        after: b,
        location: { kind: 'paragraph', paragraphIndex: leftIdx },
      });
    }
  }
}

function diffParagraph(a: DocParagraph & { type: 'paragraph' }, b: DocParagraph & { type: 'paragraph' }, leftIdx: number, out: Change[]): void {
  const base = `body[${leftIdx}]`;
  const loc = { kind: 'paragraph' as const, paragraphIndex: leftIdx };

  // Paragraph-level formatting properties (compare individually to give
  // useful change records rather than one giant "paragraph differs").
  const pPrKeys = [
    'alignment', 'indentLeft', 'indentRight', 'indentFirst',
    'spaceBefore', 'spaceAfter', 'shading', 'styleId',
  ] as const;
  for (const k of pPrKeys) {
    const av = (a as unknown as Record<string, unknown>)[k];
    const bv = (b as unknown as Record<string, unknown>)[k];
    if (!deepEqual(av, bv)) {
      out.push({ op: 'modify', path: `${base}.${k}`, kind: 'paragraph-fmt', before: av, after: bv, location: loc });
    }
  }
  if (!deepEqual(a.lineSpacing, b.lineSpacing)) {
    out.push({ op: 'modify', path: `${base}.lineSpacing`, kind: 'paragraph-fmt', before: a.lineSpacing, after: b.lineSpacing, location: loc });
  }
  if (!deepEqual(a.numbering, b.numbering)) {
    out.push({ op: 'modify', path: `${base}.numbering`, kind: 'numbering', before: a.numbering, after: b.numbering, location: loc });
  }

  // Run-level diff
  diffRuns(a.runs, b.runs, base, leftIdx, out);
}

function diffRuns(left: DocRun[], right: DocRun[], base: string, paragraphIndex: number, out: Change[]): void {
  const leftSig = left.map(runSignature);
  const rightSig = right.map(runSignature);
  const alignment = alignSequences(leftSig, rightSig, (a, b) => a === b);
  const loc = { kind: 'paragraph' as const, paragraphIndex };

  const matchedLeft = new Set(alignment.matches.map(([i]) => i));
  const matchedRight = new Set(alignment.matches.map(([, j]) => j));

  for (const [i, j] of alignment.matches) {
    if (!deepEqual(left[i], right[j])) {
      out.push({
        op: 'modify',
        path: `${base}.runs[${i}]`,
        kind: 'run',
        before: textOfRun(left[i]),
        after: textOfRun(right[j]),
        location: loc,
      });
    }
  }
  for (const i of alignment.removed) {
    if (matchedLeft.has(i)) continue;
    out.push({
      op: 'remove',
      path: `${base}.runs[${i}]`,
      kind: 'run',
      before: textOfRun(left[i]),
      location: loc,
    });
  }
  for (const j of alignment.added) {
    if (matchedRight.has(j)) continue;
    out.push({
      op: 'add',
      path: `${base}.runs[${j}]`,
      kind: 'run',
      after: textOfRun(right[j]),
      location: loc,
    });
  }
}

function elementSignature(el: BodyElement): string {
  if (el.type === 'paragraph') {
    return `P:${paragraphTextSignature(el)}`;
  }
  if (el.type === 'table') {
    return `T:${stableText(el)}`;
  }
  return `B:${el.parity ?? ''}`;
}

function paragraphTextSignature(p: DocParagraph): string {
  const parts: string[] = [];
  for (const r of p.runs) {
    if (r.type === 'text') parts.push((r as TextRun).text);
    else if (r.type === 'break') parts.push('\n');
    else if (r.type === 'field') parts.push((r as { fallbackText: string }).fallbackText);
  }
  return parts.join('');
}

function runSignature(r: DocRun): string {
  if (r.type === 'text') return `t:${(r as TextRun).text}`;
  if (r.type === 'break') return `b:${(r as { breakType: string }).breakType}`;
  if (r.type === 'field') return `f:${(r as { fallbackText: string }).fallbackText}`;
  if (r.type === 'image') return `i:`;
  if (r.type === 'shape') return `s:`;
  return 'unknown';
}

function textOfRun(r: DocRun): unknown {
  if (r.type === 'text') return (r as TextRun).text;
  if (r.type === 'field') return (r as { fallbackText: string }).fallbackText;
  return r;
}
