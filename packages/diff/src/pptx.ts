import type {
  Presentation,
  Slide,
  SlideElement,
  ShapeElement,
  PictureElement,
} from '@silurus/ooxml-pptx';
import type { Change, DiffResult, BBox } from './types.ts';
import { deepEqual } from './util/equal.ts';

/** Parser-emitted fields that aren't yet in the public TS surface. The Rust
 *  parser (CHANGELOG 0.32.0) serializes `id` / `name` / `placeholderType` /
 *  `placeholderIdx` on every shape element. */
interface IdentifiedShape extends ShapeElement {
  id?: string;
  name?: string;
  placeholderType?: string;
  placeholderIdx?: number;
}

interface ElementBBox {
  el: SlideElement;
  bbox: BBox;
}

/** Top-level entry. Produce a structural diff between two parsed presentations. */
export function diffPptx(before: Presentation, after: Presentation): DiffResult {
  const changes: Change[] = [];

  diffPresentationMeta(before, after, changes);
  diffSlides(before.slides, after.slides, changes);

  return { format: 'pptx', changes };
}

function diffPresentationMeta(a: Presentation, b: Presentation, out: Change[]): void {
  for (const key of ['slideWidth', 'slideHeight', 'defaultTextColor', 'majorFont', 'minorFont', 'hlinkColor', 'folHlinkColor'] as const) {
    const av = (a as unknown as Record<string, unknown>)[key];
    const bv = (b as unknown as Record<string, unknown>)[key];
    if (!deepEqual(av, bv)) {
      out.push({ op: 'modify', path: key, kind: 'presentation', before: av, after: bv });
    }
  }
}

function diffSlides(beforeSlides: Slide[], afterSlides: Slide[], out: Change[]): void {
  const maxLen = Math.max(beforeSlides.length, afterSlides.length);
  for (let i = 0; i < maxLen; i++) {
    const a = beforeSlides[i];
    const b = afterSlides[i];
    if (a == null && b != null) {
      out.push({
        op: 'add',
        path: `slides[${i}]`,
        kind: 'slide',
        after: b,
        location: { kind: 'slide', slideIndex: i },
      });
      continue;
    }
    if (a != null && b == null) {
      out.push({
        op: 'remove',
        path: `slides[${i}]`,
        kind: 'slide',
        before: a,
        location: { kind: 'slide', slideIndex: i },
      });
      continue;
    }
    diffSlide(a, b, i, out);
  }
}

function diffSlide(a: Slide, b: Slide, slideIndex: number, out: Change[]): void {
  if (!deepEqual(a.background, b.background)) {
    out.push({
      op: 'modify',
      path: `slides[${slideIndex}].background`,
      kind: 'slide-background',
      before: a.background,
      after: b.background,
      location: { kind: 'slide', slideIndex },
    });
  }
  diffElements(a.elements, b.elements, slideIndex, out);
}

/** Element matching strategy:
 *  1. shapes with the same `id` (parser-emitted) on both sides → matched
 *  2. shapes with the same non-empty `name` on both sides → matched
 *  3. remaining elements: greedy nearest-neighbour by type + bbox centre
 *  4. unmatched left → remove, unmatched right → add
 */
function diffElements(left: SlideElement[], right: SlideElement[], slideIndex: number, out: Change[]): void {
  const leftMatched = new Set<number>();
  const rightMatched = new Set<number>();
  const pairs: Array<[number, number]> = [];

  // Pass 1 — by id (PowerPoint reuses cNvPr ids on the same slide, so we
  // keep ALL indices per id and consume them in order rather than taking only
  // the last one. This way, two shapes with id="3" on each side pair up 1:1
  // instead of fighting over the same slot.)
  const rightById = new Map<string, number[]>();
  right.forEach((el, i) => {
    const id = (el as IdentifiedShape).id;
    if (!id) return;
    const list = rightById.get(id) ?? [];
    list.push(i);
    rightById.set(id, list);
  });
  left.forEach((el, i) => {
    const id = (el as IdentifiedShape).id;
    if (!id) return;
    const list = rightById.get(id);
    if (!list || list.length === 0) return;
    const j = list.shift()!;
    pairs.push([i, j]);
    leftMatched.add(i);
    rightMatched.add(j);
  });

  // Pass 2 — by name (same multi-index handling as pass 1)
  const rightByName = new Map<string, number[]>();
  right.forEach((el, i) => {
    if (rightMatched.has(i)) return;
    const name = (el as IdentifiedShape).name;
    if (!name) return;
    const list = rightByName.get(name) ?? [];
    list.push(i);
    rightByName.set(name, list);
  });
  left.forEach((el, i) => {
    if (leftMatched.has(i)) return;
    const name = (el as IdentifiedShape).name;
    if (!name) return;
    const list = rightByName.get(name);
    if (!list || list.length === 0) return;
    const j = list.shift()!;
    pairs.push([i, j]);
    leftMatched.add(i);
    rightMatched.add(j);
  });

  // Pass 3 — greedy nearest neighbour by type + centre distance
  for (let i = 0; i < left.length; i++) {
    if (leftMatched.has(i)) continue;
    const li = left[i];
    const lc = centre(li);
    let bestJ = -1;
    let bestD = Infinity;
    for (let j = 0; j < right.length; j++) {
      if (rightMatched.has(j)) continue;
      const rj = right[j];
      if (rj.type !== li.type) continue;
      const rc = centre(rj);
      const dx = lc.x - rc.x;
      const dy = lc.y - rc.y;
      const d = dx * dx + dy * dy;
      if (d < bestD) {
        bestD = d;
        bestJ = j;
      }
    }
    if (bestJ >= 0) {
      pairs.push([i, bestJ]);
      leftMatched.add(i);
      rightMatched.add(bestJ);
    }
  }

  // Emit changes for matched pairs
  for (const [i, j] of pairs) {
    diffElement(left[i], right[j], slideIndex, i, j, out);
  }

  // Unmatched left → remove
  for (let i = 0; i < left.length; i++) {
    if (leftMatched.has(i)) continue;
    out.push({
      op: 'remove',
      path: `slides[${slideIndex}].elements[${i}]`,
      kind: `element:${left[i].type}`,
      before: left[i],
      location: { kind: 'slide', slideIndex, bbox: bboxOf(left[i]) },
    });
  }

  // Unmatched right → add
  for (let j = 0; j < right.length; j++) {
    if (rightMatched.has(j)) continue;
    out.push({
      op: 'add',
      path: `slides[${slideIndex}].elements[${j}]`,
      kind: `element:${right[j].type}`,
      after: right[j],
      location: { kind: 'slide', slideIndex, bbox: bboxOf(right[j]) },
    });
  }
}

function diffElement(
  a: SlideElement,
  b: SlideElement,
  slideIndex: number,
  leftIdx: number,
  rightIdx: number,
  out: Change[],
): void {
  if (a.type !== b.type) {
    out.push({
      op: 'modify',
      path: `slides[${slideIndex}].elements[${leftIdx}]`,
      kind: 'element-type',
      before: a,
      after: b,
      location: { kind: 'slide', slideIndex, bbox: bboxOf(b) },
    });
    return;
  }

  // Position / size
  for (const k of ['x', 'y', 'width', 'height'] as const) {
    const av = (a as unknown as Record<string, unknown>)[k];
    const bv = (b as unknown as Record<string, unknown>)[k];
    if (av !== bv) {
      out.push({
        op: 'modify',
        path: `slides[${slideIndex}].elements[${leftIdx}].${k}`,
        kind: 'geometry',
        before: av,
        after: bv,
        location: { kind: 'slide', slideIndex, bbox: bboxOf(b) },
      });
    }
  }

  // Type-specific
  if (a.type === 'shape' && b.type === 'shape') {
    diffShape(a, b, slideIndex, leftIdx, rightIdx, out);
  } else if (a.type === 'picture' && b.type === 'picture') {
    diffPicture(a, b, slideIndex, leftIdx, out);
  } else if (a.type === 'table' && b.type === 'table') {
    if (!deepEqual(a, b)) {
      out.push({
        op: 'modify',
        path: `slides[${slideIndex}].elements[${leftIdx}]`,
        kind: 'table',
        before: a,
        after: b,
        location: { kind: 'slide', slideIndex, bbox: bboxOf(b) },
      });
    }
  } else if (a.type === 'chart' && b.type === 'chart') {
    if (!deepEqual(a, b)) {
      out.push({
        op: 'modify',
        path: `slides[${slideIndex}].elements[${leftIdx}]`,
        kind: 'chart',
        before: a,
        after: b,
        location: { kind: 'slide', slideIndex, bbox: bboxOf(b) },
      });
    }
  } else if (a.type === 'media' && b.type === 'media') {
    if (!deepEqual(a, b)) {
      out.push({
        op: 'modify',
        path: `slides[${slideIndex}].elements[${leftIdx}]`,
        kind: 'media',
        before: a,
        after: b,
        location: { kind: 'slide', slideIndex, bbox: bboxOf(b) },
      });
    }
  }
}

function diffShape(
  a: ShapeElement,
  b: ShapeElement,
  slideIndex: number,
  leftIdx: number,
  _rightIdx: number,
  out: Change[],
): void {
  const base = `slides[${slideIndex}].elements[${leftIdx}]`;
  const loc = { kind: 'slide' as const, slideIndex, bbox: bboxOf(b) };

  if (a.geometry !== b.geometry) {
    out.push({ op: 'modify', path: `${base}.geometry`, kind: 'geometry', before: a.geometry, after: b.geometry, location: loc });
  }
  if (a.rotation !== b.rotation) {
    out.push({ op: 'modify', path: `${base}.rotation`, kind: 'geometry', before: a.rotation, after: b.rotation, location: loc });
  }
  if (!deepEqual(a.fill, b.fill)) {
    out.push({ op: 'modify', path: `${base}.fill`, kind: 'fill', before: a.fill, after: b.fill, location: loc });
  }
  if (!deepEqual(a.stroke, b.stroke)) {
    out.push({ op: 'modify', path: `${base}.stroke`, kind: 'stroke', before: a.stroke, after: b.stroke, location: loc });
  }
  diffTextBody(a.textBody, b.textBody, `${base}.textBody`, loc, out);
}

function diffPicture(a: PictureElement, b: PictureElement, slideIndex: number, leftIdx: number, out: Change[]): void {
  const base = `slides[${slideIndex}].elements[${leftIdx}]`;
  const loc = { kind: 'slide' as const, slideIndex, bbox: bboxOf(b) };
  if (a.dataUrl !== b.dataUrl) {
    out.push({ op: 'modify', path: `${base}.dataUrl`, kind: 'image', before: '<image>', after: '<image>', location: loc });
  }
  if (!deepEqual(a.srcRect, b.srcRect)) {
    out.push({ op: 'modify', path: `${base}.srcRect`, kind: 'image-crop', before: a.srcRect, after: b.srcRect, location: loc });
  }
  if (a.alpha !== b.alpha) {
    out.push({ op: 'modify', path: `${base}.alpha`, kind: 'image-alpha', before: a.alpha, after: b.alpha, location: loc });
  }
}

function diffTextBody(
  a: ShapeElement['textBody'],
  b: ShapeElement['textBody'],
  path: string,
  location: { kind: 'slide'; slideIndex: number; bbox?: BBox },
  out: Change[],
): void {
  if (a == null && b == null) return;
  if (a == null || b == null) {
    out.push({ op: 'modify', path, kind: 'text', before: a, after: b, location });
    return;
  }
  if (!deepEqual(a, b)) {
    out.push({
      op: 'modify',
      path,
      kind: 'text',
      before: flattenText(a),
      after: flattenText(b),
      location,
    });
  }
}

function flattenText(textBody: NonNullable<ShapeElement['textBody']>): string {
  const lines: string[] = [];
  for (const para of textBody.paragraphs ?? []) {
    const parts: string[] = [];
    for (const r of para.runs ?? []) {
      if ('text' in r && typeof (r as { text?: unknown }).text === 'string') {
        parts.push((r as { text: string }).text);
      } else if ('lineBreak' in r) {
        parts.push('\n');
      }
    }
    lines.push(parts.join(''));
  }
  return lines.join('\n');
}

function centre(el: SlideElement): { x: number; y: number } {
  const e = el as { x: number; y: number; width: number; height: number };
  return { x: e.x + e.width / 2, y: e.y + e.height / 2 };
}

function bboxOf(el: SlideElement): BBox {
  const e = el as { x: number; y: number; width: number; height: number };
  return { x: e.x, y: e.y, width: e.width, height: e.height };
}

/** Lookup helper for the viewer. Returns one bbox per change at a given slide. */
export function bboxesForSlide(result: DiffResult, slideIndex: number): Array<{ change: Change; bbox: BBox }> {
  const out: Array<{ change: Change; bbox: BBox }> = [];
  for (const c of result.changes) {
    if (c.location?.kind !== 'slide') continue;
    if (c.location.slideIndex !== slideIndex) continue;
    if (!c.location.bbox) continue;
    out.push({ change: c, bbox: c.location.bbox });
  }
  return out;
}

export type { ElementBBox };
