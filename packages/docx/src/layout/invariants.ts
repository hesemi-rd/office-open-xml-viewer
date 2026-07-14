import { LayoutInvariantError } from './diagnostics.js';
import type {
  DeepReadonly,
  DocumentLayout,
  DrawingLayout,
  LayoutRect,
  LayoutPage,
  PaintNode,
  PointPt,
} from './types.js';

const layerNodes = (page: LayoutPage): readonly PaintNode[] => [
  ...page.layers.background,
  ...page.layers.behindText,
  ...page.layers.header,
  ...page.layers.body,
  ...page.layers.notes,
  ...page.layers.front,
  ...page.layers.footer,
];

function requireFinite(value: number, path: string): void {
  if (!Number.isFinite(value)) {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} is not finite`);
  }
}

function requirePoint(point: PointPt, path: string): void {
  requireFinite(point.xPt, `${path}.xPt`);
  requireFinite(point.yPt, `${path}.yPt`);
}

function requireRect(rect: LayoutRect, path: string): void {
  requirePoint(rect, path);
  requireFinite(rect.widthPt, `${path}.widthPt`);
  requireFinite(rect.heightPt, `${path}.heightPt`);
  if (rect.widthPt < 0 || rect.heightPt < 0) {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} has a negative extent`);
  }
}

function overlaps(a: LayoutRect, b: LayoutRect): boolean {
  return a.xPt < b.xPt + b.widthPt
    && b.xPt < a.xPt + a.widthPt
    && a.yPt < b.yPt + b.heightPt
    && b.yPt < a.yPt + a.heightPt;
}

function requireDrawingGeometry(node: DrawingLayout, path: string): void {
  if (node.transform) {
    for (const key of ['a', 'b', 'c', 'd', 'e', 'f'] as const) {
      requireFinite(node.transform[key], `${path}.transform.${key}`);
    }
  }
  if (node.clip?.kind === 'rect') requireRect(node.clip.rect, `${path}.clip.rect`);
  if (node.clip?.kind === 'polygon') {
    node.clip.points.forEach((point, index) => requirePoint(point, `${path}.clip.points[${index}]`));
  }
  node.commands.forEach((command, index) => requireRect(command.rect, `${path}.commands[${index}].rect`));
}

export function assertDocumentLayout(layout: DocumentLayout): void {
  layout.pages.forEach((page, pageIndex) => {
    requireRect(page.geometry, `pages[${pageIndex}].geometry`);
    requireFinite(page.geometry.contentTopPt, `pages[${pageIndex}].geometry.contentTopPt`);
    requireFinite(page.geometry.contentBottomPt, `pages[${pageIndex}].geometry.contentBottomPt`);

    const ordinary: PaintNode[] = [];
    layerNodes(page).forEach((node, nodeIndex) => {
      const path = `pages[${pageIndex}].nodes[${nodeIndex}]`;
      requireRect(node.flowBounds, `${path}.flowBounds`);
      requireRect(node.inkBounds, `${path}.inkBounds`);
      if (node.clipBounds) requireRect(node.clipBounds, `${path}.clipBounds`);
      requireFinite(node.advancePt, `${path}.advancePt`);
      if (node.kind === 'drawing') requireDrawingGeometry(node, path);
      if (!node.ordinaryFlow) return;
      if (node.flowBounds.yPt + node.flowBounds.heightPt > page.geometry.contentBottomPt) {
        throw new LayoutInvariantError('BOTTOM_MARGIN_INVASION', `${node.id} crosses contentBottomPt`);
      }
      ordinary.push(node);
    });

    for (let index = 0; index < ordinary.length; index += 1) {
      for (let other = index + 1; other < ordinary.length; other += 1) {
        const first = ordinary[index];
        const second = ordinary[other];
        if (first && second && overlaps(first.flowBounds, second.flowBounds)) {
          throw new LayoutInvariantError('FLOW_OVERLAP', `${first.id} overlaps ${second.id}`);
        }
      }
    }
  });
}

function canonicalize(value: unknown): unknown {
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new LayoutInvariantError('INVALID_GEOMETRY', 'fingerprint input is not finite');
    const normalized = Number(value.toFixed(6));
    return Object.is(normalized, -0) ? 0 : normalized;
  }
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return value;
  if (Array.isArray(value)) return value.map((entry) => canonicalize(entry));
  if (typeof value === 'object') {
    return Object.fromEntries(Object.entries(value as Record<string, unknown>)
      .filter(([, entry]) => entry !== undefined)
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([entryKey, entry]) => [entryKey, canonicalize(entry)]));
  }
  throw new LayoutInvariantError('INVALID_GEOMETRY', `fingerprint contains ${typeof value}`);
}

export function layoutFingerprint(layout: DocumentLayout): string {
  const value = {
    pages: layout.pages,
    diagnostics: layout.diagnostics.map(({ message: _message, ...identity }) => identity),
  };
  return JSON.stringify(canonicalize(value));
}

function deepFreeze<T>(value: T, seen: WeakSet<object>): DeepReadonly<T> {
  if (value === null || typeof value !== 'object') return value as DeepReadonly<T>;
  if (seen.has(value)) return value as DeepReadonly<T>;
  seen.add(value);
  for (const child of Object.values(value as Record<string, unknown>)) deepFreeze(child, seen);
  return Object.freeze(value) as DeepReadonly<T>;
}

export function deepFreezeDocumentLayout(layout: DocumentLayout): DeepReadonly<DocumentLayout> {
  return deepFreeze(layout, new WeakSet<object>());
}
