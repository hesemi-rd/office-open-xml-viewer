import { LayoutInvariantError } from './diagnostics.js';
import type {
  DeepReadonly,
  DocumentLayout,
  DrawingLayout,
  FlowDomain,
  LayoutRect,
  LayoutPage,
  PageLayerId,
  PaintNode,
  PointPt,
} from './types.js';

const PAGE_LAYER_IDS = [
  'background',
  'behindText',
  'header',
  'body',
  'notes',
  'front',
  'footer',
] as const satisfies readonly PageLayerId[];

type LayerNode = Readonly<{ layer: PageLayerId; node: PaintNode }>;

const layerNodes = (page: LayoutPage): readonly LayerNode[] => PAGE_LAYER_IDS.flatMap((layer) => (
  page.layers[layer].map((node) => ({ layer, node }))
));

function assertPlainData(value: unknown, path: string, ancestors = new WeakSet<object>()): void {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return;
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} is not finite`);
    }
    return;
  }
  if (typeof value !== 'object') {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} contains ${typeof value}`);
  }
  if (ancestors.has(value)) {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} contains a cycle`);
  }

  ancestors.add(value);
  try {
    if (Array.isArray(value)) {
      for (let index = 0; index < value.length; index += 1) {
        if (!(index in value)) {
          throw new LayoutInvariantError('INVALID_GEOMETRY', `${path}[${index}] is missing`);
        }
        assertPlainData(value[index], `${path}[${index}]`, ancestors);
      }
      if (Object.keys(value).length !== value.length) {
        throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} has a non-index property`);
      }
      if (Object.getOwnPropertySymbols(value).length > 0) {
        throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} has a symbol key`);
      }
      return;
    }

    const prototype = Object.getPrototypeOf(value);
    if (prototype !== Object.prototype && prototype !== null) {
      throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} is not a plain record`);
    }
    for (const key of Reflect.ownKeys(value)) {
      if (typeof key !== 'string') {
        throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} has a symbol key`);
      }
      const descriptor = Object.getOwnPropertyDescriptor(value, key);
      if (!descriptor?.enumerable || !('value' in descriptor)) {
        throw new LayoutInvariantError('INVALID_GEOMETRY', `${path}.${key} is not plain data`);
      }
      assertPlainData(descriptor.value, `${path}.${key}`, ancestors);
    }
  } finally {
    ancestors.delete(value);
  }
}

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

function contains(outer: LayoutRect, inner: LayoutRect): boolean {
  return inner.xPt >= outer.xPt
    && inner.yPt >= outer.yPt
    && inner.xPt + inner.widthPt <= outer.xPt + outer.widthPt
    && inner.yPt + inner.heightPt <= outer.yPt + outer.heightPt;
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
  assertPlainData(layout, 'layout');
  layout.pages.forEach((page, pageIndex) => {
    requireRect(page.geometry, `pages[${pageIndex}].geometry`);
    requireFinite(page.geometry.contentTopPt, `pages[${pageIndex}].geometry.contentTopPt`);
    requireFinite(page.geometry.contentBottomPt, `pages[${pageIndex}].geometry.contentBottomPt`);

    const domains = new Map<string, FlowDomain>();
    page.flowDomains.forEach((domain, domainIndex) => {
      requireRect(domain.bounds, `pages[${pageIndex}].flowDomains[${domainIndex}].bounds`);
      if (domains.has(domain.id)) {
        throw new LayoutInvariantError('INVALID_REFERENCE', `duplicate flow domain ${domain.id}`);
      }
      domains.set(domain.id, domain);
    });

    const ordinary: PaintNode[] = [];
    const nodes = new Map<string, LayerNode>();
    layerNodes(page).forEach(({ layer, node }, nodeIndex) => {
      const path = `pages[${pageIndex}].nodes[${nodeIndex}]`;
      if (nodes.has(node.id)) {
        throw new LayoutInvariantError('INVALID_REFERENCE', `duplicate node ${node.id}`);
      }
      nodes.set(node.id, { layer, node });
      requireRect(node.flowBounds, `${path}.flowBounds`);
      requireRect(node.inkBounds, `${path}.inkBounds`);
      if (node.clipBounds) requireRect(node.clipBounds, `${path}.clipBounds`);
      requireFinite(node.advancePt, `${path}.advancePt`);
      if (node.kind === 'drawing') requireDrawingGeometry(node, path);
      const domain = domains.get(node.flowDomainId);
      if (!domain) {
        throw new LayoutInvariantError('INVALID_REFERENCE', `${node.id} references missing flow domain ${node.flowDomainId}`);
      }
      if (!node.ordinaryFlow) return;
      if (domain.kind === 'body'
        && node.flowBounds.yPt + node.flowBounds.heightPt > page.geometry.contentBottomPt) {
        throw new LayoutInvariantError('BOTTOM_MARGIN_INVASION', `${node.id} crosses contentBottomPt`);
      }
      if (!contains(domain.bounds, node.flowBounds)) {
        throw new LayoutInvariantError('FLOW_DOMAIN_INVASION', `${node.id} crosses flow domain ${domain.id}`);
      }
      ordinary.push(node);
    });

    const painted = new Set<string>();
    page.layers.paintOrder.forEach((entry) => {
      const target = nodes.get(entry.nodeId);
      if (!target || target.layer !== entry.layer || painted.has(entry.nodeId)) {
        throw new LayoutInvariantError('INVALID_REFERENCE', `invalid paint reference ${entry.layer}:${entry.nodeId}`);
      }
      painted.add(entry.nodeId);
    });
    if (painted.size !== nodes.size) {
      const missing = [...nodes.keys()].find((id) => !painted.has(id));
      throw new LayoutInvariantError('INVALID_REFERENCE', `node ${missing ?? '<unknown>'} is absent from paintOrder`);
    }

    const read = new Set<string>();
    page.readingOrder.forEach((nodeId) => {
      if (!nodes.has(nodeId) || read.has(nodeId)) {
        throw new LayoutInvariantError('INVALID_REFERENCE', `invalid reading-order reference ${nodeId}`);
      }
      read.add(nodeId);
    });

    for (let index = 0; index < ordinary.length; index += 1) {
      for (let other = index + 1; other < ordinary.length; other += 1) {
        const first = ordinary[index];
        const second = ordinary[other];
        if (first && second
          && first.flowDomainId === second.flowDomainId
          && overlaps(first.flowBounds, second.flowBounds)) {
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
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([entryKey, entry]) => [entryKey, canonicalize(entry)]));
  }
  throw new LayoutInvariantError('INVALID_GEOMETRY', `fingerprint contains ${typeof value}`);
}

export function layoutFingerprint(layout: DocumentLayout): string {
  assertPlainData(layout, 'layout');
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
  assertPlainData(layout, 'layout');
  return deepFreeze(layout, new WeakSet<object>());
}
