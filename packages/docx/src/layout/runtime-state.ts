import type { DeepReadonly, DocumentLayout, LayoutServices } from './types.js';
import type { PaintResourceRegistry } from './types.js';
import type { NumberFormat } from '@silurus/ooxml-core';

export interface DocumentLayoutRuntimeState {
  services: LayoutServices | null;
  retainedErrorLayout: DeepReadonly<DocumentLayout> | null;
  readonly defaultCurrentDateMs: number;
}

const documentLayoutRuntime = Symbol('document-layout-runtime');

type RuntimeOwner = object & {
  [documentLayoutRuntime]?: DocumentLayoutRuntimeState;
};

export function attachDocumentLayoutRuntime(
  owner: object,
  defaultCurrentDateMs: number,
): void {
  Object.defineProperty(owner, documentLayoutRuntime, {
    configurable: false,
    enumerable: false,
    writable: false,
    value: { services: null, retainedErrorLayout: null, defaultCurrentDateMs },
  });
}

export function documentLayoutRuntimeOf(owner: object): DocumentLayoutRuntimeState {
  const runtime = (owner as RuntimeOwner)[documentLayoutRuntime];
  if (runtime) return runtime;
  throw new Error('Document layout runtime is not initialized; attach it explicitly');
}

export interface ImmutableResourceLookup<T> {
  readonly keys: readonly string[];
  resolve(resourceKey: string): T;
}

/** Keep browser/DOM handles in a private closure while exposing fixed immutable membership. */
export function createImmutableResourceLookup<T>(entries: ReadonlyMap<string, T>): ImmutableResourceLookup<T> {
  const snapshot = new Map(entries);
  const keys = Object.freeze([...snapshot.keys()].sort());
  return Object.freeze({
    keys,
    resolve(resourceKey: string): T {
      const value = snapshot.get(resourceKey);
      if (value === undefined) throw new Error(`Unknown runtime resource: ${resourceKey}`);
      return value;
    },
  });
}

const privateResourceLookups = new WeakMap<object, ImmutableResourceLookup<unknown>>();

export function attachPrivateResourceLookup<T>(
  owner: object,
  entries: ReadonlyMap<string, T>,
  expectedKeys: Iterable<string> = entries.keys(),
): void {
  if (privateResourceLookups.has(owner)) {
    throw new Error('Private resource lookup is already attached');
  }
  const actual = new Set(entries.keys());
  const expected = new Set(expectedKeys);
  const missing = [...expected].filter((key) => !actual.has(key)).sort();
  const extra = [...actual].filter((key) => !expected.has(key)).sort();
  if (missing.length > 0 || extra.length > 0) {
    throw new Error(
      `Runtime resource membership mismatch: missing [${missing.join(', ')}]; extra [${extra.join(', ')}]`,
    );
  }
  privateResourceLookups.set(owner, createImmutableResourceLookup(entries));
}

export function privateResourceLookupOf<T>(owner: object): ImmutableResourceLookup<T> | undefined {
  return privateResourceLookups.get(owner) as ImmutableResourceLookup<T> | undefined;
}

const paintResourceRegistries = new WeakMap<object, PaintResourceRegistry>();

export interface FieldAcquisitionContext {
  readonly totalPages: number;
  /** Resolve one PAGE field occurrence from the preceding pagination iteration. */
  readonly resolvePageField?: (
    paragraph: object,
    sourceRunIndex: number,
  ) => PageFieldAcquisitionContext | undefined;
}

export interface PageFieldAcquisitionContext {
  readonly pageIndex: number;
  readonly displayPageNumber: number;
  readonly pageNumberFormat: NumberFormat;
}

const fieldAcquisitionContexts = new WeakMap<object, FieldAcquisitionContext>();

/** Create one immutable service identity for a pagination-field iteration. */
export function createFieldAcquisitionServicesView(
  services: LayoutServices,
  context: FieldAcquisitionContext,
): LayoutServices {
  if (!Number.isInteger(context.totalPages) || context.totalPages < 1) {
    throw new RangeError('Field acquisition totalPages must be a positive integer');
  }
  const view = Object.freeze({ ...services });
  fieldAcquisitionContexts.set(view, Object.freeze({ ...context }));
  return view;
}

export function fieldAcquisitionContextOf(owner: object): FieldAcquisitionContext {
  return fieldAcquisitionContexts.get(owner) ?? Object.freeze({ totalPages: 1 });
}

/** Associate cloneable resource descriptors with their document-scoped owner
 * without widening the stable LayoutServices or public document contracts. */
export function attachPaintResourceRegistry(
  owner: object,
  registry: PaintResourceRegistry,
): void {
  if (paintResourceRegistries.has(owner)) {
    throw new Error('Paint resource registry is already attached');
  }
  paintResourceRegistries.set(owner, registry);
}

export function paintResourceRegistryOf(owner: object): PaintResourceRegistry {
  const registry = paintResourceRegistries.get(owner);
  if (!registry) throw new Error('Paint resource registry is not attached');
  return registry;
}
