import type { LayoutServices } from './types.js';

export interface DocumentLayoutRuntimeState {
  services: LayoutServices | null;
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
    value: { services: null, defaultCurrentDateMs },
  });
}

export function documentLayoutRuntimeOf(owner: object): DocumentLayoutRuntimeState {
  const runtime = (owner as RuntimeOwner)[documentLayoutRuntime];
  if (runtime) return runtime;
  // Preserve compatibility for instances constructed without the static loader;
  // the normal load path always attaches its captured load-time value explicitly.
  attachDocumentLayoutRuntime(owner, Date.now());
  return (owner as RuntimeOwner)[documentLayoutRuntime] as DocumentLayoutRuntimeState;
}
