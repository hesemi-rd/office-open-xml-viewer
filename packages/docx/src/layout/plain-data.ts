import type { DeepReadonly } from './types.js';

function assertPlainData(
  value: unknown,
  path: string,
  visiting = new WeakSet<object>(),
  completed = new WeakSet<object>(),
): void {
  if (
    value === null
    || value === undefined
    || typeof value === 'string'
    || typeof value === 'boolean'
  ) return;
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new TypeError(`${path} must contain finite numbers`);
    return;
  }
  if (typeof value !== 'object') {
    throw new TypeError(`${path} must be structured-clone-safe plain data`);
  }
  if (visiting.has(value)) {
    throw new TypeError(`${path} must be structured-clone-safe plain data`);
  }
  if (completed.has(value)) return;
  const prototype = Object.getPrototypeOf(value);
  if (!Array.isArray(value) && prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${path} must be structured-clone-safe plain data`);
  }
  visiting.add(value);
  try {
    for (const [key, child] of Object.entries(value)) {
      assertPlainData(child, `${path}.${key}`, visiting, completed);
    }
  } finally {
    visiting.delete(value);
  }
  completed.add(value);
}

export function deepFreezePlainData<T>(
  value: T,
  seen = new WeakSet<object>(),
): DeepReadonly<T> {
  if (value === null || typeof value !== 'object' || seen.has(value)) {
    return value as DeepReadonly<T>;
  }
  seen.add(value);
  for (const child of Object.values(value)) deepFreezePlainData(child, seen);
  return Object.freeze(value) as DeepReadonly<T>;
}

export function snapshotPlainData<T>(value: T, label: string): DeepReadonly<T> {
  assertPlainData(value, label);
  try {
    return deepFreezePlainData(structuredClone(value));
  } catch {
    throw new TypeError(`${label} must be structured-clone-safe plain data`);
  }
}
