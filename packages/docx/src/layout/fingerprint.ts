function canonicalValue(value: unknown): string {
  if (value === null || typeof value === 'boolean' || typeof value === 'number') {
    return JSON.stringify(value);
  }
  if (typeof value === 'string') return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(canonicalValue).join(',')}]`;
  if (typeof value === 'object') {
    const record = value as Readonly<Record<string, unknown>>;
    return `{${Object.keys(record).sort().map((key) => `${JSON.stringify(key)}:${canonicalValue(record[key])}`).join(',')}}`;
  }
  throw new TypeError(`Cannot fingerprint ${typeof value}`);
}

export function stableFingerprint(namespace: string, value: unknown): string {
  const input = canonicalValue(value);
  // This value is used as an identity, not merely as a hash-table hint. Keep
  // the canonical input itself so two resources can never alias through a
  // short digest collision (the previous FNV-1a/32 representation did).
  return `${namespace}:${encodeURIComponent(input)}`;
}
