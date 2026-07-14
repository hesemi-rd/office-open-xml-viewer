export interface CanvasFontRoute {
  /** Complete CSS family list. Consumers must not append another fallback tail. */
  readonly familyList: string;
  readonly scope: 'registered' | 'native' | 'generic';
  /** Syntax identity only; this never asserts portable glyph geometry. */
  readonly fingerprint: string;
}

export function createCanvasFontRoute(
  familyList: string,
  scope: CanvasFontRoute['scope'],
): CanvasFontRoute {
  const normalized = familyList.trim();
  if (!normalized) throw new TypeError('Canvas font route requires a family list');
  return Object.freeze({
    familyList: normalized,
    scope,
    fingerprint: `canvas-font-route-v1:${encodeURIComponent(scope)}:${encodeURIComponent(normalized)}`,
  });
}

export function canvasFontString(
  route: Readonly<CanvasFontRoute>,
  sizePx: number,
  weight: number,
  style: 'normal' | 'italic',
): string {
  if (!Number.isFinite(sizePx) || sizePx < 0) throw new RangeError('Canvas font size must be finite and non-negative');
  if (!Number.isFinite(weight) || weight < 1 || weight > 1000) {
    throw new RangeError('Canvas font weight must be finite and between 1 and 1000');
  }
  if (!route.familyList.trim()) throw new TypeError('Canvas font route requires a family list');
  return `${style} ${weight} ${sizePx}px ${route.familyList}`;
}
