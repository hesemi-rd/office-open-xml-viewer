import { describe, expect, it } from 'vitest';
import { rectsOverlap, resolveFloatOverlap } from './float-layout.js';

type Blocker = Readonly<{
  kind: 'table';
  xLeft: number;
  xRight: number;
  yTop: number;
  yBottom: number;
  paraId: number;
}>;

describe('resolveFloatOverlap convergence', () => {
  it('clears every blocker in a finite registry larger than sixteen entries', () => {
    const blockers: Blocker[] = Array.from({ length: 17 }, (_, index) => ({
      kind: 'table',
      xLeft: index * 1.1,
      xRight: (index + 1) * 1.1,
      yTop: 0,
      yBottom: 1,
      paraId: index,
    }));

    const resolved = resolveFloatOverlap(
      0, 0, 1, 1,
      0, 0, 0, 0,
      100, false, 'table', 100,
      blockers,
    );

    expect(resolved.x).toBeCloseTo(18.7);
    expect(blockers.some((blocker) => rectsOverlap(
      resolved.x,
      resolved.x + 1,
      resolved.y,
      resolved.y + 1,
      blocker.xLeft,
      blocker.xRight,
      blocker.yTop,
      blocker.yBottom,
    ))).toBe(false);
  });

  it('throws instead of returning an overlapping placement when resolution cannot converge', () => {
    let left = 0;
    const chasingBlocker = {
      kind: 'table' as const,
      get xLeft() { return left; },
      get xRight() {
        const right = left + 1.1;
        left = right;
        return right;
      },
      yTop: 0,
      yBottom: 1,
      paraId: 0,
    };

    expect(() => resolveFloatOverlap(
      0, 0, 1, 1,
      0, 0, 0, 0,
      100, false, 'table', 100,
      [chasingBlocker],
    )).toThrow('Float overlap resolution did not converge');
  });
});
