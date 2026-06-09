import { describe, it, expect, afterEach } from 'vitest';
import {
  getDefaultBidiEngine,
  setBidiEngine,
  resetBidiEngine,
  type BidiEngine,
} from './engine.js';

afterEach(() => resetBidiEngine());

describe('BidiEngine seam', () => {
  it('exposes the three UAX#9 methods on the default engine', () => {
    const engine = getDefaultBidiEngine();
    expect(typeof engine.computeLevels).toBe('function');
    expect(typeof engine.reorderVisual).toBe('function');
    expect(typeof engine.getMirror).toBe('function');
  });

  it('returns a stable singleton', () => {
    expect(getDefaultBidiEngine()).toBe(getDefaultBidiEngine());
  });

  it('can be swapped and restored', () => {
    const fake: BidiEngine = {
      computeLevels: () => ({ levels: new Uint8Array(), paragraphLevel: 0 }),
      reorderVisual: () => [42],
      getMirror: () => 0xbeef,
    };
    setBidiEngine(fake);
    expect(getDefaultBidiEngine()).toBe(fake);
    expect(getDefaultBidiEngine().reorderVisual(new Uint8Array(), 0, 0)).toEqual([42]);

    resetBidiEngine();
    expect(getDefaultBidiEngine()).not.toBe(fake);
    expect(typeof getDefaultBidiEngine().getMirror).toBe('function');
  });
});
