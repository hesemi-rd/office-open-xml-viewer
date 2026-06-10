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

  it('retains characters at resolved level MAX_DEPTH+1 (deep isolates)', () => {
    // 63 RLEs raise the explicit level to the 125 cap; an L character there
    // resolves to 126 via I2 (UAX#9 §3.3.4). It must still appear in the
    // visual order (regression: a `level <= MAX_DEPTH` filter dropped it).
    const engine = getDefaultBidiEngine();
    const text = '‫'.repeat(63) + 'a';
    const { levels } = engine.computeLevels(text, 'ltr');
    expect(levels[63]).toBe(126);
    expect(engine.reorderVisual(levels, 0, text.length)).toContain(63);
  });

  it('the default engine mirrors brackets (L4) and computes levels', () => {
    const engine = getDefaultBidiEngine();
    expect(engine.getMirror(0x28)).toBe(0x29); // ( -> )
    expect(engine.getMirror(0x3c)).toBe(0x3e); // < -> >
    expect(engine.getMirror(0x41)).toBeNull();

    // "אב" (two Hebrew letters) under auto base resolves to RTL level 1.
    const { levels, paragraphLevel } = engine.computeLevels('אב', 'auto');
    expect(paragraphLevel).toBe(1);
    expect([...levels]).toEqual([1, 1]);
    // Pure RTL line reverses under L2.
    expect(engine.reorderVisual(levels, 0, 2)).toEqual([1, 0]);
  });
});
