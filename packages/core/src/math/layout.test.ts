import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parseMathFont } from './font';
import { parseMathConstants } from './math-table';
import { layoutMath, type MathLayoutCtx } from './layout';
import type { MathNode } from '../types/math';

let ctx: MathLayoutCtx;
beforeAll(() => {
  const url = new URL('../../assets/LatinModernMath.otf', import.meta.url);
  const buf = readFileSync(fileURLToPath(url));
  const font = parseMathFont(
    buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
  );
  ctx = { font, consts: parseMathConstants(font), fontSizePx: 12, level: 'text' };
});

describe('layoutMath', () => {
  it('lays out a single run with positive width and height', () => {
    const nodes: MathNode[] = [{ kind: 'run', text: 'x', style: 'italic' }];
    const box = layoutMath(nodes, ctx);
    expect(box.width).toBeGreaterThan(0);
    expect(box.ascent).toBeGreaterThan(0);
  });

  it('concatenates runs left to right', () => {
    const one = layoutMath([{ kind: 'run', text: 'a', style: 'italic' }], ctx);
    const two = layoutMath(
      [
        { kind: 'run', text: 'a', style: 'italic' },
        { kind: 'run', text: 'b', style: 'italic' },
      ],
      ctx,
    );
    expect(two.width).toBeGreaterThan(one.width);
  });

  it('stacks a fraction taller than the numerator alone, straddling the baseline', () => {
    const num: MathNode[] = [{ kind: 'run', text: '1', style: 'italic' }];
    const den: MathNode[] = [{ kind: 'run', text: 'x', style: 'italic' }];
    const single = layoutMath(num, ctx);
    const frac = layoutMath([{ kind: 'fraction', num, den }], ctx);
    expect(frac.ascent + frac.descent).toBeGreaterThan(single.ascent + single.descent);
    expect(frac.ascent).toBeGreaterThan(0);
    expect(frac.descent).toBeGreaterThan(0);
    // a rule op should be emitted
    expect(frac.ops.some((o) => o.type === 'rule')).toBe(true);
  });

  it('raises a superscript above the base ascent', () => {
    const base: MathNode[] = [{ kind: 'run', text: 'x', style: 'italic' }];
    const plain = layoutMath(base, ctx);
    const sup = layoutMath(
      [{ kind: 'sup', base, sup: [{ kind: 'run', text: '2', style: 'italic' }] }],
      ctx,
    );
    expect(sup.ascent).toBeGreaterThan(plain.ascent);
  });

  it('drops a subscript below the base descent', () => {
    const base: MathNode[] = [{ kind: 'run', text: 'x', style: 'italic' }];
    const plain = layoutMath(base, ctx);
    const sub = layoutMath(
      [{ kind: 'sub', base, sub: [{ kind: 'run', text: 'i', style: 'italic' }] }],
      ctx,
    );
    expect(sub.descent).toBeGreaterThan(plain.descent);
  });

  it('lays out a radical with a surd glyph and a vinculum rule', () => {
    const box = layoutMath(
      [{ kind: 'radical', radicand: [{ kind: 'run', text: 'x', style: 'italic' }] }],
      ctx,
    );
    expect(box.width).toBeGreaterThan(0);
    // surd glyph drawn + at least one vinculum rule
    expect(box.ops.some((o) => o.type === 'glyph' && o.text === '√')).toBe(true);
    expect(box.ops.some((o) => o.type === 'rule')).toBe(true);
    // radical reserves ascent above the bare radicand for the rule + gap
    const bare = layoutMath([{ kind: 'run', text: 'x', style: 'italic' }], ctx);
    expect(box.ascent).toBeGreaterThan(bare.ascent);
  });

  it('lays out an n-ary sum with a body', () => {
    const box = layoutMath(
      [
        {
          kind: 'nary',
          op: '∑',
          sub: [{ kind: 'run', text: 'i', style: 'italic' }],
          sup: [{ kind: 'run', text: 'n', style: 'italic' }],
          body: [{ kind: 'run', text: 'i', style: 'italic' }],
        },
      ],
      ctx,
    );
    expect(box.width).toBeGreaterThan(0);
    expect(box.ascent).toBeGreaterThan(0);
  });
});
