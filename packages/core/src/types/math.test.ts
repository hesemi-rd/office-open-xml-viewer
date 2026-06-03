import { describe, it, expect } from 'vitest';
import { isMathNode, type MathNode } from './math';

describe('OMML AST', () => {
  it('discriminates node kinds', () => {
    const frac: MathNode = {
      kind: 'fraction',
      num: [{ kind: 'run', text: '1', style: 'italic' }],
      den: [{ kind: 'run', text: 'x', style: 'italic' }],
    };
    expect(frac.kind).toBe('fraction');
    expect(isMathNode(frac)).toBe(true);
    expect(isMathNode({ foo: 1 })).toBe(false);
  });
});
