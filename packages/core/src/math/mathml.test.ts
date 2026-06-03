import { describe, it, expect } from 'vitest';
import { mathToMathML } from './mathml';
import type { MathNode } from '../types/math';

describe('mathToMathML', () => {
  it('wraps in a math element with display mode', () => {
    const ml = mathToMathML([{ kind: 'run', text: 'x', style: 'italic' }], true);
    expect(ml).toContain('<math');
    expect(ml).toContain('display="block"');
    expect(ml).toContain('<mi>x</mi>');
  });

  it('tokenizes runs into mi/mn/mo by class', () => {
    const ml = mathToMathML([{ kind: 'run', text: 'a+12=b', style: 'italic' }], false);
    expect(ml).toContain('<mi>a</mi>');
    expect(ml).toContain('<mo>+</mo>');
    expect(ml).toContain('<mn>12</mn>');
    expect(ml).toContain('<mo>=</mo>');
    expect(ml).toContain('<mi>b</mi>');
  });

  it('marks roman style as mathvariant=normal', () => {
    const ml = mathToMathML([{ kind: 'run', text: 'd', style: 'roman' }], false);
    expect(ml).toContain('mathvariant="normal"');
  });

  it('emits mfrac / msup / msqrt / msubsup', () => {
    const nodes: MathNode[] = [
      {
        kind: 'fraction',
        num: [{ kind: 'run', text: '1', style: 'italic' }],
        den: [
          {
            kind: 'radical',
            radicand: [
              { kind: 'sup', base: [{ kind: 'run', text: 'x', style: 'italic' }], sup: [{ kind: 'run', text: '2', style: 'italic' }] },
            ],
          },
        ],
      },
    ];
    const ml = mathToMathML(nodes, true);
    expect(ml).toContain('<mfrac>');
    expect(ml).toContain('<msqrt>');
    expect(ml).toContain('<msup>');
  });

  it('builds n-ary limits with munderover', () => {
    const ml = mathToMathML(
      [
        {
          kind: 'nary',
          op: '∑',
          sub: [{ kind: 'run', text: 'i', style: 'italic' }],
          sup: [{ kind: 'run', text: 'n', style: 'italic' }],
          body: [{ kind: 'run', text: 'i', style: 'italic' }],
        },
      ],
      true,
    );
    expect(ml).toContain('<munderover>');
    expect(ml).toContain('∑');
  });

  it('emits munder for a lower limit (lim)', () => {
    const ml = mathToMathML(
      [
        {
          kind: 'limit',
          base: [{ kind: 'run', text: 'lim', style: 'roman' }],
          lower: [{ kind: 'run', text: 'n→∞', style: 'italic' }],
        },
      ],
      true,
    );
    expect(ml).toContain('<munder>');
  });

  it('emits an mtable for an eqArr with alternating alignment', () => {
    const ml = mathToMathML(
      [
        {
          kind: 'array',
          align: 'eq',
          rows: [
            [[{ kind: 'run', text: 'x', style: 'italic' }], [{ kind: 'run', text: '=1+2+3', style: 'italic' }]],
            [[], [{ kind: 'run', text: '=6', style: 'italic' }]],
          ],
        },
      ],
      true,
    );
    expect(ml).toContain('<mtable');
    expect(ml).toContain('columnalign="right left"');
    expect(ml).toContain('<mtr>');
  });

  it('emits mover accent for an accented base', () => {
    const ml = mathToMathML(
      [{ kind: 'accent', char: '^', base: [{ kind: 'run', text: 'x', style: 'italic' }] }],
      false,
    );
    expect(ml).toContain('accent="true"');
  });

  it('escapes XML metacharacters', () => {
    const ml = mathToMathML([{ kind: 'run', text: 'a<b', style: 'italic' }], false);
    expect(ml).toContain('&lt;');
    expect(ml).not.toContain('<mo><</mo>');
  });
});
