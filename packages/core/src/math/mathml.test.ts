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

  // ── §22.1.2.81 phant — the visualization bug: a hidden base must be wrapped ──
  it('wraps a hidden (show=false) phantom in mphantom so its base is invisible', () => {
    const ml = mathToMathML(
      [{ kind: 'phant', show: false, zeroDesc: true, base: [{ kind: 'run', text: 'y', style: 'italic' }] }],
      false,
    );
    // The base is enclosed in <mphantom> (invisible, occupies space) — it must NOT
    // leak as a visible <mi>y</mi> outside a phantom.
    expect(ml).toContain('<mphantom>');
    expect(ml).toContain('<mphantom><mi>y</mi></mphantom>');
    // zeroDesc → mpadded depth="0".
    expect(ml).toContain('depth="0"');
  });

  it('renders a shown (show=true) phantom base normally (spacing-only phant)', () => {
    const ml = mathToMathML(
      [{ kind: 'phant', show: true, base: [{ kind: 'run', text: 'x', style: 'italic' }] }],
      false,
    );
    expect(ml).not.toContain('<mphantom>');
    expect(ml).toContain('<mi>x</mi>');
  });

  // ── §22.1.2.99 sPre — pre-sub-superscript via mmultiscripts + mprescripts ────
  it('emits mmultiscripts with mprescripts for sPre', () => {
    const ml = mathToMathML(
      [
        {
          kind: 'sPre',
          sub: [{ kind: 'run', text: '1', style: 'italic' }],
          sup: [{ kind: 'run', text: '2', style: 'italic' }],
          base: [{ kind: 'run', text: 'A', style: 'italic' }],
        },
      ],
      false,
    );
    expect(ml).toContain('<mmultiscripts>');
    expect(ml).toContain('<mprescripts/>');
    // base precedes the prescripts marker; prescripts follow it (sub then sup).
    expect(ml.indexOf('<mi>A</mi>')).toBeLessThan(ml.indexOf('<mprescripts/>'));
    expect(ml.indexOf('<mprescripts/>')).toBeLessThan(ml.indexOf('<mn>1</mn>'));
  });

  // ── §22.1.2.13 box — logical grouping, no border (mrow) ─────────────────────
  it('renders box as a plain mrow (no border/enclosure)', () => {
    const ml = mathToMathML(
      [{ kind: 'box', base: [{ kind: 'run', text: '=', style: 'roman' }] }],
      false,
    );
    expect(ml).not.toContain('menclose');
    expect(ml).toContain('<mo>=</mo>');
  });

  // ── §22.1.2.11 borderBox — menclose with notation from hide*/strike* ─────────
  it('renders a default borderBox as menclose notation="box"', () => {
    const ml = mathToMathML(
      [{ kind: 'borderBox', base: [{ kind: 'run', text: 'abc', style: 'italic' }] }],
      false,
    );
    expect(ml).toContain('<menclose notation="box">');
  });

  it('builds borderBox notation from surviving edges + strikes', () => {
    // §22.1.2 example: left + bottom edges only (hideTop + hideRight) plus a
    // top-left→bottom-right diagonal strike.
    const ml = mathToMathML(
      [
        {
          kind: 'borderBox',
          hideTop: true,
          hideRight: true,
          strikeTlbr: true,
          base: [{ kind: 'run', text: 'x', style: 'italic' }],
        },
      ],
      false,
    );
    const m = /notation="([^"]*)"/.exec(ml);
    expect(m, 'menclose notation present').not.toBeNull();
    const tokens = (m![1] ?? '').split(' ');
    expect(tokens).toContain('bottom');
    expect(tokens).toContain('left');
    expect(tokens).not.toContain('top');
    expect(tokens).not.toContain('right');
    expect(tokens).not.toContain('box');
    expect(tokens).toContain('downdiagonalstrike');
  });
});
