import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parseMathFont } from './font';
import { parseMathConstants } from './math-table';
import { layoutMath, type MathLayoutCtx } from './layout';
import { renderMathBox, measureMathBox } from './render';
import type { MathNode } from '../types/math';

let ctx: MathLayoutCtx;
beforeAll(() => {
  const url = new URL('../../assets/LatinModernMath.otf', import.meta.url);
  const buf = readFileSync(fileURLToPath(url));
  const font = parseMathFont(
    buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
  );
  ctx = { font, consts: parseMathConstants(font), fontSizePx: 16, level: 'text' };
});

interface RecordingCtx {
  calls: string[];
}

function mockCtx(): CanvasRenderingContext2D & RecordingCtx {
  const calls: string[] = [];
  return {
    calls,
    save() {
      calls.push('save');
    },
    restore() {
      calls.push('restore');
    },
    fillText(t: string, x: number, y: number) {
      calls.push(`text:${t}@${x.toFixed(1)},${y.toFixed(1)}`);
    },
    fillRect(_x: number, _y: number, w: number, h: number) {
      calls.push(`rect:${w.toFixed(1)}x${h.toFixed(1)}`);
    },
    beginPath() {
      calls.push('beginPath');
    },
    moveTo() {},
    lineTo() {},
    stroke() {
      calls.push('stroke');
    },
    set lineWidth(_v: number) {},
    set lineJoin(_v: string) {},
    set lineCap(_v: string) {},
    set strokeStyle(_v: string) {},
    set font(_v: string) {},
    get font() {
      return '';
    },
    set fillStyle(_v: string) {},
    get fillStyle() {
      return '';
    },
  } as unknown as CanvasRenderingContext2D & RecordingCtx;
}

describe('renderMathBox', () => {
  it('draws glyphs and a fraction rule', () => {
    const nodes: MathNode[] = [
      {
        kind: 'fraction',
        num: [{ kind: 'run', text: '1', style: 'italic' }],
        den: [{ kind: 'run', text: 'x', style: 'italic' }],
      },
    ];
    const box = layoutMath(nodes, ctx);
    const m = mockCtx();
    renderMathBox(m, box, 10, 100, '#000', 'LatinModernMath');
    expect(m.calls.some((c) => c.startsWith('text:1'))).toBe(true);
    expect(m.calls.some((c) => c.startsWith('text:x'))).toBe(true);
    expect(m.calls.some((c) => c.startsWith('rect:'))).toBe(true);
    expect(m.calls[0]).toBe('save');
    expect(m.calls[m.calls.length - 1]).toBe('restore');
  });

  it('measureMathBox echoes the box metrics', () => {
    const box = layoutMath([{ kind: 'run', text: 'x', style: 'italic' }], ctx);
    expect(measureMathBox(box)).toEqual({
      width: box.width,
      ascent: box.ascent,
      descent: box.descent,
    });
  });
});
