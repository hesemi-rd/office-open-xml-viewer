import { describe, it, expect } from 'vitest';
import { evalFormulaToBool } from './formula.js';
import type { Cell } from './types.js';

function numCell(row: number, col: number, n: number): Cell {
  return { row, col, colRef: '', value: { type: 'number', number: n }, styleIndex: 0 };
}

function ctx(opts: { cells?: Cell[]; row?: number; col?: number } = {}) {
  const cellIndex = new Map<string, Cell>();
  for (const c of opts.cells ?? []) cellIndex.set(`${c.row}:${c.col}`, c);
  return {
    row: opts.row ?? 1,
    col: opts.col ?? 1,
    anchorRow: 1,
    anchorCol: 1,
    cellIndex,
    definedNames: new Map(),
    depth: 0,
  };
}

const ev = (f: string, c = ctx()) => evalFormulaToBool(f, c);

describe('evalFormulaToBool — comparisons', () => {
  it('numeric comparisons', () => {
    expect(ev('1>0')).toBe(true);
    expect(ev('1<0')).toBe(false);
    expect(ev('2>=2')).toBe(true);
    expect(ev('2<=1')).toBe(false);
    expect(ev('3=3')).toBe(true);
    expect(ev('3<>3')).toBe(false);
    expect(ev('3<>4')).toBe(true);
  });

  it('arithmetic before comparison', () => {
    expect(ev('2+2=4')).toBe(true);
    expect(ev('10-3*2=4')).toBe(true);
    expect(ev('(10-2)/4=2')).toBe(true);
  });

  it('string comparison', () => {
    expect(ev('"a"="a"')).toBe(true);
    expect(ev('"a"="b"')).toBe(false);
  });
});

describe('evalFormulaToBool — logical functions', () => {
  it('AND / OR / NOT', () => {
    expect(ev('AND(1>0,2>1)')).toBe(true);
    expect(ev('AND(1>0,2>3)')).toBe(false);
    expect(ev('OR(1>2,3>2)')).toBe(true);
    expect(ev('OR(1>2,3>4)')).toBe(false);
    expect(ev('NOT(1>2)')).toBe(true);
    expect(ev('NOT(1>0)')).toBe(false);
  });

  it('IF', () => {
    expect(ev('IF(1>0,1>0,1<0)')).toBe(true);
    expect(ev('IF(1<0,1>0,1<0)')).toBe(false);
  });

  it('nested logic', () => {
    expect(ev('AND(OR(1>2,2>1),NOT(3>4))')).toBe(true);
  });
});

describe('evalFormulaToBool — cell references', () => {
  it('reads a referenced cell (relative shift from anchor)', () => {
    // anchor A1, evaluated at A1 → A1 resolves to the cell at (1,1)
    const c = ctx({ cells: [numCell(1, 1, 90)], row: 1, col: 1 });
    expect(evalFormulaToBool('A1>=90', c)).toBe(true);
    expect(evalFormulaToBool('A1>90', c)).toBe(false);
  });
});
