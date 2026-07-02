import { describe, it, expect, vi, afterEach } from 'vitest';
import { buildDocxTextLayer } from './text-layer.js';
import type { DocxTextRunInfo } from './renderer';

// The vitest env is `node` (no document). Following renderer.textbox-image.test.ts's
// vi.stubGlobal pattern, we stub a minimal recording DOM: createElement returns a
// fake element that records its style props (via a cssText setter that parses
// `k:v;` pairs), textContent, and appended children. buildDocxTextLayer lays
// absolutely-positioned <span>s directly on the layer, one per DocxTextRunInfo.

interface FakeEl {
  tag: string;
  textContent: string;
  innerHTML: string;
  style: Record<string, string> & { cssText: string };
  children: FakeEl[];
  appendChild(c: FakeEl): void;
}

function makeEl(tag: string): FakeEl {
  const style: Record<string, string> = {};
  const el: FakeEl = {
    tag,
    textContent: '',
    innerHTML: '',
    children: [],
    // A cssText setter that parses the `k:v;k:v;` string into individual props,
    // so tests can assert either the raw cssText or a parsed key.
    style: new Proxy(style as Record<string, string> & { cssText: string }, {
      set(target, prop: string, value: string) {
        if (prop === 'cssText') {
          for (const decl of value.split(';')) {
            const idx = decl.indexOf(':');
            if (idx > 0) target[decl.slice(0, idx).trim()] = decl.slice(idx + 1).trim();
          }
          target.cssText = value;
        } else {
          target[prop] = value;
        }
        return true;
      },
    }),
    appendChild(c: FakeEl) {
      this.children.push(c);
    },
  };
  return el;
}

afterEach(() => vi.unstubAllGlobals());

function run(partial: Partial<DocxTextRunInfo>): DocxTextRunInfo {
  return {
    text: 'X',
    x: 0,
    y: 0,
    w: 10,
    h: 12,
    fontSize: 12,
    font: '12px serif',
    ...partial,
  };
}

describe('buildDocxTextLayer (extracted from DocxViewer._buildTextLayer)', () => {
  it('sizes the layer to the passed canvas dimensions and clears prior content', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    layer.innerHTML = 'STALE';
    buildDocxTextLayer(layer as unknown as HTMLDivElement, [], '640px', '480px');
    expect(layer.innerHTML).toBe(''); // cleared
    expect(layer.style.width).toBe('640px');
    expect(layer.style.height).toBe('480px');
    expect(layer.children.length).toBe(0);
  });

  it('lays one absolutely-positioned span per run with the shipped inline styles', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    const runs = [
      run({ text: 'Hello', x: 12, y: 34, h: 16, font: 'bold 14px Arial' }),
      run({ text: 'World', x: 50, y: 34, h: 16, font: '14px Arial' }),
    ];
    buildDocxTextLayer(layer as unknown as HTMLDivElement, runs, '700px', '900px');

    expect(layer.children.length).toBe(2);
    const [a, b] = layer.children;
    expect(a.tag).toBe('span');
    expect(a.textContent).toBe('Hello');
    // Shipped style contract: absolute position from run.x/run.y, the CSS `font`
    // shorthand BEFORE line-height, letter-spacing reset, transparent selectable text.
    expect(a.style.position).toBe('absolute');
    expect(a.style.left).toBe('12px');
    expect(a.style.top).toBe('34px');
    expect(a.style.font).toBe('bold 14px Arial');
    expect(a.style['line-height']).toBe('16px');
    expect(a.style['letter-spacing']).toBe('0');
    expect(a.style['white-space']).toBe('pre');
    expect(a.style.color).toBe('transparent');
    expect(a.style.cursor).toBe('text');
    expect(a.style['pointer-events']).toBe('all');
    // Declaration ORDER is load-bearing: the `font` shorthand resets line-height
    // to `normal` in real browsers, so `line-height` must come AFTER `font` in the
    // cssText. The parsed-prop asserts above cannot detect a reorder — pin it on
    // the raw string.
    expect(a.style.cssText).toMatch(/font:[^;]+;line-height:/);
    expect(b.textContent).toBe('World');
    expect(b.style.left).toBe('50px');
  });
});
