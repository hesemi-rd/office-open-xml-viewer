import { describe, expect, it } from 'vitest';
import { layoutParseErrorPage } from './error-page.js';
import { createFontResolver } from './font-service.js';
import { createTextLayoutService } from './text.js';
import { paintLayoutPage } from '../paint/canvas-page.js';

describe('parse-error page layout', () => {
  it('wraps detail text during layout and never measures text during paint', async () => {
    const text = createTextLayoutService({
      fonts: createFontResolver([]),
      measurer: {
        fingerprint: 'fixed-width-v1',
        measure: (request) => ({
          advancePt: request.text.length * request.fontSizePt * 0.5,
          ascentPt: request.fontSizePt * 0.8,
          descentPt: request.fontSizePt * 0.2,
        }),
      },
    });
    const layout = layoutParseErrorPage(
      'word/document.xml: a deliberately long parse failure that must wrap before painting',
      { widthPt: 300, heightPt: 400 },
      text,
    );
    const textCommands = layout.pages[0]?.layers.body.flatMap((node) =>
      node.kind === 'drawing' ? node.commands.filter((command) => command.kind === 'text') : [],
    ) ?? [];

    expect(textCommands.length).toBeGreaterThan(3);
    expect(textCommands.map((command) => command.kind === 'text' ? command.text : '').join(' '))
      .toContain('word/document.xml:');

    const calls: string[] = [];
    const ctx = {
      save: () => calls.push('save'),
      restore: () => calls.push('restore'),
      setTransform: () => calls.push('setTransform'),
      clearRect: () => calls.push('clearRect'),
      fillRect: () => calls.push('fillRect'),
      strokeRect: () => calls.push('strokeRect'),
      setLineDash: () => calls.push('setLineDash'),
      fillText: () => calls.push('fillText'),
      measureText: () => { throw new Error('paint measured text'); },
      fillStyle: '',
      strokeStyle: '',
      lineWidth: 1,
      font: '',
      textAlign: 'start',
      textBaseline: 'alphabetic',
    } as unknown as CanvasRenderingContext2D;
    const canvas = {
      width: 0,
      height: 0,
      getContext: () => ctx,
    } as unknown as HTMLCanvasElement;

    await expect(paintLayoutPage(layout, 0, canvas, { scale: 1, dpr: 1 })).resolves.toBeUndefined();
    expect(calls).toContain('fillText');
  });
});
