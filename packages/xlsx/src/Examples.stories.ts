import type { Meta, StoryObj } from '@storybook/html';
import { buildViewerUI } from './XlsxViewer.stories';
import { XlsxWorkbook } from './workbook';

type Args = { scale: number };

const meta: Meta<Args> = {
  title: 'XlsxViewer/Examples',
  argTypes: {
    scale: {
      control: { type: 'range', min: 0.25, max: 2, step: 0.05 },
      description: 'Cell/header scale (1 = normal size)',
    },
  },
  args: { scale: 1 },
};
export default meta;
type Story = StoryObj<Args>;

export const Demo: Story = {
  name: 'Demo — single viewer (demo.xlsx)',
  render(args) {
    const { root } = buildViewerUI(args, `${import.meta.env.BASE_URL}xlsx/demo/sample-1.xlsx`);
    return root;
  },
};

const SAMPLE_URL = `${import.meta.env.BASE_URL}xlsx/demo/sample-1.xlsx`;
// Fixed CSS viewport for the offscreen demo: there is no DOM element to measure
// in a worker, so renderViewportToBitmap requires explicit width/height.
const VIEW = { row: 1, col: 1, rows: 40, cols: 12 };
const VIEW_W = 960;
const VIEW_H = 600;

export const Offscreen: Story = {
  name: 'Offscreen — Web Worker rendering (demo.xlsx)',
  render(args) {
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';
    const heading = document.createElement('h3');
    heading.textContent = 'Offscreen — parsed and rendered in a Web Worker';
    heading.style.cssText = 'margin:0 0 8px;font-size:14px;';
    root.appendChild(heading);
    const status = document.createElement('div');
    status.style.cssText = 'color:#666;font-size:13px;margin-bottom:8px;min-height:18px;';
    status.textContent = 'Loading…';
    root.appendChild(status);

    const bar = document.createElement('div');
    bar.style.cssText = 'display:flex;align-items:center;gap:8px;margin-bottom:8px;';
    const prev = document.createElement('button');
    prev.textContent = '‹ Prev';
    const next = document.createElement('button');
    next.textContent = 'Next ›';
    const label = document.createElement('span');
    label.style.cssText = 'font-size:13px;color:#444;min-width:140px;text-align:center;';
    bar.append(prev, label, next);
    root.appendChild(bar);

    // mode: 'worker' returns the sheet viewport as an ImageBitmap; the main
    // thread only paints it through a `bitmaprenderer` context.
    const canvas = document.createElement('canvas');
    canvas.style.cssText = 'display:block;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,0.2);';
    root.appendChild(canvas);
    const ctx = canvas.getContext('bitmaprenderer') as ImageBitmapRenderingContext;

    const dpr = window.devicePixelRatio || 1;
    let wb: XlsxWorkbook | null = null;
    let index = 0;
    let busy = false;

    const paint = async () => {
      if (!wb || busy) return;
      busy = true;
      prev.disabled = next.disabled = true;
      status.textContent = `Rendering "${wb.sheetNames[index]}" in a Web Worker…`;
      try {
        const bmp = await wb.renderViewportToBitmap(index, VIEW, {
          width: VIEW_W,
          height: VIEW_H,
          dpr,
          cellScale: args.scale,
        });
        canvas.width = bmp.width;
        canvas.height = bmp.height;
        canvas.style.width = `${Math.round(bmp.width / dpr)}px`;
        canvas.style.height = `${Math.round(bmp.height / dpr)}px`;
        ctx.transferFromImageBitmap(bmp);
        label.textContent = `${wb.sheetNames[index]} (${index + 1} / ${wb.sheetCount})`;
        status.textContent = 'Rendered off the main thread — the UI never blocked.';
      } finally {
        busy = false;
        prev.disabled = index === 0;
        next.disabled = !wb || index >= wb.sheetCount - 1;
      }
    };

    prev.addEventListener('click', () => { if (index > 0) { index--; void paint(); } });
    next.addEventListener('click', () => { if (wb && index < wb.sheetCount - 1) { index++; void paint(); } });

    XlsxWorkbook.load(SAMPLE_URL, { mode: 'worker', useGoogleFonts: true })
      .then((w) => { wb = w; return paint(); })
      .catch((e: Error) => {
        status.textContent = `Error: ${e.message}`;
        status.style.color = 'red';
      });

    return root;
  },
};
