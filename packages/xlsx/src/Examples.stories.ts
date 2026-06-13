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
    // thread only paints it through a `bitmaprenderer` context. The canvas
    // stays hidden until the first frame paints — worker init + WASM + font
    // preload take a moment on cold load, and a pre-sized empty frame would
    // show at the wrong dimensions and jump once the real bitmap arrives.
    const { canvas, ctx, setBusy, reveal } = makeOffscreenStage(root);

    const dpr = window.devicePixelRatio || 1;
    let wb: XlsxWorkbook | null = null;
    let index = 0;
    let busy = false;

    const paint = async () => {
      if (!wb || busy) return;
      busy = true;
      prev.disabled = next.disabled = true;
      setBusy(true);
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
        reveal();
        label.textContent = `${wb.sheetNames[index]} (${index + 1} / ${wb.sheetCount})`;
        status.textContent = 'Rendered off the main thread — the UI never blocked.';
      } finally {
        busy = false;
        setBusy(false);
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

/** A render stage for the Offscreen story: a centred area with an indeterminate
 *  spinner overlay and a `bitmaprenderer` canvas that stays hidden until the
 *  first frame is painted (so no wrong-sized empty frame flashes during the
 *  cold-load delay). `setBusy` toggles the spinner; `reveal` shows the canvas. */
function makeOffscreenStage(root: HTMLElement): {
  canvas: HTMLCanvasElement;
  ctx: ImageBitmapRenderingContext;
  setBusy: (busy: boolean) => void;
  reveal: () => void;
} {
  if (!document.getElementById('ooxml-offscreen-spin-kf')) {
    const style = document.createElement('style');
    style.id = 'ooxml-offscreen-spin-kf';
    style.textContent = '@keyframes ooxml-offscreen-spin{to{transform:rotate(360deg)}}';
    document.head.appendChild(style);
  }
  const stage = document.createElement('div');
  stage.style.cssText =
    'position:relative;min-height:200px;display:flex;align-items:flex-start;justify-content:center;';
  const canvas = document.createElement('canvas');
  canvas.style.cssText = 'display:none;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,0.2);';
  const spinWrap = document.createElement('div');
  // Visible from mount so the spinner covers the slow cold load (worker init +
  // WASM + parse), not just the per-frame render; hidden after the first paint.
  spinWrap.style.cssText =
    'position:absolute;inset:0;display:flex;align-items:center;justify-content:center;pointer-events:none;';
  const spinner = document.createElement('div');
  spinner.style.cssText =
    'width:36px;height:36px;border:4px solid rgba(0,0,0,0.12);border-top-color:#0366d6;' +
    'border-radius:50%;animation:ooxml-offscreen-spin 0.8s linear infinite;';
  spinWrap.appendChild(spinner);
  stage.append(canvas, spinWrap);
  root.appendChild(stage);
  const ctx = canvas.getContext('bitmaprenderer') as ImageBitmapRenderingContext;
  return {
    canvas,
    ctx,
    setBusy: (busy: boolean) => { spinWrap.style.display = busy ? 'flex' : 'none'; },
    reveal: () => { canvas.style.display = 'block'; },
  };
}
