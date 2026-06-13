import type { Meta, StoryObj } from '@storybook/html';
import { buildViewerUI } from './DocxViewer.stories';
import { DocxDocument } from './document';
import { DocxViewer } from './viewer';
import type { DocxTextRunInfo } from './renderer';

type DemoArgs = { width: number };
type LayoutArgs = Record<string, never>;

const SAMPLE_URL = `${import.meta.env.BASE_URL}docx/demo/sample-1.docx`;

const meta: Meta<DemoArgs> = {
  title: 'DocxViewer/Examples',
  argTypes: {
    width: {
      control: { type: 'range', min: 400, max: 1200, step: 40 },
      description: 'Canvas render width (px) — used by the Demo story',
    },
  },
  args: { width: 700 },
};
export default meta;

type DemoStory = StoryObj<DemoArgs>;
type LayoutStory = StoryObj<LayoutArgs>;

export const Demo: DemoStory = {
  name: 'Demo — single viewer (demo.docx)',
  render(args) {
    const { root } = buildViewerUI(args, SAMPLE_URL);
    return root;
  },
};

export const Offscreen: DemoStory = {
  name: 'Offscreen — Web Worker rendering (demo.docx)',
  render(args) {
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';
    const heading = document.createElement('h3');
    heading.textContent = 'Offscreen — parsed and rendered in a Web Worker';
    heading.style.cssText = 'margin:0 0 8px;font-size:14px;';
    root.appendChild(heading);
    const status = makeStatus(root);

    const bar = document.createElement('div');
    bar.style.cssText = 'display:flex;align-items:center;gap:8px;margin-bottom:8px;';
    const prev = document.createElement('button');
    prev.textContent = '‹ Prev';
    const next = document.createElement('button');
    next.textContent = 'Next ›';
    const label = document.createElement('span');
    label.style.cssText = 'font-size:13px;color:#444;min-width:96px;text-align:center;';
    bar.append(prev, label, next);
    root.appendChild(bar);

    // mode: 'worker' returns each page as an ImageBitmap; the main thread only
    // paints it through a `bitmaprenderer` context (which consumes the bitmap).
    // The canvas stays hidden until the first frame paints — worker init + WASM
    // + font preload take a moment on cold load, and a pre-sized empty frame
    // would show at the wrong dimensions and jump once the real bitmap arrives.
    const { canvas, ctx, setBusy, reveal } = makeOffscreenStage(root);

    const dpr = window.devicePixelRatio || 1;
    let docu: DocxDocument | null = null;
    let index = 0;
    let busy = false;

    const paint = async () => {
      if (!docu || busy) return;
      busy = true;
      prev.disabled = next.disabled = true;
      setBusy(true);
      status.textContent = `Rendering page ${index + 1} in a Web Worker…`;
      try {
        const bmp = await docu.renderPageToBitmap(index, { width: args.width, dpr });
        canvas.width = bmp.width;
        canvas.height = bmp.height;
        canvas.style.width = `${Math.round(bmp.width / dpr)}px`;
        canvas.style.height = `${Math.round(bmp.height / dpr)}px`;
        ctx.transferFromImageBitmap(bmp);
        reveal();
        label.textContent = `Page ${index + 1} / ${docu.pageCount}`;
        status.textContent = 'Rendered off the main thread — the UI never blocked.';
      } finally {
        busy = false;
        setBusy(false);
        prev.disabled = index === 0;
        next.disabled = !docu || index >= docu.pageCount - 1;
      }
    };

    prev.addEventListener('click', () => { if (index > 0) { index--; void paint(); } });
    next.addEventListener('click', () => { if (docu && index < docu.pageCount - 1) { index++; void paint(); } });

    DocxDocument.load(SAMPLE_URL, { mode: 'worker', useGoogleFonts: true })
      .then((d) => { docu = d; return paint(); })
      .catch((e: Error) => {
        status.textContent = `Error: ${e.message}`;
        status.style.color = 'red';
      });

    return root;
  },
};

function makeStatus(root: HTMLElement): HTMLDivElement {
  const s = document.createElement('div');
  s.style.cssText = 'color:#666;font-size:13px;margin-bottom:8px;min-height:18px;';
  s.textContent = 'Loading…';
  root.appendChild(s);
  return s;
}

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

function buildDocxTextLayer(layer: HTMLDivElement, runs: DocxTextRunInfo[]): void {
  layer.innerHTML = '';
  for (const run of runs) {
    const span = document.createElement('span');
    span.textContent = run.text;
    span.style.cssText =
      `position:absolute;left:${run.x}px;top:${run.y}px;` +
      `font-size:${run.fontSize}px;line-height:${run.h}px;white-space:pre;color:transparent;cursor:text;pointer-events:all;`;
    layer.appendChild(span);
  }
}

export const ScrollView: LayoutStory = {
  name: 'ScrollView — stack all pages',
  render() {
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';
    const heading = document.createElement('h3');
    heading.textContent = 'ScrollView — scroll through every page';
    heading.style.cssText = 'margin:0 0 8px;font-size:14px;';
    root.appendChild(heading);
    const status = makeStatus(root);

    const scroller = document.createElement('div');
    scroller.style.cssText =
      'max-height:720px;overflow-y:auto;border:1px solid #ccc;background:#f5f5f5;padding:12px;';
    root.appendChild(scroller);

    DocxDocument.load(SAMPLE_URL)
      .then(async (doc) => {
        status.textContent = `Rendering ${doc.pageCount} pages…`;
        const widthPx = 700;

        for (let i = 0; i < doc.pageCount; i++) {
          const pageWrapper = document.createElement('div');
          pageWrapper.style.cssText =
            'position:relative;display:block;max-width:700px;margin:0 auto 12px;';

          const canvas = document.createElement('canvas');
          canvas.style.cssText =
            'display:block;width:100%;max-width:700px;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,0.2);';

          const textLayer = document.createElement('div');
          textLayer.style.cssText =
            'position:absolute;top:0;left:0;width:100%;height:100%;' +
            'overflow:hidden;pointer-events:none;user-select:text;-webkit-user-select:text;';

          pageWrapper.appendChild(canvas);
          pageWrapper.appendChild(textLayer);
          scroller.appendChild(pageWrapper);

          const runs: DocxTextRunInfo[] = [];
          await doc.renderPage(canvas, i, { width: widthPx, onTextRun: (r) => runs.push(r) });
          buildDocxTextLayer(textLayer, runs);
        }
        status.textContent = `Loaded ${doc.pageCount} pages`;
      })
      .catch((e: Error) => {
        status.textContent = `Error: ${e.message}`;
        status.style.color = 'red';
      });

    return root;
  },
};

export const ThumbnailGrid: LayoutStory = {
  name: 'ThumbnailGrid — overview of all pages',
  render() {
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';
    const heading = document.createElement('h3');
    heading.textContent = 'ThumbnailGrid — every page at a glance';
    heading.style.cssText = 'margin:0 0 8px;font-size:14px;';
    root.appendChild(heading);
    const status = makeStatus(root);

    const grid = document.createElement('div');
    grid.style.cssText =
      'display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:16px;';
    root.appendChild(grid);

    DocxDocument.load(SAMPLE_URL)
      .then(async (doc) => {
        status.textContent = `Rendering ${doc.pageCount} thumbnails…`;
        const thumbWidth = 160;
        for (let i = 0; i < doc.pageCount; i++) {
          const cell = document.createElement('div');
          cell.style.cssText = 'display:flex;flex-direction:column;align-items:center;cursor:pointer;';
          const canvas = document.createElement('canvas');
          canvas.style.cssText =
            'display:block;width:100%;max-width:160px;background:#fff;' +
            'box-shadow:0 1px 3px rgba(0,0,0,0.2);';
          const caption = document.createElement('div');
          caption.textContent = `Page ${i + 1}`;
          caption.style.cssText = 'font-size:12px;color:#444;margin-top:4px;';
          cell.append(canvas, caption);
          const idx = i;
          cell.addEventListener('click', () => {
            console.log(`[docx ThumbnailGrid] clicked page ${idx + 1}`);
          });
          grid.appendChild(cell);
          await doc.renderPage(canvas, i, { width: thumbWidth });
        }
        status.textContent = `Loaded ${doc.pageCount} pages`;
      })
      .catch((e: Error) => {
        status.textContent = `Error: ${e.message}`;
        status.style.color = 'red';
      });

    return root;
  },
};

export const MasterDetail: LayoutStory = {
  name: 'MasterDetail — thumbnails + large preview',
  render() {
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';
    const heading = document.createElement('h3');
    heading.textContent = 'MasterDetail — click a thumbnail to preview';
    heading.style.cssText = 'margin:0 0 8px;font-size:14px;';
    root.appendChild(heading);
    const status = makeStatus(root);

    const layout = document.createElement('div');
    layout.style.cssText = 'display:flex;gap:16px;height:720px;';
    root.appendChild(layout);

    const thumbCol = document.createElement('div');
    thumbCol.style.cssText =
      'flex:0 0 200px;overflow-y:auto;border:1px solid #ccc;background:#f5f5f5;padding:8px;' +
      'display:flex;flex-direction:column;gap:10px;';
    const detailCol = document.createElement('div');
    detailCol.style.cssText =
      'flex:1 1 auto;border:1px solid #ccc;background:#f5f5f5;padding:12px;overflow:auto;' +
      'display:flex;align-items:flex-start;justify-content:center;';
    layout.append(thumbCol, detailCol);

    // Detail canvas + viewer with text selection
    const detailCanvas = document.createElement('canvas');
    detailCanvas.style.cssText = 'display:block;max-width:100%;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,0.2);';
    detailCol.appendChild(detailCanvas);
    const detailViewer = new DocxViewer(detailCanvas, {
      enableTextSelection: true,
      width: 600,
    });

    // Load thumbnails using DocxDocument; detail viewer loads independently
    Promise.all([
      DocxDocument.load(SAMPLE_URL),
      detailViewer.load(SAMPLE_URL),
    ])
      .then(async ([doc]) => {
        status.textContent = `Rendering ${doc.pageCount} thumbnails…`;
        const thumbEntries: HTMLDivElement[] = [];

        const selectPage = async (i: number) => {
          for (let k = 0; k < thumbEntries.length; k++) {
            thumbEntries[k].style.outline = k === i ? '2px solid #0366d6' : 'none';
          }
          await detailViewer.goToPage(i);
        };

        for (let i = 0; i < doc.pageCount; i++) {
          const cell = document.createElement('div');
          cell.style.cssText = 'display:flex;flex-direction:column;align-items:center;cursor:pointer;padding:4px;';
          const canvas = document.createElement('canvas');
          canvas.style.cssText =
            'display:block;width:100%;max-width:180px;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,0.2);';
          const caption = document.createElement('div');
          caption.textContent = `Page ${i + 1}`;
          caption.style.cssText = 'font-size:12px;color:#444;margin-top:4px;';
          cell.append(canvas, caption);
          const idx = i;
          cell.addEventListener('click', () => {
            selectPage(idx).catch((e: Error) => {
              status.textContent = `Render error: ${e.message}`;
            });
          });
          thumbCol.appendChild(cell);
          thumbEntries.push(cell);
          await doc.renderPage(canvas, i, { width: 180 });
        }

        // Highlight first thumbnail
        if (thumbEntries.length > 0) thumbEntries[0].style.outline = '2px solid #0366d6';
        status.textContent = `Loaded ${doc.pageCount} pages`;
      })
      .catch((e: Error) => {
        status.textContent = `Error: ${e.message}`;
        status.style.color = 'red';
      });

    return root;
  },
};
