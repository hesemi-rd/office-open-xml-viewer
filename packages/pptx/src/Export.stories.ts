import type { Meta, StoryObj } from '@storybook/html';
import { PptxViewer } from './viewer';

type Args = { width: number };

const meta: Meta<Args> = {
  title: 'PptxViewer/Export',
  argTypes: {
    width: { control: { type: 'range', min: 480, max: 1600, step: 40 } },
  },
  args: { width: 960 },
};
export default meta;

type Story = StoryObj<Args>;

export const PngAndPdf: Story = {
  name: 'PNG / PDF export',
  render(args: Args) {
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';

    const toolbar = document.createElement('div');
    toolbar.style.cssText = 'display:flex;gap:8px;align-items:center;margin-bottom:10px;flex-wrap:wrap;';

    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.pptx';

    const exportPngBtn = document.createElement('button');
    exportPngBtn.textContent = 'Export current slide → PNG';
    exportPngBtn.disabled = true;

    const exportAllPngBtn = document.createElement('button');
    exportAllPngBtn.textContent = 'Export all → PNGs (download each)';
    exportAllPngBtn.disabled = true;

    const status = document.createElement('span');
    status.style.cssText = 'font-size:13px;color:#444;';

    const useDemoBtn = document.createElement('button');
    useDemoBtn.textContent = 'Load demo sample-1.pptx';

    toolbar.append(useDemoBtn, fileInput, exportPngBtn, exportAllPngBtn, status);
    root.append(toolbar);

    const container = document.createElement('div');
    container.style.cssText = `width:${args.width}px;border:1px solid #ccc;background:#f0f0f0;`;
    const canvas = document.createElement('canvas');
    container.append(canvas);
    root.append(container);

    const viewer = new PptxViewer(canvas, {
      width: args.width,
      onSlideChange: (i, total) => {
        status.textContent = `Slide ${i + 1} / ${total}`;
        exportPngBtn.disabled = false;
        exportAllPngBtn.disabled = false;
      },
      onError: (err) => { status.textContent = `Error: ${err.message}`; },
    });

    const triggerDownload = (blob: Blob, filename: string) => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    };

    useDemoBtn.addEventListener('click', () => viewer.load('/pptx/demo/sample-1.pptx'));
    fileInput.addEventListener('change', async () => {
      const f = fileInput.files?.[0];
      if (!f) return;
      const buf = await f.arrayBuffer();
      await viewer.load(buf);
    });

    exportPngBtn.addEventListener('click', async () => {
      status.textContent = 'Encoding PNG…';
      const blob = await viewer.exportCurrentSlideToPng();
      triggerDownload(blob, `slide-${viewer.slideIndex + 1}.png`);
      status.textContent = `Slide ${viewer.slideIndex + 1} exported as PNG`;
    });

    exportAllPngBtn.addEventListener('click', async () => {
      status.textContent = 'Rendering all slides…';
      const blobs = await viewer.exportAllSlidesToPng();
      blobs.forEach((b, i) => triggerDownload(b, `slide-${i + 1}.png`));
      status.textContent = `${blobs.length} PNGs queued for download`;
    });

    return root;
  },
};
