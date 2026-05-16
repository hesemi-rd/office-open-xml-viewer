import type { Meta, StoryObj } from '@storybook/html';
import { PptxDiffViewer } from './diff-viewer';
import { diffPptx } from '@silurus/ooxml-diff';

type Args = { width: number };

const meta: Meta<Args> = {
  title: 'PptxViewer/Diff',
  argTypes: {
    width: {
      control: { type: 'range', min: 320, max: 1200, step: 20 },
      description: 'Per-side canvas render width (px)',
    },
  },
  args: { width: 480 },
};
export default meta;

type Story = StoryObj<Args>;

export const SideBySide: Story = {
  name: 'Side-by-side diff (file upload)',
  render(args: Args) {
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';

    const toolbar = document.createElement('div');
    toolbar.style.cssText = 'display:flex;gap:8px;align-items:center;margin-bottom:10px;flex-wrap:wrap;';

    const beforeInput = document.createElement('input');
    beforeInput.type = 'file';
    beforeInput.accept = '.pptx';
    const afterInput = document.createElement('input');
    afterInput.type = 'file';
    afterInput.accept = '.pptx';

    const prevBtn = document.createElement('button'); prevBtn.textContent = '← Prev'; prevBtn.disabled = true;
    const nextBtn = document.createElement('button'); nextBtn.textContent = 'Next →'; nextBtn.disabled = true;
    const info = document.createElement('span'); info.style.fontSize = '13px';

    const beforeLabel = document.createElement('label');
    beforeLabel.style.cssText = 'display:flex;gap:4px;align-items:center;font-size:12px;';
    beforeLabel.append('Before:', beforeInput);
    const afterLabel = document.createElement('label');
    afterLabel.style.cssText = 'display:flex;gap:4px;align-items:center;font-size:12px;';
    afterLabel.append('After:', afterInput);

    toolbar.append(beforeLabel, afterLabel, prevBtn, nextBtn, info);
    root.append(toolbar);

    const legend = document.createElement('div');
    legend.style.cssText = 'font-size:12px;color:#444;margin-bottom:8px;';
    legend.innerHTML =
      `<span style="display:inline-block;width:10px;height:10px;background:#ef4444;margin-right:4px;"></span>removed  ` +
      `<span style="display:inline-block;width:10px;height:10px;background:#22c55e;margin-right:4px;"></span>added  ` +
      `<span style="display:inline-block;width:10px;height:10px;background:#f59e0b;margin-right:4px;"></span>modified`;
    root.append(legend);

    const stage = document.createElement('div');
    stage.style.cssText = 'display:flex;gap:12px;align-items:flex-start;flex-wrap:wrap;';
    root.append(stage);

    const leftBox = document.createElement('div');
    leftBox.style.cssText = `width:${args.width}px;border:1px solid #ccc;background:#f0f0f0;`;
    const leftCanvas = document.createElement('canvas');
    leftBox.append(leftCanvas);

    const rightBox = document.createElement('div');
    rightBox.style.cssText = `width:${args.width}px;border:1px solid #ccc;background:#f0f0f0;`;
    const rightCanvas = document.createElement('canvas');
    rightBox.append(rightCanvas);

    stage.append(leftBox, rightBox);

    const changesList = document.createElement('pre');
    changesList.style.cssText =
      'font-size:11px;background:#fafafa;border:1px solid #eee;padding:10px;margin-top:12px;' +
      'max-height:240px;overflow:auto;';
    changesList.textContent = 'Pick two .pptx files to see the diff.';
    root.append(changesList);

    let viewer: PptxDiffViewer | null = null;
    let beforeBuf: ArrayBuffer | null = null;
    let afterBuf: ArrayBuffer | null = null;

    const tryLoad = async () => {
      if (!beforeBuf || !afterBuf) return;
      viewer?.destroy();
      viewer = new PptxDiffViewer(leftCanvas, rightCanvas, {
        width: args.width,
        onSlideChange: (idx, total) => {
          info.textContent = `Slide ${idx + 1} / ${total}`;
          prevBtn.disabled = idx === 0;
          nextBtn.disabled = idx === total - 1;
        },
        onDiff: (result) => {
          if (result.changes.length === 0) {
            changesList.textContent = 'No structural differences detected.';
            return;
          }
          changesList.textContent = result.changes
            .map((c) => {
              const loc = c.location?.kind === 'slide' ? `slide ${c.location.slideIndex + 1}` : c.location?.kind ?? '';
              return `[${c.op.padEnd(6)}] ${loc.padEnd(10)} ${c.kind.padEnd(18)} ${c.path}`;
            })
            .join('\n');
        },
      });
      viewer.setDiffFn(diffPptx);
      try {
        await viewer.load(beforeBuf, afterBuf);
      } catch (err) {
        changesList.textContent = `Error: ${(err as Error).message}`;
      }
    };

    beforeInput.addEventListener('change', async () => {
      const f = beforeInput.files?.[0];
      if (!f) return;
      beforeBuf = await f.arrayBuffer();
      void tryLoad();
    });
    afterInput.addEventListener('change', async () => {
      const f = afterInput.files?.[0];
      if (!f) return;
      afterBuf = await f.arrayBuffer();
      void tryLoad();
    });

    prevBtn.addEventListener('click', () => viewer?.prevSlide());
    nextBtn.addEventListener('click', () => viewer?.nextSlide());

    return root;
  },
};

export const SelfDiff: Story = {
  name: 'Self-diff (sanity check)',
  render(args: Args) {
    const root = document.createElement('div');
    root.style.cssText = 'font-family:sans-serif;padding:16px;';

    const status = document.createElement('div');
    status.style.cssText = 'font-size:13px;margin-bottom:10px;';
    root.append(status);

    const stage = document.createElement('div');
    stage.style.cssText = 'display:flex;gap:12px;align-items:flex-start;';
    root.append(stage);

    const leftCanvas = document.createElement('canvas');
    leftCanvas.style.cssText = `width:${args.width}px;border:1px solid #ccc;`;
    const rightCanvas = document.createElement('canvas');
    rightCanvas.style.cssText = `width:${args.width}px;border:1px solid #ccc;`;
    stage.append(leftCanvas, rightCanvas);

    const viewer = new PptxDiffViewer(leftCanvas, rightCanvas, {
      width: args.width,
      onDiff: (r) => {
        status.textContent = r.changes.length === 0
          ? 'OK — diffing demo/sample-1.pptx against itself produced 0 changes.'
          : `Unexpected: got ${r.changes.length} changes diffing a file against itself.`;
      },
    });
    viewer.setDiffFn(diffPptx);
    viewer.load('/pptx/demo/sample-1.pptx', '/pptx/demo/sample-1.pptx').catch((err) => {
      status.textContent = `Load failed: ${err.message}`;
    });

    return root;
  },
};
