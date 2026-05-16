import type { Meta, StoryObj } from '@storybook/html';
import { DocxDiffViewer } from './diff-viewer';
import { diffDocx } from '@silurus/ooxml-diff';

type Args = { width: number };

const meta: Meta<Args> = {
  title: 'DocxViewer/Diff',
  argTypes: {
    width: { control: { type: 'range', min: 320, max: 1200, step: 20 } },
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
    const beforeInput = document.createElement('input'); beforeInput.type = 'file'; beforeInput.accept = '.docx';
    const afterInput = document.createElement('input'); afterInput.type = 'file'; afterInput.accept = '.docx';
    const prevBtn = document.createElement('button'); prevBtn.textContent = '← Prev'; prevBtn.disabled = true;
    const nextBtn = document.createElement('button'); nextBtn.textContent = 'Next →'; nextBtn.disabled = true;
    const info = document.createElement('span'); info.style.fontSize = '13px';
    const beforeLabel = document.createElement('label'); beforeLabel.append('Before: ', beforeInput);
    const afterLabel  = document.createElement('label'); afterLabel.append('After: ', afterInput);
    toolbar.append(beforeLabel, afterLabel, prevBtn, nextBtn, info);
    root.append(toolbar);

    const stage = document.createElement('div');
    stage.style.cssText = 'display:flex;gap:12px;align-items:flex-start;flex-wrap:wrap;';
    root.append(stage);

    const leftCanvas = document.createElement('canvas');
    const rightCanvas = document.createElement('canvas');
    const leftBox = document.createElement('div');
    leftBox.style.cssText = `width:${args.width}px;border:1px solid #ccc;background:#fff;`;
    leftBox.append(leftCanvas);
    const rightBox = document.createElement('div');
    rightBox.style.cssText = `width:${args.width}px;border:1px solid #ccc;background:#fff;`;
    rightBox.append(rightCanvas);
    stage.append(leftBox, rightBox);

    const changesList = document.createElement('pre');
    changesList.style.cssText =
      'font-size:11px;background:#fafafa;border:1px solid #eee;padding:10px;margin-top:12px;' +
      'max-height:240px;overflow:auto;';
    changesList.textContent = 'Pick two .docx files to see the diff.';
    root.append(changesList);

    let viewer: DocxDiffViewer | null = null;
    let beforeBuf: ArrayBuffer | null = null;
    let afterBuf: ArrayBuffer | null = null;

    const tryLoad = async () => {
      if (!beforeBuf || !afterBuf) return;
      viewer?.destroy();
      viewer = new DocxDiffViewer(leftCanvas, rightCanvas, {
        width: args.width,
        onPageChange: (i, total) => {
          info.textContent = `Page ${i + 1} / ${total}`;
          prevBtn.disabled = i === 0;
          nextBtn.disabled = i === total - 1;
        },
        onDiff: (result) => {
          if (result.changes.length === 0) {
            changesList.textContent = 'No structural differences detected.';
            return;
          }
          changesList.textContent = result.changes
            .map((c) => {
              const loc = c.location?.kind === 'paragraph'
                ? `¶${c.location.paragraphIndex}`
                : c.location?.kind ?? '';
              return `[${c.op.padEnd(6)}] ${loc.padEnd(8)} ${c.kind.padEnd(18)} ${c.path}`;
            })
            .join('\n');
        },
      });
      viewer.setDiffFn(diffDocx);
      try {
        await viewer.load(beforeBuf, afterBuf);
      } catch (err) {
        changesList.textContent = `Error: ${(err as Error).message}`;
      }
    };

    beforeInput.addEventListener('change', async () => {
      const f = beforeInput.files?.[0]; if (!f) return;
      beforeBuf = await f.arrayBuffer(); void tryLoad();
    });
    afterInput.addEventListener('change', async () => {
      const f = afterInput.files?.[0]; if (!f) return;
      afterBuf = await f.arrayBuffer(); void tryLoad();
    });

    prevBtn.addEventListener('click', () => viewer?.prevPage());
    nextBtn.addEventListener('click', () => viewer?.nextPage());

    return root;
  },
};
