import type { Meta, StoryObj } from '@storybook/html';
import { XlsxDiffViewer } from './diff-viewer';
import { diffXlsx } from '@silurus/ooxml-diff';

type Args = { width: number };

const meta: Meta<Args> = {
  title: 'XlsxViewer/Diff',
  argTypes: {
    width: { control: { type: 'range', min: 360, max: 1200, step: 40 } },
  },
  args: { width: 520 },
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
    const beforeInput = document.createElement('input'); beforeInput.type = 'file'; beforeInput.accept = '.xlsx';
    const afterInput = document.createElement('input'); afterInput.type = 'file'; afterInput.accept = '.xlsx';
    const beforeLabel = document.createElement('label'); beforeLabel.append('Before: ', beforeInput);
    const afterLabel  = document.createElement('label'); afterLabel.append('After: ', afterInput);
    toolbar.append(beforeLabel, afterLabel);
    root.append(toolbar);

    const stage = document.createElement('div');
    stage.style.cssText = 'display:flex;gap:12px;align-items:flex-start;flex-wrap:wrap;';
    root.append(stage);

    const leftBox = document.createElement('div');
    leftBox.style.cssText = `width:${args.width}px;height:520px;`;
    const rightBox = document.createElement('div');
    rightBox.style.cssText = `width:${args.width}px;height:520px;`;
    stage.append(leftBox, rightBox);

    const changesList = document.createElement('pre');
    changesList.style.cssText =
      'font-size:11px;background:#fafafa;border:1px solid #eee;padding:10px;margin-top:12px;' +
      'max-height:240px;overflow:auto;';
    changesList.textContent = 'Pick two .xlsx files to see the diff.';
    root.append(changesList);

    let viewer: XlsxDiffViewer | null = null;
    let beforeBuf: ArrayBuffer | null = null;
    let afterBuf: ArrayBuffer | null = null;

    const tryLoad = async () => {
      if (!beforeBuf || !afterBuf) return;
      // Clear stage and recreate containers (XlsxViewer takes a container).
      leftBox.innerHTML = '';
      rightBox.innerHTML = '';
      viewer = new XlsxDiffViewer(leftBox, rightBox, {
        onDiff: (result) => {
          if (result.changes.length === 0) {
            changesList.textContent = 'No structural differences detected.';
            return;
          }
          changesList.textContent = result.changes
            .map((c) => {
              const loc = c.location?.kind === 'cell'
                ? `${c.location.sheetName}!R${c.location.row}C${c.location.col}`
                : c.location?.kind === 'sheet'
                  ? c.location.sheetName
                  : c.location?.kind ?? '';
              return `[${c.op.padEnd(6)}] ${loc.padEnd(18)} ${c.kind.padEnd(14)} ${c.path}`;
            })
            .join('\n');
        },
      });
      viewer.setDiffFn(diffXlsx);
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

    return root;
  },
};
