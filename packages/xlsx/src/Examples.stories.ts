import type { Meta, StoryObj } from '@storybook/html';
import { buildViewerUI } from './XlsxViewer.stories';

type Args = { scale: number };

const meta: Meta<Args> = {
  title: 'XlsxViewer/Examples',
  // The viewer fills the viewport (height:100vh). Storybook's default story
  // padding would push that past the fold, hiding the bottom sheet-tab bar
  // behind a scroll. 'fullscreen' removes the padding so 100vh fits exactly and
  // the tabs stay visible.
  parameters: { layout: 'fullscreen' },
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

export const Offscreen: Story = {
  name: 'Offscreen — Web Worker rendering (demo.xlsx)',
  // The single-viewer Demo, rendered entirely in a Web Worker (mode: 'worker').
  // Identical UX — scroll, sheet tabs, cell selection — only the pixels are
  // produced off the main thread.
  render(args) {
    const { root } = buildViewerUI(args, SAMPLE_URL, { mode: 'worker' });
    return root;
  },
};
