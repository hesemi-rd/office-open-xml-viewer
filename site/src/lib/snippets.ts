// Code snippets shown in the Code section. Two groups:
//  1. Use-case recipes — what you actually want to build.
//  2. Framework integration — the same viewer in React/Vue/Svelte/Vanilla.
// All snippets are checked against the real public API.

// ── Use-case recipes ───────────────────────────────────────────────

export const embedSnippet = `import { PptxViewer } from '@silurus/ooxml/pptx';

// Drop a single deck into a <canvas> and you're done.
const canvas = document.getElementById('deck') as HTMLCanvasElement;
const viewer = new PptxViewer(canvas, { width: 960, useGoogleFonts: true });

await viewer.load('/quarterly-review.pptx');`;

export const scrollSnippet = `import { PptxPresentation } from '@silurus/ooxml/pptx';

// Render every slide stacked in a scroll container — the headless engine
// draws into any canvas you hand it.
const deck = await PptxPresentation.load('/quarterly-review.pptx');

for (let i = 0; i < deck.slideCount; i++) {
  const canvas = document.createElement('canvas');
  scroller.appendChild(canvas);
  await deck.renderSlide(canvas, i, { width: 1280 });
}`;

export const thumbnailSnippet = `import { PptxPresentation } from '@silurus/ooxml/pptx';

// A clickable thumbnail grid: render each slide small, wire up navigation.
const deck = await PptxPresentation.load('/quarterly-review.pptx');

for (let i = 0; i < deck.slideCount; i++) {
  const thumb = document.createElement('canvas');
  thumb.addEventListener('click', () => openSlide(i));
  grid.appendChild(thumb);
  await deck.renderSlide(thumb, i, { width: 240 });
}`;

export const navSnippet = `import { PptxViewer } from '@silurus/ooxml/pptx';

// Built-in page navigation — the viewer tracks the current slide for you.
const viewer = new PptxViewer(canvas, {
  width: 960,
  onSlideChange: (index, total) => {
    label.textContent = \`\${index + 1} / \${total}\`;
  },
});

await viewer.load('/quarterly-review.pptx');
prevBtn.addEventListener('click', () => viewer.prevSlide());
nextBtn.addEventListener('click', () => viewer.nextSlide());`;

export const fileSnippet = `import { DocxViewer } from '@silurus/ooxml/docx';

// Open a file the user picks — load() takes a URL or an ArrayBuffer.
const viewer = new DocxViewer(canvas, { width: 820 });

input.addEventListener('change', async () => {
  const file = input.files?.[0];
  if (file) await viewer.load(await file.arrayBuffer());
});`;

export const selectionSnippet = `import { XlsxViewer } from '@silurus/ooxml/xlsx';

// Excel-style interaction out of the box: click-drag to select a range,
// Ctrl/Cmd+C copies it as TSV, and the zoom slider rides the tab bar.
const container = document.getElementById('sheet') as HTMLElement;
const viewer = new XlsxViewer(container, {
  showZoomSlider: true,
  onSelectionChange: (range) => console.log(range),
});

await viewer.load('/forecast.xlsx');`;

// ── Framework integration ──────────────────────────────────────────
// PptxViewer is the representative viewer (canvas-based, has destroy());
// DocxViewer/XlsxViewer follow the same shape.

export const vanillaSnippet = `import { PptxViewer } from '@silurus/ooxml/pptx';

const canvas = document.getElementById('deck') as HTMLCanvasElement;
const viewer = new PptxViewer(canvas, { width: 960 });

await viewer.load('/quarterly-review.pptx');
viewer.nextSlide();`;

export const reactSnippet = `import { useEffect, useRef } from 'react';
import { PptxViewer } from '@silurus/ooxml/pptx';

export function Deck({ src }: { src: string }) {
  const ref = useRef<HTMLCanvasElement>(null);

  useEffect(() => {
    const canvas = ref.current;
    if (!canvas) return;
    const viewer = new PptxViewer(canvas, { width: 960 });
    void viewer.load(src);
    return () => viewer.destroy();
  }, [src]);

  return <canvas ref={ref} />;
}`;

export const vueSnippet = `<script setup lang="ts">
import { onMounted, onBeforeUnmount, ref } from 'vue';
import { PptxViewer } from '@silurus/ooxml/pptx';

const props = defineProps<{ src: string }>();
const canvas = ref<HTMLCanvasElement>();
let viewer: PptxViewer | undefined;

onMounted(() => {
  viewer = new PptxViewer(canvas.value as HTMLCanvasElement, { width: 960 });
  void viewer.load(props.src);
});
onBeforeUnmount(() => viewer?.destroy());
</script>

<template>
  <canvas ref="canvas" />
</template>`;

export const svelteSnippet = `<script lang="ts">
  import { onMount } from 'svelte';
  import { PptxViewer } from '@silurus/ooxml/pptx';

  export let src: string;
  let canvas: HTMLCanvasElement;

  onMount(() => {
    const viewer = new PptxViewer(canvas, { width: 960 });
    void viewer.load(src);
    return () => viewer.destroy();
  });
</script>

<canvas bind:this={canvas}></canvas>`;

export const headlessSnippet = `import { PptxPresentation } from '@silurus/ooxml/pptx';

// Headless engine — render any slide into a canvas you control. Build your
// own thumbnail grid, scroll view, or master–detail pane around it.
const deck = await PptxPresentation.load('/quarterly-review.pptx');

for (let i = 0; i < deck.slideCount; i++) {
  const canvas = document.createElement('canvas');
  thumbnails.appendChild(canvas);
  await deck.renderSlide(canvas, i, { width: 240 });
}`;
