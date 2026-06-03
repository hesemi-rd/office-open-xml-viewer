// Implementation snippets shown beside each per-format live demo. Generated
// from a small config so pptx/docx stay in sync; all checked against the API.

interface Cfg {
  Viewer: string;
  Doc: string;
  sub: 'pptx' | 'docx';
  count: string;
  render: string;
  next: string;
  prev: string;
  go: string;
}

const pptx: Cfg = {
  Viewer: 'PptxViewer', Doc: 'PptxPresentation', sub: 'pptx',
  count: 'slideCount', render: 'renderSlide', next: 'nextSlide', prev: 'prevSlide', go: 'goToSlide',
};
const docx: Cfg = {
  Viewer: 'DocxViewer', Doc: 'DocxDocument', sub: 'docx',
  count: 'pageCount', render: 'renderPage', next: 'nextPage', prev: 'prevPage', go: 'goToPage',
};

export interface DemoSnippets {
  demo: string;
  scroll: string;
  thumbnails: string;
  masterdetail: string;
}

function build(c: Cfg): DemoSnippets {
  return {
    demo: `import { ${c.Viewer} } from '@silurus/ooxml/${c.sub}';

// The built-in viewer tracks the current ${c.sub === 'pptx' ? 'slide' : 'page'} for you.
const viewer = new ${c.Viewer}(canvas, { width: 960, useGoogleFonts: true });
await viewer.load('/sample.${c.sub}');

nextBtn.addEventListener('click', () => viewer.${c.next}());
prevBtn.addEventListener('click', () => viewer.${c.prev}());`,

    scroll: `import { ${c.Doc} } from '@silurus/ooxml/${c.sub}';

// Headless engine — render every ${c.sub === 'pptx' ? 'slide' : 'page'} into a canvas you control.
const doc = await ${c.Doc}.load('/sample.${c.sub}');

for (let i = 0; i < doc.${c.count}; i++) {
  const canvas = document.createElement('canvas');
  scroller.appendChild(canvas);
  await doc.${c.render}(canvas, i, { width: 1100 });
}`,

    thumbnails: `import { ${c.Doc} } from '@silurus/ooxml/${c.sub}';

// Render each ${c.sub === 'pptx' ? 'slide' : 'page'} small, wire up navigation.
const doc = await ${c.Doc}.load('/sample.${c.sub}');

for (let i = 0; i < doc.${c.count}; i++) {
  const thumb = document.createElement('canvas');
  thumb.addEventListener('click', () => open(i));
  grid.appendChild(thumb);
  await doc.${c.render}(thumb, i, { width: 320 });
}`,

    masterdetail: `import { ${c.Doc}, ${c.Viewer} } from '@silurus/ooxml/${c.sub}';

// A large preview viewer on the right…
const viewer = new ${c.Viewer}(detailCanvas, { width: 820, enableTextSelection: true });

// …and a thumbnail rail on the left, sharing the same file.
const [doc] = await Promise.all([
  ${c.Doc}.load('/sample.${c.sub}'),
  viewer.load('/sample.${c.sub}'),
]);

for (let i = 0; i < doc.${c.count}; i++) {
  const thumb = document.createElement('canvas');
  thumb.addEventListener('click', () => viewer.${c.go}(i));  // jump the preview
  rail.appendChild(thumb);
  await doc.${c.render}(thumb, i, { width: 200 });
}`,
  };
}

export const pptxSnippets = build(pptx);
export const docxSnippets = build(docx);

export const xlsxSheetSnippet = `import { XlsxViewer } from '@silurus/ooxml/xlsx';

// XlsxViewer owns its canvas, sheet-tab bar and zoom slider — hand it a
// container element (not a canvas). Click-drag selects a range; Ctrl/Cmd+C
// copies it as TSV.
const container = document.getElementById('sheet') as HTMLElement;
const viewer = new XlsxViewer(container, { showZoomSlider: true });

await viewer.load('/sample.xlsx');`;

// ── Framework integration (per format) ─────────────────────────────
// docx/pptx take a <canvas>; xlsx takes a container <div>. pptx & xlsx expose
// destroy(); docx renders into the canvas you own and needs no teardown.
interface FwCfg {
  Viewer: string;
  sub: 'docx' | 'xlsx' | 'pptx';
  el: 'canvas' | 'container';
  tag: 'canvas' | 'div';
  RefType: 'HTMLCanvasElement' | 'HTMLDivElement';
  opts: string;
  destroy: boolean;
}

const fwPptx: FwCfg = { Viewer: 'PptxViewer', sub: 'pptx', el: 'canvas', tag: 'canvas', RefType: 'HTMLCanvasElement', opts: '{ width: 960 }', destroy: true };
const fwDocx: FwCfg = { Viewer: 'DocxViewer', sub: 'docx', el: 'canvas', tag: 'canvas', RefType: 'HTMLCanvasElement', opts: '{ width: 820 }', destroy: false };
const fwXlsx: FwCfg = { Viewer: 'XlsxViewer', sub: 'xlsx', el: 'container', tag: 'div', RefType: 'HTMLDivElement', opts: '{ showZoomSlider: true }', destroy: true };

export interface FrameworkSnippets {
  react: string;
  vue: string;
  svelte: string;
  vanilla: string;
}

function buildFw(c: FwCfg): FrameworkSnippets {
  const reactCleanup = c.destroy
    ? '    return () => viewer.destroy();'
    : `    // ${c.Viewer} renders into the ${c.el} you own — nothing to tear down.`;
  const svelteCleanup = c.destroy ? '\n    return () => viewer.destroy();' : '';
  const vueUnmount = c.destroy
    ? `onBeforeUnmount(() => viewer?.destroy());`
    : `// ${c.Viewer} needs no explicit teardown.`;

  return {
    react: `import { useEffect, useRef } from 'react';
import { ${c.Viewer} } from '@silurus/ooxml/${c.sub}';

export function Viewer({ src }: { src: string }) {
  const ref = useRef<${c.RefType}>(null);

  useEffect(() => {
    const ${c.el} = ref.current;
    if (!${c.el}) return;
    const viewer = new ${c.Viewer}(${c.el}, ${c.opts});
    void viewer.load(src);
${reactCleanup}
  }, [src]);

  return <${c.tag} ref={ref} />;
}`,

    vue: `<script setup lang="ts">
import { onMounted${c.destroy ? ', onBeforeUnmount' : ''}, ref } from 'vue';
import { ${c.Viewer} } from '@silurus/ooxml/${c.sub}';

const props = defineProps<{ src: string }>();
const ${c.el} = ref<${c.RefType}>();
let viewer: ${c.Viewer} | undefined;

onMounted(() => {
  viewer = new ${c.Viewer}(${c.el}.value as ${c.RefType}, ${c.opts});
  void viewer.load(props.src);
});
${vueUnmount}
<\/script>

<template>
  <${c.tag} ref="${c.el}" />
</template>`,

    svelte: `<script lang="ts">
  import { onMount } from 'svelte';
  import { ${c.Viewer} } from '@silurus/ooxml/${c.sub}';

  export let src: string;
  let ${c.el}: ${c.RefType};

  onMount(() => {
    const viewer = new ${c.Viewer}(${c.el}, ${c.opts});
    void viewer.load(src);${svelteCleanup}
  });
<\/script>

<${c.tag} bind:this={${c.el}}></${c.tag}>`,

    vanilla: `import { ${c.Viewer} } from '@silurus/ooxml/${c.sub}';

const ${c.el} = document.getElementById('viewer') as ${c.RefType};
const viewer = new ${c.Viewer}(${c.el}, ${c.opts});

await viewer.load('/sample.${c.sub}');`,
  };
}

export const frameworkSnippets = {
  docx: buildFw(fwDocx),
  xlsx: buildFw(fwXlsx),
  pptx: buildFw(fwPptx),
};
