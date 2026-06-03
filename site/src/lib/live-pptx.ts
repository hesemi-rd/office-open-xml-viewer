// Client-side mount for a live PPTX viewer (tech-validation of WASM/worker in Astro).
import { PptxViewer } from '@silurus/ooxml-pptx';

export function mountPptx(root: HTMLElement, url: string, width = 880): void {
  const container = document.createElement('div');
  container.className = 'lv-stage';

  const canvas = document.createElement('canvas');
  container.appendChild(canvas);

  const bar = document.createElement('div');
  bar.className = 'lv-bar';
  const prev = document.createElement('button');
  prev.className = 'lv-btn';
  prev.textContent = '‹';
  prev.disabled = true;
  const next = document.createElement('button');
  next.className = 'lv-btn';
  next.textContent = '›';
  next.disabled = true;
  const info = document.createElement('span');
  info.className = 'lv-info';
  info.textContent = 'Loading…';
  bar.append(prev, info, next);

  root.append(bar, container);

  const viewer = new PptxViewer(canvas, {
    width,
    useGoogleFonts: true,
    enableTextSelection: true,
    onSlideChange: (idx, total) => {
      info.textContent = `${idx + 1} / ${total}`;
      prev.disabled = idx === 0;
      next.disabled = idx === total - 1;
    },
    onError: (err) => { info.textContent = `Error: ${err.message}`; },
  });

  prev.addEventListener('click', () => void viewer.prevSlide());
  next.addEventListener('click', () => void viewer.nextSlide());

  viewer.load(url).catch((err: unknown) => {
    info.textContent = `Failed: ${err instanceof Error ? err.message : String(err)}`;
  });
}
