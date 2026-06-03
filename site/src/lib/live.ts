// Client-side mounts for the three real viewers, used by the Live Showcase.
// PPTX/DOCX render every page/slide stacked on a backdrop (Storybook-style),
// scrolled vertically. XLSX uses its own full viewer. Each mount returns a
// destroy() so the showcase can swap formats cleanly.
import { PptxPresentation } from '@silurus/ooxml-pptx';
import { DocxDocument } from '@silurus/ooxml-docx';
import { XlsxViewer } from '@silurus/ooxml-xlsx';

export type LiveController = { destroy: () => void };

const DPR = () => Math.min(typeof window !== 'undefined' ? window.devicePixelRatio : 1, 2);

function scroller(): HTMLDivElement {
  const s = document.createElement('div');
  s.className = 'lv-scroll';
  return s;
}

function statusLine(text: string): HTMLDivElement {
  const d = document.createElement('div');
  d.className = 'lv-status';
  d.textContent = text;
  return d;
}

export function mountPptx(root: HTMLElement, url: string): LiveController {
  root.innerHTML = '';
  const sc = scroller();
  const status = statusLine('Parsing…');
  sc.appendChild(status);
  root.append(sc);
  let destroyed = false;

  PptxPresentation.load(url, { useGoogleFonts: true })
    .then(async (deck) => {
      if (destroyed) return;
      status.remove();
      for (let i = 0; i < deck.slideCount; i++) {
        if (destroyed) return;
        const canvas = document.createElement('canvas');
        canvas.className = 'lv-page';
        sc.appendChild(canvas);
        await deck.renderSlide(canvas, i, { width: 1280, dpr: DPR() });
      }
    })
    .catch((e: unknown) => { status.textContent = msg(e); });

  return { destroy: () => { destroyed = true; root.innerHTML = ''; } };
}

export function mountDocx(root: HTMLElement, url: string): LiveController {
  root.innerHTML = '';
  const sc = scroller();
  const status = statusLine('Parsing…');
  sc.appendChild(status);
  root.append(sc);
  let destroyed = false;

  DocxDocument.load(url, { useGoogleFonts: true })
    .then(async (doc) => {
      if (destroyed) return;
      status.remove();
      for (let i = 0; i < doc.pageCount; i++) {
        if (destroyed) return;
        const canvas = document.createElement('canvas');
        canvas.className = 'lv-page';
        sc.appendChild(canvas);
        await doc.renderPage(canvas, i, { width: 1000, dpr: DPR() });
      }
    })
    .catch((e: unknown) => { status.textContent = msg(e); });

  return { destroy: () => { destroyed = true; root.innerHTML = ''; } };
}

export function mountXlsx(root: HTMLElement, url: string): LiveController {
  root.innerHTML = '';
  const host = document.createElement('div');
  host.className = 'lv-xlsx';
  root.append(host);

  const viewer = new XlsxViewer(host, {
    useGoogleFonts: true,
    showZoomSlider: true,
    onError: (err: Error) => { host.setAttribute('data-error', err.message); },
  });
  viewer.load(url).catch(() => { /* surfaced via onError */ });

  return { destroy: () => { root.innerHTML = ''; } };
}

function msg(e: unknown): string {
  return `Failed: ${e instanceof Error ? e.message : String(e)}`;
}
