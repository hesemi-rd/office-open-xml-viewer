/**
 * Browser (Chromium) regression for issue #1014 — a vo=Tr rotate-fallback mark
 * (ー prolonged sound mark) whose SUBSTITUTE font under-reports its advance via
 * `measureText` draws ink that overruns its advance-sized cell and overlaps the
 * following sideways run. Reproduces in Chrome but NOT in headless skia (skia
 * reports a full-em advance for ー), so the skia/Chrome parity gap is itself part
 * of the bug and the guard must run in a real Chromium — via Playwright here.
 *
 * No installed font on a dev host under-reports ー (every one reports advance ≥
 * ink), so the trigger is a MINIMAL synthetic font (FONT_B64 below): U+30FC has
 * an hmtx advance of 500 units (0.5em) but a glyf outline spanning x∈[40,960]
 * (~0.96em) — exactly the "advance under-reports the ink" condition the issue
 * names. The Latin letters of "controls" advance 560 with ink inside, and 話 is a
 * full-em kanji, so the run geometry mirrors the reported "…ー controls" case.
 *
 * The test drives the REAL packages/docx `drawVerticalRun` and
 * `verticalRunInkExtraPx` (esbuild-bundled to browser ESM) inside the +90° page
 * frame the renderer installs, and asserts:
 *   1. the synthetic font really under-reports (measureText advance < ink extent);
 *   2. single run: the ー ink is sized into its own (ink-extent) cell and does NOT
 *      overlap the following sideways Latin run;
 *   3. two segments (measure==paint): when the next segment is positioned by the
 *      MEASURED advance — which folds in `verticalRunInkExtraPx` — the marks clear,
 *      and WITHOUT that extra they would still overlap (the fix is load-bearing).
 *
 * Gated on Playwright Chromium being present; skips cleanly otherwise.
 */
import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, '../../..');
const verticalTextEntry = resolve(repoRoot, 'packages/docx/src/vertical-text.ts');

// Synthetic under-reporting font (family "IQ1014Test"). See the file header.
const FONT_B64 = 'AAEAAAAKAIAAAwAgT1MvMkUA1kYAAAEoAAAAYGNtYXAyioyWAAABtAAAAG5nbHlmh2Y1lwAAAjwAAADqaGVhZC/wfuYAAACsAAAANmhoZWEHCwFgAAAA5AAAACRobXR4GEwCHAAAAYgAAAAsbG9jYQEEAUUAAAIkAAAAGG1heHAADQAGAAABCAAAACBuYW1lu6hWtwAAAygAAAC3cG9zdIcOa74AAAPgAAAAgAABAAAAAQAAgZbZ8F8PPPUAAwPoAAAAAOZ5HXMAAAAA5nkdcwAoAAADwAMgAAAAAwACAAAAAAAAAAEAAAMg/zgAAAPoACj+NAPAAAEAAAAAAAAAAAAAAAAAAAALAAEAAAALAAQAAQAAAAAAAgAAAAAAAAAAAAAAAAAAAAAAAwI1AZAABQAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABCAQAAAAAAAAAAAAAPz8/PwAAACCKcQMg/zgAAAMgAMgAAAAAAAAAAAAAAAAAAAAgAAAB9AAAASwAAAH0ACgCMAA8AjAAPAIwADwCMAA8AjAAPAIwADwCMAA8A+gAUAAAAAIAAAADAAAAFAADAAEAAAAUAAQAWgAAABAAEAADAAAAIABjAGwAbwB0MPyKcf//AAAAIABjAGwAbgByMPyKcf///+H/oP+cAAAAAM8GdZkAAQAAAAAAAAAKAAwAAAAAAAAABQAEAAcACQAGAAAAAAAAAAAADQAaACcANABBAE4AWwBoAHUAAQAoAbgDwAIIAAMAABMhNSEoA5j8aAG4UAABADwAAAH0ArwAAwAAMyERITwBuP5IArwAAAEAPAAAAfQCvAADAAAzIREhPAG4/kgCvAAAAQA8AAAB9AK8AAMAADMhESE8Abj+SAK8AAABADwAAAH0ArwAAwAAMyERITwBuP5IArwAAAEAPAAAAfQCvAADAAAzIREhPAG4/kgCvAAAAQA8AAAB9AK8AAMAADMhESE8Abj+SAK8AAABADwAAAH0ArwAAwAAMyERITwBuP5IArwAAAEAUAAAA5gDIAADAAAzIREhUANI/LgDIAAAAAAAAAYATgABAAAAAAABAAoAAAABAAAAAAACAAcACgABAAAAAAAGABIAEQADAAEECQABABQAIwADAAEECQACAA4ANwADAAEECQAGACQARUlRMTAxNFRlc3RSZWd1bGFySVExMDE0VGVzdC1SZWd1bGFyAEkAUQAxADAAMQA0AFQAZQBzAHQAUgBlAGcAdQBsAGEAcgBJAFEAMQAwADEANABUAGUAcwB0AC0AUgBlAGcAdQBsAGEAcgAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALAAAAAwECAQMBBAEFAQYBBwEIAQkBCglwcm9sb25nZWQHbGF0aW5fYwdsYXRpbl9vB2xhdGluX24HbGF0aW5fdAdsYXRpbl9yB2xhdGluX2wHbGF0aW5fcwVrYW5qaQ==';

// `playwright` and `esbuild` are ROOT devDependencies (used only by this optional,
// Chromium-gated probe), NOT declared deps of @silurus/ooxml-node, so a static
// import specifier trips `tsc --noEmit` with TS2307 in the per-package typecheck.
// Import them through a runtime variable specifier: TS then types the result as
// `Promise<any>` and does not resolve the path, while Node resolves the real module
// at runtime (from the hoisted root node_modules). The suite skips when absent.
const PLAYWRIGHT = 'playwright';
const ESBUILD = 'esbuild';

/** Resolve Playwright chromium; return null (→ skip) when it or its browser is absent. */
async function loadChromium(): Promise<Any | null> {
  try {
    const pw = (await import(PLAYWRIGHT)) as Any;
    const exe = pw.chromium?.executablePath?.();
    if (typeof exe === 'string' && existsSync(exe)) return pw.chromium;
  } catch {
    /* playwright not installed */
  }
  return null;
}

/** esbuild-bundle the real drawVerticalRun + verticalRunInkExtraPx into a browser
 *  IIFE that assigns the module's exports to `window.__vt` (a classic <script>, so
 *  no dynamic `import()` — which vitest's SSR transform would rewrite). */
async function bundleVerticalText(): Promise<string> {
  const esbuild = (await import(ESBUILD)) as Any;
  const out = await esbuild.build({
    entryPoints: [verticalTextEntry],
    bundle: true,
    format: 'iife',
    globalName: '__vt',
    platform: 'browser',
    target: 'es2020',
    write: false,
    absWorkingDir: repoRoot,
  });
  return out.outputFiles[0].text as string;
}

const chromium = await loadChromium();

describe.skipIf(!chromium)('docx vertical Tr rotate-fallback ink overrun (#1014, Chromium)', () => {
  const fontPx = 100;
  let browser: Any;
  let page: Any;

  beforeAll(async () => {
    const bundle = await bundleVerticalText();
    browser = await chromium.launch();
    page = await browser.newPage();
    await page.setContent('<!doctype html><canvas id="c" width="600" height="1400"></canvas>');
    await page.evaluate(async (b64: string) => {
      const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
      const ff = new (window as Any).FontFace('IQ1014Test', bytes.buffer);
      await ff.load();
      (document as Any).fonts.add(ff);
    }, FONT_B64);
    // Classic script — the IIFE assigns the exports to `window.__vt`.
    await page.addScriptTag({ content: bundle });
  }, 60_000);

  afterAll(async () => {
    await browser?.close();
  });

  it('the synthetic font under-reports the ー advance (the Chrome-only trigger)', async () => {
    const m = await page.evaluate((px: number) => {
      const ctx = (document.getElementById('c') as HTMLCanvasElement).getContext('2d')!;
      ctx.font = px + 'px IQ1014Test';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      const tm = ctx.measureText('ー');
      return {
        advance: tm.width,
        inkExtent: tm.actualBoundingBoxLeft + tm.actualBoundingBoxRight,
      };
    }, fontPx);
    // The along-column ink extent (left+right) exceeds the advance — the report's
    // under-report. A real font gives ink ≤ advance, making the fix a no-op there.
    expect(m.inkExtent).toBeGreaterThan(m.advance);
    expect(m.advance).toBeCloseTo(fontPx * 0.5, 0);
  });

  it('single run: the ー ink stays in its ink-sized cell and does not overlap the next sideways run', async () => {
    const r = await page.evaluate((px: number) => {
      const W = 600, H = 1400, logX = 60, baseline = 300;
      const { drawVerticalRun } = (window as Any).__vt;
      const render = (text: string) => {
        const cv = document.createElement('canvas'); cv.width = W; cv.height = H;
        const ctx = cv.getContext('2d')!;
        ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, H);
        ctx.fillStyle = '#000'; ctx.font = px + 'px IQ1014Test';
        ctx.save(); ctx.translate(W, 0); ctx.rotate(Math.PI / 2);
        drawVerticalRun(ctx, text, logX, baseline, px, 0, 1, true);
        ctx.restore();
        return ctx.getImageData(0, 0, W, H).data;
      };
      const dark = (d: Uint8ClampedArray, i: number) =>
        0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2] < 128;
      // Physical-y range of ink present in `more` but not `base` (isolates the extra glyphs).
      const diffY = (base: Uint8ClampedArray, more: Uint8ClampedArray) => {
        let y0 = H, y1 = -1;
        for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
          const i = (y * W + x) * 4;
          if (dark(more, i) && !dark(base, i)) { if (y < y0) y0 = y; if (y > y1) y1 = y; break; }
        }
        return { y0, y1 };
      };
      const dashInk = diffY(render('話'), render('話ー'));                 // ー's along-column ink
      const controlsInk = diffY(render('話ー '), render('話ー controls')); // controls' ink (c first)
      return { dashMaxY: dashInk.y1, controlsMinY: controlsInk.y0 };
    }, fontPx);
    // The ー ink must end (max physical-y) strictly BEFORE the following run's ink begins.
    expect(r.dashMaxY).toBeLessThan(r.controlsMinY);
  });

  it('two segments (measure==paint): the next segment positioned by the measured advance clears the ー ink', async () => {
    const r = await page.evaluate((px: number) => {
      const W = 600, H = 1400, logX = 60, baseline = 300;
      const { drawVerticalRun, verticalRunInkExtraPx } = (window as Any).__vt;
      const ctx0 = (document.getElementById('c') as HTMLCanvasElement).getContext('2d')!;
      ctx0.font = px + 'px IQ1014Test';
      const seg1 = '話ー';
      const extra = verticalRunInkExtraPx(ctx0, seg1);
      const seg1NaturalW = ctx0.measureText(seg1).width; // whole-string advance (under-reports)
      const seg2OriginNoFix = logX + seg1NaturalW;        // where controls would start WITHOUT the fix
      const seg2Origin = logX + seg1NaturalW + extra;     // WITH the fix (measure folds in the deficit)

      const render = (draw: (ctx: CanvasRenderingContext2D) => void) => {
        const cv = document.createElement('canvas'); cv.width = W; cv.height = H;
        const ctx = cv.getContext('2d')!;
        ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, H);
        ctx.fillStyle = '#000'; ctx.font = px + 'px IQ1014Test';
        ctx.save(); ctx.translate(W, 0); ctx.rotate(Math.PI / 2); draw(ctx); ctx.restore();
        return ctx.getImageData(0, 0, W, H).data;
      };
      const dark = (d: Uint8ClampedArray, i: number) =>
        0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2] < 128;
      const maxDarkY = (d: Uint8ClampedArray) => {
        let y1 = -1;
        for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) { if (dark(d, (y * W + x) * 4)) { if (y > y1) y1 = y; break; } }
        return y1;
      };
      const minDarkY = (d: Uint8ClampedArray) => {
        for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) { if (dark(d, (y * W + x) * 4)) return y; }
        return H;
      };

      const dashMaxY = maxDarkY(render((ctx) => drawVerticalRun(ctx, seg1, logX, baseline, px, 0, 1, true)));
      const controlsMinYFix = minDarkY(render((ctx) => drawVerticalRun(ctx, 'controls', seg2Origin, baseline, px, 0, 1, true)));
      const controlsMinYNoFix = minDarkY(render((ctx) => drawVerticalRun(ctx, 'controls', seg2OriginNoFix, baseline, px, 0, 1, true)));
      return { extra, dashMaxY, controlsMinYFix, controlsMinYNoFix };
    }, fontPx);
    // The fix folds a positive deficit into the layout advance…
    expect(r.extra).toBeGreaterThan(0);
    // …so the second segment's ink clears the ー ink (no overlap)…
    expect(r.dashMaxY).toBeLessThan(r.controlsMinYFix);
    // …and WITHOUT that extra the second segment would still be overlapped by the ー
    // ink (proving the measure-side correction is load-bearing, not cosmetic).
    expect(r.dashMaxY).toBeGreaterThan(r.controlsMinYNoFix);
  });
});
