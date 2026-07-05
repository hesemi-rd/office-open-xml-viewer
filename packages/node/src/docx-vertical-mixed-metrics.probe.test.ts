/**
 * Cross-axis centering probe for docx vertical (tbRl) text (ECMA-376 §17.6.20).
 *
 * Four symptoms reported on sample-26 (the vertical newspaper):
 *   1. In a mixed column, the SIDEWAYS (Latin/digit) glyphs of "03-1234-5678"
 *      do not share the column centerline with the UPRIGHT "電話" glyphs — the
 *      digits sit off to one physical side.
 *   2. The Tr rotated fullwidth parens （）of "…日（土）" crowd the neighbouring
 *      glyph instead of sitting a full cell apart.
 *   3. The Tu comma / full stop 、。 sit TOO LOW in their cell (along-column). Word
 *      hangs them in the cell's UPPER-RIGHT corner (JIS X 4051 §4.3, PDF-verified:
 *      、 ink centroid at ≈ −0.32em along-column). The #792 along-column ink-centring
 *      force-pulled the designed corner ink back to the geometric cell centre,
 *      dropping them low; the fix skips ink-centring for the FE10–FE12 punctuation
 *      substitutes so the font's corner design stands.
 *   4. The ！ (and ？) sit shifted to the RIGHT of the column (cross-axis). Word
 *      centres them (PDF-verified: ！ ink at ≈ +0.03em cross-axis). The FE15/FE16
 *      "vertical form" substitution was corner-designed in the render font and
 *      pushed the mark right; the fix stops substituting ！／？ and draws the
 *      original fullwidth mark upright, which is centred in every font.
 *
 * This probe measures glyph INK CENTROIDS directly on the rendered canvas — the
 * cross-axis (physical x) centroid for symptoms 1 & 4, and the along-column
 * (physical y) centroids / spacing for symptoms 2 & 3 — so the fix is validated in
 * pixels, not by eye. It renders through the REAL `drawVerticalRun` inside the REAL
 * page +90° transform (translate(W,0)·rotate(+90°), same as renderDocumentToCanvas).
 *
 * Font: Hiragino Mincho ProN registered under the serif JP fallback names the
 * renderer emits for ＭＳ 明朝 (the sample's body face). A substitute font shifts
 * absolute positions but the corner-vs-centre distinction the fix asserts is a
 * DIRECTION (comma high-right, ！ centred), which holds across fonts, so the guard
 * uses generous em-fraction tolerances rather than exact PDF pixel matches.
 *
 * CI-safe: gated on skia (a devDependency). Pure `drawVerticalRun` needs no WASM.
 */
import { describe, it, expect } from 'vitest';
import { existsSync } from 'node:fs';
import { loadSkiaForTests, importForTests } from './test-imports';

const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas, FontLibrary } = (skia ?? {}) as Skia;

// Hiragino Mincho ProN — present on macOS dev hosts; skip elsewhere.
const MINCHO = '/System/Library/Fonts/ヒラギノ明朝 ProN.ttc';
const haveFont = existsSync(MINCHO);

const vtMod = await importForTests(
  () => import('../../docx/src/vertical-text.ts'),
  'packages/docx/src/vertical-text.ts',
);

function useFont(): void {
  // Register under every serif-JP fallback name normalizeFontFamily emits.
  for (const fam of ['MS Mincho', 'Hiragino Mincho ProN', 'Noto Serif JP', 'Yu Mincho']) {
    FontLibrary.use(fam, [MINCHO]);
  }
}

/** Ink weight (0..255) of a pixel: 255 = full black ink on white. */
function ink(data: Uint8ClampedArray, i: number): number {
  return 255 - (0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2]);
}

describe.skipIf(!skia || !vtMod || !haveFont)('docx vertical mixed-metrics centering (§17.6.20)', () => {
  const fontPx = 48;

  /** Render a run list with the page +90° transform and return the pixel buffer.
   *  Each entry draws one `drawVerticalRun` at logical (x, baseline). */
  function renderRuns(
    runs: Array<{ text: string; x: number; letterSpacing: number }>,
    baseline: number,
    W: number,
    H: number,
  ): { data: Uint8ClampedArray; W: number; H: number } {
    const { drawVerticalRun } = vtMod as {
      drawVerticalRun: (
        ctx: unknown,
        text: string,
        x: number,
        baseline: number,
        fontPx: number,
        letterSpacingPx: number,
      ) => void;
    };
    useFont();
    const canvas = new Canvas(W, H);
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, W, H);
    ctx.fillStyle = '#000';
    ctx.font = `${fontPx}px "MS Mincho", serif`;
    ctx.save();
    ctx.translate(W, 0);
    ctx.rotate(Math.PI / 2);
    for (const r of runs) drawVerticalRun(ctx, r.text, r.x, baseline, fontPx, r.letterSpacing);
    ctx.restore();
    const img = ctx.getImageData(0, 0, W, H);
    return { data: img.data, W, H };
  }

  /** Ink-weighted mean physical X over a band of physical rows [py0,py1). The
   *  column centerline is at physical X = W − baseline. */
  function crossAxisCentroid(
    data: Uint8ClampedArray,
    W: number,
    py0: number,
    py1: number,
  ): number {
    let sx = 0;
    let sw = 0;
    for (let py = py0; py < py1; py++) {
      for (let px = 0; px < W; px++) {
        const w = ink(data, (py * W + px) * 4);
        if (w > 20) {
          sx += px * w;
          sw += w;
        }
      }
    }
    return sw > 0 ? sx / sw : NaN;
  }

  /** Project all ink onto the physical-y (column) axis → one weight per row. */
  function inkProfileY(data: Uint8ClampedArray, W: number, H: number): number[] {
    const prof = new Array<number>(H).fill(0);
    for (let py = 0; py < H; py++) {
      let s = 0;
      for (let px = 0; px < W; px++) {
        const w = ink(data, (py * W + px) * 4);
        if (w > 20) s += w;
      }
      prof[py] = s;
    }
    return prof;
  }

  /** Ink-weighted centroid of a physical-y profile within a window [c−r, c+r]. */
  function centroidInWindow(prof: number[], c: number, r: number): number {
    let sy = 0;
    let sw = 0;
    const y0 = Math.max(0, Math.round(c - r));
    const y1 = Math.min(prof.length, Math.round(c + r));
    for (let y = y0; y < y1; y++) {
      sy += y * prof[y];
      sw += prof[y];
    }
    return sw > 0 ? sy / sw : NaN;
  }

  it('SYMPTOM 1: upright CJK and sideways digits share the column centerline (≤1px)', () => {
    const W = 400;
    const H = 400;
    const baseline = 150;
    const centerline = W - baseline;
    const cellW = fontPx; // fullwidth cell
    const logX = 40;
    // "電話" upright (2 fullwidth cells), then "0" sideways in the 3rd cell.
    const { data } = renderRuns(
      [
        { text: '電話', x: logX, letterSpacing: 0 },
        { text: '0', x: logX + 2 * cellW, letterSpacing: 0 },
      ],
      baseline,
      W,
      H,
    );
    // Physical y band of glyph i = [logX + i*cellW, logX + (i+1)*cellW).
    const cCJK1 = crossAxisCentroid(data, W, logX, logX + cellW);
    const cCJK2 = crossAxisCentroid(data, W, logX + cellW, logX + 2 * cellW);
    const cDigit = crossAxisCentroid(data, W, logX + 2 * cellW, logX + 3 * cellW);
    // eslint-disable-next-line no-console
    console.log(
      `\n[SYMPTOM1] centerline X=${centerline}\n` +
        `  電 centroidX=${cCJK1.toFixed(1)} off=${(cCJK1 - centerline).toFixed(1)}\n` +
        `  話 centroidX=${cCJK2.toFixed(1)} off=${(cCJK2 - centerline).toFixed(1)}\n` +
        `  0  centroidX=${cDigit.toFixed(1)} off=${(cDigit - centerline).toFixed(1)}\n` +
        `  CJK↔digit gap=${Math.abs(cDigit - (cCJK1 + cCJK2) / 2).toFixed(1)}px`,
    );
    const cjkMean = (cCJK1 + cCJK2) / 2;
    // The digit's cross-axis centroid must sit on the same centerline the CJK
    // glyphs do (within 1px). The em-box center is the shared reference.
    expect(Math.abs(cDigit - cjkMean)).toBeLessThanOrEqual(1);
    // And the CJK glyphs themselves must sit on the centerline.
    expect(Math.abs(cjkMean - centerline)).toBeLessThanOrEqual(1.5);
  });

  it('SYMPTOM 2: 日（土）— the Tr parens sit a full cell from their neighbours', () => {
    const W = 500;
    const H = 500;
    const baseline = 150;
    const cellW = fontPx;
    const logX = 40;
    // 日 ( 土 ) — all fullwidth cells (parens are Tr rotate, 日/土 upright).
    const text = '日（土）';
    const { data } = renderRuns([{ text, x: logX, letterSpacing: 0 }], baseline, W, H);
    // Project ink to the column axis, then take each glyph's centroid within a
    // window centered on its expected cell centre (± half a cell) so a tall
    // narrow paren's ink is not miscounted into a neighbour's band.
    const prof = inkProfileY(data, W, H);
    const centers: number[] = [];
    for (let i = 0; i < 4; i++) {
      const expected = logX + (i + 0.5) * cellW;
      centers.push(centroidInWindow(prof, expected, cellW / 2));
    }
    const gaps = [centers[1] - centers[0], centers[2] - centers[1], centers[3] - centers[2]];
    // eslint-disable-next-line no-console
    console.log(
      `\n[SYMPTOM2] cellW=${cellW} centers=${centers.map((c) => c.toFixed(1)).join(', ')}\n` +
        `  gaps 日→( , (→土 , 土→) = ${gaps.map((g) => g.toFixed(1)).join(', ')}`,
    );
    // Each glyph's ink centroid sits ~one cell (fontPx) from its neighbour. The
    // tolerance is 6px (0.125em): we correct the bracket by its geometric ink-box
    // centre (actualBoundingBox*), which differs from the ink-WEIGHTED centroid
    // this probe measures by a few px — an inherent, typographically-correct
    // residual (Word centres by designed metrics too). The pre-fix state crowded
    // 土→) to 27px (0.44 cell); the fix removes that crowding. The KEY guard is
    // that no gap collapses below ~0.8 cell (the reported crowding symptom).
    for (const g of gaps) expect(Math.abs(g - cellW)).toBeLessThanOrEqual(6);
    expect(Math.min(...gaps)).toBeGreaterThan(cellW * 0.8);
  });

  it('SYMPTOM 3: 、。 hang in the cell UPPER part, not pulled low to the centre', () => {
    // Word / JIS X 4051 place the vertical comma/full stop in the cell's upper-right
    // corner. Along the column that means the ink centroid sits ABOVE the cell
    // centre (smaller physical y). PDF ground truth (sample-26): 、 ≈ −0.32em.
    // The #792 ink-centring dropped them to ≈ 0 (the "too low" defect); the fix
    // restores the font's corner design (Hiragino: ≈ −0.33em).
    const W = 400;
    const H = 400;
    const baseline = 150;
    const cellW = fontPx;
    const logX = 40;
    // 富 、 士 。 会 — CJK anchors around each punctuation mark.
    const { data } = renderRuns([{ text: '富、士。会', x: logX, letterSpacing: 0 }], baseline, W, H);
    const prof = inkProfileY(data, W, H);
    const offsets: number[] = [];
    const labels = ['富', '、', '士', '。', '会'];
    for (let i = 0; i < 5; i++) {
      const cellCentre = logX + (i + 0.5) * cellW;
      const c = centroidInWindow(prof, cellCentre, cellW / 2);
      offsets.push((c - cellCentre) / cellW); // em-fraction, negative = above centre
    }
    // eslint-disable-next-line no-console
    console.log(
      `\n[SYMPTOM3] along-column offsets (em, −=up): ` +
        labels.map((l, i) => `${l}=${offsets[i].toFixed(3)}`).join(', ') +
        `  (PDF: 、≈-0.32, 。 font-dependent)`,
    );
    // The comma (index 1) must sit clearly ABOVE the cell centre — the corner design.
    // Pre-fix it was ≈ −0.02em (at centre). Guard: at least 0.15em up (well past the
    // pre-fix state), confirming ink-centring no longer flattens the corner offset.
    expect(offsets[1]).toBeLessThan(-0.15);
    // CJK ideographs stay ≈ centred (their ink IS centred), unaffected by the fix.
    for (const i of [0, 2, 4]) expect(Math.abs(offsets[i])).toBeLessThanOrEqual(0.06);
  });

  it('SYMPTOM 4: ！ stays on the column centreline (cross-axis), not shifted right', () => {
    // ！ FF01 is drawn upright WITHOUT FE15 substitution, so its ink centres on the
    // column like the neighbouring CJK cells. PDF ground truth: ！ ≈ +0.03em cross-
    // axis. The old FE15 substitute pushed it to ≈ +0.30em right.
    const W = 400;
    const H = 400;
    const baseline = 150;
    const cellW = fontPx;
    const logX = 40;
    const centerline = W - baseline;
    // 話 ！ 話 — CJK anchors on both sides of the exclamation.
    const { data } = renderRuns([{ text: '話！話', x: logX, letterSpacing: 0 }], baseline, W, H);
    const cCJK1 = crossAxisCentroid(data, W, logX, logX + cellW);
    const cExcl = crossAxisCentroid(data, W, logX + cellW, logX + 2 * cellW);
    const cCJK2 = crossAxisCentroid(data, W, logX + 2 * cellW, logX + 3 * cellW);
    const cjkMean = (cCJK1 + cCJK2) / 2;
    const offEm = (cExcl - cjkMean) / cellW;
    // eslint-disable-next-line no-console
    console.log(
      `\n[SYMPTOM4] centerline X=${centerline}\n` +
        `  話 centroidX=${cCJK1.toFixed(1)}  ！ centroidX=${cExcl.toFixed(1)}  話 centroidX=${cCJK2.toFixed(1)}\n` +
        `  ！ cross-axis off from CJK centreline = ${(cExcl - cjkMean).toFixed(1)}px (${offEm.toFixed(3)}em)  (PDF: +0.03em)`,
    );
    // ！ ink must share the column centreline with the CJK cells (within 0.1em). The
    // pre-fix FE15 substitute sat ≈ +0.30em right — this guard would fail on it.
    expect(Math.abs(offEm)).toBeLessThanOrEqual(0.1);
  });
});
