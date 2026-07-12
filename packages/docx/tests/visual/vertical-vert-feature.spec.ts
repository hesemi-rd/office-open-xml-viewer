import { expect, test, type Page } from '@playwright/test';

interface InkSignature {
  minX: number;
  maxX: number;
  minY: number;
  maxY: number;
  centroidX: number;
  centroidY: number;
  alpha: number;
}

interface GlyphComparison {
  ch: string;
  changedByVert: boolean;
  subject: InkSignature;
  reference: InkSignature;
  pixelDifferenceRatio: number;
}

async function compareRendererWithFontVert(
  page: Page,
  family: string,
  repertoire: string,
): Promise<GlyphComparison[] | null> {
  await page.goto('/tests/visual/vertical-vert-feature.html');
  return page.evaluate(async ({ family, repertoire }) => {
    await document.fonts.ready;
    if (!document.fonts.check(`200px "${family}"`)) return null;

    const { drawVerticalRun } = await import('/src/vertical-text.ts');
    const subject = document.querySelector('#subject') as HTMLCanvasElement;
    const reference = document.querySelector('#reference') as HTMLCanvasElement;
    const size = 420;
    const fontPx = 200;
    const originX = 210;
    const startY = 40;

    const prepare = (canvas: HTMLCanvasElement): CanvasRenderingContext2D => {
      canvas.width = size;
      canvas.height = size;
      canvas.style.fontFeatureSettings = 'normal';
      const ctx = canvas.getContext('2d', { willReadFrequently: true });
      if (!ctx) throw new Error('2D context unavailable');
      ctx.clearRect(0, 0, size, size);
      ctx.fillStyle = '#000';
      ctx.font = `${fontPx}px "${family}"`;
      return ctx;
    };
    const withVert = <T>(ctx: CanvasRenderingContext2D, draw: () => T): T => {
      const canvas = ctx.canvas as HTMLCanvasElement;
      const previous = canvas.style.fontFeatureSettings;
      canvas.style.fontFeatureSettings = '"vert" 1';
      ctx.font = ctx.font;
      try {
        return draw();
      } finally {
        canvas.style.fontFeatureSettings = previous;
        ctx.font = ctx.font;
      }
    };
    const featuredCell = (ctx: CanvasRenderingContext2D, ch: string) =>
      withVert(ctx, () => {
        const previousAlign = ctx.textAlign;
        const previousBaseline = ctx.textBaseline;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        const m = ctx.measureText(ch);
        ctx.textAlign = previousAlign;
        ctx.textBaseline = previousBaseline;
        return { advance: m.width, origin: m.width / 2 };
      });
    const alpha = (ctx: CanvasRenderingContext2D): Uint8ClampedArray => {
      const rgba = ctx.getImageData(0, 0, size, size).data;
      const result = new Uint8ClampedArray(size * size);
      for (let i = 0; i < result.length; i += 1) result[i] = rgba[i * 4 + 3];
      return result;
    };
    const signature = (values: Uint8ClampedArray): InkSignature => {
      let minX = size;
      let maxX = -1;
      let minY = size;
      let maxY = -1;
      let alphaSum = 0;
      let weightedX = 0;
      let weightedY = 0;
      for (let y = 0; y < size; y += 1) {
        for (let x = 0; x < size; x += 1) {
          const value = values[y * size + x];
          if (value === 0) continue;
          minX = Math.min(minX, x);
          maxX = Math.max(maxX, x);
          minY = Math.min(minY, y);
          maxY = Math.max(maxY, y);
          alphaSum += value;
          weightedX += x * value;
          weightedY += y * value;
        }
      }
      if (alphaSum === 0) throw new Error('glyph produced no ink');
      return {
        minX,
        maxX,
        minY,
        maxY,
        centroidX: weightedX / alphaSum,
        centroidY: weightedY / alphaSum,
        alpha: alphaSum,
      };
    };

    const comparisons: GlyphComparison[] = [];
    for (const ch of repertoire) {
      const probe = prepare(subject);
      probe.textAlign = 'center';
      probe.textBaseline = 'middle';
      probe.fillText(ch, size / 2, size / 2);
      const plainAlpha = alpha(probe);
      probe.clearRect(0, 0, size, size);
      withVert(probe, () => probe.fillText(ch, size / 2, size / 2));
      const vertProbeAlpha = alpha(probe);
      let probeDifference = 0;
      for (let i = 0; i < plainAlpha.length; i += 1) {
        probeDifference += Math.abs(plainAlpha[i] - vertProbeAlpha[i]);
      }

      const subjectCtx = prepare(subject);
      const cell = featuredCell(subjectCtx, ch);
      subjectCtx.translate(originX, startY);
      subjectCtx.rotate(Math.PI / 2);
      drawVerticalRun(subjectCtx, ch, 0, 0, fontPx, 0);
      subjectCtx.setTransform(1, 0, 0, 1, 0, 0);
      const subjectAlpha = alpha(subjectCtx);

      const referenceCtx = prepare(reference);
      referenceCtx.textAlign = 'center';
      referenceCtx.textBaseline = 'middle';
      withVert(referenceCtx, () => referenceCtx.fillText(ch, originX, startY + cell.origin));
      const referenceAlpha = alpha(referenceCtx);
      let pixelDifference = 0;
      let referenceWeight = 0;
      for (let i = 0; i < subjectAlpha.length; i += 1) {
        pixelDifference += Math.abs(subjectAlpha[i] - referenceAlpha[i]);
        referenceWeight += referenceAlpha[i];
      }
      comparisons.push({
        ch,
        changedByVert: probeDifference > 0,
        subject: signature(subjectAlpha),
        reference: signature(referenceAlpha),
        pixelDifferenceRatio: pixelDifference / Math.max(1, referenceWeight),
      });
    }
    return comparisons;
  }, { family, repertoire });
}

test('layer A: the document font keeps Word-adjudicated vertical relationships', async ({ page }) => {
  await page.goto('/tests/visual/vertical-vert-feature.html');
  const result = await page.evaluate(async () => {
    const family = 'Yu Mincho';
    const fontPx = 200;
    const columnX = 450;
    const startY = 50;
    await document.fonts.ready;
    if (!document.fonts.check(`${fontPx}px "${family}"`)) return null;

    const { drawVerticalRun } = await import('/src/vertical-text.ts');
    const canvas = document.querySelector('#subject') as HTMLCanvasElement;
    const prepare = (): CanvasRenderingContext2D => {
      canvas.width = 900;
      canvas.height = 1100;
      canvas.style.fontFeatureSettings = 'normal';
      const ctx = canvas.getContext('2d', { willReadFrequently: true });
      if (!ctx) throw new Error('2D context unavailable');
      ctx.fillStyle = '#000';
      ctx.font = `${fontPx}px "${family}"`;
      return ctx;
    };
    const withVert = <T>(ctx: CanvasRenderingContext2D, fn: () => T): T => {
      const previous = canvas.style.fontFeatureSettings;
      canvas.style.fontFeatureSettings = '"vert" 1';
      ctx.font = ctx.font;
      try {
        return fn();
      } finally {
        canvas.style.fontFeatureSettings = previous;
        ctx.font = ctx.font;
      }
    };
    const cell = (ctx: CanvasRenderingContext2D, ch: string) =>
      withVert(ctx, () => {
        const previousAlign = ctx.textAlign;
        const previousBaseline = ctx.textBaseline;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        const m = ctx.measureText(ch);
        ctx.textAlign = previousAlign;
        ctx.textBaseline = previousBaseline;
        return {
          advance: m.width,
          origin: m.width / 2,
          asc: m.actualBoundingBoxAscent,
          desc: m.actualBoundingBoxDescent,
        };
      });
    const render = (ctx: CanvasRenderingContext2D, text: string) => {
      ctx.translate(columnX, startY);
      ctx.rotate(Math.PI / 2);
      drawVerticalRun(ctx, text, 0, 0, fontPx, 0);
      ctx.setTransform(1, 0, 0, 1, 0, 0);
    };
    const rangeInk = (ctx: CanvasRenderingContext2D, fromY: number, toY: number) => {
      const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
      let minX = canvas.width;
      let maxX = -1;
      let minY = canvas.height;
      let maxY = -1;
      let weight = 0;
      let weightedX = 0;
      let weightedY = 0;
      for (let y = Math.max(0, Math.floor(fromY)); y < Math.min(canvas.height, Math.ceil(toY)); y += 1) {
        for (let x = 0; x < canvas.width; x += 1) {
          const alpha = data[(y * canvas.width + x) * 4 + 3];
          if (alpha === 0) continue;
          minX = Math.min(minX, x);
          maxX = Math.max(maxX, x);
          minY = Math.min(minY, y);
          maxY = Math.max(maxY, y);
          weight += alpha;
          weightedX += x * alpha;
          weightedY += y * alpha;
        }
      }
      if (weight === 0) throw new Error(`no ink in y range ${fromY}..${toY}`);
      return {
        minX,
        maxX,
        minY,
        maxY,
        centroidX: weightedX / weight,
        centroidY: weightedY / weight,
      };
    };
    const connectedInk = (ctx: CanvasRenderingContext2D) => {
      const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
      const components: ReturnType<typeof rangeInk>[] = [];
      let componentStart: number | null = null;
      for (let y = 0; y < canvas.height; y += 1) {
        let occupied = false;
        for (let x = 0; x < canvas.width; x += 1) {
          if (data[(y * canvas.width + x) * 4 + 3] !== 0) {
            occupied = true;
            break;
          }
        }
        if (occupied && componentStart === null) componentStart = y;
        if (!occupied && componentStart !== null) {
          components.push(rangeInk(ctx, componentStart, y));
          componentStart = null;
        }
      }
      if (componentStart !== null) components.push(rangeInk(ctx, componentStart, canvas.height));
      return components;
    };
    const pairedComponents = (ctx: CanvasRenderingContext2D, label: string) => {
      const components = connectedInk(ctx);
      if (components.length !== 2) {
        throw new Error(`${label} produced ${components.length} connected y-components, expected 2`);
      }
      return components as [ReturnType<typeof rangeInk>, ReturnType<typeof rangeInk>];
    };

    const gapCtx = prepare();
    const prolongedCell = cell(gapCtx, 'ー');
    render(gapCtx, 'ーc');
    const boundary = startY + prolongedCell.advance;
    const prolonged = rangeInk(gapCtx, startY, boundary);
    const latin = rangeInk(gapCtx, boundary, canvas.height);

    const bracketCtx = prepare();
    const openCell = cell(bracketCtx, '「');
    const closeCell = cell(bracketCtx, '」');
    render(bracketCtx, '「」');
    // A designed closing-bracket poke may cross above its nominal cell start, so
    // a cell-boundary split would assign that ink to the opening bracket. The
    // font leaves empty scanlines between the two bands; use those components.
    const [open, close] = pairedComponents(bracketCtx, 'bracket pair');

    const punctuationCtx = prepare();
    const commaCell = cell(punctuationCtx, '、');
    const periodCell = cell(punctuationCtx, '。');
    render(punctuationCtx, '、。');
    const punctuationBoundary = startY + commaCell.advance;
    // Corner punctuation can likewise poke before its cell. Component splitting
    // preserves the full designed ink rather than clipping at either cell edge.
    const [comma, period] = pairedComponents(punctuationCtx, 'punctuation pair');

    return {
      fontPx,
      gapPx: latin.minY - prolonged.maxY - 1,
      prolongedAspect: (prolonged.maxY - prolonged.minY + 1) / (prolonged.maxX - prolonged.minX + 1),
      bracketBandSeparationPx:
        (close.minY + close.maxY) / 2 - (open.minY + open.maxY) / 2,
      bracketMetricSeparationPx:
        (openCell.advance + closeCell.origin + (closeCell.desc - closeCell.asc) / 2 -
          (openCell.origin + (openCell.desc - openCell.asc) / 2)),
      bracketBands: [open, close].map((band) => ({
        alongPx: band.maxY - band.minY + 1,
        acrossPx: band.maxX - band.minX + 1,
      })),
      comma: {
        rightOfColumn: comma.centroidX > columnX,
        positionInCell: (comma.centroidY - startY) / commaCell.advance,
      },
      period: {
        rightOfColumn: period.centroidX > columnX,
        positionInCell: (period.centroidY - punctuationBoundary) / periodCell.advance,
      },
    };
  });

  test.skip(result === null, 'Yu Mincho is required for the real-document-font acceptance layer');
  if (result === null) return;
  expect(result.gapPx, 'the prolonged mark and following Latin ink have a visible gap').toBeGreaterThan(0);
  expect(result.prolongedAspect, 'the feature-selected prolonged mark is a tall stroke').toBeGreaterThan(2);
  // Compare bbox-band centres to the SAME face's featured A/D prediction. The
  // two-pixel allowance is the 2px/pt AA noise observed by the wrapper; no
  // cross-font or sample-specific placement constant participates.
  expect(
    Math.abs(result.bracketBandSeparationPx - result.bracketMetricSeparationPx),
  ).toBeLessThanOrEqual(2);
  // Independent cell origins are 1em apart; a <0.75em band separation pins the
  // feature's optical pairing while leaving the exact placement to its A/D data.
  expect(result.bracketBandSeparationPx).toBeLessThan(result.fontPx * 0.75);
  for (const band of result.bracketBands) {
    // The reachable `vert` form flips a horizontal bracket from tall/narrow to
    // flat/wide. Exact aspect ratios vary by face; substitution failure would
    // invert this relationship and leave `acrossPx < alongPx`.
    expect(band.acrossPx, 'the bracket feature form is wider than tall').toBeGreaterThan(
      band.alongPx,
    );
  }
  for (const punctuation of [result.comma, result.period]) {
    expect(punctuation.rightOfColumn).toBe(true);
    expect(punctuation.positionInCell, 'corner punctuation stays in the leading cell third').toBeLessThan(1 / 3);
  }
});

test('layer A/B: reachable glyphs reproduce each installed font own vert design', async ({ page }) => {
  const repertoire = 'ー〜～、。「」（）：；！';
  for (const family of ['Yu Mincho', 'Hiragino Mincho ProN']) {
    const comparisons = await compareRendererWithFontVert(page, family, repertoire);
    test.skip(comparisons === null, `${family} is required for this host layer`);
    if (comparisons === null) continue;
    const changed = comparisons.filter((comparison) => comparison.changedByVert);
    expect(changed.length, `${family} exposes representative vert substitutions`).toBeGreaterThanOrEqual(9);
    expect(changed.map((comparison) => comparison.ch)).not.toContain('；');
    expect(changed.map((comparison) => comparison.ch)).not.toContain('！');
    for (const comparison of changed) {
      expect(comparison.subject.centroidX).toBeCloseTo(comparison.reference.centroidX, 0);
      expect(comparison.subject.centroidY).toBeCloseTo(comparison.reference.centroidY, 0);
      expect(comparison.subject.maxX - comparison.subject.minX).toBeCloseTo(
        comparison.reference.maxX - comparison.reference.minX,
        0,
      );
      expect(comparison.subject.maxY - comparison.subject.minY).toBeCloseTo(
        comparison.reference.maxY - comparison.reference.minY,
        0,
      );
      expect(comparison.pixelDifferenceRatio).toBeLessThan(0.08);
    }
  }
});

test('forced unreachable keeps FE/upright fallbacks and plain-rotates long marks', async ({ page }) => {
  await page.goto('/tests/visual/vertical-vert-feature.html');
  const result = await page.evaluate(async () => {
    const { drawVerticalRunWithCapability } = await import('/src/vertical-text.ts');
    const canvas = document.querySelector('#subject') as HTMLCanvasElement;
    const measure = canvas.getContext('2d');
    if (!measure) throw new Error('2D context unavailable');
    measure.font = '48px serif';
    const fills: string[] = [];
    const rotations: number[] = [];
    const scales: Array<[number, number]> = [];
    const transforms: number[][] = [];
    const ctx = {
      canvas,
      font: measure.font,
      fillStyle: '#000',
      textAlign: 'start' as CanvasTextAlign,
      textBaseline: 'alphabetic' as CanvasTextBaseline,
      save() {},
      restore() {},
      translate() {},
      rotate(angle: number) { rotations.push(angle); },
      scale(x: number, y: number) { scales.push([x, y]); },
      transform(...values: number[]) { transforms.push(values); },
      measureText(text: string) { return measure.measureText(text); },
      fillText(text: string) { fills.push(text); },
    } as unknown as CanvasRenderingContext2D;
    drawVerticalRunWithCapability(
      ctx,
      '、。！；「」：ー〜～“”',
      0,
      0,
      48,
      0,
      1,
      false,
      () => false,
    );
    return { fills, rotations, scales, transforms };
  });

  expect(result.fills).toEqual(['︑', '︒', '！', '；', '﹁', '﹂', '：', 'ー', '〜', '～', '“', '”']);
  expect(result.rotations).toHaveLength(6); // FE punctuation, !, ;, and two FE brackets
  expect(result.rotations.every((angle) => angle === -Math.PI / 2)).toBe(true);
  expect(result.scales.some(([, y]) => y === -1)).toBe(false);
  expect(result.transforms).toEqual([]);
});
