import { expect, test } from '@playwright/test';

interface StrokeMeasurement {
  angleDeg: number;
  headMinusTailPx: number;
  topQuarterInkWeight: number;
  bottomQuarterInkWeight: number;
  inkWidth: number;
  inkHeight: number;
  centroidRangePx: number;
}

test('tbRl uses the font vert glyph for prolonged and wave marks', async ({ page }) => {
  await page.goto('/tests/visual/vertical-vert-feature.html');

  const result = await page.evaluate(async () => {
    const family = 'Hiragino Mincho ProN';
    await document.fonts.ready;
    if (!document.fonts.check(`200px "${family}"`)) return null;

    const { drawVerticalRun } = await import('/src/vertical-text.ts');
    const canvas = document.querySelector('canvas') as HTMLCanvasElement;
    canvas.width = 900;
    canvas.height = 1100;
    canvas.style.width = '900px';
    canvas.style.height = '1100px';
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    if (!ctx) throw new Error('2D context unavailable');
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = '#000';
    ctx.font = `200px "${family}"`;
    ctx.translate(500, 50);
    ctx.rotate(Math.PI / 2);
    drawVerticalRun(ctx, '話ーー話〜～', 0, 0, 200, 0);
    ctx.setTransform(1, 0, 0, 1, 0, 0);

    const measureCell = (index: number): StrokeMeasurement => {
      const x0 = 350;
      const y0 = 50 + index * 200;
      const width = 300;
      const height = 200;
      const data = ctx.getImageData(x0, y0, width, height).data;
      const rows: Array<{ y: number; x: number; weight: number }> = [];
      let minX = width;
      let maxX = -1;
      let minY = height;
      let maxY = -1;
      for (let y = 0; y < height; y += 1) {
        let weightedX = 0;
        let weight = 0;
        for (let x = 0; x < width; x += 1) {
          const alpha = data[(y * width + x) * 4 + 3];
          if (alpha === 0) continue;
          weightedX += x * alpha;
          weight += alpha;
          minX = Math.min(minX, x);
          maxX = Math.max(maxX, x);
          minY = Math.min(minY, y);
          maxY = Math.max(maxY, y);
        }
        if (weight > 0) rows.push({ y, x: weightedX / weight, weight });
      }
      if (rows.length < 2) throw new Error(`no measurable ink in cell ${index}`);

      // Theil-Sen over every occupied scanline: endpoint bulges remain evidence
      // instead of being removed by a fixed end trim.
      const slopes: number[] = [];
      for (let i = 0; i < rows.length; i += 1) {
        for (let j = i + 1; j < rows.length; j += 1) {
          slopes.push((rows[j].x - rows[i].x) / (rows[j].y - rows[i].y));
        }
      }
      slopes.sort((a, b) => a - b);
      const slope = slopes[Math.floor(slopes.length / 2)];
      const quarter = Math.max(1, Math.floor(rows.length / 4));
      const average = (values: Array<{ x: number }>) =>
        values.reduce((sum, value) => sum + value.x, 0) / values.length;
      const totalWeight = (values: Array<{ weight: number }>) =>
        values.reduce((sum, value) => sum + value.weight, 0);
      const topQuarter = rows.slice(0, quarter);
      const bottomQuarter = rows.slice(-quarter);
      return {
        angleDeg: (Math.atan(slope) * 180) / Math.PI,
        headMinusTailPx: average(topQuarter) - average(bottomQuarter),
        topQuarterInkWeight: totalWeight(topQuarter),
        bottomQuarterInkWeight: totalWeight(bottomQuarter),
        inkWidth: maxX - minX + 1,
        inkHeight: maxY - minY + 1,
        centroidRangePx:
          Math.max(...rows.map((row) => row.x)) - Math.min(...rows.map((row) => row.x)),
      };
    };

    return {
      prolonged: [measureCell(1), measureCell(2)],
      waveDash: measureCell(4),
      fullwidthTilde: measureCell(5),
    };
  });

  test.skip(result === null, 'Hiragino Mincho ProN is required for Word-PDF adjudication');
  if (result === null) return;

  for (const stroke of result.prolonged) {
    console.log(
      `vert prolonged angle=${stroke.angleDeg.toFixed(3)}deg head-tail=${stroke.headMinusTailPx.toFixed(2)}px (${((stroke.headMinusTailPx / stroke.inkHeight) * 100).toFixed(2)}%) top-bottom-weight=${(stroke.topQuarterInkWeight / stroke.bottomQuarterInkWeight).toFixed(2)}x`,
    );
    expect(Math.abs(stroke.angleDeg)).toBeLessThanOrEqual(0.8);
    // Word ground truth (word47-chouonpu-zoom-only.png) is -1.02% over the
    // measured ink height, so the signed quarter-centroid offset is a taper
    // check rather than evidence that the head must lie to the right.
    expect(Math.abs(stroke.headMinusTailPx)).toBeLessThanOrEqual(stroke.inkHeight * 0.025);
    // The same Word crop has 1.30x as much top-quarter ink weight. A 180-degree
    // inversion reverses this relationship even though its stroke stays vertical.
    expect(stroke.topQuarterInkWeight).toBeGreaterThan(stroke.bottomQuarterInkWeight);
    expect(stroke.inkHeight).toBeGreaterThan(stroke.inkWidth * 2);
  }
  for (const [name, wave] of [
    ['wave dash', result.waveDash],
    ['fullwidth tilde', result.fullwidthTilde],
  ] as const) {
    expect(wave.inkHeight, `${name} is the real tall vert glyph`).toBeGreaterThan(wave.inkWidth);
    expect(wave.centroidRangePx, `${name} preserves its designed waveform`).toBeGreaterThan(8);
  }
});
