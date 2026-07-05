import { test, expect } from '@playwright/test';
import { existsSync } from 'node:fs';

// ECMA-376 §17.6.12 non-regression on REAL private samples. sample-12/13/14 all
// carry `<w:pgNumType>` (sample-12: continuous start=1; sample-13: nextPage start=1
// + a continuous start=2; sample-14: nextPage start=1) plus a footer PAGE field.
// Word's PDF ground truth (measured with pdftotext) shows SEQUENTIAL footer numbers
// — none of these restarts is VISIBLE because every start value coincides with the
// natural continuation at the page where it fires: start=1 on the first section is
// the identity, and sample-13's continuous start=2 section spills onto physical
// page 2, where its restart fires but 2 equals the natural continuation (1+1)
// anyway. (A continuous restart is thus observed at a spillover page top, not
// suppressed — see page-numbering.ts; whether Word matches at a NON-coinciding
// spillover start is unverified, tracked as a follow-up issue.) So the DISPLAYED
// footer number must equal the physical page number, unchanged from before this
// feature. Skips gracefully when the (gitignored) sample is absent.
// sample-12 (continuous start=1, footer PAGE) and sample-13 (nextPage start=1 +
// continuous start=2, footer PAGE) both isolate the footer number cleanly at the
// page bottom, so their sequential numbering is a direct non-regression check.
// (sample-14 also carries pgNumType but its footer shares the bottom band with
// other numeric content, so the position heuristic can't isolate it; its structure
// — nextPage start=1 on the first section — is identical to sample-13's break[0]
// and is covered deterministically by page-numbering.test.ts.)
const CASES: { file: string; pageCount: number; width: number; expected: string[] }[] = [
  { file: 'private/sample-12', pageCount: 4, width: 595, expected: ['1', '2', '3', '4'] },
  { file: 'private/sample-13', pageCount: 6, width: 595, expected: ['1', '2', '3', '4', '5'] },
];

test.describe('page-number restart non-regression (§17.6.12)', () => {
  for (const { file, pageCount, width, expected } of CASES) {
    test(`${file}: footer PAGE numbers stay sequential`, async ({ page }) => {
      if (!existsSync(`${process.cwd()}/public/${file}.docx`)) {
        test.skip(true, `gitignored sample ${file}.docx not present`);
      }
      // Load a page in the Vite module graph first so a dynamic `import('/src/...')`
      // inside page.evaluate resolves against the dev server (the fixture imports it).
      await page.goto(`/tests/visual/fixture.html?file=demo/sample-1.docx&page=0&width=595`);
      await page.waitForFunction(
        () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
        { timeout: 30_000 },
      );
      // Render each page and collect the bottom-most SHORT text run (the footer
      // page number). A footer number is a 1–3 char numeric/roman/letter token near
      // the page bottom, so we take the lowest run whose text is <= 4 chars.
      const numbers = await page.evaluate(
        async ({ file, pageCount, width }) => {
          const { DocxDocument } = await import('/src/index.ts');
          const doc = await DocxDocument.load('/' + file + '.docx');
          const out: (string | null)[] = [];
          for (let i = 0; i < pageCount; i++) {
            const canvas = document.createElement('canvas');
            const runs: { text: string; y: number }[] = [];
            await doc.renderPage(canvas, i, {
              width, dpr: 1,
              onTextRun: (r: { text: string; y: number }) => runs.push({ text: r.text, y: r.y }),
            });
            // Footer page number = the lowest-on-page run whose trimmed text is a
            // pure page-number token (Arabic digits, or roman/letter glyphs). This
            // excludes footnote/dagger marks (†) and body text. Prefer the bottom-most.
            const isPageToken = (t: string) => /^[0-9]{1,4}$/.test(t) || /^[ivxlcdmIVXLCDM]{1,4}$/.test(t) || /^[a-zA-Z]{1,3}$/.test(t);
            const tokens = runs
              .map((r) => ({ text: r.text.trim(), y: r.y }))
              .filter((r) => isPageToken(r.text))
              .sort((a, b) => b.y - a.y);
            out.push(tokens.length ? tokens[0].text : null);
          }
          return out;
        },
        { file, pageCount, width },
      );
      // The first `expected.length` pages carry a footer; assert those.
      for (let i = 0; i < expected.length; i++) {
        expect(numbers[i], `${file} page ${i + 1} footer number`).toBe(expected[i]);
      }
    });
  }
});
