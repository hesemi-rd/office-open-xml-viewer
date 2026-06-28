import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  CellElement,
  DocParagraph,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  DocxTextRun,
  SectionProps,
} from './types';

// ECMA-376 §17.3.1.33 + §17.4.84 (vAlign): when a cell's vAlign is `center` or
// `bottom`, Word vertically aligns the cell's INKED block (the line boxes), NOT
// the inked block + leading paragraph spaceBefore + trailing paragraph
// spaceAfter. Both of those produce no ink (nothing surrounds them inside the
// cell), so including them in the vAlign block height pushes the visible block
// off centre/bottom.
//
// Before the fix, the renderer trimmed only the LAST paragraph's spaceAfter
// from `contentH` (asymmetric) and then `renderParagraph` re-consumed the FIRST
// paragraph's spaceBefore on top of `cellState.y`. With p1.spaceBefore=6 pt and
// p2.spaceAfter=6 pt the inked block landed exactly +3 pt (= spaceBefore/2)
// below the true cell centre (and symmetric for `bottom`).
//
// This file pins both directions of the symmetry: in vAlign=center / bottom,
// adding spaceBefore to the FIRST paragraph (or spaceAfter to the LAST) must
// NOT shift the painted baselines.

// ---------- recording 2D context ----------
// Records each fillText call so we can read the painted baseline y per
// paragraph. Glyph advance = charCount × fontPx; ascent/descent = 0.8/0.2 em
// (same shape as the other recording tests in this package).
interface FillTextCall { text: string; x: number; y: number; font: string; }

function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  fillTextCalls: FillTextCall[];
} {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const fillTextCalls: FillTextCall[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText(text: string, x: number, y: number) {
      fillTextCalls.push({ text, x, y, font });
    },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0, height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, fillTextCalls };
}

// ---------- model builders ----------

// A deliberately SYNTHETIC, untabled font: the mock canvas reports a clean
// 1.0 em box (0.8/0.2) for it, so these vAlign tests isolate the §17.4.84 leading
// trim without the font-metrics single-line FLOOR. (Real Latin faces like Times
// New Roman / Arial are tabled with their hhea design height — see font-metrics.ts
// — which raises the line box and is covered by font-metrics.test.ts instead.)
const TEST_FONT = 'Synthetic Untabled Serif';

function textRun(text: string): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: TEST_FONT, fontFamilyEastAsia: '',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  };
}

/** Build a paragraph wrapped as a CellElement (carries the `type: 'paragraph'`
 *  discriminator that {@link measureCellElementHeight} reads). */
function paraOf(text: string, opts: Partial<DocParagraph> = {}): CellElement {
  return {
    type: 'paragraph',
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{ type: 'text', ...textRun(text) } as DocParagraph['runs'][number]],
    defaultFontSize: 10, defaultFontFamily: TEST_FONT,
    widowControl: false,
    ...opts,
  } as unknown as CellElement;
}

function cell(
  content: CellElement[],
  vAlign: 'top' | 'center' | 'bottom',
): DocTableCell {
  return {
    content,
    colSpan: 1,
    vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background: null,
    vAlign,
    widthPt: 400,
  } as DocTableCell;
}

function row(c: DocTableCell, rowHeightPt: number, rule: 'auto' | 'atLeast' | 'exact'): DocTableRow {
  return {
    cells: [c],
    rowHeight: rowHeightPt,
    rowHeightRule: rule,
    isHeader: false,
  } as DocTableRow;
}

function tableOf(r: DocTableRow): DocTable {
  return {
    colWidths: [400],
    rows: [r],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
  } as DocTable;
}

function docWithTable(t: DocTable): DocxDocumentModel {
  return {
    section: {
      pageWidth: 400, pageHeight: 400,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [{ type: 'table', ...t } as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { [TEST_FONT]: 'roman' },
  } as unknown as DocxDocumentModel;
}

async function renderAndRead(t: DocTable) {
  const { canvas, fillTextCalls } = makeRecordingCanvas();
  await renderDocumentToCanvas(docWithTable(t), canvas, 0, {
    dpr: 1,
    width: 400, // scale = 1 (px per pt) ⇒ asserts are in pt-equivalent units
  });
  return fillTextCalls;
}

const TOL = 0.01; // px — the deliberate symmetric trim is exact arithmetic

// ============================================================================
// vAlign = center
// ============================================================================

describe('cell vAlign=center — leading spaceBefore must NOT shift the inked block (§17.4.84 / §17.3.1.33)', () => {
  it('p1.spaceBefore=6 pt does NOT shift the first paragraph baseline (vs. spaceBefore=0)', async () => {
    // Reference: both paragraphs have ZERO spacing → baselines define the
    // "true centred" positions.
    const baseRow = row(
      cell([paraOf('p1') as unknown as CellElement, paraOf('p2') as unknown as CellElement], 'center'),
      100,
      'atLeast',
    );
    const baseCalls = await renderAndRead(tableOf(baseRow));
    const baseP1 = baseCalls.find((c) => c.text === 'p1');
    const baseP2 = baseCalls.find((c) => c.text === 'p2');
    expect(baseP1).toBeDefined();
    expect(baseP2).toBeDefined();

    // Subject: p1 carries 6 pt spaceBefore, p2 carries 6 pt spaceAfter — both
    // are edge-collapsed in Word, so the visible baselines must match the
    // zero-spacing reference EXACTLY (no asymmetric shift).
    const subjRow = row(
      cell(
        [
          paraOf('p1', { spaceBefore: 6 }) as unknown as CellElement,
          paraOf('p2', { spaceAfter: 6 }) as unknown as CellElement,
        ],
        'center',
      ),
      100,
      'atLeast',
    );
    const subjCalls = await renderAndRead(tableOf(subjRow));
    const subjP1 = subjCalls.find((c) => c.text === 'p1');
    const subjP2 = subjCalls.find((c) => c.text === 'p2');
    expect(subjP1).toBeDefined();
    expect(subjP2).toBeDefined();

    // Before the fix, subjP1.y was baseP1.y + 3 (= 6 / 2 ≈ spaceBefore/2). After
    // the fix the symmetric trim + cellState.y pull-up cancel out and the
    // baselines coincide.
    expect(subjP1!.y).toBeCloseTo(baseP1!.y, 2);
    expect(subjP2!.y).toBeCloseTo(baseP2!.y, 2);
  });

  it('inked block is vertically centred in the cell (midpoint = cell midpoint)', async () => {
    // Single paragraph with 6 pt leading spaceBefore + 6 pt trailing spaceAfter.
    // With the fix, the line box is centred in the cell; without the fix the
    // line box sits +3 pt past centre.
    const t = tableOf(row(
      cell([paraOf('x', { spaceBefore: 6, spaceAfter: 6 }) as unknown as CellElement], 'center'),
      100,
      'atLeast',
    ));
    const calls = await renderAndRead(t);
    const x = calls.find((c) => c.text === 'x');
    expect(x).toBeDefined();
    // Glyph metrics (this recording canvas): 10 pt font ⇒ ascent 8, descent 2.
    // Line box height = ascent + descent = 10. Inked midpoint = baseline - asc + 5.
    const inkedTop = x!.y - 10 * 0.8;
    const inkedBottom = x!.y + 10 * 0.2;
    const inkedMid = (inkedTop + inkedBottom) / 2;
    const cellMid = 50; // row at y=0, height=100 ⇒ centre at 50.
    expect(inkedMid).toBeCloseTo(cellMid, 2);
  });

  it('first element is a nested table — leading spaceBefore is treated as 0 (no firstSpaceBefore pull-up)', async () => {
    // Nested table first + trailing structural empty paragraph (§17.4.7).
    // trimTrailingStructuralMarker drops the trailing empty paragraph, so the
    // visibleContent is [nestedTable]. No paragraph anywhere ⇒ firstSpaceBefore
    // = lastSpaceAfter = 0. The nested table should sit centred in the outer
    // cell without any spurious vertical offset.
    const nestedRow = row(
      cell([paraOf('nested') as unknown as CellElement], 'top'),
      20,
      'atLeast',
    );
    const nested = tableOf(nestedRow);
    const outerRow = row(
      cell(
        [
          { type: 'table', ...nested } as unknown as CellElement,
          paraOf('') as unknown as CellElement, // structural trailing empty paragraph
        ],
        'center',
      ),
      100,
      'atLeast',
    );
    const calls = await renderAndRead(tableOf(outerRow));
    const nestedText = calls.find((c) => c.text === 'nested');
    expect(nestedText).toBeDefined();
    // The 20 pt nested row centred in a 100 pt outer cell ⇒ nested row top at 40.
    // Inside the nested row, the single line baseline is at top + ascent (8 px @ scale 1)
    // plus any line-box leading (≈0 for an empty-spacing 10 pt paragraph).
    // We assert the baseline sits inside [40, 60] — i.e. the nested table did
    // NOT get pushed off-centre by an unrelated firstSpaceBefore.
    expect(nestedText!.y).toBeGreaterThanOrEqual(40);
    expect(nestedText!.y).toBeLessThanOrEqual(60);
  });
});

// ============================================================================
// vAlign = bottom
// ============================================================================

describe('cell vAlign=bottom — symmetric trim of leading spaceBefore (§17.4.84 / §17.3.1.33)', () => {
  // For vAlign=bottom the asymmetric-trim bug structurally cancels (renderParagraph
  // re-consumes p1.spaceBefore in the +y direction, while the buggy contentH
  // inflated by the same amount pulled cellState.y in the −y direction). The
  // visible inked bottom therefore stays at y+h-mb in both buggy and fixed code
  // for typical paragraph-first / paragraph-last cells. The symmetric trim is
  // kept for spec consistency (§17.4.84) and as defence against future changes
  // to renderParagraph's spaceBefore handling — these tests pin the absolute
  // inked-bottom invariant rather than a buggy-vs-fixed delta.
  it('inked block bottom hugs the cell bottom regardless of leading spaceBefore', async () => {
    const subjRow = row(
      cell(
        [
          paraOf('p1', { spaceBefore: 6 }) as unknown as CellElement,
          paraOf('p2', { spaceAfter: 0 }) as unknown as CellElement,
        ],
        'bottom',
      ),
      100,
      'atLeast',
    );
    const calls = await renderAndRead(tableOf(subjRow));
    const subjP2 = calls.find((c) => c.text === 'p2');
    expect(subjP2).toBeDefined();
    // Last line's descent (≈ 10 × 0.2 = 2). Inked bottom = baseline + descent.
    const inkedBottom = subjP2!.y + 10 * 0.2;
    expect(inkedBottom).toBeCloseTo(100, 2); // mb=0, cell bottom = 100.
  });

  it('p1.spaceBefore=6 pt does NOT shift the first paragraph baseline (vs. spaceBefore=0)', async () => {
    const baseRow = row(
      cell([paraOf('p1') as unknown as CellElement, paraOf('p2') as unknown as CellElement], 'bottom'),
      100,
      'atLeast',
    );
    const baseCalls = await renderAndRead(tableOf(baseRow));
    const baseP1 = baseCalls.find((c) => c.text === 'p1');
    const baseP2 = baseCalls.find((c) => c.text === 'p2');
    expect(baseP1).toBeDefined();
    expect(baseP2).toBeDefined();

    const subjRow = row(
      cell(
        [
          paraOf('p1', { spaceBefore: 6 }) as unknown as CellElement,
          paraOf('p2', { spaceAfter: 6 }) as unknown as CellElement,
        ],
        'bottom',
      ),
      100,
      'atLeast',
    );
    const subjCalls = await renderAndRead(tableOf(subjRow));
    const subjP1 = subjCalls.find((c) => c.text === 'p1');
    const subjP2 = subjCalls.find((c) => c.text === 'p2');
    expect(subjP1).toBeDefined();
    expect(subjP2).toBeDefined();

    expect(subjP1!.y).toBeCloseTo(baseP1!.y, 2);
    expect(subjP2!.y).toBeCloseTo(baseP2!.y, 2);
  });
});

// ============================================================================
// vAlign = top (regression guard)
// ============================================================================

describe('cell vAlign=top is unchanged by the symmetric trim (regression guard)', () => {
  it('top-aligned p1.spaceBefore=6 pt still shifts the first baseline down by 6 pt', async () => {
    // vAlign=top does NOT route through the centering branch — the test asserts
    // the un-touched ordinary path still consumes spaceBefore normally.
    const baseRow = row(
      cell([paraOf('p1') as unknown as CellElement], 'top'),
      100,
      'atLeast',
    );
    const baseCalls = await renderAndRead(tableOf(baseRow));
    const baseP1 = baseCalls.find((c) => c.text === 'p1');
    expect(baseP1).toBeDefined();

    const subjRow = row(
      cell([paraOf('p1', { spaceBefore: 6 }) as unknown as CellElement], 'top'),
      100,
      'atLeast',
    );
    const subjCalls = await renderAndRead(tableOf(subjRow));
    const subjP1 = subjCalls.find((c) => c.text === 'p1');
    expect(subjP1).toBeDefined();

    // Top-aligned baseline shifts down by exactly spaceBefore (6 pt @ scale 1).
    expect(subjP1!.y - baseP1!.y).toBeCloseTo(6, TOL);
  });
});
