/**
 * Section-level `btLr` renders identically to `tbRl` — ECMA-376 §17.6.20,
 * issue #988 batch-3 adjudication ①.
 *
 * Word ground truth: a section whose `<w:textDirection w:val="btLr"/>` is applied
 * at SECTION level is laid out the same as `tbRl` (CJK upright stacked top→bottom,
 * Latin/digits sideways 90° CW, columns right→left) — Word does NOT honor btLr's
 * nominal bottom-to-top / left-to-right flow. This probe renders the SAME body
 * once as `btLr` and once as `tbRl` and asserts the two produce byte-identical
 * text-run layouts (position + rotation), pinning that `btLr` routes through the
 * vertical path.
 *
 * CI-safe: gated on docx WASM + skia-canvas; skips when absent, hard-fails under
 * OOXML_REQUIRE_SKIA=1.
 */
import { describe, it, expect } from 'vitest';
import { crc32 } from 'node:zlib';
import { installImageBitmapShim, installOffscreenCanvasShim } from './render.ts';
import type { NodeCanvasFactory } from './render.ts';
import { importForTests, loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas } = (skia ?? {}) as Skia;
const docxMod = await importForTests(() => import('./docx.ts'), './docx.ts (docx WASM)');
const rendererMod = await importForTests(
  () => import('./../../docx/src/renderer.ts'),
  'packages/docx/src/renderer.ts',
);

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

const factory: NodeCanvasFactory = {
  createCanvas: (w, h) =>
    new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (() => {
    throw new Error('loadImage not needed');
  }) as unknown as NodeCanvasFactory['loadImage'],
};

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function storedZip(files: Record<string, string>): Uint8Array {
  const enc = new TextEncoder();
  const chunks: number[] = [];
  const central: number[] = [];
  const u16 = (n: number) => [n & 0xff, (n >> 8) & 0xff];
  const u32 = (n: number) => [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >> 24) & 0xff];
  let offset = 0;
  for (const [name, content] of Object.entries(files)) {
    const nameBytes = [...enc.encode(name)];
    const data = [...enc.encode(content)];
    const crc = crc32(Uint8Array.from(data)) >>> 0;
    const local = [
      ...u32(0x04034b50), ...u16(20), ...u16(0), ...u16(0), ...u16(0), ...u16(0),
      ...u32(crc), ...u32(data.length), ...u32(data.length),
      ...u16(nameBytes.length), ...u16(0), ...nameBytes, ...data,
    ];
    central.push(
      ...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0), ...u16(0), ...u16(0), ...u16(0),
      ...u32(crc), ...u32(data.length), ...u32(data.length),
      ...u16(nameBytes.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(offset),
      ...nameBytes,
    );
    chunks.push(...local);
    offset += local.length;
  }
  const centralOffset = offset;
  const end = [
    ...u32(0x06054b50), ...u16(0), ...u16(0),
    ...u16(Object.keys(files).length), ...u16(Object.keys(files).length),
    ...u32(central.length), ...u32(centralOffset), ...u16(0),
  ];
  return Uint8Array.from([...chunks, ...central, ...end]);
}

function verticalDocx(dir: 'btLr' | 'tbRl'): Uint8Array {
  const contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '</Types>';
  const rootRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';
  const para = (t: string) =>
    `<w:p><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>${t}</w:t></w:r></w:p>`;
  const document =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>` +
    para('縦横ABC混在123テスト') +
    para('日本語Wordと英語Englishの並び') +
    '<w:sectPr>' +
    '<w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
    `<w:textDirection w:val="${dir}"/>` +
    '</w:sectPr></w:body></w:document>';
  return storedZip({
    '[Content_Types].xml': contentTypes,
    '_rels/.rels': rootRels,
    'word/document.xml': document,
  });
}

interface Run { t: string; x: number; y: number; w: number; h: number; tr: string }

async function renderRuns(dir: 'btLr' | 'tbRl'): Promise<Run[]> {
  const { parseDocx } = docxMod as { parseDocx: (b: Uint8Array) => Any };
  const { renderDocumentToCanvas } = rendererMod as Any;
  const doc = parseDocx(verticalDocx(dir));
  const canvas = new Canvas(Math.round(doc.section.pageWidth), Math.round(doc.section.pageHeight));
  const runs: Run[] = [];
  const rImg = installImageBitmapShim(factory);
  const rOff = installOffscreenCanvasShim(factory);
  try {
    await renderDocumentToCanvas(doc, canvas, 0, {
      dpr: 1,
      width: doc.section.pageWidth,
      onTextRun: (r: Any) =>
        runs.push({
          t: r.text,
          x: Math.round(r.x),
          y: Math.round(r.y),
          w: Math.round(r.w),
          h: Math.round(r.h),
          tr: r.transform ?? '',
        }),
    });
  } finally {
    rOff();
    rImg();
  }
  return runs;
}

describe.skipIf(!skia || !docxMod || !rendererMod)('docx section btLr ≡ tbRl (§17.6.20)', () => {
  it('renders a btLr section identically to tbRl (vertical, same run layout)', async () => {
    // Sequential — each render installs/restores global OffscreenCanvas +
    // createImageBitmap shims; running them concurrently could restore out of order.
    const btlr = await renderRuns('btLr');
    const tbrl = await renderRuns('tbRl');
    // Both must actually be vertical (rotate transform present).
    expect(btlr.length, 'btLr produced runs').toBeGreaterThan(0);
    expect(btlr.every((r) => /rotate/.test(r.tr)), 'btLr runs are vertical').toBe(true);
    expect(tbrl.every((r) => /rotate/.test(r.tr)), 'tbRl runs are vertical').toBe(true);
    // And identical layout: same texts, positions, and rotation.
    expect(btlr).toEqual(tbrl);
  });
});
