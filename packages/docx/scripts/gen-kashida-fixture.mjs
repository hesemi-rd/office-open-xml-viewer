// Generates a synthetic .docx fixture exercising WordprocessingML kashida
// justification (ECMA-376 §17.18.44 ST_Jc: lowKashida / mediumKashida /
// highKashida) plus a `both` control, for WORD ADJUDICATION of the renderer.
//
//   node packages/docx/scripts/gen-kashida-fixture.mjs
//   -> packages/docx/tests/fixtures/kashida-justification.docx
//
// Open the .docx in Microsoft Word and export a PDF next to it
// (kashida-justification.pdf) to serve as the pixel ground truth: Word fills the
// slack of each RTL Arabic paragraph by ELONGATING words with U+0640 tatweel at
// valid joining points (more aggressively low -> medium -> high), while the
// `both` control widens inter-word spaces only.
//
// The document is authored with strict CT_PPr child ordering (bidi before jc,
// §17.3.1) so the file round-trips through a schema-validating consumer. It is a
// self-contained ZIP writer (DEFLATE via node:zlib) with no third-party deps.

import { writeFile, mkdir } from 'node:fs/promises';
import { deflateRawSync } from 'node:zlib';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const HERE = dirname(fileURLToPath(import.meta.url));
const OUT_DIR = join(HERE, '..', 'tests', 'fixtures');
const OUT = join(OUT_DIR, 'kashida-justification.docx');

// ── minimal ZIP writer ──────────────────────────────────────────────────────
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}
function zip(entries) {
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const { name, data } of entries) {
    const nameBuf = Buffer.from(name, 'utf8');
    const raw = Buffer.from(data, 'utf8');
    const comp = deflateRawSync(raw);
    const crc = crc32(raw);
    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0);
    lh.writeUInt16LE(20, 4); // version needed
    lh.writeUInt16LE(0, 6); // flags
    lh.writeUInt16LE(8, 8); // method: deflate
    lh.writeUInt16LE(0, 10); // time
    lh.writeUInt16LE(0x21, 12); // date (arbitrary, 1980-01-01ish)
    lh.writeUInt32LE(crc, 14);
    lh.writeUInt32LE(comp.length, 18);
    lh.writeUInt32LE(raw.length, 22);
    lh.writeUInt16LE(nameBuf.length, 26);
    lh.writeUInt16LE(0, 28); // extra len
    locals.push(lh, nameBuf, comp);

    const ch = Buffer.alloc(46);
    ch.writeUInt32LE(0x02014b50, 0);
    ch.writeUInt16LE(20, 4); // version made by
    ch.writeUInt16LE(20, 6); // version needed
    ch.writeUInt16LE(0, 8);
    ch.writeUInt16LE(8, 10);
    ch.writeUInt16LE(0, 12);
    ch.writeUInt16LE(0x21, 14);
    ch.writeUInt32LE(crc, 16);
    ch.writeUInt32LE(comp.length, 20);
    ch.writeUInt32LE(raw.length, 24);
    ch.writeUInt16LE(nameBuf.length, 28);
    ch.writeUInt16LE(0, 30);
    ch.writeUInt16LE(0, 32);
    ch.writeUInt16LE(0, 34);
    ch.writeUInt16LE(0, 36);
    ch.writeUInt32LE(0, 38);
    ch.writeUInt32LE(offset, 42);
    centrals.push(ch, nameBuf);

    offset += lh.length + nameBuf.length + comp.length;
  }
  const centralBuf = Buffer.concat(centrals);
  const localBuf = Buffer.concat(locals);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(centralBuf.length, 12);
  eocd.writeUInt32LE(localBuf.length, 16);
  return Buffer.concat([localBuf, centralBuf, eocd]);
}

// ── document content ────────────────────────────────────────────────────────
// A neutral Arabic sentence with many dual-joining letters (so there are ample
// tatweel insertion points). Repeated to force multi-line wrapping, so every
// non-last line is a justify candidate.
const ARABIC =
  'اللغة العربية تكتب من اليمين إلى اليسار وتتصل حروفها ببعضها البعض في الكلمة الواحدة';
const ARABIC_PARA = new Array(4).fill(ARABIC).join(' ');

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function heading(label) {
  // LTR Latin heading so the adjudicator can tell the paragraphs apart.
  return `<w:p><w:pPr><w:spacing w:before="240" w:after="60"/><w:jc w:val="left"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${label}</w:t></w:r></w:p>`;
}
function arabicPara(jc) {
  // CT_PPr order: spacing, ind, ... bidi, jc (bidi precedes jc in the schema).
  return (
    `<w:p><w:pPr><w:bidi/><w:jc w:val="${jc}"/>` +
    `<w:rPr><w:rtl/><w:rFonts w:cs="Arial"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:pPr>` +
    `<w:r><w:rPr><w:rtl/><w:rFonts w:cs="Arial"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>` +
    `<w:t xml:space="preserve">${ARABIC_PARA}</w:t></w:r></w:p>`
  );
}

const body =
  heading('1) jc = both  (control: inter-word spaces widen, no kashida)') +
  arabicPara('both') +
  heading('2) jc = lowKashida  (slight tatweel elongation)') +
  arabicPara('lowKashida') +
  heading('3) jc = mediumKashida  (moderate tatweel elongation)') +
  arabicPara('mediumKashida') +
  heading('4) jc = highKashida  (widest tatweel elongation)') +
  arabicPara('highKashida');

const documentXml =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:document xmlns:w="${W}"><w:body>${body}` +
  `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
  `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>` +
  `</w:sectPr></w:body></w:document>`;

const stylesXml =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>` +
  `<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Arial"/>` +
  `<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;

const contentTypes =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
  `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
  `</Types>`;

const rootRels =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
  `</Relationships>`;

const docRels =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
  `</Relationships>`;

async function main() {
  const buf = zip([
    { name: '[Content_Types].xml', data: contentTypes },
    { name: '_rels/.rels', data: rootRels },
    { name: 'word/_rels/document.xml.rels', data: docRels },
    { name: 'word/document.xml', data: documentXml },
    { name: 'word/styles.xml', data: stylesXml },
  ]);
  await mkdir(OUT_DIR, { recursive: true });
  await writeFile(OUT, buf);
  console.log(`wrote ${OUT} (${buf.length} bytes)`);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
