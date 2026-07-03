/**
 * Synthetic Compound File Binary (CFB / OLE2) builder for tests.
 *
 * Emits a minimal compound file — major version 3 (512-byte sector) by
 * default, or version 4 (4096-byte sector) via `opts.majorVersion` — whose FAT
 * lives in sector 0 and whose directory chain starts at sector 1, containing
 * the given directory-entry names. Just enough of [MS-CFB] to drive
 * `sniffCfb` and the `assertNotCfbContainer` load guard — not a general CFB
 * writer.
 *
 * Shared here (rather than duplicated per package) so the docx / pptx / xlsx
 * load-guard tests build their fixtures the same way. Test-only; never imported
 * by production code.
 */

const ENTRY = 128;
const FREESECT = 0xffffffff;
const ENDOFCHAIN = 0xfffffffe;
const FATSECT = 0xfffffffd;

const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

const MINI_SECTOR = 64;
const MINI_CUTOFF = 0x1000; // 4096

export interface CfbFixtureOptions {
  /** [MS-CFB] §2.2 major version: 3 (512-byte sector, SectorShift 0x0009) or
   *  4 (4096-byte sector, SectorShift 0x000C). Default 3. */
  majorVersion?: 3 | 4;
}

/**
 * Build a synthetic CFB whose directory contains `names` (entry 0 is treated as
 * the root storage, the rest as streams). The result is an `ArrayBuffer` ready
 * to hand to a `load()` factory.
 */
export function buildCfbFixture(names: string[], opts: CfbFixtureOptions = {}): ArrayBuffer {
  const majorVersion = opts.majorVersion ?? 3;
  // §2.2: SectorShift is 0x0009 (512 B) for v3, 0x000C (4096 B) for v4.
  const sectorShift = majorVersion === 4 ? 12 : 9;
  const sectorSize = 1 << sectorShift;

  const entriesPerSector = sectorSize / ENTRY;
  const dirSectors = Math.max(1, Math.ceil(names.length / entriesPerSector));
  // Logical sectors used: 0 (FAT), 1..dirSectors (directory chain).
  const lastSector = dirSectors;
  // Layout matches sniffCfb's fileOffsetOfSector: logical sector N starts at
  // byte (N + 1) * sectorSize, i.e. the header occupies "sector -1" — a full
  // `sectorSize`-sized slot, not just its own 512 bytes. For v3 that slot IS
  // exactly 512 bytes (sectorSize === the header size). For v4 (sectorSize
  // 4096) the 512-byte header is followed by 3584 bytes of padding (left
  // zeroed) before sector 0 begins at byte 4096.
  const sectorOffset = (n: number): number => (n + 1) * sectorSize;
  const buf = new ArrayBuffer(sectorOffset(lastSector) + sectorSize);
  const view = new DataView(buf);
  const bytes = new Uint8Array(buf);

  // Header (§2.2): always 512 bytes, written at the start of the first
  // sectorSize-sized slot regardless of major version.
  for (let i = 0; i < 8; i++) bytes[i] = SIGNATURE[i];
  view.setUint16(0x18, 0x003e, true); // minor version
  view.setUint16(0x1a, majorVersion, true); // major version
  view.setUint16(0x1c, 0xfffe, true); // byte order
  view.setUint16(0x1e, sectorShift, true); // sector shift
  view.setUint16(0x20, 6, true); // mini sector shift
  view.setUint32(0x2c, 1, true); // number of FAT sectors
  view.setUint32(0x30, 1, true); // first directory sector = 1
  view.setUint32(0x38, 0x00001000, true); // mini stream cutoff
  view.setUint32(0x3c, ENDOFCHAIN, true); // first mini FAT
  view.setUint32(0x44, ENDOFCHAIN, true); // first DIFAT
  view.setUint32(0x4c, 0, true); // DIFAT[0] = FAT at sector 0
  for (let i = 1; i < 109; i++) view.setUint32(0x4c + i * 4, FREESECT, true);

  // FAT (logical sector 0): sector 0 is the FAT; directory chain 1..dirSectors.
  const fatOff = sectorOffset(0);
  for (let i = 0; i < sectorSize / 4; i++) view.setUint32(fatOff + i * 4, FREESECT, true);
  view.setUint32(fatOff, FATSECT, true);
  for (let s = 1; s <= dirSectors; s++) {
    view.setUint32(fatOff + s * 4, s < dirSectors ? s + 1 : ENDOFCHAIN, true);
  }

  // Directory entries.
  names.forEach((name, idx) => {
    const logicalSector = 1 + Math.floor(idx / entriesPerSector);
    const within = idx % entriesPerSector;
    const off = sectorOffset(logicalSector) + within * ENTRY;
    const units = Math.min(name.length, 31);
    for (let i = 0; i < units; i++) view.setUint16(off + i * 2, name.charCodeAt(i), true);
    view.setUint16(off + 0x40, (units + 1) * 2, true); // name byte length incl. NUL
    view.setUint8(off + 0x42, idx === 0 ? 5 : 2); // root storage (5) / stream (2)
  });

  return buf;
}

/** A named stream with its raw bytes, for {@link buildCfbWithStreams}. */
export interface CfbStream {
  name: string;
  data: Uint8Array;
}

/**
 * Build a version-3 CFB carrying real stream *contents*, laid out the way real
 * Office produces an encrypted OOXML: streams shorter than the 4096-byte
 * mini-stream cutoff live in the root storage's mini stream (64-byte mini
 * sectors, chained through the mini FAT); larger streams live in the regular
 * FAT. This exercises {@link import('../errors/cfb-read').readCfbStream}'s mini
 * FAT / mini stream traversal, which the classification-only
 * {@link buildCfbFixture} does not.
 *
 * Fixed layout (all sectors 512 B): sector 0 = FAT, sector 1 = directory,
 * sector 2 = mini FAT, sectors 3.. = big-stream data, then the mini-stream
 * container. Test-only; mirrors the CFB writer used to generate the encrypted
 * fixtures.
 */
export function buildCfbWithStreams(streams: CfbStream[]): ArrayBuffer {
  const SECTOR = 512;
  const sectorOffset = (n: number): number => (n + 1) * SECTOR;

  const big = streams.filter((s) => s.data.length >= MINI_CUTOFF);
  const mini = streams.filter((s) => s.data.length < MINI_CUTOFF);

  // FAT plan: 0 = FAT, 1 = directory, 2 = mini FAT.
  const fat = new Map<number, number>([
    [0, FATSECT],
    [1, ENDOFCHAIN],
    [2, ENDOFCHAIN],
  ]);
  let nextSector = 3;

  // Big streams: consecutive FAT sectors.
  const bigLayout = new Map<string, { start: number; size: number; padded: Uint8Array }>();
  for (const s of big) {
    const nsec = Math.ceil(s.data.length / SECTOR);
    const start = nextSector;
    for (let i = 0; i < nsec; i++) fat.set(start + i, i < nsec - 1 ? start + i + 1 : ENDOFCHAIN);
    const padded = new Uint8Array(nsec * SECTOR);
    padded.set(s.data);
    bigLayout.set(s.name, { start, size: s.data.length, padded });
    nextSector += nsec;
  }

  // Mini streams: packed into 64-byte mini sectors, chained through the mini FAT.
  const miniFat: number[] = [];
  const miniLayout = new Map<string, { start: number; size: number }>();
  const miniChunks: Uint8Array[] = [];
  let miniIdx = 0;
  for (const s of mini) {
    const nmini = Math.ceil(s.data.length / MINI_SECTOR) || 0;
    const start = miniIdx;
    for (let i = 0; i < nmini; i++) miniFat.push(i < nmini - 1 ? start + i + 1 : ENDOFCHAIN);
    const chunk = new Uint8Array(nmini * MINI_SECTOR);
    chunk.set(s.data);
    miniChunks.push(chunk);
    miniLayout.set(s.name, { start, size: s.data.length });
    miniIdx += nmini;
  }
  const miniStream = concatBytes(miniChunks);

  // Mini-stream container in the regular FAT.
  const miniStreamSectors = Math.ceil(miniStream.length / SECTOR) || 0;
  const miniStreamStart = miniStreamSectors ? nextSector : ENDOFCHAIN;
  if (miniStreamSectors) {
    for (let i = 0; i < miniStreamSectors; i++) {
      fat.set(nextSector + i, i < miniStreamSectors - 1 ? nextSector + i + 1 : ENDOFCHAIN);
    }
    nextSector += miniStreamSectors;
  }
  const miniStreamPadded = new Uint8Array(miniStreamSectors * SECTOR);
  miniStreamPadded.set(miniStream);

  const buf = new ArrayBuffer(sectorOffset(nextSector - 1) + SECTOR);
  const view = new DataView(buf);
  const bytes = new Uint8Array(buf);

  // Header.
  for (let i = 0; i < 8; i++) bytes[i] = SIGNATURE[i];
  view.setUint16(0x18, 0x003e, true); // minor version
  view.setUint16(0x1a, 3, true); // major version 3
  view.setUint16(0x1c, 0xfffe, true); // byte order
  view.setUint16(0x1e, 9, true); // sector shift (512)
  view.setUint16(0x20, 6, true); // mini sector shift (64)
  view.setUint32(0x2c, 1, true); // number of FAT sectors
  view.setUint32(0x30, 1, true); // first directory sector
  view.setUint32(0x38, MINI_CUTOFF, true); // mini stream cutoff
  view.setUint32(0x3c, 2, true); // first mini FAT sector
  view.setUint32(0x40, 1, true); // number of mini FAT sectors
  view.setUint32(0x44, ENDOFCHAIN, true); // first DIFAT sector
  view.setUint32(0x4c, 0, true); // DIFAT[0] = FAT at sector 0
  for (let i = 1; i < 109; i++) view.setUint32(0x4c + i * 4, FREESECT, true);

  // FAT (sector 0).
  const fatOff = sectorOffset(0);
  for (let i = 0; i < SECTOR / 4; i++) view.setUint32(fatOff + i * 4, FREESECT, true);
  for (const [sec, val] of fat) view.setUint32(fatOff + sec * 4, val, true);

  // Mini FAT (sector 2).
  const miniFatOff = sectorOffset(2);
  for (let i = 0; i < SECTOR / 4; i++) view.setUint32(miniFatOff + i * 4, FREESECT, true);
  miniFat.forEach((v, i) => view.setUint32(miniFatOff + i * 4, v, true));

  // Directory (sector 1): Root Entry then one entry per stream.
  const dirOff = sectorOffset(1);
  const ordered = [...big, ...mini];
  writeDirEntry(view, dirOff, 'Root Entry', 5, miniStreamStart, miniStream.length);
  ordered.forEach((s, i) => {
    const off = dirOff + (i + 1) * ENTRY;
    const layout = bigLayout.get(s.name) ?? miniLayout.get(s.name);
    if (!layout) return;
    writeDirEntry(view, off, s.name, 2, layout.start, layout.size);
  });

  // Big stream data.
  for (const s of big) {
    const layout = bigLayout.get(s.name);
    if (layout) bytes.set(layout.padded, sectorOffset(layout.start));
  }
  // Mini stream container.
  if (miniStreamSectors) bytes.set(miniStreamPadded, sectorOffset(miniStreamStart));

  return buf;
}

function writeDirEntry(
  view: DataView,
  off: number,
  name: string,
  objType: number,
  startSector: number,
  size: number,
): void {
  const units = Math.min(name.length, 31);
  for (let i = 0; i < units; i++) view.setUint16(off + i * 2, name.charCodeAt(i), true);
  view.setUint16(off + 0x40, (units + 1) * 2, true);
  view.setUint8(off + 0x42, objType);
  view.setUint32(off + 0x74, startSector, true); // starting sector
  view.setUint32(off + 0x78, size, true); // stream size (low 32 bits)
}

function concatBytes(chunks: Uint8Array[]): Uint8Array {
  const total = chunks.reduce((n, c) => n + c.length, 0);
  const out = new Uint8Array(total);
  let o = 0;
  for (const c of chunks) {
    out.set(c, o);
    o += c.length;
  }
  return out;
}
