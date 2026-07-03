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
