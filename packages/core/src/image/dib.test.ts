import { describe, it, expect } from 'vitest';
import { decodeDib, decodePackedDib, blitDibToCtx, type DecodedDib } from './dib.js';

// ── DIB decode unit tests ────────────────────────────────────────────────────
// Both metafile players embed raster bitmaps as DIBs (BITMAPINFOHEADER + optional
// palette + pixel bits). EMF passes explicit bmi/bits offsets → decodeDib; WMF
// stores the DIB *packed* (contiguous) → decodePackedDib, which derives the
// palette/bits offsets and delegates to decodeDib. We hand-build tiny DIBs and
// assert the returned top-down RGBA matches the BGR/palette source.
//
// blitDibToCtx needs OffscreenCanvas (absent in the node test env), so it is only
// asserted to return false there — pixel output is not testable here.

/** Little-endian byte writer for crafting DIB bytes. */
class Writer {
  private bytes: number[] = [];
  u8(v: number) {
    this.bytes.push(v & 0xff);
    return this;
  }
  u16(v: number) {
    this.bytes.push(v & 0xff, (v >>> 8) & 0xff);
    return this;
  }
  i32(v: number) {
    const u = v >>> 0;
    this.bytes.push(u & 0xff, (u >>> 8) & 0xff, (u >>> 16) & 0xff, (u >>> 24) & 0xff);
    return this;
  }
  u32(v: number) {
    return this.i32(v);
  }
  build(): Uint8Array {
    return new Uint8Array(this.bytes);
  }
  get length(): number {
    return this.bytes.length;
  }
}

/** BITMAPINFOHEADER (40 bytes). height<0 ⇒ top-down. */
function bmih(width: number, height: number, bitCount: number, clrUsed = 0): Writer {
  return new Writer()
    .u32(40) // biSize
    .i32(width) // biWidth
    .i32(height) // biHeight (negative = top-down)
    .u16(1) // biPlanes
    .u16(bitCount) // biBitCount
    .u32(0) // biCompression = BI_RGB
    .u32(0) // biSizeImage
    .i32(0) // biXPelsPerMeter
    .i32(0) // biYPelsPerMeter
    .u32(clrUsed) // biClrUsed
    .u32(0); // biClrImportant
}

/** A top-down 2×2 24-bit BI_RGB DIB. Pixels are given as [r,g,b] top-down,
 *  row-major; stored on disk as BGR with each row padded to a 4-byte boundary
 *  (rowStride = ((2*24+31)>>5)<<2 = 8 bytes, i.e. 6 pixel bytes + 2 pad). */
function dib24(px: [number, number, number][]): { bytes: Uint8Array; bitsOff: number } {
  const w = new Writer();
  // header (top-down: negative height)
  const header = bmih(2, -2, 24).build();
  for (const b of header) w.u8(b);
  const bitsOff = w.length;
  // row 0 = px[0],px[1]; row 1 = px[2],px[3]; each pixel BGR; pad row to 8 bytes.
  for (let row = 0; row < 2; row++) {
    for (let col = 0; col < 2; col++) {
      const [r, g, b] = px[row * 2 + col];
      w.u8(b).u8(g).u8(r); // BGR on disk
    }
    w.u8(0).u8(0); // 2 pad bytes → 8-byte row stride
  }
  return { bytes: w.build(), bitsOff };
}

function dvOf(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

function rgbaAt(dib: DecodedDib, x: number, y: number): [number, number, number, number] {
  const i = (y * dib.width + x) * 4;
  return [dib.data[i], dib.data[i + 1], dib.data[i + 2], dib.data[i + 3]];
}

describe('decodeDib — explicit bmi/bits offsets (EMF-style)', () => {
  it('decodes a top-down 2×2 24-bit BI_RGB DIB, mapping BGR→RGBA (row 4-byte aligned)', () => {
    const pixels: [number, number, number][] = [
      [255, 0, 0], // (0,0) red
      [0, 255, 0], // (1,0) green
      [0, 0, 255], // (0,1) blue
      [10, 20, 30], // (1,1)
    ];
    const { bytes, bitsOff } = dib24(pixels);
    const dv = dvOf(bytes);
    const dib = decodeDib(dv, 0, 40, bitsOff, bytes.length - bitsOff);
    expect(dib).not.toBeNull();
    const d = dib as DecodedDib;
    expect(d.width).toBe(2);
    expect(d.height).toBe(2);
    // Top-down: dst row order matches src row order.
    expect(rgbaAt(d, 0, 0)).toEqual([255, 0, 0, 255]);
    expect(rgbaAt(d, 1, 0)).toEqual([0, 255, 0, 255]);
    expect(rgbaAt(d, 0, 1)).toEqual([0, 0, 255, 255]);
    expect(rgbaAt(d, 1, 1)).toEqual([10, 20, 30, 255]);
  });

  it('flips a BOTTOM-UP DIB so the returned data is top-down', () => {
    // Same pixels but authored bottom-up (positive height). Build manually: on
    // disk the FIRST stored row is the BOTTOM (y=1) row.
    const w = new Writer();
    for (const b of bmih(2, 2, 24).build()) w.u8(b);
    const bitsOff = w.length;
    // stored row 0 (disk) = image bottom row (y=1): blue, then [10,20,30]
    w.u8(255).u8(0).u8(0).u8(30).u8(20).u8(10).u8(0).u8(0); // BGR: blue=(0,0,255) → B=255; (10,20,30)→B=30,G=20,R=10
    // stored row 1 (disk) = image top row (y=0): red, then green
    w.u8(0).u8(0).u8(255).u8(0).u8(255).u8(0).u8(0).u8(0);
    const bytes = w.build();
    const dib = decodeDib(dvOf(bytes), 0, 40, bitsOff, bytes.length - bitsOff);
    const d = dib as DecodedDib;
    // After flip, y=0 is the top (red/green), y=1 the bottom (blue/…).
    expect(rgbaAt(d, 0, 0)).toEqual([255, 0, 0, 255]);
    expect(rgbaAt(d, 1, 0)).toEqual([0, 255, 0, 255]);
    expect(rgbaAt(d, 0, 1)).toEqual([0, 0, 255, 255]);
    expect(rgbaAt(d, 1, 1)).toEqual([10, 20, 30, 255]);
  });

  it('returns null for a non-BI_RGB (compressed) header', () => {
    const w = new Writer()
      .u32(40)
      .i32(2)
      .i32(-2)
      .u16(1)
      .u16(24)
      .u32(1) // biCompression = BI_RLE8 (non-zero → unsupported)
      .u32(0)
      .i32(0)
      .i32(0)
      .u32(0)
      .u32(0);
    for (let i = 0; i < 16; i++) w.u8(0);
    const bytes = w.build();
    expect(decodeDib(dvOf(bytes), 0, 40, 40, bytes.length - 40)).toBeNull();
  });
});

describe('decodePackedDib — contiguous header+palette+bits (WMF-style)', () => {
  it('decodes a packed 24-bit DIB by deriving bitsOff after the 40-byte header (no palette)', () => {
    const { bytes } = dib24([
      [255, 0, 0],
      [0, 255, 0],
      [0, 0, 255],
      [10, 20, 30],
    ]);
    const dib = decodePackedDib(dvOf(bytes), 0, bytes.length);
    expect(dib).not.toBeNull();
    const d = dib as DecodedDib;
    expect(d.width).toBe(2);
    expect(d.height).toBe(2);
    expect(rgbaAt(d, 0, 0)).toEqual([255, 0, 0, 255]);
    expect(rgbaAt(d, 1, 1)).toEqual([10, 20, 30, 255]);
  });

  it('decodes an 8-bit palette DIB, skipping the RGBQUAD palette to reach the bits', () => {
    // 2×2, 8bpp, top-down. Palette of 3 entries (clrUsed=3): index0=red,
    // 1=green, 2=blue (RGBQUAD is B,G,R,reserved). rowStride = ((2*8+31)>>5)<<2
    // = 4 bytes (2 index bytes + 2 pad).
    const w = new Writer();
    for (const b of bmih(2, -2, 8, 3).build()) w.u8(b);
    // palette (BGR + reserved):
    w.u8(0).u8(0).u8(255).u8(0); // idx0 red
    w.u8(0).u8(255).u8(0).u8(0); // idx1 green
    w.u8(255).u8(0).u8(0).u8(0); // idx2 blue
    // pixel rows (indices), padded to 4 bytes each:
    w.u8(0).u8(1).u8(0).u8(0); // row0: red, green
    w.u8(2).u8(1).u8(0).u8(0); // row1: blue, green
    const bytes = w.build();
    const dib = decodePackedDib(dvOf(bytes), 0, bytes.length);
    expect(dib).not.toBeNull();
    const d = dib as DecodedDib;
    expect(rgbaAt(d, 0, 0)).toEqual([255, 0, 0, 255]); // red
    expect(rgbaAt(d, 1, 0)).toEqual([0, 255, 0, 255]); // green
    expect(rgbaAt(d, 0, 1)).toEqual([0, 0, 255, 255]); // blue
    expect(rgbaAt(d, 1, 1)).toEqual([0, 255, 0, 255]); // green
  });

  it('skips an OPTIONAL optimization color table on a >8bpp DIB (biClrUsed > 0)', () => {
    // A 24-bit DIB carries no indexed palette, but MAY prepend an optional
    // optimization color table when biClrUsed > 0; the pixel bits then follow
    // AFTER it ([MS-WMF] 2.2.2.9). clrUsed=2 → 8 palette bytes that must be
    // skipped, else the bits offset misaligns and the decode is garbage.
    const w = new Writer();
    for (const b of bmih(2, -2, 24, 2).build()) w.u8(b); // clrUsed = 2
    w.u8(1).u8(2).u8(3).u8(0).u8(4).u8(5).u8(6).u8(0);    // 2 RGBQUAD entries (skipped)
    const px: [number, number, number][] = [
      [10, 20, 30], [40, 50, 60], [70, 80, 90], [100, 110, 120],
    ];
    for (let row = 0; row < 2; row++) {
      for (let col = 0; col < 2; col++) {
        const [r, g, b] = px[row * 2 + col];
        w.u8(b).u8(g).u8(r);
      }
      w.u8(0).u8(0); // row pad → 8-byte stride
    }
    const bytes = w.build();
    const dib = decodePackedDib(dvOf(bytes), 0, bytes.length);
    expect(dib).not.toBeNull();
    expect(rgbaAt(dib as DecodedDib, 0, 0)).toEqual([10, 20, 30, 255]);
    expect(rgbaAt(dib as DecodedDib, 1, 1)).toEqual([100, 110, 120, 255]);
  });

  it('returns null when dibLen is too small for a header', () => {
    const bytes = new Uint8Array(20);
    expect(decodePackedDib(dvOf(bytes), 0, 20)).toBeNull();
  });
});

describe('decodeDib — dimension / megapixel budget (DoS guard)', () => {
  // A crafted BITMAPINFOHEADER can declare enormous dimensions while carrying
  // only a few bytes of pixel data. The decode must REJECT such headers before
  // it reaches `new Uint8ClampedArray(width * height * 4)`, otherwise a 65535×
  // 65535 DIB demands a ~16 GiB RGBA buffer and OOMs the tab.

  it('returns null for a megapixel-budget-exceeding header (65535×65535, tiny pixel data)', () => {
    // Header declares a 65535×65535 image (≈4.29e9 px → ~17 GB RGBA) but the
    // buffer only holds the 40-byte header + a couple of pixel bytes. Before the
    // fix, decodeDib allocates the giant Uint8ClampedArray up front and aborts.
    const w = new Writer();
    for (const b of bmih(65535, 65535, 24).build()) w.u8(b);
    w.u8(0).u8(0).u8(0).u8(0); // 4 stray pixel bytes — nowhere near the claimed size
    const bytes = w.build();
    expect(decodeDib(dvOf(bytes), 0, 40, 40, bytes.length - 40)).toBeNull();
  });

  it('returns null when a single dimension exceeds the max canvas dimension (40000 wide, 1 tall)', () => {
    // 40000 > 32767 (browser max canvas dimension) but < the old 65536 cap, so
    // the PRE-fix code accepted it. Height is 1 and the row buffer is fully
    // supplied, so the buffer-bounds check does NOT bail — decode would succeed
    // pre-fix (returning a 40000×1 dib) and only the new dimension cap rejects
    // it. Such a DIB could never be blitted to a canvas anyway (>32767).
    const width = 40000;
    const rowStride = (((width * 24 + 31) >> 5) << 2) >>> 0; // 4-byte aligned
    const w = new Writer();
    for (const b of bmih(width, 1, 24).build()) w.u8(b); // 40000 × 1, bottom-up
    for (let i = 0; i < rowStride; i++) w.u8(0); // one full, in-bounds pixel row
    const bytes = w.build();
    expect(decodeDib(dvOf(bytes), 0, 40, 40, bytes.length - 40)).toBeNull();
  });

  it('still decodes a large-but-in-budget DIB (within dimension and megapixel caps)', () => {
    // 2×2 is trivially in budget; this asserts the new guard does not reject
    // legitimate small images. (The existing 24-bit tests already cover the
    // happy path; this is an explicit regression sentinel for the guard.)
    const { bytes, bitsOff } = dib24([
      [1, 2, 3],
      [4, 5, 6],
      [7, 8, 9],
      [10, 11, 12],
    ]);
    const dib = decodeDib(dvOf(bytes), 0, 40, bitsOff, bytes.length - bitsOff);
    expect(dib).not.toBeNull();
    expect((dib as DecodedDib).width).toBe(2);
    expect((dib as DecodedDib).height).toBe(2);
  });
});

describe('blitDibToCtx — OffscreenCanvas absent', () => {
  it('returns false when OffscreenCanvas is undefined (node test env)', () => {
    expect(typeof OffscreenCanvas).toBe('undefined');
    const dib: DecodedDib = { width: 1, height: 1, data: new Uint8ClampedArray([1, 2, 3, 4]) };
    // A minimal stub ctx — should never be touched when OffscreenCanvas is absent.
    const ctx = { drawImage() { throw new Error('should not draw'); } } as unknown as CanvasRenderingContext2D;
    expect(blitDibToCtx(ctx, dib, 0, 0, 1, 1)).toBe(false);
  });
});
