import { describe, it, expect } from 'vitest';
import { readCfbStream } from './cfb-read';
import { buildCfbWithStreams } from '../testing/cfb-fixture';

/** Deterministic pseudo-random bytes (mulberry32) so failures reproduce. */
function mulberry32(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) | 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
function randomBytes(rand: () => number, length: number): Uint8Array {
  const b = new Uint8Array(length);
  for (let i = 0; i < length; i++) b[i] = Math.floor(rand() * 256);
  return b;
}

describe('readCfbStream — happy path', () => {
  it('returns null for non-CFB (ZIP) bytes', () => {
    expect(readCfbStream(new Uint8Array([0x50, 0x4b, 0x03, 0x04]), 'EncryptionInfo')).toBeNull();
  });

  it('returns null for a too-short buffer', () => {
    expect(readCfbStream(new Uint8Array(16), 'X')).toBeNull();
  });

  it('reads a small stream stored in the mini stream (< 4096-byte cutoff)', () => {
    const small = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
    const cfb = new Uint8Array(buildCfbWithStreams([{ name: 'EncryptionInfo', data: small }]));
    expect(readCfbStream(cfb, 'EncryptionInfo')).toEqual(small);
  });

  it('reads a mini stream that spans several 64-byte mini sectors', () => {
    // 200 bytes => 4 mini sectors (64*4 = 256), exercises the mini-FAT chain.
    const rand = mulberry32(1);
    const data = randomBytes(rand, 200);
    const cfb = new Uint8Array(buildCfbWithStreams([{ name: 'EncryptionInfo', data }]));
    expect(readCfbStream(cfb, 'EncryptionInfo')).toEqual(data);
  });

  it('reads a large stream stored in the regular FAT (>= cutoff)', () => {
    const rand = mulberry32(2);
    const data = randomBytes(rand, 10_000); // spans ~20 FAT sectors
    const cfb = new Uint8Array(buildCfbWithStreams([{ name: 'EncryptedPackage', data }]));
    expect(readCfbStream(cfb, 'EncryptedPackage')).toEqual(data);
  });

  it('reads both a mini and a FAT stream from the same container', () => {
    const rand = mulberry32(3);
    const info = randomBytes(rand, 1289); // < cutoff => mini stream
    const pkg = randomBytes(rand, 5000); // >= cutoff => FAT
    const cfb = new Uint8Array(
      buildCfbWithStreams([
        { name: 'EncryptedPackage', data: pkg },
        { name: 'EncryptionInfo', data: info },
      ]),
    );
    expect(readCfbStream(cfb, 'EncryptionInfo')).toEqual(info);
    expect(readCfbStream(cfb, 'EncryptedPackage')).toEqual(pkg);
  });

  it('returns null for a stream name that is not present', () => {
    const cfb = new Uint8Array(buildCfbWithStreams([{ name: 'EncryptionInfo', data: new Uint8Array([1]) }]));
    expect(readCfbStream(cfb, 'NoSuchStream')).toBeNull();
  });

  it('returns an empty array for a zero-length stream', () => {
    const cfb = new Uint8Array(buildCfbWithStreams([{ name: 'Empty', data: new Uint8Array(0) }]));
    expect(readCfbStream(cfb, 'Empty')).toEqual(new Uint8Array(0));
  });

  it('reads a stream at exactly the cutoff boundary from the FAT', () => {
    // 4096 == cutoff => FAT (< cutoff is mini; the boundary goes to FAT).
    const rand = mulberry32(4);
    const data = randomBytes(rand, 4096);
    const cfb = new Uint8Array(buildCfbWithStreams([{ name: 'Boundary', data }]));
    expect(readCfbStream(cfb, 'Boundary')).toEqual(data);
  });
});

describe('readCfbStream — robustness (never throws, returns null on corruption)', () => {
  it('returns null when the sector shift is corrupted', () => {
    const cfb = new Uint8Array(buildCfbWithStreams([{ name: 'X', data: new Uint8Array([1]) }]));
    new DataView(cfb.buffer).setUint16(0x1e, 30, true); // 1 GiB sectors
    expect(readCfbStream(cfb, 'X')).toBeNull();
  });

  it('returns null when the target start sector points past EOF', () => {
    const rand = mulberry32(5);
    const data = randomBytes(rand, 6000);
    const cfb = new Uint8Array(buildCfbWithStreams([{ name: 'EncryptedPackage', data }]));
    // Corrupt the EncryptedPackage directory entry's starting sector (@0x74).
    // Directory is sector 1 (offset 1024); root is entry 0, package is entry 1.
    const dirOff = 2 * 512; // sectorOffset(1) = (1+1)*512
    new DataView(cfb.buffer).setUint32(dirOff + 128 + 0x74, 0x0fffffff, true);
    expect(readCfbStream(cfb, 'EncryptedPackage')).toBeNull();
  });

  it('never throws over deterministic-seed garbage with the CFB signature forced', () => {
    const rand = mulberry32(0xbeef);
    for (let i = 0; i < 5000; i++) {
      const len = 512 + Math.floor(rand() * 4096);
      const bytes = randomBytes(rand, len);
      const sig = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
      for (let j = 0; j < sig.length; j++) bytes[j] = sig[j];
      let result: Uint8Array | null | undefined;
      expect(() => {
        result = readCfbStream(bytes, 'EncryptionInfo');
      }).not.toThrow();
      expect(result === null || result instanceof Uint8Array).toBe(true);
    }
  });
});
