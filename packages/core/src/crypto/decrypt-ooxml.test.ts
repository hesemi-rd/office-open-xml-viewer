import { describe, it, expect } from 'vitest';
import { inflateRawSync } from 'node:zlib';
import { decryptOoxml } from './decrypt-ooxml';
import { encryptedDocxSpin0 } from '../testing/encrypted-fixture';
import { buildCfbFixture } from '../testing/cfb-fixture';

/**
 * Extract and inflate a stored ZIP entry from a plain (unencrypted) OOXML ZIP,
 * to prove the decrypted bytes are the *exact* original document — not merely a
 * structurally valid ZIP. Minimal local-file-header walk (ECMA ZIP): signature
 * `PK\x03\x04`, method @+8, name length @+26, extra length @+28, name @+30.
 */
function readZipEntry(zip: Uint8Array, name: string): Uint8Array | null {
  const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
  let off = 0;
  while (off + 30 <= zip.length && view.getUint32(off, true) === 0x04034b50) {
    const method = view.getUint16(off + 8, true);
    const compSize = view.getUint32(off + 18, true);
    const nameLen = view.getUint16(off + 26, true);
    const extraLen = view.getUint16(off + 28, true);
    const nameBytes = zip.subarray(off + 30, off + 30 + nameLen);
    const entryName = new TextDecoder('latin1').decode(nameBytes);
    const dataOff = off + 30 + nameLen + extraLen;
    const data = zip.subarray(dataOff, dataOff + compSize);
    if (entryName === name) {
      return method === 8 ? new Uint8Array(inflateRawSync(data)) : data;
    }
    off = dataOff + compSize;
  }
  return null;
}

describe('decryptOoxml — top-level pipeline', () => {
  it('decrypts the fixture to the exact original .docx content', async () => {
    const result = await decryptOoxml(encryptedDocxSpin0(), 'test');
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    // PK magic, and the deflated document.xml inflates to the original text.
    expect(result.data[0]).toBe(0x50);
    expect(result.data[1]).toBe(0x4b);
    const doc = readZipEntry(result.data, 'word/document.xml');
    expect(doc).not.toBeNull();
    const text = new TextDecoder('utf-8').decode(doc as Uint8Array);
    expect(text).toContain('Hello encrypted');
  });

  it('returns invalid-password for the wrong password', async () => {
    const result = await decryptOoxml(encryptedDocxSpin0(), 'WRONG');
    expect(result).toEqual({ ok: false, reason: 'invalid-password' });
  });

  it('returns corrupt when the CFB has no EncryptionInfo / EncryptedPackage', async () => {
    // A CFB with only a stray stream — classification-only fixture, no crypto
    // streams present.
    const cfb = new Uint8Array(buildCfbFixture(['Root Entry', 'Mystery']));
    const result = await decryptOoxml(cfb, 'test');
    expect(result).toEqual({ ok: false, reason: 'corrupt' });
  });

  it('returns corrupt for non-CFB bytes', async () => {
    const result = await decryptOoxml(new Uint8Array([0x50, 0x4b, 0x03, 0x04]), 'test');
    expect(result).toEqual({ ok: false, reason: 'corrupt' });
  });
});
