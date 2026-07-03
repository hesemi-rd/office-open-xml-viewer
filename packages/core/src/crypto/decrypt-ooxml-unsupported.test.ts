import { describe, it, expect } from 'vitest';
import { decryptOoxml } from './decrypt-ooxml';
import { resolveOoxmlContainer } from '../errors/cfb-guard';
import { OoxmlError } from '../errors/ooxml-error';
import { buildCfbWithStreams } from '../testing/cfb-fixture';

/**
 * End-to-end coverage for the `'unsupported-encryption'` outcome: an
 * `EncryptionInfo` stream whose `EncryptionVersionInfo` (§2.3.4.10) is *not*
 * Agile (4.4), wrapped in a real CFB alongside a (structurally irrelevant, for
 * this classification) `EncryptedPackage` stream — the same shape a real
 * Standard- or Extensible-encrypted Office file has on disk.
 *
 * `parseEncryptionInfo` (packages/core/src/crypto/encryption-info.ts) only
 * needs the first 4 bytes (vMajor/vMinor, LE16 each) to classify the stream;
 * the rest of a real Standard-encryption descriptor is a binary (non-XML)
 * structure this library never parses, so it is omitted here — only the
 * version header is under test.
 */
function encryptionInfoHeader(vMajor: number, vMinor: number): Uint8Array {
  const bytes = new Uint8Array(8 + 4); // version header + reserved, no descriptor body
  const view = new DataView(bytes.buffer);
  view.setUint16(0, vMajor, true);
  view.setUint16(2, vMinor, true);
  view.setUint32(4, 0x00000040, true); // Reserved, per §2.3.4.10
  return bytes;
}

/** A minimal, well-formed EncryptedPackage stream: an 8-byte LE64 size prefix
 *  (declaring zero plaintext bytes) and no ciphertext. Its contents are never
 *  read for a non-Agile EncryptionInfo — `decryptOoxml` returns
 *  `unsupported-encryption` right after classifying the version, before
 *  touching the package bytes — but the stream must exist for
 *  `readCfbStream` to find it (a missing stream reports `corrupt` instead). */
function dummyEncryptedPackage(): Uint8Array {
  return new Uint8Array(8);
}

describe('decryptOoxml / resolveOoxmlContainer — Standard Encryption (unsupported)', () => {
  it.each([
    ['version major 3, minor 2 (Standard, modern)', 3, 2],
    ['version major 2, minor 2 (Standard, legacy)', 2, 2],
  ])('classifies a %s EncryptionInfo header as unsupported-encryption', async (_label, vMajor, vMinor) => {
    const cfb = new Uint8Array(
      buildCfbWithStreams([
        { name: 'EncryptionInfo', data: encryptionInfoHeader(vMajor, vMinor) },
        { name: 'EncryptedPackage', data: dummyEncryptedPackage() },
      ]),
    );

    const result = await decryptOoxml(cfb, 'irrelevant');
    expect(result).toEqual({ ok: false, reason: 'unsupported-encryption' });
  });

  it('resolveOoxmlContainer surfaces the same Standard-encryption CFB as OoxmlError("unsupported-encryption")', async () => {
    const cfb = new Uint8Array(
      buildCfbWithStreams([
        { name: 'EncryptionInfo', data: encryptionInfoHeader(3, 2) },
        { name: 'EncryptedPackage', data: dummyEncryptedPackage() },
      ]),
    );

    await expect(resolveOoxmlContainer(cfb, 'irrelevant')).rejects.toBeInstanceOf(OoxmlError);
    await expect(resolveOoxmlContainer(cfb, 'irrelevant')).rejects.toMatchObject({
      code: 'unsupported-encryption',
    });
  });
});
