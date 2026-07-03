import { describe, it, expect } from 'vitest';
import { resolveOoxmlContainer } from './cfb-guard';
import { OoxmlError } from './ooxml-error';
import { buildCfbFixture } from '../testing/cfb-fixture';
import { encryptedDocxSpin0 } from '../testing/encrypted-fixture';

/**
 * Unit coverage for {@link resolveOoxmlContainer}, the decrypt-aware superset of
 * `assertNotCfbContainer` that the docx / pptx / xlsx `load()` factories call.
 * `cfb-guard.test.ts` covers the no-password `assertNotCfbContainer` path only;
 * this file exercises the password-aware branches directly (rather than only
 * indirectly through the docx/pptx/xlsx `*-decrypt.test.ts` E2E suites), reusing
 * the same fixture helpers so all of these suites build CFB bytes the same way.
 */
describe('resolveOoxmlContainer', () => {
  it('passes a non-CFB (ZIP) buffer through unchanged', async () => {
    const zip = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0, 0, 0, 0]);
    const out = await resolveOoxmlContainer(zip);
    expect(out).toBe(zip); // same reference: no copy, no decrypt attempted
  });

  it('throws OoxmlError("encrypted") for an encrypted file when no password is supplied', async () => {
    const bytes = encryptedDocxSpin0();
    await expect(resolveOoxmlContainer(bytes)).rejects.toBeInstanceOf(OoxmlError);
    await expect(resolveOoxmlContainer(bytes)).rejects.toMatchObject({ code: 'encrypted' });
  });

  it('throws OoxmlError("invalid-password") for an encrypted file with the wrong password', async () => {
    const bytes = encryptedDocxSpin0();
    await expect(resolveOoxmlContainer(bytes, 'wrong')).rejects.toBeInstanceOf(OoxmlError);
    await expect(resolveOoxmlContainer(bytes, 'wrong')).rejects.toMatchObject({
      code: 'invalid-password',
    });
  });

  it('returns the decrypted plaintext ZIP for an encrypted file with the correct password', async () => {
    const bytes = encryptedDocxSpin0();
    const out = await resolveOoxmlContainer(bytes, 'test');
    expect(out[0]).toBe(0x50); // 'P'
    expect(out[1]).toBe(0x4b); // 'K'
  });

  it('throws OoxmlError("legacy-binary-format") for a legacy binary CFB, password or not', async () => {
    const cfb = new Uint8Array(buildCfbFixture(['Root Entry', 'WordDocument']));
    await expect(resolveOoxmlContainer(cfb)).rejects.toMatchObject({
      code: 'legacy-binary-format',
    });
    await expect(resolveOoxmlContainer(cfb, 'test')).rejects.toMatchObject({
      code: 'legacy-binary-format',
    });
  });

  it('accepts an ArrayBuffer as well as a Uint8Array', async () => {
    const bytes = encryptedDocxSpin0();
    const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    await expect(resolveOoxmlContainer(ab)).rejects.toMatchObject({ code: 'encrypted' });
  });
});
