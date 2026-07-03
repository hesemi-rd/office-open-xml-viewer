import { describe, it, expect } from 'vitest';
import { DocxDocument } from './document';
import { OoxmlError, decryptOoxml } from '@silurus/ooxml-core';
import { encryptedDocxSpin0 } from '@silurus/ooxml-core/testing';

/**
 * End-to-end decryption wiring for `DocxDocument.load` ([MS-OFFCRYPTO] Agile
 * Encryption, PD8).
 *
 * The fixture is a real Agile-encrypted .docx (password `test`, spinCount 0)
 * whose bytes are base64-embedded in core so no binary Office file lands in git.
 * It was cross-verified: the independent `msoffcrypto-tool` decryptor recovers
 * the exact original .docx (see `packages/core/src/crypto/decrypt-ooxml.test.ts`,
 * which inflates the decrypted `word/document.xml` back to "Hello encrypted").
 *
 * These tests exercise the load()-level behaviour. The three outcomes are
 * distinguished on the main thread *before* the parser worker is constructed —
 * except the correct-password case, which decrypts successfully and then
 * proceeds to worker construction. In this Node test environment there is no
 * `Worker`, so a successful decrypt surfaces as a plain `ReferenceError`
 * ("Worker is not defined"), never an `OoxmlError`. That is the reliable signal
 * that decryption succeeded and the plaintext ZIP was handed onward to parse.
 *
 * Configuration coupling: this signal depends on vitest's default `environment:
 * 'node'` (no `Worker` global). If this suite's environment is ever switched to
 * `'jsdom'` / `'happy-dom'` (which polyfill `Worker`), the assertion below stops
 * proving what it claims — worker construction would no longer throw, so a
 * failure to reach the worker at all would go undetected.
 */
describe('DocxDocument.load — Agile decryption', () => {
  const fixture = () => encryptedDocxSpin0().buffer;

  it('throws OoxmlError("encrypted") when no password is supplied', async () => {
    await expect(DocxDocument.load(fixture())).rejects.toMatchObject({
      code: 'encrypted',
    });
    await expect(DocxDocument.load(fixture())).rejects.toBeInstanceOf(OoxmlError);
  });

  it('throws OoxmlError("invalid-password") for a wrong password', async () => {
    await expect(DocxDocument.load(fixture(), { password: 'wrong' })).rejects.toMatchObject({
      code: 'invalid-password',
    });
  });

  it('decrypts with the correct password and proceeds past the CFB guard to parse', async () => {
    // A successful decrypt reaches worker construction; the only failure left is
    // the missing Worker in Node — crucially NOT an OoxmlError, proving the
    // container guard accepted the decrypted plaintext.
    let error: unknown;
    try {
      await DocxDocument.load(fixture(), { password: 'test' });
    } catch (e) {
      error = e;
    }
    expect(error).toBeDefined();
    expect(error).not.toBeInstanceOf(OoxmlError);
  });

  it('the underlying decryptOoxml yields a plaintext OOXML ZIP for the same fixture', async () => {
    // Format-level sanity at the docx layer: PK magic + the expected ZIP entry
    // names (the definitive byte-for-byte check lives in core).
    const result = await decryptOoxml(encryptedDocxSpin0(), 'test');
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.data[0]).toBe(0x50); // 'P'
    expect(result.data[1]).toBe(0x4b); // 'K'
    const text = new TextDecoder('latin1').decode(result.data);
    expect(text).toContain('word/document.xml');
  });
});
