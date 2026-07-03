import { describe, it, expect } from 'vitest';
import { XlsxWorkbook } from './workbook';
import { OoxmlError, decryptOoxml } from '@silurus/ooxml-core';
import { encryptedDocxSpin0 } from '@silurus/ooxml-core/testing';

/**
 * End-to-end decryption wiring for `XlsxWorkbook.load` ([MS-OFFCRYPTO] Agile
 * Encryption, PD8). Decryption is format-independent (one core engine for docx /
 * pptx / xlsx), so this asserts the xlsx load() plumbing routes through it. A
 * correct password decrypts the shared encrypted-OOXML fixture and load()
 * proceeds to worker construction (a plain `ReferenceError` in Node, never an
 * `OoxmlError`).
 *
 * Configuration coupling: the "reaches the worker" signal relies on vitest's
 * default `environment: 'node'` (no `Worker` global). Switching this suite to
 * `'jsdom'` / `'happy-dom'` would polyfill `Worker` and silently invalidate
 * that assertion.
 */
describe('XlsxWorkbook.load — Agile decryption', () => {
  const fixture = () => encryptedDocxSpin0().buffer;

  it('throws OoxmlError("encrypted") when no password is supplied', async () => {
    await expect(XlsxWorkbook.load(fixture())).rejects.toMatchObject({ code: 'encrypted' });
  });

  it('throws OoxmlError("invalid-password") for a wrong password', async () => {
    await expect(XlsxWorkbook.load(fixture(), { password: 'wrong' })).rejects.toMatchObject({
      code: 'invalid-password',
    });
  });

  it('decrypts with the correct password and proceeds past the CFB guard', async () => {
    let error: unknown;
    try {
      await XlsxWorkbook.load(fixture(), { password: 'test' });
    } catch (e) {
      error = e;
    }
    expect(error).toBeDefined();
    expect(error).not.toBeInstanceOf(OoxmlError);
  });

  it('decryptOoxml yields a plaintext OOXML ZIP for the same fixture', async () => {
    const result = await decryptOoxml(encryptedDocxSpin0(), 'test');
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.data[0]).toBe(0x50);
    expect(result.data[1]).toBe(0x4b);
  });
});
