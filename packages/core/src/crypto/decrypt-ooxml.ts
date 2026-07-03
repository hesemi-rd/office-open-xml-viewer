/**
 * Top-level entry point: turn an encrypted OOXML CFB container + a password
 * into the plaintext OOXML ZIP bytes, ready to hand to the normal ZIP parser.
 *
 * Pipeline:
 *   1. `readCfbStream` extracts `EncryptionInfo` and `EncryptedPackage`.
 *   2. `parseEncryptionInfo` classifies the version — only Agile (4.4) is
 *      supported; Standard / Extensible / anything else yields
 *      `'unsupported-encryption'`.
 *   3. `decryptAgilePackage` verifies the password and decrypts the package.
 *
 * The result is a discriminated union so the caller (the docx / pptx / xlsx
 * `load()` guard) can map cleanly onto a typed `OoxmlError` without inspecting
 * message strings.
 */
import { readCfbStream } from '../errors/cfb-read';
import { parseEncryptionInfo } from './encryption-info';
import { decryptAgilePackage, AgileDecryptError, type DecryptFailure } from './agile';

export type DecryptResult =
  | { ok: true; data: Uint8Array }
  | { ok: false; reason: DecryptFailure };

const ENCRYPTION_INFO = 'EncryptionInfo';
const ENCRYPTED_PACKAGE = 'EncryptedPackage';

/**
 * Decrypt an Agile-encrypted OOXML container. Returns `{ ok: true, data }` with
 * the plaintext ZIP bytes, or `{ ok: false, reason }`:
 *
 *   - `'invalid-password'`        — the password did not match the verifier.
 *   - `'unsupported-encryption'`  — not Agile (Standard / Extensible / other),
 *     or an unsupported cipher / hash within an Agile descriptor.
 *   - `'corrupt'`                 — the CFB streams are missing or malformed.
 *
 * Never throws; all failures come back as `{ ok: false }`.
 */
export async function decryptOoxml(bytes: Uint8Array, password: string): Promise<DecryptResult> {
  const infoBytes = readCfbStream(bytes, ENCRYPTION_INFO);
  const packageBytes = readCfbStream(bytes, ENCRYPTED_PACKAGE);
  if (infoBytes === null || packageBytes === null) {
    return { ok: false, reason: 'corrupt' };
  }

  const info = parseEncryptionInfo(infoBytes);
  if (info.kind !== 'agile') {
    // Standard / Extensible / unknown descriptor: explicitly unsupported.
    return { ok: false, reason: 'unsupported-encryption' };
  }

  try {
    const data = await decryptAgilePackage(info.descriptor, packageBytes, password);
    return { ok: true, data };
  } catch (e) {
    if (e instanceof AgileDecryptError) return { ok: false, reason: e.reason };
    // Any unexpected WebCrypto failure (e.g. a bad-padding throw on a wrong
    // key that slipped past the verifier) is treated as corrupt input.
    return { ok: false, reason: 'corrupt' };
  }
}
