/**
 * [MS-OFFCRYPTO] Agile Encryption decryption for password-protected OOXML.
 *
 * Public surface for the load layer (`decryptOoxml`) plus the lower-level
 * primitives, exposed so they can be unit-tested against pinned test vectors.
 * All work runs on WebCrypto (`globalThis.crypto.subtle`); Rust is unchanged.
 */
export {
  parseEncryptionInfo,
  type EncryptionInfoKind,
  type AgileEncryptionDescriptor,
  type AgileCipherParams,
  type PasswordKeyEncryptor,
  type DataIntegrity,
} from './encryption-info';
export {
  deriveAgileKey,
  deriveIv,
  verifyPassword,
  deriveIntermediateKey,
  decryptPackage,
  decryptAgilePackage,
  AgileDecryptError,
  BLOCK_KEY,
  type DecryptFailure,
} from './agile';
export { decryptOoxml, type DecryptResult } from './decrypt-ooxml';
