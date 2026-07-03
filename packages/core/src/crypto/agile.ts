/**
 * [MS-OFFCRYPTO] Agile Encryption key derivation, password verification and
 * package decryption, implemented on `globalThis.crypto.subtle` (WebCrypto).
 *
 * Spec map (learn.microsoft.com/openspecs/office_file_formats/ms-offcrypto):
 *   - §2.3.4.11 Encryption Key Generation:
 *       H0    = H(salt + UTF16LE(password))
 *       Hn    = H(LE32(i) + Hn-1)   for i = 0 .. spinCount-1  (iterator PREPENDED)
 *       Hfinal= H(Hn + blockKey)
 *       key   = Hfinal truncated to keyBits/8, or right-padded with 0x36.
 *   - §2.3.4.12 IV Generation:
 *       blockKey present: IV = H(saltValue + blockKey); else IV = saltValue.
 *       Then pad with 0x36 / truncate to blockSize.
 *   - §2.3.4.13 PasswordKeyEncryptor (verifier / intermediate key blockKeys).
 *   - §2.3.4.14 DataIntegrity HMAC (blockKeys for encryptedHmacKey/Value).
 *   - §2.3.4.15 Data Encryption: EncryptedPackage = LE64(plaintextSize) then
 *       4096-byte segments, each AES-CBC with IV = H(keyData.saltValue +
 *       LE32(segmentIndex)) and the intermediate key.
 *
 * All AES-CBC here is done manually (WebCrypto's AES-CBC always applies PKCS#7
 * padding, but the Agile format is unpadded / zero-padded) by driving AES-ECB
 * over one block at a time through a tiny CBC layer — see `aesCbcDecryptNoPad`.
 */

import type {
  AgileEncryptionDescriptor,
  AgileCipherParams,
  PasswordKeyEncryptor,
  Bytes,
} from './encryption-info';

/** blockKey constants — [MS-OFFCRYPTO] §2.3.4.13 / §2.3.4.14. */
export const BLOCK_KEY = {
  /** encryptedVerifierHashInput (§2.3.4.13). */
  verifierHashInput: new Uint8Array([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]),
  /** encryptedVerifierHashValue (§2.3.4.13). */
  verifierHashValue: new Uint8Array([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]),
  /** encryptedKeyValue → intermediate key (§2.3.4.13). */
  keyValue: new Uint8Array([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]),
  /** encryptedHmacKey (§2.3.4.14). */
  hmacKey: new Uint8Array([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6]),
  /** encryptedHmacValue (§2.3.4.14). */
  hmacValue: new Uint8Array([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33]),
} as const;

const PAD_BYTE = 0x36;
const SEGMENT_LENGTH = 4096;

/** Reason a decrypt attempt failed, so the load layer can map to a typed error. */
export type DecryptFailure = 'invalid-password' | 'unsupported-encryption' | 'corrupt';

export class AgileDecryptError extends Error {
  readonly reason: DecryptFailure;
  constructor(reason: DecryptFailure, message: string) {
    super(message);
    this.name = 'AgileDecryptError';
    this.reason = reason;
  }
}

function subtle(): SubtleCrypto {
  const c = globalThis.crypto;
  if (!c || !c.subtle) {
    throw new AgileDecryptError(
      'unsupported-encryption',
      'WebCrypto (globalThis.crypto.subtle) is unavailable; cannot decrypt.',
    );
  }
  return c.subtle;
}

/** Map an Agile `hashAlgorithm` string to a WebCrypto digest name. */
function webCryptoHash(hashAlgorithm: string): 'SHA-1' | 'SHA-256' | 'SHA-384' | 'SHA-512' {
  const h = hashAlgorithm.toUpperCase().replace(/[-_]/g, '');
  switch (h) {
    case 'SHA512':
      return 'SHA-512';
    case 'SHA384':
      return 'SHA-384';
    case 'SHA256':
      return 'SHA-256';
    case 'SHA1':
      return 'SHA-1';
    default:
      throw new AgileDecryptError(
        'unsupported-encryption',
        `Unsupported hashAlgorithm "${hashAlgorithm}" (only SHA-1/256/384/512).`,
      );
  }
}

function assertSupportedCipher(p: AgileCipherParams): void {
  if (p.cipherAlgorithm.toUpperCase() !== 'AES') {
    throw new AgileDecryptError(
      'unsupported-encryption',
      `Unsupported cipherAlgorithm "${p.cipherAlgorithm}" (only AES).`,
    );
  }
  const chain = p.cipherChaining.toLowerCase();
  if (chain !== 'chainingmodecbc') {
    throw new AgileDecryptError(
      'unsupported-encryption',
      `Unsupported cipherChaining "${p.cipherChaining}" (only ChainingModeCBC).`,
    );
  }
  if (p.keyBits !== 128 && p.keyBits !== 192 && p.keyBits !== 256) {
    throw new AgileDecryptError(
      'unsupported-encryption',
      `Unsupported keyBits ${p.keyBits} (only 128/192/256).`,
    );
  }
}

function concat(...parts: Uint8Array[]): Bytes {
  const total = parts.reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(total);
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

function le32(n: number): Bytes {
  const b = new Uint8Array(4);
  new DataView(b.buffer).setUint32(0, n >>> 0, true);
  return b;
}

/** UTF-16LE bytes of a JS string (each code unit little-endian). */
function utf16le(s: string): Bytes {
  const out = new Uint8Array(s.length * 2);
  const view = new DataView(out.buffer);
  for (let i = 0; i < s.length; i++) view.setUint16(i * 2, s.charCodeAt(i), true);
  return out;
}

async function digest(hash: string, data: Bytes): Promise<Bytes> {
  return new Uint8Array(await subtle().digest(hash, data));
}

/** Truncate or right-pad (0x36) `data` to exactly `len` bytes (§2.3.4.11 /
 *  §2.3.4.12 fitting rule). */
function fitTo(data: Uint8Array, len: number): Bytes {
  if (data.length > len) return data.slice(0, len);
  const out = new Uint8Array(len);
  out.set(data);
  if (data.length < len) out.fill(PAD_BYTE, data.length);
  return out;
}

/**
 * §2.3.4.11: derive an encryption key from the password, salt, spinCount and a
 * blockKey. Returns exactly `keyBits/8` bytes.
 *
 * The spin loop is `spinCount` awaited `subtle.digest` calls; at the default
 * 100,000 iterations this costs ~1 s on a modern machine, which is acceptable
 * to run on the main thread (WebCrypto is native, non-blocking). No custom
 * synchronous hash is needed at this cost.
 */
export async function deriveAgileKey(
  password: string,
  params: AgileCipherParams,
  spinCount: number,
  blockKey: Uint8Array,
): Promise<Bytes> {
  const hash = webCryptoHash(params.hashAlgorithm);
  let h = await digest(hash, concat(params.saltValue, utf16le(password)));
  for (let i = 0; i < spinCount; i++) {
    h = await digest(hash, concat(le32(i), h));
  }
  const hFinal = await digest(hash, concat(h, blockKey));
  return fitTo(hFinal, params.keyBits / 8);
}

/** §2.3.4.12: IV from a salt and optional blockKey, fitted to blockSize. */
export async function deriveIv(
  params: AgileCipherParams,
  salt: Uint8Array,
  blockKey: Uint8Array | null,
): Promise<Bytes> {
  const iv = blockKey ? await digest(webCryptoHash(params.hashAlgorithm), concat(salt, blockKey)) : salt;
  return fitTo(iv, params.blockSize);
}

/**
 * AES-CBC decrypt WITHOUT PKCS#7 padding.
 *
 * WebCrypto's `AES-CBC` mode always validates and strips a PKCS#7 pad on
 * decrypt, but Agile ciphertext is zero-/no-padded (the block count already
 * covers the data). The standard portable workaround: append ONE synthetic
 * trailing ciphertext block that decrypts to a full, valid PKCS#7 pad, run the
 * native decrypt over `ciphertext || trailing`, and discard the extra block.
 *
 * The trailing block is built so its plaintext is `0x10 * 16` (a whole padding
 * block): `trailing = AES-ECB-encrypt(0x10*16 XOR lastCipherBlock)`. Under CBC
 * chaining WebCrypto computes `P_trailing = ECB_dec(trailing) XOR lastCipher =
 * 0x10*16`, so padding validation passes and the preceding N blocks decrypt
 * exactly as raw CBC. AES-ECB-encrypt of one block is obtained via a one-block
 * AES-CBC encrypt with a zero IV (CBC with IV=0 over a single block == ECB).
 */
async function aesCbcDecryptNoPad(
  key: Bytes,
  iv: Bytes,
  ciphertext: Uint8Array,
): Promise<Bytes> {
  const blockSize = iv.length;
  if (ciphertext.length === 0) return new Uint8Array(0);
  if (ciphertext.length % blockSize !== 0) {
    throw new AgileDecryptError('corrupt', 'ciphertext length is not a multiple of the block size');
  }

  const decKey = await subtle().importKey('raw', key, { name: 'AES-CBC' }, false, ['decrypt']);
  const encKey = await subtle().importKey('raw', key, { name: 'AES-CBC' }, false, ['encrypt']);

  const lastBlock = ciphertext.subarray(ciphertext.length - blockSize);
  const fullPad = new Uint8Array(blockSize).fill(blockSize); // 0x10 repeated (PKCS#7 whole block)
  // trailing = ECB_enc(fullPad XOR lastBlock). ECB_enc(x) == first block of
  // AES-CBC-encrypt(x, IV=0). WebCrypto appends its own pad block to the
  // single input block, so we take only the first output block.
  const ecbInput = xorBytes(fullPad, lastBlock);
  const cbcEnc = new Uint8Array(
    await subtle().encrypt({ name: 'AES-CBC', iv: new Uint8Array(blockSize) }, encKey, ecbInput),
  );
  const trailing = cbcEnc.subarray(0, blockSize);

  const padded = concat(ciphertext, trailing);
  const decrypted = new Uint8Array(await subtle().decrypt({ name: 'AES-CBC', iv }, decKey, padded));
  // `decrypted` is the N original plaintext blocks (the trailing pad block was
  // stripped by WebCrypto's PKCS#7 removal).
  return decrypted.length >= ciphertext.length ? decrypted.subarray(0, ciphertext.length) : decrypted;
}

function xorBytes(a: Uint8Array, b: Uint8Array): Bytes {
  const out = new Uint8Array(a.length);
  for (let i = 0; i < a.length; i++) out[i] = a[i] ^ b[i];
  return out;
}

/**
 * Verify the password against the PasswordKeyEncryptor verifier (§2.3.4.13):
 * decrypt encryptedVerifierHashInput and encryptedVerifierHashValue with their
 * password-derived keys (IV = saltValue), then check
 * H(decryptedInput) == decryptedValue (first hashSize bytes).
 */
export async function verifyPassword(
  password: string,
  pke: PasswordKeyEncryptor,
): Promise<boolean> {
  assertSupportedCipher(pke);
  const hash = webCryptoHash(pke.hashAlgorithm);

  const keyIn = await deriveAgileKey(password, pke, pke.spinCount, BLOCK_KEY.verifierHashInput);
  const ivIn = await deriveIv(pke, pke.saltValue, null);
  const verifierInput = await aesCbcDecryptNoPad(keyIn, ivIn, pke.encryptedVerifierHashInput);

  const keyVal = await deriveAgileKey(password, pke, pke.spinCount, BLOCK_KEY.verifierHashValue);
  const ivVal = await deriveIv(pke, pke.saltValue, null);
  const verifierValue = await aesCbcDecryptNoPad(keyVal, ivVal, pke.encryptedVerifierHashValue);

  const expected = await digest(hash, verifierInput);
  return timingSafeEqual(expected.subarray(0, pke.hashSize), verifierValue.subarray(0, pke.hashSize));
}

function timingSafeEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a[i] ^ b[i];
  return diff === 0;
}

/**
 * Decrypt the intermediate key: decrypt encryptedKeyValue with the
 * password-derived key (blockKey = keyValue, IV = pke.saltValue). §2.3.4.13.
 */
export async function deriveIntermediateKey(
  password: string,
  pke: PasswordKeyEncryptor,
): Promise<Bytes> {
  const key = await deriveAgileKey(password, pke, pke.spinCount, BLOCK_KEY.keyValue);
  const iv = await deriveIv(pke, pke.saltValue, null);
  const intermediate = await aesCbcDecryptNoPad(key, iv, pke.encryptedKeyValue);
  // Intermediate key is keyData.keyBits/8 long; encryptedKeyValue may be padded.
  return intermediate;
}

/**
 * Decrypt the EncryptedPackage stream (§2.3.4.15). The stream is
 * LE64(plaintextSize) followed by 4096-byte AES-CBC segments; each segment's IV
 * is H(keyData.saltValue + LE32(segmentIndex)) fitted to keyData.blockSize.
 * Returns exactly the declared plaintext bytes.
 */
export async function decryptPackage(
  encryptedPackage: Uint8Array,
  keyData: AgileCipherParams,
  intermediateKey: Bytes,
): Promise<Bytes> {
  assertSupportedCipher(keyData);
  if (encryptedPackage.length < 8) {
    throw new AgileDecryptError('corrupt', 'EncryptedPackage is shorter than its size prefix');
  }
  const view = new DataView(
    encryptedPackage.buffer,
    encryptedPackage.byteOffset,
    encryptedPackage.byteLength,
  );
  // getBigUint64 is unsigned, so Number(...) here is never negative — only
  // the upper bound needs checking.
  const plaintextSize = Number(view.getBigUint64(0, true));
  const ciphertext = encryptedPackage.subarray(8);
  if (plaintextSize > ciphertext.length) {
    throw new AgileDecryptError('corrupt', 'EncryptedPackage size prefix exceeds the ciphertext');
  }

  const key = intermediateKey.slice(0, keyData.keyBits / 8);
  const out = new Uint8Array(plaintextSize);
  let written = 0;
  let segIndex = 0;
  for (let off = 0; off < ciphertext.length; off += SEGMENT_LENGTH) {
    const segment = ciphertext.subarray(off, off + SEGMENT_LENGTH);
    const iv = await deriveIv(keyData, keyData.saltValue, le32(segIndex));
    const plain = await aesCbcDecryptNoPad(key, iv, segment);
    const take = Math.min(plain.length, plaintextSize - written);
    out.set(plain.subarray(0, take), written);
    written += take;
    segIndex++;
    if (written >= plaintextSize) break;
  }
  if (written !== plaintextSize) {
    throw new AgileDecryptError('corrupt', 'decrypted output is shorter than the declared size');
  }
  return out;
}

/**
 * Full Agile decryption: verify the password, derive the intermediate key, and
 * decrypt the package. Throws {@link AgileDecryptError} with reason
 * `'invalid-password'` if the password is wrong.
 */
export async function decryptAgilePackage(
  descriptor: AgileEncryptionDescriptor,
  encryptedPackage: Uint8Array,
  password: string,
): Promise<Bytes> {
  const { keyData, passwordKeyEncryptor } = descriptor;
  assertSupportedCipher(keyData);
  assertSupportedCipher(passwordKeyEncryptor);

  const ok = await verifyPassword(password, passwordKeyEncryptor);
  if (!ok) {
    throw new AgileDecryptError('invalid-password', 'The supplied password is incorrect.');
  }
  const intermediateKey = await deriveIntermediateKey(password, passwordKeyEncryptor);
  return decryptPackage(encryptedPackage, keyData, intermediateKey);
}
