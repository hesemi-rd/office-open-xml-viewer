/**
 * Parse the `\EncryptionInfo` stream of an Agile-encrypted OOXML package
 * ([MS-OFFCRYPTO] §2.3.4.10).
 *
 * The stream is an 8-byte binary header followed by a UTF-8 XML descriptor:
 *
 *   - EncryptionVersionInfo (§2.3.4.10): `vMajor` (LE16) and `vMinor` (LE16).
 *     Agile Encryption is exactly `vMajor=4, vMinor=4`. `vMajor` 3/2 is
 *     Standard/legacy binary encryption (different, binary descriptor) and
 *     `vMinor=0x0010` with `vMajor=4` is Extensible — both unsupported here.
 *   - Reserved (4 bytes): MUST be 0x00000040.
 *   - XmlEncryptionDescriptor: `<encryption>` with a `<keyData>`, an optional
 *     `<dataIntegrity>`, and a `<keyEncryptors>` holding a password
 *     `<p:encryptedKey>` (§2.3.4.10 PasswordKeyEncryptor schema).
 *
 * The descriptor is a tiny, flat, fixed schema, so rather than pull in a full
 * XML parser (core has zero runtime deps and must run in a worker without
 * DOMParser) we extract the attributes of the three known elements directly.
 * This is not a general XML reader — it targets exactly this schema.
 */

/**
 * Byte buffer backed by a plain `ArrayBuffer` (not `SharedArrayBuffer`). The
 * WebCrypto `SubtleCrypto` methods require a `BufferSource` over an
 * `ArrayBuffer`, and TypeScript's strict `Uint8Array<ArrayBufferLike>` default
 * would otherwise not satisfy that. All byte values in this module come from
 * fresh `new Uint8Array(n)` allocations (via base64 decoding), which are
 * `Uint8Array<ArrayBuffer>`, so this alias is accurate.
 */
export type Bytes = Uint8Array<ArrayBuffer>;

/** Discriminates the version of the EncryptionInfo stream. */
export type EncryptionInfoKind =
  | { kind: 'agile'; descriptor: AgileEncryptionDescriptor }
  | { kind: 'standard' } // §2.3.4.5 (vMajor 2/3/4 + vMinor 2) — binary descriptor
  | { kind: 'extensible' } // vMajor 3/4 + vMinor 0x0010
  | { kind: 'unknown' };

/** The cipher / hash parameters shared by KeyData and the PasswordKeyEncryptor
 *  (a subset of [MS-OFFCRYPTO] §2.3.4.10 attributes). */
export interface AgileCipherParams {
  saltSize: number;
  blockSize: number;
  keyBits: number;
  hashSize: number;
  cipherAlgorithm: string; // "AES"
  cipherChaining: string; // "ChainingModeCBC" | "ChainingModeCFB"
  hashAlgorithm: string; // "SHA512" | "SHA384" | "SHA256" | "SHA1" | ...
  /** Decoded `saltValue` (base64 in the XML). */
  saltValue: Bytes;
}

/** The `<p:encryptedKey>` password key encryptor (§2.3.4.10). */
export interface PasswordKeyEncryptor extends AgileCipherParams {
  spinCount: number;
  encryptedVerifierHashInput: Bytes;
  encryptedVerifierHashValue: Bytes;
  encryptedKeyValue: Bytes;
}

/** Optional `<dataIntegrity>` (§2.3.4.10 / verification §2.3.4.14). */
export interface DataIntegrity {
  encryptedHmacKey: Bytes;
  encryptedHmacValue: Bytes;
}

/** The parsed Agile descriptor. */
export interface AgileEncryptionDescriptor {
  keyData: AgileCipherParams;
  passwordKeyEncryptor: PasswordKeyEncryptor;
  dataIntegrity: DataIntegrity | null;
}

/** Decode standard base64 to bytes (no atob dependency — works in Node too). */
function base64ToBytes(b64: string): Bytes {
  // atob exists in browsers and in modern Node globals; Buffer is the Node
  // fallback. Guard both so this runs anywhere.
  if (typeof atob === 'function') {
    const bin = atob(b64);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const B = (globalThis as any).Buffer;
  if (B) return new Uint8Array(B.from(b64, 'base64'));
  throw new Error('no base64 decoder available');
}

/** Read attribute `attr` off the first `<tag ...>` occurrence in `xml`. */
function attr(xml: string, tag: string, name: string): string | null {
  // Match `<tag ... name="value" ...>` allowing any namespace prefix on the
  // tag (e.g. `p:encryptedKey`). The descriptor is machine-generated, so a
  // straightforward attribute-quote scan is sufficient and safe.
  const tagRe = new RegExp(`<(?:[\\w]+:)?${tag}\\b[^>]*>`);
  const m = tagRe.exec(xml);
  if (!m) return null;
  const el = m[0];
  const attrRe = new RegExp(`\\b${name}\\s*=\\s*"([^"]*)"`);
  const am = attrRe.exec(el);
  return am ? am[1] : null;
}

function num(v: string | null): number | null {
  if (v === null) return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function readCipherParams(xml: string, tag: string): AgileCipherParams | null {
  const saltSize = num(attr(xml, tag, 'saltSize'));
  const blockSize = num(attr(xml, tag, 'blockSize'));
  const keyBits = num(attr(xml, tag, 'keyBits'));
  const hashSize = num(attr(xml, tag, 'hashSize'));
  const cipherAlgorithm = attr(xml, tag, 'cipherAlgorithm');
  const cipherChaining = attr(xml, tag, 'cipherChaining');
  const hashAlgorithm = attr(xml, tag, 'hashAlgorithm');
  const saltValueB64 = attr(xml, tag, 'saltValue');
  if (
    saltSize === null ||
    blockSize === null ||
    keyBits === null ||
    hashSize === null ||
    !cipherAlgorithm ||
    !cipherChaining ||
    !hashAlgorithm ||
    saltValueB64 === null
  ) {
    return null;
  }
  return {
    saltSize,
    blockSize,
    keyBits,
    hashSize,
    cipherAlgorithm,
    cipherChaining,
    hashAlgorithm,
    saltValue: base64ToBytes(saltValueB64),
  };
}

/**
 * Parse the EncryptionInfo stream. Returns a discriminated result: `'agile'`
 * with the descriptor when it is Agile Encryption (v4.4), or one of the
 * unsupported markers otherwise. Never throws for a malformed descriptor — a
 * v4.4 header with an unparseable body yields `'unknown'`.
 */
export function parseEncryptionInfo(bytes: Uint8Array): EncryptionInfoKind {
  if (bytes.length < 8) return { kind: 'unknown' };
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const vMajor = view.getUint16(0, true);
  const vMinor = view.getUint16(2, true);

  // §2.3.4.10: Agile is exactly 4.4.
  if (vMajor === 4 && vMinor === 4) {
    const descriptor = parseAgileDescriptor(bytes.subarray(8));
    return descriptor ? { kind: 'agile', descriptor } : { kind: 'unknown' };
  }
  // vMinor 0x0010 (16) with vMajor 3 or 4 is Extensible Encryption (§2.3.4.2).
  if (vMinor === 0x0010 && (vMajor === 3 || vMajor === 4)) return { kind: 'extensible' };
  // vMajor 2/3/4 with vMinor 2 is Standard Encryption (§2.3.4.5), a binary
  // (non-XML) descriptor.
  if (vMinor === 2 && (vMajor === 2 || vMajor === 3 || vMajor === 4)) return { kind: 'standard' };
  return { kind: 'unknown' };
}

function parseAgileDescriptor(xmlBytes: Uint8Array): AgileEncryptionDescriptor | null {
  const xml = new TextDecoder('utf-8').decode(xmlBytes);

  const keyData = readCipherParams(xml, 'keyData');
  const pkeBase = readCipherParams(xml, 'encryptedKey');
  if (!keyData || !pkeBase) return null;

  const spinCount = num(attr(xml, 'encryptedKey', 'spinCount'));
  const eviB64 = attr(xml, 'encryptedKey', 'encryptedVerifierHashInput');
  const evvB64 = attr(xml, 'encryptedKey', 'encryptedVerifierHashValue');
  const ekvB64 = attr(xml, 'encryptedKey', 'encryptedKeyValue');
  if (spinCount === null || eviB64 === null || evvB64 === null || ekvB64 === null) return null;

  const passwordKeyEncryptor: PasswordKeyEncryptor = {
    ...pkeBase,
    spinCount,
    encryptedVerifierHashInput: base64ToBytes(eviB64),
    encryptedVerifierHashValue: base64ToBytes(evvB64),
    encryptedKeyValue: base64ToBytes(ekvB64),
  };

  let dataIntegrity: DataIntegrity | null = null;
  const hmacKeyB64 = attr(xml, 'dataIntegrity', 'encryptedHmacKey');
  const hmacValB64 = attr(xml, 'dataIntegrity', 'encryptedHmacValue');
  if (hmacKeyB64 !== null && hmacValB64 !== null) {
    dataIntegrity = {
      encryptedHmacKey: base64ToBytes(hmacKeyB64),
      encryptedHmacValue: base64ToBytes(hmacValB64),
    };
  }

  return { keyData, passwordKeyEncryptor, dataIntegrity };
}
