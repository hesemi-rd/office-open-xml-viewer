import { describe, it, expect } from 'vitest';
import {
  deriveAgileKey,
  deriveIv,
  verifyPassword,
  deriveIntermediateKey,
  decryptPackage,
  decryptAgilePackage,
  AgileDecryptError,
  BLOCK_KEY,
} from './agile';
import { parseEncryptionInfo, type AgileCipherParams, type PasswordKeyEncryptor } from './encryption-info';
import { readCfbStream } from '../errors/cfb-read';
import { encryptedDocxSpin0 } from '../testing/encrypted-fixture';

/**
 * Pinned test vectors, computed by the independent Python encryptor that
 * generated the embedded fixture (and cross-checked by msoffcrypto-tool). They
 * fix the exact key-derivation math of [MS-OFFCRYPTO] §2.3.4.11 / §2.3.4.12:
 *   password  = 'test'
 *   pwSalt    = 10 11 .. 1f    (PasswordKeyEncryptor.saltValue)
 *   keySalt   = 00 01 .. 0f    (keyData.saltValue)
 * hashAlgorithm SHA512, keyBits 256, blockSize 16.
 */
const PW = 'test';
const PW_SALT = Uint8Array.from({ length: 16 }, (_, i) => i + 16); // 10..1f
const KEY_SALT = Uint8Array.from({ length: 16 }, (_, i) => i); // 00..0f

function hex(bytes: Uint8Array): string {
  return Array.from(bytes, (b) => b.toString(16).padStart(2, '0')).join('');
}

/** Minimal SHA512/AES-CBC/256 params sharing the fixed salts above. */
function params(salt: Uint8Array): AgileCipherParams {
  return {
    saltSize: 16,
    blockSize: 16,
    keyBits: 256,
    hashSize: 64,
    cipherAlgorithm: 'AES',
    cipherChaining: 'ChainingModeCBC',
    hashAlgorithm: 'SHA512',
    saltValue: salt,
  };
}

describe('deriveAgileKey — §2.3.4.11 spin loop (pinned vectors)', () => {
  it('derives the keyValue key at spinCount 0', async () => {
    const key = await deriveAgileKey(PW, params(PW_SALT), 0, BLOCK_KEY.keyValue);
    expect(hex(key)).toBe('080bc0077d116dfd2d8daa05825c08ba6d7a90004ddbf6d1b861022fa9ccd47d');
  });

  it('derives the verifierHashInput key at spinCount 0', async () => {
    const key = await deriveAgileKey(PW, params(PW_SALT), 0, BLOCK_KEY.verifierHashInput);
    expect(hex(key)).toBe('1ec8f93d236bc7b6af8cdf4a810ed54cbabbf63746a0c245610d58f931cd548c');
  });

  it('derives the keyValue key at spinCount 3 (exercises the iterator loop)', async () => {
    const key = await deriveAgileKey(PW, params(PW_SALT), 3, BLOCK_KEY.keyValue);
    expect(hex(key)).toBe('cdfed59d5986a51270695235bc3df0ab81bb67400fa5174fdf6084d5c801f3c1');
  });

  it('derives the keyValue key at the real-world spinCount 100000', async () => {
    const key = await deriveAgileKey(PW, params(PW_SALT), 100000, BLOCK_KEY.keyValue);
    expect(hex(key)).toBe('e68661b7e2a2d211e6a0ecad75de1ac49fc02c7db6ddeb882d5379e32fcee03c');
  });
});

describe('deriveIv — §2.3.4.12', () => {
  it('IV = saltValue (no blockKey), fitted to blockSize', async () => {
    const iv = await deriveIv(params(PW_SALT), PW_SALT, null);
    expect(hex(iv)).toBe(hex(PW_SALT)); // 16-byte salt == blockSize, unchanged
  });

  it('IV = H(keySalt + LE32(0)) truncated to blockSize (data segment 0)', async () => {
    const iv = await deriveIv(params(KEY_SALT), KEY_SALT, new Uint8Array([0, 0, 0, 0]));
    expect(hex(iv)).toBe('4907038c06363497bcd63dc397144376');
  });

  it('IV = H(keySalt + LE32(1)) for data segment 1', async () => {
    const iv = await deriveIv(params(KEY_SALT), KEY_SALT, new Uint8Array([1, 0, 0, 0]));
    expect(hex(iv)).toBe('973b6d38872770a510139d8eeab971aa');
  });
});

describe('parseEncryptionInfo — §2.3.4.10', () => {
  it('parses the Agile descriptor from the fixture EncryptionInfo stream', () => {
    const info = readCfbStream(encryptedDocxSpin0(), 'EncryptionInfo');
    expect(info).not.toBeNull();
    const parsed = parseEncryptionInfo(info as Uint8Array);
    expect(parsed.kind).toBe('agile');
    if (parsed.kind !== 'agile') return;
    const { keyData, passwordKeyEncryptor, dataIntegrity } = parsed.descriptor;
    expect(keyData.hashAlgorithm).toBe('SHA512');
    expect(keyData.keyBits).toBe(256);
    expect(keyData.blockSize).toBe(16);
    expect(hex(keyData.saltValue)).toBe(hex(KEY_SALT));
    expect(passwordKeyEncryptor.spinCount).toBe(0);
    expect(hex(passwordKeyEncryptor.saltValue)).toBe(hex(PW_SALT));
    expect(passwordKeyEncryptor.encryptedKeyValue.length).toBeGreaterThan(0);
    expect(dataIntegrity).not.toBeNull();
  });

  it('classifies a Standard-encryption header (v3.2) as unsupported', () => {
    const header = new Uint8Array(8);
    new DataView(header.buffer).setUint16(0, 3, true); // vMajor 3
    new DataView(header.buffer).setUint16(2, 2, true); // vMinor 2
    expect(parseEncryptionInfo(header).kind).toBe('standard');
  });

  it('classifies an Extensible header (v4.16) as unsupported', () => {
    const header = new Uint8Array(8);
    new DataView(header.buffer).setUint16(0, 4, true);
    new DataView(header.buffer).setUint16(2, 0x0010, true);
    expect(parseEncryptionInfo(header).kind).toBe('extensible');
  });
});

describe('verifyPassword / decrypt — full pipeline against the fixture', () => {
  function fixtureDescriptor() {
    const info = readCfbStream(encryptedDocxSpin0(), 'EncryptionInfo');
    const parsed = parseEncryptionInfo(info as Uint8Array);
    if (parsed.kind !== 'agile') throw new Error('expected agile');
    return parsed.descriptor;
  }

  it('accepts the correct password', async () => {
    const pke = fixtureDescriptor().passwordKeyEncryptor as PasswordKeyEncryptor;
    expect(await verifyPassword('test', pke)).toBe(true);
  });

  it('rejects a wrong password', async () => {
    const pke = fixtureDescriptor().passwordKeyEncryptor as PasswordKeyEncryptor;
    expect(await verifyPassword('wrong', pke)).toBe(false);
  });

  it('decrypts the package to the original plaintext ZIP (PK magic + zip entries)', async () => {
    const desc = fixtureDescriptor();
    const pkg = readCfbStream(encryptedDocxSpin0(), 'EncryptedPackage') as Uint8Array;
    const plain = await decryptAgilePackage(desc, pkg, 'test');
    expect(plain[0]).toBe(0x50); // 'P'
    expect(plain[1]).toBe(0x4b); // 'K'
    // The entry names sit in the ZIP local/central headers uncompressed, so
    // they appear verbatim even though document.xml's content is deflated.
    const text = new TextDecoder('latin1').decode(plain);
    expect(text).toContain('[Content_Types].xml');
    expect(text).toContain('word/document.xml');
    expect(text).toContain('_rels/.rels');
  });

  it('throws AgileDecryptError(invalid-password) for the wrong password', async () => {
    const desc = fixtureDescriptor();
    const pkg = readCfbStream(encryptedDocxSpin0(), 'EncryptedPackage') as Uint8Array;
    await expect(decryptAgilePackage(desc, pkg, 'nope')).rejects.toMatchObject({
      name: 'AgileDecryptError',
      reason: 'invalid-password',
    });
  });

  it('deriveIntermediateKey + decryptPackage compose to the same plaintext', async () => {
    const desc = fixtureDescriptor();
    const pkg = readCfbStream(encryptedDocxSpin0(), 'EncryptedPackage') as Uint8Array;
    const ik = await deriveIntermediateKey('test', desc.passwordKeyEncryptor);
    const plain = await decryptPackage(pkg, desc.keyData, ik);
    expect(plain[0]).toBe(0x50);
    expect(plain[1]).toBe(0x4b);
  });
});

describe('AgileDecryptError — unsupported cipher / hash', () => {
  it('rejects a non-AES cipher as unsupported-encryption', async () => {
    const p: PasswordKeyEncryptor = {
      ...params(PW_SALT),
      cipherAlgorithm: 'RC2',
      spinCount: 0,
      encryptedVerifierHashInput: new Uint8Array(16),
      encryptedVerifierHashValue: new Uint8Array(64),
      encryptedKeyValue: new Uint8Array(32),
    };
    await expect(verifyPassword('x', p)).rejects.toBeInstanceOf(AgileDecryptError);
  });
});
