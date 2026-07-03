import { describe, it, expect } from 'vitest';
import { decryptOoxml } from './decrypt-ooxml';
import { resolveOoxmlContainer } from '../errors/cfb-guard';
import { OoxmlError } from '../errors/ooxml-error';
import { buildCfbWithStreams } from '../testing/cfb-fixture';

/**
 * A version 4.4 (Agile) `EncryptionInfo` header whose XML descriptor is
 * missing a required attribute exercises `parseAgileDescriptor`'s `null`
 * return path (packages/core/src/crypto/encryption-info.ts) — a structurally
 * corrupt Agile descriptor, as opposed to a well-formed non-Agile one (covered
 * by decrypt-ooxml-unsupported.test.ts). `parseEncryptionInfo` maps that
 * `null` to `{ kind: 'unknown' }`, and `decryptOoxml` treats any non-`'agile'`
 * kind — including `'unknown'` — as `'unsupported-encryption'`, the same
 * outcome a genuinely unsupported scheme produces. There is no separate
 * "malformed Agile" error code; a broken descriptor is indistinguishable from
 * an unsupported one to the caller, which is the behaviour under test here.
 */
function agileHeader(): Uint8Array {
  const bytes = new Uint8Array(8);
  const view = new DataView(bytes.buffer);
  view.setUint16(0, 4, true); // vMajor = 4
  view.setUint16(2, 4, true); // vMinor = 4 -> Agile
  view.setUint32(4, 0x00000040, true); // Reserved, per §2.3.4.10
  return bytes;
}

function utf8(s: string): Uint8Array {
  return new TextEncoder().encode(s);
}

/** A `<keyData>` with every required attribute except `saltValue` — the same
 *  omission `readCipherParams` null-checks for. */
const KEY_DATA_MISSING_SALT_VALUE =
  '<keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" ' +
  'cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512"/>';

const ENCRYPTED_KEY =
  '<p:encryptedKey spinCount="0" saltSize="16" blockSize="16" keyBits="256" hashSize="64" ' +
  'cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" ' +
  'saltValue="EBESExQVFhcYGRobHB0eHw==" ' +
  'encryptedVerifierHashInput="LmG8hnMal86Av+TI2BymxA==" ' +
  'encryptedVerifierHashValue="IOglNh3gjwzZOshzI6D1jOrrIPP2XassGj5DaRYsYhtUxemGNg0oQBHh6MBUIiRY30R+zdgdraBBm7JPi36XQw==" ' +
  'encryptedKeyValue="vraviKwQ4XwUZ9PGRI+40bO1lusZvmYIoU4DiTZqYuY="/>';

function brokenAgileDescriptorXml(): Uint8Array {
  const xml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    '<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" ' +
    'xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">' +
    KEY_DATA_MISSING_SALT_VALUE +
    '<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">' +
    ENCRYPTED_KEY +
    '</keyEncryptor></keyEncryptors></encryption>';
  return utf8(xml);
}

function encryptionInfoStream(): Uint8Array {
  const header = agileHeader();
  const xml = brokenAgileDescriptorXml();
  const out = new Uint8Array(header.length + xml.length);
  out.set(header, 0);
  out.set(xml, header.length);
  return out;
}

function dummyEncryptedPackage(): Uint8Array {
  return new Uint8Array(8);
}

describe('decryptOoxml / resolveOoxmlContainer — broken Agile XML descriptor', () => {
  it('a 4.4 header with keyData missing saltValue is treated as unsupported-encryption', async () => {
    const cfb = new Uint8Array(
      buildCfbWithStreams([
        { name: 'EncryptionInfo', data: encryptionInfoStream() },
        { name: 'EncryptedPackage', data: dummyEncryptedPackage() },
      ]),
    );

    const result = await decryptOoxml(cfb, 'test');
    expect(result).toEqual({ ok: false, reason: 'unsupported-encryption' });
  });

  it('resolveOoxmlContainer surfaces the same broken descriptor as OoxmlError("unsupported-encryption")', async () => {
    const cfb = new Uint8Array(
      buildCfbWithStreams([
        { name: 'EncryptionInfo', data: encryptionInfoStream() },
        { name: 'EncryptedPackage', data: dummyEncryptedPackage() },
      ]),
    );

    await expect(resolveOoxmlContainer(cfb, 'test')).rejects.toBeInstanceOf(OoxmlError);
    await expect(resolveOoxmlContainer(cfb, 'test')).rejects.toMatchObject({
      code: 'unsupported-encryption',
    });
  });
});
