import { describe, it, expect } from 'vitest';
import { OoxmlError, type OoxmlErrorCode } from './ooxml-error';

describe('OoxmlError', () => {
  it('is an Error subclass carrying a machine-readable code', () => {
    const err = new OoxmlError('encrypted', 'This file is password-protected.');
    expect(err).toBeInstanceOf(Error);
    expect(err).toBeInstanceOf(OoxmlError);
    expect(err.code).toBe('encrypted');
    expect(err.message).toBe('This file is password-protected.');
  });

  it('sets name to "OoxmlError" so stringification is stable', () => {
    const err = new OoxmlError('not-ooxml', 'nope');
    expect(err.name).toBe('OoxmlError');
    // Error.prototype.toString uses `name: message`.
    expect(String(err)).toBe('OoxmlError: nope');
  });

  it('keeps the code readonly at the type level and reflects each variant', () => {
    const codes: OoxmlErrorCode[] = ['encrypted', 'legacy-binary-format', 'not-ooxml'];
    for (const code of codes) {
      expect(new OoxmlError(code, code).code).toBe(code);
    }
  });

  it('captures a stack trace', () => {
    const err = new OoxmlError('legacy-binary-format', 'legacy');
    expect(typeof err.stack).toBe('string');
  });
});
