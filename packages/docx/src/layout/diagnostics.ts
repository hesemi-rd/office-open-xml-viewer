import type { LayoutDiagnosticCode } from './types.js';

export class LayoutInvariantError extends Error {
  readonly code: LayoutDiagnosticCode;

  constructor(code: LayoutDiagnosticCode, detail: string) {
    super(`${code}: ${detail}`);
    this.name = 'LayoutInvariantError';
    this.code = code;
  }
}
