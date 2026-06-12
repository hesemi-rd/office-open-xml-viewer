import { describe, expect, it, vi } from 'vitest';

// `vscode` only exists inside the extension host. `buildContentSecurityPolicy`
// does not touch it, but importing the module pulls in the top-level
// `import * as vscode` — stub it so the unit under test loads in plain Node.
vi.mock('vscode', () => ({}));

import { buildContentSecurityPolicy } from './webviewHtml';

const CSP_SOURCE = 'vscode-webview://abc';
const NONCE = 'testnonce';

describe('buildContentSecurityPolicy', () => {
  it('keeps the webview fully offline when Google Fonts are disabled', () => {
    const csp = buildContentSecurityPolicy(CSP_SOURCE, NONCE, false);

    // No outbound origin other than the extension itself.
    expect(csp).not.toContain('fonts.googleapis.com');
    expect(csp).not.toContain('fonts.gstatic.com');
    expect(csp).not.toContain('https://');

    // Baseline directives are still present.
    expect(csp).toContain("default-src 'none'");
    expect(csp).toContain(`script-src 'nonce-${NONCE}' 'wasm-unsafe-eval'`);
    expect(csp).toContain("style-src 'unsafe-inline'");
    expect(csp).toContain(`font-src ${CSP_SOURCE}`);
  });

  it('whitelists only the two font CDN origins when enabled', () => {
    const csp = buildContentSecurityPolicy(CSP_SOURCE, NONCE, true);

    // The stylesheet itself is served from fonts.googleapis.com → style-src.
    expect(csp).toMatch(/style-src[^;]*https:\/\/fonts\.googleapis\.com/);
    // The woff2 files referenced by the stylesheet come from fonts.gstatic.com → font-src.
    expect(csp).toMatch(/font-src[^;]*https:\/\/fonts\.gstatic\.com/);

    // googleapis is for the stylesheet, NOT the font binaries.
    const fontSrc = csp.match(/font-src ([^;]*)/)?.[1] ?? '';
    expect(fontSrc).toContain('https://fonts.gstatic.com');
    expect(fontSrc).not.toContain('googleapis');

    // The library never fetch()es the CDN, so connect-src must stay un-widened.
    const connectSrc = csp.match(/connect-src ([^;]*)/)?.[1] ?? '';
    expect(connectSrc).not.toContain('googleapis');
    expect(connectSrc).not.toContain('gstatic');
  });
});
