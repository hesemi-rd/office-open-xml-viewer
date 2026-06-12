import * as vscode from 'vscode';

/** Origin that serves the Google Fonts *stylesheets* (`<link rel="stylesheet">`). */
const GOOGLE_FONTS_CSS_ORIGIN = 'https://fonts.googleapis.com';
/** Origin that serves the actual `.woff2` font binaries referenced by the CSS. */
const GOOGLE_FONTS_FILES_ORIGIN = 'https://fonts.gstatic.com';

/**
 * Build the webview Content-Security-Policy string.
 *
 * Pure function (no VSCode API) so it can be unit-tested for both states.
 *
 * When `useGoogleFonts` is false the policy is fully offline: the only allowed
 * origin is the extension's own `cspSource`. When true we widen exactly two
 * directives, no more:
 *   - `style-src` gains {@link GOOGLE_FONTS_CSS_ORIGIN} because the library
 *     loads each Google Fonts CSS via an injected `<link rel="stylesheet">`.
 *   - `font-src` gains {@link GOOGLE_FONTS_FILES_ORIGIN} because the `@font-face`
 *     rules in that CSS point their `src:` at `fonts.gstatic.com` woff2 files.
 * `connect-src` is deliberately NOT widened: the preload path
 * (`packages/core/src/fonts/preload.ts`) never `fetch()`es either origin — the
 * browser font engine fetches the binaries, governed by `font-src`.
 */
export function buildContentSecurityPolicy(
  cspSource: string,
  nonce: string,
  useGoogleFonts: boolean,
): string {
  const fontSrc = useGoogleFonts
    ? `font-src ${cspSource} ${GOOGLE_FONTS_FILES_ORIGIN};`
    : `font-src ${cspSource};`;
  const styleSrc = useGoogleFonts
    ? `style-src 'unsafe-inline' ${GOOGLE_FONTS_CSS_ORIGIN};`
    : `style-src 'unsafe-inline';`;

  return [
    `default-src 'none';`,
    `img-src ${cspSource} data: blob:;`,
    `media-src ${cspSource} blob:;`,
    fontSrc,
    `script-src 'nonce-${nonce}' 'wasm-unsafe-eval';`,
    `worker-src data: blob:;`,
    styleSrc,
    `connect-src ${cspSource} data: blob:;`,
  ].join(' ');
}

/**
 * Generate the HTML for the webview panel.
 * The webview script (dist/webview.js) is allowed via the content security policy,
 * and receives the file bytes via a `ooxml-init` message posted from the extension host.
 *
 * When `useGoogleFonts` is true the CSP is widened to allow the metric-compatible
 * font CDN (see {@link buildContentSecurityPolicy}); the flag is also forwarded to
 * the viewers via the `ooxml-init` message in the editor providers.
 */
export function getWebviewHtml(
  webview: vscode.Webview,
  extensionUri: vscode.Uri,
  fileType: 'docx' | 'xlsx' | 'pptx',
  useGoogleFonts = false,
): string {
  const scriptUri = webview.asWebviewUri(
    vscode.Uri.joinPath(extensionUri, 'dist', 'webview.js'),
  );

  const nonce = getNonce();
  const csp = buildContentSecurityPolicy(webview.cspSource, nonce, useGoogleFonts);

  return /* html */ `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <meta http-equiv="Content-Security-Policy" content="${csp}" />
  <title>OOXML Viewer</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    html, body {
      width: 100%;
      height: 100%;
      background: var(--vscode-editor-background);
      color: var(--vscode-foreground);
      font-family: var(--vscode-font-family, sans-serif);
    }
    /* xlsx fills the whole viewport; docx/pptx scroll inside #viewer-root. */
    body.layout-xlsx { overflow: hidden; }
    body.layout-stack { overflow: auto; }
    #viewer-root {
      width: 100%;
      min-height: 100%;
      display: flex;
      justify-content: center;
      align-items: flex-start;
      padding: 16px;
    }
    body.layout-xlsx #viewer-root { padding: 0; height: 100%; }
    #viewer-container { max-width: 100%; width: 100%; }
    body.layout-stack #viewer-container {
      display: flex;
      flex-direction: column;
      align-items: center;
    }
    #status {
      position: fixed;
      inset: 0;
      display: flex;
      align-items: center;
      justify-content: center;
      pointer-events: none;
      z-index: 10;
    }
    #status[data-state="error"] {
      pointer-events: auto;
      color: var(--vscode-errorForeground, #f44747);
      font-size: 13px;
      padding: 16px;
      text-align: center;
    }
    .spinner {
      width: 28px;
      height: 28px;
      border: 3px solid color-mix(in srgb, var(--vscode-foreground) 20%, transparent);
      border-top-color: var(--vscode-progressBar-background, var(--vscode-foreground));
      border-radius: 50%;
      animation: spin 0.9s linear infinite;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
    /* docx / pptx scroll-stack styling */
    .page-stack {
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 16px;
      width: 100%;
    }
    .page-wrapper {
      position: relative;
      width: 100%;
      margin: 0 auto;
    }
    .page-canvas {
      display: block;
      width: 100%;
      background: #fff;
      box-shadow: 0 1px 4px rgba(0, 0, 0, 0.35);
    }
    .text-layer {
      position: absolute;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      overflow: hidden;
      pointer-events: none;
      user-select: text;
      -webkit-user-select: text;
    }
  </style>
</head>
<body class="${fileType === 'xlsx' ? 'layout-xlsx' : 'layout-stack'}">
  <div id="viewer-root">
    <div id="viewer-container">
      <div id="status"><div class="spinner"></div></div>
    </div>
  </div>
  <script nonce="${nonce}">
    window.__OOXML_FILE_TYPE__ = ${JSON.stringify(fileType)};
  </script>
  <script nonce="${nonce}" src="${scriptUri}"></script>
</body>
</html>`;
}

function getNonce(): string {
  let text = '';
  const possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  for (let i = 0; i < 32; i++) {
    text += possible.charAt(Math.floor(Math.random() * possible.length));
  }
  return text;
}
