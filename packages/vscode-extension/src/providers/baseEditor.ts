import * as vscode from 'vscode';
import { getWebviewHtml } from '../webviewHtml';
import { shouldUseGoogleFonts } from '../config';

/** Live webview panel + everything needed to rebuild its HTML, tracked so the
 *  HTML can be regenerated when the effective Google Fonts flag changes
 *  (CSP differs → full reload needed). */
interface OpenView {
  panel: vscode.WebviewPanel;
  extensionUri: vscode.Uri;
  fileType: 'docx' | 'xlsx' | 'pptx';
}

/** Registry of every open OOXML webview across all three editor providers, so a
 *  single config/trust change can refresh them all at once. */
const openViews = new Set<OpenView>();

/**
 * Rebuild the HTML of every open OOXML webview.
 *
 * Called when `ooxmlViewer.useGoogleFonts` changes or workspace trust is granted:
 * the CSP is baked into the document `<head>`, so toggling Google Fonts requires a
 * full HTML reload rather than a re-render message. Re-setting `webview.html`
 * re-runs the bootstrap, which re-requests the bytes and re-instantiates the viewer
 * with the new flag.
 */
export function refreshAllWebviews(): void {
  for (const view of openViews) {
    renderWebview(view);
  }
}

/** Set the webview HTML for a panel using the current effective flag. */
function renderWebview(view: OpenView): void {
  view.panel.webview.html = getWebviewHtml(
    view.panel.webview,
    view.extensionUri,
    view.fileType,
    shouldUseGoogleFonts(),
  );
}

/**
 * Shared implementation for the docx / xlsx / pptx readonly custom editors. The
 * three formats differ only by their `viewType` and `fileType` tag; everything
 * else (CSP-gated HTML, byte hand-off, copy bridge, Google-Fonts wiring) is common.
 */
export abstract class BaseEditorProvider
  implements vscode.CustomReadonlyEditorProvider
{
  protected abstract readonly fileType: 'docx' | 'xlsx' | 'pptx';

  constructor(protected readonly context: vscode.ExtensionContext) {}

  async openCustomDocument(uri: vscode.Uri): Promise<vscode.CustomDocument> {
    return { uri, dispose: () => undefined };
  }

  async resolveCustomEditor(
    document: vscode.CustomDocument,
    webviewPanel: vscode.WebviewPanel,
  ): Promise<void> {
    webviewPanel.webview.options = {
      enableScripts: true,
      localResourceRoots: [
        vscode.Uri.joinPath(this.context.extensionUri, 'dist'),
        vscode.Uri.joinPath(document.uri, '..'),
      ],
    };

    const view: OpenView = {
      panel: webviewPanel,
      extensionUri: this.context.extensionUri,
      fileType: this.fileType,
    };
    openViews.add(view);
    webviewPanel.onDidDispose(() => {
      openViews.delete(view);
    });

    renderWebview(view);

    const docUrl = webviewPanel.webview.asWebviewUri(document.uri).toString();

    webviewPanel.webview.onDidReceiveMessage(async (msg) => {
      if (msg.type === 'webview-ready') {
        await webviewPanel.webview.postMessage({
          type: 'ooxml-init',
          fileType: this.fileType,
          url: docUrl,
          useGoogleFonts: shouldUseGoogleFonts(),
        });
      } else if (msg.type === 'copy') {
        vscode.env.clipboard.writeText(msg.text);
      }
    });
  }
}
