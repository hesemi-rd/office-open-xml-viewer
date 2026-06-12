import * as vscode from 'vscode';
import { DocxEditorProvider } from './providers/docxEditor';
import { XlsxEditorProvider } from './providers/xlsxEditor';
import { PptxEditorProvider } from './providers/pptxEditor';
import { refreshAllWebviews } from './providers/baseEditor';
import { USE_GOOGLE_FONTS_CONFIG_ID } from './config';
import { activateMcp } from './mcp';

export function activate(context: vscode.ExtensionContext): void {
  context.subscriptions.push(
    DocxEditorProvider.register(context),
    XlsxEditorProvider.register(context),
    PptxEditorProvider.register(context),
  );

  // The Google Fonts opt-in is baked into the webview CSP, so changing it (or
  // gaining workspace trust, which can flip the effective value) requires
  // regenerating the HTML of every open OOXML preview rather than a soft
  // re-render. Re-setting `webview.html` reloads the bootstrap with the new flag.
  context.subscriptions.push(
    vscode.workspace.onDidChangeConfiguration((e) => {
      if (e.affectsConfiguration(USE_GOOGLE_FONTS_CONFIG_ID)) {
        refreshAllWebviews();
      }
    }),
    vscode.workspace.onDidGrantWorkspaceTrust(() => {
      refreshAllWebviews();
    }),
  );

  activateMcp(context);
}

export function deactivate(): void {
  // nothing to clean up
}
