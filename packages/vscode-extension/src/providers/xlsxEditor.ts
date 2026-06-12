import * as vscode from 'vscode';
import { BaseEditorProvider } from './baseEditor';

export class XlsxEditorProvider extends BaseEditorProvider {
  static readonly viewType = 'ooxmlViewer.xlsxEditor';
  protected readonly fileType = 'xlsx' as const;

  static register(context: vscode.ExtensionContext): vscode.Disposable {
    return vscode.window.registerCustomEditorProvider(
      XlsxEditorProvider.viewType,
      new XlsxEditorProvider(context),
    );
  }
}
