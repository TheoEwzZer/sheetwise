import * as vscode from 'vscode';
import { XlsxEditorProvider, VIEW_TYPE } from './provider.ts';
import { XlsxDocument } from './xlsxDocument.ts';

export function activate(context: vscode.ExtensionContext): void {
  context.subscriptions.push(XlsxEditorProvider.register(context));

  context.subscriptions.push(
    vscode.commands.registerCommand('xlsxEditor.newWorkbook', async () => {
      const folder = vscode.workspace.workspaceFolders?.[0]?.uri;
      const target = await vscode.window.showSaveDialog({
        defaultUri: folder ? vscode.Uri.joinPath(folder, 'Workbook.xlsx') : undefined,
        filters: { 'Excel Workbook': ['xlsx'] },
        saveLabel: 'Create',
      });
      if (!target) return;
      await vscode.workspace.fs.writeFile(target, await XlsxDocument.emptyWorkbookBytes());
      await vscode.commands.executeCommand('vscode.openWith', target, VIEW_TYPE);
    }),
  );
}

export function deactivate(): void {}
