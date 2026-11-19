import * as vscode from 'vscode';
import { XlsxDocument, type DocSettings } from './xlsxDocument.ts';
import type { Settings, ToHost, ToWebview } from './protocol.ts';

export const VIEW_TYPE = 'xlsxEditor.spreadsheet';

function readSettings(): DocSettings & Settings {
  const cfg = vscode.workspace.getConfiguration('xlsxEditor');
  const loc = cfg.get<string>('locale', 'auto');
  const locale = loc === 'fr' || loc === 'en' ? loc : vscode.env.language.toLowerCase().startsWith('fr') ? 'fr' : 'en';
  const gridTheme = cfg.get<string>('gridTheme', 'excel') === 'vscode' ? 'vscode' : 'excel';
  return { locale, gridTheme, recalc: cfg.get<boolean>('recalculate', true) };
}

export class XlsxEditorProvider implements vscode.CustomEditorProvider<XlsxDocument> {
  static register(context: vscode.ExtensionContext): vscode.Disposable {
    const provider = new XlsxEditorProvider(context);
    return vscode.Disposable.from(
      vscode.window.registerCustomEditorProvider(VIEW_TYPE, provider, {
        webviewOptions: { retainContextWhenHidden: true, enableFindWidget: true },
        supportsMultipleEditorsPerDocument: true,
      }),
      vscode.workspace.onDidChangeConfiguration((e) => {
        if (e.affectsConfiguration('xlsxEditor')) provider.settingsChanged();
      }),
    );
  }

  private readonly _onDidChangeCustomDocument = new vscode.EventEmitter<vscode.CustomDocumentEditEvent<XlsxDocument>>();
  readonly onDidChangeCustomDocument = this._onDidChangeCustomDocument.event;

  private readonly panels = new Map<XlsxDocument, Set<vscode.WebviewPanel>>();

  constructor(private readonly context: vscode.ExtensionContext) {}

  private settingsChanged(): void {
    const s = readSettings();
    for (const [doc, panels] of this.panels) {
      doc.updateSettings(s);
      for (const p of panels) this.post(p, { type: 'init', meta: doc.meta(), settings: { locale: s.locale, gridTheme: s.gridTheme } });
    }
  }

  async openCustomDocument(
    uri: vscode.Uri,
    openContext: vscode.CustomDocumentOpenContext,
    _token: vscode.CancellationToken,
  ): Promise<XlsxDocument> {
    const doc = await XlsxDocument.create(uri, openContext.backupId, readSettings());
    doc.onDidEdit((e) => {
      this._onDidChangeCustomDocument.fire({ document: doc, label: e.label, undo: e.undo, redo: e.redo });
    });
    doc.onDidUpdate((msg) => {
      for (const p of this.panels.get(doc) ?? []) this.post(p, msg);
    });
    doc.onDidDispose(() => this.panels.delete(doc));
    return doc;
  }

  async resolveCustomEditor(doc: XlsxDocument, panel: vscode.WebviewPanel, _token: vscode.CancellationToken): Promise<void> {
    let set = this.panels.get(doc);
    if (!set) this.panels.set(doc, (set = new Set()));
    set.add(panel);
    panel.onDidDispose(() => set!.delete(panel));

    const media = vscode.Uri.joinPath(this.context.extensionUri, 'media');
    panel.webview.options = { enableScripts: true, localResourceRoots: [media] };
    panel.webview.html = this.html(panel.webview);
    panel.webview.onDidReceiveMessage((msg: ToHost) => this.onMessage(doc, panel, msg));
  }

  private post(panel: vscode.WebviewPanel, msg: ToWebview): void {
    void panel.webview.postMessage(msg);
  }

  private onMessage(doc: XlsxDocument, panel: vscode.WebviewPanel, msg: ToHost): void {
    switch (msg.type) {
      case 'ready': {
        const s = readSettings();
        this.post(panel, { type: 'init', meta: doc.meta(), settings: { locale: s.locale, gridTheme: s.gridTheme } });
        break;
      }
      case 'getSheet':
        void doc.run(() => {
          const data = doc.sheetModel(msg.index);
          if (data) this.post(panel, { type: 'sheet', index: msg.index, data });
        });
        break;
      case 'setCells':
        void doc.setCells(msg.sheet, msg.items);
        break;
      case 'clear':
        void doc.clear(msg.sheet, msg.ranges, msg.what);
        break;
      case 'format':
        void doc.format(msg.sheet, msg.ranges, msg.patch);
        break;
      case 'paste':
        void doc.paste(msg.sheet, msg.target, msg.rows, msg.internal);
        break;
      case 'structure':
        void doc.structure(msg.sheet, msg.op, msg.at, msg.count);
        break;
      case 'merge':
        void doc.merge(msg.sheet, msg.range, msg.unmerge);
        break;
      case 'colWidth':
        void doc.colWidth(msg.sheet, msg.cols, msg.px);
        break;
      case 'rowHeight':
        void doc.rowHeight(msg.sheet, msg.rows, msg.px);
        break;
      case 'addSheet':
        void doc.addSheet();
        break;
      case 'renameSheet':
        void doc.renameSheet(msg.index, msg.name);
        break;
      case 'deleteSheet':
        void (async () => {
          const name = doc.wb.worksheets[msg.index]?.name ?? '';
          const ok = await vscode.window.showWarningMessage(
            `Delete sheet "${name}"? (Ctrl+Z to undo)`,
            { modal: true },
            'Delete',
          );
          if (ok === 'Delete') await doc.deleteSheet(msg.index);
        })();
        break;
      case 'getList':
        void doc.run(() => {
          this.post(panel, { type: 'list', r: msg.r, c: msg.c, values: doc.listValues(msg.sheet, msg.r, msg.c) });
        });
        break;
      case 'openLink': {
        const url = String(msg.url ?? '');
        if (/^(https?:|mailto:)/i.test(url)) void vscode.env.openExternal(vscode.Uri.parse(url));
        break;
      }
      case 'error':
        console.error('[xlsx-editor webview]', msg.message);
        break;
    }
  }

  private html(webview: vscode.Webview): string {
    const media = vscode.Uri.joinPath(this.context.extensionUri, 'media');
    const script = webview.asWebviewUri(vscode.Uri.joinPath(media, 'webview.js'));
    const style = webview.asWebviewUri(vscode.Uri.joinPath(media, 'webview.css'));
    const nonce = Array.from({ length: 32 }, () => Math.floor(Math.random() * 36).toString(36)).join('');
    return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src ${webview.cspSource} data:; style-src ${webview.cspSource} 'unsafe-inline'; font-src ${webview.cspSource}; script-src 'nonce-${nonce}';">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link href="${style}" rel="stylesheet">
<title>Sheetwise</title>
</head>
<body>
<div id="app"></div>
<script nonce="${nonce}" src="${script}"></script>
</body>
</html>`;
  }

  saveCustomDocument(doc: XlsxDocument, cancel: vscode.CancellationToken): Thenable<void> {
    return doc.save(cancel);
  }

  saveCustomDocumentAs(doc: XlsxDocument, destination: vscode.Uri, cancel: vscode.CancellationToken): Thenable<void> {
    return doc.saveAs(destination, cancel);
  }

  revertCustomDocument(doc: XlsxDocument, cancel: vscode.CancellationToken): Thenable<void> {
    return doc.revert(cancel);
  }

  backupCustomDocument(
    doc: XlsxDocument,
    context: vscode.CustomDocumentBackupContext,
    cancel: vscode.CancellationToken,
  ): Thenable<vscode.CustomDocumentBackup> {
    return doc.backup(context.destination, cancel);
  }
}
