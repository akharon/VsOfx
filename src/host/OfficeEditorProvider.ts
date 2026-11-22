import * as os from 'os';
import * as vscode from 'vscode';
import type { EditorCommand, HostMessage, WebviewMessage } from '../shared/messages';
import { asBytes } from '../shared/bytes';
import { pdfEditLabel } from '../shared/pdf';
import { SAFE_HREF } from '../shared/richtext';
import { formatFromPath, isEdit, type Edit, type Format } from '../shared/model';
import { OfficeDocument } from './OfficeDocument';
import { buildWebviewHtml, makeNonce } from './webviewHtml';

const SAVE_TIMEOUT_MS = 30_000;
/** Printed copies: light paper colours instead of the editor theme, nothing scrolls or clips. */
const PRINT_CSS = `
:root { color-scheme: light; }
body.vsofx-print { background: #fff; color: #000; margin: 0; padding: 16px; font-family: system-ui, sans-serif; overflow: visible; }
.vsofx-print * { overflow: visible !important; max-height: none !important; }
.vsofx-print .page { box-shadow: none !important; margin: 0 auto !important; background: #fff !important; color: #000 !important; }
.vsofx-print .print-title { font-size: 14px; margin: 0 0 8px; }
.vsofx-print table.print-sheet { border-collapse: collapse; font-size: 11px; }
.vsofx-print table.print-sheet td { border: 1px solid #999; padding: 2px 4px; white-space: pre-wrap; vertical-align: top; }
@media print { body.vsofx-print { padding: 0; } .vsofx-print .page, .vsofx-print .slide { break-inside: avoid; } }
`;
const MAX_MESSAGE_CHARS = 500;
/** Text from the webview, shortened for notifications and the log. */
export const clip = (v: unknown): string => {
  const s = String(v);
  return s.length > MAX_MESSAGE_CHARS ? `${s.slice(0, MAX_MESSAGE_CHARS)}…` : s;
};
/** Only these commands can be triggered from the webview's toolbar. */
const EDITOR_COMMANDS: Record<EditorCommand, string> = { undo: 'undo', redo: 'redo', save: 'workbench.action.files.save' };

interface PanelInfo {
  document: OfficeDocument;
  format: Format | undefined;
  readonly: boolean;
  /** Short text from the webview for the status bar (word count). */
  stats?: string;
}

interface PendingSave {
  resolve(bytes: Uint8Array): void;
  reject(err: Error): void;
  timer: ReturnType<typeof setTimeout>;
}

/**
 * Generic custom editor for all supported Office formats. The host is a thin
 * byte pipe: it reads/writes files and keeps the original bytes plus the edit
 * list. All parsing happens in the sandboxed webview (doc/arch42.md §4, §8.1).
 */
export class OfficeEditorProvider implements vscode.CustomEditorProvider<OfficeDocument> {
  static readonly viewType = 'vsofx.officeEditor';

  private readonly _onDidChange = new vscode.EventEmitter<vscode.CustomDocumentEditEvent<OfficeDocument>>();
  readonly onDidChangeCustomDocument = this._onDidChange.event;

  private readonly panels = new Map<string, Set<vscode.WebviewPanel>>();
  private readonly docs = new Map<string, OfficeDocument>();
  private readonly info = new WeakMap<vscode.WebviewPanel, PanelInfo>();
  private readonly pendingSaves = new Map<number, PendingSave>();
  private nextRequestId = 1;
  private readonly status: vscode.StatusBarItem;

  /** CSP violations and probe results reported by webviews (used by integration tests). */
  readonly violations: { directive: string; blockedURI: string }[] = [];
  readonly probeResults: Record<string, string>[] = [];
  readonly errors: string[] = [];

  constructor(
    private readonly context: vscode.ExtensionContext,
    private readonly log: vscode.OutputChannel,
  ) {
    this.status = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Right, 50);
    context.subscriptions.push(this.status, this._onDidChange);
  }

  // ---- document lifecycle ------------------------------------------------

  async openCustomDocument(
    uri: vscode.Uri,
    openContext: vscode.CustomDocumentOpenContext,
  ): Promise<OfficeDocument> {
    if (openContext.backupId) {
      try {
        const data = await vscode.workspace.fs.readFile(vscode.Uri.parse(openContext.backupId));
        return OfficeDocument.deserialize(uri, data);
      } catch (err) {
        this.log.appendLine(`Could not restore backup for ${uri.toString()}: ${errorMessage(err)}`);
        // fall through: open the file from disk
      }
    }
    const limitMB = vscode.workspace.getConfiguration('vsofx').get<number>('maxFileSizeMB', 20);
    const stat = await vscode.workspace.fs.stat(uri);
    if (stat.size > limitMB * 1024 * 1024) {
      const mb = (stat.size / 1024 / 1024).toFixed(1);
      const choice = await vscode.window.showWarningMessage(
        `${uri.path.split('/').pop()} is ${mb} MB. Opening large Office files can be slow.`,
        'Open Anyway',
        'Open as Text',
      );
      if (choice !== 'Open Anyway') {
        if (choice === 'Open as Text') void vscode.commands.executeCommand('vscode.openWith', uri, 'default');
        throw new Error('Opening cancelled');
      }
    }
    return new OfficeDocument(uri, await vscode.workspace.fs.readFile(uri));
  }

  async resolveCustomEditor(document: OfficeDocument, panel: vscode.WebviewPanel): Promise<void> {
    const key = document.uri.toString();
    const set = this.panels.get(key) ?? new Set();
    this.panels.set(key, set);
    const format = formatFromPath(document.uri.path);
    // The first editor of a document is editable, further ones are read-only mirrors.
    const readonly = format === 'pptx' || format === 'odp' || format === 'odg' || set.size > 0;
    set.add(panel);
    this.docs.set(key, document);
    this.info.set(panel, { document, format, readonly });

    const dist = vscode.Uri.joinPath(this.context.extensionUri, 'dist');
    panel.webview.options = { enableScripts: true, enableCommandUris: false, localResourceRoots: [dist] };
    panel.webview.html = await this.html(panel.webview, dist, format);

    panel.webview.onDidReceiveMessage((msg: WebviewMessage) => this.onMessage(document, panel, msg));
    panel.onDidChangeViewState(() => this.updateStatus());
    panel.onDidDispose(() => {
      set.delete(panel);
      if (set.size === 0) {
        this.panels.delete(key);
        this.docs.delete(key);
      }
      this.updateStatus();
    });
    this.updateStatus();
  }

  async saveCustomDocument(document: OfficeDocument): Promise<void> {
    try {
      const bytes = await this.requestBytes(document);
      await vscode.workspace.fs.writeFile(document.uri, bytes);
      document.markSaved();
    } catch (err) {
      this.errors.push(`save: ${errorMessage(err)}`);
      throw err;
    }
  }

  async saveCustomDocumentAs(document: OfficeDocument, destination: vscode.Uri): Promise<void> {
    const bytes = await this.requestBytes(document);
    await vscode.workspace.fs.writeFile(destination, bytes);
  }

  async revertCustomDocument(document: OfficeDocument): Promise<void> {
    document.reset(await vscode.workspace.fs.readFile(document.uri));
    for (const panel of this.panelsOf(document)) this.sendInit(panel);
  }

  async backupCustomDocument(
    document: OfficeDocument,
    context: vscode.CustomDocumentBackupContext,
  ): Promise<vscode.CustomDocumentBackup> {
    await vscode.workspace.fs.writeFile(context.destination, document.serialize());
    return {
      id: context.destination.toString(),
      delete: async () => {
        try {
          await vscode.workspace.fs.delete(context.destination);
        } catch {
          // already gone
        }
      },
    };
  }

  // ---- messaging ---------------------------------------------------------

  private onMessage(document: OfficeDocument, panel: vscode.WebviewPanel, msg: WebviewMessage): void {
    if (typeof msg !== 'object' || msg === null) return;
    switch (msg.type) {
      case 'ready':
        this.sendInit(panel);
        if (this.context.extensionMode === vscode.ExtensionMode.Test) void panel.webview.postMessage({ type: 'probe' } satisfies HostMessage);
        break;
      case 'edit':
        if (isEdit(msg.edit) && !this.info.get(panel)?.readonly) this.onUserEdit(document, panel, msg.edit);
        break;
      case 'saveResult': {
        const p = this.takePending(msg.requestId);
        if (p) {
          try {
            p.resolve(asBytes(msg.bytes));
          } catch (err) {
            p.reject(err instanceof Error ? err : new Error(String(err)));
          }
        }
        break;
      }
      case 'saveError': {
        const p = this.takePending(msg.requestId);
        if (p) p.reject(new Error(clip(msg.message)));
        break;
      }
      case 'error':
        {
        // webview text is untrusted: keep notifications and the log readable whatever it sends
        const text = clip(msg.message);
        this.log.appendLine(`Error in ${document.uri.toString()}: ${text}`);
        this.errors.push(`webview: ${text}`);
        void vscode.window
          .showErrorMessage(`VsOfx: ${text}`, 'Reopen with Text Editor')
          .then((c) => c && this.reopenAsText(document.uri));
        break;
      }
      case 'reopenAsText':
        this.reopenAsText(document.uri);
        break;
      case 'openLink': {
        // only web and mail links: never file:, command: or other schemes
        const href = String(msg.href);
        if (SAFE_HREF.test(href)) void vscode.env.openExternal(vscode.Uri.parse(href, true));
        break;
      }
      case 'command': {
        const id = EDITOR_COMMANDS[msg.command as EditorCommand];
        if (id && !this.info.get(panel)?.readonly) void vscode.commands.executeCommand(id);
        break;
      }
      case 'status': {
        const info = this.info.get(panel);
        if (info) {
          info.stats = typeof msg.text === 'string' ? msg.text.slice(0, 40) : undefined;
          this.updateStatus();
        }
        break;
      }
      case 'print':
        void this.print(document, msg).catch((err) => void vscode.window.showErrorMessage(`VsOfx: printing failed: ${clip(err instanceof Error ? err.message : err)}`));
        break;
      case 'securityViolation':
        this.log.appendLine(`Blocked by CSP: ${clip(msg.directive)} ${clip(msg.blockedURI)}`);
        this.violations.push({ directive: clip(msg.directive), blockedURI: clip(msg.blockedURI) });
        break;
      case 'probeResult':
        this.probeResults.push(msg.results);
        break;
    }
  }

  /**
   * Writes a printable copy to the temp folder and opens it in the default browser, whose print
   * dialog opens by itself (HTML) or from its PDF viewer. The HTML gets a strict CSP: only the
   * print script runs, images must be inline.
   */
  private async print(document: OfficeDocument, msg: Extract<WebviewMessage, { type: 'print' }>): Promise<void> {
    const dir = vscode.Uri.joinPath(vscode.Uri.file(os.tmpdir()), 'vsofx-print');
    const base = (document.uri.path.split('/').pop() ?? 'document').replace(/[^\w.\- ]+/g, '_');
    let file: vscode.Uri;
    if ('pdf' in msg) {
      file = vscode.Uri.joinPath(dir, base.toLowerCase().endsWith('.pdf') ? base : `${base}.pdf`);
      await vscode.workspace.fs.writeFile(file, asBytes(msg.pdf));
    } else {
      const css = new TextDecoder().decode(await vscode.workspace.fs.readFile(vscode.Uri.joinPath(this.context.extensionUri, 'dist', 'webview.css')));
      const nonce = makeNonce();
      const title = base.replace(/[<>&"]/g, '_');
      const html = `<!doctype html><html><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'; script-src 'nonce-${nonce}'">
<title>${title}</title><style>${css.replace(/<\/style/gi, '<\\/style')}
${PRINT_CSS}</style></head><body class="vsofx-print">${String(msg.html)}
<script nonce="${nonce}">addEventListener('load', () => setTimeout(() => print(), 300));</script></body></html>`;
      file = vscode.Uri.joinPath(dir, `${base}.html`);
      await vscode.workspace.fs.writeFile(file, new TextEncoder().encode(html));
    }
    await vscode.env.openExternal(file);
  }

  private onUserEdit(document: OfficeDocument, origin: vscode.WebviewPanel | undefined, edit: Edit): void {
    document.addEdit(edit);
    // Mirror the edit into other editors of the same document.
    for (const p of this.panelsOf(document)) if (p !== origin) void p.webview.postMessage({ type: 'apply', edit } satisfies HostMessage);
    this._onDidChange.fire({
      document,
      label: editLabel(edit),
      undo: () => {
        const inverse = document.undo();
        this.broadcast(document, { type: 'apply', edit: inverse });
      },
      redo: () => {
        document.redo(edit);
        this.broadcast(document, { type: 'apply', edit });
      },
    });
  }

  private sendInit(panel: vscode.WebviewPanel): void {
    const info = this.info.get(panel);
    if (!info) return;
    if (!info.format) {
      void panel.webview.postMessage({ type: 'init', format: 'docx', bytes: new Uint8Array(0), edits: [], readonly: true } satisfies HostMessage);
      return;
    }
    void panel.webview.postMessage({
      type: 'init',
      format: info.format,
      bytes: new Uint8Array(info.document.originalBytes), // a plain copy: VS Code transfers it as it is,
      edits: [...info.document.edits],
      readonly: info.readonly,
    } satisfies HostMessage);
  }

  private broadcast(document: OfficeDocument, msg: HostMessage): void {
    for (const p of this.panelsOf(document)) void p.webview.postMessage(msg);
  }

  private panelsOf(document: OfficeDocument): vscode.WebviewPanel[] {
    return [...(this.panels.get(document.uri.toString()) ?? [])];
  }

  /** Asks a live webview of the document for the saved file bytes. */
  private requestBytes(document: OfficeDocument): Promise<Uint8Array> {
    const live = this.panelsOf(document);
    const panel = live.find((p) => p.active) ?? live[0];
    if (!panel) {
      const msg = 'VsOfx: the editor for this file is not open. Reopen the file to save your changes.';
      void vscode.window.showErrorMessage(msg);
      return Promise.reject(new Error(msg));
    }
    const requestId = this.nextRequestId++;
    return new Promise<Uint8Array>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pendingSaves.delete(requestId);
        reject(new Error('VsOfx: saving timed out'));
      }, SAVE_TIMEOUT_MS);
      this.pendingSaves.set(requestId, { resolve, reject, timer });
      void panel.webview.postMessage({ type: 'requestSave', requestId } satisfies HostMessage);
    });
  }

  private takePending(requestId: number): PendingSave | undefined {
    const p = this.pendingSaves.get(requestId);
    if (p) {
      clearTimeout(p.timer);
      this.pendingSaves.delete(requestId);
    }
    return p;
  }

  reopenAsText(uri: vscode.Uri): void {
    void vscode.commands.executeCommand('vscode.openWith', uri, 'default');
  }

  /** Test hook: behaves as if the user had made `edit` in the editor (no webview is the origin). */
  injectEdit(uri: vscode.Uri, edit: Edit): void {
    const doc = this.docs.get(uri.toString());
    if (!doc) throw new Error(`No open editor for ${uri.toString()}`);
    this.onUserEdit(doc, undefined, edit);
  }

  /** Test hook: ask every open webview to run its security probe again. */
  probeAll(): void {
    for (const set of this.panels.values()) for (const p of set) void p.webview.postMessage({ type: 'probe' } satisfies HostMessage);
  }

  // ---- ui ---------------------------------------------------------------

  private updateStatus(): void {
    for (const set of this.panels.values()) {
      for (const p of set) {
        if (!p.active) continue;
        const info = this.info.get(p);
        if (!info?.format) break;
        this.status.text = `$(file) ${info.format.toUpperCase()}${info.readonly ? ' · read-only' : ''}${info.stats ? ` · ${info.stats.replace(/\$\(/g, '$ (')}` : ''}`;
        this.status.tooltip = info.readonly ? 'VsOfx: this view is read-only' : 'VsOfx editor';
        this.status.show();
        return;
      }
    }
    this.status.hide();
  }

  private async html(webview: vscode.Webview, dist: vscode.Uri, format?: Format): Promise<string> {
    const nonce = makeNonce();
    let probeUri: string | undefined;
    if (this.context.extensionMode === vscode.ExtensionMode.Test) {
      try {
        const probe = vscode.Uri.joinPath(dist, 'probe.js');
        await vscode.workspace.fs.stat(probe);
        probeUri = webview.asWebviewUri(probe).toString();
      } catch {
        // no test build
      }
    }
    return buildWebviewHtml({
      cspSource: webview.cspSource,
      scriptUri: webview.asWebviewUri(vscode.Uri.joinPath(dist, 'webview.js')).toString(),
      styleUri: webview.asWebviewUri(vscode.Uri.joinPath(dist, 'webview.css')).toString(),
      pdfChunkUri: webview.asWebviewUri(vscode.Uri.joinPath(dist, 'pdf.js')).toString(),
      probeUri,
      format,
      nonce,
    });
  }
}

function editLabel(edit: Edit): string {
  if (edit.kind === 'pdf') return pdfEditLabel(edit.op);
  if (edit.kind === 'table') return edit.op === 'insert' ? 'Insert table' : 'Remove table';
  if (edit.kind === 'newpara') return edit.op === 'insert' ? 'New paragraph' : 'Remove paragraph';
  if (edit.kind === 'delpara') return 'Delete paragraph';
  if (edit.kind === 'sheet') return edit.op === 'add' ? 'Add sheet' : edit.op === 'remove' ? 'Delete sheet' : 'Rename sheet';
  if (edit.kind === 'colwidth') return 'Column width';
  if (edit.kind === 'wrap') return 'Wrap text';
  if (edit.kind === 'align') return 'Align';
  if (edit.kind === 'lines') return `${edit.op === 'insert' ? 'Insert' : 'Delete'} ${edit.axis === 'row' ? 'row' : 'column'}${edit.count > 1 ? 's' : ''}`;
  if (edit.kind === 'batch') {
    const lines = edit.edits.find((e) => e.kind === 'lines');
    if (lines) return editLabel(lines);
    if (edit.edits.some((e) => e.kind === 'wrap')) return 'Wrap text';
    if (edit.edits.some((e) => e.kind === 'align')) return 'Align';
    if (edit.edits.some((e) => e.kind === 'delpara')) return 'Edit text';
    if (edit.edits.some((e) => e.kind === 'newpara')) return 'New paragraph';
    return edit.edits.every((e) => e.kind === 'cell' && e.value === '') ? 'Clear cells' : 'Paste';
  }
  if (edit.kind === 'paragraph' && edit.fmt && edit.oldFmt && edit.text === edit.old) return 'Format paragraph';
  return edit.kind === 'cell' ? 'Edit cell' : 'Edit paragraph';
}

function errorMessage(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}
