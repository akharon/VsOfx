import * as vscode from 'vscode';
import type { EditorCommand, HostMessage, WebviewMessage } from '../shared/messages';
import { fromByteArray, toByteArray } from '../shared/base64';
import { pdfEditLabel } from '../shared/pdf';
import { formatFromPath, isEdit, type Edit, type Format } from '../shared/model';
import { OfficeDocument } from './OfficeDocument';
import { buildWebviewHtml, makeNonce } from './webviewHtml';

const SAVE_TIMEOUT_MS = 30_000;
/** Only these commands can be triggered from the webview's toolbar. */
const EDITOR_COMMANDS: Record<EditorCommand, string> = { undo: 'undo', redo: 'redo', save: 'workbench.action.files.save' };

interface PanelInfo {
  document: OfficeDocument;
  format: Format | undefined;
  readonly: boolean;
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
    const readonly = format === 'pptx' || set.size > 0;
    set.add(panel);
    this.docs.set(key, document);
    this.info.set(panel, { document, format, readonly });

    const dist = vscode.Uri.joinPath(this.context.extensionUri, 'dist');
    panel.webview.options = { enableScripts: true, enableCommandUris: false, localResourceRoots: [dist] };
    panel.webview.html = await this.html(panel.webview, dist);

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
            p.resolve(toByteArray(String(msg.bytes)));
          } catch (err) {
            p.reject(err instanceof Error ? err : new Error(String(err)));
          }
        }
        break;
      }
      case 'saveError': {
        const p = this.takePending(msg.requestId);
        if (p) p.reject(new Error(String(msg.message)));
        break;
      }
      case 'error':
        this.log.appendLine(`Error in ${document.uri.toString()}: ${String(msg.message)}`);
        this.errors.push(`webview: ${String(msg.message)}`);
        void vscode.window
          .showErrorMessage(`VsOfx: ${String(msg.message)}`, 'Reopen with Text Editor')
          .then((c) => c && this.reopenAsText(document.uri));
        break;
      case 'reopenAsText':
        this.reopenAsText(document.uri);
        break;
      case 'command': {
        const id = EDITOR_COMMANDS[msg.command as EditorCommand];
        if (id && !this.info.get(panel)?.readonly) void vscode.commands.executeCommand(id);
        break;
      }
      case 'securityViolation':
        this.log.appendLine(`Blocked by CSP: ${String(msg.directive)} ${String(msg.blockedURI)}`);
        this.violations.push({ directive: String(msg.directive), blockedURI: String(msg.blockedURI) });
        break;
      case 'probeResult':
        this.probeResults.push(msg.results);
        break;
    }
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
      void panel.webview.postMessage({ type: 'init', format: 'docx', bytes: '', edits: [], readonly: true } satisfies HostMessage);
      return;
    }
    void panel.webview.postMessage({
      type: 'init',
      format: info.format,
      bytes: fromByteArray(info.document.originalBytes),
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
        this.status.text = `$(file) ${info.format.toUpperCase()}${info.readonly ? ' · read-only' : ''}`;
        this.status.tooltip = info.readonly ? 'VsOfx: this view is read-only' : 'VsOfx editor';
        this.status.show();
        return;
      }
    }
    this.status.hide();
  }

  private async html(webview: vscode.Webview, dist: vscode.Uri): Promise<string> {
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
      probeUri,
      nonce,
    });
  }
}

function editLabel(edit: Edit): string {
  if (edit.kind === 'pdf') return pdfEditLabel(edit.op);
  return edit.kind === 'cell' ? 'Edit cell' : 'Edit paragraph';
}

function errorMessage(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}
