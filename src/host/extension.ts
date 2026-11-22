import * as vscode from 'vscode';
import type { Edit } from '../shared/model';
import { NEW_DOCS, newDocumentBytes, type NewDocKind } from './newDocument';
import { OfficeEditorProvider } from './OfficeEditorProvider';

/** Set by esbuild.mjs: "<version> (<commit>)" of this build. */
declare const __BUILD__: string | undefined;

export function activate(context: vscode.ExtensionContext) {
  const log = vscode.window.createOutputChannel('VsOfx');
  log.appendLine(`VsOfx build ${typeof __BUILD__ === 'string' ? __BUILD__ : 'dev'}`);
  const provider = new OfficeEditorProvider(context, log);
  const retain = vscode.workspace.getConfiguration('vsofx').get<boolean>('retainContextWhenHidden', true);

  context.subscriptions.push(
    log,
    vscode.window.registerCustomEditorProvider(OfficeEditorProvider.viewType, provider, {
      webviewOptions: { retainContextWhenHidden: retain },
      supportsMultipleEditorsPerDocument: true,
    }),
    vscode.commands.registerCommand('vsofx.reopenWithTextEditor', () => {
      const input = vscode.window.tabGroups.activeTabGroup.activeTab?.input;
      if (input instanceof vscode.TabInputCustom && input.viewType === OfficeEditorProvider.viewType) {
        provider.reopenAsText(input.uri);
      } else {
        void vscode.window.showInformationMessage('VsOfx: the active tab is not an Office document.');
      }
    }),
    vscode.commands.registerCommand('vsofx.newDocument', () => newDocument()),
  );

  // Exposed to integration tests only.
  if (context.extensionMode === vscode.ExtensionMode.Test) {
    return {
      violations: provider.violations,
      probeResults: provider.probeResults,
      errors: provider.errors,
      probeAll: () => provider.probeAll(),
      injectEdit: (uri: vscode.Uri, edit: Edit) => provider.injectEdit(uri, edit),
    };
  }
  return undefined;
}

/** "VsOfx: New Document": pick a type, choose where to save, create the empty file and open it. */
async function newDocument(): Promise<void> {
  const items: (vscode.QuickPickItem & { doc: NewDocKind })[] = NEW_DOCS.map((d) => ({ label: d.label, description: d.detail, doc: d.kind }));
  const pick = await vscode.window.showQuickPick(items, { placeHolder: 'New document' });
  if (!pick) return;
  const folder = vscode.workspace.workspaceFolders?.[0]?.uri;
  const uri = await vscode.window.showSaveDialog({
    defaultUri: folder ? vscode.Uri.joinPath(folder, `Untitled.${pick.doc}`) : undefined,
    filters: { [pick.label]: [pick.doc] },
    saveLabel: 'Create',
  });
  if (!uri) return;
  const target = uri.path.toLowerCase().endsWith(`.${pick.doc}`) ? uri : uri.with({ path: `${uri.path}.${pick.doc}` });
  try {
    await vscode.workspace.fs.writeFile(target, newDocumentBytes(pick.doc));
    await vscode.commands.executeCommand('vscode.openWith', target, OfficeEditorProvider.viewType);
  } catch (err) {
    void vscode.window.showErrorMessage(`VsOfx: could not create the document: ${err instanceof Error ? err.message : String(err)}`);
  }
}

export function deactivate(): void {
  // nothing to clean up; subscriptions are disposed by VS Code
}
