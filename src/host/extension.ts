import * as vscode from 'vscode';
import type { Edit } from '../shared/model';
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

export function deactivate(): void {
  // nothing to clean up; subscriptions are disposed by VS Code
}
