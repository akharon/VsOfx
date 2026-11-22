import type { Edit, Format } from './model';

/**
 * File bytes travel as typed arrays: VS Code webviews transfer Uint8Array/ArrayBuffer between the host
 * and the page without JSON or base64 (in the webview they can be handed on to the worker as
 * transferables). The receiver reads them with asBytes() (shared/bytes.ts), which also accepts the
 * older base64 string and a serialised Node Buffer.
 */

/** Toolbar buttons that ask the host to run a standard editor command. */
export type EditorCommand = 'undo' | 'redo' | 'save';

/** host -> webview */
export type HostMessage =
  | { type: 'init'; format: Format; bytes: Uint8Array; edits: Edit[]; readonly: boolean }
  | { type: 'apply'; edit: Edit }
  | { type: 'requestSave'; requestId: number }
  | { type: 'probe' }; // only sent in extension test mode

/** webview -> host */
export type WebviewMessage =
  | { type: 'ready' }
  | { type: 'edit'; edit: Edit }
  | { type: 'saveResult'; requestId: number; bytes: Uint8Array }
  | { type: 'saveError'; requestId: number; message: string }
  | { type: 'error'; message: string }
  | { type: 'reopenAsText' }
  | { type: 'openLink'; href: string }
  | { type: 'command'; command: EditorCommand }
  | { type: 'status'; text: string } // short text for the VS Code status bar (word count)
  // printable copy, opened in the system browser (webviews cannot print themselves)
  | { type: 'print'; html: string }
  | { type: 'print'; pdf: Uint8Array }
  | { type: 'securityViolation'; directive: string; blockedURI: string }
  | { type: 'probeResult'; results: Record<string, string> };
