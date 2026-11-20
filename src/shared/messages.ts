import type { Edit, Format } from './model';

/**
 * File bytes travel as base64 strings: VS Code serialises messages as JSON, and a Node Buffer would
 * arrive as {type:'Buffer',data:[...]}.
 */

/** Toolbar buttons that ask the host to run a standard editor command. */
export type EditorCommand = 'undo' | 'redo' | 'save';

/** host -> webview */
export type HostMessage =
  | { type: 'init'; format: Format; bytes: string; edits: Edit[]; readonly: boolean }
  | { type: 'apply'; edit: Edit }
  | { type: 'requestSave'; requestId: number }
  | { type: 'diagram'; id: number; html: string | null }
  | { type: 'probe' }; // only sent in extension test mode

/** webview -> host */
export type WebviewMessage =
  | { type: 'ready' }
  | { type: 'edit'; edit: Edit }
  | { type: 'saveResult'; requestId: number; bytes: string }
  | { type: 'saveError'; requestId: number; message: string }
  | { type: 'error'; message: string }
  | { type: 'reopenAsText' }
  | { type: 'openLink'; href: string }
  | { type: 'renderDiagram'; id: number; lang: string; source: string }
  | { type: 'command'; command: EditorCommand }
  | { type: 'securityViolation'; directive: string; blockedURI: string }
  | { type: 'probeResult'; results: Record<string, string> };
