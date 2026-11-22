// Per-document view state (zoom, scroll, active sheet) kept by VS Code through the webview's
// getState / setState, and the status text (word count) the host shows in the status bar.
// Plain data only; every read and write is guarded so a missing API never breaks the editor.
export interface ViewState {
  zoom?: number;
  /** Scroll positions of the scrollable areas, by the first CSS class of the element. */
  scroll?: { cls: string; top: number; left: number }[];
  /** Index of the active sheet. */
  sheet?: number;
}

interface StateApi {
  getState(): unknown;
  setState(s: unknown): void;
}

let api: StateApi | undefined;
let state: ViewState = {};
let sink: ((text: string) => void) | undefined;

export function initViewState(a: StateApi): void {
  api = a;
  try {
    const s = a.getState();
    if (s && typeof s === 'object') state = s as ViewState;
  } catch {
    // no saved state
  }
}

export const viewState = (): ViewState => state;

export function patchViewState(p: Partial<ViewState>): void {
  state = { ...state, ...p };
  try {
    api?.setState(state);
  } catch {
    // not saved
  }
}

/** Where `reportStatus` text goes (main.ts posts it to the host). */
export function setStatusSink(f: ((text: string) => void) | undefined): void {
  sink = f;
}
export const reportStatus = (text: string): void => sink?.(text);
