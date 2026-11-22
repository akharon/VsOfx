// Webview entry (UI thread). Third-party parsing is delegated to the worker.
import './style.css';
import './code.css';
import { ownedBytes } from '../shared/bytes';
import { skeletonHtml, skeletonKind } from '../shared/skeleton';
import { workerFamily, type WorkerFamily } from './family';
import type { HostMessage, WebviewMessage } from '../shared/messages';
import { SAFE_HREF } from '../shared/richtext';
import type { ApplyResult, DocModel, Edit } from '../shared/model';
import type { LoadResult, WorkerRequest, WorkerResponse } from './protocol';
import { createToolbar } from './ui/toolbar';
import type { PdfChunk } from './pdfChunk';
import { renderSheetView } from './views/sheetView';
import { renderSlidesView } from './views/slidesView';
import { renderTextView } from './views/textView';
import type { View } from './views/types';
import { initViewState, patchViewState, setStatusSink, viewState } from './viewState';

declare function acquireVsCodeApi(): {
  postMessage(msg: WebviewMessage, transfer?: Transferable[]): void;
  getState(): unknown;
  setState(s: unknown): void;
};

const vscode = acquireVsCodeApi();
initViewState(vscode);
const app = document.getElementById('app')!;

let worker: Worker | undefined;
let view: View | undefined;
let generation = 0;
let nextId = 1;
let loaded = false;
/** Original bytes of an open PDF (pdf.js reads its own copy; the worker owns the other). */
let pdfBytes: Uint8Array | undefined;
let pdfWorker: Worker | undefined;
/** Edits from the host that arrive while the document is still loading. */
let queued: Edit[] = [];
const pending = new Map<number, { resolve(v: unknown): void; reject(e: Error): void }>();

function post(msg: WebviewMessage, transfer?: Transferable[]): void {
  vscode.postMessage(msg, transfer);
}

// ---- status bar text and remembered view (zoom, scroll, active sheet) ----------
let statusTimer = 0;
setStatusSink((text) => {
  clearTimeout(statusTimer);
  statusTimer = window.setTimeout(() => post({ type: 'status', text }), 200);
});

let restoring = false;
const scrolls = new Map<string, { cls: string; top: number; left: number }>();
let scrollTimer = 0;
document.addEventListener(
  'scroll',
  (ev) => {
    const t = ev.target;
    const cls = t instanceof HTMLElement && app.contains(t) ? t.classList[0] : undefined;
    if (!cls || restoring) return;
    scrolls.set(cls, { cls, top: Math.round((t as HTMLElement).scrollTop), left: Math.round((t as HTMLElement).scrollLeft) });
    clearTimeout(scrollTimer);
    scrollTimer = window.setTimeout(() => patchViewState({ scroll: [...scrolls.values()] }), 300);
  },
  true,
);
// remember the sheet tab that was clicked (tabs are rebuilt by the view, so count them before it reacts)
document.addEventListener(
  'click',
  (ev) => {
    const tab = (ev.target as Element | null)?.closest?.('.tabs [role="tab"]');
    if (!tab || restoring) return;
    patchViewState({ sheet: Array.from(tab.parentElement?.querySelectorAll('[role="tab"]') ?? []).indexOf(tab) });
  },
  true,
);

/** Brings back the sheet and the scroll positions of the last session of this document. */
function restoreView(): void {
  const st = viewState();
  restoring = true;
  try {
    if (typeof st.sheet === 'number' && st.sheet > 0) {
      app.querySelectorAll<HTMLElement>('.tabs [role="tab"]')[st.sheet]?.click();
    }
    for (const s of Array.isArray(st.scroll) ? st.scroll : []) {
      const target = typeof s?.cls === 'string' && /^[\w-]+$/.test(s.cls) ? app.querySelector<HTMLElement>(`.${s.cls}`) : null;
      if (!target || typeof s.top !== 'number' || typeof s.left !== 'number') continue;
      target.scrollTop = s.top;
      target.scrollLeft = s.left;
      scrolls.set(s.cls, s);
    }
  } finally {
    window.setTimeout(() => (restoring = false), 50); // scroll events of the restore itself
  }
}

// ---- links ---------------------------------------------------------------
// Links never navigate inside the webview. Ctrl/Cmd+click (or a plain click outside editable text)
// asks the host to open web and mail links in the browser / mail client.
function linkHref(a: Element): string {
  const href = a.getAttribute('data-href') ?? a.getAttribute('href') ?? '';
  return SAFE_HREF.test(href) ? href : '';
}
document.addEventListener(
  'click',
  (ev) => {
    const a = (ev.target as Element | null)?.closest?.('a');
    const href = a ? linkHref(a) : '';
    if (!a || !href) return;
    ev.preventDefault();
    if (!(a as HTMLElement).isContentEditable || ev.ctrlKey || ev.metaKey) {
      ev.stopPropagation();
      post({ type: 'openLink', href });
    }
  },
  true,
);
document.addEventListener('mouseover', (ev) => {
  const a = (ev.target as Element | null)?.closest?.('a') as HTMLElement | null;
  const href = a ? linkHref(a) : '';
  if (a && href) a.title = a.isContentEditable ? `${href}\nCtrl+click to open` : href;
});

// ---- worker --------------------------------------------------------------

function blobWorker(source: string, options?: WorkerOptions): Worker {
  // blob: workers inherit this page's CSP, so connect-src 'none' applies to them as well.
  return new Worker(URL.createObjectURL(new Blob([source], { type: 'text/javascript' })), options);
}

// ---- PDF chunk (dist/pdf.js: pdf.js, PDF view, pdf.js worker), loaded on first use -----------
// The browser hides the nonce attribute after load, but the element's nonce property stays readable.
const scriptNonce = (document.currentScript as HTMLScriptElement | null)?.nonce ?? '';
let pdfChunk: Promise<PdfChunk> | undefined;

function loadPdfChunk(): Promise<PdfChunk> {
  const ready = (globalThis as { vsofxPdf?: PdfChunk }).vsofxPdf;
  if (ready) return Promise.resolve(ready);
  pdfChunk ??= new Promise<PdfChunk>((resolve, reject) => {
    const src = app.dataset['pdfChunk'];
    if (!src) return reject(new Error('The PDF viewer is not available (no script URL).'));
    const s = document.createElement('script');
    s.nonce = scriptNonce;
    s.src = src;
    s.onload = () => {
      const c = (globalThis as { vsofxPdf?: PdfChunk }).vsofxPdf;
      if (c) resolve(c);
      else reject(new Error('The PDF viewer did not initialise.'));
    };
    s.onerror = () => reject(new Error('The PDF viewer could not be loaded.'));
    document.head.appendChild(s);
  }).catch((err) => {
    pdfChunk = undefined; // allow a retry on the next PDF
    throw err;
  });
  return pdfChunk;
}

// ---- format worker chunks (dist/worker-sheet.js, dist/worker-doc.js), loaded per family -------
// They sit next to dist/pdf.js, whose URL the host passes; the PDF worker source is inside that chunk.
const workerChunks = new Map<WorkerFamily, Promise<string>>();

function loadWorkerSource(family: WorkerFamily): Promise<string> {
  const global = globalThis as { vsofxWorkers?: Record<string, string> };
  const ready = global.vsofxWorkers?.[family];
  if (ready) return Promise.resolve(ready);
  let p = workerChunks.get(family);
  if (!p) {
    p = new Promise<string>((resolve, reject) => {
      const base = app.dataset['pdfChunk'];
      if (!base) return reject(new Error('The document worker is not available (no script URL).'));
      const s = document.createElement('script');
      s.nonce = scriptNonce;
      s.src = new URL(`worker-${family}.js`, base).toString();
      s.onload = () => {
        const src = global.vsofxWorkers?.[family];
        if (src) resolve(src);
        else reject(new Error('The document worker did not initialise.'));
      };
      s.onerror = () => reject(new Error('The document worker could not be loaded.'));
      document.head.appendChild(s);
    }).catch((err) => {
      workerChunks.delete(family); // allow a retry
      throw err;
    });
    workerChunks.set(family, p);
  }
  return p;
}

function startWorker(source: string): Worker {
  worker?.terminate();
  for (const p of pending.values()) p.reject(new Error('Worker restarted'));
  pending.clear();
  const w = blobWorker(source);
  w.addEventListener('message', (ev: MessageEvent<WorkerResponse>) => {
    const res = ev.data;
    const p = pending.get(res.id);
    if (!p) return;
    pending.delete(res.id);
    if (res.ok) p.resolve(res.result);
    else p.reject(new Error(res.message));
  });
  w.addEventListener('error', (ev) => {
    const err = new Error(ev.message || 'The worker failed');
    for (const p of pending.values()) p.reject(err);
    pending.clear();
  });
  worker = w;
  return w;
}

function rpc<T>(req: DistributiveOmit<WorkerRequest, 'id'>, transfer: Transferable[] = []): Promise<T> {
  const w = worker;
  if (!w) return Promise.reject(new Error('No worker'));
  const id = nextId++;
  return new Promise<T>((resolve, reject) => {
    pending.set(id, { resolve: resolve as (v: unknown) => void, reject });
    w.postMessage({ ...req, id }, transfer);
  });
}
type DistributiveOmit<T, K extends keyof never> = T extends unknown ? Omit<T, K> : never;

// ---- rendering -----------------------------------------------------------

function showLoading(format?: string): void {
  const kind = skeletonKind(format);
  // the host's HTML already shows this skeleton: keep it (no restart of the shimmer)
  if (app.firstElementChild?.classList.contains(`sk-${kind}`)) return;
  app.innerHTML = skeletonHtml(kind); // static markup, see shared/skeleton.ts
}

function showError(message: string): void {
  closeView();
  app.replaceChildren();
  const box = document.createElement('div');
  box.className = 'error';
  const h = document.createElement('h2');
  h.textContent = 'Cannot open this file';
  const p = document.createElement('p');
  p.textContent = message;
  const b = document.createElement('button');
  b.type = 'button';
  b.textContent = 'Open with Text Editor';
  b.addEventListener('click', () => post({ type: 'reopenAsText' }));
  box.append(h, p, b);
  app.appendChild(box);
}

function render(model: DocModel, readonly: boolean): void {
  closeView();
  app.replaceChildren();
  const onEdit = (edit: Edit) => void userEdit(edit);
  const tb = createToolbar();
  const content = document.createElement('div');
  content.className = 'content';
  app.append(tb.el, content);
  if (!readonly) {
    // pending typing becomes an edit first, so undo/redo act on what the user sees
    tb.button('', 'Undo (Ctrl+Z)', () => {
      view?.flush?.();
      post({ type: 'command', command: 'undo' });
    }, 'undo');
    tb.button('', 'Redo (Ctrl+Y)', () => {
      view?.flush?.();
      post({ type: 'command', command: 'redo' });
    }, 'redo');
    tb.button('', 'Save (Ctrl+S)', () => {
      view?.flush?.();
      post({ type: 'command', command: 'save' });
    }, 'save');
    tb.separator();
  }
  switch (model.type) {
    case 'sheet':
      view = renderSheetView(content, model, readonly, onEdit, tb);
      break;
    case 'text':
      view = renderTextView(content, model, readonly, onEdit, tb);
      break;
    case 'slides':
      view = renderSlidesView(content, model, tb);
      break;
    case 'pdf': {
      pdfWorker?.terminate();
      const chunk = (globalThis as { vsofxPdf?: PdfChunk }).vsofxPdf;
      if (!chunk) throw new Error('The PDF viewer is not loaded.');
      pdfWorker = blobWorker(chunk.workerSource, { type: 'module' });
      view = chunk.renderPdfView(content, model, readonly, onEdit, tb, pdfBytes ?? new Uint8Array(0), pdfWorker, (m) => post({ type: 'error', message: m }));
      break;
    }
  }
}

/**
 * Closes the current view without keeping its unsent input: the document is being reloaded
 * (revert, change on disk) or replaced, and an edit from the old view would land on the new content.
 */
let closing = false;
function closeView(): void {
  const old = view;
  view = undefined;
  closing = true;
  try {
    old?.dispose();
  } finally {
    closing = false;
  }
}

async function userEdit(edit: Edit): Promise<void> {
  if (closing || !view) return;
  try {
    const result = await rpc<ApplyResult>({ op: 'apply', edit });
    showEdit(edit, result);
    post({ type: 'edit', edit });
  } catch (err) {
    showError(`The edit could not be applied: ${message(err)}`);
    post({ type: 'error', message: message(err) });
  }
}

/** Passes an applied edit to the view; a batch part by part. */
function showEdit(edit: Edit, result: ApplyResult): void {
  if (edit.kind === 'batch') edit.edits.forEach((e, i) => showEdit(e, result.batch?.[i] ?? {}));
  else view?.applyEdit(edit, result);
}

function message(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

// ---- host messages -------------------------------------------------------

async function onInit(msg: Extract<HostMessage, { type: 'init' }>): Promise<void> {
  const gen = ++generation;
  loaded = false;
  queued = [];
  closeView();
  showLoading(msg.format);
  try {
    const bytes = ownedBytes(msg.bytes); // own buffer, transferred to the worker (no copy)
    // the PDF chunk carries the PDF format worker; the other families have their own chunk
    const family = workerFamily(msg.format);
    const pdf = family === 'pdf' ? await loadPdfChunk() : undefined;
    const source = pdf ? pdf.formatWorkerSource : await loadWorkerSource(family);
    if (gen !== generation) return;
    startWorker(source);
    pdfBytes = pdf ? bytes.slice() : undefined;
    const res = await rpc<LoadResult>({ op: 'load', format: msg.format, bytes, edits: msg.edits }, [bytes.buffer]);
    if (gen !== generation) return;
    render(res.model, msg.readonly || res.readonly);
    restoreView();
    loaded = true;
    const pendingEdits = queued;
    queued = [];
    for (const edit of pendingEdits) await hostEdit(edit);
  } catch (err) {
    if (gen !== generation) return;
    showError(message(err));
    post({ type: 'error', message: message(err) });
  }
}

/** Applies an edit that the host sends (undo, redo, mirrored edit); not echoed back. */
async function hostEdit(edit: Edit): Promise<void> {
  try {
    showEdit(edit, await rpc<ApplyResult>({ op: 'apply', edit }));
  } catch (err) {
    post({ type: 'error', message: message(err) });
  }
}

async function onSave(requestId: number): Promise<void> {
  if (!loaded) {
    post({ type: 'saveError', requestId, message: 'The document is not loaded yet.' });
    return;
  }
  try {
    view?.flush?.(); // its edit reaches the worker before the save request
    const bytes = await rpc<Uint8Array>({ op: 'save' });
    post({ type: 'saveResult', requestId, bytes }, [bytes.buffer]); // handed over, not copied
  } catch (err) {
    post({ type: 'saveError', requestId, message: message(err) });
  }
}

window.addEventListener('message', (ev: MessageEvent<HostMessage>) => {
  const msg = ev.data;
  switch (msg?.type) {
    case 'init':
      void onInit(msg);
      break;
    case 'apply':
      if (!loaded) queued.push(msg.edit);
      else void hostEdit(msg.edit);
      break;
    case 'requestSave':
      void onSave(msg.requestId);
      break;
    case 'probe':
      window.dispatchEvent(new CustomEvent('vsofx-probe')); // handled by dist/probe.js in test builds only
      break;
  }
});

// Every blocked request is reported; the host logs it (and tests assert on it).
window.addEventListener('securitypolicyviolation', (ev) => {
  post({ type: 'securityViolation', directive: ev.violatedDirective, blockedURI: ev.blockedURI });
});
window.addEventListener('vsofx-probe-result', (ev) => {
  post({ type: 'probeResult', results: (ev as CustomEvent<Record<string, string>>).detail });
});

post({ type: 'ready' });
