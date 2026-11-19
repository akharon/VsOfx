// Webview entry (UI thread). Third-party parsing is delegated to the worker.
import './style.css';
import workerBase64 from 'virtual:worker-bytes';
import pdfWorkerBase64 from 'virtual:pdf-worker-bytes';
import { fromByteArray, toByteArray } from '../shared/base64';
import type { HostMessage, WebviewMessage } from '../shared/messages';
import type { ApplyResult, DocModel, Edit } from '../shared/model';
import type { LoadResult, WorkerRequest, WorkerResponse } from './protocol';
import { createToolbar } from './ui/toolbar';
import { renderPdfView } from './views/pdfView';
import { renderSheetView } from './views/sheetView';
import { renderSlidesView } from './views/slidesView';
import { renderTextView } from './views/textView';
import type { View } from './views/types';

declare function acquireVsCodeApi(): { postMessage(msg: WebviewMessage, transfer?: Transferable[]): void };

const vscode = acquireVsCodeApi();
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

// ---- worker --------------------------------------------------------------

function blobWorker(base64: string, options?: WorkerOptions): Worker {
  const bin = atob(base64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  // blob: workers inherit this page's CSP, so connect-src 'none' applies to them as well.
  return new Worker(URL.createObjectURL(new Blob([bytes], { type: 'text/javascript' })), options);
}

function startWorker(): Worker {
  worker?.terminate();
  for (const p of pending.values()) p.reject(new Error('Worker restarted'));
  pending.clear();
  const w = blobWorker(workerBase64);
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

function showLoading(): void {
  app.replaceChildren();
  const d = document.createElement('div');
  d.className = 'loading';
  d.textContent = 'Loading…';
  app.appendChild(d);
}

function showError(message: string): void {
  view?.dispose();
  view = undefined;
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
  view?.dispose();
  app.replaceChildren();
  const onEdit = (edit: Edit) => void userEdit(edit);
  const tb = createToolbar();
  const content = document.createElement('div');
  content.className = 'content';
  app.append(tb.el, content);
  if (!readonly) {
    tb.button('↶ Undo', 'Undo (Ctrl+Z)', () => post({ type: 'command', command: 'undo' }));
    tb.button('↷ Redo', 'Redo (Ctrl+Y)', () => post({ type: 'command', command: 'redo' }));
    tb.button('Save', 'Save (Ctrl+S)', () => post({ type: 'command', command: 'save' }));
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
      pdfWorker = blobWorker(pdfWorkerBase64, { type: 'module' });
      view = renderPdfView(content, model, readonly, onEdit, tb, pdfBytes ?? new Uint8Array(0), pdfWorker, (m) => post({ type: 'error', message: m }));
      break;
    }
  }
}

async function userEdit(edit: Edit): Promise<void> {
  try {
    const result = await rpc<ApplyResult>({ op: 'apply', edit });
    view?.applyEdit(edit, result);
    post({ type: 'edit', edit });
  } catch (err) {
    showError(`The edit could not be applied: ${message(err)}`);
    post({ type: 'error', message: message(err) });
  }
}

function message(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

// ---- host messages -------------------------------------------------------

async function onInit(msg: Extract<HostMessage, { type: 'init' }>): Promise<void> {
  const gen = ++generation;
  loaded = false;
  queued = [];
  showLoading();
  try {
    startWorker();
    const bytes = toByteArray(msg.bytes); // own buffer, transferred to the worker
    pdfBytes = msg.format === 'pdf' ? bytes.slice() : undefined;
    const res = await rpc<LoadResult>({ op: 'load', format: msg.format, bytes, edits: msg.edits }, [bytes.buffer]);
    if (gen !== generation) return;
    render(res.model, msg.readonly || res.readonly);
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
    view?.applyEdit(edit, await rpc<ApplyResult>({ op: 'apply', edit }));
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
    const bytes = await rpc<Uint8Array>({ op: 'save' });
    post({ type: 'saveResult', requestId, bytes: fromByteArray(bytes) });
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
