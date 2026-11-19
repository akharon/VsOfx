import * as pdfjs from 'pdfjs-dist';
import type { PDFDocumentProxy, PDFPageProxy, PageViewport } from 'pdfjs-dist';
import type { ApplyResult, Edit, PdfModel } from '../../shared/model';
import {
  applyOp,
  stateFromJson,
  type FieldKind,
  type FieldSpec,
  type FieldValue,
  type PdfOp,
  type PdfState,
  type Rect,
  type TextSpec,
} from '../../shared/pdf';
import { createFindBar, domFinder } from '../ui/findBar';
import { addZoom, type Toolbar } from '../ui/toolbar';
import type { EditHandler, View } from './types';

/** A form control to draw, from an original widget or a field the user added. */
interface Widget {
  id: string;
  name: string;
  kind: 'text' | 'multiline' | 'checkbox' | 'radio' | 'dropdown' | 'listbox' | 'button' | 'signature';
  rect: Rect;
  readOnly: boolean;
  options: string[];
  radioValue?: string;
  fontSize?: number;
  maxLen?: number;
  added: boolean;
}

interface PageRec {
  index: number;
  proxy: PDFPageProxy;
  wrap: HTMLElement;
  canvas: HTMLCanvasElement;
  textLayer: HTMLElement;
  overlay: HTMLElement;
  deletedBar: HTMLElement;
  viewport: PageViewport;
  widgets: Widget[]; // original widgets
  rendered: boolean;
  textDone: boolean;
  task?: { cancel(): void };
}

type Placing = 'text' | 'checkbox' | 'dropdown' | 'freetext' | null;

const SIZES: Record<Exclude<Placing, null>, [number, number]> = {
  text: [150, 22],
  checkbox: [16, 16],
  dropdown: [130, 22],
  freetext: [150, 18],
};

/**
 * PDF view. pdf.js (own worker, no network) draws pages and text; form controls are our own HTML
 * elements placed over the widget rectangles. All changes are PdfOps (shared/pdf.ts); the file is
 * written by the worker adapter with pdf-lib.
 */
export function renderPdfView(
  container: HTMLElement,
  model: PdfModel,
  readonly: boolean,
  onEdit: EditHandler,
  tb: Toolbar,
  bytes: Uint8Array,
  pdfWorker: Worker,
  onFatal: (message: string) => void = () => {},
): View {
  container.replaceChildren();
  const state: PdfState = stateFromJson(model.state);
  const initial = new Map<string, FieldValue>();
  const names = new Set<string>();
  const recs: PageRec[] = [];
  let doc: PDFDocumentProxy | undefined;
  let loadingTask: { destroy(): Promise<void> } | undefined;
  let disposed = false;
  let baseScale = 1.25;
  let zoom = 1;
  let designMode = false;
  let placing: Placing = null;
  let selected: string | null = null;
  let currentPage = 0;
  let textGen = 0;
  let nextId = 1;
  for (const k of [...state.added.keys(), ...state.texts.keys()]) nextId = Math.max(nextId, (Number(k.slice(1)) || 0) + 1);

  const scroller = document.createElement('div');
  scroller.className = 'pdf-scroll';
  const pagesEl = document.createElement('div');
  pagesEl.className = 'pdf-pages';
  const panel = document.createElement('div');
  panel.className = 'pdf-panel';
  panel.hidden = true;
  const status = document.createElement('div');
  status.className = 'loading';
  status.textContent = 'Loading PDF…';
  scroller.append(pagesEl);
  container.append(panel, scroller, status);

  const emit = (op: PdfOp) => onEdit({ kind: 'pdf', op });
  const scale = () => baseScale * zoom;
  const current = (name: string): FieldValue => state.values.get(name) ?? initial.get(name) ?? '';
  const rectOf = (id: string, fallback: Rect): Rect => state.rects.get(id) ?? fallback;

  // ---- toolbar ----
  const info = tb.info();
  const finder = createFindBar(tb.extra, domFinder(pagesEl));
  tb.button('Find', 'Find in document (Ctrl+F)', () => finder.toggle());
  addZoom(tb, (z) => {
    zoom = z;
    layoutAll();
  });
  tb.button('◀', 'Previous page', () => goTo(currentPage - 1));
  tb.button('▶', 'Next page', () => goTo(currentPage + 1));
  if (!readonly) {
    tb.separator();
    const design = tb.toggle('Design', 'Edit the form: select, move, resize and delete fields', (on) => setDesign(on));
    const place = (label: string, title: string, kind: Exclude<Placing, null>) =>
      tb.button(label, title, () => {
        placing = kind;
        if (!designMode) design.click();
        scroller.classList.add('placing');
        info.textContent = 'Click on the page to place it';
      });
    place('＋ Text field', 'Add a text field, then click on the page', 'text');
    place('＋ Checkbox', 'Add a checkbox, then click on the page', 'checkbox');
    place('＋ Dropdown', 'Add a dropdown list, then click on the page', 'dropdown');
    place('＋ Text', 'Add fixed text to the page (printed into the page)', 'freetext');
    tb.separator();
    tb.button('Rotate page', 'Rotate the current page by 90°', () => {
      const old = state.rotations.get(currentPage) ?? 0;
      emit({ op: 'setRotation', page: currentPage, deg: (old + 90) % 360, old });
    });
    tb.button('Delete page', 'Delete the current page (on save)', () => {
      if (state.deletedPages.size + 1 >= recs.length) return;
      emit({ op: 'deletePage', page: currentPage });
    });
  }

  function setDesign(on: boolean): void {
    designMode = on;
    scroller.classList.toggle('design', on);
    if (!on) {
      placing = null;
      scroller.classList.remove('placing');
      select(null);
      updateInfo();
    }
    for (const r of recs) buildOverlay(r);
  }

  // ---- loading ----
  const observer =
    typeof IntersectionObserver === 'undefined'
      ? undefined
      : new IntersectionObserver(
          (entries) => {
            for (const e of entries) {
              if (!e.isIntersecting) continue;
              const rec = recs[Number((e.target as HTMLElement).dataset['page'])];
              if (rec) void renderPage(rec);
            }
          },
          { root: scroller, rootMargin: '800px 0px' },
        );

  void (async () => {
    try {
      pdfjs.GlobalWorkerOptions.workerPort = pdfWorker;
      const task = pdfjs.getDocument({
        data: bytes,
        useWorkerFetch: false,
        disableAutoFetch: true,
        enableXfa: false,
        useSystemFonts: true,
      });
      loadingTask = task;
      doc = await task.promise;
      if (disposed) return;
      const first = await doc.getPage(1);
      const fit = (scroller.clientWidth - 40) / first.getViewport({ scale: 1 }).width;
      if (Number.isFinite(fit) && fit > 0.3) baseScale = Math.min(2, Math.max(1, fit));
      for (let i = 0; i < doc.numPages; i++) {
        const proxy = i === 0 ? first : await doc.getPage(i + 1);
        if (disposed) return;
        const rec = createPage(i, proxy);
        recs.push(rec);
        pagesEl.appendChild(rec.wrap);
        const annots = await proxy.getAnnotations({ intent: 'display' });
        rec.widgets = annots.filter((a) => a.subtype === 'Widget').map((a) => toWidget(a, i));
        for (const w of rec.widgets) {
          names.add(w.name);
          if (!initial.has(w.name) && w.kind !== 'radio') initial.set(w.name, w.kind === 'checkbox' ? w.radioValue === 'on' : String(w.options[0] ?? ''));
        }
        layoutPage(rec);
        buildOverlay(rec);
        observer?.observe(rec.wrap);
      }
      status.remove();
      updateInfo();
      void buildAllText();
    } catch (err) {
      status.className = 'error';
      const msg = err instanceof Error ? err.message : String(err);
      status.textContent = `Cannot display this PDF: ${msg}`;
      onFatal(`Cannot display this PDF: ${msg}`);
    }
  })();

  /** Reads a pdf.js widget annotation into our own description. */
  function toWidget(a: Record<string, unknown>, _page: number): Widget {
    const type = a['fieldType'];
    const opts = (Array.isArray(a['options']) ? a['options'] : []) as { exportValue?: string; displayValue?: string }[];
    const options = opts.map((o) => String(o.displayValue ?? o.exportValue ?? ''));
    const val = Array.isArray(a['fieldValue']) ? String((a['fieldValue'] as unknown[])[0] ?? '') : String(a['fieldValue'] ?? '');
    const da = a['defaultAppearanceData'] as { fontSize?: number } | undefined;
    const w: Widget = {
      id: String(a['id']),
      name: String(a['fieldName'] ?? ''),
      kind: 'text',
      rect: a['rect'] as Rect,
      readOnly: Boolean(a['readOnly']),
      options,
      fontSize: da?.fontSize && da.fontSize > 0 ? da.fontSize : undefined,
      maxLen: typeof a['maxLen'] === 'number' ? (a['maxLen'] as number) : undefined,
      added: false,
    };
    if (type === 'Tx') {
      w.kind = a['multiLine'] ? 'multiline' : 'text';
      initial.set(w.name, val);
    } else if (type === 'Btn') {
      if (a['pushButton']) w.kind = 'button';
      else if (a['radioButton']) {
        w.kind = 'radio';
        w.radioValue = String(a['buttonValue'] ?? '');
        if (!initial.has(w.name)) initial.set(w.name, val === 'Off' ? '' : val);
      } else {
        w.kind = 'checkbox';
        const exp = String(a['exportValue'] ?? 'Yes');
        w.radioValue = val && val !== 'Off' && (val === exp || val === 'true' || val === 'Yes' || val === 'On' || val.length > 0) ? 'on' : 'off';
        initial.set(w.name, w.radioValue === 'on');
      }
    } else if (type === 'Ch') {
      w.kind = a['combo'] ? 'dropdown' : 'listbox';
      const idx = opts.findIndex((o) => o.exportValue === val || o.displayValue === val);
      initial.set(w.name, idx >= 0 ? options[idx]! : val);
    } else if (type === 'Sig') {
      w.kind = 'signature';
    }
    return w;
  }

  // ---- pages ----
  function createPage(index: number, proxy: PDFPageProxy): PageRec {
    const wrap = document.createElement('section');
    wrap.className = 'pdf-page';
    wrap.dataset['page'] = String(index);
    wrap.setAttribute('aria-label', `Page ${index + 1}`);
    const canvas = document.createElement('canvas');
    const textLayer = document.createElement('div');
    textLayer.className = 'textLayer';
    const overlay = document.createElement('div');
    overlay.className = 'pdf-overlay';
    const deletedBar = document.createElement('div');
    deletedBar.className = 'pdf-deleted';
    deletedBar.hidden = true;
    wrap.append(canvas, textLayer, overlay, deletedBar);
    const rec: PageRec = {
      index,
      proxy,
      wrap,
      canvas,
      textLayer,
      overlay,
      deletedBar,
      viewport: proxy.getViewport({ scale: 1 }),
      widgets: [],
      rendered: false,
      textDone: false,
    };
    overlay.addEventListener('pointerdown', (ev) => onPlace(rec, ev));
    return rec;
  }

  function layoutPage(rec: PageRec): void {
    const rotation = (rec.proxy.rotate + (state.rotations.get(rec.index) ?? 0)) % 360;
    rec.viewport = rec.proxy.getViewport({ scale: scale(), rotation });
    const { width, height } = rec.viewport;
    rec.wrap.style.width = `${Math.floor(width)}px`;
    rec.wrap.style.height = `${Math.floor(height)}px`;
    rec.wrap.style.setProperty('--total-scale-factor', String(scale()));
    rec.task?.cancel();
    rec.rendered = false;
    rec.textDone = false;
    rec.textLayer.replaceChildren();
    const deleted = state.deletedPages.has(rec.index);
    rec.wrap.classList.toggle('deleted', deleted);
    rec.deletedBar.hidden = !deleted;
    if (deleted) {
      rec.deletedBar.replaceChildren();
      const t = document.createElement('span');
      t.textContent = `Page ${rec.index + 1} will be removed when you save. `;
      rec.deletedBar.appendChild(t);
      if (!readonly) {
        const b = document.createElement('button');
        b.type = 'button';
        b.textContent = 'Restore';
        b.addEventListener('click', () => emit({ op: 'restorePage', page: rec.index }));
        rec.deletedBar.appendChild(b);
      }
    }
  }

  function layoutAll(): void {
    for (const r of recs) {
      layoutPage(r);
      buildOverlay(r);
    }
    renderVisible();
    void buildAllText();
  }

  function renderVisible(): void {
    const box = scroller.getBoundingClientRect();
    for (const r of recs) {
      const b = r.wrap.getBoundingClientRect();
      if (b.bottom > box.top - 800 && b.top < box.bottom + 800) void renderPage(r);
    }
  }

  async function renderPage(rec: PageRec): Promise<void> {
    if (rec.rendered || state.deletedPages.has(rec.index)) return;
    rec.rendered = true;
    const dpr = Math.min(2, window.devicePixelRatio || 1);
    const vp = rec.viewport;
    rec.canvas.width = Math.floor(vp.width * dpr);
    rec.canvas.height = Math.floor(vp.height * dpr);
    rec.canvas.style.width = `${Math.floor(vp.width)}px`;
    rec.canvas.style.height = `${Math.floor(vp.height)}px`;
    try {
      const task = rec.proxy.render({
        canvas: rec.canvas,
        viewport: vp,
        transform: dpr === 1 ? undefined : [dpr, 0, 0, dpr, 0, 0],
        annotationMode: pdfjs.AnnotationMode.ENABLE_FORMS, // form widgets are drawn by our overlay
      });
      rec.task = task;
      await task.promise;
    } catch {
      rec.rendered = false; // cancelled by a re-layout
      return;
    }
    await buildText(rec);
  }

  async function buildText(rec: PageRec): Promise<void> {
    if (rec.textDone || state.deletedPages.has(rec.index)) return;
    rec.textDone = true;
    try {
      rec.textLayer.replaceChildren();
      const layer = new pdfjs.TextLayer({
        textContentSource: rec.proxy.streamTextContent(),
        container: rec.textLayer,
        viewport: rec.viewport,
      });
      await layer.render();
    } catch {
      rec.textDone = false;
    }
  }

  /** Builds the text layers of all pages in the background so that Find covers the whole document. */
  async function buildAllText(): Promise<void> {
    const gen = ++textGen;
    for (const r of recs) {
      if (disposed || gen !== textGen) return;
      await buildText(r);
    }
  }

  function goTo(i: number): void {
    const r = recs[Math.max(0, Math.min(recs.length - 1, i))];
    r?.wrap.scrollIntoView?.({ block: 'start' });
  }

  function updateInfo(): void {
    if (placing) return;
    const total = recs.length;
    info.textContent = total ? `Page ${currentPage + 1} of ${total}` : '';
  }

  let raf = 0;
  scroller.addEventListener('scroll', () => {
    if (raf) return;
    raf = requestAnimationFrame(() => {
      raf = 0;
      const mid = scroller.getBoundingClientRect().top + scroller.clientHeight / 3;
      let best = currentPage;
      for (const r of recs) {
        const b = r.wrap.getBoundingClientRect();
        if (b.top <= mid) best = r.index;
        else break;
      }
      if (best !== currentPage) {
        currentPage = best;
        updateInfo();
      }
    });
  });

  // ---- overlays: form controls, added fields, text boxes ----
  function placeEl(el: HTMLElement, rec: PageRec, rect: Rect): void {
    const [x1, y1] = rec.viewport.convertToViewportPoint(rect[0], rect[1]) as [number, number];
    const [x2, y2] = rec.viewport.convertToViewportPoint(rect[2], rect[3]) as [number, number];
    el.style.left = `${Math.min(x1, x2)}px`;
    el.style.top = `${Math.min(y1, y2)}px`;
    el.style.width = `${Math.abs(x2 - x1)}px`;
    el.style.height = `${Math.abs(y2 - y1)}px`;
  }

  function widgetsOf(rec: PageRec): Widget[] {
    const list = rec.widgets.filter((w) => !state.removed.has(w.id)).map((w) => ({ ...w, rect: rectOf(w.id, w.rect) }));
    for (const s of state.added.values()) {
      if (s.page !== rec.index) continue;
      names.add(s.name);
      list.push({
        id: s.id,
        name: s.name,
        kind: s.kind === 'text' ? (s.multiline ? 'multiline' : 'text') : s.kind,
        rect: rectOf(s.id, s.rect),
        readOnly: false,
        options: s.options ?? [],
        added: true,
      });
    }
    return list;
  }

  function buildOverlay(rec: PageRec): void {
    rec.overlay.replaceChildren();
    if (state.deletedPages.has(rec.index)) return;
    for (const w of widgetsOf(rec)) rec.overlay.appendChild(makeWidget(rec, w));
    for (const t of state.texts.values()) if (t.page === rec.index) rec.overlay.appendChild(makeTextBox(rec, t));
  }

  function makeWidget(rec: PageRec, w: Widget): HTMLElement {
    const box = document.createElement('div');
    box.className = `pdf-w pdf-${w.kind}`;
    box.dataset['id'] = w.id;
    box.dataset['name'] = w.name;
    placeEl(box, rec, w.rect);
    const ro = readonly || w.readOnly;
    const fontPx = (w.fontSize ?? 0) > 0 ? (w.fontSize as number) * scale() : Math.max(9, Math.min(16, Math.abs(box.offsetHeight || 20) * 0.6));
    const label = `${w.name}`;
    let ctl: HTMLElement;
    switch (w.kind) {
      case 'text':
      case 'multiline': {
        const el: HTMLInputElement | HTMLTextAreaElement = w.kind === 'multiline' ? document.createElement('textarea') : document.createElement('input');
        if (el instanceof HTMLInputElement) el.type = 'text';
        el.value = String(current(w.name));
        if (w.maxLen) el.maxLength = w.maxLen;
        el.spellcheck = false;
        el.readOnly = ro;
        el.style.fontSize = `${Math.min(fontPx, 40)}px`;
        el.addEventListener('change', () => commitValue(w.name, el.value));
        (el as HTMLElement).addEventListener('keydown', (ev: KeyboardEvent) => {
          if (ev.key === 'Escape') {
            el.value = String(current(w.name));
            el.blur();
          } else if (ev.key === 'Enter' && w.kind === 'text') el.blur();
        });
        ctl = el;
        break;
      }
      case 'checkbox': {
        const el = document.createElement('input');
        el.type = 'checkbox';
        el.checked = current(w.name) === true;
        el.disabled = ro;
        el.addEventListener('change', () => commitValue(w.name, el.checked));
        ctl = el;
        break;
      }
      case 'radio': {
        const el = document.createElement('input');
        el.type = 'radio';
        el.checked = current(w.name) === w.radioValue;
        el.disabled = ro;
        el.addEventListener('change', () => el.checked && commitValue(w.name, w.radioValue ?? ''));
        ctl = el;
        break;
      }
      case 'dropdown':
      case 'listbox': {
        const el = document.createElement('select');
        if (w.kind === 'listbox') el.size = Math.max(2, Math.min(w.options.length, 6));
        const blank = document.createElement('option');
        blank.value = '';
        blank.textContent = '';
        if (w.kind === 'dropdown') el.appendChild(blank);
        for (const o of w.options) {
          const opt = document.createElement('option');
          opt.value = o;
          opt.textContent = o;
          el.appendChild(opt);
        }
        el.value = String(current(w.name));
        el.disabled = ro;
        el.style.fontSize = `${Math.min(fontPx, 40)}px`;
        el.addEventListener('change', () => commitValue(w.name, el.value));
        ctl = el;
        break;
      }
      case 'button':
      case 'signature': {
        const el = document.createElement('div');
        el.className = 'pdf-static';
        el.textContent = w.kind === 'signature' ? 'Signature' : '';
        ctl = el;
        break;
      }
    }
    ctl.setAttribute('aria-label', label);
    ctl.title = label;
    box.appendChild(ctl);
    if (!readonly) {
      box.addEventListener('pointerdown', (ev) => onDesignDown(rec, w.id, box, ev));
      if (selected === w.id) box.classList.add('selected');
      const h = document.createElement('span');
      h.className = 'pdf-handle';
      h.addEventListener('pointerdown', (ev) => onDesignDown(rec, w.id, box, ev, true));
      box.appendChild(h);
    }
    return box;
  }

  function makeTextBox(rec: PageRec, t: TextSpec): HTMLElement {
    const box = document.createElement('div');
    box.className = 'pdf-w pdf-freetext';
    box.dataset['id'] = t.id;
    placeEl(box, rec, rectOf(t.id, t.rect));
    const el = document.createElement('div');
    el.className = 'pdf-ftext';
    el.textContent = t.text;
    el.style.fontSize = `${t.size * scale()}px`;
    if (!readonly) {
      el.contentEditable = 'plaintext-only';
      el.spellcheck = false;
      el.addEventListener('blur', () => {
        const text = (el.textContent ?? '').trim() || ' ';
        const old = state.texts.get(t.id)?.text ?? '';
        if (text !== old) emit({ op: 'setText', id: t.id, text, old });
      });
      box.addEventListener('pointerdown', (ev) => onDesignDown(rec, t.id, box, ev));
      const h = document.createElement('span');
      h.className = 'pdf-handle';
      h.addEventListener('pointerdown', (ev) => onDesignDown(rec, t.id, box, ev, true));
      box.append(el, h);
      if (selected === t.id) box.classList.add('selected');
    } else box.appendChild(el);
    return box;
  }

  function commitValue(name: string, value: FieldValue): void {
    const old = current(name);
    if (old === value) return;
    emit({ op: 'setField', name, value, old });
  }

  /** Brings every control of a field in line with the state (several widgets can share a field). */
  function syncField(name: string): void {
    const v = current(name);
    for (const box of Array.from(pagesEl.querySelectorAll<HTMLElement>('.pdf-w'))) {
      if (box.dataset['name'] !== name) continue;
      const ctl = box.firstElementChild;
      if (ctl instanceof HTMLInputElement && ctl.type === 'checkbox') ctl.checked = v === true;
      else if (ctl instanceof HTMLInputElement && ctl.type === 'radio') {
        const w = recs.flatMap((r) => r.widgets).find((x) => x.id === box.dataset['id']);
        ctl.checked = v === (w?.radioValue ?? '');
      } else if (ctl instanceof HTMLInputElement || ctl instanceof HTMLTextAreaElement || ctl instanceof HTMLSelectElement) {
        if (ctl !== document.activeElement || ctl.value !== String(v)) ctl.value = String(v);
      }
    }
  }

  // ---- design: select, move, resize, delete, place ----
  function select(id: string | null): void {
    selected = id;
    for (const b of Array.from(pagesEl.querySelectorAll<HTMLElement>('.pdf-w'))) b.classList.toggle('selected', b.dataset['id'] === id);
    renderPanel();
  }

  function onDesignDown(rec: PageRec, id: string, box: HTMLElement, ev: PointerEvent, resize = false): void {
    if (!designMode || ev.button !== 0) return;
    ev.preventDefault();
    ev.stopPropagation();
    select(id);
    const startX = ev.clientX;
    const startY = ev.clientY;
    const l = box.offsetLeft;
    const t = box.offsetTop;
    const w = box.offsetWidth;
    const h = box.offsetHeight;
    const startRect = currentRect(id);
    const move = (e: PointerEvent) => {
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;
      if (resize) {
        box.style.width = `${Math.max(8, w + dx)}px`;
        box.style.height = `${Math.max(8, h + dy)}px`;
      } else {
        box.style.left = `${l + dx}px`;
        box.style.top = `${t + dy}px`;
      }
    };
    const up = () => {
      document.removeEventListener('pointermove', move);
      document.removeEventListener('pointerup', up);
      const a = rec.viewport.convertToPdfPoint(box.offsetLeft, box.offsetTop);
      const b = rec.viewport.convertToPdfPoint(box.offsetLeft + box.offsetWidth, box.offsetTop + box.offsetHeight);
      const rect: Rect = [Math.min(a[0]!, b[0]!), Math.min(a[1]!, b[1]!), Math.max(a[0]!, b[0]!), Math.max(a[1]!, b[1]!)];
      if (startRect && rect.some((v, i) => Math.abs(v - startRect[i]!) > 0.5)) emit({ op: 'setRect', id, rect, old: startRect });
    };
    document.addEventListener('pointermove', move);
    document.addEventListener('pointerup', up);
  }

  function currentRect(id: string): Rect | undefined {
    const added = state.added.get(id);
    if (added) return rectOf(id, added.rect);
    const text = state.texts.get(id);
    if (text) return rectOf(id, text.rect);
    for (const r of recs) {
      const w = r.widgets.find((x) => x.id === id);
      if (w) return rectOf(id, w.rect);
    }
    return undefined;
  }

  function onPlace(rec: PageRec, ev: PointerEvent): void {
    if (!placing || ev.target !== rec.overlay) {
      if (designMode && ev.target === rec.overlay) select(null);
      return;
    }
    const kind = placing;
    placing = null;
    scroller.classList.remove('placing');
    updateInfo();
    const bounds = rec.overlay.getBoundingClientRect();
    const [px, py] = rec.viewport.convertToPdfPoint(ev.clientX - bounds.left, ev.clientY - bounds.top) as [number, number];
    const [w, h] = SIZES[kind];
    const rect: Rect = [px - w / 2, py - h / 2, px + w / 2, py + h / 2];
    const id = `${kind === 'freetext' ? 't' : 'n'}${nextId++}`;
    if (kind === 'freetext') {
      emit({ op: 'addText', spec: { id, page: rec.index, rect, text: 'Text', size: 12 } });
    } else {
      const base = kind === 'text' ? 'Text' : kind === 'checkbox' ? 'Check' : 'Dropdown';
      let n = 1;
      while (names.has(`${base}${n}`)) n++;
      const spec: FieldSpec = {
        id,
        kind: kind as FieldKind,
        name: `${base}${n}`,
        page: rec.index,
        rect,
        ...(kind === 'dropdown' ? { options: ['Option 1', 'Option 2', 'Option 3'] } : {}),
      };
      names.add(spec.name);
      emit({ op: 'addField', spec });
    }
    select(id);
  }

  function renderPanel(): void {
    panel.replaceChildren();
    panel.hidden = !(designMode && selected);
    if (panel.hidden || !selected) return;
    const id = selected;
    const spec = state.added.get(id);
    const text = state.texts.get(id);
    const orig = recs.flatMap((r) => r.widgets).find((w) => w.id === id);
    const label = (t: string, input: HTMLElement) => {
      const l = document.createElement('label');
      l.append(`${t} `, input);
      panel.appendChild(l);
    };
    const btn = (t: string, fn: () => void) => {
      const b = document.createElement('button');
      b.type = 'button';
      b.textContent = t;
      b.addEventListener('click', fn);
      panel.appendChild(b);
      return b;
    };
    if (spec) {
      const name = document.createElement('input');
      name.value = spec.name;
      label('Name', name);
      const opts = document.createElement('input');
      opts.value = (spec.options ?? []).join(', ');
      if (spec.kind === 'dropdown') label('Options (comma separated)', opts);
      btn('Apply', () => {
        const nm = name.value.trim();
        if (!nm) return;
        const options = opts.value.split(',').map((s) => s.trim()).filter(Boolean);
        const next: FieldSpec = { ...spec, name: nm, rect: rectOf(id, spec.rect), ...(spec.kind === 'dropdown' ? { options } : {}) };
        emit({ op: 'dropField', spec: { ...spec, rect: rectOf(id, spec.rect) } });
        emit({ op: 'addField', spec: next });
      });
    } else if (orig) {
      const s = document.createElement('span');
      s.textContent = `Field "${orig.name}" (${orig.kind})`;
      panel.appendChild(s);
    } else if (text) {
      const size = document.createElement('input');
      size.type = 'number';
      size.min = '4';
      size.max = '96';
      size.value = String(text.size);
      label('Size', size);
      btn('Apply', () => {
        const n = Number(size.value);
        if (n >= 4 && n <= 96 && n !== text.size) {
          emit({ op: 'dropText', spec: { ...text, rect: rectOf(id, text.rect) } });
          emit({ op: 'addText', spec: { ...text, rect: rectOf(id, text.rect), size: n } });
        }
      });
    }
    btn('Delete', () => deleteSelected());
  }

  function deleteSelected(): void {
    const id = selected;
    if (!id) return;
    const spec = state.added.get(id);
    const text = state.texts.get(id);
    const orig = recs.flatMap((r) => r.widgets).find((w) => w.id === id);
    if (spec) emit({ op: 'dropField', spec: { ...spec, rect: rectOf(id, spec.rect) } });
    else if (text) emit({ op: 'dropText', spec: { ...text, rect: rectOf(id, text.rect) } });
    else if (orig) emit({ op: 'removeWidget', id, name: orig.name });
    select(null);
  }

  const onKey = (ev: KeyboardEvent) => {
    if ((ev.ctrlKey || ev.metaKey) && !ev.altKey && ev.key.toLowerCase() === 'f') {
      ev.preventDefault();
      finder.open();
    } else if (designMode && selected && (ev.key === 'Delete' || ev.key === 'Backspace')) {
      const t = ev.target as HTMLElement;
      if (t.closest('input, textarea, select, [contenteditable]')) return;
      ev.preventDefault();
      deleteSelected();
    } else if (ev.key === 'Escape' && placing) {
      placing = null;
      scroller.classList.remove('placing');
      updateInfo();
    }
  };
  document.addEventListener('keydown', onKey);

  return {
    applyEdit(edit: Edit, _result: ApplyResult) {
      if (edit.kind !== 'pdf') return;
      const op = edit.op;
      applyOp(state, op);
      switch (op.op) {
        case 'setField':
          syncField(op.name);
          break;
        case 'setText':
          break;
        case 'setRotation':
        case 'deletePage':
        case 'restorePage': {
          const rec = recs[op.page];
          if (rec) {
            layoutPage(rec);
            buildOverlay(rec);
            renderVisible();
            void buildAllText();
          }
          break;
        }
        default:
          for (const r of recs) buildOverlay(r);
          renderPanel();
      }
    },
    dispose() {
      disposed = true;
      textGen++;
      observer?.disconnect();
      document.removeEventListener('keydown', onKey);
      finder.close();
      for (const r of recs) r.task?.cancel();
      void loadingTask?.destroy();
    },
  };
}
