import { applyTextEdit, diffText, type CodeModel, type Edit, type TextEdit } from '../../shared/model';
import { renderMarkdown, setHeading, toggleWrap, togglePrefix, type Selection } from '../code/markdown';
import { blockToMarkdown } from '../code/mdSerialize';
import { createFindBar, domFinder, type FindTarget } from '../ui/findBar';
import { addZoom, type IconName, type Toolbar } from '../ui/toolbar';
import type { EditHandler, View } from './types';

type Mode = 'visual' | 'source';

function loadMode(): Mode {
  try {
    return localStorage.getItem('vsofx.mode.markdown') === 'source' ? 'source' : 'visual';
  } catch {
    return 'visual';
  }
}
function saveMode(m: Mode): void {
  try {
    localStorage.setItem('vsofx.mode.markdown', m);
  } catch {
    /* storage unavailable */
  }
}

/** Line (1-based) of an offset. */
function lineOf(text: string, offset: number): number {
  let n = 1;
  for (let i = text.indexOf('\n'); i >= 0 && i < offset; i = text.indexOf('\n', i + 1)) n++;
  return n;
}

/** What a rendered top-level block looked like and which source lines [start, end) it came from. */
interface BlockInfo {
  html: string;
  start: number;
  end: number;
}

/**
 * Markdown, edited like a word processor: the rendered document is editable and every change
 * is written back to the Markdown text. Only the blocks the user touched are re-serialized, so
 * the rest of the file keeps its exact formatting. "Source" shows (and edits) the raw text.
 */
export function renderCodeView(container: HTMLElement, model: CodeModel, readonly: boolean, onEdit: EditHandler, tb: Toolbar): View {
  let text = model.text; // what the worker (and the host) have
  const own = new Set<Edit>();
  let mode: Mode = loadMode();

  container.replaceChildren();
  const layout = document.createElement('div');
  layout.className = 'code-layout lang-markdown';
  const visual = document.createElement('div');
  visual.className = 'code-visual md-preview';
  const body = document.createElement('div');
  body.className = 'md-body';
  body.setAttribute('contenteditable', readonly ? 'false' : 'true');
  body.spellcheck = true;
  body.setAttribute('role', 'textbox');
  body.setAttribute('aria-multiline', 'true');
  body.setAttribute('aria-label', 'Markdown document');
  visual.appendChild(body);
  const source = document.createElement('div');
  source.className = 'code-source';
  const ta = document.createElement('textarea');
  ta.className = 'src-text';
  ta.spellcheck = true;
  ta.wrap = 'soft';
  ta.readOnly = readonly;
  ta.value = text;
  ta.setAttribute('aria-label', 'Markdown source');
  source.append(ta);
  layout.append(visual, source);
  container.append(layout);
  try {
    document.execCommand('defaultParagraphSeparator', false, 'p');
  } catch {
    /* not supported */
  }

  // ---- toolbar ----
  const finder = createFindBar(tb.extra, combinedFinder());
  tb.button('', 'Find (Ctrl+F)', () => finder.toggle(), 'search');
  tb.separator();
  addZoom(tb, (z) => layout.style.setProperty('--zoom', String(z)));

  const linkBar = document.createElement('div');
  linkBar.className = 'findbar linkbar';
  linkBar.hidden = true;
  const linkInput = document.createElement('input');
  linkInput.type = 'text';
  linkInput.placeholder = 'https://example.com';
  linkInput.setAttribute('aria-label', 'Link address');
  linkBar.append(linkInput);
  tb.extra.append(linkBar);
  let savedRange: Range | undefined;

  if (!readonly) {
    tb.separator();
    const fmt = (title: string, ic: IconName, vis: () => void, src: (s: Selection) => Selection) =>
      tb.button('', title, () => (mode === 'visual' ? (vis(), body.focus()) : transformSelection(src)), ic);
    fmt('Bold (Ctrl+B)', 'bold', () => exec('bold'), (s) => toggleWrap(s, '**'));
    fmt('Italic (Ctrl+I)', 'italic', () => exec('italic'), (s) => toggleWrap(s, '_'));
    fmt('Inline code', 'code', toggleInlineCode, (s) => toggleWrap(s, '`', 'code'));
    fmt('Heading (repeat to change level)', 'heading', cycleHeading, (s) => {
      const line = s.text.slice(s.text.lastIndexOf('\n', s.start - 1) + 1);
      const lvl = /^(#{1,6})\s/.exec(line)?.[1]?.length ?? 0;
      return setHeading(s, lvl >= 3 ? 0 : lvl + 1);
    });
    fmt('Bulleted list', 'listBullet', () => exec('insertUnorderedList'), (s) => togglePrefix(s, '- ', /^\s*[-*+]\s+(\[[ xX]\]\s+)?/));
    fmt('Numbered list', 'listNumber', () => exec('insertOrderedList'), (s) => togglePrefix(s, (i) => `${i + 1}. `, /^\s*\d+[.)]\s+/));
    fmt('Task list', 'checkbox', toggleTask, (s) => togglePrefix(s, '- [ ] ', /^\s*[-*+]\s+\[[ xX]\]\s+/));
    fmt('Quote', 'quote', toggleQuote, (s) => togglePrefix(s, '> ', /^>\s?/));
    fmt('Link (Ctrl+K)', 'link', openLinkBar, (s) => {
      const label = s.text.slice(s.start, s.end) || 'link';
      const ins = `[${label}](https://)`;
      const at = s.start + label.length + 3;
      return { text: s.text.slice(0, s.start) + ins + s.text.slice(s.end), start: at, end: at + 8 };
    });
    fmt('Table', 'table', insertTable, (s) => {
      const t = '\n| Column 1 | Column 2 |\n| -------- | -------- |\n| Cell     | Cell     |\n';
      return { text: s.text.slice(0, s.end) + t + s.text.slice(s.end), start: s.end + 3, end: s.end + 11 };
    });
  }
  tb.spacer();
  const info = tb.info();
  const sourceBtn = tb.toggle('Source', 'Show and edit the Markdown source', (on) => setMode(on ? 'source' : 'visual'), 'code');

  // ---- shared: emitting edits ----
  function emit(e: TextEdit): void {
    own.add(e);
    onEdit(e);
  }
  function updateInfo(): void {
    const words = text.trim() ? text.trim().split(/\s+/).length : 0;
    info.textContent = `${words.toLocaleString()} words`;
  }

  // ---- visual (WYSIWYG) editing ----
  let blocks = new WeakMap<Node, BlockInfo>();
  /** Source lines before the first block / from the end of the last block (front matter, final newline). */
  let headEnd = 0;
  let tailStart = 0;
  let vTimer = 0;

  function renderVisual(): void {
    if (vTimer) clearTimeout(vTimer);
    vTimer = 0;
    updateInfo();
    const y = visual.scrollTop;
    renderMarkdown(text, body);
    blocks = new WeakMap();
    const lineCount = text.split('\n').length;
    headEnd = lineCount;
    tailStart = 0;
    for (const el of Array.from(body.children) as HTMLElement[]) {
      const start = Number(el.dataset['line']);
      const end = Number(el.dataset['end']);
      if (!Number.isInteger(start) || !Number.isInteger(end)) continue;
      blocks.set(el, { html: el.outerHTML, start, end });
      headEnd = Math.min(headEnd, start);
      tailStart = Math.max(tailStart, end);
    }
    if (headEnd === lineCount) {
      headEnd = 0; // no blocks: everything is "tail"
      tailStart = 0;
    }
    visual.scrollTop = y;
  }

  /** Writes the edited rendering back to the Markdown text (one minimal splice). */
  function syncVisual(): void {
    if (vTimer) clearTimeout(vTimer);
    vTimer = 0;
    if (mode !== 'visual' || readonly) return;
    const src = text.split('\n');
    const out: string[] = src.slice(0, headEnd);
    let prev: BlockInfo | undefined;
    let first = true;
    const placed: [Node, number, number][] = [];
    for (const n of Array.from(body.childNodes)) {
      if (n.nodeType === 8 || (n.nodeType === 3 && !(n.nodeValue ?? '').trim())) continue;
      const bi = blocks.get(n);
      const same = bi && n.nodeType === 1 && (n as Element).outerHTML === bi.html;
      const lines = same ? src.slice(bi.start, bi.end) : blockToMarkdown(n).split('\n');
      if (!same && lines.join('').trim() === '') continue; // empty paragraph
      if (!first) {
        if (same && prev && prev.end <= bi.start) out.push(...src.slice(prev.end, bi.start));
        else out.push('');
      }
      const at = out.length;
      out.push(...lines);
      placed.push([n, at, out.length]);
      prev = same ? bi : undefined;
      first = false;
    }
    // whatever followed the last block (usually the final newline) is kept
    const newTail = out.length;
    out.push(...src.slice(tailStart));
    const next = out.join('\n');
    // remember what each block now corresponds to
    for (const [n, s, end] of placed) blocks.set(n, { html: n.nodeType === 1 ? (n as Element).outerHTML : '', start: s, end });
    tailStart = newTail;
    const e = diffText(text, next);
    if (!e) return;
    text = next;
    ta.value = next;
    emit(e);
    updateInfo();
  }
  const scheduleSync = () => {
    if (vTimer) clearTimeout(vTimer);
    vTimer = window.setTimeout(syncVisual, 400);
  };

  body.addEventListener('input', scheduleSync);
  body.addEventListener('blur', () => syncVisual());
  body.addEventListener('beforeinput', (ev) => {
    // undo/redo belong to VS Code (the document's edit stack), not to the browser
    if (ev.inputType === 'historyUndo' || ev.inputType === 'historyRedo') ev.preventDefault();
  });
  body.addEventListener('paste', (ev) => {
    if (readonly) return;
    ev.preventDefault(); // plain text only: no foreign HTML/styles in the document
    const t = ev.clipboardData?.getData('text/plain') ?? '';
    if (t) document.execCommand('insertText', false, t);
  });
  body.addEventListener('keydown', (ev) => {
    if (readonly) return;
    const mod = ev.ctrlKey || ev.metaKey;
    if (ev.key === 'Enter' && !mod && blockOf(currentNode())?.tagName === 'PRE') {
      ev.preventDefault();
      document.execCommand('insertText', false, '\n');
    } else if (ev.key === 'Tab' && !mod) {
      const li = currentElement()?.closest('li');
      if (li) {
        ev.preventDefault();
        exec(ev.shiftKey ? 'outdent' : 'indent');
      }
    } else if (mod && !ev.altKey && ev.key.toLowerCase() === 'k') {
      ev.preventDefault();
      openLinkBar();
    }
  });

  function exec(cmd: string, value?: string): void {
    body.focus();
    document.execCommand(cmd, false, value);
    scheduleSync();
  }
  function currentNode(): Node | null {
    const s = window.getSelection();
    return s && s.rangeCount && body.contains(s.anchorNode) ? s.anchorNode : null;
  }
  function currentElement(): HTMLElement | null {
    const n = currentNode();
    return n ? (n.nodeType === 1 ? (n as HTMLElement) : n.parentElement) : null;
  }
  /** The top-level block containing a node. */
  function blockOf(n: Node | null): HTMLElement | null {
    while (n && n.parentNode !== body) n = n.parentNode;
    return n && n.nodeType === 1 ? (n as HTMLElement) : null;
  }
  function cycleHeading(): void {
    const el = currentElement()?.closest('h1,h2,h3,h4,h5,h6,p,div,li');
    const lvl = el && /^H[1-6]$/.test(el.tagName) ? Number(el.tagName[1]) : 0;
    exec('formatBlock', lvl >= 3 ? 'p' : `h${lvl + 1}`);
  }
  function toggleQuote(): void {
    exec('formatBlock', currentElement()?.closest('blockquote') ? 'p' : 'blockquote');
  }
  function toggleInlineCode(): void {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount || !body.contains(sel.anchorNode)) return;
    const code = currentElement()?.closest('code');
    if (code && !code.closest('pre')) {
      code.replaceWith(document.createTextNode(code.textContent ?? ''));
    } else {
      const r = sel.getRangeAt(0);
      const c = document.createElement('code');
      c.textContent = r.toString() || 'code';
      r.deleteContents();
      r.insertNode(c);
      sel.selectAllChildren(c);
    }
    scheduleSync();
  }
  function toggleTask(): void {
    let li = currentElement()?.closest('li');
    if (!li) {
      exec('insertUnorderedList');
      li = currentElement()?.closest('li');
    }
    if (!li) return;
    const box = li.querySelector(':scope > input.task, :scope > p > input.task');
    if (box) {
      box.remove();
      li.classList.remove('task-item');
    } else {
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.className = 'task';
      (li.firstElementChild?.tagName === 'P' ? li.firstElementChild : li).prepend(cb, document.createTextNode(' '));
      li.classList.add('task-item');
    }
    scheduleSync();
  }
  function insertTable(): void {
    exec(
      'insertHTML',
      '<table><thead><tr><th>Column 1</th><th>Column 2</th></tr></thead><tbody><tr><td>Cell</td><td>Cell</td></tr></tbody></table><p><br></p>',
    );
  }
  function openLinkBar(): void {
    if (mode !== 'visual') return;
    const sel = window.getSelection();
    savedRange = sel && sel.rangeCount && body.contains(sel.anchorNode) ? sel.getRangeAt(0).cloneRange() : undefined;
    linkInput.value = currentElement()?.closest('a')?.getAttribute('href') ?? 'https://';
    linkBar.hidden = false;
    linkInput.focus();
    linkInput.select();
  }
  linkInput.addEventListener('keydown', (ev) => {
    if (ev.key !== 'Enter' && ev.key !== 'Escape') return;
    ev.preventDefault();
    linkBar.hidden = true;
    body.focus();
    const sel = window.getSelection();
    if (savedRange && sel) {
      sel.removeAllRanges();
      sel.addRange(savedRange);
    }
    if (ev.key === 'Escape') return;
    const url = linkInput.value.trim();
    if (!url) exec('unlink');
    else if (/^\s*(javascript|vbscript|data):/i.test(url)) return;
    else if (sel && sel.isCollapsed) exec('insertHTML', `<a href="${escapeAttr(url)}">${escapeHtml(url)}</a>`);
    else exec('createLink', url);
  });
  linkInput.addEventListener('blur', () => (linkBar.hidden = true));

  // links do not navigate; anchors scroll; task checkboxes toggle
  visual.addEventListener('click', (ev) => {
    const t = ev.target as HTMLElement;
    const a = t.closest('a');
    if (a) {
      ev.preventDefault();
      const href = a.getAttribute('href') ?? '';
      if (href.startsWith('#') && (ev.ctrlKey || ev.metaKey || readonly)) {
        const id = decodeURIComponent(href.slice(1)).toLowerCase();
        const h = Array.from(body.querySelectorAll('h1,h2,h3,h4,h5,h6')).find((x) => slug(x.textContent ?? '') === id);
        h?.scrollIntoView({ block: 'start' });
      }
      return;
    }
    if (t instanceof HTMLInputElement && t.classList.contains('task')) {
      ev.preventDefault();
      if (readonly) return;
      if (t.hasAttribute('checked')) t.removeAttribute('checked');
      else t.setAttribute('checked', '');
      t.checked = t.hasAttribute('checked');
      syncVisual();
    }
  });

  // ---- source editor ----
  let timer = 0;
  /** Commits typing in the source pane as one edit. */
  function flushSource(): void {
    if (timer) clearTimeout(timer);
    timer = 0;
    const e = diffText(text, ta.value);
    if (!e) return;
    text = ta.value;
    emit(e);
    updateInfo();
  }
  function flush(): void {
    if (mode === 'visual') syncVisual();
    else flushSource();
  }
  /** Replaces the whole text (formatting buttons) as one minimal edit. */
  function replaceAll(next: string, sel?: { start: number; end: number }): void {
    const e = diffText(text, next);
    if (!e) return;
    const before = { s: ta.selectionStart, e: ta.selectionEnd };
    text = next;
    ta.value = next;
    if (sel) ta.setSelectionRange(sel.start, sel.end);
    else ta.setSelectionRange(mapPos(before.s, e), mapPos(before.e, e));
    emit(e);
    updateInfo();
  }
  function transformSelection(fn: (s: Selection) => Selection): void {
    flushSource();
    const r = fn({ text, start: ta.selectionStart, end: ta.selectionEnd });
    replaceAll(r.text, r);
    ta.focus();
  }

  ta.addEventListener('input', () => {
    if (timer) clearTimeout(timer);
    timer = window.setTimeout(flushSource, 350);
  });
  ta.addEventListener('blur', () => flushSource());
  ta.addEventListener('keydown', (ev) => {
    const mod = ev.ctrlKey || ev.metaKey;
    if (ev.key === 'Tab' && !mod && !readonly) {
      ev.preventDefault();
      if (ev.shiftKey) {
        transformSelection((s) => togglePrefix({ ...s }, '', /^ {2}/));
      } else if (ta.selectionStart !== ta.selectionEnd && ta.value.slice(ta.selectionStart, ta.selectionEnd).includes('\n')) {
        transformSelection((s) => {
          const ls = s.text.lastIndexOf('\n', s.start - 1) + 1;
          const block = s.text.slice(ls, s.end).replace(/^/gm, '  ');
          return { text: s.text.slice(0, ls) + block + s.text.slice(s.end), start: ls, end: ls + block.length };
        });
      } else {
        ta.setRangeText('  ', ta.selectionStart, ta.selectionEnd, 'end');
        ta.dispatchEvent(new Event('input'));
      }
    } else if (ev.key === 'Enter' && !mod && !readonly) {
      // keep indentation and continue lists
      const pos = ta.selectionStart;
      const line = ta.value.slice(ta.value.lastIndexOf('\n', pos - 1) + 1, pos);
      let indent = /^[ \t]*/.exec(line)![0];
      const li = /^([ \t]*)([-*+]|\d+[.)])(\s+\[[ xX]\])?\s+/.exec(line);
      if (li && line.trim() === li[0].trim()) {
        // empty list item: end the list
        ev.preventDefault();
        ta.setRangeText('', pos - line.length + li[1]!.length, pos, 'end');
        ta.dispatchEvent(new Event('input'));
        return;
      }
      if (li) {
        const n = /^\d+/.exec(li[2]!);
        indent = li[1]! + (n ? `${Number(n[0]) + 1}${li[2]!.slice(n[0].length)}` : li[2]) + (li[3] ? ' [ ]' : '') + ' ';
      }
      ev.preventDefault();
      ta.setRangeText(`\n${indent}`, ta.selectionStart, ta.selectionEnd, 'end');
      ta.dispatchEvent(new Event('input'));
    } else if (mod && !readonly && (ev.key === 'b' || ev.key === 'i' || ev.key === 'k')) {
      ev.preventDefault();
      if (ev.key === 'b') transformSelection((s) => toggleWrap(s, '**'));
      else if (ev.key === 'i') transformSelection((s) => toggleWrap(s, '_'));
      else (tb.el.querySelector('[title^="Link"]') as HTMLButtonElement | null)?.click();
    }
  });
  // undo/redo/save are VS Code commands: commit pending typing first so they see it
  const onCapture = (ev: KeyboardEvent) => {
    if ((ev.ctrlKey || ev.metaKey) && /^[zyse]$/i.test(ev.key)) flush();
    if ((ev.ctrlKey || ev.metaKey) && !ev.altKey && ev.key.toLowerCase() === 'f') {
      ev.preventDefault();
      finder.open();
    }
  };
  document.addEventListener('keydown', onCapture, true);

  // ---- modes ----
  function setMode(m: Mode): void {
    flush();
    mode = m;
    saveMode(m);
    layout.dataset['mode'] = m;
    sourceBtn.setAttribute('aria-pressed', String(m === 'source'));
    if (m === 'visual') renderVisual();
    else {
      ta.value = text;
      updateInfo();
    }
  }

  function combinedFinder(): FindTarget {
    const dom = domFinder(body);
    let hits: number[] = [];
    let len = 0;
    let useSource = true;
    return {
      find(q) {
        useSource = mode === 'source';
        if (!useSource) return dom.find(q);
        hits = [];
        len = q.length;
        const hay = ta.value.toLowerCase();
        const needle = q.toLowerCase();
        for (let i = hay.indexOf(needle); i >= 0 && needle; i = hay.indexOf(needle, i + needle.length)) hits.push(i);
        return hits.length;
      },
      reveal(i) {
        if (!useSource) return dom.reveal(i);
        const at = hits[i];
        if (at === undefined) return;
        const lh = parseFloat(getComputedStyle(ta).lineHeight) || 18;
        ta.setSelectionRange(at, at + len);
        ta.scrollTop = Math.max(0, (lineOf(ta.value, at) - 6) * lh);
      },
      clear() {
        dom.clear();
        hits = [];
      },
    };
  }

  layout.dataset['mode'] = mode;
  sourceBtn.setAttribute('aria-pressed', String(mode === 'source'));
  if (mode === 'visual') renderVisual();
  else updateInfo();

  return {
    applyEdit(edit) {
      if (edit.kind !== 'text') return;
      if (own.delete(edit)) return; // our own edit coming back
      // undo / redo / an edit in another editor of the same file
      if (mode === 'visual') {
        if (vTimer) clearTimeout(vTimer); // the DOM is re-rendered from the new text
        vTimer = 0;
        text = applyTextEdit(text, edit);
        ta.value = text;
        renderVisual();
        return;
      }
      // source mode: rebase pending typing
      const pending = diffText(text, ta.value);
      const selS = ta.selectionStart;
      const selE = ta.selectionEnd;
      text = applyTextEdit(text, edit);
      let value = text;
      if (pending) {
        const delta = edit.ins.length - edit.del.length;
        if (pending.at >= edit.at + edit.del.length) value = applyTextEdit(text, { ...pending, at: pending.at + delta });
        else if (pending.at + pending.del.length <= edit.at) value = applyTextEdit(text, pending);
      }
      ta.value = value;
      ta.setSelectionRange(mapPos(selS, edit), mapPos(selE, edit));
      updateInfo();
    },
    flush,
    dispose() {
      flush();
      if (vTimer) clearTimeout(vTimer);
      document.removeEventListener('keydown', onCapture, true);
      finder.close();
    },
  };
}

/** Moves a caret position across a splice. */
function mapPos(p: number, e: TextEdit): number {
  if (p <= e.at) return p;
  if (p >= e.at + e.del.length) return p + e.ins.length - e.del.length;
  return e.at + e.ins.length;
}

function slug(s: string): string {
  return s.trim().toLowerCase().replace(/[^\p{L}\p{N}\s-]/gu, '').replace(/\s+/g, '-');
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);
}
const escapeAttr = escapeHtml;
