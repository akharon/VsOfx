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

  let alignHBtn: HTMLButtonElement | undefined;
  let tableBar: HTMLDivElement | undefined;
  let rowsIn!: HTMLInputElement;
  let colsIn!: HTMLInputElement;
  let tableHint!: HTMLSpanElement;
  let tableRange: Range | undefined;
  let tableSel: { start: number; end: number } | undefined;
  if (!readonly) {
    tb.separator();
    const fmt = (title: string, ic: IconName, vis: () => void, src: (s: Selection) => Selection, keepFocus = true) => {
      const b = tb.button('', title, () => (mode === 'visual' ? (vis(), keepFocus && body.focus()) : transformSelection(src)), ic);
      // the click must not take the cursor / selection out of the document
      b.addEventListener('mousedown', (ev) => ev.preventDefault());
      return b;
    };
    fmt('Bold (Ctrl+B)', 'bold', () => exec('bold'), (s) => toggleWrap(s, '**'));
    fmt('Italic (Ctrl+I)', 'italic', () => exec('italic'), (s) => toggleWrap(s, '_'));
    fmt('Inline code', 'code', toggleInlineCode, (s) => toggleWrap(s, '`', 'code'));
    fmt('Heading (repeat to change level)', 'heading', cycleHeading, (s) => {
      const line = s.text.slice(s.text.lastIndexOf('\n', s.start - 1) + 1);
      const lvl = /^(#{1,6})\s/.exec(line)?.[1]?.length ?? 0;
      return setHeading(s, lvl >= 3 ? 0 : lvl + 1);
    });
    fmt('Bulleted list', 'listBullet', () => toggleList('UL'), (s) => togglePrefix(s, '- ', /^\s*[-*+]\s+(\[[ xX]\]\s+)?/));
    fmt('Numbered list', 'listNumber', () => toggleList('OL'), (s) => togglePrefix(s, (i) => `${i + 1}. `, /^\s*\d+[.)]\s+/));
    fmt('Task list', 'checkbox', toggleTask, (s) => togglePrefix(s, '- [ ] ', /^\s*[-*+]\s+\[[ xX]\]\s+/));
    fmt('Quote', 'quote', toggleQuote, (s) => togglePrefix(s, '> ', /^>\s?/));
    fmt('Link (Ctrl+K)', 'link', openLinkBar, (s) => {
      const label = s.text.slice(s.start, s.end) || 'link';
      const ins = `[${label}](https://)`;
      const at = s.start + label.length + 3;
      return { text: s.text.slice(0, s.start) + ins + s.text.slice(s.end), start: at, end: at + 8 };
    }, false);
    const tableBtn = tb.button('', 'Table: choose its size, inserted at the cursor', () => openTableBar(), 'table');
    tableBtn.addEventListener('mousedown', (ev) => ev.preventDefault());
    alignHBtn = tb.button('', 'Alignment of the paragraph or table column with the cursor: left → center → right', () => alignColumn(), 'alignH');
    alignHBtn.addEventListener('mousedown', (ev) => ev.preventDefault());
    buildTableBar();
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
  // ---- table size bar (state declared with the toolbar) ----
  function buildTableBar(): void {
    tableBar = document.createElement('div');
    tableBar.className = 'findbar linkbar tablebar md-tablebar';
    tableBar.hidden = true;
    const num = (label: string, value: number, max: number) => {
      const wrap = document.createElement('label');
      wrap.className = 'tb-field';
      wrap.textContent = label;
      const i = document.createElement('input');
      i.type = 'number';
      i.min = '1';
      i.max = String(max);
      i.value = String(value);
      i.addEventListener('keydown', (ev) => {
        ev.stopPropagation();
        if (ev.key === 'Enter') doInsertTable();
        else if (ev.key === 'Escape') closeTableBar();
      });
      wrap.appendChild(i);
      return { wrap, i };
    };
    const rows = num('Rows', 3, 99);
    const cols = num('Columns', 3, 20);
    rowsIn = rows.i;
    colsIn = cols.i;
    const mk = (label: string, fn: () => void) => {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'tb-btn';
      b.textContent = label;
      b.addEventListener('mousedown', (ev) => ev.preventDefault());
      b.addEventListener('click', fn);
      return b;
    };
    tableHint = document.createElement('span');
    tableHint.className = 'find-status';
    tableHint.textContent = 'rows below the header row';
    tableBar.append(rows.wrap, cols.wrap, mk('Insert', doInsertTable), mk('Close', closeTableBar), tableHint);
    tb.extra.append(tableBar);
  }
  function openTableBar(): void {
    if (!tableBar) return;
    const sel = window.getSelection();
    tableRange = mode === 'visual' && sel && sel.rangeCount && body.contains(sel.anchorNode) ? sel.getRangeAt(0).cloneRange() : undefined;
    tableSel = mode === 'source' ? { start: ta.selectionStart, end: ta.selectionEnd } : undefined;
    tableBar.hidden = false;
    rowsIn.focus();
    rowsIn.select();
  }
  function closeTableBar(): void {
    if (tableBar) tableBar.hidden = true;
    if (mode === 'visual') {
      body.focus();
      const sel = window.getSelection();
      if (tableRange && sel) {
        sel.removeAllRanges();
        sel.addRange(tableRange);
      }
    } else ta.focus();
  }
  function doInsertTable(): void {
    const rows = Math.round(Number(rowsIn.value));
    const cols = Math.round(Number(colsIn.value));
    if (!(rows >= 1 && rows <= 99 && cols >= 1 && cols <= 20)) {
      tableHint.textContent = '1-99 rows, 1-20 columns';
      return;
    }
    const keep = tableSel;
    closeTableBar();
    insertTable(rows, cols, keep);
  }
  /** A table with a header row and `rows` body rows. */
  function insertTable(rows: number, cols: number, at?: { start: number; end: number }): void {
    const heads = Array.from({ length: cols }, (_, i) => `Column ${i + 1}`);
    if (mode === 'visual') {
      const th = heads.map((h) => `<th>${h}</th>`).join('');
      const tr = `<tr>${'<td><br></td>'.repeat(cols)}</tr>`.repeat(rows);
      exec('insertHTML', `<table><thead><tr>${th}</tr></thead><tbody>${tr}</tbody></table><p><br></p>`);
      return;
    }
    if (at) {
      ta.selectionStart = at.start;
      ta.selectionEnd = at.end;
    }
    transformSelection((s) => {
      const line = (cells: string[]) => `| ${cells.join(' | ')} |`;
      const t = [line(heads), line(heads.map(() => '---')), ...Array.from({ length: rows }, () => line(heads.map(() => '   ')))].join('\n');
      const before = s.text.slice(0, s.end);
      const pre = before === '' || before.endsWith('\n\n') ? '' : before.endsWith('\n') ? '\n' : '\n\n';
      const ins = `${pre}${t}\n`;
      const at0 = s.end + pre.length + 2;
      return { text: before + ins + s.text.slice(s.end), start: at0, end: at0 + heads[0]!.length };
    });
  }
  /**
   * Cycles the alignment left → center → right: of the table column with the cursor (Markdown
   * stores it per column) or of the paragraph (written as `<p align="…">text</p>`).
   */
  function alignColumn(): void {
    const cycle = (cur: string) => (cur === 'center' ? 'right' : cur === 'right' ? 'left' : 'center');
    if (mode !== 'visual') {
      transformSelection((s) => {
        const lines = s.text.split('\n');
        let pos = 0;
        let li = 0;
        while (li < lines.length - 1 && pos + lines[li]!.length < s.start) pos += lines[li++]!.length + 1;
        const line = lines[li] ?? '';
        if (/^\s*\|/.test(line)) {
          // table: the delimiter row of the table around the cursor, column under the cursor
          const cellAt = (s.text.slice(pos, s.start).match(/(?<!\\)\|/g)?.length ?? 0) - (line.trimStart().startsWith('|') ? 1 : 0);
          let top = li;
          while (top > 0 && /^\s*\|/.test(lines[top - 1]!)) top--;
          const d = top + 1;
          if (!/^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$/.test(lines[d] ?? '')) return s;
          const parts = lines[d]!.trim().replace(/^\||\|$/g, '').split('|');
          const k = Math.max(0, Math.min(parts.length - 1, cellAt));
          const p = parts[k]!.trim();
          const cur = p.startsWith(':') && p.endsWith(':') ? 'center' : p.endsWith(':') ? 'right' : 'left';
          const next = cycle(cur);
          parts[k] = ` ${next === 'center' ? ':---:' : next === 'right' ? '---:' : '---'} `;
          lines[d] = `|${parts.join('|')}|`;
          return { text: lines.join('\n'), start: s.start, end: s.start };
        }
        if (!line.trim()) return s;
        const m = /^<p align="(\w+)">(.*)<\/p>\s*$/.exec(line);
        const next = cycle(m?.[1] ?? 'left');
        const inner = m ? m[2]! : line;
        lines[li] = next === 'left' ? inner : `<p align="${next}">${inner}</p>`;
        const at = pos + lines[li]!.length;
        return { text: lines.join('\n'), start: at, end: at };
      });
      return;
    }
    const el = currentElement();
    const cell = el?.closest<HTMLTableCellElement>('td,th');
    const table = cell?.closest('table');
    if (cell && table) {
      const col = cell.cellIndex;
      const next = cycle(cell.style.textAlign || 'left');
      for (const tr of Array.from(table.rows)) {
        const c = tr.cells[col];
        if (c) c.style.textAlign = next === 'left' ? '' : next;
      }
      syncVisual();
      return;
    }
    const p = blockOf(currentNode());
    if (!p) return;
    if (p.tagName !== 'P' && p.tagName !== 'DIV') {
      if (alignHBtn) alignHBtn.title = 'Markdown can align paragraphs and table columns (not headings, lists or quotes)';
      return;
    }
    let target = p;
    if (p.tagName === 'DIV') {
      // a line the browser created: make it a paragraph first
      target = document.createElement('p');
      while (p.firstChild) target.appendChild(p.firstChild);
      p.replaceWith(target);
    }
    const next = cycle(target.style.textAlign || 'left');
    target.style.textAlign = next === 'left' ? '' : next;
    if (!target.getAttribute('style')) target.removeAttribute('style');
    placeCaretIn(target);
    syncVisual();
  }

  function placeCaretIn(el: Node, atEnd = true): void {
    const sel = window.getSelection();
    if (!sel) return;
    const r = document.createRange();
    r.selectNodeContents(el);
    r.collapse(!atEnd);
    sel.removeAllRanges();
    sel.addRange(r);
  }

  /** Bulleted / numbered list on and off for the block with the cursor (no browser list command). */
  function toggleList(tag: 'UL' | 'OL'): void {
    const blk = blockOf(currentNode());
    const li = currentElement()?.closest('li');
    if (!blk) {
      const list = document.createElement(tag);
      const item = document.createElement('li');
      item.appendChild(document.createElement('br'));
      list.appendChild(item);
      body.appendChild(list);
      placeCaretIn(item, false);
    } else if (li && body.contains(li)) {
      const list = li.parentElement!;
      if (list.tagName === tag) {
        // leave the list: the item becomes a paragraph; following items form a new list
        const p = document.createElement('p');
        for (const c of Array.from(li.childNodes)) {
          if (c.nodeType === 1 && /^(UL|OL)$/.test((c as Element).tagName)) continue;
          if (c.nodeType === 1 && (c as Element).tagName === 'P') {
            while (c.firstChild) p.appendChild(c.firstChild);
            continue;
          }
          if (c.nodeType === 1 && (c as Element).classList.contains('task')) continue;
          p.appendChild(c);
        }
        if (!p.firstChild) p.appendChild(document.createElement('br'));
        const tail = document.createElement(list.tagName);
        while (li.nextSibling) tail.appendChild(li.nextSibling);
        const nested = Array.from(li.children).filter((c) => /^(UL|OL)$/.test(c.tagName));
        li.remove();
        const outerLi = list.parentElement?.closest('li');
        if (outerLi) {
          // a nested item moves up one level as text of its own item
          const up = document.createElement('li');
          while (p.firstChild) up.appendChild(p.firstChild);
          up.append(...nested);
          if (tail.firstChild) up.appendChild(tail);
          outerLi.after(up);
          if (!list.firstChild) list.remove();
          placeCaretIn(up);
        } else {
          list.after(p);
          if (tail.firstChild) p.after(tail);
          for (const n of nested) (tail.firstChild ? tail : p).after(n);
          if (!list.firstChild) list.remove();
          placeCaretIn(p);
        }
      } else {
        // switch the kind of this list
        const other = document.createElement(tag);
        for (const a of Array.from(list.attributes)) if (a.name !== 'start' && a.name !== 'data-markup') other.setAttribute(a.name, a.value);
        while (list.firstChild) other.appendChild(list.firstChild);
        list.replaceWith(other);
        placeCaretIn(li);
      }
    } else if (/^(P|DIV|H[1-6])$/.test(blk.tagName)) {
      const item = document.createElement('li');
      while (blk.firstChild) item.appendChild(blk.firstChild);
      if (!item.firstChild) item.appendChild(document.createElement('br'));
      const prev = blk.previousElementSibling;
      if (prev && prev.tagName === tag) {
        prev.appendChild(item);
        blk.remove();
      } else {
        const list = document.createElement(tag);
        list.appendChild(item);
        blk.replaceWith(list);
      }
      placeCaretIn(item);
    } else return;
    syncVisual();
  }

  function openLinkBar(): void {
    if (mode !== 'visual') return;
    const sel = window.getSelection();
    savedRange = sel && sel.rangeCount && body.contains(sel.anchorNode) ? sel.getRangeAt(0).cloneRange() : undefined;
    if (!savedRange) {
      // no cursor in the document: the link goes to the end
      savedRange = document.createRange();
      savedRange.selectNodeContents(body.lastElementChild ?? body);
      savedRange.collapse(false);
    }
    const node = savedRange.commonAncestorContainer;
    linkInput.value = (node.nodeType === 1 ? (node as Element) : node.parentElement)?.closest('a')?.getAttribute('href') ?? 'https://';
    linkBar.hidden = false;
    linkInput.focus();
    linkInput.select();
  }
  /** Applies (or with an empty address removes) the link at the saved selection. */
  function applyLink(apply: boolean): void {
    linkBar.hidden = true;
    const r = savedRange;
    savedRange = undefined;
    body.focus();
    if (!r) return;
    const node = r.commonAncestorContainer;
    const a = (node.nodeType === 1 ? (node as Element) : node.parentElement)?.closest('a');
    const url = linkInput.value.trim();
    if (!apply) {
      placeCaretIn(r.startContainer);
      return;
    }
    if (!url || url === 'https://') {
      if (a && body.contains(a)) a.replaceWith(...Array.from(a.childNodes));
    } else if (/^\s*(javascript|vbscript|data|file):/i.test(url)) {
      return;
    } else if (a && body.contains(a)) {
      a.setAttribute('href', url);
      a.title = url;
    } else {
      const link = document.createElement('a');
      link.setAttribute('href', url);
      link.title = url;
      if (r.collapsed) link.textContent = url;
      else link.appendChild(r.extractContents());
      r.insertNode(link);
      placeCaretIn(link);
    }
    syncVisual();
  }
  linkInput.addEventListener('keydown', (ev) => {
    if (ev.key !== 'Enter' && ev.key !== 'Escape') return;
    ev.preventDefault();
    applyLink(ev.key === 'Enter');
  });
  {
    const mk = (label: string, fn: () => void) => {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'tb-btn';
      b.textContent = label;
      b.addEventListener('mousedown', (ev) => ev.preventDefault());
      b.addEventListener('click', fn);
      return b;
    };
    const hint = document.createElement('span');
    hint.className = 'find-status';
    hint.textContent = 'empty address + Apply: remove the link';
    linkBar.append(
      mk('Apply', () => applyLink(true)),
      mk('Remove', () => {
        linkInput.value = '';
        applyLink(true);
      }),
      mk('Cancel', () => applyLink(false)),
      hint,
    );
  }


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
