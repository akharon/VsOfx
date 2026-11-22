import type { ApplyResult, Edit, NewParaEdit, ParagraphEdit, TextModel } from '../../shared/model';
import {
  H_ALIGNS,
  SAFE_HREF,
  TASK_DONE,
  TASK_OPEN,
  PARA_BASE,
  cellIndex,
  fmtEquals,
  isAnchorIndex,
  isCellIndex,
  isNewParaIndex,
  normalizeRuns,
  runsText,
  runsToHtml,
  type BlockKind,
  type HAlign,
  type ParaFmt,
  type Run,
  type TableSpec,
} from '../../shared/richtext';
import { reportStatus } from '../viewState';
import { createFindBar, domFinder } from '../ui/findBar';
import { addZoom, createToolbar, type IconName, type Toolbar } from '../ui/toolbar';
import type { EditHandler, View } from './types';

/** Block kind a rendered element has by itself (without an explicit data-block). */
function naturalBlock(el: HTMLElement): BlockKind {
  const tag = el.tagName.toLowerCase();
  if (/^h[1-6]$/.test(tag)) return tag as BlockKind;
  if (tag === 'li') return el.parentElement?.tagName.toLowerCase() === 'ol' ? 'ol' : 'ul';
  return 'p';
}
const blockOf = (el: HTMLElement): BlockKind => (el.dataset['block'] as BlockKind | undefined) ?? naturalBlock(el);

/** Horizontal alignment a paragraph is shown with. */
function computedAlign(el: HTMLElement): HAlign {
  const a = el.style.textAlign || (typeof getComputedStyle === 'function' ? getComputedStyle(el).textAlign : '');
  return a === 'center' ? 'center' : a === 'right' || a === 'end' ? 'right' : a === 'justify' ? 'justify' : 'left';
}

interface RemovedEl {
  el: HTMLElement;
  /** What was taken out of the page: the block, or the list / list item that only held it. */
  unit: HTMLElement;
  parent: Node;
  prev: Node | null;
  next: Node | null;
  paras: number[];
  tables: number[];
}

/** Entries of the paragraph style drop-down. */
const STYLES: [BlockKind, string][] = [['p', 'Normal'], ['h1', 'Heading 1'], ['h2', 'Heading 2'], ['h3', 'Heading 3'], ['quote', 'Quote']];

const FLAG_KEYS =['b', 'i', 'u', 's', 'code', 'sup', 'sub'] as const;

/** Formatted runs of an editable block (nested lists and tables are not part of it). */
export function serializeRuns(el: HTMLElement): Run[] {
  const runs: Run[] = [];
  const walk = (node: Node, f: Omit<Run, 't'>) => {
    for (let n = node.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 3) {
        runs.push({ t: (n.nodeValue ?? '').replace(/​/g, '').replace(/ /g, ' '), ...f });
        continue;
      }
      if (n.nodeType !== 1) continue;
      const e = n as HTMLElement;
      const tag = e.tagName.toLowerCase();
      if (tag === 'ul' || tag === 'ol' || tag === 'table') continue;
      if (tag === 'br') {
        runs.push({ t: '\n', ...f });
        continue;
      }
      const g: Omit<Run, 't'> = { ...f };
      if (tag === 'strong' || tag === 'b') g.b = true;
      if (tag === 'em' || tag === 'i') g.i = true;
      if (tag === 'u') g.u = true;
      if (tag === 's' || tag === 'strike' || tag === 'del') g.s = true;
      if (tag === 'code') g.code = true;
      if (tag === 'sup') g.sup = true;
      if (tag === 'sub') g.sub = true;
      if (tag === 'a') {
        const href = e.dataset['href'] ?? e.getAttribute('href') ?? '';
        if (SAFE_HREF.test(href)) g.href = href;
      }
      const st = e.style;
      if (st) {
        if (st.fontWeight === 'bold' || Number(st.fontWeight) >= 600) g.b = true;
        if (st.fontStyle === 'italic') g.i = true;
        if (st.textDecoration.includes('underline') || st.textDecorationLine?.includes('underline')) g.u = true;
        if (st.textDecoration.includes('line-through') || st.textDecorationLine?.includes('line-through')) g.s = true;
      }
      walk(e, g);
    }
  };
  walk(el, {});
  // a trailing <br> that browsers keep in empty editable blocks is not content
  const out = normalizeRuns(runs);
  const last = out[out.length - 1];
  if (last && last.t.endsWith('\n') && el.lastChild?.nodeName === 'BR') {
    last.t = last.t.slice(0, -1);
    if (last.t === '') out.pop();
  }
  return out;
}

/**
 * Document text view. `model.html` is produced by the sanitizer in the worker
 * (allowlisted tags only); blocks with data-p are editable paragraphs. For Word
 * (`model.rich`) paragraphs keep bold/italic/links ..., and the toolbar can change
 * formatting, headings, lists, quotes and insert tables.
 */
export function renderTextView(
  container: HTMLElement,
  model: TextModel,
  readonly: boolean,
  onEdit: EditHandler,
  tb: Toolbar = createToolbar(),
): View {
  const rich = !!model.rich && !readonly;
  container.replaceChildren();
  const layout = document.createElement('div');
  layout.className = 'doc-layout';
  const nav = document.createElement('nav');
  nav.className = 'doc-nav';
  nav.setAttribute('aria-label', 'Document outline');
  nav.hidden = true;
  const scroll = document.createElement('div');
  scroll.className = 'doc-scroll';
  const page = document.createElement('article');
  page.className = 'page';
  page.innerHTML = model.html;
  scroll.appendChild(page);
  layout.append(nav, scroll);
  container.appendChild(layout);

  const current = new Map<number, string>(); // last committed text per paragraph
  const currentFmt = new Map<number, ParaFmt>(); // rich: last committed formatting
  const blocks = new Map<number, HTMLElement>();
  const tables = new Map<number, { spec: TableSpec; el: HTMLTableElement }>();
  const anchors = new Map<number, number>();
  /** Removed original paragraphs: the detached element and where it was (for undo). */
  const removedEls = new Map<number, RemovedEl>();
  let lastNew = 0; // highest added-paragraph index handed out (ids are never reused)
  let active: HTMLElement | null = null;
  // link editor state (a small bar under the toolbar: window.prompt is not available in webviews)
  let linkRange: Range | null = null;
  let linkInput!: HTMLInputElement;
  let linkTarget: HTMLElement | null = null;
  // table size bar state
  let tableBar!: HTMLDivElement;
  let rowsInput!: HTMLInputElement;
  let colsInput!: HTMLInputElement;
  let delTableBtn!: HTMLButtonElement;
  let tableTarget: HTMLElement | null = null;
  // paragraph style drop-down and the last selection inside the page (the drop-down takes the focus)
  let styleSelect: HTMLSelectElement | undefined;
  let savedRange: Range | null = null;

  const fmtOf = (el: HTMLElement): ParaFmt => {
    const f: ParaFmt = { runs: serializeRuns(el) };
    const b = el.dataset['block'] as BlockKind | undefined;
    if (b) f.block = b;
    const a = el.dataset['align'] as HAlign | undefined;
    if (a) f.align = a;
    return f;
  };
  /** Shows the alignment of a paragraph (after undo / redo). */
  const showAlign = (el: HTMLElement, fmt?: ParaFmt) => {
    if (!fmt) return;
    if (fmt.align) el.dataset['align'] = fmt.align;
    else delete el.dataset['align'];
  };
  const register = (el: HTMLElement, index: number) => {
    blocks.set(index, el);
    showAlign(el);
    if (rich) {
      const f = fmtOf(el);
      currentFmt.set(index, f);
      current.set(index, runsText(f.runs));
    } else current.set(index, el.textContent ?? '');
    if (readonly) return;
    // rich text: the page itself is the editing surface (set below); plain text: one host per block
    if (!rich) el.setAttribute('contenteditable', 'plaintext-only');
    el.setAttribute('spellcheck', 'false');
    el.classList.add('editable');
  };
  for (const el of Array.from(page.querySelectorAll<HTMLElement>('[data-p]'))) register(el, Number(el.dataset['p']));
  if (readonly) page.classList.add('readonly');
  for (const t of model.tables ?? []) createTable(t.table, t.cells);
  for (const p of model.paras ?? []) createParagraph(p.index, p.after, p.fmt);
  // tables that the adapter rendered itself (Writer)
  for (const el of Array.from(page.querySelectorAll<HTMLTableElement>('table.ins-table[data-table]'))) {
    const n = (k: string) => Number(el.dataset[k]);
    tables.set(n('table'), { spec: { id: n('table'), after: n('after'), rows: n('rows'), cols: n('cols') }, el });
  }
  if (rich) {
    // one editing surface: blocks that are not editable text (images, unmatched blocks) stay read-only islands
    page.setAttribute('contenteditable', 'true');
    page.setAttribute('spellcheck', 'false');
    for (const el of Array.from(page.querySelectorAll<HTMLElement>('p,h1,h2,h3,h4,h5,h6,li,pre,blockquote,img,table'))) {
      if (!el.hasAttribute('data-p') && !el.querySelector('[data-p]') && !el.closest('[data-p],[contenteditable="false"]')) {
        el.setAttribute('contenteditable', 'false');
      }
    }
  }

  // ---- toolbar: find, zoom, navigation pane, formatting marks, word count ----
  const finder = createFindBar(tb.extra, domFinder(page));
  tb.button('', 'Find in document (Ctrl+F)', () => finder.toggle(), 'search');
  tb.separator();
  addZoom(tb, (z) => layout.style.setProperty('--zoom', String(z)));
  tb.separator();
  let navOpen = false;
  tb.toggle('Outline', 'Navigation pane: headings of the document', (on) => {
    navOpen = on;
    nav.hidden = !on;
    if (on) buildNav();
  }, 'outline');
  tb.toggle('', 'Show formatting marks (paragraph ends)', (on) => page.classList.toggle('marks', on), 'pilcrow');

  // ---- formatting (Word) ----
  const linkBar = document.createElement('div');
  linkBar.className = 'findbar linkbar';
  linkBar.hidden = true;
  if (rich) {
    tb.separator();
    const fmtBtn = (title: string, ic: IconName, fn: () => void) => {
      const b = tb.button('', title, () => {
        if (!active) return;
        fn();
      }, ic);
      // keep the caret / selection in the paragraph
      b.addEventListener('mousedown', (ev) => ev.preventDefault());
      return b;
    };
    const dd = styleSelect = document.createElement('select');
    styleSelect.className = 'tb-select';
    styleSelect.title = 'Paragraph style (Ctrl+Alt+0 Normal, Ctrl+Alt+1..3 Heading 1..3)';
    styleSelect.setAttribute('aria-label', 'Paragraph style');
    for (const [v, label] of STYLES) styleSelect.add(new Option(label, v));
    const other = new Option('Other', 'other');
    other.disabled = true;
    styleSelect.add(other);
    styleSelect.addEventListener('change', () => {
      const v = dd.value as BlockKind;
      const sel = window.getSelection();
      const keep = sel && sel.rangeCount && page.contains(sel.anchorNode) ? sel.getRangeAt(0).cloneRange() : savedRange;
      page.focus?.({ preventScroll: true }); // focusing the page may move the caret: put the selection back
      if (keep) {
        sel?.removeAllRanges();
        sel?.addRange(keep);
      }
      setStyle(v);
    });
    tb.custom(styleSelect);
    fmtBtn('Bold (Ctrl+B)', 'bold', () => inline('bold'));
    fmtBtn('Italic (Ctrl+I)', 'italic', () => inline('italic'));
    fmtBtn('Underline (Ctrl+U)', 'underline', () => inline('underline'));
    fmtBtn('Strikethrough (Ctrl+Shift+X)', 'strike', () => inline('strikeThrough'));
    fmtBtn('Code', 'code', () => toggleWrap('code'));
    fmtBtn('Bulleted list (Ctrl+Shift+8)', 'listBullet', () => setBlock(active!, blockOf(active!) === 'ul' ? 'p' : 'ul'));
    fmtBtn('Numbered list (Ctrl+Shift+7)', 'listNumber', () => setBlock(active!, blockOf(active!) === 'ol' ? 'p' : 'ol'));
    fmtBtn('Task (☐ → ☑ → none)', 'checkbox', () => toggleTask(active!));
    fmtBtn('Quote', 'quote', () => setBlock(active!, blockOf(active!) === 'quote' ? 'p' : 'quote'));
    fmtBtn('Link (Ctrl+K)', 'link', () => openLinkBar());
    fmtBtn('Table: insert below the cursor paragraph, choose its size', 'table', () => openTableBar());
    fmtBtn('Horizontal alignment: left → center → right → justified', 'alignH', () => {
      const el = active!;
      const cur = (el.dataset['align'] as HAlign | undefined) ?? computedAlign(el);
      el.dataset['align'] = H_ALIGNS[(H_ALIGNS.indexOf(cur) + 1) % H_ALIGNS.length]!;
      commit(el);
    });
    buildLinkBar();
    buildTableBar();
  }
  const counter = tb.info();

  function buildNav() {
    nav.replaceChildren();
    const headings = Array.from(page.querySelectorAll<HTMLElement>('h1,h2,h3,h4,h5,h6,[data-block^="h"]')).filter((h) => /^h/.test(blockOf(h)));
    if (headings.length === 0) {
      const none = document.createElement('p');
      none.className = 'nav-empty';
      none.textContent = 'No headings in this document.';
      nav.appendChild(none);
      return;
    }
    for (const h of headings) {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'nav-item';
      b.style.paddingLeft = `${6 + (Number(blockOf(h).slice(1)) - 1) * 12}px`;
      b.textContent = (h.textContent ?? '').trim() || '(empty heading)';
      b.addEventListener('click', () => h.scrollIntoView?.({ block: 'start', behavior: 'smooth' }));
      nav.appendChild(b);
    }
  }

  function refreshStats() {
    // Count per block: textContent alone would glue neighbouring paragraphs together.
    const BLOCKS = 'p,h1,h2,h3,h4,h5,h6,li,td,th';
    let words = 0;
    for (const b of Array.from(page.querySelectorAll(BLOCKS))) {
      if (!b.querySelector(BLOCKS)) words += (b.textContent ?? '').split(/\s+/).filter(Boolean).length;
    }
    counter.textContent = `${words.toLocaleString()} word${words === 1 ? '' : 's'}`;
    reportStatus(counter.textContent);
    if (navOpen) buildNav();
  }
  refreshStats();

  // ---- editing ----
  /** The paragraph edit for what the user changed in `el` (and marks it as committed). */
  const pendingEdit = (el: HTMLElement): ParagraphEdit | null => {
    const index = Number(el.dataset['p']);
    const old = current.get(index) ?? '';
    if (!rich) {
      const text = (el.textContent ?? '').replace(/\n+/g, ' ');
      if (text === old) return null;
      current.set(index, text);
      return { kind: 'paragraph', index, text, old };
    }
    const oldFmt = currentFmt.get(index) ?? { runs: [] };
    const fmt = fmtOf(el);
    if (fmtEquals(fmt, oldFmt)) return null;
    const text = runsText(fmt.runs);
    current.set(index, text);
    currentFmt.set(index, fmt);
    return { kind: 'paragraph', index, text, old, fmt, oldFmt };
  };
  const commit = (el: HTMLElement) => {
    const e = pendingEdit(el);
    if (!e) return;
    refreshStats();
    onEdit(e);
  };
  const flush = () => {
    if (active && active.isConnected) commit(active);
  };

  function setBlock(el: HTMLElement, kind: BlockKind) {
    if (isCellIndex(Number(el.dataset['p']))) return; // table cells stay paragraphs
    el.dataset['block'] = kind;
    commit(el);
    syncStyle();
  }

  /** Shows the style of the block with the caret in the drop-down. */
  function syncStyle() {
    if (!styleSelect) return;
    const b = active && active.isConnected ? blockOf(active) : 'p';
    styleSelect.value = STYLES.some(([v]) => v === b) ? b : 'other';
  }

  /** Applies a paragraph style to the selected blocks (one undo step). */
  function setStyle(kind: BlockKind) {
    const s = readSelection();
    const first = s?.first ?? active;
    if (!first) return;
    if (!s || s.first === s.last) {
      setBlock(first, kind);
    } else {
      const list = blockList();
      for (const b of list.slice(list.indexOf(s.first), list.indexOf(s.last) + 1)) {
        if (!isCellBlock(b)) b.dataset['block'] = kind;
      }
      commitTouched();
    }
    syncStyle();
  }

  function inline(cmd: 'bold' | 'italic' | 'underline' | 'strikeThrough') {
    document.execCommand('styleWithCSS', false, 'false');
    document.execCommand(cmd);
    // formatting a selection over several blocks changes all of them: report them as one step
    const s = readSelection();
    if (s && s.first !== s.last) commitTouched();
  }
  function commitTouched() {
    const sel = window.getSelection();
    const r = sel && sel.rangeCount ? sel.getRangeAt(0) : null;
    const edits: Edit[] = [];
    for (const b of blockList()) {
      if (b !== active && !(r && r.intersectsNode(b))) continue;
      const e = pendingEdit(b);
      if (e) edits.push(e);
    }
    if (edits.length === 0) return;
    refreshStats();
    onEdit(edits.length === 1 ? edits[0]! : { kind: 'batch', edits });
  }

  /** Range of the current selection if it lies inside the active paragraph. */
  function selRange(): Range | null {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0 || !active) return null;
    const r = sel.getRangeAt(0);
    return active.contains(r.commonAncestorContainer) ? r : null;
  }

  function unwrap(e: Element) {
    const parent = e.parentNode!;
    while (e.firstChild) parent.insertBefore(e.firstChild, e);
    parent.removeChild(e);
  }

  function toggleWrap(tag: 'code') {
    const r = selRange();
    if (!r) return;
    const node = r.commonAncestorContainer;
    const inside = (node.nodeType === 1 ? (node as Element) : node.parentElement)?.closest(tag);
    if (inside && active!.contains(inside)) unwrap(inside);
    else if (!r.collapsed) {
      const w = document.createElement(tag);
      w.appendChild(r.extractContents());
      r.insertNode(w);
      const sel = window.getSelection()!;
      sel.removeAllRanges();
      const nr = document.createRange();
      nr.selectNodeContents(w);
      sel.addRange(nr);
    }
  }

  function toggleTask(el: HTMLElement) {
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    const first = walker.nextNode() as Text | null;
    const t = first?.nodeValue ?? '';
    if (first && t.startsWith(`${TASK_OPEN} `)) first.nodeValue = `${TASK_DONE} ${t.slice(2)}`;
    else if (first && t.startsWith(`${TASK_DONE} `)) first.nodeValue = t.slice(2);
    else if (first) first.nodeValue = `${TASK_OPEN} ${t}`;
    else el.textContent = `${TASK_OPEN} `;
    commit(el);
  }

  // link editor: a small bar under the toolbar (window.prompt is not available in webviews)
  function buildLinkBar() {
    linkInput = document.createElement('input');
    linkInput.type = 'text';
    linkInput.placeholder = 'https://example.com';
    linkInput.setAttribute('aria-label', 'Link address');
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
    linkBar.append(linkInput, mk('Apply', applyLink), mk('Remove', removeLink), mk('Cancel', closeLinkBar), hint);
    linkInput.addEventListener('keydown', (ev) => {
      ev.stopPropagation();
      if (ev.key === 'Enter') applyLink();
      else if (ev.key === 'Escape') closeLinkBar();
    });
    linkInput.addEventListener('input', () => {
      hint.textContent = linkInput.value && !SAFE_HREF.test(linkInput.value.trim()) ? 'Use http://, https:// or mailto:' : '';
    });
    tb.extra.appendChild(linkBar);
  }
  function openLinkBar() {
    const r = selRange();
    if (!r) return;
    linkRange = r.cloneRange();
    linkTarget = active;
    const node = r.commonAncestorContainer;
    const a = (node.nodeType === 1 ? (node as Element) : node.parentElement)?.closest<HTMLElement>('a[data-href]');
    linkInput.value = a?.dataset['href'] ?? 'https://';
    linkBar.hidden = false;
    linkInput.focus();
    linkInput.select();
  }
  function closeLinkBar() {
    linkBar.hidden = true;
    if (linkTarget && linkRange) {
      linkTarget.focus();
      const sel = window.getSelection()!;
      sel.removeAllRanges();
      sel.addRange(linkRange);
    }
  }
  function applyLink() {
    const href = linkInput.value.trim();
    if (!SAFE_HREF.test(href) || !linkRange || !linkTarget) return;
    const node = linkRange.commonAncestorContainer;
    const a = (node.nodeType === 1 ? (node as Element) : node.parentElement)?.closest<HTMLElement>('a[data-href]');
    if (a && linkTarget.contains(a)) a.dataset['href'] = href;
    else {
      const el = document.createElement('a');
      el.dataset['href'] = href;
      if (linkRange.collapsed) el.textContent = href;
      else el.appendChild(linkRange.extractContents());
      linkRange.insertNode(el);
      linkRange.selectNodeContents(el);
    }
    const target = linkTarget;
    closeLinkBar();
    commit(target);
  }
  function removeLink() {
    if (linkRange && linkTarget) {
      const node = linkRange.commonAncestorContainer;
      const a = (node.nodeType === 1 ? (node as Element) : node.parentElement)?.closest('a[data-href]');
      if (a && linkTarget.contains(a)) unwrap(a);
      const target = linkTarget;
      closeLinkBar();
      commit(target);
    }
  }

  // ---- added paragraphs (Enter) ----
  // (anchors: added paragraph -> the paragraph it was added after; declared at the top)
  /** Element for an added paragraph: a list item next to a list item, else a paragraph. */
  function createParagraph(index: number, after: number, fmt: ParaFmt): HTMLElement | undefined {
    const anchor = blocks.get(after) ?? page.querySelector<HTMLElement>(`[data-p="${after}"]`);
    if (!anchor) return undefined;
    const el = document.createElement(anchor.tagName.toLowerCase() === 'li' ? 'li' : 'p');
    el.dataset['p'] = String(index);
    if (fmt.block) el.dataset['block'] = fmt.block;
    if (fmt.align) el.dataset['align'] = fmt.align;
    el.innerHTML = runsToHtml(fmt.runs);
    anchor.after(el);
    register(el, index);
    anchors.set(index, after);
    return el;
  }
  function removeParagraph(index: number) {
    const el = blocks.get(index);
    if (!el) return;
    // paragraphs added after this one now follow its own anchor (the adapters do the same)
    const own = anchors.get(index);
    for (const [i, a] of anchors) if (a === index && own !== undefined) anchors.set(i, own);
    anchors.delete(index);
    if (active === el) active = null;
    el.remove();
    blocks.delete(index);
    current.delete(index);
    currentFmt.delete(index);
  }
  function caretRange(): Range | null {
    const sel = window.getSelection();
    return sel && sel.rangeCount ? sel.getRangeAt(0) : null;
  }
  function placeCaret(el: HTMLElement, atEnd: boolean) {
    if (rich) {
      caretAt(el, atEnd ? Infinity : 0);
      return;
    }
    el.focus();
    const r = document.createRange();
    r.selectNodeContents(el);
    r.collapse(!atEnd);
    const sel = window.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(r);
  }

  // ---- one editing surface (rich text) ----
  // The browser edits text inside one paragraph by itself. Everything that changes the paragraph
  // structure (a selection over several paragraphs, Enter, Backspace / Delete at a paragraph edge,
  // multi-line paste) is done here in `rewrite`: ONE edit (a batch when it is several) per action.
  // Limits (nothing is changed, the document cannot be corrupted):
  //  - table cells: no operation crosses into or out of a cell, a cell never gets a second paragraph;
  //  - a selection / join that has a table or a non-editable block (image) between its ends;
  //  - paragraphs that contain nested lists or tables (only plain "leaf" paragraphs are rewritten).
  const blockList = () => Array.from(page.querySelectorAll<HTMLElement>('[data-p]'));
  const indexOfBlock = (b: HTMLElement) => Number(b.dataset['p']);
  const isCellBlock = (b: HTMLElement) => !!b.closest('td,th') || isCellIndex(indexOfBlock(b));
  const isLeaf = (b: HTMLElement) => !b.querySelector('[data-p],table,ul,ol');
  const blockAt = (node: Node | null): HTMLElement | null => {
    const e = node && node.nodeType === 1 ? (node as Element) : (node?.parentElement ?? null);
    const b = e?.closest<HTMLElement>('[data-p]') ?? null;
    return b && page.contains(b) ? b : null;
  };
  const newId = () => {
    lastNew = Math.max(lastNew, PARA_BASE, ...[...blocks.keys()].filter(isNewParaIndex)) + 1;
    return lastNew;
  };
  /** The paragraph an added paragraph follows (the adapters of .odt do not tell: use the one before it). */
  const anchorOf = (index: number): number => {
    const a = anchors.get(index);
    if (a !== undefined) return a;
    const list = blockList();
    const el = blocks.get(index);
    for (let i = el ? list.indexOf(el) - 1 : -1; i >= 0; i--) {
      const n = indexOfBlock(list[i]!);
      if (isAnchorIndex(n)) return n;
    }
    return 0;
  };
  /** Whether a table or a non-editable block lies between two blocks. */
  const blocked = (a: HTMLElement, b: HTMLElement) =>
    Array.from(page.querySelectorAll('table,[contenteditable="false"]')).some(
      (n) =>
        !!(a.compareDocumentPosition(n) & Node.DOCUMENT_POSITION_FOLLOWING) &&
        !!(b.compareDocumentPosition(n) & Node.DOCUMENT_POSITION_PRECEDING) &&
        !n.contains(a) &&
        !n.contains(b),
    );
  const setRuns = (el: HTMLElement, runs: Run[]) => {
    el.innerHTML = runsToHtml(runs) || '<br>'; // the placeholder <br> is not content (see serializeRuns)
  };
  const styleOf = (r?: Run): Omit<Run, 't'> => {
    const o: Omit<Run, 't'> = {};
    if (r) for (const k of FLAG_KEYS) if (r[k]) o[k] = true;
    return o;
  };
  const textRun = (t: string, st: Omit<Run, 't'>): Run[] => (t === '' ? [] : [{ t, ...st }]);
  /** Block kind of the paragraph that follows `el` after Enter: a heading is followed by text, lists and quotes continue. */
  const nextKind = (el: HTMLElement): BlockKind | undefined => {
    const kind = blockOf(el);
    return /^h/.test(kind) ? 'p' : kind === 'quote' ? 'quote' : el.dataset['block'] === 'ul' || el.dataset['block'] === 'ol' ? kind : undefined;
  };

  interface Sel {
    first: HTMLElement;
    last: HTMLElement;
    /** Runs of `first` before the selection start / of `last` after its end. */
    before: Run[];
    after: Run[];
    collapsed: boolean;
  }
  const runsAround = (el: HTMLElement, node: Node, offset: number, side: 'before' | 'after'): Run[] => {
    const rg = document.createRange();
    if (side === 'before') {
      rg.setStart(el, 0);
      rg.setEnd(node, offset);
    } else {
      rg.setStart(node, offset);
      rg.setEnd(el, el.childNodes.length);
    }
    const holder = document.createElement('div');
    holder.appendChild(rg.cloneContents());
    return serializeRuns(holder);
  };
  function readSelection(): Sel | null {
    const s = window.getSelection();
    if (!s || s.rangeCount === 0) return null;
    const r = s.getRangeAt(0);
    if (!page.contains(r.commonAncestorContainer)) return null;
    const inside = blockList().filter((b) => r.intersectsNode(b));
    const a = blockAt(r.startContainer);
    const z = blockAt(r.endContainer);
    const first = a ?? inside[0];
    const last = z ?? inside[inside.length - 1];
    if (!first || !last) return null;
    return {
      first,
      last,
      before: a ? runsAround(a, r.startContainer, r.startOffset, 'before') : [],
      after: z ? runsAround(z, r.endContainer, r.endOffset, 'after') : [],
      collapsed: r.collapsed,
    };
  }
  const selBlock = (): HTMLElement | null => {
    const s = window.getSelection();
    return s && s.rangeCount && page.contains(s.anchorNode) ? blockAt(s.anchorNode) : null;
  };

  /** Caret at a text offset of a block (Infinity: the end). */
  function caretAt(el: HTMLElement, offset: number) {
    const sel = window.getSelection();
    if (!sel) return;
    page.focus?.({ preventScroll: true });
    const r = document.createRange();
    let left = offset;
    let done = false;
    const walk = (n: Node) => {
      for (let c = n.firstChild; c && !done; c = c.nextSibling) {
        if (c.nodeType === 3) {
          const len = (c.nodeValue ?? '').length;
          if (left <= len) {
            r.setStart(c, left);
            done = true;
          } else left -= len;
        } else if (c.nodeType === 1) {
          const tag = (c as Element).tagName;
          if (tag === 'UL' || tag === 'OL' || tag === 'TABLE') continue;
          if (tag === 'BR') {
            if (left === 0) {
              r.setStartBefore(c);
              done = true;
            } else left -= 1;
          } else walk(c);
        }
      }
    };
    walk(el);
    if (!done) {
      r.selectNodeContents(el);
      r.collapse(false);
    }
    r.collapse(true);
    sel.removeAllRanges();
    sel.addRange(r);
    el.scrollIntoView?.({ block: 'nearest' });
  }

  /** Takes an original paragraph out of the page, keeping its element so that undo can put it back. */
  function removeOriginal(index: number, to?: number) {
    const el = blocks.get(index);
    if (!el || removedEls.has(index)) return;
    let unit: HTMLElement = el;
    while (unit.parentElement && unit.parentElement !== page && /^(LI|UL|OL)$/.test(unit.parentElement.tagName) && unit.parentElement.children.length === 1) {
      unit = unit.parentElement;
    }
    const info: RemovedEl = { el, unit, parent: unit.parentNode!, prev: unit.previousSibling, next: unit.nextSibling, paras: [], tables: [] };
    if (to !== undefined) {
      for (const [i, a] of anchors) {
        if (a === index) {
          anchors.set(i, to);
          info.paras.push(i);
        }
      }
      for (const t of tables.values()) {
        if (t.spec.after === index) {
          t.spec = { ...t.spec, after: to };
          info.tables.push(t.spec.id);
        }
      }
    }
    if (active === el) active = null;
    unit.remove();
    blocks.delete(index);
    current.delete(index);
    currentFmt.delete(index);
    removedEls.set(index, info);
  }
  function restoreOriginal(index: number, fmt: ParaFmt) {
    const info = removedEls.get(index);
    if (!info) return;
    const { el, unit, parent, prev, next } = info;
    if (next && next.parentNode === parent) parent.insertBefore(unit, next);
    else if (prev && prev.parentNode === parent) parent.insertBefore(unit, prev.nextSibling);
    else parent.appendChild(unit);
    setRuns(el, fmt.runs);
    if (fmt.block) el.dataset['block'] = fmt.block;
    else delete el.dataset['block'];
    showAlign(el, fmt);
    blocks.set(index, el);
    currentFmt.set(index, fmt);
    current.set(index, runsText(fmt.runs));
    for (const i of info.paras) if (anchors.has(i)) anchors.set(i, index);
    for (const id of info.tables) {
      const t = tables.get(id);
      if (t) t.spec = { ...t.spec, after: index };
    }
    removedEls.delete(index);
  }

  /**
   * The central structural edit. The selection runs from block `first` (text before it: `before`)
   * to block `last` (text after it: `after`); everything in between is replaced by `lines`:
   * first := before + lines[0] (+ after when there is one line), the lines after the first become new
   * paragraphs, the last of them followed by `after`; the blocks after `first` up to `last` are removed.
   */
  function rewrite(first: HTMLElement, last: HTMLElement, before: Run[], after: Run[], lines: string[]): void {
    const list = blockList();
    const i = list.indexOf(first);
    const j = list.indexOf(last);
    if (i < 0 || j < i) return;
    const span = list.slice(i, j + 1);
    if (span.some(isCellBlock) && (span.length > 1 || lines.length > 1)) return; // never across / into table cells
    if (!span.every(isLeaf) || (span.length > 1 && blocked(first, last))) return;
    const edits: Edit[] = [];
    const rest = span.slice(1);
    // what the user typed into the other blocks before is reported first (only the active block can be dirty)
    for (const b of rest) {
      const e = pendingEdit(b);
      if (e) edits.push(e);
    }
    const multi = lines.length > 1;
    const st = styleOf(before[before.length - 1] ?? after[0]);
    setRuns(first, normalizeRuns([...before, ...textRun(lines[0]!, st), ...(multi ? [] : after)]));
    const head = pendingEdit(first);
    if (head) edits.push(head);
    const fi = indexOfBlock(first);
    for (const b of rest.reverse()) {
      const bi = indexOfBlock(b);
      const fmt = currentFmt.get(bi) ?? { runs: [] };
      if (isNewParaIndex(bi)) {
        edits.push({ kind: 'newpara', op: 'remove', index: bi, after: anchorOf(bi), fmt });
        removeParagraph(bi);
      } else {
        edits.push({ kind: 'delpara', op: 'remove', index: bi, fmt, to: fi });
        removeOriginal(bi, fi);
      }
    }
    let caretEl = first;
    let caretAtText = runsText(before).length + lines[0]!.length;
    if (multi) {
      const block = nextKind(first);
      let anchor = fi;
      for (let k = 1; k < lines.length; k++) {
        const text = lines[k]!;
        const runs = k === lines.length - 1 ? normalizeRuns([...textRun(text, {}), ...after]) : textRun(text, {});
        const fmt: ParaFmt = { runs };
        if (block) fmt.block = block;
        const index = newId();
        const el = createParagraph(index, anchor, fmt);
        if (!el) break;
        edits.push({ kind: 'newpara', op: 'insert', index, after: anchor, fmt });
        anchor = index;
        caretEl = el;
        caretAtText = text.length;
      }
    }
    if (edits.length > 0) {
      onEdit(edits.length === 1 ? edits[0]! : { kind: 'batch', edits });
      refreshStats();
    }
    active = caretEl;
    caretAt(caretEl, caretAtText);
  }

  /** Enter: the text after the cursor (the selection is deleted first) moves into a new paragraph below. */
  function splitSelection(s: Sel) {
    if (isCellBlock(s.first)) {
      commit(s.first); // a table cell holds one paragraph
      return;
    }
    rewrite(s.first, s.last, s.before, s.after, ['', '']);
  }
  /** Paste of plain text: the first line at the cursor, further lines as new paragraphs. */
  function pasteText(s: Sel, text: string) {
    let lines = text.replace(/\r\n?/g, '\n').split('\n');
    if (lines.length > 1 && lines[lines.length - 1] === '') lines.pop(); // text copied with its final line break
    if (isCellBlock(s.first) || isCellBlock(s.last)) lines = [lines.join(' ')];
    rewrite(s.first, s.last, s.before, s.after, lines);
  }
  /** Backspace at the start / Delete at the end of a paragraph: joins it with the neighbour. */
  function join(el: HTMLElement, dir: -1 | 1): boolean {
    const list = blockList();
    const other = list[list.indexOf(el) + dir];
    if (!other) return false;
    const [a, b] = dir < 0 ? [other, el] : [el, other];
    if (isCellBlock(a) || isCellBlock(b)) return true; // not across cells: handled (= nothing happens)
    rewrite(a, b, serializeRuns(a), serializeRuns(b), ['']);
    return true;
  }
  const onBeforeInput = (ev: InputEvent) => {
    const t = ev.inputType;
    if (t === 'insertCompositionText' || t === 'insertFromComposition') return;
    const s = readSelection();
    if (!s) return;
    const multi = s.first !== s.last;
    if (t === 'insertText' || t === 'insertReplacementText') {
      if (!multi) return; // inside one paragraph the browser does it
      ev.preventDefault();
      rewrite(s.first, s.last, s.before, s.after, [(ev.data ?? '').replace(/\r?\n/g, ' ')]);
    } else if (t === 'insertParagraph') {
      ev.preventDefault();
      splitSelection(s);
    } else if (t === 'insertFromPaste' || t === 'insertFromPasteAsQuotation') {
      ev.preventDefault();
      pasteText(s, ev.dataTransfer?.getData('text/plain') ?? ev.data ?? '');
    } else if (t.startsWith('delete')) {
      if (!s.collapsed) {
        if (!multi) return;
        ev.preventDefault();
        rewrite(s.first, s.last, s.before, s.after, ['']);
      } else if (t.endsWith('Backward') && runsText(s.before) === '') {
        if (join(s.first, -1)) ev.preventDefault();
      } else if (t.endsWith('Forward') && runsText(s.after) === '') {
        if (join(s.first, 1)) ev.preventDefault();
      }
    }
  };
  /** The block with the caret is the active one; the one the caret left is committed. */
  const onSelectionChange = () => {
    const s = window.getSelection();
    if (!s || s.rangeCount === 0 || !page.contains(s.anchorNode)) return;
    savedRange = s.getRangeAt(0).cloneRange();
    const el = blockAt(s.anchorNode);
    if (!el) return;
    if (el !== active) {
      const prev = active;
      active = el;
      prev?.classList.remove('cur');
      el.classList.add('cur');
      if (prev && prev.isConnected) commit(prev);
    }
    syncStyle();
  };
  /** Backspace in an empty added paragraph removes it again. */
  function removeEmpty(el: HTMLElement): boolean {
    const index = Number(el.dataset['p']);
    if (!isNewParaIndex(index) || (el.textContent ?? '').replace(/​/g, '') !== '') return false;
    const all = Array.from(page.querySelectorAll<HTMLElement>('[data-p]'));
    const prev = all[all.indexOf(el) - 1];
    const after = anchors.get(index);
    if (after === undefined) return false;
    const fmt = currentFmt.get(index) ?? { runs: [] };
    const e: NewParaEdit = { kind: 'newpara', op: 'remove', index, after, fmt };
    removeParagraph(index);
    onEdit(e);
    refreshStats();
    if (prev) placeCaret(prev, true);
    return true;
  }
  /** Moves the cursor to the previous / next paragraph when it is at the edge of this one. */
  function moveAcross(el: HTMLElement, ev: KeyboardEvent): boolean {
    const r = caretRange();
    if (!r || !r.collapsed) return false;
    const up = ev.key === 'ArrowUp' || ev.key === 'ArrowLeft';
    const probe = document.createRange();
    probe.selectNodeContents(el);
    if (up) probe.setEnd(r.startContainer, r.startOffset);
    else probe.setStart(r.startContainer, r.startOffset);
    const rest = probe.toString();
    let edge: boolean;
    if (ev.key === 'ArrowLeft' || ev.key === 'ArrowRight') edge = rest === '';
    else {
      // first / last line: no line break between the cursor and the edge
      const caret = r.getBoundingClientRect?.();
      const box = el.getBoundingClientRect();
      const line = parseFloat(getComputedStyle(el).lineHeight) || 20;
      edge = rest === '' || !caret || caret.height === 0 || (up ? caret.top - box.top < line * 0.9 : box.bottom - caret.bottom < line * 0.9);
    }
    if (!edge) return false;
    const all = Array.from(page.querySelectorAll<HTMLElement>('[data-p]'));
    const next = all[all.indexOf(el) + (up ? -1 : 1)];
    if (!next) return false;
    placeCaret(next, up);
    next.scrollIntoView?.({ block: 'nearest' });
    return true;
  }

  // ---- inserted tables ----
  function buildTableBar() {
    tableBar = document.createElement('div');
    tableBar.className = 'findbar linkbar tablebar';
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
        else if (ev.key === 'Escape') tableBar.hidden = true;
      });
      wrap.appendChild(i);
      return { wrap, i };
    };
    const rows = num('Rows', 3, 99);
    const cols = num('Columns', 3, 20);
    rowsInput = rows.i;
    colsInput = cols.i;
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
    hint.textContent = 'below the paragraph with the cursor';
    delTableBtn = mk('Delete this table', () => {
      const t = [...tables.values()].find((x) => tableTarget && x.el.contains(tableTarget));
      if (!t) return;
      removeTable(t.spec.id);
      onEdit({ kind: 'table', op: 'remove', table: t.spec });
      tableBar.hidden = true;
      refreshStats();
    });
    tableBar.append(rows.wrap, cols.wrap, mk('Insert', doInsertTable), delTableBtn, mk('Close', () => (tableBar.hidden = true)), hint);
    tb.extra.appendChild(tableBar);
  }
  function openTableBar() {
    tableTarget = active;
    delTableBtn.hidden = ![...tables.values()].some((x) => active && x.el.contains(active));
    tableBar.hidden = false;
    rowsInput.focus();
    rowsInput.select();
  }
  function doInsertTable() {
    const rows = Math.round(Number(rowsInput.value));
    const cols = Math.round(Number(colsInput.value));
    if (!(rows >= 1 && rows <= 99 && cols >= 1 && cols <= 20 && rows * cols <= 999) || !tableTarget) {
      rowsInput.setCustomValidity?.('1-99 rows, 1-20 columns');
      return;
    }
    tableBar.hidden = true;
    insertTable(tableTarget, rows, cols);
  }

  function createTable(spec: TableSpec, cells: string[] = []) {
    const anchor = page.querySelector<HTMLElement>(`[data-p="${spec.after}"]`);
    if (!anchor) return;
    const table = document.createElement('table');
    table.className = 'ins-table';
    table.dataset['table'] = String(spec.id);
    for (let r = 0; r < spec.rows; r++) {
      const tr = table.insertRow();
      for (let c = 0; c < spec.cols; c++) {
        const td = tr.insertCell();
        const p = document.createElement('p');
        const index = cellIndex(spec, r, c);
        p.dataset['p'] = String(index);
        p.innerHTML = cells[r * spec.cols + c] ?? '';
        td.appendChild(p);
        register(p, index);
      }
    }
    // after the anchor's list / table / earlier inserted tables
    let after: Element = anchor.closest('li') ? (anchor.closest('ul,ol') as Element) : anchor;
    while (after.nextElementSibling?.classList.contains('ins-table')) after = after.nextElementSibling;
    after.after(table);
    tables.set(spec.id, { spec, el: table });
  }
  function removeTable(id: number) {
    const t = tables.get(id);
    if (!t) return;
    for (const p of Array.from(t.el.querySelectorAll<HTMLElement>('[data-p]'))) {
      const i = Number(p.dataset['p']);
      blocks.delete(i);
      current.delete(i);
      currentFmt.delete(i);
    }
    t.el.remove();
    tables.delete(id);
  }
  function insertTable(el: HTMLElement, rows: number, cols: number) {
    let anchorIndex = Number(el.dataset['p']);
    if (isCellIndex(anchorIndex)) {
      // inside an inserted table: put the new one after the same paragraph
      const host = [...tables.values()].find((t) => t.el.contains(el));
      if (!host) return;
      anchorIndex = host.spec.after;
    }
    const id = Math.max(0, ...tables.keys()) + 1;
    const spec: TableSpec = { id, after: anchorIndex, rows, cols };
    createTable(spec);
    onEdit({ kind: 'table', op: 'insert', table: spec });
    tables.get(id)?.el.querySelector<HTMLElement>('[data-p]')?.focus();
    refreshStats();
  }

  const onKeyDown = (ev: KeyboardEvent) => {
    // one editing surface: the keyboard target is the page, the paragraph is where the caret is
    const el = (rich ? selBlock() : null) ?? (ev.target as HTMLElement).closest<HTMLElement>('[data-p]');
    if (!el) return;
    const mod = ev.ctrlKey || ev.metaKey;
    if (ev.key === 'Enter' && !ev.shiftKey) {
      ev.preventDefault();
      const s = rich ? readSelection() : null;
      if (s) splitSelection(s);
      else el.blur();
    } else if (ev.key === 'Backspace' && rich && !mod && readSelection()?.collapsed && removeEmpty(el)) {
      ev.preventDefault();
    } else if (!rich && /^Arrow(Up|Down|Left|Right)$/.test(ev.key) && !ev.shiftKey && !mod && !ev.altKey && moveAcross(el, ev)) {
      ev.preventDefault();
    } else if (ev.key === 'Enter' && ev.shiftKey && !rich) {
      ev.preventDefault();
    } else if (ev.key === 'Escape') {
      ev.preventDefault();
      const index = Number(el.dataset['p']);
      if (rich) el.innerHTML = runsToHtml(currentFmt.get(index)?.runs ?? []);
      else el.textContent = current.get(index) ?? '';
      el.blur();
    } else if (rich && mod && !ev.altKey && !ev.shiftKey && /^[biuk]$/i.test(ev.key)) {
      ev.preventDefault();
      const k = ev.key.toLowerCase();
      if (k === 'k') openLinkBar();
      else inline(k === 'b' ? 'bold' : k === 'i' ? 'italic' : 'underline');
    } else if (rich && mod && ev.shiftKey && !ev.altKey && ev.key.toLowerCase() === 'x') {
      ev.preventDefault();
      inline('strikeThrough');
    } else if (rich && mod && ev.shiftKey && !ev.altKey && (ev.code === 'Digit7' || ev.code === 'Digit8')) {
      ev.preventDefault();
      const kind = ev.code === 'Digit7' ? 'ol' : 'ul';
      setBlock(el, blockOf(el) === kind ? 'p' : kind);
    } else if (rich && mod && ev.altKey && !ev.shiftKey && /^Digit[0-3]$/.test(ev.code)) {
      ev.preventDefault(); // by key position: Ctrl+Alt is AltGr on many layouts
      setStyle((['p', 'h1', 'h2', 'h3'] as const)[Number(ev.code.slice(5))]!);
    }
  };
  const onFocusIn = (ev: FocusEvent) => {
    const el = (ev.target as HTMLElement).closest<HTMLElement>('[data-p]');
    if (el) active = el;
  };
  const onFocusOut = (ev: FocusEvent) => {
    const el = (ev.target as HTMLElement).closest<HTMLElement>('[data-p]') ?? (rich ? active : null);
    if (el && el.isConnected) commit(el);
  };
  // only plain text comes in from the clipboard (no foreign markup)
  const onPaste = (ev: ClipboardEvent) => {
    if (rich) {
      ev.preventDefault();
      const s = readSelection();
      if (s) pasteText(s, ev.clipboardData?.getData('text/plain') ?? '');
      return;
    }
    if (!(ev.target as HTMLElement).closest('[data-p]')) return;
    ev.preventDefault();
    const t = (ev.clipboardData?.getData('text/plain') ?? '').replace(/\r?\n/g, ' ');
    document.execCommand('insertText', false, t);
  };
  const onDrop = (ev: DragEvent) => ev.preventDefault();
  // undo/redo/save are VS Code commands: commit the paragraph being typed first
  const onCapture = (ev: KeyboardEvent) => {
    if ((ev.ctrlKey || ev.metaKey) && /^[zys]$/i.test(ev.key)) flush();
    if ((ev.ctrlKey || ev.metaKey) && !ev.altKey && ev.key.toLowerCase() === 'f') {
      ev.preventDefault();
      finder.open();
    }
  };
  // tickboxes: a click on ☐ / ☑ (or Word's ☒) toggles it, like a form checkbox
  const BOX_NEXT: Record<string, string> = { [TASK_OPEN]: TASK_DONE, [TASK_DONE]: TASK_OPEN, '☒': TASK_OPEN };
  const onClick = (ev: MouseEvent) => {
    const el = (ev.target as HTMLElement).closest<HTMLElement>('[data-p]');
    if (!el || ev.ctrlKey || ev.metaKey || ev.shiftKey) return;
    const pos = document.caretRangeFromPoint?.(ev.clientX, ev.clientY);
    const node = pos?.startContainer;
    if (!pos || !node || node.nodeType !== 3 || !el.contains(node)) return;
    const s = node.nodeValue ?? '';
    // the caret lands before or after the glyph depending on which half was clicked
    for (const i of [pos.startOffset, pos.startOffset - 1]) {
      const next = BOX_NEXT[s[i] ?? ''];
      if (!next) continue;
      const r = document.createRange();
      r.setStart(node, i);
      r.setEnd(node, i + 1);
      const box = r.getBoundingClientRect();
      if (ev.clientX < box.left - 2 || ev.clientX > box.right + 2 || ev.clientY < box.top - 2 || ev.clientY > box.bottom + 2) continue;
      ev.preventDefault();
      node.nodeValue = s.slice(0, i) + next + s.slice(i + 1);
      const sel = window.getSelection();
      const after = document.createRange();
      after.setStart(node, i + 1);
      sel?.removeAllRanges();
      sel?.addRange(after);
      commit(el);
      return;
    }
  };
  document.addEventListener('keydown', onCapture, true);
  if (!readonly) {
    page.addEventListener('click', onClick);
    page.addEventListener('keydown', onKeyDown);
    page.addEventListener('focusin', onFocusIn);
    page.addEventListener('focusout', onFocusOut);
    page.addEventListener('paste', onPaste);
    page.addEventListener('drop', onDrop);
    if (rich) {
      page.addEventListener('beforeinput', onBeforeInput);
      document.addEventListener('selectionchange', onSelectionChange);
    }
  }

  return {
    applyEdit(edit: Edit, _result: ApplyResult) {
      applyOne(edit);
    },
    flush,
    dispose() {
      finder.close();
      reportStatus('');
      page.removeEventListener('beforeinput', onBeforeInput);
      document.removeEventListener('selectionchange', onSelectionChange);
      document.removeEventListener('keydown', onCapture, true);
      page.removeEventListener('click', onClick);
      page.removeEventListener('keydown', onKeyDown);
      page.removeEventListener('focusin', onFocusIn);
      page.removeEventListener('focusout', onFocusOut);
      page.removeEventListener('paste', onPaste);
      page.removeEventListener('drop', onDrop);
    },
  };

  function applyOne(edit: Edit) {
    if (edit.kind === 'batch') {
      for (const e of edit.edits) applyOne(e);
      return;
    }
    if (edit.kind === 'table') {
      if (edit.op === 'insert' && !tables.has(edit.table.id)) createTable(edit.table);
      else if (edit.op === 'remove') removeTable(edit.table.id);
      refreshStats();
      return;
    }
    if (edit.kind === 'newpara') {
      if (edit.op === 'insert' && !blocks.has(edit.index)) createParagraph(edit.index, edit.after, edit.fmt);
      else if (edit.op === 'remove') removeParagraph(edit.index);
      refreshStats();
      return;
    }
    if (edit.kind === 'delpara') {
      if (edit.op === 'remove') removeOriginal(edit.index, edit.to);
      else restoreOriginal(edit.index, edit.fmt);
      refreshStats();
      return;
    }
    if (edit.kind !== 'paragraph') return;
    const el = blocks.get(edit.index);
    if (!el) return;
    applyParagraph(el, edit);
    refreshStats();
  }

  function applyParagraph(el: HTMLElement, edit: ParagraphEdit) {
    if (edit.fmt) {
      if (!fmtEquals(fmtOf(el), edit.fmt)) el.innerHTML = runsToHtml(edit.fmt.runs);
      if (edit.fmt.block) el.dataset['block'] = edit.fmt.block;
      else delete el.dataset['block'];
      showAlign(el, edit.fmt);
      currentFmt.set(edit.index, edit.fmt);
    } else el.textContent = edit.text;
    current.set(edit.index, edit.text);
    if (el === active) syncStyle();
  }
}
