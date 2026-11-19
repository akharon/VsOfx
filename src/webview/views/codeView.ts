import { applyTextEdit, diffText, type CodeModel, type Edit, type TextEdit } from '../../shared/model';
import { formatJson, jsonTree } from '../code/jsonTree';
import { renderMarkdown, setHeading, toggleWrap, togglePrefix, type Selection } from '../code/markdown';
import { lineOf, spliceSpan, type ParsedTree, type Problem, type TreeNode } from '../code/tree';
import { xmlTree } from '../code/xmlTree';
import { yamlTree } from '../code/yamlTree';
import { createFindBar, domFinder, type FindTarget } from '../ui/findBar';
import { addZoom, icon, type Toolbar } from '../ui/toolbar';
import type { EditHandler, View } from './types';

type Mode = 'visual' | 'split' | 'source';

const LANG_NAME = { json: 'JSON', yaml: 'YAML', xml: 'XML', markdown: 'Markdown' } as const;
/** Above this size the visual pane is only built on request (typing stays fast). */
const LARGE = 2_000_000;

function loadMode(lang: string): Mode | undefined {
  try {
    const m = localStorage.getItem(`vsofx.mode.${lang}`);
    return m === 'visual' || m === 'split' || m === 'source' ? m : undefined;
  } catch {
    return undefined;
  }
}
function saveMode(lang: string, m: Mode): void {
  try {
    localStorage.setItem(`vsofx.mode.${lang}`, m);
  } catch {
    /* storage unavailable */
  }
}

/**
 * JSON / YAML / XML: editable tree + source. Markdown: live preview + source with
 * formatting buttons. Both panes edit the same text; every change is a minimal splice.
 */
export function renderCodeView(container: HTMLElement, model: CodeModel, readonly: boolean, onEdit: EditHandler, tb: Toolbar): View {
  const lang = model.lang;
  const isMd = lang === 'markdown';
  let text = model.text; // what the worker (and the host) have
  const own = new Set<Edit>();
  let mode: Mode = loadMode(lang) ?? 'split';

  container.replaceChildren();
  const layout = document.createElement('div');
  layout.className = `code-layout lang-${lang}`;
  const visual = document.createElement('div');
  visual.className = isMd ? 'code-visual md-preview' : 'code-visual tree';
  visual.tabIndex = -1;
  const visualInner = document.createElement('div');
  visualInner.className = isMd ? 'md-body' : 'tree-body';
  visual.appendChild(visualInner);
  const splitter = document.createElement('div');
  splitter.className = 'code-splitter';
  splitter.setAttribute('role', 'separator');
  const source = document.createElement('div');
  source.className = 'code-source';
  const gutter = document.createElement('div');
  gutter.className = 'gutter';
  gutter.setAttribute('aria-hidden', 'true');
  const gutterInner = document.createElement('div');
  gutter.appendChild(gutterInner);
  const ta = document.createElement('textarea');
  ta.className = 'src-text';
  ta.spellcheck = isMd;
  ta.wrap = isMd ? 'soft' : 'off';
  ta.readOnly = readonly;
  ta.value = text;
  ta.setAttribute('aria-label', `${LANG_NAME[lang]} source`);
  source.append(gutter, ta);
  layout.append(visual, splitter, source);
  const problemsBar = document.createElement('div');
  problemsBar.className = 'problems';
  problemsBar.hidden = true;
  container.append(layout, problemsBar);

  // ---- toolbar ----
  const finder = createFindBar(tb.extra, combinedFinder());
  tb.button('', 'Find (Ctrl+F)', () => finder.toggle(), 'search');
  tb.separator();
  const seg = document.createElement('span');
  seg.className = 'tb-group segmented';
  const modeBtns = new Map<Mode, HTMLButtonElement>();
  const modeBtn = (m: Mode, label: string, title: string, ic: 'tree' | 'preview' | 'split' | 'code') => {
    const b = tb.button(label, title, () => setMode(m), ic);
    b.setAttribute('aria-pressed', 'false');
    seg.appendChild(b);
    modeBtns.set(m, b);
  };
  modeBtn('visual', isMd ? 'Preview' : 'Tree', isMd ? 'Rendered preview only' : 'Tree editor only', isMd ? 'preview' : 'tree');
  modeBtn('split', 'Split', 'Side by side', 'split');
  modeBtn('source', 'Source', 'Source text only', 'code');
  tb.el.querySelector('.toolbar-row')!.appendChild(seg);
  tb.separator();
  addZoom(tb, (z) => layout.style.setProperty('--zoom', String(z)));

  if (!readonly && isMd) {
    tb.separator();
    const fmt = (title: string, ic: Parameters<typeof icon>[0], fn: (s: Selection) => Selection) =>
      tb.button('', title, () => transformSelection(fn), ic);
    fmt('Bold (Ctrl+B)', 'bold', (s) => toggleWrap(s, '**'));
    fmt('Italic (Ctrl+I)', 'italic', (s) => toggleWrap(s, '_'));
    fmt('Inline code', 'code', (s) => toggleWrap(s, '`', 'code'));
    fmt('Heading (repeat to change level)', 'heading', (s) => {
      const line = s.text.slice(s.text.lastIndexOf('\n', s.start - 1) + 1);
      const lvl = /^(#{1,6})\s/.exec(line)?.[1]?.length ?? 0;
      return setHeading(s, lvl >= 3 ? 0 : lvl + 1);
    });
    fmt('Bulleted list', 'listBullet', (s) => togglePrefix(s, '- ', /^\s*[-*+]\s+(\[[ xX]\]\s+)?/));
    fmt('Numbered list', 'listNumber', (s) => togglePrefix(s, (i) => `${i + 1}. `, /^\s*\d+[.)]\s+/));
    fmt('Task list', 'checkbox', (s) => togglePrefix(s, '- [ ] ', /^\s*[-*+]\s+\[[ xX]\]\s+/));
    fmt('Quote', 'quote', (s) => togglePrefix(s, '> ', /^>\s?/));
    fmt('Link (Ctrl+K)', 'link', (s) => {
      const label = s.text.slice(s.start, s.end) || 'link';
      const ins = `[${label}](https://)`;
      const at = s.start + label.length + 3;
      return { text: s.text.slice(0, s.start) + ins + s.text.slice(s.end), start: at, end: at + 8 };
    });
    fmt('Table', 'table', (s) => {
      const t = '\n| Column 1 | Column 2 |\n| -------- | -------- |\n| Cell     | Cell     |\n';
      return { text: s.text.slice(0, s.end) + t + s.text.slice(s.end), start: s.end + 3, end: s.end + 11 };
    });
  }
  if (!readonly && lang === 'json') {
    tb.separator();
    tb.button('Format', 'Format document (pretty-print; comments are kept)', () => {
      flush();
      try {
        replaceAll(formatJson(text));
      } catch {
        /* unparsable: leave as is */
      }
    }, 'format');
  }
  if (!isMd) {
    tb.separator();
    tb.button('', 'Expand all', () => setAllOpen(true), 'expand');
    tb.button('', 'Collapse all', () => setAllOpen(false), 'collapse');
  }
  const info = tb.info();

  // ---- source editor ----
  let timer = 0;
  let lineCount = 0;
  const updateGutter = () => {
    const n = ta.value.split('\n').length;
    if (n !== lineCount) {
      lineCount = n;
      const frag: string[] = [];
      for (let i = 1; i <= n; i++) frag.push(String(i));
      gutterInner.textContent = frag.join('\n');
      gutter.style.width = `${String(n).length + 2}ch`;
    }
    gutterInner.style.transform = `translateY(${-ta.scrollTop}px)`;
  };
  /** Commits typing in the source pane as one edit. */
  function flush(): void {
    if (timer) clearTimeout(timer);
    timer = 0;
    const e = diffText(text, ta.value);
    if (!e) return;
    text = ta.value;
    emit(e);
    scheduleVisual();
  }
  function emit(e: TextEdit): void {
    own.add(e);
    onEdit(e);
  }
  /** Replaces the whole text (tree edits, Format, Markdown buttons) as one minimal edit. */
  function replaceAll(next: string, sel?: { start: number; end: number }): void {
    const e = diffText(text, next);
    if (!e) return;
    const before = { s: ta.selectionStart, e: ta.selectionEnd };
    text = next;
    ta.value = next;
    if (sel) ta.setSelectionRange(sel.start, sel.end);
    else ta.setSelectionRange(mapPos(before.s, e), mapPos(before.e, e));
    updateGutter();
    emit(e);
    // after the current event, so that focus has moved (Tab) and can be restored
    if (vTimer) clearTimeout(vTimer);
    vTimer = window.setTimeout(renderVisual, 0);
  }
  function transformSelection(fn: (s: Selection) => Selection): void {
    flush();
    const r = fn({ text, start: ta.selectionStart, end: ta.selectionEnd });
    replaceAll(r.text, r);
    ta.focus();
  }

  ta.addEventListener('input', () => {
    updateGutter();
    if (timer) clearTimeout(timer);
    timer = window.setTimeout(flush, 350);
  });
  ta.addEventListener('blur', () => flush());
  ta.addEventListener('scroll', updateGutter);
  ta.addEventListener('keydown', (ev) => {
    const mod = ev.ctrlKey || ev.metaKey;
    if (ev.key === 'Tab' && !mod && !readonly) {
      ev.preventDefault();
      const unit = '  ';
      if (ev.shiftKey) {
        transformSelection((s) => togglePrefix({ ...s }, '', new RegExp(`^${unit}`)));
      } else if (ta.selectionStart !== ta.selectionEnd && ta.value.slice(ta.selectionStart, ta.selectionEnd).includes('\n')) {
        transformSelection((s) => {
          const ls = s.text.lastIndexOf('\n', s.start - 1) + 1;
          const block = s.text.slice(ls, s.end).replace(/^/gm, unit);
          return { text: s.text.slice(0, ls) + block + s.text.slice(s.end), start: ls, end: ls + block.length };
        });
      } else {
        ta.setRangeText(unit, ta.selectionStart, ta.selectionEnd, 'end');
        ta.dispatchEvent(new Event('input'));
      }
    } else if (ev.key === 'Enter' && !mod && !readonly) {
      // keep indentation (and continue Markdown lists)
      const pos = ta.selectionStart;
      const line = ta.value.slice(ta.value.lastIndexOf('\n', pos - 1) + 1, pos);
      let indent = /^[ \t]*/.exec(line)![0];
      if (isMd) {
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
      } else if (/[{[(]\s*$/.test(line) || (lang === 'yaml' && /:\s*$/.test(line)) || (lang === 'xml' && /<[^/!?][^>]*[^/]>\s*$/.test(line))) {
        indent += '  ';
      }
      ev.preventDefault();
      ta.setRangeText(`\n${indent}`, ta.selectionStart, ta.selectionEnd, 'end');
      ta.dispatchEvent(new Event('input'));
    } else if (mod && isMd && !readonly && (ev.key === 'b' || ev.key === 'i' || ev.key === 'k')) {
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

  // ---- visual pane ----
  const expanded = new Set<string>();
  let autoDepth = 2;
  let vTimer = 0;
  function scheduleVisual(): void {
    if (vTimer) clearTimeout(vTimer);
    vTimer = window.setTimeout(renderVisual, text.length > 200_000 ? 600 : 150);
  }
  function renderVisual(): void {
    if (vTimer) clearTimeout(vTimer);
    vTimer = 0;
    if (mode === 'source') {
      updateProblems(parse()?.problems ?? []);
      return;
    }
    if (isMd) {
      const y = visual.scrollTop;
      renderMarkdown(text, visualInner);
      visual.scrollTop = y;
      updateProblems([]);
      return;
    }
    if (text.length > LARGE && visualInner.dataset['large'] !== 'ok') {
      visualInner.replaceChildren();
      const p = document.createElement('div');
      p.className = 'tree-empty';
      const b = document.createElement('button');
      b.className = 'tb-btn';
      b.textContent = 'Show tree anyway';
      b.addEventListener('click', () => {
        visualInner.dataset['large'] = 'ok';
        renderVisual();
      });
      p.append(`This file is large (${(text.length / 1e6).toFixed(1)} MB). `, b);
      visualInner.appendChild(p);
      return;
    }
    const tree = parse()!;
    updateProblems(tree.problems);
    const active = document.activeElement as HTMLElement | null;
    const focusKey = active?.dataset['focus'];
    const caret = active instanceof HTMLInputElement || active instanceof HTMLTextAreaElement ? active.selectionStart : null;
    const y = visual.scrollTop;
    visualInner.replaceChildren();
    if (tree.roots.length === 0) {
      const p = document.createElement('div');
      p.className = 'tree-empty';
      p.textContent = text.trim() ? 'Nothing to show: fix the problems below or edit the source.' : 'The document is empty. Type in the source pane to start.';
      visualInner.appendChild(p);
    }
    for (const r of tree.roots) visualInner.appendChild(renderNode(r, 0, undefined));
    autoDepth = 0; // only the first render opens levels automatically
    visual.scrollTop = y;
    if (focusKey) {
      const el = visualInner.querySelector<HTMLElement>(`[data-focus="${CSS.escape(focusKey)}"]`);
      if (el) {
        el.focus({ preventScroll: true });
        if (caret !== null && (el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement)) el.setSelectionRange(caret, caret);
      }
    }
  }
  function parse(): ParsedTree | undefined {
    try {
      if (lang === 'json') return jsonTree(text);
      if (lang === 'yaml') return yamlTree(text);
      if (lang === 'xml') return xmlTree(text);
    } catch (err) {
      return { roots: [], problems: [{ message: err instanceof Error ? err.message : String(err), offset: 0 }] };
    }
    return undefined;
  }

  function updateProblems(problems: Problem[]): void {
    if (isMd) {
      const words = text.trim() ? text.trim().split(/\s+/).length : 0;
      info.textContent = `${words.toLocaleString()} words`;
    } else {
      info.textContent = problems.length ? '' : `✓ Valid ${LANG_NAME[lang]}`;
      info.classList.toggle('ok', problems.length === 0);
    }
    problemsBar.hidden = problems.length === 0;
    problemsBar.replaceChildren();
    for (const p of problems.slice(0, 5)) {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'problem';
      b.appendChild(icon('warning'));
      const line = lineOf(text, p.offset);
      b.append(` Line ${line}: ${p.message}`);
      b.title = 'Show in source';
      b.addEventListener('click', () => revealOffset(p.offset));
      problemsBar.appendChild(b);
    }
    if (problems.length > 5) problemsBar.append(` … and ${problems.length - 5} more`);
  }

  function revealOffset(offset: number, end = offset): void {
    if (mode === 'visual') setMode('split');
    ta.focus({ preventScroll: true });
    ta.setSelectionRange(offset, end);
    const line = lineOf(text, offset);
    const lh = parseFloat(getComputedStyle(ta).lineHeight) || 18;
    ta.scrollTop = Math.max(0, (line - 4) * lh);
    updateGutter();
  }

  function renderNode(n: TreeNode, depth: number, parent: TreeNode | undefined): HTMLElement {
    const row = document.createElement('div');
    row.className = `tn-row kind-${n.kind}`;
    row.style.setProperty('--depth', String(depth));
    const hasKids = !!n.children && n.children.length > 0;
    const twisty = document.createElement('span');
    twisty.className = 'twisty';
    row.appendChild(twisty);

    // label (key / tag / index)
    if (n.labelSpan && !readonly) {
      const k = document.createElement('input');
      k.className = 'tn-key';
      k.value = n.label;
      k.dataset['focus'] = `${n.id}:k`;
      k.spellcheck = false;
      k.setAttribute('aria-label', 'Key');
      bindEdit(k, () => n.label, (v) => spliceSpan(text, n.labelSpan!, v));
      row.appendChild(k);
    } else {
      const k = document.createElement('span');
      k.className = parent?.kind === 'array' ? 'tn-index' : n.kind === 'element' ? 'tn-tag' : n.kind === 'attr' ? 'tn-attr' : 'tn-label';
      k.textContent = n.kind === 'element' ? `<${n.label}>` : n.label;
      row.appendChild(k);
    }

    if (n.value !== undefined) {
      const colon = document.createElement('span');
      colon.className = 'tn-colon';
      colon.textContent = n.kind === 'element' || n.kind === 'attr' || n.kind === 'text' ? '' : ':';
      row.appendChild(colon);
      const multi = n.value.includes('\n') || n.value.length > 120;
      const v = multi ? document.createElement('textarea') : document.createElement('input');
      v.className = `tn-val val-${n.kind}`;
      v.value = n.value;
      v.dataset['focus'] = `${n.id}:v`;
      v.spellcheck = false;
      v.setAttribute('aria-label', `Value of ${n.label}`);
      if (v instanceof HTMLTextAreaElement) v.rows = Math.min(8, n.value.split('\n').length + 1);
      if (readonly || !n.valueSpan) v.readOnly = true;
      else bindEdit(v, () => n.value ?? '', (val) => spliceSpan(text, n.valueSpan!, val));
      row.appendChild(v);
    } else if (n.summary) {
      const s = document.createElement('span');
      s.className = 'tn-summary';
      s.textContent = n.kind === 'object' ? `{ ${n.summary} }` : n.kind === 'array' ? `[ ${n.summary} ]` : n.summary;
      row.appendChild(s);
    }

    // actions
    const actions = document.createElement('span');
    actions.className = 'tn-actions';
    const act = (title: string, ic: Parameters<typeof icon>[0], fn: () => void) => {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'tb-btn icon-only';
      b.title = title;
      b.setAttribute('aria-label', title);
      b.appendChild(icon(ic));
      b.addEventListener('click', (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        fn();
      });
      actions.appendChild(b);
    };
    act('Show in source', 'code', () => revealOffset(n.at));
    if (!readonly && n.add) {
      act(n.addLabel ?? 'Add', 'plus', () => {
        flush();
        const next = n.add!();
        if (next !== null) {
          expanded.add(n.id);
          replaceAll(next);
        }
      });
    }
    if (!readonly && n.remove) {
      act('Remove', 'trash', () => {
        flush();
        const next = n.remove!();
        if (next !== null) replaceAll(next);
      });
    }
    row.appendChild(actions);

    if (!hasKids) {
      row.classList.add('leaf');
      return row;
    }
    const details = document.createElement('details');
    details.className = 'tn';
    const open = expanded.has(n.id) || depth < autoDepth;
    if (open) expanded.add(n.id);
    details.open = open;
    const summary = document.createElement('summary');
    summary.appendChild(row);
    details.appendChild(summary);
    // inputs inside <summary> must not toggle the node
    summary.addEventListener('click', (ev) => {
      if ((ev.target as HTMLElement).closest('input, textarea, button')) ev.preventDefault();
    });
    summary.addEventListener('keydown', (ev) => {
      if ((ev.target as HTMLElement).closest('input, textarea') && (ev.key === ' ' || ev.key === 'Enter')) ev.stopPropagation();
    });
    const kids = document.createElement('div');
    kids.className = 'tn-kids';
    details.appendChild(kids);
    const fill = () => {
      if (kids.childElementCount === 0) for (const c of n.children!) kids.appendChild(renderNode(c, depth + 1, n));
    };
    if (open) fill();
    details.addEventListener('toggle', () => {
      if (details.open) {
        expanded.add(n.id);
        fill();
      } else expanded.delete(n.id);
    });
    return details;
  }

  /** Commit on Enter / blur; Escape restores. */
  function bindEdit(el: HTMLInputElement | HTMLTextAreaElement, current: () => string, apply: (v: string) => string): void {
    el.addEventListener('keydown', (ev) => {
      const ke = ev as KeyboardEvent;
      if (ke.key === 'Enter' && (el instanceof HTMLInputElement || ke.ctrlKey)) {
        ke.preventDefault();
        commit();
      } else if (ke.key === 'Escape') {
        el.value = current();
        el.blur();
      }
    });
    el.addEventListener('blur', () => commit());
    let base = current();
    const commit = () => {
      // a node from an older render (its span no longer matches the text) never commits
      if (!el.isConnected || el.value === base) return;
      base = el.value;
      flush();
      replaceAll(apply(el.value));
    };
  }

  function setAllOpen(on: boolean): void {
    if (mode === 'source') setMode('split');
    if (on) {
      // expand the whole tree (rendering lazily on the way)
      for (let guard = 0; guard < 50; guard++) {
        const closed = visualInner.querySelectorAll<HTMLDetailsElement>('details:not([open])');
        if (closed.length === 0) break;
        for (const d of Array.from(closed)) d.open = true;
        if (visualInner.querySelectorAll('.tn-row').length > 20_000) break;
      }
    } else {
      expanded.clear();
      for (const d of Array.from(visualInner.querySelectorAll<HTMLDetailsElement>('details[open]'))) d.open = false;
      const first = visualInner.querySelector<HTMLDetailsElement>('details');
      if (first) first.open = true;
    }
  }

  // markdown preview: click to jump to the source, task checkboxes toggle in the source
  if (isMd) {
    visual.addEventListener('click', (ev) => {
      const t = ev.target as HTMLElement;
      const a = t.closest('a');
      if (a) {
        const href = a.getAttribute('href') ?? '';
        ev.preventDefault();
        if (href.startsWith('#')) {
          const id = decodeURIComponent(href.slice(1)).toLowerCase();
          const h = Array.from(visualInner.querySelectorAll('h1,h2,h3,h4,h5,h6')).find((x) => slug(x.textContent ?? '') === id);
          h?.scrollIntoView({ block: 'start' });
        }
        return;
      }
      if (t instanceof HTMLInputElement && t.classList.contains('task')) {
        ev.preventDefault();
        if (readonly) return;
        flush();
        const lineNo = Number(t.dataset['line']);
        const lines = text.split('\n');
        const l = lines[lineNo];
        if (l !== undefined) {
          lines[lineNo] = l.replace(/\[([ xX])\]/, (_m, c: string) => (c === ' ' ? '[x]' : '[ ]'));
          replaceAll(lines.join('\n'));
        }
        return;
      }
      const block = t.closest<HTMLElement>('[data-line]');
      if (block && mode === 'split' && window.getSelection()?.isCollapsed !== false) {
        const line = Number(block.dataset['line']);
        let off = 0;
        for (let i = 0; i < line; i++) off = text.indexOf('\n', off) + 1;
        const end = text.indexOf('\n', off);
        revealOffset(off, end < 0 ? text.length : end);
      }
    });
  }

  // ---- layout / modes ----
  function setMode(m: Mode): void {
    if (mode === 'source' && m !== 'source') flush();
    mode = m;
    saveMode(lang, m);
    layout.dataset['mode'] = m;
    for (const [k, b] of modeBtns) b.setAttribute('aria-pressed', String(k === m));
    renderVisual();
    if (m !== 'visual') updateGutter();
  }

  // draggable splitter
  splitter.addEventListener('pointerdown', (ev) => {
    ev.preventDefault();
    splitter.setPointerCapture(ev.pointerId);
    const rect = layout.getBoundingClientRect();
    const move = (e: PointerEvent) => {
      const f = Math.min(0.85, Math.max(0.15, (e.clientX - rect.left) / rect.width));
      layout.style.setProperty('--split', `${f * 100}%`);
    };
    const up = () => {
      splitter.removeEventListener('pointermove', move);
      splitter.removeEventListener('pointerup', up);
    };
    splitter.addEventListener('pointermove', move);
    splitter.addEventListener('pointerup', up);
  });

  function combinedFinder(): FindTarget {
    const dom = domFinder(visualInner);
    let hits: number[] = [];
    let len = 0;
    let useSource = true;
    return {
      find(q) {
        // the tree shows values in inputs, which a DOM text search cannot see: search the source
        useSource = !isMd || mode !== 'visual';
        if (!useSource) return dom.find(q);
        if (mode === 'visual') setMode('split');
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
        updateGutter();
      },
      clear() {
        dom.clear();
        hits = [];
      },
    };
  }

  setMode(mode);
  updateGutter();

  return {
    applyEdit(edit) {
      if (edit.kind !== 'text') return;
      if (own.delete(edit)) return; // our own edit coming back
      // undo / redo / an edit in another editor of the same file: rebase pending typing
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
      updateGutter();
      scheduleVisual();
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
