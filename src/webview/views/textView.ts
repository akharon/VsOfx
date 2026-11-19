import type { ApplyResult, Edit, TextModel } from '../../shared/model';
import { createFindBar, domFinder } from '../ui/findBar';
import { addZoom, createToolbar, type Toolbar } from '../ui/toolbar';
import type { EditHandler, View } from './types';

/**
 * Document text view. `model.html` is produced by the sanitizer in the worker
 * (allowlisted tags only); blocks with data-p are editable paragraphs.
 */
export function renderTextView(
  container: HTMLElement,
  model: TextModel,
  readonly: boolean,
  onEdit: EditHandler,
  tb: Toolbar = createToolbar(),
): View {
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
  const blocks = new Map<number, HTMLElement>();
  for (const el of Array.from(page.querySelectorAll<HTMLElement>('[data-p]'))) {
    const index = Number(el.dataset['p']);
    blocks.set(index, el);
    current.set(index, el.textContent ?? '');
    if (readonly) continue;
    el.setAttribute('contenteditable', 'plaintext-only');
    el.setAttribute('spellcheck', 'false');
    el.classList.add('editable');
  }
  if (readonly) page.classList.add('readonly');

  // ---- toolbar: find, zoom, navigation pane, formatting marks, word count ----
  const finder = createFindBar(tb.extra, domFinder(page));
  tb.button('Find', 'Find in document (Ctrl+F)', () => finder.toggle());
  addZoom(tb, (z) => layout.style.setProperty('--zoom', String(z)));
  tb.separator();
  let navOpen = false;
  tb.toggle('Outline', 'Navigation pane: headings of the document', (on) => {
    navOpen = on;
    nav.hidden = !on;
    if (on) buildNav();
  });
  tb.toggle('¶', 'Show formatting marks (paragraph ends)', (on) => page.classList.toggle('marks', on));
  const counter = tb.info();

  function buildNav() {
    nav.replaceChildren();
    const headings = Array.from(page.querySelectorAll<HTMLElement>('h1,h2,h3,h4,h5,h6'));
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
      b.style.paddingLeft = `${6 + (Number(h.tagName.slice(1)) - 1) * 12}px`;
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
    if (navOpen) buildNav();
  }
  refreshStats();

  const onFind = (ev: KeyboardEvent) => {
    if ((ev.ctrlKey || ev.metaKey) && !ev.altKey && ev.key.toLowerCase() === 'f') {
      ev.preventDefault();
      finder.open();
    }
  };
  document.addEventListener('keydown', onFind);

  // ---- editing ----
  const commit = (el: HTMLElement) => {
    const index = Number(el.dataset['p']);
    const old = current.get(index) ?? '';
    const text = (el.textContent ?? '').replace(/\n+/g, ' ');
    if (text === old) return;
    current.set(index, text);
    refreshStats();
    onEdit({ kind: 'paragraph', index, text, old });
  };

  const onKeyDown = (ev: KeyboardEvent) => {
    const el = (ev.target as HTMLElement).closest<HTMLElement>('[data-p]');
    if (!el) return;
    if (ev.key === 'Enter') {
      ev.preventDefault();
      el.blur();
    } else if (ev.key === 'Escape') {
      ev.preventDefault();
      el.textContent = current.get(Number(el.dataset['p'])) ?? '';
      el.blur();
    }
  };
  const onFocusOut = (ev: FocusEvent) => {
    const el = (ev.target as HTMLElement).closest<HTMLElement>('[data-p]');
    if (el) commit(el);
  };
  if (!readonly) {
    page.addEventListener('keydown', onKeyDown);
    page.addEventListener('focusout', onFocusOut);
  }

  return {
    applyEdit(edit: Edit, _result: ApplyResult) {
      if (edit.kind !== 'paragraph') return;
      const el = blocks.get(edit.index);
      if (!el) return;
      el.textContent = edit.text;
      current.set(edit.index, edit.text);
      refreshStats();
    },
    dispose() {
      finder.close();
      document.removeEventListener('keydown', onFind);
      page.removeEventListener('keydown', onKeyDown);
      page.removeEventListener('focusout', onFocusOut);
    },
  };
}
