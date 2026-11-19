import { icon, type IconName } from './toolbar';

/** Something that can be searched: a document, a sheet ... */
export interface FindTarget {
  /** Runs a search and returns the number of matches. */
  find(query: string): number;
  /** Shows match `index` (0-based). */
  reveal(index: number): void;
  clear(): void;
}

export interface FindBar {
  open(): void;
  close(): void;
  toggle(): void;
  isOpen(): boolean;
}

export function createFindBar(slot: HTMLElement, target: FindTarget): FindBar {
  const bar = document.createElement('div');
  bar.className = 'findbar';
  bar.hidden = true;
  const input = document.createElement('input');
  input.type = 'text';
  input.placeholder = 'Find';
  input.setAttribute('aria-label', 'Find');
  input.spellcheck = false;
  const status = document.createElement('span');
  status.className = 'find-status';
  status.setAttribute('aria-live', 'polite');
  const mk = (ic: IconName, title: string, fn: () => void) => {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'tb-btn icon-only';
    b.appendChild(icon(ic));
    b.title = title;
    b.setAttribute('aria-label', title);
    b.addEventListener('click', fn);
    return b;
  };
  let count = 0;
  let current = -1;

  const show = () => {
    status.textContent = input.value ? (count ? `${current + 1} of ${count}` : 'No results') : '';
    if (count && current >= 0) target.reveal(current);
  };
  const run = () => {
    if (!input.value) {
      target.clear();
      count = 0;
      current = -1;
      show();
      return;
    }
    count = target.find(input.value);
    current = count ? 0 : -1;
    show();
  };
  const step = (d: number) => {
    if (!count) return;
    current = (current + d + count) % count;
    show();
  };
  const close = () => {
    bar.hidden = true;
    target.clear();
    count = 0;
    current = -1;
  };

  input.addEventListener('input', run);
  input.addEventListener('keydown', (ev) => {
    ev.stopPropagation();
    if (ev.key === 'Enter') {
      ev.preventDefault();
      step(ev.shiftKey ? -1 : 1);
    } else if (ev.key === 'Escape') {
      ev.preventDefault();
      close();
    }
  });
  bar.append(
    input,
    status,
    mk('arrowUp', 'Previous match (Shift+Enter)', () => step(-1)),
    mk('arrowDown', 'Next match (Enter)', () => step(1)),
    mk('close', 'Close (Escape)', close),
  );
  slot.appendChild(bar);

  return {
    open() {
      bar.hidden = false;
      input.focus();
      input.select();
      if (input.value) run();
    },
    close,
    toggle() {
      if (bar.hidden) this.open();
      else close();
    },
    isOpen: () => !bar.hidden,
  };
}

/** Finds text inside a DOM subtree and marks it with the CSS Custom Highlight API (no DOM changes). */
export function domFinder(root: HTMLElement): FindTarget {
  type HighlightRegistry = { set(name: string, h: unknown): void; delete(name: string): void };
  const registry = (): HighlightRegistry | undefined => (globalThis as { CSS?: { highlights?: HighlightRegistry } }).CSS?.highlights;
  const HighlightCtor = (globalThis as { Highlight?: new (...r: Range[]) => unknown }).Highlight;
  let ranges: Range[] = [];

  const paint = (currentIndex: number) => {
    const reg = registry();
    if (!reg || !HighlightCtor) return;
    reg.set('vsofx-find', new HighlightCtor(...ranges));
    const cur = ranges[currentIndex];
    if (cur) reg.set('vsofx-find-current', new HighlightCtor(cur));
  };

  return {
    find(query) {
      ranges = [];
      const needle = query.toLowerCase();
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        const text = (n.nodeValue ?? '').toLowerCase();
        for (let i = text.indexOf(needle); i >= 0; i = text.indexOf(needle, i + needle.length)) {
          const r = document.createRange();
          r.setStart(n, i);
          r.setEnd(n, i + needle.length);
          ranges.push(r);
        }
      }
      paint(-1);
      return ranges.length;
    },
    reveal(index) {
      paint(index);
      ranges[index]?.startContainer.parentElement?.scrollIntoView?.({ block: 'center' });
    },
    clear() {
      ranges = [];
      registry()?.delete('vsofx-find');
      registry()?.delete('vsofx-find-current');
    },
  };
}
