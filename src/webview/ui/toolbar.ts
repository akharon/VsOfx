/**
 * A VS Code-style toolbar above the document: flat icon buttons with tooltips, optional text labels,
 * grouped controls. Icons are inline SVG built with DOM APIs (no icon fonts, no remote resources).
 * Buttons that do not fit a narrow row move into a "More" menu (see relayout()).
 */
import { patchViewState, viewState } from '../viewState';

export type IconName = keyof typeof ICONS;

export interface Toolbar {
  /** Root element; contains the button row, the "More" menu and a slot for panels such as the find bar. */
  el: HTMLElement;
  /** Panel slot below the button row. */
  extra: HTMLElement;
  /** `label` may be '' for an icon-only button (the title is then the tooltip and accessible name). */
  button(label: string, title: string, onClick: () => void, icon?: IconName): HTMLButtonElement;
  /** A button that stays pressed; `onChange` receives the new state. */
  toggle(label: string, title: string, onChange: (on: boolean) => void, icon?: IconName): HTMLButtonElement;
  separator(): void;
  /** Adds any control (for example a drop-down) to the row; it moves into "More" like the buttons. */
  custom(node: HTMLElement): void;
  /** Right-aligned status text (word count, page ...). */
  info(): HTMLElement;
  /** Pushes the following controls to the right edge. */
  spacer(): void;
  /** Moves the items that do not fit into the "More" menu, or back into the row. Runs by itself after changes and on resize. */
  relayout(): void;
}

// 16×16 stroke icons (codicon-like, drawn for this project).
const ICONS = {
  undo: ['M4 6h6.5a3.25 3.25 0 0 1 0 6.5H7', 'M6.5 3.5 4 6l2.5 2.5'],
  redo: ['M12 6H5.5a3.25 3.25 0 0 0 0 6.5H9', 'M9.5 3.5 12 6 9.5 8.5'],
  save: ['M3 2.5h8l2.5 2.5v8.5h-11z', 'M5.5 2.5v3h5v-3', 'M5 13.5v-4h6v4'],
  print: ['M4.5 6V2.5h7V6', 'M4.5 11.5h-2v-5.5h11v5.5h-2', 'M4.5 9.5h7v4h-7z'],
  search: ['M7 2.5a4.5 4.5 0 1 0 0 9a4.5 4.5 0 1 0 0-9z', 'M10.3 10.3 13.5 13.5'],
  minus: ['M4 8h8'],
  plus: ['M4 8h8', 'M8 4v8'],
  chevronLeft: ['M10 3.5 5.5 8l4.5 4.5'],
  chevronRight: ['M6 3.5 10.5 8 6 12.5'],
  edit: ['M10.5 2.8l2.7 2.7-7.7 7.7H2.8v-2.7z', 'M9 4.3l2.7 2.7'],
  textField: ['M1.5 4.5h13v7h-13z', 'M4.5 6.3v3.4'],
  checkbox: ['M2.5 2.5h11v11h-11z', 'M5 8.2l2 2 4-4.4'],
  dropdown: ['M1.5 4.5h13v7h-13z', 'M9.5 7.2l1.5 1.5 1.5-1.5'],
  text: ['M3 3.5h10', 'M8 3.5v9', 'M6 12.5h4'],
  rotate: ['M12.8 8.5A4.8 4.8 0 1 1 11.3 4.4', 'M11.8 1.8v3h-3'],
  trash: ['M2.5 4.5h11', 'M6 4.5V2.8h4v1.7', 'M4 4.5l.7 8.7h6.6l.7-8.7'],
  outline: ['M2.5 3.5h11', 'M5 8h8.5', 'M7.5 12.5h6'],
  pilcrow: ['M9.5 13V3', 'M12 13V3', 'M13 3H7.25a2.75 2.75 0 0 0 0 5.5H9.5'],
  freeze: ['M2.5 2.5h11v11h-11z', 'M2.5 6h11'],
  wrap: ['M2 3.5h12', 'M2 8h9.5a2 2 0 0 1 0 4H8', 'M9.5 10.5 8 12l1.5 1.5', 'M2 12.5h3'],
  filter: ['M2 3h12L9.3 8.6V13l-2.6-1.3V8.6z'],
  arrowUp: ['M8 13V3', 'M4 7l4-4 4 4'],
  arrowDown: ['M8 3v10', 'M4 9l4 4 4-4'],
  close: ['M4 4l8 8', 'M12 4l-8 8'],
  preview: ['M1.5 8s2.4-4.5 6.5-4.5S14.5 8 14.5 8 12.1 12.5 8 12.5 1.5 8 1.5 8z', 'M8 6a2 2 0 1 0 0 4a2 2 0 1 0 0-4z'],
  split: ['M2 2.5h12v11H2z', 'M8 2.5v11'],
  code: ['M5.5 4.5 2 8l3.5 3.5', 'M10.5 4.5 14 8l-3.5 3.5'],
  bold: ['M4.5 2.5h4.25a2.75 2.75 0 0 1 0 5.5H4.5z', 'M4.5 8h5a2.75 2.75 0 0 1 0 5.5h-5z'],
  italic: ['M7 2.5h5', 'M4 13.5h5', 'M9.5 2.5l-3 11'],
  underline: ['M4.5 2.5v5a3.5 3.5 0 0 0 7 0v-5', 'M3.5 14h9'],
  strike: ['M2.5 8h11', 'M10.8 4.8C10.4 3.6 9.4 3 8 3 6.3 3 5.2 3.8 5.2 5c0 1 .8 1.6 2.8 2', 'M5.2 11c.4 1.2 1.4 1.8 2.8 1.8 1.7 0 2.8-.8 2.8-2'],
  heading: ['M3.5 3v10', 'M10.5 3v10', 'M3.5 8h7', 'M12.5 9.5l1.5-1v4.5'],
  listBullet: ['M6 4h8', 'M6 8h8', 'M6 12h8', 'M2.5 4h.5', 'M2.5 8h.5', 'M2.5 12h.5'],
  listNumber: ['M6 4h8', 'M6 8h8', 'M6 12h8', 'M2 3l1-.5V6', 'M2 10.5a.9.9 0 0 1 1.7.4L2 13h1.8'],
  quote: ['M3 4v8', 'M6 5h7', 'M6 8h7', 'M6 11h5'],
  link: ['M6.5 9.5l3-3', 'M7.5 4.5l1-1a2.5 2.5 0 0 1 3.5 3.5l-1 1', 'M8.5 11.5l-1 1A2.5 2.5 0 0 1 4 9l1-1'],
  table: ['M2 3h12v10H2z', 'M2 6.5h12', 'M2 10h12', 'M6.5 3v10'],
  alignH: ['M2 3.5h12', 'M4 6.5h8', 'M2 9.5h12', 'M4 12.5h8'],
  image: ['M2 3h12v10H2z', 'M2 11l3.5-3.5 3 3 2-2L14 12', 'M10.5 5.5a1 1 0 1 0 0 .01'],
  sortAsc: ['M4 2.5v11', 'M1.8 11.3 4 13.5l2.2-2.2', 'M8.5 3.5h2', 'M8.5 7.5h3.5', 'M8.5 11.5h5'],
  sortDesc: ['M4 13.5v-11', 'M1.8 4.7 4 2.5l2.2 2.2', 'M8.5 3.5h5', 'M8.5 7.5h3.5', 'M8.5 11.5h2'],
  sum: ['M12 3H4l4.5 5L4 13h8'],
  rowInsert: ['M2 3h12v4H2z', 'M8 9.5v4', 'M6 11.5h4'],
  rowDelete: ['M2 3h12v4H2z', 'M6 11.5h4'],
  colInsert: ['M3 2h4v12H3z', 'M9.5 8h4', 'M11.5 6v4'],
  colDelete: ['M3 2h4v12H3z', 'M9.5 8h4'],
  more: ['M3 8h1', 'M7.5 8h1', 'M12 8h1'],
} satisfies Record<string, string[]>;

const SVG_NS = 'http://www.w3.org/2000/svg';

export function icon(name: IconName): SVGSVGElement {
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('viewBox', '0 0 16 16');
  svg.setAttribute('width', '16');
  svg.setAttribute('height', '16');
  svg.setAttribute('aria-hidden', 'true');
  svg.setAttribute('class', 'icon');
  for (const d of ICONS[name]) {
    const p = document.createElementNS(SVG_NS, 'path');
    p.setAttribute('d', d);
    svg.appendChild(p);
  }
  return svg;
}

/** The first buttons stay in the row whatever the width (Undo, Redo, Save ...). */
const PINNED = 3;

export function createToolbar(): Toolbar {
  const el = document.createElement('div');
  el.className = 'toolbar';
  el.setAttribute('role', 'toolbar');
  const row = document.createElement('div');
  row.className = 'toolbar-row';
  const extra = document.createElement('div');
  extra.className = 'toolbar-extra';

  const more = document.createElement('button');
  more.type = 'button';
  more.className = 'tb-btn icon-only tb-more';
  more.appendChild(icon('more'));
  more.title = 'More actions';
  more.setAttribute('aria-label', 'More actions');
  more.setAttribute('aria-haspopup', 'menu');
  more.setAttribute('aria-expanded', 'false');
  more.hidden = true;
  const menu = document.createElement('div');
  menu.className = 'tb-menu';
  menu.setAttribute('role', 'menu');
  menu.setAttribute('aria-label', 'More actions');
  menu.hidden = true;
  row.appendChild(more);
  el.append(row, menu, extra);

  // Items are added before the "More" button, which stays last.
  const add = (node: Element) => {
    row.insertBefore(node, more);
    schedule();
  };

  const make = (label: string, title: string, ic?: IconName): HTMLButtonElement => {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'tb-btn';
    if (ic) b.appendChild(icon(ic));
    if (label) {
      const s = document.createElement('span');
      s.className = ic ? 'tb-label collapsible' : 'tb-label';
      s.textContent = label;
      b.appendChild(s);
    } else {
      b.classList.add('icon-only');
      b.dataset['menuLabel'] = title; // shown as text inside the "More" menu
    }
    b.title = title;
    b.setAttribute('aria-label', label || title);
    add(b);
    return b;
  };

  // ---- overflow: items that do not fit move into the menu, in their original order ----
  /** Where a moved item came from in the row (an empty comment node), so it can go back there. */
  const holes = new Map<Element, Comment>();
  let queued = false;

  function schedule() {
    if (queued) return;
    queued = true;
    queueMicrotask(() => {
      queued = false;
      relayout();
    });
  }

  const menuButtons = (node: Element): HTMLButtonElement[] =>
    [node, ...Array.from(node.querySelectorAll('button'))].filter((b): b is HTMLButtonElement => b instanceof HTMLButtonElement);

  function toMenu(node: Element) {
    const hole = document.createComment('');
    node.replaceWith(hole);
    holes.set(node, hole);
    menu.prepend(node); // items are taken from the end of the row, so prepending keeps the order
    for (const b of menuButtons(node)) {
      b.setAttribute('role', 'menuitem');
      b.classList.add('tb-menu-item');
      asCheckbox(b);
    }
    if (node instanceof HTMLElement && node.classList.contains('sep')) node.hidden = true;
  }

  function toRow(node: Element) {
    holes.get(node)?.replaceWith(node);
    holes.delete(node);
    for (const b of menuButtons(node)) {
      b.removeAttribute('role');
      b.classList.remove('tb-menu-item');
      const checked = b.getAttribute('aria-checked');
      if (checked !== null) {
        b.setAttribute('aria-pressed', checked);
        b.removeAttribute('aria-checked');
      }
    }
    if (node instanceof HTMLElement && node.classList.contains('sep')) node.hidden = false;
  }

  /** In the menu a toggle is a checkable menu item (role menuitemcheckbox + aria-checked) instead of a pressed button. */
  function asCheckbox(b: HTMLButtonElement) {
    const pressed = b.getAttribute('aria-pressed');
    if (pressed === null) return;
    b.setAttribute('role', 'menuitemcheckbox');
    b.setAttribute('aria-checked', pressed);
    b.removeAttribute('aria-pressed');
  }

  const overflows = () => row.scrollWidth > row.clientWidth;

  /** The last row item that may move: not "More", not a pinned button, not the info text or the spacer. */
  function lastMovable(): Element | undefined {
    const pinned = Array.from(row.children)
      .filter((c) => c !== more && c.classList.contains('tb-btn'))
      .slice(0, PINNED);
    for (let i = row.children.length - 1; i >= 0; i--) {
      const c = row.children[i]!;
      if (c === more || pinned.includes(c) || c.classList.contains('info') || c.classList.contains('tb-spacer')) continue;
      return c;
    }
    return undefined;
  }

  function relayout() {
    for (const node of Array.from(holes.keys())) toRow(node);
    setOpen(false);
    more.hidden = true;
    if (!overflows()) return;
    more.hidden = false;
    while (overflows()) {
      const node = lastMovable();
      if (!node) break;
      toMenu(node);
    }
    if (!menu.children.length) more.hidden = true; // never show an empty menu
  }

  const menuItems = () => Array.from(menu.querySelectorAll<HTMLButtonElement>('button:not(:disabled)'));

  function setOpen(open: boolean, focusFirst = false) {
    menu.hidden = !open;
    more.setAttribute('aria-expanded', String(open));
    if (open) for (const b of menu.querySelectorAll('button')) asCheckbox(b); // a state set from outside while in the menu
    if (open && focusFirst) menuItems()[0]?.focus();
  }

  more.addEventListener('click', () => setOpen(menu.hidden, true));
  menu.addEventListener('keydown', (ev) => {
    if (ev.key === 'Escape') {
      ev.preventDefault();
      setOpen(false);
      more.focus();
      return;
    }
    if (ev.key !== 'ArrowDown' && ev.key !== 'ArrowUp') return;
    const items = menuItems();
    if (!items.length) return;
    ev.preventDefault();
    const i = items.indexOf(document.activeElement as HTMLButtonElement);
    const next = ev.key === 'ArrowDown' ? (i + 1) % items.length : i <= 0 ? items.length - 1 : i - 1;
    items[next]!.focus();
  });
  // an item runs on its own click handler (it is the same node); the menu then closes
  menu.addEventListener('click', (ev) => {
    if (!(ev.target as Element).closest('button')) return;
    const hadFocus = menu.contains(document.activeElement);
    setOpen(false);
    if (hadFocus) more.focus();
  });
  menu.addEventListener('focusout', (ev) => {
    const next = ev.relatedTarget as Node | null;
    if (!next || !(menu.contains(next) || more.contains(next))) setOpen(false);
  });
  document.addEventListener('click', (ev) => {
    if (menu.hidden) return;
    const t = ev.target as Node;
    if (!menu.contains(t) && !more.contains(t)) setOpen(false);
  });
  if (typeof ResizeObserver !== 'undefined') new ResizeObserver(() => relayout()).observe(row);

  return {
    el,
    extra,
    button(label, title, onClick, ic) {
      const b = make(label, title, ic);
      b.addEventListener('click', onClick);
      return b;
    },
    toggle(label, title, onChange, ic) {
      const b = make(label, title, ic);
      b.setAttribute('aria-pressed', 'false');
      b.addEventListener('click', () => {
        // in the "More" menu the state is aria-checked
        const attr = b.hasAttribute('aria-checked') ? 'aria-checked' : 'aria-pressed';
        const on = b.getAttribute(attr) !== 'true';
        b.setAttribute(attr, String(on));
        onChange(on);
      });
      return b;
    },
    custom(node) {
      add(node);
    },
    separator() {
      const s = document.createElement('span');
      s.className = 'sep';
      s.setAttribute('role', 'separator');
      add(s);
    },
    info() {
      const i = document.createElement('span');
      i.className = 'info';
      add(i);
      return i;
    },
    spacer() {
      const s = document.createElement('span');
      s.className = 'tb-spacer';
      add(s);
    },
    relayout,
  };
}

const ZOOM_MIN = 0.5;
const ZOOM_MAX = 3;
const STEPS = [0.5, 0.67, 0.75, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3];

/** A compact "− 100% +" group; `apply` gets the factor (1 = 100 %). Ctrl +/−/0 and Ctrl+wheel also work. */
export function addZoom(tb: Toolbar, apply: (factor: number) => void): void {
  let zoom = 1;
  const group = document.createElement('span');
  group.className = 'tb-group';
  const out = tb.button('', 'Zoom out (Ctrl+−)', () => step(-1), 'minus');
  const reset = tb.button('100%', 'Reset zoom (Ctrl+0)', () => set(1));
  reset.classList.add('zoom-value');
  const inn = tb.button('', 'Zoom in (Ctrl++)', () => step(1), 'plus');
  out.before(group);
  group.append(out, reset, inn);
  const set = (z: number) => {
    zoom = Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, Math.round(z * 100) / 100));
    reset.querySelector('.tb-label')!.textContent = `${Math.round(zoom * 100)}%`;
    out.disabled = zoom <= ZOOM_MIN;
    inn.disabled = zoom >= ZOOM_MAX;
    apply(zoom);
    patchViewState({ zoom });
  };
  const step = (dir: 1 | -1) => {
    const next = dir > 0 ? STEPS.find((s) => s > zoom + 0.001) : [...STEPS].reverse().find((s) => s < zoom - 0.001);
    set(next ?? zoom);
  };
  const saved = viewState().zoom;
  if (typeof saved === 'number' && saved >= ZOOM_MIN && saved <= ZOOM_MAX && saved !== 1) set(saved);
  // one active zoom group per document (the view is rebuilt on revert)
  zoomListeners?.abort();
  zoomListeners = new AbortController();
  const signal = zoomListeners.signal;
  document.addEventListener(
    'keydown',
    (ev) => {
      if (!(ev.ctrlKey || ev.metaKey) || ev.altKey) return;
      if (ev.key === '+' || ev.key === '=') step(1);
      else if (ev.key === '-') step(-1);
      else if (ev.key === '0') set(1);
      else return;
      ev.preventDefault();
    },
    { signal },
  );
  document.addEventListener(
    'wheel',
    (ev) => {
      if (!ev.ctrlKey) return;
      ev.preventDefault();
      step(ev.deltaY < 0 ? 1 : -1);
    },
    { passive: false, signal },
  );
}

let zoomListeners: AbortController | undefined;
