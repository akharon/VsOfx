/**
 * A VS Code-style toolbar above the document: flat icon buttons with tooltips, optional text labels,
 * grouped controls. Icons are inline SVG built with DOM APIs (no icon fonts, no remote resources).
 */
export type IconName = keyof typeof ICONS;

export interface Toolbar {
  /** Root element; contains the button row and a slot for panels such as the find bar. */
  el: HTMLElement;
  /** Panel slot below the button row. */
  extra: HTMLElement;
  /** `label` may be '' for an icon-only button (the title is then the tooltip and accessible name). */
  button(label: string, title: string, onClick: () => void, icon?: IconName): HTMLButtonElement;
  /** A button that stays pressed; `onChange` receives the new state. */
  toggle(label: string, title: string, onChange: (on: boolean) => void, icon?: IconName): HTMLButtonElement;
  separator(): void;
  /** Right-aligned status text (word count, page ...). */
  info(): HTMLElement;
  /** Pushes the following controls to the right edge. */
  spacer(): void;
}

// 16×16 stroke icons (codicon-like, drawn for this project).
const ICONS = {
  undo: ['M4 6h6.5a3.25 3.25 0 0 1 0 6.5H7', 'M6.5 3.5 4 6l2.5 2.5'],
  redo: ['M12 6H5.5a3.25 3.25 0 0 0 0 6.5H9', 'M9.5 3.5 12 6 9.5 8.5'],
  save: ['M3 2.5h8l2.5 2.5v8.5h-11z', 'M5.5 2.5v3h5v-3', 'M5 13.5v-4h6v4'],
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
  filter: ['M2 3h12L9.3 8.6V13l-2.6-1.3V8.6z'],
  arrowUp: ['M8 13V3', 'M4 7l4-4 4 4'],
  arrowDown: ['M8 3v10', 'M4 9l4 4 4-4'],
  close: ['M4 4l8 8', 'M12 4l-8 8'],
  preview: ['M1.5 8s2.4-4.5 6.5-4.5S14.5 8 14.5 8 12.1 12.5 8 12.5 1.5 8 1.5 8z', 'M8 6a2 2 0 1 0 0 4a2 2 0 1 0 0-4z'],
  split: ['M2 2.5h12v11H2z', 'M8 2.5v11'],
  code: ['M5.5 4.5 2 8l3.5 3.5', 'M10.5 4.5 14 8l-3.5 3.5'],
  bold: ['M4.5 2.5h4.25a2.75 2.75 0 0 1 0 5.5H4.5z', 'M4.5 8h5a2.75 2.75 0 0 1 0 5.5h-5z'],
  italic: ['M7 2.5h5', 'M4 13.5h5', 'M9.5 2.5l-3 11'],
  heading: ['M3.5 3v10', 'M10.5 3v10', 'M3.5 8h7', 'M12.5 9.5l1.5-1v4.5'],
  listBullet: ['M6 4h8', 'M6 8h8', 'M6 12h8', 'M2.5 4h.5', 'M2.5 8h.5', 'M2.5 12h.5'],
  listNumber: ['M6 4h8', 'M6 8h8', 'M6 12h8', 'M2 3l1-.5V6', 'M2 10.5a.9.9 0 0 1 1.7.4L2 13h1.8'],
  quote: ['M3 4v8', 'M6 5h7', 'M6 8h7', 'M6 11h5'],
  link: ['M6.5 9.5l3-3', 'M7.5 4.5l1-1a2.5 2.5 0 0 1 3.5 3.5l-1 1', 'M8.5 11.5l-1 1A2.5 2.5 0 0 1 4 9l1-1'],
  table: ['M2 3h12v10H2z', 'M2 6.5h12', 'M2 10h12', 'M6.5 3v10'],
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

export function createToolbar(): Toolbar {
  const el = document.createElement('div');
  el.className = 'toolbar';
  el.setAttribute('role', 'toolbar');
  const row = document.createElement('div');
  row.className = 'toolbar-row';
  const extra = document.createElement('div');
  extra.className = 'toolbar-extra';
  el.append(row, extra);

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
    } else b.classList.add('icon-only');
    b.title = title;
    b.setAttribute('aria-label', label || title);
    row.appendChild(b);
    return b;
  };

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
        const on = b.getAttribute('aria-pressed') !== 'true';
        b.setAttribute('aria-pressed', String(on));
        onChange(on);
      });
      return b;
    },
    separator() {
      const s = document.createElement('span');
      s.className = 'sep';
      s.setAttribute('role', 'separator');
      row.appendChild(s);
    },
    info() {
      const i = document.createElement('span');
      i.className = 'info';
      row.appendChild(i);
      return i;
    },
    spacer() {
      const s = document.createElement('span');
      s.className = 'tb-spacer';
      row.appendChild(s);
    },
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
  };
  const step = (dir: 1 | -1) => {
    const next = dir > 0 ? STEPS.find((s) => s > zoom + 0.001) : [...STEPS].reverse().find((s) => s < zoom - 0.001);
    set(next ?? zoom);
  };
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
