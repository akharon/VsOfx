/** A small button bar above the document. Plain text labels only (no icon fonts, no remote resources). */
export interface Toolbar {
  /** Root element; contains the button row and a slot for panels such as the find bar. */
  el: HTMLElement;
  /** Panel slot below the button row. */
  extra: HTMLElement;
  button(label: string, title: string, onClick: () => void): HTMLButtonElement;
  /** A button that stays pressed; `onChange` receives the new state. */
  toggle(label: string, title: string, onChange: (on: boolean) => void): HTMLButtonElement;
  separator(): void;
  /** Right-aligned status text (word count, zoom ...). */
  info(): HTMLElement;
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

  const make = (label: string, title: string): HTMLButtonElement => {
    const b = document.createElement('button');
    b.type = 'button';
    b.textContent = label;
    b.title = title;
    b.setAttribute('aria-label', title);
    row.appendChild(b);
    return b;
  };

  return {
    el,
    extra,
    button(label, title, onClick) {
      const b = make(label, title);
      b.addEventListener('click', onClick);
      return b;
    },
    toggle(label, title, onChange) {
      const b = make(label, title);
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
  };
}

const ZOOM_MIN = 0.6;
const ZOOM_MAX = 2;

/** Adds "A-", "A+" and a percentage reset button; `apply` gets the factor (1 = 100 %). */
export function addZoom(tb: Toolbar, apply: (factor: number) => void): void {
  let zoom = 1;
  const reset = tb.button('100%', 'Reset zoom', () => set(1));
  const set = (z: number) => {
    zoom = Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, Math.round(z * 10) / 10));
    reset.textContent = `${Math.round(zoom * 100)}%`;
    apply(zoom);
  };
  const out = tb.button('A−', 'Zoom out', () => set(zoom - 0.1));
  const inn = tb.button('A+', 'Zoom in', () => set(zoom + 0.1));
  // keep the order: A-, 100%, A+
  reset.before(out);
  reset.after(inn);
}
