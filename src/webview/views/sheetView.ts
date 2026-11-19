import type { ApplyResult, CellStyle, CellView, Edit, SheetData, SheetModel } from '../../shared/model';
import { MAX_COLS, MAX_ROWS } from '../../shared/model';
import { createFindBar, type FindTarget } from '../ui/findBar';
import { createToolbar, type Toolbar } from '../ui/toolbar';
import type { EditHandler, View } from './types';

const ROW_HEIGHT = 22;
const VIRTUAL_THRESHOLD = 1000; // rows
const BUFFER_ROWS = 20;
const MIN_COLS = 10;
const MIN_ROWS = 30;
const DEFAULT_COL_PX = 90;
const ROWHEAD_PX = 48;

export function columnName(index: number): string {
  let n = index + 1;
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function luminance(hex: string): number {
  const m = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(hex);
  if (!m) return 0.5;
  const [r, g, b] = [m[1]!, m[2]!, m[3]!].map((x) => parseInt(x, 16) / 255);
  return 0.2126 * r! + 0.7152 * g! + 0.0722 * b!;
}

function themeIsDark(): boolean {
  const c = document.body.classList;
  return c.contains('vscode-dark') || (c.contains('vscode-high-contrast') && !c.contains('vscode-high-contrast-light'));
}

/** Text color that stays readable on the current theme / the cell fill. */
export function readableColor(style: CellStyle | undefined): string | undefined {
  if (!style) return undefined;
  if (style.fill) return style.color ?? (luminance(style.fill) > 0.5 ? '#000000' : '#ffffff');
  if (!style.color) return undefined;
  const l = luminance(style.color);
  if (themeIsDark() ? l < 0.12 : l > 0.88) return undefined;
  return style.color;
}

interface Pos {
  r: number;
  c: number;
}

export function renderSheetView(container: HTMLElement, model: SheetModel, readonly: boolean, onEdit: EditHandler, tb: Toolbar = createToolbar()): View {
  container.replaceChildren();
  const root = el('div', 'sheet');
  const notice = el('div', 'banner');
  notice.hidden = true;
  const scroller = el('div', 'grid-scroll');
  scroller.tabIndex = 0;
  const table = el('table', 'grid');
  const colgroup = el('colgroup');
  const thead = el('thead');
  const tbody = el('tbody');
  table.append(colgroup, thead, tbody);
  scroller.appendChild(table);
  const tabs = el('div', 'tabs');
  tabs.setAttribute('role', 'tablist');
  const formulaBar = el('div', 'formula-bar');
  const nameBox = el('span', 'name-box');
  const fx = el('span', 'fx');
  fx.textContent = 'fx';
  const formulaText = el('span', 'formula-text');
  formulaBar.append(nameBox, fx, formulaText);
  root.append(formulaBar, notice, scroller, tabs);
  container.appendChild(root);

  let active = 0;
  let sel: Pos = { r: 0, c: 0 };
  let editing: { pos: Pos; input: HTMLInputElement } | undefined;
  let windowKey = '';
  let raf = 0;
  let filterOn = false;
  let freezeOn = false;
  let filters: string[] = [];

  const sheet = (): SheetData => model.sheets[active]!;
  const nCols = () => Math.min(MAX_COLS, Math.max(sheet().colCount, MIN_COLS));
  const nRows = () => Math.min(MAX_ROWS, Math.max(sheet().rowCount, MIN_ROWS));
  const virtual = () => nRows() > VIRTUAL_THRESHOLD;
  const cellAt = (r: number, c: number): CellView | null => sheet().rows[r]?.[c] ?? null;
  const findTd = (r: number, c: number) => tbody.querySelector<HTMLTableCellElement>(`td[data-r="${r}"][data-c="${c}"]`);

  // ---- toolbar: find, freeze first row, filter ----
  const finder = createFindBar(tb.extra, sheetFinder());
  tb.button('Find', 'Find in workbook (Ctrl+F)', () => finder.toggle());
  tb.separator();
  const freezeBtn = tb.toggle('Freeze row', 'Freeze the first row', (on) => {
    freezeOn = on;
    applyViewOptions();
  });
  const filterBtn = tb.toggle('Filter', 'Filter rows by text (the first row is the header)', (on) => {
    filterOn = on;
    filters = [];
    renderSheet();
  });
  function applyViewOptions() {
    const big = virtual();
    for (const b of [freezeBtn, filterBtn]) {
      b.disabled = big;
      b.title = big ? 'Not available for sheets with more than 1,000 rows' : b.getAttribute('aria-label') ?? '';
    }
    table.classList.toggle('freeze', freezeOn && !big);
    table.style.setProperty('--frozen-top', `${filterOn && !big ? 44 : 22}px`);
  }

  function sheetFinder(): FindTarget {
    let hits: { sheet: number; r: number; c: number }[] = [];
    return {
      find(query) {
        const q = query.toLowerCase();
        hits = [];
        model.sheets.forEach((s, si) =>
          s.rows.forEach((row, r) =>
            row.forEach((cell, c) => {
              if (cell && cell.text.toLowerCase().includes(q)) hits.push({ sheet: si, r, c });
            }),
          ),
        );
        return hits.length;
      },
      reveal(i) {
        const h = hits[i];
        if (!h) return;
        if (h.sheet !== active) {
          finishEdit(true);
          active = h.sheet;
          renderSheet();
          renderTabs();
        }
        select(h.r, h.c);
      },
      clear() {
        hits = [];
      },
    };
  }

  function applyFilter() {
    if (!filterOn) return;
    for (const tr of Array.from(tbody.rows)) {
      const r = tr.dataset['r'];
      if (r === undefined) continue; // spacer
      const row = Number(r);
      tr.hidden = row > 0 && filters.some((f, c) => f !== '' && !(cellAt(row, c)?.text ?? '').toLowerCase().includes(f));
    }
  }

  // ---- tabs ----
  function renderTabs() {
    tabs.replaceChildren();
    model.sheets.forEach((s, i) => {
      const b = el('button', i === active ? 'tab active' : 'tab');
      b.type = 'button';
      b.setAttribute('role', 'tab');
      b.setAttribute('aria-selected', String(i === active));
      b.textContent = s.name;
      b.addEventListener('click', () => {
        if (i === active) return;
        finishEdit(true);
        active = i;
        filters = [];
        sel = { r: 0, c: 0 };
        renderSheet();
        renderTabs();
      });
      tabs.appendChild(b);
    });
  }

  // ---- grid ----
  function renderSheet() {
    const s = sheet();
    notice.hidden = !s.truncated;
    notice.textContent = s.truncated ? `Large sheet: only the first ${MAX_ROWS.toLocaleString()} rows and ${MAX_COLS} columns are shown.` : '';
    const cols = nCols();
    colgroup.replaceChildren();
    thead.replaceChildren();
    const rh = el('col');
    rh.style.width = `${ROWHEAD_PX}px`;
    colgroup.appendChild(rh);
    let total = ROWHEAD_PX;
    const headRow = el('tr');
    headRow.appendChild(Object.assign(el('th', 'corner'), { textContent: '' }));
    for (let c = 0; c < cols; c++) {
      const w = s.colWidths[c];
      const px = w ? Math.round(w * 7 + 10) : DEFAULT_COL_PX;
      total += px;
      const col = el('col');
      col.style.width = `${px}px`;
      colgroup.appendChild(col);
      headRow.appendChild(Object.assign(el('th', 'colhead'), { textContent: columnName(c) }));
    }
    table.style.width = `${total}px`;
    thead.appendChild(headRow);
    if (filterOn && !virtual()) {
      const fr = el('tr', 'filters');
      fr.appendChild(el('th', 'corner'));
      for (let c = 0; c < cols; c++) {
        const th = el('th', 'filter');
        const input = el('input');
        input.type = 'text';
        input.placeholder = 'filter';
        input.value = filters[c] ?? '';
        input.setAttribute('aria-label', `Filter column ${columnName(c)}`);
        input.addEventListener('input', () => {
          filters[c] = input.value.toLowerCase();
          applyFilter();
        });
        th.appendChild(input);
        fr.appendChild(th);
      }
      thead.appendChild(fr);
    }
    applyViewOptions();
    windowKey = '';
    scroller.scrollTop = 0;
    renderBody();
  }

  function visibleRange(): [number, number] {
    const rows = nRows();
    if (!virtual()) return [0, rows];
    const h = scroller.clientHeight || 600;
    const start = Math.max(0, Math.floor(scroller.scrollTop / ROW_HEIGHT) - BUFFER_ROWS);
    const end = Math.min(rows, Math.ceil((scroller.scrollTop + h) / ROW_HEIGHT) + BUFFER_ROWS);
    return [start, end];
  }

  function renderBody() {
    const [start, end] = visibleRange();
    const key = `${active}:${start}:${end}`;
    if (key === windowKey) return;
    windowKey = key;
    finishEdit(true);
    const s = sheet();
    const cols = nCols();
    const rows = nRows();
    const useMerges = !virtual();
    const origin = new Map<string, { cs: number; rs: number }>();
    const covered = new Set<string>();
    if (useMerges) {
      for (const m of s.merges) {
        origin.set(`${m.r1},${m.c1}`, { cs: m.c2 - m.c1 + 1, rs: m.r2 - m.r1 + 1 });
        for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(`${r},${c}`);
      }
    }
    const frag = document.createDocumentFragment();
    if (start > 0) frag.appendChild(spacer(start * ROW_HEIGHT, cols + 1));
    for (let r = start; r < end; r++) {
      const tr = el('tr');
      tr.dataset['r'] = String(r);
      if (virtual()) tr.style.height = `${ROW_HEIGHT}px`;
      tr.appendChild(Object.assign(el('th', 'rowhead'), { textContent: String(r + 1) }));
      for (let c = 0; c < cols; c++) {
        const k = `${r},${c}`;
        if (covered.has(k)) continue;
        const td = el('td');
        td.dataset['r'] = String(r);
        td.dataset['c'] = String(c);
        const span = origin.get(k);
        if (span) {
          if (span.cs > 1) td.colSpan = span.cs;
          if (span.rs > 1) td.rowSpan = span.rs;
        }
        fillCell(td, cellAt(r, c));
        tr.appendChild(td);
      }
      frag.appendChild(tr);
    }
    if (end < rows) frag.appendChild(spacer((rows - end) * ROW_HEIGHT, cols + 1));
    tbody.replaceChildren(frag);
    applyFilter();
    markSelection();
  }

  function spacer(height: number, span: number): HTMLTableRowElement {
    const tr = el('tr', 'spacer');
    const td = el('td');
    td.colSpan = span;
    td.style.height = `${height}px`;
    td.style.padding = '0';
    td.style.border = '0';
    tr.appendChild(td);
    return tr;
  }

  function fillCell(td: HTMLTableCellElement, cell: CellView | null) {
    td.textContent = cell?.text ?? '';
    td.removeAttribute('title');
    td.removeAttribute('style');
    td.classList.remove('num');
    if (cell?.formula) td.title = `=${cell.formula}`;
    const st = cell?.style;
    if (st) {
      if (st.bold) td.style.fontWeight = 'bold';
      if (st.italic) td.style.fontStyle = 'italic';
      if (st.underline) td.style.textDecoration = 'underline';
      if (st.fill) td.style.backgroundColor = st.fill;
      const color = readableColor(st);
      if (color) td.style.color = color;
      if (st.align) td.style.textAlign = st.align;
    }
    if (!st?.align && cell && /^-?[\d.,]+%?$/.test(cell.text)) td.classList.add('num');
  }

  // ---- selection ----
  function markSelection() {
    for (const e of Array.from(tbody.querySelectorAll('td.sel'))) e.classList.remove('sel');
    findTd(sel.r, sel.c)?.classList.add('sel');
    updateBar();
  }

  function updateBar() {
    const cell = cellAt(sel.r, sel.c);
    nameBox.textContent = `${columnName(sel.c)}${sel.r + 1}`;
    formulaText.textContent = cell?.formula ? `=${cell.formula}` : (cell?.input ?? cell?.text ?? '');
  }

  function select(r: number, c: number, scrollTo = true) {
    sel = { r: Math.max(0, Math.min(nRows() - 1, r)), c: Math.max(0, Math.min(nCols() - 1, c)) };
    if (virtual() && scrollTo) {
      const top = sel.r * ROW_HEIGHT;
      const h = scroller.clientHeight || 600;
      if (top < scroller.scrollTop) scroller.scrollTop = top;
      else if (top + ROW_HEIGHT * 2 > scroller.scrollTop + h) scroller.scrollTop = top - h + ROW_HEIGHT * 3;
      renderBody();
    }
    markSelection();
    if (scrollTo) findTd(sel.r, sel.c)?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
  }

  // ---- editing ----
  function startEdit(pos: Pos, initial?: string) {
    if (readonly) return;
    finishEdit(true);
    const td = findTd(pos.r, pos.c);
    if (!td) return;
    const cell = cellAt(pos.r, pos.c);
    const input = el('input', 'cell-input');
    input.type = 'text';
    input.value = initial ?? cell?.input ?? cell?.text ?? '';
    input.spellcheck = false;
    td.replaceChildren(input);
    td.classList.add('editing');
    editing = { pos, input };
    input.focus();
    if (initial === undefined) input.select();
    input.addEventListener('keydown', (ev) => {
      ev.stopPropagation(); // keep the grid's own key handler out of it
      if (ev.key === 'Enter') {
        ev.preventDefault();
        finishEdit(true);
        select(pos.r + 1, pos.c);
        scroller.focus();
      } else if (ev.key === 'Tab') {
        ev.preventDefault();
        finishEdit(true);
        select(pos.r, pos.c + (ev.shiftKey ? -1 : 1));
        scroller.focus();
      } else if (ev.key === 'Escape') {
        ev.preventDefault();
        finishEdit(false);
        scroller.focus();
      }
    });
    input.addEventListener('blur', () => finishEdit(true));
  }

  function finishEdit(commit: boolean) {
    const cur = editing;
    if (!cur) return;
    editing = undefined;
    const { pos, input } = cur;
    const value = input.value;
    const td = findTd(pos.r, pos.c);
    td?.classList.remove('editing');
    const cell = cellAt(pos.r, pos.c);
    const old = cell?.input ?? cell?.text ?? '';
    if (commit && value !== old) {
      // Optimistic update; the worker's result (formatted text) replaces it in applyEdit().
      setCell(pos.r, pos.c, { ...(cell ?? { text: '' }), text: value, input: undefined, formula: undefined });
      onEdit({ kind: 'cell', sheet: active, row: pos.r, col: pos.c, value, old });
    } else if (td) {
      fillCell(td, cell);
    }
  }

  function setCell(r: number, c: number, cell: CellView | null, sheetIndex = active) {
    const rows = model.sheets[sheetIndex]!.rows;
    while (rows.length <= r) rows.push([]);
    const row = rows[r]!;
    while (row.length <= c) row.push(null);
    row[c] = cell;
    const s = model.sheets[sheetIndex]!;
    s.rowCount = Math.max(s.rowCount, r + 1);
    s.colCount = Math.max(s.colCount, c + 1);
    if (sheetIndex === active) {
      const td = findTd(r, c);
      if (td && !td.classList.contains('editing')) fillCell(td, cell);
      if (r === sel.r && c === sel.c) updateBar();
    }
  }

  // ---- events ----
  const onClick = (ev: MouseEvent) => {
    const td = (ev.target as HTMLElement).closest<HTMLTableCellElement>('td[data-r]');
    if (!td || td.classList.contains('editing')) return;
    finishEdit(true);
    select(Number(td.dataset['r']), Number(td.dataset['c']), false);
    scroller.focus();
  };
  const onDblClick = (ev: MouseEvent) => {
    const td = (ev.target as HTMLElement).closest<HTMLTableCellElement>('td[data-r]');
    if (td) startEdit({ r: Number(td.dataset['r']), c: Number(td.dataset['c']) });
  };
  const onKeyDown = (ev: KeyboardEvent) => {
    if (editing || ev.ctrlKey || ev.metaKey || ev.altKey) return;
    const moves: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    const mv = moves[ev.key];
    if (mv) {
      ev.preventDefault();
      select(sel.r + mv[0], sel.c + mv[1]);
    } else if (ev.key === 'Enter' || ev.key === 'F2') {
      ev.preventDefault();
      startEdit(sel);
    } else if (ev.key === 'Delete' || ev.key === 'Backspace') {
      ev.preventDefault();
      if (!readonly) {
        const cell = cellAt(sel.r, sel.c);
        const old = cell?.input ?? cell?.text ?? '';
        if (old !== '') {
          setCell(sel.r, sel.c, null);
          onEdit({ kind: 'cell', sheet: active, row: sel.r, col: sel.c, value: '', old });
        }
      }
    } else if (ev.key.length === 1) {
      ev.preventDefault();
      startEdit(sel, ev.key);
    }
  };
  const onScroll = () => {
    if (!virtual() || raf) return;
    raf = requestAnimationFrame(() => {
      raf = 0;
      renderBody();
    });
  };
  const onFind = (ev: KeyboardEvent) => {
    if ((ev.ctrlKey || ev.metaKey) && !ev.altKey && ev.key.toLowerCase() === 'f') {
      ev.preventDefault();
      finder.open();
    }
  };
  document.addEventListener('keydown', onFind);
  tbody.addEventListener('click', onClick);
  tbody.addEventListener('dblclick', onDblClick);
  scroller.addEventListener('keydown', onKeyDown);
  scroller.addEventListener('scroll', onScroll);

  renderTabs();
  renderSheet();
  markSelection();

  return {
    applyEdit(edit: Edit, result: ApplyResult) {
      if (edit.kind !== 'cell') return;
      setCell(edit.row, edit.col, result.cell ?? null, edit.sheet);
    },
    dispose() {
      finder.close();
      document.removeEventListener('keydown', onFind);
      tbody.removeEventListener('click', onClick);
      tbody.removeEventListener('dblclick', onDblClick);
      scroller.removeEventListener('keydown', onKeyDown);
      scroller.removeEventListener('scroll', onScroll);
      if (raf) cancelAnimationFrame(raf);
    },
  };
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className?: string): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  if (className) e.className = className;
  return e;
}
