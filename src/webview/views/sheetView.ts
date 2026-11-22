import type { ApplyResult, CellAlign, CellEdit, CellStyle, CellView, Edit, LinesEdit, SheetData, SheetModel } from '../../shared/model';
import { MAX_COLS, MAX_LINES_EDIT, MAX_ROWS, validSheetName } from '../../shared/model';
import { shiftFormula } from '../../shared/refShift';
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

/** Tab-separated text as Excel puts it on the clipboard (quoted fields may hold tabs / line breaks). */
export function parseClipboard(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let i = 0;
  const s = text.replace(/\r\n?/g, '\n');
  while (i < s.length) {
    const ch = s[i]!;
    if (ch === '"' && field === '') {
      // quoted field
      let j = i + 1;
      let v = '';
      for (;;) {
        const q = s.indexOf('"', j);
        if (q < 0) {
          v += s.slice(j);
          j = s.length;
          break;
        }
        v += s.slice(j, q);
        if (s[q + 1] === '"') {
          v += '"';
          j = q + 2;
        } else {
          j = q + 1;
          break;
        }
      }
      if (j >= s.length || s[j] === '\t' || s[j] === '\n') {
        field = v;
        i = j;
        continue;
      }
      field = '"'; // not a proper quoted field: take it literally
      i++;
      continue;
    }
    if (ch === '\t') {
      row.push(field);
      field = '';
    } else if (ch === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else field += ch;
    i++;
  }
  if (field !== '' || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

const quoteField = (v: string) => (/[\t\n"]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v);

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
  /** Status line at the right end of the tabs bar: Average, Count and Sum of the selection. */
  const status = el('span', 'sheet-status');
  status.setAttribute('aria-live', 'polite');
  const formulaBar = el('div', 'formula-bar');
  const nameBox = el('span', 'name-box');
  const fx = el('span', 'fx');
  fx.textContent = 'fx';
  const formulaText = el('span', 'formula-text');
  formulaBar.append(nameBox, fx, formulaText);
  root.append(formulaBar, notice, scroller, tabs);
  container.appendChild(root);

  let active = 0;
  const canSheets = !!model.sheetOps && !readonly;
  /** Formatting (wrap, column widths) is saved in the file; not for CSV. */
  const canFormat = canSheets;
  let sel: Pos = { r: 0, c: 0 };
  /** Other corner of the selected range (equal to `sel` for a single cell). */
  let ext: Pos = { r: 0, c: 0 };
  let dragging = false;
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
  tb.button('', 'Find in workbook (Ctrl+F)', () => finder.toggle(), 'search');
  tb.separator();
  const freezeBtn = tb.toggle('Freeze row', 'Freeze the first row', (on) => {
    freezeOn = on;
    applyViewOptions();
  }, 'freeze');
  const filterBtn = tb.toggle('Filter', 'Filter rows by text (the first row is the header)', (on) => {
    filterOn = on;
    filters = [];
    renderSheet();
  }, 'filter');
  const wrapBtn = readonly
    ? undefined
    : tb.button('Wrap', 'Wrap text in the selected cells (again: unwrap)', () => {
        if (!canFormat) return;
        const g = range();
        const wrap = !cellAt(sel.r, sel.c)?.style?.wrap;
        const edits: Edit[] = [];
        for (let r = g.r1; r <= g.r2; r++) {
          for (let c = g.c1; c <= g.c2; c++) {
            const old = !!cellAt(r, c)?.style?.wrap;
            if (old !== wrap) edits.push({ kind: 'wrap', sheet: active, row: r, col: c, wrap, old });
          }
        }
        if (edits.length > 2000) return;
        if (edits.length === 1) onEdit(edits[0]!);
        else if (edits.length) onEdit({ kind: 'batch', edits });
        scroller.focus();
      }, 'wrap');
  /** Sets the alignment of every selected cell to the next value after the active cell's. */
  const alignCells = () => {
    if (!canFormat) return;
    const g = range();
    const cur = cellAt(sel.r, sel.c)?.style;
    const H = ['left', 'center', 'right'] as const;
    const next = H[(H.indexOf(cur?.align ?? 'left') + 1) % 3]!;
    const edits: Edit[] = [];
    for (let r = g.r1; r <= g.r2; r++) {
      for (let c = g.c1; c <= g.c2; c++) {
        const st = cellAt(r, c)?.style;
        const old: CellAlign = {};
        if (st?.align) old.h = st.align;
        const align: CellAlign = { ...old };
        align.h = next;
        if (old.h !== align.h) edits.push({ kind: 'align', sheet: active, row: r, col: c, align, old });
      }
    }
    if (edits.length > 2000) return;
    if (edits.length === 1) onEdit(edits[0]!);
    else if (edits.length) onEdit({ kind: 'batch', edits });
    scroller.focus();
  };
  const alignBtns = readonly
    ? []
    : [
        tb.button('', 'Horizontal alignment: left → center → right', () => alignCells(), 'alignH'),
      ];
  if (!model.sheetOps) {
    for (const b of [wrapBtn, ...alignBtns]) {
      if (!b) continue;
      b.disabled = true;
      b.title = 'CSV files have no cell formatting';
    }
  }
  // ---- toolbar: sort, AutoSum, insert / delete rows and columns ----
  if (!readonly) {
    tb.separator();
    tb.button('', 'Sort A → Z by the active column (the selected cells, or the data around the active cell)', () => sortRange(false), 'sortAsc');
    tb.button('', 'Sort Z → A by the active column (the selected cells, or the data around the active cell)', () => sortRange(true), 'sortDesc');
    const sumBtn = tb.button('', 'AutoSum (Alt+=): add up the numbers selected, above or to the left', () => autoSum(), 'sum');
    if (!model.sheetOps) {
      sumBtn.disabled = true;
      sumBtn.title = 'CSV files have no formulas';
    }
    tb.button('', 'Insert row above the selection', () => selectionLines('row', 'insert', false), 'rowInsert');
    tb.button('', 'Delete the selected rows', () => selectionLines('row', 'delete', false), 'rowDelete');
    tb.button('', 'Insert column left of the selection', () => selectionLines('col', 'insert', false), 'colInsert');
    tb.button('', 'Delete the selected columns', () => selectionLines('col', 'delete', false), 'colDelete');
  }

  const truncatedText = () => (sheet().truncated ? `Large sheet: only the first ${MAX_ROWS.toLocaleString()} rows and ${MAX_COLS} columns are shown.` : '');
  let flashTimer: ReturnType<typeof setTimeout> | undefined;
  /** Shows a short message in the banner above the grid. */
  function flash(msg: string) {
    notice.hidden = false;
    notice.textContent = msg;
    clearTimeout(flashTimer);
    flashTimer = setTimeout(() => {
      notice.textContent = truncatedText();
      notice.hidden = !notice.textContent;
    }, 5000);
  }

  const isNum = (r: number, c: number) => numericValue(cellAt(r, c)?.text ?? '') !== undefined;
  const used = (r: number, c: number) => (cellAt(r, c)?.text ?? '') !== '';
  const cellRef = (r: number, c: number) => `${columnName(c)}${r + 1}`;

  /** The block of data around a cell (like Excel's current region), or undefined for an empty cell. */
  function region(r: number, c: number) {
    if (!used(r, c)) return undefined;
    const rows = sheet().rowCount;
    const cols = sheet().colCount;
    const b = { r1: r, r2: r, c1: c, c2: c };
    const any = (ra: number, rb: number, ca: number, cb: number) => {
      for (let i = ra; i <= rb; i++) for (let j = ca; j <= cb; j++) if (used(i, j)) return true;
      return false;
    };
    for (let changed = true; changed; ) {
      changed = false;
      const ca = Math.max(0, b.c1 - 1);
      const cb = Math.min(cols - 1, b.c2 + 1);
      const ra = Math.max(0, b.r1 - 1);
      const rb = Math.min(rows - 1, b.r2 + 1);
      if (b.r1 > 0 && any(b.r1 - 1, b.r1 - 1, ca, cb)) {
        b.r1--;
        changed = true;
      }
      if (b.r2 < rows - 1 && any(b.r2 + 1, b.r2 + 1, ca, cb)) {
        b.r2++;
        changed = true;
      }
      if (b.c1 > 0 && any(ra, rb, b.c1 - 1, b.c1 - 1)) {
        b.c1--;
        changed = true;
      }
      if (b.c2 < cols - 1 && any(ra, rb, b.c2 + 1, b.c2 + 1)) {
        b.c2++;
        changed = true;
      }
    }
    return b;
  }

  /** Writes values into single cells (the changed ones only) as one undo step; the first one is selected. */
  function putCells(list: { r: number; c: number; value: string }[]) {
    const edits: CellEdit[] = [];
    for (const { r, c, value } of list) {
      const cell = cellAt(r, c);
      const old = cell?.input ?? cell?.text ?? '';
      if (value === old) continue;
      edits.push({ kind: 'cell', sheet: active, row: r, col: c, value, old });
      setCell(r, c, value === '' ? null : { ...(cell ?? { text: '' }), text: value, input: undefined, formula: undefined });
    }
    emitCells(edits);
    return edits.length;
  }

  // ---- sort ----
  function sortRange(desc: boolean) {
    if (readonly) return;
    finishEdit(true);
    const s = sheet();
    if (s.merges.length) return flash('Sorting is not available for sheets with merged cells.');
    let g = range();
    let header = false;
    const wholeColumn = g.r1 === 0 && g.r2 >= nRows() - 1;
    if ((g.r1 === g.r2 && g.c1 === g.c2) || wholeColumn) {
      // a single cell or a whole column stands for the data block around it (sorting one column alone would tear the rows apart)
      let r = sel.r;
      if (wholeColumn) while (r < s.rowCount && !used(r, sel.c)) r++;
      const reg = region(r, sel.c);
      if (!reg) return flash('Select the cells to sort, or a cell that has data around it.');
      g = reg;
      header = true;
    }
    // a first row of text above numbers is a header and stays where it is
    if (header) {
      let first = g.r2 > g.r1;
      let numbers = false;
      for (let c = g.c1; c <= g.c2; c++) {
        first &&= used(g.r1, c) && !isNum(g.r1, c);
        numbers ||= g.r2 > g.r1 && isNum(g.r1 + 1, c);
      }
      header = first && numbers;
    }
    const r1 = g.r1 + (header ? 1 : 0);
    if (g.r2 <= r1 || (g.r2 - r1 + 1) * (g.c2 - g.c1 + 1) > 100_000) return flash(g.r2 <= r1 ? 'Nothing to sort.' : 'The selection is too large to sort.');
    const key = Math.min(g.c2, Math.max(g.c1, sel.c));
    const lines: { vals: string[]; key: string }[] = [];
    for (let r = r1; r <= g.r2; r++) {
      const vals: string[] = [];
      for (let c = g.c1; c <= g.c2; c++) {
        const cell = cellAt(r, c);
        if (cell?.formula) return flash('Sorting would break the formulas in the selection: sort values only.');
        vals.push(cell?.input ?? cell?.text ?? '');
      }
      lines.push({ vals, key: cellAt(r, key)?.text ?? '' });
    }
    const sign = desc ? -1 : 1;
    const num = (t: string) => numericValue(t);
    const sorted = lines
      .map((l, i) => ({ l, i }))
      .sort((a, b) => {
        const ka = a.l.key;
        const kb = b.l.key;
        if (ka === '' || kb === '') return ka === kb ? a.i - b.i : ka === '' ? 1 : -1; // empty cells stay last
        const na = num(ka);
        const nb = num(kb);
        let d: number;
        if (na !== undefined && nb !== undefined) d = na - nb;
        else if (na !== undefined || nb !== undefined) d = na !== undefined ? -1 : 1; // numbers before text
        else d = ka.localeCompare(kb, undefined, { sensitivity: 'base' });
        return d * sign || a.i - b.i;
      })
      .map((x) => x.l.vals);
    const edits = writeBlock(sorted, r1, g.c1);
    if (!edits.length) return flash('Already sorted.');
    emitCells(edits);
    scroller.focus();
  }

  // ---- AutoSum ----
  function autoSum() {
    if (readonly || !model.sheetOps) return;
    finishEdit(true);
    const g = range();
    const out: { r: number; c: number; value: string }[] = [];
    const sum = (ra: number, ca: number, rb: number, cb: number) => `=SUM(${cellRef(ra, ca)}:${cellRef(rb, cb)})`;
    if (g.r1 === g.r2 && g.c1 === g.c2) {
      let r = sel.r - 1;
      while (r >= 0 && isNum(r, sel.c)) r--;
      let c = sel.c - 1;
      while (c >= 0 && isNum(sel.r, c)) c--;
      if (r < sel.r - 1) out.push({ r: sel.r, c: sel.c, value: sum(r + 1, sel.c, sel.r - 1, sel.c) });
      else if (c < sel.c - 1) out.push({ r: sel.r, c: sel.c, value: sum(sel.r, c + 1, sel.r, sel.c - 1) });
    } else if (g.r1 === g.r2) {
      let any = false;
      for (let c = g.c1; c <= g.c2; c++) any ||= isNum(g.r1, c);
      if (any) out.push({ r: g.r1, c: g.c2 + 1, value: sum(g.r1, g.c1, g.r1, g.c2) });
    } else {
      const r2 = Math.min(g.r2, Math.max(g.r1, sheet().rowCount - 1));
      for (let c = g.c1; c <= g.c2; c++) {
        let any = false;
        for (let r = g.r1; r <= r2 && !any; r++) any = isNum(r, c);
        if (any) out.push({ r: r2 + 1, c, value: sum(g.r1, c, r2, c) });
      }
    }
    if (!out.length) return flash('AutoSum needs numbers: select them, or a cell below or right of them.');
    if (out.some((o) => o.r >= MAX_ROWS || o.c >= MAX_COLS)) return flash('There is no room for the sum at the edge of the sheet.');
    putCells(out);
    select(out[0]!.r, out[0]!.c);
    scroller.focus();
  }

  // ---- insert / delete rows and columns ----
  /** Applies a row / column command to the selected rows / columns (`after`: insert behind them, not before). */
  function selectionLines(axis: 'row' | 'col', op: 'insert' | 'delete', after: boolean) {
    const g = range();
    const [from, to] = axis === 'row' ? [g.r1, g.r2] : [g.c1, g.c2];
    changeLines(axis, op, after ? to + 1 : from, to - from + 1);
  }

  /** Cell edits that rewrite the formulas pointing at moved cells (the lines edit itself only moves cells). */
  function formulaFixes(lines: LinesEdit): CellEdit[] {
    const target = model.sheets[lines.sheet]!.name;
    const fixes: CellEdit[] = [];
    model.sheets.forEach((s, si) =>
      s.rows.forEach((row, r) =>
        row.forEach((cell, c) => {
          if (!cell?.formula) return;
          const next = shiftFormula(cell.formula, { target, home: s.name, axis: lines.axis, at: lines.at, count: lines.count, op: lines.op });
          if (next === cell.formula) return;
          let [nr, nc] = [r, c];
          if (si === lines.sheet) {
            const i = lines.axis === 'row' ? r : c;
            const end = lines.at + lines.count;
            const j = lines.op === 'insert' ? (i >= lines.at ? i + lines.count : i) : i < lines.at ? i : i >= end ? i - lines.count : -1;
            if (j < 0) return; // in a deleted line
            if (lines.axis === 'row') nr = j;
            else nc = j;
          }
          fixes.push({ kind: 'cell', sheet: si, row: nr, col: nc, value: '=' + next, old: cell.input ?? '=' + cell.formula });
        }),
      ),
    );
    return fixes;
  }

  function changeLines(axis: 'row' | 'col', op: 'insert' | 'delete', at: number, count: number) {
    if (readonly) return;
    finishEdit(true);
    const s = sheet();
    if (s.merges.length) return flash('Inserting and deleting rows or columns is not available for sheets with merged cells.');
    if (s.truncated) return flash('Inserting and deleting rows or columns is not available for sheets that are cut off.');
    const size = axis === 'row' ? s.rowCount : s.colCount;
    if (at >= size) return flash(op === 'delete' ? 'Nothing to delete: these cells are empty.' : 'There is no data from here on: nothing to move.');
    count = Math.min(count, op === 'delete' ? size - at : MAX_LINES_EDIT);
    if (count > MAX_LINES_EDIT) return flash(`Select at most ${MAX_LINES_EDIT} rows or columns.`);
    if (op === 'insert' && size + count > (axis === 'row' ? MAX_ROWS : MAX_COLS)) return flash('The sheet cannot get any bigger.');
    const lines: LinesEdit = { kind: 'lines', sheet: active, axis, op, at, count };
    if (op === 'delete') {
      lines.cells = [];
      for (let i = at; i < at + count; i++) {
        const line: string[] = [];
        for (let j = 0; j < (axis === 'row' ? s.colCount : s.rowCount); j++) {
          const cell = axis === 'row' ? cellAt(i, j) : cellAt(j, i);
          line.push(cell?.input ?? cell?.text ?? '');
        }
        while (line.length && line[line.length - 1] === '') line.pop();
        lines.cells.push(line);
      }
    }
    const fixes = formulaFixes(lines);
    onEdit(fixes.length ? { kind: 'batch', edits: [lines, ...fixes] } : lines);
    scroller.focus();
  }

  // ---- right-click menu on row / column headers ----
  let menuEl: HTMLElement | undefined;
  const closeMenu = () => {
    menuEl?.remove();
    menuEl = undefined;
  };
  const onContext = (ev: MouseEvent) => {
    if (readonly) return;
    const th = (ev.target as HTMLElement).closest<HTMLElement>('th.rowhead, th.colhead');
    if (!th) return;
    ev.preventDefault();
    finishEdit(true);
    closeMenu();
    const isCol = th.classList.contains('colhead');
    const idx = Number(isCol ? th.dataset['c'] : th.dataset['r']);
    const g = range();
    const wholeCols = g.r1 === 0 && g.r2 >= nRows() - 1;
    const wholeRows = g.c1 === 0 && g.c2 >= nCols() - 1;
    if (isCol ? !wholeCols || idx < g.c1 || idx > g.c2 : !wholeRows || idx < g.r1 || idx > g.r2) {
      sel = isCol ? { r: 0, c: idx } : { r: idx, c: 0 };
      ext = isCol ? { r: nRows() - 1, c: idx } : { r: idx, c: nCols() - 1 };
      markSelection();
    }
    const items: [string, () => void][] = isCol
      ? [
          ['Insert column left', () => selectionLines('col', 'insert', false)],
          ['Insert column right', () => selectionLines('col', 'insert', true)],
          ['Delete column', () => selectionLines('col', 'delete', false)],
          ['Sort A → Z', () => sortRange(false)],
          ['Sort Z → A', () => sortRange(true)],
        ]
      : [
          ['Insert row above', () => selectionLines('row', 'insert', false)],
          ['Insert row below', () => selectionLines('row', 'insert', true)],
          ['Delete row', () => selectionLines('row', 'delete', false)],
        ];
    const menu = el('div', 'ctx-menu');
    menu.setAttribute('role', 'menu');
    for (const [label, run] of items) {
      const b = el('button');
      b.type = 'button';
      b.setAttribute('role', 'menuitem');
      b.textContent = label;
      b.addEventListener('click', () => {
        closeMenu();
        run();
      });
      menu.appendChild(b);
    }
    menu.style.left = `${ev.clientX}px`;
    menu.style.top = `${ev.clientY}px`;
    root.appendChild(menu);
    menuEl = menu;
    menu.querySelector('button')?.focus();
  };
  const onMenuDown = (ev: MouseEvent) => {
    if (menuEl && !menuEl.contains(ev.target as Node)) closeMenu();
  };
  const onMenuKey = (ev: KeyboardEvent) => {
    if (!menuEl) return;
    if (ev.key === 'Escape') {
      closeMenu();
      scroller.focus();
    } else if (ev.key === 'ArrowDown' || ev.key === 'ArrowUp') {
      ev.preventDefault();
      const items = Array.from(menuEl.querySelectorAll('button'));
      const i = items.indexOf(document.activeElement as HTMLButtonElement);
      items[(i + (ev.key === 'ArrowDown' ? 1 : items.length - 1)) % items.length]?.focus();
    }
  };

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
        switchTo(i);
      });
      if (canSheets) {
        b.title = 'Double-click to rename';
        b.addEventListener('dblclick', () => renameTab(b, i));
      }
      tabs.appendChild(b);
    });
    if (canSheets) {
      const add = el('button', 'tab tab-action');
      add.type = 'button';
      add.textContent = '+';
      add.title = 'Add a sheet';
      add.setAttribute('aria-label', 'Add a sheet');
      add.addEventListener('click', addSheet);
      const del = el('button', 'tab tab-action');
      del.type = 'button';
      del.textContent = '🗑';
      del.title = `Delete sheet "${sheet().name}" (undo with Ctrl+Z)`;
      del.setAttribute('aria-label', 'Delete the current sheet');
      del.disabled = model.sheets.length <= 1;
      del.addEventListener('click', () => {
        if (model.sheets.length <= 1) return;
        finishEdit(true);
        const name = sheet().name;
        onEdit({ kind: 'sheet', op: 'remove', index: active, name });
      });
      tabs.append(add, del);
    }
    tabs.appendChild(status); // replaceChildren() above removed it
  }

  function switchTo(i: number) {
    finishEdit(true);
    active = Math.max(0, Math.min(model.sheets.length - 1, i));
    filters = [];
    sel = { r: 0, c: 0 };
    ext = { ...sel };
    renderSheet();
    renderTabs();
  }

  function uniqueName(base: string): string {
    const taken = new Set(model.sheets.map((s) => s.name.toLowerCase()));
    for (let n = 1; ; n++) {
      const name = `${base}${n}`;
      if (!taken.has(name.toLowerCase())) return name;
    }
  }

  function addSheet() {
    finishEdit(true);
    onEdit({ kind: 'sheet', op: 'add', index: model.sheets.length, name: uniqueName('Sheet') });
  }

  function renameTab(tab: HTMLButtonElement, i: number) {
    const input = el('input', 'tab-input');
    input.type = 'text';
    input.value = model.sheets[i]!.name;
    input.maxLength = 31;
    tab.replaceChildren(input);
    input.focus();
    input.select();
    let done = false;
    const finish = (ok: boolean) => {
      if (done) return;
      done = true;
      const name = input.value.trim();
      const old = model.sheets[i]!.name;
      const clash = model.sheets.some((s, k) => k !== i && s.name.toLowerCase() === name.toLowerCase());
      if (ok && name !== old && validSheetName(name) && !clash) onEdit({ kind: 'sheet', op: 'rename', index: i, name, old });
      renderTabs();
    };
    input.addEventListener('keydown', (ev) => {
      ev.stopPropagation();
      if (ev.key === 'Enter') finish(true);
      else if (ev.key === 'Escape') finish(false);
    });
    input.addEventListener('input', () => {
      const name = input.value.trim();
      const clash = model.sheets.some((s, k) => k !== i && s.name.toLowerCase() === name.toLowerCase());
      input.classList.toggle('invalid', !validSheetName(name) || clash);
      input.title = clash ? 'A sheet with this name exists' : validSheetName(name) ? '' : 'Not allowed: empty, longer than 31 characters, or one of \\ / ? * [ ] : \'';
    });
    input.addEventListener('blur', () => finish(true));
  }

  // ---- grid ----
  function renderSheet() {
    const s = sheet();
    notice.hidden = !s.truncated;
    notice.textContent = truncatedText();
    const cols = nCols();
    colgroup.replaceChildren();
    thead.replaceChildren();
    const rh = el('col');
    rh.style.width = `${ROWHEAD_PX}px`;
    colgroup.appendChild(rh);
    let total = ROWHEAD_PX;
    const headRow = el('tr');
    const corner = Object.assign(el('th', 'corner'), { textContent: '' });
    corner.title = 'Select all';
    headRow.appendChild(corner);
    for (let c = 0; c < cols; c++) {
      const w = s.colWidths[c];
      const px = w ? Math.round(w * 7 + 10) : DEFAULT_COL_PX;
      total += px;
      const col = el('col');
      col.style.width = `${px}px`;
      colgroup.appendChild(col);
      const th = Object.assign(el('th', 'colhead'), { textContent: columnName(c) });
      th.dataset['c'] = String(c);
      th.title = 'Click to select the column, drag the right edge to resize (double-click it: default width)';
      const grip = el('span', 'col-grip');
      grip.dataset['c'] = String(c);
      th.appendChild(grip);
      headRow.appendChild(th);
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
      const rh = Object.assign(el('th', 'rowhead'), { textContent: String(r + 1) });
      rh.dataset['r'] = String(r);
      tr.appendChild(rh);
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
    td.classList.toggle('wrap', !!st?.wrap);
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
  /** Selected range, normalised (r1 <= r2, c1 <= c2). */
  function range() {
    return { r1: Math.min(sel.r, ext.r), r2: Math.max(sel.r, ext.r), c1: Math.min(sel.c, ext.c), c2: Math.max(sel.c, ext.c) };
  }

  function markSelection() {
    for (const e of Array.from(tbody.querySelectorAll('td.sel, td.inrange'))) e.classList.remove('sel', 'inrange');
    const g = range();
    if (g.r1 !== g.r2 || g.c1 !== g.c2) {
      for (const td of Array.from(tbody.querySelectorAll<HTMLTableCellElement>('td[data-r]'))) {
        const r = Number(td.dataset['r']);
        const c = Number(td.dataset['c']);
        if (r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2) td.classList.add('inrange');
      }
    }
    findTd(sel.r, sel.c)?.classList.add('sel');
    updateBar();
  }

  function updateBar() {
    const cell = cellAt(sel.r, sel.c);
    const g = range();
    nameBox.textContent = `${columnName(sel.c)}${sel.r + 1}` + (g.r1 !== g.r2 || g.c1 !== g.c2 ? `  (${g.r2 - g.r1 + 1}R × ${g.c2 - g.c1 + 1}C)` : '');
    formulaText.textContent = cell?.formula ? `=${cell.formula}` : (cell?.input ?? cell?.text ?? '');
    updateStatus();
  }

  /** Cells read per status update at most (a whole-sheet selection is sampled). */
  const STATUS_MAX_CELLS = 200_000;

  /** Excel-like status line for 2+ selected cells: Average and Sum of the numeric cells, Count of the non-empty ones. */
  function updateStatus() {
    const g = range();
    if (g.r1 === g.r2 && g.c1 === g.c2) {
      status.textContent = '';
      return;
    }
    let seen = 0;
    let count = 0;
    let nums = 0;
    let sum = 0;
    scan: for (let r = g.r1; r <= g.r2; r++) {
      for (let c = g.c1; c <= g.c2; c++) {
        if (seen++ >= STATUS_MAX_CELLS) break scan;
        const text = cellAt(r, c)?.text ?? '';
        if (text === '') continue;
        count++;
        const v = numericValue(text);
        if (v !== undefined) {
          sum += v;
          nums++;
        }
      }
    }
    const fmt = (n: number) => n.toLocaleString(undefined, { maximumFractionDigits: 10 });
    status.textContent = nums ? `Average: ${fmt(sum / nums)}   Count: ${count}   Sum: ${fmt(sum)}` : `Count: ${count}`;
  }

  function clampPos(r: number, c: number): Pos {
    return { r: Math.max(0, Math.min(nRows() - 1, r)), c: Math.max(0, Math.min(nCols() - 1, c)) };
  }

  /** Moves the selection; with `extend` only the far corner of the range moves. */
  function select(r: number, c: number, scrollTo = true, extend = false) {
    const p = clampPos(r, c);
    if (extend) ext = p;
    else {
      sel = p;
      ext = { ...p };
    }
    const focus = extend ? ext : sel;
    if (virtual() && scrollTo) {
      const top = focus.r * ROW_HEIGHT;
      const h = scroller.clientHeight || 600;
      if (top < scroller.scrollTop) scroller.scrollTop = top;
      else if (top + ROW_HEIGHT * 2 > scroller.scrollTop + h) scroller.scrollTop = top - h + ROW_HEIGHT * 3;
      renderBody();
    }
    markSelection();
    if (scrollTo) findTd(focus.r, focus.c)?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
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
  const onMouseDown = (ev: MouseEvent) => {
    if (ev.button !== 0) return;
    const td = (ev.target as HTMLElement).closest<HTMLTableCellElement>('td[data-r]');
    if (!td || td.classList.contains('editing')) return;
    ev.preventDefault(); // no text selection while dragging over cells
    finishEdit(true);
    select(Number(td.dataset['r']), Number(td.dataset['c']), false, ev.shiftKey);
    dragging = true;
    scroller.focus();
  };
  const onClick = (ev: MouseEvent) => {
    // same as mousedown (keyboard / assistive clicks come without one); repeating it is harmless
    const td = (ev.target as HTMLElement).closest<HTMLTableCellElement>('td[data-r]');
    if (!td || td.classList.contains('editing')) return;
    finishEdit(true);
    select(Number(td.dataset['r']), Number(td.dataset['c']), false, ev.shiftKey);
    scroller.focus();
  };
  /** Column resize by dragging the grip at the right edge of a column header. */
  let resizing: { c: number; x: number; px: number; col: HTMLTableColElement } | undefined;
  const pxToChars = (px: number) => Math.max(1, Math.round(((px - 10) / 7) * 10) / 10);
  const onHeadDown = (ev: MouseEvent) => {
    if (ev.button !== 0) return;
    const t = ev.target as HTMLElement;
    if (t.classList.contains('col-grip')) {
      ev.preventDefault();
      ev.stopPropagation();
      const c = Number(t.dataset['c']);
      const col = colgroup.children[c + 1] as HTMLTableColElement;
      resizing = { c, x: ev.clientX, px: parseFloat(col.style.width) || DEFAULT_COL_PX, col };
      document.body.classList.add('col-resizing');
      return;
    }
    const th = t.closest<HTMLElement>('th');
    if (!th) return;
    finishEdit(true);
    const last = { r: nRows() - 1, c: nCols() - 1 };
    if (th.classList.contains('colhead')) {
      const c = Number(th.dataset['c']);
      if (ev.shiftKey) ext = { r: last.r, c };
      else {
        sel = { r: 0, c };
        ext = { r: last.r, c };
      }
    } else if (th.classList.contains('corner') && !th.closest('tr.filters')) {
      sel = { r: 0, c: 0 };
      ext = last;
    } else return;
    ev.preventDefault();
    markSelection();
    scroller.focus();
  };
  const onRowHeadDown = (ev: MouseEvent) => {
    const th = (ev.target as HTMLElement).closest<HTMLElement>('th.rowhead');
    if (!th || ev.button !== 0) return;
    ev.preventDefault();
    finishEdit(true);
    const r = Number(th.dataset['r']);
    const lastC = nCols() - 1;
    if (ev.shiftKey) ext = { r, c: lastC };
    else {
      sel = { r, c: 0 };
      ext = { r, c: lastC };
    }
    markSelection();
    scroller.focus();
  };
  const onResizeMove = (ev: MouseEvent) => {
    if (!resizing) return;
    const px = Math.max(24, Math.round(resizing.px + ev.clientX - resizing.x));
    const diff = px - (parseFloat(resizing.col.style.width) || px);
    resizing.col.style.width = `${px}px`;
    table.style.width = `${(parseFloat(table.style.width) || 0) + diff}px`;
  };
  const commitWidth = (c: number, width: number | null) => {
    const s = sheet();
    const old = s.colWidths[c] ?? null;
    if (old === width) return;
    while (s.colWidths.length <= c) s.colWidths.push(null);
    s.colWidths[c] = width;
    if (canFormat) onEdit({ kind: 'colwidth', sheet: active, col: c, width, old });
  };
  const onResizeUp = () => {
    if (!resizing) return;
    const { c, col } = resizing;
    resizing = undefined;
    document.body.classList.remove('col-resizing');
    commitWidth(c, pxToChars(parseFloat(col.style.width)));
  };
  const onHeadDbl = (ev: MouseEvent) => {
    const t = ev.target as HTMLElement;
    if (!t.classList.contains('col-grip')) return;
    commitWidth(Number(t.dataset['c']), null);
    renderSheet();
  };
  const onMouseOver = (ev: MouseEvent) => {
    if (!dragging) return;
    const td = (ev.target as HTMLElement).closest<HTMLTableCellElement>('td[data-r]');
    if (td) select(Number(td.dataset['r']), Number(td.dataset['c']), false, true);
  };
  const onMouseUp = () => {
    dragging = false;
  };

  /** Cell edits for a block of values written at the top-left of the selection. */
  function writeBlock(values: string[][], r0: number, c0: number): CellEdit[] {
    const edits: CellEdit[] = [];
    values.forEach((row, dr) =>
      row.forEach((value, dc) => {
        const r = r0 + dr;
        const c = c0 + dc;
        if (r >= MAX_ROWS || c >= MAX_COLS) return;
        const cell = cellAt(r, c);
        const old = cell?.input ?? cell?.text ?? '';
        if (value === old) return;
        edits.push({ kind: 'cell', sheet: active, row: r, col: c, value, old });
        setCell(r, c, value === '' ? null : { ...(cell ?? { text: '' }), text: value, input: undefined, formula: undefined });
      }),
    );
    return edits;
  }

  function emitCells(edits: CellEdit[]) {
    if (edits.length === 1) onEdit(edits[0]!);
    else if (edits.length > 1) onEdit({ kind: 'batch', edits });
  }

  function clearRange() {
    if (readonly) return;
    const g = range();
    const blank = Array.from({ length: g.r2 - g.r1 + 1 }, () => Array.from({ length: g.c2 - g.c1 + 1 }, () => ''));
    emitCells(writeBlock(blank, g.r1, g.c1));
  }

  function rangeText(): string {
    const g = range();
    const lines: string[] = [];
    for (let r = g.r1; r <= g.r2; r++) {
      const vals: string[] = [];
      for (let c = g.c1; c <= g.c2; c++) vals.push(quoteField(cellAt(r, c)?.text ?? ''));
      lines.push(vals.join('\t'));
    }
    return lines.join('\n') + (g.r1 !== g.r2 || g.c1 !== g.c2 ? '\n' : '');
  }

  const gridActive = () => !editing && (document.activeElement === scroller || scroller.contains(document.activeElement));
  const onCopy = (ev: ClipboardEvent) => {
    if (!gridActive() || !ev.clipboardData) return;
    ev.preventDefault();
    ev.clipboardData.setData('text/plain', rangeText());
    if (ev.type === 'cut') clearRange();
  };
  const onPaste = (ev: ClipboardEvent) => {
    if (!gridActive() || readonly || !ev.clipboardData) return;
    ev.preventDefault();
    let values = parseClipboard(ev.clipboardData.getData('text/plain'));
    if (values.length === 0) return;
    const g = range();
    // one value pasted into a selected range fills the range
    if (values.length === 1 && values[0]!.length === 1 && (g.r1 !== g.r2 || g.c1 !== g.c2)) {
      const v = values[0]![0]!;
      values = Array.from({ length: g.r2 - g.r1 + 1 }, () => Array.from({ length: g.c2 - g.c1 + 1 }, () => v));
    }
    emitCells(writeBlock(values, g.r1, g.c1));
    const rows = values.length;
    const cols = Math.max(...values.map((r) => r.length));
    sel = clampPos(g.r1, g.c1);
    ext = clampPos(g.r1 + rows - 1, g.c1 + cols - 1);
    markSelection();
  };
  const onDblClick = (ev: MouseEvent) => {
    const td = (ev.target as HTMLElement).closest<HTMLTableCellElement>('td[data-r]');
    if (td) startEdit({ r: Number(td.dataset['r']), c: Number(td.dataset['c']) });
  };
  const onKeyDown = (ev: KeyboardEvent) => {
    if (editing) return;
    if ((ev.ctrlKey || ev.metaKey) && !ev.altKey && ev.key.toLowerCase() === 'a') {
      ev.preventDefault();
      sel = { r: 0, c: 0 };
      ext = clampPos(Math.max(0, sheet().rowCount - 1), Math.max(0, sheet().colCount - 1));
      markSelection();
      return;
    }
    if (ev.altKey && !ev.ctrlKey && !ev.metaKey && (ev.key === '=' || ev.key === '+')) {
      ev.preventDefault();
      autoSum();
      return;
    }
    if (ev.ctrlKey || ev.metaKey || ev.altKey) return;
    const moves: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    const mv = moves[ev.key];
    if (mv) {
      ev.preventDefault();
      if (ev.shiftKey) select(ext.r + mv[0], ext.c + mv[1], true, true);
      else select(sel.r + mv[0], sel.c + mv[1]);
    } else if (ev.key === 'Tab') {
      ev.preventDefault();
      select(sel.r, sel.c + (ev.shiftKey ? -1 : 1));
    } else if (ev.key === 'Enter' || ev.key === 'F2') {
      ev.preventDefault();
      startEdit(sel);
    } else if (ev.key === 'Delete' || ev.key === 'Backspace') {
      ev.preventDefault();
      clearRange();
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
  document.addEventListener('keydown', onMenuKey);
  document.addEventListener('mousedown', onMenuDown);
  thead.addEventListener('contextmenu', onContext);
  tbody.addEventListener('contextmenu', onContext);
  document.addEventListener('copy', onCopy);
  document.addEventListener('cut', onCopy);
  document.addEventListener('paste', onPaste);
  document.addEventListener('mouseup', onMouseUp);
  tbody.addEventListener('mousedown', onMouseDown);
  tbody.addEventListener('mousedown', onRowHeadDown);
  thead.addEventListener('mousedown', onHeadDown);
  thead.addEventListener('dblclick', onHeadDbl);
  document.addEventListener('mousemove', onResizeMove);
  document.addEventListener('mouseup', onResizeUp);
  tbody.addEventListener('click', onClick);
  tbody.addEventListener('mouseover', onMouseOver);
  tbody.addEventListener('dblclick', onDblClick);
  scroller.addEventListener('keydown', onKeyDown);
  scroller.addEventListener('scroll', onScroll);

  renderTabs();
  renderSheet();
  markSelection();

  const applyChange = (edit: Edit, result: ApplyResult) => {
    if (edit.kind === 'sheet') {
      finishEdit(false);
      if (edit.op === 'rename') {
        const s = model.sheets[edit.index];
        if (s) s.name = edit.name;
        renderTabs();
        return;
      }
      if (edit.op === 'remove') {
        model.sheets.splice(edit.index, 1);
        switchTo(active > edit.index || active >= model.sheets.length ? active - 1 : active);
        return;
      }
      if (result.sheet) model.sheets.splice(edit.index, 0, result.sheet);
      switchTo(edit.index);
      return;
    }
    if (edit.kind === 'lines') {
      finishEdit(false);
      closeMenu();
      if (result.sheet) model.sheets[edit.sheet] = result.sheet;
      if (edit.sheet === active) {
        const top = scroller.scrollTop;
        renderSheet();
        scroller.scrollTop = top;
        sel = clampPos(sel.r, sel.c);
        ext = clampPos(ext.r, ext.c);
        markSelection();
      }
      for (const o of result.others ?? []) setCell(o.row, o.col, o.cell, o.sheet);
      return;
    }
    if (edit.kind === 'colwidth') {
      const s = model.sheets[edit.sheet];
      if (!s) return;
      while (s.colWidths.length <= edit.col) s.colWidths.push(null);
      const changed = s.colWidths[edit.col] !== edit.width;
      s.colWidths[edit.col] = edit.width;
      if (changed && edit.sheet === active) renderSheet(); // undo / redo
      return;
    }
    if (edit.kind === 'wrap' || edit.kind === 'align') {
      setCell(edit.row, edit.col, result.cell ?? null, edit.sheet);
      return;
    }
    if (edit.kind !== 'cell') return;
    setCell(edit.row, edit.col, result.cell ?? null, edit.sheet);
    for (const o of result.others ?? []) setCell(o.row, o.col, o.cell, o.sheet);
  };

  return {
    applyEdit(edit: Edit, result: ApplyResult) {
      applyChange(edit, result);
      updateStatus();
    },
    dispose() {
      finder.close();
      document.removeEventListener('keydown', onFind);
      document.removeEventListener('keydown', onMenuKey);
      document.removeEventListener('mousedown', onMenuDown);
      thead.removeEventListener('contextmenu', onContext);
      tbody.removeEventListener('contextmenu', onContext);
      closeMenu();
      clearTimeout(flashTimer);
      document.removeEventListener('copy', onCopy);
      document.removeEventListener('cut', onCopy);
      document.removeEventListener('paste', onPaste);
      document.removeEventListener('mouseup', onMouseUp);
      tbody.removeEventListener('mousedown', onMouseDown);
      tbody.removeEventListener('mousedown', onRowHeadDown);
      thead.removeEventListener('mousedown', onHeadDown);
      thead.removeEventListener('dblclick', onHeadDbl);
      document.removeEventListener('mousemove', onResizeMove);
      document.removeEventListener('mouseup', onResizeUp);
      tbody.removeEventListener('click', onClick);
      tbody.removeEventListener('mouseover', onMouseOver);
      tbody.removeEventListener('dblclick', onDblClick);
      scroller.removeEventListener('keydown', onKeyDown);
      scroller.removeEventListener('scroll', onScroll);
      if (raf) cancelAnimationFrame(raf);
    },
  };
}

/** The number a displayed cell text stands for ("1,234.5", "-2", "50%" = 0.5 as in Excel), or undefined for text. */
export function numericValue(text: string): number | undefined {
  const t = text.trim();
  if (!/^-?[\d,]*\d(\.\d+)?%?$/.test(t)) return undefined;
  const v = Number(t.replace(/[,%]/g, ''));
  return t.endsWith('%') ? v / 100 : v;
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className?: string): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  if (className) e.className = className;
  return e;
}
