import type { ApplyResult, CellEdit, CellView, Edit, MergeRange, SheetData, SheetModel } from '../../shared/model';
import { MAX_COLS, MAX_ROWS } from '../../shared/model';
import type { FormatAdapter } from './Adapter';
import { isEl, MIME_ODS, NS, odfText, openOdf, saveOdf, setOdfText, type OdfPackage } from './odf';
import { descendantsNS } from './xml';
import { Calc, excelToOdf, FormulaError, odfToExcel, UNSUPPORTED, type Book, type Scalar } from '../../shared/formula';
import { general } from './numFmt';

const NUMBER = /^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/;
const LEADING_ZERO = /^[-+]?0\d/;
const ROW_REPEAT = 'number-rows-repeated';
const COL_REPEAT = 'number-columns-repeated';
const VALUE_ATTRS: [string, string][] = [
  [NS.office, 'value'],
  [NS.office, 'string-value'],
  [NS.office, 'date-value'],
  [NS.office, 'time-value'],
  [NS.office, 'boolean-value'],
  [NS.table, 'formula'],
  [NS.calcext, 'value-type'],
];

interface Original {
  input: string;
  snapshot: Element;
}

/** .ods: cell values of content.xml; formulas are shown with their cached results. */
export class OdsAdapter implements FormatAdapter {
  readonly readonly = false;
  private pkg!: OdfPackage;
  private tables: Element[] = [];
  private readonly originals = new Map<string, Original>();

  async load(bytes: Uint8Array, edits: Edit[]): Promise<SheetModel> {
    this.pkg = openOdf(bytes, MIME_ODS);
    const sheet = descendantsNS(this.pkg.content, NS.office, 'spreadsheet')[0];
    if (!sheet) throw new Error('This OpenDocument file has no spreadsheet body');
    this.tables = Array.from(sheet.childNodes).filter((n): n is Element => isEl(n, NS.table, 'table'));
    for (const e of edits) this.applyEdit(e);
    return { type: 'sheet', sheets: this.tables.map((t) => this.sheetData(t)) };
  }

  applyEdit(edit: Edit): ApplyResult {
    if (edit.kind !== 'cell') throw new Error('Not a cell edit');
    const table = this.tables[edit.sheet];
    if (!table) throw new Error(`Sheet ${edit.sheet} does not exist`);
    const row = locateRow(this.pkg.content, table, edit.row);
    const cell = locateCell(this.pkg.content, row, edit.col);
    const key = `${edit.sheet}:${edit.row}:${edit.col}`;
    let orig = this.originals.get(key);
    if (!orig) {
      const view = cellView(cell);
      orig = { input: view?.input ?? view?.text ?? '', snapshot: cell.cloneNode(true) as Element };
      this.originals.set(key, orig);
    }
    if (edit.value === orig.input) restore(cell, orig.snapshot);
    else writeCell(this.pkg.content, cell, edit);
    const others = this.recalc().filter((o) => !(o.sheet === edit.sheet && o.row === edit.row && o.col === edit.col));
    const res: ApplyResult = { cell: cellView(cell) };
    if (others.length) res.others = others;
    return res;
  }

  /** Recomputes all formulas of the workbook; returns the cells whose result changed. */
  private recalc(): NonNullable<ApplyResult['others']> {
    const grids = this.tables.map(indexCells);
    const book: Book = {
      sheetIndex: (name) => this.tables.findIndex((t) => (t.getAttributeNS(NS.table, 'name') ?? '').toLowerCase() === name.toLowerCase()),
      value: (s, r, c) => {
        const el = grids[s]?.cells.get(`${r}:${c}`);
        return el ? valueOf(el) : null;
      },
      formula: (s, r, c) => {
        const el = grids[s]?.cells.get(`${r}:${c}`);
        const f = el?.getAttributeNS(NS.table, 'formula');
        return f && repeatOf(el!, COL_REPEAT) === 1 ? odfToExcel(f) : undefined;
      },
      size: (s) => ({ rows: grids[s]?.rows ?? 0, cols: grids[s]?.cols ?? 0 }),
    };
    const calc = new Calc(book);
    const changed: NonNullable<ApplyResult['others']> = [];
    grids.forEach((g, si) => {
      for (const [key, el] of g.cells) {
        const f = el.getAttributeNS(NS.table, 'formula');
        if (!f || repeatOf(el, COL_REPEAT) !== 1) continue;
        const [r, c] = key.split(':').map(Number) as [number, number];
        const v = calc.cell(si, r, c);
        if (v === UNSUPPORTED) continue;
        if (setResult(this.pkg.content, el, v) && r < MAX_ROWS && c < MAX_COLS) changed.push({ sheet: si, row: r, col: c, cell: cellView(el) });
      }
    });
    return changed;
  }

  async save(): Promise<Uint8Array> {
    return saveOdf(this.pkg);
  }

  private sheetData(table: Element): SheetData {
    const rows: (CellView | null)[][] = [];
    const merges: MergeRange[] = [];
    let r = 0;
    let lastContent = -1;
    let colCount = 0;
    let truncated = false;
    for (const rowEl of rowElements(table)) {
      const rep = repeatOf(rowEl, ROW_REPEAT);
      const cells = rowCells(rowEl, r, merges);
      const hasContent = cells.some((c) => c !== null);
      if (hasContent) {
        colCount = Math.max(colCount, cells.length);
        for (let k = 0; k < rep; k++) {
          if (r + k >= MAX_ROWS) {
            truncated = true;
            break;
          }
          rows[r + k] = cells;
          lastContent = r + k;
        }
      }
      r += rep;
    }
    const rowCount = lastContent + 1;
    const out = Array.from({ length: rowCount }, (_, i) => rows[i] ?? []);
    return {
      name: table.getAttributeNS(NS.table, 'name') ?? 'Sheet',
      rows: out,
      rowCount,
      colCount: Math.min(colCount, MAX_COLS),
      merges: merges.filter((m) => m.r1 < rowCount),
      colWidths: Array.from({ length: Math.min(colCount, MAX_COLS) }, () => null),
      truncated: truncated || colCount > MAX_COLS,
    };
  }
}

// ---- DOM helpers ---------------------------------------------------------

function repeatOf(el: Element, attr: string): number {
  const n = Number(el.getAttributeNS(NS.table, attr) || 1);
  return Number.isFinite(n) && n >= 1 ? Math.floor(n) : 1;
}

/** table:table-row elements in order, looking into row groups/header rows. */
function rowElements(parent: Element): Element[] {
  const out: Element[] = [];
  for (let c = parent.firstChild; c; c = c.nextSibling) {
    if (isEl(c, NS.table, 'table-row')) out.push(c as Element);
    else if (isEl(c, NS.table, 'table-row-group') || isEl(c, NS.table, 'table-header-rows') || isEl(c, NS.table, 'table-rows')) {
      out.push(...rowElements(c as Element));
    }
  }
  return out;
}

function cellElements(row: Element): Element[] {
  return Array.from(row.childNodes).filter((n): n is Element => isEl(n, NS.table, 'table-cell') || isEl(n, NS.table, 'covered-table-cell'));
}

function isEmptyCell(el: Element): boolean {
  return (
    !el.getAttributeNS(NS.office, 'value-type') &&
    !Array.from(el.childNodes).some((n) => isEl(n, NS.text, 'p'))
  );
}

function rowCells(rowEl: Element, rowIndex: number, merges: MergeRange[]): (CellView | null)[] {
  const out: (CellView | null)[] = [];
  let c = 0;
  for (const cell of cellElements(rowEl)) {
    const rep = repeatOf(cell, COL_REPEAT);
    if (!isEmptyCell(cell) && cell.localName === 'table-cell') {
      const view = cellView(cell);
      const cs = Number(cell.getAttributeNS(NS.table, 'number-columns-spanned') || 1);
      const rs = Number(cell.getAttributeNS(NS.table, 'number-rows-spanned') || 1);
      if (cs > 1 || rs > 1) merges.push({ r1: rowIndex, c1: c, r2: rowIndex + rs - 1, c2: c + cs - 1 });
      for (let k = 0; k < rep && c + k < MAX_COLS; k++) {
        while (out.length < c + k) out.push(null);
        out[c + k] = view;
      }
    }
    c += rep;
  }
  return out;
}

function cellView(cell: Element): CellView | null {
  const paras = Array.from(cell.childNodes).filter((n): n is Element => isEl(n, NS.text, 'p'));
  let text = paras.map(odfText).join('\n');
  const type = cell.getAttributeNS(NS.office, 'value-type');
  if (!text && type) {
    text =
      cell.getAttributeNS(NS.office, 'value') ||
      cell.getAttributeNS(NS.office, 'string-value') ||
      cell.getAttributeNS(NS.office, 'date-value') ||
      cell.getAttributeNS(NS.office, 'boolean-value') ||
      '';
  }
  if (!text && !type) return null;
  const view: CellView = { text };
  const formula = cell.getAttributeNS(NS.table, 'formula');
  if (formula) {
    view.formula = odfToExcel(formula);
    view.input = '=' + view.formula;
  }
  return view;
}

/** Returns an element that represents exactly one repetition, splitting a repeated element if needed. */
function splitRepeated(el: Element, attr: string, k: number): Element {
  const n = repeatOf(el, attr);
  if (n === 1) return el;
  const parent = el.parentNode!;
  const set = (e: Element, count: number) => {
    if (count === 1) e.removeAttributeNS(NS.table, attr);
    else e.setAttributeNS(NS.table, `${e.prefix ? e.prefix.split(':')[0] : 'table'}:${attr}`, String(count));
  };
  if (k > 0) {
    const before = el.cloneNode(true) as Element;
    set(before, k);
    parent.insertBefore(before, el);
  }
  if (k < n - 1) {
    const after = el.cloneNode(true) as Element;
    set(after, n - k - 1);
    parent.insertBefore(after, el.nextSibling);
  }
  set(el, 1);
  return el;
}

function newElement(doc: Document, like: Element, local: string, repeatAttr?: string, repeat = 1): Element {
  const prefix = like.prefix ?? 'table';
  const el = doc.createElementNS(NS.table, `${prefix}:${local}`);
  if (repeatAttr && repeat > 1) el.setAttributeNS(NS.table, `${prefix}:${repeatAttr}`, String(repeat));
  return el;
}

function locateRow(doc: Document, table: Element, r: number): Element {
  let pos = 0;
  for (const rowEl of rowElements(table)) {
    const rep = repeatOf(rowEl, ROW_REPEAT);
    if (r < pos + rep) return splitRepeated(rowEl, ROW_REPEAT, r - pos);
    pos += rep;
  }
  if (r > pos) {
    const gap = newElement(doc, table, 'table-row', ROW_REPEAT, r - pos);
    gap.appendChild(newElement(doc, table, 'table-cell'));
    table.appendChild(gap);
  }
  const row = newElement(doc, table, 'table-row');
  row.appendChild(newElement(doc, table, 'table-cell'));
  table.appendChild(row);
  return row;
}

function locateCell(doc: Document, row: Element, c: number): Element {
  let pos = 0;
  for (const cell of cellElements(row)) {
    const rep = repeatOf(cell, COL_REPEAT);
    if (c < pos + rep) {
      if (cell.localName === 'covered-table-cell') throw new Error('This cell is part of a merged cell and cannot be edited');
      return splitRepeated(cell, COL_REPEAT, c - pos);
    }
    pos += rep;
  }
  if (c > pos) row.appendChild(newElement(doc, row, 'table-cell', COL_REPEAT, c - pos));
  const cell = newElement(doc, row, 'table-cell');
  row.appendChild(cell);
  return cell;
}

function restore(cell: Element, snapshot: Element): void {
  while (cell.firstChild) cell.removeChild(cell.firstChild);
  for (const a of Array.from(cell.attributes)) cell.removeAttribute(a.name);
  for (const a of Array.from(snapshot.attributes)) cell.setAttributeNS(a.namespaceURI, a.name, a.value);
  for (const n of Array.from(snapshot.childNodes)) cell.appendChild(n.cloneNode(true));
}

/** The prefix the document already uses for a namespace. */
function prefixOf(doc: Document, ns: string, fallback: string): string {
  for (const a of Array.from(doc.documentElement.attributes)) {
    if (a.value === ns && a.name.startsWith('xmlns:')) return a.name.slice(6);
  }
  return fallback;
}

/** Cell elements by "row:col" (repeated cells/rows expanded within the display limits). */
function indexCells(table: Element): { cells: Map<string, Element>; rows: number; cols: number } {
  const cells = new Map<string, Element>();
  let r = 0;
  let cols = 0;
  let rows = 0;
  for (const rowEl of rowElements(table)) {
    const rrep = repeatOf(rowEl, ROW_REPEAT);
    let c = 0;
    const entries: [number, Element][] = [];
    for (const cell of cellElements(rowEl)) {
      const crep = repeatOf(cell, COL_REPEAT);
      if (!isEmptyCell(cell) || cell.getAttributeNS(NS.table, 'formula')) {
        for (let k = 0; k < crep && c + k < MAX_COLS; k++) entries.push([c + k, cell]);
      }
      c += crep;
    }
    if (entries.length) {
      for (let k = 0; k < rrep && r + k < MAX_ROWS; k++) {
        for (const [col, el] of entries) cells.set(`${r + k}:${col}`, el);
        rows = r + k + 1;
      }
      cols = Math.max(cols, entries[entries.length - 1]![0] + 1);
    }
    r += rrep;
  }
  return { cells, rows, cols };
}

const DAY_MS = 86400000;
const EPOCH = Date.UTC(1899, 11, 30);

/** Cell value for the formula engine (cached result for formula cells). */
function valueOf(cell: Element): Scalar {
  const type = cell.getAttributeNS(NS.office, 'value-type');
  const calcType = cell.getAttributeNS(NS.calcext, 'value-type');
  const text = () => Array.from(cell.childNodes).filter((n): n is Element => isEl(n, NS.text, 'p')).map(odfText).join('\n');
  if (calcType === 'error') return new FormulaError(text() || '#VALUE!');
  switch (type) {
    case 'float':
    case 'percentage':
    case 'currency':
      return Number(cell.getAttributeNS(NS.office, 'value'));
    case 'boolean':
      return cell.getAttributeNS(NS.office, 'boolean-value') === 'true';
    case 'date': {
      const d = cell.getAttributeNS(NS.office, 'date-value') ?? '';
      const t = Date.parse(/T/.test(d) ? d + (/Z|[+-]\d\d:\d\d$/.test(d) ? '' : 'Z') : d + 'T00:00:00Z');
      return Number.isFinite(t) ? (t - EPOCH) / DAY_MS : null;
    }
    case 'time': {
      const m = /^-?PT(?:(\d+)H)?(?:(\d+)M)?(?:([\d.]+)S)?$/.exec(cell.getAttributeNS(NS.office, 'time-value') ?? '');
      return m ? (Number(m[1] ?? 0) * 3600 + Number(m[2] ?? 0) * 60 + Number(m[3] ?? 0)) / 86400 : null;
    }
    case 'string':
      return cell.getAttributeNS(NS.office, 'string-value') ?? text();
    default: {
      const t = text();
      return t === '' ? null : t;
    }
  }
}

/** Stores a computed result as the cached value of a formula cell; false if it did not change. */
function setResult(doc: Document, cell: Element, v: Scalar): boolean {
  const office = prefixOf(doc, NS.office, 'office');
  const calcext = prefixOf(doc, NS.calcext, '');
  const oldType = cell.getAttributeNS(NS.office, 'value-type');
  let type: string;
  let display: string;
  const attrs: [string, string][] = [];
  if (v instanceof FormulaError) {
    type = 'string';
    display = v.code;
  } else if (typeof v === 'boolean') {
    type = 'boolean';
    display = v ? 'TRUE' : 'FALSE';
    attrs.push(['boolean-value', String(v)]);
  } else if (typeof v === 'string') {
    type = 'string';
    display = v;
  } else {
    const n = v ?? 0;
    type = oldType === 'percentage' || oldType === 'currency' ? oldType : 'float';
    display = type === 'percentage' ? `${general(n * 100)}%` : general(n);
    attrs.push(['value', String(n)]);
  }
  const paras = Array.from(cell.childNodes).filter((n): n is Element => isEl(n, NS.text, 'p'));
  const oldDisplay = paras.map(odfText).join('\n');
  const same =
    oldType === type &&
    oldDisplay === display &&
    attrs.every(([k, val]) => cell.getAttributeNS(NS.office, k) === val) &&
    (v instanceof FormulaError) === (cell.getAttributeNS(NS.calcext, 'value-type') === 'error');
  if (same) return false;
  for (const [ns, local] of VALUE_ATTRS) if (local !== 'formula') cell.removeAttributeNS(ns, local);
  cell.setAttributeNS(NS.office, `${office}:value-type`, type);
  for (const [k, val] of attrs) cell.setAttributeNS(NS.office, `${office}:${k}`, val);
  if (calcext) cell.setAttributeNS(NS.calcext, `${calcext}:value-type`, v instanceof FormulaError ? 'error' : type);
  for (const p of paras) cell.removeChild(p);
  const para = doc.createElementNS(NS.text, `${prefixOf(doc, NS.text, 'text')}:p`);
  setOdfText(doc, para, display);
  cell.appendChild(para);
  return true;
}

function writeCell(doc: Document, cell: Element, edit: CellEdit): void {
  const office = prefixOf(doc, NS.office, 'office');
  const text = prefixOf(doc, NS.text, 'text');
  const table = prefixOf(doc, NS.table, 'table');

  for (const [ns, local] of VALUE_ATTRS) cell.removeAttributeNS(ns, local);
  const oldType = cell.getAttributeNS(NS.office, 'value-type');
  cell.removeAttributeNS(NS.office, 'value-type');
  for (const n of Array.from(cell.childNodes)) if (isEl(n, NS.text, 'p')) cell.removeChild(n);

  const v = edit.value;
  if (v === '') return;
  if (v.startsWith('=') && v.length > 1) {
    // the result is filled in by the recalculation that follows
    cell.setAttributeNS(NS.table, `${table}:formula`, excelToOdf(v.slice(1)));
    cell.setAttributeNS(NS.office, `${office}:value-type`, 'float');
    cell.setAttributeNS(NS.office, `${office}:value`, '0');
    return;
  }
  const para = doc.createElementNS(NS.text, `${text}:p`);
  if (NUMBER.test(v) && !LEADING_ZERO.test(v) && Number.isFinite(Number(v))) {
    const type = oldType === 'percentage' || oldType === 'currency' ? oldType : 'float';
    cell.setAttributeNS(NS.office, `${office}:value-type`, type);
    cell.setAttributeNS(NS.office, `${office}:value`, String(Number(v)));
    para.appendChild(doc.createTextNode(v));
  } else {
    cell.setAttributeNS(NS.office, `${office}:value-type`, 'string');
    setOdfText(doc, para, v);
  }
  cell.appendChild(para);
}
