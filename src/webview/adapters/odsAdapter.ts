import type { ApplyResult, CellEdit, CellView, Edit, MergeRange, SheetData, SheetModel } from '../../shared/model';
import { MAX_COLS, MAX_ROWS } from '../../shared/model';
import type { FormatAdapter } from './Adapter';
import { isEl, MIME_ODS, NS, odfText, openOdf, saveOdf, setOdfText, type OdfPackage } from './odf';
import { descendantsNS } from './xml';

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
    return { cell: cellView(cell) };
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
  if (formula) view.formula = formula.replace(/^(of:)?=/, '');
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

function writeCell(doc: Document, cell: Element, edit: CellEdit): void {
  const prefixOf = (ns: string, fallback: string) => {
    // Use the prefix the document already uses for this namespace.
    for (const a of Array.from(cell.ownerDocument.documentElement.attributes)) {
      if (a.value === ns && a.name.startsWith('xmlns:')) return a.name.slice(6);
    }
    return fallback;
  };
  const office = prefixOf(NS.office, 'office');
  const text = prefixOf(NS.text, 'text');

  for (const [ns, local] of VALUE_ATTRS) cell.removeAttributeNS(ns, local);
  const oldType = cell.getAttributeNS(NS.office, 'value-type');
  cell.removeAttributeNS(NS.office, 'value-type');
  for (const n of Array.from(cell.childNodes)) if (isEl(n, NS.text, 'p')) cell.removeChild(n);

  const v = edit.value;
  if (v === '') return;
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
