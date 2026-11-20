import type { ApplyResult, CellEdit, CellStyle, CellView, Edit, MergeRange, SheetData, SheetEdit, SheetModel } from '../../shared/model';
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
  private readonly originals = new Map<Element, Map<string, Original>>();
  /** Sheets removed by the user, kept so that undo restores them with their content. */
  private readonly removed: Element[] = [];
  /** Cell style names -> wrap / alignment they set. */
  private wraps = new Map<string, CellStyle>();
  /** Column style name -> width in characters. */
  private colStyles = new Map<string, number>();

  private view(cell: Element): CellView | null {
    const v = cellView(cell);
    const st = this.wraps.get(cell.getAttributeNS(NS.table, 'style-name') ?? '');
    if (!st) return v;
    return { ...(v ?? { text: '' }), style: { ...v?.style, ...st } };
  }

  async load(bytes: Uint8Array, edits: Edit[]): Promise<SheetModel> {
    this.pkg = openOdf(bytes, MIME_ODS);
    const sheet = descendantsNS(this.pkg.content, NS.office, 'spreadsheet')[0];
    if (!sheet) throw new Error('This OpenDocument file has no spreadsheet body');
    this.tables = Array.from(sheet.childNodes).filter((n): n is Element => isEl(n, NS.table, 'table'));
    for (const e of edits) this.applyEdit(e);
    this.wraps = cellStyles(this.pkg.content);
    this.colStyles = columnStyles(this.pkg.content);
    return { type: 'sheet', sheets: this.tables.map((t) => this.sheetData(t)), sheetOps: true };
  }

  applyEdit(edit: Edit): ApplyResult {
    if (edit.kind === 'sheet') return this.applySheet(edit);
    if (edit.kind === 'colwidth') {
      const table = this.tables[edit.sheet];
      if (!table) throw new Error(`Sheet ${edit.sheet} does not exist`);
      setColumnWidth(this.pkg.content, table, edit.col, edit.width);
      return {};
    }
    if (edit.kind === 'wrap' || edit.kind === 'align') {
      const table = this.tables[edit.sheet];
      if (!table) throw new Error(`Sheet ${edit.sheet} does not exist`);
      const cell = locateCell(this.pkg.content, locateRow(this.pkg.content, table, edit.row), edit.col);
      setCellProps(this.pkg.content, cell, edit.kind === 'wrap' ? { wrap: edit.wrap } : { h: edit.align.h ?? null, v: edit.align.v ?? null });
      this.wraps = cellStyles(this.pkg.content);
      return { cell: this.view(cell) };
    }
    if (edit.kind !== 'cell') throw new Error('Not a cell edit');
    const table = this.tables[edit.sheet];
    if (!table) throw new Error(`Sheet ${edit.sheet} does not exist`);
    const row = locateRow(this.pkg.content, table, edit.row);
    const cell = locateCell(this.pkg.content, row, edit.col);
    const key = `${edit.row}:${edit.col}`;
    let originals = this.originals.get(table);
    if (!originals) this.originals.set(table, (originals = new Map()));
    let orig = originals.get(key);
    if (!orig) {
      const view = cellView(cell);
      orig = { input: view?.input ?? view?.text ?? '', snapshot: cell.cloneNode(true) as Element };
      // the cell's own style (wrap) is not part of what the user types
      orig.snapshot.removeAttributeNS(NS.table, 'style-name');
      originals.set(key, orig);
    }
    const keepStyle = cell.getAttributeNS(NS.table, 'style-name');
    if (edit.value === orig.input) restore(cell, orig.snapshot);
    else writeCell(this.pkg.content, cell, edit);
    if (keepStyle) cell.setAttributeNS(NS.table, `${prefixOf(this.pkg.content, NS.table, 'table')}:style-name`, keepStyle);
    const others = this.recalc().filter((o) => !(o.sheet === edit.sheet && o.row === edit.row && o.col === edit.col));
    const res: ApplyResult = { cell: this.view(cell) };
    if (others.length) res.others = others;
    return res;
  }

  private applySheet(edit: SheetEdit): ApplyResult {
    const doc = this.pkg.content;
    const tp = prefixOf(doc, NS.table, 'table');
    if (edit.op === 'rename') {
      const t = this.tables[edit.index];
      if (!t) throw new Error(`Sheet ${edit.index} does not exist`);
      t.setAttributeNS(NS.table, `${tp}:name`, edit.name);
      return {};
    }
    if (edit.op === 'remove') {
      const t = this.tables[edit.index];
      if (!t) throw new Error(`Sheet ${edit.index} does not exist`);
      if (this.tables.length <= 1) throw new Error('A workbook needs at least one sheet');
      t.parentNode?.removeChild(t);
      this.tables.splice(edit.index, 1);
      this.removed.push(t);
      return {};
    }
    const at = this.removed.map((t) => t.getAttributeNS(NS.table, 'name')).lastIndexOf(edit.name);
    let t: Element;
    if (at >= 0) t = this.removed.splice(at, 1)[0]!;
    else {
      const mk = (local: string) => doc.createElementNS(NS.table, `${tp}:${local}`);
      t = mk('table');
      t.setAttributeNS(NS.table, `${tp}:name`, edit.name);
      const row = mk('table-row');
      row.appendChild(mk('table-cell'));
      t.appendChild(mk('table-column'));
      t.appendChild(row);
    }
    const before = this.tables[edit.index];
    const last = this.tables[this.tables.length - 1]!;
    if (before) before.parentNode!.insertBefore(t, before);
    else last.parentNode!.insertBefore(t, last.nextSibling);
    this.tables.splice(Math.min(edit.index, this.tables.length), 0, t);
    return { sheet: this.sheetData(t) };
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
        if (setResult(this.pkg.content, el, v) && r < MAX_ROWS && c < MAX_COLS) changed.push({ sheet: si, row: r, col: c, cell: this.view(el) });
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
    // wrapped cells
    let wr = 0;
    for (const rowEl of rowElements(table)) {
      const rep = repeatOf(rowEl, ROW_REPEAT);
      let c = 0;
      for (const cell of cellElements(rowEl)) {
        const crep = repeatOf(cell, COL_REPEAT);
        const st = this.wraps.get(cell.getAttributeNS(NS.table, 'style-name') ?? '');
        if (st) {
          for (let k = 0; k < rep && wr + k < rowCount; k++) {
            const row = rows[wr + k];
            for (let j = 0; j < crep && c + j < MAX_COLS; j++) {
              const v = row?.[c + j];
              if (row && v) row[c + j] = { ...v, style: { ...v.style, ...st } };
            }
          }
        }
        c += crep;
      }
      wr += rep;
      if (wr >= rowCount) break;
    }
    const widths: (number | null)[] = [];
    let ci = 0;
    for (const col of columnElements(table)) {
      const rep = repeatOf(col, COL_REPEAT);
      const w = this.colStyles.get(col.getAttributeNS(NS.table, 'style-name') ?? '') ?? null;
      for (let k = 0; k < rep && ci < MAX_COLS; k++) widths[ci++] = w;
      if (ci >= MAX_COLS) break;
    }
    const out = Array.from({ length: rowCount }, (_, i) => rows[i] ?? []);
    return {
      name: table.getAttributeNS(NS.table, 'name') ?? 'Sheet',
      rows: out,
      rowCount,
      colCount: Math.min(colCount, MAX_COLS),
      merges: merges.filter((m) => m.r1 < rowCount),
      colWidths: Array.from({ length: Math.min(colCount, MAX_COLS) }, (_, i) => widths[i] ?? null),
      truncated: truncated || colCount > MAX_COLS,
    };
  }
}

// ---- DOM helpers ---------------------------------------------------------

/** One character of the default font, for column widths given in characters. */
const CHAR_CM = 0.2;

function automaticStyles(doc: Document): Element {
  const root = doc.documentElement;
  let auto = Array.from(root.childNodes).find((n) => isEl(n, NS.office, 'automatic-styles')) as Element | undefined;
  if (!auto) {
    auto = doc.createElementNS(NS.office, `${prefixOf(doc, NS.office, 'office')}:automatic-styles`);
    root.insertBefore(auto, Array.from(root.childNodes).find((n) => isEl(n, NS.office, 'body')) ?? null);
  }
  return auto;
}

function styleEl(doc: Document, local: string, attrs: Record<string, string>): Element {
  const sp = prefixOf(doc, NS.style, 'style');
  if (!doc.documentElement.hasAttribute(`xmlns:${sp}`)) doc.documentElement.setAttribute(`xmlns:${sp}`, NS.style);
  const fp = prefixOf(doc, NS.fo, 'fo');
  if (!doc.documentElement.hasAttribute(`xmlns:${fp}`)) doc.documentElement.setAttribute(`xmlns:${fp}`, NS.fo);
  const e = doc.createElementNS(NS.style, `${sp}:${local}`);
  for (const [k, v] of Object.entries(attrs)) {
    const [p, l] = k.split(':') as [string, string];
    e.setAttributeNS(p === 'fo' ? NS.fo : NS.style, `${p === 'fo' ? fp : sp}:${l}`, v);
  }
  return e;
}

function findStyle(doc: Document, name: string): Element | undefined {
  return descendantsNS(doc, NS.style, 'style').find((s) => s.getAttributeNS(NS.style, 'name') === name);
}

function parseLength(v: string | null): number | null {
  const m = /^([\d.]+)(cm|mm|in|pt)$/.exec(v ?? '');
  if (!m) return null;
  const cm = Number(m[1]) * ({ cm: 1, mm: 0.1, in: 2.54, pt: 2.54 / 72 } as Record<string, number>)[m[2]!]!;
  return Math.max(1, Math.round((cm / CHAR_CM) * 10) / 10);
}

function columnStyles(doc: Document): Map<string, number> {
  const map = new Map<string, number>();
  for (const s of descendantsNS(doc, NS.style, 'style')) {
    if (s.getAttributeNS(NS.style, 'family') !== 'table-column') continue;
    const p = descendantsNS(s, NS.style, 'table-column-properties')[0];
    const w = parseLength(p?.getAttributeNS(NS.style, 'column-width') ?? null);
    const name = s.getAttributeNS(NS.style, 'name');
    if (name && w) map.set(name, w);
  }
  return map;
}

/** Cell style name -> the wrap / alignment it sets (only styles that set any). */
function cellStyles(doc: Document): Map<string, CellStyle> {
  const map = new Map<string, CellStyle>();
  for (const s of descendantsNS(doc, NS.style, 'style')) {
    if (s.getAttributeNS(NS.style, 'family') !== 'table-cell') continue;
    const name = s.getAttributeNS(NS.style, 'name');
    const p = stylePropsOf(s);
    const st: CellStyle = {};
    if (p.wrap) st.wrap = true;
    if (p.h) st.align = p.h;
    if (p.v) st.valign = p.v;
    if (name && Object.keys(st).length) map.set(name, st);
  }
  return map;
}

/** table:table-column elements in order (also inside column groups / header columns). */
function columnElements(parent: Element): Element[] {
  const out: Element[] = [];
  for (let c = parent.firstChild; c; c = c.nextSibling) {
    if (isEl(c, NS.table, 'table-column')) out.push(c as Element);
    else if (isEl(c, NS.table, 'table-column-group') || isEl(c, NS.table, 'table-header-columns') || isEl(c, NS.table, 'table-columns')) {
      out.push(...columnElements(c as Element));
    }
  }
  return out;
}

function setColumnWidth(doc: Document, table: Element, col: number, width: number | null): void {
  let pos = 0;
  let target: Element | undefined;
  const cols = columnElements(table);
  for (const c of cols) {
    const rep = repeatOf(c, COL_REPEAT);
    if (col < pos + rep) {
      target = splitRepeated(c, COL_REPEAT, col - pos);
      break;
    }
    pos += rep;
  }
  const tp = prefixOf(doc, NS.table, 'table');
  if (!target) {
    // columns are declared before the rows
    const firstRow = Array.from(table.childNodes).find((n) => isEl(n, NS.table, 'table-row') || isEl(n, NS.table, 'table-row-group') || isEl(n, NS.table, 'table-header-rows') || isEl(n, NS.table, 'table-rows')) ?? null;
    const last = cols[cols.length - 1];
    const ref = last ? last.nextSibling : firstRow;
    const parent = last ? last.parentNode! : table;
    if (col > pos) parent.insertBefore(newElement(doc, table, 'table-column', COL_REPEAT, col - pos), ref);
    target = newElement(doc, table, 'table-column');
    parent.insertBefore(target, ref);
  }
  if (width === null) {
    target.removeAttributeNS(NS.table, 'style-name');
    return;
  }
  const name = `VsofxCo${String(width).replace('.', '_')}`;
  if (!findStyle(doc, name)) {
    const s = styleEl(doc, 'style', { 'style:name': name, 'style:family': 'table-column' });
    s.appendChild(styleEl(doc, 'table-column-properties', { 'style:column-width': `${(width * CHAR_CM).toFixed(3)}cm` }));
    automaticStyles(doc).appendChild(s);
  }
  target.setAttributeNS(NS.table, `${tp}:style-name`, name);
}

/** Cell formatting the editor changes: wrap and alignment (undefined = as the cell's own style). */
interface CellProps {
  wrap?: boolean;
  h?: 'left' | 'center' | 'right';
  v?: 'top' | 'middle' | 'bottom';
}
const DERIVED = /_Vsofx(W[01])?(H[lcr])?(V[tmb])?$/;
const H_CODE = { l: 'left', c: 'center', r: 'right' } as const;
const V_CODE = { t: 'top', m: 'middle', b: 'bottom' } as const;

/** Wrap / alignment as written in a cell style. */
function stylePropsOf(s: Element | undefined): CellProps {
  const out: CellProps = {};
  if (!s) return out;
  const cp = descendantsNS(s, NS.style, 'table-cell-properties')[0];
  const pp = descendantsNS(s, NS.style, 'paragraph-properties')[0];
  if (cp?.getAttributeNS(NS.fo, 'wrap-option') === 'wrap') out.wrap = true;
  const ta = pp?.getAttributeNS(NS.fo, 'text-align');
  if (ta === 'center') out.h = 'center';
  else if (ta === 'end' || ta === 'right') out.h = 'right';
  else if (ta === 'start' || ta === 'left') out.h = 'left';
  const va = cp?.getAttributeNS(NS.style, 'vertical-align');
  if (va === 'top' || va === 'middle' || va === 'bottom') out.v = va;
  return out;
}

/**
 * Changes wrap / alignment with a derived automatic style named after the cell's own style and the
 * overrides (`<base>_Vsofx[W1][Hc][Vt]`), so the cell's other formatting is kept and undo goes back.
 */
function setCellProps(doc: Document, cell: Element, change: { wrap?: boolean; h?: CellProps['h'] | null; v?: CellProps['v'] | null }): void {
  const tp = prefixOf(doc, NS.table, 'table');
  const cur = cell.getAttributeNS(NS.table, 'style-name') ?? '';
  const m = DERIVED.exec(cur);
  const base = m ? cur.slice(0, m.index) : cur;
  const own = stylePropsOf(base ? findStyle(doc, base) : undefined);
  const over: CellProps = {};
  if (m?.[1]) over.wrap = m[1] === 'W1';
  if (m?.[2]) over.h = H_CODE[m[2][1] as keyof typeof H_CODE];
  if (m?.[3]) over.v = V_CODE[m[3][1] as keyof typeof V_CODE];
  if (change.wrap !== undefined) over.wrap = change.wrap;
  if (change.h !== undefined) over.h = change.h ?? undefined;
  if (change.v !== undefined) over.v = change.v ?? undefined;
  // an override equal to the cell's own style is no override
  if (over.wrap !== undefined && over.wrap === !!own.wrap) delete over.wrap;
  if (over.h === undefined || over.h === own.h) delete over.h;
  if (over.v === undefined || over.v === own.v) delete over.v;
  if (over.wrap === undefined && !over.h && !over.v) {
    if (base) cell.setAttributeNS(NS.table, `${tp}:style-name`, base);
    else cell.removeAttributeNS(NS.table, 'style-name');
    return;
  }
  const suffix = `${over.wrap === undefined ? '' : over.wrap ? 'W1' : 'W0'}${over.h ? `H${over.h[0]}` : ''}${over.v ? `V${over.v[0]}` : ''}`;
  const name = `${base || 'Default'}_Vsofx${suffix}`;
  if (!findStyle(doc, name)) {
    const src = base ? findStyle(doc, base) : undefined;
    let s: Element;
    const auto = automaticStyles(doc);
    if (src && src.parentNode === auto) {
      s = src.cloneNode(true) as Element;
      s.setAttributeNS(NS.style, `${prefixOf(doc, NS.style, 'style')}:name`, name);
    } else {
      s = styleEl(doc, 'style', { 'style:name': name, 'style:family': 'table-cell' });
      if (base) s.setAttributeNS(NS.style, `${prefixOf(doc, NS.style, 'style')}:parent-style-name`, base);
    }
    let props = descendantsNS(s, NS.style, 'table-cell-properties')[0];
    if (!props) {
      props = styleEl(doc, 'table-cell-properties', {});
      s.insertBefore(props, s.firstChild);
    }
    const fp = prefixOf(doc, NS.fo, 'fo');
    const sp = prefixOf(doc, NS.style, 'style');
    if (over.wrap !== undefined) props.setAttributeNS(NS.fo, `${fp}:wrap-option`, over.wrap ? 'wrap' : 'no-wrap');
    if (over.v) props.setAttributeNS(NS.style, `${sp}:vertical-align`, over.v);
    if (over.h) {
      props.setAttributeNS(NS.style, `${sp}:text-align-source`, 'fix');
      let pp = descendantsNS(s, NS.style, 'paragraph-properties')[0];
      if (!pp) {
        pp = styleEl(doc, 'paragraph-properties', {});
        s.insertBefore(pp, props.nextSibling);
      }
      pp.setAttributeNS(NS.fo, `${fp}:text-align`, over.h === 'left' ? 'start' : over.h === 'right' ? 'end' : 'center');
    }
    auto.appendChild(s);
  }
  cell.setAttributeNS(NS.table, `${tp}:style-name`, name);
}

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
