import * as ExcelNS from 'exceljs';
import type { ApplyResult, CellEdit, CellStyle, CellView, Edit, MergeRange, SheetData, SheetModel } from '../../shared/model';
import { MAX_COLS, MAX_ROWS } from '../../shared/model';
import type { FormatAdapter } from './Adapter';
import { formatDate, formatNumber } from './numFmt';

// The package is CommonJS: depending on the bundler/loader the API is on `default` or on the namespace.
const Excel: typeof ExcelNS = (ExcelNS as unknown as { default?: typeof ExcelNS }).default ?? ExcelNS;

const NUMBER = /^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/;
const LEADING_ZERO = /^[-+]?0\d/;
const ISO_DATE = /^\d{4}-\d{2}-\d{2}([T ][\d:.]+Z?)?$/;

interface Original {
  raw: ExcelNS.CellValue;
  input: string;
}

export class XlsxAdapter implements FormatAdapter {
  readonly readonly = false;
  private workbook!: ExcelNS.Workbook;
  private readonly originals = new Map<string, Original>();

  async load(bytes: Uint8Array, edits: Edit[]): Promise<SheetModel> {
    this.workbook = new Excel.Workbook();
    // exceljs accepts Buffer/ArrayBuffer; a plain copy keeps it independent of the caller's memory.
    await this.workbook.xlsx.load(bytes.slice().buffer as ArrayBuffer);
    for (const e of edits) this.applyEdit(e);
    return { type: 'sheet', sheets: this.workbook.worksheets.map((ws) => this.sheetData(ws)) };
  }

  applyEdit(edit: Edit): ApplyResult {
    if (edit.kind !== 'cell') throw new Error('Not a cell edit');
    const ws = this.workbook.worksheets[edit.sheet];
    if (!ws) throw new Error(`Sheet ${edit.sheet} does not exist`);
    const cell = ws.getCell(edit.row + 1, edit.col + 1);
    const key = `${edit.sheet}:${edit.row}:${edit.col}`;
    let orig = this.originals.get(key);
    if (!orig) {
      orig = { raw: cell.value, input: inputString(cell) };
      this.originals.set(key, orig);
    }
    // Returning to the original text restores the original typed value (dates, formulas, ...).
    cell.value = edit.value === orig.input ? orig.raw : parseInput(edit, orig.raw);
    return { cell: this.cellView(cell) };
  }

  async save(): Promise<Uint8Array> {
    const out = await this.workbook.xlsx.writeBuffer();
    return new Uint8Array(out as ArrayBuffer);
  }

  // ---- model -------------------------------------------------------------

  private sheetData(ws: ExcelNS.Worksheet): SheetData {
    const rowCount = Math.min(ws.rowCount, MAX_ROWS);
    const colCount = Math.min(ws.columnCount, MAX_COLS);
    const rows: (CellView | null)[][] = Array.from({ length: rowCount }, () => []);
    ws.eachRow({ includeEmpty: false }, (row, rn) => {
      if (rn > MAX_ROWS) return;
      const out = rows[rn - 1]!;
      row.eachCell({ includeEmpty: false }, (cell, cn) => {
        if (cn > MAX_COLS) return;
        while (out.length < cn - 1) out.push(null);
        out[cn - 1] = this.cellView(cell);
      });
    });
    const colWidths: (number | null)[] = [];
    for (let c = 1; c <= colCount; c++) colWidths.push(ws.getColumn(c).width ?? null);
    return {
      name: ws.name,
      rows,
      rowCount,
      colCount,
      merges: mergesOf(ws, rowCount, colCount),
      colWidths,
      truncated: ws.rowCount > MAX_ROWS || ws.columnCount > MAX_COLS,
    };
  }

  private cellView(cell: ExcelNS.Cell): CellView | null {
    const text = displayText(cell);
    const input = inputString(cell);
    const view: CellView = { text };
    if (input !== text) view.input = input;
    if (typeof cell.formula === 'string' && cell.formula) view.formula = cell.formula;
    const style = styleOf(cell);
    if (style) view.style = style;
    if (!text && !view.style && !view.input) return null;
    return view;
  }
}

function mergesOf(ws: ExcelNS.Worksheet, rowCount: number, colCount: number): MergeRange[] {
  const merges = (ws.model as unknown as { merges?: string[] }).merges ?? [];
  const out: MergeRange[] = [];
  for (const m of merges) {
    const r = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(m);
    if (!r) continue;
    const range = { r1: +r[2]! - 1, c1: colIndex(r[1]!), r2: +r[4]! - 1, c2: colIndex(r[3]!) };
    if (range.r1 < rowCount && range.c1 < colCount) out.push(range);
  }
  return out;
}

function colIndex(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function argbToCss(argb: string | undefined): string | undefined {
  if (!argb || !/^[0-9a-f]{6}([0-9a-f]{2})?$/i.test(argb)) return undefined;
  return '#' + (argb.length === 8 ? argb.slice(2) : argb);
}

function styleOf(cell: ExcelNS.Cell): CellStyle | undefined {
  const s: CellStyle = {};
  const font = cell.font;
  if (font?.bold) s.bold = true;
  if (font?.italic) s.italic = true;
  if (font?.underline) s.underline = true;
  const color = argbToCss(font?.color?.argb);
  if (color) s.color = color;
  const fill = cell.fill;
  if (fill && fill.type === 'pattern' && fill.pattern === 'solid') {
    const f = argbToCss(fill.fgColor?.argb);
    if (f) s.fill = f;
  }
  const h = cell.alignment?.horizontal;
  if (h === 'left' || h === 'center' || h === 'right') s.align = h;
  return Object.keys(s).length ? s : undefined;
}

/** Displayed text: the value rendered with the cell's number format (exceljs does not do this). */
function displayText(cell: ExcelNS.Cell): string {
  let v: unknown = cell.value;
  if (typeof cell.formula === 'string' && cell.formula) v = cell.result;
  const fmt = cell.numFmt;
  if (typeof v === 'number') return formatNumber(v, fmt);
  if (v instanceof Date) return formatDate(v, fmt);
  try {
    return cell.text;
  } catch {
    return String(v ?? '');
  }
}

/** The text a user would type to get this cell's value. */
function inputString(cell: ExcelNS.Cell): string {
  if (typeof cell.formula === 'string' && cell.formula) return '=' + cell.formula;
  const v = cell.value;
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return isNaN(+v) ? '' : v.toISOString().replace(/(T00:00:00)?(\.000)?Z$/, '');
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number' || typeof v === 'string') return String(v);
  try {
    return cell.text;
  } catch {
    return '';
  }
}

function parseInput(edit: CellEdit, raw: ExcelNS.CellValue): ExcelNS.CellValue {
  const v = edit.value;
  if (v === '') return null;
  if (v.startsWith('=') && v.length > 1) return { formula: v.slice(1) } as ExcelNS.CellFormulaValue;
  if (raw instanceof Date && ISO_DATE.test(v)) {
    const d = new Date(/Z$|[+-]\d\d:\d\d$/.test(v) || !v.includes('T') ? v : v + 'Z');
    if (!isNaN(+d)) return d;
  }
  if (NUMBER.test(v) && !LEADING_ZERO.test(v)) {
    const n = Number(v);
    if (Number.isFinite(n)) return n;
  }
  return v;
}
