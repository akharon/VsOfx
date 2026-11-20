import * as ExcelNS from 'exceljs';
import type { ApplyResult, CellEdit, CellStyle, CellView, Edit, MergeRange, SheetData, SheetEdit, SheetModel } from '../../shared/model';
import { MAX_COLS, MAX_ROWS } from '../../shared/model';
import { Calc, FormulaError, UNSUPPORTED, type Book, type Scalar } from '../../shared/formula';
import type { FormatAdapter } from './Adapter';
import { formatDate, formatNumber } from './numFmt';
import { readZip, writeZip } from './zip';

const CP_NS = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
const CP_ELEMENTS = ['coreProperties', 'category', 'contentStatus', 'contentType', 'keywords', 'lastModifiedBy', 'lastPrinted', 'revision', 'version'];

/**
 * exceljs only understands docProps/core.xml with the usual `cp:` prefix. Some generators
 * (e.g. Microsoft 365 Copilot) declare that namespace as the default namespace instead
 * (`<coreProperties xmlns="..."><lastModifiedBy>`), which is valid XML but makes exceljs throw
 * "Unexpected xml node in parseOpen". Such a part is rewritten to the prefixed form in memory
 * before loading; every other file is passed through unchanged.
 */
export function normalizeForExcelJs(bytes: Uint8Array): Uint8Array {
  let entries;
  try {
    entries = readZip(bytes);
  } catch {
    return bytes; // let exceljs report the problem
  }
  const name = Object.keys(entries).find((n) => n.toLowerCase() === 'docprops/core.xml');
  const part = name ? entries[name] : undefined;
  if (!name || !part) return bytes;
  let xml = new TextDecoder().decode(part).replace(/^﻿/, '');
  if (/<cp:coreProperties[\s>]/.test(xml) || !xml.includes(`xmlns="${CP_NS}"`)) return bytes;
  xml = xml.replace(`xmlns="${CP_NS}"`, `xmlns:cp="${CP_NS}"`);
  xml = xml.replace(new RegExp(`<(/?)(${CP_ELEMENTS.join('|')})(?=[\\s/>])`, 'g'), '<$1cp:$2');
  return writeZip(entries, { [name]: xml });
}

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
    await this.workbook.xlsx.load(normalizeForExcelJs(bytes).slice().buffer as ArrayBuffer);
    for (const e of edits) this.applyEdit(e);
    return { type: 'sheet', sheets: this.workbook.worksheets.map((ws) => this.sheetData(ws)), sheetOps: true };
  }

  applyEdit(edit: Edit): ApplyResult {
    if (edit.kind === 'sheet') return this.applySheet(edit);
    if (edit.kind === 'colwidth') {
      const ws = this.workbook.worksheets[edit.sheet];
      if (!ws) throw new Error(`Sheet ${edit.sheet} does not exist`);
      ws.getColumn(edit.col + 1).width = edit.width ?? (undefined as unknown as number);
      return {};
    }
    if (edit.kind === 'wrap') {
      const ws = this.workbook.worksheets[edit.sheet];
      if (!ws) throw new Error(`Sheet ${edit.sheet} does not exist`);
      const cell = ws.getCell(edit.row + 1, edit.col + 1);
      const al: Partial<ExcelNS.Alignment> = { ...(cell.alignment ?? {}) };
      if (edit.wrap) al.wrapText = true;
      else delete al.wrapText;
      cell.alignment = al;
      return { cell: this.cellView(cell) };
    }
    if (edit.kind === 'align') {
      const ws = this.workbook.worksheets[edit.sheet];
      if (!ws) throw new Error(`Sheet ${edit.sheet} does not exist`);
      const cell = ws.getCell(edit.row + 1, edit.col + 1);
      const al: Partial<ExcelNS.Alignment> = { ...(cell.alignment ?? {}) };
      if (edit.align.h) al.horizontal = edit.align.h;
      else delete al.horizontal;
      if (edit.align.v) al.vertical = edit.align.v;
      else delete al.vertical;
      cell.alignment = al;
      return { cell: this.cellView(cell) };
    }
    if (edit.kind !== 'cell') throw new Error('Not a cell edit');
    const ws = this.workbook.worksheets[edit.sheet];
    if (!ws) throw new Error(`Sheet ${edit.sheet} does not exist`);
    const cell = ws.getCell(edit.row + 1, edit.col + 1);
    const key = `${ws.id}:${edit.row}:${edit.col}`;
    let orig = this.originals.get(key);
    if (!orig) {
      orig = { raw: cell.value, input: inputString(cell) };
      this.originals.set(key, orig);
    }
    // Returning to the original text restores the original typed value (dates, formulas, ...).
    cell.value = edit.value === orig.input ? orig.raw : parseInput(edit, orig.raw);
    // Excel recalculates everything when the file is opened again
    (this.workbook as unknown as { calcProperties: { fullCalcOnLoad?: boolean } }).calcProperties.fullCalcOnLoad = true;
    const others = this.recalc().filter((o) => !(o.sheet === edit.sheet && o.row === edit.row && o.col === edit.col));
    const res: ApplyResult = { cell: this.cellView(cell) };
    if (others.length) res.others = others;
    return res;
  }

  /** Sheets removed by the user, kept so that undo restores them with their content. */
  private readonly removed: ExcelNS.Worksheet[] = [];

  private applySheet(edit: SheetEdit): ApplyResult {
    const wb = this.workbook as unknown as { _worksheets: (ExcelNS.Worksheet | undefined)[]; removeWorksheetEx(ws: ExcelNS.Worksheet): void };
    const list = this.workbook.worksheets;
    const reorder = (order: ExcelNS.Worksheet[]) => order.forEach((w, i) => ((w as unknown as { orderNo: number }).orderNo = i + 1));
    if (edit.op === 'rename') {
      const ws = list[edit.index];
      if (!ws) throw new Error(`Sheet ${edit.index} does not exist`);
      ws.name = edit.name;
      return {};
    }
    if (edit.op === 'remove') {
      const ws = list[edit.index];
      if (!ws) throw new Error(`Sheet ${edit.index} does not exist`);
      if (list.length <= 1) throw new Error('A workbook needs at least one sheet');
      this.removed.push(ws);
      wb.removeWorksheetEx(ws);
      reorder(list.filter((w) => w !== ws));
      return {};
    }
    const at = this.removed.map((w) => w.name).lastIndexOf(edit.name);
    let ws: ExcelNS.Worksheet;
    if (at >= 0) {
      ws = this.removed.splice(at, 1)[0]!;
      const w = ws as unknown as { id: number };
      if (wb._worksheets[w.id]) w.id = wb._worksheets.length; // the id was given to a newer sheet
      wb._worksheets[w.id] = ws;
    } else ws = this.workbook.addWorksheet(edit.name);
    const order = this.workbook.worksheets.filter((w) => w !== ws);
    order.splice(Math.min(edit.index, order.length), 0, ws);
    reorder(order);
    return { sheet: this.sheetData(ws) };
  }

  /** Recomputes all formulas; returns the cells whose result changed. Unsupported formulas keep their cached result. */
  private recalc(): NonNullable<ApplyResult['others']> {
    const wss = this.workbook.worksheets;
    const find = (s: number, r: number, c: number) => wss[s]?.findRow(r + 1)?.findCell(c + 1);
    const book: Book = {
      sheetIndex: (name) => wss.findIndex((w) => w.name.toLowerCase() === name.toLowerCase()),
      value: (s, r, c) => scalarOf(find(s, r, c)?.value),
      formula: (s, r, c) => {
        const f = find(s, r, c)?.formula;
        return typeof f === 'string' && f ? f : undefined;
      },
      size: (s) => ({ rows: wss[s]?.rowCount ?? 0, cols: wss[s]?.columnCount ?? 0 }),
    };
    const calc = new Calc(book);
    const changed: NonNullable<ApplyResult['others']> = [];
    wss.forEach((ws, si) => {
      ws.eachRow({ includeEmpty: false }, (row, rn) => {
        row.eachCell({ includeEmpty: false }, (cell, cn) => {
          if (typeof cell.formula !== 'string' || !cell.formula) return;
          const v = calc.cell(si, rn - 1, cn - 1);
          if (v === UNSUPPORTED) return;
          const result = resultOf(v);
          const cur = cell.value as unknown as { result?: unknown };
          if (sameResult(cur.result, result)) return;
          cell.value = { ...(cell.value as object), result } as unknown as ExcelNS.CellValue;
          if (rn <= MAX_ROWS && cn <= MAX_COLS) changed.push({ sheet: si, row: rn - 1, col: cn - 1, cell: this.cellView(cell) });
        });
      });
    });
    return changed;
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
  const v = cell.alignment?.vertical;
  if (v === 'top' || v === 'bottom') s.valign = v;
  else if (v === 'middle') s.valign = 'middle';
  if (cell.alignment?.wrapText) s.wrap = true;
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

const EPOCH = Date.UTC(1899, 11, 30);

/** exceljs cell value -> formula engine value. */
function scalarOf(v: ExcelNS.CellValue | undefined): Scalar {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number' || typeof v === 'string' || typeof v === 'boolean') return v;
  if (v instanceof Date) return (v.getTime() - EPOCH) / 86400000;
  const o = v as unknown as Record<string, unknown>;
  if ('result' in o) return scalarOf(o['result'] as ExcelNS.CellValue);
  if (Array.isArray(o['richText'])) return (o['richText'] as { text: string }[]).map((t) => t.text).join('');
  if (typeof o['text'] === 'string') return o['text'];
  if (typeof o['error'] === 'string') return new FormulaError(o['error']);
  return null;
}

/** Engine value -> exceljs formula result. */
function resultOf(v: Scalar): unknown {
  if (v === null) return 0;
  if (v instanceof FormulaError) return { error: v.code };
  return v;
}

function sameResult(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (a && b && typeof a === 'object' && typeof b === 'object') return (a as { error?: string }).error === (b as { error?: string }).error;
  return false;
}
