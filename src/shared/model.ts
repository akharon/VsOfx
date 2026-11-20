// Types shared by the extension host and the webview. No runtime dependencies.
import { inverseOp, isPdfEdit, type PdfEdit, type PdfStateJson } from './pdf';
import { isAnchorIndex, isNewParaIndex, isParaFmt, isTableSpec, type ParaFmt, type TableSpec } from './richtext';

export type CodeLang = 'markdown';
export type Format = 'docx' | 'xlsx' | 'pptx' | 'odt' | 'ods' | 'odp' | 'odg' | 'pdf' | 'csv' | CodeLang;

export const FORMATS: readonly Format[] = ['docx', 'xlsx', 'pptx', 'odt', 'ods', 'odp', 'odg', 'pdf', 'csv', 'markdown'];

/** File extension -> format. */
const EXTENSIONS: Record<string, Format> = {
  docx: 'docx', xlsx: 'xlsx', pptx: 'pptx', odt: 'odt', ods: 'ods', odp: 'odp', odg: 'odg', pdf: 'pdf',
  csv: 'csv', tsv: 'csv', tab: 'csv',
  md: 'markdown', markdown: 'markdown', mdown: 'markdown', mkd: 'markdown',
};

export function formatFromPath(path: string): Format | undefined {
  const m = /\.([a-z0-9]+)$/i.exec(path);
  const ext = m?.[1]?.toLowerCase();
  return ext ? EXTENSIONS[ext] : undefined;
}

export const isCodeLang = (f: Format): f is CodeLang => f === 'markdown';
export const isTextFormat = (f: Format): boolean => f === 'csv' || isCodeLang(f);

/**
 * Splice of a text document (Markdown): at offset `at` (in the
 * "\n"-normalised text), `del` is replaced by `ins`. The inverse swaps `del` and `ins`.
 */
export interface TextEdit {
  kind: 'text';
  at: number;
  del: string;
  ins: string;
}

/** Applies a splice; throws if the text does not contain `del` at `at` (document out of sync). */
export function applyTextEdit(text: string, e: TextEdit): string {
  if (e.at > text.length || text.slice(e.at, e.at + e.del.length) !== e.del) throw new Error('The edit does not match the document (out of sync).');
  return text.slice(0, e.at) + e.ins + text.slice(e.at + e.del.length);
}

/** Smallest splice that turns `a` into `b` (common prefix and suffix are kept). */
export function diffText(a: string, b: string): TextEdit | null {
  if (a === b) return null;
  let p = 0;
  const max = Math.min(a.length, b.length);
  while (p < max && a.charCodeAt(p) === b.charCodeAt(p)) p++;
  let s = 0;
  while (s < max - p && a.charCodeAt(a.length - 1 - s) === b.charCodeAt(b.length - 1 - s)) s++;
  // do not split a surrogate pair
  if (p > 0 && /[\uD800-\uDBFF]/.test(a.charAt(p - 1))) p--;
  return { kind: 'text', at: p, del: a.slice(p, a.length - s), ins: b.slice(p, b.length - s) };
}

/** Cell edit. `value`/`old` are the strings a user would type into the cell. */
export interface CellEdit {
  kind: 'cell';
  sheet: number;
  row: number; // 0-based
  col: number; // 0-based
  value: string;
  old: string;
}

/** Paragraph edit. `index` is the paragraph's ordinal in the document body (0-based). */
export interface ParagraphEdit {
  kind: 'paragraph';
  index: number;
  text: string;
  old: string;
  /** Word only: formatted content (runs, block kind) after and before the edit. */
  fmt?: ParaFmt;
  oldFmt?: ParaFmt;
}

/** Word only: a table inserted after a paragraph (or removed again: the inverse). */
export interface TableEdit {
  kind: 'table';
  op: 'insert' | 'remove';
  table: TableSpec;
}

/** Word/Writer: a paragraph added after paragraph `after` (Enter), or removed again. */
export interface NewParaEdit {
  kind: 'newpara';
  op: 'insert' | 'remove';
  index: number; // PARA_BASE + id
  after: number;
  fmt: ParaFmt;
}

/** Spreadsheets: add, remove or rename a sheet. `old` is the previous name (rename). */
export interface SheetEdit {
  kind: 'sheet';
  op: 'add' | 'remove' | 'rename';
  index: number;
  name: string;
  old?: string;
}

/** Spreadsheets: column width in characters (null = default). */
export interface ColWidthEdit {
  kind: 'colwidth';
  sheet: number;
  col: number;
  width: number | null;
  old: number | null;
}

/** Spreadsheets: wrap text in a cell on or off. */
export interface WrapEdit {
  kind: 'wrap';
  sheet: number;
  row: number;
  col: number;
  wrap: boolean;
  old: boolean;
}

/** Cell alignment; a missing side is the default (left / bottom in most programs). */
export interface CellAlign {
  h?: 'left' | 'center' | 'right';
  v?: 'top' | 'middle' | 'bottom';
}

/** Spreadsheets: horizontal and vertical alignment of a cell. */
export interface AlignEdit {
  kind: 'align';
  sheet: number;
  row: number;
  col: number;
  align: CellAlign;
  old: CellAlign;
}

const isCellAlign = (x: unknown): x is CellAlign => {
  if (typeof x !== 'object' || x === null) return false;
  const a = x as Record<string, unknown>;
  return (
    Object.keys(a).every((k) => k === 'h' || k === 'v') &&
    (a['h'] === undefined || a['h'] === 'left' || a['h'] === 'center' || a['h'] === 'right') &&
    (a['v'] === undefined || a['v'] === 'top' || a['v'] === 'middle' || a['v'] === 'bottom')
  );
};

/** Several edits that are done and undone together (paste, clearing a range, Enter). */
export interface BatchEdit {
  kind: 'batch';
  edits: Edit[];
}

export type Edit = CellEdit | ParagraphEdit | PdfEdit | TextEdit | TableEdit | NewParaEdit | SheetEdit | ColWidthEdit | WrapEdit | AlignEdit | BatchEdit;

export const MAX_SHEET_NAME = 100;
/** Sheet names that Excel and LibreOffice both accept. */
export const validSheetName = (n: string) => n.trim() !== '' && n.length <= 31 && !/[\\/?*[\]:']/.test(n);

/** Runtime check for messages coming from the webview (the host does not trust them). */
export function isEdit(x: unknown): x is Edit {
  if (typeof x !== 'object' || x === null) return false;
  const e = x as Record<string, unknown>;
  const int = (v: unknown) => typeof v === 'number' && Number.isInteger(v) && v >= 0;
  if (e['kind'] === 'cell') {
    return int(e['sheet']) && int(e['row']) && int(e['col']) && typeof e['value'] === 'string' && typeof e['old'] === 'string';
  }
  if (e['kind'] === 'pdf') return isPdfEdit(x);
  if (e['kind'] === 'text') return int(e['at']) && typeof e['del'] === 'string' && typeof e['ins'] === 'string';
  if (e['kind'] === 'paragraph') {
    return (
      int(e['index']) &&
      typeof e['text'] === 'string' &&
      typeof e['old'] === 'string' &&
      (e['fmt'] === undefined || isParaFmt(e['fmt'])) &&
      (e['oldFmt'] === undefined || isParaFmt(e['oldFmt']))
    );
  }
  if (e['kind'] === 'table') return (e['op'] === 'insert' || e['op'] === 'remove') && isTableSpec(e['table']);
  if (e['kind'] === 'newpara') {
    return (e['op'] === 'insert' || e['op'] === 'remove') && isNewParaIndex(e['index'] as number) && isAnchorIndex(e['after'] as number) && isParaFmt(e['fmt']);
  }
  if (e['kind'] === 'sheet') {
    return (
      (e['op'] === 'add' || e['op'] === 'remove' || e['op'] === 'rename') &&
      int(e['index']) &&
      (e['index'] as number) < 1000 &&
      typeof e['name'] === 'string' &&
      (e['name'] as string).length <= MAX_SHEET_NAME &&
      (e['old'] === undefined || (typeof e['old'] === 'string' && (e['old'] as string).length <= MAX_SHEET_NAME))
    );
  }
  const width = (v: unknown) => v === null || (typeof v === 'number' && Number.isFinite(v) && v > 0 && v <= 255);
  if (e['kind'] === 'colwidth') return int(e['sheet']) && int(e['col']) && width(e['width']) && width(e['old']);
  if (e['kind'] === 'wrap') return int(e['sheet']) && int(e['row']) && int(e['col']) && typeof e['wrap'] === 'boolean' && typeof e['old'] === 'boolean';
  if (e['kind'] === 'align') return int(e['sheet']) && int(e['row']) && int(e['col']) && isCellAlign(e['align']) && isCellAlign(e['old']);
  if (e['kind'] === 'batch') {
    const list = e['edits'];
    return Array.isArray(list) && list.length > 0 && list.length <= 100_000 && list.every((x) => isEdit(x) && (x as Edit).kind !== 'batch');
  }
  return false;
}

/** The edit that reverts `e` (so the host can undo without parsing anything). */
export function inverseEdit(e: Edit): Edit {
  if (e.kind === 'pdf') return { kind: 'pdf', op: inverseOp(e.op) };
  if (e.kind === 'text') return { kind: 'text', at: e.at, del: e.ins, ins: e.del };
  if (e.kind === 'table') return { ...e, op: e.op === 'insert' ? 'remove' : 'insert' };
  if (e.kind === 'cell') return { ...e, value: e.old, old: e.value };
  if (e.kind === 'colwidth') return { ...e, width: e.old, old: e.width };
  if (e.kind === 'wrap') return { ...e, wrap: e.old, old: e.wrap };
  if (e.kind === 'align') return { ...e, align: e.old, old: e.align };
  if (e.kind === 'newpara') return { ...e, op: e.op === 'insert' ? 'remove' : 'insert' };
  if (e.kind === 'batch') return { kind: 'batch', edits: e.edits.map(inverseEdit).reverse() };
  if (e.kind === 'sheet') {
    if (e.op === 'rename') return { kind: 'sheet', op: 'rename', index: e.index, name: e.old ?? e.name, old: e.name };
    return { ...e, op: e.op === 'add' ? 'remove' : 'add' };
  }
  const inv: ParagraphEdit = { kind: 'paragraph', index: e.index, text: e.old, old: e.text };
  if (e.oldFmt) inv.fmt = e.oldFmt;
  if (e.fmt) inv.oldFmt = e.fmt;
  return inv;
}

// ---- view models (produced by adapters in the worker, rendered by views) ----

export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string; // CSS color
  fill?: string; // CSS color
  align?: 'left' | 'center' | 'right';
  valign?: 'top' | 'middle' | 'bottom';
  wrap?: boolean;
}

export interface CellView {
  /** Displayed (formatted) text. */
  text: string;
  /** What the user edits, when different from `text` (formulas, dates). */
  input?: string;
  /** Formula shown as tooltip, without leading "=". */
  formula?: string;
  style?: CellStyle;
}

export interface MergeRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number; // 0-based, inclusive
}

export interface SheetData {
  name: string;
  /** Jagged: each row only has entries up to its last non-empty cell. */
  rows: (CellView | null)[][];
  rowCount: number;
  colCount: number;
  merges: MergeRange[];
  /** Column widths in characters, null = default. */
  colWidths: (number | null)[];
  /** True if rows/columns beyond the display limit were cut off. */
  truncated: boolean;
}

export interface SheetModel {
  type: 'sheet';
  sheets: SheetData[];
  /** Sheets can be added, removed and renamed (not CSV). */
  sheetOps?: boolean;
}

export interface TextModel {
  type: 'text';
  /** Sanitized HTML; editable blocks carry `data-p="<paragraph index>"`. */
  html: string;
  /** Word: formatting (bold, lists, links, tables ...) can be edited. */
  rich?: boolean;
  /** Word: tables inserted by (replayed) edits, with the HTML of each cell (row-major). */
  tables?: { table: TableSpec; cells: string[] }[];
  /** Paragraphs added by (replayed) edits, in the order they were added. */
  paras?: { index: number; after: number; fmt: ParaFmt }[];
}

export interface SlideView {
  number: number;
  title: string;
  paragraphs: string[];
  notes: string;
  imageDataUri?: string;
}

export interface SlidesModel {
  type: 'slides';
  /** 'drawing' for .odg (pages instead of slides). */
  kind?: 'presentation' | 'drawing';
  slides: SlideView[];
}

export interface PdfModel {
  type: 'pdf';
  /** State after replaying the unsaved edits (see shared/pdf.ts). */
  state: PdfStateJson;
}

/** JSON, YAML, XML or Markdown source ("\n" line ends; the adapter restores the original ones on save). */
export interface CodeModel {
  type: 'code';
  lang: CodeLang;
  text: string;
}

/** One file of the document as text, for the Source view. */
export interface SourcePart {
  name: string;
  text: string;
}

export type DocModel = TextModel | SheetModel | SlidesModel | PdfModel | CodeModel;

/** Result of applying an edit in the worker. */
export interface ApplyResult {
  /** Updated cell after a cell edit (formatted text may differ from the input). */
  cell?: CellView | null;
  /** Other cells whose formula results changed because of the edit. */
  others?: { sheet: number; row: number; col: number; cell: CellView | null }[];
  /** Results of the parts of a batch edit, in order. */
  batch?: ApplyResult[];
  /** Sheet edits: the sheet as it is after the edit (add / restore). */
  sheet?: SheetData;
}

/** Display limits for spreadsheets. */
export const MAX_ROWS = 10_000;
export const MAX_COLS = 200;
