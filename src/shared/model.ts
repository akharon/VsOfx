// Types shared by the extension host and the webview. No runtime dependencies.
import { inverseOp, isPdfEdit, type PdfEdit, type PdfStateJson } from './pdf';

export type CodeLang = 'json' | 'yaml' | 'xml' | 'markdown';
export type Format = 'docx' | 'xlsx' | 'pptx' | 'odt' | 'ods' | 'pdf' | 'csv' | CodeLang;

export const FORMATS: readonly Format[] = ['docx', 'xlsx', 'pptx', 'odt', 'ods', 'pdf', 'csv', 'json', 'yaml', 'xml', 'markdown'];

/** File extension -> format. */
const EXTENSIONS: Record<string, Format> = {
  docx: 'docx', xlsx: 'xlsx', pptx: 'pptx', odt: 'odt', ods: 'ods', pdf: 'pdf',
  csv: 'csv', tsv: 'csv', tab: 'csv',
  json: 'json', jsonc: 'json', json5: 'json', geojson: 'json', webmanifest: 'json',
  yaml: 'yaml', yml: 'yaml',
  xml: 'xml', xsd: 'xml', xsl: 'xml', xslt: 'xml', svg: 'xml', plist: 'xml', csproj: 'xml', resx: 'xml', xaml: 'xml', config: 'xml',
  md: 'markdown', markdown: 'markdown', mdown: 'markdown', mkd: 'markdown',
};

export function formatFromPath(path: string): Format | undefined {
  const m = /\.([a-z0-9]+)$/i.exec(path);
  const ext = m?.[1]?.toLowerCase();
  return ext ? EXTENSIONS[ext] : undefined;
}

export const isCodeLang = (f: Format): f is CodeLang => f === 'json' || f === 'yaml' || f === 'xml' || f === 'markdown';
export const isTextFormat = (f: Format): boolean => f === 'csv' || isCodeLang(f);

/**
 * Splice of a text document (JSON, YAML, XML, Markdown): at offset `at` (in the
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
}

export type Edit = CellEdit | ParagraphEdit | PdfEdit | TextEdit;

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
    return int(e['index']) && typeof e['text'] === 'string' && typeof e['old'] === 'string';
  }
  return false;
}

/** The edit that reverts `e` (so the host can undo without parsing anything). */
export function inverseEdit(e: Edit): Edit {
  if (e.kind === 'pdf') return { kind: 'pdf', op: inverseOp(e.op) };
  if (e.kind === 'text') return { kind: 'text', at: e.at, del: e.ins, ins: e.del };
  return e.kind === 'cell' ? { ...e, value: e.old, old: e.value } : { ...e, text: e.old, old: e.text };
}

// ---- view models (produced by adapters in the worker, rendered by views) ----

export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string; // CSS color
  fill?: string; // CSS color
  align?: 'left' | 'center' | 'right';
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
}

export interface TextModel {
  type: 'text';
  /** Sanitized HTML; editable blocks carry `data-p="<paragraph index>"`. */
  html: string;
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

export type DocModel = TextModel | SheetModel | SlidesModel | PdfModel | CodeModel;

/** Result of applying an edit in the worker. */
export interface ApplyResult {
  /** Updated cell after a cell edit (formatted text may differ from the input). */
  cell?: CellView | null;
}

/** Display limits for spreadsheets. */
export const MAX_ROWS = 10_000;
export const MAX_COLS = 200;
