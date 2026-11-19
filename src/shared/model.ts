// Types shared by the extension host and the webview. No runtime dependencies.
import { inverseOp, isPdfEdit, type PdfEdit, type PdfStateJson } from './pdf';

export type Format = 'docx' | 'xlsx' | 'pptx' | 'odt' | 'ods' | 'pdf';

export const FORMATS: readonly Format[] = ['docx', 'xlsx', 'pptx', 'odt', 'ods', 'pdf'];

export function formatFromPath(path: string): Format | undefined {
  const m = /\.([a-z0-9]+)$/i.exec(path);
  const ext = m?.[1]?.toLowerCase();
  return FORMATS.find((f) => f === ext);
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

export type Edit = CellEdit | ParagraphEdit | PdfEdit;

/** Runtime check for messages coming from the webview (the host does not trust them). */
export function isEdit(x: unknown): x is Edit {
  if (typeof x !== 'object' || x === null) return false;
  const e = x as Record<string, unknown>;
  const int = (v: unknown) => typeof v === 'number' && Number.isInteger(v) && v >= 0;
  if (e['kind'] === 'cell') {
    return int(e['sheet']) && int(e['row']) && int(e['col']) && typeof e['value'] === 'string' && typeof e['old'] === 'string';
  }
  if (e['kind'] === 'pdf') return isPdfEdit(x);
  if (e['kind'] === 'paragraph') {
    return int(e['index']) && typeof e['text'] === 'string' && typeof e['old'] === 'string';
  }
  return false;
}

/** The edit that reverts `e` (so the host can undo without parsing anything). */
export function inverseEdit(e: Edit): Edit {
  if (e.kind === 'pdf') return { kind: 'pdf', op: inverseOp(e.op) };
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

export type DocModel = TextModel | SheetModel | SlidesModel | PdfModel;

/** Result of applying an edit in the worker. */
export interface ApplyResult {
  /** Updated cell after a cell edit (formatted text may differ from the input). */
  cell?: CellView | null;
}

/** Display limits for spreadsheets. */
export const MAX_ROWS = 10_000;
export const MAX_COLS = 200;
