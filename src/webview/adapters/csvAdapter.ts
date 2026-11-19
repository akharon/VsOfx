import Papa from 'papaparse';
import { MAX_COLS, MAX_ROWS, type ApplyResult, type CellView, type DocModel, type Edit, type SheetData } from '../../shared/model';
import type { FormatAdapter } from './Adapter';
import { decodeText, encodeText, type TextEncodingInfo } from './textCodec';

/** Delimiters tried by the auto-detection, most common first. */
export const CSV_DELIMITERS = [',', ';', '\t', '|'];

/**
 * Detects the delimiter: the candidate that splits the first lines into the most
 * consistent, non-trivial number of fields (quotes respected by papaparse).
 */
export function detectDelimiter(text: string): string {
  const sample = text.slice(0, 64 * 1024);
  let best = ',';
  let bestScore = -1;
  for (const d of CSV_DELIMITERS) {
    const rows = Papa.parse<string[]>(sample, { delimiter: d, preview: 50, skipEmptyLines: true }).data;
    if (rows.length === 0) continue;
    const counts = rows.map((r) => r.length);
    const first = counts[0]!;
    if (first < 2) continue;
    const consistent = counts.filter((c) => c === first).length / counts.length;
    const score = consistent * 1000 + first;
    if (score > bestScore) {
      bestScore = score;
      best = d;
    }
  }
  return best;
}

/** CSV / TSV shown as a single-sheet table. Cells are plain strings (no type guessing on save). */
export class CsvAdapter implements FormatAdapter {
  readonly readonly = false;
  private original: Uint8Array = new Uint8Array(0);
  private rows: string[][] = [];
  private originalRowCount = 0;
  private dirty = false;
  private info!: TextEncodingInfo;
  delimiter = ',';

  async load(bytes: Uint8Array, edits: Edit[]): Promise<DocModel> {
    const { text, info } = decodeText(bytes);
    this.original = bytes;
    this.info = info;
    this.delimiter = detectDelimiter(text);
    const parsed = Papa.parse<string[]>(text, { delimiter: this.delimiter, skipEmptyLines: false });
    const rows = parsed.data;
    // papaparse yields one empty row for the final line break
    if (info.finalNewline && rows.length > 0 && rows[rows.length - 1]!.length === 1 && rows[rows.length - 1]![0] === '') rows.pop();
    this.rows = rows;
    this.originalRowCount = rows.length;
    this.dirty = false;
    for (const e of edits) this.applyEdit(e);

    const shown = this.rows.slice(0, MAX_ROWS);
    let colCount = 0;
    for (const r of shown) colCount = Math.max(colCount, r.length);
    const cols = Math.min(colCount, MAX_COLS);
    const widths: (number | null)[] = [];
    for (let c = 0; c < cols; c++) {
      let w = 0;
      for (let r = 0; r < Math.min(shown.length, 200); r++) w = Math.max(w, shown[r]![c]?.length ?? 0);
      widths.push(Math.min(40, Math.max(8, w + 2)));
    }
    const sheet: SheetData = {
      name: this.delimiter === '\t' ? 'TSV' : `CSV (${this.delimiter === ',' ? 'comma' : this.delimiter === ';' ? 'semicolon' : this.delimiter === '|' ? 'pipe' : this.delimiter})`,
      rows: shown.map((r, i) => r.slice(0, MAX_COLS).map((v) => (v === '' ? null : cell(v, i === 0)))),
      rowCount: shown.length,
      colCount: cols,
      merges: [],
      colWidths: widths,
      truncated: this.rows.length > MAX_ROWS || colCount > MAX_COLS,
    };
    return { type: 'sheet', sheets: [sheet] };
  }

  applyEdit(edit: Edit): ApplyResult {
    if (edit.kind !== 'cell') throw new Error('Not a cell edit');
    if (edit.sheet !== 0) throw new Error('A CSV file has only one sheet');
    while (this.rows.length <= edit.row) this.rows.push([]);
    const row = this.rows[edit.row]!;
    while (row.length <= edit.col) row.push('');
    row[edit.col] = edit.value;
    // trailing empty cells / rows that were added by the user are dropped again
    while (row.length > 1 && row[row.length - 1] === '' && edit.value === '') row.pop();
    while (this.rows.length > this.originalRowCount && this.rows[this.rows.length - 1]!.every((v) => v === '')) this.rows.pop();
    this.dirty = true;
    return { cell: edit.value === '' ? null : cell(edit.value, edit.row === 0) };
  }

  async save(): Promise<Uint8Array> {
    if (!this.dirty) return this.original;
    let text = Papa.unparse(this.rows, { delimiter: this.delimiter, newline: '\n', quotes: false });
    if (this.info.finalNewline) text += '\n';
    return encodeText(text, this.info);
  }
}

function cell(text: string, header: boolean): CellView {
  return header ? { text, style: { bold: true } } : { text };
}
