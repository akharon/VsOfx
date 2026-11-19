// Formatted paragraph content for Word documents, shared by host (validation), worker (save)
// and view. A paragraph is a list of runs plus an optional block kind; `block` is only set when
// the user changed it, so untouched paragraph properties (styles, numbering) are kept on save.

export type BlockKind = 'p' | 'h1' | 'h2' | 'h3' | 'h4' | 'h5' | 'h6' | 'ul' | 'ol' | 'quote';
export const BLOCK_KINDS: readonly BlockKind[] = ['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'quote'];

export interface Run {
  t: string;
  b?: true;
  i?: true;
  u?: true;
  s?: true;
  code?: true;
  sup?: true;
  sub?: true;
  /** External link (http, https or mailto). */
  href?: string;
}

export interface ParaFmt {
  runs: Run[];
  block?: BlockKind;
}

/** Task list items are plain paragraphs that start with one of these marks. */
export const TASK_OPEN = '☐';
export const TASK_DONE = '☑';

export const SAFE_HREF = /^(https?:\/\/|mailto:)[^\s"<>]*$/i;

const FLAGS = ['b', 'i', 'u', 's', 'code', 'sup', 'sub'] as const;

export function isRun(x: unknown): x is Run {
  if (typeof x !== 'object' || x === null) return false;
  const r = x as Record<string, unknown>;
  if (typeof r['t'] !== 'string' || r['t'].length > 100_000) return false;
  for (const k of Object.keys(r)) {
    if (k === 't') continue;
    if (k === 'href') {
      if (typeof r['href'] !== 'string' || !SAFE_HREF.test(r['href'])) return false;
    } else if (!(FLAGS as readonly string[]).includes(k) || r[k] !== true) return false;
  }
  return true;
}

export function isParaFmt(x: unknown): x is ParaFmt {
  if (typeof x !== 'object' || x === null) return false;
  const f = x as Record<string, unknown>;
  return (
    Array.isArray(f['runs']) &&
    f['runs'].length <= 10_000 &&
    f['runs'].every(isRun) &&
    (f['block'] === undefined || (BLOCK_KINDS as readonly unknown[]).includes(f['block']))
  );
}

export const runsText = (runs: Run[]): string => runs.map((r) => r.t).join('');

const sameFormat = (a: Run, b: Run) => FLAGS.every((k) => a[k] === b[k]) && a.href === b.href;

/** Merges neighbouring runs with the same formatting and drops empty ones. */
export function normalizeRuns(runs: Run[]): Run[] {
  const out: Run[] = [];
  for (const r of runs) {
    if (r.t === '') continue;
    const last = out[out.length - 1];
    if (last && sameFormat(last, r)) last.t += r.t;
    else out.push({ ...r });
  }
  return out;
}

export function fmtEquals(a: ParaFmt | undefined, b: ParaFmt | undefined): boolean {
  return JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
}

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** HTML for runs (everything escaped; links are inert `data-href`). */
export function runsToHtml(runs: Run[]): string {
  return runs
    .map((r) => {
      let h = esc(r.t).replace(/\n/g, '<br>');
      if (r.code) h = `<code>${h}</code>`;
      if (r.sub) h = `<sub>${h}</sub>`;
      if (r.sup) h = `<sup>${h}</sup>`;
      if (r.s) h = `<s>${h}</s>`;
      if (r.u) h = `<u>${h}</u>`;
      if (r.i) h = `<em>${h}</em>`;
      if (r.b) h = `<strong>${h}</strong>`;
      if (r.href) h = `<a data-href="${esc(r.href)}">${h}</a>`;
      return h;
    })
    .join('');
}

/** A table inserted by the user (its cells are paragraphs with virtual indices, see cellIndex). */
export interface TableSpec {
  id: number; // 1, 2, ...
  after: number; // paragraph index the table follows
  rows: number;
  cols: number;
}

/** Paragraph indices of inserted table cells start here (real paragraphs are far below). */
export const CELL_BASE = 1_000_000;
export const MAX_TABLE_CELLS = 999;

export const cellIndex = (t: Pick<TableSpec, 'id' | 'cols'>, r: number, c: number) => CELL_BASE + t.id * 1000 + r * t.cols + c;

export function isTableSpec(x: unknown): x is TableSpec {
  if (typeof x !== 'object' || x === null) return false;
  const t = x as Record<string, unknown>;
  const int = (v: unknown, min: number, max: number) => typeof v === 'number' && Number.isInteger(v) && v >= min && v <= max;
  return int(t['id'], 1, 999) && int(t['after'], 0, CELL_BASE - 1) && int(t['rows'], 1, 99) && int(t['cols'], 1, 20) && (t['rows'] as number) * (t['cols'] as number) <= MAX_TABLE_CELLS;
}
