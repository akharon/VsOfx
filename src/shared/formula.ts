// A small spreadsheet formula engine (Excel-style syntax): used to recalculate .xlsx and .ods
// formulas after an edit. No runtime dependencies. Unknown functions yield UNSUPPORTED, so the
// caller can keep the result the file had cached instead of showing a wrong value.

export type Scalar = number | string | boolean | null | FormulaError;
export class FormulaError {
  constructor(readonly code: string) {}
  toString(): string {
    return this.code;
  }
}
const ERR = {
  div0: new FormulaError('#DIV/0!'),
  value: new FormulaError('#VALUE!'),
  ref: new FormulaError('#REF!'),
  name: new FormulaError('#NAME?'),
  na: new FormulaError('#N/A'),
  num: new FormulaError('#NUM!'),
  cycle: new FormulaError('#CYCLE!'),
};
/** Thrown (internally) when a formula uses something the engine does not implement. */
export const UNSUPPORTED = Symbol('unsupported');

/** What the engine needs from a workbook. Rows/columns are 0-based. */
export interface Book {
  sheetIndex(name: string): number; // -1 if missing
  /** Literal value of a non-formula cell. */
  value(sheet: number, row: number, col: number): Scalar;
  /** Formula (without "=") if the cell has one. */
  formula(sheet: number, row: number, col: number): string | undefined;
  /** Used size, to bound whole-column references. */
  size(sheet: number): { rows: number; cols: number };
}

// ---- tokenizer / parser -----------------------------------------------------------------------

type Node =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'bool'; v: boolean }
  | { t: 'err'; v: FormulaError }
  | { t: 'ref'; sheet?: string; r1: number; c1: number; r2: number; c2: number; range: boolean }
  | { t: 'un'; op: string; a: Node }
  | { t: 'bin'; op: string; a: Node; b: Node }
  | { t: 'pct'; a: Node }
  | { t: 'call'; name: string; args: Node[] };

interface Tok {
  k: 'num' | 'str' | 'ref' | 'name' | 'op' | 'err' | '(' | ')' | ',';
  v: string;
}

const REF_RE = /^(?:(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]*))!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3})(?![\w(])/;

function tokenize(src: string): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  while (i < src.length) {
    const ch = src[i]!;
    if (/\s/.test(ch)) {
      i++;
      continue;
    }
    const rest = src.slice(i);
    let m: RegExpExecArray | null;
    if ((m = /^(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?/.exec(rest))) {
      out.push({ k: 'num', v: m[0] });
      i += m[0].length;
    } else if (ch === '"') {
      let j = i + 1;
      let s = '';
      for (;;) {
        if (j >= src.length) throw ERR.value;
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            s += '"';
            j += 2;
            continue;
          }
          break;
        }
        s += src[j++];
      }
      out.push({ k: 'str', v: s });
      i = j + 1;
    } else if ((m = /^#(DIV\/0!|VALUE!|REF!|NAME\?|N\/A|NUM!|NULL!)/i.exec(rest))) {
      out.push({ k: 'err', v: m[0].toUpperCase() });
      i += m[0].length;
    } else if ((m = REF_RE.exec(rest))) {
      out.push({ k: 'ref', v: m[0] });
      i += m[0].length;
    } else if ((m = /^[A-Za-z_][\w.]*/.exec(rest))) {
      out.push({ k: 'name', v: m[0] });
      i += m[0].length;
    } else if ((m = /^(<>|<=|>=|[-+*/^&=<>%])/.exec(rest))) {
      out.push({ k: 'op', v: m[0] });
      i += m[0].length;
    } else if (ch === '(' || ch === ')' || ch === ',' || ch === ';') {
      out.push({ k: ch === ';' ? ',' : ch, v: ch });
      i++;
    } else throw UNSUPPORTED;
  }
  return out;
}

function colIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function parseRef(text: string): Node {
  const m = REF_RE.exec(text)!;
  const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
  const body = m[3]!.replace(/\$/g, '');
  const [a, b] = body.split(':') as [string, string | undefined];
  const cell = (s: string) => {
    const x = /^([A-Za-z]+)(\d*)$/.exec(s)!;
    return { c: colIndex(x[1]!), r: x[2] ? Number(x[2]) - 1 : -1 };
  };
  const p = cell(a);
  const q = b !== undefined ? cell(b) : p;
  const n: Node = { t: 'ref', r1: Math.min(p.r, q.r), c1: Math.min(p.c, q.c), r2: Math.max(p.r, q.r), c2: Math.max(p.c, q.c), range: b !== undefined };
  if (p.r < 0) {
    n.r1 = 0;
    n.r2 = -1; // whole column: bounded later
  }
  if (sheet !== undefined) n.sheet = sheet;
  return n;
}

const PREC: Record<string, number> = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };

export function parseFormula(src: string): Node {
  const toks = tokenize(src);
  let pos = 0;
  const peek = () => toks[pos];
  const next = () => toks[pos++];
  const expect = (k: Tok['k']) => {
    if (next()?.k !== k) throw ERR.value;
  };
  function primary(): Node {
    const t = next();
    if (!t) throw ERR.value;
    switch (t.k) {
      case 'num':
        return { t: 'num', v: Number(t.v) };
      case 'str':
        return { t: 'str', v: t.v };
      case 'err':
        return { t: 'err', v: new FormulaError(t.v) };
      case 'ref':
        return parseRef(t.v);
      case '(': {
        const e = expr(0);
        expect(')');
        return e;
      }
      case 'op':
        if (t.v === '-' || t.v === '+') return { t: 'un', op: t.v, a: unaryOperand() };
        throw ERR.value;
      case 'name': {
        const up = t.v.toUpperCase();
        if (peek()?.k === '(') {
          next();
          const args: Node[] = [];
          if (peek()?.k !== ')') {
            for (;;) {
              args.push(peek()?.k === ',' || peek()?.k === ')' ? { t: 'str', v: '' } : expr(0));
              if (peek()?.k === ',') next();
              else break;
            }
          }
          expect(')');
          return { t: 'call', name: up.replace(/^_XLFN\./, ''), args };
        }
        if (up === 'TRUE' || up === 'FALSE') return { t: 'bool', v: up === 'TRUE' };
        throw UNSUPPORTED; // defined names
      }
      default:
        throw ERR.value;
    }
  }
  function unaryOperand(): Node {
    // unary minus binds tighter than ^ in Excel
    let a = primary();
    while (peek()?.k === 'op' && peek()!.v === '%') {
      next();
      a = { t: 'pct', a };
    }
    return a;
  }
  function expr(minPrec: number): Node {
    let a = unaryOperand();
    for (;;) {
      const t = peek();
      if (!t || t.k !== 'op') break;
      const p = PREC[t.v];
      if (p === undefined || p < minPrec) break;
      next();
      const b = expr(p + 1); // all binary operators are left-associative in Excel (2^3^2 = 64)
      a = { t: 'bin', op: t.v, a, b };
    }
    return a;
  }
  const e = expr(0);
  if (pos !== toks.length) throw ERR.value;
  return e;
}

// ---- evaluation -------------------------------------------------------------------------------

type Grid = Scalar[][];
type Val = Scalar | Grid;

const isErr = (v: unknown): v is FormulaError => v instanceof FormulaError;
const isGrid = (v: Val): v is Grid => Array.isArray(v);

function toNum(v: Scalar): number {
  if (isErr(v)) throw v;
  if (v === null || v === '') return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  const n = Number(v.trim());
  if (v.trim() !== '' && Number.isFinite(n)) return n;
  throw ERR.value;
}
function toStr(v: Scalar): string {
  if (isErr(v)) throw v;
  if (v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return numText(v);
  return v;
}
function toBool(v: Scalar): boolean {
  if (isErr(v)) throw v;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'string') {
    const u = v.toUpperCase();
    if (u === 'TRUE') return true;
    if (u === 'FALSE') return false;
    throw ERR.value;
  }
  return toNum(v) !== 0;
}
/** Number as Excel's General format would show it in a formula context. */
export function numText(n: number): string {
  if (Number.isInteger(n)) return String(n);
  return String(Number(n.toPrecision(15)));
}

function scalar(v: Val): Scalar {
  if (!isGrid(v)) return v;
  return v[0]?.[0] ?? null; // implicit intersection simplified to the top-left cell
}
function flat(v: Val): Scalar[] {
  return isGrid(v) ? v.flat() : [v];
}

function compare(a: Scalar, b: Scalar): number {
  const rank = (x: Scalar) => (typeof x === 'number' || x === null ? 0 : typeof x === 'string' ? 1 : 2);
  const A = a === null ? (typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0) : a;
  const B = b === null ? (typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0) : b;
  if (rank(A) !== rank(B)) return rank(A) - rank(B);
  if (typeof A === 'string' && typeof B === 'string') {
    const x = A.toLowerCase();
    const y = B.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  return Number(A) - Number(B);
}

/** Criteria of SUMIF/COUNTIF: ">5", "<>x", "abc*", 3 ... */
function criterion(c: Scalar): (v: Scalar) => boolean {
  if (typeof c === 'number' || typeof c === 'boolean') return (v) => v !== null && compare(v, c) === 0;
  const s = toStr(c);
  const m = /^(<=|>=|<>|<|>|=)?(.*)$/s.exec(s)!;
  const op = m[1] ?? '=';
  const rhsText = m[2]!;
  const rhsNum = rhsText.trim() !== '' && Number.isFinite(Number(rhsText)) ? Number(rhsText) : undefined;
  if (rhsNum === undefined && (op === '=' || op === '<>')) {
    const re = new RegExp('^' + rhsText.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\0').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\0/g, '\\*') + '$', 'i');
    const test = (v: Scalar) => (rhsText === '' ? v === null || v === '' : re.test(v === null ? '' : toStrSafe(v)));
    return op === '=' ? test : (v) => !test(v);
  }
  const rhs: Scalar = rhsNum ?? rhsText;
  return (v) => {
    if (v === null || isErr(v)) return op === '<>';
    if (typeof rhs === 'number' && typeof v !== 'number') return op === '<>';
    const d = compare(v, rhs);
    switch (op) {
      case '<':
        return d < 0;
      case '>':
        return d > 0;
      case '<=':
        return d <= 0;
      case '>=':
        return d >= 0;
      case '<>':
        return d !== 0;
      default:
        return d === 0;
    }
  };
}
function toStrSafe(v: Scalar): string {
  try {
    return toStr(v);
  } catch {
    return '';
  }
}

const nums = (args: Val[]): number[] => {
  const out: number[] = [];
  for (const a of args) {
    if (isGrid(a)) {
      for (const v of a.flat()) {
        if (isErr(v)) throw v;
        if (typeof v === 'number') out.push(v);
      }
    } else out.push(toNum(a));
  }
  return out;
};
const round = (n: number, d: number, mode: 'round' | 'up' | 'down') => {
  const f = 10 ** d;
  const x = Math.abs(n) * f;
  const r = mode === 'round' ? Math.round(Number(x.toPrecision(15))) : mode === 'up' ? Math.ceil(Number(x.toPrecision(15))) : Math.floor(Number(x.toPrecision(15)));
  return (Math.sign(n) * r) / f;
};
const num1 = (fn: (n: number) => number) => (a: Val[]) => {
  const r = fn(toNum(scalar(a[0] ?? null)));
  if (!Number.isFinite(r)) throw ERR.num;
  return r;
};
const EPOCH = Date.UTC(1899, 11, 30);
const serialOf = (d: Date) => (d.getTime() - d.getTimezoneOffset() * 60000 - EPOCH) / 86400000;

/** Lookup in a 1-D vector (exact or approximate on sorted data). */
function lookup(key: Scalar, vec: Scalar[], exact: boolean): number {
  if (exact) {
    const m = typeof key === 'string' ? criterion('=' + key) : (v: Scalar) => v !== null && compare(v, key) === 0;
    return vec.findIndex((v) => m(v));
  }
  let found = -1;
  for (let i = 0; i < vec.length; i++) {
    const v = vec[i]!;
    if (v === null) continue;
    if (compare(v, key) <= 0) found = i;
    else break;
  }
  return found;
}

type Fn = (args: Val[], raw: Node[], ev: (n: Node) => Val) => Val;

const FUNCS: Record<string, Fn> = {
  SUM: (a) => nums(a).reduce((x, y) => x + y, 0),
  PRODUCT: (a) => nums(a).reduce((x, y) => x * y, 1),
  AVERAGE: (a) => {
    const n = nums(a);
    if (!n.length) throw ERR.div0;
    return n.reduce((x, y) => x + y, 0) / n.length;
  },
  MIN: (a) => {
    const n = nums(a);
    return n.length ? Math.min(...n) : 0;
  },
  MAX: (a) => {
    const n = nums(a);
    return n.length ? Math.max(...n) : 0;
  },
  MEDIAN: (a) => {
    const n = nums(a).sort((x, y) => x - y);
    if (!n.length) throw ERR.num;
    const h = n.length >> 1;
    return n.length % 2 ? n[h]! : (n[h - 1]! + n[h]!) / 2;
  },
  COUNT: (a) => a.flatMap(flat).filter((v) => typeof v === 'number').length,
  COUNTA: (a) => a.flatMap(flat).filter((v) => v !== null && v !== '').length,
  COUNTBLANK: (a) => a.flatMap(flat).filter((v) => v === null || v === '').length,
  ABS: num1(Math.abs),
  SQRT: num1((n) => (n < 0 ? NaN : Math.sqrt(n))),
  INT: num1(Math.floor),
  EXP: num1(Math.exp),
  LN: num1((n) => (n <= 0 ? NaN : Math.log(n))),
  LOG10: num1((n) => (n <= 0 ? NaN : Math.log10(n))),
  SIGN: num1(Math.sign),
  PI: () => Math.PI,
  LOG: (a) => {
    const n = toNum(scalar(a[0]!));
    const b = a[1] === undefined ? 10 : toNum(scalar(a[1]));
    if (n <= 0 || b <= 0 || b === 1) throw ERR.num;
    return Math.log(n) / Math.log(b);
  },
  POWER: (a) => {
    const r = toNum(scalar(a[0]!)) ** toNum(scalar(a[1]!));
    if (!Number.isFinite(r)) throw ERR.num;
    return r;
  },
  MOD: (a) => {
    const n = toNum(scalar(a[0]!));
    const d = toNum(scalar(a[1]!));
    if (d === 0) throw ERR.div0;
    return n - d * Math.floor(n / d);
  },
  ROUND: (a) => round(toNum(scalar(a[0]!)), toNum(scalar(a[1] ?? 0)), 'round'),
  ROUNDUP: (a) => round(toNum(scalar(a[0]!)), toNum(scalar(a[1] ?? 0)), 'up'),
  ROUNDDOWN: (a) => round(toNum(scalar(a[0]!)), toNum(scalar(a[1] ?? 0)), 'down'),
  TRUNC: (a) => round(toNum(scalar(a[0]!)), toNum(scalar(a[1] ?? 0)), 'down'),
  CEILING: (a) => {
    const n = toNum(scalar(a[0]!));
    const s = a[1] === undefined ? 1 : toNum(scalar(a[1]));
    return s === 0 ? 0 : Math.ceil(n / s) * s;
  },
  FLOOR: (a) => {
    const n = toNum(scalar(a[0]!));
    const s = a[1] === undefined ? 1 : toNum(scalar(a[1]));
    if (s === 0) throw ERR.div0;
    return Math.floor(n / s) * s;
  },
  IF: (_a, raw, ev) => {
    const c = toBool(scalar(ev(raw[0]!)));
    if (c) return raw[1] ? ev(raw[1]) : true;
    return raw[2] ? ev(raw[2]) : false;
  },
  IFERROR: (_a, raw, ev) => {
    try {
      const v = ev(raw[0]!);
      if (isErr(scalar(v))) return ev(raw[1]!);
      return v;
    } catch (e) {
      if (e instanceof FormulaError) return ev(raw[1]!);
      throw e;
    }
  },
  AND: (a) => a.flatMap(flat).filter((v) => v !== null && v !== '').every((v) => toBool(v)),
  OR: (a) => a.flatMap(flat).filter((v) => v !== null && v !== '').some((v) => toBool(v)),
  NOT: (a) => !toBool(scalar(a[0]!)),
  TRUE: () => true,
  FALSE: () => false,
  CONCATENATE: (a) => a.map((v) => toStr(scalar(v))).join(''),
  CONCAT: (a) => a.flatMap(flat).map(toStr).join(''),
  LEN: (a) => toStr(scalar(a[0]!)).length,
  UPPER: (a) => toStr(scalar(a[0]!)).toUpperCase(),
  LOWER: (a) => toStr(scalar(a[0]!)).toLowerCase(),
  TRIM: (a) => toStr(scalar(a[0]!)).trim().replace(/ +/g, ' '),
  LEFT: (a) => toStr(scalar(a[0]!)).slice(0, a[1] === undefined ? 1 : toNum(scalar(a[1]))),
  RIGHT: (a) => {
    const s = toStr(scalar(a[0]!));
    const n = a[1] === undefined ? 1 : toNum(scalar(a[1]));
    return n <= 0 ? '' : s.slice(-n);
  },
  MID: (a) => {
    const s = toStr(scalar(a[0]!));
    const st = toNum(scalar(a[1]!));
    if (st < 1) throw ERR.value;
    return s.substr(st - 1, toNum(scalar(a[2]!)));
  },
  VALUE: (a) => toNum(scalar(a[0]!)),
  ISBLANK: (a) => scalar(a[0]!) === null,
  ISNUMBER: (a) => typeof scalar(a[0]!) === 'number',
  ISTEXT: (a) => typeof scalar(a[0]!) === 'string',
  ISERROR: (_a, raw, ev) => {
    try {
      return isErr(scalar(ev(raw[0]!)));
    } catch (e) {
      if (e instanceof FormulaError) return true;
      throw e;
    }
  },
  SUMIF: (a) => {
    const range = isGrid(a[0]!) ? a[0] : [[a[0]!]];
    const test = criterion(scalar(a[1]!));
    const sumR = a[2] === undefined ? range : isGrid(a[2]) ? a[2] : [[a[2]]];
    let s = 0;
    range.forEach((row, r) => row.forEach((v, c) => {
      const x = sumR[r]?.[c];
      if (test(v) && typeof x === 'number') s += x;
    }));
    return s;
  },
  COUNTIF: (a) => flat(a[0]!).filter(criterion(scalar(a[1]!))).length,
  AVERAGEIF: (a) => {
    const range = isGrid(a[0]!) ? a[0] : [[a[0]!]];
    const test = criterion(scalar(a[1]!));
    const avgR = a[2] === undefined ? range : isGrid(a[2]) ? a[2] : [[a[2]]];
    const vals: number[] = [];
    range.forEach((row, r) => row.forEach((v, c) => {
      const x = avgR[r]?.[c];
      if (test(v) && typeof x === 'number') vals.push(x);
    }));
    if (!vals.length) throw ERR.div0;
    return vals.reduce((x, y) => x + y, 0) / vals.length;
  },
  VLOOKUP: (a) => {
    const key = scalar(a[0]!);
    const tbl = isGrid(a[1]!) ? a[1] : [[a[1]!]];
    const col = toNum(scalar(a[2]!));
    const exact = a[3] !== undefined && !toBool(scalar(a[3]));
    if (col < 1 || col > (tbl[0]?.length ?? 0)) throw ERR.ref;
    const i = lookup(key, tbl.map((r) => r[0] ?? null), exact);
    if (i < 0) throw ERR.na;
    return tbl[i]![col - 1] ?? null;
  },
  HLOOKUP: (a) => {
    const key = scalar(a[0]!);
    const tbl = isGrid(a[1]!) ? a[1] : [[a[1]!]];
    const row = toNum(scalar(a[2]!));
    const exact = a[3] !== undefined && !toBool(scalar(a[3]));
    if (row < 1 || row > tbl.length) throw ERR.ref;
    const i = lookup(key, tbl[0] ?? [], exact);
    if (i < 0) throw ERR.na;
    return tbl[row - 1]![i] ?? null;
  },
  MATCH: (a) => {
    const vec = flat(a[1]!);
    const type = a[2] === undefined ? 1 : toNum(scalar(a[2]));
    if (type === -1) throw UNSUPPORTED;
    const i = lookup(scalar(a[0]!), vec, type === 0);
    if (i < 0) throw ERR.na;
    return i + 1;
  },
  INDEX: (a) => {
    const tbl = isGrid(a[0]!) ? a[0] : [[a[0]!]];
    let r = toNum(scalar(a[1] ?? 0));
    let c = a[2] === undefined ? 0 : toNum(scalar(a[2]));
    if (tbl.length === 1 && a[2] === undefined) {
      c = r;
      r = 1;
    }
    if (r === 0 || c === 0) {
      if (r === 0 && c === 0) return tbl;
      if (r === 0) return tbl.map((row) => [row[c - 1] ?? null]);
      return [tbl[r - 1] ?? []];
    }
    const v = tbl[r - 1]?.[c - 1];
    if (v === undefined) throw ERR.ref;
    return v;
  },
  TODAY: () => Math.floor(serialOf(new Date())),
  NOW: () => serialOf(new Date()),
};

/** Recalculates formulas of a workbook; results are memoised until `reset()`. */
export class Calc {
  private memo = new Map<string, Scalar>();
  private busy = new Set<string>();
  constructor(private readonly book: Book) {}

  reset(): void {
    this.memo.clear();
  }

  /** Result of the formula in a cell, or UNSUPPORTED if the engine cannot compute it. */
  cell(sheet: number, row: number, col: number): Scalar | typeof UNSUPPORTED {
    try {
      return this.cellValue(sheet, row, col);
    } catch (e) {
      if (e === UNSUPPORTED) return UNSUPPORTED;
      throw e;
    }
  }

  /** Evaluates a formula (without "=") as if it were in the given cell. */
  evaluate(formula: string, sheet: number): Scalar | typeof UNSUPPORTED {
    try {
      return this.run(formula, sheet);
    } catch (e) {
      if (e === UNSUPPORTED) return UNSUPPORTED;
      throw e;
    }
  }

  private run(formula: string, sheet: number): Scalar {
    let node: Node;
    try {
      node = parseFormula(formula);
    } catch (e) {
      if (e instanceof FormulaError) return e;
      throw e;
    }
    try {
      const v = scalar(this.ev(node, sheet));
      if (typeof v === 'number' && !Number.isFinite(v)) return ERR.num;
      return v;
    } catch (e) {
      if (e instanceof FormulaError) return e;
      throw e;
    }
  }

  private cellValue(sheet: number, row: number, col: number): Scalar {
    const f = this.book.formula(sheet, row, col);
    if (f === undefined) return this.book.value(sheet, row, col);
    const key = `${sheet}:${row}:${col}`;
    const hit = this.memo.get(key);
    if (hit !== undefined || this.memo.has(key)) return hit!;
    if (this.busy.has(key)) return ERR.cycle;
    this.busy.add(key);
    try {
      const v = this.run(f, sheet);
      this.memo.set(key, v);
      return v;
    } finally {
      this.busy.delete(key);
    }
  }

  private ev = (n: Node, sheet: number): Val => {
    switch (n.t) {
      case 'num':
      case 'str':
      case 'bool':
        return n.v;
      case 'err':
        throw n.v;
      case 'ref': {
        const s = n.sheet === undefined ? sheet : this.book.sheetIndex(n.sheet);
        if (s < 0) throw ERR.ref;
        const r2 = n.r2 < 0 ? this.book.size(s).rows - 1 : n.r2;
        if (!n.range) return this.cellValue(s, n.r1, n.c1);
        if ((r2 - n.r1 + 1) * (n.c2 - n.c1 + 1) > 2_000_000) throw UNSUPPORTED;
        const g: Grid = [];
        for (let r = n.r1; r <= r2; r++) {
          const row: Scalar[] = [];
          for (let c = n.c1; c <= n.c2; c++) row.push(this.cellValue(s, r, c));
          g.push(row);
        }
        return g;
      }
      case 'un': {
        const v = toNum(scalar(this.ev(n.a, sheet)));
        return n.op === '-' ? -v : v;
      }
      case 'pct':
        return toNum(scalar(this.ev(n.a, sheet))) / 100;
      case 'bin': {
        const a = scalar(this.ev(n.a, sheet));
        const b = scalar(this.ev(n.b, sheet));
        if (isErr(a)) throw a;
        if (isErr(b)) throw b;
        switch (n.op) {
          case '+':
            return toNum(a) + toNum(b);
          case '-':
            return toNum(a) - toNum(b);
          case '*':
            return toNum(a) * toNum(b);
          case '/': {
            const d = toNum(b);
            if (d === 0) throw ERR.div0;
            return toNum(a) / d;
          }
          case '^':
            return toNum(a) ** toNum(b);
          case '&':
            return toStr(a) + toStr(b);
          case '=':
            return compare(a, b) === 0;
          case '<>':
            return compare(a, b) !== 0;
          case '<':
            return compare(a, b) < 0;
          case '>':
            return compare(a, b) > 0;
          case '<=':
            return compare(a, b) <= 0;
          case '>=':
            return compare(a, b) >= 0;
        }
        throw ERR.value;
      }
      case 'call': {
        const fn = FUNCS[n.name];
        if (!fn) throw UNSUPPORTED;
        const lazy = n.name === 'IF' || n.name === 'IFERROR' || n.name === 'ISERROR';
        const ev = (x: Node) => this.ev(x, sheet);
        const args = lazy ? [] : n.args.map(ev);
        // errors inside arguments propagate, except where a function handles them
        const v = fn(args, n.args, ev);
        return v;
      }
    }
  };
}

/** Converts an ODF formula ("of:=SUM([.A1:.B2])") to Excel syntax ("SUM(A1:B2)"). */
export function odfToExcel(f: string): string {
  let s = f.replace(/^(of:|oooc:|msoxl:)?=?/, '');
  s = mapOutsideStrings(s, (part) =>
    part
      .replace(/\[\$?('(?:[^']|'')+'|[^.\]:]*)?\.(\$?[A-Z]+\$?\d+)(?::\$?('(?:[^']|'')+'|[^.\]:]*)?\.(\$?[A-Z]+\$?\d+))?\]/gi, (_m, sh: string | undefined, a: string, _sh2: string | undefined, b: string | undefined) => {
        const sheet = sh ? `${/^'/.test(sh) || /^[A-Za-z_][\w]*$/.test(sh) ? sh : `'${sh}'`}!` : '';
        return sheet + a + (b ? ':' + b : '');
      })
      .replace(/;/g, ','),
  );
  return s;
}

/** Converts an Excel-syntax formula (without "=") to ODF ("of:=..."). */
export function excelToOdf(f: string): string {
  const out = mapOutsideStrings(f, (part) =>
    part
      .replace(new RegExp(`(?<![\\w.$'!])${REF_RE.source.slice(1)}`, 'g'), (m) => {
        const r = /^(?:(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]*))!)?(.*)$/.exec(m)!;
        const sheet = r[1] !== undefined ? `$'${r[1]}'` : r[2] !== undefined ? `$${r[2]}` : '';
        const [a, b] = r[3]!.split(':');
        return `[${sheet}.${a}${b ? `:.${b}` : ''}]`;
      })
      .replace(/,/g, ';'),
  );
  return 'of:=' + out;
}

function mapOutsideStrings(s: string, fn: (part: string) => string): string {
  return s.split(/("(?:[^"]|"")*")/).map((p, i) => (i % 2 ? p : fn(p))).join('');
}
