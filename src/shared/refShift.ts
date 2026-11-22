// Rewrites the cell references of an Excel-syntax formula (without "=") after rows or columns were
// inserted or deleted. Pure text; no dependencies. Whole-column / whole-row references (A:A, 1:1) are left alone.

export interface Shift {
  /** Sheet whose rows/columns moved (the formula may live in another sheet). */
  target: string;
  /** Sheet the formula is in (unqualified references point there). */
  home: string;
  axis: 'row' | 'col';
  /** 0-based first inserted / deleted line. */
  at: number;
  count: number;
  op: 'insert' | 'delete';
}

const REF = /(?<![\w.$'!])((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?(?![\w(!])/g;

const colNum = (s: string) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
function colName(i: number): string {
  let n = i + 1;
  let s = '';
  while (n > 0) {
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** New ends of a reference after the change (a range shrinks around deleted lines; undefined = all of it was deleted). */
function moveRange(a: number, b: number, s: Shift): [number, number] | undefined {
  if (s.op === 'insert') return [a >= s.at ? a + s.count : a, b >= s.at ? b + s.count : b];
  const end = s.at + s.count;
  if (a >= s.at && b < end) return undefined;
  const na = a < s.at ? a : a >= end ? a - s.count : s.at;
  const nb = b < s.at ? b : b >= end ? b - s.count : s.at - 1;
  return [na, nb];
}

export function shiftFormula(formula: string, s: Shift): string {
  return formula
    .split(/("(?:[^"]|"")*")/)
    .map((part, k) => {
      if (k % 2) return part; // string literal
      return part.replace(
        REF,
        (m, sheet: string | undefined, d1: string, c1: string, d2: string, r1: string, d3: string | undefined, c2: string | undefined, d4: string | undefined, r2: string | undefined) => {
          let name = s.home;
          if (sheet) {
            const raw = sheet.slice(0, -1);
            name = raw.startsWith("'") ? raw.slice(1, -1).replace(/''/g, "'") : raw;
          }
          if (name.toLowerCase() !== s.target.toLowerCase()) return m;
          const a = { col: colNum(c1), row: Number(r1) - 1 };
          const b = c2 !== undefined ? { col: colNum(c2), row: Number(r2) - 1 } : a;
          const moved = moveRange(a[s.axis], b[s.axis], s);
          if (!moved) return '#REF!';
          const na = { ...a, [s.axis]: moved[0] };
          const nb = { ...b, [s.axis]: moved[1] };
          const one = (p: { col: number; row: number }, dc: string, dr: string) => `${dc}${colName(p.col)}${dr}${p.row + 1}`;
          return (sheet ?? '') + one(na, d1, d2) + (c2 !== undefined ? ':' + one(nb, d3 ?? '', d4 ?? '') : '');
        },
      );
    })
    .join('');
}
