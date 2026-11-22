import { strFromU8 } from 'fflate';
import { describe, expect, it } from 'vitest';
import { Calc, excelToOdf, FormulaError, odfToExcel, UNSUPPORTED, type Book, type Scalar } from '../src/shared/formula.ts';
import type { SheetModel, SlidesModel } from '../src/shared/model.ts';
import { OdpAdapter } from '../src/webview/adapters/odpAdapter.ts';
import { OdsAdapter } from '../src/webview/adapters/odsAdapter.ts';
import { XlsxAdapter } from '../src/webview/adapters/xlsxAdapter.ts';
import { readZip } from '../src/webview/adapters/zip.ts';
import { makeOdp, makeOdsFormulas, makeXlsx } from './helpers/fixtures.ts';

/** A one-sheet book from a grid; strings starting with "=" are formulas. */
function book(grid: (Scalar | string)[][], names = ['Sheet1']): Book {
  return {
    sheetIndex: (n) => names.findIndex((x) => x.toLowerCase() === n.toLowerCase()),
    value: (_s, r, c) => {
      const v = grid[r]?.[c];
      return typeof v === 'string' && v.startsWith('=') ? null : (v ?? null);
    },
    formula: (_s, r, c) => {
      const v = grid[r]?.[c];
      return typeof v === 'string' && v.startsWith('=') ? v.slice(1) : undefined;
    },
    size: () => ({ rows: grid.length, cols: Math.max(...grid.map((r) => r.length)) }),
  };
}
const ev = (f: string, grid: (Scalar | string)[][] = [[]]) => new Calc(book(grid)).evaluate(f, 0);

describe('formula engine', () => {
  it('arithmetic, precedence, comparison and text', () => {
    expect(ev('1+2*3')).toBe(7);
    expect(ev('(1+2)*3')).toBe(9);
    expect(ev('-2^2')).toBe(4); // Excel: unary minus first
    expect(ev('2^3^2')).toBe(64);
    expect(ev('50%')).toBe(0.5);
    expect(ev('"a"&1&TRUE')).toBe('a1TRUE');
    expect(ev('1<2')).toBe(true);
    expect(ev('"abc"="ABC"')).toBe(true);
    expect((ev('1/0') as FormulaError).code).toBe('#DIV/0!');
    expect((ev('"x"+1') as FormulaError).code).toBe('#VALUE!');
  });

  it('references, ranges and functions', () => {
    const g = [
      [1, 2, '=A1+B1'],
      [10, 20, '=SUM(A1:B2)'],
      ['a', null, '=C1*2'],
    ];
    const c = new Calc(book(g));
    expect(c.cell(0, 0, 2)).toBe(3);
    expect(c.cell(0, 1, 2)).toBe(33);
    expect(c.cell(0, 2, 2)).toBe(6);
    expect(ev('AVERAGE(A1:B2)', g)).toBe(8.25);
    expect(ev('COUNT(A1:B3)', g)).toBe(4);
    expect(ev('COUNTA(A1:B3)', g)).toBe(5);
    expect(ev('IF(A1>0,"pos","neg")', g)).toBe('pos');
    expect(ev('SUMIF(A1:A2,">5",B1:B2)', g)).toBe(20);
    expect(ev('COUNTIF(A1:A3,"a")', g)).toBe(1);
    expect(ev('VLOOKUP(10,A1:B2,2,FALSE)', g)).toBe(20);
    expect(ev('INDEX(A1:B2,2,1)', g)).toBe(10);
    expect(ev('MATCH(20,B1:B2,0)', g)).toBe(2);
    expect(ev('ROUND(2.345,2)')).toBe(2.35);
    expect(ev('IFERROR(1/0,"x")')).toBe('x');
    expect(ev('SUM(A:A)', g)).toBe(11);
    expect(ev('$A$1+Sheet1!B1', g)).toBe(3);
  });

  it('reports cycles and unsupported functions', () => {
    const c = new Calc(book([['=B1', '=A1']]));
    expect((c.cell(0, 0, 0) as FormulaError).code).toBe('#CYCLE!');
    expect(ev('WEIRDFUNC(1)')).toBe(UNSUPPORTED);
  });

  it('converts between ODF and Excel syntax', () => {
    expect(odfToExcel('of:=SUM([.A1:.B2]; [$Other.C3])')).toBe('SUM(A1:B2, Other!C3)');
    expect(odfToExcel('of:=[.A1]&";"')).toBe('A1&";"');
    expect(excelToOdf('SUM(A1:B2,Other!C3)')).toBe('of:=SUM([.A1:.B2];[$Other.C3])');
    expect(excelToOdf('IF(A1>0,"a,b",1)')).toBe('of:=IF([.A1]>0;"a,b";1)');
  });
});

describe('formulas in spreadsheets', () => {
  it('xlsx: a typed formula is computed and dependents are recalculated', async () => {
    const a = new XlsxAdapter();
    await a.load(await makeXlsx('formulas'), []);
    // C1 = A1+B1 (3), C2 = SUM(A2:B2) (30)
    const r1 = a.applyEdit({ kind: 'cell', sheet: 0, row: 0, col: 0, value: '5', old: '1' });
    expect(r1.others?.find((o) => o.row === 0 && o.col === 2)?.cell?.text).toBe('7');
    const r2 = a.applyEdit({ kind: 'cell', sheet: 0, row: 2, col: 0, value: '=C1*C2', old: '' });
    expect(r2.cell?.text).toBe('210');
    expect(r2.cell?.formula).toBe('C1*C2');
    const back = await new XlsxAdapter().load(await a.save(), []);
    expect((back as SheetModel).sheets[0]!.rows[2]![0]!.text).toBe('210');
  });

  it('ods: formulas are shown in Excel syntax, typed formulas are stored as OpenFormula', async () => {
    const a = new OdsAdapter();
    const m = (await a.load(makeOdsFormulas(), [])) as SheetModel;
    const c = m.sheets[0]!.rows[0]![2]!;
    expect(c.formula).toBe('SUM(A1:B1)');
    expect(c.input).toBe('=SUM(A1:B1)');
    const r = a.applyEdit({ kind: 'cell', sheet: 0, row: 0, col: 0, value: '10', old: '2' });
    expect(r.others?.[0]?.cell?.text).toBe('13');
    const r2 = a.applyEdit({ kind: 'cell', sheet: 0, row: 1, col: 0, value: '=C1/2', old: '' });
    expect(r2.cell?.text).toBe('6.5');
    const xml = strFromU8(readZip(await a.save())['content.xml']!);
    expect(xml).toContain('table:formula="of:=[.C1]/2"');
    expect(xml).toContain('office:value="6.5"');
  });
});

describe('OpenDocument presentations and drawings', () => {
  it('odp: titles, text, notes', async () => {
    const m = (await new OdpAdapter('presentation').load(makeOdp('presentation'))) as SlidesModel;
    expect(m.slides).toHaveLength(1);
    expect(m.slides[0]!.title).toBe('Hello ODP');
    expect(m.slides[0]!.paragraphs).toEqual(['Point one', 'Shape text']);
    expect(m.slides[0]!.notes).toBe('Speaker note');
  });

  it('odg: pages', async () => {
    const m = (await new OdpAdapter('drawing').load(makeOdp('drawing'))) as SlidesModel;
    expect(m.kind).toBe('drawing');
    expect(m.slides[0]!.paragraphs).toContain('Shape text');
  });
});

