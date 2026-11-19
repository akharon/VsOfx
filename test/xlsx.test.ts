import { describe, expect, it } from 'vitest';
import type { SheetModel } from '../src/shared/model.ts';
import { XlsxAdapter } from '../src/webview/adapters/xlsxAdapter.ts';
import { makeXlsx } from './helpers/fixtures.ts';

async function load(kind: Parameters<typeof makeXlsx>[0], edits: Parameters<XlsxAdapter['load']>[1] = []) {
  const a = new XlsxAdapter();
  const model = (await a.load(await makeXlsx(kind), edits)) as SheetModel;
  return { a, model };
}

describe('xlsx adapter', () => {
  it('loads cells and sheet names', async () => {
    const { model } = await load('simple');
    expect(model.type).toBe('sheet');
    expect(model.sheets[0]!.name).toBe('Sheet1');
    expect(model.sheets[0]!.rows[0]!.map((c) => c?.text)).toEqual(['Name', 'Qty', 'Note']);
    expect(model.sheets[0]!.rows[1]![1]!.text).toBe('3');
    expect(model.sheets[0]!.rowCount).toBe(3);
    expect(model.sheets[0]!.colCount).toBe(3);
  });

  it('loads all sheets', async () => {
    const { model } = await load('multi-sheet');
    expect(model.sheets.map((s) => s.name)).toEqual(['Sheet1', 'Second', 'Third']);
  });

  it('exposes styles, merges, widths and formatted text', async () => {
    const { model } = await load('styles');
    const s = model.sheets[0]!;
    expect(s.rows[0]![0]!.style).toMatchObject({ bold: true, color: '#FFFFFF', fill: '#1F4E79' });
    expect(s.rows[1]![0]!.style?.italic).toBe(true);
    expect(s.rows[1]![1]!.text).toBe('1,234.50');
    expect(s.rows[1]![1]!.style?.align).toBe('center');
    expect(s.rows[1]![2]!.text).toBe('2024-01-15');
    expect(s.colWidths[0]).toBe(24);
    expect(s.merges).toEqual([{ r1: 3, c1: 0, r2: 3, c2: 2 }]);
  });

  it('shows formulas with cached results', async () => {
    const { model } = await load('formulas');
    const c = model.sheets[0]!.rows[0]![2]!;
    expect(c.text).toBe('3');
    expect(c.formula).toBe('A1+B1');
    expect(c.input).toBe('=A1+B1');
  });

  it('round-trips an edit: B2 changes, other styles stay', async () => {
    const { a } = await load('styles');
    const r = a.applyEdit({ kind: 'cell', sheet: 0, row: 1, col: 1, value: '99.5', old: '1234.5' });
    expect(r.cell?.text).toBe('99.50'); // number format is kept
    const saved = await a.save();

    const b = new XlsxAdapter();
    const model = (await b.load(saved, [])) as SheetModel;
    const s = model.sheets[0]!;
    expect(s.rows[1]![1]!.text).toBe('99.50');
    expect(s.rows[0]![0]!.style).toMatchObject({ bold: true, fill: '#1F4E79' });
    expect(s.rows[1]![0]!.style?.italic).toBe(true);
    expect(s.merges).toHaveLength(1);
    expect(s.colWidths[0]).toBe(24);
  });

  it('interprets typed input: numbers, formulas, text, empty', async () => {
    const { a } = await load('simple');
    const at = (row: number, col: number, value: string) => a.applyEdit({ kind: 'cell', sheet: 0, row, col, value, old: '' }).cell;
    expect(at(2, 0, '42')?.text).toBe('42');
    expect(at(2, 1, '007')?.text).toBe('007'); // leading zeros stay text
    expect(at(2, 2, '=B2*2')?.formula).toBe('B2*2');
    expect(at(1, 2, '')).toBeNull();
    const reloaded = (await new XlsxAdapter().load(await a.save(), [])) as SheetModel;
    const rows = reloaded.sheets[0]!.rows;
    expect(rows[2]![0]!.text).toBe('42');
    expect(rows[2]![2]!.formula).toBe('B2*2');
    expect(rows[1]![2] ?? null).toBeNull();
  });

  it('undo restores the original typed value, including dates and formulas', async () => {
    const { a } = await load('styles');
    const dateEdit = { kind: 'cell', sheet: 0, row: 1, col: 2, value: 'hello', old: '2024-01-15' } as const;
    expect(a.applyEdit(dateEdit).cell?.text).toBe('hello');
    expect(a.applyEdit({ ...dateEdit, value: '2024-01-15', old: 'hello' }).cell?.text).toBe('2024-01-15');
    const reloaded = (await new XlsxAdapter().load(await a.save(), [])) as SheetModel;
    expect(reloaded.sheets[0]!.rows[1]![2]!.text).toBe('2024-01-15');

    const f = await load('formulas');
    f.a.applyEdit({ kind: 'cell', sheet: 0, row: 0, col: 2, value: '9', old: '=A1+B1' });
    const back = f.a.applyEdit({ kind: 'cell', sheet: 0, row: 0, col: 2, value: '=A1+B1', old: '9' });
    expect(back.cell?.formula).toBe('A1+B1');
  });

  it('replays edits on load (restoring unsaved state)', async () => {
    const { model } = await load('simple', [{ kind: 'cell', sheet: 0, row: 1, col: 0, value: 'Bananas', old: 'Apples' }]);
    expect(model.sheets[0]!.rows[1]![0]!.text).toBe('Bananas');
  });

  it('rejects edits to missing sheets and non-cell edits', async () => {
    const { a } = await load('simple');
    expect(() => a.applyEdit({ kind: 'cell', sheet: 5, row: 0, col: 0, value: 'x', old: '' })).toThrow();
    expect(() => a.applyEdit({ kind: 'paragraph', index: 0, text: 'x', old: '' })).toThrow();
  });

  it('limits large sheets and reports it', async () => {
    const { model } = await load('large');
    const s = model.sheets[0]!;
    expect(s.rowCount).toBe(10_000);
    expect(s.truncated).toBe(true);
    expect(s.rows[9_999]![1]!.text).toBe('row 10000');
  }, 120_000);
});
