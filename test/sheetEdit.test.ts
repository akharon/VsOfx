// @vitest-environment jsdom
// Sort, insert / delete rows and columns, AutoSum: the view and the adapters together, with save / reload.
import ExcelJS from 'exceljs';
import { beforeEach, describe, expect, it } from 'vitest';
import { inverseEdit, isEdit, type ApplyResult, type Edit, type Format, type SheetModel } from '../src/shared/model.ts';
import { shiftFormula, type Shift } from '../src/shared/refShift.ts';
import { createAdapter } from '../src/webview/adapters/index.ts';
import { createToolbar, type Toolbar } from '../src/webview/ui/toolbar.ts';
import { renderSheetView } from '../src/webview/views/sheetView.ts';
import { makeOdsData } from './helpers/fixtures.ts';

let root: HTMLElement;
beforeEach(() => {
  document.body.innerHTML = '<div id="root"></div>';
  root = document.getElementById('root')!;
});

const enc = (s: string) => new TextEncoder().encode(s);

async function xlsxData(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.addRow(['Item', 'Qty']);
  ws.addRow(['b', 3]);
  ws.addRow(['a', 10]);
  ws.addRow(['c', 2]);
  ws.addRow([]);
  ws.addRow(['', { formula: 'SUM(B2:B4)', result: 15 }]);
  ws.getCell('A3').font = { bold: true };
  ws.getCell('B3').numFmt = '0.00';
  const other = wb.addWorksheet('Other');
  other.addRow([{ formula: 'S!B3*2', result: 20 }]);
  return new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
}

const FORMATS: { format: Format; bytes: () => Promise<Uint8Array> | Uint8Array; formulas: boolean }[] = [
  { format: 'xlsx', bytes: xlsxData, formulas: true },
  { format: 'ods', bytes: makeOdsData, formulas: true },
  { format: 'csv', bytes: () => enc('Item,Qty\nb,3\na,10\nc,2\n'), formulas: false },
];

/** What a user would type into every cell of the first sheet. */
const grid = (m: SheetModel, sheet = 0) => {
  const s = m.sheets[sheet]!;
  return Array.from({ length: s.rowCount }, (_, r) => Array.from({ length: s.colCount }, (_, c) => s.rows[r]?.[c]?.input ?? s.rows[r]?.[c]?.text ?? ''));
};

/** A view wired to an adapter the way main.ts does it, with an undo / redo stack. */
async function session(format: Format, bytes: Uint8Array) {
  const adapter = createAdapter(format);
  const model = (await adapter.load(bytes, [])) as SheetModel;
  const tb: Toolbar = createToolbar();
  document.body.append(tb.el); // outside the view's container, as in the real page
  const done: Edit[] = [];
  const undone: Edit[] = [];
  const show = (e: Edit, res: ApplyResult): void => {
    if (e.kind === 'batch') e.edits.forEach((x, i) => show(x, res.batch?.[i] ?? {}));
    else view.applyEdit(e, res);
  };
  const run = (e: Edit) => show(e, adapter.applyEdit(e));
  const view = renderSheetView(
    root,
    model,
    false,
    (e) => {
      expect(isEdit(e)).toBe(true);
      done.push(e);
      undone.length = 0;
      run(e);
    },
    tb,
  );
  return {
    adapter,
    model,
    view,
    done,
    undo() {
      const e = done.pop()!;
      undone.push(e);
      run(inverseEdit(e));
    },
    redo() {
      const e = undone.pop()!;
      done.push(e);
      run(e);
    },
    async reload() {
      return (await createAdapter(format).load(await adapter.save(), [])) as SheetModel;
    },
    click: (title: string) => document.body.querySelector<HTMLButtonElement>(`button[title^="${title}"]`)!.click(),
  };
}

const td = (r: number, c: number) => root.querySelector<HTMLElement>(`td[data-r="${r}"][data-c="${c}"]`)!;
const mouse = (el: Element, type: string, init: MouseEventInit = {}) => el.dispatchEvent(new MouseEvent(type, { bubbles: true, button: 0, ...init }));
const selectCell = (r: number, c: number, r2 = r, c2 = c) => {
  mouse(td(r, c), 'mousedown');
  if (r !== r2 || c !== c2) mouse(td(r2, c2), 'mousedown', { shiftKey: true });
};
const selectRow = (r: number) => mouse(root.querySelector(`th.rowhead[data-r="${r}"]`)!, 'mousedown');
const selectCol = (c: number) => mouse(root.querySelector(`th.colhead[data-c="${c}"]`)!, 'mousedown');

describe('shiftFormula', () => {
  const sh = (over: Partial<Shift>): Shift => ({ target: 'S', home: 'S', axis: 'row', at: 2, count: 1, op: 'insert', ...over });
  it('inserts: moves references at or below the line, grows ranges that contain it', () => {
    expect(shiftFormula('SUM(B2:B4)+A1+$A$3', sh({}))).toBe('SUM(B2:B5)+A1+$A$4');
    expect(shiftFormula('B3', sh({ at: 2, count: 2 }))).toBe('B5');
    expect(shiftFormula('C1+AA2', sh({ axis: 'col', at: 2 }))).toBe('D1+AB2');
  });
  it('deletes: moves, shrinks ranges, #REF! when the target is gone', () => {
    expect(shiftFormula('SUM(B2:B4)', sh({ op: 'delete' }))).toBe('SUM(B2:B3)');
    expect(shiftFormula('B3+B4', sh({ op: 'delete' }))).toBe('#REF!+B3');
    expect(shiftFormula('SUM(B3:B3)', sh({ op: 'delete' }))).toBe('SUM(#REF!)');
    expect(shiftFormula('SUM(B3:B6)', sh({ op: 'delete', count: 2, at: 1 }))).toBe('SUM(B2:B4)');
  });
  it('respects sheet names, strings and function names', () => {
    expect(shiftFormula('Other!B3', sh({}))).toBe('Other!B3');
    expect(shiftFormula("S!B3+'S'!B3", sh({ home: 'Other' }))).toBe("S!B4+'S'!B4");
    expect(shiftFormula('"B3"&LOG10(B3)', sh({}))).toBe('"B3"&LOG10(B4)');
    expect(shiftFormula("'My ''Sheet'!A5", sh({ target: "My 'Sheet", home: 'X' }))).toBe("'My ''Sheet'!A6");
  });
});

describe('edit validation', () => {
  it('accepts a lines edit and rejects nonsense', () => {
    const ok: Edit = { kind: 'lines', sheet: 0, axis: 'row', op: 'delete', at: 1, count: 2, cells: [['a', 'b'], []] };
    expect(isEdit(ok)).toBe(true);
    expect(inverseEdit(ok)).toEqual({ ...ok, op: 'insert' });
    expect(isEdit({ ...ok, count: 0 })).toBe(false);
    expect(isEdit({ ...ok, axis: 'diag' })).toBe(false);
    expect(isEdit({ ...ok, cells: [[1]] })).toBe(false);
  });
});

describe.each(FORMATS)('$format: sort', ({ format, bytes }) => {
  it('sorts the data around the active cell by the active column (numbers numerically), keeps the header, one undo step', async () => {
    const s = await session(format, await bytes());
    const before = grid(s.model);
    selectCell(2, 1);
    s.click('Sort A');
    expect(grid(s.model).slice(0, 4)).toEqual([['Item', 'Qty'], ['c', '2'], ['b', '3'], ['a', '10']]);
    expect(s.done).toHaveLength(1);
    const saved = grid(await s.reload());
    expect(saved.slice(0, 4)).toEqual([['Item', 'Qty'], ['c', '2'], ['b', '3'], ['a', '10']]);
    s.click('Sort Z');
    expect(grid(s.model).slice(1, 4)).toEqual([['a', '10'], ['b', '3'], ['c', '2']]); // 10, 3, 2: numbers, not text
    s.undo();
    s.undo();
    expect(grid(s.model)).toEqual(before);
    s.redo();
    expect(grid(s.model).slice(1, 4)).toEqual([['c', '2'], ['b', '3'], ['a', '10']]);
  });

  it('sorts text A to Z and Z to A within a selected range', async () => {
    const s = await session(format, await bytes());
    selectCell(1, 0, 3, 1); // b 3, a 10, c 2
    s.click('Sort A');
    expect(grid(s.model).slice(1, 4)).toEqual([['a', '10'], ['b', '3'], ['c', '2']]);
    s.click('Sort Z');
    expect(grid(s.model).slice(1, 4)).toEqual([['c', '2'], ['b', '3'], ['a', '10']]);
    expect(grid(await s.reload()).slice(1, 4)).toEqual([['c', '2'], ['b', '3'], ['a', '10']]);
  });

  if (format !== 'csv') {
    it('refuses to sort over formulas', async () => {
      const s = await session(format, await bytes());
      selectCell(3, 1, 5, 1);
      s.click('Sort A');
      expect(s.done).toHaveLength(0);
      expect(root.querySelector('.banner')!.textContent).toMatch(/formulas/);
    });
  }
});

describe.each(FORMATS.filter((f) => f.formulas))('$format: AutoSum', ({ format, bytes }) => {
  it('sums the selected column below it, and the block above the active cell', async () => {
    const s = await session(format, await bytes());
    selectCell(1, 1, 3, 1);
    s.click('AutoSum');
    expect(s.model.sheets[0]!.rows[4]![1]!.input).toBe('=SUM(B2:B4)');
    expect(s.model.sheets[0]!.rows[4]![1]!.text).toBe('15');
    s.undo();
    expect(grid(s.model)[4]).toEqual(['', '']);
    selectCell(4, 1);
    s.click('AutoSum');
    expect(s.model.sheets[0]!.rows[4]![1]!.input).toBe('=SUM(B2:B4)');
    const reloaded = await s.reload();
    expect(reloaded.sheets[0]!.rows[4]![1]!.formula).toBe('SUM(B2:B4)');
    expect(reloaded.sheets[0]!.rows[4]![1]!.text).toBe('15');
  });

  it('sums a selected row to its right', async () => {
    const s = await session(format, await bytes());
    selectCell(1, 0); // "b": no number above or to its left
    s.click('AutoSum');
    expect(s.done).toHaveLength(0);
    selectCell(2, 0, 2, 1); // a, 10
    s.click('AutoSum');
    expect(s.model.sheets[0]!.rows[2]![2]!.input).toBe('=SUM(A3:B3)');
  });

  it('works with Alt+=', async () => {
    const s = await session(format, await bytes());
    selectCell(1, 1, 3, 1);
    root.querySelector('.grid-scroll')!.dispatchEvent(new KeyboardEvent('keydown', { key: '=', altKey: true, bubbles: true }));
    expect(s.done).toHaveLength(1);
  });
});

describe('csv: AutoSum is off', () => {
  it('disables the button', async () => {
    await session('csv', enc('a,b\n1,2\n'));
    expect(document.body.querySelector<HTMLButtonElement>('button[title^="CSV files have no formulas"]')!.disabled).toBe(true);
  });
});

describe.each(FORMATS)('$format: insert and delete rows and columns', ({ format, bytes, formulas }) => {
  it('inserts a row above the selected row; undo and redo are exact; saves and reloads', async () => {
    const s = await session(format, await bytes());
    const before = grid(s.model);
    selectRow(2);
    s.click('Insert row above');
    const after = grid(s.model);
    expect(after.slice(0, 5)).toEqual([before[0], before[1], ['', ''], before[2], before[3]]);
    if (formulas) expect(s.model.sheets[0]!.rows[6]![1]!.input).toBe('=SUM(B2:B5)'); // the range grew, the formula moved down
    expect(grid(await s.reload())).toEqual(after);
    s.undo();
    expect(grid(s.model)).toEqual(before);
    expect(grid(await s.reload())).toEqual(before);
    s.redo();
    expect(grid(s.model)).toEqual(after);
  });

  it('deletes rows (2 selected) and undo puts them back', async () => {
    const s = await session(format, await bytes());
    const before = grid(s.model);
    mouse(root.querySelector('th.rowhead[data-r="1"]')!, 'mousedown');
    mouse(root.querySelector('th.rowhead[data-r="2"]')!, 'mousedown', { shiftKey: true });
    s.click('Delete the selected rows');
    const after = grid(s.model);
    expect(after[1]).toEqual(['c', '2']);
    if (formulas) expect(s.model.sheets[0]!.rows[2]![0]).toBeFalsy();
    expect(grid(await s.reload())).toEqual(after);
    s.undo();
    expect(grid(s.model)).toEqual(before);
    expect(grid(await s.reload())).toEqual(before);
    s.redo();
    expect(grid(s.model)).toEqual(after);
  });

  it('inserts and deletes a column', async () => {
    const s = await session(format, await bytes());
    const before = grid(s.model);
    selectCol(1);
    s.click('Insert column left');
    const inserted = grid(s.model);
    expect(inserted[1]).toEqual(['b', '', '3']);
    if (formulas) expect(s.model.sheets[0]!.rows[5]![2]!.input).toBe('=SUM(C2:C4)');
    expect(grid(await s.reload())).toEqual(inserted);
    s.undo();
    expect(grid(s.model)).toEqual(before);
    selectCol(0);
    s.click('Delete the selected columns');
    expect(grid(s.model)[1]).toEqual(['3']);
    if (formulas) expect(s.model.sheets[0]!.rows[5]![0]!.input).toBe('=SUM(A2:A4)');
    expect(grid(await s.reload())).toEqual(grid(s.model));
    s.undo();
    expect(grid(s.model)).toEqual(before);
    expect(grid(await s.reload())).toEqual(before);
  });

  it('works from the right-click menu on a row and a column header', async () => {
    const s = await session(format, await bytes());
    const before = grid(s.model);
    root.querySelector('th.rowhead[data-r="1"]')!.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
    const items = Array.from(root.querySelectorAll('.ctx-menu button')).map((b) => b.textContent);
    expect(items).toEqual(['Insert row above', 'Insert row below', 'Delete row']);
    (root.querySelectorAll('.ctx-menu button')[1] as HTMLButtonElement).click(); // below row 2
    expect(grid(s.model)[2]).toEqual(['', '']);
    expect(root.querySelector('.ctx-menu')).toBeNull();
    s.undo();
    expect(grid(s.model)).toEqual(before);
    root.querySelector('th.colhead[data-c="0"]')!.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
    expect(Array.from(root.querySelectorAll('.ctx-menu button')).map((b) => b.textContent)).toEqual(['Insert column left', 'Insert column right', 'Delete column', 'Sort A → Z', 'Sort Z → A']);
    (root.querySelectorAll('.ctx-menu button')[3] as HTMLButtonElement).click();
    expect(grid(s.model).slice(1, 4).map((r) => r[0])).toEqual(['a', 'b', 'c']);
  });
});

describe('formulas are rewritten and undone', () => {
  it('xlsx: deleting a referenced row gives #REF!, undo restores the formula; other sheets follow', async () => {
    const s = await session('xlsx', await xlsxData());
    const sheetOther = () => s.model.sheets[1]!.rows[0]![0]!.input;
    expect(sheetOther()).toBe('=S!B3*2');
    selectRow(1);
    s.click('Insert row above');
    expect(sheetOther()).toBe('=S!B4*2'); // the formula on the other sheet followed the cell
    expect(s.model.sheets[1]!.rows[0]![0]!.text).toBe('20');
    s.undo();
    expect(sheetOther()).toBe('=S!B3*2');
    selectRow(2);
    s.click('Delete the selected rows'); // deletes the row B3 lives in
    expect(sheetOther()).toBe('=#REF!*2');
    expect(s.model.sheets[0]!.rows[4]![1]!.input).toBe('=SUM(B2:B3)');
    s.undo();
    expect(sheetOther()).toBe('=S!B3*2');
    expect(grid(s.model, 0)[5]).toEqual(['', '=SUM(B2:B4)']);
    const re = await s.reload();
    expect(re.sheets[1]!.rows[0]![0]!.input).toBe('=S!B3*2');
    expect(re.sheets[0]!.rows[5]![1]!.input).toBe('=SUM(B2:B4)');
  });

  it('xlsx: the styles of a deleted row come back with undo', async () => {
    const s = await session('xlsx', await xlsxData());
    selectRow(2);
    s.click('Delete the selected rows');
    s.undo();
    const re = await s.reload();
    expect(re.sheets[0]!.rows[2]![0]!.style?.bold).toBe(true);
    expect(re.sheets[0]!.rows[2]![1]!.text).toBe('10.00'); // number format kept
  });

  it('xlsx: shared formulas survive an insert', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.addRow([1, 2]);
    ws.addRow([3, 4]);
    ws.getCell('C1').value = { formula: 'A1+B1', result: 3 };
    ws.getCell('C2').value = { sharedFormula: 'C1', result: 7 };
    const bytes = new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
    const s = await session('xlsx', bytes);
    expect(s.model.sheets[0]!.rows[1]![2]!.formula).toBe('A2+B2');
    selectRow(0);
    s.click('Insert row above');
    const re = await s.reload();
    expect(re.sheets[0]!.rows[1]![2]!.formula).toBe('A2+B2');
    expect(re.sheets[0]!.rows[2]![2]!.formula).toBe('A3+B3');
  });

  it('ods: formulas are rewritten and saved', async () => {
    const s = await session('ods', makeOdsData());
    selectRow(2);
    s.click('Insert row above');
    const re = await s.reload();
    expect(re.sheets[0]!.rows[6]![1]!.formula).toBe('SUM(B2:B5)');
    expect(re.sheets[0]!.rows[6]![1]!.text).toBe('15');
    s.undo();
    expect((await s.reload()).sheets[0]!.rows[5]![1]!.formula).toBe('SUM(B2:B4)');
  });
});

describe('limits', () => {
  it('refuses to insert or delete in a sheet with merged cells', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.addRow(['a', 'b']);
    ws.addRow(['c', 'd']);
    ws.mergeCells('A1:B1');
    const s = await session('xlsx', new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer));
    selectRow(1);
    s.click('Insert row above');
    expect(s.done).toHaveLength(0);
    expect(root.querySelector('.banner')!.textContent).toMatch(/merged/);
  });

  it('does nothing below the data', async () => {
    const s = await session('csv', enc('a,b\n1,2\n'));
    selectRow(10);
    s.click('Delete the selected rows');
    expect(s.done).toHaveLength(0);
  });
});
