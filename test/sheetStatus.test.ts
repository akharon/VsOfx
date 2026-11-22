// @vitest-environment jsdom
import { beforeEach, describe, expect, it } from 'vitest';
import type { Edit, SheetModel } from '../src/shared/model.ts';
import { createToolbar } from '../src/webview/ui/toolbar.ts';
import { numericValue, renderSheetView } from '../src/webview/views/sheetView.ts';

let root: HTMLElement;
beforeEach(() => {
  document.body.innerHTML = '<div id="root"></div>';
  root = document.getElementById('root')!;
});

const model = (): SheetModel => ({
  type: 'sheet',
  sheetOps: true,
  sheets: [
    {
      name: 'S',
      rows: [
        [{ text: '10' }, { text: 'Apple' }],
        [{ text: '-2.5' }, null],
        [{ text: '1,500' }, { text: 'x' }],
      ],
      rowCount: 3,
      colCount: 2,
      merges: [],
      colWidths: [],
      truncated: false,
    },
    {
      name: 'T',
      rows: [[{ text: '1' }, { text: '2' }]],
      rowCount: 1,
      colCount: 2,
      merges: [],
      colWidths: [],
      truncated: false,
    },
  ],
});

const td = (r: number, c: number) => root.querySelector<HTMLElement>(`td[data-r="${r}"][data-c="${c}"]`)!;
const status = () => root.querySelector('.sheet-status')!.textContent;
/** Click the first cell, then shift-click the other corner (mouse selection of a range). */
function select(r1: number, c1: number, r2 = r1, c2 = c1) {
  td(r1, c1).dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
  if (r1 !== r2 || c1 !== c2) td(r2, c2).dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0, shiftKey: true }));
}
const f = (n: number) => n.toLocaleString(undefined, { maximumFractionDigits: 10 });

describe('sheet status line', () => {
  it('shows Average, Count and Sum of the selection, Sum and Average over numbers only', () => {
    renderSheetView(root, model(), false, () => {}, createToolbar());
    select(0, 0, 2, 1); // 10, Apple, -2.5, (empty), 1,500, x
    expect(root.querySelector('.tabs .sheet-status')!.getAttribute('aria-live')).toBe('polite');
    expect(status()).toBe(`Average: ${f(502.5)}   Count: 5   Sum: ${f(1507.5)}`);
  });

  it('shows only Count when no numeric cell is selected', () => {
    renderSheetView(root, model(), false, () => {}, createToolbar());
    select(0, 1, 2, 1); // Apple, (empty), x
    expect(status()).toBe('Count: 2');
  });

  it('shows nothing for a single cell', () => {
    renderSheetView(root, model(), false, () => {}, createToolbar());
    select(0, 0, 2, 1);
    expect(status()).not.toBe('');
    select(1, 0);
    expect(status()).toBe('');
  });

  it('updates after a cell edit', () => {
    const edits: Edit[] = [];
    const view = renderSheetView(root, model(), false, (e) => edits.push(e), createToolbar());
    select(0, 0, 2, 1);
    const edit: Edit = { kind: 'cell', sheet: 0, row: 1, col: 0, value: '5', old: '-2.5' };
    view.applyEdit(edit, { cell: { text: '5' } }); // 10, 5, 1,500 are now the numbers
    expect(status()).toBe(`Average: ${f(505)}   Count: 5   Sum: ${f(1515)}`);
  });

  it('follows the active sheet', () => {
    renderSheetView(root, model(), false, () => {}, createToolbar());
    select(0, 0, 0, 1);
    expect(status()).toBe(`Average: ${f(10)}   Count: 2   Sum: ${f(10)}`); // 10 and Apple: one number
    root.querySelectorAll<HTMLElement>('[role="tab"]')[1]!.click(); // selection resets to A1 on sheet T
    expect(status()).toBe('');
    select(0, 0, 0, 1);
    expect(status()).toBe(`Average: ${f(1.5)}   Count: 2   Sum: ${f(3)}`);
  });

  it('parses displayed numbers: separators, minus, percent; text is not a number', () => {
    expect(numericValue('1,234.5')).toBe(1234.5);
    expect(numericValue('-3')).toBe(-3);
    expect(numericValue('50%')).toBe(0.5);
    expect(numericValue('Apple')).toBeUndefined();
    expect(numericValue('2024-01-01')).toBeUndefined();
  });
});
