// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { Edit, SheetModel, TextModel } from '../src/shared/model.ts';
import { columnName, readableColor, renderSheetView } from '../src/webview/views/sheetView.ts';
import { renderSlidesView } from '../src/webview/views/slidesView.ts';
import { renderTextView } from '../src/webview/views/textView.ts';

const sheetModel = (): SheetModel => ({
  type: 'sheet',
  sheets: [
    {
      name: 'One',
      rows: [[{ text: 'a' }, { text: 'b', style: { bold: true, fill: '#ff0000' } }], [null, { text: '5', formula: 'A1+1', input: '=A1+1' }]],
      rowCount: 2,
      colCount: 2,
      merges: [{ r1: 0, c1: 0, r2: 0, c2: 1 }],
      colWidths: [20, null],
      truncated: false,
    },
    { name: 'Two', rows: [[{ text: 'z' }]], rowCount: 1, colCount: 1, merges: [], colWidths: [null], truncated: true },
  ],
});

let root: HTMLElement;
beforeEach(() => {
  document.body.className = 'vscode-dark';
  document.body.innerHTML = '<div id="root"></div>';
  root = document.getElementById('root')!;
});

const td = (r: number, c: number) => root.querySelector<HTMLElement>(`td[data-r="${r}"][data-c="${c}"]`)!;

describe('sheet view', () => {
  it('names columns like a spreadsheet', () => {
    expect([0, 25, 26, 27, 51, 52, 701, 702].map(columnName)).toEqual(['A', 'Z', 'AA', 'AB', 'AZ', 'BA', 'ZZ', 'AAA']);
  });

  it('renders cells, merges, styles, tooltips and tabs', () => {
    renderSheetView(root, sheetModel(), false, () => {});
    expect(td(0, 0).textContent).toBe('a');
    expect((td(0, 0) as HTMLTableCellElement).colSpan).toBe(2);
    expect(td(0, 1)).toBeNull(); // covered by the merge
    expect(td(1, 1).title).toBe('=A1+1');
    expect(root.querySelectorAll('.tab').length).toBe(2);
    expect(root.querySelectorAll('thead th.colhead').length).toBeGreaterThanOrEqual(10);
    renderSheetView(root, { type: 'sheet', sheets: [{ ...sheetModel().sheets[0]!, merges: [] }] }, false, () => {});
    expect(td(0, 1).style.fontWeight).toBe('bold');
    expect(td(0, 1).style.backgroundColor).not.toBe('');
  });

  it('switches sheets and shows the truncation notice', () => {
    renderSheetView(root, sheetModel(), false, () => {});
    expect(root.querySelector<HTMLElement>('.banner')!.hidden).toBe(true);
    root.querySelectorAll<HTMLButtonElement>('.tab')[1]!.click();
    expect(td(0, 0).textContent).toBe('z');
    expect(root.querySelector<HTMLElement>('.banner')!.hidden).toBe(false);
  });

  it('edits a cell on double click: Enter commits and reports the edit with the old value', () => {
    const edits: Edit[] = [];
    renderSheetView(root, sheetModel(), false, (e) => edits.push(e));
    td(1, 1).dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    const input = root.querySelector<HTMLInputElement>('input.cell-input')!;
    expect(input.value).toBe('=A1+1'); // formulas are edited as typed
    input.value = '42';
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(edits).toEqual([{ kind: 'cell', sheet: 0, row: 1, col: 1, value: '42', old: '=A1+1' }]);
    expect(td(1, 1).textContent).toBe('42'); // optimistic update
    expect(root.querySelector('input.cell-input')).toBeNull();
  });

  it('Escape cancels, unchanged values report nothing', () => {
    const edits: Edit[] = [];
    renderSheetView(root, sheetModel(), false, (e) => edits.push(e));
    td(0, 0).dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    const input = root.querySelector<HTMLInputElement>('input.cell-input')!;
    input.value = 'changed';
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(edits).toEqual([]);
    expect(td(0, 0).textContent).toBe('a');
    td(0, 0).dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    root.querySelector<HTMLInputElement>('input.cell-input')!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(edits).toEqual([]);
  });

  it('typing on a selected cell starts editing; Delete clears', () => {
    const edits: Edit[] = [];
    renderSheetView(root, sheetModel(), false, (e) => edits.push(e));
    const scroller = root.querySelector<HTMLElement>('.grid-scroll')!;
    td(0, 0).dispatchEvent(new MouseEvent('click', { bubbles: true }));
    scroller.dispatchEvent(new KeyboardEvent('keydown', { key: 'x', bubbles: true }));
    expect(root.querySelector<HTMLInputElement>('input.cell-input')!.value).toBe('x');
    root.querySelector<HTMLInputElement>('input.cell-input')!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    td(1, 1).dispatchEvent(new MouseEvent('click', { bubbles: true }));
    scroller.dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', bubbles: true }));
    expect(edits).toEqual([{ kind: 'cell', sheet: 0, row: 1, col: 1, value: '', old: '=A1+1' }]);
  });

  it('is not editable when read-only', () => {
    const onEdit = vi.fn();
    renderSheetView(root, sheetModel(), true, onEdit);
    td(0, 0).dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    expect(root.querySelector('input.cell-input')).toBeNull();
    expect(onEdit).not.toHaveBeenCalled();
  });

  it('applies results and host edits (undo/redo) to the displayed cell', () => {
    const view = renderSheetView(root, sheetModel(), false, () => {});
    view.applyEdit({ kind: 'cell', sheet: 0, row: 0, col: 0, value: 'new', old: 'a' }, { cell: { text: 'NEW' } });
    expect(td(0, 0).textContent).toBe('NEW');
    view.applyEdit({ kind: 'cell', sheet: 0, row: 1, col: 1, value: '', old: '5' }, { cell: null });
    expect(td(1, 1).textContent).toBe('');
    // edit in a sheet that is not shown only updates the data
    view.applyEdit({ kind: 'cell', sheet: 1, row: 0, col: 0, value: 'q', old: 'z' }, { cell: { text: 'q' } });
    root.querySelectorAll<HTMLButtonElement>('.tab')[1]!.click();
    expect(td(0, 0).textContent).toBe('q');
  });

  it('virtualises large sheets: only a window of rows is in the DOM', () => {
    const rows = Array.from({ length: 5000 }, (_, i) => [{ text: `r${i}` }]);
    renderSheetView(root, { type: 'sheet', sheets: [{ name: 'Big', rows, rowCount: 5000, colCount: 1, merges: [], colWidths: [null], truncated: false }] }, false, () => {});
    const rendered = root.querySelectorAll('tbody tr:not(.spacer)').length;
    expect(rendered).toBeGreaterThan(10);
    expect(rendered).toBeLessThan(200);
    expect(td(0, 0).textContent).toBe('r0');
    expect(td(4000, 0)).toBeNull();
    expect(root.querySelector('tr.spacer')).not.toBeNull();
  });

  it('keeps text readable in both themes', () => {
    document.body.className = 'vscode-dark';
    expect(readableColor({ color: '#000000' })).toBeUndefined(); // black on dark theme is dropped
    expect(readableColor({ color: '#ff0000' })).toBe('#ff0000');
    expect(readableColor({ fill: '#000000' })).toBe('#ffffff');
    expect(readableColor({ fill: '#ffffff' })).toBe('#000000');
    document.body.className = 'vscode-light';
    expect(readableColor({ color: '#000000' })).toBe('#000000');
    expect(readableColor({ color: '#ffffff' })).toBeUndefined();
  });
});

describe('text view', () => {
  const model: TextModel = { type: 'text', html: '<h1 data-p="0">Title</h1><p data-p="1">Body</p><p>Read only</p>' };

  it('makes mapped paragraphs editable and leaves others alone', () => {
    renderTextView(root, model, false, () => {});
    expect(root.querySelector('[data-p="0"]')!.getAttribute('contenteditable')).toBe('plaintext-only');
    expect(root.querySelector('p:not([data-p])')!.hasAttribute('contenteditable')).toBe(false);
  });

  it('reports an edit on focus out, with the old text', () => {
    const edits: Edit[] = [];
    renderTextView(root, model, false, (e) => edits.push(e));
    const p = root.querySelector<HTMLElement>('[data-p="1"]')!;
    p.textContent = 'Body changed';
    p.dispatchEvent(new FocusEvent('focusout', { bubbles: true }));
    expect(edits).toEqual([{ kind: 'paragraph', index: 1, text: 'Body changed', old: 'Body' }]);
    p.dispatchEvent(new FocusEvent('focusout', { bubbles: true }));
    expect(edits).toHaveLength(1); // unchanged text: no second edit
    p.textContent = 'Again';
    p.dispatchEvent(new FocusEvent('focusout', { bubbles: true }));
    expect(edits[1]).toEqual({ kind: 'paragraph', index: 1, text: 'Again', old: 'Body changed' });
  });

  it('Escape restores the text and reports nothing', () => {
    const edits: Edit[] = [];
    renderTextView(root, model, false, (e) => edits.push(e));
    const p = root.querySelector<HTMLElement>('[data-p="1"]')!;
    p.textContent = 'oops';
    p.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(p.textContent).toBe('Body');
    expect(edits).toEqual([]);
  });

  it('is read-only on demand and applies host edits', () => {
    const view = renderTextView(root, model, true, () => {});
    expect(root.querySelector('[contenteditable]')).toBeNull();
    view.applyEdit({ kind: 'paragraph', index: 1, text: 'From host', old: 'Body' }, {});
    expect(root.querySelector('[data-p="1"]')!.textContent).toBe('From host');
  });
});

describe('slides view', () => {
  it('renders text with textContent only (no HTML injection)', () => {
    renderSlidesView(root, {
      type: 'slides',
      slides: [{ number: 1, title: '<img src=x onerror=alert(1)>', paragraphs: ['<b>x</b>'], notes: 'n', imageDataUri: 'data:image/png;base64,AAAA' }],
    });
    expect(root.querySelector('h2')!.textContent).toContain('<img src=x onerror=alert(1)>');
    expect(root.querySelectorAll('img').length).toBe(1); // only the real picture
    expect(root.querySelector('li')!.textContent).toBe('<b>x</b>');
    expect(root.querySelector('details')).not.toBeNull();
  });
});
