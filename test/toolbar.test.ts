// @vitest-environment jsdom
import { beforeEach, describe, expect, it } from 'vitest';
import type { SheetModel, TextModel } from '../src/shared/model.ts';
import { createToolbar } from '../src/webview/ui/toolbar.ts';
import { renderSheetView } from '../src/webview/views/sheetView.ts';
import { renderTextView } from '../src/webview/views/textView.ts';

let root: HTMLElement;
beforeEach(() => {
  document.body.innerHTML = '<div id="root"></div>';
  root = document.getElementById('root')!;
});

const btn = (tb: ReturnType<typeof createToolbar>, label: string) =>
  Array.from(tb.el.querySelectorAll('button')).find(
    (b) => b.textContent === label || b.getAttribute('aria-label') === label || b.title.startsWith(label),
  )!;

const sheet = (): SheetModel => ({
  type: 'sheet',
  sheets: [
    {
      name: 'S',
      rows: [[{ text: 'Name' }, { text: 'Qty' }], [{ text: 'Apple' }, { text: '3' }], [{ text: 'Pear' }, { text: '4', formula: 'B2+1' }], [{ text: 'apple pie' }, { text: '9' }]],
      rowCount: 4,
      colCount: 2,
      merges: [],
      colWidths: [null, null],
      truncated: false,
    },
    { name: 'T', rows: [[{ text: 'needle' }]], rowCount: 1, colCount: 1, merges: [], colWidths: [null], truncated: false },
  ],
});

describe('sheet toolbar', () => {
  it('shows the address and formula of the selected cell', () => {
    renderSheetView(root, sheet(), false, () => {});
    root.querySelector<HTMLElement>('td[data-r="2"][data-c="1"]')!.click();
    expect(root.querySelector('.name-box')!.textContent).toBe('B3');
    expect(root.querySelector('.formula-text')!.textContent).toBe('=B2+1');
  });

  it('filters rows by column text and keeps the header row', () => {
    const tb = createToolbar();
    renderSheetView(root, sheet(), false, () => {}, tb);
    btn(tb, 'Filter').click();
    const input = root.querySelector<HTMLInputElement>('th.filter input')!;
    input.value = 'APPLE';
    input.dispatchEvent(new Event('input'));
    const hidden = (r: number) => root.querySelector<HTMLElement>(`tr[data-r="${r}"]`)!.hidden;
    expect([hidden(0), hidden(1), hidden(2), hidden(3)]).toEqual([false, false, true, false]);
    btn(tb, 'Filter').click(); // off again: everything visible
    expect(root.querySelectorAll('tr[hidden]').length).toBe(0);
  });

  it('freezes the first row', () => {
    const tb = createToolbar();
    renderSheetView(root, sheet(), false, () => {}, tb);
    btn(tb, 'Freeze row').click();
    expect(root.querySelector('table.grid')!.classList.contains('freeze')).toBe(true);
  });

  it('finds text across sheets and selects the match', () => {
    const tb = createToolbar();
    renderSheetView(root, sheet(), false, () => {}, tb);
    btn(tb, 'Find').click();
    const input = tb.extra.querySelector<HTMLInputElement>('input')!;
    input.value = 'needle';
    input.dispatchEvent(new Event('input'));
    expect(tb.extra.querySelector('.find-status')!.textContent).toBe('1 of 1');
    expect(root.querySelector('.tab.active')!.textContent).toBe('T');
    expect(root.querySelector('td.sel')!.textContent).toBe('needle');
    input.value = 'zzz';
    input.dispatchEvent(new Event('input'));
    expect(tb.extra.querySelector('.find-status')!.textContent).toBe('No results');
  });

  it('cycles through matches with Enter', () => {
    const tb = createToolbar();
    renderSheetView(root, sheet(), false, () => {}, tb);
    btn(tb, 'Find').click();
    const input = tb.extra.querySelector<HTMLInputElement>('input')!;
    input.value = 'apple';
    input.dispatchEvent(new Event('input'));
    expect(root.querySelector('td.sel')!.textContent).toBe('Apple');
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(root.querySelector('td.sel')!.textContent).toBe('apple pie');
  });
});

describe('document toolbar', () => {
  const doc = (): TextModel => ({ type: 'text', html: '<h1 data-p="0">Title</h1><p data-p="1">one two three</p><h2 data-p="2">Part</h2>' });

  it('counts words and updates after an edit', () => {
    const tb = createToolbar();
    const view = renderTextView(root, doc(), false, () => {}, tb);
    expect(tb.el.querySelector('.info')!.textContent).toBe('5 words');
    view.applyEdit({ kind: 'paragraph', index: 1, text: 'x', old: 'one two three' }, {});
    expect(tb.el.querySelector('.info')!.textContent).toBe('3 words');
  });

  it('lists headings in the navigation pane', () => {
    const tb = createToolbar();
    renderTextView(root, doc(), false, () => {}, tb);
    const nav = root.querySelector<HTMLElement>('.doc-nav')!;
    expect(nav.hidden).toBe(true);
    btn(tb, 'Outline').click();
    expect(nav.hidden).toBe(false);
    expect(Array.from(nav.querySelectorAll('.nav-item')).map((b) => b.textContent)).toEqual(['Title', 'Part']);
  });

  it('toggles formatting marks and zoom', () => {
    const tb = createToolbar();
    renderTextView(root, doc(), false, () => {}, tb);
    btn(tb, 'Show formatting marks').click();
    expect(root.querySelector('.page')!.classList.contains('marks')).toBe(true);
    btn(tb, 'Zoom in').click();
    expect(root.querySelector<HTMLElement>('.doc-layout')!.style.getPropertyValue('--zoom')).toBe('1.1');
    expect(btn(tb, '110%')).toBeTruthy();
    btn(tb, '110%').click();
    expect(root.querySelector<HTMLElement>('.doc-layout')!.style.getPropertyValue('--zoom')).toBe('1');
  });

  it('opens the find bar with Ctrl+F and counts matches', () => {
    const tb = createToolbar();
    renderTextView(root, doc(), false, () => {}, tb);
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'f', ctrlKey: true, bubbles: true }));
    const bar = tb.extra.querySelector<HTMLElement>('.findbar')!;
    expect(bar.hidden).toBe(false);
    const input = bar.querySelector('input')!;
    input.value = 'T';
    input.dispatchEvent(new Event('input'));
    expect(bar.querySelector('.find-status')!.textContent).toBe('1 of 5'); // Title(2), two(1), three(1), Part(1)
  });
});
