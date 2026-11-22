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

/** jsdom has no layout: buttons are 30px wide, separators 1px, the info text and the spacer 0px. Sets the row width and relays out. */
function fit(tb: ReturnType<typeof createToolbar>, width: number) {
  const row = tb.el.querySelector<HTMLElement>('.toolbar-row')!;
  for (const e of Array.from(tb.el.querySelectorAll<HTMLElement>('button, .sep, .info, .tb-spacer'))) {
    const w = e.classList.contains('sep') ? 1 : e.classList.contains('info') || e.classList.contains('tb-spacer') ? 0 : 30;
    Object.defineProperty(e, 'offsetWidth', { value: w, configurable: true });
  }
  Object.defineProperty(row, 'clientWidth', { value: width, configurable: true });
  Object.defineProperty(row, 'scrollWidth', {
    configurable: true,
    get: () => Array.from(row.children).reduce((sum, c) => sum + (c as HTMLElement).offsetWidth, 0),
  });
  tb.relayout();
}

/** Undo, Redo, Save (pinned), |, Alpha, Wrap (toggle), Last, info, spacer. */
function overflowToolbar() {
  const log: string[] = [];
  const tb = createToolbar();
  tb.button('', 'Undo', () => log.push('undo'), 'undo');
  tb.button('', 'Redo', () => log.push('redo'), 'redo');
  tb.button('', 'Save', () => log.push('save'), 'save');
  tb.separator();
  tb.button('', 'Alpha', () => log.push('alpha'), 'bold');
  tb.toggle('Wrap', 'Wrap text', (on) => log.push(`wrap:${on}`), 'wrap');
  tb.button('', 'Last', () => log.push('last'), 'table');
  tb.info();
  tb.spacer();
  root.append(tb.el);
  const row = tb.el.querySelector<HTMLElement>('.toolbar-row')!;
  const more = tb.el.querySelector<HTMLButtonElement>('.tb-more')!;
  const menu = tb.el.querySelector<HTMLElement>('.tb-menu')!;
  const rowTitles = () => Array.from(row.children).filter((c) => c.classList.contains('tb-btn') && c !== more).map((c) => c.getAttribute('title'));
  const menuTitles = () => Array.from(menu.querySelectorAll('button')).map((b) => b.getAttribute('title'));
  return { tb, log, row, more, menu, rowTitles, menuTitles };
}

describe('toolbar overflow menu', () => {
  it('shows no More button when everything fits', () => {
    const t = overflowToolbar();
    fit(t.tb, 400);
    expect(t.more.hidden).toBe(true);
    expect(t.menu.children).toHaveLength(0);
    expect(t.rowTitles()).toEqual(['Undo', 'Redo', 'Save', 'Alpha', 'Wrap text', 'Last']);
  });

  it('moves the last buttons into the menu, in their order, and shows More', () => {
    const t = overflowToolbar();
    fit(t.tb, 160);
    expect(t.more.hidden).toBe(false);
    expect(t.more.getAttribute('aria-haspopup')).toBe('menu');
    expect(t.rowTitles()).toEqual(['Undo', 'Redo', 'Save', 'Alpha']);
    expect(t.menuTitles()).toEqual(['Wrap text', 'Last']);
    // icon-only items show their tooltip as text in the menu
    expect(t.menu.querySelector<HTMLElement>('[title="Last"]')!.dataset['menuLabel']).toBe('Last');
  });

  it('runs a moved button and closes the menu', () => {
    const t = overflowToolbar();
    fit(t.tb, 160);
    t.more.click();
    expect(t.menu.hidden).toBe(false);
    expect(t.more.getAttribute('aria-expanded')).toBe('true');
    t.menu.querySelector<HTMLButtonElement>('[title="Last"]')!.click();
    expect(t.log).toEqual(['last']);
    expect(t.menu.hidden).toBe(true);
    expect(t.more.getAttribute('aria-expanded')).toBe('false');
  });

  it('shows a moved toggle as a menuitemcheckbox with aria-checked, and back as a pressed button', () => {
    const t = overflowToolbar();
    fit(t.tb, 160);
    const wrap = t.menu.querySelector<HTMLButtonElement>('[title="Wrap text"]')!;
    expect(wrap.getAttribute('role')).toBe('menuitemcheckbox');
    expect(wrap.hasAttribute('aria-pressed')).toBe(false);
    expect(wrap.getAttribute('aria-checked')).toBe('false');
    expect(t.menu.querySelector('[title="Last"]')!.getAttribute('role')).toBe('menuitem');
    wrap.click();
    expect(t.log).toEqual(['wrap:true']);
    expect(wrap.getAttribute('aria-checked')).toBe('true');
    fit(t.tb, 400); // back in the row, still pressed
    expect(t.tb.el.querySelector('[title="Wrap text"]')!.hasAttribute('aria-checked')).toBe(false);
    expect(t.tb.el.querySelector('[title="Wrap text"]')!.getAttribute('aria-pressed')).toBe('true');
  });

  it('moves the buttons back in their original order when widened', () => {
    const t = overflowToolbar();
    fit(t.tb, 160);
    fit(t.tb, 400);
    expect(t.more.hidden).toBe(true);
    expect(t.menu.children).toHaveLength(0);
    expect(t.rowTitles()).toEqual(['Undo', 'Redo', 'Save', 'Alpha', 'Wrap text', 'Last']);
    expect(t.row.querySelector<HTMLElement>('.sep')!.hidden).toBe(false);
    fit(t.tb, 160); // and narrow again: same items, same order
    expect(t.menuTitles()).toEqual(['Wrap text', 'Last']);
  });

  it('never shows an empty menu', () => {
    const log: string[] = [];
    const tb = createToolbar();
    tb.button('', 'Undo', () => log.push('undo'), 'undo');
    tb.button('', 'Redo', () => log.push('redo'), 'redo');
    tb.button('', 'Save', () => log.push('save'), 'save');
    root.append(tb.el);
    fit(tb, 40); // only pinned buttons, nothing may move
    const more = tb.el.querySelector<HTMLElement>('.tb-more')!;
    expect(more.hidden).toBe(true);
    expect(tb.el.querySelector('.tb-menu')!.children).toHaveLength(0);
  });

  it('closes on Escape and returns focus to More', () => {
    const t = overflowToolbar();
    fit(t.tb, 160);
    t.more.click();
    const first = t.menu.querySelector<HTMLButtonElement>('button')!;
    expect(document.activeElement).toBe(first);
    first.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(t.menu.hidden).toBe(true);
    expect(document.activeElement).toBe(t.more);
  });

  it('moves between items with the arrow keys', () => {
    const t = overflowToolbar();
    fit(t.tb, 160);
    t.more.click();
    const [wrap, last] = Array.from(t.menu.querySelectorAll('button'));
    expect(document.activeElement).toBe(wrap);
    t.menu.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(document.activeElement).toBe(last);
    t.menu.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(document.activeElement).toBe(wrap); // wraps around
    t.menu.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowUp', bubbles: true }));
    expect(document.activeElement).toBe(last);
  });

  it('closes when clicking outside', () => {
    const t = overflowToolbar();
    fit(t.tb, 160);
    t.more.click();
    expect(t.menu.hidden).toBe(false);
    document.body.click();
    expect(t.menu.hidden).toBe(true);
  });
});
