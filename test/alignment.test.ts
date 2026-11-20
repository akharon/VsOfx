// @vitest-environment jsdom
import { strFromU8 } from 'fflate';
import { describe, expect, it } from 'vitest';
import { applyTextEdit, inverseEdit, isEdit, type Edit, type SheetModel, type TextEdit, type TextModel } from '../src/shared/model.ts';
import { cellIndex, type ParaFmt } from '../src/shared/richtext.ts';
import { createAdapter } from '../src/webview/adapters/index.ts';
import { readZip } from '../src/webview/adapters/zip.ts';
import { createToolbar } from '../src/webview/ui/toolbar.ts';
import { renderCodeView } from '../src/webview/views/codeView.ts';
import { renderTextView } from '../src/webview/views/textView.ts';
import { makeDocx, makeOds, makeOdt, makeXlsx } from './helpers/fixtures.ts';

const fmt = (t: string, extra: Partial<ParaFmt> = {}): ParaFmt => ({ runs: [{ t }], ...extra });
const content = (b: Uint8Array) => strFromU8(readZip(b)['content.xml']!);
const docXml = (b: Uint8Array) => strFromU8(readZip(b)['word/document.xml']!);

describe('alignment edits', () => {
  it('are validated and inverted', () => {
    const e: Edit = { kind: 'align', sheet: 0, row: 1, col: 2, align: { h: 'center', v: 'top' }, old: {} };
    expect(isEdit(e)).toBe(true);
    expect(inverseEdit(inverseEdit(e))).toEqual(e);
    expect(isEdit({ ...e, align: { h: 'middle' } })).toBe(false);
    expect(isEdit({ kind: 'paragraph', index: 0, text: 'a', old: 'a', fmt: fmt('a', { align: 'justify', valign: 'middle' }) })).toBe(true);
    expect(isEdit({ kind: 'paragraph', index: 0, text: 'a', old: 'a', fmt: { runs: [], align: 'up' } })).toBe(false);
  });
});

describe('spreadsheets: alignment', () => {
  for (const [format, bytes] of [
    ['xlsx', () => makeXlsx('simple')],
    ['ods', () => makeOds()],
  ] as const) {
    it(`${format}: saves horizontal and vertical alignment with wrap; undo restores`, async () => {
      const a = createAdapter(format);
      await a.load(await bytes(), []);
      const al: Edit = { kind: 'align', sheet: 0, row: 0, col: 0, align: { h: 'center', v: 'middle' }, old: {} };
      expect(a.applyEdit(al).cell?.style).toMatchObject({ align: 'center', valign: 'middle' });
      a.applyEdit({ kind: 'wrap', sheet: 0, row: 0, col: 0, wrap: true, old: false });
      let back = (await createAdapter(format).load(await a.save(), [])) as SheetModel;
      expect(back.sheets[0]!.rows[0]![0]?.style).toMatchObject({ align: 'center', valign: 'middle', wrap: true });
      a.applyEdit(inverseEdit(al));
      back = (await createAdapter(format).load(await a.save(), [])) as SheetModel;
      const st = back.sheets[0]!.rows[0]![0]?.style;
      expect(st?.align).toBeUndefined();
      expect(st?.valign).toBeUndefined();
      expect(st?.wrap).toBe(true);
    });
  }
});

describe('documents: alignment', () => {
  it('Word: writes w:jc and the cell w:vAlign of an inserted table', async () => {
    const a = createAdapter('docx');
    await a.load(makeDocx('simple'), []);
    a.applyEdit({ kind: 'paragraph', index: 0, text: 'First paragraph.', old: 'First paragraph.', fmt: fmt('First paragraph.', { align: 'center' }), oldFmt: fmt('First paragraph.') });
    const spec = { id: 1, after: 0, rows: 1, cols: 1 };
    a.applyEdit({ kind: 'table', op: 'insert', table: spec });
    a.applyEdit({ kind: 'paragraph', index: cellIndex(spec, 0, 0), text: 'X', old: '', fmt: fmt('X', { valign: 'bottom', align: 'right' }), oldFmt: fmt('') });
    const xml = docXml(await a.save());
    expect(xml).toMatch(/<w:jc w:val="center"\/>.*First paragraph/);
    expect(xml).toContain('<w:vAlign w:val="bottom"/>');
    expect(xml).toContain('<w:jc w:val="right"/>');
  });

  it('Writer: writes derived styles and renders them again', async () => {
    const a = createAdapter('odt');
    await a.load(makeOdt(), []);
    a.applyEdit({ kind: 'paragraph', index: 1, text: 'Mid', old: '', fmt: fmt('Mid', { align: 'center' }), oldFmt: fmt('') });
    const spec = { id: 1, after: 1, rows: 1, cols: 1 };
    a.applyEdit({ kind: 'table', op: 'insert', table: spec });
    a.applyEdit({ kind: 'paragraph', index: cellIndex(spec, 0, 0), text: 'C', old: '', fmt: fmt('C', { valign: 'middle' }), oldFmt: fmt('') });
    const saved = await a.save();
    expect(content(saved)).toMatch(/fo:text-align="center"/);
    expect(content(saved)).toMatch(/style:vertical-align="middle"/);
    const m = (await createAdapter('odt').load(saved, [])) as TextModel;
    expect(m.html).toMatch(/<p data-p="1" data-align="center">Mid<\/p>/);
    expect(m.html).toContain('data-valign="middle"');
  });

  it('view: the alignment buttons change the paragraph and the cell', () => {
    document.body.innerHTML = '<div id="root"></div>';
    const root = document.getElementById('root')!;
    const edits: Edit[] = [];
    const model: TextModel = { type: 'text', html: '<p data-p="0">One</p><table><tbody><tr><td><p data-p="1">Cell</p></td></tr></tbody></table>', rich: true };
    const tb = createToolbar();
    document.body.appendChild(tb.el);
    const view = renderTextView(root, model, false, (e) => edits.push(e), tb);
    const btn = (t: string) => Array.from(tb.el.querySelectorAll('button')).find((b) => b.title.startsWith(t))!;
    const p0 = root.querySelector<HTMLElement>('[data-p="0"]')!;
    p0.dispatchEvent(new FocusEvent('focusin', { bubbles: true }));
    btn('Horizontal').click();
    expect(p0.dataset['align']).toBe('center');
    expect(edits.at(-1)).toMatchObject({ kind: 'paragraph', index: 0, fmt: { align: 'center' } });
    const p1 = root.querySelector<HTMLElement>('[data-p="1"]')!;
    p1.dispatchEvent(new FocusEvent('focusin', { bubbles: true }));
    btn('Vertical').click();
    expect(p1.closest('td')!.dataset['valign']).toBe('middle');
    expect(edits.at(-1)).toMatchObject({ kind: 'paragraph', index: 1, fmt: { valign: 'middle' } });
    view.applyEdit(inverseEdit(edits.at(-1)!), {});
    expect(p1.closest('td')!.dataset['valign']).toBeUndefined();
    view.dispose();
  });
});

describe('Markdown: table size and column alignment', () => {
  const setup = (src: string, source = false) => {
    document.body.innerHTML = '<div id="root"></div>';
    localStorage.setItem('vsofx.mode.markdown', source ? 'source' : 'visual');
    const root = document.getElementById('root')!;
    const edits: Edit[] = [];
    const tb = createToolbar();
    document.body.appendChild(tb.el);
    const view = renderCodeView(root, { type: 'code', lang: 'markdown', text: src }, false, (e) => edits.push(e), tb);
    const btn = (t: string) => Array.from(tb.el.querySelectorAll('button')).find((b) => b.title.startsWith(t))!;
    const bar = tb.el.querySelector<HTMLElement>('.md-tablebar')!;
    let text = src;
    const result = () => {
      view.flush?.();
      for (const e of edits.splice(0)) text = applyTextEdit(text, e as TextEdit);
      return text;
    };
    return { root, btn, bar, view, result };
  };

  it('source: asks for the size and inserts that table', () => {
    const { btn, bar, result, view } = setup('Intro\n', true);
    const ta = document.querySelector('textarea')!;
    ta.selectionStart = ta.selectionEnd = 5;
    btn('Table').click();
    expect(bar.hidden).toBe(false);
    const [rows, cols] = Array.from(bar.querySelectorAll('input'));
    rows!.value = '2';
    cols!.value = '4';
    Array.from(bar.querySelectorAll('button')).find((b) => b.textContent === 'Insert')!.click();
    const out = result();
    expect(out).toContain('| Column 1 | Column 2 | Column 3 | Column 4 |');
    expect(out.split('\n').filter((l) => l.startsWith('|'))).toHaveLength(4); // header, delimiter, 2 rows
    view.dispose();
    localStorage.removeItem('vsofx.mode.markdown');
  });

  it('visual: alignment cycles the table column; the table button opens the size bar', () => {
    const { root, btn, bar, result, view } = setup('| A | B |\n| --- | --- |\n| 1 | 2 |\n');
    const td = root.querySelector('td:nth-child(2)')!;
    const r = document.createRange();
    r.selectNodeContents(td);
    window.getSelection()!.removeAllRanges();
    window.getSelection()!.addRange(r);
    btn('Horizontal').click();
    expect(result()).toMatch(/\| --- \| :---: \|/);
    btn('Table').click();
    expect(bar.hidden).toBe(false);
    view.dispose();
    localStorage.removeItem('vsofx.mode.markdown');
  });
});
