// @vitest-environment jsdom
import { strFromU8 } from 'fflate';
import { describe, expect, it } from 'vitest';
import { inverseEdit, isEdit, type Edit, type SheetModel, type TextModel } from '../src/shared/model.ts';
import { PARA_BASE, cellIndex, type ParaFmt } from '../src/shared/richtext.ts';
import { createAdapter } from '../src/webview/adapters/index.ts';
import { readZip } from '../src/webview/adapters/zip.ts';
import { createToolbar } from '../src/webview/ui/toolbar.ts';
import { parseClipboard, renderSheetView } from '../src/webview/views/sheetView.ts';
import { renderTextView } from '../src/webview/views/textView.ts';
import { makeDocx, makeOds, makeOdt, makeXlsx } from './helpers/fixtures.ts';


const content = (bytes: Uint8Array) => strFromU8(readZip(bytes)['content.xml']!);
const docXml = (bytes: Uint8Array) => strFromU8(readZip(bytes)['word/document.xml']!);
const fmt = (t: string, extra: Partial<ParaFmt> = {}): ParaFmt => ({ runs: [{ t }], ...extra });

describe('edit model', () => {
  it('validates and inverts new edit kinds', () => {
    const np: Edit = { kind: 'newpara', op: 'insert', index: PARA_BASE + 1, after: 0, fmt: fmt('x') };
    const sh: Edit = { kind: 'sheet', op: 'rename', index: 0, name: 'B', old: 'A' };
    const batch: Edit = { kind: 'batch', edits: [np, { kind: 'cell', sheet: 0, row: 0, col: 0, value: 'a', old: '' }] };
    for (const e of [np, sh, batch]) {
      expect(isEdit(e)).toBe(true);
      expect(inverseEdit(inverseEdit(e))).toEqual(e);
    }
    expect(inverseEdit(sh)).toEqual({ kind: 'sheet', op: 'rename', index: 0, name: 'A', old: 'B' });
    expect((inverseEdit(batch) as { edits: Edit[] }).edits[0]!.kind).toBe('cell'); // reversed order
    expect(isEdit({ ...np, index: 5 })).toBe(false);
    expect(isEdit({ kind: 'batch', edits: [batch] })).toBe(false);
  });
});

describe('Word: Enter adds a paragraph', () => {
  it('inserts the paragraph after its anchor, with formatting, and removes it on undo', async () => {
    const a = createAdapter('docx');
    await a.load(makeDocx('simple'), []);
    const ins: Edit = { kind: 'newpara', op: 'insert', index: PARA_BASE + 1, after: 0, fmt: { runs: [{ t: 'Brand new', b: true }] } };
    a.applyEdit({ kind: 'batch', edits: [ins] });
    let xml = docXml(await a.save());
    expect(xml).toMatch(/First paragraph\.<\/w:t><\/w:r><\/w:p><w:p>(<w:pPr>.*?<\/w:pPr>)?<w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">Brand new/);
    a.applyEdit(inverseEdit(ins));
    xml = docXml(await a.save());
    expect(xml).not.toContain('Brand new');
  });

  it('replays added paragraphs into the model', async () => {
    const ins: Edit = { kind: 'newpara', op: 'insert', index: PARA_BASE + 1, after: 0, fmt: fmt('Added') };
    const m = (await createAdapter('docx').load(makeDocx('simple'), [ins])) as TextModel;
    expect(m.paras).toEqual([{ index: PARA_BASE + 1, after: 0, fmt: fmt('Added') }]);
  });
});

describe('Writer (odt) formatting', () => {
  it('is rich, writes runs, headings, lists and links, and renders them again', async () => {
    const a = createAdapter('odt');
    const m = (await a.load(makeOdt(), [])) as TextModel;
    expect(m.rich).toBe(true);
    a.applyEdit({ kind: 'paragraph', index: 1, text: 'Bold link', old: '', fmt: { runs: [{ t: 'Bold', b: true }, { t: ' link', href: 'https://e.com' }] }, oldFmt: fmt('') });
    a.applyEdit({ kind: 'paragraph', index: 4, text: 'Head', old: '', fmt: { runs: [{ t: 'Head' }], block: 'h2' }, oldFmt: fmt('') });
    const saved = await a.save();
    const xml = content(saved);
    expect(xml).toContain('style:name="VsofxT_b"');
    expect(xml).toContain('xlink:href="https://e.com"');
    expect(xml).toMatch(/<text:h text:outline-level="2" text:style-name="VsofxH2">Head<\/text:h>/);
    const again = (await createAdapter('odt').load(saved, [])) as TextModel;
    expect(again.html).toContain('<strong>Bold</strong><a data-href="https://e.com">&nbsp;link</a>');
    expect(again.html).toContain('<h2 data-p="4">Head</h2>');
  });

  it('turns a paragraph into a list item and back', async () => {
    const a = createAdapter('odt');
    await a.load(makeOdt(), []);
    a.applyEdit({ kind: 'paragraph', index: 4, text: 'Last one', old: 'Last one', fmt: fmt('Last one', { block: 'ol' }), oldFmt: fmt('Last one') });
    expect(content(await a.save())).toMatch(/<text:list text:style-name="VsofxOL"><text:list-item><text:p[^>]*>Last one/);
    const html = ((await createAdapter('odt').load(await a.save(), [])) as TextModel).html;
    expect(html).toContain('<ol><li><p data-p="4">Last one</p></li></ol>');
    a.applyEdit({ kind: 'paragraph', index: 4, text: 'Last one', old: 'Last one', fmt: fmt('Last one'), oldFmt: fmt('Last one', { block: 'ol' }) });
    expect(content(await a.save())).not.toContain('VsofxOL"><text:list-item>');
  });

  it('adds paragraphs and tables after an anchor', async () => {
    const a = createAdapter('odt');
    await a.load(makeOdt(), []);
    a.applyEdit({ kind: 'newpara', op: 'insert', index: PARA_BASE + 1, after: 1, fmt: fmt('New one') });
    const spec = { id: 1, after: PARA_BASE + 1, rows: 2, cols: 2 };
    a.applyEdit({ kind: 'table', op: 'insert', table: spec });
    a.applyEdit({ kind: 'paragraph', index: cellIndex(spec, 1, 1), text: 'D', old: '', fmt: fmt('D'), oldFmt: fmt('') });
    const saved = await a.save();
    const m = (await createAdapter('odt').load(saved, [])) as TextModel;
    expect(m.html.indexOf('New one')).toBeGreaterThan(m.html.indexOf('gap.'));
    expect(content(saved)).toMatch(/New one<\/text:p><table:table table:name="VsofxTable1"/);
    expect(content(saved)).toContain('>D</text:p>');
    a.applyEdit({ kind: 'table', op: 'remove', table: spec });
    a.applyEdit({ kind: 'newpara', op: 'remove', index: PARA_BASE + 1, after: 1, fmt: fmt('New one') });
    expect(content(await a.save())).not.toMatch(/New one|VsofxTable1/);
  });
});

describe('document view: Enter', () => {
  it('splits the paragraph at the cursor into a new one (one undo step)', () => {
    document.body.innerHTML = '<div id="root"></div>';
    const root = document.getElementById('root')!;
    const edits: Edit[] = [];
    const view = renderTextView(root, { type: 'text', html: '<p data-p="0">Hello world</p>', rich: true }, false, (e) => edits.push(e), createToolbar());
    const p = root.querySelector<HTMLElement>('[data-p="0"]')!;
    p.focus();
    const r = document.createRange();
    r.setStart(p.firstChild!, 5);
    r.collapse(true);
    getSelection()!.removeAllRanges();
    getSelection()!.addRange(r);
    p.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(edits).toHaveLength(1);
    const b = edits[0]!;
    expect(b.kind).toBe('batch');
    const [first, add] = (b as { edits: Edit[] }).edits;
    expect(first).toMatchObject({ kind: 'paragraph', index: 0, text: 'Hello' });
    expect(add).toMatchObject({ kind: 'newpara', op: 'insert', after: 0, fmt: { runs: [{ t: ' world' }] } });
    expect(root.querySelectorAll('[data-p]')).toHaveLength(2);
    view.applyEdit(inverseEdit(add!), {});
    expect(root.querySelectorAll('[data-p]')).toHaveLength(1);
    view.dispose();
  });
});

describe('spreadsheets', () => {
  it('parses clipboard text from Excel', () => {
    expect(parseClipboard('a\tb\r\n1\t2\r\n')).toEqual([['a', 'b'], ['1', '2']]);
    expect(parseClipboard('"x\ty"\t"say ""hi"""\n')).toEqual([['x\ty', 'say "hi"']]);
    expect(parseClipboard('one')).toEqual([['one']]);
  });

  for (const [format, bytes] of [
    ['xlsx', () => makeXlsx('multi-sheet')],
    ['ods', () => makeOds()],
  ] as const) {
    it(`${format}: adds, renames and removes sheets; undo restores the content`, async () => {
      const a = createAdapter(format);
      const m = (await a.load(await bytes(), [])) as SheetModel;
      expect(m.sheetOps).toBe(true);
      const n = m.sheets.length;
      const firstName = m.sheets[0]!.name;
      const firstRows = JSON.stringify(m.sheets[0]!.rows);
      const add: Edit = { kind: 'sheet', op: 'add', index: n, name: 'Extra' };
      expect(a.applyEdit(add).sheet?.name).toBe('Extra');
      a.applyEdit({ kind: 'cell', sheet: n, row: 0, col: 0, value: 'hi', old: '' });
      a.applyEdit({ kind: 'sheet', op: 'rename', index: n, name: 'Renamed', old: 'Extra' });
      const rm: Edit = { kind: 'sheet', op: 'remove', index: 0, name: firstName };
      a.applyEdit(rm);
      let saved = (await createAdapter(format).load(await a.save(), [])) as SheetModel;
      expect(saved.sheets.map((s) => s.name)).not.toContain(firstName);
      expect(saved.sheets[saved.sheets.length - 1]).toMatchObject({ name: 'Renamed' });
      expect(saved.sheets[saved.sheets.length - 1]!.rows[0]![0]!.text).toBe('hi');
      const back = a.applyEdit(inverseEdit(rm));
      expect(JSON.stringify(back.sheet!.rows)).toBe(firstRows);
      saved = (await createAdapter(format).load(await a.save(), [])) as SheetModel;
      expect(saved.sheets[0]!.name).toBe(firstName);
      expect(saved.sheets).toHaveLength(n + 1);
    });
  }

  it('selects a range, copies it as TSV and pastes a block as one edit', () => {
    document.body.innerHTML = '<div id="root"></div>';
    const root = document.getElementById('root')!;
    const edits: Edit[] = [];
    const model: SheetModel = {
      type: 'sheet',
      sheetOps: true,
      sheets: [{ name: 'S', rows: [[{ text: 'a' }, { text: 'b' }], [{ text: 'c' }, { text: 'd' }]], rowCount: 2, colCount: 2, merges: [], colWidths: [], truncated: false }],
    };
    const view = renderSheetView(root, model, false, (e) => edits.push(e), createToolbar());
    const td = (r: number, c: number) => root.querySelector<HTMLElement>(`td[data-r="${r}"][data-c="${c}"]`)!;
    td(0, 0).dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
    td(1, 1).dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0, shiftKey: true }));
    expect(root.querySelectorAll('td.inrange')).toHaveLength(4);
    const data = new Map<string, string>();
    const clip = { setData: (t: string, v: string) => data.set(t, v), getData: (t: string) => data.get(t) ?? '' };
    const fire = (type: string) => {
      const ev = new Event(type, { bubbles: true, cancelable: true }) as ClipboardEvent;
      Object.defineProperty(ev, 'clipboardData', { value: clip });
      document.dispatchEvent(ev);
    };
    fire('copy');
    expect(data.get('text/plain')).toBe('a\tb\nc\td\n');
    td(2, 0).dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }));
    fire('paste');
    expect(edits).toHaveLength(1);
    expect(edits[0]!.kind).toBe('batch');
    expect((edits[0] as { edits: Edit[] }).edits.map((e) => (e as { value: string }).value)).toEqual(['a', 'b', 'c', 'd']);
    // Delete clears the pasted range in one step
    root.querySelector<HTMLElement>('.grid-scroll')!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', bubbles: true }));
    expect(edits).toHaveLength(2);
    expect((edits[1] as { edits: Edit[] }).edits).toHaveLength(4);
    view.dispose();
  });
});

describe('spreadsheets: column width and wrap', () => {
  for (const [format, bytes] of [
    ['xlsx', () => makeXlsx('simple')],
    ['ods', () => makeOds()],
  ] as const) {
    it(`${format}: saves column widths and wrapped cells; undo restores`, async () => {
      const a = createAdapter(format);
      const m = (await a.load(await bytes(), [])) as SheetModel;
      const old = m.sheets[0]!.colWidths[1] ?? null;
      const w: Edit = { kind: 'colwidth', sheet: 0, col: 1, width: 30, old };
      a.applyEdit(w);
      const wrap: Edit = { kind: 'wrap', sheet: 0, row: 0, col: 0, wrap: true, old: false };
      expect(a.applyEdit(wrap).cell?.style?.wrap).toBe(true);
      a.applyEdit({ kind: 'cell', sheet: 0, row: 0, col: 0, value: 'still wrapped', old: m.sheets[0]!.rows[0]![0]?.text ?? '' });
      let back = (await createAdapter(format).load(await a.save(), [])) as SheetModel;
      expect(back.sheets[0]!.colWidths[1]).toBeCloseTo(30, 0);
      expect(back.sheets[0]!.rows[0]![0]).toMatchObject({ text: 'still wrapped', style: { wrap: true } });
      a.applyEdit(inverseEdit(wrap));
      a.applyEdit(inverseEdit(w));
      back = (await createAdapter(format).load(await a.save(), [])) as SheetModel;
      expect(back.sheets[0]!.rows[0]![0]?.style?.wrap).toBeFalsy();
      expect(back.sheets[0]!.colWidths[1] ?? null).toBe(old);
    });
  }

  it('header clicks select a whole column / row; dragging the edge resizes', () => {
    document.body.innerHTML = '<div id="root"></div>';
    const root = document.getElementById('root')!;
    const edits: Edit[] = [];
    const model: SheetModel = {
      type: 'sheet',
      sheetOps: true,
      sheets: [{ name: 'S', rows: [[{ text: 'a' }]], rowCount: 3, colCount: 2, merges: [], colWidths: [], truncated: false }],
    };
    const view = renderSheetView(root, model, false, (e) => edits.push(e), createToolbar());
    const down = (t: Element, init: MouseEventInit = {}) => t.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0, ...init }));
    down(root.querySelector('th.colhead[data-c="1"]')!);
    const inCol = Array.from(root.querySelectorAll<HTMLElement>('td.inrange, td.sel'));
    expect(inCol.every((td) => td.dataset['c'] === '1')).toBe(true);
    expect(inCol.length).toBe(30); // all shown rows (minimum grid size)
    down(root.querySelector('th.rowhead[data-r="2"]')!);
    expect(Array.from(root.querySelectorAll<HTMLElement>('td.inrange, td.sel')).every((td) => td.dataset['r'] === '2')).toBe(true);
    down(root.querySelector('th.colhead[data-c="0"] .col-grip')!, { clientX: 100 });
    document.dispatchEvent(new MouseEvent('mousemove', { clientX: 170 }));
    document.dispatchEvent(new MouseEvent('mouseup', {}));
    expect(edits.at(-1)).toMatchObject({ kind: 'colwidth', sheet: 0, col: 0, old: null });
    expect((edits.at(-1) as { width: number }).width).toBe(21.4); // (90 + 70 - 10) / 7
    view.dispose();
  });
});
