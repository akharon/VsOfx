// @vitest-environment jsdom
// Phase 9 Wave 2: paragraph styles, shortcuts, underline / strikethrough, status text,
// remembered view state and "New Document".
import { strFromU8 } from 'fflate';
import { afterEach, describe, expect, it } from 'vitest';
import { NEW_DOCS, newDocumentBytes, storeZip } from '../src/host/newDocument.ts';
import type { Edit, SheetModel, TextModel } from '../src/shared/model.ts';
import type { ParaFmt } from '../src/shared/richtext.ts';
import { DocxAdapter } from '../src/webview/adapters/docxAdapter.ts';
import { OdsAdapter } from '../src/webview/adapters/odsAdapter.ts';
import { OdtAdapter } from '../src/webview/adapters/odtAdapter.ts';
import { XlsxAdapter } from '../src/webview/adapters/xlsxAdapter.ts';
import { readZip } from '../src/webview/adapters/zip.ts';
import { addZoom, createToolbar } from '../src/webview/ui/toolbar.ts';
import { initViewState, patchViewState, setStatusSink, viewState } from '../src/webview/viewState.ts';
import { renderTextView } from '../src/webview/views/textView.ts';
import { makeDocx, makeOdt } from './helpers/fixtures.ts';

const para = (index: number, fmt: ParaFmt, old: string, oldFmt: ParaFmt = { runs: [{ t: old }] }): Edit => ({
  kind: 'paragraph',
  index,
  text: fmt.runs.map((r) => r.t).join(''),
  old,
  fmt,
  oldFmt,
});

const ADAPTERS = [
  { name: 'docx', make: () => new DocxAdapter(), bytes: () => makeDocx('simple'), first: 'First paragraph.', main: 'word/document.xml', u: '<w:u w:val="single"/>', s: '<w:strike/>' },
  { name: 'odt', make: () => new OdtAdapter(), bytes: () => makeOdt(), first: 'Plain paragraph with bold part and a   gap.', main: 'content.xml', u: 'text-underline-style="solid"', s: 'text-line-through-style="solid"' },
] as const;

for (const f of ADAPTERS) {
  describe(`${f.name}: underline, strikethrough and styles survive save and reload`, () => {
    it('writes underline and strikethrough and shows them again', async () => {
      const a = f.make();
      const m = (await a.load(f.bytes(), [])) as TextModel;
      const idx = f.name === 'docx' ? 0 : 1;
      const old = f.name === 'docx' ? f.first : 'Plain paragraph with bold part and a   gap.';
      void m;
      a.applyEdit(para(idx, { runs: [{ t: 'Under', u: true }, { t: ' and ' }, { t: 'struck', s: true }, { t: ' both', u: true, s: true }] }, old, { runs: [{ t: old }] }));
      const bytes = await a.save();
      const xml = strFromU8(readZip(bytes)[f.main]!);
      expect(xml).toContain(f.u);
      expect(xml).toContain(f.s);
      const again = f.make();
      const model = (await again.load(bytes, [])) as TextModel;
      expect(model.html).toContain('<u>Under</u>');
      expect(model.html).toMatch(/<s>struck<\/s>/);
      expect(model.html).toMatch(/<u><s>(&nbsp;| )both<\/s><\/u>|<s><u>(&nbsp;| )both<\/u><\/s>/);
    });

    it('applies a style as one edit, saves it, and "Normal" removes it again', async () => {
      const idx = f.name === 'docx' ? 0 : 1;
      const old = f.name === 'docx' ? f.first : 'Plain paragraph with bold part and a   gap.';
      const fmt = { runs: [{ t: old }] };
      const a = f.make();
      await a.load(f.bytes(), []);
      a.applyEdit(para(idx, { ...fmt, block: 'h2' }, old));
      const bytes = await a.save();
      const model = (await f.make().load(bytes, [])) as TextModel;
      expect(model.html).toMatch(new RegExp(`<h2[^>]*data-p="${idx}"|<h2 data-p="${idx}"`));
      // back to Normal
      const b = f.make();
      await b.load(bytes, []);
      b.applyEdit(para(idx, { ...fmt, block: 'p' }, old, { ...fmt, block: 'h2' }));
      const back = (await f.make().load(await b.save(), [])) as TextModel;
      expect(back.html).not.toMatch(new RegExp(`<h2[^>]*data-p="${idx}"`));
    });
  });
}

describe('paragraph style drop-down and shortcuts', () => {
  const setup = (html = '<p data-p="0">Hello world</p><p data-p="1">Second one</p>') => {
    document.body.innerHTML = '<div id="root"></div>';
    const root = document.getElementById('root')!;
    const edits: Edit[] = [];
    const tb = createToolbar();
    const view = renderTextView(root, { type: 'text', html, rich: true }, false, (e) => edits.push(e), tb);
    const p0 = root.querySelector<HTMLElement>('[data-p="0"]')!;
    const select = tb.el.querySelector<HTMLSelectElement>('select[aria-label="Paragraph style"]')!;
    const focus = (p: HTMLElement) => p.dispatchEvent(new FocusEvent('focusin', { bubbles: true }));
    return { root, edits, tb, view, p0, select, focus };
  };

  it('has Normal, Heading 1-3 and Quote; choosing one is one edit', () => {
    const { edits, p0, select, focus, view } = setup();
    expect(Array.from(select.options).map((o) => o.textContent).slice(0, 5)).toEqual(['Normal', 'Heading 1', 'Heading 2', 'Heading 3', 'Quote']);
    focus(p0);
    select.value = 'h3';
    select.dispatchEvent(new Event('change'));
    expect(edits).toHaveLength(1);
    expect(edits[0]).toMatchObject({ kind: 'paragraph', index: 0, fmt: { block: 'h3' } });
    expect(p0.dataset['block']).toBe('h3');
    view.dispose();
  });

  it('Ctrl+Alt+digit sets the style by key position, Ctrl+Shift+8 toggles a bullet list', () => {
    const { edits, p0, focus, view, select } = setup();
    focus(p0);
    const key = (init: KeyboardEventInit) => p0.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init }));
    key({ ctrlKey: true, altKey: true, code: 'Digit1', key: '¹' });
    expect(p0.dataset['block']).toBe('h1');
    expect(select.value).toBe('h1');
    key({ ctrlKey: true, altKey: true, code: 'Digit0', key: '0' });
    expect(p0.dataset['block']).toBe('p');
    key({ ctrlKey: true, shiftKey: true, code: 'Digit8', key: '*' });
    expect(p0.dataset['block']).toBe('ul');
    expect(select.value).toBe('other');
    expect(edits).toHaveLength(3);
    view.dispose();
  });

  it('a style over several paragraphs is one batch (one undo step)', () => {
    const { root, edits, p0, select, focus, view } = setup();
    focus(p0);
    const p1 = root.querySelector<HTMLElement>('[data-p="1"]')!;
    window.getSelection()!.setBaseAndExtent(p0.firstChild!, 2, p1.firstChild!, 3);
    select.value = 'quote';
    select.dispatchEvent(new Event('change'));
    expect(edits).toHaveLength(1);
    expect(edits[0]!.kind).toBe('batch');
    expect((edits[0] as { edits: Edit[] }).edits).toHaveLength(2);
    expect(p1.dataset['block']).toBe('quote');
    view.dispose();
  });

  it('shows the shortcuts in the tooltips', () => {
    const { tb, view } = setup();
    for (const t of ['Bold (Ctrl+B)', 'Italic (Ctrl+I)', 'Underline (Ctrl+U)', 'Strikethrough (Ctrl+Shift+X)', 'Link (Ctrl+K)']) {
      expect(tb.el.querySelector(`[title="${t}"]`), t).not.toBeNull();
    }
    view.dispose();
  });

  it('reports the word count for the status bar', () => {
    const seen: string[] = [];
    setStatusSink((t) => seen.push(t));
    const { view } = setup();
    expect(seen.at(-1)).toBe('4 words');
    view.dispose();
    expect(seen.at(-1)).toBe('');
    setStatusSink(undefined);
  });
});

describe('remembered view state', () => {
  afterEach(() => initViewState({ getState: () => undefined, setState: () => {} }));

  it('saves the zoom and applies it to the next toolbar', () => {
    let stored: unknown;
    initViewState({ getState: () => stored, setState: (s) => (stored = s) });
    let z = 1;
    addZoom(createToolbar(), (f) => (z = f));
    document.dispatchEvent(new KeyboardEvent('keydown', { key: '+', ctrlKey: true }));
    expect(z).toBeGreaterThan(1);
    expect(stored).toMatchObject({ zoom: z });
    const saved = z;
    z = 1;
    initViewState({ getState: () => stored, setState: () => {} });
    expect(viewState().zoom).toBe(saved);
    addZoom(createToolbar(), (f) => (z = f));
    expect(z).toBe(saved);
  });

  it('survives a broken state object', () => {
    initViewState({ getState: () => { throw new Error('no'); }, setState: () => { throw new Error('no'); } });
    expect(() => patchViewState({ sheet: 1 })).not.toThrow();
  });
});

describe('New Document', () => {
  it('writes a valid stored archive', () => {
    const zip = storeZip([['a.txt', 'hello'], ['dir/b.txt', 'wörld']]);
    const parts = readZip(zip);
    expect(Object.keys(parts)).toEqual(['a.txt', 'dir/b.txt']);
    expect(strFromU8(parts['dir/b.txt']!)).toBe('wörld');
  });

  it('offers the four types', () => {
    expect(NEW_DOCS.map((d) => d.kind)).toEqual(['docx', 'xlsx', 'odt', 'ods']);
  });

  it('a new Word document opens, takes text, saves and reloads', async () => {
    const bytes = newDocumentBytes('docx');
    expect(bytes.length).toBeLessThan(3000);
    const a = new DocxAdapter();
    const m = (await a.load(bytes, [])) as TextModel;
    expect(m.type).toBe('text');
    a.applyEdit(para(0, { runs: [{ t: 'Hello', b: true }, { t: ' new' }], block: 'h1' }, ''));
    const model = (await new DocxAdapter().load(await a.save(), [])) as TextModel;
    expect(model.html).toContain('<strong>Hello</strong> new');
    expect(model.html).toContain('<h1');
  });

  it('a new Writer document opens, takes text, saves and reloads', async () => {
    const bytes = newDocumentBytes('odt');
    const parts = readZip(bytes);
    expect(Object.keys(parts)[0]).toBe('mimetype');
    const a = new OdtAdapter();
    await a.load(bytes, []);
    a.applyEdit(para(0, { runs: [{ t: 'Hi ' }, { t: 'there', u: true }] }, ''));
    const model = (await new OdtAdapter().load(await a.save(), [])) as TextModel;
    expect(model.html).toContain('Hi <u>there</u>');
  });

  const cell = (): Edit => ({ kind: 'cell', sheet: 0, row: 0, col: 0, value: '42', old: '' });
  it.each([
    ['xlsx', () => new XlsxAdapter()],
    ['ods', () => new OdsAdapter()],
  ] as const)('a new %s spreadsheet opens, takes a cell, saves and reloads', async (kind, make) => {
    const bytes = newDocumentBytes(kind);
    const a = make();
    const m = (await a.load(bytes, [])) as SheetModel;
    expect(m.sheets).toHaveLength(1);
    expect(m.sheets[0]!.name).toBe('Sheet1');
    a.applyEdit(cell());
    const again = (await make().load(await a.save(), [])) as SheetModel;
    expect(again.sheets[0]!.rows[0]![0]!.text).toBe('42');
  });
});
