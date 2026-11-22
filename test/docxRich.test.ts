// @vitest-environment jsdom
import { strFromU8 } from 'fflate';
import { describe, expect, it } from 'vitest';
import { inverseEdit, isEdit, type Edit, type TextModel } from '../src/shared/model.ts';
import { cellIndex, runsToHtml, type ParaFmt, type TableSpec } from '../src/shared/richtext.ts';
import { DocxAdapter } from '../src/webview/adapters/docxAdapter.ts';
import { readZip, writeZip } from '../src/webview/adapters/zip.ts';
import { createToolbar } from '../src/webview/ui/toolbar.ts';
import { renderTextView, serializeRuns } from '../src/webview/views/textView.ts';
import { makeDocx } from './helpers/fixtures.ts';

const plain = (t: string): ParaFmt => ({ runs: [{ t }] });
const para = (index: number, fmt: ParaFmt, oldText: string, oldFmt: ParaFmt = plain(oldText)): Edit => ({
  kind: 'paragraph',
  index,
  text: fmt.runs.map((r) => r.t).join(''),
  old: oldText,
  fmt,
  oldFmt,
});

async function saved(edits: Edit[], bytes = makeDocx('simple')) {
  const a = new DocxAdapter();
  await a.load(bytes, []);
  for (const e of edits) a.applyEdit(e);
  const parts = readZip(await a.save());
  const s = (n: string) => (parts[n] ? strFromU8(parts[n]!) : '');
  return { doc: s('word/document.xml'), styles: s('word/styles.xml'), numbering: s('word/numbering.xml'), rels: s('word/_rels/document.xml.rels'), ct: s('[Content_Types].xml'), a };
}

describe('Word formatting: edit model', () => {
  it('validates and inverts rich paragraph and table edits', () => {
    const e = para(0, { runs: [{ t: 'Hi', b: true }], block: 'h2' }, 'First paragraph.');
    expect(isEdit(e)).toBe(true);
    expect(inverseEdit(inverseEdit(e))).toEqual(e);
    expect((inverseEdit(e) as { fmt: ParaFmt }).fmt).toEqual(plain('First paragraph.'));
    expect(isEdit({ ...e, fmt: { runs: [{ t: 'x', href: 'javascript:alert(1)' }] } })).toBe(false);
    expect(isEdit({ ...e, fmt: { runs: [{ t: 'x', evil: true }] } })).toBe(false);
    const t: Edit = { kind: 'table', op: 'insert', table: { id: 1, after: 0, rows: 2, cols: 2 } };
    expect(isEdit(t)).toBe(true);
    expect(inverseEdit(t)).toEqual({ ...t, op: 'remove' });
    expect(isEdit({ kind: 'table', op: 'insert', table: { id: 1, after: 0, rows: 100, cols: 20 } })).toBe(false);
  });
});

describe('Word formatting: saving', () => {
  it('writes bold, italic, code and links as runs with schema-ordered properties', async () => {
    const { doc, rels } = await saved([
      para(0, { runs: [{ t: 'Bold', b: true }, { t: ' and ' }, { t: 'it', i: true }, { t: ' ' }, { t: 'x()', code: true }, { t: ' ' }, { t: 'site', href: 'https://example.com/a?b=1&c=2' }] }, 'First paragraph.'),
    ]);
    expect(doc).toMatch(/<w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">Bold<\/w:t><\/w:r>/);
    expect(doc).toMatch(/<w:rPr><w:i\/><\/w:rPr><w:t xml:space="preserve">it</);
    expect(doc).toMatch(/<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"\/><w:shd [^>]*\/><\/w:rPr><w:t xml:space="preserve">x\(\)/);
    const id = /<w:hyperlink r:id="(rIdVsofx\d+)"/.exec(doc)?.[1];
    expect(id).toBeTruthy();
    expect(doc).toMatch(/<w:color w:val="0563C1"\/><w:u w:val="single"\/>/); // color before u
    expect(rels).toContain(`Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/a?b=1&amp;c=2" TargetMode="External"`);
    expect(doc).not.toContain('First paragraph.');
  });

  it('turns paragraphs into headings, lists and quotes and adds missing styles and numbering', async () => {
    const { doc, styles, numbering } = await saved([
      para(0, { runs: [{ t: 'First paragraph.' }], block: 'h1' }, 'First paragraph.'),
      para(1, { runs: [{ t: 'Item' }], block: 'ul' }, 'Second paragraph with <angle> & ampersand.'),
      para(3, { runs: [{ t: 'Said' }], block: 'quote' }, 'Last paragraph.'),
    ]);
    expect(doc).toMatch(/<w:p><w:pPr><w:pStyle w:val="Heading1"\/><\/w:pPr>/);
    expect(doc).toMatch(/<w:pStyle w:val="ListParagraph"\/><w:numPr><w:ilvl w:val="0"\/><w:numId w:val="(\d+)"\/><\/w:numPr>/);
    expect(doc).toMatch(/<w:pStyle w:val="Quote"\/>/);
    expect(styles).toContain('w:styleId="Heading1"');
    expect(styles).toContain('w:styleId="Quote"');
    expect(styles).toContain('w:styleId="ListParagraph"');
    const numId = /<w:numId w:val="(\d+)"/.exec(doc)![1];
    expect(numbering).toContain(`<w:num w:numId="${numId}">`);
    // abstractNum definitions come before num definitions (schema order)
    expect(numbering.lastIndexOf('<w:abstractNum ')).toBeLessThan(numbering.indexOf('<w:num '));
  });

  it('creates numbering.xml with relationship and content type when the file has none', async () => {
    const parts = readZip(makeDocx('simple'));
    delete parts['word/numbering.xml'];
    const ct = strFromU8(parts['[Content_Types].xml']!).replace(/<Override PartName="\/word\/numbering.xml"[^>]*\/>/, '');
    const bytes = writeZip(parts, { '[Content_Types].xml': ct });
    const out = await saved([para(0, { runs: [{ t: 'One' }], block: 'ol' }, 'First paragraph.')], bytes);
    expect(out.numbering).toContain('<w:numFmt w:val="decimal"/>');
    expect(out.rels).toContain('relationships/numbering" Target="numbering.xml"');
    expect(out.ct).toContain('PartName="/word/numbering.xml"');
  });

  it('keeps paragraph properties when only the text formatting changed', async () => {
    const bytes = makeDocx('styles');
    const { doc } = await saved([para(0, { runs: [{ t: 'Main ' }, { t: 'Title', i: true }] }, 'Main Title')], bytes);
    expect(doc).toMatch(/<w:pStyle w:val="Heading1"\/>.*Main .*<w:i\/>.*Title/);
  });

  it('formatting back to the original keeps the original XML', async () => {
    const bytes = makeDocx('simple');
    const e = para(0, { runs: [{ t: 'First paragraph.', b: true }] }, 'First paragraph.');
    const { doc } = await saved([e, inverseEdit(e)], bytes);
    expect(doc).toBe(strFromU8(readZip(bytes)['word/document.xml']!));
  });

  it('inserts a table with formatted cells after a paragraph; undo removes it', async () => {
    const spec: TableSpec = { id: 1, after: 0, rows: 2, cols: 2 };
    const cell: Edit = para(cellIndex(spec, 1, 1), { runs: [{ t: 'B2', b: true }] }, '', { runs: [] });
    const out = await saved([{ kind: 'table', op: 'insert', table: spec }, cell]);
    expect(out.doc).toMatch(/First paragraph\.<\/w:t><\/w:r><\/w:p><w:tbl><w:tblPr>/);
    expect((out.doc.match(/<w:tc>/g) ?? []).length).toBe(4);
    expect(out.doc).toContain('<w:tblBorders>');
    expect(out.doc).toMatch(/<w:b\/><\/w:rPr><w:t xml:space="preserve">B2</);
    const undone = await saved([{ kind: 'table', op: 'insert', table: spec }, cell, { kind: 'table', op: 'remove', table: spec }]);
    expect(undone.doc).not.toContain('<w:tbl>');
  });

  it('replays formatting and tables into the model (hot exit)', async () => {
    const spec: TableSpec = { id: 1, after: 1, rows: 1, cols: 2 };
    const a = new DocxAdapter();
    const model = (await a.load(makeDocx('simple'), [
      para(0, { runs: [{ t: 'Big', b: true }], block: 'h2' }, 'First paragraph.'),
      { kind: 'table', op: 'insert', table: spec },
      para(cellIndex(spec, 0, 1), { runs: [{ t: 'cell' }] }, '', { runs: [] }),
    ])) as TextModel;
    expect(model.rich).toBe(true);
    expect(model.html).toContain('<p data-p="0" data-block="h2"><strong>Big</strong></p>');
    expect(model.tables).toEqual([{ table: spec, cells: ['', 'cell'] }]);
  });
});

describe('Word formatting: view', () => {
  it('serialises formatted HTML to runs', () => {
    const el = document.createElement('p');
    el.innerHTML = 'a <strong>b <em>c</em></strong><code>d</code> <a data-href="https://x.example">e</a><span style="font-weight: bold">f</span>';
    expect(serializeRuns(el)).toEqual([
      { t: 'a ' },
      { t: 'b ', b: true },
      { t: 'c', b: true, i: true },
      { t: 'd', code: true },
      { t: ' ' },
      { t: 'e', href: 'https://x.example' },
      { t: 'f', b: true },
    ]);
    el.innerHTML = runsToHtml([{ t: '<x>', b: true, href: 'https://y.example' }]);
    expect(serializeRuns(el)).toEqual([{ t: '<x>', b: true, href: 'https://y.example' }]);
  });

  it('block buttons emit rich edits; host undo restores; table button inserts an editable table', () => {
    document.body.innerHTML = '<div id="root"></div>';
    const root = document.getElementById('root')!;
    const edits: Edit[] = [];
    const tb = createToolbar();
    const view = renderTextView(root, { type: 'text', html: '<p data-p="0">Hello <strong>world</strong></p>', rich: true }, false, (e) => edits.push(e), tb);
    const p = root.querySelector<HTMLElement>('[data-p="0"]')!;
    p.focus();
    p.dispatchEvent(new FocusEvent('focusin', { bubbles: true }));
    const click = (title: string) => tb.el.querySelector<HTMLButtonElement>(`[title^="${title}"]`)!.click();
    const style = tb.el.querySelector<HTMLSelectElement>('select[aria-label="Paragraph style"]')!;
    style.value = 'h1';
    style.dispatchEvent(new Event('change'));
    expect(edits[0]).toMatchObject({ kind: 'paragraph', index: 0, fmt: { block: 'h1', runs: [{ t: 'Hello ' }, { t: 'world', b: true }] }, oldFmt: { runs: [{ t: 'Hello ' }, { t: 'world', b: true }] } });
    expect(p.dataset['block']).toBe('h1');
    view.applyEdit(inverseEdit(edits[0]!), {});
    expect(p.dataset['block']).toBeUndefined();
    click('Task');
    expect(p.textContent).toBe('☐ Hello world');
    click('Table');
    const bar = tb.extra.querySelector('.tablebar')!;
    const [rows, cols] = Array.from(bar.querySelectorAll<HTMLInputElement>('input[type="number"]'));
    rows!.value = '2';
    cols!.value = '4';
    Array.from(bar.querySelectorAll('button')).find((b) => b.textContent === 'Insert')!.click();
    expect(edits.at(-1)).toEqual({ kind: 'table', op: 'insert', table: { id: 1, after: 0, rows: 2, cols: 4 } });
    expect(root.querySelectorAll('table.ins-table [data-p]').length).toBe(8);
    expect(root.querySelector('article.page')!.getAttribute('contenteditable')).toBe('true');
    view.applyEdit({ kind: 'table', op: 'remove', table: { id: 1, after: 0, rows: 2, cols: 4 } }, {});
    expect(root.querySelector('table.ins-table')).toBeNull();
    view.dispose();
  });

  it('documents without rich support (ODT) show no formatting buttons', () => {
    document.body.innerHTML = '<div id="root"></div>';
    const tb = createToolbar();
    renderTextView(document.getElementById('root')!, { type: 'text', html: '<p data-p="0">x</p>' }, false, () => {}, tb).dispose();
    expect(tb.el.querySelector('[title^="Bold"]')).toBeNull();
  });
});
