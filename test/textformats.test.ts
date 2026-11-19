// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { applyTextEdit, diffText, formatFromPath, inverseEdit, isEdit, type CodeModel, type Edit, type SheetModel, type TextEdit } from '../src/shared/model.ts';
import { CodeAdapter } from '../src/webview/adapters/codeAdapter.ts';
import { CsvAdapter, detectDelimiter } from '../src/webview/adapters/csvAdapter.ts';
import { renderMarkdown, setHeading, toggleWrap, togglePrefix } from '../src/webview/code/markdown.ts';
import { createToolbar } from '../src/webview/ui/toolbar.ts';
import { renderCodeView } from '../src/webview/views/codeView.ts';

const enc = (s: string) => new TextEncoder().encode(s);
const dec = (b: Uint8Array) => new TextDecoder().decode(b);

describe('text edits', () => {
  it('diff, apply and inverse round-trip', () => {
    const a = 'hello world';
    const b = 'hello brave new world';
    const e = diffText(a, b)!;
    expect(e).toEqual({ kind: 'text', at: 6, del: '', ins: 'brave new ' });
    expect(applyTextEdit(a, e)).toBe(b);
    expect(applyTextEdit(b, inverseEdit(e) as TextEdit)).toBe(a);
    expect(isEdit(e)).toBe(true);
    expect(isEdit({ kind: 'text', at: -1, del: '', ins: 'x' })).toBe(false);
    expect(diffText(a, a)).toBeNull();
  });

  it('refuses an edit that does not match the text', () => {
    expect(() => applyTextEdit('abc', { kind: 'text', at: 1, del: 'x', ins: 'y' })).toThrow(/out of sync/);
  });

  it('maps file extensions', () => {
    expect(formatFromPath('/a/b.yml')).toBeUndefined();
    expect(formatFromPath('x.tsv')).toBe('csv');
    expect(formatFromPath('README.md')).toBe('markdown');
    expect(formatFromPath('a.json')).toBeUndefined();
    expect(formatFromPath('a.txt')).toBeUndefined();
  });
});

describe('CodeAdapter', () => {
  it('saves untouched files byte for byte', async () => {
    const bytes = enc('﻿{"a": 1}\r\n');
    const a = new CodeAdapter('markdown');
    await a.load(bytes, []);
    expect(await a.save()).toBe(bytes);
  });

  it('keeps BOM and CRLF line ends, replays edits', async () => {
    const a = new CodeAdapter('markdown');
    const edit: Edit = { kind: 'text', at: 3, del: '1', ins: '2' };
    const model = (await a.load(enc('﻿a: 1\r\nb: x\r\n'), [edit])) as CodeModel;
    expect(model.text).toBe('a: 2\nb: x\n');
    a.applyEdit({ kind: 'text', at: 10, del: '', ins: 'c: 3\n' });
    const out = await a.save();
    expect(Array.from(out.subarray(0, 3))).toEqual([0xef, 0xbb, 0xbf]);
    expect(dec(out.subarray(3))).toBe('a: 2\r\nb: x\r\nc: 3\r\n');
  });

  it('round-trips UTF-16 files', async () => {
    const s = 'a: Ö\n';
    const bytes = new Uint8Array(2 + s.length * 2);
    bytes.set([0xff, 0xfe]);
    for (let i = 0; i < s.length; i++) bytes[2 + i * 2] = s.charCodeAt(i) & 0xff;
    const a = new CodeAdapter('markdown');
    expect(((await a.load(bytes, [])) as CodeModel).text).toBe(s);
    a.applyEdit({ kind: 'text', at: 3, del: 'Ö', ins: 'Ü' });
    const out = await a.save();
    expect(new TextDecoder('utf-16le').decode(out.subarray(2))).toBe('a: Ü\n');
  });
});

describe('CsvAdapter', () => {
  it('detects the delimiter', () => {
    expect(detectDelimiter('a,b,c\n1,2,3\n')).toBe(',');
    expect(detectDelimiter('Name;Preis\n"Müller, A";1,50\nB;2,00\n')).toBe(';');
    expect(detectDelimiter('a\tb\n1\t2\n')).toBe('\t');
    expect(detectDelimiter('a|b|c\n1|2|3\n')).toBe('|');
  });

  it('shows a table and saves edits with the same delimiter and quoting rules', async () => {
    const a = new CsvAdapter();
    const src = 'Name;Preis\r\n"Müller; A";1,50\r\nB;2,00\r\n';
    const m = (await a.load(enc(src), [])) as SheetModel;
    const s = m.sheets[0]!;
    expect(s.name).toContain('semicolon');
    expect(s.rows[1]![0]!.text).toBe('Müller; A');
    expect(s.rows[0]![0]!.style?.bold).toBe(true);
    expect(await a.save()).toEqual(enc(src)); // untouched
    a.applyEdit({ kind: 'cell', sheet: 0, row: 2, col: 1, value: '3,25', old: '2,00' });
    a.applyEdit({ kind: 'cell', sheet: 0, row: 3, col: 0, value: 'New "x"', old: '' });
    expect(dec(await a.save())).toBe('Name;Preis\r\n"Müller; A";1,50\r\nB;3,25\r\n"New ""x"""\r\n');
  });

  it('undoing an appended row removes it again', async () => {
    const a = new CsvAdapter();
    await a.load(enc('a,b\n1,2\n'), []);
    const e: Edit = { kind: 'cell', sheet: 0, row: 5, col: 3, value: 'x', old: '' };
    a.applyEdit(e);
    a.applyEdit(inverseEdit(e));
    expect(dec(await a.save())).toBe('a,b\n1,2\n');
  });
});

describe('Markdown', () => {
  it('renders safely: no raw HTML, no script links, task lists', () => {
    const el = document.createElement('div');
    renderMarkdown('# Title\n\n<script>alert(1)</script>\n\n[x](javascript:alert(1))\n\n- [x] done\n- [ ] todo\n\n![pic](https://example.com/a.png)', el);
    expect(el.querySelector('script')).toBeNull();
    expect(el.textContent).toContain('<script>alert(1)</script>'); // shown as text
    expect(el.querySelector('a[href^="javascript"]')).toBeNull();
    expect(el.querySelectorAll('input.task').length).toBe(2);
    expect(el.querySelector('img')).toBeNull(); // remote image -> alt text
    expect(el.querySelector('h1')!.dataset['line']).toBe('0');
  });

  it('formatting helpers', () => {
    expect(toggleWrap({ text: 'a word b', start: 2, end: 6 }, '**').text).toBe('a **word** b');
    expect(toggleWrap({ text: 'a **word** b', start: 4, end: 8 }, '**').text).toBe('a word b');
    expect(setHeading({ text: 'Title', start: 0, end: 0 }, 2).text).toBe('## Title');
    expect(togglePrefix({ text: 'a\nb', start: 0, end: 3 }, '- ', /^\s*[-*+]\s+/).text).toBe('- a\n- b');
    expect(togglePrefix({ text: '- a\n- b', start: 0, end: 7 }, '- ', /^\s*[-*+]\s+/).text).toBe('a\nb');
  });
});

describe('code view', () => {
  it('a formatting button emits a minimal text edit; host undo reverts it', () => {
    document.body.innerHTML = '<div id="root"></div>';
    const root = document.getElementById('root')!;
    const edits: Edit[] = [];
    const tb = createToolbar();
    const view = renderCodeView(root, { type: 'code', lang: 'markdown', text: 'a word b' }, false, (e) => edits.push(e), tb);
    tb.el.querySelector<HTMLButtonElement>('[aria-label="Source"]')!.click();
    const ta = root.querySelector<HTMLTextAreaElement>('.src-text')!;
    ta.setSelectionRange(2, 6);
    tb.el.querySelector<HTMLButtonElement>('[title^="Bold"]')!.click();
    expect(ta.value).toBe('a **word** b');
    expect(edits).toEqual([{ kind: 'text', at: 2, del: 'word', ins: '**word**' }]);
    view.applyEdit(edits[0]!, {}); // own edit echoed: ignored
    view.applyEdit(inverseEdit(edits[0]!), {});
    expect(ta.value).toBe('a word b');
    view.dispose();
  });

  it('typing in the source is committed on flush', () => {
    document.body.innerHTML = '<div id="root"></div>';
    const root = document.getElementById('root')!;
    const edits: Edit[] = [];
    const tb = createToolbar();
    const view = renderCodeView(root, { type: 'code', lang: 'markdown', text: 'Hello' }, false, (e) => edits.push(e), tb);
    tb.el.querySelector<HTMLButtonElement>('[aria-label="Source"]')!.click();
    const ta = root.querySelector<HTMLTextAreaElement>('.src-text')!;
    ta.value = 'Hello world';
    ta.dispatchEvent(new Event('input'));
    expect(edits).toEqual([]);
    view.flush!();
    expect(edits).toEqual([{ kind: 'text', at: 5, del: '', ins: ' world' }]);
    view.dispose();
  });
});

describe('Markdown WYSIWYG', () => {
  const setup = (text: string) => {
    document.body.innerHTML = '<div id="root"></div>';
    try {
      localStorage.removeItem('vsofx.mode.markdown');
    } catch {
      /* ignore */
    }
    const root = document.getElementById('root')!;
    const edits: Edit[] = [];
    const view = renderCodeView(root, { type: 'code', lang: 'markdown', text }, false, (e) => edits.push(e), createToolbar());
    return { root, edits, view, body: root.querySelector<HTMLElement>('.md-body')! };
  };

  it('opens as an editable document, not as source', () => {
    const { root, body, view } = setup('# Title\n');
    expect(body.isContentEditable || body.getAttribute('contenteditable') === 'true').toBe(true);
    expect((root.querySelector('.code-layout') as HTMLElement).dataset['mode']).toBe('visual');
    view.dispose();
  });

  it('editing one block rewrites only that block', () => {
    const src = '# Title\n\nSome  *text*  here.\n\n* one\n* two\n';
    const { body, edits, view } = setup(src);
    body.querySelector('h1')!.textContent = 'New title';
    view.flush!();
    expect(edits).toHaveLength(1);
    expect(applyTextEdit(src, edits[0] as TextEdit)).toBe('# New title\n\nSome  *text*  here.\n\n* one\n* two\n');
    view.dispose();
  });

  it('serializes formatting typed in the editor', () => {
    const src = 'Para\n';
    const { body, edits, view } = setup(src);
    body.querySelector('p')!.innerHTML = 'A <b>bold</b> and <i>italic</i> <code>x</code> <a href="https://e.com">link</a>';
    view.flush!();
    expect(applyTextEdit(src, edits[0] as TextEdit)).toBe('A **bold** and _italic_ `x` [link](https://e.com)\n');
    view.dispose();
  });

  it('new paragraphs and lists are added with blank lines', () => {
    const src = 'First\n';
    const { body, edits, view } = setup(src);
    const ul = document.createElement('ul');
    ul.innerHTML = '<li>a</li><li>b</li>';
    body.append(ul);
    view.flush!();
    expect(applyTextEdit(src, edits[0] as TextEdit)).toBe('First\n\n- a\n- b\n');
    view.dispose();
  });
});
