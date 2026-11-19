// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { applyTextEdit, diffText, formatFromPath, inverseEdit, isEdit, type CodeModel, type Edit, type SheetModel, type TextEdit } from '../src/shared/model.ts';
import { CodeAdapter } from '../src/webview/adapters/codeAdapter.ts';
import { CsvAdapter, detectDelimiter } from '../src/webview/adapters/csvAdapter.ts';
import { formatJson, jsonTree } from '../src/webview/code/jsonTree.ts';
import { renderMarkdown, setHeading, toggleWrap, togglePrefix } from '../src/webview/code/markdown.ts';
import { spliceSpan, type TreeNode } from '../src/webview/code/tree.ts';
import { xmlTree } from '../src/webview/code/xmlTree.ts';
import { yamlTree } from '../src/webview/code/yamlTree.ts';
import { createToolbar } from '../src/webview/ui/toolbar.ts';
import { renderCodeView } from '../src/webview/views/codeView.ts';

const enc = (s: string) => new TextEncoder().encode(s);
const dec = (b: Uint8Array) => new TextDecoder().decode(b);
/** Finds a node by its label path, e.g. find(roots, 'a', 'b'). */
function find(nodes: TreeNode[], ...labels: string[]): TreeNode {
  let cur: TreeNode | undefined = { id: '', kind: 'object', label: '', at: 0, children: nodes };
  for (const l of labels) cur = cur?.children?.find((c) => c.label === l);
  if (!cur) throw new Error(`no node ${labels.join('/')}`);
  return cur;
}

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
    expect(formatFromPath('/a/b.YML')).toBe('yaml');
    expect(formatFromPath('x.tsv')).toBe('csv');
    expect(formatFromPath('README.md')).toBe('markdown');
    expect(formatFromPath('a.svg')).toBe('xml');
    expect(formatFromPath('a.txt')).toBeUndefined();
  });
});

describe('CodeAdapter', () => {
  it('saves untouched files byte for byte', async () => {
    const bytes = enc('﻿{"a": 1}\r\n');
    const a = new CodeAdapter('json');
    await a.load(bytes, []);
    expect(await a.save()).toBe(bytes);
  });

  it('keeps BOM and CRLF line ends, replays edits', async () => {
    const a = new CodeAdapter('yaml');
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
    const a = new CodeAdapter('yaml');
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

describe('JSON tree', () => {
  const src = '{\n  // settings\n  "name": "VsOfx",\n  "count": 3,\n  "tags": ["a", "b"],\n  "deep": { "on": true }\n}\n';

  it('edits values and keys in place, keeping comments', () => {
    const t = jsonTree(src);
    expect(t.problems).toEqual([]);
    const name = find(t.roots, '(root)', 'name');
    let out = spliceSpan(src, name.valueSpan!, 'Say "hi"');
    expect(out).toContain('"name": "Say \\"hi\\""');
    expect(out).toContain('// settings');
    out = spliceSpan(src, find(t.roots, '(root)', 'count').valueSpan!, '42');
    expect(JSON.parse(out.replace('// settings', '')).count).toBe(42);
    out = spliceSpan(src, find(t.roots, '(root)', 'count').valueSpan!, 'not a number');
    expect(out).toContain('"count": "not a number"');
    out = spliceSpan(src, find(t.roots, '(root)', 'deep').labelSpan!, 'deeper');
    expect(out).toContain('"deeper": { "on": true }');
  });

  it('removes and adds entries', () => {
    const t = jsonTree(src);
    const removed = find(t.roots, '(root)', 'tags').remove!()!;
    expect(removed).not.toContain('tags');
    expect(jsonTree(removed).problems).toEqual([]);
    const added = find(t.roots, '(root)', 'tags').add!()!;
    expect(JSON.parse(added.replace('// settings', '')).tags).toEqual(['a', 'b', '']);
    const withKey = find(t.roots, '(root)').add!()!;
    expect(JSON.parse(withKey.replace('// settings', '')).key).toBe('');
  });

  it('reports syntax errors and formats', () => {
    expect(jsonTree('{"a": }').problems.length).toBeGreaterThan(0);
    expect(formatJson('{"a":[1,2]}')).toBe('{\n  "a": [\n    1,\n    2\n  ]\n}');
  });
});

describe('YAML tree', () => {
  const src = '# config\nname: VsOfx # inline\nport: 8080\nlist:\n  - one\n  - two\nnested:\n  flag: true\n';

  it('edits scalars keeping comments and types', () => {
    const t = yamlTree(src);
    expect(t.problems).toEqual([]);
    let out = spliceSpan(src, find(t.roots, '(root)', 'name').valueSpan!, 'New name');
    expect(out).toContain('name: New name # inline');
    out = spliceSpan(src, find(t.roots, '(root)', 'name').valueSpan!, 'yes: really');
    expect(out).toContain('name: "yes: really" # inline');
    out = spliceSpan(src, find(t.roots, '(root)', 'port').valueSpan!, '9090');
    expect(out).toContain('port: 9090\n');
    out = spliceSpan(src, find(t.roots, '(root)', 'name').valueSpan!, '123');
    expect(out).toContain('name: "123"'); // stays a string
  });

  it('removes and adds items with the right indentation', () => {
    const t = yamlTree(src);
    expect(find(t.roots, '(root)', 'list', '0').remove!()).toBe(src.replace('  - one\n', ''));
    expect(find(t.roots, '(root)', 'port').remove!()).toBe(src.replace('port: 8080\n', ''));
    expect(find(t.roots, '(root)', 'list').add!()).toContain("  - two\n  - ''\n");
    expect(find(t.roots, '(root)', 'nested').add!()).toContain("  flag: true\n  key: ''\n");
  });

  it('reports errors', () => {
    expect(yamlTree('a: [1, 2\nb: 3\n').problems.length).toBeGreaterThan(0);
  });
});

describe('XML tree', () => {
  const src = '<?xml version="1.0"?>\n<root id="1">\n  <!-- c -->\n  <name lang="de">Müller &amp; Co</name>\n  <empty/>\n  <list>\n    <item>a</item>\n  </list>\n</root>\n';

  it('edits text and attributes with escaping', () => {
    const t = xmlTree(src);
    expect(t.problems).toEqual([]);
    const name = find(t.roots, 'root', 'name');
    expect(name.value).toBe('Müller & Co');
    let out = spliceSpan(src, name.valueSpan!, 'A < B & C');
    expect(out).toContain('<name lang="de">A &lt; B &amp; C</name>');
    out = spliceSpan(src, find(t.roots, 'root', 'name', '@lang').valueSpan!, 'say "x"');
    expect(out).toContain('<name lang="say &quot;x&quot;">');
    out = spliceSpan(src, find(t.roots, 'root', 'empty').valueSpan!, 'filled');
    expect(out).toContain('<empty>filled</empty>');
    expect(out).toContain('<!-- c -->');
  });

  it('removes elements and attributes', () => {
    const t = xmlTree(src);
    expect(find(t.roots, 'root', 'empty').remove!()).toBe(src.replace('  <empty/>\n', ''));
    expect(find(t.roots, 'root', '@id').remove!()).toContain('<root>');
    expect(find(t.roots, 'root', 'list').add!()).toContain('<list name="">');
  });

  it('reports malformed XML', () => {
    expect(xmlTree('<a><b></a>').problems.length).toBeGreaterThan(0);
    expect(xmlTree('<a>').problems[0]!.message).toMatch(/not closed/);
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
  it('a tree edit emits a minimal text edit and updates the source', () => {
    document.body.innerHTML = '<div id="root"></div>';
    const root = document.getElementById('root')!;
    const edits: Edit[] = [];
    const tb = createToolbar();
    const view = renderCodeView(root, { type: 'code', lang: 'json', text: '{\n  "a": 1\n}\n' }, false, (e) => edits.push(e), tb);
    const input = root.querySelector<HTMLInputElement>('.tn-val')!;
    expect(input.value).toBe('1');
    input.value = '2';
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }));
    expect(edits).toEqual([{ kind: 'text', at: 9, del: '1', ins: '2' }]);
    expect(root.querySelector<HTMLTextAreaElement>('.src-text')!.value).toBe('{\n  "a": 2\n}\n');
    // the host undoes it
    view.applyEdit(edits[0]!, {}); // own edit echoed: ignored
    view.applyEdit(inverseEdit(edits[0]!), {});
    expect(root.querySelector<HTMLTextAreaElement>('.src-text')!.value).toBe('{\n  "a": 1\n}\n');
    view.dispose();
  });

  it('typing in the source is committed on flush', () => {
    document.body.innerHTML = '<div id="root"></div>';
    const root = document.getElementById('root')!;
    const edits: Edit[] = [];
    const view = renderCodeView(root, { type: 'code', lang: 'markdown', text: 'Hello' }, false, (e) => edits.push(e), createToolbar());
    const ta = root.querySelector<HTMLTextAreaElement>('.src-text')!;
    ta.value = 'Hello world';
    ta.dispatchEvent(new Event('input'));
    expect(edits).toEqual([]);
    view.flush!();
    expect(edits).toEqual([{ kind: 'text', at: 5, del: '', ins: ' world' }]);
    view.dispose();
  });
});
