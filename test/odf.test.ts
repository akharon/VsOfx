import { strFromU8 } from 'fflate';
import { describe, expect, it } from 'vitest';
import type { SheetModel, TextModel } from '../src/shared/model.ts';
import { OdsAdapter } from '../src/webview/adapters/odsAdapter.ts';
import { OdtAdapter } from '../src/webview/adapters/odtAdapter.ts';
import { assertOfficePackage } from '../src/webview/adapters/index.ts';
import { odfText, setOdfText } from '../src/webview/adapters/odf.ts';
import { parseXml } from '../src/webview/adapters/xml.ts';
import { isZip, readZip } from '../src/webview/adapters/zip.ts';
import { makeOds, makeOdt } from './helpers/fixtures.ts';

const content = (bytes: Uint8Array) => strFromU8(readZip(bytes)['content.xml']!);

describe('odt adapter', () => {
  it('renders headings, spans, spaces, lists and tables', async () => {
    const model = (await new OdtAdapter().load(makeOdt(), [])) as TextModel;
    expect(model.html).toContain('<h1 data-p="0">Report title</h1>');
    expect(model.html).toContain('<strong>bold part</strong>');
    expect(model.html).toContain('a&nbsp;&nbsp;&nbsp;gap.');
    expect(model.html).toContain('<ul><li><p data-p="2">List item</p></li></ul>');
    expect(model.html).toContain('<td><p data-p="3">Cell text</p></td>');
  });

  it('round-trips an edit and puts mimetype first, stored', async () => {
    const a = new OdtAdapter();
    await a.load(makeOdt(), []);
    a.applyEdit({ kind: 'paragraph', index: 1, text: 'Edited  with   spaces & <tags>', old: '' });
    const saved = await a.save();
    expect(isZip(saved)).toBe(true);
    expect(Object.keys(readZip(saved))[0]).toBe('mimetype');
    // local header of the first entry: compression method (offset 8) must be 0 = stored
    expect(saved[8]).toBe(0);
    expect(saved[9]).toBe(0);
    expect(strFromU8(readZip(saved)['mimetype']!)).toBe('application/vnd.oasis.opendocument.text');
    expect(content(saved)).toContain('<text:s text:c="2"/>');

    const model = (await new OdtAdapter().load(saved, [])) as TextModel;
    expect(model.html).toContain('<p data-p="1">Edited &nbsp;with &nbsp;&nbsp;spaces &amp; &lt;tags&gt;</p>');
  });

  it('undo to the original text restores spans', async () => {
    const a = new OdtAdapter();
    await a.load(makeOdt(), []);
    const original = 'Plain paragraph with bold part and a   gap.';
    a.applyEdit({ kind: 'paragraph', index: 1, text: 'x', old: original });
    a.applyEdit({ kind: 'paragraph', index: 1, text: original, old: 'x' });
    expect(content(await a.save())).toContain('<text:span text:style-name="T1">bold part</text:span>');
  });

  it('rejects a spreadsheet opened as text and replays edits', async () => {
    await expect(new OdtAdapter().load(makeOds(), [])).rejects.toThrow(/Unexpected OpenDocument type/);
    const model = (await new OdtAdapter().load(makeOdt(), [{ kind: 'paragraph', index: 4, text: 'Replayed', old: 'Last one' }])) as TextModel;
    expect(model.html).toContain('<p data-p="4">Replayed</p>');
  });

  it('encodes and decodes whitespace', () => {
    const doc = parseXml('<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"/>');
    const p = doc.documentElement;
    for (const s of ['a b', '  lead', 'a    b', 'tab\there', 'line1\nline2', 'trail  ']) {
      setOdfText(doc, p, s);
      expect(odfText(p), JSON.stringify(s)).toBe(s);
    }
  });
});

describe('ods adapter', () => {
  const sheet0 = (m: SheetModel) => m.sheets[0]!;

  it('expands repeated rows and cells, ignoring trailing empty repeats', async () => {
    const model = (await new OdsAdapter().load(makeOds(), [])) as SheetModel;
    const s = sheet0(model);
    expect(model.sheets.map((x) => x.name)).toEqual(['Data', 'Second']);
    expect(s.rowCount).toBe(4); // header + 3 repeated rows; the 1,048,000 empty rows are dropped
    expect(s.rows[0]!.map((c) => c?.text)).toEqual(['Name', 'Value']);
    expect(s.rows[3]!.map((c) => c?.text)).toEqual(['7', '7']);
    expect(s.colCount).toBe(2);
    expect(s.truncated).toBe(false);
  });

  it('edits one cell inside a repeated range (split) and saves', async () => {
    const a = new OdsAdapter();
    await a.load(makeOds(), []);
    const r = a.applyEdit({ kind: 'cell', sheet: 0, row: 2, col: 1, value: '99', old: '7' });
    expect(r.cell?.text).toBe('99');
    const saved = await a.save();
    expect(Object.keys(readZip(saved))[0]).toBe('mimetype');
    const model = (await new OdsAdapter().load(saved, [])) as SheetModel;
    const rows = sheet0(model).rows;
    expect(rows[1]!.map((c) => c?.text)).toEqual(['7', '7']);
    expect(rows[2]!.map((c) => c?.text)).toEqual(['7', '99']);
    expect(rows[3]!.map((c) => c?.text)).toEqual(['7', '7']);
    expect(content(saved)).toContain('office:value="99"');
  });

  it('writes strings, numbers, empty values and cells beyond the data', async () => {
    const a = new OdsAdapter();
    await a.load(makeOds(), []);
    a.applyEdit({ kind: 'cell', sheet: 0, row: 0, col: 0, value: 'Title  two', old: 'Name' });
    a.applyEdit({ kind: 'cell', sheet: 0, row: 1, col: 0, value: '', old: '7' });
    a.applyEdit({ kind: 'cell', sheet: 0, row: 6, col: 4, value: '3.5', old: '' });
    a.applyEdit({ kind: 'cell', sheet: 1, row: 0, col: 0, value: 'y', old: 'x' });
    const model = (await new OdsAdapter().load(await a.save(), [])) as SheetModel;
    const rows = sheet0(model).rows;
    expect(rows[0]![0]!.text).toBe('Title  two');
    expect(rows[1]![0] ?? null).toBeNull();
    expect(rows[6]![4]!.text).toBe('3.5');
    expect(sheet0(model).rowCount).toBe(7);
    expect(model.sheets[1]!.rows[0]![0]!.text).toBe('y');
  });

  it('undo restores the original cell', async () => {
    const a = new OdsAdapter();
    await a.load(makeOds(), []);
    a.applyEdit({ kind: 'cell', sheet: 0, row: 2, col: 1, value: 'z', old: '7' });
    const r = a.applyEdit({ kind: 'cell', sheet: 0, row: 2, col: 1, value: '7', old: 'z' });
    expect(r.cell?.text).toBe('7');
    expect(content(await a.save())).toContain('office:value="7"');
  });

  it('replays edits on load and rejects invalid ones', async () => {
    const model = (await new OdsAdapter().load(makeOds(), [{ kind: 'cell', sheet: 0, row: 0, col: 1, value: 'Amount', old: 'Value' }])) as SheetModel;
    expect(sheet0(model).rows[0]![1]!.text).toBe('Amount');
    const a = new OdsAdapter();
    await a.load(makeOds(), []);
    expect(() => a.applyEdit({ kind: 'cell', sheet: 9, row: 0, col: 0, value: '', old: '' })).toThrow();
  });
});

describe('package checks', () => {
  it('detects non-zip input', () => {
    expect(() => assertOfficePackage(new TextEncoder().encode('just text'))).toThrow(/not a valid Office document/);
    expect(() => assertOfficePackage(new Uint8Array())).toThrow(/empty/);
    expect(() => assertOfficePackage(makeOdt())).not.toThrow();
  });
});
