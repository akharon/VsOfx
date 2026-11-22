// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { formatFromPath, inverseEdit, type Edit, type SheetModel } from '../src/shared/model.ts';
import { CsvAdapter, detectDelimiter } from '../src/webview/adapters/csvAdapter.ts';

const enc = (s: string) => new TextEncoder().encode(s);
const dec = (b: Uint8Array) => new TextDecoder().decode(b);

describe('file extensions', () => {
  it('maps file extensions', () => {
    expect(formatFromPath('/a/b.yml')).toBeUndefined();
    expect(formatFromPath('x.tsv')).toBe('csv');
    expect(formatFromPath('README.md')).toBeUndefined();
    expect(formatFromPath('a.json')).toBeUndefined();
    expect(formatFromPath('a.txt')).toBeUndefined();
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

describe('text encodings', () => {
  it('round-trips UTF-16 CSV files', async () => {
    const s = 'a,b\nÖ,2\n';
    const bytes = new Uint8Array(2 + s.length * 2);
    bytes.set([0xff, 0xfe]);
    for (let i = 0; i < s.length; i++) bytes[2 + i * 2] = s.charCodeAt(i) & 0xff;
    const a = new CsvAdapter();
    const m = (await a.load(bytes, [])) as SheetModel;
    expect(m.sheets[0]!.rows[1]![0]!.text).toBe('Ö');
    a.applyEdit({ kind: 'cell', sheet: 0, row: 1, col: 0, value: 'Ü', old: 'Ö' });
    const out = await a.save();
    expect(new TextDecoder('utf-16le').decode(out.subarray(2))).toBe('a,b\nÜ,2\n');
  });
});
