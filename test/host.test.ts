import { describe, expect, it } from 'vitest';
import { OfficeDocument } from '../src/host/OfficeDocument.ts';
import { buildCsp, buildWebviewHtml, makeNonce } from '../src/host/webviewHtml.ts';
import { formatFromPath, inverseEdit, isEdit, type Edit } from '../src/shared/model.ts';

const uri = { toString: () => 'file:///a.xlsx' } as never;
const edit = (value: string, old: string): Edit => ({ kind: 'cell', sheet: 0, row: 1, col: 2, value, old });

describe('OfficeDocument', () => {
  it('tracks edits with undo and redo', () => {
    const doc = new OfficeDocument(uri, new Uint8Array([1, 2, 3]));
    const e1 = edit('b', 'a');
    const e2 = edit('c', 'b');
    doc.addEdit(e1);
    doc.addEdit(e2);
    expect(doc.edits).toEqual([e1, e2]);

    const inverse = doc.undo();
    expect(inverse).toEqual(edit('b', 'c'));
    expect(doc.edits).toEqual([e1]);
    doc.redo(e2);
    expect(doc.edits).toEqual([e1, e2]);
  });

  it('undo on an empty list throws', () => {
    expect(() => new OfficeDocument(uri, new Uint8Array()).undo()).toThrow();
  });

  it('remembers the saved state and resets on revert', () => {
    const doc = new OfficeDocument(uri, new Uint8Array([1]));
    doc.addEdit(edit('b', 'a'));
    doc.markSaved();
    expect(doc.savedEditCount).toBe(1);
    doc.reset(new Uint8Array([9]));
    expect(doc.edits).toEqual([]);
    expect(Array.from(doc.originalBytes)).toEqual([9]);
    expect(doc.savedEditCount).toBe(0);
  });

  it('round-trips through a hot-exit backup', () => {
    const doc = new OfficeDocument(uri, new Uint8Array([5, 6, 7, 8]));
    const p: Edit = { kind: 'paragraph', index: 3, text: 'new', old: 'old' };
    doc.addEdit(edit('b', 'a'));
    doc.addEdit(p);
    doc.markSaved();
    const back = OfficeDocument.deserialize(uri, doc.serialize());
    expect(back.edits).toEqual(doc.edits);
    expect(Array.from(back.originalBytes)).toEqual([5, 6, 7, 8]);
    expect(back.savedEditCount).toBe(2);
  });
});

describe('shared model', () => {
  it('inverts edits', () => {
    const p: Edit = { kind: 'paragraph', index: 0, text: 'b', old: 'a' };
    expect(inverseEdit(p)).toEqual({ kind: 'paragraph', index: 0, text: 'a', old: 'b' });
    expect(inverseEdit(inverseEdit(p))).toEqual(p);
  });

  it('validates edits coming from the webview', () => {
    expect(isEdit(edit('a', 'b'))).toBe(true);
    expect(isEdit({ kind: 'paragraph', index: 1, text: 'x', old: 'y' })).toBe(true);
    expect(isEdit({ kind: 'cell', sheet: -1, row: 0, col: 0, value: '', old: '' })).toBe(false);
    expect(isEdit({ kind: 'cell', sheet: 0, row: 0.5, col: 0, value: '', old: '' })).toBe(false);
    expect(isEdit({ kind: 'paragraph', index: 1, text: 5, old: '' })).toBe(false);
    expect(isEdit({ kind: 'other' })).toBe(false);
    expect(isEdit(null)).toBe(false);
  });

  it('detects the format from the file name', () => {
    expect(formatFromPath('/x/Report.DOCX')).toBe('docx');
    expect(formatFromPath('a.ods')).toBe('ods');
    expect(formatFromPath('a.txt')).toBeUndefined();
  });
});

describe('webview HTML (security)', () => {
  const csp = buildCsp('vscode-resource:', 'abc');

  it('uses exactly the policy of doc/arch42.md §8.1', () => {
    expect(csp).toBe(
      "default-src 'none'; script-src 'nonce-abc'; style-src vscode-resource: 'unsafe-inline'; img-src data: blob:; " +
        "font-src vscode-resource:; worker-src blob: vscode-resource:; connect-src 'none'; form-action 'none'; frame-src 'none'",
    );
  });

  it('blocks network, eval and remote content', () => {
    expect(csp).toContain("connect-src 'none'");
    expect(csp).not.toContain('unsafe-eval');
    expect(csp).not.toMatch(/https?:/);
    expect(csp).not.toContain('*');
    expect(csp).toContain("default-src 'none'");
  });

  it('embeds the policy and loads only nonce-tagged scripts', () => {
    const html = buildWebviewHtml({ cspSource: 'vscode-resource:', scriptUri: 'vscode-resource:/webview.js', styleUri: 'vscode-resource:/webview.css', nonce: 'abc' });
    expect(html).toContain(`content="${csp}"`);
    const scripts = html.match(/<script[^>]*>/g) ?? [];
    expect(scripts.length).toBe(1);
    expect(scripts[0]).toContain('nonce="abc"');
    expect(html).not.toMatch(/<script>[^<]/);
    expect(html).not.toMatch(/\son\w+=/);
  });

  it('passes the PDF chunk URL as data, without a second script tag or a looser policy', () => {
    const base = { cspSource: 'vscode-resource:', scriptUri: 'vscode-resource:/webview.js', styleUri: 'vscode-resource:/webview.css', nonce: 'abc' };
    const html = buildWebviewHtml({ ...base, pdfChunkUri: 'vscode-resource:/pdf.js' });
    expect(html).toContain('<div id="app" aria-live="polite" data-pdf-chunk="vscode-resource:/pdf.js" data-skeleton="page">');
    expect(html.match(/<script[^>]*>/g)).toHaveLength(1);
    expect(html).toContain(`content="${buildCsp('vscode-resource:', 'abc')}"`);
    expect(buildCsp('vscode-resource:', 'abc')).toContain("script-src 'nonce-abc';");
    expect(buildCsp('vscode-resource:', 'abc')).toContain("connect-src 'none'");
    expect(buildWebviewHtml(base)).not.toContain('data-pdf-chunk');
  });

  it('creates unpredictable nonces', () => {
    const a = makeNonce();
    expect(a).toMatch(/^[0-9a-f]{32}$/);
    expect(makeNonce()).not.toBe(a);
  });
});

describe('base64 transport', () => {
  it('round-trips bytes of every length remainder, and matches Node', async () => {
    const { fromByteArray, toByteArray } = await import('../src/shared/base64.ts');
    for (const len of [0, 1, 2, 3, 4, 5, 255, 1000]) {
      const bytes = Uint8Array.from({ length: len }, (_, i) => (i * 37 + 11) & 255);
      const b64 = fromByteArray(bytes);
      expect(b64).toBe(Buffer.from(bytes).toString('base64'));
      expect(Array.from(toByteArray(b64))).toEqual(Array.from(bytes));
    }
    expect(() => toByteArray('ab$d')).toThrow(/Invalid base64/);
  });
});
