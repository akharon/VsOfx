import { strFromU8 } from 'fflate';
import { describe, expect, it } from 'vitest';
import type { TextModel } from '../src/shared/model.ts';
import { DocxAdapter } from '../src/webview/adapters/docxAdapter.ts';
import { readZip } from '../src/webview/adapters/zip.ts';
import { makeDocx, type DocxKind } from './helpers/fixtures.ts';

async function load(kind: DocxKind, edits: Parameters<DocxAdapter['load']>[1] = []) {
  const a = new DocxAdapter();
  const bytes = makeDocx(kind);
  const model = (await a.load(bytes, edits)) as TextModel;
  return { a, model, bytes };
}

describe('docx adapter', () => {
  it('renders paragraphs with escaped text and editable markers', async () => {
    const { model } = await load('simple');
    expect(model.html).toContain('<p data-p="0">First paragraph.</p>');
    expect(model.html).toContain('<p data-p="1">Second paragraph with &lt;angle&gt; &amp; ampersand.</p>');
    expect(model.html).toContain('<p data-p="3">Last paragraph.</p>'); // empty paragraph 2 is skipped
  });

  it('renders headings, lists, tables and images', async () => {
    const styles = (await load('styles')).model.html;
    expect(styles).toMatch(/<h1 data-p="0">Main Title<\/h1>/);
    expect(styles).toMatch(/<ul><li data-p="2">Bullet one<\/li><li data-p="3">Bullet two<\/li><\/ul>/);
    expect(styles).toMatch(/<h2 data-p="4">Subtitle<\/h2>/);

    const tables = (await load('tables')).model.html;
    expect(tables).toMatch(/<table>.*<td><p data-p="1">A1<\/p><\/td>/);

    const images = (await load('images')).model.html;
    expect(images).toMatch(/<img src="data:image\/png;base64,/);
  });

  it('maps at least 95% of text blocks to paragraphs in every fixture', async () => {
    for (const kind of ['simple', 'styles', 'tables', 'images', 'split-runs'] as const) {
      const { a } = await load(kind);
      expect(a.mappedRatio, kind).toBeGreaterThanOrEqual(0.95);
    }
  });

  it('never emits script, handlers or external URLs', async () => {
    for (const kind of ['simple', 'styles', 'tables', 'images', 'split-runs'] as const) {
      const html = (await load(kind)).model.html;
      expect(html).not.toMatch(/<script|onerror|onclick|javascript:|https?:/i);
    }
  });

  it('saves an edit: new text present, other parts byte-identical', async () => {
    const { a, bytes } = await load('simple');
    a.applyEdit({ kind: 'paragraph', index: 1, text: 'Changed <&> text', old: 'x' });
    const saved = await a.save();
    const doc = strFromU8(readZip(saved)['word/document.xml']!);
    expect(doc).toContain('Changed &lt;&amp;&gt; text');
    expect(doc).not.toContain('Second paragraph');
    expect(doc).toContain('First paragraph.');
    const before = readZip(bytes);
    const after = readZip(saved);
    expect(Object.keys(after)).toEqual(Object.keys(before));
    for (const name of Object.keys(before)) {
      if (name !== 'word/document.xml') expect(after[name], name).toEqual(before[name]);
    }
    // reload shows the change
    const reloaded = (await new DocxAdapter().load(saved, [])) as TextModel;
    expect(reloaded.html).toContain('Changed &lt;&amp;&gt; text');
  });

  it('keeps paragraph properties and the first run formatting when patching', async () => {
    const { a } = await load('split-runs');
    a.applyEdit({ kind: 'paragraph', index: 0, text: 'Brand new', old: 'Hello world' });
    const doc = strFromU8(readZip(await a.save())['word/document.xml']!);
    expect(doc).toMatch(/<w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">Brand new<\/w:t><\/w:r>/);
    expect(doc).not.toContain('<w:i/>');
    expect(doc).toContain('Plain second paragraph.');
  });

  it('keeps pPr (heading style) when patching', async () => {
    const { a } = await load('styles');
    a.applyEdit({ kind: 'paragraph', index: 0, text: 'New Title', old: 'Main Title' });
    const doc = strFromU8(readZip(await a.save())['word/document.xml']!);
    expect(doc).toContain('<w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">New Title</w:t></w:r>');
  });

  it('undo back to the original text keeps the original runs', async () => {
    const { a, bytes } = await load('split-runs');
    a.applyEdit({ kind: 'paragraph', index: 0, text: 'x', old: 'Hello world' });
    a.applyEdit({ kind: 'paragraph', index: 0, text: 'Hello world', old: 'x' });
    const saved = await a.save();
    // nothing left to patch, so the document part is untouched
    expect(readZip(saved)['word/document.xml']).toEqual(readZip(bytes)['word/document.xml']);
    expect(strFromU8(readZip(saved)['word/document.xml']!)).toContain('<w:i/>');
  });

  it('shows replayed edits and edited text in the model', async () => {
    const { model } = await load('simple', [{ kind: 'paragraph', index: 0, text: 'Replayed', old: 'First paragraph.' }]);
    expect(model.html).toContain('<p data-p="0">Replayed</p>');
  });

  it('rejects unknown paragraphs and wrong edit kinds', async () => {
    const { a } = await load('simple');
    expect(() => a.applyEdit({ kind: 'paragraph', index: 99, text: 'x', old: '' })).toThrow();
    expect(() => a.applyEdit({ kind: 'cell', sheet: 0, row: 0, col: 0, value: '', old: '' })).toThrow();
  });

  it('fails clearly on a package without document.xml', async () => {
    const { zipSync, strToU8 } = await import('fflate');
    const bytes = zipSync({ 'x.txt': strToU8('hi') });
    await expect(new DocxAdapter().load(bytes, [])).rejects.toThrow(/document\.xml/);
  });
});
