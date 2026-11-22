// @vitest-environment jsdom
// Word / Writer: the page is one editing surface. Structural input (selection over several
// paragraphs, Enter, Backspace / Delete at a paragraph edge, multi-line paste) is done by the view
// in `beforeinput` and reported as ONE edit; the adapters write / replay it.
import { strFromU8 } from 'fflate';
import { describe, expect, it } from 'vitest';
import { inverseEdit, isEdit, type Edit, type TextModel } from '../src/shared/model.ts';
import { DocxAdapter } from '../src/webview/adapters/docxAdapter.ts';
import { OdtAdapter } from '../src/webview/adapters/odtAdapter.ts';
import { readZip } from '../src/webview/adapters/zip.ts';
import { createToolbar } from '../src/webview/ui/toolbar.ts';
import { renderTextView } from '../src/webview/views/textView.ts';
import { makeDocx, makeOdt, zip } from './helpers/fixtures.ts';

const flat = (e: Edit): Edit[] => (e.kind === 'batch' ? e.edits : [e]);

const ODF_NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"';
const MIME = 'application/vnd.oasis.opendocument.text';
/** Writer document: four plain paragraphs, the second with a bold part. */
function makeWriter(): Uint8Array {
  const body =
    '<office:automatic-styles><style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style></office:automatic-styles>' +
    '<office:body><office:text>' +
    '<text:p>Alpha one</text:p>' +
    '<text:p>Beta <text:span text:style-name="T1">bold</text:span> two</text:p>' +
    '<text:p>Gamma three</text:p>' +
    '<text:p>Delta four</text:p>' +
    '</office:text></office:body>';
  return zip(
    {
      mimetype: MIME,
      'content.xml': `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${ODF_NS} office:version="1.3">${body}</office:document-content>`,
      'META-INF/manifest.xml': `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="${MIME}"/></manifest:manifest>`,
    },
    ['content.xml', 'META-INF/manifest.xml', 'mimetype'],
  );
}

interface Adapter {
  load(bytes: Uint8Array, edits: Edit[]): Promise<TextModel>;
  applyEdit(e: Edit): unknown;
  save(): Promise<Uint8Array>;
}
interface Fmt {
  name: string;
  /** Four paragraphs; the second contains `Second` / `Beta` text to select in, see `sel`. */
  bytes: () => Uint8Array;
  make: () => Adapter;
  /** [paragraph, offset] of the selection start (inside paragraph 2 = index 1), end (inside paragraph 4 = index 3). */
  from: number;
  to: number;
  expectJoined: string[];
  original: string[];
}
const FORMATS: Fmt[] = [
  {
    name: 'Word (.docx)',
    bytes: () => makeDocx('simple'),
    make: () => new DocxAdapter(),
    from: 6, // "Second| paragraph with ..."
    to: 4, // "Last| paragraph."
    expectJoined: ['First paragraph.', 'Secondx paragraph.'],
    original: ['First paragraph.', 'Second paragraph with <angle> & ampersand.', '', 'Last paragraph.'],
  },
  {
    name: 'Writer (.odt)',
    bytes: makeWriter,
    make: () => new OdtAdapter(),
    from: 2, // "Be|ta ..."
    to: 3, // "Del|ta four"
    expectJoined: ['Alpha one', 'Bexta four'],
    original: ['Alpha one', 'Beta bold two', 'Gamma three', 'Delta four'],
  },
];

const texts = (html: string) => {
  const d = document.createElement('div');
  d.innerHTML = html;
  return Array.from(d.querySelectorAll<HTMLElement>('[data-p]')).map((p) => p.textContent ?? '');
};
const viewTexts = (root: HTMLElement) => Array.from(root.querySelectorAll<HTMLElement>('[data-p]')).map((p) => p.textContent ?? '');

async function mount(model: TextModel) {
  document.body.innerHTML = '<div id="root"></div>';
  const root = document.getElementById('root')!;
  const edits: Edit[] = [];
  const view = renderTextView(root, model, false, (e) => edits.push(e), createToolbar());
  return { root, edits, view };
}
function select(a: Node, ao: number, b: Node = a, bo: number = ao) {
  const r = document.createRange();
  r.setStart(a, ao);
  r.setEnd(b, bo);
  const s = getSelection()!;
  s.removeAllRanges();
  s.addRange(r);
}
const input = (target: Element, inputType: string, data?: string) => {
  const ev = new InputEvent('beforeinput', { inputType, data, cancelable: true, bubbles: true });
  target.dispatchEvent(ev);
  return ev;
};
const para = (root: HTMLElement, i: number) => root.querySelector<HTMLElement>(`[data-p="${i}"]`)!;
/** Applies edits to a fresh adapter, saves, loads the result. */
async function roundTrip(f: Fmt, edits: Edit[]): Promise<{ texts: string[]; model: TextModel; bytes: Uint8Array }> {
  const a = f.make();
  await a.load(f.bytes(), []);
  for (const e of edits.flatMap(flat)) a.applyEdit(e);
  const bytes = await a.save();
  const model = await f.make().load(bytes, []);
  return { texts: texts(model.html), model, bytes };
}

for (const f of FORMATS) {
  describe(`one editing surface: ${f.name}`, () => {
    const setup = async () => mount(await f.make().load(f.bytes(), []));

    it('makes the page the editing host (rich) and leaves blocks without data-p read-only', async () => {
      const { root } = await setup();
      expect(root.querySelector('article.page')!.getAttribute('contenteditable')).toBe('true');
      expect(para(root, 0).hasAttribute('contenteditable')).toBe(false);
    });

    it('typing over a selection from paragraph 2 into paragraph 4 is one batch; undo restores', async () => {
      const { root, edits, view } = await setup();
      expect(viewTexts(root)).toEqual(f.original);
      select(para(root, 1).firstChild!, f.from, para(root, 3).firstChild!, f.to);
      expect(input(para(root, 1), 'insertText', 'x').defaultPrevented).toBe(true);
      expect(viewTexts(root)).toEqual(f.expectJoined);
      expect(edits).toHaveLength(1);
      const batch = edits[0]!;
      expect(batch.kind).toBe('batch');
      expect(isEdit(batch)).toBe(true);
      const kinds = flat(batch).map((e) => e.kind + ('op' in e ? `:${e.op}` : ''));
      expect(kinds).toEqual(['paragraph', 'delpara:remove', 'delpara:remove']);
      // the removed paragraphs carry their content for the view to restore
      expect(flat(batch).slice(1)).toMatchObject([{ index: 3, fmt: { runs: [{ t: f.original[3] }] } }, { index: 2 }]);

      view.applyEdit(inverseEdit(batch), {});
      expect(viewTexts(root)).toEqual(f.original);
      expect(Array.from(root.querySelectorAll('[data-p]')).map((p) => (p as HTMLElement).dataset['p'])).toEqual(['0', '1', '2', '3']);
      view.applyEdit(batch, {}); // redo
      expect(viewTexts(root)).toEqual(f.expectJoined);
      view.dispose();
    });

    it('adapter: the batch is saved and replayed; the inverse brings the original back', async () => {
      const { root, edits } = await setup();
      select(para(root, 1).firstChild!, f.from, para(root, 3).firstChild!, f.to);
      input(para(root, 1), 'insertText', 'x');
      const batch = edits[0]!;
      expect((await roundTrip(f, [batch])).texts).toEqual(f.expectJoined);
      // replay on load
      const replay = await f.make().load(f.bytes(), flat(batch));
      expect(texts(replay.html)).toEqual(f.expectJoined);
      // inverse: apply, save, reload = original
      expect((await roundTrip(f, [batch, inverseEdit(batch)])).texts).toEqual(f.original);
    });

    it('Backspace at the start of a paragraph joins it to the previous one', async () => {
      const { root, edits, view } = await setup();
      select(para(root, 2).firstChild ?? para(root, 2), 0);
      const ev = input(para(root, 2), 'deleteContentBackward');
      expect(ev.defaultPrevented).toBe(true);
      expect(viewTexts(root)).toEqual(f.original.filter((_, i) => i !== 2).map((t, i) => (i === 1 ? t + f.original[2] : t)));
      expect(edits).toHaveLength(1);
      expect(flat(edits[0]!).map((e) => e.kind).at(-1)).toBe('delpara');
      view.applyEdit(inverseEdit(edits[0]!), {});
      expect(viewTexts(root)).toEqual(f.original);
      view.dispose();
    });

    it('Delete at the end of a paragraph joins the next one; undo restores it', async () => {
      const { root, edits, view } = await setup();
      const last = para(root, 2).lastChild;
      if (last) select(last, (last.nodeValue ?? '').length);
      else select(para(root, 2), 0);
      const second = f.original[2]!;
      const ev = input(para(root, 2), 'deleteContentForward');
      expect(ev.defaultPrevented).toBe(true);
      expect(viewTexts(root)[2]).toBe(second + f.original[3]);
      expect(viewTexts(root)).toHaveLength(3);
      const joined = (await roundTrip(f, edits)).texts;
      expect(joined[2]).toBe(second + f.original[3]);
      view.applyEdit(inverseEdit(edits[0]!), {});
      expect(viewTexts(root)).toEqual(f.original);
      view.dispose();
    });

    it('Enter in the middle splits: the tail moves to a new paragraph (one batch)', async () => {
      const { root, edits, view } = await setup();
      const t = para(root, 0).firstChild!;
      select(t, 3);
      expect(input(para(root, 0), 'insertParagraph').defaultPrevented).toBe(true);
      const head = f.original[0]!.slice(0, 3);
      const tail = f.original[0]!.slice(3);
      expect(viewTexts(root).slice(0, 3)).toEqual([head, tail, f.original[1]]);
      expect(edits).toHaveLength(1);
      expect(flat(edits[0]!).map((e) => e.kind)).toEqual(['paragraph', 'newpara']);
      expect((await roundTrip(f, edits)).texts.slice(0, 3)).toEqual([head, tail, f.original[1]]);
      view.applyEdit(inverseEdit(edits[0]!), {});
      expect(viewTexts(root)).toEqual(f.original);
      view.dispose();
    });

    it('Backspace at the start of an added paragraph removes it again', async () => {
      const { root, edits, view } = await setup();
      select(para(root, 0).firstChild!, 3);
      input(para(root, 0), 'insertParagraph');
      // the caret is now at the start of the added paragraph
      const added = root.querySelectorAll<HTMLElement>('[data-p]')[1]!;
      expect(Number(added.dataset['p'])).toBeGreaterThanOrEqual(2_000_000);
      input(added, 'deleteContentBackward');
      expect(viewTexts(root)).toEqual(f.original);
      expect(edits).toHaveLength(2);
      expect(flat(edits[1]!).map((e) => e.kind)).toEqual(['paragraph', 'newpara']);
      expect((await roundTrip(f, edits)).texts).toEqual(f.original);
      view.dispose();
    });

    it('pasting three lines makes three paragraphs: first line at the caret, the tail after the last', async () => {
      const { root, edits, view } = await setup();
      select(para(root, 0).firstChild!, 5);
      expect(input(para(root, 0), 'insertFromPaste', 'A\r\nB\nC').defaultPrevented).toBe(true);
      const o = f.original[0]!;
      expect(viewTexts(root).slice(0, 4)).toEqual([`${o.slice(0, 5)}A`, 'B', `C${o.slice(5)}`, f.original[1]]);
      expect(edits).toHaveLength(1);
      expect(flat(edits[0]!).map((e) => e.kind)).toEqual(['paragraph', 'newpara', 'newpara']);
      expect((await roundTrip(f, edits)).texts.slice(0, 4)).toEqual([`${o.slice(0, 5)}A`, 'B', `C${o.slice(5)}`, f.original[1]]);
      view.applyEdit(inverseEdit(edits[0]!), {});
      expect(viewTexts(root)).toEqual(f.original);
      view.dispose();
    });

    it('a single-line paste or typing inside one paragraph is left to the browser / stays one paragraph edit', async () => {
      const { root, edits, view } = await setup();
      select(para(root, 0).firstChild!, 2);
      expect(input(para(root, 0), 'insertText', 'z').defaultPrevented).toBe(false); // native
      expect(edits).toHaveLength(0);
      expect(input(para(root, 0), 'insertFromPaste', 'one line').defaultPrevented).toBe(true);
      expect(viewTexts(root)[0]).toBe(`${f.original[0]!.slice(0, 2)}one line${f.original[0]!.slice(2)}`);
      expect(viewTexts(root)).toHaveLength(4);
      expect(edits).toHaveLength(1);
      expect(edits[0]!.kind).toBe('paragraph');
      view.dispose();
    });
  });
}

describe('one editing surface: formatting survives a merge', () => {
  const bold = async () => {
    const a = new DocxAdapter();
    const bytes = makeDocx('split-runs'); // "Hello " bold, "wor" italic, "ld"; then "Plain second paragraph."
    return { model: await a.load(bytes, []), bytes };
  };
  it('Backspace at the start of paragraph 2 keeps the bold / italic runs of paragraph 1', async () => {
    const { model, bytes } = await bold();
    const { root, edits, view } = await mount(model);
    select(para(root, 1).firstChild!, 0);
    input(para(root, 1), 'deleteContentBackward');
    expect(para(root, 0).innerHTML).toBe('<strong>Hello </strong><em>wor</em>ldPlain second paragraph.');
    expect(root.querySelectorAll('[data-p]')).toHaveLength(1);
    expect(edits[0]).toMatchObject({ kind: 'batch' });
    // saved and reloaded: the runs are still bold / italic
    const a = new DocxAdapter();
    await a.load(bytes, []);
    for (const e of flat(edits[0]!)) a.applyEdit(e);
    const saved = await a.save();
    const xml = strFromU8(readZip(saved)['word/document.xml']!);
    expect(xml).toMatch(/<w:b\/>[^]*Hello /);
    expect(xml).not.toContain('Plain second paragraph.</w:t></w:r></w:p><w:p>');
    const again = await new DocxAdapter().load(saved, []);
    expect(again.html).toContain('<strong>Hello </strong>');
    expect(texts(again.html)).toEqual(['Hello worldPlain second paragraph.']);
    view.dispose();
  });

  it('Enter inside a formatted run keeps the formatting on both halves', async () => {
    const { model, bytes } = await bold();
    const { root, edits, view } = await mount(model);
    const em = para(root, 0).querySelector('em')!;
    select(em.firstChild!, 1); // "w|or"
    input(para(root, 0), 'insertParagraph');
    expect(para(root, 0).innerHTML).toBe('<strong>Hello </strong><em>w</em>');
    const added = root.querySelectorAll<HTMLElement>('[data-p]')[1]!;
    expect(added.innerHTML).toBe('<em>or</em>ld');
    const a = new DocxAdapter();
    await a.load(bytes, []);
    for (const e of flat(edits[0]!)) a.applyEdit(e);
    const again = await new DocxAdapter().load(await a.save(), []);
    expect(texts(again.html)).toEqual(['Hello w', 'orld', 'Plain second paragraph.']);
    view.dispose();
  });
});

describe('one editing surface: tables and edges (nothing is corrupted)', () => {
  const odt = async () => mount(await new OdtAdapter().load(makeOdt(), [])); // 0 h, 1 p, 2 list item, 3 table cell, 4 last

  it('a selection that crosses a table is ignored', async () => {
    const { root, edits, view } = await odt();
    const before = viewTexts(root);
    select(para(root, 1).firstChild!, 2, para(root, 4).firstChild!, 2);
    expect(input(para(root, 1), 'insertText', 'x').defaultPrevented).toBe(true);
    expect(viewTexts(root)).toEqual(before);
    expect(edits).toHaveLength(0);
    view.dispose();
  });

  it('Backspace at the start of the paragraph after a table does not join into the cell', async () => {
    const { root, edits, view } = await odt();
    const before = viewTexts(root);
    select(para(root, 4).firstChild!, 0);
    expect(input(para(root, 4), 'deleteContentBackward').defaultPrevented).toBe(true);
    expect(viewTexts(root)).toEqual(before);
    expect(edits).toHaveLength(0);
    view.dispose();
  });

  it('Backspace at the start / Delete at the end of a table cell does nothing, Enter keeps one paragraph', async () => {
    const { root, edits, view } = await odt();
    const cell = para(root, 3);
    select(cell.firstChild!, 0);
    expect(input(cell, 'deleteContentBackward').defaultPrevented).toBe(true);
    select(cell.firstChild!, (cell.firstChild!.nodeValue ?? '').length);
    expect(input(cell, 'deleteContentForward').defaultPrevented).toBe(true);
    expect(input(cell, 'insertParagraph').defaultPrevented).toBe(true);
    expect(root.querySelectorAll('[data-p]')).toHaveLength(5);
    expect(edits).toHaveLength(0);
    view.dispose();
  });

  it('a multi-line paste into a table cell becomes one line', async () => {
    const { root, edits, view } = await odt();
    const cell = para(root, 3);
    select(cell.firstChild!, 0);
    input(cell, 'insertFromPaste', 'a\nb');
    expect(cell.textContent).toBe('a bCell text');
    expect(root.querySelectorAll('[data-p]')).toHaveLength(5);
    expect(edits).toHaveLength(1);
    view.dispose();
  });

  it('deleting a whole selection (Delete over paragraphs 2-4) removes them, in a document with a list item', async () => {
    const { root, edits, view } = await mount(await new DocxAdapter().load(makeDocx('styles'), []));
    // 0 title, 1 intro, 2 bullet, 3 bullet, 4 subtitle, 5 closing
    select(para(root, 1).firstChild!, 2, para(root, 4).firstChild!, 3);
    input(para(root, 1), 'deleteContentBackward');
    expect(viewTexts(root)).toEqual(['Main Title', 'Intitle', 'Closing text.']);
    const batch = edits[0]!;
    const a = new DocxAdapter();
    await a.load(makeDocx('styles'), []);
    for (const e of flat(batch)) a.applyEdit(e);
    expect(texts((await new DocxAdapter().load(await a.save(), [])).html)).toEqual(['Main Title', 'Intitle', 'Closing text.']);
    view.applyEdit(inverseEdit(batch), {});
    expect(viewTexts(root)).toEqual(['Main Title', 'Intro text.', 'Bullet one', 'Bullet two', 'Subtitle', 'Closing text.']);
    view.dispose();
  });

  it('delpara validation and inverse', () => {
    const e: Edit = { kind: 'delpara', op: 'remove', index: 3, fmt: { runs: [{ t: 'x' }] }, to: 1 };
    expect(isEdit(e)).toBe(true);
    expect(inverseEdit(e)).toEqual({ ...e, op: 'restore' });
    expect(isEdit({ ...e, index: 1_000_000 })).toBe(false); // cells are never removed
    expect(isEdit({ ...e, op: 'delete' })).toBe(false);
    expect(isEdit({ ...e, to: -1 })).toBe(false);
    expect(isEdit({ ...e, fmt: { runs: [{ t: 'x', evil: true }] } })).toBe(false);
  });
});
