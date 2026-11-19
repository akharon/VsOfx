import { PDFCheckBox, PDFDocument, PDFDropdown, PDFRadioGroup, PDFTextField } from '@cantoo/pdf-lib';
import { describe, expect, it } from 'vitest';
import { PdfAdapter, assertPdf } from '../src/webview/adapters/pdfAdapter';
import { inverseEdit, isEdit, type Edit } from '../src/shared/model';
import { applyOp, emptyPdfState, inverseOp, isPdfOp, reduceOps, type FieldSpec, type PdfOp } from '../src/shared/pdf';

/** A two-page PDF with a text field, checkbox, dropdown and radio group. */
async function makeForm(): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const p1 = doc.addPage([400, 400]);
  doc.addPage([400, 400]);
  const form = doc.getForm();
  form.createTextField('name').addToPage(p1, { x: 20, y: 340, width: 200, height: 24 });
  form.createCheckBox('agree').addToPage(p1, { x: 20, y: 300, width: 16, height: 16 });
  const dd = form.createDropdown('country');
  dd.setOptions(['DE', 'FR', 'TR']);
  dd.addToPage(p1, { x: 20, y: 260, width: 120, height: 22 });
  const rg = form.createRadioGroup('size');
  rg.addOptionToPage('S', p1, { x: 20, y: 220, width: 14, height: 14 });
  rg.addOptionToPage('L', p1, { x: 60, y: 220, width: 14, height: 14 });
  return doc.save();
}

const edit = (op: PdfOp): Edit => ({ kind: 'pdf', op });
const spec: FieldSpec = { id: 'n1', kind: 'text', name: 'extra', page: 0, rect: [20, 100, 160, 124] };

describe('pdf edit model', () => {
  const ops: PdfOp[] = [
    { op: 'setField', name: 'a', value: 'x', old: '' },
    { op: 'addField', spec },
    { op: 'removeWidget', id: '12R', name: 'a' },
    { op: 'setRect', id: 'n1', rect: [1, 2, 3, 4], old: [0, 0, 1, 1] },
    { op: 'addText', spec: { id: 't1', page: 0, rect: [0, 0, 10, 10], text: 'hi', size: 12 } },
    { op: 'setText', id: 't1', text: 'b', old: 'a' },
    { op: 'setRotation', page: 0, deg: 90, old: 0 },
    { op: 'deletePage', page: 1 },
  ];

  it('inverse of inverse is the original, and every op validates', () => {
    for (const o of ops) {
      expect(isPdfOp(o)).toBe(true);
      expect(inverseOp(inverseOp(o))).toEqual(o);
      expect(isEdit(edit(o))).toBe(true);
      expect(inverseEdit(inverseEdit(edit(o)))).toEqual(edit(o));
    }
  });

  it('an op followed by its inverse restores the state', () => {
    for (const o of ops) {
      const s = reduceOps([o, inverseOp(o)]);
      // value-like ops end with the old value assigned; structural ops are gone again
      if (o.op === 'setField') expect(s.values.get(o.name)).toBe(o.old);
      else if (o.op === 'setRect') expect(s.rects.get(o.id)).toEqual(o.old);
      else if (o.op === 'setText') expect(s.texts.size).toBe(0);
      else if (o.op === 'setRotation') expect(s.rotations.size).toBe(0);
      else expect(s).toEqual(emptyPdfState());
    }
  });

  it('rejects malformed ops from the webview', () => {
    expect(isPdfOp({ op: 'setField', name: 1, value: 'x', old: '' })).toBe(false);
    expect(isPdfOp({ op: 'setRect', id: 'a', rect: [1, 2, 3], old: [1, 2, 3, 4] })).toBe(false);
    expect(isPdfOp({ op: 'deletePage', page: -1 })).toBe(false);
    expect(isPdfOp({ op: 'nope' })).toBe(false);
    expect(isEdit({ kind: 'pdf', op: { op: 'addField', spec: { ...spec, kind: 'script' } } })).toBe(false);
  });

  it('state is order independent for assignments', () => {
    const s = emptyPdfState();
    applyOp(s, { op: 'setField', name: 'a', value: '1', old: '' });
    applyOp(s, { op: 'setField', name: 'a', value: '2', old: '1' });
    expect(s.values.get('a')).toBe('2');
  });
});

describe('PdfAdapter', () => {
  it('rejects files that are not PDFs', () => {
    expect(() => assertPdf(new TextEncoder().encode('PK\u0003\u0004 not a pdf'))).toThrow(/not a valid PDF/);
    expect(() => assertPdf(new Uint8Array(0))).toThrow(/empty/);
  });

  it('saves untouched documents byte for byte', async () => {
    const bytes = await makeForm();
    const a = new PdfAdapter();
    await a.load(bytes, []);
    expect(await a.save()).toBe(bytes);
  });

  it('fills text, checkbox, dropdown and radio fields', async () => {
    const a = new PdfAdapter();
    await a.load(await makeForm(), []);
    a.applyEdit(edit({ op: 'setField', name: 'name', value: 'Mustafa Ö', old: '' }));
    a.applyEdit(edit({ op: 'setField', name: 'agree', value: true, old: false }));
    a.applyEdit(edit({ op: 'setField', name: 'country', value: 'TR', old: '' }));
    a.applyEdit(edit({ op: 'setField', name: 'size', value: 'L', old: '' }));
    const out = await PDFDocument.load(await a.save());
    const form = out.getForm();
    expect(form.getField('name')).toBeInstanceOf(PDFTextField);
    expect((form.getField('name') as PDFTextField).getText()).toBe('Mustafa Ö');
    expect((form.getField('agree') as PDFCheckBox).isChecked()).toBe(true);
    expect((form.getField('country') as PDFDropdown).getSelected()).toEqual(['TR']);
    expect((form.getField('size') as PDFRadioGroup).getSelected()).toBe('L');
  });

  it('undo (inverse edit) gives the same result as never editing', async () => {
    const bytes = await makeForm();
    const a = new PdfAdapter();
    await a.load(bytes, []);
    const e = edit({ op: 'setField', name: 'name', value: 'x', old: '' });
    a.applyEdit(e);
    a.applyEdit(inverseEdit(e));
    const s = await a.save();
    const form = (await PDFDocument.load(s)).getForm();
    expect((form.getField('name') as PDFTextField).getText() ?? '').toBe('');
  });

  it('replays unsaved edits on load (hot exit / revert of other editors)', async () => {
    const a = new PdfAdapter();
    const model = await a.load(await makeForm(), [edit({ op: 'setField', name: 'name', value: 'Replayed', old: '' })]);
    expect(model.type === 'pdf' && model.state.values).toEqual([['name', 'Replayed']]);
    const form = (await PDFDocument.load(await a.save())).getForm();
    expect((form.getField('name') as PDFTextField).getText()).toBe('Replayed');
  });

  it('adds, moves and removes fields', async () => {
    const a = new PdfAdapter();
    await a.load(await makeForm(), []);
    const dropdown: FieldSpec = { id: 'n2', kind: 'dropdown', name: 'pick', page: 1, rect: [10, 10, 110, 32], options: ['A', 'B'] };
    a.applyEdit(edit({ op: 'addField', spec }));
    a.applyEdit(edit({ op: 'addField', spec: dropdown }));
    a.applyEdit(edit({ op: 'addField', spec: { id: 'n3', kind: 'checkbox', name: 'tick', page: 0, rect: [200, 100, 216, 116] } }));
    a.applyEdit(edit({ op: 'setField', name: 'extra', value: 'hello', old: '' }));
    a.applyEdit(edit({ op: 'dropField', spec: { id: 'n3', kind: 'checkbox', name: 'tick', page: 0, rect: [200, 100, 216, 116] } }));
    const out = await PDFDocument.load(await a.save());
    const names = out.getForm().getFields().map((f) => f.getName()).sort();
    expect(names).toEqual(['agree', 'country', 'extra', 'name', 'pick', 'size']);
    expect((out.getForm().getField('extra') as PDFTextField).getText()).toBe('hello');
    expect((out.getForm().getField('pick') as PDFDropdown).getOptions()).toEqual(['A', 'B']);
  });

  it('removes an original field by widget id and moves another', async () => {
    const src = await makeForm();
    // find the object numbers pdf.js would report as "<n>R"
    const doc = await PDFDocument.load(src);
    const widgetRef = (field: string) => {
      const f = doc.getForm().getField(field);
      return doc.context.getObjectRef(f.acroField.getWidgets()[0]!.dict)!.objectNumber;
    };
    const a = new PdfAdapter();
    await a.load(src, []);
    a.applyEdit(edit({ op: 'removeWidget', id: `${widgetRef('agree')}R`, name: 'agree' }));
    a.applyEdit(edit({ op: 'setRect', id: `${widgetRef('name')}R`, rect: [50, 50, 250, 80], old: [20, 340, 220, 364] }));
    const out = await PDFDocument.load(await a.save());
    expect(out.getForm().getFields().map((f) => f.getName())).not.toContain('agree');
    expect(out.getForm().getTextField('name').acroField.getWidgets()[0]!.getRectangle()).toMatchObject({ x: 50, y: 50, width: 200, height: 30 });
  });

  it('prints free text, rotates and deletes pages', async () => {
    const a = new PdfAdapter();
    await a.load(await makeForm(), []);
    a.applyEdit(edit({ op: 'addText', spec: { id: 't1', page: 0, rect: [30, 30, 200, 50], text: 'Signed: Ö ✓', size: 14 } }));
    a.applyEdit(edit({ op: 'setRotation', page: 0, deg: 90, old: 0 }));
    a.applyEdit(edit({ op: 'deletePage', page: 1 }));
    const out = await PDFDocument.load(await a.save());
    expect(out.getPageCount()).toBe(1);
    expect(out.getPage(0).getRotation().angle).toBe(90);
  });

  it('refuses to delete every page', async () => {
    const a = new PdfAdapter();
    await a.load(await makeForm(), []);
    a.applyEdit(edit({ op: 'deletePage', page: 0 }));
    a.applyEdit(edit({ op: 'deletePage', page: 1 }));
    await expect(a.save()).rejects.toThrow(/at least one page/);
  });
});
