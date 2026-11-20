import {
  PDFCheckBox,
  PDFDocument,
  PDFDropdown,
  PDFOptionList,
  PDFRadioGroup,
  PDFTextField,
  StandardFonts,
  degrees,
  rgb,
} from '@cantoo/pdf-lib';
import type { ApplyResult, DocModel, Edit } from '../../shared/model';
import { applyOp, emptyPdfState, stateToJson, type FieldSpec, type PdfState, type Rect } from '../../shared/pdf';
import type { FormatAdapter } from './Adapter';

/** True if the bytes start like a PDF ("%PDF-" within the first KiB). */
export function isPdf(bytes: Uint8Array): boolean {
  const head = new TextDecoder('latin1').decode(bytes.subarray(0, 1024));
  return head.includes('%PDF-');
}

export function assertPdf(bytes: Uint8Array): void {
  if (bytes.length === 0) throw new Error('The file is empty.');
  if (!isPdf(bytes)) throw new Error('This file is not a valid PDF document.');
}

const box = (r: Rect) => ({
  x: Math.min(r[0], r[2]),
  y: Math.min(r[1], r[3]),
  width: Math.abs(r[2] - r[0]),
  height: Math.abs(r[3] - r[1]),
});

/** Object number from a pdf.js annotation id such as "12R". */
function objNum(id: string): number | undefined {
  const m = /^(\d+)R/.exec(id);
  return m ? Number(m[1]) : undefined;
}

/**
 * PDF: the view (pdf.js) shows the original bytes; this adapter keeps the edit state
 * (shared/pdf.ts) and writes "original + state" with pdf-lib on save.
 */
export class PdfAdapter implements FormatAdapter {
  readonly readonly = false;
  private bytes: Uint8Array = new Uint8Array(0);
  private state: PdfState = emptyPdfState();

  async load(bytes: Uint8Array, edits: Edit[]): Promise<DocModel> {
    assertPdf(bytes);
    this.bytes = bytes;
    this.state = emptyPdfState();
    for (const e of edits) if (e.kind === 'pdf') applyOp(this.state, e.op);
    return { type: 'pdf', state: stateToJson(this.state) };
  }

  applyEdit(edit: Edit): ApplyResult {
    if (edit.kind !== 'pdf') throw new Error('Not a PDF edit');
    applyOp(this.state, edit.op);
    return {};
  }

  async save(): Promise<Uint8Array> {
    const s = this.state;
    const untouched =
      s.values.size + s.added.size + s.removed.size + s.rects.size + s.texts.size + s.images.size + s.rotations.size + s.deletedPages.size === 0;
    if (untouched) return this.bytes;

    const doc = await PDFDocument.load(this.bytes, { ignoreEncryption: true, updateMetadata: false });
    const form = doc.getForm();
    const pages = doc.getPages();
    const page = (i: number) => {
      const p = pages[i];
      if (!p) throw new Error(`Page ${i + 1} does not exist`);
      return p;
    };

    // 1. new fields
    for (const spec of s.added.values()) await this.addField(doc, spec, s.rects.get(spec.id) ?? spec.rect, page(spec.page));

    // 2. original widgets: move/resize and remove
    const wanted = new Map<number, Rect>();
    for (const [id, r] of s.rects) {
      const n = objNum(id);
      if (n !== undefined) wanted.set(n, r);
    }
    const removedNums = new Set<number>();
    for (const id of s.removed) {
      const n = objNum(id);
      if (n !== undefined) removedNums.add(n);
    }
    if (wanted.size > 0 || removedNums.size > 0) {
      for (const field of form.getFields()) {
        const widgets = field.acroField.getWidgets();
        // backwards so that removing a widget does not shift the indices still to visit
        for (let i = widgets.length - 1; i >= 0; i--) {
          const w = widgets[i]!;
          const wref = doc.context.getObjectRef(w.dict);
          if (!wref) continue;
          const r = wanted.get(wref.objectNumber);
          if (r) w.setRectangle(box(r));
          if (removedNums.has(wref.objectNumber)) {
            if (field.acroField.getWidgets().length === 1) {
              form.removeField(field);
              break;
            }
            for (const p of pages) p.node.removeAnnot(wref);
            field.acroField.removeWidget(i);
          }
        }
      }
    }

    // 3. values
    for (const [name, value] of s.values) {
      let field;
      try {
        field = form.getField(name);
      } catch {
        continue; // the field was removed
      }
      try {
        if (field instanceof PDFTextField) {
          const t = String(value);
          if (t === '') field.setText(undefined);
          else field.setText(t);
        } else if (field instanceof PDFCheckBox) {
          if (value === true || value === 'true') field.check();
          else field.uncheck();
        } else if (field instanceof PDFRadioGroup || field instanceof PDFDropdown || field instanceof PDFOptionList) {
          if (value === '' || value === false) field.clear();
          else field.select(String(value));
        }
      } catch (err) {
        throw new Error(`Field "${name}": ${err instanceof Error ? err.message : String(err)}`);
      }
    }

    // 4. free text
    if (s.texts.size > 0) {
      const fonts = new Map<string, Awaited<ReturnType<PDFDocument['embedFont']>>>();
      const fontFor = async (t: { bold?: boolean; italic?: boolean }) => {
        const name =
          t.bold && t.italic ? StandardFonts.HelveticaBoldOblique : t.bold ? StandardFonts.HelveticaBold : t.italic ? StandardFonts.HelveticaOblique : StandardFonts.Helvetica;
        let f = fonts.get(name);
        if (!f) fonts.set(name, (f = await doc.embedFont(name)));
        return f;
      };
      for (const t of s.texts.values()) {
        const font = await fontFor(t);
        const known = new Set(font.getCharacterSet());
        const hex = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(t.color ?? '');
        const color = hex ? rgb(parseInt(hex[1]!, 16) / 255, parseInt(hex[2]!, 16) / 255, parseInt(hex[3]!, 16) / 255) : rgb(0, 0, 0);
        const r = box(s.rects.get(t.id) ?? t.rect);
        const text = [...t.text].map((c) => (known.has(c.codePointAt(0)!) || c === '\n' ? c : '?')).join('');
        // From the top of the box, wrapped at its width: the same layout the editor shows.
        const ascent = font.heightAtSize(t.size, { descender: false });
        page(t.page).drawText(text, {
          x: r.x,
          y: r.y + r.height - ascent,
          size: t.size,
          font,
          color,
          lineHeight: t.size * 1.2,
          maxWidth: Math.max(r.width, t.size),
          wordBreaks: [' ', '-'],
        });
      }
    }

    // 4b. images
    for (const im of s.images.values()) {
      const raw = Uint8Array.from(atob(im.data), (c) => c.charCodeAt(0));
      const img = im.mime === 'image/png' ? await doc.embedPng(raw) : await doc.embedJpg(raw);
      page(im.page).drawImage(img, box(s.rects.get(im.id) ?? im.rect));
    }

    // 5. rotation, then deletion (highest index first)
    for (const [i, deg] of s.rotations) {
      const p = page(i);
      p.setRotation(degrees((p.getRotation().angle + deg) % 360));
    }
    for (const i of [...s.deletedPages].sort((a, b) => b - a)) {
      if (i < doc.getPageCount()) doc.removePage(i);
    }
    if (doc.getPageCount() === 0) throw new Error('A PDF needs at least one page. Restore a deleted page before saving.');
    return doc.save();
  }

  private async addField(doc: PDFDocument, spec: FieldSpec, r: Rect, page: ReturnType<PDFDocument['getPage']>): Promise<void> {
    const form = doc.getForm();
    const opts = { ...box(r), borderWidth: 1 };
    switch (spec.kind) {
      case 'text': {
        const f = form.createTextField(spec.name);
        if (spec.multiline) f.enableMultiline();
        const name =
          spec.bold && spec.italic ? StandardFonts.HelveticaBoldOblique : spec.bold ? StandardFonts.HelveticaBold : spec.italic ? StandardFonts.HelveticaOblique : StandardFonts.Helvetica;
        const hex = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(spec.color ?? '');
        f.addToPage(page, {
          ...opts,
          font: await doc.embedFont(name),
          ...(hex ? { textColor: rgb(parseInt(hex[1]!, 16) / 255, parseInt(hex[2]!, 16) / 255, parseInt(hex[3]!, 16) / 255) } : {}),
        });
        if (spec.size) f.setFontSize(spec.size);
        break;
      }
      case 'checkbox':
        form.createCheckBox(spec.name).addToPage(page, opts);
        break;
      case 'dropdown': {
        const f = form.createDropdown(spec.name);
        f.setOptions(spec.options && spec.options.length > 0 ? spec.options : ['Option 1']);
        f.addToPage(page, opts);
        break;
      }
    }
  }
}
