import * as MammothNS from 'mammoth';
import type { ApplyResult, Edit, TextModel } from '../../shared/model';
import { CELL_BASE, cellIndex, fmtEquals, runsToHtml, type ParaFmt, type TableSpec } from '../../shared/richtext';
import { buildTable, DocxParts, insertTablesAfter, writeParagraph } from './docxWrite';
import { sanitizeChildren, parseHtmlFragment } from '../sanitize';
import type { FormatAdapter } from './Adapter';
import { descendantsNS, hasAncestor, normText, parseXml, serializeXml, xmlSafe } from './xml';
import { readZip, text, writeZip, type ZipEntries } from './zip';

const mammoth: typeof MammothNS = (MammothNS as unknown as { default?: typeof MammothNS }).default ?? MammothNS;

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const KEEP_IN_PARAGRAPH = new Set(['pPr', 'bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd']);
/** How far ahead to look for the paragraph that matches a rendered block. */
const LOOKAHEAD = 200;

/** Paragraphs of the body in document order (including table cells, excluding text boxes). */
function bodyParagraphs(doc: Document): Element[] {
  return descendantsNS(doc, W, 'p').filter((p) => !hasAncestor(p, W, 'txbxContent'));
}

/** Visible text of a w:p (what mammoth renders), tabs as \t, deleted text excluded. */
function paragraphText(p: Element): string {
  let out = '';
  const walk = (n: Node) => {
    for (let c = n.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      const el = c as Element;
      if (el.namespaceURI !== W) {
        walk(el);
        continue;
      }
      switch (el.localName) {
        case 't':
          out += el.textContent ?? '';
          break;
        case 'tab':
          out += '\t';
          break;
        case 'br':
        case 'cr':
          out += '\n';
          break;
        case 'del':
        case 'delText':
        case 'instrText':
          break;
        default:
          walk(el);
      }
    }
  };
  walk(p);
  return out;
}

interface ParaState {
  text: string;
  fmt?: ParaFmt;
}

export class DocxAdapter implements FormatAdapter {
  readonly readonly = false;
  private entries!: ZipEntries;
  private paragraphs: string[] = []; // original text per body paragraph
  private html = '';
  /** paragraph index -> edited content (only paragraphs that differ from the original). */
  private readonly edited = new Map<number, ParaState>();
  /** State before the first edit of a paragraph, to detect "back at the original". */
  private readonly originals = new Map<number, ParaState>();
  /** Tables inserted by the user, by id. Their cells are paragraphs with virtual indices. */
  private readonly tables = new Map<number, TableSpec>();
  /** Fraction of non-empty paragraphs that could be matched to a rendered block (for tests/diagnostics). */
  mappedRatio = 1;

  async load(bytes: Uint8Array, edits: Edit[]): Promise<TextModel> {
    this.entries = readZip(bytes);
    const doc = parseXml(text(this.entries, 'word/document.xml'));
    this.paragraphs = bodyParagraphs(doc).map(paragraphText);

    const ab = bytes.slice().buffer as ArrayBuffer;
    const input: { arrayBuffer: ArrayBuffer; buffer?: unknown } = { arrayBuffer: ab };
    // The Node build of mammoth (used by unit tests) reads `buffer` instead of `arrayBuffer`.
    const B = (globalThis as { Buffer?: { from(b: ArrayBuffer): unknown } }).Buffer;
    if (B) input.buffer = B.from(ab);
    const result = await mammoth.convertToHtml(input as { arrayBuffer: ArrayBuffer });
    this.html = result.value;

    for (const e of edits) this.applyEdit(e);
    const tables = [...this.tables.values()].map((table) => {
      const cells: string[] = [];
      for (let r = 0; r < table.rows; r++) {
        for (let c = 0; c < table.cols; c++) {
          const s = this.edited.get(cellIndex(table, r, c));
          cells.push(s?.fmt ? runsToHtml(s.fmt.runs) : '');
        }
      }
      return { table, cells };
    });
    return { type: 'text', html: this.render(), rich: true, tables };
  }

  applyEdit(edit: Edit): ApplyResult {
    if (edit.kind === 'table') {
      if (edit.op === 'insert') this.tables.set(edit.table.id, edit.table);
      else this.tables.delete(edit.table.id);
      return {};
    }
    if (edit.kind !== 'paragraph') throw new Error('Not a paragraph edit');
    const isCell = edit.index >= CELL_BASE;
    const original = this.paragraphs[edit.index];
    if (original === undefined && !isCell) throw new Error(`Paragraph ${edit.index} does not exist`);
    if (!this.originals.has(edit.index)) {
      this.originals.set(edit.index, edit.oldFmt ? { text: edit.old, fmt: edit.oldFmt } : { text: original ?? '' });
    }
    const first = this.originals.get(edit.index)!;
    // Back at the original: forget the edit so the original runs are kept on save.
    const same = edit.fmt || first.fmt ? fmtEquals(edit.fmt, first.fmt) : normText(edit.text) === normText(first.text);
    if (same && !isCell) this.edited.delete(edit.index);
    else this.edited.set(edit.index, edit.fmt ? { text: edit.text, fmt: edit.fmt } : { text: edit.text });
    return {};
  }

  async save(): Promise<Uint8Array> {
    if (this.edited.size === 0 && this.tables.size === 0) return writeZip(this.entries, {});
    const doc = parseXml(text(this.entries, 'word/document.xml'));
    const paras = bodyParagraphs(doc);
    const parts = new DocxParts(this.entries, doc);
    for (const [index, s] of this.edited) {
      const p = paras[index];
      if (!p) continue; // table cells are written below
      if (s.fmt) writeParagraph(parts, p, s.fmt);
      else patchParagraph(doc, p, s.text);
    }
    const byAnchor = new Map<number, TableSpec[]>();
    for (const t of [...this.tables.values()].sort((a, b) => a.id - b.id)) {
      byAnchor.set(t.after, [...(byAnchor.get(t.after) ?? []), t]);
    }
    for (const [after, specs] of byAnchor) {
      const p = paras[after];
      if (!p) continue;
      const pfx = p.prefix ? `${p.prefix}:` : '';
      const tables = specs.map((spec) =>
        buildTable(parts, pfx, spec, (r, c) => {
          const s = this.edited.get(cellIndex(spec, r, c));
          return s?.fmt ?? (s ? { runs: [{ t: s.text }] } : undefined);
        }),
      );
      insertTablesAfter(p, tables);
    }
    return writeZip(this.entries, { ...parts.replacements(), 'word/document.xml': serializeXml(doc) });
  }

  /** Renders the sanitized HTML, tagging blocks that map to paragraphs with data-p. */
  private render(): string {
    const root = parseHtmlFragment(this.html);
    const edited = this.edited;
    let cursor = 0;
    let mapped = 0;
    let candidates = 0;
    const assigned = new Map<Element, number>();

    const hooks = {
      attrs: (el: Element, tag: string) => {
        if (!/^(p|h[1-6]|li)$/.test(tag)) return undefined;
        const t = normText(blockText(el));
        if (!t) return undefined;
        candidates++;
        const limit = Math.min(this.paragraphs.length, cursor + LOOKAHEAD);
        for (let j = cursor; j < limit; j++) {
          const original = normText(this.paragraphs[j] ?? '');
          if (original === t || (edited.has(j) && normText(edited.get(j)?.text ?? '') === t)) {
            cursor = j + 1;
            mapped++;
            assigned.set(el, j);
            const block = edited.get(j)?.fmt?.block;
            const attrs: Record<string, string> = { 'data-p': String(j) };
            if (block) attrs['data-block'] = block;
            return attrs;
          }
        }
        return undefined;
      },
      htmlOverride: (el: Element) => {
        const j = assigned.get(el);
        const fmt = j !== undefined ? edited.get(j)?.fmt : undefined;
        return fmt ? runsToHtml(fmt.runs) : undefined;
      },
      textOverride: (el: Element) => {
        const j = assigned.get(el);
        return j !== undefined && edited.has(j) ? edited.get(j)!.text : undefined;
      },
    };
    const html = sanitizeChildren(root, hooks);
    this.mappedRatio = candidates === 0 ? 1 : mapped / candidates;
    return html;
  }
}

/** Text of a block without nested lists/tables. */
function blockText(el: Element): string {
  let out = '';
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 3 || n.nodeType === 4) out += n.nodeValue ?? '';
    else if (n.nodeType === 1) {
      const tag = ((n as Element).localName || '').toLowerCase();
      if (tag === 'ul' || tag === 'ol' || tag === 'table') continue;
      out += tag === 'br' ? ' ' : blockText(n as Element);
    }
  }
  return out;
}

/**
 * Replaces a paragraph's content by one run with `newText`. Keeps the paragraph
 * properties and the formatting (w:rPr) of the paragraph's first run.
 */
function patchParagraph(doc: Document, p: Element, newText: string): void {
  const prefix = p.prefix ? `${p.prefix}:` : '';
  const firstRun = descendantsNS(p, W, 'r')[0];
  let rPr: Element | undefined;
  if (firstRun) {
    for (let n = firstRun.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 1 && (n as Element).namespaceURI === W && (n as Element).localName === 'rPr') rPr = n as Element;
    }
  }
  const rPrClone = rPr ? (rPr.cloneNode(true) as Element) : undefined;

  for (const child of Array.from(p.childNodes)) {
    const keep = child.nodeType === 1 && (child as Element).namespaceURI === W && KEEP_IN_PARAGRAPH.has((child as Element).localName);
    if (!keep) p.removeChild(child);
  }

  const run = doc.createElementNS(W, `${prefix}r`);
  if (rPrClone) run.appendChild(rPrClone);
  const parts = xmlSafe(newText).split(/(\t|\r\n|\n|\r)/);
  for (const part of parts) {
    if (part === '') continue;
    if (part === '\t') run.appendChild(doc.createElementNS(W, `${prefix}tab`));
    else if (/^[\r\n]/.test(part)) run.appendChild(doc.createElementNS(W, `${prefix}br`));
    else {
      const t = doc.createElementNS(W, `${prefix}t`);
      t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
      t.appendChild(doc.createTextNode(part));
      run.appendChild(t);
    }
  }
  p.appendChild(run);
}
