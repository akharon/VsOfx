import * as MammothNS from 'mammoth';
import type { ApplyResult, Edit, TextModel } from '../../shared/model';
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

export class DocxAdapter implements FormatAdapter {
  readonly readonly = false;
  private entries!: ZipEntries;
  private paragraphs: string[] = []; // original text per body paragraph
  private html = '';
  /** paragraph index -> edited text (only paragraphs that differ from the original). */
  private readonly edited = new Map<number, string>();
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
    return { type: 'text', html: this.render() };
  }

  applyEdit(edit: Edit): ApplyResult {
    if (edit.kind !== 'paragraph') throw new Error('Not a paragraph edit');
    const original = this.paragraphs[edit.index];
    if (original === undefined) throw new Error(`Paragraph ${edit.index} does not exist`);
    // Back at the original text: forget the edit so the original runs are kept on save.
    if (normText(edit.text) === normText(original)) this.edited.delete(edit.index);
    else this.edited.set(edit.index, edit.text);
    return {};
  }

  async save(): Promise<Uint8Array> {
    if (this.edited.size === 0) return writeZip(this.entries, {});
    const doc = parseXml(text(this.entries, 'word/document.xml'));
    const paras = bodyParagraphs(doc);
    for (const [index, t] of this.edited) {
      const p = paras[index];
      if (p) patchParagraph(doc, p, t);
    }
    return writeZip(this.entries, { 'word/document.xml': serializeXml(doc) });
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
          if (original === t || (edited.has(j) && normText(edited.get(j) ?? '') === t)) {
            cursor = j + 1;
            mapped++;
            assigned.set(el, j);
            return { 'data-p': String(j) };
          }
        }
        return undefined;
      },
      textOverride: (el: Element) => {
        const j = assigned.get(el);
        return j !== undefined && edited.has(j) ? edited.get(j) : undefined;
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
