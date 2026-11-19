import type { ApplyResult, Edit, TextModel } from '../../shared/model';
import { escapeHtml } from '../sanitize';
import type { FormatAdapter } from './Adapter';
import { collectParagraphs, isEl, MIME_ODT, NS, odfText, openOdf, saveOdf, setOdfText, type OdfPackage } from './odf';
import { descendantsNS } from './xml';

interface Original {
  text: string;
  nodes: Node[];
}

interface InlineStyle {
  bold: boolean;
  italic: boolean;
  underline: boolean;
}

/** .odt: paragraphs and headings of content.xml, text only (no images, no layout). */
export class OdtAdapter implements FormatAdapter {
  readonly readonly = false;
  private pkg!: OdfPackage;
  private paras: Element[] = [];
  private readonly originals = new Map<number, Original>();
  private styles = new Map<string, InlineStyle>();

  async load(bytes: Uint8Array, edits: Edit[]): Promise<TextModel> {
    this.pkg = openOdf(bytes, MIME_ODT);
    const body = descendantsNS(this.pkg.content, NS.office, 'text')[0];
    if (!body) throw new Error('This OpenDocument file has no text body');
    this.paras = collectParagraphs(body);
    this.styles = readStyles(this.pkg.content);
    for (const e of edits) this.applyEdit(e);
    return { type: 'text', html: this.render(body) };
  }

  applyEdit(edit: Edit): ApplyResult {
    if (edit.kind !== 'paragraph') throw new Error('Not a paragraph edit');
    const el = this.paras[edit.index];
    if (!el) throw new Error(`Paragraph ${edit.index} does not exist`);
    let orig = this.originals.get(edit.index);
    if (!orig) {
      orig = { text: odfText(el), nodes: Array.from(el.childNodes).map((n) => n.cloneNode(true)) };
      this.originals.set(edit.index, orig);
    }
    if (edit.text === orig.text) {
      while (el.firstChild) el.removeChild(el.firstChild);
      for (const n of orig.nodes) el.appendChild(n.cloneNode(true));
    } else {
      setOdfText(this.pkg.content, el, edit.text);
    }
    return {};
  }

  async save(): Promise<Uint8Array> {
    return saveOdf(this.pkg);
  }

  // ---- rendering ---------------------------------------------------------

  private render(body: Element): string {
    const index = new Map<Element, number>();
    this.paras.forEach((p, i) => index.set(p, i));
    const blocks = (parent: Node): string => {
      let out = '';
      for (let c = parent.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        const e = c as Element;
        if (e.namespaceURI === NS.text && (e.localName === 'p' || e.localName === 'h')) {
          const i = index.get(e);
          const tag = e.localName === 'h' ? `h${Math.min(6, Math.max(1, Number(e.getAttributeNS(NS.text, 'outline-level') || 1)))}` : 'p';
          const attr = i === undefined ? '' : ` data-p="${i}"`;
          out += `<${tag}${attr}>${this.inline(e, this.styleOf(e))}</${tag}>`;
        } else if (isEl(e, NS.text, 'list')) {
          out += `<ul>${blocks(e)}</ul>`;
        } else if (isEl(e, NS.text, 'list-item') || isEl(e, NS.text, 'list-header')) {
          out += `<li>${blocks(e)}</li>`;
        } else if (isEl(e, NS.table, 'table')) {
          out += `<table><tbody>${blocks(e)}</tbody></table>`;
        } else if (isEl(e, NS.table, 'table-row')) {
          out += `<tr>${blocks(e)}</tr>`;
        } else if (isEl(e, NS.table, 'table-cell')) {
          const span = Number(e.getAttributeNS(NS.table, 'number-columns-spanned') || 1);
          out += `<td${span > 1 && span < 1000 ? ` colspan="${span}"` : ''}>${blocks(e)}</td>`;
        } else if (isEl(e, NS.draw, 'frame') || isEl(e, NS.office, 'annotation') || isEl(e, NS.text, 'note')) {
          continue;
        } else {
          out += blocks(e); // sections, table groups, indexes, ...
        }
      }
      return out;
    };
    return blocks(body);
  }

  private styleOf(el: Element): InlineStyle | undefined {
    const name = el.getAttributeNS(NS.text, 'style-name');
    return name ? this.styles.get(name) : undefined;
  }

  private inline(el: Element, inherited?: InlineStyle): string {
    let out = '';
    for (let c = el.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 3 || c.nodeType === 4) {
        out += escapeHtml(c.nodeValue ?? '');
        continue;
      }
      if (c.nodeType !== 1) continue;
      const e = c as Element;
      if (e.namespaceURI === NS.text) {
        if (e.localName === 's') {
          out += '&nbsp;'.repeat(Math.min(1000, Number(e.getAttributeNS(NS.text, 'c') || 1)));
          continue;
        }
        if (e.localName === 'tab') {
          out += '&emsp;';
          continue;
        }
        if (e.localName === 'line-break') {
          out += '<br>';
          continue;
        }
        if (e.localName === 'note' || e.localName === 'soft-page-break') continue;
        if (e.localName === 'span') {
          const s = this.styleOf(e);
          let html = this.inline(e, s);
          if (s?.bold && !inherited?.bold) html = `<strong>${html}</strong>`;
          if (s?.italic && !inherited?.italic) html = `<em>${html}</em>`;
          if (s?.underline && !inherited?.underline) html = `<u>${html}</u>`;
          out += html;
          continue;
        }
      }
      if (isEl(e, NS.office, 'annotation') || isEl(e, NS.draw, 'frame')) continue;
      out += this.inline(e, inherited);
    }
    return out;
  }
}

function readStyles(doc: Document): Map<string, InlineStyle> {
  const map = new Map<string, InlineStyle>();
  for (const style of descendantsNS(doc, NS.style, 'style')) {
    const name = style.getAttributeNS(NS.style, 'name');
    const props = descendantsNS(style, NS.style, 'text-properties')[0];
    if (!name || !props) continue;
    const weight = props.getAttributeNS(NS.fo, 'font-weight');
    const fstyle = props.getAttributeNS(NS.fo, 'font-style');
    const underline = props.getAttributeNS(NS.style, 'text-underline-style');
    map.set(name, {
      bold: weight === 'bold' || /^[6-9]00$/.test(weight ?? ''),
      italic: fstyle === 'italic' || fstyle === 'oblique',
      underline: !!underline && underline !== 'none',
    });
  }
  return map;
}
