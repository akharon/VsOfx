// Shared helpers for the OpenDocument adapters (.odt, .ods).
import { parseXml, serializeXml, xmlSafe } from './xml';
import { readZip, text, writeZip, type ZipEntries } from './zip';

export const NS = {
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
  style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  calcext: 'urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0',
} as const;

export const MIME_ODT = 'application/vnd.oasis.opendocument.text';
export const MIME_ODS = 'application/vnd.oasis.opendocument.spreadsheet';

export function isEl(n: Node | null, ns: string, local: string): boolean {
  return !!n && n.nodeType === 1 && (n as Element).namespaceURI === ns && (n as Element).localName === local;
}

export interface OdfPackage {
  entries: ZipEntries;
  content: Document;
}

/** Unzips, checks the mimetype entry and parses content.xml. */
export function openOdf(bytes: Uint8Array, expectedMime: string): OdfPackage {
  const entries = readZip(bytes);
  const mime = entries['mimetype'] ? new TextDecoder().decode(entries['mimetype']).trim() : '';
  if (mime !== expectedMime) throw new Error(`Unexpected OpenDocument type "${mime || 'missing mimetype'}"`);
  return { entries, content: parseXml(text(entries, 'content.xml')) };
}

export function saveOdf(pkg: OdfPackage): Uint8Array {
  return writeZip(pkg.entries, { 'content.xml': serializeXml(pkg.content) }, { mimetypeFirst: true });
}

/** Plain text of a text:p / text:h element, expanding spaces, tabs and line breaks. */
export function odfText(el: Element): string {
  let out = '';
  const walk = (n: Node) => {
    for (let c = n.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 3 || c.nodeType === 4) {
        out += c.nodeValue ?? '';
        continue;
      }
      if (c.nodeType !== 1) continue;
      const e = c as Element;
      if (e.namespaceURI === NS.text) {
        switch (e.localName) {
          case 's':
            out += ' '.repeat(Math.min(1000, Number(e.getAttributeNS(NS.text, 'c') || 1)));
            continue;
          case 'tab':
            out += '\t';
            continue;
          case 'line-break':
            out += '\n';
            continue;
          case 'note':
          case 'soft-page-break':
            continue;
        }
      }
      if (e.namespaceURI === NS.office && e.localName === 'annotation') continue;
      walk(e);
    }
  };
  walk(el);
  return out;
}

/** Replaces the content of a paragraph/heading by `value` (spaces, tabs and newlines encoded per ODF). */
export function setOdfText(doc: Document, el: Element, value: string): void {
  while (el.firstChild) el.removeChild(el.firstChild);
  const prefix = (el.prefix ? el.prefix + ':' : 'text:');
  const tp = prefix; // text: elements share the paragraph's prefix
  let buf = '';
  const flush = () => {
    if (buf) el.appendChild(doc.createTextNode(buf));
    buf = '';
  };
  const s = xmlSafe(value);
  let i = 0;
  while (i < s.length) {
    const ch = s[i]!;
    if (ch === '\t') {
      flush();
      el.appendChild(doc.createElementNS(NS.text, `${tp}tab`));
      i++;
    } else if (ch === '\n' || ch === '\r') {
      flush();
      el.appendChild(doc.createElementNS(NS.text, `${tp}line-break`));
      i += ch === '\r' && s[i + 1] === '\n' ? 2 : 1;
    } else if (ch === ' ') {
      let n = 0;
      while (s[i + n] === ' ') n++;
      // A single space inside text is literal; leading spaces and runs use text:s.
      let literal = 0;
      if (buf.length > 0 && !buf.endsWith(' ') && n >= 1) {
        buf += ' ';
        literal = 1;
      }
      const rest = n - literal;
      if (rest > 0) {
        flush();
        const sp = doc.createElementNS(NS.text, `${tp}s`);
        if (rest > 1) sp.setAttributeNS(NS.text, `${tp}c`, String(rest));
        el.appendChild(sp);
      }
      i += n;
    } else {
      buf += ch;
      i++;
    }
  }
  flush();
}

/** Collects text:p and text:h in document order, skipping frames, notes and annotations. */
export function collectParagraphs(root: Element): Element[] {
  const out: Element[] = [];
  const walk = (n: Node) => {
    for (let c = n.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      const e = c as Element;
      if (e.namespaceURI === NS.text && (e.localName === 'p' || e.localName === 'h')) {
        out.push(e);
        continue; // paragraphs do not nest
      }
      if (e.namespaceURI === NS.draw && e.localName === 'frame') continue;
      if (e.namespaceURI === NS.office && e.localName === 'annotation') continue;
      if (e.namespaceURI === NS.text && e.localName === 'note') continue;
      walk(e);
    }
  };
  walk(root);
  return out;
}
