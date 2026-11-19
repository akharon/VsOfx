// XML helpers on top of @xmldom/xmldom. DOMParser does not exist in Web Workers,
// and xmldom is already part of the bundle (dependency of mammoth).
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

const handler = {
  warning() {
    /* ignore */
  },
  error(msg: string) {
    throw new Error(`Invalid XML: ${msg}`);
  },
  fatalError(msg: string) {
    throw new Error(`Invalid XML: ${msg}`);
  },
};

export function parseXml(source: string): Document {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const parser = new (DOMParser as any)({ errorHandler: handler }) as DOMParser;
  const doc = parser.parseFromString(source, 'text/xml');
  if (!doc || !doc.documentElement) throw new Error('Invalid XML: empty document');
  return doc;
}

export function serializeXml(doc: Document): string {
  const s = new XMLSerializer().serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, '');
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s;
}

/** Child elements with the given namespace and local name. */
export function childrenNS(parent: Node, ns: string, local: string): Element[] {
  const out: Element[] = [];
  for (let n = parent.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1 && (n as Element).namespaceURI === ns && (n as Element).localName === local) out.push(n as Element);
  }
  return out;
}

export function firstChildNS(parent: Node, ns: string, local: string): Element | undefined {
  return childrenNS(parent, ns, local)[0];
}

/** All descendant elements with the given namespace and local name, as a real array. */
export function descendantsNS(parent: Document | Element, ns: string, local: string): Element[] {
  return Array.from(parent.getElementsByTagNameNS(ns, local));
}

export function hasAncestor(el: Node, ns: string, local: string): boolean {
  for (let n = el.parentNode; n; n = n.parentNode) {
    if (n.nodeType === 1 && (n as Element).namespaceURI === ns && (n as Element).localName === local) return true;
  }
  return false;
}

/** Removes characters that are not allowed in XML 1.0. */
export function xmlSafe(s: string): string {
  return s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F￾￿]/g, '');
}

export function normText(s: string): string {
  return s.replace(/\s+/g, ' ').trim();
}
