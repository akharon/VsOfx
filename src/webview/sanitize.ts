// Allowlist HTML sanitizer. It never copies markup: it walks a parsed tree and
// emits a *new* HTML string containing only allowed elements and attributes
// (doc/arch42.md §8.1 item 4). Works on xmldom trees, so it runs in the worker.
import { DOMParser } from '@xmldom/xmldom';

const ALLOWED_TAGS = new Set([
  'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'strong', 'em', 'u', 's', 'sub', 'sup',
  'ul', 'ol', 'li', 'table', 'thead', 'tbody', 'tr', 'th', 'td', 'a', 'img', 'br', 'span',
]);
const VOID_TAGS = new Set(['br', 'img']);
/** Elements removed together with everything inside them. */
const DROP_WITH_CONTENT = new Set([
  'script', 'style', 'iframe', 'object', 'embed', 'noscript', 'template', 'svg', 'math',
  'link', 'meta', 'base', 'form', 'input', 'button', 'textarea', 'select', 'audio', 'video', 'canvas', 'applet',
]);
const DATA_IMAGE = /^data:image\/(png|jpe?g|gif|bmp|webp);base64,[a-z0-9+/=\s]*$/i;

export interface SanitizeHooks {
  /** Extra attributes for an emitted element (called once per allowed element, in document order). */
  attrs?(el: Element, tag: string): Record<string, string> | undefined;
  /** If it returns a string, the element's content is replaced by that plain text. */
  textOverride?(el: Element, tag: string): string | undefined;
}

export function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** Parses mammoth-style (X)HTML. Strict XML first, lenient HTML parser as fallback. */
export function parseHtmlFragment(source: string): Element {
  const quiet = { warning() {}, error() {}, fatalError() {} };
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const make = () => new (DOMParser as any)({ errorHandler: quiet }) as DOMParser;
  try {
    const strict = {
      warning() {},
      error(m: string) {
        throw new Error(m);
      },
      fatalError(m: string) {
        throw new Error(m);
      },
    };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const doc = new (DOMParser as any)({ errorHandler: strict }).parseFromString(`<root>${source}</root>`, 'text/xml') as Document;
    return doc.documentElement;
  } catch {
    const doc = make().parseFromString(`<root>${source}</root>`, 'text/html');
    return doc.documentElement;
  }
}

export function sanitizeChildren(parent: Node, hooks: SanitizeHooks = {}): string {
  let out = '';
  for (let n = parent.firstChild; n; n = n.nextSibling) out += sanitizeNode(n, hooks);
  return out;
}

export function sanitizeHtml(source: string, hooks: SanitizeHooks = {}): string {
  return sanitizeChildren(parseHtmlFragment(source), hooks);
}

function sanitizeNode(node: Node, hooks: SanitizeHooks): string {
  if (node.nodeType === 3 || node.nodeType === 4) return escapeHtml(node.nodeValue ?? '');
  if (node.nodeType !== 1) return ''; // comments, processing instructions, ...
  const el = node as Element;
  const tag = (el.localName || el.nodeName).toLowerCase();
  if (DROP_WITH_CONTENT.has(tag)) return '';
  if (!ALLOWED_TAGS.has(tag)) return sanitizeChildren(el, hooks); // unwrap
  if (tag === 'a') {
    const href = el.getAttribute('href') ?? '';
    // Only in-document anchors survive; external links are rendered as plain text.
    if (!/^#[\w.:-]*$/.test(href)) return sanitizeChildren(el, hooks);
  }
  if (tag === 'img' && !DATA_IMAGE.test((el.getAttribute('src') ?? '').trim())) return '';

  let attrs = '';
  const add = (k: string, v: string) => (attrs += ` ${k}="${escapeHtml(v)}"`);
  if (tag === 'a') add('href', el.getAttribute('href') ?? '');
  if (tag === 'img') {
    add('src', (el.getAttribute('src') ?? '').trim());
    add('alt', el.getAttribute('alt') ?? '');
  }
  if (tag === 'td' || tag === 'th') {
    for (const k of ['colspan', 'rowspan']) {
      const v = el.getAttribute(k);
      if (v && /^\d{1,4}$/.test(v)) add(k, v);
    }
  }
  const extra = hooks.attrs?.(el, tag);
  if (extra) for (const [k, v] of Object.entries(extra)) add(k, v);

  if (VOID_TAGS.has(tag)) return `<${tag}${attrs}>`;
  const override = hooks.textOverride?.(el, tag);
  const inner = override !== undefined ? escapeHtml(override) : sanitizeChildren(el, hooks);
  return `<${tag}${attrs}>${inner}</${tag}>`;
}
