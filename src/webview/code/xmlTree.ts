import { wholeLines, type ParsedTree, type Problem, type TreeNode } from './tree';

// A tolerant XML scanner that keeps exact source offsets (DOM parsers do not), so values
// can be edited in place. Well-formedness is additionally checked with the browser's parser.

export interface XAttr {
  name: string;
  /** Offset of the whole attribute (name) and of its value (inside the quotes). */
  start: number;
  valueAt: number;
  valueLen: number;
  value: string;
  quote: '"' | "'";
}

export interface XText {
  type: 'text';
  start: number;
  end: number;
  value: string;
  cdata: boolean;
}

export interface XEl {
  type: 'el';
  name: string;
  start: number;
  /** Offset just after the start tag's ">". */
  openEnd: number;
  /** Offset just after the end tag (or the self-closing tag). */
  end: number;
  selfClosing: boolean;
  attrs: XAttr[];
  children: (XEl | XText)[];
}

const ENTITIES: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

export function decodeEntities(s: string): string {
  return s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (m, e: string) => {
    if (e[0] === '#') {
      const n = e[1] === 'x' || e[1] === 'X' ? Number.parseInt(e.slice(2), 16) : Number.parseInt(e.slice(1), 10);
      return Number.isFinite(n) && n <= 0x10ffff ? String.fromCodePoint(n) : m;
    }
    return ENTITIES[e] ?? m;
  });
}

export const escapeText = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
export const escapeAttr = (s: string, q: '"' | "'") =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(q === '"' ? /"/g : /'/g, q === '"' ? '&quot;' : '&apos;');

export function scanXml(text: string): { roots: XEl[]; problems: Problem[] } {
  const problems: Problem[] = [];
  const top: XEl = { type: 'el', name: '#doc', start: 0, openEnd: 0, end: text.length, selfClosing: false, attrs: [], children: [] };
  const stack: XEl[] = [top];
  const cur = () => stack[stack.length - 1]!;
  const NAME = /[^\s/>="']+/y;
  const ATTR = /\s*([^\s=/>]+)\s*(?:=\s*("([^"]*)"|'([^']*)'))?/y;
  let i = 0;
  while (i < text.length) {
    const lt = text.indexOf('<', i);
    const textEnd = lt < 0 ? text.length : lt;
    if (textEnd > i) cur().children.push({ type: 'text', start: i, end: textEnd, value: decodeEntities(text.slice(i, textEnd)), cdata: false });
    if (lt < 0) break;
    i = lt;
    if (text.startsWith('<!--', i)) {
      const e = text.indexOf('-->', i + 4);
      i = e < 0 ? text.length : e + 3;
    } else if (text.startsWith('<![CDATA[', i)) {
      const e = text.indexOf(']]>', i + 9);
      const end = e < 0 ? text.length : e;
      cur().children.push({ type: 'text', start: i + 9, end, value: text.slice(i + 9, end), cdata: true });
      i = e < 0 ? text.length : e + 3;
    } else if (text.startsWith('<?', i)) {
      const e = text.indexOf('?>', i + 2);
      i = e < 0 ? text.length : e + 2;
    } else if (text.startsWith('<!', i)) {
      let depth = 0;
      let j = i + 2;
      for (; j < text.length; j++) {
        const c = text[j];
        if (c === '[') depth++;
        else if (c === ']') depth--;
        else if (c === '>' && depth <= 0) break;
      }
      i = j + 1;
    } else if (text.startsWith('</', i)) {
      const e = text.indexOf('>', i);
      const name = text.slice(i + 2, e < 0 ? text.length : e).trim();
      const end = e < 0 ? text.length : e + 1;
      const idx = stack.map((s) => s.name).lastIndexOf(name);
      if (idx <= 0) problems.push({ message: `Unexpected end tag </${name}>`, offset: i });
      else {
        if (idx !== stack.length - 1) problems.push({ message: `<${cur().name}> is not closed`, offset: cur().start });
        while (stack.length - 1 > idx) stack.pop()!.end = i;
        stack.pop()!.end = end;
      }
      i = end;
    } else {
      NAME.lastIndex = i + 1;
      const nm = NAME.exec(text);
      if (!nm) {
        problems.push({ message: 'Stray "<"', offset: i });
        cur().children.push({ type: 'text', start: i, end: i + 1, value: '<', cdata: false });
        i++;
        continue;
      }
      const el: XEl = { type: 'el', name: nm[0], start: i, openEnd: 0, end: 0, selfClosing: false, attrs: [], children: [] };
      let j = NAME.lastIndex;
      for (;;) {
        ATTR.lastIndex = j;
        const m = ATTR.exec(text);
        if (!m || m[0].length === 0 || m.index !== j) break;
        if (m[2] !== undefined) {
          const quote = m[2][0] as '"' | "'";
          const raw = m[3] ?? m[4] ?? '';
          const valueAt = j + m[0].length - 1 - raw.length;
          el.attrs.push({ name: m[1]!, start: j + m[0].indexOf(m[1]!), valueAt, valueLen: raw.length, value: decodeEntities(raw), quote });
        }
        j = ATTR.lastIndex;
      }
      while (j < text.length && /\s/.test(text[j]!)) j++;
      if (text.startsWith('/>', j)) {
        el.selfClosing = true;
        el.openEnd = el.end = j + 2;
        cur().children.push(el);
      } else if (text[j] === '>') {
        el.openEnd = j + 1;
        cur().children.push(el);
        stack.push(el);
      } else {
        problems.push({ message: `Malformed tag <${el.name}>`, offset: i });
        const e = text.indexOf('>', j);
        el.openEnd = el.end = e < 0 ? text.length : e + 1;
        cur().children.push(el);
      }
      i = el.openEnd;
    }
  }
  while (stack.length > 1) {
    const open = stack.pop()!;
    problems.push({ message: `<${open.name}> is not closed`, offset: open.start });
    open.end = text.length;
  }
  return { roots: top.children.filter((c): c is XEl => c.type === 'el'), problems };
}

/** The browser's own XML parser: first well-formedness error, if any. */
function domCheck(text: string): string | undefined {
  if (typeof DOMParser === 'undefined' || !text.trim()) return undefined;
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  const err = doc.getElementsByTagName('parsererror')[0];
  if (!err) return undefined;
  return (err.textContent ?? 'Not well-formed XML').replace(/\s+/g, ' ').replace(/^.*?error[^:]*:\s*/i, '').trim().slice(0, 200);
}

export function xmlTree(text: string): ParsedTree {
  const { roots, problems } = scanXml(text);
  if (problems.length === 0) {
    const msg = domCheck(text);
    if (msg) {
      const line = /line (\d+)/i.exec(msg);
      let offset = 0;
      if (line) for (let n = Number(line[1]); n > 1 && offset >= 0; n--) offset = text.indexOf('\n', offset) + 1;
      problems.push({ message: msg, offset: Math.max(0, offset) });
    }
  }
  if (roots.length > 1) problems.push({ message: 'An XML document must have exactly one root element', offset: roots[1]!.start });

  const removeRange = (s: number, e: number) => () => {
    const [a, b] = wholeLines(text, s, e);
    return text.slice(0, a) + text.slice(b);
  };

  const build = (el: XEl, id: string): TreeNode => {
    const texts = el.children.filter((c): c is XText => c.type === 'text' && (c.cdata || c.value.trim() !== ''));
    const elems = el.children.filter((c): c is XEl => c.type === 'el');
    const node: TreeNode = { id, kind: 'element', label: el.name, at: el.start, remove: removeRange(el.start, el.end) };
    const attrs: TreeNode[] = el.attrs.map((a, i) => ({
      id: `${id}@${i}`,
      kind: 'attr',
      label: `@${a.name}`,
      at: a.start,
      value: a.value,
      valueSpan: { at: a.valueAt, len: a.valueLen, encode: (s) => escapeAttr(s, a.quote) },
      remove: () => {
        let s = a.start;
        while (s > 0 && /\s/.test(text[s - 1]!)) s--;
        return text.slice(0, s) + text.slice(a.valueAt + a.valueLen + 1);
      },
    }));
    // <name>value</name> and <empty/>: shown as one editable leaf
    if (elems.length === 0 && texts.length <= 1 && el.children.every((c) => c.type === 'text')) {
      const t = texts[0];
      if (t && !t.cdata) {
        // edit the trimmed text; keep surrounding whitespace
        const raw = text.slice(t.start, t.end);
        const lead = /^\s*/.exec(raw)![0].length;
        const trail = /\s*$/.exec(raw)![0].length;
        node.value = t.value.trim();
        node.valueSpan = { at: t.start + lead, len: t.end - t.start - lead - trail, encode: escapeText };
      } else if (t) {
        node.value = t.value;
        node.valueSpan = { at: t.start, len: t.end - t.start, encode: (s) => s.replace(/]]>/g, ']]]]><![CDATA[>') };
      } else if (el.selfClosing) {
        node.value = '';
        const slash = text.lastIndexOf('/>', el.end);
        node.valueSpan = { at: slash, len: 2, encode: (s) => (s === '' ? '/>' : `>${escapeText(s)}</${el.name}>`) };
      } else {
        node.value = '';
        node.valueSpan = { at: el.openEnd, len: Math.max(0, text.lastIndexOf('</', el.end) - el.openEnd), encode: escapeText };
      }
      if (attrs.length) node.children = attrs;
      return node;
    }
    const kids: TreeNode[] = [...attrs];
    let ti = 0;
    let ei = 0;
    for (const c of el.children) {
      if (c.type === 'el') kids.push(build(c, `${id}/${ei++}`));
      else if (c.cdata || c.value.trim() !== '') {
        const raw = text.slice(c.start, c.end);
        const lead = c.cdata ? 0 : /^\s*/.exec(raw)![0].length;
        const trail = c.cdata ? 0 : /\s*$/.exec(raw)![0].length;
        kids.push({
          id: `${id}#${ti++}`,
          kind: 'text',
          label: c.cdata ? '#cdata' : '#text',
          at: c.start,
          value: c.cdata ? c.value : c.value.trim(),
          valueSpan: {
            at: c.start + lead,
            len: c.end - c.start - lead - trail,
            encode: c.cdata ? (s) => s.replace(/]]>/g, ']]]]><![CDATA[>') : escapeText,
          },
        });
      }
    }
    node.children = kids;
    node.summary = `${elems.length} ${elems.length === 1 ? 'child' : 'children'}${attrs.length ? `, ${attrs.length} attr` : ''}`;
    node.addLabel = 'Add attribute';
    node.add = () => {
      const names = new Set(el.attrs.map((a) => a.name));
      let n = 'name';
      for (let i = 2; names.has(n); i++) n = `name${i}`;
      const at = el.openEnd - (el.selfClosing ? 2 : 1);
      return `${text.slice(0, at)} ${n}=""${text.slice(at)}`;
    };
    return node;
  };

  return { roots: roots.map((r, i) => build(r, `x${i}`)), problems };
}
