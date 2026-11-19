import { isAlias, isMap, isPair, isScalar, isSeq, parse, parseAllDocuments, type Document, type Node, type Scalar } from 'yaml';
import { wholeLines, type NodeKind, type ParsedTree, type Span, type TreeNode } from './tree';

type Path = (string | number)[];

/** Parses `s` as a YAML scalar; undefined if it is not a single plain scalar. */
function scalarValue(s: string): unknown {
  try {
    const v: unknown = parse(s);
    return v !== null && typeof v === 'object' ? undefined : v;
  } catch {
    return undefined;
  }
}

/** Source text for a scalar, keeping the original quoting style where possible. */
function encodeScalar(s: string, orig: Scalar): string {
  const origIsString = typeof orig.value === 'string';
  if (s.includes('\n')) return JSON.stringify(s);
  if (orig.type === 'QUOTE_SINGLE') return `'${s.replace(/'/g, "''")}'`;
  if (orig.type === 'QUOTE_DOUBLE') return JSON.stringify(s);
  if (s === '') return origIsString ? "''" : '';
  const parsed = scalarValue(s);
  if (!origIsString && parsed !== undefined) return s; // number, bool, null, date: as typed
  return parsed === s ? s : JSON.stringify(s);
}

function kindOf(v: unknown): NodeKind {
  if (v === null || v === undefined) return 'null';
  if (typeof v === 'string') return 'string';
  if (typeof v === 'number' || typeof v === 'bigint') return 'number';
  if (typeof v === 'boolean') return 'boolean';
  return 'other';
}

export function yamlTree(text: string): ParsedTree {
  const docs = parseAllDocuments(text, { prettyErrors: false, keepSourceTokens: false });
  const list: Document[] = Array.isArray(docs) ? docs : [docs as unknown as Document];
  const problems = list.flatMap((d) => d.errors.map((e) => ({ message: e.message.split('\n')[0] ?? e.message, offset: e.pos[0] })));

  const scalarSpan = (n: Scalar): Span | undefined => {
    const r = n.range;
    if (!r) return undefined;
    const del = text.slice(r[0], r[1]);
    const trailing = /\s*$/.exec(del)![0]; // block scalars include their line break
    const needSpace = r[1] === r[0] && r[0] > 0 && text[r[0] - 1] === ':';
    return {
      at: r[0],
      len: r[1] - r[0],
      encode: (s) => (needSpace ? ' ' : '') + encodeScalar(s, n) + (trailing.includes('\n') ? trailing : ''),
    };
  };

  const removal = (doc: Document, path: Path, start: number, end: number, flow: boolean) => (): string | null => {
    if (flow) {
      const copy = doc.clone();
      copy.deleteIn(path);
      return copy.toString();
    }
    const [s, e] = wholeLines(text, start, end);
    return text.slice(0, s) + text.slice(e);
  };

  const build = (doc: Document, node: unknown, id: string, label: string, path: Path, at: number): TreeNode => {
    if (isAlias(node)) return { id, kind: 'other', label, value: `*${node.source}`, at: node.range?.[0] ?? at };
    if (isMap(node)) {
      const items = node.items;
      return {
        id,
        kind: 'object',
        label,
        at: node.range?.[0] ?? at,
        summary: `${items.length} ${items.length === 1 ? 'key' : 'keys'}`,
        children: items.filter(isPair).map((p, i) => {
          const keyNode = p.key as Node | null;
          const key = isScalar(keyNode) ? String(keyNode.value) : String(keyNode);
          const child = build(doc, p.value, `${id}/${i}`, key, [...path, isScalar(keyNode) ? (keyNode.value as string) : key], keyNode?.range?.[0] ?? at);
          if (isScalar(keyNode)) {
            const span = scalarSpan(keyNode);
            if (span) child.labelSpan = span;
          }
          const start = keyNode?.range?.[0];
          const end = (p.value as Node | null)?.range?.[2] ?? keyNode?.range?.[2];
          if (start !== undefined && end !== undefined) child.remove = removal(doc, [...path, isScalar(keyNode) ? (keyNode.value as string) : key], start, end, !!node.flow);
          return child;
        }),
        addLabel: 'Add key',
        add: () => {
          const copy = doc.clone();
          const keys = new Set(items.map((p) => String(isScalar(p.key) ? p.key.value : p.key)));
          let k = 'key';
          for (let i = 2; keys.has(k); i++) k = `key${i}`;
          if (path.length === 0) copy.set(k, '');
          else copy.setIn([...path, k], '');
          return node.flow || items.length === 0 ? copy.toString() : insertAfter(text, items[items.length - 1]!, `${k}: ''`);
        },
      };
    }
    if (isSeq(node)) {
      const items = node.items;
      return {
        id,
        kind: 'array',
        label,
        at: node.range?.[0] ?? at,
        summary: `${items.length} ${items.length === 1 ? 'item' : 'items'}`,
        children: items.map((it, i) => {
          const child = build(doc, it, `${id}/${i}`, String(i), [...path, i], at);
          const r = (it as Node | null)?.range;
          if (r) {
            // include the "- " marker of block sequences
            let s = r[0];
            if (!node.flow) {
              const before = text.lastIndexOf('-', r[0]);
              if (before >= 0 && /^-\s*$/.test(text.slice(before, r[0]))) s = before;
            }
            child.remove = removal(doc, [...path, i], s, r[2], !!node.flow);
          }
          return child;
        }),
        addLabel: 'Add item',
        add: () => {
          const copy = doc.clone();
          if (path.length === 0) copy.add('');
          else copy.addIn(path, '');
          return node.flow || items.length === 0 ? copy.toString() : insertAfter(text, { value: items[items.length - 1] }, "- ''", true);
        },
      };
    }
    if (isScalar(node)) {
      const kind = kindOf(node.value);
      const span = scalarSpan(node);
      return {
        id,
        kind,
        label,
        at: node.range?.[0] ?? at,
        value: node.value === null ? (span && span.len > 0 ? text.slice(span.at, span.at + span.len).trim() : '') : String(node.value),
        ...(span ? { valueSpan: span } : {}),
      };
    }
    // empty value ("key:" with nothing after it)
    return { id, kind: 'null', label, at, value: '' };
  };

  /** Inserts a new line with the same indentation as the last item of a block collection. */
  function insertAfter(src: string, last: { key?: unknown; value?: unknown }, line: string, seq = false): string {
    const k = (last.key as Node | undefined)?.range?.[0];
    const v = last.value as Node | undefined;
    const startOfItem = seq ? (v?.range?.[0] ?? 0) : (k ?? 0);
    const lineStart = src.lastIndexOf('\n', startOfItem - 1) + 1;
    const prefix = /^[ \t]*(?:-[ \t]+)*/.exec(src.slice(lineStart, startOfItem))![0];
    // "  - key: v" -> new key at "    "; "  - item" -> new item at "  "
    const dash = prefix.lastIndexOf('-');
    const indent = seq ? (dash >= 0 ? prefix.slice(0, dash) : prefix) : prefix.replace(/-/g, ' ');
    const end0 = v?.range?.[2] ?? (last.key as Node | undefined)?.range?.[2] ?? startOfItem;
    const nl = src.indexOf('\n', Math.max(end0 - 1, startOfItem));
    const end = nl < 0 ? src.length : nl;
    return `${src.slice(0, end)}\n${indent}${line}${src.slice(end)}`;
  }

  const roots = list.map((d, i) =>
    build(d, d.contents, `d${i}`, list.length > 1 ? `Document ${i + 1}` : '(root)', [], d.contents?.range?.[0] ?? 0),
  );
  return { roots, problems };
}
