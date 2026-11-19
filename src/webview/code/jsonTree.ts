import { applyEdits, format, getNodePath, modify, parseTree, printParseErrorCode, type FormattingOptions, type Node, type ParseError } from 'jsonc-parser';
import type { NodeKind, ParsedTree, TreeNode } from './tree';

/** Indentation used by the document (first indented line), for inserts and "Format". */
export function detectIndent(text: string): FormattingOptions {
  const m = /^\n?([ \t]+)\S/m.exec(text.slice(0, 10_000));
  const ws = m?.[1] ?? '  ';
  return ws.startsWith('\t') ? { insertSpaces: false, tabSize: 4, eol: '\n' } : { insertSpaces: true, tabSize: Math.min(8, ws.length) || 2, eol: '\n' };
}

/** Pretty-prints JSON (comments kept) with the document's own indentation. */
export function formatJson(text: string): string {
  return applyEdits(text, format(text, undefined, detectIndent(text)));
}

/** A typed value: strings are always quoted; other kinds accept any JSON literal, else become a string. */
function encodeValue(kind: NodeKind, s: string): string {
  if (kind !== 'string') {
    const t = s.trim();
    try {
      JSON.parse(t);
      return t;
    } catch {
      /* not a literal: store as string */
    }
  }
  return JSON.stringify(s);
}

export function jsonTree(text: string): ParsedTree {
  const errors: ParseError[] = [];
  const root = parseTree(text, errors, { allowTrailingComma: true, disallowComments: false, allowEmptyContent: true });
  const problems = errors.map((e) => ({ message: printParseErrorCode(e.error).replace(/([a-z])([A-Z])/g, '$1 $2'), offset: e.offset }));
  const opts = detectIndent(text);

  const remove = (n: Node) => () => applyEdits(text, modify(text, getNodePath(n), undefined, { formattingOptions: opts }));

  const build = (n: Node, id: string, label: string, keyNode?: Node): TreeNode => {
    const base = {
      id,
      label,
      at: keyNode?.offset ?? n.offset,
      ...(keyNode
        ? { labelSpan: { at: keyNode.offset, len: keyNode.length, encode: (s: string) => JSON.stringify(s) } }
        : {}),
      ...(n.parent ? { remove: remove(n) } : {}),
    };
    if (n.type === 'object') {
      const props = n.children ?? [];
      return {
        ...base,
        kind: 'object',
        summary: `${props.length} ${props.length === 1 ? 'key' : 'keys'}`,
        children: props.flatMap((p, i) => {
          const [k, v] = p.children ?? [];
          if (!k) return [];
          const name = String(k.value);
          // a property without a value (while typing) is shown as null
          return [v ? build(v, `${id}/${i}`, name, k) : { id: `${id}/${i}`, kind: 'null' as const, label: name, at: k.offset }];
        }),
        addLabel: 'Add key',
        add: () => {
          const keys = new Set(props.map((p) => String(p.children?.[0]?.value)));
          let key = 'key';
          for (let i = 2; keys.has(key); i++) key = `key${i}`;
          return applyEdits(text, modify(text, [...getNodePath(n), key], '', { formattingOptions: opts }));
        },
      };
    }
    if (n.type === 'array') {
      const items = n.children ?? [];
      return {
        ...base,
        kind: 'array',
        summary: `${items.length} ${items.length === 1 ? 'item' : 'items'}`,
        children: items.map((c, i) => build(c, `${id}/${i}`, String(i))),
        addLabel: 'Add item',
        add: () => applyEdits(text, modify(text, [...getNodePath(n), -1], '', { formattingOptions: opts, isArrayInsertion: true })),
      };
    }
    const kind: NodeKind = n.type === 'string' ? 'string' : n.type === 'number' ? 'number' : n.type === 'boolean' ? 'boolean' : n.type === 'null' ? 'null' : 'other';
    return {
      ...base,
      kind,
      value: kind === 'string' ? String(n.value) : text.slice(n.offset, n.offset + n.length),
      valueSpan: { at: n.offset, len: n.length, encode: (s) => encodeValue(kind, s) },
    };
  };

  return { roots: root ? [build(root, 'r', '(root)')] : [], problems };
}
