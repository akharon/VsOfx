// Structured view of a text document (JSON, YAML, XML). Every editable part knows the
// exact span of source text it came from, so an edit is a minimal splice: formatting,
// comments and the order of everything else stay as the user wrote them.

export type NodeKind = 'object' | 'array' | 'element' | 'string' | 'number' | 'boolean' | 'null' | 'text' | 'attr' | 'other';

/** A replaceable piece of source text. `encode` turns what the user typed into source text. */
export interface Span {
  at: number;
  len: number;
  encode(value: string): string;
}

export interface TreeNode {
  /** Stable id (path), used to keep expanded state and focus across re-renders. */
  id: string;
  kind: NodeKind;
  /** Key, index, tag or attribute name. */
  label: string;
  /** If set, the label can be renamed. */
  labelSpan?: Span;
  /** Leaf value as the user sees and edits it. */
  value?: string;
  valueSpan?: Span;
  children?: TreeNode[];
  /** Short summary for containers ("3 keys", "5 items"). */
  summary?: string;
  /** Source offset of the node (for "reveal in source"). */
  at: number;
  /** Returns the full new text with this node removed, or null if not possible. */
  remove?: () => string | null;
  /** Returns the full new text with a new child added (containers). */
  add?: () => string | null;
  /** Label for the add button ("Add key", "Add item"). */
  addLabel?: string;
}

export interface Problem {
  message: string;
  offset: number;
}

export interface ParsedTree {
  roots: TreeNode[];
  problems: Problem[];
}

/** Line (1-based) of an offset. */
export function lineOf(text: string, offset: number): number {
  let n = 1;
  for (let i = text.indexOf('\n'); i >= 0 && i < offset; i = text.indexOf('\n', i + 1)) n++;
  return n;
}

/**
 * Widens [start, end) to whole lines when only whitespace (and `prefix` chars such as "- ")
 * surround it, so removing a node does not leave an empty line behind.
 */
export function wholeLines(text: string, start: number, end: number): [number, number] {
  // some parsers include trailing line breaks in a node's range
  while (end > start && text[end - 1] === '\n') end--;
  let s = start;
  while (s > 0 && (text[s - 1] === ' ' || text[s - 1] === '\t')) s--;
  if (s > 0 && text[s - 1] !== '\n') return [start, end];
  let e = end;
  while (e < text.length && (text[e] === ' ' || text[e] === '\t')) e++;
  if (e < text.length && text[e] !== '\n') return [start, end];
  return [s, e < text.length ? e + 1 : e];
}

/** Applies a span edit and returns the new text. */
export function spliceSpan(text: string, span: Span, value: string): string {
  return text.slice(0, span.at) + span.encode(value) + text.slice(span.at + span.len);
}
