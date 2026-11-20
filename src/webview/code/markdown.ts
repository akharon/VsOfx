import MarkdownIt from 'markdown-it';
import { attachDiagrams } from './diagrams';

// Raw HTML is disabled: everything in the source is escaped, and markdown-it's link
// validation drops javascript:/vbscript:/file: URLs. Remote images cannot load under the
// webview CSP, so they are replaced by their alt text instead of producing blocked requests.

const md = new MarkdownIt({ html: false, linkify: true, typographer: false, breaks: false });

// Aligned paragraphs: the one piece of HTML that GitHub and most renderers accept for alignment,
// `<p align="center">text</p>` on one line. Only this exact form is recognised (raw HTML stays off).
const ALIGNED = /^<p align="(left|center|right|justify)">(.*)<\/p>\s*$/;
md.block.ruler.before('paragraph', 'aligned_paragraph', (state, startLine, _endLine, silent) => {
  if (state.sCount[startLine]! - state.blkIndent >= 4) return false;
  const line = state.src.slice(state.bMarks[startLine]! + state.tShift[startLine]!, state.eMarks[startLine]);
  const m = ALIGNED.exec(line);
  if (!m) return false;
  if (silent) return true;
  const open = state.push('paragraph_open', 'p', 1);
  open.map = [startLine, startLine + 1];
  open.attrSet('style', `text-align: ${m[1]}`);
  const inl = state.push('inline', '', 0);
  inl.content = m[2]!.trim();
  inl.map = [startLine, startLine + 1];
  inl.children = [];
  state.push('paragraph_close', 'p', -1);
  state.line = startLine + 1;
  return true;
});

// source line of each block, so the preview can jump to the source
md.core.ruler.push('source_lines', (state) => {
  for (const t of state.tokens) {
    if (t.map && t.nesting >= 0 && t.block) {
      t.attrSet('data-line', String(t.map[0]));
      if (t.level === 0) t.attrSet('data-end', String(t.map[1]));
    }
    if ((t.type === 'bullet_list_open' || t.type === 'ordered_list_open') && t.markup) t.attrSet('data-markup', t.markup);
  }
  return true;
});

// GitHub task lists: "- [ ] todo" / "- [x] done"
md.core.ruler.push('task_lists', (state) => {
  const toks = state.tokens;
  for (let i = 2; i < toks.length; i++) {
    const inline = toks[i]!;
    if (inline.type !== 'inline' || toks[i - 1]!.type !== 'paragraph_open' || toks[i - 2]!.type !== 'list_item_open') continue;
    const m = /^\[([ xX])\]\s/.exec(inline.content);
    const first = inline.children?.[0];
    if (!m || !first || first.type !== 'text') continue;
    first.content = first.content.slice(m[0].length);
    const box = new state.Token('html_inline', '', 0);
    box.content = `<input type="checkbox" class="task" data-line="${toks[i - 2]!.map?.[0] ?? 0}"${m[1] === ' ' ? '' : ' checked'}> `;
    inline.children!.unshift(box);
    toks[i - 2]!.attrJoin('class', 'task-item');
  }
  return true;
});
// html_inline is only produced by the rule above (html: false escapes user HTML)
md.renderer.rules['html_inline'] = (tokens, idx) => tokens[idx]!.content;

/** Renders Markdown into `target` (replacing its content). */
export function renderMarkdown(source: string, target: HTMLElement): void {
  const html = md.render(source);
  const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
  for (const img of Array.from(doc.querySelectorAll('img'))) {
    const src = img.getAttribute('src') ?? '';
    if (!src.startsWith('data:image/')) {
      const alt = doc.createElement('span');
      alt.className = 'img-alt';
      alt.textContent = `🖼 ${img.getAttribute('alt') || src}`;
      alt.title = src;
      alt.dataset['src'] = src;
      alt.dataset['alt'] = img.getAttribute('alt') ?? '';
      alt.setAttribute('contenteditable', 'false');
      img.replaceWith(alt);
    }
  }
  for (const a of Array.from(doc.querySelectorAll('a'))) a.setAttribute('title', a.getAttribute('href') ?? '');
  // fenced code: markdown-it puts the attributes on <code>; the block is the <pre>
  for (const pre of Array.from(doc.querySelectorAll('body > pre'))) {
    const code = pre.querySelector('code');
    for (const k of ['data-line', 'data-end']) {
      const v = code?.getAttribute(k);
      if (v !== null && v !== undefined) {
        pre.setAttribute(k, v);
        code!.removeAttribute(k);
      }
    }
  }
  target.replaceChildren(...Array.from(doc.body.childNodes).map((n) => document.importNode(n, true)));
  attachDiagrams(target);
}

// ---- formatting commands on the source -------------------------------------------------------

export interface Selection {
  text: string;
  start: number;
  end: number;
}

/** Wraps the selection in `mark` (toggles if already wrapped). Returns new text and selection. */
export function toggleWrap(s: Selection, mark: string, placeholder = 'text'): Selection {
  const { text, start, end } = s;
  const before = text.slice(start - mark.length, start);
  const after = text.slice(end, end + mark.length);
  if (before === mark && after === mark) {
    return { text: text.slice(0, start - mark.length) + text.slice(start, end) + text.slice(end + mark.length), start: start - mark.length, end: end - mark.length };
  }
  const inner = text.slice(start, end) || placeholder;
  return { text: text.slice(0, start) + mark + inner + mark + text.slice(end), start: start + mark.length, end: start + mark.length + inner.length };
}

/** Applies `fn` to every line touched by the selection. */
export function mapLines(s: Selection, fn: (line: string, i: number) => string): Selection {
  const ls = s.text.lastIndexOf('\n', s.start - 1) + 1;
  let le = s.text.indexOf('\n', Math.max(s.end - (s.end > s.start ? 1 : 0), s.start));
  if (le < 0) le = s.text.length;
  const block = s.text.slice(ls, le).split('\n').map(fn).join('\n');
  return { text: s.text.slice(0, ls) + block + s.text.slice(le), start: ls, end: ls + block.length };
}

/** Sets heading level 1-6 on the selected lines (0 = plain paragraph, same level again = toggle off). */
export function setHeading(s: Selection, level: number): Selection {
  return mapLines(s, (l) => {
    const m = /^(#{1,6})\s+/.exec(l);
    const body = m ? l.slice(m[0].length) : l;
    if (level === 0 || (m && m[1]!.length === level)) return body;
    return `${'#'.repeat(level)} ${body}`;
  });
}

/** Toggles a line prefix such as "- ", "1. ", "> ". */
export function togglePrefix(s: Selection, prefix: string | ((i: number) => string), pattern: RegExp): Selection {
  const lines = s.text.slice(s.text.lastIndexOf('\n', s.start - 1) + 1, s.end).split('\n');
  const all = lines.every((l) => pattern.test(l) || l.trim() === '');
  return mapLines(s, (l, i) => {
    if (all) return l.replace(pattern, '');
    if (l.trim() === '') return l;
    return (typeof prefix === 'string' ? prefix : prefix(i)) + l.replace(pattern, '');
  });
}
