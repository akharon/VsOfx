// Turns the edited (contenteditable) Markdown rendering back into Markdown source.
// Only blocks the user changed are serialized (see codeView.ts); this covers what the
// renderer produces and what the browser's editing commands insert.

/** Markdown for one top-level block element (or a stray text node). */
export function blockToMarkdown(n: Node): string {
  if (n.nodeType === 3) return inline(n).trim();
  if (n.nodeType !== 1) return '';
  return block(n as HTMLElement).replace(/\n+$/, '');
}

function block(el: HTMLElement): string {
  const tag = el.tagName.toLowerCase();
  switch (tag) {
    case 'h1':
    case 'h2':
    case 'h3':
    case 'h4':
    case 'h5':
    case 'h6':
      return `${'#'.repeat(Number(tag[1]))} ${inlineChildren(el).replace(/\n/g, ' ').trim()}`;
    case 'ul':
    case 'ol':
      return list(el);
    case 'blockquote':
      return children(el)
        .split('\n')
        .map((l) => (l ? `> ${l}` : '>'))
        .join('\n');
    case 'pre': {
      const code = el.querySelector('code');
      const lang = /language-(\S+)/.exec(code?.className ?? '')?.[1] ?? '';
      const body = (el.textContent ?? '').replace(/\n$/, '');
      const fence = /```/.test(body) ? '~~~~' : '```';
      return `${fence}${lang}\n${body}\n${fence}`;
    }
    case 'hr':
      return '---';
    case 'table':
      return table(el);
    case 'br':
      return '';
    case 'div':
      // the browser wraps lines in <div> in some cases; a div with blocks inside is a container
      return Array.from(el.children).some((c) => isBlock(c)) ? children(el) : escapeLineStart(inlineChildren(el).trim());
    default:
      return escapeLineStart(inlineChildren(el).trim());
  }
}

const BLOCKS = new Set(['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'UL', 'OL', 'LI', 'BLOCKQUOTE', 'PRE', 'HR', 'TABLE']);
const isBlock = (n: Node) => n.nodeType === 1 && BLOCKS.has((n as Element).tagName);

/** Block children joined by blank lines; runs of inline content become paragraphs. */
function children(el: HTMLElement, sep = '\n\n'): string {
  const out: string[] = [];
  let run: Node[] = [];
  const flushRun = () => {
    const s = run.map(inline).join('').trim();
    if (s) out.push(s);
    run = [];
  };
  for (const c of Array.from(el.childNodes)) {
    if (isBlock(c)) {
      flushRun();
      const s = block(c as HTMLElement);
      if (s.trim()) out.push(s);
    } else run.push(c);
  }
  flushRun();
  return out.join(sep);
}

function list(el: HTMLElement): string {
  const ordered = el.tagName === 'OL';
  let n = Number(el.getAttribute('start') ?? 1) || 1;
  const bullet = el.dataset['markup'] && /^[-*+]$/.test(el.dataset['markup']) ? el.dataset['markup'] : '-';
  const delim = el.dataset['markup'] === ')' ? ')' : '.';
  const items: string[] = [];
  for (const li of Array.from(el.children)) {
    if (li.tagName !== 'LI') {
      // nested list typed directly into a list (browser indent command)
      if (li.tagName === 'UL' || li.tagName === 'OL') {
        const nested = list(li as HTMLElement).replace(/^/gm, '   ');
        items.push(nested);
      }
      continue;
    }
    const marker = ordered ? `${n++}${delim} ` : `${bullet} `;
    const box = li.querySelector(':scope > input.task, :scope > p > input.task') as HTMLInputElement | null;
    const task = box ? (box.hasAttribute('checked') || box.checked ? '[x] ' : '[ ] ') : '';
    // a tight item (no <p>) keeps its text and nested list on consecutive lines, else the list turns loose
    const tight = !li.querySelector(':scope > p');
    const body = children(li as HTMLElement, tight ? '\n' : '\n\n').replace(/^\s+/, '');
    const pad = ' '.repeat(marker.length);
    const lines = body.split('\n');
    items.push(marker + task + lines[0] + lines.slice(1).map((l) => '\n' + (l ? pad + l : '')).join(''));
  }
  return items.join('\n');
}

function table(el: HTMLElement): string {
  const rows = Array.from(el.querySelectorAll('tr'));
  if (!rows.length) return '';
  const cells = rows.map((tr) => Array.from(tr.children).map((c) => inline(c).replace(/\n/g, ' ').replace(/\|/g, '\\|').trim()));
  const cols = Math.max(...cells.map((r) => r.length));
  const aligns = Array.from(rows[0]!.children).map((c) => (c as HTMLElement).style.textAlign);
  const line = (r: string[]) => `| ${Array.from({ length: cols }, (_, i) => r[i] ?? '').join(' | ')} |`;
  const sep = `| ${Array.from({ length: cols }, (_, i) => {
    const a = aligns[i];
    return a === 'center' ? ':---:' : a === 'right' ? '---:' : a === 'left' ? ':---' : '---';
  }).join(' | ')} |`;
  return [line(cells[0]!), sep, ...cells.slice(1).map(line)].join('\n');
}

// ---- inline --------------------------------------------------------------------------------

function inline(n: Node): string {
  if (n.nodeType === 3) return escape((n.nodeValue ?? '').replace(/ /g, ' ').replace(/\s*\n\s*/g, ' '));
  if (n.nodeType !== 1) return '';
  const el = n as HTMLElement;
  const tag = el.tagName.toLowerCase();
  const inner = () => Array.from(el.childNodes).map(inline).join('');
  switch (tag) {
    case 'strong':
    case 'b':
      return wrap(inner(), '**');
    case 'em':
    case 'i':
      return wrap(inner(), '_');
    case 's':
    case 'del':
    case 'strike':
      return wrap(inner(), '~~');
    case 'code': {
      const t = el.textContent ?? '';
      if (!t) return '';
      const tick = t.includes('`') ? '``' : '`';
      return `${tick}${tick.length > 1 ? ' ' : ''}${t}${tick.length > 1 ? ' ' : ''}${tick}`;
    }
    case 'a': {
      const href = el.getAttribute('href') ?? '';
      const text = inner();
      if (!href) return text;
      return `[${text || href}](${href.replace(/[()\s]/g, (c) => encodeURIComponent(c))})`;
    }
    case 'img':
      return `![${el.getAttribute('alt') ?? ''}](${el.getAttribute('src') ?? ''})`;
    case 'br':
      return '\\\n';
    case 'input':
      return ''; // task checkboxes are written by the list serializer
    case 'span':
      if (el.classList.contains('img-alt')) return `![${el.dataset['alt'] ?? ''}](${el.dataset['src'] ?? ''})`;
      return inner();
    case 'mark':
    case 'u':
    case 'font':
    default:
      if (isBlock(el)) return block(el);
      return inner();
  }
}

/** Wraps in a delimiter, keeping surrounding spaces outside (else it is not Markdown). */
function wrap(s: string, mark: string): string {
  const m = /^(\s*)([\s\S]*?)(\s*)$/.exec(s)!;
  if (!m[2]) return s;
  return `${m[1]}${mark}${m[2]}${mark}${m[3]}`;
}

function escape(s: string): string {
  return s
    .replace(/([\\`*[\]<])/g, '\\$1')
    .replace(/(^|[^\p{L}\p{N}])_|_(?=$|[^\p{L}\p{N}])/gu, (m) => m.replace('_', '\\_'));
}

/** Escapes what would turn the start of a paragraph into another block (heading, list ...). */
export function escapeLineStart(s: string): string {
  return s.replace(/^(\s*)(#{1,6}\s|[-+]\s|\d+[.)]\s|>|={3,}|-{3,})/, (_m, sp: string, x: string) => `${sp}\\${x}`);
}

/** Inline Markdown of an element's content (not of the element itself: that would recurse for blocks). */
function inlineChildren(el: Element): string {
  return Array.from(el.childNodes).map(inline).join('');
}
