// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { applyTextEdit, inverseEdit, type Edit, type TextEdit } from '../src/shared/model.ts';
import { renderMarkdown } from '../src/webview/code/markdown.ts';
import { blockToMarkdown } from '../src/webview/code/mdSerialize.ts';
import { createToolbar } from '../src/webview/ui/toolbar.ts';
import { renderCodeView } from '../src/webview/views/codeView.ts';

const SAMPLES: Record<string, string> = {
  headings: '# One\n\n## Two *em*\n\n###### Six',
  inline: 'Text with **bold**, _italic_, ~~strike~~, `code` and [a link](https://e.com).',
  bullets: '- a\n- b\n  - nested\n- c',
  numbers: '1. first\n2. second\n3. third',
  tasks: '- [ ] open\n- [x] done',
  quote: '> quoted **text**\n>\n> second paragraph',
  fence: '```ts\nconst a = 1;\n```',
  table: '| A | B |\n| --- | --- |\n| 1 | 2 |',
  rule: '---',
  breaks: 'line one\\\nline two',
  special: 'Not a heading: \\# hash, 1\\. not a list',
};

/** HTML of Markdown text, for comparing what two sources mean. */
function html(md: string): string {
  const el = document.createElement('div');
  renderMarkdown(md, el);
  for (const e of Array.from(el.querySelectorAll('[data-line], [data-end]'))) {
    e.removeAttribute('data-line');
    e.removeAttribute('data-end');
  }
  return el.innerHTML.replace(/\s+/g, ' ').trim();
}

describe('Markdown: rendered blocks serialize back to the same document', () => {
  for (const [name, md] of Object.entries(SAMPLES)) {
    it(name, () => {
      const el = document.createElement('div');
      renderMarkdown(md, el);
      const out = Array.from(el.children)
        .map((c) => blockToMarkdown(c))
        .filter((s) => s !== '')
        .join('\n\n');
      expect(html(out)).toBe(html(md));
    });
  }
});

describe('Markdown editor: undo of visual edits', () => {
  it('host undo/redo of a visual edit restores the text and the view', () => {
    document.body.innerHTML = '<div id="root"></div>';
    localStorage.removeItem('vsofx.mode.markdown');
    const root = document.getElementById('root')!;
    const edits: Edit[] = [];
    const src = '# Title\n\nBody text.\n';
    const view = renderCodeView(root, { type: 'code', lang: 'markdown', text: src }, false, (e) => edits.push(e), createToolbar());
    root.querySelector('.md-body p')!.textContent = 'Changed body.';
    view.flush!();
    expect(edits).toHaveLength(1);
    const after = applyTextEdit(src, edits[0] as TextEdit);
    expect(after).toBe('# Title\n\nChanged body.\n');
    view.applyEdit(edits[0]!, {}); // own edit echoed back: ignored
    view.applyEdit(inverseEdit(edits[0]!), {}); // undo
    expect(root.querySelector('.md-body p')!.textContent).toBe('Body text.');
    view.flush!();
    expect(edits).toHaveLength(1); // nothing new to commit after an undo
    view.applyEdit(edits[0]!, {}); // redo (a host edit, not an echo any more)
    expect(root.querySelector('.md-body p')!.textContent).toBe('Changed body.');
    view.dispose();
  });
});
