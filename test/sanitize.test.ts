import { describe, expect, it } from 'vitest';
import { sanitizeHtml } from '../src/webview/sanitize.ts';

describe('sanitizer', () => {
  it('keeps allowed formatting', () => {
    expect(sanitizeHtml('<h1>T</h1><p>a <strong>b</strong> <em>c</em></p><ul><li>x</li></ul>')).toBe(
      '<h1>T</h1><p>a <strong>b</strong> <em>c</em></p><ul><li>x</li></ul>',
    );
  });

  it('drops scripts together with their content', () => {
    const out = sanitizeHtml('<p>ok<script>alert(1)</script></p>');
    expect(out).toBe('<p>ok</p>');
  });

  it('removes event handlers and style attributes', () => {
    const out = sanitizeHtml('<p onclick="x()" style="color:red" id="a" class="b">t</p><img src="data:image/png;base64,AAAA" onerror="x()"/>');
    expect(out).not.toMatch(/onclick|onerror|style=|id=|class=/);
    expect(out).toContain('<p>t</p>');
  });

  it('rejects javascript: and external links but keeps the text', () => {
    expect(sanitizeHtml('<p><a href="javascript:alert(1)">click</a></p>')).toBe('<p>click</p>');
    expect(sanitizeHtml('<p><a href="https://evil.example/x">site</a></p>')).toBe('<p>site</p>');
    expect(sanitizeHtml('<p><a href="#bookmark">jump</a></p>')).toBe('<p><a href="#bookmark">jump</a></p>');
  });

  it('only allows data: images', () => {
    expect(sanitizeHtml('<p><img src="https://evil.example/a.png"/></p>')).toBe('<p></p>');
    expect(sanitizeHtml('<p><img src="//evil.example/a.png"/></p>')).toBe('<p></p>');
    expect(sanitizeHtml('<p><img src="data:text/html;base64,PHNjcmlwdD4="/></p>')).toBe('<p></p>');
    expect(sanitizeHtml('<p><img src="data:image/svg+xml;base64,AAAA"/></p>')).toBe('<p></p>');
    expect(sanitizeHtml('<p><img src="data:image/png;base64,AAAA" alt="x"/></p>')).toBe('<p><img src="data:image/png;base64,AAAA" alt="x"></p>');
  });

  it('drops dangerous containers and unwraps unknown tags', () => {
    expect(sanitizeHtml('<p>a</p><iframe src="x">z</iframe><object>z</object><svg><script/></svg><form><input/></form>')).toBe('<p>a</p>');
    expect(sanitizeHtml('<p><font color="red">text</font></p>')).toBe('<p>text</p>');
  });

  it('escapes text', () => {
    expect(sanitizeHtml('<p>a &lt;b&gt; &amp; "q"</p>')).toBe('<p>a &lt;b&gt; &amp; &quot;q&quot;</p>');
  });

  it('survives non-XML input via the lenient parser', () => {
    const out = sanitizeHtml('<p>one<br>two<script>x</script>');
    expect(out).not.toContain('script');
    expect(out).toContain('one');
  });

  it('keeps table spans as numbers only', () => {
    expect(sanitizeHtml('<table><tr><td colspan="2" onclick="x">c</td><td rowspan="a2">d</td></tr></table>')).toBe(
      '<table><tr><td colspan="2">c</td><td>d</td></tr></table>',
    );
  });
});
