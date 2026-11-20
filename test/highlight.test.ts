import { describe, expect, it } from 'vitest';
import { highlightLang, tokenize } from '../src/webview/code/highlight';

const SAMPLES: Record<string, string> = {
  python: '@dec\ndef f(x: int) -> str:\n    """doc"""\n    return f"{x}" # c\n',
  c: '#include <stdio.h>\nint main(void) { /* c */ printf("%d\n", 0x1F); return 0; }\n',
  cpp: 'template <typename T> class A { std::vector<T> v; }; // x\n',
  java: '@Override public String toString() { return "a" + 1.5e3; }',
  rust: "fn f<'a>(s: &'a str) -> Option<u8> { println!(\"{}\", 'x'); None }",
  javascript: 'const t = `a ${b}`; // c\nclass X extends Y {}',
  typescript: 'interface P { a: number } type Q = keyof P;',
  dart: "void main() { final s = '''multi\nline'''; print(s); }",
  ruby: "def ok?\n  :sym # c\nend\n=begin\nx\n=end\n",
  gdscript: 'extends Node2D\n@export var speed: float = 1.0\nfunc _ready():\n\tpass',
  json: '{ "a": [1, true, null], "b": "s" // c\n}',
  yaml: '# c\nkey: value\nlist:\n  - 1\n  - "q" # c\nblock: |\n  text\n  more\nnext: &a yes\n',
  xml: '<?xml version="1.0"?>\n<!-- c --><a x="1"><b/>t&amp;<![CDATA[<x>]]></a>',
};

describe('code highlighter', () => {
  it('knows the supported languages and aliases, and nothing else', () => {
    expect(highlightLang('py')).toBe('python');
    expect(highlightLang('C++')).toBe('cpp');
    expect(highlightLang('yml')).toBe('yaml');
    expect(highlightLang('gd')).toBe('gdscript');
    expect(highlightLang('go')).toBeUndefined();
    expect(highlightLang('html')).toBeUndefined();
  });

  for (const [lang, src] of Object.entries(SAMPLES)) {
    it(`keeps the ${lang} text unchanged and marks tokens`, () => {
      const toks = tokenize(src, lang);
      expect(toks.map((t) => t[1]).join('')).toBe(src);
      expect(toks.some((t) => t[0] !== null)).toBe(true);
    });
  }

  it('marks keywords, strings and comments', () => {
    const t = tokenize('def f(): return "x" # c', 'python');
    expect(t).toContainEqual(['kw', 'def']);
    expect(t).toContainEqual(['str', '"x"']);
    expect(t).toContainEqual(['com', '# c']);
    expect(tokenize('{"k": 1}', 'json')).toContainEqual(['key', '"k"']);
  });
});
