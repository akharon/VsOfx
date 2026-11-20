// Small syntax highlighter for fenced code in the Markdown view. Only these languages:
// code - Python, C, C++, Java, Rust, JavaScript, TypeScript, Dart, Ruby, GDScript;
// config - XML, JSON, YAML. Tokens become <span class="tk-…">; the text is unchanged, so the
// serializer (which reads textContent) is not affected.

type Kind = 'kw' | 'type' | 'str' | 'num' | 'com' | 'fn' | 'lit' | 'meta' | 'key' | 'tag' | 'attr' | 'punct';
type Tok = [Kind | null, string];

interface CodeLang {
  kind: 'code';
  name: string;
  keywords: Set<string>;
  types: Set<string>;
  literals: Set<string>;
  line: string[]; // line comment starts
  block?: [string, string];
  hashMeta?: boolean; // C/C++ preprocessor lines
  decorator?: boolean; // @name
  rawStr?: boolean; // python/ruby/dart triple quotes, backticks
  symbols?: boolean; // ruby :sym
  lifetimes?: boolean; // rust 'a
}

const w = (s: string) => new Set(s.split(/\s+/).filter(Boolean));

const C_KW = 'auto break case const continue default do else enum extern for goto if inline register restrict return sizeof static struct switch typedef union volatile while _Alignas _Alignof _Atomic _Bool _Generic _Noreturn _Static_assert _Thread_local';
const C_TYPES = 'void char short int long float double signed unsigned bool size_t ssize_t int8_t int16_t int32_t int64_t uint8_t uint16_t uint32_t uint64_t uintptr_t intptr_t ptrdiff_t FILE wchar_t';
const JS_KW = 'async await break case catch class const continue debugger default delete do else export extends finally for from function get if import in instanceof let new of return set static super switch this throw try typeof var void while with yield as';

const LANGS: Record<string, CodeLang> = {
  python: {
    kind: 'code', name: 'Python',
    keywords: w('and as assert async await break class continue def del elif else except finally for from global if import in is lambda nonlocal not or pass raise return try while with yield match case self cls'),
    types: w('int float str bool bytes list dict set tuple object type complex frozenset bytearray Exception'),
    literals: w('True False None'),
    line: ['#'], decorator: true, rawStr: true,
  },
  c: {
    kind: 'code', name: 'C', keywords: w(C_KW), types: w(C_TYPES), literals: w('NULL true false'),
    line: ['//'], block: ['/*', '*/'], hashMeta: true,
  },
  cpp: {
    kind: 'code', name: 'C++',
    keywords: w(`${C_KW} alignas alignof catch class concept consteval constexpr constinit const_cast co_await co_return co_yield decltype delete dynamic_cast explicit export friend mutable namespace new noexcept operator override final private protected public reinterpret_cast requires static_assert static_cast template this thread_local throw try typeid typename using virtual`),
    types: w(`${C_TYPES} std string vector map unordered_map set array unique_ptr shared_ptr optional auto char8_t char16_t char32_t`),
    literals: w('nullptr NULL true false'),
    line: ['//'], block: ['/*', '*/'], hashMeta: true,
  },
  java: {
    kind: 'code', name: 'Java',
    keywords: w('abstract assert break case catch class continue default do else enum extends final finally for if implements import instanceof interface native new package private protected public record return sealed permits static strictfp super switch synchronized this throw throws transient try var volatile while yield'),
    types: w('void boolean byte char short int long float double String Object Integer Long Double Boolean List Map Set'),
    literals: w('true false null'),
    line: ['//'], block: ['/*', '*/'], decorator: true,
  },
  rust: {
    kind: 'code', name: 'Rust',
    keywords: w('as async await break const continue crate dyn else enum extern fn for if impl in let loop match mod move mut pub ref return self Self static struct super trait type unsafe use where while macro_rules'),
    types: w('i8 i16 i32 i64 i128 isize u8 u16 u32 u64 u128 usize f32 f64 bool char str String Vec Option Result Box Rc Arc HashMap HashSet'),
    literals: w('true false None Some Ok Err'),
    line: ['//'], block: ['/*', '*/'], lifetimes: true,
  },
  javascript: {
    kind: 'code', name: 'JavaScript', keywords: w(JS_KW), types: w('Array Object String Number Boolean Promise Map Set Date RegExp Error JSON Math Symbol BigInt'),
    literals: w('true false null undefined NaN Infinity'), line: ['//'], block: ['/*', '*/'], rawStr: true, decorator: true,
  },
  typescript: {
    kind: 'code', name: 'TypeScript',
    keywords: w(`${JS_KW} abstract declare enum implements interface keyof namespace private protected public readonly type satisfies infer is asserts override`),
    types: w('string number boolean any unknown never void object bigint symbol Array Record Partial Readonly Promise Map Set Date Error'),
    literals: w('true false null undefined NaN Infinity'), line: ['//'], block: ['/*', '*/'], rawStr: true, decorator: true,
  },
  dart: {
    kind: 'code', name: 'Dart',
    keywords: w('abstract as assert async await base break case catch class const continue covariant default deferred do dynamic else enum export extends extension external factory final finally for get hide if implements import in interface is late library mixin new on operator part required rethrow return sealed set show static super switch sync this throw try typedef var when while with yield'),
    types: w('void int double num bool String List Map Set Future Stream Object dynamic Iterable Function Never'),
    literals: w('true false null'), line: ['//'], block: ['/*', '*/'], decorator: true, rawStr: true,
  },
  ruby: {
    kind: 'code', name: 'Ruby',
    keywords: w('alias and begin break case class def defined? do else elsif end ensure for if in module next not or redo rescue retry return self super then undef unless until when while yield require require_relative attr_accessor attr_reader attr_writer private protected public lambda proc puts'),
    types: w('String Integer Float Array Hash Symbol Proc Object Struct Kernel'),
    literals: w('true false nil'), line: ['#'], block: ['=begin', '=end'], symbols: true,
  },
  gdscript: {
    kind: 'code', name: 'GDScript',
    keywords: w('and as assert await break breakpoint class class_name const continue elif else enum extends for func if in is match not or pass preload return self signal static super tool var void while yield onready export setget get set'),
    types: w('int float bool String StringName NodePath Vector2 Vector2i Vector3 Vector3i Vector4 Color Rect2 Transform2D Transform3D Basis Quaternion Array Dictionary Node Node2D Node3D Object Resource PackedScene Callable Signal'),
    literals: w('true false null PI TAU INF NAN'), line: ['#'], decorator: true, rawStr: true,
  },
};

const ALIASES: Record<string, string> = {
  py: 'python', python3: 'python', h: 'c', 'c++': 'cpp', cc: 'cpp', cxx: 'cpp', hpp: 'cpp', hh: 'cpp',
  js: 'javascript', jsx: 'javascript', mjs: 'javascript', cjs: 'javascript', node: 'javascript',
  ts: 'typescript', tsx: 'typescript', mts: 'typescript', rs: 'rust', rb: 'ruby',
  gd: 'gdscript', godot: 'gdscript', gdscript: 'gdscript', jsonc: 'json', json5: 'json', yml: 'yaml',
};

const CONFIG = new Set(['json', 'yaml', 'xml']);
const CONFIG_NAMES: Record<string, string> = { json: 'JSON', yaml: 'YAML', xml: 'XML' };

/** Canonical language for a fence tag, or undefined when it is not one we highlight. */
export function highlightLang(tag: string): string | undefined {
  const t = tag.toLowerCase();
  const k = ALIASES[t] ?? t;
  return LANGS[k] || CONFIG.has(k) ? k : undefined;
}

export function langLabel(lang: string): string {
  return LANGS[lang]?.name ?? CONFIG_NAMES[lang] ?? lang;
}

// ---- tokenizers --------------------------------------------------------------------------------

const NUM = /^(0[xX][0-9a-fA-F_]+|0[bB][01_]+|0[oO][0-7_]+|(\d[\d_]*\.?[\d_]*|\.\d[\d_]*)([eE][+-]?\d+)?)[a-zA-Z_]*/;
const IDENT = /^[A-Za-z_$][\w$]*[?!]?/;

function tokCode(src: string, L: CodeLang): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  let lineStart = true;
  let prevWord = '';
  const push = (k: Kind | null, s: string) => {
    if (!s) return;
    const last = out[out.length - 1];
    if (last && last[0] === k) last[1] += s;
    else out.push([k, s]);
  };
  while (i < src.length) {
    const rest = src.slice(i);
    const ch = src[i]!;
    if (ch === '\n') {
      push(null, ch);
      i++;
      lineStart = true;
      continue;
    }
    if (ch === ' ' || ch === '\t' || ch === '\r') {
      push(null, ch);
      i++;
      continue;
    }
    const atLineStart = lineStart;
    lineStart = false;
    // comments
    if (L.block && rest.startsWith(L.block[0]) && (L.block[0] !== '=begin' || atLineStart)) {
      const e = src.indexOf(L.block[1], i + L.block[0].length);
      const end = e < 0 ? src.length : e + L.block[1].length;
      push('com', src.slice(i, end));
      i = end;
      continue;
    }
    if (L.line.some((c) => rest.startsWith(c))) {
      const e = src.indexOf('\n', i);
      const end = e < 0 ? src.length : e;
      push('com', src.slice(i, end));
      i = end;
      continue;
    }
    if (L.hashMeta && ch === '#' && atLineStart) {
      let end = i;
      // continue over backslash line ends
      for (;;) {
        const e = src.indexOf('\n', end);
        if (e < 0) {
          end = src.length;
          break;
        }
        end = e;
        if (src[e - 1] !== '\\') break;
        end = e + 1;
      }
      push('meta', src.slice(i, end));
      i = end;
      continue;
    }
    // strings
    const pre = /^(?:[rbfuRBFU]{1,2}|r#*)?(?=["'`])/.exec(rest)?.[0] ?? '';
    const q = src[i + pre.length];
    if ((q === '"' || q === "'" || (q === '`' && L.rawStr)) && !(q === "'" && L.lifetimes && /^'[A-Za-z_]\w*(?!')/.test(rest) && !/^'.'/.test(rest))) {
      if (pre && !/^[rbfuRBFU]+$|^r#*$/.test(pre)) {
        // not a string prefix
      } else {
        const start = i + pre.length;
        const triple = L.rawStr && src.startsWith(q.repeat(3), start) && q !== '`';
        const close = triple ? q.repeat(3) : q;
        let j = start + close.length;
        const multi = triple || q === '`';
        while (j < src.length) {
          if (src[j] === '\\') {
            j += 2;
            continue;
          }
          if (src.startsWith(close, j)) {
            j += close.length;
            break;
          }
          if (src[j] === '\n' && !multi) break;
          j++;
        }
        push('str', src.slice(i, j));
        i = j;
        continue;
      }
    }
    if (L.lifetimes && ch === "'") {
      const m = /^'[A-Za-z_]\w*/.exec(rest);
      if (m) {
        push('meta', m[0]);
        i += m[0].length;
        continue;
      }
    }
    if (L.decorator && ch === '@') {
      const m = /^@[\w.]+/.exec(rest);
      if (m) {
        push('meta', m[0]);
        i += m[0].length;
        continue;
      }
    }
    if (L.symbols && ch === ':' && /^:[A-Za-z_]/.test(rest) && src[i - 1] !== ':') {
      const m = /^:[A-Za-z_]\w*[?!]?/.exec(rest)!;
      push('lit', m[0]);
      i += m[0].length;
      continue;
    }
    // numbers
    if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(src[i + 1] ?? ''))) {
      const m = NUM.exec(rest);
      if (m) {
        push('num', m[0]);
        i += m[0].length;
        continue;
      }
    }
    // words
    const m = IDENT.exec(rest);
    if (m) {
      let word = m[0];
      if (word.endsWith('?') || word.endsWith('!')) {
        // only Ruby methods and Rust macros keep the trailing mark
        if (!(L.symbols || (L.name === 'Rust' && word.endsWith('!')))) word = word.slice(0, -1);
      }
      const after = src.slice(i + word.length).match(/^\s*(.)/)?.[1];
      let k: Kind | null = null;
      if (L.keywords.has(word)) k = 'kw';
      else if (L.literals.has(word)) k = 'lit';
      else if (L.types.has(word)) k = 'type';
      else if (L.name === 'Rust' && word.endsWith('!')) k = 'meta';
      else if (after === '(' || ['def', 'fn', 'func', 'function'].includes(prevWord)) k = 'fn';
      else if (/^[A-Z][a-z0-9]\w*$/.test(word) || ['class', 'struct', 'enum', 'interface', 'trait', 'impl', 'extends', 'new', 'class_name', 'module'].includes(prevWord)) k = 'type';
      else if (/^[A-Z][A-Z0-9_]+$/.test(word) && word.length > 1) k = 'lit';
      push(k, word);
      prevWord = word;
      i += word.length;
      continue;
    }
    push(/[{}()[\];,.]/.test(ch) ? 'punct' : null, ch);
    i++;
  }
  return out;
}

function tokJson(src: string): Tok[] {
  const out: Tok[] = [];
  const re = /("(?:[^"\\\n]|\\.)*"?)(\s*:)?|(\/\/[^\n]*|\/\*[\s\S]*?(?:\*\/|$))|(-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)|\b(true|false|null)\b|([{}[\],:])|(\s+|[^\s"/\d{}[\],:a-z-]+|.)/gy;
  let m: RegExpExecArray | null;
  while ((m = re.exec(src))) {
    if (m[1] !== undefined) {
      out.push([m[2] ? 'key' : 'str', m[1]]);
      if (m[2]) out.push(['punct', m[2]]);
    } else if (m[3] !== undefined) out.push(['com', m[3]]);
    else if (m[4] !== undefined) out.push(['num', m[4]]);
    else if (m[5] !== undefined) out.push(['lit', m[5]]);
    else if (m[6] !== undefined) out.push(['punct', m[6]]);
    else out.push([null, m[0]]);
    if (m[0] === '') re.lastIndex++;
  }
  return out;
}

function yamlValue(v: string, out: Tok[]): void {
  // value part of a line (no comment): scalar kinds
  const lead = /^\s*/.exec(v)![0];
  const body = v.slice(lead.length);
  const trail = /\s*$/.exec(body)![0];
  const core = body.slice(0, body.length - trail.length);
  out.push([null, lead]);
  if (!core) {
    /* nothing */
  } else if (/^["']/.test(core)) out.push(['str', core]);
  else if (/^[&*][\w-]+/.test(core)) {
    const a = /^[&*][\w-]+/.exec(core)![0];
    out.push(['meta', a]);
    yamlValue(core.slice(a.length), out);
  } else if (/^!!?[\w/-]*/.test(core)) {
    const a = /^!!?[\w/-]*/.exec(core)![0];
    out.push(['type', a]);
    yamlValue(core.slice(a.length), out);
  } else if (/^(true|false|yes|no|on|off|null|~)$/i.test(core)) out.push(['lit', core]);
  else if (/^[-+]?(\d[\d_]*(\.\d*)?([eE][+-]?\d+)?|0x[0-9a-fA-F]+|\.inf|\.nan)$/i.test(core)) out.push(['num', core]);
  else if (/^[|>][-+]?\d*$/.test(core)) out.push(['punct', core]);
  else if (/^[[{]/.test(core)) out.push(...tokJson(core));
  else out.push(['str', core]);
  out.push([null, trail]);
}

function tokYaml(src: string): Tok[] {
  const out: Tok[] = [];
  const lines = src.split('\n');
  let block = -1; // indent of an open | or > block scalar
  lines.forEach((line, n) => {
    if (n > 0) out.push([null, '\n']);
    const indent = /^\s*/.exec(line)![0].length;
    if (block >= 0) {
      if (line.trim() === '' || indent > block) {
        out.push(['str', line]);
        return;
      }
      block = -1;
    }
    if (/^(---|\.\.\.)\s*$/.test(line)) {
      out.push(['meta', line]);
      return;
    }
    // split off a comment (a # that starts the line or follows a space, outside quotes)
    let code = line;
    let comment = '';
    let quote = '';
    for (let i = 0; i < line.length; i++) {
      const c = line[i]!;
      if (quote) {
        if (c === quote) quote = '';
      } else if (c === '"' || c === "'") quote = c;
      else if (c === '#' && (i === 0 || /\s/.test(line[i - 1]!))) {
        code = line.slice(0, i);
        comment = line.slice(i);
        break;
      }
    }
    let rest = code;
    const lead = /^\s*(- )*/.exec(rest)![0];
    if (lead) {
      out.push([null, /^\s*/.exec(lead)![0]]);
      const dashes = lead.slice(/^\s*/.exec(lead)![0].length);
      if (dashes) out.push(['punct', dashes]);
      rest = rest.slice(lead.length);
    }
    const km = /^((?:"[^"]*"|'[^']*'|[^\s:#][^:#]*?))(\s*:)(?=\s|$)/.exec(rest);
    if (km) {
      out.push(['key', km[1]!], ['punct', km[2]!]);
      rest = rest.slice(km[0].length);
    }
    if (/^\s*[|>][-+]?\d*\s*$/.test(rest)) block = indent + (lead.length - /^\s*/.exec(lead)![0].length);
    yamlValue(rest, out);
    if (comment) out.push(['com', comment]);
  });
  return out;
}

function tokXml(src: string): Tok[] {
  const out: Tok[] = [];
  const re = /(<!--[\s\S]*?(?:-->|$))|(<!\[CDATA\[[\s\S]*?(?:\]\]>|$))|(<[?!][\s\S]*?(?:>|$))|(<\/?)([\w:.-]+)([^>]*?)(\/?>)|(&[#\w]+;)|([^<&]+|[<&])/gy;
  let m: RegExpExecArray | null;
  while ((m = re.exec(src))) {
    if (m[1]) out.push(['com', m[1]]);
    else if (m[2]) out.push(['str', m[2]]);
    else if (m[3]) out.push(['meta', m[3]]);
    else if (m[4]) {
      out.push(['punct', m[4]], ['tag', m[5]!]);
      const attrs = /([\w:.-]+)(\s*=\s*)("[^"]*"?|'[^']*'?)|(\s+)|([^\s]+?)/gy;
      let a: RegExpExecArray | null;
      const s = m[6]!;
      while ((a = attrs.exec(s))) {
        if (a[1]) out.push(['attr', a[1]], ['punct', a[2]!], ['str', a[3]!]);
        else out.push([null, a[0]]);
        if (a[0] === '') attrs.lastIndex++;
      }
      out.push(['punct', m[7]!]);
    } else if (m[8]) out.push(['lit', m[8]]);
    else out.push([null, m[0]]);
    if (m[0] === '') re.lastIndex++;
  }
  return out;
}

export function tokenize(src: string, lang: string): Tok[] {
  if (lang === 'json') return tokJson(src);
  if (lang === 'yaml') return tokYaml(src);
  if (lang === 'xml') return tokXml(src);
  const L = LANGS[lang];
  return L ? tokCode(src, L) : [[null, src]];
}

/** Replaces the content of `code` with highlighted spans for `lang` (text stays identical). */
export function highlightInto(code: HTMLElement, lang: string): void {
  const src = code.textContent ?? '';
  const frag = document.createDocumentFragment();
  for (const [k, s] of tokenize(src, lang)) {
    if (!s) continue;
    if (!k) frag.appendChild(document.createTextNode(s));
    else {
      const sp = document.createElement('span');
      sp.className = `tk-${k}`;
      sp.textContent = s;
      frag.appendChild(sp);
    }
  }
  code.replaceChildren(frag);
}

/** Highlights every supported fenced block below `root` and labels it with its language. */
export function highlightBlocks(root: ParentNode): void {
  const pres = root instanceof HTMLElement && root.matches('pre') ? [root] : Array.from(root.querySelectorAll<HTMLElement>('pre'));
  for (const pre of pres) {
    const code = pre.querySelector<HTMLElement>('code');
    const tag = /language-(\S+)/.exec(code?.className ?? '')?.[1] ?? '';
    const lang = highlightLang(tag);
    if (!code || !lang) continue;
    pre.classList.add('hl', CONFIG.has(lang) ? 'hl-config' : 'hl-code');
    pre.dataset['lang'] = langLabel(lang);
    highlightInto(code, lang);
  }
}
