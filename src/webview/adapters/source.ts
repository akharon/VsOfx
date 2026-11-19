import type { Format, SourcePart } from '../../shared/model';
import { decodeText } from './textCodec';
import { isZip, readZip } from './zip';

const MAX_PART_CHARS = 2_000_000;
const TEXT_PART = /\.(xml|rels|vml|txt|json|rdf|csv)$|^mimetype$/i;

/** Text parts of a (saved) document for the read-only Source view. */
export function sourceParts(format: Format, bytes: Uint8Array): SourcePart[] {
  if (format === 'pdf') return [{ name: 'PDF', text: cap(latin1(bytes.subarray(0, MAX_PART_CHARS)), bytes.length > MAX_PART_CHARS) }];
  if (!isZip(bytes)) return [{ name: 'Text', text: cap(decodeText(bytes).text, false) }];
  const entries = readZip(bytes);
  const names = Object.keys(entries).filter((n) => TEXT_PART.test(n));
  // the main part first
  const main = ['word/document.xml', 'content.xml', 'xl/workbook.xml', 'ppt/presentation.xml'];
  names.sort((a, b) => rank(a, main) - rank(b, main));
  return names.map((name) => {
    const text = new TextDecoder().decode(entries[name]!);
    return { name, text: cap(/\.(xml|rels|vml|rdf)$/i.test(name) ? prettyXml(text) : text, false) };
  });
}

function rank(n: string, main: string[]): number {
  const i = main.indexOf(n);
  if (i >= 0) return i;
  if (/^(xl\/worksheets|ppt\/slides)\//.test(n)) return 10;
  return 20;
}

function cap(text: string, truncated: boolean): string {
  if (text.length > MAX_PART_CHARS) return text.slice(0, MAX_PART_CHARS) + '\n… (truncated)';
  return truncated ? text + '\n… (truncated)' : text;
}

function latin1(b: Uint8Array): string {
  let s = '';
  for (let i = 0; i < b.length; i += 8192) s += String.fromCharCode(...b.subarray(i, i + 8192));
  // keep it readable: control characters other than line breaks / tabs become dots
  return s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]/g, '·');
}

/** Indents XML one element per line; text content stays on its element's line. */
export function prettyXml(xml: string): string {
  const out: string[] = [];
  let depth = 0;
  let inline = false; // the last line is an open tag that may still get text and its close tag
  for (const m of xml.matchAll(/<!\[CDATA\[[\s\S]*?\]\]>|<!--[\s\S]*?-->|<[^>]+>|[^<]+/g)) {
    const t = m[0];
    if (t.startsWith('</')) {
      depth = Math.max(0, depth - 1);
      if (inline) out[out.length - 1] += t;
      else out.push('  '.repeat(depth) + t);
      inline = false;
    } else if (t.startsWith('<?') || t.startsWith('<!')) {
      if (inline && t.startsWith('<![CDATA[')) out[out.length - 1] += t;
      else out.push('  '.repeat(depth) + t);
    } else if (t.startsWith('<')) {
      out.push('  '.repeat(depth) + t);
      if (t.endsWith('/>')) inline = false;
      else {
        depth++;
        inline = true;
      }
    } else {
      if (!t.trim()) continue;
      if (inline) out[out.length - 1] += t;
      else out.push('  '.repeat(depth) + t.trim());
    }
  }
  return out.join('\n');
}
