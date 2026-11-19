/** How a text file was stored, so that saving writes it back the same way. */
export interface TextEncodingInfo {
  bom: boolean;
  utf16: 'le' | 'be' | null;
  eol: '\n' | '\r\n' | '\r';
  /** True if the original text ended with a line break. */
  finalNewline: boolean;
}

/** Decodes bytes (UTF-8, or UTF-16 with BOM) and normalises line ends to "\n". */
export function decodeText(bytes: Uint8Array): { text: string; info: TextEncodingInfo } {
  let utf16: TextEncodingInfo['utf16'] = null;
  let bom = false;
  let raw: string;
  if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    utf16 = 'le';
    bom = true;
    raw = new TextDecoder('utf-16le').decode(bytes.subarray(2));
  } else if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    utf16 = 'be';
    bom = true;
    raw = new TextDecoder('utf-16be').decode(bytes.subarray(2));
  } else {
    bom = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf;
    raw = new TextDecoder('utf-8').decode(bom ? bytes.subarray(3) : bytes);
  }
  if (raw.includes('\u0000') && utf16 === null) throw new Error('This file looks binary, not text.');
  const crlf = (raw.match(/\r\n/g) ?? []).length;
  const lf = (raw.match(/(?<!\r)\n/g) ?? []).length;
  const cr = (raw.match(/\r(?!\n)/g) ?? []).length;
  const eol: TextEncodingInfo['eol'] = crlf >= lf && crlf >= cr && crlf > 0 ? '\r\n' : cr > lf ? '\r' : '\n';
  const text = raw.replace(/\r\n?/g, '\n');
  return { text, info: { bom, utf16, eol, finalNewline: text.endsWith('\n') } };
}

export function encodeText(text: string, info: TextEncodingInfo): Uint8Array {
  const out = info.eol === '\n' ? text : text.replace(/\n/g, info.eol);
  if (info.utf16) {
    const buf = new Uint8Array(2 + out.length * 2);
    const le = info.utf16 === 'le';
    buf[0] = le ? 0xff : 0xfe;
    buf[1] = le ? 0xfe : 0xff;
    for (let i = 0; i < out.length; i++) {
      const c = out.charCodeAt(i);
      buf[2 + i * 2] = le ? c & 0xff : c >> 8;
      buf[3 + i * 2] = le ? c >> 8 : c & 0xff;
    }
    return buf;
  }
  const body = new TextEncoder().encode(out);
  if (!info.bom) return body;
  const buf = new Uint8Array(body.length + 3);
  buf.set([0xef, 0xbb, 0xbf]);
  buf.set(body, 3);
  return buf;
}
