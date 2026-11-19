// Minimal Excel number-format renderer (exceljs returns raw values, not formatted text).
// Covers the common cases: General, 0 / 0.00 / #,##0.00, %, scientific, currency literals,
// negative/zero sections and date/time patterns. Anything exotic falls back to General.

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

export function general(n: number): string {
  if (!Number.isFinite(n)) return String(n);
  const s = String(Number(n.toPrecision(15)));
  return s;
}

/** Splits on `;` outside quotes. */
function sections(fmt: string): string[] {
  const out: string[] = [];
  let cur = '';
  let quoted = false;
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i]!;
    if (ch === '"') quoted = !quoted;
    if (ch === ';' && !quoted) {
      out.push(cur);
      cur = '';
    } else cur += ch;
  }
  out.push(cur);
  return out;
}

/** Removes quotes/brackets and returns display literals. `[$€-407]` becomes `€`. */
function literal(s: string): string {
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const ch = s[i]!;
    if (ch === '"') {
      const end = s.indexOf('"', i + 1);
      out += s.slice(i + 1, end < 0 ? undefined : end);
      i = end < 0 ? s.length : end;
    } else if (ch === '[') {
      const end = s.indexOf(']', i);
      const inner = s.slice(i + 1, end < 0 ? undefined : end);
      const cur = /^\$([^-\]]*)/.exec(inner);
      if (cur) out += cur[1];
      i = end < 0 ? s.length : end;
    } else if (ch === '\\') {
      out += s[++i] ?? '';
    } else if (ch === '_') {
      out += ' ';
      i++;
    } else if (ch === '*') {
      i++;
    } else out += ch;
  }
  return out;
}

function stripQuoted(s: string): string {
  return s.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '');
}

export function isDateFormat(fmt: string | undefined): boolean {
  if (!fmt || fmt === 'General') return false;
  return /[ymdhs]/i.test(stripQuoted(fmt).replace(/\\./g, '').replace(/[_*]./g, ''));
}

/** Index of the first digit placeholder outside quotes and brackets, or -1. */
function findPlaceholder(section: string): number {
  let quoted = false;
  let bracket = false;
  for (let i = 0; i < section.length; i++) {
    const ch = section[i]!;
    if (ch === '"' && !bracket) quoted = !quoted;
    else if (ch === '[' && !quoted) bracket = true;
    else if (ch === ']' && !quoted) bracket = false;
    else if (ch === '\\' || ch === '_' || ch === '*') i++;
    else if (!quoted && !bracket && (/[0#?]/.test(ch) || (ch === '.' && /[0#?]/.test(section[i + 1] ?? '')))) return i;
  }
  return -1;
}

export function formatNumber(n: number, fmt: string | undefined): string {
  if (!fmt || fmt === 'General' || !Number.isFinite(n)) return general(n);
  const parts = sections(fmt);
  let section = parts[0]!;
  let value = Math.abs(n); // the sign is handled by the section choice or added below
  if (n < 0 && parts.length > 1) section = parts[1]!;
  else if (n === 0 && parts.length > 2) section = parts[2]!;
  const bare = stripQuoted(section);
  if (/[?#0]\s*\/\s*[?#0]/.test(bare) || /@/.test(bare)) return general(n);

  const idx = findPlaceholder(section);
  if (idx < 0) return literal(section) || general(n);
  let end = idx;
  while (end < section.length && /[0#?,.]/.test(section[end]!)) end++;
  let sci: string | undefined;
  const e = /^E[+-]0+/.exec(section.slice(end));
  if (e) {
    sci = e[0];
    end += e[0].length;
  }
  const prefix = literal(section.slice(0, idx));
  const suffix = literal(section.slice(end));
  const pattern = section.slice(idx, sci ? end - sci.length : end);

  if (/%/.test(stripQuoted(section))) value *= 100;
  const [intPat, decPat = ''] = pattern.split('.');
  const maxDec = (decPat.match(/[0#?]/g) ?? []).length;
  const minDec = (decPat.match(/0/g) ?? []).length;
  const minInt = (intPat!.match(/0/g) ?? []).length;
  const grouping = /,/.test(intPat!);

  let body: string;
  if (sci) {
    const digits = sci.length - 2;
    const [mant, exp] = value.toExponential(maxDec).split('e') as [string, string];
    const sign = exp.startsWith('-') ? '-' : sci[1] === '+' ? '+' : '';
    body = `${mant}E${sign || (sci[1] === '+' ? '+' : '')}${exp.replace(/^[+-]/, '').padStart(digits, '0')}`;
  } else {
    let s = value.toFixed(maxDec);
    let [ip, dp = ''] = s.split('.') as [string, string?];
    dp = dp ?? '';
    while (dp.length > minDec && dp.endsWith('0')) dp = dp.slice(0, -1);
    ip = ip.padStart(minInt, '0');
    if (minInt === 0 && ip === '0' && maxDec > 0 && dp) ip = '';
    if (grouping) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    s = dp ? `${ip}.${dp}` : ip;
    body = s;
  }
  const negative = n < 0 && parts.length === 1;
  return `${negative ? '-' : ''}${prefix}${body}${suffix}`;
}

type DateToken = { t: string; lit?: boolean };

function tokenizeDate(fmt: string): DateToken[] {
  const tokens: DateToken[] = [];
  const f = sections(fmt)[0]!;
  for (let i = 0; i < f.length; ) {
    const ch = f[i]!;
    if (ch === '"') {
      const end = f.indexOf('"', i + 1);
      tokens.push({ t: f.slice(i + 1, end < 0 ? undefined : end), lit: true });
      i = end < 0 ? f.length : end + 1;
    } else if (ch === '[') {
      const end = f.indexOf(']', i);
      i = end < 0 ? f.length : end + 1; // colors, locale tags, elapsed-time brackets are ignored
    } else if (ch === '\\') {
      tokens.push({ t: f[i + 1] ?? '', lit: true });
      i += 2;
    } else if (ch === '_' || ch === '*') {
      i += 2;
    } else {
      const m = /^(yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM|A\/P)/i.exec(f.slice(i));
      if (m) {
        tokens.push({ t: m[0] });
        i += m[0].length;
      } else {
        tokens.push({ t: ch, lit: true });
        i++;
      }
    }
  }
  return tokens;
}

export function formatDate(d: Date, fmt: string | undefined): string {
  if (isNaN(+d)) return '';
  if (!fmt || !isDateFormat(fmt)) return d.toISOString().replace(/(T00:00:00)?(\.000)?Z$/, '').replace('T', ' ');
  const tokens = tokenizeDate(fmt);
  const ampm = tokens.some((x) => !x.lit && /^(AM\/PM|A\/P)$/i.test(x.t));
  const pad = (n: number, w = 2) => String(n).padStart(w, '0');
  let out = '';
  tokens.forEach((tok, i) => {
    if (tok.lit) {
      out += tok.t;
      return;
    }
    const t = tok.t.toLowerCase();
    const prev = tokens.slice(0, i).reverse().find((x) => !x.lit);
    const next = tokens.slice(i + 1).find((x) => !x.lit);
    const minute = (t === 'm' || t === 'mm') && ((prev && /^h/i.test(prev.t)) || (next && /^s/i.test(next.t)));
    const h24 = d.getUTCHours();
    switch (true) {
      case t === 'yyyy': out += pad(d.getUTCFullYear(), 4); break;
      case t === 'yy': out += pad(d.getUTCFullYear() % 100); break;
      case minute && t === 'm': out += d.getUTCMinutes(); break;
      case minute: out += pad(d.getUTCMinutes()); break;
      case t === 'mmmmm': out += MONTHS[d.getUTCMonth()]![0]; break;
      case t === 'mmmm': out += MONTHS[d.getUTCMonth()]; break;
      case t === 'mmm': out += MONTHS[d.getUTCMonth()]!.slice(0, 3); break;
      case t === 'mm': out += pad(d.getUTCMonth() + 1); break;
      case t === 'm': out += d.getUTCMonth() + 1; break;
      case t === 'dddd': out += DAYS[d.getUTCDay()]; break;
      case t === 'ddd': out += DAYS[d.getUTCDay()]!.slice(0, 3); break;
      case t === 'dd': out += pad(d.getUTCDate()); break;
      case t === 'd': out += d.getUTCDate(); break;
      case t === 'hh': out += pad(ampm ? h24 % 12 || 12 : h24); break;
      case t === 'h': out += ampm ? h24 % 12 || 12 : h24; break;
      case t === 'ss': out += pad(d.getUTCSeconds()); break;
      case t === 's': out += d.getUTCSeconds(); break;
      case t === 'am/pm': out += h24 < 12 ? 'AM' : 'PM'; break;
      case t === 'a/p': out += h24 < 12 ? 'A' : 'P'; break;
    }
  });
  return out;
}
