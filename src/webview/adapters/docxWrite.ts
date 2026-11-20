// Writes formatted paragraphs, block kinds (headings, lists, quote) and new tables into
// word/document.xml. Element order inside w:rPr / w:pPr follows the OOXML schema, because
// Word refuses files with misordered properties. Styles, numbering definitions and hyperlink
// relationships that a document does not have yet are added to their parts.
import { type BlockKind, type HAlign, type ParaFmt, type Run, type TableSpec } from '../../shared/richtext';
import { childrenNS, firstChildNS, parseXml, serializeXml, xmlSafe } from './xml';
import { optionalText, type ZipEntries } from './zip';

export const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const REL_HYPERLINK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const REL_NUMBERING = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';
const REL_STYLES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';

const RPR_ORDER = [
  'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss',
  'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs',
  'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout',
  'specVanish', 'oMath',
];
const PPR_ORDER = [
  'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr',
  'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE',
  'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents',
  'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle',
  'rPr', 'sectPr', 'pPrChange',
];
/** Paragraph children that stay when the content is replaced. */
const KEEP_IN_PARAGRAPH = new Set(['pPr', 'bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd']);
/** Run properties managed by the editor (removed from the inherited w:rPr before applying runs). */
const MANAGED_RPR = new Set(['b', 'bCs', 'i', 'iCs', 'strike', 'dstrike', 'u', 'vertAlign']);

const HEADING_SIZE: Record<string, number> = { h1: 32, h2: 28, h3: 26, h4: 24, h5: 22, h6: 22 };

/** Lazily loaded parts that the writer may change (styles, numbering, relationships, content types). */
export class DocxParts {
  private readonly docs = new Map<string, Document>();
  private relCounter = 0;
  private numIds: Partial<Record<'ul' | 'ol', string>> = {};

  constructor(
    private readonly entries: ZipEntries,
    readonly doc: Document,
  ) {}

  private part(name: string, create?: () => string): Document | undefined {
    let d = this.docs.get(name);
    if (d) return d;
    const src = optionalText(this.entries, name) ?? create?.();
    if (src === undefined) return undefined;
    d = parseXml(src);
    this.docs.set(name, d);
    return d;
  }

  /** Changed parts as replacements for writeZip. */
  replacements(): Record<string, string> {
    const out: Record<string, string> = {};
    for (const [name, d] of this.docs) out[name] = serializeXml(d);
    return out;
  }

  private rels(): Document {
    return this.part('word/_rels/document.xml.rels', () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PR}"/>`)!;
  }

  private addRel(type: string, target: string, external: boolean): string {
    const rels = this.rels();
    const root = rels.documentElement;
    const ids = new Set(childrenNS(root, PR, 'Relationship').map((r) => r.getAttribute('Id')));
    let id: string;
    do id = `rIdVsofx${++this.relCounter}`;
    while (ids.has(id));
    const r = rels.createElementNS(PR, 'Relationship');
    r.setAttribute('Id', id);
    r.setAttribute('Type', type);
    r.setAttribute('Target', target);
    if (external) r.setAttribute('TargetMode', 'External');
    root.appendChild(r);
    return id;
  }

  private hasRel(type: string): boolean {
    return childrenNS(this.rels().documentElement, PR, 'Relationship').some((r) => r.getAttribute('Type') === type);
  }

  private addOverride(partName: string, contentType: string): void {
    const ct = this.part('[Content_Types].xml');
    if (!ct) return;
    const root = ct.documentElement;
    if (childrenNS(root, CT, 'Override').some((o) => o.getAttribute('PartName') === partName)) return;
    const o = ct.createElementNS(CT, 'Override');
    o.setAttribute('PartName', partName);
    o.setAttribute('ContentType', contentType);
    root.appendChild(o);
  }

  hyperlinkRel(href: string): string {
    return this.addRel(REL_HYPERLINK, href, true);
  }

  /** Makes sure a paragraph style exists (Heading1..6, Quote, ListParagraph). */
  ensureStyle(id: string): void {
    const styles = this.part('word/styles.xml', () => {
      if (!this.hasRel(REL_STYLES)) this.addRel(REL_STYLES, 'styles.xml', false);
      this.addOverride('/word/styles.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml');
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W}"/>`;
    })!;
    const root = styles.documentElement;
    if (childrenNS(root, W, 'style').some((s) => s.getAttributeNS(W, 'styleId') === id || s.getAttribute('w:styleId') === id)) return;
    const pfx = root.prefix ? `${root.prefix}:` : 'w:';
    const el = (name: string, attrs: Record<string, string> = {}, ...kids: Element[]) => {
      const e = styles.createElementNS(W, `${pfx}${name}`);
      for (const [k, v] of Object.entries(attrs)) e.setAttributeNS(W, `${pfx}${k}`, v);
      for (const k of kids) e.appendChild(k);
      return e;
    };
    const h = /^Heading([1-6])$/.exec(id);
    let style: Element;
    if (h) {
      const n = Number(h[1]);
      style = el(
        'style',
        { type: 'paragraph', styleId: id },
        el('name', { val: `heading ${n}` }),
        el('basedOn', { val: 'Normal' }),
        el('next', { val: 'Normal' }),
        el('uiPriority', { val: '9' }),
        el('qFormat'),
        el('pPr', {}, el('keepNext'), el('keepLines'), el('spacing', { before: n === 1 ? '360' : '240', after: '80' }), el('outlineLvl', { val: String(n - 1) })),
        el('rPr', {}, el('b'), el('bCs'), el('sz', { val: String(HEADING_SIZE[`h${n}`]) }), el('szCs', { val: String(HEADING_SIZE[`h${n}`]) })),
      );
    } else if (id === 'Quote') {
      style = el(
        'style',
        { type: 'paragraph', styleId: id },
        el('name', { val: 'Quote' }),
        el('basedOn', { val: 'Normal' }),
        el('next', { val: 'Normal' }),
        el('uiPriority', { val: '29' }),
        el('qFormat'),
        el('pPr', {}, el('spacing', { before: '120', after: '120' }), el('ind', { left: '720', right: '720' })),
        el('rPr', {}, el('i'), el('iCs'), el('color', { val: '595959' })),
      );
    } else {
      style = el(
        'style',
        { type: 'paragraph', styleId: id },
        el('name', { val: 'List Paragraph' }),
        el('basedOn', { val: 'Normal' }),
        el('uiPriority', { val: '34' }),
        el('qFormat'),
        el('pPr', {}, el('ind', { left: '720' }), el('contextualSpacing')),
      );
    }
    root.appendChild(style);
  }

  /** numId of a bullet ('ul') or decimal ('ol') list definition, created on first use. */
  numId(kind: 'ul' | 'ol'): string {
    const cached = this.numIds[kind];
    if (cached) return cached;
    const numbering = this.part('word/numbering.xml', () => {
      if (!this.hasRel(REL_NUMBERING)) this.addRel(REL_NUMBERING, 'numbering.xml', false);
      this.addOverride('/word/numbering.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml');
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}"/>`;
    })!;
    const root = numbering.documentElement;
    const pfx = root.prefix ? `${root.prefix}:` : 'w:';
    const el = (name: string, attrs: Record<string, string> = {}, ...kids: Element[]) => {
      const e = numbering.createElementNS(W, `${pfx}${name}`);
      for (const [k, v] of Object.entries(attrs)) e.setAttributeNS(W, `${pfx}${k}`, v);
      for (const k of kids) e.appendChild(k);
      return e;
    };
    const attr = (e: Element, k: string) => Number(e.getAttributeNS(W, k) ?? e.getAttribute(`w:${k}`) ?? 0);
    const abstractId = String(Math.max(-1, ...childrenNS(root, W, 'abstractNum').map((a) => attr(a, 'abstractNumId'))) + 1);
    const numId = String(Math.max(0, ...childrenNS(root, W, 'num').map((n) => attr(n, 'numId'))) + 1);
    const levels: Element[] = [];
    for (let l = 0; l < 9; l++) {
      const bullet = kind === 'ul';
      levels.push(
        el(
          'lvl',
          { ilvl: String(l) },
          el('start', { val: '1' }),
          el('numFmt', { val: bullet ? 'bullet' : l % 3 === 1 ? 'lowerLetter' : l % 3 === 2 ? 'lowerRoman' : 'decimal' }),
          el('lvlText', { val: bullet ? ['•', '◦', '▪'][l % 3]! : `%${l + 1}.` }),
          el('lvlJc', { val: 'left' }),
          el('pPr', {}, el('ind', { left: String(720 * (l + 1)), hanging: '360' })),
        ),
      );
    }
    const abstract = el('abstractNum', { abstractNumId: abstractId }, el('multiLevelType', { val: 'hybridMultilevel' }), ...levels);
    // schema: all abstractNum elements come before the num elements
    const firstNum = childrenNS(root, W, 'num')[0];
    if (firstNum) root.insertBefore(abstract, firstNum);
    else root.appendChild(abstract);
    root.appendChild(el('num', { numId }, el('abstractNumId', { val: abstractId })));
    this.numIds[kind] = numId;
    return numId;
  }
}

/** Inserts `child` into `parent` at the position the schema order requires. */
function insertOrdered(parent: Element, child: Element, order: string[]): void {
  const idx = order.indexOf(child.localName);
  for (let n = parent.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const i = order.indexOf((n as Element).localName);
    if (i > idx || i === -1) {
      parent.insertBefore(child, n);
      return;
    }
  }
  parent.appendChild(child);
}

function removeChildren(parent: Element, names: Set<string> | string[]): void {
  const set = names instanceof Set ? names : new Set(names);
  for (const c of Array.from(parent.childNodes)) {
    if (c.nodeType === 1 && (c as Element).namespaceURI === W && set.has((c as Element).localName)) parent.removeChild(c);
  }
}

const prefixOf = (el: Element) => (el.prefix ? `${el.prefix}:` : '');

function wEl(doc: Document, pfx: string, name: string, attrs: Record<string, string> = {}): Element {
  const e = doc.createElementNS(W, `${pfx}${name}`);
  for (const [k, v] of Object.entries(attrs)) e.setAttributeNS(W, `${pfx || 'w:'}${k}`, v);
  return e;
}

/** w:rPr of the paragraph's first run, without the properties the editor manages. */
function baseRunProps(p: Element): Element | undefined {
  for (const r of Array.from(p.getElementsByTagNameNS(W, 'r'))) {
    const rPr = firstChildNS(r, W, 'rPr');
    if (!rPr) return undefined;
    const clone = rPr.cloneNode(true) as Element;
    removeChildren(clone, MANAGED_RPR);
    removeChildren(clone, ['rPrChange']);
    return clone;
  }
  return undefined;
}

function runElement(parts: DocxParts, pfx: string, base: Element | undefined, run: Run): Element {
  const doc = parts.doc;
  const r = wEl(doc, pfx, 'r');
  const rPr = base ? (base.cloneNode(true) as Element) : wEl(doc, pfx, 'rPr');
  const add = (name: string, attrs: Record<string, string> = {}) => {
    removeChildren(rPr, [name]);
    insertOrdered(rPr, wEl(doc, pfx, name, attrs), RPR_ORDER);
  };
  if (run.b) add('b');
  if (run.i) add('i');
  if (run.s) add('strike');
  if (run.u || run.href) add('u', { val: 'single' });
  if (run.sup) add('vertAlign', { val: 'superscript' });
  else if (run.sub) add('vertAlign', { val: 'subscript' });
  if (run.href) add('color', { val: '0563C1' });
  if (run.code) {
    add('rFonts', { ascii: 'Consolas', hAnsi: 'Consolas', cs: 'Consolas' });
    add('shd', { val: 'clear', color: 'auto', fill: 'F2F2F2' });
  }
  if (rPr.firstChild) r.appendChild(rPr);
  for (const part of xmlSafe(run.t).split(/(\t|\n)/)) {
    if (part === '') continue;
    if (part === '\t') r.appendChild(wEl(doc, pfx, 'tab'));
    else if (part === '\n') r.appendChild(wEl(doc, pfx, 'br'));
    else {
      const t = wEl(doc, pfx, 't');
      t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
      t.appendChild(doc.createTextNode(part));
      r.appendChild(t);
    }
  }
  return r;
}

/** Replaces the paragraph's content by the formatted runs (and applies the block kind if set). */
export function writeParagraph(parts: DocxParts, p: Element, fmt: ParaFmt): void {
  const pfx = prefixOf(p);
  const base = baseRunProps(p);
  for (const child of Array.from(p.childNodes)) {
    const keep = child.nodeType === 1 && (child as Element).namespaceURI === W && KEEP_IN_PARAGRAPH.has((child as Element).localName);
    if (!keep) p.removeChild(child);
  }
  for (const run of fmt.runs) {
    const r = runElement(parts, pfx, base, run);
    if (run.href) {
      const h = wEl(parts.doc, pfx, 'hyperlink');
      h.setAttributeNS(R, 'r:id', parts.hyperlinkRel(run.href));
      h.appendChild(r);
      p.appendChild(h);
    } else p.appendChild(r);
  }
  if (fmt.block) applyBlock(parts, p, fmt.block);
  if (fmt.align) applyAlign(parts, p, fmt.align);
}

/** Paragraph alignment (w:jc). */
export function applyAlign(parts: DocxParts, p: Element, align: HAlign): void {
  const pfx = prefixOf(p);
  let pPr = firstChildNS(p, W, 'pPr');
  if (!pPr) {
    pPr = wEl(parts.doc, pfx, 'pPr');
    p.insertBefore(pPr, p.firstChild);
  }
  removeChildren(pPr, ['jc']);
  insertOrdered(pPr, wEl(parts.doc, pfx, 'jc', { val: align === 'justify' ? 'both' : align }), PPR_ORDER);
}


export function applyBlock(parts: DocxParts, p: Element, block: BlockKind): void {
  const pfx = prefixOf(p);
  let pPr = firstChildNS(p, W, 'pPr');
  if (!pPr) {
    pPr = wEl(parts.doc, pfx, 'pPr');
    p.insertBefore(pPr, p.firstChild);
  }
  removeChildren(pPr, ['pStyle', 'numPr', 'outlineLvl']);
  const set = (name: string, attrs: Record<string, string>, ...kids: Element[]) => {
    const e = wEl(parts.doc, pfx, name, attrs);
    for (const k of kids) e.appendChild(k);
    insertOrdered(pPr, e, PPR_ORDER);
  };
  if (/^h[1-6]$/.test(block)) {
    const id = `Heading${block[1]}`;
    parts.ensureStyle(id);
    set('pStyle', { val: id });
  } else if (block === 'quote') {
    parts.ensureStyle('Quote');
    set('pStyle', { val: 'Quote' });
  } else if (block === 'ul' || block === 'ol') {
    parts.ensureStyle('ListParagraph');
    set('pStyle', { val: 'ListParagraph' });
    set('numPr', {}, wEl(parts.doc, pfx, 'ilvl', { val: '0' }), wEl(parts.doc, pfx, 'numId', { val: parts.numId(block) }));
  }
  if (!pPr.firstChild) p.removeChild(pPr);
}

/**
 * A new paragraph after `anchor` with the anchor's paragraph properties (style, list) and the
 * character formatting of its first run; then `fmt` is written into it.
 */
export function newParagraphAfter(parts: DocxParts, anchor: Element, fmt: ParaFmt): Element {
  const pfx = prefixOf(anchor);
  const p = wEl(parts.doc, pfx, 'p');
  const pPr = firstChildNS(anchor, W, 'pPr');
  if (pPr) {
    const clone = pPr.cloneNode(true) as Element;
    removeChildren(clone, ['sectPr', 'pPrChange']); // the section break stays with the anchor
    p.appendChild(clone);
  }
  const base = baseRunProps(anchor);
  if (base) {
    // carrier for the inherited run properties; writeParagraph replaces it
    const r = wEl(parts.doc, pfx, 'r');
    r.appendChild(base);
    p.appendChild(r);
  }
  anchor.parentNode!.insertBefore(p, anchor.nextSibling);
  writeParagraph(parts, p, fmt);
  return p;
}

/** A new bordered table; `cell(r, c)` gives the formatted content of each cell. */
export function buildTable(parts: DocxParts, pfx: string, spec: TableSpec, cell: (r: number, c: number) => ParaFmt | undefined): Element {
  const doc = parts.doc;
  const width = Math.floor(9000 / spec.cols);
  const tbl = wEl(doc, pfx, 'tbl');
  const tblPr = wEl(doc, pfx, 'tblPr');
  tblPr.appendChild(wEl(doc, pfx, 'tblW', { w: '0', type: 'auto' }));
  const borders = wEl(doc, pfx, 'tblBorders');
  for (const side of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']) {
    borders.appendChild(wEl(doc, pfx, side, { val: 'single', sz: '4', space: '0', color: 'auto' }));
  }
  tblPr.appendChild(borders);
  tblPr.appendChild(wEl(doc, pfx, 'tblLook', { val: '04A0', firstRow: '1', lastRow: '0', firstColumn: '1', lastColumn: '0', noHBand: '0', noVBand: '1' }));
  tbl.appendChild(tblPr);
  const grid = wEl(doc, pfx, 'tblGrid');
  for (let c = 0; c < spec.cols; c++) grid.appendChild(wEl(doc, pfx, 'gridCol', { w: String(width) }));
  tbl.appendChild(grid);
  for (let r = 0; r < spec.rows; r++) {
    const tr = wEl(doc, pfx, 'tr');
    for (let c = 0; c < spec.cols; c++) {
      const tc = wEl(doc, pfx, 'tc');
      const tcPr = wEl(doc, pfx, 'tcPr');
      tcPr.appendChild(wEl(doc, pfx, 'tcW', { w: String(width), type: 'dxa' }));
      tc.appendChild(tcPr);
      const p = wEl(doc, pfx, 'p');
      tc.appendChild(p);
      const f = cell(r, c);
      if (f) writeParagraph(parts, p, r === 0 ? { ...f, runs: f.runs.map((x) => ({ ...x, b: true as const })) } : f);
      tr.appendChild(tc);
    }
    tbl.appendChild(tr);
  }
  return tbl;
}

/** Inserts tables after paragraph `p`, in the given order. A cell must end with a paragraph. */
export function insertTablesAfter(p: Element, tables: Element[]): void {
  const parent = p.parentNode as Element | null;
  if (!parent) return;
  const ref: Node | null = p.nextSibling;
  for (const t of tables) parent.insertBefore(t, ref);
  if (parent.namespaceURI === W && parent.localName === 'tc') {
    let last: Node | null = parent.lastChild;
    while (last && last.nodeType !== 1) last = last.previousSibling;
    if (last && (last as Element).localName === 'tbl') parent.appendChild(wEl(p.ownerDocument!, prefixOf(p), 'p'));
  }
}
