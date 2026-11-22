import type { ApplyResult, Edit, ParagraphEdit, TextModel } from '../../shared/model';
import { cellIndex, fmtEquals, isCellIndex, isNewParaIndex, SAFE_HREF, type BlockKind, type HAlign, type ParaFmt, type Run, type TableSpec } from '../../shared/richtext';
import { escapeHtml } from '../sanitize';
import type { FormatAdapter } from './Adapter';
import { collectParagraphs, isEl, MIME_ODT, NS, odfText, openOdf, saveOdf, setOdfText, type OdfPackage } from './odf';
import { descendantsNS } from './xml';

const XLINK = 'http://www.w3.org/1999/xlink';

interface InlineStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  code?: boolean;
  sup?: boolean;
  sub?: boolean;
}

interface Original {
  /** Clone of the paragraph element before the first edit. */
  el: Element;
  text: string;
  fmt?: ParaFmt;
  block: BlockKind;
}

/** Automatic styles this editor adds (names start with "Vsofx"). */
const HEADING_PT = [24, 20, 16, 14, 12, 12];
const FLAG_KEYS = ['b', 'i', 'u', 's', 'code', 'sup', 'sub'] as const;
/** Suffixes of the derived alignment styles (see applyAlign). */
const ALIGN_SUFFIX = /_VsofxA[lcrj]$/;

/**
 * .odt: paragraphs and headings of content.xml with their character formatting.
 * Edits change the XML in place: formatted runs become text:span with automatic styles,
 * headings text:h, lists text:list; added paragraphs and tables are inserted after their anchor.
 */
export class OdtAdapter implements FormatAdapter {
  readonly readonly = false;
  private pkg!: OdfPackage;
  private body!: Element;
  /** Editable paragraphs by index: document paragraphs, added ones, cells of added tables. */
  private readonly paras = new Map<number, Element>();
  private readonly originals = new Map<number, Original>();
  private readonly tables = new Map<number, { spec: TableSpec; el: Element }>();
  private styles = new Map<string, InlineStyle>();
  private listKinds = new Map<string, 'ul' | 'ol'>();
  private paraAlign = new Map<string, HAlign>();
  private readonly created = new Set<string>();
  /** Removed original paragraphs: the detached node and its place (see deleteOriginal). */
  private readonly removedParas = new Map<number, { unit: Element; parent: Node; prev: Node | null; next: Node | null; tables: number[] }>();

  async load(bytes: Uint8Array, edits: Edit[]): Promise<TextModel> {
    this.pkg = openOdf(bytes, MIME_ODT);
    const body = descendantsNS(this.pkg.content, NS.office, 'text')[0];
    if (!body) throw new Error('This OpenDocument file has no text body');
    this.body = body;
    collectParagraphs(body).forEach((p, i) => this.paras.set(i, p));
    this.readStyles();
    for (const e of edits) this.applyEdit(e);
    this.readStyles();
    return { type: 'text', html: this.render(body), rich: true };
  }

  applyEdit(edit: Edit): ApplyResult {
    if (edit.kind === 'table') {
      if (edit.op === 'insert') this.insertTable(edit.table);
      else this.removeTable(edit.table.id);
      return {};
    }
    if (edit.kind === 'newpara') {
      if (edit.op === 'insert') this.insertParagraph(edit.index, edit.after, edit.fmt);
      else this.removeParagraph(edit.index);
      return {};
    }
    if (edit.kind === 'delpara') {
      if (edit.op === 'remove') this.deleteOriginal(edit.index, edit.to);
      else this.restoreOriginal(edit.index);
      return {};
    }
    if (edit.kind !== 'paragraph') throw new Error('Not a paragraph edit');
    this.editParagraph(edit);
    return {};
  }

  /**
   * Takes an original paragraph out of the document but keeps its node (and where it was), so
   * that restoreOriginal can put it back. A list item / list that only holds it goes with it.
   */
  private deleteOriginal(index: number, to: number | undefined): void {
    const el = this.paras.get(index);
    if (!el || isCellIndex(index) || isNewParaIndex(index) || this.removedParas.has(index)) throw new Error(`Paragraph ${index} cannot be removed`);
    for (let a: Node | null = el.parentNode; a; a = a.parentNode) {
      if (isEl(a, NS.table, 'table-cell')) throw new Error(`Paragraph ${index} is in a table`);
    }
    let unit: Element = el;
    for (;;) {
      const parent = unit.parentNode as Element | null;
      if (!parent || !(isEl(parent, NS.text, 'list-item') || isEl(parent, NS.text, 'list'))) break;
      if (firstEl(parent) !== unit || nextEl(unit)) break;
      unit = parent;
    }
    const parent = unit.parentNode!;
    const info = { unit, parent, prev: unit.previousSibling, next: unit.nextSibling, tables: [] as number[] };
    parent.removeChild(unit);
    if (to !== undefined) {
      for (const [id, t] of this.tables) {
        if (t.spec.after === index) {
          this.tables.set(id, { ...t, spec: { ...t.spec, after: to } });
          info.tables.push(id);
        }
      }
    }
    this.removedParas.set(index, info);
  }

  private restoreOriginal(index: number): void {
    const info = this.removedParas.get(index);
    if (!info) throw new Error(`Paragraph ${index} was not removed`);
    const { unit, parent, prev, next } = info;
    if (next && next.parentNode === parent) parent.insertBefore(unit, next);
    else if (prev && prev.parentNode === parent) parent.insertBefore(unit, prev.nextSibling);
    else parent.appendChild(unit);
    for (const id of info.tables) {
      const t = this.tables.get(id);
      if (t) this.tables.set(id, { ...t, spec: { ...t.spec, after: index } });
    }
    this.removedParas.delete(index);
  }

  async save(): Promise<Uint8Array> {
    return saveOdf(this.pkg);
  }

  // ---- editing -------------------------------------------------------------

  private editParagraph(edit: ParagraphEdit): void {
    const el = this.paras.get(edit.index);
    if (!el) throw new Error(`Paragraph ${edit.index} does not exist`);
    let orig = this.originals.get(edit.index);
    if (!orig) {
      orig = { el: el.cloneNode(true) as Element, text: odfText(el), block: this.naturalBlock(el) };
      if (edit.oldFmt) orig.fmt = edit.oldFmt;
      this.originals.set(edit.index, orig);
    }
    const back = edit.fmt || orig.fmt ? fmtEquals(edit.fmt, orig.fmt) : edit.text === orig.text;
    if (back) {
      // the original element (with all its own formatting) in the original structure
      const clone = orig.el.cloneNode(true) as Element;
      el.parentNode!.replaceChild(clone, el);
      this.paras.set(edit.index, clone);
      this.placeBlock(edit.index, orig.block);
      return;
    }
    if (!edit.fmt) {
      setOdfText(this.pkg.content, el, edit.text);
      return;
    }
    if (edit.fmt.block) this.placeBlock(edit.index, edit.fmt.block);
    this.writeRuns(this.paras.get(edit.index)!, edit.fmt.runs);
    this.applyAlign(this.paras.get(edit.index)!, edit.fmt);
  }

  /** Paragraph alignment through a derived automatic style. */
  private applyAlign(el: Element, fmt: ParaFmt): void {
    const doc = this.pkg.content;
    if (fmt.align) {
      const cur = el.getAttributeNS(NS.text, 'style-name') ?? '';
      const base = cur.replace(ALIGN_SUFFIX, '');
      const name = this.derivedStyle(base, `_VsofxA${fmt.align[0]}`, 'paragraph', 'paragraph-properties', (p) =>
        p.setAttributeNS(NS.fo, `${prefixFor(doc, NS.fo, 'fo')}:text-align`, fmt.align === 'left' ? 'start' : fmt.align === 'right' ? 'end' : fmt.align!),
      );
      el.setAttributeNS(NS.text, `${pfx(el)}style-name`, name);
    }
  }

  /** `<base><suffix>`: a copy of an automatic style (or a child of a common one) with one property changed. */
  private derivedStyle(base: string, suffix: string, family: string, propsName: string, set: (props: Element) => void): string {
    const doc = this.pkg.content;
    const name = `${base || (family === 'paragraph' ? 'Standard' : 'Default')}${suffix}`;
    const all = descendantsNS(doc, NS.style, 'style');
    if (all.some((s) => s.getAttributeNS(NS.style, 'name') === name)) return name;
    const sp = prefixFor(doc, NS.style, 'style');
    const auto = this.automaticStyles();
    const src = base ? all.find((s) => s.getAttributeNS(NS.style, 'name') === base) : undefined;
    let s: Element;
    if (src && src.parentNode === auto) {
      s = src.cloneNode(true) as Element;
      s.setAttributeNS(NS.style, `${sp}:name`, name);
    } else {
      s = doc.createElementNS(NS.style, `${sp}:style`);
      s.setAttributeNS(NS.style, `${sp}:name`, name);
      s.setAttributeNS(NS.style, `${sp}:family`, family);
      if (base) s.setAttributeNS(NS.style, `${sp}:parent-style-name`, base);
    }
    let props = descendantsNS(s, NS.style, propsName)[0];
    if (!props) {
      props = doc.createElementNS(NS.style, `${sp}:${propsName}`);
      s.insertBefore(props, s.firstChild);
    }
    set(props);
    auto.appendChild(s);
    return name;
  }

  private automaticStyles(): Element {
    this.ensureStyle('VsofxCell'); // creates office:automatic-styles and the namespaces if needed
    const root = this.pkg.content.documentElement;
    return Array.from(root.childNodes).find((n) => isEl(n, NS.office, 'automatic-styles')) as Element;
  }

  private insertParagraph(index: number, after: number, fmt: ParaFmt): void {
    if (!isNewParaIndex(index) || this.paras.has(index)) throw new Error(`Paragraph ${index} cannot be added`);
    const anchor = this.paras.get(after);
    if (!anchor) throw new Error(`Paragraph ${after} does not exist`);
    const doc = this.pkg.content;
    const p = doc.createElementNS(NS.text, `${anchor.prefix ? anchor.prefix + ':' : 'text:'}p`);
    const style = anchor.getAttributeNS(NS.text, 'style-name');
    if (anchor.localName === 'p' && style) p.setAttributeNS(NS.text, `${pfx(anchor)}style-name`, style);
    const item = anchor.parentNode as Element;
    if (isEl(item, NS.text, 'list-item') && firstEl(item) === anchor) {
      // a new item of the same list
      const li = doc.createElementNS(NS.text, item.tagName);
      li.appendChild(p);
      item.parentNode!.insertBefore(li, item.nextSibling);
    } else anchor.parentNode!.insertBefore(p, anchor.nextSibling);
    this.paras.set(index, p);
    if (fmt.block) this.placeBlock(index, fmt.block);
    this.writeRuns(this.paras.get(index)!, fmt.runs);
    this.applyAlign(this.paras.get(index)!, fmt);
  }

  private removeParagraph(index: number): void {
    const el = this.paras.get(index);
    if (!el || !isNewParaIndex(index)) throw new Error(`Paragraph ${index} cannot be removed`);
    const parent = el.parentNode as Element;
    parent.removeChild(el);
    if (isEl(parent, NS.text, 'list-item') && !firstEl(parent)) {
      const list = parent.parentNode as Element;
      list.removeChild(parent);
      if (!firstEl(list)) list.parentNode?.removeChild(list);
    }
    this.paras.delete(index);
    this.originals.delete(index);
  }

  private insertTable(spec: TableSpec): void {
    const anchor = this.paras.get(spec.after);
    if (!anchor) throw new Error(`Paragraph ${spec.after} does not exist`);
    if (this.tables.has(spec.id)) return;
    const doc = this.pkg.content;
    const tp = prefixFor(doc, NS.table, 'table');
    const xp = prefixFor(doc, NS.text, 'text');
    const t = (local: string) => doc.createElementNS(NS.table, `${tp}:${local}`);
    const table = t('table');
    table.setAttributeNS(NS.table, `${tp}:name`, `VsofxTable${spec.id}`);
    table.setAttributeNS(NS.table, `${tp}:style-name`, this.ensureStyle('VsofxTable'));
    const col = t('table-column');
    col.setAttributeNS(NS.table, `${tp}:number-columns-repeated`, String(spec.cols));
    table.appendChild(col);
    for (let r = 0; r < spec.rows; r++) {
      const row = t('table-row');
      for (let c = 0; c < spec.cols; c++) {
        const cell = t('table-cell');
        cell.setAttributeNS(NS.table, `${tp}:style-name`, this.ensureStyle('VsofxCell'));
        cell.setAttributeNS(NS.office, `${prefixFor(doc, NS.office, 'office')}:value-type`, 'string');
        const p = doc.createElementNS(NS.text, `${xp}:p`);
        cell.appendChild(p);
        row.appendChild(cell);
        this.paras.set(cellIndex(spec, r, c), p);
      }
      table.appendChild(row);
    }
    // tables cannot be inside a paragraph or a list: after the anchor's top-level list
    let ref: Element = anchor;
    while (ref.parentNode && (isEl(ref.parentNode, NS.text, 'list-item') || isEl(ref.parentNode, NS.text, 'list'))) ref = ref.parentNode as Element;
    ref.parentNode!.insertBefore(table, ref.nextSibling);
    this.tables.set(spec.id, { spec, el: table });
  }

  private removeTable(id: number): void {
    const t = this.tables.get(id);
    if (!t) return;
    t.el.parentNode?.removeChild(t.el);
    for (let r = 0; r < t.spec.rows; r++) {
      for (let c = 0; c < t.spec.cols; c++) {
        this.paras.delete(cellIndex(t.spec, r, c));
        this.originals.delete(cellIndex(t.spec, r, c));
      }
    }
    this.tables.delete(id);
  }

  /** Block kind of a paragraph as it is in the document. */
  private naturalBlock(el: Element): BlockKind {
    if (el.localName === 'h') return `h${outlineLevel(el)}` as BlockKind;
    const item = el.parentNode;
    if (isEl(item, NS.text, 'list-item')) return this.listKindOf(item!.parentNode as Element);
    if ((el.getAttributeNS(NS.text, 'style-name') ?? '').replace(ALIGN_SUFFIX, '') === 'VsofxQuote') return 'quote';
    return 'p';
  }

  private listKindOf(list: Element): 'ul' | 'ol' {
    for (let l: Element | null = list; l; l = l.parentNode as Element | null) {
      if (!isEl(l, NS.text, 'list')) continue;
      const k = this.listKinds.get(l.getAttributeNS(NS.text, 'style-name') ?? '');
      if (k) return k;
    }
    return 'ul';
  }

  /** Puts paragraph `index` into the structure of `kind` (heading, list item, quote, paragraph). */
  private placeBlock(index: number, kind: BlockKind): void {
    let el = this.paras.get(index)!;
    if (isCellIndex(index)) return; // cells of added tables stay paragraphs
    const doc = this.pkg.content;
    const xp = pfx(el);
    const isList = kind === 'ul' || kind === 'ol';
    const item = el.parentNode as Element;
    const inList = isEl(item, NS.text, 'list-item');
    if (inList && (!isList || this.listKindOf(item.parentNode as Element) !== kind)) {
      el = this.leaveList(el);
    }
    // element type and paragraph style
    const heading = /^h[1-6]$/.test(kind);
    if (heading !== (el.localName === 'h')) {
      const next = doc.createElementNS(NS.text, `${xp}${heading ? 'h' : 'p'}`);
      for (const a of Array.from(el.attributes)) {
        if (a.namespaceURI === NS.text && (a.localName === 'outline-level' || a.localName === 'style-name')) continue;
        next.setAttributeNS(a.namespaceURI, a.name, a.value);
      }
      while (el.firstChild) next.appendChild(el.firstChild);
      el.parentNode!.replaceChild(next, el);
      el = next;
    }
    if (heading) {
      const n = Number(kind[1]);
      el.setAttributeNS(NS.text, `${xp}outline-level`, String(n));
      el.setAttributeNS(NS.text, `${xp}style-name`, this.ensureStyle(`VsofxH${n}`));
    } else if (kind === 'quote') {
      el.setAttributeNS(NS.text, `${xp}style-name`, this.ensureStyle('VsofxQuote'));
    } else if (/^Vsofx(H\d|Quote)(_VsofxA[lcrj])?$/.test(el.getAttributeNS(NS.text, 'style-name') ?? '')) {
      el.removeAttributeNS(NS.text, 'style-name');
    }
    if (isList && !isEl(el.parentNode, NS.text, 'list-item')) el = this.joinList(el, kind);
    this.paras.set(index, el);
  }

  /** Wraps a paragraph into a list item: appended to a directly preceding list of the same kind, or a new list. */
  private joinList(el: Element, kind: 'ul' | 'ol'): Element {
    const doc = this.pkg.content;
    const xp = pfx(el);
    const styleName = this.ensureStyle(kind === 'ul' ? 'VsofxUL' : 'VsofxOL');
    let prev = el.previousSibling;
    while (prev && prev.nodeType !== 1) prev = prev.previousSibling;
    const li = doc.createElementNS(NS.text, `${xp}list-item`);
    if (prev && isEl(prev, NS.text, 'list') && (prev as Element).getAttributeNS(NS.text, 'style-name') === styleName) {
      el.parentNode!.removeChild(el);
      prev.appendChild(li);
    } else {
      const list = doc.createElementNS(NS.text, `${xp}list`);
      list.setAttributeNS(NS.text, `${xp}style-name`, styleName);
      el.parentNode!.replaceChild(list, el);
      list.appendChild(li);
    }
    li.appendChild(el);
    return el;
  }

  /** Moves a paragraph out of its (top-level) list, splitting the list around it. */
  private leaveList(el: Element): Element {
    const item = el.parentNode as Element;
    const list = item.parentNode as Element;
    const tail = list.cloneNode(false) as Element;
    tail.removeAttributeNS('http://www.w3.org/XML/1998/namespace', 'id');
    for (const a of Array.from(tail.attributes)) if (a.localName === 'id') tail.removeAttribute(a.name);
    while (item.nextSibling) tail.appendChild(item.nextSibling);
    const host = list.parentNode!;
    host.insertBefore(el, list.nextSibling);
    if (firstEl(tail)) host.insertBefore(tail, el.nextSibling);
    if (!firstEl(item)) list.removeChild(item);
    if (!firstEl(list)) host.removeChild(list);
    return el;
  }

  /** Replaces the paragraph's content by formatted runs. */
  private writeRuns(el: Element, runs: Run[]): void {
    const doc = this.pkg.content;
    const xp = pfx(el);
    while (el.firstChild) el.removeChild(el.firstChild);
    for (const run of runs) {
      let target: Element = el;
      if (run.href) {
        const a = doc.createElementNS(NS.text, `${xp}a`);
        const xl = prefixFor(doc, XLINK, 'xlink');
        if (!doc.documentElement.hasAttribute(`xmlns:${xl}`)) doc.documentElement.setAttribute(`xmlns:${xl}`, XLINK);
        a.setAttributeNS(XLINK, `${xl}:type`, 'simple');
        a.setAttributeNS(XLINK, `${xl}:href`, run.href);
        el.appendChild(a);
        target = a;
      }
      const flags = FLAG_KEYS.filter((k) => run[k]);
      if (flags.length) {
        const span = doc.createElementNS(NS.text, `${xp}span`);
        span.setAttributeNS(NS.text, `${xp}style-name`, this.ensureStyle(`VsofxT_${flags.join('_')}`));
        target.appendChild(span);
        target = span;
      }
      const tmp = doc.createElementNS(NS.text, `${xp}span`);
      setOdfText(doc, tmp, run.t);
      while (tmp.firstChild) target.appendChild(tmp.firstChild);
    }
  }

  /** Adds one of the editor's automatic styles to content.xml (once) and returns its name. */
  private ensureStyle(name: string): string {
    if (this.created.has(name)) return name;
    const doc = this.pkg.content;
    const root = doc.documentElement;
    const op = prefixFor(doc, NS.office, 'office');
    const sp = prefixFor(doc, NS.style, 'style');
    const fp = prefixFor(doc, NS.fo, 'fo');
    const tp = prefixFor(doc, NS.text, 'text');
    if (!root.hasAttribute(`xmlns:${sp}`)) root.setAttribute(`xmlns:${sp}`, NS.style);
    if (!root.hasAttribute(`xmlns:${fp}`)) root.setAttribute(`xmlns:${fp}`, NS.fo);
    let auto = Array.from(root.childNodes).find((n) => isEl(n, NS.office, 'automatic-styles')) as Element | undefined;
    if (!auto) {
      auto = doc.createElementNS(NS.office, `${op}:automatic-styles`);
      const bodyEl = Array.from(root.childNodes).find((n) => isEl(n, NS.office, 'body')) ?? null;
      root.insertBefore(auto, bodyEl);
    }
    const exists = Array.from(auto.childNodes).some((n) => n.nodeType === 1 && (n as Element).getAttributeNS(NS.style, 'name') === name);
    this.created.add(name);
    if (exists) return name;
    const s = (local: string, attrs: Record<string, string> = {}, ns: string = NS.style, p = sp) => {
      const e = doc.createElementNS(ns, `${p}:${local}`);
      for (const [k, v] of Object.entries(attrs)) {
        const [kp, kl] = k.split(':') as [string, string];
        const kns = kp === 'fo' ? NS.fo : kp === 'text' ? NS.text : kp === 'table' ? NS.table : NS.style;
        const kpre = kp === 'fo' ? fp : kp === 'text' ? tp : kp === 'table' ? prefixFor(doc, NS.table, 'table') : sp;
        e.setAttributeNS(kns, `${kpre}:${kl}`, v);
      }
      return e;
    };
    let el: Element;
    const h = /^VsofxH(\d)$/.exec(name);
    if (name.startsWith('VsofxT_')) {
      const f = new Set(name.slice(7).split('_'));
      const props: Record<string, string> = {};
      if (f.has('b')) Object.assign(props, { 'fo:font-weight': 'bold', 'style:font-weight-asian': 'bold', 'style:font-weight-complex': 'bold' });
      if (f.has('i')) Object.assign(props, { 'fo:font-style': 'italic', 'style:font-style-asian': 'italic', 'style:font-style-complex': 'italic' });
      if (f.has('u')) Object.assign(props, { 'style:text-underline-style': 'solid', 'style:text-underline-width': 'auto', 'style:text-underline-color': 'font-color' });
      if (f.has('s')) Object.assign(props, { 'style:text-line-through-style': 'solid' });
      if (f.has('code')) Object.assign(props, { 'fo:font-family': "'Courier New'", 'style:font-family-generic': 'modern', 'style:font-pitch': 'fixed', 'fo:background-color': '#f2f2f2' });
      if (f.has('sup')) props['style:text-position'] = 'super 58%';
      else if (f.has('sub')) props['style:text-position'] = 'sub 58%';
      el = s('style', { 'style:name': name, 'style:family': 'text' });
      el.appendChild(s('text-properties', props));
    } else if (h) {
      const n = Number(h[1]);
      const pt = `${HEADING_PT[n - 1]}pt`;
      el = s('style', { 'style:name': name, 'style:family': 'paragraph', 'style:default-outline-level': String(n) });
      el.appendChild(s('paragraph-properties', { 'fo:margin-top': '0.42cm', 'fo:margin-bottom': '0.21cm', 'fo:keep-with-next': 'always' }));
      el.appendChild(s('text-properties', { 'fo:font-size': pt, 'fo:font-weight': 'bold', 'style:font-size-asian': pt, 'style:font-weight-asian': 'bold', 'style:font-size-complex': pt, 'style:font-weight-complex': 'bold' }));
    } else if (name === 'VsofxQuote') {
      el = s('style', { 'style:name': name, 'style:family': 'paragraph' });
      el.appendChild(s('paragraph-properties', { 'fo:margin-left': '1cm', 'fo:margin-right': '1cm', 'fo:margin-top': '0.2cm', 'fo:margin-bottom': '0.2cm' }));
      el.appendChild(s('text-properties', { 'fo:font-style': 'italic', 'fo:color': '#595959' }));
    } else if (name === 'VsofxUL' || name === 'VsofxOL') {
      el = s('list-style', { 'style:name': name }, NS.text, tp);
      for (let l = 1; l <= 10; l++) {
        const lvl =
          name === 'VsofxUL'
            ? s('list-level-style-bullet', { 'text:level': String(l), 'text:bullet-char': ['•', '◦', '▪'][(l - 1) % 3]! }, NS.text, tp)
            : s('list-level-style-number', { 'text:level': String(l), 'style:num-suffix': '.', 'style:num-format': '1' }, NS.text, tp);
        const props = s('list-level-properties', { 'text:list-level-position-and-space-mode': 'label-alignment' });
        props.appendChild(s('list-level-label-alignment', { 'text:label-followed-by': 'listtab', 'text:list-tab-stop-position': `${(0.635 * (l + 1)).toFixed(3)}cm`, 'fo:text-indent': '-0.635cm', 'fo:margin-left': `${(0.635 * (l + 1)).toFixed(3)}cm` }));
        lvl.appendChild(props);
        el.appendChild(lvl);
      }
    } else if (name === 'VsofxTable') {
      el = s('style', { 'style:name': name, 'style:family': 'table' });
      el.appendChild(s('table-properties', { 'style:width': '17cm', 'table:align': 'margins' }));
    } else {
      el = s('style', { 'style:name': name, 'style:family': 'table-cell' });
      el.appendChild(s('table-cell-properties', { 'fo:padding': '0.1cm', 'fo:border': '0.5pt solid #000000' }));
    }
    auto.appendChild(el);
    return name;
  }

  // ---- rendering ---------------------------------------------------------

  private readStyles(): void {
    this.styles = readStyles(this.pkg.content);
    this.paraAlign = new Map();
    for (const s of descendantsNS(this.pkg.content, NS.style, 'style')) {
      const name = s.getAttributeNS(NS.style, 'name');
      if (!name) continue;
      const ta = descendantsNS(s, NS.style, 'paragraph-properties')[0]?.getAttributeNS(NS.fo, 'text-align');
      const h: HAlign | undefined = ta === 'center' ? 'center' : ta === 'justify' ? 'justify' : ta === 'end' || ta === 'right' ? 'right' : ta === 'start' || ta === 'left' ? 'left' : undefined;
      if (h && s.getAttributeNS(NS.style, 'family') === 'paragraph') this.paraAlign.set(name, h);
    }
    this.listKinds = new Map();
    for (const ls of descendantsNS(this.pkg.content, NS.text, 'list-style')) {
      const name = ls.getAttributeNS(NS.style, 'name');
      const first = Array.from(ls.childNodes).find((n) => n.nodeType === 1) as Element | undefined;
      if (name) this.listKinds.set(name, first?.localName === 'list-level-style-number' ? 'ol' : 'ul');
    }
  }

  private render(body: Element): string {
    const index = new Map<Element, number>();
    for (const [i, p] of this.paras) index.set(p, i);
    const tableIds = new Map<Element, TableSpec>();
    for (const t of this.tables.values()) tableIds.set(t.el, t.spec);
    const blocks = (parent: Node): string => {
      let out = '';
      for (let c = parent.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        const e = c as Element;
        if (e.namespaceURI === NS.text && (e.localName === 'p' || e.localName === 'h')) {
          const i = index.get(e);
          const tag = e.localName === 'h' ? `h${outlineLevel(e)}` : 'p';
          let attr = i === undefined ? '' : ` data-p="${i}"`;
          const sn = e.getAttributeNS(NS.text, 'style-name') ?? '';
          if (sn.replace(ALIGN_SUFFIX, '') === 'VsofxQuote') attr += ' data-block="quote"';
          const al = this.paraAlign.get(sn);
          if (al && i !== undefined) attr += ` data-align="${al}"`;
          out += `<${tag}${attr}>${this.inline(e, this.styleOf(e))}</${tag}>`;
        } else if (isEl(e, NS.text, 'list')) {
          const tag = this.listKindOf(e);
          out += `<${tag}>${blocks(e)}</${tag}>`;
        } else if (isEl(e, NS.text, 'list-item') || isEl(e, NS.text, 'list-header')) {
          out += `<li>${blocks(e)}</li>`;
        } else if (isEl(e, NS.table, 'table')) {
          const spec = tableIds.get(e);
          const attr = spec ? ` class="ins-table" data-table="${spec.id}" data-after="${spec.after}" data-rows="${spec.rows}" data-cols="${spec.cols}"` : '';
          out += `<table${attr}><tbody>${blocks(e)}</tbody></table>`;
        } else if (isEl(e, NS.table, 'table-row')) {
          out += `<tr>${blocks(e)}</tr>`;
        } else if (isEl(e, NS.table, 'table-cell')) {
          const span = Number(e.getAttributeNS(NS.table, 'number-columns-spanned') || 1);
          out += `<td${span > 1 && span < 1000 ? ` colspan="${span}"` : ''}>${blocks(e)}</td>`;
        } else if (isEl(e, NS.draw, 'frame') || isEl(e, NS.office, 'annotation') || isEl(e, NS.text, 'note')) {
          continue;
        } else {
          out += blocks(e); // sections, table groups, indexes, ...
        }
      }
      return out;
    };
    return blocks(body);
  }

  private styleOf(el: Element): InlineStyle | undefined {
    const name = el.getAttributeNS(NS.text, 'style-name');
    return name ? this.styles.get(name) : undefined;
  }

  private inline(el: Element, inherited?: InlineStyle): string {
    let out = '';
    for (let c = el.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 3 || c.nodeType === 4) {
        out += escapeHtml(c.nodeValue ?? '');
        continue;
      }
      if (c.nodeType !== 1) continue;
      const e = c as Element;
      if (e.namespaceURI === NS.text) {
        if (e.localName === 's') {
          out += '&nbsp;'.repeat(Math.min(1000, Number(e.getAttributeNS(NS.text, 'c') || 1)));
          continue;
        }
        if (e.localName === 'tab') {
          out += '<span class="tab">\t</span>';
          continue;
        }
        if (e.localName === 'line-break') {
          out += '<br>';
          continue;
        }
        if (e.localName === 'note' || e.localName === 'soft-page-break') continue;
        if (e.localName === 'span') {
          const s = this.styleOf(e);
          const merged = { ...inherited, ...s };
          let html = this.inline(e, merged);
          if (s?.code && !inherited?.code) html = `<code>${html}</code>`;
          if (s?.sub && !inherited?.sub) html = `<sub>${html}</sub>`;
          if (s?.sup && !inherited?.sup) html = `<sup>${html}</sup>`;
          if (s?.strike && !inherited?.strike) html = `<s>${html}</s>`;
          if (s?.underline && !inherited?.underline) html = `<u>${html}</u>`;
          if (s?.italic && !inherited?.italic) html = `<em>${html}</em>`;
          if (s?.bold && !inherited?.bold) html = `<strong>${html}</strong>`;
          out += html;
          continue;
        }
        if (e.localName === 'a') {
          const href = e.getAttributeNS(XLINK, 'href') ?? '';
          const inner = this.inline(e, inherited);
          out += SAFE_HREF.test(href) ? `<a data-href="${escapeHtml(href)}">${inner}</a>` : inner;
          continue;
        }
      }
      if (isEl(e, NS.office, 'annotation') || isEl(e, NS.draw, 'frame')) continue;
      out += this.inline(e, inherited);
    }
    return out;
  }
}

const pfx = (el: Element) => (el.prefix ? `${el.prefix}:` : '');

function outlineLevel(el: Element): number {
  return Math.min(6, Math.max(1, Number(el.getAttributeNS(NS.text, 'outline-level') || 1)));
}

/** The prefix the document already uses for a namespace. */
function prefixFor(doc: Document, ns: string, fallback: string): string {
  for (const a of Array.from(doc.documentElement.attributes)) {
    if (a.value === ns && a.name.startsWith('xmlns:')) return a.name.slice(6);
  }
  return fallback;
}

function readStyles(doc: Document): Map<string, InlineStyle> {
  const map = new Map<string, InlineStyle>();
  for (const style of descendantsNS(doc, NS.style, 'style')) {
    const name = style.getAttributeNS(NS.style, 'name');
    const props = descendantsNS(style, NS.style, 'text-properties')[0];
    if (!name || !props) continue;
    const weight = props.getAttributeNS(NS.fo, 'font-weight');
    const fstyle = props.getAttributeNS(NS.fo, 'font-style');
    const underline = props.getAttributeNS(NS.style, 'text-underline-style');
    const strike = props.getAttributeNS(NS.style, 'text-line-through-style');
    const family = `${props.getAttributeNS(NS.fo, 'font-family') ?? ''} ${props.getAttributeNS(NS.style, 'font-name') ?? ''}`;
    const pos = props.getAttributeNS(NS.style, 'text-position') ?? '';
    const s: InlineStyle = {};
    if (weight === 'bold' || /^[6-9]00$/.test(weight ?? '')) s.bold = true;
    if (fstyle === 'italic' || fstyle === 'oblique') s.italic = true;
    if (underline && underline !== 'none') s.underline = true;
    if (strike && strike !== 'none') s.strike = true;
    if (/courier|mono|consolas/i.test(family) || props.getAttributeNS(NS.style, 'font-pitch') === 'fixed') s.code = true;
    if (/^(super|[1-9])/.test(pos)) s.sup = true;
    else if (/^(sub|-)/.test(pos)) s.sub = true;
    map.set(name, s);
  }
  return map;
}

function nextEl(n: Node): Element | null {
  for (let c = n.nextSibling; c; c = c.nextSibling) if (c.nodeType === 1) return c as Element;
  return null;
}

function firstEl(n: Node): Element | null {
  for (let c = n.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) return c as Element;
  return null;
}
