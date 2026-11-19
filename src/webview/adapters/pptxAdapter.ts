import { fromByteArray } from '../../shared/base64';
import type { ApplyResult, Edit, SlideView, SlidesModel } from '../../shared/model';
import type { FormatAdapter } from './Adapter';
import { descendantsNS, normText, parseXml } from './xml';
import { optionalText, readZip, text, writeZip, type ZipEntries } from './zip';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const MAX_IMAGE_BYTES = 4 * 1024 * 1024;
const MIME: Record<string, string> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp' };

/** Slide viewer (read-only): text, notes and the first picture of each slide. */
export class PptxAdapter implements FormatAdapter {
  readonly readonly = true;
  private bytes!: Uint8Array;

  async load(bytes: Uint8Array): Promise<SlidesModel> {
    this.bytes = bytes;
    const entries = readZip(bytes);
    const slides: SlideView[] = [];
    const paths = slidePaths(entries);
    paths.forEach((path, i) => {
      const doc = parseXml(text(entries, path));
      const rels = relationships(entries, relsPathOf(path));
      const view: SlideView = { number: i + 1, title: '', paragraphs: [], notes: '' };

      for (const sp of descendantsNS(doc, P, 'sp')) {
        const ph = descendantsNS(sp, P, 'ph')[0];
        const type = ph?.getAttribute('type') ?? '';
        const lines = shapeParagraphs(sp);
        if (type === 'title' || type === 'ctrTitle') view.title = view.title || lines.join(' ');
        else if (type !== 'sldNum' && type !== 'ftr' && type !== 'dt') view.paragraphs.push(...lines);
      }
      for (const tbl of descendantsNS(doc, A, 'tbl')) {
        for (const tr of descendantsNS(tbl, A, 'tr')) {
          const cells = descendantsNS(tr, A, 'tc').map((tc) => descendantsNS(tc, A, 'p').map(paraText).join(' ').trim());
          const line = cells.join(' | ').trim();
          if (line.replace(/\|/g, '').trim()) view.paragraphs.push(line);
        }
      }

      const notesTarget = rels.find((r) => r.type.endsWith('/notesSlide'))?.target;
      if (notesTarget) {
        const notesXml = optionalText(entries, resolve(path, notesTarget));
        if (notesXml) {
          const notes = parseXml(notesXml);
          const lines: string[] = [];
          for (const sp of descendantsNS(notes, P, 'sp')) {
            const ph = descendantsNS(sp, P, 'ph')[0];
            if (ph?.getAttribute('type') === 'body') lines.push(...shapeParagraphs(sp));
          }
          view.notes = lines.join('\n');
        }
      }

      const blip = descendantsNS(doc, A, 'blip')[0];
      const rid = blip?.getAttributeNS(R, 'embed');
      const target = rid ? rels.find((r) => r.id === rid)?.target : undefined;
      if (target) {
        const mediaPath = resolve(path, target);
        const data = entries[mediaPath];
        const mime = MIME[mediaPath.split('.').pop()?.toLowerCase() ?? ''];
        if (data && mime && data.length <= MAX_IMAGE_BYTES) view.imageDataUri = `data:${mime};base64,${fromByteArray(data)}`;
      }
      slides.push(view);
    });
    return { type: 'slides', slides };
  }

  applyEdit(_edit: Edit): ApplyResult {
    throw new Error('Presentations are read-only');
  }

  async save(): Promise<Uint8Array> {
    return writeZip(readZip(this.bytes), {});
  }
}

function paraText(p: Element): string {
  let out = '';
  const walk = (n: Node) => {
    for (let c = n.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      const el = c as Element;
      if (el.namespaceURI === A && el.localName === 't') out += el.textContent ?? '';
      else if (el.namespaceURI === A && el.localName === 'br') out += ' ';
      else walk(el);
    }
  };
  walk(p);
  return out;
}

function shapeParagraphs(sp: Element): string[] {
  return descendantsNS(sp, A, 'p')
    .map((p) => normText(paraText(p)))
    .filter((s) => s.length > 0);
}

interface Rel {
  id: string;
  type: string;
  target: string;
}

function relsPathOf(part: string): string {
  const i = part.lastIndexOf('/');
  return `${part.slice(0, i + 1)}_rels/${part.slice(i + 1)}.rels`;
}

function relationships(entries: ZipEntries, relsPath: string): Rel[] {
  const xml = optionalText(entries, relsPath);
  if (!xml) return [];
  return descendantsNS(parseXml(xml), REL_NS, 'Relationship')
    .filter((r) => r.getAttribute('TargetMode') !== 'External')
    .map((r) => ({ id: r.getAttribute('Id') ?? '', type: r.getAttribute('Type') ?? '', target: r.getAttribute('Target') ?? '' }));
}

/** Resolves a relationship target relative to the part that owns it. */
function resolve(fromPart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = fromPart.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.' && seg !== '') parts.push(seg);
  }
  return parts.join('/');
}

/** Slide part paths in presentation order. */
function slidePaths(entries: ZipEntries): string[] {
  const pres = parseXml(text(entries, 'ppt/presentation.xml'));
  const rels = relationships(entries, 'ppt/_rels/presentation.xml.rels');
  const out: string[] = [];
  for (const id of descendantsNS(pres, P, 'sldId')) {
    const rid = id.getAttributeNS(R, 'id');
    const target = rels.find((r) => r.id === rid)?.target;
    if (target) {
      const path = resolve('ppt/presentation.xml', target);
      if (entries[path]) out.push(path);
    }
  }
  return out;
}
