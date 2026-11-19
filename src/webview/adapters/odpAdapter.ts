import { fromByteArray } from '../../shared/base64';
import type { ApplyResult, Edit, SlideView, SlidesModel } from '../../shared/model';
import type { FormatAdapter } from './Adapter';
import { isEl, MIME_ODG, MIME_ODP, NS, odfText, openOdf } from './odf';
import { descendantsNS, normText } from './xml';
import { readZip, writeZip } from './zip';

const PRESENTATION_NS = 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0';
const XLINK_NS = 'http://www.w3.org/1999/xlink';
const MAX_IMAGE_BYTES = 4 * 1024 * 1024;
const MIME: Record<string, string> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml' };

/** .odp presentations and .odg drawings (read-only): text, notes and the first picture of each page. */
export class OdpAdapter implements FormatAdapter {
  readonly readonly = true;
  private bytes!: Uint8Array;

  constructor(private readonly kind: 'presentation' | 'drawing') {}

  async load(bytes: Uint8Array): Promise<SlidesModel> {
    this.bytes = bytes;
    const pkg = openOdf(bytes, this.kind === 'drawing' ? MIME_ODG : MIME_ODP);
    const body = descendantsNS(pkg.content, NS.office, this.kind)[0];
    if (!body) throw new Error(`This OpenDocument file has no ${this.kind} body`);
    const pages = Array.from(body.childNodes).filter((n): n is Element => isEl(n, NS.draw, 'page'));
    const slides = pages.map((page, i) => {
      const view: SlideView = { number: i + 1, title: '', paragraphs: [], notes: '' };
      for (const frame of Array.from(page.childNodes)) {
        if (frame.nodeType !== 1) continue;
        const el = frame as Element;
        if (el.namespaceURI === PRESENTATION_NS && el.localName === 'notes') {
          view.notes = paragraphs(el).join('\n');
          continue;
        }
        const lines = paragraphs(el);
        const cls = el.getAttributeNS(PRESENTATION_NS, 'class');
        if (cls === 'title' && !view.title) view.title = lines.join(' ');
        else if (cls !== 'page-number' && cls !== 'footer' && cls !== 'date-time') view.paragraphs.push(...lines);
      }
      if (!view.title && this.kind === 'drawing') view.title = page.getAttributeNS(NS.draw, 'name') ?? '';
      const img = descendantsNS(page, NS.draw, 'image').find((im) => !hasNotesAncestor(im));
      const href = img?.getAttributeNS(XLINK_NS, 'href') ?? '';
      const path = href.replace(/^\.\//, '');
      const data = path ? pkg.entries[path] : undefined;
      const mime = MIME[path.split('.').pop()?.toLowerCase() ?? ''];
      if (data && mime && data.length <= MAX_IMAGE_BYTES) view.imageDataUri = `data:${mime};base64,${fromByteArray(data)}`;
      return view;
    });
    return { type: 'slides', kind: this.kind, slides };
  }

  applyEdit(_edit: Edit): ApplyResult {
    throw new Error(this.kind === 'drawing' ? 'Drawings are read-only' : 'Presentations are read-only');
  }

  async save(): Promise<Uint8Array> {
    return writeZip(readZip(this.bytes), {}, { mimetypeFirst: true });
  }
}

function hasNotesAncestor(n: Node): boolean {
  for (let p = n.parentNode; p; p = p.parentNode) {
    if (p.nodeType === 1 && (p as Element).namespaceURI === PRESENTATION_NS && (p as Element).localName === 'notes') return true;
  }
  return false;
}

/** Non-empty text:p / text:h lines inside a frame or shape, in document order. */
function paragraphs(root: Element): string[] {
  const out: string[] = [];
  const walk = (n: Node) => {
    for (let c = n.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      const e = c as Element;
      if (e.namespaceURI === NS.text && (e.localName === 'p' || e.localName === 'h')) {
        const t = normText(odfText(e));
        if (t) out.push(t);
        continue;
      }
      if (e.namespaceURI === NS.office && e.localName === 'annotation') continue;
      walk(e);
    }
  };
  walk(root);
  return out;
}
