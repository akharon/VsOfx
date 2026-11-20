// Diagram fences (```seldia, ```mermaid ...) are drawn by whichever extension extends VS Code's
// Markdown renderer (e.g. SelDiA). The host renders the fence with `markdown.api.render`; the webview
// keeps only the first <svg> and shows it as an <img> (an SVG image cannot run scripts or load URLs).

export const DIAGRAM_LANG = /^(seldia(-[\w-]+)?|mermaid|plantuml)$/;

type Poster = (msg: { type: 'renderDiagram'; id: number; lang: string; source: string }) => void;

let post: Poster | undefined;
let nextId = 1;
const waiting = new Map<number, (svg: string | null) => void>();
const cache = new Map<string, Promise<string | null>>();

export function setDiagramPoster(p: Poster): void {
  post = p;
}

/** Called by main.ts for each host reply. */
export function onDiagram(id: number, html: string | null): void {
  const done = waiting.get(id);
  waiting.delete(id);
  done?.(html === null ? null : svgDataUrl(html));
}

export function requestDiagram(lang: string, source: string): Promise<string | null> {
  const key = `${lang}\n${source}`;
  let p = cache.get(key);
  if (!p) {
    const send = post;
    if (!send) return Promise.resolve(null);
    p = new Promise((resolve) => {
      const id = nextId++;
      waiting.set(id, resolve);
      send({ type: 'renderDiagram', id, lang, source });
    });
    cache.set(key, p);
    if (cache.size > 200) cache.delete(cache.keys().next().value!);
  }
  return p;
}

function svgDataUrl(html: string): string | null {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  // The diagram itself, not a sprite sheet of shared symbols that may come first.
  const all = Array.from(doc.querySelectorAll('svg')).filter((s) => !s.parentElement?.closest('svg'));
  const svg =
    doc.querySelector<SVGSVGElement>('.seldia-diagram-svg > svg') ??
    all.sort((a, b) => b.outerHTML.length - a.outerHTML.length).find((s) => s.querySelector('g, path, rect, text'));
  if (!svg) return null;
  // Symbols the diagram <use>s from other SVGs on the page go into the image too.
  for (const other of all) {
    if (other === svg) continue;
    const defs = doc.createElementNS('http://www.w3.org/2000/svg', 'defs');
    for (const d of Array.from(other.querySelectorAll('symbol, defs > *'))) defs.appendChild(d.cloneNode(true));
    if (defs.childNodes.length) svg.insertBefore(defs, svg.firstChild);
  }
  // CSS a preview stylesheet would have supplied: keep any <style> blocks of the rendering.
  for (const st of Array.from(doc.querySelectorAll('style'))) if (!svg.contains(st)) svg.insertBefore(st.cloneNode(true), svg.firstChild);
  if (!svg.getAttribute('xmlns')) svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
  const text = new XMLSerializer().serializeToString(svg);
  return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(text)}`;
}

/** Puts a rendered image into each diagram <pre> of `root` (the code stays, hidden, for editing). */
export function attachDiagrams(root: HTMLElement): void {
  for (const pre of Array.from(root.querySelectorAll<HTMLElement>('pre'))) {
    const code = pre.querySelector('code');
    const lang = /language-(\S+)/.exec(code?.className ?? '')?.[1] ?? '';
    if (!code || !DIAGRAM_LANG.test(lang)) continue;
    const source = code.textContent ?? '';
    void requestDiagram(lang, source).then((url) => {
      if (!url || !pre.isConnected || code.textContent !== source) return;
      const img = document.createElement('img');
      img.className = 'md-diagram';
      img.src = url;
      img.alt = `${lang} diagram`;
      img.title = 'Click to show or hide the diagram source';
      img.setAttribute('contenteditable', 'false');
      img.addEventListener('click', () => pre.classList.toggle('show-code'));
      pre.querySelector('img.md-diagram')?.remove();
      pre.prepend(img);
      pre.classList.add('has-diagram');
    });
  }
}
