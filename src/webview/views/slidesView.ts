import type { SlidesModel } from '../../shared/model';
import { createFindBar, domFinder } from '../ui/findBar';
import { addZoom, createToolbar, type Toolbar } from '../ui/toolbar';
import type { View } from './types';

/** Read-only list of slides (title, text, first picture, notes). Uses textContent only. */
export function renderSlidesView(container: HTMLElement, model: SlidesModel, tb: Toolbar = createToolbar()): View {
  container.replaceChildren();
  const banner = document.createElement('div');
  banner.className = 'banner';
  banner.textContent = 'Presentations are shown read-only (text and first picture of each slide).';
  const list = document.createElement('div');
  list.className = 'slides';
  container.append(banner, list);

  const finder = createFindBar(tb.extra, domFinder(list));
  tb.button('', 'Find in presentation (Ctrl+F)', () => finder.toggle(), 'search');
  tb.separator();
  addZoom(tb, (z) => list.style.setProperty('--zoom', String(z)));
  const onFind = (ev: KeyboardEvent) => {
    if ((ev.ctrlKey || ev.metaKey) && !ev.altKey && ev.key.toLowerCase() === 'f') {
      ev.preventDefault();
      finder.open();
    }
  };
  document.addEventListener('keydown', onFind);

  if (model.slides.length === 0) {
    const empty = document.createElement('p');
    empty.textContent = 'This presentation has no slides.';
    list.appendChild(empty);
  }
  for (const slide of model.slides) {
    const card = document.createElement('section');
    card.className = 'slide';
    const head = document.createElement('h2');
    const num = document.createElement('span');
    num.className = 'slide-num';
    num.textContent = String(slide.number);
    const title = document.createElement('span');
    title.textContent = slide.title || 'Untitled slide';
    if (!slide.title) title.className = 'muted';
    head.append(num, title);
    card.appendChild(head);
    if (slide.paragraphs.length) {
      const ul = document.createElement('ul');
      for (const p of slide.paragraphs) {
        const li = document.createElement('li');
        li.textContent = p;
        ul.appendChild(li);
      }
      card.appendChild(ul);
    }
    if (slide.imageDataUri) {
      const img = document.createElement('img');
      img.src = slide.imageDataUri;
      img.alt = `Picture on slide ${slide.number}`;
      card.appendChild(img);
    }
    if (slide.notes) {
      const details = document.createElement('details');
      const summary = document.createElement('summary');
      summary.textContent = 'Notes';
      const pre = document.createElement('p');
      pre.className = 'notes';
      pre.textContent = slide.notes;
      details.append(summary, pre);
      card.appendChild(details);
    }
    list.appendChild(card);
  }
  return {
    applyEdit() {},
    dispose() {
      finder.close();
      document.removeEventListener('keydown', onFind);
    },
  };
}
