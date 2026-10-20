import { h } from '../ui/dom.js';
import { renderSlide } from '../render/render.js';
import { slideSize } from '../model/deck.js';

/**
 * Print (or "Save as PDF") one slide per page, at the slide's own size, so
 * the PDF is vector text and shapes rather than pictures of slides.
 */
export function printDeck(deck, { includeHidden = false } = {}) {
  const { w, h: H } = slideSize(deck);
  const host = h('div', { class: 'print-host' });
  for (const slide of deck.slides) {
    if (slide.hidden && !includeHidden) continue;
    const page = h('div', { class: 'print-page' });
    page.append(renderSlide(slide, deck));
    host.append(page);
  }
  // Slides are laid out at 1px per point; a CSS point is 96/72 px, so the slide is scaled up to fill the page.
  const style = h('style', {}, `@page { size: ${w}pt ${H}pt; margin: 0; } .print-page { width: ${w}pt; height: ${H}pt; } .print-page .slide { transform: scale(${96 / 72}); transform-origin: 0 0; }`);
  document.head.append(style);
  document.body.append(host);
  document.body.classList.add('is-printing');
  const done = () => {
    host.remove();
    style.remove();
    document.body.classList.remove('is-printing');
    window.removeEventListener('afterprint', done);
  };
  window.addEventListener('afterprint', done);
  // Let pictures decode before the print snapshot.
  const imgs = [...host.querySelectorAll('img')];
  Promise.all(imgs.map((img) => (img.complete ? null : new Promise((r) => { img.onload = img.onerror = r; })))).then(() => {
    requestAnimationFrame(() => {
      window.print();
    });
  });
}
