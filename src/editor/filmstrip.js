import { h, showPopover, menu } from '../ui/dom.js';
import { icon } from '../ui/icons.js';
import { renderThumb } from '../render/render.js';
import { LAYOUTS, slideTitle } from '../model/deck.js';
import { openComments } from './comments.js';

const THUMB_W = 168;

/**
 * The slide list on the left: thumbnails you can click, multi-select,
 * reorder by dragging and manage from the keyboard or a context menu.
 */
export class Filmstrip {
  constructor(app) {
    this.app = app;
    this.list = h('ol', { class: 'filmstrip-list', role: 'listbox', 'aria-label': 'Slides', 'aria-multiselectable': 'true' });
    this.addBtn = h('button', {
      type: 'button', class: 'btn filmstrip-add', 'aria-haspopup': 'menu', html: `${icon('plus')}<span>New slide</span>`,
      onclick: () => this.layoutMenu(this.addBtn),
    });
    this.el = h('nav', { class: 'filmstrip', 'aria-label': 'Slides' }, this.list, this.addBtn);
    this.items = new Map(); // slide id → li
    this.dropIndex = null;
    this.list.addEventListener('keydown', (e) => this.onKey(e));
    this.list.addEventListener('dragover', (e) => this.onDragOver(e));
    this.list.addEventListener('drop', (e) => this.onDrop(e));
    this.list.addEventListener('dragleave', (e) => {
      if (!this.list.contains(e.relatedTarget)) this.setDropIndex(null);
    });
  }

  layoutMenu(anchor, at = null) {
    showPopover(anchor, menu(LAYOUTS.map((l) => ({ label: l.name, run: () => this.app.addSlide(l.id) }))), { at, label: 'New slide layout' });
  }

  /** Rebuild the whole list (after adding, removing or reordering slides). */
  render() {
    const { deck } = this.app;
    this.items.clear();
    this.list.replaceChildren(...deck.slides.map((s, i) => this.item(s, i)));
    this.syncSelection();
  }

  item(slide, i) {
    const { deck } = this.app;
    const thumb = renderThumb(slide, deck, THUMB_W);
    const title = slideTitle(slide);
    const comments = openComments(slide);
    const commentsLabel = `${comments} open comment${comments === 1 ? '' : 's'}`;
    const li = h('li', {
      class: `film-item${slide.hidden ? ' is-hidden' : ''}`, role: 'option', tabindex: '-1', draggable: 'true',
      'aria-label': `Slide ${i + 1}${title ? `: ${title}` : ''}${slide.hidden ? ' (skipped)' : ''}${comments ? `, ${commentsLabel}` : ''}`,
      dataset: { id: slide.id },
    },
    h('span', { class: 'film-num' }, String(i + 1), slide.hidden ? h('span', { class: 'film-hidden', html: icon('eyeOff') }) : null),
    h('div', { class: 'film-thumb' }, thumb),
    comments ? h('span', { class: 'film-comments', title: commentsLabel, html: icon('comment') }, String(comments)) : null);
    let pending = false;
    li.addEventListener('pointerdown', (e) => {
      const selected = this.app.slideSelection.has(slide.id);
      if (e.button === 2 && selected) return;
      // Pressing on one of several selected slides may start dragging them all:
      // narrow the selection to it only on a plain click.
      if (selected && this.app.slideSelection.size > 1 && !e.shiftKey && !e.ctrlKey && !e.metaKey) {
        pending = true;
        return;
      }
      this.click(slide.id, e);
    });
    li.addEventListener('click', (e) => {
      if (!pending) return;
      pending = false;
      this.click(slide.id, e);
    });
    li.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      this.contextMenu({ x: e.clientX, y: e.clientY });
    });
    li.addEventListener('dragstart', (e) => {
      pending = false;
      if (!this.app.slideSelection.has(slide.id)) this.app.goToSlide(this.app.deck.slides.indexOf(slide));
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('application/x-libreslides-slides', 'move');
      li.classList.add('is-dragging');
    });
    li.addEventListener('dragend', () => {
      li.classList.remove('is-dragging');
      this.setDropIndex(null);
    });
    this.items.set(slide.id, li);
    return li;
  }

  /** Redraw one slide's thumbnail. */
  refreshThumb(slideId) {
    const li = this.items.get(slideId);
    const i = this.app.deck.slides.findIndex((s) => s.id === slideId);
    if (!li || i < 0) return;
    const fresh = this.item(this.app.deck.slides[i], i);
    li.replaceWith(fresh);
    this.syncSelection();
  }

  click(id, e) {
    const { app } = this;
    const idx = app.deck.slides.findIndex((s) => s.id === id);
    if (e.shiftKey) {
      const cur = app.slideIndex;
      const [a, b] = [Math.min(cur, idx), Math.max(cur, idx)];
      app.slideSelection = new Set(app.deck.slides.slice(a, b + 1).map((s) => s.id));
      this.syncSelection();
      return;
    }
    if (e.ctrlKey || e.metaKey) {
      const sel = new Set(app.slideSelection);
      if (sel.has(id) && sel.size > 1) sel.delete(id);
      else sel.add(id);
      app.slideSelection = sel;
      if (!sel.has(app.slide.id)) app.goToSlide(app.deck.slides.findIndex((s) => sel.has(s.id)), { keepSlideSelection: true });
      this.syncSelection();
      return;
    }
    app.goToSlide(idx);
  }

  syncSelection() {
    const { app } = this;
    for (const [id, li] of this.items) {
      const current = id === app.slide.id;
      li.classList.toggle('is-current', current);
      li.classList.toggle('is-selected', app.slideSelection.has(id));
      li.setAttribute('aria-selected', String(app.slideSelection.has(id)));
      li.tabIndex = current ? 0 : -1;
    }
  }

  focusCurrent() {
    const li = this.items.get(this.app.slide.id);
    li?.focus({ preventScroll: true });
    li?.scrollIntoView({ block: 'nearest' });
  }

  scrollToCurrent() {
    this.items.get(this.app.slide.id)?.scrollIntoView({ block: 'nearest' });
  }

  onKey(e) {
    const { app } = this;
    const mod = e.ctrlKey || e.metaKey;
    const go = (i) => {
      e.preventDefault();
      app.goToSlide(Math.max(0, Math.min(app.deck.slides.length - 1, i)));
      this.focusCurrent();
    };
    if (e.key === 'ArrowDown' && mod) { e.preventDefault(); app.moveSlides(e.shiftKey ? Infinity : 1); this.focusCurrent(); }
    else if (e.key === 'ArrowUp' && mod) { e.preventDefault(); app.moveSlides(e.shiftKey ? -Infinity : -1); this.focusCurrent(); }
    else if (e.key === 'ArrowDown' || e.key === 'ArrowRight') go(app.slideIndex + 1);
    else if (e.key === 'ArrowUp' || e.key === 'ArrowLeft') go(app.slideIndex - 1);
    else if (e.key === 'Home') go(0);
    else if (e.key === 'End') go(app.deck.slides.length - 1);
    else if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); app.deleteSlides(); this.focusCurrent(); }
    else if (e.key === 'Enter' && !mod) { e.preventDefault(); app.addSlide(); this.focusCurrent(); }
    else if (mod && e.key.toLowerCase() === 'd') { e.preventDefault(); app.duplicateSlides(); this.focusCurrent(); }
    else if (mod && e.key.toLowerCase() === 'c') { e.preventDefault(); app.copySlides(); }
    else if (mod && e.key.toLowerCase() === 'x') { e.preventDefault(); app.copySlides(); app.deleteSlides(); }
    else if (mod && e.key.toLowerCase() === 'a') {
      e.preventDefault();
      app.slideSelection = new Set(app.deck.slides.map((s) => s.id));
      this.syncSelection();
    } else return;
    e.stopPropagation();
  }

  contextMenu(at) {
    const { app } = this;
    const many = app.slideSelection.size > 1;
    const allHidden = [...app.slideSelection].every((id) => app.deck.slides.find((s) => s.id === id)?.hidden);
    showPopover(null, menu([
      { label: 'New slide', icon: 'plus', shortcut: 'Mod-M', run: () => app.addSlide() },
      { label: many ? 'Duplicate slides' : 'Duplicate slide', icon: 'duplicate', shortcut: 'Mod-D', run: () => app.duplicateSlides() },
      { label: many ? 'Delete slides' : 'Delete slide', icon: 'trash', shortcut: 'Delete', run: () => app.deleteSlides() },
      'separator',
      { label: 'Cut', icon: 'cut', run: () => { app.copySlides(); app.deleteSlides(); } },
      { label: 'Copy', icon: 'copy', run: () => app.copySlides() },
      { label: 'Paste', icon: 'paste', disabled: !app.hasSlideClipboard(), run: () => app.pasteSlides() },
      'separator',
      { label: allHidden ? 'Unskip slide' : 'Skip slide', icon: allHidden ? 'eye' : 'eyeOff', run: () => app.toggleSkip() },
      { label: 'Move up', icon: 'up', shortcut: 'Mod-Up', disabled: app.slideIndex === 0, run: () => app.moveSlides(-1) },
      { label: 'Move down', icon: 'down', shortcut: 'Mod-Down', disabled: app.slideIndex === app.deck.slides.length - 1, run: () => app.moveSlides(1) },
      'separator',
      { label: 'Change background…', icon: 'image', run: () => app.openPanel('design') },
      { label: 'Transition…', icon: 'animate', run: () => app.openPanel('transition') },
    ], { iconFn: (n) => icon(n) }), { at, label: 'Slide actions' });
  }

  setDropIndex(i) {
    this.dropIndex = i;
    for (const li of this.items.values()) li.classList.remove('drop-before', 'drop-after');
    if (i == null) return;
    const lis = [...this.list.children];
    if (i < lis.length) lis[i].classList.add('drop-before');
    else lis[lis.length - 1]?.classList.add('drop-after');
  }

  onDragOver(e) {
    if (!e.dataTransfer.types.includes('application/x-libreslides-slides')) return;
    e.preventDefault();
    const lis = [...this.list.children];
    let i = lis.length;
    for (let k = 0; k < lis.length; k++) {
      const r = lis[k].getBoundingClientRect();
      if (e.clientY < r.top + r.height / 2) { i = k; break; }
    }
    this.setDropIndex(i);
  }

  onDrop(e) {
    if (!e.dataTransfer.types.includes('application/x-libreslides-slides')) return;
    e.preventDefault();
    const i = this.dropIndex;
    this.setDropIndex(null);
    if (i != null) this.app.moveSlidesTo(i);
  }
}
