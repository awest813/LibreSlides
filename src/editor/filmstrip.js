import { h, showPopover, menu } from '../ui/dom.js';
import { icon } from '../ui/icons.js';
import { renderThumb } from '../render/render.js';
import { LAYOUTS, slideTitle, slideSections, hasSections, sectionAt, DEFAULT_SECTION } from '../model/deck.js';
import { openComments } from './comments.js';

const THUMB_W = 168;

/**
 * The slide list on the left: thumbnails you can click, multi-select,
 * reorder by dragging and manage from the keyboard or a context menu. When
 * the deck has sections, a header row before each one names it and folds it
 * away (folding is a view setting, not part of the deck).
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
    this.collapsed = new Set(); // ids of the first slides of folded sections
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
    const { slides } = this.app.deck;
    this.items.clear();
    const rows = slides.map((s, i) => this.item(s, i));
    if (hasSections(slides)) {
      for (const group of slideSections(slides).reverse()) {
        const folded = this.collapsed.has(slides[group.start].id);
        if (folded) for (const li of rows.slice(group.start, group.end)) li.classList.add('in-collapsed');
        rows.splice(group.start, 0, this.sectionHeader(group, folded));
      }
    }
    this.list.replaceChildren(...rows);
    this.syncSelection();
  }

  sectionHeader(group, folded) {
    const { slides } = this.app.deck;
    const name = group.name || DEFAULT_SECTION;
    const count = group.end - group.start;
    const countLabel = `${count} slide${count === 1 ? '' : 's'}`;
    const li = h('li', { class: `film-section${folded ? ' is-collapsed' : ''}${group.name ? '' : ' is-default'}`, role: 'none', dataset: { start: slides[group.start].id } },
      h('button', {
        type: 'button', class: 'film-section-toggle', 'aria-expanded': String(!folded), title: `${name} (${countLabel})`,
        'aria-label': `Section ${name}, ${countLabel}`, onclick: () => this.toggleSection(slides[group.start].id),
      },
      h('span', { class: 'film-section-chevron', html: icon('chevronDown') }),
      h('span', { class: 'film-section-name' }, name),
      h('span', { class: 'film-section-count' }, String(count))));
    li.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      this.sectionMenu(this.app.deck.slides.findIndex((s) => s.id === li.dataset.start), { x: e.clientX, y: e.clientY });
    });
    return li;
  }

  /** Fold or unfold the section that starts with slide `id`. */
  toggleSection(id, open = this.collapsed.has(id)) {
    if (open) this.collapsed.delete(id);
    else this.collapsed.add(id);
    const hadFocus = !!document.activeElement?.closest?.('.film-section');
    this.render();
    if (hadFocus) this.list.querySelector(`.film-section[data-start="${CSS.escape(id)}"] button`)?.focus();
  }

  /** Unfold the section of the current slide, so it can be seen. */
  revealCurrent() {
    const { slides } = this.app.deck;
    const group = sectionAt(slides, this.app.slideIndex);
    if (group && this.collapsed.has(slides[group.start].id)) this.toggleSection(slides[group.start].id, true);
  }

  sectionMenu(i, at) {
    const { app } = this;
    const { slides } = app.deck;
    const groups = slideSections(slides);
    const k = groups.findIndex((g) => i >= g.start && i < g.end);
    const group = groups[k];
    if (!group) return;
    const allFolded = groups.every((g) => this.collapsed.has(slides[g.start].id));
    showPopover(null, menu([
      { label: 'Rename section…', icon: 'rename', run: () => app.renameSection(i) },
      { label: 'Remove section', icon: 'minus', disabled: !group.name, run: () => app.removeSection(i) },
      { label: 'Select slides in section', icon: 'select', run: () => app.selectSection(i) },
      'separator',
      { label: 'Move section up', icon: 'up', disabled: !group.name || !groups[k - 1]?.name, run: () => app.moveSection(i, -1) },
      { label: 'Move section down', icon: 'down', disabled: !group.name || k === groups.length - 1, run: () => app.moveSection(i, 1) },
      'separator',
      {
        label: allFolded ? 'Expand all' : 'Collapse all', icon: allFolded ? 'chevronDown' : 'chevronRight',
        run: () => {
          this.collapsed = allFolded ? new Set() : new Set(groups.map((g) => slides[g.start].id));
          this.render();
        },
      },
      { label: 'Remove all sections', icon: 'trash', run: () => app.removeAllSections() },
    ], { iconFn: (n) => icon(n) }), { at, label: 'Section actions' });
  }

  item(slide, i) {
    const { deck } = this.app;
    const thumb = renderThumb(slide, deck, THUMB_W);
    const title = slideTitle(slide);
    const comments = openComments(slide);
    const commentsLabel = `${comments} open comment${comments === 1 ? '' : 's'}`;
    const li = h('li', {
      class: `film-item${slide.hidden ? ' is-hidden' : ''}`, role: 'option', tabindex: '-1', draggable: 'true',
      'aria-label': `Slide ${i + 1}${title ? `: ${title}` : ''}${slide.hidden ? ' (skipped)' : ''}${comments ? `, ${commentsLabel}` : ''}${slide.section ? `, starts section ${slide.section}` : ''}`,
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
    if (li.classList.contains('in-collapsed')) fresh.classList.add('in-collapsed');
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
    this.revealCurrent();
    const li = this.items.get(this.app.slide.id);
    li?.focus({ preventScroll: true });
    li?.scrollIntoView({ block: 'nearest' });
  }

  scrollToCurrent() {
    this.revealCurrent();
    this.items.get(this.app.slide.id)?.scrollIntoView({ block: 'nearest' });
  }

  onKey(e) {
    if (e.target.closest('.film-section')) return; // a section header's button
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
      { label: 'Cut', icon: 'cut', shortcut: 'Mod-X', run: () => { app.copySlides(); app.deleteSlides(); } },
      { label: 'Copy', icon: 'copy', shortcut: 'Mod-C', run: () => app.copySlides() },
      { label: 'Paste', icon: 'paste', shortcut: 'Mod-V', disabled: !app.hasSlideClipboard(), run: () => app.pasteSlides() },
      'separator',
      { label: `${allHidden ? 'Unskip' : 'Skip'} ${many ? 'slides' : 'slide'}`, icon: allHidden ? 'eye' : 'eyeOff', run: () => app.toggleSkip() },
      { label: many ? 'Move slides up' : 'Move slide up', icon: 'up', shortcut: 'Mod-Up', disabled: app.slideIndices()[0] === 0, run: () => app.moveSlides(-1) },
      { label: many ? 'Move slides down' : 'Move slide down', icon: 'down', shortcut: 'Mod-Down', disabled: app.slideIndices().at(-1) === app.deck.slides.length - 1, run: () => app.moveSlides(1) },
      'separator',
      ...this.sectionItems(),
      'separator',
      { label: 'Change background…', icon: 'image', run: () => app.openPanel('design') },
      { label: 'Transition…', icon: 'animate', run: () => app.openPanel('transition') },
    ], { iconFn: (n) => icon(n) }), { at, label: 'Slide actions' });
  }

  /** Section actions for the current slide, for the slide menus. */
  sectionItems() {
    const { app } = this;
    const group = sectionAt(app.deck.slides, app.slideIndex);
    const sectioned = hasSections(app.deck.slides);
    return [
      app.slide.section ? null : { label: 'Add section…', icon: 'section', run: () => app.addSection() },
      sectioned ? { label: 'Rename section…', icon: 'rename', run: () => app.renameSection(app.slideIndex) } : null,
      group?.name ? { label: 'Remove section', icon: 'minus', run: () => app.removeSection(app.slideIndex) } : null,
    ];
  }

  slideRows() {
    return [...this.list.querySelectorAll('.film-item')];
  }

  setDropIndex(i) {
    this.dropIndex = i;
    for (const el of this.list.querySelectorAll('.drop-before, .drop-after')) el.classList.remove('drop-before', 'drop-after');
    if (i == null) return;
    const lis = this.slideRows();
    if (i < lis.length) {
      // Slides dropped before a section's first slide go at the end of the section before (or, at the top, into the first).
      const header = lis[i].previousElementSibling;
      (i > 0 && header?.classList.contains('film-section') ? header : lis[i]).classList.add('drop-before');
    } else [...this.list.children].filter((el) => !el.classList.contains('in-collapsed')).at(-1)?.classList.add('drop-after');
  }

  onDragOver(e) {
    if (!e.dataTransfer.types.includes('application/x-libreslides-slides')) return;
    e.preventDefault();
    const lis = this.slideRows();
    let i = lis.length;
    for (let k = 0; k < lis.length; k++) {
      if (lis[k].classList.contains('in-collapsed')) continue;
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
