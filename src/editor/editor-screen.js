import { h, toast, debounce, showPopover, menu, setFocusFallback, isPopoverOpen, pickFileInput, closePopover, shortcutLabel, downloadBlob } from '../ui/dom.js';
import { icon, LOGO_SVG } from '../ui/icons.js';
import { openDialog, promptDialog, confirmDialog } from '../ui/dialog.js';
import {
  canSaveToFiles, pickSaveLocation, ensureWritePermission, writeToHandle, fileModified, formatOfName, fileNameFor,
  handleFromDataTransfer, FORMAT_NOTES,
} from '../io/file-access.js';
import { deckBlob } from '../io/files.js';
import { isDark, toggleTheme } from '../ui/theme.js';
import { showInstallHelp } from '../ui/install.js';
import { getDeck, saveDeck, onStorageChange } from '../storage/db.js';
import {
  slideSize, newSlide, layoutElements, cloneSlide, cloneElement, canHoldText, boxOf, moveElement, addMedia, newImage, newTable, newChart,
  htmlToText, SIZES, uid,
} from '../model/deck.js';
import { themeById, resolveColor } from '../model/themes.js';
import { FONT_SIZES } from '../model/fonts.js';
import { textStyle } from '../render/render.js';
import { sanitizeHtml, sanitizeElement } from '../io/sanitize.js';
import { readImageFile } from '../io/images.js';
import { SlideCanvas, aabb } from './canvas.js';
import { Filmstrip } from './filmstrip.js';
import { Toolbar } from './toolbar.js';
import { SidePanel } from './sidepanel.js';
import { History } from './history.js';
import { stripInline, exec, execFontSize, caretFormat, highlightAll } from './richtext.js';
import { VERSION, BUILD, diagnostics } from '../diagnostics.js';

const CLIP_TYPE = 'application/x-libreslides+json';
const isTextual = (el) => canHoldText(el) || el.type === 'table';
// Where keys type text (so Ctrl+Z and friends are the field's own); not checkboxes, sliders or lists.
const typingTarget = (t) => !!t && (t.isContentEditable || t.tagName === 'TEXTAREA'
  || (t.tagName === 'INPUT' && !/^(checkbox|radio|range|color|button|submit|reset|file|image)$/.test(t.type)));

let elementClipboard = null; // { elements, media } — shared by every editor in this tab
let slideClipboard = null; // { slides, media }

/** Wrap paragraphs in a list, or unwrap a list back into paragraphs. */
function toggleListHtml(html, kind) {
  const tpl = document.createElement('template');
  tpl.innerHTML = html || '';
  const top = [...tpl.content.children];
  const lines = [];
  for (const n of top) {
    if (n.tagName === 'UL' || n.tagName === 'OL') for (const li of n.children) lines.push({ html: li.innerHTML, style: li.getAttribute('style') });
    else lines.push({ html: n.innerHTML, style: n.getAttribute('style') });
  }
  if (!lines.length) lines.push({ html: '<br>', style: null });
  const already = top.length && top.every((n) => n.tagName === kind.toUpperCase());
  const attr = (s) => (s ? ` style="${s.replace(/"/g, '&quot;')}"` : '');
  return already
    ? lines.map((l) => `<p${attr(l.style)}>${l.html}</p>`).join('')
    : `<${kind}>${lines.map((l) => `<li${attr(l.style)}>${l.html}</li>`).join('')}</${kind}>`;
}

export class EditorScreen {
  constructor(root, { deckId, onHome, onOpenFile, onNewDeck, onCopyDeck, onPresent }) {
    this.root = root;
    this.deckId = deckId;
    this.options = { onHome, onOpenFile, onNewDeck, onCopyDeck, onPresent };
    this.selection = [];
    this.slideSelection = new Set();
    this.slideIndex = 0;
    this.history = new History();
    this.view = { notes: true };
    try {
      this.view.notes = localStorage.getItem('ls:notes') !== '0';
    } catch { /* storage blocked */ }
    this.saveState = 'saved';
    this.dirty = false;
    this.conflict = false;
    this.saveSoon = debounce(() => this.save(), 700);
    this.changedSlides = new Set();
    this.thumbSoon = debounce(() => {
      for (const id of this.changedSlides) this.filmstrip.refreshThumb(id);
      this.changedSlides.clear();
    }, 250);
    this.uiSoon = () => {
      if (this.uiFrame) return;
      this.uiFrame = requestAnimationFrame(() => {
        this.uiFrame = 0;
        this.toolbar.update();
      });
    };
  }

  // ---------------------------------------------------------------- lifecycle

  async mount() {
    const deck = await getDeck(this.deckId);
    if (!deck) return false;
    this.deck = deck;
    this.rev = deck.rev || 0;
    this.slideSelection = new Set([deck.slides[0].id]);
    this.build();
    this.goToSlide(0);
    this.filmstrip.render();
    this.canvas.layout();
    this.unsubscribe = onStorageChange((msg) => this.onStorageMessage(msg));
    return true;
  }

  build() {
    this.canvas = new SlideCanvas(this);
    this.filmstrip = new Filmstrip(this);
    this.toolbar = new Toolbar(this);
    this.panel = new SidePanel(this);

    this.titleInput = h('input', { class: 'doc-title-input', 'aria-label': 'Presentation name', value: this.deck.title, spellcheck: 'false' });
    this.titleInput.addEventListener('change', () => this.rename(this.titleInput.value));
    this.titleInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === 'Escape') {
        if (e.key === 'Escape') this.titleInput.value = this.deck.title;
        this.focusCanvas();
      }
    });
    this.saveLabel = h('span', { class: 'save-state', role: 'status' });
    // Where the presentation saves to on this device, and whether that file is up to date.
    this.fileChip = h('button', { type: 'button', class: 'file-chip', onclick: () => this.saveToFile() });
    const presentMain = h('button', { type: 'button', class: 'present-main', html: `${icon('present')}<span>Slideshow</span>`, title: `Present from this slide (${shortcutLabel('Mod-Enter')})`, onclick: () => this.present({ fromStart: false }) });
    const presentMore = h('button', { type: 'button', class: 'present-more', 'aria-label': 'Slideshow options', 'aria-haspopup': 'menu', html: icon('chevronDown') });
    presentMore.addEventListener('click', () => showPopover(presentMore, menu([
      { label: 'Present from this slide', icon: 'present', shortcut: 'Mod-Enter', run: () => this.present({ fromStart: false }) },
      { label: 'Present from beginning', icon: 'play', shortcut: 'F5', run: () => this.present({ fromStart: true }) },
      { label: 'Presenter view', icon: 'presenter', run: () => this.presenterView() },
    ], { iconFn: (n) => icon(n) }), { placement: 'bottom-end' }));
    this.darkBtn = h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Toggle dark mode', title: 'Dark mode', html: icon(isDark() ? 'sun' : 'moon'), onclick: () => this.toggleDarkUi() });

    const titlebar = h('header', { class: 'titlebar editor-titlebar' },
      h('button', { type: 'button', class: 'app-logo', 'aria-label': 'All presentations', title: 'All presentations', html: LOGO_SVG('#fff', 'transparent', 28), onclick: () => this.goHome() }),
      h('div', { class: 'title-stack' },
        h('div', { class: 'title-row' }, this.titleInput, this.fileChip, this.saveLabel),
        this.toolbar.menubar),
      h('div', { class: 'titlebar-right' },
        h('div', { class: 'present-btn' }, presentMain, presentMore),
        this.darkBtn));

    this.notes = h('textarea', { class: 'notes-input', placeholder: 'Click to add speaker notes', 'aria-label': 'Speaker notes', spellcheck: 'true' });
    this.notes.addEventListener('input', () => {
      const s = this.slide;
      const before = this.snapshot();
      s.notes = this.notes.value;
      this.history.push(before, `notes:${s.id}`);
      this.markDirty();
    });
    this.notesPane = h('div', { class: 'notes-pane', hidden: !this.view.notes }, this.notes);
    const center = h('div', { class: 'center-col' }, this.canvas.wrap, this.notesPane);
    const workspace = h('div', { class: 'workspace' }, this.filmstrip.el, center, this.panel.el);

    this.el = h('div', { class: 'editor-screen' }, titlebar, this.toolbar.toolbar, workspace);
    this.root.replaceChildren(this.el);
    setFocusFallback(() => this.focusCanvas());

    this.onKeyDown = (e) => this.handleKey(e);
    this.onCopy = (e) => this.handleCopy(e, false);
    this.onCut = (e) => this.handleCopy(e, true);
    this.onPaste = (e) => this.handlePaste(e);
    this.onBeforeUnload = (e) => {
      if (this.dirty) this.save();
      // Changes not yet written to the linked file: let the browser ask before closing.
      if (this.deck.file?.dirty) {
        e.preventDefault();
        e.returnValue = '';
      }
    };
    this.onVisibility = () => {
      if (document.visibilityState === 'hidden') this.flush();
    };
    document.addEventListener('keydown', this.onKeyDown);
    document.addEventListener('copy', this.onCopy);
    document.addEventListener('cut', this.onCut);
    document.addEventListener('paste', this.onPaste);
    window.addEventListener('beforeunload', this.onBeforeUnload);
    document.addEventListener('visibilitychange', this.onVisibility);

    // Pictures dropped on the slide.
    this.canvas.wrap.addEventListener('dragover', (e) => {
      if (e.dataTransfer?.types?.includes('Files')) {
        e.preventDefault();
        e.stopPropagation();
      }
    });
    this.canvas.wrap.addEventListener('drop', (e) => {
      const files = [...(e.dataTransfer?.files || [])];
      if (!files.length) return;
      e.preventDefault();
      e.stopPropagation();
      const images = files.filter((f) => f.type.startsWith('image/'));
      if (images.length) this.insertImageFiles(images, this.canvas.toSlide(e));
      else {
        // The handle must be requested synchronously, during the drop event.
        handleFromDataTransfer(e.dataTransfer, files[0]).then((handle) => this.options.onOpenFile?.(files[0], handle));
      }
    });
    this.updateSaveLabel();
    this.updateFileChip();
  }

  async destroy({ save = true } = {}) {
    if (save) await this.flush();
    this.saveSoon.cancel();
    this.thumbSoon.cancel();
    this.canvas.destroy();
    this.unsubscribe?.();
    this.presentation?.exit?.();
    document.removeEventListener('keydown', this.onKeyDown);
    document.removeEventListener('copy', this.onCopy);
    document.removeEventListener('cut', this.onCut);
    document.removeEventListener('paste', this.onPaste);
    window.removeEventListener('beforeunload', this.onBeforeUnload);
    document.removeEventListener('visibilitychange', this.onVisibility);
    setFocusFallback(null);
    closePopover();
  }

  // ---------------------------------------------------------------- state

  get slide() { return this.deck.slides[this.slideIndex]; }
  get theme() { return themeById(this.deck.themeId); }
  get size() { return slideSize(this.deck); }

  selectedElements() {
    const set = new Set(this.selection);
    return this.slide.elements.filter((e) => set.has(e.id));
  }

  firstSelected() { return this.selectedElements()[0] || null; }

  setSelection(ids, { quiet = false } = {}) {
    if (this.canvas.editing && !(ids.length === 1 && ids[0] === this.canvas.editing.id)) this.canvas.endEdit();
    this.selection = ids.filter((id) => this.slide.elements.some((e) => e.id === id));
    this.canvas.drawOverlay();
    if (!quiet) this.onSelectionChanged();
  }

  onSelectionChanged() {
    this.toolbar.update();
    this.panel.update();
  }

  onTextSelection() { this.uiSoon(); }
  onZoom() { this.toolbar.update(); }
  onToolChange() { this.toolbar.update(); }

  focusCanvas() {
    if (this.canvas.editing) this.canvas.restoreSelection();
    else this.canvas.wrap.focus({ preventScroll: true });
  }

  snapshot() {
    return JSON.stringify({ slides: this.deck.slides, themeId: this.deck.themeId, size: this.deck.size, slideId: this.slide?.id, selection: this.selection });
  }

  restore(snap) {
    const s = JSON.parse(snap);
    this.deck.slides = s.slides;
    this.deck.themeId = s.themeId;
    this.deck.size = s.size;
    let i = this.deck.slides.findIndex((x) => x.id === s.slideId);
    if (i < 0) i = Math.min(this.slideIndex, this.deck.slides.length - 1);
    this.slideIndex = i;
    this.slideSelection = new Set([this.slide.id]);
    this.selection = s.selection.filter((id) => this.slide.elements.some((e) => e.id === id));
    this.notesSlideId = null;
    this.filmstrip.render();
    this.refresh();
    this.markDirty();
  }

  /** Record `before` as an undo step for a change already made, then save and redraw. */
  commit(before, { coalesce = null, render = true, structure = false } = {}) {
    const after = this.snapshot();
    if (after !== before) {
      // Text typed in the box being edited is its own undo step, before this change.
      const ed = this.canvas.editing;
      if (ed && ed.before !== before) this.history.push(ed.before);
      this.history.push(before, coalesce);
      if (ed) {
        const el = this.canvas.find(ed.id);
        ed.before = after;
        ed.original = JSON.stringify([el?.text, el?.cells]);
      }
      this.markDirty({ thumbs: !structure });
    }
    if (structure) this.filmstrip.render();
    if (render) this.refresh();
    else this.onSelectionChanged();
  }

  /** Make a change to the deck as one undo step. */
  change(fn, opts = {}) {
    const before = this.snapshot();
    const result = fn();
    this.commit(before, opts);
    return result;
  }

  refresh() {
    this.canvas.render();
    this.toolbar.update();
    this.panel.update();
    // Keep the notes the user is typing, unless they belong to another slide (or undo/redo changed them).
    if (document.activeElement !== this.notes || this.notesSlideId !== this.slide.id) {
      this.notes.value = this.slide.notes || '';
      this.notesSlideId = this.slide.id;
    }
  }

  markDirty({ thumbs = false } = {}) {
    this.dirty = true;
    this.changeCount = (this.changeCount || 0) + 1;
    if (this.deck.file && !this.deck.file.dirty) {
      this.deck.file.dirty = true;
      this.updateFileChip();
    }
    this.setSaveState('unsaved');
    this.saveSoon();
    if (thumbs) {
      this.changedSlides.add(this.slide.id);
      this.thumbSoon();
    }
  }

  setSaveState(state) {
    this.saveState = state;
    this.updateSaveLabel();
  }

  updateSaveLabel() {
    if (!this.saveLabel) return;
    const text = { saved: 'Saved on this device', unsaved: 'Saving…', saving: 'Saving…', conflict: 'Not saved: changed in another tab', error: 'Couldn’t save' }[this.saveState];
    this.saveLabel.textContent = text;
    this.saveLabel.classList.toggle('is-error', this.saveState === 'conflict' || this.saveState === 'error');
  }

  async save() {
    if (!this.dirty || this.conflict || this.saving) {
      if (this.saving) this.saveSoon();
      return;
    }
    this.saving = true;
    this.dirty = false;
    this.setSaveState('saving');
    try {
      const r = await saveDeck(this.deck, { expectRev: this.rev });
      if (r.conflict) {
        this.conflict = true;
        this.dirty = true;
        this.setSaveState('conflict');
        toast('This presentation was changed in another tab, so changes here are no longer saved.', {
          type: 'error', timeout: 60000, action: { label: 'Reload', run: () => this.reload() },
        });
      } else {
        this.rev = r.rev;
        this.setSaveState(this.dirty ? 'unsaved' : 'saved');
      }
    } catch (err) {
      console.error(err);
      this.dirty = true;
      this.setSaveState('error');
      toast(`Couldn’t save: ${err.message || err}`, { type: 'error', timeout: 8000 });
    } finally {
      this.saving = false;
    }
  }

  async flush() {
    this.saveSoon.cancel();
    if (this.canvas?.editing) this.canvas.endEdit();
    while (this.saving) await new Promise((r) => setTimeout(r, 20));
    await this.save();
  }

  async reload() {
    this.canvas.cancelEdit();
    const deck = await getDeck(this.deckId);
    if (!deck) return;
    this.deck = deck;
    this.rev = deck.rev || 0;
    this.conflict = false;
    this.dirty = false;
    this.history = new History();
    this.titleInput.value = deck.title;
    this.slideIndex = Math.min(this.slideIndex, deck.slides.length - 1);
    this.selection = [];
    this.slideSelection = new Set([this.slide.id]);
    this.notesSlideId = null;
    this.filmstrip.render();
    this.refresh();
    this.setSaveState('saved');
  }

  onStorageMessage(msg) {
    if (msg.id !== this.deckId) return;
    if (msg.type === 'deleted') {
      toast('This presentation was deleted in another tab.', { type: 'error', timeout: 10000 });
      this.conflict = true;
      this.setSaveState('conflict');
    } else if (msg.type === 'saved' && msg.rev > this.rev && !this.dirty && !this.canvas.editing && !this.saving) {
      this.reload(); // nothing unsaved here: just show the newer version
    }
  }

  // ---------------------------------------------------------------- navigation

  goToSlide(i, { keepSlideSelection = false } = {}) {
    if (this.canvas.editing) this.canvas.endEdit();
    this.slideIndex = Math.max(0, Math.min(this.deck.slides.length - 1, i));
    this.selection = [];
    if (!keepSlideSelection) this.slideSelection = new Set([this.slide.id]);
    this.history.seal();
    this.filmstrip.syncSelection();
    this.filmstrip.scrollToCurrent();
    this.refresh();
  }

  goHome() { this.options.onHome(); }

  rename(title) {
    const t = title.trim() || 'Untitled presentation';
    this.titleInput.value = t;
    if (t === this.deck.title) return;
    this.deck.title = t;
    document.title = `${t} – LibreSlides`;
    this.markDirty();
  }

  async renamePrompt() {
    const r = await promptDialog({ title: 'Rename presentation', fields: [{ name: 'title', label: 'Name', value: this.deck.title }], confirmLabel: 'Rename' });
    if (r) this.rename(r.title);
  }

  // ---------------------------------------------------------------- slides

  slideIndices() {
    return this.deck.slides.map((s, i) => (this.slideSelection.has(s.id) ? i : -1)).filter((i) => i >= 0);
  }

  addSlide(layout = null) {
    const cur = this.slide;
    const l = layout || (cur.layout === 'title' || cur.layout === 'section' ? 'titleContent' : cur.layout === 'blank' ? 'blank' : cur.layout || 'titleContent');
    const s = newSlide(l, this.deck);
    s.transition = { ...(cur.transition || { type: 'none', duration: 0.6 }) };
    this.change(() => this.deck.slides.splice(this.slideIndex + 1, 0, s), { structure: true, render: false });
    this.goToSlide(this.slideIndex + 1);
  }

  duplicateSlides() {
    const idx = this.slideIndices();
    const copies = idx.map((i) => cloneSlide(this.deck.slides[i]));
    const at = idx[idx.length - 1] + 1;
    this.change(() => this.deck.slides.splice(at, 0, ...copies), { structure: true, render: false });
    this.slideSelection = new Set(copies.map((s) => s.id));
    this.goToSlide(at, { keepSlideSelection: true });
  }

  deleteSlides() {
    this.canvas.endEdit(); // keep the text being typed (and its undo step)
    const idx = this.slideIndices();
    const first = idx[0];
    this.change(() => {
      this.deck.slides = this.deck.slides.filter((s) => !this.slideSelection.has(s.id));
      if (!this.deck.slides.length) this.deck.slides.push(newSlide('title', this.deck));
    }, { structure: true, render: false });
    this.goToSlide(Math.min(first, this.deck.slides.length - 1));
  }

  moveSlides(delta) {
    const idx = this.slideIndices();
    if (!idx.length) return;
    let target;
    if (delta === -Infinity) target = 0;
    else if (delta === Infinity) target = this.deck.slides.length;
    else target = delta < 0 ? idx[0] + delta : idx[idx.length - 1] + 1 + delta;
    this.moveSlidesTo(Math.max(0, Math.min(this.deck.slides.length, target)));
  }

  /** Move the selected slides so they start at position `index` (counted before the move). */
  moveSlidesTo(index) {
    const ids = this.slideSelection;
    const moving = this.deck.slides.filter((s) => ids.has(s.id));
    const before = this.deck.slides.slice(0, index).filter((s) => !ids.has(s.id));
    const after = this.deck.slides.slice(index).filter((s) => !ids.has(s.id));
    const order = [...before, ...moving, ...after];
    if (order.every((s, i) => s === this.deck.slides[i])) return;
    const currentId = this.slide.id;
    this.change(() => { this.deck.slides = order; }, { structure: true, render: false });
    this.goToSlide(order.findIndex((s) => s.id === currentId), { keepSlideSelection: true });
  }

  toggleSkip() {
    const sel = this.deck.slides.filter((s) => this.slideSelection.has(s.id));
    const hide = !sel.every((s) => s.hidden);
    this.change(() => sel.forEach((s) => { s.hidden = hide; }), { structure: true });
    toast(hide ? 'Skipped in the slideshow' : 'Shown in the slideshow');
  }

  copySlides() {
    const slides = this.deck.slides.filter((s) => this.slideSelection.has(s.id));
    slideClipboard = { slides: JSON.parse(JSON.stringify(slides)), media: this.mediaFor(slides.flatMap((s) => [...s.elements, s.background || {}])) };
    toast(`Copied ${slides.length} slide${slides.length === 1 ? '' : 's'}`);
  }

  hasSlideClipboard() { return !!slideClipboard; }

  pasteSlides() {
    if (!slideClipboard) return;
    const copies = slideClipboard.slides.map((s) => cloneSlide(s));
    this.change(() => {
      this.importMedia(slideClipboard.media, copies.flatMap((s) => [...s.elements, s.background || {}]));
      this.deck.slides.splice(this.slideIndex + 1, 0, ...copies);
    }, { structure: true, render: false });
    this.slideSelection = new Set(copies.map((s) => s.id));
    this.goToSlide(this.slideIndex + 1, { keepSlideSelection: true });
  }

  changeSlide(fn) {
    this.change(() => fn(this.slide));
  }

  applyLayout(layout) {
    this.canvas.endEdit(); // the placeholder being edited is replaced below
    this.change(() => {
      const s = this.slide;
      const { w, h: H } = this.size;
      const fresh = layoutElements(layout, w, H);
      const old = s.elements.filter((e) => e.ph);
      const take = (role) => {
        const i = old.findIndex((e) => e.ph === role || (role === 'body' && e.ph === 'subtitle') || (role === 'subtitle' && e.ph === 'body'));
        return i < 0 ? null : old.splice(i, 1)[0];
      };
      for (const el of fresh) {
        const prev = take(el.ph);
        if (prev) el.text = prev.text;
      }
      // Placeholder text the new layout has no room for stays as an ordinary text box.
      const leftovers = old.filter((e) => e.text).map((e) => ({ ...e, ph: undefined }));
      s.elements = [...fresh, ...s.elements.filter((e) => !e.ph), ...leftovers];
      s.layout = layout;
      this.selection = [];
    }, { structure: true });
  }

  setTheme(id) {
    this.change(() => { this.deck.themeId = id; }, { structure: true });
  }

  setSize(size) {
    if (!SIZES[size] || size === this.deck.size) return;
    const sx = SIZES[size].w / this.size.w;
    this.change(() => {
      this.deck.size = size;
      for (const s of this.deck.slides) {
        for (const el of s.elements) {
          if (el.type === 'line') {
            el.x1 *= sx; el.x2 *= sx;
          } else if (el.type === 'image' || (el.type === 'shape' && el.lockRatio)) {
            // Keep proportions: scale about the element's centre.
            const k = Math.min(1, sx);
            const cx = (el.x + el.w / 2) * sx;
            const cy = el.y + el.h / 2;
            el.w *= k; el.h *= k;
            el.x = cx - el.w / 2; el.y = cy - el.h / 2;
          } else {
            el.x *= sx; el.w *= sx;
          }
        }
      }
    }, { structure: true });
    this.canvas.layout();
  }

  setBackground(bg, { all = false } = {}) {
    this.change(() => {
      for (const s of all ? this.deck.slides : [this.slide]) s.background = bg ? JSON.parse(JSON.stringify(bg)) : null;
    }, { structure: all });
    if (all) toast('Background applied to all slides');
  }

  async pickBackgroundImage() {
    const file = await pickFileInput('image/*');
    if (!file) return;
    try {
      const img = await readImageFile(file, { maxSide: 2560 });
      const media = addMedia(this.deck, img.dataUrl);
      this.setBackground({ type: 'image', media, color: '#000000' });
    } catch (err) {
      toast(err.message || 'Couldn’t read that picture.', { type: 'error' });
    }
  }

  setTransition(t, { all = false, coalesce = null } = {}) {
    this.change(() => {
      for (const s of all ? this.deck.slides : [this.slide]) s.transition = { ...t };
    }, { coalesce, render: false });
    if (!coalesce) this.panel.update(true);
    if (all) toast('Transition applied to all slides');
  }

  // ---------------------------------------------------------------- elements

  addElements(els) {
    this.change(() => this.slide.elements.push(...els));
    this.setSelection(els.map((e) => e.id));
  }

  deleteSelection() {
    if (!this.selection.length) return;
    const ids = new Set(this.selection);
    this.change(() => { this.slide.elements = this.slide.elements.filter((e) => !ids.has(e.id)); });
    this.setSelection([]);
  }

  duplicateSelection() {
    const els = this.selectedElements();
    if (!els.length) return;
    const copies = els.map((e) => {
      const c = cloneElement(e);
      moveElement(c, 16, 16);
      if (c.anim) c.anim.order = this.nextAnimOrder() + (e.anim.order || 0);
      return c;
    });
    this.addElements(copies);
  }

  selectAll() {
    this.setSelection(this.slide.elements.map((e) => e.id));
  }

  /** Pictures referenced by elements, so a copy carries them along. */
  mediaFor(els) {
    const media = {};
    for (const el of els) if (el.media && this.deck.media[el.media]) media[el.media] = this.deck.media[el.media];
    return media;
  }

  /** Add copied pictures to this deck and point the elements at them. */
  importMedia(media, els) {
    const map = {};
    for (const [id, data] of Object.entries(media || {})) map[id] = addMedia(this.deck, data);
    for (const el of els) if (el.media && map[el.media]) el.media = map[el.media];
  }

  clipboardPayload() {
    const els = this.selectedElements();
    return { app: 'LibreSlides', elements: JSON.parse(JSON.stringify(els)), media: this.mediaFor(els) };
  }

  copySelection({ cut = false } = {}) {
    if (!this.selection.length) return;
    elementClipboard = this.clipboardPayload();
    const text = this.selectedElements().map((e) => htmlToText(e.text || '')).filter(Boolean).join('\n');
    navigator.clipboard?.writeText?.(text || ' ').catch(() => {});
    if (cut) this.deleteSelection();
  }

  pasteElements(payload) {
    const { w, h: H } = this.size;
    const copies = payload.elements.map((e) => cloneElement(e));
    // Pasting onto the same spot again steps each copy down and right.
    const sameSpot = copies.some((c) => this.slide.elements.some((e) => {
      const a = boxOf(e);
      const b = boxOf(c);
      return Math.abs(a.x - b.x) < 1 && Math.abs(a.y - b.y) < 1;
    }));
    for (const c of copies) {
      if (sameSpot) moveElement(c, 16, 16);
      const b = boxOf(c);
      if (b.x > w - 10 || b.y > H - 10) moveElement(c, -b.x + 20, -b.y + 20);
      delete c.anim;
    }
    this.change(() => {
      this.importMedia(payload.media, copies);
      this.slide.elements.push(...copies);
    });
    this.setSelection(copies.map((c) => c.id));
  }

  async pasteFromMenu() {
    if (elementClipboard) {
      this.pasteElements(elementClipboard);
      return;
    }
    try {
      const text = await navigator.clipboard.readText();
      if (text) this.pasteText(text);
    } catch {
      toast(`Use ${shortcutLabel('Mod-V')} to paste from other apps.`);
    }
  }

  pasteText(text) {
    const { w, h: H } = this.size;
    const lines = text.replace(/\r\n?/g, '\n').split('\n');
    const el = { ...newTextBoxAt(w, H), text: sanitizeHtml(lines.map((l) => `<p>${escapeText(l) || '<br>'}</p>`).join('')) };
    this.addElements([el]);
    requestAnimationFrame(() => this.canvas.autoGrow(el));
  }

  handleCopy(e, cut) {
    if (!this.isCanvasContext(e.target) || !this.selection.length || this.canvas.editing) return;
    e.preventDefault();
    const payload = this.clipboardPayload();
    elementClipboard = payload;
    e.clipboardData.setData(CLIP_TYPE, JSON.stringify(payload));
    const text = this.selectedElements().map((x) => htmlToText(x.text || '')).filter(Boolean).join('\n');
    e.clipboardData.setData('text/plain', text);
    if (cut) this.deleteSelection();
  }

  handlePaste(e) {
    if (this.canvas.editing || typingTarget(e.target) || document.querySelector('dialog[open]')) return;
    const dt = e.clipboardData;
    if (this.filmstrip.el.contains(document.activeElement)) {
      if (slideClipboard) {
        e.preventDefault();
        this.pasteSlides();
      }
      return;
    }
    e.preventDefault();
    const custom = dt?.getData(CLIP_TYPE);
    if (custom) {
      try {
        const payload = JSON.parse(custom);
        // Clipboard data can come from any page: clean text and colours like an opened file.
        for (const el of payload.elements || []) sanitizeElement(el);
        payload.media = Object.fromEntries(Object.entries(payload.media || {}).filter(([, v]) => /^data:image\//.test(v)));
        this.pasteElements(payload);
        return;
      } catch { /* fall through to text */ }
    }
    const files = [...(dt?.files || [])].filter((f) => f.type.startsWith('image/'));
    if (files.length) {
      this.insertImageFiles(files);
      return;
    }
    const text = dt?.getData('text/plain');
    if (text && text.trim()) this.pasteText(text);
    else if (elementClipboard) this.pasteElements(elementClipboard);
  }

  isCanvasContext(target) {
    return !typingTarget(target) && !this.filmstrip.el.contains(target) && !this.panel.el.contains(target) && !document.querySelector('dialog[open]');
  }

  updateSelected(fn, { coalesce = null } = {}) {
    const els = this.selectedElements();
    if (!els.length) return;
    this.change(() => els.forEach(fn), { coalesce });
  }

  setElementProp(prop, value) {
    this.updateSelected((e) => { if (e.type !== 'line' && e.type !== 'image') e[prop] = value; });
  }

  setStroke(partial) {
    this.updateSelected((e) => { e.stroke = { color: 'theme:text', width: 2, dash: 'solid', ...(e.stroke || {}), ...partial }; });
  }

  setLineEnds(start, end) {
    this.updateSelected((e) => { if (e.type === 'line') { e.arrowStart = start; e.arrowEnd = end; } });
  }

  order(action) {
    const ids = new Set(this.selection);
    if (!ids.size) return;
    this.change(() => {
      const list = this.slide.elements;
      const sel = list.filter((e) => ids.has(e.id));
      const rest = list.filter((e) => !ids.has(e.id));
      if (action === 'front') this.slide.elements = [...rest, ...sel];
      else if (action === 'back') this.slide.elements = [...sel, ...rest];
      else {
        const arr = [...list];
        const step = action === 'forward' ? 1 : -1;
        const indices = arr.map((e, i) => (ids.has(e.id) ? i : -1)).filter((i) => i >= 0);
        if (step > 0) indices.reverse();
        for (const i of indices) {
          const j = i + step;
          if (j < 0 || j >= arr.length || ids.has(arr[j].id)) continue;
          [arr[i], arr[j]] = [arr[j], arr[i]];
        }
        this.slide.elements = arr;
      }
    });
  }

  align(kind, { toSlide = false } = {}) {
    const els = this.selectedElements().filter((e) => !e.locked);
    if (!els.length) return;
    const { w, h: H } = this.size;
    const ref = els.length === 1 || toSlide ? { x: 0, y: 0, w, h: H } : (() => {
      const bs = els.map(aabb);
      const x = Math.min(...bs.map((b) => b.x));
      const y = Math.min(...bs.map((b) => b.y));
      return { x, y, w: Math.max(...bs.map((b) => b.x + b.w)) - x, h: Math.max(...bs.map((b) => b.y + b.h)) - y };
    })();
    this.change(() => {
      for (const el of els) {
        const b = aabb(el);
        let dx = 0;
        let dy = 0;
        if (kind === 'left') dx = ref.x - b.x;
        if (kind === 'center') dx = ref.x + ref.w / 2 - (b.x + b.w / 2);
        if (kind === 'right') dx = ref.x + ref.w - (b.x + b.w);
        if (kind === 'top') dy = ref.y - b.y;
        if (kind === 'middle') dy = ref.y + ref.h / 2 - (b.y + b.h / 2);
        if (kind === 'bottom') dy = ref.y + ref.h - (b.y + b.h);
        moveElement(el, Math.round(dx * 10) / 10, Math.round(dy * 10) / 10);
      }
    });
  }

  distribute(axis) {
    const els = this.selectedElements().filter((e) => !e.locked);
    if (els.length < 3) return;
    const key = axis === 'x' ? 'x' : 'y';
    const len = axis === 'x' ? 'w' : 'h';
    const items = els.map((el) => ({ el, b: aabb(el) })).sort((a, b) => a.b[key] - b.b[key]);
    const start = items[0].b[key];
    const end = Math.max(...items.map((i) => i.b[key] + i.b[len]));
    const total = items.reduce((s, i) => s + i.b[len], 0);
    const gap = (end - start - total) / (items.length - 1);
    this.change(() => {
      let pos = start;
      for (const { el, b } of items) {
        const d = pos - b[key];
        moveElement(el, axis === 'x' ? d : 0, axis === 'x' ? 0 : d);
        pos += b[len] + gap;
      }
    });
  }

  rotateBy(deg) {
    this.updateSelected((e) => {
      if (e.type === 'line') {
        const cx = (e.x1 + e.x2) / 2;
        const cy = (e.y1 + e.y2) / 2;
        const r = (deg * Math.PI) / 180;
        const rot = (x, y) => [cx + (x - cx) * Math.cos(r) - (y - cy) * Math.sin(r), cy + (x - cx) * Math.sin(r) + (y - cy) * Math.cos(r)];
        [e.x1, e.y1] = rot(e.x1, e.y1);
        [e.x2, e.y2] = rot(e.x2, e.y2);
      } else e.rot = (((e.rot || 0) + deg) % 360 + 360) % 360;
    });
  }

  flip(axis) {
    this.updateSelected((e) => {
      if (e.type === 'line') {
        if (axis === 'h') [e.x1, e.x2] = [e.x2, e.x1];
        else [e.y1, e.y2] = [e.y2, e.y1];
      } else if (axis === 'h') e.flipH = !e.flipH;
      else e.flipV = !e.flipV;
    });
  }

  toggleLock() {
    const els = this.selectedElements();
    const lock = !els.every((e) => e.locked);
    this.updateSelected((e) => { e.locked = lock; });
  }

  // ---------------------------------------------------------------- text formatting

  hasTextTarget() {
    return this.canvas.isEditing || this.selectedElements().some(isTextual);
  }

  textElements() {
    return this.selectedElements().filter(isTextual);
  }

  firstSelectedTextStyle() {
    const el = this.textElements()[0];
    return el ? textStyle(el, this.theme) : null;
  }

  currentTextFormat() {
    if (this.canvas.editing) {
      const sel = document.getSelection();
      const body = this.canvas.editing.bodies.find((b) => b.contains(sel?.anchorNode)) || this.canvas.editing.bodies[0];
      const f = caretFormat(body);
      if (f) {
        let node = sel.anchorNode.nodeType === 3 ? sel.anchorNode.parentElement : sel.anchorNode;
        f.colorCss = getComputedStyle(node).color;
        return f;
      }
    }
    const el = this.textElements()[0];
    if (!el) return null;
    const t = textStyle(el, this.theme);
    const html = el.text || '';
    return { ...t, colorCss: resolveColor(t.color, this.theme, '#000'), bullets: html.startsWith('<ul') || (!html && el.bullets), numbers: html.startsWith('<ol') };
  }

  currentTextColor() {
    if (this.canvas.editing) return null;
    return this.firstSelectedTextStyle()?.color || null;
  }

  /** Set a text property: on the selected text while typing, otherwise on the whole of each selected box. */
  setTextProp(prop, value) {
    const c = this.canvas;
    const inline = c.editing && ['font', 'size', 'color', 'highlight', 'align', 'bold', 'italic', 'underline', 'strike'].includes(prop);
    if (inline) {
      c.execText((body) => {
        if (prop === 'font') exec('fontName', value);
        else if (prop === 'size') execFontSize(body, value);
        else if (prop === 'color') exec('foreColor', resolveColor(value, this.theme, '#000000'));
        else if (prop === 'highlight') exec('hiliteColor', value ? resolveColor(value, this.theme) : 'transparent');
        else if (prop === 'align') exec({ left: 'justifyLeft', center: 'justifyCenter', right: 'justifyRight', justify: 'justifyFull' }[value]);
        else exec({ bold: 'bold', italic: 'italic', underline: 'underline', strike: 'strikeThrough' }[prop]);
      });
      this.toolbar.update();
      return;
    }
    const els = c.editing ? [c.find(c.editing.id)].filter(Boolean) : this.textElements();
    if (!els.length) return;
    if (prop === 'highlight') {
      const color = value ? resolveColor(value, this.theme) : null;
      this.change(() => {
        for (const el of els) {
          if (el.text) el.text = highlightAll(el.text, color);
          if (el.cells) for (const row of el.cells) for (const cell of row) cell.text = highlightAll(cell.text, color);
        }
      });
      return;
    }
    this.change(() => {
      for (const el of els) {
        el.style = { ...(el.style || {}), [prop]: value };
        if (value == null) delete el.style[prop];
        if (el.text) el.text = stripInline(el.text, prop);
        if (el.cells) for (const row of el.cells) for (const cell of row) cell.text = stripInline(cell.text, prop);
      }
    }, { coalesce: `text-${prop}` });
  }

  toggleTextProp(prop) {
    if (this.canvas.editing) {
      this.setTextProp(prop, null);
      return;
    }
    const t = this.firstSelectedTextStyle();
    if (t) this.setTextProp(prop, !t[prop]);
  }

  stepFontSize(dir) {
    const next = (n) => (dir > 0 ? FONT_SIZES.find((s) => s > n + 0.01) || n + 8 : [...FONT_SIZES].reverse().find((s) => s < n - 0.01) || Math.max(1, n - 1));
    if (this.canvas.editing) {
      const f = this.currentTextFormat();
      this.setTextProp('size', next(f?.size || 18));
      return;
    }
    const els = this.textElements();
    if (!els.length) return;
    this.change(() => {
      for (const el of els) {
        el.style = { ...el.style, size: next(textStyle(el, this.theme).size) };
        if (el.text) el.text = stripInline(el.text, 'size');
      }
    }, { coalesce: 'text-size' });
  }

  toggleList(kind) {
    if (this.canvas.editing) {
      this.canvas.execText(() => exec(kind === 'ul' ? 'insertUnorderedList' : 'insertOrderedList'));
      this.toolbar.update();
      return;
    }
    const els = this.textElements().filter(canHoldText);
    if (!els.length) return;
    this.change(() => {
      for (const el of els) {
        el.text = toggleListHtml(el.text, kind);
        el.bullets = el.text.startsWith('<ul');
      }
    });
  }

  indent(dir) {
    if (this.canvas.editing) this.canvas.execText(() => exec(dir > 0 ? 'indent' : 'outdent'));
  }

  clearFormatting() {
    if (this.canvas.editing) {
      this.canvas.execText(() => exec('removeFormat'));
      return;
    }
    const props = ['font', 'color', 'bold', 'italic', 'underline', 'strike', 'highlight', 'size', 'align'];
    this.change(() => {
      for (const el of this.textElements()) {
        const keep = { valign: el.style?.valign, lineHeight: el.style?.lineHeight, spacing: el.style?.spacing };
        // Placeholders go back to their layout size; other boxes keep theirs.
        if (!el.ph && el.style?.size) keep.size = el.style.size;
        if (el.ph) {
          const fresh = layoutElements(this.slide.layout, this.size.w, this.size.h).find((p) => p.ph === el.ph);
          if (fresh?.style?.size) keep.size = fresh.style.size;
        }
        el.style = Object.fromEntries(Object.entries(keep).filter(([, v]) => v != null));
        for (const p of props) {
          if (el.text) el.text = stripInline(el.text, p);
          if (el.cells) for (const row of el.cells) for (const cell of row) cell.text = stripInline(cell.text, p);
        }
      }
    });
  }

  /** Shortcuts that work while typing in a text box (the rest are the browser's). */
  textShortcut(e) {
    const k = (e.key || '').toLowerCase();
    const run = (fn) => { e.preventDefault(); fn(); };
    if (e.shiftKey && (k === '>' || k === '.')) run(() => this.stepFontSize(1));
    else if (e.shiftKey && (k === '<' || k === ',')) run(() => this.stepFontSize(-1));
    else if (e.shiftKey && k === 'l') run(() => this.setTextProp('align', 'left'));
    else if (e.shiftKey && k === 'e') run(() => this.setTextProp('align', 'center'));
    else if (e.shiftKey && k === 'r') run(() => this.setTextProp('align', 'right'));
    else if (e.shiftKey && k === 'j') run(() => this.setTextProp('align', 'justify'));
    else if (e.shiftKey && (k === '8' || k === '*')) run(() => this.toggleList('ul'));
    else if (e.shiftKey && (k === '7' || k === '&')) run(() => this.toggleList('ol'));
    else if (k === ']') run(() => this.indent(1));
    else if (k === '[') run(() => this.indent(-1));
    else if (k === '\\') run(() => this.clearFormatting());
    else if (k === 's') run(() => this.saveToFile({ as: e.shiftKey }));
    else if (k === 'p') run(() => this.printDeck());
    else if (k === 'enter') run(() => this.present({ fromStart: false }));
    else if (k === 'm') run(() => this.addSlide());
    else return;
    e.stopPropagation();
    this.toolbar.update();
  }

  // ---------------------------------------------------------------- pictures & tables

  async insertImage() {
    const file = await pickFileInput('image/*');
    if (file) this.insertImageFiles([file]);
  }

  async insertImageFiles(files, at = null) {
    const { w, h: H } = this.size;
    const els = [];
    for (const [i, file] of files.entries()) {
      try {
        const img = await readImageFile(file);
        const media = addMedia(this.deck, img.dataUrl);
        const k = Math.min(1, (w * 0.6) / img.width, (H * 0.6) / img.height);
        const iw = Math.round(img.width * k);
        const ih = Math.round(img.height * k);
        const cx = at ? at.x : w / 2;
        const cy = at ? at.y : H / 2;
        const el = newImage(media, Math.round(cx - iw / 2 + i * 16), Math.round(cy - ih / 2 + i * 16), iw, ih);
        el.alt = file.name.replace(/\.[^.]+$/, '');
        el.natural = { w: img.width, h: img.height };
        els.push(el);
      } catch (err) {
        toast(`${file.name}: ${err.message || 'couldn’t read this picture'}`, { type: 'error' });
      }
    }
    if (els.length) this.addElements(els);
  }

  async replaceImage() {
    const el = this.selectedElements().find((e) => e.type === 'image');
    if (!el) return;
    const file = await pickFileInput('image/*');
    if (!file) return;
    try {
      const img = await readImageFile(file);
      const media = addMedia(this.deck, img.dataUrl);
      this.change(() => {
        el.media = media;
        el.natural = { w: img.width, h: img.height };
        delete el.crop; // the old picture's crop means nothing for the new one
        // Keep the width and position, take the new picture's proportions.
        const cy = el.y + el.h / 2;
        el.h = Math.round((el.w * img.height) / img.width);
        el.y = Math.round(cy - el.h / 2);
      });
    } catch (err) {
      toast(err.message || 'Couldn’t read that picture.', { type: 'error' });
    }
  }

  resetImageSize() {
    const el = this.selectedElements().find((e) => e.type === 'image');
    if (!el) return;
    const probe = new Image();
    probe.onload = () => {
      const { w, h: H } = this.size;
      const k = Math.min(1, w / probe.naturalWidth, H / probe.naturalHeight);
      this.change(() => {
        const cx = el.x + el.w / 2;
        const cy = el.y + el.h / 2;
        el.w = Math.round(probe.naturalWidth * k);
        el.h = Math.round(probe.naturalHeight * k);
        el.x = Math.round(cx - el.w / 2);
        el.y = Math.round(cy - el.h / 2);
        el.rot = 0;
      });
    };
    probe.src = this.deck.media[el.media];
  }

  insertTable(rows, cols) {
    const { w, h: H } = this.size;
    const tw = Math.min(w * 0.84, cols * 140);
    const th = rows * 36;
    this.addElements([newTable(rows, cols, Math.round((w - tw) / 2), Math.round((H - th) / 2), Math.round(tw), th)]);
  }

  /** A sample column chart in the middle of the slide, with its data open in the side panel. */
  insertChart() {
    const { w, h: H } = this.size;
    const cw = Math.round(Math.min(w * 0.6, 560));
    const ch = Math.round(Math.min(H * 0.6, 340));
    this.addElements([newChart(Math.round((w - cw) / 2), Math.round((H - ch) / 2), cw, ch)]);
    this.openPanel('format');
  }

  tableOp(op) {
    const el = this.selectedElements().find((e) => e.type === 'table');
    if (!el) return;
    let r = el.cells.length - 1;
    let c = el.cells[0].length - 1;
    if (this.canvas.editing) {
      const sel = document.getSelection();
      const node = sel?.anchorNode?.nodeType === 3 ? sel.anchorNode.parentElement : sel?.anchorNode;
      const td = node?.closest?.('td');
      if (td) { r = +td.dataset.r; c = +td.dataset.c; }
      this.canvas.endEdit();
    } else if (op === 'rowAbove' || op === 'colLeft') { r = 0; c = 0; }
    const rowH = el.h / el.cells.length;
    const cols = el.cells[0].length;
    this.change(() => {
      if (op === 'rowAbove' || op === 'rowBelow') {
        el.cells.splice(op === 'rowAbove' ? r : r + 1, 0, Array.from({ length: cols }, () => ({ text: '' })));
        el.h += rowH;
      } else if (op === 'colLeft' || op === 'colRight') {
        const at = op === 'colLeft' ? c : c + 1;
        for (const row of el.cells) row.splice(at, 0, { text: '' });
        el.colWidths = Array.from({ length: cols + 1 }, () => 1 / (cols + 1));
      } else if (op === 'deleteRow') {
        if (el.cells.length === 1) return;
        el.cells.splice(r, 1);
        el.h -= rowH;
      } else if (op === 'deleteCol') {
        if (cols === 1) return;
        for (const row of el.cells) row.splice(c, 1);
        el.colWidths = Array.from({ length: cols - 1 }, () => 1 / (cols - 1));
      }
    });
  }

  // ---------------------------------------------------------------- animations

  nextAnimOrder() {
    return Math.max(0, ...this.slide.elements.map((e) => e.anim?.order || 0)) + 1;
  }

  setAnimation(props, ids = null) {
    const set = new Set(ids || this.selection);
    const els = this.slide.elements.filter((e) => set.has(e.id));
    if (!els.length) return;
    this.change(() => {
      let order = this.nextAnimOrder();
      for (const el of els) {
        if (props.effect === 'none') {
          delete el.anim;
          continue;
        }
        // Duration or start alone doesn't animate something that isn't animated.
        if (!el.anim && !props.effect) continue;
        const fresh = !el.anim;
        el.anim = { effect: 'fade', duration: 0.5, start: 'click', ...(el.anim || {}), ...props };
        if (fresh) el.anim.order = order++;
      }
    }, { render: true });
    this.panel.update(true);
  }

  reorderAnimation(id, dir) {
    const list = this.slide.elements.filter((e) => e.anim).sort((a, b) => a.anim.order - b.anim.order);
    const i = list.findIndex((e) => e.id === id);
    const j = i + dir;
    if (i < 0 || j < 0 || j >= list.length) return;
    this.change(() => {
      [list[i], list[j]] = [list[j], list[i]];
      list.forEach((e, k) => { e.anim.order = k + 1; });
    });
    this.panel.update(true);
  }

  async previewAnimations() {
    const { previewOnStage } = await import('../present/present.js');
    previewOnStage(this, { animations: true });
  }

  async previewTransition() {
    const { previewOnStage } = await import('../present/present.js');
    previewOnStage(this, { transition: true });
  }

  // ---------------------------------------------------------------- panels, view

  openPanel(name) {
    if (this.panel.current === name) this.panel.update(true);
    else this.panel.open(name);
  }

  closePanel() { this.panel.close(); }

  toggleNotes() {
    this.view.notes = !this.view.notes;
    this.notesPane.hidden = !this.view.notes;
    try { localStorage.setItem('ls:notes', this.view.notes ? '1' : '0'); } catch { /* storage blocked */ }
    this.canvas.layout();
  }

  toggleSnap() {
    this.canvas.snapEnabled = !this.canvas.snapEnabled;
    toast(this.canvas.snapEnabled ? 'Snapping to guides is on' : 'Snapping to guides is off (hold Alt while dragging to turn it off temporarily)');
  }

  zoomBy(k) {
    this.canvas.setZoom(Math.max(0.1, Math.min(4, this.canvas.scale * k)));
  }

  isDarkUi() { return isDark(); }

  toggleDarkUi() {
    toggleTheme();
    this.darkBtn.innerHTML = icon(isDark() ? 'sun' : 'moon');
  }

  toggleFullscreen() {
    if (document.fullscreenElement) document.exitFullscreen?.();
    else document.documentElement.requestFullscreen?.().catch(() => {});
  }

  canvasMenu(at) {
    const has = this.selection.length > 0;
    const items = has ? [
      { label: 'Cut', icon: 'cut', shortcut: 'Mod-X', run: () => this.copySelection({ cut: true }) },
      { label: 'Copy', icon: 'copy', shortcut: 'Mod-C', run: () => this.copySelection() },
      { label: 'Paste', icon: 'paste', shortcut: 'Mod-V', run: () => this.pasteFromMenu() },
      { label: 'Duplicate', icon: 'duplicate', shortcut: 'Mod-D', run: () => this.duplicateSelection() },
      { label: 'Delete', icon: 'trash', shortcut: 'Delete', run: () => this.deleteSelection() },
      'separator',
      { label: 'Bring to front', icon: 'toFront', shortcut: 'Mod-Shift-Up', run: () => this.order('front') },
      { label: 'Send to back', icon: 'toBack', shortcut: 'Mod-Shift-Down', run: () => this.order('back') },
      'separator',
      { label: 'Animate…', icon: 'animate', run: () => this.openPanel('animate') },
      { label: 'Format options…', icon: 'format', run: () => this.openPanel('format') },
    ] : [
      { label: 'Paste', icon: 'paste', shortcut: 'Mod-V', run: () => this.pasteFromMenu() },
      { label: 'Select all', shortcut: 'Mod-A', run: () => this.selectAll() },
      'separator',
      { label: 'New slide', icon: 'plus', shortcut: 'Mod-M', run: () => this.addSlide() },
      { label: 'Change background…', icon: 'image', run: () => this.openPanel('design') },
      { label: 'Change theme…', icon: 'theme', run: () => this.openPanel('design') },
      { label: 'Transition…', icon: 'animate', run: () => this.openPanel('transition') },
    ];
    showPopover(null, menu(items, { iconFn: (n) => icon(n) }), { at, label: 'Actions' });
  }

  // ---------------------------------------------------------------- keyboard

  handleKey(e) {
    if (this.presentation?.active) return;
    if (document.querySelector('dialog[open]') || isPopoverOpen()) return;
    if (!e.key) return;
    const mod = e.ctrlKey || e.metaKey;
    const k = e.key.toLowerCase();
    const run = (fn) => { e.preventDefault(); fn(); };

    // Anywhere in the editor
    if (mod && k === 'o') return run(() => this.openFile());
    if (e.key === 'F5') return run(() => this.present({ fromStart: !e.shiftKey }));
    if (mod && k === 'enter') return run(() => this.present({ fromStart: e.shiftKey }));
    if (mod && k === 's') return run(() => this.saveToFile({ as: e.shiftKey }));
    if (mod && k === 'p') return run(() => this.printDeck());
    if (mod && k === 'm') return run(() => this.addSlide());
    if (mod && k === '/') return run(() => this.showShortcuts());
    if (mod && (k === '=' || k === '+')) return run(() => this.zoomBy(1.25));
    if (mod && k === '-') return run(() => this.zoomBy(0.8));
    // By key position: Alt changes the character typed on a Mac (Alt+0 is º).
    if (mod && e.altKey && e.code === 'Digit0') return run(() => this.canvas.setZoom('fit'));

    // Strikethrough also while typing on the slide, but Alt+Shift+5 still types in other fields.
    if (e.altKey && e.shiftKey && !mod && e.code === 'Digit5' && this.hasTextTarget() && (!typingTarget(e.target) || this.canvas.wrap.contains(e.target))) {
      return run(() => this.toggleTextProp('strike'));
    }

    if (typingTarget(e.target)) return;
    if (mod && !e.shiftKey && k === 'z') return run(() => this.undo());
    if (mod && (k === 'y' || (e.shiftKey && k === 'z'))) return run(() => this.redo());
    if (this.filmstrip.el.contains(e.target) || this.panel.el.contains(e.target)) return;
    if (this.toolbar.toolbar.contains(e.target) || this.toolbar.menubar.contains(e.target)) {
      if (e.key !== 'Escape') return;
    }
    // Space and Enter press a focused button (in the title bar…); arrows and Delete aren't for the slide either.
    if (e.key !== 'Escape' && e.target.closest?.('button, a, [role=button]')) return;

    const els = this.selectedElements();
    if (e.key === 'Escape') {
      if (this.canvas.tool) return run(() => this.canvas.setTool(null));
      if (this.selection.length) return run(() => this.setSelection([]));
      return;
    }
    if (e.key === 'PageDown') return run(() => this.goToSlide(this.slideIndex + 1));
    if (e.key === 'PageUp') return run(() => this.goToSlide(this.slideIndex - 1));
    if (!els.length) {
      if (e.key === 'ArrowDown' || e.key === 'ArrowRight') return run(() => this.goToSlide(this.slideIndex + 1));
      if (e.key === 'ArrowUp' || e.key === 'ArrowLeft') return run(() => this.goToSlide(this.slideIndex - 1));
      if (e.key === 'Home') return run(() => this.goToSlide(0));
      if (e.key === 'End') return run(() => this.goToSlide(this.deck.slides.length - 1));
      if (mod && k === 'a') return run(() => this.selectAll());
      if (e.key === 'Tab' && !e.shiftKey && this.slide.elements.length && e.target === this.canvas.wrap) return run(() => this.setSelection([this.slide.elements[0].id]));
      if (mod && k === 'd') return run(() => this.duplicateSlides());
      return;
    }
    if (e.key === 'Delete' || e.key === 'Backspace') return run(() => this.deleteSelection());
    if (mod && k === 'a') return run(() => this.selectAll());
    if (mod && k === 'd') return run(() => this.duplicateSelection());
    if (e.key === 'Tab') {
      const list = this.slide.elements;
      const i = list.findIndex((x) => x.id === this.selection[this.selection.length - 1]) + (e.shiftKey ? -1 : 1);
      // Past the first or last object, focus leaves the canvas as usual.
      if (i < 0 || i >= list.length) {
        this.setSelection([]);
        return;
      }
      return run(() => this.setSelection([list[i].id]));
    }
    if (mod && e.key === 'ArrowUp') return run(() => this.order(e.shiftKey ? 'front' : 'forward'));
    if (mod && e.key === 'ArrowDown') return run(() => this.order(e.shiftKey ? 'back' : 'backward'));
    if (e.key.startsWith('Arrow') && !mod) {
      const step = e.shiftKey ? 10 : e.altKey ? 0.5 : 1;
      const dx = e.key === 'ArrowLeft' ? -step : e.key === 'ArrowRight' ? step : 0;
      const dy = e.key === 'ArrowUp' ? -step : e.key === 'ArrowDown' ? step : 0;
      return run(() => this.updateSelected((el) => { if (!el.locked) moveElement(el, dx, dy); }, { coalesce: 'nudge' }));
    }
    const single = els.length === 1 ? els[0] : null;
    if (single && isTextual(single) && (e.key === 'Enter' || e.key === 'F2')) return run(() => this.canvas.startEdit(single.id, { selectAll: e.key === 'F2' }));
    if (mod && !e.altKey) {
      if (!e.shiftKey && k === 'b') return run(() => this.toggleTextProp('bold'));
      if (!e.shiftKey && k === 'i') return run(() => this.toggleTextProp('italic'));
      if (!e.shiftKey && k === 'u') return run(() => this.toggleTextProp('underline'));
      if (this.hasTextTarget()) {
        const before = e.defaultPrevented;
        this.textShortcut(e);
        if (e.defaultPrevented && !before) return;
      }
    }
    // Typing on a selected text box starts editing it.
    if (single && canHoldText(single) && !mod && !e.altKey && e.key.length === 1 && !single.locked) {
      e.preventDefault();
      this.canvas.startEdit(single.id);
      exec('insertText', e.key);
    }
  }

  undo() {
    if (this.canvas.editing) this.canvas.endEdit();
    const snap = this.history.undo(this.snapshot());
    if (snap) this.restore(snap);
  }

  redo() {
    if (this.canvas.editing) this.canvas.endEdit();
    const snap = this.history.redo(this.snapshot());
    if (snap) this.restore(snap);
  }

  // ---------------------------------------------------------------- saving to files

  updateFileChip() {
    const f = this.deck.file;
    const chip = this.fileChip;
    if (!chip) return;
    chip.classList.toggle('is-dirty', !!f?.dirty);
    chip.classList.toggle('is-unlinked', !f);
    if (f) {
      chip.innerHTML = `${icon('fileDown')}<span class="file-chip-name"></span><span class="file-chip-state">${f.dirty ? 'Unsaved changes' : 'Saved'}</span>`;
      chip.querySelector('.file-chip-name').textContent = f.name;
      chip.title = f.dirty ? `Save changes to ${f.name} (${shortcutLabel('Mod-S')})` : `Saved to ${f.name} on this device`;
      chip.setAttribute('aria-label', `${f.name}: ${f.dirty ? 'unsaved changes, click to save' : 'saved'}`);
    } else {
      chip.innerHTML = `${icon('fileDown')}<span class="file-chip-name">Save to file…</span>`;
      chip.title = `Save this presentation as a file on your device (${shortcutLabel('Mod-Shift-S')})`;
      chip.setAttribute('aria-label', 'Save to a file on this device');
    }
  }

  /**
   * Save: write to the linked file in its format, or ask where to save
   * (Save As) when there's none. Without the File System Access API, a file
   * is downloaded. Resolves true when the file was written.
   */
  async saveToFile({ as = false } = {}) {
    if (this.canvas.editing) this.canvas.endEdit();
    if (this.writingFile) return false;
    const link = this.deck.file;
    if (!as && link?.handle) {
      // Ask for permission first, while the click or key press still counts as a user gesture.
      if (!(await ensureWritePermission(link.handle))) {
        toast(`LibreSlides wasn’t allowed to save to “${link.name}”.`, { type: 'error', timeout: 8000, action: { label: 'Save as…', run: () => this.saveToFile({ as: true }) } });
        return false;
      }
      const modified = await fileModified(link.handle);
      if (modified == null) {
        toast(`“${link.name}” can’t be found any more (moved or deleted?). Choose where to save it.`, { timeout: 7000 });
        return this.saveToFile({ as: true });
      }
      if (link.lastModified && Math.abs(modified - link.lastModified) > 1000) {
        const ok = await confirmDialog(`“${link.name}” was changed by another app since LibreSlides last opened or saved it. Replace it with this version?`, { title: 'File changed on disk', confirmLabel: 'Replace' });
        if (!ok) return false;
      }
      return this.writeFile(link.handle, link.format, link.name);
    }
    const preferred = link?.format || 'pptx';
    if (!canSaveToFiles()) return this.downloadFile(preferred);
    let handle;
    try {
      handle = await pickSaveLocation(fileNameFor(this.deck.title, preferred), preferred);
    } catch (err) {
      console.warn(err);
      return this.downloadFile(preferred);
    }
    if (!handle) return false;
    const format = formatOfName(handle.name) || preferred;
    return this.writeFile(handle, format, handle.name, { rename: true });
  }

  async writeFile(handle, format, name, { rename = false } = {}) {
    this.writingFile = true;
    const changesBefore = this.changeCount || 0;
    const firstTime = this.deck.file?.format !== format || this.deck.file?.handle !== handle;
    try {
      const blob = await deckBlob(this.deck, format);
      const lastModified = await writeToHandle(handle, blob);
      if (rename) {
        const title = name.replace(/\.[^.]+$/, '');
        this.deck.title = title;
        this.titleInput.value = title;
        document.title = `${title} – LibreSlides`;
      }
      // Edits made while the file was being written aren't in it yet.
      const dirty = (this.changeCount || 0) !== changesBefore;
      this.deck.file = { handle, name, format, lastModified, dirty };
      this.updateFileChip();
      this.dirty = true;
      await this.save(); // remember the link
      const note = firstTime ? FORMAT_NOTES[format] : null;
      toast(note ? `Saved to “${name}”. ${note}` : `Saved to “${name}”`, { type: 'success', timeout: note ? 8000 : 2500 });
      return true;
    } catch (err) {
      console.error(err);
      toast(`Couldn’t save “${name}”: ${err.message || err}`, { type: 'error', timeout: 8000, action: { label: 'Save as…', run: () => this.saveToFile({ as: true }) } });
      return false;
    } finally {
      this.writingFile = false;
    }
  }

  /** Browsers without the File System Access API: choose a format and download it. */
  async downloadFile(preferred = 'pptx') {
    const choice = await openDialog({
      title: 'Save to a file',
      body: h('div', {},
        h('p', {}, 'Your browser saves files as downloads. Your work is also kept in LibreSlides on this device.'),
        h('p', { class: 'pf-hint' }, FORMAT_NOTES.pptx)),
      buttons: [
        { label: 'Cancel', value: null },
        { label: 'LibreSlides (.lslides)', value: 'lslides' },
        { label: 'PowerPoint (.pptx)', value: 'pptx', primary: true },
      ],
      initialFocus: preferred === 'lslides' ? '.dialog-actions .btn:nth-child(2)' : '.btn-primary',
    });
    if (!choice) return false;
    const format = choice === true ? 'pptx' : choice;
    try {
      downloadBlob(await deckBlob(this.deck, format), fileNameFor(this.deck.title, format));
      return true;
    } catch (err) {
      toast(`Couldn’t create the file: ${err.message || err}`, { type: 'error' });
      return false;
    }
  }

  /** Before closing the presentation: offer to save changes not yet written to its file. */
  async confirmLeave() {
    if (this.canvas?.editing) this.canvas.endEdit();
    const f = this.deck?.file;
    if (!f?.dirty) return true;
    const choice = await openDialog({
      title: 'Save changes?',
      body: h('p', {}, `Save your changes to “${f.name}” before closing? They’re kept in LibreSlides either way.`),
      buttons: [
        { label: 'Cancel', value: null },
        { label: 'Don’t save', value: 'discard' },
        { label: 'Save', value: 'save', primary: true },
      ],
    });
    if (choice === 'save') return this.saveToFile();
    return choice === 'discard';
  }

  // ---------------------------------------------------------------- files

  openFile() { this.options.onOpenFile?.(); }
  newPresentation() { this.options.onNewDeck?.(); }

  async makeCopy() {
    await this.flush();
    this.options.onCopyDeck?.(this.deck.id);
  }

  async downloadPptx() {
    await this.flush();
    toast('Preparing PowerPoint file…', { timeout: 2000 });
    try {
      const { exportPptx } = await import('../io/pptx-export.js');
      await exportPptx(this.deck);
    } catch (err) {
      console.error(err);
      toast(`Couldn’t create the PowerPoint file: ${err.message || err}`, { type: 'error', timeout: 8000 });
    }
  }

  async downloadNative() {
    await this.flush();
    const { saveNativeFile } = await import('../io/files.js');
    saveNativeFile(this.deck);
  }

  async printDeck({ pdf = false } = {}) {
    await this.flush();
    const { printDeck } = await import('../io/print.js');
    if (pdf) toast('Choose “Save as PDF” as the printer to download a PDF.', { timeout: 6000 });
    printDeck(this.deck);
  }

  // ---------------------------------------------------------------- presenting

  async present({ fromStart = false } = {}) {
    // One show at a time (a held F5 repeats while the show is still loading).
    if (this.presentation || this.presentPending) return;
    this.presentPending = true;
    try {
      await this.flush();
      const { startSlideshow } = await import('../present/present.js');
      this.presentation = startSlideshow(this.deck, {
        start: fromStart ? 0 : this.slideIndex,
        onExit: (i) => {
          this.presentation = null;
          if (i != null) this.goToSlide(i);
          this.focusCanvas();
        },
      });
    } finally {
      this.presentPending = false;
    }
  }

  async presenterView() {
    if (this.presentation || this.presentPending) return;
    this.presentPending = true;
    try {
      await this.flush();
      const { startPresenterView } = await import('../present/present.js');
      this.presentation = startPresenterView(this.deck, {
        start: this.slideIndex,
        onExit: (i) => {
          this.presentation = null;
          if (i != null) this.goToSlide(i);
          this.focusCanvas();
        },
      });
    } finally {
      this.presentPending = false;
    }
  }

  // ---------------------------------------------------------------- dialogs

  showShortcuts() {
    const rows = [
      ['Present from this slide / from the beginning', 'Mod-Enter', 'F5'],
      ['New slide', 'Mod-M'],
      ['Duplicate', 'Mod-D'],
      ['Undo / Redo', 'Mod-Z', 'Mod-Y'],
      ['Cut / Copy / Paste', 'Mod-X', 'Mod-C', 'Mod-V'],
      ['Select all', 'Mod-A'],
      ['Edit the selected text box', 'Enter'],
      ['Stop editing text / deselect', 'Escape'],
      ['Select the next object', 'Tab'],
      ['Nudge (fine / large steps)', 'Arrow keys', 'Alt-Arrow', 'Shift-Arrow'],
      ['Bring forward / to front', 'Mod-Up', 'Mod-Shift-Up'],
      ['Send backward / to back', 'Mod-Down', 'Mod-Shift-Down'],
      ['Bold / Italic / Underline', 'Mod-B', 'Mod-I', 'Mod-U'],
      ['Bigger / smaller text', 'Mod-Shift-.', 'Mod-Shift-,'],
      ['Align left / center / right', 'Mod-Shift-L', 'Mod-Shift-E', 'Mod-Shift-R'],
      ['Bulleted / numbered list', 'Mod-Shift-8', 'Mod-Shift-7'],
      ['Previous / next slide', 'Page Up', 'Page Down'],
      ['Move slide up / down (in the slide list)', 'Mod-Up', 'Mod-Down'],
      ['Zoom in / out / fit', 'Mod-=', 'Mod--', 'Mod-Alt-0'],
      ['Keep proportions while resizing; draw straight lines', 'Shift-Drag'],
      ['Move without snapping', 'Alt-Drag'],
      ['Print / Open', 'Mod-P', 'Mod-O'],
    ];
    const table = h('table', { class: 'shortcut-table' }, h('tbody', {}, ...rows.map(([label, ...keys]) => h('tr', {}, h('td', {}, label), h('td', {}, ...keys.flatMap((kk, i) => [i ? ' ' : null, h('kbd', {}, /^[A-Z][a-z]/.test(kk) && !kk.includes('-') ? kk : shortcutLabel(kk))]))))));
    const present = h('table', { class: 'shortcut-table' }, h('tbody', {}, ...[
      ['Next / previous', '→ Space Enter', '← Backspace'],
      ['First / last slide', 'Home', 'End'],
      ['Go to slide number', 'Number then Enter'],
      ['Black / white screen', 'B', 'W'],
      ['Show speaker notes', 'S'],
      ['End the slideshow', 'Esc'],
    ].map(([label, ...keys]) => h('tr', {}, h('td', {}, label), h('td', {}, ...keys.flatMap((kk, i) => [i ? ' ' : null, h('kbd', {}, kk)]))))));
    openDialog({ title: 'Keyboard shortcuts', className: 'wide', body: h('div', {}, table, h('h3', { class: 'dialog-subhead' }, 'During a slideshow'), present) });
  }

  showInstall() { showInstallHelp(); }

  async showAbout() {
    const report = await diagnostics(this);
    const pre = h('pre', { class: 'diagnostics' }, report);
    openDialog({
      title: 'About LibreSlides',
      className: 'wide',
      body: h('div', {},
        h('p', {}, `LibreSlides ${VERSION} (build ${BUILD}) — a presentation editor that runs entirely in your browser and works offline. Presentations are kept on this device; nothing is uploaded.`),
        h('details', {}, h('summary', {}, 'Details for a bug report'), pre)),
      buttons: [
        { label: 'Copy details', value: 'copy' },
        { label: 'Close', value: null, primary: true },
      ],
      onSubmit: (v) => {
        if (v === 'copy') {
          navigator.clipboard?.writeText(report).then(() => toast('Copied'), () => {});
          return false;
        }
        return true;
      },
    });
  }
}

function escapeText(s) {
  return String(s).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
}

/** A text box in the middle of the slide, for pasted text. */
function newTextBoxAt(w, H) {
  return { id: uid(), type: 'text', x: Math.round(w * 0.15), y: Math.round(H * 0.4), w: Math.round(w * 0.7), h: 60, rot: 0, text: '', fill: 'none', stroke: null, style: { size: 18 } };
}
