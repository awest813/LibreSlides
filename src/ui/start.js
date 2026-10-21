import { h, formatDate, toast, debounce, showPopover, menu, downloadBlob, pickFileInput } from './dom.js';
import { icon, LOGO_SVG } from './icons.js';
import { listDecks, deleteDeck, duplicateDeck, renameDeck, exportBackup, importBackup, onStorageChange } from '../storage/db.js';
import { TEMPLATES } from '../model/templates.js';
import { renderThumb } from '../render/render.js';
import { sanitizeDeck } from '../io/sanitize.js';
import { confirmDialog, promptDialog } from './dialog.js';
import { isInstalledWindow, showInstallHelp } from './install.js';
import { isDark } from './theme.js';

const THUMB_W = 220;
const templateDecks = new Map(); // built once per session, for the thumbnails

/** Thumbnails are drawn THUMB_W wide; cards are narrower on phones, so size each one to its box. */
function fitThumb(box) {
  const frame = box.querySelector(':scope > .slide-thumb');
  const slide = frame?.firstElementChild;
  const width = box.clientWidth;
  if (!slide || !width) return;
  const k = width / parseFloat(slide.style.width);
  frame.style.width = `${width}px`;
  frame.style.height = `${Math.round(parseFloat(slide.style.height) * k)}px`;
  slide.style.transform = `scale(${k})`;
}
const fitThumbs = (root) => root.querySelectorAll('.template-thumb, .deck-cover').forEach(fitThumb);

function templateTiles({ onTemplate, onImport }) {
  const row = h('div', { class: 'template-row' });
  for (const t of TEMPLATES) {
    if (!templateDecks.has(t.id)) templateDecks.set(t.id, t.build());
    const deck = templateDecks.get(t.id);
    row.append(h('div', { class: 'template-tile' },
      h('div', { class: 'template-thumb', 'aria-hidden': 'true' }, renderThumb(deck.slides[0], deck, THUMB_W)),
      h('button', { type: 'button', class: 'template-card', onclick: () => onTemplate(t), 'aria-label': `New ${t.name}` }, h('span', { class: 'label' }, t.name))));
  }
  row.append(h('div', { class: 'template-tile' },
    h('div', { class: 'template-thumb is-import', 'aria-hidden': 'true', html: `<div class="import-inner">${icon('upload', 'icon icon-lg')}<span class="import-formats">PowerPoint (.pptx) or LibreSlides file</span></div>` }),
    h('button', { type: 'button', class: 'template-card', onclick: onImport, 'aria-label': 'Open from device', 'aria-description': 'PowerPoint or LibreSlides file' }, h('span', { class: 'label' }, 'Open from device…'))));
  return row;
}

function deckGrid({ onOpen, onDownload }) {
  const wrap = h('section', { 'aria-labelledby': 'recent-title' });
  const filter = h('input', { type: 'search', placeholder: 'Search presentations', 'aria-label': 'Search presentations' });
  const more = h('button', { type: 'button', class: 'icon-btn', title: 'Backup & restore', 'aria-label': 'Backup and restore', html: icon('more') });
  more.addEventListener('click', () => showPopover(more, menu([
    { label: 'Back up all presentations', icon: 'download', run: backup },
    { label: 'Restore from backup…', icon: 'upload', run: restore },
  ], { iconFn: (n) => icon(n) }), { placement: 'bottom-end' }));
  const body = h('div', {});
  wrap.append(
    h('div', { class: 'start-section-title' },
      h('h2', { id: 'recent-title' }, 'Recent presentations'),
      h('div', { class: 'doc-toolbar' }, h('label', { class: 'doc-filter', html: icon('search') }, filter), more)),
    body);
  let decks = [];

  const render = () => {
    const q = filter.value.trim().toLowerCase();
    const shown = q ? decks.filter((d) => d.title.toLowerCase().includes(q) || (d.preview || '').toLowerCase().includes(q)) : decks;
    body.replaceChildren();
    if (!decks.length) {
      body.append(h('div', { class: 'empty-state', html: `${icon('present', 'icon')}<p>No presentations yet. Pick a template above to get started.</p>` }));
      return;
    }
    if (!shown.length) {
      body.append(h('div', { class: 'empty-state' }, h('p', {}, `No presentations match “${filter.value}”.`)));
      return;
    }
    const grid = h('ul', { class: 'deck-grid' });
    for (const d of shown) {
      const actions = h('button', { type: 'button', class: 'icon-btn', 'aria-label': `More actions for ${d.title}`, html: icon('more') });
      actions.addEventListener('click', (e) => {
        e.stopPropagation();
        showPopover(actions, menu([
          { label: 'Open', icon: 'open', run: () => onOpen(d.id) },
          { label: 'Rename…', icon: 'rename', run: () => rename(d) },
          { label: 'Make a copy', icon: 'duplicate', run: () => duplicate(d) },
          { label: 'Download as PowerPoint', icon: 'fileDown', run: () => onDownload(d.id) },
          'separator',
          { label: 'Delete', icon: 'trash', run: () => remove(d) },
        ], { iconFn: (n) => icon(n) }), { placement: 'bottom-end' });
      });
      const cover = d.cover
        ? renderThumb(d.cover.slide, { themeId: d.cover.themeId, size: d.cover.size, media: d.cover.media || {} }, THUMB_W)
        : h('div', { class: 'slide-thumb' });
      const open = h('button', { type: 'button', class: 'deck-open', onclick: () => onOpen(d.id) }, h('span', { class: 'visually-hidden' }, `Open ${d.title}`));
      grid.append(h('li', {
        class: 'deck-card',
        onkeydown: (e) => {
          if (e.target !== open) return;
          if (e.key === 'Delete') remove(d);
        },
      },
      h('div', { class: 'deck-cover', 'aria-hidden': 'true' }, cover),
      h('div', { class: 'deck-meta' },
        h('div', { class: 'deck-text' },
          h('strong', {}, d.title || 'Untitled presentation'),
          h('small', {}, `${formatDate(d.updatedAt)} · ${d.slideCount} slide${d.slideCount === 1 ? '' : 's'}`),
          d.file ? h('small', { class: 'deck-file', title: `Saves to ${d.file.name} on this device` }, `${d.file.dirty ? '● ' : ''}${d.file.name}`) : null),
        actions),
      open));
    }
    body.append(grid);
    fitThumbs(grid);
  };

  const refresh = async () => {
    decks = await listDecks();
    render();
  };

  const rename = async (d) => {
    const r = await promptDialog({ title: 'Rename presentation', fields: [{ name: 'title', label: 'Name', value: d.title }], confirmLabel: 'Rename' });
    if (r && r.title.trim()) {
      await renameDeck(d.id, r.title.trim());
      refresh();
    }
  };
  const duplicate = async (d) => {
    await duplicateDeck(d.id);
    toast(`Copied “${d.title}”`);
    refresh();
  };
  const remove = async (d) => {
    if (!(await confirmDialog(`Delete “${d.title}”? This can’t be undone.`, { title: 'Delete presentation', confirmLabel: 'Delete', danger: true }))) return;
    await deleteDeck(d.id);
    toast(`Deleted “${d.title}”`);
    refresh();
  };

  async function backup() {
    const data = await exportBackup();
    const stamp = new Date().toISOString().slice(0, 10);
    downloadBlob(new Blob([JSON.stringify(data)], { type: 'application/json' }), `libreslides-backup-${stamp}.json`);
    toast(`Backed up ${data.decks.length} presentation${data.decks.length === 1 ? '' : 's'}.`, { type: 'success' });
  }

  async function restore() {
    const file = await pickFileInput('.json,application/json');
    if (!file) return;
    try {
      const r = await importBackup(JSON.parse(await file.text()), sanitizeDeck);
      toast(`Restored ${r.added} new and ${r.updated} updated presentation${r.added + r.updated === 1 ? '' : 's'}${r.skipped ? ` (${r.skipped} already up to date)` : ''}.`, { type: 'success', timeout: 5000 });
      refresh();
    } catch (err) {
      toast(err instanceof SyntaxError ? 'That file is not valid JSON.' : err.message, { type: 'error', timeout: 6000 });
    }
  }

  filter.addEventListener('input', debounce(render, 80));
  refresh();
  return { el: wrap, refresh };
}

let cleanup = null;

export function renderStartScreen(root, { onOpen, onTemplate, onImport, onToggleTheme, onDownload }) {
  cleanup?.();
  const grid = deckGrid({ onOpen, onDownload });
  // Hidden once LibreSlides runs as an installed app.
  const installBtn = h('button', {
    type: 'button', class: 'titlebar-btn', hidden: isInstalledWindow(),
    title: 'Install LibreSlides as an app: it gets its own window, works offline and opens .pptx files from your file manager',
    html: `${icon('download')}<span>Install app</span>`,
    onclick: () => showInstallHelp(),
  });
  // The same switch as in the editor: a sun in dark mode, a moon in light.
  const darkBtn = h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Dark mode', title: 'Dark mode' });
  const showTheme = () => {
    darkBtn.innerHTML = icon(isDark() ? 'sun' : 'moon');
    darkBtn.setAttribute('aria-pressed', String(isDark()));
  };
  darkBtn.addEventListener('click', () => {
    onToggleTheme();
    showTheme();
  });
  showTheme();
  const screen = h('div', { class: 'start-screen' },
    h('header', { class: 'titlebar' },
      h('span', { class: 'app-logo', html: LOGO_SVG('#c43e1c', '#fff', 28) }),
      h('h1', { class: 'app-name' }, 'LibreSlides'),
      h('div', { style: { flex: 1 } }),
      h('div', { class: 'titlebar-right' }, installBtn, darkBtn)),
    h('div', { class: 'start-scroll' },
      h('main', { class: 'start-inner' },
        h('div', { class: 'start-section-title' }, h('h2', {}, 'Start a new presentation')),
        templateTiles({ onTemplate, onImport }),
        grid.el,
        h('p', { class: 'start-foot' }, 'Presentations are saved in this browser and work offline. Nothing is uploaded.'))));
  root.replaceChildren(screen);
  const resize = new ResizeObserver(() => fitThumbs(screen));
  resize.observe(screen.querySelector('.start-inner'));
  const unStorage = onStorageChange(() => grid.refresh());
  const onInstalled = () => { installBtn.hidden = true; };
  window.addEventListener('appinstalled', onInstalled);
  cleanup = () => {
    resize.disconnect();
    unStorage();
    window.removeEventListener('appinstalled', onInstalled);
  };
  return { refresh: grid.refresh, destroy: () => cleanup?.() };
}
