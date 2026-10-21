import './styles/app.css';
import { createDeck, getDeck, duplicateDeck, saveDeck, findDeckByFile } from './storage/db.js';
import { importFile, extOf } from './io/files.js';
import { pickFileToOpen, formatOfName, handleFromDataTransfer } from './io/file-access.js';
import { renderStartScreen } from './ui/start.js';
import { TEMPLATES } from './model/templates.js';
import { toast, h, shortcutLabel } from './ui/dom.js';
import { toggleTheme, syncThemeColor } from './ui/theme.js';
import { initInstall } from './ui/install.js';
import { watchErrors, noteError } from './diagnostics.js';
import { releaseMediaUrls } from './render/render.js';

watchErrors();
const root = document.getElementById('app');
let screen = null; // the open EditorScreen
let startScreen = null; // its listeners stop when the editor opens
let routing = Promise.resolve();

const go = (hash) => {
  if (location.hash === hash) route();
  else location.hash = hash;
};
const openDeck = (id) => go(`#/p/${encodeURIComponent(id)}`);

async function newFromTemplate(t) {
  const deck = t.build();
  openDeck(await createDeck(deck));
}

/**
 * Open a file from the device. With a file handle (Chrome, Edge, installed
 * app), the presentation stays linked to the file so Save writes back to it;
 * otherwise it's opened as a copy. Opening a file that's already linked to a
 * presentation goes back to that presentation. Resolves to the deck id or null.
 */
async function importAndOpen(file, handle = null, { show = true } = {}) {
  const open = (id) => {
    if (show) openDeck(id);
    return id;
  };
  try {
    const format = handle ? formatOfName(file.name) : null;
    const link = format ? { handle, name: file.name, format, lastModified: file.lastModified, dirty: false } : null;
    const existing = link ? await findDeckByFile(handle) : null;
    if (existing) {
      if (screen?.deckId === existing.id) await screen.flush();
      const deck = await getDeck(existing.id);
      // Unchanged on disk: reuse the presentation (it keeps everything the file format can't, and unsaved edits).
      if (deck.file && Math.abs((deck.file.lastModified || 0) - file.lastModified) < 1000) return open(deck.id);
      if (!deck.file?.dirty) {
        // Changed on disk and nothing newer here: reload the same presentation from the file.
        const fresh = await importFile(file);
        Object.assign(deck, { slides: fresh.slides, media: fresh.media, size: fresh.size, themeId: fresh.themeId, importedTheme: fresh.importedTheme, file: link });
        if (!deck.importedTheme) delete deck.importedTheme;
        await saveDeck(deck);
        await dropOpenCopy(deck.id);
        toast(`Reloaded “${file.name}”, which changed on disk`, { type: 'success' });
        return open(deck.id);
      }
      // Both changed: keep the unsaved version as its own presentation and open the file fresh.
      deck.file = null;
      deck.title = `${deck.title} (unsaved changes)`;
      await saveDeck(deck);
      await dropOpenCopy(deck.id);
      toast(`“${file.name}” changed on disk. Your unsaved version was kept as “${deck.title}”.`, { timeout: 8000 });
    }
    const deck = await importFile(file);
    if (link) {
      deck.file = link;
      deck.title = file.name.replace(/\.[^.]+$/, '');
    }
    const id = await createDeck(deck);
    if (show) {
      toast(link
        ? `Opened “${file.name}”. Save (${shortcutLabel('Mod-S')}) writes your changes back to it.`
        : `Opened a copy of “${file.name}”. Use File › Save as to save it to a file.`, { type: 'success', timeout: 6000 });
    }
    return open(id);
  } catch (err) {
    if (err?.code) console.warn(err);
    else console.error(err);
    noteError(err, `opening ${extOf(file.name) || 'file'}`);
    toast(`${file.name}: ${err.message || 'Couldn’t open that file.'}`, { type: 'error', timeout: 7000 });
    return null;
  }
}

/** After changing a deck in storage behind the open editor, close it without saving its stale copy. */
async function dropOpenCopy(id) {
  if (screen?.deckId !== id) return false;
  const stale = screen;
  screen = null;
  window.libreslides = null;
  await stale.destroy({ save: false });
  route();
  return true;
}

async function openFromDevice(file = null, handle = null) {
  if (file) return importAndOpen(file, handle);
  const picked = await pickFileToOpen();
  return picked ? importAndOpen(picked.file, picked.handle) : null;
}

async function downloadPptx(id) {
  const deck = await getDeck(id);
  if (!deck) return;
  const { exportPptx } = await import('./io/pptx-export.js');
  exportPptx(deck).catch((err) => toast(`Couldn’t create the PowerPoint file: ${err.message}`, { type: 'error' }));
}

const prefetchEditor = () => import('./editor/editor-screen.js').catch(() => {});

function showStart() {
  document.title = 'LibreSlides';
  (window.requestIdleCallback || ((fn) => setTimeout(fn, 1500)))(prefetchEditor);
  startScreen = renderStartScreen(root, {
    onOpen: openDeck,
    onTemplate: newFromTemplate,
    onImport: () => openFromDevice(),
    onToggleTheme: toggleTheme,
    onDownload: downloadPptx,
  });
}

async function showEditor(id) {
  const { EditorScreen } = await import('./editor/editor-screen.js');
  const s = new EditorScreen(root, {
    deckId: id,
    onHome: () => go('#/'),
    onOpenFile: (file, handle) => openFromDevice(file, handle),
    onNewDeck: () => newFromTemplate(TEMPLATES[0]),
    onCopyDeck: async (deckId) => {
      const copy = await duplicateDeck(deckId);
      if (copy) {
        toast('Opened a copy');
        openDeck(copy);
      }
    },
  });
  const ok = await s.mount();
  if (!ok) {
    toast('That presentation no longer exists.', { type: 'error' });
    go('#/');
    return;
  }
  screen = s;
  document.title = `${s.deck.title} – LibreSlides`;
  window.libreslides = { screen: s };
}

/** The audience window of presenter view. */
async function showAudience(id) {
  const deck = await getDeck(id);
  if (!deck) {
    root.replaceChildren(h('p', { class: 'empty-state' }, 'That presentation isn’t available.'));
    return;
  }
  document.title = `${deck.title} – Slideshow`;
  root.replaceChildren();
  const { runAudience } = await import('./present/present.js');
  runAudience(deck);
}

function route() {
  routing = routing.then(routeNow).catch((err) => {
    console.error(err);
    noteError(err, 'route');
    toast(`Something went wrong: ${err.message || err}`, { type: 'error', timeout: 8000 });
  });
  return routing;
}

/** Show what the address asks for (route() queues these, so one finishes before the next starts). */
async function routeNow() {
  const m = /^#\/(p|show)\/(.+)$/.exec(location.hash);
  let id = null;
  try {
    id = m ? decodeURIComponent(m[2]) : null;
  } catch {
    id = '';
  }
  // The "New presentation" app shortcut.
  if (location.hash === '#/new') {
    if (screen && !(await screen.confirmLeave())) {
      history.replaceState(null, '', `#/p/${encodeURIComponent(screen.deckId)}`);
      return;
    }
    const id = await createDeck(TEMPLATES[0].build());
    history.replaceState(null, '', `#/p/${encodeURIComponent(id)}`);
    // Not route(): that would wait for this step to finish.
    return routeNow();
  }
  if (m?.[1] === 'show') {
    root.removeAttribute('aria-busy');
    await showAudience(id);
    return;
  }
  if (screen && screen.deckId === id) return;
  // Back button, links, file launches…: offer to save to the linked file first.
  if (screen && !(await screen.confirmLeave())) {
    history.replaceState(null, '', `#/p/${encodeURIComponent(screen.deckId)}`);
    return;
  }
  if (screen) {
    await screen.destroy();
    screen = null;
    window.libreslides = null;
  }
  root.removeAttribute('aria-busy');
  if (m && !id) {
    history.replaceState(null, '', '#/');
    toast('That presentation link isn’t valid.', { type: 'error' });
  }
  startScreen?.destroy();
  startScreen = null;
  // The pictures of what was shown (a presentation, the list's covers) are no longer needed.
  releaseMediaUrls();
  if (id) await showEditor(id);
  else showStart();
}

window.addEventListener('libreslides:db-blocked', () => toast('Close other LibreSlides tabs to finish updating.', { timeout: 10000 }));
window.addEventListener('libreslides:db-outdated', () => toast('LibreSlides was updated in another tab. Reload this tab to keep saving.', {
  type: 'error', timeout: 60000, action: { label: 'Reload', run: () => location.reload() },
}));

syncThemeColor();
matchMedia('(prefers-color-scheme: dark)').addEventListener?.('change', syncThemeColor);
window.addEventListener('hashchange', route);
route();

// Drop a .pptx or .lslides file anywhere on the start screen to open it.
let dragDepth = 0;
let overlay = null;
const clearOverlay = () => {
  dragDepth = 0;
  overlay?.remove();
  overlay = null;
};
window.addEventListener('dragenter', (e) => {
  if (screen || !e.dataTransfer?.types?.includes('Files')) return;
  dragDepth++;
  if (!overlay) {
    overlay = h('div', { class: 'drop-overlay' }, 'Drop to open in LibreSlides');
    document.body.append(overlay);
  }
});
window.addEventListener('dragleave', () => {
  if (--dragDepth <= 0) clearOverlay();
});
window.addEventListener('dragover', (e) => {
  if (e.dataTransfer?.types?.includes('Files')) e.preventDefault();
});
window.addEventListener('drop', (e) => {
  clearOverlay();
  if (!e.dataTransfer?.files?.length) return;
  e.preventDefault(); // never navigate away to the dropped file
  if (screen) return; // the editor handles drops on its slide
  const file = e.dataTransfer.files[0];
  // The handle must be requested synchronously, during the drop event.
  handleFromDataTransfer(e.dataTransfer, file).then((handle) => openFromDevice(file, handle));
});

// "Open with LibreSlides" from the file manager once installed. Several files
// at once open the first and add the rest to the presentation list.
if ('launchQueue' in window) {
  window.launchQueue.setConsumer(async (params) => {
    let shown = false;
    let added = 0;
    for (const handle of params.files || []) {
      if (handle.kind !== 'file') continue;
      let file;
      try {
        file = await handle.getFile();
      } catch (err) {
        toast(`Couldn’t read “${handle.name}”: ${err.message || err}`, { type: 'error', timeout: 6000 });
        continue;
      }
      const id = await importAndOpen(file, handle, { show: !shown });
      if (id && shown) added++;
      if (id) shown = true;
    }
    if (added) toast(`${added} more presentation${added === 1 ? ' was' : 's were'} added to your list.`, { timeout: 6000 });
  });
}

// Ctrl+O on the start screen (the editor has its own shortcuts).
window.addEventListener('keydown', (e) => {
  if (screen || !(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey || e.key.toLowerCase() !== 'o') return;
  if (document.querySelector('dialog[open]')) return;
  e.preventDefault();
  openFromDevice();
});

initInstall({
  onInstalled: () => toast('LibreSlides is installed. Open it from your apps — it works offline, and can open .pptx files from your file manager.', { type: 'success', timeout: 9000 }),
});

// Service worker: offline support. Registered after load so it never competes with the first paint.
if ('serviceWorker' in navigator && import.meta.env.PROD) {
  window.addEventListener('load', () => {
    import('virtual:pwa-register').then(({ registerSW }) => {
      const updateSW = registerSW({
        immediate: true,
        onNeedRefresh: () => toast('A new version of LibreSlides is ready.', {
          timeout: 60000, action: { label: 'Update', run: () => updateSW() },
        }),
        onNeedReload: () => (screen ? screen.flush() : Promise.resolve()).catch(() => {}).finally(() => location.reload()),
        onOfflineReady: () => toast('LibreSlides is ready to work offline.', { type: 'success' }),
      });
    }).catch(() => {});
  });
}

// Ask the browser to keep our data even under storage pressure.
navigator.storage?.persist?.().catch(() => {});
