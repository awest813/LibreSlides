/**
 * Details for bug reports: version, browser, storage and the last errors,
 * copied from Help › About. Nothing from the presentation's text is included.
 */
/* global __APP_VERSION__, __BUILD_COMMIT__, __BUILD_DATE__ */
import { installState } from './ui/install.js';
export const VERSION = typeof __APP_VERSION__ !== 'undefined' ? __APP_VERSION__ : 'dev';
export const BUILD = typeof __BUILD_COMMIT__ !== 'undefined' ? `${__BUILD_COMMIT__} (${__BUILD_DATE__})` : 'dev';

const recent = [];
const MAX = 20;

/**
 * An error as text: its stack where it has one, else its name and message
 * (a DOMException such as QuotaExceededError has no stack), else what was
 * thrown or rejected with, as JSON when it's a plain object.
 */
function describe(err) {
  if (err?.stack) return String(err.stack);
  if (err?.message) return err.name ? `${err.name}: ${err.message}` : String(err.message);
  if (err && typeof err === 'object') {
    try {
      return JSON.stringify(err).slice(0, 300);
    } catch { /* circular */ }
  }
  return String(err);
}

/** Remember an error for the report (the browser console still gets it). */
export function noteError(err, where = '') {
  const message = describe(err);
  recent.push(`${new Date().toISOString()} ${where ? `[${where}] ` : ''}${message.split('\n').slice(0, 4).join(' ⏎ ')}`);
  if (recent.length > MAX) recent.shift();
}

/** Collect uncaught errors and rejections from LibreSlides' own code. */
export function watchErrors() {
  window.addEventListener('error', (e) => {
    // Extensions and cross-origin scripts report as "Script error." with no file: not ours.
    if (e.filename && !e.filename.startsWith(location.origin)) return;
    noteError(e.error || e.message, 'uncaught');
  });
  window.addEventListener('unhandledrejection', (e) => noteError(e.reason, 'unhandled promise'));
}

/** Plain-text report about the app, the browser and (optionally) the open presentation. */
export async function diagnostics(app = null) {
  const lines = [`LibreSlides ${VERSION}, build ${BUILD}`];
  lines.push(`Browser: ${navigator.userAgent}`);
  const installed = ['standalone', 'window-controls-overlay', 'minimal-ui'].some((m) => matchMedia(`(display-mode: ${m})`).matches);
  lines.push(`Window: ${installed ? 'installed app' : 'browser tab'}, ${innerWidth}×${innerHeight} at ${devicePixelRatio}x, ${navigator.onLine ? 'online' : 'offline'}`);
  try {
    const est = await navigator.storage?.estimate?.();
    const persisted = await navigator.storage?.persisted?.();
    if (est) lines.push(`Storage: ${(est.usage / 1048576).toFixed(1)} MB used of ${(est.quota / 1048576).toFixed(0)} MB${persisted ? ', persistent' : ''}`);
  } catch { /* not available */ }
  lines.push(`Service worker: ${navigator.serviceWorker?.controller ? 'active (works offline)' : 'not active'}`);
  lines.push(`Install: ${installState()}; saving to files: ${typeof window.showSaveFilePicker === 'function' ? 'saves back to linked files' : 'downloads'}`);
  if (app?.deck) {
    const counts = {};
    for (const s of app.deck.slides) for (const el of s.elements) counts[el.type] = (counts[el.type] || 0) + 1;
    lines.push(`Presentation: ${app.deck.slides.length} slides, ${app.deck.size}, theme ${app.deck.themeId}, ${Object.keys(app.deck.media || {}).length} pictures`);
    lines.push(`Content: ${Object.entries(counts).map(([k, v]) => `${k} ${v}`).join(', ') || 'empty'}`);
    lines.push(`Autosave: ${app.saveState}; linked file: ${app.deck.file ? `${app.deck.file.format}${app.deck.file.dirty ? ', unsaved changes' : ', saved'}` : 'none'}`);
  }
  lines.push(recent.length ? `Recent errors:\n  ${recent.join('\n  ')}` : 'Recent errors: none');
  return lines.join('\n');
}
