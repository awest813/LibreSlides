import { downloadBlob, safeFileName } from '../ui/dom.js';
import { normalizeDeck, pruneMedia, uid } from '../model/deck.js';
import { sanitizeDeck } from './sanitize.js';

/**
 * Opening and saving presentation files: the native .lslides format (the
 * deck as JSON, pictures included) and PowerPoint .pptx.
 */
export const OPEN_ACCEPT = '.pptx,.lslides,.json,application/vnd.openxmlformats-officedocument.presentationml.presentation';

/** Version of the .lslides format this build writes (and the newest it reads). */
const NATIVE_FORMAT = 1;

export const extOf = (name = '') => (/\.([^.]+)$/.exec(name)?.[1] || '').toLowerCase();

/** Save a file: the system "Save as" dialog where there is one (Chrome, Edge), a download elsewhere. */
export async function saveFile(blob, name, { description, accept }) {
  if (typeof window.showSaveFilePicker === 'function') {
    try {
      const handle = await window.showSaveFilePicker({ suggestedName: name, types: [{ description, accept }] });
      const w = await handle.createWritable();
      await w.write(blob);
      await w.close();
      return true;
    } catch (err) {
      if (err?.name === 'AbortError') return false;
      // Not allowed here (e.g. inside a frame): fall back to a download.
    }
  }
  downloadBlob(blob, name);
  return true;
}

/** A deck as a plain copy for files and backups: no link to a file on this device. */
export function portableCopy(deck) {
  const { file: _linked, ...rest } = deck;
  return JSON.parse(JSON.stringify(rest));
}

/** The .lslides file for a deck. */
export function nativeBlob(deck) {
  const data = { app: 'LibreSlides', format: NATIVE_FORMAT, deck: pruneMedia(portableCopy(deck)) };
  return new Blob([JSON.stringify(data)], { type: 'application/vnd.libreslides+json' });
}

/** The deck as a file in `format` ('pptx' or 'lslides'). */
export async function deckBlob(deck, format) {
  if (format === 'lslides') return nativeBlob(deck);
  const { buildPptx } = await import('./pptx-export.js');
  return buildPptx(deck);
}

export function saveNativeFile(deck) {
  return saveFile(nativeBlob(deck), `${safeFileName(deck.title)}.lslides`, { description: 'LibreSlides presentation', accept: { 'application/vnd.libreslides+json': ['.lslides'] } });
}

/** Read a presentation file into a new deck (with a fresh id). */
export async function importFile(file) {
  const ext = extOf(file.name);
  if (ext === 'pptx' || ext === 'potx' || ext === 'ppsx') {
    const { importPptx } = await import('./pptx-import.js');
    const deck = await importPptx(await file.arrayBuffer(), file.name.replace(/\.[^.]+$/, ''));
    return normalizeDeck(sanitizeDeck(deck));
  }
  if (ext === 'lslides' || ext === 'json') {
    let data;
    try {
      data = JSON.parse(await file.text());
    } catch {
      throw Object.assign(new Error('This file is damaged or isn’t a LibreSlides presentation.'), { code: 'bad-file' });
    }
    if (data?.app === 'LibreSlides' && data.format > NATIVE_FORMAT) {
      throw Object.assign(new Error('This file was made by a newer version of LibreSlides. Reload the app to update it, then try again.'), { code: 'newer' });
    }
    const deck = data?.deck;
    if (data?.app !== 'LibreSlides' || !deck || typeof deck !== 'object' || !Array.isArray(deck.slides)) {
      if (Array.isArray(data?.decks)) throw Object.assign(new Error('This is a backup of several presentations. Restore it from the start screen (⋯ › Restore from backup).'), { code: 'backup' });
      throw Object.assign(new Error('This isn’t a LibreSlides presentation.'), { code: 'bad-file' });
    }
    deck.id = uid();
    delete deck.file;
    return normalizeDeck(sanitizeDeck(deck));
  }
  if (ext === 'ppt') throw Object.assign(new Error('Old PowerPoint 97–2003 files (.ppt) can’t be opened. Save it as .pptx in PowerPoint, LibreOffice or Google Slides first.'), { code: 'unsupported' });
  throw Object.assign(new Error('LibreSlides opens PowerPoint (.pptx) and LibreSlides (.lslides) files.'), { code: 'unsupported' });
}
