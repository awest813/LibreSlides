/**
 * Working with files on the user's device.
 *
 * Where the File System Access API is available (Chrome, Edge, Opera, and
 * installed apps on ChromeOS, Windows, macOS and Linux), a presentation opened
 * from disk keeps a handle to its file: Save writes back to that file in its
 * own format, and Save As links the presentation to the new file. Other
 * browsers import files as copies, and saving downloads a file instead.
 */
import { pickFileInput, safeFileName } from '../ui/dom.js';
import { OPEN_ACCEPT, extOf } from './files.js';

export const FILE_FORMATS = {
  pptx: {
    label: 'PowerPoint presentation',
    mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    exts: ['.pptx'],
  },
  lslides: {
    label: 'LibreSlides presentation',
    mime: 'application/vnd.libreslides+json',
    exts: ['.lslides'],
  },
};

/** What saving to a format can't keep, shown the first time a deck is saved that way. */
export const FORMAT_NOTES = {
  pptx: 'PowerPoint files keep slides, text, shapes, pictures, tables, notes and transitions, but not LibreSlides animations or theme switching.',
};

export const canSaveToFiles = () => typeof window !== 'undefined' && typeof window.showSaveFilePicker === 'function';
export const canOpenFileHandles = () => typeof window !== 'undefined' && typeof window.showOpenFilePicker === 'function';

/** Writable format of a file name, or null. */
export function formatOfName(name = '') {
  const ext = `.${extOf(name)}`;
  return Object.keys(FILE_FORMATS).find((id) => FILE_FORMATS[id].exts.includes(ext)) || null;
}

/** "Q3 review" → "Q3 review.pptx" (an existing presentation extension is replaced). */
export function fileNameFor(title, format) {
  const raw = String(title || 'Untitled presentation');
  const base = safeFileName(/\.(pptx|lslides|potx|ppsx)$/i.test(raw) ? raw.replace(/\.[^.]+$/, '') : raw);
  return `${base}${FILE_FORMATS[format]?.exts[0] || '.pptx'}`;
}

const isAbort = (err) => err?.name === 'AbortError';

/** Let the user pick a file to open. Resolves to { file, handle } (handle null without the API) or null. */
export async function pickFileToOpen() {
  if (canOpenFileHandles()) {
    try {
      const [handle] = await window.showOpenFilePicker({
        id: 'libreslides-open',
        multiple: false,
        types: [{
          description: 'Presentations',
          accept: {
            [FILE_FORMATS.pptx.mime]: ['.pptx'],
            [FILE_FORMATS.lslides.mime]: ['.lslides'],
          },
        }],
      });
      return { file: await handle.getFile(), handle };
    } catch (err) {
      if (isAbort(err)) return null;
      console.warn('showOpenFilePicker failed; using a file input instead', err);
    }
  }
  const file = await pickFileInput(OPEN_ACCEPT);
  return file ? { file, handle: null } : null;
}

/** Ask where to save. Resolves to a file handle, or null if cancelled. */
export async function pickSaveLocation(suggestedName, preferred = 'pptx') {
  const order = [preferred, ...Object.keys(FILE_FORMATS).filter((f) => f !== preferred)];
  try {
    return await window.showSaveFilePicker({
      id: 'libreslides-save',
      suggestedName,
      types: order.map((id) => ({ description: FILE_FORMATS[id].label, accept: { [FILE_FORMATS[id].mime]: FILE_FORMATS[id].exts } })),
    });
  } catch (err) {
    if (isAbort(err)) return null;
    throw err;
  }
}

/**
 * Make sure we may write to the file. Browsers forget the permission between
 * sessions; asking again needs a user gesture (the Save click or Ctrl+S), so
 * call this before any slow work.
 */
export async function ensureWritePermission(handle) {
  if (!handle.queryPermission) return true;
  const opts = { mode: 'readwrite' };
  try {
    if ((await handle.queryPermission(opts)) === 'granted') return true;
    return (await handle.requestPermission(opts)) === 'granted';
  } catch {
    return false;
  }
}

/** Write a blob to the file; returns the file's new lastModified time. */
export async function writeToHandle(handle, blob) {
  const writable = await handle.createWritable();
  try {
    await writable.write(blob);
    await writable.close();
  } catch (err) {
    await writable.abort?.().catch(() => {});
    throw err;
  }
  try {
    return (await handle.getFile()).lastModified;
  } catch {
    return Date.now();
  }
}

/** The file's current lastModified, or null if it can't be read (moved, deleted, no permission). */
export async function fileModified(handle) {
  try {
    return (await handle.getFile()).lastModified;
  } catch {
    return null;
  }
}

/**
 * The FileSystemFileHandle for a dropped file (Chromium). Must be called
 * synchronously inside the drop event; only a handle with the same name as
 * `file` is returned, so a multi-file drop can never link the wrong file.
 */
export function handleFromDataTransfer(dataTransfer, file) {
  const items = [...(dataTransfer?.items || [])].filter((i) => i.kind === 'file' && i.getAsFileSystemHandle);
  const pending = items.map((i) => i.getAsFileSystemHandle().catch(() => null));
  if (!file || !pending.length) return Promise.resolve(null);
  return Promise.all(pending).then((handles) => handles.find((h) => h?.kind === 'file' && h.name === file.name) || null);
}
