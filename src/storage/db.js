import { openDB } from 'idb';
import { normalizeDeck, slideText, uid, pruneMedia } from '../model/deck.js';

/**
 * Presentations are kept in IndexedDB in two stores:
 *   meta    – small records for the presentation list (title, dates, cover slide)
 *   content – the full deck
 * so the start screen never loads every deck.
 *
 * Each save bumps `rev`; an editor that finds a newer rev than the one it
 * loaded knows another tab changed the deck and stops autosaving over it.
 */
const DB_NAME = 'libreslides';
const DB_VERSION = 1;

const channel = typeof BroadcastChannel !== 'undefined' ? new BroadcastChannel('libreslides-db') : null;
const announce = (msg) => channel?.postMessage(msg);

/** Call `fn({type, id, rev})` when any tab changes stored presentations. Returns an unsubscribe function. */
export function onStorageChange(fn) {
  if (!channel) return () => {};
  const handler = (e) => fn(e.data);
  channel.addEventListener('message', handler);
  return () => channel.removeEventListener('message', handler);
}

let dbPromise;
export const getDb = () => {
  dbPromise ||= openDB(DB_NAME, DB_VERSION, {
    upgrade(db) {
      if (!db.objectStoreNames.contains('meta')) db.createObjectStore('meta', { keyPath: 'id' }).createIndex('updatedAt', 'updatedAt');
      if (!db.objectStoreNames.contains('content')) db.createObjectStore('content', { keyPath: 'id' });
    },
    blocked: () => globalThis.dispatchEvent?.(new CustomEvent('libreslides:db-blocked')),
    // A newer version in another tab wants to upgrade: get out of its way. This tab
    // can't reopen the database at its old version, so it needs a reload.
    blocking: (_oldVersion, _newVersion, event) => {
      event.target.close();
      dbPromise = null;
      globalThis.dispatchEvent?.(new CustomEvent('libreslides:db-outdated'));
    },
    terminated: () => { dbPromise = null; },
  }).catch((err) => {
    dbPromise = null;
    throw err;
  });
  return dbPromise;
};

/** The list record for a deck: enough to draw its cover thumbnail. */
function metaOf(deck) {
  const cover = deck.slides.find((s) => !s.hidden) || deck.slides[0];
  const media = {};
  if (cover) {
    if (cover.background?.media) media[cover.background.media] = deck.media?.[cover.background.media];
    for (const el of cover.elements) if (el.media && deck.media?.[el.media]) media[el.media] = deck.media[el.media];
  }
  return {
    id: deck.id,
    title: deck.title,
    createdAt: deck.createdAt,
    updatedAt: deck.updatedAt,
    rev: deck.rev || 0,
    slideCount: deck.slides.length,
    preview: deck.slides.slice(0, 3).map(slideText).join(' · ').slice(0, 200),
    // The file on this device the deck saves to (a FileSystemFileHandle survives in IndexedDB).
    file: deck.file ? { handle: deck.file.handle, name: deck.file.name, format: deck.file.format, dirty: !!deck.file.dirty } : null,
    cover: cover ? { slide: cover, themeId: deck.themeId, size: deck.size, media } : null,
  };
}

export async function listDecks() {
  const db = await getDb();
  const all = await db.getAll('meta');
  return all.sort((a, b) => b.updatedAt - a.updatedAt);
}

export async function getMeta(id) {
  return (await getDb()).get('meta', id);
}

export async function getDeck(id) {
  const db = await getDb();
  const row = await db.get('content', id);
  if (!row) return null;
  const deck = normalizeDeck(row.deck);
  // Nothing can undo back to pictures that were deleted before this session.
  pruneMedia(deck);
  return deck;
}

/**
 * Store a deck. With `expectRev`, refuses (returns {conflict: true, rev}) if
 * another tab saved a newer version since. Returns {rev} on success.
 */
export async function saveDeck(deck, { expectRev = null, touch = true } = {}) {
  const db = await getDb();
  const tx = db.transaction(['meta', 'content'], 'readwrite');
  const current = await tx.objectStore('meta').get(deck.id);
  if (expectRev != null && current && (current.rev || 0) !== expectRev) {
    await tx.done;
    return { conflict: true, rev: current.rev || 0 };
  }
  deck.rev = (current?.rev || 0) + 1;
  if (touch) deck.updatedAt = Date.now();
  await tx.objectStore('content').put({ id: deck.id, deck });
  await tx.objectStore('meta').put(metaOf(deck));
  await tx.done;
  announce({ type: 'saved', id: deck.id, rev: deck.rev });
  return { rev: deck.rev };
}

export async function createDeck(deck) {
  deck.id ||= uid();
  deck.createdAt = deck.updatedAt = Date.now();
  deck.rev = 0;
  await saveDeck(deck);
  return deck.id;
}

export async function renameDeck(id, title) {
  const deck = await getDeck(id);
  if (!deck) return;
  deck.title = title;
  await saveDeck(deck);
}

export async function duplicateDeck(id) {
  const deck = await getDeck(id);
  if (!deck) return null;
  const { file: _linked, ...rest } = deck;
  const copy = JSON.parse(JSON.stringify(rest));
  copy.id = uid();
  copy.title = `Copy of ${deck.title}`;
  return createDeck(copy);
}

/** The deck linked to the same file as `handle` (opening a file again finds it). */
export async function findDeckByFile(handle) {
  if (!handle?.isSameEntry) return null;
  for (const m of await listDecks()) {
    try {
      if (m.file?.handle && (await m.file.handle.isSameEntry(handle))) return m;
    } catch { /* a stale handle */ }
  }
  return null;
}

export async function deleteDeck(id) {
  const db = await getDb();
  const tx = db.transaction(['meta', 'content'], 'readwrite');
  await tx.objectStore('meta').delete(id);
  await tx.objectStore('content').delete(id);
  await tx.done;
  announce({ type: 'deleted', id });
}

/** Every deck in one object, for "Back up all presentations". */
export async function exportBackup() {
  const db = await getDb();
  const rows = await db.getAll('content');
  const decks = rows.map(({ deck: { file: _linked, ...rest } }) => rest);
  return { app: 'LibreSlides', version: 1, exportedAt: new Date().toISOString(), decks };
}

/** Restore a backup: adds new decks, replaces older copies, keeps newer ones. */
export async function importBackup(data, sanitize = (d) => d) {
  if (!data || !Array.isArray(data.decks)) throw new Error('That file isn’t a LibreSlides backup.');
  let added = 0;
  let updated = 0;
  let skipped = 0;
  for (const raw of data.decks) {
    if (!raw?.id) continue;
    const deck = normalizeDeck(sanitize(raw));
    const existing = await getMeta(deck.id);
    // A restored copy keeps the link to its file if this device already had it.
    delete deck.file;
    if (existing?.file) deck.file = { ...existing.file, dirty: true };
    if (existing && existing.updatedAt >= deck.updatedAt) {
      skipped++;
      continue;
    }
    await saveDeck(deck, { touch: false });
    if (existing) updated++;
    else added++;
  }
  announce({ type: 'restored' });
  return { added, updated, skipped };
}
