import { FONT_FACES } from '../model/fonts.js';

/**
 * Fonts that came embedded in a PowerPoint file (deck.fonts) are handed to
 * the browser with the FontFace API the first time one of the deck's slides
 * is drawn, so the editor, slideshow, thumbnails and print all use them.
 * Each face is added once per document (the audience window of the presenter
 * view is a document of its own); a face the browser refuses is dropped and
 * its text keeps the fallback font.
 */
const added = new WeakMap(); // document → keys of the faces added to it

const bytesOf = (url) => {
  const s = atob(url.slice(url.indexOf(',') + 1));
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
  return out.buffer;
};

function addFace(doc, family, url, descriptors) {
  try {
    // The name as it is: browsers take it literally (quotes would become part of it).
    const face = new FontFace(family, bytesOf(url), descriptors);
    doc.fonts.add(face);
    face.load().catch((err) => {
      doc.fonts.delete(face);
      console.warn(`The embedded font “${family}” couldn’t be used`, err);
    });
  } catch (err) {
    console.warn(`The embedded font “${family}” couldn’t be used`, err);
  }
}

/** Make a deck's embedded fonts available in `doc` (cheap to call again). */
export function registerDeckFonts(deck, doc = globalThis.document) {
  const fonts = deck?.fonts;
  if (!fonts || typeof FontFace !== 'function' || !doc?.fonts) return;
  let seen = added.get(doc);
  if (!seen) added.set(doc, (seen = new Set()));
  for (const [family, faces] of Object.entries(fonts)) {
    for (const [face, url] of Object.entries(faces || {})) {
      if (!Object.hasOwn(FONT_FACES, face) || typeof url !== 'string') continue;
      // The same face of the same family with the same data is only added once.
      const key = `${family}\n${face}\n${url.length}\n${url.slice(-64)}`;
      if (seen.has(key)) continue;
      seen.add(key);
      addFace(doc, family, url, FONT_FACES[face]);
    }
  }
}
