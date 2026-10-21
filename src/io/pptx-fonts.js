import { BUNDLED_FONTS, FONT_FACES, embedFontsSetting } from '../model/fonts.js';
import { themeById } from '../model/themes.js';
import { desc, kid, attr, rid } from './ooxml.js';
import { safeFontFamily } from './sanitize.js';
import { BUNDLED_FACES } from './bundled-fonts.js';
import { woffToSfnt, wrapEot, fontInfo, decodeFontData, toBase64, dataUrlBytes } from './font-data.js';

/**
 * Fonts embedded in PowerPoint files. A presentation lists them in
 * p:embeddedFontLst (ppt/presentation.xml): per family, up to four faces
 * (regular, bold, italic, bold italic), each a font part (usually
 * ppt/fonts/fontN.fntdata, in Embedded OpenType; see font-data.js).
 *
 * Exported decks carry the bundled fonts they use, and any fonts that came
 * embedded in the file they were opened from, so they look the same in
 * PowerPoint without those fonts installed. Opened decks keep their embedded
 * fonts in deck.fonts; render/deck-fonts.js hands them to the browser.
 */
const FACES = Object.keys(FONT_FACES);
const FONT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font';
const FONT_TYPE = 'application/x-fontdata';
const MAX_FONT_BYTES = 12 * 1024 * 1024;

const escXml = (s) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const warned = new Set();
const warnOnce = (key, ...args) => {
  if (warned.has(key)) return;
  warned.add(key);
  console.warn(...args);
};

// ------------------------------------------------------------------ export

/** The first family of a CSS font-family list, without quotes. */
const firstFamily = (list) => list.split(',')[0].replace(/["']/g, '').trim();

/** Font families named in a piece of slide HTML (inline font-family and <font face>). */
function htmlFonts(html, into) {
  if (!html) return;
  const s = String(html).replace(/&quot;|&#0*34;/g, '"').replace(/&#0*39;|&apos;/g, "'");
  for (const m of s.matchAll(/font-family\s*:\s*([^;>]+)/gi)) into.add(firstFamily(m[1]));
  for (const m of s.matchAll(/\bface\s*=\s*(["'])(.*?)\1/gi)) into.add(firstFamily(m[2]));
}

/** Every font family a deck uses: the theme's, elements' own and those inside their text. */
export function usedFonts(deck) {
  const fonts = new Set();
  const theme = themeById(deck.themeId);
  fonts.add(theme.fonts.title);
  fonts.add(theme.fonts.body);
  for (const slide of deck.slides || []) {
    for (const el of slide.elements || []) {
      if (el.style?.font) fonts.add(el.style.font);
      htmlFonts(el.text, fonts);
      for (const row of el.cells || []) for (const cell of row || []) htmlFonts(cell?.text, fonts);
    }
  }
  fonts.delete('');
  return fonts;
}

const fetched = new Map(); // URL → promise of the TrueType file

async function fetchWoff(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
  return woffToSfnt(new Uint8Array(await res.arrayBuffer()));
}

/**
 * A face of a family as a TrueType file, or null if there's none to embed:
 * the deck's own embedded fonts first, then the bundled fonts (fetched once).
 */
export function loadFace(family, face, deck) {
  const own = deck.fonts?.[family];
  if (own && Object.hasOwn(own, face)) return dataUrlBytes(own[face]);
  const url = Object.hasOwn(BUNDLED_FACES, family) ? BUNDLED_FACES[family][face] : null;
  if (!url) return null;
  if (!fetched.has(url)) fetched.set(url, fetchWoff(url).catch((err) => { fetched.delete(url); throw err; }));
  return fetched.get(url);
}

/** PowerPoint's pitch-and-family byte from a font's PANOSE classification, or null if it doesn't say. */
function pitchFamily(panose) {
  if (panose[0] === 3) return 0x42; // script
  if (panose[0] === 4 || panose[0] === 5) return 0x52; // decorative
  if (panose[0] !== 2) return null;
  if (panose[3] === 9) return 0x31; // monospaced: modern, fixed pitch
  if (panose[1] >= 11) return 0x22; // sans serif (swiss), variable pitch
  return panose[1] >= 2 ? 0x12 : null; // serif (roman), variable pitch
}

/** The p:font element for a family, with what PowerPoint uses to pick a stand-in when it can't use the font. */
function fontElement(family, sfnt) {
  let extra = '';
  try {
    const { panose } = fontInfo(sfnt);
    if (panose.some(Boolean)) extra += ` panose="${[...panose].map((b) => b.toString(16).padStart(2, '0')).join('').toUpperCase()}"`;
    const pf = pitchFamily(panose);
    if (pf != null) extra += ` pitchFamily="${pf}"`;
  } catch { /* the name is enough */ }
  return `<p:font typeface="${escXml(family)}"${extra} charset="0"/>`;
}

/**
 * Put p:embeddedFontLst where the schema wants it among p:presentation's
 * children (after sldSz, notesSz and smartTags; before custShowLst,
 * photoAlbum, custDataLst, kinsoku, defaultTextStyle, modifyVerifier and
 * extLst), and say the presentation embeds fonts.
 */
export function insertFontList(presXml, list) {
  let out = presXml.replace(/<p:embeddedFontLst\/>|<p:embeddedFontLst>[\s\S]*?<\/p:embeddedFontLst>/, '');
  const next = /<p:(custShowLst|photoAlbum|custDataLst|kinsoku|defaultTextStyle|modifyVerifier|extLst)\b/.exec(out);
  const at = next ? next.index : out.lastIndexOf('</p:presentation>');
  out = out.slice(0, at) + list + out.slice(at);
  return out.replace(/<p:presentation\b(?![^>]*embedTrueTypeFonts)/, '<p:presentation embedTrueTypeFonts="1"');
}

/** Add the font parts, their relationships, content type and the font list to the package. */
async function addFontParts(zip, entries) {
  const relsPath = 'ppt/_rels/presentation.xml.rels';
  let rels = await zip.file(relsPath).async('string');
  let nextRel = Math.max(0, ...[...rels.matchAll(/Id="rId(\d+)"/g)].map((m) => +m[1])) + 1;
  let n = 1;
  const list = [];
  for (const { family, faces } of entries) {
    let refs = '';
    for (const face of FACES) {
      if (!faces[face]) continue;
      while (zip.file(`ppt/fonts/font${n}.fntdata`)) n++;
      const id = `rId${nextRel++}`;
      zip.file(`ppt/fonts/font${n}.fntdata`, faces[face].eot);
      rels = rels.replace('</Relationships>', `<Relationship Id="${id}" Type="${FONT_REL}" Target="fonts/font${n}.fntdata"/></Relationships>`);
      refs += `<p:${face} r:id="${id}"/>`;
    }
    const sample = (faces.regular || Object.values(faces)[0]).sfnt;
    list.push(`<p:embeddedFont>${fontElement(family, sample)}${refs}</p:embeddedFont>`);
  }
  zip.file(relsPath, rels);

  let types = await zip.file('[Content_Types].xml').async('string');
  if (!/Extension="fntdata"/i.test(types)) types = types.replace(/<Types\b[^>]*>/, (m) => `${m}<Default Extension="fntdata" ContentType="${FONT_TYPE}"/>`);
  zip.file('[Content_Types].xml', types);

  const pres = await zip.file('ppt/presentation.xml').async('string');
  zip.file('ppt/presentation.xml', insertFontList(pres, `<p:embeddedFontLst>${list.join('')}</p:embeddedFontLst>`));
}

/**
 * Embed the fonts a deck uses into the generated .pptx (a JSZip): the faces
 * of each family that LibreSlides has the file for. A face that can't be
 * loaded (offline before the app was cached, say) is left out; the file is
 * still written. `load(family, face, deck)` returns a TrueType file or null.
 */
export async function embedFonts(zip, deck, { enabled = embedFontsSetting(), load = loadFace } = {}) {
  if (!enabled) return;
  const entries = [];
  for (const family of usedFonts(deck)) {
    const faces = {};
    for (const face of FACES) {
      try {
        const sfnt = await load(family, face, deck);
        if (sfnt) faces[face] = { sfnt, eot: wrapEot(sfnt) };
      } catch (err) {
        warnOnce(`${family}/${face}`, `Couldn’t embed the font ${family} (${face})`, err);
      }
    }
    if (Object.keys(faces).length) entries.push({ family, faces });
  }
  if (entries.length) await addFontParts(zip, entries);
}

// ------------------------------------------------------------------ import

/**
 * Read the fonts embedded in a presentation into deck.fonts as font data
 * URLs. Fonts LibreSlides bundles anyway are skipped, and so is anything that
 * isn't a readable font. PowerPoint's own compressed (MicroType Express)
 * fonts are unpacked into TrueType; text in a font that can't be read shows
 * in a stand-in font.
 */
export async function readEmbeddedFonts({ pkg, pres, presRels, deck }) {
  for (const entry of desc(pres, 'embeddedFont')) {
    const family = safeFontFamily(attr(kid(entry, 'font'), 'typeface'));
    if (!family || BUNDLED_FONTS.includes(family)) continue;
    for (const face of FACES) {
      const rel = presRels[rid(kid(entry, face), 'id')];
      const file = rel && !rel.external ? pkg.zip.file(rel.path) : null;
      if (!file) continue;
      try {
        const bytes = await file.async('uint8array');
        if (bytes.length > MAX_FONT_BYTES) throw new Error('The font file is too large');
        const { data, mime } = decodeFontData(bytes);
        deck.fonts ||= {};
        if (!Object.hasOwn(deck.fonts, family)) deck.fonts[family] = {};
        deck.fonts[family][face] = `data:${mime};base64,${toBase64(data)}`;
      } catch (err) {
        if (err.code === 'mtx') warnOnce(`mtx/${family}`, `The embedded font “${family}” couldn’t be unpacked (MicroType Express); its text shows in a similar font.`, err);
        else console.warn(`Couldn’t read the embedded font “${family}” (${face})`, err);
      }
    }
  }
}
