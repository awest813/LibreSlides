import { uid } from '../model/deck.js';
import { toBase64 } from './font-data.js';

/**
 * Office Open XML packages (.pptx): reading them (XML lookups that ignore
 * namespace prefixes, relationships, pictures and DrawingML colours, with
 * limits on how much a part may unpack to) and escaping text for the XML
 * that's written. Shared by the PowerPoint reader, writer and their parts.
 */
export const EMU_PER_PT = 12700;
/** Points as whole EMU, as the XML wants them. */
export const emu = (pt) => Math.round(pt * EMU_PER_PT);
export const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', svg: 'image/svg+xml', bmp: 'image/bmp', webp: 'image/webp' };

export const xml = (s) => new DOMParser().parseFromString(s, 'application/xml');
/** Child elements by local name (namespace prefixes vary between producers). */
export const kids = (el, name) => (el ? [...el.children].filter((c) => c.localName === name) : []);
export const kid = (el, name) => {
  for (let c = el?.firstElementChild; c; c = c.nextElementSibling) if (c.localName === name) return c;
  return null;
};
export const path = (el, ...names) => names.reduce((e, n) => kid(e, n), el);
export const desc = (el, name) => (el ? [...el.getElementsByTagNameNS('*', name)] : []);
export const attr = (el, name) => el?.getAttribute(name) ?? null;
export const rid = (el, name) => el?.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', name) || el?.getAttribute(`r:${name}`) || null;

export function resolvePath(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

export const relsPath = (p) => p.replace(/([^/]+)$/, '_rels/$1.rels');

// ------------------------------------------------------------------ colours

function hexToHsl(hex) {
  const n = parseInt(hex, 16);
  const r = ((n >> 16) & 255) / 255;
  const g = ((n >> 8) & 255) / 255;
  const b = (n & 255) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const hh = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [hh / 6, s, l];
}

function hslToHex([hh, s, l]) {
  const f = (n) => {
    const k = (n + hh * 12) % 12;
    const a = s * Math.min(l, 1 - l);
    return Math.round(255 * (l - a * Math.max(-1, Math.min(k - 3, 9 - k, 1))));
  };
  return [f(0), f(8), f(4)].map((v) => v.toString(16).padStart(2, '0')).join('');
}

// Tint and shade work on linear (not gamma-encoded) RGB, which is why
// PowerPoint's "tx1 tint 75%" subtitle grey is #898989 rather than #404040.
const toLinear = (c) => (c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
const fromLinear = (c) => (c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055);
function mapLinear(hex, fn) {
  const n = parseInt(hex, 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255]
    .map((v) => Math.round(255 * Math.max(0, Math.min(1, fromLinear(fn(toLinear(v / 255)))))).toString(16).padStart(2, '0'))
    .join('');
}

// Colour changes that work on hue (0), saturation (1) or luminance (2).
const HSL_MODS = {
  hueMod: [0, (x, v) => x * v], hueOff: [0, (x, v) => x + v],
  satMod: [1, (x, v) => x * v], satOff: [1, (x, v) => x + v],
  lumMod: [2, (x, v) => x * v], lumOff: [2, (x, v) => x + v],
};

export const SCHEME_ALIAS = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' };

/** The colour of an element holding a colour choice (srgbClr, schemeClr…), as "#rrggbb". */
export function colorOf(holder, scheme) {
  if (!holder) return null;
  const c = [...holder.children].find((e) => /Clr$/.test(e.localName));
  if (!c) return null;
  let hex = null;
  if (c.localName === 'srgbClr') hex = attr(c, 'val');
  else if (c.localName === 'schemeClr') {
    const v = attr(c, 'val');
    hex = scheme[v] || scheme[SCHEME_ALIAS[v]] || null;
  } else if (c.localName === 'sysClr') hex = attr(c, 'lastClr') || (attr(c, 'val') === 'window' ? 'FFFFFF' : '000000');
  else if (c.localName === 'prstClr') hex = { black: '000000', white: 'FFFFFF', red: 'FF0000', blue: '0000FF', green: '00FF00' }[attr(c, 'val')] || '000000';
  if (!hex || !/^[0-9a-f]{6}$/i.test(hex)) return null;
  // Modifiers apply in document order.
  for (const m of c.children) {
    const v = +attr(m, 'val') / 100000;
    if (!Number.isFinite(v)) continue;
    if (m.localName === 'tint') hex = mapLinear(hex, (x) => x * v + (1 - v));
    else if (m.localName === 'shade') hex = mapLinear(hex, (x) => x * v);
    else if (HSL_MODS[m.localName]) {
      const [i, change] = HSL_MODS[m.localName];
      const hsl = hexToHsl(hex);
      // Hue offsets are in 60000ths of a degree (val / 100000 * 100000 / 60000 / 360 of a turn).
      hsl[i] = i === 0 ? ((change(hsl[0], m.localName === 'hueOff' ? (v * 100000) / 60000 / 360 : v) % 1) + 1) % 1 : Math.max(0, Math.min(1, change(hsl[i], v)));
      hex = hslToHex(hsl);
    }
  }
  return `#${hex.toLowerCase()}`;
}

export const alphaOf = (holder) => {
  const c = holder && [...holder.children].find((e) => /Clr$/.test(e.localName));
  const a = kid(c, 'alpha');
  return a ? Math.max(0, Math.min(1, +attr(a, 'val') / 100000)) : 1;
};

/** `color` seen at `alpha` over `under` (both "#rrggbb"). */
export function blend(color, under, alpha) {
  const p = (h) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16));
  const [a, b] = [p(color), p(under)];
  return `#${a.map((v, i) => Math.round(v * alpha + b[i] * (1 - alpha)).toString(16).padStart(2, '0')).join('')}`;
}

/** Text for XML: escaped, without the control characters (or halves of emoji) XML can't hold. */
export const xmlText = (s) => String(s ?? '')
  .replace(/\r\n?/g, '\n')
  .replace(/[\ud800-\udbff][\udc00-\udfff]|[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff\ud800-\udfff]/g, (c) => (c.length === 2 ? c : ''))
  .replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c]));

/** A GUID as Office writes them: {8-4-4-4-12}, upper-case hex. */
export function guid() {
  const random = () => 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    return (c === 'x' ? r : (r & 3) | 8).toString(16);
  });
  return `{${(globalThis.crypto?.randomUUID?.() || random()).toUpperCase()}}`;
}

// ------------------------------------------------------------------ reading the package

/**
 * The most a package's parts may unpack to: an XML part, any part (videos
 * can be big) and all of them together. More is a damaged file or a zip
 * bomb, and reading stops as soon as the data passes the limit, whatever
 * sizes the zip claims.
 */
export const PART_LIMITS = { xml: 64 * 1024 * 1024, part: 256 * 1024 * 1024, total: 1024 * 1024 * 1024 };

const tooBig = () => Object.assign(new Error('This file is too large to open, or it’s damaged.'), { code: 'too-big' });

/** A zip entry's bytes, stopping with a "too-big" error past `limit`. */
function readEntry(file, limit) {
  // The size the zip claims, which can lie (checked again as the data comes in).
  if (file._data?.uncompressedSize > limit) return Promise.reject(tooBig());
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    let done = false;
    const stream = file.internalStream('uint8array');
    stream.on('data', (chunk) => {
      if (done) return;
      size += chunk.length;
      if (size > limit) {
        done = true;
        stream.pause();
        reject(tooBig());
      } else chunks.push(chunk);
    }).on('error', (err) => {
      if (!done) reject(err);
      done = true;
    }).on('end', () => {
      if (done) return;
      done = true;
      const out = new Uint8Array(size);
      let at = 0;
      for (const c of chunks) {
        out.set(c, at);
        at += c.length;
      }
      resolve(out);
    }).resume();
  });
}

/** Text of an XML part: UTF-8, or UTF-16 with a byte order mark. */
function decodeText(bytes) {
  const [a, b] = bytes;
  const encoding = a === 0xff && b === 0xfe ? 'utf-16le' : a === 0xfe && b === 0xff ? 'utf-16be' : 'utf-8';
  return new TextDecoder(encoding).decode(bytes);
}

export class Package {
  constructor(zip, limits = PART_LIMITS) {
    this.zip = zip;
    this.limits = limits;
    this.used = 0; // bytes unpacked so far
    this.cache = new Map();
    this.urls = new Map();
  }

  /**
   * A part's bytes, or null when there's no such part. Throws a "too-big"
   * error past the limits; past the total, the package is `full` and
   * nothing more should be read from it.
   */
  async bytes(p, limit = this.limits.part) {
    const f = this.zip.file(p);
    if (!f) return null;
    const room = this.limits.total - this.used;
    try {
      const bytes = await readEntry(f, Math.min(limit, room));
      this.used += bytes.length;
      return bytes;
    } catch (err) {
      if (err.code === 'too-big' && room <= limit) this.full = true;
      throw err;
    }
  }

  async text(p) {
    const bytes = await this.bytes(p, this.limits.xml);
    return bytes && decodeText(bytes);
  }

  async doc(p) {
    if (!this.cache.has(p)) this.cache.set(p, this.text(p).then((t) => (t ? xml(t) : null)));
    return this.cache.get(p);
  }

  /** rId → { path, type, external } for a part (relationships without an id or target are left out). */
  async rels(p) {
    const d = await this.doc(relsPath(p));
    const map = Object.create(null);
    for (const r of desc(d, 'Relationship')) {
      const id = attr(r, 'Id');
      const target = attr(r, 'Target');
      if (!id || target == null) continue;
      // An external target is only ever kept as a link, never read or fetched.
      const external = attr(r, 'TargetMode') === 'External';
      map[id] = { path: external ? target : resolvePath(p, target), type: (attr(r, 'Type') || '').split('/').pop(), external };
    }
    return map;
  }

  /**
   * A part as a data URL, of type `mime` or else the one its extension says;
   * null when there's no such part or browsers can't draw it (EMF, WMF…).
   * Read once however often it's asked for (a master's logo, on every slide).
   */
  dataUrl(p, mime = MIME[(String(p).split('.').pop() || '').toLowerCase()]) {
    if (!mime || !this.zip.file(p)) return Promise.resolve(null);
    const key = `${mime} ${p}`;
    if (!this.urls.has(key)) {
      // A picture or video too big to keep is left out, as long as the package as a whole isn't.
      this.urls.set(key, this.bytes(p).then((b) => `data:${mime};base64,${toBase64(b)}`, (err) => {
        if (err.code === 'too-big' && !this.full) return null;
        throw err;
      }));
    }
    return this.urls.get(key);
  }
}

// The deck's pictures by their data, so each is stored once (see addMediaTo).
const mediaIndex = new WeakMap();

/** Store a picture (or other media data URL) in the deck once, returning its id. */
export function addMediaTo(deck, data) {
  let index = mediaIndex.get(deck.media);
  if (!index) {
    index = new Map(Object.entries(deck.media).map(([id, d]) => [d, id]));
    mediaIndex.set(deck.media, index);
  }
  const known = index.get(data);
  if (known && deck.media[known] === data) return known;
  const id = `m${uid()}`;
  deck.media[id] = data;
  index.set(data, id);
  return id;
}
