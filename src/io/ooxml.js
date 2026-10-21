import { uid } from '../model/deck.js';

/**
 * Reading Office Open XML packages (.pptx): XML lookups that ignore namespace
 * prefixes, relationships, pictures and DrawingML colours. Shared by the
 * PowerPoint reader and its parts (charts, media, fonts).
 */
export const EMU_PER_PT = 12700;
export const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', svg: 'image/svg+xml', bmp: 'image/bmp', webp: 'image/webp' };

export const xml = (s) => new DOMParser().parseFromString(s, 'application/xml');
/** Child elements by local name (namespace prefixes vary between producers). */
export const kids = (el, name) => (el ? [...el.children].filter((c) => c.localName === name) : []);
export const kid = (el, name) => kids(el, name)[0] || null;
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

// ------------------------------------------------------------------ reading the package

export class Package {
  constructor(zip) {
    this.zip = zip;
    this.cache = new Map();
  }

  async text(p) {
    const f = this.zip.file(p);
    return f ? f.async('string') : null;
  }

  async doc(p) {
    if (!this.cache.has(p)) this.cache.set(p, this.text(p).then((t) => (t ? xml(t) : null)));
    return this.cache.get(p);
  }

  /** rId → { target path, type } for a part. */
  async rels(p) {
    const d = await this.doc(relsPath(p));
    const map = {};
    for (const r of desc(d, 'Relationship')) {
      const target = attr(r, 'Target');
      map[attr(r, 'Id')] = { path: attr(r, 'TargetMode') === 'External' ? target : resolvePath(p, target), type: (attr(r, 'Type') || '').split('/').pop(), external: attr(r, 'TargetMode') === 'External' };
    }
    return map;
  }

  async dataUrl(p) {
    const f = this.zip.file(p);
    if (!f) return null;
    const ext = (p.split('.').pop() || '').toLowerCase();
    const mime = MIME[ext];
    if (!mime) return null; // EMF/WMF and other formats browsers can't draw
    return `data:${mime};base64,${await f.async('base64')}`;
  }
}

/** Store a picture (or other media data URL) in the deck once, returning its id. */
export function addMediaTo(deck, data) {
  for (const [id, d] of Object.entries(deck.media)) if (d === data) return id;
  const id = `m${uid()}`;
  deck.media[id] = data;
  return id;
}
