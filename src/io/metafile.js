import { encodePng } from './png.js';
import { toBase64 } from './font-data.js';

/**
 * Windows metafiles (WMF and EMF), the vector pictures older PowerPoint
 * files are full of (clip art, logos, previews of embedded objects), drawn
 * as SVG since browsers can't show them. The common records are played:
 * lines, polygons, rectangles, ellipses and arcs, Bézier curves, paths,
 * pens, brushes (solid and hatched), text, bitmaps (DIBs), gradients,
 * clipping, mapping modes, world transforms and saving and restoring the
 * drawing state. EMF+ records are skipped (a "dual" file has plain GDI
 * records as well).
 *
 * The data is untrusted: each record is read through a DataView of its own
 * length, so nothing outside it can be read, a record that's damaged is
 * skipped, and the number of records, shapes, points and bitmap pixels is
 * limited, so a hostile file can't hang the import or run it out of memory.
 */

export const METAFILE_LIMITS = { records: 500000, elements: 100000, points: 2000000, pixels: 24e6, text: 200000 };

const fail = (message) => Object.assign(new Error(`Metafile: ${message}`), { code: 'bad-picture' });

/** 'emf', 'wmf' or null, from a file's first bytes. */
export function metafileKind(bytes) {
  if (bytes.length < 18) return null;
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (v.getUint32(0, true) === 1 && bytes.length >= 88 && v.getUint32(40, true) === 0x464d4520) return 'emf';
  if (v.getUint32(0, true) === 0x9ac6cdd7) return 'wmf';
  const type = v.getUint16(0, true);
  if ((type === 1 || type === 2) && v.getUint16(2, true) === 9) return 'wmf';
  return null;
}

// ------------------------------------------------------------------ small helpers

const num = (v) => {
  const r = Math.round(v * 100) / 100;
  return Object.is(r, -0) || !Number.isFinite(r) ? '0' : String(r);
};
const hex2 = (n) => n.toString(16).padStart(2, '0');
const rgbOf = (r, g, b) => `#${hex2(r & 255)}${hex2(g & 255)}${hex2(b & 255)}`;
/** A COLORREF (red, green, blue, flags) at byte `o`. */
const colorAt = (v, o) => rgbOf(v.getUint8(o), v.getUint8(o + 1), v.getUint8(o + 2));
const esc = (s) => String(s)
  .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]|[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/g, '')
  .replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

/** Matrices as GDI writes them (row vectors): x' = a·x + c·y + e, y' = b·x + d·y + f. */
const IDENTITY = Object.freeze({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
/** A then B. */
const mul = (A, B) => ({
  a: A.a * B.a + A.b * B.c, b: A.a * B.b + A.b * B.d,
  c: A.c * B.a + A.d * B.c, d: A.c * B.b + A.d * B.d,
  e: A.e * B.a + A.f * B.c + B.e, f: A.e * B.b + A.f * B.d + B.f,
});
const finiteMatrix = (m) => Object.values(m).every(Number.isFinite);

const CHARSETS = {
  128: 'shift_jis', 129: 'euc-kr', 134: 'gbk', 136: 'big5', 161: 'windows-1253', 162: 'windows-1254', 163: 'windows-1258',
  177: 'windows-1255', 178: 'windows-1256', 186: 'windows-1257', 204: 'windows-1251', 222: 'windows-874', 238: 'windows-1250',
};
const decoders = new Map();
function decode8(bytes, charset) {
  const label = CHARSETS[charset] || 'windows-1252';
  if (!decoders.has(label)) {
    try {
      decoders.set(label, new TextDecoder(label));
    } catch {
      decoders.set(label, new TextDecoder('windows-1252'));
    }
  }
  return decoders.get(label).decode(bytes);
}

/**
 * The symbol fonts' characters as Unicode (the fonts themselves are rarely
 * there to draw them): Symbol's Greek and maths, and the bullets, boxes,
 * checks and arrows of Wingdings and Webdings.
 */
const SYMBOL_FONTS = {
  symbol: {
    ...Object.fromEntries([...'ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ'].map((c, i) => [65 + i, c])),
    ...Object.fromEntries([...'αβχδεφγηιϕκλμνοπθρστυϖωξψζ'].map((c, i) => [97 + i, c])),
    0x22: '∀', 0x24: '∃', 0x27: '∋', 0x2a: '∗', 0x2d: '−', 0x40: '≅', 0x5c: '∴', 0x5e: '⊥', 0xa3: '≤', 0xa5: '∞', 0xa7: '♣', 0xa8: '♦', 0xa9: '♥',
    0xaa: '♠', 0xab: '↔', 0xac: '←', 0xad: '↑', 0xae: '→', 0xaf: '↓', 0xb0: '°', 0xb1: '±', 0xb3: '≥', 0xb4: '×', 0xb5: '∝', 0xb6: '∂', 0xb7: '•',
    0xb8: '÷', 0xb9: '≠', 0xba: '≡', 0xbb: '≈', 0xbc: '…', 0xc5: '⊕', 0xc7: '∩', 0xc8: '∪', 0xce: '∈', 0xd6: '√', 0xd7: '⋅', 0xd8: '¬', 0xd9: '∧',
    0xda: '∨', 0xdb: '⇔', 0xdc: '⇐', 0xdd: '⇑', 0xde: '⇒', 0xdf: '⇓', 0xe5: '∑', 0xf2: '∫',
  },
  wingdings: {
    0x21: '✏', 0x22: '✂', 0x28: '☎', 0x2a: '✉', 0x36: '⌛', 0x3f: '✍', 0x46: '☞', 0x4a: '☺', 0x4c: '☹', 0x4e: '☠', 0x52: '☼', 0x54: '❄',
    0x58: '✠', 0x5b: '☯', 0x6c: '●', 0x6d: '❍', 0x6e: '■', 0x6f: '□', 0x70: '◻', 0x71: '❑', 0x72: '❒', 0x73: '⬧', 0x74: '⧫', 0x75: '◆',
    0x76: '❖', 0x77: '⬥', 0x9f: '•', 0xa1: '○', 0xa7: '▪', 0xa8: '◻', 0xab: '★', 0xd8: '➢', 0xe0: '➔', 0xe8: '➔', 0xef: '⇦', 0xf0: '⇨',
    0xfb: '✗', 0xfc: '✓', 0xfd: '☒', 0xfe: '☑',
  },
  webdings: { 0x33: '◀', 0x34: '▶', 0x35: '▲', 0x36: '▼', 0x61: '✓', 0x63: '□', 0x67: '■', 0x6e: '●', 0x72: '✕' },
};

/** Text in a symbol font as Unicode (characters may come in the private use area, U+F020 and on). */
function symbolText(face, str) {
  const table = SYMBOL_FONTS[String(face || '').toLowerCase().replace(/\s+\d*$/, '')];
  if (!table) return null;
  return [...str].map((c) => {
    let code = c.codePointAt(0);
    if (code >= 0xf000 && code <= 0xf0ff) code -= 0xf000;
    return table[code] || (code < 0x80 ? c : ' ');
  }).join('');
}

/** Font names to CSS, with a generic family to fall back on. */
function fontFamily(face) {
  const name = String(face || '').replace(/["'<>&;{}\\]/g, '').trim() || 'Arial';
  const generic = /times|serif|roman|georgia|garamond|cambria|book|palatino|minion/i.test(name) && !/sans/i.test(name) ? 'serif'
    : /courier|mono|consol|fixed|terminal|lucida console/i.test(name) ? 'monospace' : 'sans-serif';
  return `'${name}', ${generic}`;
}

// ------------------------------------------------------------------ bitmaps

/**
 * A device-independent bitmap (BITMAPINFO at `bmi`, pixels at `bits`, in
 * view `v`) as { width, height, rgba } (top row first), or { png | jpeg:
 * bytes } for a DIB that holds a whole PNG or JPEG file; null if damaged.
 */
export function decodeDib(v, bmi, bmiSize, bits, bitsSize, { alpha = false, budget } = {}) {
  // `bits` null: a packed DIB, its pixels straight after its header and colours.
  if (bmiSize < 12 || bmi < 0 || bmi + bmiSize > v.byteLength || (bits != null && (bits < 0 || bits > v.byteLength))) return null;
  const headerSize = v.getUint32(bmi, true);
  let width;
  let height;
  let bpp;
  let compression = 0;
  let colors = 0;
  let paletteAt;
  let entry = 4;
  if (headerSize === 12) {
    width = v.getUint16(bmi + 4, true);
    height = v.getInt16(bmi + 6, true);
    bpp = v.getUint16(bmi + 10, true);
    paletteAt = bmi + 12;
    entry = 3;
  } else if (headerSize >= 40 && headerSize <= bmiSize + 4096) {
    width = v.getInt32(bmi + 4, true);
    height = v.getInt32(bmi + 8, true);
    bpp = v.getUint16(bmi + 14, true);
    compression = v.getUint32(bmi + 16, true);
    colors = v.getUint32(bmi + 32, true);
    paletteAt = bmi + headerSize;
  } else return null;
  const topDown = height < 0;
  height = Math.abs(height);
  // Colour masks: BI_BITFIELDS after a 40-byte header, or inside a bigger one.
  let masks = null;
  if (compression === 3 || compression === 6) {
    const at = bmi + 40;
    if (at + 12 > v.byteLength) return null;
    masks = [v.getUint32(at, true), v.getUint32(at + 4, true), v.getUint32(at + 8, true), compression === 6 || (headerSize >= 56 && at + 16 <= v.byteLength) ? v.getUint32(at + 12, true) : 0];
    if (headerSize === 40) paletteAt += compression === 6 ? 16 : 12;
  } else if (bpp === 16) masks = [0x7c00, 0x3e0, 0x1f, 0];
  const paletteSize = bpp <= 8 ? Math.min(1 << bpp, colors || 1 << bpp) : 0;
  if (bits == null) {
    bits = paletteAt + paletteSize * entry;
    bitsSize = v.byteLength - bits;
  }
  if (bits > v.byteLength) return null;
  bitsSize = Math.max(0, Math.min(bitsSize, v.byteLength - bits));
  if (compression === 4 || compression === 5) {
    const data = new Uint8Array(v.buffer, v.byteOffset + bits, bitsSize);
    return compression === 5 ? { png: data } : { jpeg: data };
  }
  if (!(width > 0 && height > 0) || width * height > (budget?.left ?? METAFILE_LIMITS.pixels)) return null;
  if (budget) budget.left -= width * height;
  const palette = [];
  for (let i = 0; i < paletteSize; i++) {
    const o = paletteAt + i * entry;
    palette.push(o + 3 <= v.byteLength ? [v.getUint8(o + 2), v.getUint8(o + 1), v.getUint8(o)] : [0, 0, 0]);
  }
  const rgba = new Uint8Array(width * height * 4);
  const put = (x, y, r, g, b, a = 255) => {
    if (x < 0 || x >= width || y < 0 || y >= height) return;
    const o = ((topDown ? y : height - 1 - y) * width + x) * 4;
    rgba[o] = r;
    rgba[o + 1] = g;
    rgba[o + 2] = b;
    rgba[o + 3] = a;
  };
  const end = bits + bitsSize;
  const byte = (o) => (o < end ? v.getUint8(o) : 0);
  if (compression === 1 || compression === 2) {
    // Run-length encoded 8- or 4-bit pixels; unset pixels stay transparent.
    const four = compression === 2;
    const color = (i) => palette[i] || [0, 0, 0];
    let x = 0;
    let y = 0;
    for (let o = bits; o + 1 < end && y < height;) {
      const n = byte(o);
      const c = byte(o + 1);
      o += 2;
      if (n) {
        for (let k = 0; k < n; k++) put(x++, y, ...color(four ? (k & 1 ? c & 15 : c >> 4) : c));
      } else if (c === 0) {
        x = 0;
        y++;
      } else if (c === 1) break;
      else if (c === 2) {
        x += byte(o);
        y += byte(o + 1);
        o += 2;
      } else {
        for (let k = 0; k < c; k++) put(x++, y, ...color(four ? (k & 1 ? byte(o + (k >> 1)) & 15 : byte(o + (k >> 1)) >> 4) : byte(o + k)));
        const used = four ? Math.ceil(c / 2) : c;
        o += used + (used & 1);
      }
    }
    return { width, height, rgba };
  }
  if (![1, 2, 4, 8, 16, 24, 32].includes(bpp)) return null;
  const stride = (((width * bpp + 31) >>> 5) << 2);
  const shiftOf = (m) => (m ? 31 - Math.clz32(m & -m) : 0);
  const scaleOf = (m) => (m ? 255 / (m >>> shiftOf(m)) : 0);
  const ms = masks?.map(shiftOf);
  const mk = masks?.map(scaleOf);
  let anyAlpha = false;
  for (let y = 0; y < height; y++) {
    const row = bits + y * stride;
    if (row >= end) break;
    for (let x = 0; x < width; x++) {
      if (bpp <= 8) {
        const bit = x * bpp;
        const i = (byte(row + (bit >> 3)) >> (8 - bpp - (bit & 7))) & ((1 << bpp) - 1);
        put(x, y, ...(palette[i] || [0, 0, 0]));
      } else if (bpp === 24) {
        const o = row + x * 3;
        put(x, y, byte(o + 2), byte(o + 1), byte(o));
      } else {
        const o = row + x * (bpp >> 3);
        const p = bpp === 16 ? byte(o) | (byte(o + 1) << 8) : (byte(o) | (byte(o + 1) << 8) | (byte(o + 2) << 16) | (byte(o + 3) << 24)) >>> 0;
        if (masks && (bpp === 16 || compression === 3 || compression === 6)) {
          const ch = (k) => Math.round(((p & masks[k]) >>> ms[k]) * mk[k]);
          const a = masks[3] ? ch(3) : 255;
          put(x, y, ch(0), ch(1), ch(2), a);
        } else {
          const a = byte(o + 3);
          if (a) anyAlpha = true;
          put(x, y, byte(o + 2), byte(o + 1), byte(o), a);
        }
      }
    }
  }
  // 32-bit pixels: their fourth byte is alpha only when the picture is drawn with it (AlphaBlend), and not all zero.
  if (bpp === 32 && !masks) {
    if (alpha && anyAlpha) {
      for (let i = 0; i < rgba.length; i += 4) {
        const a = rgba[i + 3];
        if (a && a < 255) for (let k = 0; k < 3; k++) rgba[i + k] = Math.min(255, Math.round((rgba[i + k] * 255) / a));
      }
    } else for (let i = 3; i < rgba.length; i += 4) rgba[i] = 255;
  }
  return { width, height, rgba };
}

/** Part of a picture: the rectangle at sx, sy of sw × sh pixels (clipped to it). */
function cropPixels(img, sx, sy, sw, sh) {
  sx = Math.max(0, Math.round(sx));
  sy = Math.max(0, Math.round(sy));
  sw = Math.min(img.width - sx, Math.round(sw));
  sh = Math.min(img.height - sy, Math.round(sh));
  if (sw <= 0 || sh <= 0) return null;
  if (sx === 0 && sy === 0 && sw === img.width && sh === img.height) return img;
  const out = new Uint8Array(sw * sh * 4);
  for (let y = 0; y < sh; y++) out.set(img.rgba.subarray(((sy + y) * img.width + sx) * 4, ((sy + y) * img.width + sx + sw) * 4), y * sw * 4);
  return { width: sw, height: sh, rgba: out };
}

// ------------------------------------------------------------------ the drawing

const STOCK = {
  0: { kind: 'brush', style: 'solid', color: '#ffffff' }, 1: { kind: 'brush', style: 'solid', color: '#c0c0c0' },
  2: { kind: 'brush', style: 'solid', color: '#808080' }, 3: { kind: 'brush', style: 'solid', color: '#404040' },
  4: { kind: 'brush', style: 'solid', color: '#000000' }, 5: { kind: 'brush', style: 'null' },
  6: { kind: 'pen', style: 0, width: 0, color: '#ffffff' }, 7: { kind: 'pen', style: 0, width: 0, color: '#000000' },
  8: { kind: 'pen', style: 5, width: 0, color: '#000000' },
  10: { kind: 'font', height: -12, face: 'Courier New', weight: 400 }, 11: { kind: 'font', height: -12, face: 'Courier New', weight: 400 },
  12: { kind: 'font', height: -12, face: 'Arial', weight: 400 }, 13: { kind: 'font', height: -16, face: 'Arial', weight: 700 },
  14: { kind: 'font', height: -16, face: 'Arial', weight: 400 }, 16: { kind: 'font', height: -16, face: 'Courier New', weight: 400 },
  17: { kind: 'font', height: -11, face: 'Segoe UI', weight: 400 },
  18: { kind: 'brush', style: 'solid', color: '#ffffff' }, 19: { kind: 'pen', style: 0, width: 0, color: '#000000' },
};

const HATCHES = {
  0: 'M0,4H8', 1: 'M4,0V8', 2: 'M0,0L8,8M-1,7L1,9M7,-1L9,1', 3: 'M0,8L8,0M-1,1L1,-1M7,9L9,7', 4: 'M0,4H8M4,0V8', 5: 'M0,0L8,8M0,8L8,0',
};

/**
 * The drawing state and the SVG being made. Coordinates go through the
 * logical-to-device mapping as they come, so the SVG is in device units.
 */
class Drawing {
  constructor() {
    this.dc = {
      mapMode: 1, winOrg: [0, 0], winExt: [1, 1], vpOrg: [0, 0], vpExt: [1, 1], world: IDENTITY,
      pen: STOCK[7], brush: STOCK[0], font: STOCK[14], textColor: '#000000', bkColor: '#ffffff', bkMode: 2,
      textAlign: 0, fillRule: 'evenodd', clip: [], arcClockwise: false, miter: 10, charExtra: 0,
    };
    this.cp = [0, 0];
    this.stack = [];
    this.out = []; // { clip: [...], svg: string }
    this.defs = [];
    this.images = []; // { index, picture } to encode once played
    this.ids = 0;
    this.points = 0;
    this.path = null; // the path being built between BeginPath and EndPath (device coordinates)
    this.line = null; // MoveTo/LineTo strokes not yet written: { d, from }
    this.pxUnit = 1; // one device pixel, in the SVG's units
    this.devPerMm = [3.78, 3.78];
    this.budget = { left: METAFILE_LIMITS.pixels };
    this.full = false;
  }

  // ---------------------------------------------------------------- mapping

  /** The logical-to-device matrix. */
  matrix() {
    if (this.cached) return this.cached;
    const dc = this.dc;
    let sx = 1;
    let sy = 1;
    const metric = { 2: 0.1, 3: 0.01, 4: 0.254, 5: 0.0254, 6: 25.4 / 1440 }[dc.mapMode];
    if (metric) {
      sx = metric * this.devPerMm[0];
      sy = -metric * this.devPerMm[1];
    } else if (dc.mapMode === 7 || dc.mapMode === 8) {
      sx = dc.vpExt[0] / (dc.winExt[0] || 1);
      sy = dc.vpExt[1] / (dc.winExt[1] || 1);
      if (dc.mapMode === 7) {
        const s = Math.min(Math.abs(sx), Math.abs(sy));
        sx = Math.sign(sx || 1) * s;
        sy = Math.sign(sy || 1) * s;
      }
    }
    const page = { a: sx, b: 0, c: 0, d: sy, e: dc.vpOrg[0] - dc.winOrg[0] * sx, f: dc.vpOrg[1] - dc.winOrg[1] * sy };
    const m = mul(dc.world, page);
    this.cached = finiteMatrix(m) ? m : IDENTITY;
    return this.cached;
  }

  set(changes) {
    this.flushLine();
    this.dc = { ...this.dc, ...changes };
    this.cached = null;
  }

  pt(x, y) {
    const m = this.matrix();
    return [m.a * x + m.c * y + m.e, m.b * x + m.d * y + m.f];
  }

  /** How much the mapping scales lengths (its area scale's square root). */
  scale() {
    const m = this.matrix();
    return Math.sqrt(Math.abs(m.a * m.d - m.b * m.c)) || 1;
  }

  // ---------------------------------------------------------------- state

  save() {
    this.flushLine();
    if (this.stack.length < 1000) this.stack.push(this.dc);
  }

  restore(n) {
    this.flushLine();
    // A negative number counts back from the last saved state; a positive one is the state's place in the stack (from 1).
    const index = n < 0 ? this.stack.length + n : n - 1;
    if (index < 0 || index >= this.stack.length) return;
    this.dc = this.stack[index];
    this.stack.length = index;
    this.cached = null;
  }

  // ---------------------------------------------------------------- output

  emit(svg) {
    if (this.out.length >= METAFILE_LIMITS.elements) {
      this.full = true;
      return -1;
    }
    this.out.push({ clip: this.dc.clip, svg });
    return this.out.length - 1;
  }

  id(prefix) {
    return `${prefix}${++this.ids}`;
  }

  countPoints(n) {
    this.points += n;
    if (this.points > METAFILE_LIMITS.points) this.full = true;
    return !this.full;
  }

  strokeAttrs() {
    const pen = this.dc.pen;
    if (!pen || pen.style === 5) return 'stroke="none"';
    const scale = this.scale();
    const w = pen.cosmetic || !pen.width ? this.pxUnit : Math.max(this.pxUnit, pen.width * scale);
    const unit = Math.max(w, this.pxUnit * 2);
    const dash = pen.dashes?.length ? pen.dashes.map((d) => num(Math.max(0.1, d) * (pen.cosmetic ? this.pxUnit : scale))).join(' ')
      : { 1: [3, 1], 2: [1, 1], 3: [3, 1, 1, 1], 4: [3, 1, 1, 1, 1, 1] }[pen.style]?.map((d) => num(d * unit)).join(' ');
    const cap = { 0: 'round', 1: 'square', 2: 'butt' }[pen.cap ?? 0];
    const join = { 0: 'round', 1: 'bevel', 2: 'miter' }[pen.join ?? 0];
    return `stroke="${pen.color}" stroke-width="${num(w)}"${dash ? ` stroke-dasharray="${dash}"` : ''} stroke-linecap="${cap}" stroke-linejoin="${join}"${join === 'miter' ? ` stroke-miterlimit="${num(Math.max(1, this.dc.miter))}"` : ''}`;
  }

  fillAttrs() {
    const brush = this.dc.brush;
    if (!brush || brush.style === 'null') return 'fill="none"';
    if (brush.style === 'hatch') {
      const id = this.id('h');
      const s = num(8 * this.pxUnit);
      const bg = this.dc.bkMode === 2 ? `<rect width="8" height="8" fill="${this.dc.bkColor}"/>` : '';
      this.defs.push(`<pattern id="${id}" patternUnits="userSpaceOnUse" width="${s}" height="${s}"><g transform="scale(${num(this.pxUnit)})">${bg}<path d="${HATCHES[brush.hatch] || HATCHES[0]}" stroke="${brush.color}" stroke-width="1" fill="none"/></g></pattern>`);
      return `fill="url(#${id})" fill-rule="${this.dc.fillRule}"`;
    }
    return `fill="${brush.color}" fill-rule="${this.dc.fillRule}"`;
  }

  /**
   * A shape: path data in device units, filled with the brush and/or
   * outlined with the pen. Inside BeginPath…EndPath it's added to the path instead.
   */
  shape(d, { fill = true, stroke = true } = {}) {
    if (!d) return;
    this.flushLine();
    if (this.path) {
      this.path.push(d);
      return;
    }
    const f = fill ? this.fillAttrs() : 'fill="none"';
    const s = stroke ? this.strokeAttrs() : 'stroke="none"';
    if (f === 'fill="none"' && s === 'stroke="none"') return;
    this.emit(`<path d="${d}" ${f} ${s}/>`);
  }

  /** Path data for logical points (`close`: a closed figure). */
  poly(points, close) {
    if (points.length < 2 || !this.countPoints(points.length)) return '';
    let d = '';
    for (let i = 0; i < points.length; i++) {
      const [x, y] = this.pt(points[i][0], points[i][1]);
      d += `${i ? 'L' : 'M'}${num(x)},${num(y)}`;
    }
    return close ? `${d}Z` : d;
  }

  /** Path data for Bézier curves through logical points (a start and three per curve). */
  bezier(points, { from = null, close = false } = {}) {
    const all = from ? [from, ...points] : points;
    if (all.length < 4 || !this.countPoints(all.length)) return '';
    const p = all.map(([x, y]) => this.pt(x, y).map(num).join(','));
    let d = `M${p[0]}`;
    for (let i = 1; i + 2 < p.length; i += 3) d += `C${p[i]} ${p[i + 1]} ${p[i + 2]}`;
    return close ? `${d}Z` : d;
  }

  moveTo(x, y) {
    this.cp = [x, y];
    if (this.path) this.path.push(`M${this.pt(x, y).map(num).join(',')}`);
  }

  lineTo(x, y) {
    if (this.path) {
      const [px, py] = this.pt(x, y);
      if (!this.path.length) this.path.push(`M${this.pt(...this.cp).map(num).join(',')}`);
      this.path.push(`L${num(px)},${num(py)}`);
      this.cp = [x, y];
      return;
    }
    if (!this.countPoints(1)) return;
    // Runs of LineTo are written as one polyline.
    if (!this.line || this.line.at[0] !== this.cp[0] || this.line.at[1] !== this.cp[1]) {
      this.flushLine();
      this.line = { d: `M${this.pt(...this.cp).map(num).join(',')}`, at: this.cp };
    }
    this.line.d += `L${this.pt(x, y).map(num).join(',')}`;
    this.cp = [x, y];
    this.line.at = this.cp;
  }

  flushLine() {
    if (!this.line) return;
    const { d } = this.line;
    this.line = null;
    if (this.dc.pen?.style !== 5) this.emit(`<path d="${d}" fill="none" ${this.strokeAttrs()}/>`);
  }

  /** An ellipse's arc in logical units as Bézier path data, from angle t0 sweeping by `sweep` (radians, parametric). */
  arcPath(cx, cy, rx, ry, t0, sweep, { move = true } = {}) {
    // At most once round (a damaged file can ask for any number of turns).
    if (![cx, cy, rx, ry, t0, sweep].every(Number.isFinite)) return '';
    sweep = Math.max(-Math.PI * 2, Math.min(Math.PI * 2, sweep));
    const n = Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 2) - 1e-9));
    const dt = sweep / n;
    const k = (4 / 3) * Math.tan(dt / 4);
    const at = (t) => [cx + rx * Math.cos(t), cy + ry * Math.sin(t)];
    const fmt = (p) => this.pt(...p).map(num).join(',');
    let d = move ? `M${fmt(at(t0))}` : `L${fmt(at(t0))}`;
    for (let i = 0; i < n; i++) {
      const a = t0 + i * dt;
      const b = a + dt;
      const [x0, y0] = at(a);
      const [x3, y3] = at(b);
      d += `C${fmt([x0 - k * rx * Math.sin(a), y0 + k * ry * Math.cos(a)])} ${fmt([x3 + k * rx * Math.sin(b), y3 - k * ry * Math.cos(b)])} ${fmt([x3, y3])}`;
    }
    this.countPoints(n * 3);
    return d;
  }

  ellipse(l, t, r, b) {
    const d = this.arcPath((l + r) / 2, (t + b) / 2, Math.abs(r - l) / 2, Math.abs(b - t) / 2, 0, Math.PI * 2);
    this.shape(`${d}Z`);
  }

  rect(l, t, r, b, opts) {
    this.shape(this.poly([[l, t], [r, t], [r, b], [l, b]], true), opts);
  }

  roundRect(l, t, r, b, w, h) {
    [l, r] = [Math.min(l, r), Math.max(l, r)];
    [t, b] = [Math.min(t, b), Math.max(t, b)];
    const rx = Math.min(Math.abs(w) / 2, (r - l) / 2);
    const ry = Math.min(Math.abs(h) / 2, (b - t) / 2);
    if (!rx || !ry) return this.rect(l, t, r, b);
    const H = Math.PI / 2;
    const d = this.arcPath(r - rx, t + ry, rx, ry, -H, H)
      + this.arcPath(r - rx, b - ry, rx, ry, 0, H, { move: false })
      + this.arcPath(l + rx, b - ry, rx, ry, H, H, { move: false })
      + this.arcPath(l + rx, t + ry, rx, ry, Math.PI, H, { move: false });
    this.shape(`${d}Z`);
  }

  /** Arc, chord ('chord') or pie ('pie') of the ellipse in a box, between the radials through two points. */
  arc(l, t, r, b, xs, ys, xe, ye, kind) {
    const cx = (l + r) / 2;
    const cy = (t + b) / 2;
    const rx = Math.abs(r - l) / 2;
    const ry = Math.abs(b - t) / 2;
    if (!rx || !ry) return;
    const t0 = Math.atan2((ys - cy) / ry, (xs - cx) / rx);
    let t1 = Math.atan2((ye - cy) / ry, (xe - cx) / rx);
    // GDI draws arcs counter-clockwise as seen with y down (decreasing angles), unless told otherwise.
    const ccw = !this.dc.arcClockwise;
    let sweep = t1 - t0;
    if (ccw) {
      if (sweep >= 0) sweep -= Math.PI * 2;
    } else if (sweep <= 0) sweep += Math.PI * 2;
    if (Math.abs(sweep) < 1e-9) sweep = ccw ? -Math.PI * 2 : Math.PI * 2;
    t1 = t0 + sweep;
    let d = this.arcPath(cx, cy, rx, ry, t0, sweep);
    if (kind === 'pie') d += `L${this.pt(cx, cy).map(num).join(',')}Z`;
    else if (kind === 'chord') d += 'Z';
    this.shape(d, { fill: kind !== 'arc' });
    return [cx + rx * Math.cos(t1), cy + ry * Math.sin(t1)];
  }

  // ---------------------------------------------------------------- paths and clipping

  beginPath() {
    this.flushLine();
    this.path = [];
  }

  endPath() {
    if (this.path) this.finished = this.path.join('');
    this.path = null;
  }

  drawPath(fill, stroke) {
    const d = this.path ? this.path.join('') : this.finished;
    this.path = null;
    this.finished = null;
    if (d) this.shape(d, { fill, stroke });
  }

  clipTo(d, rule, mode = 1) {
    this.flushLine();
    if (!d) {
      if (mode === 5) this.dc = { ...this.dc, clip: [] };
      return;
    }
    const id = this.id('c');
    this.defs.push(`<clipPath id="${id}" clipPathUnits="userSpaceOnUse"><path d="${d}" clip-rule="${rule}"/></clipPath>`);
    // Intersecting: one more clip (groups nest); anything else replaces them (unions and differences are approximated).
    this.dc = { ...this.dc, clip: mode === 1 ? [...this.dc.clip, id].slice(-16) : [id] };
  }

  /** A rectangle (logical units) as the only area drawn in, within what's clipped already. */
  intersectClip(l, t, r, b) {
    this.clipTo(this.poly([[l, t], [r, t], [r, b], [l, b]], true), 'nonzero', 1);
  }

  excludeClip(l, t, r, b) {
    const big = 'M-1e6,-1e6H1e6V1e6H-1e6Z';
    this.clipTo(big + this.poly([[l, t], [r, t], [r, b], [l, b]], true), 'evenodd', 1);
  }

  // ---------------------------------------------------------------- text

  /**
   * Text at a logical point: `dx` gives each character's advance (logical
   * units) where the file has them. Drawn in the font, colour and alignment
   * of the drawing state, turned as the font and the mapping say.
   */
  text(x, y, str, dx = null, { opaqueRect = null } = {}) {
    this.flushLine();
    if (!str || this.full) return;
    if (opaqueRect) this.fillRect(...opaqueRect, this.dc.bkColor);
    const font = this.dc.font || STOCK[14];
    const align = this.dc.textAlign;
    if (align & 1) [x, y] = this.cp; // TA_UPDATECP
    let size = Math.abs(font.height || 0);
    if (font.height > 0) size *= 0.85; // a cell height, internal leading and all
    if (!size) size = 12 / this.scale();
    const symbols = symbolText(font.face, str);
    const chars = [...(symbols ?? str)].slice(0, 4000);
    // Each character's place along the baseline.
    let advances = null;
    if (dx?.length) {
      advances = [];
      let at = 0;
      let k = 0;
      for (const ch of chars) {
        advances.push(at);
        at += (dx[k] || 0) + (ch.length > 1 ? dx[k + 1] || 0 : 0) + (this.dc.charExtra || 0);
        k += ch.length;
      }
      advances.push(at);
    }
    const total = advances ? advances[advances.length - 1] : null;
    const h = align & 6; // TA_CENTER 6, TA_RIGHT 2
    let shift = 0;
    let anchor = 'start';
    if (h === 6) {
      if (total != null) shift = -total / 2;
      else anchor = 'middle';
    } else if (h === 2) {
      if (total != null) shift = -total;
      else anchor = 'end';
    }
    const v = align & 24; // TA_BOTTOM 8, TA_BASELINE 24
    const baseline = v === 24 ? 0 : v === 8 ? -0.21 * size : 0.9 * size;
    if (align & 1) this.cp = [x + (total ?? chars.length * size * 0.5), y];
    // Glyphs stay upright on the page whichever way the mapping turns y.
    const m = this.matrix();
    const flip = m.a * m.d - m.b * m.c < 0 ? -1 : 1;
    const angle = ((font.escapement || 0) / 10) * (Math.PI / 180);
    const cos = Math.cos(angle);
    const sin = Math.sin(angle);
    // Text space: x along the baseline, y down the page; then turned, flipped back and mapped.
    const turn = { a: cos, b: -sin * flip, c: sin, d: cos * flip, e: x, f: y };
    const t = mul({ a: 1, b: 0, c: 0, d: flip, e: 0, f: 0 }, mul(turn, m));
    if (!finiteMatrix(t)) return;
    const xs = advances ? advances.slice(0, chars.length).map((a) => num(a + shift)).join(' ') : num(shift);
    const style = [
      `font-family="${esc(fontFamily(symbols != null ? 'Arial' : font.face))}"`, `font-size="${num(size)}"`,
      font.weight >= 600 ? `font-weight="${font.weight >= 800 ? 900 : 700}"` : '', font.italic ? 'font-style="italic"' : '',
      font.underline || font.strike ? `text-decoration="${[font.underline && 'underline', font.strike && 'line-through'].filter(Boolean).join(' ')}"` : '',
      anchor !== 'start' ? `text-anchor="${anchor}"` : '',
    ].filter(Boolean).join(' ');
    this.chars = (this.chars || 0) + chars.length;
    if (this.chars > METAFILE_LIMITS.text) return;
    this.emit(`<text transform="matrix(${[t.a, t.b, t.c, t.d, t.e, t.f].map((n) => Math.round(n * 10000) / 10000).join(' ')})" x="${xs}" y="${num(baseline)}" ${style} fill="${this.dc.textColor}" xml:space="preserve">${esc(chars.join(''))}</text>`);
  }

  fillRect(l, t, r, b, color) {
    const d = this.poly([[l, t], [r, t], [r, b], [l, b]], true);
    if (d) this.emit(`<path d="${d}" fill="${color}"/>`);
  }

  // ---------------------------------------------------------------- pictures

  /**
   * A bitmap stretched over a logical rectangle. `rop` is the raster
   * operation: copies are drawn as they are, AND, OR and XOR (used for
   * masks) as blending modes; operations with no picture fill the rectangle.
   */
  picture(img, l, t, w, h, rop = 0xcc0020) {
    this.flushLine();
    const op = (rop >>> 16) & 0xff;
    if (!img) {
      if (op === 0xf0) this.shape(this.poly([[l, t], [l + w, t], [l + w, t + h], [l, t + h]], true), { stroke: false });
      else if (op === 0x00 || op === 0xff) this.fillRect(l, t, l + w, t + h, op ? '#ffffff' : '#000000');
      return;
    }
    if (op === 0xaa) return; // leaves the page as it is
    const blend = { 0x88: 'multiply', 0xee: 'lighten', 0x66: 'difference', 0x44: 'multiply' }[op];
    // The rectangle as an affine map of the unit square.
    const [x0, y0] = this.pt(l, t);
    const [x1, y1] = this.pt(l + w, t);
    const [x2, y2] = this.pt(l, t + h);
    const index = this.emit('');
    if (index < 0) return;
    this.images.push({ index, img, invert: op === 0x33, matrix: [x1 - x0, y1 - y0, x2 - x0, y2 - y0, x0, y0], blend });
  }

  /** The SVG, once its pictures are encoded. */
  async finish(viewBox, width, height) {
    this.flushLine();
    for (const job of this.images) {
      let href = null;
      try {
        const { img } = job;
        if (img.png) href = `data:image/png;base64,${toBase64(img.png)}`;
        else if (img.jpeg) href = `data:image/jpeg;base64,${toBase64(img.jpeg)}`;
        else {
          if (job.invert) for (let i = 0; i < img.rgba.length; i += 4) for (let k = 0; k < 3; k++) img.rgba[i + k] = 255 - img.rgba[i + k];
          href = `data:image/png;base64,${toBase64(await encodePng(img.width, img.height, img.rgba))}`;
        }
      } catch {
        href = null;
      }
      const [a, b, c, d, e, f] = job.matrix.map((n) => Math.round(n * 1000) / 1000);
      if (href) this.out[job.index].svg = `<image width="1" height="1" preserveAspectRatio="none" transform="matrix(${a} ${b} ${c} ${d} ${e} ${f})"${job.blend ? ` style="mix-blend-mode:${job.blend}"` : ''} href="${href}"/>`;
    }
    // Elements in their clip groups (consecutive ones with the same clips share them).
    let body = '';
    let open = [];
    for (const { clip, svg } of this.out) {
      if (!svg) continue;
      let same = 0;
      while (same < open.length && same < clip.length && open[same] === clip[same]) same++;
      body += '</g>'.repeat(open.length - same);
      for (let i = same; i < clip.length; i++) body += `<g clip-path="url(#${clip[i]})">`;
      open = clip;
      body += svg;
    }
    body += '</g>'.repeat(open.length);
    const [vx, vy, vw, vh] = viewBox.map(num);
    return `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="${vx} ${vy} ${vw} ${vh}" width="${num(width)}" height="${num(height)}" preserveAspectRatio="none" overflow="hidden">`
      + `${this.defs.length ? `<defs>${this.defs.join('')}</defs>` : ''}${body}</svg>`;
  }
}

// ------------------------------------------------------------------ EMF

const EMF_PLUS = 0x2b464d45;

/** The records of an EMF, each as [type, DataView of the record]. */
function* records(bytes, read, min = 8) {
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let count = 0;
  for (let at = 0; at + 8 <= bytes.length;) {
    const [type, size] = read(v, at);
    if (size < min || at + size > bytes.length || ++count > METAFILE_LIMITS.records) return;
    yield [type, new DataView(bytes.buffer, bytes.byteOffset + at, size), at];
    at += size;
  }
}

function emfToSvg(bytes) {
  const g = new Drawing();
  const head = new DataView(bytes.buffer, bytes.byteOffset, Math.min(bytes.length, 108));
  const rect = (o) => [head.getInt32(o, true), head.getInt32(o + 4, true), head.getInt32(o + 8, true), head.getInt32(o + 12, true)];
  const bounds = rect(8);
  const frame = rect(24); // in 0.01 mm
  const device = [head.getInt32(72, true), head.getInt32(76, true)];
  const mm = [head.getInt32(80, true), head.getInt32(84, true)];
  if (device[0] > 0 && device[1] > 0 && mm[0] > 0 && mm[1] > 0) g.devPerMm = [device[0] / mm[0], device[1] / mm[1]];
  const objects = new Map();
  let plus = null; // EMF+: whether it's "dual"
  let drew = false;
  const i32 = (v, o) => v.getInt32(o, true);
  const u32 = (v, o) => v.getUint32(o, true);
  const pts = (v, o, n, small) => {
    const size = small ? 4 : 8;
    n = Math.min(n, Math.floor((v.byteLength - o) / size));
    const out = new Array(Math.max(0, n));
    for (let i = 0; i < n; i++) out[i] = small ? [v.getInt16(o + i * 4, true), v.getInt16(o + i * 4 + 2, true)] : [i32(v, o + i * 8), i32(v, o + i * 8 + 4)];
    return out;
  };
  const select = (index) => {
    const obj = index & 0x80000000 ? STOCK[index & 0x7fffffff] : objects.get(index);
    if (!obj) return;
    if (obj.kind === 'pen') g.set({ pen: obj });
    else if (obj.kind === 'brush') g.set({ brush: obj });
    else if (obj.kind === 'font') g.set({ font: obj });
  };
  const logFont = (v, o) => {
    let face = '';
    for (let i = 0; i < 32 && o + 28 + i * 2 + 2 <= v.byteLength; i++) {
      const c = v.getUint16(o + 28 + i * 2, true);
      if (!c) break;
      face += String.fromCharCode(c);
    }
    return {
      kind: 'font', height: i32(v, o), escapement: i32(v, o + 8), weight: i32(v, o + 16), italic: !!v.getUint8(o + 20),
      underline: !!v.getUint8(o + 21), strike: !!v.getUint8(o + 22), charset: v.getUint8(o + 23), face,
    };
  };
  const dib = (v, bmi, bmiSize, bits, bitsSize, alpha = false) => (bmiSize ? decodeDib(v, bmi, bmiSize, bits, bitsSize, { alpha, budget: g.budget }) : null);
  const crop = (img, sx, sy, sw, sh) => (img && !img.png && !img.jpeg ? cropPixels(img, sx, sy, sw, sh) : img);
  const region = (v, o, logical) => {
    // RGNDATA: a header, then its rectangles.
    const n = Math.min(u32(v, o + 8), Math.floor((v.byteLength - o - 32) / 16), 10000);
    let d = '';
    for (let i = 0; i < n; i++) {
      const r = o + 32 + i * 16;
      const [l, t, rr, b] = [i32(v, r), i32(v, r + 4), i32(v, r + 8), i32(v, r + 12)];
      if (logical) d += g.poly([[l, t], [rr, t], [rr, b], [l, b]], true);
      else d += `M${l},${t}H${rr}V${b}H${l}Z`;
    }
    return d;
  };
  const text = (v, wide, small = false) => {
    let x;
    let y;
    let n;
    let options;
    let off;
    let dxAt = 0;
    let rc = null;
    if (small) {
      x = i32(v, 8);
      y = i32(v, 12);
      n = u32(v, 16);
      options = u32(v, 20);
      off = 36 + (options & 0x100 ? 0 : 16);
      wide = !(options & 0x200);
    } else {
      x = i32(v, 36);
      y = i32(v, 40);
      n = u32(v, 44);
      off = u32(v, 48);
      options = u32(v, 52);
      rc = [i32(v, 56), i32(v, 60), i32(v, 64), i32(v, 68)];
      dxAt = u32(v, 72);
    }
    n = Math.min(n, 8192);
    const len = n * (wide ? 2 : 1);
    if (off < 8 || off + len > v.byteLength) return;
    let str;
    const raw = new Uint8Array(v.buffer, v.byteOffset + off, len);
    if (wide) {
      str = '';
      for (let i = 0; i < n; i++) str += String.fromCharCode(v.getUint16(off + i * 2, true));
    } else str = decode8(raw, g.dc.font?.charset);
    if (options & 0x10) return; // ETO_GLYPH_INDEX: glyph numbers, not characters
    let dx = null;
    const step = options & 0x2000 ? 8 : 4; // ETO_PDY: an x and a y per character
    if (dxAt >= 8 && dxAt + n * step <= v.byteLength) {
      dx = [];
      for (let i = 0; i < n; i++) dx.push(i32(v, dxAt + i * step));
    }
    const opaque = options & 2 && rc && rc[2] > rc[0] && rc[3] > rc[1] ? rc : null;
    g.text(x, y, str, dx, { opaqueRect: opaque });
    drew = true;
  };

  for (const [type, v] of records(bytes, (dv, at) => [dv.getUint32(at, true), dv.getUint32(at + 4, true)])) {
    if (g.full) break;
    if (type === 14) break; // EOF
    try {
      switch (type) {
        case 70: { // GDICOMMENT: EMF+ records ride in these
          if (v.byteLength >= 20 && u32(v, 12) === EMF_PLUS && plus == null && v.getUint16(16, true) === 0x4001) plus = !!(v.getUint16(18, true) & 1);
          break;
        }
        case 9: g.set({ winExt: [i32(v, 8), i32(v, 12)] }); break;
        case 10: g.set({ winOrg: [i32(v, 8), i32(v, 12)] }); break;
        case 11: g.set({ vpExt: [i32(v, 8), i32(v, 12)] }); break;
        case 12: g.set({ vpOrg: [i32(v, 8), i32(v, 12)] }); break;
        case 31: case 32: { // SCALEVIEWPORTEXTEX, SCALEWINDOWEXTEX
          const [xn, xd, yn, yd] = [i32(v, 8), i32(v, 12), i32(v, 16), i32(v, 20)];
          if (!xd || !yd) break;
          const key = type === 31 ? 'vpExt' : 'winExt';
          g.set({ [key]: [(g.dc[key][0] * xn) / xd, (g.dc[key][1] * yn) / yd] });
          break;
        }
        case 17: g.set({ mapMode: u32(v, 8) }); break;
        case 18: g.set({ bkMode: u32(v, 8) }); break;
        case 19: g.set({ fillRule: u32(v, 8) === 2 ? 'nonzero' : 'evenodd' }); break;
        case 22: g.set({ textAlign: u32(v, 8) }); break;
        case 24: g.set({ textColor: colorAt(v, 8) }); break;
        case 25: g.set({ bkColor: colorAt(v, 8) }); break;
        case 57: g.set({ arcClockwise: u32(v, 8) === 2 }); break;
        case 58: g.set({ miter: v.getFloat32(8, true) }); break;
        case 27: g.moveTo(i32(v, 8), i32(v, 12)); break;
        case 54: g.lineTo(i32(v, 8), i32(v, 12)); break;
        case 33: g.save(); break;
        case 34: g.restore(i32(v, 8)); break;
        case 35: case 36: { // SETWORLDTRANSFORM, MODIFYWORLDTRANSFORM
          const x = { a: v.getFloat32(8, true), b: v.getFloat32(12, true), c: v.getFloat32(16, true), d: v.getFloat32(20, true), e: v.getFloat32(24, true), f: v.getFloat32(28, true) };
          const mode = type === 35 ? 4 : u32(v, 32);
          if (mode === 1) g.set({ world: IDENTITY });
          else if (!finiteMatrix(x)) break;
          else if (mode === 2) g.set({ world: mul(x, g.dc.world) });
          else if (mode === 3) g.set({ world: mul(g.dc.world, x) });
          else if (mode === 4) g.set({ world: x });
          break;
        }
        case 37: select(u32(v, 8)); break;
        case 40: objects.delete(u32(v, 8)); break;
        case 38: { // CREATEPEN
          const style = u32(v, 12);
          objects.set(u32(v, 8), { kind: 'pen', style: style & 15, width: i32(v, 16), color: colorAt(v, 24), cosmetic: i32(v, 16) <= 1 });
          break;
        }
        case 95: { // EXTCREATEPEN
          const style = u32(v, 28);
          const n = Math.min(u32(v, 48), 16);
          const dashes = (style & 15) === 7 ? Array.from({ length: n }, (_, i) => u32(v, 52 + i * 4)) : null;
          const geometric = !!(style & 0x10000);
          objects.set(u32(v, 8), {
            kind: 'pen', style: style & 15, width: geometric ? u32(v, 32) : 0, color: colorAt(v, 40), cosmetic: !geometric,
            cap: { 0: 0, 0x100: 1, 0x200: 2 }[style & 0xf00] ?? 0, join: { 0: 0, 0x1000: 1, 0x2000: 2 }[style & 0xf000] ?? 0, dashes,
          });
          break;
        }
        case 39: { // CREATEBRUSHINDIRECT
          const style = u32(v, 12);
          objects.set(u32(v, 8), style === 1 ? { kind: 'brush', style: 'null' } : style === 2 ? { kind: 'brush', style: 'hatch', color: colorAt(v, 16), hatch: u32(v, 20) } : { kind: 'brush', style: 'solid', color: colorAt(v, 16) });
          break;
        }
        case 93: case 94: { // CREATEMONOBRUSH, CREATEDIBPATTERNBRUSHPT: the pattern's average colour
          const img = decodeDib(v, u32(v, 16), u32(v, 20), u32(v, 24), u32(v, 28), { budget: g.budget });
          let color = '#808080';
          if (img?.rgba) {
            const sum = [0, 0, 0];
            const n = img.width * img.height;
            for (let i = 0; i < img.rgba.length; i += 4) for (let k = 0; k < 3; k++) sum[k] += img.rgba[i + k];
            color = rgbOf(...sum.map((s) => Math.round(s / n)));
            if (type === 93) color = sum[0] / n > 127 ? g.dc.bkColor : g.dc.textColor;
          }
          objects.set(u32(v, 8), { kind: 'brush', style: 'solid', color });
          break;
        }
        case 82: objects.set(u32(v, 8), logFont(v, 12)); break;
        case 2: case 85: g.shape(g.bezier(pts(v, 28, u32(v, 24), type === 85)), { fill: false }); drew = true; break;
        case 3: case 86: g.shape(g.poly(pts(v, 28, u32(v, 24), type === 86), true)); drew = true; break;
        case 4: case 87: g.shape(g.poly(pts(v, 28, u32(v, 24), type === 87), false), { fill: false }); drew = true; break;
        case 5: case 88: { // POLYBEZIERTO
          const p = pts(v, 28, u32(v, 24), type === 88);
          const d = g.bezier(p, { from: g.cp });
          if (g.path) g.path.push(g.path.length ? d.replace(/^M[^C]*/, '') : d);
          else g.shape(d, { fill: false });
          if (p.length) g.cp = p[p.length - 1];
          drew = true;
          break;
        }
        case 6: case 89: { // POLYLINETO
          for (const [x, y] of pts(v, 28, u32(v, 24), type === 89)) g.lineTo(x, y);
          drew = true;
          break;
        }
        case 7: case 8: case 90: case 91: { // POLYPOLYLINE, POLYPOLYGON
          const small = type >= 90;
          const n = Math.min(u32(v, 24), 100000, Math.floor((v.byteLength - 32) / 4));
          const counts = Array.from({ length: n }, (_, i) => u32(v, 32 + i * 4));
          const all = pts(v, 32 + n * 4, u32(v, 28), small);
          let at = 0;
          let d = '';
          for (const c of counts) {
            if (at + c > all.length) break;
            d += g.poly(all.slice(at, at + c), type === 8 || type === 91);
            at += c;
          }
          g.shape(d, { fill: type === 8 || type === 91 });
          drew = true;
          break;
        }
        case 56: case 92: { // POLYDRAW
          const n = u32(v, 24);
          const small = type === 92;
          const p = pts(v, 28, n, small);
          const typesAt = 28 + p.length * (small ? 4 : 8);
          const kindOf = (i) => (typesAt + i < v.byteLength ? v.getUint8(typesAt + i) : 2);
          let d = `M${g.pt(...g.cp).map(num).join(',')}`;
          for (let i = 0; i < p.length; i++) {
            const kind = kindOf(i);
            const [x, y] = g.pt(...p[i]);
            if ((kind & 6) === 6) d += `M${num(x)},${num(y)}`;
            else if ((kind & 6) === 4 && i + 2 < p.length) {
              d += `C${[p[i], p[i + 1], p[i + 2]].map((q) => g.pt(...q).map(num).join(',')).join(' ')}`;
              i += 2;
            } else d += `L${num(x)},${num(y)}`;
            if (kindOf(i) & 1) d += 'Z';
          }
          if (p.length) g.cp = p[p.length - 1];
          g.shape(d, { fill: false });
          drew = true;
          break;
        }
        case 42: g.ellipse(i32(v, 8), i32(v, 12), i32(v, 16), i32(v, 20)); drew = true; break;
        case 43: g.rect(i32(v, 8), i32(v, 12), i32(v, 16), i32(v, 20)); drew = true; break;
        case 44: g.roundRect(i32(v, 8), i32(v, 12), i32(v, 16), i32(v, 20), i32(v, 24), i32(v, 28)); drew = true; break;
        case 45: case 46: case 47: case 55: { // ARC, CHORD, PIE, ARCTO
          const end = g.arc(i32(v, 8), i32(v, 12), i32(v, 16), i32(v, 20), i32(v, 24), i32(v, 28), i32(v, 32), i32(v, 36), { 45: 'arc', 46: 'chord', 47: 'pie', 55: 'arc' }[type]);
          if (type === 55 && end) g.cp = end;
          drew = true;
          break;
        }
        case 41: { // ANGLEARC
          const [cx, cy, r] = [i32(v, 8), i32(v, 12), u32(v, 16)];
          const [a0, sw] = [v.getFloat32(20, true), v.getFloat32(24, true)].map((a) => (a * Math.PI) / 180);
          if (!Number.isFinite(a0) || !Number.isFinite(sw)) break;
          g.lineTo(cx + r * Math.cos(a0), cy - r * Math.sin(a0));
          g.shape(g.arcPath(cx, cy, r, r, -a0, -sw), { fill: false });
          g.cp = [cx + r * Math.cos(a0 + sw), cy - r * Math.sin(a0 + sw)];
          drew = true;
          break;
        }
        case 15: { // SETPIXELV
          const [x, y] = [i32(v, 8), i32(v, 12)];
          g.fillRect(x, y, x + 1, y + 1, colorAt(v, 16));
          break;
        }
        case 59: g.beginPath(); break;
        case 60: g.endPath(); break;
        case 61: if (g.path) g.path.push('Z'); break;
        case 68: g.path = null; g.finished = null; break;
        case 62: g.drawPath(true, false); drew = true; break;
        case 63: g.drawPath(true, true); drew = true; break;
        case 64: g.drawPath(false, true); drew = true; break;
        case 67: { // SELECTCLIPPATH
          const d = g.path ? g.path.join('') : g.finished;
          g.path = null;
          g.finished = null;
          g.clipTo(d, g.dc.fillRule, u32(v, 8));
          break;
        }
        case 30: g.intersectClip(i32(v, 8), i32(v, 12), i32(v, 16), i32(v, 20)); break;
        case 29: g.excludeClip(i32(v, 8), i32(v, 12), i32(v, 16), i32(v, 20)); break;
        case 75: { // EXTSELECTCLIPRGN (device units)
          const size = u32(v, 8);
          const mode = u32(v, 12);
          if (!size) g.clipTo(null, 'nonzero', 5);
          else g.clipTo(region(v, 16, false), 'nonzero', mode);
          break;
        }
        case 71: case 74: { // FILLRGN, PAINTRGN
          const brush = type === 71 ? objects.get(u32(v, 28)) || (u32(v, 28) & 0x80000000 ? STOCK[u32(v, 28) & 0x7fffffff] : null) : g.dc.brush;
          const d = region(v, type === 71 ? 32 : 28, true);
          const saved = g.dc;
          if (brush?.kind === 'brush') g.dc = { ...g.dc, brush };
          g.shape(d, { stroke: false });
          g.dc = saved;
          drew = true;
          break;
        }
        case 76: case 77: { // BITBLT, STRETCHBLT
          const [x, y, w, h, rop] = [i32(v, 24), i32(v, 28), i32(v, 32), i32(v, 36), u32(v, 40)];
          const [xs, ys] = [i32(v, 44), i32(v, 48)];
          const img = v.byteLength >= 100 ? dib(v, u32(v, 84), u32(v, 88), u32(v, 92), u32(v, 96)) : null;
          const [ws, hs] = type === 77 && v.byteLength >= 108 ? [i32(v, 100), i32(v, 104)] : [w, h];
          const top = img?.rgba && v.getInt32(u32(v, 84) + 8, true) > 0 ? img.height - ys - hs : ys;
          g.picture(crop(img, xs, top, ws, hs), x, y, w, h, rop);
          drew = true;
          break;
        }
        case 81: { // STRETCHDIBITS
          const [x, y, xs, ys, ws, hs] = [i32(v, 24), i32(v, 28), i32(v, 32), i32(v, 36), i32(v, 40), i32(v, 44)];
          const img = dib(v, u32(v, 48), u32(v, 52), u32(v, 56), u32(v, 60));
          // Source rectangles count from the bottom of a bottom-up bitmap.
          const top = img?.rgba && v.getInt32(u32(v, 48) + 8, true) > 0 ? img.height - ys - hs : ys;
          g.picture(crop(img, xs, top, ws, hs), x, y, i32(v, 72), i32(v, 76), u32(v, 68));
          drew = true;
          break;
        }
        case 80: { // SETDIBITSTODEVICE
          const [x, y, xs, ys, ws, hs] = [i32(v, 8 + 16), i32(v, 28), i32(v, 32), i32(v, 36), i32(v, 40), i32(v, 44)];
          const img = dib(v, u32(v, 48), u32(v, 52), u32(v, 56), u32(v, 60));
          const top = img?.rgba && v.getInt32(u32(v, 48) + 8, true) > 0 ? img.height - ys - hs : ys;
          g.picture(crop(img, xs, top, ws, hs), x, y, ws, hs);
          drew = true;
          break;
        }
        case 114: case 116: { // ALPHABLEND, TRANSPARENTBLT
          const [x, y, w, h] = [i32(v, 24), i32(v, 28), i32(v, 32), i32(v, 36)];
          const blend = u32(v, 40);
          const [xs, ys, ws, hs] = [i32(v, 44), i32(v, 48), i32(v, 100), i32(v, 104)];
          const img = dib(v, u32(v, 84), u32(v, 88), u32(v, 92), u32(v, 96), type === 114 && !!((blend >>> 24) & 1));
          const top = img?.rgba && v.getInt32(u32(v, 84) + 8, true) > 0 ? img.height - ys - hs : ys;
          const part = crop(img, xs, top, ws, hs);
          if (part?.rgba) {
            if (type === 114) {
              const k = ((blend >>> 16) & 255) / 255;
              if (k < 1) for (let i = 3; i < part.rgba.length; i += 4) part.rgba[i] = Math.round(part.rgba[i] * k);
            } else {
              const key = [blend & 255, (blend >> 8) & 255, (blend >> 16) & 255];
              for (let i = 0; i < part.rgba.length; i += 4) if (part.rgba[i] === key[0] && part.rgba[i + 1] === key[1] && part.rgba[i + 2] === key[2]) part.rgba[i + 3] = 0;
            }
          }
          g.picture(part, x, y, w, h);
          drew = true;
          break;
        }
        case 78: { // MASKBLT: the source where the mask is set
          const [x, y, w, h] = [i32(v, 24), i32(v, 28), i32(v, 32), i32(v, 36)];
          const [xs, ys] = [i32(v, 44), i32(v, 48)];
          let img = crop(dib(v, u32(v, 84), u32(v, 88), u32(v, 92), u32(v, 96)), xs, ys, w, h);
          const mask = v.byteLength >= 128 ? crop(dib(v, u32(v, 112), u32(v, 116), u32(v, 120), u32(v, 124)), i32(v, 100), i32(v, 104), w, h) : null;
          if (img?.rgba && mask?.rgba && mask.width === img.width && mask.height === img.height) {
            img = { ...img, rgba: img.rgba.slice() };
            for (let i = 0; i < img.rgba.length; i += 4) img.rgba[i + 3] = mask.rgba[i] > 127 ? 255 : 0;
          }
          g.picture(img, x, y, w, h);
          drew = true;
          break;
        }
        case 118: { // GRADIENTFILL
          const nv = Math.min(u32(v, 24), 10000);
          const n = Math.min(u32(v, 28), 10000);
          const mode = u32(v, 32);
          const vert = (i) => {
            const o = 36 + i * 16;
            return { x: i32(v, o), y: i32(v, o + 4), color: rgbOf(v.getUint16(o + 8, true) >> 8, v.getUint16(o + 10, true) >> 8, v.getUint16(o + 12, true) >> 8) };
          };
          const at = 36 + nv * 16;
          for (let i = 0; i < n; i++) {
            if (mode === 2) {
              const tri = [0, 1, 2].map((k) => u32(v, at + i * 12 + k * 4)).filter((k) => k < nv).map(vert);
              if (tri.length === 3) g.emit(`<path d="${g.poly(tri.map((p) => [p.x, p.y]), true)}" fill="${tri[0].color}"/>`);
            } else {
              const [a, b] = [u32(v, at + i * 8), u32(v, at + i * 8 + 4)];
              if (a >= nv || b >= nv) continue;
              const [p, q] = [vert(a), vert(b)];
              const id = g.id('g');
              const [x1, y1] = g.pt(p.x, p.y);
              const [x2, y2] = g.pt(mode === 0 ? q.x : p.x, mode === 0 ? p.y : q.y);
              g.defs.push(`<linearGradient id="${id}" gradientUnits="userSpaceOnUse" x1="${num(x1)}" y1="${num(y1)}" x2="${num(x2)}" y2="${num(y2)}"><stop offset="0" stop-color="${p.color}"/><stop offset="1" stop-color="${q.color}"/></linearGradient>`);
              g.emit(`<path d="${g.poly([[p.x, p.y], [q.x, p.y], [q.x, q.y], [p.x, q.y]], true)}" fill="url(#${id})"/>`);
            }
          }
          drew = true;
          break;
        }
        case 83: text(v, false); break;
        case 84: text(v, true); break;
        case 108: text(v, true, true); break;
        default: break;
      }
    } catch {
      // A damaged record: skipped.
    }
  }
  if (plus === false && !drew) throw fail('EMF+ only');
  // The picture's frame: its frame rectangle in device units, else its bounds.
  const toDevice = (r) => [(r[0] / 100) * g.devPerMm[0], (r[1] / 100) * g.devPerMm[1], (r[2] / 100) * g.devPerMm[0], (r[3] / 100) * g.devPerMm[1]];
  let box = frame[2] > frame[0] && frame[3] > frame[1] ? toDevice(frame) : null;
  if (!box && bounds[2] >= bounds[0] && bounds[3] >= bounds[1]) box = [bounds[0], bounds[1], bounds[2] + 1, bounds[3] + 1];
  if (!box) throw fail('no size');
  const [w, h] = [box[2] - box[0], box[3] - box[1]];
  const px = [((frame[2] - frame[0]) / 2540) * 96, ((frame[3] - frame[1]) / 2540) * 96];
  const size = px[0] > 0 && px[1] > 0 ? px : [w, h];
  return g.finish([box[0], box[1], w, h], Math.max(1, Math.min(4000, size[0])), Math.max(1, Math.min(4000, size[1])));
}

// ------------------------------------------------------------------ WMF

/** WMF records that only set things up (windows, colours, objects…) rather than draw. */
const WMF_STATE = new Set([
  0x020b, 0x020c, 0x020f, 0x0410, 0x020d, 0x020e, 0x0211, 0x0412, 0x0103, 0x0102, 0x0104, 0x0106, 0x0107, 0x0108, 0x012e,
  0x0209, 0x0201, 0x001e, 0x0127, 0x0214, 0x02fa, 0x02fb, 0x02fc, 0x0142, 0x01f9, 0x00f7, 0x06ff, 0x012d, 0x01f0,
  0x0234, 0x0035, 0x0626, 0x0231, 0x0105, 0x0037, 0x0139, 0x0436,
]);

async function wmfToSvg(bytes, { embedded = true } = {}) {
  const g = new Drawing();
  const v0 = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let at = 0;
  let bbox = null;
  let inch = 1440;
  if (v0.getUint32(0, true) === 0x9ac6cdd7) {
    const [l, t, r, b] = [6, 8, 10, 12].map((o) => v0.getInt16(o, true));
    if (r !== l && b !== t) bbox = [Math.min(l, r), Math.min(t, b), Math.abs(r - l), Math.abs(b - t)];
    inch = v0.getUint16(14, true) || 1440;
    at = 22;
  }
  if (at + 18 > bytes.length) throw fail('cut short');
  const headerWords = v0.getUint16(at + 2, true);
  at += Math.max(9, headerWords) * 2;
  // The player's viewport is fixed (the picture's frame); the file's window maps into it.
  g.set({ mapMode: 8 });
  let frame = null;
  /** Fix the frame at the first drawing: the window the file has set by then (else the placeable header's box). */
  const setFrame = () => {
    if (frame) return;
    const ext = g.dc.windowSet ? g.dc.winExt : bbox ? [bbox[2], bbox[3]] : null;
    if (!ext) return;
    frame = [Math.abs(ext[0]) || 1, Math.abs(ext[1]) || 1];
    g.set({ vpExt: frame });
    // A device pixel (for hairlines and hatching): the header says how many units make an inch.
    g.pxUnit = bbox ? Math.max(0.01, inch / 96) : Math.max(1, frame[0] / 1000);
  };
  // Until the file sets a window, logical units are the frame's.
  if (bbox) g.set({ winOrg: [bbox[0], bbox[1]], winExt: [bbox[2], bbox[3]], vpExt: [bbox[2], bbox[3]] });
  const objects = [];
  const add = (obj) => {
    let i = objects.indexOf(null);
    if (i < 0) i = objects.length;
    if (i < 4096) objects[i] = obj;
  };
  let drew = false;
  const emf = { parts: [], total: 0 };
  const wmfRecords =records(bytes.subarray(at), (dv, p) => [dv.getUint16(p + 4, true), dv.getUint32(p, true) * 2], 6);
  for (const [fn, v] of wmfRecords) {
    if (g.full) break;
    if (fn === 0) break;
    const s16 = (i) => v.getInt16(6 + i * 2, true);
    const u16 = (i) => v.getUint16(6 + i * 2, true);
    const n = (v.byteLength - 6) >> 1; // parameter words
    const points = (from, count) => {
      count = Math.min(count, Math.floor((n - from) / 2));
      return Array.from({ length: Math.max(0, count) }, (_, i) => [s16(from + i * 2), s16(from + i * 2 + 1)]);
    };
    const dibAt = (word) => decodeDib(v, 6 + word * 2, v.byteLength - 6 - word * 2, null, 0, { budget: g.budget });
    try {
      if (!frame && !WMF_STATE.has(fn)) setFrame();
      switch (fn) {
        case 0x0626: { // ESCAPE: an EMF of the same picture can ride in MFCOMMENT ("WMFC") chunks
          if (u16(0) !== 0x000f || v.byteLength < 50 || v.getUint32(10, true) !== 0x43464d57) break;
          const total = v.getUint32(40, true);
          const size = Math.min(v.getUint32(32, true), v.byteLength - 50 + 6);
          if (!emf.parts.length) emf.total = total;
          if (total !== emf.total || total > 64e6) break;
          emf.parts.push(new Uint8Array(v.buffer, v.byteOffset + 44, Math.max(0, Math.min(size, v.byteLength - 44))));
          break;
        }
        case 0x020b: g.set({ winOrg: [s16(1), s16(0)] }); break;
        case 0x020c: g.set({ winExt: [s16(1), s16(0)], windowSet: true }); break;
        case 0x020f: g.set({ winOrg: [g.dc.winOrg[0] + s16(1), g.dc.winOrg[1] + s16(0)] }); break;
        case 0x0410: { // SCALEWINDOWEXT
          const [yd, yn, xd, xn] = [s16(0), s16(1), s16(2), s16(3)];
          if (xd && yd) g.set({ winExt: [(g.dc.winExt[0] * xn) / xd, (g.dc.winExt[1] * yn) / yd], windowSet: true });
          break;
        }
        case 0x0103: g.set({ mapMode: u16(0) === 7 ? 7 : 8 }); break;
        case 0x0102: g.set({ bkMode: u16(0) }); break;
        case 0x0106: g.set({ fillRule: u16(0) === 2 ? 'nonzero' : 'evenodd' }); break;
        case 0x012e: g.set({ textAlign: u16(0) }); break;
        case 0x0209: g.set({ textColor: colorAt(v, 6) }); break;
        case 0x0201: g.set({ bkColor: colorAt(v, 6) }); break;
        case 0x0108: g.set({ charExtra: s16(0) }); break;
        case 0x001e: g.save(); break;
        case 0x0127: g.restore(s16(0)); break;
        case 0x0214: g.moveTo(s16(1), s16(0)); break;
        case 0x0213: g.lineTo(s16(1), s16(0)); drew = true; break;
        case 0x02fa: { // CREATEPENINDIRECT
          const style = u16(0);
          add({ kind: 'pen', style: style & 15, width: s16(1), color: colorAt(v, 12), cosmetic: s16(1) <= 1, cap: { 0: 0, 0x100: 1, 0x200: 2 }[style & 0xf00] ?? 0, join: { 0: 0, 0x1000: 1, 0x2000: 2 }[style & 0xf000] ?? 0 });
          break;
        }
        case 0x02fc: { // CREATEBRUSHINDIRECT
          const style = u16(0);
          add(style === 1 ? { kind: 'brush', style: 'null' } : style === 2 ? { kind: 'brush', style: 'hatch', color: colorAt(v, 8), hatch: u16(3) } : { kind: 'brush', style: 'solid', color: colorAt(v, 8) });
          break;
        }
        case 0x0142: case 0x01f9: { // DIBCREATEPATTERNBRUSH, CREATEPATTERNBRUSH: the pattern's average colour
          const img = fn === 0x0142 ? dibAt(2) : null;
          let color = '#808080';
          if (img?.rgba) {
            const sum = [0, 0, 0];
            for (let i = 0; i < img.rgba.length; i += 4) for (let k = 0; k < 3; k++) sum[k] += img.rgba[i + k];
            color = rgbOf(...sum.map((x) => Math.round(x / (img.width * img.height))));
          }
          add({ kind: 'brush', style: 'solid', color });
          break;
        }
        case 0x02fb: { // CREATEFONTINDIRECT
          let face = '';
          const bytesAt = 6 + 18;
          const end = Math.min(v.byteLength, bytesAt + 32);
          const raw = [];
          for (let i = bytesAt; i < end; i++) {
            const c = v.getUint8(i);
            if (!c) break;
            raw.push(c);
          }
          face = decode8(new Uint8Array(raw), 0);
          add({ kind: 'font', height: s16(0), escapement: s16(2), weight: s16(4), italic: !!v.getUint8(16), underline: !!v.getUint8(17), strike: !!v.getUint8(18), charset: v.getUint8(19), face });
          break;
        }
        case 0x00f7: case 0x06ff: add({ kind: 'other' }); break; // CREATEPALETTE, CREATEREGION
        case 0x012d: { // SELECTOBJECT
          const obj = objects[u16(0)];
          if (obj?.kind === 'pen') g.set({ pen: obj });
          else if (obj?.kind === 'brush') g.set({ brush: obj });
          else if (obj?.kind === 'font') g.set({ font: obj });
          break;
        }
        case 0x01f0: if (u16(0) < objects.length) objects[u16(0)] = null; break;
        case 0x0418: g.ellipse(s16(3), s16(2), s16(1), s16(0)); drew = true; break;
        case 0x041b: g.rect(s16(3), s16(2), s16(1), s16(0)); drew = true; break;
        case 0x061c: g.roundRect(s16(5), s16(4), s16(3), s16(2), s16(1), s16(0)); drew = true; break;
        case 0x0817: case 0x0830: case 0x081a: // ARC, CHORD, PIE
          g.arc(s16(7), s16(6), s16(5), s16(4), s16(3), s16(2), s16(1), s16(0), { 0x0817: 'arc', 0x0830: 'chord', 0x081a: 'pie' }[fn]);
          drew = true;
          break;
        case 0x0324: g.shape(g.poly(points(1, s16(0)), true)); drew = true; break;
        case 0x0325: g.shape(g.poly(points(1, s16(0)), false), { fill: false }); drew = true; break;
        case 0x0538: { // POLYPOLYGON
          const count = Math.min(u16(0), n - 1);
          let from = 1 + count;
          let d = '';
          for (let i = 0; i < count; i++) {
            const c = u16(1 + i);
            d += g.poly(points(from, c), true);
            from += c * 2;
          }
          g.shape(d);
          drew = true;
          break;
        }
        case 0x041f: g.fillRect(s16(3), s16(2), s16(3) + 1, s16(2) + 1, colorAt(v, 6)); break; // SETPIXEL
        case 0x0416: g.intersectClip(s16(3), s16(2), s16(1), s16(0)); break;
        case 0x0415: g.excludeClip(s16(3), s16(2), s16(1), s16(0)); break;
        case 0x012c: g.clipTo(null, 'nonzero', 5); break; // SELECTCLIPREGION: back to no clipping
        case 0x061d: { // PATBLT
          const rop = u16(0) | (u16(1) << 16);
          g.picture(null, s16(5), s16(4), s16(3), s16(2), rop);
          drew = true;
          break;
        }
        case 0x0521: { // TEXTOUT
          const len = Math.min(s16(0), (n - 1) * 2);
          if (len <= 0) break;
          const words = Math.ceil(len / 2);
          const str = decode8(new Uint8Array(v.buffer, v.byteOffset + 8, len), g.dc.font?.charset);
          g.text(s16(2 + words), s16(1 + words), str);
          drew = true;
          break;
        }
        case 0x0a32: { // EXTTEXTOUT
          const [y, x, len, options] = [s16(0), s16(1), Math.max(0, s16(2)), u16(3)];
          let word = 4;
          let rc = null;
          if (options & 6) {
            rc = [s16(4), s16(5), s16(6), s16(7)];
            word = 8;
          }
          const count = Math.min(len, (n - word) * 2);
          if (count <= 0) break;
          const str = decode8(new Uint8Array(v.buffer, v.byteOffset + 6 + word * 2, count), g.dc.font?.charset);
          const dxAt = word + Math.ceil(count / 2);
          const dx = n - dxAt >= count ? Array.from({ length: count }, (_, i) => s16(dxAt + i)) : null;
          g.text(x, y, str, dx, { opaqueRect: options & 2 && rc ? rc : null });
          drew = true;
          break;
        }
        case 0x0b41: case 0x0940: case 0x0f43: case 0x0d33: { // DIBSTRETCHBLT, DIBBITBLT, STRETCHDIB, SETDIBTODEV
          const words = v.byteLength / 2;
          const noBitmap = (fn === 0x0b41 || fn === 0x0940) && words === (fn >> 8) + 3;
          let rop = u16(0) | (u16(1) << 16);
          let p; // [xDest, yDest, wDest, hDest, xSrc, ySrc, wSrc, hSrc, dibWord]
          if (fn === 0x0b41) p = noBitmap ? [s16(10), s16(9), s16(8), s16(7)] : [s16(9), s16(8), s16(7), s16(6), s16(5), s16(4), s16(3), s16(2), 10];
          else if (fn === 0x0940) p = noBitmap ? [s16(8), s16(7), s16(6), s16(5), s16(3), s16(2)] : [s16(7), s16(6), s16(5), s16(4), s16(3), s16(2), s16(5), s16(4), 8];
          else if (fn === 0x0f43) p = [s16(10), s16(9), s16(8), s16(7), s16(6), s16(5), s16(4), s16(3), 11];
          else {
            rop = 0xcc0020;
            p = [s16(8), s16(7), s16(6), s16(5), s16(4), s16(3), s16(6), s16(5), 9];
          }
          let img = !noBitmap && p[8] ? dibAt(p[8]) : null;
          if (img?.rgba) {
            const bottomUp = v.getInt32(6 + p[8] * 2 + 8, true) > 0;
            img = cropPixels(img, p[4], bottomUp ? img.height - p[5] - p[7] : p[5], p[6], p[7]);
          }
          g.picture(img, p[0], p[1], p[2], p[3], rop);
          drew = true;
          break;
        }
        default: break;
      }
    } catch {
      // A damaged record: skipped.
    }
  }
  // The EMF that came along, whole: it's the better copy (Office draws it rather than the WMF).
  const emfSize = emf.parts.reduce((n, p) => n + p.length, 0);
  if (embedded && emfSize && emfSize === emf.total) {
    const whole = new Uint8Array(emfSize);
    let k = 0;
    for (const p of emf.parts) {
      whole.set(p, k);
      k += p.length;
    }
    if (metafileKind(whole) === 'emf') {
      try {
        return await emfToSvg(whole);
      } catch { /* the WMF records, then */ }
    }
  }
  setFrame();
  const ext = frame || [1, 1];
  if (!drew) throw fail('nothing drawn');
  // The frame in device units starts at 0, 0 (where the window's origin goes).
  const px = [(ext[0] / inch) * 96, (ext[1] / inch) * 96];
  return g.finish([0, 0, ext[0], ext[1]], Math.max(1, Math.min(4000, px[0])), Math.max(1, Math.min(4000, px[1])));
}

/** A WMF or EMF as SVG text; fails with an error for anything else, or a file with nothing to show. */
export async function metafileToSvg(input, options) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const kind = metafileKind(bytes);
  if (kind === 'emf') return emfToSvg(bytes);
  if (kind === 'wmf') return wmfToSvg(bytes, options);
  throw fail('not a metafile');
}
