/**
 * Font files as bytes, for fonts embedded in PowerPoint files.
 *
 * PowerPoint keeps each embedded font (ppt/fonts/*.fntdata) as Embedded
 * OpenType (EOT): a little-endian header – sizes, version, flags, a few OS/2
 * and head values, the font's names – followed by the TrueType file. The
 * header flags can say the TrueType data is XOR-obfuscated (each byte ^ 0x50)
 * or MicroType Express (MTX) compressed. PowerPoint itself writes compressed
 * fonts, which LibreSlides can't unpack; LibreOffice writes them uncompressed,
 * as LibreSlides does (see vcl/source/font/EOTConverter.cxx in LibreOffice and
 * the W3C EOT submission).
 *
 * The app's own fonts are bundled as WOFF, which is TrueType with each table
 * zlib-compressed, so they're rebuilt into a TrueType file before wrapping.
 */
const view = (b) => new DataView(b.buffer, b.byteOffset, b.byteLength);
const tagAt = (b, o) => String.fromCharCode(b[o], b[o + 1], b[o + 2], b[o + 3]);
const fail = (message) => { throw new Error(message); };

// TrueType ('\0\1\0\0' or 'true') and OpenType with CFF outlines ('OTTO').
const TRUETYPE = new Set([0x00010000, 0x74727565]);
const OPENTYPE_CFF = 0x4f54544f;

const EOT_VERSIONS = new Set([0x00010000, 0x00020001, 0x00020002]);
const EOT_MAGIC = 0x504c;
export const EOT_COMPRESSED = 0x00000004; // MicroType Express
export const EOT_XOR = 0x10000000;

// ------------------------------------------------------------------ TrueType tables

/** Is this a TrueType or OpenType file? */
export const isSfnt = (bytes) => bytes.length >= 12 && (TRUETYPE.has(view(bytes).getUint32(0)) || view(bytes).getUint32(0) === OPENTYPE_CFF);

/** The tables of a TrueType/OpenType file, tag → bytes. Throws if it isn't one or is damaged. */
export function sfntTables(bytes) {
  if (!isSfnt(bytes)) fail('Not a TrueType or OpenType font');
  const v = view(bytes);
  const count = v.getUint16(4);
  if (!count || 12 + count * 16 > bytes.length) fail('Damaged font table directory');
  const tables = new Map();
  for (let i = 0; i < count; i++) {
    const rec = 12 + i * 16;
    const offset = v.getUint32(rec + 8);
    const length = v.getUint32(rec + 12);
    if (offset + length > bytes.length) fail('Damaged font table directory');
    tables.set(tagAt(bytes, rec), bytes.subarray(offset, offset + length));
  }
  return tables;
}

/** The TrueType checksum: the sum of the data as big-endian 32-bit numbers (zero-padded). */
export function checksum(bytes) {
  let sum = 0;
  for (let i = 0; i < bytes.length; i += 4) {
    const word = ((bytes[i] << 24) | ((bytes[i + 1] || 0) << 16) | ((bytes[i + 2] || 0) << 8) | (bytes[i + 3] || 0)) >>> 0;
    sum = (sum + word) >>> 0;
  }
  return sum;
}

/**
 * A TrueType file from its tables (tag → bytes): the table directory in tag
 * order, each table on a 4-byte boundary, the checksums worked out afresh and
 * the head table's checkSumAdjustment set so the whole file sums to 0xB1B0AFBA.
 */
export function buildSfnt(flavor, tables) {
  const tags = [...tables.keys()].sort();
  const count = tags.length;
  let entrySelector = 0;
  while (2 ** (entrySelector + 1) <= count) entrySelector++;
  const searchRange = 2 ** entrySelector * 16;
  const offsets = [];
  let size = 12 + count * 16;
  for (const tag of tags) {
    offsets.push(size);
    size += (tables.get(tag).length + 3) & ~3;
  }
  const out = new Uint8Array(size);
  const v = view(out);
  v.setUint32(0, flavor);
  v.setUint16(4, count);
  v.setUint16(6, searchRange);
  v.setUint16(8, entrySelector);
  v.setUint16(10, count * 16 - searchRange);
  tags.forEach((tag, i) => {
    const data = tables.get(tag);
    out.set(data, offsets[i]);
    // The head table's checksum is taken with checkSumAdjustment zeroed.
    if (tag === 'head' && data.length >= 12) v.setUint32(offsets[i] + 8, 0);
    const rec = 12 + i * 16;
    for (let k = 0; k < 4; k++) out[rec + k] = tag.charCodeAt(k);
    v.setUint32(rec + 4, checksum(out.subarray(offsets[i], offsets[i] + data.length)));
    v.setUint32(rec + 8, offsets[i]);
    v.setUint32(rec + 12, data.length);
  });
  const head = tags.indexOf('head');
  if (head >= 0 && tables.get('head').length >= 12) v.setUint32(offsets[head] + 8, (0xb1b0afba - checksum(out)) >>> 0);
  return out;
}

/** Inflate zlib data (WOFF tables) with the browser's DecompressionStream. */
export async function inflateZlib(data) {
  const stream = new Response(data).body.pipeThrough(new DecompressionStream('deflate'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/** A WOFF 1.0 font as the TrueType/OpenType file it was made from. */
export async function woffToSfnt(bytes, inflate = inflateZlib) {
  const v = view(bytes);
  if (bytes.length < 44 || v.getUint32(0) !== 0x774f4646) fail('Not a WOFF font'); // 'wOFF'
  const count = v.getUint16(12);
  if (!count || 44 + count * 20 > bytes.length) fail('Damaged WOFF table directory');
  const tables = new Map();
  for (let i = 0; i < count; i++) {
    const entry = 44 + i * 20;
    const offset = v.getUint32(entry + 4);
    const compLength = v.getUint32(entry + 8);
    const origLength = v.getUint32(entry + 12);
    if (offset + compLength > bytes.length || compLength > origLength) fail('Damaged WOFF table directory');
    const raw = bytes.subarray(offset, offset + compLength);
    // A table is stored compressed only when that made it smaller.
    const data = compLength < origLength ? await inflate(raw) : raw;
    if (data.length !== origLength) fail('Damaged WOFF table');
    tables.set(tagAt(bytes, entry), data);
  }
  return buildSfnt(v.getUint32(4), tables);
}

// ------------------------------------------------------------------ names and metrics

const NAME_IDS = { family: 1, style: 2, fullName: 4, version: 5 };

const utf16be = (b) => {
  let s = '';
  for (let i = 0; i + 1 < b.length; i += 2) s += String.fromCharCode((b[i] << 8) | b[i + 1]);
  return s;
};

/** The family, style, full and version names of a font (English Windows names preferred). */
export function fontNames(tables) {
  const t = tables.get('name');
  const best = {};
  if (t && t.length >= 6) {
    const v = view(t);
    const count = v.getUint16(2);
    const strings = v.getUint16(4);
    for (let i = 0; i < count && 6 + i * 12 + 12 <= t.length; i++) {
      const r = 6 + i * 12;
      const [platform, encoding, language, id, length, offset] = [0, 2, 4, 6, 8, 10].map((k) => v.getUint16(r + k));
      const start = strings + offset;
      if (start + length > t.length) continue;
      const bytes = t.subarray(start, start + length);
      let score;
      let text;
      if (platform === 3 && (encoding === 0 || encoding === 1 || encoding === 10)) [score, text] = [language === 0x409 ? 3 : 2, utf16be(bytes)];
      else if (platform === 0) [score, text] = [1, utf16be(bytes)];
      else if (platform === 1 && encoding === 0) [score, text] = [0.5, String.fromCharCode(...bytes)];
      else continue;
      if (!best[id] || best[id].score < score) best[id] = { score, text };
    }
  }
  return Object.fromEntries(Object.entries(NAME_IDS).map(([key, id]) => [key, best[id]?.text || '']));
}

/** What a font says about itself in its OS/2 and head tables (zeros where they're missing). */
export function fontInfo(sfnt) {
  const tables = sfntTables(sfnt);
  const os2 = tables.get('OS/2');
  const head = tables.get('head');
  const info = {
    ...fontNames(tables),
    trueType: TRUETYPE.has(view(sfnt).getUint32(0)),
    weight: 400, italic: false, fsType: 0, panose: new Uint8Array(10),
    unicodeRange: [0, 0, 0, 0], codePageRange: [0, 0], checkSumAdjustment: 0,
  };
  if (os2 && os2.length >= 78) {
    const v = view(os2);
    info.weight = v.getUint16(4);
    info.fsType = v.getUint16(8);
    info.panose = os2.slice(32, 42);
    info.unicodeRange = [42, 46, 50, 54].map((o) => v.getUint32(o));
    info.italic = !!(v.getUint16(62) & 1);
    // Code page ranges came with OS/2 version 1.
    if (v.getUint16(0) >= 1 && os2.length >= 86) info.codePageRange = [78, 82].map((o) => v.getUint32(o));
  }
  if (head && head.length >= 12) info.checkSumAdjustment = view(head).getUint32(8);
  return info;
}

// ------------------------------------------------------------------ Embedded OpenType

/** A name as EOT stores it: UTF-16LE with a terminating zero counted in its size (as PowerPoint writes it). */
function eotString(s) {
  if (!s) return new Uint8Array(0);
  const out = new Uint8Array(s.length * 2 + 2);
  for (let i = 0; i < s.length; i++) {
    out[i * 2] = s.charCodeAt(i) & 0xff;
    out[i * 2 + 1] = s.charCodeAt(i) >> 8;
  }
  return out;
}

/**
 * Wrap a TrueType font in an uncompressed, unobfuscated EOT (version
 * 0x00020002), field for field as LibreOffice's EOTConverter does.
 */
export function wrapEot(sfnt) {
  const info = fontInfo(sfnt);
  if (!info.trueType) fail('Only TrueType fonts can be embedded in PowerPoint files');
  const names = [info.family, info.style, info.version, info.fullName].map(eotString);
  // Fixed header (80 bytes), four names each after 2 bytes of padding and a 2-byte size, then the
  // version 2.2 tail: padding, root string size, root string checksum, EUDC code page, padding,
  // signature size, EUDC flags and EUDC font size.
  const headerSize = 80 + names.reduce((n, s) => n + 4 + s.length, 0) + 2 + 2 + 4 + 4 + 2 + 2 + 4 + 4;
  const out = new Uint8Array(headerSize + sfnt.length);
  const v = view(out);
  v.setUint32(0, out.length, true);
  v.setUint32(4, sfnt.length, true);
  v.setUint32(8, 0x00020002, true);
  v.setUint32(12, 0, true); // flags: not compressed, not obfuscated
  out.set(info.panose, 16);
  out[26] = 0; // charset: ANSI, as PowerPoint and LibreOffice write it
  out[27] = info.italic ? 1 : 0;
  v.setUint32(28, info.weight, true);
  v.setUint16(32, info.fsType, true);
  v.setUint16(34, EOT_MAGIC, true);
  info.unicodeRange.forEach((r, i) => v.setUint32(36 + i * 4, r, true));
  info.codePageRange.forEach((r, i) => v.setUint32(52 + i * 4, r, true));
  v.setUint32(60, info.checkSumAdjustment, true);
  let p = 80; // after four reserved zeros
  for (const s of names) {
    v.setUint16(p + 2, s.length, true);
    out.set(s, p + 4);
    p += 4 + s.length;
  }
  p += 4; // padding, no root string
  v.setUint32(p, 0x50475342, true); // root string checksum of nothing
  v.setUint32(p + 4, 1252, true); // EUDC code page, as LibreOffice writes it
  // The padding, signature size, EUDC flags and EUDC font size that follow are all zero.
  out.set(sfnt, headerSize);
  return out;
}

/**
 * The font inside an EOT file: { data, flags, family }. XOR obfuscation is
 * undone; MicroType Express compressed data throws an error with code 'mtx'.
 */
export function unwrapEot(bytes) {
  if (bytes.length < 82) fail('Not an Embedded OpenType font');
  const v = view(bytes);
  const dataSize = v.getUint32(4, true);
  const version = v.getUint32(8, true);
  const flags = v.getUint32(12, true);
  if (!EOT_VERSIONS.has(version) || v.getUint16(34, true) !== EOT_MAGIC) fail('Not an Embedded OpenType font');
  let p = 80;
  const need = (n) => { if (p + n > bytes.length) fail('Damaged Embedded OpenType font'); };
  // A name: 2 bytes of padding, a 2-byte size and UTF-16LE text.
  const readString = () => {
    need(4);
    const size = v.getUint16(p + 2, true);
    p += 4;
    need(size);
    let s = '';
    for (let i = 0; i + 1 < size; i += 2) s += String.fromCharCode(v.getUint16(p + i, true));
    p += size;
    return s.replace(/\0+$/, '');
  };
  const family = readString();
  readString(); // style
  readString(); // version
  readString(); // full name
  if (version >= 0x00020001) readString(); // root string (sites allowed to use a web font)
  if (version === 0x00020002) {
    p += 10; // root string checksum, EUDC code page, padding
    need(2);
    p += 2 + v.getUint16(p, true); // signature
    p += 4; // EUDC flags
    need(4);
    p += 4 + v.getUint32(p, true); // EUDC font
  }
  need(dataSize);
  const data = bytes.slice(p, p + dataSize);
  if (flags & EOT_XOR) for (let i = 0; i < data.length; i++) data[i] ^= 0x50;
  if (flags & EOT_COMPRESSED) throw Object.assign(new Error('The font is MicroType Express compressed'), { code: 'mtx', family });
  return { data, flags, family };
}

/**
 * An embedded font part (.fntdata) as a font file browsers load: EOT is
 * unwrapped, a plain TrueType/OpenType file is taken as it is. Returns
 * { data, mime }; throws for anything else or a damaged font.
 */
export function decodeFontData(bytes) {
  const data = isSfnt(bytes) ? bytes : unwrapEot(bytes).data;
  sfntTables(data); // throws if it isn't a usable font
  return { data, mime: view(data).getUint32(0) === OPENTYPE_CFF ? 'font/otf' : 'font/ttf' };
}

// ------------------------------------------------------------------ base64

export function toBase64(bytes) {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

export function fromBase64(text) {
  const s = atob(text);
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
  return out;
}

/** The bytes of a base64 data URL. */
export const dataUrlBytes = (url) => fromBase64(url.slice(url.indexOf(',') + 1));
