import { readThrough } from './png.js';

/**
 * A TIFF decoder for pictures in PowerPoint files, which browsers can't
 * show: baseline TIFF (uncompressed, PackBits), LZW and Deflate (with or
 * without the horizontal predictor), in strips or tiles, chunky or planar,
 * as bilevel, grey, palette, RGB (with alpha) or CMYK pixels of 1 to 16
 * bits. Only the first image of the file is read. The data is untrusted:
 * every offset and length is checked, and a damaged file fails with an
 * error (or, inside its pixel data, leaves the rest of the picture blank)
 * rather than hanging or reading out of bounds.
 */

export const TIFF_MAX_PIXELS = 25e6;
/** The most bytes of samples a picture (or one strip or tile of it) may unpack to. */
const TIFF_MAX_BYTES = 200 * 1024 * 1024;

const fail = (message) => Object.assign(new Error(`TIFF: ${message}`), { code: 'bad-picture' });

/** The size in bytes of a value of each field type. */
const TYPE_SIZE = { 1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 6: 1, 7: 1, 8: 2, 9: 4, 10: 8, 11: 4, 12: 8 };

/** The fields of a TIFF's first image: tag → array of numbers (null if the file isn't TIFF). */
function readIfd(bytes) {
  if (bytes.length < 8) throw fail('too short');
  const little = bytes[0] === 0x49 && bytes[1] === 0x49;
  if (!little && !(bytes[0] === 0x4d && bytes[1] === 0x4d)) throw fail('not a TIFF');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u16 = (o) => view.getUint16(o, little);
  const u32 = (o) => view.getUint32(o, little);
  if (u16(2) !== 42) throw fail('not a classic TIFF');
  const at = u32(4);
  if (at < 8 || at + 2 > bytes.length) throw fail('no image');
  const n = u16(at);
  if (at + 2 + n * 12 > bytes.length) throw fail('cut short');
  const fields = new Map();
  for (let i = 0; i < Math.min(n, 1000); i++) {
    const e = at + 2 + i * 12;
    const tag = u16(e);
    const type = u16(e + 2);
    const count = u32(e + 4);
    const size = TYPE_SIZE[type];
    if (!size || !count || count > 1 << 24) continue;
    const total = size * count;
    const from = total <= 4 ? e + 8 : u32(e + 8);
    if (from + total > bytes.length) continue;
    const read = {
      1: (o) => bytes[o], 2: (o) => bytes[o], 6: (o) => (bytes[o] << 24) >> 24, 7: (o) => bytes[o],
      3: u16, 8: (o) => view.getInt16(o, little), 4: u32, 9: (o) => view.getInt32(o, little),
      5: (o) => u32(o) / (u32(o + 4) || 1), 10: (o) => view.getInt32(o, little) / (view.getInt32(o + 4, little) || 1),
      11: (o) => view.getFloat32(o, little), 12: (o) => view.getFloat64(o, little),
    }[type];
    const values = new Array(count);
    for (let k = 0; k < count; k++) values[k] = read(from + k * size);
    fields.set(tag, values);
  }
  return { fields, little };
}

// ------------------------------------------------------------------ decompression

/** PackBits runs, up to `size` bytes. */
function unpackBits(src, size) {
  const out = new Uint8Array(size);
  let o = 0;
  for (let i = 0; i < src.length && o < size;) {
    const n = (src[i++] << 24) >> 24;
    if (n >= 0) {
      const end = Math.min(i + n + 1, src.length);
      for (; i < end && o < size; i++) out[o++] = src[i];
      i = end;
    } else if (n !== -128) {
      const v = src[i++] ?? 0;
      for (let k = 0; k < 1 - n && o < size; k++) out[o++] = v;
    }
  }
  return out;
}

/** TIFF's LZW (most significant bit first, codes widening one early), up to `size` bytes. */
function unLzw(src, size) {
  const out = new Uint8Array(size);
  let o = 0;
  // The table: each code's last byte, first byte, prefix code and length.
  const suffix = new Uint8Array(4096);
  const first = new Uint8Array(4096);
  const prefix = new Uint16Array(4096);
  const length = new Uint16Array(4096);
  for (let i = 0; i < 256; i++) {
    suffix[i] = i;
    first[i] = i;
    length[i] = 1;
  }
  let next = 258;
  let width = 9;
  let bitPos = 0;
  const bits = src.length * 8;
  let old = -1;
  const read = () => {
    if (bitPos + width > bits) return 257;
    let v = 0;
    for (let k = 0; k < width; k++, bitPos++) v = (v << 1) | ((src[bitPos >> 3] >> (7 - (bitPos & 7))) & 1);
    return v;
  };
  /** Write a code's string (as much of it as fits). */
  const emit = (code) => {
    let c = code;
    for (let k = length[code] - 1; k >= 0; k--) {
      if (o + k < size) out[o + k] = suffix[c];
      c = prefix[c];
    }
    o += length[code];
  };
  while (o < size) {
    const code = read();
    if (code === 257) break;
    if (code === 256) {
      next = 258;
      width = 9;
      old = -1;
      continue;
    }
    if (old < 0) {
      if (code > 255) break; // damaged: a string before the table has any
      emit(code);
      old = code;
      continue;
    }
    if (code > next || (code > 255 && code < 258)) break; // damaged
    if (next < 4096) {
      // The new string: the previous one and the first byte of this one (of itself, when it's the one being added).
      prefix[next] = old;
      suffix[next] = code === next ? first[old] : first[code];
      first[next] = first[old];
      length[next] = Math.min(4096, length[old] + 1);
      next++;
    } else if (code === next) break;
    emit(code);
    old = code;
    if (next >= (1 << width) - 1 && width < 12) width++;
  }
  return out;
}

/** zlib data inflated by the browser, stopping at `size` bytes. */
async function inflate(src, size) {
  const out = new Uint8Array(size);
  let o = 0;
  if (typeof DecompressionStream !== 'function') throw fail('no Deflate here');
  const reader = readThrough(new DecompressionStream('deflate'), src);
  try {
    while (o < size) {
      const { value, done } = await reader.read();
      if (done) break;
      const n = Math.min(value.length, size - o);
      out.set(value.subarray(0, n), o);
      o += n;
    }
  } catch {
    // Damaged data: what was inflated before it stays.
  } finally {
    reader.cancel().catch(() => {});
  }
  return out;
}

const reverseBits = (b) => {
  let r = 0;
  for (let i = 0; i < 8; i++) r = (r << 1) | ((b >> i) & 1);
  return r;
};

// ------------------------------------------------------------------ pixels

/**
 * Decode a TIFF into { width, height, rgba } (8-bit RGBA, not
 * premultiplied). Async because Deflate is inflated by the browser.
 */
export async function decodeTiff(input) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const { fields, little } = readIfd(bytes);
  const one = (tag, fallback) => fields.get(tag)?.[0] ?? fallback;
  const width = one(256, 0);
  const height = one(257, 0);
  if (!(width > 0 && height > 0 && width <= 65535 && height <= 65535)) throw fail('no size');
  if (width * height > TIFF_MAX_PIXELS) throw fail('too large');
  const spp = Math.max(1, Math.min(8, one(277, 1)));
  const bps = one(258, 1);
  if (![1, 2, 4, 8, 16].includes(bps) || (fields.get(258) || []).some((b) => b !== bps)) throw fail(`${bps}-bit samples`);
  if (one(339, 1) === 3) throw fail('floating-point samples');
  const compression = one(259, 1);
  if (![1, 5, 8, 32773, 32946].includes(compression)) throw fail(`compression ${compression}`);
  const photometric = one(262, spp >= 3 ? 2 : 1);
  if (![0, 1, 2, 3, 5].includes(photometric)) throw fail(`photometric ${photometric}`);
  const planar = spp > 1 && one(284, 1) === 2;
  const predictor = one(317, 1);
  const fillReversed = one(266, 1) === 2;
  const planes = planar ? spp : 1;
  const perPlane = planar ? 1 : spp; // samples in each plane's pixel

  // The image in pieces: strips (full rows) or tiles.
  const tiled = fields.has(322) && fields.has(324);
  const pieceW = tiled ? one(322, 0) : width;
  const pieceH = tiled ? one(323, 0) : Math.min(height, one(278, height) || height);
  if (!(pieceW > 0 && pieceH > 0) || pieceW > 65535 || pieceH > 65535) throw fail('bad pieces');
  if (tiled && (bps < 8 || pieceW * pieceH > TIFF_MAX_PIXELS)) throw fail('bad tiles');
  const offsets = fields.get(tiled ? 324 : 273) || [];
  const counts = fields.get(tiled ? 325 : 279) || [];
  const across = Math.ceil(width / pieceW);
  const down = Math.ceil(height / pieceH);
  const pieceRow = Math.ceil((pieceW * perPlane * bps) / 8);
  const rowBytes = Math.ceil((width * perPlane * bps) / 8);
  // Sensible amounts of work and memory, whatever the file says.
  if (across * down * planes > 200000) throw fail('too many pieces');
  if (pieceRow * pieceH > TIFF_MAX_BYTES || rowBytes * height * planes > TIFF_MAX_BYTES) throw fail('too large');
  const data = Array.from({ length: planes }, () => new Uint8Array(rowBytes * height));

  for (let p = 0; p < planes; p++) {
    for (let ty = 0; ty < down; ty++) {
      for (let tx = 0; tx < across; tx++) {
        const index = p * across * down + ty * across + tx;
        const from = offsets[index];
        if (!(from >= 0) || from >= bytes.length) continue;
        let raw = bytes.subarray(from, Math.min(bytes.length, from + (counts[index] > 0 ? counts[index] : bytes.length)));
        if (fillReversed) raw = raw.map(reverseBits);
        const size = pieceRow * pieceH;
        let piece;
        if (compression === 1) piece = raw;
        else if (compression === 32773) piece = unpackBits(raw, size);
        else if (compression === 5) piece = unLzw(raw, size);
        else piece = await inflate(raw, size);
        if (predictor === 2) unpredict(piece, pieceRow, pieceH, pieceW, perPlane, bps, little);
        // Into the plane, clipped to the image.
        const rows = Math.min(pieceH, height - ty * pieceH);
        const x0 = Math.floor((tx * pieceW * perPlane * bps) / 8);
        const len = Math.min(pieceRow, rowBytes - x0);
        for (let r = 0; r < rows; r++) {
          const s = r * pieceRow;
          if (s >= piece.length) break;
          data[p].set(piece.subarray(s, Math.min(piece.length, s + len)), (ty * pieceH + r) * rowBytes + x0);
        }
      }
    }
  }
  return { width, height, rgba: toRgba(data, { width, height, rowBytes, spp, perPlane, planar, bps, little, photometric, fields }) };
}

/** Undo the horizontal predictor: each sample was stored as its difference from the one before it in the row. */
function unpredict(piece, pieceRow, rows, w, spp, bps, little) {
  if (bps === 8) {
    for (let r = 0; r < rows; r++) {
      const s = r * pieceRow;
      const end = Math.min(piece.length, s + w * spp);
      for (let i = s + spp; i < end; i++) piece[i] = (piece[i] + piece[i - spp]) & 255;
    }
  } else if (bps === 16) {
    const view = new DataView(piece.buffer, piece.byteOffset, piece.byteLength);
    for (let r = 0; r < rows; r++) {
      const s = r * pieceRow;
      const end = Math.min(piece.length - 1, s + w * spp * 2);
      for (let i = s + spp * 2; i < end; i += 2) view.setUint16(i, (view.getUint16(i, little) + view.getUint16(i - spp * 2, little)) & 0xffff, little);
    }
  }
}

function toRgba(data, { width, height, rowBytes, spp, perPlane, planar, bps, little, photometric, fields }) {
  const out = new Uint8Array(width * height * 4);
  const max = (1 << bps) - 1;
  // Sample `s` of pixel x in row y, as 0…255 (or the raw value for palettes).
  const sample = (y, x, s, raw = false) => {
    const plane = planar ? data[s] : data[0];
    const i = planar ? x : x * perPlane + s;
    const row = y * rowBytes;
    let v;
    if (bps === 8) v = plane[row + i];
    else if (bps === 16) {
      const o = row + i * 2;
      v = little ? plane[o] | (plane[o + 1] << 8) : (plane[o] << 8) | plane[o + 1];
    } else {
      const bit = i * bps;
      v = (plane[row + (bit >> 3)] >> (8 - bps - (bit & 7))) & max;
    }
    return raw ? v : Math.round((v * 255) / max);
  };
  const extra = fields.get(338) || [];
  const colorSamples = photometric === 2 ? 3 : photometric === 5 ? 4 : 1;
  const hasAlpha = spp > colorSamples;
  const premultiplied = hasAlpha && extra[0] === 1;
  const map = fields.get(320);
  const entries = 1 << bps;
  if (photometric === 3 && (!map || map.length < entries * 3)) throw fail('no colour map');
  if (spp < colorSamples) throw fail('too few samples');
  for (let y = 0, o = 0; y < height; y++) {
    for (let x = 0; x < width; x++, o += 4) {
      let r;
      let g;
      let b;
      if (photometric === 2) [r, g, b] = [sample(y, x, 0), sample(y, x, 1), sample(y, x, 2)];
      else if (photometric === 3) {
        const v = sample(y, x, 0, true);
        [r, g, b] = [map[v] >> 8, map[entries + v] >> 8, map[2 * entries + v] >> 8];
      } else if (photometric === 5) {
        const k = 255 - sample(y, x, 3);
        [r, g, b] = [0, 1, 2].map((s) => Math.round(((255 - sample(y, x, s)) * k) / 255));
      } else {
        const v = sample(y, x, 0);
        r = g = b = photometric === 0 ? 255 - v : v;
      }
      let a = hasAlpha ? sample(y, x, colorSamples) : 255;
      if (premultiplied && a < 255) {
        const k = a ? 255 / a : 0;
        [r, g, b] = [Math.min(255, r * k), Math.min(255, g * k), Math.min(255, b * k)];
      }
      out[o] = r;
      out[o + 1] = g;
      out[o + 2] = b;
      out[o + 3] = a;
    }
  }
  return out;
}

/** For tests: the decompressors. */
export const _internal = { unpackBits, unLzw, inflate };
