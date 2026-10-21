import { sfntTables, buildSfnt } from './font-data.js';

/**
 * MicroType Express (MTX): the compressed form PowerPoint keeps its embedded
 * fonts in (EOT with the TTEMBED_TTCOMPRESSED flag; see font-data.js). This
 * follows the W3C Member Submission "MicroType Express (MTX) Font Format"
 * (https://www.w3.org/Submission/MTX/).
 *
 * The font data starts with a 10-byte header – a version byte, the copy limit
 * the compressor used, and where the second and third blocks start (24-bit
 * big-endian numbers) – and then three blocks, each LZCOMP-compressed:
 *   1. the font itself, with its glyf table in Compact Table Format (CTF),
 *      an empty loca table and its cvt table delta-coded;
 *   2. the values each glyph's instructions start by pushing ("push data");
 *   3. the rest of each glyph's instructions ("code").
 * Unpacking turns it back into a plain TrueType file. Everything read is
 * checked against the data's bounds, so a damaged or hostile file throws an
 * error rather than producing garbage or running away.
 */
const fail = (message) => { throw new Error(message); };
const damaged = () => fail('Damaged MicroType Express font');

const MAX_BLOCK = 1 << 24; // the most a block's 24-bit length can say (run-length encoding is held to it too)
const MAX_GLYF = 64 * 1024 * 1024;

// ------------------------------------------------------------------ bytes in, bytes out

/** Reads big-endian numbers from a byte array, failing past its end. */
function byteReader(bytes) {
  let pos = 0;
  const need = (n) => { if (pos + n > bytes.length) damaged(); };
  const r = {
    left: () => bytes.length - pos,
    u8: () => { need(1); return bytes[pos++]; },
    u16: () => { need(2); pos += 2; return (bytes[pos - 2] << 8) | bytes[pos - 1]; },
    i16: () => (r.u16() << 16) >> 16,
    bytes: (n) => { need(n); pos += n; return bytes.subarray(pos - n, pos); },
    // A 255UShort: one byte for 0–252, more for bigger numbers.
    u255: () => {
      const code = r.u8();
      if (code === 253) return r.u16();
      if (code === 254) return 506 + r.u8();
      if (code === 255) return 253 + r.u8();
      return code;
    },
  };
  return r;
}

/** A byte array that grows as it's written to (big-endian numbers). */
function byteWriter(size = 1024) {
  let buf = new Uint8Array(size);
  let length = 0;
  const room = (n) => {
    if (length + n <= buf.length) return;
    const bigger = new Uint8Array(Math.max(buf.length * 2, length + n));
    bigger.set(buf.subarray(0, length));
    buf = bigger;
  };
  const w = {
    get length() { return length; },
    u8: (v) => { room(1); buf[length++] = v; },
    u16: (v) => { room(2); buf[length++] = (v >> 8) & 0xff; buf[length++] = v & 0xff; },
    bytes: (b) => { room(b.length); buf.set(b, length); length += b.length; },
    done: () => buf.slice(0, length),
  };
  return w;
}

/** Reads a byte array one bit at a time, most significant bit first. */
function bitReader(bytes) {
  let pos = 0;
  const end = bytes.length * 8;
  const bit = () => {
    if (pos >= end) damaged();
    const b = (bytes[pos >> 3] >> (7 - (pos & 7))) & 1;
    pos++;
    return b;
  };
  const value = (n) => {
    let v = 0;
    for (let i = 0; i < n; i++) v = v * 2 + bit();
    return v;
  };
  return { bit, value };
}

// ------------------------------------------------------------------ LZCOMP

/**
 * An adaptive Huffman code for the symbols 0…range-1. The tree's nodes sit in
 * an array in order of weight, heaviest first (the root is node 1, the
 * symbols start as leaves range…2·range-1, each weighing `leafWeight`). Each
 * time a symbol is read its leaf and every node above it gain one, each first
 * trading places with the first node of the same weight, which keeps the
 * order – so symbols seen often move up the tree and get shorter codes.
 */
function adaptiveHuffman(range, leafWeight) {
  const size = 2 * range;
  const up = new Int32Array(size);
  const left = new Int32Array(size);
  const right = new Int32Array(size);
  const symbol = new Int32Array(size).fill(-1);
  const weight = new Float64Array(size);
  const leaf = new Int32Array(range);
  for (let i = 2; i < size; i++) up[i] = i >> 1;
  for (let i = 1; i < range; i++) {
    left[i] = 2 * i;
    right[i] = 2 * i + 1;
  }
  for (let s = 0; s < range; s++) {
    symbol[range + s] = s;
    weight[range + s] = leafWeight;
    leaf[s] = range + s;
  }
  for (let i = range - 1; i >= 1; i--) weight[i] = weight[2 * i] + weight[2 * i + 1];

  // Trade the subtrees at two places; each place keeps its parent.
  const swapIn = (arr, a, b) => {
    const t = arr[a];
    arr[a] = arr[b];
    arr[b] = t;
  };
  const relink = (n) => {
    if (symbol[n] >= 0) leaf[symbol[n]] = n;
    else up[left[n]] = up[right[n]] = n;
  };
  const swap = (a, b) => {
    swapIn(left, a, b);
    swapIn(right, a, b);
    swapIn(symbol, a, b);
    swapIn(weight, a, b);
    relink(a);
    relink(b);
  };

  const update = (s) => {
    let n = leaf[s];
    while (n !== 1) {
      let first = n;
      while (weight[first - 1] === weight[n]) first--;
      if (first !== n) {
        swap(n, first);
        n = first;
      }
      weight[n]++;
      n = up[n];
    }
    weight[1]++;
  };

  const read = (bits) => {
    let n = 1;
    while (symbol[n] < 0) n = bits.bit() ? right[n] : left[n];
    const s = symbol[n];
    update(s);
    return s;
  };

  return { read, update };
}

// The window starts out holding every pair of a byte below 32 and one below
// 96, then every byte four times over, so early repeats can be copies too.
const PRELOAD = (() => {
  const p = new Uint8Array(2 * 32 * 96 + 4 * 256);
  let i = 0;
  for (let a = 0; a < 32; a++) {
    for (let b = 0; b < 96; b++) {
      p[i++] = a;
      p[i++] = b;
    }
  }
  for (let a = 0; a < 256; a++) p.fill(a, i, (i += 4));
  return p;
})();

const LENGTH_MIN = 2;
const LONG_DISTANCE = 512; // copies from this far back are at least a byte longer

/**
 * Unpack an LZCOMP block. It starts with a bit saying whether the result is
 * also run-length encoded and the 24-bit length of the result, then a stream
 * of symbols from three adaptive Huffman codes:
 *  - the main code: a literal byte (0–255); a copy, whose number says how many
 *    3-bit pieces its distance has and the first 2-bit piece of its length;
 *    or one of three symbols that repeat the byte 2, 4 or 6 bytes back;
 *  - further length pieces, each with a "more follows" bit;
 *  - the distance pieces.
 * The number of distance pieces a copy may have is set by the block's length.
 */
export function lzcompDecode(bytes) {
  const bits = bitReader(bytes);
  const runLength = bits.bit();
  const length = bits.value(24);
  let pieces = 1;
  while (8 ** pieces < length) pieces++;
  const copies = 8 * pieces; // copy symbols: 256 … 256 + copies - 1
  const dups = 256 + copies; // then the three one-byte repeats
  const symbols = adaptiveHuffman(dups + 3, 1);
  // The shortest copies and the two-byte repeats start out as if seen before.
  for (const s of [256, 257, ...Array(12).fill(dups), ...Array(6).fill(dups + 1)]) symbols.update(s);
  const lengths = adaptiveHuffman(8, 3);
  const distances = adaptiveHuffman(8, 3);

  const out = new Uint8Array(PRELOAD.length + length);
  out.set(PRELOAD);
  let pos = PRELOAD.length;
  while (pos < out.length) {
    const s = symbols.read(bits);
    if (s < 256) {
      out[pos++] = s;
    } else if (s >= dups) {
      out[pos] = out[pos - 2 * (s - dups + 1)];
      pos++;
    } else {
      let piece = (s - 256) & 7;
      let n = piece & 3;
      while (piece & 4) {
        piece = lengths.read(bits);
        n = n * 4 + (piece & 3);
        if (n > length) damaged();
      }
      n += LENGTH_MIN;
      let distance = 0;
      for (let i = ((s - 256) >> 3) + 1; i > 0; i--) distance = distance * 8 + distances.read(bits);
      distance += 1;
      if (distance >= LONG_DISTANCE) n++;
      // The distance counts back from the copy's last byte.
      const from = pos - distance - n + 1;
      if (from < 0 || pos + n > out.length) damaged();
      out.copyWithin(pos, from, from + n);
      pos += n;
    }
  }
  const data = out.subarray(PRELOAD.length);
  return runLength ? unRunLength(data) : data;
}

/**
 * Undo run-length encoding: the first byte is the escape; the escape then a
 * count and a byte stands for that byte repeated, the escape then 0 for the
 * escape itself. An unfinished escape at the end is dropped.
 */
function unRunLength(data) {
  const escape = data[0];
  const out = byteWriter(data.length * 2);
  for (let i = 1; i < data.length;) {
    if (data[i] !== escape) out.u8(data[i++]);
    else if (i + 1 < data.length && data[i + 1] === 0) {
      out.u8(escape);
      i += 2;
    } else if (i + 2 < data.length) {
      if (out.length + data[i + 1] > MAX_BLOCK) damaged();
      for (let k = 0; k < data[i + 1]; k++) out.u8(data[i + 2]);
      i += 3;
    } else break;
  }
  return out.done();
}

// ------------------------------------------------------------------ push data

const HOP3 = 251; // A X A B A … written as A X HOP3 B: the value two back, B, the value two back
const HOP4 = 252; // A X A B A C A … written as A X HOP4 B C

/** A 255Short: one byte for 0–249, more for bigger or negative numbers. */
function read255Short(r, code = r.u8()) {
  if (code === 250) return -read255Short(r); // the next number, negated
  if (code === 253) return r.i16();
  if (code === 254) return 500 + r.u8();
  if (code === 255) return 250 + r.u8();
  return code;
}

/** The next `count` values of the push data. */
function readPushes(r, count) {
  if (count > r.left() * 5) damaged(); // no byte stands for more than five values
  const values = [];
  while (values.length < count) {
    const code = r.u8();
    if (code === HOP3 || code === HOP4) {
      if (values.length < 2) damaged();
      const a = values[values.length - 2];
      values.push(a, read255Short(r), a);
      if (code === HOP4) values.push(read255Short(r), a);
    } else {
      values.push(read255Short(r, code));
    }
  }
  if (values.length !== count) damaged();
  return values;
}

/**
 * TrueType instructions that push these values: runs of values that fit in a
 * byte as PUSHB/NPUSHB, the others as PUSHW/NPUSHW, 255 at most at a time
 * (PUSHB and PUSHW for runs of up to 7, as fonts usually have them).
 */
function pushInstructions(values, w) {
  for (let i = 0; i < values.length;) {
    const word = values[i] < 0 || values[i] > 255;
    let end = i + 1;
    while (end < values.length && end - i < 255 && (values[end] < 0 || values[end] > 255) === word) end++;
    const n = end - i;
    if (n < 8) w.u8((word ? 0xb8 : 0xb0) + n - 1); // PUSHW[n-1] / PUSHB[n-1]
    else {
      w.u8(word ? 0x41 : 0x40); // NPUSHW / NPUSHB
      w.u8(n);
    }
    for (; i < end; i++) {
      if (word) w.u16(values[i] & 0xffff);
      else w.u8(values[i]);
    }
  }
}

/**
 * Write a glyph's instructions, preceded by their length: its push values
 * (how many, then how many bytes of code follow, as 255UShorts in the glyph
 * record) as push instructions, then its code.
 */
function writeInstructions(r, streams, out) {
  const pushCount = r.u255();
  const codeSize = r.u255();
  const w = byteWriter(pushCount * 3 + codeSize + 16);
  pushInstructions(readPushes(streams.push, pushCount), w);
  w.bytes(streams.code.bytes(codeSize));
  if (w.length > 0xffff) damaged();
  streams.longest = Math.max(streams.longest, w.length);
  out.u16(w.length);
  out.bytes(w.done());
}

// ------------------------------------------------------------------ glyphs

/**
 * A point's move from the one before, by its flag (the low 7 bits: 128
 * shapes) and the 1–4 bytes after it in the "triplet" encoding: how many
 * bits the x and y moves take, what's added to them, and their signs.
 */
function readMove(r, flag) {
  const f = flag & 0x7f;
  const signed = (v, positive) => (positive ? v : -v);
  if (f < 10) return [0, signed(((f >> 1) << 8) + r.u8(), f & 1)];
  if (f < 20) return [signed((((f - 10) >> 1) << 8) + r.u8(), f & 1), 0];
  if (f < 84) {
    const k = f - 20;
    const b = r.u8();
    return [signed(1 + ((k >> 4) << 4) + (b >> 4), k & 1), signed(1 + (((k >> 2) & 3) << 4) + (b & 15), k & 2)];
  }
  if (f < 120) {
    const k = f - 84;
    const q = k >> 2;
    return [signed(1 + Math.floor(q / 3) * 256 + r.u8(), k & 1), signed(1 + (q % 3) * 256 + r.u8(), k & 2)];
  }
  if (f < 124) {
    const [a, b, c] = r.bytes(3);
    return [signed((a << 4) | (b >> 4), f & 1), signed(((b & 15) << 8) | c, f & 2)];
  }
  return [signed(r.u16(), f & 1), signed(r.u16(), f & 2)];
}

const ON_CURVE = 1;
const X_SHORT = 2;
const Y_SHORT = 4;
const REPEAT = 8;
const X_SAME = 0x10; // or, for a short x, positive
const Y_SAME = 0x20;

/** A coordinate move in TrueType's form: writes its bytes, returns its flag bits. */
function writeMove(d, short, same, bytes) {
  if (d === 0) return same;
  if (d >= -255 && d <= 255) {
    bytes.u8(Math.abs(d));
    return short | (d > 0 ? same : 0);
  }
  if (d < -0x8000 || d > 0x7fff) damaged();
  bytes.u16(d & 0xffff);
  return 0;
}

/**
 * A simple glyph: the end point of each contour (the first, then how far on
 * each next one is, as 255UShorts), a flag per point (top bit set for points
 * off the curve), the points' moves, then the instructions. Its bounding box
 * is worked out from the points unless the record gave one.
 */
function simpleGlyph(r, contours, bbox, streams, out) {
  const ends = [];
  for (let i = 0, end = 0; i < contours; i++) {
    end = i ? end + r.u255() : r.u255();
    if (end > 0xfffe) damaged();
    ends.push(end);
  }
  const count = ends[contours - 1] + 1;
  if (count > r.left()) damaged(); // each point has a flag byte
  const flags = r.bytes(count);
  const xs = new Int32Array(count);
  const ys = new Int32Array(count);
  for (let i = 0, x = 0, y = 0; i < count; i++) {
    const [dx, dy] = readMove(r, flags[i]);
    xs[i] = x += dx;
    ys[i] = y += dy;
    if (x < -0x8000 || x > 0x7fff || y < -0x8000 || y > 0x7fff) damaged();
  }

  // TrueType keeps the flags, with repeats run together, then all the x moves, then all the y moves.
  const xBytes = byteWriter(count * 2);
  const yBytes = byteWriter(count * 2);
  const tt = new Uint8Array(count);
  for (let i = 0; i < count; i++) {
    tt[i] = (flags[i] & 0x80 ? 0 : ON_CURVE)
      | writeMove(xs[i] - (i ? xs[i - 1] : 0), X_SHORT, X_SAME, xBytes)
      | writeMove(ys[i] - (i ? ys[i - 1] : 0), Y_SHORT, Y_SAME, yBytes);
  }
  let [xMin, yMin, xMax, yMax] = bbox || [xs[0], ys[0], xs[0], ys[0]];
  for (let i = 1; i < count && !bbox; i++) {
    xMin = Math.min(xMin, xs[i]);
    yMin = Math.min(yMin, ys[i]);
    xMax = Math.max(xMax, xs[i]);
    yMax = Math.max(yMax, ys[i]);
  }
  out.u16(contours);
  for (const v of [xMin, yMin, xMax, yMax]) out.u16(v & 0xffff);
  for (const end of ends) out.u16(end);
  writeInstructions(r, streams, out);
  for (let i = 0; i < count;) {
    let run = 1;
    while (i + run < count && tt[i + run] === tt[i] && run < 256) run++;
    if (run > 2) {
      out.u8(tt[i] | REPEAT);
      out.u8(run - 1);
      i += run;
    } else out.u8(tt[i++]);
  }
  out.bytes(xBytes.done());
  out.bytes(yBytes.done());
}

const ARGS_ARE_WORDS = 0x1;
const HAVE_SCALE = 0x8;
const MORE_COMPONENTS = 0x20;
const HAVE_XY_SCALE = 0x40;
const HAVE_2X2 = 0x80;
const HAVE_INSTRUCTIONS = 0x100;

/**
 * A composite glyph is kept as TrueType has it – its bounding box and its
 * components – except that its instructions, if it has any, are stored like
 * a simple glyph's.
 */
function compositeGlyph(r, streams, out) {
  out.u16(0xffff);
  out.bytes(r.bytes(8)); // bounding box
  let flags;
  let instructed = false;
  do {
    flags = r.u16();
    out.u16(flags);
    out.u16(r.u16()); // glyph index
    let size = flags & ARGS_ARE_WORDS ? 4 : 2;
    if (flags & HAVE_SCALE) size += 2;
    else if (flags & HAVE_XY_SCALE) size += 4;
    else if (flags & HAVE_2X2) size += 8;
    out.bytes(r.bytes(size));
    if (flags & HAVE_INSTRUCTIONS) instructed = true;
  } while (flags & MORE_COMPONENTS);
  if (instructed) writeInstructions(r, streams, out);
}

// In place of the contour count: the real count follows, then the bounding box, then the glyph.
const EXPLICIT_BBOX = 0x7fff;

/** One glyph from the CTF glyf table, written out as TrueType. */
function glyph(r, streams, out) {
  let contours = r.i16();
  let bbox = null;
  if (contours === EXPLICIT_BBOX) {
    contours = r.i16();
    bbox = [r.i16(), r.i16(), r.i16(), r.i16()];
  }
  if (contours > 0) simpleGlyph(r, contours, bbox, streams, out);
  else if (contours < 0) compositeGlyph(r, streams, out);
  // A glyph without contours has no data.
}

/**
 * The glyf and loca tables from the CTF glyph records, one per glyph in
 * order. Glyphs are padded to even lengths so short offsets can be used.
 * Returns the tables, whether the offsets are long and the longest
 * instructions.
 */
function expandGlyphs(ctf, numGlyphs, pushData, codeData, longOffsets) {
  const r = byteReader(ctf);
  const streams = { push: byteReader(pushData), code: byteReader(codeData), longest: 0 };
  const out = byteWriter(ctf.length * 3);
  const offsets = new Uint32Array(numGlyphs + 1);
  for (let g = 0; g < numGlyphs; g++) {
    offsets[g] = out.length;
    glyph(r, streams, out);
    if (out.length & 1) out.u8(0);
    if (out.length > MAX_GLYF) fail('The font is too large');
  }
  offsets[numGlyphs] = out.length;
  const long = longOffsets || out.length > 0x1fffe;
  const loca = new Uint8Array((numGlyphs + 1) * (long ? 4 : 2));
  const v = new DataView(loca.buffer);
  offsets.forEach((o, i) => (long ? v.setUint32(i * 4, o) : v.setUint16(i * 2, o / 2)));
  return { glyf: out.done(), loca, long, longest: streams.longest };
}

// ------------------------------------------------------------------ other tables

/**
 * The cvt table: a count, then each value as a change from the one before –
 * 0–237 as they are, 238 then a 16-bit number, and 239–247 or 248–255 then a
 * byte for bigger falls or rises (in steps of 238).
 */
function expandCvt(data) {
  const r = byteReader(data);
  const count = r.u16();
  if (count > r.left()) damaged();
  const out = new Uint8Array(count * 2);
  for (let i = 0, value = 0; i < count; i++) {
    const code = r.u8();
    let change = code;
    if (code === 238) change = r.i16();
    else if (code >= 248) change = (code - 247) * 238 + r.u8();
    else if (code >= 239) change = -((code - 239) * 238 + r.u8());
    value = ((value + change) << 16) >> 16;
    out[i * 2] = (value >> 8) & 0xff;
    out[i * 2 + 1] = value & 0xff;
  }
  return out;
}

/** Is this a well-formed hdmx table (device widths per pixel size) for so many glyphs? */
function plainHdmx(t, numGlyphs) {
  if (t.length < 8) return false;
  const v = new DataView(t.buffer, t.byteOffset, t.byteLength);
  const records = v.getInt16(2);
  const size = v.getInt32(4);
  return v.getUint16(0) === 0 && records >= 0 && size === ((numGlyphs + 2 + 3) & ~3) && 8 + records * size <= t.length;
}

/** Is this a well-formed VDMX table (vertical extents per pixel size)? */
function plainVdmx(t) {
  if (t.length < 6) return false;
  const v = new DataView(t.buffer, t.byteOffset, t.byteLength);
  const ratios = v.getUint16(4);
  if (v.getUint16(0) > 1 || 6 + ratios * 6 > t.length) return false;
  for (let i = 0; i < ratios; i++) {
    const group = v.getUint16(6 + ratios * 4 + i * 2);
    if (group + 4 > t.length || group + 4 + v.getUint16(group) * 6 > t.length) return false;
  }
  return true;
}

// ------------------------------------------------------------------ the font

/**
 * Unpack MTX-compressed font data (an EOT's font data, once any XOR is
 * undone) into a TrueType file. Throws an error if it isn't MTX data or is
 * damaged.
 */
export function mtxToSfnt(bytes) {
  if (bytes.length < 10 || (bytes[0] !== 2 && bytes[0] !== 3)) fail('Not MicroType Express data');
  const u24 = (o) => (bytes[o] << 16) | (bytes[o + 1] << 8) | bytes[o + 2];
  const second = u24(4);
  const third = u24(7);
  if (second < 10 || third < second || third > bytes.length) damaged();
  const font = lzcompDecode(bytes.subarray(10, second));
  const pushData = lzcompDecode(bytes.subarray(second, third));
  const code = lzcompDecode(bytes.subarray(third));

  if (font.length < 12) damaged();
  const flavor = new DataView(font.buffer, font.byteOffset, 4).getUint32(0);
  const tables = sfntTables(font);
  const head = tables.get('head')?.slice();
  const maxp = tables.get('maxp')?.slice();
  if (!head || head.length < 54 || !maxp || maxp.length < 6) damaged();
  const numGlyphs = (maxp[4] << 8) | maxp[5];
  tables.delete('loca');
  if (tables.has('glyf')) {
    const { glyf, loca, long, longest } = expandGlyphs(tables.get('glyf'), numGlyphs, pushData, code, head[51] === 1);
    tables.set('glyf', glyf);
    tables.set('loca', loca);
    head[50] = 0;
    head[51] = long ? 1 : 0; // indexToLocFormat
    // Push instructions are made afresh, which can make a glyph's instructions longer than they were.
    if (maxp.length >= 32 && longest > ((maxp[26] << 8) | maxp[27])) {
      maxp[26] = longest >> 8;
      maxp[27] = longest & 0xff;
    }
  }
  tables.set('head', head);
  tables.set('maxp', maxp);
  if (tables.has('cvt ')) tables.set('cvt ', expandCvt(tables.get('cvt ')));
  // hdmx and VDMX only list sizes for hinting; one that isn't in its plain form is left out.
  if (tables.has('hdmx') && !plainHdmx(tables.get('hdmx'), numGlyphs)) tables.delete('hdmx');
  if (tables.has('VDMX') && !plainVdmx(tables.get('VDMX'))) tables.delete('VDMX');
  return buildSfnt(flavor, tables);
}
