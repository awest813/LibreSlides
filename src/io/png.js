/**
 * A small PNG writer for pictures decoded here (TIFF, the bitmaps inside
 * WMF and EMF files): 8-bit RGBA, zlib-compressed with the browser's
 * CompressionStream, or stored uncompressed where there isn't one.
 */

let crcTable = null;
function crc32(bytes, crc = -1) {
  if (!crcTable) {
    crcTable = new Int32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      crcTable[n] = c;
    }
  }
  for (let i = 0; i < bytes.length; i++) crc = crcTable[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
  return crc;
}

function adler32(bytes) {
  let a = 1;
  let b = 0;
  for (let i = 0; i < bytes.length;) {
    const end = Math.min(bytes.length, i + 5552);
    for (; i < end; i++) {
      a += bytes[i];
      b += a;
    }
    a %= 65521;
    b %= 65521;
  }
  return ((b << 16) | a) >>> 0;
}

/** Bytes as a zlib stream of stored (uncompressed) blocks. */
function storedZlib(bytes) {
  const blocks = Math.max(1, Math.ceil(bytes.length / 65535));
  const out = new Uint8Array(2 + bytes.length + blocks * 5 + 4);
  out[0] = 0x78;
  out[1] = 0x01;
  let at = 2;
  for (let i = 0; i < blocks; i++) {
    const part = bytes.subarray(i * 65535, (i + 1) * 65535);
    out[at++] = i === blocks - 1 ? 1 : 0;
    out[at++] = part.length & 0xff;
    out[at++] = part.length >> 8;
    out[at++] = ~part.length & 0xff;
    out[at++] = (~part.length >> 8) & 0xff;
    out.set(part, at);
    at += part.length;
  }
  new DataView(out.buffer).setUint32(at, adler32(bytes));
  return out;
}

/**
 * A reader of what a CompressionStream or DecompressionStream (`stream`)
 * makes of `bytes`. Errors in the data come out of the reader.
 */
export function readThrough(stream, bytes) {
  const writer = stream.writable.getWriter();
  writer.write(bytes).catch(() => {});
  writer.close().catch(() => {});
  return stream.readable.getReader();
}

async function zlib(bytes) {
  if (typeof CompressionStream === 'function') {
    try {
      const reader = readThrough(new CompressionStream('deflate'), bytes);
      const chunks = [];
      let size = 0;
      for (;;) {
        const { value, done } = await reader.read();
        if (done) break;
        chunks.push(value);
        size += value.length;
      }
      const out = new Uint8Array(size);
      let at = 0;
      for (const c of chunks) {
        out.set(c, at);
        at += c.length;
      }
      return out;
    } catch { /* stored below */ }
  }
  return storedZlib(bytes);
}

function chunk(type, data) {
  const out = new Uint8Array(12 + data.length);
  const view = new DataView(out.buffer);
  view.setUint32(0, data.length);
  for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
  out.set(data, 8);
  view.setUint32(8 + data.length, ~crc32(out.subarray(4, 8 + data.length)) >>> 0);
  return out;
}

/** An RGBA picture (`rgba`: width × height × 4 bytes) as PNG bytes. */
export async function encodePng(width, height, rgba) {
  // Each row starts with its filter type (0, none); fully opaque pictures are written as RGB.
  let opaque = true;
  for (let i = 3; i < rgba.length; i += 4) {
    if (rgba[i] !== 255) {
      opaque = false;
      break;
    }
  }
  const channels = opaque ? 3 : 4;
  const stride = width * channels + 1;
  const raw = new Uint8Array(stride * height);
  for (let y = 0, s = 0; y < height; y++) {
    let d = y * stride + 1;
    if (opaque) {
      for (let x = 0; x < width; x++, s += 4) {
        raw[d++] = rgba[s];
        raw[d++] = rgba[s + 1];
        raw[d++] = rgba[s + 2];
      }
    } else {
      raw.set(rgba.subarray(s, s + width * 4), d);
      s += width * 4;
    }
  }
  const header = new Uint8Array(13);
  const hv = new DataView(header.buffer);
  hv.setUint32(0, width);
  hv.setUint32(4, height);
  header.set([8, opaque ? 2 : 6, 0, 0, 0], 8);
  const parts = [new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', header), chunk('IDAT', await zlib(raw)), chunk('IEND', new Uint8Array(0))];
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const p of parts) {
    out.set(p, at);
    at += p.length;
  }
  return out;
}
