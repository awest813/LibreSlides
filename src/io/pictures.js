import { decodeTiff } from './tiff.js';
import { metafileKind, metafileToSvg } from './metafile.js';
import { encodePng, readThrough } from './png.js';
import { toBase64 } from './font-data.js';

/**
 * Pictures in PowerPoint files that browsers can't show — TIFF, and the
 * Windows metafiles (WMF, EMF) older decks are full of — turned into ones
 * they can when the file is opened: TIFF into PNG, metafiles into SVG. This
 * module (and the decoders) is loaded only for files that have them.
 */

/** 'tiff', 'emf', 'wmf' or null, from a picture's first bytes. */
export function pictureKind(bytes) {
  if (bytes.length >= 8 && ((bytes[0] === 0x49 && bytes[1] === 0x49 && bytes[2] === 42 && bytes[3] === 0) || (bytes[0] === 0x4d && bytes[1] === 0x4d && bytes[2] === 0 && bytes[3] === 42))) return 'tiff';
  return metafileKind(bytes);
}

/** gzip data (a compressed metafile, .emz or .wmz) unpacked, up to `limit` bytes. */
async function gunzip(bytes, limit = 64 * 1024 * 1024) {
  const reader = readThrough(new DecompressionStream('gzip'), bytes);
  const chunks = [];
  let size = 0;
  try {
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      size += value.length;
      if (size > limit) throw Object.assign(new Error('Picture too large'), { code: 'bad-picture' });
      chunks.push(value);
    }
  } finally {
    reader.cancel().catch(() => {});
  }
  const out = new Uint8Array(size);
  let at = 0;
  for (const c of chunks) {
    out.set(c, at);
    at += c.length;
  }
  return out;
}

/** A picture as a data URL a browser can show (PNG or SVG); fails with an error when it can't be read. */
export async function convertPicture(bytes) {
  if (bytes[0] === 0x1f && bytes[1] === 0x8b && typeof DecompressionStream === 'function') bytes = await gunzip(bytes);
  const kind = pictureKind(bytes);
  if (kind === 'tiff') {
    const { width, height, rgba } = await decodeTiff(bytes);
    return `data:image/png;base64,${toBase64(await encodePng(width, height, rgba))}`;
  }
  if (kind === 'emf' || kind === 'wmf') {
    const svg = await metafileToSvg(bytes);
    return `data:image/svg+xml;base64,${toBase64(new TextEncoder().encode(svg))}`;
  }
  throw Object.assign(new Error('Not a picture LibreSlides can convert'), { code: 'bad-picture' });
}

/** What's shown in place of a picture that can't be read: a grey frame with a broken-picture sign. */
export const PICTURE_PLACEHOLDER = `data:image/svg+xml;base64,${btoa('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 160 120" width="160" height="120" preserveAspectRatio="none">'
  + '<rect x="1" y="1" width="158" height="118" fill="#f1f3f5" stroke="#adb5bd" stroke-width="2" stroke-dasharray="6 4"/>'
  + '<g fill="none" stroke="#868e96" stroke-width="4" stroke-linejoin="round"><path d="M58 40h44v40H58z"/><path d="M60 76l14-16 10 10 6-6 10 12"/></g>'
  + '<path d="M50 88L110 32" stroke="#868e96" stroke-width="4" stroke-linecap="round"/></svg>')}`;
