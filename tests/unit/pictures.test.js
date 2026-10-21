import { describe, it, expect } from 'vitest';
import { readFileSync, readdirSync } from 'node:fs';
import { resolve } from 'node:path';
import { inflateSync } from 'node:zlib';
import JSZip from 'jszip';
import { decodeTiff } from '../../src/io/tiff.js';
import { encodePng } from '../../src/io/png.js';
import { metafileToSvg, metafileKind, decodeDib } from '../../src/io/metafile.js';
import { convertPicture, pictureKind, PICTURE_PLACEHOLDER } from '../../src/io/pictures.js';
import { buildPptx } from '../../src/io/pptx-export.js';
import { importPptx } from '../../src/io/pptx-import.js';
import { sanitizeDeck } from '../../src/io/sanitize.js';
import { newDeck, newImage } from '../../src/model/deck.js';

// Browsers have these; the test's VM context may not.
globalThis.DecompressionStream ??= (await import('node:stream/web')).DecompressionStream;
globalThis.CompressionStream ??= (await import('node:stream/web')).CompressionStream;

// Pictures made for these tests (see tests/fixtures/README.md).
const DIR = resolve('tests/fixtures/pictures');
const fixture = (name) => new Uint8Array(readFileSync(resolve(DIR, name)));
const TIFFS = readdirSync(DIR).filter((f) => f.endsWith('.tif'));
const toBuffer = (blob) => (blob.arrayBuffer ? blob.arrayBuffer() : new Response(blob).arrayBuffer());
const svgOf = (dataUrl) => (dataUrl.includes(';base64,') ? atob(dataUrl.split(',')[1]) : decodeURIComponent(dataUrl.split(',')[1]));

/** A small, repeatable random number source (mulberry32). */
function random(seed) {
  return () => {
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

describe('TIFF', () => {
  const expected = JSON.parse(readFileSync(resolve(DIR, 'tiff-pixels.json'), 'utf8'));

  it.each(TIFFS)('decodes %s as Pillow does', async (name) => {
    const { width, height, rgba } = await decodeTiff(fixture(name));
    expect([width, height]).toEqual([40, 30]);
    for (const [x, y, ...want] of expected[name.replace(/\.tif$/, '')]) {
      const o = (y * width + x) * 4;
      const got = [...rgba.subarray(o, o + 4)];
      // Colour under a see-through pixel doesn't matter.
      for (let k = 0; k < 4; k++) if (want[3] > 0 || k === 3) expect(Math.abs(got[k] - want[k]), `${name} (${x}, ${y})`).toBeLessThanOrEqual(1);
    }
  });

  it('refuses what isn\'t TIFF, or is too big, with an error', async () => {
    await expect(decodeTiff(new Uint8Array([1, 2, 3]))).rejects.toThrow(/TIFF/);
    const huge = fixture('rgb-lzw-pred.tif').slice();
    // Width and height (the first two fields) made 60000 × 60000.
    const v = new DataView(huge.buffer);
    const little = huge[0] === 0x49;
    const ifd = v.getUint32(4, little);
    for (let i = 0; i < v.getUint16(ifd, little); i++) {
      const e = ifd + 2 + i * 12;
      const tag = v.getUint16(e, little);
      if (tag === 256 || tag === 257) {
        v.setUint16(e + 2, 4, little);
        v.setUint32(e + 8, 60000, little);
      }
    }
    await expect(decodeTiff(huge)).rejects.toThrow(/too large/);
    // A big picture in tiles of one pixel: refused rather than read piece by piece.
    const tiny = fixture('tiled.tif').slice();
    const tv = new DataView(tiny.buffer);
    const tifd = tv.getUint32(4, little);
    for (let i = 0; i < tv.getUint16(tifd, little); i++) {
      const e = tifd + 2 + i * 12;
      const tag = tv.getUint16(e, little);
      if (tag === 256 || tag === 257) {
        tv.setUint16(e + 2, 4, little);
        tv.setUint32(e + 8, 4000, little);
      }
      if (tag === 322 || tag === 323) tv.setUint16(e + 8, 1, little);
    }
    const t0 = Date.now();
    await expect(decodeTiff(tiny)).rejects.toThrow(/TIFF/);
    expect(Date.now() - t0).toBeLessThan(1000);
  });
});

describe('PNG writer', () => {
  it('writes the pixels it\'s given, with alpha only where it\'s needed', async () => {
    const rgba = new Uint8Array([255, 0, 0, 255, 0, 255, 0, 128, 0, 0, 255, 255, 10, 20, 30, 0]);
    const png = await encodePng(2, 2, rgba);
    expect([...png.subarray(1, 4)].map((c) => String.fromCharCode(c)).join('')).toBe('PNG');
    const view = new DataView(png.buffer, png.byteOffset);
    expect([view.getUint32(16), view.getUint32(20), png[25]]).toEqual([2, 2, 6]);
    const idat = png.indexOf(0x49, 33); // "IDAT"
    const raw = inflateSync(png.subarray(idat + 4, idat + 4 + view.getUint32(idat - 4)));
    expect([...raw]).toEqual([0, ...rgba.subarray(0, 8), 0, ...rgba.subarray(8)]);
    const opaque = await encodePng(1, 1, new Uint8Array([1, 2, 3, 255]));
    expect(opaque[25]).toBe(2);
  });
});

describe('Windows metafiles', () => {
  it('knows them by their first bytes', () => {
    expect(metafileKind(fixture('drawing.emf'))).toBe('emf');
    expect(metafileKind(fixture('drawing.wmf'))).toBe('wmf');
    expect(pictureKind(fixture('tiled.tif'))).toBe('tiff');
    expect(metafileKind(fixture('tiled.tif'))).toBeNull();
  });

  it.each([['drawing.emf', true], ['drawing.wmf', true], ['drawing.wmf', false]])('draws %s (EMF inside a WMF read: %s) as SVG', async (name, embedded) => {
    const svg = await metafileToSvg(fixture(name), { embedded });
    const doc = new DOMParser().parseFromString(svg, 'image/svg+xml');
    expect(doc.querySelector('parsererror')).toBeNull();
    const [, , w, h] = doc.documentElement.getAttribute('viewBox').split(' ').map(Number);
    expect(w / h).toBeCloseTo(4 / 3, 1);
    for (const color of ['#3366cc', '#ff9900', '#cc0033', '#009933']) expect(svg).toContain(color);
    expect(doc.querySelector('text').textContent).toBe('Metafile');
    expect(doc.querySelector('image').getAttribute('href')).toMatch(/^data:image\/png;base64,/);
  });

  /** An EMF of some records (each [type, ...32-bit values]), with a header framing 100 × 100 device pixels. */
  function emf(...records) {
    const parts = [];
    const header = new DataView(new ArrayBuffer(88));
    header.setUint32(0, 1, true);
    header.setUint32(4, 88, true);
    [0, 0, 99, 99].forEach((n, i) => header.setInt32(8 + i * 4, n, true));
    [0, 0, 2645, 2645].forEach((n, i) => header.setInt32(24 + i * 4, n, true)); // 26.45 mm at 96 dpi: 100 pixels
    header.setUint32(40, 0x464d4520, true);
    [1000, 1000, 264, 264].forEach((n, i) => header.setInt32(72 + i * 4, n, true));
    parts.push(new Uint8Array(header.buffer));
    for (const [type, ...values] of [...records, [14, 0, 0, 0]]) {
      const v = new DataView(new ArrayBuffer(8 + values.length * 4));
      v.setUint32(0, type, true);
      v.setUint32(4, v.byteLength, true);
      values.forEach((n, i) => (Number.isInteger(n) ? v.setInt32(8 + i * 4, n, true) : v.setFloat32(8 + i * 4, n, true)));
      parts.push(new Uint8Array(v.buffer));
    }
    const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
    let at = 0;
    for (const p of parts) {
      out.set(p, at);
      at += p.length;
    }
    return out;
  }
  const pathsOf = (svg) => [...svg.matchAll(/<path d="([^"]*)"/g)].map((m) => m[1]);

  it('maps through world transforms, and saving and restoring the drawing state', async () => {
    const svg = await metafileToSvg(emf(
      [39, 1, 0, 0x0000ff, 0], [37, 1], // a red brush
      [33], [35, 2.5, 0, 0, 2.5, 0.5, 0.5], // SaveDC, scale by 2.5
      [43, 0, 0, 10, 10], // a rectangle
      [34, -1], // RestoreDC
      [43, 0, 0, 10, 10],
    ));
    const [scaled, plain] = pathsOf(svg);
    expect(scaled).toBe('M0.5,0.5L25.5,0.5L25.5,25.5L0.5,25.5Z');
    expect(plain).toBe('M0,0L10,0L10,10L0,10Z');
    expect(svg).toContain('fill="#ff0000"');
  });

  it('maps anisotropic windows into viewports, and draws arcs as curves', async () => {
    const svg = await metafileToSvg(emf(
      [17, 8], [9, 1000, 1000], [11, 100, -100], [12, 0, 100], // y up, 1000 units across 100 pixels
      [42, 0, 0, 1000, 1000], // an ellipse filling the frame
    ));
    const d = pathsOf(svg)[0];
    expect(d).toMatch(/^M100,50C/);
    expect(d.match(/C/g)).toHaveLength(4);
    const ys = [...d.matchAll(/,(-?[\d.]+)/g)].map((m) => +m[1]);
    expect(Math.min(...ys)).toBeCloseTo(0, 0);
    expect(Math.max(...ys)).toBeCloseTo(100, 0);
  });

  it('draws text at its characters\' places, upright, in its font', async () => {
    const font = new Array(23).fill(0);
    font[0] = 1; // handle
    font[1] = -20; // height
    font[5] = 700; // weight (the LOGFONT starts at font[1])
    // "Arial" in UTF-16, from the face name's first word.
    const face = [0x00720041, 0x00610069, 0x0000006c];
    face.forEach((n, i) => { font[8 + i] = n; });
    const text = [0, 0, 0, 0, 1, 1.5, 1.5, 10, 50, 2, 76, 0, 0, 0, 0, 0, 80, 0x00690048, 6, 6];
    const svg = await metafileToSvg(emf([82, ...font], [37, 1], [22, 24], [24, 0x0000ff], [84, ...text]));
    const t = new DOMParser().parseFromString(svg, 'image/svg+xml').querySelector('text');
    expect(t.textContent).toBe('Hi');
    expect(t.getAttribute('x')).toBe('0 6');
    expect(t.getAttribute('font-weight')).toBe('700');
    expect(t.getAttribute('font-family')).toContain('Arial');
    expect(t.getAttribute('transform')).toBe('matrix(1 0 0 1 10 50)');
  });

  it('draws at most one turn of an arc, however many a damaged file asks for', async () => {
    const t0 = Date.now();
    const svg = await metafileToSvg(emf([27, 60, 50], [41, 50, 50, 10, 0.0, 1e30]));
    expect(Date.now() - t0).toBeLessThan(1000);
    expect(pathsOf(svg).join('').match(/C/g).length).toBeLessThanOrEqual(4);
  });

  it('reads bitmaps of every depth, upside down or not', () => {
    // A 2 × 2, 24-bit, bottom-up DIB: rows are padded to 4 bytes.
    const bytes = new Uint8Array(40 + 16);
    const v = new DataView(bytes.buffer);
    v.setUint32(0, 40, true);
    v.setInt32(4, 2, true);
    v.setInt32(8, 2, true);
    v.setUint16(12, 1, true);
    v.setUint16(14, 24, true);
    // Bottom row: blue, green; top row: red, white (BGR).
    bytes.set([255, 0, 0, 0, 255, 0, 0, 0, 0, 0, 255, 255, 255, 255, 0, 0], 40);
    const img = decodeDib(v, 0, 40, 40, 16);
    expect([...img.rgba]).toEqual([255, 0, 0, 255, 255, 255, 255, 255, 0, 0, 255, 255, 0, 255, 0, 255]);
    // A damaged header: nothing, rather than an error.
    v.setUint32(0, 7, true);
    expect(decodeDib(v, 0, 40, 40, 16)).toBeNull();
  });
});

describe('damaged pictures', () => {
  /** Ways of damaging a file: flipped bytes, huge numbers, cut short, a piece repeated. */
  function damage(bytes, rnd) {
    const out = bytes.slice();
    const kind = Math.floor(rnd() * 4);
    if (kind === 0) for (let i = 0, n = 1 + Math.floor(rnd() * 20); i < n; i++) out[Math.floor(rnd() * out.length)] = Math.floor(rnd() * 256);
    else if (kind === 1) {
      const v = new DataView(out.buffer);
      for (let i = 0, n = 1 + Math.floor(rnd() * 4); i < n; i++) v.setUint32(Math.floor(rnd() * (out.length - 4)), [0xffffffff, 0x7fffffff, 0x80000000, 0x10000, 0][Math.floor(rnd() * 5)], true);
    } else if (kind === 2) return out.slice(0, Math.floor(rnd() * out.length));
    else {
      const a = Math.floor(rnd() * out.length);
      const b = Math.min(out.length, a + Math.floor(rnd() * 400));
      const piece = out.slice(a, b);
      const merged = new Uint8Array(out.length + piece.length);
      merged.set(out.subarray(0, b));
      merged.set(piece, b);
      merged.set(out.subarray(b), b + piece.length);
      return merged;
    }
    return out;
  }

  it.each([...TIFFS, 'drawing.emf', 'drawing.wmf'])('never hangs or fails other than cleanly on a damaged %s', async (name) => {
    const rnd = random(name.length * 7919);
    const original = fixture(name);
    const started = Date.now();
    for (let i = 0; i < 150; i++) {
      const bytes = damage(original, rnd);
      const t0 = Date.now();
      try {
        const url = await convertPicture(bytes);
        expect(url).toMatch(/^data:image\/(png|svg\+xml);base64,/);
      } catch (err) {
        expect(err).toBeInstanceOf(Error);
        expect(err.code).toBe('bad-picture');
      }
      expect(Date.now() - t0).toBeLessThan(3000);
    }
    expect(Date.now() - started).toBeLessThan(30000);
  }, 60000);
});

describe('pictures in PowerPoint files', () => {
  /** A deck with one picture, whose part is swapped for `bytes` named `name`. */
  async function deckWithPicture(bytes, name) {
    const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
    const deck = newDeck({ slides: [] });
    deck.media = { p: PNG };
    deck.slides.push({ id: 's', layout: 'blank', background: null, elements: [Object.assign(newImage('p', 100, 80), { x: 10, y: 10 })], notes: '', transition: { type: 'none', duration: 0.6 }, hidden: false });
    const zip = await JSZip.loadAsync(await toBuffer(await buildPptx(deck)));
    const part = Object.keys(zip.files).find((f) => /^ppt\/media\/[^/]+$/.test(f));
    zip.remove(part);
    zip.file(`ppt/media/${name}`, bytes);
    for (const rels of Object.keys(zip.files).filter((f) => f.endsWith('.rels'))) {
      const text = await zip.file(rels).async('string');
      zip.file(rels, text.replaceAll(part.split('/').pop(), name));
    }
    return importPptx(await zip.generateAsync({ type: 'arraybuffer' }), 'x');
  }

  it('brings in EMF and TIFF pictures, converted, and they survive cleaning', async () => {
    const emfDeck = await deckWithPicture(fixture('drawing.emf'), 'image1.emf');
    const [el] = emfDeck.slides[0].elements;
    expect(el.type).toBe('image');
    expect(emfDeck.media[el.media]).toMatch(/^data:image\/svg\+xml/);
    expect(emfDeck.importNotes).toBeUndefined();
    sanitizeDeck(emfDeck);
    const svg = svgOf(emfDeck.media[el.media]);
    expect(svg).toContain('Metafile');
    expect(svg).toMatch(/<image[^>]+href="data:image\/png;base64,/);
    const tiffDeck = await deckWithPicture(fixture('tiled.tif'), 'image1.tiff');
    expect(tiffDeck.media[tiffDeck.slides[0].elements[0].media]).toMatch(/^data:image\/png;base64,/);
  });

  it('shows a placeholder for a picture that can\'t be read, and says so', async () => {
    const deck = await deckWithPicture(new Uint8Array([1, 0, 0, 0, 9, 9, 9]), 'image1.wmf');
    expect(deck.media[deck.slides[0].elements[0].media]).toBe(PICTURE_PLACEHOLDER);
    expect(deck.importNotes).toEqual(['picture']);
    expect(Object.keys(deck)).not.toContain('importNotes');
  });
});
