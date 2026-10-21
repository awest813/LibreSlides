import { describe, it, expect, vi, beforeAll, afterAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { inflateSync } from 'node:zlib';
import { DecompressionStream } from 'node:stream/web';
import JSZip from 'jszip';
import { woffToSfnt, sfntTables, checksum, fontNames, wrapEot, unwrapEot, decodeFontData, toBase64, EOT_XOR, EOT_COMPRESSED } from '../../src/io/font-data.js';
import { embedFonts, usedFonts, insertFontList } from '../../src/io/pptx-fonts.js';
import { buildPptx } from '../../src/io/pptx-export.js';
import { importPptx } from '../../src/io/pptx-import.js';
import { sanitizeDeck } from '../../src/io/sanitize.js';
import { registerDeckFonts } from '../../src/render/deck-fonts.js';
import { newDeck, newTextBox } from '../../src/model/deck.js';

// The bundled font files: node_modules/@fontsource/<family>/files/<family>-latin-<weight>-<style>.woff
const fontFile = (name) => readFileSync(resolve('node_modules/@fontsource', name.replace(/-latin.*/, ''), 'files', name));
const woffFile = (name) => new Uint8Array(fontFile(`${name}.woff`));
const inflate = async (data) => new Uint8Array(inflateSync(data));
const ttf = (name) => woffToSfnt(woffFile(name), inflate);
const u32 = (b, o) => new DataView(b.buffer, b.byteOffset).getUint32(o);
const toBuffer = (blob) => (blob.arrayBuffer ? blob.arrayBuffer() : new Response(blob).arrayBuffer());
const exportZip = async (deck) => JSZip.loadAsync(await toBuffer(await buildPptx(deck)));

function deckWith(...elements) {
  const deck = newDeck({ slides: [] });
  deck.slides.push({ id: 's1', layout: 'blank', background: null, elements, notes: '', transition: { type: 'none', duration: 0.6 }, hidden: false });
  return deck;
}

function textBox(html, font) {
  const box = newTextBox(10, 10);
  box.text = html;
  if (font) box.style.font = font;
  return box;
}

describe('WOFF to TrueType', () => {
  it('rebuilds a bundled WOFF into a valid TrueType file', async () => {
    const woff = woffFile('oswald-latin-400-normal');
    const font = await woffToSfnt(woff, inflate);
    expect(u32(font, 0)).toBe(0x00010000);
    const tables = sfntTables(font);
    expect(tables.size).toBe(new DataView(woff.buffer).getUint16(12));
    for (const tag of ['cmap', 'glyf', 'head', 'hhea', 'hmtx', 'loca', 'maxp', 'name', 'OS/2', 'post']) expect(tables.has(tag)).toBe(true);
    // Each table's checksum is right (the head table's is taken with checkSumAdjustment zeroed)…
    const dir = [...Array(tables.size).keys()].map((i) => 12 + i * 16);
    for (const rec of dir) {
      const tag = String.fromCharCode(...font.subarray(rec, rec + 4));
      const data = font.slice(u32(font, rec + 8), u32(font, rec + 8) + u32(font, rec + 12));
      if (tag === 'head') data.fill(0, 8, 12);
      expect(u32(font, rec + 4)).toBe(checksum(data));
      expect(u32(font, rec + 8) % 4).toBe(0);
    }
    // …and the whole file sums to the magic number.
    expect(checksum(font)).toBe(0xb1b0afba);
    expect(fontNames(tables)).toMatchObject({ family: 'Oswald', style: 'Regular' });
  });

  it('turns down files that aren’t WOFF', async () => {
    await expect(woffToSfnt(new Uint8Array(64), inflate)).rejects.toThrow(/WOFF/);
  });
});

describe('Embedded OpenType', () => {
  it('wraps TrueType in EOT and unwraps it again', async () => {
    const font = await ttf('oswald-latin-700-normal');
    const eot = wrapEot(font);
    const v = new DataView(eot.buffer);
    expect(v.getUint32(0, true)).toBe(eot.length);
    expect(v.getUint32(4, true)).toBe(font.length);
    expect(v.getUint32(8, true)).toBe(0x00020002);
    expect(v.getUint32(12, true)).toBe(0); // not compressed, not obfuscated
    expect(v.getUint32(28, true)).toBe(700);
    expect(v.getUint16(34, true)).toBe(0x504c);
    const back = unwrapEot(eot);
    expect(back.family).toBe('Oswald');
    expect(back.data).toEqual(font);
    expect(decodeFontData(eot).mime).toBe('font/ttf');
  });

  it('undoes XOR obfuscation and refuses compressed fonts', async () => {
    const font = await ttf('oswald-latin-400-normal');
    const eot = wrapEot(font);
    const v = new DataView(eot.buffer);
    const start = eot.length - font.length;
    const xored = eot.slice();
    new DataView(xored.buffer).setUint32(12, EOT_XOR, true);
    for (let i = start; i < xored.length; i++) xored[i] ^= 0x50;
    expect(unwrapEot(xored).data).toEqual(font);
    v.setUint32(12, EOT_COMPRESSED, true);
    expect(() => unwrapEot(eot)).toThrow(expect.objectContaining({ code: 'mtx' }));
  });

  it('turns down data that isn’t a font', () => {
    expect(() => decodeFontData(new Uint8Array(200))).toThrow();
    expect(() => decodeFontData(new TextEncoder().encode('<html><script>alert(1)</script></html>'.repeat(4)))).toThrow();
  });
});

describe('Embedding fonts in PowerPoint files', () => {
  // The bundled font files, served the way the built app serves them.
  beforeAll(() => {
    vi.stubGlobal('DecompressionStream', DecompressionStream);
    vi.stubGlobal('fetch', async (url) => {
      const name = String(url).split('/').pop();
      return new Response(fontFile(name));
    });
  });
  afterAll(() => vi.unstubAllGlobals());

  it('finds the fonts a deck uses', () => {
    const deck = deckWith(textBox('<p>a</p>', 'Oswald'), textBox('<p><span style="font-family: &quot;Playfair Display&quot;, serif">b</span><font face="Lora">c</font></p>'));
    deck.slides[0].elements.push({ type: 'table', cells: [[{ text: '<span style="font-family: \'Roboto Slab\'">d</span>' }]] });
    expect([...usedFonts(deck)].sort()).toEqual(['Inter', 'Lora', 'Oswald', 'Playfair Display', 'Roboto Slab']);
  });

  it('embeds the bundled fonts a deck uses, with parts, relationships, content type and font list', async () => {
    const deck = deckWith(textBox('<p>Big</p>', 'Oswald'), textBox('<p><span style="font-family: \'Playfair Display\'">fancy</span></p>'));
    const zip = await exportZip(deck);
    const pres = await zip.file('ppt/presentation.xml').async('string');
    const rels = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
    const types = await zip.file('[Content_Types].xml').async('string');

    expect(types).toContain('<Default Extension="fntdata" ContentType="application/x-fontdata"/>');
    expect(pres).toMatch(/<p:presentation embedTrueTypeFonts="1"/);
    // In schema order: after the slide sizes, before the default text style.
    const at = (s) => pres.indexOf(s);
    expect(at('<p:notesSz')).toBeLessThan(at('<p:embeddedFontLst>'));
    expect(at('</p:embeddedFontLst>')).toBeLessThan(at('<p:defaultTextStyle>'));
    const doc = new DOMParser().parseFromString(pres, 'application/xml');
    expect(doc.getElementsByTagName('parsererror')).toHaveLength(0);

    const fonts = Object.fromEntries([...doc.getElementsByTagName('p:embeddedFont')].map((f) => [
      f.getElementsByTagName('p:font')[0].getAttribute('typeface'),
      [...f.children].slice(1).map((c) => [c.localName, c.getAttribute('r:id')]),
    ]));
    // Inter is the theme's font; Playfair Display has an italic, Oswald doesn't.
    expect(Object.keys(fonts).sort()).toEqual(['Inter', 'Oswald', 'Playfair Display']);
    expect(fonts.Oswald.map(([face]) => face)).toEqual(['regular', 'bold']);
    expect(fonts['Playfair Display'].map(([face]) => face)).toEqual(['regular', 'bold', 'italic']);
    expect(fonts.Inter.map(([face]) => face)).toEqual(['regular', 'bold', 'italic', 'boldItalic']);
    expect(pres).toContain('<p:font typeface="Oswald" charset="0"/>'); // Oswald has no PANOSE classification
    expect(pres).toMatch(/<p:font typeface="Inter" panose="[0-9A-F]{20}" charset="0"\/>/);

    for (const [, faces] of Object.entries(fonts)) {
      for (const [, id] of faces) {
        const target = new RegExp(`Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/(font\\d+\\.fntdata)"`).exec(rels)?.[1];
        expect(target).toBeTruthy();
        const part = await zip.file(`ppt/fonts/${target}`).async('uint8array');
        expect(unwrapEot(part).data.length).toBeGreaterThan(1000);
      }
    }
    // Relationship ids stay unique.
    const ids = [...rels.matchAll(/Id="(rId\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it('leaves fonts out when turned off or when they can’t be loaded', async () => {
    const pres = '<p:presentation><p:notesSz/><p:defaultTextStyle/></p:presentation>';
    const zip = new JSZip();
    zip.file('ppt/presentation.xml', pres);
    zip.file('ppt/_rels/presentation.xml.rels', '<Relationships></Relationships>');
    zip.file('[Content_Types].xml', '<Types></Types>');
    const deck = deckWith(textBox('<p>x</p>', 'Oswald'));
    await embedFonts(zip, deck, { enabled: false });
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    await embedFonts(zip, deck, { load: () => { throw new Error('offline'); } });
    await embedFonts(zip, deck, { load: async () => null });
    warn.mockRestore();
    expect(await zip.file('ppt/presentation.xml').async('string')).toBe(pres);
    expect(Object.keys(zip.files).some((f) => f.startsWith('ppt/fonts/'))).toBe(false);
  });

  it('puts the font list in schema order', () => {
    const list = '<p:embeddedFontLst/>';
    expect(insertFontList('<p:presentation x="1"><p:sldSz/><p:notesSz/><p:custShowLst/><p:extLst/></p:presentation>', list))
      .toBe('<p:presentation embedTrueTypeFonts="1" x="1"><p:sldSz/><p:notesSz/><p:embeddedFontLst/><p:custShowLst/><p:extLst/></p:presentation>');
    expect(insertFontList('<p:presentation embedTrueTypeFonts="1"><p:notesSz/><p:embeddedFontLst><old/></p:embeddedFontLst></p:presentation>', list))
      .toBe('<p:presentation embedTrueTypeFonts="1"><p:notesSz/><p:embeddedFontLst/></p:presentation>');
  });
});

describe('Reading embedded fonts', () => {
  it('brings embedded fonts into the deck and writes them out again', async () => {
    const font = await ttf('lora-latin-400-italic');
    const dataUrl = `data:font/ttf;base64,${toBase64(font)}`;
    const deck = deckWith(textBox('<p>x</p>', 'House Serif'));
    deck.fonts = { 'House Serif': { italic: dataUrl } };
    const zip = await exportZip(deck);
    const pres = await zip.file('ppt/presentation.xml').async('string');
    expect(pres).toMatch(/<p:font typeface="House Serif"[^>]*\/><p:italic r:id="rId\d+"\/><\/p:embeddedFont>/);

    const back = await importPptx(await zip.generateAsync({ type: 'arraybuffer' }), 'x');
    expect(back.fonts['House Serif']).toEqual({ italic: dataUrl });
    // Bundled fonts (Inter, the theme's) aren't stored again.
    expect(Object.keys(back.fonts)).toEqual(['House Serif']);
    expect(back.slides[0].elements[0].style.font).toBe('House Serif');
    expect(sanitizeDeck(back).fonts['House Serif'].italic).toBe(dataUrl);
  });

  it('reads obfuscated and plain TrueType parts, and skips compressed or broken ones', async () => {
    const font = await ttf('roboto-slab-latin-700-normal');
    const zip = await exportZip(deckWith(textBox('<p>x</p>')));
    const eot = wrapEot(font);
    const xored = eot.slice();
    new DataView(xored.buffer).setUint32(12, EOT_XOR, true);
    for (let i = eot.length - font.length; i < xored.length; i++) xored[i] ^= 0x50;
    const mtx = eot.slice();
    new DataView(mtx.buffer).setUint32(12, EOT_COMPRESSED, true);
    const parts = { bold: xored, italic: font, regular: mtx, boldItalic: new TextEncoder().encode('not a font at all, sorry') };
    let rels = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
    let refs = '';
    Object.entries(parts).forEach(([face, bytes], i) => {
      zip.file(`ppt/fonts/font${i + 1}.fntdata`, bytes);
      rels = rels.replace('</Relationships>', `<Relationship Id="rId9${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font${i + 1}.fntdata"/></Relationships>`);
      refs += `<p:${face} r:id="rId9${i}"/>`;
    });
    zip.file('ppt/_rels/presentation.xml.rels', rels);
    const pres = await zip.file('ppt/presentation.xml').async('string');
    zip.file('ppt/presentation.xml', insertFontList(pres, `<p:embeddedFontLst><p:embeddedFont><p:font typeface="Slab &quot;Pro&quot;; }"/>${refs}</p:embeddedFont></p:embeddedFontLst>`));

    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const deck = await importPptx(await zip.generateAsync({ type: 'arraybuffer' }), 'x');
    expect(warn.mock.calls.some(([msg]) => /MicroType Express/.test(msg))).toBe(true);
    warn.mockRestore();
    // The name keeps only safe characters.
    expect(Object.keys(deck.fonts)).toEqual(['Slab Pro']);
    const faces = deck.fonts['Slab Pro'];
    expect(Object.keys(faces).sort()).toEqual(['bold', 'italic']);
    expect(faces.bold).toBe(`data:font/ttf;base64,${toBase64(font)}`);
    expect(faces.italic).toBe(faces.bold);
  });
});

describe('Untrusted embedded fonts', () => {
  it('keeps only font data under safe names', () => {
    const ttfUrl = 'data:font/ttf;base64,AAEAAA==';
    const deck = sanitizeDeck({
      slides: [],
      fonts: {
        Good: { regular: ttfUrl, bold: 'data:font/woff2;base64,d09GMg==', italic: 'data:text/html;base64,PHNjcmlwdD4=', boldItalic: 'data:font/ttf;base64,AA"});x' },
        'Bad"; } body { x: url(y)': { regular: ttfUrl, extra: ttfUrl },
        constructor: { regular: ttfUrl },
        '!!!': { regular: ttfUrl },
        Empty: { regular: 'data:image/png;base64,AAAA' },
        Odd: 'data:font/ttf;base64,AAEAAA==',
      },
    });
    expect(deck.fonts).toEqual({
      Good: { regular: ttfUrl, bold: 'data:font/woff2;base64,d09GMg==' },
      'Bad  body  x urly': { regular: ttfUrl },
      constructor: { regular: ttfUrl },
    });
    expect(Object.getPrototypeOf(deck.fonts)).toBe(Object.prototype);
    expect(sanitizeDeck({ slides: [], fonts: 'x' }).fonts).toBeUndefined();
    expect(sanitizeDeck({ slides: [], fonts: { A: { regular: 'nope' } } }).fonts).toBeUndefined();
  });
});

describe('Showing embedded fonts', () => {
  it('adds each face to the document once, and drops faces the browser refuses', async () => {
    const added = [];
    const deleted = [];
    const fonts = { add: (f) => added.push(f), delete: (f) => deleted.push(f) };
    class FakeFontFace {
      constructor(family, data, descriptors) {
        Object.assign(this, { family, data, descriptors });
      }

      load() { return this.descriptors.style === 'italic' ? Promise.reject(new Error('bad font')) : Promise.resolve(this); }
    }
    vi.stubGlobal('FontFace', FakeFontFace);
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const doc = { fonts };
    const deck = { fonts: { 'House Serif': { regular: 'data:font/ttf;base64,AAEAAA==', italic: 'data:font/ttf;base64,AAEAAB==' } } };
    registerDeckFonts(deck, doc);
    registerDeckFonts(deck, doc);
    registerDeckFonts({ ...deck }, doc);
    await new Promise((r) => setTimeout(r, 0));
    warn.mockRestore();
    vi.unstubAllGlobals();
    expect(added.map((f) => [f.family, f.descriptors])).toEqual([
      ['House Serif', { weight: '400', style: 'normal' }],
      ['House Serif', { weight: '400', style: 'italic' }],
    ]);
    expect(new Uint8Array(added[0].data)).toEqual(new Uint8Array([0, 1, 0, 0]));
    expect(deleted).toEqual([added[1]]);
    // A deck without embedded fonts costs nothing.
    expect(() => registerDeckFonts({}, doc)).not.toThrow();
  });
});
