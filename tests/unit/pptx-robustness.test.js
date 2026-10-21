import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { buildPptx } from '../../src/io/pptx-export.js';
import { importPptx } from '../../src/io/pptx-import.js';
import { Package, addMediaTo } from '../../src/io/ooxml.js';
import { newDeck, newTextBox, newImage } from '../../src/model/deck.js';

const toBuffer = (blob) => (blob.arrayBuffer ? blob.arrayBuffer() : new Response(blob).arrayBuffer());
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

function deckWith(...elements) {
  const deck = newDeck({ slides: [] });
  deck.slides.push({ id: 's1', layout: 'blank', background: null, elements, notes: '', transition: { type: 'none', duration: 0.6 }, hidden: false });
  return deck;
}

/** A .pptx of `deck`, its parts changed by `edit`, as bytes. */
async function edited(deck, edit) {
  const zip = await JSZip.loadAsync(await toBuffer(await buildPptx(deck)));
  await edit(zip);
  return zip.generateAsync({ type: 'arraybuffer' });
}

const box = (text) => Object.assign(newTextBox(10, 10), { text: `<p>${text}</p>` });

describe('damaged and hostile PowerPoint files', () => {
  it('fails with a clean error, never a TypeError', async () => {
    const cases = [
      // A relationship without a target, and one called "__proto__".
      (zip) => zip.file('ppt/_rels/presentation.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x"/><Relationship Id="__proto__" Type="y" Target="z"/></Relationships>'),
      (zip) => zip.file('ppt/presentation.xml', '<oops/>'),
      (zip) => zip.file('ppt/presentation.xml', 'not xml at all <'),
      (zip) => zip.remove('ppt/slides/slide1.xml'),
    ];
    for (const edit of cases) {
      const err = await importPptx(await edited(deckWith(box('a')), edit), 'x').catch((e) => e);
      expect(err).toBeInstanceOf(Error);
      expect(err.code).toBe('bad-file');
    }
  });

  it('reads files whose numbers are nonsense', async () => {
    const bytes = await edited(deckWith(box('hello')), async (zip) => {
      const xml = await zip.file('ppt/slides/slide1.xml').async('string');
      zip.file('ppt/slides/slide1.xml', xml.replace(/<a:off x="\d+" y="\d+"\/><a:ext cx="\d+" cy="\d+"\/>/, '<a:off x="abc" y="-1e999"/><a:ext cx="-5" cy="NaN"/>'));
      const pres = await zip.file('ppt/presentation.xml').async('string');
      zip.file('ppt/presentation.xml', pres.replace(/<p:sldSz cx="\d+" cy="\d+"/, '<p:sldSz cx="-10" cy="0"'));
    });
    const deck = await importPptx(bytes, 'x');
    const el = deck.slides[0].elements.find((e) => e.text?.includes('hello'));
    for (const k of ['x', 'y', 'w', 'h']) expect(Number.isFinite(el[k])).toBe(true);
    expect(el.w).toBeGreaterThanOrEqual(0);
    expect(deck.size).toBe('16:9');
  });

  it('leaves out groups nested beyond reason', async () => {
    const depth = 40;
    const g = '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="90" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></p:grpSpPr>';
    const deep = '<p:sp><p:nvSpPr><p:cNvPr id="91" name="deep"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr></p:sp>';
    const bytes = await edited(deckWith(box('top')), async (zip) => {
      const xml = await zip.file('ppt/slides/slide1.xml').async('string');
      zip.file('ppt/slides/slide1.xml', xml.replace('</p:spTree>', `${g.repeat(depth)}${deep}${'</p:grpSp>'.repeat(depth)}</p:spTree>`));
    });
    const deck = await importPptx(bytes, 'x');
    expect(deck.slides[0].elements.some((e) => e.text?.includes('top'))).toBe(true);
    expect(deck.slides[0].elements.some((e) => e.fill === '#ff0000')).toBe(false);
  });

  it('never reads pictures that are elsewhere (external relationships)', async () => {
    const deck = deckWith(Object.assign(newImage(PNG, 100, 100), { x: 10, y: 10 }));
    deck.media = { [deck.slides[0].elements[0].media]: PNG };
    const bytes = await edited(deck, async (zip) => {
      const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('string');
      zip.file('ppt/slides/_rels/slide1.xml.rels', rels.replace(/Target="\.\.\/media\/[^"]+"/, 'Target="https://example.invalid/x.png" TargetMode="External"'));
    });
    const fetched = [];
    const real = globalThis.fetch;
    globalThis.fetch = (url) => { fetched.push(url); return Promise.reject(new Error('no')); };
    try {
      const out = await importPptx(bytes, 'x');
      expect(out.slides[0].elements.filter((e) => e.type === 'image')).toHaveLength(0);
    } finally {
      globalThis.fetch = real;
    }
    expect(fetched).toEqual([]);
  });

  it('stops unpacking a part that is bigger than it may be (zip bombs)', async () => {
    const zip = new JSZip();
    zip.file('big.xml', `<a>${' '.repeat(5000)}</a>`);
    zip.file('small.xml', '<a/>');
    const loaded = await JSZip.loadAsync(await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' }));
    const pkg = new Package(loaded, { xml: 1000, part: 1000, total: 1e6 });
    await expect(pkg.text('big.xml')).rejects.toMatchObject({ code: 'too-big' });
    expect(await pkg.text('small.xml')).toBe('<a/>');
    // A zip that claims a part is small is still stopped once the data says otherwise.
    loaded.file('big.xml')._data.uncompressedSize = 10;
    await expect(new Package(loaded, { xml: 1000, part: 1000, total: 1e6 }).text('big.xml')).rejects.toMatchObject({ code: 'too-big' });
    // Past the total, the package is full.
    const full = new Package(loaded, { xml: 1e6, part: 1e6, total: 2000 });
    await expect(full.text('big.xml')).rejects.toMatchObject({ code: 'too-big' });
    expect(full.full).toBe(true);
  });

  it('reads each picture once, and stores it once', async () => {
    const zip = new JSZip();
    zip.file('ppt/media/a.png', PNG.split(',')[1], { base64: true });
    const pkg = new Package(await JSZip.loadAsync(await zip.generateAsync({ type: 'arraybuffer' })));
    const [a, b] = await Promise.all([pkg.dataUrl('ppt/media/a.png'), pkg.dataUrl('ppt/media/a.png')]);
    expect(a).toBe(PNG);
    expect(b).toBe(a);
    expect(pkg.used).toBe(atob(PNG.split(',')[1]).length);
    const deck = { media: { old: PNG } };
    expect(addMediaTo(deck, PNG)).toBe('old');
    const id = addMediaTo(deck, 'data:image/png;base64,AAAA');
    expect(addMediaTo(deck, 'data:image/png;base64,AAAA')).toBe(id);
    expect(Object.keys(deck.media)).toHaveLength(2);
  });
});

describe('text that could break the XML of exported files', () => {
  it('escapes font names (which pptxgenjs writes as they are) and ignores unknown transitions', async () => {
    const box = Object.assign(newTextBox(10, 10), { text: '<p><span style="font-family: \'A&lt;B&amp;C\'">x</span> y</p>' });
    box.style = { ...box.style, font: 'F"/><evil/><a:latin typeface="' };
    const deck = deckWith(box);
    deck.slides[0].transition = { type: 'constructor', duration: 'slow' };
    deck.slides[0].notes = 'a bell \u0007 rings';
    const zip = await JSZip.loadAsync(await toBuffer(await buildPptx(deck)));
    for (const name of Object.keys(zip.files).filter((n) => /\.(xml|rels)$/.test(n))) {
      const xml = await zip.file(name).async('string');
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      expect(doc.getElementsByTagName('parsererror')).toHaveLength(0);
      expect(xml).not.toMatch(/<evil|native code|\u0007/);
    }
    const back = await importPptx(await zip.generateAsync({ type: 'arraybuffer' }), 'x');
    // The names come back as names (less what can't go into CSS).
    expect(back.slides[0].elements[0].text).toContain("font-family: 'F/evil/a:latin typeface='");
  });
});
