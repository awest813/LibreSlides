import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { buildPptx, pptxTheme, pattFillXml, tileXml } from '../../src/io/pptx-export.js';
import { importPptx } from '../../src/io/pptx-import.js';
import { fixSchemeColors } from '../../src/io/pptx-diagram.js';
import { pictureSize } from '../../src/io/images.js';
import { sanitizeDeck } from '../../src/io/sanitize.js';
import { renderElement, renderSlide } from '../../src/render/render.js';
import { themeById, backgroundCss, backgroundGradient, editBackgroundGradient, tileAlign } from '../../src/model/themes.js';
import { PATTERNS, patternPath } from '../../src/model/patterns.js';
import { newDeck, newShape, newTextBox } from '../../src/model/deck.js';

/**
 * Themes, colours and backgrounds in PowerPoint files: each slide's own
 * master and theme, colour map overrides, the theme's background pictures,
 * gradients, tiles, patterns, see-through fills, link colours, and the
 * theme the files LibreSlides writes carry.
 */
const toBuffer = (blob) => (blob.arrayBuffer ? blob.arrayBuffer() : new Response(blob).arrayBuffer());
// A 1 × 1 PNG.
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const blank = (id, elements = []) => ({ id, layout: 'blank', background: null, elements, notes: '', transition: { type: 'none', duration: 0.6 }, hidden: false });

function deckWith(...elements) {
  const deck = newDeck({ slides: [] });
  deck.slides.push(blank('s1', elements));
  return deck;
}

const zipOf = async (deck) => JSZip.loadAsync(await toBuffer(await buildPptx(deck)));
const reread = async (zip) => importPptx(await zip.generateAsync({ type: 'arraybuffer' }), 'x');

async function roundTrip(deck, edit) {
  const zip = await zipOf(deck);
  if (edit) await edit(zip);
  return reread(zip);
}

const text = (zip, name) => zip.file(name).async('string');
async function change(zip, name, fn) {
  zip.file(name, fn(await text(zip, name)));
}

/** A shape's spPr fill (the first a:solidFill in it) swapped for `fill`. */
const withFill = (xml, fill) => xml.replace(/(<a:prstGeom[\s\S]*?<\/a:prstGeom>)<a:solidFill>[\s\S]*?<\/a:solidFill>/, `$1${fill}`);

/**
 * Give an exported package a second master (with its own theme, whose
 * accent 1 is `accent1`, and one layout) and put slide 2 on it. The second
 * master comes first in presentation.xml's relationships, as many files have
 * it, though it's the second in the list of masters.
 */
async function addSecondMaster(zip, { accent1, layoutXml = (x) => x }) {
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  zip.file('ppt/theme/theme2.xml', (await text(zip, 'ppt/theme/theme1.xml')).replace(/(<a:accent1>)[\s\S]*?(<\/a:accent1>)/, `$1<a:srgbClr val="${accent1}"/>$2`));
  zip.file('ppt/slideMasters/slideMaster2.xml', await text(zip, 'ppt/slideMasters/slideMaster1.xml'));
  zip.file('ppt/slideMasters/_rels/slideMaster2.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout9.xml"/><Relationship Id="rId2" Type="${REL}/theme" Target="../theme/theme2.xml"/></Relationships>`);
  zip.file('ppt/slideLayouts/slideLayout9.xml', layoutXml(await text(zip, 'ppt/slideLayouts/slideLayout1.xml')));
  zip.file('ppt/slideLayouts/_rels/slideLayout9.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/slideMaster" Target="../slideMasters/slideMaster2.xml"/></Relationships>`);
  await change(zip, 'ppt/_rels/presentation.xml.rels', (x) => x.replace(/(<Relationships[^>]*>)/, `$1<Relationship Id="rIdM2" Type="${REL}/slideMaster" Target="slideMasters/slideMaster2.xml"/>`));
  await change(zip, 'ppt/presentation.xml', (x) => x.replace('</p:sldMasterIdLst>', '<p:sldMasterId id="2147483900" r:id="rIdM2"/></p:sldMasterIdLst>'));
  await change(zip, 'ppt/slides/_rels/slide2.xml.rels', (x) => x.replace(/slideLayouts\/slideLayout\d+\.xml/, 'slideLayouts/slideLayout9.xml'));
  await change(zip, '[Content_Types].xml', (x) => x.replace('</Types>', '<Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
    + '<Override PartName="/ppt/slideMasters/slideMaster2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
    + '<Override PartName="/ppt/slideLayouts/slideLayout9.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/></Types>'));
}

/** A two-slide deck, each with a rectangle filled with the theme's accent 1. */
async function twoMasterDeck(options) {
  const deck = deckWith(newShape('rect', 10, 10, 100, 50));
  deck.slides.push(blank('s2', [newShape('rect', 10, 10, 100, 50)]));
  return roundTrip(deck, async (zip) => {
    for (const n of [1, 2]) await change(zip, `ppt/slides/slide${n}.xml`, (x) => withFill(x, '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>'));
    await addSecondMaster(zip, options);
  });
}

describe('PowerPoint masters and themes', () => {
  it('reads each slide in its own master\'s theme', async () => {
    const back = await twoMasterDeck({ accent1: '00AA00' });
    const accent1 = themeById('simple-light').colors.accent1;
    expect(back.slides[0].elements[0].fill).toBe(accent1);
    expect(back.slides[1].elements[0].fill).toBe('#00aa00');
    // The presentation's theme is the first slide's.
    expect(back.importedTheme.colors.accent1).toBe(accent1);
  });

  it('applies a layout\'s colour map override', async () => {
    const inverted = '<p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>';
    const deck = deckWith(newShape('rect', 10, 10, 100, 50));
    deck.slides.push(blank('s2', [newShape('rect', 10, 10, 100, 50)]));
    const back = await roundTrip(deck, async (zip) => {
      for (const n of [1, 2]) await change(zip, `ppt/slides/slide${n}.xml`, (x) => withFill(x, '<a:solidFill><a:schemeClr val="bg1"/></a:solidFill>'));
      await addSecondMaster(zip, { accent1: '00AA00', layoutXml: (x) => x.replace(/<p:clrMapOvr>[\s\S]*?<\/p:clrMapOvr>|(?=<\/p:sldLayout>)/, inverted) });
    });
    const t = pptxTheme(deckWith());
    expect(back.slides[0].elements[0].fill).toBe(t.colors.lt1);
    // bg1 is dark 1 on the slide whose layout swaps them.
    expect(back.slides[1].elements[0].fill).toBe(t.colors.dk1);
  });

  it('takes a slide\'s colour map override over its layout\'s', async () => {
    const back = await roundTrip(deckWith(newShape('rect', 10, 10, 100, 50)), async (zip) => {
      await change(zip, 'ppt/slides/slide1.xml', (x) => withFill(x, '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>')
        .replace(/<p:clrMapOvr>[\s\S]*?<\/p:clrMapOvr>/, '<p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent2" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>'));
    });
    expect(back.slides[0].elements[0].fill).toBe(pptxTheme(deckWith()).colors.lt1);
  });
});

describe('PowerPoint backgrounds', () => {
  it('shows the theme\'s background picture a background refers to', async () => {
    const back = await roundTrip(deckWith(newTextBox(10, 10)), async (zip) => {
      await change(zip, 'ppt/theme/theme1.xml', (x) => x.replace(/<a:bgFillStyleLst>[\s\S]*?<\/a:bgFillStyleLst>/, '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        + '<a:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></a:bgFillStyleLst>'));
      zip.file('ppt/theme/_rels/theme1.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/wood.png"/></Relationships>');
      zip.file('ppt/media/wood.png', PNG.split(',')[1], { base64: true });
      await change(zip, 'ppt/slides/slide1.xml', (x) => x.replace(/<p:bg>[\s\S]*?<\/p:bg>/, '').replace(/<p:cSld[^>]*>/, '$&<p:bg><p:bgRef idx="1003"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'));
    });
    const bg = back.slides[0].background;
    expect(bg.type).toBe('image');
    expect(back.media[bg.media]).toBe(PNG);
  });

  it('keeps every stop of a gradient background, and radial ones radial', async () => {
    const deck = deckWith(newTextBox(10, 10));
    const stops = [{ pos: 0, color: '#ff0000' }, { pos: 0.4, color: '#00ff00', opacity: 0.5 }, { pos: 1, color: '#0000ff' }];
    deck.slides[0].background = { type: 'gradient', from: '#ff0000', to: '#0000ff', angle: 90, gradient: { type: 'linear', angle: 90, stops } };
    deck.slides.push(blank('s2'));
    deck.slides[1].background = { type: 'gradient', from: '#ffffff', to: '#000000', angle: 180, gradient: { type: 'radial', cx: 0.5, cy: -0.8, stops: [{ pos: 0, color: '#ffffff' }, { pos: 1, color: '#000000' }] } };
    const zip = await zipOf(deck);
    const xml = await text(zip, 'ppt/slides/slide1.xml');
    expect(xml.match(/<a:gs /g)).toHaveLength(3);
    // A path gradient's first stop is its centre.
    expect(await text(zip, 'ppt/slides/slide2.xml')).toMatch(/<a:gs pos="0"><a:srgbClr val="FFFFFF">[\s\S]*<a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"\/>/);
    const back = await reread(zip);
    expect(back.slides[0].background).toEqual(deck.slides[0].background);
    expect(back.slides[1].background).toEqual(deck.slides[1].background);
    // Two stops end to end are the plain kind.
    const plain = { type: 'gradient', from: '#112233', to: '#445566', angle: 135 };
    deck.slides[0].background = plain;
    expect((await roundTrip(deck)).slides[0].background).toEqual(plain);
  });

  it('draws gradient backgrounds with all their stops, and edits keep them', () => {
    const t = themeById('simple-light');
    const bg = { type: 'gradient', from: '#ff0000', to: '#0000ff', angle: 90, gradient: { type: 'radial', cx: 0.25, cy: 0.5, stops: [{ pos: 0, color: '#ff0000' }, { pos: 0.5, color: '#00ff00', opacity: 0.5 }, { pos: 1, color: '#0000ff' }] } };
    expect(backgroundCss(bg, t)).toBe('radial-gradient(farthest-corner at 25% 50%, #ff0000 0%, rgba(0, 255, 0, 0.5) 50%, #0000ff 100%)');
    const edited = editBackgroundGradient(bg, { to: '#000000' });
    expect(edited.to).toBe('#000000');
    expect(edited.gradient.stops.map((s) => s.color)).toEqual(['#ff0000', '#00ff00', '#000000']);
    expect(editBackgroundGradient(bg, { angle: 180 }).gradient).toMatchObject({ type: 'linear', angle: 180 });
    expect(bg.gradient.stops[2].color).toBe('#0000ff');
    expect(backgroundGradient({ type: 'gradient', from: '#000000', to: '#ffffff', angle: 45 })).toEqual({ type: 'linear', angle: 45, stops: [{ pos: 0, color: '#000000' }, { pos: 1, color: '#ffffff' }] });
  });

  it('keeps pattern backgrounds', async () => {
    const deck = deckWith(newTextBox(10, 10));
    deck.slides[0].background = { type: 'pattern', pattern: { preset: 'dkUpDiag', fg: '#112233', bg: '#ffeedd' } };
    const zip = await zipOf(deck);
    expect(await text(zip, 'ppt/slides/slide1.xml')).toContain('<p:bg><p:bgPr><a:pattFill prst="dkUpDiag"><a:fgClr><a:srgbClr val="112233"></a:srgbClr></a:fgClr><a:bgClr><a:srgbClr val="FFEEDD"></a:srgbClr></a:bgClr></a:pattFill>');
    expect((await reread(zip)).slides[0].background).toEqual(deck.slides[0].background);
    expect(backgroundCss(deck.slides[0].background, themeById('simple-light'))).toMatch(/^#ffeedd url\("data:image\/svg\+xml,[^"]*"\) 0 0 \/ 8px 8px repeat$/);
  });

  it('tiles tiled pictures, at their size and from where they line up', async () => {
    const deck = deckWith(newTextBox(10, 10));
    deck.media.m1 = PNG;
    // The 1 × 1 picture is ¾ pt at 96 dpi: at 400 % its tiles are 3 pt.
    deck.slides[0].background = { type: 'image', media: 'm1', color: '#000000', tile: { w: 3, h: 3, x: 1, y: 2, align: 'ctr', flip: 'xy' } };
    const zip = await zipOf(deck);
    expect(await text(zip, 'ppt/slides/slide1.xml')).toMatch(/<p:bg><p:bgPr><a:blipFill[^>]*><a:blip [^>]*>[\s\S]*?<\/a:blip><a:tile tx="12700" ty="25400" sx="400000" sy="400000" flip="xy" algn="ctr"\/><\/a:blipFill>/);
    const back = await reread(zip);
    expect(back.slides[0].background).toEqual({ ...deck.slides[0].background, media: back.slides[0].background.media });
    expect(back.media[back.slides[0].background.media]).toBe(PNG);
    const css = backgroundCss(deck.slides[0].background, themeById('simple-light'));
    expect(css).toBe('#000000 url("m1") calc(50% + 1px) calc(50% + 2px) / 3px 3px repeat');
    expect(tileAlign('br')).toEqual({ x: 1, y: 1 });
    expect(tileAlign('t')).toEqual({ x: 0.5, y: 0 });
    expect(tileAlign('nonsense')).toEqual({ x: 0, y: 0 });
  });

  it('reads a slide\'s tiled picture background at the picture\'s size', async () => {
    const back = await roundTrip(deckWith(newTextBox(10, 10)), async (zip) => {
      zip.file('ppt/media/tile.png', PNG.split(',')[1], { base64: true });
      await change(zip, 'ppt/slides/_rels/slide1.xml.rels', (x) => x.replace('</Relationships>', '<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/tile.png"/></Relationships>'));
      await change(zip, 'ppt/slides/slide1.xml', (x) => x.replace(/<p:bg>[\s\S]*?<\/p:bg>/, '').replace(/<p:cSld[^>]*>/, '$&<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdT"/><a:tile/></a:blipFill><a:effectLst/></p:bgPr></p:bg>'));
    });
    expect(back.slides[0].background.tile).toEqual({ w: 0.75, h: 0.75, x: 0, y: 0, align: 'tl' });
  });
});

describe('PowerPoint fills', () => {
  it('keeps pattern fills as patterns, drawn and written back', async () => {
    const shape = Object.assign(newShape('rect', 10, 10, 100, 50), { fill: '#2cb731', pattern: { preset: 'ltHorz', fg: '#2cb731', bg: '#ffffff' } });
    const zip = await zipOf(deckWith(shape));
    expect(await text(zip, 'ppt/slides/slide1.xml')).toContain('<a:pattFill prst="ltHorz"><a:fgClr><a:srgbClr val="2CB731"></a:srgbClr></a:fgClr>');
    const el = (await reread(zip)).slides[0].elements[0];
    expect(el.pattern).toEqual(shape.pattern);
    expect(el.fill).toBe('#2cb731');
    const div = renderElement(el, { theme: themeById('simple-light'), deck: { media: {} } });
    expect(div.querySelector('pattern path').getAttribute('d')).toBe(patternPath('ltHorz'));
    expect(div.querySelector('svg > path').getAttribute('fill')).toMatch(/^url\(#ls-patt-\d+\)$/);
  });

  it('has a drawing of every preset pattern', () => {
    const presets = ['pct5', 'pct10', 'pct20', 'pct25', 'pct30', 'pct40', 'pct50', 'pct60', 'pct70', 'pct75', 'pct80', 'pct90', 'horz', 'vert', 'ltHorz', 'ltVert', 'dkHorz', 'dkVert', 'narHorz', 'narVert', 'dashHorz', 'dashVert', 'cross',
      'dnDiag', 'upDiag', 'ltDnDiag', 'ltUpDiag', 'dkDnDiag', 'dkUpDiag', 'wdDnDiag', 'wdUpDiag', 'dashDnDiag', 'dashUpDiag', 'diagCross', 'smCheck', 'lgCheck', 'smGrid', 'lgGrid', 'dotGrid', 'smConfetti', 'lgConfetti',
      'horzBrick', 'diagBrick', 'solidDmnd', 'openDmnd', 'dotDmnd', 'plaid', 'sphere', 'weave', 'divot', 'shingle', 'wave', 'trellis', 'zigZag'];
    for (const p of presets) expect(PATTERNS[p], p).toHaveLength(8);
    const cells = (p) => PATTERNS[p].reduce((n, row) => n + row.toString(2).split('1').length - 1, 0);
    // Denser as the percentage goes up.
    const pct = ['pct5', 'pct10', 'pct20', 'pct25', 'pct30', 'pct40', 'pct50', 'pct60', 'pct70', 'pct75', 'pct80', 'pct90'].map(cells);
    expect(pct).toEqual([...pct].sort((a, b) => a - b));
    expect(pattFillXml({ preset: '<bad>', fg: '#000000', bg: '#ffffff' }, themeById('simple-light'))).toContain('prst="pct50"');
  });

  it('keeps a see-through fill see-through, not mixed with the background', async () => {
    const shape = Object.assign(newShape('rect', 10, 10, 100, 50), { text: '<p>Hi</p>', fill: 'rgba(255, 0, 0, 0.49)' });
    const zip = await zipOf(deckWith(shape));
    expect(await text(zip, 'ppt/slides/slide1.xml')).toMatch(/<a:srgbClr val="FF0000"><a:alpha val="49000"\/><\/a:srgbClr>/);
    expect((await reread(zip)).slides[0].elements[0].fill).toBe('rgba(255, 0, 0, 0.49)');
  });

  it('keeps a see-through outline see-through', async () => {
    const shape = Object.assign(newShape('rect', 10, 10, 100, 50), { fill: '#ffffff', stroke: { color: 'rgba(0, 0, 255, 0.12)', width: 1, dash: 'solid' } });
    const zip = await zipOf(deckWith(shape));
    expect(await text(zip, 'ppt/slides/slide1.xml')).toMatch(/<a:ln [^>]*><a:solidFill><a:srgbClr val="0000FF"><a:alpha val="12000"\/>/);
    expect((await reread(zip)).slides[0].elements[0].stroke.color).toBe('rgba(0, 0, 255, 0.12)');
  });

  it('keeps tiled and see-through picture fills', async () => {
    const shape = Object.assign(newShape('ellipse', 10, 10, 100, 50), { fill: '#ffffff', fillImage: { media: 'm1', opacity: 0.6, tile: { w: 1.5, h: 0.75, x: 0, y: 0, align: 'tl', flip: 'x' } } });
    const deck = deckWith(shape);
    deck.media.m1 = PNG;
    const zip = await zipOf(deck);
    expect(await text(zip, 'ppt/slides/slide1.xml')).toMatch(/<a:alphaModFix amt="60000"\/><\/a:blip><a:tile tx="0" ty="0" sx="200000" sy="100000" flip="x" algn="tl"\/>/);
    const back = await reread(zip);
    expect(back.slides[0].elements[0].fillImage).toEqual({ ...shape.fillImage, media: back.slides[0].elements[0].fillImage.media });
    // Drawn as tiles, every other one mirrored.
    const div = renderElement(back.slides[0].elements[0], { theme: themeById('simple-light'), deck: back });
    const pattern = div.querySelector('pattern');
    expect(pattern.getAttribute('width')).toBe('3');
    expect([...pattern.querySelectorAll('image')].map((i) => i.getAttribute('opacity'))).toEqual(['0.6', '0.6']);
  });

  it('takes a shape\'s line width and dash from the theme\'s line style', async () => {
    const back = await roundTrip(deckWith(newShape('rect', 10, 10, 100, 50)), async (zip) => {
      await change(zip, 'ppt/theme/theme1.xml', (x) => x.replace(/<a:lnStyleLst>[\s\S]*?<\/a:lnStyleLst>/, '<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
        + '<a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"><a:shade val="50000"/></a:schemeClr></a:solidFill><a:prstDash val="dash"/></a:ln></a:lnStyleLst>'));
      await change(zip, 'ppt/slides/slide1.xml', (x) => x.replace(/<a:ln\b[\s\S]*?<\/a:ln>/, '').replace('</p:spPr>', '</p:spPr><p:style><a:lnRef idx="2"><a:srgbClr val="FF0000"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>'));
    });
    expect(back.slides[0].elements[0].stroke).toEqual({ color: '#bc0000', width: 3, dash: 'dash' });
  });
});

describe('PowerPoint link colours', () => {
  it('gives links the theme\'s link colour, unless they have their own', async () => {
    const box = Object.assign(newTextBox(10, 10, 400, 60), { text: '<p>see <a href="https://example.com/a">a</a> and <a href="https://example.com/b">b</a></p>' });
    const back = await roundTrip(deckWith(box), async (zip) => {
      await change(zip, 'ppt/theme/theme1.xml', (x) => x.replace(/(<a:hlink>)[\s\S]*?(<\/a:hlink>)/, '$1<a:srgbClr val="0563C1"/>$2'));
      // Links in PowerPoint's own colour (without the extension saying to keep the text's): the
      // first without a colour of its own, the second red.
      await change(zip, 'ppt/slides/slide1.xml', (x) => x.split('<a:r>').map((run) => {
        const plain = run.replace(/<a:extLst>[\s\S]*?<\/a:extLst>/, '');
        if (run.includes('<a:t>a</a:t>')) return plain.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, '');
        if (run.includes('<a:t>b</a:t>')) return plain.replace(/<a:srgbClr val="[0-9A-F]+"\/>/, '<a:srgbClr val="FF0000"/>');
        return run;
      }).join('<a:r>'));
    });
    const html = back.slides[0].elements[0].text;
    expect(html).toMatch(/<a href="https:\/\/example.com\/a"><span style="color: #0563c1[;"][^>]*>a<\/span><\/a>/);
    expect(html).toMatch(/<a href="https:\/\/example.com\/b"><span style="color: #ff0000[;"][^>]*>b<\/span><\/a>/);
  });
});

describe('The theme in PowerPoint files LibreSlides writes', () => {
  it('writes the deck\'s colours into the theme', async () => {
    const deck = deckWith(newTextBox(10, 10));
    deck.themeId = 'coral';
    const zip = await zipOf(deck);
    const theme = await text(zip, 'ppt/theme/theme1.xml');
    const coral = themeById('coral');
    expect(theme).toContain(`<a:accent1><a:srgbClr val="${coral.colors.accent1.slice(1).toUpperCase()}"/></a:accent1>`);
    expect(theme).toContain(`<a:dk1><a:srgbClr val="${coral.colors.text.slice(1).toUpperCase()}"/></a:dk1>`);
    expect(theme).toContain(`<a:lt1><a:srgbClr val="${coral.background.color.slice(1).toUpperCase()}"/></a:lt1>`);
    expect(theme).toContain('<a:latin typeface="Montserrat"/>');
    expect(await text(zip, 'ppt/slideMasters/slideMaster1.xml')).toMatch(/<p:clrMap bg1="lt1" tx1="dk1"/);
  });

  it('swaps text and background in a dark theme\'s colour map', async () => {
    const deck = deckWith(newTextBox(10, 10));
    deck.themeId = 'plum';
    const zip = await zipOf(deck);
    expect(await text(zip, 'ppt/theme/theme1.xml')).toContain('<a:dk1><a:srgbClr val="2E1065"/></a:dk1>');
    expect(await text(zip, 'ppt/slideMasters/slideMaster1.xml')).toMatch(/<p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2"/);
    const back = await reread(zip);
    expect(back.importedTheme).toMatchObject({ dark: true, colors: { dk1: '#2e1065', lt1: themeById('plum').colors.text } });
  });

  it('writes back the theme a PowerPoint file came with', async () => {
    const first = await twoMasterDeck({ accent1: '00AA00' });
    first.importedTheme.colors.accent2 = '#123456';
    first.importedTheme.fonts = { major: 'Georgia', minor: 'Verdana' };
    const zip = await zipOf(first);
    const theme = await text(zip, 'ppt/theme/theme1.xml');
    expect(theme).toContain('<a:accent2><a:srgbClr val="123456"/></a:accent2>');
    expect(theme).toMatch(/<a:majorFont><a:latin typeface="Georgia"/);
    expect(theme).toMatch(/<a:minorFont><a:latin typeface="Verdana"/);
    expect((await reread(zip)).importedTheme.colors.accent2).toBe('#123456');
  });

  it('writes SmartArt\'s theme colours out as the colours they were', () => {
    const scheme = { accent1: '4F81BD', lt1: 'FFFFFF', dk1: '000000', bg1: 'FFFFFF' };
    const part = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:solidFill><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:solidFill><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></dsp:drawing>';
    const out = fixSchemeColors(part, scheme);
    expect(out.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<dsp:drawing')).toBe(true);
    expect(out).toContain('<a:srgbClr val="4F81BD"><a:shade val="50000"/></a:srgbClr>');
    expect(out).toContain('<a:srgbClr val="FFFFFF"/>');
    expect(out).toContain('<a:schemeClr val="phClr"/>');
    expect(fixSchemeColors('<a:x/>', scheme)).toBe('<a:x/>');
  });
});

describe('Theme and background odds and ends', () => {
  it('reads pictures\' sizes and resolutions from their headers', () => {
    expect(pictureSize(PNG)).toEqual({ w: 1, h: 1, dpiX: 96, dpiY: 96 });
    // A JPEG at 300 dpi (JFIF), 640 × 480.
    const jpeg = [0xff, 0xd8, 0xff, 0xe0, 0, 16, 0x4a, 0x46, 0x49, 0x46, 0, 1, 1, 1, 0x01, 0x2c, 0x01, 0x2c, 0, 0,
      0xff, 0xc0, 0, 17, 8, 0x01, 0xe0, 0x02, 0x80, 3, 1, 0x22, 0, 2, 0x11, 1, 3, 0x11, 1];
    expect(pictureSize(`data:image/jpeg;base64,${btoa(String.fromCharCode(...jpeg))}`)).toEqual({ w: 640, h: 480, dpiX: 300, dpiY: 300 });
    expect(pictureSize('data:image/png;base64,AAAA')).toBeNull();
    expect(tileXml({ w: 1.5, h: 3, x: 0, y: 0, align: 'zz' }, PNG)).toBe('<a:tile tx="0" ty="0" sx="200000" sy="400000" flip="none" algn="tl"/>');
  });

  it('cleans backgrounds, patterns, tiles and themes from untrusted files', () => {
    const deck = sanitizeDeck({
      media: { m1: PNG },
      importedTheme: { name: 'x', colors: { dk1: '#000000' } },
      slides: [
        { background: { type: 'gradient', from: '#000', to: '#fff', angle: 0, gradient: { type: 'linear', angle: 0, stops: [{ pos: 0, color: 'url(x)' }, { pos: 2, color: '#ffffff' }, { pos: 0.5, color: '#000000', opacity: 0.5 }] } } },
        { background: { type: 'pattern', pattern: { preset: 'zz', fg: '#000', bg: '#fff' } } },
        { background: { type: 'pattern', pattern: { preset: 'pct5', fg: 'url(x)', bg: '#fff' } } },
        { background: { type: 'image', media: 'm1', tile: { w: 'x', h: 1 } } },
        { background: { type: 'image', media: 'm1', tile: { w: 4, h: 4, x: 1e9, align: 'nowhere', flip: 'q' } }, elements: [{ type: 'shape', x: 0, y: 0, w: 10, h: 10, fill: '#fff', pattern: { preset: 'nope' }, fillImage: { media: 'm1', opacity: 'x', tile: { w: 2, h: 2, flip: 'xy' } } }] },
      ],
    });
    expect(deck.importedTheme).toBeUndefined();
    expect(deck.slides[0].background.gradient).toEqual({ type: 'linear', angle: 0, stops: [{ pos: 1, color: '#ffffff' }, { pos: 0.5, color: '#000000', opacity: 0.5 }] });
    expect(deck.slides[1].background).toBeNull();
    expect(deck.slides[2].background.pattern.fg).toBe('#000000');
    expect(deck.slides[3].background.tile).toBeUndefined();
    expect(deck.slides[4].background.tile).toEqual({ w: 4, h: 4, x: 1e5, y: 0, align: 'tl' });
    const el = deck.slides[4].elements[0];
    expect(el.pattern).toBeUndefined();
    expect(el.fillImage).toEqual({ media: 'm1', tile: { w: 2, h: 2, x: 0, y: 0, align: 'tl', flip: 'xy' } });
  });

  it('renders a slide whose background is a pattern', () => {
    const deck = deckWith();
    deck.slides[0].background = { type: 'pattern', pattern: { preset: 'pct50', fg: 'theme:accent1', bg: '#ffffff' } };
    const slide = renderSlide(deck.slides[0], deck);
    expect(slide.style.background).toContain('data:image/svg+xml');
  });
});
