import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { buildPptx } from '../../src/io/pptx-export.js';
import { importPptx } from '../../src/io/pptx-import.js';
import { chooseAlternate, xml } from '../../src/io/ooxml.js';
import { vmlImageId } from '../../src/io/pptx-ole.js';
import { newDeck, newTextBox } from '../../src/model/deck.js';

const toBuffer = (blob) => (blob.arrayBuffer ? blob.arrayBuffer() : new Response(blob).arrayBuffer());
// A 1 × 1 PNG.
const PNG_B64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const pngBytes = () => Uint8Array.from(atob(PNG_B64), (c) => c.charCodeAt(0));

function deckWith(...elements) {
  const deck = newDeck({ slides: [] });
  deck.slides.push({ id: 's1', layout: 'blank', background: null, elements, notes: '', transition: { type: 'none', duration: 0.6 }, hidden: false });
  return deck;
}

/** A one-slide .pptx (with a text box) whose slide gets `shapes` added to its shape tree, and `edit` the rest. */
async function withShapes(shapes, edit) {
  const zip = await JSZip.loadAsync(await toBuffer(await buildPptx(deckWith(Object.assign(newTextBox(10, 10), { text: '<p>x</p>' })))));
  const xmlText = await zip.file('ppt/slides/slide1.xml').async('string');
  zip.file('ppt/slides/slide1.xml', xmlText.replace('</p:spTree>', `${shapes}</p:spTree>`));
  if (edit) await edit(zip);
  return importPptx(await zip.generateAsync({ type: 'arraybuffer' }), 'x');
}

/** Add relationships to the first slide's .rels. */
async function addRels(zip, rels, part = 'ppt/slides/_rels/slide1.xml.rels') {
  const text = await zip.file(part).async('string');
  zip.file(part, text.replace('</Relationships>', `${rels}</Relationships>`));
}

const NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
const xfrm = (tag, x, y, w, h) => `<${tag}><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></${tag}>`;
const pic = (id, rId, box = [0, 0, 1270000, 1270000]) => `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}" descr="pic ${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrm('a:xfrm', ...box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

describe('markup compatibility (mc:AlternateContent)', () => {
  it('takes the first choice it understands, else the fallback', () => {
    const doc = xml(`<r ${MC} xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
      <mc:AlternateContent><mc:Choice Requires="cx"><one/></mc:Choice><mc:Choice Requires="p14"><two/></mc:Choice><mc:Fallback><three/></mc:Fallback></mc:AlternateContent>
      <mc:AlternateContent><mc:Choice Requires="cx p14"><one/></mc:Choice><mc:Fallback><three/></mc:Fallback></mc:AlternateContent>
      <mc:AlternateContent><mc:Choice Requires="v" xmlns:v="urn:schemas-microsoft-com:vml"><one/></mc:Choice></mc:AlternateContent></r>`);
    const [a, b, c] = [...doc.documentElement.children];
    expect(chooseAlternate(a).firstElementChild.localName).toBe('two');
    expect(chooseAlternate(b).firstElementChild.localName).toBe('three');
    expect(chooseAlternate(c)).toBeNull();
  });

  it('reads the shapes in an AlternateContent: a chart PowerPoint can\'t draw is its fallback picture', async () => {
    const shapes = `<mc:AlternateContent ${MC} ${NS}><mc:Choice xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" Requires="cx2"><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="50" name="Chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>${xfrm('p:xfrm', 0, 0, 100, 100)}<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"/></a:graphic></p:graphicFrame></mc:Choice>
      <mc:Fallback>${pic(51, 'rIdPic', [1270000, 1270000, 2540000, 1270000])}</mc:Fallback></mc:AlternateContent>`;
    const deck = await withShapes(shapes, async (zip) => {
      zip.file('ppt/media/fallback.png', pngBytes());
      await addRels(zip, `<Relationship Id="rIdPic" Type="${IMAGE_REL}" Target="../media/fallback.png"/>`);
    });
    const images = deck.slides[0].elements.filter((e) => e.type === 'image');
    expect(images).toHaveLength(1);
    expect(images[0]).toMatchObject({ alt: 'pic 51' });
    expect(images[0].w).toBeGreaterThan(images[0].h * 1.9);
  });
});

describe('OLE objects', () => {
  const frame = (inner, box = [1270000, 635000, 2540000, 1270000]) => `<p:graphicFrame ${NS}><p:nvGraphicFramePr><p:cNvPr id="60" name="Object 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>${xfrm('p:xfrm', ...box)}<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">${inner}</a:graphicData></a:graphic></p:graphicFrame>`;

  it('shows an embedded object as its preview picture, in the frame\'s box', async () => {
    // As PowerPoint 2010 and later write them: VML for older readers, a picture in the fallback.
    const inner = `<mc:AlternateContent ${MC}><mc:Choice xmlns:v="urn:schemas-microsoft-com:vml" Requires="v"><p:oleObj spid="_x0000_s1027" name="Document" r:id="rIdOle" progId="Word.Document.12"><p:embed/></p:oleObj></mc:Choice>
      <mc:Fallback><p:oleObj name="Document" r:id="rIdOle" progId="Word.Document.12"><p:embed/>${pic(0, 'rIdPrev', [0, 0, 10, 10])}</p:oleObj></mc:Fallback></mc:AlternateContent>`;
    const deck = await withShapes(frame(inner), async (zip) => {
      zip.file('ppt/media/preview.png', pngBytes());
      await addRels(zip, `<Relationship Id="rIdPrev" Type="${IMAGE_REL}" Target="../media/preview.png"/>`);
    });
    const img = deck.slides[0].elements.find((e) => e.type === 'image');
    expect(img).toBeTruthy();
    // 2,540,000 EMU is 200pt, scaled to LibreSlides' slide.
    const k = img.w / 200;
    expect(img.h / k).toBeCloseTo(100, 0);
    expect(img.x / k).toBeCloseTo(100, 0);
  });

  it('shows an old-style embedded object as the picture of its shape in the VML drawing', async () => {
    const inner = '<p:oleObj spid="_x0000_s4097" name="Slide" r:id="rIdOle" imgW="10" imgH="10" progId="PowerPoint.Slide.8"><p:embed/></p:oleObj>';
    const vml = `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
 <v:shape id="_x0000_s4096" type="#_x0000_t75" style='position:absolute'><v:imagedata o:relid="rId9" o:title=""/></v:shape>
 <v:shape id="_x0000_s4097" type="#_x0000_t75" style='position:absolute;left:0;top:0;width:84pt;height:102pt'
  fillcolor="yellow [1]"><v:imagedata o:relid="rId1" o:title=""/><br></v:shape></xml>`;
    const deck = await withShapes(frame(inner), async (zip) => {
      zip.file('ppt/media/old.png', pngBytes());
      zip.file('ppt/drawings/vmlDrawing1.vml', vml);
      zip.file('ppt/drawings/_rels/vmlDrawing1.vml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${IMAGE_REL}" Target="../media/old.png"/></Relationships>`);
      await addRels(zip, '<Relationship Id="rIdVml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>');
    });
    expect(deck.slides[0].elements.filter((e) => e.type === 'image')).toHaveLength(1);
  });

  it('finds a VML shape\'s picture whatever the quoting, and nothing for a shape that isn\'t there', () => {
    expect(vmlImageId('<v:shape id=\'a\'><v:imagedata r:id="rId4"/></v:shape>', 'a')).toBe('rId4');
    expect(vmlImageId('<v:shape id="a"><v:imagedata o:relid="rId4"/></v:shape>', 'b')).toBeNull();
    expect(vmlImageId('<v:shape id="a">', 'a')).toBeNull();
  });
});
