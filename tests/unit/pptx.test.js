import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { buildPptx, htmlToRuns, toHex, fixParagraphs } from '../../src/io/pptx-export.js';
import { importPptx } from '../../src/io/pptx-import.js';
import { renderElement } from '../../src/render/render.js';
import { themeById } from '../../src/model/themes.js';
import { TEMPLATES } from '../../src/model/templates.js';
import { newDeck, newLine, newTable, newTextBox, newShape, newImage, htmlToText } from '../../src/model/deck.js';

const toBuffer = (blob) => (blob.arrayBuffer ? blob.arrayBuffer() : new Response(blob).arrayBuffer());
// A 1 × 1 PNG.
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

/** A one-slide deck holding `elements`. */
function deckWith(...elements) {
  const deck = newDeck({ slides: [] });
  deck.slides.push({ id: 's1', layout: 'blank', background: null, elements, notes: '', transition: { type: 'none', duration: 0.6 }, hidden: false });
  return deck;
}

/** Export a deck, let `edit` change the files of the .pptx, and read it back. */
async function roundTrip(deck, edit) {
  const zip = await JSZip.loadAsync(await toBuffer(await buildPptx(deck)));
  if (edit) await edit(zip);
  return importPptx(await zip.generateAsync({ type: 'arraybuffer' }), 'x');
}

const slideXml = async (deck) => (await JSZip.loadAsync(await toBuffer(await buildPptx(deck)))).file('ppt/slides/slide1.xml').async('string');

describe('PowerPoint export', () => {
  it('converts colours', () => {
    expect(toHex('#abc')).toBe('AABBCC');
    expect(toHex('rgb(255, 0, 16)')).toBe('FF0010');
    expect(toHex('rgba(0, 0, 0, 0)')).toBeNull();
    expect(toHex('transparent')).toBeNull();
  });

  it('turns HTML into paragraphs of runs', () => {
    const runs = htmlToRuns('<p>a <b>b</b></p><ul><li>c<ul><li>d</li></ul></li></ul>');
    expect(runs.map((r) => r.text)).toEqual(['a ', 'b', 'c', 'd']);
    expect(runs[1].options).toMatchObject({ bold: true, breakLine: true });
    expect(runs[2].options).toMatchObject({ bullet: true });
    expect(runs[3].options).toMatchObject({ bullet: true, indentLevel: 1 });
  });
});

describe('PowerPoint text', () => {
  it('writes one set of paragraph properties per paragraph, as PowerPoint requires', async () => {
    const box = newTextBox(10, 10);
    box.text = '<p>plain <b>bold</b> end</p><ul><li>one <i>two</i></li></ul>';
    const xml = await slideXml(deckWith(box));
    const paras = xml.match(/<a:p>[\s\S]*?<\/a:p>/g).filter((p) => p.includes('<a:r>'));
    expect(paras).toHaveLength(2);
    for (const p of paras) {
      expect(p.match(/<a:pPr\b/g)).toHaveLength(1);
      expect(p.startsWith('<a:p><a:pPr')).toBe(true);
    }
    expect(fixParagraphs('<a:p><a:pPr algn="l"/><a:r/><a:pPr algn="l"><a:buNone/></a:pPr><a:r/></a:p>')).toBe('<a:p><a:pPr algn="l"/><a:r/><a:r/></a:p>');
  });

  it('keeps line breaks inside a paragraph', async () => {
    const box = newTextBox(10, 10);
    box.text = '<p>one<br>two<br><br>four</p><p>next</p>';
    const xml = await slideXml(deckWith(box));
    expect(xml.match(/<a:br\/>/g)).toHaveLength(3);
    const back = await roundTrip(deckWith(box));
    expect(back.slides[0].elements[0].text).toBe('<p>one<br>two<br><br>four</p><p>next</p>');
  });

  it('keeps plain text plain next to bold text', async () => {
    const box = newTextBox(10, 10);
    box.text = '<p><b>12k</b></p><p>Weekly active users</p>';
    const el = (await roundTrip(deckWith(box))).slides[0].elements[0];
    expect(el.style.bold).toBe(true);
    expect(el.text).toContain('<span style="font-weight: 400">Weekly active users</span>');
  });
});

describe('PowerPoint pictures and shapes', () => {
  it('keeps a picture\'s crop', async () => {
    const pic = newImage('m1', 100, 50, 200, 150);
    pic.crop = { l: 0.5, t: 0.25, r: 0, b: 0.1 };
    const deck = deckWith(pic);
    deck.media.m1 = PNG;
    const xml = await slideXml(deck);
    expect(xml).toContain('<a:srcRect l="50000" r="0" t="25000" b="10000"/>');
    const back = (await roundTrip(deck)).slides[0].elements[0];
    expect(back.crop).toEqual({ l: 0.5, t: 0.25, r: 0, b: 0.1 });
    expect([back.x, back.y, back.w, back.h].map(Math.round)).toEqual([100, 50, 200, 150]);
  });

  it('keeps see-through shapes see-through', async () => {
    const shape = newShape('ellipse', 10, 10);
    shape.fill = '#ff0000';
    shape.opacity = 0.5;
    const back = (await roundTrip(deckWith(shape))).slides[0].elements[0];
    expect(back.opacity).toBe(0.5);
    expect(back.fill).toBe('#ff0000');
  });

  it('keeps a table header\'s white bold text', async () => {
    const table = newTable(2, 2, 100, 300, 400, 80);
    table.cells[0][0].text = '<p>Head</p>';
    table.cells[1][0].text = '<p>Body</p>';
    const t = (await roundTrip(deckWith(table))).slides[0].elements[0];
    expect(t.cells[0][0].text).toMatch(/color: #ffffff/);
    expect(t.cells[0][0].text).toMatch(/font-weight: 700/);
    expect(t.cells[1][0].text).toBe('<p>Body</p>');
  });

  it('shows the graphics of the master and layout, unless the slide hides them', async () => {
    const box = newTextBox(10, 10);
    box.text = '<p>Hello</p>';
    // A footer bar on the master, filled with the text colour lightened 75%.
    const bar = '<p:sp><p:nvSpPr><p:cNvPr id="99" name="Bar"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="6350000"/><a:ext cx="12192000" cy="508000"/></a:xfrm>'
      + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></p:spPr></p:sp>';
    const addBar = async (zip) => {
      const name = Object.keys(zip.files).find((n) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(n));
      zip.file(name, (await zip.file(name).async('string')).replace('</p:spTree>', `${bar}</p:spTree>`));
    };
    const shown = (await roundTrip(deckWith(box), addBar)).slides[0].elements;
    const footer = shown.find((e) => e.type === 'shape');
    expect(footer).toMatchObject({ fill: '#898989', x: 0, y: 500, w: 960, h: 40 });
    expect(shown.indexOf(footer)).toBeLessThan(shown.findIndex((e) => e.text === '<p>Hello</p>'));

    const hidden = (await roundTrip(deckWith(box), async (zip) => {
      await addBar(zip);
      zip.file('ppt/slides/slide1.xml', (await zip.file('ppt/slides/slide1.xml').async('string')).replace(/<p:sld(\s)/, '<p:sld showMasterSp="0"$1'));
    })).slides[0].elements;
    expect(hidden.some((e) => e.type === 'shape')).toBe(false);
  });
});

describe('PowerPoint gradients and shadows', () => {
  it('keeps linear and radial gradient fills as real gradients', async () => {
    const a = newShape('rect', 10, 10, 200, 100);
    a.gradient = { type: 'linear', angle: 90, stops: [{ pos: 0, color: '#ff0000' }, { pos: 0.5, color: '#00ff00', opacity: 0.5 }, { pos: 1, color: '#0000ff' }] };
    a.fill = '#ff0000';
    const b = newShape('ellipse', 300, 10, 100, 100);
    b.gradient = { type: 'radial', cx: 0.25, cy: 0.5, stops: [{ pos: 0, color: '#ffffff' }, { pos: 1, color: '#000000' }] };
    b.fill = '#ffffff';
    b.shadow = true;
    const xml = await slideXml(deckWith(a, b));
    expect(xml.match(/<a:gradFill/g)).toHaveLength(2);
    expect(xml).toContain('<a:lin ang="0" scaled="0"/>');
    expect(xml).toContain('<a:alpha val="50000"/>');
    expect(xml).toContain('<a:path path="circle"><a:fillToRect l="25000" t="50000" r="75000" b="50000"/></a:path>');
    const [ba, bb] = (await roundTrip(deckWith(a, b))).slides[0].elements;
    expect(ba.gradient).toEqual(a.gradient);
    expect(bb.gradient).toEqual(b.gradient);
    expect(bb.shadow).toBe(true);
    expect(ba.shadow).toBeUndefined();
  });

  it('writes gradient backgrounds as gradients, not pictures', async () => {
    const deck = deckWith(newTextBox(10, 10));
    deck.slides[0].background = { type: 'gradient', from: '#112233', to: '#445566', angle: 135 };
    const zip = await JSZip.loadAsync(await toBuffer(await buildPptx(deck)));
    expect(await zip.file('ppt/slides/slide1.xml').async('string')).toMatch(/<p:bg><p:bgPr><a:gradFill rotWithShape="1">[\s\S]*<a:lin ang="2700000"/);
    expect(Object.values(zip.files).some((f) => f.name.startsWith('ppt/media/') && !f.dir)).toBe(false);
    const back = await importPptx(await zip.generateAsync({ type: 'arraybuffer' }));
    expect(back.slides[0].background).toEqual({ type: 'gradient', from: '#112233', to: '#445566', angle: 135 });
  });

  it('takes a shape\'s gradient from the theme\'s fill style', async () => {
    const shape = newShape('rect', 0, 0, 200, 100);
    shape.fill = '#3366cc';
    const back = await roundTrip(deckWith(shape), async (zip) => {
      const xml = await zip.file('ppt/slides/slide1.xml').async('string');
      // No fill of its own; the theme's second fill style in accent 1, as PowerPoint's default shapes have.
      const styled = xml.replace(/(<a:prstGeom[\s\S]*?<\/a:prstGeom>)<a:solidFill>[\s\S]*?<\/a:solidFill>/, '$1')
        .replace('</p:spPr>', '</p:spPr><p:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="2"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>');
      zip.file('ppt/slides/slide1.xml', styled);
    });
    const el = back.slides[0].elements[0];
    expect(el.gradient?.stops).toHaveLength(3);
    expect(el.fill).toBe(el.gradient.stops[0].color);
  });
});

describe('PowerPoint pictures in shapes and SmartArt', () => {
  it('keeps a picture filling a shape, crop and all', async () => {
    const shape = newShape('ellipse', 10, 20, 200, 100);
    shape.fill = '#ffffff';
    shape.fillImage = { media: 'm1', crop: { l: 0.1, t: 0, r: 0.2, b: 0 } };
    const deck = deckWith(shape);
    deck.media.m1 = PNG;
    const zip = await JSZip.loadAsync(await toBuffer(await buildPptx(deck)));
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');
    expect(xml).toContain('<a:blipFill rotWithShape="1"><a:blip r:embed="rIdLsFill1"></a:blip><a:srcRect l="10000" t="0" r="20000" b="0"/>');
    expect(await zip.file('ppt/slides/_rels/slide1.xml.rels').async('string')).toContain('Target="../media/ls-fill-1-1.png"');
    expect(zip.file('ppt/media/ls-fill-1-1.png')).toBeTruthy();
    const back = await importPptx(await zip.generateAsync({ type: 'arraybuffer' }));
    const el = back.slides[0].elements[0];
    expect(el.shape).toBe('ellipse');
    expect(back.media[el.fillImage.media]).toBe(PNG);
    expect(el.fillImage.crop).toEqual({ l: 0.1, t: 0, r: 0.2, b: 0 });
  });

  /** A slide with a SmartArt diagram in a frame at (1 in, 1 in); `drawing` is the drawn copy's shapes (or none). */
  const withSmartArt = (drawing) => roundTrip(deckWith(), async (zip) => {
    const ns = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const frame = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Diagram"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rIdDm" r:lo="rIdLo" r:qs="rIdQs" r:cs="rIdCs"/></a:graphicData></a:graphic></p:graphicFrame>`;
    const slide = await zip.file('ppt/slides/slide1.xml').async('string');
    zip.file('ppt/slides/slide1.xml', slide.replace('</p:spTree>', `${frame}</p:spTree>`));
    const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('string');
    zip.file('ppt/slides/_rels/slide1.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rIdDm" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>'
      + '<Relationship Id="rIdDr" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing1.xml"/></Relationships>'));
    zip.file('ppt/diagrams/data1.xml', `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" ${ns}><dgm:ptLst><dgm:pt modelId="0" type="doc"><dgm:t><a:p><a:r><a:t>Doc</a:t></a:r></a:p></dgm:t></dgm:pt>`
      + '<dgm:pt modelId="1"><dgm:t><a:bodyPr/><a:p><a:r><a:t>First</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="2"><dgm:t><a:p><a:r><a:t>Second</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst>'
      + '<dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rIdDr"/></a:ext></dgm:extLst></dgm:dataModel>');
    zip.file('ppt/diagrams/drawing1.xml', `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" ${ns}><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>${drawing}</dsp:spTree></dsp:drawing>`);
  });

  it('brings in SmartArt\'s drawn shapes, in its frame', async () => {
    const back = await withSmartArt('<dsp:sp modelId="{1}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></dsp:spPr>'
      + '<dsp:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1800"/><a:t>First</a:t></a:r></a:p></dsp:txBody><dsp:txXfrm><a:off x="914400" y="0"/><a:ext cx="1828800" cy="914400"/></dsp:txXfrm></dsp:sp>');
    const [circle, label] = back.slides[0].elements;
    expect(circle).toMatchObject({ type: 'shape', shape: 'ellipse', fill: '#4472c4', text: '', x: 72, y: 72, w: 72, h: 72 });
    expect(label).toMatchObject({ type: 'text', fill: 'none', x: 144, y: 72, w: 144, h: 72 });
    expect(htmlToText(label.text)).toBe('First');
  });

  it('lists SmartArt\'s text when there\'s no drawn copy', async () => {
    const [list] = (await withSmartArt('')).slides[0].elements;
    expect(list).toMatchObject({ type: 'text', x: 72, y: 72, w: 288, h: 144 });
    expect(list.text).toBe('<ul><li>First</li><li>Second</li></ul>');
  });
});

describe('PowerPoint groups', () => {
  it('writes groups as groups and reads them back', async () => {
    const [a, b, c] = [newShape('rect', 100, 100, 100, 50), newTextBox(300, 200, 100, 50), newShape('ellipse', 600, 100)];
    b.text = '<p>Label</p>';
    a.groupId = 'g';
    b.groupId = 'g';
    const xml = await slideXml(deckWith(a, b, c));
    const group = xml.match(/<p:grpSp>[\s\S]*?<\/p:grpSp>/)[0];
    expect(group.match(/<p:sp>/g)).toHaveLength(2);
    // The group's frame is its members' box, in EMU.
    expect(group).toContain(`<a:off x="${100 * 12700}" y="${100 * 12700}"/><a:ext cx="${300 * 12700}" cy="${150 * 12700}"/>`);
    expect(xml.indexOf('<p:grpSp>')).toBeLessThan(xml.lastIndexOf('<p:sp>'));
    const els = (await roundTrip(deckWith(a, b, c))).slides[0].elements;
    expect(els.map((e) => e.type)).toEqual(['shape', 'text', 'shape']);
    expect(els[0].groupId).toBeTruthy();
    expect(els[1].groupId).toBe(els[0].groupId);
    expect(els[2].groupId).toBeUndefined();
    expect([els[1].x, els[1].y].map(Math.round)).toEqual([300, 200]);
  });

  it('turns the members of a turned group about its centre', async () => {
    const [a, b] = [newShape('rect', 0, 100, 100, 100), newShape('rect', 200, 100, 100, 100)];
    a.groupId = 'g';
    b.groupId = 'g';
    const back = await roundTrip(deckWith(a, b), async (zip) => {
      const xml = await zip.file('ppt/slides/slide1.xml').async('string');
      // Turn the group (0–300 × 100–200, centre 150, 150) a quarter turn.
      zip.file('ppt/slides/slide1.xml', xml.replace(/(<p:grpSp>[\s\S]*?<p:grpSpPr><a:xfrm)>/, '$1 rot="5400000">'));
    });
    const [ta, tb] = back.slides[0].elements;
    // The left square ends up above the centre, the right one below, both turned.
    expect([ta.x, ta.y, ta.rot].map(Math.round)).toEqual([100, 0, 90]);
    expect([tb.x, tb.y, tb.rot].map(Math.round)).toEqual([100, 200, 90]);
  });
});

describe('PowerPoint freeform shapes', () => {
  /** Export one shape, swap its preset outline for `geometry` (or another preset), and read it back. */
  const withGeometry = async (geometry) => {
    const shape = newShape('rect', 0, 0, 200, 100);
    shape.fill = '#3366cc';
    const back = await roundTrip(deckWith(shape), async (zip) => {
      const xml = await zip.file('ppt/slides/slide1.xml').async('string');
      zip.file('ppt/slides/slide1.xml', xml.replace(/<a:prstGeom prst="rect">[\s\S]*?<\/a:prstGeom>/, geometry));
    });
    return back.slides[0].elements[0];
  };
  const near = (cmds) => cmds.map(([op, ...n]) => [op, ...n.map((v) => Math.round(v * 1000) / 1000)]);

  it('reads lines, curves, arcs and named points', async () => {
    const el = await withGeometry('<a:custGeom><a:avLst/><a:gdLst/><a:pathLst>'
      + '<a:path w="200" h="100"><a:moveTo><a:pt x="0" y="100"/></a:moveTo><a:arcTo wR="100" hR="100" stAng="10800000" swAng="10800000"/><a:close/></a:path>'
      // Named points are in the shape's own units (a path without a size of its own).
      + '<a:path fill="none"><a:moveTo><a:pt x="l" y="b"/></a:moveTo><a:cubicBezTo><a:pt x="wd4" y="t"/><a:pt x="hc" y="t"/><a:pt x="r" y="vc"/></a:cubicBezTo><a:quadBezTo><a:pt x="hc" y="t"/><a:pt x="0" y="0"/></a:quadBezTo></a:path>'
      + '</a:pathLst></a:custGeom>');
    expect(el.shape).toBe('custom');
    const [half, open] = el.paths;
    // A half circle over the top, from the bottom left to the bottom right corner, in two quarter curves.
    expect(near(half.cmds)).toEqual([['M', 0, 1], ['C', 0, 0.448, 0.224, 0, 0.5, 0], ['C', 0.776, 0, 1, 0.448, 1, 1], ['Z']]);
    expect(half.fill).toBeUndefined();
    expect(open).toEqual({ cmds: [['M', 0, 1], ['C', 0.25, 0, 0.5, 0, 1, 0.5], ['Q', 0.5, 0, 0, 0]], fill: false });
  });

  it('falls back to a rectangle for formulas it can\'t work out', async () => {
    const el = await withGeometry('<a:custGeom><a:pathLst><a:path w="10" h="10"><a:moveTo><a:pt x="x1" y="0"/></a:moveTo><a:lnTo><a:pt x="10" y="10"/></a:lnTo></a:path></a:pathLst></a:custGeom>');
    expect(el.shape).toBe('rect');
    expect(el.paths).toBeUndefined();
  });

  it('draws any PowerPoint preset, with its adjust values', async () => {
    const cloud = await withGeometry('<a:prstGeom prst="cloud"><a:avLst/></a:prstGeom>');
    expect(cloud).toMatchObject({ shape: 'custom', preset: { name: 'cloud' } });
    expect(cloud.paths.length).toBeGreaterThan(1);
    // A rounded rectangle with its own corner size is drawn with it, and remembers it.
    const round = await withGeometry('<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst></a:prstGeom>');
    expect(round).toMatchObject({ shape: 'custom', preset: { name: 'roundRect', adj: { adj: 50000 } } });
    expect(round.paths[0].cmds[0]).toEqual(['M', 0, 0.5]);
    // Without its own adjust values it's LibreSlides' own shape.
    expect((await withGeometry('<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>')).shape).toBe('roundRect');
  });

  it('writes presets back as presets, with their adjust values', async () => {
    const back = await withGeometry('<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst></a:prstGeom>');
    const deck = deckWith(back);
    const cloud = { ...(await withGeometry('<a:prstGeom prst="cloud"><a:avLst/></a:prstGeom>')), id: 'c' };
    deck.slides[0].elements.push(cloud);
    const xml = await slideXml(deck);
    expect(xml).toContain('<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst></a:prstGeom>');
    expect(xml).toMatch(/<a:prstGeom prst="cloud"><a:avLst>\s*<\/a:avLst>|<a:prstGeom prst="cloud"><a:avLst\/>/);
    expect(xml).not.toContain('<a:custGeom>');
  });

  it('keeps an elbow connector\'s route and arrowhead', async () => {
    const line = newLine(0, 0, 100, 50, { arrowEnd: true });
    const back = await roundTrip(deckWith(line), async (zip) => {
      const xml = await zip.file('ppt/slides/slide1.xml').async('string');
      zip.file('ppt/slides/slide1.xml', xml.replace('prst="line"', 'prst="bentConnector3"'));
    });
    const el = back.slides[0].elements[0];
    expect(el).toMatchObject({ type: 'shape', shape: 'custom', fill: 'none', arrowEnd: true, preset: { name: 'bentConnector3' } });
    // Across, down, across.
    expect(el.paths[0].cmds.map((c) => c[0])).toEqual(['M', 'L', 'L', 'L']);
    expect(el.paths[0].fill).toBe(false);
    // Drawn with its arrowhead at the end of the route.
    const node = renderElement(el, { theme: themeById(back.themeId), deck: back });
    const head = node.querySelector('polygon').getAttribute('points').split(' ')[0].split(',').map(Number);
    expect(head.map(Math.round)).toEqual([Math.round(el.w), Math.round(el.h)]);
    const again = await slideXml(deckWith(el));
    expect(again).toContain('prst="bentConnector3"');
    expect(again).toMatch(/<a:tailEnd type="triangle"\/>/);
  });

  it('writes freeforms back as PowerPoint custom shapes, text and all', async () => {
    const shape = newShape('rect', 100, 50, 200, 100);
    shape.shape = 'custom';
    shape.fill = '#ff0000';
    shape.stroke = { color: '#000000', width: 2, dash: 'solid' };
    shape.text = '<p>Inside</p>';
    shape.paths = [
      { cmds: [['M', 0, 1], ['C', 0, 0.5, 0.25, 0, 0.5, 0], ['Q', 1, 0, 1, 1], ['Z']] },
      { cmds: [['M', 0, 0], ['L', 1, 0.5]], fill: false },
    ];
    const xml = await slideXml(deckWith(shape));
    expect(xml.match(/<a:custGeom>/g)).toHaveLength(2);
    expect(xml).toContain('<a:cubicBezTo>');
    expect(xml).toContain('<a:quadBezTo>');
    const els = (await roundTrip(deckWith(shape))).slides[0].elements;
    expect(els).toHaveLength(2);
    expect(els[0]).toMatchObject({ shape: 'custom', fill: '#ff0000', text: '<p>Inside</p>' });
    expect([els[0].x, els[0].y, els[0].w, els[0].h].map(Math.round)).toEqual([100, 50, 200, 100]);
    expect(near(els[0].paths[0].cmds)).toEqual(shape.paths[0].cmds);
    expect(els[1]).toMatchObject({ shape: 'custom', fill: 'none' });
    expect(near(els[1].paths[0].cmds)).toEqual([['M', 0, 0], ['L', 1, 0.5]]);
  });
});

describe('PowerPoint round trip', () => {
  it('keeps text, layout, notes, transitions, lines and tables', async () => {
    const deck = TEMPLATES.find((t) => t.id === 'update').build();
    deck.slides[0].notes = 'Speaker note';
    deck.slides[1].hidden = true;
    deck.slides[1].transition = { type: 'fade', duration: 0.6 };
    deck.slides[2].elements.push(newLine(10, 20, 300, 120, { arrowEnd: true }));
    const table = newTable(2, 2, 100, 300, 400, 80);
    table.cells[0][0].text = '<p>Head</p>';
    deck.slides[3].elements.push(table);

    const buffer = await toBuffer(await buildPptx(deck));
    const zip = await JSZip.loadAsync(buffer);
    const slide2 = await zip.file('ppt/slides/slide2.xml').async('string');
    expect(slide2).toContain('show="0"');
    expect(slide2).toContain('<p:transition spd="med"><p:fade/></p:transition>');

    const back = await importPptx(buffer, 'x');
    expect(back.slides).toHaveLength(deck.slides.length);
    expect(back.size).toBe('16:9');
    expect(back.slides[0].notes).toBe('Speaker note');
    expect(back.slides[1].hidden).toBe(true);
    expect(back.slides[1].transition.type).toBe('fade');

    const titleIn = deck.slides[1].elements.find((e) => e.ph === 'title');
    const titleOut = back.slides[1].elements.find((e) => e.text && htmlToText(e.text) === htmlToText(titleIn.text));
    expect(titleOut).toBeTruthy();
    expect(Math.round(titleOut.x)).toBe(Math.round(titleIn.x));
    expect(Math.round(titleOut.y)).toBe(Math.round(titleIn.y));
    expect(titleOut.style.size).toBe(34);
    expect(titleOut.style.font).toBe('Roboto Slab');
    const bodyOut = back.slides[1].elements.find((e) => e.text?.startsWith('<ul>'));
    expect(htmlToText(bodyOut.text)).toBe(htmlToText(deck.slides[1].elements.find((e) => e.ph === 'body').text));

    const line = back.slides[2].elements.find((e) => e.type === 'line');
    expect([line.x1, line.y1, line.x2, line.y2].map(Math.round)).toEqual([10, 20, 300, 120]);
    expect(line.arrowEnd).toBe(true);
    expect(back.slides[2].elements.filter((e) => e.shape === 'chevron')).toHaveLength(4);

    const t = back.slides[3].elements.find((e) => e.type === 'table');
    expect(t.cells.length).toBe(2);
    expect(htmlToText(t.cells[0][0].text)).toBe('Head');
  });

  it('reads rotated lines', async () => {
    const PptxGenJS = (await import('pptxgenjs')).default;
    const p = new PptxGenJS();
    p.layout = 'LAYOUT_WIDE';
    // A horizontal 2 in line turned 90° about its centre (1, 1) becomes vertical.
    p.addSlide().addShape('line', { x: 0, y: 1, w: 2, h: 0, rotate: 90, line: { color: '000000', width: 1 } });
    const deck = await importPptx(await p.write({ outputType: 'arraybuffer' }));
    const l = deck.slides[0].elements.find((e) => e.type === 'line');
    expect([l.x1, l.y1, l.x2, l.y2].map(Math.round)).toEqual([72, 0, 72, 144]);
  });

  it('rejects files that are not presentations', async () => {
    await expect(importPptx(new TextEncoder().encode('not a zip').buffer)).rejects.toThrow(/damaged|PowerPoint/);
  });
});
