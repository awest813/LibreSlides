import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { buildPptx } from '../../src/io/pptx-export.js';
import { importPptx } from '../../src/io/pptx-import.js';
import { sanitizeDeck, sanitizeDiagram, isXmlPart, DIAGRAM_LIMITS } from '../../src/io/sanitize.js';
import { layoutKind, colorsKind, readOutline, editedDataXml, stretchDrawingXml } from '../../src/io/pptx-diagram.js';
import { renderElement } from '../../src/render/render.js';
import { themeById } from '../../src/model/themes.js';
import { newDeck, newDiagram, DIAGRAM_LAYOUTS, mediaIdsOf, slideText, xmlDataUrl, xmlOfDataUrl } from '../../src/model/deck.js';
import { layoutDiagram, diagramChildren, diagramAsShapes, fitFontSize, textFits } from '../../src/model/diagram-layout.js';
import {
  normalizeLevels, parseOutline, outlineText, addNode, removeNode, promote, demote, moveUp, moveDown, outlineTree, lineAt,
} from '../../src/model/diagram-outline.js';

const toBuffer = (blob) => (blob.arrayBuffer ? blob.arrayBuffer() : new Response(blob).arrayBuffer());
const nodes = (...list) => list.map(([text, level], i) => ({ id: `n${i}`, text, level }));
const diagram = (layout, list, extra = {}) => ({ id: 'd', type: 'diagram', x: 100, y: 50, w: 600, h: 320, rot: 0, diagram: { layout, colors: 'accent1', style: 'flat', nodes: list, ...extra } });

function deckWith(...elements) {
  const deck = newDeck({ slides: [] });
  deck.slides.push({ id: 's1', layout: 'blank', background: null, elements, notes: '', transition: { type: 'none', duration: 0.6 }, hidden: false });
  return deck;
}

describe('diagram outline', () => {
  const list = () => nodes(['A', 0], ['A1', 1], ['A2', 1], ['A2x', 2], ['B', 0], ['B1', 1]);

  it('keeps levels an outline can have', () => {
    expect(normalizeLevels(nodes(['a', 2], ['b', 5], ['c', 1])).map((n) => n.level)).toEqual([0, 1, 1]);
  });

  it('reads and writes the text pane, a line per node indented with tabs (or pairs of spaces)', () => {
    expect(outlineText(list())).toBe('A\n\tA1\n\tA2\n\t\tA2x\nB\n\tB1');
    const parsed = parseOutline('One\n\tTwo\n    Three\n  Four');
    expect(parsed.map((n) => [n.text, n.level])).toEqual([['One', 0], ['Two', 1], ['Three', 2], ['Four', 1]]);
    expect(parseOutline('  \n ')).toEqual([]);
    expect(lineAt('a\nb\nc', 3)).toBe(1);
  });

  it('keeps nodes their ids as their lines are edited, added and removed', () => {
    const before = list();
    const after = parseOutline('A\n\tA1 changed\n\tA2\n\t\tA2x\nNew\nB', before);
    expect(after.map((n) => n.id)).toEqual(['n0', 'n1', 'n2', 'n3', after[4].id, 'n4']);
    expect(after[4].id).not.toMatch(/^n\d$/);
  });

  it('adds, removes, promotes, demotes and moves nodes with what is under them', () => {
    const l = list();
    expect(addNode(l, 1).nodes.map((n) => n.text)).toEqual(['A', 'A1', '', 'A2', 'A2x', 'B', 'B1']);
    expect(addNode(l, 2)).toMatchObject({ index: 4 });
    expect(addNode([], -1).nodes).toHaveLength(1);
    // Removing a node moves the ones under it up.
    expect(removeNode(l, 2).nodes.map((n) => [n.text, n.level])).toEqual([['A', 0], ['A1', 1], ['A2x', 1], ['B', 0], ['B1', 1]]);
    expect(demote(l, 2).nodes.map((n) => n.level)).toEqual([0, 1, 2, 3, 0, 1]);
    expect(demote(l, 1).nodes).toBe(l); // a first child has nothing to go under
    expect(promote(l, 2).nodes.map((n) => n.level)).toEqual([0, 1, 0, 1, 0, 1]);
    expect(promote(l, 0).nodes).toBe(l);
    const up = moveUp(l, 4);
    expect(up.nodes.map((n) => n.text)).toEqual(['B', 'B1', 'A', 'A1', 'A2', 'A2x']);
    expect(up.index).toBe(0);
    const down = moveDown(l, 1);
    expect(down.nodes.map((n) => n.text)).toEqual(['A', 'A2', 'A2x', 'A1', 'B', 'B1']);
    expect(down.index).toBe(3);
    expect(moveDown(l, 4).nodes).toBe(l);
    expect(outlineTree(l).map((t) => [t.node.text, t.children.length])).toEqual([['A', 2], ['B', 1]]);
  });
});

describe('diagram layout', () => {
  it('draws a shape per top-level node, the deeper ones as bullets in it', () => {
    const out = layoutDiagram(diagram('list', nodes(['One', 0], ['Sub', 1], ['Two', 0], ['Three', 0])));
    expect(out.map((e) => e.shape)).toEqual(['rect', 'rect', 'rect']);
    expect(out[0].text).toBe('<p style="text-align: center">One</p><ul><li>Sub</li></ul>');
    expect(out[1].text).toBe('<p>Two</p>');
  });

  it('puts arrows between the steps of a process, and lines between the levels of a hierarchy', () => {
    const process = layoutDiagram(diagram('process', nodes(['a', 0], ['b', 0], ['c', 0])));
    expect(process.map((e) => e.shape)).toEqual(['roundRect', 'rightArrow', 'roundRect', 'rightArrow', 'roundRect']);
    const tree = layoutDiagram(diagram('hierarchy', nodes(['Boss', 0], ['A', 1], ['A1', 2], ['B', 1])));
    const boxes = tree.filter((e) => e.type === 'shape');
    expect(boxes.map((e) => htmlText(e.text))).toEqual(['A1', 'A', 'B', 'Boss']);
    const y = Object.fromEntries(boxes.map((e) => [htmlText(e.text), e.y]));
    expect(y.Boss).toBeLessThan(y.A);
    expect(y.A).toBeLessThan(y.A1);
    expect(y.A).toBe(y.B);
    expect(tree.filter((e) => e.type === 'line').length).toBeGreaterThanOrEqual(4);
  });

  it('keeps every layout inside its box, with one text size for its shapes', () => {
    for (const { id } of DIAGRAM_LAYOUTS) {
      const el = { ...newDiagram(id, 0, 0, 500, 300) };
      const out = layoutDiagram(el);
      expect(out.length, id).toBeGreaterThan(0);
      for (const c of out) {
        const box = c.type === 'line' ? { x: Math.min(c.x1, c.x2), y: Math.min(c.y1, c.y2), w: Math.abs(c.x2 - c.x1), h: Math.abs(c.y2 - c.y1) } : c;
        expect(box.x, id).toBeGreaterThanOrEqual(-0.5);
        expect(box.y, id).toBeGreaterThanOrEqual(-0.5);
        expect(box.x + box.w, id).toBeLessThanOrEqual(500.5);
        expect(box.y + box.h, id).toBeLessThanOrEqual(300.5);
      }
      const sizes = new Set(out.filter((c) => c.text && c.type === 'shape').map((c) => c.style.size));
      expect(sizes.size, id).toBeLessThanOrEqual(1);
    }
  });

  it('shrinks text so the longest label fits its shape, without breaking words', () => {
    expect(textFits([{ text: 'Strengths' }], 100, 40, 12)).toBe(true);
    expect(textFits([{ text: 'Strengths' }], 30, 400, 12)).toBe(false);
    const short = fitFontSize([{ w: 160, h: 90, paras: [{ text: 'Plan' }] }]);
    const long = fitFontSize([{ w: 160, h: 90, paras: [{ text: 'Internationalisation' }] }]);
    expect(long).toBeLessThan(short);
    const a = layoutDiagram(diagram('process', nodes(['a', 0], ['b', 0], ['c', 0], ['d', 0])));
    const b = layoutDiagram(diagram('process', nodes(['a', 0], ['Internationalisation and localisation', 0], ['c', 0], ['d', 0])));
    expect(b[0].style.size).toBeLessThan(a[0].style.size);
    expect(b[0].style.size).toBe(b[2].style.size);
  });

  it('colours shapes by the diagram colours and style', () => {
    const colorful = layoutDiagram(diagram('matrix', nodes(['a', 0], ['b', 0], ['c', 0]), { colors: 'colorful' }));
    expect(colorful.map((e) => e.fill)).toEqual(['theme:accent1', 'theme:accent2', 'theme:accent3']);
    const outline = layoutDiagram(diagram('matrix', nodes(['a', 0]), { colors: 'accent3', style: 'outline' }));
    expect(outline[0]).toMatchObject({ fill: 'none', stroke: { color: 'theme:accent3' }, style: { color: 'theme:text' } });
    const imported = layoutDiagram(diagram('matrix', nodes(['a', 0]), { palette: ['#111111', '#222222', '#333333', '#444444', '#555555', '#666666'], colors: 'accent2', font: 'Calibri' }));
    expect(imported[0]).toMatchObject({ fill: '#222222', style: { font: 'Calibri' } });
  });

  it('shows PowerPoint’s drawing, stretched to the box, until edited', () => {
    const el = diagram('list', nodes(['a', 0]), { drawn: { w: 300, h: 160, elements: [{ id: 'c', type: 'shape', shape: 'ellipse', x: 10, y: 20, w: 100, h: 50, text: '' }] } });
    expect(diagramChildren(el)).toEqual([expect.objectContaining({ shape: 'ellipse', x: 20, y: 40, w: 200, h: 100 })]);
    const node = renderElement(el, { theme: themeById('simple-light'), deck: newDeck() });
    expect(node.querySelectorAll('.el-diagram-shapes > .el')).toHaveLength(1);
    delete el.diagram.drawn;
    expect(diagramChildren(el)[0].shape).toBe('rect');
  });

  it('turns into grouped shapes on the slide, keeping its animation', () => {
    const el = { ...diagram('process', nodes(['a', 0], ['b', 0])), anim: { effect: 'fade', duration: 0.5, start: 'click', order: 1 } };
    const shapes = diagramAsShapes(el);
    expect(shapes).toHaveLength(3);
    expect(new Set(shapes.map((s) => s.groupId)).size).toBe(1);
    expect(shapes[0].x).toBeGreaterThanOrEqual(100);
    expect(shapes.map((s) => s.anim.start)).toEqual(['click', 'with', 'with']);
  });

  it('counts its text in the slide’s text and its pictures as used', () => {
    const el = diagram('list', nodes(['Hello', 0]), { drawn: { w: 1, h: 1, elements: [{ id: 'p', type: 'image', media: 'm1', x: 0, y: 0, w: 1, h: 1 }] }, source: { data: 'x1', layout: 'x2', media: { data: [{ id: 'rId1', media: 'm2' }] } } });
    expect(slideText({ elements: [el] })).toBe('Hello');
    expect(mediaIdsOf(el)).toEqual(['m1', 'x1', 'x2', 'm2']);
  });
});

const htmlText = (html) => new DOMParser().parseFromString(html, 'text/html').body.textContent;

// ------------------------------------------------------------------ PowerPoint

const NS = 'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const pt = (id, text, type = null) => `<dgm:pt modelId="${id}"${type ? ` type="${type}"` : ''}><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p>${text ? `<a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r>` : ''}</a:p></dgm:t></dgm:pt>`;
const cxn = (src, dest, ord, type = null) => `<dgm:cxn modelId="c${src}${dest}"${type ? ` type="${type}"` : ''} srcId="${src}" destId="${dest}" srcOrd="${ord}" destOrd="0"/>`;
// A hand-made data model: the doc point "0", "Second" (2) listed before "First" (1), "Child" (3) under "First", and a presentation point.
const DATA = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<dgm:dataModel ${NS}><dgm:ptLst>${pt('0', '', 'doc')}${pt('2', 'Second')}${pt('1', 'First')}${pt('3', 'Child')}${pt('9', '', 'pres')}</dgm:ptLst>`
  + `<dgm:cxnLst>${cxn('0', '2', 1)}${cxn('0', '1', 0)}${cxn('1', '3', 0)}${cxn('0', '9', 0, 'presOf')}</dgm:cxnLst><dgm:bg/><dgm:whole/>`
  + '<dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rIdDr"/></a:ext></dgm:extLst></dgm:dataModel>';
const LAYOUT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<dgm:layoutDef ${NS} uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/process1"><dgm:title val=""/><dgm:layoutNode name="diagram"/></dgm:layoutDef>`;
const STYLE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<dgm:styleDef ${NS} uniqueId="urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1"><dgm:title val=""/></dgm:styleDef>`;
const COLORS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<dgm:colorsDef ${NS} uniqueId="urn:microsoft.com/office/officeart/2005/8/colors/colorful1"><dgm:title val=""/></dgm:colorsDef>`;
const DRAWING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" ${NS}><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>`
  + '<dsp:sp modelId="{1}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></dsp:spPr>'
  + '<dsp:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1800"/><a:t>First</a:t></a:r></a:p></dsp:txBody><dsp:txXfrm><a:off x="914400" y="0"/><a:ext cx="1828800" cy="914400"/></dsp:txXfrm></dsp:sp></dsp:spTree></dsp:drawing>';

/** A deck exported and given a SmartArt diagram in a frame at (1 in, 1 in), 4 × 2 in, read back. */
async function withSmartArt({ drawing = DRAWING, parts = true } = {}) {
  const zip = await JSZip.loadAsync(await toBuffer(await buildPptx(deckWith())));
  const frame = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Diagram"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rIdDm" r:lo="rIdLo" r:qs="rIdQs" r:cs="rIdCs"/></a:graphicData></a:graphic></p:graphicFrame>`;
  const slide = await zip.file('ppt/slides/slide1.xml').async('string');
  zip.file('ppt/slides/slide1.xml', slide.replace('</p:spTree>', `${frame}</p:spTree>`));
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const rel = (id, type, target) => `<Relationship Id="${id}" Type="${type}" Target="../diagrams/${target}"/>`;
  const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('string');
  zip.file('ppt/slides/_rels/slide1.xml.rels', rels.replace('</Relationships>', `${rel('rIdDm', `${REL}/diagramData`, 'data1.xml')}${parts ? rel('rIdLo', `${REL}/diagramLayout`, 'layout1.xml') + rel('rIdQs', `${REL}/diagramQuickStyle`, 'quickStyle1.xml') + rel('rIdCs', `${REL}/diagramColors`, 'colors1.xml') : ''}${drawing ? rel('rIdDr', 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing', 'drawing1.xml') : ''}</Relationships>`));
  zip.file('ppt/diagrams/data1.xml', DATA);
  if (parts) {
    zip.file('ppt/diagrams/layout1.xml', LAYOUT);
    zip.file('ppt/diagrams/quickStyle1.xml', STYLE);
    zip.file('ppt/diagrams/colors1.xml', COLORS);
  }
  if (drawing) zip.file('ppt/diagrams/drawing1.xml', drawing);
  return importPptx(await zip.generateAsync({ type: 'arraybuffer' }), 'x');
}

describe('SmartArt from PowerPoint', () => {
  it('knows PowerPoint’s layouts and colours by the closest of its own', () => {
    const u = (n) => `urn:microsoft.com/office/officeart/2005/8/layout/${n}`;
    expect(['default', 'vList2', 'vList5', 'process1', 'process2', 'chevron1', 'hChevron3', 'cycle1', 'cycle2', 'hierarchy1', 'orgChart1', 'pyramid1', 'venn1', 'matrix1', 'radial1', 'pList2'].map((n) => layoutKind(u(n))))
      .toEqual(['list', 'vlist', 'vlist', 'process', 'process', 'chevron', 'chevron', 'cycle', 'cycle', 'hierarchy', 'hierarchy', 'pyramid', 'venn', 'matrix', 'radial', 'list']);
    expect(['accent1_2', 'accent4_3', 'colorful1', 'dark2'].map((n) => colorsKind(`urn:microsoft.com/office/officeart/2005/8/colors/${n}`))).toEqual(['accent1', 'accent4', 'colorful', 'accent1']);
  });

  it('reads the outline from the data model, children in their order', () => {
    const doc = new DOMParser().parseFromString(DATA, 'application/xml');
    expect(readOutline(doc)).toEqual([{ id: '1', text: 'First', level: 0 }, { id: '3', text: 'Child', level: 1 }, { id: '2', text: 'Second', level: 0 }]);
  });

  it('brings a diagram in looking as PowerPoint drew it, with its outline, layout and colours', async () => {
    const deck = await withSmartArt();
    const [el] = deck.slides[0].elements;
    expect(el).toMatchObject({ type: 'diagram', x: 72, y: 72, w: 288, h: 144 });
    expect(el.diagram).toMatchObject({ layout: 'process', colors: 'colorful', style: 'flat' });
    expect(el.diagram.nodes.map((n) => [n.text, n.level])).toEqual([['First', 0], ['Child', 1], ['Second', 0]]);
    // Its SmartArt parts are kept in deck.media, out of the slides (and their undo history).
    const { source } = el.diagram;
    expect(source.look).toEqual({ layout: 'process', colors: 'colorful', style: 'flat' });
    expect(Object.fromEntries(['data', 'layout', 'style', 'colors', 'drawing'].map((key) => [key, xmlOfDataUrl(deck.media[source[key]])])))
      .toEqual({ data: DATA, layout: LAYOUT, style: STYLE, colors: COLORS, drawing: DRAWING });
    // The drawing's shapes, in the diagram's box.
    const [circle, label] = el.diagram.drawn.elements;
    expect(el.diagram.drawn).toMatchObject({ w: 288, h: 144 });
    expect(circle).toMatchObject({ type: 'shape', shape: 'ellipse', fill: '#4472c4', x: 0, y: 0, w: 72, h: 72 });
    expect(label).toMatchObject({ type: 'text', x: 72, y: 0, w: 144, h: 72 });
  });

  it('lays out a diagram without a drawing itself, and one without its parts too', async () => {
    const [el] = (await withSmartArt({ drawing: null })).slides[0].elements;
    expect(el.diagram.drawn).toBeUndefined();
    expect(diagramChildren(el).filter((c) => c.shape === 'roundRect')).toHaveLength(2);
    const [bare] = (await withSmartArt({ drawing: null, parts: false })).slides[0].elements;
    expect(bare.diagram).toMatchObject({ layout: 'list', nodes: [{ text: 'First' }, { text: 'Child' }, { text: 'Second' }] });
    expect(bare.diagram.source).toBeUndefined();
  });

  it('goes back as the same SmartArt, its parts unchanged, until it’s edited', async () => {
    const deck = await withSmartArt();
    const zip = await JSZip.loadAsync(await toBuffer(await buildPptx(deck)));
    expect(await zip.file('ppt/diagrams/data1.xml').async('string')).toBe(DATA);
    expect(await zip.file('ppt/diagrams/layout1.xml').async('string')).toBe(LAYOUT);
    expect(await zip.file('ppt/diagrams/quickStyle1.xml').async('string')).toBe(STYLE);
    expect(await zip.file('ppt/diagrams/colors1.xml').async('string')).toBe(COLORS);
    expect(await zip.file('ppt/diagrams/drawing1.xml').async('string')).toBe(DRAWING);
    const slide = await zip.file('ppt/slides/slide1.xml').async('string');
    expect(slide).toMatch(/<p:graphicFrame>.*<a:graphicData uri="http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/diagram"><dgm:relIds [^>]*r:dm="rIdLsDm1" r:lo="rIdLsLo1" r:qs="rIdLsQs1" r:cs="rIdLsCs1"\/>/);
    expect(slide).toContain('<a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/>');
    const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('string');
    expect(rels).toContain('Id="rIdDr" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing1.xml"');
    const types = await zip.file('[Content_Types].xml').async('string');
    for (const t of ['drawingml.diagramData+xml', 'drawingml.diagramLayout+xml', 'drawingml.diagramStyle+xml', 'drawingml.diagramColors+xml', 'ms-office.drawingml.diagramDrawing+xml']) expect(types).toContain(t);
    // And reads back the same.
    const back = (await importPptx(await zip.generateAsync({ type: 'arraybuffer' }))).slides[0].elements[0];
    expect(back.diagram.nodes.map((n) => n.text)).toEqual(['First', 'Child', 'Second']);
    expect(back.diagram.drawn.elements).toHaveLength(2);
  });

  it('writes an edited diagram’s outline into its data, and its own layout as the drawing', async () => {
    const deck = await withSmartArt();
    const d = deck.slides[0].elements[0].diagram;
    d.nodes = [{ ...d.nodes[0], text: 'First, edited' }, d.nodes[2], { id: 'new', text: 'Third & last', level: 1 }];
    delete d.drawn;
    d.edited = true;
    const zip = await JSZip.loadAsync(await toBuffer(await buildPptx(deck)));
    const data = await zip.file('ppt/diagrams/data1.xml').async('string');
    expect(data).toContain('First, edited');
    expect(data).toContain('Third &amp; last');
    expect(data).not.toContain('Child');
    expect(data).not.toContain('type="pres"');
    expect(data).toContain('modelId="1"');
    expect(await zip.file('ppt/diagrams/layout1.xml').async('string')).toBe(LAYOUT);
    const drawing = await zip.file('ppt/diagrams/drawing1.xml').async('string');
    expect(drawing).toMatch(/^<\?xml[^>]*\?>\n<dsp:drawing /);
    expect(drawing).toContain('<a:t>Third &amp; last</a:t>');
    expect(data).toMatch(new RegExp(`relId="${/Id="([^"]+)" Type="[^"]*diagramDrawing"/.exec(await zip.file('ppt/slides/_rels/slide1.xml.rels').async('string'))[1]}"`));
    const back = (await importPptx(await zip.generateAsync({ type: 'arraybuffer' }))).slides[0].elements[0];
    expect(back.diagram.nodes.map((n) => [n.text, n.level])).toEqual([['First, edited', 0], ['Second', 0], ['Third & last', 1]]);
    expect(back.diagram.drawn.elements.some((e) => /Third &amp; last/.test(e.text))).toBe(true);
  });

  it('edits a data model: kept nodes keep their points, new ones get their own, connections made afresh', () => {
    const xml = editedDataXml(DATA, [{ id: '2', text: 'Second', level: 0 }, { id: 'x', text: 'New', level: 1 }], 'rIdNew');
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    expect(readOutline(doc).map((n) => [n.text, n.level])).toEqual([['Second', 0], ['New', 1]]);
    expect(readOutline(doc)[0].id).toBe('2');
    expect(xml).toContain('relId="rIdNew"');
    expect(doc.getElementsByTagName('parsererror')).toHaveLength(0);
  });

  it('saves a diagram given a layout of LibreSlides’ own, or a new one, as grouped shapes', async () => {
    const deck = await withSmartArt();
    Object.assign(deck.slides[0].elements[0].diagram, { layout: 'cycle', edited: true });
    delete deck.slides[0].elements[0].diagram.drawn;
    deck.slides[0].elements.push(newDiagram('process', 400, 300, 300, 150));
    const zip = await JSZip.loadAsync(await toBuffer(await buildPptx(deck)));
    expect(Object.keys(zip.files).filter((n) => n.startsWith('ppt/diagrams/'))).toEqual([]);
    const slide = await zip.file('ppt/slides/slide1.xml').async('string');
    expect(slide.match(/<p:grpSp>/g)).toHaveLength(2);
    expect(slide).not.toContain('graphicFrame');
  });

  it('stretches PowerPoint’s drawing with a resized diagram', async () => {
    expect(stretchDrawingXml('<a:off x="10" y="20"/><a:ext cx="100" cy="50"/><a:ext uri="x"/>', 2, 0.5)).toBe('<a:off x="20" y="10"/><a:ext cx="200" cy="25"/><a:ext uri="x"/>');
    const deck = await withSmartArt();
    deck.slides[0].elements[0].w *= 2;
    const zip = await JSZip.loadAsync(await toBuffer(await buildPptx(deck)));
    expect(await zip.file('ppt/diagrams/drawing1.xml').async('string')).toContain('<a:off x="1828800" y="0"/><a:ext cx="3657600" cy="914400"/>');
  });
});

describe('untrusted diagrams', () => {
  it('keeps known layouts, colours and styles, and plain-text nodes of sensible size and number', () => {
    const d = sanitizeDiagram({
      layout: 'spiral', colors: 'url(x)', style: 3, font: 'Calibri"; x:url(y)', palette: ['#fff'],
      nodes: [{ id: 'a', text: 'One\u0000two', level: 3 }, { id: 'a', text: 7, level: 'x' }, { text: 'x'.repeat(5000), level: 1.5 }, null],
    });
    expect(d).toMatchObject({ layout: 'list', colors: 'accent1', style: 'flat', font: 'Calibri xurly' });
    expect(d.palette).toBeUndefined();
    expect(d.nodes.map((n) => [n.text.length, n.level])).toEqual([[7, 0], [1, 0], [DIAGRAM_LIMITS.text, 0]]);
    expect(d.nodes[0].text).toBe('One two');
    expect(new Set(d.nodes.map((n) => n.id)).size).toBe(3);
    expect(sanitizeDiagram({ nodes: Array.from({ length: 2000 }, () => ({ text: 'x', level: 0 })) }).nodes).toHaveLength(DIAGRAM_LIMITS.nodes);
    expect(sanitizeDiagram(null)).toEqual({ layout: 'list', nodes: [], colors: 'accent1', style: 'flat' });
  });

  it('cleans PowerPoint’s drawing like any other elements', () => {
    const d = sanitizeDiagram({ drawn: { w: 10, h: 10, elements: [{ type: 'text', text: '<img src=x onerror=alert(1)><b>ok</b>', fill: 'url(x)' }, { type: 'diagram', diagram: {} }, { type: 'shape', shape: 'custom', paths: 'x' }] } });
    expect(d.drawn.elements).toHaveLength(2);
    expect(d.drawn.elements[0]).toMatchObject({ text: '<b>ok</b>', fill: 'none' });
    expect(sanitizeDiagram({ drawn: { w: -1, h: 10, elements: [{ type: 'text' }] } }).drawn).toBeUndefined();
  });

  it('keeps SmartArt parts only as XML that parses, and a diagram’s only when it has them all', () => {
    expect(isXmlPart(xmlDataUrl(DATA))).toBe(true);
    expect(isXmlPart(xmlDataUrl('<dgm:layoutDef'))).toBe(false);
    expect(isXmlPart(xmlDataUrl('not xml'))).toBe(false);
    expect(isXmlPart(xmlDataUrl(`<a>${'x'.repeat(DIAGRAM_LIMITS.part)}</a>`))).toBe(false);
    expect(isXmlPart('data:application/xml;base64,!!!')).toBe(false);
    const parts = { d: xmlDataUrl(DATA), l: xmlDataUrl(LAYOUT), s: xmlDataUrl(STYLE), c: xmlDataUrl(COLORS), bad: xmlDataUrl('<a><b></a>'), raw: 'data:application/xml,<a/>' };
    const deck = sanitizeDeck({
      media: parts,
      slides: [{ elements: [
        { type: 'diagram', diagram: { layout: 'process', source: { data: 'd', layout: 'l', style: 's', colors: 'c', drawing: 'bad', look: { layout: 'cycle', colors: 'x' } } } },
        { type: 'diagram', diagram: { source: { data: 'd', layout: 'bad', style: 's', colors: 'c' } } },
        { type: 'diagram', diagram: { source: { data: 'd', layout: '../x', style: 's', colors: 'c' } } },
        { type: 'diagram', diagram: { source: { data: DATA, layout: LAYOUT, style: STYLE, colors: COLORS } } },
      ] }],
    });
    expect(Object.keys(deck.media).sort()).toEqual(['c', 'd', 'l', 's']);
    const [a, b, c, raw] = deck.slides[0].elements.map((e) => e.diagram.source);
    expect(a).toEqual({ data: 'd', layout: 'l', style: 's', colors: 'c', look: { layout: 'cycle', colors: 'accent1', style: 'flat' } });
    expect([b, c, raw]).toEqual([undefined, undefined, undefined]);
  }, 30000); // encodes a 4 MB part

  it('drops a diagram’s SmartArt parts when the pictures they need are missing, and pictures missing from its drawing', () => {
    const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
    const parts = { data: 'd', layout: 'l', style: 's', colors: 'c' };
    const deck = sanitizeDeck({
      media: { m1: PNG, d: xmlDataUrl(DATA), l: xmlDataUrl(LAYOUT), s: xmlDataUrl(STYLE), c: xmlDataUrl(COLORS) },
      slides: [{ elements: [
        { type: 'diagram', diagram: { source: { ...parts, media: { data: [{ id: 'rId1', media: 'm1' }] } }, drawn: { w: 1, h: 1, elements: [{ type: 'image', media: 'gone', x: 0, y: 0, w: 1, h: 1 }, { type: 'shape', x: 0, y: 0, w: 1, h: 1 }] } } },
        { type: 'diagram', diagram: { source: { ...parts, media: { data: [{ id: 'rId1', media: 'gone' }] } } } },
      ] }],
    });
    const [a, b] = deck.slides[0].elements;
    expect(a.diagram.source.media).toEqual({ data: [{ id: 'rId1', media: 'm1' }] });
    expect(a.diagram.drawn.elements.map((e) => e.type)).toEqual(['shape']);
    expect(b.diagram.source).toBeUndefined();
  });
});
