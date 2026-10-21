import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { buildPptx, htmlToRuns, toHex, fixParagraphs } from '../../src/io/pptx-export.js';
import { importPptx } from '../../src/io/pptx-import.js';
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
