import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { buildPptx, htmlToRuns, toHex } from '../../src/io/pptx-export.js';
import { importPptx } from '../../src/io/pptx-import.js';
import { TEMPLATES } from '../../src/model/templates.js';
import { newLine, newTable, htmlToText } from '../../src/model/deck.js';

const toBuffer = (blob) => (blob.arrayBuffer ? blob.arrayBuffer() : new Response(blob).arrayBuffer());

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
