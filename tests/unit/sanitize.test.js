import { describe, it, expect } from 'vitest';
import { sanitizeHtml, sanitizeDeck } from '../../src/io/sanitize.js';
import { importFile } from '../../src/io/files.js';

const lslides = (data) => new File([JSON.stringify(data)], 'talk.lslides');

describe('untrusted presentations', () => {
  it('drops inline styles that could load something from the network', () => {
    expect(sanitizeHtml('<span style="color: red">a</span>')).toBe('<span style="color: red">a</span>');
    for (const style of ['background: url(https://t/x)', "background-image: image-set('https://t/x' 1x)", 'background: \\75 rl(https://t/x)']) {
      expect(sanitizeHtml(`<span style="${style}">a</span>`)).toBe('<span>a</span>');
    }
  });

  it('only lets a number through as a gradient angle', () => {
    const deck = sanitizeDeck({ slides: [{ background: { type: 'gradient', from: '#000', to: '#fff', angle: '0deg,#fff,#fff),url(https://t/x),linear-gradient(0' } }, { background: { type: 'gradient', from: '#000', to: '#fff', angle: 45 } }] });
    expect(deck.slides[0].background.angle).toBe(180);
    expect(deck.slides[1].background.angle).toBe(45);
  });

  it('keeps font names in any script', () => {
    const deck = sanitizeDeck({ slides: [{ elements: [{ type: 'text', style: { font: '맑은 고딕' } }, { type: 'text', style: { font: 'Inter"; x:url(y)' } }] }] });
    expect(deck.slides[0].elements[0].style.font).toBe('맑은 고딕');
    expect(deck.slides[0].elements[1].style.font).toBe('Inter xurly');
  });

  it('keeps only safe table cell borders and sensible margins', () => {
    const deck = sanitizeDeck({ slides: [{ elements: [{ type: 'table', cells: [[
      { text: '', borders: { t: { color: 'url(https://t/x)', width: 2 }, b: { color: '#f00', width: 999, dash: 'wavy' }, l: 'x' }, margin: { t: -4, l: 'x', r: 9 } },
      { text: '', borders: 'x', margin: 7 },
    ]] }] }] });
    const [a, b] = deck.slides[0].elements[0].cells[0];
    expect(a.borders).toEqual({ b: { color: '#f00', width: 50, dash: 'solid' } });
    expect(a.margin).toEqual({ t: 0, r: 9, b: 3.6, l: 7.2 });
    expect(b.borders).toBeUndefined();
    expect(b.margin).toBeUndefined();
  });

  it('copes with damaged structure instead of crashing', () => {
    const deck = sanitizeDeck({ title: 7, media: 'abc', slides: [null, { elements: [null, { type: 'table', cells: [null, [null, { text: '<b>x</b>' }]] }, { type: 'image', media: 'm', crop: { l: 'x', t: 2 } }] }] });
    expect(deck.title).toBe('');
    expect(deck.media).toEqual({});
    expect(deck.slides).toHaveLength(1);
    const [table] = deck.slides[0].elements;
    expect(table.cells).toEqual([[{ text: '' }, { text: '<b>x</b>' }]]);
  });

  it('keeps only well-formed freeform outlines', () => {
    const deck = sanitizeDeck({ slides: [{ elements: [
      { type: 'shape', shape: 'custom', paths: [{ cmds: [['M', 0, 0], ['L', 1, 'x'], ['X', 1], ['C', 1, 2, 3], ['Q', 0.5, 0, 1, 1], ['Z']], fill: false, stroke: 'yes' }, 'junk'] },
      { type: 'shape', shape: 'custom', paths: [{ cmds: [['L', NaN, 0]] }] },
    ] }] });
    const [good, bad] = deck.slides[0].elements;
    expect(good.paths).toEqual([{ cmds: [['M', 0, 0], ['Q', 0.5, 0, 1, 1], ['Z']], fill: false }]);
    expect(bad.shape).toBe('rect');
    expect(bad.paths).toBeUndefined();
  });

  it('refuses .lslides files from a newer version or another app', async () => {
    const deck = { title: 'T', slides: [] };
    await expect(importFile(lslides({ app: 'LibreSlides', format: 99, deck }))).rejects.toThrow(/newer version/);
    await expect(importFile(lslides({ app: 'Other', format: 1, deck }))).rejects.toThrow(/isn’t a LibreSlides/);
    const ok = await importFile(lslides({ app: 'LibreSlides', format: 1, deck }));
    expect(ok.title).toBe('T');
  });

  it('keeps section names as plain text of sensible length', () => {
    const deck = sanitizeDeck({ slides: [{ section: '  Intro\n<b>x</b> ' }, { section: 7 }, { section: '   ' }, { section: 'x'.repeat(1000) }] });
    expect(deck.slides.map((s) => s.section?.length ?? s.section)).toEqual([14, undefined, undefined, 255]);
    expect(deck.slides[0].section).toBe('Intro <b>x</b>');
  });

  it('keeps numbers numbers, and names and choices known ones', () => {
    const deck = sanitizeDeck({
      size: '__proto__',
      slides: [{
        id: '"><x', transition: { type: 'constructor', duration: 'slow' }, hidden: 'yes',
        elements: [
          { type: 'shape', shape: { a: 1 }, x: '10', y: -1e308, w: -5, rot: NaN, opacity: 7, text: { a: 1 }, stroke: { color: 'red', width: 'x', dash: 'constructor' }, style: { align: 3, valign: [], size: 'big', lineHeight: 1e9 } },
          { type: 'line', x1: 0, y1: null, x2: '5', stroke: 'thin' },
          { type: 'table', cells: '', colWidths: [1, 3] },
          { type: 'table', cells: [[{ text: 'a' }, { text: 'b' }]], colWidths: [1, 3] },
        ],
        comments: [{ text: 't', x: -1e308, y: 2 }],
      }],
    });
    expect(deck.size).toBe('16:9');
    const [slide] = deck.slides;
    expect(slide.id).toMatch(/^[\w-]+$/);
    expect(slide.transition).toEqual({ type: 'none', duration: 0.6 });
    expect(slide.hidden).toBe(false);
    const [shape, line, empty, table] = slide.elements;
    expect(shape).toMatchObject({ shape: 'rect', x: 0, y: -1e5, w: 0, h: 0, rot: 0, opacity: 1, text: '', stroke: { color: 'none', dash: 'solid' } });
    expect(shape.stroke.width).toBeUndefined();
    expect(shape.style).toEqual({ lineHeight: 10 });
    expect(line).toMatchObject({ x1: 0, y1: 0, x2: 0, y2: 0, stroke: null });
    expect(empty.cells).toEqual([[{ text: '' }]]);
    expect(table.colWidths).toEqual([0.25, 0.75]);
    expect(slide.comments[0]).toMatchObject({ x: -1e5, y: 2 });
  });
});
