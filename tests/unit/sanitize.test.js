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
});
