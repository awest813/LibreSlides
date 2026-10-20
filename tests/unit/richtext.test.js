import { describe, it, expect } from 'vitest';
import { stripInline, tidyHtml, isEmptyHtml, normalizeFontTags } from '../../src/editor/richtext.js';
import { sanitizeHtml, sanitizeDeck } from '../../src/io/sanitize.js';

describe('rich text', () => {
  it('removes inline formatting that a whole-box format replaces', () => {
    expect(stripInline('<p><b>a</b> <span style="color: red; font-size: 20px">b</span></p>', 'bold')).toBe('<p>a <span style="color: red; font-size: 20px">b</span></p>');
    expect(stripInline('<p><span style="color: red">b</span></p>', 'color')).toBe('<p>b</p>');
    expect(stripInline('<p style="text-align: center">x</p>', 'align')).toBe('<p>x</p>');
  });

  it('tidies edited HTML', () => {
    expect(tidyHtml('hello')).toBe('<p>hello</p>');
    expect(tidyHtml('<div>a</div><div>b</div>')).toBe('<p>a</p><p>b</p>');
    expect(tidyHtml('<p><br></p>')).toBe('');
    expect(isEmptyHtml('<ul><li><br></li></ul>')).toBe(true);
  });

  it('turns size markers into point sizes', () => {
    const div = document.createElement('div');
    div.innerHTML = '<p>a <font size="7">big <span style="font-size: 10px">x</span></font></p>';
    normalizeFontTags(div);
    expect(div.querySelector('font')).not.toBeNull(); // left for the caller that is setting the size
    const spans = normalizeFontTags(div, 40);
    expect(spans).toHaveLength(1);
    expect(div.querySelector('span').style.fontSize).toBe('40px');
    expect(div.querySelector('span span').style.fontSize).toBe('');
  });

  it('sanitizes untrusted text and pictures', () => {
    expect(sanitizeHtml('<p onclick="x()">a<script>alert(1)</script><img src=x onerror=y></p>')).toBe('<p>a</p>');
    expect(sanitizeHtml('<span style="background: url(http://evil)">a</span>')).toBe('<span>a</span>');
    const deck = { slides: [{ elements: [{ text: '<b>ok</b><iframe></iframe>' }] }], media: { a: 'javascript:alert(1)', b: 'data:image/png;base64,AAAA' } };
    sanitizeDeck(deck);
    expect(deck.slides[0].elements[0].text).toBe('<b>ok</b>');
    expect(Object.keys(deck.media)).toEqual(['b']);
  });
});
