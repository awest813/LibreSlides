import { describe, it, expect } from 'vitest';
import {
  newDeck, newSlide, layoutElements, cloneSlide, fillPlaceholders, addMedia, pruneMedia, normalizeDeck, slideText, slideTitle,
  textToHtml, htmlToText, lineBox, moveElement, newLine, LAYOUTS, SIZES,
} from '../../src/model/deck.js';
import { TEMPLATES } from '../../src/model/templates.js';
import { THEMES, resolveColor, backgroundCss, themeById } from '../../src/model/themes.js';
import { SHAPES, shapePath, customPath } from '../../src/model/shapes.js';

describe('deck model', () => {
  it('creates a deck with a title slide', () => {
    const d = newDeck();
    expect(d.slides).toHaveLength(1);
    expect(d.slides[0].layout).toBe('title');
    expect(d.slides[0].elements.map((e) => e.ph)).toEqual(['title', 'subtitle']);
  });

  it('lays out every layout inside the slide', () => {
    for (const size of Object.values(SIZES)) {
      for (const l of LAYOUTS) {
        for (const el of layoutElements(l.id, size.w, size.h)) {
          expect(el.x).toBeGreaterThanOrEqual(0);
          expect(el.y).toBeGreaterThanOrEqual(0);
          expect(el.x + el.w).toBeLessThanOrEqual(size.w);
          expect(el.y + el.h).toBeLessThanOrEqual(size.h);
        }
      }
    }
  });

  it('copies slides with new ids', () => {
    const s = newSlide('twoContent');
    const c = cloneSlide(s);
    expect(c.id).not.toBe(s.id);
    expect(c.elements.map((e) => e.id)).not.toEqual(s.elements.map((e) => e.id));
    expect(c.elements.map((e) => e.ph)).toEqual(s.elements.map((e) => e.ph));
  });

  it('fills placeholders with text and bullets', () => {
    const s = fillPlaceholders(newSlide('titleContent'), { title: 'Hello', body: ['One', 'Two & <three>'] });
    expect(slideTitle(s)).toBe('Hello');
    expect(s.elements[1].text).toBe('<ul><li>One</li><li>Two &amp; &lt;three&gt;</li></ul>');
    expect(slideText(s)).toBe('Hello One Two & <three>');
  });

  it('converts between text and HTML', () => {
    expect(textToHtml('a\n\nb')).toBe('<p>a</p><p><br></p><p>b</p>');
    expect(htmlToText('<p>a<br>b</p><ul><li>c</li></ul>')).toBe('a\nb\nc');
  });

  it('keeps one copy of each picture and prunes unused ones', () => {
    const d = newDeck();
    const a = addMedia(d, 'data:image/png;base64,AAAA');
    expect(addMedia(d, 'data:image/png;base64,AAAA')).toBe(a);
    const b = addMedia(d, 'data:image/png;base64,BBBB');
    d.slides[0].elements.push({ id: 'i', type: 'image', media: a, x: 0, y: 0, w: 1, h: 1 });
    pruneMedia(d);
    expect(Object.keys(d.media)).toEqual([a]);
    expect(d.media[b]).toBeUndefined();
  });

  it('moves lines by their end points', () => {
    const l = newLine(10, 20, 110, 70);
    moveElement(l, 5, -5);
    expect(lineBox(l)).toEqual({ x: 15, y: 15, w: 100, h: 50 });
  });

  it('repairs decks from older or foreign files', () => {
    const d = normalizeDeck({ slides: [{ elements: [{ type: 'text', x: 0, y: 0, w: 10, h: 10 }] }], size: 'weird', themeId: 'nope' });
    expect(d.size).toBe('16:9');
    expect(d.themeId).toBe(THEMES[0].id);
    expect(d.slides[0].id).toBeTruthy();
    expect(d.slides[0].elements[0].style).toEqual({});
  });

  it('builds every template', () => {
    for (const t of TEMPLATES) {
      const d = t.build();
      expect(d.slides.length).toBeGreaterThan(0);
      expect(themeById(d.themeId).id).toBe(d.themeId);
    }
  });
});

describe('themes and shapes', () => {
  it('resolves theme colours', () => {
    const t = themeById('coral');
    expect(resolveColor('theme:accent1', t)).toBe(t.colors.accent1);
    expect(resolveColor('#123456', t)).toBe('#123456');
    expect(resolveColor('none', t)).toBe('transparent');
    expect(resolveColor(null, t, 'x')).toBe('x');
    expect(backgroundCss({ type: 'gradient', from: '#000000', to: 'theme:accent1', angle: 90 }, t)).toBe(`linear-gradient(90deg, #000000, ${t.colors.accent1})`);
  });

  it('draws a closed path for every shape', () => {
    for (const s of SHAPES) {
      const d = shapePath(s.id, 120, 80);
      expect(d).toMatch(/^M/);
      expect(d).toMatch(/Z$/);
      expect(d).not.toMatch(/NaN/);
    }
  });

  it('stretches freeform outlines to the box', () => {
    expect(customPath([['M', 0, 1], ['L', 0.5, 0], ['C', 0, 0, 1, 1, 1, 0.5], ['Q', 0.25, 0.25, 0, 0], ['Z']], 200, 100))
      .toBe('M0 100L100 0C0 0 200 100 200 50Q50 25 0 0Z');
  });
});
