import { describe, it, expect } from 'vitest';
import {
  newDeck, newSlide, layoutElements, cloneSlide, fillPlaceholders, addMedia, pruneMedia, normalizeDeck, slideText, slideTitle,
  textToHtml, htmlToText, lineBox, moveElement, newLine, LAYOUTS, SIZES, slideSize,
  slideSections, moveSlidesWithSections, removeSlidesWithSections, setSections,
} from '../../src/model/deck.js';
import { TEMPLATES } from '../../src/model/templates.js';
import { THEMES, resolveColor, backgroundCss, themeById } from '../../src/model/themes.js';
import { SHAPES, shapePath, customPath } from '../../src/model/shapes.js';

describe('deck model', () => {
  it('takes only real size names as sizes', () => {
    expect(slideSize({ size: '__proto__' })).toEqual(SIZES['16:9']);
    expect(slideSize({ size: 'toString' })).toEqual(SIZES['16:9']);
    expect(slideSize({ size: 'a4' })).toEqual(SIZES.a4);
  });

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

describe('sections', () => {
  /** Slides a…, with "A:a" naming a section that starts at slide a. */
  const slides = (spec) => spec.split(' ').map((t) => {
    const [name, id] = t.includes(':') ? t.split(':') : [null, t];
    return { id, ...(name && { section: name }) };
  });
  const show = (list) => list.map((s) => (s.section ? `${s.section}:${s.id}` : s.id)).join(' ');
  const move = (spec, ids, index) => {
    const list = slides(spec);
    const m = moveSlidesWithSections(list, new Set(ids), index);
    if (!m) return null;
    setSections(m.order, m.sections);
    return show(m.order);
  };

  it('groups slides by the section their first slide names', () => {
    expect(slideSections(slides('a A:b c B:d'))).toEqual([
      { name: null, start: 0, end: 1 }, { name: 'A', start: 1, end: 3 }, { name: 'B', start: 3, end: 4 },
    ]);
    expect(slideSections(slides('A:a b'))).toEqual([{ name: 'A', start: 0, end: 2 }]);
  });

  it('doesn’t copy a section with its first slide', () => {
    expect(cloneSlide({ id: 'a', section: 'A', elements: [] }).section).toBeUndefined();
  });

  it('passes a section on to its next slide when its first slide is deleted or moved away', () => {
    expect(show(removeSlidesWithSections(slides('A:a b c B:d'), new Set(['a'])))).toBe('A:b c B:d');
    expect(show(removeSlidesWithSections(slides('A:a b B:c'), new Set(['a', 'b'])))).toBe('B:c');
    expect(move('A:a b c B:d e', ['a'], 5)).toBe('A:b c B:d e a');
  });

  it('puts moved slides into the section they land in', () => {
    expect(move('A:a b B:c d', ['d'], 2)).toBe('A:a b d B:c');
    expect(move('A:a b B:c d', ['b'], 4)).toBe('A:a B:c d b');
    expect(move('A:a b B:c d', ['c'], 4)).toBe('A:a b B:d c');
    // At the very top, into the first section rather than one of no name.
    expect(move('A:a b B:c d', ['c'], 0)).toBe('A:c a b B:d');
    expect(move('x A:a b', ['b'], 0)).toBe('b x A:a');
  });

  it('moves whole sections past whole sections, after the slides in no section', () => {
    expect(move('A:a b B:c d C:e', ['c', 'd'], 1)).toBe('B:c d A:a b C:e');
    expect(move('A:a b B:c d C:e', ['a', 'b'], 3)).toBe('B:c d A:a b C:e');
    expect(move('x y B:c d', ['c', 'd'], 1)).toBeNull();
    expect(move('A:a b', ['a'], 0)).toBeNull();
  });
});
