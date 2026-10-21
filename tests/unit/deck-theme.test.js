import { describe, it, expect } from 'vitest';
import { deckTheme, themeById, resolveColor } from '../../src/model/themes.js';
import { renderSlide } from '../../src/render/render.js';
import { newDeck } from '../../src/model/deck.js';

const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const imported = {
  name: 'Facet', dark: false, fonts: { major: 'Trebuchet MS', minor: 'Trebuchet MS' },
  colors: { dk1: '#000000', lt1: '#ffffff', dk2: '#2c3c43', lt2: '#ebebeb', accent1: '#90c226', accent2: '#54a021', accent3: '#e6b91e', accent4: '#e76618', accent5: '#c42f1a', accent6: '#918655', hlink: '#99ca3c', folHlink: '#b9d181' },
};

describe('a deck’s theme', () => {
  it('is its LibreSlides theme, or the colours, background and fonts of the file it came from', () => {
    const deck = newDeck();
    expect(deckTheme(deck)).toBe(themeById(deck.themeId));
    deck.importedTheme = imported;
    const t = deckTheme(deck);
    expect(t.name).toBe('Facet');
    expect(resolveColor('theme:accent1', t)).toBe('#90c226');
    expect(resolveColor('theme:text', t)).toBe('#000000');
    expect(t.background).toEqual({ type: 'solid', color: '#ffffff' });
    expect(t.fonts).toEqual({ title: 'Trebuchet MS', body: 'Trebuchet MS' });
    expect(deckTheme(deck)).toBe(t);
    expect(resolveColor('theme:text', deckTheme({ ...deck, importedTheme: { ...imported, dark: true } }))).toBe('#ffffff');
  });

  it('draws a tiled background picture whose tiles mirror as mirrored tiles', () => {
    const deck = newDeck();
    deck.media = { m1: PNG };
    const slide = deck.slides[0];
    slide.background = { type: 'image', media: 'm1', tile: { w: 40, h: 30, x: 0, y: 0, flip: 'xy' } };
    const root = renderSlide(slide, deck);
    const images = root.querySelectorAll('.slide-bg-tiles pattern image');
    expect(images).toHaveLength(4);
    expect(images[3].getAttribute('transform')).toBe('translate(80 60) scale(-1 -1)');
    slide.background.tile.flip = 'none';
    expect(renderSlide(slide, deck).querySelector('.slide-bg-tiles')).toBeNull();
  });
});
