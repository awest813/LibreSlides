import { describe, it, expect } from 'vitest';
import { renderElement, renderSlide, renderThumb } from '../../src/render/render.js';
import {
  newDeck, newSlide, newShape, newTextBox, newImage, newLine, newTable, newChart, newMedia, newDiagram, addMedia, DIAGRAM_LAYOUTS, CHART_KINDS,
} from '../../src/model/deck.js';
import { THEMES, themeById } from '../../src/model/themes.js';

const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+ip1sAAAAASUVORK5CYII=';

/** A slide with one of every kind of element, in most of their variations. */
function everything(deck) {
  const slide = newSlide('title', deck);
  const pic = addMedia(deck, PNG);
  const els = [];
  for (const [i, kind] of ['rect', 'ellipse', 'triangle', 'star5', 'rightArrow', 'custom'].entries()) {
    const s = newShape(kind, i * 20, 0, 100, 80);
    if (kind === 'custom') Object.assign(s, { paths: [{ cmds: [['M', 0, 0], ['L', 1, 1]], fill: false }], arrowEnd: 'stealth', arrowStart: 'oval' });
    Object.assign(s, { rot: i * 30, flipH: i % 2 === 1, flipV: i % 3 === 1, shadow: i === 2, text: '<p>Hi <span data-field="slidenum">?</span></p>' });
    if (i === 3) s.gradient = { type: 'radial', stops: [{ pos: 0, color: '#fff' }, { pos: 1, color: 'theme:accent2' }] };
    if (i === 4) s.fillImage = { media: pic, crop: { l: 0.1, t: 0, r: 0, b: 0.2 } };
    els.push(s);
  }
  const t = newTextBox(0, 100, 200, 100);
  Object.assign(t, { text: '<p>Columns</p>', style: { columns: 2, vertical: 'vert270', inset: { l: 1, t: 2, r: 3, b: 4 }, wrap: false } });
  const w = newTextBox(0, 200, 200, 100);
  Object.assign(w, { text: '<p>Warped <span data-field="slidenum">?</span></p>', style: { warp: { preset: 'textArchUp' } } });
  els.push(t, w);
  els.push(Object.assign(newImage(pic, 0, 0, 50, 50), { flipH: true, clip: { shape: 'ellipse' }, stroke: { color: '#000', width: 2 }, adjust: { grayscale: true }, crop: { l: 0.1, t: 0.1, r: 0, b: 0 } }));
  els.push(newLine(0, 0, 0, 100, { arrowEnd: 'diamond' }), newLine(5, 5, 100, 5));
  const table = newTable(3, 3, 0, 300, 300, 100);
  table.cells[0][0].colSpan = 2;
  table.cells[0][1].merged = true;
  els.push(table);
  for (const { id } of CHART_KINDS) els.push(Object.assign(newChart(0, 0, 200, 150), { chart: { ...newChart(0, 0).chart, kind: id, dataLabels: true, title: id } }));
  els.push(newMedia('video', pic, pic, 0, 0, 100, 60), newMedia('audio', pic, null, 0, 0, 40, 40));
  for (const { id } of DIAGRAM_LAYOUTS) els.push(newDiagram(id, 0, 0, 300, 200));
  slide.elements.push(...els);
  return slide;
}

describe('drawing slides', () => {
  it('draws every kind of element, in every theme and slide size, without NaN', () => {
    const cases = [...THEMES.map((t) => [t.id, '16:9']), ...['4:3', 'a4-portrait', '500x700'].map((size) => ['simple-dark', size])];
    for (const [themeId, size] of cases) {
      const deck = newDeck({ themeId, size });
      deck.slides.push(everything(deck));
      const node = renderSlide(deck.slides[1], deck, { editing: true });
      expect(node.querySelectorAll(':scope > .el')).toHaveLength(deck.slides[1].elements.length);
      expect(node.innerHTML).not.toMatch(/NaN|undefined|Infinity/);
      // Fields show the slide's number.
      expect([...node.querySelectorAll('[data-field="slidenum"]')].every((f) => f.textContent === '2')).toBe(true);
    }
    const deck = newDeck();
    deck.slides.push(everything(deck));
    expect(renderSlide(deck.slides[1], deck, { live: true }).querySelectorAll('video, audio')).toHaveLength(2);
  }, 20000);

  it('flips the border of a picture cropped to a shape with the picture', () => {
    const deck = newDeck();
    const el = Object.assign(newImage(addMedia(deck, PNG), 0, 0, 100, 80), { flipH: true, clip: { shape: 'rtTriangle' }, stroke: { color: '#000000', width: 3, dash: 'solid' } });
    const node = renderElement(el, { theme: themeById(deck.themeId), deck });
    const [frame, border] = [node.querySelector('.el-img'), node.querySelector('svg.el-svg')];
    expect(frame.style.clipPath).toMatch(/^path\(/);
    expect(frame.style.transform).toBe('scale(-1, 1)');
    expect(border.style.transform).toBe('scale(-1, 1)');
  });

  it('leaves thumbnails off screen for the browser to skip', () => {
    const deck = newDeck();
    const thumb = renderThumb(deck.slides[0], deck, 160);
    expect(thumb.style.contentVisibility).toBe('auto');
    expect(thumb.style.height).toBe('90px');
  });
});
