import { describe, it, expect } from 'vitest';
import { WARPS, WARP_GROUPS, cleanWarp, isPlainWarp, mainAdjust, warpDefaults } from '../../src/model/warps.js';
import { placeCharacters, textLines, warpPaths, gradientOfCss } from '../../src/render/warp.js';
import { renderElement } from '../../src/render/render.js';
import { themeById } from '../../src/model/themes.js';
import { newTextBox } from '../../src/model/deck.js';
import { sanitizeElement } from '../../src/io/sanitize.js';

const base = { font: 'Inter', size: 40, color: '#000' };
const linesOf = (html) => {
  const div = document.createElement('div');
  div.innerHTML = html;
  return textLines(div, base);
};
// Where a placed character's transform puts it.
const where = (c) => {
  const m = /translate\(([-\d.]+) ([-\d.]+)\)/.exec(c.transform) || /matrix\((?:[-\d.]+ ){4}([-\d.]+) ([-\d.]+)\)/.exec(c.transform);
  return [+m[1], +m[2]];
};

describe('text warps', () => {
  it('works out every preset, and names each in the editor', () => {
    expect(Object.keys(WARPS).length).toBeGreaterThanOrEqual(40);
    const named = WARP_GROUPS.flatMap(([, list]) => list.map(([id]) => id));
    for (const id of named) expect(WARPS[id], id).toBeTruthy();
    for (const id of Object.keys(WARPS)) if (id !== 'textPlain') expect(named, id).toContain(id);
    for (const preset of Object.keys(WARPS)) {
      const paths = warpPaths({ preset }, 300, 120);
      expect(paths.length, preset).toBeGreaterThan(0);
      for (const p of paths) for (const pt of p.pts) expect(pt.every(Number.isFinite), preset).toBe(true);
    }
  });

  it('keeps only known presets with adjust values in range', () => {
    expect(cleanWarp({ preset: 'textArchUp', adjust: { adj: 99e9, bogus: 3 } })).toEqual({ preset: 'textArchUp', adjust: { adj: 21599999 } });
    expect(cleanWarp({ preset: 'textWave1', adjust: { adj1: NaN, adj2: '5' }, stretch: true })).toEqual({ preset: 'textWave1', stretch: true });
    expect(cleanWarp({ preset: '__proto__' })).toBeNull();
    expect(cleanWarp({ preset: 'textNoShape' })).toBeNull();
    expect(cleanWarp({ preset: 'textPlain', adjust: { adj: 50000 } })).toBeNull();
    expect(cleanWarp({ preset: 'textPlain', adjust: { adj: 40000 } })).toEqual({ preset: 'textPlain', adjust: { adj: 40000 } });
    expect(isPlainWarp(undefined)).toBe(true);
    expect(warpDefaults('textArchUp')).toEqual({ adj: 10800000 });
    expect(mainAdjust({ preset: 'textInflate' })).toEqual({ name: 'adj', min: 0, max: 20000, value: 18750 });
    const el = sanitizeElement({ ...newTextBox(0, 0), style: { warp: { preset: 'textCircle', adjust: { adj: -5 } } } });
    expect(el.style.warp).toEqual({ preset: 'textCircle', adjust: { adj: 0 } });
    expect(sanitizeElement({ ...newTextBox(0, 0), style: { warp: { preset: '<script>' } } }).style.warp).toBeUndefined();
  });

  it('reads the lines of a box with each run’s look', () => {
    const lines = linesOf('<p>ab <b>c</b></p><p><span style="color: red; font-size: 20px; -webkit-text-stroke: 2px blue">d</span><br>e</p>');
    expect(lines.map((l) => l.map((c) => c.ch).join(''))).toEqual(['ab c', 'd', 'e']);
    expect(lines[0][3].style.bold).toBe(true);
    expect(lines[1][0].style).toMatchObject({ color: 'red', size: 20, outline: { width: 2, color: 'blue' } });
  });

  it('runs text along an arch, turning each character', () => {
    const chars = placeCharacters({ preset: 'textArchUp' }, linesOf('<p>ARCHED TEXT</p>'), { w: 400, h: 200, align: 'center' }).filter((c) => c.ch !== ' ');
    const ys = chars.map((c) => where(c)[1]);
    const angles = chars.map((c) => +/rotate\(([-\d.]+)\)/.exec(c.transform)[1]);
    // Higher in the middle than at the ends, and turning from rising to falling.
    expect(Math.min(...ys)).toBeLessThan(ys[0] - 10);
    expect(Math.min(...ys)).toBeLessThan(ys.at(-1) - 10);
    expect(angles[0]).toBeLessThan(-10);
    expect(angles.at(-1)).toBeGreaterThan(10);
    // In order from left to right.
    const xs = chars.map((c) => where(c)[0]);
    expect(xs).toEqual([...xs].sort((a, b) => a - b));
  });

  it('stretches text between an envelope’s top and bottom', () => {
    const chars = placeCharacters({ preset: 'textInflate' }, linesOf('<p>INFLATE</p>'), { w: 420, h: 150 });
    const scaleY = chars.map((c) => +c.transform.match(/[-\d.]+/g)[3]);
    // Taller in the middle (the envelope bulges) than at the ends, and filling the width.
    const mid = scaleY[Math.floor(chars.length / 2)];
    expect(mid).toBeGreaterThan(scaleY[0] * 1.2);
    expect(mid).toBeGreaterThan(scaleY.at(-1) * 1.2);
    expect(where(chars[0])[0]).toBeLessThan(60);
    expect(where(chars.at(-1))[0]).toBeGreaterThan(360);
    // Two lines share the height: the second line sits below the first.
    const two = placeCharacters({ preset: 'textInflate' }, linesOf('<p>ONE</p><p>TWO</p>'), { w: 420, h: 150 });
    expect(where(two[4])[1]).toBeGreaterThan(where(two[1])[1] + 30);
  });

  it('slants text with a slant warp', () => {
    const chars = placeCharacters({ preset: 'textSlantUp' }, linesOf('<p>SLANTED</p>'), { w: 400, h: 200 });
    const ys = chars.map((c) => where(c)[1]);
    expect(ys[0]).toBeGreaterThan(ys.at(-1) + 50);
  });

  it('draws a warped box as positioned SVG characters over hidden text', () => {
    const el = newTextBox(10, 20, 400, 160);
    el.text = '<p>Hello <span style="color: #ff0000; text-shadow: rgba(0, 0, 0, 0.5) 2px 2px 3px">world</span></p>';
    el.style = { ...el.style, warp: { preset: 'textWave1' } };
    const node = renderElement(el, { theme: themeById('simple-light') });
    const box = node.querySelector('.el-textbox');
    expect(box.classList.contains('is-warped')).toBe(true);
    expect(box.querySelector('.el-text-body').textContent).toBe('Hello world');
    const glyphs = [...box.querySelectorAll('svg.el-warp text')];
    expect(glyphs.map((g) => g.textContent).join('')).toBe('Helloworld');
    expect(glyphs.every((g) => /^matrix\(/.test(g.getAttribute('transform')))).toBe(true);
    expect(glyphs.at(-1).getAttribute('fill')).toBe('#ff0000');
    // The shadow is an SVG filter on the run's characters.
    expect(glyphs.at(-1).parentNode.getAttribute('filter')).toMatch(/^url\(#/);
    expect(box.querySelector('feDropShadow').getAttribute('dx')).toBe('2');
    // A gradient shows through the characters, across the whole box.
    el.text = '<p><span style="color: #f00; background-image: linear-gradient(90deg, #f00 0%, #00f 100%); -webkit-background-clip: text; background-clip: text; -webkit-text-fill-color: transparent">Hi</span></p>';
    const svg = renderElement(el, { theme: themeById('simple-light') }).querySelector('svg.el-warp');
    expect([...svg.querySelectorAll('mask text')].map((g) => g.textContent).join('')).toBe('Hi');
    expect(svg.querySelector('rect[mask]').getAttribute('fill')).toMatch(/^url\(#ls-grad-/);
    expect(svg.querySelector('linearGradient stop:last-child').getAttribute('stop-color')).toBe('#00f');
    // Without a warp, no drawing.
    delete el.style.warp;
    expect(renderElement(el, { theme: themeById('simple-light') }).querySelector('.el-warp')).toBeNull();
  });

  it('reads CSS gradients for gradient-filled characters', () => {
    expect(gradientOfCss('linear-gradient(90deg, rgba(1, 2, 3, 0.5) 0%, #fff 100%)')).toEqual({ type: 'linear', angle: 90, stops: [{ pos: 0, color: 'rgba(1, 2, 3, 0.5)' }, { pos: 1, color: '#fff' }] });
    expect(gradientOfCss('radial-gradient(circle at 50% 25%, red 0%, blue 100%)')).toMatchObject({ type: 'radial', cx: 0.5, cy: 0.25 });
  });
});
