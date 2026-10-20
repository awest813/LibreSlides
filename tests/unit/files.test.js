import { describe, it, expect } from 'vitest';
import { formatOfName, fileNameFor } from '../../src/io/file-access.js';
import { portableCopy, nativeBlob, importFile } from '../../src/io/files.js';
import { safeColor, sanitizeDeck } from '../../src/io/sanitize.js';
import { highlightAll } from '../../src/editor/richtext.js';
import { buildSteps, nextPos, prevPos, firstPos } from '../../src/present/present.js';
import { newDeck, newSlide, newShape } from '../../src/model/deck.js';

describe('files on the device', () => {
  it('knows which formats it can save back to', () => {
    expect(formatOfName('Talk.PPTX')).toBe('pptx');
    expect(formatOfName('talk.lslides')).toBe('lslides');
    expect(formatOfName('talk.potx')).toBeNull();
    expect(formatOfName('talk.ppt')).toBeNull();
  });

  it('names files after the presentation', () => {
    expect(fileNameFor('Q3 vs. Q4', 'pptx')).toBe('Q3 vs. Q4.pptx');
    expect(fileNameFor('deck.pptx', 'lslides')).toBe('deck.lslides');
    expect(fileNameFor('a/b:c', 'pptx')).toBe('a-b-c.pptx');
  });

  it('never writes the link to a local file into files or copies', async () => {
    const deck = newDeck({ title: 'Linked' });
    deck.file = { handle: { kind: 'file' }, name: 'Linked.pptx', format: 'pptx', dirty: true };
    expect(portableCopy(deck).file).toBeUndefined();
    const text = await nativeBlob(deck).text();
    expect(JSON.parse(text).deck.file).toBeUndefined();
    const reopened = await importFile(new File([text], 'Linked.lslides'));
    expect(reopened.file).toBeUndefined();
    expect(reopened.id).not.toBe(deck.id);
    expect(reopened.title).toBe('Linked');
  });
});

describe('untrusted colours', () => {
  it('only accepts colours that draw nothing from the network', () => {
    expect(safeColor('#a1b2c3')).toBe('#a1b2c3');
    expect(safeColor('rgba(0, 0, 0, 0.5)')).toBe('rgba(0, 0, 0, 0.5)');
    expect(safeColor('theme:accent1')).toBe('theme:accent1');
    expect(safeColor('red url(https://example.com/x)')).toBeNull();
    expect(safeColor({})).toBeNull();
  });

  it('cleans colours across a deck and drops pictures that were removed', () => {
    const deck = {
      slides: [{
        background: { type: 'solid', color: 'url(https://t/x)' },
        elements: [
          { type: 'shape', fill: 'url(x)', stroke: { color: 'expression(1)', width: 1 }, style: { color: 'url(y)', font: 'Inter"; background:url(z)' }, text: '' },
          { type: 'image', media: 'gone', x: 0, y: 0, w: 1, h: 1 },
        ],
      }],
      media: { gone: 'https://example.com/a.png' },
    };
    sanitizeDeck(deck);
    const [shape] = deck.slides[0].elements;
    expect(deck.slides[0].background.color).toBe('#ffffff');
    expect(shape.fill).toBe('none');
    expect(shape.stroke.color).toBe('none');
    expect(shape.style.color).toBeNull();
    expect(shape.style.font).not.toMatch(/[;:()"]/);
    expect(deck.slides[0].elements).toHaveLength(1);
  });
});

describe('box-wide highlight', () => {
  it('wraps each paragraph once and can be removed', () => {
    const html = highlightAll('<p>a <b>b</b></p><ul><li>c</li></ul>', '#ffff00');
    expect(html).toBe('<p><span style="background-color: rgb(255, 255, 0);">a <b>b</b></span></p><ul><li><span style="background-color: rgb(255, 255, 0);">c</span></li></ul>');
    expect(highlightAll(html, null)).toBe('<p>a <b>b</b></p><ul><li>c</li></ul>');
  });
});

describe('animations that start on their own', () => {
  it('plays a first "with previous" animation as the slide appears', () => {
    const d = newDeck();
    const s = newSlide('blank', d);
    const a = newShape('rect', 0, 0);
    const b = newShape('rect', 0, 0);
    a.anim = { effect: 'fade', order: 1, start: 'after' };
    b.anim = { effect: 'zoom', order: 2 };
    s.elements.push(a, b);
    d.slides.push(s);
    expect(buildSteps(s).auto).toBe(true);
    expect(nextPos(d, { i: 0, step: 0 })).toEqual({ i: 1, step: 1 });
    expect(nextPos(d, { i: 1, step: 1 })).toEqual({ i: 1, step: 2 });
    expect(prevPos(d, { i: 1, step: 1 })).toEqual({ i: 0, step: 0 });
    expect(firstPos(d, 1)).toEqual({ i: 1, step: 1 });
  });
});
