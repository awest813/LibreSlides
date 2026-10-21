import { describe, it, expect } from 'vitest';
import { metaOf } from '../../src/storage/db.js';
import { newDeck, newSlide, newShape, newImage, newMedia, newDiagram } from '../../src/model/deck.js';

describe('the presentation list', () => {
  it('keeps the pictures the cover thumbnail draws, and only those', () => {
    const deck = newDeck();
    const cover = newSlide('blank', deck);
    deck.media = { pic: 'data:image/png;base64,AA==', fill: 'data:image/png;base64,AB==', bg: 'data:image/png;base64,AC==', video: 'data:video/mp4;base64,AD==', poster: 'data:image/png;base64,AE==', xml: 'data:application/xml;base64,PHgvPg==', drawn: 'data:image/png;base64,AF==', later: 'data:image/png;base64,AG==' };
    cover.background = { type: 'image', media: 'bg' };
    const shape = newShape('rect', 0, 0);
    shape.fillImage = { media: 'fill' };
    const diagram = newDiagram('list', 0, 0);
    diagram.diagram.drawn = { w: 100, h: 100, elements: [newImage('drawn', 0, 0, 10, 10)] };
    diagram.diagram.source = { data: 'xml', layout: 'xml', style: 'xml', colors: 'xml' };
    cover.elements.push(newImage('pic', 0, 0, 10, 10), shape, newMedia('video', 'video', 'poster', 0, 0, 10, 10), diagram);
    const second = newSlide('blank', deck);
    second.elements.push(newImage('later', 0, 0, 10, 10));
    deck.slides = [cover, second];
    expect(Object.keys(metaOf(deck).cover.media).sort()).toEqual(['bg', 'drawn', 'fill', 'pic', 'poster']);
  });
});
