import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { newDeck, newShape, newTextBox, newTable, newImage } from '../../src/model/deck.js';
import { buildPptx } from '../../src/io/pptx-export.js';
import { importPptx } from '../../src/io/pptx-import.js';
import { readAnimations } from '../../src/io/pptx-timing.js';

const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

function deckWith(...elements) {
  const deck = newDeck({ slides: [] });
  deck.slides.push({ id: 's', layout: 'blank', background: null, elements, notes: '', transition: { type: 'none', duration: 0.6 }, hidden: false });
  deck.media.m1 = PNG;
  return deck;
}
const exported = async (deck) => JSZip.loadAsync(await (await buildPptx(deck)).arrayBuffer());
const slideXml = async (deck) => (await exported(deck)).file('ppt/slides/slide1.xml').async('string');
const parse = (xml) => new DOMParser().parseFromString(xml, 'application/xml');

/** Four elements: text, shape, picture and table, animated click / with / after / click. */
function animatedDeck() {
  const a = newTextBox(10, 10);
  a.text = '<p>Title</p>';
  a.anim = { effect: 'fade', duration: 0.5, start: 'click', order: 1 };
  const b = newShape('rect', 10, 100);
  b.anim = { effect: 'fly-left', duration: 0.75, start: 'with', order: 2 };
  const c = newImage('m1', 300, 100, 100, 100);
  c.anim = { effect: 'zoom', duration: 1, start: 'after', order: 3 };
  const d = newTable(2, 2, 10, 300, 300, 60);
  d.anim = { effect: 'wipe', duration: 0.5, start: 'click', order: 4 };
  return deckWith(a, b, c, d);
}

describe('animations in PowerPoint files', () => {
  it('writes entrance effects as PowerPoint writes them', async () => {
    const doc = parse(await slideXml(animatedDeck()));
    const effects = [...doc.getElementsByTagNameNS('*', 'cTn')].filter((c) => c.getAttribute('presetClass') === 'entr');
    expect(effects.map((e) => [e.getAttribute('presetID'), e.getAttribute('presetSubtype'), e.getAttribute('nodeType')])).toEqual([
      ['10', '0', 'clickEffect'], ['2', '8', 'withEffect'], ['53', '16', 'afterEffect'], ['22', '8', 'clickEffect'],
    ]);
    // Each effect points at its own shape, by an id no other shape has.
    const ids = [...doc.getElementsByTagNameNS('*', 'cNvPr')].map((c) => c.getAttribute('id'));
    expect(new Set(ids).size).toBe(ids.length);
    const targets = effects.map((e) => e.getElementsByTagNameNS('*', 'spTgt')[0].getAttribute('spid'));
    const names = targets.map((id) => [...doc.getElementsByTagNameNS('*', 'cNvPr')].find((c) => c.getAttribute('id') === id)?.parentElement.parentElement.localName);
    expect(names).toEqual(['sp', 'sp', 'pic', 'graphicFrame']);
    // The zoom starts after the fly-in ends (0.75 s into the click).
    const groups = effects.map((e) => e.parentElement.parentElement.parentElement.firstElementChild.firstElementChild.getAttribute('delay'));
    expect(groups.slice(0, 3)).toEqual(['0', '0', '750']);
    expect(doc.getElementsByTagNameNS('*', 'bldP')[0].getAttribute('spid')).toBe(targets[0]);
  });

  it('reads them back in order', async () => {
    const zip = await exported(animatedDeck());
    const back = await importPptx(await zip.generateAsync({ type: 'arraybuffer' }));
    expect(back.slides[0].elements.map((e) => e.anim)).toEqual([
      { effect: 'fade', duration: 0.5, start: 'click', order: 1 },
      { effect: 'fly-left', duration: 0.75, start: 'with', order: 2 },
      { effect: 'zoom', duration: 1, start: 'after', order: 3 },
      { effect: 'wipe', duration: 0.5, start: 'click', order: 4 },
    ]);
  });

  it('writes and reads emphasis and exit effects', async () => {
    const a = newShape('rect', 10, 10);
    a.anim = { effect: 'fade', duration: 0.5, start: 'click', order: 1 };
    a.emphasis = { effect: 'pulse', duration: 0.8, start: 'click', order: 2 };
    a.exit = { effect: 'fly-left', duration: 0.5, start: 'after', order: 3 };
    const xml = await slideXml(deckWith(a));
    expect(xml).toContain('presetID="26" presetClass="emph"');
    expect(xml).toContain('presetID="2" presetClass="exit" presetSubtype="8"');
    // Going out, it's hidden once it has gone.
    expect(xml).toMatch(/<p:cond delay="499"\/><\/p:stCondLst><\/p:cTn><p:tgtEl><p:spTgt spid="\d+"\/><\/p:tgtEl><p:attrNameLst><p:attrName>style.visibility<\/p:attrName><\/p:attrNameLst><\/p:cBhvr><p:to><p:strVal val="hidden"\/>/);
    const zip = await exported(deckWith(a));
    const back = (await importPptx(await zip.generateAsync({ type: 'arraybuffer' }))).slides[0].elements[0];
    expect([back.anim, back.emphasis, back.exit]).toEqual([
      { effect: 'fade', duration: 0.5, start: 'click', order: 1 },
      { effect: 'pulse', duration: 0.8, start: 'click', order: 2 },
      { effect: 'fly-left', duration: 0.5, start: 'after', order: 3 },
    ]);
  });

  it('starts by itself when the first effect follows the slide', async () => {
    const a = newShape('rect', 10, 10);
    a.anim = { effect: 'fly-down', duration: 0.5, start: 'after', order: 1 };
    const xml = await slideXml(deckWith(a));
    expect(xml).toMatch(/<p:cond delay="indefinite"\/><p:cond evt="onBegin" delay="0"><p:tn val="2"\/><\/p:cond>/);
    expect(xml).toContain('presetID="2" presetClass="entr" presetSubtype="1"');
  });

  it('puts the timeline after the transition', async () => {
    const deck = animatedDeck();
    deck.slides[0].transition = { type: 'fade', duration: 0.6 };
    const xml = await slideXml(deck);
    expect(xml.indexOf('<p:transition')).toBeLessThan(xml.indexOf('<p:timing>'));
    expect(xml.indexOf('</p:timing>')).toBeLessThan(xml.indexOf('</p:sld>'));
  });

  it('reads PowerPoint\'s effects into the closest ones LibreSlides has', () => {
    const effect = (spid, id, sub, type, extra = '', cls = 'entr') => `<p:par><p:cTn presetID="${id}" presetClass="${cls}" presetSubtype="${sub}" nodeType="${type}"><p:childTnLst>`
      + `<p:animEffect><p:cBhvr><p:cTn dur="2000"/><p:tgtEl><p:spTgt spid="${spid}">${extra}</p:spTgt></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`;
    const doc = parse(`<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:timing><p:tnLst><p:par><p:cTn><p:childTnLst><p:seq><p:cTn nodeType="mainSeq"><p:childTnLst>${
      effect(4, 3, 10, 'clickEffect') // blinds
      + effect(4, 3, 10, 'withEffect', '<p:txEl><p:pRg st="1" end="1"/></p:txEl>') // the same shape's next paragraph
      + effect(5, 10, 0, 'afterEffect', '', 'exit')
      + effect(6, 2, 2, 'withEffect')
      + effect(6, 32, 0, 'clickEffect', '', 'emph') // teeter
      + effect(6, 99, 0, 'clickEffect', '', 'path') // a motion path: left out
    }</p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing></p:sld>`);
    expect(readAnimations(doc)).toEqual([
      { spid: '4', kind: 'anim', effect: 'fade', duration: 2, start: 'click', order: 1 },
      { spid: '5', kind: 'exit', effect: 'fade', duration: 2, start: 'after', order: 2 },
      { spid: '6', kind: 'anim', effect: 'fly-right', duration: 2, start: 'with', order: 3 },
      { spid: '6', kind: 'emphasis', effect: 'teeter', duration: 2, start: 'click', order: 4 },
    ]);
  });
});
