import { describe, it, expect, vi, afterEach } from 'vitest';
import JSZip from 'jszip';
import { buildPptx } from '../../src/io/pptx-export.js';
import { importPptx } from '../../src/io/pptx-import.js';
import { playbackOf } from '../../src/io/pptx-media.js';
import { sanitizeDeck } from '../../src/io/sanitize.js';
import { isMediaData, mediaTypeOf } from '../../src/io/media-files.js';
import { renderMedia } from '../../src/render/media.js';
import { renderSlide } from '../../src/render/render.js';
import { Player, startSlideshow, firstPos, nextPos } from '../../src/present/present.js';
import { newDeck, newMedia, newShape, addMedia, usedMedia, pruneMedia } from '../../src/model/deck.js';
import { themeById } from '../../src/model/themes.js';

const toBuffer = (blob) => (blob.arrayBuffer ? blob.arrayBuffer() : new Response(blob).arrayBuffer());
// A 1 × 1 PNG, and stand-ins for a video and a sound (the bytes needn't play).
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const MP4 = 'data:video/mp4;base64,AAAAIGZ0eXBpc29tAAACAGlzb21pc28yYXZjMW1wNDEAAAAIZnJlZQ==';
const MP3 = 'data:audio/mpeg;base64,SUQzBAAAAAAAI1RTU0UAAAAPAAADTGF2ZjU4Ljc2LjEwMAAAAAAA';

/** A deck with a video (with a poster, playing by itself and looping) and a sound on its first slide. */
function mediaDeck() {
  const deck = newDeck({ slides: [] });
  const video = newMedia('video', addMedia(deck, MP4), addMedia(deck, PNG), 96, 54, 480, 270);
  video.autoplay = true;
  video.loop = true;
  const sound = newMedia('audio', addMedia(deck, MP3), null, 700, 400, 64, 64);
  deck.slides.push({ id: 's1', layout: 'blank', background: null, elements: [newShape('rect', 10, 10), video, sound], notes: '', transition: { type: 'none', duration: 0.6 }, hidden: false });
  return { deck, video, sound };
}

describe('video and sound in PowerPoint files', () => {
  it('exports them as real PowerPoint media and reads them back', async () => {
    const { deck, video, sound } = mediaDeck();
    const zip = await JSZip.loadAsync(await toBuffer(await buildPptx(deck)));
    const files = Object.keys(zip.files);
    expect(files.some((f) => /^ppt\/media\/[^/]+\.mp4$/.test(f))).toBe(true);
    expect(files.some((f) => /^ppt\/media\/[^/]+\.mp3$/.test(f))).toBe(true);

    const types = await zip.file('[Content_Types].xml').async('string');
    expect(types).toContain('<Default Extension="mp4" ContentType="video/mp4"/>');
    expect(types).toContain('<Default Extension="mp3" ContentType="audio/mpeg"/>');

    const xml = await zip.file('ppt/slides/slide1.xml').async('string');
    expect(xml).toMatch(/<a:videoFile r:link="rId\d+"\/>/);
    expect(xml).toMatch(/<a:audioFile r:link="rId\d+"\/>/);
    expect(xml.match(/<p14:media [^>]*r:embed="rId\d+"/g)).toHaveLength(2);
    // Shape ids stay unique on the slide, and the timeline plays and loops the video by its id.
    const ids = xml.match(/<p:cNvPr id="(\d+)"/g).map((m) => m.match(/\d+/)[0]);
    expect(new Set(ids).size).toBe(ids.length);
    expect(xml).toMatch(/<p:timing>[\s\S]*cmd="playFrom\(0.0\)"[\s\S]*repeatCount="indefinite"[\s\S]*<\/p:timing><\/p:sld>/);

    const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('string');
    expect(rels).toMatch(/relationships\/video" Target="\.\.\/media\/[^"]+\.mp4"/);
    expect(rels).toMatch(/relationships\/audio" Target="\.\.\/media\/[^"]+\.mp3"/);
    expect(rels.match(/office\/2007\/relationships\/media"/g)).toHaveLength(2);

    const back = await importPptx(await zip.generateAsync({ type: 'arraybuffer' }), 'x');
    const media = back.slides[0].elements.filter((e) => e.type === 'media');
    expect(media.map((e) => e.kind)).toEqual(['video', 'audio']);
    const [v, a] = media;
    expect(back.media[v.media]).toBe(MP4);
    expect(back.media[v.poster]).toBe(PNG);
    expect(back.media[a.media]).toBe(MP3);
    expect(v).toMatchObject({ autoplay: true, loop: true });
    expect(a).toMatchObject({ autoplay: false, loop: false });
    expect(v.x).toBeCloseTo(video.x, 0);
    expect(v.w).toBeCloseTo(video.w, 0);
    expect(a.y).toBeCloseTo(sound.y, 0);
    // pptxgenjs's play button stands in for the sound's missing poster.
    expect(back.media[a.poster]).toMatch(/^data:image\/png;base64,/);
  });

  it('leaves videos it can\'t play as their poster picture', async () => {
    const { deck } = mediaDeck();
    const zip = await JSZip.loadAsync(await toBuffer(await buildPptx(deck)));
    // Rename the video to a Windows Media file, as an old presentation might have.
    const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('string');
    const target = /Target="\.\.\/media\/([^"]+\.mp4)"/.exec(rels)[1];
    zip.file(`ppt/media/${target.replace('.mp4', '.wmv')}`, await zip.file(`ppt/media/${target}`).async('uint8array'));
    zip.file('ppt/slides/_rels/slide1.xml.rels', rels.replaceAll(target, target.replace('.mp4', '.wmv')));
    const back = await importPptx(await zip.generateAsync({ type: 'arraybuffer' }), 'x');
    const els = back.slides[0].elements;
    expect(els.filter((e) => e.type === 'media').map((e) => e.kind)).toEqual(['audio']);
    expect(els.some((e) => e.type === 'image' && back.media[e.media] === PNG)).toBe(true);
  });

  it('reads "start automatically" only from the first step when it needs no click', () => {
    const doc = (nodeType) => new DOMParser().parseFromString(`<p:sld xmlns:p="p"><p:timing><p:tnLst><p:par><p:cTn id="1" nodeType="tmRoot"><p:childTnLst>
      <p:seq><p:cTn id="2" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3"><p:childTnLst><p:par><p:cTn id="4"><p:childTnLst>
      <p:par><p:cTn id="5" presetClass="mediacall" nodeType="${nodeType}"><p:childTnLst><p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="6" dur="4500"/>
      <p:tgtEl><p:spTgt spid="7"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq>
      </p:childTnLst></p:cTn></p:par></p:tnLst></p:timing></p:sld>`, 'application/xml');
    expect(playbackOf(doc('afterEffect'), '7')).toEqual({ autoplay: true, loop: false, duration: 4.5 });
    expect(playbackOf(doc('clickEffect'), '7').autoplay).toBe(false);
    expect(playbackOf(doc('withEffect'), '8').autoplay).toBe(false);
  });
});

describe('video and sound in the model and untrusted files', () => {
  it('counts posters as used', () => {
    const { deck, video } = mediaDeck();
    expect(usedMedia(deck).has(video.poster)).toBe(true);
    addMedia(deck, 'data:image/png;base64,AAAA');
    pruneMedia(deck);
    expect(Object.keys(deck.media)).toHaveLength(3);
  });

  it('keeps playable media and drops anything else', () => {
    expect(isMediaData(MP4)).toBe(true);
    expect(isMediaData('data:video/x-ms-wmv;base64,AAAA')).toBe(false);
    expect(isMediaData('data:text/html;base64,PHNjcmlwdD4=')).toBe(false);
    expect(mediaTypeOf({ type: '', name: 'clip.M4V' })).toBe('video/mp4');
    expect(mediaTypeOf({ type: 'audio/x-m4a', name: 'a' })).toBe('audio/x-m4a');
    const deck = sanitizeDeck({
      media: { v: MP4, a: MP3, p: PNG, bad: 'data:video/x-ms-wmv;base64,AAAA', html: 'data:text/html;base64,PHNjcmlwdD4=' },
      slides: [{ elements: [
        { type: 'media', kind: 'video', media: 'v', poster: 'a', autoplay: 'yes', x: 0, y: 0, w: 10, h: 10 },
        { type: 'media', kind: 'audio', media: 'a', poster: 'p', loop: true, x: 0, y: 0, w: 10, h: 10 },
        { type: 'media', kind: 'video', media: 'bad', x: 0, y: 0, w: 10, h: 10 },
        { type: 'media', kind: 'script', media: 'v', x: 0, y: 0, w: 10, h: 10 },
        { type: 'media', kind: 'video', media: 'missing', x: 0, y: 0, w: 10, h: 10 },
        { type: 'image', media: 'p', x: 0, y: 0, w: 10, h: 10 },
      ] }],
    });
    expect(Object.keys(deck.media).sort()).toEqual(['a', 'p', 'v']);
    const [v, a, img] = deck.slides[0].elements;
    expect(deck.slides[0].elements).toHaveLength(3);
    expect(v).toMatchObject({ kind: 'video', media: 'v', autoplay: false });
    expect(v.poster).toBeUndefined(); // a sound isn't a picture
    expect(a).toMatchObject({ kind: 'audio', poster: 'p', loop: true });
    expect(img.type).toBe('image');
  });
});

// jsdom lacks these.
globalThis.ResizeObserver ||= class { observe() {} disconnect() {} unobserve() {} };
globalThis.matchMedia ||= () => ({ matches: false, addEventListener() {}, removeEventListener() {} });
globalThis.CSS ||= { escape: (s) => String(s).replace(/["\\]/g, '\\$&') };

afterEach(() => {
  document.body.replaceChildren();
  vi.restoreAllMocks();
});

describe('drawing video and sound', () => {
  const ctx = (deck, opts) => ({ theme: themeById(deck.themeId), deck, ...opts });

  it('shows the poster while editing and in thumbnails, with no player to take the clicks', () => {
    const { deck, video, sound } = mediaDeck();
    const editing = renderMedia(video, ctx(deck, { editing: true }));
    expect(editing.querySelector('video, audio')).toBeNull();
    expect(editing.querySelector('img.el-img').getAttribute('draggable')).toBe('false');
    expect(editing.querySelector('.el-media-badge')).not.toBeNull();
    expect(renderMedia(video, ctx(deck, {})).querySelector('.el-media-badge')).toBeNull();
    const speaker = renderMedia(sound, ctx(deck, { editing: true }));
    expect(speaker.querySelector('audio, img')).toBeNull();
    expect(speaker.querySelector('.el-media-badge svg')).not.toBeNull();
    // A video without a poster is a dark frame with a play mark.
    const bare = renderMedia({ ...video, poster: undefined }, ctx(deck, {}));
    expect(bare.querySelector('.el-media-blank')).not.toBeNull();
    expect(bare.querySelector('.el-media-badge')).not.toBeNull();
  });

  it('plays in the slideshow', () => {
    const { deck, video, sound } = mediaDeck();
    const player = renderMedia(video, ctx(deck, { live: true })).querySelector('video');
    expect(player.controls).toBe(true);
    expect(player.loop).toBe(true);
    expect(player.dataset.autoplay).toBe('1');
    expect(player.getAttribute('src')).toBeTruthy();
    expect(player.getAttribute('poster')).toBeTruthy();
    const live = renderMedia(sound, ctx(deck, { live: true }));
    expect(live.querySelector('audio').dataset.autoplay).toBeUndefined();
    const play = vi.spyOn(window.HTMLMediaElement.prototype, 'play').mockImplementation(() => Promise.resolve());
    live.querySelector('.el-media-toggle').click();
    expect(play).toHaveBeenCalledTimes(1);
    // Thumbnails and the editor never get a player.
    expect(renderSlide(deck.slides[0], deck).querySelector('video, audio')).toBeNull();
    expect(renderSlide(deck.slides[0], deck, { live: true }).querySelectorAll('video, audio')).toHaveLength(2);
  });

  it('starts videos set to play by themselves, stops them on leaving the slide, and keeps clicks on them', async () => {
    const { deck } = mediaDeck();
    deck.slides.push({ id: 's2', layout: 'blank', background: null, elements: [], notes: '', transition: { type: 'none', duration: 0.6 }, hidden: false });
    const play = vi.spyOn(window.HTMLMediaElement.prototype, 'play').mockImplementation(() => Promise.resolve());
    const pause = vi.spyOn(window.HTMLMediaElement.prototype, 'pause').mockImplementation(() => {});
    const show = startSlideshow(deck, { start: 0 });
    await Promise.resolve();
    await Promise.resolve();
    expect(play).toHaveBeenCalledTimes(1); // the video, not the sound
    const root = document.querySelector('.show');
    root.querySelector('video').dispatchEvent(new MouseEvent('click', { bubbles: true }));
    root.querySelector('.el-media-toggle').dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(show.pos).toEqual({ i: 0, step: 0 });
    root.click();
    expect(show.pos).toEqual({ i: 1, step: 0 });
    expect(pause).toHaveBeenCalled();
    show.exit();
  });

  it('keeps the presenter\'s console quiet', () => {
    const { deck } = mediaDeck();
    const host = document.createElement('div');
    document.body.append(host);
    const player = new Player(deck, host, { animate: false, media: false });
    player.go(firstPos(deck, 0), 'cut');
    expect(host.querySelector('video, audio')).toBeNull();
    expect(nextPos(deck, player.pos)).toBeNull();
    player.destroy();
  });
});
