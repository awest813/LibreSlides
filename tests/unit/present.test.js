import { describe, it, expect, vi, afterEach } from 'vitest';
import { buildSteps, nextPos, prevPos, firstPos, Player, startSlideshow } from '../../src/present/present.js';
import { newDeck, newSlide, newShape } from '../../src/model/deck.js';

function deckWithAnimations() {
  const d = newDeck();
  const s = newSlide('blank', d);
  const a = newShape('rect', 0, 0);
  const b = newShape('rect', 0, 0);
  const c = newShape('rect', 0, 0);
  a.anim = { effect: 'fade', order: 1, duration: 0.5 };
  b.anim = { effect: 'zoom', order: 2, duration: 0.4, start: 'after' };
  c.anim = { effect: 'fly-up', order: 3, duration: 0.5 };
  s.elements.push(c, b, a);
  const hidden = newSlide('blank', d);
  hidden.hidden = true;
  d.slides.push(s, hidden, newSlide('title', d));
  return d;
}

describe('slideshow positions', () => {
  it('groups animations into click steps with delays', () => {
    const d = deckWithAnimations();
    const steps = buildSteps(d.slides[1]);
    expect(steps).toHaveLength(2);
    expect(steps[0].map((x) => x.effect)).toEqual(['fade', 'zoom']);
    expect(steps[0][1].delay).toBe(0.5);
    expect(steps[1][0].effect).toBe('fly-up');
  });

  it('steps through builds and skips hidden slides', () => {
    const d = deckWithAnimations();
    let p = firstPos(d, 0);
    const seen = [];
    while (p) {
      seen.push(`${p.i}.${p.step}`);
      p = nextPos(d, p);
    }
    expect(seen).toEqual(['0.0', '1.0', '1.1', '1.2', '3.0']);
    expect(prevPos(d, { i: 3, step: 0 })).toEqual({ i: 1, step: 2 });
    expect(prevPos(d, { i: 0, step: 0 })).toBeNull();
    expect(firstPos(d, 2)).toEqual({ i: 3, step: 0 });
  });
});

// jsdom lacks these.
globalThis.ResizeObserver ||= class { observe() {} disconnect() {} unobserve() {} };
globalThis.matchMedia ||= () => ({ matches: false, addEventListener() {}, removeEventListener() {} });
globalThis.CSS ||= { escape: (s) => String(s).replace(/["\\]/g, '\\$&') };

afterEach(() => {
  document.body.replaceChildren();
  delete Element.prototype.animate;
});

describe('playing a slideshow', () => {
  it('plays the transition into a slide whose first animation starts by itself', () => {
    const d = deckWithAnimations();
    const s = d.slides[1];
    const fade = s.elements[2];
    fade.anim.start = 'with';
    fade.opacity = 0.5;
    s.transition = { type: 'fade', duration: 0.6 };
    const animate = vi.fn(() => ({ finished: Promise.resolve(), finish() {} }));
    Element.prototype.animate = animate;
    const host = document.createElement('div');
    document.body.append(host);
    const player = new Player(d, host);
    player.go(firstPos(d, 0));
    player.go(nextPos(d, player.pos));
    expect(player.pos).toEqual({ i: 1, step: 1 });
    const layer = player.frame.lastElementChild;
    expect(animate.mock.contexts).toContain(layer);
    // The fade ends at the object's own opacity instead of jumping back to it.
    const node = layer.querySelector(`[data-id="${fade.id}"]`);
    const call = animate.mock.calls[animate.mock.contexts.indexOf(node)];
    expect(call[0].at(-1).opacity).toBe(0.5);
  });

  it('brings the slide back on the first click after blacking out, and leaves Ctrl combinations alone', () => {
    const show = startSlideshow(deckWithAnimations(), { start: 0 });
    const root = document.querySelector('.show');
    const curtain = root.querySelector('.show-curtain');
    const key = (init) => {
      const e = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init });
      root.dispatchEvent(e);
      return e;
    };
    key({ key: 'b' });
    expect(curtain.hidden).toBe(false);
    root.click();
    expect(curtain.hidden).toBe(true);
    expect(show.pos).toEqual({ i: 0, step: 0 });
    expect(key({ key: 'p', ctrlKey: true }).defaultPrevented).toBe(false);
    expect(key({ key: 'n', metaKey: true }).defaultPrevented).toBe(false);
    expect(show.pos).toEqual({ i: 0, step: 0 });
    root.click();
    expect(show.pos).toEqual({ i: 1, step: 0 });
    show.exit();
  });
});

describe('emphasis and exit effects', () => {
  it('shows an object until its exit, and only from its entrance', () => {
    const d = newDeck();
    const s = newSlide('blank', d);
    const a = newShape('rect', 0, 0);
    const b = newShape('rect', 0, 0);
    a.anim = { effect: 'fade', order: 1, duration: 0.5 };
    a.emphasis = { effect: 'pulse', order: 2, duration: 0.5 };
    b.exit = { effect: 'fade', order: 3, duration: 0.5 };
    s.elements.push(a, b);
    d.slides = [s];
    const steps = buildSteps(s);
    expect(steps.map((st) => st.map((x) => `${x.kind}:${x.effect}`))).toEqual([['anim:fade'], ['emphasis:pulse'], ['exit:fade']]);
    const p = new Player(d, document.createElement('div'));
    const visible = (step) => {
      const layer = p.makeLayer(0, step);
      return [a, b].map((el) => layer.querySelector(`[data-id="${el.id}"]`).style.visibility !== 'hidden');
    };
    expect(visible(0)).toEqual([false, true]);
    expect(visible(1)).toEqual([true, true]);
    expect(visible(2)).toEqual([true, true]);
    expect(visible(3)).toEqual([true, false]);
  });
});

describe('motion paths', () => {
  it('leaves an object where its path ended, once that step has played', () => {
    const d = newDeck();
    const s = newSlide('blank', d);
    const a = newShape('rect', 0, 0);
    a.motion = { effect: 'right', order: 1, duration: 1 };
    s.elements.push(a);
    d.slides = [s];
    const p = new Player(d, document.createElement('div'));
    const at = (step) => p.makeLayer(0, step).querySelector(`[data-id="${a.id}"]`).style.translate;
    expect(at(0)).toBe('');
    expect(at(1)).toBe('240px 0px'); // a quarter of the 960 pt slide
  });
});

describe('links in the slideshow', () => {
  it('finds where a link to another slide goes', async () => {
    const { linkTarget } = await import('../../src/present/present.js');
    const d = deckWithAnimations(); // slide 2 is hidden
    expect(linkTarget(d, 'next', { i: 0, step: 0 })).toBe('next');
    expect(linkTarget(d, 'last', { i: 0, step: 0 })).toEqual({ i: 3, step: 0 });
    expect(linkTarget(d, d.slides[1].id, { i: 0, step: 0 })).toEqual({ i: 1, step: 0 });
    expect(linkTarget(d, 'nope', { i: 0, step: 0 })).toBeNull();
  });
});

describe('finding your way in the slideshow', () => {
  const key = (root, k) => root.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true }));

  it('goes to the slide with the number typed, hidden ones too, and counts slides as the filmstrip does', () => {
    const d = deckWithAnimations(); // four slides, the third hidden
    const show = startSlideshow(d, { start: 0 });
    const root = document.querySelector('.show');
    expect(root.querySelector('.show-count').textContent).toBe('1 / 4');
    key(root, '4');
    key(root, 'Enter');
    expect(show.pos).toEqual({ i: 3, step: 0 });
    expect(root.querySelector('.show-count').textContent).toBe('4 / 4');
    key(root, '3');
    key(root, 'Enter');
    expect(show.pos).toEqual({ i: 2, step: 0 });
    key(root, '9');
    key(root, 'Enter');
    expect(show.pos.i).toBe(3);
    show.exit();
  });

  it('moves on and back with a swipe, without the tap that may follow counting too', () => {
    const show = startSlideshow(deckWithAnimations(), { start: 0 });
    const root = document.querySelector('.show');
    const touch = (type, x, y, touches) => {
      const e = new Event(type, { bubbles: true });
      Object.assign(e, { touches: touches ? [{ clientX: x, clientY: y }] : [], changedTouches: [{ clientX: x, clientY: y }] });
      root.dispatchEvent(e);
    };
    const swipe = (dx, dy = 0) => {
      touch('touchstart', 500, 300, true);
      touch('touchend', 500 + dx, 300 + dy, false);
    };
    swipe(-120);
    expect(show.pos).toEqual({ i: 1, step: 0 });
    root.click(); // sent by the browser for the same touch
    expect(show.pos).toEqual({ i: 1, step: 0 });
    swipe(-30); // too short
    swipe(-80, 200); // mostly up and down
    expect(show.pos).toEqual({ i: 1, step: 0 });
    swipe(150);
    expect(show.pos).toEqual({ i: 0, step: 0 });
    show.exit();
  });

  it('doesn’t start a slide’s sound once the show has moved past it', async () => {
    vi.useFakeTimers();
    const play = vi.fn(() => Promise.resolve());
    const proto = globalThis.HTMLMediaElement.prototype;
    const original = Object.getOwnPropertyDescriptor(proto, 'play');
    Object.defineProperty(proto, 'play', { value: play, configurable: true, writable: true });
    Element.prototype.animate = () => ({ finished: new Promise(() => {}), finish() {}, cancel() {} });
    try {
      const d = newDeck();
      const s = newSlide('blank', d);
      s.transition = { type: 'fade', duration: 1 };
      d.media.snd = 'data:audio/wav;base64,UklGRg==';
      s.elements.push({ id: 'snd', type: 'media', kind: 'audio', media: 'snd', x: 0, y: 0, w: 50, h: 50, rot: 0, autoplay: true });
      const after = newSlide('blank', d);
      after.transition = { type: 'fade', duration: 1 }; // the sound's slide is still there as it fades in
      d.slides.push(s, after);
      const host = document.createElement('div');
      document.body.append(host);
      const player = new Player(d, host);
      player.go({ i: 0, step: 0 });
      player.go({ i: 1, step: 0 }); // the sound starts once the transition is over…
      player.go({ i: 2, step: 0 }); // …but the show moves on first
      await vi.advanceTimersByTimeAsync(2000);
      expect(play).not.toHaveBeenCalled();
      player.go({ i: 1, step: 0 }, 'cut');
      await vi.advanceTimersByTimeAsync(10);
      expect(play).toHaveBeenCalledTimes(1);
    } finally {
      vi.useRealTimers();
      if (original) Object.defineProperty(proto, 'play', original);
      else delete proto.play;
    }
  });
});
