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
