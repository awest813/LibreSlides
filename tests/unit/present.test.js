import { describe, it, expect } from 'vitest';
import { buildSteps, nextPos, prevPos, firstPos } from '../../src/present/present.js';
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
