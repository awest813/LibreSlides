import { describe, it, expect } from 'vitest';
import PRESETS from '../../src/model/preset-shapes.js';
import { evaluateGeometry, guides } from '../../src/model/geometry.js';

describe('shape geometry', () => {
  it('works out every PowerPoint preset shape', () => {
    expect(Object.keys(PRESETS).length).toBeGreaterThan(180);
    for (const [name, def] of Object.entries(PRESETS)) {
      for (const [w, h] of [[200, 100], [100, 300]]) {
        const { paths, textRect } = evaluateGeometry(def, w * 12700, h * 12700);
        expect(paths.length, name).toBeGreaterThan(0);
        for (const p of paths) for (const [, ...n] of p.cmds) for (const v of n) expect(Number.isFinite(v), name).toBe(true);
        if (textRect) for (const v of Object.values(textRect)) expect(Number.isFinite(v), name).toBe(true);
      }
    }
  });

  it('evaluates formulas and adjust values', () => {
    const def = { av: [['adj', 'val 25000']], gd: [['x1', '*/ ss adj 100000'], ['x2', '+- r 0 x1'], ['a', 'at2 1 1'], ['p', 'pin 0 x2 50']] };
    const { g } = guides(def, 200, 100);
    expect([g.x1, g.x2, g.a, g.p]).toEqual([25, 175, 2700000, 50]);
    expect(guides(def, 200, 100, { adj: 50000 }).g.x1).toBe(50);
  });

  it("draws a rounded rectangle's corners from its adjust value", () => {
    const { paths, textRect } = evaluateGeometry(PRESETS.roundRect, 200, 100, { adj: 50000 });
    // Corner radius: half the short side, so the ends are semicircles.
    expect(paths[0].cmds[0]).toEqual(['M', 0, 0.5]);
    expect(textRect.l).toBeGreaterThan(0.07);
  });
});
