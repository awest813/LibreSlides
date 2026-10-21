import { describe, it, expect } from 'vitest';
import {
  slideSections, moveSlidesWithSections, removeSlidesWithSections, setSections, newTable, newDiagram, DIAGRAM_LAYOUTS,
  gatherGroups, normalizeDeck,
} from '../../src/model/deck.js';
import { insertRow, insertCol, deleteRow, deleteCol, mergeCells, unmergeCell, normalizeTable } from '../../src/model/tables.js';
import { layoutDiagram } from '../../src/model/diagram-layout.js';
import { evaluateGeometry } from '../../src/model/geometry.js';
import PRESETS from '../../src/model/preset-shapes.js';
import { WARPS } from '../../src/model/warps.js';
import { placeCharacters } from '../../src/render/warp.js';

/**
 * Property-style checks of the model's invariants: many random cases from a
 * seeded generator (so a failure repeats), each checked against what must
 * always hold.
 */

/** A small seeded random number generator (mulberry32). */
function rng(seed) {
  let a = seed >>> 0;
  const next = () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
  next.int = (n) => Math.floor(next() * n);
  next.pick = (list) => list[next.int(list.length)];
  return next;
}

const finite = (v) => typeof v === 'number' && Number.isFinite(v);

// ------------------------------------------------------------------ sections

describe('sections survive moves and deletes', () => {
  /** Slides s0…, some starting sections with names of their own. */
  const randomSlides = (r) => {
    const n = 1 + r.int(9);
    let k = 0;
    return Array.from({ length: n }, (_, i) => ({ id: `s${i}`, ...(r() < 0.35 && { section: `S${k++}` }) }));
  };
  const sectionOf = (slides) => {
    const out = new Map();
    for (const g of slideSections(slides)) for (let i = g.start; i < g.end; i++) out.set(slides[i].id, g.name);
    return out;
  };
  const randomIds = (r, slides, { notAll = false } = {}) => {
    const ids = new Set(slides.filter(() => r() < 0.4).map((s) => s.id));
    if (!ids.size) ids.add(r.pick(slides).id);
    if (notAll && ids.size === slides.length) ids.delete(slides[0].id);
    return ids;
  };

  it('keeps the slides that stay in their sections when others are deleted', () => {
    const r = rng(1);
    for (let n = 0; n < 500; n++) {
      const slides = randomSlides(r);
      if (slides.length < 2) continue;
      const ids = randomIds(r, slides, { notAll: true });
      const before = sectionOf(slides);
      const rest = removeSlidesWithSections(slides.map((s) => ({ ...s })), ids);
      const after = sectionOf(rest);
      for (const s of rest) expect(after.get(s.id)).toBe(before.get(s.id));
      // Every name appears once, in the order it had.
      const names = rest.map((s) => s.section).filter(Boolean);
      expect(names).toEqual([...new Set(names)]);
      expect(names).toEqual(slides.map((s) => s.section).filter((x) => names.includes(x)));
    }
  });

  it('keeps every section, and the slides that stay in theirs, when slides move', () => {
    const r = rng(2);
    for (let n = 0; n < 1000; n++) {
      const slides = randomSlides(r);
      const ids = randomIds(r, slides);
      const index = r.int(slides.length + 1);
      const before = sectionOf(slides);
      const copy = slides.map((s) => ({ ...s }));
      const m = moveSlidesWithSections(copy, ids, index);
      if (!m) continue;
      setSections(m.order, m.sections);
      expect(m.order.map((s) => s.id).sort()).toEqual(slides.map((s) => s.id).sort());
      const after = sectionOf(m.order);
      for (const s of m.order) if (!ids.has(s.id)) expect(after.get(s.id)).toBe(before.get(s.id));
      // No section is lost or made twice, and slides in no section stay first.
      const names = m.order.map((s) => s.section).filter(Boolean);
      expect(names.sort()).toEqual(slides.map((s) => s.section).filter(Boolean).sort());
      const firstNamed = m.order.findIndex((s) => s.section);
      if (firstNamed > 0) for (const s of m.order.slice(0, firstNamed)) expect(before.get(s.id) == null || ids.has(s.id)).toBe(true);
    }
  });
});

// ------------------------------------------------------------------ groups

describe('groups stay together', () => {
  const contiguous = (els) => {
    const seen = new Set();
    let prev = null;
    for (const e of els) {
      if (e.groupId && e.groupId !== prev && seen.has(e.groupId)) return false;
      if (e.groupId) seen.add(e.groupId);
      prev = e.groupId || null;
    }
    return true;
  };
  const order = (els, keep) => els.filter(keep).map((e) => e.id).join(' ');

  it('gathers each group where its topmost member is, keeping everything else in order', () => {
    const r = rng(5);
    for (let n = 0; n < 500; n++) {
      const els = Array.from({ length: r.int(12) }, (_, i) => ({ id: `e${i}`, ...(r() < 0.6 && { groupId: r.pick(['A', 'B', 'C']) }) }));
      const out = gatherGroups(els);
      expect(out.map((e) => e.id).sort()).toEqual(els.map((e) => e.id).sort());
      expect(contiguous(out)).toBe(true);
      expect(order(out, (e) => !e.groupId)).toBe(order(els, (e) => !e.groupId));
      for (const g of ['A', 'B', 'C']) expect(order(out, (e) => e.groupId === g)).toBe(order(els, (e) => e.groupId === g));
      if (contiguous(els)) expect(out.map((e) => e.id)).toEqual(els.map((e) => e.id));
      // The topmost member stays where it was among the ungrouped elements before and after it.
      for (const g of ['A', 'B', 'C']) {
        const last = els.map((e) => e.groupId).lastIndexOf(g);
        if (last < 0) continue;
        const below = els.slice(0, last).filter((e) => !e.groupId).length;
        expect(out.slice(0, out.findIndex((e) => e.groupId === g)).filter((e) => !e.groupId)).toHaveLength(below);
      }
    }
  });

  it('when a deck is opened', () => {
    const deck = normalizeDeck({ slides: [{ id: 's', elements: [{ id: 'a', groupId: 'g' }, { id: 'x' }, { id: 'b', groupId: 'g' }] }] });
    expect(deck.slides[0].elements.map((e) => e.id)).toEqual(['x', 'a', 'b']);
  });
});

// ------------------------------------------------------------------ tables

/** What's wrong with a table, if anything: it should be a full grid, each cell covered by exactly one block that fits it. */
function tableProblems(el) {
  const rows = el.cells.length;
  const cols = el.cells[0]?.length || 0;
  const out = [];
  const sum = (list) => list.reduce((s, v) => s + v, 0);
  if (!rows || !cols) out.push('empty');
  if (el.cells.some((row) => row.length !== cols)) out.push('ragged rows');
  if (el.colWidths.length !== cols || Math.abs(sum(el.colWidths) - 1) > 1e-6) out.push('column widths');
  if (el.rowHeights && (el.rowHeights.length !== rows || Math.abs(sum(el.rowHeights) - 1) > 1e-6)) out.push('row heights');
  const owner = el.cells.map((row) => row.map(() => 0));
  el.cells.forEach((row, r) => row.forEach((cell, c) => {
    if (cell.merged) {
      if (cell.rowSpan != null || cell.colSpan != null) out.push(`covered cell ${r},${c} has a span`);
      return;
    }
    const rs = cell.rowSpan || 1;
    const cs = cell.colSpan || 1;
    if (r + rs > rows || c + cs > cols) out.push(`block ${r},${c} runs off the table`);
    for (let i = r; i < Math.min(rows, r + rs); i++) {
      for (let j = c; j < Math.min(cols, c + cs); j++) {
        owner[i][j]++;
        if ((i !== r || j !== c) && !el.cells[i][j].merged) out.push(`${i},${j} in block ${r},${c} isn't covered`);
      }
    }
  }));
  owner.forEach((row, i) => row.forEach((n, j) => { if (n !== 1) out.push(`${i},${j} is in ${n} blocks`); }));
  return out;
}

describe('tables stay normal', () => {
  it('after any mix of merges, splits, inserts and deletes', () => {
    const r = rng(3);
    for (let n = 0; n < 150; n++) {
      const el = newTable(1 + r.int(5), 1 + r.int(5), 0, 0, 400, 200);
      if (r() < 0.5) el.rowHeights = el.cells.map(() => 1 / el.cells.length);
      for (let k = 0; k < 25; k++) {
        const rows = el.cells.length;
        const cols = el.cells[0].length;
        const [i, j] = [r.int(rows), r.int(cols)];
        switch (r.int(7)) {
          case 0: mergeCells(el, i, j, 'right'); break;
          case 1: mergeCells(el, i, j, 'down'); break;
          case 2: unmergeCell(el, i, j); break;
          case 3: insertRow(el, r.int(rows + 1)); break;
          case 4: insertCol(el, r.int(cols + 1)); break;
          case 5: deleteRow(el, i); break;
          default: deleteCol(el, j);
        }
        const problems = tableProblems(el);
        if (problems.length) expect({ n, k, problems }).toEqual({ n, k, problems: [] });
      }
    }
  });

  it('repairs overlapping or oversized spans', () => {
    const r = rng(4);
    for (let n = 0; n < 200; n++) {
      const el = newTable(1 + r.int(5), 1 + r.int(5), 0, 0, 400, 200);
      for (const row of el.cells) {
        for (const cell of row) {
          if (r() < 0.3) cell.rowSpan = r.int(7) - 1;
          if (r() < 0.3) cell.colSpan = r() < 0.1 ? 'x' : r.int(7);
          if (r() < 0.2) cell.merged = true;
        }
      }
      expect(tableProblems(normalizeTable(el))).toEqual([]);
    }
  });
});

// ------------------------------------------------------------------ diagrams

describe('diagrams lay out sensibly', () => {
  const inside = (el, c) => {
    const pad = 1;
    if (c.type === 'line') {
      for (const [x, y] of [[c.x1, c.y1], [c.x2, c.y2]]) {
        expect(x).toBeGreaterThanOrEqual(-pad);
        expect(y).toBeGreaterThanOrEqual(-pad);
        expect(x).toBeLessThanOrEqual(el.w + pad);
        expect(y).toBeLessThanOrEqual(el.h + pad);
      }
      return;
    }
    for (const k of ['x', 'y', 'w', 'h']) expect(finite(c[k])).toBe(true);
    expect(c.w).toBeGreaterThanOrEqual(0);
    expect(c.h).toBeGreaterThanOrEqual(0);
    // Turned arrows (cycles) may poke out by a little; everything else stays in.
    const slack = c.rot ? Math.max(c.w, c.h) / 2 : pad;
    expect(c.x).toBeGreaterThanOrEqual(-slack);
    expect(c.y).toBeGreaterThanOrEqual(-slack);
    expect(c.x + c.w).toBeLessThanOrEqual(el.w + slack);
    expect(c.y + c.h).toBeLessThanOrEqual(el.h + slack);
    if (c.style) expect(c.style.size).toBeGreaterThan(0);
  };
  const outlines = {
    none: [],
    one: [['Only', 0]],
    fifty: Array.from({ length: 50 }, (_, i) => [`Node ${i}`, i % 3 ? 1 : 0]),
    'fifty top-level': Array.from({ length: 50 }, (_, i) => [`N${i}`, 0]),
    deep: Array.from({ length: 30 }, (_, i) => [`Level ${i % 9}`, i % 9]),
    long: [['A very long label that has many words in it and more besides', 0], ['Supercalifragilisticexpialidocious', 1]],
  };

  for (const [name, list] of Object.entries(outlines)) {
    it(`with ${name === 'none' ? 'no' : name} nodes, in every layout and box shape`, () => {
      for (const { id } of DIAGRAM_LAYOUTS) {
        for (const [w, h] of [[560, 320], [120, 600], [900, 60], [1, 1]]) {
          const el = newDiagram(id, 0, 0, w, h);
          el.diagram.nodes = list.map(([text, level], i) => ({ id: `n${i}`, text, level }));
          const out = layoutDiagram(el);
          if (!list.length) expect(out).toEqual([]);
          for (const c of out) inside(el, c);
        }
      }
    });
  }
});

// ------------------------------------------------------------------ geometry

describe('preset geometry', () => {
  const sizes = [[0, 0], [0, 100], [100, 0], [1, 1], [100, 100], [1e5, 1], [1, 1e5]];
  const values = [undefined, -1e7, 0, 50000, 1e7];
  const check = (defs) => {
    const bad = [];
    for (const [name, def] of Object.entries(defs)) {
      for (const [w, h] of sizes) {
        for (const v of values) {
          const adj = {};
          if (v !== undefined) for (const [name] of def.av || []) adj[name] = v;
          const { paths, textRect } = evaluateGeometry(def, w, h, adj);
          const nums = [...paths.flatMap((p) => p.cmds.flatMap((c) => c.slice(1))), ...Object.values(textRect || {})];
          if (!nums.every(finite)) bad.push(`${name} ${w}×${h} ${v}`);
        }
      }
    }
    expect(bad).toEqual([]);
  };

  it('works out every preset shape, at any size and adjust value, without NaN', () => check(PRESETS));
  it('works out every warp, at any size and adjust value, without NaN', () => check(WARPS));
});

// ------------------------------------------------------------------ warps

describe('warped text of any length', () => {
  const style = { font: 'Inter', size: 24, color: '#000' };
  const line = (text) => [...text].map((ch) => ({ ch, style }));
  const texts = { empty: [], one: [line('A')], spaces: [line('   ')], long: [line('word '.repeat(150))], lines: Array.from({ length: 40 }, () => line('Hi')) };

  for (const [name, lines] of Object.entries(texts)) {
    it(`places ${name} text under every warp`, () => {
      const bad = [];
      for (const preset of Object.keys(WARPS)) {
        for (const [w, h] of [[300, 120], [4, 4], [0, 50]]) {
          const out = placeCharacters({ preset }, lines, { w, h });
          if (out.some((c) => /NaN|Infinity/.test(c.transform))) bad.push(`${preset} ${w}×${h}`);
        }
      }
      expect(bad).toEqual([]);
    });
  }
});
