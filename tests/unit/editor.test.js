import { describe, it, expect, vi, afterEach } from 'vitest';

vi.mock('../../src/storage/db.js', () => ({
  getDeck: async () => null,
  saveDeck: async (deck) => ({ rev: (deck.rev || 0) + 1 }),
  onStorageChange: () => () => {},
}));

import { EditorScreen } from '../../src/editor/editor-screen.js';
import { newDeck, newSlide, newShape, newTextBox } from '../../src/model/deck.js';

// jsdom lacks these.
globalThis.ResizeObserver ||= class { observe() {} disconnect() {} unobserve() {} };
globalThis.matchMedia ||= () => ({ matches: false, addEventListener() {}, removeEventListener() {} });
globalThis.CSS ||= { escape: (s) => String(s).replace(/["\\]/g, '\\$&') };
Element.prototype.scrollIntoView ||= () => {};
document.queryCommandState ||= () => false;
document.execCommand ||= () => false;

let editor;
function openEditor(deck) {
  const root = document.createElement('div');
  document.body.append(root);
  editor = new EditorScreen(root, { deckId: deck.id });
  editor.deck = deck;
  editor.slideSelection = new Set([deck.slides[0].id]);
  editor.build();
  editor.goToSlide(0);
  editor.filmstrip.render();
  return editor;
}

function deckWith(...elements) {
  const d = newDeck();
  d.slides[0] = newSlide('blank', d);
  d.slides[0].elements.push(...elements);
  d.slides.push(newSlide('titleContent', d));
  return d;
}

/** Type into the text box being edited, as the browser would. */
function type(ed, html) {
  const body = ed.canvas.editing.bodies[0];
  body.innerHTML = html;
  body.dispatchEvent(new Event('input'));
}

const key = (target, init) => {
  const e = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init });
  target.dispatchEvent(e);
  return e;
};

afterEach(async () => {
  await editor?.destroy({ save: false });
  editor = null;
  document.body.replaceChildren();
});

describe('editor: speaker notes', () => {
  it('loads the new slide’s notes when the slide changes while the notes box has focus', () => {
    const d = deckWith();
    d.slides[0].notes = 'one';
    d.slides[1].notes = 'two';
    const ed = openEditor(d);
    ed.notes.focus();
    expect(ed.notes.value).toBe('one');
    ed.goToSlide(1);
    expect(ed.notes.value).toBe('two');
    ed.notes.value += '!';
    ed.notes.dispatchEvent(new Event('input'));
    expect(d.slides.map((s) => s.notes)).toEqual(['one', 'two!']);
  });
});

describe('editor: undo', () => {
  it('undoes text typed before a formatting change as its own, earlier step', () => {
    const box = newTextBox(10, 10);
    box.text = '<p>old</p>';
    const ed = openEditor(deckWith(box));
    ed.setSelection([box.id]);
    ed.canvas.startEdit(box.id);
    type(ed, '<p>new</p>');
    ed.setElementProp('fill', '#ff0000');
    ed.canvas.endEdit();
    ed.undo();
    expect([ed.slide.elements[0].text, ed.slide.elements[0].fill]).toEqual(['<p>new</p>', 'none']);
    ed.undo();
    expect(ed.slide.elements[0].text).toBe('<p>old</p>');
    expect(ed.history.canUndo).toBe(false);
  });

  it('records nothing for a change that changes nothing', () => {
    const ed = openEditor(deckWith(newShape('rect', 0, 0)));
    ed.change(() => {});
    ed.setElementProp('fill', 'theme:accent1'); // nothing selected
    expect(ed.history.canUndo).toBe(false);
  });
});

describe('editor: changing slides while typing', () => {
  it('applies a layout to a slide whose placeholder is being edited', () => {
    const d = deckWith();
    const ed = openEditor(d);
    ed.goToSlide(1);
    const title = ed.slide.elements.find((e) => e.ph === 'title');
    ed.canvas.startEdit(title.id);
    type(ed, '<p>Hello</p>');
    ed.applyLayout('title');
    expect(ed.canvas.editing).toBeNull();
    expect(ed.slide.elements.find((e) => e.ph === 'title').text).toBe('<p>Hello</p>');
    // No leftover node for the replaced placeholder.
    expect(ed.canvas.slideEl.querySelectorAll(':scope > .el')).toHaveLength(ed.slide.elements.length);
  });
});

describe('editor: animations', () => {
  it('doesn’t animate an object when only the duration or start changes', () => {
    const a = newShape('rect', 0, 0);
    const b = newShape('rect', 200, 0);
    b.anim = { effect: 'zoom', duration: 0.5, start: 'click', order: 1 };
    const ed = openEditor(deckWith(a, b));
    ed.setSelection([a.id, b.id]);
    ed.setAnimation({ duration: 1.2 });
    expect(a.anim).toBeUndefined();
    expect(ed.slide.elements[1].anim).toMatchObject({ effect: 'zoom', duration: 1.2 });
    ed.setAnimation({ effect: 'fade' });
    expect(ed.slide.elements[0].anim).toMatchObject({ effect: 'fade', order: 2 });
  });
});

describe('editor: side panel', () => {
  it('shows the controls for a background fill chosen from the focused list', () => {
    const ed = openEditor(deckWith());
    ed.openPanel('design');
    const fill = ed.panel.body.querySelector('select');
    fill.focus();
    fill.value = 'solid';
    fill.dispatchEvent(new Event('change'));
    expect(ed.slide.background?.type).toBe('solid');
    expect([...ed.panel.body.querySelectorAll('.pf-color')].map((b) => b.textContent)).toContain('Colour');
  });

  it('makes one undo step for Enter in a number field, and updates the locked other side', () => {
    const shape = { ...newShape('rect', 0, 0, 200, 100), lockRatio: true };
    const ed = openEditor(deckWith(shape));
    ed.setSelection([shape.id]);
    ed.openPanel('format');
    const width = ed.panel.body.querySelector('input[type=number]');
    width.focus();
    width.value = '400';
    key(width, { key: 'Enter' });
    document.activeElement.dispatchEvent(new Event('change'));
    width.dispatchEvent(new Event('change'));
    expect(ed.slide.elements[0]).toMatchObject({ w: 400, h: 200 });
    const [, height] = ed.panel.body.querySelectorAll('input[type=number]');
    expect(height.value).toBe('200');
    ed.undo();
    expect(ed.slide.elements[0]).toMatchObject({ w: 200, h: 100 });
    expect(ed.history.canUndo).toBe(false);
  });
});

describe('editor: keyboard', () => {
  it('undoes with Ctrl+Z after using a checkbox in the side panel', () => {
    const shape = newShape('rect', 0, 0);
    const ed = openEditor(deckWith(shape));
    ed.setSelection([shape.id]);
    ed.openPanel('format');
    const shadow = [...ed.panel.body.querySelectorAll('input[type=checkbox]')].find((c) => c.parentElement.textContent === 'Drop shadow');
    shadow.focus();
    shadow.checked = true;
    shadow.dispatchEvent(new Event('change'));
    expect(ed.slide.elements[0].shadow).toBe(true);
    const e = key(document.activeElement, { key: 'z', ctrlKey: true });
    expect(e.defaultPrevented).toBe(true);
    expect(ed.slide.elements[0].shadow).toBeFalsy();
  });

  it('leaves Delete and Space to a focused title-bar button', () => {
    const shape = newShape('rect', 0, 0);
    const ed = openEditor(deckWith(shape));
    ed.setSelection([shape.id]);
    const btn = ed.el.querySelector('.titlebar-right .icon-btn');
    btn.focus();
    expect(key(btn, { key: 'Delete' }).defaultPrevented).toBe(false);
    expect(key(btn, { key: ' ' }).defaultPrevented).toBe(false);
    expect(ed.slide.elements).toHaveLength(1);
  });

  it('types Alt+Shift+5 in the notes instead of striking through the selected box', () => {
    const box = newTextBox(0, 0);
    const ed = openEditor(deckWith(box));
    ed.setSelection([box.id]);
    ed.notes.focus();
    const e = key(ed.notes, { key: '%', code: 'Digit5', altKey: true, shiftKey: true });
    expect(e.defaultPrevented).toBe(false);
    expect(ed.slide.elements[0].style?.strike).toBeUndefined();
  });

  it('starts one slideshow when F5 repeats while it loads', async () => {
    const ed = openEditor(deckWith());
    await Promise.all([ed.present(), ed.present(), ed.present({ fromStart: true })]);
    expect(document.querySelectorAll('.show')).toHaveLength(1);
    ed.presentation.exit();
  });
});

describe('editor: charts', () => {
  it('inserts a chart and edits its data in the format panel, one undo step at a time', () => {
    const ed = openEditor(deckWith());
    ed.insertChart();
    const chart = () => ed.slide.elements.find((e) => e.type === 'chart');
    expect(ed.selection).toEqual([chart().id]);
    expect(ed.panel.current).toBe('format');
    const panel = ed.panel.el;
    const cell = panel.querySelector('[aria-label="Series 1, Category 2"]');
    cell.value = '9';
    cell.dispatchEvent(new Event('change'));
    expect(chart().chart.series[0].values[1]).toBe(9);
    [...panel.querySelectorAll('button')].find((b) => b.textContent === 'Series').click();
    expect(chart().chart.series).toHaveLength(4);
    const kind = [...panel.querySelectorAll('select')].find((s) => [...s.options].some((o) => o.value === 'doughnut'));
    kind.value = 'pie';
    kind.dispatchEvent(new Event('change'));
    expect(chart().chart.kind).toBe('pie');
    expect(ed.canvas.stage.querySelectorAll('.chart-slice').length).toBeGreaterThan(0);
    ed.undo();
    ed.undo();
    expect(chart().chart.kind).toBe('column');
    expect(chart().chart.series).toHaveLength(3);
    expect(chart().chart.series[0].values[1]).toBe(9);
  });
});

describe('editor: groups', () => {
  it('groups objects together in the stacking order, and copies get groups of their own', () => {
    const [a, b, c] = [newShape('rect', 0, 0), newShape('ellipse', 200, 0), newShape('rect', 400, 0)];
    const ed = openEditor(deckWith(a, b, c));
    ed.setSelection([a.id, c.id]);
    key(ed.canvas.wrap, { key: 'g', ctrlKey: true });
    // The group sits where its topmost member was: b, then a and c.
    expect(ed.slide.elements.map((e) => e.id)).toEqual([b.id, a.id, c.id]);
    expect(a.groupId).toBeTruthy();
    expect(c.groupId).toBe(a.groupId);
    ed.duplicateSelection();
    const copies = ed.slide.elements.slice(3);
    expect(copies).toHaveLength(2);
    expect(copies[0].groupId).toBe(copies[1].groupId);
    expect(copies[0].groupId).not.toBe(a.groupId);
    ed.setSelection([a.id, c.id]);
    key(ed.canvas.wrap, { key: 'G', ctrlKey: true, shiftKey: true });
    expect(a.groupId).toBeUndefined();
    ed.undo();
    expect(ed.slide.elements.find((e) => e.id === a.id).groupId).toBeTruthy();
  });

  it('selects a whole group with a click, and one of its objects with another', () => {
    const [a, b] = [newShape('rect', 0, 0), newShape('ellipse', 200, 0)];
    a.groupId = 'g1';
    b.groupId = 'g1';
    const ed = openEditor(deckWith(a, b));
    const click = () => {
      const node = ed.canvas.slideEl.querySelector(`[data-id="${a.id}"]`);
      node.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, button: 0, clientX: 10, clientY: 10, pointerId: 1 }));
      window.dispatchEvent(new PointerEvent('pointerup', { bubbles: true, button: 0, clientX: 10, clientY: 10, pointerId: 1 }));
      document.dispatchEvent(new PointerEvent('pointerup', { bubbles: true, button: 0, clientX: 10, clientY: 10, pointerId: 1 }));
    };
    click();
    expect([...ed.selection].sort()).toEqual([a.id, b.id].sort());
    click();
    expect(ed.selection).toEqual([a.id]);
  });
});
