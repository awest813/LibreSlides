import { describe, it, expect, vi, afterEach } from 'vitest';

vi.mock('../../src/storage/db.js', () => ({
  getDeck: async () => null,
  saveDeck: async (deck) => ({ rev: (deck.rev || 0) + 1 }),
  onStorageChange: () => () => {},
}));

import { EditorScreen, stepOrder } from '../../src/editor/editor-screen.js';
import { newDeck, newSlide, newShape, newTextBox, newTable } from '../../src/model/deck.js';

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

describe('editor: Edit menu while typing', () => {
  it('cuts the selected text, not the whole box being edited', () => {
    const box = newTextBox(10, 10);
    box.text = '<p>Hello world</p>';
    const ed = openEditor(deckWith(box));
    ed.setSelection([box.id]);
    ed.canvas.startEdit(box.id);
    const text = ed.canvas.editing.bodies[0].querySelector('p').firstChild;
    const range = document.createRange();
    range.setStart(text, 0);
    range.setEnd(text, 5);
    ed.canvas.savedRange = range;
    const writeText = vi.fn(async () => {});
    Object.defineProperty(navigator, 'clipboard', { value: { writeText }, configurable: true });
    const exec = vi.spyOn(document, 'execCommand');
    // The menu offers Cut for the selected text.
    ed.toolbar.menuButtons.Edit.click();
    const cut = [...document.querySelectorAll('.popover .menu-item')].find((b) => b.textContent.startsWith('Cut'));
    expect(cut.disabled).toBe(false);
    cut.click();
    expect(ed.slide.elements).toHaveLength(1);
    expect(ed.canvas.editing?.id).toBe(box.id);
    expect(exec).toHaveBeenCalledWith('cut');
    // jsdom can't cut: the text goes to the clipboard and is deleted.
    expect(writeText).toHaveBeenCalledWith('Hello');
    expect(exec).toHaveBeenCalledWith('delete');
    exec.mockRestore();
    delete navigator.clipboard;
  });

  it('deletes the selected objects from the menu when nothing is being typed', () => {
    const shape = newShape('rect', 0, 0);
    const ed = openEditor(deckWith(shape));
    ed.setSelection([shape.id]);
    ed.editCommand('delete');
    expect(ed.slide.elements).toHaveLength(0);
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

  it('gives the cell being typed in borders and margins from the format panel, or the whole table', () => {
    const table = newTable(2, 2, 0, 0, 200, 80);
    const ed = openEditor(deckWith(table));
    ed.setSelection([table.id]);
    ed.openPanel('format');
    const field = (label) => [...ed.panel.body.querySelectorAll('label.pf')].find((l) => l.firstChild.textContent === label).querySelector('select, input');
    const borders = field('Borders');
    borders.value = 'outside';
    borders.dispatchEvent(new Event('change'));
    const t = () => ed.slide.elements[0];
    expect(t().cells[0][0].borders).toMatchObject({ t: { width: 1 }, l: { width: 1 } });
    expect(t().cells[0][0].borders.r).toBeUndefined();
    expect(field('Borders').value).toBe(''); // ready for the next choice
    // Typing in the bottom-right cell: only it changes.
    ed.canvas.startEdit(table.id);
    const body = ed.canvas.editing.bodies[3];
    const range = document.createRange();
    range.selectNodeContents(body);
    ed.canvas.savedRange = range;
    const left = field('Left margin');
    left.value = '20';
    left.dispatchEvent(new Event('change'));
    expect(t().cells[1][1].margin).toMatchObject({ l: 20 });
    expect(t().cells[0][0].margin).toBeUndefined();
    ed.undo();
    expect(t().cells[1][1].margin).toBeUndefined();
  });

  it('gives a borderless box a visible border when its border weight is set', () => {
    const box = newTextBox(0, 0);
    box.stroke = { color: 'none', width: 2 };
    const ed = openEditor(deckWith(box));
    ed.setSelection([box.id]);
    ed.openPanel('format');
    const weight = [...ed.panel.body.querySelectorAll('label.pf')].find((l) => l.firstChild.textContent === 'Border weight').querySelector('input');
    expect(weight.value).toBe('0'); // no border shows
    expect(ed.panel.body.querySelector('.pf-color .pf-swatch.is-none')).toBeTruthy();
    weight.value = '3';
    weight.dispatchEvent(new Event('change'));
    expect(ed.slide.elements[0].stroke).toMatchObject({ color: 'theme:text', width: 3 });
    ed.undo();
    expect(ed.slide.elements[0].stroke).toEqual({ color: 'none', width: 2 });
  });

  it('shows an animation’s own duration when it isn’t one of the choices', () => {
    const shape = newShape('rect', 0, 0);
    shape.anim = { effect: 'fade', duration: 1, start: 'click', order: 1 };
    const ed = openEditor(deckWith(shape));
    ed.setSelection([shape.id]);
    ed.openPanel('animate');
    const duration = [...ed.panel.body.querySelectorAll('label.pf')].find((l) => l.firstChild.textContent === 'Duration').querySelector('select');
    expect(duration.selectedOptions[0].textContent).toBe('1 s');
    const labels = [...duration.options].map((o) => o.textContent);
    expect(new Set(labels).size).toBe(labels.length);
  });

  it('names every control in the panel', () => {
    const table = newTable(2, 2, 0, 0, 200, 80);
    const ed = openEditor(deckWith(table));
    const unnamed = () => [...ed.panel.body.querySelectorAll('input, select, textarea, button')].filter((c) => !c.getAttribute('aria-label') && !c.closest('label') && !c.textContent.trim());
    ed.openPanel('design');
    expect(unnamed()).toEqual([]);
    ed.setSelection([table.id]);
    ed.openPanel('format');
    expect(unnamed()).toEqual([]);
    ed.insertChart();
    ed.updateSelected((e) => { e.chart.series[1].kind = 'line'; });
    expect(unnamed()).toEqual([]);
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

  it('moves through the menu bar with the arrow keys, and between open menus', () => {
    const ed = openEditor(deckWith());
    const items = [...ed.toolbar.menubar.querySelectorAll('.menubar-item')];
    expect(items.filter((b) => b.tabIndex === 0)).toEqual([items[0]]); // one tab stop
    items[0].focus();
    key(items[0], { key: 'ArrowRight' });
    expect(document.activeElement).toBe(items[1]);
    key(items[1], { key: 'ArrowDown' });
    expect(items[1].getAttribute('aria-expanded')).toBe('true');
    expect(document.activeElement.closest('.popover')).toBeTruthy();
    key(document.activeElement, { key: 'ArrowRight' });
    expect(items[2].getAttribute('aria-expanded')).toBe('true');
    expect(items[1].getAttribute('aria-expanded')).toBe('false');
    // Tab leaves the menu from its button.
    key(document.activeElement, { key: 'Tab' });
    expect(document.querySelector('.popover')).toBeNull();
    expect(document.activeElement).toBe(items[2]);
    expect(items[2].tabIndex).toBe(0);
  });

  it('makes the toolbar one tab stop that the arrow keys move along, skipping hidden controls', () => {
    const box = newTextBox(0, 0);
    const ed = openEditor(deckWith(box));
    const stops = () => [...ed.toolbar.toolbar.querySelectorAll('button')].filter((b) => b.tabIndex === 0);
    expect(stops()).toHaveLength(1);
    const first = stops()[0];
    first.focus();
    key(first, { key: 'End' });
    const last = document.activeElement;
    expect(last.closest('[hidden]')).toBeNull();
    expect(stops()).toEqual([last]);
    // Selecting text shows the text controls, with Bold reported as on or off.
    ed.setSelection([box.id]);
    expect(ed.toolbar.boldBtn.getAttribute('aria-pressed')).toBe('false');
    ed.toggleTextProp('bold');
    expect(ed.toolbar.boldBtn.getAttribute('aria-pressed')).toBe('true');
    // Hiding the focused control's group moves the tab stop to a visible one.
    ed.setSelection([]);
    expect(stops()).toHaveLength(1);
    expect(stops()[0].closest('[hidden]')).toBeNull();
  });

  it('opens the side panel’s tabs with the arrow keys', () => {
    const ed = openEditor(deckWith());
    ed.openPanel('design');
    const tab = ed.panel.tabs.querySelector('[aria-selected="true"]');
    expect(tab.tabIndex).toBe(0);
    tab.focus();
    key(tab, { key: 'ArrowRight' });
    expect(ed.panel.current).toBe('transition');
    expect(document.activeElement.dataset.tab).toBe('transition');
    expect(ed.panel.body.getAttribute('role')).toBe('tabpanel');
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

describe('editor: tables', () => {
  const menuItem = (label) => [...document.querySelectorAll('.popover .menu-item')].find((b) => b.querySelector('.menu-label').textContent === label);

  it('inserts a row at the right-clicked cell from the slide’s context menu', () => {
    const table = newTable(2, 2, 0, 0, 200, 80);
    table.cells[0][0].text = '<p>a</p>';
    table.cells[1][0].text = '<p>b</p>';
    const ed = openEditor(deckWith(table));
    const td = ed.canvas.slideEl.querySelector('td[data-r="0"][data-c="0"]');
    td.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 5, clientY: 5 }));
    expect(ed.selection).toEqual([table.id]);
    menuItem('Insert row below').click();
    const t = ed.slide.elements[0];
    expect(t.cells.map((row) => row[0].text)).toEqual(['<p>a</p>', '', '<p>b</p>']);
    ed.undo();
    expect(ed.slide.elements[0].cells).toHaveLength(2);
  });

  it('keeps row and column actions in one toolbar menu', () => {
    const table = newTable(1, 2, 0, 0, 200, 40);
    const ed = openEditor(deckWith(table));
    ed.setSelection([table.id]);
    const group = ed.toolbar.groups.table;
    expect(group.hidden).toBe(false);
    expect(group.querySelectorAll('button')).toHaveLength(2);
    // Lists and indents do nothing to a table until its text is typed in.
    expect([ed.toolbar.bulletBtn.disabled, ed.toolbar.indentBtn.disabled]).toEqual([true, true]);
    ed.canvas.startEdit(table.id);
    expect([ed.toolbar.bulletBtn.disabled, ed.toolbar.indentBtn.disabled]).toEqual([false, false]);
    ed.canvas.endEdit();
    group.querySelector('[aria-label="Rows and columns"]').click();
    expect(menuItem('Delete row').disabled).toBe(true); // the only row
    menuItem('Insert column right').click();
    expect(ed.slide.elements[0].cells[0]).toHaveLength(3);
  });

  it('chooses a table size with the arrow keys', () => {
    const ed = openEditor(deckWith());
    ed.toolbar.tableBtn.click();
    const first = document.querySelector('.table-grid button');
    first.focus();
    key(first, { key: 'ArrowRight' });
    key(document.activeElement, { key: 'ArrowDown' });
    expect(document.activeElement.getAttribute('aria-label')).toBe('2 by 2 table');
    document.activeElement.click();
    expect(ed.slide.elements.at(-1).cells.map((r) => r.length)).toEqual([2, 2]);
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

describe('editor: fields and text effects', () => {
  it('inserts a slide number that shows each slide’s own number', () => {
    const d = deckWith();
    const ed = openEditor(d);
    ed.goToSlide(1);
    ed.insertField('slidenum');
    const box = ed.slide.elements.at(-1);
    expect(box.text).toBe('<p><span data-field="slidenum">2</span></p>');
    expect(ed.canvas.slideEl.querySelector('[data-field=slidenum]').textContent).toBe('2');
    d.slides.reverse();
    ed.goToSlide(0);
    expect(ed.canvas.slideEl.querySelector('[data-field=slidenum]').textContent).toBe('1');
  });

  it('gives the selected boxes a text effect, and takes it away', () => {
    const box = newTextBox(0, 0);
    box.text = '<p>One <b>two</b></p><ul><li>three</li></ul>';
    const ed = openEditor(deckWith(box));
    ed.setSelection([box.id]);
    ed.setTextEffect('outline');
    expect(ed.slide.elements[0].text).toMatch(/^<p><span style="-webkit-text-stroke: 1px [^"]+">One <b>two<\/b><\/span><\/p><ul><li><span style="-webkit-text-stroke/);
    ed.setTextEffect('shadow');
    expect(ed.slide.elements[0].text).not.toContain('stroke');
    expect(ed.slide.elements[0].text).toContain('text-shadow');
    ed.setTextEffect('none');
    expect(ed.slide.elements[0].text).toBe('<p>One <b>two</b></p><ul><li>three</li></ul>');
  });
});

describe('editor: sections', () => {
  // jsdom has <dialog> but not its modal methods.
  HTMLDialogElement.prototype.showModal ||= function showModal() { this.open = true; };
  HTMLDialogElement.prototype.close ||= function close() {
    this.open = false;
    this.dispatchEvent(new Event('close'));
  };
  /** Answer the open prompt with `value`. */
  const answer = async (value) => {
    const dlg = document.querySelector('dialog');
    dlg.querySelector('input').value = value;
    dlg.querySelector('form').dispatchEvent(new Event('submit', { cancelable: true }));
    if (dlg.open) dlg.close();
    await new Promise((r) => setTimeout(r));
  };
  const headers = (ed) => [...ed.filmstrip.list.querySelectorAll('.film-section-name')].map((e) => e.textContent);

  it('adds, renames and removes sections as undoable steps, with a header for each', async () => {
    const d = deckWith();
    d.slides.push(newSlide('blank', d), newSlide('blank', d));
    const ed = openEditor(d);
    expect(headers(ed)).toEqual([]);
    ed.goToSlide(1);
    const adding = ed.addSection();
    await answer('  Middle ');
    await adding;
    expect(ed.deck.slides.map((s) => s.section)).toEqual([undefined, 'Middle', undefined, undefined]);
    expect(headers(ed)).toEqual(['Default Section', 'Middle']);
    // The header row comes right before the section's first slide.
    expect(ed.filmstrip.list.children[2].classList.contains('film-section')).toBe(true);
    expect(ed.filmstrip.list.children[3].dataset.id).toBe(ed.deck.slides[1].id);

    const renaming = ed.renameSection(3);
    await answer('Second half');
    await renaming;
    expect(headers(ed)).toEqual(['Default Section', 'Second half']);
    ed.undo();
    expect(headers(ed)).toEqual(['Default Section', 'Middle']);
    ed.removeSection(2);
    expect(ed.deck.slides.some((s) => s.section)).toBe(false);
    expect(headers(ed)).toEqual([]);
    ed.undo();
    expect(headers(ed)).toEqual(['Default Section', 'Middle']);
  });

  it('keeps one section name when slides are duplicated, and passes it on when the first slide is deleted', () => {
    const d = deckWith();
    d.slides[1].section = 'Two';
    d.slides.push(newSlide('blank', d));
    const ed = openEditor(d);
    ed.goToSlide(1);
    ed.duplicateSlides();
    expect(ed.deck.slides.map((s) => s.section)).toEqual([undefined, 'Two', undefined, undefined]);
    ed.goToSlide(1);
    ed.deleteSlides();
    expect(ed.deck.slides.map((s) => s.section)).toEqual([undefined, 'Two', undefined]);
    expect(headers(ed)).toEqual(['Default Section', 'Two']);
  });

  it('folds a section away from its header, and unfolds it to show a slide in it', () => {
    const d = deckWith();
    d.slides[1].section = 'Two';
    const ed = openEditor(d);
    ed.filmstrip.list.querySelectorAll('.film-section-toggle')[1].click();
    const item = () => ed.filmstrip.items.get(d.slides[1].id);
    expect(item().classList.contains('in-collapsed')).toBe(true);
    expect(ed.filmstrip.list.querySelectorAll('.film-section-toggle')[1].getAttribute('aria-expanded')).toBe('false');
    ed.goToSlide(1);
    expect(item().classList.contains('in-collapsed')).toBe(false);
  });
});

describe('editor: stacking order', () => {
  const ids = (list) => list.map((e) => e.id).join(' ');
  const els = (spec) => spec.split(' ').map((s) => ({ id: s, ...(s.length > 1 && { groupId: s[0] }) }));

  it('steps a group over a whole group, never splitting either', () => {
    const list = els('A1 A2 B1 B2 c');
    expect(ids(stepOrder(list, new Set(['A1', 'A2']), 1))).toBe('B1 B2 A1 A2 c');
    expect(ids(stepOrder(list, new Set(['c']), -1))).toBe('A1 A2 c B1 B2');
    expect(ids(stepOrder(list, new Set(['c']), 1))).toBe('A1 A2 B1 B2 c');
  });

  it('moves an object chosen inside a group within it', () => {
    expect(ids(stepOrder(els('A1 A2 A3 b'), new Set(['A1']), 1))).toBe('A2 A1 A3 b');
    expect(ids(stepOrder(els('A1 A2 A3 b'), new Set(['A3']), 1))).toBe('A1 A2 A3 b');
  });
});
