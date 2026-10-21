import { describe, it, expect, vi, afterEach, beforeAll } from 'vitest';

vi.mock('../../src/storage/db.js', () => ({
  getDeck: async () => null,
  saveDeck: async (deck) => ({ rev: (deck.rev || 0) + 1 }),
  onStorageChange: () => () => {},
}));

import { EditorScreen } from '../../src/editor/editor-screen.js';
import { newDeck, newSlide, newDiagram } from '../../src/model/deck.js';

// jsdom lacks these.
globalThis.ResizeObserver ||= class { observe() {} disconnect() {} unobserve() {} };
globalThis.matchMedia ||= () => ({ matches: false, addEventListener() {}, removeEventListener() {} });
globalThis.CSS ||= { escape: (s) => String(s).replace(/["\\]/g, '\\$&') };
Element.prototype.scrollIntoView ||= () => {};
document.queryCommandState ||= () => false;
document.execCommand ||= () => false;

beforeAll(() => {
  HTMLDialogElement.prototype.showModal ||= function showModal() { this.open = true; };
  HTMLDialogElement.prototype.close ||= function close() {
    this.open = false;
    this.dispatchEvent(new Event('close'));
  };
});

let editor;
function openEditor(...elements) {
  const deck = newDeck();
  deck.slides[0] = newSlide('blank', deck);
  deck.slides[0].elements.push(...elements);
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

afterEach(async () => {
  await editor?.destroy({ save: false });
  editor = null;
  document.body.replaceChildren();
});

const pane = () => editor.panel.body.querySelector('.diagram-outline');
const texts = (el) => el.diagram.nodes.map((n) => '\t'.repeat(n.level) + n.text);
const type = (area, value) => {
  area.value = value;
  area.dispatchEvent(new Event('input', { bubbles: true }));
};

describe('editor: diagrams', () => {
  it('inserts a diagram of the layout picked, its text ready to type in', async () => {
    openEditor();
    const done = editor.insertDiagram();
    const cards = [...document.querySelectorAll('dialog .diagram-card')];
    expect(cards.map((c) => c.lastElementChild.textContent)).toContain('Chevron process');
    cards.find((c) => c.lastElementChild.textContent === 'Chevron process').click();
    await done;
    const [el] = editor.slide.elements;
    expect(el).toMatchObject({ type: 'diagram', diagram: { layout: 'chevron' } });
    expect(editor.selection).toEqual([el.id]);
    expect(editor.panel.current).toBe('format');
    expect(document.activeElement).toBe(pane());
    expect(editor.canvas.nodeOf(el.id).querySelectorAll('.el-diagram-shapes > .el').length).toBe(el.diagram.nodes.length);
  });

  it('edits the outline as text, re-laying the diagram out, and undoes it in one step', () => {
    const el = newDiagram('list', 100, 100, 500, 300);
    openEditor(el);
    editor.setSelection([el.id]);
    editor.editDiagramText();
    const before = texts(editor.slide.elements[0]);
    type(pane(), 'Alpha\n\tAlpha one\nBeta');
    type(pane(), 'Alpha\n\tAlpha one\nBeta\nGamma');
    const now = editor.slide.elements[0];
    expect(texts(now)).toEqual(['Alpha', '\tAlpha one', 'Beta', 'Gamma']);
    expect(now.diagram.edited).toBe(true);
    expect(editor.canvas.nodeOf(el.id).querySelectorAll('.el-diagram-shapes > .el')).toHaveLength(3);
    editor.undo();
    expect(texts(editor.slide.elements[0])).toEqual(before);
  });

  it('indents with Tab and Shift+Tab, and moves and adds shapes with the buttons', () => {
    const el = newDiagram('list', 100, 100, 500, 300);
    el.diagram.nodes = [{ id: 'a', text: 'A', level: 0 }, { id: 'b', text: 'B', level: 0 }, { id: 'c', text: 'C', level: 0 }];
    openEditor(el);
    editor.setSelection([el.id]);
    editor.editDiagramText();
    const area = pane();
    area.setSelectionRange(3, 3); // on "B"
    area.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true }));
    expect(texts(editor.slide.elements[0])).toEqual(['A', '\tB', 'C']);
    expect(area.value).toBe('A\n\tB\nC');
    area.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', shiftKey: true, bubbles: true, cancelable: true }));
    expect(texts(editor.slide.elements[0])).toEqual(['A', 'B', 'C']);
    // Move "B" up, with the button, then add a shape after it.
    area.setSelectionRange(3, 3);
    editor.panel.body.querySelector('[aria-label="Move up"]').click();
    expect(texts(editor.slide.elements[0])).toEqual(['B', 'A', 'C']);
    pane().setSelectionRange(0, 0);
    editor.panel.body.querySelector('[aria-label="Add a shape after this one"]').click();
    expect(texts(editor.slide.elements[0])).toEqual(['B', '', 'A', 'C']);
    editor.undo();
    expect(texts(editor.slide.elements[0])).toEqual(['B', 'A', 'C']);
  });

  it('changes layout and colours from the format options', () => {
    const el = newDiagram('list', 100, 100, 500, 300);
    openEditor(el);
    editor.setSelection([el.id]);
    editor.openPanel('format');
    const [layout, colors] = editor.panel.body.querySelectorAll('.diagram-section select');
    layout.value = 'cycle';
    layout.dispatchEvent(new Event('change'));
    expect(editor.slide.elements[0].diagram.layout).toBe('cycle');
    expect(editor.canvas.nodeOf(el.id).querySelector('.el-shape svg path')).toBeTruthy();
    colors.value = 'colorful';
    colors.dispatchEvent(new Event('change'));
    expect(editor.slide.elements[0].diagram.colors).toBe('colorful');
    editor.undo();
    editor.undo();
    expect(editor.slide.elements[0].diagram).toMatchObject({ layout: 'list', colors: 'accent1' });
  });

  it('puts the keyboard in the text pane on a double-click', () => {
    const el = newDiagram('process', 100, 100, 500, 300);
    openEditor(el);
    editor.setSelection([el.id]);
    editor.canvas.nodeOf(el.id).querySelector('.el-shape').dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    expect(editor.panel.current).toBe('format');
    expect(document.activeElement).toBe(pane());
  });

  it('converts a diagram to grouped shapes, keeping PowerPoint’s drawing as it was', () => {
    const el = newDiagram('process', 100, 100, 500, 300);
    el.diagram.drawn = { w: 250, h: 150, elements: [{ id: 'x', type: 'shape', shape: 'ellipse', x: 10, y: 10, w: 50, h: 50, text: '', fill: '#ff0000', style: {} }] };
    openEditor(el);
    editor.setSelection([el.id]);
    editor.convertDiagramToShapes();
    const els = editor.slide.elements;
    expect(els).toHaveLength(1);
    expect(els[0]).toMatchObject({ type: 'shape', shape: 'ellipse', x: 120, y: 120, w: 100, h: 100 });
    expect(editor.selection).toEqual([els[0].id]);
    editor.undo();
    expect(editor.slide.elements[0].type).toBe('diagram');
  });
});
