import { h } from '../ui/dom.js';
import { icon } from '../ui/icons.js';
import { openDialog } from '../ui/dialog.js';
import { DIAGRAM_LAYOUTS, DIAGRAM_COLORS, DIAGRAM_STYLES, newDiagram } from '../model/deck.js';
import { themeById, backgroundCss } from '../model/themes.js';
import { outlineText, parseOutline, lineAt, addNode, removeNode, promote, demote, moveUp, moveDown } from '../model/diagram-outline.js';
import { renderElement } from '../render/render.js';
import { exportsAsSmartArt } from '../io/pptx-diagram.js';

/**
 * Diagrams in the editor: the layout picker (Insert › Diagram…) and the
 * Diagram section of the format options, whose text pane edits the outline
 * a diagram is drawn from (see model/diagram-outline.js).
 */

/** Insert › Diagram…: a card for each layout, drawn in the deck's theme. Resolves with the layout picked, or null. */
export async function pickDiagramLayout(deck) {
  const theme = themeById(deck.themeId);
  let chosen = null;
  const cards = DIAGRAM_LAYOUTS.map((kind) => {
    const sample = renderElement(newDiagram(kind.id, 0, 0, 240, 150), { theme, deck });
    sample.style.transform = 'scale(0.5)';
    sample.style.transformOrigin = '0 0';
    const preview = h('div', { class: 'diagram-card-preview', 'aria-hidden': 'true', style: { background: backgroundCss(null, theme) } }, sample);
    const card = h('button', { type: 'button', class: 'diagram-card' }, preview, h('span', {}, kind.name));
    card.addEventListener('click', () => {
      chosen = kind.id;
      card.closest('dialog')?.close();
    });
    return card;
  });
  await openDialog({
    title: 'Insert diagram', className: 'wide', initialFocus: '.diagram-card',
    body: h('div', { class: 'diagram-picker' }, ...cards),
    buttons: [{ label: 'Cancel', value: null }],
  });
  return chosen;
}

const select = (label, options, value, onValue) => {
  const s = h('select', {}, ...options.map((o) => h('option', { value: o.id, selected: o.id === value || null }, o.name)));
  s.addEventListener('change', () => onValue(s.value));
  return h('label', { class: 'pf' }, h('span', { class: 'pf-label' }, label), s);
};

/** The caret at the end of line `line` of a text area. */
function caretToLine(area, line) {
  const lines = area.value.split('\n');
  const at = lines.slice(0, Math.max(0, line) + 1).join('\n').length;
  area.setSelectionRange(at, at);
}

/**
 * The Diagram section of the format options for diagram `el`: its layout,
 * colours and style, its text as an outline (a line per shape, indented with
 * Tab), buttons to add, remove and move its shapes, and "Convert to shapes".
 * Each change is one undo step; typing coalesces.
 */
export function diagramSection(panel, el) {
  const app = panel.app;
  const d = el.diagram;
  const current = () => app.selectedElements().find((e) => e.id === el.id)?.diagram;

  const outline = h('textarea', {
    class: 'diagram-outline', rows: Math.min(12, Math.max(5, d.nodes.length + 1)), spellcheck: 'true',
    placeholder: 'A line for each shape', 'aria-label': 'Diagram text: a line for each shape. Tab moves a line in, Shift+Tab out.',
  });
  outline.value = outlineText(d.nodes);
  outline.addEventListener('input', () => {
    const now = current();
    if (!now) return;
    const nodes = parseOutline(outline.value, now.nodes);
    app.editDiagram((g) => { g.nodes = nodes; }, { coalesce: `diagram-text:${el.id}` });
  });

  /** An outline operation on the line with the caret, keeping the caret on the node it moved. */
  const run = (fn, { needsNode = true } = {}) => {
    const now = current();
    if (!now) return;
    const nodes = parseOutline(outline.value, now.nodes);
    const i = Math.min(lineAt(outline.value, outline.selectionStart), nodes.length - 1);
    if (needsNode && i < 0) return;
    const res = fn(nodes, i);
    if (res.nodes === nodes) return;
    panel.diagramCaret = { id: el.id, line: res.index };
    app.editDiagram((g) => { g.nodes = res.nodes; });
    // Typing here, the pane isn't rebuilt: show the change in place.
    if (document.activeElement === outline) {
      outline.value = outlineText(res.nodes);
      caretToLine(outline, res.index);
      panel.diagramCaret = null;
    }
  };
  outline.addEventListener('keydown', (e) => {
    if (e.key !== 'Tab' || e.ctrlKey || e.metaKey || e.altKey) return;
    e.preventDefault();
    run(e.shiftKey ? promote : demote);
  });
  // After a button's change rebuilt the pane, the caret is where the node went.
  if (panel.diagramCaret?.id === el.id) {
    caretToLine(outline, panel.diagramCaret.line);
    panel.diagramCaret = null;
  }

  const tool = (name, label, fn, opts) => h('button', { type: 'button', class: 'icon-btn', title: label, 'aria-label': label, html: icon(name), onclick: () => run(fn, opts) });
  const tools = h('div', { class: 'pf-row diagram-tools', role: 'group', 'aria-label': 'Shapes' },
    tool('plus', 'Add a shape after this one', (nodes, i) => addNode(nodes, i), { needsNode: false }),
    tool('trash', 'Remove this shape', removeNode),
    tool('outdent', 'Promote (Shift+Tab)', promote),
    tool('indent', 'Demote (Tab)', demote),
    tool('up', 'Move up', moveUp),
    tool('down', 'Move down', moveDown));

  const asSmartArt = exportsAsSmartArt(el);
  const notes = [
    d.drawn ? 'Shown as PowerPoint drew it until you change it.' : null,
    asSmartArt ? 'Saved to PowerPoint as SmartArt.' : 'Saved to PowerPoint as grouped shapes.',
    d.source && !asSmartArt ? 'It goes back as SmartArt only with the layout, colours and style it came with.' : null,
  ].filter(Boolean).join(' ');
  return h('section', { class: 'panel-section diagram-section' }, h('h3', {}, 'Diagram'),
    h('div', { class: 'pf-stack' },
      select('Layout', DIAGRAM_LAYOUTS, d.layout, (v) => app.editDiagram((g) => { g.layout = v; })),
      select('Colours', DIAGRAM_COLORS, d.colors, (v) => app.editDiagram((g) => { g.colors = v; })),
      select('Style', DIAGRAM_STYLES, d.style, (v) => app.editDiagram((g) => { g.style = v; })),
      h('label', { class: 'pf' }, h('span', { class: 'pf-label' }, 'Text'), outline),
      tools,
      h('small', { class: 'pf-hint' }, 'A line for each shape. Tab and Shift+Tab move a line in and out, under the shape before it.'),
      h('button', { type: 'button', class: 'btn', onclick: () => app.convertDiagramToShapes() }, 'Convert to shapes'),
      h('small', { class: 'pf-hint' }, notes)));
}

/** Put the keyboard in a diagram's text pane (after opening the format options): at the end, or of its first line. */
export function focusDiagramText(panel, { start = false } = {}) {
  const area = panel.body.querySelector('.diagram-outline');
  if (!area) return;
  area.focus({ preventScroll: false });
  if (start) caretToLine(area, 0);
  else area.setSelectionRange(area.value.length, area.value.length);
}
