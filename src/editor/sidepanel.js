import { h } from '../ui/dom.js';
import { icon } from '../ui/icons.js';
import { openColorPicker } from '../ui/color-picker.js';
import { THEMES, resolveColor, backgroundBaseColor } from '../model/themes.js';
import { SIZES, LAYOUTS, TRANSITIONS, ANIMATIONS, CHART_KINDS, newSlide, fillPlaceholders, canHoldText, slideText } from '../model/deck.js';
import { renderThumb, textStyle } from '../render/render.js';
import { seriesColor, pointColor } from '../render/chart.js';
import { commentsPanel } from './comments.js';

const TITLES = { design: 'Design', transition: 'Transition', animate: 'Animate', format: 'Format options', comments: 'Comments' };
const TAB_LABELS = { design: 'Design', transition: 'Transition', animate: 'Animate', format: 'Format', comments: 'Comments' };

const field = (label, control, hint = null) => h('label', { class: 'pf' }, h('span', { class: 'pf-label' }, label), control, hint ? h('small', { class: 'pf-hint' }, hint) : null);
const section = (title, ...children) => h('section', { class: 'panel-section' }, h('h3', {}, title), ...children);

function numberInput(value, onValue, { step = 1, min = null, max = null, suffix = '' } = {}) {
  const input = h('input', { type: 'number', step, min, max, value: Number.isFinite(value) ? String(Math.round(value * 10) / 10) : '' });
  let last = input.value;
  const fire = () => {
    // Enter fires this, then the change event does too: one undo step.
    if (input.value === last) return;
    last = input.value;
    const n = parseFloat(input.value);
    if (Number.isFinite(n)) onValue(min != null ? Math.max(min, max != null ? Math.min(max, n) : n) : n);
  };
  input.addEventListener('change', fire);
  input.addEventListener('keydown', (e) => { if (e.key === 'Enter') fire(); });
  return suffix ? h('span', { class: 'pf-num' }, input, h('span', { class: 'pf-suffix' }, suffix)) : input;
}

function select(options, value, onValue) {
  const s = h('select', {}, ...options.map(([v, label]) => h('option', { value: v, selected: String(v) === String(value) || null }, label)));
  s.addEventListener('change', () => onValue(s.value));
  return s;
}

function checkbox(label, checked, onValue) {
  const input = h('input', { type: 'checkbox', checked: checked || null });
  input.addEventListener('change', () => onValue(input.checked));
  return h('label', { class: 'pf-check' }, input, label);
}

function colorButton(label, value, theme, onPick, opts = {}) {
  const sw = h('span', { class: 'pf-swatch', style: { background: resolveColor(value, theme, 'transparent') } });
  const b = h('button', { type: 'button', class: 'btn pf-color', 'aria-label': label, 'aria-haspopup': 'true' }, sw, h('span', {}, label));
  b.addEventListener('click', () => openColorPicker(b, { theme, current: value, onPick, ...opts }));
  return b;
}

/**
 * The right-hand panel: themes and backgrounds, transitions, animations,
 * the detailed format options for the selection and the slide's comments.
 */
/** Gradient fill controls for a shape: on/off, its end colours and its direction. */
function gradientControls(app, el, theme) {
  const g = el.gradient;
  const toggle = checkbox('Gradient', !!g, (v) => app.setGradient(v
    ? { type: 'linear', angle: 180, stops: [{ pos: 0, color: el.fill || 'theme:accent1' }, { pos: 1, color: 'theme:accent2' }] }
    : null));
  if (!g) return [toggle];
  const stops = g.stops;
  const withStop = (i, color) => ({ ...g, stops: stops.map((s, k) => (k === i ? { ...s, color } : s)) });
  return [
    toggle,
    h('div', { class: 'pf-row' },
      colorButton('From', stops[0].color, theme, (v) => app.setGradient(withStop(0, v))),
      colorButton('To', stops[stops.length - 1].color, theme, (v) => app.setGradient(withStop(stops.length - 1, v)))),
    field('Direction', select([
      ['linear:180', 'Top to bottom'], ['linear:0', 'Bottom to top'], ['linear:90', 'Left to right'], ['linear:270', 'Right to left'],
      ['linear:135', 'Diagonal'], ['linear:225', 'Diagonal (reverse)'], ['radial', 'From the centre'],
    ], g.type === 'radial' ? 'radial' : `linear:${g.angle ?? 180}`, (v) => {
      const [type, angle] = v.split(':');
      app.setGradient({ ...g, type, angle: angle == null ? g.angle : +angle });
    })),
  ];
}

export class SidePanel {
  constructor(app) {
    this.app = app;
    this.current = null;
    this.key = '';
    this.activeComment = null; // the comment thread highlighted here and on the slide
    this.title = h('h2', {});
    this.body = h('div', { class: 'panel-body' });
    this.tabs = h('div', { class: 'panel-tabs', role: 'tablist' }, ...Object.entries(TAB_LABELS).map(([id, name]) => {
      const t = h('button', { type: 'button', role: 'tab', class: 'panel-tab', dataset: { tab: id }, onclick: () => this.open(id) }, name);
      return t;
    }));
    this.el = h('aside', { class: 'sidepanel', hidden: true, 'aria-label': 'Side panel' },
      h('header', { class: 'panel-header' }, this.title, h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Close panel', html: icon('close'), onclick: () => this.close() })),
      this.tabs, this.body);
  }

  open(name) {
    this.current = name;
    this.el.hidden = false;
    this.key = '';
    this.update(true);
    this.app.canvas.layout();
  }

  close() {
    this.current = null;
    this.el.hidden = true;
    this.app.canvas.layout();
    this.app.focusCanvas();
  }

  update(force = false) {
    if (!this.current) return;
    const app = this.app;
    const key = `${this.current}|${app.slide.id}|${app.selection.join(',')}`;
    // Don't rebuild under the user's fingers while they type or drag a slider here. Selects,
    // checkboxes and number fields only change things once a choice is made, so they rebuild
    // (showing the controls that go with it, or the other side of a locked aspect ratio).
    const typing = this.el.contains(document.activeElement) && document.activeElement.matches('textarea, input:not([type=number], [type=checkbox])');
    if (!force && key === this.key && typing) return;
    this.key = key;
    this.title.textContent = TITLES[this.current];
    for (const t of this.tabs.children) {
      t.setAttribute('aria-selected', String(t.dataset.tab === this.current));
      t.classList.toggle('is-active', t.dataset.tab === this.current);
    }
    // The tabs scroll sideways when the panel is narrow: keep the open one in view.
    const tab = this.tabs.querySelector('.is-active')?.getBoundingClientRect();
    const strip = this.tabs.getBoundingClientRect();
    if (tab && tab.right > strip.right) this.tabs.scrollLeft += tab.right - strip.right;
    else if (tab && tab.left < strip.left) this.tabs.scrollLeft -= strip.left - tab.left;
    // Keyboard focus stays on the same control across the rebuild.
    const controls = () => [...this.body.querySelectorAll('button, input, select, textarea')];
    const focusAt = controls().indexOf(document.activeElement);
    const scroll = this.body.scrollTop;
    this.body.replaceChildren(...this[this.current]());
    this.body.scrollTop = scroll;
    if (focusAt >= 0) controls()[focusAt]?.focus({ preventScroll: true });
  }

  // ---------------------------------------------------------------- design

  design() {
    const app = this.app;
    const { deck, slide, theme } = app;
    const themes = h('div', { class: 'theme-grid' });
    for (const t of THEMES) {
      const sample = { themeId: t.id, size: deck.size, media: {} };
      const s = fillPlaceholders(newSlide('title', sample), { title: t.name, subtitle: 'Aa' });
      const b = h('button', { type: 'button', class: `theme-card${t.id === deck.themeId ? ' is-active' : ''}`, 'aria-label': `Theme: ${t.name}`, 'aria-pressed': String(t.id === deck.themeId), onclick: () => app.setTheme(t.id) },
        renderThumb(s, sample, 112));
      themes.append(b);
    }

    const bg = slide.background;
    const bgType = bg?.type || 'theme';
    const bgControls = h('div', { class: 'pf-stack' });
    bgControls.append(field('Fill', select([['theme', 'Theme default'], ['solid', 'Solid colour'], ['gradient', 'Gradient'], ['image', 'Picture']], bgType, (v) => {
      if (v === 'theme') app.setBackground(null);
      else if (v === 'solid') app.setBackground({ type: 'solid', color: backgroundBaseColor(theme.background) });
      else if (v === 'gradient') app.setBackground({ type: 'gradient', from: 'theme:accent1', to: 'theme:accent2', angle: 135 });
      else app.pickBackgroundImage();
    })));
    if (bgType === 'solid') bgControls.append(colorButton('Colour', bg.color, theme, (v) => app.setBackground({ type: 'solid', color: v })));
    if (bgType === 'gradient') {
      bgControls.append(
        h('div', { class: 'pf-row' },
          colorButton('From', bg.from, theme, (v) => app.setBackground({ ...bg, from: v })),
          colorButton('To', bg.to, theme, (v) => app.setBackground({ ...bg, to: v }))),
        field('Angle', select([[0, 'Bottom to top'], [90, 'Left to right'], [135, 'Diagonal'], [180, 'Top to bottom'], [225, 'Diagonal (reverse)']], bg.angle ?? 180, (v) => app.setBackground({ ...bg, angle: +v }))),
      );
    }
    if (bgType === 'image') bgControls.append(h('button', { type: 'button', class: 'btn', onclick: () => app.pickBackgroundImage() }, 'Choose another picture…'));
    bgControls.append(
      checkbox('Show theme graphics', !slide.hideDecor, (v) => app.changeSlide((s) => { s.hideDecor = !v; })),
      h('div', { class: 'pf-row' },
        h('button', { type: 'button', class: 'btn', onclick: () => app.setBackground(slide.background, { all: true }) }, 'Apply to all slides'),
        h('button', { type: 'button', class: 'btn btn-ghost', disabled: !slide.background || null, onclick: () => app.setBackground(null) }, 'Reset')),
    );

    const layouts = h('div', { class: 'layout-grid' });
    for (const l of LAYOUTS) {
      const sample = { themeId: deck.themeId, size: deck.size, media: {} };
      const s = newSlide(l.id, sample);
      for (const el of s.elements) el.text = el.ph === 'title' ? '<p>Title</p>' : el.bullets ? '<ul><li>Text</li><li>Text</li></ul>' : '<p>Text</p>';
      layouts.append(h('button', { type: 'button', class: `layout-card${slide.layout === l.id ? ' is-active' : ''}`, title: l.name, 'aria-label': `Layout: ${l.name}`, onclick: () => app.applyLayout(l.id) },
        renderThumb(s, sample, 72), h('span', {}, l.name)));
    }

    return [
      section('Theme', themes),
      section('Background (this slide)', bgControls),
      section('Layout (this slide)', layouts),
      section('Slide size', select(Object.entries(SIZES).map(([id, s]) => [id, s.label]), deck.size, (v) => app.setSize(v)),
        h('p', { class: 'pf-hint' }, 'Changing the size scales everything on every slide.')),
    ];
  }

  // ---------------------------------------------------------------- transitions

  transition() {
    const app = this.app;
    const t = app.slide.transition || { type: 'none', duration: 0.6 };
    const list = h('div', { class: 'choice-list', role: 'radiogroup', 'aria-label': 'Transition' });
    for (const tr of TRANSITIONS) {
      list.append(h('button', { type: 'button', role: 'radio', 'aria-checked': String(t.type === tr.id), class: `choice${t.type === tr.id ? ' is-active' : ''}`, onclick: () => app.setTransition({ ...t, type: tr.id }) }, tr.name));
    }
    const dur = h('input', { type: 'range', min: '0.2', max: '2', step: '0.1', value: String(t.duration ?? 0.6) });
    const out = h('output', {}, `${t.duration ?? 0.6} s`);
    dur.addEventListener('input', () => {
      out.textContent = `${dur.value} s`;
      app.setTransition({ ...app.slide.transition, duration: +dur.value }, { coalesce: 'transition-duration' });
    });
    return [
      section(`Slide ${app.slideIndex + 1}`, list, field('Duration', h('div', { class: 'pf-row' }, dur, out))),
      h('div', { class: 'pf-row' },
        h('button', { type: 'button', class: 'btn', onclick: () => app.setTransition(app.slide.transition, { all: true }) }, 'Apply to all slides'),
        h('button', { type: 'button', class: 'btn btn-primary', disabled: t.type === 'none' || null, html: `${icon('play')}<span>Preview</span>`, onclick: () => app.previewTransition() })),
    ];
  }

  // ---------------------------------------------------------------- animations

  animate() {
    const app = this.app;
    const els = app.selectedElements();
    const out = [];
    if (els.length) {
      const a = els[0].anim || { effect: 'none', duration: 0.5 };
      out.push(section(els.length > 1 ? `${els.length} objects` : 'Selected object',
        field('Effect', select(ANIMATIONS.map((x) => [x.id, x.name]), a.effect, (v) => app.setAnimation({ effect: v }))),
        field('Duration', select([[0.3, 'Fast (0.3 s)'], [0.5, 'Medium (0.5 s)'], [0.8, 'Slow (0.8 s)'], [1.2, 'Very slow (1.2 s)']], a.duration ?? 0.5, (v) => app.setAnimation({ duration: +v }))),
        field('Start', select([['click', 'On click'], ['with', 'With previous'], ['after', 'After previous']], a.start || 'click', (v) => app.setAnimation({ start: v })))));
    } else {
      out.push(h('p', { class: 'pf-hint' }, 'Select an object on the slide to animate it.'));
    }
    const animated = app.slide.elements.filter((e) => e.anim && e.anim.effect !== 'none').sort((x, y) => (x.anim.order || 0) - (y.anim.order || 0));
    const list = h('ol', { class: 'anim-list' });
    animated.forEach((el, i) => {
      const name = (canHoldText(el) && slideText({ elements: [el] })) || ({ image: 'Picture', line: 'Line', table: 'Table', shape: 'Shape', text: 'Text box', chart: 'Chart', media: el.kind === 'audio' ? 'Sound' : 'Video' }[el.type]) || 'Object';
      const effect = ANIMATIONS.find((x) => x.id === el.anim.effect)?.name;
      list.append(h('li', { class: `anim-item${app.selection.includes(el.id) ? ' is-selected' : ''}` },
        h('button', { type: 'button', class: 'anim-name', onclick: () => app.setSelection([el.id]) }, h('b', {}, name.slice(0, 40)), h('small', {}, `${effect}${el.anim.start && el.anim.start !== 'click' ? ` · ${el.anim.start} previous` : ''}`)),
        h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Move earlier', disabled: i === 0 || null, html: icon('up'), onclick: () => app.reorderAnimation(el.id, -1) }),
        h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Move later', disabled: i === animated.length - 1 || null, html: icon('down'), onclick: () => app.reorderAnimation(el.id, 1) }),
        h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Remove animation', html: icon('close'), onclick: () => app.setAnimation({ effect: 'none' }, [el.id]) })));
    });
    out.push(section('Order on this slide', animated.length ? list : h('p', { class: 'pf-hint' }, 'Nothing on this slide is animated yet.')));
    if (animated.length) out.push(h('button', { type: 'button', class: 'btn btn-primary', html: `${icon('play')}<span>Preview</span>`, onclick: () => app.previewAnimations() }));
    return out;
  }

  // ---------------------------------------------------------------- comments

  comments() { return commentsPanel(this.app); }

  // ---------------------------------------------------------------- format options

  format() {
    const app = this.app;
    const els = app.selectedElements();
    if (!els.length) return [h('p', { class: 'pf-hint' }, 'Select something on the slide to see its format options.')];
    const el = els[0];
    const theme = app.theme;
    const single = els.length === 1;
    const out = [];
    const upd = (fn, coalesce) => app.updateSelected(fn, { coalesce });

    if (single && el.type === 'line') {
      out.push(section('Position',
        h('div', { class: 'pf-grid' },
          field('Start X', numberInput(el.x1, (v) => upd((e) => { e.x1 = v; }))), field('Start Y', numberInput(el.y1, (v) => upd((e) => { e.y1 = v; }))),
          field('End X', numberInput(el.x2, (v) => upd((e) => { e.x2 = v; }))), field('End Y', numberInput(el.y2, (v) => upd((e) => { e.y2 = v; }))))));
      out.push(section('Line',
        colorButton('Colour', el.stroke?.color, theme, (v) => app.setStroke({ color: v })),
        field('Weight', numberInput(el.stroke?.width || 2, (v) => app.setStroke({ width: v }), { min: 0.5, max: 50, step: 0.5, suffix: 'pt' })),
        checkbox('Arrow at start', el.arrowStart, (v) => upd((e) => { e.arrowStart = v; })),
        checkbox('Arrow at end', el.arrowEnd, (v) => upd((e) => { e.arrowEnd = v; }))));
    } else if (single) {
      const ratio = el.w / el.h;
      out.push(section('Size & rotation',
        h('div', { class: 'pf-grid' },
          field('Width', numberInput(el.w, (v) => upd((e) => { if (e.lockRatio) e.h = v / ratio; e.w = v; }), { min: 4, suffix: 'pt' })),
          field('Height', numberInput(el.h, (v) => upd((e) => { if (e.lockRatio) e.w = v * ratio; e.h = v; }), { min: 4, suffix: 'pt' })),
          field('Rotation', numberInput(el.rot || 0, (v) => upd((e) => { e.rot = ((v % 360) + 360) % 360; }), { suffix: '°' }))),
        checkbox('Lock aspect ratio', el.lockRatio, (v) => upd((e) => { e.lockRatio = v; }))));
      out.push(section('Position',
        h('div', { class: 'pf-grid' },
          field('X', numberInput(el.x, (v) => upd((e) => { e.x = v; }), { suffix: 'pt' })),
          field('Y', numberInput(el.y, (v) => upd((e) => { e.y = v; }), { suffix: 'pt' }))),
        h('div', { class: 'pf-row' },
          h('button', { type: 'button', class: 'btn', onclick: () => app.align('center', { toSlide: true }) }, 'Center horizontally'),
          h('button', { type: 'button', class: 'btn', onclick: () => app.align('middle', { toSlide: true }) }, 'Center vertically'))));
    }

    if (els.every((e) => e.type !== 'line')) {
      const fillable = els.some((e) => e.type === 'shape' || e.type === 'text' || e.type === 'table');
      const look = h('div', { class: 'pf-stack' });
      if (fillable) look.append(colorButton(el.type === 'table' ? 'Header colour' : 'Fill', el.fill, theme, (v) => app.setElementProp('fill', v), { allowNone: true }));
      if (fillable && el.type !== 'table' && el.fill !== 'none' && !el.fillImage) look.append(...gradientControls(app, el, theme));
      if (fillable && el.type !== 'table') {
        look.append(h('div', { class: 'pf-row' },
          h('button', { type: 'button', class: 'btn', onclick: () => app.pickFillImage() }, el.fillImage ? 'Change fill picture…' : 'Fill with a picture…'),
          el.fillImage ? h('button', { type: 'button', class: 'btn btn-ghost', onclick: () => app.removeFillImage() }, 'Remove picture') : null));
      }
      look.append(
        colorButton('Border', el.stroke?.width ? el.stroke.color : 'none', theme, (v) => app.setStroke({ color: v, width: Math.max(1, el.stroke?.width || 2) }), { allowNone: true, noneLabel: 'No border' }),
        field('Border weight', numberInput(el.stroke?.width || 0, (v) => app.setStroke({ width: v }), { min: 0, max: 50, step: 0.5, suffix: 'pt' })),
      );
      const op = h('input', { type: 'range', min: '0', max: '100', value: String(Math.round((el.opacity ?? 1) * 100)) });
      const opOut = h('output', {}, `${op.value}%`);
      op.addEventListener('input', () => {
        opOut.textContent = `${op.value}%`;
        upd((e) => { e.opacity = +op.value / 100; }, 'opacity');
      });
      look.append(field('Opacity', h('div', { class: 'pf-row' }, op, opOut)), checkbox('Drop shadow', el.shadow, (v) => upd((e) => { e.shadow = v; })));
      out.push(section('Appearance', look));
    }

    if (els.some((e) => canHoldText(e) || e.type === 'table')) {
      const t = textStyle(el, theme);
      out.push(section('Text',
        field('Vertical alignment', select([['top', 'Top'], ['middle', 'Middle'], ['bottom', 'Bottom']], t.valign, (v) => app.setTextProp('valign', v))),
        field('Line spacing', select([[1, 'Single'], [1.15, '1.15'], [1.25, '1.25'], [1.5, '1.5'], [2, 'Double']], t.lineHeight, (v) => app.setTextProp('lineHeight', +v))),
        field('Space between paragraphs', select([[0, 'None'], [0.35, 'Small'], [0.8, 'Medium'], [1.4, 'Large']], t.spacing, (v) => app.setTextProp('spacing', +v)))));
    }

    if (single && el.type === 'image') {
      const alt = h('textarea', { rows: 2, placeholder: 'Describe the picture for people who can’t see it' });
      alt.value = el.alt || '';
      alt.addEventListener('change', () => upd((e) => { e.alt = alt.value; }));
      out.push(section('Picture',
        field('Fit', select([['fill', 'Stretch to box'], ['cover', 'Fill box (crop)'], ['contain', 'Fit inside box']], el.fit || 'fill', (v) => upd((e) => { e.fit = v; }))),
        field('Corner radius', numberInput(el.radius || 0, (v) => upd((e) => { e.radius = v; }), { min: 0, suffix: 'pt' })),
        h('div', { class: 'pf-row' },
          h('button', { type: 'button', class: 'btn', onclick: () => app.flip('h') }, 'Flip horizontally'),
          h('button', { type: 'button', class: 'btn', onclick: () => app.flip('v') }, 'Flip vertically')),
        h('button', { type: 'button', class: 'btn', onclick: () => app.resetImageSize() }, 'Reset size'),
        field('Alt text', alt)));
    }

    if (single && el.type === 'media') {
      out.push(section(el.kind === 'audio' ? 'Sound' : 'Video',
        checkbox('Play automatically in the slideshow', el.autoplay, (v) => upd((e) => { e.autoplay = v; })),
        checkbox('Loop until the slide changes', el.loop, (v) => upd((e) => { e.loop = v; })),
        h('p', { class: 'pf-hint' }, el.autoplay ? 'Starts when the slide (or its animation) appears.' : 'Click it in the slideshow to play.')));
    }

    if (single && el.type === 'table') {
      out.push(section('Table',
        checkbox('Header row', el.header, (v) => upd((e) => { e.header = v; })),
        checkbox('Banded rows', el.banded, (v) => upd((e) => { e.banded = v; })),
        h('p', { class: 'pf-hint' }, `${el.cells.length} rows × ${el.cells[0]?.length || 0} columns. Double-click the table to type in it; Tab moves between cells.`)));
    }

    if (single && el.type === 'chart' && el.chart) out.push(...chartSections(el, theme, upd));

    out.push(section('Arrange', checkbox('Lock position', els.every((e) => e.locked), (v) => upd((e) => { e.locked = v; }))));
    return out;
  }
}

// ------------------------------------------------------------------ charts

/** A number typed into the chart data: "1,234.5", "3,5" or blank (no point). */
function parseValue(text) {
  const t = String(text).trim().replace(/\s/g, '');
  if (!t) return null;
  const n = Number(/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(t) ? t.replace(/,/g, '') : t.replace(',', '.'));
  return Number.isFinite(n) ? n : null;
}

/** A small field of the data grid; it changes the chart when it's left (or Enter is pressed). */
function cellInput(value, label, onValue, numeric = false) {
  const input = h('input', { type: 'text', class: numeric ? 'is-num' : null, inputmode: numeric ? 'decimal' : null, 'aria-label': label });
  input.value = value ?? '';
  input.addEventListener('change', () => onValue(input.value));
  return input;
}

function swatch(label, value, theme, onPick) {
  const b = h('button', { type: 'button', class: 'chart-swatch', title: label, 'aria-label': label, 'aria-haspopup': 'true', style: { background: value } });
  b.addEventListener('click', () => openColorPicker(b, { theme, current: value, onPick }));
  return b;
}

const removeButton = (label, run) => h('button', { type: 'button', class: 'icon-btn chart-remove', title: label, 'aria-label': label, html: icon('close'), onclick: run });

/** The Chart and Chart data sections of the format options. */
function chartSections(el, theme, upd) {
  const c = el.chart;
  const edit = (fn) => upd((e) => { if (e.type === 'chart' && e.chart) fn(e.chart, e); });
  const pie = c.kind === 'pie' || c.kind === 'doughnut';
  const title = h('input', { type: 'text', placeholder: 'No title' });
  title.value = c.title || '';
  title.addEventListener('change', () => edit((ch) => { ch.title = title.value; }));
  const options = h('div', { class: 'pf-stack' },
    field('Type', select(CHART_KINDS.map((k) => [k.id, k.name]), c.kind, (v) => edit((ch) => { ch.kind = v; }))),
    ['column', 'bar', 'area'].includes(c.kind)
      ? field('Grouping', select([['clustered', c.kind === 'area' ? 'Standard' : 'Clustered'], ['stacked', 'Stacked'], ['percent', '100% stacked']], c.grouping || 'clustered', (v) => edit((ch) => { ch.grouping = v; })))
      : null,
    field('Title', title),
    field('Legend', select([['right', 'Right'], ['bottom', 'Bottom'], ['top', 'Top'], ['left', 'Left'], ['none', 'None']], c.legend || 'none', (v) => edit((ch) => { ch.legend = v; }))),
    checkbox('Data labels', c.dataLabels, (v) => edit((ch) => { ch.dataLabels = v; })),
    c.kind === 'line' ? checkbox('Markers', c.markers, (v) => edit((ch) => { ch.markers = v; })) : null,
    field('Text size', numberInput(el.style?.size || 12, (v) => edit((ch, e) => { e.style = { ...e.style, size: v }; }), { min: 4, max: 72, suffix: 'pt' })),
    h('div', { class: 'pf-row' },
      colorButton('Text colour', el.style?.color || 'theme:text', theme, (v) => edit((ch, e) => { e.style = { ...e.style, color: v }; })),
      colorButton('Background', el.fill || 'none', theme, (v) => edit((ch, e) => { e.fill = v; }), { allowNone: true, noneLabel: 'No background' })),
  );

  // The data: a row per category and a column per series (only the first for pies).
  const series = pie ? c.series.slice(0, 1) : c.series;
  const head = h('tr', {}, h('th', {}, c.kind === 'scatter' ? 'X' : 'Category'), ...series.map((s, k) => h('th', {},
    cellInput(s.name, `Name of series ${k + 1}`, (v) => edit((ch) => { ch.series[k].name = v; })),
    pie ? null : h('div', { class: 'chart-head' },
      swatch(`Colour of ${s.name}`, seriesColor(c, k, theme), theme, (v) => edit((ch) => { ch.series[k].color = v; })),
      c.series.length > 1 ? removeButton(`Remove ${s.name}`, () => edit((ch) => { ch.series.splice(k, 1); })) : null))),
  pie ? h('th', {}) : null, h('th', {}));
  const rows = c.categories.map((cat, i) => h('tr', {},
    h('td', {}, cellInput(cat, `Category ${i + 1}`, (v) => edit((ch) => { ch.categories[i] = v; }))),
    ...series.map((s, k) => h('td', {}, cellInput(s.values[i] ?? '', `${s.name}, ${cat}`, (v) => edit((ch) => {
      const list = ch.series[k].values;
      while (list.length <= i) list.push(null);
      list[i] = parseValue(v);
    }), true))),
    pie ? h('td', {}, swatch(`Colour of ${cat}`, pointColor(c, i, theme), theme, (v) => edit((ch) => {
      const first = ch.series[0];
      first.pointColors = ch.categories.map((_, j) => first.pointColors?.[j] || null);
      first.pointColors[i] = v;
    }))) : null,
    h('td', {}, c.categories.length > 1 ? removeButton(`Remove ${cat || `row ${i + 1}`}`, () => edit((ch) => {
      ch.categories.splice(i, 1);
      for (const s of ch.series) {
        s.values.splice(i, 1);
        s.pointColors?.splice(i, 1);
      }
    })) : null)));
  const grid = h('div', { class: 'chart-data' }, h('table', {}, h('thead', {}, head), h('tbody', {}, ...rows)));
  const buttons = h('div', { class: 'pf-row' },
    h('button', { type: 'button', class: 'btn', html: `${icon('plus')}<span>Row</span>`, onclick: () => edit((ch) => {
      ch.categories.push(ch.kind === 'scatter' ? String(ch.categories.length + 1) : `Category ${ch.categories.length + 1}`);
      for (const s of ch.series) while (s.values.length < ch.categories.length) s.values.push(null);
    }) }),
    pie ? null : h('button', { type: 'button', class: 'btn', html: `${icon('plus')}<span>Series</span>`, onclick: () => edit((ch) => {
      ch.series.push({ name: `Series ${ch.series.length + 1}`, values: ch.categories.map(() => null) });
    }) }));
  return [
    section('Chart', options),
    section('Chart data', grid, buttons, h('p', { class: 'pf-hint' }, pie ? 'Pie and doughnut charts show the first series.' : 'Leave a cell empty for no point.')),
  ];
}
