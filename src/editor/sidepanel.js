import { h, rovingFocus } from '../ui/dom.js';
import { icon } from '../ui/icons.js';
import { openColorPicker } from '../ui/color-picker.js';
import { THEMES, backgroundBaseColor, editBackgroundGradient } from '../model/themes.js';
import { SIZES, namedSize, slideSize, LAYOUTS, TRANSITIONS, ANIMATIONS, EMPHASES, EXITS, MOTIONS, slideAnimations, CHART_KINDS, newSlide, fillPlaceholders, canHoldText, slideText } from '../model/deck.js';
import { SHAPES } from '../model/shapes.js';
import { BORDER_PRESETS, cellMargin } from '../model/tables.js';
import { renderThumb, textStyle } from '../render/render.js';
import { seriesColor, pointColor, seriesKind, onSecondaryAxis, labelParts, scatterLines, showsMarkers, COMBO_KINDS } from '../render/chart.js';
import { formatValue } from '../render/number-format.js';
import { commentsPanel } from './comments.js';
import { diagramSection } from './diagrams.js';
import { textEffectOf } from './richtext.js';
import { WARP_GROUPS, mainAdjust } from '../model/warps.js';
import { field, section, numberInput, select, groupedSelect, checkbox, colorButton, DASH_STYLES } from './fields.js';

const TITLES = { design: 'Design', transition: 'Transition', animate: 'Animate', format: 'Format options', comments: 'Comments' };
const TAB_LABELS = { design: 'Design', transition: 'Transition', animate: 'Animate', format: 'Format', comments: 'Comments' };

// Line ends, as PowerPoint names them.
const ARROW_ENDS = [['none', 'None'], ['triangle', 'Arrow'], ['arrow', 'Open arrow'], ['stealth', 'Stealth arrow'], ['diamond', 'Diamond'], ['oval', 'Round']];
const endValue = (end) => (!end ? 'none' : end === true ? 'triangle' : end);
const endOf = (v) => (v === 'none' ? false : v === 'triangle' ? true : v);

/** Set or clear one of a picture's colour adjustments. */
function setAdjust(e, key, value) {
  e.adjust = { ...(e.adjust || {}), [key]: value };
  if (!e.adjust[key]) delete e.adjust[key];
  if (!Object.keys(e.adjust).length) delete e.adjust;
}

/** A slider for a picture's brightness or contrast (-100 to 100, 0 unchanged). */
function adjustSlider(label, el, key, upd) {
  const input = h('input', { type: 'range', min: '-100', max: '100', value: String(Math.round((el.adjust?.[key] || 0) * 100)) });
  const out = h('output', {}, input.value);
  input.addEventListener('input', () => {
    out.textContent = input.value;
    upd((e) => { setAdjust(e, key, +input.value / 100 || undefined); }, key);
  });
  return field(label, h('div', { class: 'pf-row' }, input, out));
}

/** The WordArt warp of text boxes (see model/warps.js), and how much it bends, for one. */
function warpControls(app, els) {
  const presets = new Set(els.map((e) => e.style?.warp?.preset || 'none'));
  const value = presets.size === 1 ? [...presets][0] : 'mixed';
  // A plain warp narrowed at one end comes only from PowerPoint files.
  const extra = value === 'mixed' ? [['mixed', 'Mixed']] : value === 'textPlain' ? [['textPlain', 'Plain (narrowed)']] : [];
  const out = [field('Warp', groupedSelect([...extra, ['none', 'None']], WARP_GROUPS, value,
    (v) => app.setWarp(v === 'none' ? null : { preset: v })))];
  const amount = els.length === 1 && mainAdjust(els[0].style?.warp);
  if (amount) {
    const warp = els[0].style.warp;
    const input = h('input', { type: 'range', min: String(amount.min), max: String(amount.max), step: 'any', value: String(amount.value) });
    const pct = () => `${Math.round(((input.value - amount.min) / (amount.max - amount.min || 1)) * 100)}%`;
    const shown = h('output', {}, pct());
    input.addEventListener('input', () => {
      shown.textContent = pct();
      app.setWarp({ ...warp, adjust: { ...warp.adjust, [amount.name]: +input.value } }, { coalesce: 'warp-amount' });
    });
    out.push(field('Amount', h('div', { class: 'pf-row' }, input, shown)));
  }
  return out;
}

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

/** How long an animation takes, with the one it has if it isn't one of these (from a PowerPoint file). */
function durationOptions(current) {
  const list = [[0.3, 'Very fast (0.3 s)'], [0.5, 'Fast (0.5 s)'], [0.8, 'Medium (0.8 s)'], [1.2, 'Slow (1.2 s)'], [2, 'Very slow (2 s)']];
  return list.some(([v]) => v === current) || !Number.isFinite(current) ? list : [...list, [current, `${Math.round(current * 100) / 100} s`]].sort((a, b) => a[0] - b[0]);
}

/**
 * The right-hand panel: themes and backgrounds, transitions, animations,
 * the detailed format options for the selection and the slide's comments.
 */
export class SidePanel {
  constructor(app) {
    this.app = app;
    this.current = null;
    this.key = '';
    this.activeComment = null; // the comment thread highlighted here and on the slide
    const ids = `panel-${Math.random().toString(36).slice(2)}`;
    this.title = h('h2', { id: `${ids}-title` });
    this.body = h('div', { class: 'panel-body', id: `${ids}-body`, role: 'tabpanel', 'aria-labelledby': `${ids}-title` });
    this.tabs = h('div', { class: 'panel-tabs', role: 'tablist', 'aria-label': 'Side panel' }, ...Object.entries(TAB_LABELS).map(([id, name]) => (
      h('button', { type: 'button', role: 'tab', class: 'panel-tab', 'aria-controls': this.body.id, dataset: { tab: id }, onclick: () => this.open(id) }, name))));
    // One tab stop: the arrow keys move between the tabs and open them.
    rovingFocus(this.tabs, () => [...this.tabs.children], { onMove: (t) => this.open(t.dataset.tab) });
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
      t.tabIndex = t.dataset.tab === this.current ? 0 : -1;
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
    // A pattern (from a PowerPoint file) can be kept, or changed for another kind of fill.
    const bgKinds = [['theme', 'Theme default'], ['solid', 'Solid colour'], ['gradient', 'Gradient'], ['image', 'Picture'], ...(bgType === 'pattern' ? [['pattern', 'Pattern']] : [])];
    bgControls.append(field('Fill', select(bgKinds, bgType, (v) => {
      if (v === 'pattern') return;
      if (v === 'theme') app.setBackground(null);
      else if (v === 'solid') app.setBackground({ type: 'solid', color: backgroundBaseColor(theme.background) });
      else if (v === 'gradient') app.setBackground({ type: 'gradient', from: 'theme:accent1', to: 'theme:accent2', angle: 135 });
      else app.pickBackgroundImage();
    })));
    if (bgType === 'solid') bgControls.append(colorButton('Colour', bg.color, theme, (v) => app.setBackground({ type: 'solid', color: v })));
    if (bgType === 'gradient') {
      bgControls.append(
        h('div', { class: 'pf-row' },
          colorButton('From', bg.from, theme, (v) => app.setBackground(editBackgroundGradient(bg, { from: v }))),
          colorButton('To', bg.to, theme, (v) => app.setBackground(editBackgroundGradient(bg, { to: v })))),
        field('Angle', select([[0, 'Bottom to top'], [90, 'Left to right'], [135, 'Diagonal'], [180, 'Top to bottom'], [225, 'Diagonal (reverse)']], bg.angle ?? 180, (v) => app.setBackground(editBackgroundGradient(bg, { angle: +v })))),
      );
    }
    if (bgType === 'pattern') {
      bgControls.append(h('div', { class: 'pf-row' },
        colorButton('Foreground', bg.pattern.fg, theme, (v) => app.setBackground({ ...bg, pattern: { ...bg.pattern, fg: v } })),
        colorButton('Background', bg.pattern.bg, theme, (v) => app.setBackground({ ...bg, pattern: { ...bg.pattern, bg: v } }))));
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
      section('Slide size', select([
        ...Object.entries(SIZES).map(([id, s]) => [id, s.label]),
        ...(namedSize(deck.size) ? [] : [[deck.size, slideSize(deck).label]]),
        ['custom', 'Custom…'],
      ], deck.size, (v) => (v === 'custom' ? app.customSizePrompt() : app.setSize(v)), 'Slide size'),
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
      // One section each for how it comes in, draws attention and goes out.
      const motions = MOTIONS.filter((m) => m.id !== 'custom' || els[0].motion?.effect === 'custom');
      for (const [kind, title, effects] of [['anim', 'Entrance', ANIMATIONS], ['emphasis', 'Emphasis', EMPHASES], ['exit', 'Exit', EXITS], ['motion', 'Motion path', motions]]) {
        const a = els[0][kind] || { effect: 'none', duration: 0.5 };
        const on = a.effect && a.effect !== 'none';
        out.push(section(els.length > 1 ? `${title} (${els.length} objects)` : title,
          field('Effect', select(effects.map((x) => [x.id, x.name]), a.effect, (v) => app.setAnimation({ effect: v }, null, kind))),
          on ? field('Duration', select(durationOptions(a.duration ?? 0.5), a.duration ?? 0.5, (v) => app.setAnimation({ duration: +v }, null, kind))) : null,
          on ? field('Start', select([['click', 'On click'], ['with', 'With previous'], ['after', 'After previous']], a.start || 'click', (v) => app.setAnimation({ start: v }, null, kind))) : null));
      }
    } else {
      out.push(h('p', { class: 'pf-hint' }, 'Select an object on the slide to animate it.'));
    }
    const animated = slideAnimations(app.slide);
    const list = h('ol', { class: 'anim-list' });
    const effectName = { anim: ANIMATIONS, emphasis: EMPHASES, exit: EXITS, motion: MOTIONS };
    animated.forEach(({ el, kind, anim }, i) => {
      const name = (canHoldText(el) && slideText({ elements: [el] })) || ({ image: 'Picture', line: 'Line', table: 'Table', shape: 'Shape', text: 'Text box', chart: 'Chart', diagram: 'Diagram', media: el.kind === 'audio' ? 'Sound' : 'Video' }[el.type]) || 'Object';
      const effect = effectName[kind].find((x) => x.id === anim.effect)?.name;
      list.append(h('li', { class: `anim-item${app.selection.includes(el.id) ? ' is-selected' : ''}` },
        h('button', { type: 'button', class: 'anim-name', onclick: () => app.setSelection([el.id]) }, h('b', {}, name.slice(0, 40)), h('small', {}, `${effect}${anim.start && anim.start !== 'click' ? ` · ${anim.start} previous` : ''}`)),
        h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Move earlier', disabled: i === 0 || null, html: icon('up'), onclick: () => app.reorderAnimation(el.id, -1, kind) }),
        h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Move later', disabled: i === animated.length - 1 || null, html: icon('down'), onclick: () => app.reorderAnimation(el.id, 1, kind) }),
        h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Remove animation', html: icon('close'), onclick: () => app.setAnimation({ effect: 'none' }, [el.id], kind) })));
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
        field('Start', select(ARROW_ENDS, endValue(el.arrowStart), (v) => upd((e) => { e.arrowStart = endOf(v); }))),
        field('End', select(ARROW_ENDS, endValue(el.arrowEnd), (v) => upd((e) => { e.arrowEnd = endOf(v); })))));
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
          h('button', { type: 'button', class: 'btn', onclick: () => app.align('center', { toSlide: true }) }, 'Centre horizontally'),
          h('button', { type: 'button', class: 'btn', onclick: () => app.align('middle', { toSlide: true }) }, 'Centre vertically'))));
    }

    if (els.every((e) => e.type !== 'line' && e.type !== 'diagram')) {
      const fillable = els.some((e) => e.type === 'shape' || e.type === 'text' || e.type === 'table');
      const look = h('div', { class: 'pf-stack' });
      if (fillable) look.append(colorButton(el.type === 'table' ? 'Header colour' : 'Fill', el.fill, theme, (v) => app.setElementProp('fill', v), { allowNone: true }));
      if (fillable && el.type !== 'table' && el.fill !== 'none' && !el.fillImage) look.append(...gradientControls(app, el, theme));
      if (fillable && el.type !== 'table') {
        look.append(h('div', { class: 'pf-row' },
          h('button', { type: 'button', class: 'btn', onclick: () => app.pickFillImage() }, el.fillImage ? 'Change fill picture…' : 'Fill with a picture…'),
          el.fillImage ? h('button', { type: 'button', class: 'btn btn-ghost', onclick: () => app.removeFillImage() }, 'Remove picture') : null));
      }
      const border = el.stroke?.width && el.stroke.color && el.stroke.color !== 'none' ? el.stroke : null;
      look.append(
        colorButton('Border', border ? border.color : 'none', theme, (v) => app.setStroke({ color: v, width: Math.max(1, el.stroke?.width || 2) }), { allowNone: true, noneLabel: 'No border' }),
        field('Border weight', numberInput(border ? border.width : 0, (v) => app.setBorderWeight(v), { min: 0, max: 50, step: 0.5, suffix: 'pt' })),
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
      const effects = new Set(els.map((e) => textEffectOf(e.text || (e.cells || []).flat().map((c) => c.text).join(''))));
      const effect = effects.size === 1 ? [...effects][0] : 'mixed';
      out.push(section('Text',
        field('Vertical alignment', select([['top', 'Top'], ['middle', 'Middle'], ['bottom', 'Bottom']], t.valign, (v) => app.setTextProp('valign', v))),
        field('Line spacing', select([[1, 'Single'], [1.15, '1.15'], [1.25, '1.25'], [1.5, '1.5'], [2, 'Double']], t.lineHeight, (v) => app.setTextProp('lineHeight', +v))),
        field('Space between paragraphs', select([[0, 'None'], [0.35, 'Small'], [0.8, 'Medium'], [1.4, 'Large']], t.spacing, (v) => app.setTextProp('spacing', +v))),
        field('Effect', select([
          ...(effect === 'mixed' ? [['mixed', 'Mixed']] : []),
          ['none', 'None'], ['shadow', 'Shadow'], ['glow', 'Glow'], ['outline', 'Outline'], ['hollow', 'Hollow'], ['gradient', 'Gradient'],
        ], effect, (v) => app.setTextEffect(v))),
        ...(els.every(canHoldText) ? warpControls(app, els) : [])));
    }

    if (single && el.type === 'image') {
      const alt = h('textarea', { rows: 2, placeholder: 'Describe the picture for people who can’t see it' });
      alt.value = el.alt || '';
      alt.addEventListener('change', () => upd((e) => { e.alt = alt.value; }));
      out.push(section('Picture',
        field('Fit', select([['fill', 'Stretch to box'], ['cover', 'Fill box (crop)'], ['contain', 'Fit inside box']], el.fit || 'fill', (v) => upd((e) => { e.fit = v; }))),
        field('Corner radius', numberInput(el.radius || 0, (v) => upd((e) => { e.radius = v; }), { min: 0, suffix: 'pt' })),
        field('Crop to shape', select([
          ['rect', 'None'],
          ...(el.clip?.shape === 'custom' ? [['custom', 'PowerPoint shape']] : []),
          ...SHAPES.filter((s) => s.id !== 'rect' && s.group !== 'Callouts').map((s) => [s.id, s.name]),
        ], el.clip?.shape || 'rect', (v) => upd((e) => {
          if (v === 'custom') return;
          if (v === 'rect') delete e.clip;
          else e.clip = { shape: v };
        }))),
        checkbox('Greyscale', el.adjust?.grayscale, (v) => upd((e) => { setAdjust(e, 'grayscale', v || undefined); })),
        adjustSlider('Brightness', el, 'brightness', upd),
        adjustSlider('Contrast', el, 'contrast', upd),
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
      out.push(cellSection(this, el, theme));
    }

    if (single && el.type === 'chart' && el.chart) out.push(...chartSections(el, theme, upd));
    // A diagram's text and layout first: they're what it's edited for.
    if (single && el.type === 'diagram' && el.diagram) out.unshift(diagramSection(this, el));

    out.push(section('Arrange', checkbox('Lock position', els.every((e) => e.locked), (v) => upd((e) => { e.locked = v; }))));
    return out;
  }
}

// ------------------------------------------------------------------ table cells

/**
 * Borders and margins for the cells being edited (from the cursor's cell to
 * the end of the selection), or the whole table. The border's colour, weight
 * and style are a pen kept by the panel; a choice of borders draws with it.
 */
function cellSection(panel, el, theme) {
  const app = panel.app;
  const st = el.stroke?.color && el.stroke.color !== 'none' && el.stroke.width ? el.stroke : null;
  const pen = (panel.borderPen ||= { color: st?.color || 'theme:text', width: st?.width || 1, dash: 'solid' });
  const borders = select([['', 'Choose…'], ...BORDER_PRESETS], '', (v) => {
    if (v) app.setCellBorders(v, pen);
  });
  const at = app.tableCells();
  const m = cellMargin(at ? el.cells[at.range.r0][at.range.c0] : null);
  const margin = (label, side) => field(label, numberInput(m[side], (v) => app.setCellMargins({ [side]: v }), { min: 0, max: 200, step: 0.5, suffix: 'pt' }));
  return section('Cells',
    field('Borders', borders),
    colorButton('Border colour', pen.color, theme, (v) => { pen.color = v; panel.update(true); }),
    field('Border weight', numberInput(pen.width, (v) => { pen.width = v; }, { min: 0.25, max: 20, step: 0.25, suffix: 'pt' })),
    field('Border dash', select(DASH_STYLES, pen.dash, (v) => { pen.dash = v; })),
    h('div', { class: 'pf-grid' }, margin('Left margin', 'l'), margin('Right margin', 'r'), margin('Top margin', 't'), margin('Bottom margin', 'b')),
    h('p', { class: 'pf-hint' }, 'Applies to the cells you’re typing in (or have selected), or to the whole table.'));
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
  const lines = c.kind === 'line' || c.series.some((s, k) => seriesKind(c, k) === 'line');
  const options = h('div', { class: 'pf-stack' },
    field('Type', select(CHART_KINDS.map((k) => [k.id, k.name]), c.kind, (v) => edit((ch) => { ch.kind = v; }))),
    ['column', 'bar', 'area', 'line'].includes(c.kind)
      ? field('Grouping', select([['clustered', c.kind === 'area' || c.kind === 'line' ? 'Standard' : 'Clustered'], ['stacked', 'Stacked'], ['percent', '100% stacked']], c.grouping || 'clustered', (v) => edit((ch) => { ch.grouping = v; })))
      : null,
    field('Title', title),
    field('Legend', select([['right', 'Right'], ['bottom', 'Bottom'], ['top', 'Top'], ['left', 'Left'], ['none', 'None']], c.legend || 'none', (v) => edit((ch) => { ch.legend = v; }))),
    checkbox('Data labels', c.dataLabels, (v) => edit((ch) => { ch.dataLabels = v; })),
    pie && c.dataLabels ? field('Labels show', select(LABEL_CONTENTS, labelParts(c).join(), (v) => edit((ch) => {
      ch.labelContent = v.split(',');
      if (v === 'value') delete ch.labelContent;
    }))) : null,
    c.kind === 'scatter' ? checkbox('Lines', scatterLines(c), (v) => edit((ch) => { ch.lines = v || undefined; })) : null,
    lines || c.kind === 'scatter' ? checkbox('Markers', showsMarkers(c), (v) => edit((ch) => { ch.markers = v; })) : null,
    lines || scatterLines(c) ? checkbox('Smooth lines', c.smooth, (v) => edit((ch) => { ch.smooth = v || undefined; })) : null,
    field('Text size', numberInput(el.style?.size || 12, (v) => edit((ch, e) => { e.style = { ...e.style, size: v }; }), { min: 4, max: 72, suffix: 'pt' })),
    h('div', { class: 'pf-row' },
      colorButton('Text colour', el.style?.color || 'theme:text', theme, (v) => edit((ch, e) => { e.style = { ...e.style, color: v }; })),
      colorButton('Background', el.fill || 'none', theme, (v) => edit((ch, e) => { e.fill = v; }), { allowNone: true, noneLabel: 'No background' })),
  );

  // The data: a row per category and a column per series (only the first for pies; a ring each for doughnuts).
  const series = c.kind === 'pie' ? c.series.slice(0, 1) : c.series;
  const head = h('tr', {}, h('th', {}, c.kind === 'scatter' ? 'X' : 'Category'), ...series.map((s, k) => h('th', {},
    cellInput(s.name, `Name of series ${k + 1}`, (v) => edit((ch) => { ch.series[k].name = v; })),
    c.kind === 'pie' ? null : h('div', { class: 'chart-head' },
      pie ? null : swatch(`Colour of ${s.name}`, seriesColor(c, k, theme), theme, (v) => edit((ch) => { ch.series[k].color = v; })),
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
    c.kind === 'pie' ? null : h('button', { type: 'button', class: 'btn', html: `${icon('plus')}<span>Series</span>`, onclick: () => edit((ch) => {
      ch.series.push({ name: `Series ${ch.series.length + 1}`, values: ch.categories.map(() => null) });
    }) }));
  return [
    section('Chart', options),
    chartMore(c, edit),
    section('Chart data', grid, buttons, h('p', { class: 'pf-hint' }, c.kind === 'pie' ? 'Pie charts show the first series.' : c.kind === 'doughnut' ? 'Each series is a ring, the first innermost.' : 'Leave a cell empty for no point.')),
  ];
}

/** Number formats offered for a chart's values: Excel codes with an example of each. */
const NUMBER_FORMATS = [['General', 'General'], ['#,##0', '1,234'], ['#,##0.0', '1,234.5'], ['#,##0.00', '1,234.56'], ['0%', '12%'], ['0.0%', '12.3%'], ['$#,##0', '$1,234'], ['#,##0 €', '1,234 €']];
const LABEL_CONTENTS = [['value', 'Value'], ['percent', 'Percentage'], ['category', 'Category name'], ['category,percent', 'Category and percentage'], ['category,value', 'Category and value']];
/** Where data labels can go, by kind of chart ('' is PowerPoint's automatic place). */
const LABEL_PLACES = {
  bar: [['', 'Automatic'], ['outEnd', 'Outside end'], ['inEnd', 'Inside end'], ['ctr', 'Centre'], ['inBase', 'Inside base']],
  line: [['', 'Automatic'], ['t', 'Above'], ['b', 'Below'], ['l', 'Left'], ['r', 'Right'], ['ctr', 'Centre']],
  pie: [['', 'Best fit'], ['outEnd', 'Outside end'], ['inEnd', 'Inside end'], ['ctr', 'Centre']],
};

// Whether "More chart options" is open, kept as the panel is redrawn after each change.
let chartMoreOpen = false;

/** An axis bound or unit: a number, or blank for automatic. */
function autoNumber(value, label, onValue) {
  const input = h('input', { type: 'number', step: 'any', placeholder: 'Auto', 'aria-label': label, value: Number.isFinite(value) ? String(value) : '' });
  input.addEventListener('change', () => onValue(input.value.trim() === '' || !Number.isFinite(+input.value) ? undefined : +input.value));
  return input;
}

/** Set (or clear, when blank) one setting of an axis, dropping the axis when it has none left. */
function setAxis(ch, key, prop, value) {
  const axis = { ...(ch[key] || {}), [prop]: value };
  if (value === undefined || value === '') delete axis[prop];
  if (Object.keys(axis).length) ch[key] = axis;
  else delete ch[key];
}

/** The advanced chart options: number format, label placement, axes, bar spacing and combinations. */
function chartMore(c, edit) {
  const pie = c.kind === 'pie' || c.kind === 'doughnut';
  const bars = c.kind === 'column' || c.kind === 'bar' || c.series.some((s, k) => seriesKind(c, k) === 'column');
  const stacked = c.grouping === 'stacked' || c.grouping === 'percent';
  const text = (value, label, onValue) => {
    const input = h('input', { type: 'text', placeholder: 'None', 'aria-label': label });
    input.value = value || '';
    input.addEventListener('change', () => onValue(input.value.trim()));
    return input;
  };
  const formats = NUMBER_FORMATS.some(([code]) => code === (c.numberFormat || 'General')) ? NUMBER_FORMATS : [...NUMBER_FORMATS, [c.numberFormat, `${formatValue(1234.5, c.numberFormat)} (from file)`]];
  const places = c.kind === 'pie' ? LABEL_PLACES.pie
    : c.kind === 'column' || c.kind === 'bar' ? LABEL_PLACES.bar.filter(([v]) => !stacked || v !== 'outEnd')
      : c.kind === 'line' || c.kind === 'scatter' ? LABEL_PLACES.line : null;
  const axisRow = (key) => h('div', { class: 'pf-row chart-axis-scale' },
    ...[['min', 'Minimum'], ['max', 'Maximum'], ['unit', 'Major unit']].map(([prop, label]) => field(label, autoNumber(c[key]?.[prop], label, (v) => edit((ch) => setAxis(ch, key, prop, v))))));
  const combo = COMBO_KINDS.includes(c.kind) && c.series.length > 1;
  const secondary = c.series.some((s, k) => onSecondaryAxis(c, k));

  const body = h('div', { class: 'pf-stack' },
    field('Number format', select(formats, c.numberFormat || 'General', (v) => edit((ch) => {
      if (v === 'General') delete ch.numberFormat;
      else ch.numberFormat = v;
      delete ch.labelFormat;
    }))),
    c.dataLabels && places ? field('Label position', select(places, places.some(([v]) => v === c.labelPosition) ? c.labelPosition : '', (v) => edit((ch) => {
      if (v) ch.labelPosition = v;
      else delete ch.labelPosition;
    }))) : null,
    pie ? null : field(c.kind === 'scatter' ? 'Y axis title' : 'Value axis title', text(c.valueAxis?.title, 'Value axis title', (v) => edit((ch) => setAxis(ch, 'valueAxis', 'title', v)))),
    pie ? null : field(c.kind === 'scatter' ? 'X axis title' : 'Category axis title', text(c.categoryAxis?.title, 'Category axis title', (v) => edit((ch) => setAxis(ch, 'categoryAxis', 'title', v)))),
    pie || c.grouping === 'percent' ? null : axisRow('valueAxis'),
    pie ? null : checkbox('Gridlines', c.valueAxis?.gridlines !== false, (v) => edit((ch) => setAxis(ch, 'valueAxis', 'gridlines', v ? undefined : false))),
    bars ? h('div', { class: 'pf-row chart-axis-scale' },
      field('Gap width', numberInput(c.gapWidth ?? 150, (v) => edit((ch) => { ch.gapWidth = v; }), { min: 0, max: 500, suffix: '%' })),
      stacked ? null : field('Overlap', numberInput(c.overlap ?? 0, (v) => edit((ch) => { ch.overlap = v; }), { min: -100, max: 100, suffix: '%' }))) : null,
    c.kind === 'doughnut' ? field('Hole size', numberInput(c.holeSize ?? 50, (v) => edit((ch) => { ch.holeSize = v; }), { min: 10, max: 90, suffix: '%' })) : null,
    // Combination charts: each series as columns, a line or an area, on either value axis.
    combo ? h('div', { class: 'pf-stack chart-combo' }, h('span', { class: 'pf-label' }, 'Series'),
      ...c.series.map((s, k) => h('div', { class: 'pf-row' },
        h('span', { class: 'chart-combo-name' }, s.name || `Series ${k + 1}`),
        select([['', 'Chart type'], ['column', 'Columns'], ['line', 'Line'], ['area', 'Area']], s.kind && s.kind !== c.kind ? s.kind : '', (v) => edit((ch) => {
          if (v && v !== ch.kind) ch.series[k].kind = v;
          else delete ch.series[k].kind;
        }), `How ${s.name || `series ${k + 1}`} is drawn`),
        checkbox('Secondary axis', s.secondary, (v) => edit((ch) => {
          if (v) ch.series[k].secondary = true;
          else delete ch.series[k].secondary;
        }))))) : null,
    secondary ? field('Secondary axis title', text(c.secondaryAxis?.title, 'Secondary axis title', (v) => edit((ch) => setAxis(ch, 'secondaryAxis', 'title', v)))) : null,
  );
  const details = h('details', { class: 'panel-section chart-more', open: chartMoreOpen || null }, h('summary', {}, 'More chart options'), body);
  details.addEventListener('toggle', () => { chartMoreOpen = details.open; });
  return details;
}
