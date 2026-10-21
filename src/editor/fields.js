import { h } from '../ui/dom.js';
import { openColorPicker } from '../ui/color-picker.js';
import { resolveColor } from '../model/themes.js';

/**
 * Form controls for the side panel (format options, diagrams, charts…):
 * each change is reported through a callback once it's made, so every
 * change is one undo step.
 */

/** A labelled control: the label above it, and an optional hint below. */
export const field = (label, control, hint = null) => h('label', { class: 'pf' }, h('span', { class: 'pf-label' }, label), control, hint ? h('small', { class: 'pf-hint' }, hint) : null);

/** A titled section of the panel. */
export const section = (title, ...children) => h('section', { class: 'panel-section' }, h('h3', {}, title), ...children);

/** A number field, clamped to min and max; Enter or leaving it applies the value. */
export function numberInput(value, onValue, { step = 1, min = null, max = null, suffix = '' } = {}) {
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

/** A list of [value, text] options; `label` names it when it isn't inside a field(). */
export function select(options, value, onValue, label = null) {
  const s = h('select', { 'aria-label': label }, ...options.map(([v, text]) => h('option', { value: v, selected: String(v) === String(value) || null }, text)));
  s.addEventListener('change', () => onValue(s.value));
  return s;
}

/** A select of options in groups: [[group label, [[value, label]…]]…], with ungrouped `first` options before them. */
export function groupedSelect(first, groups, value, onValue) {
  const option = ([v, label]) => h('option', { value: v, selected: String(v) === String(value) || null }, label);
  const s = h('select', {}, ...first.map(option), ...groups.map(([label, options]) => h('optgroup', { label }, ...options.map(option))));
  s.addEventListener('change', () => onValue(s.value));
  return s;
}

export function checkbox(label, checked, onValue) {
  const input = h('input', { type: 'checkbox', checked: checked || null });
  input.addEventListener('change', () => onValue(input.checked));
  return h('label', { class: 'pf-check' }, input, label);
}

/** A button showing a colour that opens the colour menu (`opts` are the menu's: allowNone, noneLabel…). */
export function colorButton(label, value, theme, onPick, opts = {}) {
  // No colour shows as a struck-out swatch, as in the colour menu (an empty one is invisible in dark mode).
  const none = !value || value === 'none';
  const sw = h('span', { class: `pf-swatch${none ? ' is-none' : ''}`, style: none ? null : { background: resolveColor(value, theme, 'transparent') } });
  const b = h('button', { type: 'button', class: 'btn pf-color', 'aria-label': label, 'aria-haspopup': 'true' }, sw, h('span', {}, label));
  b.addEventListener('click', () => openColorPicker(b, { theme, current: value, onPick, ...opts }));
  return b;
}

/** Border and line dashes, as the toolbar's dash menu and the table cell borders name them. */
export const DASH_STYLES = [['solid', 'Solid'], ['dash', 'Dash'], ['dot', 'Dot'], ['dashDot', 'Dash-dot'], ['longDash', 'Long dash']];
