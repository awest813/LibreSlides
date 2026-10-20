import { h, showPopover, closePopover } from './dom.js';
import { THEME_COLOR_KEYS, resolveColor } from '../model/themes.js';

const STANDARD = [
  '#000000', '#434343', '#666666', '#999999', '#b7b7b7', '#cccccc', '#d9d9d9', '#efefef', '#f3f3f3', '#ffffff',
  '#980000', '#ff0000', '#ff9900', '#ffff00', '#00ff00', '#00ffff', '#4a86e8', '#0000ff', '#9900ff', '#ff00ff',
  '#e6b8af', '#f4cccc', '#fce5cd', '#fff2cc', '#d9ead3', '#d0e0e3', '#c9daf8', '#cfe2f3', '#d9d2e9', '#ead1dc',
  '#dd7e6b', '#ea9999', '#f9cb9c', '#ffe599', '#b6d7a8', '#a2c4c9', '#a4c2f4', '#9fc5e8', '#b4a7d6', '#d5a6bd',
  '#cc4125', '#e06666', '#f6b26b', '#ffd966', '#93c47d', '#76a5af', '#6d9eeb', '#6fa8dc', '#8e7cc3', '#c27ba0',
  '#a61c00', '#cc0000', '#e69138', '#f1c232', '#6aa84f', '#45818e', '#3c78d8', '#3d85c6', '#674ea7', '#a64d79',
  '#5b0f00', '#660000', '#783f04', '#7f6000', '#274e13', '#0c343d', '#1c4587', '#073763', '#20124d', '#4c1130',
];

const NAMES = { title: 'Title', text: 'Text', accent1: 'Accent 1', accent2: 'Accent 2', accent3: 'Accent 3', accent4: 'Accent 4', accent5: 'Accent 5', accent6: 'Accent 6' };

let recent = [];

/**
 * A colour menu anchored to `anchor`: theme colours (which follow the theme),
 * a standard palette, recent and custom colours. Calls onPick(value) with
 * "theme:<key>", "#rrggbb" or "none".
 */
export function openColorPicker(anchor, { theme, onPick, allowNone = false, noneLabel = 'Transparent', current = null, onClose = null }) {
  const pick = (v) => {
    closePopover();
    if (v && v.startsWith('#')) recent = [v, ...recent.filter((c) => c !== v)].slice(0, 10);
    onPick(v);
  };
  const chip = (value, label) => {
    const color = resolveColor(value, theme);
    return h('button', {
      type: 'button', class: `color-chip${value === current ? ' is-current' : ''}`, title: label, 'aria-label': label,
      style: { background: color }, onclick: () => pick(value),
    });
  };
  const panel = h('div', { class: 'color-panel' });
  if (allowNone) {
    panel.append(h('button', { type: 'button', class: 'menu-item', onclick: () => pick('none') }, h('span', { class: 'menu-icon', html: '<span class="none-swatch"></span>' }), h('span', { class: 'menu-label' }, noneLabel), h('span')));
  }
  panel.append(h('div', { class: 'color-heading' }, 'Theme'));
  panel.append(h('div', { class: 'color-grid cols-8' }, ...THEME_COLOR_KEYS.map((k) => chip(`theme:${k}`, NAMES[k]))));
  panel.append(h('div', { class: 'color-heading' }, 'Standard'));
  panel.append(h('div', { class: 'color-grid' }, ...STANDARD.map((c) => chip(c, c))));
  if (recent.length) {
    panel.append(h('div', { class: 'color-heading' }, 'Recent'));
    panel.append(h('div', { class: 'color-grid' }, ...recent.map((c) => chip(c, c))));
  }
  const start = current && current.startsWith?.('#') ? current : resolveColor(current, theme, '#000000');
  const input = h('input', { type: 'color', value: /^#[0-9a-f]{6}$/i.test(start) ? start : '#000000', 'aria-label': 'Custom colour' });
  input.addEventListener('change', () => pick(input.value));
  panel.append(h('label', { class: 'color-custom' }, input, 'Custom…'));
  return showPopover(anchor, panel, { label: anchor?.getAttribute?.('aria-label') || 'Colour', onClose });
}
