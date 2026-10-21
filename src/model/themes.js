import { patternSvg, PATTERN_SIZE } from './patterns.js';

/**
 * Themes set a deck's default background, colours and fonts. Elements refer
 * to theme colours as "theme:accent1", "theme:text"… so switching the theme
 * restyles everything that wasn't given an explicit colour.
 *
 * `decor` shapes are drawn behind a slide's content (positions are fractions
 * of the slide size); `title` decor is used on title and section slides,
 * `content` decor on the others.
 */
export const THEMES = [
  {
    id: 'simple-light', name: 'Simple Light',
    background: { type: 'solid', color: '#ffffff' },
    colors: { text: '#374151', title: '#111827', accent1: '#2563eb', accent2: '#f59e0b', accent3: '#10b981', accent4: '#ef4444', accent5: '#8b5cf6', accent6: '#06b6d4' },
    fonts: { title: 'Inter', body: 'Inter' }, titleBold: true,
    decor: { title: [{ shape: 'rect', x: 0.108, y: 0.625, w: 0.08, h: 0.008, fill: 'theme:accent1' }], content: [] },
  },
  {
    id: 'simple-dark', name: 'Simple Dark',
    background: { type: 'solid', color: '#111827' },
    colors: { text: '#d1d5db', title: '#ffffff', accent1: '#60a5fa', accent2: '#fbbf24', accent3: '#34d399', accent4: '#f87171', accent5: '#a78bfa', accent6: '#22d3ee' },
    fonts: { title: 'Inter', body: 'Inter' }, titleBold: true, dark: true,
    decor: { title: [{ shape: 'rect', x: 0.108, y: 0.625, w: 0.08, h: 0.008, fill: 'theme:accent1' }], content: [] },
  },
  {
    id: 'coral', name: 'Coral',
    background: { type: 'solid', color: '#fff8f3' },
    colors: { text: '#44403c', title: '#9a3412', accent1: '#f97316', accent2: '#e11d48', accent3: '#0d9488', accent4: '#facc15', accent5: '#7c3aed', accent6: '#0284c7' },
    fonts: { title: 'Montserrat', body: 'Inter' }, titleBold: true,
    decor: {
      title: [{ shape: 'ellipse', x: 0.68, y: -0.25, w: 0.5, h: 0.89, fill: 'theme:accent1', opacity: 0.9 }, { shape: 'ellipse', x: 0.82, y: 0.55, w: 0.25, h: 0.44, fill: 'theme:accent2', opacity: 0.85 }],
      content: [{ shape: 'rect', x: 0, y: 0, w: 0.012, h: 1, fill: 'theme:accent1' }],
    },
  },
  {
    id: 'midnight', name: 'Midnight',
    background: { type: 'gradient', from: '#0f172a', to: '#1e3a8a', angle: 135 },
    colors: { text: '#cbd5e1', title: '#ffffff', accent1: '#38bdf8', accent2: '#f472b6', accent3: '#a3e635', accent4: '#fb923c', accent5: '#c084fc', accent6: '#2dd4bf' },
    fonts: { title: 'Montserrat', body: 'Inter' }, titleBold: true, dark: true,
    decor: {
      title: [{ shape: 'rect', x: 0, y: 0.94, w: 1, h: 0.06, fill: 'theme:accent1' }],
      content: [{ shape: 'rect', x: 0.06, y: 0.235, w: 0.06, h: 0.007, fill: 'theme:accent1' }],
    },
  },
  {
    id: 'editorial', name: 'Editorial',
    background: { type: 'solid', color: '#faf7f0' },
    colors: { text: '#3f3f46', title: '#1c1917', accent1: '#9a3412', accent2: '#166534', accent3: '#1e40af', accent4: '#a16207', accent5: '#6b21a8', accent6: '#0f766e' },
    fonts: { title: 'Playfair Display', body: 'Lora' }, titleBold: false,
    decor: {
      title: [{ shape: 'rect', x: 0.05, y: 0.06, w: 0.9, h: 0.003, fill: 'theme:title' }, { shape: 'rect', x: 0.05, y: 0.94, w: 0.9, h: 0.003, fill: 'theme:title' }],
      content: [{ shape: 'rect', x: 0.06, y: 0.93, w: 0.88, h: 0.002, fill: 'theme:accent1' }],
    },
  },
  {
    id: 'forest', name: 'Forest',
    background: { type: 'solid', color: '#f1f6f2' },
    colors: { text: '#334155', title: '#14532d', accent1: '#16a34a', accent2: '#ca8a04', accent3: '#0369a1', accent4: '#b91c1c', accent5: '#7e22ce', accent6: '#0f766e' },
    fonts: { title: 'Roboto Slab', body: 'Inter' }, titleBold: true,
    decor: {
      title: [{ shape: 'rect', x: 0, y: 0, w: 0.36, h: 1, fill: 'theme:accent1', opacity: 0.12 }],
      content: [{ shape: 'rect', x: 0, y: 0, w: 1, h: 0.025, fill: 'theme:accent1' }],
    },
  },
  {
    id: 'bold', name: 'Bold',
    background: { type: 'solid', color: '#ffd23f' },
    colors: { text: '#1f1f1f', title: '#000000', accent1: '#000000', accent2: '#ee4266', accent3: '#3bceac', accent4: '#540d6e', accent5: '#0ead69', accent6: '#1d4ed8' },
    fonts: { title: 'Oswald', body: 'Inter' }, titleBold: true,
    decor: {
      title: [{ shape: 'rect', x: 0.08, y: 0.08, w: 0.84, h: 0.84, fill: 'none', stroke: 'theme:title', strokeWidth: 4 }],
      content: [{ shape: 'rect', x: 0, y: 0.9, w: 1, h: 0.1, fill: 'theme:accent1' }],
    },
  },
  {
    id: 'ocean', name: 'Ocean',
    background: { type: 'gradient', from: '#e0f2fe', to: '#ffffff', angle: 180 },
    colors: { text: '#334155', title: '#075985', accent1: '#0284c7', accent2: '#14b8a6', accent3: '#6366f1', accent4: '#f43f5e', accent5: '#eab308', accent6: '#22c55e' },
    fonts: { title: 'Montserrat', body: 'Inter' }, titleBold: true,
    decor: {
      title: [{ shape: 'wave', x: 0, y: 0.72, w: 1, h: 0.28, fill: 'theme:accent1', opacity: 0.9 }],
      content: [{ shape: 'rect', x: 0.94, y: 0, w: 0.06, h: 1, fill: 'theme:accent1', opacity: 0.15 }],
    },
  },
  {
    id: 'plum', name: 'Plum',
    background: { type: 'solid', color: '#2e1065' },
    colors: { text: '#e9d5ff', title: '#fde68a', accent1: '#f59e0b', accent2: '#ec4899', accent3: '#22d3ee', accent4: '#a3e635', accent5: '#f97316', accent6: '#818cf8' },
    fonts: { title: 'Playfair Display', body: 'Inter' }, titleBold: true, dark: true,
    decor: {
      title: [{ shape: 'ellipse', x: -0.12, y: 0.55, w: 0.4, h: 0.71, fill: 'theme:accent2', opacity: 0.35 }],
      content: [],
    },
  },
];

export const DEFAULT_THEME = 'simple-light';

export const themeById = (id) => THEMES.find((t) => t.id === id) || THEMES[0];

export const THEME_COLOR_KEYS = ['title', 'text', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];

/** A colour value: "#rrggbb", "rgba(…)", "theme:<key>", "none" or null (the fallback). */
export function resolveColor(value, theme, fallback = 'transparent') {
  if (value == null || value === '') return fallback;
  if (value === 'none') return 'transparent';
  if (typeof value === 'string' && value.startsWith('theme:')) {
    const key = value.slice(6);
    if (key === 'bg') return backgroundBaseColor(theme.background);
    return theme.colors[key] || fallback;
  }
  return value;
}

/** The colour a background mostly is: a gradient's first, a pattern's background, else its colour. */
export const backgroundBaseColor = (bg) => (bg?.type === 'gradient' ? bg.from : bg?.type === 'pattern' ? bg.pattern?.bg : bg?.color) || '#ffffff';

/**
 * A gradient background's gradient, as shapes have them (see model/deck.js):
 * its own (`gradient`: more stops, or from the centre out) or the one its
 * `from`, `to` and `angle` make.
 */
export const backgroundGradient = (bg) => (bg.gradient?.stops?.length > 1 ? bg.gradient
  : { type: 'linear', angle: bg.angle ?? 180, stops: [{ pos: 0, color: bg.from }, { pos: 1, color: bg.to }] });

/** A gradient background with another first or last colour or angle (its own gradient changed to match). */
export function editBackgroundGradient(bg, { from, to, angle }) {
  const out = { ...bg };
  const g = bg.gradient?.stops?.length > 1 ? { ...bg.gradient, stops: bg.gradient.stops.map((s) => ({ ...s })) } : null;
  if (from != null) out.from = from;
  if (to != null) out.to = to;
  if (angle != null) out.angle = angle;
  if (g) {
    if (from != null) g.stops[0].color = from;
    if (to != null) g.stops[g.stops.length - 1].color = to;
    if (angle != null) {
      g.type = 'linear';
      g.angle = angle;
      delete g.cx;
      delete g.cy;
    }
    out.gradient = g;
  }
  return out;
}

/** "#rrggbb" seen at `opacity` (0–1), as CSS. */
export const withOpacity = (color, opacity = 1) => (opacity < 1 && /^#[0-9a-f]{6}$/i.test(color)
  ? `rgba(${[1, 3, 5].map((i) => parseInt(color.slice(i, i + 2), 16)).join(', ')}, ${Math.round(opacity * 1000) / 1000})` : color);

/** A gradient (see model/deck.js) as a CSS image, radial ones out to the farthest corner. */
export function gradientCss(g, theme) {
  const pct = (v) => Math.round(Math.max(0, Math.min(1, v)) * 1000) / 10;
  const stops = (g.stops || []).map((s) => `${withOpacity(resolveColor(s.color, theme, '#000000'), s.opacity ?? 1)} ${pct(s.pos)}%`).join(', ');
  return g.type === 'radial'
    ? `radial-gradient(farthest-corner at ${pct(g.cx ?? 0.5)}% ${pct(g.cy ?? 0.5)}%, ${stops})`
    : `linear-gradient(${g.angle ?? 180}deg, ${stops})`;
}

/**
 * Where a tiled picture's tiles line up, as fractions of the box: OOXML's
 * alignments (tl, t, tr, l, ctr, r, bl, b, br; top left when unknown).
 */
export function tileAlign(align) {
  const a = ['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br'].includes(align) ? align : 'tl';
  const x = a === 't' || a === 'ctr' || a === 'b' ? 0.5 : a.endsWith('r') ? 1 : 0;
  const y = a.startsWith('t') ? 0 : a.startsWith('b') ? 1 : 0.5;
  return { x, y };
}

/**
 * CSS `background` for a background spec ({type: solid|gradient|image|pattern}).
 * A gradient may have a `gradient` of its own (see backgroundGradient). A
 * picture covers the slide, or is tiled when it has `tile` { w, h (a tile's
 * size), x, y (its offset from where `align` puts it, see tileAlign), flip?
 * (none|x|y|xy) } in points. A pattern is `pattern` { preset, fg, bg } (see
 * model/patterns.js).
 */
export function backgroundCss(bg, theme, mediaUrl = (x) => x) {
  if (!bg) bg = theme.background;
  if (bg.type === 'gradient') {
    if (bg.gradient?.stops?.length > 1) return gradientCss(bg.gradient, theme);
    return `linear-gradient(${bg.angle ?? 180}deg, ${resolveColor(bg.from, theme)}, ${resolveColor(bg.to, theme)})`;
  }
  if (bg.type === 'image' && bg.media) {
    const under = resolveColor(bg.color || '#000', theme);
    const t = bg.tile;
    if (t?.w > 0 && t?.h > 0) {
      const a = tileAlign(t.align);
      const round = (v) => Math.round((v || 0) * 100) / 100;
      const at = (f, v) => `calc(${f * 100}% + ${round(v)}px)`;
      return `${under} url("${mediaUrl(bg.media)}") ${at(a.x, t.x)} ${at(a.y, t.y)} / ${round(t.w)}px ${round(t.h)}px repeat`;
    }
    return `${under} url("${mediaUrl(bg.media)}") center / cover no-repeat`;
  }
  if (bg.type === 'pattern' && bg.pattern) {
    const back = resolveColor(bg.pattern.bg, theme, '#ffffff');
    const svg = patternSvg(bg.pattern.preset, resolveColor(bg.pattern.fg, theme, '#000000'), back);
    return `${back} url("data:image/svg+xml,${encodeURIComponent(svg)}") 0 0 / ${PATTERN_SIZE}px ${PATTERN_SIZE}px repeat`;
  }
  return resolveColor(bg.color, theme, '#ffffff');
}

/** Rough luminance test, used to pick readable text over a fill. */
export function isDarkColor(hex) {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex || '');
  if (!m) return false;
  const n = parseInt(m[1], 16);
  const [r, g, b] = [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  return 0.2126 * r + 0.7152 * g + 0.0722 * b < 140;
}
