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

export const backgroundBaseColor = (bg) => (bg?.type === 'gradient' ? bg.from : bg?.color || '#ffffff');

/** CSS `background` for a background spec ({type: solid|gradient|image}). */
export function backgroundCss(bg, theme, mediaUrl = (x) => x) {
  if (!bg) bg = theme.background;
  if (bg.type === 'gradient') return `linear-gradient(${bg.angle ?? 180}deg, ${resolveColor(bg.from, theme)}, ${resolveColor(bg.to, theme)})`;
  if (bg.type === 'image' && bg.media) return `${resolveColor(bg.color || '#000', theme)} url("${mediaUrl(bg.media)}") center / cover no-repeat`;
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
