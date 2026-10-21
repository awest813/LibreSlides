import { DEFAULT_THEME, themeById } from './themes.js';

/**
 * The presentation model. Coordinates are in points on a slide 960 × 540 pt
 * (16:9, PowerPoint's 13.33 × 7.5 in widescreen), 720 × 540 pt (4:3), another
 * of SIZES, or a size of its own ("WxH" in points: see customSize); font
 * sizes are in points too, so a slide renders at 1px = 1pt and is scaled.
 *
 * deck    { id, title, themeId, size, slides[], media{id: dataUrl}, createdAt, updatedAt }
 * slide   { id, layout, background|null, elements[], notes, transition{type, duration}, hidden }
 * element { id, type: text|shape|image|line|table|chart|media, x, y, w, h, rot, opacity, … }
 * media   { id, type: 'media', kind: 'video'|'audio', media, poster?, x, y, w, h, rot, autoplay?, loop? }
 *
 * A chart keeps its data in the element, so it can be edited and exported as
 * a real PowerPoint chart:
 *   chart { kind: column|bar|line|area|pie|doughnut|scatter,
 *           grouping: clustered|stacked|percent (columns, bars and areas),
 *           title, categories[], series[{ name, values[number|null], color?, pointColors? }],
 *           legend: right|bottom|top|left|none, dataLabels, markers (lines) }
 * Series without a colour take the theme's accents in turn; pie and doughnut
 * slices use the first series, coloured by its pointColors. A scatter chart's
 * categories are its X values. `style` { size, color, font } sets its text,
 * and `fill` and `stroke` the chart area's background and border.
 *
 * Pictures may have colour adjustments, `adjust: { grayscale?, brightness?,
 * contrast? }` (-1 to 1), and be cropped to a shape, `clip: { shape, paths?,
 * preset? }` as shape elements have (an oval, a PowerPoint preset…).
 *
 * Elements can have an entrance (`anim`), an emphasis and an exit animation
 * (see ANIM_KINDS).
 *
 * Elements sharing a `groupId` are a group: selected, moved and saved
 * together, next to each other in the stacking order.
 *
 * Shapes and text boxes may also have a gradient fill, `gradient: { type:
 * 'linear'|'radial', angle (CSS degrees: 90 runs left to right) | cx, cy
 * (centre, fractions of the box), stops: [{ pos 0–1, color, opacity? }] }`,
 * with `fill` holding its first colour, or a picture fill, `fillImage:
 * { media, crop? }`. Tables are cells[row][col] of { text, fill?, valign?,
 * rowSpan?, colSpan?, merged? } (see tables.js).
 *
 * Pictures, videos and sounds live in deck.media (as data URLs) and are
 * referenced by id, so undo snapshots of the slides stay small however many
 * a deck has. A video or sound element points at its file (`media`) and at
 * the picture shown until it plays (`poster`, optional). It plays when the
 * slide appears if `autoplay`, otherwise when its control is clicked.
 *
 * Text may hold fields PowerPoint fills in as the slideshow runs: the
 * slide's number or today's date, `<span data-field="slidenum|datetime…">`
 * (see FIELD and fieldText). The text inside is only what was last shown.
 *
 * Runs of text may have WordArt-like effects as inline CSS: an outline
 * (-webkit-text-stroke), a shadow or glow (text-shadow) and a gradient
 * (a linear-gradient background clipped to the text, with `color` holding
 * its first colour).
 *
 * A slide may carry review comments, as PowerPoint's do. They're plain text
 * (never HTML), shown only in the editor, and kept through undo, copies and
 * .pptx files (see io/pptx-comments.js):
 *   slide.comments = [{ id, author, initials?, date (ISO string), text,
 *     x?, y? (where its marker sits, in slide points; none: on the whole slide),
 *     resolved?, replies?: [{ id, author, initials?, date, text }] }]
 */
export const SIZES = {
  '16:9': { w: 960, h: 540, label: 'Widescreen (16:9)' },
  '16:10': { w: 864, h: 540, label: 'Widescreen (16:10)' },
  '4:3': { w: 720, h: 540, label: 'Standard (4:3)' },
  a4: { w: 780, h: 540, label: 'A4 paper' },
  'a4-portrait': { w: 540, h: 780, label: 'A4 paper (portrait)' },
  'letter-portrait': { w: 540, h: 720, label: 'Letter paper (portrait)' },
};

/**
 * A size of its own, "WxH" in points (from a PowerPoint file, or chosen),
 * within PowerPoint's limits of 1 to 56 inches.
 */
export function customSize(size) {
  const m = /^(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)$/.exec(String(size || ''));
  if (!m) return null;
  const [w, h] = [+m[1], +m[2]];
  if (!(w >= 72 && w <= 4032 && h >= 72 && h <= 4032)) return null;
  const inch = (v) => Math.round((v / 72) * 100) / 100;
  return { w, h, label: `Custom (${inch(w)} × ${inch(h)} in)` };
}

export const slideSize = (deck) => SIZES[deck?.size] || customSize(deck?.size) || SIZES['16:9'];

export const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-3);

export const LAYOUTS = [
  { id: 'title', name: 'Title slide' },
  { id: 'titleContent', name: 'Title and body' },
  { id: 'section', name: 'Section header' },
  { id: 'twoContent', name: 'Title and two columns' },
  { id: 'titleOnly', name: 'Title only' },
  { id: 'bigNumber', name: 'Big number' },
  { id: 'quote', name: 'Quote' },
  { id: 'blank', name: 'Blank' },
];

export const TRANSITIONS = [
  { id: 'none', name: 'None' },
  { id: 'fade', name: 'Fade' },
  { id: 'slide', name: 'Slide from right' },
  { id: 'push', name: 'Push' },
  { id: 'cover', name: 'Cover from bottom' },
  { id: 'zoom', name: 'Zoom' },
  { id: 'flip', name: 'Flip' },
];

export const ANIMATIONS = [
  { id: 'none', name: 'None' },
  { id: 'appear', name: 'Appear' },
  { id: 'fade', name: 'Fade in' },
  { id: 'fly-up', name: 'Fly in from bottom' },
  { id: 'fly-down', name: 'Fly in from top' },
  { id: 'fly-left', name: 'Fly in from left' },
  { id: 'fly-right', name: 'Fly in from right' },
  { id: 'zoom', name: 'Zoom in' },
  { id: 'wipe', name: 'Wipe from left' },
  { id: 'spin', name: 'Spin in' },
];

// Effects that draw attention to something already there, and that take it away.
export const EMPHASES = [
  { id: 'none', name: 'None' },
  { id: 'pulse', name: 'Pulse' },
  { id: 'spin', name: 'Spin' },
  { id: 'teeter', name: 'Teeter' },
  { id: 'flash', name: 'Flash' },
];

export const EXITS = [
  { id: 'none', name: 'None' },
  { id: 'appear', name: 'Disappear' },
  { id: 'fade', name: 'Fade out' },
  { id: 'fly-up', name: 'Fly out to bottom' },
  { id: 'fly-down', name: 'Fly out to top' },
  { id: 'fly-left', name: 'Fly out to left' },
  { id: 'fly-right', name: 'Fly out to right' },
  { id: 'zoom', name: 'Zoom out' },
  { id: 'wipe', name: 'Wipe to left' },
  { id: 'spin', name: 'Spin out' },
];

// Motion paths: where an object travels, as points relative to where it starts,
// in fractions of the slide's width and height (as PowerPoint keeps them).
const circle = Array.from({ length: 25 }, (_, i) => {
  const a = (i / 24) * 2 * Math.PI;
  return [Math.round(0.1 * Math.sin(a) * 1000) / 1000, Math.round(0.15 * (1 - Math.cos(a)) * 1000) / 1000];
});
export const MOTIONS = [
  { id: 'none', name: 'None' },
  { id: 'right', name: 'Move right', path: [[0, 0], [0.25, 0]] },
  { id: 'left', name: 'Move left', path: [[0, 0], [-0.25, 0]] },
  { id: 'up', name: 'Move up', path: [[0, 0], [0, -0.25]] },
  { id: 'down', name: 'Move down', path: [[0, 0], [0, 0.25]] },
  { id: 'circle', name: 'Circle', path: circle },
  { id: 'custom', name: 'Custom path' },
];

/** A motion animation's path (see MOTIONS). */
export const motionPath = (m) => (m?.effect === 'custom' ? m.path : MOTIONS.find((x) => x.id === m?.effect)?.path) || [[0, 0]];

/**
 * An element's animations: `anim` (its entrance), `emphasis`, `exit` and
 * `motion` (a path, with `path` for a custom one), each { effect, duration,
 * start: click|with|after, order }, sharing one playing order on the slide.
 */
/** PowerPoint's field types LibreSlides fills in: the slide number, and the date in one of its 13 formats. */
export const FIELD = /^(slidenum|datetime(1[0-3]|[1-9])?)$/;

// How each date field is written, as Intl.DateTimeFormat options (PowerPoint's datetime1…13).
const DATE_FORMATS = [
  { dateStyle: 'short' }, { dateStyle: 'full' }, { day: 'numeric', month: 'long', year: 'numeric' }, { dateStyle: 'long' },
  { day: '2-digit', month: 'short', year: '2-digit' }, { month: 'long', year: '2-digit' }, { month: 'short', year: '2-digit' },
  { dateStyle: 'short', timeStyle: 'short' }, { dateStyle: 'short', timeStyle: 'medium' },
  { hour: '2-digit', minute: '2-digit', hourCycle: 'h23' }, { hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' },
  { hour: 'numeric', minute: '2-digit', hour12: true }, { hour: 'numeric', minute: '2-digit', second: '2-digit', hour12: true },
];

/** What a field shows on slide number `n` (from 1). */
export function fieldText(type, n, date = new Date()) {
  if (type === 'slidenum') return String(n);
  const i = Math.max(1, +type.slice(8) || 1) - 1;
  return new Intl.DateTimeFormat(undefined, DATE_FORMATS[i]).format(date);
}

export const ANIM_KINDS = ['anim', 'emphasis', 'exit', 'motion'];

/** Every animation on a slide, in playing order: { el, kind, anim }. */
export function slideAnimations(slide) {
  return slide.elements.flatMap((el) => ANIM_KINDS.filter((kind) => el[kind]?.effect && el[kind].effect !== 'none').map((kind) => ({ el, kind, anim: el[kind] })))
    .sort((a, b) => (a.anim.order || 0) - (b.anim.order || 0));
}

/** Placeholder hint text shown on empty placeholders while editing. */
export const PLACEHOLDER_HINTS = {
  title: 'Click to add title',
  subtitle: 'Click to add subtitle',
  body: 'Click to add text',
};

const text = (props) => ({ id: uid(), type: 'text', rot: 0, text: '', fill: 'none', stroke: null, ...props, style: { ...props.style } });

/** The placeholder text boxes of a layout, for a slide W × H. */
export function layoutElements(layout, W, H) {
  const X = (f) => Math.round(f * W);
  const Y = (f) => Math.round(f * H);
  const title = (geom, style = {}) => text({ ph: 'title', ...geom, style: { size: 34, valign: 'bottom', ...style } });
  const body = (geom, style = {}) => text({ ph: 'body', bullets: true, ...geom, style: { size: 20, valign: 'top', ...style } });
  const subtitle = (geom, style = {}) => text({ ph: 'subtitle', ...geom, style: { size: 22, valign: 'top', ...style } });
  const topTitle = () => title({ x: X(0.06), y: Y(0.06), w: X(0.88), h: Y(0.16) });
  switch (layout) {
    case 'title':
      return [
        title({ x: X(0.1), y: Y(0.27), w: X(0.8), h: Y(0.33) }, { size: 50 }),
        subtitle({ x: X(0.1), y: Y(0.65), w: X(0.8), h: Y(0.14) }),
      ];
    case 'titleContent':
      return [topTitle(), body({ x: X(0.06), y: Y(0.26), w: X(0.88), h: Y(0.64) })];
    case 'section':
      return [
        title({ x: X(0.1), y: Y(0.3), w: X(0.8), h: Y(0.3) }, { size: 44 }),
        subtitle({ x: X(0.1), y: Y(0.63), w: X(0.8), h: Y(0.14) }, { size: 20 }),
      ];
    case 'twoContent':
      return [
        topTitle(),
        body({ x: X(0.06), y: Y(0.26), w: X(0.43), h: Y(0.64) }),
        body({ x: X(0.51), y: Y(0.26), w: X(0.43), h: Y(0.64) }),
      ];
    case 'titleOnly':
      return [topTitle()];
    case 'bigNumber':
      return [
        title({ x: X(0.06), y: Y(0.12), w: X(0.88), h: Y(0.46) }, { size: 120, color: 'theme:accent1' }),
        body({ x: X(0.06), y: Y(0.62), w: X(0.88), h: Y(0.24), bullets: false }, { size: 24 }),
      ];
    case 'quote':
      return [
        title({ x: X(0.12), y: Y(0.2), w: X(0.76), h: Y(0.44) }, { size: 36, italic: true, align: 'center', valign: 'middle' }),
        subtitle({ x: X(0.12), y: Y(0.68), w: X(0.76), h: Y(0.1) }, { size: 18, align: 'center' }),
      ];
    default:
      return [];
  }
}

/** Title-style layouts get the theme's "title" decoration. */
export const isTitleLayout = (layout) => layout === 'title' || layout === 'section' || layout === 'quote';

export function newSlide(layout = 'titleContent', deck = null) {
  const { w, h } = slideSize(deck);
  return { id: uid(), layout, background: null, elements: layoutElements(layout, w, h), notes: '', transition: { type: 'none', duration: 0.6 }, hidden: false };
}

export function newDeck({ title = 'Untitled presentation', themeId = DEFAULT_THEME, size = '16:9', slides = null } = {}) {
  const deck = { id: uid(), title, themeId, size, slides: [], media: {}, createdAt: Date.now(), updatedAt: Date.now() };
  deck.slides = slides || [newSlide('title', deck)];
  return deck;
}

// ------------------------------------------------------------------ elements

export function newTextBox(x, y, w = 300, h = 60) {
  return text({ x, y, w, h, style: { size: 18 } });
}

export function newShape(kind, x, y, w = 160, h = 120) {
  return {
    id: uid(), type: 'shape', shape: kind, x, y, w, h, rot: 0, text: '',
    fill: 'theme:accent1', stroke: { color: 'none', width: 2, dash: 'solid' },
    style: { size: 18, color: '#ffffff', align: 'center', valign: 'middle' },
  };
}

export function newLine(x1, y1, x2, y2, { arrowEnd = false } = {}) {
  return { id: uid(), type: 'line', x1, y1, x2, y2, rot: 0, stroke: { color: 'theme:text', width: 2, dash: 'solid' }, arrowStart: false, arrowEnd };
}

export function newImage(media, x, y, w, h) {
  return { id: uid(), type: 'image', media, x, y, w, h, rot: 0, stroke: { color: 'none', width: 2, dash: 'solid' }, radius: 0 };
}

/** A video or sound: `media` is its file, `poster` the picture shown until it plays. */
export function newMedia(kind, media, poster, x, y, w, h) {
  const el = { id: uid(), type: 'media', kind, media, x, y, w, h, rot: 0, autoplay: false, loop: false };
  if (poster) el.poster = poster;
  return el;
}

export function newTable(rows, cols, x, y, w, h) {
  return {
    id: uid(), type: 'table', x, y, w, h, rot: 0, header: true,
    cells: Array.from({ length: rows }, () => Array.from({ length: cols }, () => ({ text: '' }))),
    colWidths: Array.from({ length: cols }, () => 1 / cols),
    fill: 'theme:accent1', stroke: { color: '#94a3b8', width: 1, dash: 'solid' },
    style: { size: 14 },
  };
}

export const CHART_KINDS = [
  { id: 'column', name: 'Column' },
  { id: 'bar', name: 'Bar' },
  { id: 'line', name: 'Line' },
  { id: 'area', name: 'Area' },
  { id: 'pie', name: 'Pie' },
  { id: 'doughnut', name: 'Doughnut' },
  { id: 'scatter', name: 'Scatter (XY)' },
];

/** A column chart with PowerPoint's sample data, ready to be edited. */
export function newChart(x, y, w = 480, h = 300) {
  return {
    id: uid(), type: 'chart', x, y, w, h, rot: 0,
    chart: {
      kind: 'column', grouping: 'clustered', title: '',
      categories: ['Category 1', 'Category 2', 'Category 3', 'Category 4'],
      series: [
        { name: 'Series 1', values: [4.3, 2.5, 3.5, 4.5] },
        { name: 'Series 2', values: [2.4, 4.4, 1.8, 2.8] },
        { name: 'Series 3', values: [2, 2, 3, 5] },
      ],
      legend: 'bottom', dataLabels: false,
    },
    fill: 'none', stroke: { color: 'none', width: 1, dash: 'solid' },
    style: { size: 12 },
  };
}

/** A line's bounding box (lines store their two end points). */
export function lineBox(el) {
  const x = Math.min(el.x1, el.x2);
  const y = Math.min(el.y1, el.y2);
  return { x, y, w: Math.abs(el.x2 - el.x1), h: Math.abs(el.y2 - el.y1) };
}

/** Bounding box of any element (unrotated). */
export const boxOf = (el) => (el.type === 'line' ? lineBox(el) : { x: el.x, y: el.y, w: el.w, h: el.h });

/** Move an element by dx, dy (in place). */
export function moveElement(el, dx, dy) {
  if (el.type === 'line') {
    el.x1 += dx; el.x2 += dx; el.y1 += dy; el.y2 += dy;
  } else {
    el.x += dx; el.y += dy;
  }
}

export const canHoldText = (el) => el.type === 'text' || el.type === 'shape';

const deepCopy = (v) => JSON.parse(JSON.stringify(v));

/** The elements grouped with `el` on its slide (itself alone when it isn't in a group of two or more). */
export function groupOf(slide, el) {
  if (!el?.groupId) return [el];
  const members = slide.elements.filter((e) => e.groupId === el.groupId);
  return members.length > 1 ? members : [el];
}

/** Give copied elements groups of their own, so they don't join the groups they were copied from. */
export function regroup(els) {
  const ids = new Map();
  for (const el of els) {
    if (!el.groupId) continue;
    if (!ids.has(el.groupId)) ids.set(el.groupId, uid());
    el.groupId = ids.get(el.groupId);
  }
  return els;
}

/** "Ada Lovelace" → "AL": the initials PowerPoint shows for a comment's author. */
export const initialsOf = (name) => String(name).split(/\s+/).map((w) => w.match(/[\p{L}\p{N}]/u)?.[0]).filter(Boolean).slice(0, 2).join('').toUpperCase();

export function cloneElement(el) {
  return { ...deepCopy(el), id: uid() };
}

export function cloneSlide(slide) {
  const copy = deepCopy(slide);
  copy.id = uid();
  copy.elements = copy.elements.map((e) => ({ ...e, id: uid() }));
  if (Array.isArray(copy.comments)) {
    copy.comments = copy.comments.map((c) => ({ ...c, id: uid(), ...(c.replies ? { replies: c.replies.map((r) => ({ ...r, id: uid() })) } : {}) }));
  }
  return copy;
}

// ------------------------------------------------------------------ media

/** Store a picture (data URL) in the deck, reusing an identical one. Returns its id. */
export function addMedia(deck, dataUrl) {
  deck.media ||= {};
  for (const [id, url] of Object.entries(deck.media)) if (url === dataUrl) return id;
  const id = `m${uid()}`;
  deck.media[id] = dataUrl;
  return id;
}

/** Where an element (or slide background) refers to deck.media: [holder, key] pairs. */
export function mediaSlots(el) {
  const slots = [];
  if (el?.media) slots.push([el, 'media']);
  if (el?.poster) slots.push([el, 'poster']);
  if (el?.fillImage?.media) slots.push([el.fillImage, 'media']);
  return slots;
}

/** The deck.media ids an element refers to (a picture, a video and its poster, a fill picture). */
export const mediaIdsOf = (el) => mediaSlots(el).map(([o, key]) => o[key]);

/** Ids of the pictures, videos and sounds the deck still uses. */
export function usedMedia(deck) {
  const used = new Set();
  for (const s of deck.slides) {
    for (const holder of [s.background, ...s.elements]) for (const id of mediaIdsOf(holder)) used.add(id);
  }
  return used;
}

/** Drop pictures no slide uses any more (call when no undo history refers to them). */
export function pruneMedia(deck) {
  const used = usedMedia(deck);
  for (const id of Object.keys(deck.media || {})) if (!used.has(id)) delete deck.media[id];
  return deck;
}

// ------------------------------------------------------------------ text helpers

/** Plain text of an HTML fragment (for previews, search and notes). */
export function htmlToText(html = '') {
  if (!html) return '';
  const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
  doc.body.querySelectorAll('br').forEach((br) => br.replaceWith('\n'));
  doc.body.querySelectorAll('p, li, div, h1, h2, h3').forEach((el) => el.append('\n'));
  return (doc.body.textContent || '').replace(/\n{2,}/g, '\n').trim();
}

const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

/** HTML for plain paragraphs, or a bullet list when `bullets`. */
export function textToHtml(value, { bullets = false } = {}) {
  const lines = (Array.isArray(value) ? value : String(value).split('\n')).map((l) => String(l));
  if (bullets) return `<ul>${lines.map((l) => `<li>${esc(l) || '<br>'}</li>`).join('')}</ul>`;
  return lines.map((l) => `<p>${esc(l) || '<br>'}</p>`).join('');
}

/** Fill a slide's placeholders: { title, subtitle, body, body2 } (strings or arrays of bullets). */
export function fillPlaceholders(slide, content) {
  const bodies = slide.elements.filter((e) => e.ph === 'body');
  for (const el of slide.elements) {
    if (el.ph === 'title' && content.title != null) el.text = textToHtml(content.title);
    if (el.ph === 'subtitle' && content.subtitle != null) el.text = textToHtml(content.subtitle);
  }
  const bodyContent = [content.body, content.body2];
  bodies.forEach((el, i) => {
    const v = bodyContent[i];
    if (v != null) el.text = textToHtml(v, { bullets: el.bullets && Array.isArray(v) });
  });
  return slide;
}

/** Plain text of a slide, title first (for the deck list and search). */
export function slideText(slide) {
  const els = [...slide.elements].sort((a, b) => (a.ph === 'title' ? -1 : b.ph === 'title' ? 1 : 0));
  const parts = [];
  for (const el of els) {
    if (el.text) parts.push(htmlToText(el.text));
    if (el.cells) for (const row of el.cells) for (const c of row) if (c.text) parts.push(htmlToText(c.text));
    if (el.chart?.title) parts.push(el.chart.title);
  }
  return parts.join(' ').replace(/\s+/g, ' ').trim();
}

/** Title of a slide, for the outline and the presenter view. */
export function slideTitle(slide) {
  const t = slide.elements.find((e) => e.ph === 'title' && e.text);
  return t ? htmlToText(t.text).split('\n')[0] : '';
}

/** Make a loaded deck safe to use: fill in fields added by later versions. */
export function normalizeDeck(deck) {
  deck.media ||= {};
  deck.size = SIZES[deck.size] || customSize(deck.size) ? deck.size : '16:9';
  deck.themeId = themeById(deck.themeId).id;
  deck.title ||= 'Untitled presentation';
  if (!Array.isArray(deck.slides) || !deck.slides.length) deck.slides = [newSlide('title', deck)];
  for (const s of deck.slides) {
    s.id ||= uid();
    s.elements ||= [];
    s.notes ||= '';
    s.transition ||= { type: 'none', duration: 0.6 };
    for (const el of s.elements) {
      el.id ||= uid();
      el.rot ||= 0;
      if (canHoldText(el)) el.style ||= {};
    }
  }
  return deck;
}
