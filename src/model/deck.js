import { DEFAULT_THEME, themeById } from './themes.js';

/**
 * The presentation model. Coordinates are in points on a slide 960 × 540 pt
 * (16:9, PowerPoint's 13.33 × 7.5 in widescreen), 720 × 540 pt (4:3), another
 * of SIZES, or a size of its own ("WxH" in points: see customSize); font
 * sizes are in points too, so a slide renders at 1px = 1pt and is scaled.
 *
 * deck    { id, title, themeId, size, slides[], media{id: dataUrl}, createdAt, updatedAt,
 *           rev (bumped by each save, see storage/db.js), fonts? (embedded, see fonts.js FONT_FACES),
 *           file? { handle, name, format, lastModified, dirty } (the file on this device it saves to) }
 * slide   { id, layout (LAYOUTS), background|null ({ type: solid|gradient|image, … }, see
 *           themes.js backgroundCss; null: the theme's), elements[] (back to front), notes (plain
 *           text), transition { type (TRANSITIONS), duration (s) }, hidden, hideDecor?, section?, comments? }
 * element { id, type: text|shape|image|line|table|chart|media|diagram, x, y, w, h (the unrotated
 *           box), rot (degrees clockwise about its centre), opacity? (0–1), flipH?, flipV?, shadow?,
 *           groupId?, locked?, lockRatio?, link? ("https://…", "mailto:…" or "#slide:<id>|first|last|next|prev"),
 *           anim?, emphasis?, exit?, motion? (see ANIM_KINDS), … }
 *   Lines have no box: { x1, y1, x2, y2, stroke, arrowStart, arrowEnd (false, true or triangle|arrow|
 *   stealth|diamond|oval) } (see lineBox). Strokes are { color, width (pt), dash: solid|dash|dot|
 *   dashDot|longDash }. Colours are "#rrggbb", "rgba(…)", "theme:<key>" or "none" (see themes.js).
 *   text, shape { text (HTML), ph? (title|subtitle|body: a layout's placeholder), bullets?, fill, stroke,
 *           style { font, size (pt), color, bold, italic, underline, strike, align, valign, lineHeight,
 *             spacing (between paragraphs, em), inset? { l, t, r, b } (pt), wrap? (false: no wrapping),
 *             vertical? (vert|vert270|eaVert), columns?, columnGap? (pt), warp? (see below) } }
 *   shape  { shape (SHAPES id, or 'custom' with paths, see shapes.js), paths?, preset? { name, adj? }
 *           (the PowerPoint preset it was opened as), textRect? { l, t, r, b } (where its text goes,
 *           fractions of the box), arrowStart?, arrowEnd? (open freeforms: connectors) }
 *   image  { media, alt?, fit? (fill|cover|contain), crop? { l, t, r, b } (fractions cut off), radius?,
 *           stroke, adjust?, clip? (see below) }
 *   media  { kind: 'video'|'audio', media, poster?, autoplay?, loop?, stroke? }
 *
 * A diagram (PowerPoint's SmartArt) keeps its text as an outline and is
 * drawn from it by diagram-layout.js, so it re-lays itself out as it's edited:
 *   diagram { layout: one of DIAGRAM_LAYOUTS (list, vlist, process, chevron, cycle,
 *               hierarchy, pyramid, venn, matrix, radial),
 *             nodes: [{ id, text (plain text), level (0–8, as in PowerPoint's text pane) }],
 *             colors: accent1…accent6|colorful, style: flat|outline|raised (DIAGRAM_COLORS, DIAGRAM_STYLES),
 *             font?, palette? (the six accent colours of the file it came from, used for its accents),
 *             drawn?: { w, h, elements[] } (PowerPoint's own drawing of it: shapes, lines and
 *               pictures in the box's coordinates when it was w × h; shown, stretched to the box,
 *               until the diagram is edited),
 *             edited? (changed since it was opened from a file),
 *             source?: { data, layout, style, colors, drawing? (the SmartArt parts, their XML in
 *                 deck.media as "data:application/xml" URLs, see xmlDataUrl),
 *               look: { layout, colors, style } (as opened),
 *               media?: { data?|drawing?: [{ id (relationship), media }] } (pictures the parts refer to) } }
 * A diagram from a PowerPoint file goes back as SmartArt (its parts as they were, or its
 * data with the new outline when edited); new ones, or ones given another layout, colours
 * or style, go as a group of plain shapes (see io/pptx-diagram.js).
 *
 * A chart keeps its data in the element, so it can be edited and exported as
 * a real PowerPoint chart:
 *   chart { kind: column|bar|line|area|pie|doughnut|scatter|bubble|radar,
 *           grouping: clustered|stacked|percent (columns, bars, areas and lines),
 *           title, categories[], series[{ name, values[number|null], color?, pointColors?, sizes? (bubbles),
 *             kind?: column|line|area (combination charts), secondary? (right-hand axis) }],
 *           legend: right|bottom|top|left|none, dataLabels, markers (lines, radar; scatter unless false),
 *           filled? (radar), lineWidth? (pt, lines and radar; default 2.25),
 *           numberFormat?, labelFormat? (Excel codes: "#,##0.00", "0%", "$#,##0"…),
 *           labelContent?: [category, value, percent] (default value), labelPosition?
 *             (outEnd|inEnd|ctr|inBase|bestFit|t|b|l|r), lines?, smooth? (scatter, lines),
 *           gapWidth? (0–500, default 150), overlap? (-100–100), holeSize? (10–90, default 50),
 *           valueAxis?, secondaryAxis? { min?, max?, unit?, title?, gridlines?, numberFormat? (secondary) },
 *           categoryAxis? { title } }
 * Every optional field may be missing (older decks): the defaults are PowerPoint's.
 * Series without a colour take the theme's accents in turn; pie slices use the
 * first series and doughnuts a ring per series (the first innermost), coloured
 * by the first series' pointColors. A scatter or bubble chart's categories are
 * its X values (a bubble's size is in its series' sizes). `style` { size, color, font } sets its text, and `fill` and
 * `stroke` the chart area's background and border.
 *
 * Shapes, text boxes and pictures may have `softEdge` (pt): edges that fade
 * out over that distance (PowerPoint's soft edges effect).
 *
 * Pictures may have colour adjustments, `adjust: { grayscale?, brightness?,
 * contrast? }` (-1 to 1), and be cropped to a shape, `clip: { shape, paths?,
 * preset? }` as shape elements have (an oval, a PowerPoint preset…).
 *
 * Elements can have an entrance (`anim`), an emphasis and an exit animation,
 * and a motion path (see ANIM_KINDS and MOTIONS).
 *
 * Elements sharing a `groupId` are a group: selected, moved and saved
 * together, next to each other in the stacking order (see gatherGroups).
 *
 * Shapes and text boxes may also have a gradient fill, `gradient: { type:
 * 'linear'|'radial', angle (CSS degrees: 90 runs left to right) | cx, cy
 * (centre, fractions of the box), stops: [{ pos 0–1, color, opacity? }] }`,
 * with `fill` holding its first colour, or a picture fill, `fillImage:
 * { media, crop? }`. Tables are { cells, colWidths, rowHeights? (fractions of
 * the box), header? (first row in the table's `fill`), banded?, fill, stroke,
 * style }, cells[row][col] being { text, fill?, valign?, rowSpan?, colSpan?,
 * merged?, borders?, margin? } (see tables.js). A cell's
 * `borders: { t?, r?, b?, l? }` are strokes { color, width, dash } (colour
 * 'none': no border) for its sides, the table's `stroke` standing in for a
 * side it doesn't set; `margin: { t, r, b, l }` is the space around its text
 * in points (CELL_MARGIN when there's none).
 *
 * Pictures, videos and sounds live in deck.media (as data URLs) and are
 * referenced by id, so undo snapshots of the slides stay small however many
 * a deck has. So do the SmartArt parts that diagrams from PowerPoint keep. A
 * video or sound element points at its file (`media`) and at the picture
 * shown until it plays (`poster`, optional). It plays when the slide appears
 * if `autoplay`, otherwise when its control is clicked.
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
 * A text box (or shape's text) may be bent into a WordArt shape, as
 * PowerPoint's Text Effects › Transform does: `style.warp = { preset
 * (textArchUp, textWave1… the OOXML names), adjust?: { adj: number, … }
 * (in the preset's own units), stretch? (along "follow path" warps too) }`.
 * See model/warps.js; it's drawn by render/warp.js.
 *
 * A slide may carry review comments, as PowerPoint's do. They're plain text
 * (never HTML), shown only in the editor, and kept through undo, copies and
 * .pptx files (see io/pptx-comments.js):
 *   slide.comments = [{ id, author, initials?, date (ISO string), text,
 *     x?, y? (where its marker sits, in slide points; none: on the whole slide),
 *     resolved?, replies?: [{ id, author, initials?, date, text }] }]
 *
 * Slides can be grouped into named sections, as in PowerPoint's slide
 * sorter. The first slide of a section carries its name, `slide.section`
 * (plain text), and the section runs until the next slide that has one.
 * Slides before the first such slide are in no section (PowerPoint's
 * "Default Section"). A section without slides can't be represented, so it
 * goes: a copied slide doesn't take the name along (see cloneSlide), and
 * moving or deleting a section's first slide passes the name to the next of
 * its slides (see moveSlidesWithSections and removeSlidesWithSections).
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

/** One of SIZES by name (own names only, so "__proto__" and the like aren't sizes). */
export const namedSize = (name) => (typeof name === 'string' && Object.hasOwn(SIZES, name) ? SIZES[name] : null);

export const slideSize = (deck) => namedSize(deck?.size) || customSize(deck?.size) || SIZES['16:9'];

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

const dateFormatters = [];

/** What a field shows on slide number `n` (from 1). */
export function fieldText(type, n, date = new Date()) {
  if (type === 'slidenum') return String(n);
  const i = Math.max(1, +type.slice(8) || 1) - 1;
  // Made once each: every slide drawn (every thumbnail) fills its fields.
  dateFormatters[i] ||= new Intl.DateTimeFormat(undefined, DATE_FORMATS[i]);
  return dateFormatters[i].format(date);
}

/**
 * An element's animations: `anim` (its entrance, ANIMATIONS), `emphasis`
 * (EMPHASES), `exit` (EXITS) and `motion` (a path, MOTIONS, with `path` for a
 * custom one), each { effect, duration (s), start: click|with|after, order },
 * sharing one playing order on the slide.
 */
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
  { id: 'bubble', name: 'Bubble' },
  { id: 'radar', name: 'Radar' },
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

// Diagrams (see diagram-layout.js): the kinds LibreSlides lays out, their colours and styles.
export const DIAGRAM_LAYOUTS = [
  { id: 'list', name: 'Basic list' },
  { id: 'vlist', name: 'Vertical list' },
  { id: 'process', name: 'Process' },
  { id: 'chevron', name: 'Chevron process' },
  { id: 'cycle', name: 'Cycle' },
  { id: 'hierarchy', name: 'Hierarchy' },
  { id: 'pyramid', name: 'Pyramid' },
  { id: 'venn', name: 'Basic Venn' },
  { id: 'matrix', name: 'Matrix' },
  { id: 'radial', name: 'Radial' },
];

export const DIAGRAM_COLORS = [
  ...[1, 2, 3, 4, 5, 6].map((n) => ({ id: `accent${n}`, name: `Accent ${n}` })),
  { id: 'colorful', name: 'Colourful' },
];

export const DIAGRAM_STYLES = [
  { id: 'flat', name: 'Flat' },
  { id: 'outline', name: 'Outline' },
  { id: 'raised', name: 'Raised (shadow)' },
];

// Text to start a new diagram of each kind with: [text, level].
const DIAGRAM_SAMPLES = {
  list: [['Research', 0], ['Ideas', 0], ['Feedback', 0], ['Results', 0]],
  vlist: [['Goals', 0], ['Grow the team', 1], ['Plan', 0], ['One step at a time', 1]],
  process: [['Plan', 0], ['Build', 0], ['Launch', 0]],
  chevron: [['Plan', 0], ['Design', 0], ['Build', 0], ['Launch', 0]],
  cycle: [['Plan', 0], ['Do', 0], ['Check', 0], ['Act', 0]],
  hierarchy: [['Director', 0], ['Manager', 1], ['Team', 2], ['Team', 2], ['Manager', 1], ['Team', 2]],
  pyramid: [['Vision', 0], ['Strategy', 0], ['Tactics', 0]],
  venn: [['People', 0], ['Process', 0], ['Technology', 0]],
  matrix: [['Strengths', 0], ['Weaknesses', 0], ['Opportunities', 0], ['Threats', 0]],
  radial: [['Core', 0], ['Idea', 1], ['Idea', 1], ['Idea', 1], ['Idea', 1]],
};

/** A diagram of the given kind with some sample text, ready to be edited. */
export function newDiagram(layout, x, y, w = 560, h = 320) {
  const kind = DIAGRAM_LAYOUTS.some((k) => k.id === layout) ? layout : 'list';
  return {
    id: uid(), type: 'diagram', x, y, w, h, rot: 0,
    diagram: { layout: kind, nodes: DIAGRAM_SAMPLES[kind].map(([t, level]) => ({ id: uid(), text: t, level })), colors: 'accent1', style: 'flat' },
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

/**
 * Elements in stacking order with each group's members next to each other,
 * where its topmost member is (as grouping puts them). A group can come
 * apart in a file from elsewhere, or an older version.
 */
export function gatherGroups(elements) {
  const top = new Map(); // groupId → index of its topmost member
  const members = new Map();
  elements.forEach((e, i) => {
    if (!e.groupId) return;
    top.set(e.groupId, i);
    if (!members.has(e.groupId)) members.set(e.groupId, []);
    members.get(e.groupId).push(e);
  });
  if (!top.size) return elements;
  return elements.flatMap((e, i) => (!e.groupId ? [e] : top.get(e.groupId) === i ? members.get(e.groupId) : []));
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
  delete copy.section; // a copy joins the section it's put in
  copy.elements = copy.elements.map((e) => ({ ...e, id: uid() }));
  if (Array.isArray(copy.comments)) {
    copy.comments = copy.comments.map((c) => ({ ...c, id: uid(), ...(c.replies ? { replies: c.replies.map((r) => ({ ...r, id: uid() })) } : {}) }));
  }
  return copy;
}

// ------------------------------------------------------------------ sections

/** PowerPoint's name for the slides before the first section. */
export const DEFAULT_SECTION = 'Default Section';

/**
 * The slides grouped by section, in order: [{ name, start, end }] (indices,
 * end exclusive). Slides before the first section make a group with no name.
 */
export function slideSections(slides) {
  const out = [];
  slides.forEach((s, i) => {
    if (i && !s.section) return;
    if (out.length) out.at(-1).end = i;
    out.push({ name: s.section || null, start: i, end: slides.length });
  });
  return out;
}

export const hasSections = (slides) => slides.some((s) => s.section);

/** The section group (see slideSections) that slide `i` is in. */
export const sectionAt = (slides, i) => slideSections(slides).find((g) => i >= g.start && i < g.end) || null;

/**
 * Section names by slide id once the slides with `ids` are taken out: a
 * section whose first slides go passes its name to its first slide that
 * stays. A section none of whose slides stay keeps it on its first slide.
 */
function liftSections(slides, ids) {
  const names = new Map(slides.map((s) => [s.id, s.section || null]));
  let pending = null; // the leaving first slide of the section we're in
  for (const s of slides) {
    if (ids.has(s.id)) {
      if (names.get(s.id)) pending = s.id;
    } else if (names.get(s.id)) pending = null;
    else if (pending) {
      names.set(s.id, names.get(pending));
      names.set(pending, null);
      pending = null;
    }
  }
  return names;
}

/** Give slides the section names in `names` (by slide id), in place. */
export function setSections(slides, names) {
  for (const s of slides) {
    if (names.get(s.id)) s.section = names.get(s.id);
    else delete s.section;
  }
}

/** The slides without those in `ids`, with the sections they started passed on to the slides that stay. */
export function removeSlidesWithSections(slides, ids) {
  const names = liftSections(slides, ids);
  const rest = slides.filter((s) => !ids.has(s.id));
  setSections(rest, names);
  return rest;
}

/**
 * Where the slides with `ids` go when moved to start at position `index`
 * (counted before the move), keeping sections sensible:
 * - the slides that stay keep their sections (see liftSections);
 * - moved slides join the section they're put after, or the first section
 *   when put at the very top, so a move never makes a section of no name;
 * - whole sections that move go between sections, never into the middle of
 *   one (above or below it, by the direction of the move), and after the
 *   slides that are in no section.
 * Returns { order, sections } (slides in their new order and section names
 * by slide id, for setSections), or null when nothing moves.
 */
export function moveSlidesWithSections(slides, ids, index) {
  const moving = slides.filter((s) => ids.has(s.id));
  if (!moving.length) return null;
  const names = liftSections(slides, ids);
  const rest = slides.filter((s) => !ids.has(s.id));
  const starts = (i) => !!names.get(rest[i].id);
  let at = slides.slice(0, index).filter((s) => !ids.has(s.id)).length;
  const from = slides.indexOf(moving[0]); // where the first of them was, among the slides that stay
  if (moving.some((s) => names.get(s.id))) {
    if (at < rest.length && !starts(at)) {
      if (at <= from) while (at > 0 && !starts(at)) at--;
      if (!starts(at)) while (at < rest.length && !starts(at)) at++;
    }
  } else if (at === 0 && rest.length && starts(0)) {
    names.set(moving[0].id, names.get(rest[0].id));
    names.set(rest[0].id, null);
  }
  const order = [...rest.slice(0, at), ...moving, ...rest.slice(at)];
  if (order.every((s, i) => s === slides[i])) return null;
  return { order, sections: names };
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
  // A diagram's drawing from PowerPoint, and its SmartArt parts, may have pictures in them.
  for (const child of el?.diagram?.drawn?.elements || []) slots.push(...mediaSlots(child));
  const source = el?.diagram?.source;
  for (const key of DIAGRAM_PARTS) if (source?.[key]) slots.push([source, key]);
  for (const refs of Object.values(source?.media || {})) for (const ref of refs) slots.push([ref, 'media']);
  return slots;
}

/** The SmartArt parts a diagram keeps (see diagram.source). */
export const DIAGRAM_PARTS = ['data', 'layout', 'style', 'colors', 'drawing'];

/** XML (a SmartArt part) as a data URL, to keep in deck.media. */
export function xmlDataUrl(xml) {
  const bytes = new TextEncoder().encode(xml);
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return `data:application/xml;base64,${btoa(bin)}`;
}

/** The XML in a data URL made by xmlDataUrl, or null. */
export function xmlOfDataUrl(url) {
  const m = /^data:application\/xml;base64,([a-z0-9+/]*={0,2})$/i.exec(url || '');
  if (!m) return null;
  try {
    return new TextDecoder().decode(Uint8Array.from(atob(m[1]), (c) => c.charCodeAt(0)));
  } catch {
    return null;
  }
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
    for (const n of el.diagram?.nodes || []) parts.push(n.text);
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
  deck.size = namedSize(deck.size) || customSize(deck.size) ? deck.size : '16:9';
  deck.themeId = themeById(deck.themeId).id;
  deck.title ||= 'Untitled presentation';
  if (!Array.isArray(deck.slides) || !deck.slides.length) deck.slides = [newSlide('title', deck)];
  for (const s of deck.slides) {
    s.id ||= uid();
    s.elements ||= [];
    s.notes ||= '';
    s.transition ||= { type: 'none', duration: 0.6 };
    if (!s.section) delete s.section;
    s.elements = gatherGroups(s.elements);
    for (const el of s.elements) {
      el.id ||= uid();
      el.rot ||= 0;
      if (canHoldText(el)) el.style ||= {};
    }
  }
  return deck;
}
