import DOMPurify from 'dompurify';
import { PATH_ARGS } from '../model/shapes.js';
import { normalizeTable, SIDES, CELL_MARGIN } from '../model/tables.js';
import { isMediaData } from './media-files.js';
import {
  CHART_KINDS, TRANSITIONS, SIZES, customSize, uid, ANIM_KINDS, ANIMATIONS, EMPHASES, EXITS, MOTIONS, FIELD, DIAGRAM_LAYOUTS, DIAGRAM_COLORS, DIAGRAM_STYLES, DIAGRAM_PARTS, xmlOfDataUrl,
} from '../model/deck.js';
import { normalizeLevels, MAX_LEVEL } from '../model/diagram-outline.js';
import { FONT_FACES } from '../model/fonts.js';
import { cleanWarp } from '../model/warps.js';
import { isPattern } from '../model/patterns.js';

/**
 * Text in slides is HTML. Anything that comes from outside (opened files,
 * restored backups, the clipboard) goes through here first.
 */
const TEXT_CONFIG = {
  ALLOWED_TAGS: ['p', 'div', 'br', 'span', 'b', 'strong', 'i', 'em', 'u', 's', 'strike', 'sub', 'sup', 'ul', 'ol', 'li', 'a', 'font'],
  ALLOWED_ATTR: ['style', 'href', 'color', 'face', 'size', 'start', 'data-field'],
  ALLOW_DATA_ATTR: false,
};

const UNSAFE_STYLE = /url\s*\(|expression\s*\(|image(-set)?\s*\(|cross-fade\s*\(|element\s*\(|@import|\\/i;

let hooked = false;
function ensureHooks() {
  if (hooked) return;
  hooked = true;
  DOMPurify.addHook('afterSanitizeAttributes', (node) => {
    // No external resources from inline styles (CSS escapes could spell url( in disguise); links open in a new tab.
    if (node.hasAttribute?.('style') && UNSAFE_STYLE.test(node.getAttribute('style'))) node.removeAttribute('style');
    if (node.hasAttribute?.('data-field') && !FIELD.test(node.getAttribute('data-field'))) node.removeAttribute('data-field');
    if (node.tagName === 'A') {
      node.setAttribute('target', '_blank');
      node.setAttribute('rel', 'noopener noreferrer');
    }
  });
}

export function sanitizeHtml(html) {
  if (!html) return '';
  ensureHooks();
  return DOMPurify.sanitize(String(html), { ...TEXT_CONFIG, ADD_ATTR: ['target', 'rel'] });
}

const isImageData = (v) => typeof v === 'string' && /^data:image\/(png|jpe?g|gif|webp|svg\+xml|bmp|avif);/i.test(v);

/**
 * A video or sound element: a known kind and a playable file, else it's
 * dropped. The poster must be a picture, and the play settings plain values.
 */
function keepMediaElement(el, deck) {
  if (el.kind !== 'video' && el.kind !== 'audio') return false;
  if (!isMediaData(deck.media?.[el.media])) return false;
  if (el.poster != null && !isImageData(deck.media?.[el.poster])) delete el.poster;
  el.autoplay = el.autoplay === true;
  el.loop = el.loop === true;
  if (el.duration != null && !(Number.isFinite(el.duration) && el.duration > 0)) delete el.duration;
  return true;
}

/**
 * A colour value we draw: hex, rgb()/rgba(), a theme colour, none or transparent.
 * Anything else (such as `url(…)`, which would fetch from the network) is dropped.
 */
const COLOR = /^(#[0-9a-f]{3,8}|rgba?\([\d\s.,%/]+\)|theme:[a-z0-9]+|none|transparent)$/i;
export const safeColor = (v) => (typeof v === 'string' && COLOR.test(v.trim()) ? v.trim() : null);

const isObject = (v) => !!v && typeof v === 'object' && !Array.isArray(v);
const fraction = (v) => (Number.isFinite(v) ? Math.max(0, Math.min(0.99, v)) : 0);
const DASHES = ['solid', 'dash', 'dot', 'dashDot', 'longDash'];

/** A table cell's own borders and margins: safe colours, known dashes and sensible numbers, or nothing. */
function cleanCellLook(c) {
  if (c.borders != null) {
    const b = isObject(c.borders) ? c.borders : {};
    const borders = {};
    for (const side of SIDES) {
      const s = b[side];
      const color = isObject(s) && safeColor(s.color);
      if (!color) continue;
      borders[side] = { color, width: Number.isFinite(s.width) ? Math.max(0, Math.min(50, s.width)) : 1, dash: DASHES.includes(s.dash) ? s.dash : 'solid' };
    }
    if (Object.keys(borders).length) c.borders = borders;
    else delete c.borders;
  }
  if (c.margin != null) {
    if (isObject(c.margin)) c.margin = Object.fromEntries(SIDES.map((side) => [side, Number.isFinite(c.margin[side]) ? Math.max(0, Math.min(200, c.margin[side])) : CELL_MARGIN[side]]));
    else delete c.margin;
  }
}

// How far from the slide anything may be, in points (slides are at most 56 inches a side).
const REACH = 1e5;
const clamp = (v, min, max) => Math.max(min, Math.min(max, v));
/** An id: a short plain string. */
const isId = (v) => typeof v === 'string' && /^[\w-]{1,64}$/.test(v);

/**
 * An element's numbers: places, sizes and turns finite and within reach of
 * the slide (they go into CSS and into .pptx XML as they are), see-through
 * between 0 and 1, text sizes and spacing sensible. Ones that aren't numbers
 * at all are put right or left out.
 */
function cleanNumbers(el) {
  // Lines go from point to point; everything else has a box.
  const needed = el.type === 'line' ? ['x1', 'y1', 'x2', 'y2'] : ['x', 'y', 'w', 'h'];
  for (const key of ['x', 'y', 'w', 'h', 'x1', 'y1', 'x2', 'y2', 'rot']) {
    if (el[key] !== undefined || needed.includes(key)) el[key] = Number.isFinite(el[key]) ? clamp(el[key], key === 'w' || key === 'h' ? 0 : -REACH, REACH) : 0;
  }
  if (el.opacity != null) {
    if (Number.isFinite(el.opacity)) el.opacity = clamp(el.opacity, 0, 1);
    else delete el.opacity;
  }
  if (el.radius != null && !(Number.isFinite(el.radius) && el.radius >= 0)) delete el.radius;
  if (el.natural != null && !(isObject(el.natural) && [el.natural.w, el.natural.h].every((v) => Number.isFinite(v) && v > 0))) delete el.natural;
  if (el.id != null && !isId(el.id)) el.id = uid();
  if (el.groupId != null && !isId(el.groupId)) delete el.groupId;
  const s = el.style;
  if (isObject(s)) {
    for (const [key, min, max] of [['size', 1, 4000], ['lineHeight', 0.1, 10], ['spacing', 0, 50], ['columnGap', 0, REACH]]) {
      if (s[key] == null) continue;
      if (Number.isFinite(s[key])) s[key] = clamp(s[key], min, max);
      else delete s[key];
    }
  } else if (el.style != null) delete el.style;
}

/** A gradient (see model/deck.js): a known type, finite numbers, 2–10 stops of safe colours; null for anything else. */
function cleanGradient(g) {
  const stops = isObject(g) && Array.isArray(g.stops) ? g.stops.filter(isObject).slice(0, 10).map((s) => {
    const stop = { pos: Number.isFinite(s.pos) ? clamp(s.pos, 0, 1) : 0, color: safeColor(s.color) };
    if (Number.isFinite(s.opacity) && s.opacity < 1) stop.opacity = Math.max(0, s.opacity);
    return stop;
  }).filter((s) => s.color) : [];
  if (stops.length < 2) return null;
  const out = { type: g.type === 'radial' ? 'radial' : 'linear', stops };
  // A radial gradient's centre may be outside the box (PowerPoint's can).
  if (out.type === 'radial') Object.assign(out, { cx: Number.isFinite(g.cx) ? clamp(g.cx, -2, 3) : 0.5, cy: Number.isFinite(g.cy) ? clamp(g.cy, -2, 3) : 0.5 });
  else out.angle = Number.isFinite(g.angle) ? g.angle : 180;
  return out;
}

/** A pattern fill (see model/patterns.js): a known pattern in safe colours, or null. */
function cleanPattern(p) {
  if (!isObject(p) || !isPattern(p.preset)) return null;
  return { preset: p.preset, fg: safeColor(p.fg) || '#000000', bg: safeColor(p.bg) || '#ffffff' };
}

/** A picture's tiling (see themes.js backgroundCss): finite sizes and offsets, a known alignment and flip; null for anything else. */
function cleanTile(t) {
  if (!isObject(t) || !(Number.isFinite(t.w) && t.w >= 0.5 && Number.isFinite(t.h) && t.h >= 0.5)) return null;
  const out = { w: Math.min(REACH, t.w), h: Math.min(REACH, t.h), x: Number.isFinite(t.x) ? clamp(t.x, -REACH, REACH) : 0, y: Number.isFinite(t.y) ? clamp(t.y, -REACH, REACH) : 0 };
  out.align = ['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br'].includes(t.align) ? t.align : 'tl';
  if (['x', 'y', 'xy'].includes(t.flip)) out.flip = t.flip;
  return out;
}

/** A slide background (see themes.js backgroundCss) with only safe colours, numbers and parts; null for anything else. */
function cleanBackground(bg) {
  if (!isObject(bg)) return null;
  if (bg.type === 'pattern') {
    const pattern = cleanPattern(bg.pattern);
    return pattern ? { type: 'pattern', pattern } : null;
  }
  for (const key of ['color', 'from', 'to']) if (bg[key] != null) bg[key] = safeColor(bg[key]) || '#ffffff';
  // The angle goes into CSS as is, so it must be a number.
  if (bg.angle != null) bg.angle = Number.isFinite(+bg.angle) ? +bg.angle : 180;
  if (bg.gradient != null) {
    bg.gradient = cleanGradient(bg.gradient);
    if (!bg.gradient) delete bg.gradient;
  }
  if (bg.tile != null) {
    bg.tile = cleanTile(bg.tile);
    if (!bg.tile) delete bg.tile;
  }
  return bg;
}

/** The theme a deck had in PowerPoint (see model/deck.js): its twelve colours and two fonts, or null. */
const THEME_COLOR_NAMES = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
export function cleanImportedTheme(t) {
  if (!isObject(t) || !isObject(t.colors)) return null;
  const colors = {};
  for (const name of THEME_COLOR_NAMES) {
    if (typeof t.colors[name] !== 'string' || !/^#[0-9a-f]{6}$/i.test(t.colors[name])) return null;
    colors[name] = t.colors[name].toLowerCase();
  }
  const fonts = isObject(t.fonts) ? t.fonts : {};
  return {
    name: plainLine(t.name, 100),
    colors,
    dark: t.dark === true,
    fonts: { major: safeFontFamily(fonts.major) || 'Calibri', minor: safeFontFamily(fonts.minor) || 'Calibri' },
  };
}

/** Clean an element from an untrusted source (file, backup, clipboard) in place. */
export function sanitizeElement(el) {
  cleanNumbers(el);
  if (el.text != null) el.text = typeof el.text === 'string' ? sanitizeHtml(el.text) : '';
  if (el.type === 'table' || el.cells != null) {
    el.cells = (Array.isArray(el.cells) ? el.cells : []).filter(Array.isArray).map((row) => row.map((c) => (isObject(c) ? c : { text: '' })));
    if (!el.cells.length) el.cells = [[{ text: '' }]];
    // A rectangular grid (merged cells keep their place in it).
    const cols = Math.max(1, ...el.cells.map((row) => row.length));
    for (const row of el.cells) while (row.length < cols) row.push({ text: '' });
    for (const row of el.cells) {
      for (const c of row) {
        c.text = typeof c.text === 'string' ? sanitizeHtml(c.text) : '';
        if (c.fill != null) c.fill = safeColor(c.fill);
        if (c.valign != null && !['top', 'middle', 'bottom'].includes(c.valign)) delete c.valign;
        cleanCellLook(c);
      }
    }
    for (const key of ['colWidths', 'rowHeights']) {
      const list = el[key];
      const n = key === 'colWidths' ? cols : el.cells.length;
      if (list == null) continue;
      // Shares of the table's width or height: positive, adding up to 1.
      const total = Array.isArray(list) && list.length === n && list.every((v) => Number.isFinite(v) && v > 0) ? list.reduce((a, v) => a + v, 0) : NaN;
      if (Number.isFinite(total)) el[key] = list.map((v) => v / total);
      else delete el[key];
    }
    if (!el.colWidths) el.colWidths = Array.from({ length: cols }, () => 1 / cols);
    normalizeTable(el);
  }
  if (el.paths != null) {
    // Freeform outlines: known commands with the right number of finite coordinates, within reason.
    let budget = 20000;
    el.paths = (Array.isArray(el.paths) ? el.paths : []).filter(isObject).map((p) => {
      const cmds = (Array.isArray(p.cmds) ? p.cmds : []).filter((c) => Array.isArray(c) && Object.hasOwn(PATH_ARGS, c[0])
        && c.length === PATH_ARGS[c[0]] + 1 && c.slice(1).every(Number.isFinite) && budget-- > 0);
      const out = { cmds: cmds.map((c) => [...c]) };
      if (p.fill === false) out.fill = false;
      if (p.stroke === false) out.stroke = false;
      if (['darken', 'darkenLess', 'lighten', 'lightenLess'].includes(p.shade)) out.shade = p.shade;
      return out;
    }).filter((p) => p.cmds.length);
    if (!el.paths.length) {
      delete el.paths;
      if (el.shape === 'custom') el.shape = 'rect';
    }
  }
  if (el.gradient != null) {
    el.gradient = cleanGradient(el.gradient);
    if (!el.gradient) delete el.gradient;
  }
  if (el.pattern != null) {
    el.pattern = cleanPattern(el.pattern);
    if (!el.pattern) delete el.pattern;
  }
  if (el.adjust != null) {
    const a = el.adjust;
    el.adjust = isObject(a) ? Object.fromEntries([['grayscale', a.grayscale === true || undefined], ...['brightness', 'contrast'].map((k) => [k, Number.isFinite(a[k]) ? Math.max(-1, Math.min(1, a[k])) : undefined])].filter(([, v]) => v)) : {};
    if (!Object.keys(el.adjust).length) delete el.adjust;
  }
  if (el.clip != null) {
    // Cropped to a shape: a shape name, or a freeform outline (checked as shapes' are).
    const c = el.clip;
    if (!isObject(c) || typeof c.shape !== 'string' || !/^[a-zA-Z0-9]{1,40}$/.test(c.shape)) delete el.clip;
    else if (c.shape === 'custom') {
      const holder = sanitizeElement({ shape: 'custom', paths: c.paths, preset: c.preset });
      if (holder.shape !== 'custom') delete el.clip;
      else el.clip = { shape: 'custom', paths: holder.paths, ...(holder.preset && { preset: holder.preset }) };
    } else el.clip = { shape: c.shape };
  }
  // Animations: known effects and starts, sensible numbers (the slideshow times them).
  for (const kind of ANIM_KINDS) {
    const a = el[kind];
    if (a == null) continue;
    const effects = { anim: ANIMATIONS, emphasis: EMPHASES, exit: EXITS, motion: MOTIONS }[kind];
    if (!isObject(a) || !effects.some((x) => x.id === a.effect) || a.effect === 'none') {
      delete el[kind];
      continue;
    }
    el[kind] = {
      effect: a.effect,
      duration: Number.isFinite(a.duration) ? Math.max(0.05, Math.min(30, a.duration)) : 0.5,
      start: ['click', 'with', 'after'].includes(a.start) ? a.start : 'click',
      order: Number.isFinite(a.order) ? a.order : 0,
    };
    if (kind === 'motion' && a.effect === 'custom') {
      // A path of up to 500 points, each within a few slides of the start.
      const path = (Array.isArray(a.path) ? a.path : []).filter((p) => Array.isArray(p) && p.length === 2 && p.every((v) => Number.isFinite(v) && Math.abs(v) <= 10)).slice(0, 500);
      if (path.length < 2) delete el[kind];
      else el[kind].path = path.map(([x, y]) => [x, y]);
    }
  }
  // A click on the object goes to a web page or another slide.
  if (el.link != null && !(typeof el.link === 'string' && el.link.length < 2000 && /^(https?:\/\/|mailto:|#slide:[\w-]+$)/i.test(el.link))) delete el.link;
  for (const end of ['arrowStart', 'arrowEnd']) {
    if (el[end] != null && el[end] !== true && el[end] !== false && !['arrow', 'stealth', 'diamond', 'oval', 'triangle'].includes(el[end])) el[end] = !!el[end];
  }
  if (el.textRect != null) {
    if (isObject(el.textRect)) el.textRect = Object.fromEntries(['l', 't', 'r', 'b'].map((k) => [k, Number.isFinite(el.textRect[k]) ? Math.max(-1, Math.min(1, el.textRect[k])) : 0]));
    else delete el.textRect;
  }
  if (el.preset != null) {
    // A PowerPoint preset name and its adjust values, written back into .pptx files.
    if (isObject(el.preset) && typeof el.preset.name === 'string' && /^[a-zA-Z0-9]{1,40}$/.test(el.preset.name)) {
      const adj = isObject(el.preset.adj) ? Object.fromEntries(Object.entries(el.preset.adj).filter(([k, v]) => /^adj\d{0,2}$/.test(k) && Number.isFinite(v))) : {};
      el.preset = Object.keys(adj).length ? { name: el.preset.name, adj } : { name: el.preset.name };
    } else delete el.preset;
  }
  if (el.type === 'chart' || el.chart != null) el.chart = sanitizeChart(el.chart);
  if (el.type === 'diagram' || el.diagram != null) el.diagram = sanitizeDiagram(el.diagram);
  if (el.crop != null) {
    if (isObject(el.crop)) for (const side of ['l', 't', 'r', 'b']) el.crop[side] = fraction(el.crop[side]);
    else delete el.crop;
  }
  if (el.fill != null) el.fill = safeColor(el.fill) || 'none';
  if (isObject(el.stroke)) {
    const { color, width, dash } = el.stroke;
    el.stroke = { color: safeColor(color) || 'none' };
    if (Number.isFinite(width)) el.stroke.width = clamp(width, 0, 1000);
    if (dash != null) el.stroke.dash = DASHES.includes(dash) ? dash : 'solid';
  } else if (el.stroke != null) el.stroke = null;
  // A shape's outline: a shape name (PowerPoint's presets among them).
  if (el.shape != null && !(typeof el.shape === 'string' && /^[a-zA-Z0-9]{1,40}$/.test(el.shape))) el.shape = 'rect';
  if (el.style && typeof el.style === 'object') {
    if (el.style.color != null) el.style.color = safeColor(el.style.color);
    if (typeof el.style.font === 'string') el.style.font = el.style.font.replace(/[^\p{L}\p{N}\s-]/gu, '').slice(0, 64);
    // Text box layout (see render.js textLayout): only what draws safely.
    const s = el.style;
    if (s.inset != null) {
      if (isObject(s.inset)) s.inset = Object.fromEntries(['l', 't', 'r', 'b'].map((k) => [k, Number.isFinite(s.inset[k]) ? Math.max(0, Math.min(500, s.inset[k])) : 0]));
      else delete s.inset;
    }
    if (s.vertical != null && !['vert', 'vert270', 'eaVert'].includes(s.vertical)) delete s.vertical;
    if (s.align != null && !['left', 'center', 'right', 'justify'].includes(s.align)) delete s.align;
    if (s.valign != null && !['top', 'middle', 'bottom'].includes(s.valign)) delete s.valign;
    if (s.font != null && typeof s.font !== 'string') delete s.font;
    if (s.wrap != null && s.wrap !== false) delete s.wrap;
    if (s.columns != null) s.columns = Math.max(1, Math.min(16, Math.round(+s.columns) || 1));
    // A WordArt warp: a known preset with adjust values in its range.
    if (s.warp != null) {
      const warp = cleanWarp(s.warp);
      if (warp) s.warp = warp;
      else delete s.warp;
    }
  }
  return el;
}

/**
 * Chart data from an untrusted source: a known kind, plain strings, finite
 * numbers (anything else is a blank point), safe colours and sensible sizes.
 */
const CHART_LIMITS = { series: 200, points: 1000, text: 255 };
const chartText = (v) => (typeof v === 'string' || Number.isFinite(v) ? String(v).slice(0, CHART_LIMITS.text) : '');

export function sanitizeChart(chart) {
  const c = isObject(chart) ? chart : {};
  const out = {
    kind: CHART_KINDS.some((k) => k.id === c.kind) ? c.kind : 'column',
    title: chartText(c.title),
    categories: (Array.isArray(c.categories) ? c.categories : []).slice(0, CHART_LIMITS.points).map(chartText),
    series: (Array.isArray(c.series) ? c.series : []).filter(isObject).slice(0, CHART_LIMITS.series).map((s, i) => {
      const ser = {
        name: chartText(s.name) || `Series ${i + 1}`,
        values: (Array.isArray(s.values) ? s.values : []).slice(0, CHART_LIMITS.points).map((v) => (typeof v === 'number' && Number.isFinite(v) ? v : null)),
      };
      if (safeColor(s.color)) ser.color = safeColor(s.color);
      if (Array.isArray(s.pointColors)) ser.pointColors = s.pointColors.slice(0, CHART_LIMITS.points).map((p) => safeColor(p) || null);
      if (['column', 'line', 'area'].includes(s.kind)) ser.kind = s.kind;
      if (s.secondary === true) ser.secondary = true;
      return ser;
    }),
    legend: ['right', 'bottom', 'top', 'left', 'none'].includes(c.legend) ? c.legend : 'none',
    dataLabels: c.dataLabels === true,
  };
  if (['clustered', 'stacked', 'percent'].includes(c.grouping)) out.grouping = c.grouping;
  if (typeof c.markers === 'boolean' && (c.markers || c.kind === 'scatter')) out.markers = c.markers;
  for (const key of ['lines', 'smooth']) if (c[key] === true) out[key] = true;
  for (const key of ['numberFormat', 'labelFormat']) if (chartFormat(c[key])) out[key] = c[key];
  if (Array.isArray(c.labelContent)) {
    const parts = LABEL_PARTS.filter((p) => c.labelContent.includes(p));
    if (parts.length) out.labelContent = parts;
  }
  if (LABEL_POSITIONS.includes(c.labelPosition)) out.labelPosition = c.labelPosition;
  for (const [key, min, max] of [['gapWidth', 0, 500], ['overlap', -100, 100], ['holeSize', 10, 90]]) {
    if (Number.isFinite(c[key])) out[key] = Math.min(max, Math.max(min, c[key]));
  }
  for (const key of ['valueAxis', 'secondaryAxis', 'categoryAxis']) {
    const axis = chartAxis(c[key], key !== 'categoryAxis', key === 'secondaryAxis');
    if (axis) out[key] = axis;
  }
  return out;
}

/**
 * A diagram from an untrusted source (see model/deck.js): a known layout,
 * colours and style, an outline of plain-text nodes of sensible length and
 * number, and PowerPoint's drawing as elements cleaned like any other. The
 * SmartArt parts it came with are checked in deck.media (see isXmlPart), and
 * dropped with the pictures they need by sanitizeDeck when they're missing.
 */
export const DIAGRAM_LIMITS = { nodes: 500, text: 1000, id: 64, drawn: 2000, part: 4e6 };
const known = (list, v, fallback) => (list.some((k) => k.id === v) ? v : fallback);
const plainLine = (v, max) => (typeof v === 'string' || Number.isFinite(v) ? String(v).replace(/[\u0000-\u001f\u007f]+/g, ' ').slice(0, max) : '');
const DRAWN_TYPES = ['text', 'shape', 'line', 'image'];
const mediaId = (v) => typeof v === 'string' && /^[\w-]{1,64}$/.test(v);

/** A SmartArt part in deck.media: XML (as a data URL, see xmlDataUrl) of sensible size that starts like XML and parses as it. */
export function isXmlPart(v) {
  if (typeof v !== 'string' || v.length > DIAGRAM_LIMITS.part * 1.4) return false;
  const xml = xmlOfDataUrl(v);
  if (!xml || xml.length > DIAGRAM_LIMITS.part || !/^\ufeff?\s*</.test(xml)) return false;
  try {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    return !!doc.documentElement && !doc.getElementsByTagName('parsererror').length;
  } catch {
    return false;
  }
}

export function sanitizeDiagram(diagram) {
  const d = isObject(diagram) ? diagram : {};
  const ids = new Set();
  const nodes = (Array.isArray(d.nodes) ? d.nodes : []).filter(isObject).slice(0, DIAGRAM_LIMITS.nodes).map((n) => {
    let id = typeof n.id === 'string' && n.id.length <= DIAGRAM_LIMITS.id ? n.id : '';
    if (!id || ids.has(id)) id = uid();
    ids.add(id);
    return { id, text: plainLine(n.text, DIAGRAM_LIMITS.text), level: Number.isInteger(n.level) ? Math.max(0, Math.min(MAX_LEVEL, n.level)) : 0 };
  });
  const out = {
    layout: known(DIAGRAM_LAYOUTS, d.layout, 'list'),
    nodes: normalizeLevels(nodes),
    colors: known(DIAGRAM_COLORS, d.colors, 'accent1'),
    style: known(DIAGRAM_STYLES, d.style, 'flat'),
  };
  if (typeof d.font === 'string' && safeFontFamily(d.font)) out.font = safeFontFamily(d.font);
  if (Array.isArray(d.palette) && d.palette.length === 6 && d.palette.every((c) => typeof c === 'string' && /^#[0-9a-f]{6}$/i.test(c))) out.palette = [...d.palette];
  if (d.edited === true) out.edited = true;
  const drawn = d.drawn;
  if (isObject(drawn) && Array.isArray(drawn.elements) && Number.isFinite(drawn.w) && Number.isFinite(drawn.h) && drawn.w > 0 && drawn.h > 0) {
    const elements = drawn.elements.filter((e) => isObject(e) && DRAWN_TYPES.includes(e.type)).slice(0, DIAGRAM_LIMITS.drawn).map(sanitizeElement);
    if (elements.length) out.drawn = { w: drawn.w, h: drawn.h, elements };
  }
  const src = d.source;
  if (isObject(src)) {
    const source = {};
    for (const key of DIAGRAM_PARTS) if (mediaId(src[key])) source[key] = src[key];
    const look = isObject(src.look) ? src.look : {};
    source.look = { layout: known(DIAGRAM_LAYOUTS, look.layout, out.layout), colors: known(DIAGRAM_COLORS, look.colors, out.colors), style: known(DIAGRAM_STYLES, look.style, out.style) };
    // Pictures the parts refer to: relationship ids and deck pictures (checked against deck.media in sanitizeDeck).
    if (isObject(src.media)) {
      const media = {};
      for (const key of ['data', 'drawing']) {
        const refs = Array.isArray(src.media[key]) ? src.media[key].filter((r) => isObject(r) && typeof r.id === 'string' && /^[\w-]{1,40}$/.test(r.id) && mediaId(r.media)) : [];
        if (refs.length) media[key] = refs.slice(0, 200).map((r) => ({ id: r.id, media: r.media }));
      }
      if (Object.keys(media).length) source.media = media;
    }
    if (source.data && source.layout && source.style && source.colors) out.source = source;
  }
  return out;
}

/** A number format code: a short line of text (it goes into the chart's XML escaped). */
const chartFormat = (v) => typeof v === 'string' && v.length > 0 && v.length <= 100 && !/[\u0000-\u001f]/.test(v);
const LABEL_PARTS = ['category', 'value', 'percent'];
const LABEL_POSITIONS = ['outEnd', 'inEnd', 'ctr', 'inBase', 'bestFit', 't', 'b', 'l', 'r'];

/** A chart axis's title, (value axes) scale and gridlines and (the secondary one) number format, or null when it sets none. */
function chartAxis(a, scaled, formatted) {
  if (!isObject(a)) return null;
  const out = {};
  if (chartText(a.title)) out.title = chartText(a.title);
  if (formatted && chartFormat(a.numberFormat)) out.numberFormat = a.numberFormat;
  if (scaled) {
    for (const key of ['min', 'max']) if (Number.isFinite(a[key])) out[key] = a[key];
    if (Number.isFinite(a.unit) && a.unit > 0) out.unit = a.unit;
    if (typeof a.gridlines === 'boolean') out.gridlines = a.gridlines;
  }
  return Object.keys(out).length ? out : null;
}

/**
 * Slide comments from an untrusted source: plain objects with plain-text
 * strings of sensible length, finite positions and a capped number of
 * comments and replies. Anything else is dropped.
 */
const COMMENT_LIMITS = { comments: 2000, replies: 500, text: 10000, author: 100, initials: 10, date: 40, id: 64 };
const str = (v, max) => (typeof v === 'string' ? v.slice(0, max) : '');

function commentEntry(c) {
  const out = { id: str(c.id, COMMENT_LIMITS.id) || uid(), author: str(c.author, COMMENT_LIMITS.author), text: str(c.text, COMMENT_LIMITS.text) };
  if (str(c.initials, COMMENT_LIMITS.initials)) out.initials = str(c.initials, COMMENT_LIMITS.initials);
  // Dates stay ISO strings; one that can't be read is left out.
  if (typeof c.date === 'string' && c.date.length <= COMMENT_LIMITS.date && !Number.isNaN(Date.parse(c.date))) out.date = c.date;
  return out;
}

export function sanitizeComments(list) {
  return (Array.isArray(list) ? list : []).filter(isObject).slice(0, COMMENT_LIMITS.comments).map((c) => {
    const out = commentEntry(c);
    if (Number.isFinite(c.x) && Number.isFinite(c.y)) Object.assign(out, { x: clamp(c.x, -REACH, REACH), y: clamp(c.y, -REACH, REACH) });
    if (c.resolved === true) out.resolved = true;
    const replies = (Array.isArray(c.replies) ? c.replies : []).filter(isObject).slice(0, COMMENT_LIMITS.replies).map(commentEntry);
    if (replies.length) out.replies = replies;
    return out;
  });
}

/** A section name (plain text on one line, of sensible length), or '' for anything else. */
export const SECTION_NAME_MAX = 255;
export const sectionName = (v) => (typeof v === 'string' ? v.replace(/[\u0000-\u001f\u007f]+/g, ' ').trim().slice(0, SECTION_NAME_MAX).trim() : '');

/** Clean every piece of HTML, colour and picture in a deck from an untrusted source. */
export function sanitizeDeck(deck) {
  // Damaged or hand-edited files: keep only what has the expected shape.
  if (typeof deck.title !== 'string') deck.title = '';
  if (!(typeof deck.size === 'string' && (Object.hasOwn(SIZES, deck.size) || customSize(deck.size)))) deck.size = '16:9';
  if (!isObject(deck.media)) deck.media = {};
  deck.slides = (Array.isArray(deck.slides) ? deck.slides : []).filter(isObject);
  for (const s of deck.slides) {
    if (!isId(s.id)) s.id = uid();
    s.notes = typeof s.notes === 'string' ? s.notes : '';
    s.hidden = s.hidden === true;
    const t = s.transition;
    s.transition = isObject(t) && TRANSITIONS.some((x) => x.id === t.type)
      ? { type: t.type, duration: Number.isFinite(t.duration) ? clamp(t.duration, 0, 10) : 0.6 }
      : { type: 'none', duration: 0.6 };
    s.elements = (Array.isArray(s.elements) ? s.elements : []).filter(isObject);
    for (const el of s.elements) sanitizeElement(el);
    if (s.comments != null) s.comments = sanitizeComments(s.comments);
    if (!s.comments?.length) delete s.comments;
    if (s.section != null) s.section = sectionName(s.section);
    if (!s.section) delete s.section;
    s.background = cleanBackground(s.background);
  }
  for (const [id, data] of Object.entries(deck.media || {})) {
    if (isMediaData(data)) continue; // videos and sounds only ever play through <video> and <audio>
    if (/^data:application\/xml[;,]/i.test(data)) {
      if (!isXmlPart(data)) delete deck.media[id]; // diagrams' SmartArt parts, only ever written into .pptx files
    } else if (!isImageData(data)) delete deck.media[id];
    else if (/^data:image\/svg\+xml/i.test(data)) {
      // SVG can carry script; it only ever renders through <img>, which doesn't run it, but keep it tidy.
      try {
        const [head, body] = data.split(',');
        const svg = head.endsWith(';base64') ? atob(body) : decodeURIComponent(body);
        deck.media[id] = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(DOMPurify.sanitize(svg, { USE_PROFILES: { svg: true, svgFilters: true } }))}`;
      } catch {
        delete deck.media[id];
      }
    }
  }
  // Pictures that were dropped (unsupported or unsafe) leave nothing behind.
  const keepPictures = (list) => {
    const out = list.filter((el) => el.type !== 'image' || deck.media?.[el.media]);
    for (const el of out) {
      if (el.fillImage == null) continue;
      if (!isObject(el.fillImage) || !deck.media?.[el.fillImage.media]) delete el.fillImage;
      else {
        const { crop, tile, opacity } = el.fillImage;
        el.fillImage = { media: el.fillImage.media };
        if (isObject(crop)) el.fillImage.crop = Object.fromEntries(['l', 't', 'r', 'b'].map((k) => [k, fraction(crop[k])]));
        if (cleanTile(tile)) el.fillImage.tile = cleanTile(tile);
        if (Number.isFinite(opacity) && opacity < 1) el.fillImage.opacity = Math.max(0, opacity);
      }
    }
    return out;
  };
  for (const s of deck.slides || []) {
    s.elements = keepPictures(s.elements || []);
    for (const el of s.elements) {
      const drawn = el.diagram?.drawn;
      if (drawn) drawn.elements = keepPictures(drawn.elements);
      if (drawn && !drawn.elements.length) delete el.diagram.drawn;
      // SmartArt parts that are missing, or whose pictures are, can't be written back.
      const source = el.diagram?.source;
      if (!source) continue;
      if (source.drawing && !/^data:application\/xml/.test(deck.media?.[source.drawing] || '')) delete source.drawing;
      const parts = DIAGRAM_PARTS.filter((key) => source[key]).every((key) => /^data:application\/xml/.test(deck.media?.[source[key]] || ''));
      const refs = Object.values(source.media || {}).flat();
      if (!parts || refs.some((r) => !isImageData(deck.media?.[r.media]))) delete el.diagram.source;
    }
    s.elements = s.elements.filter((el) => el.type !== 'media' || keepMediaElement(el, deck));
    if (s.background?.type === 'image' && !deck.media?.[s.background.media]) s.background = null;
  }
  sanitizeFonts(deck);
  if (deck.importedTheme != null) {
    deck.importedTheme = cleanImportedTheme(deck.importedTheme);
    if (!deck.importedTheme) delete deck.importedTheme;
  }
  return deck;
}

// ------------------------------------------------------------------ embedded fonts

/** A font family name with only letters, digits, spaces and hyphens (as element fonts are kept), or ''. */
export const safeFontFamily = (name) => (typeof name === 'string' ? name.replace(/[^\p{L}\p{N}\s-]/gu, '').trim().slice(0, 64) : '');

const FONT_DATA = /^data:font\/(ttf|otf|woff2?);base64,[a-z0-9+/]*={0,2}$/i;
export const isFontData = (v) => typeof v === 'string' && v.length < 17e6 && FONT_DATA.test(v);

/** deck.fonts ({ family: { regular|bold|italic|boldItalic: data URL } }): only font data under safe names. */
function sanitizeFonts(deck) {
  if (deck.fonts == null) return;
  const fonts = {};
  for (const [name, faces] of Object.entries(isObject(deck.fonts) ? deck.fonts : {})) {
    const family = safeFontFamily(name);
    if (!family || !isObject(faces)) continue;
    for (const face of Object.keys(FONT_FACES)) {
      if (!Object.hasOwn(faces, face) || !isFontData(faces[face])) continue;
      if (!Object.hasOwn(fonts, family)) fonts[family] = {};
      fonts[family][face] = faces[face];
    }
  }
  if (Object.keys(fonts).length) deck.fonts = fonts;
  else delete deck.fonts;
}
