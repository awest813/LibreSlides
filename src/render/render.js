import { themeById, resolveColor, backgroundCss } from '../model/themes.js';
import { slideSize, isTitleLayout, PLACEHOLDER_HINTS, canHoldText } from '../model/deck.js';
import { shapePath, customPath, textInset } from '../model/shapes.js';
import { renderChart } from './chart.js';
import { renderMedia } from './media.js';
import { fontStack } from '../model/fonts.js';
import { registerDeckFonts } from './deck-fonts.js';

/**
 * Renders slides to DOM at 1px = 1pt. The same renderer draws the editor
 * canvas, thumbnails, the slideshow and print, so they always match.
 */

const SVG = 'http://www.w3.org/2000/svg';
const svgEl = (tag, attrs = {}) => {
  const el = document.createElementNS(SVG, tag);
  for (const [k, v] of Object.entries(attrs)) if (v != null) el.setAttribute(k, v);
  return el;
};

// ------------------------------------------------------------------ media

const blobUrls = new Map(); // media id → { data, url }

/** An object URL for a deck picture, video or sound (cheaper to reuse across thumbnails than its data URL). */
export function mediaUrl(deck, id) {
  const data = deck?.media?.[id];
  if (!data) return '';
  const hit = blobUrls.get(id);
  if (hit && hit.data === data) return hit.url;
  let url = data;
  try {
    const comma = data.indexOf(',');
    const head = data.slice(5, comma);
    const mime = head.split(';')[0] || 'application/octet-stream';
    let bytes;
    if (head.endsWith(';base64')) {
      const bin = atob(data.slice(comma + 1));
      bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    } else {
      bytes = new TextEncoder().encode(decodeURIComponent(data.slice(comma + 1)));
    }
    url = URL.createObjectURL(new Blob([bytes], { type: mime }));
  } catch {
    /* keep the data URL */
  }
  if (hit && hit.url !== hit.data) URL.revokeObjectURL(hit.url);
  blobUrls.set(id, { data, url });
  return url;
}

// ------------------------------------------------------------------ text

/** The effective text style of an element: its own settings over the theme's defaults. */
export function textStyle(el, theme) {
  const s = el.style || {};
  const isTitle = el.ph === 'title';
  const isHeading = isTitle || el.ph === 'subtitle';
  return {
    font: s.font || (isTitle ? theme.fonts.title : theme.fonts.body),
    size: s.size || 18,
    color: s.color || (isTitle ? 'theme:title' : el.ph === 'subtitle' ? 'theme:text' : 'theme:text'),
    bold: s.bold ?? (isTitle ? !!theme.titleBold : false),
    italic: !!s.italic,
    underline: !!s.underline,
    strike: !!s.strike,
    align: s.align || 'left',
    valign: s.valign || (el.type === 'shape' ? 'middle' : 'top'),
    lineHeight: s.lineHeight || (isHeading ? 1.1 : 1.25),
    spacing: s.spacing ?? (isHeading ? 0 : 0.35),
  };
}

export function applyTextStyle(box, el, theme) {
  const t = textStyle(el, theme);
  Object.assign(box.style, {
    fontFamily: fontStack(t.font),
    fontSize: `${t.size}px`,
    color: resolveColor(t.color, theme, '#000'),
    fontWeight: t.bold ? '700' : '400',
    fontStyle: t.italic ? 'italic' : 'normal',
    textDecoration: [t.underline && 'underline', t.strike && 'line-through'].filter(Boolean).join(' ') || 'none',
    textAlign: t.align,
    justifyContent: t.valign === 'middle' ? 'center' : t.valign === 'bottom' ? 'flex-end' : 'flex-start',
    lineHeight: String(t.lineHeight),
  });
  box.style.setProperty('--para-gap', `${t.spacing}em`);
  return t;
}

/**
 * Text box options from PowerPoint: margins (`inset` { l, t, r, b } in
 * points), no wrapping, vertical text ('vert' turned a quarter clockwise,
 * 'vert270' anticlockwise, 'eaVert' East Asian, upright) and columns.
 */
function textLayout(box, body, s) {
  if (s.inset) box.style.padding = `${s.inset.t}px ${s.inset.r}px ${s.inset.b}px ${s.inset.l}px`;
  if (s.wrap === false) box.style.whiteSpace = 'pre';
  if (s.vertical) {
    box.style.writingMode = 'vertical-rl';
    box.style.textOrientation = s.vertical === 'eaVert' ? 'upright' : 'sideways';
    if (s.vertical === 'vert270') box.style.transform = 'rotate(180deg)';
  }
  if (s.columns > 1) {
    Object.assign(body.style, { columnCount: String(s.columns), columnGap: `${s.columnGap ?? 7.2}px`, columnFill: 'auto', height: '100%' });
  }
}

function textBox(el, ctx) {
  const box = document.createElement('div');
  box.className = 'el-textbox';
  if (el.type === 'shape') {
    // Preset shapes from PowerPoint say where their text goes; LibreSlides' own shapes keep it inside the outline.
    const ins = el.textRect || (({ x, y, b }) => ({ l: x, r: x, t: y, b }))(textInset(el.shape));
    box.style.left = `${ins.l * el.w}px`;
    box.style.right = `${ins.r * el.w}px`;
    box.style.top = `${ins.t * el.h}px`;
    box.style.bottom = `${ins.b * el.h}px`;
  }
  applyTextStyle(box, el, ctx.theme);
  const body = document.createElement('div');
  body.className = 'el-text-body';
  if (el.bullets) body.classList.add('has-bullets');
  body.innerHTML = el.text || '';
  textLayout(box, body, el.style || {});
  if (ctx.editing && el.ph && !el.text) body.dataset.placeholder = PLACEHOLDER_HINTS[el.ph] || '';
  box.append(body);
  return box;
}

// ------------------------------------------------------------------ elements

const DASHES = { solid: null, dash: '8 5', dot: '2 4', dashDot: '8 4 2 4', longDash: '14 6' };
export const dashArray = (dash, width = 1) => {
  const d = DASHES[dash];
  return d ? d.split(' ').map((n) => n * Math.max(1, width / 2)).join(' ') : null;
};

function strokeAttrs(stroke, theme) {
  if (!stroke || !stroke.width || stroke.color === 'none' || !stroke.color) return { stroke: 'none' };
  return {
    stroke: resolveColor(stroke.color, theme, '#000'),
    'stroke-width': stroke.width,
    'stroke-dasharray': dashArray(stroke.dash, stroke.width),
    'stroke-linejoin': 'round',
  };
}

function arrowHead(x1, y1, x2, y2, width, color) {
  const size = Math.max(7, width * 3.2);
  const a = Math.atan2(y2 - y1, x2 - x1);
  const p = (ang) => `${x2 - size * Math.cos(a + ang)},${y2 - size * Math.sin(a + ang)}`;
  return svgEl('polygon', { points: `${x2},${y2} ${p(0.45)} ${p(-0.45)}`, fill: color });
}

let gradientCount = 0;

/**
 * Add a shape's gradient (see model/deck.js) to its SVG and return the paint
 * that uses it. Angles are CSS ones: 90 runs left to right, 180 top to bottom.
 */
export function gradientPaint(svg, g, w, h, theme) {
  const id = `ls-grad-${++gradientCount}`;
  let paint;
  if (g.type === 'radial') {
    const cx = (g.cx ?? 0.5) * w;
    const cy = (g.cy ?? 0.5) * h;
    // Out to the farthest corner, as PowerPoint's "from centre" gradients go.
    const r = Math.max(...[[0, 0], [w, 0], [0, h], [w, h]].map(([x, y]) => Math.hypot(x - cx, y - cy)));
    paint = svgEl('radialGradient', { id, gradientUnits: 'userSpaceOnUse', cx, cy, r });
  } else {
    const a = (((g.angle ?? 180) - 90) * Math.PI) / 180;
    const r = Math.abs((w / 2) * Math.cos(a)) + Math.abs((h / 2) * Math.sin(a));
    paint = svgEl('linearGradient', {
      id, gradientUnits: 'userSpaceOnUse',
      x1: w / 2 - Math.cos(a) * r, y1: h / 2 - Math.sin(a) * r, x2: w / 2 + Math.cos(a) * r, y2: h / 2 + Math.sin(a) * r,
    });
  }
  for (const s of g.stops || []) {
    paint.append(svgEl('stop', { offset: Math.max(0, Math.min(1, s.pos)), 'stop-color': resolveColor(s.color, theme, '#000'), 'stop-opacity': s.opacity ?? 1 }));
  }
  const defs = svgEl('defs');
  defs.append(paint);
  svg.append(defs);
  return `url(#${id})`;
}

/** The outline a picture is cropped to (see deck.js), as SVG path data for a w × h box. */
function clipPath(clip, w, h) {
  if (clip.shape === 'custom' && Array.isArray(clip.paths)) return clip.paths.map((p) => customPath(p.cmds, w, h)).join('');
  return clip.shape && clip.shape !== 'rect' ? shapePath(clip.shape, w, h) : null;
}

/** A picture's colour adjustments as CSS filters: greyscale, brightness and contrast (-1 to 1, 0 unchanged). */
export function pictureFilters(adjust) {
  if (!adjust) return [];
  const out = [];
  if (adjust.grayscale) out.push('grayscale(1)');
  if (adjust.brightness) out.push(`brightness(${Math.max(0, 1 + adjust.brightness)})`);
  if (adjust.contrast) out.push(`contrast(${Math.max(0, 1 + adjust.contrast)})`);
  return out;
}

/** A picture stretched over a shape (keeping only the part its crop leaves), as an SVG paint. */
function picturePaint(svg, el, url) {
  const id = `ls-pic-${++gradientCount}`;
  const { l = 0, t = 0, r = 0, b = 0 } = el.fillImage.crop || {};
  const iw = el.w / Math.max(0.01, 1 - l - r);
  const ih = el.h / Math.max(0.01, 1 - t - b);
  const pattern = svgEl('pattern', { id, patternUnits: 'userSpaceOnUse', width: el.w, height: el.h });
  pattern.append(svgEl('image', { href: url, x: -l * iw, y: -t * ih, width: iw, height: ih, preserveAspectRatio: 'none' }));
  const defs = svgEl('defs');
  defs.append(pattern);
  svg.append(defs);
  return `url(#${id})`;
}

const SHADES = { darken: 'rgba(0,0,0,0.25)', darkenLess: 'rgba(0,0,0,0.12)', lighten: 'rgba(255,255,255,0.4)', lightenLess: 'rgba(255,255,255,0.2)' };

/**
 * Where a freeform's line starts and ends, each as [from x, from y, tip x, tip y]
 * along its direction there, for arrowheads (connectors' routes).
 */
function pathEnds(paths, w, h) {
  const cmds = paths.flatMap((p) => p.cmds).filter((c) => c[0] !== 'Z');
  if (cmds.length < 2) return null;
  const pt = (c, i) => [c[i] * w, c[i + 1] * h];
  const last = cmds[cmds.length - 1];
  const before = cmds[cmds.length - 2];
  // A command's own point is its last two numbers.
  const end = (c) => pt(c, c.length - 2);
  const tip = end(last);
  // Towards the tip from the last control point, or from the point before.
  const from = last[0] === 'C' ? pt(last, 3) : last[0] === 'Q' ? pt(last, 1) : end(before);
  const first = end(cmds[0]);
  const next = cmds[1][0] === 'C' || cmds[1][0] === 'Q' ? pt(cmds[1], 1) : end(cmds[1]);
  return { end: [...from, ...tip], start: [...next, ...first] };
}

/** The DOM for one element. */
export function renderElement(el, ctx) {
  const { theme } = ctx;
  const div = document.createElement('div');
  div.className = `el el-${el.type}`;
  div.dataset.id = el.id;
  if (el.anim && el.anim.effect !== 'none') div.dataset.anim = el.anim.effect;
  if (el.opacity != null && el.opacity < 1) div.style.opacity = String(el.opacity);

  if (el.type === 'line') {
    const minX = Math.min(el.x1, el.x2);
    const minY = Math.min(el.y1, el.y2);
    Object.assign(div.style, { left: `${minX}px`, top: `${minY}px`, width: `${Math.abs(el.x2 - el.x1)}px`, height: `${Math.abs(el.y2 - el.y1)}px` });
    const svg = svgEl('svg', { class: 'el-svg', overflow: 'visible', width: '100%', height: '100%' });
    const [x1, y1, x2, y2] = [el.x1 - minX, el.y1 - minY, el.x2 - minX, el.y2 - minY];
    const sa = strokeAttrs({ ...el.stroke, color: el.stroke?.color === 'none' ? 'theme:text' : el.stroke?.color }, theme);
    // A wide invisible stroke makes thin lines easy to click.
    svg.append(svgEl('line', { x1, y1, x2, y2, class: 'hit', stroke: 'transparent', 'stroke-width': Math.max(12, (el.stroke?.width || 2) + 10) }));
    svg.append(svgEl('line', { x1, y1, x2, y2, ...sa, 'stroke-linecap': 'round' }));
    if (el.arrowEnd) svg.append(arrowHead(x1, y1, x2, y2, el.stroke?.width || 2, sa.stroke));
    if (el.arrowStart) svg.append(arrowHead(x2, y2, x1, y1, el.stroke?.width || 2, sa.stroke));
    div.append(svg);
    return div;
  }

  Object.assign(div.style, { left: `${el.x}px`, top: `${el.y}px`, width: `${el.w}px`, height: `${el.h}px` });
  if (el.rot) div.style.transform = `rotate(${el.rot}deg)`;
  const shadow = el.shadow ? 'drop-shadow(0 4px 8px rgba(0,0,0,0.3))' : null;

  if (el.type === 'shape' || (el.type === 'text' && ((el.fill && el.fill !== 'none') || (el.stroke && el.stroke.color !== 'none' && el.stroke.width)))) {
    const svg = svgEl('svg', { class: 'el-svg', overflow: 'visible', width: el.w, height: el.h, viewBox: `0 0 ${el.w} ${el.h}`, preserveAspectRatio: 'none' });
    const picture = el.fillImage && ctx.deck?.media?.[el.fillImage.media];
    const fill = picture ? picturePaint(svg, el, mediaUrl(ctx.deck, el.fillImage.media))
      : el.gradient && el.fill !== 'none' ? gradientPaint(svg, el.gradient, el.w, el.h, theme)
        : resolveColor(el.fill, theme, 'transparent');
    const stroke = strokeAttrs(el.stroke, theme);
    if (el.type === 'shape' && el.shape === 'custom' && Array.isArray(el.paths)) {
      for (const p of el.paths) {
        const d = customPath(p.cmds, el.w, el.h);
        svg.append(svgEl('path', { d, fill: p.fill === false ? 'none' : fill, ...(p.stroke === false ? { stroke: 'none' } : stroke) }));
        // Faces drawn lighter or darker (a can's top, a cube's side).
        if (SHADES[p.shade] && p.fill !== false && el.fill !== 'none') svg.append(svgEl('path', { d, fill: SHADES[p.shade], stroke: 'none' }));
      }
      if ((el.arrowStart || el.arrowEnd) && stroke.stroke !== 'none') {
        const ends = pathEnds(el.paths, el.w, el.h);
        if (ends && el.arrowEnd) svg.append(arrowHead(...ends.end, el.stroke?.width || 2, stroke.stroke));
        if (ends && el.arrowStart) svg.append(arrowHead(...ends.start, el.stroke?.width || 2, stroke.stroke));
      }
    } else svg.append(svgEl('path', { d: shapePath(el.type === 'shape' ? el.shape : 'rect', el.w, el.h), fill, ...stroke }));
    if (shadow) svg.style.filter = shadow;
    if (el.flipH || el.flipV) svg.style.transform = `scale(${el.flipH ? -1 : 1}, ${el.flipV ? -1 : 1})`;
    div.append(svg);
  }
  if (canHoldText(el)) div.append(textBox(el, ctx));

  if (el.type === 'image') {
    const img = document.createElement('img');
    img.src = mediaUrl(ctx.deck, el.media);
    img.alt = el.alt || '';
    img.draggable = false;
    img.className = 'el-img';
    img.style.objectFit = el.fit === 'cover' ? 'cover' : el.fit === 'contain' ? 'contain' : 'fill';
    // A cropped picture (from PowerPoint): the whole picture, positioned so the box shows the kept part.
    let frame = img;
    const crop = (!el.fit || el.fit === 'fill') && el.crop;
    if (crop) {
      const fw = Math.max(0.01, 1 - crop.l - crop.r);
      const fh = Math.max(0.01, 1 - crop.t - crop.b);
      Object.assign(img.style, { position: 'absolute', maxWidth: 'none', width: `${100 / fw}%`, height: `${100 / fh}%`, left: `${(-100 * crop.l) / fw}%`, top: `${(-100 * crop.t) / fh}%` });
      frame = document.createElement('div');
      Object.assign(frame.style, { position: 'relative', width: '100%', height: '100%', overflow: 'hidden' });
      frame.append(img);
    }
    if (el.radius) frame.style.borderRadius = `${el.radius}px`;
    // A picture cropped to a shape (a circle…): clipped to its outline, which also carries the border.
    const clip = el.clip && clipPath(el.clip, el.w, el.h);
    const border = el.stroke && el.stroke.color !== 'none' && el.stroke.width;
    if (clip) frame.style.clipPath = `path('${clip}')`;
    else if (border) {
      frame.style.outline = `${el.stroke.width}px ${el.stroke.dash && el.stroke.dash !== 'solid' ? 'dashed' : 'solid'} ${resolveColor(el.stroke.color, theme)}`;
      frame.style.outlineOffset = `-${el.stroke.width / 2}px`;
    }
    const filters = [...pictureFilters(el.adjust), shadow].filter(Boolean);
    if (filters.length) frame.style.filter = filters.join(' ');
    if (el.flipH || el.flipV) frame.style.transform = `scale(${el.flipH ? -1 : 1}, ${el.flipV ? -1 : 1})`;
    div.append(frame);
    if (clip && border) {
      const svg = svgEl('svg', { class: 'el-svg', overflow: 'visible', width: el.w, height: el.h });
      svg.append(svgEl('path', { d: clip, fill: 'none', ...strokeAttrs(el.stroke, theme) }));
      div.append(svg);
    }
  }

  if (el.type === 'table') div.append(renderTable(el, ctx));
  if (el.type === 'chart') div.append(renderChart(el, ctx));
  if (el.type === 'media') div.append(renderMedia(el, ctx));
  return div;
}

function renderTable(el, ctx) {
  const { theme } = ctx;
  const table = document.createElement('table');
  table.className = 'el-table-grid';
  applyTextStyle(table, { ...el, type: 'text' }, theme);
  const border = el.stroke && el.stroke.color !== 'none' ? `${el.stroke.width || 1}px solid ${resolveColor(el.stroke.color, theme)}` : 'none';
  const colgroup = document.createElement('colgroup');
  for (const w of el.colWidths || []) {
    const col = document.createElement('col');
    col.style.width = `${w * 100}%`;
    colgroup.append(col);
  }
  table.append(colgroup);
  const headerFill = resolveColor(el.fill || 'theme:accent1', theme);
  el.cells.forEach((row, r) => {
    const tr = document.createElement('tr');
    if (el.rowHeights?.[r]) tr.style.height = `${el.rowHeights[r] * 100}%`;
    row.forEach((cell, c) => {
      if (cell.merged) return; // covered by a merged cell
      const td = document.createElement('td');
      if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
      if (cell.colSpan > 1) td.colSpan = cell.colSpan;
      if (cell.valign) td.style.verticalAlign = cell.valign === 'bottom' ? 'bottom' : cell.valign === 'top' ? 'top' : 'middle';
      td.style.border = border;
      td.dataset.r = r;
      td.dataset.c = c;
      if (cell.fill) td.style.background = resolveColor(cell.fill, theme);
      else if (el.header && r === 0) {
        td.style.background = headerFill;
        td.style.color = '#ffffff';
        td.style.fontWeight = '700';
      } else if (el.banded && r % 2 === 0) td.style.background = 'rgba(127,127,127,0.12)';
      const body = document.createElement('div');
      body.className = 'cell-body';
      body.innerHTML = cell.text || '';
      td.append(body);
      tr.append(td);
    });
    table.append(tr);
  });
  return table;
}

// ------------------------------------------------------------------ slides

function decorElements(slide, theme, W, H) {
  if (slide.hideDecor) return [];
  const list = isTitleLayout(slide.layout) ? theme.decor?.title : theme.decor?.content;
  return (list || []).map((d, i) => ({
    id: `decor-${i}`, type: 'shape', shape: d.shape, x: d.x * W, y: d.y * H, w: d.w * W, h: d.h * H, rot: 0,
    fill: d.fill, opacity: d.opacity, stroke: d.stroke ? { color: d.stroke, width: d.strokeWidth || 2 } : null,
  }));
}

/** Theme decorations as plain elements (for export). */
export const themeDecor = (slide, deck) => {
  const { w, h } = slideSize(deck);
  return decorElements(slide, themeById(deck.themeId), w, h);
};

/**
 * Render a whole slide. Options:
 *   editing – show placeholder hints on empty placeholders
 *   live    – videos and sounds can play (the slideshow), instead of showing their poster
 */
export function renderSlide(slide, deck, { editing = false, live = false } = {}) {
  registerDeckFonts(deck);
  const theme = themeById(deck.themeId);
  const { w, h } = slideSize(deck);
  const root = document.createElement('div');
  root.className = `slide${editing ? ' is-editing' : ''}`;
  root.style.width = `${w}px`;
  root.style.height = `${h}px`;
  root.style.background = backgroundCss(slide.background, theme, (id) => mediaUrl(deck, id));
  const ctx = { theme, deck, editing, live };
  const decor = document.createElement('div');
  decor.className = 'slide-decor';
  for (const d of decorElements(slide, theme, w, h)) decor.append(renderElement(d, ctx));
  root.append(decor);
  for (const el of slide.elements) root.append(renderElement(el, ctx));
  return root;
}

/** A slide scaled into a box `width` px wide (thumbnails, previews). */
export function renderThumb(slide, deck, width) {
  const { w, h } = slideSize(deck);
  const scale = width / w;
  const frame = document.createElement('div');
  frame.className = 'slide-thumb';
  frame.style.width = `${width}px`;
  frame.style.height = `${Math.round(h * scale)}px`;
  const s = renderSlide(slide, deck);
  s.style.transform = `scale(${scale})`;
  s.style.transformOrigin = '0 0';
  frame.append(s);
  return frame;
}
