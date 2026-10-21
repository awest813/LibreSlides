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

/** An object URL for a deck picture (cheaper to reuse across thumbnails than its data URL). */
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

function textBox(el, ctx) {
  const box = document.createElement('div');
  box.className = 'el-textbox';
  if (el.type === 'shape') {
    const ins = textInset(el.shape);
    box.style.left = `${ins.x * el.w}px`;
    box.style.right = `${ins.x * el.w}px`;
    box.style.top = `${ins.y * el.h}px`;
    box.style.bottom = `${ins.b * el.h}px`;
  }
  applyTextStyle(box, el, ctx.theme);
  const body = document.createElement('div');
  body.className = 'el-text-body';
  if (el.bullets) body.classList.add('has-bullets');
  body.innerHTML = el.text || '';
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
    const fill = resolveColor(el.fill, theme, 'transparent');
    const stroke = strokeAttrs(el.stroke, theme);
    if (el.type === 'shape' && el.shape === 'custom' && Array.isArray(el.paths)) {
      for (const p of el.paths) {
        svg.append(svgEl('path', { d: customPath(p.cmds, el.w, el.h), fill: p.fill === false ? 'none' : fill, ...(p.stroke === false ? { stroke: 'none' } : stroke) }));
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
    if (el.stroke && el.stroke.color !== 'none' && el.stroke.width) {
      frame.style.outline = `${el.stroke.width}px ${el.stroke.dash && el.stroke.dash !== 'solid' ? 'dashed' : 'solid'} ${resolveColor(el.stroke.color, theme)}`;
      frame.style.outlineOffset = `-${el.stroke.width / 2}px`;
    }
    if (shadow) frame.style.filter = shadow;
    if (el.flipH || el.flipV) frame.style.transform = `scale(${el.flipH ? -1 : 1}, ${el.flipV ? -1 : 1})`;
    div.append(frame);
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
    row.forEach((cell, c) => {
      const td = document.createElement('td');
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
 */
export function renderSlide(slide, deck, { editing = false } = {}) {
  registerDeckFonts(deck);
  const theme = themeById(deck.themeId);
  const { w, h } = slideSize(deck);
  const root = document.createElement('div');
  root.className = `slide${editing ? ' is-editing' : ''}`;
  root.style.width = `${w}px`;
  root.style.height = `${h}px`;
  root.style.background = backgroundCss(slide.background, theme, (id) => mediaUrl(deck, id));
  const ctx = { theme, deck, editing };
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
