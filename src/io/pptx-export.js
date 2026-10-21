import PptxGenJS from 'pptxgenjs';
import JSZip from 'jszip';
import { slideSize, boxOf, htmlToText, ANIM_KINDS, motionPath, FIELD, fieldText } from '../model/deck.js';
import { cellBorder, cellMargin } from '../model/tables.js';
import { isPlainWarp } from '../model/warps.js';
import { themeById, resolveColor, backgroundBaseColor } from '../model/themes.js';
import { textStyle, themeDecor } from '../render/render.js';
import { safeFileName } from '../ui/dom.js';
import { saveFile } from './files.js';
import { addChart, finishCharts } from './pptx-chart.js';
import { addMediaElement, finishMedia } from './pptx-media.js';
import { embedFonts } from './pptx-fonts.js';
import { timingXml, insertTiming } from './pptx-timing.js';
import { writeComments } from './pptx-comments.js';
import { writeSections } from './pptx-sections.js';

/**
 * PowerPoint export with pptxgenjs. Slides become real PowerPoint text boxes,
 * shapes, pictures and tables (editable in PowerPoint, Keynote, Google Slides
 * and LibreOffice). Transitions and skipped slides are added to the XML
 * afterwards, since pptxgenjs doesn't write them.
 */
const IN = 72; // points per inch

/** "#rrggbb", "rgb(…)" → "RRGGBB" (null for transparent). */
export function toHex(css) {
  if (!css || css === 'transparent' || css === 'none') return null;
  let m = /^#([0-9a-f]{3})$/i.exec(css);
  if (m) return m[1].split('').map((c) => c + c).join('').toUpperCase();
  m = /^#([0-9a-f]{6})/i.exec(css);
  if (m) return m[1].toUpperCase();
  m = /rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)(?:[\s,/]+([\d.]+))?/i.exec(css);
  if (m) {
    if (m[4] !== undefined && parseFloat(m[4]) === 0) return null;
    return [m[1], m[2], m[3]].map((n) => (+n).toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  return null;
}

/** A line end (true or a style name) as pptxgenjs's arrow type. */
const arrowType = (end) => (['arrow', 'stealth', 'diamond', 'oval'].includes(end) ? end : 'triangle');

const DASH = { solid: 'solid', dash: 'dash', dot: 'sysDot', dashDot: 'dashDot', longDash: 'lgDash' };

function lineOpts(stroke, theme) {
  if (!stroke || !stroke.width || !stroke.color || stroke.color === 'none') return { type: 'none' };
  return { color: toHex(resolveColor(stroke.color, theme)) || '000000', width: stroke.width, dashType: DASH[stroke.dash] || 'solid' };
}

/**
 * A table cell's margins for pptxgenjs, [top, right, bottom, left]. It reads
 * them as points when the top one is at least 1, else as inches.
 */
function cellMargins(cell) {
  const m = cellMargin(cell);
  const list = [m.t, m.r, m.b, m.l];
  return m.t >= 1 ? list : list.map((v) => v / IN);
}

function fillOpts(fill, theme, opacity = 1) {
  const hex = toHex(resolveColor(fill, theme, 'transparent'));
  if (!hex) return { type: 'none' };
  const f = { color: hex };
  if (opacity < 1) f.transparency = Math.round((1 - opacity) * 100);
  return f;
}

const geom = (b) => ({ x: b.x / IN, y: b.y / IN, w: Math.max(b.w, 0.5) / IN, h: Math.max(b.h, 0.5) / IN });

// ------------------------------------------------------------------ text

/** The value of `prop` in an inline style, as written (parentheses and all), or null. */
function styleValue(css, prop) {
  const m = new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*`, 'i').exec(css);
  if (!m) return null;
  let depth = 0;
  let i = m.index + m[0].length;
  const from = i;
  for (; i < css.length; i++) {
    if (css[i] === '(') depth++;
    else if (css[i] === ')') depth--;
    else if (css[i] === ';' && depth <= 0) break;
  }
  return css.slice(from, i).trim() || null;
}

/** Split a CSS list on its top-level commas. */
const cssList = (v) => v.split(/,(?![^(]*\))/).map((x) => x.trim()).filter(Boolean);

/** A CSS colour and its opacity: { color (hex), alpha }. */
function cssColor(v) {
  const hex = toHex(v);
  if (!hex) return null;
  const a = /rgba\([^)]*[,/]\s*([\d.]+)\s*\)/i.exec(v);
  return { color: hex, alpha: a ? Math.max(0, Math.min(1, +a[1])) : 1 };
}

const COLOR_RE = /#[0-9a-f]{3,8}\b|rgba?\([^)]*\)|\b[a-z]+\b/i;

/**
 * Text effects in a run's inline style (see model/deck.js): an outline
 * { width, color, alpha }, shadows [{ x, y, blur, color, alpha }] and a gradient.
 */
function textEffectsOf(css, s) {
  const stroke = styleValue(css, '-webkit-text-stroke');
  const strokeWidth = styleValue(css, '-webkit-text-stroke-width');
  if (stroke || strokeWidth) {
    const width = parseFloat(stroke || strokeWidth) || 0;
    const color = cssColor((stroke || '').replace(/^[\d.]+px\s*/, '') || styleValue(css, '-webkit-text-stroke-color') || '');
    s.outline = width > 0 && color ? { width, ...color } : null;
  }
  const shadow = styleValue(css, 'text-shadow');
  if (shadow) {
    s.shadows = shadow === 'none' ? null : cssList(shadow).map((one) => {
      const color = cssColor(COLOR_RE.exec(one.replace(/-?[\d.]+px/g, ''))?.[0] || '') || { color: '000000', alpha: 1 };
      const [x = 0, y = 0, blur = 0] = (one.match(/-?[\d.]+px/g) || []).map(parseFloat);
      return { x, y, blur, ...color };
    });
  }
  const image = styleValue(css, 'background-image');
  if (image) s.gradient = /gradient\(/.test(image) && /transparent/.test(styleValue(css, '-webkit-text-fill-color') || '') ? gradientOfCss(image) : null;
  if (/^(transparent|rgba\([^)]*,\s*0\))$/i.test(styleValue(css, 'color') || '')) s.noFill = true;
}

/** A CSS linear or radial gradient as a LibreSlides one (see model/deck.js). */
function gradientOfCss(image) {
  const m = /^(linear|radial)-gradient\((.*)\)$/i.exec(image.trim());
  if (!m) return null;
  const parts = cssList(m[2]);
  const g = { type: m[1].toLowerCase(), stops: [] };
  if (g.type === 'linear') g.angle = /deg$/.test(parts[0]) ? (parts.shift(), parseFloat(m[2])) : 180;
  else {
    const at = /at\s+([\d.]+)%\s+([\d.]+)%/.exec(parts[0]);
    if (at || /circle|ellipse/.test(parts[0])) parts.shift();
    Object.assign(g, { cx: at ? +at[1] / 100 : 0.5, cy: at ? +at[2] / 100 : 0.5 });
  }
  parts.forEach((stop, i) => {
    const color = cssColor(stop.replace(/\s+[\d.]+%$/, ''));
    if (!color) return;
    const pos = /([\d.]+)%$/.exec(stop);
    g.stops.push({ pos: pos ? +pos[1] / 100 : i / Math.max(1, parts.length - 1), color: `#${color.color}`, ...(color.alpha < 1 && { opacity: color.alpha }) });
  });
  return g.stops.length ? g : null;
}

/** Formatting of a run from its element and its ancestors up to the text root. */
function runStyle(node, root, base, links) {
  const s = { ...base };
  const chain = [];
  for (let n = node.parentElement; n && n !== root; n = n.parentElement) chain.unshift(n);
  for (const el of chain) {
    const tag = el.tagName;
    if (tag === 'B' || tag === 'STRONG') s.bold = true;
    if (tag === 'I' || tag === 'EM') s.italic = true;
    if (tag === 'U') s.underline = true;
    if (tag === 'S' || tag === 'STRIKE') s.strike = true;
    if (tag === 'SUB') s.subscript = true;
    if (tag === 'SUP') s.superscript = true;
    if (tag === 'A' && el.getAttribute('href')) s.hyperlink = links ? links(el.getAttribute('href')) : { url: el.getAttribute('href') };
    if (FIELD.test(el.dataset?.field || '')) s.field = el.dataset.field;
    if (tag === 'FONT') {
      if (el.getAttribute('color')) s.color = toHex(el.getAttribute('color')) || s.color;
      if (el.getAttribute('face')) s.fontFace = el.getAttribute('face');
    }
    const st = el.style;
    if (!st) continue;
    if (st.fontWeight) s.bold = st.fontWeight === 'bold' || parseInt(st.fontWeight, 10) >= 600;
    if (st.fontStyle) s.italic = st.fontStyle === 'italic';
    const deco = `${st.textDecoration || ''} ${st.textDecorationLine || ''}`;
    if (/underline/.test(deco)) s.underline = true;
    if (/line-through/.test(deco)) s.strike = true;
    if (st.color && toHex(st.color)) Object.assign(s, { color: toHex(st.color), noFill: false });
    if (st.backgroundColor) s.highlight = toHex(st.backgroundColor) || s.highlight;
    if (st.fontSize && /px|pt$/.test(st.fontSize)) s.fontSize = parseFloat(st.fontSize);
    if (st.fontFamily) s.fontFace = st.fontFamily.split(',')[0].replace(/["']/g, '').trim();
    if (st.letterSpacing && parseFloat(st.letterSpacing)) s.charSpacing = parseFloat(st.letterSpacing);
    if (st.textTransform === 'uppercase') s.caps = 'all';
    if (/small-caps/.test(st.fontVariant || '')) s.caps = 'small';
    textEffectsOf(el.getAttribute('style') || '', s);
  }
  return s;
}

/**
 * What pptxgenjs can't write — capitals, fields, outlines, gradient text,
 * shadows and glows — travels in the run's language as "x-ls-" and hex-coded
 * JSON, and is written into the XML afterwards (see fixRuns).
 */
const EXTRAS_LANG = 'x-ls-';
const hexOf = (str) => [...new TextEncoder().encode(str)].map((b) => b.toString(16).padStart(2, '0')).join('');
const unhex = (hex) => new TextDecoder().decode(new Uint8Array(hex.match(/../g).map((h) => parseInt(h, 16))));

function toRunOptions(s) {
  const o = {};
  if (s.bold) o.bold = true;
  if (s.italic) o.italic = true;
  if (s.underline) o.underline = { style: 'sng' };
  if (s.strike) o.strike = 'sngStrike';
  if (s.color) o.color = s.color;
  if (s.highlight) o.highlight = s.highlight;
  if (s.fontSize) o.fontSize = s.fontSize;
  if (s.fontFace) o.fontFace = s.fontFace;
  if (s.subscript) o.subscript = true;
  if (s.superscript) o.superscript = true;
  if (s.hyperlink) o.hyperlink = s.hyperlink;
  if (s.charSpacing) o.charSpacing = s.charSpacing;
  const extras = {};
  for (const key of ['caps', 'field', 'outline', 'gradient', 'shadows']) if (s[key]) extras[key] = s[key];
  if (s.noFill && !s.gradient) extras.noFill = true;
  if (Object.keys(extras).length) o.lang = EXTRAS_LANG + hexOf(JSON.stringify(extras));
  return o;
}

const emu = (pt) => Math.round(pt * 12700);
const clrXml = ({ color, alpha }) => (alpha < 1 ? `<a:srgbClr val="${color}"><a:alpha val="${Math.round(alpha * 100000)}"/></a:srgbClr>` : `<a:srgbClr val="${color}"/>`);

/** A run's a:rPr with its extras (see EXTRAS_LANG) written in. */
function runWithExtras(hex, attrs, body) {
  let extras;
  try {
    extras = JSON.parse(unhex(hex));
  } catch {
    return `<a:rPr lang="en-US"${attrs}>${body}</a:rPr>`;
  }
  if (extras.caps === 'all' || extras.caps === 'small') attrs += ` cap="${extras.caps}"`;
  if (extras.field) attrs += ` x-ls-fld="${extras.field}"`;
  const [, solid = '', rest] = /^(<a:solidFill>[\s\S]*?<\/a:solidFill>)?([\s\S]*)$/.exec(body);
  const o = extras.outline;
  const ln = o ? `<a:ln w="${emu(o.width)}"><a:solidFill>${clrXml(o)}</a:solidFill></a:ln>` : '';
  const fill = extras.gradient ? gradFillXml(extras.gradient) : extras.noFill ? '<a:noFill/>' : solid;
  // Shadows with no offset are glows.
  const glow = extras.shadows?.find((sh) => !sh.x && !sh.y);
  const shadow = extras.shadows?.find((sh) => sh.x || sh.y);
  let effects = '';
  if (glow) effects += `<a:glow rad="${emu(glow.blur)}">${clrXml(glow)}</a:glow>`;
  if (shadow) {
    const dist = Math.hypot(shadow.x, shadow.y);
    const dir = Math.round(((((Math.atan2(shadow.y, shadow.x) * 180) / Math.PI) % 360) + 360) % 360 * 60000);
    effects += `<a:outerShdw blurRad="${emu(shadow.blur)}" dist="${emu(dist)}" dir="${dir}" algn="ctr" rotWithShape="0">${clrXml(shadow)}</a:outerShdw>`;
  }
  return `<a:rPr lang="en-US"${attrs}>${ln}${fill}${effects ? `<a:effectLst>${effects}</a:effectLst>` : ''}${rest}</a:rPr>`;
}

let fieldCount = 0;

/** Text with its fields showing slide number `n` and today's date, as PowerPoint shows them until it updates them. */
const withFields = (html, n) => (html || '').replace(/(<span data-field="(\w+)">)[^<]*(<\/span>)/g, (m, open, type, close) => (FIELD.test(type) ? open + fieldText(type, n) + close : m));

/** Write run extras (see EXTRAS_LANG) into a slide's XML, and make field runs fields. */
function fixRuns(xml) {
  xml = xml.replace(/<a:rPr lang="x-ls-([0-9a-f]+)"(?: altLang="en-US")?([^>]*)>([\s\S]*?)<\/a:rPr>/g, (m, hex, attrs, body) => runWithExtras(hex, attrs, body));
  return xml.replace(/<a:r>(<a:rPr\b[^>]*?) x-ls-fld="(\w+)"([\s\S]*?)<\/a:r>/g, (m, head, type, tail) => {
    const id = `{5A3B1C2D-0000-4000-8000-${(++fieldCount).toString(16).toUpperCase().padStart(12, '0')}}`;
    return `<a:fld id="${id}" type="${type}">${head}${tail}</a:fld>`;
  });
}

// CSS numbering styles as PowerPoint's.
const NUMBER_STYLES = { 'lower-alpha': 'alphaLcPeriod', 'upper-alpha': 'alphaUcPeriod', 'lower-roman': 'romanLcPeriod', 'upper-roman': 'romanUcPeriod' };

/** A list's bullet for pptxgenjs: numbered (with its style and first number), or a bullet character of its own. */
function listBullet(list) {
  const type = list.style?.listStyleType || '';
  if (list.tagName === 'OL') {
    const b = { type: 'number' };
    if (NUMBER_STYLES[type]) b.style = NUMBER_STYLES[type];
    if (+list.getAttribute('start') > 1) b.numberStartAt = +list.getAttribute('start');
    return b;
  }
  // A string bullet ('– '): its first character.
  const ch = /^["'](.)/u.exec(type)?.[1];
  return ch && ch.codePointAt(0) <= 0xffff ? { characterCode: ch.codePointAt(0).toString(16).padStart(4, '0').toUpperCase() } : true;
}

/** Paragraphs of an HTML fragment as pptxgenjs text runs. */
export function htmlToRuns(html, base = {}, links = null) {
  const doc = new DOMParser().parseFromString(`<body>${html || ''}</body>`, 'text/html');
  const root = doc.body;
  // Loose inline content at the top level forms paragraphs of its own.
  let loose = null;
  for (const node of [...root.childNodes]) {
    if (node.nodeType === 1 && /^(P|DIV|H\d|UL|OL)$/.test(node.tagName)) {
      loose = null;
      continue;
    }
    if (!loose) {
      loose = doc.createElement('p');
      node.before(loose);
    }
    loose.append(node);
  }

  const paragraphs = [];
  const collect = (block, para) => {
    const runs = [];
    // A <br> is a line break inside the paragraph: the next run starts on a new line (<a:br/>).
    let breaks = 0;
    const walk = (n) => {
      for (const c of n.childNodes) {
        if (c.nodeType === 3) {
          const text = c.textContent.replace(/​/g, '');
          if (!text) continue;
          const options = toRunOptions(runStyle(c, root, base, links));
          if (breaks && !runs.length) runs.push({ text: '', options }); // the paragraph starts with an empty line
          for (; breaks > 1; breaks--) runs.push({ text: '', options: { ...options, softBreakBefore: true } });
          runs.push({ text, options: breaks ? { ...options, softBreakBefore: true } : options });
          breaks = 0;
        } else if (c.nodeName === 'BR') breaks++;
        else if (c.nodeName !== 'UL' && c.nodeName !== 'OL') walk(c);
      }
    };
    walk(block);
    // Trailing <br>s only keep an empty line open in the editor.
    paragraphs.push({ runs, para });
  };
  const blockAlign = (el) => {
    const a = el.style?.textAlign;
    return ['left', 'center', 'right', 'justify'].includes(a) ? a : null;
  };
  const visitList = (list, level) => {
    const bullet = listBullet(list);
    for (const li of list.children) {
      if (li.tagName !== 'LI') continue;
      collect(li, { bullet, indentLevel: level, align: blockAlign(li) });
      for (const sub of li.children) if (sub.tagName === 'UL' || sub.tagName === 'OL') visitList(sub, level + 1);
    }
  };
  for (const node of root.children) {
    if (node.tagName === 'UL' || node.tagName === 'OL') visitList(node, 0);
    else collect(node, { align: blockAlign(node) });
  }

  const out = [];
  paragraphs.forEach(({ runs, para }, i) => {
    const pOpts = {};
    if (para.bullet) pOpts.bullet = para.bullet;
    if (para.indentLevel) pOpts.indentLevel = para.indentLevel;
    if (para.align) pOpts.align = para.align;
    if (!runs.length) runs.push({ text: '', options: toRunOptions(base) });
    runs.forEach((r, k) => {
      const options = { ...r.options, ...pOpts };
      if (k === runs.length - 1 && i < paragraphs.length - 1) options.breakLine = true;
      out.push({ text: r.text, options });
    });
  });
  return out;
}

// ------------------------------------------------------------------ pictures

const blobToDataUrl = (blob) => new Promise((resolve, reject) => {
  const r = new FileReader();
  r.onload = () => resolve(r.result);
  r.onerror = () => reject(r.error);
  r.readAsDataURL(blob);
});

/** PowerPoint can't read SVG reliably, so vector pictures are exported as PNG. */
async function rasterize(dataUrl, w, h) {
  const img = new Image();
  img.src = dataUrl;
  await img.decode();
  const k = Math.min(4, 2000 / Math.max(w, h, 1)) || 2;
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(w * k));
  canvas.height = Math.max(1, Math.round(h * k));
  canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
  return canvas.toDataURL('image/png');
}

async function exportableImage(dataUrl, w = 400, h = 300) {
  if (!dataUrl) return null;
  if (/^data:image\/svg/i.test(dataUrl) || !/;base64,/.test(dataUrl)) {
    try {
      return await rasterize(dataUrl, w, h);
    } catch {
      return null;
    }
  }
  if (/^data:image\/(webp|avif|bmp)/i.test(dataUrl)) {
    try {
      return await blobToDataUrl(await (await fetch(await rasterize(dataUrl, w, h))).blob());
    } catch {
      return null;
    }
  }
  return dataUrl;
}

const dataForPptx = (url) => url.replace(/^data:/, '');

/** A picture's pixel size, or null where it can't be decoded (tests, broken data). */
async function imageSize(dataUrl) {
  try {
    const img = new Image();
    img.src = dataUrl;
    await img.decode();
    return { w: img.naturalWidth, h: img.naturalHeight };
  } catch {
    return null;
  }
}

// ------------------------------------------------------------------ slides

async function addElement(pptx, s, el, deck, theme, marks) {
  const opacity = el.opacity ?? 1;
  const shadow = el.shadow ? { type: 'outer', blur: 8, offset: 4, angle: 90, color: '000000', opacity: 0.3 } : undefined;
  // Animated elements, media that plays by itself and group members are named so the
  // slide's timeline and groups can find them.
  const timed = ANIM_KINDS.some((kind) => el[kind]?.effect && el[kind].effect !== 'none') || (el.type === 'media' && (el.autoplay || el.loop));
  const grouped = el.groupId && marks.groupSizes.get(el.groupId) > 1;
  const objectName = timed || grouped ? `LibreSlides shape ${++marks.count}` : undefined;
  if (timed) marks.timeline.push({ name: objectName, el });
  if (grouped) marks.groups.push({ name: objectName, groupId: el.groupId, box: boxOf(el) });
  if (el.type === 'line') {
    let { x1, y1, x2, y2, arrowStart, arrowEnd } = el;
    if (x1 > x2) {
      [x1, x2, y1, y2] = [x2, x1, y2, y1];
      [arrowStart, arrowEnd] = [arrowEnd, arrowStart];
    }
    const line = lineOpts({ color: 'theme:text', width: 2, ...el.stroke, color: el.stroke?.color === 'none' ? 'theme:text' : el.stroke?.color }, theme);
    if (arrowStart) line.beginArrowType = arrowType(arrowStart);
    if (arrowEnd) line.endArrowType = arrowType(arrowEnd);
    s.addShape('line', { ...geom({ x: x1, y: Math.min(y1, y2), w: x2 - x1, h: Math.abs(y2 - y1) }), flipV: y2 < y1, line, objectName, hyperlink: el.link ? marks.links(el.link) || undefined : undefined });
    return;
  }
  const g = geom(boxOf(el));
  const hyperlink = el.link ? marks.links(el.link) || undefined : undefined;
  const common = { ...g, rotate: el.rot || 0, flipH: !!el.flipH, flipV: !!el.flipV, shadow, objectName, hyperlink };

  if (el.type === 'image') {
    const data = await exportableImage(deck.media?.[el.media], el.w, el.h);
    if (!data) return;
    const opts = { ...common, data: dataForPptx(data), altText: el.alt || '' };
    const crop = (!el.fit || el.fit === 'fill') && el.crop;
    if (crop) {
      // The whole picture at the size it's shown, and the part of it the box shows.
      opts.w = g.w / Math.max(0.01, 1 - crop.l - crop.r);
      opts.h = g.h / Math.max(0.01, 1 - crop.t - crop.b);
      opts.sizing = { type: 'crop', x: crop.l * opts.w, y: crop.t * opts.h, w: g.w, h: g.h };
    } else if (el.fit === 'cover' || el.fit === 'contain') {
      // pptxgenjs crops or fits using the picture's own proportions.
      const natural = el.natural?.w ? el.natural : await imageSize(data);
      if (natural?.w && natural?.h) {
        opts.w = g.w;
        opts.h = (g.w * natural.h) / natural.w;
        opts.sizing = { type: el.fit, w: g.w, h: g.h };
      }
    }
    if (el.radius) opts.rounding = el.radius >= Math.min(el.w, el.h) / 2;
    if (opacity < 1) opts.transparency = Math.round((1 - opacity) * 100);
    // Cropped to a shape, or colour adjustments: added to the picture's XML afterwards.
    if (el.clip || el.adjust) {
      opts.objectName ||= `LibreSlides shape ${++marks.count}`;
      marks.pictureLooks.set(opts.objectName, { clip: el.clip, adjust: el.adjust });
    }
    s.addImage(opts);
    return;
  }

  if (el.type === 'chart') {
    addChart(pptx, s, el, theme, g, objectName);
    return;
  }
  if (el.type === 'media') {
    await addMediaElement(pptx, s, el, deck, g, objectName);
    return;
  }

  if (el.type === 'table') {
    const t = textStyle({ ...el, type: 'text' }, theme);
    const base = { fontFace: t.font, fontSize: t.size, color: toHex(resolveColor(t.color, theme)) || '000000' };
    const header = toHex(resolveColor(el.fill || 'theme:accent1', theme));
    const cols = el.cells[0].length;
    const dashes = []; // borders pptxgenjs can only write solid or dashed: { index (of the cell in the grid), side, dash }
    // Cells covered by a merged one are left out; pptxgenjs works out the rest from the spans.
    const rows = el.cells.map((row, r) => row.map((cell, c) => {
      if (cell.merged) return null;
      const options = {};
      if (cell.rowSpan > 1) options.rowspan = cell.rowSpan;
      if (cell.colSpan > 1) options.colspan = cell.colSpan;
      if (cell.valign) options.valign = cell.valign;
      if (cell.fill) options.fill = { color: toHex(resolveColor(cell.fill, theme)) };
      else if (el.header && r === 0 && header) {
        options.fill = { color: header };
        options.color = 'FFFFFF';
        options.bold = true;
      } else if (el.banded && r % 2 === 0) options.fill = { color: 'EFEFEF' };
      // Each side's border, pptxgenjs's way round (top, right, bottom, left); dotted and other dashes are set afterwards.
      options.border = ['t', 'r', 'b', 'l'].map((side) => {
        const line = lineOpts(cellBorder(el, cell, side), theme);
        if (line.type === 'none') return { type: 'none' };
        if (line.dashType !== 'solid') dashes.push({ index: r * cols + c, side, dash: line.dashType });
        return { type: line.dashType === 'solid' ? 'solid' : 'dash', pt: line.width, color: line.color };
      });
      if (cell.margin) options.margin = cellMargins(cell);
      const runs = htmlToRuns(withFields(cell.text, marks.number), {}, marks.links);
      return { text: runs.length ? runs : '', options };
    }).filter(Boolean));
    const name = objectName || (dashes.length ? `LibreSlides shape ${++marks.count}` : undefined);
    if (dashes.length) marks.tableDashes.set(name, dashes);
    s.addTable(rows, {
      x: g.x, y: g.y, w: g.w, colW: (el.colWidths || []).map((f) => f * g.w),
      rowH: el.rowHeights?.length === el.cells.length ? el.rowHeights.map((f) => f * g.h) : g.h / el.cells.length,
      ...base, valign: 'middle', margin: cellMargins({}), objectName: name,
    });
    return;
  }

  // Text boxes and shapes (with or without text).
  const t = textStyle(el, theme);
  let shape = el.type === 'shape' ? el.shape : 'rect';
  const hasText = !!htmlToText(el.text || '');
  const opts = {
    ...common,
    shape,
    fill: fillOpts(el.fill, theme, opacity),
    line: lineOpts(el.stroke, theme),
  };
  // What pptxgenjs can't write is fixed in its XML afterwards (see postProcess), found by the shape's name.
  const mark = () => (opts.objectName ||= `LibreSlides shape ${++marks.count}`);
  const picture = el.fillImage && (await exportableImage(deck.media?.[el.fillImage.media], el.w, el.h));
  if (picture) {
    if (opts.fill.type === 'none') opts.fill = { color: 'FFFFFF' };
    marks.pictures.set(mark(), { data: picture, crop: el.fillImage.crop, opacity });
  } else if (el.gradient && opts.fill.type !== 'none') marks.gradients.set(mark(), gradFillXml(el.gradient, theme, opacity));
  let extras = [];
  if (shape === 'custom' && el.preset?.name) {
    // A PowerPoint preset goes back as that preset (pptxgenjs writes any preset name), with its adjust values.
    shape = el.preset.name;
    opts.shape = shape;
    if (el.preset.adj && Object.keys(el.preset.adj).length) marks.presets.set(mark(), el.preset.adj);
    if (el.arrowStart) opts.line.beginArrowType = arrowType(el.arrowStart);
    if (el.arrowEnd) opts.line.endArrowType = arrowType(el.arrowEnd);
  } else if (shape === 'custom') {
    // A freeform: one PowerPoint shape per kind of path (filled or not, outlined or not);
    // the first carries the text.
    const kinds = new Map();
    for (const p of el.paths || []) {
      const key = `${p.fill !== false}${p.stroke !== false}`;
      if (!kinds.has(key)) kinds.set(key, []);
      kinds.get(key).push(p);
    }
    const parts = [...kinds.values()].map((paths) => ({
      points: customPoints(paths, g),
      fill: paths[0].fill === false ? { type: 'none' } : opts.fill,
      line: paths[0].stroke === false ? { type: 'none' } : opts.line,
    }));
    if (!parts.length) parts.push({ points: customPoints([{ cmds: [['M', 0, 0], ['L', 1, 0], ['L', 1, 1], ['L', 0, 1], ['Z']] }], g), fill: opts.fill, line: opts.line });
    shape = 'custGeom';
    Object.assign(opts, { shape, ...parts[0] });
    extras = parts.slice(1);
  }
  // Drawn over the main shape, in the order they were in.
  const addExtras = () => { for (const extra of extras) s.addShape(shape, { ...common, objectName: undefined, ...extra }); };
  // Titles, subtitles and body text are PowerPoint placeholders (named, then marked as such afterwards);
  // an empty one stays, so PowerPoint shows "Click to add title" there too.
  if (el.ph && el.type === 'text') marks.placeholders.set(mark(), el.ph);
  if (!hasText && !(el.ph && el.type === 'text')) {
    if (el.type === 'shape' || opts.fill.type !== 'none' || opts.line.type !== 'none') s.addShape(shape, opts);
    addExtras();
    return;
  }
  const base = {
    fontFace: t.font,
    fontSize: t.size,
    color: toHex(resolveColor(t.color, theme)) || '000000',
    bold: t.bold || undefined,
    italic: t.italic || undefined,
  };
  Object.assign(opts, {
    ...base,
    underline: t.underline ? { style: 'sng' } : undefined,
    strike: t.strike ? 'sngStrike' : undefined,
    align: t.align,
    valign: t.valign,
    lineSpacingMultiple: Math.round((t.lineHeight / 1.2) * 100) / 100,
    paraSpaceAfter: t.spacing ? Math.round(t.spacing * t.size) : 0,
    fit: 'none',
    isTextBox: el.type === 'text',
  });
  // Text box layout (see render.js textLayout).
  const box = el.style || {};
  if (box.inset) opts.margin = [box.inset.l, box.inset.r, box.inset.b, box.inset.t];
  if (box.wrap === false) opts.wrap = false;
  if (box.vertical) opts.vert = box.vertical;
  if (box.columns > 1) marks.columns.set(mark(), { n: box.columns, gap: box.columnGap });
  if (!isPlainWarp(box.warp) && hasText) marks.warps.set(mark(), box.warp);
  // Text with no fill of its own (hollow, outlined WordArt) has a transparent colour.
  const runBase = resolveColor(t.color, theme) === 'transparent' ? { noFill: true } : {};
  s.addText(hasText ? htmlToRuns(withFields(el.text, marks.number), runBase, marks.links) : [{ text: '', options: {} }], opts);
  addExtras();
}

/** Freeform paths (fractions of the box) as pptxgenjs custom geometry points (inches from the box's corner). */
function customPoints(paths, g) {
  const pts = [];
  for (const { cmds } of paths) {
    for (const [op, ...n] of cmds) {
      const x = (i) => n[i] * g.w;
      const y = (i) => n[i] * g.h;
      if (op === 'M') pts.push({ x: x(0), y: y(1), moveTo: true });
      else if (op === 'L') pts.push({ x: x(0), y: y(1) });
      else if (op === 'C') pts.push({ x: x(4), y: y(5), curve: { type: 'cubic', x1: x(0), y1: y(1), x2: x(2), y2: y(3) } });
      else if (op === 'Q') pts.push({ x: x(2), y: y(3), curve: { type: 'quadratic', x1: x(0), y1: y(1) } });
      else if (op === 'Z') pts.push({ close: true });
    }
  }
  return pts;
}

const TRANSITION_XML = {
  fade: '<p:fade/>',
  slide: '<p:cover dir="l"/>',
  push: '<p:push dir="l"/>',
  cover: '<p:cover dir="u"/>',
  zoom: '<p:zoom/>',
  flip: '<p:fade thruBlk="1"/>',
};

/**
 * pptxgenjs writes paragraph properties before every run of a paragraph, but
 * PowerPoint only allows them first (it offers to "repair" the file
 * otherwise). Keep the first <a:pPr> of each paragraph and drop the rest.
 */
export function fixParagraphs(slideXml) {
  slideXml = fixRuns(slideXml);
  return slideXml.replace(/<a:p>([\s\S]*?)<\/a:p>/g, (para, inner) => {
    let first = true;
    const body = inner.replace(/<a:pPr\b[^>]*\/>|<a:pPr\b[^>]*>[\s\S]*?<\/a:pPr>/g, (pPr, at) => {
      const keep = first && at === 0;
      first = false;
      return keep ? pPr : '';
    });
    return `<a:p>${body}</a:p>`;
  });
}

/** A LibreSlides gradient (see model/deck.js) as DrawingML, faded by `opacity`. */
export function gradFillXml(g, theme, opacity = 1) {
  let stops = g.stops || [];
  // PowerPoint runs path gradients from the edge (first stop) in to the centre.
  if (g.type === 'radial') stops = stops.map((s) => ({ ...s, pos: 1 - s.pos })).reverse();
  const gs = stops.map((s) => {
    const alpha = (s.opacity ?? 1) * opacity;
    const hex = toHex(resolveColor(s.color, theme)) || '000000';
    return `<a:gs pos="${Math.round(Math.max(0, Math.min(1, s.pos)) * 100000)}"><a:srgbClr val="${hex}">${alpha < 1 ? `<a:alpha val="${Math.round(alpha * 100000)}"/>` : ''}</a:srgbClr></a:gs>`;
  }).join('');
  let dir;
  if (g.type === 'radial') {
    const pct = (v) => Math.round(Math.max(0, Math.min(1, v)) * 100000);
    const cx = g.cx ?? 0.5;
    const cy = g.cy ?? 0.5;
    dir = `<a:path path="circle"><a:fillToRect l="${pct(cx)}" t="${pct(cy)}" r="${pct(1 - cx)}" b="${pct(1 - cy)}"/></a:path>`;
  } else dir = `<a:lin ang="${Math.round(((((g.angle ?? 180) - 90) % 360) + 360) % 360) * 60000}" scaled="0"/>`;
  return `<a:gradFill rotWithShape="1"><a:gsLst>${gs}</a:gsLst>${dir}</a:gradFill>`;
}

/** Swap the solid fill of the shape called `name` for `fillXml`. */
function replaceShapeFill(xml, name, fillXml) {
  const at = xml.indexOf(`name="${name}"`);
  if (at < 0) return xml;
  const start = xml.lastIndexOf('<p:sp>', at);
  const end = xml.indexOf('</p:sp>', at);
  const spPr = xml.slice(start, end).match(/<p:spPr>[\s\S]*?<\/p:spPr>/);
  if (start < 0 || end < 0 || !spPr) return xml;
  // The shape's fill comes before its outline (whose own fill is inside <a:ln>).
  const fixed = spPr[0].replace(/^([\s\S]*?)<a:solidFill>[\s\S]*?<\/a:solidFill>/, (whole, before) => (before.includes('<a:ln') ? whole : `${before}${fillXml}`));
  return xml.slice(0, start) + xml.slice(start, end).replace(spPr[0], fixed) + xml.slice(end);
}

/** A text warp (see model/warps.js) as an a:prstTxWarp. */
export function warpXml(warp) {
  const gds = Object.entries(warp.adjust || {}).filter(([k, v]) => /^adj\d*$/.test(k) && Number.isFinite(v)).map(([k, v]) => `<a:gd name="${k}" fmla="val ${Math.round(v)}"/>`).join('');
  return `<a:prstTxWarp prst="${warp.preset}">${gds ? `<a:avLst>${gds}</a:avLst>` : '<a:avLst/>'}</a:prstTxWarp>`;
}

/** Turn the solid fills of the shapes named for gradients (and a gradient background) into real gradients. */
export function applyGradients(xml, marks) {
  if (!marks) return xml;
  // Text in columns.
  for (const [name, { n, gap }] of marks.columns || []) {
    const at = xml.indexOf(`name="${name}"`);
    const end = xml.indexOf('</p:sp>', at);
    if (at < 0 || end < 0) continue;
    const block = xml.slice(at, end).replace('<a:bodyPr', `<a:bodyPr numCol="${Math.round(n)}"${gap != null ? ` spcCol="${Math.round(gap * 12700)}"` : ''}`);
    xml = xml.slice(0, at) + block + xml.slice(end);
  }
  for (const [name, grad] of marks.gradients) xml = replaceShapeFill(xml, name, grad);
  // Preset shapes' adjust values.
  for (const [name, adj] of marks.presets || []) {
    const at = xml.indexOf(`name="${name}"`);
    const end = xml.indexOf('</p:sp>', at);
    if (at < 0 || end < 0) continue;
    const gds = Object.entries(adj).filter(([k, v]) => /^adj\d*$/.test(k) && Number.isFinite(v)).map(([k, v]) => `<a:gd name="${k}" fmla="val ${Math.round(v)}"/>`).join('');
    const block = xml.slice(at, end).replace(/<a:avLst\s*\/>|<a:avLst>[\s\S]*?<\/a:avLst>/, `<a:avLst>${gds}</a:avLst>`);
    xml = xml.slice(0, at) + block + xml.slice(end);
  }
  // WordArt text warps: first in the text box's settings, before its autofit.
  for (const [name, warp] of marks.warps || []) {
    const at = xml.indexOf(`name="${name}"`);
    const end = xml.indexOf('</p:sp>', at);
    if (at < 0 || end < 0) continue;
    const block = xml.slice(at, end).replace(/<a:bodyPr\b([^>]*?)(\/?)>/, (tag, attrs, empty) => {
      const stretch = warp.stretch && !/fromWordArt=/.test(attrs) ? ' fromWordArt="1"' : '';
      return `<a:bodyPr${attrs}${stretch}>${warpXml(warp)}${empty ? '</a:bodyPr>' : ''}`;
    });
    xml = xml.slice(0, at) + block + xml.slice(end);
  }
  if (marks.background) xml = xml.replace(/<p:bg>[\s\S]*?<\/p:bg>/, `<p:bg><p:bgPr>${marks.background}<a:effectLst/></p:bgPr></p:bg>`);
  return xml;
}

/** Put the pictures that fill shapes into the package, and point those shapes' fills at them. */
async function applyPictureFills(zip, slideNo, xml, marks) {
  if (!marks?.pictures.size) return xml;
  const relsPath = `ppt/slides/_rels/slide${slideNo}.xml.rels`;
  let rels = await zip.file(relsPath).async('string');
  let types = await zip.file('[Content_Types].xml').async('string');
  let n = 0;
  for (const [name, { data, crop, opacity }] of marks.pictures) {
    const m = /^data:image\/(png|jpeg|gif);base64,(.*)$/s.exec(data);
    if (!m) continue;
    const ext = m[1] === 'jpeg' ? 'jpg' : m[1];
    const id = `rIdLsFill${++n}`;
    zip.file(`ppt/media/ls-fill-${slideNo}-${n}.${ext}`, m[2], { base64: true });
    rels = rels.replace('</Relationships>', `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/ls-fill-${slideNo}-${n}.${ext}"/></Relationships>`);
    if (!new RegExp(`Extension="${ext}"`, 'i').test(types)) types = types.replace('<Default ', `<Default Extension="${ext}" ContentType="image/${m[1]}"/><Default `);
    const pct = (v) => Math.round((v || 0) * 100000);
    const src = crop ? `<a:srcRect l="${pct(crop.l)}" t="${pct(crop.t)}" r="${pct(crop.r)}" b="${pct(crop.b)}"/>` : '';
    const alpha = opacity < 1 ? `<a:alphaModFix amt="${pct(opacity)}"/>` : '';
    xml = replaceShapeFill(xml, name, `<a:blipFill rotWithShape="1"><a:blip r:embed="${id}">${alpha}</a:blip>${src}<a:stretch><a:fillRect/></a:stretch></a:blipFill>`);
  }
  zip.file(relsPath, rels);
  zip.file('[Content_Types].xml', types);
  return xml;
}

/** Where the shape called `name` is in a slide's XML: the closest <p:sp>, <p:pic> or <p:graphicFrame> opened before its name. */
function shapeBlock(xml, name) {
  const at = xml.indexOf(`name="${name}"`);
  if (at < 0) return null;
  const [tag, start] = ['<p:sp>', '<p:pic>', '<p:graphicFrame>'].map((t) => [t, xml.lastIndexOf(t, at)]).sort((a, b) => b[1] - a[1])[0];
  const close = `</${tag.slice(1)}`;
  const end = xml.indexOf(close, at);
  return start < 0 || end < 0 ? null : { start, end: end + close.length };
}

/** Change the XML of the shape called `name`. */
function editShape(xml, name, change) {
  const b = shapeBlock(xml, name);
  return b ? xml.slice(0, b.start) + change(xml.slice(b.start, b.end)) + xml.slice(b.end) : xml;
}

/** Mark title, subtitle and body text boxes as PowerPoint placeholders (ids from 1 for those that need one). */
function applyPlaceholders(xml, placeholders) {
  let idx = 0;
  for (const [name, role] of placeholders || []) {
    const ph = role === 'title' ? '<p:ph type="title"/>' : role === 'subtitle' ? `<p:ph type="subTitle" idx="${++idx}"/>` : `<p:ph type="body" idx="${++idx}"/>`;
    xml = editShape(xml, name, (sp) => sp.replace(/<p:cNvSpPr txBox="1"\/>/, '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>')
      .replace(/<p:nvPr\/>|<p:nvPr><\/p:nvPr>/, `<p:nvPr>${ph}</p:nvPr>`));
  }
  return xml;
}

/** A pictures' colour adjustments (see render.js pictureFilters) and the shape it's cropped to. */
function applyPictureLooks(xml, looks) {
  for (const [name, { clip, adjust }] of looks || []) {
    xml = editShape(xml, name, (pic) => {
      const a = adjust || {};
      const pct = (v) => Math.round(v * 100000);
      const effects = (a.grayscale ? '<a:grayscl/>' : '') + (a.brightness || a.contrast ? `<a:lum${a.brightness ? ` bright="${pct(a.brightness)}"` : ''}${a.contrast ? ` contrast="${pct(a.contrast)}"` : ''}/>` : '');
      if (effects) pic = pic.replace(/<a:blip ([^>]*?)\/>|<a:blip ([^>]*)>/, (m, selfClosing, open) => (selfClosing != null ? `<a:blip ${selfClosing}>${effects}</a:blip>` : `<a:blip ${open}>${effects}`));
      const shape = clip?.preset?.name || (clip?.shape !== 'custom' && clip?.shape);
      if (shape) {
        const adj = clip.preset?.adj ? Object.entries(clip.preset.adj).map(([k, v]) => `<a:gd name="${k}" fmla="val ${Math.round(v)}"/>`).join('') : '';
        pic = pic.replace(/<a:prstGeom prst="[^"]*">[\s\S]*?<\/a:prstGeom>|<a:prstGeom prst="[^"]*"\/>/, `<a:prstGeom prst="${shape}"><a:avLst>${adj}</a:avLst></a:prstGeom>`);
      }
      return pic;
    });
  }
  return xml;
}

/**
 * pptxgenjs has no groups: gather each group's shapes (found by name, in their
 * stacking order) into a <p:grpSp> where the first of them was.
 */
export function applyGroups(xml, members) {
  const byGroup = new Map();
  for (const m of members || []) byGroup.set(m.groupId, [...(byGroup.get(m.groupId) || []), m]);
  let next = Math.max(0, ...(xml.match(/<p:cNvPr id="\d+"/g) || []).map((s) => +s.match(/\d+/)[0])) + 1;
  for (const list of byGroup.values()) {
    const blocks = [];
    for (const { name } of list) {
      const b = shapeBlock(xml, name);
      if (b) blocks.push({ ...b, xml: xml.slice(b.start, b.end) });
    }
    if (blocks.length < 2) continue;
    blocks.sort((a, b) => a.start - b.start);
    const emu = (v) => Math.round(v * 12700);
    const boxes = list.map((m) => m.box);
    const x = Math.min(...boxes.map((b) => b.x));
    const y = Math.min(...boxes.map((b) => b.y));
    const w = Math.max(...boxes.map((b) => b.x + b.w)) - x;
    const h = Math.max(...boxes.map((b) => b.y + b.h)) - y;
    const frame = `<a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/><a:chOff x="${emu(x)}" y="${emu(y)}"/><a:chExt cx="${emu(w)}" cy="${emu(h)}"/>`;
    const id = next++;
    const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="Group ${id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm>${frame}</a:xfrm></p:grpSpPr>${blocks.map((b) => b.xml).join('')}</p:grpSp>`;
    // Take the shapes out (last first, so earlier positions hold), then put the group where the first was.
    for (const b of [...blocks].reverse()) xml = xml.slice(0, b.start) + xml.slice(b.end);
    xml = xml.slice(0, blocks[0].start) + group + xml.slice(blocks[0].start);
  }
  return xml;
}

const BORDER_TAG = { t: 'lnT', r: 'lnR', b: 'lnB', l: 'lnL' };

/** Table borders' dashes, which pptxgenjs writes as solid or "sysDash" only (cells in grid order, merged ones too). */
export function applyTableDashes(xml, tables) {
  for (const [name, dashes] of tables || []) {
    xml = editShape(xml, name, (frame) => {
      const cells = frame.split(/(?=<a:tc[ >])/);
      for (const { index, side, dash } of dashes) {
        const i = index + 1; // before the first cell: the frame and grid
        if (!cells[i]) continue;
        const tag = BORDER_TAG[side];
        cells[i] = cells[i].replace(new RegExp(`(<a:${tag} [^>]*>[\\s\\S]*?<a:prstDash val=")[^"]*(")`), `$1${dash}$2`);
      }
      return cells.join('');
    });
  }
  return xml;
}

/** Give shapes that pptxgenjs numbered the same (tables, media) ids of their own, so animations find the right one. */
function uniqueShapeIds(xml) {
  const ids = (xml.match(/<p:cNvPr id="\d+"/g) || []).map((m) => +m.match(/\d+/)[0]);
  let next = Math.max(0, ...ids) + 1;
  const seen = new Set();
  return xml.replace(/<p:cNvPr id="(\d+)"/g, (m, id) => {
    if (!seen.has(id)) {
      seen.add(id);
      return m;
    }
    return `<p:cNvPr id="${next++}"`;
  });
}

/** The animations and media of a slide's timeline (see pptx-timing.js), found in its XML by name. */
function timelineOf(xml, timeline) {
  const anims = [];
  const media = [];
  for (const { name, el } of timeline) {
    const spid = new RegExp(`<p:cNvPr id="(\\d+)" name="${name}"`).exec(xml)?.[1];
    if (!spid) continue;
    if (el.type === 'media') media.push({ spid, kind: el.kind === 'audio' ? 'audio' : 'video', autoplay: !!el.autoplay, loop: !!el.loop, duration: el.duration });
    for (const kind of ANIM_KINDS) {
      const a = el[kind];
      if (!a?.effect || a.effect === 'none') continue;
      anims.push({ spid, kind, order: a.order || 0, effect: a.effect, duration: a.duration ?? 0.5, start: a.start || 'click', hasText: !!htmlToText(el.text || ''), ...(kind === 'motion' && { path: motionPath(a) }) });
    }
  }
  anims.sort((a, b) => a.order - b.order);
  return { anims, media };
}

/** Tidy the paragraphs, and add transitions and "skip slide" to the generated XML. */
async function postProcess(buffer, deck, allMarks) {
  const zip = await JSZip.loadAsync(buffer);
  await finishMedia(zip, deck);
  await finishCharts(zip, deck);
  for (const [i, slide] of deck.slides.entries()) {
    const path = `ppt/slides/slide${i + 1}.xml`;
    const file = zip.file(path);
    if (!file) continue;
    let xml = applyPictureLooks(applyGradients(fixParagraphs(await file.async('string')), allMarks[i]), allMarks[i].pictureLooks);
    xml = applyPlaceholders(xml, allMarks[i].placeholders);
    xml = applyTableDashes(xml, allMarks[i].tableDashes);
    xml = await applyPictureFills(zip, i + 1, xml, allMarks[i]);
    const t = slide.transition;
    if (t && TRANSITION_XML[t.type]) {
      const d = t.duration ?? 0.6;
      const spd = d < 0.45 ? 'fast' : d < 0.9 ? 'med' : 'slow';
      const tr = `<p:transition spd="${spd}">${TRANSITION_XML[t.type]}</p:transition>`;
      xml = xml.includes('</p:clrMapOvr>') ? xml.replace('</p:clrMapOvr>', `</p:clrMapOvr>${tr}`) : xml.replace('</p:sld>', `${tr}</p:sld>`);
    }
    if (slide.hidden) xml = xml.replace(/<p:sld(\s)/, '<p:sld show="0"$1');
    xml = applyGroups(uniqueShapeIds(xml), allMarks[i].groups);
    xml = insertTiming(xml, timingXml(timelineOf(xml, allMarks[i].timeline)));
    zip.file(path, xml);
  }
  await writeComments(zip, deck);
  await writeSections(zip, deck);
  await embedFonts(zip, deck);
  return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', compression: 'DEFLATE' });
}

/** Build the .pptx. Returns a Blob. */
export async function buildPptx(deck) {
  const theme = themeById(deck.themeId);
  const { w, h } = slideSize(deck);
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'LIBRESLIDES', width: w / IN, height: h / IN });
  pptx.layout = 'LIBRESLIDES';
  pptx.title = deck.title;
  pptx.company = 'LibreSlides';
  pptx.theme = { headFontFace: theme.fonts.title, bodyFontFace: theme.fonts.body };

  // What each slide needs fixing in the XML pptxgenjs writes (see applyGradients).
  const allMarks = [];
  for (const slide of deck.slides) {
    const s = pptx.addSlide();
    // Links as pptxgenjs takes them: a web address, or a slide by number.
    const index = deck.slides.indexOf(slide);
    const links = (href) => {
      if (/^(https?:|mailto:)/i.test(href)) return { url: href };
      const to = /^#slide:(.+)$/.exec(href)?.[1];
      const n = { next: index + 2, prev: index, first: 1, last: deck.slides.length }[to] ?? deck.slides.findIndex((x) => x.id === to) + 1;
      return to && n >= 1 && n <= deck.slides.length ? { slide: n } : null;
    };
    const marks = { number: allMarks.length + 1, links, count: 0, gradients: new Map(), pictures: new Map(), presets: new Map(), columns: new Map(), warps: new Map(), pictureLooks: new Map(), placeholders: new Map(), tableDashes: new Map(), timeline: [], groups: [], background: null };
    marks.groupSizes = slide.elements.reduce((m, e) => (e.groupId ? m.set(e.groupId, (m.get(e.groupId) || 0) + 1) : m), new Map());
    allMarks.push(marks);
    const bg = slide.background || theme.background;
    if (bg.type === 'image' && deck.media?.[bg.media]) {
      const data = await exportableImage(deck.media[bg.media], w, h);
      if (data) s.background = { data: dataForPptx(data) };
    } else if (bg.type === 'gradient') {
      // Written as a solid colour, then made a real gradient afterwards.
      s.background = { color: toHex(resolveColor(bg.from, theme)) || 'FFFFFF' };
      marks.background = gradFillXml({ type: 'linear', angle: bg.angle, stops: [{ pos: 0, color: bg.from }, { pos: 1, color: bg.to }] }, theme);
    } else {
      s.background = { color: toHex(resolveColor(bg.color, theme, backgroundBaseColor(theme.background))) || 'FFFFFF' };
    }
    for (const d of themeDecor(slide, deck)) await addElement(pptx, s, d, deck, theme, marks);
    for (const el of slide.elements) await addElement(pptx, s, el, deck, theme, marks);
    if (slide.notes) s.addNotes(slide.notes);
  }
  const buffer = await pptx.write({ outputType: 'arraybuffer' });
  return postProcess(buffer, deck, allMarks);
}

export async function exportPptx(deck) {
  const blob = await buildPptx(deck);
  return saveFile(blob, `${safeFileName(deck.title)}.pptx`, {
    description: 'PowerPoint presentation',
    accept: { 'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'] },
  });
}
