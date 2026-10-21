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
import { exportsAsSmartArt, writeDiagram } from './pptx-diagram.js';
import { xmlText, emu } from './ooxml.js';
import { symbolBullet, EAST_ASIAN, COMPLEX_SCRIPT } from './pptx-text.js';
import { diagramAsShapes } from '../model/diagram-layout.js';

/**
 * PowerPoint export with pptxgenjs. Slides become real PowerPoint text boxes,
 * shapes, pictures, tables, charts and media (editable in PowerPoint,
 * Keynote, Google Slides and LibreOffice). What pptxgenjs can't write
 * (gradients, groups, placeholders, text effects and warps, transitions,
 * animations, comments, sections, SmartArt, fonts…) is added to its XML
 * afterwards: shapes that need it are given names ("LibreSlides shape N") to
 * be found by, and postProcess finishes each part.
 */
const IN = 72; // points per inch
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

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

/** A font name as pptxgenjs takes it: it writes font names into the XML as they are, so they're escaped here. */
const fontFace = (name) => (name ? xmlText(name) : undefined);

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

const GENERIC_FONTS = new Set(['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui', 'inherit', 'initial']);

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
    if (/\bnone\b/.test(deco) && !/underline|line-through/.test(deco)) Object.assign(s, { underline: false, strike: false });
    if (st.color && toHex(st.color)) Object.assign(s, { color: toHex(st.color), noFill: false });
    if (st.backgroundColor) s.highlight = toHex(st.backgroundColor) || s.highlight;
    if (st.fontSize && /px|pt$/.test(st.fontSize)) s.fontSize = parseFloat(st.fontSize);
    if (st.fontFamily) {
      // The first font is the run's; ones after it are for its East Asian or complex-script text (see fontsOf in the reader).
      const fonts = cssList(st.fontFamily).map((f) => f.replace(/["']/g, '').trim()).filter((f) => f && !GENERIC_FONTS.has(f));
      s.fontFace = fonts[0] || s.fontFace;
      s.moreFonts = fonts.slice(1);
    }
    if (st.letterSpacing && parseFloat(st.letterSpacing)) s.charSpacing = parseFloat(st.letterSpacing);
    if (st.textTransform === 'uppercase') s.caps = 'all';
    if (/small-caps/.test(st.fontVariant || '')) s.caps = 'small';
    textEffectsOf(el.getAttribute('style') || '', s);
  }
  return s;
}

/**
 * What pptxgenjs can't write — capitals, fields, outlines, gradient text,
 * shadows and glows, bold or italic turned off in a box that has them, and
 * fonts of their own for East Asian and complex-script text — travels in
 * the run's language as "x-ls-" and hex-coded JSON, and is
 * written into the XML afterwards (see fixRuns).
 */
const EXTRAS_LANG = 'x-ls-';
const hexOf = (str) => [...new TextEncoder().encode(str)].map((b) => b.toString(16).padStart(2, '0')).join('');
const unhex = (hex) => new TextDecoder().decode(new Uint8Array(hex.match(/../g).map((h) => parseInt(h, 16))));

function toRunOptions(s, text = '') {
  // Formatting a run turns off is written as off, or it would take the box's
  // (pptxgenjs gives every run the box's options, and only writes b and i when they're on).
  const o = {};
  if (s.bold != null) o.bold = !!s.bold;
  if (s.italic != null) o.italic = !!s.italic;
  if (s.underline != null) o.underline = { style: s.underline ? 'sng' : 'none' };
  if (s.strike != null) o.strike = s.strike ? 'sngStrike' : 'noStrike';
  if (s.color) o.color = s.color;
  if (s.highlight) o.highlight = s.highlight;
  if (s.fontSize) o.fontSize = s.fontSize;
  if (s.fontFace) o.fontFace = fontFace(s.fontFace);
  if (s.subscript) o.subscript = true;
  if (s.superscript) o.superscript = true;
  if (s.hyperlink) o.hyperlink = s.hyperlink;
  if (s.charSpacing) o.charSpacing = s.charSpacing;
  const extras = {};
  for (const key of ['caps', 'field', 'outline', 'gradient', 'shadows']) if (s[key]) extras[key] = s[key];
  if (s.noFill && !s.gradient) extras.noFill = true;
  // Fonts for East Asian and complex-script text, in the order they follow the run's own.
  const more = [...(s.moreFonts || [])];
  if (more.length && EAST_ASIAN.test(text)) extras.ea = more.shift();
  if (more.length && COMPLEX_SCRIPT.test(text)) extras.cs = more.shift();
  const off = [s.bold === false && 'b', s.italic === false && 'i'].filter(Boolean);
  if (off.length) extras.off = off;
  if (Object.keys(extras).length) o.lang = EXTRAS_LANG + hexOf(JSON.stringify(extras));
  return o;
}

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
  for (const name of extras.off || []) if (name === 'b' || name === 'i') attrs = `${attrs.replace(new RegExp(` ${name}="[^"]*"`), '')} ${name}="0"`;
  if (extras.field) attrs += ` x-ls-fld="${extras.field}"`;
  for (const tag of ['ea', 'cs']) if (typeof extras[tag] === 'string') body = body.replace(new RegExp(`<a:${tag} typeface="[^"]*"`), `<a:${tag} typeface="${xmlText(extras[tag])}"`);
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

/**
 * A list's bullet for pptxgenjs: numbered (with its style and first number),
 * or a bullet character of its own; with the character in Wingdings for a
 * symbol that has one there (see symbolBullet).
 */
function listBullet(list) {
  const type = list.style?.listStyleType || '';
  if (list.tagName === 'OL') {
    const b = { type: 'number' };
    if (NUMBER_STYLES[type]) b.style = NUMBER_STYLES[type];
    if (+list.getAttribute('start') > 1) b.numberStartAt = +list.getAttribute('start');
    return { bullet: b };
  }
  // A string bullet ('– '): its first character.
  const ch = /^["'](.)/u.exec(type)?.[1];
  if (!ch || ch.codePointAt(0) > 0xffff) return { bullet: true };
  return { bullet: { characterCode: ch.codePointAt(0).toString(16).padStart(4, '0').toUpperCase() }, symbol: symbolBullet(ch) };
}

/** A length in an inline style in points (1px is 1pt): px, pt, or em of `size`; null when unset or in other units. */
function lengthOf(css, prop, size) {
  const m = /^(-?[\d.]+)(px|pt|em)$/.exec(styleValue(css, prop) || '');
  return !m ? null : m[2] === 'em' ? +m[1] * size : +m[1];
}

/**
 * Paragraph settings pptxgenjs can't write (margins and indents, direction,
 * the bullet's colour, size and font, spacing set to none) travel in a tab
 * stop's alignment as "x-ls-" and hex-coded JSON, and are written into the
 * paragraph's a:pPr afterwards (see paraWithExtras).
 */
const paraTab = (extras) => [{ position: 1, alignment: EXTRAS_LANG + hexOf(JSON.stringify(extras)) }];

/** Paragraphs of an HTML fragment as pptxgenjs text runs; `box` is the text's size and spacing (see render.js textStyle). */
export function htmlToRuns(html, base = {}, links = null, box = {}) {
  const doc = new DOMParser().parseFromString(`<body>${html || ''}</body>`, 'text/html');
  const root = doc.body;
  // Loose inline content at the top level forms paragraphs of its own.
  let loose = null;
  for (const node of [...root.childNodes]) {
    if (node.nodeType === 1 && /^(P|DIV|H\d|UL|OL)$/.test(node.nodeName)) {
      loose = null;
      continue;
    }
    if (!loose) {
      loose = doc.createElement('p');
      node.before(loose);
    }
    loose.append(node);
  }
  const boxSize = box.size || 18;
  const isBlock = (n) => n.nodeType === 1 && /^(P|DIV|H\d|UL|OL)$/.test(n.nodeName);
  const styleOf = (el) => el.getAttribute?.('style') || '';
  // The text size an element is in: its own or the nearest one set around it.
  const sizeOf = (el) => {
    for (let n = el; n && n !== root; n = n.parentElement) {
      const v = lengthOf(styleOf(n), 'font-size', boxSize);
      if (v) return v;
    }
    return boxSize;
  };
  // Lists and paragraphs set in (see textBody in the reader): their own value, else the default.
  const startOf = (el, fallback) => lengthOf(styleOf(el), 'margin-inline-start', sizeOf(el))
    ?? lengthOf(styleOf(el), el.closest('[dir]')?.getAttribute('dir') === 'rtl' ? 'margin-right' : 'margin-left', sizeOf(el)) ?? fallback;
  const insetOf = (list) => lengthOf(styleOf(list), 'padding-inline-start', sizeOf(list))
    ?? lengthOf(styleOf(list), list.closest('[dir]')?.getAttribute('dir') === 'rtl' ? 'padding-right' : 'padding-left', sizeOf(list))
    ?? 1.25 * sizeOf(list);
  // The space above a paragraph: its margin and, where it comes first, those of the lists it starts (as CSS collapses them), else app.css's gap.
  const gapOf = (el) => {
    let gap = 0;
    for (let n = el; n && n !== root; n = n.parentElement) {
      const own = lengthOf(styleOf(n), 'margin-top', sizeOf(n));
      const parent = n.parentElement;
      const stacked = (parent === root && n.previousElementSibling) || (n.tagName === 'LI' && n.previousElementSibling?.tagName === 'LI')
        || (parent?.tagName === 'LI' && /^(UL|OL|P|DIV)$/.test(n.tagName));
      gap = Math.max(gap, own ?? (stacked ? (box.spacing || 0) * sizeOf(n) : 0));
      if (n.previousElementSibling || (parent !== root && [...parent.childNodes].some((c) => c !== n && c.nodeType === 3 && c.textContent.trim() && n.compareDocumentPosition(c) & 2))) break;
    }
    return gap;
  };

  const paragraphs = [];
  // A paragraph's runs: a list item's own text, leaving out the lists and paragraphs inside it (paragraphs of their own).
  const collect = (block, para) => {
    const runs = [];
    // A <br> is a line break inside the paragraph: the next run starts on a new line (<a:br/>).
    let breaks = 0;
    const walk = (n) => {
      for (const c of n.childNodes) {
        if (c.nodeType === 3) {
          const text = c.textContent.replace(/​/g, '');
          if (!text) continue;
          const options = toRunOptions(runStyle(c, root, base, links), text);
          if (breaks && !runs.length) runs.push({ text: '', options }); // the paragraph starts with an empty line
          for (; breaks > 1; breaks--) runs.push({ text: '', options: { ...options, softBreakBefore: true } });
          runs.push({ text, options: breaks ? { ...options, softBreakBefore: true } : options });
          breaks = 0;
        } else if (c.nodeName === 'BR') breaks++;
        else if (c.nodeName !== 'UL' && c.nodeName !== 'OL' && !para.own.includes(c)) walk(c);
      }
    };
    walk(block);
    // Trailing <br>s only keep an empty line open in the editor.
    paragraphs.push({ runs, para, el: block });
  };
  const blockAlign = (el) => {
    const a = el.style?.textAlign;
    return ['left', 'center', 'right', 'justify'].includes(a) ? a : null;
  };
  // Blocks inside a list item: ones after its own text are paragraphs of their own, as is any but a first one before it.
  const ownBlocks = (li) => {
    const blocks = [...li.children].filter(isBlock);
    const lead = blocks[0]?.tagName !== 'UL' && blocks[0]?.tagName !== 'OL' && ![...li.childNodes].slice(0, [...li.childNodes].indexOf(blocks[0]))
      .some((c) => c.nodeType === 3 ? c.textContent.trim() : !isBlock(c));
    return lead ? blocks.slice(1) : blocks;
  };
  const visitBlocks = (blocks, level, from) => {
    for (const b of blocks) {
      if (b.tagName === 'UL' || b.tagName === 'OL') visitList(b, level, from);
      else {
        const marL = from + startOf(b, 0);
        collect(b, { own: [], level, marL, indent: lengthOf(styleOf(b), 'text-indent', sizeOf(b)) || 0, align: blockAlign(b) });
      }
    }
  };
  const visitList = (list, level, from) => {
    const { bullet, symbol } = listBullet(list);
    const listStart = from + insetOf(list);
    for (const item of list.children) {
      // A list straight inside a list (as browsers indent) is the next level down.
      if (item.tagName === 'UL' || item.tagName === 'OL') visitList(item, level + 1, listStart);
      if (item.tagName !== 'LI') continue;
      const textStart = listStart + startOf(item, 0);
      const own = ownBlocks(item);
      // The bullet hangs where it's drawn: just before the text, at most a list's usual inset (1.25em) away.
      const hang = Math.min(textStart - from, 1.25 * sizeOf(item));
      collect(item, { own, bullet, symbol, level, marL: textStart, indent: -hang, align: blockAlign(item) });
      visitBlocks(own, level + 1, textStart);
    }
  };
  visitBlocks([...root.children], 0, 0);

  // Space between paragraphs: the box's, unless any has its own, when each one's is written as its space before.
  const gaps = paragraphs.map(({ el }, i) => (i ? gapOf(el) : 0));
  const ownGaps = paragraphs.some(({ el }, i) => i && Math.abs(gaps[i] - (box.spacing || 0) * sizeOf(el)) > 0.5);

  const out = [];
  paragraphs.forEach(({ runs, para, el }, i) => {
    const pOpts = {};
    const css = styleOf(el);
    if (para.bullet) pOpts.bullet = para.bullet;
    if (para.level) pOpts.indentLevel = para.level;
    if (para.align) pOpts.align = para.align;
    const line = styleValue(css, 'line-height');
    if (/^[\d.]+$/.test(line || '') && +line > 0) pOpts.lineSpacingMultiple = Math.round((+line / 1.2) * 100) / 100;
    else if (lengthOf(css, 'line-height', sizeOf(el)) > 0) pOpts.lineSpacing = lengthOf(css, 'line-height', sizeOf(el));
    const extras = {};
    const marL = Math.max(0, para.marL);
    if (marL || para.indent || para.bullet) Object.assign(extras, { marL, indent: Math.max(-marL, para.indent) });
    if (el.closest('[dir]')?.getAttribute('dir') === 'rtl') extras.rtl = true;
    if (ownGaps) Object.assign(extras, { spcBef: Math.round(gaps[i] * 10) / 10, spcAft: 0 });
    if (para.bullet) {
      // A bullet takes its paragraph's first run's colour and size, unless it has its own (see textBody in the reader).
      const first = runs.find((r) => r.text)?.options || {};
      const color = toHex(styleValue(css, '--bu-color') || '') || toHex(el.style?.color || '');
      if (color && color !== (first.color || toHex(base.color || ''))) extras.buClr = color;
      const scale = (parseFloat(styleValue(css, '--bu-size')) || 100) / 100;
      const pct = (sizeOf(el) * scale) / (first.fontSize || sizeOf(el));
      if (Math.abs(pct - 1) > 0.01) extras.buSzPct = Math.round(Math.max(0.25, Math.min(4, pct)) * 100000);
      if (para.symbol) Object.assign(extras, { buFont: para.symbol.font, buChar: para.symbol.char });
    }
    if (!runs.length) runs.push({ text: '', options: toRunOptions(base) });
    runs.forEach((r, k) => {
      const options = { ...r.options, ...pOpts };
      if (k === 0 && Object.keys(extras).length) options.tabStops = paraTab(extras);
      if (k === runs.length - 1 && i < paragraphs.length - 1) options.breakLine = true;
      out.push({ text: r.text, options });
    });
  });
  return out;
}

// ------------------------------------------------------------------ pictures

/** A picture drawn into a PNG at about the size it's shown (sharper, up to 2,000 pixels across). */
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

/**
 * A picture as PowerPoint can show it: SVG (which it can't read reliably),
 * WebP, AVIF and BMP are drawn into PNGs, each once per export and size.
 */
function exportableImage(dataUrl, w = 400, h = 300, cache = null) {
  if (!dataUrl) return null;
  if (!/^data:image\/(svg|webp|avif|bmp)/i.test(dataUrl) && /;base64,/.test(dataUrl)) return dataUrl;
  const key = `${Math.round(w)}x${Math.round(h)} ${dataUrl}`;
  if (cache?.has(key)) return cache.get(key);
  const png = rasterize(dataUrl, w, h).catch(() => null);
  cache?.set(key, png);
  return png;
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
    const data = await exportableImage(deck.media?.[el.media], el.w, el.h, marks.pngs);
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
  if (el.type === 'diagram') {
    // SmartArt from PowerPoint: a stand-in shape, swapped for the diagram's graphic frame afterwards (see applyDiagrams).
    const name = objectName || `LibreSlides shape ${++marks.count}`;
    marks.diagrams.set(name, el);
    s.addShape('rect', { ...common, rotate: 0, objectName: name, fill: { type: 'none' }, line: { type: 'none' } });
    return;
  }

  if (el.type === 'table') {
    const t = textStyle({ ...el, type: 'text' }, theme);
    const base = { fontFace: fontFace(t.font), fontSize: t.size, color: toHex(resolveColor(t.color, theme)) || '000000' };
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
      const runs = htmlToRuns(withFields(cell.text, marks.number), {}, marks.links, { size: t.size });
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
  const picture = el.fillImage && (await exportableImage(deck.media?.[el.fillImage.media], el.w, el.h, marks.pngs));
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
    fontFace: fontFace(t.font),
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
  // A text box's own link is added afterwards (see applyTextLinks).
  if (hyperlink) {
    marks.textLinks.set(mark(), hyperlink);
    opts.hyperlink = undefined;
  }
  s.addText(hasText ? htmlToRuns(withFields(el.text, marks.number), runBase, marks.links, { size: t.size, spacing: t.spacing }) : [{ text: '', options: {} }], opts);
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

/** A paragraph's a:pPr with its extras (see paraTab) written in, in the order PowerPoint requires. */
function paraWithExtras(pPr) {
  const marker = /<a:tabLst><a:tab pos="\d+" algn="x-ls-([0-9a-f]+)"\/><\/a:tabLst>/.exec(pPr);
  if (!marker) return pPr;
  pPr = pPr.replace(marker[0], '');
  let x;
  try {
    x = JSON.parse(unhex(marker[1]));
  } catch {
    return pPr;
  }
  const m = /^<a:pPr\b([^>]*?)\/?>([\s\S]*?)(?:<\/a:pPr>)?$/.exec(pPr);
  if (!m) return pPr;
  let [, attrs, inner] = m;
  const set = (name, v) => {
    attrs = `${attrs.replace(new RegExp(`\\s${name}="[^"]*"`), '')} ${name}="${v}"`;
  };
  if (Number.isFinite(x.marL)) set('marL', emu(x.marL));
  if (Number.isFinite(x.indent)) set('indent', emu(x.indent));
  if (x.rtl === true) set('rtl', '1');
  const take = (...tags) => {
    for (const tag of tags) {
      const found = new RegExp(`<a:${tag}\\b[^>]*/>|<a:${tag}\\b[^>]*>[\\s\\S]*?</a:${tag}>`).exec(inner);
      if (found) {
        inner = inner.replace(found[0], '');
        return found[0];
      }
    }
    return '';
  };
  const lnSpc = take('lnSpc');
  let spcBef = take('spcBef');
  let spcAft = take('spcAft');
  let clr = take('buClrTx', 'buClr');
  let size = take('buSzTx', 'buSzPct', 'buSzPts');
  let font = take('buFontTx', 'buFont');
  let bullet = take('buNone', 'buAutoNum', 'buChar', 'buBlip');
  const pts = (v) => `<a:spcPts val="${Math.round(Math.max(0, Math.min(1584, v)) * 100)}"/>`;
  if (Number.isFinite(x.spcBef)) spcBef = `<a:spcBef>${pts(x.spcBef)}</a:spcBef>`;
  if (Number.isFinite(x.spcAft)) spcAft = `<a:spcAft>${pts(x.spcAft)}</a:spcAft>`;
  if (/^[0-9A-F]{6}$/i.test(x.buClr || '')) clr = `<a:buClr><a:srgbClr val="${x.buClr.toUpperCase()}"/></a:buClr>`;
  if (Number.isFinite(x.buSzPct)) size = `<a:buSzPct val="${Math.round(x.buSzPct)}"/>`;
  if (typeof x.buFont === 'string') font = `<a:buFont typeface="${xmlText(x.buFont)}" pitchFamily="2" charset="2"/>`;
  if (typeof x.buChar === 'string' && bullet.startsWith('<a:buChar')) bullet = `<a:buChar char="${xmlText(x.buChar)}"/>`;
  return `<a:pPr${attrs}>${lnSpc}${spcBef}${spcAft}${clr}${size}${font}${bullet}${inner}</a:pPr>`;
}

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
      return keep ? paraWithExtras(pPr) : '';
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

/**
 * What pptxgenjs can't write into a shape's own XML, for the shapes named
 * in `marks`: text columns, gradient fills (and a gradient background),
 * preset shapes' adjust values and text warps.
 */
export function applyGradients(xml, marks) {
  if (!marks) return xml;
  // Text in columns.
  for (const [name, { n, gap }] of marks.columns || []) {
    const at = xml.indexOf(`name="${name}"`);
    const end = xml.indexOf('</p:sp>', at);
    if (at < 0 || end < 0) continue;
    const block = xml.slice(at, end).replace('<a:bodyPr', `<a:bodyPr numCol="${Math.round(n)}"${gap != null ? ` spcCol="${emu(gap)}"` : ''}`);
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

/**
 * Text boxes' own links (a click anywhere on the box): a web address, or a
 * slide by number. pptxgenjs gives them no relationship (r:id="rIdundefined")
 * and puts them on every run of the text that has no link of its own.
 */
async function applyTextLinks(zip, slideNo, xml, links) {
  if (!links?.size) return xml;
  const relsPath = `ppt/slides/_rels/slide${slideNo}.xml.rels`;
  let rels = await zip.file(relsPath).async('string');
  let n = 0;
  for (const [name, link] of links) {
    const id = `rIdLsLink${++n}`;
    const rel = link.url
      ? `<Relationship Id="${id}" Type="${REL}/hyperlink" Target="${xmlText(link.url)}" TargetMode="External"/>`
      : `<Relationship Id="${id}" Type="${REL}/slide" Target="slide${link.slide}.xml"/>`;
    rels = rels.replace('</Relationships>', `${rel}</Relationships>`);
    const click = `<a:hlinkClick r:id="${id}"${link.url ? '' : ' action="ppaction://hlinksldjump"'}/>`;
    xml = editShape(xml, name, (sp) => sp.replace(/<p:cNvPr\b([^>]*?)(\/?)>/, (m, attrs, empty) => `<p:cNvPr${attrs}>${click}${empty ? '</p:cNvPr>' : ''}`));
  }
  zip.file(relsPath, rels);
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

/** Swap the stand-ins for SmartArt diagrams (see addElement) for their graphic frames, adding their parts. */
async function applyDiagrams(zip, slideNo, xml, diagrams, deck) {
  for (const [name, el] of diagrams || []) {
    const b = shapeBlock(xml, name);
    if (!b) continue;
    const cNvPr = /<p:cNvPr\b[^>]*?(?:\/>|>[\s\S]*?<\/p:cNvPr>)/.exec(xml.slice(b.start, b.end))?.[0];
    const frame = await writeDiagram(zip, slideNo, el, themeById(deck.themeId), cNvPr, deck.media);
    xml = xml.slice(0, b.start) + frame + xml.slice(b.end);
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

// Picture formats by their first bytes, for parts named after another format.
const PICTURE_MAGIC = [
  { bytes: [0x89, 0x50, 0x4e, 0x47], ext: 'png', type: 'image/png' },
  { bytes: [0xff, 0xd8, 0xff], ext: 'jpg', type: 'image/jpeg', also: ['jpeg'] },
  { bytes: [0x47, 0x49, 0x46, 0x38], ext: 'gif', type: 'image/gif' },
];
const sameBytes = (a, b) => a.length === b.length && a.every((v, i) => v === b[i]);
// Characters XML 1.0 can't hold (control characters, lone surrogates), which text may carry into pptxgenjs's XML.
const NOT_XML = /[\ud800-\udbff][\udc00-\udfff]|[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff\ud800-\udfff]/g;

/**
 * Last touches to the package: each picture stored once (pptxgenjs writes a
 * copy per use, so a master's logo would be on every slide) under the name
 * of its real format (it calls slide backgrounds .png whatever they are),
 * no content types for parts that aren't there, and no characters XML can't hold.
 */
export async function tidyPackage(zip) {
  const renamed = new Map(); // ppt/media/… → the name it now has
  const kept = new Map(); // byte length → [{ name, bytes }]
  for (const file of zip.file(/^ppt\/media\/[^/]+$/)) {
    const bytes = await file.async('uint8array');
    const ext = (/\.([^.]+)$/.exec(file.name)?.[1] || '').toLowerCase();
    const real = PICTURE_MAGIC.find((m) => m.bytes.every((b, i) => bytes[i] === b));
    let name = file.name;
    if (real && ext !== real.ext && !real.also?.includes(ext)) {
      name = `${file.name.replace(/\.[^.]*$/, '')}.${real.ext}`;
      while (zip.file(name) || [...renamed.values()].includes(name)) name = name.replace(/(\.[^.]+)$/, '-1$1');
    }
    const same = (kept.get(bytes.length) || []).find((k) => k.name.split('.').pop() === name.split('.').pop() && sameBytes(k.bytes, bytes));
    if (same) name = same.name;
    else (kept.get(bytes.length) || kept.set(bytes.length, []).get(bytes.length)).push({ name, bytes });
    if (name === file.name) continue;
    renamed.set(file.name, name);
    zip.remove(file.name);
    if (!same) zip.file(name, bytes);
  }
  const typesPath = '[Content_Types].xml';
  let types = await zip.file(typesPath).async('string');
  if (renamed.size) {
    for (const file of zip.file(/\.rels$/)) {
      const rels = await file.async('string');
      // Relationships point at media from parts one folder below ppt/ (slides, layouts, diagrams…).
      const fixed = rels.replace(/Target="\.\.\/media\/([^"]+)"/g, (m, base) => {
        const to = renamed.get(`ppt/media/${base}`);
        return to ? `Target="../media/${to.split('/').pop()}"` : m;
      });
      if (fixed !== rels) zip.file(file.name, fixed);
    }
    for (const m of PICTURE_MAGIC) {
      const used = [...renamed.values()].some((n) => n.endsWith(`.${m.ext}`));
      if (used && !new RegExp(`<Default Extension="${m.ext}"`, 'i').test(types)) types = types.replace(/<Types\b[^>]*>/, (t) => `${t}<Default Extension="${m.ext}" ContentType="${m.type}"/>`);
    }
  }
  // pptxgenjs lists a slide master per slide, though there's one.
  types = types.replace(/<Override PartName="\/([^"]+)"[^>]*\/>/g, (o, part) => (zip.file(part) ? o : ''));
  zip.file(typesPath, types);
  for (const file of zip.file(/\.(xml|rels)$/)) {
    const xml = await file.async('string');
    const clean = xml.replace(NOT_XML, (c) => (c.length === 2 ? c : ''));
    if (clean !== xml) zip.file(file.name, clean);
  }
}

/**
 * Finish the package pptxgenjs wrote: media and charts, then each slide's
 * XML (paragraphs, the shapes' marks, transitions, "skip slide", groups and
 * animations), comments, sections, fonts, and last tidyPackage.
 */
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
    xml = await applyTextLinks(zip, i + 1, xml, allMarks[i].textLinks);
    xml = await applyDiagrams(zip, i + 1, xml, allMarks[i].diagrams, deck);
    const t = slide.transition;
    if (t && Object.hasOwn(TRANSITION_XML, t.type)) {
      const d = Number.isFinite(t.duration) ? t.duration : 0.6;
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
  await tidyPackage(zip);
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

  // What each slide needs fixing in the XML pptxgenjs writes (see postProcess).
  const allMarks = [];
  const pngs = new Map(); // pictures drawn into PNGs (see exportableImage)
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
    const marks = { number: allMarks.length + 1, links, pngs, count: 0, gradients: new Map(), pictures: new Map(), presets: new Map(), columns: new Map(), warps: new Map(), pictureLooks: new Map(), placeholders: new Map(), tableDashes: new Map(), textLinks: new Map(), diagrams: new Map(), timeline: [], groups: [], background: null };
    // Diagrams that don't go as SmartArt go as the shapes they're drawn with.
    const elements = slide.elements.flatMap((el) => (el.type === 'diagram' && !exportsAsSmartArt(el, deck.media) ? diagramAsShapes(el) : [el]));
    marks.groupSizes = elements.reduce((m, e) => (e.groupId ? m.set(e.groupId, (m.get(e.groupId) || 0) + 1) : m), new Map());
    allMarks.push(marks);
    const bg = slide.background || theme.background;
    if (bg.type === 'image' && deck.media?.[bg.media]) {
      const data = await exportableImage(deck.media[bg.media], w, h, pngs);
      if (data) s.background = { data: dataForPptx(data) };
    } else if (bg.type === 'gradient') {
      // Written as a solid colour, then made a real gradient afterwards.
      s.background = { color: toHex(resolveColor(bg.from, theme)) || 'FFFFFF' };
      marks.background = gradFillXml({ type: 'linear', angle: bg.angle, stops: [{ pos: 0, color: bg.from }, { pos: 1, color: bg.to }] }, theme);
    } else {
      s.background = { color: toHex(resolveColor(bg.color, theme, backgroundBaseColor(theme.background))) || 'FFFFFF' };
    }
    for (const d of themeDecor(slide, deck)) await addElement(pptx, s, d, deck, theme, marks);
    for (const el of elements) await addElement(pptx, s, el, deck, theme, marks);
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
