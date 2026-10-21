import PptxGenJS from 'pptxgenjs';
import JSZip from 'jszip';
import { slideSize, boxOf, htmlToText } from '../model/deck.js';
import { themeById, resolveColor, backgroundBaseColor } from '../model/themes.js';
import { textStyle, themeDecor } from '../render/render.js';
import { safeFileName } from '../ui/dom.js';
import { saveFile } from './files.js';
import { addChart } from './pptx-chart.js';
import { addMediaElement } from './pptx-media.js';
import { embedFonts } from './pptx-fonts.js';

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

const DASH = { solid: 'solid', dash: 'dash', dot: 'sysDot', dashDot: 'dashDot', longDash: 'lgDash' };

function lineOpts(stroke, theme) {
  if (!stroke || !stroke.width || !stroke.color || stroke.color === 'none') return { type: 'none' };
  return { color: toHex(resolveColor(stroke.color, theme)) || '000000', width: stroke.width, dashType: DASH[stroke.dash] || 'solid' };
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

/** Formatting of a run from its element and its ancestors up to the text root. */
function runStyle(node, root, base) {
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
    if (tag === 'A' && el.getAttribute('href')) s.hyperlink = { url: el.getAttribute('href') };
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
    if (st.color) s.color = toHex(st.color) || s.color;
    if (st.backgroundColor) s.highlight = toHex(st.backgroundColor) || s.highlight;
    if (st.fontSize && /px|pt$/.test(st.fontSize)) s.fontSize = parseFloat(st.fontSize);
    if (st.fontFamily) s.fontFace = st.fontFamily.split(',')[0].replace(/["']/g, '').trim();
  }
  return s;
}

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
  return o;
}

/** Paragraphs of an HTML fragment as pptxgenjs text runs. */
export function htmlToRuns(html, base = {}) {
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
          const options = toRunOptions(runStyle(c, root, base));
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
    for (const li of list.children) {
      if (li.tagName !== 'LI') continue;
      collect(li, { bullet: list.tagName === 'OL' ? { type: 'number' } : true, indentLevel: level, align: blockAlign(li) });
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

/** A gradient background as a picture (JPEG keeps it small; the same gradient is drawn once). */
const gradientCache = new Map();
function gradientImage(bg, theme, w, h) {
  const key = JSON.stringify([bg.from, bg.to, bg.angle, theme.id, w, h]);
  if (!gradientCache.has(key)) gradientCache.set(key, drawGradient(bg, theme, w, h));
  return gradientCache.get(key);
}

function drawGradient(bg, theme, w, h) {
  const canvas = document.createElement('canvas');
  canvas.width = 1280;
  canvas.height = Math.round((1280 * h) / w);
  const ctx = canvas.getContext('2d');
  const a = (((bg.angle ?? 180) - 90) * Math.PI) / 180;
  const cx = canvas.width / 2;
  const cy = canvas.height / 2;
  const r = Math.abs(cx * Math.cos(a)) + Math.abs(cy * Math.sin(a));
  const g = ctx.createLinearGradient(cx - Math.cos(a) * r, cy - Math.sin(a) * r, cx + Math.cos(a) * r, cy + Math.sin(a) * r);
  g.addColorStop(0, resolveColor(bg.from, theme));
  g.addColorStop(1, resolveColor(bg.to, theme));
  ctx.fillStyle = g;
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  return canvas.toDataURL('image/jpeg', 0.92);
}

// ------------------------------------------------------------------ slides

async function addElement(pptx, s, el, deck, theme) {
  const opacity = el.opacity ?? 1;
  const shadow = el.shadow ? { type: 'outer', blur: 8, offset: 4, angle: 90, color: '000000', opacity: 0.3 } : undefined;
  if (el.type === 'line') {
    let { x1, y1, x2, y2, arrowStart, arrowEnd } = el;
    if (x1 > x2) {
      [x1, x2, y1, y2] = [x2, x1, y2, y1];
      [arrowStart, arrowEnd] = [arrowEnd, arrowStart];
    }
    const line = lineOpts({ color: 'theme:text', width: 2, ...el.stroke, color: el.stroke?.color === 'none' ? 'theme:text' : el.stroke?.color }, theme);
    if (arrowStart) line.beginArrowType = 'triangle';
    if (arrowEnd) line.endArrowType = 'triangle';
    s.addShape('line', { ...geom({ x: x1, y: Math.min(y1, y2), w: x2 - x1, h: Math.abs(y2 - y1) }), flipV: y2 < y1, line });
    return;
  }
  const g = geom(boxOf(el));
  const common = { ...g, rotate: el.rot || 0, flipH: !!el.flipH, flipV: !!el.flipV, shadow };

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
    s.addImage(opts);
    return;
  }

  if (el.type === 'chart') {
    addChart(pptx, s, el, theme, g);
    return;
  }
  if (el.type === 'media') {
    await addMediaElement(pptx, s, el, deck, g);
    return;
  }

  if (el.type === 'table') {
    const t = textStyle({ ...el, type: 'text' }, theme);
    const base = { fontFace: t.font, fontSize: t.size, color: toHex(resolveColor(t.color, theme)) || '000000' };
    const header = toHex(resolveColor(el.fill || 'theme:accent1', theme));
    const rows = el.cells.map((row, r) => row.map((cell) => {
      const options = {};
      if (cell.fill) options.fill = { color: toHex(resolveColor(cell.fill, theme)) };
      else if (el.header && r === 0 && header) {
        options.fill = { color: header };
        options.color = 'FFFFFF';
        options.bold = true;
      } else if (el.banded && r % 2 === 0) options.fill = { color: 'EFEFEF' };
      const runs = htmlToRuns(cell.text, {});
      return { text: runs.length ? runs : '', options };
    }));
    const border = lineOpts(el.stroke, theme);
    s.addTable(rows, {
      x: g.x, y: g.y, w: g.w, colW: (el.colWidths || []).map((f) => f * g.w), rowH: g.h / el.cells.length,
      ...base, valign: 'middle', margin: 4,
      border: border.type === 'none' ? { type: 'none' } : { type: 'solid', pt: border.width, color: border.color },
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
  let extras = [];
  if (shape === 'custom') {
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
  const addExtras = () => { for (const extra of extras) s.addShape(shape, { ...common, ...extra }); };
  if (!hasText) {
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
  s.addText(htmlToRuns(el.text, {}), opts);
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

/** Tidy the paragraphs, and add transitions and "skip slide" to the generated XML. */
async function postProcess(buffer, deck) {
  const zip = await JSZip.loadAsync(buffer);
  for (const [i, slide] of deck.slides.entries()) {
    const path = `ppt/slides/slide${i + 1}.xml`;
    const file = zip.file(path);
    if (!file) continue;
    let xml = fixParagraphs(await file.async('string'));
    const t = slide.transition;
    if (t && TRANSITION_XML[t.type]) {
      const d = t.duration ?? 0.6;
      const spd = d < 0.45 ? 'fast' : d < 0.9 ? 'med' : 'slow';
      const tr = `<p:transition spd="${spd}">${TRANSITION_XML[t.type]}</p:transition>`;
      xml = xml.includes('</p:clrMapOvr>') ? xml.replace('</p:clrMapOvr>', `</p:clrMapOvr>${tr}`) : xml.replace('</p:sld>', `${tr}</p:sld>`);
    }
    if (slide.hidden) xml = xml.replace(/<p:sld(\s)/, '<p:sld show="0"$1');
    zip.file(path, xml);
  }
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

  for (const slide of deck.slides) {
    const s = pptx.addSlide();
    const bg = slide.background || theme.background;
    if (bg.type === 'image' && deck.media?.[bg.media]) {
      const data = await exportableImage(deck.media[bg.media], w, h);
      if (data) s.background = { data: dataForPptx(data) };
    } else if (bg.type === 'gradient') {
      try {
        s.background = { data: dataForPptx(gradientImage(bg, theme, w, h)) };
      } catch {
        s.background = { color: toHex(resolveColor(bg.from, theme)) };
      }
    } else {
      s.background = { color: toHex(resolveColor(bg.color, theme, backgroundBaseColor(theme.background))) || 'FFFFFF' };
    }
    for (const d of themeDecor(slide, deck)) await addElement(pptx, s, d, deck, theme);
    for (const el of slide.elements) await addElement(pptx, s, el, deck, theme);
    if (slide.notes) s.addNotes(slide.notes);
  }
  const buffer = await pptx.write({ outputType: 'arraybuffer' });
  return postProcess(buffer, deck);
}

export async function exportPptx(deck) {
  const blob = await buildPptx(deck);
  return saveFile(blob, `${safeFileName(deck.title)}.pptx`, {
    description: 'PowerPoint presentation',
    accept: { 'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'] },
  });
}
