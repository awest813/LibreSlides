import { ANIM_KINDS } from './deck.js';
import { normalizeLevels, outlineTree } from './diagram-outline.js';

/**
 * Diagram layout: a diagram element (see deck.js) as the plain elements it
 * is drawn with — shapes with text, arrows and connector lines, using
 * LibreSlides' own shape kinds — in its box's coordinates (0, 0 at its
 * top-left corner). Pure: the renderer draws what this returns inside the
 * diagram's box, and the PowerPoint writer saves it.
 *
 * Text is sized as SmartArt sizes it: one font size for all the diagram's
 * shapes, the largest at which the longest text fits its shape (estimated
 * from the characters' typical widths, since there's no page to measure on).
 */

// ------------------------------------------------------------------ fitting text

const LINE = 1.2; // line height, in ems
const PAD = { x: 7.2, y: 3.6 }; // a text box's margins (render.js .el-textbox)
const BULLET_INDENT = 1.25; // ems per list level (app.css)

const glyphWidths = new Map();

/** About how wide a character is, in ems (erring wide, as the fonts LibreSlides offers are). */
function glyphWidth(ch) {
  let w = glyphWidths.get(ch);
  if (w === undefined) {
    w = guessWidth(ch);
    if (glyphWidths.size < 5000) glyphWidths.set(ch, w);
  }
  return w;
}

function guessWidth(ch) {
  if (/[\u1100-\u115f\u2e80-\ua4cf\uac00-\ud7a3\uf900-\ufaff\ufe30-\ufe4f\uff00-\uff60\uffe0-\uffe6]/u.test(ch) || ch.length > 1) return 1.05;
  if (/[iljtfrI.,:;'|!()[\] ]/.test(ch)) return 0.36;
  if (/[mwMW@%]/.test(ch)) return 0.95;
  if (/[A-Z0-9#&?]/.test(ch)) return 0.72;
  return 0.61;
}

/** About how wide `text` is at `size` points. */
export const textWidth = (text, size) => [...String(text)].reduce((w, ch) => w + glyphWidth(ch), 0) * size;

/** How many lines a paragraph takes in a column `width` wide, wrapping between words (and inside a word too long for a line). */
export function lineCount(text, width, size) {
  if (width <= 0) return Infinity;
  const space = 0.28 * size;
  let lines = 1;
  let x = 0;
  for (const word of String(text).split(/\s+/).filter(Boolean)) {
    const w = textWidth(word, size);
    if (x > 0 && x + space + w <= width) {
      x += space + w;
      continue;
    }
    if (x > 0) lines++;
    const extra = Math.max(0, Math.ceil(w / width) - 1);
    lines += extra;
    x = w - extra * width;
  }
  return lines;
}

/** Whether paragraphs [{ text, indent (list levels) }] fit a w × h text area at `size`, without breaking a word. */
export function textFits(paras, w, h, size) {
  let height = 0;
  for (const p of paras) {
    const width = w - (p.indent || 0) * BULLET_INDENT * size;
    if (String(p.text).split(/\s+/).some((word) => textWidth(word, size) > width)) return false;
    height += lineCount(p.text, width, size) * size * LINE;
  }
  return height <= h + 0.01;
}

/**
 * The largest font size (in half points, from `max` down to `min`) at which
 * every box's text fits it: boxes [{ w, h, paras }] are text areas.
 */
export function fitFontSize(boxes, max = 32, min = 6) {
  const list = boxes.filter((b) => b.paras.some((p) => p.text));
  for (let size = Math.floor(max * 2) / 2; size > min; size -= 0.5) {
    if (list.every((b) => textFits(b.paras, b.w - 2 * PAD.x, b.h - 2 * PAD.y, size))) return size;
  }
  return min;
}

// ------------------------------------------------------------------ elements

const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

/** Lines [{ text, depth }] as nested bulleted lists. */
function bulletsHtml(lines) {
  let html = '';
  let depth = -1;
  for (const l of lines) {
    const d = Math.max(0, Math.min(l.depth, depth + 1));
    if (d > depth) html += '<ul><li>'.repeat(d - depth);
    else html += `${'</li></ul>'.repeat(depth - d)}</li><li>`;
    html += esc(l.text) || '<br>';
    depth = d;
  }
  return html + '</li></ul>'.repeat(depth + 1);
}

/** A shape's text: its node's text, centred, with the nodes under it as bullets. */
function textHtml(text, lines = []) {
  const p = esc(text) || '<br>';
  return lines.length ? `<p style="text-align: center">${p}</p>${bulletsHtml(lines)}` : `<p>${p}</p>`;
}

/** A shape's text as paragraphs to fit (see textFits). */
const parasOf = (text, lines = []) => [{ text, indent: 0 }, ...lines.map((l) => ({ text: l.text, indent: l.depth + 1 }))];

/** A shape's descendants as bullet lines: [{ text, depth }] (depth 0 under it). */
function linesUnder(item, depth = 0) {
  return item.children.flatMap((c) => [{ text: c.node.text, depth }, ...linesUnder(c, depth + 1)]);
}

/** The diagram's colours: the fill for the `i`th shape (in the order the layout colours them). */
function colorOf(d, i) {
  const accent = (n) => d.palette?.[n - 1] || `theme:accent${n}`;
  if (d.colors === 'colorful') return accent((i % 6) + 1);
  const m = /^accent([1-6])$/.exec(d.colors || '');
  return accent(m ? +m[1] : 1);
}

/**
 * Builds a layout's elements: shapes holding text, plain text, lines and
 * arrows, styled by the diagram's colours and style. Text sizes are filled
 * in at the end (see finish).
 */
class Builder {
  constructor(el) {
    this.el = el;
    this.d = el.diagram;
    this.out = [];
    this.fits = []; // { el, area, paras, group }
  }

  /** A shape with text in it; `area` is its text area in fractions of the box (l, t, r, b insets). */
  shape(kind, box, i, text = '', lines = [], { area = null, group = 'main', textRect = null, paths = null, plain = false } = {}) {
    const fill = colorOf(this.d, i);
    const style = this.d.style;
    const el = {
      id: `${this.el.id}-${this.out.length}`, type: 'shape', shape: kind, ...box, rot: box.rot || 0,
      text: plain ? '' : textHtml(text, lines),
      fill: style === 'outline' ? 'none' : fill,
      stroke: style === 'outline' ? { color: fill, width: 2, dash: 'solid' } : null,
      style: this.textStyle(style === 'outline' ? 'theme:text' : '#ffffff', lines.length ? 'left' : 'center'),
    };
    if (paths) el.paths = paths;
    if (textRect) el.textRect = textRect;
    if (style === 'raised') el.shadow = true;
    this.out.push(el);
    if (!plain) this.fit(el, text, lines, area || textRect, group);
    return el;
  }

  /** Text on its own (beside or under a shape). */
  text(box, text, lines, { group = 'side', align = 'left', valign = 'top' } = {}) {
    const el = {
      id: `${this.el.id}-${this.out.length}`, type: 'text', ...box, rot: 0, fill: 'none', stroke: null,
      text: text == null ? bulletsHtml(lines) : textHtml(text, lines),
      style: { ...this.textStyle('theme:text', align), valign },
    };
    this.out.push(el);
    this.fit(el, text, lines, null, group);
    return el;
  }

  /** A shape that holds no text (Venn circles, arrows), faded to `opacity`. */
  mark(kind, box, i, { opacity = 1, color = null } = {}) {
    const el = this.shape(kind, box, i, '', [], { plain: true });
    if (color) el.fill = color;
    if (opacity < 1) el.opacity = opacity;
    return el;
  }

  /** A connector line, in the first shape's colour. */
  line(x1, y1, x2, y2) {
    this.out.push({
      id: `${this.el.id}-${this.out.length}`, type: 'line', x1, y1, x2, y2, rot: 0,
      stroke: { color: colorOf(this.d, 0), width: 1.5, dash: 'solid' }, arrowStart: false, arrowEnd: false,
    });
  }

  textStyle(color, align) {
    const s = { size: 18, color, align, valign: 'middle', lineHeight: 1.1, spacing: 0 };
    if (this.d.font) s.font = this.d.font;
    return s;
  }

  fit(el, text, lines, area, group) {
    const a = area || { l: 0, t: 0, r: 0, b: 0 };
    this.fits.push({ el, group, paras: text == null ? lines.map((l) => ({ text: l.text, indent: l.depth + 1 })) : parasOf(text, lines), w: el.w * (1 - a.l - a.r), h: el.h * (1 - a.t - a.b) });
  }

  /** Give each group of text one size: the largest at which all of it fits (text beside shapes no bigger than in them). */
  finish() {
    const max = Math.max(8, Math.min(36, Math.min(this.el.w, this.el.h) / 5));
    const main = fitFontSize(this.fits.filter((f) => f.group === 'main'), max);
    const side = fitFontSize(this.fits.filter((f) => f.group === 'side'), Math.min(max, main * 0.85 || max));
    for (const f of this.fits) f.el.style.size = f.group === 'main' ? main : side;
    return this.out;
  }
}

// ------------------------------------------------------------------ layouts

/** The top-level nodes, each with the nodes under it as bullet lines: [{ text, lines }]. */
const topItems = (nodes) => outlineTree(nodes).map((item) => ({ text: item.node.text, lines: linesUnder(item) }));

/** A grid of boxes about 1.6 times as wide as tall, as large as fits, the last row centred (SmartArt's "Basic Block List"). */
function basicList(b, W, H) {
  const items = topItems(b.d.nodes);
  const n = items.length;
  const g = Math.min(W, H) * 0.04;
  let best = null;
  for (let cols = 1; cols <= n; cols++) {
    const rows = Math.ceil(n / cols);
    let bw = (W - (cols - 1) * g) / cols;
    let bh = (H - (rows - 1) * g) / rows;
    if (bw * 0.6 > bh) bw = bh / 0.6;
    else bh = bw * 0.6;
    if (!best || bw * bh > best.bw * best.bh) best = { cols, rows, bw, bh };
  }
  if (!best) return;
  const { cols, rows, bw, bh } = best;
  const y0 = (H - rows * bh - (rows - 1) * g) / 2;
  items.forEach((it, i) => {
    const r = Math.floor(i / cols);
    const inRow = Math.min(cols, n - r * cols);
    const x0 = (W - inRow * bw - (inRow - 1) * g) / 2;
    b.shape('rect', { x: x0 + (i % cols) * (bw + g), y: y0 + r * (bh + g), w: bw, h: bh }, i, it.text, it.lines);
  });
}

/** Full-width bars, one under another, each with the text under it as a list below it (SmartArt's "Vertical Bullet List"). */
function verticalList(b, W, H) {
  const items = topItems(b.d.nodes);
  // In units of a bar's height: the bars, the lists (0.7 a line) and the space between.
  const units = items.reduce((s, it) => s + 1 + it.lines.length * 0.7, 0) + Math.max(0, items.length - 1) * 0.15;
  if (!units) return;
  const u = Math.min(H / units, H / 3);
  let y = (H - units * u) / 2;
  items.forEach((it, i) => {
    b.shape('roundRect', { x: 0, y, w: W, h: u }, i, it.text);
    y += u;
    if (it.lines.length) {
      const h = it.lines.length * 0.7 * u;
      b.text({ x: W * 0.04, y, w: W * 0.96, h }, null, it.lines);
      y += h;
    }
    y += 0.15 * u;
  });
}

/** Boxes in a row with arrows between them (SmartArt's "Basic Process"). */
function process(b, W, H) {
  const items = topItems(b.d.nodes);
  const n = items.length;
  if (!n) return;
  // Each arrow is a quarter of a box wide, with a little space either side.
  const bw = W / (n + (n - 1) * 0.41);
  const bh = Math.min(H, bw * 0.75);
  const y = (H - bh) / 2;
  const aw = bw * 0.25;
  const ah = Math.min(aw * 1.2, bh * 0.6);
  items.forEach((it, i) => {
    const x = i * bw * 1.41;
    b.shape('roundRect', { x, y, w: bw, h: bh }, i, it.text, it.lines);
    if (i < n - 1) b.mark('rightArrow', { x: x + bw * 1.08, y: (H - ah) / 2, w: aw, h: ah }, i, { opacity: 0.6 });
  });
}

/** Chevrons nested in a row; the text under each step goes below it (SmartArt's "Basic Chevron Process"). */
function chevrons(b, W, H) {
  const items = topItems(b.d.nodes);
  const n = items.length;
  if (!n) return;
  const lists = items.some((it) => it.lines.length);
  // The notch of each chevron takes in the point of the one before (d is the shape's depth, half its height).
  let bh = lists ? H * 0.3 : H;
  let bw = (W + (n - 1) * (bh / 2 - bh * 0.1)) / n;
  if (bh > bw * 0.4) {
    bh = bw * 0.4;
    bw = (W + (n - 1) * (bh / 2 - bh * 0.1)) / n;
  }
  // Many steps: that still leaves them taller than wide, which the overlap above doesn't
  // allow for. Chevrons 0.4 as tall as wide fill the width exactly.
  if (bh > bw) {
    bw = W / (n - 0.16 * (n - 1));
    bh = bw * 0.4;
  }
  const d = Math.min(bw, bh) / 2;
  const step = bw - d + bh * 0.1;
  const y = lists ? 0 : (H - bh) / 2;
  // Text clears the notch, which cuts in deepest at the middle, and runs a little into the point.
  items.forEach((it, i) => {
    const x = i * step;
    const first = i === 0;
    b.shape(first ? 'homePlate' : 'chevron', { x, y, w: bw, h: bh }, i, it.text, [], { textRect: { l: first ? 0.04 : (d * 0.95) / bw, t: 0, r: (d * 0.35) / bw, b: 0 } });
    if (it.lines.length) b.text({ x: x + (first ? 0 : d / 2), y: bh + H * 0.04, w: step - (first ? 0 : d / 2), h: H - bh - H * 0.04 }, null, it.lines);
  });
}

/** Circles round a ring with arrows between them, clockwise from the top (SmartArt's "Basic Cycle"). */
function cycle(b, W, H) {
  const items = topItems(b.d.nodes);
  const n = items.length;
  if (!n) return;
  const m = Math.min(W, H);
  const [cx, cy] = [W / 2, H / 2];
  if (n === 1) {
    const s = m * 0.6;
    b.shape('ellipse', { x: cx - s / 2, y: cy - s / 2, w: s, h: s }, 0, items[0].text, items[0].lines, { area: ELLIPSE_AREA });
    return;
  }
  // Circles take seven tenths of the space between their centres, the arrows the rest.
  const sin = Math.sin(Math.PI / n);
  const R = m / 2 / (1 + 0.7 * sin);
  const s = Math.min(1.4 * R * sin, m * 0.42);
  // Wider boxes stretch the ring into an ellipse.
  const [rx, ry] = [R + (W - m) * 0.3, R + (H - m) * 0.3];
  const at = (a) => [cx + rx * Math.cos(a), cy + ry * Math.sin(a)];
  items.forEach((it, i) => {
    const [x, y] = at(-Math.PI / 2 + (i * 2 * Math.PI) / n);
    b.shape('ellipse', { x: x - s / 2, y: y - s / 2, w: s, h: s }, i, it.text, it.lines, { area: ELLIPSE_AREA });
  });
  const aw = Math.min(s * 0.32, 2 * R * sin - s);
  if (aw < 4) return;
  items.forEach((_, i) => {
    const a = -Math.PI / 2 + ((i + 0.5) * 2 * Math.PI) / n;
    const [x, y] = at(a);
    const ah = aw * 0.7;
    b.mark('rightArrow', { x: x - aw / 2, y: y - ah / 2, w: aw, h: ah, rot: Math.round(((a * 180) / Math.PI + 90 + 360) % 360) }, i, { opacity: 0.6 });
  });
}

// Where text goes in a circle, in fractions of its box (render.js textInset for ellipses).
const ELLIPSE_AREA = { l: 0.15, t: 0.15, r: 0.15, b: 0.15 };

/** A tree of boxes, each over the ones under it, joined by elbow lines (SmartArt's "Hierarchy" and "Organization Chart"). */
function hierarchy(b, W, H) {
  const roots = outlineTree(b.d.nodes);
  if (!roots.length) return;
  const leaves = (it) => (it.children.length ? it.children.reduce((s, c) => s + leaves(c), 0) : 1);
  const depth = (it) => 1 + Math.max(0, ...it.children.map(depth));
  const total = roots.reduce((s, r) => s + leaves(r), 0);
  const levels = Math.max(...roots.map(depth));
  const bh = H / (levels + (levels - 1) * 0.6);
  const slot = W / total;
  const bw = Math.min(slot * 0.86, bh * 3);
  const yOf = (lvl) => lvl * bh * 1.6;
  // Place each subtree in its run of leaf slots; a box centres over its children.
  const place = (it, lvl, from) => {
    const n = leaves(it);
    let cx = from + (n * slot) / 2;
    const kids = [];
    let at = from;
    for (const c of it.children) {
      kids.push(place(c, lvl + 1, at));
      at += leaves(c) * slot;
    }
    if (kids.length) cx = (kids[0].cx + kids[kids.length - 1].cx) / 2;
    const y = yOf(lvl);
    b.shape('roundRect', { x: cx - bw / 2, y, w: bw, h: bh }, b.d.colors === 'colorful' ? lvl : 0, it.node.text);
    if (kids.length) {
      const mid = y + bh * 1.3;
      b.line(cx, y + bh, cx, mid);
      if (kids.length > 1) b.line(kids[0].cx, mid, kids[kids.length - 1].cx, mid);
      for (const k of kids) b.line(k.cx, mid, k.cx, yOf(lvl + 1));
    }
    return { cx };
  };
  let from = 0;
  for (const r of roots) {
    place(r, 0, from);
    from += leaves(r) * slot;
  }
}

/** Slices of a triangle, the first at the top (SmartArt's "Basic Pyramid"). */
function pyramid(b, W, H) {
  const items = topItems(b.d.nodes);
  const n = items.length;
  if (!n) return;
  const pw = Math.min(W, H * 1.4);
  const cx = W / 2;
  const g = Math.min(4, H * 0.015);
  const half = (y) => (y / H) * (pw / 2);
  items.forEach((it, i) => {
    const top = (i * H) / n + (i ? g / 2 : 0);
    const bottom = ((i + 1) * H) / n - (i < n - 1 ? g / 2 : 0);
    const wb = half(bottom);
    const a = (wb - half(top)) / (2 * wb);
    const paths = [{ cmds: [['M', a, 0], ['L', 1 - a, 0], ['L', 1, 1], ['L', 0, 1], ['Z']] }];
    // Text in the slice's middle; the top one's in its wider lower part.
    const textRect = i ? { l: a / 2 + 0.02, t: 0, r: a / 2 + 0.02, b: 0 } : { l: 0.22, t: 0.42, r: 0.22, b: 0 };
    b.shape('custom', { x: cx - wb, y: top, w: 2 * wb, h: bottom - top }, i, it.text, it.lines, { paths, textRect });
  });
}

/** Overlapping see-through circles, the text in each away from the middle (SmartArt's "Basic Venn"). */
function venn(b, W, H) {
  const items = topItems(b.d.nodes);
  const n = items.length;
  if (!n) return;
  const m = Math.min(W, H);
  const [cx, cy] = [W / 2, H / 2];
  const sin = n > 1 ? Math.sin(Math.PI / n) : 0;
  const R = n > 1 ? m / 2 / (1 + 1.4 * sin) : 0;
  const s = n > 1 ? 2.8 * R * sin : m * 0.8;
  const angle = (i) => (n === 2 ? Math.PI * (1 - i) : -Math.PI / 2 + (i * 2 * Math.PI) / n);
  const [rx, ry] = n > 1 ? [R + (W - m) * 0.2, R + (H - m) * 0.2] : [0, 0];
  const outline = b.d.style === 'outline';
  items.forEach((it, i) => {
    const a = angle(i);
    const [x, y] = [cx + rx * Math.cos(a), cy + ry * Math.sin(a)];
    b.mark('ellipse', { x: x - s / 2, y: y - s / 2, w: s, h: s }, i, { opacity: outline ? 1 : 0.5 });
  });
  items.forEach((it, i) => {
    const a = angle(i);
    // The text sits in the part of the circle the others don't cover.
    const r = n > 1 ? R + s * 0.18 : 0;
    const tw = s * (n > 1 ? 0.56 : 0.7);
    const th = s * (n > 1 ? 0.4 : 0.6);
    const [x, y] = [cx + (rx + r - R) * Math.cos(a), cy + (ry + r - R) * Math.sin(a)];
    b.text({ x: x - tw / 2, y: y - th / 2, w: tw, h: th }, it.text, it.lines, { group: 'main', align: 'center', valign: 'middle' });
  });
}

/** Boxes in a square grid, two by two for four (SmartArt's "Basic Matrix"). */
function matrix(b, W, H) {
  const items = topItems(b.d.nodes);
  const n = items.length;
  if (!n) return;
  const cols = Math.ceil(Math.sqrt(n));
  const rows = Math.ceil(n / cols);
  const g = Math.min(W, H) * 0.03;
  let cw = (W - (cols - 1) * g) / cols;
  let ch = (H - (rows - 1) * g) / rows;
  cw = Math.min(cw, ch * 1.5);
  ch = Math.min(ch, cw);
  const x0 = (W - cols * cw - (cols - 1) * g) / 2;
  const y0 = (H - rows * ch - (rows - 1) * g) / 2;
  items.forEach((it, i) => {
    b.shape('roundRect', { x: x0 + (i % cols) * (cw + g), y: y0 + Math.floor(i / cols) * (ch + g), w: cw, h: ch }, i, it.text, it.lines);
  });
}

/** The first node in a circle in the middle, the others round it, joined to it by lines (SmartArt's "Basic Radial"). */
function radial(b, W, H) {
  const nodes = normalizeLevels(b.d.nodes);
  if (!nodes.length) return;
  // Round the middle: the first node's children and any other top-level nodes, each with what's under it as bullets.
  const outer = [];
  for (const n of nodes.slice(1)) {
    if (n.level <= 1) outer.push({ text: n.text, lines: [] });
    else outer[outer.length - 1].lines.push({ text: n.text, depth: n.level - 2 });
  }
  const k = outer.length;
  const m = Math.min(W, H);
  const [cx, cy] = [W / 2, H / 2];
  if (!k) {
    const c = m * 0.5;
    b.shape('ellipse', { x: cx - c / 2, y: cy - c / 2, w: c, h: c }, 0, nodes[0].text, [], { area: ELLIPSE_AREA });
    return;
  }
  let s = m * 0.26;
  let R = m / 2 - s / 2;
  if (k > 1 && 2 * R * Math.sin(Math.PI / k) * 0.86 < s) {
    s = (m * Math.sin(Math.PI / k) * 0.86) / (1 + Math.sin(Math.PI / k) * 0.86);
    R = m / 2 - s / 2;
  }
  const c = Math.min(s * 1.3, (R - s / 2) * 2 * 0.86);
  const [rx, ry] = [R + (W - m) * 0.3, R + (H - m) * 0.3];
  const spots = outer.map((_, i) => {
    const a = -Math.PI / 2 + (i * 2 * Math.PI) / k;
    return { a, x: cx + rx * Math.cos(a), y: cy + ry * Math.sin(a) };
  });
  // Lines first, so the circles cover their ends.
  for (const p of spots) b.line(cx + (c / 2) * Math.cos(p.a), cy + (c / 2) * Math.sin(p.a), p.x - (s / 2) * Math.cos(p.a), p.y - (s / 2) * Math.sin(p.a));
  b.shape('ellipse', { x: cx - c / 2, y: cy - c / 2, w: c, h: c }, 0, nodes[0].text, [], { area: ELLIPSE_AREA });
  outer.forEach((o, i) => b.shape('ellipse', { x: spots[i].x - s / 2, y: spots[i].y - s / 2, w: s, h: s }, i + 1, o.text, o.lines, { area: ELLIPSE_AREA }));
}

const LAYOUTS = { list: basicList, vlist: verticalList, process, chevron: chevrons, cycle, hierarchy, pyramid, venn, matrix, radial };

/** The elements a diagram element is drawn with by its own layout, in its box's coordinates. */
export function layoutDiagram(el) {
  const d = el.diagram || {};
  const b = new Builder({ ...el, diagram: { ...d, nodes: normalizeLevels(d.nodes || []) } });
  (LAYOUTS[d.layout] || basicList)(b, Math.max(1, el.w), Math.max(1, el.h));
  return b.finish();
}

// ------------------------------------------------------------------ what's drawn

/** PowerPoint's drawing of a diagram (see deck.js), stretched from the box it was made for to `w` × `h`. */
function stretchDrawn(drawn, w, h) {
  const sx = drawn.w > 0 ? w / drawn.w : 1;
  const sy = drawn.h > 0 ? h / drawn.h : 1;
  return drawn.elements.map((c) => (c.type === 'line'
    ? { ...c, x1: c.x1 * sx, y1: c.y1 * sy, x2: c.x2 * sx, y2: c.y2 * sy }
    : { ...c, x: c.x * sx, y: c.y * sy, w: c.w * sx, h: c.h * sy }));
}

/** What a diagram element shows, in its box's coordinates: PowerPoint's drawing of it until it's edited, else its layout. */
export function diagramChildren(el) {
  const drawn = el.diagram?.drawn;
  return drawn?.elements?.length ? stretchDrawn(drawn, el.w, el.h) : layoutDiagram(el);
}

/**
 * The diagram's elements on the slide, turned with it: plain elements with
 * ids of their own (see diagramAsShapes).
 */
export function diagramShapes(el, newId = (i) => `${el.id}-${i}`) {
  const a = ((el.rot || 0) * Math.PI) / 180;
  const [cx, cy] = [el.x + el.w / 2, el.y + el.h / 2];
  const turn = (x, y) => (a ? [cx + (x - cx) * Math.cos(a) - (y - cy) * Math.sin(a), cy + (x - cx) * Math.sin(a) + (y - cy) * Math.cos(a)] : [x, y]);
  return diagramChildren(el).map((c, i) => {
    const copy = JSON.parse(JSON.stringify(c));
    copy.id = newId(i);
    if (c.type === 'line') {
      [copy.x1, copy.y1] = turn(el.x + c.x1, el.y + c.y1);
      [copy.x2, copy.y2] = turn(el.x + c.x2, el.y + c.y2);
    } else {
      const [x, y] = turn(el.x + c.x + c.w / 2, el.y + c.y + c.h / 2);
      Object.assign(copy, { x: x - c.w / 2, y: y - c.h / 2, rot: ((c.rot || 0) + (el.rot || 0)) % 360 });
    }
    return copy;
  });
}

/**
 * A diagram as plain shapes ("Convert to shapes", and saving to PowerPoint
 * when it can't go as SmartArt): its elements on the slide, grouped (in the
 * group it's in, if any), each with its animations (playing together) and link.
 */
export function diagramAsShapes(el, { groupId = el.groupId || `diagram-${el.id}`, newId } = {}) {
  return diagramShapes(el, newId).map((c, i) => {
    c.groupId = groupId;
    if (el.link) c.link = el.link;
    for (const kind of ANIM_KINDS) {
      if (el[kind]?.effect && el[kind].effect !== 'none') c[kind] = { ...el[kind], start: i ? 'with' : el[kind].start };
    }
    return c;
  });
}
