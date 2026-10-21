import { WARPS, FOLLOW_PATH, isPlainWarp } from '../model/warps.js';
import { evaluateGeometry } from '../model/geometry.js';
import { fontStack } from '../model/fonts.js';
import { resolveColor } from '../model/themes.js';
import { gradientPaint } from './render.js';

/**
 * WordArt text warps (see model/warps.js), drawn as an SVG over the text box
 * with each character placed on its own. There are no glyph outlines to bend,
 * so a character is moved, turned, scaled and skewed to follow the warp:
 *
 * - "Follow path" warps (one path, or three for Button): each line of text
 *   runs along a path, its characters turned to the path's direction.
 * - The others come in pairs of paths, the top and bottom of an envelope:
 *   each line is stretched across it, a character at fraction u of the line
 *   standing between top(u) and bottom(u), with several lines sharing the
 *   envelope's height.
 *
 * With more paths (or pairs) than one, lines are shared out between them in
 * order, as PowerPoint does. The plain text stays in the box underneath,
 * hidden, and is shown instead while it's being edited (see app.css).
 */

const SVG = 'http://www.w3.org/2000/svg';
const svgEl = (tag, attrs = {}) => {
  const el = document.createElementNS(SVG, tag);
  for (const [k, v] of Object.entries(attrs)) if (v != null) el.setAttribute(k, v);
  return el;
};

const DEFAULT_INSET = { l: 7.2, t: 3.6, r: 7.2, b: 3.6 }; // the text box's padding (app.css)
const EMU = 12700; // geometry is worked out in EMUs, as the formulas expect

// ------------------------------------------------------------------ text

const BLOCKS = /^(P|DIV|LI|H\d|UL|OL)$/;

// What reads as one character: an accented letter or an emoji sequence stays together.
const segmenter = typeof Intl.Segmenter === 'function' ? new Intl.Segmenter(undefined, { granularity: 'grapheme' }) : null;
const graphemes = (s) => (segmenter ? Array.from(segmenter.segment(s), (g) => g.segment) : [...s]);

/** The look of a run of text: its box's style, then each enclosing tag's and inline style's. */
function charStyle(node, s) {
  const out = { ...s };
  const tag = node.nodeName;
  if (tag === 'B' || tag === 'STRONG') out.bold = true;
  if (tag === 'I' || tag === 'EM') out.italic = true;
  if (tag === 'U') out.underline = true;
  if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') out.strike = true;
  if (tag === 'SUB' || tag === 'SUP') out.size = s.size * 0.65;
  if (tag === 'FONT') {
    if (node.getAttribute('face')) out.font = node.getAttribute('face');
    if (node.getAttribute('color')) out.color = node.getAttribute('color');
  }
  for (const decl of (node.getAttribute('style') || '').split(';')) {
    const at = decl.indexOf(':');
    if (at < 0) continue;
    const prop = decl.slice(0, at).trim().toLowerCase();
    const value = decl.slice(at + 1).trim();
    if (prop === 'font-family') out.font = value.split(',')[0].replace(/["']/g, '').trim() || out.font;
    else if (prop === 'font-size' && parseFloat(value) > 0) out.size = parseFloat(value);
    else if (prop === 'color') out.color = value;
    else if (prop === 'font-weight') out.bold = value === 'bold' || +value >= 600;
    else if (prop === 'font-style') out.italic = value === 'italic' || value === 'oblique';
    else if (prop === 'text-decoration' || prop === 'text-decoration-line') {
      out.underline = value.includes('underline');
      out.strike = value.includes('line-through');
    } else if (prop === 'text-transform') out.caps = value === 'uppercase';
    else if (prop === 'font-variant') out.caps = value === 'small-caps';
    else if (prop === 'letter-spacing') out.letterSpacing = parseFloat(value) || 0;
    else if (prop === '-webkit-text-stroke') out.outline = strokeOf(value);
    else if (prop === 'text-shadow') out.shadow = value === 'none' ? null : value;
    else if (prop === 'background-image') out.gradient = /gradient\(/.test(value) ? value : null;
    else if (prop === '-webkit-text-fill-color' && value === 'transparent') out.fillNone = true;
  }
  return out;
}

/** -webkit-text-stroke ("1px #c00") as { width, color }. */
function strokeOf(value) {
  const m = /^([\d.]+)px\s+(.+)$/.exec(value);
  return m && +m[1] > 0 ? { width: +m[1], color: m[2] } : null;
}

/**
 * The text of a box as lines, each a list of characters with their looks:
 * a line per paragraph, list item or line break.
 */
export function textLines(body, base) {
  const lines = [[]];
  const breakLine = (always) => { if (always || lines.at(-1).length) lines.push([]); };
  const walk = (node, style) => {
    for (const n of node.childNodes) {
      if (n.nodeType === 3) {
        if (/^(UL|OL)$/.test(node.nodeName)) continue; // spacing between list items
        for (const ch of graphemes(n.data)) {
          if (ch === '\n') breakLine(true);
          else if (ch !== '\u200b') lines.at(-1).push({ ch: ch === '\t' ? ' ' : ch, style });
        }
      } else if (n.nodeType === 1) {
        if (n.nodeName === 'BR') {
          breakLine(true);
          continue;
        }
        const block = BLOCKS.test(n.nodeName);
        if (block) breakLine(false);
        walk(n, charStyle(n, style));
        if (block) breakLine(false);
      }
    }
  };
  walk(body, base);
  while (lines.length && !lines.at(-1).length) lines.pop();
  return lines;
}

// ------------------------------------------------------------------ measuring

let measurer;
const measured = new Map();

const fontCss = (s) => `${s.italic ? 'italic ' : ''}${s.bold ? 700 : 400} ${s.size}px ${fontStack(s.font)}`;

/**
 * Text measured in a font, as fractions of its size: how far it advances
 * and how far its ink reaches above and below the baseline. The canvas
 * measures it where there is one; elsewhere (as in tests) it's a rough guess.
 */
function measure(text, s) {
  if (measurer === undefined) measurer = typeof OffscreenCanvas === 'function' ? new OffscreenCanvas(1, 1).getContext('2d') : null;
  const font = fontCss(s);
  const key = `${font}|${text}`;
  let m = measured.get(key);
  if (!m) {
    if (measurer) {
      measurer.font = font;
      const tm = measurer.measureText(text);
      m = { width: tm.width / s.size, ascent: tm.actualBoundingBoxAscent / s.size, descent: tm.actualBoundingBoxDescent / s.size };
    } else {
      const guess = (ch) => (ch === ' ' ? 0.28 : /[iljtf.,;:'!|]/.test(ch) ? 0.3 : /[mwMW]/.test(ch) ? 0.85 : /[A-Z0-9]/.test(ch) ? 0.68 : 0.55);
      m = {
        width: [...text].reduce((sum, ch) => sum + guess(ch), 0),
        ascent: /[A-Z0-9bdfhklt]/.test(text) ? 0.72 : /\S/.test(text) ? 0.53 : 0,
        descent: /[gjpqy,;]/.test(text) ? 0.2 : 0,
      };
    }
    if (measured.size > 5000) measured.clear();
    measured.set(key, m);
  }
  return m;
}

// Text measured before a font arrived was measured in another.
globalThis.document?.fonts?.addEventListener?.('loadingdone', () => measured.clear());

/** Whether a web font face (a FontFace) is the one a style of text is drawn in. */
function faceOf(face, s) {
  if (face.family.replace(/^["']|["']$/g, '') !== s.font) return false;
  if ((face.style === 'italic' || face.style === 'oblique') !== !!s.italic) return false;
  const [lo, hi = lo] = String(face.weight).split(/\s+/).map((v) => (v === 'bold' ? 700 : v === 'normal' ? 400 : +v));
  const weight = s.bold ? 700 : 400;
  return weight >= lo && weight <= hi;
}

/**
 * The fonts (CSS font shorthands, with the characters wanted) of some text
 * that the browser hasn't loaded yet. The canvas measures with a fallback
 * until a web font (the bundled ones, a deck's embedded ones) has loaded, and
 * doesn't load it itself. (FontFaceSet.check() can't tell: Chrome says yes
 * when the fallback is there.)
 */
function fontsToLoad(chars) {
  const set = globalThis.document?.fonts;
  if (typeof set?.load !== 'function' || typeof set.forEach !== 'function') return [];
  const pending = [];
  set.forEach((face) => { if (face.status !== 'loaded' && face.status !== 'error') pending.push(face); });
  if (!pending.length) return [];
  const text = new Map(); // font → [a style, its characters]
  for (const { ch, style } of chars) {
    const font = fontCss(style);
    const seen = text.get(font);
    if (seen) seen[1] += ch;
    else text.set(font, [style, ch]);
  }
  const out = [];
  for (const [font, [style, chs]] of text) if (pending.some((face) => faceOf(face, style))) out.push([font, chs]);
  return out;
}

/** Draw a warp again once its fonts have loaded, in place of `svg`, so it's laid out with them. */
function redrawWhenLoaded(svg, fonts, redraw) {
  Promise.all(fonts.map(([font, chs]) => globalThis.document.fonts.load(font, chs))).then(() => {
    if (!svg.isConnected) return; // drawn again since, or never shown
    measured.clear();
    const fresh = redraw();
    if (fresh) svg.replaceWith(fresh);
  }).catch(() => { /* the fallback font stays */ });
}

const advance = (ch, s) => measure(ch, s).width * s.size + (s.letterSpacing || 0);

/**
 * Lay a line out flat: each character's x (from the line's start) and
 * advance, the line's width and height (from one baseline to the next), and
 * how far its ink reaches above and below the baseline.
 */
function measureLine(chars) {
  let x = 0;
  const out = chars.map(({ ch, style }) => {
    const glyph = style.caps ? ch.toUpperCase() : ch;
    const c = { ch: glyph, style, x, adv: advance(glyph, style) };
    x += c.adv;
    return c;
  });
  // Spaces at the ends take no room.
  while (out.length && out.at(-1).ch === ' ') x -= out.pop().adv;
  const line = { chars: out, width: x, height: 0, ascent: 0, descent: 0 };
  // The ink of each run of characters in one style.
  for (let i = 0; i < out.length;) {
    let j = i;
    while (j < out.length && out[j].style === out[i].style) j++;
    const { style } = out[i];
    const ink = measure(out.slice(i, j).map((c) => c.ch).join(''), style);
    line.ascent = Math.max(line.ascent, ink.ascent * style.size);
    line.descent = Math.max(line.descent, ink.descent * style.size);
    line.height = Math.max(line.height, style.size * 1.2);
    i = j;
  }
  // Text with no ink (spaces) still has a size.
  if (!line.ascent) line.ascent = out[0] ? out[0].style.size * 0.7 : 0;
  return line;
}

/** Lines too wide for `width` broken at spaces, as the box would wrap them. */
function wrapLines(lines, width) {
  const out = [];
  for (const line of lines) {
    let rest = line;
    while (rest.length) {
      let x = 0;
      let cut = -1;
      let i = 0;
      for (; i < rest.length; i++) {
        x += advance(rest[i].ch, rest[i].style);
        if (rest[i].ch === ' ') cut = i;
        if (x > width && cut > 0) break;
      }
      if (i >= rest.length) {
        out.push(rest);
        break;
      }
      out.push(rest.slice(0, cut));
      rest = rest.slice(cut + 1);
    }
    if (!line.length) out.push(line);
  }
  return out;
}

// ------------------------------------------------------------------ paths

/** A path (fractions of a w × h box) as a polyline with the distance along it at each point. */
function flatten(cmds, w, h) {
  const pts = [];
  let start = null;
  const at = (x, y) => pts.push([x * w, y * h]);
  for (const [op, ...n] of cmds) {
    const from = pts.at(-1);
    if (op === 'M') {
      at(n[0], n[1]);
      start = pts.at(-1);
    } else if (op === 'L') at(n[0], n[1]);
    else if ((op === 'C' || op === 'Q') && from) {
      const ctrl = [];
      for (let i = 0; i < n.length; i += 2) ctrl.push([n[i] * w, n[i + 1] * h]);
      const steps = 24;
      for (let k = 1; k <= steps; k++) {
        const t = k / steps;
        const u = 1 - t;
        if (op === 'Q') pts.push([0, 1].map((d) => u * u * from[d] + 2 * u * t * ctrl[0][d] + t * t * ctrl[1][d]));
        else pts.push([0, 1].map((d) => u * u * u * from[d] + 3 * u * u * t * ctrl[0][d] + 3 * u * t * t * ctrl[1][d] + t * t * t * ctrl[2][d]));
      }
    } else if (op === 'Z' && start) pts.push([...start]);
  }
  return withLengths(pts);
}

function withLengths(pts) {
  const len = [0];
  for (let i = 1; i < pts.length; i++) len.push(len[i - 1] + Math.hypot(pts[i][0] - pts[i - 1][0], pts[i][1] - pts[i - 1][1]));
  return { pts, len, total: len.at(-1) || 0 };
}

/** The point `s` along a polyline, and its direction there (radians). */
function pointAt(poly, s) {
  const { pts, len } = poly;
  if (pts.length < 2) return { x: pts[0]?.[0] || 0, y: pts[0]?.[1] || 0, angle: 0 };
  s = Math.max(0, Math.min(poly.total, s));
  let lo = 0;
  let hi = len.length - 1;
  while (hi - lo > 1) {
    const mid = (lo + hi) >> 1;
    if (len[mid] <= s) lo = mid;
    else hi = mid;
  }
  // Skip repeated points, which have no direction.
  while (hi < pts.length - 1 && len[hi] === len[lo]) hi++;
  const [a, b] = [pts[lo], pts[hi]];
  const span = len[hi] - len[lo];
  const t = span ? (s - len[lo]) / span : 0;
  return { x: a[0] + (b[0] - a[0]) * t, y: a[1] + (b[1] - a[1]) * t, angle: Math.atan2(b[1] - a[1], b[0] - a[0]) };
}

/** A polyline moved `d` to the right of its direction (down, for one running left to right). */
function offset(poly, d) {
  if (!d) return poly;
  return withLengths(poly.pts.map((p, i) => {
    const { angle } = pointAt(poly, poly.len[i] + (i === poly.pts.length - 1 ? -0.01 : 0.01));
    return [p[0] - d * Math.sin(angle), p[1] + d * Math.cos(angle)];
  }));
}

// ------------------------------------------------------------------ placing

/** A character drawn at (x, y), turned `angle` and stretched `sx` along its line. */
const along = (x, y, angle, sx) => `translate(${r(x)} ${r(y)}) rotate(${r((angle * 180) / Math.PI)}) scale(${r(sx, 1000)} 1)`;
const r = (v, k = 100) => Math.round(v * k) / k || 0;

/**
 * Lines running along a path, the first standing on it and the others below
 * (on the inside of an arch, outside one along the bottom), as LibreOffice
 * draws them.
 */
function placeAlong(poly, lines, { align, stretch }) {
  const out = [];
  lines.forEach((line, j) => {
    const d = lines.slice(0, j).reduce((sum, l) => sum + l.height, 0);
    const path = offset(poly, d);
    if (!line.width || !path.total) return;
    const sx = stretch || line.width > path.total ? path.total / line.width : 1;
    const room = path.total - line.width * sx;
    const start = align === 'center' ? room / 2 : align === 'right' ? room : 0;
    for (const c of line.chars) {
      const p = pointAt(path, start + (c.x + c.adv / 2) * sx);
      out.push({ ...c, transform: along(p.x, p.y, p.angle, sx) });
    }
  });
  return out;
}

/** Lines stretched across the envelope between two paths, sharing its height. */
function placeBetween(top, bottom, lines) {
  const out = [];
  // The point a fraction u across and v down the envelope.
  const at = (u, v) => {
    const a = pointAt(top, u * top.total);
    const b = pointAt(bottom, u * bottom.total);
    return [a.x + (b.x - a.x) * v, a.y + (b.y - a.y) * v];
  };
  lines.forEach((line, j) => {
    if (!line.width) return;
    const v0 = j / lines.length;
    const band = 1 / lines.length;
    const vBase = v0 + band * (line.ascent / (line.ascent + line.descent));
    for (const c of line.chars) {
      // (A character with no advance, a lone combining mark, is given a sliver of one.)
      const adv = Math.max(c.adv, 1e-3);
      const [u0, u1, uc] = [c.x / line.width, (c.x + adv) / line.width, (c.x + adv / 2) / line.width];
      const o = at(uc, vBase);
      const [left, right, up] = [at(u0, vBase), at(u1, vBase), at(uc, v0)];
      // Across: the line's direction there, a character's advance long; down: from the line's top to its baseline.
      const xa = [(right[0] - left[0]) / adv, (right[1] - left[1]) / adv];
      const ya = [(o[0] - up[0]) / line.ascent, (o[1] - up[1]) / line.ascent];
      out.push({ ...c, transform: `matrix(${r(xa[0], 1e4)} ${r(xa[1], 1e4)} ${r(ya[0], 1e4)} ${r(ya[1], 1e4)} ${r(o[0])} ${r(o[1])})` });
    }
  });
  return out;
}

/** Share lines out between `n` paths or envelopes, in order, as evenly as they go. */
function shareOut(lines, n) {
  const per = Math.max(1, Math.ceil(lines.length / n));
  return Array.from({ length: n }, (_, i) => lines.slice(i * per, (i + 1) * per));
}

/** The paths of a warp drawn w × h, as polylines. */
export function warpPaths(warp, w, h) {
  const def = WARPS[warp.preset];
  try {
    return evaluateGeometry(def, w * EMU, h * EMU, warp.adjust || {}).paths.map((p) => flatten(p.cmds, w, h));
  } catch {
    return [];
  }
}

/**
 * Where each character of a box's text goes under its warp: the characters,
 * each with its look and an SVG transform from a character drawn centred on
 * (0, 0) (its baseline) to its place in the w × h area.
 */
export function placeCharacters(warp, lines, { w, h, align = 'left', wrap = true }) {
  const paths = warpPaths(warp, w, h);
  if (!paths.length || w <= 0 || h <= 0) return [];
  const follow = FOLLOW_PATH.has(warp.preset);
  // Text wraps where it would in the box, or where it would run off the end of the path.
  const laid = (wrap ? wrapLines(lines, follow ? paths[0].total : w) : lines).map(measureLine);
  if (follow) {
    return shareOut(laid, paths.length).flatMap((ls, i) => placeAlong(paths[i], ls, { align, stretch: !!warp.stretch }));
  }
  const pairs = Math.floor(paths.length / 2);
  return shareOut(laid, pairs).flatMap((ls, i) => placeBetween(paths[2 * i], paths[2 * i + 1], ls));
}

// ------------------------------------------------------------------ drawing

/** A text-shadow (one or more "color dx dy blur") as an SVG filter. */
function shadowFilter(css, id) {
  const filter = svgEl('filter', { id, x: '-50%', y: '-50%', width: '200%', height: '200%', 'color-interpolation-filters': 'sRGB' });
  // Split at commas outside brackets (colours can be rgba(…)).
  const shadows = css.split(/,(?![^(]*\))/).map((sh) => sh.trim()).filter(Boolean);
  const merge = svgEl('feMerge');
  shadows.forEach((sh, i) => {
    const color = (/(rgba?\([^)]*\)|#[0-9a-f]{3,8}|\b[a-z]+\b)/i.exec(sh) || ['#000'])[0];
    const [dx = 0, dy = 0, blur = 0] = (sh.replace(color, '').match(/-?[\d.]+/g) || []).map(Number);
    filter.append(svgEl('feDropShadow', { in: 'SourceGraphic', dx, dy, stdDeviation: blur / 2, 'flood-color': color, result: `s${i}` }));
    merge.append(svgEl('feMergeNode', { in: `s${i}` }));
  });
  filter.append(merge);
  return filter;
}

/** A CSS linear or radial gradient (as the text effects write them) as a LibreSlides gradient (see deck.js). */
export function gradientOfCss(css) {
  const m = /(linear|radial)-gradient\((.*)\)\s*$/.exec(css || '');
  if (!m) return null;
  const parts = m[2].split(/,(?![^(]*\))/).map((s) => s.trim());
  const g = { type: m[1], stops: [] };
  if (g.type === 'linear' && /deg$/.test(parts[0])) g.angle = parseFloat(parts.shift());
  else if (g.type === 'radial' && /^circle|^ellipse|\bat\b/.test(parts[0])) {
    const at = /at\s+([\d.]+)%\s+([\d.]+)%/.exec(parts.shift());
    if (at) Object.assign(g, { cx: at[1] / 100, cy: at[2] / 100 });
  }
  parts.forEach((p, i) => {
    const pos = /\s([\d.]+)%$/.exec(p);
    g.stops.push({ pos: pos ? pos[1] / 100 : i / Math.max(1, parts.length - 1), color: pos ? p.slice(0, pos.index).trim() : p });
  });
  return g.stops.length ? g : null;
}

let warpCount = 0;

/**
 * The SVG drawing a box's text warped, for a box `bw` × `bh` with the text
 * box style `t` (render.js textStyle), or null when there's nothing to draw.
 */
export function renderWarp(el, body, opts) {
  const { t, theme, bw, bh } = opts;
  const warp = el.style?.warp;
  if (isPlainWarp(warp)) return null;
  const inset = el.style?.inset || DEFAULT_INSET;
  const w = bw - inset.l - inset.r;
  const h = bh - inset.t - inset.b;
  const base = { font: t.font, size: t.size, color: resolveColor(t.color, theme, '#000'), bold: t.bold, italic: t.italic, underline: t.underline, strike: t.strike };
  const lines = textLines(body, base);
  if (!lines.length) return null;
  const chars = placeCharacters(warp, lines, { w, h, align: t.align, wrap: el.style?.wrap !== false });
  const svg = svgEl('svg', { class: 'el-warp', width: bw, height: bh, overflow: 'visible', 'aria-hidden': 'true' });
  // Measured in a fallback font: laid out again when the real one has loaded (once).
  const loading = opts.retry === false ? [] : fontsToLoad(lines.flat());
  if (loading.length) redrawWhenLoaded(svg, loading, () => renderWarp(el, body, { ...opts, retry: false }));
  const defs = svgEl('defs');
  const area = svgEl('g', { transform: `translate(${inset.l} ${inset.t})` });
  svg.append(defs, area);
  // Characters sharing a shadow and a gradient are drawn as one group with them.
  const groups = new Map();
  const groupFor = (s) => {
    const key = `${s.shadow || ''}|${s.gradient || ''}`;
    if (groups.has(key)) return groups.get(key);
    const group = { g: svgEl('g') };
    if (s.shadow) {
      const id = `ls-warp-${++warpCount}`;
      defs.append(shadowFilter(s.shadow, id));
      group.g.setAttribute('filter', `url(#${id})`);
    }
    const gradient = gradientOfCss(s.gradient);
    if (gradient) {
      // The gradient runs across the whole area (close to the run it covers in a box of WordArt),
      // showing through the characters.
      const id = `ls-warp-${++warpCount}`;
      const room = { x: -w, y: -h, width: 3 * w, height: 3 * h };
      group.mask = svgEl('mask', { id, maskUnits: 'userSpaceOnUse', ...room });
      defs.append(group.mask);
      group.g.append(svgEl('rect', { ...room, fill: gradientPaint(svg, gradient, w, h, theme), mask: `url(#${id})` }));
    }
    area.append(group.g);
    groups.set(key, group);
    return group;
  };
  for (const c of chars) {
    if (c.ch === ' ') continue;
    const s = c.style;
    const group = groupFor(s);
    const glyph = (attrs) => {
      const text = svgEl('text', {
        x: 0, y: 0, transform: c.transform, 'text-anchor': 'middle',
        'font-family': fontStack(s.font), 'font-size': r(s.size), 'font-weight': s.bold ? 700 : 400, 'font-style': s.italic ? 'italic' : null,
        'text-decoration': [s.underline && 'underline', s.strike && 'line-through'].filter(Boolean).join(' ') || null,
        ...attrs,
      });
      text.textContent = c.ch;
      return text;
    };
    const outline = s.outline ? { stroke: s.outline.color, 'stroke-width': s.outline.width, 'stroke-linejoin': 'round' } : {};
    if (group.mask) {
      group.mask.append(glyph({ fill: '#fff' }));
      if (s.outline) group.g.append(glyph({ fill: 'none', ...outline }));
    } else group.g.append(glyph({ fill: s.color === 'transparent' || s.fillNone ? 'none' : s.color, ...outline }));
  }
  return svg;
}
