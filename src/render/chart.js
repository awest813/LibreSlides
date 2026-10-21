import { resolveColor, isDarkColor } from '../model/themes.js';
import { fontStack } from '../model/fonts.js';

/**
 * Charts, drawn as inline SVG filling the element's box at 1px = 1pt. Text
 * widths are estimated rather than measured, so drawing is pure and quick
 * enough for thumbnails, and looks the same in the editor, the slideshow and
 * print. Layout follows PowerPoint's defaults: title on top, the legend
 * beside or under the plot, a rounded value scale with gridlines, the first
 * bar of a bar chart at the bottom and pies starting at twelve o'clock.
 */

const SVG = 'http://www.w3.org/2000/svg';
const svgEl = (tag, attrs = {}) => {
  const el = document.createElementNS(SVG, tag);
  for (const [k, v] of Object.entries(attrs)) if (v != null) el.setAttribute(k, v);
  return el;
};
const r1 = (n) => Math.round(n * 10) / 10;

// ------------------------------------------------------------------ colours

const ACCENTS = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];

/** `hex` moved `amount` of the way towards `to` (both "#rrggbb"). */
function mix(hex, to, amount) {
  const p = (h) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16));
  if (!/^#[0-9a-f]{6}$/i.test(hex)) return hex;
  const [a, b] = [p(hex), p(to)];
  return `#${a.map((v, i) => Math.round(v + (b[i] - v) * amount).toString(16).padStart(2, '0')).join('')}`;
}

/** The theme's accents in turn, darker then lighter on each later round (as PowerPoint does). */
export function paletteColor(i, theme) {
  const base = resolveColor(`theme:${ACCENTS[i % ACCENTS.length]}`, theme, '#4472c4');
  const round = Math.floor(i / ACCENTS.length);
  if (!round) return base;
  return mix(base, round % 2 ? '#000000' : '#ffffff', Math.min(0.6, 0.2 + 0.1 * round));
}

/** The colour of series `i`: its own, or the theme's. */
export const seriesColor = (chart, i, theme) => resolveColor(chart.series?.[i]?.color, theme, null) || paletteColor(i, theme);

/** The colour of slice `i` of a pie or doughnut. */
export const pointColor = (chart, i, theme) => resolveColor(chart.series?.[0]?.pointColors?.[i], theme, null) || paletteColor(i, theme);

const contrastText = (fill) => (isDarkColor(fill) ? '#ffffff' : '#262626');

// ------------------------------------------------------------------ numbers

const numberFormats = new Map();
/** A number for a label, with up to `digits` decimals and the reader's digit grouping. */
export function formatNumber(v, digits = 2) {
  if (!numberFormats.has(digits)) numberFormats.set(digits, new Intl.NumberFormat(undefined, { maximumFractionDigits: digits }));
  return numberFormats.get(digits).format(v);
}

const isNum = (v) => typeof v === 'number' && Number.isFinite(v);

/**
 * A rounded scale covering min…max in at most `maxTicks` steps of 1, 2 or 5
 * (times a power of ten), like PowerPoint's automatic axis.
 */
export function niceScale(min, max, maxTicks = 6) {
  if (!isNum(min) || !isNum(max)) [min, max] = [0, 1];
  if (max - min < 1e-12) {
    if (max === 0) max = 1;
    else if (max > 0) min = 0;
    else max = 0;
  }
  const rough = (max - min) / Math.max(1, maxTicks);
  const mag = 10 ** Math.floor(Math.log10(rough));
  const step = [1, 2, 5, 10].map((m) => m * mag).find((s) => s >= rough * 0.999) || 10 * mag;
  const lo = Math.floor(min / step + 1e-9) * step;
  const hi = Math.ceil(max / step - 1e-9) * step;
  return { min: lo, max: hi === lo ? lo + step : hi, step, digits: Math.max(0, -Math.floor(Math.log10(step) + 1e-9)) };
}

const ticksOf = (scale) => {
  const out = [];
  for (let v = scale.min; v <= scale.max + scale.step / 1e6; v += scale.step) out.push(Math.abs(v) < scale.step / 1e6 ? 0 : v);
  return out;
};

// ------------------------------------------------------------------ text

/** Rough width of a label (no layout is available in thumbnails and tests). */
const textWidth = (s, size) => String(s).length * size * 0.55;

/** `s` cut short with an ellipsis to fit `width`. */
function fit(s, width, size) {
  s = String(s ?? '');
  if (textWidth(s, size) <= width + 0.5) return s;
  const n = Math.max(0, Math.floor(width / (size * 0.55)) - 1);
  return n ? `${s.slice(0, n).trimEnd()}…` : '';
}

function label(x, y, s, { size, anchor = 'middle', fill = null, opacity = null, cls = null, weight = null } = {}) {
  const t = svgEl('text', { x: r1(x), y: r1(y + size * 0.35), 'font-size': r1(size), 'text-anchor': anchor, fill, 'fill-opacity': opacity, class: cls, 'font-weight': weight });
  t.textContent = s;
  return t;
}

// ------------------------------------------------------------------ data

/** How many categories a chart has (series may be longer or shorter). */
const categoryCount = (chart) => (chart.categories?.length ? chart.categories.length : Math.max(0, ...(chart.series || []).map((s) => s.values?.length || 0)));

/**
 * Each series' segments as { from, to, value } per category: from zero for
 * clustered charts, piled up for stacked ones (bars pile positives up and
 * negatives down; areas add up), as percentages of the category's total for
 * 100% stacked ones.
 */
function segments(chart, n, { stacked, percent, separateSigns }) {
  const series = chart.series || [];
  const out = series.map(() => []);
  for (let i = 0; i < n; i++) {
    const total = percent ? series.reduce((t, s) => t + Math.abs(isNum(s.values?.[i]) ? s.values[i] : 0), 0) : 0;
    let up = 0;
    let down = 0;
    series.forEach((s, k) => {
      const raw = s.values?.[i];
      if (!isNum(raw)) {
        out[k].push(stacked ? { from: up, to: up, value: null } : null);
        return;
      }
      const v = percent ? (total ? (raw / total) * 100 : 0) : raw;
      if (!stacked) out[k].push({ from: 0, to: v, value: raw });
      else if (separateSigns && v < 0) {
        out[k].push({ from: down, to: down + v, value: raw });
        down += v;
      } else {
        out[k].push({ from: up, to: up + v, value: raw });
        up += v;
      }
    });
  }
  return out;
}

/** The value range of some segments, with zero in it where the chart kind wants it. */
function valueRange(segs, { zero = true } = {}) {
  let min = Infinity;
  let max = -Infinity;
  for (const list of segs) {
    for (const s of list) {
      if (!s || s.value == null) continue;
      min = Math.min(min, s.from, s.to);
      max = Math.max(max, s.from, s.to);
    }
  }
  if (min === Infinity) return [0, 1];
  // Like PowerPoint, lines start from zero unless their values sit far above (or below) it.
  if (zero || min <= (max * 5) / 6) min = Math.min(min, 0);
  if (zero || max >= (min * 5) / 6) max = Math.max(max, 0);
  return [min, max];
}

// ------------------------------------------------------------------ legend

function legendEntries(chart, theme) {
  const kind = chart.kind;
  if (kind === 'pie' || kind === 'doughnut') {
    return (chart.categories || []).map((name, i) => ({ name, color: pointColor(chart, i, theme), mark: 'box' }));
  }
  const mark = kind === 'line' ? 'line' : kind === 'scatter' ? 'dot' : 'box';
  return (chart.series || []).map((s, i) => ({ name: s.name || `Series ${i + 1}`, color: seriesColor(chart, i, theme), mark }));
}

/** Draw the legend at `pos` inside `area`, returning the area left for the plot. */
function drawLegend(g, entries, pos, area, size, textOpts) {
  if (!entries.length || pos === 'none') return area;
  const key = size * 0.75;
  const gap = size * 0.4;
  const rowH = size * 1.5;
  const legend = svgEl('g', { class: 'chart-legend' });
  const entry = (x, y, e, maxW) => {
    if (e.mark === 'line') legend.append(svgEl('line', { x1: r1(x), y1: r1(y), x2: r1(x + key), y2: r1(y), stroke: e.color, 'stroke-width': 2.25 }));
    else if (e.mark === 'dot') legend.append(svgEl('circle', { cx: r1(x + key / 2), cy: r1(y), r: r1(key * 0.35), fill: e.color }));
    else legend.append(svgEl('rect', { x: r1(x), y: r1(y - key / 2), width: r1(key), height: r1(key), fill: e.color }));
    legend.append(label(x + key + gap, y, fit(e.name, maxW, size), { size, anchor: 'start', ...textOpts }));
  };
  g.append(legend);
  if (pos === 'left' || pos === 'right') {
    const textW = Math.min(area.w * 0.4 - key - gap, Math.max(...entries.map((e) => textWidth(e.name, size))));
    const w = key + gap + Math.max(0, textW);
    const shown = entries.slice(0, Math.max(1, Math.floor(area.h / rowH)));
    const x = pos === 'right' ? area.x + area.w - w : area.x;
    let y = area.y + (area.h - shown.length * rowH) / 2 + rowH / 2;
    for (const e of shown) {
      entry(x, y, e, textW);
      y += rowH;
    }
    return { ...area, x: pos === 'left' ? area.x + w + size : area.x, w: area.w - w - size };
  }
  // Top or bottom: entries side by side, in as many rows as they need.
  const widths = entries.map((e) => key + gap + Math.min(textWidth(e.name, size), area.w * 0.6) + size);
  const rows = [[]];
  let used = 0;
  entries.forEach((e, i) => {
    if (used + widths[i] > area.w && rows[rows.length - 1].length) {
      rows.push([]);
      used = 0;
    }
    rows[rows.length - 1].push(i);
    used += widths[i];
  });
  const kept = rows.slice(0, Math.max(1, Math.floor((area.h * 0.4) / rowH)));
  const h = kept.length * rowH;
  let y = (pos === 'top' ? area.y : area.y + area.h - h) + rowH / 2;
  for (const row of kept) {
    const rowW = row.reduce((t, i) => t + widths[i], 0) - size;
    let x = area.x + (area.w - rowW) / 2;
    for (const i of row) {
      entry(x, y, entries[i], widths[i] - key - gap - size);
      x += widths[i];
    }
    y += rowH;
  }
  return { ...area, y: pos === 'top' ? area.y + h + size * 0.5 : area.y, h: area.h - h - size * 0.5 };
}

// ------------------------------------------------------------------ pies

function arcPath(cx, cy, r, inner, a0, a1) {
  const pt = (rad, a) => `${r1(cx + rad * Math.cos(a))},${r1(cy + rad * Math.sin(a))}`;
  const large = a1 - a0 > Math.PI ? 1 : 0;
  if (a1 - a0 >= Math.PI * 2 - 1e-6) {
    // A whole circle (or ring): two half arcs, as one arc can't go all the way round.
    const ring = (rad) => `M${pt(rad, a0)}A${r1(rad)},${r1(rad)} 0 1 1 ${pt(rad, a0 + Math.PI)}A${r1(rad)},${r1(rad)} 0 1 1 ${pt(rad, a0)}Z`;
    return inner ? `${ring(r)}${ring(inner)}` : ring(r);
  }
  if (!inner) return `M${r1(cx)},${r1(cy)}L${pt(r, a0)}A${r1(r)},${r1(r)} 0 ${large} 1 ${pt(r, a1)}Z`;
  return `M${pt(r, a0)}A${r1(r)},${r1(r)} 0 ${large} 1 ${pt(r, a1)}L${pt(inner, a1)}A${r1(inner)},${r1(inner)} 0 ${large} 0 ${pt(inner, a0)}Z`;
}

function drawPie(g, chart, area, size, theme) {
  const values = (chart.series?.[0]?.values || []).map((v) => (isNum(v) ? Math.abs(v) : 0));
  const total = values.reduce((t, v) => t + v, 0);
  if (!total || area.w <= 0 || area.h <= 0) return;
  const r = (Math.min(area.w, area.h) / 2) * 0.92;
  const inner = chart.kind === 'doughnut' ? r * 0.5 : 0;
  const cx = area.x + area.w / 2;
  const cy = area.y + area.h / 2;
  const edge = resolveColor('theme:bg', theme, '#ffffff');
  let a = -Math.PI / 2;
  const labels = [];
  values.forEach((v, i) => {
    if (!v) return;
    const sweep = (v / total) * Math.PI * 2;
    const fill = pointColor(chart, i, theme);
    g.append(svgEl('path', { class: 'chart-slice', d: arcPath(cx, cy, r, inner, a, a + sweep), fill, 'fill-rule': 'evenodd', stroke: edge, 'stroke-width': 1, 'stroke-linejoin': 'round' }));
    if (chart.dataLabels) {
      const mid = a + sweep / 2;
      const at = inner ? (r + inner) / 2 : r * 0.62;
      labels.push(label(cx + at * Math.cos(mid), cy + at * Math.sin(mid), formatNumber(chart.series[0].values[i]), { size, fill: contrastText(fill), cls: 'chart-label' }));
    }
    a += sweep;
  });
  g.append(...labels);
}

// ------------------------------------------------------------------ axes

function drawAxes(g, chart, area, size, theme, textOpts) {
  const kind = chart.kind;
  const n = categoryCount(chart);
  const series = chart.series || [];
  if (!n || !series.length || area.w <= size || area.h <= size) return;
  const horizontal = kind === 'bar';
  const stacked = (kind === 'column' || kind === 'bar' || kind === 'area') && (chart.grouping === 'stacked' || chart.grouping === 'percent');
  const percent = stacked && chart.grouping === 'percent';
  const segs = segments(chart, n, { stacked, percent, separateSigns: kind !== 'area' });
  const [vmin, vmax] = valueRange(segs, { zero: kind !== 'line' && kind !== 'scatter' });
  const ink = textOpts.fill;
  const muted = { ...textOpts, opacity: 0.75 };
  const grid = svgEl('g', { class: 'chart-grid', stroke: ink, 'stroke-opacity': 0.18, 'stroke-width': 0.75 });
  const plotG = svgEl('g', { class: 'chart-plot' });
  const labelsG = svgEl('g', { class: 'chart-labels' });
  g.append(grid, plotG, labelsG);
  const valueText = (v, scale) => (percent ? `${formatNumber(v, 0)}%` : formatNumber(v, scale.digits));

  // Scatter charts have numbers along the bottom too: their categories are X values.
  const xs = kind === 'scatter' ? (chart.categories || []).map((c, i) => (isNum(parseFloat(c)) ? parseFloat(c) : i + 1)) : null;
  while (xs && xs.length < n) xs.push(xs.length + 1);
  const cats = Array.from({ length: n }, (_, i) => String(chart.categories?.[i] ?? ''));

  if (horizontal) {
    const catW = Math.min(area.w * 0.35, Math.max(...cats.map((c) => textWidth(c, size))) + size * 0.6);
    const plot = { x: area.x + catW, y: area.y, w: area.w - catW - size * 0.8, h: area.h - size * 1.7 };
    if (plot.w <= 0 || plot.h <= 0) return;
    const probe = niceScale(vmin, vmax, 6);
    const tickW = Math.max(textWidth(valueText(probe.max, probe), size), textWidth(valueText(probe.min, probe), size)) + size * 1.2;
    const scale = niceScale(vmin, vmax, Math.max(2, Math.min(10, Math.floor(plot.w / tickW))));
    const X = (v) => plot.x + ((v - scale.min) / (scale.max - scale.min)) * plot.w;
    for (const t of ticksOf(scale)) {
      grid.append(svgEl('line', { x1: r1(X(t)), y1: r1(plot.y), x2: r1(X(t)), y2: r1(plot.y + plot.h) }));
      labelsG.append(label(X(t), plot.y + plot.h + size * 0.9, valueText(t, scale), { size, ...muted }));
    }
    const slot = plot.h / n;
    const x0 = X(Math.max(scale.min, Math.min(scale.max, 0)));
    grid.append(svgEl('line', { x1: r1(x0), y1: r1(plot.y), x2: r1(x0), y2: r1(plot.y + plot.h), 'stroke-opacity': 0.45, class: 'chart-axis' }));
    const step = Math.max(1, Math.ceil((size * 1.3) / slot));
    cats.forEach((c, i) => {
      if (i % step) return;
      // The first category is at the bottom, as in PowerPoint.
      labelsG.append(label(plot.x - size * 0.4, plot.y + plot.h - (i + 0.5) * slot, fit(c, catW - size * 0.6, size), { size, anchor: 'end', ...muted }));
    });
    drawBars(plotG, labelsG, chart, segs, { n, slot, stacked, horizontal, theme, size, textOpts, toPx: X, place: (i) => plot.y + plot.h - (i + 1) * slot });
    return;
  }

  const probe = niceScale(vmin, vmax, Math.max(2, Math.floor(area.h / (size * 2.6))));
  const tickW = Math.max(...ticksOf(probe).map((t) => textWidth(valueText(t, probe), size))) + size * 0.6;
  const catH = size * 1.7;
  const plot = { x: area.x + tickW, y: area.y + size * 0.5, w: area.w - tickW - size * 0.5, h: area.h - catH - size * 0.5 };
  if (plot.w <= 0 || plot.h <= 0) return;
  const scale = niceScale(vmin, vmax, Math.max(2, Math.min(10, Math.floor(plot.h / (size * 2.4)))));
  const Y = (v) => plot.y + plot.h - ((v - scale.min) / (scale.max - scale.min)) * plot.h;
  for (const t of ticksOf(scale)) {
    grid.append(svgEl('line', { x1: r1(plot.x), y1: r1(Y(t)), x2: r1(plot.x + plot.w), y2: r1(Y(t)) }));
    labelsG.append(label(plot.x - size * 0.4, Y(t), valueText(t, scale), { size, anchor: 'end', ...muted }));
  }
  const y0 = Y(Math.max(scale.min, Math.min(scale.max, 0)));
  grid.append(svgEl('line', { x1: r1(plot.x), y1: r1(y0), x2: r1(plot.x + plot.w), y2: r1(y0), 'stroke-opacity': 0.45, class: 'chart-axis' }));

  if (xs) {
    // Scatter: a second value scale across.
    const [xmin, xmax] = valueRange([xs.map((x) => ({ from: x, to: x, value: x }))], { zero: false });
    const xr = niceScale(xmin, xmax, Math.max(2, Math.min(10, Math.floor(plot.w / (size * 5)))));
    const X = (v) => plot.x + ((v - xr.min) / (xr.max - xr.min)) * plot.w;
    for (const t of ticksOf(xr)) labelsG.append(label(X(t), plot.y + plot.h + size * 0.9, formatNumber(t, xr.digits), { size, ...muted }));
    series.forEach((s, k) => {
      const color = seriesColor(chart, k, theme);
      xs.forEach((x, i) => {
        const v = s.values?.[i];
        if (!isNum(v)) return;
        plotG.append(svgEl('circle', { class: 'chart-marker', cx: r1(X(x)), cy: r1(Y(v)), r: r1(Math.max(2, size * 0.3)), fill: color }));
        if (chart.dataLabels) labelsG.append(label(X(x), Y(v) - size * 0.9, formatNumber(v), { size, cls: 'chart-label', ...textOpts }));
      });
    });
    return;
  }

  // Categories along the bottom: in the middle of their slot, or on the
  // gridlines for area charts (which reach from edge to edge).
  const edgeToEdge = kind === 'area' && n > 1;
  const slot = plot.w / (edgeToEdge ? n - 1 : n);
  const cx = (i) => plot.x + (edgeToEdge ? i * slot : (i + 0.5) * slot);
  const widest = Math.max(...cats.map((c) => textWidth(c, size)));
  const step = Math.max(1, Math.ceil((Math.min(widest, size * 12) + size * 0.6) / slot));
  cats.forEach((c, i) => {
    if (i % step) return;
    labelsG.append(label(cx(i), plot.y + plot.h + size * 0.9, fit(c, slot * step - size * 0.3, size), { size, ...muted }));
  });

  if (kind === 'column') {
    drawBars(plotG, labelsG, chart, segs, { n, slot, stacked, horizontal: false, theme, size, textOpts, toPx: Y, place: (i) => plot.x + i * slot });
  } else if (kind === 'area') {
    drawAreas(plotG, labelsG, chart, segs, { cx, Y, y0, stacked, theme, size, textOpts });
  } else {
    drawLines(plotG, labelsG, chart, segs, { cx, Y, theme, size, textOpts });
  }
}

/**
 * Columns or bars: `place(i)` is where category i's slot starts (its left
 * edge, or its top edge for bars) and `toPx` turns a value into a position.
 */
function drawBars(g, labels, chart, segs, { n, slot, stacked, horizontal, theme, size, textOpts, toPx, place }) {
  const series = chart.series;
  // PowerPoint's default gap between categories is 150% of a bar's width.
  const barW = slot / ((stacked ? 1 : series.length) + 1.5);
  const groupW = stacked ? barW : barW * series.length;
  series.forEach((s, k) => {
    const fill = seriesColor(chart, k, theme);
    for (let i = 0; i < n; i++) {
      const seg = segs[k][i];
      if (!seg || seg.value == null) continue;
      // The first series is on the left of its group, or at the bottom for bars.
      const offset = (slot - groupW) / 2 + (stacked ? 0 : k * barW);
      const pos = horizontal ? place(i) + slot - offset - barW : place(i) + offset;
      const a = toPx(seg.from);
      const b = toPx(seg.to);
      const lo = Math.min(a, b);
      const len = Math.abs(b - a);
      g.append(svgEl('rect', horizontal
        ? { class: 'chart-bar', x: r1(lo), y: r1(pos), width: r1(len), height: r1(barW), fill }
        : { class: 'chart-bar', x: r1(pos), y: r1(lo), width: r1(barW), height: r1(len), fill }));
      if (!chart.dataLabels) continue;
      const text = formatNumber(seg.value);
      const mid = pos + barW / 2;
      if (stacked) {
        if (len < size * 1.1) continue;
        labels.append(label(horizontal ? lo + len / 2 : mid, horizontal ? mid : lo + len / 2, text, { size, fill: contrastText(fill), cls: 'chart-label' }));
      } else if (horizontal) {
        const neg = seg.to < seg.from;
        labels.append(label(neg ? lo - size * 0.3 : lo + len + size * 0.3, mid, text, { size, anchor: neg ? 'end' : 'start', cls: 'chart-label', ...textOpts }));
      } else {
        const neg = seg.to < seg.from;
        labels.append(label(mid, neg ? lo + len + size * 0.8 : lo - size * 0.8, text, { size, cls: 'chart-label', ...textOpts }));
      }
    }
  });
}

function drawLines(g, labels, chart, segs, { cx, Y, theme, size, textOpts }) {
  chart.series.forEach((s, k) => {
    const color = seriesColor(chart, k, theme);
    // Blank values leave a gap in the line.
    let run = [];
    const runs = [run];
    segs[k].forEach((seg, i) => {
      if (!seg || seg.value == null) runs.push((run = []));
      else run.push([cx(i), Y(seg.to), seg.value]);
    });
    for (const pts of runs) {
      if (pts.length > 1) g.append(svgEl('polyline', { class: 'chart-line', points: pts.map(([x, y]) => `${r1(x)},${r1(y)}`).join(' '), fill: 'none', stroke: color, 'stroke-width': 2.25, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));
      for (const [x, y, v] of pts) {
        if (chart.markers || pts.length === 1) g.append(svgEl('circle', { class: 'chart-marker', cx: r1(x), cy: r1(y), r: r1(Math.max(2, size * 0.3)), fill: color }));
        if (chart.dataLabels) labels.append(label(x, y - size * 0.9, formatNumber(v), { size, cls: 'chart-label', ...textOpts }));
      }
    }
  });
}

function drawAreas(g, labels, chart, segs, { cx, Y, y0, stacked, theme, size, textOpts }) {
  chart.series.forEach((s, k) => {
    const fill = seriesColor(chart, k, theme);
    const top = segs[k].map((seg, i) => [cx(i), Y(seg ? seg.to : 0)]);
    const bottom = stacked ? segs[k].map((seg, i) => [cx(i), Y(seg ? seg.from : 0)]).reverse() : [[top[top.length - 1][0], y0], [top[0][0], y0]];
    g.append(svgEl('polygon', { class: 'chart-area', points: [...top, ...bottom].map(([x, y]) => `${r1(x)},${r1(y)}`).join(' '), fill }));
    if (!chart.dataLabels) return;
    segs[k].forEach((seg, i) => {
      if (!seg || seg.value == null) return;
      const y = stacked ? (Y(seg.from) + Y(seg.to)) / 2 : Y(seg.to) - size * 0.9;
      labels.append(label(cx(i), y, formatNumber(seg.value), { size, cls: 'chart-label', ...(stacked ? { fill: contrastText(fill) } : textOpts) }));
    });
  });
}

// ------------------------------------------------------------------ the chart

/** The SVG for a chart element. */
export function renderChart(el, ctx) {
  const { theme } = ctx;
  const W = Math.max(1, el.w || 1);
  const H = Math.max(1, el.h || 1);
  const chart = el.chart || {};
  const ink = resolveColor(el.style?.color || 'theme:text', theme, '#000000');
  const size = Math.max(4, Math.min(el.style?.size || 12, H / 5));
  const svg = svgEl('svg', {
    class: 'el-chart-svg', width: W, height: H, viewBox: `0 0 ${W} ${H}`,
    'font-family': fontStack(el.style?.font || theme.fonts.body), fill: ink,
    style: 'position:absolute;left:0;top:0;overflow:hidden',
  });
  const textOpts = { fill: ink };
  // The chart area's background and border.
  const fill = resolveColor(el.fill, theme, 'transparent');
  const stroke = el.stroke?.width && el.stroke.color && el.stroke.color !== 'none' ? resolveColor(el.stroke.color, theme) : null;
  if (fill !== 'transparent' || stroke) {
    const sw = stroke ? el.stroke.width : 0;
    svg.append(svgEl('rect', { x: sw / 2, y: sw / 2, width: Math.max(0, W - sw), height: Math.max(0, H - sw), fill: fill === 'transparent' ? 'none' : fill, stroke: stroke || 'none', 'stroke-width': sw || null }));
  }

  const pad = size * 0.6;
  let area = { x: pad, y: pad, w: W - pad * 2, h: H - pad * 2 };
  if (chart.title) {
    const ts = size * 1.5;
    svg.append(label(W / 2, area.y + ts / 2, fit(chart.title, area.w, ts), { size: ts, cls: 'chart-title' }));
    area = { ...area, y: area.y + ts * 1.5, h: area.h - ts * 1.5 };
  }
  const g = svgEl('g');
  svg.append(g);
  area = drawLegend(g, legendEntries(chart, theme), chart.legend || 'none', area, size, { fill: ink, opacity: 0.85 });
  if (chart.kind === 'pie' || chart.kind === 'doughnut') drawPie(g, chart, area, size, theme);
  else drawAxes(g, chart, area, size, theme, textOpts);
  return svg;
}
