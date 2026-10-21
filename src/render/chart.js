import { resolveColor, isDarkColor } from '../model/themes.js';
import { fontStack } from '../model/fonts.js';
import { formatValue, isGeneral, isPercentFormat } from './number-format.js';
import { svgEl } from './svg.js';

/**
 * Charts, drawn as inline SVG filling the element's box at 1px = 1pt. Text
 * widths are estimated rather than measured, so drawing is pure and quick
 * enough for thumbnails, and looks the same in the editor, the slideshow and
 * print. Layout follows PowerPoint's defaults: title on top, the legend
 * beside or under the plot, a rounded value scale with gridlines, the first
 * bar of a bar chart at the bottom and pies starting at twelve o'clock.
 */

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

// ------------------------------------------------------------------ kinds

/** Chart kinds whose series can be drawn as one another (a combination chart). */
export const COMBO_KINDS = ['column', 'line', 'area'];

/** How series `k` is drawn: as its own kind in a combination chart, else as the chart's. */
export function seriesKind(chart, k) {
  const own = chart.series?.[k]?.kind;
  return COMBO_KINDS.includes(chart.kind) && COMBO_KINDS.includes(own) ? own : chart.kind;
}

/** Whether series `k` is plotted against the secondary value axis (on the right). */
export const onSecondaryAxis = (chart, k) => COMBO_KINDS.includes(chart.kind) && chart.series?.[k]?.secondary === true;

/** Whether a scatter chart joins its points with lines (and marks them). */
export const scatterLines = (chart) => chart.lines === true;
export const showsMarkers = (chart) => (chart.kind === 'scatter' ? chart.markers !== false : chart.markers === true);

// ------------------------------------------------------------------ numbers

const numberFormats = new Map();
/** A number for a label, with up to `digits` decimals and the reader's digit grouping. */
export function formatNumber(v, digits = 2) {
  if (!numberFormats.has(digits)) numberFormats.set(digits, new Intl.NumberFormat(undefined, { maximumFractionDigits: digits }));
  return numberFormats.get(digits).format(v);
}

const isNum = (v) => typeof v === 'number' && Number.isFinite(v);

/** What a data label shows: some of the category, the value and (pies) the share. */
export const LABEL_PARTS = ['category', 'value', 'percent'];
export function labelParts(chart) {
  const list = Array.isArray(chart.labelContent) ? LABEL_PARTS.filter((p) => chart.labelContent.includes(p)) : [];
  return list.length ? list : ['value'];
}

/**
 * A data label's text: its parts in PowerPoint's order, in the labels'
 * number format, on lines of their own in pies and doughnuts (as
 * LibreOffice shows them) and separated by commas elsewhere.
 */
export function labelText(chart, value, category, share = null) {
  const fmt = chart.labelFormat ?? chart.numberFormat;
  const out = [];
  for (const part of labelParts(chart)) {
    if (part === 'category') out.push(String(category ?? ''));
    else if (part === 'value') out.push(formatValue(value, fmt));
    else if (share != null) out.push(formatValue(share, isPercentFormat(fmt) ? fmt : '0%'));
  }
  return out.filter((s) => s !== '').join(chart.kind === 'pie' || chart.kind === 'doughnut' ? '\n' : ', ');
}

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
  return { min: lo, max: hi === lo ? lo + step : hi, step, digits: stepDigits(step) };
}

const stepDigits = (step) => Math.min(10, Math.max(0, -Math.floor(Math.log10(step) + 1e-9), (String(+step.toPrecision(12)).split('.')[1] || '').length));

/**
 * The scale of a value axis: automatic, with the axis's own minimum, maximum
 * and major unit where it sets them (`k` turns them into the scale's units:
 * 100 for 100% charts, whose axis runs 0…1 in PowerPoint).
 */
export function axisScale(lo, hi, maxTicks, axis = {}, k = 1) {
  const own = (v) => (isNum(v) ? v * k : null);
  const min = own(axis?.min);
  const max = own(axis?.max);
  const unit = own(axis?.unit) > 0 ? own(axis.unit) : null;
  const auto = niceScale(min ?? lo, max ?? hi, maxTicks);
  if (min == null && max == null && !unit) return auto;
  const step = unit || auto.step;
  const a = min ?? (unit ? Math.floor(Math.min(lo, max ?? lo) / unit + 1e-9) * unit : auto.min);
  const b = max ?? (unit ? Math.ceil(Math.max(hi, a + unit) / unit - 1e-9) * unit : auto.max);
  // A unit too small for the range would draw thousands of gridlines.
  if (!(b > a) || (b - a) / step > 500) return auto;
  return { min: a, max: b, step, digits: stepDigits(step) };
}

const ticksOf = (scale) => {
  const out = [];
  for (let v = scale.min; v <= scale.max + scale.step / 1e6; v += scale.step) out.push(Math.abs(v) < scale.step / 1e6 ? 0 : v);
  return out;
};

/** A tick label of a value axis in its number format (or as a percentage for 100% charts). */
function tickText(fmt, v, scale, percent) {
  if (percent) return isPercentFormat(fmt) ? formatValue(v / 100, fmt) : `${formatNumber(v, scale.digits)}%`;
  return isGeneral(fmt) ? formatNumber(v, scale.digits) : formatValue(v, fmt);
}

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

/** A text centred on y; lines broken at "\n" are centred together. */
function label(x, y, s, { size, anchor = 'middle', fill = null, opacity = null, cls = null, weight = null } = {}) {
  const lines = String(s).split('\n');
  const top = y - ((lines.length - 1) * size * 1.2) / 2;
  const t = svgEl('text', { x: r1(x), y: r1(top + size * 0.35), 'font-size': r1(size), 'text-anchor': anchor, fill, 'fill-opacity': opacity, class: cls, 'font-weight': weight });
  if (lines.length === 1) t.textContent = s;
  else t.append(...lines.map((line, i) => Object.assign(svgEl('tspan', { x: r1(x), dy: i ? r1(size * 1.2) : null }), { textContent: line })));
  return t;
}

// ------------------------------------------------------------------ data

/** How many categories a chart has (series may be longer or shorter). */
const categoryCount = (chart) => (chart.categories?.length ? chart.categories.length : Math.max(0, ...(chart.series || []).map((s) => s.values?.length || 0)));

/**
 * Each series' segments as { from, to, value } per category: from zero for
 * clustered charts, piled up for stacked ones (bars pile positives up and
 * negatives down; areas and lines add up), as percentages of the category's
 * total for 100% stacked ones.
 */
function segments(series, n, { stacked, percent, separateSigns }) {
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

/**
 * The series drawn together: one layer per kind and value axis, areas under
 * columns under lines. The chart's grouping stacks the layers of its own kind.
 */
function layersOf(chart, n) {
  const series = chart.series || [];
  const byKey = new Map();
  series.forEach((s, k) => {
    const kind = seriesKind(chart, k);
    const secondary = onSecondaryAxis(chart, k);
    const key = `${kind}/${secondary}`;
    if (!byKey.has(key)) byKey.set(key, { kind, secondary, idx: [] });
    byKey.get(key).idx.push(k);
  });
  const order = ['area', 'column', 'bar', 'line', 'scatter'];
  const stacking = chart.grouping === 'stacked' || chart.grouping === 'percent';
  return [...byKey.values()].sort((a, b) => order.indexOf(a.kind) - order.indexOf(b.kind)).map((layer) => {
    const stacked = stacking && layer.kind === chart.kind && layer.kind !== 'scatter';
    const percent = stacked && chart.grouping === 'percent';
    const segs = segments(layer.idx.map((k) => series[k]), n, { stacked, percent, separateSigns: layer.kind === 'column' || layer.kind === 'bar' });
    return { ...layer, stacked, percent, segs };
  });
}

/** The value range of some layers. */
function layersRange(layers) {
  let lo = Infinity;
  let hi = -Infinity;
  for (const l of layers) {
    const [a, b] = valueRange(l.segs, { zero: l.kind !== 'line' && l.kind !== 'scatter' });
    lo = Math.min(lo, a);
    hi = Math.max(hi, b);
  }
  return lo === Infinity ? [0, 1] : [lo, hi];
}

// ------------------------------------------------------------------ legend

function legendEntries(chart, theme) {
  const kind = chart.kind;
  if (kind === 'pie' || kind === 'doughnut') {
    return (chart.categories || []).map((name, i) => ({ name, color: pointColor(chart, i, theme), mark: 'box' }));
  }
  const markOf = (k) => {
    const own = seriesKind(chart, k);
    if (own === 'scatter') return scatterLines(chart) && !showsMarkers(chart) ? 'line' : 'dot';
    return own === 'line' ? 'line' : 'box';
  };
  return (chart.series || []).map((s, i) => ({ name: s.name || `Series ${i + 1}`, color: seriesColor(chart, i, theme), mark: markOf(i) }));
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

/**
 * Pies, and doughnuts with a ring per series (the first series innermost,
 * as in PowerPoint). Labels sit inside the slices, or just outside them
 * when they're placed at the outside end.
 */
function drawPie(g, chart, area, size, theme, textOpts) {
  if (area.w <= 0 || area.h <= 0) return;
  const doughnut = chart.kind === 'doughnut';
  const rings = (chart.series || []).slice(0, doughnut ? Infinity : 1);
  const outside = !doughnut && chart.labelPosition === 'outEnd';
  const r = (Math.min(area.w, area.h) / 2) * (outside && chart.dataLabels ? 0.75 : 0.92);
  const hole = doughnut ? r * (Math.min(90, Math.max(10, chart.holeSize ?? 50)) / 100) : 0;
  const band = (r - hole) / Math.max(1, rings.length);
  const cx = area.x + area.w / 2;
  const cy = area.y + area.h / 2;
  const edge = resolveColor('theme:bg', theme, '#ffffff');
  const labels = [];
  rings.forEach((s, k) => {
    const values = (s.values || []).map((v) => (isNum(v) ? Math.abs(v) : 0));
    const total = values.reduce((t, v) => t + v, 0);
    if (!total) return;
    const inner = hole + k * band;
    const outer = inner + band;
    // Where labels go along the radius: PowerPoint's centre, inside end or best fit.
    const at = doughnut ? (inner + outer) / 2 : { ctr: r * 0.5, inEnd: r * 0.8 }[chart.labelPosition] ?? r * 0.62;
    let a = -Math.PI / 2;
    values.forEach((v, i) => {
      if (!v) return;
      const sweep = (v / total) * Math.PI * 2;
      const fill = pointColor(chart, i, theme);
      g.append(svgEl('path', { class: 'chart-slice', d: arcPath(cx, cy, outer, inner, a, a + sweep), fill, 'fill-rule': 'evenodd', stroke: edge, 'stroke-width': 1, 'stroke-linejoin': 'round' }));
      if (chart.dataLabels) {
        const mid = a + sweep / 2;
        const text = labelText(chart, s.values[i], chart.categories?.[i], v / total);
        if (outside) {
          const right = Math.cos(mid) >= 0;
          labels.push(label(cx + (r + size * 0.4) * Math.cos(mid), cy + (r + size * 0.7) * Math.sin(mid), text, { size, anchor: Math.abs(Math.cos(mid)) < 0.2 ? 'middle' : right ? 'start' : 'end', cls: 'chart-label', ...textOpts }));
        } else labels.push(label(cx + at * Math.cos(mid), cy + at * Math.sin(mid), text, { size, fill: contrastText(fill), cls: 'chart-label' }));
      }
      a += sweep;
    });
  });
  g.append(...labels);
}

// ------------------------------------------------------------------ axes

/** Axis titles on the left, bottom and right of `area`, returning the area left inside them. */
function drawAxisTitles(g, { left, bottom, right }, area, size, textOpts) {
  let { x, y, w, h } = area;
  const room = size * 1.5;
  const turned = (cx, cy, s, angle, length) => {
    const t = label(0, 0, fit(s, length, size), { size, cls: 'chart-axis-title', ...textOpts });
    t.setAttribute('transform', `translate(${r1(cx)},${r1(cy)}) rotate(${angle})`);
    g.append(t);
  };
  if (bottom && h > room * 2) {
    g.append(label(x + w / 2, y + h - size * 0.6, fit(bottom, w, size), { size, cls: 'chart-axis-title', ...textOpts }));
    h -= room;
  }
  if (left && w > room * 2) {
    turned(x + size * 0.6, y + h / 2, left, -90, h);
    x += room;
    w -= room;
  }
  if (right && w > room * 2) {
    turned(x + w - size * 0.6, y + h / 2, right, -90, h);
    w -= room;
  }
  return { x, y, w, h };
}

function drawAxes(g, chart, area, size, theme, textOpts) {
  const kind = chart.kind;
  const n = categoryCount(chart);
  const series = chart.series || [];
  if (!n || !series.length || area.w <= size || area.h <= size) return;
  const horizontal = kind === 'bar';
  const layers = layersOf(chart, n);
  let main = layers.filter((l) => !l.secondary);
  let second = layers.filter((l) => l.secondary);
  if (!main.length) [main, second] = [second, []];
  const percent = main.some((l) => l.percent);
  const percent2 = second.some((l) => l.percent);
  const [vmin, vmax] = layersRange(main);
  const valueAxis = chart.valueAxis || {};
  const catTitle = chart.categoryAxis?.title || '';
  area = drawAxisTitles(g, {
    left: horizontal ? catTitle : valueAxis.title,
    bottom: horizontal ? valueAxis.title : catTitle,
    right: second.length ? chart.secondaryAxis?.title : '',
  }, area, size, textOpts);
  if (area.w <= size || area.h <= size) return;
  const gridlines = valueAxis.gridlines !== false;
  const ink = textOpts.fill;
  const muted = { ...textOpts, opacity: 0.75 };
  const grid = svgEl('g', { class: 'chart-grid', stroke: ink, 'stroke-opacity': 0.18, 'stroke-width': 0.75 });
  const plotG = svgEl('g', { class: 'chart-plot' });
  const labelsG = svgEl('g', { class: 'chart-labels' });
  g.append(grid, plotG, labelsG);
  const valueText = (v, scale, pct = percent, fmt = chart.numberFormat) => tickText(fmt, v, scale, pct);

  // Scatter charts have numbers along the bottom too: their categories are X values.
  const xs = kind === 'scatter' ? (chart.categories || []).map((c, i) => (isNum(parseFloat(c)) ? parseFloat(c) : i + 1)) : null;
  while (xs && xs.length < n) xs.push(xs.length + 1);
  const cats = Array.from({ length: n }, (_, i) => String(chart.categories?.[i] ?? ''));

  if (horizontal) {
    const catW = Math.min(area.w * 0.35, Math.max(...cats.map((c) => textWidth(c, size))) + size * 0.6);
    const plot = { x: area.x + catW, y: area.y, w: area.w - catW - size * 0.8, h: area.h - size * 1.7 };
    if (plot.w <= 0 || plot.h <= 0) return;
    const probe = axisScale(vmin, vmax, 6, valueAxis, percent ? 100 : 1);
    const tickW = Math.max(textWidth(valueText(probe.max, probe), size), textWidth(valueText(probe.min, probe), size)) + size * 1.2;
    const scale = axisScale(vmin, vmax, Math.max(2, Math.min(10, Math.floor(plot.w / tickW))), valueAxis, percent ? 100 : 1);
    const X = (v) => plot.x + ((v - scale.min) / (scale.max - scale.min)) * plot.w;
    for (const t of ticksOf(scale)) {
      if (gridlines) grid.append(svgEl('line', { x1: r1(X(t)), y1: r1(plot.y), x2: r1(X(t)), y2: r1(plot.y + plot.h) }));
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
    for (const layer of main) drawBars(plotG, labelsG, chart, layer, { n, slot, horizontal, theme, size, textOpts, toPx: X, place: (i) => plot.y + plot.h - (i + 1) * slot });
    return;
  }

  const ticksFor = (h) => Math.max(2, Math.min(10, Math.floor(h / (size * 2.4))));
  const probe = axisScale(vmin, vmax, Math.max(2, Math.floor(area.h / (size * 2.6))), valueAxis, percent ? 100 : 1);
  const tickW = Math.max(...ticksOf(probe).map((t) => textWidth(valueText(t, probe), size))) + size * 0.6;
  // A secondary value axis on the right, with a scale of its own.
  const [smin, smax] = second.length ? layersRange(second) : [0, 1];
  const secondAxis = chart.secondaryAxis || {};
  const secondFormat = secondAxis.numberFormat || 'General';
  const probe2 = second.length ? axisScale(smin, smax, Math.max(2, Math.floor(area.h / (size * 2.6))), secondAxis, percent2 ? 100 : 1) : null;
  const tickW2 = probe2 ? Math.max(...ticksOf(probe2).map((t) => textWidth(valueText(t, probe2, percent2, secondFormat), size))) + size * 0.6 : size * 0.5;
  const catH = size * 1.7;
  const plot = { x: area.x + tickW, y: area.y + size * 0.5, w: area.w - tickW - tickW2, h: area.h - catH - size * 0.5 };
  if (plot.w <= 0 || plot.h <= 0) return;
  const scale = axisScale(vmin, vmax, ticksFor(plot.h), valueAxis, percent ? 100 : 1);
  const toY = (s) => (v) => plot.y + plot.h - ((v - s.min) / (s.max - s.min)) * plot.h;
  const Y = toY(scale);
  for (const t of ticksOf(scale)) {
    if (gridlines) grid.append(svgEl('line', { x1: r1(plot.x), y1: r1(Y(t)), x2: r1(plot.x + plot.w), y2: r1(Y(t)) }));
    labelsG.append(label(plot.x - size * 0.4, Y(t), valueText(t, scale), { size, anchor: 'end', ...muted }));
  }
  const y0 = Y(Math.max(scale.min, Math.min(scale.max, 0)));
  grid.append(svgEl('line', { x1: r1(plot.x), y1: r1(y0), x2: r1(plot.x + plot.w), y2: r1(y0), 'stroke-opacity': 0.45, class: 'chart-axis' }));
  let Y2 = Y;
  if (probe2) {
    const scale2 = axisScale(smin, smax, ticksFor(plot.h), secondAxis, percent2 ? 100 : 1);
    Y2 = toY(scale2);
    for (const t of ticksOf(scale2)) labelsG.append(label(plot.x + plot.w + size * 0.4, Y2(t), valueText(t, scale2, percent2, secondFormat), { size, anchor: 'start', ...muted }));
  }

  if (xs) {
    // Scatter: a second value scale across.
    const [xmin, xmax] = valueRange([xs.map((x) => ({ from: x, to: x, value: x }))], { zero: false });
    const xr = niceScale(xmin, xmax, Math.max(2, Math.min(10, Math.floor(plot.w / (size * 5)))));
    const X = (v) => plot.x + ((v - xr.min) / (xr.max - xr.min)) * plot.w;
    for (const t of ticksOf(xr)) labelsG.append(label(X(t), plot.y + plot.h + size * 0.9, formatNumber(t, xr.digits), { size, ...muted }));
    for (const layer of main) drawLines(plotG, labelsG, chart, layer, { cx: (i) => X(xs[i]), Y, theme, size, textOpts });
    return;
  }

  // Categories along the bottom: in the middle of their slot, or on the
  // gridlines for area charts (which reach from edge to edge).
  const edgeToEdge = layers.every((l) => l.kind === 'area') && n > 1;
  const slot = plot.w / (edgeToEdge ? n - 1 : n);
  const cx = (i) => plot.x + (edgeToEdge ? i * slot : (i + 0.5) * slot);
  const widest = Math.max(...cats.map((c) => textWidth(c, size)));
  const step = Math.max(1, Math.ceil((Math.min(widest, size * 12) + size * 0.6) / slot));
  cats.forEach((c, i) => {
    if (i % step) return;
    labelsG.append(label(cx(i), plot.y + plot.h + size * 0.9, fit(c, slot * step - size * 0.3, size), { size, ...muted }));
  });

  for (const layer of layers) {
    const toPx = layer.secondary && second.length ? Y2 : Y;
    if (layer.kind === 'column') {
      drawBars(plotG, labelsG, chart, layer, { n, slot, horizontal: false, theme, size, textOpts, toPx, place: (i) => plot.x + i * slot });
    } else if (layer.kind === 'area') {
      drawAreas(plotG, labelsG, chart, layer, { cx, Y: toPx, y0: Math.min(plot.y + plot.h, Math.max(plot.y, toPx(0))), theme, size, textOpts });
    } else {
      drawLines(plotG, labelsG, chart, layer, { cx, Y: toPx, theme, size, textOpts });
    }
  }
}

/**
 * Columns or bars: `place(i)` is where category i's slot starts (its left
 * edge, or its top edge for bars) and `toPx` turns a value into a position.
 * Bars are as wide as the chart's gap width leaves them (150% of a bar
 * between categories by default) and overlap by its overlap.
 */
function drawBars(g, labels, chart, layer, { n, slot, horizontal, theme, size, textOpts, toPx, place }) {
  const { idx, segs, stacked } = layer;
  const gap = Math.min(500, Math.max(0, chart.gapWidth ?? 150)) / 100;
  const overlap = stacked ? 1 : Math.min(100, Math.max(-100, chart.overlap ?? 0)) / 100;
  const m = stacked ? 1 : idx.length;
  const barW = slot / (m - (m - 1) * overlap + gap);
  const groupW = barW * (1 + (m - 1) * (1 - overlap));
  // Where labels go: outside the end by default, or in the middle of stacked bars.
  const where = stacked
    ? (['inEnd', 'inBase', 'ctr'].includes(chart.labelPosition) ? chart.labelPosition : 'ctr')
    : (['inEnd', 'inBase', 'ctr', 'outEnd'].includes(chart.labelPosition) ? chart.labelPosition : 'outEnd');
  idx.forEach((k, j) => {
    const fill = seriesColor(chart, k, theme);
    for (let i = 0; i < n; i++) {
      const seg = segs[j][i];
      if (!seg || seg.value == null) continue;
      // The first series is on the left of its group, or at the bottom for bars.
      const offset = (slot - groupW) / 2 + (stacked ? 0 : j * barW * (1 - overlap));
      const pos = horizontal ? place(i) + slot - offset - barW : place(i) + offset;
      const a = toPx(seg.from);
      const b = toPx(seg.to);
      const lo = Math.min(a, b);
      const len = Math.abs(b - a);
      g.append(svgEl('rect', horizontal
        ? { class: 'chart-bar', x: r1(lo), y: r1(pos), width: r1(len), height: r1(barW), fill }
        : { class: 'chart-bar', x: r1(pos), y: r1(lo), width: r1(barW), height: r1(len), fill }));
      if (!chart.dataLabels) continue;
      if (where !== 'outEnd' && stacked && len < size * 1.1) continue;
      const text = labelText(chart, seg.value, chart.categories?.[i]);
      const mid = pos + barW / 2;
      // `dir` runs from the bar's base to its end, in pixels.
      const dir = b !== a ? Math.sign(b - a) : horizontal ? 1 : -1;
      const inside = where !== 'outEnd';
      const style = inside ? { fill: contrastText(fill) } : textOpts;
      const pad = horizontal ? size * 0.3 : size * 0.8;
      const along = { outEnd: b + dir * pad, inEnd: b - dir * pad, inBase: a + dir * pad, ctr: (a + b) / 2 }[where];
      if (horizontal) {
        const forward = where === 'outEnd' || where === 'inBase' ? dir > 0 : dir < 0;
        labels.append(label(along, mid, text, { size, anchor: where === 'ctr' ? 'middle' : forward ? 'start' : 'end', cls: 'chart-label', ...style }));
      } else labels.append(label(mid, along, text, { size, cls: 'chart-label', ...style }));
    }
  });
}

/** A line through some points: straight, or a smooth curve through them. */
function linePath(pts, smooth) {
  const p = (x, y) => `${r1(x)},${r1(y)}`;
  if (!smooth || pts.length < 3) return `M${pts.map(([x, y]) => p(x, y)).join('L')}`;
  // Catmull-Rom splines as cubic Béziers.
  let d = `M${p(pts[0][0], pts[0][1])}`;
  for (let i = 0; i < pts.length - 1; i++) {
    const [p0, p1, p2, p3] = [pts[i - 1] || pts[i], pts[i], pts[i + 1], pts[i + 2] || pts[i + 1]];
    d += `C${p(p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6)} ${p(p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6)} ${p(p2[0], p2[1])}`;
  }
  return d;
}

/** Lines (and scatter points): `cx(i)` is where category i is across. */
function drawLines(g, labels, chart, layer, { cx, Y, theme, size, textOpts }) {
  const scatter = layer.kind === 'scatter';
  const joined = !scatter || scatterLines(chart);
  const markers = showsMarkers(chart);
  const where = ['t', 'b', 'l', 'r', 'ctr'].includes(chart.labelPosition) ? chart.labelPosition : scatter ? 'r' : 't';
  const at = { t: [0, -0.9, 'middle'], b: [0, 0.9, 'middle'], l: [-0.5, 0, 'end'], r: [0.5, 0, 'start'], ctr: [0, 0, 'middle'] }[where];
  layer.idx.forEach((k, j) => {
    const color = seriesColor(chart, k, theme);
    // Blank values leave a gap in the line.
    let run = [];
    const runs = [run];
    layer.segs[j].forEach((seg, i) => {
      if (!seg || seg.value == null) runs.push((run = []));
      else run.push([cx(i), Y(seg.to), seg.value, i]);
    });
    for (const pts of runs) {
      if (joined && pts.length > 1) g.append(svgEl('path', { class: 'chart-line', d: linePath(pts, chart.smooth === true), fill: 'none', stroke: color, 'stroke-width': 2.25, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));
      for (const [x, y, v, i] of pts) {
        if (markers || (!scatter && pts.length === 1)) g.append(svgEl('circle', { class: 'chart-marker', cx: r1(x), cy: r1(y), r: r1(Math.max(2, size * 0.3)), fill: color }));
        if (chart.dataLabels) labels.append(label(x + at[0] * size, y + at[1] * size, labelText(chart, v, chart.categories?.[i]), { size, anchor: at[2], cls: 'chart-label', ...textOpts }));
      }
    }
  });
}

function drawAreas(g, labels, chart, layer, { cx, Y, y0, theme, size, textOpts }) {
  const { stacked } = layer;
  layer.idx.forEach((k, j) => {
    const segs = layer.segs[j];
    const fill = seriesColor(chart, k, theme);
    const top = segs.map((seg, i) => [cx(i), Y(seg ? seg.to : 0)]);
    const bottom = stacked ? segs.map((seg, i) => [cx(i), Y(seg ? seg.from : 0)]).reverse() : [[top[top.length - 1][0], y0], [top[0][0], y0]];
    g.append(svgEl('polygon', { class: 'chart-area', points: [...top, ...bottom].map(([x, y]) => `${r1(x)},${r1(y)}`).join(' '), fill }));
    if (!chart.dataLabels) return;
    segs.forEach((seg, i) => {
      if (!seg || seg.value == null) return;
      const y = stacked ? (Y(seg.from) + Y(seg.to)) / 2 : Y(seg.to) - size * 0.9;
      labels.append(label(cx(i), y, labelText(chart, seg.value, chart.categories?.[i]), { size, cls: 'chart-label', ...(stacked ? { fill: contrastText(fill) } : textOpts) }));
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
  if (chart.kind === 'pie' || chart.kind === 'doughnut') drawPie(g, chart, area, size, theme, textOpts);
  else drawAxes(g, chart, area, size, theme, textOpts);
  return svg;
}
