import { uid, CHART_KINDS } from '../model/deck.js';
import { resolveColor, backgroundBaseColor } from '../model/themes.js';
import { seriesColor, pointColor, seriesKind, onSecondaryAxis, labelParts, showsMarkers, COMBO_KINDS } from '../render/chart.js';
import { formatValue, isGeneral, isPercentFormat } from '../render/number-format.js';
import { EMU_PER_PT, kids, kid, path, desc, attr, rid, colorOf, blend } from './ooxml.js';

/**
 * PowerPoint charts (c:chartSpace parts) to and from LibreSlides chart
 * elements. Reading takes the numbers cached in the chart (the embedded
 * workbook isn't needed) with the kind, grouping, colours, title, legend and
 * data labels; writing makes a real PowerPoint chart with pptxgenjs, so it
 * stays editable in PowerPoint, Keynote and LibreOffice.
 */
export const MAX_SERIES = 200;
export const MAX_POINTS = 1000;

const isNum = (v) => typeof v === 'number' && Number.isFinite(v);
/** A c: boolean: <c:showVal/> means yes, as does a missing val attribute. */
const isOn = (el) => !!el && !['0', 'false'].includes(attr(el, 'val'));

// ------------------------------------------------------------------ reading

/** Chart kinds by the plot element that holds them; 3-D charts are read as flat ones. */
const PLOTS = {
  barChart: 'column', bar3DChart: 'column', lineChart: 'line', line3DChart: 'line', areaChart: 'area', area3DChart: 'area',
  pieChart: 'pie', pie3DChart: 'pie', ofPieChart: 'pie', doughnutChart: 'doughnut', scatterChart: 'scatter',
};
// Charts LibreSlides can't draw (radar, bubble, stock, surface): left out of imports.
const OTHER_PLOTS = new Set(['radarChart', 'bubbleChart', 'stockChart', 'surfaceChart', 'surface3DChart']);
const GROUPING = { stacked: 'stacked', percentStacked: 'percent' };
const LEGEND = { r: 'right', tr: 'right', l: 'left', t: 'top', b: 'bottom' };
const LABEL_POSITIONS = new Set(['outEnd', 'inEnd', 'ctr', 'inBase', 'bestFit', 't', 'b', 'l', 'r']);

/** The cache of a data reference (c:cat, c:val, c:tx…): numbers or strings, cached or literal. */
function cacheOf(holder) {
  if (!holder) return null;
  for (const [ref, cache] of [['numRef', 'numCache'], ['strRef', 'strCache'], ['multiLvlStrRef', 'multiLvlStrCache']]) {
    const r = kid(holder, ref);
    if (r) return kid(r, cache);
  }
  return kid(holder, 'numLit') || kid(holder, 'strLit');
}

/** The points of a data reference as strings by index (null where there's none). */
function pointsOf(holder) {
  const cache = cacheOf(holder);
  if (!cache) return { values: [], format: '' };
  // Multi-level categories: the first level is the one next to the axis.
  const list = cache.localName === 'multiLvlStrCache' ? kid(cache, 'lvl') : cache;
  const pts = kids(list, 'pt');
  const count = Math.max(+attr(kid(cache, 'ptCount'), 'val') || 0, ...pts.map((p) => +attr(p, 'idx') + 1 || 0));
  const out = new Array(Math.min(count, MAX_POINTS)).fill(null);
  for (const pt of pts) {
    const i = +attr(pt, 'idx');
    if (i >= 0 && i < out.length) out[i] = kid(pt, 'v')?.textContent ?? null;
  }
  return { values: out, format: kid(cache, 'formatCode')?.textContent || '' };
}

const toNumber = (v) => (v == null || v === '' || !Number.isFinite(+v) ? null : +v);

/** A number format: an element's own code, or the data's where it's linked to the source. */
function formatOf(numFmt, source = '') {
  if (!numFmt) return source;
  return attr(numFmt, 'sourceLinked') === '1' ? source || attr(numFmt, 'formatCode') || '' : attr(numFmt, 'formatCode') || '';
}

/** Category labels: numbers (dates among them) in their axis's or data's format, text as it is. */
function categoryLabels(holder, date1904, axis) {
  const { values, format } = pointsOf(holder);
  const code = formatOf(kid(axis, 'numFmt'), format);
  return values.map((v) => {
    if (v == null) return '';
    if (!isGeneral(code) && v !== '' && Number.isFinite(+v)) return formatValue(+v, code, { date1904 });
    return String(v);
  });
}

const textOf = (holder) => {
  const rich = kid(holder, 'rich');
  if (rich) return kids(rich, 'p').map((p) => desc(p, 't').map((t) => t.textContent).join('')).join(' ').trim();
  const v = kid(holder, 'v');
  if (v) return v.textContent.trim();
  return pointsOf(holder).values.filter((x) => x != null).join(' ').trim();
};

/** A fill's colour: solid, or the first stop of a gradient. */
function fillColor(spPr, scheme) {
  if (!spPr) return null;
  const solid = kid(spPr, 'solidFill');
  if (solid) return colorOf(solid, scheme);
  const grad = kid(spPr, 'gradFill');
  return grad ? colorOf(desc(grad, 'gs')[0], scheme) : null;
}
const lineColor = (spPr, scheme) => colorOf(kid(kid(spPr, 'ln'), 'solidFill'), scheme);

/** The colour PowerPoint gives a series (or slice) that doesn't set one: the deck's accents in turn. */
const accent = (i, scheme) => (scheme[`accent${(i % 6) + 1}`] ? `#${scheme[`accent${(i % 6) + 1}`].toLowerCase()}` : null);

function seriesColorOf(ser, kind, i, scheme) {
  const spPr = kid(ser, 'spPr');
  const marker = kid(kid(ser, 'marker'), 'spPr');
  let color;
  if (kind === 'line') color = lineColor(spPr, scheme) || fillColor(marker, scheme);
  else if (kind === 'scatter') color = fillColor(marker, scheme) || lineColor(spPr, scheme) || fillColor(spPr, scheme);
  else color = fillColor(spPr, scheme);
  const idx = +attr(kid(ser, 'idx'), 'val');
  return color || accent(Number.isFinite(idx) ? idx : i, scheme);
}

/**
 * What the data labels of a plot (or its series) show: some of the category
 * name, value and percentage, where, and in what number format; null for none.
 * A label of its own (c:dLbl) overrides the plot's for its point.
 */
function labelsOf(plot, ser, n) {
  const dLbls = kid(ser, 'dLbls') || kid(plot, 'dLbls');
  if (!dLbls || isOn(kid(dLbls, 'delete'))) return null;
  const own = kids(dLbls, 'dLbl').filter((d) => !isOn(kid(d, 'delete')));
  const everyPoint = own.length >= n;
  const flag = (name) => (!everyPoint && isOn(kid(dLbls, name))) || own.some((d) => isOn(kid(d, name)));
  const content = [['category', 'showCatName'], ['value', 'showVal'], ['percent', 'showPercent']].filter(([, name]) => flag(name)).map(([part]) => part);
  if (!content.length) return null;
  // Labels of their own for every point leave the plot's settings unused (pptxgenjs writes a centred one there).
  const pos = attr(kid(own[0], 'dLblPos'), 'val') || (everyPoint ? null : attr(kid(dLbls, 'dLblPos'), 'val'));
  return { content, position: LABEL_POSITIONS.has(pos) ? pos : null, numFmt: kid(dLbls, 'numFmt') || kid(own[0], 'numFmt') };
}

/** The axes of a plot area by id. */
const axesById = (plotArea) => new Map(['valAx', 'catAx', 'dateAx', 'serAx'].flatMap((n) => kids(plotArea, n)).map((ax) => [attr(kid(ax, 'axId'), 'val'), ax]));

/** A plot's value axis (a scatter chart's upright one) and category axis (its X axis). */
function axesOf(plot, axes) {
  const own = kids(plot, 'axId').map((a) => axes.get(attr(a, 'val'))).filter(Boolean);
  const values = own.filter((a) => a.localName === 'valAx');
  if (plot.localName === 'scatterChart') {
    const across = (a) => ['b', 't'].includes(attr(kid(a, 'axPos'), 'val'));
    return { val: values.find((a) => !across(a)) || values[1] || null, cat: values.find(across) || values[0] || null };
  }
  return { val: values[0] || null, cat: own.find((a) => a.localName === 'catAx' || a.localName === 'dateAx') || null };
}

const axisTitle = (ax) => (kid(ax, 'title') ? textOf(kid(kid(ax, 'title'), 'tx')) || 'Axis Title' : '');

/** A value axis's own scale, title and gridlines (PowerPoint draws none unless they're there). */
function valueAxisOf(ax, { gridlines = true } = {}) {
  const out = {};
  if (!ax) return out;
  const scaling = kid(ax, 'scaling');
  const min = toNumber(attr(kid(scaling, 'min'), 'val'));
  const max = toNumber(attr(kid(scaling, 'max'), 'val'));
  const unit = toNumber(attr(kid(ax, 'majorUnit'), 'val'));
  if (min != null) out.min = min;
  if (max != null) out.max = max;
  if (unit > 0) out.unit = unit;
  if (axisTitle(ax)) out.title = axisTitle(ax);
  if (gridlines && !kid(ax, 'majorGridlines')) out.gridlines = false;
  return out;
}

const keepFormat = (code) => (isGeneral(code) ? null : code.slice(0, 100));
const isEmpty = (o) => !Object.keys(o).length;

/** The chart's text style from its axes, legend, the whole chart or a pie's first data label (in that order). */
function textStyleOf(space, chartEl, ctx) {
  const plotArea = kid(chartEl, 'plotArea');
  const pieLabel = ['pieChart', 'pie3DChart', 'doughnutChart', 'ofPieChart'].map((n) => kid(plotArea, n)).find(Boolean);
  const sources = [...kids(plotArea, 'catAx'), ...kids(plotArea, 'valAx'), ...kids(plotArea, 'dateAx'), kid(chartEl, 'legend'), space, path(pieLabel, 'ser', 'dLbls', 'dLbl') || path(pieLabel, 'dLbls', 'dLbl')];
  const runs = sources.map((s) => desc(kid(s, 'txPr'), 'defRPr')[0]).filter(Boolean);
  const pick = (fn) => runs.map(fn).find((v) => v != null) ?? null;
  const size = pick((r) => (attr(r, 'sz') ? +attr(r, 'sz') / 100 : null));
  const color = pick((r) => colorOf(kid(r, 'solidFill'), ctx.scheme));
  const face = pick((r) => attr(kid(r, 'latin'), 'typeface'));
  const style = { size: Math.round((size || 10) * ctx.k * 10) / 10, font: face === '+mj-lt' ? ctx.fonts.major : face && !face.startsWith('+') ? face : ctx.fonts.minor };
  if (color) style.color = color;
  return style;
}

/** Scatter series: one shared list of X values, with blanks where a series has no point. */
function mergeScatter(list) {
  const same = list.every((s) => s.x.length === list[0].x.length && s.x.every((x, i) => x === list[0].x[i]));
  if (same) return { categories: list[0].x.map((x) => (x == null ? '' : String(x))), values: list.map((s) => s.y) };
  const xs = [...new Set(list.flatMap((s) => s.x.filter(isNum)))].sort((a, b) => a - b).slice(0, MAX_POINTS);
  return {
    categories: xs.map(String),
    values: list.map((s) => xs.map((x) => {
      const i = s.x.lastIndexOf(x);
      return i >= 0 ? s.y[i] ?? null : null;
    })),
  };
}

const kindOf = (plot) => (PLOTS[plot.localName] === 'column' && attr(kid(plot, 'barDir'), 'val') === 'bar' ? 'bar' : PLOTS[plot.localName]);

/**
 * A scatter chart's lines and markers: series join their points unless their
 * line has no fill (or the style is markers only), and are marked unless
 * every series' marker is "none".
 */
function readScatterStyle(chart, plot, sers) {
  const style = attr(kid(plot, 'scatterStyle'), 'val') || 'marker';
  const joined = (ser) => {
    const ln = path(ser, 'spPr', 'ln');
    return ln ? !kid(ln, 'noFill') : !['marker', 'none'].includes(style);
  };
  if (sers.some(joined)) chart.lines = true;
  if (sers.length && sers.every((ser) => attr(path(ser, 'marker', 'symbol'), 'val') === 'none')) chart.markers = false;
  if (chart.lines && (sers.some((ser) => isOn(kid(ser, 'smooth'))) || /^smooth/.test(style))) chart.smooth = true;
}

/** The chart data of a c:chartSpace, or null for a chart LibreSlides can't draw. */
export function parseChart(space, ctx) {
  const chartEl = kid(space, 'chart');
  const plotArea = kid(chartEl, 'plotArea');
  const plots = [...(plotArea?.children || [])].filter((c) => PLOTS[c.localName] || OTHER_PLOTS.has(c.localName));
  const first = plots[0];
  if (!first || !PLOTS[first.localName]) return null;
  const kind = kindOf(first);
  const date1904 = isOn(kid(space, 'date1904'));
  const pie = kind === 'pie' || kind === 'doughnut';
  const axes = axesById(plotArea);
  const mainAxes = axesOf(first, axes);
  // Combination charts: columns, lines and areas keep their own kind (and
  // value axis); other mixes are read as the first kind, with every series.
  const sers = plots.filter((p) => PLOTS[p.localName]).flatMap((p) => {
    const combo = COMBO_KINDS.includes(kind) && COMBO_KINDS.includes(kindOf(p));
    const val = axesOf(p, axes).val;
    return kids(p, 'ser').map((ser) => ({ ser, plot: p, kind: combo && kindOf(p) !== kind ? kindOf(p) : null, secondary: combo && !!val && val !== mainAxes.val }));
  });
  const chosen = (kind === 'pie' ? sers.slice(0, 1) : sers).slice(0, MAX_SERIES);

  let categories = [];
  const series = chosen.map(({ ser, kind: own }, i) => ({
    name: textOf(kid(ser, 'tx')) || `Series ${i + 1}`,
    values: pointsOf(kid(ser, 'val') || kid(ser, 'yVal')).values.map(toNumber),
    color: seriesColorOf(ser, own || kind, i, ctx.scheme),
  }));
  if (kind === 'scatter') {
    const merged = mergeScatter(chosen.map(({ ser }, i) => ({ x: pointsOf(kid(ser, 'xVal')).values.map(toNumber), y: series[i].values })));
    categories = merged.categories;
    series.forEach((s, i) => { s.values = merged.values[i]; });
  } else {
    const withCats = chosen.find(({ ser }) => kid(ser, 'cat'));
    categories = withCats ? categoryLabels(kid(withCats.ser, 'cat'), date1904, mainAxes.cat) : [];
  }
  chosen.forEach(({ kind: own, secondary }, i) => {
    if (own) series[i].kind = own;
    if (secondary) series[i].secondary = true;
  });
  // Every series as long as the categories (PowerPoint numbers them 1, 2, 3… if there are none).
  const n = Math.min(MAX_POINTS, Math.max(categories.length, ...series.map((s) => s.values.length)));
  while (categories.length < n) categories.push(String(categories.length + 1));
  for (const s of series) {
    s.values = Array.from({ length: n }, (_, i) => (isNum(s.values[i]) ? s.values[i] : null));
    if (!s.color) delete s.color;
  }

  const chart = { kind, categories, series, legend: 'none', dataLabels: false, title: '' };
  if (['column', 'bar', 'area', 'line'].includes(kind)) chart.grouping = GROUPING[attr(kid(first, 'grouping'), 'val')] || 'clustered';
  if (pie && series[0]) {
    // Slices: their own colours, else the accents in turn (or all the series' colour if they don't vary).
    const ser = chosen[0].ser;
    const vary = isOn(kid(first, 'varyColors')) || !kid(first, 'varyColors');
    const own = {};
    for (const dPt of kids(ser, 'dPt')) own[+attr(kid(dPt, 'idx'), 'val')] = fillColor(kid(dPt, 'spPr'), ctx.scheme);
    const plain = fillColor(kid(ser, 'spPr'), ctx.scheme);
    series[0].pointColors = categories.map((_, i) => own[i] || (vary ? accent(i, ctx.scheme) : plain) || accent(i, ctx.scheme) || '#4472c4');
    for (const s of series) delete s.color;
  }
  const hole = toNumber(attr(kid(first, 'holeSize'), 'val'));
  if (kind === 'doughnut' && hole != null && hole !== 50) chart.holeSize = Math.min(90, Math.max(10, hole));
  // Lines (a combination chart's too) are marked as their first series is.
  const lineSer = chosen.find((s) => (s.kind || kind) === 'line');
  if (lineSer) chart.markers = attr(path(lineSer.ser, 'marker', 'symbol'), 'val') !== 'none';
  else if (kind === 'line') chart.markers = false;
  if (kind === 'scatter') readScatterStyle(chart, first, chosen.map(({ ser }) => ser));
  else if (chosen.some(({ ser }) => isOn(kid(ser, 'smooth')))) chart.smooth = true;

  // Columns and bars: their gap width and overlap (a stacked chart's bars always overlap).
  const bars = plots.find((p) => PLOTS[p.localName] === 'column');
  if (bars) {
    const gap = toNumber(attr(kid(bars, 'gapWidth'), 'val'));
    const overlap = toNumber(attr(kid(bars, 'overlap'), 'val'));
    if (gap != null && gap !== 150) chart.gapWidth = Math.min(500, Math.max(0, gap));
    if (overlap && !GROUPING[attr(kid(bars, 'grouping'), 'val')]) chart.overlap = Math.min(100, Math.max(-100, overlap));
  }

  // Data labels, and the number formats of the value axis and the labels.
  const labels = chosen.length ? labelsOf(chosen[0].plot, chosen[0].ser, categories.length) : null;
  chart.dataLabels = !!labels;
  const dataFormat = chosen.length ? pointsOf(kid(chosen[0].ser, 'val') || kid(chosen[0].ser, 'yVal')).format : '';
  const axisFormat = formatOf(kid(mainAxes.val, 'numFmt'), dataFormat);
  if (!pie && keepFormat(axisFormat)) chart.numberFormat = keepFormat(axisFormat);
  if (labels) {
    if (labels.content.join() !== 'value') chart.labelContent = labels.content;
    if (labels.position) chart.labelPosition = labels.position;
    const labelFormat = keepFormat(formatOf(labels.numFmt, dataFormat)) || 'General';
    if (labelFormat !== (chart.numberFormat || 'General')) chart.labelFormat = labelFormat;
  }

  // Axes: their scale, titles and gridlines.
  if (!pie) {
    const valueAxis = valueAxisOf(mainAxes.val);
    if (!isEmpty(valueAxis)) chart.valueAxis = valueAxis;
    if (axisTitle(mainAxes.cat)) chart.categoryAxis = { title: axisTitle(mainAxes.cat) };
    const secondary = chosen.find((s) => s.secondary);
    const secondAxis = secondary && axesOf(secondary.plot, axes).val;
    const second = secondary ? valueAxisOf(secondAxis, { gridlines: false }) : {};
    const secondFormat = secondary && keepFormat(formatOf(kid(secondAxis, 'numFmt'), pointsOf(kid(secondary.ser, 'val')).format));
    if (secondFormat) second.numberFormat = secondFormat;
    if (!isEmpty(second)) chart.secondaryAxis = second;
  }

  // The title: its own text, or PowerPoint's automatic one (the series name for a single series).
  const titleEl = kid(chartEl, 'title');
  const auto = series.length === 1 ? series[0].name : 'Chart Title';
  if (titleEl) chart.title = textOf(kid(titleEl, 'tx')) || auto;
  else if (attr(kid(chartEl, 'autoTitleDeleted'), 'val') === '0' && series.length === 1) chart.title = auto;

  const legend = kid(chartEl, 'legend');
  if (legend) chart.legend = LEGEND[attr(kid(legend, 'legendPos'), 'val') || 'r'] || 'right';
  return chart;
}

/** A chart graphic frame as a chart element, or null to leave it out. */
export async function readChart({ pkg, node, rels, box, ctx }) {
  const rel = rels?.[rid(desc(node, 'chart')[0], 'id')];
  if (!box || !rel || rel.external) return null;
  const space = (await pkg.doc(rel.path))?.documentElement;
  if (space?.localName !== 'chartSpace') return null;
  const chart = parseChart(space, ctx);
  if (!chart) return null;
  // The chart area's own background and border (usually none).
  const spPr = kid(space, 'spPr');
  const ln = kid(spPr, 'ln');
  const border = lineColor(spPr, ctx.scheme);
  return {
    id: uid(), type: 'chart', x: box.x, y: box.y, w: box.w, h: box.h, rot: 0, chart,
    fill: fillColor(spPr, ctx.scheme) || 'none',
    stroke: border ? { color: border, width: Math.round(((+attr(ln, 'w') || 9525) / EMU_PER_PT) * ctx.k * 10) / 10, dash: 'solid' } : { color: 'none', width: 1, dash: 'solid' },
    style: textStyleOf(space, kid(space, 'chart'), ctx),
  };
}

// ------------------------------------------------------------------ writing

const TYPES = { column: 'bar', bar: 'bar', line: 'line', area: 'area', pie: 'pie', doughnut: 'doughnut', scatter: 'scatter' };
const LEGEND_POS = { right: 'r', left: 'l', top: 't', bottom: 'b' };
const hex = (c) => (/^#[0-9a-f]{6}$/i.test(c || '') ? c.slice(1).toUpperCase() : null);
const xmlText = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/**
 * Data label positions each kind of plot takes (as pptxgenjs checks them).
 * Outside the end is clustered bars' default, so it needn't be written.
 */
const LABEL_SPOTS = { bar: ['ctr', 'inBase', 'inEnd'], pie: ['bestFit', 'ctr', 'inEnd', 'outEnd'], line: ['t', 'b', 'l', 'r', 'ctr'], scatter: ['t', 'b', 'l', 'r', 'ctr'] };
const labelSpot = (type, pos) => (LABEL_SPOTS[type]?.includes(pos) ? pos : undefined);

/** The series by how they're plotted (kind and value axis), the chart's own kind first. */
export function plotGroups(chart) {
  const groups = new Map();
  (chart.series || []).forEach((s, k) => {
    const kind = seriesKind(chart, k);
    const secondary = onSecondaryAxis(chart, k);
    const key = `${kind}/${secondary}`;
    if (!groups.has(key)) groups.set(key, { kind, secondary, idx: [] });
    groups.get(key).idx.push(k);
  });
  const main = (g) => (g.kind === chart.kind && !g.secondary ? 1 : 0);
  return [...groups.values()].sort((a, b) => main(b) - main(a));
}

/** A value axis's pptxgenjs options: its scale, gridlines and title. */
function valueAxisOptions(axis = {}, gridlines, grid) {
  return {
    valAxisMinVal: isNum(axis.min) ? axis.min : undefined,
    valAxisMaxVal: isNum(axis.max) ? axis.max : undefined,
    valAxisMajorUnit: axis.unit > 0 ? axis.unit : undefined,
    valGridLine: gridlines ? { color: grid, size: 0.75 } : { style: 'none' },
    showValAxisTitle: !!axis.title,
    valAxisTitle: axis.title || undefined,
  };
}

/** Add a chart element to a pptxgenjs slide, as a native PowerPoint chart. */
export function addChart(pptx, slide, el, theme, geom, objectName) {
  const chart = el.chart || {};
  const kind = CHART_KINDS.some((k) => k.id === chart.kind) ? chart.kind : 'column';
  const pie = kind === 'pie' || kind === 'doughnut';
  let series = (chart.series || []).filter((s) => s && typeof s === 'object');
  if (kind === 'pie') series = series.slice(0, 1);
  let categories = (chart.categories || []).map((c) => String(c ?? ''));
  const n = categories.length || Math.max(0, ...series.map((s) => s.values?.length || 0));
  while (categories.length < n) categories.push(String(categories.length + 1));
  // pptxgenjs needs some data: an empty chart keeps its place with one blank point.
  if (!n) categories = [''];
  if (!series.length) series = [{ name: 'Series 1', values: [] }];
  const view = { ...chart, kind, series };
  const values = (s) => categories.map((_, i) => (isNum(s.values?.[i]) ? s.values[i] : null));
  const named = (s, i) => String(s.name ?? '') || `Series ${i + 1}`;

  // A doughnut's outer rings are added to its XML afterwards (see finishCharts); the workbook has them all.
  let data = series.map((s, i) => ({ name: named(s, i), labels: categories, values: values(s) }));
  if (kind === 'scatter') {
    // The categories are the X values (numbered 1, 2, 3… where they aren't numbers).
    const xs = categories.map((c, i) => (Number.isFinite(parseFloat(c)) ? parseFloat(c) : i + 1));
    data = [{ name: 'X', values: xs }, ...series.map((s, i) => ({ name: named(s, i), values: values(s) }))];
  }
  const colors = pie ? categories.map((_, i) => hex(pointColor(view, i, theme))) : series.map((_, i) => hex(seriesColor(view, i, theme)));

  const ink = resolveColor(el.style?.color || 'theme:text', theme, '#000000');
  const text = hex(ink) || '000000';
  const grid = hex(/^#[0-9a-f]{6}$/i.test(ink) ? blend(ink, backgroundBaseColor(theme.background), 0.2) : null) || 'D9D9D9';
  const size = el.style?.size || 12;
  // Font names go into the chart's XML as they are.
  const font = (el.style?.font || theme.fonts.body).replace(/[&<>"]/g, '');
  const fill = hex(resolveColor(el.fill, theme, null));
  const border = el.stroke?.width && el.stroke.color !== 'none' ? hex(resolveColor(el.stroke.color, theme, null)) : null;
  const grouping = chart.grouping === 'percent' ? 'percentStacked' : chart.grouping === 'stacked' ? 'stacked' : null;
  // pptxgenjs only stacks areas one way: 100% stacked areas (and stacked lines) are finished afterwards.
  const barGrouping = (k) => (k === 'area' ? (grouping ? 'stacked' : 'standard') : k === 'line' ? undefined : grouping || 'clustered');
  const format = chart.numberFormat || 'General';
  const percentAxis = grouping === 'percentStacked' && !pie && kind !== 'scatter';
  const parts = labelParts(chart);
  const labels = !!chart.dataLabels;
  const catTitle = chart.categoryAxis?.title;

  const options = {
    ...geom,
    barDir: kind === 'bar' ? 'bar' : 'col',
    barGrouping: barGrouping(kind),
    barGapWidthPct: isNum(chart.gapWidth) ? chart.gapWidth : 150,
    barOverlapPct: isNum(chart.overlap) ? chart.overlap : 0,
    holeSize: isNum(chart.holeSize) ? chart.holeSize : 50,
    showTitle: !!chart.title, title: chart.title || undefined, titleFontFace: font, titleFontSize: Math.round(size * 1.5), titleColor: text,
    showLegend: !!LEGEND_POS[chart.legend], legendPos: LEGEND_POS[chart.legend] || 'r', legendFontFace: font, legendFontSize: size, legendColor: text,
    showValue: labels && parts.includes('value'), showPercent: pie && labels && parts.includes('percent'), showLabel: pie && labels && parts.includes('category'), showLeaderLines: false,
    dataLabelFormatCode: chart.labelFormat || format, dataLabelFontFace: font, dataLabelFontSize: size, dataLabelColor: text,
    dataLabelPosition: labelSpot(kind === 'column' ? 'bar' : kind, chart.labelPosition),
    catAxisLabelFontFace: font, catAxisLabelFontSize: size, catAxisLabelColor: text, catAxisLineColor: grid,
    showCatAxisTitle: !pie && !!catTitle, catAxisTitle: catTitle || undefined, catAxisTitleFontFace: font, catAxisTitleFontSize: size, catAxisTitleColor: text,
    valAxisLabelFontFace: font, valAxisLabelFontSize: size, valAxisLabelColor: text, valAxisLineShow: false,
    valAxisLabelFormatCode: percentAxis && !isPercentFormat(format) ? '0%' : format,
    valLabelFormatCode: xmlText(format),
    ...valueAxisOptions(chart.valueAxis, chart.valueAxis?.gridlines !== false, grid),
    valAxisTitleFontFace: font, valAxisTitleFontSize: size, valAxisTitleColor: text,
    // Scatter charts have no vertical gridlines (as LibreSlides draws them).
    catGridLine: kind === 'scatter' ? { style: 'none' } : undefined,
    lineSize: kind === 'scatter' && !chart.lines ? 0 : 2.25,
    lineDataSymbol: showsMarkers(view) ? 'circle' : 'none',
    lineDataSymbolSize: 6,
    lineSmooth: chart.smooth === true,
    chartArea: { fill: fill ? { color: fill } : undefined, border: border ? { color: border, pt: el.stroke.width } : undefined, roundedCorners: false },
    plotArea: { fill: { color: null } },
    ...(objectName && { objectName }),
  };

  // Combination charts: a plot per kind and value axis, with a hidden second category axis for the secondary one.
  const groups = pie || kind === 'scatter' || kind === 'bar' ? [] : plotGroups(view);
  if (groups.length > 1) {
    const types = groups.map((g) => ({
      type: TYPES[g.kind],
      data: g.idx.map((k) => data[k]),
      options: {
        chartColors: g.idx.map((k) => colors[k] || '4472C4'),
        barGrouping: g.kind === kind ? barGrouping(g.kind) : g.kind === 'area' ? 'standard' : g.kind === 'line' ? undefined : 'clustered',
        dataLabelPosition: labelSpot(g.kind === 'column' ? 'bar' : g.kind, chart.labelPosition),
        secondaryValAxis: g.secondary, secondaryCatAxis: g.secondary,
      },
    }));
    if (groups.some((g) => g.secondary)) {
      const second = chart.secondaryAxis || {};
      options.valAxes = [{}, { ...valueAxisOptions(second, second.gridlines === true, grid), valAxisLabelFormatCode: second.numberFormat || 'General' }];
      options.catAxes = [{}, { catAxisHidden: true, showCatAxisTitle: false }];
    }
    slide.addChart(types, options);
    return;
  }
  slide.addChart(TYPES[kind], data, { ...options, chartColors: colors.map((c) => c || '4472C4') });
}

// ------------------------------------------------------------------ finishing

/** A spreadsheet column's letters (1 is A). */
const columnName = (n) => (n <= 26 ? String.fromCharCode(64 + n) : columnName(Math.floor((n - 1) / 26)) + String.fromCharCode(65 + ((n - 1) % 26)));

/** A doughnut's outer rings: copies of its first series, with each series' name and values. */
function addRings(xml, chart) {
  const first = /<c:doughnutChart>[\s\S]*?(<c:ser>[\s\S]*?<\/c:ser>)/.exec(xml)?.[1];
  if (!first) return xml;
  const rings = chart.series.slice(1, MAX_SERIES).map((s, i) => {
    const k = i + 1;
    const col = columnName(k + 2);
    const pts = (chart.categories || []).map((_, j) => `<c:pt idx="${j}"><c:v>${isNum(s.values?.[j]) ? s.values[j] : ''}</c:v></c:pt>`).join('');
    return first
      .replace(/<c:idx val="0"\/>(\s*)<c:order val="0"\/>/, `<c:idx val="${k}"/>$1<c:order val="${k}"/>`)
      .replace(/<c:tx>[\s\S]*?<\/c:tx>/, (tx) => tx.replace('$B$1', () => `$${col}$1`).replace(/<c:v>[\s\S]*?<\/c:v>/, () => `<c:v>${xmlText(s.name || `Series ${k + 1}`)}</c:v>`))
      .replace(/<c:val>[\s\S]*?<\/c:val>/, (val) => val
        .replace(/\$B\$(\d+):\$B\$(\d+)/, (_, a, b) => `$${col}$${a}:$${col}$${b}`)
        .replace(/(<c:ptCount val="\d+"\/>)[\s\S]*?(<\/c:numCache>)/, (_, count, end) => count + pts + end));
  });
  return xml.replace(first, () => first + rings.join(''));
}

/**
 * A chart part as pptxgenjs wrote it, with what it can't write: 100% stacked
 * areas, stacked lines, doughnut rings beyond the first, category names on
 * the labels of other charts and, for combination charts, line markers in
 * their series' colour and columns between the category ticks.
 */
export function fixChartXml(xml, chart) {
  const grouping = { stacked: 'stacked', percent: 'percentStacked' }[chart.grouping];
  if (chart.kind === 'area' && grouping === 'percentStacked') xml = xml.replace('<c:areaChart><c:grouping val="stacked"/>', '<c:areaChart><c:grouping val="percentStacked"/>');
  if (chart.kind === 'line' && grouping) xml = xml.replace('<c:lineChart>', `<c:lineChart><c:grouping val="${grouping}"/>`);
  if (chart.kind === 'doughnut' && chart.series?.length > 1) xml = addRings(xml, chart);
  const pie = chart.kind === 'pie' || chart.kind === 'doughnut';
  if (!pie && chart.dataLabels && labelParts(chart).includes('category')) xml = xml.replace(/<c:showCatName val="0"\/>/g, '<c:showCatName val="1"/>');
  const groups = COMBO_KINDS.includes(chart.kind) ? plotGroups(chart) : [];
  if (groups.length > 1) {
    xml = xml.replace(/<c:ser>[\s\S]*?<\/c:ser>/g, (ser) => {
      const color = /<c:spPr>\s*<a:solidFill>\s*(<a:srgbClr val="[0-9A-F]{6}"\/>)/.exec(ser)?.[1];
      return color ? ser.replace(/(<c:marker>[\s\S]*?<c:spPr>\s*<a:solidFill>)[\s\S]*?(<\/a:solidFill>)/, (_, a, b) => a + color + b) : ser;
    });
    if (groups.some((g) => g.kind === 'column')) xml = xml.replace(/<c:crossBetween val="midCat"\/>/g, '<c:crossBetween val="between"/>');
  }
  return xml;
}

/**
 * Finish the chart parts of an exported .pptx (see fixChartXml). Each
 * slide's charts are found from its graphic frames, which are in the order
 * of its elements.
 */
export async function finishCharts(zip, deck) {
  for (const [i, slide] of (deck.slides || []).entries()) {
    const charts = (slide.elements || []).filter((e) => e.type === 'chart');
    if (!charts.length) continue;
    const xml = await zip.file(`ppt/slides/slide${i + 1}.xml`)?.async('string');
    const rels = await zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`)?.async('string');
    if (!xml || !rels) continue;
    const targets = [...xml.matchAll(/<c:chart\b[^>]*\br:id="([^"]+)"/g)].map(([, id]) => {
      const rel = [...rels.matchAll(/<Relationship\b[^>]*>/g)].map(([r]) => r).find((r) => r.includes(`Id="${id}"`));
      const target = /Target="([^"]+)"/.exec(rel || '')?.[1];
      // Targets are absolute ("/ppt/charts/chart1.xml") or relative to the slide ("../charts/chart1.xml").
      return target && (target.startsWith('/') ? target.slice(1) : `ppt/${target.replace(/^\.\.\//, '')}`);
    });
    for (const [k, el] of charts.entries()) {
      const file = targets[k] && zip.file(targets[k]);
      if (!file) continue;
      const before = await file.async('string');
      const after = fixChartXml(before, el.chart || {});
      if (after !== before) zip.file(targets[k], after);
    }
  }
}
