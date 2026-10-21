import { uid, CHART_KINDS } from '../model/deck.js';
import { resolveColor, backgroundBaseColor } from '../model/themes.js';
import { seriesColor, pointColor } from '../render/chart.js';
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

/** Category labels: dates (numbers in a date format) as dates, other numbers as they are. */
function categoryLabels(holder, date1904) {
  const { values, format } = pointsOf(holder);
  const isDate = /[dmy]/i.test(format.replace(/"[^"]*"|\[[^\]]*\]/g, '')) && !/general/i.test(format);
  return values.map((v) => {
    if (v == null) return '';
    if (isDate && Number.isFinite(+v)) {
      const day = Date.UTC(date1904 ? 1904 : 1899, date1904 ? 0 : 11, date1904 ? 1 : 30) + +v * 86400000;
      return new Date(day).toLocaleDateString(undefined, { timeZone: 'UTC', year: 'numeric', month: 'short', day: 'numeric' });
    }
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

/** Whether a plot (or one of its series) shows values on its points. */
function showsValues(plot, ser) {
  const dLbls = kid(ser, 'dLbls') || kid(plot, 'dLbls');
  if (!dLbls || isOn(kid(dLbls, 'delete'))) return false;
  return isOn(kid(dLbls, 'showVal')) || kids(dLbls, 'dLbl').some((d) => isOn(kid(d, 'showVal')));
}

/** The chart's text style from its axes, legend or the whole chart (in that order). */
function textStyleOf(space, chartEl, ctx) {
  const plotArea = kid(chartEl, 'plotArea');
  const sources = [...kids(plotArea, 'catAx'), ...kids(plotArea, 'valAx'), ...kids(plotArea, 'dateAx'), kid(chartEl, 'legend'), space];
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

/** The chart data of a c:chartSpace, or null for a chart LibreSlides can't draw. */
export function parseChart(space, ctx) {
  const chartEl = kid(space, 'chart');
  const plotArea = kid(chartEl, 'plotArea');
  // Combination charts: the first kind, with every series.
  const plots = [...(plotArea?.children || [])].filter((c) => PLOTS[c.localName] || OTHER_PLOTS.has(c.localName));
  const first = plots[0];
  if (!first || !PLOTS[first.localName]) return null;
  let kind = PLOTS[first.localName];
  if (kind === 'column' && attr(kid(first, 'barDir'), 'val') === 'bar') kind = 'bar';
  const date1904 = isOn(kid(space, 'date1904'));
  const pie = kind === 'pie' || kind === 'doughnut';
  const sers = plots.filter((p) => PLOTS[p.localName]).flatMap((p) => kids(p, 'ser').map((ser) => ({ ser, plot: p })));
  const chosen = (pie ? sers.slice(0, 1) : sers).slice(0, MAX_SERIES);

  let categories = [];
  const series = chosen.map(({ ser }, i) => ({
    name: textOf(kid(ser, 'tx')) || `Series ${i + 1}`,
    values: pointsOf(kid(ser, 'val') || kid(ser, 'yVal')).values.map(toNumber),
    color: seriesColorOf(ser, kind, i, ctx.scheme),
  }));
  if (kind === 'scatter') {
    const merged = mergeScatter(chosen.map(({ ser }, i) => ({ x: pointsOf(kid(ser, 'xVal')).values.map(toNumber), y: series[i].values })));
    categories = merged.categories;
    series.forEach((s, i) => { s.values = merged.values[i]; });
  } else {
    const withCats = chosen.find(({ ser }) => kid(ser, 'cat'));
    categories = withCats ? categoryLabels(kid(withCats.ser, 'cat'), date1904) : [];
  }
  // Every series as long as the categories (PowerPoint numbers them 1, 2, 3… if there are none).
  const n = Math.min(MAX_POINTS, Math.max(categories.length, ...series.map((s) => s.values.length)));
  while (categories.length < n) categories.push(String(categories.length + 1));
  for (const s of series) {
    s.values = Array.from({ length: n }, (_, i) => (isNum(s.values[i]) ? s.values[i] : null));
    if (!s.color) delete s.color;
  }

  const chart = { kind, categories, series, legend: 'none', dataLabels: false, title: '' };
  if (kind === 'column' || kind === 'bar' || kind === 'area') chart.grouping = GROUPING[attr(kid(first, 'grouping'), 'val')] || 'clustered';
  if (pie && series[0]) {
    // Slices: their own colours, else the accents in turn (or all the series' colour if they don't vary).
    const ser = chosen[0].ser;
    const vary = isOn(kid(first, 'varyColors')) || !kid(first, 'varyColors');
    const own = {};
    for (const dPt of kids(ser, 'dPt')) own[+attr(kid(dPt, 'idx'), 'val')] = fillColor(kid(dPt, 'spPr'), ctx.scheme);
    const plain = fillColor(kid(ser, 'spPr'), ctx.scheme);
    series[0].pointColors = categories.map((_, i) => own[i] || (vary ? accent(i, ctx.scheme) : plain) || accent(i, ctx.scheme) || '#4472c4');
    delete series[0].color;
  }
  if (kind === 'line') {
    const symbol = attr(path(chosen[0]?.ser, 'marker', 'symbol'), 'val');
    chart.markers = !!chosen.length && symbol !== 'none';
  }
  chart.dataLabels = !!chosen.length && showsValues(chosen[0].plot, chosen[0].ser);

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

/** Add a chart element to a pptxgenjs slide, as a native PowerPoint chart. */
export function addChart(pptx, slide, el, theme, geom) {
  const chart = el.chart || {};
  const kind = CHART_KINDS.some((k) => k.id === chart.kind) ? chart.kind : 'column';
  const pie = kind === 'pie' || kind === 'doughnut';
  let series = (chart.series || []).filter((s) => s && typeof s === 'object');
  if (pie) series = series.slice(0, 1);
  let categories = (chart.categories || []).map((c) => String(c ?? ''));
  const n = categories.length || Math.max(0, ...series.map((s) => s.values?.length || 0));
  while (categories.length < n) categories.push(String(categories.length + 1));
  // pptxgenjs needs some data: an empty chart keeps its place with one blank point.
  if (!n) categories = [''];
  if (!series.length) series = [{ name: 'Series 1', values: [] }];
  const values = (s) => categories.map((_, i) => (isNum(s.values?.[i]) ? s.values[i] : null));
  const named = (s, i) => String(s.name ?? '') || `Series ${i + 1}`;

  let data = series.map((s, i) => ({ name: named(s, i), labels: categories, values: values(s) }));
  if (kind === 'scatter') {
    // The categories are the X values (numbered 1, 2, 3… where they aren't numbers).
    const xs = categories.map((c, i) => (Number.isFinite(parseFloat(c)) ? parseFloat(c) : i + 1));
    data = [{ name: 'X', values: xs }, ...series.map((s, i) => ({ name: named(s, i), values: values(s) }))];
  }
  const colors = pie ? categories.map((_, i) => hex(pointColor(chart, i, theme))) : series.map((_, i) => hex(seriesColor(chart, i, theme)));

  const ink = resolveColor(el.style?.color || 'theme:text', theme, '#000000');
  const text = hex(ink) || '000000';
  const grid = hex(/^#[0-9a-f]{6}$/i.test(ink) ? blend(ink, backgroundBaseColor(theme.background), 0.2) : null) || 'D9D9D9';
  const size = el.style?.size || 12;
  // Font names go into the chart's XML as they are.
  const font = (el.style?.font || theme.fonts.body).replace(/[&<>"]/g, '');
  const fill = hex(resolveColor(el.fill, theme, null));
  const border = el.stroke?.width && el.stroke.color !== 'none' ? hex(resolveColor(el.stroke.color, theme, null)) : null;
  const grouping = chart.grouping === 'percent' ? 'percentStacked' : chart.grouping === 'stacked' ? 'stacked' : null;

  slide.addChart(TYPES[kind], data, {
    ...geom,
    chartColors: colors.map((c) => c || '4472C4'),
    barDir: kind === 'bar' ? 'bar' : 'col',
    // pptxgenjs only stacks areas one way; 100% stacked areas are written as stacked ones.
    barGrouping: kind === 'area' ? (grouping ? 'stacked' : 'standard') : grouping || 'clustered',
    barGapWidthPct: 150,
    holeSize: 50,
    showTitle: !!chart.title, title: chart.title || undefined, titleFontFace: font, titleFontSize: Math.round(size * 1.5), titleColor: text,
    showLegend: !!LEGEND_POS[chart.legend], legendPos: LEGEND_POS[chart.legend] || 'r', legendFontFace: font, legendFontSize: size, legendColor: text,
    showValue: !!chart.dataLabels, showPercent: false, showLabel: false, showLeaderLines: false,
    dataLabelFormatCode: 'General', dataLabelFontFace: font, dataLabelFontSize: size, dataLabelColor: text,
    catAxisLabelFontFace: font, catAxisLabelFontSize: size, catAxisLabelColor: text, catAxisLineColor: grid,
    valAxisLabelFontFace: font, valAxisLabelFontSize: size, valAxisLabelColor: text, valAxisLineShow: false,
    valAxisLabelFormatCode: grouping === 'percentStacked' ? '0%' : 'General',
    valGridLine: { color: grid, size: 0.75 },
    lineSize: kind === 'scatter' ? 0 : 2.25,
    lineDataSymbol: kind === 'scatter' || chart.markers ? 'circle' : 'none',
    lineDataSymbolSize: 6,
    chartArea: { fill: fill ? { color: fill } : undefined, border: border ? { color: border, pt: el.stroke.width } : undefined, roundedCorners: false },
    plotArea: { fill: { color: null } },
  });
}
