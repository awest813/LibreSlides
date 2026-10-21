import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import PptxGenJS from 'pptxgenjs';
import { buildPptx } from '../../src/io/pptx-export.js';
import { importPptx } from '../../src/io/pptx-import.js';
import { sanitizeDeck } from '../../src/io/sanitize.js';
import { renderChart, niceScale, paletteColor } from '../../src/render/chart.js';
import { newDeck, newChart, slideText } from '../../src/model/deck.js';
import { themeById } from '../../src/model/themes.js';

const toBuffer = (blob) => (blob.arrayBuffer ? blob.arrayBuffer() : new Response(blob).arrayBuffer());
const theme = themeById('simple-light');

function deckWith(...elements) {
  const deck = newDeck({ slides: [] });
  deck.slides.push({ id: 's1', layout: 'blank', background: null, elements, notes: '', transition: { type: 'none', duration: 0.6 }, hidden: false });
  return deck;
}

/** A chart element with `chart` settings over the sample column chart. */
function chartEl(chart = {}, props = {}) {
  const el = newChart(60, 40, 480, 300);
  Object.assign(el.chart, chart);
  return Object.assign(el, props);
}

const exportZip = async (deck) => JSZip.loadAsync(await toBuffer(await buildPptx(deck)));

/** Export a deck, let `edit` change the files of the .pptx, and read it back. */
async function roundTrip(deck, edit) {
  const zip = await exportZip(deck);
  if (edit) await edit(zip);
  return importPptx(await zip.generateAsync({ type: 'arraybuffer' }), 'x');
}

// pptxgenjs numbers chart parts across presentations, so they're found by listing them.
const chartParts = (zip) => Object.keys(zip.files).filter((f) => /^ppt\/charts\/chart\d+\.xml$/.test(f)).sort((a, b) => parseInt(a.match(/\d+/)[0], 10) - parseInt(b.match(/\d+/)[0], 10));
const chartsOf = (deck) => deck.slides.flatMap((s) => s.elements.filter((e) => e.type === 'chart'));
const render = (el) => renderChart(el, { theme });

describe('chart export', () => {
  it('writes a native PowerPoint chart part', async () => {
    const zip = await exportZip(deckWith(chartEl({ title: 'Sales' })));
    const xml = await zip.file(chartParts(zip)[0]).async('string');
    expect(xml).toContain('<c:barChart>');
    expect(xml).toContain('<c:barDir val="col"/>');
    expect(xml).toContain('<c:grouping val="clustered"/>');
    expect(xml).toContain('<c:v>Series 2</c:v>');
    expect(xml).toContain('<a:t>Sales</a:t>');
    expect(xml).toContain('<c:legendPos val="b"/>');
    // The slide refers to it from a graphic frame, and the data travels in an embedded workbook.
    const slide = await zip.file('ppt/slides/slide1.xml').async('string');
    expect(slide).toMatch(/graphicData uri="http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/chart"/);
    expect(Object.keys(zip.files).some((f) => /^ppt\/embeddings\/.*\.xlsx$/.test(f))).toBe(true);
  });

  it('writes each kind of chart', async () => {
    const plots = { column: 'barChart', bar: 'barChart', line: 'lineChart', area: 'areaChart', pie: 'pieChart', doughnut: 'doughnutChart', scatter: 'scatterChart' };
    const kinds = Object.keys(plots);
    const zip = await exportZip(deckWith(...kinds.map((kind) => chartEl({ kind }))));
    for (const [i, kind] of kinds.entries()) {
      const xml = await zip.file(chartParts(zip)[i]).async('string');
      expect(xml).toContain(`<c:${plots[kind]}>`);
    }
  });

  it('keeps an empty chart in place', async () => {
    const zip = await exportZip(deckWith(chartEl({ categories: [], series: [] })));
    expect(chartParts(zip)).toHaveLength(1);
  });
});

describe('chart round trip', () => {
  it('keeps the kind, data, colours, title, legend and labels', async () => {
    const column = chartEl({ title: 'Revenue', legend: 'right', dataLabels: true, categories: ['North', 'South', 'East'],
      series: [{ name: 'Q1', values: [1.5, -2, 3], color: '#123456' }, { name: 'Q2', values: [4, null, 6], color: '#abcdef' }] });
    const bars = chartEl({ kind: 'bar', grouping: 'stacked', legend: 'top' });
    const percent = chartEl({ kind: 'column', grouping: 'percent', legend: 'left' });
    const line = chartEl({ kind: 'line', markers: true, legend: 'none' });
    const pie = chartEl({ kind: 'pie', categories: ['A', 'B', 'C'], series: [{ name: 'Share', values: [5, 3, 2], pointColors: ['#ff0000', '#00ff00', '#0000ff'] }] });
    const doughnut = chartEl({ kind: 'doughnut' });
    const scatter = chartEl({ kind: 'scatter', categories: ['1', '2.5', '4'], series: [{ name: 'Y', values: [2, 4, 8], color: '#336699' }] });
    const area = chartEl({ kind: 'area', grouping: 'stacked' });
    const back = chartsOf(await roundTrip(deckWith(column, bars, percent, line, pie, doughnut, scatter, area)));
    expect(back).toHaveLength(8);

    const [c, b, p, l, pi, d, sc, a] = back.map((e) => e.chart);
    expect(c).toMatchObject({ kind: 'column', grouping: 'clustered', title: 'Revenue', legend: 'right', dataLabels: true, categories: ['North', 'South', 'East'] });
    expect(c.series).toEqual([{ name: 'Q1', values: [1.5, -2, 3], color: '#123456' }, { name: 'Q2', values: [4, null, 6], color: '#abcdef' }]);
    expect(b).toMatchObject({ kind: 'bar', grouping: 'stacked', legend: 'top', title: '', dataLabels: false });
    expect(p).toMatchObject({ kind: 'column', grouping: 'percent', legend: 'left' });
    expect(l).toMatchObject({ kind: 'line', markers: true, legend: 'none' });
    expect(l.series.map((s) => s.values)).toEqual(line.chart.series.map((s) => s.values));
    expect(pi).toMatchObject({ kind: 'pie', categories: ['A', 'B', 'C'] });
    expect(pi.series).toEqual([{ name: 'Share', values: [5, 3, 2], pointColors: ['#ff0000', '#00ff00', '#0000ff'] }]);
    expect(d.kind).toBe('doughnut');
    expect(sc).toMatchObject({ kind: 'scatter', categories: ['1', '2.5', '4'] });
    expect(sc.series).toEqual([{ name: 'Y', values: [2, 4, 8], color: '#336699' }]);
    expect(a).toMatchObject({ kind: 'area', grouping: 'stacked' });

    // Position and text style come back too.
    expect(back[0]).toMatchObject({ x: 60, y: 40, w: 480, h: 300 });
    expect(back[0].style.size).toBe(12);
  });

  it('gives series without a colour the theme’s accents', async () => {
    const [el] = chartsOf(await roundTrip(deckWith(chartEl())));
    expect(el.chart.series.map((s) => s.color)).toEqual([0, 1, 2].map((i) => paletteColor(i, theme).toLowerCase()));
  });
});

describe('reading PowerPoint charts', () => {
  it('reads a chart made by pptxgenjs', async () => {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    const slide = pptx.addSlide();
    slide.addChart('line', [
      { name: 'Visitors', labels: ['Mon', 'Tue', 'Wed'], values: [120, 95, 140] },
      { name: 'Sign-ups', labels: ['Mon', 'Tue', 'Wed'], values: [12, 9, 17] },
    ], { x: 1, y: 1, w: 6, h: 4, showLegend: true, legendPos: 'b', showTitle: true, title: 'Traffic', chartColors: ['FF8800', '0088FF'], lineDataSymbol: 'none' });
    const deck = await importPptx(await pptx.write({ outputType: 'arraybuffer' }), 'x');
    const [el] = chartsOf(deck);
    expect(el).toMatchObject({ type: 'chart', x: 72, y: 72, w: 432, h: 288 });
    expect(el.chart).toMatchObject({ kind: 'line', title: 'Traffic', legend: 'bottom', markers: false, categories: ['Mon', 'Tue', 'Wed'] });
    expect(el.chart.series).toEqual([
      { name: 'Visitors', values: [120, 95, 140], color: '#ff8800' },
      { name: 'Sign-ups', values: [12, 9, 17], color: '#0088ff' },
    ]);
    expect(slideText(deck.slides[0])).toContain('Traffic');
  });

  /** A deck whose first chart part is replaced by `plot` (inside c:plotArea) and `extra` (inside c:chart). */
  const withChartXml = (plot, extra = '') => roundTrip(deckWith(chartEl()), (zip) => zip.file(chartParts(zip)[0],
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    + `<c:chart>${extra}<c:plotArea>${plot}</c:plotArea></c:chart></c:chartSpace>`));
  const ser = (name, vals, more = '') => `<c:ser><c:idx val="0"/><c:tx><c:v>${name}</c:v></c:tx>${more}<c:val><c:numLit><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numLit></c:val></c:ser>`;

  it('reads 3-D, combination and dated charts as flat ones', async () => {
    const cat = '<c:cat><c:numRef><c:numCache><c:formatCode>m/d/yyyy</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>45292</c:v></c:pt><c:pt idx="1"><c:v>45323</c:v></c:pt></c:numCache></c:numRef></c:cat>';
    const deck = await withChartXml(
      `<c:bar3DChart><c:barDir val="bar"/><c:grouping val="percentStacked"/>${ser('A', [1, 2], `<c:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></c:spPr>${cat}`)}</c:bar3DChart>`
      + `<c:lineChart>${ser('B', [3, 4])}</c:lineChart>`,
      '<c:autoTitleDeleted val="0"/><c:legend/>',
    );
    const { chart } = chartsOf(deck)[0];
    // No title of its own, and more than one series: PowerPoint shows none.
    expect(chart).toMatchObject({ kind: 'bar', grouping: 'percent', legend: 'right', title: '' });
    expect(chart.series.map((s) => [s.name, s.values])).toEqual([['A', [1, 2]], ['B', [3, 4]]]);
    expect(chart.series[0].color).toMatch(/^#[0-9a-f]{6}$/);
    expect(chart.categories).toHaveLength(2);
    expect(chart.categories[0]).toMatch(/2024/);
  });

  it('reads pie slice colours and labels', async () => {
    const dPt = '<c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></c:spPr></c:dPt>';
    const deck = await withChartXml(`<c:pie3DChart><c:varyColors val="1"/>${ser('Fruit', [3, 2, 1], dPt)}<c:dLbls><c:showVal val="1"/></c:dLbls></c:pie3DChart>`, '<c:title><c:tx><c:rich><a:p><a:r><a:t>Fruit</a:t></a:r></a:p></c:rich></c:tx></c:title>');
    const { chart } = chartsOf(deck)[0];
    expect(chart).toMatchObject({ kind: 'pie', title: 'Fruit', dataLabels: true, legend: 'none', categories: ['1', '2', '3'] });
    expect(chart.series[0].pointColors[1]).toBe('#00aa00');
    expect(chart.series[0].pointColors).toHaveLength(3);
  });

  it('leaves out charts it can’t draw', async () => {
    const deck = await withChartXml(`<c:radarChart>${ser('A', [1, 2])}</c:radarChart>`);
    expect(chartsOf(deck)).toHaveLength(0);
  });
});

describe('drawing charts', () => {
  it('draws a bar per value and a legend entry per series', () => {
    const svg = render(chartEl({ title: 'T', dataLabels: true }));
    expect(svg.tagName).toBe('svg');
    expect(svg.querySelectorAll('.chart-bar')).toHaveLength(12);
    expect(svg.querySelectorAll('.chart-legend rect')).toHaveLength(3);
    expect(svg.querySelectorAll('.chart-label')).toHaveLength(12);
    expect(svg.querySelector('.chart-title').textContent).toBe('T');
  });

  it('skips blank values and copes with negative ones', () => {
    const el = chartEl({ kind: 'bar', categories: ['a', 'b', 'c'], series: [{ name: 'S', values: [3, null, -2] }] });
    const svg = render(el);
    expect(svg.querySelectorAll('.chart-bar')).toHaveLength(2);
    const labels = [...svg.querySelectorAll('.chart-labels text')].map((t) => t.textContent);
    expect(labels).toContain('-2');
  });

  it('draws stacked and 100% stacked charts', () => {
    const stacked = render(chartEl({ grouping: 'stacked' }));
    expect(stacked.querySelectorAll('.chart-bar')).toHaveLength(12);
    const percent = render(chartEl({ grouping: 'percent' }));
    const ticks = [...percent.querySelectorAll('.chart-labels text')].map((t) => t.textContent).filter((t) => t.endsWith('%'));
    expect(ticks).toContain('100%');
    const area = render(chartEl({ kind: 'area', grouping: 'stacked' }));
    expect(area.querySelectorAll('.chart-area')).toHaveLength(3);
  });

  it('draws pie and doughnut slices', () => {
    const pie = chartEl({ kind: 'pie', categories: ['a', 'b', 'c', 'd'], series: [{ name: 'S', values: [1, 2, 0, 4] }] });
    expect(render(pie).querySelectorAll('.chart-slice')).toHaveLength(3);
    expect(render(pie).querySelectorAll('.chart-legend rect')).toHaveLength(4);
    const whole = chartEl({ kind: 'doughnut', categories: ['a'], series: [{ name: 'S', values: [5] }] });
    expect(render(whole).querySelector('.chart-slice').getAttribute('d')).toMatch(/Z.*Z/);
  });

  it('draws lines with gaps, and scatter points', () => {
    const line = render(chartEl({ kind: 'line', series: [{ name: 'S', values: [1, 2, null, 4, 5] }], categories: ['a', 'b', 'c', 'd', 'e'] }));
    expect(line.querySelectorAll('.chart-line')).toHaveLength(2);
    const scatter = render(chartEl({ kind: 'scatter', categories: ['1', '2', '3'], series: [{ name: 'S', values: [1, null, 3] }] }));
    expect(scatter.querySelectorAll('.chart-marker')).toHaveLength(2);
  });

  it('doesn’t break on empty, tiny or damaged charts', () => {
    for (const chart of [{ categories: [], series: [] }, { series: [{ name: 'S', values: [] }] }, { kind: 'pie', series: [] }, { kind: 'line', series: [{ name: 'S', values: [null, null, null, null] }] }]) {
      expect(() => render(chartEl(chart))).not.toThrow();
    }
    expect(() => render(chartEl({}, { w: 3, h: 2 }))).not.toThrow();
    expect(() => render({ id: 'x', type: 'chart', x: 0, y: 0, w: 100, h: 100 })).not.toThrow();
  });

  it('rounds the value scale', () => {
    expect(niceScale(0, 28.6, 6)).toMatchObject({ min: 0, max: 30, step: 5 });
    expect(niceScale(-4.5, 28.6, 8)).toMatchObject({ min: -5, max: 30, step: 5 });
    expect(niceScale(0, 0.37, 5)).toMatchObject({ min: 0, max: 0.4, step: 0.1, digits: 1 });
    expect(niceScale(0, 0, 5)).toMatchObject({ min: 0, max: 1 });
  });
});

describe('untrusted chart data', () => {
  it('keeps only well-formed chart data', () => {
    const deck = sanitizeDeck({ slides: [{ elements: [
      { type: 'chart', x: 0, y: 0, w: 10, h: 10, chart: {
        kind: 'radar', grouping: 'sideways', title: { toString: () => 'x' }, legend: 'middle', dataLabels: 'yes', markers: true,
        categories: ['a', 2, null, { x: 1 }],
        series: [{ name: '<b>S</b>', values: [1, 'two', NaN, Infinity, -3], color: 'url(https://t/x)', pointColors: ['#fff', 'red;x'] }, 'junk', { values: 7 }],
      } },
      { type: 'chart', chart: 'nonsense' },
    ] }] });
    const [a, b] = deck.slides[0].elements;
    expect(a.chart).toEqual({
      kind: 'column', title: '', legend: 'none', dataLabels: false, markers: true,
      categories: ['a', '2', '', ''],
      series: [{ name: '<b>S</b>', values: [1, null, null, null, -3], pointColors: ['#fff', null] }, { name: 'Series 2', values: [] }],
    });
    expect(b.chart).toEqual({ kind: 'column', title: '', categories: [], series: [], legend: 'none', dataLabels: false });
  });

  it('caps the size of chart data', () => {
    const big = { kind: 'line', categories: Array.from({ length: 5000 }, (_, i) => `c${i}`), series: Array.from({ length: 500 }, () => ({ name: 'x'.repeat(1000), values: Array(5000).fill(1) })) };
    const deck = sanitizeDeck({ slides: [{ elements: [{ type: 'chart', chart: big }] }] });
    const { chart } = deck.slides[0].elements[0];
    expect(chart.categories).toHaveLength(1000);
    expect(chart.series).toHaveLength(200);
    expect(chart.series[0].values).toHaveLength(1000);
    expect(chart.series[0].name).toHaveLength(255);
  });
});
