import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import PptxGenJS from 'pptxgenjs';
import { buildPptx } from '../../src/io/pptx-export.js';
import { importPptx } from '../../src/io/pptx-import.js';
import { sanitizeDeck } from '../../src/io/sanitize.js';
import { renderChart, niceScale, axisScale, paletteColor } from '../../src/render/chart.js';
import { formatValue, isDateFormat } from '../../src/render/number-format.js';
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

/** A deck whose first chart part is replaced by `plot` (inside c:plotArea) and `extra` (inside c:chart). */
const withChartXml = (plot, extra = '') => roundTrip(deckWith(chartEl()), (zip) => zip.file(chartParts(zip)[0],
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
  + `<c:chart>${extra}<c:plotArea>${plot}</c:plotArea></c:chart></c:chartSpace>`));
const ser = (name, vals, more = '') => `<c:ser><c:idx val="0"/><c:tx><c:v>${name}</c:v></c:tx>${more}<c:val><c:numLit><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numLit></c:val></c:ser>`;

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

  it('reads radar and bubble charts, and stock and surface charts as lines', async () => {
    const radar = chartsOf(await withChartXml(`<c:radarChart><c:radarStyle val="filled"/>${ser('A', [1, 2, 3])}</c:radarChart>`))[0];
    expect(radar.chart).toMatchObject({ kind: 'radar', filled: true, series: [{ name: 'A', values: [1, 2, 3] }] });
    const xy = (tag, vals) => `<c:${tag}><c:numLit><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numLit></c:${tag}>`;
    const bubble = chartsOf(await withChartXml(`<c:bubbleChart><c:ser><c:idx val="0"/><c:tx><c:v>B</c:v></c:tx>${xy('xVal', [1, 2])}${xy('yVal', [5, 6])}${xy('bubbleSize', [10, 40])}</c:ser></c:bubbleChart>`))[0];
    expect(bubble.chart).toMatchObject({ kind: 'bubble', categories: ['1', '2'], series: [{ values: [5, 6], sizes: [10, 40] }] });
    for (const plot of ['stockChart', 'surfaceChart']) {
      const chart = chartsOf(await withChartXml(`<c:${plot}>${ser('A', [1, 2])}${ser('B', [3, 4])}</c:${plot}>`))[0];
      expect(chart.chart).toMatchObject({ kind: 'line', series: [{ values: [1, 2] }, { values: [3, 4] }] });
    }
  });

  it('colours series as the chart\'s style says: accents in turn, or shades of one', async () => {
    const noIdx = (x) => x.replace(/<c:idx val="0"\/>/g, '');
    const two = chartsOf(await withChartXml(noIdx(`<c:barChart>${ser('A', [1])}${ser('B', [2])}</c:barChart>`)))[0];
    const [a1, a2] = two.chart.series.map((s) => s.color);
    expect(a1).not.toBe(a2);
    // Style 12: shades of accent 2, darkest first (Office 2010 writes it as c14:style 112).
    const style = '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="112"/></mc:Choice><mc:Fallback><c:style val="12"/></mc:Fallback></mc:AlternateContent>';
    const deck = await roundTrip(deckWith(chartEl()), async (zip) => {
      const part = chartParts(zip)[0];
      const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        + `${style}<c:chart><c:plotArea><c:barChart>${ser('A', [1])}${ser('B', [2])}${ser('C', [3])}</c:barChart></c:plotArea></c:chart></c:chartSpace>`;
      zip.file(part, noIdx(xml));
    });
    const colors = chartsOf(deck)[0].chart.series.map((s) => s.color);
    const lum = (c) => [1, 3, 5].reduce((t, i) => t + parseInt(c.slice(i, i + 2), 16), 0);
    expect(new Set(colors).size).toBe(3);
    expect(lum(colors[0])).toBeLessThan(lum(colors[1]));
    expect(lum(colors[1])).toBeLessThan(lum(colors[2]));
    expect(colors[1]).toBe(a2);
  });
});

describe('radar and bubble charts', () => {
  const radar = () => chartEl({ kind: 'radar', categories: ['a', 'b', 'c', 'd', 'e'], series: [{ name: 'S', values: [1, 4, 2, 5, 3] }, { name: 'T', values: [2, 2, 2, 2, 2] }], legend: 'right' });
  const bubble = () => chartEl({ kind: 'bubble', categories: ['1', '2', '3'], series: [{ name: 'B', values: [3, 1, 2], sizes: [10, 40, 0] }] });

  it('draws a radar as a web with a closed line per series, or filled', () => {
    const svg = render(radar());
    expect(svg.querySelectorAll('.chart-line')).toHaveLength(2);
    expect(svg.querySelector('.chart-line').getAttribute('d')).toMatch(/Z$/);
    expect(svg.querySelectorAll('.chart-grid line')).toHaveLength(5);
    expect(svg.querySelectorAll('.chart-legend line')).toHaveLength(2);
    const filled = render(chartEl({ ...radar().chart, filled: true }));
    expect(filled.querySelectorAll('.chart-area')).toHaveLength(2);
  });

  it('draws bubbles with areas as their sizes', () => {
    const svg = render(bubble());
    const r = [...svg.querySelectorAll('.chart-bubble')].map((c) => +c.getAttribute('r'));
    expect(r).toHaveLength(2); // a size of 0 has no bubble
    expect(r[1] / r[0]).toBeCloseTo(2, 1);
  });

  it('writes them as PowerPoint radar and bubble charts, and reads them back', async () => {
    const zip = await exportZip(deckWith(radar(), bubble()));
    const xml = await Promise.all(chartParts(zip).map((p) => zip.file(p).async('string')));
    expect(xml[0]).toContain('<c:radarChart><c:radarStyle val="standard"/>');
    expect(xml[1]).toContain('<c:bubbleChart>');
    expect(xml[1]).toMatch(/<c:bubbleSize>[\s\S]*<c:v>40<\/c:v>/);
    const back = chartsOf(await importPptx(await zip.generateAsync({ type: 'arraybuffer' }), 'x'));
    expect(back[0].chart).toMatchObject({ kind: 'radar', series: [{ values: [1, 4, 2, 5, 3] }, { values: [2, 2, 2, 2, 2] }] });
    expect(back[1].chart).toMatchObject({ kind: 'bubble', categories: ['1', '2', '3'], series: [{ values: [3, 1, 2], sizes: [10, 40, null] }] });
  });
});

describe('chart area pictures', () => {
  const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

  it('draws, writes and reads back a picture filling the chart area', async () => {
    const deck = deckWith(chartEl({}, { fillImage: { media: 'bg' } }));
    deck.media = { bg: PNG };
    const svg = renderChart(deck.slides[0].elements[0], { theme, deck }, PNG);
    expect(svg.querySelector('image').getAttribute('href')).toBe(PNG);
    const zip = await exportZip(deck);
    const part = chartParts(zip)[0];
    const xml = await zip.file(part).async('string');
    expect(xml).toMatch(/<\/c:chart><c:spPr><a:blipFill><a:blip r:embed="rIdLsPicture"\/>/);
    const rels = await zip.file(part.replace(/([^/]+)$/, '_rels/$1.rels')).async('string');
    const target = /Id="rIdLsPicture"[^>]*Target="\.\.\/media\/([^"]+)"/.exec(rels)[1];
    expect(zip.file(`ppt/media/${target}`)).toBeTruthy();
    const back = await importPptx(await zip.generateAsync({ type: 'arraybuffer' }), 'x');
    const el = chartsOf(back)[0];
    expect(back.media[el.fillImage.media]).toBe(PNG);
  });
});

describe('drawing charts', () => {
  it('wraps or turns category labels that don\'t fit, rather than leaving them out', () => {
    const cats = ['North region', 'South region', 'East region', 'West region'];
    const series = [{ name: 'S', values: [1, 2, 3, 4] }];
    const texts = (svg) => [...svg.querySelectorAll('.chart-labels text')].map((t) => t.textContent);
    const wrapped = render(chartEl({ categories: cats, series }, { w: 260, h: 200 }));
    expect(texts(wrapped)).toEqual(expect.arrayContaining(['Northregion', 'Westregion']));
    const turned = render(chartEl({ categories: ['Categorisation 1', 'Categorisation 2', 'Categorisation 3', 'Categorisation 4', 'Categorisation 5', 'Categorisation 6'], series }, { w: 200, h: 220 }));
    const labels = [...turned.querySelectorAll('.chart-labels text[transform*="rotate(-45)"]')];
    expect(labels).toHaveLength(6);
  });

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
        kind: 'funnel', grouping: 'sideways', title: { toString: () => 'x' }, legend: 'middle', dataLabels: 'yes', markers: true,
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

// ------------------------------------------------------------------ number formats, labels, axes, rings, combinations

/** A number with exactly `d` decimals, as the reader's locale writes it. */
const num = (v, d = 0, grouping = true) => new Intl.NumberFormat(undefined, { minimumFractionDigits: d, maximumFractionDigits: d, useGrouping: grouping }).format(v);
const texts = (svg, sel) => [...svg.querySelectorAll(sel)].map((t) => t.textContent);
const partXml = async (zip, i = 0) => zip.file(chartParts(zip)[i]).async('string');

describe('number formats', () => {
  it('formats numbers as Excel codes do', () => {
    expect(formatValue(1234.567, 'General')).toBe(num(1234.567, 2).replace(/0+$/, ''));
    expect(formatValue(1234.567, '0')).toBe(num(1235, 0, false));
    expect(formatValue(1234.567, '0.0')).toBe(num(1234.6, 1, false));
    expect(formatValue(1234.5, '0.00')).toBe(num(1234.5, 2, false));
    expect(formatValue(1234567.8, '#,##0')).toBe(num(1234568));
    expect(formatValue(1234.5, '#,##0.00')).toBe(num(1234.5, 2));
    expect(formatValue(0.123, '0%')).toBe('12%');
    expect(formatValue(0.1234, '0.0%')).toBe(`${num(12.3, 1)}%`);
    expect(formatValue(1234, '$#,##0')).toBe(`$${num(1234)}`);
    expect(formatValue(-1234, '$#,##0')).toBe(`-$${num(1234)}`);
    expect(formatValue(1234, '#,##0 €')).toBe(`${num(1234)} €`);
    expect(formatValue(1234, '#,##0\\ "€"')).toBe(`${num(1234)} €`);
    expect(formatValue(1234, '[$€-407]#,##0')).toBe(`€${num(1234)}`);
    expect(formatValue(-5, '#,##0;(#,##0)')).toBe('(5)');
    expect(formatValue(0, '#,##0;-#,##0;"zero"')).toBe('zero');
    expect(formatValue(1234567, '#,##0,"K"')).toBe(`${num(1235)}K`);
  });

  it('formats dates', () => {
    expect(isDateFormat('m/d/yyyy')).toBe(true);
    expect(isDateFormat('#,##0.00')).toBe(false);
    expect(isDateFormat('"days" 0')).toBe(false);
    expect(formatValue(45292, 'm/d/yyyy')).toBe('1/1/2024');
    expect(formatValue(45292, 'yyyy-mm-dd')).toBe('2024-01-01');
    expect(formatValue(45292.75, 'h:mm AM/PM')).toBe('6:00 PM');
    expect(formatValue(45292, 'mmm yy')).toMatch(/24$/);
  });
});

describe('chart number formats, labels and axes', () => {
  const cats = ['North', 'South', 'East'];
  const sales = (more = {}) => chartEl({ categories: cats, series: [{ name: 'S', values: [1250, 2300, 1800] }], legend: 'none', ...more });

  it('draws ticks and labels in the chart’s number formats', () => {
    const svg = render(sales({ numberFormat: '$#,##0', labelFormat: '#,##0.0', dataLabels: true }));
    expect(texts(svg, '.chart-labels > text:not(.chart-label)')).toContain(`$${num(2500)}`);
    expect(texts(svg, '.chart-label')).toEqual([num(1250, 1), num(2300, 1), num(1800, 1)]);
  });

  it('shows categories, values and percentages on pie labels', () => {
    const pie = chartEl({ kind: 'pie', categories: ['A', 'B'], series: [{ name: 'S', values: [3, 1] }], dataLabels: true, labelContent: ['category', 'percent'] });
    const labels = [...render(pie).querySelectorAll('.chart-label')].map((t) => [...t.querySelectorAll('tspan')].map((s) => s.textContent));
    expect(labels).toEqual([['A', '75%'], ['B', '25%']]);
    // Old decks (true / false) show values.
    expect(texts(render({ ...pie, chart: { ...pie.chart, labelContent: undefined } }), '.chart-label')).toEqual(['3', '1']);
  });

  it('follows the value axis’s bounds, unit, gridlines and titles', () => {
    expect(axisScale(0, 2300, 6, { min: 0, max: 4000, unit: 500 })).toMatchObject({ min: 0, max: 4000, step: 500 });
    expect(axisScale(0, 2300, 6, { max: 3000 })).toMatchObject({ min: 0, max: 3000 });
    expect(axisScale(0, 0.5, 6, { min: 0, max: 0.6 }, 100)).toMatchObject({ min: 0, max: 60 });
    // Units too fine for the range fall back to the automatic scale.
    expect(axisScale(0, 2300, 6, { unit: 0.0001 })).toMatchObject({ step: 500 });
    const svg = render(sales({ valueAxis: { min: 0, max: 4000, unit: 500, gridlines: false, title: 'Revenue' }, categoryAxis: { title: 'Region' } }));
    const ticks = texts(svg, '.chart-labels > text');
    expect(ticks).toContain(num(4000, 0, true));
    expect(ticks.filter((t) => /^\d/.test(t))).toHaveLength(9);
    expect(svg.querySelectorAll('.chart-grid line:not(.chart-axis)')).toHaveLength(0);
    expect(texts(svg, '.chart-axis-title')).toEqual(['Region', 'Revenue']);
  });

  it('spaces bars by their gap width and overlap', () => {
    const widths = (more) => [...render(chartEl({ categories: ['a'], series: [{ name: 'A', values: [1] }, { name: 'B', values: [2] }], legend: 'none', ...more })).querySelectorAll('.chart-bar')].map((b) => [+b.getAttribute('x'), +b.getAttribute('width')]);
    const [[x1, w1], [x2]] = widths({});
    expect(x2 - x1).toBeCloseTo(w1, 0);
    const [[a, wide], [b]] = widths({ gapWidth: 0, overlap: 50 });
    expect(wide).toBeGreaterThan(w1);
    expect(b - a).toBeCloseTo(wide / 2, 0);
  });

  it('places data labels where the chart says', () => {
    const at = (labelPosition) => {
      const svg = render(sales({ dataLabels: true, labelPosition }));
      const bar = svg.querySelector('.chart-bar');
      const y = +svg.querySelector('.chart-label').getAttribute('y');
      return y - +bar.getAttribute('y');
    };
    expect(at('outEnd')).toBeLessThan(0);
    expect(at('inEnd')).toBeGreaterThan(0);
    expect(at('inBase')).toBeGreaterThan(at('ctr'));
    expect(at('ctr')).toBeGreaterThan(at('inEnd'));
  });

  it('draws a ring per doughnut series, the first innermost', () => {
    const svg = render(chartEl({ kind: 'doughnut', categories: ['a', 'b'], series: [{ name: 'X', values: [1, 1] }, { name: 'Y', values: [2, 1] }, { name: 'Z', values: [1, 3] }] }));
    expect(svg.querySelectorAll('.chart-slice')).toHaveLength(6);
    // Pies still show their first series only.
    expect(render(chartEl({ kind: 'pie', categories: ['a', 'b'], series: [{ name: 'X', values: [1, 1] }, { name: 'Y', values: [2, 1] }] })).querySelectorAll('.chart-slice')).toHaveLength(2);
  });

  it('joins scatter points with lines when asked', () => {
    const scatter = (more) => render(chartEl({ kind: 'scatter', categories: ['1', '2', '3'], series: [{ name: 'S', values: [1, 3, 2] }], ...more }));
    expect(scatter({}).querySelectorAll('.chart-line')).toHaveLength(0);
    expect(scatter({ lines: true }).querySelectorAll('.chart-line')).toHaveLength(1);
    expect(scatter({ lines: true, markers: false }).querySelectorAll('.chart-marker')).toHaveLength(0);
    expect(scatter({ lines: true, smooth: true }).querySelector('.chart-line').getAttribute('d')).toContain('C');
  });

  it('stacks lines and draws 100% stacked areas to the top', () => {
    const data = { categories: ['a', 'b'], series: [{ name: 'A', values: [1, 2] }, { name: 'B', values: [3, 2] }], legend: 'none' };
    const top = (svg) => Math.min(...[...svg.querySelectorAll('.chart-line')].flatMap((l) => l.getAttribute('d').match(/[\d.]+,[\d.]+/g).map((p) => +p.split(',')[1])));
    const plain = render(chartEl({ kind: 'line', ...data }));
    const stacked = render(chartEl({ kind: 'line', grouping: 'stacked', ...data }));
    expect(texts(stacked, '.chart-labels > text')).toContain('4');
    expect(top(stacked)).toBeLessThanOrEqual(top(plain));
    const percent = render(chartEl({ kind: 'area', grouping: 'percent', ...data }));
    expect(texts(percent, '.chart-labels > text')).toContain('100%');
  });

  it('draws combination charts with a secondary axis', () => {
    const svg = render(chartEl({
      categories: ['a', 'b', 'c'], legend: 'bottom',
      series: [{ name: 'Units', values: [120, 150, 90] }, { name: 'Margin', values: [0.1, 0.2, 0.15], kind: 'line', secondary: true }],
      secondaryAxis: { numberFormat: '0%', title: 'Margin' },
    }));
    expect(svg.querySelectorAll('.chart-bar')).toHaveLength(3);
    expect(svg.querySelectorAll('.chart-line')).toHaveLength(1);
    expect(svg.querySelectorAll('.chart-legend line')).toHaveLength(1);
    const right = [...svg.querySelectorAll('.chart-labels > text')].filter((t) => t.getAttribute('text-anchor') === 'start').map((t) => t.textContent);
    expect(right).toContain('20%');
    expect(texts(svg, '.chart-axis-title')).toEqual(['Margin']);
  });
});

describe('exporting chart details', () => {
  it('writes number formats, axis settings, gap width and label positions', async () => {
    const zip = await exportZip(deckWith(chartEl({
      numberFormat: '$#,##0', labelFormat: '0.0', dataLabels: true, labelPosition: 'inEnd', gapWidth: 60, overlap: -20,
      valueAxis: { min: 0, max: 8, unit: 2, gridlines: false, title: 'Value' }, categoryAxis: { title: 'Kind' },
    })));
    const xml = await partXml(zip);
    expect(xml).toContain('<c:numFmt formatCode="$#,##0" sourceLinked="0"/>');
    expect(xml).toContain('<c:numFmt formatCode="0.0" sourceLinked="0"/>');
    expect(xml).toMatch(/<c:max val="8"\/><c:min val="0"\/>/);
    expect(xml).toContain('<c:majorUnit val="2"/>');
    expect(xml).not.toContain('<c:majorGridlines>');
    expect(xml).toContain('<a:t>Value</a:t>');
    expect(xml).toContain('<a:t>Kind</a:t>');
    expect(xml).toContain('<c:gapWidth val="60"/>');
    expect(xml).toContain('<c:overlap val="-20"/>');
    expect(xml).toContain('<c:dLblPos val="inEnd"/>');
  });

  it('writes pie percentages, doughnut rings, 100% areas, stacked lines and scatter lines', async () => {
    const zip = await exportZip(deckWith(
      chartEl({ kind: 'pie', dataLabels: true, labelContent: ['category', 'percent'], labelPosition: 'outEnd' }),
      chartEl({ kind: 'doughnut', holeSize: 60 }),
      chartEl({ kind: 'area', grouping: 'percent' }),
      chartEl({ kind: 'line', grouping: 'stacked' }),
      chartEl({ kind: 'scatter', categories: ['1', '2', '3', '4'], lines: true, smooth: true, markers: false }),
      chartEl({ kind: 'column', dataLabels: true, labelContent: ['category', 'value'] }),
    ));
    const [pie, doughnut, area, line, scatter, column] = await Promise.all([0, 1, 2, 3, 4, 5].map((i) => partXml(zip, i)));
    expect(pie).toMatch(/<c:dLbl>.*?<c:dLblPos val="outEnd"\/>.*?<c:showVal val="0"\/>\s*<c:showCatName val="1"\/>.*?<c:showPercent val="1"\/>/s);
    expect(doughnut.match(/<c:ser>/g)).toHaveLength(3);
    expect(doughnut).toContain('<c:holeSize val="60"/>');
    expect(doughnut).toContain('<c:v>Series 3</c:v>');
    expect(doughnut).toContain('Sheet1!$D$2:$D$5');
    expect(area).toContain('<c:grouping val="percentStacked"/>');
    expect(line).toContain('<c:lineChart><c:grouping val="stacked"/>');
    expect(scatter).toContain('<c:smooth val="1"/>');
    expect(scatter).toContain('<c:symbol val="none"/>');
    expect(scatter).toMatch(/<c:ser>[^]*?<c:spPr><a:solidFill>[^]*?<a:ln w="28575"/);
    expect(column).toContain('<c:showCatName val="1"/>');
  });

  it('writes combination charts with a secondary axis', async () => {
    const zip = await exportZip(deckWith(chartEl({
      series: [{ name: 'A', values: [1, 2, 3, 4], color: '#112233' }, { name: 'B', values: [2, 3, 4, 5] }, { name: 'C', values: [0.1, 0.2, 0.3, 0.4], kind: 'line', secondary: true, color: '#aa0000' }],
      markers: true, secondaryAxis: { numberFormat: '0%' },
    })));
    const xml = await partXml(zip);
    expect(xml.match(/<c:barChart>/g)).toHaveLength(1);
    expect(xml.match(/<c:lineChart>/g)).toHaveLength(1);
    expect(xml.match(/<c:valAx>/g)).toHaveLength(2);
    expect(xml).toContain('<c:numFmt formatCode="0%" sourceLinked="0"/>');
    // The line's markers are its own colour.
    expect(/<c:lineChart>[\s\S]*<c:marker>[\s\S]*?<a:solidFill><a:srgbClr val="([0-9A-F]+)"/.exec(xml)[1]).toBe('AA0000');
  });

  it('keeps every new setting through a round trip', async () => {
    const els = chartsOf(await roundTrip(deckWith(
      chartEl({ numberFormat: '#,##0.00', dataLabels: true, labelPosition: 'ctr', gapWidth: 80, overlap: 25, valueAxis: { min: 1, max: 9, unit: 2, gridlines: false, title: 'V' }, categoryAxis: { title: 'C' } }),
      chartEl({ kind: 'pie', dataLabels: true, labelContent: ['category', 'percent'], labelPosition: 'outEnd', legend: 'none', categories: ['A', 'B'], series: [{ name: 'S', values: [1, 2] }] }),
      chartEl({ kind: 'doughnut', holeSize: 70 }),
      chartEl({ kind: 'area', grouping: 'percent' }),
      chartEl({ kind: 'line', grouping: 'stacked', smooth: true }),
      chartEl({ kind: 'scatter', categories: ['1', '2', '3', '4'], lines: true, markers: false }),
      chartEl({ series: [{ name: 'A', values: [1, 2, 3, 4] }, { name: 'B', values: [4, 3, 2, 1], kind: 'line', secondary: true }], secondaryAxis: { numberFormat: '0.0', title: 'Right' } }),
    )));
    const [column, pie, doughnut, area, line, scatter, combo] = els.map((e) => e.chart);
    // A pie without a legend keeps its text style (from its data labels).
    expect(els[1].style).toMatchObject({ size: 12, font: els[0].style.font });
    expect(column).toMatchObject({ numberFormat: '#,##0.00', dataLabels: true, labelPosition: 'ctr', gapWidth: 80, overlap: 25, valueAxis: { min: 1, max: 9, unit: 2, gridlines: false, title: 'V' }, categoryAxis: { title: 'C' } });
    expect(column.labelFormat).toBeUndefined();
    expect(pie).toMatchObject({ dataLabels: true, labelContent: ['category', 'percent'], labelPosition: 'outEnd' });
    expect(doughnut.holeSize).toBe(70);
    expect(doughnut.series.map((s) => s.values)).toEqual(chartEl().chart.series.map((s) => s.values));
    expect(area.grouping).toBe('percent');
    expect(line).toMatchObject({ grouping: 'stacked', smooth: true });
    expect(scatter).toMatchObject({ lines: true, markers: false });
    expect(combo.series.map((s) => [s.name, s.kind, s.secondary])).toEqual([['A', undefined, undefined], ['B', 'line', true]]);
    expect(combo.secondaryAxis).toEqual({ numberFormat: '0.0', title: 'Right' });
  });
});

describe('reading chart details back', () => {
  it('leaves pie labels at their automatic place', async () => {
    const [el] = chartsOf(await roundTrip(deckWith(chartEl({ kind: 'pie', dataLabels: true, labelContent: ['percent'] }))));
    expect(el.chart.labelPosition).toBeUndefined();
    expect(el.chart.labelContent).toEqual(['percent']);
  });
});

describe('reading chart details from PowerPoint', () => {
  const ax = (tag, id, cross, more = '') => `<c:${tag}><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/>${tag === 'valAx' && more.includes('SCALE') ? '<c:max val="50"/><c:min val="-10"/>' : ''}</c:scaling>${more.replace('SCALE', '')}<c:crossAx val="${cross}"/></c:${tag}>`;
  const title = (text) => `<c:title><c:tx><c:rich><a:p><a:r><a:t>${text}</a:t></a:r></a:p></c:rich></c:tx></c:title>`;
  const cache = (vals, format) => `<c:val><c:numRef><c:numCache><c:formatCode>${format}</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache></c:numRef></c:val>`;

  it('reads number formats, axis scale, titles, gridlines, gaps and label positions', async () => {
    const labels = '<c:dLbls><c:numFmt formatCode="0.0" sourceLinked="0"/><c:dLblPos val="inBase"/><c:showVal val="1"/></c:dLbls>';
    const deck = await withChartXml(
      `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:tx><c:v>S</c:v></c:tx>${cache([1, 2], '#,##0 €')}</c:ser>${labels}<c:gapWidth val="40"/><c:overlap val="-10"/><c:axId val="1"/><c:axId val="2"/></c:barChart>`
      + ax('catAx', 1, 2, title('Month')) + ax('valAx', 2, 1, `SCALE<c:majorGridlines/>${title('Euros')}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorUnit val="5"/>`),
    );
    expect(chartsOf(deck)[0].chart).toMatchObject({
      numberFormat: '#,##0 €', labelFormat: '0.0', labelPosition: 'inBase', gapWidth: 40, overlap: -10, dataLabels: true,
      valueAxis: { min: -10, max: 50, unit: 5, title: 'Euros' }, categoryAxis: { title: 'Month' },
    });
    expect(chartsOf(deck)[0].chart.valueAxis.gridlines).toBeUndefined();
  });

  it('reads pie percentages and category names, and every doughnut ring', async () => {
    const pie = await withChartXml(`<c:pieChart><c:varyColors val="1"/>${ser('Fruit', [3, 1])}<c:dLbls><c:showVal val="0"/><c:showCatName val="1"/><c:showPercent val="1"/></c:dLbls></c:pieChart>`);
    expect(chartsOf(pie)[0].chart).toMatchObject({ dataLabels: true, labelContent: ['category', 'percent'] });
    const ring = await withChartXml(`<c:doughnutChart><c:varyColors val="1"/>${ser('A', [1, 2])}${ser('B', [3, 4])}<c:holeSize val="75"/></c:doughnutChart>`);
    const { chart } = chartsOf(ring)[0];
    expect(chart.series.map((s) => s.values)).toEqual([[1, 2], [3, 4]]);
    expect(chart.holeSize).toBe(75);
  });

  it('reads scatter lines and markers, and stacked lines', async () => {
    const xy = (ln, symbol) => `<c:ser><c:idx val="0"/><c:spPr>${ln}</c:spPr><c:marker><c:symbol val="${symbol}"/></c:marker><c:xVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:yVal><c:smooth val="1"/></c:ser>`;
    const lines = await withChartXml(`<c:scatterChart><c:scatterStyle val="smoothMarker"/>${xy('<a:ln w="19050"/>', 'none')}</c:scatterChart>`);
    expect(chartsOf(lines)[0].chart).toMatchObject({ kind: 'scatter', lines: true, smooth: true, markers: false });
    const dots = await withChartXml(`<c:scatterChart><c:scatterStyle val="lineMarker"/>${xy('<a:ln><a:noFill/></a:ln>', 'circle')}</c:scatterChart>`);
    expect(chartsOf(dots)[0].chart.lines).toBeUndefined();
    const stacked = await withChartXml(`<c:lineChart><c:grouping val="percentStacked"/>${ser('A', [1, 2])}</c:lineChart>`);
    expect(chartsOf(stacked)[0].chart.grouping).toBe('percent');
  });

  it('reads column and line combinations, with a secondary axis', async () => {
    const deck = await withChartXml(
      `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('Units', [5, 6])}<c:axId val="1"/><c:axId val="2"/></c:barChart>`
      + `<c:lineChart><c:grouping val="standard"/>${ser('Rate', [0.1, 0.2], '<c:marker><c:symbol val="circle"/></c:marker>')}<c:axId val="3"/><c:axId val="4"/></c:lineChart>`
      + ax('catAx', 1, 2) + ax('valAx', 2, 1, '<c:majorGridlines/>') + ax('catAx', 3, 4, '<c:delete val="1"/>') + ax('valAx', 4, 3, `${title('Rate')}<c:numFmt formatCode="0%" sourceLinked="0"/>`),
    );
    const { chart } = chartsOf(deck)[0];
    expect(chart.kind).toBe('column');
    expect(chart.series.map((s) => [s.name, s.kind, s.secondary])).toEqual([['Units', undefined, undefined], ['Rate', 'line', true]]);
    expect(chart.secondaryAxis).toEqual({ title: 'Rate', numberFormat: '0%' });
    expect(chart.markers).toBe(true);
  });
});

describe('untrusted chart details', () => {
  it('keeps well-formed settings and drops the rest', () => {
    const deck = sanitizeDeck({ slides: [{ elements: [{ type: 'chart', chart: {
      kind: 'scatter', markers: false, lines: true, smooth: 'yes', numberFormat: '$#,##0', labelFormat: 'x'.repeat(200),
      labelContent: ['percent', 'nonsense', 'category'], labelPosition: 'middle', gapWidth: 900, overlap: -150, holeSize: 'big',
      valueAxis: { min: 0, max: Infinity, unit: -1, title: 7, gridlines: 'no', numberFormat: '0%' }, categoryAxis: { title: { x: 1 } }, secondaryAxis: { numberFormat: '0%' },
      series: [{ name: 'S', values: [1], kind: 'pie', secondary: 'yes' }, { name: 'T', values: [2], kind: 'line', secondary: true }],
    } }] }] });
    const { chart } = deck.slides[0].elements[0];
    expect(chart).toMatchObject({ markers: false, lines: true, numberFormat: '$#,##0', labelContent: ['category', 'percent'], gapWidth: 500, overlap: -100, valueAxis: { min: 0, title: '7' }, secondaryAxis: { numberFormat: '0%' } });
    for (const key of ['smooth', 'labelFormat', 'labelPosition', 'holeSize', 'categoryAxis']) expect(chart[key]).toBeUndefined();
    expect(chart.valueAxis).toEqual({ min: 0, title: '7' });
    expect(chart.series).toEqual([{ name: 'S', values: [1] }, { name: 'T', values: [2], kind: 'line', secondary: true }]);
  });
});
