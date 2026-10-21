import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { newDeck, newTable } from '../../src/model/deck.js';
import { normalizeTable, anchorOf, insertRow, insertCol, deleteRow, deleteCol, mergeCells, unmergeCell, setBorders, setMargins, cellBorder, blockRange, CELL_MARGIN } from '../../src/model/tables.js';
import { renderElement } from '../../src/render/render.js';
import { themeById } from '../../src/model/themes.js';
import { buildPptx } from '../../src/io/pptx-export.js';
import { importPptx } from '../../src/io/pptx-import.js';

/** A rows × cols table whose cells say where they are ("r,c"). */
function grid(rows, cols) {
  const t = newTable(rows, cols, 0, 0, 300, 30 * rows);
  t.cells.forEach((row, r) => row.forEach((cell, c) => { cell.text = `${r},${c}`; }));
  return t;
}
/** The grid as text, '·' for covered cells and "text+rows×cols" for merged ones. */
const show = (t) => t.cells.map((row) => row.map((c) => (c.merged ? '·' : c.rowSpan || c.colSpan ? `${c.text}+${c.rowSpan || 1}x${c.colSpan || 1}` : c.text)).join(' | '));

describe('merged table cells', () => {
  it('merges right and down, keeping the text, and unmerges', () => {
    const t = grid(3, 3);
    mergeCells(t, 0, 0, 'right');
    expect(show(t)[0]).toBe('0,00,1+1x2 | · | 0,2');
    mergeCells(t, 0, 1, 'down'); // from a covered cell: the whole block grows
    expect(show(t).slice(0, 2)).toEqual(['0,00,11,01,1+2x2 | · | 0,2', '· | · | 1,2']);
    expect(anchorOf(t, 1, 1)).toMatchObject({ r: 0, c: 0 });
    unmergeCell(t, 1, 1);
    expect(show(t).slice(0, 2)).toEqual(['0,00,11,01,1 |  | 0,2', ' |  | 1,2']);
  });

  it('grows a block when a row or column is added inside it, and shrinks it on delete', () => {
    const t = grid(3, 3);
    mergeCells(t, 0, 0, 'down');
    insertRow(t, 1);
    expect(t.cells[0][0].rowSpan).toBe(3);
    expect(t.cells[1].map((c) => !!c.merged)).toEqual([true, false, false]);
    insertRow(t, 0); // above the block: it just moves down
    expect(t.cells[1][0].rowSpan).toBe(3);
    deleteRow(t, 1); // the block's top row: the next row takes its text
    expect(show(t)[1]).toBe('0,01,0+2x1 |  | ');
    mergeCells(t, 0, 0, 'right');
    insertCol(t, 1);
    expect(t.cells[0][0].colSpan).toBe(3);
    deleteCol(t, 0);
    expect(t.cells[0][0]).toMatchObject({ colSpan: 2 });
    expect(t.colWidths).toHaveLength(3);
    expect(t.colWidths.reduce((s, v) => s + v, 0)).toBeCloseTo(1);
  });

  it('keeps blocks inside the table and apart from each other', () => {
    const t = grid(2, 2);
    t.cells[0][0].colSpan = 5;
    t.cells[1][1].rowSpan = 3;
    normalizeTable(t);
    expect(show(t)).toEqual(['0,0+1x2 | ·', '1,0 | 1,1']);
  });

  it('draws merged cells across rows and columns', () => {
    const t = grid(2, 3);
    mergeCells(t, 0, 0, 'right');
    mergeCells(t, 0, 2, 'down');
    t.rowHeights = [0.25, 0.75];
    const deck = newDeck();
    const node = renderElement(t, { theme: themeById(deck.themeId), deck });
    const rows = [...node.querySelectorAll('tr')];
    expect(rows.map((tr) => tr.children.length)).toEqual([2, 2]);
    expect(rows[0].children[0].colSpan).toBe(2);
    expect(rows[0].children[1].rowSpan).toBe(2);
    expect(rows[1].style.height).toBe('75%');
  });
});

describe('cell borders and margins', () => {
  const red = { color: '#ff0000', width: 2, dash: 'solid' };
  /** Which sides of each cell have a border of their own ("trbl", '-' for none of its own). */
  const sides = (t) => t.cells.map((row) => row.map((c) => ['t', 'r', 'b', 'l'].filter((s) => c.borders?.[s]).join('') || '-').join(' '));

  it('draws outside borders round a range, with the cells across each edge in step', () => {
    const t = grid(3, 3);
    setBorders(t, { r0: 1, c0: 1, r1: 1, c1: 1 }, 'outside', red);
    expect(sides(t)).toEqual(['- b -', 'r trbl l', '- t -']);
    expect(cellBorder(t, t.cells[0][1], 'b')).toEqual(red);
    expect(cellBorder(t, t.cells[0][0], 'b')).toEqual(t.stroke); // the table's border
  });

  it('draws inside borders, one side, or none, and takes in whole merged blocks', () => {
    const t = grid(2, 3);
    setBorders(t, { r0: 0, c0: 0, r1: 1, c1: 2 }, 'inside', red);
    expect(sides(t)).toEqual(['rb rbl bl', 'tr trl tl']);
    setBorders(t, { r0: 0, c0: 0, r1: 1, c1: 2 }, 'none');
    expect(t.cells.flat().every((c) => ['t', 'r', 'b', 'l'].every((s) => c.borders[s].color === 'none'))).toBe(true);
    const u = grid(3, 3);
    mergeCells(u, 0, 0, 'down');
    expect(blockRange(u, 1, 1, 1, 0)).toEqual({ r0: 0, c0: 0, r1: 1, c1: 1 });
    setBorders(u, { r0: 1, c0: 0, r1: 1, c1: 0 }, 'bottom', red); // the bottom of the merged block
    expect(u.cells[0][0].borders.b).toEqual(red);
    expect(u.cells[2][0].borders.t).toEqual(red);
  });

  it('sets margins, dropping them when they are back to the defaults', () => {
    const t = grid(2, 2);
    setMargins(t, { r0: 0, c0: 0, r1: 0, c1: 1 }, { l: 20 });
    expect(t.cells[0][1].margin).toEqual({ ...CELL_MARGIN, l: 20 });
    expect(t.cells[1][0].margin).toBeUndefined();
    setMargins(t, { r0: 0, c0: 0, r1: 0, c1: 1 }, { l: CELL_MARGIN.l });
    expect(t.cells[0][1].margin).toBeUndefined();
  });

  it('draws each cell\'s own borders and margins', () => {
    const t = grid(1, 2);
    t.stroke = { color: '#0000ff', width: 1, dash: 'solid' };
    t.cells[0][0].borders = { r: { color: '#ff0000', width: 3, dash: 'dot' }, b: { color: 'none', width: 0, dash: 'solid' } };
    t.cells[0][0].margin = { t: 1, r: 2, b: 3, l: 4 };
    const deck = newDeck();
    const td = renderElement(t, { theme: themeById(deck.themeId), deck }).querySelector('td');
    expect(td.style.borderRight).toBe('3px dotted rgb(255, 0, 0)');
    expect(td.style.borderBottomStyle).toBe('none');
    expect(td.style.borderTop).toBe('1px solid rgb(0, 0, 255)');
    expect(td.style.padding).toBe('1px 2px 3px 4px');
  });
});

describe('tables in PowerPoint files', () => {
  const toBuffer = (blob) => blob.arrayBuffer();
  const deckWith = (el) => {
    const deck = newDeck({ slides: [] });
    deck.slides.push({ id: 's', layout: 'blank', background: null, elements: [el], notes: '', transition: { type: 'none', duration: 0.6 }, hidden: false });
    return deck;
  };

  it('keeps merged cells, row heights and cell alignment', async () => {
    const t = grid(3, 3);
    t.header = false;
    mergeCells(t, 0, 0, 'right');
    mergeCells(t, 1, 2, 'down');
    t.cells[1][0].valign = 'bottom';
    t.rowHeights = [0.2, 0.3, 0.5];
    const zip = await JSZip.loadAsync(await toBuffer(await buildPptx(deckWith(t))));
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');
    expect(xml).toContain('gridSpan="2"');
    expect(xml).toContain('hMerge="1"');
    expect(xml).toContain('rowSpan="2"');
    expect(xml).toContain('vMerge="1"');
    const back = (await importPptx(await zip.generateAsync({ type: 'arraybuffer' }))).slides[0].elements[0];
    expect(show(back).map((r) => r.replace(/<\/?p>/g, ''))).toEqual(['0,00,1+1x2 | · | 0,2', '1,0 | 1,1 | 1,22,2+2x1', '2,0 | 2,1 | ·']);
    expect(back.cells[1][0].valign).toBe('bottom');
    expect(back.rowHeights.map((v) => Math.round(v * 10) / 10)).toEqual([0.2, 0.3, 0.5]);
  });

  it('keeps cell borders and margins', async () => {
    const t = grid(3, 3);
    t.header = false;
    t.stroke = { color: '#999999', width: 1, dash: 'solid' };
    mergeCells(t, 1, 1, 'right');
    setBorders(t, { r0: 0, c0: 0, r1: 0, c1: 2 }, 'bottom', { color: '#ff0000', width: 3, dash: 'solid' });
    setBorders(t, { r0: 1, c0: 1, r1: 1, c1: 2 }, 'outside', { color: '#00aa00', width: 1.5, dash: 'dot' });
    setBorders(t, { r0: 2, c0: 0, r1: 2, c1: 0 }, 'none');
    setMargins(t, { r0: 0, c0: 0, r1: 0, c1: 0 }, { l: 20, t: 0 });
    setMargins(t, { r0: 2, c0: 2, r1: 2, c1: 2 }, { r: 12.5 });
    const zip = await JSZip.loadAsync(await toBuffer(await buildPptx(deckWith(t))));
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');
    expect(xml).toContain('marL="254000" marR="91440" marT="0" marB="45720"');
    expect(xml).toContain('<a:prstDash val="sysDot"/>');
    expect(xml).not.toContain('sysDash');
    const back = (await importPptx(await zip.generateAsync({ type: 'arraybuffer' }))).slides[0].elements[0];
    expect(back.stroke).toEqual(t.stroke);
    const view = (x) => x.cells.map((row) => row.map((c) => (c.merged ? null : { borders: c.borders, margin: c.margin })));
    expect(view(back)).toEqual(JSON.parse(JSON.stringify(view(t))));
  });

  it('reads borders cell by cell and side by side from PowerPoint tables and their styles', async () => {
    const t = grid(3, 2);
    t.header = false;
    const zip = await JSZip.loadAsync(await toBuffer(await buildPptx(deckWith(t))));
    let xml = await zip.file('ppt/slides/slide1.xml').async('string');
    // As PowerPoint writes it: a style with lines between the rows and a thick one under the header,
    // a cell with no top border, one with a blue dashed right border, and a cell with margins of its own.
    const ln = (tag, inner) => `<a:${tag} w="25400"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>${inner}</a:${tag}>`;
    const cellPr = [
      '<a:tcPr/>', `<a:tcPr marL="0" marR="182880">${ln('lnR', '<a:prstDash val="dash"/>')}</a:tcPr>`,
      '<a:tcPr><a:lnT w="12700"><a:noFill/></a:lnT></a:tcPr>', '<a:tcPr/>',
      '<a:tcPr/>', '<a:tcPr/>',
    ];
    let i = 0;
    xml = xml.replace(/<a:tcPr[^>]*>[\s\S]*?<\/a:tcPr>|<a:tcPr[^>]*\/>/g, () => cellPr[i++] || '<a:tcPr/>');
    xml = xml.replace(/<a:tblPr[^>]*\/>|<a:tblPr[^>]*>[\s\S]*?<\/a:tblPr>/, '<a:tblPr firstRow="1"><a:tableStyleId>{TEST-STYLE}</a:tableStyleId></a:tblPr>');
    zip.file('ppt/slides/slide1.xml', xml);
    const line = (w, color) => `<a:ln w="${w}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:ln>`;
    zip.file('ppt/tableStyles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{TEST-STYLE}"><a:tblStyle styleId="{TEST-STYLE}" styleName="Test">
<a:wholeTbl><a:tcStyle><a:tcBdr><a:left><a:ln><a:noFill/></a:ln></a:left><a:right><a:ln><a:noFill/></a:ln></a:right><a:top>${line(12700, '000000')}</a:top><a:bottom>${line(12700, '000000')}</a:bottom><a:insideH>${line(12700, '000000')}</a:insideH><a:insideV><a:ln><a:noFill/></a:ln></a:insideV></a:tcBdr></a:tcStyle></a:wholeTbl>
<a:firstRow><a:tcStyle><a:tcBdr><a:bottom>${line(38100, '000000')}</a:bottom></a:tcBdr></a:tcStyle></a:firstRow></a:tblStyle></a:tblStyleLst>`);
    const back = (await importPptx(await zip.generateAsync({ type: 'arraybuffer' }))).slides[0].elements[0];
    const thin = { color: '#000000', width: 1, dash: 'solid' };
    const none = { color: 'none', width: 0, dash: 'solid' };
    expect(back.stroke).toEqual(none); // most sides: the style's blank verticals
    const b = back.cells.map((row) => row.map((c) => c.borders));
    expect(b[0][0]).toEqual({ t: thin, b: { ...thin, width: 3 } });
    expect(b[0][1]).toEqual({ t: thin, r: { color: '#0000ff', width: 2, dash: 'dash' }, b: { ...thin, width: 3 } });
    expect(b[1][0]).toEqual({ b: thin }); // its top: none of its own
    expect(b[2][1]).toEqual({ t: thin, b: thin });
    expect(back.cells[0][1].margin).toEqual({ t: 3.6, r: 14.4, b: 3.6, l: 0 });
    expect(back.cells[0][0].margin).toBeUndefined();
  });

  it('paints the cells the way PowerPoint\'s table style does', async () => {
    const t = grid(4, 2);
    t.header = false;
    const zip = await JSZip.loadAsync(await toBuffer(await buildPptx(deckWith(t))));
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');
    // As PowerPoint writes it: the default style, with a header row and banded rows.
    zip.file('ppt/slides/slide1.xml', xml.replace(/<a:tblPr[^>]*\/>|<a:tblPr[^>]*>[\s\S]*?<\/a:tblPr>/, '<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr>'));
    const back = (await importPptx(await zip.generateAsync({ type: 'arraybuffer' }))).slides[0].elements[0];
    const fills = back.cells.map((row) => row[0].fill);
    expect(fills[0]).toBe(back.fill); // accent 1
    expect(fills[1]).not.toBe(fills[2]); // banded
    expect(fills[1]).toBe(fills[3]);
    expect(back.cells[0][0].text).toMatch(/font-weight: 700/);
    expect(back.header).toBe(false);
  });
});
