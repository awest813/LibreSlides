import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { newDeck, newTable } from '../../src/model/deck.js';
import { normalizeTable, anchorOf, insertRow, insertCol, deleteRow, deleteCol, mergeCells, unmergeCell } from '../../src/model/tables.js';
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
