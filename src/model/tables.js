/**
 * Table structure: merged cells and row and column changes that keep them
 * consistent. A merged block is held by its top-left cell, which has
 * `rowSpan` and/or `colSpan`; the cells it covers stay in the grid (so every
 * row has a cell for every column) with `merged: true`, and draw nothing.
 * `colWidths` and the optional `rowHeights` are fractions of the table's size.
 * Cells may also have borders and margins of their own (see setBorders).
 */

const blank = () => ({ text: '' });

/** Recompute which cells are covered from the merged blocks, keeping spans inside the table. */
export function normalizeTable(el) {
  const rows = el.cells.length;
  const cols = el.cells[0]?.length || 0;
  for (const row of el.cells) for (const cell of row) delete cell.merged;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const cell = el.cells[r][c];
      if (cell.merged) {
        delete cell.rowSpan;
        delete cell.colSpan;
        continue;
      }
      let rs = Math.max(1, Math.min(rows - r, Math.floor(cell.rowSpan) || 1));
      let cs = Math.max(1, Math.min(cols - c, Math.floor(cell.colSpan) || 1));
      // A block can't reach into one that's already there.
      while (cs > 1 && el.cells[r].slice(c + 1, c + cs).some((x) => x.merged)) cs--;
      while (rs > 1 && el.cells[r + rs - 1].slice(c, c + cs).some((x) => x.merged)) rs--;
      if (rs > 1) cell.rowSpan = rs;
      else delete cell.rowSpan;
      if (cs > 1) cell.colSpan = cs;
      else delete cell.colSpan;
      for (let i = r; i < r + rs; i++) {
        for (let j = c; j < c + cs; j++) {
          if (i === r && j === c) continue;
          const covered = el.cells[i][j];
          covered.merged = true;
          delete covered.rowSpan;
          delete covered.colSpan;
        }
      }
    }
  }
  return el;
}

/** The cell whose block covers row r, column c (itself when not merged). */
export function anchorOf(el, r, c) {
  for (let i = r; i >= 0; i--) {
    for (let j = c; j >= 0; j--) {
      const cell = el.cells[i][j];
      if (!cell.merged && i + (cell.rowSpan || 1) > r && j + (cell.colSpan || 1) > c) return { r: i, c: j, cell };
    }
  }
  return { r, c, cell: el.cells[r][c] };
}

const fractions = (list, at, count) => {
  if (!list) return list;
  const out = [...list];
  if (count > 0) out.splice(at, 0, ...Array(count).fill(1 / out.length));
  else out.splice(at, -count);
  const sum = out.reduce((s, v) => s + v, 0) || 1;
  return out.map((v) => v / sum);
};

/** Insert a row at index `at` (cells inside a merged block join it). */
export function insertRow(el, at) {
  const cols = el.cells[0].length;
  const row = Array.from({ length: cols }, blank);
  for (let c = 0; c < cols; c++) {
    if (at === 0 || at >= el.cells.length) continue;
    const a = anchorOf(el, at - 1, c);
    if (a.r < at && a.r + (a.cell.rowSpan || 1) > at) {
      row[c].merged = true;
      if (a.c === c) a.cell.rowSpan = (a.cell.rowSpan || 1) + 1;
    }
  }
  el.cells.splice(at, 0, row);
  el.rowHeights = fractions(el.rowHeights, at, 1);
  return normalizeTable(el);
}

/** Insert a column at index `at`. */
export function insertCol(el, at) {
  const cols = el.cells[0].length;
  el.cells.forEach((row, r) => {
    const cell = blank();
    if (at > 0 && at < cols) {
      const a = anchorOf(el, r, at - 1);
      if (a.c < at && a.c + (a.cell.colSpan || 1) > at) {
        cell.merged = true;
        if (a.r === r) a.cell.colSpan = (a.cell.colSpan || 1) + 1;
      }
    }
    row.splice(at, 0, cell);
  });
  el.colWidths = fractions(el.colWidths, at, 1);
  return normalizeTable(el);
}

/** Delete row r; a block starting there moves its text to the row below. */
export function deleteRow(el, r) {
  if (el.cells.length < 2) return el;
  const cols = el.cells[0].length;
  for (let c = 0; c < cols; c++) {
    const a = anchorOf(el, r, c);
    if (a.c !== c) continue;
    const span = a.cell.rowSpan || 1;
    if (span < 2) continue;
    if (a.r === r) {
      // The block loses its top row: the cell below takes over.
      const below = el.cells[r + 1][c];
      Object.assign(below, { ...a.cell, merged: false, rowSpan: span - 1 });
    } else a.cell.rowSpan = span - 1;
  }
  el.cells.splice(r, 1);
  el.rowHeights = fractions(el.rowHeights, r, -1);
  return normalizeTable(el);
}

/** Delete column c; a block starting there moves its text to the column on the right. */
export function deleteCol(el, c) {
  if (el.cells[0].length < 2) return el;
  el.cells.forEach((row, r) => {
    const a = anchorOf(el, r, c);
    if (a.r !== r) return;
    const span = a.cell.colSpan || 1;
    if (span < 2) return;
    if (a.c === c) Object.assign(row[c + 1], { ...a.cell, merged: false, colSpan: span - 1 });
    else a.cell.colSpan = span - 1;
  });
  for (const row of el.cells) row.splice(c, 1);
  el.colWidths = fractions(el.colWidths, c, -1);
  return normalizeTable(el);
}

/** Merge the block holding (r, c) with the cells to its right ('right') or below ('down'). */
export function mergeCells(el, r, c, direction) {
  const a = anchorOf(el, r, c);
  const rs = a.cell.rowSpan || 1;
  const cs = a.cell.colSpan || 1;
  const next = direction === 'right' ? { r: a.r, c: a.c + cs } : { r: a.r + rs, c: a.c };
  if (next.r >= el.cells.length || next.c >= el.cells[0].length) return false;
  const b = anchorOf(el, next.r, next.c);
  const bottom = Math.max(a.r + rs, b.r + (b.cell.rowSpan || 1));
  const right = Math.max(a.c + cs, b.c + (b.cell.colSpan || 1));
  // Text of every cell joining the block moves into it.
  const texts = [];
  for (let i = a.r; i < bottom; i++) {
    for (let j = a.c; j < right; j++) {
      const cell = el.cells[i][j];
      if (cell === a.cell || cell.merged) continue;
      if (cell.text) texts.push(cell.text);
      cell.text = '';
      cell.merged = true;
    }
  }
  if (texts.length) a.cell.text = [a.cell.text, ...texts].filter(Boolean).join('');
  a.cell.rowSpan = bottom - a.r;
  a.cell.colSpan = right - a.c;
  normalizeTable(el);
  return true;
}

/** Split the block holding (r, c) back into single cells. */
export function unmergeCell(el, r, c) {
  const a = anchorOf(el, r, c);
  if ((a.cell.rowSpan || 1) < 2 && (a.cell.colSpan || 1) < 2) return false;
  delete a.cell.rowSpan;
  delete a.cell.colSpan;
  normalizeTable(el);
  return true;
}

// ------------------------------------------------------------------ borders and margins

/*
 * A cell's `borders` ({ t?, r?, b?, l? }, strokes) stand in for the table's
 * `stroke` side by side. An edge two cells share is kept on both, as
 * PowerPoint keeps it, so setting one side sets the facing side across it.
 * `margin` ({ t, r, b, l }, points) is the space around the cell's text.
 */

// PowerPoint's default cell margins: 0.1 in left and right, 0.05 in above and below.
export const CELL_MARGIN = { t: 3.6, r: 7.2, b: 3.6, l: 7.2 };
export const NO_BORDER = { color: 'none', width: 0, dash: 'solid' };
export const SIDES = ['t', 'r', 'b', 'l'];
const OPPOSITE = { t: 'b', r: 'l', b: 't', l: 'r' };

/** A cell's margins, the defaults standing in for any it doesn't set. */
export const cellMargin = (cell) => ({ ...CELL_MARGIN, ...(cell?.margin || {}) });

/** The stroke on one side ('t', 'r', 'b' or 'l') of a cell: its own, else the table's. */
export const cellBorder = (el, cell, side) => cell?.borders?.[side] || el.stroke || NO_BORDER;

/** Rows r0–r1 and columns c0–c1 (either way round), grown to take in whole merged blocks. */
export function blockRange(el, r0, c0, r1, c1) {
  let range = { r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1) };
  for (let grown = true; grown;) {
    grown = false;
    for (let r = range.r0; r <= range.r1; r++) {
      for (let c = range.c0; c <= range.c1; c++) {
        const a = anchorOf(el, r, c);
        const next = {
          r0: Math.min(range.r0, a.r), c0: Math.min(range.c0, a.c),
          r1: Math.max(range.r1, a.r + (a.cell.rowSpan || 1) - 1), c1: Math.max(range.c1, a.c + (a.cell.colSpan || 1) - 1),
        };
        if (next.r0 !== range.r0 || next.c0 !== range.c0 || next.r1 !== range.r1 || next.c1 !== range.c1) {
          range = next;
          grown = true;
        }
      }
    }
  }
  return range;
}

/** The blocks (merged or single cells) of a range: { r, c, rs, cs, cell }. */
function blocksIn(el, range) {
  const out = [];
  for (let r = range.r0; r <= range.r1; r++) {
    for (let c = range.c0; c <= range.c1; c++) {
      const cell = el.cells[r][c];
      if (!cell.merged) out.push({ r, c, rs: cell.rowSpan || 1, cs: cell.colSpan || 1, cell });
    }
  }
  return out;
}

/** Set one side of a block, and the facing side of each block across that edge. */
function setSide(el, block, side, stroke) {
  const set = (cell, s) => { cell.borders = { ...(cell.borders || {}), [s]: { ...stroke } }; };
  set(block.cell, side);
  const { r, c, rs, cs } = block;
  const rows = el.cells.length;
  const cols = el.cells[0].length;
  const across = [];
  if (side === 't' && r > 0) for (let j = c; j < c + cs; j++) across.push([r - 1, j]);
  if (side === 'b' && r + rs < rows) for (let j = c; j < c + cs; j++) across.push([r + rs, j]);
  if (side === 'l' && c > 0) for (let i = r; i < r + rs; i++) across.push([i, c - 1]);
  if (side === 'r' && c + cs < cols) for (let i = r; i < r + rs; i++) across.push([i, c + cs]);
  const done = new Set();
  for (const [i, j] of across) {
    const { cell } = anchorOf(el, i, j);
    if (done.has(cell)) continue;
    done.add(cell);
    set(cell, OPPOSITE[side]);
  }
}

// What PowerPoint's Borders menu offers.
export const BORDER_PRESETS = [
  ['all', 'All borders'], ['outside', 'Outside borders'], ['inside', 'Inside borders'], ['none', 'No borders'],
  ['top', 'Top border'], ['bottom', 'Bottom border'], ['left', 'Left border'], ['right', 'Right border'],
];

/**
 * Draw `preset` borders (see BORDER_PRESETS) with `stroke` around and
 * between the cells of a range ({ r0, c0, r1, c1 }).
 */
export function setBorders(el, range, preset, stroke) {
  const R = blockRange(el, range.r0, range.c0, range.r1, range.c1);
  const pen = preset === 'none' || !stroke?.width || !stroke.color || stroke.color === 'none' ? NO_BORDER
    : { color: stroke.color, width: stroke.width, dash: stroke.dash || 'solid' };
  const one = { top: 't', bottom: 'b', left: 'l', right: 'r' }[preset];
  for (const block of blocksIn(el, R)) {
    const edge = { t: block.r === R.r0, b: block.r + block.rs - 1 === R.r1, l: block.c === R.c0, r: block.c + block.cs - 1 === R.c1 };
    for (const side of SIDES) {
      const on = preset === 'all' || preset === 'none' || (preset === 'outside' && edge[side]) || (preset === 'inside' && !edge[side]) || (one === side && edge[side]);
      if (on) setSide(el, block, side, pen);
    }
  }
  return el;
}

/** Change some of the margins ({ t?, r?, b?, l? }, points) of the cells in a range. */
export function setMargins(el, range, margin) {
  const R = blockRange(el, range.r0, range.c0, range.r1, range.c1);
  for (const { cell } of blocksIn(el, R)) {
    const m = { ...cellMargin(cell), ...margin };
    if (SIDES.every((s) => m[s] === CELL_MARGIN[s])) delete cell.margin;
    else cell.margin = m;
  }
  return el;
}
