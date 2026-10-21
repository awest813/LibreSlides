import { uid } from './deck.js';

/**
 * A diagram's text as an outline, as in PowerPoint's SmartArt text pane:
 * nodes [{ id, text, level }] in order, each at most one level deeper than
 * the one before it. A node's "subtree" is it and the deeper nodes after it.
 * The editing operations return a new list (and where the node they moved
 * went), leaving the one they were given alone.
 */
export const MAX_LEVEL = 8;

/** Levels that make an outline: the first node at the top, none more than one deeper than the one before. */
export function normalizeLevels(nodes) {
  let prev = -1;
  return nodes.map((n) => {
    const level = Math.max(0, Math.min(MAX_LEVEL, prev + 1, Math.round(+n.level) || 0));
    prev = level;
    return level === n.level ? n : { ...n, level };
  });
}

/** Where the subtree of node `i` ends (the index after its last descendant). */
export function subtreeEnd(nodes, i) {
  let j = i + 1;
  while (j < nodes.length && nodes[j].level > nodes[i].level) j++;
  return j;
}

/** The index of the sibling just before node `i` (same level, same parent), or -1. */
function previousSibling(nodes, i) {
  for (let j = i - 1; j >= 0; j--) {
    if (nodes[j].level < nodes[i].level) return -1;
    if (nodes[j].level === nodes[i].level) return j;
  }
  return -1;
}

const shift = (list, by) => list.map((n) => ({ ...n, level: Math.max(0, Math.min(MAX_LEVEL, n.level + by)) }));

/** A new node after node `i` and its subtree, at its level (at the end, at the top level, when `i` is -1). */
export function addNode(nodes, i, text = '') {
  const at = i < 0 ? nodes.length : subtreeEnd(nodes, i);
  const node = { id: uid(), text, level: i < 0 ? 0 : nodes[i].level };
  return { nodes: [...nodes.slice(0, at), node, ...nodes.slice(at)], index: at };
}

/** Take node `i` out; the nodes under it move up a level, as PowerPoint does. */
export function removeNode(nodes, i) {
  if (i < 0 || i >= nodes.length) return { nodes, index: i };
  const end = subtreeEnd(nodes, i);
  const out = normalizeLevels([...nodes.slice(0, i), ...shift(nodes.slice(i + 1, end), -1), ...nodes.slice(end)]);
  return { nodes: out, index: Math.min(i, out.length - 1) };
}

/** Node `i` and its subtree one level deeper, under the sibling before it (no change for a first child). */
export function demote(nodes, i) {
  if (previousSibling(nodes, i) < 0 || nodes[i].level >= MAX_LEVEL) return { nodes, index: i };
  const end = subtreeEnd(nodes, i);
  return { nodes: [...nodes.slice(0, i), ...shift(nodes.slice(i, end), 1), ...nodes.slice(end)], index: i };
}

/** Node `i` and its subtree one level up (no change at the top level). */
export function promote(nodes, i) {
  if (!nodes[i]?.level) return { nodes, index: i };
  const end = subtreeEnd(nodes, i);
  return { nodes: normalizeLevels([...nodes.slice(0, i), ...shift(nodes.slice(i, end), -1), ...nodes.slice(end)]), index: i };
}

/** Node `i` and its subtree before the sibling before it. */
export function moveUp(nodes, i) {
  const j = previousSibling(nodes, i);
  if (j < 0) return { nodes, index: i };
  const end = subtreeEnd(nodes, i);
  return { nodes: [...nodes.slice(0, j), ...nodes.slice(i, end), ...nodes.slice(j, i), ...nodes.slice(end)], index: j };
}

/** Node `i` and its subtree after the sibling after it. */
export function moveDown(nodes, i) {
  const end = subtreeEnd(nodes, i);
  if (end >= nodes.length || nodes[end].level !== nodes[i].level) return { nodes, index: i };
  const after = subtreeEnd(nodes, end);
  return { nodes: [...nodes.slice(0, i), ...nodes.slice(end, after), ...nodes.slice(i, end), ...nodes.slice(after)], index: i + after - end };
}

/** The outline as text: a line per node, indented with a tab per level. */
export const outlineText = (nodes) => nodes.map((n) => '\t'.repeat(n.level) + String(n.text).replace(/[\r\n\t]+/g, ' ')).join('\n');

/**
 * Nodes from outline text (a line each; leading tabs, or pairs of spaces,
 * give the level). Nodes of `previous` keep their ids where their lines
 * are still there, or a line took their place (so an edited node stays the
 * same node). Blank text gives no nodes.
 */
export function parseOutline(text, previous = []) {
  if (!String(text).trim()) return [];
  const lines = String(text).replace(/\r\n?/g, '\n').split('\n').map((line) => {
    const lead = /^[\t ]*/.exec(line)[0];
    const level = [...lead].reduce((n, c) => n + (c === '\t' ? 2 : 1), 0) >> 1;
    return { text: line.slice(lead.length).trimEnd(), level };
  });
  // Lines the same as before keep their node, in order; the lines between keep the nodes between.
  const ids = new Array(lines.length).fill(null);
  const anchors = [[-1, -1]];
  let from = 0;
  lines.forEach((l, i) => {
    for (let j = from; j < previous.length; j++) {
      if (previous[j].text === l.text && previous[j].level === l.level) {
        ids[i] = previous[j].id;
        anchors.push([i, j]);
        from = j + 1;
        return;
      }
    }
  });
  anchors.push([lines.length, previous.length]);
  for (let a = 0; a < anchors.length - 1; a++) {
    const [i0, j0] = anchors[a];
    const [i1, j1] = anchors[a + 1];
    for (let k = 1; i0 + k < i1 && j0 + k < j1; k++) ids[i0 + k] = previous[j0 + k].id;
  }
  return normalizeLevels(lines.map((l, i) => ({ id: ids[i] || uid(), text: l.text, level: l.level })));
}

/** The node a line of outline text is (0-based), from a caret position in it. */
export const lineAt = (text, caret) => String(text).slice(0, caret).split('\n').length - 1;

/**
 * The outline as a tree: [{ node, children: [...] }] for its top-level
 * nodes. (Layouts that use the text inside a shape see deeper nodes as its
 * bullet points instead.)
 */
export function outlineTree(nodes) {
  const roots = [];
  const stack = [];
  for (const node of normalizeLevels(nodes)) {
    const item = { node, children: [] };
    while (stack.length > node.level) stack.pop();
    (stack.length ? stack[stack.length - 1].children : roots).push(item);
    stack.push(item);
  }
  return roots;
}
