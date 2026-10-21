/**
 * Shape geometry. Shape ids are the OOXML preset names (prstGeom), so they
 * map one-to-one to PowerPoint shapes on export and import.
 */
export const SHAPES = [
  { id: 'rect', name: 'Rectangle', group: 'Basic' },
  { id: 'roundRect', name: 'Rounded rectangle', group: 'Basic' },
  { id: 'ellipse', name: 'Oval', group: 'Basic' },
  { id: 'triangle', name: 'Triangle', group: 'Basic' },
  { id: 'rtTriangle', name: 'Right triangle', group: 'Basic' },
  { id: 'diamond', name: 'Diamond', group: 'Basic' },
  { id: 'parallelogram', name: 'Parallelogram', group: 'Basic' },
  { id: 'trapezoid', name: 'Trapezoid', group: 'Basic' },
  { id: 'pentagon', name: 'Pentagon', group: 'Basic' },
  { id: 'hexagon', name: 'Hexagon', group: 'Basic' },
  { id: 'octagon', name: 'Octagon', group: 'Basic' },
  { id: 'plus', name: 'Plus', group: 'Basic' },
  { id: 'heart', name: 'Heart', group: 'Basic' },
  { id: 'star4', name: '4-point star', group: 'Basic' },
  { id: 'star5', name: '5-point star', group: 'Basic' },
  { id: 'wave', name: 'Wave', group: 'Basic' },
  { id: 'rightArrow', name: 'Right arrow', group: 'Arrows' },
  { id: 'leftArrow', name: 'Left arrow', group: 'Arrows' },
  { id: 'upArrow', name: 'Up arrow', group: 'Arrows' },
  { id: 'downArrow', name: 'Down arrow', group: 'Arrows' },
  { id: 'chevron', name: 'Chevron', group: 'Arrows' },
  { id: 'homePlate', name: 'Pentagon arrow', group: 'Arrows' },
  { id: 'wedgeRectCallout', name: 'Speech bubble', group: 'Callouts' },
  { id: 'wedgeEllipseCallout', name: 'Thought bubble', group: 'Callouts' },
];

const f = (n) => Math.round(n * 100) / 100;
const poly = (pts) => `M${pts.map(([x, y]) => `${f(x)} ${f(y)}`).join('L')}Z`;

function star(w, h, points, inner) {
  const pts = [];
  for (let i = 0; i < points * 2; i++) {
    const a = -Math.PI / 2 + (i * Math.PI) / points;
    const r = i % 2 ? inner : 1;
    pts.push([w / 2 + (Math.cos(a) * r * w) / 2, h / 2 + (Math.sin(a) * r * h) / 2]);
  }
  return poly(pts);
}

function regular(w, h, n) {
  const pts = [];
  for (let i = 0; i < n; i++) {
    const a = -Math.PI / 2 + (i * 2 * Math.PI) / n;
    pts.push([w / 2 + (Math.cos(a) * w) / 2, h / 2 + (Math.sin(a) * h) / 2]);
  }
  return poly(pts);
}

const ellipse = (w, h) => `M0 ${f(h / 2)}A${f(w / 2)} ${f(h / 2)} 0 1 1 ${f(w)} ${f(h / 2)}A${f(w / 2)} ${f(h / 2)} 0 1 1 0 ${f(h / 2)}Z`;

/** SVG path data for a shape filling a w × h box. */
export function shapePath(kind, w, h) {
  const m = Math.min(w, h);
  switch (kind) {
    case 'roundRect': {
      const r = Math.min(m * 0.16, w / 2, h / 2);
      return `M${f(r)} 0H${f(w - r)}A${f(r)} ${f(r)} 0 0 1 ${f(w)} ${f(r)}V${f(h - r)}A${f(r)} ${f(r)} 0 0 1 ${f(w - r)} ${f(h)}H${f(r)}A${f(r)} ${f(r)} 0 0 1 0 ${f(h - r)}V${f(r)}A${f(r)} ${f(r)} 0 0 1 ${f(r)} 0Z`;
    }
    case 'ellipse': return ellipse(w, h);
    case 'triangle': return poly([[w / 2, 0], [w, h], [0, h]]);
    case 'rtTriangle': return poly([[0, 0], [w, h], [0, h]]);
    case 'diamond': return poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
    case 'parallelogram': { const o = m * 0.25; return poly([[o, 0], [w, 0], [w - o, h], [0, h]]); }
    case 'trapezoid': { const o = m * 0.25; return poly([[o, 0], [w - o, 0], [w, h], [0, h]]); }
    case 'pentagon': return regular(w, h, 5);
    case 'hexagon': { const o = m * 0.25; return poly([[o, 0], [w - o, 0], [w, h / 2], [w - o, h], [o, h], [0, h / 2]]); }
    case 'octagon': { const c = m * 0.29; return poly([[c, 0], [w - c, 0], [w, c], [w, h - c], [w - c, h], [c, h], [0, h - c], [0, c]]); }
    case 'plus': {
      // Equal-thickness arms centred in the box.
      const a = m * 0.33;
      const x0 = (w - a) / 2;
      const y0 = (h - a) / 2;
      return poly([[x0, 0], [x0 + a, 0], [x0 + a, y0], [w, y0], [w, y0 + a], [x0 + a, y0 + a], [x0 + a, h], [x0, h], [x0, y0 + a], [0, y0 + a], [0, y0], [x0, y0]]);
    }
    case 'heart':
      return `M${f(w / 2)} ${f(h * 0.28)}C${f(w * 0.5)} ${f(h * 0.02)} 0 ${f(-h * 0.02)} 0 ${f(h * 0.32)}C0 ${f(h * 0.58)} ${f(w * 0.32)} ${f(h * 0.78)} ${f(w / 2)} ${f(h)}C${f(w * 0.68)} ${f(h * 0.78)} ${f(w)} ${f(h * 0.58)} ${f(w)} ${f(h * 0.32)}C${f(w)} ${f(-h * 0.02)} ${f(w * 0.5)} ${f(h * 0.02)} ${f(w / 2)} ${f(h * 0.28)}Z`;
    case 'star4': return star(w, h, 4, 0.32);
    case 'star5': return star(w, h, 5, 0.4);
    case 'wave':
      return `M0 ${f(h * 0.3)}C${f(w * 0.3)} ${f(-h * 0.15)} ${f(w * 0.6)} ${f(h * 0.7)} ${f(w)} ${f(h * 0.15)}V${f(h)}H0Z`;
    case 'rightArrow': { const hd = Math.min(w * 0.5, h * 0.6); return poly([[0, h * 0.25], [w - hd, h * 0.25], [w - hd, 0], [w, h / 2], [w - hd, h], [w - hd, h * 0.75], [0, h * 0.75]]); }
    case 'leftArrow': { const hd = Math.min(w * 0.5, h * 0.6); return poly([[w, h * 0.25], [hd, h * 0.25], [hd, 0], [0, h / 2], [hd, h], [hd, h * 0.75], [w, h * 0.75]]); }
    case 'upArrow': { const hd = Math.min(h * 0.5, w * 0.6); return poly([[w * 0.25, h], [w * 0.25, hd], [0, hd], [w / 2, 0], [w, hd], [w * 0.75, hd], [w * 0.75, h]]); }
    case 'downArrow': { const hd = Math.min(h * 0.5, w * 0.6); return poly([[w * 0.25, 0], [w * 0.75, 0], [w * 0.75, h - hd], [w, h - hd], [w / 2, h], [0, h - hd], [w * 0.25, h - hd]]); }
    case 'chevron': { const d = Math.min(w * 0.5, h * 0.5); return poly([[0, 0], [w - d, 0], [w, h / 2], [w - d, h], [0, h], [d, h / 2]]); }
    case 'homePlate': { const d = Math.min(w * 0.5, h * 0.5); return poly([[0, 0], [w - d, 0], [w, h / 2], [w - d, h], [0, h]]); }
    case 'wedgeRectCallout': return poly([[0, 0], [w, 0], [w, h * 0.78], [w * 0.42, h * 0.78], [w * 0.2, h], [w * 0.24, h * 0.78], [0, h * 0.78]]);
    case 'wedgeEllipseCallout': {
      const eh = h * 0.8;
      return `${ellipse(w, eh)}M${f(w * 0.22)} ${f(eh * 0.8)}L${f(w * 0.12)} ${f(h)}L${f(w * 0.38)} ${f(eh * 0.95)}Z`;
    }
    default: return `M0 0H${f(w)}V${f(h)}H0Z`;
  }
}

/**
 * Freeform shapes (from PowerPoint's custom geometry) keep their outline as
 * `paths: [{ cmds, fill, stroke }]`. Each command is ['M', x, y], ['L', x, y],
 * ['C', x1, y1, x2, y2, x, y], ['Q', x1, y1, x, y] or ['Z'], with coordinates
 * as fractions of the box, so the outline stretches with it. A path with
 * `fill: false` or `stroke: false` isn't filled or outlined.
 */
export const PATH_ARGS = { M: 2, L: 2, C: 6, Q: 4, Z: 0 };

/** SVG path data for freeform commands drawn in a w × h box. */
export function customPath(cmds, w, h) {
  return cmds.map(([op, ...n]) => op + n.map((v, i) => f(v * (i % 2 ? h : w))).join(' ')).join('');
}

/** Where text sits inside a shape (fractions of the box), so it stays inside the outline. */
export function textInset(kind) {
  switch (kind) {
    case 'ellipse': case 'wedgeEllipseCallout': return { x: 0.15, y: 0.15, b: kind === 'wedgeEllipseCallout' ? 0.3 : 0.15 };
    case 'triangle': return { x: 0.25, y: 0.45, b: 0.05 };
    case 'diamond': return { x: 0.25, y: 0.25, b: 0.25 };
    case 'star5': case 'star4': return { x: 0.3, y: 0.32, b: 0.25 };
    case 'wedgeRectCallout': return { x: 0.04, y: 0.04, b: 0.26 };
    case 'heart': return { x: 0.2, y: 0.2, b: 0.25 };
    default: return { x: 0, y: 0, b: 0 };
  }
}
