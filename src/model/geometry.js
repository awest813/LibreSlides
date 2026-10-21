/**
 * DrawingML shape geometry: PowerPoint's preset shapes and custom shapes are
 * outlines built from formulas ("guides") over the shape's size and its
 * adjust values. This works them out for one size, giving freeform paths
 * (see shapes.js) and the rectangle text goes in.
 *
 * A definition is { av: [[name, formula]], gd: [[name, formula]],
 * rect?: [l, t, r, b], paths: [{ w?, h?, fill?, stroke?, cmds }] }, where a
 * command is ['M', x, y], ['L', x, y], ['C', x1, y1, x2, y2, x, y],
 * ['Q', x1, y1, x, y], ['A', wR, hR, stAng, swAng] or ['Z'], and every
 * argument is a number or a guide name (as in the OOXML).
 */

const DEG = 60000; // angles are in 60000ths of a degree
const rad = (a) => (a / DEG) * (Math.PI / 180);

// The formula operators (ECMA-376 20.1.9.11).
const OPS = {
  '*/': (x, y, z) => (x * y) / z,
  '+-': (x, y, z) => x + y - z,
  '+/': (x, y, z) => (x + y) / z,
  '?:': (x, y, z) => (x > 0 ? y : z),
  abs: (x) => Math.abs(x),
  at2: (x, y) => (Math.atan2(y, x) * 180 * DEG) / Math.PI,
  cat2: (x, y, z) => x * Math.cos(Math.atan2(z, y)),
  cos: (x, y) => x * Math.cos(rad(y)),
  max: (x, y) => Math.max(x, y),
  min: (x, y) => Math.min(x, y),
  mod: (x, y, z) => Math.sqrt(x * x + y * y + z * z),
  pin: (x, y, z) => (y < x ? x : y > z ? z : y),
  sat2: (x, y, z) => x * Math.sin(Math.atan2(z, y)),
  sin: (x, y) => x * Math.sin(rad(y)),
  sqrt: (x) => Math.sqrt(x),
  tan: (x, y) => x * Math.tan(rad(y)),
  val: (x) => x,
};

/** The guides every shape has, for a w × h shape. */
function builtIns(w, h) {
  const ss = Math.min(w, h);
  const g = {
    l: 0, t: 0, r: w, b: h, w, h, hc: w / 2, vc: h / 2, ss, ls: Math.max(w, h),
    cd2: 10800000, cd4: 5400000, cd8: 2700000, '3cd4': 16200000, '3cd8': 8100000, '5cd8': 13500000, '7cd8': 18900000,
  };
  for (const n of [2, 3, 4, 5, 6, 8, 10, 12, 32]) g[`wd${n}`] = w / n;
  for (const n of [2, 3, 4, 5, 6, 8, 10, 12, 32]) g[`hd${n}`] = h / n;
  for (const n of [2, 4, 6, 8, 16, 32]) g[`ssd${n}`] = ss / n;
  return g;
}

/** Evaluate a definition's guides for a w × h shape with adjust values `adj` ({ adj1: 25000, … }). */
export function guides(def, w, h, adj = {}) {
  const g = builtIns(w, h);
  const value = (arg) => {
    if (typeof arg === 'number') return arg;
    if (arg in g) return g[arg];
    const n = Number(arg);
    if (Number.isFinite(n)) return n;
    throw new Error(`Unknown guide ${arg}`);
  };
  const run = (formula) => {
    const [op, ...args] = String(formula).trim().split(/\s+/);
    const fn = OPS[op];
    if (!fn) throw new Error(`Unknown formula ${op}`);
    return fn(...args.map(value));
  };
  for (const [name, formula] of def.av || []) g[name] = name in adj && Number.isFinite(adj[name]) ? adj[name] : run(formula);
  for (const [name, formula] of def.gd || []) g[name] = run(formula);
  return { g, value };
}

/**
 * The outline of a definition drawn w × h, as freeform paths with coordinates
 * as fractions of the box, and the text rectangle as fractions too. Throws on
 * formulas it can't work out.
 */
export function evaluateGeometry(def, w, h, adj = {}) {
  const { value } = guides(def, w, h, adj);
  const paths = [];
  for (const p of def.paths || []) {
    // A path may use its own coordinate space, stretched over the shape.
    const pw = value(p.w ?? w) || w;
    const ph = value(p.h ?? h) || h;
    const sx = w / pw;
    const sy = h / ph;
    const cmds = [];
    let cur = [0, 0];
    let start = [0, 0];
    for (const [op, ...args] of p.cmds) {
      const n = args.map(value);
      if (op === 'M' || op === 'L') {
        cur = n;
        if (op === 'M') start = n;
        cmds.push([op, ...n]);
      } else if (op === 'C' || op === 'Q') {
        cmds.push([op, ...n]);
        cur = n.slice(-2);
      } else if (op === 'A') {
        const arc = arcToCurves(cur, n[0], n[1], n[2] / DEG, n[3] / DEG);
        cmds.push(...arc.cmds);
        cur = arc.end;
      } else if (op === 'Z') {
        cmds.push(['Z']);
        cur = start;
      }
    }
    if (!cmds.length) continue;
    const frac = (v, i) => Math.round(((v * (i % 2 ? sy : sx)) / (i % 2 ? h : w)) * 10000) / 10000 || 0;
    const path = { cmds: cmds.map(([op, ...n]) => [op, ...n.map(frac)]) };
    if (path.cmds.some((c) => c.slice(1).some((v) => !Number.isFinite(v)))) throw new Error('Bad path');
    if (p.fill === 'none') path.fill = false;
    else if (p.fill && p.fill !== 'norm') path.shade = p.fill; // lighten, darken…: drawn lighter or darker
    if (p.stroke === false || p.stroke === '0' || p.stroke === 'false') path.stroke = false;
    paths.push(path);
  }
  let textRect = null;
  if (def.rect) {
    const [l, t, r, b] = def.rect.map(value);
    textRect = { l: l / w || 0, t: t / h || 0, r: 1 - r / w || 0, b: 1 - b / h || 0 };
    if (Object.values(textRect).some((v) => !Number.isFinite(v))) textRect = null;
  }
  return { paths, textRect };
}

/**
 * An arc from point `from` as cubic Béziers. The angles (degrees) are where
 * the arc starts and how far it turns, measured as seen on the ellipse, which
 * is what PowerPoint means by them.
 */
export function arcToCurves(from, wR, hR, stAng, swAng) {
  if (![wR, hR, stAng, swAng].every(Number.isFinite) || !wR || !hR || !swAng) return { cmds: [], end: from };
  const r = (d) => (d * Math.PI) / 180;
  // The ellipse's own (parametric) angle for an angle seen from its centre.
  const param = (a) => Math.atan2(wR * Math.sin(r(a)), hR * Math.cos(r(a)));
  const t0 = param(stAng);
  let sweep = param(stAng + swAng) - t0;
  if (Math.abs(swAng) >= 360) sweep = Math.sign(swAng) * 2 * Math.PI;
  else {
    while (swAng > 0 && sweep <= 0) sweep += 2 * Math.PI;
    while (swAng < 0 && sweep >= 0) sweep -= 2 * Math.PI;
  }
  const cx = from[0] - wR * Math.cos(t0);
  const cy = from[1] - hR * Math.sin(t0);
  const at = (t) => [cx + wR * Math.cos(t), cy + hR * Math.sin(t)];
  const steps = Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 2)));
  const d = sweep / steps;
  const k = (4 / 3) * Math.tan(d / 4);
  const cmds = [];
  let t = t0;
  for (let i = 0; i < steps; i++) {
    const [x1, y1] = at(t);
    const [x2, y2] = at(t + d);
    cmds.push(['C', x1 - k * wR * Math.sin(t), y1 + k * hR * Math.cos(t), x2 + k * wR * Math.sin(t + d), y2 - k * hR * Math.cos(t + d), x2, y2]);
    t += d;
  }
  return { cmds, end: at(t0 + sweep) };
}

/** A definition (as above) read from an a:custGeom or a preset definition's XML element. */
export function definitionFromXml(el) {
  const kid = (e, n) => (e ? [...e.children].find((c) => c.localName === n) || null : null);
  const kids = (e, n) => (e ? [...e.children].filter((c) => c.localName === n) : []);
  const gds = (list) => kids(kid(el, list), 'gd').map((gd) => [gd.getAttribute('name'), gd.getAttribute('fmla')]);
  const pt = (e) => {
    const p = kid(e, 'pt');
    return [p?.getAttribute('x'), p?.getAttribute('y')];
  };
  const rect = kid(el, 'rect');
  return {
    av: gds('avLst'),
    gd: gds('gdLst'),
    rect: rect ? ['l', 't', 'r', 'b'].map((s) => rect.getAttribute(s)) : undefined,
    paths: kids(kid(el, 'pathLst'), 'path').map((p) => {
      const out = { cmds: [] };
      for (const a of ['w', 'h', 'fill', 'stroke']) if (p.getAttribute(a) != null) out[a] = p.getAttribute(a);
      for (const c of p.children) {
        const n = c.localName;
        if (n === 'moveTo') out.cmds.push(['M', ...pt(c)]);
        else if (n === 'lnTo') out.cmds.push(['L', ...pt(c)]);
        else if (n === 'cubicBezTo' || n === 'quadBezTo') out.cmds.push([n === 'cubicBezTo' ? 'C' : 'Q', ...kids(c, 'pt').flatMap((p2) => [p2.getAttribute('x'), p2.getAttribute('y')])]);
        else if (n === 'arcTo') out.cmds.push(['A', ...['wR', 'hR', 'stAng', 'swAng'].map((a) => c.getAttribute(a))]);
        else if (n === 'close') out.cmds.push(['Z']);
      }
      return out;
    }),
  };
}
