import JSZip from 'jszip';
import { uid, newDeck, SIZES } from '../model/deck.js';
import { SHAPES } from '../model/shapes.js';
import { isDarkColor } from '../model/themes.js';
import { readChart } from './pptx-chart.js';
import { readMedia } from './pptx-media.js';
import { readEmbeddedFonts } from './pptx-fonts.js';
import { EMU_PER_PT, kids, kid, path, desc, attr, rid, SCHEME_ALIAS, colorOf, alphaOf, blend, Package, addMediaTo } from './ooxml.js';

/**
 * PowerPoint (.pptx) reader. Brings in what LibreSlides can edit: text boxes
 * and placeholders with their formatting, shapes, pictures (with their crop),
 * lines and arrows, tables, backgrounds, speaker notes, transitions and
 * skipped slides. Positions, text styles and the graphics on layouts and
 * masters are resolved onto each slide. Charts, SmartArt and video aren't
 * supported and are left out.
 */
const SHAPE_IDS = new Set(SHAPES.map((s) => s.id));

// ------------------------------------------------------------------ text

const ALIGN = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' };
const esc = (s) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const isTrue = (v) => v === '1' || v === 'true';

/** Size, colour, font… of a run from a:rPr or a:defRPr. */
function runProps(rPr, ctx) {
  if (!rPr) return {};
  const p = {};
  if (attr(rPr, 'sz')) p.size = (+attr(rPr, 'sz') / 100) * ctx.k * ctx.fontScale;
  if (attr(rPr, 'b') != null) p.bold = isTrue(attr(rPr, 'b'));
  if (attr(rPr, 'i') != null) p.italic = isTrue(attr(rPr, 'i'));
  if (attr(rPr, 'u') != null) p.underline = attr(rPr, 'u') !== 'none';
  if (attr(rPr, 'strike') != null) p.strike = attr(rPr, 'strike') !== 'noStrike';
  const fill = kid(rPr, 'solidFill');
  if (fill) p.color = colorOf(fill, ctx.scheme);
  const hl = kid(rPr, 'highlight');
  if (hl) p.highlight = colorOf(hl, ctx.scheme);
  const latin = kid(rPr, 'latin');
  if (latin) {
    const face = attr(latin, 'typeface');
    if (face === '+mj-lt') p.font = ctx.fonts.major;
    else if (face === '+mn-lt') p.font = ctx.fonts.minor;
    else if (face && !face.startsWith('+')) p.font = face;
  }
  const link = rid(kid(rPr, 'hlinkClick'), 'id');
  const target = link && ctx.rels?.[link];
  if (target?.external && /^(https?:|mailto:)/i.test(target.path)) p.link = target.path;
  return p;
}

/**
 * Paragraph settings from an a:pPr or a list level (a:lvl1pPr…): alignment,
 * bullet, line and paragraph spacing, and default run formatting.
 */
function paraProps(pPr, ctx) {
  if (!pPr) return null;
  const p = { run: runProps(kid(pPr, 'defRPr'), ctx) };
  const align = ALIGN[attr(pPr, 'algn')];
  if (align) p.align = align;
  if (kid(pPr, 'buNone')) p.bullet = 'none';
  else if (kid(pPr, 'buAutoNum')) p.bullet = 'ol';
  else if (kid(pPr, 'buChar') || kid(pPr, 'buBlip')) p.bullet = 'ul';
  const ln = path(pPr, 'lnSpc', 'spcPct');
  if (ln) p.lineHeight = (+attr(ln, 'val') / 100000) * 1.2; // "single" spacing is about 1.2 lines
  for (const side of ['spcAft', 'spcBef']) {
    const pts = path(pPr, side, 'spcPts');
    const pct = path(pPr, side, 'spcPct');
    if (pts) p[side] = { pt: (+attr(pts, 'val') / 100) * ctx.k };
    else if (pct) p[side] = { em: (+attr(pct, 'val') / 100000) * 1.2 };
  }
  return p;
}

/**
 * Where a shape's text formatting comes from, lowest priority first: the
 * master's text styles, the master's and layout's placeholder, the shape's
 * own list style. Each entry gives the settings for a list level (0-8).
 */
const lstStyleLevels = (lstStyle, ctx) => (lvl) => paraProps(kid(lstStyle, `lvl${lvl + 1}pPr`), ctx);

function levelProps(chain, lvl) {
  const out = { run: {} };
  for (const get of chain) {
    const p = get(lvl);
    if (!p) continue;
    const { run, ...rest } = p;
    Object.assign(out.run, run);
    for (const [key, v] of Object.entries(rest)) if (v !== undefined) out[key] = v;
  }
  return out;
}

const emOf = (space, size) => (!space ? null : space.em ?? (size ? space.pt / size : 0));

/** Convert a txBody to HTML; returns { html, base, align, lineHeight, spacing } (base is the box-level run style). */
function textBody(txBody, ctx, { chain = [], defaults = {}, base: fixedBase = null } = {}) {
  const paras = kids(txBody, 'p').map((p) => {
    const pPr = kid(p, 'pPr');
    const lvl = Math.min(8, +(attr(pPr, 'lvl') || 0));
    const props = levelProps([...chain, () => paraProps(pPr, ctx)], lvl);
    const runs = [];
    for (const r of p.children) {
      if (r.localName === 'r' || r.localName === 'fld') {
        const t = kid(r, 't')?.textContent || '';
        if (t) runs.push({ text: t, props: { ...defaults, ...props.run, ...runProps(kid(r, 'rPr'), ctx) } });
      } else if (r.localName === 'br') runs.push({ br: true });
    }
    return { lvl, props, runs };
  });
  // Alignment shared by every paragraph belongs to the box, not to each paragraph.
  const aligns = paras.map((p) => p.props.align || null);
  const uniform = aligns.every((a) => a === aligns[0]);
  const first = paras.flatMap((p) => p.runs).find((r) => !r.br);
  const base = fixedBase || (first ? { ...first.props } : { ...defaults, ...(paras[0]?.props.run || {}) });
  const out = [];
  const listStack = []; // open list tags
  const closeLists = (depth) => {
    while (listStack.length > depth) out.push(`</li></${listStack.pop()}>`);
  };
  for (const { lvl, props, runs } of paras) {
    const bullet = runs.some((r) => !r.br) && (props.bullet === 'ul' || props.bullet === 'ol') ? props.bullet : null;
    const inner = runs.map((r) => (r.br ? '<br>' : spanFor(r.text, r.props, base))).join('') || '<br>';
    const style = props.align && !uniform ? ` style="text-align: ${props.align}"` : '';
    if (bullet) {
      const depth = lvl + 1;
      if (listStack.length > depth) closeLists(depth);
      if (listStack.length === depth) out.push(`</li><li${style}>${inner}`);
      else {
        while (listStack.length < depth) {
          out.push(`<${bullet}><li${listStack.length === depth - 1 ? style : ''}>`);
          listStack.push(bullet);
        }
        out.push(inner);
      }
    } else {
      closeLists(0);
      out.push(`<p${style}>${inner}</p>`);
    }
  }
  closeLists(0);
  const lead = paras[0]?.props || {};
  const size = base.size || defaults.size;
  return {
    html: out.join('').replace(/<li>(<[uo]l>)/g, '<li><br>$1'),
    base,
    align: uniform ? aligns[0] : null,
    lineHeight: lead.lineHeight,
    spacing: emOf(lead.spcAft, size) ?? emOf(lead.spcBef, size),
  };
}

/** A font name made safe to put in a style attribute. */
const cssFont = (name) => String(name).replace(/["'\\;<>{}()]/g, '').trim();

/** A run as HTML, with inline styles only where it differs from the box's base style. */
function spanFor(text, props, base) {
  const css = [];
  if (props.size && Math.abs(props.size - (base.size || 0)) > 0.4) css.push(`font-size: ${Math.round(props.size * 10) / 10}px`);
  if (props.color && props.color !== base.color) css.push(`color: ${props.color}`);
  if (props.font && props.font !== base.font && cssFont(props.font)) css.push(`font-family: '${cssFont(props.font)}'`);
  if (props.highlight) css.push(`background-color: ${props.highlight}`);
  // Formatting left unset is off, so compare as booleans.
  if (!!props.bold !== !!base.bold) css.push(`font-weight: ${props.bold ? 700 : 400}`);
  if (!!props.italic !== !!base.italic) css.push(`font-style: ${props.italic ? 'italic' : 'normal'}`);
  if (!!props.underline !== !!base.underline || !!props.strike !== !!base.strike) {
    css.push(`text-decoration: ${[props.underline && 'underline', props.strike && 'line-through'].filter(Boolean).join(' ') || 'none'}`);
  }
  let html = esc(text);
  if (css.length) html = `<span style="${css.join('; ')}">${html}</span>`;
  return props.link ? `<a href="${esc(props.link)}">${html}</a>` : html;
}

// ------------------------------------------------------------------ shapes

function xfrmOf(spPr) {
  const x = kid(spPr, 'xfrm');
  if (!x) return null;
  const off = kid(x, 'off');
  const ext = kid(x, 'ext');
  if (!off || !ext) return null;
  return {
    x: +attr(off, 'x') / EMU_PER_PT, y: +attr(off, 'y') / EMU_PER_PT,
    w: +attr(ext, 'cx') / EMU_PER_PT, h: +attr(ext, 'cy') / EMU_PER_PT,
    rot: (+(attr(x, 'rot') || 0)) / 60000, flipH: attr(x, 'flipH') === '1', flipV: attr(x, 'flipV') === '1',
  };
}

function phOf(sp) {
  const nv = [...sp.children].find((c) => /^nv/.test(c.localName));
  const ph = nv && desc(nv, 'ph')[0];
  if (!ph) return null;
  return { type: attr(ph, 'type') || 'body', idx: attr(ph, 'idx') };
}

const PH_ROLE = { title: 'title', ctrTitle: 'title', subTitle: 'subtitle', body: 'body', obj: 'body' };

/** Placeholder shapes of a layout or master, for looking up inherited positions and text styles. */
function placeholders(doc) {
  const list = [];
  for (const sp of desc(doc, 'sp')) {
    const ph = phOf(sp);
    if (ph) list.push({ ...ph, sp });
  }
  return list;
}

function findPh(list, ph) {
  if (!list) return null;
  const sameType = (a) => a.type === ph.type || (PH_ROLE[a.type] && PH_ROLE[a.type] === PH_ROLE[ph.type]);
  return (ph.idx != null && list.find((p) => p.idx === ph.idx && sameType(p)))
    || (ph.idx != null && list.find((p) => p.idx === ph.idx))
    || list.find((p) => p.type === ph.type)
    || list.find(sameType)
    || null;
}

function lineOf(ln, scheme, k) {
  if (!ln || kid(ln, 'noFill')) return null;
  const fill = kid(ln, 'solidFill') || kid(ln, 'gradFill');
  // An outline with no fill of its own takes the shape style's line (or has none).
  if (!fill) return null;
  const color = colorOf(fill.localName === 'gradFill' ? desc(fill, 'gs')[0] : fill, scheme);
  const w = attr(ln, 'w') ? (+attr(ln, 'w') / EMU_PER_PT) * k : 0.75;
  const dash = attr(kid(ln, 'prstDash'), 'val');
  return { color: color || '#000000', width: Math.round(w * 10) / 10, dash: !dash || dash === 'solid' ? 'solid' : /dot/i.test(dash) && !/dash/i.test(dash) ? 'dot' : /lgDash/.test(dash) ? 'longDash' : /dashDot/i.test(dash) ? 'dashDot' : 'dash' };
}

/** A shape's fill as { color, alpha } ('none' for no fill), or undefined when it takes the style's. */
function fillOf(spPr, scheme) {
  if (!spPr) return undefined;
  if (kid(spPr, 'noFill')) return { color: 'none', alpha: 1 };
  const solid = kid(spPr, 'solidFill');
  if (solid) {
    const color = colorOf(solid, scheme);
    return color ? { color, alpha: alphaOf(solid) } : undefined;
  }
  const grad = kid(spPr, 'gradFill');
  if (grad) {
    const stop = desc(grad, 'gs')[0];
    const color = colorOf(stop, scheme);
    return color ? { color, alpha: alphaOf(stop) } : undefined;
  }
  return undefined;
}

/** Theme default fill/line/font colour from <p:style> references. */
function styleRefs(sp, scheme) {
  const st = kid(sp, 'style');
  if (!st) return {};
  return {
    fill: colorOf(kid(st, 'fillRef'), scheme),
    line: colorOf(kid(st, 'lnRef'), scheme),
    font: colorOf(kid(st, 'fontRef'), scheme),
    hasLine: +(attr(kid(st, 'lnRef'), 'idx') || 0) > 0,
    hasFill: +(attr(kid(st, 'fillRef'), 'idx') || 0) > 0,
  };
}

/** The part of a picture shown (a:srcRect), as fractions cut from each side; null when uncropped. */
function cropOf(blipFill) {
  const r = kid(blipFill, 'srcRect');
  if (!r) return null;
  const side = (n) => Math.max(0, Math.min(0.99, (+attr(r, n) || 0) / 100000));
  const crop = { l: side('l'), t: side('t'), r: side('r'), b: side('b') };
  if (crop.l + crop.r >= 0.99 || crop.t + crop.b >= 0.99) return null;
  return Object.values(crop).some((v) => v > 0.0005) ? crop : null;
}

/** Closest LibreSlides shape for PowerPoint presets it doesn't draw itself. */
const PRESET_ALIAS = {
  flowChartProcess: 'rect', flowChartAlternateProcess: 'roundRect', flowChartTerminator: 'roundRect', snipRoundRect: 'roundRect', round2SameRect: 'roundRect',
  flowChartDecision: 'diamond', flowChartConnector: 'ellipse', flowChartInputOutput: 'parallelogram', flowChartPreparation: 'hexagon',
  notchedRightArrow: 'rightArrow', stripedRightArrow: 'rightArrow', star6: 'star5', star7: 'star5', star8: 'star5',
};

/**
 * A custom geometry (a:custGeom) as freeform paths (see model/shapes.js), or
 * null when it uses formulas this reader doesn't evaluate. `w` and `h` are
 * the shape's size in EMU, the units of paths that don't give their own.
 */
function customGeometry(custGeom, w, h) {
  const guides = {};
  for (const gd of desc(kid(custGeom, 'gdLst'), 'gd')) {
    const m = /^val\s+(-?[\d.]+)$/.exec(attr(gd, 'fmla') || '');
    if (m) guides[attr(gd, 'name')] = +m[1];
  }
  const paths = [];
  let count = 0;
  for (const p of kids(kid(custGeom, 'pathLst'), 'path')) {
    const pw = +attr(p, 'w') || w;
    const ph = +attr(p, 'h') || h;
    if (!pw || !ph) return null;
    const known = { l: 0, t: 0, r: pw, b: ph, w: pw, h: ph, hc: pw / 2, vc: ph / 2, wd2: pw / 2, hd2: ph / 2 };
    const num = (v) => (v == null ? NaN : /^-?[\d.]+$/.test(v) ? +v : known[v] ?? guides[v] ?? NaN);
    const pts = (el) => kids(el, 'pt').flatMap((pt) => [num(attr(pt, 'x')), num(attr(pt, 'y'))]);
    const cmds = [];
    let cur = [0, 0];
    let start = [0, 0];
    for (const c of p.children) {
      const name = c.localName;
      if (name === 'moveTo' || name === 'lnTo') {
        cur = pts(c);
        if (name === 'moveTo') start = cur;
        cmds.push([name === 'moveTo' ? 'M' : 'L', ...cur]);
      } else if (name === 'cubicBezTo' || name === 'quadBezTo') {
        const n = pts(c);
        cmds.push([name === 'cubicBezTo' ? 'C' : 'Q', ...n]);
        cur = n.slice(-2);
      } else if (name === 'arcTo') {
        const arc = arcToCurves(cur, num(attr(c, 'wR')), num(attr(c, 'hR')), num(attr(c, 'stAng')) / 60000, num(attr(c, 'swAng')) / 60000);
        cmds.push(...arc.cmds);
        cur = arc.end;
      } else if (name === 'close') {
        cmds.push(['Z']);
        cur = start;
      }
    }
    const flat = cmds.flat().filter((v) => typeof v === 'number');
    if (!cmds.length || flat.some((v) => !Number.isFinite(v))) return null;
    count += cmds.length;
    const path = { cmds: cmds.map(([op, ...n]) => [op, ...n.map((v, i) => Math.round((v / (i % 2 ? ph : pw)) * 10000) / 10000 || 0)]) };
    if (attr(p, 'fill') === 'none') path.fill = false;
    if (attr(p, 'stroke') === '0' || attr(p, 'stroke') === 'false') path.stroke = false;
    paths.push(path);
  }
  return paths.length && count <= 20000 ? paths : null;
}

/**
 * An a:arcTo from point `from` as cubic Béziers. The angles (degrees) are where
 * the arc starts and how far it turns, measured as seen on the ellipse, which
 * is what PowerPoint means by them.
 */
function arcToCurves(from, wR, hR, stAng, swAng) {
  if (![wR, hR, stAng, swAng].every(Number.isFinite) || !wR || !hR || !swAng) return { cmds: [], end: from };
  const rad = (d) => (d * Math.PI) / 180;
  // The ellipse's own (parametric) angle for an angle seen from its centre.
  const param = (a) => Math.atan2(wR * Math.sin(rad(a)), hR * Math.cos(rad(a)));
  const t0 = param(stAng);
  let sweep = param(stAng + swAng) - t0;
  const full = Math.abs(swAng) >= 360;
  if (full) sweep = Math.sign(swAng) * 2 * Math.PI;
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

// ------------------------------------------------------------------ slides

async function readSlide(pkg, slidePath, ctx, deck) {
  const doc = await pkg.doc(slidePath);
  const rels = await pkg.rels(slidePath);
  const layoutRel = Object.values(rels).find((r) => r.type === 'slideLayout');
  const layoutDoc = layoutRel ? await pkg.doc(layoutRel.path) : null;
  const layoutRels = layoutRel ? await pkg.rels(layoutRel.path) : {};
  const masterRel = Object.values(layoutRels).find((r) => r.type === 'slideMaster');
  const masterDoc = masterRel ? await pkg.doc(masterRel.path) : null;
  const masterRels = masterRel ? await pkg.rels(masterRel.path) : {};
  const layoutPh = layoutDoc ? placeholders(layoutDoc) : [];
  const masterPh = masterDoc ? placeholders(masterDoc) : [];
  const { k } = ctx;
  const slide = { id: uid(), layout: 'blank', background: null, elements: [], notes: '', transition: { type: 'none', duration: 0.6 }, hidden: attr(doc.documentElement, 'show') === '0', hideDecor: true };

  // Background: the slide's own, else the layout's, else the master's.
  for (const [d, r] of [[doc, rels], [layoutDoc, layoutRels], [masterDoc, masterRels]]) {
    const bg = d && path(d.documentElement, 'cSld', 'bg');
    if (!bg) continue;
    const bgPr = kid(bg, 'bgPr');
    const bgRef = kid(bg, 'bgRef');
    if (bgPr && kid(bgPr, 'blipFill')) {
      const blip = desc(bgPr, 'blip')[0];
      const target = r[rid(blip, 'embed')];
      const data = target && (await pkg.dataUrl(target.path));
      if (data) {
        slide.background = { type: 'image', media: addMediaTo(deck, data), color: '#000000' };
        break;
      }
    }
    const grad = bgPr && kid(bgPr, 'gradFill');
    if (grad) {
      const stops = desc(grad, 'gs');
      const lin = kid(grad, 'lin');
      const ang = lin ? (+attr(lin, 'ang') / 60000 + 90) % 360 : 180;
      slide.background = { type: 'gradient', from: colorOf(stops[0], ctx.scheme) || '#ffffff', to: colorOf(stops[stops.length - 1], ctx.scheme) || '#ffffff', angle: Math.round(ang) };
      break;
    }
    const color = colorOf(bgPr && kid(bgPr, 'solidFill'), ctx.scheme) || colorOf(bgRef, ctx.scheme);
    if (color) {
      slide.background = { type: 'solid', color };
      break;
    }
  }
  const bgColor = slide.background?.type === 'solid' ? slide.background.color : slide.background?.from || '#ffffff';
  const darkBg = isDarkColor(bgColor);
  const defaultText = darkBg ? (ctx.scheme.lt1 ? `#${ctx.scheme.lt1}` : '#ffffff') : (ctx.scheme.dk1 ? `#${ctx.scheme.dk1}` : '#000000');
  const masterStyle = (name, c) => lstStyleLevels(kid(ctx.txStyles, name), c);

  /** A straight line or connector, from its box and flips. */
  const lineElement = (node, b) => {
    if (!b) return null;
    const ln = kid(kid(node, 'spPr'), 'ln');
    const refs = styleRefs(node, ctx.scheme);
    const stroke = lineOf(ln, ctx.scheme, k) || { color: refs.line || defaultText, width: attr(ln, 'w') ? (+attr(ln, 'w') / EMU_PER_PT) * k : 1, dash: 'solid' };
    let [x1, x2] = [b.x, b.x + b.w];
    let [y1, y2] = [b.y, b.y + b.h];
    if (b.flipH) [x1, x2] = [x2, x1];
    if (b.flipV) [y1, y2] = [y2, y1];
    if (b.rot) {
      const cx = b.x + b.w / 2;
      const cy = b.y + b.h / 2;
      const r = (b.rot * Math.PI) / 180;
      const turn = (x, y) => [cx + (x - cx) * Math.cos(r) - (y - cy) * Math.sin(r), cy + (x - cx) * Math.sin(r) + (y - cy) * Math.cos(r)];
      [x1, y1] = turn(x1, y1);
      [x2, y2] = turn(x2, y2);
    }
    const head = attr(kid(ln, 'headEnd'), 'type');
    const tail = attr(kid(ln, 'tailEnd'), 'type');
    return {
      id: uid(), type: 'line', x1, y1, x2, y2, rot: 0, stroke: { color: stroke.color, width: Math.round(stroke.width * 10) / 10, dash: stroke.dash },
      arrowStart: !!head && head !== 'none', arrowEnd: !!tail && tail !== 'none',
    };
  };

  /**
   * Add the shapes of a shape tree. `part` is the slide, or the layout or
   * master whose graphics show on it (their placeholders are only templates
   * for the slide's own, so they're skipped).
   */
  const walk = async (container, group, part) => {
    const tctx = { ...ctx, rels: part.rels };
    for (const node of container.children) {
      const name = node.localName;
      if (name === 'grpSp') {
        const gx = xfrmOf(kid(node, 'grpSpPr'));
        const ch = kid(kid(node, 'grpSpPr'), 'xfrm');
        const chOff = kid(ch, 'chOff');
        const chExt = kid(ch, 'chExt');
        let g = group;
        if (gx && chOff && chExt) {
          const cw = +attr(chExt, 'cx') / EMU_PER_PT || 1;
          const chh = +attr(chExt, 'cy') / EMU_PER_PT || 1;
          const inner = { sx: gx.w / cw, sy: gx.h / chh, ox: gx.x - (+attr(chOff, 'x') / EMU_PER_PT) * (gx.w / cw), oy: gx.y - (+attr(chOff, 'y') / EMU_PER_PT) * (gx.h / chh) };
          g = group ? { sx: group.sx * inner.sx, sy: group.sy * inner.sy, ox: group.ox + inner.ox * group.sx, oy: group.oy + inner.oy * group.sy } : inner;
        }
        await walk(node, g, part);
        continue;
      }
      const place = (b) => {
        if (!b) return null;
        let { x, y, w, h: hh } = b;
        if (group) {
          x = group.ox + x * group.sx;
          y = group.oy + y * group.sy;
          w *= group.sx;
          hh *= group.sy;
        }
        return { ...b, x: x * k + ctx.ox, y: y * k + ctx.oy, w: w * k, h: hh * k };
      };
      if (!part.isSlide && (name === 'sp' || name === 'pic') && phOf(node)) continue;

      if (name === 'sp' && /^(line|straightConnector1)$/.test(attr(desc(node, 'prstGeom')[0], 'prst') || '')) {
        const line = lineElement(node, place(xfrmOf(kid(node, 'spPr'))));
        if (line) slide.elements.push(line);
      } else if (name === 'sp') {
        const spPr = kid(node, 'spPr');
        const ph = phOf(node);
        let box = xfrmOf(spPr);
        const inherited = ph ? (findPh(layoutPh, ph) || findPh(masterPh, ph)) : null;
        const masterMatch = ph ? findPh(masterPh, ph) : null;
        if (!box && inherited) box = xfrmOf(kid(inherited.sp, 'spPr')) || (masterMatch && xfrmOf(kid(masterMatch.sp, 'spPr')));
        box = place(box);
        if (!box) continue;
        const txBody = kid(node, 'txBody');
        const role = ph ? PH_ROLE[ph.type] : null;
        if (ph && !role && !txBody?.textContent.trim()) continue; // empty date/footer/slide number
        const refs = styleRefs(node, ctx.scheme);
        const prst = attr(kid(spPr, 'prstGeom'), 'prst') || 'rect';
        const own = xfrmOf(spPr);
        const paths = kid(spPr, 'custGeom') && own ? customGeometry(kid(spPr, 'custGeom'), own.w * EMU_PER_PT, own.h * EMU_PER_PT) : null;
        const ownFill = fillOf(spPr, ctx.scheme);
        let fill = ownFill?.color;
        if (fill === undefined) fill = refs.hasFill ? refs.fill : 'none';
        const alpha = ownFill?.alpha ?? 1;
        const ln = kid(spPr, 'ln');
        let stroke = lineOf(ln, ctx.scheme, k);
        if (!stroke && refs.hasLine && refs.line && !kid(ln, 'noFill')) {
          stroke = { color: refs.line, width: attr(ln, 'w') ? Math.round((+attr(ln, 'w') / EMU_PER_PT) * k * 10) / 10 : 0.75, dash: 'solid' };
        }
        const bodyPr = kid(txBody, 'bodyPr');
        const fontScale = +(attr(kid(bodyPr, 'normAutofit'), 'fontScale') || 100000) / 100000;
        const lctx = { ...tctx, fontScale };
        const isTitle = role === 'title';
        // Text formatting, lowest priority first: the master's text styles (or
        // the presentation's defaults for ordinary text boxes), the master's
        // and the layout's placeholder, the shape style's font colour, the shape's own list style.
        const chain = ph
          ? [masterStyle(isTitle ? 'titleStyle' : role ? 'bodyStyle' : 'otherStyle', lctx)]
          : [ctx.defaultTextStyle ? lstStyleLevels(ctx.defaultTextStyle, lctx) : masterStyle('otherStyle', lctx)];
        if (masterMatch) chain.push(lstStyleLevels(path(masterMatch.sp, 'txBody', 'lstStyle'), lctx));
        if (inherited && inherited !== masterMatch) chain.push(lstStyleLevels(path(inherited.sp, 'txBody', 'lstStyle'), lctx));
        if (refs.font) chain.push(() => ({ run: { color: refs.font } }));
        chain.push(lstStyleLevels(kid(txBody, 'lstStyle'), lctx));
        const defaults = {
          size: (isTitle ? 44 : role ? 24 : 18) * k * fontScale,
          font: isTitle ? ctx.fonts.major : ctx.fonts.minor,
          color: fill && fill !== 'none' && prst !== 'rect' ? '#ffffff' : defaultText,
        };
        const anchor = attr(bodyPr, 'anchor') || (inherited && attr(path(inherited.sp, 'txBody', 'bodyPr'), 'anchor')) || (masterMatch && attr(path(masterMatch.sp, 'txBody', 'bodyPr'), 'anchor')) || (isTitle ? 'ctr' : 't');
        const text = txBody ? textBody(txBody, lctx, { chain, defaults }) : { html: '', base: defaults };
        const hasText = !!txBody?.textContent.trim();
        const { base } = text;
        // A see-through fill: shown through the element's opacity when there's no text to fade with it.
        const fadeAll = alpha < 1 && !hasText && fill !== 'none' && (!stroke || !stroke.width);
        if (alpha < 1 && !fadeAll && fill && fill !== 'none') fill = blend(fill, bgColor, alpha);
        const plainBox = !paths && prst === 'rect' && fill === 'none' && !stroke;
        const el = {
          id: uid(), type: plainBox ? 'text' : 'shape', x: box.x, y: box.y, w: box.w, h: box.h, rot: box.rot, flipH: box.flipH || undefined, flipV: box.flipV || undefined,
          text: hasText ? text.html : '',
          fill: fill || 'none', stroke: stroke ? { color: stroke.color, width: stroke.width, dash: stroke.dash } : null,
          style: {
            size: Math.round((base.size || defaults.size) * 10) / 10, font: base.font || defaults.font, color: base.color || defaults.color,
            bold: !!base.bold, italic: base.italic || undefined, underline: base.underline || undefined, strike: base.strike || undefined,
            align: text.align || 'left',
            valign: anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : 'top',
            lineHeight: Math.round((text.lineHeight || 1.2) * 100) / 100,
            spacing: Math.round((text.spacing || 0) * 100) / 100,
          },
        };
        if (fadeAll) el.opacity = Math.round(alpha * 100) / 100;
        if (paths) {
          el.shape = 'custom';
          el.paths = paths;
        } else if (!plainBox) el.shape = SHAPE_IDS.has(prst) ? prst : PRESET_ALIAS[prst] || 'rect';
        if (role && hasText) el.ph = role;
        if (!hasText && (plainBox || (paths && fill === 'none' && !stroke))) continue; // nothing to show
        slide.elements.push(el);
      } else if (name === 'pic') {
        const media = await readMedia({ pkg, node, rels: part.rels, box: place(xfrmOf(kid(node, 'spPr'))), ctx: tctx, deck });
        if (media) {
          slide.elements.push(media);
          continue;
        }
        const blipFill = kid(node, 'blipFill');
        const blip = desc(node, 'blip')[0];
        const target = part.rels[rid(blip, 'embed')];
        const data = target && !target.external ? await pkg.dataUrl(target.path) : null;
        const ph = phOf(node);
        let box = xfrmOf(kid(node, 'spPr'));
        if (!box && ph) {
          const inherited = findPh(layoutPh, ph) || findPh(masterPh, ph);
          box = inherited && xfrmOf(kid(inherited.sp, 'spPr'));
        }
        box = place(box);
        if (!data || !box) continue;
        const descr = attr(desc(node, 'cNvPr')[0], 'descr');
        const amt = attr(kid(blip, 'alphaModFix'), 'amt');
        const el = {
          id: uid(), type: 'image', media: addMediaTo(deck, data), x: box.x, y: box.y, w: box.w, h: box.h, rot: box.rot,
          flipH: box.flipH || undefined, flipV: box.flipV || undefined, alt: descr || '', fit: 'fill',
          stroke: lineOf(kid(kid(node, 'spPr'), 'ln'), ctx.scheme, k) || { color: 'none', width: 2, dash: 'solid' },
        };
        const crop = cropOf(blipFill);
        if (crop) el.crop = crop;
        if (amt != null && +amt < 100000) el.opacity = Math.max(0, Math.round(+amt / 1000) / 100);
        slide.elements.push(el);
      } else if (name === 'cxnSp') {
        const line = lineElement(node, place(xfrmOf(kid(node, 'spPr'))));
        if (line) slide.elements.push(line);
      } else if (name === 'graphicFrame') {
        const uri = attr(desc(node, 'graphicData')[0], 'uri') || '';
        if (/\/chart$/.test(uri)) {
          const chart = await readChart({ pkg, node, rels: part.rels, box: place(xfrmOf(node)), ctx: tctx, deck });
          if (chart) slide.elements.push(chart);
          continue;
        }
        const tbl = desc(node, 'tbl')[0];
        const b = place(xfrmOf(node)); // a graphic frame's <p:xfrm> sits directly inside it
        if (!tbl || !b) continue;
        const cols = desc(kid(tbl, 'tblGrid'), 'gridCol').map((g) => +attr(g, 'w'));
        const total = cols.reduce((s, w) => s + w, 0) || 1;
        const tblPr = kid(tbl, 'tblPr');
        // The table's own text style: the formatting of its last row (the header is often set apart).
        const rows = kids(tbl, 'tr');
        const sample = { ...runProps(desc(rows[0], 'rPr')[0], tctx), ...runProps(desc(rows[rows.length - 1], 'rPr')[0], tctx) };
        const tableStyle = { size: Math.round(sample.size || 18 * k), color: sample.color || defaultText, font: sample.font, bold: false };
        const cells = rows.map((tr) => kids(tr, 'tc').map((tc) => {
          const body = kid(tc, 'txBody');
          // Formatting is relative to the table's, so a header's white bold text is kept.
          const t = body ? textBody(body, { ...tctx, fontScale: 1 }, { defaults: tableStyle, base: tableStyle }) : { html: '' };
          const cell = { text: body?.textContent.trim() ? t.html : '' };
          const fill = colorOf(kid(kid(tc, 'tcPr'), 'solidFill'), ctx.scheme);
          if (fill) cell.fill = fill;
          return cell;
        }));
        if (!cells.length) continue;
        const width = Math.max(...cells.map((r) => r.length));
        for (const row of cells) while (row.length < width) row.push({ text: '' });
        slide.elements.push({
          id: uid(), type: 'table', x: b.x, y: b.y, w: b.w, h: b.h, rot: 0,
          header: attr(tblPr, 'firstRow') === '1' && !cells[0].some((c) => c.fill),
          banded: attr(tblPr, 'bandRow') === '1' || undefined,
          cells, colWidths: cols.length === width ? cols.map((w) => w / total) : Array.from({ length: width }, () => 1 / width),
          fill: `#${ctx.scheme.accent1 || '4472c4'}`, stroke: { color: '#94a3b8', width: 1, dash: 'solid' },
          style: { size: tableStyle.size, color: tableStyle.color, font: tableStyle.font },
        });
      }
    }
  };

  // "Hide background graphics" on the slide hides the layout's and the master's; on the layout, the master's.
  const tree = (d) => path(d.documentElement, 'cSld', 'spTree');
  if (attr(doc.documentElement, 'showMasterSp') !== '0') {
    if (masterDoc && attr(layoutDoc?.documentElement, 'showMasterSp') !== '0') await walk(tree(masterDoc), null, { rels: masterRels });
    if (layoutDoc) await walk(tree(layoutDoc), null, { rels: layoutRels });
  }
  await walk(tree(doc), null, { rels, isSlide: true });

  // Transition (PowerPoint 2010+ effects sit inside mc:AlternateContent; the fallback is plain).
  const tr = desc(doc, 'transition')[0];
  if (tr) {
    const effect = [...tr.children].find((c) => c.localName !== 'sndAc' && c.localName !== 'extLst');
    const map = { fade: 'fade', push: 'push', cover: 'cover', pull: 'slide', wipe: 'slide', zoom: 'zoom', split: 'fade', dissolve: 'fade', flip: 'flip', cut: 'none' };
    let type = effect ? map[effect.localName] || 'fade' : 'fade';
    // "Cover from the right" is what LibreSlides calls "Slide from right".
    if (effect?.localName === 'cover' && attr(effect, 'dir') === 'l') type = 'slide';
    const spd = attr(tr, 'spd');
    const dur = attr(tr, 'p14:dur') || attr(tr, 'dur');
    slide.transition = {
      type,
      duration: dur ? Math.min(2, +dur / 1000) : spd === 'fast' ? 0.4 : spd === 'slow' ? 1 : 0.6,
    };
  }

  // Speaker notes.
  const notesRel = Object.values(rels).find((r) => r.type === 'notesSlide');
  if (notesRel) {
    const nd = await pkg.doc(notesRel.path);
    for (const sp of desc(nd, 'sp')) {
      const ph = phOf(sp);
      if (ph?.type !== 'body') continue;
      slide.notes = kids(kid(sp, 'txBody'), 'p').map((p) => desc(p, 't').map((t) => t.textContent).join('')).join('\n').trim();
    }
  }

  // Give the slide a layout name that matches what's on it, for "Apply layout".
  const roles = slide.elements.map((e) => e.ph).filter(Boolean);
  slide.layout = roles.includes('title') ? (roles.includes('body') ? 'titleContent' : roles.includes('subtitle') ? 'title' : 'titleOnly') : 'blank';
  return slide;
}

/** Read a .pptx (ArrayBuffer) into a deck. */
export async function importPptx(buffer, fallbackTitle = 'Presentation') {
  let zip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch {
    throw Object.assign(new Error('This file is damaged or isn’t a PowerPoint presentation.'), { code: 'bad-file' });
  }
  const pkg = new Package(zip);
  const presPath = 'ppt/presentation.xml';
  const pres = await pkg.doc(presPath);
  if (!pres) throw Object.assign(new Error('This isn’t a PowerPoint presentation (.pptx).'), { code: 'bad-file' });
  const presRels = await pkg.rels(presPath);

  // The closest LibreSlides size; anything else (A4, portrait…) is scaled to fit and centred.
  const sz = desc(pres, 'sldSz')[0];
  const W = (sz && +attr(sz, 'cx') / EMU_PER_PT) || 960;
  const H = (sz && +attr(sz, 'cy') / EMU_PER_PT) || 540;
  const size = Object.keys(SIZES).reduce((best, id) => (Math.abs(SIZES[id].w / SIZES[id].h - W / H) < Math.abs(SIZES[best].w / SIZES[best].h - W / H) ? id : best), '16:9');
  const k = Math.min(SIZES[size].w / W, SIZES[size].h / H);
  const ox = (SIZES[size].w - W * k) / 2;
  const oy = (SIZES[size].h - H * k) / 2;

  // Theme colours and fonts, from the first master's theme; the master's colour map says which is text and background.
  const scheme = {};
  const fonts = { major: 'Calibri', minor: 'Calibri' };
  const masterRel = Object.values(presRels).find((r) => r.type === 'slideMaster');
  const masterRels = masterRel ? await pkg.rels(masterRel.path) : {};
  const themeRel = Object.values(masterRels).find((r) => r.type === 'theme') || Object.values(presRels).find((r) => r.type === 'theme');
  const themeDoc = themeRel ? await pkg.doc(themeRel.path) : null;
  const clr = desc(themeDoc, 'clrScheme')[0];
  for (const c of clr ? [...clr.children] : []) {
    const v = [...c.children][0];
    scheme[c.localName] = v?.localName === 'srgbClr' ? attr(v, 'val') : v?.localName === 'sysClr' ? attr(v, 'lastClr') : null;
  }
  const masterDoc = masterRel ? await pkg.doc(masterRel.path) : null;
  const clrMap = desc(masterDoc, 'clrMap')[0];
  for (const [name, fallback] of Object.entries(SCHEME_ALIAS)) scheme[name] = scheme[attr(clrMap, name) || fallback] || scheme[fallback];
  const major = attr(path(desc(themeDoc, 'majorFont')[0], 'latin'), 'typeface');
  const minor = attr(path(desc(themeDoc, 'minorFont')[0], 'latin'), 'typeface');
  if (major) fonts.major = major;
  if (minor) fonts.minor = minor;
  const txStyles = masterDoc ? desc(masterDoc, 'txStyles')[0] : null;
  const defaultTextStyle = desc(pres, 'defaultTextStyle')[0] || null;

  const core = await pkg.doc('docProps/core.xml');
  const titleEl = core && desc(core, 'title')[0];
  const deck = newDeck({ title: titleEl?.textContent?.trim() || fallbackTitle, size, slides: [] });
  const ctx = { k, ox, oy, scheme, fonts, txStyles, defaultTextStyle, fontScale: 1 };
  await readEmbeddedFonts({ pkg, pres, presRels, deck });

  const ids = desc(desc(pres, 'sldIdLst')[0], 'sldId');
  for (const s of ids) {
    const rel = presRels[rid(s, 'id')];
    if (!rel) continue;
    try {
      deck.slides.push(await readSlide(pkg, rel.path, ctx, deck));
    } catch (err) {
      console.warn('Skipped a slide that could not be read', err);
    }
  }
  if (!deck.slides.length) throw Object.assign(new Error('No slides could be read from this file.'), { code: 'bad-file' });
  // Dark presentations get the dark theme, so new slides match (a picture background says nothing either way).
  const firstBg = deck.slides[0].background;
  if (firstBg && firstBg.type !== 'image' && isDarkColor(firstBg.color || firstBg.from)) deck.themeId = 'simple-dark';
  return deck;
}
