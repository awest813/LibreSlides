import JSZip from 'jszip';
import { uid, newDeck, SIZES } from '../model/deck.js';
import { SHAPES } from '../model/shapes.js';
import { isDarkColor } from '../model/themes.js';

/**
 * PowerPoint (.pptx) reader. Brings in what LibreSlides can edit: text boxes
 * and placeholders with their formatting, shapes, pictures, lines and
 * arrows, tables, backgrounds, speaker notes, transitions and skipped
 * slides. Positions inherited from layouts and masters are resolved. Charts,
 * SmartArt and video aren't supported and are left out.
 */
const EMU_PER_PT = 12700;
const SHAPE_IDS = new Set(SHAPES.map((s) => s.id));
const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', svg: 'image/svg+xml', bmp: 'image/bmp', webp: 'image/webp' };

const xml = (s) => new DOMParser().parseFromString(s, 'application/xml');
/** Child elements by local name (namespace prefixes vary between producers). */
const kids = (el, name) => (el ? [...el.children].filter((c) => c.localName === name) : []);
const kid = (el, name) => kids(el, name)[0] || null;
const path = (el, ...names) => names.reduce((e, n) => kid(e, n), el);
const desc = (el, name) => (el ? [...el.getElementsByTagNameNS('*', name)] : []);
const attr = (el, name) => el?.getAttribute(name) ?? null;
const rid = (el, name) => el?.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', name) || el?.getAttribute(`r:${name}`) || null;

function resolvePath(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

const relsPath = (p) => p.replace(/([^/]+)$/, '_rels/$1.rels');

// ------------------------------------------------------------------ colours

function hexToHsl(hex) {
  const n = parseInt(hex, 16);
  const r = ((n >> 16) & 255) / 255;
  const g = ((n >> 8) & 255) / 255;
  const b = (n & 255) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const hh = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [hh / 6, s, l];
}

function hslToHex([hh, s, l]) {
  const f = (n) => {
    const k = (n + hh * 12) % 12;
    const a = s * Math.min(l, 1 - l);
    return Math.round(255 * (l - a * Math.max(-1, Math.min(k - 3, 9 - k, 1))));
  };
  return [f(0), f(8), f(4)].map((v) => v.toString(16).padStart(2, '0')).join('');
}

const SCHEME_ALIAS = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' };

/** The colour of an element holding a colour choice (srgbClr, schemeClr…), as "#rrggbb". */
function colorOf(holder, scheme) {
  if (!holder) return null;
  const c = [...holder.children].find((e) => /Clr$/.test(e.localName));
  if (!c) return null;
  let hex = null;
  if (c.localName === 'srgbClr') hex = attr(c, 'val');
  else if (c.localName === 'schemeClr') {
    const v = attr(c, 'val');
    hex = scheme[SCHEME_ALIAS[v] || v] || null;
  } else if (c.localName === 'sysClr') hex = attr(c, 'lastClr') || (attr(c, 'val') === 'window' ? 'FFFFFF' : '000000');
  else if (c.localName === 'prstClr') hex = { black: '000000', white: 'FFFFFF', red: 'FF0000', blue: '0000FF', green: '00FF00' }[attr(c, 'val')] || '000000';
  if (!hex || !/^[0-9a-f]{6}$/i.test(hex)) return null;
  const mod = (name) => {
    const e = kid(c, name);
    return e ? +attr(e, 'val') / 100000 : null;
  };
  const lumMod = mod('lumMod');
  const lumOff = mod('lumOff');
  const tint = mod('tint');
  const shade = mod('shade');
  if (lumMod != null || lumOff != null || tint != null || shade != null) {
    const hsl = hexToHsl(hex);
    if (lumMod != null) hsl[2] *= lumMod;
    if (lumOff != null) hsl[2] += lumOff;
    if (tint != null) hsl[2] += (1 - hsl[2]) * (1 - tint);
    if (shade != null) hsl[2] *= shade;
    hsl[2] = Math.max(0, Math.min(1, hsl[2]));
    hex = hslToHex(hsl);
  }
  return `#${hex.toLowerCase()}`;
}

const alphaOf = (holder) => {
  const c = holder && [...holder.children].find((e) => /Clr$/.test(e.localName));
  const a = kid(c, 'alpha');
  return a ? +attr(a, 'val') / 100000 : 1;
};

// ------------------------------------------------------------------ reading the package

class Package {
  constructor(zip) {
    this.zip = zip;
    this.cache = new Map();
  }

  async text(p) {
    const f = this.zip.file(p);
    return f ? f.async('string') : null;
  }

  async doc(p) {
    if (!this.cache.has(p)) this.cache.set(p, this.text(p).then((t) => (t ? xml(t) : null)));
    return this.cache.get(p);
  }

  /** rId → { target path, type } for a part. */
  async rels(p) {
    const d = await this.doc(relsPath(p));
    const map = {};
    for (const r of desc(d, 'Relationship')) {
      const target = attr(r, 'Target');
      map[attr(r, 'Id')] = { path: attr(r, 'TargetMode') === 'External' ? target : resolvePath(p, target), type: (attr(r, 'Type') || '').split('/').pop(), external: attr(r, 'TargetMode') === 'External' };
    }
    return map;
  }

  async dataUrl(p) {
    const f = this.zip.file(p);
    if (!f) return null;
    const ext = (p.split('.').pop() || '').toLowerCase();
    const mime = MIME[ext];
    if (!mime) return null; // EMF/WMF and other formats browsers can't draw
    return `data:${mime};base64,${await f.async('base64')}`;
  }
}

// ------------------------------------------------------------------ text

const ALIGN = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' };
const esc = (s) => s.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));

/** Size, colour, font… of a run from a:rPr or a:defRPr. */
function runProps(rPr, ctx) {
  if (!rPr) return {};
  const p = {};
  if (attr(rPr, 'sz')) p.size = (+attr(rPr, 'sz') / 100) * ctx.k * ctx.fontScale;
  if (attr(rPr, 'b') != null) p.bold = attr(rPr, 'b') === '1' || attr(rPr, 'b') === 'true';
  if (attr(rPr, 'i') != null) p.italic = attr(rPr, 'i') === '1' || attr(rPr, 'i') === 'true';
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
  return p;
}

/** Convert a txBody to HTML; returns { html, base } where base is the box-level style. */
function textBody(txBody, ctx, { defaults = {}, bulletsByDefault = false } = {}) {
  const paras = kids(txBody, 'p');
  // Alignment shared by every paragraph belongs to the box, not to each paragraph.
  const aligns = paras.map((p) => ALIGN[attr(kid(p, 'pPr'), 'algn')] || null);
  const uniform = aligns.every((a) => a === aligns[0]);
  const out = [];
  let base = null;
  const listStack = []; // open list tags
  const closeLists = (depth) => {
    while (listStack.length > depth) out.push(`</li></${listStack.pop()}>`);
  };
  for (const p of paras) {
    const pPr = kid(p, 'pPr');
    const lvl = +(attr(pPr, 'lvl') || 0);
    const align = ALIGN[attr(pPr, 'algn')] || null;
    const runs = [];
    for (const r of p.children) {
      if (r.localName === 'r' || r.localName === 'fld') {
        const props = { ...defaults, ...runProps(kid(r, 'rPr'), ctx) };
        const t = kid(r, 't')?.textContent || '';
        if (t) runs.push({ text: t, props });
      } else if (r.localName === 'br') runs.push({ br: true });
    }
    const first = runs.find((r) => !r.br);
    if (!base && first) base = { ...first.props };
    const bullet = kid(pPr, 'buNone') ? null : kid(pPr, 'buAutoNum') ? 'ol' : kid(pPr, 'buChar') || kid(pPr, 'buBlip') || (bulletsByDefault && runs.length) ? 'ul' : null;
    const inner = runs.map((r) => (r.br ? '<br>' : spanFor(r.text, r.props, () => base))).join('') || '<br>';
    const style = align && !uniform ? ` style="text-align: ${align}"` : '';
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
  return { html: out.join('').replace(/<li>(<[uo]l>)/g, '<li><br>$1'), base: base || { ...defaults }, align: uniform ? aligns[0] : null };
}

/** A run as HTML, with inline styles only where it differs from the box's base style. */
function spanFor(text, props, getBase) {
  const base = getBase() || {};
  const css = [];
  if (props.size && Math.abs(props.size - (base.size || 0)) > 0.4) css.push(`font-size: ${Math.round(props.size * 10) / 10}px`);
  if (props.color && props.color !== base.color) css.push(`color: ${props.color}`);
  if (props.font && props.font !== base.font) css.push(`font-family: ${props.font}`);
  if (props.highlight) css.push(`background-color: ${props.highlight}`);
  if (props.bold !== undefined && !!props.bold !== !!base.bold) css.push(`font-weight: ${props.bold ? 700 : 400}`);
  if (props.italic !== undefined && !!props.italic !== !!base.italic) css.push(`font-style: ${props.italic ? 'italic' : 'normal'}`);
  const deco = [props.underline && 'underline', props.strike && 'line-through'].filter(Boolean);
  if (deco.length && (!!props.underline !== !!base.underline || !!props.strike !== !!base.strike)) css.push(`text-decoration: ${deco.join(' ')}`);
  const t = esc(text);
  return css.length ? `<span style="${css.join('; ')}">${t}</span>` : t;
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

/** Placeholder shapes of a layout or master, for looking up inherited positions and text sizes. */
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

function fillOf(spPr, scheme) {
  if (!spPr) return undefined;
  if (kid(spPr, 'noFill')) return 'none';
  const solid = kid(spPr, 'solidFill');
  if (solid) return colorOf(solid, scheme) || undefined;
  const grad = kid(spPr, 'gradFill');
  if (grad) return colorOf(desc(grad, 'gs')[0], scheme) || undefined;
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

// ------------------------------------------------------------------ slides

async function readSlide(pkg, slidePath, ctx, deck) {
  const doc = await pkg.doc(slidePath);
  const rels = await pkg.rels(slidePath);
  const layoutRel = Object.values(rels).find((r) => r.type === 'slideLayout');
  const layoutDoc = layoutRel ? await pkg.doc(layoutRel.path) : null;
  const layoutRels = layoutRel ? await pkg.rels(layoutRel.path) : {};
  const masterRel = Object.values(layoutRels).find((r) => r.type === 'slideMaster');
  const masterDoc = masterRel ? await pkg.doc(masterRel.path) : null;
  const layoutPh = layoutDoc ? placeholders(layoutDoc) : [];
  const masterPh = masterDoc ? placeholders(masterDoc) : [];
  const k = ctx.k;
  const slide = { id: uid(), layout: 'blank', background: null, elements: [], notes: '', transition: { type: 'none', duration: 0.6 }, hidden: attr(doc.documentElement, 'show') === '0', hideDecor: true };

  // Background: the slide's own, else the layout's, else the master's.
  for (const [d, r] of [[doc, rels], [layoutDoc, layoutRels], [masterDoc, masterRel ? await pkg.rels(masterRel.path) : {}]]) {
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

  const walk = async (container, group) => {
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
        await walk(node, g);
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
        return { ...b, x: x * k, y: y * k, w: w * k, h: hh * k };
      };

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
        let fill = fillOf(spPr, ctx.scheme);
        if (fill === undefined) fill = refs.hasFill ? refs.fill : 'none';
        const ln = kid(spPr, 'ln');
        let stroke = lineOf(ln, ctx.scheme, k);
        if (!stroke && refs.hasLine && refs.line && !kid(ln, 'noFill')) {
          stroke = { color: refs.line, width: attr(ln, 'w') ? Math.round((+attr(ln, 'w') / EMU_PER_PT) * k * 10) / 10 : 0.75, dash: 'solid' };
        }
        // Default text size: the placeholder's own list style, then the master's text styles.
        const bodyPr = kid(txBody, 'bodyPr');
        const fontScale = +(attr(kid(bodyPr, 'normAutofit'), 'fontScale') || 100000) / 100000;
        const lctx = { ...ctx, fontScale };
        const lvl1 = (d) => d && runProps(path(d, 'lstStyle', 'lvl1pPr', 'defRPr'), lctx);
        const isTitle = role === 'title';
        const masterStyle = runProps(path(ctx.txStyles, isTitle ? 'titleStyle' : role ? 'bodyStyle' : 'otherStyle', 'lvl1pPr', 'defRPr'), lctx);
        const defaults = { ...masterStyle, ...(masterMatch && lvl1(kid(masterMatch.sp, 'txBody'))), ...(inherited && lvl1(kid(inherited.sp, 'txBody'))), ...lvl1(txBody) };
        if (!defaults.size) defaults.size = (isTitle ? 44 : role ? 24 : 18) * k * fontScale;
        if (!defaults.font) defaults.font = isTitle ? ctx.fonts.major : ctx.fonts.minor;
        if (!defaults.color) defaults.color = refs.font || (fill && fill !== 'none' && prst !== 'rect' ? '#ffffff' : defaultText);
        const anchor = attr(bodyPr, 'anchor') || (inherited && attr(path(inherited.sp, 'txBody', 'bodyPr'), 'anchor')) || (isTitle ? 'ctr' : 't');
        const text = txBody ? textBody(txBody, lctx, { defaults, bulletsByDefault: role === 'body' }) : { html: '', base: defaults };
        const hasText = !!txBody?.textContent.trim();
        const base = text.base;
        const plainBox = prst === 'rect' && fill === 'none' && !stroke;
        const el = {
          id: uid(), type: plainBox ? 'text' : 'shape', x: box.x, y: box.y, w: box.w, h: box.h, rot: box.rot, flipH: box.flipH || undefined, flipV: box.flipV || undefined,
          text: hasText ? text.html : '',
          fill: fill || 'none', stroke: stroke ? { color: stroke.color, width: stroke.width, dash: stroke.dash } : null,
          style: {
            size: Math.round((base.size || defaults.size) * 10) / 10, font: base.font || defaults.font, color: base.color || defaults.color,
            bold: base.bold ?? defaults.bold ?? false, italic: base.italic || undefined, underline: base.underline || undefined,
            align: text.align || (isTitle && ph?.type === 'ctrTitle' ? 'center' : 'left'),
            valign: anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : 'top',
          },
        };
        if (!plainBox) el.shape = SHAPE_IDS.has(prst) ? prst : prst === 'flowChartProcess' ? 'rect' : prst === 'flowChartAlternateProcess' ? 'roundRect' : 'rect';
        if (role && hasText) el.ph = role;
        if (!hasText && plainBox) continue; // an empty placeholder: nothing to show
        slide.elements.push(el);
      } else if (name === 'pic') {
        const blip = desc(node, 'blip')[0];
        const target = rels[rid(blip, 'embed')];
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
        slide.elements.push({
          id: uid(), type: 'image', media: addMediaTo(deck, data), x: box.x, y: box.y, w: box.w, h: box.h, rot: box.rot,
          flipH: box.flipH || undefined, flipV: box.flipV || undefined, alt: descr || '', fit: 'fill',
          stroke: lineOf(kid(kid(node, 'spPr'), 'ln'), ctx.scheme, k) || { color: 'none', width: 2, dash: 'solid' },
        });
      } else if (name === 'cxnSp') {
        const line = lineElement(node, place(xfrmOf(kid(node, 'spPr'))));
        if (line) slide.elements.push(line);
      } else if (name === 'graphicFrame') {
        const tbl = desc(node, 'tbl')[0];
        const b = place(xfrmOf(node)); // a graphic frame's <p:xfrm> sits directly inside it
        if (!tbl || !b) continue;
        const cols = desc(kid(tbl, 'tblGrid'), 'gridCol').map((g) => +attr(g, 'w'));
        const total = cols.reduce((s, w) => s + w, 0) || 1;
        const firstRow = attr(kid(tbl, 'tblPr'), 'firstRow') === '1';
        const cells = kids(tbl, 'tr').map((tr) => kids(tr, 'tc').map((tc) => {
          const t = kid(tc, 'txBody') ? textBody(kid(tc, 'txBody'), { ...ctx, fontScale: 1 }, { defaults: { size: 14 * k } }) : { html: '' };
          const cell = { text: kid(tc, 'txBody')?.textContent.trim() ? t.html : '' };
          const fill = colorOf(kid(kid(tc, 'tcPr'), 'solidFill'), ctx.scheme);
          if (fill) cell.fill = fill;
          return cell;
        }));
        if (!cells.length) continue;
        const width = Math.max(...cells.map((r) => r.length));
        for (const row of cells) while (row.length < width) row.push({ text: '' });
        const firstRun = desc(tbl, 'rPr').find((r) => attr(r, 'sz'));
        slide.elements.push({
          id: uid(), type: 'table', x: b.x, y: b.y, w: b.w, h: b.h, rot: 0, header: firstRow && !cells[0].some((c) => c.fill),
          cells, colWidths: cols.length === width ? cols.map((w) => w / total) : Array.from({ length: width }, () => 1 / width),
          fill: `#${ctx.scheme.accent1 || '4472c4'}`, stroke: { color: '#94a3b8', width: 1, dash: 'solid' },
          style: { size: firstRun ? Math.round((+attr(firstRun, 'sz') / 100) * k) : Math.round(14 * k), color: defaultText },
        });
      }
    }
  };
  await walk(path(doc.documentElement, 'cSld', 'spTree'), null);

  // Transition (PowerPoint 2010+ effects sit inside mc:AlternateContent; the fallback is plain).
  const tr = desc(doc, 'transition')[0];
  if (tr) {
    const effect = [...tr.children].find((c) => c.localName !== 'sndAc' && c.localName !== 'extLst');
    const map = { fade: 'fade', push: 'push', cover: 'cover', pull: 'slide', wipe: 'slide', zoom: 'zoom', split: 'fade', dissolve: 'fade', flip: 'flip', cut: 'none' };
    const spd = attr(tr, 'spd');
    const dur = attr(tr, 'p14:dur') || attr(tr, 'dur');
    slide.transition = {
      type: effect ? map[effect.localName] || 'fade' : 'fade',
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

function addMediaTo(deck, data) {
  for (const [id, d] of Object.entries(deck.media)) if (d === data) return id;
  const id = `m${uid()}`;
  deck.media[id] = data;
  return id;
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

  const sz = desc(pres, 'sldSz')[0];
  const W = sz ? +attr(sz, 'cx') / EMU_PER_PT : 960;
  const H = sz ? +attr(sz, 'cy') / EMU_PER_PT : 540;
  const ratio = W / H;
  const size = Math.abs(ratio - 4 / 3) < 0.05 ? '4:3' : Math.abs(ratio - 1.6) < 0.05 ? '16:10' : '16:9';
  const k = SIZES[size].h / H;

  // Theme colours and fonts, from the first master's theme.
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
  const major = attr(path(desc(themeDoc, 'majorFont')[0], 'latin'), 'typeface');
  const minor = attr(path(desc(themeDoc, 'minorFont')[0], 'latin'), 'typeface');
  if (major) fonts.major = major;
  if (minor) fonts.minor = minor;
  const masterDoc = masterRel ? await pkg.doc(masterRel.path) : null;
  const txStyles = masterDoc ? desc(masterDoc, 'txStyles')[0] : null;

  const titleEl = (await pkg.doc('docProps/core.xml')) && desc(await pkg.doc('docProps/core.xml'), 'title')[0];
  const deck = newDeck({ title: titleEl?.textContent?.trim() || fallbackTitle, size, slides: [] });
  const ctx = { k, scheme, fonts, txStyles, fontScale: 1 };

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
  // Dark presentations get the dark theme, so new slides match.
  const firstBg = deck.slides[0].background;
  if (firstBg && isDarkColor(firstBg.color || firstBg.from)) deck.themeId = 'simple-dark';
  return deck;
}
