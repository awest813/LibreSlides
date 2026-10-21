import { kids, kid, desc, attr, rid, addMediaTo } from './ooxml.js';
import { xmlText } from './pptx-comments.js';
import { normalizeLevels } from '../model/diagram-outline.js';
import { xmlDataUrl, xmlOfDataUrl, DIAGRAM_PARTS } from '../model/deck.js';
import { diagramChildren } from '../model/diagram-layout.js';
import { resolveColor } from '../model/themes.js';

/**
 * SmartArt in PowerPoint files to and from LibreSlides diagrams (see
 * model/deck.js).
 *
 * A SmartArt graphic frame points at four parts: its data (the outline as
 * points and connections), a layout definition, a quick style and a colour
 * definition, and PowerPoint adds a drawing of it (dsp:drawing) that other
 * apps show. LibreSlides reads the outline from the data, maps the layout,
 * colours and style to its own, and keeps the parts so the diagram goes back
 * as the same SmartArt. Edited, its data gets the new outline and the drawing
 * LibreSlides' layout of it. Microsoft's layout definitions aren't
 * LibreSlides' to hand out, so diagrams made in LibreSlides (or given a
 * layout, colours or style of its own) are saved as grouped shapes instead.
 */
const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const DSP = 'http://schemas.microsoft.com/office/drawing/2008/diagram';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

// ------------------------------------------------------------------ reading

/** The last part of a SmartArt definition's id ("urn:…/layout/process1" → "process1"). */
const idName = (uniqueId) => String(uniqueId || '').split('/').pop();

// PowerPoint's layouts by name, as the LibreSlides layout closest to each (first match wins).
const LAYOUT_KINDS = [
  [/chevron/i, 'chevron'],
  [/venn/i, 'venn'],
  [/matrix/i, 'matrix'],
  [/radial/i, 'radial'],
  [/hierarchy|orgChart/i, 'hierarchy'],
  [/pyramid/i, 'pyramid'],
  [/cycle/i, 'cycle'],
  [/process|arrow|equation|funnel/i, 'process'],
  [/^vList|^list1$|^vBullet|^stack/i, 'vlist'],
];

/** The LibreSlides layout for a SmartArt layout definition's id (a basic list for any it doesn't know). */
export const layoutKind = (uniqueId) => LAYOUT_KINDS.find(([re]) => re.test(idName(uniqueId)))?.[1] || 'list';

/** LibreSlides colours for a SmartArt colour definition's id ("accent2_1" → accent2, "colorful3" → colorful). */
export function colorsKind(uniqueId) {
  const name = idName(uniqueId);
  if (/^colorful/i.test(name)) return 'colorful';
  return `accent${/^accent([1-6])/.exec(name)?.[1] || 1}`;
}

/** LibreSlides style for a SmartArt quick style's id: the plain ones flat, those with effects raised. */
export const styleKind = (uniqueId) => (/^(moderate|intense|polished|inset|cartoon|powder|brick|sunset|birdsEye|3d)/i.test(idName(uniqueId)) ? 'raised' : 'flat');

/** A data point's text: its paragraphs' runs, the paragraphs joined by spaces. */
function pointText(pt) {
  return kids(kid(pt, 't'), 'p')
    .map((p) => desc(p, 't').filter((t) => /^(r|fld)$/.test(t.parentElement?.localName || '')).map((t) => t.textContent).join(''))
    .filter(Boolean).join(' ').replace(/[\u0000-\u001f\u007f]+/g, ' ').trim();
}

const isNode = (pt) => !attr(pt, 'type') || attr(pt, 'type') === 'node' || attr(pt, 'type') === 'asst';

/**
 * The outline in a SmartArt data model: from the document point down its
 * parent-of connections, children in their order, deeper ones a level down.
 * Points that aren't connected (older or hand-made files) come in order at
 * the top level.
 */
export function readOutline(data, limit = 500) {
  const root = data?.documentElement;
  const pts = kids(kid(root, 'ptLst'), 'pt');
  const byId = new Map(pts.map((p) => [attr(p, 'modelId'), p]));
  const under = new Map();
  for (const c of kids(kid(root, 'cxnLst'), 'cxn')) {
    if (attr(c, 'type') && attr(c, 'type') !== 'parOf') continue;
    const src = attr(c, 'srcId');
    if (!under.has(src)) under.set(src, []);
    under.get(src).push({ id: attr(c, 'destId'), ord: +attr(c, 'srcOrd') || 0 });
  }
  const nodes = [];
  const seen = new Set();
  const visit = (id, level) => {
    for (const { id: child } of (under.get(id) || []).sort((a, b) => a.ord - b.ord)) {
      const pt = byId.get(child);
      if (!pt || seen.has(child) || nodes.length >= limit) continue;
      seen.add(child);
      if (!isNode(pt)) continue;
      nodes.push({ id: child, text: pointText(pt), level });
      visit(child, level + 1);
    }
  };
  const doc = pts.find((p) => attr(p, 'type') === 'doc');
  if (doc) visit(attr(doc, 'modelId'), 0);
  if (!nodes.length) for (const pt of pts.filter(isNode).slice(0, limit)) nodes.push({ id: attr(pt, 'modelId') || String(nodes.length), text: pointText(pt), level: 0 });
  return normalizeLevels(nodes);
}

/**
 * A SmartArt graphic frame's diagram (see model/deck.js), without
 * PowerPoint's drawing, which the slide reader brings in as shapes: null
 * when it has no data. `drawingPath` is the drawing part, if there is one.
 * The parts are kept (in deck.media) when they're all there, so they can be
 * written back as they are, with the pictures they refer to:
 * source.media = { data?: [{ id (relationship), media }], drawing?: […] }.
 */
export async function readDiagram({ pkg, node, rels, scheme, font, deck }) {
  const ids = desc(node, 'relIds')[0];
  const partOf = (name) => {
    const r = rels?.[rid(ids, name)];
    return r && !r.external ? r.path : null;
  };
  const dataPath = partOf('dm');
  const data = dataPath && (await pkg.doc(dataPath));
  if (!data?.documentElement || data.documentElement.localName !== 'dataModel') return null;
  // The data names its drawing; older files only have one drawing per slide to go on.
  const drawingId = attr(desc(data, 'dataModelExt')[0], 'relId');
  const drawings = Object.values(rels || {}).filter((r) => r.type === 'diagramDrawing');
  const drawingPath = (rels?.[drawingId] || (drawings.length === 1 ? drawings[0] : null))?.path || null;
  const paths = { data: dataPath, layout: partOf('lo'), style: partOf('qs'), colors: partOf('cs') };
  const docs = {};
  for (const [key, p] of Object.entries(paths)) docs[key] = p ? await pkg.doc(p) : null;
  // The document point names the definitions too, for files without their parts.
  const prSet = kid(kids(kid(data.documentElement, 'ptLst'), 'pt').find((p) => attr(p, 'type') === 'doc'), 'prSet');
  const defId = (key, tag, fallback) => attr(desc(docs[key], tag)[0], 'uniqueId') || attr(prSet, fallback);
  const diagram = {
    layout: layoutKind(defId('layout', 'layoutDef', 'loTypeId')),
    nodes: readOutline(data),
    colors: colorsKind(defId('colors', 'colorsDef', 'csTypeId')),
    style: styleKind(defId('style', 'styleDef', 'qsTypeId')),
  };
  if (font) diagram.font = font;
  const accents = [1, 2, 3, 4, 5, 6].map((n) => scheme?.[`accent${n}`]);
  if (accents.every((c) => /^[0-9a-f]{6}$/i.test(c || ''))) diagram.palette = accents.map((c) => `#${c.toLowerCase()}`);
  const source = { look: { layout: diagram.layout, colors: diagram.colors, style: diagram.style } };
  let keep = !!deck;
  for (const [key, p] of Object.entries({ ...paths, drawing: drawingPath })) {
    const xml = p && (await pkg.text(p));
    if (!xml || !deck) continue;
    source[key] = addMediaTo(deck, xmlDataUrl(xml));
    // Pictures the part refers to travel with it; parts that refer to anything else can't.
    for (const [id, r] of Object.entries(await pkg.rels(p))) {
      const data = r.type === 'image' && !r.external && /^[\w-]{1,40}$/.test(id) ? await pkg.dataUrl(r.path) : null;
      if (!data || !deck) keep = false;
      else ((source.media ||= {})[key] ||= []).push({ id, media: addMediaTo(deck, data) });
    }
  }
  if (keep && source.data && source.layout && source.style && source.colors) diagram.source = source;
  return { diagram, drawingPath };
}

// ------------------------------------------------------------------ writing

/**
 * Whether a diagram goes into a .pptx as SmartArt: one from PowerPoint, still
 * with its own layout, colours and style, and (given the deck's `media`) with
 * its parts and pictures there.
 */
export function exportsAsSmartArt(el, media = null) {
  const d = el?.diagram;
  const src = d?.source;
  if (!src?.data || !src.layout || !src.style || !src.colors) return false;
  if (media) {
    const parts = DIAGRAM_PARTS.filter((key) => src[key]);
    if (!parts.every((key) => xmlOfDataUrl(media[src[key]]))) return false;
    if (!Object.values(src.media || {}).flat().every((r) => /^data:image\/[\w+.-]+;base64,/.test(media[r.media] || ''))) return false;
  }
  if (!d.edited) return true;
  const look = src.look || {};
  return look.layout === d.layout && look.colors === d.colors && look.style === d.style;
}

const emu = (pt) => Math.round(pt * 12700);
const guid = () => {
  const hex = globalThis.crypto?.randomUUID?.() || 'xxxxxxxx-xxxx-4xxx-8xxx-xxxxxxxxxxxx'.replace(/x/g, () => Math.floor(Math.random() * 16).toString(16));
  return `{${hex.toUpperCase()}}`;
};

/** A colour as PowerPoint's RRGGBB (theme colours resolved), or null for none. */
function hexOf(color, theme) {
  const c = resolveColor(color, theme, null);
  const m = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(c || '');
  if (!m) return null;
  return (m[1].length === 3 ? m[1].replace(/./g, '$&$&') : m[1]).toUpperCase();
}

const solid = (hex, alpha = 1) => `<a:solidFill><a:srgbClr val="${hex}">${alpha < 1 ? `<a:alpha val="${Math.round(alpha * 100000)}"/>` : ''}</a:srgbClr></a:solidFill>`;

/** A freeform outline (fractions of the box) as DrawingML custom geometry, with its text area. */
function custGeomXml(el) {
  const [W, H] = [emu(el.w), emu(el.h)];
  const pt = (x, y) => `<a:pt x="${Math.round(x * W)}" y="${Math.round(y * H)}"/>`;
  const cmds = (el.paths || []).flatMap((p) => p.cmds).map(([op, ...n]) => {
    if (op === 'M') return `<a:moveTo>${pt(n[0], n[1])}</a:moveTo>`;
    if (op === 'L') return `<a:lnTo>${pt(n[0], n[1])}</a:lnTo>`;
    if (op === 'C') return `<a:cubicBezTo>${pt(n[0], n[1])}${pt(n[2], n[3])}${pt(n[4], n[5])}</a:cubicBezTo>`;
    if (op === 'Q') return `<a:quadBezTo>${pt(n[0], n[1])}${pt(n[2], n[3])}</a:quadBezTo>`;
    return '<a:close/>';
  }).join('');
  const r = el.textRect;
  const rect = r ? `<a:rect l="${Math.round(r.l * W)}" t="${Math.round(r.t * H)}" r="${Math.round((1 - r.r) * W)}" b="${Math.round((1 - r.b) * H)}"/>` : '<a:rect l="l" t="t" r="r" b="b"/>';
  return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>${rect}<a:pathLst><a:path w="${W}" h="${H}">${cmds}</a:path></a:pathLst></a:custGeom>`;
}

/** The paragraphs of a diagram shape's text (see diagram-layout.js): [{ text, depth (-1: not a bullet), align }]. */
function paragraphsOf(html) {
  const body = new DOMParser().parseFromString(`<body>${html || ''}</body>`, 'text/html').body;
  const out = [];
  const list = (ul, depth) => {
    for (const li of ul.children) {
      if (li.tagName !== 'LI') continue;
      out.push({ text: [...li.childNodes].filter((n) => n.nodeName !== 'UL').map((n) => n.textContent).join(''), depth });
      for (const sub of li.children) if (sub.tagName === 'UL') list(sub, depth + 1);
    }
  };
  for (const n of body.children) {
    if (n.tagName === 'UL') list(n, 0);
    else out.push({ text: n.textContent, depth: -1, align: n.style?.textAlign || null });
  }
  return out;
}

const ALIGN = { left: 'l', center: 'ctr', right: 'r', justify: 'just' };

/** A shape's text as a dsp:txBody. */
function txBodyXml(el, theme) {
  const s = el.style || {};
  const size = s.size || 18;
  const color = hexOf(s.color || 'theme:text', theme) || '000000';
  const font = xmlText(s.font || theme.fonts.body);
  const rPr = `<a:rPr lang="en-US" sz="${Math.round(size * 100)}"${s.bold ? ' b="1"' : ''}>${solid(color)}<a:latin typeface="${font}"/></a:rPr>`;
  const indent = emu(size * 1.25);
  const paras = paragraphsOf(el.text).map((p) => {
    const pPr = p.depth >= 0
      ? `<a:pPr marL="${indent * (p.depth + 1)}" lvl="${Math.min(8, p.depth + 1)}" indent="-${indent}" algn="l"><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>`
      : `<a:pPr algn="${ALIGN[p.align || s.align] || 'ctr'}"/>`;
    return `<a:p>${pPr}${p.text ? `<a:r>${rPr}<a:t>${xmlText(p.text)}</a:t></a:r>` : `<a:endParaRPr lang="en-US" sz="${Math.round(size * 100)}"/>`}</a:p>`;
  }).join('');
  const anchor = { top: 't', bottom: 'b' }[s.valign] || 'ctr';
  return `<dsp:txBody><a:bodyPr spcFirstLastPara="0" vert="horz" wrap="square" numCol="1" spcCol="0" anchor="${anchor}" anchorCtr="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>${paras || '<a:p><a:endParaRPr lang="en-US"/></a:p>'}</dsp:txBody>`;
}

const xfrmXml = (x, y, w, h, { rot = 0, flipV = false } = {}) => `<a:xfrm${rot ? ` rot="${Math.round((((rot % 360) + 360) % 360) * 60000)}"` : ''}${flipV ? ' flipV="1"' : ''}><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>`;
const nvSpPr = '<dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>';
const PRESETS = new Set(['rect', 'roundRect', 'ellipse', 'chevron', 'homePlate', 'rightArrow', 'triangle', 'diamond']);

/** One of a diagram's elements (box coordinates) as a dsp:sp of a SmartArt drawing. */
function drawingShapeXml(el, theme) {
  if (el.type === 'line') {
    const [x, y] = [Math.min(el.x1, el.x2), Math.min(el.y1, el.y2)];
    const flipV = (el.x2 - el.x1) * (el.y2 - el.y1) < 0;
    const hex = hexOf(el.stroke?.color, theme) || '000000';
    return `<dsp:sp modelId="${guid()}">${nvSpPr}<dsp:spPr>${xfrmXml(x, y, Math.abs(el.x2 - el.x1), Math.abs(el.y2 - el.y1), { flipV })}<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="${emu(el.stroke?.width || 1)}">${solid(hex)}</a:ln></dsp:spPr></dsp:sp>`;
  }
  const geom = el.type === 'shape' && el.shape === 'custom' ? custGeomXml(el)
    : `<a:prstGeom prst="${el.type === 'shape' && PRESETS.has(el.shape) ? el.shape : 'rect'}"><a:avLst/></a:prstGeom>`;
  const fillHex = el.fill && el.fill !== 'none' ? hexOf(el.fill, theme) : null;
  const opacity = el.opacity ?? 1;
  const fill = fillHex ? solid(fillHex, opacity) : '<a:noFill/>';
  const strokeHex = el.stroke?.width && el.stroke.color !== 'none' ? hexOf(el.stroke.color, theme) : null;
  const ln = strokeHex ? `<a:ln w="${emu(el.stroke.width)}">${solid(strokeHex, opacity)}</a:ln>` : '<a:ln><a:noFill/></a:ln>';
  const effects = el.shadow ? `<a:effectLst><a:outerShdw blurRad="101600" dist="50800" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="30000"/></a:srgbClr></a:outerShdw></a:effectLst>` : '';
  const hasText = !!el.text && !!paragraphsOf(el.text).some((p) => p.text);
  const r = el.textRect;
  const txXfrm = hasText && r ? `<dsp:txXfrm><a:off x="${emu(el.x + r.l * el.w)}" y="${emu(el.y + r.t * el.h)}"/><a:ext cx="${emu(el.w * (1 - r.l - r.r))}" cy="${emu(el.h * (1 - r.t - r.b))}"/></dsp:txXfrm>` : '';
  return `<dsp:sp modelId="${guid()}">${nvSpPr}<dsp:spPr>${xfrmXml(el.x, el.y, el.w, el.h, { rot: el.rot })}${geom}${fill}${ln}${effects}</dsp:spPr>${hasText ? txBodyXml(el, theme) + txXfrm : ''}</dsp:sp>`;
}

/** A SmartArt drawing (dsp:drawing) of a diagram's elements, in its frame's coordinates. */
export function drawingXml(elements, theme) {
  const shapes = elements.map((el) => drawingShapeXml(el, theme)).join('');
  return `${XML_HEAD}<dsp:drawing xmlns:dgm="${DGM}" xmlns:dsp="${DSP}" xmlns:a="${A}"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>${shapes}</dsp:spTree></dsp:drawing>`;
}

/** A SmartArt drawing made for a frame of another size, its shapes' places and sizes stretched by sx, sy. */
export function stretchDrawingXml(xml, sx, sy) {
  if (Math.abs(sx - 1) < 0.001 && Math.abs(sy - 1) < 0.001) return xml;
  return xml
    .replace(/<a:(off|chOff) x="(-?\d+)" y="(-?\d+)"/g, (m, tag, x, y) => `<a:${tag} x="${Math.round(x * sx)}" y="${Math.round(y * sy)}"`)
    .replace(/<a:(ext|chExt) cx="(\d+)" cy="(\d+)"/g, (m, tag, cx, cy) => `<a:${tag} cx="${Math.round(cx * sx)}" cy="${Math.round(cy * sy)}"`);
}

/**
 * A SmartArt data model with a new outline: the document point and the
 * nodes still there are kept (with new text where it changed), new nodes
 * added, and the connections made afresh. The presentation points, which
 * PowerPoint keeps for its layout of the old outline, are left out for it
 * to make again. `drawingRelId` is the drawing's relationship from the slide.
 */
export function editedDataXml(dataXml, nodes, drawingRelId) {
  const doc = new DOMParser().parseFromString(dataXml, 'application/xml');
  const root = doc.documentElement;
  const prefix = root.prefix ? `${root.prefix}:` : '';
  const make = (name, attrs = {}) => {
    const e = doc.createElementNS(DGM, prefix + name);
    for (const [k, v] of Object.entries(attrs)) if (v != null) e.setAttribute(k, v);
    return e;
  };
  const drawing = (name) => doc.createElementNS(A, `a:${name}`);
  /** A dgm:t holding one paragraph of `text` (in the first run's formatting, if it had one). */
  const textOf = (text, rPr) => {
    const t = make('t');
    const p = drawing('p');
    t.append(drawing('bodyPr'), drawing('lstStyle'), p);
    if (text) {
      const r = drawing('r');
      const run = rPr ? rPr.cloneNode(true) : drawing('rPr');
      if (!rPr) run.setAttribute('lang', 'en-US');
      const at = drawing('t');
      at.textContent = text;
      r.append(run, at);
      p.append(r);
    } else {
      const end = drawing('endParaRPr');
      end.setAttribute('lang', 'en-US');
      p.append(end);
    }
    return t;
  };
  const point = (modelId, type, extra = {}) => {
    const pt = make('pt', { modelId, type, ...extra });
    pt.append(make('prSet', type ? {} : { phldrT: '[Text]' }), make('spPr'), textOf(''));
    return pt;
  };

  let ptLst = kid(root, 'ptLst');
  if (!ptLst) root.prepend((ptLst = make('ptLst')));
  let cxnLst = kid(root, 'cxnLst');
  if (!cxnLst) ptLst.after((cxnLst = make('cxnLst')));
  const pts = kids(ptLst, 'pt');
  const docPt = pts.find((p) => attr(p, 'type') === 'doc') || point(guid(), 'doc');
  const old = new Map(pts.filter(isNode).map((p) => [attr(p, 'modelId'), p]));
  ptLst.replaceChildren(docPt);
  cxnLst.replaceChildren();

  const stack = [attr(docPt, 'modelId')];
  const order = new Map(); // parent id → children so far
  const used = new Set();
  for (const node of normalizeLevels(nodes)) {
    while (stack.length > node.level + 1) stack.pop();
    const parent = stack[stack.length - 1];
    let pt = !used.has(node.id) && old.get(node.id);
    if (pt) {
      if (pointText(pt) !== node.text) {
        const rPr = desc(kid(pt, 't'), 'rPr')[0] || null;
        const t = textOf(node.text, rPr);
        const before = kid(pt, 't');
        if (before) before.replaceWith(t);
        else (kid(pt, 'extLst') ? kid(pt, 'extLst').before(t) : pt.append(t));
      }
    } else {
      pt = point(guid());
      pt.replaceChild(textOf(node.text), kid(pt, 't'));
    }
    const id = attr(pt, 'modelId');
    used.add(id);
    const cxnId = guid();
    const [parTrans, sibTrans] = [guid(), guid()];
    ptLst.append(pt, point(parTrans, 'parTrans', { cxnId }), point(sibTrans, 'sibTrans', { cxnId }));
    const n = order.get(parent) || 0;
    order.set(parent, n + 1);
    cxnLst.append(make('cxn', { modelId: cxnId, srcId: parent, destId: id, srcOrd: String(n), destOrd: '0', parTransId: parTrans, sibTransId: sibTrans }));
    stack.push(id);
  }

  // Where the drawing is: the data model's extension names its relationship from the slide.
  const ext = desc(root, 'dataModelExt')[0];
  if (ext && drawingRelId) ext.setAttribute('relId', drawingRelId);
  else if (ext?.parentElement?.localName === 'ext') ext.parentElement.remove();
  else if (drawingRelId) {
    let extLst = kid(root, 'extLst');
    if (!extLst) root.append((extLst = make('extLst')));
    const aExt = drawing('ext');
    aExt.setAttribute('uri', 'http://schemas.microsoft.com/office/drawing/2008/diagram');
    const dm = doc.createElementNS(DSP, 'dsp:dataModelExt');
    dm.setAttribute('relId', drawingRelId);
    dm.setAttribute('minVer', 'http://schemas.openxmlformats.org/drawingml/2006/diagram');
    aExt.append(dm);
    extLst.append(aExt);
  }
  return XML_HEAD + new XMLSerializer().serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, '');
}

const CONTENT_TYPE = {
  data: 'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml',
  layout: 'application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml',
  style: 'application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml',
  colors: 'application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml',
  drawing: 'application/vnd.ms-office.drawingml.diagramDrawing+xml',
};
const PART_NAME = { data: 'data', layout: 'layout', style: 'quickStyle', colors: 'colors', drawing: 'drawing' };
const REL_TYPE = {
  data: `${REL}/diagramData`, layout: `${REL}/diagramLayout`, style: `${REL}/diagramQuickStyle`, colors: `${REL}/diagramColors`,
  drawing: 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing',
};
const DATA_MODEL_REL = /(<(?:\w+:)?dataModelExt\b[^>]*?\brelId=")([^"]*)(")/;

/**
 * Put a diagram that goes as SmartArt (see exportsAsSmartArt) into an
 * exported package for slide `slideNo`: its parts, their relationships and
 * content types, and the pictures (from `media`, the deck's) they refer to.
 * Returns the p:graphicFrame that shows it, named by `cNvPr` (the p:cNvPr
 * of the shape it takes the place of).
 */
export async function writeDiagram(zip, slideNo, el, theme, cNvPr, media = {}) {
  const d = el.diagram;
  const src = Object.fromEntries(DIAGRAM_PARTS.map((key) => [key, xmlOfDataUrl(media[d.source[key]])]));
  let n = 1;
  while (zip.file(`ppt/diagrams/data${n}.xml`)) n++;
  const relsPath = `ppt/slides/_rels/slide${slideNo}.xml.rels`;
  let rels = (await zip.file(relsPath)?.async('string')) || `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  const taken = (id) => rels.includes(`Id="${id}"`);

  // The drawing: as it was (stretched if the diagram was), or LibreSlides' layout once edited.
  let drawing = null;
  if (!d.edited && src.drawing) drawing = d.drawn?.w && d.drawn?.h ? stretchDrawingXml(src.drawing, el.w / d.drawn.w, el.h / d.drawn.h) : src.drawing;
  else if (d.edited) drawing = drawingXml(diagramChildren(el), theme);
  const oldId = DATA_MODEL_REL.exec(src.data)?.[2];
  const ids = { data: `rIdLsDm${n}`, layout: `rIdLsLo${n}`, style: `rIdLsQs${n}`, colors: `rIdLsCs${n}`, drawing: drawing ? (!d.edited && oldId && !taken(oldId) ? oldId : `rIdLsDr${n}`) : null };
  let data = src.data;
  if (d.edited) data = editedDataXml(src.data, d.nodes, ids.drawing);
  else if (drawing && oldId !== ids.drawing) data = data.replace(DATA_MODEL_REL, (m, a, id, b) => a + ids.drawing + b);
  const parts = { data, layout: src.layout, style: src.style, colors: src.colors, drawing };
  const source = d.source;

  let types = await zip.file('[Content_Types].xml').async('string');
  let pictures = 0;
  for (const [key, xml] of Object.entries(parts)) {
    if (!xml) continue;
    const name = `${PART_NAME[key]}${n}.xml`;
    zip.file(`ppt/diagrams/${name}`, xml);
    rels = rels.replace('</Relationships>', `<Relationship Id="${ids[key]}" Type="${REL_TYPE[key]}" Target="../diagrams/${name}"/></Relationships>`);
    types = types.replace('</Types>', `<Override PartName="/ppt/diagrams/${name}" ContentType="${CONTENT_TYPE[key]}"/></Types>`);
    // The pictures the part refers to (not LibreSlides' own drawing, which has none).
    const refs = key === 'drawing' && d.edited ? [] : source.media?.[key] || [];
    const own = [];
    for (const { id, media: mediaId } of refs) {
      const m = /^data:image\/(png|jpeg|gif|bmp|webp|svg\+xml);base64,(.*)$/s.exec(media?.[mediaId] || '');
      if (!m) continue;
      const ext = { jpeg: 'jpg', 'svg+xml': 'svg' }[m[1]] || m[1];
      const file = `ls-diagram-${n}-${++pictures}.${ext}`;
      zip.file(`ppt/media/${file}`, m[2], { base64: true });
      own.push(`<Relationship Id="${id}" Type="${REL}/image" Target="../media/${file}"/>`);
      if (!new RegExp(`<Default Extension="${ext}"`, 'i').test(types)) types = types.replace('</Types>', `<Default Extension="${ext}" ContentType="image/${m[1]}"/></Types>`);
    }
    if (own.length) zip.file(`ppt/diagrams/_rels/${name}.rels`, `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${own.join('')}</Relationships>`);
  }
  zip.file(relsPath, rels);
  zip.file('[Content_Types].xml', types);

  const name = cNvPr || `<p:cNvPr id="${1000 + n}" name="Diagram ${n}"/>`;
  return `<p:graphicFrame><p:nvGraphicFramePr>${name}<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>`
    + `<p:xfrm><a:off x="${emu(el.x)}" y="${emu(el.y)}"/><a:ext cx="${emu(el.w)}" cy="${emu(el.h)}"/></p:xfrm>`
    + `<a:graphic><a:graphicData uri="${DGM}"><dgm:relIds xmlns:dgm="${DGM}" xmlns:r="${REL}" r:dm="${ids.data}" r:lo="${ids.layout}" r:qs="${ids.style}" r:cs="${ids.colors}"/></a:graphicData></a:graphic></p:graphicFrame>`;
}
