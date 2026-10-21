import { uid, initialsOf } from '../model/deck.js';
import { EMU_PER_PT, kids, kid, desc, attr } from './ooxml.js';

/**
 * Slide comments in .pptx files.
 *
 * PowerPoint has two kinds. The classic ones (ECMA-376) sit in
 * ppt/comments/commentN.xml, one p:cm per comment, with the authors in
 * ppt/commentAuthors.xml; their positions are in "master units" of 1/576 inch
 * (1/8 point: LibreOffice reads them as EMU × 15.87 and writes mm × 576 / 25.4).
 * They have no threads of their own, but PowerPoint 2013 and later mark a
 * reply with p15:threadingInfo/p15:parentCm (its thread's author and number).
 * Microsoft 365 writes threaded "modern" comments instead
 * (ppt/comments/modernComment_*.xml, authors in ppt/authors.xml): p188:cm
 * with a p188:replyLst, a status ("resolved"), text as DrawingML paragraphs,
 * and either a position in EMU or the object it's attached to.
 *
 * Both kinds are read. The classic kind is written, since PowerPoint,
 * LibreOffice and Google Slides all read it; it can't say a thread is
 * resolved, so resolved threads are written as open ones.
 */
const MODERN_NS = 'http://schemas.microsoft.com/office/powerpoint/2018/8/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const TYPES = 'application/vnd.openxmlformats-officedocument.presentationml';
const THREADING_EXT = '{C676402C-5697-4E1C-873F-D02D1690AC5C}';
const P15 = 'http://schemas.microsoft.com/office/powerpoint/2012/main';
const MASTER_UNITS_PER_PT = 8;

// ------------------------------------------------------------------ reading

/** Authors of both kinds by id (classic ones are numbers, modern ones GUIDs), read once per file. */
const authorLists = new WeakMap();
function authorsOf(pkg) {
  if (!authorLists.has(pkg)) {
    authorLists.set(pkg, (async () => {
      const map = new Map();
      for (const r of Object.values(await pkg.rels('ppt/presentation.xml'))) {
        if (r.type !== 'commentAuthors' && r.type !== 'authors') continue;
        for (const a of desc(await pkg.doc(r.path), r.type === 'authors' ? 'author' : 'cmAuthor')) {
          map.set(attr(a, 'id'), { name: attr(a, 'name') || '', initials: attr(a, 'initials') || '' });
        }
      }
      return map;
    })());
  }
  return authorLists.get(pkg);
}

/** An ISO date from a file's date, which may lack a time zone (`biasMinutes`: UTC minus that local time). */
function isoDate(value, biasMinutes = 0) {
  if (!value) return null;
  const zoned = /(z|[+-]\d\d:?\d\d)$/i.test(value);
  const t = Date.parse(zoned ? value : `${value}Z`);
  return Number.isFinite(t) ? new Date(t + (zoned ? 0 : biasMinutes * 60000)).toISOString() : null;
}

/** The author, date and text every comment and reply has. */
function entry(authors, authorId, date, text) {
  const a = authors.get(authorId);
  const out = { id: uid(), author: a?.name || 'Anonymous', text };
  if (a?.initials) out.initials = a.initials;
  if (date) out.date = date;
  return out;
}

/** A point in the file's units as slide points (scaled and centred like everything else); (0, 0) means the whole slide. */
function place(x, y, unitPt, ctx) {
  if (!Number.isFinite(x) || !Number.isFinite(y) || (x === 0 && y === 0)) return {};
  const { k = 1, ox = 0, oy = 0 } = ctx || {};
  const round = (v) => Math.round(v * 100) / 100;
  return { x: round(x * unitPt * k + ox), y: round(y * unitPt * k + oy) };
}

const pointOf = (pos, unitPt, ctx) => (pos ? place(+attr(pos, 'x'), +attr(pos, 'y'), unitPt, ctx) : {});

/** Classic comments: replies (marked with their thread's author and number) join that thread. */
function readClassic(doc, authors, ctx) {
  const items = kids(doc.documentElement, 'cm').map((cm) => {
    const threading = desc(cm, 'threadingInfo')[0];
    const parent = kid(threading, 'parentCm');
    return {
      key: `${attr(cm, 'authorId')}:${attr(cm, 'idx')}`,
      parentKey: parent ? `${attr(parent, 'authorId')}:${attr(parent, 'idx')}` : null,
      comment: { ...entry(authors, attr(cm, 'authorId'), isoDate(attr(cm, 'dt'), +attr(threading, 'timeZoneBias') || 0), kid(cm, 'text')?.textContent || ''), ...pointOf(kid(cm, 'pos'), 1 / MASTER_UNITS_PER_PT, ctx) },
    };
  });
  const byKey = new Map(items.map((i) => [i.key, i]));
  // The thread a comment belongs to: follow the parents up (a reply to a reply joins the same thread).
  const rootOf = (item) => {
    const seen = new Set();
    while (item.parentKey && byKey.has(item.parentKey) && !seen.has(item)) {
      seen.add(item);
      item = byKey.get(item.parentKey);
    }
    return item;
  };
  const out = [];
  for (const item of items) {
    const root = rootOf(item);
    if (root === item) {
      out.push(item.comment);
      continue;
    }
    const { x, y, ...reply } = item.comment;
    (root.comment.replies ||= []).push(reply);
  }
  return out;
}

/** Plain text of a modern comment's paragraphs. */
const bodyText = (txBody) => kids(txBody, 'p')
  .map((p) => [...p.children].map((r) => (r.localName === 'br' ? '\n' : desc(r, 't').map((t) => t.textContent).join(''))).join(''))
  .join('\n');

/** Where a modern comment attached to an object goes: that object's top-left corner (top-level objects with a position only). */
function objectOrigin(slideDoc, cm, ctx) {
  const list = [...cm.children].find((c) => c.localName === 'deMkLst');
  const mk = list && [...list.children].find((c) => c.localName !== 'docMk' && c.localName !== 'sldMk' && attr(c, 'id'));
  if (!mk || !slideDoc) return null;
  const nv = desc(slideDoc, 'cNvPr').find((c) => attr(c, 'id') === attr(mk, 'id'));
  const shape = nv?.parentElement?.parentElement;
  if (!shape || shape.parentElement?.localName !== 'spTree') return null;
  const off = desc(shape, 'off')[0];
  return off ? place(+attr(off, 'x'), +attr(off, 'y'), 1 / EMU_PER_PT, ctx) : null;
}

function readModern(doc, authors, slideDoc, ctx) {
  return kids(doc.documentElement, 'cm').map((cm) => {
    const c = entry(authors, attr(cm, 'authorId'), isoDate(attr(cm, 'created')), bodyText(kid(cm, 'txBody')));
    Object.assign(c, objectOrigin(slideDoc, cm, ctx) || pointOf(kid(cm, 'pos'), 1 / EMU_PER_PT, ctx));
    if (/^(resolved|closed)$/i.test(attr(cm, 'status') || '')) c.resolved = true;
    const replies = kids(kid(cm, 'replyLst'), 'reply').map((r) => entry(authors, attr(r, 'authorId'), isoDate(attr(r, 'created')), bodyText(kid(r, 'txBody'))));
    if (replies.length) c.replies = replies;
    return c;
  });
}

/**
 * A slide's comments (`rels`: the slide's relationships; `ctx`: the reader's
 * scale `k` and offsets `ox`, `oy`). Modern comments win when a slide has both
 * kinds, since PowerPoint keeps the classic ones only for older readers.
 */
export async function readComments(pkg, slidePath, rels, ctx) {
  try {
    const parts = Object.values(rels).filter((r) => r.type === 'comments' && !r.external);
    const docs = (await Promise.all(parts.map((r) => pkg.doc(r.path)))).filter((d) => d?.documentElement?.localName === 'cmLst');
    if (!docs.length) return [];
    const authors = await authorsOf(pkg);
    const modern = docs.filter((d) => d.documentElement.namespaceURI === MODERN_NS);
    if (modern.length) {
      const slideDoc = await pkg.doc(slidePath);
      return modern.flatMap((d) => readModern(d, authors, slideDoc, ctx));
    }
    return docs.flatMap((d) => readClassic(d, authors, ctx));
  } catch (err) {
    console.warn('Skipped comments that could not be read', err);
    return [];
  }
}

// ------------------------------------------------------------------ writing

/** Text for XML: escaped, without the control characters (or halves of emoji) XML can't hold. */
export const xmlText = (s) => String(s ?? '')
  .replace(/\r\n?/g, '\n')
  .replace(/[\ud800-\udbff][\udc00-\udfff]|[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff\ud800-\udfff]/g, (c) => (c.length === 2 ? c : ''))
  .replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c]));

/**
 * A classic comment's date: local time with no zone, as PowerPoint writes it,
 * and the minutes to add to get UTC (PowerPoint's timeZoneBias).
 */
function classicDate(iso) {
  const d = new Date(iso);
  if (!iso || Number.isNaN(d.getTime())) return null;
  const bias = d.getTimezoneOffset();
  return { dt: new Date(d.getTime() - bias * 60000).toISOString().slice(0, 23), bias };
}

/** One p:cm; a reply names its thread's first comment. */
function cmXml(authorId, idx, c, pos, parent) {
  const date = classicDate(c.date);
  const parentXml = parent ? `<p15:parentCm authorId="${parent.authorId}" idx="${parent.idx}"/>` : '';
  const threading = date || parent
    ? `<p:extLst><p:ext uri="${THREADING_EXT}"><p15:threadingInfo xmlns:p15="${P15}"${date ? ` timeZoneBias="${date.bias}"` : ''}>${parentXml}</p15:threadingInfo></p:ext></p:extLst>`
    : '';
  return `<p:cm authorId="${authorId}"${date ? ` dt="${date.dt}"` : ''} idx="${idx}"><p:pos x="${pos.x}" y="${pos.y}"/><p:text>${xmlText(c.text)}</p:text>${threading}</p:cm>`;
}

/** Add a relationship to a .rels part (creating it if needed). */
async function addRelationship(zip, relsPath, id, type, target) {
  const rel = `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`;
  const file = zip.file(relsPath);
  const xml = file ? await file.async('string') : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  zip.file(relsPath, xml.replace('</Relationships>', `${rel}</Relationships>`));
}

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const P_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

/**
 * Write the deck's comments into an exported package (slides named
 * ppt/slides/slideN.xml in deck order): a comments part per slide, the
 * authors, and their relationships and content types. A reply follows its
 * thread at the same place, by its own author.
 */
export async function writeComments(zip, deck) {
  const authors = new Map(); // name → { id, initials, lastIdx }
  const authorOf = (c) => {
    const name = String(c.author || '').trim() || 'Anonymous';
    if (!authors.has(name)) authors.set(name, { id: authors.size, initials: String(c.initials || '').trim() || initialsOf(name) || 'A', lastIdx: 0 });
    return authors.get(name);
  };
  const overrides = [];
  for (const [i, slide] of deck.slides.entries()) {
    const list = Array.isArray(slide.comments) ? slide.comments.filter((c) => c && typeof c === 'object') : [];
    if (!list.length || !zip.file(`ppt/slides/slide${i + 1}.xml`)) continue;
    const cms = [];
    for (const c of list) {
      const a = authorOf(c);
      const idx = ++a.lastIdx;
      // In master units; (0, 0) is a comment on the whole slide.
      const anchored = Number.isFinite(c.x) && Number.isFinite(c.y);
      const pos = anchored ? { x: Math.round(c.x * MASTER_UNITS_PER_PT), y: Math.round(c.y * MASTER_UNITS_PER_PT) } : { x: 0, y: 0 };
      cms.push(cmXml(a.id, idx, c, pos, null));
      for (const r of Array.isArray(c.replies) ? c.replies : []) {
        if (!r || typeof r !== 'object') continue;
        const ra = authorOf(r);
        cms.push(cmXml(ra.id, ++ra.lastIdx, r, pos, { authorId: a.id, idx }));
      }
    }
    const name = `comment${i + 1}.xml`;
    zip.file(`ppt/comments/${name}`, `${XML_HEAD}<p:cmLst ${P_NS}>${cms.join('')}</p:cmLst>`);
    await addRelationship(zip, `ppt/slides/_rels/slide${i + 1}.xml.rels`, 'rIdLsComments', 'comments', `../comments/${name}`);
    overrides.push(`<Override PartName="/ppt/comments/${name}" ContentType="${TYPES}.comments+xml"/>`);
  }
  if (!authors.size) return;
  const list = [...authors].map(([name, a]) => `<p:cmAuthor id="${a.id}" name="${xmlText(name)}" initials="${xmlText(a.initials)}" lastIdx="${a.lastIdx}" clrIdx="${a.id % 8}"/>`);
  zip.file('ppt/commentAuthors.xml', `${XML_HEAD}<p:cmAuthorLst ${P_NS}>${list.join('')}</p:cmAuthorLst>`);
  await addRelationship(zip, 'ppt/_rels/presentation.xml.rels', 'rIdLsCommentAuthors', 'commentAuthors', 'commentAuthors.xml');
  overrides.push(`<Override PartName="/ppt/commentAuthors.xml" ContentType="${TYPES}.commentAuthors+xml"/>`);
  const types = await zip.file('[Content_Types].xml').async('string');
  zip.file('[Content_Types].xml', types.replace('</Types>', `${overrides.join('')}</Types>`));
}
