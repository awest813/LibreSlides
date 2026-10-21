import { kid, attr, contentOf } from './ooxml.js';

/**
 * OLE objects (an embedded workbook, document, slide or old-style picture)
 * aren't something LibreSlides can open, but PowerPoint keeps a picture of
 * each, which is what it shows until the object is opened: that picture is
 * brought in in the object's place.
 */

/** The p:oleObj of an OLE object's graphic data (directly, or in the branch of mc:AlternateContent that's read). */
export const oleObjectOf = (graphicData) => contentOf(graphicData).find((c) => c.localName === 'oleObj') || null;

const attrOf = (tag, name) => new RegExp(`\\s${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(tag)?.slice(1).find((v) => v != null) ?? null;

/**
 * The picture part of the VML shape `spid` in a legacy drawing (VML is
 * often not well-formed XML, so it's searched rather than parsed), or null.
 */
export function vmlImageId(text, spid) {
  if (!text || !spid) return null;
  for (const m of text.matchAll(/<v:shape\b[^>]*>/g)) {
    const tag = m[0];
    if (attrOf(tag, 'id') !== spid && attrOf(tag, 'o:spid') !== spid) continue;
    const end = text.indexOf('</v:shape>', m.index);
    const body = text.slice(m.index, end < 0 ? text.length : end);
    const data = /<v:imagedata\b[^>]*>/.exec(body)?.[0];
    return data ? attrOf(data, 'o:relid') || attrOf(data, 'r:id') || attrOf(data, 'relid') : null;
  }
  return null;
}

/**
 * What an OLE object shows: { pic } (a p:pic inside the p:oleObj, as files
 * since PowerPoint 2010 have it) or { path } (the picture part of its shape
 * in the slide's VML drawing, as older files have it); null when there's none.
 */
export async function olePreview(pkg, ole, rels) {
  if (!ole) return null;
  const pic = kid(ole, 'pic');
  if (pic) return { pic };
  const spid = attr(ole, 'spid');
  if (!spid) return null;
  for (const rel of Object.values(rels || {})) {
    if (rel.type !== 'vmlDrawing' || rel.external) continue;
    let text;
    try {
      text = await pkg.text(rel.path);
    } catch (err) {
      if (pkg.full) throw err;
      continue;
    }
    const id = vmlImageId(text, spid);
    if (!id) continue;
    const target = (await pkg.rels(rel.path))[id];
    if (target && !target.external) return { path: target.path };
  }
  return null;
}
