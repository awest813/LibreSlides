import { slideSections, hasSections, DEFAULT_SECTION } from '../model/deck.js';
import { desc, attr, xmlText, guid } from './ooxml.js';

/**
 * Slide sections in .pptx files. PowerPoint 2010 and later list them in an
 * extension of ppt/presentation.xml:
 *
 *   <p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">
 *     <p14:sectionLst xmlns:p14="…/powerpoint/2010/main">
 *       <p14:section name="Intro" id="{GUID}"><p14:sldIdLst><p14:sldId id="256"/>…
 *
 * where the slide ids are those of p:sldIdLst. Every slide is in a section
 * once there are any; the ones before the first section that has a name of
 * its own are in "Default Section". In a deck, a section is the name on its
 * first slide (see model/deck.js).
 */
export const SECTIONS_EXT = '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}';
const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';

// ------------------------------------------------------------------ reading

/**
 * Put the sections of presentation.xml (`pres`, a document) on the slides
 * they start. `sldIds` holds the p:sldId id of each of `slides`. Sections
 * without slides (or none that could be read) are left out, and so is a
 * first section called "Default Section", which is what PowerPoint calls the
 * slides that are in no section.
 */
export function readSections(pres, slides, sldIds) {
  const index = new Map(sldIds.map((id, i) => [String(id), i]));
  const list = desc(pres, 'sectionLst')[0];
  if (!list) return;
  let first = true;
  for (const sec of desc(list, 'section')) {
    const at = desc(sec, 'sldId').map((s) => index.get(attr(s, 'id'))).filter((i) => i != null);
    if (!at.length) continue;
    const start = Math.min(...at);
    const name = (attr(sec, 'name') || '').trim() || 'Untitled Section';
    if (!(first && start === 0 && name === DEFAULT_SECTION)) slides[start].section = name;
    first = false;
  }
}

// ------------------------------------------------------------------ writing

/** Add a p:ext to presentation.xml's p:extLst (its last child), making the list if there's none. */
export function addPresentationExt(presXml, ext) {
  const end = presXml.lastIndexOf('</p:presentation>');
  if (end < 0) return presXml;
  const head = presXml.slice(0, end).trimEnd();
  if (head.endsWith('</p:extLst>')) {
    const at = head.length - '</p:extLst>'.length;
    return presXml.slice(0, at) + ext + presXml.slice(at);
  }
  return `${presXml.slice(0, end)}<p:extLst>${ext}</p:extLst>${presXml.slice(end)}`;
}

/** The p14:sectionLst extension for a deck whose slides have the p:sldId ids `sldIds`. */
export function sectionsXml(slides, sldIds) {
  const sections = slideSections(slides).map((g) => {
    const ids = sldIds.slice(g.start, g.end).map((id) => `<p14:sldId id="${id}"/>`).join('');
    return `<p14:section name="${xmlText(g.name || DEFAULT_SECTION)}" id="${guid()}"><p14:sldIdLst>${ids}</p14:sldIdLst></p14:section>`;
  });
  return `<p:ext uri="${SECTIONS_EXT}"><p14:sectionLst xmlns:p14="${P14}">${sections.join('')}</p14:sectionLst></p:ext>`;
}

/** Write the deck's sections into the generated .pptx (a JSZip), if it has any. */
export async function writeSections(zip, deck) {
  if (!hasSections(deck.slides)) return;
  const path = 'ppt/presentation.xml';
  const pres = await zip.file(path)?.async('string');
  if (!pres) return;
  const list = /<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/.exec(pres)?.[1] || '';
  const ids = [...list.matchAll(/<p:sldId\b[^>]*?\sid="(\d+)"/g)].map((m) => m[1]);
  if (ids.length !== deck.slides.length) return; // not one slide each: better no sections than wrong ones
  zip.file(path, addPresentationExt(pres, sectionsXml(deck.slides, ids)));
}
