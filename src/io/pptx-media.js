import { uid } from '../model/deck.js';
import { kid, kids, path, desc, attr, rid, addMediaTo } from './ooxml.js';
import { MEDIA_TYPES, MEDIA_CONTENT_TYPES, dataType } from './media-files.js';

/**
 * Video and audio in PowerPoint files to and from LibreSlides media elements.
 *
 * PowerPoint keeps a video as a picture (its poster frame) that also points
 * at the video file: <a:videoFile r:link> for older readers and
 * <p14:media r:embed> for the copy inside the package. Sounds use
 * <a:audioFile>. Whether it plays by itself and loops is in the slide's
 * timeline (<p:timing>). Formats browsers can't play (Windows Media, AVI,
 * MPEG-1…) and videos linked from elsewhere stay as their poster picture.
 */

// ------------------------------------------------------------------ reading

/** The types browsers play, by file extension. */
const EXT_TYPE = { mp4: 'video/mp4', m4v: 'video/mp4', mov: 'video/quicktime', webm: 'video/webm', ogv: 'video/ogg', ogg: 'audio/ogg', oga: 'audio/ogg', mp3: 'audio/mpeg', m4a: 'audio/mp4', wav: 'audio/wav' };

/** The media file a relationship points at, as a data URL; null if it's elsewhere or can't play. */
function mediaData(pkg, target, kind) {
  if (!target || target.external) return null;
  let type = EXT_TYPE[(target.path.split('.').pop() || '').toLowerCase()];
  if (!type) return null;
  // An .mp4 that PowerPoint calls a sound is a sound (and an .ogg it calls a video, a video).
  if (kind === 'audio' && type === 'video/mp4') type = 'audio/mp4';
  if (kind === 'video' && type === 'audio/ogg') type = 'video/ogg';
  return pkg.dataUrl(target.path, type);
}

/** Whether the slide's timeline plays a shape as the slide appears, and loops it. */
export function playbackOf(doc, spid) {
  const timing = doc && spid ? desc(doc.documentElement, 'timing')[0] : null;
  const out = { autoplay: false, loop: false, duration: null };
  if (!timing) return out;
  const targets = (el) => desc(el, 'spTgt').some((t) => attr(t, 'spid') === spid);
  // Looping: the shape's media node repeats until stopped.
  out.loop = desc(timing, 'cMediaNode').some((n) => targets(n) && attr(kid(n, 'cTn'), 'repeatCount') === 'indefinite');
  // The "play" commands; the length PowerPoint gives them is the media's.
  const plays = desc(timing, 'cmd').filter((c) => /^playFrom/.test(attr(c, 'cmd') || '') && targets(c));
  const dur = +attr(path(plays[0], 'cBhvr', 'cTn'), 'dur');
  if (dur > 1) out.duration = dur / 1000;
  // Playing by itself: a play command in the main sequence's first step, when that step starts without a click.
  const main = desc(timing, 'cTn').find((c) => attr(c, 'nodeType') === 'mainSeq');
  const first = kids(kid(main, 'childTnLst'), 'par')[0];
  const firstEffect = desc(first, 'cTn').find((c) => /Effect$/.test(attr(c, 'nodeType') || ''));
  out.autoplay = !!firstEffect && attr(firstEffect, 'nodeType') !== 'clickEffect' && plays.some((c) => first.contains(c));
  return out;
}

/** A p:pic carrying a video or audio file as a media element, or null to treat it as a plain picture. */
export async function readMedia({ pkg, node, rels, box, deck }) {
  const nvPr = path(node, 'nvPicPr', 'nvPr');
  const link = ['videoFile', 'quickTimeFile', 'audioFile', 'wavAudioFile'].map((n) => kid(nvPr, n)).find(Boolean);
  const embed = desc(nvPr, 'media')[0]; // <p14:media>, in the extension list
  if ((!link && !embed) || !box) return null;
  let kind = /audio/i.test(link?.localName || '') ? 'audio' : 'video';
  // The copy inside the package first; the older link may point at a file elsewhere.
  let data = await mediaData(pkg, rels[rid(embed, 'embed')], kind);
  if (!data) data = await mediaData(pkg, rels[rid(link, 'link') || rid(link, 'embed')], kind);
  if (!data) return null;
  if (dataType(data).startsWith('audio/')) kind = 'audio';

  const blip = desc(kid(node, 'blipFill'), 'blip')[0];
  const posterTarget = rels[rid(blip, 'embed')];
  const poster = posterTarget && !posterTarget.external ? await pkg.dataUrl(posterTarget.path) : null;
  const spid = attr(path(node, 'nvPicPr', 'cNvPr'), 'id');
  const play = playbackOf(node.ownerDocument, spid);
  const el = {
    id: uid(), type: 'media', kind, media: addMediaTo(deck, data), x: box.x, y: box.y, w: box.w, h: box.h, rot: box.rot || 0,
    autoplay: play.autoplay, loop: play.loop,
  };
  if (poster) el.poster = addMediaTo(deck, poster);
  if (play.duration) el.duration = play.duration;
  return el;
}

// ------------------------------------------------------------------ writing

/** The file of a media element as pptxgenjs takes it ("video/mp4;base64,…"), or null to leave it out. */
function exportable(el, deck) {
  const data = deck.media?.[el.media];
  const extn = MEDIA_TYPES[dataType(data)];
  return extn && /;base64,/.test(data) ? { data: data.replace(/^data:/, ''), extn } : null;
}

/** pptxgenjs files the poster as a PNG, so other pictures are converted (kept as they are where that fails). */
async function coverImage(dataUrl) {
  if (!dataUrl || !/^data:image\//.test(dataUrl)) return null;
  if (/^data:image\/png;base64,/.test(dataUrl)) return dataUrl.replace(/^data:/, '');
  try {
    const img = new Image();
    img.src = dataUrl;
    await img.decode();
    const canvas = document.createElement('canvas');
    canvas.width = img.naturalWidth || 640;
    canvas.height = img.naturalHeight || 360;
    canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
    return canvas.toDataURL('image/png').replace(/^data:/, '');
  } catch {
    return /;base64,/.test(dataUrl) && !/svg/.test(dataUrl) ? dataUrl.replace(/^data:/, '') : null;
  }
}

/** Add a media element to a pptxgenjs slide (geom in inches). */
export async function addMediaElement(pptx, slide, el, deck, geom, objectName) {
  const file = exportable(el, deck);
  if (!file) return;
  const opts = { type: el.kind === 'audio' ? 'audio' : 'video', data: file.data, extn: file.extn, x: geom.x, y: geom.y, w: geom.w, h: geom.h };
  // Without a poster pptxgenjs shows its own play button.
  const cover = await coverImage(deck.media?.[el.poster]);
  if (cover) opts.cover = cover;
  if (objectName) opts.objectName = objectName;
  slide.addMedia(opts);
}

/**
 * Finish the media pptxgenjs wrote: sounds marked as audio (it writes every
 * file as a video), rotation and PowerPoint's content types. (Shape ids are
 * made unique for the whole slide afterwards, and the timeline, for
 * autoplay and looping, is written with the slide's animations: see
 * pptx-export.js and pptx-timing.js.)
 */
export async function finishMedia(zip, deck) {
  let any = false;
  for (const [i, slide] of deck.slides.entries()) {
    const els = slide.elements.filter((el) => el.type === 'media' && exportable(el, deck));
    const file = els.length && zip.file(`ppt/slides/slide${i + 1}.xml`);
    if (!file) continue;
    any = true;
    let xml = await file.async('string');
    let k = 0;
    // The media pictures are in the same order as the elements they came from.
    xml = xml.replace(/<p:pic>(?:(?!<\/p:pic>)[\s\S])*?<a:videoFile[\s\S]*?<\/p:pic>/g, (pic) => {
      const el = els[k++];
      if (!el) return pic;
      if (el.kind === 'audio') pic = pic.replace('<a:videoFile ', '<a:audioFile ');
      if (el.rot) pic = pic.replace(/<a:xfrm(?=[\s>])/, `<a:xfrm rot="${Math.round((((el.rot % 360) + 360) % 360) * 60000)}"`);
      return pic;
    });
    zip.file(file.name, xml);
  }
  if (!any) return;
  // pptxgenjs names a type after the extension ("audio/mp3"); PowerPoint expects the standard ones.
  const types = zip.file('[Content_Types].xml');
  let ct = await types.async('string');
  ct = ct.replace(/<Default Extension="([a-z0-9]+)" ContentType="[^"]*"\/>/gi, (m, ext) => (MEDIA_CONTENT_TYPES[ext.toLowerCase()]
    ? `<Default Extension="${ext}" ContentType="${MEDIA_CONTENT_TYPES[ext.toLowerCase()]}"/>` : m));
  zip.file('[Content_Types].xml', ct);
}
