/**
 * Video and audio in PowerPoint files to and from LibreSlides media elements.
 * Not written yet: a video comes in as its poster picture.
 */

/** A p:pic carrying a video or audio file as a media element, or null to treat it as a plain picture. */
export async function readMedia(/* { pkg, node, rels, box, ctx, deck } */) {
  return null;
}

/** Add a media element to a pptxgenjs slide. */
export async function addMediaElement(/* pptx, slide, el, deck, geom */) {}
