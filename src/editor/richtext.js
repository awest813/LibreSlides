/**
 * Rich-text helpers for text boxes. While a box is being edited, formatting
 * goes through the browser's editing commands on the selection; when whole
 * boxes are selected instead, formatting is set on the element's style and
 * any conflicting inline formatting inside it is removed (as in Slides).
 */

const PROP_CSS = {
  font: ['font-family'],
  size: ['font-size'],
  color: ['color'],
  bold: ['font-weight'],
  italic: ['font-style'],
  underline: ['text-decoration', 'text-decoration-line'],
  strike: ['text-decoration', 'text-decoration-line'],
  highlight: ['background-color', 'background'],
  align: ['text-align'],
  // A box's line and paragraph spacing replace what its paragraphs had of their own (from PowerPoint).
  lineHeight: ['line-height'],
  spacing: ['margin-top', 'margin-bottom'],
  effect: ['-webkit-text-stroke', '-webkit-text-stroke-width', '-webkit-text-stroke-color', 'text-shadow', 'background-image', '-webkit-background-clip', 'background-clip', '-webkit-text-fill-color'],
};
const PROP_TAGS = { bold: ['B', 'STRONG'], italic: ['I', 'EM'], underline: ['U'], strike: ['S', 'STRIKE'] };

function unwrap(el) {
  el.replaceWith(...el.childNodes);
}

/** Remove inline formatting for `prop` from an HTML fragment. */
export function stripInline(html, prop) {
  if (!html) return html;
  const tpl = document.createElement('template');
  tpl.innerHTML = html;
  const root = tpl.content;
  for (const tag of PROP_TAGS[prop] || []) root.querySelectorAll(tag).forEach(unwrap);
  if (prop === 'font') root.querySelectorAll('font[face]').forEach((f) => f.removeAttribute('face'));
  if (prop === 'color') root.querySelectorAll('font[color]').forEach((f) => f.removeAttribute('color'));
  if (prop === 'size') root.querySelectorAll('font[size]').forEach((f) => f.removeAttribute('size'));
  for (const el of root.querySelectorAll('[style]')) {
    for (const css of PROP_CSS[prop] || []) el.style.removeProperty(css);
    if (!el.getAttribute('style')?.trim()) el.removeAttribute('style');
  }
  // Spans and fonts left with no attributes do nothing.
  root.querySelectorAll('span:not([style]), font:not([face]):not([color]):not([size])').forEach(unwrap);
  const div = document.createElement('div');
  div.append(root);
  return div.innerHTML;
}

/**
 * Turn <font> tags (from execCommand) into styled spans. <font size="7"> marks
 * text whose size is being set: with `markerSize` it becomes that size (and
 * the spans are returned); without, it is left for the caller setting it.
 */
export function normalizeFontTags(root, markerSize = null) {
  const sized = [];
  for (const f of root.querySelectorAll('font')) {
    const marker = f.getAttribute('size') === '7';
    if (marker && !markerSize) continue;
    const span = document.createElement('span');
    if (marker) {
      span.style.fontSize = `${markerSize}px`;
      for (const inner of f.querySelectorAll('[style]')) inner.style.removeProperty('font-size');
      sized.push(span);
    }
    if (f.hasAttribute('face')) span.style.fontFamily = f.getAttribute('face');
    if (f.hasAttribute('color')) span.style.color = f.getAttribute('color');
    span.append(...f.childNodes);
    f.replaceWith(span);
  }
  // execCommand('fontSize') with CSS styling writes keyword sizes; those never come from us.
  root.querySelectorAll('span[style*="xxx-large"]').forEach((s) => s.style.removeProperty('font-size'));
  return sized;
}

export function isEmptyHtml(html) {
  if (!html) return true;
  const tpl = document.createElement('template');
  tpl.innerHTML = html;
  return !tpl.content.textContent.replace(/​/g, '').trim() && !tpl.content.querySelector('img');
}

/** Clean up the HTML of an edited text box. */
export function tidyHtml(html) {
  if (isEmptyHtml(html)) return '';
  const tpl = document.createElement('template');
  tpl.innerHTML = html;
  normalizeFontTags(tpl.content);
  // Zero-width spaces that held a caret for a new font size aren't needed once text was typed there.
  const walker = document.createTreeWalker(tpl.content, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (n.data.length > 1 && n.data.includes('​')) n.data = n.data.replace(/​/g, '');
  }
  // Bare text at the top level (typed into an empty box) goes in a paragraph.
  const out = document.createElement('div');
  let p = null;
  for (const node of [...tpl.content.childNodes]) {
    const inline = node.nodeType === 3 || (node.nodeType === 1 && !/^(P|DIV|UL|OL|H\d)$/.test(node.nodeName));
    if (inline) {
      if (!p) { p = document.createElement('p'); out.append(p); }
      p.append(node);
    } else {
      p = null;
      if (node.nodeName === 'DIV' && !node.querySelector('p, ul, ol, div')) {
        const para = document.createElement('p');
        for (const a of node.attributes) para.setAttribute(a.name, a.value);
        para.append(...node.childNodes);
        out.append(para);
      } else out.append(node);
    }
  }
  return out.innerHTML;
}

/** Run an editing command with CSS styling (spans with styles, not <b>/<font>). */
export function exec(cmd, value = null, { css = true } = {}) {
  try {
    document.execCommand('styleWithCSS', false, css);
  } catch { /* old browsers */ }
  return document.execCommand(cmd, false, value);
}

/**
 * Set the paragraph at the caret in (dir 1) or out (-1) by a list's inset,
 * as PowerPoint does, when it isn't a list item (the browser's indent
 * command would wrap it in a quote). Returns whether it did.
 */
export function indentParagraph(dir) {
  const node = document.getSelection()?.anchorNode;
  const block = (node?.nodeType === 1 ? node : node?.parentElement)?.closest('p, div, li');
  if (!block || block.tagName === 'LI' || block.isContentEditable === false || block.contentEditable === 'true') return false;
  const step = 1.25 * (parseFloat(getComputedStyle(block).fontSize) || 18);
  const next = Math.max(0, Math.round(((parseFloat(block.style.marginInlineStart) || 0) + dir * step) * 10) / 10);
  block.style.marginInlineStart = next ? `${next}px` : '';
  if (!block.getAttribute('style')) block.removeAttribute('style');
  return true;
}

/** Set the font size (pt) of the selected text, or of what's typed next at the caret. */
export function execFontSize(body, size) {
  const sel = document.getSelection();
  if (sel.rangeCount && sel.isCollapsed) {
    // An empty sized span holding a zero-width space, with the caret inside it.
    const span = document.createElement('span');
    span.style.fontSize = `${size}px`;
    span.textContent = '​';
    const range = sel.getRangeAt(0);
    range.insertNode(span);
    range.setStart(span.firstChild, 1);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
    return;
  }
  exec('fontSize', '7', { css: false });
  const spans = normalizeFontTags(body, size);
  // The marker tags were replaced: select the resized text again.
  if (spans.length) {
    const range = document.createRange();
    range.setStartBefore(spans[0].firstChild || spans[0]);
    range.setEndAfter(spans[spans.length - 1].lastChild || spans[spans.length - 1]);
    const sel = document.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  }
}

/**
 * Highlight all the text of a box (or remove highlights with `color` null):
 * each paragraph's content is wrapped in one highlighted span.
 */
export function highlightAll(html, color) {
  html = stripInline(html, 'highlight');
  if (!html || !color) return html;
  const tpl = document.createElement('template');
  tpl.innerHTML = html;
  for (const block of tpl.content.querySelectorAll('p, li, div')) {
    const inline = [...block.childNodes].filter((n) => !(n.nodeType === 1 && /^(UL|OL|P|DIV)$/.test(n.nodeName)));
    if (!inline.length || !inline.some((n) => n.textContent.trim())) continue;
    const span = document.createElement('span');
    span.style.backgroundColor = color;
    inline[0].before(span);
    span.append(...inline);
  }
  const div = document.createElement('div');
  div.append(tpl.content);
  return div.innerHTML;
}

/**
 * Give all the text in an HTML fragment one text effect (see TEXT_EFFECTS),
 * replacing any it had: each paragraph's content is wrapped in a styled span.
 */
export function effectAll(html, css) {
  html = stripInline(html, 'effect');
  if (!html || !css) return html;
  const tpl = document.createElement('template');
  tpl.innerHTML = html;
  for (const block of tpl.content.querySelectorAll('p, li, div')) {
    const inline = [...block.childNodes].filter((n) => !(n.nodeType === 1 && /^(UL|OL|P|DIV)$/.test(n.nodeName)));
    if (!inline.length || !inline.some((n) => n.textContent.trim())) continue;
    const span = document.createElement('span');
    span.setAttribute('style', css);
    inline[0].before(span);
    span.append(...inline);
  }
  const div = document.createElement('div');
  div.append(tpl.content);
  return div.innerHTML;
}

/** WordArt-like effects for whole text boxes, as inline CSS made from the text and accent colours. */
export const TEXT_EFFECTS = {
  shadow: () => 'text-shadow: rgba(0, 0, 0, 0.45) 2px 2px 3px',
  glow: ({ accent }) => `text-shadow: ${accent} 0px 0px 8px`,
  outline: ({ accent }) => `-webkit-text-stroke: 1px ${accent}`,
  hollow: ({ text }) => `color: transparent; -webkit-text-stroke: 1px ${text}`,
  gradient: ({ accent, accent2 }) => `color: ${accent}; background-image: linear-gradient(90deg, ${accent} 0%, ${accent2} 100%); -webkit-background-clip: text; background-clip: text; -webkit-text-fill-color: transparent`,
};

/** Which of TEXT_EFFECTS an HTML fragment's text has, if any (or 'mixed'). */
export function textEffectOf(html) {
  const found = new Set();
  const css = (html || '').match(/style="[^"]*"/g) || [];
  for (const st of css) {
    if (/-webkit-background-clip: text|background-clip: text/.test(st)) found.add('gradient');
    else if (/-webkit-text-stroke/.test(st)) found.add(/color: transparent/.test(st) ? 'hollow' : 'outline');
    else if (/text-shadow: [^;"]*\b0px 0px/.test(st)) found.add('glow');
    else if (/text-shadow/.test(st)) found.add('shadow');
  }
  return found.size > 1 ? 'mixed' : [...found][0] || 'none';
}

/** Formatting at the caret, for the toolbar. */
export function caretFormat(body) {
  const sel = document.getSelection();
  let node = sel?.anchorNode;
  if (!node || !body.contains(node)) return null;
  if (node.nodeType === 3) node = node.parentElement;
  const cs = getComputedStyle(node);
  const block = node.closest('li, p, div');
  return {
    font: cs.fontFamily.split(',')[0].replace(/["']/g, '').trim(),
    size: Math.round(parseFloat(cs.fontSize)),
    bold: document.queryCommandState('bold'),
    italic: document.queryCommandState('italic'),
    underline: document.queryCommandState('underline'),
    strike: document.queryCommandState('strikeThrough'),
    align: block ? getComputedStyle(block).textAlign.replace('start', 'left').replace('end', 'right') : cs.textAlign,
    bullets: !!node.closest('ul'),
    numbers: !!node.closest('ol'),
  };
}
