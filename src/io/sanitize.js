import DOMPurify from 'dompurify';

/**
 * Text in slides is HTML. Anything that comes from outside (opened files,
 * restored backups, the clipboard) goes through here first.
 */
const TEXT_CONFIG = {
  ALLOWED_TAGS: ['p', 'div', 'br', 'span', 'b', 'strong', 'i', 'em', 'u', 's', 'strike', 'sub', 'sup', 'ul', 'ol', 'li', 'a', 'font'],
  ALLOWED_ATTR: ['style', 'href', 'color', 'face', 'size'],
  ALLOW_DATA_ATTR: false,
};

let hooked = false;
function ensureHooks() {
  if (hooked) return;
  hooked = true;
  DOMPurify.addHook('afterSanitizeAttributes', (node) => {
    // No external resources from inline styles; links open in a new tab.
    if (node.hasAttribute?.('style') && /url\s*\(|expression\s*\(/i.test(node.getAttribute('style'))) node.removeAttribute('style');
    if (node.tagName === 'A') {
      node.setAttribute('target', '_blank');
      node.setAttribute('rel', 'noopener noreferrer');
    }
  });
}

export function sanitizeHtml(html) {
  if (!html) return '';
  ensureHooks();
  return DOMPurify.sanitize(String(html), { ...TEXT_CONFIG, ADD_ATTR: ['target', 'rel'] });
}

const isImageData = (v) => typeof v === 'string' && /^data:image\/(png|jpe?g|gif|webp|svg\+xml|bmp|avif);/i.test(v);

/**
 * A colour value we draw: hex, rgb()/rgba(), a theme colour, none or transparent.
 * Anything else (such as `url(…)`, which would fetch from the network) is dropped.
 */
const COLOR = /^(#[0-9a-f]{3,8}|rgba?\([\d\s.,%/]+\)|theme:[a-z0-9]+|none|transparent)$/i;
export const safeColor = (v) => (typeof v === 'string' && COLOR.test(v.trim()) ? v.trim() : null);

/** Clean an element from an untrusted source (file, backup, clipboard) in place. */
export function sanitizeElement(el) {
  if (el.text) el.text = sanitizeHtml(el.text);
  if (Array.isArray(el.cells)) {
    for (const row of el.cells) {
      for (const c of row) {
        c.text = sanitizeHtml(c.text);
        if (c.fill != null) c.fill = safeColor(c.fill);
      }
    }
  }
  if (el.fill != null) el.fill = safeColor(el.fill) || 'none';
  if (el.stroke && typeof el.stroke === 'object') el.stroke.color = safeColor(el.stroke.color) || 'none';
  if (el.style && typeof el.style === 'object') {
    if (el.style.color != null) el.style.color = safeColor(el.style.color);
    if (typeof el.style.font === 'string') el.style.font = el.style.font.replace(/[^\w\s-]/g, '').slice(0, 64);
  }
  return el;
}

/** Clean every piece of HTML, colour and picture in a deck from an untrusted source. */
export function sanitizeDeck(deck) {
  for (const s of deck.slides || []) {
    s.notes = typeof s.notes === 'string' ? s.notes : '';
    for (const el of s.elements || []) sanitizeElement(el);
    const bg = s.background;
    if (bg && typeof bg === 'object') {
      for (const key of ['color', 'from', 'to']) if (bg[key] != null) bg[key] = safeColor(bg[key]) || '#ffffff';
    } else s.background = null;
  }
  for (const [id, data] of Object.entries(deck.media || {})) {
    if (!isImageData(data)) delete deck.media[id];
    else if (/^data:image\/svg\+xml/i.test(data)) {
      // SVG can carry script; it only ever renders through <img>, which doesn't run it, but keep it tidy.
      try {
        const [head, body] = data.split(',');
        const svg = head.endsWith(';base64') ? atob(body) : decodeURIComponent(body);
        deck.media[id] = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(DOMPurify.sanitize(svg, { USE_PROFILES: { svg: true, svgFilters: true } }))}`;
      } catch {
        delete deck.media[id];
      }
    }
  }
  // Pictures that were dropped (unsupported or unsafe) leave nothing behind.
  for (const s of deck.slides || []) {
    s.elements = (s.elements || []).filter((el) => el.type !== 'image' || deck.media?.[el.media]);
    if (s.background?.type === 'image' && !deck.media?.[s.background.media]) s.background = null;
  }
  return deck;
}
