import { h, formatDate } from '../ui/dom.js';
import { icon } from '../ui/icons.js';
import { promptDialog } from '../ui/dialog.js';
import { uid, boxOf, initialsOf } from '../model/deck.js';

/**
 * Slide comments in the editor: the Comments tab of the side panel (threads
 * with replies, resolving and deleting), the markers on the canvas and the
 * name comments are signed with. Every change is one undo step. Comments are
 * plain text, so they only ever reach the page as text nodes.
 */
const AUTHOR_KEY = 'ls:author';
let sessionAuthor = null; // when storage is blocked, the name lasts until the page closes

/** The name the user signs comments with, if they've given one. */
export function savedAuthor() {
  try {
    return localStorage.getItem(AUTHOR_KEY) || sessionAuthor;
  } catch {
    return sessionAuthor;
  }
}

/** Ask for the user's name and remember it. Resolves to the name, or null when cancelled. */
export async function askAuthor() {
  const r = await promptDialog({
    title: 'Your name',
    fields: [{ name: 'name', label: 'Name shown on your comments', value: savedAuthor() || 'You' }],
    confirmLabel: 'Save',
  });
  if (!r) return null;
  const name = r.name.trim().slice(0, 100) || 'You';
  sessionAuthor = name;
  try { localStorage.setItem(AUTHOR_KEY, name); } catch { /* storage blocked */ }
  return name;
}

/** A new comment or reply signed by the user (asked for their name the first time). */
async function signed(text) {
  const author = savedAuthor() || await askAuthor();
  if (!author) return null;
  const initials = initialsOf(author);
  return { id: uid(), author, ...(initials ? { initials } : {}), date: new Date().toISOString(), text };
}

const isAnchored = (c) => Number.isFinite(c.x) && Number.isFinite(c.y);
const findComment = (slide, id) => slide.comments?.find((c) => c.id === id) || null;

/** How many comments on a slide are still open. */
export const openComments = (slide) => (Array.isArray(slide.comments) ? slide.comments.filter((c) => !c.resolved).length : 0);

// ------------------------------------------------------------------ changes

/** Comment on the current slide, at the selected object's top-left corner if one is selected. Resolves to its id. */
export async function addComment(app, text) {
  const c = await signed(text);
  if (!c) return null;
  const el = app.firstSelected();
  if (el) {
    const b = boxOf(el);
    Object.assign(c, { x: Math.round(b.x), y: Math.round(b.y) });
  }
  app.changeSlide((s) => { (s.comments ||= []).push(c); });
  return c.id;
}

export async function replyToComment(app, id, text) {
  if (!findComment(app.slide, id)) return null;
  const r = await signed(text);
  if (!r) return null;
  app.changeSlide((s) => {
    const c = findComment(s, id);
    if (c) (c.replies ||= []).push(r);
  });
  return r.id;
}

export function setResolved(app, id, resolved) {
  app.changeSlide((s) => {
    const c = findComment(s, id);
    if (!c) return;
    if (resolved) c.resolved = true;
    else delete c.resolved;
  });
  // Rebuilt even while a reply box keeps the focus (Safari doesn't focus clicked buttons).
  app.panel.update(true);
}

/** Delete a whole thread, or one reply in it. */
export function deleteComment(app, id, replyId = null) {
  app.changeSlide((s) => {
    const c = findComment(s, id);
    if (!c) return;
    if (replyId) {
      c.replies = (c.replies || []).filter((r) => r.id !== replyId);
      if (!c.replies.length) delete c.replies;
    } else {
      s.comments = s.comments.filter((x) => x !== c);
      if (!s.comments.length) delete s.comments;
    }
  });
  app.panel.update(true);
}

// ------------------------------------------------------------------ opening the panel

/** Insert › Comment: the Comments tab, with the cursor in the new-comment box. */
export function startComment(app) {
  app.openPanel('comments');
  app.panel.body.querySelector('[data-composer="new"] textarea')?.focus();
}

/** Show one thread in the Comments tab (from its marker on the slide). */
export function showThread(app, id) {
  app.panel.activeComment = id;
  app.openPanel('comments');
  app.canvas.drawOverlay();
  app.panel.body.querySelector(`[data-thread="${CSS.escape(id)}"]`)?.scrollIntoView({ block: 'nearest' });
}

// ------------------------------------------------------------------ the canvas

/**
 * Markers for the open comments that have a place on the slide, drawn in the
 * canvas overlay (so never in thumbnails, the slideshow, print or exports),
 * numbered as in the Comments tab, and under the selection's handles.
 */
export function drawCommentMarkers(canvas, overlay) {
  const { app, scale } = canvas;
  const list = app.slide?.comments;
  if (!Array.isArray(list)) return;
  const stacked = new Map(); // markers for the same point line up to the left
  const markers = [];
  list.forEach((c, i) => {
    if (c.resolved || !isAnchored(c)) return;
    const at = `${c.x},${c.y}`;
    const n = stacked.get(at) || 0;
    stacked.set(at, n + 1);
    markers.push(h('button', {
      type: 'button', tabindex: '-1', class: `comment-marker${app.panel?.activeComment === c.id ? ' is-active' : ''}`,
      style: { left: `${c.x * scale - n * 36}px`, top: `${c.y * scale}px` }, title: `${c.author}: ${String(c.text).slice(0, 120)}`, html: icon('comment'),
      // Not a click on the slide: no deselecting or marquee.
      onpointerdown: (e) => e.stopPropagation(),
      onclick: () => showThread(app, c.id),
    }, h('span', {}, String(i + 1))));
  });
  overlay.prepend(...markers);
}

// ------------------------------------------------------------------ the Comments tab

/** What the user has typed but not posted yet, by box ("new" or a thread id), kept across rebuilds of the panel. */
const drafts = new WeakMap();
const draftsOf = (app) => {
  if (!drafts.has(app)) drafts.set(app, new Map());
  return drafts.get(app);
};

/** A box to type a comment or reply in; Ctrl+Enter (⌘Enter) or the button posts it. */
function composer(app, key, { placeholder, label, rows, post }) {
  const store = draftsOf(app);
  const input = h('textarea', { rows, placeholder, 'aria-label': placeholder, maxlength: '10000' });
  input.value = store.get(key) || '';
  const button = h('button', { type: 'button', class: `btn${key === 'new' ? ' btn-primary' : ''}`, disabled: !input.value.trim() || null }, label);
  const send = async () => {
    const text = input.value.trim();
    if (!text || !await post(text)) return;
    store.delete(key);
    app.panel.update(true);
    app.panel.body.querySelector(`[data-composer="${CSS.escape(key)}"] textarea`)?.focus();
  };
  input.addEventListener('input', () => {
    store.set(key, input.value);
    button.disabled = !input.value.trim();
  });
  input.addEventListener('keydown', (e) => {
    if (e.key !== 'Enter' || !(e.ctrlKey || e.metaKey)) return;
    // Not the editor's Ctrl+Enter (start the slideshow).
    e.preventDefault();
    e.stopPropagation();
    send();
  });
  button.addEventListener('click', send);
  return h('div', { class: 'comment-composer', dataset: { composer: key } }, input, button);
}

/** Author and date line of a comment or reply. */
function byline(c, number = null) {
  const date = c.date && !Number.isNaN(Date.parse(c.date)) ? h('time', { datetime: c.date, title: new Date(c.date).toLocaleString() }, formatDate(c.date)) : null;
  return h('div', { class: 'comment-head' },
    number ? h('span', { class: 'comment-num', title: 'Marked on the slide' }, String(number)) : null,
    h('b', { class: 'comment-author' }, c.author || 'Anonymous'),
    date);
}

function thread(app, c, i) {
  const active = app.panel.activeComment === c.id;
  const replies = (c.replies || []).map((r) => h('div', { class: 'comment-reply' },
    byline(r),
    h('p', { class: 'comment-text' }, r.text),
    h('button', { type: 'button', class: 'icon-btn comment-delete', 'aria-label': `Delete reply by ${r.author || 'Anonymous'}`, title: 'Delete reply', html: icon('trash'), onclick: () => deleteComment(app, c.id, r.id) })));
  const actions = h('div', { class: 'comment-actions' },
    c.resolved ? h('span', { class: 'comment-tag' }, 'Resolved') : null,
    h('button', { type: 'button', class: 'btn btn-ghost', html: `${icon(c.resolved ? 'rotateCcw' : 'check')}<span>${c.resolved ? 'Reopen' : 'Resolve'}</span>`, onclick: () => setResolved(app, c.id, !c.resolved) }),
    h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Delete comment', title: 'Delete comment and its replies', html: icon('trash'), onclick: () => deleteComment(app, c.id) }));
  const el = h('article', { class: `comment-thread${c.resolved ? ' is-resolved' : ''}${active ? ' is-active' : ''}`, dataset: { thread: c.id } },
    byline(c, isAnchored(c) ? i + 1 : null),
    h('p', { class: 'comment-text' }, c.text),
    replies.length ? h('div', { class: 'comment-replies' }, ...replies) : null,
    c.resolved ? null : composer(app, c.id, { placeholder: 'Reply…', label: 'Reply', rows: 1, post: (text) => replyToComment(app, c.id, text) }),
    actions);
  // Working in a thread highlights its marker on the slide.
  el.addEventListener('focusin', () => {
    if (app.panel.activeComment === c.id) return;
    app.panel.activeComment = c.id;
    app.panel.body.querySelectorAll('.comment-thread.is-active').forEach((t) => t.classList.remove('is-active'));
    el.classList.add('is-active');
    app.canvas.drawOverlay();
  });
  return el;
}

/** The Comments tab: a box for a new comment, then this slide's threads. */
export function commentsPanel(app) {
  const { slide } = app;
  const list = Array.isArray(slide.comments) ? slide.comments : [];
  const name = savedAuthor();
  const selected = app.firstSelected();
  const out = [h('section', { class: 'panel-section' },
    h('h3', {}, 'New comment'),
    composer(app, 'new', { placeholder: selected ? 'Comment on the selected object…' : 'Comment on this slide…', label: 'Comment', rows: 3, post: (text) => addComment(app, text) }),
    name ? h('p', { class: 'pf-hint' }, `Commenting as ${name}. `, h('button', { type: 'button', class: 'link-btn', onclick: async () => { if (await askAuthor()) app.panel.update(true); } }, 'Change name')) : null)];
  const threads = list.map((c, i) => thread(app, c, i));
  const open = list.filter((c) => !c.resolved).length;
  out.push(h('section', { class: 'panel-section' },
    h('h3', {}, list.length ? `Slide ${app.slideIndex + 1} · ${open} open` : `Slide ${app.slideIndex + 1}`),
    threads.length ? h('div', { class: 'comment-list' }, ...threads) : h('p', { class: 'pf-hint' }, 'No comments on this slide yet.')));
  // Reviewing: the next slide that still has open comments.
  const slides = app.deck.slides;
  const next = [...slides.slice(app.slideIndex + 1), ...slides.slice(0, app.slideIndex)].find((s) => openComments(s));
  if (next) {
    const n = slides.indexOf(next);
    out.push(h('button', { type: 'button', class: 'btn', onclick: () => app.goToSlide(n) }, `Next comments: slide ${n + 1}`));
  }
  return out;
}
