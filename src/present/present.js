import { h, toast } from '../ui/dom.js';
import { icon } from '../ui/icons.js';
import { renderSlide, renderThumb } from '../render/render.js';
import { slideSize, slideTitle, slideAnimations, motionPath } from '../model/deck.js';

/**
 * Slideshow: plays a deck full screen with transitions and click-by-click
 * animations. The same player drives the presenter view (a console here, the
 * audience in a second window kept in step over a BroadcastChannel) and the
 * editor's Preview buttons.
 *
 * A position is { i, step }: slide index, and how many click-steps of that
 * slide's animations have played.
 */

// ------------------------------------------------------------------ positions

/** The slide's animations grouped into click-steps; each entry has delays for "with/after previous". */
export function buildSteps(slide) {
  const animated = slideAnimations(slide);
  const steps = [];
  let cur = null;
  let lastStart = 0;
  let lastEnd = 0;
  for (const { el, kind, anim } of animated) {
    const dur = anim.duration ?? 0.5;
    const start = anim.start || 'click';
    if (start === 'click' || !cur) {
      cur = [];
      steps.push(cur);
      lastStart = 0;
      lastEnd = 0;
    }
    const delay = start === 'after' ? lastEnd : start === 'with' ? lastStart : 0;
    cur.push({ id: el.id, kind, effect: anim.effect, duration: dur, delay, ...(kind === 'motion' && { path: motionPath(anim) }) });
    lastStart = delay;
    lastEnd = delay + dur;
  }
  // A first animation set to "with" or "after previous" plays as the slide appears, without a click.
  steps.auto = animated.length > 0 && (animated[0].anim.start || 'click') !== 'click';
  return steps;
}

/**
 * Where a link to "#slide:…" goes from position `pos`: 'next', 'prev', or the
 * start of a slide (first, last, or by id); null when there's no such slide.
 */
export function linkTarget(deck, target, pos) {
  if (target === 'next' || target === 'prev') return target;
  const vis = visibleIndices(deck);
  const i = target === 'first' ? vis[0] : target === 'last' ? vis[vis.length - 1] : deck.slides.findIndex((s) => s.id === target);
  return i == null || i < 0 || i === pos?.i ? null : { i, step: minStep(deck.slides[i]) };
}

/** The first position of a slide: past its automatic step, if it has one. */
const minStep = (slide) => (buildSteps(slide).auto ? 1 : 0);

const visibleIndices = (deck) => deck.slides.map((s, i) => (s.hidden ? -1 : i)).filter((i) => i >= 0);

export function nextPos(deck, pos) {
  const steps = buildSteps(deck.slides[pos.i]);
  if (pos.step < steps.length) return { i: pos.i, step: pos.step + 1 };
  const vis = visibleIndices(deck);
  const n = vis.find((i) => i > pos.i);
  return n == null ? null : { i: n, step: minStep(deck.slides[n]) };
}

export function prevPos(deck, pos) {
  if (pos.step > minStep(deck.slides[pos.i])) return { i: pos.i, step: pos.step - 1 };
  const vis = visibleIndices(deck);
  const p = [...vis].reverse().find((i) => i < pos.i);
  return p == null ? null : { i: p, step: buildSteps(deck.slides[p]).length };
}

export function firstPos(deck, from = 0) {
  const vis = visibleIndices(deck);
  const i = Math.max(0, vis.find((x) => x >= from) ?? vis[vis.length - 1] ?? from);
  return { i, step: deck.slides[i] ? minStep(deck.slides[i]) : 0 };
}

// ------------------------------------------------------------------ animation effects

/** A point of a motion path (fractions of the slide) as a CSS translate. */
const translateAt = ([x, y] = [0, 0], w, h) => `${x * w}px ${y * h}px`;

/** Keyframes for an emphasis: something already on the slide, drawing the eye and settling back. */
function emphasis(effect, opacity = 1) {
  switch (effect) {
    case 'pulse': return [{ transform: 'scale(1)' }, { transform: 'scale(1.12)' }, { transform: 'scale(1)' }];
    case 'spin': return [{ transform: 'rotate(0)' }, { transform: 'rotate(360deg)' }];
    case 'teeter': return [0, 4, -4, 4, -4, 0].map((a) => ({ transform: `rotate(${a}deg)` }));
    case 'flash': return [{ opacity }, { opacity: 0 }, { opacity }];
    default: return null;
  }
}

/** Keyframes for an entrance; `to` is the element's own opacity, where fades end. */
function entrance(effect, W, H, to = 1) {
  switch (effect) {
    case 'fade': case 'appear': return [{ opacity: 0 }, { opacity: to }];
    case 'fly-up': return [{ transform: `translateY(${H}px)`, opacity: 0 }, { transform: 'translateY(0)', opacity: to }];
    case 'fly-down': return [{ transform: `translateY(${-H}px)`, opacity: 0 }, { transform: 'translateY(0)', opacity: to }];
    case 'fly-left': return [{ transform: `translateX(${-W}px)`, opacity: 0 }, { transform: 'translateX(0)', opacity: to }];
    case 'fly-right': return [{ transform: `translateX(${W}px)`, opacity: 0 }, { transform: 'translateX(0)', opacity: to }];
    case 'zoom': return [{ transform: 'scale(0.3)', opacity: 0 }, { transform: 'scale(1)', opacity: to }];
    case 'spin': return [{ transform: 'rotate(-200deg) scale(0.3)', opacity: 0 }, { transform: 'rotate(0) scale(1)', opacity: to }];
    case 'wipe': return [{ clipPath: 'inset(0 100% 0 0)' }, { clipPath: 'inset(0 0 0 0)' }];
    default: return null;
  }
}

function transitionFrames(type) {
  switch (type) {
    case 'fade': return { incoming: [{ opacity: 0 }, { opacity: 1 }] };
    case 'slide': return { incoming: [{ transform: 'translateX(100%)' }, { transform: 'translateX(0)' }] };
    case 'push': return { incoming: [{ transform: 'translateX(100%)' }, { transform: 'translateX(0)' }], outgoing: [{ transform: 'translateX(0)' }, { transform: 'translateX(-100%)' }] };
    case 'cover': return { incoming: [{ transform: 'translateY(100%)' }, { transform: 'translateY(0)' }] };
    case 'zoom': return { incoming: [{ transform: 'scale(0.6)', opacity: 0 }, { transform: 'scale(1)', opacity: 1 }] };
    case 'flip': return {
      incoming: [{ transform: 'rotateY(-90deg)', offset: 0 }, { transform: 'rotateY(-90deg)', offset: 0.5 }, { transform: 'rotateY(0)', offset: 1 }],
      outgoing: [{ transform: 'rotateY(0)', offset: 0 }, { transform: 'rotateY(90deg)', offset: 0.5 }, { transform: 'rotateY(90deg)', offset: 1 }],
    };
    default: return null;
  }
}

const reducedMotion = () => matchMedia('(prefers-reduced-motion: reduce)').matches;

// ------------------------------------------------------------------ player

/** Pause every video and sound under `root` (a slide that's going away keeps playing otherwise). */
function stopMedia(root) {
  for (const m of root?.querySelectorAll('video, audio') || []) {
    try {
      m.pause();
    } catch { /* not a real media element (tests) */ }
  }
}

/**
 * Draws a deck position into `host`, scaled to fit, with transitions and
 * builds. With `media`, videos and sounds play (the slideshow and the
 * audience window); otherwise they show their poster (the presenter's
 * console and previews in the editor, which shouldn't make a sound).
 */
export class Player {
  constructor(deck, host, { animate = true, media = true } = {}) {
    this.deck = deck;
    this.host = host;
    this.animate = animate;
    this.media = media;
    this.pos = null;
    this.layer = null;
    this.frame = h('div', { class: 'show-frame' });
    host.append(this.frame);
    this.resize = () => this.fit();
    this.observer = new ResizeObserver(this.resize);
    this.observer.observe(host);
  }

  destroy() {
    this.observer.disconnect();
    stopMedia(this.frame);
    this.frame.remove();
  }

  fit() {
    const { w, h: H } = slideSize(this.deck);
    const r = this.host.getBoundingClientRect();
    const s = Math.min(r.width / w, r.height / H) || 1;
    this.scale = s;
    Object.assign(this.frame.style, { width: `${w * s}px`, height: `${H * s}px` });
    for (const layer of this.frame.children) {
      if (layer.firstElementChild) layer.firstElementChild.style.transform = `scale(${s})`;
    }
  }

  makeLayer(i, step) {
    const slide = this.deck.slides[i];
    const node = renderSlide(slide, this.deck, { live: this.media });
    node.style.transformOrigin = '0 0';
    node.style.transform = `scale(${this.scale || 1})`;
    const steps = buildSteps(slide);
    // Objects that come in on later clicks start hidden, those that went out on earlier ones are gone,
    // and those that moved along a path on earlier ones are where it ended.
    const { w, h: H } = slideSize(this.deck);
    for (const [k, items] of steps.entries()) {
      for (const it of items) {
        const el = node.querySelector(`:scope > .el[data-id="${CSS.escape(it.id)}"]`);
        if (!el) continue;
        if (it.kind === 'motion') {
          if (k < step) el.style.translate = translateAt(it.path.at(-1), w, H);
          continue;
        }
        if (it.kind === 'emphasis' || (it.kind === 'exit') === (k >= step)) continue;
        el.style.visibility = 'hidden';
      }
    }
    return h('div', { class: 'show-layer' }, node);
  }

  /** Show a position. `how`: 'cut' (no animation), 'auto' (transition or build when moving forward by one). */
  go(pos, how = 'auto') {
    if (!pos) return;
    const prev = this.pos;
    this.pos = pos;
    if (!this.scale) this.fit();
    const anim = this.animate && how === 'auto' && !reducedMotion() && prev;
    // One more click on the same slide: play that step's animations.
    if (anim && prev.i === pos.i && pos.step === prev.step + 1 && this.layer) {
      this.playStep(pos.i, prev.step);
      return;
    }
    const steps = buildSteps(this.deck.slides[pos.i]);
    const autoPlay = this.animate && how === 'auto' && !reducedMotion() && (!prev || prev.i !== pos.i) && steps.auto && pos.step === 1;
    const layer = this.makeLayer(pos.i, autoPlay ? 0 : pos.step);
    const old = this.layer;
    this.layer = layer;
    this.frame.append(layer);
    const tr = this.deck.slides[pos.i].transition;
    // Arriving at a slide (at its first position: past an automatic first build) plays its transition.
    const forward = prev && pos.i > prev.i && pos.step === minStep(this.deck.slides[pos.i]);
    const frames = anim && forward && tr ? transitionFrames(tr.type) : null;
    for (const a of this.running || []) a.finish();
    this.running = [];
    stopMedia(old);
    if (frames) {
      const opts = { duration: (tr.duration ?? 0.6) * 1000, easing: 'cubic-bezier(.4,0,.2,1)', fill: 'both' };
      this.running.push(layer.animate(frames.incoming, opts));
      if (old && frames.outgoing) this.running.push(old.animate(frames.outgoing, opts));
      this.running[0].finished.catch(() => {}).then(() => old?.remove());
    } else old?.remove();
    const settle = frames ? tr.duration ?? 0.6 : 0;
    // Videos and sounds set to play by themselves start once the slide is in (those built in later, as they appear).
    this.startMedia(layer, settle * 1000);
    if (autoPlay) this.playStep(pos.i, 0, settle);
  }

  /** Play the videos and sounds under `root` that start by themselves, after `delay` ms. */
  startMedia(root, delay = 0) {
    if (!this.media) return;
    for (const m of root.querySelectorAll('[data-autoplay]')) {
      if (m.closest('.el')?.style.visibility === 'hidden') continue;
      const play = () => {
        // Browsers may refuse to play sound before the viewer has clicked; the controls still work.
        if (m.isConnected) Promise.resolve().then(() => m.play()).catch(() => {});
      };
      if (delay > 0) setTimeout(play, delay);
      else play();
    }
  }

  playStep(i, k, extraDelay = 0) {
    const { w, h: H } = slideSize(this.deck);
    const items = buildSteps(this.deck.slides[i])[k] || [];
    const node = this.layer.firstElementChild;
    for (const it of items) {
      const el = node.querySelector(`:scope > .el[data-id="${CSS.escape(it.id)}"]`);
      if (!el) continue;
      const delay = (it.delay + extraDelay) * 1000;
      const opacity = +(el.style.opacity || 1);
      if (it.kind === 'motion') {
        // Along its path, and staying where it ends.
        const frames = it.path.map((pt) => ({ translate: translateAt(pt, w, H) }));
        el.animate(frames, { duration: it.duration * 1000, delay, easing: 'ease-in-out', fill: 'forwards' })
          .finished.then(() => { el.style.translate = frames.at(-1).translate; }).catch(() => {});
        continue;
      }
      if (it.kind === 'emphasis') {
        const frames = emphasis(it.effect, opacity);
        if (frames) el.animate(frames, { duration: it.duration * 1000, delay, easing: 'ease-in-out', composite: frames[0].transform ? 'add' : 'replace' });
        continue;
      }
      if (it.kind === 'exit') {
        // The entrance played backwards, then gone.
        const frames = entrance(it.effect, w, H, opacity)?.reverse();
        const out = el.animate(frames || [{ opacity }, { opacity: 0 }], { duration: it.effect === 'appear' ? 1 : it.duration * 1000, delay, easing: 'cubic-bezier(.4,0,.8,.3)', fill: 'forwards', composite: frames?.[0].transform ? 'add' : 'replace' });
        out.finished.then(() => { el.style.visibility = 'hidden'; out.cancel(); }).catch(() => {});
        (this.running ||= []).push(out);
        continue;
      }
      el.style.visibility = '';
      this.startMedia(el, delay);
      const frames = entrance(it.effect, w, H, opacity);
      if (!frames) continue;
      const duration = it.effect === 'appear' ? 1 : it.duration * 1000;
      el.animate(frames, { duration, delay, easing: 'cubic-bezier(.2,.7,.3,1)', fill: 'backwards', composite: frames[0].transform ? 'add' : 'replace' });
    }
  }
}

// ------------------------------------------------------------------ fullscreen slideshow

function enterFullscreen(el) {
  if (document.fullscreenElement) return;
  el.requestFullscreen?.({ navigationUI: 'hide' }).catch(() => {});
}

/**
 * Present `deck` in this window. Returns a controller with `active` and `exit()`.
 * `onExit(slideIndex)` is called when the show ends.
 */
export function startSlideshow(deck, { start = 0, onExit = null, remote = null } = {}) {
  const host = h('div', { class: 'show-host' });
  const root = h('div', { class: 'show', role: 'application', 'aria-roledescription': 'slideshow', tabindex: '-1' }, host);
  const player = new Player(deck, host);
  const counter = h('span', { class: 'show-count' });
  const notesBox = h('div', { class: 'show-notes', hidden: true });
  const curtain = h('div', { class: 'show-curtain', hidden: true });
  const endScreen = h('div', { class: 'show-end', hidden: true }, 'End of slideshow. Click or press any key to exit.');
  const controls = h('div', { class: 'show-controls' },
    h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Previous', html: icon('chevronLeft'), onclick: (e) => { e.stopPropagation(); prev(); } }),
    counter,
    h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Next', html: icon('chevronRight'), onclick: (e) => { e.stopPropagation(); next(); } }),
    h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Speaker notes', title: 'Speaker notes (S)', html: icon('notes'), onclick: (e) => { e.stopPropagation(); toggleNotes(); } }),
    h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Full screen', title: 'Full screen (F)', html: icon('fullscreen'), onclick: (e) => { e.stopPropagation(); toggleFs(); } }),
    h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Exit slideshow', title: 'Exit (Esc)', html: icon('close'), onclick: (e) => { e.stopPropagation(); exit(); } }));
  root.append(curtain, endScreen, notesBox, controls);
  document.body.append(root);
  root.focus();
  if (!remote) enterFullscreen(root);

  let pos = firstPos(deck, start);
  let ended = false;
  let typed = '';
  let idle;
  const ctl = { active: true, exit: () => exit(), get pos() { return pos; }, go: (p, how) => show(p, how), blank: (color) => blank(color) };

  function show(p, how = 'auto') {
    if (!p) return;
    pos = p;
    ended = false;
    endScreen.hidden = true;
    player.go(p, how);
    const vis = visibleIndices(deck);
    counter.textContent = `${Math.max(1, vis.indexOf(p.i) + 1)} / ${vis.length}`;
    notesBox.textContent = deck.slides[p.i].notes || 'No speaker notes for this slide.';
    remote?.send?.(p);
  }
  const nav = remote?.nav || null;
  function next() {
    if (nav) return nav('next');
    if (ended) return exit();
    const n = nextPos(deck, pos);
    if (n) show(n);
    else {
      ended = true;
      endScreen.hidden = false;
    }
  }
  function prev() {
    if (nav) return nav('prev');
    if (ended) {
      ended = false;
      endScreen.hidden = true;
      return;
    }
    show(prevPos(deck, pos), 'cut');
  }
  function jump(p) {
    if (nav) nav('goto', p);
    else show(p, 'cut');
  }
  function toggleNotes() { notesBox.hidden = !notesBox.hidden; }
  function toggleFs() {
    if (document.fullscreenElement) {
      delete root.dataset.wasFs; // leaving full screen on purpose doesn't end the show
      document.exitFullscreen?.();
    } else enterFullscreen(root);
  }
  /** Black or white out the screen (again to bring the slide back; no colour: back). */
  function blank(color) {
    if (!color || (!curtain.hidden && curtain.dataset.color === color)) curtain.hidden = true;
    else {
      curtain.hidden = false;
      curtain.dataset.color = color;
      curtain.style.background = color;
    }
  }
  function exit() {
    if (!ctl.active) return;
    ctl.active = false;
    document.removeEventListener('keydown', onKey, true);
    document.removeEventListener('fullscreenchange', onFsChange);
    clearTimeout(idle);
    player.destroy();
    root.remove();
    if (document.fullscreenElement) document.exitFullscreen?.().catch(() => {});
    remote?.end?.();
    onExit?.(pos?.i ?? null);
  }

  function onKey(e) {
    if (!ctl.active) return;
    const k = e.key;
    e.stopPropagation();
    // Ctrl, ⌘ and Alt combinations are the browser's (new window, print, back…), not slideshow keys.
    if (e.ctrlKey || e.metaKey || e.altKey) return;
    if (ended && k !== 'Escape' && !['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p', 'P'].includes(k)) {
      e.preventDefault();
      exit();
      return;
    }
    if (!curtain.hidden && !['b', 'B', '.', 'w', 'W', ',', 'Escape'].includes(k)) curtain.hidden = true;
    if (/^\d$/.test(k)) {
      typed += k;
      return;
    }
    if (k === 'Enter' && typed) {
      e.preventDefault();
      const vis = visibleIndices(deck);
      const i = vis[Math.min(vis.length, Math.max(1, +typed)) - 1];
      typed = '';
      if (i != null) jump({ i, step: minStep(deck.slides[i]) });
      return;
    }
    typed = '';
    const run = (fn) => { e.preventDefault(); fn(); };
    if (['ArrowRight', 'ArrowDown', ' ', 'Enter', 'PageDown', 'n', 'N'].includes(k)) run(next);
    else if (['ArrowLeft', 'ArrowUp', 'Backspace', 'PageUp', 'p', 'P'].includes(k)) run(prev);
    else if (k === 'Home') run(() => jump(firstPos(deck, 0)));
    else if (k === 'End') run(() => { const vis = visibleIndices(deck); jump({ i: vis[vis.length - 1], step: buildSteps(deck.slides[vis[vis.length - 1]]).length }); });
    else if (k === 'b' || k === 'B' || k === '.') run(() => blank('#000'));
    else if (k === 'w' || k === 'W' || k === ',') run(() => blank('#fff'));
    else if (k === 's' || k === 'S') run(toggleNotes);
    else if (k === 'f' || k === 'F') run(toggleFs);
    else if (k === 'Escape') run(exit);
  }
  const onFsChange = () => {
    // Leaving full screen with Esc ends the show, as in Slides and PowerPoint.
    if (!document.fullscreenElement && root.dataset.wasFs && !remote) exit();
    if (document.fullscreenElement) root.dataset.wasFs = '1';
  };
  document.addEventListener('keydown', onKey, true);
  document.addEventListener('fullscreenchange', onFsChange);
  // With the screen blacked or whited out, the first click only brings the slide back.
  const uncover = () => {
    if (curtain.hidden) return false;
    curtain.hidden = true;
    return true;
  };
  root.addEventListener('click', (e) => {
    // Links: to another slide, or a web page (in a new tab; links in text open there by themselves).
    const link = e.target.closest('a[href], [data-link]');
    if (link) {
      const href = link.getAttribute('href') || link.dataset.link;
      const to = /^#slide:/.test(href) && linkTarget(deck, href.slice(7), pos);
      if (/^#/.test(href)) e.preventDefault();
      if (to === 'next') next();
      else if (to === 'prev') prev();
      else if (to) jump(to);
      else if (!link.matches('a') && /^(https?:|mailto:)/i.test(href)) window.open(href, '_blank', 'noopener');
      return;
    }
    if (e.target.closest('.show-controls, .show-notes')) return;
    // Clicks on a video's controls or a sound's button play and pause it, not the slideshow.
    if (e.target.closest('video, audio, .el-media-toggle')) return;
    if (!uncover()) next();
  });
  root.addEventListener('contextmenu', (e) => {
    e.preventDefault();
    if (!uncover()) prev();
  });
  const showPointer = () => {
    root.classList.add('show-pointer');
    clearTimeout(idle);
    idle = setTimeout(() => root.classList.remove('show-pointer'), 2200);
  };
  root.addEventListener('pointermove', showPointer);
  // Touch screens have no pointer moving about: a tap shows the controls too.
  root.addEventListener('pointerdown', (e) => {
    if (e.pointerType === 'touch') showPointer();
  });

  show(pos, nav ? 'cut' : 'auto');
  return ctl;
}

// ------------------------------------------------------------------ presenter view

const fmtTime = (ms) => {
  const s = Math.floor(ms / 1000);
  const hh = Math.floor(s / 3600);
  const mm = Math.floor((s % 3600) / 60);
  const ss = s % 60;
  return `${hh ? `${hh}:` : ''}${String(mm).padStart(hh ? 2 : 1, '0')}:${String(ss).padStart(2, '0')}`;
};

/**
 * Presenter view: the audience sees the slideshow in a second window (move
 * it to the projector and make it full screen); this window shows the
 * current and next slide, speaker notes and a timer.
 */
export function startPresenterView(deck, { start = 0, onExit = null } = {}) {
  const channel = new BroadcastChannel(`libreslides-show-${deck.id}`);
  const url = `${location.pathname}${location.search}#/show/${encodeURIComponent(deck.id)}`;
  let audience = null;
  const openAudience = () => {
    audience = window.open(url, `libreslides-audience-${deck.id}`, 'popup,width=1024,height=600');
    if (!audience) toast('Your browser blocked the audience window. Allow pop-ups for this site to present on a second screen.', { type: 'error', timeout: 8000 });
  };
  openAudience();

  let pos = firstPos(deck, start);
  const currentHost = h('div', { class: 'pv-current' });
  const nextHost = h('div', { class: 'pv-next-slide' });
  const player = new Player(deck, currentHost, { animate: false, media: false });
  const notes = h('div', { class: 'pv-notes', tabindex: '0', 'aria-label': 'Speaker notes' });
  let notesSize = 22;
  const timerEl = h('span', { class: 'pv-timer' }, '0:00');
  const clockEl = h('span', { class: 'pv-clock' });
  const counter = h('span', { class: 'pv-count' });
  const titleEl = h('span', { class: 'pv-title' });
  let started = Date.now();
  let pausedAt = null;
  const pauseBtn = h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Pause timer', html: icon('pause') });
  pauseBtn.addEventListener('click', () => {
    if (pausedAt) {
      started += Date.now() - pausedAt;
      pausedAt = null;
      pauseBtn.innerHTML = icon('pause');
      pauseBtn.setAttribute('aria-label', 'Pause timer');
    } else {
      pausedAt = Date.now();
      pauseBtn.innerHTML = icon('play');
      pauseBtn.setAttribute('aria-label', 'Resume timer');
    }
  });
  const tick = setInterval(() => {
    timerEl.textContent = fmtTime((pausedAt || Date.now()) - started);
    clockEl.textContent = new Date().toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });
  }, 500);

  const root = h('div', { class: 'presenter', tabindex: '-1', role: 'application', 'aria-label': 'Presenter view' },
    h('header', { class: 'pv-bar' },
      titleEl,
      h('div', { class: 'pv-time' }, timerEl, pauseBtn,
        h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Restart timer', html: icon('rotateCcw'), onclick: () => { started = Date.now(); if (pausedAt) pausedAt = Date.now(); } }),
        clockEl),
      h('div', { class: 'pv-actions' },
        // Opens it again if it was closed; it asks where the show is when it starts.
        h('button', { type: 'button', class: 'btn', onclick: () => (!audience || audience.closed ? openAudience() : audience.focus?.()) }, 'Show audience window'),
        h('button', { type: 'button', class: 'btn btn-primary', onclick: () => exit() }, 'End show'))),
    h('div', { class: 'pv-main' },
      h('div', { class: 'pv-left' }, currentHost,
        h('div', { class: 'pv-nav' },
          h('button', { type: 'button', class: 'btn', html: `${icon('chevronLeft')}<span>Previous</span>`, onclick: () => go(prevPos(deck, pos), 'cut') }),
          counter,
          h('button', { type: 'button', class: 'btn btn-primary', html: `<span>Next</span>${icon('chevronRight')}`, onclick: () => go(nextPos(deck, pos)) }))),
      h('div', { class: 'pv-right' },
        h('div', { class: 'pv-label' }, 'Next'), nextHost,
        h('div', { class: 'pv-label pv-notes-head' }, 'Speaker notes',
          h('span', {},
            h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Smaller notes', html: icon('minus'), onclick: () => { notesSize = Math.max(12, notesSize - 2); notes.style.fontSize = `${notesSize}px`; } }),
            h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Larger notes', html: icon('plus'), onclick: () => { notesSize = Math.min(48, notesSize + 2); notes.style.fontSize = `${notesSize}px`; } }))),
        notes)));
  document.body.append(root);
  root.focus();

  function go(p, how = 'auto') {
    if (!p) return;
    pos = p;
    player.go(p, 'cut');
    const vis = visibleIndices(deck);
    counter.textContent = `Slide ${vis.indexOf(p.i) + 1} of ${vis.length}`;
    titleEl.textContent = slideTitle(deck.slides[p.i]) || deck.title;
    notes.textContent = deck.slides[p.i].notes || 'No speaker notes for this slide.';
    notes.classList.toggle('is-empty', !deck.slides[p.i].notes);
    const n = nextPos(deck, { i: p.i, step: buildSteps(deck.slides[p.i]).length });
    nextHost.replaceChildren(p.step < buildSteps(deck.slides[p.i]).length
      ? h('div', { class: 'pv-next-text' }, 'Next click: animation on this slide')
      : n ? renderThumb(deck.slides[n.i], deck, Math.min(360, nextHost.clientWidth || 320)) : h('div', { class: 'pv-next-text' }, 'End of slideshow'));
    channel.postMessage({ type: 'state', pos: p, how });
  }

  function exit() {
    if (!ctl.active) return;
    ctl.active = false;
    clearInterval(tick);
    channel.postMessage({ type: 'end' });
    channel.close();
    document.removeEventListener('keydown', onKey, true);
    window.removeEventListener('pagehide', onPageHide);
    player.destroy();
    root.remove();
    onExit?.(pos.i);
  }

  channel.onmessage = (e) => {
    const m = e.data;
    if (m.type === 'hello') channel.postMessage({ type: 'state', pos, how: 'cut' });
    else if (m.type === 'nav') {
      if (m.action === 'next') go(nextPos(deck, pos));
      else if (m.action === 'prev') go(prevPos(deck, pos), 'cut');
      else if (m.action === 'goto' && m.pos) go(m.pos, 'cut');
      else if (m.action === 'exit') exit();
    }
  };

  // The same keys as the slideshow (see startSlideshow).
  let typed = '';
  function onKey(e) {
    if (!ctl.active || e.target.closest?.('input, textarea')) return;
    const k = e.key;
    e.stopPropagation();
    if (e.ctrlKey || e.metaKey || e.altKey) return;
    // Space and Enter press the focused button (pause, notes size…) instead.
    if ((k === ' ' || k === 'Enter') && e.target.closest?.('button')) return;
    if (/^\d$/.test(k)) {
      typed += k;
      return;
    }
    const run = (fn) => { e.preventDefault(); fn(); };
    if (k === 'Enter' && typed) {
      const vis = visibleIndices(deck);
      const i = vis[Math.min(vis.length, Math.max(1, +typed)) - 1];
      typed = '';
      if (i != null) run(() => go({ i, step: minStep(deck.slides[i]) }, 'cut'));
      return;
    }
    typed = '';
    if (['ArrowRight', 'ArrowDown', ' ', 'Enter', 'PageDown', 'n', 'N'].includes(k)) run(() => go(nextPos(deck, pos)));
    else if (['ArrowLeft', 'ArrowUp', 'Backspace', 'PageUp', 'p', 'P'].includes(k)) run(() => go(prevPos(deck, pos), 'cut'));
    else if (k === 'Home') run(() => go(firstPos(deck, 0), 'cut'));
    else if (k === 'End') run(() => { const vis = visibleIndices(deck); go({ i: vis[vis.length - 1], step: buildSteps(deck.slides[vis[vis.length - 1]]).length }, 'cut'); });
    else if (k === 'b' || k === 'B' || k === '.') run(() => channel.postMessage({ type: 'blank', color: '#000' }));
    else if (k === 'w' || k === 'W' || k === ',') run(() => channel.postMessage({ type: 'blank', color: '#fff' }));
    else if (k === 'Escape') run(exit);
  }
  // Closing or reloading this window ends the show in the audience window too.
  const onPageHide = () => channel.postMessage({ type: 'end' });
  document.addEventListener('keydown', onKey, true);
  window.addEventListener('pagehide', onPageHide);
  const ctl = { active: true, exit };
  go(pos, 'cut');
  return ctl;
}

/**
 * The audience window of the presenter view: shows whatever the console
 * says, and sends clicks and keys back to it.
 */
export function runAudience(deck) {
  const channel = new BroadcastChannel(`libreslides-show-${deck.id}`);
  const send = (msg) => channel.postMessage(msg);
  const ctl = startSlideshow(deck, {
    start: 0,
    remote: {
      // Every move here goes through the console, which then tells us where to go.
      nav: (action, pos = null) => send({ type: 'nav', action, pos }),
      end: () => { send({ type: 'nav', action: 'exit' }); channel.close(); window.close(); },
    },
    onExit: () => {},
  });
  const showEl = document.querySelector('.show');
  const hint = h('button', { type: 'button', class: 'show-fs-hint', onclick: (e) => { e.stopPropagation(); showEl.requestFullscreen?.().catch(() => {}); hint.remove(); } }, 'Click here to go full screen');
  showEl.append(hint);
  setTimeout(() => hint.remove(), 8000);
  channel.onmessage = (e) => {
    const m = e.data;
    if (m.type === 'state') {
      ctl.blank(null); // moving on brings the slide back, as in the slideshow
      ctl.go(m.pos, m.how);
    } else if (m.type === 'blank') ctl.blank(m.color);
    else if (m.type === 'end') {
      channel.close();
      window.close();
    }
  };
  send({ type: 'hello' });
  return ctl;
}

// ------------------------------------------------------------------ previews in the editor

/** Play the current slide's transition or animations over the editor's slide. */
export function previewOnStage(app, { animations = false, transition = false } = {}) {
  const stage = app.canvas.stage;
  const host = h('div', { class: 'preview-host', title: 'Click to stop the preview' });
  stage.append(host);
  const player = new Player(app.deck, host, { media: false });
  const i = app.slideIndex;
  const done = () => {
    clearTimeout(timer);
    player.destroy();
    host.remove();
  };
  host.addEventListener('pointerdown', (e) => { e.stopPropagation(); done(); });
  let timer;
  if (transition) {
    // From the previous slide (or black) into this one.
    if (i > 0) player.go({ i: i - 1, step: buildSteps(app.deck.slides[i - 1]).length }, 'cut');
    else player.pos = { i: -1, step: 0 };
    requestAnimationFrame(() => player.go({ i, step: minStep(app.slide) }));
    timer = setTimeout(done, ((app.slide.transition?.duration ?? 0.6) + 0.6) * 1000);
  } else if (animations) {
    const steps = buildSteps(app.slide);
    player.go({ i, step: 0 }, 'cut');
    let k = 0;
    const play = () => {
      if (k >= steps.length) {
        timer = setTimeout(done, 900);
        return;
      }
      player.go({ i, step: k + 1 });
      const longest = Math.max(...steps[k].map((s) => s.delay + s.duration));
      k++;
      timer = setTimeout(play, longest * 1000 + 450);
    };
    timer = setTimeout(play, 350);
  }
}
