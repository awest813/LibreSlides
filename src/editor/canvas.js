import { h } from '../ui/dom.js';
import { renderSlide, renderElement, fillFields, slideNumber } from '../render/render.js';
import { slideSize, boxOf, moveElement, canHoldText, newTextBox, newShape, newLine, groupOf } from '../model/deck.js';
import { themeById } from '../model/themes.js';
import { sanitizeHtml } from '../io/sanitize.js';
import { tidyHtml, normalizeFontTags, exec } from './richtext.js';
import { drawCommentMarkers } from './comments.js';

/**
 * The editing surface: the current slide at the zoom level, with a selection
 * overlay (handles, snapping guides, marquee) drawn on top in screen pixels
 * so handles keep their size at any zoom.
 */
const HANDLES = [
  ['nw', -1, -1], ['n', 0, -1], ['ne', 1, -1], ['e', 1, 0],
  ['se', 1, 1], ['s', 0, 1], ['sw', -1, 1], ['w', -1, 0],
];
const SNAP_PX = 6;
const PAD = 40; // space around the slide at "fit" zoom

const rad = (deg) => (deg * Math.PI) / 180;
const rotateVec = (x, y, deg) => {
  const c = Math.cos(rad(deg));
  const s = Math.sin(rad(deg));
  return { x: x * c - y * s, y: x * s + y * c };
};

/** Axis-aligned bounds of a (possibly rotated) element. */
export function aabb(el) {
  const b = boxOf(el);
  if (!el.rot || el.type === 'line') return b;
  const cx = b.x + b.w / 2;
  const cy = b.y + b.h / 2;
  const pts = [[-1, -1], [1, -1], [1, 1], [-1, 1]].map(([sx, sy]) => rotateVec((sx * b.w) / 2, (sy * b.h) / 2, el.rot));
  const xs = pts.map((p) => cx + p.x);
  const ys = pts.map((p) => cy + p.y);
  const x = Math.min(...xs);
  const y = Math.min(...ys);
  return { x, y, w: Math.max(...xs) - x, h: Math.max(...ys) - y };
}

const unionBox = (boxes) => {
  const x = Math.min(...boxes.map((b) => b.x));
  const y = Math.min(...boxes.map((b) => b.y));
  return { x, y, w: Math.max(...boxes.map((b) => b.x + b.w)) - x, h: Math.max(...boxes.map((b) => b.y + b.h)) - y };
};

function placeCaretAt(x, y) {
  let range = null;
  if (document.caretPositionFromPoint) {
    const p = document.caretPositionFromPoint(x, y);
    if (p) {
      range = document.createRange();
      range.setStart(p.offsetNode, p.offset);
    }
  } else if (document.caretRangeFromPoint) range = document.caretRangeFromPoint(x, y);
  if (!range) return false;
  range.collapse(true);
  const sel = document.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  return true;
}

function caretToEnd(body, selectAll = false) {
  const range = document.createRange();
  range.selectNodeContents(body);
  if (!selectAll) range.collapse(false);
  const sel = document.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
}

export class SlideCanvas {
  constructor(app) {
    this.app = app;
    this.zoom = 'fit';
    this.scale = 1;
    this.tool = null; // { type: 'text' | 'shape' | 'line', shape?, arrow? }
    this.editing = null; // { id, original, before, bodies }
    this.snapEnabled = true;
    this.guides = [];
    this.marquee = null;
    this.draft = null;

    this.slideEl = null;
    this.overlay = h('div', { class: 'overlay', 'aria-hidden': 'true' });
    this.stage = h('div', { class: 'stage' }, this.overlay);
    this.wrap = h('div', { class: 'canvas-wrap', tabindex: '0', role: 'region', 'aria-label': 'Slide canvas. Tab selects objects; arrow keys move them.' }, this.stage);

    this.stage.addEventListener('pointerdown', (e) => this.onPointerDown(e));
    this.stage.addEventListener('dblclick', (e) => this.onDblClick(e));
    // Links in text are followed in the slideshow, not while editing.
    this.stage.addEventListener('click', (e) => { if (e.target.closest?.('a[href]')) e.preventDefault(); });
    this.stage.addEventListener('contextmenu', (e) => this.onContextMenu(e));
    this.wrap.addEventListener('pointerdown', (e) => {
      // The grey area around the slide: deselect, or draw there with a tool.
      if (e.target === this.wrap) this.onPointerDown(e);
    });
    this.wrap.addEventListener('wheel', (e) => {
      if (!(e.ctrlKey || e.metaKey)) return;
      e.preventDefault();
      this.setZoom(Math.max(0.1, Math.min(4, this.scale * (e.deltaY < 0 ? 1.1 : 1 / 1.1))));
    }, { passive: false });
    this.resizeObserver = new ResizeObserver(() => this.layout());
    this.resizeObserver.observe(this.wrap);
    this.onSelectionChange = () => {
      if (!this.editing) return;
      const sel = document.getSelection();
      if (sel.rangeCount && this.editing.bodies.some((b) => b.contains(sel.anchorNode))) {
        this.savedRange = sel.getRangeAt(0).cloneRange();
        this.app.onTextSelection();
      }
    };
    document.addEventListener('selectionchange', this.onSelectionChange);
  }

  destroy() {
    this.resizeObserver.disconnect();
    document.removeEventListener('selectionchange', this.onSelectionChange);
  }

  get deck() { return this.app.deck; }
  get slide() { return this.app.slide; }
  get size() { return slideSize(this.deck); }
  find(id) { return this.slide.elements.find((e) => e.id === id); }
  nodeOf(id) { return this.slideEl?.querySelector(`:scope > .el[data-id="${CSS.escape(id)}"]`); }

  // ---------------------------------------------------------------- layout & rendering

  setZoom(z) {
    this.zoom = z;
    this.layout();
    this.app.onZoom?.();
  }

  fitScale() {
    const { w, h: H } = this.size;
    const availW = Math.max(100, this.wrap.clientWidth - PAD * 2);
    const availH = Math.max(80, this.wrap.clientHeight - PAD * 2);
    return Math.min(availW / w, availH / H);
  }

  layout() {
    const { w, h: H } = this.size;
    this.scale = this.zoom === 'fit' ? this.fitScale() : this.zoom;
    this.stage.style.width = `${Math.round(w * this.scale)}px`;
    this.stage.style.height = `${Math.round(H * this.scale)}px`;
    this.wrap.classList.toggle('is-zoomed', this.zoom !== 'fit' && this.scale > this.fitScale());
    if (this.slideEl) this.slideEl.style.transform = `scale(${this.scale})`;
    this.drawOverlay();
  }

  /**
   * Draw the slide. The element being edited keeps its DOM node (so the caret
   * and focus survive); everything else is redrawn from the model.
   */
  render() {
    const fresh = renderSlide(this.slide, this.deck, { editing: true });
    fresh.style.transform = `scale(${this.scale})`;
    if (!this.slideEl || !this.editing) {
      if (this.slideEl) this.slideEl.replaceWith(fresh);
      else this.stage.prepend(fresh);
      this.slideEl = fresh;
    } else {
      const live = this.nodeOf(this.editing.id);
      this.slideEl.className = fresh.className;
      this.slideEl.style.cssText = fresh.style.cssText;
      for (const child of [...this.slideEl.children]) if (child !== live) child.remove();
      let before = true;
      for (const child of [...fresh.children]) {
        if (live && child.dataset.id === live.dataset.id) {
          this.patchLive(live, child);
          before = false;
        } else if (before && live) this.slideEl.insertBefore(child, live);
        else this.slideEl.append(child);
      }
    }
    this.layout();
  }

  /** Update the edited element's box and shape without touching its text. */
  patchLive(live, fresh) {
    live.style.cssText = fresh.style.cssText;
    live.querySelector(':scope > svg')?.remove();
    const svg = fresh.querySelector(':scope > svg');
    if (svg) live.prepend(svg);
    const box = live.querySelector(':scope > .el-textbox');
    const fbox = fresh.querySelector(':scope > .el-textbox');
    if (box && fbox) {
      box.style.cssText = fbox.style.cssText;
      box.firstElementChild?.classList.toggle('has-bullets', fbox.firstElementChild?.classList.contains('has-bullets'));
    }
    const table = live.querySelector(':scope > table');
    const ftable = fresh.querySelector(':scope > table');
    if (table && ftable) {
      table.style.cssText = ftable.style.cssText;
      table.querySelector('colgroup')?.replaceWith(ftable.querySelector('colgroup'));
      const fcells = ftable.querySelectorAll('td');
      table.querySelectorAll('td').forEach((td, i) => {
        if (fcells[i]) td.style.cssText = fcells[i].style.cssText;
      });
    }
  }

  /** Redraw one element (after moving or resizing it). */
  refreshNode(el) {
    const node = this.nodeOf(el.id);
    if (!node) return;
    const fresh = renderElement(el, { theme: themeById(this.deck.themeId), deck: this.deck, editing: true, slideNo: slideNumber(this.slide, this.deck) });
    fillFields(fresh, this.slide, this.deck);
    if (this.editing?.id === el.id) this.patchLive(node, fresh);
    else node.replaceWith(fresh);
  }

  // ---------------------------------------------------------------- overlay

  drawOverlay() {
    const s = this.scale;
    const ov = this.overlay;
    ov.replaceChildren();
    const selected = this.app.selectedElements();
    const single = selected.length === 1 ? selected[0] : null;
    for (const el of selected) {
      if (el.type === 'line') {
        const line = h('div', { class: 'sel-line' });
        const len = Math.hypot(el.x2 - el.x1, el.y2 - el.y1) * s;
        const ang = (Math.atan2(el.y2 - el.y1, el.x2 - el.x1) * 180) / Math.PI;
        Object.assign(line.style, { left: `${el.x1 * s}px`, top: `${el.y1 * s}px`, width: `${len}px`, transform: `rotate(${ang}deg)` });
        ov.append(line);
        if (single) {
          for (const end of [1, 2]) {
            ov.append(h('div', { class: 'handle handle-end', dataset: { handle: `end${end}` }, style: { left: `${el[`x${end}`] * s}px`, top: `${el[`y${end}`] * s}px` } }));
          }
        }
        continue;
      }
      const box = h('div', { class: `sel-box${this.editing?.id === el.id ? ' is-editing' : ''}${el.locked ? ' is-locked' : ''}` });
      Object.assign(box.style, { left: `${el.x * s}px`, top: `${el.y * s}px`, width: `${el.w * s}px`, height: `${el.h * s}px`, transform: el.rot ? `rotate(${el.rot}deg)` : '' });
      if (single && !el.locked) {
        for (const [name] of HANDLES) box.append(h('div', { class: `handle handle-${name}`, dataset: { handle: name } }));
        box.append(h('div', { class: 'handle handle-rotate', dataset: { handle: 'rotate' }, title: 'Rotate' }));
      }
      ov.append(box);
    }
    if (selected.length > 1) {
      const u = unionBox(selected.map(aabb));
      ov.append(h('div', { class: 'sel-group', style: { left: `${u.x * s}px`, top: `${u.y * s}px`, width: `${u.w * s}px`, height: `${u.h * s}px` } }));
    }
    for (const g of this.guides) {
      ov.append(h('div', { class: `guide guide-${g.axis}`, style: g.axis === 'x' ? { left: `${g.at * s}px` } : { top: `${g.at * s}px` } }));
    }
    for (const r of [this.marquee, this.draft]) {
      if (!r) continue;
      if (r.line) {
        const len = Math.hypot(r.x2 - r.x1, r.y2 - r.y1) * s;
        const ang = (Math.atan2(r.y2 - r.y1, r.x2 - r.x1) * 180) / Math.PI;
        ov.append(h('div', { class: 'sel-line draft', style: { left: `${r.x1 * s}px`, top: `${r.y1 * s}px`, width: `${len}px`, transform: `rotate(${ang}deg)` } }));
      } else {
        ov.append(h('div', { class: r === this.marquee ? 'marquee' : 'draft-box', style: { left: `${r.x * s}px`, top: `${r.y * s}px`, width: `${r.w * s}px`, height: `${r.h * s}px` } }));
      }
    }
    drawCommentMarkers(this, ov);
  }

  // ---------------------------------------------------------------- pointer input

  toSlide(e) {
    const r = this.stage.getBoundingClientRect();
    return { x: (e.clientX - r.left) / this.scale, y: (e.clientY - r.top) / this.scale };
  }

  /** Follow the pointer until it's released. */
  track(e, { move, up }) {
    const start = { x: e.clientX, y: e.clientY };
    let moved = false;
    const onMove = (ev) => {
      if (!moved && Math.hypot(ev.clientX - start.x, ev.clientY - start.y) < 3) return;
      moved = true;
      move?.(ev);
    };
    const onUp = (ev) => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      window.removeEventListener('pointercancel', onUp);
      up?.(ev, moved);
    };
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
    window.addEventListener('pointercancel', onUp);
  }

  onPointerDown(e) {
    if (e.button !== 0 && e.button !== 2) return;
    const handle = e.target.closest?.('[data-handle]');
    if (handle && e.button === 0) {
      e.preventDefault();
      this.startHandleDrag(e, handle.dataset.handle);
      return;
    }
    if (this.tool && e.button === 0) {
      e.preventDefault();
      this.endEdit();
      this.startDraw(e);
      return;
    }
    const node = e.target.closest?.('.slide > .el');
    const id = node?.dataset.id;
    if (this.editing) {
      if (id === this.editing.id) return; // clicks inside the text being edited place the caret
      this.endEdit();
    }
    this.wrap.focus({ preventScroll: true });
    if (!id) {
      if (e.button !== 0) return;
      e.preventDefault();
      if (!e.shiftKey) this.app.setSelection([]);
      this.startMarquee(e);
      return;
    }
    const el = this.find(id);
    const additive = e.shiftKey || e.ctrlKey || e.metaKey;
    const wasSelected = this.app.selection.includes(id);
    // A click selects the whole group an object is in; a second click, just that object.
    const group = groupOf(this.slide, el).map((m) => m.id);
    const groupSelected = group.length > 1 && this.app.selection.length === group.length && group.every((m) => this.app.selection.includes(m));
    if (additive) {
      e.preventDefault();
      this.app.setSelection(wasSelected ? this.app.selection.filter((x) => !group.includes(x)) : [...new Set([...this.app.selection, ...group])]);
      return;
    }
    if (!wasSelected) this.app.setSelection(group);
    if (e.button !== 0) return;
    e.preventDefault();
    if (el.locked) return;
    this.startMove(e, () => {
      if (groupSelected) {
        this.app.setSelection([id]);
        return;
      }
      // A click (no drag) on a text box that was already selected starts typing in it.
      if (wasSelected && this.app.selection.length === 1 && (canHoldText(el) || el.type === 'table')) this.startEdit(id, { point: e });
    });
  }

  onDblClick(e) {
    const node = e.target.closest?.('.slide > .el');
    const el = node && this.find(node.dataset.id);
    if (!el || el.locked) return;
    if (canHoldText(el) || el.type === 'table') {
      if (this.editing?.id !== el.id) this.startEdit(el.id, { point: e });
    } else if (el.type === 'image') this.app.replaceImage?.();
    else if (el.type === 'chart') this.app.openPanel('format');
    else if (el.type === 'diagram') this.app.editDiagramText();
  }

  onContextMenu(e) {
    if (this.editing && e.target.closest?.('[contenteditable="true"]')) return; // the browser's menu (spelling…)
    e.preventDefault();
    const node = e.target.closest?.('.slide > .el');
    if (node && !this.app.selection.includes(node.dataset.id)) this.app.setSelection([node.dataset.id]);
    if (!node) this.app.setSelection([]);
    const td = e.target.closest?.('td[data-r]');
    this.app.canvasMenu({ x: e.clientX, y: e.clientY }, td ? { r: +td.dataset.r, c: +td.dataset.c } : null);
  }

  snapCandidates(exclude) {
    const { w, h: H } = this.size;
    const xs = [0, w / 2, w];
    const ys = [0, H / 2, H];
    for (const el of this.slide.elements) {
      if (exclude.has(el.id)) continue;
      const b = aabb(el);
      xs.push(b.x, b.x + b.w / 2, b.x + b.w);
      ys.push(b.y, b.y + b.h / 2, b.y + b.h);
    }
    return { xs, ys };
  }

  /** Smallest correction that lines one of `edges` up with a candidate. */
  snapAxis(edges, cands) {
    const tol = SNAP_PX / this.scale;
    let best = null;
    for (const e of edges) {
      for (const c of cands) {
        const d = c - e;
        if (Math.abs(d) <= tol && (!best || Math.abs(d) < Math.abs(best.d))) best = { d, at: c };
      }
    }
    return best;
  }

  startMove(e, onClick) {
    const els = this.app.selectedElements().filter((el) => !el.locked);
    if (!els.length) return;
    const p0 = this.toSlide(e);
    const before = this.app.snapshot();
    const orig = els.map((el) => JSON.parse(JSON.stringify(el)));
    const box0 = unionBox(els.map(aabb));
    const cands = this.snapCandidates(new Set(els.map((x) => x.id)));
    this.track(e, {
      move: (ev) => {
        const p = this.toSlide(ev);
        let dx = p.x - p0.x;
        let dy = p.y - p0.y;
        if (ev.shiftKey) {
          if (Math.abs(dx) > Math.abs(dy)) dy = 0;
          else dx = 0;
        }
        this.guides = [];
        if (this.snapEnabled && !ev.altKey) {
          const bx = box0.x + dx;
          const by = box0.y + dy;
          const sx = this.snapAxis([bx, bx + box0.w / 2, bx + box0.w], cands.xs);
          const sy = this.snapAxis([by, by + box0.h / 2, by + box0.h], cands.ys);
          if (sx && !(ev.shiftKey && dx === 0)) { dx += sx.d; this.guides.push({ axis: 'x', at: sx.at }); }
          if (sy && !(ev.shiftKey && dy === 0)) { dy += sy.d; this.guides.push({ axis: 'y', at: sy.at }); }
        }
        els.forEach((el, i) => {
          Object.assign(el, JSON.parse(JSON.stringify(orig[i])));
          moveElement(el, Math.round(dx * 10) / 10, Math.round(dy * 10) / 10);
          const node = this.nodeOf(el.id);
          if (el.type === 'line') this.refreshNode(el);
          else if (node) {
            node.style.left = `${el.x}px`;
            node.style.top = `${el.y}px`;
          }
        });
        this.drawOverlay();
      },
      up: (_ev, moved) => {
        this.guides = [];
        if (moved) this.app.commit(before);
        else onClick?.();
        this.drawOverlay();
      },
    });
  }

  startHandleDrag(e, handle) {
    const el = this.app.selectedElements()[0];
    if (!el) return;
    const before = this.app.snapshot();
    const orig = JSON.parse(JSON.stringify(el));
    const cands = this.snapCandidates(new Set([el.id]));

    if (handle.startsWith('end')) {
      const n = handle.slice(3);
      const other = n === '1' ? '2' : '1';
      this.track(e, {
        move: (ev) => {
          let p = this.toSlide(ev);
          if (ev.shiftKey) {
            // Lock to 45° steps.
            const ox = el[`x${other}`];
            const oy = el[`y${other}`];
            const len = Math.hypot(p.x - ox, p.y - oy);
            const ang = Math.round(Math.atan2(p.y - oy, p.x - ox) / (Math.PI / 4)) * (Math.PI / 4);
            p = { x: ox + len * Math.cos(ang), y: oy + len * Math.sin(ang) };
          } else if (this.snapEnabled && !ev.altKey) {
            const sx = this.snapAxis([p.x], cands.xs);
            const sy = this.snapAxis([p.y], cands.ys);
            if (sx) p.x += sx.d;
            if (sy) p.y += sy.d;
          }
          el[`x${n}`] = Math.round(p.x * 10) / 10;
          el[`y${n}`] = Math.round(p.y * 10) / 10;
          this.refreshNode(el);
          this.drawOverlay();
        },
        up: (_ev, moved) => moved && this.app.commit(before),
      });
      return;
    }

    if (handle === 'rotate') {
      const cx = orig.x + orig.w / 2;
      const cy = orig.y + orig.h / 2;
      this.track(e, {
        move: (ev) => {
          const p = this.toSlide(ev);
          let ang = (Math.atan2(p.y - cy, p.x - cx) * 180) / Math.PI + 90;
          if (ev.shiftKey) ang = Math.round(ang / 15) * 15;
          else for (const t of [0, 90, 180, 270, 360]) if (Math.abs(ang - t) < 3) ang = t;
          ang = ((Math.round(ang) % 360) + 360) % 360;
          el.rot = ang;
          const node = this.nodeOf(el.id);
          if (node) node.style.transform = ang ? `rotate(${ang}deg)` : '';
          this.drawOverlay();
        },
        up: (_ev, moved) => moved && this.app.commit(before),
      });
      return;
    }

    const [, sx, sy] = HANDLES.find(([n]) => n === handle);
    const p0 = this.toSlide(e);
    const corner = sx !== 0 && sy !== 0;
    // Pictures and videos keep their proportions from the corners unless Shift is held; other shapes only with Shift.
    const lockDefault = el.type === 'image' || el.type === 'media' || !!el.lockRatio;
    this.track(e, {
      move: (ev) => {
        const p = this.toSlide(ev);
        const d = rotateVec(p.x - p0.x, p.y - p0.y, -(orig.rot || 0));
        let w = orig.w + sx * d.x;
        let hh = orig.h + sy * d.y;
        const lock = corner && (lockDefault ? !ev.shiftKey : ev.shiftKey);
        if (lock) {
          const k = Math.max(w / orig.w, hh / orig.h);
          w = orig.w * k;
          hh = orig.h * k;
        }
        w = Math.max(8, w);
        hh = Math.max(8, hh);
        // Keep the opposite side (or corner) where it was.
        const rot = orig.rot || 0;
        const anchor = rotateVec((-sx * orig.w) / 2, (-sy * orig.h) / 2, rot);
        const ax = orig.x + orig.w / 2 + anchor.x;
        const ay = orig.y + orig.h / 2 + anchor.y;
        const back = rotateVec((-sx * w) / 2, (-sy * hh) / 2, rot);
        let x = ax - back.x - w / 2;
        let y = ay - back.y - hh / 2;
        this.guides = [];
        if (!rot && this.snapEnabled && !ev.altKey && !lock) {
          if (sx) {
            const edge = sx > 0 ? x + w : x;
            const s = this.snapAxis([edge], cands.xs);
            if (s) {
              if (sx > 0) w += s.d;
              else { x += s.d; w -= s.d; }
              this.guides.push({ axis: 'x', at: s.at });
            }
          }
          if (sy) {
            const edge = sy > 0 ? y + hh : y;
            const s = this.snapAxis([edge], cands.ys);
            if (s) {
              if (sy > 0) hh += s.d;
              else { y += s.d; hh -= s.d; }
              this.guides.push({ axis: 'y', at: s.at });
            }
          }
        }
        Object.assign(el, { x: Math.round(x * 10) / 10, y: Math.round(y * 10) / 10, w: Math.round(w * 10) / 10, h: Math.round(hh * 10) / 10 });
        this.refreshNode(el);
        this.drawOverlay();
      },
      up: (_ev, moved) => {
        this.guides = [];
        if (moved) this.app.commit(before);
        this.drawOverlay();
      },
    });
  }

  startMarquee(e) {
    const p0 = this.toSlide(e);
    const base = e.shiftKey ? [...this.app.selection] : [];
    this.track(e, {
      move: (ev) => {
        const p = this.toSlide(ev);
        this.marquee = { x: Math.min(p0.x, p.x), y: Math.min(p0.y, p.y), w: Math.abs(p.x - p0.x), h: Math.abs(p.y - p0.y) };
        const m = this.marquee;
        // Objects inside the box, with the rest of their groups.
        const hit = this.slide.elements.filter((el) => {
          const b = aabb(el);
          return b.x >= m.x && b.y >= m.y && b.x + b.w <= m.x + m.w && b.y + b.h <= m.y + m.h;
        }).flatMap((el) => groupOf(this.slide, el)).map((el) => el.id);
        this.app.setSelection([...new Set([...base, ...hit])], { quiet: true });
        this.drawOverlay();
      },
      up: () => {
        this.marquee = null;
        this.app.setSelection(this.app.selection);
      },
    });
  }

  // ---------------------------------------------------------------- drawing new elements

  setTool(tool) {
    this.tool = tool;
    this.wrap.classList.toggle('is-drawing', !!tool);
    this.app.onToolChange?.();
  }

  startDraw(e) {
    const tool = this.tool;
    const p0 = this.toSlide(e);
    const isLine = tool.type === 'line';
    this.track(e, {
      move: (ev) => {
        const p = this.toSlide(ev);
        if (isLine) {
          let { x, y } = p;
          if (ev.shiftKey) {
            const len = Math.hypot(x - p0.x, y - p0.y);
            const ang = Math.round(Math.atan2(y - p0.y, x - p0.x) / (Math.PI / 4)) * (Math.PI / 4);
            x = p0.x + len * Math.cos(ang);
            y = p0.y + len * Math.sin(ang);
          }
          this.draft = { line: true, x1: p0.x, y1: p0.y, x2: x, y2: y };
        } else {
          let w = p.x - p0.x;
          let hh = p.y - p0.y;
          if (ev.shiftKey) {
            const m = Math.max(Math.abs(w), Math.abs(hh));
            w = Math.sign(w || 1) * m;
            hh = Math.sign(hh || 1) * m;
          }
          this.draft = { x: Math.min(p0.x, p0.x + w), y: Math.min(p0.y, p0.y + hh), w: Math.abs(w), h: Math.abs(hh) };
        }
        this.drawOverlay();
      },
      up: (_ev, moved) => {
        const d = moved ? this.draft : null;
        this.draft = null;
        this.setTool(null);
        let el;
        const round = (n) => Math.round(n);
        if (isLine) {
          el = d && Math.hypot(d.x2 - d.x1, d.y2 - d.y1) > 4
            ? newLine(round(d.x1), round(d.y1), round(d.x2), round(d.y2), { arrowEnd: tool.arrow })
            : newLine(round(p0.x), round(p0.y), round(p0.x + 160), round(p0.y), { arrowEnd: tool.arrow });
        } else {
          const big = d && d.w > 4 && d.h > 4;
          if (tool.type === 'text') el = big ? newTextBox(round(d.x), round(d.y), round(d.w), round(Math.max(d.h, 30))) : newTextBox(round(p0.x), round(p0.y - 20), 300, 44);
          else el = big ? newShape(tool.shape, round(d.x), round(d.y), round(d.w), round(d.h)) : newShape(tool.shape, round(p0.x - 80), round(p0.y - 60));
        }
        this.app.addElements([el]);
        if (tool.type === 'text') this.startEdit(el.id);
      },
    });
  }

  // ---------------------------------------------------------------- text editing

  get isEditing() { return !!this.editing; }

  startEdit(id, { point = null, selectAll = false } = {}) {
    const el = this.find(id);
    if (!el || !(canHoldText(el) || el.type === 'table')) return;
    if (this.editing?.id === id) return;
    this.endEdit();
    this.app.setSelection([id], { quiet: true });
    this.editing = { id, before: this.app.snapshot(), original: JSON.stringify([el.text, el.cells]), bodies: [] };
    this.render();
    const node = this.nodeOf(id);
    const bodies = el.type === 'table' ? [...node.querySelectorAll('.cell-body')] : [node.querySelector('.el-text-body')];
    this.editing.bodies = bodies;
    for (const body of bodies) {
      body.contentEditable = 'true';
      body.spellcheck = true;
      delete body.dataset.placeholder;
      if (el.type !== 'table' && !el.text) body.innerHTML = el.bullets ? '<ul><li><br></li></ul>' : '<p><br></p>';
      body.addEventListener('input', () => this.onInput(el, body));
      body.addEventListener('keydown', (ev) => this.onEditKey(ev, body));
      body.addEventListener('paste', (ev) => this.onEditPaste(ev));
      body.addEventListener('drop', (ev) => ev.preventDefault());
    }
    let target = bodies[0];
    if (point) {
      const cell = document.elementFromPoint(point.clientX, point.clientY)?.closest?.('.cell-body');
      if (cell && bodies.includes(cell)) target = cell;
    }
    target.focus({ preventScroll: true });
    if (selectAll || !point || !placeCaretAt(point.clientX, point.clientY) || !target.contains(document.getSelection().anchorNode)) {
      caretToEnd(target, selectAll);
    }
    this.drawOverlay();
    this.app.onSelectionChanged();
  }

  onInput(el, body) {
    if (body.querySelector('font')) normalizeFontTags(body);
    if (el.type === 'table') {
      const td = body.closest('td');
      el.cells[+td.dataset.r][+td.dataset.c].text = body.innerHTML;
      this.autoGrow(el);
    } else {
      el.text = body.innerHTML;
      this.autoGrow(el);
    }
    this.app.markDirty({ thumbs: true });
  }

  /** Text boxes grow to fit their text, as in Slides and PowerPoint. */
  autoGrow(el) {
    const node = this.nodeOf(el.id);
    if (!node) return;
    if (el.type === 'table') {
      const table = node.querySelector('table');
      if (table && table.offsetHeight > el.h + 0.5) {
        el.h = Math.ceil(table.offsetHeight);
        node.style.height = `${el.h}px`;
        this.drawOverlay();
      }
      return;
    }
    if (el.type !== 'text' || el.ph) return;
    const box = node.querySelector('.el-textbox');
    const body = box?.firstElementChild;
    if (!body) return;
    const cs = getComputedStyle(box);
    const need = body.offsetHeight + parseFloat(cs.paddingTop) + parseFloat(cs.paddingBottom);
    if (need > el.h + 0.5) {
      el.h = Math.ceil(need);
      node.style.height = `${el.h}px`;
      this.drawOverlay();
    }
  }

  onEditKey(e, body) {
    if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      this.endEdit();
      this.wrap.focus({ preventScroll: true });
    } else if (e.key === 'Tab') {
      e.preventDefault();
      const sel = document.getSelection();
      const inList = sel.anchorNode && (sel.anchorNode.nodeType === 1 ? sel.anchorNode : sel.anchorNode.parentElement)?.closest('li');
      const el = this.find(this.editing.id);
      if (el?.type === 'table') {
        // Tab moves between cells.
        const bodies = this.editing.bodies;
        const i = bodies.indexOf(body) + (e.shiftKey ? -1 : 1);
        if (bodies[i]) {
          bodies[i].focus();
          caretToEnd(bodies[i], true);
        }
      } else if (inList) exec(e.shiftKey ? 'outdent' : 'indent');
      else if (!e.shiftKey) exec('insertText', '\t');
    }
    // Formatting shortcuts while typing; app-wide ones (F5, Ctrl+S, Ctrl+O…) carry on to the editor.
    if ((e.ctrlKey || e.metaKey) && !e.altKey) this.app.textShortcut?.(e);
  }

  onEditPaste(e) {
    const dt = e.clipboardData;
    const files = [...(dt?.files || [])].filter((f) => f.type.startsWith('image/'));
    e.preventDefault();
    if (files.length) {
      this.app.insertImageFiles(files);
      return;
    }
    const text = dt?.getData('text/plain');
    if (text) exec('insertText', text.replace(/\r\n?/g, '\n'));
  }

  /** Put the remembered text selection back (after clicking a toolbar control). */
  restoreSelection() {
    if (!this.editing) return null;
    const range = this.savedRange;
    const body = (range && this.editing.bodies.find((b) => b.contains(range.startContainer))) || this.editing.bodies[0];
    body.focus({ preventScroll: true });
    if (range && body.contains(range.startContainer)) {
      const sel = document.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
    }
    return body;
  }

  /** Apply an editing command to the text being edited. */
  execText(fn) {
    const body = this.restoreSelection();
    if (!body) return;
    fn(body);
    body.dispatchEvent(new Event('input'));
    this.onSelectionChange();
  }

  /** Stop editing without keeping the typed text (the deck is being replaced). */
  cancelEdit() {
    if (!this.editing) return;
    for (const b of this.editing.bodies) b.contentEditable = 'false';
    this.editing = null;
    this.savedRange = null;
  }

  endEdit() {
    if (!this.editing) return;
    const { id, before, original, bodies } = this.editing;
    this.editing = null;
    this.savedRange = null;
    const el = this.find(id);
    for (const b of bodies) b.contentEditable = 'false';
    if (el) {
      if (el.type === 'table') {
        for (const b of bodies) {
          const td = b.closest('td');
          const cell = el.cells[+td.dataset.r]?.[+td.dataset.c];
          if (cell) cell.text = sanitizeHtml(tidyHtml(b.innerHTML));
        }
      } else {
        el.text = sanitizeHtml(tidyHtml(bodies[0]?.innerHTML || ''));
      }
      if (JSON.stringify([el.text, el.cells]) !== original) this.app.commit(before, { render: false });
    }
    document.getSelection()?.removeAllRanges();
    this.render();
    this.app.onSelectionChanged();
  }
}
