/** Tiny DOM helpers — no framework needed for an app this size. */

export function h(tag, props = {}, ...children) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(props || {})) {
    if (v == null || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k === 'html') el.innerHTML = v;
    else if (k === 'text') el.textContent = v;
    else if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
    else if (k === 'dataset') Object.assign(el.dataset, v);
    else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2).toLowerCase(), v);
    else if (v === true) el.setAttribute(k, '');
    else el.setAttribute(k, v);
  }
  for (const c of children.flat()) {
    if (c == null || c === false) continue;
    el.append(c instanceof Node ? c : document.createTextNode(String(c)));
  }
  return el;
}

export const $ = (sel, root = document) => root.querySelector(sel);
export const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

export const escapeHtml = (value = '') =>
  String(value).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

export const isMac = /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent);
export const modKey = isMac ? '⌘' : 'Ctrl';
/** "Mod-Shift-L" → "Ctrl+Shift+L" (or "⌘⇧L" on a Mac), like PowerPoint's menus. */
export const shortcutLabel = (s) => {
  if (!s) return '';
  const label = s.replace(/Mod/g, modKey).replace(/Shift/g, isMac ? '⇧' : 'Shift').replace(/Alt/g, isMac ? '⌥' : 'Alt');
  // Join modifier chains with "+" (Mac uses no separator); keep a literal trailing "-" key.
  return label.replace(/(\S)-(?=\S)/g, isMac ? '$1' : '$1+');
};

// ---------------------------------------------------------------------------
// Popovers (dropdown menus, galleries, colour pickers)
// ---------------------------------------------------------------------------

let openPopover = null;

/**
 * Where keyboard focus goes when the element holding it disappears (a menu
 * item that was clicked, a dialog opened from one): the editor screen points
 * this at the document so typing carries on there, as in PowerPoint.
 */
let focusFallback = null;
export const setFocusFallback = (fn) => { focusFallback = fn; };
const focusIsLost = () => !document.activeElement || document.activeElement === document.body;
/** Give focus to the fallback once the current task settles, unless something else took it. */
export function restoreLostFocus() {
  setTimeout(() => {
    if (focusIsLost() && !openPopover && !document.querySelector('dialog[open]')) focusFallback?.();
  });
}

/** The button that reports the popover as expanded: the anchor, or the caret of a split button. */
const expanderOf = (anchor) => (anchor?.matches?.('button, [role]') ? anchor : anchor?.querySelector?.('[aria-haspopup]') || null);

export function closePopover() {
  if (!openPopover) return;
  const { el, anchor, onClose, cleanup } = openPopover;
  openPopover = null;
  cleanup();
  const hadFocus = el.contains(document.activeElement);
  el.remove();
  if (hadFocus) restoreLostFocus();
  expanderOf(anchor)?.setAttribute('aria-expanded', 'false');
  anchor?.classList.remove('is-open');
  onClose?.();
}

export const isPopoverOpen = () => Boolean(openPopover);

/**
 * Show `content` in a floating panel anchored below `anchor` (or at {x, y}).
 * Clicking outside or pressing Escape closes it.
 */
export function showPopover(anchor, content, { at = null, onClose = null, className = '', focus = true, placement = 'bottom-start', label = null } = {}) {
  const sameAnchor = openPopover && anchor && openPopover.anchor === anchor;
  closePopover();
  if (sameAnchor) return null; // clicking the trigger again toggles it shut

  // A menu is its own widget; anything else is a small dialog named after its button.
  const isMenu = content.getAttribute?.('role') === 'menu';
  const name = label || anchor?.getAttribute?.('aria-label') || anchor?.getAttribute?.('title') || anchor?.textContent?.trim() || 'Options';
  const el = h('div', { class: `popover ${className}`, role: isMenu ? null : 'dialog', 'aria-label': isMenu ? null : name });
  el.append(content);
  document.body.append(el);

  const place = () => {
    const pr = el.getBoundingClientRect();
    let x;
    let y;
    if (at) {
      x = at.x;
      y = at.y;
    } else {
      const r = anchor.getBoundingClientRect();
      x = placement === 'bottom-end' ? r.right - pr.width : r.left;
      y = r.bottom + 2;
      if (y + pr.height > window.innerHeight - 8 && r.top - pr.height - 2 > 8) y = r.top - pr.height - 2;
    }
    x = Math.max(8, Math.min(x, window.innerWidth - pr.width - 8));
    y = Math.max(8, Math.min(y, window.innerHeight - pr.height - 8));
    el.style.left = `${Math.round(x)}px`;
    el.style.top = `${Math.round(y)}px`;
  };
  place();

  const onDown = (e) => {
    if (!el.contains(e.target) && !(anchor && anchor.contains(e.target))) closePopover();
  };
  const onKey = (e) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      const a = anchor;
      closePopover();
      a?.focus?.();
    } else if (e.key === 'Tab' && isMenu) {
      // Tab leaves a menu: it closes, and focus moves on from its button (or back to the document).
      const a = anchor?.isConnected ? expanderOf(anchor) || anchor : null;
      closePopover();
      if (a) a.focus();
      else e.preventDefault();
    } else if (['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(e.key)) {
      const items = $$('[role="menuitem"]:not([disabled]),[role="menuitemcheckbox"]:not([disabled])', el);
      // Other controls (the table size grid) handle their own arrow keys.
      const active = document.activeElement;
      if (!items.length || (el.contains(active) && active !== el && !items.includes(active))) return;
      e.preventDefault();
      const i = items.indexOf(document.activeElement);
      const next = { Home: 0, End: items.length - 1, ArrowDown: (i + 1) % items.length, ArrowUp: (i - 1 + items.length) % items.length }[e.key];
      items[next].focus();
    }
  };
  // A panel (colour picker, shapes…) closes when the keyboard moves on from it.
  el.addEventListener('focusout', (e) => {
    const to = e.relatedTarget;
    if (to && !el.contains(to) && !(anchor && anchor.contains(to)) && openPopover?.el === el) closePopover();
  });
  const onResize = () => closePopover();
  document.addEventListener('pointerdown', onDown, true);
  document.addEventListener('keydown', onKey, true);
  window.addEventListener('resize', onResize);
  window.addEventListener('blur', onResize);

  openPopover = {
    el,
    anchor,
    onClose,
    cleanup: () => {
      document.removeEventListener('pointerdown', onDown, true);
      document.removeEventListener('keydown', onKey, true);
      window.removeEventListener('resize', onResize);
      window.removeEventListener('blur', onResize);
    },
  };
  expanderOf(anchor)?.setAttribute('aria-expanded', 'true');
  anchor?.classList.add('is-open');
  if (focus) {
    const first = el.querySelector('[role="menuitem"],[role="menuitemcheckbox"],input,button');
    first?.focus({ preventScroll: true });
  }
  return el;
}

/**
 * Build a menu. Items: { label, icon?, shortcut?, checked?, disabled?, run }
 * or 'separator' or { heading }.
 */
export function menu(items, { iconFn } = {}) {
  const list = h('div', { class: 'menu', role: 'menu' });
  for (const item of items) {
    if (!item) continue;
    if (item === 'separator') {
      list.append(h('div', { class: 'menu-sep', role: 'separator' }));
      continue;
    }
    if (item.heading) {
      list.append(h('div', { class: 'menu-heading' }, item.heading));
      continue;
    }
    const btn = h(
      'button',
      {
        type: 'button',
        class: `menu-item${item.checked ? ' is-checked' : ''}`,
        role: item.checked !== undefined ? 'menuitemcheckbox' : 'menuitem',
        tabindex: '-1', // the arrow keys move between items; Tab leaves the menu
        'aria-checked': item.checked !== undefined ? String(Boolean(item.checked)) : null,
        disabled: item.disabled || null,
        style: item.style || null,
        onclick: () => {
          closePopover();
          item.run?.();
        },
      },
      h('span', { class: 'menu-icon', html: item.icon && iconFn ? iconFn(item.icon) : item.checked ? '✓' : '' }),
      h('span', { class: 'menu-label', style: item.labelStyle || null }, item.label),
      h('span', { class: 'menu-shortcut' }, shortcutLabel(item.shortcut || '')),
    );
    list.append(btn);
  }
  return list;
}

/**
 * Make the controls `items()` of `container` (a toolbar, menu bar or tab
 * list) a single tab stop that the arrow keys, Home and End move along,
 * skipping hidden and disabled ones. `onMove(control)` runs after a move.
 * Returns sync(), to call after controls are shown, hidden or disabled.
 */
export function rovingFocus(container, items, { prev = 'ArrowLeft', next = 'ArrowRight', onMove = null } = {}) {
  const usable = () => items().filter((b) => !b.disabled && !b.closest('[hidden]'));
  const setCurrent = (cur) => { for (const b of items()) b.tabIndex = b === cur ? 0 : -1; };
  const sync = () => {
    const list = usable();
    setCurrent(list.find((b) => b.tabIndex === 0) || list[0]);
  };
  container.addEventListener('focusin', (e) => {
    if (items().includes(e.target)) setCurrent(e.target);
  });
  container.addEventListener('keydown', (e) => {
    const list = usable();
    const i = list.indexOf(e.target);
    if (i < 0 || e.altKey || e.ctrlKey || e.metaKey) return;
    const moves = { [next]: (i + 1) % list.length, [prev]: (i - 1 + list.length) % list.length, Home: 0, End: list.length - 1 };
    if (!Object.hasOwn(moves, e.key)) return;
    const j = moves[e.key];
    e.preventDefault();
    list[j].focus();
    onMove?.(list[j]);
  });
  sync();
  return sync;
}

// ---------------------------------------------------------------------------
// Toasts
// ---------------------------------------------------------------------------

let toastHost;
export function toast(message, { type = 'info', timeout = 3200, action = null } = {}) {
  if (!toastHost) {
    toastHost = h('div', { class: 'toast-host', role: 'status', 'aria-live': 'polite' });
    document.body.append(toastHost);
  }
  const el = h('div', { class: `toast toast-${type}` }, h('span', {}, message));
  if (action) {
    el.append(h('button', { type: 'button', class: 'toast-action', onclick: () => { action.run(); el.remove(); } }, action.label));
  }
  toastHost.append(el);
  requestAnimationFrame(() => el.classList.add('is-visible'));
  setTimeout(() => {
    el.classList.remove('is-visible');
    setTimeout(() => el.remove(), 250);
  }, timeout);
}

export function debounce(fn, ms) {
  let t;
  const wrapped = (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), ms);
  };
  wrapped.cancel = () => clearTimeout(t);
  return wrapped;
}

export function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = h('a', { href: url, download: filename, style: { display: 'none' } });
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}

/** Show the browser's file chooser. Resolves to the chosen File, or null if cancelled. */
export function pickFileInput(accept = '') {
  return new Promise((resolve) => {
    const input = h('input', { type: 'file', accept, style: { display: 'none' } });
    const done = (file) => {
      resolve(file);
      input.remove();
    };
    input.addEventListener('change', () => done(input.files?.[0] || null));
    input.addEventListener('cancel', () => done(null));
    document.body.append(input);
    input.click();
  });
}

export const safeFileName = (value = 'presentation') => {
  const name = String(value).trim().replace(/[<>:"/\\|?*\u0000-\u001f]+/g, '-').replace(/[. ]+$/g, '');
  return (name || 'presentation').slice(0, 120);
};

export function formatDate(ts) {
  const d = new Date(ts);
  const now = new Date();
  const sameDay = d.toDateString() === now.toDateString();
  if (sameDay) return d.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });
  const yesterday = new Date(now);
  yesterday.setDate(now.getDate() - 1);
  if (d.toDateString() === yesterday.toDateString()) return 'Yesterday';
  return d.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: d.getFullYear() === now.getFullYear() ? undefined : 'numeric' });
}
