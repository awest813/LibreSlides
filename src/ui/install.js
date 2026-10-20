/**
 * "Install app" support. Chrome and Edge (including on Chromebooks and
 * Android) fire `beforeinstallprompt` when LibreSlides can be installed; we
 * keep the event so our Install button can show the browser's install dialog.
 * Other browsers install from their own menus, so we explain how. Once
 * installed, LibreSlides gets its own window and launcher icon, works offline
 * and appears in the file manager's "Open with" menu for .pptx files.
 */
import { h, toast } from './dom.js';
import { openDialog } from './dialog.js';

let deferred = null;

export const isInstalledWindow = () =>
  typeof matchMedia === 'function' && ['standalone', 'window-controls-overlay', 'minimal-ui'].some((m) => matchMedia(`(display-mode: ${m})`).matches);

/** For bug reports: whether the browser has offered to install, or it's already installed. */
export const installState = () => (isInstalledWindow() ? 'installed app' : deferred ? 'installable (the browser offered it)' : 'not offered by this browser yet');

export function initInstall({ onInstalled } = {}) {
  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferred = e;
  });
  window.addEventListener('appinstalled', () => {
    deferred = null;
    onInstalled?.();
  });
}

/** Show the browser's install dialog. Resolves true if the user accepted. */
async function promptInstall() {
  const e = deferred;
  if (!e) return false;
  // The event can only be used once.
  deferred = null;
  e.prompt();
  const { outcome } = await e.userChoice.catch(() => ({ outcome: 'dismissed' }));
  return outcome === 'accepted';
}

/** Install now if the browser lets us, otherwise explain how to install from the browser's menu. */
export async function showInstallHelp() {
  if (isInstalledWindow()) {
    toast('LibreSlides is already running as an installed app.');
    return;
  }
  if (deferred) {
    await promptInstall();
    return;
  }
  const ua = navigator.userAgent;
  const ios = /iPhone|iPad|iPod/.test(ua) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
  const safari = /Safari/.test(ua) && !/Chrome|Chromium|Edg|OPR|Firefox|FxiOS|CriOS/.test(ua);
  const firefox = /Firefox|FxiOS/.test(ua);
  let steps;
  if (ios) steps = ['Tap the Share button.', 'Choose “Add to Home Screen”.', 'Open LibreSlides from your home screen; it then works offline.'];
  else if (safari) steps = ['In the menu bar, choose File › Add to Dock (Safari 17 or later).', 'Open LibreSlides from the Dock or Launchpad; it then works offline.'];
  else if (firefox) steps = ['Firefox on computers can’t install web apps. LibreSlides still works offline in this tab once it has loaded.', 'To install it as an app, open this page in Chrome, Edge or Safari.'];
  else steps = ['Click the install icon at the right of the address bar, or open the browser menu (⋮ or …) and choose “Install LibreSlides” (it may be under “Cast, save and share” or “Apps”).', 'If there’s no install option yet, reload the page once and try again.'];
  await openDialog({
    title: 'Install LibreSlides',
    body: h('div', {},
      h('p', {}, 'Installed, LibreSlides gets its own window and icon, works without an internet connection, and can open PowerPoint files from your file manager.'),
      h('ol', { class: 'install-steps' }, ...steps.map((t) => h('li', {}, t)))),
  });
}
