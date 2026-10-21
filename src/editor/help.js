import { h, toast, shortcutLabel } from '../ui/dom.js';
import { openDialog } from '../ui/dialog.js';
import { VERSION, BUILD, diagnostics } from '../diagnostics.js';

/** The Help menu's dialogs: keyboard shortcuts (Ctrl+/) and About, with details for a bug report. */

/** Help › Keyboard shortcuts: the editor's, then the slideshow's. */
export function showShortcuts() {
  const rows = [
    ['Present from this slide', 'Mod-Enter', 'Shift-F5'],
    ['Present from the beginning', 'F5', 'Mod-Shift-Enter'],
    ['Save / Save as', 'Mod-S', 'Mod-Shift-S'],
    ['Open / Print', 'Mod-O', 'Mod-P'],
    ['New slide', 'Mod-M'],
    ['New comment', 'Mod-Alt-M'],
    ['Duplicate the selected objects (or slides)', 'Mod-D'],
    ['Group / Ungroup', 'Mod-G', 'Mod-Shift-G'],
    ['Link', 'Mod-K'],
    ['Undo / Redo', 'Mod-Z', 'Mod-Y'],
    ['Cut / Copy / Paste', 'Mod-X', 'Mod-C', 'Mod-V'],
    ['Select all', 'Mod-A'],
    ['Edit the selected text box (or diagram’s text)', 'Enter'],
    ['Stop editing text / deselect', 'Escape'],
    ['Select the next object', 'Tab'],
    ['Nudge (fine / large steps)', 'Arrow keys', 'Alt-Arrow', 'Shift-Arrow'],
    ['Bring forward / to front', 'Mod-Up', 'Mod-Shift-Up'],
    ['Send backward / to back', 'Mod-Down', 'Mod-Shift-Down'],
    ['Bold / Italic / Underline', 'Mod-B', 'Mod-I', 'Mod-U'],
    ['Strikethrough', 'Alt-Shift-5'],
    ['Increase / decrease font size', 'Mod-Shift-.', 'Mod-Shift-,'],
    ['Align left / centre / right / justify', 'Mod-Shift-L', 'Mod-Shift-E', 'Mod-Shift-R', 'Mod-Shift-J'],
    ['Bulleted / numbered list', 'Mod-Shift-8', 'Mod-Shift-7'],
    ['Decrease / increase indent (while typing)', 'Mod-[', 'Mod-]'],
    ['Clear formatting', 'Mod-\\'],
    ['Previous / next slide', 'Page Up', 'Page Down'],
    ['Move slide up / down (in the slide list)', 'Mod-Up', 'Mod-Down'],
    ['Zoom in / out / fit', 'Mod-=', 'Mod--', 'Mod-Alt-0'],
    ['Keep proportions while resizing; draw straight lines', 'Shift-Drag'],
    ['Move without snapping', 'Alt-Drag'],
    ['Keyboard shortcuts', 'Mod-/'],
  ];
  const table = h('table', { class: 'shortcut-table' }, h('tbody', {}, ...rows.map(([label, ...keys]) => h('tr', {}, h('td', {}, label), h('td', {}, ...keys.flatMap((kk, i) => [i ? ' ' : null, h('kbd', {}, /^[A-Z][a-z]/.test(kk) && !kk.includes('-') ? kk : shortcutLabel(kk))]))))));
  const present = h('table', { class: 'shortcut-table' }, h('tbody', {}, ...[
    ['Next / previous', '→ Space Enter', '← Backspace'],
    ['First / last slide', 'Home', 'End'],
    ['Go to slide number', 'Number then Enter'],
    ['Black / white screen', 'B', 'W'],
    ['Show speaker notes', 'S'],
    ['End the slideshow', 'Esc'],
  ].map(([label, ...keys]) => h('tr', {}, h('td', {}, label), h('td', {}, ...keys.flatMap((kk, i) => [i ? ' ' : null, h('kbd', {}, kk)]))))));
  openDialog({ title: 'Keyboard shortcuts', className: 'wide', body: h('div', {}, table, h('h3', { class: 'dialog-subhead' }, 'During a slideshow'), present) });
}

/** Help › About LibreSlides: the version, and the diagnostics report to copy into a bug report. */
export async function showAbout(app) {
  const report = await diagnostics(app);
  const pre = h('pre', { class: 'diagnostics' }, report);
  openDialog({
    title: 'About LibreSlides',
    className: 'wide',
    body: h('div', {},
      h('p', {}, `LibreSlides ${VERSION} (build ${BUILD}) — a presentation editor that runs entirely in your browser and works offline. Presentations are kept on this device; nothing is uploaded.`),
      h('details', {}, h('summary', {}, 'Details for a bug report'), pre)),
    buttons: [
      { label: 'Copy details', value: 'copy' },
      { label: 'Close', value: null, primary: true },
    ],
    onSubmit: (v) => {
      if (v === 'copy') {
        navigator.clipboard?.writeText(report).then(() => toast('Copied'), () => {});
        return false;
      }
      return true;
    },
  });
}
