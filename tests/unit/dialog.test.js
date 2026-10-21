import { describe, it, expect, beforeAll, afterEach } from 'vitest';
import { openDialog, confirmDialog } from '../../src/ui/dialog.js';
import { safeFileName } from '../../src/ui/dom.js';

beforeAll(() => {
  // jsdom has <dialog> but not its modal methods.
  HTMLDialogElement.prototype.showModal ||= function showModal() { this.open = true; };
  HTMLDialogElement.prototype.close ||= function close() {
    this.open = false;
    this.dispatchEvent(new Event('close'));
  };
});

afterEach(() => {
  for (const d of document.querySelectorAll('dialog')) d.close();
});

const press = (el, type) => el.dispatchEvent(new Event(type, { bubbles: true }));

describe('dialogs', () => {
  it('focuses the main button of a dialog without fields, so Enter does it', () => {
    openDialog({
      title: 'Save changes?',
      body: document.createElement('p'),
      buttons: [{ label: 'Cancel', value: null }, { label: 'Don’t save', value: 'discard' }, { label: 'Save', value: 'save', primary: true }],
    });
    expect(document.activeElement.textContent).toBe('Save');
  });

  it('focuses Cancel rather than a dangerous main button', () => {
    confirmDialog('Delete it?', { confirmLabel: 'Delete', danger: true });
    expect(document.activeElement.textContent).toBe('Cancel');
    expect(document.querySelector('dialog h2').textContent).toBe('LibreSlides');
  });

  it('stays open when a text selection drag ends on the backdrop', async () => {
    const input = document.createElement('input');
    const result = openDialog({ title: 'Rename', body: input });
    const dlg = document.querySelector('dialog');
    press(input, 'pointerdown');
    press(dlg, 'click');
    expect(dlg.open).toBe(true);
    press(dlg, 'pointerdown');
    press(dlg, 'click');
    expect(dlg.open).toBe(false);
    expect(await result).toBeNull();
  });

  it('names an untitled file “presentation”', () => {
    expect(safeFileName('')).toBe('presentation');
  });
});
