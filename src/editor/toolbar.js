import { h, showPopover, menu, closePopover, shortcutLabel } from '../ui/dom.js';
import { icon } from '../ui/icons.js';
import { openColorPicker } from '../ui/color-picker.js';
import { FONTS, FONT_SIZES, fontStack, embedFontsSetting, setEmbedFontsSetting } from '../model/fonts.js';
import { SHAPES, shapePath } from '../model/shapes.js';
import { LAYOUTS } from '../model/deck.js';
import { resolveColor } from '../model/themes.js';

const iconFn = (n) => icon(n);
const keep = (e) => e.preventDefault(); // keep focus (and the text selection) where it is

function tb(iconName, label, run, { shortcut = '', cls = '', haspopup = false, text = null } = {}) {
  const tip = shortcut ? `${label} (${shortcutLabel(shortcut)})` : label;
  const b = h('button', {
    type: 'button', class: `tb ${cls}`, title: tip, 'aria-label': label, 'aria-haspopup': haspopup ? 'true' : null,
    'aria-expanded': haspopup ? 'false' : null, html: `${iconName ? icon(iconName) : ''}${text ? `<span class="tb-text">${text}</span>` : ''}${haspopup ? icon('chevronDown', 'icon tb-caret') : ''}`,
    onmousedown: keep,
  });
  b.addEventListener('click', () => run(b));
  return b;
}

const sep = () => h('span', { class: 'tb-sep', role: 'separator' });

/** A toolbar button with a coloured bar under its icon showing the current colour. */
function swatchButton(iconName, label, run) {
  const bar = h('span', { class: 'swatch-bar' });
  const b = h('button', { type: 'button', class: 'tb tb-swatch', title: label, 'aria-label': label, 'aria-haspopup': 'true', onmousedown: keep });
  b.innerHTML = icon(iconName);
  b.append(bar);
  b.addEventListener('click', () => run(b));
  b.setColor = (c) => { bar.style.background = c; };
  return b;
}

/**
 * The menu bar (File, Edit, View…) and the contextual toolbar below it, like
 * Google Slides: insert tools are always there; formatting controls appear
 * for whatever is selected.
 */
export class Toolbar {
  constructor(app) {
    this.app = app;
    this.menubar = this.buildMenubar();
    this.toolbar = this.buildToolbar();
  }

  // ---------------------------------------------------------------- menus

  buildMenubar() {
    const app = this.app;
    const menus = {
      File: () => [
        { label: 'New presentation', icon: 'newDoc', run: () => app.newPresentation() },
        { label: 'Open…', icon: 'open', shortcut: 'Mod-O', run: () => app.openFile() },
        'separator',
        { label: app.deck.file ? `Save to “${app.deck.file.name}”` : 'Save to file…', icon: 'save', shortcut: 'Mod-S', run: () => app.saveToFile() },
        { label: 'Save as…', icon: 'fileDown', shortcut: 'Mod-Shift-S', run: () => app.saveToFile({ as: true }) },
        { label: 'Make a copy', icon: 'duplicate', run: () => app.makeCopy() },
        'separator',
        { heading: 'Download a copy' },
        { label: 'PowerPoint (.pptx)', icon: 'fileDown', run: () => app.downloadPptx() },
        { label: 'Embed fonts in PowerPoint files', checked: embedFontsSetting(), run: () => setEmbedFontsSetting(!embedFontsSetting()) },
        { label: 'PDF document (.pdf)', icon: 'fileDown', run: () => app.printDeck({ pdf: true }) },
        { label: 'LibreSlides file (.lslides)', icon: 'fileDown', run: () => app.downloadNative() },
        'separator',
        { label: 'Rename…', icon: 'rename', run: () => app.renamePrompt() },
        { label: 'Slide size…', icon: 'layout', run: () => app.openPanel('design') },
        { label: 'Print', icon: 'print', shortcut: 'Mod-P', run: () => app.printDeck() },
        'separator',
        { label: 'Back to all presentations', icon: 'home', run: () => app.goHome() },
      ],
      Edit: () => {
        // While typing, these act on the selected text.
        const typing = app.canvas.isEditing;
        const none = typing ? !app.hasTextSelection() : !app.selection.length;
        return [
          { label: 'Undo', icon: 'undo', shortcut: 'Mod-Z', disabled: !app.history.canUndo && !typing, run: () => app.undo() },
          { label: 'Redo', icon: 'redo', shortcut: 'Mod-Y', disabled: !app.history.canRedo, run: () => app.redo() },
          'separator',
          { label: 'Cut', icon: 'cut', shortcut: 'Mod-X', disabled: none, run: () => app.editCommand('cut') },
          { label: 'Copy', icon: 'copy', shortcut: 'Mod-C', disabled: none, run: () => app.editCommand('copy') },
          { label: 'Paste', icon: 'paste', shortcut: 'Mod-V', run: () => app.editCommand('paste') },
          { label: 'Duplicate', icon: 'duplicate', shortcut: 'Mod-D', disabled: typing || !app.selection.length, run: () => app.duplicateSelection() },
          { label: 'Delete', icon: 'trash', shortcut: 'Delete', disabled: none, run: () => app.editCommand('delete') },
          'separator',
          { label: 'Select all', shortcut: 'Mod-A', run: () => app.selectAll() },
          { label: 'Select none', run: () => app.setSelection([]) },
        ];
      },
      View: () => [
        { label: 'Slideshow', icon: 'present', shortcut: 'Mod-Enter', run: () => app.present({ fromStart: false }) },
        { label: 'Slideshow from beginning', icon: 'play', shortcut: 'F5', run: () => app.present({ fromStart: true }) },
        { label: 'Presenter view', icon: 'presenter', run: () => app.presenterView() },
        'separator',
        { label: 'Zoom in', icon: 'zoomIn', shortcut: 'Mod-=', run: () => app.zoomBy(1.25) },
        { label: 'Zoom out', icon: 'zoomOut', shortcut: 'Mod--', run: () => app.zoomBy(0.8) },
        { label: 'Fit', shortcut: 'Mod-Alt-0', checked: app.canvas.zoom === 'fit', run: () => app.canvas.setZoom('fit') },
        'separator',
        { label: 'Speaker notes', checked: app.view.notes, run: () => app.toggleNotes() },
        { label: 'Snap to guides', checked: app.canvas.snapEnabled, run: () => app.toggleSnap() },
        { label: 'Side panel', checked: !!app.panel.current, run: () => (app.panel.current ? app.closePanel() : app.openPanel('design')) },
        { label: 'Dark mode', checked: app.isDarkUi(), run: () => app.toggleDarkUi() },
        { label: 'Full screen', icon: 'fullscreen', run: () => app.toggleFullscreen() },
      ],
      Insert: () => [
        { label: 'Text box', icon: 'textBox', run: () => app.canvas.setTool({ type: 'text' }) },
        { label: 'Image…', icon: 'image', run: () => app.insertImage() },
        { label: 'Link…', icon: 'link', shortcut: 'Mod-K', run: () => app.insertLink() },
        { label: 'Slide number', run: () => app.insertField('slidenum') },
        { label: 'Date', run: () => app.insertField('datetime1') },
        { label: 'Video or audio…', icon: 'media', run: () => app.insertMedia() },
        { label: 'Table…', icon: 'table', run: () => this.tableMenu(this.menuButtons.Insert) },
        { label: 'Chart…', icon: 'chart', run: () => app.insertChart() },
        { label: 'Diagram…', icon: 'diagram', run: () => app.insertDiagram() },
        'separator',
        ...SHAPES.slice(0, 6).map((s) => ({ label: s.name, run: () => app.canvas.setTool({ type: 'shape', shape: s.id }) })),
        { label: 'More shapes…', icon: 'shapes', run: () => this.shapeMenu(this.shapeBtn) },
        { label: 'Line', icon: 'line', run: () => app.canvas.setTool({ type: 'line' }) },
        { label: 'Arrow', icon: 'arrow', run: () => app.canvas.setTool({ type: 'line', arrow: true }) },
        'separator',
        { label: 'Comment', icon: 'comment', shortcut: 'Mod-Alt-M', run: () => app.insertComment() },
        { label: 'New slide', icon: 'plus', shortcut: 'Mod-M', run: () => app.addSlide() },
      ],
      Format: () => {
        const t = app.hasTextTarget();
        return [
          { label: 'Bold', icon: 'bold', shortcut: 'Mod-B', disabled: !t, run: () => app.toggleTextProp('bold') },
          { label: 'Italic', icon: 'italic', shortcut: 'Mod-I', disabled: !t, run: () => app.toggleTextProp('italic') },
          { label: 'Underline', icon: 'underline', shortcut: 'Mod-U', disabled: !t, run: () => app.toggleTextProp('underline') },
          { label: 'Strikethrough', icon: 'strike', shortcut: 'Alt-Shift-5', disabled: !t, run: () => app.toggleTextProp('strike') },
          { label: 'Increase font size', shortcut: 'Mod-Shift-.', disabled: !t, run: () => app.stepFontSize(1) },
          { label: 'Decrease font size', shortcut: 'Mod-Shift-,', disabled: !t, run: () => app.stepFontSize(-1) },
          'separator',
          { label: 'Align left', icon: 'alignLeft', shortcut: 'Mod-Shift-L', disabled: !t, run: () => app.setTextProp('align', 'left') },
          { label: 'Center', icon: 'alignCenter', shortcut: 'Mod-Shift-E', disabled: !t, run: () => app.setTextProp('align', 'center') },
          { label: 'Align right', icon: 'alignRight', shortcut: 'Mod-Shift-R', disabled: !t, run: () => app.setTextProp('align', 'right') },
          { label: 'Justify', icon: 'alignJustify', shortcut: 'Mod-Shift-J', disabled: !t, run: () => app.setTextProp('align', 'justify') },
          'separator',
          { label: 'Bulleted list', icon: 'bulletList', shortcut: 'Mod-Shift-8', disabled: !t, run: () => app.toggleList('ul') },
          { label: 'Numbered list', icon: 'orderedList', shortcut: 'Mod-Shift-7', disabled: !t, run: () => app.toggleList('ol') },
          { label: 'Clear formatting', icon: 'clearFormat', shortcut: 'Mod-\\', disabled: !t, run: () => app.clearFormatting() },
          'separator',
          { label: 'Format options…', icon: 'format', disabled: !app.selection.length, run: () => app.openPanel('format') },
        ];
      },
      Slide: () => [
        { label: 'New slide', icon: 'plus', shortcut: 'Mod-M', run: () => app.addSlide() },
        { label: 'Duplicate slide', icon: 'duplicate', run: () => app.duplicateSlides() },
        { label: 'Delete slide', icon: 'trash', run: () => app.deleteSlides() },
        { label: app.slide.hidden ? 'Unskip slide' : 'Skip slide', icon: app.slide.hidden ? 'eye' : 'eyeOff', run: () => app.toggleSkip() },
        'separator',
        { label: 'Move slide up', icon: 'up', disabled: app.slideIndex === 0, run: () => app.moveSlides(-1) },
        { label: 'Move slide down', icon: 'down', disabled: app.slideIndex === app.deck.slides.length - 1, run: () => app.moveSlides(1) },
        'separator',
        ...app.filmstrip.sectionItems(),
        'separator',
        { label: 'Change background…', icon: 'image', run: () => app.openPanel('design') },
        { label: 'Apply layout…', icon: 'layout', run: () => this.layoutMenu(this.menuButtons.Slide) },
        { label: 'Transition…', icon: 'animate', run: () => app.openPanel('transition') },
        { label: 'Change theme…', icon: 'theme', run: () => app.openPanel('design') },
      ],
      Arrange: () => this.arrangeItems(),
      Help: () => [
        { label: 'Keyboard shortcuts', icon: 'keyboard', shortcut: 'Mod-/', run: () => app.showShortcuts() },
        { label: 'Install LibreSlides…', icon: 'download', run: () => app.showInstall() },
        { label: 'About LibreSlides', icon: 'info', run: () => app.showAbout() },
      ],
    };
    this.menuButtons = {};
    const bar = h('div', { class: 'menubar', role: 'menubar', 'aria-label': 'Menu' });
    const names = Object.keys(menus);
    for (const name of names) {
      const b = h('button', { type: 'button', class: 'menubar-item', role: 'menuitem', 'aria-haspopup': 'menu', 'aria-expanded': 'false', onmousedown: keep }, name);
      b.addEventListener('click', () => showPopover(b, menu(menus[name](), { iconFn }), { label: name }));
      b.addEventListener('keydown', (e) => {
        if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
        const next = this.menuButtons[names[(names.indexOf(name) + (e.key === 'ArrowRight' ? 1 : names.length - 1)) % names.length]];
        next.focus();
      });
      this.menuButtons[name] = b;
      bar.append(b);
    }
    // Moving the pointer across the bar while a menu is open switches menus, as in desktop apps.
    bar.addEventListener('pointerover', (e) => {
      const b = e.target.closest('.menubar-item');
      const open = bar.querySelector('.menubar-item.is-open');
      if (b && open && b !== open) b.click();
    });
    return bar;
  }

  arrangeItems() {
    const app = this.app;
    const n = app.selection.length;
    const none = !n;
    return [
      { label: 'Bring to front', icon: 'toFront', shortcut: 'Mod-Shift-Up', disabled: none, run: () => app.order('front') },
      { label: 'Bring forward', shortcut: 'Mod-Up', disabled: none, run: () => app.order('forward') },
      { label: 'Send backward', shortcut: 'Mod-Down', disabled: none, run: () => app.order('backward') },
      { label: 'Send to back', icon: 'toBack', shortcut: 'Mod-Shift-Down', disabled: none, run: () => app.order('back') },
      'separator',
      { label: 'Group', shortcut: 'Mod-G', disabled: n < 2, run: () => app.groupSelection() },
      { label: 'Ungroup', shortcut: 'Mod-Shift-G', disabled: !app.selectedElements().some((e) => e.groupId), run: () => app.ungroupSelection() },
      'separator',
      { heading: n > 1 ? 'Align to each other' : 'Align on slide' },
      { label: 'Left', icon: 'alignLeftEdge', disabled: none, run: () => app.align('left') },
      { label: 'Center', icon: 'alignHCenter', disabled: none, run: () => app.align('center') },
      { label: 'Right', icon: 'alignRightEdge', disabled: none, run: () => app.align('right') },
      { label: 'Top', icon: 'alignTopEdge', disabled: none, run: () => app.align('top') },
      { label: 'Middle', icon: 'alignVCenter', disabled: none, run: () => app.align('middle') },
      { label: 'Bottom', icon: 'alignBottomEdge', disabled: none, run: () => app.align('bottom') },
      { label: 'Distribute horizontally', icon: 'distributeH', disabled: n < 3, run: () => app.distribute('x') },
      { label: 'Distribute vertically', icon: 'distributeV', disabled: n < 3, run: () => app.distribute('y') },
      'separator',
      { label: 'Rotate clockwise 90°', icon: 'rotateCw', disabled: none, run: () => app.rotateBy(90) },
      { label: 'Rotate counter-clockwise 90°', icon: 'rotateCcw', disabled: none, run: () => app.rotateBy(-90) },
      { label: 'Flip horizontally', icon: 'flipH', disabled: none, run: () => app.flip('h') },
      { label: 'Flip vertically', icon: 'flipV', disabled: none, run: () => app.flip('v') },
      'separator',
      { label: app.selectedElements().every((e) => e.locked) && n ? 'Unlock position' : 'Lock position', disabled: none, run: () => app.toggleLock() },
    ];
  }

  shapeMenu(anchor) {
    const app = this.app;
    const panel = h('div', { class: 'shape-panel' });
    let group = null;
    for (const s of SHAPES) {
      if (s.group !== group) {
        group = s.group;
        panel.append(h('div', { class: 'color-heading' }, group));
      }
      const b = h('button', {
        type: 'button', class: 'shape-chip', title: s.name, 'aria-label': s.name,
        html: shapeIcon(s.id),
        onclick: () => { closePopover(); app.canvas.setTool({ type: 'shape', shape: s.id }); },
      });
      panel.append(b);
    }
    showPopover(anchor, panel, { label: 'Shapes' });
  }

  lineMenu(anchor) {
    showPopover(anchor, menu([
      { label: 'Line', icon: 'line', run: () => this.app.canvas.setTool({ type: 'line' }) },
      { label: 'Arrow', icon: 'arrow', run: () => this.app.canvas.setTool({ type: 'line', arrow: true }) },
    ], { iconFn }), { label: 'Lines' });
  }

  tableMenu(anchor) {
    const label = h('div', { class: 'table-grid-label' }, 'Insert table');
    const grid = h('div', { class: 'table-grid', role: 'grid' });
    const cells = [];
    const mark = (r, c) => {
      cells.forEach((b) => b.classList.toggle('is-on', +b.dataset.r <= r && +b.dataset.c <= c));
      label.textContent = `${c + 1} × ${r + 1} table`;
    };
    for (let r = 0; r < 8; r++) {
      for (let c = 0; c < 10; c++) {
        const b = h('button', { type: 'button', 'aria-label': `${c + 1} by ${r + 1} table`, dataset: { r, c } });
        b.addEventListener('pointerenter', () => mark(r, c));
        b.addEventListener('focus', () => mark(r, c));
        b.addEventListener('click', () => { closePopover(); this.app.insertTable(r + 1, c + 1); });
        cells.push(b);
        grid.append(b);
      }
    }
    showPopover(anchor, h('div', { class: 'table-grid-picker' }, label, grid), { label: 'Insert table' });
  }

  layoutMenu(anchor) {
    showPopover(anchor, menu(LAYOUTS.map((l) => ({ label: l.name, checked: this.app.slide.layout === l.id, run: () => this.app.applyLayout(l.id) }))), { label: 'Layout' });
  }

  zoomMenu(anchor) {
    const app = this.app;
    showPopover(anchor, menu([
      { label: 'Fit', checked: app.canvas.zoom === 'fit', run: () => app.canvas.setZoom('fit') },
      'separator',
      ...[0.5, 0.75, 1, 1.5, 2].map((z) => ({ label: `${z * 100}%`, checked: app.canvas.zoom === z, run: () => app.canvas.setZoom(z) })),
    ]), { label: 'Zoom' });
  }

  // ---------------------------------------------------------------- toolbar

  buildToolbar() {
    const app = this.app;
    const groups = {};
    const group = (name, ...children) => {
      groups[name] = h('div', { class: 'tb-group', dataset: { group: name } }, ...children);
      return groups[name];
    };

    this.undoBtn = tb('undo', 'Undo', () => app.undo(), { shortcut: 'Mod-Z' });
    this.redoBtn = tb('redo', 'Redo', () => app.redo(), { shortcut: 'Mod-Y' });
    this.zoomBtn = tb(null, 'Zoom', (b) => this.zoomMenu(b), { haspopup: true, text: 'Fit', cls: 'tb-zoom' });

    this.selectBtn = tb('select', 'Select', () => app.canvas.setTool(null));
    this.textBtn = tb('textBox', 'Text box', () => app.canvas.setTool(app.canvas.tool?.type === 'text' ? null : { type: 'text' }));
    this.imageBtn = tb('image', 'Insert image', () => app.insertImage());
    this.shapeBtn = tb('shapes', 'Shape', (b) => this.shapeMenu(b), { haspopup: true });
    this.lineBtn = tb('line', 'Line', (b) => this.lineMenu(b), { haspopup: true });
    this.tableBtn = tb('table', 'Table', (b) => this.tableMenu(b), { haspopup: true });
    this.chartBtn = tb('chart', 'Insert chart', () => app.insertChart());

    // Slide (nothing selected)
    group('slide',
      tb('image', 'Background', () => app.openPanel('design'), { text: 'Background' }),
      tb('layout', 'Layout', (b) => this.layoutMenu(b), { haspopup: true, text: 'Layout' }),
      tb('theme', 'Theme', () => app.openPanel('design'), { text: 'Theme' }),
      tb('animate', 'Transition', () => app.openPanel('transition'), { text: 'Transition' }),
    );

    // Fill and border
    this.fillBtn = swatchButton('fill', 'Fill colour', (b) => openColorPicker(b, { theme: app.theme, allowNone: true, current: app.firstSelected()?.fill, onPick: (v) => app.setElementProp('fill', v) }));
    this.borderBtn = swatchButton('borderColor', 'Border colour', (b) => openColorPicker(b, { theme: app.theme, allowNone: true, current: app.firstSelected()?.stroke?.color, onPick: (v) => app.setStroke({ color: v, width: Math.max(1, app.firstSelected()?.stroke?.width || 2) }) }));
    group('fill', this.fillBtn, this.borderBtn,
      tb('borderWeight', 'Border weight', (b) => this.weightMenu(b), { haspopup: true }),
      tb('borderDash', 'Border dash', (b) => this.dashMenu(b), { haspopup: true }));

    // Lines
    this.lineColorBtn = swatchButton('borderColor', 'Line colour', (b) => openColorPicker(b, { theme: app.theme, current: app.firstSelected()?.stroke?.color, onPick: (v) => app.setStroke({ color: v }) }));
    const arrowsBtn = tb('arrow', 'Arrows', (b) => showPopover(b, menu([
      { label: 'No arrows', run: () => app.setLineEnds(false, false) },
      { label: 'Arrow at end', run: () => app.setLineEnds(false, true) },
      { label: 'Arrow at start', run: () => app.setLineEnds(true, false) },
      { label: 'Arrows at both ends', run: () => app.setLineEnds(true, true) },
    ])), { haspopup: true });
    group('line', this.lineColorBtn,
      tb('borderWeight', 'Line weight', (b) => this.weightMenu(b), { haspopup: true }),
      tb('borderDash', 'Line dash', (b) => this.dashMenu(b), { haspopup: true }),
      arrowsBtn);

    // Text
    this.fontBtn = tb(null, 'Font', (b) => this.fontMenu(b), { haspopup: true, text: 'Inter', cls: 'tb-font' });
    this.sizeInput = h('input', { class: 'tb-size', type: 'text', inputmode: 'numeric', 'aria-label': 'Font size', value: '18', title: 'Font size' });
    this.sizeInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        const n = parseFloat(this.sizeInput.value);
        if (n >= 1 && n <= 400) app.setTextProp('size', n);
        app.focusCanvas();
      } else if (e.key === 'Escape') app.focusCanvas();
    });
    this.sizeInput.addEventListener('focus', () => this.sizeInput.select());
    const sizeBox = h('div', { class: 'tb-sizebox' },
      tb('minus', 'Decrease font size', () => app.stepFontSize(-1), { shortcut: 'Mod-Shift-,', cls: 'tb-small' }),
      this.sizeInput,
      tb('plus', 'Increase font size', () => app.stepFontSize(1), { shortcut: 'Mod-Shift-.', cls: 'tb-small' }));
    this.boldBtn = tb('bold', 'Bold', () => app.toggleTextProp('bold'), { shortcut: 'Mod-B' });
    this.italicBtn = tb('italic', 'Italic', () => app.toggleTextProp('italic'), { shortcut: 'Mod-I' });
    this.underlineBtn = tb('underline', 'Underline', () => app.toggleTextProp('underline'), { shortcut: 'Mod-U' });
    this.colorBtn = swatchButton('textColor', 'Text colour', (b) => openColorPicker(b, { theme: app.theme, current: app.currentTextColor(), onPick: (v) => app.setTextProp('color', v) }));
    this.hiliteBtn = swatchButton('highlight', 'Highlight colour', (b) => openColorPicker(b, { theme: app.theme, allowNone: true, noneLabel: 'None', onPick: (v) => app.setTextProp('highlight', v === 'none' ? null : v) }));
    this.hiliteBtn.setColor('transparent');
    this.alignBtn = tb('alignLeft', 'Align', (b) => showPopover(b, h('div', { class: 'tb-row-panel' },
      ...[['left', 'alignLeft', 'Left', 'Mod-Shift-L'], ['center', 'alignCenter', 'Center', 'Mod-Shift-E'], ['right', 'alignRight', 'Right', 'Mod-Shift-R'], ['justify', 'alignJustify', 'Justify', 'Mod-Shift-J']]
        .map(([v, ic, name, sc]) => tb(ic, name, () => { closePopover(); app.setTextProp('align', v); }, { shortcut: sc })),
      h('span', { class: 'tb-sep' }),
      ...[['top', 'alignTopEdge', 'Top'], ['middle', 'alignVCenter', 'Middle'], ['bottom', 'alignBottomEdge', 'Bottom']]
        .map(([v, ic, name]) => tb(ic, `Vertical: ${name}`, () => { closePopover(); app.setTextProp('valign', v); }))), { label: 'Align' }), { haspopup: true });
    const spacingBtn = tb('lineSpacing', 'Line & paragraph spacing', (b) => showPopover(b, menu([
      { heading: 'Line spacing' },
      ...[1, 1.15, 1.25, 1.5, 2].map((v) => ({ label: String(v), checked: app.firstSelectedTextStyle()?.lineHeight === v, run: () => app.setTextProp('lineHeight', v) })),
      { heading: 'Space between paragraphs' },
      ...[[0, 'None'], [0.35, 'Small'], [0.8, 'Medium'], [1.4, 'Large']].map(([v, name]) => ({ label: name, checked: app.firstSelectedTextStyle()?.spacing === v, run: () => app.setTextProp('spacing', v) })),
    ])), { haspopup: true });
    this.bulletBtn = tb('bulletList', 'Bulleted list', () => app.toggleList('ul'), { shortcut: 'Mod-Shift-8' });
    this.numberBtn = tb('orderedList', 'Numbered list', () => app.toggleList('ol'), { shortcut: 'Mod-Shift-7' });
    group('text', this.fontBtn, sizeBox, sep(), this.boldBtn, this.italicBtn, this.underlineBtn, this.colorBtn, this.hiliteBtn, sep(),
      this.alignBtn, spacingBtn, this.bulletBtn, this.numberBtn,
      tb('outdent', 'Decrease indent', () => app.indent(-1), { shortcut: 'Mod-[' }),
      tb('indent', 'Increase indent', () => app.indent(1), { shortcut: 'Mod-]' }),
      tb('clearFormat', 'Clear formatting', () => app.clearFormatting(), { shortcut: 'Mod-\\' }));

    // Pictures
    group('image', tb('imagePlus', 'Replace image', () => app.replaceImage(), { text: 'Replace image' }));

    // Tables
    group('table',
      tb('rowAbove', 'Insert row above', () => app.tableOp('rowAbove')),
      tb('rowBelow', 'Insert row below', () => app.tableOp('rowBelow')),
      tb('colLeft', 'Insert column left', () => app.tableOp('colLeft')),
      tb('colRight', 'Insert column right', () => app.tableOp('colRight')),
      tb('trash', 'Delete row or column', (b) => showPopover(b, menu([
        { label: 'Delete row', run: () => app.tableOp('deleteRow') },
        { label: 'Delete column', run: () => app.tableOp('deleteCol') },
      ])), { haspopup: true }),
      tb('merge', 'Merge cells', (b) => showPopover(b, menu([
        { heading: 'The cell with the cursor' },
        { label: 'Merge with the cell on the right', run: () => app.tableOp('mergeRight') },
        { label: 'Merge with the cell below', run: () => app.tableOp('mergeDown') },
        { label: 'Unmerge', run: () => app.tableOp('unmerge') },
      ])), { haspopup: true }));

    // Charts
    group('chart', tb('chart', 'Edit chart', () => app.openPanel('format'), { text: 'Edit chart' }));

    // Diagrams
    group('diagram', tb('diagram', 'Edit diagram text', () => app.editDiagramText(), { text: 'Edit diagram' }));

    // Any selection
    group('element',
      tb('arrange', 'Arrange', (b) => showPopover(b, menu(this.arrangeItems(), { iconFn }), { label: 'Arrange' }), { haspopup: true }),
      tb('animate', 'Animate', () => app.openPanel('animate'), { text: 'Animate' }),
      tb('format', 'Format options', () => app.openPanel('format'), { text: 'Format options' }));

    this.groups = groups;
    const bar = h('div', { class: 'toolbar', role: 'toolbar', 'aria-label': 'Formatting' },
      this.undoBtn, this.redoBtn, this.zoomBtn, sep(),
      this.selectBtn, this.textBtn, this.imageBtn, this.shapeBtn, this.lineBtn, this.tableBtn, this.chartBtn, sep(),
      ...Object.values(groups));
    // Arrow keys move between toolbar controls (one tab stop).
    bar.addEventListener('keydown', (e) => {
      if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
      if (e.target === this.sizeInput) return;
      const items = [...bar.querySelectorAll('button, input')].filter((b) => b.offsetParent);
      const i = items.indexOf(e.target);
      if (i < 0) return;
      e.preventDefault();
      items[(i + (e.key === 'ArrowRight' ? 1 : items.length - 1)) % items.length].focus();
    });
    return bar;
  }

  weightMenu(anchor) {
    const app = this.app;
    const stroke = app.firstSelected()?.stroke;
    showPopover(anchor, menu([1, 2, 3, 4, 6, 8, 12, 16].map((w) => ({
      label: `${w} pt`, checked: stroke?.width === w && stroke?.color !== 'none', run: () => app.setBorderWeight(w),
    }))), { label: 'Weight' });
  }

  dashMenu(anchor) {
    const app = this.app;
    const current = app.firstSelected()?.stroke?.dash || 'solid';
    showPopover(anchor, menu([['solid', 'Solid'], ['dash', 'Dash'], ['dot', 'Dot'], ['dashDot', 'Dash-dot'], ['longDash', 'Long dash']]
      .map(([id, name]) => ({ label: name, checked: current === id, run: () => app.setStroke({ dash: id }) }))), { label: 'Dash' });
  }

  fontMenu(anchor) {
    const app = this.app;
    const current = app.firstSelectedTextStyle()?.font;
    const theme = app.theme;
    const items = [
      { heading: 'Theme fonts' },
      ...[...new Set([theme.fonts.title, theme.fonts.body])].map((f) => ({ label: f, checked: f === current, labelStyle: `font-family:${fontStack(f)}`, run: () => app.setTextProp('font', f) })),
      { heading: 'All fonts' },
      ...FONTS.map((f) => ({ label: f, checked: f === current, labelStyle: `font-family:${fontStack(f)}`, run: () => app.setTextProp('font', f) })),
    ];
    const m = menu(items);
    m.classList.add('font-menu');
    showPopover(anchor, m, { label: 'Font' });
  }

  /** Show the controls that apply to the selection, and their current state. */
  update() {
    const app = this.app;
    const els = app.selectedElements();
    const editing = app.canvas.isEditing;
    const has = (fn) => els.length > 0 && els.some(fn);
    const show = (name, on) => { this.groups[name].hidden = !on; };
    const textTarget = app.hasTextTarget();
    show('slide', !els.length);
    show('fill', has((e) => e.type === 'shape' || e.type === 'text' || e.type === 'table') && !has((e) => e.type === 'line'));
    show('line', els.length > 0 && els.every((e) => e.type === 'line'));
    show('text', textTarget);
    show('image', els.length === 1 && els[0].type === 'image');
    show('table', els.length === 1 && els[0].type === 'table');
    show('chart', els.length === 1 && els[0].type === 'chart');
    show('diagram', els.length === 1 && els[0].type === 'diagram');
    show('element', els.length > 0);
    this.undoBtn.disabled = !app.history.canUndo && !editing;
    this.redoBtn.disabled = !app.history.canRedo && !editing;
    this.zoomBtn.querySelector('.tb-text').textContent = app.canvas.zoom === 'fit' ? 'Fit' : `${Math.round(app.canvas.scale * 100)}%`;

    const tool = app.canvas.tool;
    this.selectBtn.classList.toggle('is-active', !tool);
    this.textBtn.classList.toggle('is-active', tool?.type === 'text');
    this.shapeBtn.classList.toggle('is-active', tool?.type === 'shape');
    this.lineBtn.classList.toggle('is-active', tool?.type === 'line');

    const first = els[0];
    if (first) {
      const theme = app.theme;
      this.fillBtn.setColor(resolveColor(first.fill, theme, 'transparent'));
      this.borderBtn.setColor(resolveColor(first.stroke?.width ? first.stroke.color : 'none', theme, 'transparent'));
      this.lineColorBtn.setColor(resolveColor(first.stroke?.color, theme, '#000'));
    }
    if (textTarget) {
      const f = app.currentTextFormat();
      if (f) {
        this.fontBtn.querySelector('.tb-text').textContent = f.font;
        if (document.activeElement !== this.sizeInput) this.sizeInput.value = f.size ? String(Math.round(f.size * 10) / 10) : '';
        this.boldBtn.classList.toggle('is-active', !!f.bold);
        this.italicBtn.classList.toggle('is-active', !!f.italic);
        this.underlineBtn.classList.toggle('is-active', !!f.underline);
        this.bulletBtn.classList.toggle('is-active', !!f.bullets);
        this.numberBtn.classList.toggle('is-active', !!f.numbers);
        this.colorBtn.setColor(f.colorCss || '#000');
        this.alignBtn.innerHTML = icon({ left: 'alignLeft', center: 'alignCenter', right: 'alignRight', justify: 'alignJustify' }[f.align] || 'alignLeft') + icon('chevronDown', 'icon tb-caret');
      }
    }
  }
}

/** Small outline of a shape for the shape picker. */
function shapeIcon(id) {
  const tall = ['upArrow', 'downArrow'].includes(id);
  const wide = ['rightArrow', 'leftArrow', 'chevron', 'homePlate', 'parallelogram', 'wave'].includes(id);
  const w = tall ? 16 : 24;
  const hh = wide ? 16 : 24;
  return `<svg viewBox="-2 -2 28 28" width="24" height="24" aria-hidden="true"><path transform="translate(${(24 - w) / 2} ${(24 - hh) / 2})" d="${shapePath(id, w, hh)}" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"/></svg>`;
}
