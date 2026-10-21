import {
  Bold, Italic, Underline, Strikethrough, Baseline, Highlighter, List, ListOrdered, ListIndentIncrease, ListIndentDecrease,
  TextAlignStart, TextAlignCenter, TextAlignEnd, TextAlignJustify, Undo2, Redo2, Scissors, Copy, ClipboardPaste, CopyPlus,
  RemoveFormatting, Search, Table, Image, ImagePlus, Link, Minus, FilePlus, FolderOpen, Printer, Download, Upload, FileDown,
  ZoomIn, ZoomOut, Moon, Sun, Trash2, Plus, X, ChevronDown, ChevronUp, ChevronLeft, ChevronRight, Check, Keyboard, Info,
  Maximize, Minimize, Square, House, Ellipsis, Pencil, Clock, Rows3, Columns3, BetweenHorizontalStart, BetweenHorizontalEnd,
  BetweenVerticalStart, BetweenVerticalEnd, MousePointer2, Shapes, Slash, MoveUpRight, PaintBucket, PenLine, Layers, Sparkles,
  SlidersHorizontal, Palette, Presentation, Play, Eye, EyeOff, BringToFront, SendToBack, AlignStartVertical, AlignCenterVertical,
  AlignEndVertical, AlignStartHorizontal, AlignCenterHorizontal, AlignEndHorizontal, AlignHorizontalDistributeCenter,
  AlignVerticalDistributeCenter, RotateCw, RotateCcw, FlipHorizontal2, FlipVertical2, StickyNote, Timer, Pause, Grid3x3, Magnet,
  PanelRight, ArrowUp, ArrowDown, ListCollapse, LayoutTemplate, Type, SquareDashed, MonitorPlay, Save,
  Clapperboard,
} from 'lucide';

const ICONS = {
  bold: Bold, italic: Italic, underline: Underline, strike: Strikethrough, textColor: Baseline, highlight: Highlighter,
  bulletList: List, orderedList: ListOrdered, indent: ListIndentIncrease, outdent: ListIndentDecrease,
  alignLeft: TextAlignStart, alignCenter: TextAlignCenter, alignRight: TextAlignEnd, alignJustify: TextAlignJustify,
  undo: Undo2, redo: Redo2, cut: Scissors, copy: Copy, paste: ClipboardPaste, duplicate: CopyPlus, clearFormat: RemoveFormatting,
  search: Search, table: Table, image: Image, imagePlus: ImagePlus, link: Link, minus: Minus, newDoc: FilePlus, open: FolderOpen,
  print: Printer, download: Download, upload: Upload, fileDown: FileDown, zoomIn: ZoomIn, zoomOut: ZoomOut, moon: Moon, sun: Sun,
  trash: Trash2, plus: Plus, close: X, chevronDown: ChevronDown, chevronUp: ChevronUp, chevronLeft: ChevronLeft,
  chevronRight: ChevronRight, check: Check, keyboard: Keyboard, info: Info, fullscreen: Maximize, exitFullscreen: Minimize,
  square: Square, home: House, more: Ellipsis, rename: Pencil, clock: Clock, rows: Rows3, columns: Columns3,
  rowAbove: BetweenHorizontalStart, rowBelow: BetweenHorizontalEnd, colLeft: BetweenVerticalStart, colRight: BetweenVerticalEnd,
  select: MousePointer2, shapes: Shapes, line: Slash, arrow: MoveUpRight, fill: PaintBucket, borderColor: PenLine,
  arrange: Layers, animate: Sparkles, format: SlidersHorizontal, theme: Palette, present: Presentation, play: Play,
  eye: Eye, eyeOff: EyeOff, toFront: BringToFront, toBack: SendToBack,
  alignLeftEdge: AlignStartVertical, alignHCenter: AlignCenterVertical, alignRightEdge: AlignEndVertical,
  alignTopEdge: AlignStartHorizontal, alignVCenter: AlignCenterHorizontal, alignBottomEdge: AlignEndHorizontal,
  distributeH: AlignHorizontalDistributeCenter, distributeV: AlignVerticalDistributeCenter,
  rotateCw: RotateCw, rotateCcw: RotateCcw, flipH: FlipHorizontal2, flipV: FlipVertical2, notes: StickyNote, timer: Timer,
  pause: Pause, grid: Grid3x3, snap: Magnet, panelRight: PanelRight, up: ArrowUp, down: ArrowDown, lineSpacing: ListCollapse,
  layout: LayoutTemplate, save: Save, textBox: Type, borderWeight: Minus, borderDash: SquareDashed, presenter: MonitorPlay,
  media: Clapperboard,
};

const cache = new Map();

const ATTRS = 'xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false"';

function nodeToString([tag, attrs, children]) {
  const a = Object.entries(attrs || {}).map(([k, v]) => `${k}="${v}"`).join(' ');
  const inner = Array.isArray(children) ? children.map(nodeToString).join('') : '';
  return `<${tag} ${a}>${inner}</${tag}>`;
}

/** SVG markup for an icon name (falls back to an empty square). */
export function icon(name, cls = 'icon') {
  const key = `${name}|${cls}`;
  if (!cache.has(key)) {
    const def = ICONS[name] || Square;
    cache.set(key, `<svg ${ATTRS} class="${cls}">${def.map(nodeToString).join('')}</svg>`);
  }
  return cache.get(key);
}

/** The app logo: a slide with a play mark, on the brand colour. */
export const LOGO_SVG = (fg = '#fff', bg = '#c43e1c', size = 26) =>
  `<svg width="${size}" height="${size}" viewBox="0 0 48 48" aria-hidden="true"><rect width="48" height="48" rx="10" fill="${bg}"/><rect x="9" y="12" width="30" height="20" rx="2.5" fill="none" stroke="${fg}" stroke-width="3.2"/><path d="M21 17.5v9l7.5-4.5z" fill="${fg}"/><path d="M24 32v5M18 38h12" stroke="${fg}" stroke-width="3.2" stroke-linecap="round"/></svg>`;
