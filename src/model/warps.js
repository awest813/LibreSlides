import WARPS from './preset-warps.js';

/**
 * WordArt text warps: PowerPoint's Text Effects › Transform, which bend a
 * box's text along a path or squeeze it into a shape (a:prstTxWarp). A box
 * with one has `style.warp = { preset, adjust?, stretch? }`: `preset` is the
 * OOXML name (textArchUp…), `adjust` its adjust values ({ adj: 10800000 },
 * in the definition's own units) and `stretch` (PowerPoint's fromWordArt)
 * stretches text along the "follow path" warps to fill the path, as the
 * other warps always do. See render/warp.js for how it's drawn.
 */
export { WARPS };

// PowerPoint's Transform gallery, with readable names.
export const WARP_GROUPS = [
  ['Follow path', [['textArchUp', 'Arch up'], ['textArchDown', 'Arch down'], ['textCircle', 'Circle'], ['textButton', 'Button']]],
  ['Arches and rings', [
    ['textArchUpPour', 'Arch up (filled)'], ['textArchDownPour', 'Arch down (filled)'], ['textCirclePour', 'Circle (filled)'], ['textButtonPour', 'Button (filled)'],
    ['textRingInside', 'Ring inside'], ['textRingOutside', 'Ring outside'],
  ]],
  ['Shapes', [
    ['textStop', 'Stop'], ['textTriangle', 'Triangle up'], ['textTriangleInverted', 'Triangle down'], ['textChevron', 'Chevron up'],
    ['textChevronInverted', 'Chevron down'], ['textCanUp', 'Can up'], ['textCanDown', 'Can down'],
  ]],
  ['Waves and curves', [
    ['textWave1', 'Wave'], ['textWave2', 'Wave (reversed)'], ['textDoubleWave1', 'Double wave'], ['textWave4', 'Double wave (reversed)'],
    ['textCurveUp', 'Curve up'], ['textCurveDown', 'Curve down'],
  ]],
  ['Inflate and deflate', [
    ['textInflate', 'Inflate'], ['textDeflate', 'Deflate'], ['textInflateTop', 'Inflate top'], ['textDeflateTop', 'Deflate top'],
    ['textInflateBottom', 'Inflate bottom'], ['textDeflateBottom', 'Deflate bottom'], ['textDeflateInflate', 'Deflate-inflate'],
    ['textDeflateInflateDeflate', 'Deflate-inflate-deflate'],
  ]],
  ['Perspective', [
    ['textFadeRight', 'Fade right'], ['textFadeLeft', 'Fade left'], ['textFadeUp', 'Fade up'], ['textFadeDown', 'Fade down'],
    ['textSlantUp', 'Slant up'], ['textSlantDown', 'Slant down'], ['textCascadeUp', 'Cascade up'], ['textCascadeDown', 'Cascade down'],
  ]],
];

/** Warps whose text runs along each path; the others fit it between pairs of paths. */
export const FOLLOW_PATH = new Set(['textArchUp', 'textArchDown', 'textCircle', 'textButton']);

/** A preset's adjust values when it has none of its own ({ adj: 10800000 }), where they're plain numbers. */
export function warpDefaults(preset) {
  const out = {};
  for (const [name, formula] of WARPS[preset]?.av || []) {
    const m = /^val (-?\d+)$/.exec(formula);
    if (m) out[name] = +m[1];
    else if (formula === 'val cd2') out[name] = 10800000;
  }
  return out;
}

/** The adjust value the editor's "Amount" slider changes, with its range: { name, min, max, value }. */
export function mainAdjust(warp) {
  const def = WARPS[warp?.preset];
  const name = def?.av?.[0]?.[0];
  const range = name && def.ranges?.[name];
  if (!range) return null;
  return { name, min: range[0], max: range[1], value: warp.adjust?.[name] ?? warpDefaults(warp.preset)[name] ?? range[0] };
}

/** Whether a warp leaves text as it is (no warp, textNoShape, or textPlain at its default). */
export function isPlainWarp(warp) {
  if (!warp || !WARPS[warp.preset]) return true;
  if (warp.preset !== 'textPlain') return false;
  const defaults = warpDefaults('textPlain');
  return Object.entries(warp.adjust || {}).every(([k, v]) => v === defaults[k]);
}

/**
 * A warp from an untrusted source (a file, an old deck), or null: a known
 * preset that changes the text, with adjust values it has, finite and
 * within its handles' ranges.
 */
export function cleanWarp(warp) {
  if (!warp || typeof warp !== 'object' || !Object.hasOwn(WARPS, warp.preset)) return null;
  const def = WARPS[warp.preset];
  const out = { preset: warp.preset };
  const adjust = {};
  for (const [name] of def.av || []) {
    const v = warp.adjust?.[name];
    if (typeof v !== 'number' || !Number.isFinite(v)) continue;
    const [min, max] = def.ranges?.[name] || [-21600000, 21600000];
    adjust[name] = Math.round(Math.max(min, Math.min(max, v)));
  }
  if (Object.keys(adjust).length) out.adjust = adjust;
  if (warp.stretch === true) out.stretch = true;
  return isPlainWarp(out) ? null : out;
}
