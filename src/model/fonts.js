/**
 * Fonts offered in the font menu. The first group is bundled with the app
 * (see src/styles/fonts.css), so it looks the same on every device and
 * offline; the rest are common system fonts with sensible fallbacks.
 */
export const BUNDLED_FONTS = ['Inter', 'Montserrat', 'Playfair Display', 'Lora', 'Roboto Slab', 'Oswald'];

export const SYSTEM_FONTS = [
  'Arial', 'Calibri', 'Georgia', 'Times New Roman', 'Trebuchet MS', 'Verdana', 'Segoe UI',
  'Courier New', 'Comic Sans MS', 'Impact',
];

export const FONTS = [...BUNDLED_FONTS, ...SYSTEM_FONTS];

const SERIF = new Set(['Playfair Display', 'Lora', 'Roboto Slab', 'Georgia', 'Times New Roman']);
const MONO = new Set(['Courier New']);

/** CSS font-family for a font name, with a fallback of the same kind. */
export function fontStack(name) {
  if (!name) return 'Inter, system-ui, sans-serif';
  const generic = MONO.has(name) ? 'monospace' : SERIF.has(name) ? 'serif' : 'sans-serif';
  return `"${name.replace(/"/g, '')}", ${generic}`;
}

export const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 96, 120];

/**
 * The four faces a font family can carry in a PowerPoint file, with the CSS
 * weight and style each one is drawn with. Fonts that came embedded in an
 * opened file are kept in the deck as deck.fonts: { family: { face: data URL } }.
 */
export const FONT_FACES = {
  regular: { weight: '400', style: 'normal' },
  bold: { weight: '700', style: 'normal' },
  italic: { weight: '400', style: 'italic' },
  boldItalic: { weight: '700', style: 'italic' },
};

/** Whether PowerPoint files carry the fonts they use (File › Embed fonts). On unless turned off on this device. */
export function embedFontsSetting() {
  try {
    return localStorage.getItem('ls:embedFonts') !== '0';
  } catch {
    return true;
  }
}

export function setEmbedFontsSetting(on) {
  try { localStorage.setItem('ls:embedFonts', on ? '1' : '0'); } catch { /* storage blocked */ }
}
