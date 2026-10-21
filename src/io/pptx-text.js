/**
 * Text helpers shared by the PowerPoint reader and writer.
 *
 * Bullets in PowerPoint's symbol fonts (Wingdings, Wingdings 2 and 3,
 * Symbol) are letters that only look right in that font: Wingdings' "l" is
 * a round bullet, "§" a small square. They're shown as the Unicode
 * characters that look like them, which any font can draw, and written back
 * in Wingdings where there's one (see symbolBullet).
 *
 * Each table maps the font's character codes (0x20-0xFF; PowerPoint also
 * writes them in the private use area as U+F020-U+F0FF) to a lookalike.
 */
const table = (pairs) => new Map(Object.entries(pairs).map(([code, ch]) => [Number(code), ch]));

const WINGDINGS = table({
  0x22: '✂', 0x23: '✁', 0x28: '☎', 0x29: '✆', 0x2a: '✉', 0x36: '⌛', 0x3e: '✇', 0x3f: '✍',
  0x41: '✌', 0x45: '☜', 0x46: '☞', 0x47: '☝', 0x48: '☟', 0x4a: '☺', 0x4c: '☹', 0x4e: '☠',
  0x51: '✈', 0x52: '☼', 0x54: '❄', 0x56: '✞', 0x58: '✠', 0x59: '✡', 0x5a: '☪', 0x5b: '☯', 0x5d: '☸',
  0x6c: '●', 0x6d: '❍', 0x6e: '■', 0x6f: '□', 0x70: '◻', 0x71: '❑', 0x72: '❒', 0x73: '⬧', 0x74: '⧫',
  0x75: '◆', 0x76: '❖', 0x77: '⬥', 0x78: '⌧', 0x7a: '⌘', 0x7b: '❀', 0x7c: '✿', 0x7d: '❝', 0x7e: '❞',
  0x9e: '·', 0x9f: '•', 0xa0: '▪', 0xa1: '○', 0xa4: '◉', 0xa5: '◎', 0xa7: '▪', 0xa8: '◻',
  0xaa: '✦', 0xab: '★', 0xac: '✶', 0xad: '✴', 0xae: '✹', 0xaf: '✵', 0xb1: '⌖', 0xb2: '⟡', 0xb5: '✪', 0xb6: '✰',
  0xd5: '⌫', 0xd6: '⌦', 0xd8: '➢', 0xdf: '←', 0xe0: '→', 0xe1: '↑', 0xe2: '↓', 0xe3: '↖', 0xe4: '↗',
  0xe5: '↙', 0xe6: '↘', 0xe8: '➔', 0xef: '⇦', 0xf0: '⇨', 0xf1: '⇧', 0xf2: '⇩', 0xf3: '⬄', 0xf4: '⇳',
  0xfb: '✗', 0xfc: '✓', 0xfd: '☒', 0xfe: '☑',
});

const WINGDINGS_2 = table({
  0x4f: '✗', 0x50: '✓', 0x52: '☑', 0x54: '☒',
  0x96: '•', 0x97: '•', 0x98: '●', 0x99: '●', 0x9a: '●', 0x9b: '●', 0x9c: '○', 0x9d: '○', 0x9e: '○', 0x9f: '◯',
  0xa0: '▪', 0xa1: '▪', 0xa2: '■', 0xa3: '■', 0xa4: '□', 0xa5: '□', 0xa6: '▣', 0xa7: '▣',
  0xae: '◆', 0xaf: '◆', 0xb0: '◇', 0xb1: '◈',
});

const WINGDINGS_3 = table({
  0x70: '▲', 0x71: '▼', 0x72: '◀', 0x73: '▶', 0x74: '◀', 0x75: '▶', 0x76: '▲', 0x77: '▼',
  0x5b: '⬅', 0x5c: '➡', 0x5d: '⬆', 0x5e: '⬇', 0x7d: '▶', 0x7e: '◀',
});

// Symbol is a Greek and mathematical font: its letters are Greek ones.
const GREEK = 'ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ';
const SYMBOL = table({
  ...Object.fromEntries([...GREEK].map((ch, i) => [0x41 + i, ch])),
  ...Object.fromEntries([...GREEK.toLowerCase().replace('ϑ', 'ϕ').replace('ς', 'ϖ')].map((ch, i) => [0x61 + i, ch])),
  0x2a: '∗', 0x2d: '−', 0xa5: '∞', 0xa7: '♣', 0xa8: '♦', 0xa9: '♥', 0xaa: '♠', 0xab: '↔', 0xac: '←', 0xad: '↑',
  0xae: '→', 0xaf: '↓', 0xb0: '°', 0xb1: '±', 0xb7: '•', 0xd7: '⋅', 0xde: '⇒', 0xe0: '◊',
});

const FONTS = {
  wingdings: [WINGDINGS], 'wingdings 2': [WINGDINGS_2], 'wingdings 3': [WINGDINGS_3], symbol: [SYMBOL],
  // LibreOffice's symbol fonts hold Wingdings' bullets and Symbol's signs at these codes.
  opensymbol: [WINGDINGS, SYMBOL], starsymbol: [WINGDINGS, SYMBOL],
};
// What's shown for a character of a symbol font that isn't in its table.
const FALLBACK = { 'wingdings 3': '▶' };

/**
 * The character to show for bullet `ch` in font `font`: a lookalike for a
 * symbol font's character, the character itself for any other font. A
 * character in the private use area that can't be told is a plain bullet.
 */
export function bulletGlyph(ch, font) {
  if (!ch) return ch;
  const name = String(font || '').trim().toLowerCase();
  let code = ch.codePointAt(0);
  const tables = FONTS[name];
  const pua = code >= 0xf020 && code <= 0xf0ff;
  // Characters beyond the font's 256 are Unicode already.
  if ((tables && code <= 0xff) || pua) {
    if (pua) code -= 0xf000;
    for (const t of tables || [WINGDINGS, SYMBOL]) if (t.has(code)) return t.get(code);
    if (tables) return FALLBACK[name] || '•';
  }
  return ch.codePointAt(0) >= 0xe000 && ch.codePointAt(0) <= 0xf8ff ? '•' : ch;
}

// Lookalikes written back in Wingdings, which PowerPoint always has (plain bullets and dots stay as they are).
let wingdingsOf = null;

/** A bullet character as PowerPoint writes it: { font, char } in Wingdings for one of its lookalikes, else null. */
export function symbolBullet(ch) {
  if (!wingdingsOf) {
    wingdingsOf = new Map();
    for (const [code, glyph] of WINGDINGS) if (!wingdingsOf.has(glyph) && !'•·'.includes(glyph)) wingdingsOf.set(glyph, String.fromCharCode(code));
  }
  const char = wingdingsOf.get(ch);
  return char ? { font: 'Wingdings', char } : null;
}
