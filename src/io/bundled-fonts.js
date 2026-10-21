/**
 * The font files bundled with the app (the Latin subsets src/styles/fonts.css
 * loads), by family and face, for embedding in PowerPoint files. These are
 * the WOFF versions: browsers can't unpack WOFF2 (Brotli) for us. Vite gives
 * the same hashed files the stylesheet uses, and the service worker keeps
 * them, so embedding works offline.
 */
import interRegular from '@fontsource/inter/files/inter-latin-400-normal.woff?url';
import interItalic from '@fontsource/inter/files/inter-latin-400-italic.woff?url';
import interBold from '@fontsource/inter/files/inter-latin-700-normal.woff?url';
import interBoldItalic from '@fontsource/inter/files/inter-latin-700-italic.woff?url';
import montserratRegular from '@fontsource/montserrat/files/montserrat-latin-400-normal.woff?url';
import montserratBold from '@fontsource/montserrat/files/montserrat-latin-700-normal.woff?url';
import playfairRegular from '@fontsource/playfair-display/files/playfair-display-latin-400-normal.woff?url';
import playfairItalic from '@fontsource/playfair-display/files/playfair-display-latin-400-italic.woff?url';
import playfairBold from '@fontsource/playfair-display/files/playfair-display-latin-700-normal.woff?url';
import loraRegular from '@fontsource/lora/files/lora-latin-400-normal.woff?url';
import loraItalic from '@fontsource/lora/files/lora-latin-400-italic.woff?url';
import loraBold from '@fontsource/lora/files/lora-latin-700-normal.woff?url';
import robotoSlabRegular from '@fontsource/roboto-slab/files/roboto-slab-latin-400-normal.woff?url';
import robotoSlabBold from '@fontsource/roboto-slab/files/roboto-slab-latin-700-normal.woff?url';
import oswaldRegular from '@fontsource/oswald/files/oswald-latin-400-normal.woff?url';
import oswaldBold from '@fontsource/oswald/files/oswald-latin-700-normal.woff?url';

export const BUNDLED_FACES = {
  Inter: { regular: interRegular, bold: interBold, italic: interItalic, boldItalic: interBoldItalic },
  Montserrat: { regular: montserratRegular, bold: montserratBold },
  'Playfair Display': { regular: playfairRegular, bold: playfairBold, italic: playfairItalic },
  Lora: { regular: loraRegular, bold: loraBold, italic: loraItalic },
  'Roboto Slab': { regular: robotoSlabRegular, bold: robotoSlabBold },
  Oswald: { regular: oswaldRegular, bold: oswaldBold },
};
