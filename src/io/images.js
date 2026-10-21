/**
 * Reading pictures into a deck. Big photos are scaled down (a slide never
 * needs more than about 2,000 pixels across), and formats PowerPoint can't
 * show (WebP, AVIF, BMP) are converted to PNG, so exported files open
 * everywhere.
 */
export const readAsDataUrl = (blob) => new Promise((resolve, reject) => {
  const r = new FileReader();
  r.onload = () => resolve(r.result);
  r.onerror = () => reject(r.error || new Error('Couldn’t read the file.'));
  r.readAsDataURL(blob);
});

/**
 * A PNG, JPEG or GIF's size in pixels and its resolution, read from the
 * start of its data URL: { w, h, dpiX, dpiY } (96 dpi when it doesn't say,
 * as Office takes it), or null for anything else.
 */
export function pictureSize(dataUrl) {
  const m = /^data:image\/[\w+.-]+;base64,/i.exec(dataUrl || '');
  if (!m) return null;
  let b;
  try {
    // Enough of the start for the header (and a JPEG's metadata before its size).
    const head = dataUrl.slice(m[0].length, m[0].length + 349524);
    b = Uint8Array.from(atob(head.slice(0, head.length - (head.length % 4))), (c) => c.charCodeAt(0));
  } catch {
    return null;
  }
  const u16 = (i) => (b[i] << 8) | b[i + 1];
  const u32 = (i) => ((b[i] << 24) >>> 0) + ((b[i + 1] << 16) | (b[i + 2] << 8) | b[i + 3]);
  const out = (w, h, dpiX = 96, dpiY = dpiX) => (w > 0 && h > 0 ? { w, h, dpiX: dpiX > 0 ? dpiX : 96, dpiY: dpiY > 0 ? dpiY : 96 } : null);
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) {
    const [w, h] = [u32(16), u32(20)];
    // Pixels per metre, if it says (pHYs comes before the image data).
    for (let i = 8; i + 8 < b.length;) {
      const len = u32(i);
      const type = String.fromCharCode(b[i + 4], b[i + 5], b[i + 6], b[i + 7]);
      if (type === 'IDAT') break;
      if (type === 'pHYs' && b[i + 16] === 1) return out(w, h, u32(i + 8) * 0.0254, u32(i + 12) * 0.0254);
      i += 12 + len;
    }
    return out(w, h);
  }
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return out(b[6] | (b[7] << 8), b[8] | (b[9] << 8));
  if (b[0] === 0xff && b[1] === 0xd8) {
    let dpi = null;
    for (let i = 2; i + 9 < b.length;) {
      if (b[i] !== 0xff) return null;
      const marker = b[i + 1];
      if (marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd7) || marker === 0x01) {
        i += 2;
        continue;
      }
      const len = u16(i + 2);
      // JFIF's density (in dots per inch, or per centimetre).
      if (marker === 0xe0 && String.fromCharCode(...b.subarray(i + 4, i + 8)) === 'JFIF') {
        const unit = b[i + 11];
        const [x, y] = [u16(i + 12), u16(i + 14)];
        if (unit === 1) dpi = [x, y];
        else if (unit === 2) dpi = [x * 2.54, y * 2.54];
      }
      // Start of frame (not DHT, JPG or DAC): its height and width.
      if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker)) return out(u16(i + 7), u16(i + 5), ...(dpi || []));
      i += 2 + len;
    }
  }
  return null;
}

/** A picture's natural size in points (at its own resolution), or null. */
export function pictureSizePt(dataUrl) {
  const s = pictureSize(dataUrl);
  return s && { w: (s.w * 72) / s.dpiX, h: (s.h * 72) / s.dpiY };
}

const loadImage = (src) => new Promise((resolve, reject) => {
  const img = new Image();
  img.onload = () => resolve(img);
  img.onerror = () => reject(new Error('That file isn’t a picture this browser can show.'));
  img.src = src;
});

/**
 * A picture recoloured as PowerPoint's duotone effect does: each pixel's
 * brightness picks a colour between `dark` and `light` ("#rrggbb"). Null
 * where pictures can't be drawn (no canvas) or this one can't be read.
 */
export async function duotoneImage(dataUrl, dark, light, { maxSide = 2560 } = {}) {
  const canvas = typeof document !== 'undefined' ? document.createElement('canvas') : null;
  const g = canvas?.getContext?.('2d', { willReadFrequently: true });
  if (!g) return null;
  try {
    const img = await Promise.race([loadImage(dataUrl), new Promise((resolve) => { setTimeout(resolve, 10000, null); })]);
    if (!img?.naturalWidth || !img.naturalHeight) return null;
    const k = Math.min(1, maxSide / Math.max(img.naturalWidth, img.naturalHeight));
    canvas.width = Math.max(1, Math.round(img.naturalWidth * k));
    canvas.height = Math.max(1, Math.round(img.naturalHeight * k));
    g.drawImage(img, 0, 0, canvas.width, canvas.height);
    const pixels = g.getImageData(0, 0, canvas.width, canvas.height);
    const rgb = (hex) => [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16));
    const [a, b] = [rgb(dark), rgb(light)];
    const d = pixels.data;
    for (let i = 0; i < d.length; i += 4) {
      const lum = (0.2126 * d[i] + 0.7152 * d[i + 1] + 0.0722 * d[i + 2]) / 255;
      for (let c = 0; c < 3; c++) d[i + c] = a[c] + (b[c] - a[c]) * lum;
    }
    g.putImageData(pixels, 0, 0);
    return /^data:image\/jpeg/i.test(dataUrl) ? canvas.toDataURL('image/jpeg', 0.9) : canvas.toDataURL('image/png');
  } catch {
    return null;
  }
}

export async function readImageFile(file, { maxSide = 1920 } = {}) {
  if (!file.type.startsWith('image/') && !/\.(png|jpe?g|gif|svg|webp|bmp|avif)$/i.test(file.name || '')) {
    throw new Error('That file isn’t a picture.');
  }
  const dataUrl = await readAsDataUrl(file);
  const img = await loadImage(dataUrl);
  let width = img.naturalWidth || 400;
  let height = img.naturalHeight || 300;
  const type = file.type || '';
  // SVG scales by itself; GIFs may be animated. Keep both as they are.
  if (/svg|gif/.test(type)) return { dataUrl, width, height };
  const tooBig = Math.max(width, height) > maxSide;
  const convert = !/^image\/(png|jpeg)$/.test(type);
  if (!tooBig && !convert && file.size < 1.5 * 1024 * 1024) return { dataUrl, width, height };
  const k = Math.min(1, maxSide / Math.max(width, height));
  width = Math.round(width * k);
  height = Math.round(height * k);
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  canvas.getContext('2d').drawImage(img, 0, 0, width, height);
  const out = type === 'image/jpeg' ? canvas.toDataURL('image/jpeg', 0.9) : canvas.toDataURL('image/png');
  return { dataUrl: out, width, height };
}
