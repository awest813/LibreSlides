/**
 * Reading pictures into a deck. Big photos are scaled down (a slide never
 * needs more than about 2,000 pixels across), and formats PowerPoint can't
 * show (WebP, AVIF, BMP) are converted to PNG, so exported files open
 * everywhere.
 */
const readAsDataUrl = (blob) => new Promise((resolve, reject) => {
  const r = new FileReader();
  r.onload = () => resolve(r.result);
  r.onerror = () => reject(r.error || new Error('Couldn’t read the file.'));
  r.readAsDataURL(blob);
});

const loadImage = (src) => new Promise((resolve, reject) => {
  const img = new Image();
  img.onload = () => resolve(img);
  img.onerror = () => reject(new Error('That file isn’t a picture this browser can show.'));
  img.src = src;
});

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
