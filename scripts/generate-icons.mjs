// Draws the app icons (public/icons/icon-192.png, icon-512.png) from the logo
// geometry in public/favicon.svg, with no dependencies: shapes are rasterised
// with signed distances and 4×4 supersampling, then written as PNG.
// Run with `npm run icons` after changing the logo.
import { writeFileSync, mkdirSync } from 'node:fs';
import { deflateSync } from 'node:zlib';

const BRAND = [0xc4, 0x3e, 0x1c];
const WHITE = [255, 255, 255];

// Logo in its 48 × 48 SVG units.
const roundRect = (x, y, cx, cy, hw, hh, r) => {
  const qx = Math.abs(x - cx) - hw + r;
  const qy = Math.abs(y - cy) - hh + r;
  return Math.hypot(Math.max(qx, 0), Math.max(qy, 0)) + Math.min(Math.max(qx, qy), 0) - r;
};
const segment = (x, y, ax, ay, bx, by) => {
  const px = x - ax;
  const py = y - ay;
  const dx = bx - ax;
  const dy = by - ay;
  const t = Math.max(0, Math.min(1, (px * dx + py * dy) / (dx * dx + dy * dy)));
  return Math.hypot(px - dx * t, py - dy * t);
};
const inTriangle = (x, y) => {
  const [a, b, c] = [[21, 17.5], [21, 26.5], [28.5, 22]];
  const s = (p, q, r) => (p[0] - r[0]) * (q[1] - r[1]) - (q[0] - r[0]) * (p[1] - r[1]);
  const d1 = s([x, y], a, b);
  const d2 = s([x, y], b, c);
  const d3 = s([x, y], c, a);
  return !((d1 < 0 || d2 < 0 || d3 < 0) && (d1 > 0 || d2 > 0 || d3 > 0));
};
const STROKE = 1.6; // half of stroke-width 3.2
const isWhite = (x, y) =>
  Math.abs(roundRect(x, y, 24, 22, 15, 10, 2.5)) <= STROKE
  || inTriangle(x, y)
  || segment(x, y, 24, 32, 24, 37) <= STROKE
  || segment(x, y, 18, 38, 30, 38) <= STROKE;

function render(size) {
  // Maskable-safe: full-bleed brand background, logo in the middle 72%.
  const inner = size * 0.72;
  const off = (size - inner) / 2;
  const k = 48 / inner;
  const SS = 4;
  const rows = [];
  for (let py = 0; py < size; py++) {
    const row = Buffer.alloc(1 + size * 3); // filter byte + RGB
    for (let px = 0; px < size; px++) {
      let hits = 0;
      for (let sy = 0; sy < SS; sy++) {
        for (let sx = 0; sx < SS; sx++) {
          const x = (px + (sx + 0.5) / SS - off) * k;
          const y = (py + (sy + 0.5) / SS - off) * k;
          if (isWhite(x, y)) hits++;
        }
      }
      const a = hits / (SS * SS);
      for (let c = 0; c < 3; c++) row[1 + px * 3 + c] = Math.round(BRAND[c] * (1 - a) + WHITE[c] * a);
    }
    rows.push(row);
  }
  return png(size, size, Buffer.concat(rows));
}

const CRC = new Int32Array(256).map((_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c;
});
const crc32 = (buf) => {
  let c = -1;
  for (const b of buf) c = CRC[(c ^ b) & 255] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
};
function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([len, body, crc]);
}
function png(w, h, raw) {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0);
  ihdr.writeUInt32BE(h, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 2; // RGB
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', deflateSync(raw, { level: 9 })),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

mkdirSync('public/icons', { recursive: true });
for (const size of [192, 512]) {
  writeFileSync(`public/icons/icon-${size}.png`, render(size));
  console.log(`public/icons/icon-${size}.png`);
}
