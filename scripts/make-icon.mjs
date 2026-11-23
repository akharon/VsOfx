// node scripts/make-icon.mjs [out.png] -> media/icon.png (128x128, the Marketplace icon).
// Drawn in code so the icon has no external source file: a blue rounded tile with a white page and
// a folded corner. On the page is the ThalesoMilas emblem (Thales' theorem): a circle with an
// inscribed right triangle on its diameter, and concentric half-rings below the diameter.
import { mkdirSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { deflateSync } from 'node:zlib';

const N = 128;
const S = 4; // supersampling per axis
const px = new Float64Array(N * N * 4);

const inRound = (x, y, x0, y0, x1, y1, r) => {
  if (x < x0 || x > x1 || y < y0 || y > y1) return false;
  const cx = Math.min(Math.max(x, x0 + r), x1 - r);
  const cy = Math.min(Math.max(y, y0 + r), y1 - r);
  return (x - cx) ** 2 + (y - cy) ** 2 <= r * r;
};

/** Colour of a point, or null for transparent. Later shapes paint over earlier ones. */
function shade(x, y) {
  let c = null;
  if (inRound(x, y, 4, 4, 124, 124, 22)) c = [0x1e, 0x5f, 0xb4];
  // page with the top-right corner cut off
  const [x0, y0, x1, y1, fold] = [24, 14, 104, 114, 20];
  if (inRound(x, y, x0, y0, x1, y1, 6) && !(x > x1 - fold && y < y0 + fold && x - (x1 - fold) > y - y0)) c = [0xff, 0xff, 0xff];
  // folded corner triangle
  if (x >= x1 - fold && x <= x1 && y >= y0 && y <= y0 + fold && x - (x1 - fold) <= y - y0) c = [0xc9, 0xd8, 0xee];
  if (c && c[0] === 0xff && inEmblem(x, y)) c = [0x1e, 0x5f, 0xb4];
  return c;
}

// The emblem, centred on the page below the fold.
const E = { x: 64, y: 70, r: 31, w: 4 };
const segDist = (x, y, [ax, ay], [bx, by]) => {
  const dx = bx - ax, dy = by - ay;
  const t = Math.min(1, Math.max(0, ((x - ax) * dx + (y - ay) * dy) / (dx * dx + dy * dy)));
  return Math.hypot(x - ax - t * dx, y - ay - t * dy);
};
const L = [E.x - E.r, E.y]; // diameter ends
const R = [E.x + E.r, E.y];
const ang = (122 * Math.PI) / 180; // apex on the upper-left of the circle, as in the logo
const A = [E.x + E.r * Math.cos(ang), E.y - E.r * Math.sin(ang)];
// right-angle marker: a small square in the apex corner, along both legs
const unit = ([px, py], [qx, qy]) => { const d = Math.hypot(qx - px, qy - py); return [(qx - px) / d, (qy - py) / d]; };
const uL = unit(A, L), uR = unit(A, R), m = 7;
const M1 = [A[0] + uL[0] * m, A[1] + uL[1] * m];
const M2 = [A[0] + uR[0] * m, A[1] + uR[1] * m];
const M3 = [M1[0] + uR[0] * m, M1[1] + uR[1] * m];

function inEmblem(x, y) {
  const h = E.w / 2;
  const d = Math.hypot(x - E.x, y - E.y);
  if (Math.abs(d - E.r) <= h) return true; // outer circle
  if (segDist(x, y, L, R) <= h || segDist(x, y, L, A) <= h || segDist(x, y, A, R) <= h) return true;
  if (segDist(x, y, M1, M3) <= h * 0.7 || segDist(x, y, M2, M3) <= h * 0.7) return true;
  if (y >= E.y && (Math.abs(d - E.r * 0.66) <= h || Math.abs(d - E.r * 0.4) <= h)) return true; // half-rings
  return d <= E.r * 0.1; // centre dot
}

for (let j = 0; j < N; j++)
  for (let i = 0; i < N; i++) {
    let r = 0, g = 0, b = 0, a = 0;
    for (let sy = 0; sy < S; sy++)
      for (let sx = 0; sx < S; sx++) {
        const c = shade(i + (sx + 0.5) / S, j + (sy + 0.5) / S);
        if (c) { r += c[0]; g += c[1]; b += c[2]; a++; }
      }
    const k = (j * N + i) * 4;
    if (a) { px[k] = r / a; px[k + 1] = g / a; px[k + 2] = b / a; }
    px[k + 3] = (255 * a) / (S * S);
  }

const raw = Buffer.alloc(N * (N * 4 + 1));
for (let j = 0; j < N; j++) {
  raw[j * (N * 4 + 1)] = 0;
  for (let i = 0; i < N * 4; i++) raw[j * (N * 4 + 1) + 1 + i] = Math.round(px[j * N * 4 + i]);
}

const crcTable = Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});
const crc = (buf) => {
  let c = 0xffffffff;
  for (const b of buf) c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
};
const chunk = (type, data) => {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type, 'ascii'), data]);
  const c = Buffer.alloc(4);
  c.writeUInt32BE(crc(td));
  return Buffer.concat([len, td, c]);
};
const ihdr = Buffer.alloc(13);
ihdr.writeUInt32BE(N, 0);
ihdr.writeUInt32BE(N, 4);
ihdr[8] = 8; // bit depth
ihdr[9] = 6; // RGBA
const png = Buffer.concat([
  Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
  chunk('IHDR', ihdr),
  chunk('IDAT', deflateSync(raw, { level: 9 })),
  chunk('IEND', Buffer.alloc(0)),
]);

const out = process.argv[2] ?? path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'media', 'icon.png');
mkdirSync(path.dirname(out), { recursive: true });
writeFileSync(out, png);
console.log(`wrote ${out} (${png.length} bytes)`);
