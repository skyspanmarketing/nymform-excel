// Draws the Nymform ribbon and manifest icons as PNGs with no image dependencies.
// The mark: a rounded square in the accent color with two groups of lines, the right group
// offset from the left, standing for "original data -> a representation of it".
import { writeFileSync, mkdirSync } from "node:fs";
import { deflateSync } from "node:zlib";

const ACCENT = [0x4b, 0x3f, 0xd1]; // indigo, the one accent color
const SIZES = [16, 32, 64, 80, 128];
const SS = 8; // supersampling per axis

function crc32(buf) {
  let c, crc = 0xffffffff;
  for (let n = 0; n < buf.length; n++) {
    c = (crc ^ buf[n]) & 0xff;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crc = (crc >>> 8) ^ c;
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type, "ascii"), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(td));
  return Buffer.concat([len, td, crc]);
}

function png(size, rgba) {
  const raw = Buffer.alloc(size * (size * 4 + 1));
  for (let y = 0; y < size; y++) {
    raw[y * (size * 4 + 1)] = 0;
    rgba.copy(raw, y * (size * 4 + 1) + 1, y * size * 4, (y + 1) * size * 4);
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; ihdr[9] = 6; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", deflateSync(raw, { level: 9 })),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

// Shape test in unit coordinates (0..1). Returns "bg", "line" or null.
function shape(x, y) {
  const r = 0.2;
  const inset = 0.04;
  const lo = inset, hi = 1 - inset;
  if (x < lo || x > hi || y < lo || y > hi) return null;
  const cx = Math.min(Math.max(x, lo + r), hi - r);
  const cy = Math.min(Math.max(y, lo + r), hi - r);
  if ((x - cx) ** 2 + (y - cy) ** 2 > r * r) return null;
  // Three lines on the left, three on the right shifted down by half a pitch.
  const t = 0.085; // line thickness
  const rows = [0.3, 0.5, 0.7];
  for (const ry of rows) {
    if (x >= 0.2 && x <= 0.45 && Math.abs(y - ry) <= t / 2) return "line";
    if (x >= 0.55 && x <= 0.8 && Math.abs(y - (ry + 0.1)) <= t / 2 && ry + 0.1 < 0.85) return "line";
  }
  return "bg";
}

mkdirSync("assets", { recursive: true });
for (const size of SIZES) {
  const buf = Buffer.alloc(size * size * 4);
  for (let py = 0; py < size; py++) {
    for (let px = 0; px < size; px++) {
      let bg = 0, line = 0;
      for (let sy = 0; sy < SS; sy++) {
        for (let sx = 0; sx < SS; sx++) {
          const s = shape((px + (sx + 0.5) / SS) / size, (py + (sy + 0.5) / SS) / size);
          if (s === "bg") bg++;
          else if (s === "line") line++;
        }
      }
      const total = SS * SS;
      const a = (bg + line) / total;
      const lineShare = bg + line ? line / (bg + line) : 0;
      const i = (py * size + px) * 4;
      for (let ch = 0; ch < 3; ch++) buf[i + ch] = Math.round(ACCENT[ch] * (1 - lineShare) + 255 * lineShare);
      buf[i + 3] = Math.round(a * 255);
    }
  }
  writeFileSync(`assets/icon-${size}.png`, png(size, buf));
}
console.log(`Wrote ${SIZES.map((s) => `assets/icon-${s}.png`).join(", ")}`);
