import { deflateSync } from "node:zlib";

/**
 * Minimal, dependency-free PNG encoder — used instead of a native-binding image library (sharp/
 * canvas) specifically because this app must `npm install` and build in any environment without
 * assuming a C++ toolchain is available (see backend/package.json — every other dependency here is
 * pure JS/WASM). Produces a real, valid, openable PNG: signature + IHDR + one IDAT (zlib-deflated
 * raw RGB scanlines, filter type 0) + IEND, with correct CRC32 per chunk.
 */

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(buf: Buffer): number {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]!) & 0xff]! ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type: string, data: Buffer): Buffer {
  const typeBuf = Buffer.from(type, "ascii");
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const body = Buffer.concat([typeBuf, data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body), 0);
  return Buffer.concat([len, body, crc]);
}

/** A simple deterministic "business chart" motif: a light background with a few colored bars — reads as a placeholder chart image, not a blank rectangle, without needing a real charting/image library. */
export function generateBusinessChartPng(width = 480, height = 320, seed = 1): Buffer {
  const bg: [number, number, number] = [246, 248, 251];
  const barColors: [number, number, number][] = [
    [27, 47, 196],
    [16, 133, 92],
    [214, 158, 46],
    [176, 58, 46],
  ];
  const barCount = 4 + (seed % 3);
  const margin = Math.round(width * 0.08);
  const plotWidth = width - margin * 2;
  const plotHeight = height - margin * 2;
  const barWidth = Math.floor(plotWidth / (barCount * 1.6));
  const gap = Math.floor((plotWidth - barWidth * barCount) / (barCount + 1));

  const raw = Buffer.alloc((width * 3 + 1) * height);
  for (let y = 0; y < height; y++) {
    const rowStart = y * (width * 3 + 1);
    raw[rowStart] = 0; // filter type 0 (none)
    for (let x = 0; x < width; x++) {
      let [r, g, b] = bg;
      for (let i = 0; i < barCount; i++) {
        const barX0 = margin + gap * (i + 1) + barWidth * i;
        const barX1 = barX0 + barWidth;
        const heightFrac = 0.3 + ((i * 37 + seed * 13) % 60) / 100;
        const barY0 = margin + Math.round(plotHeight * (1 - heightFrac));
        const barY1 = height - margin;
        if (x >= barX0 && x < barX1 && y >= barY0 && y < barY1) {
          [r, g, b] = barColors[i % barColors.length]!;
        }
      }
      const px = rowStart + 1 + x * 3;
      raw[px] = r;
      raw[px + 1] = g;
      raw[px + 2] = b;
    }
  }

  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 2; // color type: RGB
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  const idat = deflateSync(raw);
  const signature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

  return Buffer.concat([signature, chunk("IHDR", ihdr), chunk("IDAT", idat), chunk("IEND", Buffer.alloc(0))]);
}
