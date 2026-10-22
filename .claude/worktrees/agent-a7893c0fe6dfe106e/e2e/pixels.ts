import { inflateSync } from "node:zlib";

// Decodes a Playwright `page.screenshot()` PNG buffer so a test can assert
// on COMPOSITED pixels — what the user actually sees after the browser
// stacks the shape canvases over the text DOM — rather than on a single
// canvas's backing store (SPEC-073 AC-7/AC-9: transparency compositing and
// the shadow pass are only observable in the final composite). Plain
// unfiltering of the IDAT stream via node:zlib; no dependency needed.

export interface DecodedScreenshot {
  width: number;
  height: number;
  /** The composited [r, g, b] at device pixel (x, y). */
  rgbAt(x: number, y: number): [number, number, number];
}

export function decodePng(buffer: Buffer): DecodedScreenshot {
  let pos = 8;
  const idat: Buffer[] = [];
  let width = 0;
  let height = 0;
  let colorType = 6;
  while (pos < buffer.length) {
    const length = buffer.readUInt32BE(pos);
    const type = buffer.toString("ascii", pos + 4, pos + 8);
    if (type === "IHDR") {
      width = buffer.readUInt32BE(pos + 8);
      height = buffer.readUInt32BE(pos + 12);
      colorType = buffer[pos + 8 + 9];
    } else if (type === "IDAT") {
      idat.push(buffer.subarray(pos + 8, pos + 8 + length));
    }
    pos += length + 12;
  }
  const raw = inflateSync(Buffer.concat(idat));
  const bpp = colorType === 6 ? 4 : 3; // truecolour with or without alpha
  const stride = width * bpp;
  const out = Buffer.alloc(height * stride);
  let rp = 0;
  for (let y = 0; y < height; y++) {
    const filter = raw[rp++];
    for (let x = 0; x < stride; x++) {
      const cur = raw[rp++];
      const a = x >= bpp ? out[y * stride + x - bpp] : 0;
      const b = y > 0 ? out[(y - 1) * stride + x] : 0;
      const c = x >= bpp && y > 0 ? out[(y - 1) * stride + x - bpp] : 0;
      let v: number;
      if (filter === 0) {
        v = cur;
      } else if (filter === 1) {
        v = cur + a;
      } else if (filter === 2) {
        v = cur + b;
      } else if (filter === 3) {
        v = cur + ((a + b) >> 1);
      } else {
        const p = a + b - c;
        const pa = Math.abs(p - a);
        const pb = Math.abs(p - b);
        const pc = Math.abs(p - c);
        v = cur + (pa <= pb && pa <= pc ? a : pb <= pc ? b : c);
      }
      out[y * stride + x] = v & 255;
    }
  }
  return {
    width,
    height,
    rgbAt(x: number, y: number) {
      const cx = Math.min(width - 1, Math.max(0, Math.round(x)));
      const cy = Math.min(height - 1, Math.max(0, Math.round(y)));
      const i = cy * stride + cx * bpp;
      return [out[i], out[i + 1], out[i + 2]];
    },
  };
}
