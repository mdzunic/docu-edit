import { describe, it, expect } from "vitest";
import { decodeImageDimensions, gifFrameCount } from "./imageDimensions";

// --- byte-level fixture builders --------------------------------------------

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

function writeUint32BE(bytes: Uint8Array, offset: number, value: number): void {
  bytes[offset] = (value >>> 24) & 0xff;
  bytes[offset + 1] = (value >>> 16) & 0xff;
  bytes[offset + 2] = (value >>> 8) & 0xff;
  bytes[offset + 3] = value & 0xff;
}

function writeUint16BE(bytes: Uint8Array, offset: number, value: number): void {
  bytes[offset] = (value >>> 8) & 0xff;
  bytes[offset + 1] = value & 0xff;
}

function writeUint16LE(bytes: Uint8Array, offset: number, value: number): void {
  bytes[offset] = value & 0xff;
  bytes[offset + 1] = (value >>> 8) & 0xff;
}

/** A minimal (but format-valid enough) PNG byte buffer: signature + IHDR
 * chunk header carrying only what decodeImageDimensions reads (width/height
 * at fixed offsets 16/20) — the module never validates chunk length/CRC. */
function buildPng(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(24);
  bytes.set(PNG_SIGNATURE, 0);
  writeUint32BE(bytes, 8, 13); // IHDR chunk length (unused by the decoder)
  bytes.set([0x49, 0x48, 0x44, 0x52], 12); // "IHDR"
  writeUint32BE(bytes, 16, width);
  writeUint32BE(bytes, 20, height);
  return bytes;
}

function buildGif(width: number, height: number, signature = "GIF89a"): Uint8Array {
  const bytes = new Uint8Array(10);
  for (let i = 0; i < 6; i++) {
    bytes[i] = signature.charCodeAt(i);
  }
  writeUint16LE(bytes, 6, width);
  writeUint16LE(bytes, 8, height);
  return bytes;
}

/** SOI + a single SOF0 marker segment carrying width/height at the fixed
 * offsets decodeJpeg reads relative to the marker. */
function buildJpeg(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(11);
  bytes[0] = 0xff;
  bytes[1] = 0xd8; // SOI
  bytes[2] = 0xff;
  bytes[3] = 0xc0; // SOF0
  writeUint16BE(bytes, 4, 8); // segment length (unused directly)
  bytes[6] = 0x08; // precision
  writeUint16BE(bytes, 7, height);
  writeUint16BE(bytes, 9, width);
  return bytes;
}

/** SOI, then a throwaway APP0 segment (exercises the marker-skipping loop),
 * then SOF0 carrying width/height. */
function buildJpegWithApp0(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(29);
  bytes[0] = 0xff;
  bytes[1] = 0xd8; // SOI
  bytes[2] = 0xff;
  bytes[3] = 0xe0; // APP0
  writeUint16BE(bytes, 4, 16); // segment length: 2 (length field) + 14 payload
  // bytes 6..19 are the APP0 payload — contents don't matter.
  bytes[20] = 0xff;
  bytes[21] = 0xc0; // SOF0
  writeUint16BE(bytes, 22, 8);
  bytes[24] = 0x08; // precision
  writeUint16BE(bytes, 25, height);
  writeUint16BE(bytes, 27, width);
  return bytes;
}

function utf8(text: string): Uint8Array {
  return new TextEncoder().encode(text);
}

describe("decodeImageDimensions — PNG (AC-21)", () => {
  it("reads width/height from the IHDR chunk", () => {
    expect(decodeImageDimensions(buildPng(800, 600), "image/png")).toEqual({
      widthPx: 800,
      heightPx: 600,
    });
  });

  it("reads a 1x1 image", () => {
    expect(decodeImageDimensions(buildPng(1, 1), "image/png")).toEqual({
      widthPx: 1,
      heightPx: 1,
    });
  });

  it("returns null for a truncated buffer (shorter than the IHDR fields)", () => {
    const truncated = buildPng(800, 600).subarray(0, 23);
    expect(decodeImageDimensions(truncated, "image/png")).toBeNull();
  });

  it("returns null when the signature doesn't match", () => {
    const bytes = buildPng(800, 600);
    bytes[0] = 0x00;
    expect(decodeImageDimensions(bytes, "image/png")).toBeNull();
  });

  it("returns null for a zero width or height", () => {
    expect(decodeImageDimensions(buildPng(0, 600), "image/png")).toBeNull();
    expect(decodeImageDimensions(buildPng(800, 0), "image/png")).toBeNull();
  });
});

describe("decodeImageDimensions — GIF (AC-21)", () => {
  it("reads width/height from the logical screen descriptor (GIF89a)", () => {
    expect(decodeImageDimensions(buildGif(320, 240, "GIF89a"), "image/gif")).toEqual({
      widthPx: 320,
      heightPx: 240,
    });
  });

  it("accepts the GIF87a signature too", () => {
    expect(decodeImageDimensions(buildGif(64, 32, "GIF87a"), "image/gif")).toEqual({
      widthPx: 64,
      heightPx: 32,
    });
  });

  it("returns null for a truncated buffer", () => {
    const truncated = buildGif(320, 240).subarray(0, 9);
    expect(decodeImageDimensions(truncated, "image/gif")).toBeNull();
  });

  it("returns null for an unrecognised signature", () => {
    expect(decodeImageDimensions(buildGif(320, 240, "GIF88a"), "image/gif")).toBeNull();
  });

  it("returns null for a zero width or height", () => {
    expect(decodeImageDimensions(buildGif(0, 240), "image/gif")).toBeNull();
  });
});

describe("decodeImageDimensions — JPEG (AC-21)", () => {
  it("reads width/height from an SOF0 segment straight after SOI", () => {
    expect(decodeImageDimensions(buildJpeg(1024, 768), "image/jpeg")).toEqual({
      widthPx: 1024,
      heightPx: 768,
    });
  });

  it("skips over other marker segments (APP0) to find SOF0", () => {
    expect(decodeImageDimensions(buildJpegWithApp0(500, 400), "image/jpeg")).toEqual({
      widthPx: 500,
      heightPx: 400,
    });
  });

  it("returns null when the SOI marker is missing", () => {
    const bytes = buildJpeg(1024, 768);
    bytes[0] = 0x00;
    expect(decodeImageDimensions(bytes, "image/jpeg")).toBeNull();
  });

  it("returns null when no SOF marker is ever found", () => {
    const bytes = new Uint8Array([0xff, 0xd8, 0xff, 0xd9]); // SOI immediately followed by EOI
    expect(decodeImageDimensions(bytes, "image/jpeg")).toBeNull();
  });

  it("returns null for a segment truncated before the dimension bytes", () => {
    const truncated = buildJpeg(1024, 768).subarray(0, 10);
    expect(decodeImageDimensions(truncated, "image/jpeg")).toBeNull();
  });
});

describe("decodeImageDimensions — SVG (AC-21)", () => {
  it("reads explicit width/height attributes (unitless = px)", () => {
    const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100"></svg>';
    expect(decodeImageDimensions(utf8(svg), "image/svg+xml")).toEqual({
      widthPx: 200,
      heightPx: 100,
    });
  });

  it("converts an 'in' suffix at 96 DPI", () => {
    const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="2in" height="1in"></svg>';
    expect(decodeImageDimensions(utf8(svg), "image/svg+xml")).toEqual({
      widthPx: 192,
      heightPx: 96,
    });
  });

  it("converts a 'cm' suffix at 96 DPI / 2.54", () => {
    const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="2.54cm" height="2.54cm"></svg>';
    const result = decodeImageDimensions(utf8(svg), "image/svg+xml");
    expect(result?.widthPx).toBeCloseTo(96, 5);
    expect(result?.heightPx).toBeCloseTo(96, 5);
  });

  it("converts a 'pt' suffix at 96/72 DPI", () => {
    const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="72pt" height="36pt"></svg>';
    const result = decodeImageDimensions(utf8(svg), "image/svg+xml");
    expect(result?.widthPx).toBeCloseTo(96, 5);
    expect(result?.heightPx).toBeCloseTo(48, 5);
  });

  it("falls back to the viewBox when width/height are absent", () => {
    const svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 300 150"></svg>';
    expect(decodeImageDimensions(utf8(svg), "image/svg+xml")).toEqual({
      widthPx: 300,
      heightPx: 150,
    });
  });

  it("prefers explicit width/height over the viewBox when both are present", () => {
    const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="50" height="25" viewBox="0 0 300 150"></svg>';
    expect(decodeImageDimensions(utf8(svg), "image/svg+xml")).toEqual({
      widthPx: 50,
      heightPx: 25,
    });
  });

  it("returns null when neither width/height nor viewBox are present", () => {
    const svg = '<svg xmlns="http://www.w3.org/2000/svg"></svg>';
    expect(decodeImageDimensions(utf8(svg), "image/svg+xml")).toBeNull();
  });

  it("returns null for a malformed viewBox", () => {
    const svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 300"></svg>';
    expect(decodeImageDimensions(utf8(svg), "image/svg+xml")).toBeNull();
  });

  it("returns null for malformed XML", () => {
    const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10">';
    expect(decodeImageDimensions(utf8(svg), "image/svg+xml")).toBeNull();
  });

  it("returns null when the root element isn't <svg>", () => {
    const notSvg = '<root xmlns="http://example.com" width="10" height="10"></root>';
    expect(decodeImageDimensions(utf8(notSvg), "image/svg+xml")).toBeNull();
  });
});

describe("decodeImageDimensions — unrecognised/malformed input", () => {
  it("sniffs PNG/GIF/JPEG straight from the bytes regardless of the declared content type", () => {
    // Only SVG is content-type-gated (it has no fixed magic bytes); the
    // other three formats are recognised by their own signature no matter
    // what contentType the caller passes.
    expect(decodeImageDimensions(buildPng(10, 10), "image/bmp")).toEqual({ widthPx: 10, heightPx: 10 });
  });

  it("returns null for bytes matching no format's signature, whatever the declared content type", () => {
    const garbage = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
    expect(decodeImageDimensions(garbage, "image/bmp")).toBeNull();
  });

  it("returns null for empty bytes", () => {
    expect(decodeImageDimensions(new Uint8Array(0), "image/png")).toBeNull();
  });

  it("returns null for random bytes matching no format's signature", () => {
    const garbage = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
    expect(decodeImageDimensions(garbage, "image/png")).toBeNull();
  });
});

const fourCC = (text: string) => [...text].map((c) => c.charCodeAt(0));

function riff(chunk: string, payload: number[]): Uint8Array {
  const size = payload.length;
  return new Uint8Array([
    ...fourCC("RIFF"),
    (12 + size) & 0xff, ((12 + size) >> 8) & 0xff, 0, 0,
    ...fourCC("WEBP"),
    ...fourCC(chunk),
    size & 0xff, (size >> 8) & 0xff, 0, 0,
    ...payload,
  ]);
}

describe("decodeImageDimensions — WebP (SPEC-107)", () => {
  it("reads a lossy VP8 frame header", () => {
    // 3-byte frame tag, 9D 01 2A start code, then 14-bit LE width/height.
    const payload = [0, 0, 0, 0x9d, 0x01, 0x2a, 320 & 0xff, 320 >> 8, 240 & 0xff, 240 >> 8];
    expect(decodeImageDimensions(riff("VP8 ", payload), "image/webp")).toEqual({
      widthPx: 320,
      heightPx: 240,
    });
  });

  it("reads a lossless VP8L header", () => {
    // 0x2F signature, then width-1 and height-1 as consecutive 14-bit
    // fields of a little-endian 32-bit value.
    const packed = (48 - 1) | ((30 - 1) << 14);
    const payload = [
      0x2f,
      packed & 0xff, (packed >> 8) & 0xff, (packed >> 16) & 0xff, (packed >>> 24) & 0xff,
      0,
    ];
    expect(decodeImageDimensions(riff("VP8L", payload), "image/webp")).toEqual({
      widthPx: 48,
      heightPx: 30,
    });
  });

  it("reads an extended VP8X canvas size", () => {
    // Flags + 3 reserved bytes, then 24-bit LE width-1 / height-1.
    const payload = [
      0x02, 0, 0, 0,
      (700 - 1) & 0xff, ((700 - 1) >> 8) & 0xff, 0,
      (350 - 1) & 0xff, ((350 - 1) >> 8) & 0xff, 0,
    ];
    expect(decodeImageDimensions(riff("VP8X", payload), "image/webp")).toEqual({
      widthPx: 700,
      heightPx: 350,
    });
  });

  it("returns null for a malformed VP8 frame (bad start code) and unknown chunks", () => {
    const badStart = [0, 0, 0, 0xff, 0xff, 0xff, 10, 0, 10, 0];
    expect(decodeImageDimensions(riff("VP8 ", badStart), "image/webp")).toBeNull();
    expect(decodeImageDimensions(riff("ANMF", [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]), "image/webp")).toBeNull();
  });

  it("returns null for a VP8L header with a bad signature or a non-zero version", () => {
    const packed = (48 - 1) | ((30 - 1) << 14);
    const header = [packed & 0xff, (packed >> 8) & 0xff, (packed >> 16) & 0xff, (packed >>> 24) & 0xff];
    expect(decodeImageDimensions(riff("VP8L", [0x2e, ...header, 0]), "image/webp")).toBeNull();
    const versioned = [...header];
    versioned[3] |= 0x20; // version 1
    expect(decodeImageDimensions(riff("VP8L", [0x2f, ...versioned, 0]), "image/webp")).toBeNull();
  });

  it("returns null for a truncated RIFF container", () => {
    expect(decodeImageDimensions(utf8("RIFF1234WEBP"), "image/webp")).toBeNull();
  });
});

/** A 1 × 1 GIF89a with `frames` image descriptors (no colour tables), a
 * graphic-control extension before each frame, then the trailer. */
function gif(frames: number): Uint8Array {
  const bytes = [0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 1, 0, 1, 0, 0, 0, 0];
  for (let i = 0; i < frames; i += 1) {
    bytes.push(0x21, 0xf9, 4, 0, 10, 0, 0, 0);
    bytes.push(0x2c, 0, 0, 0, 0, 1, 0, 1, 0, 0, 2, 2, 0x4c, 0x01, 0);
  }
  bytes.push(0x3b);
  return new Uint8Array(bytes);
}

describe("gifFrameCount (SPEC-112: an animated GIF is edited as its first frame)", () => {
  it("counts image descriptors by walking the blocks, not by scanning bytes", () => {
    expect(gifFrameCount(gif(1))).toBe(1);
    expect(gifFrameCount(gif(2))).toBe(2);
    // A truncated file counts the frames seen before the cut.
    expect(gifFrameCount(gif(3).subarray(0, 50))).toBe(2);
    expect(gifFrameCount(new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61]))).toBe(0);
    expect(gifFrameCount(new Uint8Array(32))).toBe(0);
  });
});
