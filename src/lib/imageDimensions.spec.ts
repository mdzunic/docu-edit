import { describe, it, expect } from "vitest";
import { decodeImageDimensions } from "./imageDimensions";

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
