import { describe, expect, it } from "vitest";
import { DEFAULT_MAX_DPI, dataUriBytes, parseJpegHeader, sampleSize } from "./imageEncode";

/** Assembles a JPEG far enough to carry a frame header: SOI, an APP0 segment
 * that must be skipped by its length, then the frame. */
function jpeg(
  frameMarker: number,
  { width = 40, height = 30, components = 3, precision = 8 } = {},
): Uint8Array {
  return new Uint8Array([
    0xff, 0xd8, // SOI
    0xff, 0xe0, 0x00, 0x06, 0x4a, 0x46, 0x49, 0x46, // APP0, length 6
    0xff, frameMarker, 0x00, 0x11,
    precision,
    (height >> 8) & 0xff, height & 0xff,
    (width >> 8) & 0xff, width & 0xff,
    components,
  ]);
}

describe("parseJpegHeader", () => {
  it("reads dimensions and component count past a skipped segment", () => {
    expect(parseJpegHeader(jpeg(0xc0, { width: 640, height: 480 }))).toEqual({
      widthPx: 640,
      heightPx: 480,
      components: 3,
      baseline: true,
    });
  });

  it("treats SOF0 and SOF1 as baseline, SOF2 (progressive) as not", () => {
    expect(parseJpegHeader(jpeg(0xc0))?.baseline).toBe(true);
    expect(parseJpegHeader(jpeg(0xc1))?.baseline).toBe(true);
    expect(parseJpegHeader(jpeg(0xc2))?.baseline).toBe(false);
  });

  it("reports greyscale and CMYK component counts", () => {
    expect(parseJpegHeader(jpeg(0xc0, { components: 1 }))?.components).toBe(1);
    expect(parseJpegHeader(jpeg(0xc0, { components: 4 }))?.components).toBe(4);
  });

  it("rejects non-JPEG bytes, truncated files and unreadable frames", () => {
    // A PNG signature.
    expect(parseJpegHeader(new Uint8Array([0x89, 0x50, 0x4e, 0x47]))).toBeNull();
    expect(parseJpegHeader(new Uint8Array([0xff, 0xd8]))).toBeNull();
    // 12-bit precision is not something /DCTDecode is promised.
    expect(parseJpegHeader(jpeg(0xc0, { precision: 12 }))).toBeNull();
    // Marker-misaligned bytes rather than a segment.
    expect(parseJpegHeader(new Uint8Array([0xff, 0xd8, 0x00, 0x00, 0x00, 0x00]))).toBeNull();
  });

  it("returns null when the scan starts before any frame header", () => {
    expect(parseJpegHeader(new Uint8Array([0xff, 0xd8, 0xff, 0xda, 0x00, 0x02]))).toBeNull();
  });
});

describe("dataUriBytes", () => {
  it("decodes base64 payloads", () => {
    expect(Array.from(dataUriBytes("data:image/png;base64,AQID") ?? [])).toEqual([1, 2, 3]);
  });

  it("returns null for other schemes and non-base64 data URIs", () => {
    expect(dataUriBytes("blob:http://localhost/abc")).toBeNull();
    expect(dataUriBytes("https://example.com/a.png")).toBeNull();
    expect(dataUriBytes("data:image/svg+xml,<svg/>")).toBeNull();
  });
});

describe("sampleSize", () => {
  it("caps at 300 dpi of the printed size by default (SPEC-034)", () => {
    expect(DEFAULT_MAX_DPI).toBe(300);
    // A 3000 × 1500 px photo printed 72 pt (1 in) wide: 300 px across.
    expect(sampleSize(3000, 1500, 72, 36)).toEqual({ width: 300, height: 150 });
  });

  it("takes a higher ceiling when asked, so 600 dpi is not downsampled to 300 (SPEC-108)", () => {
    expect(sampleSize(3000, 1500, 72, 36, 600)).toEqual({ width: 600, height: 300 });
  });

  it("never upscales past the image's own pixels", () => {
    expect(sampleSize(64, 40, 595, 842, 600)).toEqual({ width: 64, height: 40 });
  });
});
