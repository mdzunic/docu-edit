// Turns the images a page places into PDF image XObject payloads
// (SPEC-034). Browser-bound like domExtract.ts — it needs the platform image
// decoder, a canvas and CompressionStream — but free of Vue imports, so the
// pure format logic below stays unit-testable and the rest is covered e2e.
//
// Two encodings, picked per image:
//  - a baseline JPEG is passed through untouched as /DCTDecode. PDF carries
//    JPEG natively, so this costs nothing and keeps a photo at its original
//    size rather than exploding it into raw samples.
//  - everything else (PNG, GIF, converted SVG, progressive or CMYK JPEG) is
//    drawn to a canvas and embedded as zlib-compressed RGB samples, with the
//    alpha channel split out into an /SMask when the image is transparent.
import type { PdfImageData } from "./buildPdf";
// Shared with buildPdf.ts, which deflates its font and content streams —
// see core/deflate.ts for why it lives there rather than here.
import { deflate } from "../core/deflate";

/** The frame header of a JPEG: what /DCTDecode needs to know before the
 * original bytes can be handed to a PDF viewer unchanged. */
export interface JpegHeader {
  widthPx: number;
  heightPx: number;
  /** 1 (greyscale), 3 (YCbCr) or 4 (CMYK/YCCK). */
  components: number;
  /** SOF0/SOF1. Progressive (SOF2) and arithmetic-coded frames are not what
   * /DCTDecode promises to decode, so they take the canvas path instead. */
  baseline: boolean;
}

/** Markers carrying no payload length: padding, restart markers and SOI. */
function isStandaloneMarker(marker: number): boolean {
  return marker === 0x01 || marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd7);
}

/** Start-of-frame markers are 0xC0..0xCF except the three that mean
 * something else: DHT (C4), JPG (C8) and DAC (CC). */
function isFrameMarker(marker: number): boolean {
  return marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
}

/** Reads a JPEG's frame header, or null when the bytes are not a JPEG whose
 * dimensions can be read before the scan. */
export function parseJpegHeader(bytes: Uint8Array): JpegHeader | null {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) {
    return null;
  }
  let offset = 2;
  while (offset + 1 < bytes.length) {
    if (bytes[offset] !== 0xff) {
      return null; // Not marker-aligned: refuse rather than guess.
    }
    const marker = bytes[offset + 1];
    if (isStandaloneMarker(marker)) {
      offset += 2;
      continue;
    }
    // The scan (and end of image) means no frame header was found first.
    if (marker === 0xda || marker === 0xd9) {
      return null;
    }
    if (offset + 3 >= bytes.length) {
      return null;
    }
    const length = (bytes[offset + 2] << 8) | bytes[offset + 3];
    if (length < 2) {
      return null;
    }
    if (isFrameMarker(marker)) {
      // [length 2][precision 1][height 2][width 2][components 1]
      if (offset + 9 >= bytes.length) {
        return null;
      }
      const precision = bytes[offset + 4];
      const heightPx = (bytes[offset + 5] << 8) | bytes[offset + 6];
      const widthPx = (bytes[offset + 7] << 8) | bytes[offset + 8];
      const components = bytes[offset + 9];
      if (precision !== 8 || widthPx === 0 || heightPx === 0) {
        return null;
      }
      return { widthPx, heightPx, components, baseline: marker === 0xc0 || marker === 0xc1 };
    }
    offset += 2 + length;
  }
  return null;
}


/** The bytes behind a `data:` URI, or null for any other scheme. */
export function dataUriBytes(src: string): Uint8Array | null {
  const comma = src.indexOf(",");
  if (!src.startsWith("data:") || comma < 0) {
    return null;
  }
  const meta = src.slice(5, comma);
  if (!meta.endsWith(";base64")) {
    return null;
  }
  const binary = atob(src.slice(comma + 1));
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

function loadImage(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.addEventListener("load", () => resolve(image));
    image.addEventListener("error", () => reject(new Error("image could not be decoded")));
    image.src = src;
  });
}

/** Cuts the `crop` rectangle (source pixels) out of an image and returns it
 * as a lossless PNG data: URI at the source's own resolution — SPEC-108's
 * cover/actual crop, applied BEFORE encodeImage so a cropped-away region
 * never reaches the output file. Throws when the image cannot be decoded. */
export async function cropImage(
  src: string,
  crop: { xPx: number; yPx: number; widthPx: number; heightPx: number },
): Promise<string> {
  const image = await loadImage(src);
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(crop.widthPx));
  canvas.height = Math.max(1, Math.round(crop.heightPx));
  const context = canvas.getContext("2d");
  if (!context) {
    throw new Error("This browser cannot crop images");
  }
  context.drawImage(image, crop.xPx, crop.yPx, crop.widthPx, crop.heightPx, 0, 0, canvas.width, canvas.height);
  return canvas.toDataURL("image/png");
}

/** SPEC-034's sampling ceiling: enough for print. */
export const DEFAULT_MAX_DPI = 300;

/** Samples an image at no more than `maxDpi` (300 unless the caller asks —
 * SPEC-108's 600 dpi sheet does) of the size it is printed at: enough for
 * print, and it keeps a huge source photo from becoming a huge PDF. Never
 * upscales past the image's own pixels. */
export function sampleSize(
  naturalWidth: number,
  naturalHeight: number,
  maxWidthPt: number,
  maxHeightPt: number,
  maxDpi: number = DEFAULT_MAX_DPI,
): { width: number; height: number } {
  const capPx = Math.max(1, Math.ceil((Math.max(maxWidthPt, maxHeightPt) / 72) * maxDpi));
  const longEdge = Math.max(naturalWidth, naturalHeight);
  const scale = longEdge > capPx ? capPx / longEdge : 1;
  return {
    width: Math.max(1, Math.round(naturalWidth * scale)),
    height: Math.max(1, Math.round(naturalHeight * scale)),
  };
}

/** Encodes one image for embedding, or null when it cannot be decoded (a
 * broken src must not fail the whole export — the page still exports without
 * it). `maxWidthPt`/`maxHeightPt` are the largest printed size the document
 * gives this image, which sets the sampling ceiling together with `maxDpi`
 * (default 300, SPEC-034's cap). */
export async function encodeImage(
  src: string,
  maxWidthPt: number,
  maxHeightPt: number,
  maxDpi: number = DEFAULT_MAX_DPI,
): Promise<PdfImageData | null> {
  try {
    // A baseline JPEG rides into the PDF as-is.
    const raw = src.startsWith("data:image/jpeg") ? dataUriBytes(src) : null;
    if (raw) {
      const header = parseJpegHeader(raw);
      if (header && header.baseline && (header.components === 3 || header.components === 1)) {
        return {
          widthPx: header.widthPx,
          heightPx: header.heightPx,
          filter: "DCTDecode",
          colorSpace: header.components === 1 ? "DeviceGray" : "DeviceRGB",
          data: raw,
        };
      }
    }

    const image = await loadImage(src);
    const { width, height } = sampleSize(
      image.naturalWidth || image.width,
      image.naturalHeight || image.height,
      maxWidthPt,
      maxHeightPt,
      maxDpi,
    );
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d");
    if (!context) {
      return null;
    }
    context.drawImage(image, 0, 0, width, height);
    const { data } = context.getImageData(0, 0, width, height);

    const rgb = new Uint8Array(width * height * 3);
    const alpha = new Uint8Array(width * height);
    let opaque = true;
    for (let pixel = 0; pixel < width * height; pixel += 1) {
      const source = pixel * 4;
      const target = pixel * 3;
      rgb[target] = data[source];
      rgb[target + 1] = data[source + 1];
      rgb[target + 2] = data[source + 2];
      alpha[pixel] = data[source + 3];
      if (data[source + 3] !== 255) {
        opaque = false;
      }
    }
    return {
      widthPx: width,
      heightPx: height,
      filter: "FlateDecode",
      colorSpace: "DeviceRGB",
      data: await deflate(rgb),
      alpha: opaque ? undefined : await deflate(alpha),
    };
  } catch {
    return null;
  }
}
