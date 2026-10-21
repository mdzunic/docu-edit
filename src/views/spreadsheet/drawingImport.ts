// Reading an image for Insert ▸ Image… and an image paste (SPEC-131 "Insert ▸
// Image", "Paste"): what useSheetDrawings hands drawings.ts' insertImage — the
// bytes, the content type and the natural size in px (96 dpi) — or the
// message that says why nothing went in.
//
// - The type must be one ACCEPTED_IMAGE_MIME_TYPES names; a file over
//   isImageTooLarge's limit is refused with IMAGE_TOO_LARGE_MESSAGE before a
//   byte is read (and again after an SVG's rasterisation).
// - An SVG is rasterised by convertSvgToPng for an .xlsx workbook (Excel's
//   pictures are bitmaps; its SVG extension is not written here) and kept as
//   SVG for .ods, which holds SVG pictures natively.
// - The natural size is decodeImageDimensions' (no decode needed); a file it
//   cannot read is loaded into an Image as a fallback, and one that will not
//   load is refused as broken.
// Browser-only beyond the decode (convertSvgToPng and the Image fallback), so
// the tests stub the loader.
import { decodeImageDimensions } from "../../lib/core/imageDimensions";
import { IMAGE_TOO_LARGE_MESSAGE, isAcceptedImageType, isImageTooLarge } from "../../lib/core/media";
import { convertSvgToPng } from "../../lib/core/svgToPng";
import type { SpreadsheetFormat } from "../../lib/spreadsheet/model";

export const UNSUPPORTED_IMAGE_MESSAGE = "That file isn't a supported image (PNG, JPEG, GIF or SVG).";
export const UNREADABLE_IMAGE_MESSAGE = "That image is broken or unreadable.";

/** How long the Image fallback waits before calling an image broken. */
const LOAD_TIMEOUT_MS = 5000;

export interface PreparedImage {
  bytes: Uint8Array;
  contentType: string;
  naturalPx: { width: number; height: number };
}

/** What a picked or pasted image needs: a File or a Blob with its type. */
export type ImageSource = Pick<Blob, "type" | "size" | "arrayBuffer">;

/** The natural size of bytes the decoder could not read, by loading them
 * into an Image; null when it will not load. */
export type ImageLoader = (bytes: Uint8Array, contentType: string) => Promise<{ width: number; height: number } | null>;

export const loadNaturalSize: ImageLoader = (bytes, contentType) => {
  if (typeof Image === "undefined" || typeof URL.createObjectURL !== "function") {
    return Promise.resolve(null);
  }
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: contentType }));
  return new Promise((resolve) => {
    const image = new Image();
    const done = (size: { width: number; height: number } | null) => {
      clearTimeout(timer);
      URL.revokeObjectURL(url);
      resolve(size);
    };
    const timer = setTimeout(() => done(null), LOAD_TIMEOUT_MS);
    image.addEventListener("load", () => done(image.naturalWidth > 0 && image.naturalHeight > 0 ? { width: image.naturalWidth, height: image.naturalHeight } : null));
    image.addEventListener("error", () => done(null));
    image.src = url;
  });
};

/** The image ready for insertImage, or the message to show (see the module
 * header). */
export async function prepareImage(file: ImageSource, format: SpreadsheetFormat, load: ImageLoader = loadNaturalSize): Promise<PreparedImage | string> {
  if (!isAcceptedImageType(file.type)) {
    return UNSUPPORTED_IMAGE_MESSAGE;
  }
  if (isImageTooLarge(file.size)) {
    return IMAGE_TOO_LARGE_MESSAGE;
  }
  let bytes: Uint8Array;
  try {
    bytes = new Uint8Array(await file.arrayBuffer());
  } catch {
    return UNREADABLE_IMAGE_MESSAGE;
  }
  let contentType = file.type;
  if (contentType === "image/svg+xml" && format !== "ods") {
    try {
      const png = await convertSvgToPng(bytes);
      bytes = png.data;
      contentType = "image/png";
    } catch (error) {
      return error instanceof Error && error.message ? error.message : UNREADABLE_IMAGE_MESSAGE;
    }
    if (isImageTooLarge(bytes.length)) {
      return IMAGE_TOO_LARGE_MESSAGE;
    }
  }
  const decoded = decodeImageDimensions(bytes, contentType);
  const naturalPx = decoded ? { width: decoded.widthPx, height: decoded.heightPx } : await load(bytes, contentType);
  if (!naturalPx) {
    return UNREADABLE_IMAGE_MESSAGE;
  }
  return { bytes, contentType, naturalPx };
}
