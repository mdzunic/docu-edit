// Turns a picked/dropped/pasted File into what lib/images.ts's insertImage
// command needs (SPEC-024 AC-1, AC-2, AC-33, AC-34): validates type and size
// before ever reading bytes, converts SVG to PNG when the target document is
// .docx-bound (AC-2), and decodes the natural pixel size. Not framework-free
// like the rest of lib/ — it's the browser-only orchestration DocumentView.vue
// delegates to so its own script stays a thin caller, same role
// openDocument.ts plays for the file-open path (both read their File through
// core/fileReader.ts's readFileAsArrayBuffer).
import { decodeImageDimensions } from "./core/imageDimensions";
import { readFileAsArrayBuffer } from "./core/fileReader";
import { type NewImageMedia } from "./images";
import { IMAGE_TOO_LARGE_MESSAGE, isAcceptedImageType, isImageTooLarge } from "./core/media";
import { bytesToBase64 } from "./opaque";
import { convertSvgToPng, SvgConversionError } from "./core/svgToPng";
import type { DocumentFormat } from "./documentModel";

export class ImageInsertError extends Error {}

export const UNSUPPORTED_IMAGE_TYPE_MESSAGE =
  "That file isn't a supported image type (PNG, JPEG, GIF or SVG)";
export const IMAGE_READ_ERROR_MESSAGE = "That image could not be read";

export interface PreparedImageInsert {
  media: NewImageMedia;
  naturalWidthPx: number;
  naturalHeightPx: number;
}

/** Validates and prepares a File for insertImage. Throws ImageInsertError
 * with a user-facing message on any failure (unsupported type, over the 10
 * MB limit, unreadable/malformed bytes, or a failed SVG conversion) — the
 * caller shows it as a toast and inserts nothing, logging nothing to the
 * console (AC-34). `targetFormat` is the document's effective save format
 * (the caller's own `doc.format === "odt" ? "odt" : "docx"` rule, the same
 * one the default download format uses) — only it decides whether an SVG is
 * converted (AC-2) or kept as-is (AC-3). */
export async function prepareImageInsert(file: File, targetFormat: DocumentFormat): Promise<PreparedImageInsert> {
  if (!isAcceptedImageType(file.type)) {
    throw new ImageInsertError(UNSUPPORTED_IMAGE_TYPE_MESSAGE);
  }
  if (isImageTooLarge(file.size)) {
    throw new ImageInsertError(IMAGE_TOO_LARGE_MESSAGE);
  }

  let bytes: Uint8Array;
  try {
    bytes = new Uint8Array(await readFileAsArrayBuffer(file));
  } catch {
    throw new ImageInsertError(IMAGE_READ_ERROR_MESSAGE);
  }
  if (isImageTooLarge(bytes.length)) {
    throw new ImageInsertError(IMAGE_TOO_LARGE_MESSAGE);
  }

  if (file.type === "image/svg+xml" && targetFormat === "docx") {
    let converted;
    try {
      converted = await convertSvgToPng(bytes);
    } catch (error) {
      if (error instanceof SvgConversionError) {
        throw new ImageInsertError(error.message);
      }
      throw error;
    }
    return {
      media: { mediaId: crypto.randomUUID(), contentType: "image/png", data: bytesToBase64(converted.data) },
      naturalWidthPx: converted.widthPx,
      naturalHeightPx: converted.heightPx,
    };
  }

  const dims = decodeImageDimensions(bytes, file.type);
  if (!dims) {
    throw new ImageInsertError(IMAGE_READ_ERROR_MESSAGE);
  }
  return {
    media: { mediaId: crypto.randomUUID(), contentType: file.type, data: bytesToBase64(bytes) },
    naturalWidthPx: dims.widthPx,
    naturalHeightPx: dims.heightPx,
  };
}
