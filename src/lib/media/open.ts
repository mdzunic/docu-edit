// Assembles opened media bytes into a MediaModel (SPEC-107): the media
// product's counterpart of assembleDocumentModel/assemblePresentation. The
// bytes are stored verbatim (base64 in a SPEC-024 MediaPart — the store's
// media split persists them out of the shell record); everything else on the
// model is what could be learned about the file without rendering it.
// Framework-free and store-free per the media layer's boundary.
import { bytesToBase64 } from "../core/packageParts";
import { decodeImageDimensions } from "../core/imageDimensions";
import type { DocumentProperties } from "../core/documentProperties";
import type { MediaPart } from "../core/media";
import { parsePdf, parsePdfDate, pdfTextOf, type PdfDocument } from "../core/pdfDocument";
import { MEDIA_CONTENT_TYPES, type MediaFormat, type MediaModel, type PdfSummary } from "./model";

/** The structural summary a parsed PDF contributes to its MediaModel. */
export function summarisePdf(doc: PdfDocument): PdfSummary {
  return {
    version: doc.version,
    pageCount: doc.pages.length,
    pages: doc.pages.map((page) => ({
      widthPt: page.widthPt,
      heightPt: page.heightPt,
      rotate: page.rotate,
    })),
    linearised: doc.linearised,
    reconstructed: doc.reconstructed,
  };
}

/** /Info mapped onto the product-neutral DocumentProperties vocabulary.
 * Producer has no dedicated slot, so it travels as an extra field (the
 * read-only shape SPEC-041 defined for exactly this). */
export function pdfInfoProperties(doc: PdfDocument): DocumentProperties {
  const properties: DocumentProperties = {};
  const info = doc.info;
  if (!info) {
    return properties;
  }
  const text = (key: string): string => pdfTextOf(info.get(key)).trim();
  const title = text("Title");
  if (title) {
    properties.title = title;
  }
  const subject = text("Subject");
  if (subject) {
    properties.subject = subject;
  }
  const keywords = text("Keywords");
  if (keywords) {
    properties.keywords = keywords;
  }
  const author = text("Author");
  if (author) {
    properties.author = author;
  }
  const created = parsePdfDate(text("CreationDate"));
  if (created) {
    properties.created = created;
  }
  const modified = parsePdfDate(text("ModDate"));
  if (modified) {
    properties.modified = modified;
  }
  const producer = text("Producer");
  if (producer) {
    properties.extraFields = [{ name: "Producer", xml: "", value: producer }];
  }
  return properties;
}

/** Builds a MediaModel from a media file's bytes. For a PDF this parses the
 * structure (throwing PdfEncryptedError / PdfParseError for files that are
 * refused — the open path surfaces the message); for an image it reads the
 * intrinsic dimensions, tolerating files whose size cannot be read.
 * `parsed` is `parsePdf(bytes)` when the caller has just run it over these
 * very bytes (SPEC-110's self-check), so a large PDF is not parsed twice. */
export async function assembleMediaModel(
  title: string,
  format: MediaFormat,
  bytes: Uint8Array,
  parsed?: PdfDocument,
): Promise<MediaModel> {
  const contentType = MEDIA_CONTENT_TYPES[format];
  const sourceMediaId = crypto.randomUUID();
  const part: MediaPart = {
    mediaId: sourceMediaId,
    contentType,
    data: bytesToBase64(bytes),
  };
  const model: MediaModel = {
    kind: "media",
    id: crypto.randomUUID(),
    title,
    format,
    contentType,
    sourceMediaId,
    media: [part],
    byteLength: bytes.length,
  };
  if (format === "pdf") {
    const doc = parsed ?? (await parsePdf(bytes));
    model.pdf = summarisePdf(doc);
    const properties = pdfInfoProperties(doc);
    if (Object.keys(properties).length > 0) {
      model.properties = properties;
    }
  } else {
    const dimensions = decodeImageDimensions(bytes, contentType);
    if (dimensions) {
      model.image = { widthPx: dimensions.widthPx, heightPx: dimensions.heightPx };
    }
  }
  return model;
}
